# hotcall — Project Library: the HOT-CALL ASN entry driver (the "One Cycle Entry" / out-of-cycle urgent
# shipment the operator keys by hand), reimplemented as PURE Jython-2.7/CPython logic + a thin gateway
# driver — the producer-recipe pattern (sibling of asn/code.py create_asn, the SECOND ASN producer).
#
# Source truth (read in full):
#   HotCallEntry.pas                                     — the legacy "One Cycle Entry" form (THE target)
#   docs/analysis/edi/hotcall-coverage-analysis.md       — the P12 coverage analysis (seam COVERED, entry GAP)
#   docs/analysis/edi/project-library/asn/code.py        — the M1 ASN write seam this REUSES verbatim
#   docs/analysis/edi/spike-asndetail-rekey.sql          — the Q1 re-key (the @HotCall=1 always-INSERT branch)
#   docs/analysis/edi/856/project-library/edi856/code.py — the 856 builder a hot-call ASN flows through (8HC)
#
# ROLE: a hot-call ASN is create_asn (asn/code.py) with the ENTIRE AD_FRSPULL->ratio->manifest-gen stage
# DELETED and a manual `[(part, qty)] + operator-typed manifest` substituted, plus @HotCall=1 and the -1/-1
# seq sentinel (hotcall-coverage-analysis.md §2). NO AD_FRSPULL, NO SELECT_ForecastDetailBCASN, NO ratio
# explode, NO banker's rounding — the operator types the final parts + qtys directly. It is the SIMPLEST
# producer in the suite.
#
# This module holds BOTH:
#   * computeHotcallDetails(...) + the validators — the PURE part. Imports nothing, references no gateway
#     global at import time -> CPython-importable for the unit test (scripts/e2e/test_hotcall_build.py).
#   * create_hotcall_asn(...) — the GATEWAY DRIVER. ONE Inventory transaction, reusing the M1 ASN write
#     seam (INSERT_ASNInfo + INSERT_ASNDetail). References `system` only at CALL time (like asn.create_asn /
#     edi856.send_856), so importing the module never needs the gateway runtime; it is driven headlessly
#     through scripts/e2e/jython_shim.py (the R8 system.db shim) for end-to-end parity
#     (scripts/e2e/test_hotcall_e2e.py).
#
# Jython 2.7 (8.1 and 8.3, no delta — pure dict assembly + system.db; no Python-3-only libs).
#
# DIVERGENCES from the legacy (ratified, consistent with M1 — hotcall-coverage-analysis.md §4 hazards):
#   * EIN-at-SEND, NOT EIN-at-create (P14). Legacy stamps SiteEIN+1 at create + bumps the ALC counter in
#     the create tx (HotCallEntry.pas:250-251,290-292). We follow the M1 at-send model: write IN_ASN_EIN=0
#     here; the per-site EIN is allocated atomically from INV_SITES.IN_EIN_SEQ at 856 SEND (edi856.send_856).
#     Same decision already ratified for create_asn — extended to hot-calls for ONE consistent EIN mechanism.
#   * Header IN_QTY = SUM of the detail qtys (P14), FIXING the legacy stale-loop-var @QTY garbage
#     (HotCallEntry.pas:246-247 passes whatever `qty` the last validation iteration left in the var — NOT a
#     deliberate total). No wire impact (the 856/810 read DETAIL qty), and the sum is the obviously-correct
#     value a header total should carry.


class HotcallError(Exception):
    """Raised on an entry-validation error (the legacy ShowMessage+exit guards: manifest length, broadcast
    manifest, missing part/qty). The driver does NOT open a transaction until validation passes, so a raise
    writes nothing — matching the Delphi `exit` before BeginTrans (HotCallEntry.pas:143-218 are all before
    :221 BeginTrans)."""
    pass


# =================================================================================================
# PURE part — validation + the manual rows -> ASN detail rows.
# =================================================================================================

def _validate_manifest(manifest):
    """The two legacy manifest guards (HotCallEntry.pas), reproduced:

      1. length >= 8 (HotCallEntry.pas:157-162, ShowMessage('Manifest number must 8 characters')).
         The .dfm caps the field at MaxLength=8 (hotcall-coverage-analysis.md §1), so in practice the
         manifest is EXACTLY 8 chars; the Pascal test is `length < 8 -> reject`, i.e. >= 8 passes. We
         enforce the SAME >= 8 lower bound. (The numeric check is COMMENTED OUT in the .pas, :150-155 —
         so a hot-call manifest is NOT required to be numeric. We do NOT add a numeric check.)

      2. NON-broadcast: a hot-call manifest is operator-typed and must NOT collide with the generated
         broadcast '7'-prefixed manifests (asn._manifest -> '7'+...). A '7'-prefix manifest would emit as
         M391 on the wire (edi856/edi810 IT101 = 'M391' if manifest starts '7' else 'M390') — a hot-call
         is intrinsically M390 (hotcall-coverage-analysis.md §1/§4). The legacy relies on the operator
         not typing a '7' manifest; the rebuild makes that explicit (faithful intent, fail-fast).

    Returns the manifest unchanged on success; raises HotcallError otherwise."""
    if manifest is None:
        raise HotcallError("Manifest number must 8 characters")  # legacy wording (HotCallEntry.pas:159)
    m = str(manifest)
    if len(m) < 8:
        raise HotcallError("Manifest number must 8 characters")  # HotCallEntry.pas:159 (verbatim wording)
    if m.startswith("7"):
        # M390-vs-M391 hazard: a hot-call is M390 (non-'7'). A '7'-prefix would mis-emit as M391 broadcast.
        raise HotcallError(
            "Hot-call manifest %r must not start with '7' (that is a broadcast/M391 manifest)" % (m,))
    return m


def _validate_item(part, qty):
    """Per-row validation mirroring the two legacy loops:
      * qty must be numeric & > 0 (HotCallEntry.pas:184-197: TryStrToInt then Qty<=0 reject).
      * a part is REQUIRED when a qty is present (HotCallEntry.pas:204-218: a non-empty qty Edit with an
        empty part ComboBox -> 'Part number required').
    Returns (part, intQty) on success; raises HotcallError otherwise. The part is upper-cased to match the
    legacy AssyPartsCodeN_ComboBox (CharCase=ecUpperCase, hotcall-coverage-analysis.md §1)."""
    if part is None or str(part).strip() == "":
        raise HotcallError("Part number required")               # HotCallEntry.pas:212
    try:
        q = int(qty)
    except (TypeError, ValueError):
        raise HotcallError("Qty must be numeric")                # HotCallEntry.pas:186
    if q <= 0:
        raise HotcallError("Qty must be greater than 0")         # HotCallEntry.pas:193
    return (str(part).strip().upper(), q)


def computeHotcallDetails(items, manifest):
    """Turn the operator's manual [(part, qty), ...] + the typed manifest into the ASN detail rows that
    feed INSERT_ASNDetail (@HotCall=1). PURE — no DB, no fan-out, no ratio, no rounding (the whole point of
    a hot-call: the operator types the FINAL qty per part — hotcall-coverage-analysis.md §2).

    Reproduces the legacy detail loop (HotCallEntry.pas:258-285): one row per non-empty entry, every row
    carrying the SAME operator-typed manifest, @HotCall=1 (always-INSERT, never the per-manifest accumulate
    — spike-asndetail-rekey.sql:85-90). The asnId is stamped by the DRIVER once INSERT_ASNInfo returns it.

    Parameters
    ----------
    items : list of (part, qty) pairs (or dicts {'part','qty'}). Blank entries (empty part AND empty qty)
        are SKIPPED, exactly as the legacy loop skips a row whose qty Edit is '' (:262). A non-blank row
        is validated (_validate_item) and emitted.
    manifest : the operator-typed 8-char manifest (validated by _validate_manifest before this is called).

    Returns
    -------
    list of detail dicts {'manifest', 'partNumber', 'qty', 'hotCall'(=1)} in entry order. The header
    IN_QTY the driver writes = sum of these qtys (P14 — fixes the legacy @QTY garbage).

    Raises HotcallError on: no items at all; a row with a qty but no part (or vice versa); a non-numeric or
    <= 0 qty. (The legacy validates qty across ALL rows BEFORE inserting any — we do the same by validating
    every non-blank row here, before the driver opens its transaction.)
    """
    norm = []
    for it in items:
        if isinstance(it, dict):
            part = it.get("part", it.get("partNumber"))
            qty = it.get("qty")
        else:
            part, qty = it[0], it[1]
        # SKIP a fully-blank row (no part, no qty) — the legacy loop ignores an empty qty Edit (:262).
        partBlank = part is None or str(part).strip() == ""
        qtyBlank = qty is None or str(qty).strip() == ""
        if partBlank and qtyBlank:
            continue
        # A half-filled row (qty present, part blank — or vice versa) is the legacy error case, not a skip.
        p, q = _validate_item(part, qty)
        norm.append((p, q))

    if not norm:
        # the legacy has no explicit "at least one item" guard, but a hot-call with zero parts would create
        # an empty ASN (header only) the 856 then rejects (no feed rows). Fail fast here — the operator must
        # enter at least one part. (Faithful intent: a hot-call exists to ship parts.)
        raise HotcallError("Enter at least one part and quantity")

    details = []
    for (part, qty) in norm:
        details.append({
            "manifest": manifest,
            "partNumber": part,
            "qty": qty,
            "hotCall": 1,
        })
    return details


def _detailQtySum(details):
    """The header IN_QTY (P14): SUM of the detail qtys — the obviously-correct total that REPLACES the
    legacy stale-loop-var @QTY garbage (HotCallEntry.pas:246-247)."""
    return sum(int(d["qty"]) for d in details)


# =================================================================================================
# create_hotcall_asn — the GATEWAY DRIVER. ONE Inventory transaction, reusing the M1 ASN write seam.
# =================================================================================================
#
# IG81-COMPAT: every gateway API used here (runScalarPrepQuery(...,tx=) / runPrepUpdate(...,tx=) /
# beginTransaction / commitTransaction / rollbackTransaction / closeTransaction / getLogger) is identical
# on 8.1.52 and 8.3 — no version guard needed (same set asn.create_asn / edi856.send_856 use).

from db_shared import CONNECTION as DATABASE   # centralized DB-conn name (default Inventory_Spike; single prod-rename point)

# The hot-call header sentinel (HotCallEntry.pas:238-243): VC_START_SEQ_NUMBER = VC_END_SEQ_NUMBER = '-1'.
# This is what marks a header as a hot-call — SELECT_ASNSeq's `<> -1` filter excludes these from the
# morning-create idempotency lock, and the 856 filename branch (_filename_856) keys on it for '8HC...'.
HOTCALL_SEQ = "-1"


def create_hotcall_asn(line, prodDate, manifest, items, site=1, database=None):
    """Create one HOT-CALL ASN header + its manual detail rows in a single Inventory transaction.

    Reproduces the legacy "One Cycle Entry" Add (HotCallEntry.pas:136-312) with the M1 decisions:
      * EIN-at-SEND (P14): write IN_ASN_EIN=0 here; the real per-site EIN is allocated from
        INV_SITES.IN_EIN_SEQ at 856 send (edi856.send_856) — NOT SiteEIN+1-at-create. So the legacy
        AD_UpdateEIN counter bump in the create tx (:290-292) is DROPPED here: the EIN-at-send model owns
        the per-site counter, so bumping it at create would double-allocate. (hotcall-coverage-analysis.md
        §4: 'AD_UpdateEIN-equivalent ... already done by the EIN-at-send alloc'.)
      * Header IN_QTY = SUM of the detail qtys (P14), fixing the @QTY garbage (:246-247).
      * Status 'C' (created), VC_ASSEMBLY_LINE='' (:236-237), VC_START_SEQ_NUMBER=VC_END_SEQ_NUMBER='-1'
        (:238-243), the operator-typed 8-char manifest + line + production date.
      * @HotCall=1 on every INSERT_ASNDetail (:279-280) — the always-INSERT branch (no dedup), preserved
        by the Q1 re-key (spike-asndetail-rekey.sql:85-90).

    NO AD_FRSPULL, NO SELECT_ForecastDetailBCASN, NO computeAsnDetails, NO rounding — the operator's parts
    + qtys go in verbatim. This is the thinnest producer in the suite.

    Parameters
    ----------
    line     : VC_LINE_NAME (the operator's line pick, from AD_GetLines — HotCallEntry.pas:101/112).
    prodDate : 'yyyymmdd' production date (the operator's ASN-date, formatdatetime('yyyymmdd', ...) :248-249).
    manifest : the operator-typed 8-char manifest (validated: length >= 8, non-'7' broadcast).
    items    : list of (part, qty) pairs (or {'part','qty'} dicts) — the manual entries.
    site     : INV_SITES.IN_SITE_ID owning the EIN sequence (multi-site D1: derives from the session,
               NEVER a client param; the EIN itself is allocated at SEND, not here). Default 1.
    database : Inventory connection name (defaults to DATABASE).

    Returns a dict: {'asnId', 'details', 'qty', 'manifest', 'line', 'prodDate'}.

    Validation (BEFORE any write — the legacy guards all `exit` before BeginTrans, :143-218 < :221):
      * manifest length >= 8 and non-'7' (broadcast) — _validate_manifest.
      * each item: numeric qty > 0, part required — _validate_item (inside computeHotcallDetails).
      A HotcallError raised here writes NOTHING (the transaction is not yet open).

    Transaction order (ONE tx, all-or-nothing, the M1 pattern):
      1. INSERT_ASNInfo (status 'C', @StartSeq='-1', @EndSeq='-1', @AssyLine='', @Qty=<detail sum>,
         @Ein=0, OUTPUT @ASNID) — captured via the M1 DECLARE/EXEC ... @ASNID=@id OUTPUT/SELECT @id pattern
         on the SAME tx (verbatim from asn.create_asn).
      2. per detail: INSERT_ASNDetail @ASNID, @EIN=0, @Manifest, @PartNumber, @Qty, @HotCall=1.
      commit. On ANY error: rollback + re-raise (the partial header unwinds with the details).
    """
    db = database if database is not None else DATABASE
    log = system.util.getLogger("SPIKE.create_hotcall_asn")

    # --- VALIDATION (pre-transaction; a raise here writes nothing) -------------------------------
    manifest = _validate_manifest(manifest)
    details = computeHotcallDetails(items, manifest)   # validates each item; raises before any write
    headerQty = _detailQtySum(details)                 # P14 — header IN_QTY = sum of detail qtys

    prodDate = str(prodDate)
    if len(prodDate) != 8:
        raise HotcallError("create_hotcall_asn: prodDate must be yyyymmdd (8 chars), got %r" % (prodDate,))

    # --- ONE Inventory transaction: header + details (the M1 write seam) -------------------------
    tx = system.db.beginTransaction(db)
    try:
        # INSERT_ASNInfo — the M1 OUTPUT-capture pattern (asn.create_asn step 5, verbatim). @ASNID is an
        # OUTPUT param, NOT an auto key: SCOPE_IDENTITY after a bare EXEC is NULL (the proc's INSERT is a
        # child scope), and createSProcCall+registerOutParam can't join a gateway transaction. So DECLARE
        # @id + EXEC ... @ASNID=@id OUTPUT + SELECT @id, read back via runScalarPrepQuery on the SAME tx.
        # @StartSeq/@EndSeq are varchar(4) in the proc (verified live) -> pass the '-1' STRING sentinel.
        # @Ein=0 (EIN-at-send, P14). @Qty = the detail sum (P14, fixes the @QTY garbage).
        asnId = system.db.runScalarPrepQuery(
            "DECLARE @id int; "
            "EXEC INSERT_ASNInfo @ASNID=@id OUTPUT, @LineName=?, @AssyLine=?, @StartSeq=?, "
            "@DTStartSeq=?, @EndSeq=?, @DTEndSeq=?, @Qty=?, @PDate=?, @Ein=?; "
            "SELECT @id",
            [line, "", HOTCALL_SEQ, prodDate, HOTCALL_SEQ, prodDate, int(headerQty), prodDate, 0],
            db, tx)
        if asnId is None:
            raise HotcallError("create_hotcall_asn: INSERT_ASNInfo returned no ASN id")
        asnId = int(asnId)

        # INSERT_ASNDetail per manual row — @HotCall=1 (always-INSERT, no dedup). @EIN=0 (at-send).
        for d in details:
            system.db.runPrepUpdate(
                "EXEC INSERT_ASNDetail @ASNID=?, @EIN=?, @Manifest=?, @PartNumber=?, @Qty=?, @HotCall=?",
                [asnId, 0, d["manifest"], d["partNumber"], int(d["qty"]), 1], db, tx=tx)
            d["asnId"] = asnId

        system.db.commitTransaction(tx)
    except Exception:
        # roll back the partial header (+ any details) — all-or-nothing, the M1 atomicity pattern.
        system.db.rollbackTransaction(tx)
        raise
    finally:
        system.db.closeTransaction(tx)

    log.info("create_hotcall_asn line=%s prodDate=%s manifest=%s -> ASN %s, %d detail rows, qty=%d "
             "(EIN allocated at SEND)" % (line, prodDate, manifest, asnId, len(details), headerQty))
    return {"asnId": asnId, "details": details, "qty": headerQty, "manifest": manifest,
            "line": line, "prodDate": prodDate}
