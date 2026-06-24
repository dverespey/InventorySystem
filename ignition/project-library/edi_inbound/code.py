# edi_inbound — Project Library: the M1 INBOUND EDI processor (997 functional ack + 824 application
# advice), the loop-closer for the outbound 856/810 we ship. Producer-recipe pattern, the twin of
# edi856/code.py and edi810/code.py: PURE Jython-2.7/CPython parsers + a thin gateway DRIVER.
#
# SOURCE TRUTH (all on disk, cited per function):
#   docs/analysis/edi/inbound/997-824-inbound-spec.md          — the behavioral spec (THE target)
#   docs/analysis/edi/inbound/einstatus-status-flow-analysis.md — the DB-side status flow + the AK map
#   EDIUpload.pas (LIVE, InventorySystem.dpr:53)               — the legacy parser:
#       997 branch :186-252 (AK1 loop :192, EDIType=copy(fcl,5,2):194, EIN=copy(fcl,8,9):195,
#                            Status=copy(fcl,5,1) of the NEXT line :197)
#       824 branch :253-305 (NTE loop :279-291, errorText=copy(fcl,9,50):283, manifest=copy(fcl,60,8):284,
#                            part=copy(fcl,69,12):285)
#       ISA/DUNS  :64-88     (splitString('*',...):71, delSL[4] trading-partner DUNS :73-78, ST id
#                            data=copy(fcl,4,3):88)
#   UPDATE_EINStatus (/tmp/inv_utf8.sql:1711; CreateInventory.sql:1711-1730)
#
# LOCKED DECISIONS realized here (Q6/Q10/Q11 — already approved):
#   Q6   997 AK9 A/E/P/R -> DISTINCT statuses (NOT the legacy binary A-vs-rest collapse, which stored
#        the raw char so E/P then render BLANK in the status CASEs). ak9_to_status: A->A, E->E, P->P,
#        R->R, anything else (M/W/X/garbage) -> R. Tolerate AK2/AK3/AK4 between AK1 and AK9 (read the
#        REAL AK9, not "the next line" — fixes the legacy fragility at EDIUpload.pas:196-197).
#   Q10  824 -> auto-flag the referenced ASN(s) Rejected ('R', D3 recoverable) + write an alarm/reject
#        record (manifest/part/errorText x N + a home-hub-surfaced flag) for the main-screen alarm +
#        click-to-detail. (Legacy flagged NOTHING — Excel report only.)
#   Q11  scheduled gateway poll (NOT the legacy manual button) + DUNS guard. A no-DUNS-match file is
#        QUARANTINED (moved aside + logged), NOT silently dropped/left-to-rescan (legacy :439).
#   Site-scoped status updates: INV_ASN_MST / INV_INV_MST have NO IN_SITE_ID column yet -> the site
#        predicate is a `-- M4` marker (same re-key convention as spike-asndetail-rekey.sql). Today the
#        WHERE is EIN+type only, but the resolved site is THREADED through so the M4 add is one line.
#
# BUGS WE DO **NOT** CARRY (spec §5.4/§6 + the status-flow analysis §4 traps):
#   * EIN-carryover  (legacy resets EIN once at Execute top, not per file -> a 824 after a 997 archives
#                     under the stale 997 EIN). We parse each file FRESH; no shared mutable EIN state.
#   * no-archive re-ingest (a failed move leaves the file to be re-processed next run). We track
#                     processed files in a DB ledger so a re-run/re-drop is a NO-OP.
#   * P12 #8 wrong-target retry (UpdateEINStatus on error fires UpdateRecProdRejInfo against
#                     INV_REJECT_INF). We never do that; an ack failure rolls back the ACK only.
#   * no-@@ROWCOUNT silent-success (UPDATE_EINStatus no-ops on an unknown EIN yet reports success). We
#                     CHECK @@ROWCOUNT (rowsUpdated): 0 rows -> an alarm, NOT a silent success.
#
# Jython 2.7 (8.1 and 8.3, no delta — pure string slicing + system.db/system.file; no Py3-only libs).
#
# This module holds BOTH:
#   * parse_997 / parse_824 / ak9_to_status — PURE. Import nothing, reference no gateway global at import
#     time -> CPython-importable for the unit test (scripts/e2e/test_edi_inbound_build.py). No DB, no I/O.
#   * process_997 / process_824 / process_inbound — the GATEWAY DRIVER. References `system` only at CALL
#     time (like edi856.send_856), so importing the module never needs the gateway runtime; it is driven
#     headlessly through scripts/e2e/jython_shim.py (the R8 system.db/system.file shim).


# =================================================================================================
# PURE parsers + the AK9 status map
# =================================================================================================

class EdiInboundError(Exception):
    """Raised on a structural parse/processing error. The driver rolls back the per-file tx on it
    (the file is NOT marked processed, so a fixed re-drop can re-run it)."""
    pass


def _seg_id(line):
    """The 3-char segment id = the legacy copy(fcl,1,3) (EDIUpload.pas:191,247,277,290). Pascal copy is
    1-based start 1, len 3 -> Python line[0:3]. Returns '' for a short/blank line (so a trailing blank
    line in the file ends a loop cleanly rather than IndexError-ing)."""
    return (line or "")[0:3]


def _is_se(line):
    """True for the SE transaction-set trailer segment. The segment id is 'SE' followed by the element
    separator (an 'SE*...' line). Matching on the 2-char 'SE' prefix + a separator/space avoids treating
    some other 'SE...'-prefixed value as the trailer while still catching 'SE*<count>*<ctrl#>'."""
    s = line or ""
    return s[0:2] == "SE" and (len(s) < 3 or s[2] in ("*", " ", "\t"))


def ak9_to_status(code):
    """Q6 — map an X12 AK9/AK5 acknowledgment code to the rebuild status char. A STRICT SUPERSET of the
    legacy (which only branched A-vs-rest for the history text and stored the raw char verbatim):

        A  Accepted               -> 'A'
        E  Accepted, with errors  -> 'E'   (legacy stored 'E' but the status CASE renders it BLANK)
        P  Partially accepted     -> 'P'   (legacy stored 'P' -> BLANK)
        R  Rejected               -> 'R'
        anything else (M/W/X auth/content failures, garbage, empty) -> 'R'  (rejected, the safe default)

    `code` is the single AK901/AK501 char (already extracted by parse_997). Case-normalized to upper
    (TEMA emits upper; defensive). A multi-char input takes the first char (an AK9 code is one char).

    !!! REQUIRED M1 FOLLOW-ON (adversary SHOULD-FIX-3 — the E/P fix is HALF-DONE until this lands):
    'E' and 'P' are STORED distinctly here, but the only status-render decode that exists anywhere today
    (legacy SELECT_ASNStatus CreateInventory.sql:1955-1959; invoice :3254-3257) maps ONLY A/S/C/R — so a
    stored 'E' or 'P' renders BLANK downstream, which is EXACTLY the legacy bug Q6 set out to fix. There
    is NO rebuild status-render NQ/view on disk yet to extend (repo-wide search found none), so this is
    recorded as a HARD REQUIREMENT, not silently dropped:

        WHEN the ASN/invoice status-DISPLAY NQ/view is built, it MUST add decode arms:
            'E' -> "Accepted w/ errors"   'P' -> "Partial"   ('R' -> "Rejected" must be present too)
        alongside the existing A/S/C arms (einstatus-status-flow-analysis.md §3; spec §2.5 + the
        "required follow-on" note). Without these arms the Q6 fix is incomplete.

    Confirm against a golden whether TEMA ever sends E/P to this supplier; Live is 100% 'A', so until a
    golden arrives the E/P path is exercised only via synthetic fixtures (the gap is latent on parity
    data but real the moment TEMA sends an E/P)."""
    if code is None:
        return "R"
    c = str(code).strip().upper()
    if not c:
        return "R"
    c = c[0]
    if c in ("A", "E", "P", "R"):
        return c
    return "R"          # M/W/X and anything unrecognized -> Rejected (safe default)


# Functional-group id -> the EINType the UPDATE routes on (the legacy @EINType). 'SH' -> ASN (856),
# anything else (incl. 'IN') -> invoice (810), matching UPDATE_EINStatus's `if @EINType='SH' ... else`.
def _functional_to_eintype(fnId):
    """Normalize the AK1 functional-id element. Legacy reads copy(fcl,5,2) = chars 5-6 (2 chars) and
    passes it straight as @EINType; UPDATE_EINStatus tests `= 'SH'` (else = invoice). We keep that exact
    two-branch semantics: return 'SH' for the ship-notice group, else 'IN' (so the driver routes
    SH->ASN, everything-else->invoice, byte-for-byte the legacy branch)."""
    s = ("" if fnId is None else str(fnId)).strip().upper()
    return "SH" if s == "SH" else "IN"


def _lines(text):
    """Split an EDI file body into segment lines, dropping trailing CR and empty lines. TEMA writes one
    X12 segment per physical line (the legacy Readln walk relies on this — EDIUpload.pas reads line by
    line, never on a '~' segment terminator). We split on \\n and strip \\r so CRLF and LF both work."""
    if text is None:
        return []
    return [ln.rstrip("\r") for ln in text.split("\n")]


def parse_997(text):
    """PURE 997 functional-ack parser. Returns a list of acks, ONE per AK1 functional group:
        {'functionalCode': 'SH'|'IN', 'ein': int, 'ak9': <raw AK9/AK5 char or ''>}

    Reproduces the legacy AK1 loop (EDIUpload.pas:190-248) by the SAME fixed element offsets (the parity
    oracle, spec §2.3) — but FIXES the legacy fragility: legacy reads the status from char 5 of the
    *immediately following* line with no AK9 guard (:196-197), so an interleaved AK2/AK3/AK4 detail
    segment would make it read the WRONG char. We instead, after each AK1, scan FORWARD to the real
    'AK9' segment (tolerating AK2/AK3/AK4 between them) and read AK901 = char 5 there. If no AK9 appears
    before the next AK1 / SE / EOF, ak9 is '' (-> ak9_to_status -> 'R', a safe reject for a malformed
    group rather than a false accept).

    Element offsets (1-based copy -> 0-based slice), EDIUpload.pas:194-197 + spec §2.3:
      AK1 functionalCode = copy(fcl,5,2)  = chars 5-6  -> line[4:6]   (group functional id 'SH'/'IN')
      AK1 ein            = copy(fcl,8,9)  = chars 8-16 -> line[7:16]  (group control # == our EIN, %09d)
      AK9 ak9            = copy(fcl,5,1)  = char 5      -> line[4:5]   (AK901 ack code)
    The offsets assume TEMA's AK1 = 'AK1*<2-char fnid>*<9-digit ctrl#>' (fnid at 4:6, '*' at 6, ctrl#
    at 7:16). This is data-dependent vs a real golden 997 (spec §2.3 #1 — the #1 thing to confirm). The
    parse is offset-faithful to the legacy AND segment-aware for the AK9 (the deliberate fix).

    Parameters
    ----------
    text : the full 997 file body (str). Already DUNS-guarded + envelope-checked by the driver.

    Returns the list of ack dicts in file order. An empty list means no AK1 groups were found (an
    envelope-only / non-997 body) — the driver treats that as nothing-to-do, not an error."""
    lines = _lines(text)
    acks = []
    i = 0
    n = len(lines)
    while i < n:
        if _seg_id(lines[i]) == "AK1":
            ak1 = lines[i]
            fnId = ak1[4:6]                          # copy(fcl,5,2) chars 5-6
            einRaw = ak1[7:16].strip()               # copy(fcl,8,9) chars 8-16
            try:
                ein = int(einRaw)
            except (ValueError, TypeError):
                raise EdiInboundError(
                    "parse_997: AK1 group control number (EIN) is not an integer: %r "
                    "(segment %r)" % (einRaw, ak1))
            # Scan forward to the REAL AK9 (tolerate AK2/AK3/AK4 detail between AK1 and AK9 — the fix).
            # Stop at the next AK1 (next group), SE (trailer), or EOF.
            ak9 = ""
            j = i + 1
            while j < n:
                sid = _seg_id(lines[j])
                if sid == "AK9":
                    ak9 = lines[j][4:5]              # copy(fcl,5,1) AK901
                    break
                if sid == "AK1" or _is_se(lines[j]):
                    break                           # next group / trailer -> no AK9 for this group
                j += 1                              # tolerate AK2/AK3/AK4 detail (the legacy fix)
            acks.append({
                "functionalCode": _functional_to_eintype(fnId),
                "ein": ein,
                "ak9": ak9,
            })
            # advance to where the AK9 scan stopped (so we don't re-read AK2/3/4 as nothing); if we hit
            # AK9 we resume after it, else resume at j (the next AK1/SE we found).
            i = (j + 1) if (j < n and _seg_id(lines[j]) == "AK9") else j
        else:
            i += 1
    return acks


def parse_824(text):
    """PURE 824 application-advice parser. Returns:
        {'refs': [{'manifest': <str>, 'part': <str>, 'errorText': <str>}, ...], 'count': <int>}

    Reproduces the legacy NTE loop (EDIUpload.pas:279-291): walk lines until the SE trailer, and for each
    'NTE' segment extract the three FIXED-WIDTH fields TEMA packs (spec §3.1 — the parity oracle):
      errorText = copy(fcl,9,50)  = chars 9-58  -> line[8:58]   (50 chars, the free-text reason)
      manifest  = copy(fcl,60,8)  = chars 60-67 -> line[59:67]  (8 chars, the join key to INV_ASN_MST)
      part      = copy(fcl,69,12) = chars 69-80 -> line[68:80]  (12 chars)
    Each field is .strip()-ed (the legacy wrote the fixed-width slice verbatim into Excel; the rebuild
    stores trimmed values for a clean join/display — the manifest join needs the bare 8-char id).

    The loop ends at the first segment whose id is 'SE' (legacy `while data <> 'SE*'`, :279 — but legacy
    compared the 3-char copy to the 4-char literal 'SE*' which NEVER equals it, so the legacy loop only
    actually terminated at EOF/exception; we end on the real SE trailer OR EOF, which is the intended
    behavior). 'count' = len(refs) (legacy logged x-3 = the NTE count, :301).

    These fixed offsets are the single most fragile thing in the 824 path (any width drift silently
    mis-columns every field — spec §3.1). Data-dependent vs a real golden 824. We keep the legacy byte
    map exactly as the oracle.

    Parameters
    ----------
    text : the full 824 file body (str). Already DUNS-guarded + envelope-checked by the driver.

    Returns the refs dict. An 824 with zero NTE lines yields {'refs': [], 'count': 0} (the driver then
    has nothing to flag — it still marks the file processed, but raises no alarm)."""
    lines = _lines(text)
    refs = []
    for ln in lines:
        if _is_se(ln):
            break                                    # SE trailer ends the transaction set
        if _seg_id(ln) == "NTE":
            refs.append({
                "errorText": ln[8:58].strip(),       # copy(fcl,9,50)  chars 9-58
                "manifest": ln[59:67].strip(),       # copy(fcl,60,8)  chars 60-67
                "part": ln[68:80].strip(),           # copy(fcl,69,12) chars 69-80
            })
    return {"refs": refs, "count": len(refs)}


def _split_isa(line):
    """Split the ISA line on '*' the SAME way the legacy splitString helper does (EDIUpload.pas:473-493):
    one token per '*', the text BEFORE each separator, starting delSL[0]='ISA'. Python str.split('*')
    yields exactly that (plus a trailing token after the last '*'), so delSL[4] == split('*')[4] = the
    5th element. Returns the list. The DUNS guard reads index 4 — authoritative-BY-CODE (spec §5.2):
    the exact X12 element index 4 lands on is data-dependent vs a real inbound ISA; the rebuild
    replicates the index, not the X12 label."""
    return (line or "").split("*")


def _detect_type(lines):
    """Return the transaction-set id = the legacy data:=copy(fcl,4,3) read from LINE 3 (the ST segment),
    EDIUpload.pas:86-88. Pascal copy(fcl,4,3) = chars 4-6 -> line[3:6] of 'ST*997' -> '997'. The legacy
    Readln's line1 (ISA), line2 (GS), line3 (ST) then slices line3. We find the ST segment (line index 2
    in a well-formed file; we scan for the first 'ST' segment to be robust to an extra envelope line)."""
    for ln in lines:
        if ln[0:2] == "ST" and (len(ln) < 3 or ln[2] in ("*", " ", "\t")):
            return ln[3:6]                            # copy(fcl,4,3) -> chars 4-6
    return ""


# =================================================================================================
# THE GATEWAY DRIVER — process_997 / process_824 / process_inbound. ONE transaction per file.
# =================================================================================================
#
# IG81-COMPAT: every gateway API used here (runPrepQuery / runScalarPrepQuery(...,tx=) / runPrepUpdate
# (...,tx=) / beginTransaction / system.file / getLogger) is identical on 8.1.52 and 8.3 — no version
# guard needed. The SCHEDULED POLL host (Q11) is a gateway TIMER/event script on 8.1 (NOT 8.3 Event
# Streams) — IG83-TODO below.

from db_shared import CONNECTION as DATABASE   # centralized DB-conn name (default Inventory_Spike; single prod-rename point)


def _resolve_site_by_duns(db, duns, tx):
    """DUNS guard (Q11): resolve an inbound file's ISA delSL[4] to a site by the trading-partner DUNS.
    Under D1 the legacy ALC-DB AD_GetSiteTMMDUNS allow-list collapses into a per-site DUNS attribute on
    INV_SITES. delSL[4] is matched against VC_TMM_DUNS (the TEMA/TMM trading-partner DUNS — the receiver
    DUNS the legacy validated, edi-upload.md §4.2). Returns the IN_SITE_ID, or None if no site matches
    (-> the driver QUARANTINES the file; NOT silently dropped, unlike legacy :439).

    NB — UNPROVABLE WITHOUT A GOLDEN INBOUND ISA (adversary BLOCKER-1 / KEY FINDING #2): the rebuild
    matches VC_TMM_DUNS, which is LEGACY-FAITHFUL (legacy AD_GetSiteTMMDUNS matches delSL[4] against
    Site.SiteTMMDUNS — the same column semantics). But the X12 ELEMENT delSL[4] lands on for a REAL
    inbound TEMA ISA is NOT known from available data. On OUR OWN OUTBOUND ISA (EDI856Object.pas/
    EDI810Object.pas) delSL[4] = ISA04 = OUR SiteDUNS, while the TMM DUNS is at delSL[8]/ISA08 — so if
    TEMA's inbound ISA is X12-standard, delSL[4] would NOT carry the TMM DUNS and this guard would
    quarantine every TEMA file (same latent risk in the legacy). The CODE is faithful; the ELEMENT is
    PENDING A CAPTURED GOLDEN INBOUND 997/824 ISA (same honest stance as the 856/810 byte-parity gap).
    The spike INV_SITES carries placeholder DUNS values."""
    if duns is None:
        return None
    rows = system.db.runPrepQuery(
        "SELECT IN_SITE_ID FROM INV_SITES WHERE VC_TMM_DUNS = ?", [str(duns).strip()], db, tx)
    if not len(rows):
        return None
    return int(rows[0]["IN_SITE_ID"])


def _already_processed(db, fileName, fileHash, tx):
    """Idempotency ledger check (kills the legacy re-ingest hazard, spec §5.4/§8.8). CONTENT IS THE
    IDENTITY: a file is processed once PER CONTENT HASH, regardless of filename. We key on VC_FILE_HASH
    ALONE (NOT name+hash) so:
      * identical content re-delivered under a NEW name (VAN/mailer re-stamps the filename on a
        redelivery — common) is a NO-OP, not a re-process. The old name-AND-hash key let a renamed
        redelivery slip through -> for an 824 that meant DUPLICATE main-screen alarms + a redundant
        'R' re-flip (adversary SHOULD-FIX-1, live-proven). Content-only dedup closes that hole.
      * a DIFFERENT-content file that happens to REUSE a name IS reprocessed (different hash -> no
        ledger hit) -> correct: the content is new, so it is a new event. (Decided: content is the
        identity; a same-name/new-content file MUST reprocess.)
    The ledger still RECORDS fileName (audit / archive traceability), but the DEDUP key is the hash.
    `fileName` is kept in the signature for that audit record; it is intentionally not part of the seek.
    Returns True if a ledger row already records this CONTENT as successfully processed."""
    rows = system.db.runPrepQuery(
        "SELECT 1 FROM INV_EDI_INBOUND_LOG WHERE VC_FILE_HASH = ? "
        "AND VC_STATUS = 'PROCESSED'", [str(fileHash)], db, tx)
    return len(rows) > 0


def _record_processed(db, fileName, fileHash, ediType, siteId, status, detail, tx):
    """Write the processed-files ledger row INSIDE the per-file tx (so it commits atomically with the
    status flips / reject rows — a re-run can never see a 'PROCESSED' row for a file whose effects rolled
    back). status is 'PROCESSED' (success) or 'QUARANTINED' (no DUNS match / not-EDI). siteId may be NULL
    for a quarantined file (no site resolved)."""
    system.db.runPrepUpdate(
        "INSERT INTO INV_EDI_INBOUND_LOG "
        "(VC_FILE_NAME, VC_FILE_HASH, VC_EDI_TYPE, IN_SITE_ID, VC_STATUS, VC_DETAIL, DT_PROCESSED) "
        "VALUES (?, ?, ?, ?, ?, ?, GETDATE())",
        [str(fileName), str(fileHash), ediType, siteId, status, (detail or "")[:400]], db, tx=tx)


def _hash_text(text):
    """A stable content hash = THE ledger dedupe key. CONTENT is the file identity: the same content
    re-dropped is detected even under a DIFFERENT NAME (so a renamed redelivery is a no-op), and a
    same-NAME-but-changed file is NOT falsely skipped (different content -> different hash -> a new
    event). _already_processed seeks on this hash alone (NOT name+hash). Jython-2.7 safe (hashlib is in
    the stdlib on both 8.1 and 8.3 Jython). Hex digest, 64 chars (fits varchar(64))."""
    import hashlib
    return hashlib.sha256((text or "").encode("utf-8")).hexdigest()


def _apply_997_ack(db, ack, siteId, tx):
    """Apply ONE 997 ack: the site-scoped UPDATE_EINStatus, in-line (we wrap the proc's two-branch UPDATE
    rather than EXEC the proc, so we can CHECK @@ROWCOUNT and add the site predicate — the proc has
    neither). Returns rowsUpdated (so the driver can alarm on an unknown EIN = 0 rows).

    Routing (UPDATE_EINStatus CreateInventory.sql:1722-1728): functionalCode 'SH' -> INV_ASN_MST/
    VC_ASN_STATUS; else -> INV_INV_MST/VC_INV_STATUS. Status = ak9_to_status(ack['ak9']) (Q6 — distinct
    A/E/P/R; NOT the raw char). Site scoping is a -- M4 marker (the tables have no IN_SITE_ID yet); the
    resolved siteId is threaded so the M4 add is one line in each WHERE."""
    status = ak9_to_status(ack["ak9"])
    ein = int(ack["ein"])
    if ack["functionalCode"] == "SH":
        # 856/ASN ack. -- M4: add  AND IN_SITE_ID = ?  (params [status, ein, siteId]) once the column exists.
        rows = system.db.runPrepUpdate(
            "UPDATE INV_ASN_MST SET VC_ASN_STATUS = ? WHERE IN_ASN_EIN = ?"
            "  /* -- M4: AND IN_SITE_ID = %d */" % (siteId,),
            [status, ein], db, tx=tx)
    else:
        # 810/invoice ack (the else-default, incl. 'IN'). -- M4: add  AND IN_SITE_ID = ?  here too.
        rows = system.db.runPrepUpdate(
            "UPDATE INV_INV_MST SET VC_INV_STATUS = ? WHERE IN_INV_EIN = ?"
            "  /* -- M4: AND IN_SITE_ID = %d */" % (siteId,),
            [status, ein], db, tx=tx)
    return int(rows) if rows is not None else 0


def process_997(text, fileName, database=None, fileHash=None, siteId=None, tx=None):
    """Process ONE 997 file body (already DUNS-guarded — siteId resolved by the caller). For each AK1
    group: apply the site-scoped status flip via _apply_997_ack and CHECK @@ROWCOUNT. An unknown EIN
    (0 rows) raises an ALARM record (a reject row, IN_EDI_ALARM_REJ) — it does NOT silently succeed
    (the legacy no-@@ROWCOUNT bug, §8). Idempotent: applying the same ack twice writes the same status.

    Runs INSIDE the per-file tx the caller (process_inbound) opened, so all acks + alarms + the ledger
    row commit atomically. Returns a summary dict: {'acks': [...], 'applied': N, 'unknownEins': [...]}.

    If called standalone (tx=None, for a unit/integration harness), it opens + commits its own tx and
    resolves the site from siteId (required)."""
    db = database if database is not None else DATABASE
    log = system.util.getLogger("SPIKE.process_997")
    if fileHash is None:
        fileHash = _hash_text(text)
    acks = parse_997(text)

    ownTx = tx is None
    if ownTx:
        if siteId is None:
            raise EdiInboundError("process_997: siteId is required when called without a tx")
        tx = system.db.beginTransaction(db)
    applied = 0
    unknownEins = []
    try:
        for ack in acks:
            rows = _apply_997_ack(db, ack, siteId, tx)
            if rows > 0:
                applied += 1
            else:
                # unknown EIN: 0 rows updated. Alarm — NOT a silent success (legacy bug §8). Record a
                # reject/alarm row the home hub surfaces; the operator sees "ack for unknown EIN".
                unknownEins.append({"ein": ack["ein"], "functionalCode": ack["functionalCode"]})
                _write_alarm(db, siteId, "997_UNKNOWN_EIN",
                             "997 ack for unknown %s EIN %d (0 rows updated)"
                             % (ack["functionalCode"], ack["ein"]),
                             manifest=None, part=None, errorText=None,
                             ein=ack["ein"], tx=tx)
        if ownTx:
            _record_processed(db, fileName, fileHash, "997", siteId, "PROCESSED",
                              "%d acks, %d applied, %d unknown" % (len(acks), applied, len(unknownEins)),
                              tx)
            system.db.commitTransaction(tx)
    except Exception:
        if ownTx:
            system.db.rollbackTransaction(tx)
        raise
    finally:
        if ownTx:
            system.db.closeTransaction(tx)

    log.info("process_997 file=%s site=%s -> %d acks, %d applied, %d unknown-EIN alarms"
             % (fileName, siteId, len(acks), applied, len(unknownEins)))
    return {"acks": acks, "applied": applied, "unknownEins": unknownEins, "fileHash": fileHash}


def _write_alarm(db, siteId, alarmType, summary, manifest, part, errorText, ein, tx):
    """Write ONE alarm/reject row (IN_EDI_ALARM_REJ) — the DATA/flag side of the main-screen alarm +
    click-to-detail (Q10). The real gateway NATIVE alarm pipeline (system.alarm / an alarm journal tag)
    is the PROD step; here we implement the DATA so it is testable headlessly, and a home-hub NQ surfaces
    BIT_RESOLVED=0 rows as the main-screen alarm. (Prod follow-on: wire system.alarm off these rows.)

    BIT_RESOLVED defaults 0 (active/unacknowledged). alarmType is '997_UNKNOWN_EIN' or '824_ASN_REJECT'.
    One row per (manifest, part, errorText) for an 824 -> the click-to-detail shows every reject line."""
    system.db.runPrepUpdate(
        "INSERT INTO INV_EDI_ALARM_REJ "
        "(IN_SITE_ID, VC_ALARM_TYPE, VC_SUMMARY, VC_MANIFEST_NUMBER, VC_ASSY_PART_NUMBER, "
        " VC_ERROR_TEXT, IN_EIN, BIT_RESOLVED, DT_RAISED) "
        "VALUES (?, ?, ?, ?, ?, ?, ?, 0, GETDATE())",
        [siteId, alarmType, (summary or "")[:200], manifest, part,
         (errorText or "")[:200] if errorText is not None else None,
         ein], db, tx=tx)


def process_824(text, fileName, database=None, fileHash=None, siteId=None, tx=None):
    """Process ONE 824 application-advice file body (already DUNS-guarded). Per Q10:
      1. parse_824 -> the (manifest, part, errorText) reject lines.
      2. For each DISTINCT manifest, flip the referenced ASN(s) to 'R' (recoverable reject, D3): match
         the manifest on INV_ASN_DETAIL_MST.VC_MANIFEST_NUMBER -> parent INV_ASN_MST (the manifest lives
         on the DETAIL row, NOT the header — confirmed against sys.columns). Site-scoped is a -- M4
         marker. The 824 is a REAL status event (legacy did NOT flip — left the row stuck at 'S', the 6
         stuck-S invoices on Live; einstatus-status-flow-analysis.md §4 trap).
      3. Write an alarm/reject row PER reject line (IN_EDI_ALARM_REJ) capturing manifest/part/errorText,
         for the main-screen alarm + click-to-detail.

    INTENDED SCOPE — FLAG EVERY ASN CARRYING THE REJECTED MANIFEST (adversary SHOULD-FIX-2, spec §3.2):
    a single manifest can legitimately appear on MORE THAN ONE ASN's detail rows (proven on Live: a
    manifest maps to up to 3 distinct IN_ASN_ID; e.g. manifest 52066074 -> 7 detail rows across 3 ASNs).
    The flip UPDATE is a SET-based join, so it flips ALL ASNs whose detail carries the rejected manifest
    in one statement -> asnsFlagged is the @@ROWCOUNT (can be 2 or 3). THIS IS DELIBERATE: a reject is
    recoverable (D3), so OVER-FLAGGING (flag every shipment that carried the rejected manifest) is the
    SAFE choice — missing one would leave a genuinely-rejected ASN looking accepted. The alarm rows
    capture the reject per LINE regardless of how many ASNs matched. Site scoping (-- M4) will further
    narrow this to the resolving site's ASNs once INV_ASN_MST carries IN_SITE_ID.

    Runs INSIDE the per-file tx (or opens its own if tx=None). Returns a summary dict:
    {'count': N, 'manifestsFlagged': [...], 'asnsFlagged': N (TOTAL ASN rows flipped, may be >1 per
    manifest), 'refs': [...]}."""
    db = database if database is not None else DATABASE
    log = system.util.getLogger("SPIKE.process_824")
    if fileHash is None:
        fileHash = _hash_text(text)
    parsed = parse_824(text)
    refs = parsed["refs"]

    ownTx = tx is None
    if ownTx:
        if siteId is None:
            raise EdiInboundError("process_824: siteId is required when called without a tx")
        tx = system.db.beginTransaction(db)
    asnsFlagged = 0
    manifestsFlagged = []
    try:
        # 2. flip each referenced ASN to 'R' (one UPDATE per distinct manifest -> the parent ASN).
        seen = set()
        for ref in refs:
            manifest = ref["manifest"]
            if not manifest or manifest in seen:
                continue
            seen.add(manifest)
            # FLAG EVERY ASN carrying this manifest (intended over-flag, SHOULD-FIX-2): the set-based
            # join flips ALL matching ASNs in one statement -> rows can be >1 (Live: up to 3 ASNs per
            # manifest). A reject is recoverable (D3); over-flagging is safer than missing a shipment.
            # -- M4: add  AND a.IN_SITE_ID = ?  (the resolved siteId) once INV_ASN_MST has the column.
            rows = system.db.runPrepUpdate(
                "UPDATE a SET a.VC_ASN_STATUS = 'R' "
                "FROM INV_ASN_MST a JOIN INV_ASN_DETAIL_MST d ON a.IN_ASN_ID = d.IN_ASN_ID "
                "WHERE d.VC_MANIFEST_NUMBER = ?"
                "  /* -- M4: AND a.IN_SITE_ID = %d */" % (siteId,),
                [manifest], db, tx=tx)
            rows = int(rows) if rows is not None else 0
            asnsFlagged += rows
            manifestsFlagged.append({"manifest": manifest, "asnRows": rows})
            if rows == 0:
                # an 824 referencing a manifest we don't have -> alarm (the operator must still see the
                # reject text; the ASN match failing is itself notable). NOT a silent drop.
                log.warn("process_824 file=%s: manifest %s matched no ASN (0 rows flagged)"
                         % (fileName, manifest))
        # 3. one alarm/reject row per reject LINE (the click-to-detail payload).
        for ref in refs:
            _write_alarm(db, siteId, "824_ASN_REJECT",
                         "824 ASN reject: manifest %s part %s" % (ref["manifest"], ref["part"]),
                         manifest=ref["manifest"], part=ref["part"], errorText=ref["errorText"],
                         ein=None, tx=tx)
        if ownTx:
            _record_processed(db, fileName, fileHash, "824", siteId, "PROCESSED",
                              "%d reject lines, %d ASN rows flagged R" % (len(refs), asnsFlagged), tx)
            system.db.commitTransaction(tx)
    except Exception:
        if ownTx:
            system.db.rollbackTransaction(tx)
        raise
    finally:
        if ownTx:
            system.db.closeTransaction(tx)

    log.info("process_824 file=%s site=%s -> %d reject lines, %d distinct manifests, %d ASN rows -> 'R'"
             % (fileName, siteId, len(refs), len(manifestsFlagged), asnsFlagged))
    return {"count": parsed["count"], "manifestsFlagged": manifestsFlagged,
            "asnsFlagged": asnsFlagged, "refs": refs, "fileHash": fileHash}


def _quarantine_file(srcPath, quarantineDir, log):
    """Move a no-DUNS-match / not-EDI file aside (Q11 — NOT silently dropped/left-to-rescan like legacy
    :439). Best-effort; logs on failure. IG83-TODO: on 8.3 prefer system.file move; here os.rename
    (atomic on one filesystem). Returns the destination path, or None if not moved (no quarantineDir)."""
    if srcPath is None or quarantineDir is None:
        return None
    import os as _os
    try:
        if not _os.path.isdir(quarantineDir):
            _os.makedirs(quarantineDir)
        dest = _os.path.join(quarantineDir, _os.path.basename(srcPath))
        _os.rename(srcPath, dest)
        return dest
    except Exception as e:
        log.warn("quarantine: failed to move %s -> %s: %s" % (srcPath, quarantineDir, e))
        return None


def _archive_file(srcPath, archiveDir, log):
    """Move a successfully-processed file to the archive (post-commit). The DB ledger is the AUTHORITATIVE
    idempotency guard (so even if the move fails, a re-drop is a NO-OP per _already_processed); the
    archive move is housekeeping. Best-effort; logs on failure. IG83-TODO: 8.3 system.file move +
    archive-on-success."""
    if srcPath is None or archiveDir is None:
        return None
    import os as _os
    try:
        if not _os.path.isdir(archiveDir):
            _os.makedirs(archiveDir)
        dest = _os.path.join(archiveDir, _os.path.basename(srcPath))
        _os.rename(srcPath, dest)
        return dest
    except Exception as e:
        log.warn("archive: failed to move %s -> %s: %s" % (srcPath, archiveDir, e))
        return None


def process_one_file(text, fileName, database=None, srcPath=None, archiveDir=None, quarantineDir=None):
    """Process ONE inbound file end-to-end (the unit the poll loop calls per file). ONE transaction:
      1. envelope check: line 1 must contain 'ISA' (legacy :65); else -> quarantine as not-EDI.
      2. DUNS guard: split the ISA, resolve delSL[4] -> site by VC_TMM_DUNS. No match -> quarantine
         (NOT silently dropped, Q11). Logged either way.
      3. idempotency: if the CONTENT HASH is already PROCESSED in the ledger -> NO-OP (kills the
         re-ingest hazard, §5.4). Content is the identity: a renamed redelivery (same content, new name)
         is skipped; a same-name/new-content file reprocesses.
      4. dispatch by transaction-set id (ST, copy(fcl,4,3)): '997' -> process_997, '824' -> process_824.
         Any other type (830/862/820/unknown) is OUT OF M1 SCOPE -> quarantine + log (M2 backlog).
      5. record the ledger row + commit. THEN (post-commit) archive the file.
      On any error: rollback (the file is NOT marked processed -> a fixed re-drop re-runs it), re-raise.

    Returns a result dict: {'type', 'siteId', 'outcome', 'detail', 'result'(the 997/824 summary)}.
    outcome in {'PROCESSED', 'QUARANTINED_NOT_EDI', 'QUARANTINED_NO_DUNS', 'QUARANTINED_OUT_OF_SCOPE',
                'SKIPPED_ALREADY_PROCESSED'}."""
    db = database if database is not None else DATABASE
    log = system.util.getLogger("SPIKE.process_inbound")
    fileHash = _hash_text(text)
    lines = _lines(text)
    line1 = lines[0] if lines else ""

    # 1. envelope check (legacy :65 — pos('ISA',fcl)>0).
    if "ISA" not in line1:
        log.warn("process_inbound file=%s: not an EDI document (no ISA on line 1) -> quarantine"
                 % (fileName,))
        # record a quarantine ledger row in its own tx (so a re-drop of the same junk is a no-op too).
        tx = system.db.beginTransaction(db)
        try:
            _record_processed(db, fileName, fileHash, "??", None, "QUARANTINED",
                              "not an EDI document (no ISA)", tx)
            system.db.commitTransaction(tx)
        finally:
            system.db.closeTransaction(tx)
        _quarantine_file(srcPath, quarantineDir, log)
        return {"type": None, "siteId": None, "outcome": "QUARANTINED_NOT_EDI",
                "detail": "no ISA on line 1", "result": None}

    delSL = _split_isa(line1)
    duns = delSL[4] if len(delSL) > 4 else None

    tx = system.db.beginTransaction(db)
    committed = False
    outcome = None
    detail = None
    siteId = None
    typ = None
    result = None
    try:
        # 2. DUNS guard.
        siteId = _resolve_site_by_duns(db, duns, tx)
        if siteId is None:
            log.warn("process_inbound file=%s: no site for trading-partner DUNS %r -> quarantine"
                     % (fileName, duns))
            _record_processed(db, fileName, fileHash, "??", None, "QUARANTINED",
                              "no site match for DUNS %r" % (duns,), tx)
            system.db.commitTransaction(tx)
            committed = True
            system.db.closeTransaction(tx)
            _quarantine_file(srcPath, quarantineDir, log)   # move aside (Q11 — NOT silently dropped)
            return {"type": None, "siteId": None, "outcome": "QUARANTINED_NO_DUNS",
                    "detail": "no site for DUNS %r" % (duns,), "result": None}

        # 3. idempotency.
        if _already_processed(db, fileName, fileHash, tx):
            log.info("process_inbound file=%s: already processed (ledger hit) -> skip" % (fileName,))
            system.db.rollbackTransaction(tx)     # nothing to commit; release the tx cleanly
            committed = True                      # (we did not write -> nothing to roll back beyond this)
            outcome = "SKIPPED_ALREADY_PROCESSED"
            detail = "ledger hit"
            return {"type": None, "siteId": siteId, "outcome": outcome, "detail": detail, "result": None}

        # 4. dispatch by transaction-set id.
        typ = _detect_type(lines)
        if typ == "997":
            result = process_997(text, fileName, database=db, fileHash=fileHash, siteId=siteId, tx=tx)
            detail = "%d acks, %d applied, %d unknown" % (
                len(result["acks"]), result["applied"], len(result["unknownEins"]))
        elif typ == "824":
            result = process_824(text, fileName, database=db, fileHash=fileHash, siteId=siteId, tx=tx)
            detail = "%d reject lines, %d ASN rows -> R" % (result["count"], result["asnsFlagged"])
        else:
            # 830/862/820/unknown -> M2 backlog (spec §4/§7). Quarantine + log; NOT processed as M1.
            log.warn("process_inbound file=%s: transaction-set %r is out of M1 scope -> quarantine"
                     % (fileName, typ))
            _record_processed(db, fileName, fileHash, typ or "??", siteId, "QUARANTINED",
                              "out of M1 scope (type %r)" % (typ,), tx)
            system.db.commitTransaction(tx)
            committed = True
            _quarantine_file(srcPath, quarantineDir, log)
            return {"type": typ, "siteId": siteId, "outcome": "QUARANTINED_OUT_OF_SCOPE",
                    "detail": "type %r out of M1 scope" % (typ,), "result": None}

        # 5. ledger row + commit.
        _record_processed(db, fileName, fileHash, typ, siteId, "PROCESSED", detail, tx)
        system.db.commitTransaction(tx)
        committed = True
    except Exception:
        if not committed:
            system.db.rollbackTransaction(tx)
        raise
    finally:
        system.db.closeTransaction(tx)

    # post-commit housekeeping: archive the processed file. The DB ledger is the authoritative
    # idempotency guard, so an archive-move failure does not risk a double-process.
    _archive_file(srcPath, archiveDir, log)
    outcome = "PROCESSED"
    log.info("process_inbound file=%s type=%s site=%s -> PROCESSED (%s)"
             % (fileName, typ, siteId, detail))
    return {"type": typ, "siteId": siteId, "outcome": outcome, "detail": detail, "result": result}


def process_inbound(inDir, database=None, archiveDir=None, quarantineDir=None):
    """The SCHEDULED POLL entry (Q11 — replaces the legacy manual button). List the EDIIn directory, and
    for each FILE (not subdir) read the body + process_one_file it. Reproduces the legacy FindFirst poll
    (EDIUpload.pas:53-56) but driven by a gateway scheduled TIMER/event script (8.1-safe) instead of a
    blocking button.

    IG83-TODO: on 8.3, prefer an Event Stream + system.file move/archive for the file I/O; this 8.1-safe
    shape uses a gateway timer script calling process_inbound + os-level directory listing/rename.
    IG81-COMPAT: gateway timer scripts + the system.db/system.file APIs here are identical on 8.1/8.3.

    On the gateway the inDir = [DIRECTORIES] EDIIn; for the spike test a temp dir. archiveDir defaults to
    <inDir>/Archive, quarantineDir to <inDir>/Quarantine (created on demand). Returns a list of the
    per-file result dicts. A failure on ONE file is logged and the poll CONTINUES to the next file (one
    bad file must not block the rest — the failed file is NOT marked processed, so the next poll retries
    it)."""
    db = database if database is not None else DATABASE
    log = system.util.getLogger("SPIKE.process_inbound")
    import os as _os
    if archiveDir is None:
        archiveDir = _os.path.join(inDir, "Archive")
    if quarantineDir is None:
        quarantineDir = _os.path.join(inDir, "Quarantine")

    results = []
    try:
        names = sorted(_os.listdir(inDir))
    except OSError as e:
        log.error("process_inbound: cannot list inDir %s: %s" % (inDir, e))
        return results

    for name in names:
        srcPath = _os.path.join(inDir, name)
        if not _os.path.isfile(srcPath):
            continue                              # skip subdirs (Archive/Quarantine)
        try:
            with open(srcPath, "rb") as fh:
                text = fh.read().decode("utf-8", "replace")
        except Exception as e:
            log.error("process_inbound: cannot read %s: %s" % (srcPath, e))
            results.append({"type": None, "siteId": None, "outcome": "READ_ERROR",
                            "detail": str(e), "result": None, "fileName": name})
            continue
        try:
            r = process_one_file(text, name, database=db, srcPath=srcPath,
                                 archiveDir=archiveDir, quarantineDir=quarantineDir)
            r["fileName"] = name
            results.append(r)
        except Exception as e:
            # one bad file must not block the rest; it is NOT marked processed -> next poll retries it.
            log.error("process_inbound: FAILED processing %s: %s" % (name, e))
            results.append({"type": None, "siteId": None, "outcome": "ERROR",
                            "detail": str(e), "result": None, "fileName": name})
    log.info("process_inbound poll of %s -> %d files: %s"
             % (inDir, len(results), ", ".join("%s=%s" % (r.get("fileName"), r["outcome"]) for r in results)))
    return results
