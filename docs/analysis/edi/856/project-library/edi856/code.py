# edi856 — Project Library: the OUTBOUND 856 ASN X12 builder (M1 Rank-2), reimplemented as PURE
# Jython-2.7/CPython logic + a thin gateway driver, the producer-recipe pattern (twin of asn/code.py
# computeAsnDetails + create_asn).
#
# BUILD TARGET — BYTE-FAITHFUL to the legacy 856 that TEMA currently accepts. The legacy builder is
# EDI856Object.pas (T856EDI; LIVE per InventorySystem.dpr:48). This module PORTS its segment logic
# verbatim (every quirk reproduced); it does NOT wrap REPORT_EDI856 (which self-flips status and
# hardcodes WHERE IN_ASN_EIN=6440 — see report-edi856-data-analysis.md).
#
# Source truth (both complete on disk):
#   docs/analysis/edi/856/edi856-wire-format.md          — the byte-exact segment spec (THE target)
#   docs/analysis/edi/856/report-edi856-data-analysis.md — the data feed + the AVOID list + decisions A-E
#   EDI856Object.pas                                     — the original builder (line refs cited per segment)
#   docs/analysis/edi/asn-invoice.md §4.2                — the segment map cross-check
#
# This module holds BOTH:
#   * build_856(...)  — the PURE segment builder. Imports nothing, references no gateway global at import
#     time -> CPython-importable for the unit test (scripts/e2e/test_edi856_build.py). Given a header,
#     the detail feed rows, the site dict, the EIN and a file timestamp, it returns the ORDERED list of
#     segment strings (the driver joins them with CRLF). No DB, no I/O.
#   * send_856(...)  — the GATEWAY DRIVER. ONE transaction: allocate the per-site EIN from
#     INV_SITES.IN_EIN_SEQ (atomic), stamp it on the ASN header, read the feed (the side-effect-free
#     SELECT in spike-edi856-feed.sql), build_856(...), write the file, flip the ASN C->S per-ASN,
#     commit / rollback. References `system` only at CALL time (like asn.create_asn / order.commitOrders),
#     so importing the module never needs the gateway runtime; it is driven headlessly through
#     scripts/e2e/jython_shim.py (the R8 system.db shim).
#
# Jython 2.7 (8.1 and 8.3, no delta — pure string assembly + system.db; no Python-3-only libs).
#
# LOCKED DECISIONS realized here (byte-faithful — David locked byte parity with the TEMA-accepted output):
#   A  KEEP the cost INNER join (feed filters to cost-covered lines, like legacy) — feed SQL.
#   B  INNER forecast join for Kanban — the SAME JOIN INV_FORECAST_DETAIL_INF the legacy proc uses (a part
#      with >1 forecast row fans out, GROUP-BY keeps every distinct kanban). NOT a CROSS APPLY TOP 1 — feed SQL.
#   C  GROUP-BY collapse (NO sum), KEEPING m.MO_PRICE as a GROUP-BY key (legacy splits rows on it even
#      though the 856 never carries a price) — feed SQL.
#   D  emit CRLF ONLY (NO '~' segment terminator) — the driver joins with "\r\n"; the segment terminator
#      VC_SEP_SEGMENT is READ into the site dict but NEVER emitted (legacy Trap 1 / EDI856Object.pas:13).
#   E  filename = the OPERATIONAL-SENDER pattern (MainMenu.pas:2718/2723, the C->S-flip send path; NOT the
#      recreate button). NORMAL: 856 + copy(prodDate,5,4)=MMDD + LineName + .txt. HOT-CALL (seq='-1'):
#      8HC + copy(prodDate,4,5)=Y+MMDD + IntToStr(y) + LineName + .txt. See _filename_856() below.
#   EIN  allocated per-site from INV_SITES.IN_EIN_SEQ AT SEND (atomic), %09d in ISA13/GS06/ST02/SE02/
#        GE02/IEA02, BSN02 = yyyymmdd + %09d.  NOT the 6440 literal; NOT EIN-at-create.
#   Status flip  per-ASN C->S at send, DECOUPLED — NEVER REPORT_EDI856's self-flip, NEVER the blanket
#        UPDATE_ASNStatus (WHERE VC_ASN_STATUS='C' across all sites).


# =================================================================================================
# PURE segment builder
# =================================================================================================

class Edi856BuildError(Exception):
    """Raised by build_856 on a structural input error (e.g. a feed row missing a required column).
    The driver rolls back on it, matching the Delphi try/except -> Execute returns False."""
    pass


def _fmt_ein(ein):
    """%9.9d — 9 chars, zero-padded (EDI856Object.pas:159 et al., EVERYWHERE the EIN appears).
    Python's %09d == Pascal's %9.9d for a non-negative int (zero-pad to width 9). EIN is int NOT NULL
    (report-edi856-data-analysis.md), range 3502-9057, so never negative. We do NOT truncate: legacy
    Format('%9.9d') would WIDEN past 9 if the EIN exceeded 999,999,999 — far outside the real range —
    so a faithful port keeps the same widen-not-truncate behaviour. (%09d widens identically.)"""
    return "%09d" % int(ein)


def _fixed(s, width):
    """Pascal Format('%-Ns', [s]) — left-justify in a fixed field, space-pad to `width`. The ISA is a
    POSITIONAL segment: ISA02/04 = 10, ISA06/08 = 15. EDI856Object.pas does NOT truncate over-width
    values (Pascal %-Ns only PADS, never cuts), so an over-wide DUNS/abbr would push the ISA past its
    fixed width — a latent legacy bug (Trap 6). The wire-format doc says the rebuild MUST hard-truncate
    to keep the ISA byte-exact-width; we truncate to `width` AND pad, so the field is EXACTLY `width`
    chars whatever the input. (For the real site values — abbr 'MAS', DUNS 9 digits, DUNS-supplier
    '969009112-71930' = 15 — this truncation never fires, so it is byte-identical to legacy on real
    data; it only diverges on the pathological over-wide input legacy would have mis-emitted anyway.)"""
    s = "" if s is None else str(s)
    if len(s) > width:
        s = s[:width]                      # hard-truncate (Trap 6 fix; never fires on real data)
    return s + (" " * (width - len(s)))    # left-justify, space-pad to exactly `width`


def _one_char(s, field):
    """X12 ISA15 (usage indicator) and ISA16 (component-element separator) are single-character fields.
    Legacy ISA15 = Site.SiteEDIMode (a 1-char 'P'/'T'); ISA16 = the sub-element separator char. The
    spike INV_SITES placeholder carries VC_EDI_MODE='PROD' (4 chars) which would emit a MALFORMED 4-char
    ISA15 (X12 mandates exactly one char) — so enforce single-char here, taking the FIRST char (legacy
    intent: 'P'). Raise on empty (a missing usage indicator is a hard X12 error, not something to widen).
    Real loaded data ('P'/'T', '#'/'>') is already 1 char so this never alters it; it only catches a bad
    load (defect #7 — placeholder VC_EDI_MODE='PROD')."""
    s = "" if s is None else str(s)
    if len(s) < 1:
        raise Edi856BuildError("edi856: ISA %s must be a single character, got empty" % (field,))
    return s[0]


def _yymmdd(prodDate):
    """ISA09 = copy(fPickupDate,3,6) = yymmdd (century dropped) — UNLIKE GS04/BSN03/DTM02 which are full
    yyyymmdd. EDI856Object.pas:154. Pascal copy(s,3,6) is 1-based -> chars 3..8 of 'yyyymmdd' = the last
    two year digits + MM + DD. '20260618' -> '260618'."""
    prodDate = str(prodDate)
    if len(prodDate) != 8:
        raise Edi856BuildError(
            "edi856: production date must be yyyymmdd (8 chars), got %r" % (prodDate,))
    return prodDate[2:8]


def build_856(header, detailRows, site, ein, fileTime, trailerId="1234567890"):
    """Build the byte-faithful 856 as an ORDERED list of segment strings. The driver joins them with
    CRLF (decision D — no segment terminator emitted).

    Reproduces EDI856Object.pas T856EDI.Execute (the ISA..IEA envelope + the HL S->O->I hierarchy).

    Parameters
    ----------
    header : dict for the ASN header. Keys:
        'prodDate'  : VC_PRODUCTION_DATE, varchar(8) yyyymmdd. Drives ISA09 (yymmdd), GS04/BSN02/BSN03/
                      DTM02 (yyyymmdd). EDI856Object.pas reads it as fPickupDate (:111).
    detailRows : list of dicts, one per surviving feed row (the spike-edi856-feed.sql projection), IN
        FEED ORDER (the driver passes them ordered by Manifest, then a stable detail key, so the S->O->I
        grouping is deterministic — legacy relied on REPORT_EDI856's GROUP BY ... order). Keys:
        'Manifest'   : d.VC_MANIFEST_NUMBER  -> PRF (Order HL) + the new-Order break key.
        'PartNumber' : d.VC_ASSY_PART_NUMBER -> LIN BP.
        'ShipQty'    : d.IN_QTY (int)        -> SN1 (no leading zeros).
        'Kanban'     : f.VC_ASSY_KANBAN_NUMBER -> LIN RC.
        (UnitPrice / MO_PRICE is NOT emitted — the 856 carries no price; referenced 0x in EDI856Object.pas.)
    site : dict from the site row (INV_SITES). Keys (legacy AD_GetSite field -> INV_SITES column):
        'abbr'        SiteAbbr               (VC_SITE_ABBR)        -> ISA02
        'duns'        SiteDUNS               (VC_DUNS)            -> ISA04, ISA06 prefix, GS02
        'supplierCode' SiteSupplierCode      (VC_SUPPLIER_CODE)   -> ISA06 suffix (duns-supplierCode)
        'tmmDuns'     SiteTMMDUNS            (VC_TMM_DUNS)        -> ISA08, GS03
        'ediMode'     SiteEDIMode            (VC_EDI_MODE)        -> ISA15 (T/P)
        'deliveryMethodCode' SiteDeliveryMethodCode (VC_DELIVERY_METHOD_CODE) -> TD5 carrier code
        'sepElement'  SiteSepElement         (VC_SEP_ELEMENT)     -> the element separator ('*')
        'sepSubElement' SiteSepSubElement    (VC_SEP_SUBELEMENT)  -> ISA16 (sub-element sep)
        'sepSegment'  SiteSepSegment         (VC_SEP_SEGMENT)     -> READ but NEVER EMITTED (decision D).
    ein : int — the per-site EIN allocated at send. Formatted %09d in all 7 control positions; BSN02 =
        prodDate + %09d.
    fileTime : a 'HHMM' string (the ONE shared timestamp, captured once per file — legacy f810Time).
        The driver passes system.date.format(system.date.now(), "HHmm"); the pure builder takes the
        already-formatted 4-char string so it stays I/O-free and deterministically testable.
    trailerId : the TD3 trailer/equipment id. Legacy literal '1234567890' (Trap 2) — defaulted for
        parity but PARAMETERIZED per the build spec (a wrong literal is not a TEMA reject, but the spec
        wants the magic value liftable). EDI856Object.pas:308.

    Returns the list of segment strings in emission order (EDI856Object.pas:116-126):
        ISA, GS, ST, BSN, DTM, HL(S), TD1, TD5, TD3, [per manifest: HL(O), PRF], [per detail: HL(I),
        LIN, SN1], CTT, SE, GE, IEA.

    SE01 = count of segments ST..SE inclusive (counted, never a magic offset). CTT01 = HL count only.
    """
    sep = site["sepElement"]                 # element separator ('*')
    subSep = site["sepSubElement"]           # sub-element separator (ISA16 only)
    prodDate = str(header["prodDate"])       # yyyymmdd
    isaDate = _yymmdd(prodDate)              # ISA09 = yymmdd
    hhmm = str(fileTime)                     # the shared HHMM
    einStr = _fmt_ein(ein)

    segs = []

    # --- ISA (EDI856Object.pas:140-164) — positional, %-Ns fixed widths. ---------------------------
    # ISA*00*<abbr %-10>*01*<duns %-10>*ZZ*<duns-supplier %-15>*01*<tmmDuns %-15>*<yymmdd>*<hhmm>*U*
    #    00400*<%09d EIN>*0*<ediMode>*<subSep>
    isa06 = str(site["duns"]) + "-" + str(site["supplierCode"])   # :151 — the '-' is data
    segs.append(sep.join([
        "ISA",
        "00",                               # :146 ISA01 auth-info qualifier
        _fixed(site["abbr"], 10),           # :147 ISA02 auth info (SiteAbbr, %-10s)
        "01",                               # :148 ISA03 security-info qualifier
        _fixed(site["duns"], 10),           # :149 ISA04 security info (SiteDUNS, %-10s)
        "ZZ",                               # :150 ISA05 interchange-id qualifier (sender)
        _fixed(isa06, 15),                  # :151 ISA06 sender id (DUNS-SupplierCode, %-15s)
        "01",                               # :152 ISA07 interchange-id qualifier (receiver = DUNS)
        _fixed(site["tmmDuns"], 15),        # :153 ISA08 receiver id (SiteTMMDUNS, %-15s)
        isaDate,                            # :154 ISA09 date = copy(prodDate,3,6) = yymmdd
        hhmm,                               # :155 ISA10 time
        "U",                                # :156 ISA11 interchange control-standards id
        "00400",                            # :157 ISA12 interchange control-version
        einStr,                             # :159 ISA13 interchange control # = EIN
        "0",                                # :160 ISA14 ack-requested
        _one_char(site["ediMode"], "ISA15 (usage indicator)"),   # :161 ISA15 (T/P) — single char
        _one_char(subSep, "ISA16 (sub-element separator)"),      # :162 ISA16 — single char
    ]))

    # --- GS (EDI856Object.pas:174-189) — GS06 = EIN (group control #). NO trailing element sep. -----
    segs.append(sep.join([
        "GS",
        "SH",                               # functional id (Ship Notice)
        str(site["duns"]),                  # :181 GS02 application sender = DUNS
        str(site["tmmDuns"]),               # :182 GS03 application receiver = TMM DUNS
        prodDate,                           # :183 GS04 date = full yyyymmdd
        hhmm,                               # :184 GS05 time
        einStr,                             # :185 GS06 group control # = EIN
        "X",                                # :186 GS07 responsible-agency
        "004010",                           # :187 GS08 version
    ]))

    # --- ST (EDI856Object.pas:199-208) — ST02 = EIN, NOT a 1-based counter. ------------------------
    segs.append(sep.join(["ST", "856", einStr]))

    # --- BSN (EDI856Object.pas:219-231) — BSN02 = prodDate + %09d EIN (17 chars, the shipment id). --
    segs.append(sep.join([
        "BSN",
        "00",                               # purpose (Original)
        prodDate + einStr,                  # :227 BSN02 shipment id = yyyymmdd + %09d EIN (17 chars)
        prodDate,                           # :228 BSN03 date = yyyymmdd
        hhmm,                               # :229 BSN04 time (NO trailing sep — legacy commented it out)
    ]))

    # --- DTM (EDI856Object.pas:242-254) — shipped date/time; ET hardcoded. ------------------------
    segs.append(sep.join(["DTM", "011", prodDate, hhmm, "ET"]))

    # --- HL hierarchy S -> O -> I (EDI856Object.pas:265-378) --------------------------------------
    # Built into a buffer (HLList in legacy) so CTT/SE counts are computed from the ACTUAL emitted
    # segments. fHID is the HL id sequence (1-based, increments per HL). hlCount = # of HL segments
    # only (S + #orders + #items) -> CTT01.
    hl, hlCount, hid = [], 0, 1

    # Shipment HL (:285). HL*1**S*1 — HL02 empty (no parent). Then TD1, TD5, TD3.
    hl.append(sep.join(["HL", str(hid), "", "S", "1"]))   # HL02 empty (two seps -> '')
    hid += 1
    hlCount += 1
    # TD1 (:294-296): 'TD1'+sep THEN +sep -> 'TD1**' (TWO empty trailing elements). The .pas builds
    # 'TD1'+fSepElement (line 294) then +fSepElement (line 295) with NO field appended after either, so
    # the segment ends with two separators. join(["TD1","",""]) == 'TD1**'.
    hl.append(sep.join(["TD1", "", ""]))                   # :294-296 TD1** (two empty elements)
    hl.append(sep.join(["TD5", "B", "25", "00000",
                        str(site["deliveryMethodCode"])])) # :298 TD5*B*25*00000*<deliveryMethodCode>
    hl.append(sep.join(["TD3", "TL", "", str(trailerId)])) # :305 TD3*TL**<trailerId> (Trap 2)

    # Order HL per distinct manifest (:317-336) + Item HL per detail (:338-359).
    lastManifest = None
    orderParent = None
    for row in detailRows:
        for k in ("Manifest", "PartNumber", "ShipQty", "Kanban"):
            if k not in row:
                raise Edi856BuildError("edi856: feed row missing column %r: %r" % (k, row))
        manifest = str(row["Manifest"])
        if manifest != lastManifest:
            # New Order HL (:320). HL*<id>*1*O*1 — HL02 hardcoded '1' (always parents the shipment,
            # Trap 3 — NOT the running parent). PRF*<Manifest>-<Manifest> (the '-' is data).
            hl.append(sep.join(["HL", str(hid), "1", "O", "1"]))
            orderParent = hid
            hid += 1
            hlCount += 1
            hl.append("PRF" + sep + manifest + "-" + manifest)   # :330 PRF (single element)
            lastManifest = manifest

        # Item HL (:338). HL*<id>*<orderParent>*I*0.
        hl.append(sep.join(["HL", str(hid), str(orderParent), "I", "0"]))
        hid += 1
        hlCount += 1
        # LIN**BP*<PartNumber>*RC*<Kanban>*  (:347-352 — LIN01 empty; TRAILING sep after the kanban).
        # The .pas appends +fSepElement after EVERY field INCLUDING the kanban (line 352 ends with
        # '...Kanban...+fSepElement'), so the segment ends in a trailing separator (one empty trailing
        # element). join([...,"RC",kanban]) drops it; append a trailing "" to reproduce the byte.
        hl.append(sep.join(
            ["LIN", "", "BP", str(row["PartNumber"]), "RC", str(row["Kanban"]), ""]))
        # SN1**<ShipQty>*PC  (:355 — SN101 empty; qty = IN_QTY, no leading zeros; PC = pieces)
        hl.append(sep.join(["SN1", "", str(int(row["ShipQty"])), "PC"]))

    segs.extend(hl)

    # --- CTT (EDI856Object.pas:380) — count of HL segments ONLY (NOT line count). -----------------
    segs.append(sep.join(["CTT", str(hlCount)]))

    # --- SE (EDI856Object.pas:400) — SE01 = count of segments ST..SE INCLUSIVE (SE counts itself). -
    # Compute by counting the actual emitted segments from ST to SE, never a magic offset (the TEMA-
    # reject trap). ST..(DTM) = 3, + the HL buffer, + CTT, + SE itself.
    #   index of ST in segs is 2 (ISA, GS, ST). Everything from there to the CTT we already appended
    #   is the ST..CTT run; +1 for the SE we are about to add.
    st_index = 2
    se01 = (len(segs) - st_index) + 1       # segments ST..CTT already in `segs`, plus SE itself
    segs.append(sep.join(["SE", str(se01), einStr]))

    # --- GE (EDI856Object.pas:421) GE*1*<%09d EIN> (GE02 must = GS06). -----------------------------
    segs.append(sep.join(["GE", "1", einStr]))
    # --- IEA (EDI856Object.pas:441) IEA*1*<%09d EIN> (IEA02 must = ISA13). -------------------------
    segs.append(sep.join(["IEA", "1", einStr]))

    return segs


def _filename_856(prodDate, lineName, startSeq=None, counter=1):
    """The 856 output filename. TWO branches, keyed on the hot-call sentinel — reproduced from the
    OPERATIONAL SENDER (MainMenu.ResendMarkedEDIsClick, the C->S-flip send path), NOT the recreate button.

    The operational sender is the live daily path: it builds the 856 for every queued ASN and flips it
    C->S. The earlier build was anchored on the ASNInvoice recreate button (a different code path with a
    DIFFERENT, latently-inconsistent offset) — that was WRONG for both branches: both omitted LineName,
    and the normal branch used the wrong date offset. The canonical patterns are MainMenu.pas:

        if Data_Module.EDI856DataSet.FieldByName('StartSeq').AsString <> '-1' then      // NORMAL ASN
            AssignFile(fcf, ...+'\856'+copy(EDI856.PickupDate,5,4)+EDI856.LineName+'.txt');  // :2718
        else                                                                            // HOT-CALL ASN
            AssignFile(fcf, ...+'\8HC'+copy(EDI856.PickupDate,4,5)+IntToStr(y)+EDI856.LineName+'.txt'); // :2723
            INC(y);                                                                      // :2724

    NORMAL (startSeq != '-1', MainMenu.pas:2718):
        '856' + copy(PickupDate,5,4) + LineName + '.txt'.
        Pascal copy(s,5,4) is 1-based -> chars 5..8 of 'yyyymmdd' = MM + DD (4 chars) = prodDate[4:8]
        0-based. '20260618' / 'COROLLA' -> '856' + '0618' + 'COROLLA' + '.txt' = '8560618COROLLA.txt'.

    HOT-CALL (startSeq == '-1', the VC_START_SEQ_NUMBER='-1' sentinel HotCallEntry stamps,
        HotCallEntry.pas:238-243; MainMenu.pas:2722-2724):
        '8HC' + copy(PickupDate,4,5) + IntToStr(y) + LineName + '.txt'.
        copy(PickupDate,4,5) = prodDate[3:8] (last-year-digit + MM + DD, 5 chars) — a DIFFERENT offset
        from the normal branch (preserve this legacy asymmetry: normal [4:8]=MMDD; hot-call [3:8]=Y+MMDD).
        IntToStr(y) is the per-send-batch hot-call counter, NOT a literal '1': y init 1 (:2702), INC only
        in the hot-call branch (:2724). '20260618' / y=1 / 'COROLLA' -> '8HC' + '60618' + '1' + 'COROLLA'
        + '.txt' = '8HC606181COROLLA.txt'.

    BYTE-FAITHFUL (P13): TEMA's dispatcher may key on the '8HC' prefix to route hot-calls, so the exact
    pattern matters. A hot-call ASN MUST produce '8HC...<counter><LineName>.txt'; a normal ASN MUST
    produce '856<MMDD><LineName>.txt'.

    lineName : the ASN header's VC_LINE_NAME (str) — appended to BOTH branches (legacy EDI856.LineName).
    startSeq : the ASN header's VC_START_SEQ_NUMBER (str). '-1' -> hot-call branch; anything else (incl.
        None, for back-compat with callers that only build normal ASNs) -> the normal '856' pattern.
    counter  : the hot-call y-equivalent (1-based). Only used in the hot-call branch. The legacy y is a
        per-SEND-BATCH counter (Nth hot-call file in one ResendMarkedEDIs run); the rebuild sends per-ASN,
        so the driver derives a deterministic, collision-free per-ASN equivalent (see send_856). The exact
        y RANGE is golden-pending (P13 cutover check) — byte-faithful per source, range unverified."""
    prodDate = str(prodDate)
    lineName = "" if lineName is None else str(lineName)
    if len(prodDate) != 8:
        raise Edi856BuildError("edi856: production date must be yyyymmdd, got %r" % (prodDate,))
    if startSeq is not None and str(startSeq) == "-1":
        # HOT-CALL — MainMenu.pas:2723: '8HC' + copy(pd,4,5) + IntToStr(y) + LineName + '.txt'
        return "8HC" + prodDate[3:8] + str(int(counter)) + lineName + ".txt"
    # NORMAL — MainMenu.pas:2718: '856' + copy(pd,5,4) + LineName + '.txt'
    return "856" + prodDate[4:8] + lineName + ".txt"   # copy(s,5,4) = [4:8] 0-based = MMDD


def _cleanup_tmp(tmpPath):
    """Best-effort delete of the orphan .tmp on a rolled-back send (atomicity). Never raises — the caller
    is mid-propagation of the REAL failure; we must not mask it. (If the .tmp can't be removed it is at
    worst a harmless leftover the mailer ignores — never a FINAL 856 the DB didn't commit.)"""
    import os as _os
    try:
        if _os.path.exists(tmpPath):
            _os.remove(tmpPath)
    except Exception:
        pass


# =================================================================================================
# send_856 — the GATEWAY DRIVER. ONE transaction.
# =================================================================================================
#
# IG81-COMPAT: every gateway API used here (runPrepQuery / runScalarPrepQuery(...,tx=) / runPrepUpdate
# (...,tx=) / beginTransaction / system.date / system.file.writeFile / getLogger) is identical on
# 8.1.52 and 8.3 — no version guard needed.

from db_shared import CONNECTION as DATABASE   # centralized DB-conn name (default Inventory_Spike; single prod-rename point)

# The feed SQL (spike-edi856-feed.sql) inlined here so the driver carries no external-file dependency at
# runtime (the .sql file is the reviewable canonical copy). This string is BYTE-IDENTICAL to the .sql
# file's SELECT body EXCEPT the single bind: the .sql declares @ASNID and the inline uses runPrepQuery's
# `?`. test_edi856_e2e.py enforces `_FEED_SQL == <.sql SELECT body with '@ASNID' -> '?'>` so the two
# cannot drift. Reproduces legacy REPORT_EDI856 @EIN=0 (NOT the 6440 literal; NO self-flip):
#   A  INNER cost join (feed filters to cost-covered lines, like legacy).
#   B  INNER forecast join — the SAME JOIN INV_FORECAST_DETAIL_INF f ON ... the legacy proc uses (fan-out
#      a part with >1 forecast row to multiple rows, then GROUP-BY keeps each distinct kanban). NOT a
#      CROSS APPLY TOP 1 (which dropped legacy lines + was nondeterministic over a heap).
#   C  GROUP-BY collapse (no sum) — and the GROUP BY KEEPS m.MO_PRICE exactly as legacy does, so a part
#      with 2 price-distinct overlapping cost windows emits 2 detail rows (legacy splits on MO_PRICE even
#      though the 856 never carries the price). Dropping it collapsed legacy's 2 rows to 1.
#   inclusive cost window (<= / >=) — matches the live proc.
# 5-col projection (the 856 emits no price; legacy's UnitPrice/SiteEIN/StartSeq/LineName are unused by
# the builder). MO_PRICE is in the GROUP BY but NOT the SELECT — it changes cardinality, not the wire.
_FEED_SQL = (
    "SELECT d.VC_MANIFEST_NUMBER AS Manifest, "
    "d.VC_ASSY_PART_NUMBER AS PartNumber, "
    "d.IN_QTY AS ShipQty, "
    "a.VC_PRODUCTION_DATE AS PickUpDate, "
    "f.VC_ASSY_KANBAN_NUMBER AS Kanban "
    "FROM INV_ASN_MST a "
    "JOIN INV_ASN_DETAIL_MST d ON a.IN_ASN_ID = d.IN_ASN_ID "
    "JOIN INV_MANIFEST_COST_MST m ON d.VC_ASSY_PART_NUMBER = m.VC_ASSY_PART_NUMBER_CODE "
    "JOIN INV_FORECAST_DETAIL_INF f ON d.VC_ASSY_PART_NUMBER = f.VC_ASSY_PART_NUMBER_CODE "
    "WHERE a.IN_ASN_ID = ? "
    "AND m.VC_START_MANIFEST <= a.VC_PRODUCTION_DATE "
    "AND m.VC_END_MANIFEST   >= a.VC_PRODUCTION_DATE "
    "GROUP BY d.VC_MANIFEST_NUMBER, d.VC_ASSY_PART_NUMBER, m.MO_PRICE, d.IN_QTY, "
    "a.VC_PRODUCTION_DATE, f.VC_ASSY_KANBAN_NUMBER "
    "ORDER BY d.VC_MANIFEST_NUMBER, d.VC_ASSY_PART_NUMBER"
)


def _read_site(database, siteId, tx):
    """Read the site identity row (INV_SITES) on the SAME transaction and map the columns to the dict
    build_856 expects. We read it inside the tx (runPrepQuery accepts tx on 8.1+) so it sees the EIN
    allocation's own row (we UPDATE IN_EIN_SEQ on this row earlier in the tx) without blocking on the
    uncommitted row lock. Returns the siteDict.

    NB on byte-fidelity: the legacy site identity actually lives in VehicleOrder.Site (AD_GetSite); the
    rebuild relocates it to INV_SITES (the sites-CRUD decision). The COLUMN VALUES in the spike's
    Inventory.INV_SITES are PLACEHOLDER (abbr 'MAS', DUNS '000000001', supplier 'MAS', sep '>' subelem,
    EDI mode 'PROD') — NOT the real legacy values (VehicleOrder.Site: DUNS 969009112, supplier 71930,
    TMM DUNS 808369495, EDI mode 'P', sub-elem '#'). So the BYTES this driver emits on the spike reflect
    the placeholder site row, not the TEMA wire. The test asserts STRUCTURE + self-consistency against
    whatever site row is configured; true wire bytes require the real site values be loaded into
    INV_SITES at cutover. (Honest verification — see the test docstring.)"""
    rows = system.db.runPrepQuery(
        "SELECT VC_SITE_ABBR, VC_DUNS, VC_SUPPLIER_CODE, VC_TMM_DUNS, VC_EDI_MODE, "
        "VC_DELIVERY_METHOD_CODE, VC_SEP_ELEMENT, VC_SEP_SUBELEMENT, VC_SEP_SEGMENT "
        "FROM INV_SITES WHERE IN_SITE_ID = ?", [int(siteId)], database, tx)
    if not len(rows):
        raise Edi856BuildError("send_856: no INV_SITES row for site id %r" % (siteId,))
    r = rows[0]
    return {
        "abbr": r["VC_SITE_ABBR"],
        "duns": r["VC_DUNS"],
        "supplierCode": r["VC_SUPPLIER_CODE"],
        "tmmDuns": r["VC_TMM_DUNS"],
        "ediMode": r["VC_EDI_MODE"],
        "deliveryMethodCode": r["VC_DELIVERY_METHOD_CODE"],
        "sepElement": r["VC_SEP_ELEMENT"],
        "sepSubElement": r["VC_SEP_SUBELEMENT"],
        "sepSegment": r["VC_SEP_SEGMENT"],     # read; NEVER emitted (decision D)
    }


def send_856(asnId, site, database=None, outDir=None, trailerId="1234567890", fileTime=None):
    """Send one ASN's 856: allocate the per-site EIN, stamp the header, build the file, write it, flip
    the ASN C->S. ONE Inventory transaction; commit on success, rollback + re-raise on any error.

    Reproduces the legacy create/recreate send path (ASNSelect.CreateASN / ASNInvoice recreate) with the
    M1 decisions: EIN-at-send from INV_SITES.IN_EIN_SEQ (NOT SiteEIN+1-at-create); a DECOUPLED per-ASN
    C->S flip (NOT REPORT_EDI856's self-flip, NOT the blanket UPDATE_ASNStatus); the feed via the pure
    SELECT (NOT the proc).

    Parameters
    ----------
    asnId    : INV_ASN_MST.IN_ASN_ID to send.
    site     : the INV_SITES.IN_SITE_ID owning the EIN sequence. Site-scoped EIN allocation; site id
               derives from the ASN/session, NEVER a client param (multi-site D1). Today the ASN header
               carries no IN_SITE_ID column (added at M4), so the caller passes the site explicitly.
    database : Inventory connection name (defaults to DATABASE).
    outDir   : directory to write the 856 file into ([DIRECTORIES] EDIOut on the gateway). For the spike
               test a temp dir. If None, the file is NOT written (the function still returns the bytes,
               for a dry-run / structural assertion).
    trailerId: TD3 trailer id (default the legacy literal '1234567890' for parity; parameterized).
    fileTime : override the HHMM stamp (testing determinism); default system.date.now() -> 'HHmm'.

    Returns a dict: {'asnId', 'ein', 'filename', 'path', 'segments'(list), 'text'(str), 'rowCount'}.

    Transaction order (ONE tx) + temp-then-rename file atomicity:
      1. allocate EIN atomically from INV_SITES.IN_EIN_SEQ (site-scoped UPDATE ... OUTPUT the bumped
         value) — the at-send increment. NEVER the 6440 literal.
      2. stamp IN_ASN_EIN on the ASN header.
      3. read the feed (pure SELECT, @ASNID-parameterized; NO self-flip) + the header prodDate.
      4. build_856(...) (pure).
      5. write the file TO A .tmp (NOT the final name).
      6. flip the ASN VC_ASN_STATUS C->S, per-ASN (WHERE IN_ASN_ID = @asnId) — DECOUPLED.
      commit.
      7. ONLY after the commit succeeds: rename <name>.tmp -> <name>.
      On ANY error: rollback + delete the orphan .tmp + re-raise (the EIN bump, the stamp, and the flip
      all unwind together AND the final 856 file never appears — so a rolled-back send can never leave a
      phantom ASN on disk under an uncommitted EIN).
    """
    db = database if database is not None else DATABASE
    log = system.util.getLogger("SPIKE.send_856")
    if fileTime is None:
        # ONE shared HHMM for the whole file (legacy f810Time captured once). system.date is gateway-only,
        # referenced at CALL time so import stays runtime-free.
        fileTime = system.date.format(system.date.now(), "HHmm")

    tmpPath = None          # the <name>.tmp written inside the tx; renamed to the final name post-commit
    committed = False        # only after a successful commitTransaction do we expose the final 856 file
    tx = system.db.beginTransaction(db)
    try:
        # --- 1. allocate the per-site EIN atomically (INV_SITES.IN_EIN_SEQ) ------------------------
        # UPDATE ... SET IN_EIN_SEQ = IN_EIN_SEQ + 1 OUTPUT INSERTED.IN_EIN_SEQ — a single atomic
        # statement: the increment and the read are one operation under the row lock, so two concurrent
        # sends never get the same EIN (the legacy AD_UpdateEIN had no WHERE -> a D1 cross-site hazard;
        # this is site-scoped by IN_SITE_ID). On the tx so it unwinds with the rest on rollback.
        ein = system.db.runScalarPrepQuery(
            "UPDATE INV_SITES SET IN_EIN_SEQ = IN_EIN_SEQ + 1 "
            "OUTPUT INSERTED.IN_EIN_SEQ WHERE IN_SITE_ID = ?", [int(site)], db, tx)
        if ein is None:
            raise Edi856BuildError("send_856: no INV_SITES row for site id %r (EIN alloc)" % (site,))
        ein = int(ein)

        # --- 2. stamp the EIN on the ASN header ---------------------------------------------------
        system.db.runPrepUpdate(
            "UPDATE INV_ASN_MST SET IN_ASN_EIN = ? WHERE IN_ASN_ID = ?",
            [ein, int(asnId)], db, tx=tx)

        # --- 3. read the feed + the header prodDate (pure SELECT; NO self-flip) --------------------
        # Both reads touch INV_ASN_MST, the row we just stamped the EIN onto in THIS tx. They MUST run
        # ON the tx (runPrepQuery's tx arg, 8.1+) — an autocommit read on a separate connection would
        # block on the uncommitted row lock and hang. Routing through the tx also makes them read-
        # consistent with the in-flight EIN stamp.
        # Also read VC_START_SEQ_NUMBER — the hot-call sentinel ('-1') that picks the 8HC filename branch
        # (P13, the operational sender MainMenu.pas:2715/2722) — and VC_LINE_NAME, which the operational
        # sender appends to BOTH filename branches (EDI856.LineName, MainMenu.pas:2718/2723). A normal ASN
        # carries a real seq; a hot-call carries '-1'.
        prodRows = system.db.runPrepQuery(
            "SELECT VC_PRODUCTION_DATE, VC_START_SEQ_NUMBER, VC_LINE_NAME "
            "FROM INV_ASN_MST WHERE IN_ASN_ID = ?",
            [int(asnId)], db, tx)
        if not len(prodRows):
            raise Edi856BuildError("send_856: ASN %r not found" % (asnId,))
        prodDate = str(prodRows[0]["VC_PRODUCTION_DATE"])
        startSeq = prodRows[0]["VC_START_SEQ_NUMBER"]
        startSeq = None if startSeq is None else str(startSeq).strip()
        lineName = prodRows[0]["VC_LINE_NAME"]
        lineName = "" if lineName is None else str(lineName)

        feedDs = system.db.runPrepQuery(_FEED_SQL, [int(asnId)], db, tx)
        detailRows = [{
            "Manifest": r["Manifest"],
            "PartNumber": r["PartNumber"],
            "ShipQty": int(r["ShipQty"]),
            "Kanban": r["Kanban"],
        } for r in feedDs]
        if not detailRows:
            # legacy guards on RecordCount>0 (EDI856Object.pas:113) -> 'No records to process'. An ASN
            # with zero surviving feed rows has nothing to ship; abort rather than emit an empty file.
            raise Edi856BuildError(
                "send_856: ASN %r has no surviving 856 feed rows (cost/forecast joins dropped all "
                "detail) — nothing to send" % (asnId,))

        # --- 4. build the segments (pure) + the file text -----------------------------------------
        siteDict = _read_site(db, site, tx)
        segments = build_856({"prodDate": prodDate}, detailRows, siteDict, ein, fileTime,
                             trailerId=trailerId)
        # decision D — join with CRLF only; NO segment terminator. A trailing CRLF after the last
        # segment matches the legacy Writeln (every line, incl. the last, gets a CRLF).
        text = "\r\n".join(segments) + "\r\n"
        # P13: hot-call ASNs (VC_START_SEQ_NUMBER='-1') -> '8HC<Y+MMDD><counter><LineName>.txt'; normal ->
        # '856<MMDD><LineName>.txt' (operational sender MainMenu.pas:2718/2723).
        #
        # The hot-call COUNTER (the deterministic y-equivalent). Legacy y is a per-SEND-BATCH counter:
        # init 1 before the ResendMarkedEDIs loop, INC only in the hot-call branch — so the Nth hot-call
        # file emitted in ONE batch run gets y=N (MainMenu.pas:2702/2724). The rebuild sends per-ASN (not
        # per-batch), so we reproduce "the Nth hot-call of the day for this line" deterministically:
        # counter = 1 + (count of same-day, same-line hot-call ASNs ALREADY flipped to 'S'). This ASN is
        # still 'C' at this point (the flip is step 6), so it is excluded from the count; the first hot-call
        # of the day -> counter 1 (matches legacy single-send y=1), a second -> counter 2 (no collision).
        # Read ON the tx so it is read-consistent with the in-flight flip(s).
        # GOLDEN-PENDING (P13 cutover check): byte-faithful to the source pattern, but the exact y RANGE
        # legacy actually produced (vs send order / multi-batch days) is unverified until a golden 8HC file
        # exists. The pattern + the no-collision property hold; the precise counter VALUE is the open check.
        counter = 1
        if startSeq == "-1":
            counter = 1 + int(system.db.runScalarPrepQuery(
                "SELECT COUNT(*) FROM INV_ASN_MST "
                "WHERE VC_START_SEQ_NUMBER = '-1' AND VC_ASN_STATUS = 'S' "
                "AND VC_LINE_NAME = ? AND VC_PRODUCTION_DATE = ? AND IN_ASN_ID <> ?",
                [lineName, prodDate, int(asnId)], db, tx))
        filename = _filename_856(prodDate, lineName, startSeq, counter)

        # --- 5. write the file to a .tmp (NOT the final name) — temp-then-rename atomicity --------
        # ATOMICITY (adversary blocker): the FINAL 856 file must NOT exist unless the DB committed. If
        # we wrote the final name here (step 5) and the C->S flip (step 6) or the commit then FAILED, a
        # rollback would unwind the EIN bump / stamp / flip but leave a REAL 856 on disk under an EIN the
        # DB never committed — a dispatcher would send a phantom ASN. So we write to <name>.tmp inside
        # the tx, flip + COMMIT, and only AFTER the commit succeeds do we rename <name>.tmp -> <name>.
        # A DB failure (flip/commit) rolls back AND the finally-clause deletes the orphan .tmp, so the
        # final name never appears. (A crash between commit and rename leaves only a harmless .tmp the
        # mailer ignores — never a final file the DB didn't commit.)
        path = None
        tmpPath = None
        if outDir is not None:
            import os as _os
            path = _os.path.join(outDir, filename)
            tmpPath = path + ".tmp"
            system.file.writeFile(tmpPath, text)        # write the TEMP file only (8.1+ safe)

        # --- 6. flip the ASN C->S, per-ASN, DECOUPLED ---------------------------------------------
        # NEVER REPORT_EDI856's self-flip (WHERE IN_ASN_EIN=@EIN over the wrong rows); NEVER the blanket
        # UPDATE_ASNStatus (WHERE VC_ASN_STATUS='C' across ALL sites/EINs). Just this ASN by its id.
        system.db.runPrepUpdate(
            "UPDATE INV_ASN_MST SET VC_ASN_STATUS = 'S' WHERE IN_ASN_ID = ?", [int(asnId)], db, tx=tx)

        system.db.commitTransaction(tx)
        committed = True
    except Exception:
        system.db.rollbackTransaction(tx)
        # the DB unwound — make sure no orphan .tmp (and definitely no final file) survives.
        if tmpPath is not None:
            _cleanup_tmp(tmpPath)
        raise
    finally:
        system.db.closeTransaction(tx)

    # --- 7. POST-COMMIT: rename <name>.tmp -> <name>. The final 856 exists ONLY now (DB committed). ---
    # IG83-TODO: on 8.3, prefer system.file move/rename + archive-on-success; here we use os.rename which
    #            is atomic on a single filesystem (the EDIOut dir is one volume). 8.1-safe.
    if committed and tmpPath is not None:
        import os as _os
        # os.rename replaces an existing target on POSIX; on the gateway the final name is unique per
        # send (filename carries the prod date), so no collision is expected.
        _os.rename(tmpPath, path)

    log.info("send_856 asn=%s site=%s -> EIN %09d, %d feed rows, %d segments, file=%s"
             % (asnId, site, ein, len(detailRows), len(segments), filename))
    return {"asnId": int(asnId), "ein": ein, "filename": filename, "path": path,
            "segments": segments, "text": text, "rowCount": len(detailRows)}
