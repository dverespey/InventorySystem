# edi810 — Project Library: the OUTBOUND 810 INVOICE X12 builder (M1 Rank-4), reimplemented as PURE
# Jython-2.7/CPython logic + a thin gateway driver. Twin of edi856/code.py (build_856 + send_856), and
# the same producer-recipe pattern as asn/code.py / order/code.py.
#
# BUILD TARGET — BYTE-FAITHFUL to the legacy 810 WIRE FORMAT (the segment shapes TEMA currently accepts),
# with TWO DELIBERATE, David-LOCKED divergences in the MONEY fields (see CLEAN MONEY below). The legacy
# builder is EDI810Object.pas (T810EDI; LIVE per InventorySystem.dpr:47). This module PORTS its segment
# logic (every separator/empty-element quirk reproduced — the 810 happens to be CLEAN: zero trailing
# seps); it does NOT wrap REPORT_EDI810 (which self-flips status — see report-edi810-data-analysis.md).
#
# Source truth (all on disk):
#   docs/analysis/edi/810/edi810-wire-format.md          — the byte-exact segment spec (THE target)
#   docs/analysis/edi/810/report-edi810-data-analysis.md — the data feed + the AVOID list + D6 pricing
#   EDI810Object.pas                                     — the original builder (line refs cited per segment)
#   docs/analysis/reporting/spike-report-procs-d6.sql:60-92 — the D6 REPORT_EDI810 feed (window-aware)
#
# This module holds:
#   * build_810(...)      — the PURE segment builder. Imports nothing (decimal is stdlib, Jython 2.7-safe),
#     references no gateway global at import time -> CPython-importable for the unit test
#     (scripts/e2e/test_edi810_build.py). Given the invoice feed rows, the site dict, the EIN and a file
#     timestamp it returns the ORDERED list of segment strings (the driver joins them with CRLF). No DB,
#     no I/O.
#   * create_invoice(...) — the daily-invoicing CREATE driver. ONE Inventory transaction: read the @EIN=0
#     unbilled feed, allocate the per-site EIN from INV_SITES.IN_EIN_SEQ (atomic, the SAME shared sequence
#     as the 856), INSERT_INVInfo (header, status 'S', the EIN, OUTPUT IN_INV_ID), link the detail
#     (UPDATE_INVItems set IN_INV_ID), build_810, write .tmp -> rename-after-commit.
#   * recreate_invoice(...) — re-emit an EXISTING invoice's 810 (read the @EIN<>0 feed, REUSE its EIN, NO
#     new alloc, NO status mutation on read; just rebuild the file).
#   * unsend_invoice(...)  — Carry-5 IN-PLACE unsend: revert INV_INV_MST.VC_INV_STATUS to 'C' (the
#     unsent/recreate state) + re-pool the detail (IN_INV_ID=null), KEEPING the header+EIN+audit. NOT the
#     legacy HARD-DELETE (UPDATE_INVUnsend), NOT a hard-delete.
#   All three reference `system` only at CALL time (like edi856.send_856), so importing the module never
#   needs the gateway runtime; they are driven headlessly through scripts/e2e/jython_shim.py (R8 shim).
#
# Jython 2.7 (8.1 and 8.3, no delta — pure string assembly + decimal + system.db; no Python-3-only libs).
#
# LOCKED DECISIONS realized here:
#   GS01='IN'        (856 uses 'SH') — edi810-wire-format.md:19, EDI810Object.pas:181.
#   BIG02=SupplierCode (the "invoice number"; constant per site — real uniqueness is the EIN) — :228.
#   IT101 M391/M390  by manifest '7'-prefix (broadcast vs hot-call) — EDI810Object.pas:291-294.
#   CLEAN dates off NOW (fileTime), NOT the pickup date (a 856 difference) — ISA09 yymmdd / GS04 / BIG01.
#   CRLF only         — the driver joins with "\r\n"; VC_SEP_SEGMENT is READ but NEVER emitted (legacy bug).
#   EIN %09d in all 7 control positions — allocated at invoice-CREATE from INV_SITES.IN_EIN_SEQ (the
#                     SHARED per-site counter, also used by the 856 — faithful, interleaved); REUSED at
#                     recreate (do NOT re-allocate, or the 997-ack UPDATE_EINStatus won't land).
#   SE01/CTT01        COUNTED from the emitted segments, never a magic offset (the 856 lesson).
#
# CLEAN MONEY — David LOCKED. Two legacy bugs are DELIBERATELY NOT reproduced (a wrong invoice TOTAL is a
# TEMA-reject / mis-bill; same spirit as the D6 over-bill fix):
#   IT104 unit price — legacy FloatToStr(MO_PRICE.AsCurrency) (EDI810Object.pas:298) emits a literal '.',
#     the OS-LOCALE separator, and a VARIABLE fraction count ('12.5' / '12.55' / '12.5500'). Rebuild:
#     a CLEAN fixed-scale decimal — locale-independent, consistent scale, no variable trailing. MO_PRICE
#     is money/scale-4 -> we emit scale-4 with a literal '.' (e.g. '12.5000'). DECISION 810-2. The exact
#     IT104 scale (4 vs trimmed) is flagged PENDING A GOLDEN 810; scale-4 is the consistent, defensible
#     choice (matches MO_PRICE's native scale and the TDS scale).
#   TDS01 total      — legacy hand-rolls an implied-decimal scale-4 (NO point) with BUGS (EDI810Object.pas
#     :339-350): a 1-digit fraction is NOT padded ('1234.5' -> '12345', an implied 1.2345 = OFF BY 10000x)
#     and a whole-dollar total yields a malformed string. Rebuild: TDS01 = the integer round(total*10000)
#     (correct implied-decimal scale-4, no point). DECISION 810-1. The unit test asserts the CLEAN values
#     (incl. the 1-digit-fraction case legacy got wrong) with a comment noting the legacy divergence.

from decimal import Decimal, ROUND_HALF_UP


# =================================================================================================
# PURE segment builder
# =================================================================================================

class Edi810BuildError(Exception):
    """Raised by build_810 on a structural input error (e.g. a feed row missing a required column).
    The driver rolls back on it, matching the Delphi try/except -> Execute returns False."""
    pass


def _fmt_ein(ein):
    """%9.9d — 9 chars, zero-padded (EDI810Object.pas:160 et al., EVERYWHERE the EIN appears: ISA13,
    GS06, ST02, SE02, GE02, IEA02). Python's %09d == Pascal's %9.9d for a non-negative int. IN_INV_EIN is
    int NOT NULL, unique-per-invoice (live 29-9058), so never negative. We do NOT truncate: Format('%9.9d')
    WIDENS past 9 if the EIN exceeded 999,999,999 (far outside the real range) — %09d widens identically,
    so the port is faithful at the edge too."""
    return "%09d" % int(ein)


def _fixed(s, width):
    """Pascal Format('%-Ns', [s]) — left-justify in a fixed field, space-pad to `width`. The ISA is a
    POSITIONAL segment: ISA02/04 = 10, ISA06/08 = 15 (EDI810Object.pas:142/144/146/148). Pascal %-Ns only
    PADS, never cuts -> an over-wide DUNS/abbr would push the ISA past its fixed width (a latent legacy bug).
    The wire-format spec wants the ISA byte-exact-width, so we hard-truncate to `width` AND pad: the field
    is EXACTLY `width` chars whatever the input. On the real site values (abbr <=10, DUNS 9-13, DUNS-supplier
    <=15) the truncation never fires, so it is byte-identical to legacy on real data; it only diverges on a
    pathological over-wide input legacy would have mis-emitted anyway. (Same _fixed as edi856.)"""
    s = "" if s is None else str(s)
    if len(s) > width:
        s = s[:width]                      # hard-truncate (never fires on real data)
    return s + (" " * (width - len(s)))    # left-justify, space-pad to exactly `width`


def _one_char(s, field):
    """X12 ISA15 (usage indicator, EDI810Object.pas:162 = SiteEDIMode 'P'/'T') and ISA16 (component-element
    separator, :163 = SiteSepSubElement) are single-character fields. Enforce single-char here (take the
    FIRST char); raise on empty (a missing usage indicator is a hard X12 error). Real loaded data ('P'/'T',
    '>') is already 1 char so this never alters it; it only catches a bad load (the same defect the 856 hit:
    INV_SITES.VC_EDI_MODE placeholder once carried 'PROD')."""
    s = "" if s is None else str(s)
    if len(s) < 1:
        raise Edi810BuildError("edi810: ISA %s must be a single character, got empty" % (field,))
    return s[0]


def _yymmdd(yyyymmdd):
    """ISA09 = formatdatetime('yymmdd', f810Time) = yymmdd, century dropped (EDI810Object.pas:149). UNLIKE
    GS04/BIG01 which are full yyyymmdd. We take an already-formatted 8-char yyyymmdd (the driver formats
    NOW once) and drop the century: '20260620' -> '260620'. NOTE: the 810 dates off NOW (fileTime), not the
    pickup date — a 856 difference (edi810-wire-format.md:17)."""
    yyyymmdd = str(yyyymmdd)
    if len(yyyymmdd) != 8:
        raise Edi810BuildError("edi810: file date must be yyyymmdd (8 chars), got %r" % (yyyymmdd,))
    return yyyymmdd[2:8]


# --- CLEAN MONEY helpers (David locked — legacy bugs deliberately NOT reproduced) -----------------

def _price_decimal(price):
    """Coerce a feed MO_PRICE (money/scale-4; may arrive as a Decimal, float, int or numeric string) to a
    Decimal for exact arithmetic. We never go through float for the wire string (FloatToStr was the legacy
    locale/variable-digit hazard). MO_PRICE is money so it is exact to 4 decimals."""
    if isinstance(price, Decimal):
        return price
    return Decimal(str(price))


def _it104(price):
    """IT104 unit price — DECISION 810-2 (CLEAN, locale-independent). Legacy EDI810Object.pas:298 emitted
    FloatToStr(MO_PRICE.AsCurrency): a literal '.', the OS-LOCALE separator, and a VARIABLE fraction count.
    Rebuild: a fixed scale-4 decimal with a literal '.' and NO locale separator (e.g. 12.5 -> '12.5000',
    12.55 -> '12.5500'). Scale 4 matches MO_PRICE's native money scale and the TDS scale; the exact scale
    (4 vs trimmed) is flagged PENDING A GOLDEN 810. Deterministic + portable (no FloatToStr / locale)."""
    q = _price_decimal(price).quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP)
    return "{0:.4f}".format(q)


def _tds_total(total):
    """TDS01 invoice total — DECISION 810-1 (CLEAN implied-decimal scale-4 integer). Legacy
    EDI810Object.pas:339-350 hand-rolled an implied-decimal-4 string WITH BUGS: a 1-digit fraction was NOT
    padded ('1234.5' -> '12345', an implied 1.2345 = OFF BY 10000x) and a whole-dollar total produced a
    malformed string. Rebuild: TDS01 = round(total * 10000) as an integer string (correct implied-decimal
    scale-4, no point). E.g. 1234.5 -> '12345000' (NOT the legacy '12345'); 12.55 -> '125500'; 100 ->
    '1000000' (NOT legacy's malformed whole-dollar). `total` is a Decimal (sum of price*qty, both exact)."""
    scaled = (Decimal(total) * 10000).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
    return str(int(scaled))


def build_810(invoiceRows, site, ein, fileTime):
    """Build the byte-faithful 810 invoice as an ORDERED list of segment strings. The driver joins them
    with CRLF (the segment terminator VC_SEP_SEGMENT is read into the site dict but NEVER emitted — the
    legacy CRLF-only bug, edi810-wire-format.md:9).

    Reproduces EDI810Object.pas T810EDI.Execute (the ISA..IEA envelope + the IT1 loop with per-manifest
    REF/DTM breaks), with the CLEAN money of DECISIONS 810-1/810-2.

    Parameters
    ----------
    invoiceRows : list of dicts, one per surviving feed row (the D6 REPORT_EDI810 projection), IN FEED
        ORDER (the driver passes them ORDER BY manifest so the IT1 manifest-break is deterministic — legacy
        relied on REPORT_EDI810's ORDER BY VC_MANIFEST_NUMBER). Keys (legacy EDI810DataSet field):
        'Manifest'   : d.VC_MANIFEST_NUMBER  -> REF*MK + IT101 (M391 if it starts '7' else M390). :258/291.
        'PartNumber' : d.VC_ASSY_PART_NUMBER -> IT107 (PN). :301.
        'ShipQty'    : d.IN_QTY (int)        -> IT102 (no leading zeros). :296.
        'UnitPrice'  : m.MO_PRICE (money/4)  -> IT104 (CLEAN scale-4) + the TDS total. :298/307.
        'PickUpDate' : a.VC_PRODUCTION_DATE  -> DTM*050 (yyyymmdd). :283/330. ALL rows share one pickup
                       date per file (legacy breaks the IT1 loop on a pickup-date change — the driver
                       feeds one invoice's rows, which the @EIN<>0 read returns for one EIN).
    site : dict from the site row (INV_SITES). Keys (legacy SiteDataset field -> INV_SITES column):
        'abbr'         SiteAbbr          (VC_SITE_ABBR)       -> ISA02
        'duns'         SiteDUNS          (VC_DUNS)            -> ISA04, ISA06 prefix, GS02
        'supplierCode' SiteSupplierCode  (VC_SUPPLIER_CODE)  -> ISA06 suffix (duns-supplier) + BIG02
        'tmmDuns'      SiteTMMDUNS       (VC_TMM_DUNS)        -> ISA08, GS03
        'ediMode'      SiteEDIMode       (VC_EDI_MODE)        -> ISA15 (T/P)
        'dockCode'     SiteDockCode      (VC_DOCK_CODE)       -> IT111 (ZZ <dock>) — the 810-specific field
        'sepElement'   SiteSepElement    (VC_SEP_ELEMENT)    -> the element separator ('*')
        'sepSubElement' SiteSepSubElement (VC_SEP_SUBELEMENT)-> ISA16 (sub-element sep)
        'sepSegment'   SiteSepSegment    (VC_SEP_SEGMENT)    -> READ but NEVER EMITTED (CRLF only).
    ein : int — the invoice EIN (allocated at create, reused at recreate). %09d in all 7 control positions.
    fileTime : a dict {'yyyymmdd': '20260620', 'hhmm': '1430'} — the ONE shared timestamp, captured once
        per file off NOW (legacy f810Time:=now). The pure builder takes the already-formatted strings so it
        stays I/O-free and deterministically testable. ISA09 = yymmdd of yyyymmdd; GS04/BIG01 = yyyymmdd.

    Returns the list of segment strings in emission order (EDI810Object.pas:112-120):
        ISA, GS, ST, BIG, [per-manifest break: REF*MK, DTM*050] (interior breaks), IT1(xN), trailing
        REF*MK, DTM*050, TDS, CTT, SE, GE, IEA.

    SE01 = count of segments ST..SE inclusive (counted, never a magic offset). CTT01 = #IT1 lines.
    """
    if not invoiceRows:
        # legacy guards on RecordCount>0 (EDI810Object.pas:109) -> 'No records to process'. An invoice
        # with zero surviving feed rows has nothing to bill; the driver aborts before calling us, but
        # guard here too so the pure builder never emits an empty (CTT*0 / no-IT1) envelope.
        raise Edi810BuildError("edi810: no invoice feed rows to build (nothing to bill)")

    sep = site["sepElement"]                 # element separator ('*')
    subSep = site["sepSubElement"]           # sub-element separator (ISA16 only)
    yyyymmdd = str(fileTime["yyyymmdd"])     # NOW, full yyyymmdd
    isaDate = _yymmdd(yyyymmdd)              # ISA09 = yymmdd of NOW
    hhmm = str(fileTime["hhmm"])             # the shared HHMM of NOW
    einStr = _fmt_ein(ein)

    segs = []

    # --- ISA (EDI810Object.pas:139-165) — positional, %-Ns fixed widths. ----------------------------
    # ISA*00*<abbr %-10>*01*<duns %-10>*ZZ*<duns-supplier %-15>*01*<tmmDuns %-15>*<yymmdd NOW>*<hhmm>*U*
    #    00400*<%09d EIN>*0*<ediMode>*<subSep>
    isa06 = str(site["duns"]) + "-" + str(site["supplierCode"])   # :146 — the '-' is data
    segs.append(sep.join([
        "ISA",
        "00",                               # :141 ISA01 auth-info qualifier
        _fixed(site["abbr"], 10),           # :142 ISA02 auth info (SiteAbbr, %-10s)
        "01",                               # :143 ISA03 security-info qualifier
        _fixed(site["duns"], 10),           # :144 ISA04 security info (SiteDUNS, %-10s)
        "ZZ",                               # :145 ISA05 interchange-id qualifier (sender)
        _fixed(isa06, 15),                  # :146 ISA06 sender id (DUNS-SupplierCode, %-15s)
        "01",                               # :147 ISA07 interchange-id qualifier (receiver = DUNS)
        _fixed(site["tmmDuns"], 15),        # :148 ISA08 receiver id (SiteTMMDUNS, %-15s)
        isaDate,                            # :149 ISA09 date = yymmdd of NOW (NOT pickup date)
        hhmm,                               # :150 ISA10 time
        "U",                                # :151 ISA11 interchange control-standards id
        "00400",                            # :152 ISA12 interchange control-version
        einStr,                             # :160 ISA13 interchange control # = EIN
        "0",                                # :161 ISA14 ack-requested
        _one_char(site["ediMode"], "ISA15 (usage indicator)"),   # :162 ISA15 (T/P) — single char
        _one_char(subSep, "ISA16 (sub-element separator)"),      # :163 ISA16 — single char
    ]))

    # --- GS (EDI810Object.pas:180-188) — GS01='IN' (856 uses 'SH'); GS04 = NOW yyyymmdd; GS06 = EIN. ---
    segs.append(sep.join([
        "GS",
        "IN",                               # :181 GS01 functional id = INVOICE (the 810 marker, NOT 'SH')
        str(site["duns"]),                  # :182 GS02 application sender = DUNS
        str(site["tmmDuns"]),               # :183 GS03 application receiver = TMM DUNS
        yyyymmdd,                           # :184 GS04 date = full yyyymmdd of NOW
        hhmm,                               # :185 GS05 time
        einStr,                             # :186 GS06 group control # = EIN
        "X",                                # :187 GS07 responsible-agency
        "004010",                           # :188 GS08 version
    ]))

    # --- ST (EDI810Object.pas:205-207) — ST*810*<%09d EIN>, NOT a 1-based counter. -------------------
    segs.append(sep.join(["ST", "810", einStr]))

    # --- BIG (EDI810Object.pas:226-228) — BIG*<yyyymmdd NOW>*<SiteSupplierCode>. ---------------------
    # BIG01 = invoice date (NOW); BIG02 = SiteSupplierCode (the "invoice number" — constant per site; real
    # uniqueness is the EIN in ISA13/ST02). NO trailing sep (ends on the supplier code).
    segs.append(sep.join(["BIG", yyyymmdd, str(site["supplierCode"])]))

    # --- IT1 loop (EDI810Object.pas:241-369) --------------------------------------------------------
    # Built into a buffer (IT1List in legacy) so CTT/SE counts are computed from the ACTUAL emitted
    # segments. The legacy loop:
    #   * tracks lastManifest; on a CHANGE it emits REF*MK*<PREVIOUS manifest> + DTM*050*<new PickUpDate>
    #     BEFORE the new line (the interior break, :268-287).
    #   * emits one IT1 per detail row.
    #   * after the loop, a TRAILING REF*MK*<last manifest> + DTM*050*<last PickUpDate> (:320-333).
    #   * then TDS (the invoice total).
    # NOTE the interior REF carries the PREVIOUS manifest (lastManifest), while the trailing REF carries the
    # FINAL manifest — both faithfully reproduced. dockCode is IT111 (ZZ qualifier). No trailing sep on IT1.
    it1, it1Count = [], 0
    total = Decimal("0")
    lastManifest = str(invoiceRows[0]["Manifest"])
    lastPickUp = str(invoiceRows[0]["PickUpDate"])

    for row in invoiceRows:
        for k in ("Manifest", "PartNumber", "ShipQty", "UnitPrice", "PickUpDate"):
            if k not in row:
                raise Edi810BuildError("edi810: feed row missing column %r: %r" % (k, row))
        manifest = str(row["Manifest"])
        pickUp = str(row["PickUpDate"])

        # interior manifest break (EDI810Object.pas:268-287): REF carries the PREVIOUS manifest, DTM the
        # NEW pickup date (legacy reads PickUpDate AFTER advancing lastManifest — :283 uses the current row).
        if manifest != lastManifest:
            it1.append(sep.join(["REF", "MK", lastManifest]))          # :271-273 (PREVIOUS manifest)
            it1.append(sep.join(["DTM", "050", pickUp]))               # :281-283 (current row's pickup)
            lastManifest = manifest
            lastPickUp = pickUp

        # IT1 (EDI810Object.pas:289-305). IT101 = M391 if manifest starts '7' (broadcast) else M390
        # (hot-call). IT102 = ShipQty (no leading zeros). IT104 = CLEAN unit price. IT107 = part, qualifiers
        # EA/QT/PN/PK/1/ZZ; IT111 = SiteDockCode. NO trailing sep (ends on the dock code).
        it101 = "M391" if manifest[:1] == "7" else "M390"
        qty = int(row["ShipQty"])
        price = _price_decimal(row["UnitPrice"])
        it1.append(sep.join([
            "IT1",
            it101,                          # :291-294 M391/M390 by manifest '7'-prefix
            str(qty),                       # :296 IT102 qty (no leading zeros)
            "EA",                           # :297 IT103 unit-of-measure
            _it104(price),                  # :298 IT104 unit price — CLEAN scale-4 (DECISION 810-2)
            "QT",                           # :299 IT105 basis-of-unit-price code
            "PN",                           # :300 IT106 product-id qualifier
            str(row["PartNumber"]),         # :301 IT107 part number
            "PK",                           # :302 IT108 product-id qualifier
            "1",                            # :303 IT109 (packaging count literal)
            "ZZ",                           # :304 IT110 product-id qualifier
            str(site["dockCode"]),          # :305 IT111 dock code (NO trailing sep)
        ]))
        total += price * qty               # :307 (Decimal — exact; clean money)
        it1Count += 1
        # track the last pickup date for the trailing DTM (legacy updates lastPickUp at EOF, :316).
        lastManifest = manifest
        lastPickUp = pickUp

    # trailing REF*MK*<last manifest> + DTM*050*<last PickUpDate> (EDI810Object.pas:320-333).
    it1.append(sep.join(["REF", "MK", lastManifest]))                  # :320-322
    it1.append(sep.join(["DTM", "050", lastPickUp]))                   # :328-330

    # TDS (EDI810Object.pas:337-353) — CLEAN implied-decimal scale-4 integer (DECISION 810-1).
    it1.append(sep.join(["TDS", _tds_total(total)]))

    segs.extend(it1)

    # --- CTT (EDI810Object.pas:377-378) — CTT01 = #IT1 lines (fID1Count). ----------------------------
    segs.append(sep.join(["CTT", str(it1Count)]))

    # --- SE (EDI810Object.pas:398-400) — SE01 = count of segments ST..SE INCLUSIVE (SE counts itself). -
    # Compute by counting the actual emitted segments from ST to SE, never a magic offset (the TEMA-reject
    # trap). index of ST in segs is 2 (ISA, GS, ST). Everything from ST to the CTT we already appended is
    # the ST..CTT run; +1 for the SE we are about to add. (Legacy fSECount is incremented per-segment from
    # ST through SE; counting the buffer here is the faithful equivalent.)
    st_index = 2
    se01 = (len(segs) - st_index) + 1       # segments ST..CTT already in `segs`, plus SE itself
    segs.append(sep.join(["SE", str(se01), einStr]))

    # --- GE (EDI810Object.pas:418-420) GE*1*<%09d EIN> (GE02 must = GS06). ---------------------------
    segs.append(sep.join(["GE", "1", einStr]))
    # --- IEA (EDI810Object.pas:438-440) IEA*1*<%09d EIN> (IEA02 must = ISA13). -----------------------
    segs.append(sep.join(["IEA", "1", einStr]))

    return segs


def _filename_810(pickUpDate):
    """RECREATE filename pattern (ASNInvoice.pas:872): '810' + copy(PickUpDate,5,4) + '.txt' =
    '810<mmdd>.txt' (drops the year -> cross-year collisions, noted in the wire-format spec). Pascal
    copy(s,5,4) is 1-based -> chars 5..8 of 'yyyymmdd' = MM+DD. '20260620' -> '0620' -> '8100620.txt'.
    One deterministic pattern, matching the legacy recreate path."""
    pickUpDate = str(pickUpDate)
    if len(pickUpDate) != 8:
        raise Edi810BuildError("edi810: pickup date must be yyyymmdd (8 chars), got %r" % (pickUpDate,))
    return "810" + pickUpDate[4:8] + ".txt"   # copy(s,5,4) = [4:8] 0-based


def _cleanup_tmp(tmpPath):
    """Best-effort delete of the orphan .tmp on a rolled-back send (atomicity). Never raises — the caller
    is mid-propagation of the REAL failure; we must not mask it. (Same as edi856._cleanup_tmp.)"""
    import os as _os
    try:
        if _os.path.exists(tmpPath):
            _os.remove(tmpPath)
    except Exception:
        pass


# =================================================================================================
# The gateway drivers. THREE modes, each ONE transaction.
# =================================================================================================
#
# IG81-COMPAT: every gateway API used here (runPrepQuery / runScalarPrepQuery(...,tx=) / runPrepUpdate
# (...,tx=) / beginTransaction / system.date / system.file.writeFile / getLogger) is identical on 8.1.52
# and 8.3 — no version guard needed.

DATABASE = "Inventory_Spike"            # the Inventory rebuild connection name

# ISA15 usage indicator carried in INV_SITES.VC_EDI_MODE.

# The CREATE feed (REPORT_EDI810 @EIN=0) and the RECREATE feed (REPORT_EDI810 @EIN<>0) inlined here so the
# drivers carry no external-file dependency at runtime (spike-edi810-feed.sql is the reviewable canonical
# copy; test_edi810_e2e.py enforces these strings are byte-identical to the .sql's marked bodies modulo the
# @EIN<->? / @SITEID bind, so they cannot drift). BOTH reproduce REPORT_EDI810 (NO self-flip):
#
# CREATE (@EIN=0): the unbilled read. INNER JOIN keyed on the DETAIL flag IN_INV_ID IS NULL + the header
#   VC_ASN_STATUS='A' (the analysis: the unbilled flag is on the detail). CROSS APPLY fn_ManifestCostAt for
#   the D6 window-aware MO_PRICE. ORDER BY manifest. 6-col (Manifest/Part/Price/Qty/PickUp/ASNid). NO
#   GROUP BY (unlike the 856) — one row per (header x surviving detail). Site-scoped externally (M4 adds
#   IN_SITE_ID to the ASN header; today the create reads all unbilled 'A' detail, scoped to the one site).
_CREATE_FEED_SQL = (
    "SELECT d.VC_MANIFEST_NUMBER AS Manifest, "
    "d.VC_ASSY_PART_NUMBER AS PartNumber, "
    "m.MO_PRICE AS UnitPrice, "
    "d.IN_QTY AS ShipQty, "
    "a.VC_PRODUCTION_DATE AS PickUpDate, "
    "d.IN_ASN_ID AS ASNid "
    "FROM INV_ASN_MST a "
    "JOIN INV_ASN_DETAIL_MST d "
    "ON a.IN_ASN_ID = d.IN_ASN_ID AND a.VC_ASN_STATUS = 'A' AND d.IN_INV_ID IS NULL "
    "CROSS APPLY dbo.fn_ManifestCostAt(d.VC_ASSY_PART_NUMBER, a.VC_PRODUCTION_DATE) m "
    "ORDER BY d.VC_MANIFEST_NUMBER"
)

# RECREATE (@EIN<>0): the existing-invoice read. + JOIN INV_INV_MST iim ON d.IN_INV_ID = iim.IN_INV_ID,
#   WHERE iim.IN_INV_EIN = @EIN. The proc's @EIN<>0 branch ALSO self-flips (UPDATE ... SET VC_INV_STATUS=
#   'S' WHERE IN_INV_EIN=@EIN) — we DROP that (the flip is the driver's job, decoupled). 7-col (+EIN).
_RECREATE_FEED_SQL = (
    "SELECT d.VC_MANIFEST_NUMBER AS Manifest, "
    "d.VC_ASSY_PART_NUMBER AS PartNumber, "
    "m.MO_PRICE AS UnitPrice, "
    "d.IN_QTY AS ShipQty, "
    "a.VC_PRODUCTION_DATE AS PickUpDate, "
    "d.IN_ASN_ID AS ASNid, "
    "iim.IN_INV_EIN AS InvEIN "
    "FROM INV_ASN_MST a "
    "JOIN INV_ASN_DETAIL_MST d ON a.IN_ASN_ID = d.IN_ASN_ID "
    "JOIN INV_INV_MST iim ON d.IN_INV_ID = iim.IN_INV_ID "
    "CROSS APPLY dbo.fn_ManifestCostAt(d.VC_ASSY_PART_NUMBER, a.VC_PRODUCTION_DATE) m "
    "WHERE iim.IN_INV_EIN = ? "
    "ORDER BY d.VC_MANIFEST_NUMBER"
)


def _read_site(database, siteId, tx):
    """Read the site identity row (INV_SITES) on the SAME transaction and map the columns to the dict
    build_810 expects (incl. dockCode -> IT111, the 810-specific field the 856 doesn't carry). Reading on
    the tx lets it see the EIN-allocation's own uncommitted row without blocking. (Twin of
    edi856._read_site, plus VC_DOCK_CODE.)

    NB on byte-fidelity: the spike's Inventory.INV_SITES values are PLACEHOLDER (abbr 'MAS', DUNS
    '000000001', dock 'D01', sep '>' subelem, EDI mode 'P') — NOT the real legacy values. So the BYTES this
    driver emits on the spike reflect the placeholder site row, not the TEMA wire. The test asserts
    STRUCTURE + self-consistency against whatever site row is configured; true wire bytes require the real
    site values be loaded into INV_SITES at cutover."""
    rows = system.db.runPrepQuery(
        "SELECT VC_SITE_ABBR, VC_DUNS, VC_SUPPLIER_CODE, VC_TMM_DUNS, VC_EDI_MODE, VC_DOCK_CODE, "
        "VC_SEP_ELEMENT, VC_SEP_SUBELEMENT, VC_SEP_SEGMENT "
        "FROM INV_SITES WHERE IN_SITE_ID = ?", [int(siteId)], database, tx)
    if not len(rows):
        raise Edi810BuildError("edi810: no INV_SITES row for site id %r" % (siteId,))
    r = rows[0]
    return {
        "abbr": r["VC_SITE_ABBR"],
        "duns": r["VC_DUNS"],
        "supplierCode": r["VC_SUPPLIER_CODE"],
        "tmmDuns": r["VC_TMM_DUNS"],
        "ediMode": r["VC_EDI_MODE"],
        "dockCode": r["VC_DOCK_CODE"],          # IT111 (810-specific)
        "sepElement": r["VC_SEP_ELEMENT"],
        "sepSubElement": r["VC_SEP_SUBELEMENT"],
        "sepSegment": r["VC_SEP_SEGMENT"],      # read; NEVER emitted (CRLF only)
    }


def _file_time(fileTime):
    """Resolve the ONE shared file timestamp into {'yyyymmdd', 'hhmm'} (NOW, captured once — legacy
    f810Time:=now). Accepts an explicit dict (testing determinism) OR None -> system.date.now(). The 810
    dates off NOW, not the pickup date."""
    if fileTime is not None:
        return {"yyyymmdd": str(fileTime["yyyymmdd"]), "hhmm": str(fileTime["hhmm"])}
    now = system.date.now()
    return {"yyyymmdd": system.date.format(now, "yyyyMMdd"), "hhmm": system.date.format(now, "HHmm")}


def _rows_from_feed(feedDs):
    """Map a feed PyDataset to the list of dicts the driver works with. build_810 consumes
    Manifest/Part/Qty/UnitPrice/PickUp; ASNid is carried for the per-pickup-date grouping + the scoped
    detail link (build_810 ignores extra keys). FEED ORDER is preserved (ORDER BY manifest) so the IT1
    manifest-break stays deterministic within each pickup-date group."""
    return [{
        "Manifest": r["Manifest"],
        "PartNumber": r["PartNumber"],
        "ShipQty": int(r["ShipQty"]),
        "UnitPrice": r["UnitPrice"],
        "PickUpDate": r["PickUpDate"],
        "ASNid": int(r["ASNid"]),
    } for r in feedDs]


def _group_by_pickup_date(invoiceRows):
    """Partition the unbilled feed into ONE group per distinct VC_PRODUCTION_DATE (PickUpDate), preserving
    the feed's manifest order WITHIN each group. Returns a list of (pickUpDate, rows) pairs, ORDERED by
    pickUpDate ascending so the N invoices are emitted in a stable, sequential order (N pickup dates ->
    N invoices / N EINs / N files).

    LEGACY ANCHOR (the per-pickup-date file split): MainMenu.CreateINVOICEClick (MainMenu.pas:2613-2654)
    walks the REPORT_EDI810 @EIN=0 cursor (ordered BY VC_MANIFEST_NUMBER); each EDI810.Execute consumes
    rows until the IT1 loop BREAKS on a pickup-date change (EDI810Object.pas:263-266), then writes a
    SEPARATE file + allocates a NEW EIN + INSERT_INVInfo. The break is on PickUpDate ALONE — there is no
    line dimension in the @EIN=0 feed (it returns 6 cols: Manifest/Part/Price/Qty/PickUp/ASNid, NO
    LineName — verified against CreateInventory.sql:3739-3745). So the grouping is per DISTINCT PickUpDate.

    DIVERGENCE NOTE (faithful, documented): legacy breaks on CONTIGUOUS-run boundaries in the manifest-
    ordered cursor, so if one pickup date were interleaved across non-adjacent manifests it would emit >1
    file for that date. We group by DISTINCT pickup-date value (all of a date's rows in one invoice), which
    is the natural/correct grouping and matches "N pickup dates -> N invoices". Because VC_PRODUCTION_DATE
    lives on the ASN HEADER (one date per ASN — verified: INV_ASN_DETAIL_MST has no VC_PRODUCTION_DATE),
    contiguous-run vs distinct-value can only diverge when two ASNs of the SAME date sort non-adjacently by
    manifest; collapsing them into one invoice for the date is the intended behaviour, not a regression."""
    order = []
    groups = {}
    for row in invoiceRows:
        pud = str(row["PickUpDate"])
        if pud not in groups:
            groups[pud] = []
            order.append(pud)
        groups[pud].append(row)
    return [(pud, groups[pud]) for pud in sorted(order)]


def _write_and_emit(invoiceRows, siteDict, ein, ftime, outDir):
    """Shared pure-build + .tmp-write step (used by create + recreate). Returns
    (segments, text, filename, path, tmpPath). Writes ONLY the .tmp (the caller renames post-commit)."""
    segments = build_810(invoiceRows, siteDict, ein, ftime)
    # CRLF only; a trailing CRLF after the last segment matches the legacy Writeln (every line, incl. the
    # last, gets a CRLF). The segment terminator is never emitted.
    text = "\r\n".join(segments) + "\r\n"
    filename = _filename_810(invoiceRows[0]["PickUpDate"])
    path = None
    tmpPath = None
    if outDir is not None:
        import os as _os
        path = _os.path.join(outDir, filename)
        tmpPath = path + ".tmp"
        system.file.writeFile(tmpPath, text)        # write the TEMP file only (8.1+ safe)
    return segments, text, filename, path, tmpPath


def _create_one_invoice(db, site, pickUpDate, groupRows, ftime, outDir, log):
    """Create ONE invoice for ONE pickup-date group (the per-EDI810.Execute unit of the legacy loop): its
    OWN EIN, INSERT_INVInfo header, the date-scoped detail link, build_810, .tmp -> rename-after-commit.
    ONE Inventory transaction per invoice (so a failure on date K leaves dates < K already committed — the
    legacy loop likewise committed each invoice via InsertINVInfo before moving on). Returns the per-invoice
    result dict. Re-raises on any error after rolling back + cleaning the orphan .tmp.

    Reproduces ONE iteration of MainMenu.CreateINVOICEClick:
      EDI810.EIN := SiteEIN+1 (alloc) -> AD_UpdateEIN (bump) -> file write -> InsertINVInfo
      (INSERT_INVInfo header + UPDATE_INVItems detail link). The rebuild's atomic OUTPUT alloc is the
      EIN bump+read in one statement (strictly safer than the legacy read-then-bump; faithful on the
      single-threaded daily run)."""
    asnIds = sorted(set(int(r["ASNid"]) for r in groupRows))

    tmpPath = None
    path = None
    committed = False
    tx = system.db.beginTransaction(db)
    try:
        # --- allocate the per-site EIN atomically (INV_SITES.IN_EIN_SEQ) — the SHARED 856/810 seq ----
        # Each pickup-date group burns its OWN EIN (N dates -> N sequential EINs), matching the legacy
        # per-Execute SiteEIN+1 / AD_UpdateEIN bump.
        ein = system.db.runScalarPrepQuery(
            "UPDATE INV_SITES SET IN_EIN_SEQ = IN_EIN_SEQ + 1 "
            "OUTPUT INSERTED.IN_EIN_SEQ WHERE IN_SITE_ID = ?", [int(site)], db, tx)
        if ein is None:
            raise Edi810BuildError(
                "create_invoice: no INV_SITES row for site id %r (EIN alloc)" % (site,))
        ein = int(ein)

        # --- INSERT_INVInfo: invoice header (the EIN, status 'S'), OUTPUT IN_INV_ID -----------------
        # The legacy proc returns @INVID = SCOPE_IDENTITY() (an OUTPUT param). Through the shim/JDBC we
        # capture the new id with a trailing SELECT SCOPE_IDENTITY() on the SAME tx. Reproduce the proc's
        # INSERT shape (IN_INV_EIN, 'S', VC_LAST_UPDATE=VC_ADD = the legacy yyyymmddHHMMSS add stamp,
        # 14 chars right-padded to the column's 16 — INV_INV_MST.VC_LAST_UPDATE/VC_ADD are varchar(16)).
        # Derived from the ONE resolved file time (ftime) — no extra NOW calls; HHMM + '00' seconds.
        addStamp = (ftime["yyyymmdd"] + ftime["hhmm"] + "00" + "00")[:16]   # yyyymmddHHMMSS -> 16
        invId = system.db.runScalarPrepQuery(
            "INSERT INTO INV_INV_MST (IN_INV_EIN, VC_INV_STATUS, VC_LAST_UPDATE, VC_ADD) "
            "VALUES (?, 'S', ?, ?); SELECT CAST(SCOPE_IDENTITY() AS int)",
            [ein, addStamp, addStamp], db, tx)
        if invId is None:
            raise Edi810BuildError("create_invoice: INSERT_INVInfo returned no IN_INV_ID")
        invId = int(invId)

        # --- link ONLY THIS DATE's unbilled detail to this invoice (date-scoped UPDATE_INVItems) -----
        # Legacy UPDATE_INVItems(@INVID,@ASNID) links ALL of an ASN's still-unbilled detail regardless of
        # pickup date. We SCOPE the link to THIS group's pickup date (join the ASN header for
        # VC_PRODUCTION_DATE — the date lives on INV_ASN_MST, not the detail) so a date-K invoice never
        # claims a date-K' (K'!=K) row. In practice the two coincide (one VC_PRODUCTION_DATE per ASN
        # header), so this is a correctness-preserving tightening of the legacy per-ASN link, not a
        # behaviour change on real data; it is the "keep it correct" the task calls for. Per distinct ASN
        # id (the legacy granularity), but AND-ed with the header date.
        for asnId in asnIds:
            system.db.runPrepUpdate(
                "UPDATE d SET d.IN_INV_ID = ?, d.VC_LAST_UPDATE = ? "
                "FROM INV_ASN_DETAIL_MST d JOIN INV_ASN_MST a ON a.IN_ASN_ID = d.IN_ASN_ID "
                "WHERE d.IN_INV_ID IS NULL AND d.IN_ASN_ID = ? AND a.VC_PRODUCTION_DATE = ?",
                [invId, addStamp, asnId, pickUpDate], db, tx=tx)

        # --- build + write the .tmp (NOT the final name) — ONLY this group's lines + the right DTM ---
        siteDict = _read_site(db, site, tx)
        segments, text, filename, path, tmpPath = _write_and_emit(groupRows, siteDict, ein, ftime, outDir)

        system.db.commitTransaction(tx)
        committed = True
    except Exception:
        system.db.rollbackTransaction(tx)
        if tmpPath is not None:
            _cleanup_tmp(tmpPath)
        raise
    finally:
        system.db.closeTransaction(tx)

    # --- POST-COMMIT: rename <name>.tmp -> <name>. The final 810 exists ONLY now (DB committed). ------
    # IG83-TODO: on 8.3, prefer system.file move/rename + archive-on-success; here os.rename (atomic on a
    #            single filesystem — the EDIOut dir is one volume). 8.1-safe.
    if committed and tmpPath is not None:
        import os as _os
        _os.rename(tmpPath, path)

    log.info("create_invoice site=%s pickUp=%s -> invId %s EIN %09d, %d feed rows over %d ASNs, "
             "%d segments, file=%s"
             % (site, pickUpDate, invId, ein, len(groupRows), len(asnIds), len(segments), filename))
    return {"invId": invId, "ein": ein, "asnIds": asnIds, "pickUpDate": pickUpDate, "filename": filename,
            "path": path, "segments": segments, "text": text, "rowCount": len(groupRows)}


def create_invoice(site, database=None, outDir=None, fileTime=None):
    """CREATE the daily invoices from the unbilled feed: read the @EIN=0 unbilled feed, GROUP it by pickup
    date, and emit ONE invoice (one EIN, one INSERT_INVInfo header, one date-scoped detail link, one 810
    file) PER distinct pickup date. N distinct pickup dates -> N invoices / N sequential EINs / N files.
    Each invoice is its OWN transaction (commit-per-invoice), with the .tmp-then-rename-after-commit
    atomicity preserved per file.

    Reproduces MainMenu.CreateINVOICEClick (MainMenu.pas:2613-2654): a while-not-Eof walk of the
    REPORT_EDI810 @EIN=0 cursor (ordered BY VC_MANIFEST_NUMBER) where each EDI810.Execute consumes one
    pickup-date run (the IT1 loop BREAKS on a pickup-date change, EDI810Object.pas:263-266) and then writes
    a SEPARATE file + allocates a NEW EIN + INSERT_INVInfo. The grouping is per DISTINCT PickUpDate (the
    @EIN=0 feed has NO LineName/line dimension — 6 cols, verified CreateInventory.sql:3739-3745); see
    _group_by_pickup_date for the contiguous-run-vs-distinct-value divergence note. EIN-at-create from
    INV_SITES.IN_EIN_SEQ (the SHARED per-site counter, also the 856's); the feed via the pure @EIN=0 SELECT
    (NOT the proc, which would self-flip); D6 window-aware pricing.

    Parameters
    ----------
    site     : the INV_SITES.IN_SITE_ID owning the EIN sequence. Site-scoped EIN allocation; site id
               derives from the session, NEVER a client param (multi-site D1). Today the create read is
               site-implicit (the spike is one site); M4 adds IN_SITE_ID to the ASN header to scope it.
    database : Inventory connection name (defaults to DATABASE).
    outDir   : directory to write the 810 into ([DIRECTORIES] EDIOut on the gateway). If None, the files
               are NOT written (still returns the bytes per invoice, for a dry-run / structural assertion).
    fileTime : override {'yyyymmdd','hhmm'} (testing determinism); default system.date.now() off NOW. The
               SAME fileTime stamps every invoice in the run (legacy captures f810Time once per Execute;
               on the single daily run they coincide to the minute — one shared NOW is faithful + keeps the
               control dates consistent across the batch).

    Returns {'mode':'create', 'invoiceCount', 'invoices'(list of per-invoice dicts), 'eins'(list),
             'rowCount'(total feed rows)}. Each per-invoice dict carries
             {'invId','ein','asnIds','pickUpDate','filename','path','segments','text','rowCount'}.

    Per-invoice transaction order + temp-then-rename atomicity (identical discipline to edi856.send_856):
      0. read the @EIN=0 unbilled feed ONCE (pure SELECT; NO mutation). Abort if empty (nothing to bill).
      1. group the feed by pickup date.
      then PER GROUP, in its own transaction:
        2. allocate the EIN atomically from INV_SITES.IN_EIN_SEQ (site-scoped UPDATE ... OUTPUT).
        3. INSERT_INVInfo (header: the EIN, status 'S', OUTPUT IN_INV_ID).
        4. link ONLY this date's unbilled detail (date-scoped UPDATE_INVItems via the ASN header date).
        5. build_810 (this group's lines + its DTM) + write the file TO A .tmp.
        commit.
        6. ONLY after the commit succeeds: rename <name>.tmp -> <name>.
        On ANY error in a group: rollback (its EIN bump + header insert + detail link all unwind) + delete
        the orphan .tmp + re-raise -> a rolled-back invoice can never leave a phantom 810 on disk under an
        uncommitted EIN/invoice. Earlier groups that already committed STAY committed (the legacy loop
        likewise persisted each invoice before the next).
    """
    db = database if database is not None else DATABASE
    log = system.util.getLogger("SPIKE.edi810.create_invoice")
    ftime = _file_time(fileTime)

    # --- 0. read the unbilled feed ONCE (pure SELECT; NO self-flip) -------------------------------
    # A single autocommit read (no tx) — it mutates nothing, and reading outside the per-invoice tx keeps
    # the grouping decision independent of any one invoice's transaction.
    feedDs = system.db.runPrepQuery(_CREATE_FEED_SQL, [], db)
    invoiceRows = _rows_from_feed(feedDs)
    if not invoiceRows:
        # legacy: nothing unbilled -> no invoice. Abort cleanly (no header, no EIN burned).
        raise Edi810BuildError(
            "create_invoice: no unbilled detail (the @EIN=0 feed is empty) — nothing to bill")

    # --- 1. group by pickup date -> one invoice per distinct pickup date --------------------------
    groups = _group_by_pickup_date(invoiceRows)

    invoices = []
    for pickUpDate, groupRows in groups:
        invoices.append(_create_one_invoice(db, site, pickUpDate, groupRows, ftime, outDir, log))

    eins = [inv["ein"] for inv in invoices]
    log.info("create_invoice site=%s -> %d invoice(s) over %d pickup date(s), EINs=%s, %d total feed rows"
             % (site, len(invoices), len(groups), eins, len(invoiceRows)))
    return {"mode": "create", "invoiceCount": len(invoices), "invoices": invoices, "eins": eins,
            "rowCount": len(invoiceRows)}


def recreate_invoice(ein, site, database=None, outDir=None, fileTime=None):
    """RE-EMIT an EXISTING invoice's 810 (recreate): read the @EIN<>0 feed for the invoice, REUSE its EIN
    (do NOT re-allocate, or the 997-ack UPDATE_EINStatus WHERE IN_INV_EIN=@EIN won't land), rebuild the
    file. ONE Inventory transaction (the read + file write); commit / rollback + cleanup on error.

    Reproduces ASNInvoice.pas:857-895 (the recreate path) MINUS REPORT_EDI810's @EIN<>0 self-flip
    (UPDATE INV_INV_MST SET VC_INV_STATUS='S' WHERE IN_INV_EIN=@EIN — dropped; the read is pure). NO new
    EIN, NO status mutation (the invoice is already sent; recreate just regenerates the same wire).

    Parameters
    ----------
    ein      : the EXISTING invoice's IN_INV_EIN (filters the read; emitted verbatim — NOT re-allocated).
    site     : the INV_SITES.IN_SITE_ID (for the site identity / separators / dock code).
    database / outDir / fileTime : as create_invoice.

    Returns {'mode':'recreate', 'ein', 'asnIds', 'filename', 'path', 'segments', 'text', 'rowCount'}.
    """
    db = database if database is not None else DATABASE
    log = system.util.getLogger("SPIKE.edi810.recreate_invoice")
    ftime = _file_time(fileTime)
    ein = int(ein)

    tmpPath = None
    path = None
    committed = False
    tx = system.db.beginTransaction(db)
    try:
        feedDs = system.db.runPrepQuery(_RECREATE_FEED_SQL, [ein], db, tx)
        invoiceRows = _rows_from_feed(feedDs)
        if not invoiceRows:
            raise Edi810BuildError(
                "recreate_invoice: no detail for EIN %r (the @EIN<>0 feed is empty) — nothing to re-emit"
                % (ein,))
        asnIds = sorted(set(int(r["ASNid"]) for r in feedDs))

        siteDict = _read_site(db, site, tx)
        # REUSE the EIN verbatim (do NOT allocate). build_810 emits it in all 7 control positions.
        segments, text, filename, path, tmpPath = _write_and_emit(invoiceRows, siteDict, ein, ftime, outDir)

        system.db.commitTransaction(tx)
        committed = True
    except Exception:
        system.db.rollbackTransaction(tx)
        if tmpPath is not None:
            _cleanup_tmp(tmpPath)
        raise
    finally:
        system.db.closeTransaction(tx)

    if committed and tmpPath is not None:
        import os as _os
        _os.rename(tmpPath, path)

    log.info("recreate_invoice EIN %09d site=%s -> %d feed rows over %d ASNs, %d segments, file=%s (EIN reused)"
             % (ein, site, len(invoiceRows), len(asnIds), len(segments), filename))
    return {"mode": "recreate", "ein": ein, "asnIds": asnIds, "filename": filename, "path": path,
            "segments": segments, "text": text, "rowCount": len(invoiceRows)}


def unsend_invoice(invId, database=None):
    """UNSEND an invoice IN-PLACE (Carry 5): revert INV_INV_MST.VC_INV_STATUS to 'C' (the unsent/recreate
    state) and re-pool the detail (IN_INV_ID = null), KEEPING the header + EIN + audit. ONE Inventory
    transaction.

    DELIBERATELY NOT the legacy HARD-DELETE: UPDATE_INVUnsend (:3387) does
        UPDATE INV_ASN_DETAIL_MST SET IN_INV_ID = null WHERE IN_INV_ID = @INVid;
        DELETE FROM INV_INV_MST WHERE IN_INV_ID = @INVid          -- destroys header + EIN + audit
    A commented-out line in the legacy proc shows the original intent was a status revert
        --UPDATE INV_INV_MST SET VC_INV_STATUS = 'C' WHERE IN_INV_ID = @INVid
    The Carry-5 rebuild RESTORES that intent: status -> 'C' (recreate-able), detail re-pooled, header kept.
    The EIN stays on the header so the 997-ack still has a target and the audit trail survives; the 'C'
    status is the "recreate the 810" flag (UPDATE_INVRecreate's 'C'->'S' grammar — but we NEVER run that
    blanket proc; this is per-invoice). NEVER the self-flip, NEVER UPDATE_INVRecreate.

    Parameters
    ----------
    invId    : INV_INV_MST.IN_INV_ID to unsend.
    database : Inventory connection name (defaults to DATABASE).

    Returns {'mode':'unsend', 'invId', 'ein', 'status', 'repooled'(detail rows re-pooled)}.
    """
    db = database if database is not None else DATABASE
    log = system.util.getLogger("SPIKE.edi810.unsend_invoice")
    invId = int(invId)

    tx = system.db.beginTransaction(db)
    try:
        # the invoice must exist (capture its EIN for the audit return + to prove the header survives).
        einRows = system.db.runPrepQuery(
            "SELECT IN_INV_EIN FROM INV_INV_MST WHERE IN_INV_ID = ?", [invId], db, tx)
        if not len(einRows):
            raise Edi810BuildError("unsend_invoice: no INV_INV_MST row for invoice id %r" % (invId,))
        ein = int(einRows[0]["IN_INV_EIN"])

        # re-pool the detail (IN_INV_ID = null) — same as the legacy re-pool step.
        repooled = system.db.runPrepUpdate(
            "UPDATE INV_ASN_DETAIL_MST SET IN_INV_ID = null WHERE IN_INV_ID = ?", [invId], db, tx=tx)

        # IN-PLACE status revert to 'C' (the unsent/recreate state) — KEEP the header + EIN + audit.
        # NOT the legacy DELETE. Per-invoice (WHERE IN_INV_ID=@invId), never the blanket UPDATE_INVRecreate.
        system.db.runPrepUpdate(
            "UPDATE INV_INV_MST SET VC_INV_STATUS = 'C' WHERE IN_INV_ID = ?", [invId], db, tx=tx)

        system.db.commitTransaction(tx)
    except Exception:
        system.db.rollbackTransaction(tx)
        raise
    finally:
        system.db.closeTransaction(tx)

    log.info("unsend_invoice invId=%s -> status 'C' (in-place), EIN %09d kept, detail re-pooled"
             % (invId, ein))
    return {"mode": "unsend", "invId": invId, "ein": ein, "status": "C", "repooled": repooled}
