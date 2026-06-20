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
#   B  INNER forecast join for Kanban, capped to one row (CROSS APPLY (SELECT TOP 1 ...)) — feed SQL.
#   C  GROUP-BY collapse (NO sum) — feed SQL.
#   D  emit CRLF ONLY (NO '~' segment terminator) — the driver joins with "\r\n"; the segment terminator
#      VC_SEP_SEGMENT is READ into the site dict but NEVER emitted (legacy Trap 1 / EDI856Object.pas:13).
#   E  filename = the CREATE pattern  856<copy(prodDate,4,5)>.txt  — _filename_856() below.
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
        str(site["ediMode"]),               # :161 ISA15 usage indicator (T/P)
        subSep,                             # :162 ISA16 component-element separator
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
    hl.append(sep.join(["TD1", ""]))                       # :294 TD1** (empty)
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
        # LIN**BP*<PartNumber>*RC*<Kanban>  (:347 — LIN01 empty)
        hl.append(sep.join(["LIN", "", "BP", str(row["PartNumber"]), "RC", str(row["Kanban"])]))
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


def _filename_856(prodDate):
    """Decision E — the CREATE filename pattern (ASNSelect.pas:457):
        856 + copy(ProductionDate,4,5) + '.txt'
    Pascal copy(s,4,5) is 1-based -> chars 4..8 of 'yyyymmdd' = last-year-digit + MM + DD (5 chars).
    '20260618' -> '60618' -> '85660618.txt'? No: '856' + '60618' = '85660618.txt'. (The RECREATE
    variant ASNInvoice.pas:817 uses copy(PickupDate,5,4) = MM+DD = a DIFFERENT offset; the build spec
    locks the CREATE pattern as the ONE deterministic choice — decision E.)"""
    prodDate = str(prodDate)
    if len(prodDate) != 8:
        raise Edi856BuildError("edi856: production date must be yyyymmdd, got %r" % (prodDate,))
    return "856" + prodDate[3:8] + ".txt"   # copy(s,4,5) = [3:8] 0-based


# =================================================================================================
# send_856 — the GATEWAY DRIVER. ONE transaction.
# =================================================================================================
#
# IG81-COMPAT: every gateway API used here (runPrepQuery / runScalarPrepQuery(...,tx=) / runPrepUpdate
# (...,tx=) / beginTransaction / system.date / system.file.writeFile / getLogger) is identical on
# 8.1.52 and 8.3 — no version guard needed.

DATABASE = "Inventory_Spike"            # the Inventory rebuild connection name

# The feed SQL (spike-edi856-feed.sql) inlined here so the driver carries no external-file dependency at
# runtime (the .sql file is the reviewable canonical copy; this string MUST stay byte-identical to its
# SELECT body). Parameterized by @ASNID (NOT the 6440 literal); NO self-flip; INNER cost (A) + CROSS
# APPLY TOP 1 forecast (B) + inclusive cost window + GROUP-BY collapse (C). 5-col projection (the 856
# emits no price; the legacy 9-col feed's UnitPrice/SiteEIN/StartSeq/LineName are not consumed by the
# builder — we read only what build_856 needs, and read the header prodDate from the header row).
_FEED_SQL = (
    "SELECT d.VC_MANIFEST_NUMBER AS Manifest, "
    "d.VC_ASSY_PART_NUMBER AS PartNumber, "
    "d.IN_QTY AS ShipQty, "
    "a.VC_PRODUCTION_DATE AS PickUpDate, "
    "f.VC_ASSY_KANBAN_NUMBER AS Kanban "
    "FROM INV_ASN_MST a "
    "JOIN INV_ASN_DETAIL_MST d ON a.IN_ASN_ID = d.IN_ASN_ID "
    "JOIN INV_MANIFEST_COST_MST m ON d.VC_ASSY_PART_NUMBER = m.VC_ASSY_PART_NUMBER_CODE "
    "CROSS APPLY (SELECT TOP 1 f1.VC_ASSY_KANBAN_NUMBER "
    "            FROM INV_FORECAST_DETAIL_INF f1 "
    "            WHERE f1.VC_ASSY_PART_NUMBER_CODE = d.VC_ASSY_PART_NUMBER) f "
    "WHERE a.IN_ASN_ID = ? "
    "AND m.VC_START_MANIFEST <= a.VC_PRODUCTION_DATE "
    "AND m.VC_END_MANIFEST   >= a.VC_PRODUCTION_DATE "
    "GROUP BY d.VC_MANIFEST_NUMBER, d.VC_ASSY_PART_NUMBER, d.IN_QTY, "
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

    Transaction order (ONE tx):
      1. allocate EIN atomically from INV_SITES.IN_EIN_SEQ (site-scoped UPDATE ... OUTPUT the bumped
         value) — the at-send increment. NEVER the 6440 literal.
      2. stamp IN_ASN_EIN on the ASN header.
      3. read the feed (pure SELECT, @ASNID-parameterized; NO self-flip) + the header prodDate.
      4. build_856(...) (pure).
      5. write the file (temp-then-rename atomic on the gateway path; here a single writeFile to outDir).
      6. flip the ASN VC_ASN_STATUS C->S, per-ASN (WHERE IN_ASN_ID = @asnId) — DECOUPLED.
      commit; rollback + re-raise on any error (nothing partially sent — the EIN bump, the stamp, and
      the flip all unwind together; the file write is the only non-transactional step, ordered LAST
      before the flip so a write failure rolls the DB back).
    """
    db = database if database is not None else DATABASE
    log = system.util.getLogger("SPIKE.send_856")
    if fileTime is None:
        # ONE shared HHMM for the whole file (legacy f810Time captured once). system.date is gateway-only,
        # referenced at CALL time so import stays runtime-free.
        fileTime = system.date.format(system.date.now(), "HHmm")

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
        prodRows = system.db.runPrepQuery(
            "SELECT VC_PRODUCTION_DATE FROM INV_ASN_MST WHERE IN_ASN_ID = ?", [int(asnId)], db, tx)
        if not len(prodRows):
            raise Edi856BuildError("send_856: ASN %r not found" % (asnId,))
        prodDate = str(prodRows[0]["VC_PRODUCTION_DATE"])

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
        filename = _filename_856(prodDate)

        # --- 5. write the file (LAST non-DB step before the flip) ---------------------------------
        path = None
        if outDir is not None:
            # gateway path is [DIRECTORIES] EDIOut. system.file.writeFile is 8.1+ safe. We write the
            # final name directly here; the production gateway path should temp-then-rename + archive-
            # on-success for atomic file I/O (IG83-TODO below).
            # IG83-TODO: on 8.3, prefer the temp-then-rename + archive-on-success atomic pattern (write
            #            <name>.tmp, fsync, rename to <name>) so a partial file is never picked up by the
            #            mailer. Kept a single writeFile here for the 8.1-safe spike.
            import os as _os
            path = _os.path.join(outDir, filename)
            system.file.writeFile(path, text)

        # --- 6. flip the ASN C->S, per-ASN, DECOUPLED ---------------------------------------------
        # NEVER REPORT_EDI856's self-flip (WHERE IN_ASN_EIN=@EIN over the wrong rows); NEVER the blanket
        # UPDATE_ASNStatus (WHERE VC_ASN_STATUS='C' across ALL sites/EINs). Just this ASN by its id.
        system.db.runPrepUpdate(
            "UPDATE INV_ASN_MST SET VC_ASN_STATUS = 'S' WHERE IN_ASN_ID = ?", [int(asnId)], db, tx=tx)

        system.db.commitTransaction(tx)
    except Exception:
        system.db.rollbackTransaction(tx)
        raise
    finally:
        system.db.closeTransaction(tx)

    log.info("send_856 asn=%s site=%s -> EIN %09d, %d feed rows, %d segments, file=%s"
             % (asnId, site, ein, len(detailRows), len(segments), filename))
    return {"asnId": int(asnId), "ein": ein, "filename": filename, "path": path,
            "segments": segments, "text": text, "rowCount": len(detailRows)}
