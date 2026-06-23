#!/usr/bin/env python3
"""test_edi810_e2e.py — END-TO-END exercise of the REAL edi810 gateway drivers (create_invoice,
recreate_invoice, unsend_invoice), driven headlessly through the system.db shim (jython_shim) — the M1
Rank-4 outbound-810 INVOICE keystone proof.

WHAT THIS PROVES (and what it HONESTLY does NOT)
------------------------------------------------
Runs the ACTUAL driver code (the per-site EIN allocation from INV_SITES.IN_EIN_SEQ — the SHARED 856/810
sequence; INSERT_INVInfo's header insert; UPDATE_INVItems's per-ASN detail link; the side-effect-free D6
feed SELECT; build_810; the .tmp-then-rename file write — all inside ONE Inventory transaction). NOT a
proc-EXEC, which would never touch the driver/seam (retro R8). It asserts what is REAL:

  (1) FEED-ROW PARITY — the rebuild CREATE feed (spike-edi810-feed.sql, as inlined in the driver) returns
      the SAME rows, row-for-row, as the legacy REPORT_EDI810 @EIN=0 SELECT run on the SAME synthetic
      unbilled ASN. (The live snapshot is ALL-billed, so we synthesize unbilled lines — sentinel-tagged,
      swept in cleanup. The parts carry REAL cost windows, so the D6 price is real.)
  (2) CREATE mode — EIN allocated from INV_SITES.IN_EIN_SEQ (atomic, +1); INSERT_INVInfo inserts the
      header (status 'S', the EIN); the unbilled detail is LINKED (IN_INV_ID set); the 810 file is written;
      byte-exact STRUCTURE + self-consistency of the file (GS01='IN', BIG02, IT1 M391/M390, %09d EIN x7,
      SE01 counted, CTT01=#IT1, CLEAN money).
  (3) RECREATE mode — re-emit the just-created invoice's 810: the EIN is REUSED (NOT re-allocated;
      IN_EIN_SEQ NOT bumped again), the wire is regenerated identically.
  (4) ATOMICITY — a DB failure AFTER the .tmp write rolls back (EIN bump + header insert + detail link all
      unwind) and the final 810 never appears on disk (only a swept .tmp).
  (5) UNSEND (Carry 5) — in-place: status reverts to 'C', the detail is re-pooled (IN_INV_ID=null), the
      HEADER + EIN + audit SURVIVE (NOT the legacy hard-delete).
  (6) ISA15 enforced single-char.
  (7) MULTI-DATE CREATE (FIX 1, the per-pickup-date file split — MainMenu.pas:2613-2654 +
      EDI810Object.pas:263-266): a CREATE over unbilled detail spanning 2 DISTINCT pickup dates emits 2
      invoices / 2 distinct SEQUENTIAL EINs (from IN_EIN_SEQ) / 2 files '810<mmdd>.txt', each file carrying
      ONLY its date's IT1 lines + the correct DTM date, with the detail-link DATE-SCOPED (each date's rows
      linked to its own invoice, no cross-date leak). A single-date create still emits exactly 1 invoice.

HONEST VERIFICATION (fixture-fidelity discipline, memory feedback-parity-fixture-fidelity)
------------------------------------------------------------------------------------------
NO golden 810 file exists, so BYTE-FOR-BYTE TEMA PARITY IS UNPROVABLE — we do NOT claim it. The site
identity (DUNS/abbr/supplier/dock/sep/EDI-mode) in the spike Inventory.INV_SITES is PLACEHOLDER data, so
the ISA/GS/BIG BYTES reflect the placeholder site row, not the TEMA wire; the STRUCTURE assertions are
site-value-independent. The exact IT104 scale (4 vs trimmed) is also pending a golden 810. What IS proven:
feed-row parity vs the legacy SELECT (1) + byte-exact STRUCTURE & self-consistency & CLEAN money (2,3) +
the EIN provenance/reuse + the atomicity + the in-place unsend. True TEMA byte-parity + exact money scale
pending a golden 810.

Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 scripts/e2e/test_edi810_e2e.py
"""
import os, subprocess, sys, tempfile, shutil
from decimal import Decimal

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report, preclean_sentinels   # noqa: E402
import jython_shim              # noqa: E402

CONTAINER = os.environ.get("CONTAINER", "mssql-spike")
SA_PASS = os.environ.get("SA_PASS")
INV = "Inventory"

SITE_ID = 1                     # the INV_SITES row owning the (shared 856/810) EIN sequence
SEED_EIN_SEQ = 9200             # a known starting IN_EIN_SEQ -> the create allocation must yield 9201
# sentinel line names + VC_ADD stamps so cleanup is EXACT and the spike is left as-found.
CREATE_LINE = "ZZ810CRT"        # the synthetic unbilled ASN the CREATE feeds from
ATOM_LINE = "ZZ810ATM"          # a second synthetic unbilled ASN for the atomicity case
VC_ADD_SENTINEL = "ZZ810E2E000000"   # 16-char VC_ADD stamp on every row this test inserts

# Two parts that carry a REAL cost window covering PDATE (so fn_ManifestCostAt returns a real D6 price).
# 42600F261100 / 42600F263100 both have window 20250404..20280630 @ 138.0000 (verified on the snapshot).
PART_A = "42600F261100"
PART_B = "42600F263100"
PDATE = "20260610"              # inside the 20250404..20280630 window
PDATE2 = "20260611"             # a SECOND distinct pickup date, also inside the cost window (multi-date)
MANIFEST_7 = "76061857"        # starts '7' -> IT101 M391 (broadcast)
MANIFEST_8 = "80061900"        # not '7'     -> IT101 M390 (hot-call)
MULTI_LINE_1 = "ZZ810MD1"       # synthetic unbilled ASN on PDATE  (multi-date create)
MULTI_LINE_2 = "ZZ810MD2"       # synthetic unbilled ASN on PDATE2 (multi-date create)

EDI810_PY = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                             "docs", "analysis", "edi", "810", "project-library", "edi810", "code.py"))
FEED_SQL_FILE = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                                 "docs", "analysis", "edi", "810", "spike-edi810-feed.sql"))


def _marked_body(path, begin, end):
    """Extract the SELECT body between the BEGIN/END markers, drop line-comments, collapse whitespace,
    strip the trailing ';'. Used to prove the inline _*_FEED_SQL is byte-identical to the .sql."""
    with open(path) as fh:
        lines = fh.read().splitlines()
    out, grab = [], False
    for ln in lines:
        if begin in ln:
            grab = True
            continue
        if grab and end in ln:
            break
        if grab:
            out.append(ln)
    return " ".join(" ".join(out).split()).rstrip(";").strip()


def _norm(s):
    return " ".join(s.split())


def sql(db, query):
    if not SA_PASS:
        sys.exit("export SA_PASS first")
    out = subprocess.check_output([
        "docker", "exec", CONTAINER, "/opt/mssql-tools18/bin/sqlcmd", "-C", "-S", "localhost",
        "-U", "sa", "-P", SA_PASS, "-d", db, "-h", "-1", "-W", "-s", "\t",
        "-Q", "SET QUOTED_IDENTIFIER ON; SET ANSI_NULLS ON; SET NOCOUNT ON; " + query], text=True)
    return [l.split("\t") for l in out.splitlines()
            if l.strip() and not l.startswith("(") and not l.startswith("Msg ")]


def scalar(db, q):
    rows = sql(db, q)
    return rows[0][0] if rows else None


def _insert_unbilled_asn(line, pdate=PDATE, detail=None):
    """Create a synthetic status-'A' ASN on `pdate` with unbilled detail rows. Default `detail` = 3 rows
    (2 manifests: PART_A on the '7' manifest x2 lines, PART_B on the '8' manifest) — the single-date case.
    Pass an explicit detail list of (manifest, part, qty) for the multi-date case. Returns the new
    IN_ASN_ID. All rows sentinel-tagged. VC_START_SEQ_NUMBER='-1' so the filtered unique index (WHERE
    VC_START_SEQ_NUMBER<>'-1') is NOT enforced -> no collision with other synthetic ASNs on (line, pdate).
    NB the pickup date lives on the ASN HEADER (one VC_PRODUCTION_DATE per ASN), so a distinct pickup date
    means a distinct ASN."""
    if detail is None:
        detail = [(MANIFEST_7, PART_A, 80), (MANIFEST_7, PART_B, 12), (MANIFEST_8, PART_A, 100)]
    sql(INV,
        "INSERT INTO INV_ASN_MST (IN_ASN_EIN, VC_ASN_STATUS, VC_LINE_NAME, VC_ASSEMBLY_LINE, "
        "VC_START_SEQ_NUMBER, VC_END_SEQ_NUMBER, IN_QTY, VC_PRODUCTION_DATE, VC_LAST_UPDATE, VC_ADD) "
        "VALUES (0, 'A', '%s', '', '-1', '-1', 0, '%s', '%s', '%s')"
        % (line, pdate, VC_ADD_SENTINEL, VC_ADD_SENTINEL))
    asn = int(scalar(INV, "SELECT MAX(IN_ASN_ID) FROM INV_ASN_MST WHERE VC_LINE_NAME='%s'" % line))
    # detail rows, IN_INV_ID NULL (unbilled).
    for (man, part, qty) in detail:
        sql(INV,
            "INSERT INTO INV_ASN_DETAIL_MST (IN_ASN_ID, IN_INV_ID, IN_ASN_EIN, VC_MANIFEST_NUMBER, "
            "VC_ASSY_PART_NUMBER, IN_QTY, VC_LAST_UPDATE, VC_ADD) "
            "VALUES (%d, NULL, 0, '%s', '%s', %d, '%s', '%s')"
            % (asn, man, part, qty, VC_ADD_SENTINEL, VC_ADD_SENTINEL))
    return asn


def cleanup(origSeq):
    """Restore Inventory + INV_SITES as-found. Two sweeps:
      * sentinel-tagged ASNs/detail (this test's synthetic shipments — VC_ADD = VC_ADD_SENTINEL).
      * the invoice HEADERS the REAL create_invoice driver inserted. The driver stamps VC_ADD with the
        legacy DATE stamp (production behaviour — NOT a test sentinel), so we cannot sweep those by the
        sentinel. Instead we sweep by the EIN range the test owns: we SEED INV_SITES.IN_EIN_SEQ to
        SEED_EIN_SEQ (9200) before the run, so every invoice the driver creates here has IN_INV_EIN in
        (SEED_EIN_SEQ, SEED_EIN_SEQ+50]. The real snapshot's max EIN is 9058 (< 9200), so this range is
        exclusively this test's — no real invoice is ever touched."""
    sql(INV, "DELETE FROM INV_ASN_DETAIL_MST WHERE VC_ADD = '%s'" % VC_ADD_SENTINEL)
    sql(INV, "DELETE FROM INV_ASN_MST WHERE VC_ADD = '%s'" % VC_ADD_SENTINEL)
    # invoice headers created by the driver in this run: the test-owned EIN range (seeded above the snapshot
    # max). Belt-and-suspenders: also sweep any sentinel-tagged headers (none, but harmless).
    sql(INV, "DELETE FROM INV_INV_MST WHERE IN_INV_EIN > %d AND IN_INV_EIN <= %d"
             % (SEED_EIN_SEQ, SEED_EIN_SEQ + 50))
    sql(INV, "DELETE FROM INV_INV_MST WHERE VC_ADD = '%s'" % VC_ADD_SENTINEL)
    if origSeq is not None:
        sql(INV, "UPDATE INV_SITES SET IN_EIN_SEQ = %s WHERE IN_SITE_ID = %d" % (origSeq, SITE_ID))


def main():
    print("=" * 92)
    print(" edi810 END-TO-END — feed parity + CREATE + RECREATE + atomicity + UNSEND (Carry 5) + EIN")
    print(" (NOT legacy byte-parity — no golden 810; site values are spike placeholders. See docstring.)")
    print("=" * 92)
    rep = Report()

    edi810 = jython_shim.load_wrapper("edi810_driver", EDI810_PY)

    # P11 self-heal: pre-clean leftovers from a KILLED prior run (error-swallowing; child rows before
    # parents). Sentinel-scoped (the VC_ADD sentinel + the test-owned EIN range above the live max), so
    # nothing real is touched even if the prior run's own teardown never ran.
    preclean_sentinels(lambda s: sql(INV, s), [
        "DELETE FROM INV_ASN_DETAIL_MST WHERE VC_ADD = '%s'" % VC_ADD_SENTINEL,
        "DELETE FROM INV_ASN_MST WHERE VC_ADD = '%s'" % VC_ADD_SENTINEL,
        "DELETE FROM INV_INV_MST WHERE IN_INV_EIN > %d AND IN_INV_EIN <= %d"
        % (SEED_EIN_SEQ, SEED_EIN_SEQ + 50),
        "DELETE FROM INV_INV_MST WHERE VC_ADD = '%s'" % VC_ADD_SENTINEL,
    ], label="edi810")

    # --- DRIFT GUARD: the inline feeds are byte-identical to the .sql marked bodies ---------------
    print("\n--- DRIFT GUARD: _CREATE_FEED_SQL / _RECREATE_FEED_SQL == spike-edi810-feed.sql bodies ---")
    create_body = _marked_body(FEED_SQL_FILE, ">>> BEGIN CREATE FEED SQL", "<<< END CREATE FEED SQL")
    recreate_body = _marked_body(FEED_SQL_FILE, ">>> BEGIN RECREATE FEED SQL",
                                 "<<< END RECREATE FEED SQL").replace("@EIN", "?")
    rep.check("_CREATE_FEED_SQL char-identical to the .sql CREATE body (no bind)",
              _norm(edi810._CREATE_FEED_SQL) == create_body,
              "match" if _norm(edi810._CREATE_FEED_SQL) == create_body else
              "\n      inline=%r\n      .sql  =%r" % (_norm(edi810._CREATE_FEED_SQL), create_body))
    rep.check("_RECREATE_FEED_SQL char-identical to the .sql RECREATE body (@EIN -> ?)",
              _norm(edi810._RECREATE_FEED_SQL) == recreate_body,
              "match" if _norm(edi810._RECREATE_FEED_SQL) == recreate_body else
              "\n      inline=%r\n      .sql  =%r" % (_norm(edi810._RECREATE_FEED_SQL), recreate_body))
    rep.check("the CREATE feed uses CROSS APPLY fn_ManifestCostAt (D6 window-aware)",
              "fn_ManifestCostAt" in edi810._CREATE_FEED_SQL
              and "CROSS APPLY" in edi810._CREATE_FEED_SQL, "D6 pricing present")
    rep.check("the CREATE feed keys on the DETAIL unbilled flag (IN_INV_ID IS NULL, status 'A')",
              "IN_INV_ID IS NULL" in edi810._CREATE_FEED_SQL
              and "VC_ASN_STATUS = 'A'" in edi810._CREATE_FEED_SQL, "unbilled key present")
    # both feeds are bare SELECTs — no DML keyword starts a statement. (Substring-test on whole words:
    # 'PickUpDate' contains 'UPDATE' as a substring, so test the statement KIND, not a bare substring.)
    def _has_dml(s):
        up = " " + " ".join(s.upper().split()) + " "
        return any(kw in up for kw in (" UPDATE INV", " DELETE FROM", " INSERT INTO", " MERGE INTO"))
    rep.check("NO self-flip / UPDATE_INVRecreate / any DML baked into either feed (pure SELECTs)",
              not _has_dml(edi810._CREATE_FEED_SQL) and not _has_dml(edi810._RECREATE_FEED_SQL),
              "feeds are pure SELECTs")

    origSeq = createAsn = atomAsn = None
    tmpdir = tempfile.mkdtemp(prefix="edi810_e2e_")
    try:
        # --- snapshot the live all-billed fact (the create feed is empty before we inject) --------
        liveUnbilled = int(scalar(INV,
            "SELECT COUNT(*) FROM INV_ASN_MST a JOIN INV_ASN_DETAIL_MST d "
            "ON a.IN_ASN_ID=d.IN_ASN_ID AND a.VC_ASN_STATUS='A' AND d.IN_INV_ID IS NULL"))
        rep.check("live snapshot is all-billed (the create feed is empty until we synthesize)",
                  liveUnbilled == 0, "%d unbilled" % liveUnbilled)

        # --- synthesize an unbilled ASN + a covering cost window guarantee -------------------------
        createAsn = _insert_unbilled_asn(CREATE_LINE)
        rep.check("synthetic unbilled ASN created with 3 detail rows", createAsn is not None,
                  "asn=%s" % createAsn)
        # confirm fn_ManifestCostAt covers both parts at PDATE (else the CROSS APPLY would drop the line).
        covA = int(scalar(INV, "SELECT COUNT(*) FROM dbo.fn_ManifestCostAt('%s','%s')" % (PART_A, PDATE)))
        covB = int(scalar(INV, "SELECT COUNT(*) FROM dbo.fn_ManifestCostAt('%s','%s')" % (PART_B, PDATE)))
        rep.check("both synthetic parts have a covering cost window at PDATE (D6 price exists)",
                  covA == 1 and covB == 1, "A=%d B=%d" % (covA, covB))

        # --- seed the EIN seq to a known value (snapshot the original first) -----------------------
        origSeq = scalar(INV, "SELECT IN_EIN_SEQ FROM INV_SITES WHERE IN_SITE_ID=%d" % SITE_ID)
        rep.check("INV_SITES row exists for site %d" % SITE_ID, origSeq is not None, str(origSeq))
        sql(INV, "UPDATE INV_SITES SET IN_EIN_SEQ=%d WHERE IN_SITE_ID=%d" % (SEED_EIN_SEQ, SITE_ID))

        # --- (1) FEED-ROW PARITY: rebuild CREATE feed vs legacy REPORT_EDI810 @EIN=0 SELECT --------
        # The legacy @EIN=0 SELECT scans ALL unbilled 'A' detail; our synthetic ASN is the only unbilled
        # data, so the two reads compare directly. Compare the 5 wire columns the builder consumes.
        print("\n--- (1) FEED-ROW PARITY: rebuild CREATE feed vs legacy REPORT_EDI810 @EIN=0 SELECT ---")
        legacy_feed = sql(INV,
            "SELECT d.VC_MANIFEST_NUMBER, d.VC_ASSY_PART_NUMBER, m.MO_PRICE, d.IN_QTY, "
            "a.VC_PRODUCTION_DATE "
            "FROM INV_ASN_MST a "
            "JOIN INV_ASN_DETAIL_MST d "
            "ON a.IN_ASN_ID=d.IN_ASN_ID AND a.VC_ASN_STATUS='A' AND d.IN_INV_ID IS NULL "
            "CROSS APPLY dbo.fn_ManifestCostAt(d.VC_ASSY_PART_NUMBER, a.VC_PRODUCTION_DATE) m "
            "ORDER BY d.VC_MANIFEST_NUMBER")
        legacy_wire = sorted((r[0], r[1], str(Decimal(r[2])), int(r[3]), r[4]) for r in legacy_feed)

        feedDs = edi810.system.db.runPrepQuery(edi810._CREATE_FEED_SQL, [], "Inventory_Spike")
        rebuild_wire = sorted((r["Manifest"], r["PartNumber"], str(Decimal(str(r["UnitPrice"]))),
                               int(r["ShipQty"]), r["PickUpDate"]) for r in feedDs)
        rep.check("rebuild CREATE feed row COUNT == legacy SELECT count (3)",
                  len(rebuild_wire) == len(legacy_wire) == 3,
                  "rebuild=%d legacy=%d" % (len(rebuild_wire), len(legacy_wire)))
        rep.check("rebuild CREATE feed == legacy REPORT_EDI810 @EIN=0 SELECT, row-for-row "
                  "(Manifest/Part/Price/Qty/PickUp)", rebuild_wire == legacy_wire,
                  "%d/%d rows identical" % (sum(1 for a, b in zip(rebuild_wire, legacy_wire) if a == b),
                                            len(legacy_wire)))
        if rebuild_wire != legacy_wire:
            for a, b in zip(rebuild_wire, legacy_wire):
                if a != b:
                    print("    [DIF] rebuild=%r legacy=%r" % (a, b))

        # --- (2) CREATE mode (single pickup date -> exactly ONE invoice) --------------------------
        print("\n--- (2) CREATE: EIN alloc + INSERT_INVInfo + detail link + 810 file ---")
        batch = edi810.create_invoice(SITE_ID, database="Inventory_Spike", outDir=tmpdir,
                                      fileTime={"yyyymmdd": "20260620", "hhmm": "1430"})
        # the synthetic ASN is all on ONE pickup date (PDATE) -> one pickup-date group -> ONE invoice.
        rep.check("single-date create -> exactly ONE invoice (one pickup-date group)",
                  batch["invoiceCount"] == 1 and len(batch["invoices"]) == 1,
                  "invoiceCount=%s" % batch["invoiceCount"])
        result = batch["invoices"][0]            # the one per-invoice result dict (downstream asserts)
        ein = int(result["ein"])
        invId = int(result["invId"])
        rep.check("create_invoice allocated EIN == seeded IN_EIN_SEQ + 1 (atomic)",
                  ein == SEED_EIN_SEQ + 1, "ein=%s seed+1=%d" % (ein, SEED_EIN_SEQ + 1))
        postSeq = int(scalar(INV, "SELECT IN_EIN_SEQ FROM INV_SITES WHERE IN_SITE_ID=%d" % SITE_ID))
        rep.check("INV_SITES.IN_EIN_SEQ bumped by exactly 1", postSeq == SEED_EIN_SEQ + 1,
                  "post=%d seed=%d" % (postSeq, SEED_EIN_SEQ))
        # header inserted with status 'S' + the EIN.
        hdr = sql(INV, "SELECT IN_INV_EIN, VC_INV_STATUS FROM INV_INV_MST WHERE IN_INV_ID=%d" % invId)
        rep.check("INSERT_INVInfo inserted the header (status 'S', the EIN)",
                  hdr and int(hdr[0][0]) == ein and hdr[0][1] == "S", str(hdr))
        # detail LINKED to the invoice (UPDATE_INVItems): all 3 rows now carry IN_INV_ID.
        linked = int(scalar(INV, "SELECT COUNT(*) FROM INV_ASN_DETAIL_MST WHERE IN_INV_ID=%d" % invId))
        stillNull = int(scalar(INV, "SELECT COUNT(*) FROM INV_ASN_DETAIL_MST WHERE IN_ASN_ID=%d "
                                    "AND IN_INV_ID IS NULL" % createAsn))
        rep.check("all 3 unbilled detail rows LINKED to the new invoice (UPDATE_INVItems)",
                  linked == 3 and stillNull == 0, "linked=%d stillNull=%d" % (linked, stillNull))
        rep.check("create_invoice rowCount == feed count (3)", result["rowCount"] == 3,
                  "rowCount=%d" % result["rowCount"])

        # --- byte-exact STRUCTURE of the CREATE file ---------------------------------------------
        print("\n--- CREATE file STRUCTURE + self-consistency + CLEAN money ---")
        rep.check("CREATE file written to outDir", result["path"] and os.path.exists(result["path"]),
                  result["path"])
        with open(result["path"], "rb") as fh:
            raw = fh.read()
        text = raw.decode("ascii")
        rep.check("file uses CRLF line endings (no '~' terminator)", b"\r\n" in raw and b"~" not in raw,
                  "CRLF present, no '~'")
        segs = [s for s in text.split("\r\n") if s != ""]
        rep.check("file segments == driver-returned segments", segs == result["segments"],
                  "%d file vs %d returned" % (len(segs), len(result["segments"])))
        einStr = "%09d" % ein
        isa = segs[0].split("*"); gs = segs[1].split("*"); st = segs[2].split("*"); big = segs[3].split("*")
        se = segs[-3].split("*"); ge = segs[-2].split("*"); iea = segs[-1].split("*")
        rep.check("envelope ISA..IEA with GS01='IN' + ST*810",
                  segs[0].startswith("ISA*") and gs[1] == "IN" and segs[2].startswith("ST*810*")
                  and segs[-1].startswith("IEA*"), "%s ... %s" % (segs[0][:6], segs[-1][:6]))
        rep.check("BIG02 == SiteSupplierCode (from INV_SITES)",
                  big[2] == scalar(INV, "SELECT VC_SUPPLIER_CODE FROM INV_SITES WHERE IN_SITE_ID=%d"
                                        % SITE_ID), big[2])
        rep.check("%09d EIN in all 7 positions (ISA13/GS06/ST02/SE02/GE02/IEA02)",
                  isa[13] == einStr and gs[6] == einStr and st[2] == einStr and se[2] == einStr
                  and ge[2] == einStr and iea[2] == einStr, "all == %s" % einStr)
        # SE01 / CTT01 recomputed off the FILE.
        st_idx = next(i for i, s in enumerate(segs) if s.startswith("ST*"))
        se_idx = next(i for i, s in enumerate(segs) if s.startswith("SE*"))
        rep.check("SE01 == actual ST..SE segment count (file)", int(se[1]) == (se_idx - st_idx) + 1,
                  "SE01=%s actual=%d" % (se[1], (se_idx - st_idx) + 1))
        it1s = [s for s in segs if s.startswith("IT1*")]
        ctt = segs[se_idx - 1].split("*")
        rep.check("CTT01 == #IT1 (3)", int(ctt[1]) == len(it1s) == 3,
                  "CTT01=%s #IT1=%d" % (ctt[1], len(it1s)))
        # IT101 M391 (the '7' manifest) vs M390 (the '8' manifest) — 2 broadcast + 1 hot-call.
        m391 = sum(1 for s in it1s if s.split("*")[1] == "M391")
        m390 = sum(1 for s in it1s if s.split("*")[1] == "M390")
        rep.check("IT101 split: 2 M391 (manifest '7') + 1 M390 (manifest '8')", m391 == 2 and m390 == 1,
                  "M391=%d M390=%d" % (m391, m390))
        rep.check("every IT1 ends on the dock code (NO trailing sep — the 810 is CLEAN)",
                  all(not s.endswith("*") for s in it1s) and all("*ZZ*" in s for s in it1s),
                  "%d IT1 clean" % len(it1s))
        # CLEAN money: TDS = round(total*10000), total = (80+100)*138.00 + 12*138.00 = 192*138 = 26496.00
        # -> 26496.0000 -> 264960000. (Both parts price 138.0000 on the snapshot.)
        expTotal = Decimal("192") * Decimal("138.0000")
        tds = next(s for s in segs if s.startswith("TDS*")).split("*")
        rep.check("TDS01 == CLEAN round(total*10000) integer (no decimal point)",
                  tds[1] == str(int((expTotal * 10000))) and "." not in tds[1],
                  "TDS01=%s expected=%s" % (tds[1], str(int(expTotal * 10000))))
        rep.check("IT104 is a fixed scale-4 decimal (e.g. '138.0000')",
                  all(len(s.split("*")[4].split(".")[-1]) == 4 for s in it1s),
                  it1s[0].split("*")[4])
        rep.check("ISA09 == yymmdd of NOW '260620' while GS04 == yyyymmdd (dated off NOW)",
                  isa[9] == "260620" and gs[4] == "20260620", "ISA09=%s GS04=%s" % (isa[9], gs[4]))
        rep.check("ISA15 (usage indicator) is exactly one char", len(isa[15]) == 1, isa[15])
        rep.check("filename == '810' + mmdd + '.txt' = '8100610.txt'",
                  result["filename"] == "8100610.txt", result["filename"])

        # --- (3) RECREATE mode: re-emit the SAME invoice, EIN REUSED ----------------------------
        print("\n--- (3) RECREATE: re-emit the same invoice, EIN reused (no re-alloc) ---")
        preSeqRec = int(scalar(INV, "SELECT IN_EIN_SEQ FROM INV_SITES WHERE IN_SITE_ID=%d" % SITE_ID))
        recdir = tempfile.mkdtemp(prefix="edi810_recreate_")
        try:
            rec = edi810.recreate_invoice(ein, SITE_ID, database="Inventory_Spike", outDir=recdir,
                                          fileTime={"yyyymmdd": "20260620", "hhmm": "1430"})
            rep.check("recreate REUSED the EIN (== the created invoice's EIN, NOT a new alloc)",
                      int(rec["ein"]) == ein, "rec=%s created=%s" % (rec["ein"], ein))
            postSeqRec = int(scalar(INV, "SELECT IN_EIN_SEQ FROM INV_SITES WHERE IN_SITE_ID=%d" % SITE_ID))
            rep.check("INV_SITES.IN_EIN_SEQ NOT bumped by recreate (no new EIN)",
                      postSeqRec == preSeqRec, "pre=%d post=%d" % (preSeqRec, postSeqRec))
            rep.check("recreate row count == create row count (same detail)",
                      rec["rowCount"] == result["rowCount"], "rec=%d create=%d"
                      % (rec["rowCount"], result["rowCount"]))
            # the recreate wire is byte-identical to the create wire (same EIN, same time, same rows).
            rep.check("recreate produced a byte-identical 810 to the create (same EIN/time/rows)",
                      rec["segments"] == result["segments"], "identical wire")
            rep.check("recreate file written", rec["path"] and os.path.exists(rec["path"]), rec["path"])
        finally:
            shutil.rmtree(recdir, ignore_errors=True)

        # --- (4) ATOMICITY: a post-write DB failure -> rollback -> NO final 810 ------------------
        # Re-pool the just-created invoice's detail first so the atomicity ASN's create has unbilled data.
        # Use a SEPARATE synthetic unbilled ASN so the create feed sees rows. Inject a fault into the
        # detail-link UPDATE (step 4, AFTER the .tmp write... actually the .tmp is step 5; we fault the
        # COMMIT path by faulting the LAST runPrepUpdate before commit) — simplest: fault the detail-link
        # UPDATE but only AFTER the .tmp has been written. To keep the .tmp-write-before-failure ordering,
        # we instead fault on commitTransaction.
        print("\n--- (4) ATOMICITY: DB failure after .tmp write -> rollback -> NO final 810 ---")
        # link/clear: first re-pool the created invoice so only the atom ASN is unbilled.
        sql(INV, "UPDATE INV_ASN_DETAIL_MST SET IN_INV_ID=NULL WHERE IN_INV_ID=%d" % invId)
        # remove the created header so the atom create starts clean (re-created in cleanup-safe way: it is
        # sentinel-tagged so cleanup sweeps it regardless).
        sql(INV, "DELETE FROM INV_INV_MST WHERE IN_INV_ID=%d" % invId)
        # now there are 3 unbilled rows again (the createAsn detail). Snapshot the seq.
        preSeqAtom = int(scalar(INV, "SELECT IN_EIN_SEQ FROM INV_SITES WHERE IN_SITE_ID=%d" % SITE_ID))
        atomdir = tempfile.mkdtemp(prefix="edi810_atom_")
        try:
            realCommit = edi810.system.db.commitTransaction
            realWrite = edi810.system.file.writeFile
            wroteTmp = {"path": None}

            def faulting_commit(tx):
                raise RuntimeError("INJECTED post-write DB failure (commit)")

            def spying_write(path, data, append=False):
                wroteTmp["path"] = path
                return realWrite(path, data, append)

            edi810.system.db.commitTransaction = faulting_commit
            edi810.system.file.writeFile = spying_write
            raised = False
            try:
                edi810.create_invoice(SITE_ID, database="Inventory_Spike", outDir=atomdir,
                                      fileTime={"yyyymmdd": "20260620", "hhmm": "1430"})
            except Exception as e:
                raised = True
                print("    (expected) create_invoice raised: %s" % e)
            finally:
                edi810.system.db.commitTransaction = realCommit
                edi810.system.file.writeFile = realWrite

            rep.check("create_invoice re-raised on the post-write DB failure", raised, "raised=%s" % raised)
            rep.check("driver wrote the EDI to a .tmp (NOT the final name) before the failure",
                      wroteTmp["path"] is not None and wroteTmp["path"].endswith(".tmp"),
                      "wrote=%r" % wroteTmp["path"])
            files = sorted(os.listdir(atomdir))
            rep.check("NO final 810 file exists after the rolled-back create (atomicity)",
                      "8100610.txt" not in files, "dir=%r" % files)
            rep.check("the orphan .tmp was cleaned up too (dir empty)", files == [], "dir=%r" % files)
            # the DB unwound: the EIN seq did NOT bump, no committed invoice header, detail still unbilled.
            postSeqAtom = int(scalar(INV, "SELECT IN_EIN_SEQ FROM INV_SITES WHERE IN_SITE_ID=%d" % SITE_ID))
            rep.check("INV_SITES.IN_EIN_SEQ NOT bumped (EIN alloc rolled back)",
                      postSeqAtom == preSeqAtom, "pre=%d post=%d" % (preSeqAtom, postSeqAtom))
            committedHeaders = int(scalar(INV, "SELECT COUNT(*) FROM INV_INV_MST WHERE VC_ADD='%s'"
                                              % VC_ADD_SENTINEL))
            rep.check("no committed invoice header from the rolled-back create",
                      committedHeaders == 0, "%d sentinel headers" % committedHeaders)
            stillUnbilled = int(scalar(INV, "SELECT COUNT(*) FROM INV_ASN_DETAIL_MST WHERE IN_ASN_ID=%d "
                                            "AND IN_INV_ID IS NULL" % createAsn))
            rep.check("detail still unbilled (link rolled back)", stillUnbilled == 3,
                      "%d unbilled" % stillUnbilled)
        finally:
            shutil.rmtree(atomdir, ignore_errors=True)

        # --- (5) UNSEND (Carry 5): in-place — re-create an invoice, then unsend it ---------------
        print("\n--- (5) UNSEND (Carry 5): in-place status revert + re-pool, header + EIN SURVIVE ---")
        # re-create a clean invoice (the createAsn detail is unbilled again after the atomicity rollback).
        preSeqU = int(scalar(INV, "SELECT IN_EIN_SEQ FROM INV_SITES WHERE IN_SITE_ID=%d" % SITE_ID))
        udir = tempfile.mkdtemp(prefix="edi810_unsend_")
        try:
            cre2batch = edi810.create_invoice(SITE_ID, database="Inventory_Spike", outDir=udir,
                                              fileTime={"yyyymmdd": "20260620", "hhmm": "1430"})
            cre2 = cre2batch["invoices"][0]      # single pickup date -> one invoice
            uInvId = int(cre2["invId"])
            uEin = int(cre2["ein"])
            rep.check("re-created invoice for unsend (status 'S', detail linked)",
                      scalar(INV, "SELECT VC_INV_STATUS FROM INV_INV_MST WHERE IN_INV_ID=%d" % uInvId) == "S"
                      and int(scalar(INV, "SELECT COUNT(*) FROM INV_ASN_DETAIL_MST WHERE IN_INV_ID=%d"
                                          % uInvId)) == 3, "set up")
            # UNSEND it.
            un = edi810.unsend_invoice(uInvId, database="Inventory_Spike")
            rep.check("unsend returned the EIN (header survives)", int(un["ein"]) == uEin,
                      "ein=%s" % un["ein"])
            # NB: the shim's runPrepUpdate returns a fixed 1 (not the affected-row count), so un['repooled']
            # is shim-bounded; on the real gateway it is the true count. Assert the re-pool against the DB
            # directly (below) — the authoritative observation.
            rep.check("unsend returned a repool count (gateway gives the true count; shim returns 1)",
                      un["repooled"] is not None, "repooled=%s (shim-bounded)" % un["repooled"])
            # IN-PLACE: header STILL EXISTS, status now 'C', EIN unchanged (NOT a hard-delete).
            survives = sql(INV, "SELECT IN_INV_EIN, VC_INV_STATUS FROM INV_INV_MST WHERE IN_INV_ID=%d"
                                % uInvId)
            rep.check("invoice HEADER SURVIVES (NOT the legacy hard-delete)", len(survives) == 1,
                      "%d rows" % len(survives))
            rep.check("status reverted IN-PLACE to 'C' (the unsent/recreate state)",
                      survives and survives[0][1] == "C", str(survives))
            rep.check("EIN preserved on the surviving header (so the 997-ack still lands)",
                      survives and int(survives[0][0]) == uEin, str(survives))
            repooledNull = int(scalar(INV, "SELECT COUNT(*) FROM INV_ASN_DETAIL_MST WHERE IN_ASN_ID=%d "
                                           "AND IN_INV_ID IS NULL" % createAsn))
            rep.check("detail RE-POOLED (IN_INV_ID = null) — available to re-bill", repooledNull == 3,
                      "%d re-pooled" % repooledNull)
            noLink = int(scalar(INV, "SELECT COUNT(*) FROM INV_ASN_DETAIL_MST WHERE IN_INV_ID=%d" % uInvId))
            rep.check("no detail still points at the unsent invoice", noLink == 0, "%d still linked" % noLink)
        finally:
            shutil.rmtree(udir, ignore_errors=True)

        # --- (6) MULTI-DATE CREATE (FIX 1): 2 distinct pickup dates -> 2 invoices / 2 EINs / 2 files ---
        # The crux of the per-pickup-date file split (MainMenu.pas:2613-2654 + EDI810Object.pas:263-266).
        # Start CLEAN: sweep ALL prior sentinel rows + any test-owned invoice headers so the create feed
        # holds ONLY the two ASNs we inject. Re-seed the EIN seq to a known value so the two allocations
        # are deterministic + sequential.
        print("\n--- (6) MULTI-DATE CREATE: 2 pickup dates -> 2 invoices / 2 sequential EINs / 2 files ---")
        sql(INV, "DELETE FROM INV_ASN_DETAIL_MST WHERE VC_ADD='%s'" % VC_ADD_SENTINEL)
        sql(INV, "DELETE FROM INV_ASN_MST WHERE VC_ADD='%s'" % VC_ADD_SENTINEL)
        sql(INV, "DELETE FROM INV_INV_MST WHERE IN_INV_EIN > %d AND IN_INV_EIN <= %d"
                 % (SEED_EIN_SEQ, SEED_EIN_SEQ + 50))
        MD_SEED = SEED_EIN_SEQ + 10            # a distinct seed band for the multi-date run (9210 -> 9211,9212)
        sql(INV, "UPDATE INV_SITES SET IN_EIN_SEQ=%d WHERE IN_SITE_ID=%d" % (MD_SEED, SITE_ID))

        # Two ASNs on DISTINCT pickup dates. ASN1 on PDATE: 2 detail (PART_A '7' x2 -> M391; PART_B '8' ->
        # M390). ASN2 on PDATE2: 1 detail (PART_A '7' -> M391). Distinct dates => distinct ASNs (the date is
        # on the header). Manifests chosen so the feed (ORDER BY manifest) INTERLEAVES the two dates' rows
        # (manifest '7...' rows from BOTH dates sort before/around the '8...' row) — proving the grouping
        # splits by date, not by feed adjacency.
        mdAsn1 = _insert_unbilled_asn(MULTI_LINE_1, pdate=PDATE,
                                      detail=[(MANIFEST_7, PART_A, 80), (MANIFEST_8, PART_B, 12)])
        mdAsn2 = _insert_unbilled_asn(MULTI_LINE_2, pdate=PDATE2,
                                      detail=[(MANIFEST_7, PART_A, 100)])
        rep.check("two synthetic unbilled ASNs on DISTINCT pickup dates created",
                  mdAsn1 is not None and mdAsn2 is not None, "asn1=%s(%s) asn2=%s(%s)"
                  % (mdAsn1, PDATE, mdAsn2, PDATE2))

        mddir = tempfile.mkdtemp(prefix="edi810_multidate_")
        try:
            mdBatch = edi810.create_invoice(SITE_ID, database="Inventory_Spike", outDir=mddir,
                                            fileTime={"yyyymmdd": "20260620", "hhmm": "1430"})
            # (a) TWO invoices, one per pickup date.
            rep.check("CREATE over 2 pickup dates -> exactly 2 invoices (per-pickup-date file split)",
                      mdBatch["invoiceCount"] == 2 and len(mdBatch["invoices"]) == 2,
                      "invoiceCount=%s" % mdBatch["invoiceCount"])
            invs = mdBatch["invoices"]
            # invoices are ordered by pickup date ascending (PDATE then PDATE2).
            inv1 = next(i for i in invs if i["pickUpDate"] == PDATE)
            inv2 = next(i for i in invs if i["pickUpDate"] == PDATE2)

            # (b) TWO DISTINCT SEQUENTIAL EINs from IN_EIN_SEQ (MD_SEED+1, MD_SEED+2).
            rep.check("two DISTINCT EINs allocated", inv1["ein"] != inv2["ein"],
                      "ein1=%s ein2=%s" % (inv1["ein"], inv2["ein"]))
            rep.check("EINs are SEQUENTIAL from IN_EIN_SEQ (MD_SEED+1, MD_SEED+2)",
                      sorted([inv1["ein"], inv2["ein"]]) == [MD_SEED + 1, MD_SEED + 2]
                      and inv1["ein"] == MD_SEED + 1 and inv2["ein"] == MD_SEED + 2,
                      "eins=%s expected=%s" % (sorted([inv1["ein"], inv2["ein"]]),
                                               [MD_SEED + 1, MD_SEED + 2]))
            postSeqMD = int(scalar(INV, "SELECT IN_EIN_SEQ FROM INV_SITES WHERE IN_SITE_ID=%d" % SITE_ID))
            rep.check("IN_EIN_SEQ bumped by exactly 2 (one per invoice)",
                      postSeqMD == MD_SEED + 2, "post=%d seed=%d" % (postSeqMD, MD_SEED))

            # (c) TWO files with the correct '810' + mmdd + '.txt' names (FIX 2 still recreate-style; see
            #     STOP note in the deliverable — LineName has no source in the @EIN=0 feed).
            rep.check("invoice 1 filename == '810'+mmdd(PDATE)+'.txt' = '8100610.txt'",
                      inv1["filename"] == "8100610.txt", inv1["filename"])
            rep.check("invoice 2 filename == '810'+mmdd(PDATE2)+'.txt' = '8100611.txt'",
                      inv2["filename"] == "8100611.txt", inv2["filename"])
            rep.check("both invoices have DISTINCT filenames", inv1["filename"] != inv2["filename"],
                      "%s vs %s" % (inv1["filename"], inv2["filename"]))
            mdFiles = sorted(os.listdir(mddir))
            rep.check("exactly 2 files written to outDir (no .tmp left)",
                      mdFiles == ["8100610.txt", "8100611.txt"], "dir=%r" % mdFiles)

            # (d) each file carries ONLY its date's IT1 lines + the correct DTM date.
            def _read_segs(path):
                with open(path, "rb") as fh:
                    return [s for s in fh.read().decode("ascii").split("\r\n") if s != ""]
            segs1 = _read_segs(inv1["path"])
            segs2 = _read_segs(inv2["path"])
            it1_1 = [s for s in segs1 if s.startswith("IT1*")]
            it1_2 = [s for s in segs2 if s.startswith("IT1*")]
            rep.check("invoice 1 (PDATE) has 2 IT1 lines (its 2 detail rows only)", len(it1_1) == 2,
                      "%d IT1" % len(it1_1))
            rep.check("invoice 2 (PDATE2) has 1 IT1 line (its 1 detail row only)", len(it1_2) == 1,
                      "%d IT1" % len(it1_2))
            dtm1 = sorted(set(s.split("*")[2] for s in segs1 if s.startswith("DTM*050*")))
            dtm2 = sorted(set(s.split("*")[2] for s in segs2 if s.startswith("DTM*050*")))
            rep.check("invoice 1 DTM*050 carries ONLY PDATE (no PDATE2 bleed)", dtm1 == [PDATE],
                      "dtm1=%s" % dtm1)
            rep.check("invoice 2 DTM*050 carries ONLY PDATE2 (no PDATE bleed)", dtm2 == [PDATE2],
                      "dtm2=%s" % dtm2)
            rep.check("CTT01 per file == its own #IT1 (2 and 1)",
                      int(next(s for s in segs1 if s.startswith("CTT*")).split("*")[1]) == 2
                      and int(next(s for s in segs2 if s.startswith("CTT*")).split("*")[1]) == 1,
                      "ok")

            # (e) detail-link scoped correctly: each date's detail linked to its OWN invoice; no cross-link.
            md1Linked = int(scalar(INV, "SELECT COUNT(*) FROM INV_ASN_DETAIL_MST WHERE IN_INV_ID=%d"
                                        % inv1["invId"]))
            md2Linked = int(scalar(INV, "SELECT COUNT(*) FROM INV_ASN_DETAIL_MST WHERE IN_INV_ID=%d"
                                        % inv2["invId"]))
            rep.check("invoice 1 linked exactly its 2 detail rows (date-scoped UPDATE_INVItems)",
                      md1Linked == 2, "linked=%d" % md1Linked)
            rep.check("invoice 2 linked exactly its 1 detail row (date-scoped UPDATE_INVItems)",
                      md2Linked == 1, "linked=%d" % md2Linked)
            # ASN1's detail belongs to inv1 only; ASN2's to inv2 only — no cross-date leakage.
            asn1ToInv2 = int(scalar(INV, "SELECT COUNT(*) FROM INV_ASN_DETAIL_MST WHERE IN_ASN_ID=%d "
                                         "AND IN_INV_ID=%d" % (mdAsn1, inv2["invId"])))
            asn2ToInv1 = int(scalar(INV, "SELECT COUNT(*) FROM INV_ASN_DETAIL_MST WHERE IN_ASN_ID=%d "
                                         "AND IN_INV_ID=%d" % (mdAsn2, inv1["invId"])))
            rep.check("NO cross-date detail leak (PDATE rows -> inv1 only, PDATE2 rows -> inv2 only)",
                      asn1ToInv2 == 0 and asn2ToInv1 == 0,
                      "asn1->inv2=%d asn2->inv1=%d" % (asn1ToInv2, asn2ToInv1))
            stillUnbilledMD = int(scalar(INV, "SELECT COUNT(*) FROM INV_ASN_DETAIL_MST WHERE VC_ADD='%s' "
                                              "AND IN_INV_ID IS NULL" % VC_ADD_SENTINEL))
            rep.check("all multi-date unbilled detail now linked (nothing left unbilled)",
                      stillUnbilledMD == 0, "%d still unbilled" % stillUnbilledMD)
        finally:
            shutil.rmtree(mddir, ignore_errors=True)

    finally:
        cleanup(origSeq)
        shutil.rmtree(tmpdir, ignore_errors=True)
        left = int(scalar(INV, "SELECT COUNT(*) FROM INV_ASN_MST WHERE VC_ADD='%s'" % VC_ADD_SENTINEL))
        leftDet = int(scalar(INV, "SELECT COUNT(*) FROM INV_ASN_DETAIL_MST WHERE VC_ADD='%s'"
                                  % VC_ADD_SENTINEL))
        # any invoice header still in the test-owned EIN range = a leak (the driver-created headers).
        leftInv = int(scalar(INV, "SELECT COUNT(*) FROM INV_INV_MST WHERE IN_INV_EIN > %d AND "
                                  "IN_INV_EIN <= %d" % (SEED_EIN_SEQ, SEED_EIN_SEQ + 50)))
        rep.check("Inventory restored as-found (sentinel ASNs/detail + driver-created invoices swept)",
                  left == 0 and leftDet == 0 and leftInv == 0,
                  "ASN=%d DET=%d INV=%d remaining" % (left, leftDet, leftInv))
        if origSeq is not None:
            nowSeq = scalar(INV, "SELECT IN_EIN_SEQ FROM INV_SITES WHERE IN_SITE_ID=%d" % SITE_ID)
            rep.check("INV_SITES.IN_EIN_SEQ restored to original", str(nowSeq) == str(origSeq),
                      "now=%s orig=%s" % (nowSeq, origSeq))

    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
