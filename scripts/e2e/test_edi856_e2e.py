#!/usr/bin/env python3
"""test_edi856_e2e.py — END-TO-END exercise of the REAL edi856.send_856 gateway driver, driven
headlessly through the system.db shim (jython_shim) — the M1 Rank-2 outbound-856 keystone proof.

WHAT THIS PROVES (and what it HONESTLY does NOT)
------------------------------------------------
Runs the ACTUAL driver code (the per-site EIN allocation from INV_SITES.IN_EIN_SEQ, the header EIN
stamp, the side-effect-free feed SELECT, build_856, the file write, the per-ASN C->S flip — all inside
ONE Inventory transaction) — NOT a proc-EXEC, which would never touch the driver/seam (retro R8). It
asserts what is REAL:

  (1) FEED-ROW PARITY — the rebuild feed (spike-edi856-feed.sql, as inlined in the driver) returns the
      SAME rows, row-for-row, as the legacy REPORT_EDI856 @EIN=0 PREVIEW SELECT run on the SAME ASN
      (status forced 'C' so the preview branch fires). This proves the read is faithful to the proc
      MINUS the proc's hazards (no self-flip, no 6440 literal). A real cross-check against the legacy
      SELECT, not self-assertion.

  (2) BYTE-EXACT STRUCTURE + SELF-CONSISTENCY of the emitted file — segment order; the %09d EIN in all
      7 positions; ISA13 == GS06 == ST02 == SE02 == GE02 == IEA02; SE01 == the actual ST..SE segment
      count; CTT01 == the HL count; one Order HL per distinct manifest with PRF; one Item HL per feed
      row with LIN/SN1; CRLF-only (no '~'); the filename pattern.

  (3) THE EIN CAME FROM INV_SITES — the EIN stamped on the header and emitted in the file equals the
      pre-send INV_SITES.IN_EIN_SEQ + 1 (allocated at send, atomic, site-scoped), and INV_SITES.IN_EIN_SEQ
      was bumped by exactly 1. NOT the 6440 literal; NOT EIN-at-create.

  (4) THE FLIP IS PER-ASN, DECOUPLED — the target ASN went 'C' -> 'S'; a SECOND 'C' ASN that is NOT the
      target stayed 'C' (proving we did NOT use the blanket UPDATE_ASNStatus WHERE VC_ASN_STATUS='C').

HONEST VERIFICATION (fixture-fidelity discipline, memory feedback-parity-fixture-fidelity)
------------------------------------------------------------------------------------------
NO golden 856 file exists for this ASN, so BYTE-FOR-BYTE LEGACY PARITY IS UNPROVABLE — we do NOT claim
it. Two further caveats stated plainly:
  * The site identity (DUNS/abbr/supplier/TMM/EDI-mode/separators) in the spike Inventory.INV_SITES is
    PLACEHOLDER data (the real legacy site identity lives in VehicleOrder.Site, relocated to INV_SITES at
    the sites-CRUD step but not yet loaded with the real values). So the ISA/GS BYTES reflect the
    placeholder site row, not the TEMA wire. The STRUCTURE assertions are site-value-independent; the
    exact ISA bytes will only be wire-correct once the real site values are loaded.
  * Legacy ASN 4721 was frozen under a different forecast-recipe vintage (see test_create_asn_parity.py),
    but THAT does not affect the 856 BUILD: the 856 reads the ASN's STORED detail rows verbatim (it does
    not re-run the fan-out), so copying 4721's exact detail rows gives the builder the real shipped lines.

What IS proven: feed-row parity vs the legacy SELECT (1) + byte-exact STRUCTURE & self-consistency (2) +
the EIN provenance (3) + the decoupled per-ASN flip (4). Byte-parity pending a golden 856.

METHOD
------
  1. Copy ASN 4721's header (status forced 'C', EIN 0) + its 16 detail rows from Inventory_Live into the
     rebuild Inventory under a NEW identity id (no collision). Also create a SECOND 'C' decoy ASN to
     prove the flip is per-ASN.
  2. Snapshot + set INV_SITES.IN_EIN_SEQ to a known value for the target site so the allocation is
     observable, then restore it at the end.
  3. FEED PARITY: run the legacy REPORT_EDI856 @EIN=0 preview SELECT (filtered to the copied ASN) and the
     rebuild feed on the SAME ASN; diff row-for-row.
  4. Run the REAL send_856 (writes the file to a temp dir).
  5. Structurally validate the emitted file; assert the EIN provenance + the per-ASN/decoupled flip.
  6. Restore Inventory + INV_SITES as-found (delete the copied ASNs, reset IN_EIN_SEQ).

Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 scripts/e2e/test_edi856_e2e.py
"""
import os, subprocess, sys, tempfile, shutil

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report          # noqa: E402
import jython_shim              # noqa: E402

CONTAINER = os.environ.get("CONTAINER", "mssql-spike")
SA_PASS = os.environ.get("SA_PASS")
INV = "Inventory"
LIVE = "Inventory_Live"

LEGACY_ASN = 4721               # COROLLA / 20260618 / 16 detail rows
LINE = "COROLLA"
PDATE = "20260618"
SITE_ID = 1                     # the INV_SITES row owning the EIN sequence
SEED_EIN_SEQ = 9100             # a known starting IN_EIN_SEQ -> allocation must yield 9101
# a sentinel line name for the copied/decoy ASNs so cleanup is exact and nothing real is touched.
TEST_LINE = "ZZ856E2E"
DECOY_LINE = "ZZ856DEC"

EDI856_PY = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                             "docs", "analysis", "edi", "856", "project-library", "edi856", "code.py"))


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


def cleanup(testAsn, decoyAsn, origSeq):
    """Restore Inventory + INV_SITES as-found."""
    sql(INV, "DELETE FROM INV_ASN_DETAIL_MST WHERE IN_ASN_ID IN "
             "(SELECT IN_ASN_ID FROM INV_ASN_MST WHERE VC_LINE_NAME IN ('%s','%s')); "
             "DELETE FROM INV_ASN_MST WHERE VC_LINE_NAME IN ('%s','%s')"
             % (TEST_LINE, DECOY_LINE, TEST_LINE, DECOY_LINE))
    if origSeq is not None:
        sql(INV, "UPDATE INV_SITES SET IN_EIN_SEQ = %s WHERE IN_SITE_ID = %d"
                 % (origSeq, SITE_ID))


def main():
    print("=" * 88)
    print(" send_856 END-TO-END — feed-row parity + byte-exact STRUCTURE + EIN provenance + decoupled flip")
    print(" (NOT legacy byte-parity — no golden 856; site values are spike placeholders. See docstring.)")
    print("=" * 88)
    rep = Report()

    edi856 = jython_shim.load_wrapper("edi856_driver", EDI856_PY)

    # --- pre-clean any leftovers from a prior aborted run ---------------------------------------
    sql(INV, "DELETE FROM INV_ASN_DETAIL_MST WHERE IN_ASN_ID IN "
             "(SELECT IN_ASN_ID FROM INV_ASN_MST WHERE VC_LINE_NAME IN ('%s','%s')); "
             "DELETE FROM INV_ASN_MST WHERE VC_LINE_NAME IN ('%s','%s')"
             % (TEST_LINE, DECOY_LINE, TEST_LINE, DECOY_LINE))

    # --- 0. legacy ground truth present ---------------------------------------------------------
    lcount = scalar(LIVE, "SELECT COUNT(*) FROM INV_ASN_DETAIL_MST WHERE IN_ASN_ID=%d" % LEGACY_ASN)
    rep.check("legacy ASN %d present in Inventory_Live with detail" % LEGACY_ASN, int(lcount) > 0,
              "%s detail rows" % lcount)

    testAsn = decoyAsn = origSeq = None
    tmpdir = tempfile.mkdtemp(prefix="edi856_e2e_")
    try:
        # --- 1. copy the header (status 'C', EIN 0) under the sentinel line, capture new id ------
        # cross-DB copy: Inventory_Live -> Inventory. Force status 'C' so the @EIN=0 preview branch
        # fires AND the C->S flip is observable. Keep the REAL detail rows (the 856 ships stored detail).
        sql(INV,
            "INSERT INTO INV_ASN_MST (IN_ASN_EIN, VC_ASN_STATUS, VC_LINE_NAME, VC_ASSEMBLY_LINE, "
            "VC_START_SEQ_NUMBER, DT_START_SEQ, VC_END_SEQ_NUMBER, DT_END_SEQ, IN_QTY, "
            "VC_PRODUCTION_DATE, VC_LAST_UPDATE, VC_ADD) "
            "SELECT 0, 'C', '%s', VC_ASSEMBLY_LINE, VC_START_SEQ_NUMBER, DT_START_SEQ, "
            "VC_END_SEQ_NUMBER, DT_END_SEQ, IN_QTY, VC_PRODUCTION_DATE, VC_LAST_UPDATE, VC_ADD "
            "FROM %s.dbo.INV_ASN_MST WHERE IN_ASN_ID=%d" % (TEST_LINE, LIVE, LEGACY_ASN))
        testAsn = int(scalar(INV, "SELECT IN_ASN_ID FROM INV_ASN_MST WHERE VC_LINE_NAME='%s'" % TEST_LINE))
        sql(INV,
            "INSERT INTO INV_ASN_DETAIL_MST (IN_ASN_ID, IN_INV_ID, IN_ASN_EIN, VC_MANIFEST_NUMBER, "
            "VC_ASSY_PART_NUMBER, IN_QTY, VC_LAST_UPDATE, VC_ADD) "
            "SELECT %d, NULL, 0, VC_MANIFEST_NUMBER, VC_ASSY_PART_NUMBER, IN_QTY, VC_LAST_UPDATE, VC_ADD "
            "FROM %s.dbo.INV_ASN_DETAIL_MST WHERE IN_ASN_ID=%d" % (testAsn, LIVE, LEGACY_ASN))
        copied = int(scalar(INV, "SELECT COUNT(*) FROM INV_ASN_DETAIL_MST WHERE IN_ASN_ID=%d" % testAsn))
        rep.check("copied ASN %d into Inventory with all detail rows" % testAsn, copied == int(lcount),
                  "%d copied vs %s legacy" % (copied, lcount))

        # --- a SECOND 'C' decoy ASN (proves the flip is per-ASN, not blanket) -------------------
        sql(INV,
            "INSERT INTO INV_ASN_MST (IN_ASN_EIN, VC_ASN_STATUS, VC_LINE_NAME, VC_ASSEMBLY_LINE, "
            "VC_START_SEQ_NUMBER, VC_END_SEQ_NUMBER, IN_QTY, VC_PRODUCTION_DATE, VC_LAST_UPDATE, VC_ADD) "
            "VALUES (0, 'C', '%s', '', '0001', '0002', 0, '%s', '0000000000000000', '0000000000000000')"
            % (DECOY_LINE, PDATE))
        decoyAsn = int(scalar(INV, "SELECT IN_ASN_ID FROM INV_ASN_MST WHERE VC_LINE_NAME='%s'" % DECOY_LINE))

        # --- 2. seed INV_SITES.IN_EIN_SEQ to a known value (snapshot the original first) ---------
        origSeq = scalar(INV, "SELECT IN_EIN_SEQ FROM INV_SITES WHERE IN_SITE_ID=%d" % SITE_ID)
        rep.check("INV_SITES row exists for site %d" % SITE_ID, origSeq is not None, str(origSeq))
        sql(INV, "UPDATE INV_SITES SET IN_EIN_SEQ=%d WHERE IN_SITE_ID=%d" % (SEED_EIN_SEQ, SITE_ID))

        # --- 3. FEED-ROW PARITY: legacy REPORT_EDI856 @EIN=0 preview vs the rebuild feed ---------
        # The legacy preview branch selects ALL status-'C' ASNs; filter to our copied ASN by id via a
        # wrapper SELECT over the proc is not possible (it's a proc), so we run the legacy SELECT BODY
        # (the @EIN=0 branch, byte-identical to REPORT_EDI856) with an extra a.IN_ASN_ID predicate. This
        # is the legacy read, not the rebuild's — the comparison is rebuild-feed vs legacy-SELECT.
        print("\n--- (1) FEED-ROW PARITY: rebuild feed vs legacy REPORT_EDI856 @EIN=0 SELECT (same ASN) ---")
        legacy_feed = sql(INV,
            "SELECT d.VC_MANIFEST_NUMBER, d.VC_ASSY_PART_NUMBER, m.MO_PRICE, d.IN_QTY, "
            "a.VC_PRODUCTION_DATE, f.VC_ASSY_KANBAN_NUMBER "
            "FROM INV_ASN_MST a "
            "JOIN INV_ASN_DETAIL_MST d ON a.IN_ASN_ID=d.IN_ASN_ID "
            "JOIN INV_MANIFEST_COST_MST m ON d.VC_ASSY_PART_NUMBER=m.VC_ASSY_PART_NUMBER_CODE "
            "JOIN INV_FORECAST_DETAIL_INF f ON d.VC_ASSY_PART_NUMBER=f.VC_ASSY_PART_NUMBER_CODE "
            "WHERE a.VC_ASN_STATUS='C' AND a.IN_ASN_ID=%d "
            "AND m.VC_START_MANIFEST <= a.VC_PRODUCTION_DATE "
            "AND m.VC_END_MANIFEST   >= a.VC_PRODUCTION_DATE "
            "GROUP BY a.IN_ASN_EIN, d.VC_MANIFEST_NUMBER, d.VC_ASSY_PART_NUMBER, m.MO_PRICE, d.IN_QTY, "
            "a.VC_PRODUCTION_DATE, f.VC_ASSY_KANBAN_NUMBER, a.VC_START_SEQ_NUMBER, a.VC_LINE_NAME "
            "ORDER BY d.VC_MANIFEST_NUMBER, d.VC_ASSY_PART_NUMBER" % testAsn)
        # legacy emits the 9-col GROUP BY; the rebuild reads Manifest/Part/Qty/Kanban (+ prodDate from
        # the header). Compare on the 4 wire columns the builder consumes — the price isn't emitted, and
        # the CROSS APPLY TOP 1 must reproduce the SAME Kanban the legacy INNER join's GROUP BY produced
        # (for these parts there is exactly one forecast row, so TOP 1 == the unique row).
        legacy_wire = sorted((r[0], r[1], int(r[3]), r[5]) for r in legacy_feed)

        feedDs = edi856.system.db.runPrepQuery(edi856._FEED_SQL, [testAsn], "Inventory_Spike")
        rebuild_wire = sorted((r["Manifest"], r["PartNumber"], int(r["ShipQty"]), r["Kanban"])
                              for r in feedDs)
        rep.check("rebuild feed row COUNT == legacy SELECT row count",
                  len(rebuild_wire) == len(legacy_wire),
                  "rebuild=%d legacy=%d" % (len(rebuild_wire), len(legacy_wire)))
        rep.check("rebuild feed == legacy REPORT_EDI856 @EIN=0 SELECT, row-for-row "
                  "(Manifest/Part/Qty/Kanban)", rebuild_wire == legacy_wire,
                  "%d/%d rows identical" % (sum(1 for a, b in zip(rebuild_wire, legacy_wire) if a == b),
                                            len(legacy_wire)))
        if rebuild_wire != legacy_wire:
            for a, b in zip(rebuild_wire, legacy_wire):
                if a != b:
                    print("    [DIF] rebuild=%r legacy=%r" % (a, b))

        # --- 4. run the REAL send_856 (writes the file to tmpdir) -------------------------------
        print("\n--- (run) send_856 ---")
        result = edi856.send_856(testAsn, SITE_ID, database="Inventory_Spike", outDir=tmpdir,
                                 fileTime="1430")
        ein = result["ein"]
        rep.check("send_856 returned an EIN", ein is not None, "ein=%s" % ein)
        rep.check("send_856 row count == feed row count == legacy detail count",
                  result["rowCount"] == len(legacy_wire) == int(lcount),
                  "send=%d feed=%d legacy=%d" % (result["rowCount"], len(legacy_wire), int(lcount)))

        # --- 5a. EIN provenance: from INV_SITES.IN_EIN_SEQ (NOT 6440, NOT at-create) ------------
        print("\n--- (3) EIN provenance (INV_SITES.IN_EIN_SEQ, allocated at send) ---")
        rep.check("allocated EIN == seeded IN_EIN_SEQ + 1 (atomic at-send allocation)",
                  int(ein) == SEED_EIN_SEQ + 1, "ein=%s seed+1=%d" % (ein, SEED_EIN_SEQ + 1))
        postSeq = int(scalar(INV, "SELECT IN_EIN_SEQ FROM INV_SITES WHERE IN_SITE_ID=%d" % SITE_ID))
        rep.check("INV_SITES.IN_EIN_SEQ bumped by exactly 1", postSeq == SEED_EIN_SEQ + 1,
                  "post=%d seed=%d" % (postSeq, SEED_EIN_SEQ))
        rep.check("allocated EIN is NOT the 6440 literal", int(ein) != 6440, str(ein))
        hdrEin = int(scalar(INV, "SELECT IN_ASN_EIN FROM INV_ASN_MST WHERE IN_ASN_ID=%d" % testAsn))
        rep.check("the EIN was STAMPED on the ASN header", hdrEin == int(ein),
                  "header=%d alloc=%d" % (hdrEin, ein))

        # --- 5b. read the emitted file off disk + validate STRUCTURE ----------------------------
        print("\n--- (2) byte-exact STRUCTURE + self-consistency of the emitted file ---")
        rep.check("file written to outDir", result["path"] and os.path.exists(result["path"]),
                  result["path"])
        with open(result["path"], "rb") as fh:
            raw = fh.read()
        text = raw.decode("ascii")
        # CRLF-only (decision D): every line ends CRLF; no bare '~' terminator anywhere.
        rep.check("file uses CRLF line endings (decision D)", b"\r\n" in raw and b"~" not in raw,
                  "CRLF present, no '~'")
        segs = [s for s in text.split("\r\n") if s != ""]
        # also confirm the returned segments match the file (driver join is faithful)
        rep.check("file segments == driver-returned segments", segs == result["segments"],
                  "%d file vs %d returned" % (len(segs), len(result["segments"])))

        einStr = "%09d" % int(ein)
        isa = segs[0].split("*"); gs = segs[1].split("*"); st = segs[2].split("*")
        bsn = segs[3].split("*"); se = segs[-3].split("*"); ge = segs[-2].split("*"); iea = segs[-1].split("*")
        rep.check("envelope order ISA..IEA", segs[0].startswith("ISA*") and segs[1].startswith("GS*")
                  and segs[2].startswith("ST*856*") and segs[-1].startswith("IEA*"),
                  "%s ... %s" % (segs[0][:6], segs[-1][:6]))
        rep.check("%09d EIN in all 7 positions (ISA13/GS06/ST02/SE02/GE02/IEA02 + BSN02 suffix)",
                  isa[13] == einStr and gs[6] == einStr and st[2] == einStr and se[2] == einStr
                  and ge[2] == einStr and iea[2] == einStr and bsn[2] == PDATE + einStr,
                  "all == %s" % einStr)
        rep.check("control numbers consistent: ISA13 == GS06 == ST02 == SE02 == GE02 == IEA02",
                  len({isa[13], gs[6], st[2], se[2], ge[2], iea[2]}) == 1, isa[13])

        # SE01 == actual ST..SE count; CTT01 == HL count (recomputed off the FILE, not the builder).
        st_idx = next(i for i, s in enumerate(segs) if s.startswith("ST*"))
        se_idx = next(i for i, s in enumerate(segs) if s.startswith("SE*"))
        rep.check("SE01 == actual ST..SE segment count (file)", int(se[1]) == (se_idx - st_idx) + 1,
                  "SE01=%s actual=%d" % (se[1], (se_idx - st_idx) + 1))
        ctt = segs[se_idx - 1].split("*")
        hlCount = sum(1 for s in segs if s.startswith("HL*"))
        rep.check("CTT01 == HL count (file)", int(ctt[1]) == hlCount,
                  "CTT01=%s HL=%d" % (ctt[1], hlCount))

        # one Order HL + PRF per distinct manifest; one Item HL + LIN + SN1 per feed row.
        manifests = sorted(set(r["Manifest"] for r in
                          edi856.system.db.runPrepQuery(edi856._FEED_SQL, [testAsn], "Inventory_Spike")))
        orderHls = [s for s in segs if s.startswith("HL*") and "*O*" in s]
        prfs = [s for s in segs if s.startswith("PRF*")]
        itemHls = [s for s in segs if s.startswith("HL*") and "*I*" in s]
        lins = [s for s in segs if s.startswith("LIN*")]
        sn1s = [s for s in segs if s.startswith("SN1*")]
        rep.check("one Order HL + one PRF per distinct manifest",
                  len(orderHls) == len(manifests) == len(prfs),
                  "O=%d PRF=%d manifests=%d" % (len(orderHls), len(prfs), len(manifests)))
        rep.check("one Item HL + LIN + SN1 per feed row",
                  len(itemHls) == len(lins) == len(sn1s) == result["rowCount"],
                  "I=%d LIN=%d SN1=%d rows=%d" % (len(itemHls), len(lins), len(sn1s), result["rowCount"]))
        rep.check("ISA09 == yymmdd '260618' (century dropped) while GS04 == yyyymmdd",
                  isa[9] == "260618" and gs[4] == PDATE, "ISA09=%s GS04=%s" % (isa[9], gs[4]))
        rep.check("filename == decision-E pattern '85660618.txt'",
                  result["filename"] == "85660618.txt", result["filename"])

        # --- 6. DECOUPLED PER-ASN FLIP ---------------------------------------------------------
        print("\n--- (4) the C->S flip is PER-ASN, decoupled (NOT the blanket UPDATE_ASNStatus) ---")
        tgtStatus = scalar(INV, "SELECT VC_ASN_STATUS FROM INV_ASN_MST WHERE IN_ASN_ID=%d" % testAsn)
        decStatus = scalar(INV, "SELECT VC_ASN_STATUS FROM INV_ASN_MST WHERE IN_ASN_ID=%d" % decoyAsn)
        rep.check("target ASN flipped C -> S", tgtStatus == "S", tgtStatus)
        rep.check("decoy 'C' ASN UNTOUCHED (still 'C') — flip is per-ASN, NOT WHERE status='C'",
                  decStatus == "C", decStatus)

    finally:
        cleanup(testAsn, decoyAsn, origSeq)
        shutil.rmtree(tmpdir, ignore_errors=True)
        # confirm restoration
        left = scalar(INV, "SELECT COUNT(*) FROM INV_ASN_MST WHERE VC_LINE_NAME IN ('%s','%s')"
                           % (TEST_LINE, DECOY_LINE))
        rep.check("Inventory restored as-found (test + decoy ASNs swept)", int(left) == 0,
                  "%s remaining" % left)
        if origSeq is not None:
            nowSeq = scalar(INV, "SELECT IN_EIN_SEQ FROM INV_SITES WHERE IN_SITE_ID=%d" % SITE_ID)
            rep.check("INV_SITES.IN_EIN_SEQ restored to original", str(nowSeq) == str(origSeq),
                      "now=%s orig=%s" % (nowSeq, origSeq))

    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
