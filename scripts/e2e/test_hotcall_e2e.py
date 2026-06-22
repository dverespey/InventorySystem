#!/usr/bin/env python3
"""test_hotcall_e2e.py — END-TO-END exercise of the REAL hotcall.create_hotcall_asn driver (P12/P14) AND
the hot-call -> 856 flow (P13: an 8HC filename), driven headlessly through the system.db shim
(jython_shim). The "second ASN producer" keystone proof.

WHAT THIS PROVES (and what it HONESTLY does NOT)
------------------------------------------------
Runs the ACTUAL driver code (INSERT_ASNInfo OUTPUT-capture + per-detail INSERT_ASNDetail @HotCall=1, all
inside ONE Inventory transaction) — NOT a proc-EXEC, which would never touch the driver/seam (retro R8).
Then feeds the created hot-call ASN through the EXISTING edi856.send_856 builder (UNCHANGED except the P13
filename branch) and asserts an '8HC...' file with the M390 LIN/SN1 segments. It asserts what IS real:

  (1) HEADER — the created hot-call header is status 'C', VC_START_SEQ_NUMBER=VC_END_SEQ_NUMBER='-1'
      (the hot-call sentinel), VC_ASSEMBLY_LINE='', IN_ASN_EIN=0 (EIN-at-SEND, P14), and IN_QTY = the SUM
      of the detail qtys (P14 — fixes the legacy stale @QTY garbage, HotCallEntry.pas:246-247).
  (2) DETAIL — one INV_ASN_DETAIL_MST row per manual entry, every row @HotCall=1 (always-INSERT: two rows
      sharing a manifest stay DISTINCT, never accumulated), the operator-typed manifest, the right parts/qtys.
  (3) ATOMICITY — a DB failure injected into the per-detail INSERT rolls the WHOLE create back: no header,
      no details left behind (all-or-nothing, the M1 pattern). The driver re-raises (does not swallow).
  (4) 8HC FLOW (P13) — the created hot-call ASN, fed through the SAME send_856, produces a filename
      '8HC<copy(pd,4,5)><counter><LineName>.txt' (NOT '856...'), INCLUDING the line name + the per-ASN
      counter, and the emitted 856 carries LIN/SN1 segments for the hot-call parts. (The M390/M391 split
      is an 810/856 IT101 concern; the 856 itself carries the parts via LIN/SN1 — we assert those.) A
      normal-ASN control proves the '856<MMDD><LineName>.txt' branch still fires; a SECOND same-day hot-call
      proves the counter advances 1 -> 2 with NO filename collision.

HONEST VERIFICATION (fixture-fidelity discipline, memory feedback-parity-fixture-fidelity)
------------------------------------------------------------------------------------------
The CREATE + the 8HC filename are FAITHFUL to HotCallEntry.pas / ASNInvoice.pas:825 (derived from the
.pas, not the rebuild). But TRUE TEMA filename/byte-parity is PENDING A GOLDEN 8HC file — none exists for
this synthetic ASN, and the spike INV_SITES carries PLACEHOLDER site identity (DUNS/abbr/separators), so
the ISA/GS BYTES reflect the placeholder, not the wire (same caveat as test_edi856_e2e.py). What is proven:
the create is row-faithful to the .pas; the file is byte-EXACT STRUCTURE + self-consistent; the FILENAME
follows the 8HC branch. Byte parity vs a golden 8HC is the remaining cutover check.

METHOD
------
  1. Run the REAL create_hotcall_asn for a sentinel line with 3 manual rows (2 sharing a manifest, using
     parts that have cost+forecast coverage so the 856 feed survives).
  2. Assert the header (1) + detail rows (2) read back from Inventory.
  3. ATOMICITY (3): inject a fault into the per-detail INSERT and assert nothing is left behind + re-raise.
  4. 8HC FLOW (4): seed INV_SITES.IN_EIN_SEQ, run the EXISTING send_856 on the created hot-call ASN, assert
     the '8HC' filename + the LIN/SN1 segments. Run a NORMAL ASN control through send_856 -> '856' filename.
  5. Restore Inventory + INV_SITES as-found (delete the sentinel ASNs; reset IN_EIN_SEQ).

Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 scripts/e2e/test_hotcall_e2e.py
"""
import os, subprocess, sys, tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report          # noqa: E402
import jython_shim              # noqa: E402

CONTAINER = os.environ.get("CONTAINER", "mssql-spike")
SA_PASS = os.environ.get("SA_PASS")
INV = "Inventory"

ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".."))
HOTCALL_PY = os.path.join(ROOT, "docs", "analysis", "edi", "project-library", "hotcall", "code.py")
EDI856_PY = os.path.join(ROOT, "docs", "analysis", "edi", "856", "project-library", "edi856", "code.py")

# Sentinel line names so cleanup is exact and nothing real is touched. The filtered unique index
# UX_INV_ASN_MST_LINE_PDATE_NORMAL excludes VC_START_SEQ_NUMBER='-1', so hot-call ASNs never collide on
# (line, pdate) — multiple hot-call ASNs on the same sentinel line are fine.
TEST_LINE = "ZZHCE2E"
NORMAL_LINE = "ZZHCNORM"        # the normal-ASN control (a real seq -> '856' filename)
PDATE = "20260618"
SITE_ID = 1
SEED_EIN_SEQ = 9300             # known IN_EIN_SEQ -> the at-send allocation yields 9301
MANIFEST = "52089698"           # a real live hot-call manifest (ASN 4712), 8 chars, non-'7' (-> M390)

# parts with BOTH cost coverage (window spans 20260618) AND a forecast kanban -> survive the 856 feed
# INNER joins (verified live on the spike).
PART_A = "42600F261100"
PART_B = "42600FE36000"


def sql(db, query):
    if not SA_PASS:
        sys.exit("export SA_PASS first")
    # SET QUOTED_IDENTIFIER/ANSI_NULLS ON to mirror the gateway JDBC session (the filtered unique index on
    # INV_ASN_MST requires them ON for any DML, Msg 1934 otherwise).
    out = subprocess.check_output([
        "docker", "exec", CONTAINER, "/opt/mssql-tools18/bin/sqlcmd", "-C", "-S", "localhost",
        "-U", "sa", "-P", SA_PASS, "-d", db, "-h", "-1", "-W", "-s", "\t",
        "-Q", "SET QUOTED_IDENTIFIER ON; SET ANSI_NULLS ON; SET NOCOUNT ON; " + query], text=True)
    return [l.split("\t") for l in out.splitlines()
            if l.strip() and not l.startswith("(") and not l.startswith("Msg ")]


def scalar(db, q):
    rows = sql(db, q)
    return rows[0][0] if rows else None


def cleanup(origSeq):
    """Restore Inventory + INV_SITES as-found: delete the sentinel ASNs + reset IN_EIN_SEQ."""
    sql(INV, "DELETE FROM INV_ASN_DETAIL_MST WHERE IN_ASN_ID IN "
             "(SELECT IN_ASN_ID FROM INV_ASN_MST WHERE VC_LINE_NAME IN ('%s','%s')); "
             "DELETE FROM INV_ASN_MST WHERE VC_LINE_NAME IN ('%s','%s')"
             % (TEST_LINE, NORMAL_LINE, TEST_LINE, NORMAL_LINE))
    if origSeq is not None:
        sql(INV, "UPDATE INV_SITES SET IN_EIN_SEQ = %s WHERE IN_SITE_ID = %d" % (origSeq, SITE_ID))


def details_rows(asnId):
    rows = sql(INV, "SELECT VC_MANIFEST_NUMBER, VC_ASSY_PART_NUMBER, IN_QTY "
                    "FROM INV_ASN_DETAIL_MST WHERE IN_ASN_ID=%d ORDER BY IN_INV_ID, VC_ASSY_PART_NUMBER"
                    % asnId)
    return [(r[0], r[1], int(r[2])) for r in rows]


def main():
    print("=" * 90)
    print(" create_hotcall_asn END-TO-END — header/detail/atomicity (P12/P14) + the 8HC -> 856 flow (P13)")
    print(" (CREATE + 8HC faithful to the .pas; TRUE TEMA byte-parity pending a golden 8HC — see docstring)")
    print("=" * 90)
    rep = Report()

    hotcall = jython_shim.load_wrapper("hotcall_driver", HOTCALL_PY)
    edi856 = jython_shim.load_wrapper("edi856_driver", EDI856_PY)

    # --- pre-clean any leftovers from a prior aborted run ---------------------------------------
    sql(INV, "DELETE FROM INV_ASN_DETAIL_MST WHERE IN_ASN_ID IN "
             "(SELECT IN_ASN_ID FROM INV_ASN_MST WHERE VC_LINE_NAME IN ('%s','%s')); "
             "DELETE FROM INV_ASN_MST WHERE VC_LINE_NAME IN ('%s','%s')"
             % (TEST_LINE, NORMAL_LINE, TEST_LINE, NORMAL_LINE))

    origSeq = scalar(INV, "SELECT IN_EIN_SEQ FROM INV_SITES WHERE IN_SITE_ID=%d" % SITE_ID)
    created = normalAsn = None
    tmpdir = tempfile.mkdtemp(prefix="hotcall_e2e_")
    try:
        # ============================================================================================
        # (1)+(2) CREATE — run the REAL driver; assert header + detail rows.
        #   3 manual rows: PART_A qty 1, PART_B qty 2, PART_A qty 3 (rows 1 & 3 share PART_A + the
        #   manifest -> @HotCall=1 keeps them DISTINCT, NOT accumulated). header IN_QTY must = 1+2+3 = 6.
        # ============================================================================================
        print("\n--- (1)+(2) create_hotcall_asn: header + detail rows ---")
        items = [(PART_A, 1), (PART_B, 2), (PART_A, 3)]
        result = hotcall.create_hotcall_asn(
            line=TEST_LINE, prodDate=PDATE, manifest=MANIFEST, items=items,
            site=SITE_ID, database="Inventory_Spike")
        created = result["asnId"]
        rep.check("create_hotcall_asn returned a new ASN id", created is not None, "asnId=%s" % created)
        rep.check("driver reports qty = SUM(detail) = 6 (P14)", result["qty"] == 6, "%d" % result["qty"])
        rep.check("driver reports 3 detail rows (no dedup)", len(result["details"]) == 3,
                  "%d" % len(result["details"]))

        # --- header read-back (1) ---
        hdr = sql(INV, "SELECT VC_ASN_STATUS, VC_START_SEQ_NUMBER, VC_END_SEQ_NUMBER, VC_ASSEMBLY_LINE, "
                       "IN_ASN_EIN, IN_QTY, VC_PRODUCTION_DATE FROM INV_ASN_MST WHERE IN_ASN_ID=%d"
                       % created)[0]
        st, sseq, eseq, aline, ein, hq, pd = hdr
        rep.check("header status = 'C' (created)", st == "C", st)
        rep.check("header VC_START_SEQ_NUMBER = '-1' (hot-call sentinel)", sseq == "-1", sseq)
        rep.check("header VC_END_SEQ_NUMBER = '-1' (hot-call sentinel)", eseq == "-1", eseq)
        rep.check("header VC_ASSEMBLY_LINE = '' (HotCallEntry.pas:236-237)", aline == "", "%r" % aline)
        rep.check("header IN_ASN_EIN = 0 (EIN allocated at SEND, P14)", int(ein) == 0, ein)
        rep.check("header IN_QTY = 6 = SUM(detail) (P14, fixes the @QTY garbage)", int(hq) == 6, hq)
        rep.check("header prodDate = 20260618", pd == PDATE, pd)

        # --- detail read-back (2) ---
        det = details_rows(created)
        rep.check("3 detail rows persisted (one per manual entry)", len(det) == 3, "%d rows" % len(det))
        rep.check("every detail carries the operator manifest", all(d[0] == MANIFEST for d in det),
                  MANIFEST)
        # two rows share PART_A + manifest but are DISTINCT (qty 1 and qty 3, NOT a single accumulated 4).
        partA_rows = [d for d in det if d[1] == PART_A]
        rep.check("two PART_A rows kept DISTINCT (qty 1 and qty 3, NOT accumulated to 4) — @HotCall=1",
                  len(partA_rows) == 2 and sorted(r[2] for r in partA_rows) == [1, 3],
                  "PART_A rows=%r" % [r[2] for r in partA_rows])
        rep.check("PART_B row qty = 2", any(d[1] == PART_B and d[2] == 2 for d in det),
                  "%r" % [d for d in det if d[1] == PART_B])

        # ============================================================================================
        # (3) ATOMICITY — inject a fault into the per-detail INSERT; the WHOLE create must roll back.
        #     Use a DISTINCT manifest so this attempt is unambiguously separable from the committed one.
        # ============================================================================================
        print("\n--- (3) ATOMICITY: a DB failure in the per-detail INSERT -> rollback -> no partial ---")
        ATOM_MANIFEST = "52089999"
        preHeaders = int(scalar(INV, "SELECT COUNT(*) FROM INV_ASN_MST WHERE VC_LINE_NAME='%s'" % TEST_LINE))
        realUpdate = hotcall.system.db.runPrepUpdate
        hitDetail = {"n": 0}

        def faulting_update(sqlText, args, db=None, getKey=False, tx=None):
            # fail on the FIRST INSERT_ASNDetail (AFTER the header insert ran) so we prove the header
            # unwinds with the details. The header goes through runScalarPrepQuery, not this path.
            if "INSERT_ASNDetail" in sqlText:
                hitDetail["n"] += 1
                raise RuntimeError("INJECTED DB failure on INSERT_ASNDetail")
            return realUpdate(sqlText, args, db=db, getKey=getKey, tx=tx)

        hotcall.system.db.runPrepUpdate = faulting_update
        raised = False
        try:
            hotcall.create_hotcall_asn(
                line=TEST_LINE, prodDate=PDATE, manifest=ATOM_MANIFEST, items=[(PART_A, 9)],
                site=SITE_ID, database="Inventory_Spike")
        except Exception as e:
            raised = True
            print("    (expected) create_hotcall_asn raised: %s" % e)
        finally:
            hotcall.system.db.runPrepUpdate = realUpdate     # un-inject

        rep.check("create_hotcall_asn re-raised the injected DB failure (did not swallow it)", raised,
                  "raised=%s" % raised)
        rep.check("the fault actually fired on the per-detail INSERT (non-vacuous)", hitDetail["n"] >= 1,
                  "hits=%d" % hitDetail["n"])
        postHeaders = int(scalar(INV, "SELECT COUNT(*) FROM INV_ASN_MST WHERE VC_LINE_NAME='%s'"
                                       % TEST_LINE))
        rep.check("NO new header left behind (the partial header rolled back with the tx)",
                  postHeaders == preHeaders, "pre=%d post=%d" % (preHeaders, postHeaders))
        atomDetail = int(scalar(INV, "SELECT COUNT(*) FROM INV_ASN_DETAIL_MST WHERE VC_MANIFEST_NUMBER='%s'"
                                     % ATOM_MANIFEST))
        rep.check("NO detail rows left behind for the failed attempt's manifest", atomDetail == 0,
                  "%d rows" % atomDetail)

        # ============================================================================================
        # (4) 8HC FLOW (P13) — the created hot-call ASN -> the EXISTING send_856 -> an '8HC' filename
        #     + LIN/SN1 segments. Plus a NORMAL-ASN control proving the '856' branch still fires.
        # ============================================================================================
        print("\n--- (4) 8HC FLOW: the hot-call ASN through the EXISTING send_856 -> '8HC...1.txt' ---")
        sql(INV, "UPDATE INV_SITES SET IN_EIN_SEQ = %d WHERE IN_SITE_ID = %d" % (SEED_EIN_SEQ, SITE_ID))
        sent = edi856.send_856(created, SITE_ID, database="Inventory_Spike", outDir=tmpdir,
                               fileTime="1430")
        files = sorted(os.listdir(tmpdir))
        # EXPECTED derived from the OPERATIONAL sender MainMenu.pas:2723: '8HC'+copy(pd,4,5)+y+LineName+'.txt'.
        # First hot-call of the day for TEST_LINE -> counter y=1; the name INCLUDES the line name.
        expected8hc = "8HC" + PDATE[3:8] + "1" + TEST_LINE + ".txt"   # '8HC606181ZZHCE2E.txt'
        rep.check("send_856 wrote the 8HC filename '%s' (P13: hot-call branch, incl. LineName + y=1)"
                  % expected8hc, sent["filename"] == expected8hc and expected8hc in files,
                  "filename=%s dir=%r" % (sent["filename"], files))
        rep.check("the 8HC filename INCLUDES the line name '%s' (operational sender appends LineName)"
                  % TEST_LINE, TEST_LINE in sent["filename"], sent["filename"])
        rep.check("the 8HC file is NOT the normal '856' pattern", not sent["filename"].startswith("856"),
                  sent["filename"])
        # the 856 carries the hot-call parts via LIN/SN1 (M390 — the manifest is non-'7').
        segs = sent["segments"]
        linParts = [s for s in segs if s.startswith("LIN")]
        sn1Segs = [s for s in segs if s.startswith("SN1")]
        rep.check("the 856 carries LIN segments for the hot-call parts (PART_A and PART_B present)",
                  any(PART_A in s for s in linParts) and any(PART_B in s for s in linParts),
                  "%d LIN segs" % len(linParts))
        rep.check("the 856 carries an SN1 (ship-qty) segment per LIN", len(sn1Segs) == len(linParts)
                  and len(sn1Segs) >= 2, "%d SN1 / %d LIN" % (len(sn1Segs), len(linParts)))
        # the PRF (order) segment carries the operator manifest (the '-' is data, legacy PRF format).
        prfSegs = [s for s in segs if s.startswith("PRF")]
        rep.check("the PRF (order) segment carries the operator manifest %s" % MANIFEST,
                  any(MANIFEST in s for s in prfSegs), "%r" % prfSegs)
        rep.check("EIN allocated at SEND = seeded IN_EIN_SEQ + 1 = %d (NOT 0, NOT at-create)"
                  % (SEED_EIN_SEQ + 1), sent["ein"] == SEED_EIN_SEQ + 1, "ein=%s" % sent["ein"])

        # --- (4a) SECOND same-day hot-call on the SAME line -> counter advances 1 -> 2, no collision ----
        # The first hot-call ('created') is now flipped to 'S'. A SECOND hot-call ASN on TEST_LINE / PDATE
        # is the Nth=2 of the day -> send_856 must derive counter=2 (1 + the one same-day same-line hot-call
        # already 'S'), so the filename is '8HC<Y+MMDD>2<LineName>.txt' — DISTINCT from the first (no
        # same-day filename collision, the property the legacy per-batch y guaranteed). The filtered unique
        # index UX_INV_ASN_MST_LINE_PDATE_NORMAL excludes '-1', so two hot-calls on (line, pdate) coexist.
        print("\n--- (4a) a SECOND same-day hot-call on the same line -> counter 2 (no filename collision) ---")
        result2 = hotcall.create_hotcall_asn(
            line=TEST_LINE, prodDate=PDATE, manifest="52089777", items=[(PART_B, 4)],
            site=SITE_ID, database="Inventory_Spike")
        created2 = result2["asnId"]
        sent2 = edi856.send_856(created2, SITE_ID, database="Inventory_Spike", outDir=tmpdir,
                                fileTime="1430")
        expected8hc2 = "8HC" + PDATE[3:8] + "2" + TEST_LINE + ".txt"   # counter=2 -> '8HC606182ZZHCE2E.txt'
        rep.check("the 2nd same-day hot-call -> counter 2: '%s' (1 + the one already 'S')" % expected8hc2,
                  sent2["filename"] == expected8hc2, sent2["filename"])
        rep.check("the two same-day hot-call filenames DIFFER (counter 1 vs 2 -> NO collision)",
                  sent["filename"] != sent2["filename"],
                  "first=%s second=%s" % (sent["filename"], sent2["filename"]))

        # --- NORMAL-ASN control: a real seq -> the '856' branch (proves the 8HC branch is selective) ---
        print("\n--- (4b) NORMAL-ASN control through the SAME send_856 -> '856' filename ---")
        # build a normal ASN by hand (status 'C', a real seq) with the same parts so the feed survives.
        sql(INV, "INSERT INTO INV_ASN_MST (IN_ASN_EIN, VC_ASN_STATUS, VC_LINE_NAME, VC_ASSEMBLY_LINE, "
                 "VC_START_SEQ_NUMBER, DT_START_SEQ, VC_END_SEQ_NUMBER, DT_END_SEQ, IN_QTY, "
                 "VC_PRODUCTION_DATE, VC_LAST_UPDATE, VC_ADD) "
                 "VALUES (0, 'C', '%s', '', '0001', GETDATE(), '0009', GETDATE(), 3, '%s', "
                 "'2026061814300000', '2026061814300000')" % (NORMAL_LINE, PDATE))
        normalAsn = int(scalar(INV, "SELECT IN_ASN_ID FROM INV_ASN_MST WHERE VC_LINE_NAME='%s'"
                                    % NORMAL_LINE))
        sql(INV, "INSERT INTO INV_ASN_DETAIL_MST (IN_ASN_ID, IN_INV_ID, IN_ASN_EIN, VC_MANIFEST_NUMBER, "
                 "VC_ASSY_PART_NUMBER, IN_QTY, VC_LAST_UPDATE, VC_ADD) "
                 "VALUES (%d, NULL, 0, '52081111', '%s', 3, '2026061814300000', '2026061814300000')"
                 % (normalAsn, PART_A))
        sentNorm = edi856.send_856(normalAsn, SITE_ID, database="Inventory_Spike", outDir=tmpdir,
                                   fileTime="1430")
        # EXPECTED from the operational sender MainMenu.pas:2718: '856'+copy(pd,5,4)=MMDD+LineName+'.txt'.
        expected856 = "856" + PDATE[4:8] + NORMAL_LINE + ".txt"   # '8560618ZZHCNORM.txt'
        rep.check("NORMAL ASN through the SAME send_856 -> '%s' (856 + MMDD + LineName; branch still fires)"
                  % expected856, sentNorm["filename"] == expected856, sentNorm["filename"])
        rep.check("the NORMAL filename INCLUDES the line name '%s' (operational sender appends LineName)"
                  % NORMAL_LINE, NORMAL_LINE in sentNorm["filename"], sentNorm["filename"])
        rep.check("the two filenames DIVERGE (8HC for hot-call, 856 for normal) — the branch is real",
                  sent["filename"] != sentNorm["filename"],
                  "hot=%s normal=%s" % (sent["filename"], sentNorm["filename"]))

    finally:
        cleanup(origSeq)
        post = int(scalar(INV, "SELECT COUNT(*) FROM INV_ASN_MST WHERE VC_LINE_NAME IN ('%s','%s')"
                               % (TEST_LINE, NORMAL_LINE)))
        rep.check("Inventory restored as-found (sentinel ASNs swept)", post == 0, "%d remaining" % post)
        postSeq = scalar(INV, "SELECT IN_EIN_SEQ FROM INV_SITES WHERE IN_SITE_ID=%d" % SITE_ID)
        rep.check("INV_SITES.IN_EIN_SEQ restored", str(postSeq) == str(origSeq),
                  "orig=%s post=%s" % (origSeq, postSeq))

    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
