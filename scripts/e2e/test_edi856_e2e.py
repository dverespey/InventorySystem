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
# the latent-divergence probe ASN gets its OWN line name: INV_ASN_MST has a filtered UNIQUE index
# UX_INV_ASN_MST_LINE_PDATE_NORMAL on (VC_LINE_NAME, VC_PRODUCTION_DATE) WHERE VC_START_SEQ_NUMBER<>'-1'
# (spike-asn-unique-guard.sql), so the probe MUST differ from testAsn on line-or-date or the INSERT is
# rejected as a duplicate. A distinct line keeps cleanup exact.
PROBE_LINE = "ZZ856PRB"
# sentinel VC_ADD stamps tagging the cost/forecast rows the latent-divergence cases inject, so cleanup is
# exact (delete WHERE VC_ADD = sentinel) and the spike is left as-found even though these are real rows.
COST_SENTINEL = "ZZ856COST0000000"      # 16-char VC_ADD stamp on injected INV_MANIFEST_COST_MST rows
FC_SENTINEL = "ZZ856FCST0000000"        # 16-char VC_ADD stamp on injected INV_FORECAST_DETAIL_INF rows

EDI856_PY = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                             "docs", "analysis", "edi", "856", "project-library", "edi856", "code.py"))
FEED_SQL_FILE = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                                 "docs", "analysis", "edi", "856", "spike-edi856-feed.sql"))


def _marked_feed_body(path):
    """Extract the canonical SELECT body between the BEGIN/END FEED SQL markers in spike-edi856-feed.sql,
    drop line-comments + collapse all whitespace runs to single spaces, strip the trailing ';'. Used to
    prove code.py:_FEED_SQL is byte-identical to the .sql (modulo the @ASNID<->? bind)."""
    with open(path) as fh:
        lines = fh.read().splitlines()
    out, grab = [], False
    for ln in lines:
        if ">>> BEGIN FEED SQL" in ln:
            grab = True
            continue
        if "<<< END FEED SQL" in ln:
            break
        if grab:
            out.append(ln)
    body = " ".join(out)
    body = " ".join(body.split())           # collapse whitespace
    return body.rstrip(";").strip()


def _normalize_sql(s):
    """Collapse whitespace runs to single spaces for a whitespace-insensitive char compare."""
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


def cleanup(testAsn, decoyAsn, origSeq):
    """Restore Inventory + INV_SITES as-found."""
    sql(INV, "DELETE FROM INV_ASN_DETAIL_MST WHERE IN_ASN_ID IN "
             "(SELECT IN_ASN_ID FROM INV_ASN_MST WHERE VC_LINE_NAME IN ('%s','%s','%s')); "
             "DELETE FROM INV_ASN_MST WHERE VC_LINE_NAME IN ('%s','%s','%s')"
             % (TEST_LINE, DECOY_LINE, PROBE_LINE, TEST_LINE, DECOY_LINE, PROBE_LINE))
    # sweep any cost/forecast rows the latent-divergence cases injected (tagged by sentinel VC_ADD).
    sql(INV, "DELETE FROM INV_MANIFEST_COST_MST WHERE VC_ADD = '%s'" % COST_SENTINEL)
    sql(INV, "DELETE FROM INV_FORECAST_DETAIL_INF WHERE VC_ADD = '%s'" % FC_SENTINEL)
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
             "(SELECT IN_ASN_ID FROM INV_ASN_MST WHERE VC_LINE_NAME IN ('%s','%s','%s')); "
             "DELETE FROM INV_ASN_MST WHERE VC_LINE_NAME IN ('%s','%s','%s')"
             % (TEST_LINE, DECOY_LINE, PROBE_LINE, TEST_LINE, DECOY_LINE, PROBE_LINE))
    sql(INV, "DELETE FROM INV_MANIFEST_COST_MST WHERE VC_ADD = '%s'" % COST_SENTINEL)
    sql(INV, "DELETE FROM INV_FORECAST_DETAIL_INF WHERE VC_ADD = '%s'" % FC_SENTINEL)

    # --- (#8) DRIFT GUARD: code.py:_FEED_SQL is byte-identical to the .sql's marked SELECT body --
    # (modulo the single @ASNID<->? bind). If either is edited without the other, this fails.
    print("\n--- (#8) _FEED_SQL == spike-edi856-feed.sql marked body (modulo @ASNID<->?) ---")
    sql_body = _marked_feed_body(FEED_SQL_FILE).replace("@ASNID", "?")
    inline_body = _normalize_sql(edi856._FEED_SQL)
    rep.check("_FEED_SQL char-identical to spike-edi856-feed.sql SELECT body (@ASNID -> ?)",
              inline_body == sql_body,
              "match" if inline_body == sql_body else
              "\n      inline=%r\n      .sql  =%r" % (inline_body, sql_body))
    rep.check("the feed now uses an INNER forecast JOIN (not CROSS APPLY TOP 1)",
              "JOIN INV_FORECAST_DETAIL_INF" in inline_body and "CROSS APPLY" not in inline_body,
              "INNER join present, CROSS APPLY absent")
    rep.check("the feed GROUP BY keeps m.MO_PRICE (legacy splits rows on it)",
              "m.MO_PRICE" in inline_body, "MO_PRICE in GROUP BY")

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
        # the rebuild's INNER forecast JOIN + 9-key GROUP BY (incl. MO_PRICE) must reproduce the legacy
        # rows row-for-row. On THIS ASN every part has 1 cost row + 1 forecast row, so the count is 16==16;
        # the >1-row fan-out/split is exercised separately by cases #5 (multi-price) and #6 (multi-kanban).
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
        # (#7) ISA15 = the X12 usage indicator, EXACTLY ONE char. INV_SITES.VC_EDI_MODE is now seeded 'P'
        # (was 'PROD', which emitted a malformed 4-char ISA15). The builder also enforces single-char.
        rep.check("ISA15 (usage indicator) is exactly one char (was 'PROD' -> now 'P')",
                  len(isa[15]) == 1 and isa[15] in ("P", "T"),
                  "ISA15=%r len=%d" % (isa[15], len(isa[15])))
        rep.check("ISA16 (component-element separator) is exactly one char",
                  len(isa[16]) == 1, "ISA16=%r" % isa[16])
        # the TD1/LIN byte fixes appear in the REAL emitted file too (not just the unit test).
        rep.check("TD1 segment == 'TD1**' (two empty elements, .pas:294-296)",
                  next(s for s in segs if s.startswith("TD1*")) == "TD1**",
                  next(s for s in segs if s.startswith("TD1*")))
        rep.check("every LIN ends with the trailing element separator (.pas:352)",
                  all(s.endswith("*") for s in lins), "%d LINs all trailing-sep" % len(lins))
        rep.check("filename == decision-E pattern '85660618.txt'",
                  result["filename"] == "85660618.txt", result["filename"])

        # --- 6. DECOUPLED PER-ASN FLIP ---------------------------------------------------------
        print("\n--- (4) the C->S flip is PER-ASN, decoupled (NOT the blanket UPDATE_ASNStatus) ---")
        tgtStatus = scalar(INV, "SELECT VC_ASN_STATUS FROM INV_ASN_MST WHERE IN_ASN_ID=%d" % testAsn)
        decStatus = scalar(INV, "SELECT VC_ASN_STATUS FROM INV_ASN_MST WHERE IN_ASN_ID=%d" % decoyAsn)
        rep.check("target ASN flipped C -> S", tgtStatus == "S", tgtStatus)
        rep.check("decoy 'C' ASN UNTOUCHED (still 'C') — flip is per-ASN, NOT WHERE status='C'",
                  decStatus == "C", decStatus)

        # ========================================================================================
        # LATENT FEED-DIVERGENCE CASES — these prove the feed reproduces the LEGACY fan-out/split
        # that the prior build collapsed, by INJECTING the >1-row shape the sparse snapshot lacks.
        # Each compares the rebuild feed against the legacy 9-col GROUP-BY SELECT (the byte-faithful
        # REPORT_EDI856 @EIN=0 read, scoped by ASN id), row-for-row. Injected cost/forecast rows are
        # sentinel-tagged (VC_ADD) and swept in cleanup() so the spike is left as-found.
        #
        # The probe ASN ships a SINGLE detail line for part PROBE_PART (which has exactly ONE cost row
        # and ONE forecast row in the snapshot), so a 2nd cost window / 2nd forecast row makes the row
        # count jump 1 -> 2 unambiguously.
        PROBE_PART = "42600FEK5000"     # ASN-4721 part; 1 cost row + 1 forecast row in the snapshot
        PROBE_KANBAN = "JZUX"           # its existing forecast kanban (the snapshot's single row)
        PROBE_MANIFEST = "76061857"

        # a single-line probe ASN (status 'C', EIN 0, prodDate within the cost window). Its OWN line name
        # (PROBE_LINE) so it does not collide with testAsn on the (line, prodDate) filtered unique index.
        sql(INV,
            "INSERT INTO INV_ASN_MST (IN_ASN_EIN, VC_ASN_STATUS, VC_LINE_NAME, VC_ASSEMBLY_LINE, "
            "VC_START_SEQ_NUMBER, VC_END_SEQ_NUMBER, IN_QTY, VC_PRODUCTION_DATE, VC_LAST_UPDATE, VC_ADD) "
            "VALUES (0, 'C', '%s', '', '0001', '0002', 0, '%s', '0000000000000000', '0000000000000000')"
            % (PROBE_LINE, PDATE))
        probeAsn = int(scalar(INV, "SELECT MAX(IN_ASN_ID) FROM INV_ASN_MST WHERE VC_LINE_NAME='%s'"
                                   % PROBE_LINE))
        rep.check("probe ASN created under its own line (distinct from testAsn on the unique index)",
                  probeAsn != testAsn, "probe=%d testAsn=%d" % (probeAsn, testAsn))
        sql(INV,
            "INSERT INTO INV_ASN_DETAIL_MST (IN_ASN_ID, IN_INV_ID, IN_ASN_EIN, VC_MANIFEST_NUMBER, "
            "VC_ASSY_PART_NUMBER, IN_QTY, VC_LAST_UPDATE, VC_ADD) "
            "VALUES (%d, NULL, 0, '%s', '%s', 80, '0000000000000000', '0000000000000000')"
            % (probeAsn, PROBE_MANIFEST, PROBE_PART))

        # the legacy 9-col GROUP-BY SELECT (REPORT_EDI856 @EIN=0 body), scoped by ASN id (status drop —
        # we compare the feed READ, not the status branch). Returns the 4 wire cols sorted.
        def legacy_wire_for(asn):
            rows = sql(INV,
                "SELECT d.VC_MANIFEST_NUMBER, d.VC_ASSY_PART_NUMBER, m.MO_PRICE, d.IN_QTY, "
                "a.VC_PRODUCTION_DATE, f.VC_ASSY_KANBAN_NUMBER "
                "FROM INV_ASN_MST a "
                "JOIN INV_ASN_DETAIL_MST d ON a.IN_ASN_ID=d.IN_ASN_ID "
                "JOIN INV_MANIFEST_COST_MST m ON d.VC_ASSY_PART_NUMBER=m.VC_ASSY_PART_NUMBER_CODE "
                "JOIN INV_FORECAST_DETAIL_INF f ON d.VC_ASSY_PART_NUMBER=f.VC_ASSY_PART_NUMBER_CODE "
                "WHERE a.IN_ASN_ID=%d "
                "AND m.VC_START_MANIFEST <= a.VC_PRODUCTION_DATE "
                "AND m.VC_END_MANIFEST   >= a.VC_PRODUCTION_DATE "
                "GROUP BY a.IN_ASN_EIN, d.VC_MANIFEST_NUMBER, d.VC_ASSY_PART_NUMBER, m.MO_PRICE, "
                "d.IN_QTY, a.VC_PRODUCTION_DATE, f.VC_ASSY_KANBAN_NUMBER, a.VC_START_SEQ_NUMBER, "
                "a.VC_LINE_NAME "
                "ORDER BY d.VC_MANIFEST_NUMBER, d.VC_ASSY_PART_NUMBER" % asn)
            return sorted((r[0], r[1], int(r[3]), r[5]) for r in rows)

        def rebuild_wire_for(asn):
            ds = edi856.system.db.runPrepQuery(edi856._FEED_SQL, [asn], "Inventory_Spike")
            return sorted((r["Manifest"], r["PartNumber"], int(r["ShipQty"]), r["Kanban"]) for r in ds)

        # --- (#5) MULTI-PRICE: MO_PRICE in the GROUP BY splits a 2-cost-window part into 2 rows ----
        # The snapshot part has one cost window covering PDATE. Inject a SECOND, price-distinct window
        # that ALSO covers PDATE. Legacy keeps m.MO_PRICE in the GROUP BY -> 2 rows; the rebuild must
        # too (dropping MO_PRICE would collapse to 1, BLOCKER-3). The 856 never carries the price, but
        # MO_PRICE changes the ROW COUNT the wire DOES carry (2 LIN/SN1 lines vs 1).
        print("\n--- (#5) MULTI-PRICE: a 2nd price-distinct overlapping cost window -> 2 feed rows ---")
        base1 = rebuild_wire_for(probeAsn)
        rep.check("probe ASN feeds exactly 1 row BEFORE the 2nd cost window (single line)",
                  len(base1) == 1, "%d rows: %r" % (len(base1), base1))
        # clone the existing covering cost row's window, but with a DIFFERENT price + the sentinel VC_ADD.
        sql(INV,
            "INSERT INTO INV_MANIFEST_COST_MST (VC_ASSY_PART_NUMBER_CODE, VC_ASSY_MANIFEST_NUMBER, "
            "VC_START_MANIFEST, VC_END_MANIFEST, MO_PRICE, VC_LAST_UPDATE, VC_ADD) "
            "SELECT TOP 1 VC_ASSY_PART_NUMBER_CODE, VC_ASSY_MANIFEST_NUMBER, VC_START_MANIFEST, "
            "VC_END_MANIFEST, MO_PRICE + 11.11, VC_LAST_UPDATE, '%s' "
            "FROM INV_MANIFEST_COST_MST WHERE VC_ASSY_PART_NUMBER_CODE='%s' "
            "AND VC_START_MANIFEST <= '%s' AND VC_END_MANIFEST >= '%s'"
            % (COST_SENTINEL, PROBE_PART, PDATE, PDATE))
        legP = legacy_wire_for(probeAsn)
        rebP = rebuild_wire_for(probeAsn)
        rep.check("legacy SELECT now emits 2 rows for the 2-cost-window part (MO_PRICE split)",
                  len(legP) == 2, "%d rows: %r" % (len(legP), legP))
        rep.check("rebuild feed MATCHES legacy: 2 rows (KEEPS MO_PRICE in GROUP BY — BLOCKER-3 fixed)",
                  rebP == legP and len(rebP) == 2,
                  "rebuild=%r legacy=%r" % (rebP, legP))
        # sweep the injected cost row so case #6 starts clean.
        sql(INV, "DELETE FROM INV_MANIFEST_COST_MST WHERE VC_ADD='%s'" % COST_SENTINEL)
        afterP = rebuild_wire_for(probeAsn)
        rep.check("after sweeping the 2nd cost window the feed is back to 1 row", len(afterP) == 1,
                  "%d rows" % len(afterP))

        # --- (#6) MULTI-KANBAN: the INNER forecast JOIN fans out 2 distinct-kanban rows -> 2 lines -
        # The snapshot part has one forecast row (kanban PROBE_KANBAN). Inject a SECOND forecast row
        # for the SAME part with a DISTINCT kanban. The legacy INNER JOIN INV_FORECAST_DETAIL_INF fans
        # the detail line to 2 rows; the 9-col GROUP BY keeps each distinct kanban -> 2 LIN lines. The
        # prior CROSS APPLY (SELECT TOP 1 ...) kept ONE (nondeterministically) — SHOULD-FIX 4. The
        # rebuild's INNER join must now keep BOTH.
        print("\n--- (#6) MULTI-KANBAN: a 2nd distinct-kanban forecast row -> 2 feed rows (fan-out) ---")
        base6 = rebuild_wire_for(probeAsn)
        rep.check("probe ASN feeds exactly 1 row BEFORE the 2nd forecast row", len(base6) == 1,
                  "%d rows: %r" % (len(base6), base6))
        NEW_KANBAN = "ZZZ9"
        rep.check("the injected kanban %r differs from the snapshot kanban %r" % (NEW_KANBAN, PROBE_KANBAN),
                  NEW_KANBAN != PROBE_KANBAN, "%s != %s" % (NEW_KANBAN, PROBE_KANBAN))
        # clone the existing forecast row but with a DISTINCT kanban + the sentinel VC_ADD (all other
        # NOT-NULL columns copied verbatim so the row is schema-valid).
        sql(INV,
            "INSERT INTO INV_FORECAST_DETAIL_INF (VC_ASSY_PART_NUMBER_CODE, VC_EFFECTIVE_MONTH, "
            "VC_ASSY_KANBAN_NUMBER, VC_TIRE_PART_NUMBER_CODE, IN_TIRE_RATIO, VC_WHEEL_PART_NUMBER_CODE, "
            "IN_WHEEL_RATIO, IN_RATIO, IN_ASSY_QTY, VC_BROADCAST_CODE, VC_LAST_UPDATE, VC_ADD) "
            "SELECT TOP 1 VC_ASSY_PART_NUMBER_CODE, VC_EFFECTIVE_MONTH, '%s', VC_TIRE_PART_NUMBER_CODE, "
            "IN_TIRE_RATIO, VC_WHEEL_PART_NUMBER_CODE, IN_WHEEL_RATIO, IN_RATIO, IN_ASSY_QTY, "
            "VC_BROADCAST_CODE, VC_LAST_UPDATE, '%s' "
            "FROM INV_FORECAST_DETAIL_INF WHERE VC_ASSY_PART_NUMBER_CODE='%s'"
            % (NEW_KANBAN, FC_SENTINEL, PROBE_PART))
        legK = legacy_wire_for(probeAsn)
        rebK = rebuild_wire_for(probeAsn)
        rep.check("legacy SELECT now emits 2 rows for the 2-forecast-row part (distinct kanbans)",
                  len(legK) == 2, "%d rows: %r" % (len(legK), legK))
        rep.check("both kanbans present in the legacy rows (%r + %r)" % (PROBE_KANBAN, NEW_KANBAN),
                  set(r[3] for r in legK) == {PROBE_KANBAN, NEW_KANBAN},
                  str([r[3] for r in legK]))
        rep.check("rebuild feed MATCHES legacy: 2 rows, BOTH kanbans (INNER join fan-out, not TOP 1)",
                  rebK == legK and len(rebK) == 2,
                  "rebuild=%r legacy=%r" % (rebK, legK))
        sql(INV, "DELETE FROM INV_FORECAST_DETAIL_INF WHERE VC_ADD='%s'" % FC_SENTINEL)
        afterK = rebuild_wire_for(probeAsn)
        rep.check("after sweeping the 2nd forecast row the feed is back to 1 row", len(afterK) == 1,
                  "%d rows" % len(afterK))

        # --- (#4) ATOMICITY: a post-write DB failure must leave NO final 856 file on disk ---------
        # send_856 writes the EDI to <name>.tmp INSIDE the tx, flips C->S, COMMITs, and ONLY THEN
        # renames .tmp -> final. Force the flip (step 6, AFTER the .tmp write) to raise -> the driver
        # rolls back, deletes the orphan .tmp, and never renames -> the final 856 must NOT exist. (If
        # the final file were written before the commit, a rolled-back send would leave a phantom ASN
        # on disk under an uncommitted EIN — the adversary blocker.)
        print("\n--- (#4) ATOMICITY: post-write DB failure -> rollback -> NO final 856 file ---")
        atomdir = tempfile.mkdtemp(prefix="edi856_atom_")
        try:
            # re-arm the probe ASN to 'C' (untouched here) and snapshot the EIN seq so we can prove the
            # bump unwound too. Inject a fault into the REAL driver's flip UPDATE via the shim.
            preSeq = int(scalar(INV, "SELECT IN_EIN_SEQ FROM INV_SITES WHERE IN_SITE_ID=%d" % SITE_ID))
            realUpdate = edi856.system.db.runPrepUpdate
            realWrite = edi856.system.file.writeFile
            wroteTmp = {"path": None}      # spy: record that the .tmp WAS actually written (step 5)

            def faulting_update(sqlText, args, db=None, getKey=False, tx=None):
                # fail ONLY on the C->S flip (step 6), AFTER the .tmp has been written (step 5). The EIN
                # stamp UPDATE (step 2) must still run so the driver reaches the .tmp write.
                if "VC_ASN_STATUS" in sqlText and "= 'S'" in sqlText:
                    raise RuntimeError("INJECTED post-write DB failure (C->S flip)")
                return realUpdate(sqlText, args, db=db, getKey=getKey, tx=tx)

            def spying_write(path, data, append=False):
                # confirm the driver DID write the .tmp before the failure — so the cleanup proof below
                # is non-vacuous (a real file was created then removed, not a skipped write path).
                wroteTmp["path"] = path
                return realWrite(path, data, append)

            edi856.system.db.runPrepUpdate = faulting_update
            edi856.system.file.writeFile = spying_write
            raised = False
            try:
                edi856.send_856(probeAsn, SITE_ID, database="Inventory_Spike", outDir=atomdir,
                                fileTime="1430")
            except Exception as e:
                raised = True
                print("    (expected) send_856 raised: %s" % e)
            finally:
                edi856.system.db.runPrepUpdate = realUpdate     # un-inject
                edi856.system.file.writeFile = realWrite

            rep.check("send_856 re-raised on the post-write DB failure (did not swallow it)", raised,
                      "raised=%s" % raised)
            # NON-VACUOUS: the .tmp WAS written (step 5 ran) AND it targeted the .tmp, not the final name.
            rep.check("driver wrote the EDI to a .tmp (NOT the final name) before the failure",
                      wroteTmp["path"] is not None and wroteTmp["path"].endswith(".tmp"),
                      "wrote=%r" % wroteTmp["path"])
            files = sorted(os.listdir(atomdir))
            finalName = edi856._filename_856(PDATE)             # '85660618.txt'
            rep.check("NO final 856 file exists after the rolled-back send (atomicity)",
                      finalName not in files, "dir=%r" % files)
            rep.check("the orphan .tmp was cleaned up too (the .tmp that WAS written is now gone)",
                      (finalName + ".tmp") not in files and files == [], "dir=%r" % files)
            # the DB unwound: status stayed 'C', EIN seq NOT bumped, header EIN not left stamped+committed.
            stAtom = scalar(INV, "SELECT VC_ASN_STATUS FROM INV_ASN_MST WHERE IN_ASN_ID=%d" % probeAsn)
            rep.check("probe ASN status stayed 'C' (flip rolled back)", stAtom == "C", stAtom)
            postSeqAtom = int(scalar(INV, "SELECT IN_EIN_SEQ FROM INV_SITES WHERE IN_SITE_ID=%d"
                                         % SITE_ID))
            rep.check("INV_SITES.IN_EIN_SEQ NOT bumped (EIN alloc rolled back with the tx)",
                      postSeqAtom == preSeq, "pre=%d post=%d" % (preSeq, postSeqAtom))
        finally:
            shutil.rmtree(atomdir, ignore_errors=True)

    finally:
        cleanup(testAsn, decoyAsn, origSeq)
        shutil.rmtree(tmpdir, ignore_errors=True)
        # confirm restoration
        left = scalar(INV, "SELECT COUNT(*) FROM INV_ASN_MST WHERE VC_LINE_NAME IN ('%s','%s','%s')"
                           % (TEST_LINE, DECOY_LINE, PROBE_LINE))
        rep.check("Inventory restored as-found (test + decoy ASNs swept)", int(left) == 0,
                  "%s remaining" % left)
        if origSeq is not None:
            nowSeq = scalar(INV, "SELECT IN_EIN_SEQ FROM INV_SITES WHERE IN_SITE_ID=%d" % SITE_ID)
            rep.check("INV_SITES.IN_EIN_SEQ restored to original", str(nowSeq) == str(origSeq),
                      "now=%s orig=%s" % (nowSeq, origSeq))

    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
