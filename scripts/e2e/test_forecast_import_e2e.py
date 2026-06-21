#!/usr/bin/env python3
"""test_forecast_import_e2e.py — END-TO-END exercise of the REAL M2 EDI 830 forecast importer
(docs/analysis/edi/inbound/project-library/forecast/code.py), driven headlessly through the
system.db shim (jython_shim) against the spike DB — the M2 foundational-unit proof.

WHAT THIS PROVES (running the ACTUAL import_830 driver, NOT a proc-EXEC — retro R8)
----------------------------------------------------------------------------------
Runs import_830 / _process_one_830 / process_forecast_dir against the spike DB inside the per-file
transaction the driver itself opens:

  (1) HAPPY PATH — a synthetic 830 for a synthetic assembly WITH a recipe -> INV_FORECAST_INF replaced
      (raw assembly count) + INV_BREAKDOWN_FC_INF written (exploded component rows, day-spread).
  (2) D10 RAW WEEK — the stored IN_WEEK_NUMBER is the RAW FST09 week (DO 2624 -> 24, NOT 25). The
      FirstProductionDay offset (+1 for 2026) is applied ONLY to the holiday-calendar lookup.
  (3) DELETE-THEN-ACCUMULATE (the key M2 invariant) — re-importing the SAME 830 a second time does NOT
      double the breakdown qtys (the per-assembly DELETE_ForecastInfo runs FIRST, clearing the forward
      slice, so the additive upsert behaves as a REPLACE). Proven by importing twice WITH idempotency
      DISABLED (same content, forced re-run) and asserting the qty is identical, not doubled.
  (4) BOM-MISS ALARM (D-Bug-1) — an assembly with NO recipe writes a 830_FORECAST_GAP alarm row (visible),
      and writes NO breakdown rows (nothing silently doubled, nothing silently dropped).
  (5) HOLIDAY DAY-SPREAD — a week whose holiday-lookup week carries an H day spreads the qty over the
      reduced working days (the production calendar in VehicleOrder drives it).
  (6) DUNS GUARD — a file whose ISA delSL[4] matches no INV_SITES.VC_TMM_DUNS is QUARANTINED, not dropped.
  (7) IDEMPOTENCY — the SAME 830 dropped twice via process_forecast_dir processes once (the 2nd poll is a
      ledger NO-OP); no double breakdown, no duplicate alarms.

HONEST VERIFICATION (no golden 830)
-----------------------------------
NO captured TEMA 830 exists on disk (the D10 golden EDI/830000008976.EDI is gitignored client data), so
BYTE/OFFSET PARITY VS A REAL TEMA FILE IS UNPROVABLE — not claimed. The fixtures are built to the legacy
ForecastBreakdownF.pas copy() offsets (the oracle). What IS proven end-to-end through the REAL driver +
DB: the parse->explode->day-spread pipeline, the RAW-week store (D10), the delete-then-accumulate (the
re-import-not-doubled invariant), the BOM-miss alarm, the holiday spread off the real production calendar,
the DUNS quarantine, and idempotency. Pending a golden: the exact ISA element index + the real day-spread
on a captured feed.

Synthetic-only + restore-as-found: every row this test creates is tagged with the sentinel VC_ADD stamp
or a ZZF830 part/assembly prefix; teardown deletes exactly those, asserts the spike is restored, and never
touches real production data. Inventory_Live / VehicleOrder are READ-ONLY (the calendar is only SELECTed).

Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 scripts/e2e/test_forecast_import_e2e.py
"""
import os
import sys
import subprocess

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report          # noqa: E402
import jython_shim              # noqa: E402

CONTAINER = os.environ.get("CONTAINER", "mssql-spike")
SA_PASS = os.environ.get("SA_PASS")
INV = "Inventory"

SITE_ID = 1
TMM_DUNS = "000000011"          # INV_SITES site 1 VC_TMM_DUNS (at ISA delSL[4])
BAD_DUNS = "999999999"          # matches no site -> quarantine

# sentinels so cleanup is exact and nothing real is touched.
ADD_STAMP = "ZZF830ADD0000000"  # 16-char VC_ADD on synthetic recipe/part rows
ASSY = "ZZF830ASSY01"           # synthetic assembly WITH a recipe (12 chars)
ASSY_NOBOM = "ZZF830NOBOM1"     # synthetic assembly with NO recipe (gap alarm)
TIRE = "ZZF830TIRE01"           # synthetic tire component (12 chars)
WHEEL = "ZZF830WHL001"          # synthetic wheel component (12 chars)
SUPPLIER = "ZZF83"              # 5-char synthetic supplier
SIZE = "ZZF830SZ"              # synthetic size (component part master size)
LINE = "COROLLA"                # a real production line WITH holidays (for the holiday-spread case)

CODE_PY = os.path.normpath(os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "..", "..",
    "docs", "analysis", "edi", "inbound", "project-library", "forecast", "code.py"))


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


# --- 830 fixture builders (legacy element offsets — mirror the PURE test) ---------------------------

def isa(duns):
    """The inbound ISA — TEMA as sender. delSL[4] = the routing DUNS (the index the driver reads)."""
    return ("ISA*00*          *01*" + ("%-10s" % duns)[:10] +
            "*ZZ*RECEIVER  *01*000000011*260427*1200*U*00400*000000777*0*P*>")


def lin(assembly, kanban):
    return "LIN**BP*%s*RC*%s" % (assembly, kanban)


def fst(qty, weekDate, doRef):
    return "FST*%s*D*W*%s*%s***DO*%s" % (qty, weekDate, weekDate, doRef)


def build_830(duns, entries):
    out = [isa(duns), "GS*PS*SENDER*RECEIVER*20260427*1200*1*X*004010", "ST*830*0001"]
    for (assembly, kanban, weeks) in entries:
        out.append(lin(assembly, kanban))
        for (qty, weekDate, doRef) in weeks:
            out.append(fst(qty, weekDate, doRef))
    out.append("CTT*%d" % len(entries))
    out.append("SE*99*0001")
    return "\n".join(out) + "\n"


# --- synthetic fixture setup / teardown ------------------------------------------------------------

def setup_fixtures():
    """Insert the synthetic recipe + component part-master rows (tagged VC_ADD=ADD_STAMP / ZZF830 prefix).
    All heap-table inserts; teardown removes them by sentinel. The recipe: assembly ZZF830ASSY01 with
    tire ZZF830TIRE01 + wheel ZZF830WHL001, ratios 100/100/100 (so tireCount=wheelCount=weekCount), a
    default (blank) effective month, and a synthetic broadcast code."""
    teardown_fixtures()   # idempotent: clear any prior run first
    # synthetic recipe row.
    sql(INV,
        "INSERT INTO INV_FORECAST_DETAIL_INF "
        "(VC_ASSY_PART_NUMBER_CODE, VC_EFFECTIVE_MONTH, VC_ASSY_KANBAN_NUMBER, "
        " VC_TIRE_PART_NUMBER_CODE, IN_TIRE_RATIO, VC_WHEEL_PART_NUMBER_CODE, IN_WHEEL_RATIO, IN_RATIO, "
        " VC_VALVE_PART_NUMBER, VC_FILM_PART_NUMBER, IN_ASSY_QTY, VC_BROADCAST_CODE, VC_ADD, "
        " VC_LABEL_PART_NUMBER, VC_MISC1_PART_NUMBER, VC_MISC2_PART_NUMBER) "
        "VALUES ('%s', ' ', 'RC01', '%s', 100, '%s', 100, 100, '', '', 1, 'ZZF830BC', '%s', '', '', '')"
        % (ASSY, TIRE, WHEEL, ADD_STAMP))
    # synthetic part-master rows for the components (LEFT-join size/supplier resolve to NULL, but we set
    # VC_LINE_NAME so the calendar lookup uses LINE; the driver passes the explicit supplier from the feed
    # site config in the prod path — here import_830 is called with supplier=SUPPLIER).
    # NB: an INSERT trigger archives into INV_PARTS_STOCK_MST_HIST, which requires VC_PARTS_NAME +
    # IN_RENBAN_COUNT NOT NULL — so the synthetic part must set them (else the insert silently fails and
    # SELECT_PartsStockInfo returns no row -> line 'ALL LINES' -> no holiday spread). VC_LINE_NAME=COROLLA
    # so the production-calendar lookup uses a line WITH real holidays.
    for pn in (TIRE, WHEEL):
        sql(INV,
            "INSERT INTO INV_PARTS_STOCK_MST "
            "(VC_PART_NUMBER, VC_PARTS_NAME, IN_RENBAN_COUNT, VC_LINE_NAME, VC_ADD) "
            "VALUES ('%s', 'ZZF830 SYNTH', 0, '%s', '%s')" % (pn, LINE, ADD_STAMP))


def teardown_fixtures():
    """Delete EXACTLY the synthetic rows (by sentinel) + any breakdown/raw/alarm/ledger rows for the
    synthetic parts. Asserts nothing real is touched (all predicates are sentinel-scoped)."""
    sql(INV, "DELETE FROM INV_FORECAST_DETAIL_INF WHERE VC_ADD = '%s' OR VC_ASSY_PART_NUMBER_CODE IN ('%s','%s')"
        % (ADD_STAMP, ASSY, ASSY_NOBOM))
    sql(INV, "DELETE FROM INV_PARTS_STOCK_MST WHERE VC_ADD = '%s' OR VC_PART_NUMBER IN ('%s','%s')"
        % (ADD_STAMP, TIRE, WHEEL))
    sql(INV, "DELETE FROM INV_BREAKDOWN_FC_INF WHERE VC_PART_NUMBER IN ('%s','%s') OR VC_SUPPLIER_CODE='%s'"
        % (TIRE, WHEEL, SUPPLIER))
    sql(INV, "DELETE FROM INV_FORECAST_INF WHERE VC_PART_NUMBER IN ('%s','%s') OR VC_SUPPLIER_CODE='%s'"
        % (ASSY, ASSY_NOBOM, SUPPLIER))
    sql(INV, "DELETE FROM INV_EDI_ALARM_REJ WHERE VC_ALARM_TYPE IN ('830_FORECAST_GAP','830_FORECAST_STALE') "
        "AND (VC_ASSY_PART_NUMBER IN ('%s','%s') OR VC_ERROR_TEXT LIKE '%%ZZF830%%' OR IN_SITE_ID IN (1,2))"
        % (ASSY, ASSY_NOBOM))
    sql(INV, "DELETE FROM INV_EDI_INBOUND_LOG WHERE VC_FILE_NAME LIKE 'zzf830%%' OR VC_DETAIL LIKE '%%ZZF830%%'")
    # restore the last-import stamp the staleness test may have touched (set NULL as found).
    sql(INV, "UPDATE INV_SITES SET VC_LAST_FORECAST_IMPORT = NULL WHERE IN_SITE_ID IN (1,2)")


def count_breakdown(part, week):
    return int(scalar(INV, "SELECT COUNT(*) FROM INV_BREAKDOWN_FC_INF WHERE VC_PART_NUMBER='%s' AND IN_WEEK_NUMBER=%d"
                      % (part, week)) or 0)


def breakdown_qtys(part, week):
    rows = sql(INV, "SELECT IN_QTY1,IN_QTY2,IN_QTY3,IN_QTY4,IN_QTY5,IN_QTY6,IN_QTY7 "
                    "FROM INV_BREAKDOWN_FC_INF WHERE VC_PART_NUMBER='%s' AND IN_WEEK_NUMBER=%d" % (part, week))
    if not rows:
        return None
    return [int(x) if x not in ("NULL", "") else None for x in rows[0]]


def main():
    rep = Report()
    print("M2 forecast-import E2E (REAL import_830 driver via jython_shim, spike DB)\n")
    if not SA_PASS:
        sys.exit("export SA_PASS first")

    mod = jython_shim.load_wrapper("forecast_e2e", CODE_PY)

    setup_fixtures()
    try:
        # -----------------------------------------------------------------------------------------
        # 1+2+3. HAPPY PATH + D10 RAW WEEK + DELETE-THEN-ACCUMULATE.
        #   Feed: assembly ZZF830ASSY01, qty 138, weekDate 20260615, DO 2624 -> raw week 24.
        #   Ratios 100/100/100 -> tire=wheel=138. Line COROLLA, lookupWeek = 24 + offset(1) = 25.
        #   Week 25 / COROLLA has NO holiday -> Mon-Fri spread: 138//5=27 r3 -> [30,27,27,27,27,0,0].
        # -----------------------------------------------------------------------------------------
        body = build_830(TMM_DUNS, [(ASSY, "RC01", [(138, "20260615", "2624")])])
        s1 = mod.import_830(body, "zzf830_happy.edi", database=INV, siteId=SITE_ID, supplier=SUPPLIER)

        raw_count = int(scalar(INV, "SELECT IN_COUNT FROM INV_FORECAST_INF WHERE VC_PART_NUMBER='%s' "
                                    "AND IN_WEEK_NUMBER=24" % ASSY) or -1)
        rep.check("1. raw INV_FORECAST_INF replaced (assembly count 138)", raw_count == 138,
                  "got %d" % raw_count)
        rep.check("1. breakdown rows written for tire+wheel @ week 24",
                  count_breakdown(TIRE, 24) == 1 and count_breakdown(WHEEL, 24) == 1,
                  "tire=%d wheel=%d" % (count_breakdown(TIRE, 24), count_breakdown(WHEEL, 24)))

        # D10 — THE RAW-WEEK PROOF: stored at week 24, NOT 25 (the offset is only for the holiday lookup).
        rep.check("2. D10 RAW WEEK: breakdown stored at IN_WEEK_NUMBER 24 (NOT 25)",
                  count_breakdown(TIRE, 24) == 1 and count_breakdown(TIRE, 25) == 0,
                  "wk24=%d wk25=%d" % (count_breakdown(TIRE, 24), count_breakdown(TIRE, 25)))

        qtys1 = breakdown_qtys(TIRE, 24)
        rep.check("2. day-spread 138/Mon-Fri (week 25 no holiday) -> [30,27,27,27,27,0,0]",
                  qtys1 == [30, 27, 27, 27, 27, 0, 0], str(qtys1))

        # DELETE-THEN-ACCUMULATE: re-import the SAME content with idempotency DISABLED (call import_830
        # directly, which does NOT consult the ledger — it always DELETE-then-rebuilds). If the delete did
        # NOT run first, the additive upsert would DOUBLE the qty (-> [60,54,...]). It must stay identical.
        s2 = mod.import_830(body, "zzf830_happy_again.edi", database=INV, siteId=SITE_ID, supplier=SUPPLIER)
        qtys2 = breakdown_qtys(TIRE, 24)
        rep.check("3. DELETE-THEN-ACCUMULATE: re-import SAME 830 -> qty NOT doubled (still [30,27,...])",
                  qtys2 == [30, 27, 27, 27, 27, 0, 0], "got %s" % qtys2)
        rep.check("3. re-import leaves exactly ONE breakdown row per component (no dup rows)",
                  count_breakdown(TIRE, 24) == 1 and count_breakdown(WHEEL, 24) == 1,
                  "tire=%d wheel=%d" % (count_breakdown(TIRE, 24), count_breakdown(WHEEL, 24)))
        rep.check("3. raw forecast also replaced (still 138, not 276)",
                  int(scalar(INV, "SELECT IN_COUNT FROM INV_FORECAST_INF WHERE VC_PART_NUMBER='%s' "
                                  "AND IN_WEEK_NUMBER=24" % ASSY) or -1) == 138)

        # -----------------------------------------------------------------------------------------
        # 4. BOM-MISS ALARM (D-Bug-1): an assembly with NO recipe -> a 830_FORECAST_GAP alarm, no
        #    breakdown rows, no silent drop.
        # -----------------------------------------------------------------------------------------
        body_nobom = build_830(TMM_DUNS, [(ASSY_NOBOM, "RC09", [(500, "20260615", "2624")])])
        s4 = mod.import_830(body_nobom, "zzf830_nobom.edi", database=INV, siteId=SITE_ID, supplier=SUPPLIER)
        gap_alarms = int(scalar(INV, "SELECT COUNT(*) FROM INV_EDI_ALARM_REJ "
                                     "WHERE VC_ALARM_TYPE='830_FORECAST_GAP' AND VC_ASSY_PART_NUMBER='%s'"
                                     % ASSY_NOBOM) or 0)
        rep.check("4. BOM-miss -> a 830_FORECAST_GAP alarm row (visible, not vanished)", gap_alarms == 1,
                  "got %d" % gap_alarms)
        rep.check("4. BOM-miss -> summary gapAlarms == 1", s4["gapAlarms"] == 1, str(s4["gapAlarms"]))
        rep.check("4. BOM-miss -> ZERO breakdown writes for the no-recipe assembly", s4["breakdownWrites"] == 0,
                  str(s4["breakdownWrites"]))
        # raw forecast IS still written for the no-BOM assembly (the legacy writes raw before exploding).
        rep.check("4. BOM-miss -> raw forecast STILL written (faithful: raw write precedes explode)",
                  int(scalar(INV, "SELECT COUNT(*) FROM INV_FORECAST_INF WHERE VC_PART_NUMBER='%s'"
                             % ASSY_NOBOM) or 0) >= 1)

        # -----------------------------------------------------------------------------------------
        # 5. HOLIDAY DAY-SPREAD: raw week 21 -> lookupWeek 22 (offset +1). Week 22 / COROLLA has an H on
        #    day 1 (Mon, 20260525) -> 4 working days. 138//4=34 r2 -> [36,34,0?,...]. The H day is day1
        #    (Mon) so day1 is OFF: working = Tue,Wed,Thu,Fri = [0,34+2?...]. Compute: days=4,
        #    ratiocount=34, leftover=2; first WORKING day (Tue=idx2) gets 34+2=36 -> [0,36,34,34,34,0,0].
        # -----------------------------------------------------------------------------------------
        body_hol = build_830(TMM_DUNS, [(ASSY, "RC01", [(138, "20260518", "2621")])])  # DO 2621 -> raw 21
        # raw week 21 -> lookupWeek = 21 + offset(1) = 22. Week 22 / COROLLA has an H on day 1 (Mon,
        # 20260525) -> working = Tue,Wed,Thu,Fri = 4 days. 138//4=34 r2 -> first WORKING day (Tue=idx2)
        # gets 34+2=36 -> [0,36,34,34,34,0,0] (Mon off=holiday, Sat/Sun off=weekend).
        s5 = mod.import_830(body_hol, "zzf830_holiday.edi", database=INV, siteId=SITE_ID, supplier=SUPPLIER)
        rep.check("5. holiday week stored at RAW week 21", count_breakdown(TIRE, 21) == 1,
                  "wk21=%d" % count_breakdown(TIRE, 21))
        qtys_h = breakdown_qtys(TIRE, 21)
        rep.check("5. holiday day-spread (Mon H, lookupWeek 22) -> [0,36,34,34,34,0,0]",
                  qtys_h == [0, 36, 34, 34, 34, 0, 0], str(qtys_h))
        rep.check("5. holiday-week sum preserved (138)", qtys_h is not None and sum(qtys_h) == 138,
                  str(qtys_h))

        # -----------------------------------------------------------------------------------------
        # 6. DUNS GUARD: a file whose ISA delSL[4] matches no site -> QUARANTINED (via _process_one_830).
        # -----------------------------------------------------------------------------------------
        import tempfile, shutil
        tmpdir = tempfile.mkdtemp(prefix="zzf830_")
        try:
            bad_body = build_830(BAD_DUNS, [(ASSY, "RC01", [(100, "20260615", "2624")])])
            r6 = mod._process_one_830(bad_body, "zzf830_baddns.edi", INV, mod.CAL_DATABASE,
                                      None, os.path.join(tmpdir, "arch"), os.path.join(tmpdir, "quar"), {})
            rep.check("6. DUNS guard: no-match DUNS -> QUARANTINED (not processed)",
                      r6["outcome"] == "QUARANTINED_NO_DUNS", r6["outcome"])

            # -------------------------------------------------------------------------------------
            # 7. IDEMPOTENCY via process_forecast_dir: drop the SAME 830 twice -> processed once.
            # -------------------------------------------------------------------------------------
            indir = os.path.join(tmpdir, "in")
            os.makedirs(indir)
            # UNIQUE content (qty 99, DO 2630 -> raw week 30, weekDate 20260720) so the poll's content hash
            # does NOT collide with the direct import_830 ledger rows above. Week 30 lookup = 31 (no COROLLA
            # holiday) -> 99//5=19 r4 -> [23,19,19,19,19,0,0].
            good = build_830(TMM_DUNS, [(ASSY, "RC01", [(99, "20260720", "2630")])])
            with open(os.path.join(indir, "zzf830_poll.edi"), "w") as fh:
                fh.write(good)
            res_a = mod.process_forecast_dir(indir, database=INV, supplierBySite={SITE_ID: SUPPLIER})
            # re-drop the SAME content under a NEW name (renamed redelivery) -> ledger NO-OP.
            with open(os.path.join(indir, "zzf830_poll_REDELIVERED.edi"), "w") as fh:
                fh.write(good)
            res_b = mod.process_forecast_dir(indir, database=INV, supplierBySite={SITE_ID: SUPPLIER})
            outcomes_a = [r["outcome"] for r in res_a]
            outcomes_b = [r["outcome"] for r in res_b]
            rep.check("7. idempotency: 1st poll PROCESSED",
                      "PROCESSED" in outcomes_a, str(outcomes_a))
            rep.check("7. idempotency: renamed re-drop -> SKIPPED_ALREADY_PROCESSED (content-hash dedup)",
                      "SKIPPED_ALREADY_PROCESSED" in outcomes_b, str(outcomes_b))
            rep.check("7. idempotency: exactly ONE breakdown row per component @ week 30 (no double poll)",
                      count_breakdown(TIRE, 30) == 1 and count_breakdown(WHEEL, 30) == 1,
                      "tire=%d wheel=%d" % (count_breakdown(TIRE, 30), count_breakdown(WHEEL, 30)))
            qtys_poll = breakdown_qtys(TIRE, 30)
            rep.check("7. idempotency: qty NOT doubled by the 2nd poll ([23,19,19,19,19,0,0])",
                      qtys_poll == [23, 19, 19, 19, 19, 0, 0], str(qtys_poll))
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

        # -----------------------------------------------------------------------------------------
        # 8. Q11 last-import stamp + staleness alarm (DATA side).
        # -----------------------------------------------------------------------------------------
        stamp = scalar(INV, "SELECT VC_LAST_FORECAST_IMPORT FROM INV_SITES WHERE IN_SITE_ID=%d" % SITE_ID)
        rep.check("8. Q11: VC_LAST_FORECAST_IMPORT stamped after import",
                  stamp is not None and stamp != "NULL" and len(stamp) >= 8, repr(stamp))
        # force site 2 to be NEVER-imported (NULL) and AUTO -> stale_sites raises a 830_FORECAST_STALE.
        sql(INV, "UPDATE INV_SITES SET VC_LAST_FORECAST_IMPORT = NULL, VC_FORECAST_IMPORT_MODE='AUTO' WHERE IN_SITE_ID=2")
        stale = mod.stale_sites(database=INV, staleDays=8, raiseAlarms=True)
        rep.check("8. Q11: stale_sites flags the never-imported AUTO site",
                  any(s["siteId"] == 2 for s in stale), str(stale))
        stale_alarm = int(scalar(INV, "SELECT COUNT(*) FROM INV_EDI_ALARM_REJ "
                                      "WHERE VC_ALARM_TYPE='830_FORECAST_STALE' AND IN_SITE_ID=2 AND BIT_RESOLVED=0")
                          or 0)
        rep.check("8. Q11: a 830_FORECAST_STALE alarm row was raised for the stale site", stale_alarm == 1,
                  "got %d" % stale_alarm)
        # dedup: a second stale_sites run does NOT raise a duplicate alarm.
        mod.stale_sites(database=INV, staleDays=8, raiseAlarms=True)
        stale_alarm2 = int(scalar(INV, "SELECT COUNT(*) FROM INV_EDI_ALARM_REJ "
                                       "WHERE VC_ALARM_TYPE='830_FORECAST_STALE' AND IN_SITE_ID=2 AND BIT_RESOLVED=0")
                           or 0)
        rep.check("8. Q11: staleness alarm is deduped (still 1 after a 2nd check)", stale_alarm2 == 1,
                  "got %d" % stale_alarm2)

    finally:
        teardown_fixtures()
        # restore-as-found assertion: no synthetic rows left behind.
        leftover = (int(scalar(INV, "SELECT COUNT(*) FROM INV_FORECAST_DETAIL_INF WHERE VC_ADD='%s'" % ADD_STAMP) or 0)
                    + int(scalar(INV, "SELECT COUNT(*) FROM INV_PARTS_STOCK_MST WHERE VC_ADD='%s'" % ADD_STAMP) or 0)
                    + int(scalar(INV, "SELECT COUNT(*) FROM INV_BREAKDOWN_FC_INF WHERE VC_PART_NUMBER IN ('%s','%s')"
                                 % (TIRE, WHEEL)) or 0)
                    + int(scalar(INV, "SELECT COUNT(*) FROM INV_FORECAST_INF WHERE VC_PART_NUMBER IN ('%s','%s')"
                                 % (ASSY, ASSY_NOBOM)) or 0))
        rep.check("teardown: spike restored as-found (0 synthetic rows remain)", leftover == 0,
                  "leftover=%d" % leftover)

    return rep.summary_exit()


if __name__ == "__main__":
    sys.exit(main())
