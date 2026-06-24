#!/usr/bin/env python3
"""test_report_procs_d6.py — prove the D6 migration of the window-blind report procs.

The migrated procs (db/migrations/spike-report-procs-d6.sql) replace each window-blind
`JOIN INV_MANIFEST_COST_MST` with `CROSS APPLY dbo.fn_ManifestCostAt(part, production_date)`. This test
deploys _D6-suffixed copies SIDE-BY-SIDE with the still-live legacy procs (non-destructive, repeatable)
and proves, against the live spike DB (fixture-disciplined):
  1. the migrated proc returns the WINDOW-AWARE set (== an independent CROSS APPLY count);
  2. legacy >= migrated (the window-blind JOIN over-includes — never drops);
  3. HEADLINE D6 FIX: injecting a 2nd (non-covering, gap) price window for a part doubles the LEGACY
     output for that part (the over-bill) while the migrated output is UNCHANGED (picks the one
     covering window);
  4. EDI856 @EIN!=0 no longer hardcodes 6440 (the SELECT uses @EIN, matching its UPDATE);
  5. the EDI810/EDI856/Monthly read-branches run clean.

Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 e2e/test_report_procs_d6.py
"""
import os, subprocess, sys, re

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report  # noqa: E402

CONTAINER = os.environ.get("CONTAINER", "mssql-spike")
SA_PASS = os.environ.get("SA_PASS")
DB = "Inventory"
PROCS = ["REPORT_INVOICESSummary", "REPORT_MonthlyINVOICESSummary", "REPORT_EDI810", "REPORT_EDI856"]
MIGRATION_SQL = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..",
                                 "db", "migrations", "spike-report-procs-d6.sql"))


def sqlq(query):
    if not SA_PASS:
        sys.exit("export SA_PASS first")
    out = subprocess.check_output([
        "docker", "exec", CONTAINER, "/opt/mssql-tools18/bin/sqlcmd", "-C", "-S", "localhost",
        "-U", "sa", "-P", SA_PASS, "-d", DB, "-h", "-1", "-W", "-s", "\t",
        # SET QUOTED_IDENTIFIER ON is REQUIRED: INV_ASN_MST carries a FILTERED unique index
        # (UX_INV_ASN_MST_LINE_PDATE_NORMAL, filter [VC_START_SEQ_NUMBER]<>'-1'). sqlcmd defaults to
        # QUOTED_IDENTIFIER OFF, under which any DML against a filtered-index table fails with Msg 1934.
        # Without this the 856 seed INSERT (line ~144) failed silently and asnId corrupted to "INSERT".
        "-Q", "SET NOCOUNT ON; SET QUOTED_IDENTIFIER ON; " + query], text=True)
    return [l.split("\t") for l in out.splitlines()
            if l.strip() and not l.startswith("(") and not l.startswith("Msg ")]


def scalar(query):
    r = sqlq(query)
    return r[0][0] if r else None


def deploy_d6_copies():
    """Derive _D6-suffixed copies of the migrated procs and deploy them alongside the legacy procs.

    The procs are CREATEd under SET QUOTED_IDENTIFIER ON: a proc captures its QI setting at CREATE time
    and applies it at EXEC time. sqlcmd -i defaults to QI OFF, which would bake QI=OFF into the _D6 procs
    and make them FAIL at EXEC with Msg 1934 the moment they touch the filtered-index INV_ASN_MST table
    (the legacy procs were created QI ON in production, so they work — the divergence would be a TEST
    artifact, not a proc defect). Prepend the SET so every CREATE PROCEDURE batch captures QI=ON.
    """
    src = open(MIGRATION_SQL).read()
    for p in PROCS:
        src = re.sub(r'\b' + p + r'\b', p + "_D6", src)
    src = "SET QUOTED_IDENTIFIER ON;\nGO\n" + src
    tmp = "/tmp/report_d6.sql"
    open("/tmp/_report_d6_local.sql", "w").write(src)
    subprocess.check_call(["docker", "cp", "/tmp/_report_d6_local.sql", "%s:%s" % (CONTAINER, tmp)])
    subprocess.check_output(["docker", "exec", CONTAINER, "/opt/mssql-tools18/bin/sqlcmd", "-C",
                             "-S", "localhost", "-U", "sa", "-P", SA_PASS, "-d", DB, "-i", tmp], text=True)


def drop_d6_copies():
    sqlq("; ".join("IF OBJECT_ID('dbo.%s_D6','P') IS NOT NULL DROP PROCEDURE dbo.%s_D6" % (p, p) for p in PROCS))


def main():
    print("=" * 78)
    print(" REPORT PROCS D6 migration — window-blind JOIN -> CROSS APPLY fn_ManifestCostAt")
    print("=" * 78)
    rep = Report()

    # a production date with the most invoiced + in-window detail rows (non-vacuous parity subject)
    PDATE = scalar(
        "SELECT TOP 1 a.VC_PRODUCTION_DATE FROM INV_ASN_MST a "
        "JOIN INV_ASN_DETAIL_MST d ON a.IN_ASN_ID=d.IN_ASN_ID "
        "JOIN INV_INV_MST i ON i.IN_INV_ID=d.IN_INV_ID "
        "CROSS APPLY dbo.fn_ManifestCostAt(d.VC_ASSY_PART_NUMBER, a.VC_PRODUCTION_DATE) m "
        "GROUP BY a.VC_PRODUCTION_DATE ORDER BY COUNT(*) DESC")
    if not PDATE:
        rep.check("found a priced production date", False, "no invoiced+in-window rows")
        sys.exit(rep.summary_exit())
    print("  parity production date = %s" % PDATE)

    def legacy_count():
        return len(sqlq("EXEC REPORT_INVOICESSummary @PDate='%s'" % PDATE))

    def d6_count():
        return len(sqlq("EXEC REPORT_INVOICESSummary_D6 @PDate='%s'" % PDATE))

    deploy_d6_copies()
    try:
        # independent window-aware truth (the set the migrated proc must return)
        truth = int(scalar(
            "SELECT COUNT(*) FROM INV_ASN_MST a JOIN INV_ASN_DETAIL_MST d ON a.IN_ASN_ID=d.IN_ASN_ID "
            "JOIN INV_INV_MST i ON i.IN_INV_ID=d.IN_INV_ID "
            "CROSS APPLY dbo.fn_ManifestCostAt(d.VC_ASSY_PART_NUMBER, a.VC_PRODUCTION_DATE) m "
            "WHERE a.VC_PRODUCTION_DATE='%s'" % PDATE))
        lc, d6 = legacy_count(), d6_count()
        print("  legacy=%d  migrated=%d  window-aware-truth=%d" % (lc, d6, truth))
        rep.check("migrated proc returns the window-aware row set", d6 == truth and d6 > 0,
                  "migrated=%d truth=%d" % (d6, truth))
        rep.check("legacy (window-blind) >= migrated (never drops; over-includes out-of-window)",
                  lc >= d6, "legacy=%d migrated=%d" % (lc, d6))

        # HEADLINE: inject a 2nd, non-covering (gap) window for a part on PDATE -> legacy doubles that
        # part's lines, migrated unchanged.
        part = scalar("SELECT TOP 1 d.VC_ASSY_PART_NUMBER FROM INV_ASN_MST a "
                      "JOIN INV_ASN_DETAIL_MST d ON a.IN_ASN_ID=d.IN_ASN_ID "
                      "JOIN INV_INV_MST i ON i.IN_INV_ID=d.IN_INV_ID "
                      "CROSS APPLY dbo.fn_ManifestCostAt(d.VC_ASSY_PART_NUMBER, a.VC_PRODUCTION_DATE) m "
                      "WHERE a.VC_PRODUCTION_DATE='%s' "
                      "GROUP BY d.VC_ASSY_PART_NUMBER ORDER BY COUNT(*) DESC" % PDATE)
        part_lines = int(scalar("SELECT COUNT(*) FROM INV_ASN_MST a JOIN INV_ASN_DETAIL_MST d ON a.IN_ASN_ID=d.IN_ASN_ID "
                                "JOIN INV_INV_MST i ON i.IN_INV_ID=d.IN_INV_ID WHERE a.VC_PRODUCTION_DATE='%s' "
                                "AND d.VC_ASSY_PART_NUMBER='%s'" % (PDATE, part)))
        print("  injecting a 2nd (gap) window for part %s (%d invoiced line(s) on %s)" % (part, part_lines, PDATE))
        sqlq("INSERT INTO INV_MANIFEST_COST_MST (VC_ASSY_PART_NUMBER_CODE, VC_ASSY_MANIFEST_NUMBER, "
             "VC_START_MANIFEST, VC_END_MANIFEST, MO_PRICE) VALUES ('%s','D6','20100101','20101231',9.99)" % part)
        try:
            lc2, d62 = legacy_count(), d6_count()
            print("  after 2nd window: legacy=%d  migrated=%d" % (lc2, d62))
            rep.check("HEADLINE D6 FIX: 2nd window DOUBLES the legacy lines for that part (over-bill)",
                      lc2 == lc + part_lines, "legacy %d -> %d (+%d expected)" % (lc, lc2, part_lines))
            rep.check("HEADLINE D6 FIX: migrated UNCHANGED (picks the one covering window)",
                      d62 == d6, "migrated %d -> %d" % (d6, d62))
        finally:
            sqlq("DELETE FROM INV_MANIFEST_COST_MST WHERE VC_ASSY_PART_NUMBER_CODE='%s' AND VC_ASSY_MANIFEST_NUMBER='D6'" % part)

        # EDI856 @EIN!=0: the hardcoded 6440 is gone; the SELECT filters on @EIN (matches its UPDATE).
        # Match SERVER-SIDE via LIKE (OBJECT_DEFINITION is multi-line; sqlcmd -W would split it).
        has6440 = scalar("SELECT CASE WHEN OBJECT_DEFINITION(OBJECT_ID('dbo.REPORT_EDI856_D6')) "
                         "LIKE '%6440%' THEN 1 ELSE 0 END")
        usesEin = scalar("SELECT CASE WHEN OBJECT_DEFINITION(OBJECT_ID('dbo.REPORT_EDI856_D6')) "
                         "LIKE '%IN_ASN_EIN = @EIN%' THEN 1 ELSE 0 END")
        rep.check("EDI856 @EIN!=0 no longer hardcodes 6440", has6440 == "0", "has6440=%s" % has6440)
        rep.check("EDI856 SELECT filters on @EIN (the param)", usesEin == "1", "usesEin=%s" % usesEin)

        # ---- EDI856 forecast fan-out + GROUP BY (untested otherwise: the spike has 0 'C'-status ASNs) ----
        # Seed ONE 'C' ASN + detail + TWO forecast rows (distinct kanban = the f-join fan-out the GROUP BY
        # collapses) + a covering manifest window. EDI856 @EIN=0 must return 2 rows (one per kanban), and
        # legacy == migrated (EDI856 @EIN=0 was ALREADY windowed -> the JOIN->TVF swap is a faithful no-op).
        P8, EIN8, ST = "ZZ856PART", 99999, "2026061512000000"

        def sweep_856_fixtures():
            # Robust teardown keyed on the STABLE fixture markers (the ZZ856* part code + the EIN8 ASN),
            # NOT on asnId — so it always sweeps even if the seed INSERT failed and asnId is garbage. The
            # detail delete is by part code (independent of asnId), the ASN delete by IN_ASN_EIN.
            sqlq("DELETE FROM INV_MANIFEST_COST_MST WHERE VC_ASSY_PART_NUMBER_CODE LIKE 'ZZ856%%'; "
                 "DELETE FROM INV_FORECAST_DETAIL_INF WHERE VC_ASSY_PART_NUMBER_CODE LIKE 'ZZ856%%'; "
                 "DELETE FROM INV_ASN_DETAIL_MST WHERE VC_ASSY_PART_NUMBER LIKE 'ZZ856%%'; "
                 "DELETE FROM INV_ASN_DETAIL_MST WHERE IN_ASN_EIN=%d; "
                 "DELETE FROM INV_ASN_MST WHERE IN_ASN_EIN=%d" % (EIN8, EIN8))

        sweep_856_fixtures()  # pre-clean: a prior aborted run cannot leak into this one
        asnId = scalar("INSERT INTO INV_ASN_MST (IN_ASN_EIN,VC_ASN_STATUS,VC_LINE_NAME,VC_ASSEMBLY_LINE,"
                       "VC_START_SEQ_NUMBER,VC_END_SEQ_NUMBER,IN_QTY,VC_PRODUCTION_DATE,VC_LAST_UPDATE,VC_ADD) "
                       "OUTPUT INSERTED.IN_ASN_ID VALUES (%d,'C','ZZ856LINE','Z','0001','0002',1,'20260615','%s','%s')"
                       % (EIN8, ST, ST))
        # GUARD: under a broken QI the seed INSERT silently fails and asnId is non-numeric ("INSERT").
        # Fail LOUDLY rather than run the 856 asserts non-functionally + leak fixtures.
        try:
            asnId_i = int(asnId)
        except (TypeError, ValueError):
            asnId_i = None
        if asnId_i is None:
            rep.check("EDI856 seed INSERT succeeded (numeric IN_ASN_ID — QI ON, filtered index OK)",
                      False, "asnId=%r (filtered-index INSERT failed; check SET QUOTED_IDENTIFIER ON)" % (asnId,))
            sweep_856_fixtures()
        else:
            sqlq("INSERT INTO INV_ASN_DETAIL_MST (IN_ASN_ID,IN_ASN_EIN,VC_MANIFEST_NUMBER,VC_ASSY_PART_NUMBER,IN_QTY,VC_LAST_UPDATE) "
                 "VALUES (%d,%d,'Z9','%s',10,'%s')" % (asnId_i, EIN8, P8, ST))
            for k in ("K1", "K2"):
                sqlq("INSERT INTO INV_FORECAST_DETAIL_INF (VC_ASSY_PART_NUMBER_CODE,VC_EFFECTIVE_MONTH,"
                     "VC_TIRE_PART_NUMBER_CODE,VC_WHEEL_PART_NUMBER_CODE,VC_BROADCAST_CODE,VC_ASSY_KANBAN_NUMBER) "
                     "VALUES ('%s','202606','T','W','BC','%s')" % (P8, k))
            sqlq("INSERT INTO INV_MANIFEST_COST_MST (VC_ASSY_PART_NUMBER_CODE,VC_ASSY_MANIFEST_NUMBER,"
                 "VC_START_MANIFEST,VC_END_MANIFEST,MO_PRICE) VALUES ('%s','Z9','20250101','20281231',100)" % P8)
            try:
                leg856 = len(sqlq("EXEC REPORT_EDI856 @EIN=0"))
                mig856 = len(sqlq("EXEC REPORT_EDI856_D6 @EIN=0"))
                rep.check("EDI856 @EIN=0 exercises forecast fan-out + GROUP BY (2 kanbans -> 2 rows)",
                          mig856 == 2, "migrated rows=%d" % mig856)
                rep.check("EDI856 @EIN=0 parity: legacy == migrated (windowed JOIN -> TVF faithful)",
                          leg856 == mig856, "legacy=%d migrated=%d" % (leg856, mig856))
            finally:
                sweep_856_fixtures()

        # EDI810 @EIN=0 + Monthly: PARITY (legacy == migrated) — non-vacuous (catches any divergence the
        # CROSS APPLY swap might introduce), unlike a bare len>=0. Monthly has rows for the parity month.
        rep.check("EDI810 @EIN=0 parity: legacy == migrated",
                  len(sqlq("EXEC REPORT_EDI810 @EIN=0")) == len(sqlq("EXEC REPORT_EDI810_D6 @EIN=0")))
        lm = len(sqlq("EXEC REPORT_MonthlyINVOICESSummary @PDate='%s'" % PDATE[:6]))
        mm = len(sqlq("EXEC REPORT_MonthlyINVOICESSummary_D6 @PDate='%s'" % PDATE[:6]))
        rep.check("Monthly parity: legacy == migrated (month %s, %d rows)" % (PDATE[:6], mm), lm == mm,
                  "legacy=%d migrated=%d" % (lm, mm))
    finally:
        drop_d6_copies()

    leftover = int(scalar("SELECT COUNT(*) FROM sys.procedures WHERE name LIKE 'REPORT_%%_D6'") or 0)
    inj = int(scalar("SELECT COUNT(*) FROM INV_MANIFEST_COST_MST WHERE VC_ASSY_MANIFEST_NUMBER='D6'") or 0)
    # also confirm NO ZZ856* fixture rows leaked (the QI-OFF bug used to leak cost+forecast rows).
    zz = int(scalar(
        "SELECT (SELECT COUNT(*) FROM INV_MANIFEST_COST_MST WHERE VC_ASSY_PART_NUMBER_CODE LIKE 'ZZ856%%') "
        "+ (SELECT COUNT(*) FROM INV_FORECAST_DETAIL_INF WHERE VC_ASSY_PART_NUMBER_CODE LIKE 'ZZ856%%') "
        "+ (SELECT COUNT(*) FROM INV_ASN_DETAIL_MST WHERE VC_ASSY_PART_NUMBER LIKE 'ZZ856%%') "
        "+ (SELECT COUNT(*) FROM INV_ASN_MST WHERE IN_ASN_EIN=99999)") or 0)
    rep.check("fixture clean (_D6 procs dropped, injected window removed, ZZ856* swept)",
              leftover == 0 and inj == 0 and zz == 0,
              "_D6 procs=%d injected=%d zz856=%d" % (leftover, inj, zz))
    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
