#!/usr/bin/env python3
"""test_stocktaking_writepost.py — row-level PARITY integration test for the stocktaking write path.

Proves, against the LIVE spike DB (fixture-disciplined — leaves the DB exactly as found), that the
REBUILD write-then-post path (source-row write + POST_StockMovement) moves INV_PARTS_STOCK_MST.IN_QTY
IDENTICALLY to the LEGACY stocktaking triggers, for insert / amend / delete. This is the first
row-level rebuild==legacy proof for a producer (complements the aggregate reconciliation harness
test_stock_ledger_parity.py), and the first time POST_StockMovement is exercised end-to-end behind a
real source-table write.

Why two arms: the rebuild write-then-post and the legacy triggers BOTH write IN_QTY, so they cannot
co-run without double-counting. The legacy arm runs with the three stocktaking triggers ENABLED (the
trigger moves IN_QTY); the rebuild arm runs with them DISABLED (only POST_StockMovement moves IN_QTY).
UPDATE_PartNumber on INV_PARTS_STOCK_MST is disabled for BOTH arms so neither leaves INV_PART_QTY_INF
audit residue (same precaution as the parity harness self-test). Everything is restored at the end.

Mirrors the gateway wrapper (stocktaking.insertAdjustment/amendAdjustment/deleteAdjustment) at the SQL
level — the wrapper is thin Jython glue over exactly this proc/statement sequence (its delta logic is
unit-tested in test_stocktaking_posts.py; this proves the DB composition + legacy parity).

Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 e2e/test_stocktaking_writepost.py
"""
import os, subprocess, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report  # noqa: E402

CONTAINER = os.environ.get("CONTAINER", "mssql-spike")
SA_PASS = os.environ.get("SA_PASS")
DB = "Inventory"
PART = 16                       # '478930223000' — the part the ledger parity self-test also uses
PARTNUM = "478930223000"
REASON = "WRITEPOST_TEST"       # unique marker so we can find/clean only our rows
Q1, Q2 = 137, 200               # insert qty, then amend-to qty (signed adjustments, D5)


def sql(query):
    if not SA_PASS:
        sys.exit("export SA_PASS first (see scripts/spike-db.sh)")
    out = subprocess.check_output([
        "docker", "exec", CONTAINER, "/opt/mssql-tools18/bin/sqlcmd", "-C", "-S", "localhost",
        "-U", "sa", "-P", SA_PASS, "-d", DB, "-h", "-1", "-W", "-s", "\t",
        "-Q", "SET NOCOUNT ON; " + query], text=True)
    rows = []
    for line in out.splitlines():
        line = line.rstrip("\n")
        if not line or line.startswith("(") or line.startswith("Msg "):
            continue
        rows.append(line.split("\t"))
    return rows


def scalar(query):
    r = sql(query)
    return r[0][0] if r else None


def qty():
    return int(scalar("SELECT IN_QTY FROM INV_PARTS_STOCK_MST WHERE IN_PART_ID=%d" % PART))


STAMP = ("CONVERT(varchar,getdate(),112)+SUBSTRING(CONVERT(varchar,getdate(),114),1,2)"
         "+SUBSTRING(CONVERT(varchar,getdate(),114),4,2)+SUBSTRING(CONVERT(varchar,getdate(),114),7,2)"
         "+SUBSTRING(CONVERT(varchar,getdate(),114),10,2)")


def cleanup(base):
    """Remove any test rows + ledger rows, restore IN_QTY, re-enable all triggers. Idempotent."""
    sql("DELETE FROM INV_STOCK_LEDGER WHERE VC_SOURCE_EVENT LIKE 'STOCKTAKING:stk=%%' "
        "AND VC_REASON='%s'; "
        "DELETE FROM INV_STOCKTAKING_INF WHERE VC_REASON='%s'; "
        "UPDATE INV_PARTS_STOCK_MST SET IN_QTY=%d WHERE IN_PART_ID=%d; "
        "ENABLE TRIGGER INSERT_Stocktaking ON INV_STOCKTAKING_INF; "
        "ENABLE TRIGGER UPDATE_Stocktaking ON INV_STOCKTAKING_INF; "
        "ENABLE TRIGGER DELETE_Stocktaking ON INV_STOCKTAKING_INF; "
        "ENABLE TRIGGER UPDATE_PartNumber ON INV_PARTS_STOCK_MST"
        % (REASON, REASON, base, PART))


def legacy_arm(rep, base):
    print("\n--- ARM 1: LEGACY (stocktaking triggers ENABLED) ---")
    # INSERT via the legacy proc -> INSERT_Stocktaking adds +Q1.
    sql("EXEC INSERT_StockTakingInfo @PartNumber='%s', @QTY=%d, @Reason='%s'" % (PARTNUM, Q1, REASON))
    ins = qty() - base
    stk = scalar("SELECT MAX(IN_STOCKTAKING_ID) FROM INV_STOCKTAKING_INF WHERE VC_REASON='%s'" % REASON)
    # AMEND Q1->Q2 via the legacy proc -> UPDATE_Stocktaking nets +(Q2-Q1).
    sql("EXEC UPDATE_StockTakingInfo @SupCode='', @PartCode='%s', @QTY=%d, @Reason='%s', @StocktakingID=%s"
        % (PARTNUM, Q2, REASON, stk))
    amend = qty() - base - Q1
    # DELETE via the legacy proc -> DELETE_Stocktaking reverses -Q2.
    sql("EXEC DELETE_StocktakingInfo @StocktakingID=%s" % stk)
    dele = qty() - base
    rep.check("legacy insert delta == +Q1", ins == Q1, "%d" % ins)
    rep.check("legacy amend delta == +(Q2-Q1)", amend == (Q2 - Q1), "%d" % amend)
    rep.check("legacy after delete back to base", dele == 0, "delta %d" % dele)
    return {"insert": ins, "amend": amend, "delete_to": qty()}


def rebuild_arm(rep, base):
    print("\n--- ARM 2: REBUILD (stocktaking triggers DISABLED; only POST_StockMovement moves IN_QTY) ---")
    sql("DISABLE TRIGGER INSERT_Stocktaking ON INV_STOCKTAKING_INF; "
        "DISABLE TRIGGER UPDATE_Stocktaking ON INV_STOCKTAKING_INF; "
        "DISABLE TRIGGER DELETE_Stocktaking ON INV_STOCKTAKING_INF")

    # INSERT source row (no trigger) + post +Q1.
    sql("INSERT INTO INV_STOCKTAKING_INF (IN_PART_ID, IN_QTY, VC_REASON, VC_LAST_UPDATE, VC_ADD) "
        "VALUES (%d, %d, '%s', %s, %s)" % (PART, Q1, REASON, STAMP, STAMP))
    stk = scalar("SELECT MAX(IN_STOCKTAKING_ID) FROM INV_STOCKTAKING_INF WHERE VC_REASON='%s'" % REASON)
    sql("EXEC POST_StockMovement @partId=%d, @delta=%d, @sourceEnum='STOCKTAKING', @sourceRowId=%s, "
        "@eventKey='STOCKTAKING:stk=%s:ins', @reason='%s', @site=1, @purge=0"
        % (PART, Q1, stk, stk, REASON))
    ins = qty() - base
    led_ins = int(scalar("SELECT IN_QTY_CHANGE FROM INV_STOCK_LEDGER WHERE VC_SOURCE_EVENT='STOCKTAKING:stk=%s:ins'" % stk))

    # AMEND source row to Q2 (fresh non-NULL stamp; no trigger) + post net +(Q2-Q1).
    sql("UPDATE INV_STOCKTAKING_INF SET IN_QTY=%d, VC_LAST_UPDATE=%s WHERE IN_STOCKTAKING_ID=%s"
        % (Q2, STAMP, stk))
    nullstamp = int(scalar("SELECT COUNT(*) FROM INV_STOCKTAKING_INF WHERE IN_STOCKTAKING_ID=%s "
                           "AND VC_LAST_UPDATE IS NOT NULL AND LEN(VC_LAST_UPDATE)=16" % stk))
    sql("EXEC POST_StockMovement @partId=%d, @delta=%d, @sourceEnum='STOCKTAKING', @sourceRowId=%s, "
        "@eventKey='STOCKTAKING:stk=%s:upd', @reason='%s', @site=1, @purge=0"
        % (PART, Q2 - Q1, stk, stk, REASON))
    amend = qty() - base - Q1

    # DELETE: post reversal -Q2, then remove the source row (no trigger). NOTE: this test asserts the
    # qty DELTAS, not the wrapper's exact statement order (the wrapper reads-old -> deletes -> posts
    # from the snapshot); both orders are qty-equivalent because the reversal is computed from `old`.
    sql("EXEC POST_StockMovement @partId=%d, @delta=%d, @sourceEnum='STOCKTAKING', @sourceRowId=%s, "
        "@eventKey='STOCKTAKING:stk=%s:del', @reason='%s', @site=1, @purge=0"
        % (PART, -Q2, stk, stk, REASON))
    sql("DELETE FROM INV_STOCKTAKING_INF WHERE IN_STOCKTAKING_ID=%s" % stk)
    dele = qty() - base
    led_sum = int(scalar("SELECT ISNULL(SUM(IN_QTY_CHANGE),0) FROM INV_STOCK_LEDGER "
                         "WHERE VC_SOURCE_EVENT LIKE 'STOCKTAKING:stk=%s:%%'" % stk))

    rep.check("rebuild insert delta == +Q1 (POST_StockMovement bumped IN_QTY)", ins == Q1, "%d" % ins)
    rep.check("rebuild ledger row recorded +Q1", led_ins == Q1, "%d" % led_ins)
    rep.check("rebuild amend writes a non-NULL 16-char VC_LAST_UPDATE (fixes Bug2)", nullstamp == 1)
    rep.check("rebuild amend delta == +(Q2-Q1)", amend == (Q2 - Q1), "%d" % amend)
    rep.check("rebuild after delete back to base", dele == 0, "delta %d" % dele)
    rep.check("rebuild ledger nets to 0 over insert+amend+delete", led_sum == 0, "%d" % led_sum)
    return {"insert": ins, "amend": amend, "delete_to": qty()}


def main():
    print("=" * 78)
    print(" STOCKTAKING write-then-post — row-level PARITY integration test (rebuild == legacy)")
    print("=" * 78)
    rep = Report()
    base = qty()
    print("  part %d base IN_QTY = %d" % (PART, base))
    # Disable the part-master audit trigger for BOTH arms (no INV_PART_QTY_INF residue).
    sql("DISABLE TRIGGER UPDATE_PartNumber ON INV_PARTS_STOCK_MST")
    try:
        leg = legacy_arm(rep, base)
        sql("UPDATE INV_PARTS_STOCK_MST SET IN_QTY=%d WHERE IN_PART_ID=%d" % (base, PART))  # reset between arms
        reb = rebuild_arm(rep, base)
        print("\n--- PARITY: rebuild deltas vs legacy deltas ---")
        rep.check("insert parity (rebuild == legacy)", reb["insert"] == leg["insert"],
                  "rebuild %d vs legacy %d" % (reb["insert"], leg["insert"]))
        rep.check("amend parity (rebuild == legacy)", reb["amend"] == leg["amend"],
                  "rebuild %d vs legacy %d" % (reb["amend"], leg["amend"]))
        rep.check("delete parity (both restore to base)", reb["delete_to"] == leg["delete_to"] == base,
                  "rebuild %d / legacy %d / base %d" % (reb["delete_to"], leg["delete_to"], base))
    finally:
        cleanup(base)

    # Fixture discipline: assert the DB is exactly as found.
    restored = qty()
    leftover_rows = int(scalar("SELECT COUNT(*) FROM INV_STOCKTAKING_INF WHERE VC_REASON='%s'" % REASON))
    leftover_led = int(scalar("SELECT COUNT(*) FROM INV_STOCK_LEDGER WHERE VC_REASON='%s'" % REASON))
    trig_off = int(scalar("SELECT COUNT(*) FROM sys.triggers WHERE is_disabled=1 AND name IN "
                          "('INSERT_Stocktaking','UPDATE_Stocktaking','DELETE_Stocktaking','UPDATE_PartNumber')"))
    rep.check("DB restored as found (IN_QTY, no test rows/ledger, triggers re-enabled)",
              restored == base and leftover_rows == 0 and leftover_led == 0 and trig_off == 0,
              "IN_QTY=%d rows=%d led=%d trig_off=%d" % (restored, leftover_rows, leftover_led, trig_off))
    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
