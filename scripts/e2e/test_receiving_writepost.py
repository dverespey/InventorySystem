#!/usr/bin/env python3
"""test_receiving_writepost.py — row-level PARITY integration test for the receiving write path.

Proves, against the LIVE spike DB (fixture-disciplined), that the REBUILD write-then-post path
(open-order write + POST_StockMovement) moves INV_PARTS_STOCK_MST.IN_QTY IDENTICALLY to the LEGACY
RecConfStat qty triggers, for the reachable 'S' add-point path: insert a SHIPPED order (+qty), amend
its qty (net), delete it (-qty). All 47 parts in this snapshot are 'S' add-point, so the 'A'-path
divergences (D7 arrival-add, D8(3) reversal, D12#3 yards) are NOT integration-reachable here — they are
covered by the pure unit test (test_receiving_posts.py) with synthetic rows, and are intentional
DIVERGENCES from legacy (not parity) anyway.

Two arms (rebuild path + legacy triggers both write IN_QTY -> can't co-run):
  legacy  — RecConfStat qty triggers ENABLED; direct INV_OPEN_ORDER_INF DML fires them.
  rebuild — those triggers DISABLED; only POST_StockMovement moves IN_QTY.
UPDATE_PartNumber disabled both arms (no audit residue). HIST rows the legacy INSERT/UPDATE triggers
snapshot are cleaned by the renban marker. Everything restored in finally.

Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 scripts/e2e/test_receiving_writepost.py
"""
import os, subprocess, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report  # noqa: E402

CONTAINER = os.environ.get("CONTAINER", "mssql-spike")
SA_PASS = os.environ.get("SA_PASS")
DB = "Inventory"
PART = 16
PARTNUM = "478930223000"        # part 16, a real 'S' add-point part
RENBAN = "WPTEST01"             # unique marker (VC_RENBAN_NUMBER) for find/clean
SHIP = "20260618"               # a non-blank ship-status stamp => 'shipped'
Q1, Q2 = 137, 200

STAMP = ("CONVERT(varchar,getdate(),112)+SUBSTRING(CONVERT(varchar,getdate(),114),1,2)"
         "+SUBSTRING(CONVERT(varchar,getdate(),114),4,2)+SUBSTRING(CONVERT(varchar,getdate(),114),7,2)"
         "+SUBSTRING(CONVERT(varchar,getdate(),114),10,2)")
# All NOT-NULL varchars get '' unless set; a SHIPPED 'S' order sets VC_STATUS_SUPPLIER_SHIPPING.
INS = ("INSERT INTO INV_OPEN_ORDER_INF (VC_SUPPLIER_CODE, VC_PART_NUMBER, VC_FRS_NUMBER, "
       "VC_RENBAN_NUMBER, IN_QTY, VC_STATUS_SUPPLIER_SHIPPING, VC_ARRIVAL, VC_TRAILER_NUMBER, "
       "VC_STATUS_PLANT_YARD, VC_PLANT_PARKING, VC_STATUS_ASSEMBLER_YARD, VC_ASSEMBLER_LOCATION, "
       "VC_STATUS_EMPTY_TRAILER, VC_DETENTION, VC_ORDER_DATE, VC_WAREHOUSE, VC_TERMINATED, "
       "VC_SHIP_DATE, VC_KANBAN_NUMBER, VC_LAST_UPDATE, VC_ADD) VALUES "
       "('WPT', '%s', 'WPT0001', '%s', %d, '%s', '', '', '', '', '', '', '', '', '', '', '', '', '', %s, %s)")


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


LEDGER_REASON = "WPTEST_RECV"   # unique marker on the rebuild arm's ledger posts (for clean teardown)


def cleanup(base):
    sql("DELETE FROM INV_STOCK_LEDGER WHERE VC_REASON='%s'; "
        "DELETE FROM INV_OPEN_ORDER_INF WHERE VC_RENBAN_NUMBER='%s'; "
        "DELETE FROM INV_OPEN_ORDER_INF_HIST WHERE VC_RENBAN_NUMBER='%s'; "
        "UPDATE INV_PARTS_STOCK_MST SET IN_QTY=%d WHERE IN_PART_ID=%d; "
        "ENABLE TRIGGER INSERT_RecConfStatPartsStockMstQTY ON INV_OPEN_ORDER_INF; "
        "ENABLE TRIGGER UPDATE_RecConfStatPartsStockMstQTY ON INV_OPEN_ORDER_INF; "
        "ENABLE TRIGGER DELETE_RecConfStatPartsStockMstQTY ON INV_OPEN_ORDER_INF; "
        "ENABLE TRIGGER UPDATE_PartNumber ON INV_PARTS_STOCK_MST"
        % (LEDGER_REASON, RENBAN, RENBAN, base, PART))


def legacy_arm(rep, base):
    print("\n--- ARM 1: LEGACY (RecConfStat qty triggers ENABLED) ---")
    sql(INS % (PARTNUM, RENBAN, Q1, SHIP, STAMP, STAMP))
    ins = qty() - base
    oid = scalar("SELECT MAX(IN_ORDER_ID) FROM INV_OPEN_ORDER_INF WHERE VC_RENBAN_NUMBER='%s'" % RENBAN)
    sql("UPDATE INV_OPEN_ORDER_INF SET IN_QTY=%d, VC_LAST_UPDATE=%s WHERE IN_ORDER_ID=%s" % (Q2, STAMP, oid))
    amend = qty() - base - Q1
    sql("DELETE FROM INV_OPEN_ORDER_INF WHERE IN_ORDER_ID=%s" % oid)
    dele = qty() - base
    rep.check("legacy insert shipped 'S' -> +Q1", ins == Q1, "%d" % ins)
    rep.check("legacy amend qty -> +(Q2-Q1)", amend == (Q2 - Q1), "%d" % amend)
    rep.check("legacy delete active shipped -> back to base", dele == 0, "delta %d" % dele)
    return {"insert": ins, "amend": amend, "delete_to": qty()}


def rebuild_arm(rep, base):
    print("\n--- ARM 2: REBUILD (RecConfStat qty triggers DISABLED) ---")
    sql("DISABLE TRIGGER INSERT_RecConfStatPartsStockMstQTY ON INV_OPEN_ORDER_INF; "
        "DISABLE TRIGGER UPDATE_RecConfStatPartsStockMstQTY ON INV_OPEN_ORDER_INF; "
        "DISABLE TRIGGER DELETE_RecConfStatPartsStockMstQTY ON INV_OPEN_ORDER_INF")
    sql(INS % (PARTNUM, RENBAN, Q1, SHIP, STAMP, STAMP))
    oid = scalar("SELECT MAX(IN_ORDER_ID) FROM INV_OPEN_ORDER_INF WHERE VC_RENBAN_NUMBER='%s'" % RENBAN)
    sql("EXEC POST_StockMovement @partId=%d, @delta=%d, @sourceEnum='RECEIVING_SHIP', @sourceRowId=%s, "
        "@eventKey='RECEIVING_SHIP:ord=%s:ins', @reason='WPTEST_RECV', @site=1, @purge=0"
        % (PART, Q1, oid, oid))
    ins = qty() - base
    led_ins = int(scalar("SELECT IN_QTY_CHANGE FROM INV_STOCK_LEDGER WHERE VC_SOURCE_EVENT='RECEIVING_SHIP:ord=%s:ins'" % oid))
    sql("UPDATE INV_OPEN_ORDER_INF SET IN_QTY=%d, VC_LAST_UPDATE=%s WHERE IN_ORDER_ID=%s" % (Q2, STAMP, oid))
    sql("EXEC POST_StockMovement @partId=%d, @delta=%d, @sourceEnum='RECEIVING_SHIP', @sourceRowId=%s, "
        "@eventKey='RECEIVING_SHIP:ord=%s:upd', @reason='WPTEST_RECV', @site=1, @purge=0"
        % (PART, Q2 - Q1, oid, oid))
    amend = qty() - base - Q1
    sql("EXEC POST_StockMovement @partId=%d, @delta=%d, @sourceEnum='RECEIVING_SHIP', @sourceRowId=%s, "
        "@eventKey='RECEIVING_SHIP:ord=%s:del', @reason='WPTEST_RECV', @site=1, @purge=0"
        % (PART, -Q2, oid, oid))
    sql("DELETE FROM INV_OPEN_ORDER_INF WHERE IN_ORDER_ID=%s" % oid)
    dele = qty() - base
    led_sum = int(scalar("SELECT ISNULL(SUM(IN_QTY_CHANGE),0) FROM INV_STOCK_LEDGER "
                         "WHERE VC_SOURCE_EVENT LIKE 'RECEIVING_SHIP:ord=%s:%%'" % oid))
    rep.check("rebuild insert shipped 'S' -> +Q1", ins == Q1, "%d" % ins)
    rep.check("rebuild ledger row recorded +Q1", led_ins == Q1, "%d" % led_ins)
    rep.check("rebuild amend qty -> +(Q2-Q1)", amend == (Q2 - Q1), "%d" % amend)
    rep.check("rebuild delete -> back to base", dele == 0, "delta %d" % dele)
    rep.check("rebuild ledger nets to 0 over insert+amend+delete", led_sum == 0, "%d" % led_sum)
    return {"insert": ins, "amend": amend, "delete_to": qty()}


def main():
    print("=" * 78)
    print(" RECEIVING write-then-post — row-level PARITY integration test ('S' path, rebuild == legacy)")
    print("=" * 78)
    rep = Report()
    base = qty()
    print("  part %d ('%s', add-point 'S') base IN_QTY = %d" % (PART, PARTNUM, base))
    sql("DISABLE TRIGGER UPDATE_PartNumber ON INV_PARTS_STOCK_MST")
    try:
        leg = legacy_arm(rep, base)
        sql("UPDATE INV_PARTS_STOCK_MST SET IN_QTY=%d WHERE IN_PART_ID=%d" % (base, PART))
        reb = rebuild_arm(rep, base)
        print("\n--- PARITY: rebuild deltas vs legacy deltas ---")
        rep.check("insert parity", reb["insert"] == leg["insert"], "rebuild %d vs legacy %d" % (reb["insert"], leg["insert"]))
        rep.check("amend parity", reb["amend"] == leg["amend"], "rebuild %d vs legacy %d" % (reb["amend"], leg["amend"]))
        rep.check("delete parity (both restore to base)", reb["delete_to"] == leg["delete_to"] == base,
                  "rebuild %d / legacy %d / base %d" % (reb["delete_to"], leg["delete_to"], base))
    finally:
        cleanup(base)

    restored = qty()
    oo = int(scalar("SELECT COUNT(*) FROM INV_OPEN_ORDER_INF WHERE VC_RENBAN_NUMBER='%s'" % RENBAN))
    hist = int(scalar("SELECT COUNT(*) FROM INV_OPEN_ORDER_INF_HIST WHERE VC_RENBAN_NUMBER='%s'" % RENBAN))
    led = int(scalar("SELECT COUNT(*) FROM INV_STOCK_LEDGER WHERE VC_REASON='WPTEST_RECV'"))
    trig_off = int(scalar("SELECT COUNT(*) FROM sys.triggers WHERE is_disabled=1 AND name LIKE '%RecConfStat%'")) \
        + int(scalar("SELECT COUNT(*) FROM sys.triggers WHERE is_disabled=1 AND name='UPDATE_PartNumber'"))
    rep.check("DB restored as found (IN_QTY, no test rows/HIST/ledger, triggers re-enabled)",
              restored == base and oo == 0 and hist == 0 and led == 0 and trig_off == 0,
              "IN_QTY=%d oo=%d hist=%d led=%d trig_off=%d" % (restored, oo, hist, led, trig_off))
    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
