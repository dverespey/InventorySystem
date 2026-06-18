#!/usr/bin/env python3
"""test_order_commit_integration.py — the Order commit path end-to-end against the live spike DB.

Drives the REAL order.computeOrderRecords output through the legacy INSERT_OpenOrder proc and the
renban counter (SELECT/UPDATE_PartsStockRenban), then verifies the written INV_OPEN_ORDER_INF rows
(FRS#, renban, qty) match what computeOrderRecords produced, the counter advanced correctly, and — since
order-create writes BLANK-status orders ("to order", not shipped) — on-hand IN_QTY did NOT move (the
receiving INSERT trigger's add-point gate doesn't fire on a blank-status row). Fixture-disciplined.

Uses a far-future FRS date (20291215 -> prefix '91215') so the test rows are uniquely identifiable and
collide with nothing real. Lot-sized scenario (bit 0, 2 lots, no renban group) to exercise the
multi-record + sequential-renban + counter-bump path.

Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 scripts/e2e/test_order_commit_integration.py
"""
import os, subprocess, sys, importlib.util

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report  # noqa: E402

CONTAINER = os.environ.get("CONTAINER", "mssql-spike")
SA_PASS = os.environ.get("SA_PASS")
DB = "Inventory"
PART = 16
PARTNUM = "478930223000"
FRSD = "20291215"               # -> FRS prefix '91215'
FRSPREFIX = "91215"
ORDER_PY = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                            "docs", "analysis", "order", "project-library", "order", "code.py"))


def _load():
    spec = importlib.util.spec_from_file_location("order_lib", ORDER_PY)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


order = _load()


def sql(query):
    if not SA_PASS:
        sys.exit("export SA_PASS first")
    out = subprocess.check_output([
        "docker", "exec", CONTAINER, "/opt/mssql-tools18/bin/sqlcmd", "-C", "-S", "localhost",
        "-U", "sa", "-P", SA_PASS, "-d", DB, "-h", "-1", "-W", "-s", "\t",
        "-Q", "SET NOCOUNT ON; " + query], text=True)
    return [l.split("\t") for l in out.splitlines()
            if l.strip() and not l.startswith("(") and not l.startswith("Msg ")]


def scalar(query):
    r = sql(query)
    return r[0][0] if r else None


def q(s):
    return "'" + str(s).replace("'", "''") + "'"


def cleanup(count0):
    sql("DELETE FROM INV_OPEN_ORDER_INF WHERE VC_FRS_NUMBER LIKE '%s%%'; "
        "DELETE FROM INV_OPEN_ORDER_INF_HIST WHERE VC_FRS_NUMBER LIKE '%s%%'; "
        "UPDATE INV_PARTS_STOCK_MST SET IN_RENBAN_COUNT=%s WHERE IN_PART_ID=%d"
        % (FRSPREFIX, FRSPREFIX, count0, PART))


def main():
    print("=" * 78)
    print(" ORDER commit/write path — integration (computeOrderRecords -> INSERT_OpenOrder, live DB)")
    print("=" * 78)
    rep = Report()
    kanban = scalar("SELECT VC_KANBAN_NUMBER FROM INV_PARTS_STOCK_MST WHERE IN_PART_ID=%d" % PART)
    count0 = int(scalar("SELECT IN_RENBAN_COUNT FROM INV_PARTS_STOCK_MST WHERE IN_PART_ID=%d" % PART))
    base_qty = int(scalar("SELECT IN_QTY FROM INV_PARTS_STOCK_MST WHERE IN_PART_ID=%d" % PART))
    print("  part %d kanban=%s renban_count=%d IN_QTY=%d" % (PART, kanban, count0, base_qty))

    # compute the REAL records for a lot-sized 2-lot order (bit 0, no renban group)
    recs, newCount = order.computeOrderRecords("WPT", PARTNUM, kanban, 0, "", lotCount=2,
                                               qtyQ=500, lotQty=250, frsDateYmd=FRSD, renbanCount=count0)
    try:
        # write each via the legacy proc + persist the advanced counter (mirrors commitOrders' DB steps)
        for r in recs:
            sql("EXEC INSERT_OpenOrder @SupCode=%s, @PartNum=%s, @KanbanNum=%s, @FRSNum=%s, "
                "@RenbanNum=%s, @Qty=%d" % (q(r["supCode"]), q(r["partNum"]), q(r["kanban"]),
                                            q(r["frsNum"]), q(r["renbanNum"]), r["qty"]))
        sql("EXEC UPDATE_PartsStockRenban @PartNum=%s, @RenbanCount=%d" % (q(PARTNUM), newCount))

        rows = sql("SELECT VC_FRS_NUMBER, VC_RENBAN_NUMBER, IN_QTY FROM INV_OPEN_ORDER_INF "
                   "WHERE VC_FRS_NUMBER LIKE '%s%%' ORDER BY VC_FRS_NUMBER" % FRSPREFIX)
        got = [(r[0], r[1], int(r[2])) for r in rows]
        want = [(r["frsNum"], r["renbanNum"], r["qty"]) for r in recs]
        rep.check("both order rows written with the computed FRS/renban/qty", got == want,
                  "got=%s want=%s" % (got, want))
        rep.check("FRS sequential (9121501, 9121502)",
                  [g[0] for g in got] == ["9121501", "9121502"], str([g[0] for g in got]))
        rep.check("renban sequential (kanban + counter from %d)" % count0,
                  [g[1] for g in got] == ["%s%03d" % (kanban, count0), "%s%03d" % (kanban, count0 + 1)],
                  str([g[1] for g in got]))
        cnt_after = int(scalar("SELECT IN_RENBAN_COUNT FROM INV_PARTS_STOCK_MST WHERE IN_PART_ID=%d" % PART))
        rep.check("renban counter advanced by 2 in the DB", cnt_after == count0 + 2,
                  "%d -> %d" % (count0, cnt_after))
        qty_after = int(scalar("SELECT IN_QTY FROM INV_PARTS_STOCK_MST WHERE IN_PART_ID=%d" % PART))
        rep.check("on-hand IN_QTY UNCHANGED (order-create writes blank-status rows -> no stock move)",
                  qty_after == base_qty, "%d -> %d" % (base_qty, qty_after))
    finally:
        cleanup(count0)

    leftover = int(scalar("SELECT COUNT(*) FROM INV_OPEN_ORDER_INF WHERE VC_FRS_NUMBER LIKE '%s%%'" % FRSPREFIX))
    cnt_restored = int(scalar("SELECT IN_RENBAN_COUNT FROM INV_PARTS_STOCK_MST WHERE IN_PART_ID=%d" % PART))
    rep.check("DB restored as found (no test orders, counter restored)",
              leftover == 0 and cnt_restored == count0, "leftover=%d count=%d" % (leftover, cnt_restored))
    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
