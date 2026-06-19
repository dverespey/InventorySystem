#!/usr/bin/env python3
"""test_seam_driver_order.py — drive the REAL `order.commitOrders` write path end-to-end through the
system.db shim's PERSISTENT-SESSION transaction extension (cutover-runbook §7, the last R8 gap).

The producer seams (test_seam_driver.py) are autocommit; Order's commitOrders is different — it opens a
transaction with `beginTransaction(db)` that SPANS statements (N× INSERT_OpenOrder + UPDATE_PartsStockRenban)
and rolls the whole thing back on any mid-loop failure (Order.pas:686/757). An autocommit shim (a fresh
`docker exec -Q` per call = a new connection each time) physically CANNOT carry a BEGIN TRAN across
statements, so it could never prove that atomicity. jython_shim's `_TxSession` holds ONE `docker exec -i`
connection open and feeds it framed batches, so the REAL commitOrders runs against a real transaction.

This file is a SIBLING of test_seam_driver.py (kept separate so the tx-session lifecycle doesn't entangle
the autocommit producers). It drives the ACTUAL commitOrders (not a SQL re-implementation):

  1. HAPPY PATH — a lot-sized 2-lot order (bit 0, no renban group) for a real part (16 / 478930223000).
     Asserts: exactly len(records) rows landed in INV_OPEN_ORDER_INF with the FRS#/renban/qty
     computeOrderRecords produced; IN_RENBAN_COUNT advanced to newCount via UPDATE_PartsStockRenban;
     IN_QTY did NOT move (order-create writes blank-status rows -> the receiving trigger's add-point gate
     does not fire); and it all COMMITTED (rows persist after closeTransaction).

  2. ROLLBACK PROOF — inject a real SQL error on the 2nd INSERT (a bad statement routed onto the SAME live
     tx session, exactly as a constraint violation would behave). Asserts: the first INSERT was VISIBLE
     INSIDE the open tx before the failure, yet NO order rows persist after and IN_RENBAN_COUNT is
     UNCHANGED — atomicity across statements, the thing the autocommit shim could never show.

Far-future FRS date (20291215 -> prefix '91215') keeps the test rows uniquely identifiable. DB restored
exactly as-found in a finally and asserted at the end, matching the other seam tests.

Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 scripts/e2e/test_seam_driver_order.py
"""
import os, subprocess, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report          # noqa: E402
import jython_shim              # noqa: E402

CONTAINER = os.environ.get("CONTAINER", "mssql-spike")
SA_PASS = os.environ.get("SA_PASS")
DB = "Inventory"
PART = 16
PARTNUM = "478930223000"        # part 16's VC_PART_NUMBER (same fixture as test_order_commit_integration)
SUPCODE = "WPT"
FRSD = "20291215"               # order-by date -> FRS prefix '91215' (yyyymmdd[3:8]); far-future = unique
FRSPREFIX = "91215"
ORDER_PY = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                            "docs", "analysis", "order", "project-library", "order", "code.py"))


def sql(query):
    if not SA_PASS:
        sys.exit("export SA_PASS first")
    out = subprocess.check_output([
        "docker", "exec", CONTAINER, "/opt/mssql-tools18/bin/sqlcmd", "-C", "-S", "localhost",
        "-U", "sa", "-P", SA_PASS, "-d", DB, "-h", "-1", "-W", "-s", "\t",
        "-Q", "SET NOCOUNT ON; " + query], text=True)
    return [l.split("\t") for l in out.splitlines()
            if l.strip() and not l.startswith("(") and not l.startswith("Msg ")]


def scalar(q):
    r = sql(q)
    return r[0][0] if r else None


def load_order():
    """The REAL order library with the shim's `system` (incl. the tx-capable db) injected."""
    return jython_shim.load_wrapper("order_real", ORDER_PY)


def cleanup(count0, base_qty):
    sql("DELETE FROM INV_OPEN_ORDER_INF WHERE VC_FRS_NUMBER LIKE '%s%%'; "
        "DELETE FROM INV_OPEN_ORDER_INF_HIST WHERE VC_FRS_NUMBER LIKE '%s%%'; "
        "UPDATE INV_PARTS_STOCK_MST SET IN_RENBAN_COUNT=%d, IN_QTY=%d WHERE IN_PART_ID=%d"
        % (FRSPREFIX, FRSPREFIX, count0, base_qty, PART))


def main():
    print("=" * 78)
    print(" SEAM DRIVER (Order) — REAL order.commitOrders via the shim's PERSISTENT-SESSION tx (R8 §7)")
    print("=" * 78)
    rep = Report()
    if not SA_PASS:
        sys.exit("export SA_PASS first")

    order = load_order()
    rep.check("real order module loads under the shim (commitOrders present)",
              hasattr(order, "commitOrders") and hasattr(order, "computeOrderRecords"))

    kanban = scalar("SELECT VC_KANBAN_NUMBER FROM INV_PARTS_STOCK_MST WHERE IN_PART_ID=%d" % PART)
    count0 = int(scalar("SELECT IN_RENBAN_COUNT FROM INV_PARTS_STOCK_MST WHERE IN_PART_ID=%d" % PART))
    base_qty = int(scalar("SELECT IN_QTY FROM INV_PARTS_STOCK_MST WHERE IN_PART_ID=%d" % PART))
    print("  part %d kanban=%s renban_count=%d IN_QTY=%d" % (PART, kanban, count0, base_qty))

    # lot-sized line (bit 0, no renban group): 2 lots -> 2 records + a counter bump => exercises the loop
    # AND the conditional UPDATE_PartsStockRenban inside ONE transaction.
    line = {"supCode": SUPCODE, "partNum": PARTNUM, "kanban": kanban, "lotSizedBit": 0,
            "renbanGroup": "", "lotCount": 2, "qtyQ": 500, "lotQty": 250, "frsDateYmd": FRSD}

    try:
        # ---- 1. HAPPY PATH: drive the REAL commitOrders through the tx session ----------------------
        print("\n--- happy path (real commitOrders -> begin/INSERT x2/UPDATE_renban/commit) ---")
        order.system = jython_shim._System()
        records = order.commitOrders(line, DB)
        rep.check("commitOrders returned 2 records (lot-sized 2 lots)", len(records) == 2,
                  str([r["frsNum"] for r in records]))

        rows = sql("SELECT VC_FRS_NUMBER, VC_RENBAN_NUMBER, IN_QTY FROM INV_OPEN_ORDER_INF "
                   "WHERE VC_FRS_NUMBER LIKE '%s%%' ORDER BY VC_FRS_NUMBER" % FRSPREFIX)
        got = [(r[0], r[1], int(r[2])) for r in rows]
        want = [(r["frsNum"], r["renbanNum"], r["qty"]) for r in records]
        rep.check("exactly len(records) rows COMMITTED with the computed FRS/renban/qty "
                  "(persist after closeTransaction)", got == want, "got=%s want=%s" % (got, want))
        rep.check("FRS sequential (9121501, 9121502)",
                  [g[0] for g in got] == ["9121501", "9121502"], str([g[0] for g in got]))
        rep.check("renban sequential (kanban + counter from %d via the tx)" % count0,
                  [g[1] for g in got] == ["%s%03d" % (kanban, count0), "%s%03d" % (kanban, count0 + 1)],
                  str([g[1] for g in got]))

        cnt_after = int(scalar("SELECT IN_RENBAN_COUNT FROM INV_PARTS_STOCK_MST WHERE IN_PART_ID=%d" % PART))
        rep.check("IN_RENBAN_COUNT advanced to newCount via UPDATE_PartsStockRenban (committed)",
                  cnt_after == count0 + 2, "%d -> %d (want %d)" % (count0, cnt_after, count0 + 2))

        qty_after = int(scalar("SELECT IN_QTY FROM INV_PARTS_STOCK_MST WHERE IN_PART_ID=%d" % PART))
        rep.check("IN_QTY did NOT move (Order is a write path, not a stock move)",
                  qty_after == base_qty, "%d -> %d" % (base_qty, qty_after))

        # reset to baseline before the rollback scenario (so its assertions start clean)
        cleanup(count0, base_qty)
        rep.check("baseline reset between scenarios (no rows, counter back to %d)" % count0,
                  len(sql("SELECT 1 FROM INV_OPEN_ORDER_INF WHERE VC_FRS_NUMBER LIKE '%s%%'" % FRSPREFIX)) == 0
                  and int(scalar("SELECT IN_RENBAN_COUNT FROM INV_PARTS_STOCK_MST WHERE IN_PART_ID=%d" % PART)) == count0)

        # ---- 2. ROLLBACK PROOF: mid-transaction failure must unwind the earlier INSERT --------------
        # Inject a REAL SQL error on the 2nd INSERT by routing a bad statement onto the SAME live tx
        # session (exactly how a constraint violation would surface), while still driving the ACTUAL
        # commitOrders begin/try/except-rollback/finally-close lifecycle. Before failing we peek that
        # the FIRST insert is visible INSIDE the open tx — so the post-rollback "0 rows" is genuinely a
        # rollback of a committed-into-tx statement, not a never-ran statement.
        print("\n--- rollback proof (real commitOrders; injected error on the 2nd INSERT, same tx) ---")
        sysobj = jython_shim._System()
        order.system = sysobj
        real_update = sysobj.db.runPrepUpdate
        state = {"n": 0, "seen_inside": None}

        def faulty_update(sqltext, args, db=None, getKey=False, tx=None):
            state["n"] += 1
            if state["n"] == 2:
                # the first INSERT already ran on this tx; confirm it's visible WITHIN the open tx
                _cols, _rows = tx._batch(
                    "SELECT COUNT(*) c FROM INV_OPEN_ORDER_INF WHERE VC_FRS_NUMBER LIKE '%s%%'" % FRSPREFIX,
                    want_rows=True)
                state["seen_inside"] = int(_rows[0][0])
                # a guaranteed real error on the same connection -> SqlError -> commitOrders rollback
                return tx.exec_update("INSERT INTO INV_OPEN_ORDER_INF (IN_QTY) VALUES ('not-an-int-zz')")
            return real_update(sqltext, args, db, getKey, tx)

        sysobj.db.runPrepUpdate = faulty_update

        threw = False
        try:
            order.commitOrders(line, DB)
        except Exception as e:          # noqa: E722 - commitOrders re-raises after rollback
            threw = True
            print("    commitOrders raised (expected): %s" % str(e)[:70])
        rep.check("commitOrders re-raised the mid-tx failure", threw)
        rep.check("the FIRST INSERT was visible INSIDE the open transaction before the failure",
                  state["seen_inside"] == 1, "count_in_tx=%r" % state["seen_inside"])

        persisted = len(sql("SELECT 1 FROM INV_OPEN_ORDER_INF WHERE VC_FRS_NUMBER LIKE '%s%%'" % FRSPREFIX))
        rep.check("NO order rows persisted — the earlier INSERT rolled back (cross-statement atomicity)",
                  persisted == 0, "persisted=%d" % persisted)
        cnt_rb = int(scalar("SELECT IN_RENBAN_COUNT FROM INV_PARTS_STOCK_MST WHERE IN_PART_ID=%d" % PART))
        rep.check("IN_RENBAN_COUNT UNCHANGED after rollback (no partial counter advance)",
                  cnt_rb == count0, "%d (want %d)" % (cnt_rb, count0))
    finally:
        cleanup(count0, base_qty)

    leftover = int(scalar("SELECT COUNT(*) FROM INV_OPEN_ORDER_INF WHERE VC_FRS_NUMBER LIKE '%s%%'" % FRSPREFIX))
    leftover_hist = int(scalar("SELECT COUNT(*) FROM INV_OPEN_ORDER_INF_HIST WHERE VC_FRS_NUMBER LIKE '%s%%'" % FRSPREFIX))
    cnt_restored = int(scalar("SELECT IN_RENBAN_COUNT FROM INV_PARTS_STOCK_MST WHERE IN_PART_ID=%d" % PART))
    qty_restored = int(scalar("SELECT IN_QTY FROM INV_PARTS_STOCK_MST WHERE IN_PART_ID=%d" % PART))
    rep.check("DB restored as found (no test orders/hist, counter + IN_QTY restored)",
              leftover == 0 and leftover_hist == 0 and cnt_restored == count0 and qty_restored == base_qty,
              "orders=%d hist=%d counter=%d qty=%d" % (leftover, leftover_hist, cnt_restored, qty_restored))
    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
