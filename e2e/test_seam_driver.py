#!/usr/bin/env python3
"""test_seam_driver.py — run the REAL gateway-side Jython write-then-post wrappers end-to-end (retro R8).

Closes the standing coverage gap: the *_writepost.py integration tests proved the pure computePosts +
the POST_StockMovement proc, but hand-rolled the orchestration in SQL — they never executed the actual
Jython driver functions (insertAdjustment / amendAdjustment / deleteAdjustment, the dynamic SQL, getKey
round-trip, _readRow re-read, the stockLedger.post funnel). This test loads those REAL modules with a
`system.db` shim (jython_shim.py) routing their DB calls to the spike DB, disables the legacy triggers
(so only the wrappers' POST_StockMovement moves IN_QTY), and drives insert/amend/delete through the
ACTUAL functions — asserting IN_QTY + the ledger. Fixture-disciplined.

Covers the transaction-free producers (stocktaking, reject, shipping, receiving) — the proven runner
mechanism. Each section disables the legacy qty triggers for that table so only the wrappers'
POST_StockMovement moves IN_QTY, drives the REAL driver functions (insert/amend/delete + shipping's
header-delete cascade and receiving's resolveAddPoint add-point gate), asserts IN_QTY + the ledger, and
restores the DB as found. Order's commitOrders uses beginTransaction and needs the shim's
persistent-session extension (still a documented extension — noted in jython_shim.py).

Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 e2e/test_seam_driver.py
"""
import os, subprocess, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report          # noqa: E402
import jython_shim              # noqa: E402

CONTAINER = os.environ.get("CONTAINER", "mssql-spike")
SA_PASS = os.environ.get("SA_PASS")
DB = "Inventory"
PART = 16
PARTNUM = "478930223000"        # part 16's VC_PART_NUMBER; add-point 'S' (verified against the live join)
REASON = "SHIMTEST"
SHIP_DATE = "29991231"          # distinctive VC_PRODUCTION_DATE for the shipping part-line fixture
SHIP_LINE = "SHIMSHIPLINE"      # distinctive VC_LINE_NAME for the INV_SHIPPING_INF header fixture
RECV_RENBAN = "SHIMRC01"        # distinctive VC_RENBAN_NUMBER for the open-order fixture (col max 8)
RECV_SHIP_STAMP = "20260618"    # a non-blank VC_STATUS_SUPPLIER_SHIPPING => 'S' add-point counts (moves IN_QTY)
PLIB = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..",
                                     "ignition", "project-library"))


def sql(query):
    out = subprocess.check_output([
        "docker", "exec", CONTAINER, "/opt/mssql-tools18/bin/sqlcmd", "-C", "-S", "localhost",
        "-U", "sa", "-P", SA_PASS, "-d", DB, "-h", "-1", "-W", "-s", "\t",
        "-Q", "SET NOCOUNT ON; " + query], text=True)
    return [l.split("\t") for l in out.splitlines()
            if l.strip() and not l.startswith("(") and not l.startswith("Msg ")]


def scalar(q):
    r = sql(q)
    return r[0][0] if r else None


def qty():
    return int(scalar("SELECT IN_QTY FROM INV_PARTS_STOCK_MST WHERE IN_PART_ID=%d" % PART))


def load_real_modules():
    """The REAL stockLedger + stocktaking + reject + shipping + receiving wrappers, with the shim's
    `system` injected and the real stockLedger threaded into the producer modules."""
    sl = jython_shim.load_wrapper("sl_real", os.path.join(PLIB, "stockLedger", "code.py"))
    stk = jython_shim.load_wrapper("stk_real", os.path.join(PLIB, "stocktaking", "code.py"),
                                   extra_globals={"stockLedger": sl})
    rej = jython_shim.load_wrapper("rej_real", os.path.join(PLIB, "reject", "code.py"),
                                   extra_globals={"stockLedger": sl})
    shp = jython_shim.load_wrapper("shp_real", os.path.join(PLIB, "shipping", "code.py"),
                                   extra_globals={"stockLedger": sl})
    rcv = jython_shim.load_wrapper("rcv_real", os.path.join(PLIB, "receiving", "code.py"),
                                   extra_globals={"stockLedger": sl})
    return stk, rej, shp, rcv


LEDGER_FILTER = "IN_PART_ID=%d AND VC_SOURCE_ENUM IN ('STOCKTAKING','REJECT')" % PART  # ledger rows carry
# computePosts's own VC_REASON (not our SHIMTEST marker), so key the ledger on part+enum (part 16 has no
# production ledger rows in the spike — baseline empty).

# Shipping/receiving drivers write their OWN reason strings ("shipping: ...", "receiving ...") not our
# markers, so — exactly like the stocktaking/reject filter above — key those ledgers on part+enum.
# Verified baseline: part 16 has ZERO rows in any of these enums (clean teardown target).
SHIP_LEDGER_FILTER = "IN_PART_ID=%d AND VC_SOURCE_ENUM='SHIPPING'" % PART
RECV_LEDGER_FILTER = ("IN_PART_ID=%d AND VC_SOURCE_ENUM IN "
                      "('RECEIVING_SHIP','RECEIVING_ARRIVAL','RECEIVING_REVERSAL')") % PART


def make_ship_header():
    """Create a valid INV_SHIPPING_INF header (the FK target INV_PART_SHIPPING_INF requires) and return
    its IN_SHIPPING_ID. Mirrors test_shipping_writepost.ins_header."""
    return int(scalar(
        "INSERT INTO INV_SHIPPING_INF (VC_LINE_NAME, VC_START_SEQ_NUMBER, VC_END_SEQ_NUMBER, "
        "VC_PRODUCTION_DATE, VC_ADD) OUTPUT INSERTED.IN_SHIPPING_ID VALUES "
        "('%s', '0', '0', '%s', CONVERT(varchar,getdate(),112))" % (SHIP_LINE, SHIP_DATE)))


def cleanup(base):
    sql(("DELETE FROM INV_STOCK_LEDGER WHERE " + LEDGER_FILTER + "; "
         "DELETE FROM INV_STOCK_LEDGER WHERE " + SHIP_LEDGER_FILTER + "; "
         "DELETE FROM INV_STOCK_LEDGER WHERE " + RECV_LEDGER_FILTER + "; "
         "DELETE FROM INV_STOCKTAKING_INF WHERE VC_REASON='%s'; "
         "DELETE FROM INV_REJECT_INF WHERE VC_REASON='%s'; "
         "DELETE FROM INV_PART_SHIPPING_INF WHERE VC_PRODUCTION_DATE='%s'; "
         "DELETE FROM INV_SHIPPING_INF WHERE VC_LINE_NAME='%s'; "
         "DELETE FROM INV_OPEN_ORDER_INF WHERE VC_RENBAN_NUMBER='%s'; "
         "DELETE FROM INV_OPEN_ORDER_INF_HIST WHERE VC_RENBAN_NUMBER='%s'; "
         "UPDATE INV_PARTS_STOCK_MST SET IN_QTY=%d WHERE IN_PART_ID=%d; "
         "ENABLE TRIGGER INSERT_Stocktaking ON INV_STOCKTAKING_INF; "
         "ENABLE TRIGGER UPDATE_Stocktaking ON INV_STOCKTAKING_INF; "
         "ENABLE TRIGGER DELETE_Stocktaking ON INV_STOCKTAKING_INF; "
         "ENABLE TRIGGER INSERT_RejectParts ON INV_REJECT_INF; "
         "ENABLE TRIGGER UPDATE_RejectParts ON INV_REJECT_INF; "
         "ENABLE TRIGGER DELETE_RejectParts ON INV_REJECT_INF; "
         "ENABLE TRIGGER InsertPartShipping ON INV_PART_SHIPPING_INF; "
         "ENABLE TRIGGER UpdatePartShipping ON INV_PART_SHIPPING_INF; "
         "ENABLE TRIGGER DeletePartShipping ON INV_PART_SHIPPING_INF; "
         "ENABLE TRIGGER DeleteShipDate ON INV_SHIPPING_INF; "
         "ENABLE TRIGGER INSERT_RecConfStatPartsStockMstQTY ON INV_OPEN_ORDER_INF; "
         "ENABLE TRIGGER UPDATE_RecConfStatPartsStockMstQTY ON INV_OPEN_ORDER_INF; "
         "ENABLE TRIGGER DELETE_RecConfStatPartsStockMstQTY ON INV_OPEN_ORDER_INF; "
         "ENABLE TRIGGER UPDATE_PartNumber ON INV_PARTS_STOCK_MST")
        % (REASON, REASON, SHIP_DATE, SHIP_LINE, RECV_RENBAN, RECV_RENBAN, base, PART))


def main():
    print("=" * 78)
    print(" SEAM DRIVER — running the REAL Jython write-then-post wrappers via the system.db shim (R8)")
    print("=" * 78)
    rep = Report()
    if not SA_PASS:
        sys.exit("export SA_PASS first")
    stk, rej, shp, rcv = load_real_modules()
    rep.check("real wrapper modules load under the shim (stocktaking + reject + shipping + receiving)",
              hasattr(stk, "insertAdjustment") and hasattr(rej, "insertReject")
              and hasattr(shp, "insertPartShipping") and hasattr(rcv, "insertOpenOrder"))
    base = qty()
    print("  part %d (%s, add-point 'S') base IN_QTY = %d" % (PART, PARTNUM, base))
    sql("DISABLE TRIGGER INSERT_Stocktaking ON INV_STOCKTAKING_INF; "
        "DISABLE TRIGGER UPDATE_Stocktaking ON INV_STOCKTAKING_INF; "
        "DISABLE TRIGGER DELETE_Stocktaking ON INV_STOCKTAKING_INF; "
        "DISABLE TRIGGER INSERT_RejectParts ON INV_REJECT_INF; "
        "DISABLE TRIGGER UPDATE_RejectParts ON INV_REJECT_INF; "
        "DISABLE TRIGGER DELETE_RejectParts ON INV_REJECT_INF; "
        "DISABLE TRIGGER InsertPartShipping ON INV_PART_SHIPPING_INF; "
        "DISABLE TRIGGER UpdatePartShipping ON INV_PART_SHIPPING_INF; "
        "DISABLE TRIGGER DeletePartShipping ON INV_PART_SHIPPING_INF; "
        "DISABLE TRIGGER DeleteShipDate ON INV_SHIPPING_INF; "
        "DISABLE TRIGGER INSERT_RecConfStatPartsStockMstQTY ON INV_OPEN_ORDER_INF; "
        "DISABLE TRIGGER UPDATE_RecConfStatPartsStockMstQTY ON INV_OPEN_ORDER_INF; "
        "DISABLE TRIGGER DELETE_RecConfStatPartsStockMstQTY ON INV_OPEN_ORDER_INF; "
        "DISABLE TRIGGER UPDATE_PartNumber ON INV_PARTS_STOCK_MST")
    try:
        # ---- STOCKTAKING: drive the REAL insert/amend/delete ----
        print("\n--- stocktaking (real insertAdjustment/amendAdjustment/deleteAdjustment) ---")
        sid = stk.insertAdjustment(PART, 50, REASON)
        rep.check("real insertAdjustment +50 -> IN_QTY base+50 + ledger row", qty() == base + 50 and sid,
                  "IN_QTY=%d id=%s" % (qty(), sid))
        led = int(scalar("SELECT IN_QTY_CHANGE FROM INV_STOCK_LEDGER "
                         "WHERE VC_SOURCE_EVENT='STOCKTAKING:stk=%s:ins'" % sid) or 0)
        rep.check("real insert wrote the ledger movement (+50)", led == 50, "ledger=%d" % led)
        stk.amendAdjustment(sid, 80, REASON)
        rep.check("real amendAdjustment 50->80 -> IN_QTY base+80 (net +30 posted)", qty() == base + 80,
                  "IN_QTY=%d" % qty())
        stk.deleteAdjustment(sid)
        rep.check("real deleteAdjustment -> IN_QTY back to base", qty() == base, "IN_QTY=%d" % qty())

        # ---- REJECT: drive the REAL insert/amend/delete (effect=-IN_QTY, int-keyed) ----
        print("\n--- reject (real insertReject/amendReject/deleteReject) ---")
        rid = rej.insertReject(PART, 30, "T", REASON)
        rep.check("real insertReject 30 -> IN_QTY base-30", qty() == base - 30, "IN_QTY=%d id=%s" % (qty(), rid))
        rej.amendReject(rid, 45, REASON)
        rep.check("real amendReject 30->45 -> IN_QTY base-45", qty() == base - 45, "IN_QTY=%d" % qty())
        rej.deleteReject(rid)
        rep.check("real deleteReject (un-reject) -> IN_QTY back to base", qty() == base, "IN_QTY=%d" % qty())

        # the whole point: these moved IN_QTY through the ACTUAL wrapper code, not a SQL re-implementation
        rep.check("ledger has the real-driver movements and nets to 0 (insert+amend+delete x2)",
                  int(scalar("SELECT COUNT(*) FROM INV_STOCK_LEDGER WHERE " + LEDGER_FILTER) or 0) >= 6 and
                  int(scalar("SELECT ISNULL(SUM(IN_QTY_CHANGE),0) FROM INV_STOCK_LEDGER WHERE " + LEDGER_FILTER) or 0) == 0)

        # ---- SHIPPING: drive the REAL insertPartShipping/amendPartShipping/deletePartShipping +
        #      deleteShipmentHeader (effect = -IN_QTY; surrogate IN_PART_SHIPPING_ID; FK header required) ----
        print("\n--- shipping (real insertPartShipping/amendPartShipping/deletePartShipping/"
              "deleteShipmentHeader) ---")
        shipId = make_ship_header()
        # insert a consumed line -> -50
        psid = shp.insertPartShipping(shipId, PARTNUM, SHIP_DATE, 50)
        rep.check("real insertPartShipping 50 -> IN_QTY base-50 + ledger row",
                  qty() == base - 50 and psid, "IN_QTY=%d id=%s" % (qty(), psid))
        sled = int(scalar("SELECT IN_QTY_CHANGE FROM INV_STOCK_LEDGER "
                          "WHERE VC_SOURCE_EVENT='SHIPPING:psh=%s:ins'" % psid) or 0)
        rep.check("real insert wrote the ledger movement (-50)", sled == -50, "ledger=%d" % sled)
        # amend qty 50->80 (consume more) -> net -30 more -> base-80
        shp.amendPartShipping(psid, 80)
        rep.check("real amendPartShipping 50->80 -> IN_QTY base-80 (net -30 posted)",
                  qty() == base - 80, "IN_QTY=%d" % qty())
        # delete the line -> restore +80 -> base
        shp.deletePartShipping(psid)
        rep.check("real deletePartShipping -> IN_QTY back to base", qty() == base, "IN_QTY=%d" % qty())

        # header-delete cascade: 2 same-part rows on one header, deleteShipmentHeader restores BOTH
        shipId2 = make_ship_header()
        pa = shp.insertPartShipping(shipId2, PARTNUM, SHIP_DATE, 30)
        pb = shp.insertPartShipping(shipId2, PARTNUM, SHIP_DATE, 20)
        rep.check("real header cascade: 2 lines consumed -> IN_QTY base-50",
                  qty() == base - 50, "IN_QTY=%d ids=%s,%s" % (qty(), pa, pb))
        shp.deleteShipmentHeader(shipId2)
        rep.check("real deleteShipmentHeader restores BOTH lines (multi-row-safe) -> base",
                  qty() == base, "IN_QTY=%d" % qty())
        rep.check("shipping ledger has the real-driver movements and nets to 0",
                  int(scalar("SELECT COUNT(*) FROM INV_STOCK_LEDGER WHERE " + SHIP_LEDGER_FILTER) or 0) >= 6 and
                  int(scalar("SELECT ISNULL(SUM(IN_QTY_CHANGE),0) FROM INV_STOCK_LEDGER WHERE " + SHIP_LEDGER_FILTER) or 0) == 0)

        # ---- RECEIVING: drive the REAL insertOpenOrder/updateOpenOrder/deleteOpenOrder. The 'S'
        #      add-point gate is resolved BY THE REAL resolveAddPoint join; the row is SHIPPED so it
        #      actually COUNTS (+IN_QTY) — a vacuous (non-counting) order would move nothing. ----
        print("\n--- receiving (real resolveAddPoint + insertOpenOrder/updateOpenOrder/deleteOpenOrder) ---")
        ap = rcv.resolveAddPoint(PARTNUM)
        rep.check("real resolveAddPoint(%s) -> 'S' (so a shipped order counts)" % PARTNUM,
                  ap == "S", "add-point=%r" % ap)
        oo = {"VC_SUPPLIER_CODE": "SHM", "VC_PART_NUMBER": PARTNUM, "VC_FRS_NUMBER": "SHM0001",
              "VC_RENBAN_NUMBER": RECV_RENBAN, "IN_QTY": 60,
              "VC_STATUS_SUPPLIER_SHIPPING": RECV_SHIP_STAMP}   # shipped => 'S' counts => +qty
        oid = rcv.insertOpenOrder(oo)
        rep.check("real insertOpenOrder (shipped 'S', counted) -> IN_QTY base+60 (NON-zero delta)",
                  qty() == base + 60 and oid, "IN_QTY=%d id=%s" % (qty(), oid))
        rled = int(scalar("SELECT IN_QTY_CHANGE FROM INV_STOCK_LEDGER "
                          "WHERE VC_SOURCE_EVENT='RECEIVING_SHIP:ord=%s:ins'" % oid) or 0)
        rep.check("real insert wrote the ledger movement (+60)", rled == 60, "ledger=%d" % rled)
        # amend qty 60->90 -> net +30 -> base+90
        rcv.updateOpenOrder(oid, {"IN_QTY": 90})
        rep.check("real updateOpenOrder 60->90 -> IN_QTY base+90 (net +30 posted)",
                  qty() == base + 90, "IN_QTY=%d" % qty())
        # delete the (still-active, counted) order -> reverse -90 -> base
        rcv.deleteOpenOrder(oid)
        rep.check("real deleteOpenOrder (active counted) -> IN_QTY back to base",
                  qty() == base, "IN_QTY=%d" % qty())
        rep.check("receiving ledger has the real-driver movements and nets to 0",
                  int(scalar("SELECT COUNT(*) FROM INV_STOCK_LEDGER WHERE " + RECV_LEDGER_FILTER) or 0) >= 3 and
                  int(scalar("SELECT ISNULL(SUM(IN_QTY_CHANGE),0) FROM INV_STOCK_LEDGER WHERE " + RECV_LEDGER_FILTER) or 0) == 0)
    finally:
        cleanup(base)

    restored = qty()
    leftover = int(scalar("SELECT COUNT(*) FROM INV_STOCK_LEDGER WHERE (" + LEDGER_FILTER + ") OR ("
                          + SHIP_LEDGER_FILTER + ") OR (" + RECV_LEDGER_FILTER + ")") or 0)
    fixtures = (int(scalar("SELECT COUNT(*) FROM INV_PART_SHIPPING_INF WHERE VC_PRODUCTION_DATE='%s'" % SHIP_DATE) or 0)
                + int(scalar("SELECT COUNT(*) FROM INV_SHIPPING_INF WHERE VC_LINE_NAME='%s'" % SHIP_LINE) or 0)
                + int(scalar("SELECT COUNT(*) FROM INV_OPEN_ORDER_INF WHERE VC_RENBAN_NUMBER='%s'" % RECV_RENBAN) or 0)
                + int(scalar("SELECT COUNT(*) FROM INV_OPEN_ORDER_INF_HIST WHERE VC_RENBAN_NUMBER='%s'" % RECV_RENBAN) or 0))
    trig = int(scalar("SELECT COUNT(*) FROM sys.triggers WHERE is_disabled=1 AND name IN "
                      "('INSERT_Stocktaking','UPDATE_Stocktaking','DELETE_Stocktaking','INSERT_RejectParts',"
                      "'UPDATE_RejectParts','DELETE_RejectParts','InsertPartShipping','UpdatePartShipping',"
                      "'DeletePartShipping','DeleteShipDate','INSERT_RecConfStatPartsStockMstQTY',"
                      "'UPDATE_RecConfStatPartsStockMstQTY','DELETE_RecConfStatPartsStockMstQTY',"
                      "'UPDATE_PartNumber')") or 0)
    rep.check("DB restored as found (IN_QTY, no test rows/ledger/fixtures, all triggers re-enabled)",
              restored == base and leftover == 0 and fixtures == 0 and trig == 0,
              "IN_QTY=%d ledger=%d fixtures=%d trig_off=%d" % (restored, leftover, fixtures, trig))
    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
