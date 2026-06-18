#!/usr/bin/env python3
"""test_seam_driver.py — run the REAL gateway-side Jython write-then-post wrappers end-to-end (retro R8).

Closes the standing coverage gap: the *_writepost.py integration tests proved the pure computePosts +
the POST_StockMovement proc, but hand-rolled the orchestration in SQL — they never executed the actual
Jython driver functions (insertAdjustment / amendAdjustment / deleteAdjustment, the dynamic SQL, getKey
round-trip, _readRow re-read, the stockLedger.post funnel). This test loads those REAL modules with a
`system.db` shim (jython_shim.py) routing their DB calls to the spike DB, disables the legacy triggers
(so only the wrappers' POST_StockMovement moves IN_QTY), and drives insert/amend/delete through the
ACTUAL functions — asserting IN_QTY + the ledger. Fixture-disciplined.

Covers the transaction-free producers (stocktaking, reject) — the proven runner mechanism. Shipping /
receiving extend mechanically (same shim); Order's commitOrders uses beginTransaction and needs the
shim's persistent-session extension (noted in jython_shim.py).

Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 scripts/e2e/test_seam_driver.py
"""
import os, subprocess, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report          # noqa: E402
import jython_shim              # noqa: E402

CONTAINER = os.environ.get("CONTAINER", "mssql-spike")
SA_PASS = os.environ.get("SA_PASS")
DB = "Inventory"
PART = 16
REASON = "SHIMTEST"
PLIB = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                                     "docs", "analysis", "inventory-stock", "project-library"))


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
    """The REAL stockLedger + stocktaking + reject wrappers, with the shim's `system` injected and
    the real stockLedger threaded into the producer modules."""
    sl = jython_shim.load_wrapper("sl_real", os.path.join(PLIB, "stockLedger", "code.py"))
    stk = jython_shim.load_wrapper("stk_real", os.path.join(PLIB, "stocktaking", "code.py"),
                                   extra_globals={"stockLedger": sl})
    rej = jython_shim.load_wrapper("rej_real", os.path.join(PLIB, "reject", "code.py"),
                                   extra_globals={"stockLedger": sl})
    return stk, rej


LEDGER_FILTER = "IN_PART_ID=%d AND VC_SOURCE_ENUM IN ('STOCKTAKING','REJECT')" % PART  # ledger rows carry
# computePosts's own VC_REASON (not our SHIMTEST marker), so key the ledger on part+enum (part 16 has no
# production ledger rows in the spike — baseline empty).


def cleanup(base):
    sql(("DELETE FROM INV_STOCK_LEDGER WHERE " + LEDGER_FILTER + "; "
         "DELETE FROM INV_STOCKTAKING_INF WHERE VC_REASON='%s'; "
         "DELETE FROM INV_REJECT_INF WHERE VC_REASON='%s'; "
         "UPDATE INV_PARTS_STOCK_MST SET IN_QTY=%d WHERE IN_PART_ID=%d; "
         "ENABLE TRIGGER INSERT_Stocktaking ON INV_STOCKTAKING_INF; "
         "ENABLE TRIGGER UPDATE_Stocktaking ON INV_STOCKTAKING_INF; "
         "ENABLE TRIGGER DELETE_Stocktaking ON INV_STOCKTAKING_INF; "
         "ENABLE TRIGGER INSERT_RejectParts ON INV_REJECT_INF; "
         "ENABLE TRIGGER UPDATE_RejectParts ON INV_REJECT_INF; "
         "ENABLE TRIGGER DELETE_RejectParts ON INV_REJECT_INF; "
         "ENABLE TRIGGER UPDATE_PartNumber ON INV_PARTS_STOCK_MST")
        % (REASON, REASON, base, PART))


def main():
    print("=" * 78)
    print(" SEAM DRIVER — running the REAL Jython write-then-post wrappers via the system.db shim (R8)")
    print("=" * 78)
    rep = Report()
    if not SA_PASS:
        sys.exit("export SA_PASS first")
    stk, rej = load_real_modules()
    rep.check("real wrapper modules load under the shim (stocktaking + reject)",
              hasattr(stk, "insertAdjustment") and hasattr(rej, "insertReject"))
    base = qty()
    print("  part %d base IN_QTY = %d" % (PART, base))
    sql("DISABLE TRIGGER INSERT_Stocktaking ON INV_STOCKTAKING_INF; "
        "DISABLE TRIGGER UPDATE_Stocktaking ON INV_STOCKTAKING_INF; "
        "DISABLE TRIGGER DELETE_Stocktaking ON INV_STOCKTAKING_INF; "
        "DISABLE TRIGGER INSERT_RejectParts ON INV_REJECT_INF; "
        "DISABLE TRIGGER UPDATE_RejectParts ON INV_REJECT_INF; "
        "DISABLE TRIGGER DELETE_RejectParts ON INV_REJECT_INF; "
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
    finally:
        cleanup(base)

    restored = qty()
    leftover = int(scalar("SELECT COUNT(*) FROM INV_STOCK_LEDGER WHERE " + LEDGER_FILTER) or 0)
    trig = int(scalar("SELECT COUNT(*) FROM sys.triggers WHERE is_disabled=1 AND name IN "
                      "('INSERT_Stocktaking','UPDATE_Stocktaking','DELETE_Stocktaking','INSERT_RejectParts',"
                      "'UPDATE_RejectParts','DELETE_RejectParts','UPDATE_PartNumber')") or 0)
    rep.check("DB restored as found (IN_QTY, no test rows/ledger, triggers re-enabled)",
              restored == base and leftover == 0 and trig == 0,
              "IN_QTY=%d ledger=%d trig_off=%d" % (restored, leftover, trig))
    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
