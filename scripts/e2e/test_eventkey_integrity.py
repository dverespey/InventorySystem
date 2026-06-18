#!/usr/bin/env python3
"""test_eventkey_integrity.py — the REAL computePosts event keys must fit VC_SOURCE_EVENT and
round-trip through POST_StockMovement without truncation.

Closes the harness blind-spot the producer-seam reviews caught: the per-producer write tests hand-roll
SHORT event keys for their parity arithmetic, so none of them exercised the ACTUAL key `computePosts`
emits. The amend keys (':upd:to=<effect>:v=<16-char stamp>', longest ENUM RECEIVING_REVERSAL) run
~45-64 chars and silently truncated in the original VC_SOURCE_EVENT varchar(40) / @eventKey varchar(40)
— which dropped the per-edit stamp and re-introduced the very amend-collision the ':to=' guard prevents.
The column + param were widened to varchar(100); this test imports each producer's real computePosts,
takes the longest amend/F5 keys, asserts len<=100, and posts each (delta=0 -> no IN_QTY change) then
reads VC_SOURCE_EVENT back and asserts it stored INTACT. Fixture-disciplined (deletes its ledger rows).

Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 scripts/e2e/test_eventkey_integrity.py
"""
import os, subprocess, sys, importlib.util

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report  # noqa: E402

CONTAINER = os.environ.get("CONTAINER", "mssql-spike")
SA_PASS = os.environ.get("SA_PASS")
DB = "Inventory"
PART = 16
REASON = "EVTKEY_TEST"
COLW = 100
PLIB = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                                     "docs", "analysis", "inventory-stock", "project-library"))
STAMP = "20260618235959ff"     # a full 16-char stamp (worst case for length)


def _load(name, rel):
    spec = importlib.util.spec_from_file_location(name, os.path.join(PLIB, rel, "code.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


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


def collect_keys():
    """The full set of REAL event keys the four producers emit (longest forms: amend + F5 + reversal)."""
    reject = _load("reject_lib", "reject")
    stk = _load("stk_lib", "stocktaking")
    shipping = _load("ship_lib", "shipping")
    receiving = _load("recv_lib", "receiving")
    keys = []
    # reject amend (int-keyed)
    keys += [p["eventKey"] for p in reject.computePosts(
        "update", {"IN_PART_ID": PART, "IN_QTY": 100, "IN_REJECT_ID": 9999999, "VC_LAST_UPDATE": STAMP},
        {"IN_PART_ID": PART, "IN_QTY": 140, "IN_REJECT_ID": 9999999, "VC_LAST_UPDATE": STAMP})]
    # stocktaking amend
    keys += [p["eventKey"] for p in stk.computePosts(
        "update", {"IN_PART_ID": PART, "IN_QTY": 100, "IN_STOCKTAKING_ID": 9999999, "VC_LAST_UPDATE": STAMP},
        {"IN_PART_ID": PART, "IN_QTY": 140, "IN_STOCKTAKING_ID": 9999999, "VC_LAST_UPDATE": STAMP})]
    # shipping amend + F5 part-change (two longest keys: del-old / add-new)
    keys += [p["eventKey"] for p in shipping.computePosts(
        "update", {"VC_PART_NUMBER": "P", "IN_QTY": 100, "IN_PART_SHIPPING_ID": 9999999, "VC_ADD": STAMP},
        {"VC_PART_NUMBER": "P", "IN_QTY": 140, "IN_PART_SHIPPING_ID": 9999999, "VC_ADD": STAMP})]
    keys += [p["eventKey"] for p in shipping.computePosts(
        "update", {"VC_PART_NUMBER": "OLDP", "IN_QTY": 100, "IN_PART_SHIPPING_ID": 9999999, "VC_ADD": STAMP},
        {"VC_PART_NUMBER": "NEWP", "IN_QTY": 140, "IN_PART_SHIPPING_ID": 9999999, "VC_ADD": STAMP})]
    # receiving amend (ship) + the longest-ENUM RECEIVING_REVERSAL ('A' arrival set->cleared) + F5
    oo = lambda **k: dict({"VC_PART_NUMBER": "P", "IN_QTY": 100, "IN_ORDER_ID": 9999999,
                           "VC_LAST_UPDATE": STAMP, "VC_STATUS_SUPPLIER_SHIPPING": "", "VC_ARRIVAL": "",
                           "VC_STATUS_PLANT_YARD": "", "VC_STATUS_ASSEMBLER_YARD": "", "VC_WAREHOUSE": "",
                           "VC_TERMINATED": "", "VC_STATUS_EMPTY_TRAILER": ""}, **k)
    keys += [p["eventKey"] for p in receiving.computePosts(
        "update", oo(VC_STATUS_SUPPLIER_SHIPPING="20260618"), oo(IN_QTY=140, VC_STATUS_SUPPLIER_SHIPPING="20260618"),
        addPoint="S")]
    keys += [p["eventKey"] for p in receiving.computePosts(   # RECEIVING_REVERSAL (arrival cleared, 'A')
        "update", oo(VC_ARRIVAL="20260618"), oo(VC_ARRIVAL=""), addPoint="A")]
    keys += [p["eventKey"] for p in receiving.computePosts(   # F5 part-change two-post
        "update", oo(VC_PART_NUMBER="OLDP", VC_STATUS_SUPPLIER_SHIPPING="20260618"),
        oo(VC_PART_NUMBER="NEWP", VC_STATUS_SUPPLIER_SHIPPING="20260618"), addPointOld="S", addPointNew="S")]
    return keys


def main():
    print("=" * 78)
    print(" EVENT-KEY INTEGRITY — real computePosts keys fit VC_SOURCE_EVENT(100) + round-trip intact")
    print("=" * 78)
    rep = Report()
    keys = collect_keys()
    longest = max(keys, key=len)
    print("  collected %d real event keys; longest = %d chars: %s" % (len(keys), len(longest), longest))
    rep.check("all real event keys fit VC_SOURCE_EVENT(%d)" % COLW, all(len(k) <= COLW for k in keys),
              "longest %d chars" % len(longest))
    try:
        for k in keys:
            sql("EXEC POST_StockMovement @partId=%d, @delta=0, @sourceEnum='STOCKTAKING', "
                "@eventKey='%s', @reason='%s', @site=1, @purge=0" % (PART, k, REASON))
        stored = set(r[0] for r in sql("SELECT VC_SOURCE_EVENT FROM INV_STOCK_LEDGER WHERE VC_REASON='%s'" % REASON))
        intact = [k for k in keys if k in stored]
        rep.check("every key round-tripped through POST_StockMovement INTACT (no truncation)",
                  len(intact) == len(set(keys)), "%d/%d stored intact" % (len(intact), len(set(keys))))
        # belt-and-suspenders: a 45-char key would have FAILED the old varchar(40) -> prove the width matters
        rep.check("longest key exceeds the old varchar(40) (so the widening was load-bearing)",
                  len(longest) > 40, "%d > 40" % len(longest))
    finally:
        sql("DELETE FROM INV_STOCK_LEDGER WHERE VC_REASON='%s'" % REASON)
    leftover = int(scalar("SELECT COUNT(*) FROM INV_STOCK_LEDGER WHERE VC_REASON='%s'" % REASON))
    base = int(scalar("SELECT IN_QTY FROM INV_PARTS_STOCK_MST WHERE IN_PART_ID=%d" % PART))
    rep.check("fixture clean (delta=0 posts left IN_QTY untouched, ledger rows removed)", leftover == 0,
              "leftover=%d IN_QTY=%d" % (leftover, base))
    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
