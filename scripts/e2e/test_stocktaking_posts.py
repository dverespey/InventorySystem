#!/usr/bin/env python3
"""test_stocktaking_posts.py — unit test for the stocktaking -> stock-ledger wiring (design §3 #12-#14).

Pure: imports `stocktaking.computePosts` straight from the on-disk project library (no DB, no env).

Stocktaking is int-keyed and a SIGNED ADJUSTMENT (D5): effect=+IN_QTY posted verbatim (a -5 row
removes 5). IN_QTY is NULL-able -> NULL is a no-op. Verified vs the live INSERT_/UPDATE_/DELETE_
Stocktaking triggers (all join on IN_PART_ID).

Run:  python3 scripts/e2e/test_stocktaking_posts.py
"""
import os, sys, importlib.util

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report  # noqa: E402

PLIB = os.path.normpath(os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "..", "..",
    "docs", "analysis", "inventory-stock", "project-library"))


def _load(name, rel):
    spec = importlib.util.spec_from_file_location(name, os.path.join(PLIB, rel, "code.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


stk = _load("stocktaking_lib", "stocktaking")
stockLedger = _load("stockledger_lib", "stockLedger")

_COLS = ("IN_PART_ID", "IN_QTY", "IN_STOCKTAKING_ID", "VC_LAST_UPDATE")


def row(part=16, qty=50, stkid=903, upd="20260618120000ff"):
    return dict(zip(_COLS, (part, qty, stkid, upd)))


def only(posts):
    assert len(posts) == 1, "expected 1 post, got %d: %r" % (len(posts), posts)
    return posts[0]


def main():
    print("=" * 78)
    print(" STOCKTAKING -> STOCK-LEDGER wiring — pure computePosts unit test (design §3 #12-#14)")
    print("=" * 78)
    rep = Report()
    S, cp = stk, stk.computePosts

    print("\n--- enum contract ---")
    rep.check("ENUM_STOCKTAKING matches stockLedger",
              S.ENUM_STOCKTAKING == stockLedger.ENUM_STOCKTAKING, S.ENUM_STOCKTAKING)

    print("\n--- INSERT (signed adjustment, posted verbatim — D5) ---")
    p = only(cp("insert", new=row(qty=50)))
    rep.check("insert +50 -> +50 (positive adjustment)",
              p["delta"] == 50 and p["sourceEnum"] == S.ENUM_STOCKTAKING, "d=%d key=%s" % (p["delta"], p["eventKey"]))
    rep.check("insert key ':ins' on IN_STOCKTAKING_ID", p["eventKey"] == "STOCKTAKING:stk=903:ins", p["eventKey"])
    rep.check("post carries partId (int) + sourceRowId",
              p["partId"] == 16 and p["sourceRowId"] == 903, "partId=%s stk=%s" % (p["partId"], p["sourceRowId"]))
    p = only(cp("insert", new=row(qty=-5)))
    rep.check("insert -5 -> -5 (NEGATIVE adjustment posted verbatim, D5)", p["delta"] == -5, str(p))
    rep.check("insert NULL qty -> no post (no NULL-poison)", cp("insert", new=row(qty=None)) == [])
    rep.check("insert blank qty -> no post", cp("insert", new=row(qty="")) == [])

    print("\n--- DELETE (reverse the adjustment) ---")
    p = only(cp("delete", old=row(qty=50)))
    rep.check("delete +50 -> -50 (reverse)", p["delta"] == -50 and p["eventKey"] == "STOCKTAKING:stk=903:del", str(p))
    p = only(cp("delete", old=row(qty=-5)))
    rep.check("delete -5 -> +5 (reverse a negative adj.)", p["delta"] == 5, str(p))

    print("\n--- UPDATE (amend, same part) ---")
    p = only(cp("update", old=row(qty=50), new=row(qty=80)))
    rep.check("amend +50->+80 -> +30 (net new-old)", p["delta"] == 30, str(p))
    rep.check("amend key ':upd:v=<ver>'", p["eventKey"] == "STOCKTAKING:stk=903:upd:v=20260618120000ff", p["eventKey"])
    p = only(cp("update", old=row(qty=50), new=row(qty=20)))
    rep.check("amend +50->+20 -> -30", p["delta"] == -30, str(p))
    rep.check("amend no change -> no post", cp("update", old=row(qty=50), new=row(qty=50)) == [])

    print("\n--- UPDATE (part-id change -> faithful two-post) ---")
    posts = cp("update", old=row(part=7, qty=50), new=row(part=9, qty=30))
    byPart = dict((x["partId"], x) for x in posts)
    rep.check("part change emits TWO posts", len(posts) == 2, str([x["partId"] for x in posts]))
    rep.check("backs out old part (-50)", byPart.get(7, {}).get("delta") == -50, str(byPart.get(7)))
    rep.check("applies new part (+30)", byPart.get(9, {}).get("delta") == 30, str(byPart.get(9)))

    print("\n--- idempotency keys ---")
    k1 = only(cp("update", old=row(qty=50), new=row(qty=60, upd="AAA")))["eventKey"]
    k2 = only(cp("update", old=row(qty=60), new=row(qty=70, upd="BBB")))["eventKey"]
    rep.check("successive amends distinct keys (version = VC_LAST_UPDATE)", k1 != k2, "%s vs %s" % (k1, k2))
    ka = only(cp("insert", new=row(stkid=11)))["eventKey"]
    kb = only(cp("insert", new=row(stkid=22)))["eventKey"]
    rep.check("different stocktaking rows distinct insert keys", ka != kb, "%s vs %s" % (ka, kb))

    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
