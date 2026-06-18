#!/usr/bin/env python3
"""test_reject_posts.py — unit test for the reject -> stock-ledger wiring (design §3 #9-#11).

Pure: imports `reject.computePosts` straight from the on-disk project library (no DB, no env).

Reject is int-keyed stock-OUT (effect=-IN_QTY, no gate); the delete always restores (design F2 — no
purge case). Verified vs the live INSERT_/UPDATE_/DELETE_RejectParts triggers (all join on IN_PART_ID).

Run:  python3 scripts/e2e/test_reject_posts.py
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


reject = _load("reject_lib", "reject")
stockLedger = _load("stockledger_lib", "stockLedger")

_COLS = ("IN_PART_ID", "IN_QTY", "IN_REJECT_ID", "VC_LAST_UPDATE")


def row(part=16, qty=100, rej=551, upd="20260618120000ff"):
    return dict(zip(_COLS, (part, qty, rej, upd)))


def only(posts):
    assert len(posts) == 1, "expected 1 post, got %d: %r" % (len(posts), posts)
    return posts[0]


def main():
    print("=" * 78)
    print(" REJECT -> STOCK-LEDGER wiring — pure computePosts unit test (design §3 #9-#11)")
    print("=" * 78)
    rep = Report()
    R, cp = reject, reject.computePosts

    print("\n--- enum contract ---")
    rep.check("ENUM_REJECT matches stockLedger", R.ENUM_REJECT == stockLedger.ENUM_REJECT, R.ENUM_REJECT)

    print("\n--- INSERT (remove) ---")
    p = only(cp("insert", new=row(qty=100)))
    rep.check("insert -> -qty (remove from on-hand)",
              p["delta"] == -100 and p["sourceEnum"] == R.ENUM_REJECT, "d=%d key=%s" % (p["delta"], p["eventKey"]))
    rep.check("insert key ':ins' on IN_REJECT_ID", p["eventKey"] == "REJECT:rej=551:ins", p["eventKey"])
    rep.check("post carries partId (int, no resolution) + sourceRowId",
              p["partId"] == 16 and p["sourceRowId"] == 551, "partId=%s rej=%s" % (p["partId"], p["sourceRowId"]))
    rep.check("insert qty 0 -> no post", cp("insert", new=row(qty=0)) == [])

    print("\n--- DELETE (un-reject restores; always, F2 no purge) ---")
    p = only(cp("delete", old=row(qty=100)))
    rep.check("delete -> +qty (restore)", p["delta"] == 100 and p["eventKey"] == "REJECT:rej=551:del", str(p))

    print("\n--- UPDATE (amend, same part) ---")
    p = only(cp("update", old=row(qty=100), new=row(qty=140)))
    rep.check("amend 100->140 -> -40 (reject 40 more)", p["delta"] == -40, str(p))
    rep.check("amend key ':upd:to=<target>:v=<ver>' (collision-safe)",
              p["eventKey"] == "REJECT:rej=551:upd:to=-140:v=20260618120000ff", p["eventKey"])
    p = only(cp("update", old=row(qty=140), new=row(qty=100)))
    rep.check("amend 140->100 -> +40 (restore 40)", p["delta"] == 40, str(p))
    rep.check("amend no change -> no post", cp("update", old=row(qty=100), new=row(qty=100)) == [])

    print("\n--- UPDATE (part-id change -> faithful two-post; domain-forbidden but handled) ---")
    posts = cp("update", old=row(part=7, qty=100), new=row(part=9, qty=70))
    byPart = dict((x["partId"], x) for x in posts)
    rep.check("part change emits TWO posts", len(posts) == 2, str([x["partId"] for x in posts]))
    rep.check("restores old part (+100)", byPart.get(7, {}).get("delta") == 100, str(byPart.get(7)))
    rep.check("removes from new part (-70)", byPart.get(9, {}).get("delta") == -70, str(byPart.get(9)))

    print("\n--- idempotency keys ---")
    k1 = only(cp("update", old=row(qty=100), new=row(qty=120, upd="AAA")))["eventKey"]
    k2 = only(cp("update", old=row(qty=120), new=row(qty=130, upd="BBB")))["eventKey"]
    rep.check("successive amends distinct keys (version = VC_LAST_UPDATE)", k1 != k2, "%s vs %s" % (k1, k2))
    # collision-safety: SAME stamp, different targets -> distinct keys (the :to= discriminator).
    s1 = only(cp("update", old=row(qty=100), new=row(qty=140, upd="SAME")))["eventKey"]
    s2 = only(cp("update", old=row(qty=140), new=row(qty=180, upd="SAME")))["eventKey"]
    rep.check("same-stamp distinct-target amends get distinct keys (:to= collision-safety)",
              s1 != s2, "%s vs %s" % (s1, s2))
    ka = only(cp("insert", new=row(rej=11)))["eventKey"]
    kb = only(cp("insert", new=row(rej=22)))["eventKey"]
    rep.check("different reject rows distinct insert keys", ka != kb, "%s vs %s" % (ka, kb))

    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
