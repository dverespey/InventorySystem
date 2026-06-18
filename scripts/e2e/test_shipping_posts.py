#!/usr/bin/env python3
"""test_shipping_posts.py — unit test for the shipping -> stock-ledger wiring.

Exercises the PURE decision logic `shipping.computePosts` / `computeHeaderDeletePosts` (the §3
#15-#17 mapping) with NO database: the functions import nothing and are CPython-importable straight
from the on-disk project library.

Design: docs/analysis/inventory-stock/IGNITION-stock-ledger-design.md §3 #15-#17 / §4;
source spec: docs/analysis/shipping/shipping.md §2 (the four live trigger bodies), §4, §6.

Asserts: stock-OUT on insert (-qty), restore on delete (+qty), net delta on amend, the part-number
amend two-post (parity with the independent-join UpdatePartShipping), and the header-delete cascade
restoring EVERY row (multi-row-safe F3 fix), scoped by the driver to one header's rows.

Run:  python3 scripts/e2e/test_shipping_posts.py     (no DB, no env needed)
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


shipping = _load("shipping_lib", "shipping")
stockLedger = _load("stockledger_lib", "stockLedger")

_COLS = ("VC_PART_NUMBER", "IN_QTY", "IN_PART_SHIPPING_ID", "IN_SHIPPING_ID", "VC_ADD")


def row(part="P1", qty=100, psh=12077, ship=900, add="20260618120000ff"):
    return dict(zip(_COLS, (part, qty, psh, ship, add)))


def only(posts):
    assert len(posts) == 1, "expected 1 post, got %d: %r" % (len(posts), posts)
    return posts[0]


def main():
    print("=" * 78)
    print(" SHIPPING -> STOCK-LEDGER wiring — pure computePosts unit test (design §3 #15-#17)")
    print("=" * 78)
    rep = Report()
    S, cp = shipping, shipping.computePosts

    print("\n--- enum contract ---")
    rep.check("ENUM_SHIPPING matches stockLedger",
              S.ENUM_SHIPPING == stockLedger.ENUM_SHIPPING, S.ENUM_SHIPPING)

    # --- INSERT (§3 #15: always subtract, no add-point gate) ------------------------------------
    print("\n--- INSERT (consume) ---")
    p = only(cp("insert", new=row(qty=100)))
    rep.check("insert -> -qty (stock-OUT, no add-point gate)",
              p["delta"] == -100 and p["sourceEnum"] == S.ENUM_SHIPPING,
              "%s d=%d key=%s" % (p["sourceEnum"], p["delta"], p["eventKey"]))
    rep.check("insert event key is ':ins' keyed on IN_PART_SHIPPING_ID",
              p["eventKey"] == "SHIPPING:psh=12077:ins", p["eventKey"])
    rep.check("insert sets sourceRowId = IN_PART_SHIPPING_ID", p["sourceRowId"] == 12077, str(p["sourceRowId"]))
    rep.check("insert qty 0 -> no post", cp("insert", new=row(qty=0)) == [])

    # --- DELETE (§3 #17: restore) ---------------------------------------------------------------
    print("\n--- DELETE (restore) ---")
    p = only(cp("delete", old=row(qty=100)))
    rep.check("delete -> +qty (restore)", p["delta"] == 100, str(p))
    rep.check("delete event key is ':del'", p["eventKey"] == "SHIPPING:psh=12077:del", p["eventKey"])

    # --- UPDATE same part (§3 #16: net delta = old - new) ---------------------------------------
    print("\n--- UPDATE (amend, same part) ---")
    p = only(cp("update", old=row(qty=100), new=row(qty=140)))
    rep.check("amend qty 100->140 -> -40 (consume 40 more)", p["delta"] == -40, str(p))
    rep.check("amend event key is ':upd:v=<ver>'",
              p["eventKey"] == "SHIPPING:psh=12077:upd:v=20260618120000ff", p["eventKey"])
    p = only(cp("update", old=row(qty=140), new=row(qty=100)))
    rep.check("amend qty 140->100 -> +40 (restore 40)", p["delta"] == 40, str(p))
    rep.check("amend with no qty change -> no post", cp("update", old=row(qty=100), new=row(qty=100)) == [])

    # --- UPDATE part-number amend (two posts; PARITY with independent-join UpdatePartShipping) ---
    print("\n--- UPDATE (part-number amend -> two posts) ---")
    posts = cp("update", old=row(part="OLD", qty=100), new=row(part="NEW", qty=70))
    byPart = dict((x["partNumber"], x) for x in posts)
    rep.check("part-number amend emits TWO posts", len(posts) == 2, str([x["partNumber"] for x in posts]))
    rep.check("restores OLD part (+old qty)", byPart.get("OLD", {}).get("delta") == 100, str(byPart.get("OLD")))
    rep.check("consumes NEW part (-new qty)", byPart.get("NEW", {}).get("delta") == -70, str(byPart.get("NEW")))
    rep.check("part-amend keys distinct (del-old / add-new)",
              byPart["OLD"]["eventKey"].endswith("upd:del-old:v=20260618120000ff")
              and byPart["NEW"]["eventKey"].endswith("upd:add-new:v=20260618120000ff"),
              "%s | %s" % (byPart["OLD"]["eventKey"], byPart["NEW"]["eventKey"]))

    # --- header-delete cascade (DeleteShipDate): restore EVERY row, multi-row-safe (F3 fix) ------
    print("\n--- header-delete cascade ---")
    chd = S.computeHeaderDeletePosts
    posts = chd([row(part="A", qty=100, psh=1), row(part="B", qty=50, psh=2)])
    rep.check("cascade restores each distinct row (+qty each)",
              len(posts) == 2 and {p["partNumber"]: p["delta"] for p in posts} == {"A": 100, "B": 50},
              str([(p["partNumber"], p["delta"]) for p in posts]))
    # F3: two rows for the SAME part in one cascade -> BOTH restored (legacy under-restored to one).
    posts = chd([row(part="SAME", qty=30, psh=3), row(part="SAME", qty=20, psh=4)])
    total = sum(p["delta"] for p in posts)
    rep.check("cascade is MULTI-ROW-SAFE: 2 same-part rows -> BOTH restored (F3 fix, +50 total)",
              len(posts) == 2 and total == 50, "posts=%d total=%d" % (len(posts), total))
    rep.check("cascade same-part keys distinct (per IN_PART_SHIPPING_ID)",
              posts[0]["eventKey"] != posts[1]["eventKey"],
              "%s vs %s" % (posts[0]["eventKey"], posts[1]["eventKey"]))
    rep.check("empty cascade -> no posts", chd([]) == [])

    # --- idempotency keys -----------------------------------------------------------------------
    print("\n--- idempotency keys ---")
    k1 = only(cp("update", old=row(qty=100), new=row(qty=120, add="AAA")))["eventKey"]
    k2 = only(cp("update", old=row(qty=120), new=row(qty=130, add="BBB")))["eventKey"]
    rep.check("successive amends get DISTINCT keys (version = VC_ADD)", k1 != k2, "%s vs %s" % (k1, k2))
    ka = only(cp("insert", new=row(qty=10, psh=11)))["eventKey"]
    kb = only(cp("insert", new=row(qty=10, psh=22)))["eventKey"]
    rep.check("different rows get distinct insert keys", ka != kb, "%s vs %s" % (ka, kb))

    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
