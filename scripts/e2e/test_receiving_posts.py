#!/usr/bin/env python3
"""test_receiving_posts.py — unit test for the receiving -> stock-ledger wiring.

Exercises the PURE decision logic `receiving.computePosts` (the §3 #1-#8 mapping) with NO database:
the function imports nothing and is CPython-importable straight from the on-disk project library.

Design: docs/analysis/inventory-stock/IGNITION-stock-ledger-design.md §3 / §4;
source spec: docs/analysis/receiving/recconfstat.md §2-§4 (D7 / D8(3) / D12#3 / F5).

Asserts each receiving path posts the correct part / signed delta / VC_SOURCE_ENUM / event key —
including the divergences the legacy triggers get wrong (D8(3) arrival reversal, D12#3 yard-on-edit,
F5 part-change two-post) and the gates that post nothing (terminated / empty-trailer / not-counted).

Run:  python3 scripts/e2e/test_receiving_posts.py     (no DB, no env needed)
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


receiving = _load("receiving_lib", "receiving")
stockLedger = _load("stockledger_lib", "stockLedger")


# ---- row builder -----------------------------------------------------------------------------
_COLS = ("VC_PART_NUMBER", "IN_QTY", "VC_STATUS_SUPPLIER_SHIPPING", "VC_ARRIVAL",
         "VC_STATUS_PLANT_YARD", "VC_STATUS_ASSEMBLER_YARD", "VC_WAREHOUSE",
         "VC_TERMINATED", "VC_STATUS_EMPTY_TRAILER", "VC_LAST_UPDATE", "IN_ORDER_ID")
D = "20260618120000"   # an 8/16-char-ish date stamp, "set" for any status column


def row(part="P1", qty=100, ship="", arr="", plant="", assy="", wh="",
        term="", empty="", ts="20260618120000ff", oid=8842):
    return dict(zip(_COLS, (part, qty, ship, arr, plant, assy, wh, term, empty, ts, oid)))


def only(posts):
    """Assert exactly one post and return it (fails loudly otherwise)."""
    assert len(posts) == 1, "expected 1 post, got %d: %r" % (len(posts), posts)
    return posts[0]


def main():
    print("=" * 78)
    print(" RECEIVING -> STOCK-LEDGER wiring — pure computePosts unit test (design §3)")
    print("=" * 78)
    rep = Report()
    R, cp = receiving, receiving.computePosts

    # --- enum contract: receiving's local mirrors MUST equal stockLedger's (no silent drift) ----
    print("\n--- enum contract ---")
    rep.check("ENUM_RECEIVING_SHIP matches stockLedger",
              R.ENUM_RECEIVING_SHIP == stockLedger.ENUM_RECEIVING_SHIP, R.ENUM_RECEIVING_SHIP)
    rep.check("ENUM_RECEIVING_ARRIVAL matches stockLedger",
              R.ENUM_RECEIVING_ARRIVAL == stockLedger.ENUM_RECEIVING_ARRIVAL, R.ENUM_RECEIVING_ARRIVAL)
    rep.check("ENUM_RECEIVING_REVERSAL matches stockLedger",
              R.ENUM_RECEIVING_REVERSAL == stockLedger.ENUM_RECEIVING_REVERSAL, R.ENUM_RECEIVING_REVERSAL)

    # --- INSERT (§3 #1/#2) ----------------------------------------------------------------------
    print("\n--- INSERT ---")
    p = only(cp("insert", new=row(qty=100, ship=D), addPoint="S"))
    rep.check("insert 'S' shipped -> +qty SHIP",
              p["delta"] == 100 and p["sourceEnum"] == R.ENUM_RECEIVING_SHIP,
              "%s d=%d key=%s" % (p["sourceEnum"], p["delta"], p["eventKey"]))
    rep.check("insert event key is ':ins'", p["eventKey"] == "RECEIVING_SHIP:ord=8842:ins", p["eventKey"])
    rep.check("insert sets sourceRowId = IN_ORDER_ID", p["sourceRowId"] == 8842, str(p["sourceRowId"]))

    p = only(cp("insert", new=row(qty=60, arr=D), addPoint="A"))
    rep.check("insert 'A' arrived(arrival) -> +qty ARRIVAL",
              p["delta"] == 60 and p["sourceEnum"] == R.ENUM_RECEIVING_ARRIVAL, str(p))

    p = only(cp("insert", new=row(qty=60, plant=D), addPoint="A"))
    rep.check("insert 'A' arrived via PLANT-YARD only -> +qty (D12#3)",
              p["delta"] == 60 and p["sourceEnum"] == R.ENUM_RECEIVING_ARRIVAL, str(p))

    rep.check("insert 'S' NOT shipped -> no post", cp("insert", new=row(qty=100), addPoint="S") == [])
    rep.check("insert add-point other -> no post (no stock move)",
              cp("insert", new=row(qty=100, ship=D), addPoint="X") == [])

    # --- UPDATE same part (§3 #3-#7) ------------------------------------------------------------
    print("\n--- UPDATE (same part) ---")
    p = only(cp("update", old=row(qty=100, ship=D), new=row(qty=140, ship=D), addPoint="S"))
    rep.check("edit qty while shipped 'S' 100->140 -> +40 SHIP (delta, not double-count)",
              p["delta"] == 40 and p["sourceEnum"] == R.ENUM_RECEIVING_SHIP, str(p))
    rep.check("edit event key is ':upd:v=<ts>'",
              p["eventKey"] == "RECEIVING_SHIP:ord=8842:upd:v=20260618120000ff", p["eventKey"])

    p = only(cp("update", old=row(qty=100), new=row(qty=100, ship=D), addPoint="S"))
    rep.check("ship-status set 'S' blank->set -> +qty SHIP", p["delta"] == 100, str(p))
    p = only(cp("update", old=row(qty=100, ship=D), new=row(qty=100), addPoint="S"))
    rep.check("ship-status cleared 'S' set->blank -> -qty SHIP", p["delta"] == -100, str(p))

    p = only(cp("update", old=row(qty=80), new=row(qty=80, arr=D), addPoint="A"))
    rep.check("arrival set 'A' blank->set -> +qty ARRIVAL (D7)",
              p["delta"] == 80 and p["sourceEnum"] == R.ENUM_RECEIVING_ARRIVAL, str(p))

    p = only(cp("update", old=row(qty=80, arr=D), new=row(qty=80), addPoint="A"))
    rep.check("arrival cleared 'A' set->blank -> -qty REVERSAL (D8(3) — DEAD in legacy)",
              p["delta"] == -80 and p["sourceEnum"] == R.ENUM_RECEIVING_REVERSAL, str(p))

    rep.check("arrival cleared but PLANT-YARD still set -> no post (still arrived)",
              cp("update", old=row(qty=80, arr=D, plant=D), new=row(qty=80, plant=D), addPoint="A") == [])

    p = only(cp("update", old=row(qty=80, arr=D), new=row(qty=50, arr=D), addPoint="A"))
    rep.check("edit qty while arrived 'A' 80->50 -> -30 ARRIVAL (still arrived, not a reversal)",
              p["delta"] == -30 and p["sourceEnum"] == R.ENUM_RECEIVING_ARRIVAL, str(p))

    p = only(cp("update", old=row(qty=80), new=row(qty=80, plant=D), addPoint="A"))
    rep.check("yard set on EDIT 'A' blank->plant-yard -> +qty ARRIVAL (D12#3 under-count fix)",
              p["delta"] == 80, str(p))

    rep.check("edit with no on-hand-affecting change -> no post",
              cp("update", old=row(qty=100, ship=D), new=row(qty=100, ship=D, term=D), addPoint="S") == [])

    # --- UPDATE part re-pointed (§4.3 F5: two posts, two parts) ---------------------------------
    print("\n--- UPDATE (F5 part-number change) ---")
    posts = cp("update", old=row(part="OLD", qty=100, ship=D), new=row(part="NEW", qty=70, ship=D),
               addPointOld="S", addPointNew="S")
    byPart = dict((x["partNumber"], x) for x in posts)
    rep.check("F5 part change emits TWO posts (old + new part)", len(posts) == 2, str([p["partNumber"] for p in posts]))
    rep.check("F5 removes old part's effect", byPart.get("OLD", {}).get("delta") == -100, str(byPart.get("OLD")))
    rep.check("F5 adds new part's effect", byPart.get("NEW", {}).get("delta") == 70, str(byPart.get("NEW")))
    rep.check("F5 keys are distinct (del-old / add-new)",
              byPart["OLD"]["eventKey"].endswith("upd:del-old:v=20260618120000ff")
              and byPart["NEW"]["eventKey"].endswith("upd:add-new:v=20260618120000ff"),
              "%s | %s" % (byPart["OLD"]["eventKey"], byPart["NEW"]["eventKey"]))

    posts = cp("update", old=row(part="OLD", qty=100), new=row(part="NEW", qty=70, ship=D),
               addPointOld="S", addPointNew="S")
    rep.check("F5 with old part NOT counted -> single +new post",
              len(posts) == 1 and posts[0]["partNumber"] == "NEW" and posts[0]["delta"] == 70, str(posts))

    # --- DELETE (§3 #8 + §3.1 gates) ------------------------------------------------------------
    print("\n--- DELETE ---")
    p = only(cp("delete", old=row(qty=100, ship=D), addPoint="S"))
    rep.check("delete active shipped 'S' -> -qty SHIP", p["delta"] == -100, str(p))
    rep.check("delete event key is ':del'", p["eventKey"] == "RECEIVING_SHIP:ord=8842:del", p["eventKey"])

    p = only(cp("delete", old=row(qty=60, arr=D), addPoint="A"))
    rep.check("delete active arrived 'A' -> -qty ARRIVAL", p["delta"] == -60, str(p))

    rep.check("delete TERMINATED order -> no post (§3.1 gate / DELETE_AutoPurge pre-stamp)",
              cp("delete", old=row(qty=100, ship=D, term=D), addPoint="S") == [])
    rep.check("delete EMPTY-TRAILER order -> no post (live gate)",
              cp("delete", old=row(qty=100, ship=D, empty=D), addPoint="S") == [])
    rep.check("delete NOT-counted order -> no post",
              cp("delete", old=row(qty=100), addPoint="S") == [])

    # --- idempotency-key behaviour --------------------------------------------------------------
    print("\n--- idempotency keys ---")
    k1 = only(cp("update", old=row(qty=100, ship=D), new=row(qty=120, ship=D, ts="AAA"), addPoint="S"))["eventKey"]
    k2 = only(cp("update", old=row(qty=120, ship=D), new=row(qty=130, ship=D, ts="BBB"), addPoint="S"))["eventKey"]
    rep.check("successive edits get DISTINCT event keys (version = VC_LAST_UPDATE)", k1 != k2, "%s vs %s" % (k1, k2))
    ka = only(cp("insert", new=row(qty=10, ship=D, oid=11), addPoint="S"))["eventKey"]
    kb = only(cp("insert", new=row(qty=10, ship=D, oid=22), addPoint="S"))["eventKey"]
    rep.check("different orders get distinct insert keys", ka != kb, "%s vs %s" % (ka, kb))

    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
