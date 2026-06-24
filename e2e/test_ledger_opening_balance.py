#!/usr/bin/env python3
"""test_ledger_opening_balance.py — the cutover opening-balance backfill + forward parity (thread b,
option 2 — David's choice: "start clean, go forward from there").

Since the from-zero reconstruction is impossible on the purged snapshot (DELETE_AutoPurge aged out the
receiving IN-movements, F2), the cutover seeds the ledger with each part's CURRENT IN_QTY as one
OPENING_BALANCE movement (no IN_QTY bump — the materialized column already holds it), then runs forward.
This validator proves the load-bearing invariant SUM(ledger qty_delta) == IN_QTY:
  1. holds immediately AFTER seeding (by construction),
  2. is IDEMPOTENT (re-seed adds nothing),
  3. STAYS true as forward movements post — exercised with the streams that actually matter in current
     use: receiving (+) and shipping (-). (Reject/stocktaking are almost never used today — David.)
  4. PROC_RebuildStockBalance re-derives the same IN_QTY from the ledger.
Fixture-disciplined (removes its ledger rows, restores IN_QTY, re-enables the audit trigger).

Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 e2e/test_ledger_opening_balance.py
"""
import os, subprocess, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report  # noqa: E402

CONTAINER = os.environ.get("CONTAINER", "mssql-spike")
SA_PASS = os.environ.get("SA_PASS")
DB = "Inventory"
PART = 16
FWD_ADD, FWD_SUB = 50, 20       # a forward receiving (+) and shipping (-) movement


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


def qty():
    return int(scalar("SELECT IN_QTY FROM INV_PARTS_STOCK_MST WHERE IN_PART_ID=%d" % PART))


def ledger_sum():
    return int(scalar("SELECT ISNULL(SUM(IN_QTY_CHANGE),0) FROM INV_STOCK_LEDGER WHERE IN_PART_ID=%d" % PART) or 0)


def ledger_count():
    return int(scalar("SELECT COUNT(*) FROM INV_STOCK_LEDGER WHERE IN_PART_ID=%d" % PART) or 0)


def post(delta, enum, evt):
    sql("EXEC POST_StockMovement @partId=%d, @delta=%d, @sourceEnum='%s', @eventKey='%s', "
        "@reason='OPENBAL_TEST', @site=1, @purge=0" % (PART, delta, enum, evt))


def cleanup(base):
    sql("DELETE FROM INV_STOCK_LEDGER WHERE IN_PART_ID=%d AND (VC_SOURCE_ENUM='OPENING_BALANCE' "
        "OR VC_REASON='OPENBAL_TEST'); "
        "UPDATE INV_PARTS_STOCK_MST SET IN_QTY=%d WHERE IN_PART_ID=%d; "
        "ENABLE TRIGGER UPDATE_PartNumber ON INV_PARTS_STOCK_MST" % (PART, base, PART))


def main():
    print("=" * 78)
    print(" LEDGER OPENING-BALANCE backfill + forward parity (thread b / option 2)")
    print("=" * 78)
    rep = Report()
    base = qty()
    print("  part %d base IN_QTY = %d ; pre-existing ledger rows = %d" % (PART, base, ledger_count()))
    # forward posts bump IN_QTY -> the legacy UPDATE_PartNumber audit trigger would fire; suppress it.
    sql("DISABLE TRIGGER UPDATE_PartNumber ON INV_PARTS_STOCK_MST")
    try:
        # 0. GENESIS GUARD: seeding AFTER a forward post would double-count -> the proc must REFUSE.
        post(10, "RECEIVING_SHIP", "OPENBAL_TEST:guard")           # a forward movement lands first
        threw = scalar("BEGIN TRY EXEC SEED_OpeningBalance @partId=%d; SELECT 'NO' END TRY "
                       "BEGIN CATCH SELECT 'YES' END CATCH" % PART)
        opening_after = int(scalar("SELECT COUNT(*) FROM INV_STOCK_LEDGER WHERE IN_PART_ID=%d "
                                   "AND VC_SOURCE_ENUM='OPENING_BALANCE'" % PART) or 0)
        rep.check("genesis guard: seed AFTER a forward post THROWS and writes no opening row",
                  threw == "YES" and opening_after == 0, "threw=%s opening=%d" % (threw, opening_after))
        # reset part 16 to a clean ledger for the happy-path tests below
        sql("DELETE FROM INV_STOCK_LEDGER WHERE IN_PART_ID=%d; "
            "UPDATE INV_PARTS_STOCK_MST SET IN_QTY=%d WHERE IN_PART_ID=%d" % (PART, base, PART))

        # 1. seed: records the opening balance WITHOUT bumping IN_QTY
        sql("EXEC SEED_OpeningBalance @partId=%d, @site=1" % PART)
        rep.check("seed records ONE opening movement = IN_QTY, NO bump",
                  ledger_count() == 1 and ledger_sum() == base and qty() == base,
                  "rows=%d sum=%d IN_QTY=%d (base=%d)" % (ledger_count(), ledger_sum(), qty(), base))
        rep.check("invariant SUM(ledger) == IN_QTY holds at t0", ledger_sum() == qty())

        # 2. idempotent re-seed
        sql("EXEC SEED_OpeningBalance @partId=%d, @site=1" % PART)
        rep.check("re-seed is idempotent (still one opening row)", ledger_count() == 1, "rows=%d" % ledger_count())

        # 3. forward movements (receiving + / shipping -) keep the invariant
        post(FWD_ADD, "RECEIVING_SHIP", "OPENBAL_TEST:recv")
        rep.check("forward receiving +%d: SUM==IN_QTY==base+%d" % (FWD_ADD, FWD_ADD),
                  ledger_sum() == qty() == base + FWD_ADD, "sum=%d IN_QTY=%d" % (ledger_sum(), qty()))
        post(-FWD_SUB, "SHIPPING", "OPENBAL_TEST:ship")
        rep.check("forward shipping -%d: SUM==IN_QTY==base+%d" % (FWD_SUB, FWD_ADD - FWD_SUB),
                  ledger_sum() == qty() == base + FWD_ADD - FWD_SUB, "sum=%d IN_QTY=%d" % (ledger_sum(), qty()))

        # 4. rebuildBalance re-derives the same IN_QTY from the ledger (opening + forwards)
        sql("EXEC PROC_RebuildStockBalance @partId=%d" % PART)
        rep.check("rebuildBalance re-derives IN_QTY from ledger (== SUM, unchanged)",
                  qty() == base + FWD_ADD - FWD_SUB == ledger_sum(), "IN_QTY=%d sum=%d" % (qty(), ledger_sum()))
    finally:
        cleanup(base)

    restored = qty()
    leftover = ledger_count()
    trig = int(scalar("SELECT COUNT(*) FROM sys.triggers WHERE is_disabled=1 AND name='UPDATE_PartNumber'"))
    rep.check("DB restored as found (IN_QTY, no ledger rows, trigger re-enabled)",
              restored == base and leftover == 0 and trig == 0,
              "IN_QTY=%d ledger=%d trig_off=%d" % (restored, leftover, trig))

    # ---- FULL-SNAPSHOT: seed EVERY part, prove ledger SUM == IN_QTY for ALL parts (the cutover close) ----
    # This is the all-parts version of the invariant: after the backfill the whole ledger reconciles to
    # IN_QTY by construction. (Set-based mirror of SEED_OpeningBalance; no IN_QTY bump.) The ledger is
    # empty here, so a re-derived SUM per part must equal IN_QTY for every part.
    print("\n--- FULL-SNAPSHOT cutover backfill (all parts, via the shipped SEED_AllOpeningBalances proc) ---")
    try:
        sql("EXEC SEED_AllOpeningBalances @site=1")   # the exact set-based proc the gateway wrapper calls
        total = int(scalar("SELECT COUNT(*) FROM INV_PARTS_STOCK_MST"))
        seeded = int(scalar("SELECT COUNT(*) FROM INV_STOCK_LEDGER WHERE VC_SOURCE_ENUM='OPENING_BALANCE'"))
        drift = int(scalar(
            "SELECT COUNT(*) FROM INV_PARTS_STOCK_MST p WHERE p.IN_QTY <> "
            "(SELECT ISNULL(SUM(IN_QTY_CHANGE),0) FROM INV_STOCK_LEDGER l WHERE l.IN_PART_ID=p.IN_PART_ID)"))
        rep.check("opening row seeded for every part", seeded == total, "seeded=%d / parts=%d" % (seeded, total))
        rep.check("after backfill, ledger SUM == IN_QTY for ALL parts (0 drift) — cutover parity closed",
                  drift == 0, "%d parts with drift" % drift)
    finally:
        sql("DELETE FROM INV_STOCK_LEDGER WHERE VC_SOURCE_ENUM='OPENING_BALANCE'")
    rep.check("full-snapshot fixture clean (all opening rows removed)",
              int(scalar("SELECT COUNT(*) FROM INV_STOCK_LEDGER WHERE VC_SOURCE_ENUM='OPENING_BALANCE'") or 0) == 0)
    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
