#!/usr/bin/env python3
"""dress_rehearsal_smoke.py — Phase-C forward-post smoke test for the cutover DRESS REHEARSAL.

Runs AFTER the 13 qty-triggers are DROPPED and SEED_AllOpeningBalances has seeded the ledger
(IN_QTY == SUM(ledger), zero drift). Proves the producer SEAMS are now the SOLE IN_QTY writer
(no double-count) by driving ONE real receiving insert and ONE real shipping insert through the
ACTUAL gateway-side Jython wrappers (via jython_shim) — NOT EXEC-ing the procs.

Unlike test_seam_driver.py this does NOT disable/enable triggers: in the rehearsal the qty-triggers
are already DROPPED, so the seam's POST_StockMovement is the only thing that moves IN_QTY.

Asserts per op: IN_QTY moves by EXACTLY the delta ONCE, a ledger row is written, and the global
invariant IN_QTY == SUM(ledger) STILL holds for the test part after. Then asserts the SEED genesis
guard THROWs if you try to seed a part that now has a forward ledger row.

Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 e2e/dress_rehearsal_smoke.py
"""
import os, subprocess, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import jython_shim  # noqa: E402

CONTAINER = os.environ.get("CONTAINER", "mssql-spike")
SA_PASS = os.environ.get("SA_PASS")
DB = "Inventory"
PART = 16
PARTNUM = "478930223000"        # part 16's VC_PART_NUMBER; add-point 'S' (shipped order counts)
RECV_RENBAN = "DRSRC01"
RECV_SHIP_STAMP = "20260618"    # non-blank => 'S' add-point counts (+IN_QTY)
SHIP_DATE = "29991230"          # distinctive VC_PRODUCTION_DATE for the shipping fixture
SHIP_LINE = "DRSSHIPLINE"
PLIB = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..",
                                     "ignition", "project-library"))

_passed = []
_failed = []


def check(label, cond, detail=""):
    (_passed if cond else _failed).append(label)
    print(("  PASS " if cond else "  FAIL ") + label + (("  [%s]" % detail) if detail else ""))


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


def ledger_sum():
    return int(scalar("SELECT ISNULL(SUM(IN_QTY_CHANGE),0) FROM INV_STOCK_LEDGER WHERE IN_PART_ID=%d" % PART) or 0)


def make_ship_header():
    return int(scalar(
        "INSERT INTO INV_SHIPPING_INF (VC_LINE_NAME, VC_START_SEQ_NUMBER, VC_END_SEQ_NUMBER, "
        "VC_PRODUCTION_DATE, VC_ADD) OUTPUT INSERTED.IN_SHIPPING_ID VALUES "
        "('%s', '0', '0', '%s', CONVERT(varchar,getdate(),112))" % (SHIP_LINE, SHIP_DATE)))


def main():
    print("=" * 78)
    print(" DRESS-REHEARSAL Phase-C smoke: seams are the SOLE IN_QTY writer (triggers DROPPED)")
    print("=" * 78)
    if not SA_PASS:
        sys.exit("export SA_PASS first")

    sl = jython_shim.load_wrapper("sl", os.path.join(PLIB, "stockLedger", "code.py"))
    rcv = jython_shim.load_wrapper("rcv", os.path.join(PLIB, "receiving", "code.py"),
                                   extra_globals={"stockLedger": sl})
    shp = jython_shim.load_wrapper("shp", os.path.join(PLIB, "shipping", "code.py"),
                                   extra_globals={"stockLedger": sl})

    base = qty()
    base_ls = ledger_sum()
    print("  part %d (%s) base IN_QTY=%d  ledger_sum=%d  (invariant holds: %s)"
          % (PART, PARTNUM, base, base_ls, base == base_ls))
    check("pre-smoke invariant: IN_QTY == SUM(ledger) for test part (post-seed)", base == base_ls,
          "qty=%d sum=%d" % (base, base_ls))

    recv_oid = None
    ship_id = ship_hdr = None
    try:
        # ---- RECEIVING insert (+60) through the REAL wrapper; triggers are GONE ----
        print("\n--- receiving (real insertOpenOrder; qty-triggers DROPPED) ---")
        ap = rcv.resolveAddPoint(PARTNUM)
        check("resolveAddPoint -> 'S' (shipped order counts)", ap == "S", "ap=%r" % ap)
        oo = {"VC_SUPPLIER_CODE": "DRS", "VC_PART_NUMBER": PARTNUM, "VC_FRS_NUMBER": "DRS0001",
              "VC_RENBAN_NUMBER": RECV_RENBAN, "IN_QTY": 60,
              "VC_STATUS_SUPPLIER_SHIPPING": RECV_SHIP_STAMP}
        recv_oid = rcv.insertOpenOrder(oo)
        check("real insertOpenOrder +60 moves IN_QTY by EXACTLY +60 ONCE (no trigger double-count)",
              qty() == base + 60, "IN_QTY=%d (expected %d)" % (qty(), base + 60))
        rled = int(scalar("SELECT IN_QTY_CHANGE FROM INV_STOCK_LEDGER "
                          "WHERE VC_SOURCE_EVENT='RECEIVING_SHIP:ord=%s:ins'" % recv_oid) or 0)
        check("receiving wrote exactly its ledger row (+60)", rled == 60, "ledger=%d" % rled)
        check("invariant STILL holds after receiving post", qty() == ledger_sum(),
              "qty=%d sum=%d" % (qty(), ledger_sum()))

        # ---- SHIPPING insert (-50) through the REAL wrapper; triggers are GONE ----
        print("\n--- shipping (real insertPartShipping; qty-triggers DROPPED) ---")
        ship_hdr = make_ship_header()
        ship_id = shp.insertPartShipping(ship_hdr, PARTNUM, SHIP_DATE, 50)
        check("real insertPartShipping -50 moves IN_QTY by EXACTLY -50 ONCE (no trigger double-count)",
              qty() == base + 60 - 50, "IN_QTY=%d (expected %d)" % (qty(), base + 10))
        sled = int(scalar("SELECT IN_QTY_CHANGE FROM INV_STOCK_LEDGER "
                          "WHERE VC_SOURCE_EVENT='SHIPPING:psh=%s:ins'" % ship_id) or 0)
        check("shipping wrote exactly its ledger row (-50)", sled == -50, "ledger=%d" % sled)
        check("invariant STILL holds after shipping post", qty() == ledger_sum(),
              "qty=%d sum=%d" % (qty(), ledger_sum()))

        # ---- GENESIS GUARD: seeding a part that has a forward row (and NO opening row) must THROW ----
        # NOTE: SEED_OpeningBalance checks the IDEMPOTENCY guard (opening row already exists -> RETURN)
        # BEFORE the genesis guard, so calling it on an ALREADY-SEEDED part short-circuits silently — that
        # is correct/idempotent, NOT the guard. To genuinely exercise the genesis guard we must present a
        # part that has a forward (non-OPENING) ledger row but NO opening row (a misordered cutover).
        print("\n--- genesis guard: SEED on a part with a forward row but NO opening row must THROW ---")
        sql("DELETE FROM INV_STOCK_LEDGER WHERE IN_PART_ID=%d AND VC_SOURCE_ENUM='OPENING_BALANCE'; "
            "INSERT INTO INV_STOCK_LEDGER (IN_PART_ID,IN_QTY_CHANGE,IN_BALANCE_AFTER,VC_SOURCE_ENUM,"
            "IN_SOURCE_ROW_ID,VC_SOURCE_EVENT,site_id,VC_REASON,TS_POSTED,VC_ADD) VALUES "
            "(%d,5,5,'RECEIVING_SHIP',999999,'DRS_GUARD_PROBE',1,'guard probe',"
            "'2026061900000000','2026061900000000')" % (PART, PART))
        out = subprocess.run([
            "docker", "exec", CONTAINER, "/opt/mssql-tools18/bin/sqlcmd", "-C", "-S", "localhost",
            "-U", "sa", "-P", SA_PASS, "-d", DB, "-b",
            "-Q", "SET NOCOUNT ON; EXEC dbo.SEED_OpeningBalance @partId=%d;" % PART],
            capture_output=True, text=True)
        threw = (out.returncode != 0) and ("must be seeded BEFORE" in (out.stdout + out.stderr))
        check("SEED_OpeningBalance THROWs on a part with forward ledger rows (genesis guard)",
              threw, "rc=%d msg=%r" % (out.returncode, (out.stdout + out.stderr).strip()[:90]))
        # restore: remove the probe forward row and re-create the opening row to match IN_QTY
        sql("DELETE FROM INV_STOCK_LEDGER WHERE IN_PART_ID=%d AND VC_SOURCE_EVENT='DRS_GUARD_PROBE'; "
            "INSERT INTO INV_STOCK_LEDGER (IN_PART_ID,IN_QTY_CHANGE,IN_BALANCE_AFTER,VC_SOURCE_ENUM,"
            "IN_SOURCE_ROW_ID,VC_SOURCE_EVENT,site_id,VC_REASON,TS_POSTED,VC_ADD) "
            "SELECT %d,IN_QTY,IN_QTY,'OPENING_BALANCE',NULL,'OPENING:part=%d',1,'cutover opening balance',"
            "'2026061900000000','2026061900000000' FROM INV_PARTS_STOCK_MST WHERE IN_PART_ID=%d"
            % (PART, PART, PART, PART))
    finally:
        # teardown: reverse the two forward posts + remove fixtures/ledger so the part returns to base.
        # (The DB-level RESTORE in step 5 is the real reset; this keeps the rehearsal self-consistent.)
        cleanup = []
        cleanup.append("DELETE FROM INV_STOCK_LEDGER WHERE IN_PART_ID=%d AND VC_SOURCE_ENUM IN "
                       "('RECEIVING_SHIP','RECEIVING_ARRIVAL','RECEIVING_REVERSAL','SHIPPING')" % PART)
        cleanup.append("DELETE FROM INV_OPEN_ORDER_INF WHERE VC_RENBAN_NUMBER='%s'" % RECV_RENBAN)
        cleanup.append("DELETE FROM INV_OPEN_ORDER_INF_HIST WHERE VC_RENBAN_NUMBER='%s'" % RECV_RENBAN)
        cleanup.append("DELETE FROM INV_PART_SHIPPING_INF WHERE VC_PRODUCTION_DATE='%s'" % SHIP_DATE)
        cleanup.append("DELETE FROM INV_SHIPPING_INF WHERE VC_LINE_NAME='%s'" % SHIP_LINE)
        cleanup.append("UPDATE INV_PARTS_STOCK_MST SET IN_QTY=%d WHERE IN_PART_ID=%d" % (base, PART))
        sql("; ".join(cleanup))

    after = qty()
    after_ls = ledger_sum()
    check("teardown: part back to base IN_QTY and seed-only ledger (invariant intact)",
          after == base and after == after_ls, "qty=%d base=%d sum=%d" % (after, base, after_ls))

    print("\n" + "=" * 78)
    print(" SMOKE RESULT: %d passed, %d failed" % (len(_passed), len(_failed)))
    if _failed:
        for f in _failed:
            print("   FAILED: " + f)
    print("=" * 78)
    sys.exit(1 if _failed else 0)


if __name__ == "__main__":
    main()
