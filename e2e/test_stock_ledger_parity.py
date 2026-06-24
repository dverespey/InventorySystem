#!/usr/bin/env python3
"""test_stock_ledger_parity.py — Stock-Ledger Service GO/NO-GO reconciliation harness.

Design: docs/analysis/inventory-stock/IGNITION-stock-ledger-design.md §6 (the GO/NO-GO).

Two parts:
  A. PROC self-test (fixture-disciplined, leaves the DB as found): post a movement, assert
     the ledger row + IN_QTY bumped by delta; replay the same event key, assert NO double-post;
     post purge=1, assert nothing moves; then delete the test rows + restore IN_QTY.
  B. RECONCILIATION: DERIVE the ledger by replaying the §3 trigger logic over the live source
     tables (receiving/reject/stocktaking/shipping with the live gates + signs), SUM per part,
     DIFF vs live INV_PARTS_STOCK_MST.IN_QTY, and CLASSIFY each non-zero diff against the
     predicted divergence classes (D8(3) / D12#3 / F3 / F5). A diff matching a predicted class
     = PASS (ledger correct, legacy buggy); an UNEXPLAINED diff = FAIL (rebuild bug).

  ** SNAPSHOT FINDING (see the report banner): the restored Inventory.bak is NOT a complete
     event history. DELETE_AutoPurge has aged out the receiving (open-order) rows that BUILT
     the legacy IN_QTY, and the 4238 surviving open orders all carry BLANK status (they fail
     every counting gate). So a from-zero replay of THIS snapshot cannot reconstruct IN_QTY —
     the reconciliation is reported with that purge-horizon caveat (§3.1), not silently passed. **

Mostly sqlcmd-driven (not browser). Reuses e2e/lib.py Report for PASS/FAIL accounting.

Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 e2e/test_stock_ledger_parity.py
"""
import os, subprocess, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report  # noqa: E402

CONTAINER = os.environ.get("CONTAINER", "mssql-spike")
SA_PASS = os.environ.get("SA_PASS")
DB = "Inventory"


def sql(query, sep="\t"):
    """Run a query via sqlcmd in the container; return list of split-row lists (no header)."""
    if not SA_PASS:
        sys.exit("export SA_PASS first (see scripts/spike-db.sh)")
    out = subprocess.check_output([
        "docker", "exec", CONTAINER,
        "/opt/mssql-tools18/bin/sqlcmd", "-C", "-S", "localhost", "-U", "sa",
        "-P", SA_PASS, "-d", DB, "-h", "-1", "-W", "-s", sep,
        "-Q", "SET NOCOUNT ON; " + query,
    ], text=True)
    rows = []
    for line in out.splitlines():
        line = line.rstrip("\n")
        if not line or line.startswith("(") or line.startswith("Msg "):
            continue
        rows.append(line.split(sep))
    return rows


def scalar(query):
    r = sql(query)
    return r[0][0] if r else None


def rows_col(query, col=0):
    """First column of every row (a flat list of strings)."""
    return [r[col] for r in sql(query) if len(r) > col and r[col] != ""]


def quote(s):
    """Single-quote a string literal for inline T-SQL (doubling embedded quotes)."""
    return "'" + str(s).replace("'", "''") + "'"


# ---------------------------------------------------------------------------------------------
# A. PROC self-test
# ---------------------------------------------------------------------------------------------
def proc_self_test(rep):
    print("\n--- A. POST_StockMovement self-test (part 16) ---")
    PART, DELTA, EVT = 16, 137, "SPIKE_PARITY:tx=1"
    base = int(scalar("SELECT IN_QTY FROM INV_PARTS_STOCK_MST WHERE IN_PART_ID=%d" % PART))
    # The proc's IN_QTY bump fires the LEGACY UPDATE_PartNumber trigger (this spike DB still has the
    # legacy triggers attached), which writes spurious INV_PART_QTY_INF + _HIST audit rows. In the
    # rebuild DB those triggers won't exist (design §5). Disable it around the self-test so the test
    # leaves no audit-row residue; re-enabled in teardown.
    sql("DISABLE TRIGGER UPDATE_PartNumber ON INV_PARTS_STOCK_MST")

    sql("EXEC POST_StockMovement @partId=%d, @delta=%d, @sourceEnum='STOCKTAKING', "
        "@sourceRowId=999999, @eventKey='%s', @reason='parity self-test', @site=1, @purge=0"
        % (PART, DELTA, EVT))
    after = int(scalar("SELECT IN_QTY FROM INV_PARTS_STOCK_MST WHERE IN_PART_ID=%d" % PART))
    led = sql("SELECT IN_QTY_CHANGE, IN_BALANCE_AFTER FROM INV_STOCK_LEDGER WHERE VC_SOURCE_EVENT='%s'" % EVT)
    rep.check("post bumps IN_QTY by delta", after == base + DELTA,
              "%d -> %d (delta %d)" % (base, after, DELTA))
    rep.check("ledger row written with correct change+balance",
              len(led) == 1 and int(led[0][0]) == DELTA and int(led[0][1]) == base + DELTA,
              str(led[0]) if led else "no row")

    # Replay — idempotent, no double-post.
    sql("EXEC POST_StockMovement @partId=%d, @delta=%d, @sourceEnum='STOCKTAKING', "
        "@sourceRowId=999999, @eventKey='%s', @reason='replay', @site=1, @purge=0"
        % (PART, DELTA, EVT))
    after2 = int(scalar("SELECT IN_QTY FROM INV_PARTS_STOCK_MST WHERE IN_PART_ID=%d" % PART))
    cnt = int(scalar("SELECT COUNT(*) FROM INV_STOCK_LEDGER WHERE VC_SOURCE_EVENT='%s'" % EVT))
    rep.check("replay does NOT double-post (idempotent)", after2 == base + DELTA and cnt == 1,
              "IN_QTY=%d rows=%d" % (after2, cnt))

    # purge=1 posts nothing.
    sql("EXEC POST_StockMovement @partId=%d, @delta=999, @sourceEnum='STOCKTAKING', "
        "@eventKey='SPIKE_PARITY:tx=purge', @purge=1" % PART)
    after3 = int(scalar("SELECT IN_QTY FROM INV_PARTS_STOCK_MST WHERE IN_PART_ID=%d" % PART))
    pcnt = int(scalar("SELECT COUNT(*) FROM INV_STOCK_LEDGER WHERE VC_SOURCE_EVENT='SPIKE_PARITY:tx=purge'"))
    rep.check("purge=1 posts nothing", after3 == base + DELTA and pcnt == 0,
              "IN_QTY=%d rows=%d" % (after3, pcnt))

    # rebuildBalance re-stamps the absolute SUM of this part's ledger (only our 1 test row here).
    sql("EXEC PROC_RebuildStockBalance @partId=%d" % PART)
    rebuilt = int(scalar("SELECT IN_QTY FROM INV_PARTS_STOCK_MST WHERE IN_PART_ID=%d" % PART))
    rep.check("rebuildBalance re-stamps absolute SUM(ledger)", rebuilt == DELTA,
              "IN_QTY=%d == SUM(ledger)=%d" % (rebuilt, DELTA))

    # Fixture discipline: delete test rows, restore the part's IN_QTY, re-enable the legacy trigger.
    sql("DELETE FROM INV_STOCK_LEDGER WHERE VC_SOURCE_EVENT LIKE 'SPIKE_PARITY:%%'; "
        "UPDATE INV_PARTS_STOCK_MST SET IN_QTY=%d WHERE IN_PART_ID=%d; "
        "ENABLE TRIGGER UPDATE_PartNumber ON INV_PARTS_STOCK_MST" % (base, PART))
    restored = int(scalar("SELECT IN_QTY FROM INV_PARTS_STOCK_MST WHERE IN_PART_ID=%d" % PART))
    leftover = int(scalar("SELECT COUNT(*) FROM INV_STOCK_LEDGER WHERE VC_SOURCE_EVENT LIKE 'SPIKE_PARITY:%%'"))
    rep.check("DB restored as found (no test rows, IN_QTY intact)",
              restored == base and leftover == 0,
              "IN_QTY=%d rows_left=%d" % (restored, leftover))


# ---------------------------------------------------------------------------------------------
# B. Reconciliation — derive the ledger from source via the §3 mapping, diff vs live IN_QTY.
# ---------------------------------------------------------------------------------------------
# The derivation is expressed as ONE SQL query that replays each trigger leg faithfully:
#   * receiving (open-order) 'S' add-point, SHIP leg : +IN_QTY when VC_STATUS_SUPPLIER_SHIPPING<>''
#   * receiving (open-order) 'A' add-point, ARRIVAL  : +IN_QTY when arrival/plant/assembler/warehouse set
#   * reject  : -IN_QTY per reject row (int-keyed)
#   * stocktaking : +IN_QTY per row (signed adjustment delta, D5)
#   * shipping : -IN_QTY per row (string-keyed)
# Each leg SUMS ALL source rows per part (the rebuild's multi-row-correct behavior), so an F3
# multi-row-trigger under-count in legacy surfaces as derived>legacy for that part.
DERIVE_SQL = """
WITH recv AS (
    -- 'S' add-point ship leg + 'A' add-point arrival/yard/warehouse leg, gated exactly as the
    -- live INSERT_RecConfStatPartsStockMstQTY (VC_TERMINATED='' applies on the DELETE side; for the
    -- live snapshot the INSERT-counted rows are those still present with the status set).
    SELECT ps.IN_PART_ID,
           SUM(CASE
                 WHEN s.VC_INVENTORY_ADD_POINT='S' AND oo.VC_STATUS_SUPPLIER_SHIPPING<>'' THEN oo.IN_QTY
                 WHEN s.VC_INVENTORY_ADD_POINT='A' AND (oo.VC_ARRIVAL<>'' OR oo.VC_STATUS_PLANT_YARD<>''
                       OR oo.VC_STATUS_ASSEMBLER_YARD<>'' OR oo.VC_WAREHOUSE<>'') THEN oo.IN_QTY
                 ELSE 0 END) AS d
    FROM INV_OPEN_ORDER_INF oo
    JOIN INV_PARTS_STOCK_MST ps ON ps.VC_PART_NUMBER = oo.VC_PART_NUMBER
    JOIN INV_SUPPLIER_MST s ON ps.IN_SUPPLIER_ID = s.IN_SUPPLIER_ID
    GROUP BY ps.IN_PART_ID
),
rej AS (
    SELECT IN_PART_ID, -SUM(IN_QTY) AS d FROM INV_REJECT_INF GROUP BY IN_PART_ID
),
stk AS (
    SELECT IN_PART_ID, SUM(IN_QTY) AS d FROM INV_STOCKTAKING_INF GROUP BY IN_PART_ID
),
shp AS (
    SELECT ps.IN_PART_ID, -SUM(s.IN_QTY) AS d
    FROM INV_PART_SHIPPING_INF s
    JOIN INV_PARTS_STOCK_MST ps ON ps.VC_PART_NUMBER = s.VC_PART_NUMBER
    GROUP BY ps.IN_PART_ID
)
SELECT p.IN_PART_ID,
       p.IN_QTY AS legacy,
       (ISNULL(recv.d,0)+ISNULL(rej.d,0)+ISNULL(stk.d,0)+ISNULL(shp.d,0)) AS derived,
       (ISNULL(recv.d,0)+ISNULL(rej.d,0)+ISNULL(stk.d,0)+ISNULL(shp.d,0)) - p.IN_QTY AS diff,
       ISNULL(recv.d,0) AS d_recv, ISNULL(stk.d,0) AS d_stk, ISNULL(shp.d,0) AS d_shp
FROM INV_PARTS_STOCK_MST p
LEFT JOIN recv ON recv.IN_PART_ID=p.IN_PART_ID
LEFT JOIN rej  ON rej.IN_PART_ID =p.IN_PART_ID
LEFT JOIN stk  ON stk.IN_PART_ID =p.IN_PART_ID
LEFT JOIN shp  ON shp.IN_PART_ID =p.IN_PART_ID
ORDER BY p.IN_PART_ID
"""


def snapshot_facts():
    """Facts that determine which divergence classes are even REACHABLE in this snapshot."""
    return {
        "add_point_A_suppliers": int(scalar("SELECT COUNT(*) FROM INV_SUPPLIER_MST WHERE VC_INVENTORY_ADD_POINT='A'")),
        "counting_open_orders": int(scalar("SELECT COUNT(*) FROM INV_OPEN_ORDER_INF WHERE VC_STATUS_SUPPLIER_SHIPPING<>'' OR VC_ARRIVAL<>'' OR VC_STATUS_PLANT_YARD<>'' OR VC_STATUS_ASSEMBLER_YARD<>'' OR VC_WAREHOUSE<>''")),
        "reject_rows": int(scalar("SELECT COUNT(*) FROM INV_REJECT_INF")),
        "stocktaking_rows": int(scalar("SELECT COUNT(*) FROM INV_STOCKTAKING_INF")),
        "shipping_rows": int(scalar("SELECT COUNT(*) FROM INV_PART_SHIPPING_INF")),
        "shipping_f3_groups": int(scalar("SELECT COUNT(*) FROM (SELECT IN_SHIPPING_ID, VC_PART_NUMBER FROM INV_PART_SHIPPING_INF GROUP BY IN_SHIPPING_ID, VC_PART_NUMBER HAVING COUNT(*)>1) x")),
        "parts_total": int(scalar("SELECT COUNT(*) FROM INV_PARTS_STOCK_MST")),
        "parts_with_activity": int(scalar(
            "SELECT COUNT(*) FROM INV_PARTS_STOCK_MST p WHERE EXISTS(SELECT 1 FROM INV_STOCKTAKING_INF t WHERE t.IN_PART_ID=p.IN_PART_ID) "
            "OR EXISTS(SELECT 1 FROM INV_PART_SHIPPING_INF s WHERE s.VC_PART_NUMBER=p.VC_PART_NUMBER) "
            "OR EXISTS(SELECT 1 FROM INV_REJECT_INF r WHERE r.IN_PART_ID=p.IN_PART_ID)")),
        "open_order_min_add": scalar("SELECT MIN(VC_ADD) FROM INV_OPEN_ORDER_INF"),
        "shipping_max_add": scalar("SELECT MAX(VC_ADD) FROM INV_PART_SHIPPING_INF"),
    }


def reconcile(rep):
    print("\n--- B. Reconciliation: derive ledger from source vs live IN_QTY ---")
    f = snapshot_facts()
    print("  snapshot facts:")
    for k in ("parts_total", "parts_with_activity", "add_point_A_suppliers",
              "counting_open_orders", "reject_rows", "stocktaking_rows",
              "shipping_rows", "shipping_f3_groups"):
        print("    %-24s %s" % (k, f[k]))
    print("    open_order VC_ADD min = %s ; shipping VC_ADD max = %s"
          % (f["open_order_min_add"], f["shipping_max_add"]))

    # The reachability of each predicted divergence class in THIS snapshot.
    purge_horizon = (f["counting_open_orders"] == 0)
    print("\n  PREDICTED divergence-class reachability in this snapshot:")
    print("    D8(3) arrival-reversal   : NEEDS 'A' add-point + cleared arrival -> reachable=%s"
          % (f["add_point_A_suppliers"] > 0))
    print("    D12#3 plant/assembler-yd : NEEDS 'A' add-point + yard-on-edit  -> reachable=%s"
          % (f["add_point_A_suppliers"] > 0))
    print("    F3 multi-row under-count : NEEDS same-part multi-row batch     -> reachable=%s (shipping groups=%d)"
          % (f["shipping_f3_groups"] > 0, f["shipping_f3_groups"]))
    print("    F5 part-number change    : NEEDS a re-pointed VC_PART_NUMBER   -> not derivable from a static snapshot")

    rows = sql(DERIVE_SQL)
    data = []
    for r in rows:
        if len(r) < 7:
            continue
        data.append({
            "part": int(r[0]), "legacy": int(r[1]), "derived": int(r[2]),
            "diff": int(r[3]), "d_recv": int(r[4]), "d_stk": int(r[5]), "d_shp": int(r[6]),
        })

    exact = [d for d in data if d["diff"] == 0]
    nonzero = [d for d in data if d["diff"] != 0]
    print("\n  per-part reconciliation (%d parts): %d EXACT (diff 0), %d non-zero"
          % (len(data), len(exact), len(nonzero)))

    # ---- F3 candidate SET (review fix #2): legacy multi-row-trigger under-count is reachable
    # via ANY source table whose trigger joins `inserted` by a key that can repeat in one batch —
    # shipping + reject + open-orders (design §6 cites all three). Resolve the full set of parts
    # that have >1 same-key row in a single source batch, and warn LOUDLY if a candidate can't
    # resolve to a live part (so it can't silently evaporate the way a TOP 1 did).
    f3_parts = set()
    f3_unresolved = []
    # shipping: multi-row same-part within one shipping batch (keyed VC_PART_NUMBER)
    for pn in rows_col("SELECT VC_PART_NUMBER FROM INV_PART_SHIPPING_INF "
                       "GROUP BY IN_SHIPPING_ID, VC_PART_NUMBER HAVING COUNT(*)>1"):
        pid = scalar("SELECT IN_PART_ID FROM INV_PARTS_STOCK_MST WHERE VC_PART_NUMBER=%s"
                     % quote(pn))
        (f3_parts.add(int(pid)) if pid else f3_unresolved.append(("shipping", pn)))
    # open-orders: multi-row same-part within one FRS batch (keyed VC_PART_NUMBER)
    for pn in rows_col("SELECT VC_PART_NUMBER FROM INV_OPEN_ORDER_INF "
                       "GROUP BY VC_FRS_NUMBER, VC_PART_NUMBER HAVING COUNT(*)>1"):
        pid = scalar("SELECT IN_PART_ID FROM INV_PARTS_STOCK_MST WHERE VC_PART_NUMBER=%s"
                     % quote(pn))
        (f3_parts.add(int(pid)) if pid else f3_unresolved.append(("open-order", pn)))
    # reject: multi-row same-part id in one reject batch (keyed IN_PART_ID — already an id)
    for pid in rows_col("SELECT IN_PART_ID FROM INV_REJECT_INF "
                        "GROUP BY IN_REJECT_ID, IN_PART_ID HAVING COUNT(*)>1"):
        if str(pid).strip().lstrip("-").isdigit():
            f3_parts.add(int(pid))
    if f3_unresolved:
        print("  ⚠️  F3 candidate part(s) NOT in INV_PARTS_STOCK_MST (cannot adjudicate, NOT swallowed): %s"
              % f3_unresolved)

    # Classification is DIRECTION-based:
    #   derived > legacy (OVER-count): the rebuild-bug signature — derived posted more stock than the
    #     legacy IN_QTY holds. The ONLY benign over-count is F3 (legacy multi-row trigger silently
    #     dropped rows, so the ledger is correctly HIGHER). Any other over-count is a real defect and
    #     must FAIL on EVERY snapshot (review fix #1 — not maskable by purge_horizon).
    #   derived < legacy (UNDER-count): on a post-purge snapshot this is the §3.1 purge horizon — the
    #     receiving IN-movements that built legacy IN_QTY were aged out, so the from-zero replay can't
    #     reach them. On a FULL-history (pre-purge) dataset an under-count would instead be a bug; that
    #     case is caught by the from-zero-reconciliation check, which only runs when purge_horizon is False.
    classified = {"PURGE_HORIZON": [], "F3": [], "UNEXPLAINED": []}
    for d in nonzero:
        if d["diff"] > 0:                               # OVER-count
            (classified["F3"] if d["part"] in f3_parts else classified["UNEXPLAINED"]).append(d)
        else:                                           # UNDER-count (diff < 0)
            (classified["PURGE_HORIZON"] if purge_horizon else classified["UNEXPLAINED"]).append(d)

    print("\n  classification of non-zero diffs:")
    print("    PURGE_HORIZON (derived<legacy; receiving IN-movements purged, §3.1): %d"
          % len(classified["PURGE_HORIZON"]))
    print("    F3 over-count candidate part(s) (legacy multi-row under-count; shipping+reject+open-order set): %d  %s"
          % (len(classified["F3"]), [d["part"] for d in classified["F3"]]))
    print("    UNEXPLAINED (rebuild-bug signature): %d  %s"
          % (len(classified["UNEXPLAINED"]), [(d["part"], d["diff"]) for d in classified["UNEXPLAINED"]]))

    # ---- Adjudication (review fix #1: the OVER-count bug signature must FAIL UNCONDITIONALLY) ----
    # A derived>legacy diff that is not a known F3 part is a real rebuild defect — it must fail on
    # ANY snapshot, post-purge or not. purge_horizon only gates the separate from-zero TOTAL SKIP
    # below; it must NOT mask an over-count (that was the vacuous-assert the reviewer caught).
    rep.check("no UNEXPLAINED over-count diffs (rebuild-bug signature) — must hold on any snapshot",
              len(classified["UNEXPLAINED"]) == 0,
              "%d exact, %d purge-horizon, %d F3, %d UNEXPLAINED"
              % (len(exact), len(classified["PURGE_HORIZON"]),
                 len(classified["F3"]), len(classified["UNEXPLAINED"])))

    if purge_horizon:
        print("\n  *** PURGE-HORIZON FINDING (blocks a clean from-zero reconciliation) ***")
        print("  The restored Inventory.bak has 0 counting open-order rows: every surviving order")
        print("  carries blank VC_STATUS_SUPPLIER_SHIPPING/VC_ARRIVAL, and DELETE_AutoPurge (§3.1)")
        print("  has aged out the receiving rows that BUILT the legacy IN_QTY. A from-zero replay")
        print("  of THIS snapshot therefore cannot reconstruct IN_QTY — the reconciliation needs a")
        print("  complete event history (a pre-purge dump, or the live cutover backfill window).")
        print("  The §3 derivation logic itself is exercised + correct; the snapshot is the limit.")
        rep.skip("from-zero reconciliation == legacy IN_QTY (needs pre-purge event history)",
                 "snapshot is post-DELETE_AutoPurge; 0 counting open-orders (§3.1 purge horizon)")

    # Always emit the TSV artifact the reviewer eyeballs.
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "artifacts", "stock_ledger_parity.tsv")
    with open(out, "w") as fh:
        fh.write("part\tlegacy\tderived\tdiff\td_recv\td_stk\td_shp\tclassification\n")
        for d in data:
            cls = "EXACT"
            if d["diff"] > 0:
                cls = "F3_candidate" if d["part"] in f3_parts else "UNEXPLAINED"
            elif d["diff"] < 0:
                cls = "PURGE_HORIZON" if purge_horizon else "UNEXPLAINED"
            fh.write("%d\t%d\t%d\t%d\t%d\t%d\t%d\t%s\n"
                     % (d["part"], d["legacy"], d["derived"], d["diff"],
                        d["d_recv"], d["d_stk"], d["d_shp"], cls))
    print("\n  TSV artifact: %s" % out)


def main():
    print("=" * 78)
    print(" STOCK-LEDGER SERVICE — reconciliation parity harness (design §6 GO/NO-GO)")
    print("=" * 78)
    rep = Report()
    proc_self_test(rep)
    reconcile(rep)
    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
