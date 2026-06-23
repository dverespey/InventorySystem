#!/usr/bin/env python3
"""test_renban_e2e.py — drive the REAL renban.commit_renban_breakdown write-back end-to-end against the
live spike DB through jython_shim's PERSISTENT-SESSION transaction extension (the R8 seam-driver pattern).

The renban breakdown's commit is delete-then-reinsert + a counter bump, ALL in ONE transaction
(RenbanOrder.pas:416/436/473). An autocommit shim (a fresh `docker exec -Q` per call) physically cannot
carry a BEGIN TRAN across the DELETE_OrderRenban / INSERT_OpenOrder / UPDATE_RenbanGroupCount sequence,
so it could never prove the atomicity. The _TxSession holds one open connection so the ACTUAL
commit_renban_breakdown runs against a real transaction (not a SQL re-implementation).

Scenario (a real renban group + real grouped parts so the SELECT_OrderNoRenban join resolves):
  1. Seed BLANK-renban placeholder open orders for CMWA parts (FRS prefix '91231' -> far-future, unique).
  2. Drive the real commit_renban_breakdown(group, trailers, pallets) — it reads the blank orders
     (aliased loadBlankRenbanOrders), computes the trailer split, then DELETE-then-reINSERTs + bumps the
     group counter in one tx.
  3. Assert: every re-inserted row has a NON-BLANK renban (eligible for order-file gen) AND blank
     VC_ORDER_DATE; the per-trailer FRS (5-char prefix + 2-digit ordinal) + renban (group + %.3d) match
     what compute produced; the placeholders are GONE; the group counter advanced to next_count; on-hand
     IN_QTY did NOT move (status-empty rows -> triggers gated off); a re-run is a clean no-op (idempotent).

Far-future FRS prefix '91231' tags every test row. DB restored EXACTLY as-found in a finally, asserted
at the end. Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 scripts/e2e/test_renban_e2e.py
"""
import os, subprocess, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report          # noqa: E402
import jython_shim              # noqa: E402

CONTAINER = os.environ.get("CONTAINER", "mssql-spike")
SA_PASS = os.environ.get("SA_PASS")
DB = "Inventory"
GROUP = "CMWA"
FRSPREFIX = "91231"             # far-future FRS prefix (Y=9, MMDD=1231) -> placeholder FRS '9123101'
PLACEHOLDER_FRS = FRSPREFIX + "01"
RENBAN_PY = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                             "docs", "analysis", "order", "project-library", "renban", "code.py"))
AUTH_PY = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                           "docs", "analysis", "production-readiness", "project-library", "auth", "code.py"))

# SERVER-SIDE WRITE GATE (BLOCKER-2): the driver authorizes auth.requireWrite(session) before any write.
WRITE_SESSION = {"auth": {"user": {"userName": "op1", "roles": ["ProductionControl"]}}}


def sql(query):
    if not SA_PASS:
        sys.exit("export SA_PASS first")
    out = subprocess.check_output([
        "docker", "exec", CONTAINER, "/opt/mssql-tools18/bin/sqlcmd", "-C", "-S", "localhost",
        "-U", "sa", "-P", SA_PASS, "-d", DB, "-h", "-1", "-W", "-s", "\t",
        "-Q", "SET NOCOUNT ON; " + query], text=True)
    return [l.split("\t") for l in out.splitlines()
            if l.strip() and not l.startswith("(") and not l.startswith("Msg ")]


def scalar(q):
    r = sql(q)
    return r[0][0] if r else None


def q(s):
    return "'" + str(s).replace("'", "''") + "'"


def load_renban():
    """The REAL renban library with the shim's tx-capable `system` injected."""
    # Register the REAL auth module so the driver's `import auth as A; A.requireWrite(session)` resolves.
    if "auth" not in sys.modules:
        sys.modules["auth"] = jython_shim.load_wrapper("auth", AUTH_PY)
    return jython_shim.load_wrapper("renban_real", RENBAN_PY)


def seed_placeholders(parts):
    """Insert blank-renban placeholder open orders for the given (part, kanban, supplier, qty) tuples,
    directly (a raw INSERT, NOT INSERT_OpenOrder, so the FRS is exactly our tagged placeholder and no
    server FRS-derivation interferes with the seed). Status fields blank -> stock-neutral."""
    for (part, kanban, sup, qty) in parts:
        sql("INSERT INTO INV_OPEN_ORDER_INF "
            "(VC_SUPPLIER_CODE, VC_PART_NUMBER, VC_KANBAN_NUMBER, VC_FRS_NUMBER, VC_RENBAN_NUMBER, "
            " IN_QTY, VC_ORDER_DATE, VC_FRS_DATE, VC_ADD) VALUES "
            "(%s, %s, %s, %s, '', %d, '', %s, '20261231000000')"
            % (q(sup), q(part), q(kanban), q(PLACEHOLDER_FRS), qty, q("2026" + FRSPREFIX[1:])))


def cleanup():
    sql("DELETE FROM INV_OPEN_ORDER_INF WHERE VC_FRS_NUMBER LIKE '%s%%'; "
        "DELETE FROM INV_OPEN_ORDER_INF_HIST WHERE VC_FRS_NUMBER LIKE '%s%%'" % (FRSPREFIX, FRSPREFIX))
    # P4: the clean-wrap override path raises a RENBAN_COLLISION alarm. Sweep any test-raised alarm rows
    # (this suite only writes RENBAN_COLLISION alarms for the CMWA group via the override scenario above).
    sql("DELETE FROM INV_EDI_ALARM_REJ WHERE VC_ALARM_TYPE='RENBAN_COLLISION' "
        "AND VC_MANIFEST_NUMBER LIKE '%s%%'" % GROUP)


def main():
    print("=" * 78)
    print(" RENBAN BREAKDOWN E2E — REAL commit_renban_breakdown via the shim's persistent-session tx")
    print("=" * 78)
    rep = Report()
    if not SA_PASS:
        sys.exit("export SA_PASS first")

    renban = load_renban()
    rep.check("real renban module loads under the shim (compute + commit present)",
              hasattr(renban, "commit_renban_breakdown") and hasattr(renban, "compute_trailer_breakdown"))

    # real CMWA parts (proved on the spike): lot qtys 40/30/40. Use 3 distinct parts so a truck carries
    # multiple part numbers (one-renban-per-trailer / multi-part-per-trailer, the domain shape).
    parts_meta = [("4261102Q4000", "15D", "0572B", 40),    # lotqty 40 -> 200 = 5 lots
                  ("4261102Q5000", "16E", "0572B", 30),    # lotqty 30 -> 300 = 10 lots
                  ("4261102Q8000", "M1",  "0572B", 40)]    # lotqty 40 -> 360 = 9 lots
    # order qtys chosen as exact lot multiples (the only thing this stage ever sees): 5 + 10 + 9 = 24 lots
    seed = [("4261102Q4000", "15D", "0572B", 200),
            ("4261102Q5000", "16E", "0572B", 300),
            ("4261102Q8000", "M1",  "0572B", 360)]

    group_count0 = scalar("SELECT VC_RENBAN_GROUP_COUNT FROM INV_RENBAN_GROUP_MST WHERE VC_RENBAN_GROUP_CODE='%s'" % GROUP)
    onhand0 = {p[0]: int(scalar("SELECT IN_QTY FROM INV_PARTS_STOCK_MST WHERE VC_PART_NUMBER=%s" % q(p[0])))
               for p in parts_meta}
    print("  group %s count=%s ; on-hand %s" % (GROUP, group_count0, onhand0))
    seed3 = int(str(group_count0)[-3:])

    # pre-clean any stray tagged rows, then make sure CMWA has no OTHER blank-renban rows that our
    # breakdown would sweep up (steady state proved = 0; assert it so the test is isolated).
    cleanup()
    other_blank = int(scalar(
        "SELECT COUNT(*) FROM INV_OPEN_ORDER_INF o JOIN INV_PARTS_STOCK_MST p ON o.VC_PART_NUMBER=p.VC_PART_NUMBER "
        "JOIN INV_RENBAN_GROUP_MST r ON r.IN_RENBAN_ID=p.IN_RENBAN_ID AND r.VC_RENBAN_GROUP_CODE='%s' "
        "WHERE o.VC_RENBAN_NUMBER=''" % GROUP))
    rep.check("isolated: CMWA has no pre-existing blank-renban orders (steady state) before seeding",
              other_blank == 0, "blank=%d" % other_blank)

    try:
        seed_placeholders(seed)
        seeded = int(scalar("SELECT COUNT(*) FROM INV_OPEN_ORDER_INF WHERE VC_FRS_NUMBER='%s' AND VC_RENBAN_NUMBER=''"
                            % PLACEHOLDER_FRS))
        rep.check("seeded 3 blank-renban CMWA placeholders", seeded == 3, "seeded=%d" % seeded)

        # ---- drive the REAL commit through the tx session: 3 trailers x 10 pallets (cap 30 >= 24) ----
        renban.system = jython_shim._System()
        TRAILERS, PALLETS = 3, 10
        result = renban.commit_renban_breakdown(GROUP, TRAILERS, PALLETS, DB, session=WRITE_SESSION)
        rep.check("commit returned a compute result with trailer rows + next_count",
                  result["rows"] and result["next_count"] is not None,
                  "rows=%d next_count=%s total_lots=%d" % (len(result["rows"]), result["next_count"], result["total_lots"]))
        rep.check("total_lots = 5 + 10 + 9 = 24", result["total_lots"] == 24, str(result["total_lots"]))

        # all placeholders deleted (blank-renban rows for the part are gone)
        remaining_blank = int(scalar("SELECT COUNT(*) FROM INV_OPEN_ORDER_INF WHERE VC_FRS_NUMBER='%s' AND VC_RENBAN_NUMBER=''"
                                     % PLACEHOLDER_FRS))
        rep.check("delete-then-reinsert: all blank placeholders REMOVED (remaining_blank=0)",
                  remaining_blank == 0, "remaining=%d" % remaining_blank)

        # the re-inserted rows: every one has a NON-BLANK renban and blank order-date (eligible for emit)
        db_rows = sql("SELECT VC_FRS_NUMBER, VC_RENBAN_NUMBER, VC_PART_NUMBER, IN_QTY, "
                      "  CASE WHEN VC_ORDER_DATE IS NULL OR VC_ORDER_DATE='' THEN 1 ELSE 0 END AS ordblank, "
                      "  CASE WHEN VC_STATUS_SUPPLIER_SHIPPING IS NULL OR VC_STATUS_SUPPLIER_SHIPPING='' THEN 1 ELSE 0 END AS shipblank "
                      "FROM INV_OPEN_ORDER_INF WHERE VC_FRS_NUMBER LIKE '%s%%' ORDER BY VC_FRS_NUMBER, VC_PART_NUMBER"
                      % FRSPREFIX)
        all_renban_nonblank = all(r[1].strip() != "" for r in db_rows) and len(db_rows) > 0
        rep.check("every re-inserted row has a NON-BLANK renban (eligible for SELECT_OrderNotOrdered)",
                  all_renban_nonblank, "renbans=%s" % [r[1] for r in db_rows])
        rep.check("every re-inserted row has blank VC_ORDER_DATE (the other emit gate)",
                  all(r[4] == "1" for r in db_rows))
        rep.check("every re-inserted row is status-empty (stock-neutral precondition)",
                  all(r[5] == "1" for r in db_rows))

        # the per-trailer FRS + renban + qty in the DB == what compute produced (the .pas read-out)
        db_tuples = sorted((r[0], r[1], r[2], int(r[3])) for r in db_rows)
        want_tuples = sorted((row["frs"], row["renban"], row["part"], row["qty"]) for row in result["rows"])
        rep.check("DB (FRS, renban, part, qty) per trailer == compute output (proven vs the .pas)",
                  db_tuples == want_tuples, "db=%s want=%s" % (db_tuples[:3], want_tuples[:3]))

        # FRS suffixes are the trailer ordinals 01/02/03 (5-char prefix + zero-padded ordinal)
        frs_suffixes = sorted(set(r[0][-2:] for r in db_rows))
        rep.check("FRS trailer suffixes are 01/02/03 (3 trailers, prefix %s)" % FRSPREFIX,
                  frs_suffixes == ["01", "02", "03"], str(frs_suffixes))
        # renban numbers are CMWA<seed3>, <seed3+1>, <seed3+2>
        want_renbans = sorted({"%s%03d" % (GROUP, seed3 + k) for k in range(TRAILERS)})
        got_renbans = sorted(set(r[1] for r in db_rows))
        rep.check("renban numbers are %s (group + seed%d..seed+%d)" % (want_renbans, seed3, TRAILERS - 1),
                  got_renbans == want_renbans, "got=%s" % got_renbans)

        # group counter advanced to next_count (= last rcount + 1 = seed3 + 3)
        cnt_after = scalar("SELECT VC_RENBAN_GROUP_COUNT FROM INV_RENBAN_GROUP_MST WHERE VC_RENBAN_GROUP_CODE='%s'" % GROUP)
        rep.check("group counter advanced to next_count (%s -> %s, = seed3+3=%d)"
                  % (group_count0, cnt_after, seed3 + 3),
                  int(cnt_after) == result["next_count"] and int(cnt_after) == seed3 + 3,
                  "cnt_after=%s next_count=%s" % (cnt_after, result["next_count"]))

        # stock neutrality: on-hand IN_QTY for every part unchanged
        neutral = all(int(scalar("SELECT IN_QTY FROM INV_PARTS_STOCK_MST WHERE VC_PART_NUMBER=%s" % q(p))) == onhand0[p]
                      for p in onhand0)
        rep.check("STOCK NEUTRAL: on-hand IN_QTY unchanged for every part (status-empty rows)", neutral)

        # idempotency: re-run the breakdown -> no blank placeholders left -> a clean no-op (no new rows,
        # counter NOT advanced again). Mirrors the legacy "No records for this Renban Group" (:648).
        cnt_before_rerun = cnt_after
        rows_before_rerun = int(scalar("SELECT COUNT(*) FROM INV_OPEN_ORDER_INF WHERE VC_FRS_NUMBER LIKE '%s%%'" % FRSPREFIX))
        renban.system = jython_shim._System()
        rerun = renban.commit_renban_breakdown(GROUP, TRAILERS, PALLETS, DB, session=WRITE_SESSION)
        rows_after_rerun = int(scalar("SELECT COUNT(*) FROM INV_OPEN_ORDER_INF WHERE VC_FRS_NUMBER LIKE '%s%%'" % FRSPREFIX))
        cnt_after_rerun = scalar("SELECT VC_RENBAN_GROUP_COUNT FROM INV_RENBAN_GROUP_MST WHERE VC_RENBAN_GROUP_CODE='%s'" % GROUP)
        rep.check("idempotent re-run: no blank orders left -> no-op (no new rows, counter unchanged)",
                  rerun["rows"] == [] and rows_after_rerun == rows_before_rerun and cnt_after_rerun == cnt_before_rerun,
                  "rerun_rows=%d before=%d after=%d cnt=%s->%s"
                  % (len(rerun["rows"]), rows_before_rerun, rows_after_rerun, cnt_before_rerun, cnt_after_rerun))

        # ================================================================================
        # SHOULD-FIX 2 — PARTIAL-LOT (lots=0) placeholder SURVIVES (delete scope = emitted rows).
        # Seed a part with 0 < qty < lotqty (lots = qty div lotqty = 0 -> lands on NO truck -> emits NO
        # trailer row) alongside a normal part. The legacy NEVER deletes the lots=0 part's placeholder
        # (its commit loop walks only the emitted grid rows, RenbanOrder.pas:417/506); the OLD rebuild
        # (parts_seen from the full feed) WOULD have deleted it -> silent order loss. Assert it SURVIVES.
        # ================================================================================
        cleanup()
        PART_ZERO = "4261102Q4000"      # lotqty 40 -> qty 20 = 0 lots (partial)
        PART_NORM = "4261102Q5000"      # lotqty 30 -> qty 300 = 10 lots (normal)
        seed_placeholders([(PART_ZERO, "15D", "0572B", 20),       # 20 < 40 -> lots 0
                           (PART_NORM, "16E", "0572B", 300)])     # 300/30 = 10 lots
        seeded2 = int(scalar("SELECT COUNT(*) FROM INV_OPEN_ORDER_INF WHERE VC_FRS_NUMBER='%s' AND VC_RENBAN_NUMBER=''"
                             % PLACEHOLDER_FRS))
        rep.check("partial-lot: seeded 2 blank placeholders (1 sub-lot + 1 normal)", seeded2 == 2,
                  "seeded=%d" % seeded2)
        cnt_before_pl = scalar("SELECT VC_RENBAN_GROUP_COUNT FROM INV_RENBAN_GROUP_MST WHERE VC_RENBAN_GROUP_CODE='%s'" % GROUP)

        renban.system = jython_shim._System()
        res_pl = renban.commit_renban_breakdown(GROUP, 3, 10, DB, session=WRITE_SESSION)   # cap 30 >= 10 lots
        rep.check("partial-lot: total_lots = 10 (the sub-lot part contributes 0)",
                  res_pl["total_lots"] == 10, str(res_pl["total_lots"]))
        rep.check("partial-lot: the sub-lot part emits NO trailer row (only the normal part is grouped)",
                  all(r["part"] == PART_NORM for r in res_pl["rows"]) and len(res_pl["rows"]) > 0,
                  "row_parts=%s" % sorted(set(r["part"] for r in res_pl["rows"])))

        # the sub-lot part's BLANK placeholder must STILL EXIST (not deleted, not re-inserted)
        zero_blank = int(scalar("SELECT COUNT(*) FROM INV_OPEN_ORDER_INF WHERE VC_PART_NUMBER=%s "
                                "AND VC_FRS_NUMBER='%s' AND VC_RENBAN_NUMBER=''" % (q(PART_ZERO), PLACEHOLDER_FRS)))
        rep.check("partial-lot: the lots=0 part's BLANK placeholder SURVIVES (legacy preserves it)",
                  zero_blank == 1, "zero_blank=%d (was deleted -> ORDER LOSS under the OLD full-feed bug)" % zero_blank)
        zero_total = int(scalar("SELECT COUNT(*) FROM INV_OPEN_ORDER_INF WHERE VC_PART_NUMBER=%s "
                                "AND VC_FRS_NUMBER LIKE '%s%%'" % (q(PART_ZERO), FRSPREFIX)))
        rep.check("partial-lot: the sub-lot part has EXACTLY its 1 untouched placeholder (no grouped rows)",
                  zero_total == 1, "zero_total=%d" % zero_total)
        # its qty is intact (the order is not lost)
        zero_qty = scalar("SELECT IN_QTY FROM INV_OPEN_ORDER_INF WHERE VC_PART_NUMBER=%s AND VC_FRS_NUMBER='%s' "
                          "AND VC_RENBAN_NUMBER=''" % (q(PART_ZERO), PLACEHOLDER_FRS))
        rep.check("partial-lot: the surviving placeholder still carries its qty (20) — order intact",
                  int(zero_qty) == 20, "qty=%s" % zero_qty)

        # the NORMAL part is grouped: its blank placeholder is gone, replaced by non-blank renban rows
        norm_blank = int(scalar("SELECT COUNT(*) FROM INV_OPEN_ORDER_INF WHERE VC_PART_NUMBER=%s "
                                "AND VC_FRS_NUMBER LIKE '%s%%' AND VC_RENBAN_NUMBER=''" % (q(PART_NORM), FRSPREFIX)))
        norm_grouped = int(scalar("SELECT COUNT(*) FROM INV_OPEN_ORDER_INF WHERE VC_PART_NUMBER=%s "
                                  "AND VC_FRS_NUMBER LIKE '%s%%' AND VC_RENBAN_NUMBER<>''" % (q(PART_NORM), FRSPREFIX)))
        rep.check("partial-lot: the NORMAL part is grouped (blank placeholder gone, non-blank renban rows present)",
                  norm_blank == 0 and norm_grouped == len([r for r in res_pl["rows"] if r["part"] == PART_NORM])
                  and norm_grouped > 0,
                  "norm_blank=%d norm_grouped=%d" % (norm_blank, norm_grouped))
        # DELETE_OrderRenban was issued ONLY for the emitted (normal) part, never for the sub-lot part
        # (proven by the survival above). The counter advanced for the normal part's trailers only.
        cnt_after_pl = scalar("SELECT VC_RENBAN_GROUP_COUNT FROM INV_RENBAN_GROUP_MST WHERE VC_RENBAN_GROUP_CODE='%s'" % GROUP)
        rep.check("partial-lot: counter advanced to next_count (from the emitted normal-part trailers only)",
                  int(cnt_after_pl) == res_pl["next_count"],
                  "cnt %s->%s next_count=%s" % (cnt_before_pl, cnt_after_pl, res_pl["next_count"]))

        # ================================================================================
        # P4 CLEAN WRAP (999->000) — DB PERSIST + ring-wrapped renban strings (replaces BLOCKER-1 parity).
        # Set the group count to 997, drive 5 trailers x 2 pallets with a 5-lot part -> RAW rcounts
        # 997,998,999,1000,1001 -> next_count = 1002. The renban STRINGS ring-wrap: CMWA997/998/999/000/001
        # (NO 4-digit CMWA1000). On the REAL CMWA group CMWA997/998/999 are resident (proved on the spike),
        # so resolution=None CORRECTLY returns COLLISION (the allocator fires); we then OVERRIDE to force the
        # write and read back the PERSISTED count = '002' (the clean ring wrap 1002 % 1000), NOT the
        # pre-cutover '100'. Expected values derived from spec §12.7 ring math (NOT the rebuild). The persist-
        # step non-vacuity lives HERE: revert step (c) -> the DB stores '100' -> this check flips red.
        # ================================================================================
        cleanup()
        ROLL_PART = "4261102Q5000"      # lotqty 30 -> qty 150 = 5 lots
        sql("EXEC UPDATE_RenbanGroupCount @RenbanCode=%s, @RenbanCount='997'" % q(GROUP))   # restored in finally
        seed_placeholders([(ROLL_PART, "16E", "0572B", 150)])     # 150/30 = 5 lots
        renban.system = jython_shim._System()
        # resolution=None: the candidates straddle 999 onto resident CMWA997/998/999 -> COLLISION (no write).
        res_roll = renban.commit_renban_breakdown(GROUP, 5, 2, DB, session=WRITE_SESSION)
        rep.check("clean-wrap rollover: next_count carries the RAW count across 1000 (seed 997 -> 1002)",
                  res_roll["next_count"] == 1002, "next_count=%s" % res_roll["next_count"])
        rep.check("clean-wrap rollover: resolution=None DETECTS the collision (CMWA997/998/999 resident) -> no write",
                  res_roll["status"] == "COLLISION" and res_roll.get("alarm_id"),
                  "status=%s collisions=%s alarm=%s" % (res_roll["status"],
                  [c["renban"] for c in res_roll["collisions"]], res_roll.get("alarm_id")))
        cnt_after_warn = scalar("SELECT VC_RENBAN_GROUP_COUNT FROM INV_RENBAN_GROUP_MST WHERE VC_RENBAN_GROUP_CODE='%s'" % GROUP)
        rep.check("clean-wrap rollover: COLLISION did NOT write (counter still 997, no trailer rows)",
                  cnt_after_warn == "997"
                  and int(scalar("SELECT COUNT(*) FROM INV_OPEN_ORDER_INF WHERE VC_FRS_NUMBER LIKE '%s%%' AND VC_RENBAN_NUMBER<>''" % FRSPREFIX)) == 0,
                  "cnt=%s" % cnt_after_warn)
        # OVERRIDE: force the write through (the candidates ring-wrap to CMWA997/998/999/000/001) + read the
        # PERSISTED count. (Override re-checks in-tx; 997/998/999 are STILL resident but override is the
        # operator's explicit decision to reuse — the in-tx re-check guards a NEWLY-taken number, not this
        # deliberate reuse.) Expected DB persist = '002' (1002 % 1000), NOT the pre-cutover '100'.
        # SHOULD-FIX-2 CONTRACT: the dialog passes back the WARN's collision set as `acknowledged` (the
        # production override script does `acknowledged=list(p.acknowledged)`); we mirror that here so the
        # deliberately-reused 997/998/999 are excluded from the in-tx re-check. (The None default now fails
        # CLOSED — acknowledges nothing — so this MUST pass the explicit set, exactly as the UI does.)
        renban.system = jython_shim._System()
        ack_roll = [c["renban"] for c in res_roll["collisions"]]   # the WARN payload the dialog returns
        res_ovr = renban.commit_renban_breakdown(GROUP, 5, 2, DB,
                                                 resolution={"action": "override", "alarm_id": res_roll["alarm_id"],
                                                             "acknowledged": ack_roll},
                                                 actor="e2e", session=WRITE_SESSION)
        rep.check("clean-wrap rollover: override COMMITTED the breakdown", res_ovr["status"] == "COMMITTED",
                  "status=%s" % res_ovr["status"])
        cnt_roll = scalar("SELECT VC_RENBAN_GROUP_COUNT FROM INV_RENBAN_GROUP_MST WHERE VC_RENBAN_GROUP_CODE='%s'" % GROUP)
        rep.check("clean-wrap rollover: DB persists '002' (ring wrap 1002 %% 1000), NOT the pre-cutover '100'",
                  cnt_roll == "002", "persisted=%r (pre-cutover truncation would store '100')" % cnt_roll)
        # the renban STRINGS in the DB ring-wrap: CMWA997/998/999/000/001 — NO 4-digit CMWA1000.
        roll_renbans = sql("SELECT DISTINCT VC_RENBAN_NUMBER FROM INV_OPEN_ORDER_INF "
                           "WHERE VC_PART_NUMBER=%s AND VC_FRS_NUMBER LIKE '%s%%' AND VC_RENBAN_NUMBER<>'' "
                           "ORDER BY VC_RENBAN_NUMBER" % (q(ROLL_PART), FRSPREFIX))
        got_roll = sorted(r[0] for r in roll_renbans)
        rep.check("clean-wrap rollover: renban STRINGS ring-wrap CMWA997/998/999/000/001 (NO 4-digit CMWA1000)",
                  got_roll == ["CMWA000", "CMWA001", "CMWA997", "CMWA998", "CMWA999"] and "CMWA1000" not in got_roll,
                  "got=%s" % got_roll)

    finally:
        cleanup()
        # restore the group counter (the commit advanced it) so the spike is left EXACTLY as-found
        sql("EXEC UPDATE_RenbanGroupCount @RenbanCode=%s, @RenbanCount=%s" % (q(GROUP), q(group_count0)))

    leftover = int(scalar("SELECT COUNT(*) FROM INV_OPEN_ORDER_INF WHERE VC_FRS_NUMBER LIKE '%s%%'" % FRSPREFIX))
    leftover_hist = int(scalar("SELECT COUNT(*) FROM INV_OPEN_ORDER_INF_HIST WHERE VC_FRS_NUMBER LIKE '%s%%'" % FRSPREFIX))
    leftover_alarm = int(scalar("SELECT COUNT(*) FROM INV_EDI_ALARM_REJ WHERE VC_ALARM_TYPE='RENBAN_COLLISION'"))
    rep.check("spike restored as-found: no leftover RENBAN_COLLISION alarm rows", leftover_alarm == 0,
              "alarm_rows=%d" % leftover_alarm)
    cnt_restored = scalar("SELECT VC_RENBAN_GROUP_COUNT FROM INV_RENBAN_GROUP_MST WHERE VC_RENBAN_GROUP_CODE='%s'" % GROUP)
    onhand_restored = all(int(scalar("SELECT IN_QTY FROM INV_PARTS_STOCK_MST WHERE VC_PART_NUMBER=%s" % q(p))) == onhand0[p]
                          for p in onhand0)
    rep.check("spike restored as-found (no test rows/hist, counter + on-hand back to baseline)",
              leftover == 0 and leftover_hist == 0 and cnt_restored == group_count0 and onhand_restored,
              "orders=%d hist=%d count=%s(want %s) onhand_ok=%s"
              % (leftover, leftover_hist, cnt_restored, group_count0, onhand_restored))
    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
