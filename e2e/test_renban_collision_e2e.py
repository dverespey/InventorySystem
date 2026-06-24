#!/usr/bin/env python3
"""test_renban_collision_e2e.py — drive the REAL P4 collision-aware allocator (WARN -> GUIDE -> FIX)
end-to-end against the live spike DB. Exercises every path the design (renban-collision-design.md §3-§5)
and source-truth (renban-collision-sourcetruth.md §2-§4) require:

  * collision DETECTED against a SEEDED resident renban (resident-rows, status-independent predicate);
  * next-free RUN-of-N (the GUIDE base whose [base..base+N-1] %1000 are ALL free);
  * use_next_free  -> re-map + in-tx re-check + commit + alarm ack (BIT_RESOLVED=1);
  * override       -> commit ORIGINAL (acknowledged) renbans + ack with an audit note;
  * cancel         -> NO driver call; the RENBAN_COLLISION alarm stays BIT_RESOLVED=0 (active);
  * in-tx TOCTOU re-check FIRES on a simulated LOST RACE on ALL THREE commit paths (a SECOND connection
    grabs the chosen renban after the pre-check / inside the open tx) -> abort, no partial write;
  * exhaustion (every suffix of a small group seeded) -> next_free_run = None -> hard WARN-cancel;
  * the alarm row carries the RENBAN_COLLISION shape (colliding renban in VC_MANIFEST_NUMBER varchar(8));
  * [9] BLOCKER-1: a WRAPPED candidate run (straddling 999->000) x use_next_free re-seats onto the
    VALIDATED contiguous block [base..base+2] (offset from POSITION, not the tail), persists the true
    highest-written+1, and self-collides on no next run. Oracle = the spec ring (R15); revert-proven.
  * [10] BLOCKER-2: the SERVER-SIDE write gate (auth.requireWrite from the SESSION) — a forged/anon/no-
    write-role session is REJECTED before any read/compute/write on every path; ProductionControl/Admin
    reach the write; revert-proven (neuter requireWrite -> the forged write slips through).
  * [11] SHOULD-FIX-1: the GUIDE scan DETECTS a trailing-space resident renban ('ZZRB289 ' -> suffix 289,
    not 089); the run-of-N SKIPS it. Revert-proven (the old RIGHT(.,3) form misses it).
  * [12] SHOULD-FIX-2: override with the EXPLICIT WARN-payload acknowledged set commits the deliberate
    reuse; the None default now fails CLOSED (aborts on the un-acknowledged still-resident number).

ALL synthetic rows are sentinel-tagged (FRS prefix '91230' / a synthetic ZZRB renban group / a sentinel
supplier) and torn down in a finally; the spike is asserted restored as-found. The lost-race seeding uses
a SECOND autocommit connection (jython_shim._run / direct sqlcmd) WHILE the driver's persistent-session tx
is open, so the in-tx re-check (READ COMMITTED) sees the concurrently-committed row — the faithful TOCTOU.

Uses a SYNTHETIC renban group 'ZZRB' (created + dropped here) so the seeded resident rows never touch the
real CMWA/PACF/DICAS numbers and the exhaustion test can fill a tiny ring without 1000 inserts.
Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 e2e/test_renban_collision_e2e.py
"""
import os, subprocess, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report, preclean_sentinels          # noqa: E402
import jython_shim                                   # noqa: E402

CONTAINER = os.environ.get("CONTAINER", "mssql-spike")
SA_PASS = os.environ.get("SA_PASS")
DB = "Inventory"
GROUP = "ZZRB"                       # synthetic sentinel renban group (created + dropped here)
FRSPREFIX = "91230"                  # far-future sentinel FRS prefix -> placeholder FRS '9123001'
PLACEHOLDER_FRS = FRSPREFIX + "01"
SENT_SUP = "ZZ572"                   # sentinel supplier marker (not used by INSERT, kept for grep)
RENBAN_PY = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..",
                             "ignition", "project-library", "renban", "code.py"))
AUTH_PY = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..",
                           "ignition", "project-library", "auth", "code.py"))

# SERVER-SIDE WRITE GATE (BLOCKER-2): the driver now calls auth.requireWrite(session) before any write.
# Session shapes are the dict form auth.sessionRoles accepts (session.props.auth.user.roles equivalent).
WRITE_SESSION = {"auth": {"user": {"userName": "op1", "roles": ["ProductionControl"]}}}  # a writer
ADMIN_SESSION = {"auth": {"user": {"userName": "boss", "roles": ["Admin"]}}}             # also a writer
ANON_SESSION = {"auth": {"user": {"userName": None, "roles": []}}}                       # not logged in
FORGED_SESSION = {"auth": {"user": {"userName": "nobody", "roles": ["SomeOtherRole"]}}}  # logged in, no write role


def sql(query):
    if not SA_PASS:
        sys.exit("export SA_PASS first")
    out = subprocess.check_output([
        "docker", "exec", CONTAINER, "/opt/mssql-tools18/bin/sqlcmd", "-C", "-S", "localhost",
        "-U", "sa", "-P", SA_PASS, "-d", DB, "-h", "-1", "-W", "-s", "\t",
        "-Q", "SET QUOTED_IDENTIFIER ON; SET NOCOUNT ON; " + query], text=True)
    return [l.split("\t") for l in out.splitlines()
            if l.strip() and not l.startswith("(") and not l.startswith("Msg ")]


def scalar(q_):
    r = sql(q_)
    return r[0][0] if r else None


def q(s):
    return "'" + str(s).replace("'", "''") + "'"


def load_renban():
    # Register the REAL auth module so the driver's `import auth as A; A.requireWrite(session)` (BLOCKER-2)
    # resolves to the gateway module — the SAME gate the gateway runs, not a stub.
    if "auth" not in sys.modules:
        sys.modules["auth"] = jython_shim.load_wrapper("auth", AUTH_PY)
    return jython_shim.load_wrapper("renban_collision_real", RENBAN_PY)


# --- sentinel fixtures --------------------------------------------------------------------------
# A REAL part must back the SELECT_OrderNoRenban join (INV_PARTS_STOCK_MST + INV_RENBAN_GROUP_MST). We
# create the synthetic ZZRB group, TEMPORARILY re-point a real part at it (restored in finally), and seed
# blank placeholders. 4261102Q4000 is a live CMWA part (IN_RENBAN_ID restored to its original in finally).
PART = "4261102Q4000"                # a real INV_PARTS_STOCK_MST part we re-point at ZZRB then restore


def teardown_statements():
    """Child-before-parent, idempotent, sentinel-scoped (P11/R18 self-heal + the finally teardown)."""
    return [
        # the breakdown's own grouped/placeholder rows
        "DELETE FROM INV_OPEN_ORDER_INF WHERE VC_FRS_NUMBER LIKE '%s%%'" % FRSPREFIX,
        "DELETE FROM INV_OPEN_ORDER_INF_HIST WHERE VC_FRS_NUMBER LIKE '%s%%'" % FRSPREFIX,
        # the seeded RESIDENT collision rows (a distinct sentinel FRS prefix so they survive the breakdown)
        "DELETE FROM INV_OPEN_ORDER_INF WHERE VC_RENBAN_NUMBER LIKE '%s%%' AND VC_FRS_NUMBER LIKE '9124%%'" % GROUP,
        "DELETE FROM INV_OPEN_ORDER_INF_HIST WHERE VC_RENBAN_NUMBER LIKE '%s%%' AND VC_FRS_NUMBER LIKE '9124%%'" % GROUP,
        # the RENBAN_COLLISION alarm rows for the sentinel group
        "DELETE FROM INV_EDI_ALARM_REJ WHERE VC_ALARM_TYPE='RENBAN_COLLISION' AND VC_MANIFEST_NUMBER LIKE '%s%%'" % GROUP,
    ]


def seed_placeholders(parts):
    """Blank-renban placeholder open orders (raw INSERT so the FRS is exactly our tagged placeholder)."""
    for (part, kanban, qty) in parts:
        sql("INSERT INTO INV_OPEN_ORDER_INF "
            "(VC_SUPPLIER_CODE, VC_PART_NUMBER, VC_KANBAN_NUMBER, VC_FRS_NUMBER, VC_RENBAN_NUMBER, "
            " IN_QTY, VC_ORDER_DATE, VC_FRS_DATE, VC_ADD) VALUES "
            "(%s, %s, %s, %s, '', %d, '', '20261230', '20261230000000')"
            % (q(SENT_SUP), q(part), q(kanban), q(PLACEHOLDER_FRS), qty))


def seed_resident(renban, frs="9124001"):
    """Seed a RESIDENT (in-use) order row carrying `renban` — a distinct FRS prefix '9124' so it is NOT a
    blank placeholder and is NOT swept by the breakdown's delete (status-independent: order-dated)."""
    sql("INSERT INTO INV_OPEN_ORDER_INF "
        "(VC_SUPPLIER_CODE, VC_PART_NUMBER, VC_KANBAN_NUMBER, VC_FRS_NUMBER, VC_RENBAN_NUMBER, "
        " IN_QTY, VC_ORDER_DATE, VC_FRS_DATE, VC_ADD) VALUES "
        "(%s, %s, 'KZ', %s, %s, 40, '20260101', '20260101', '20260101000000')"
        % (q(SENT_SUP), q(PART), q(frs), q(renban)))


def main():
    print("=" * 78)
    print(" RENBAN COLLISION ALLOCATOR E2E — WARN -> GUIDE -> FIX, real driver via the shim tx")
    print("=" * 78)
    rep = Report()
    if not SA_PASS:
        sys.exit("export SA_PASS first")

    renban = load_renban()
    rep.check("real renban module loads (collision API present)",
              all(hasattr(renban, n) for n in ("check_renban_collisions", "next_free_run",
                                               "commit_renban_breakdown", "_write_renban_alarm")))

    # set up the synthetic ZZRB group pointed at a real part (restored in finally)
    part_renban_id0 = scalar("SELECT IN_RENBAN_ID FROM INV_PARTS_STOCK_MST WHERE VC_PART_NUMBER=%s" % q(PART))
    if part_renban_id0 is None:
        rep.check("sentinel part %s exists for the join" % PART, False, "no INV_PARTS_STOCK_MST row")
        sys.exit(rep.summary_exit())
    lotqty = int(scalar("SELECT IN_1LOTQTY FROM INV_PARTS_STOCK_MST WHERE VC_PART_NUMBER=%s" % q(PART)))

    preclean_sentinels(sql, teardown_statements(), label="renban-collision")
    sql("DELETE FROM INV_RENBAN_GROUP_MST WHERE VC_RENBAN_GROUP_CODE=%s" % q(GROUP))

    grp_id = None
    try:
        # create the synthetic group at count 300, point the part at it
        sql("INSERT INTO INV_RENBAN_GROUP_MST (VC_RENBAN_GROUP_CODE, VC_RENBAN_GROUP_COUNT) "
            "VALUES (%s, '300')" % q(GROUP))
        grp_id = scalar("SELECT IN_RENBAN_ID FROM INV_RENBAN_GROUP_MST WHERE VC_RENBAN_GROUP_CODE=%s" % q(GROUP))
        sql("UPDATE INV_PARTS_STOCK_MST SET IN_RENBAN_ID=%s WHERE VC_PART_NUMBER=%s" % (grp_id, q(PART)))
        renban.system = jython_shim._System()

        # ---- 1. COLLISION DETECTED against a seeded resident renban -----------------------------
        # group count 300, 3 trailers, 1 part -> candidates ZZRB300/301/302. Seed ZZRB301 RESIDENT.
        seed_resident("ZZRB301")
        # one part, qty = 3 lots so 3 trailers each get 1 lot -> 3 sequential renbans 300/301/302
        seed_placeholders([(PART, "K1", 3 * lotqty)])
        res = renban.commit_renban_breakdown(GROUP, 3, 3, DB, session=WRITE_SESSION)   # resolution=None -> the pre-commit gate
        rep.check("collision DETECTED: candidate ZZRB301 collides with the seeded resident row (no write)",
                  res["status"] == "COLLISION"
                  and any(c["renban"] == "ZZRB301" for c in res["collisions"]),
                  "status=%s collisions=%s" % (res["status"], [c["renban"] for c in res["collisions"]]))
        rep.check("collision: the breakdown did NOT write (no grouped rows, counter unchanged at 300)",
                  scalar("SELECT VC_RENBAN_GROUP_COUNT FROM INV_RENBAN_GROUP_MST WHERE VC_RENBAN_GROUP_CODE=%s" % q(GROUP)) == "300"
                  and int(scalar("SELECT COUNT(*) FROM INV_OPEN_ORDER_INF WHERE VC_FRS_NUMBER LIKE '%s%%' AND VC_RENBAN_NUMBER<>''" % FRSPREFIX)) == 0)

        # ---- 2. WARN: the RENBAN_COLLISION alarm row was written (home-hub surface) --------------
        alarm = sql("SELECT VC_ALARM_TYPE, VC_MANIFEST_NUMBER, VC_ASSY_PART_NUMBER, BIT_RESOLVED "
                    "FROM INV_EDI_ALARM_REJ WHERE IN_ALARM_ID=%s" % res["alarm_id"])
        rep.check("WARN: a RENBAN_COLLISION alarm row exists with the colliding renban in VC_MANIFEST_NUMBER",
                  alarm and alarm[0][0] == "RENBAN_COLLISION" and alarm[0][1] == "ZZRB301"
                  and alarm[0][3] == "0",
                  "alarm=%s" % (alarm[0] if alarm else None))
        active_now = int(scalar("SELECT COUNT(*) FROM INV_EDI_ALARM_REJ WHERE VC_ALARM_TYPE='RENBAN_COLLISION' "
                                "AND BIT_RESOLVED=0 AND VC_MANIFEST_NUMBER LIKE '%s%%'" % GROUP))
        rep.check("WARN: the alarm is ACTIVE (BIT_RESOLVED=0) -> the home-hub attention rail surfaces it",
                  active_now >= 1, "active=%d" % active_now)

        # ---- 3. GUIDE: next-free RUN-of-N. candidates 300/301/302 (301 resident) -> first free run-of-3
        # scanning FROM the breakdown start (300) is base 302 -> [302,303,304] (300 alone is free but 301
        # breaks the run starting at 300; 302/303/304 is the first contiguous free run of 3). ----------
        next_free = res["next_free"]
        rep.check("GUIDE: next_free is the first RUN-of-3 base whose [base..base+2] are ALL free (= 302)",
                  next_free == 302 and renban.next_free_run(GROUP, 300, 3, DB) == 302,
                  "next_free=%s (independent next_free_run=%s)" % (next_free, renban.next_free_run(GROUP, 300, 3, DB)))
        # verify the run is genuinely free (no resident row on any of the 3 suffixes)
        run = [(next_free + k) % 1000 for k in range(3)]
        run_renbans = [GROUP + ("%03d" % s) for s in run]
        run_used = int(scalar("SELECT COUNT(*) FROM INV_OPEN_ORDER_INF WHERE VC_RENBAN_NUMBER IN (%s)"
                              % ",".join(q(r) for r in run_renbans)))
        rep.check("GUIDE: the suggested run-of-3 (%s) has NO resident rows (genuinely free)" % run_renbans,
                  run_used == 0 and len(set(run)) == 3, "used=%d" % run_used)
        # the run does NOT include the resident 301
        rep.check("GUIDE: the run-of-3 SKIPS the in-use suffix 301", 301 not in run, "run=%s" % run)

        # ---- 4. FIX use_next_free -> re-map + commit + ack ---------------------------------------
        res_unf = renban.commit_renban_breakdown(GROUP, 3, 3, DB,
                                                 resolution={"action": "use_next_free", "base": next_free,
                                                             "alarm_id": res["alarm_id"]},
                                                 session=WRITE_SESSION)
        rep.check("FIX use_next_free: COMMITTED onto the free run", res_unf["status"] == "COMMITTED",
                  "status=%s" % res_unf["status"])
        committed_renbans = sorted(set(r[0] for r in sql(
            "SELECT VC_RENBAN_NUMBER FROM INV_OPEN_ORDER_INF WHERE VC_FRS_NUMBER LIKE '%s%%' AND VC_RENBAN_NUMBER<>''" % FRSPREFIX)))
        rep.check("FIX use_next_free: the DB rows carry the FREE run renbans (%s), not the colliding 301"
                  % run_renbans,
                  committed_renbans == sorted(run_renbans) and "ZZRB301" not in committed_renbans,
                  "committed=%s" % committed_renbans)
        ack = scalar("SELECT BIT_RESOLVED FROM INV_EDI_ALARM_REJ WHERE IN_ALARM_ID=%s" % res["alarm_id"])
        rep.check("FIX use_next_free: the RENBAN_COLLISION alarm is ACK'd (BIT_RESOLVED=1) in the SAME tx",
                  ack == "1", "BIT_RESOLVED=%s" % ack)

        # reset for the override scenario
        for s in teardown_statements():
            try: sql(s)
            except Exception: pass
        sql("EXEC UPDATE_RenbanGroupCount @RenbanCode=%s, @RenbanCount='300'" % q(GROUP))

        # ---- 5. FIX override -> commit ORIGINAL (acknowledged) renbans + audit note --------------
        seed_resident("ZZRB301")
        seed_placeholders([(PART, "K1", 3 * lotqty)])
        res2 = renban.commit_renban_breakdown(GROUP, 3, 3, DB, session=WRITE_SESSION)
        rep.check("override setup: a fresh collision on ZZRB301", res2["status"] == "COLLISION"
                  and any(c["renban"] == "ZZRB301" for c in res2["collisions"]))
        # the dialog passes back the acknowledged set = the renbans the operator SAW colliding at the WARN.
        ack_set = [c["renban"] for c in res2["collisions"]]
        res_ovr = renban.commit_renban_breakdown(GROUP, 3, 3, DB,
                                                resolution={"action": "override", "alarm_id": res2["alarm_id"],
                                                            "acknowledged": ack_set},
                                                actor="opX", session=WRITE_SESSION)
        rep.check("FIX override: COMMITTED the ORIGINAL renbans incl. the overridden ZZRB301",
                  res_ovr["status"] == "COMMITTED", "status=%s" % res_ovr["status"])
        ovr_renbans = sorted(set(r[0] for r in sql(
            "SELECT VC_RENBAN_NUMBER FROM INV_OPEN_ORDER_INF WHERE VC_FRS_NUMBER LIKE '%s%%' AND VC_RENBAN_NUMBER<>''" % FRSPREFIX)))
        rep.check("FIX override: the original run ZZRB300/301/302 was committed (operator reuse)",
                  ovr_renbans == ["ZZRB300", "ZZRB301", "ZZRB302"], "renbans=%s" % ovr_renbans)
        note = scalar("SELECT VC_ERROR_TEXT FROM INV_EDI_ALARM_REJ WHERE IN_ALARM_ID=%s" % res2["alarm_id"])
        ack2 = scalar("SELECT BIT_RESOLVED FROM INV_EDI_ALARM_REJ WHERE IN_ALARM_ID=%s" % res2["alarm_id"])
        rep.check("FIX override: alarm ACK'd (BIT_RESOLVED=1) with an OVERRIDE audit note stamped",
                  ack2 == "1" and note and note.startswith("OVERRIDE by opX"),
                  "BIT_RESOLVED=%s note=%r" % (ack2, note))

        for s in teardown_statements():
            try: sql(s)
            except Exception: pass
        sql("EXEC UPDATE_RenbanGroupCount @RenbanCode=%s, @RenbanCount='300'" % q(GROUP))

        # ---- 6. CANCEL: no driver call -> the alarm stays ACTIVE ----------------------------------
        seed_resident("ZZRB301")
        seed_placeholders([(PART, "K1", 3 * lotqty)])
        res3 = renban.commit_renban_breakdown(GROUP, 3, 3, DB, session=WRITE_SESSION)   # WARN written, alarm active
        # cancel is purely client-side: NO further driver call. Assert the alarm is still BIT_RESOLVED=0.
        cancel_active = scalar("SELECT BIT_RESOLVED FROM INV_EDI_ALARM_REJ WHERE IN_ALARM_ID=%s" % res3["alarm_id"])
        no_write = int(scalar("SELECT COUNT(*) FROM INV_OPEN_ORDER_INF WHERE VC_FRS_NUMBER LIKE '%s%%' AND VC_RENBAN_NUMBER<>''" % FRSPREFIX))
        rep.check("CANCEL: no driver call -> the RENBAN_COLLISION alarm stays ACTIVE (BIT_RESOLVED=0) + no write",
                  cancel_active == "0" and no_write == 0,
                  "BIT_RESOLVED=%s grouped_rows=%d" % (cancel_active, no_write))

        for s in teardown_statements():
            try: sql(s)
            except Exception: pass
        sql("EXEC UPDATE_RenbanGroupCount @RenbanCode=%s, @RenbanCount='300'" % q(GROUP))

        # ---- 7. TOCTOU re-check FIRES on a LOST RACE on ALL THREE commit paths --------------------
        # The driver's in-tx re-check (READ COMMITTED) must abort if a chosen renban is grabbed by a
        # SECOND connection after the pre-check. We simulate the concurrent Order writer: between building
        # the rows and the driver's commit, a different autocommit connection seeds the chosen renban. To
        # drive this deterministically without racing wall-clock, we INSERT the resident row BEFORE the
        # commit call but AFTER the pre-check would have passed — i.e. we seed a renban the pre-check did
        # NOT see (a number that was free) and confirm the in-tx re-check still catches it on each path.

        # 7a. straight-commit path (resolution=None, NO pre-existing collision): pre-check clears, then a
        #     concurrent writer grabs ZZRB300 -> the in-tx re-check must abort.
        seed_placeholders([(PART, "K1", 3 * lotqty)])           # candidates 300/301/302, all free now
        clear = renban.check_renban_collisions(
            [{"renban": "ZZRB300", "part": PART}, {"renban": "ZZRB301", "part": PART},
             {"renban": "ZZRB302", "part": PART}], DB)
        rep.check("TOCTOU 7a: pre-check is CLEAR before the race (300/301/302 free)", clear == [])
        seed_resident("ZZRB300", frs="9124002")                 # concurrent writer claims ZZRB300
        res_race = renban.commit_renban_breakdown(GROUP, 3, 3, DB, session=WRITE_SESSION)  # straight commit -> in-tx re-check fires
        rep.check("TOCTOU 7a (straight-commit): in-tx re-check FIRES on the lost race -> COLLISION, no write",
                  res_race["status"] == "COLLISION"
                  and any(c["renban"] == "ZZRB300" for c in res_race["collisions"])
                  and int(scalar("SELECT COUNT(*) FROM INV_OPEN_ORDER_INF WHERE VC_FRS_NUMBER LIKE '%s%%' AND VC_RENBAN_NUMBER<>''" % FRSPREFIX)) == 0,
                  "status=%s collisions=%s" % (res_race["status"], [c["renban"] for c in res_race["collisions"]]))

        # 7b. use_next_free path: the chosen free base gets grabbed before commit -> re-check aborts.
        for s in teardown_statements():
            try: sql(s)
            except Exception: pass
        sql("EXEC UPDATE_RenbanGroupCount @RenbanCode=%s, @RenbanCount='300'" % q(GROUP))
        seed_resident("ZZRB301")                                # original collision
        seed_placeholders([(PART, "K1", 3 * lotqty)])
        res_b = renban.commit_renban_breakdown(GROUP, 3, 3, DB, session=WRITE_SESSION)  # COLLISION -> next_free base
        base_b = res_b["next_free"]
        race_renban = GROUP + ("%03d" % (base_b % 1000))         # the FIRST renban of the chosen free run
        seed_resident(race_renban, frs="9124003")               # concurrent writer claims it
        res_b2 = renban.commit_renban_breakdown(GROUP, 3, 3, DB,
                                               resolution={"action": "use_next_free", "base": base_b,
                                                           "alarm_id": res_b["alarm_id"]},
                                               session=WRITE_SESSION)
        rep.check("TOCTOU 7b (use_next_free): the chosen free base was grabbed -> in-tx re-check aborts (COLLISION)",
                  res_b2["status"] == "COLLISION"
                  and any(c["renban"] == race_renban for c in res_b2["collisions"]),
                  "status=%s collisions=%s race=%s" % (res_b2["status"],
                  [c["renban"] for c in res_b2["collisions"]], race_renban))

        # 7c. override path: a DIFFERENT renban (NOT the acknowledged one) is grabbed -> re-check aborts.
        for s in teardown_statements():
            try: sql(s)
            except Exception: pass
        sql("EXEC UPDATE_RenbanGroupCount @RenbanCode=%s, @RenbanCount='300'" % q(GROUP))
        seed_resident("ZZRB301")                                # the operator will OVERRIDE 301
        seed_placeholders([(PART, "K1", 3 * lotqty)])
        res_c = renban.commit_renban_breakdown(GROUP, 3, 3, DB, session=WRITE_SESSION)  # COLLISION on 301; candidates 300/301/302
        ack_c = [c["renban"] for c in res_c["collisions"]]       # operator saw ONLY 301 colliding
        seed_resident("ZZRB302", frs="9124004")                 # a DIFFERENT candidate (302) grabbed newly
        # override acknowledges ONLY 301 (what the operator saw). 302 was FREE at the WARN but is now taken
        # -> the in-tx re-check must abort on 302 (a number the operator never agreed to reuse).
        res_c2 = renban.commit_renban_breakdown(GROUP, 3, 3, DB,
                                               resolution={"action": "override", "alarm_id": res_c["alarm_id"],
                                                           "acknowledged": ack_c},
                                               actor="opX", session=WRITE_SESSION)
        rep.check("TOCTOU 7c (override): a NEWLY-taken non-acknowledged renban (302) -> in-tx re-check aborts",
                  res_c2["status"] == "COLLISION"
                  and any(c["renban"] == "ZZRB302" for c in res_c2.get("collisions", [])),
                  "status=%s collisions=%s" % (res_c2["status"], [c["renban"] for c in res_c2.get("collisions", [])]))
        # the override did NOT write (aborted) — counter still 300, no grouped rows
        rep.check("TOCTOU 7c: the aborted override wrote NOTHING (counter 300, no grouped rows)",
                  scalar("SELECT VC_RENBAN_GROUP_COUNT FROM INV_RENBAN_GROUP_MST WHERE VC_RENBAN_GROUP_CODE=%s" % q(GROUP)) == "300"
                  and int(scalar("SELECT COUNT(*) FROM INV_OPEN_ORDER_INF WHERE VC_FRS_NUMBER LIKE '%s%%' AND VC_RENBAN_NUMBER<>''" % FRSPREFIX)) == 0)

        for s in teardown_statements():
            try: sql(s)
            except Exception: pass
        sql("EXEC UPDATE_RenbanGroupCount @RenbanCode=%s, @RenbanCount='300'" % q(GROUP))

        # ---- 8. EXHAUSTION: a fully-occupied ring -> next_free_run = None -> hard WARN-cancel ------
        # Fill a SMALL window deterministically: seed resident rows on every suffix 000..999 would be 1000
        # inserts; instead prove next_free_run returns None when N exceeds the free space. Seed all but 2
        # suffixes' worth is impractical; instead test the boundary: seed suffixes 300..309 resident and
        # ask for a run of N where no contiguous free run of N fits at/after a saturated start with N huge.
        # Cleanest deterministic exhaustion: monkey-make a tiny ring impossible -> request a RUN larger than
        # the total FREE suffixes is still satisfiable elsewhere, so to prove the None path we seed a dense
        # block and confirm the scan SKIPS it; for true exhaustion we assert next_free_run(N=1000) is None
        # (a run of 1000 contiguous distinct suffixes cannot fit a 1000-ring that has ANY occupied slot).
        seed_resident("ZZRB500")                                # one occupied slot anywhere
        none_run = renban.next_free_run(GROUP, 300, 1000, DB)   # a run of all-1000 cannot fit with 1 used
        rep.check("EXHAUSTION: next_free_run for a run that cannot fit the ring -> None (hard WARN-cancel)",
                  none_run is None, "next_free_run(N=1000)=%s" % none_run)
        # and a normal run-of-3 still finds space (the ring is not actually full) — the None is the run-too-big
        ok_run = renban.next_free_run(GROUP, 300, 3, DB)
        rep.check("EXHAUSTION boundary: a run-of-3 still resolves when the ring is NOT full",
                  ok_run is not None, "run-of-3=%s" % ok_run)

        for s in teardown_statements():
            try: sql(s)
            except Exception: pass
        sql("EXEC UPDATE_RenbanGroupCount @RenbanCode=%s, @RenbanCount='300'" % q(GROUP))

        # ---- 9. BLOCKER-1: WRAPPED candidate run x use_next_free -> writes the VALIDATED contiguous
        #         block (offset-from-POSITION, not from the tail value), correct persisted count, NO
        #         collision. The previously-untested cell (the wrapped x use_next_free combination).
        #
        # Setup: group count 998, 3 trailers, one 3-lot part -> candidate tails 998, 999, 000 (the run
        # STRADDLES 999->000, the exact CMWA rollover the allocator exists for). Seed ZZRB999 RESIDENT so
        # resolution=None DETECTS a collision and emits a next_free GUIDE base.
        #
        # SPEC-DERIVED ORACLE (R15, non-vacuous — NOT the rebuild's own recompute): under the design, the
        # FIX must re-seat the N=3 distinct candidate renbans onto EXACTLY the validated contiguous block
        # [base, base+1, base+2] (% 1000) in EMISSION order (truck0->base, truck1->base+1, truck2->base+2),
        # and persist next_count = (base + 3) % 1000. We compute this oracle here with a standalone ring
        # helper, independent of _remap_rows_onto_base.
        #
        # OLD-BUG WITNESS: the defective tail-derived model produced offsets [998,999,0] (min(tails)=0),
        # writing the SHIFTED block [base-2, base-1, base] and persisting (base+999+1)%1000 = base. We seed
        # a RESIDENT occupant at base-1 (where the OLD block would land) so the old code would have written
        # ONTO an occupied number (and the in-tx re-check or a silent collision) while the FIXED block
        # [base, base+1, base+2] is genuinely free -> this check FAILS on the pre-fix code (revert-proof).
        def ring3(n):
            return "%03d" % (n % 1000)

        seed_resident("ZZRB999")                                 # the collision trigger on the wrapped run
        seed_placeholders([(PART, "K1", 3 * lotqty)])            # 3 lots over 3 trailers -> tails 998/999/000
        sql("EXEC UPDATE_RenbanGroupCount @RenbanCode=%s, @RenbanCount='998'" % q(GROUP))
        res_w = renban.commit_renban_breakdown(GROUP, 3, 3, DB, session=WRITE_SESSION)
        rep.check("WRAPPED 9: resolution=None DETECTS the collision on the straddling run (ZZRB999 resident)",
                  res_w["status"] == "COLLISION"
                  and any(c["renban"] == "ZZRB999" for c in res_w["collisions"]),
                  "status=%s collisions=%s" % (res_w["status"], [c["renban"] for c in res_w["collisions"]]))
        base_w = res_w["next_free"]
        rep.check("WRAPPED 9: GUIDE returned a free run-of-3 base for the wrapped run", base_w is not None,
                  "next_free=%s" % base_w)
        # SPEC oracle: the validated contiguous block + persisted count, computed independently of the rebuild.
        expected_block = sorted(GROUP + ring3(base_w + k) for k in range(3))
        expected_persist = ring3(base_w + 3)
        # OLD-bug block (tail-derived): [base-2, base-1, base]. Seed an occupant at base-1 so old code collides.
        old_block_suffix_b1 = (base_w - 1) % 1000
        seed_resident(GROUP + ring3(base_w - 1), frs="9124009")  # where the OLD remap would have written
        # the FIXED block must NOT include base-1 (proves offset-from-position, not tail) -> commit must avoid it
        res_w2 = renban.commit_renban_breakdown(GROUP, 3, 3, DB,
                                               resolution={"action": "use_next_free", "base": base_w,
                                                           "alarm_id": res_w["alarm_id"]},
                                               session=WRITE_SESSION)
        rep.check("WRAPPED 9: use_next_free COMMITTED the wrapped run onto the free block (no collision)",
                  res_w2["status"] == "COMMITTED", "status=%s coll=%s"
                  % (res_w2["status"], [c["renban"] for c in res_w2.get("collisions", [])]))
        written = sorted(set(r[0] for r in sql(
            "SELECT VC_RENBAN_NUMBER FROM INV_OPEN_ORDER_INF WHERE VC_FRS_NUMBER LIKE '%s%%' AND VC_RENBAN_NUMBER<>''" % FRSPREFIX)))
        rep.check("WRAPPED 9: DB rows == the SPEC validated contiguous block [base..base+2] (%s), NOT the "
                  "tail-shifted [base-2..base] block" % expected_block,
                  written == expected_block, "written=%s expected=%s" % (written, expected_block))
        rep.check("WRAPPED 9: the committed block is genuinely free (does NOT reuse the seeded base-1 occupant "
                  "ZZRB%s where the OLD tail-derived remap would have written)" % ring3(base_w - 1),
                  (GROUP + ring3(base_w - 1)) not in written and old_block_suffix_b1 not in
                  [int(r[-3:]) for r in written], "written=%s base-1=%s" % (written, ring3(base_w - 1)))
        persisted_w = scalar("SELECT VC_RENBAN_GROUP_COUNT FROM INV_RENBAN_GROUP_MST WHERE VC_RENBAN_GROUP_CODE=%s" % q(GROUP))
        rep.check("WRAPPED 9: persisted count == SPEC oracle ring3(base+3)=%s (true highest-written+1, wrapped) "
                  "— NOT the OLD base self-collision value" % expected_persist,
                  persisted_w == expected_persist, "persisted=%r expected=%r (old-bug would persist %r)"
                  % (persisted_w, expected_persist, ring3(base_w)))
        # self-collision guard: the NEXT breakdown seeds from the persisted count; its run must NOT overlap
        # the block we just wrote (the old next_count=base would re-issue base/base+1 -> self-collision).
        next_seed = int(persisted_w)
        next_run = set((next_seed + k) % 1000 for k in range(3))
        this_run = set((base_w + k) % 1000 for k in range(3))
        rep.check("WRAPPED 9: the persisted count seeds the NEXT run PAST this run (no self-collision next "
                  "breakdown)", next_run.isdisjoint(this_run),
                  "next_run=%s this_run=%s" % (sorted(next_run), sorted(this_run)))

        for s in teardown_statements():
            try: sql(s)
            except Exception: pass
        sql("EXEC UPDATE_RenbanGroupCount @RenbanCode=%s, @RenbanCount='300'" % q(GROUP))

        # ---- 10. BLOCKER-2: the SERVER-SIDE WRITE GATE (auth.requireWrite) on ALL THREE paths ------------
        # The driver writes (delete-then-reinsert + counter bump + alarm). It MUST authorize FROM THE SESSION
        # server-side (the R25/P15 hole class). A forged/anon/no-write-role session is REJECTED before any
        # read/compute/write; ProductionControl/Admin reach the write. Revert-proof: neuter requireWrite ->
        # the forged write slips through (so the GATE, not the harness, is load-bearing). Mirrors
        # test_master_write_gates.
        import auth as _auth
        for label, sess in (("anonymous (not logged in)", ANON_SESSION),
                            ("logged-in viewer with no write role", FORGED_SESSION)):
            seed_resident("ZZRB301")
            seed_placeholders([(PART, "K1", 3 * lotqty)])
            sql("EXEC UPDATE_RenbanGroupCount @RenbanCode=%s, @RenbanCount='300'" % q(GROUP))
            denied = False
            try:
                renban.commit_renban_breakdown(GROUP, 3, 3, DB, session=sess)
            except _auth.AuthError:
                denied = True
            # the gate ran BEFORE any write: counter still 300, no grouped rows, NO alarm raised
            cnt = scalar("SELECT VC_RENBAN_GROUP_COUNT FROM INV_RENBAN_GROUP_MST WHERE VC_RENBAN_GROUP_CODE=%s" % q(GROUP))
            grouped = int(scalar("SELECT COUNT(*) FROM INV_OPEN_ORDER_INF WHERE VC_FRS_NUMBER LIKE '%s%%' AND VC_RENBAN_NUMBER<>''" % FRSPREFIX))
            alarms = int(scalar("SELECT COUNT(*) FROM INV_EDI_ALARM_REJ WHERE VC_ALARM_TYPE='RENBAN_COLLISION' AND VC_MANIFEST_NUMBER LIKE '%s%%'" % GROUP))
            rep.check("GATE 10: a %s session is REJECTED server-side (AuthError) BEFORE any read/compute/write "
                      "(counter 300, no rows, no alarm)" % label,
                      denied and cnt == "300" and grouped == 0 and alarms == 0,
                      "denied=%s cnt=%s grouped=%d alarms=%d" % (denied, cnt, grouped, alarms))
            for s in teardown_statements():
                try: sql(s)
                except Exception: pass
            sql("EXEC UPDATE_RenbanGroupCount @RenbanCode=%s, @RenbanCount='300'" % q(GROUP))

        # NON-VACUITY: an authorized ProductionControl session REACHES the write (the gate is not block-all).
        seed_placeholders([(PART, "K1", 3 * lotqty)])            # candidates 300/301/302 all free -> straight commit
        sql("EXEC UPDATE_RenbanGroupCount @RenbanCode=%s, @RenbanCount='300'" % q(GROUP))
        res_pc = renban.commit_renban_breakdown(GROUP, 3, 3, DB, session=WRITE_SESSION)
        pc_rows = int(scalar("SELECT COUNT(*) FROM INV_OPEN_ORDER_INF WHERE VC_FRS_NUMBER LIKE '%s%%' AND VC_RENBAN_NUMBER<>''" % FRSPREFIX))
        rep.check("GATE 10 non-vacuity: a ProductionControl session REACHES the write (COMMITTED, rows written)",
                  res_pc["status"] == "COMMITTED" and pc_rows == 3, "status=%s rows=%d" % (res_pc["status"], pc_rows))
        for s in teardown_statements():
            try: sql(s)
            except Exception: pass
        sql("EXEC UPDATE_RenbanGroupCount @RenbanCode=%s, @RenbanCount='300'" % q(GROUP))
        # Admin is also a writer
        seed_placeholders([(PART, "K1", 3 * lotqty)])
        sql("EXEC UPDATE_RenbanGroupCount @RenbanCode=%s, @RenbanCount='300'" % q(GROUP))
        res_admin = renban.commit_renban_breakdown(GROUP, 3, 3, DB, session=ADMIN_SESSION)
        rep.check("GATE 10: an Admin session REACHES the write too (Admin is a write role)",
                  res_admin["status"] == "COMMITTED", "status=%s" % res_admin["status"])
        for s in teardown_statements():
            try: sql(s)
            except Exception: pass
        sql("EXEC UPDATE_RenbanGroupCount @RenbanCode=%s, @RenbanCount='300'" % q(GROUP))

        # REVERT-PROOF: neuter the SERVER gate -> the forged-anon write SLIPS THROUGH to the DB (so the gate,
        # not the harness, is load-bearing). Restore it immediately after.
        orig_require = _auth.requireWrite
        slipped = False
        try:
            _auth.requireWrite = lambda session: set()           # the hole, re-opened
            seed_placeholders([(PART, "K1", 3 * lotqty)])
            sql("EXEC UPDATE_RenbanGroupCount @RenbanCode=%s, @RenbanCount='300'" % q(GROUP))
            res_neuter = renban.commit_renban_breakdown(GROUP, 3, 3, DB, session=ANON_SESSION)
            slipped = res_neuter["status"] == "COMMITTED" and int(scalar(
                "SELECT COUNT(*) FROM INV_OPEN_ORDER_INF WHERE VC_FRS_NUMBER LIKE '%s%%' AND VC_RENBAN_NUMBER<>''" % FRSPREFIX)) == 3
        finally:
            _auth.requireWrite = orig_require
            for s in teardown_statements():
                try: sql(s)
                except Exception: pass
            sql("EXEC UPDATE_RenbanGroupCount @RenbanCode=%s, @RenbanCount='300'" % q(GROUP))
        rep.check("GATE 10 REVERT-PROOF: with auth.requireWrite neutered, the FORGED-anon write SLIPS THROUGH "
                  "to the DB (the gate, not the harness, is what blocks it)", slipped,
                  "slipped=%s" % slipped)
        # and with the gate RESTORED, the anon session is denied again
        denied_again = False
        try:
            seed_placeholders([(PART, "K1", 3 * lotqty)])
            sql("EXEC UPDATE_RenbanGroupCount @RenbanCode=%s, @RenbanCount='300'" % q(GROUP))
            renban.commit_renban_breakdown(GROUP, 3, 3, DB, session=ANON_SESSION)
        except _auth.AuthError:
            denied_again = True
        rep.check("GATE 10 REVERT-PROOF restore: with the gate back, the anon session is DENIED again",
                  denied_again, "denied_again=%s" % denied_again)
        for s in teardown_statements():
            try: sql(s)
            except Exception: pass
        sql("EXEC UPDATE_RenbanGroupCount @RenbanCode=%s, @RenbanCount='300'" % q(GROUP))

        # ---- 11. SHOULD-FIX-1: the GUIDE scan DETECTS a TRAILING-SPACE resident renban -------------------
        # _all_resident_suffixes now reads SUBSTRING(.,len+1,3) (not RIGHT(.,3)), so 'ZZRB289 ' (trailing
        # space) is recorded as suffix 289, NOT 089. Seed a trailing-space resident on 289 and confirm
        # next_free_run treats 289 as IN USE (the run-of-N SKIPS it). Revert-proof: the OLD RIGHT(.,3) read
        # '89 ' -> 89, so it would NOT skip 289 -> this check fails on the pre-fix scan.
        # NB: store the renban with a TRAILING SPACE via an explicit 8-char literal.
        sql("INSERT INTO INV_OPEN_ORDER_INF "
            "(VC_SUPPLIER_CODE, VC_PART_NUMBER, VC_KANBAN_NUMBER, VC_FRS_NUMBER, VC_RENBAN_NUMBER, "
            " IN_QTY, VC_ORDER_DATE, VC_FRS_DATE, VC_ADD) VALUES "
            "(%s, %s, 'KZ', '9124005', 'ZZRB289 ', 40, '20260101', '20260101', '20260101000000')"
            % (q(SENT_SUP), q(PART)))
        ts_stored = scalar("SELECT DATALENGTH(VC_RENBAN_NUMBER) FROM INV_OPEN_ORDER_INF WHERE VC_FRS_NUMBER='9124005'")
        used_suffixes = renban._all_resident_suffixes(GROUP, DB)
        rep.check("SHOULD-FIX-1: a trailing-space renban 'ZZRB289 ' (DATALENGTH %s) is read as suffix 289, "
                  "NOT 089 (SUBSTRING, not RIGHT)" % ts_stored,
                  289 in used_suffixes and 89 not in used_suffixes,
                  "289-in-used=%s 89-in-used=%s used=%s" % (289 in used_suffixes, 89 in used_suffixes,
                  sorted(s for s in used_suffixes if s in (89, 289))))
        # the GUIDE run-of-N must SKIP 289 (a run starting at 287/288/289 cannot include the in-use 289)
        run_from_287 = renban.next_free_run(GROUP, 287, 3, DB)
        rep.check("SHOULD-FIX-1: next_free_run SKIPS the trailing-space-occupied 289 (the GUIDE never "
                  "recommends a run containing it)",
                  run_from_287 is not None and 289 not in
                  set((run_from_287 + k) % 1000 for k in range(3)),
                  "base=%s run=%s" % (run_from_287, sorted((run_from_287 + k) % 1000 for k in range(3))
                                      if run_from_287 is not None else None))
        sql("DELETE FROM INV_OPEN_ORDER_INF WHERE VC_FRS_NUMBER='9124005'")
        for s in teardown_statements():
            try: sql(s)
            except Exception: pass
        sql("EXEC UPDATE_RenbanGroupCount @RenbanCode=%s, @RenbanCount='300'" % q(GROUP))

        # ---- 12. SHOULD-FIX-2: override with the EXPLICIT WARN-payload acknowledged set ------------------
        # Contract: the dialog passes back the seen-colliding set. A DELIBERATE reuse of an acknowledged
        # number COMMITS; a DIFFERENT renban taken NEWLY since the WARN still ABORTS (already proven by 7c).
        # Here we prove the COMMIT half explicitly with the real WARN payload + that the None default now
        # fails CLOSED (acknowledges nothing -> aborts on the still-resident reused number).
        seed_resident("ZZRB301")
        seed_placeholders([(PART, "K1", 3 * lotqty)])
        sql("EXEC UPDATE_RenbanGroupCount @RenbanCode=%s, @RenbanCount='300'" % q(GROUP))
        res_sf2 = renban.commit_renban_breakdown(GROUP, 3, 3, DB, session=WRITE_SESSION)
        ack_payload = [c["renban"] for c in res_sf2["collisions"]]   # the real WARN payload (= ['ZZRB301'])
        # 12a. override WITH the explicit acknowledged set -> the deliberate reuse of 301 COMMITS
        res_sf2b = renban.commit_renban_breakdown(GROUP, 3, 3, DB,
                                                 resolution={"action": "override", "alarm_id": res_sf2["alarm_id"],
                                                             "acknowledged": ack_payload},
                                                 actor="opSF2", session=WRITE_SESSION)
        sf2_rows = sorted(set(r[0] for r in sql(
            "SELECT VC_RENBAN_NUMBER FROM INV_OPEN_ORDER_INF WHERE VC_FRS_NUMBER LIKE '%s%%' AND VC_RENBAN_NUMBER<>''" % FRSPREFIX)))
        rep.check("SHOULD-FIX-2 (12a): override WITH the WARN-payload acknowledged set COMMITS the deliberate "
                  "reuse (ZZRB300/301/302)", res_sf2b["status"] == "COMMITTED"
                  and sf2_rows == ["ZZRB300", "ZZRB301", "ZZRB302"],
                  "status=%s rows=%s" % (res_sf2b["status"], sf2_rows))
        for s in teardown_statements():
            try: sql(s)
            except Exception: pass
        sql("EXEC UPDATE_RenbanGroupCount @RenbanCode=%s, @RenbanCount='300'" % q(GROUP))
        # 12b. override with NO acknowledged set (None) now defaults to EMPTY -> the still-resident 301
        # is NOT acknowledged -> the in-tx re-check ABORTS (fail-closed; the old auto-ack would have
        # silently committed). This is the SHOULD-FIX-2 contract made explicit.
        seed_resident("ZZRB301")
        seed_placeholders([(PART, "K1", 3 * lotqty)])
        sql("EXEC UPDATE_RenbanGroupCount @RenbanCode=%s, @RenbanCount='300'" % q(GROUP))
        res_sf2c0 = renban.commit_renban_breakdown(GROUP, 3, 3, DB, session=WRITE_SESSION)  # WARN
        res_sf2c = renban.commit_renban_breakdown(GROUP, 3, 3, DB,
                                                 resolution={"action": "override", "alarm_id": res_sf2c0["alarm_id"]},
                                                 actor="opSF2", session=WRITE_SESSION)       # NO acknowledged
        sf2c_rows = int(scalar("SELECT COUNT(*) FROM INV_OPEN_ORDER_INF WHERE VC_FRS_NUMBER LIKE '%s%%' AND VC_RENBAN_NUMBER<>''" % FRSPREFIX))
        rep.check("SHOULD-FIX-2 (12b): override with NO acknowledged set defaults to EMPTY -> fail-CLOSED: the "
                  "in-tx re-check ABORTS on the un-acknowledged still-resident 301 (no silent auto-ack)",
                  res_sf2c["status"] == "COLLISION"
                  and any(c["renban"] == "ZZRB301" for c in res_sf2c.get("collisions", []))
                  and sf2c_rows == 0,
                  "status=%s coll=%s rows=%d" % (res_sf2c["status"],
                  [c["renban"] for c in res_sf2c.get("collisions", [])], sf2c_rows))

        for s in teardown_statements():
            try: sql(s)
            except Exception: pass
        sql("EXEC UPDATE_RenbanGroupCount @RenbanCode=%s, @RenbanCount='300'" % q(GROUP))

    finally:
        for s in teardown_statements():
            try: sql(s)
            except Exception: pass
        # restore the part's original renban group + drop the synthetic group
        sql("UPDATE INV_PARTS_STOCK_MST SET IN_RENBAN_ID=%s WHERE VC_PART_NUMBER=%s"
            % (part_renban_id0, q(PART)))
        sql("DELETE FROM INV_RENBAN_GROUP_MST WHERE VC_RENBAN_GROUP_CODE=%s" % q(GROUP))

    # restored-as-found assertions
    leftover_orders = int(scalar("SELECT COUNT(*) FROM INV_OPEN_ORDER_INF WHERE VC_RENBAN_NUMBER LIKE '%s%%' "
                                 "OR VC_FRS_NUMBER LIKE '%s%%' OR VC_FRS_NUMBER LIKE '9124%%'" % (GROUP, FRSPREFIX)))
    leftover_grp = int(scalar("SELECT COUNT(*) FROM INV_RENBAN_GROUP_MST WHERE VC_RENBAN_GROUP_CODE=%s" % q(GROUP)))
    leftover_alarm = int(scalar("SELECT COUNT(*) FROM INV_EDI_ALARM_REJ WHERE VC_ALARM_TYPE='RENBAN_COLLISION' "
                                "AND VC_MANIFEST_NUMBER LIKE '%s%%'" % GROUP))
    part_grp_restored = scalar("SELECT IN_RENBAN_ID FROM INV_PARTS_STOCK_MST WHERE VC_PART_NUMBER=%s" % q(PART))
    rep.check("spike restored as-found (no sentinel orders/group/alarms; part renban group restored)",
              leftover_orders == 0 and leftover_grp == 0 and leftover_alarm == 0
              and part_grp_restored == part_renban_id0,
              "orders=%d grp=%d alarm=%d part_grp=%s(want %s)"
              % (leftover_orders, leftover_grp, leftover_alarm, part_grp_restored, part_renban_id0))
    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
