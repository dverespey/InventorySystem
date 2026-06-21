#!/usr/bin/env python3
"""test_order_file_e2e.py — END-TO-END exercise of the REAL order-FILE generator
(docs/analysis/order/project-library/order_file/code.py) driven headlessly through the system.db shim
(jython_shim) against the spike DB — the M2 unit-2 proof.

WHAT THIS PROVES (running the ACTUAL generate_order_files driver, NOT a proc-EXEC re-implementation — R8)
-------------------------------------------------------------------------------------------------------
Synthesizes not-yet-ordered open-order rows against a REAL part (so the feed's parts/supplier/renban
joins resolve), then drives generate_order_files end-to-end inside the ONE stamp transaction the driver
itself opens:

  (1) .ord BYTES — the published file's exact bytes equal format_ord_line for each row (supplier raw +
      FRS raw + %8s renban + part raw + %05d qty + yyyymmdd ship), CRLF-terminated; one line per row.
  (2) 3-DESTINATION FAN-OUT — LocalFTP=False writes ONE copy (supplier dir only); LocalFTP=True writes
      supplier + Archive (+ logistics ONLY when the resolved logistics dir is not 'NONE'). The 'NONE'
      sentinel skips the logistics copy (Q8). Proven by toggling LocalFTP + the logistics resolution.
  (3) COLUMN ALIASING (legacy H10) — the emitted qty is the OPEN-ORDER IN_QTY, NOT the parts-stock
      on-hand IN_QTY (they differ on the fixture by construction); proves the aliased feed picked i.IN_QTY.
  (4) EMIT-THEN-STAMP / re-emit guard (legacy H11) — after a run the open-order rows carry VC_ORDER_DATE;
      a 2nd run re-reads the feed, finds them stamped-out, and emits NOTHING (idempotent). The stamp
      covers all same part+FRS rows (the proc has no renban filter) from the SNAPSHOT, never re-querying.
  (5) ATOMICITY (legacy H1 FIX) — inject a DB failure AFTER the .tmp files are written (during the stamp
      tx): the driver rolls back, deletes every orphan .tmp, and publishes NO final .ord. The open-order
      rows stay UN-stamped (a re-run would re-emit). The thing the legacy could not do (files mid-tx).

HONEST VERIFICATION
-------------------
The .ord is BYTE-FAITHFUL to OrderFormCreateF.pas:556-585 (proven exact-bytes here + in the pure test).
PENDING a golden: (a) whether the receiving sub-supplier parser expects always-full-width supplier/FRS/
part columns (the legacy writes them RAW — H3); (b) the M4 multi-site leading field (INV_SITES.
VC_SUPPLIER_CODE) width; (c) the Excel order-form VISUAL layout (templates not on disk). The ship date
uses the REAL AD_GetSpecialDate holiday calendar in VehicleOrder (read-only).

Synthetic-only + restore-as-found: every row is tagged with a far-future FRS prefix ('98')/sentinel
VC_ADD; teardown deletes exactly those + asserts the spike is restored; Inventory_Live / VehicleOrder
are never written. Files go to a tempfile.mkdtemp, removed at the end.

Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 scripts/e2e/test_order_file_e2e.py
"""
import os, sys, subprocess, tempfile, shutil, datetime, glob

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report          # noqa: E402
import jython_shim              # noqa: E402

CONTAINER = os.environ.get("CONTAINER", "mssql-spike")
SA_PASS = os.environ.get("SA_PASS")
INV = "Inventory"

# A REAL part whose parts/supplier joins resolve in the feed. We attach SYNTHETIC open-order rows to it.
PART = "426070205000"           # supplier 11111, kanban TPMSS, 1lotqty 180, IN_SHIP_DAYS 0 (-> ship offset 0)
SUPCODE = "11111"               # the part's master supplier (what the feed groups/writes)
ADD_STAMP = "ZZOFADD0000000A"   # 16-char VC_ADD sentinel on synthetic rows (exact teardown)
FRS_PREFIX = "98"               # far-future FRS prefix -> our rows are uniquely identifiable
STOCK_ONHAND_SENTINEL = None    # captured at runtime (the part's real p.IN_QTY)

CODE_PY = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                           "docs", "analysis", "order", "project-library", "order_file", "code.py"))
FEED_SQL_FILE = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                                 "docs", "analysis", "order", "spike-order-file-feed.sql"))


def sql(query, db=INV):
    if not SA_PASS:
        sys.exit("export SA_PASS first")
    out = subprocess.check_output([
        "docker", "exec", CONTAINER, "/opt/mssql-tools18/bin/sqlcmd", "-C", "-S", "localhost",
        "-U", "sa", "-P", SA_PASS, "-d", db, "-h", "-1", "-W", "-s", "\t",
        "-Q", "SET NOCOUNT ON; " + query], text=True)
    return [l.split("\t") for l in out.splitlines()
            if l.strip() and not l.startswith("(") and not l.startswith("Msg ")]


def scalar(q, db=INV):
    r = sql(q, db)
    return r[0][0] if r else None


def sql_uncommitted(query, db=INV):
    """READ UNCOMMITTED variant for asserting AUTOCOMMITTED rows right after a rolled-back persistent-
    session tx: the shim holds the rollback connection open until close() drains, and a READ COMMITTED
    read fired the same instant can BLOCK on that connection's lingering locks (a harness artifact, not a
    driver behaviour — see the (5) settle comment). The synthetic rows were committed BEFORE the driver
    began, so a dirty read of them is correct here and immune to the close-race."""
    return sql("SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED; " + query, db)


def load_order_file():
    return jython_shim.load_wrapper("order_file_real", CODE_PY)


def seed_orders(renbans, qty):
    """Insert synthetic NOT-ORDERED open-order rows for PART under FRS_PREFIX, distinct renbans, one shared
    part+FRS (so the same-part+FRS stamp fan-out is exercised), VC_ORDER_DATE='' (not ordered)."""
    frs = FRS_PREFIX + "00001"      # 7-char FRS (all renbans share it -> stamp marks all on one call)
    vals = ", ".join(
        "('%s','%s','%s','%s',%d,'','TPMSS','%s')" % (SUPCODE, PART, frs, rb, qty, ADD_STAMP)
        for rb in renbans)
    sql("INSERT INTO INV_OPEN_ORDER_INF "
        "(VC_SUPPLIER_CODE, VC_PART_NUMBER, VC_FRS_NUMBER, VC_RENBAN_NUMBER, IN_QTY, VC_ORDER_DATE, "
        "VC_KANBAN_NUMBER, VC_ADD) VALUES %s" % vals)


def cleanup():
    sql("DELETE FROM INV_OPEN_ORDER_INF WHERE VC_ADD='%s'; "
        "DELETE FROM INV_OPEN_ORDER_INF_HIST WHERE VC_ADD='%s'" % (ADD_STAMP, ADD_STAMP))


def list_ord(d):
    return sorted(os.path.basename(p) for p in glob.glob(os.path.join(d, "**", "*"), recursive=True)
                  if os.path.isfile(p))


def main():
    print("=" * 82)
    print(" ORDER-FILE generator E2E — REAL generate_order_files via the shim (M2 unit-2, R8 §7)")
    print("=" * 82)
    rep = Report()
    if not SA_PASS:
        sys.exit("export SA_PASS first")

    of = load_order_file()
    rep.check("real order_file module loads under the shim (generate_order_files present)",
              hasattr(of, "generate_order_files") and hasattr(of, "format_ord_line"))

    # _FEED_SQL must not drift from the canonical .sql (the aliased-feed contract).
    canon = open(FEED_SQL_FILE).read()
    # extract the SELECT...; body from the .sql (strip the comment block + trailing ;), normalize ws/comments
    def _norm(s):
        import re
        s = re.sub(r"--[^\n]*", "", s)          # strip line comments
        s = re.sub(r"\s+", " ", s).strip().rstrip(";")
        return s
    sel = canon[canon.upper().index("SELECT I."):]
    feedMatch = _norm(of._FEED_SQL) == _norm(sel)
    rep.check("_FEED_SQL matches the canonical spike-order-file-feed.sql (no drift)", feedMatch,
              "" if feedMatch else "drift:\n  code=%r\n  sql =%r" % (_norm(of._FEED_SQL), _norm(sel)))

    cleanup()                                   # start clean
    onhand = int(scalar("SELECT IN_QTY FROM INV_PARTS_STOCK_MST WHERE VC_PART_NUMBER='%s'" % PART))
    ORDER_QTY = 1200                            # DELIBERATELY != the part's on-hand stock qty (aliasing proof)
    rep.check("fixture: order qty %d differs from parts-stock on-hand %d (so aliasing is observable)"
              % (ORDER_QTY, onhand), ORDER_QTY != onhand, "onhand=%d" % onhand)

    runDate = datetime.date(2026, 6, 22)        # a Monday
    orderDateStr = runDate.strftime("%Y%m%d")
    # The part is in renban group 7, so SELECT_PartShipDays returns the GROUP's ship-days (Monday lead 13),
    # NOT the part's own (0) — this itself proves the renban-group ship-day override + the GetShip scan.
    # Compute the expected ship date via the driver's OWN get_ship_offset (against the real VehicleOrder
    # holiday calendar) so the test is self-consistent and non-circular for the OTHER assertions.
    of.system = jython_shim._System()
    shipOffset = of.get_ship_offset(PART, runDate, "VehicleOrder", INV)
    shipDate = (runDate + datetime.timedelta(days=shipOffset)).strftime("%Y%m%d")
    rep.check("ship offset is non-zero (renban-group ship-days 13 working days -> GetShip scan applied)",
              shipOffset > 0, "offset=%d shipDate=%s" % (shipOffset, shipDate))
    # NON-CIRCULAR ship-date anchor (adversary BLOCKER 2 method flaw): the rebuild's get_ship_offset is
    # checked AGAINST a legacy-derived constant, not only against itself. The adversary independently
    # re-compiled GetShip (FPC) for THIS input — part 426070205000 (renban group 7, Monday lead 13) on
    # runDate 2026-06-22 against the REAL VehicleOrder 'H' window (2026-07-03 + the 07-13..17 summer
    # shutdown) — and got offset 18 -> ship 20260710. If the rebuild ever drifts from GetShip on this
    # fixture, this FAILS (it is not a rebuild-vs-rebuild self-consistency check). The fixture is pinned
    # to the spike's VehicleOrder calendar; if that calendar changes, update LEGACY_SHIP_* with a fresh
    # GetShip re-derivation (do not just copy the rebuild's output).
    LEGACY_SHIP_OFFSET = 18                      # GetShip(lead 13, Mon 2026-06-22, real 'H' window), FPC-proved
    LEGACY_SHIP_DATE = "20260710"
    rep.check("ship offset == legacy GetShip offset 18 (FPC-derived, non-circular vs rebuild — BLOCKER 2)",
              shipOffset == LEGACY_SHIP_OFFSET, "rebuild offset=%d, legacy GetShip=%d (FPC); "
              "if the spike VehicleOrder 'H' window changed, re-derive — DO NOT copy the rebuild value"
              % (shipOffset, LEGACY_SHIP_OFFSET))
    rep.check("ship date == legacy GetShip ship date %s (non-circular)" % LEGACY_SHIP_DATE,
              shipDate == LEGACY_SHIP_DATE, "rebuild=%s legacy=%s" % (shipDate, LEGACY_SHIP_DATE))
    tmproot = tempfile.mkdtemp(prefix="orderfile_e2e_")

    try:
        # ============ (1) HAPPY PATH: emit, assert .ord bytes, LocalFTP=False -> ONE copy =============
        print("\n--- (1) emit .ord (LocalFTP=False -> supplier dir only) + exact bytes ---")
        seed_orders(["98R001", "98R002"], ORDER_QTY)
        of.system = jython_shim._System()
        supdir = os.path.join(tmproot, "sup_off")
        res = of.generate_order_files(database=INV, calDatabase="VehicleOrder", localFtp=False,
                                      runDate=runDate, dirOverride={SUPCODE: supdir})
        rep.check("driver processed exactly our 2 synthetic rows", res["rowCount"] == 2, str(res["rowCount"]))
        rep.check("LocalFTP=False -> exactly ONE file published (supplier dir only)",
                  len(res["files"]) == 1, "files=%s" % res["files"])
        published = res["files"][0]
        rep.check("published file is a final .ord (NOT a .tmp)",
                  published.endswith(".ord") and os.path.exists(published) and
                  not os.path.exists(published + ".tmp"), published)
        # BLOCKER-1 + non-circular FILENAME assertion (adversary BLOCKER 1 + 2): supplier 11111 is
        # PACIFIC_MFG with BIT_ORDER_FILE_TIMESTAMP=0 on the live/spike config. The legacy no-timestamp
        # branch (OrderFormCreateF.pas:321) is the STABLE name <name>-<sup>.ord, NOT a timestamped one.
        # We compare the PUBLISHED basename to the legacy-derived expectation (independent of the rebuild's
        # own _ord_filename), so a regression to bool('0')==True (timestamped) FAILS here.
        legacy_name = "PACIFIC_MFG-11111.ord"      # legacy :321 form for sup 11111 (PACIFIC_MFG, bit 0)
        rep.check("published .ord NAME = legacy stable '%s' (bit 0 -> NO timestamp — BLOCKER-1)"
                  % legacy_name, os.path.basename(published) == legacy_name,
                  "got=%r want=%r" % (os.path.basename(published), legacy_name))
        rep.check("published .ord NAME carries NO yyyymmddhhmmss timestamp (bit 0 stable name)",
                  orderDateStr not in os.path.basename(published)[len("PACIFIC_MFG-11111"):],
                  os.path.basename(published))

        body = open(published, "rb").read().decode("ascii")
        # expected bytes: each row uses the computed ship date (shipOffset from the renban-group lead)
        want_lines = [of.format_ord_line({"supCode": SUPCODE, "frsNum": FRS_PREFIX + "00001",
                                          "renban": rb, "partNum": PART, "qty": ORDER_QTY,
                                          "shipDate": shipDate}) for rb in ["98R001", "98R002"]]
        want = "\r\n".join(want_lines) + "\r\n"
        rep.check("published .ord bytes are EXACTLY format_ord_line per row, CRLF-terminated",
                  body == want, "got=%r\n      want=%r" % (body, want))
        rep.check("the emitted qty is the ORDER qty %d, NOT the on-hand %d (aliasing — H10)"
                  % (ORDER_QTY, onhand),
                  ("%05d" % ORDER_QTY) in body and ("%05d" % onhand) not in body,
                  "order=%s onhand=%s" % ("%05d" % ORDER_QTY, "%05d" % onhand))
        # The trailing 8 bytes of each line ARE the legacy-derived ship date (non-circular — anchored to
        # the FPC-proved GetShip constant above, not the rebuild's own formatter output).
        line0 = body.split("\r\n")[0]
        rep.check("the .ord ship-date field (last 8 bytes) == legacy GetShip date %s (BLOCKER 2)"
                  % LEGACY_SHIP_DATE, line0[-8:] == LEGACY_SHIP_DATE, "got=%r" % line0[-8:])

        # ============ (4) EMIT-THEN-STAMP: rows now order-dated; a 2nd run emits NOTHING ==============
        print("\n--- (4) emit-then-stamp: rows stamped; re-run is a no-op (re-emit guard, H11) ---")
        stamped = sql("SELECT VC_RENBAN_NUMBER, VC_ORDER_DATE, VC_SHIP_DATE FROM INV_OPEN_ORDER_INF "
                      "WHERE VC_ADD='%s' ORDER BY VC_RENBAN_NUMBER" % ADD_STAMP)
        rep.check("BOTH same-part+FRS rows stamped with the order date (one stamp marks all — H11)",
                  all(r[1] == orderDateStr for r in stamped) and len(stamped) == 2, str(stamped))
        rep.check("VC_SHIP_DATE stamped too (= computed ship date from the renban-group lead)",
                  all(r[2] == shipDate for r in stamped), "shipDate=%s rows=%s" % (shipDate, stamped))

        of.system = jython_shim._System()
        supdir2 = os.path.join(tmproot, "sup_rerun")
        res2 = of.generate_order_files(database=INV, calDatabase="VehicleOrder", localFtp=False,
                                       runDate=runDate, dirOverride={SUPCODE: supdir2})
        rep.check("2nd run emits NOTHING (the stamped rows dropped out of the feed)",
                  res2["rowCount"] == 0 and res2["files"] == [], str(res2))
        rep.check("no .ord written on the no-op re-run",
                  not os.path.isdir(supdir2) or list_ord(supdir2) == [], "dir=%s" % supdir2)

        cleanup()

        # ============ (2) 3-DESTINATION FAN-OUT: LocalFTP=True -> supplier + logistics + archive =======
        print("\n--- (2) LocalFTP=True fan-out (supplier + logistics + archive; 'NONE' skips logistics) ---")
        seed_orders(["98R010"], ORDER_QTY)
        of.system = jython_shim._System()
        ftpdir = os.path.join(tmproot, "ftp_on")
        logdir = os.path.join(tmproot, "logistics")
        # Force a real (non-'NONE') logistics dir by stubbing resolve_logistics_dir for this scenario.
        orig_resolve = of.resolve_logistics_dir
        of.resolve_logistics_dir = lambda partNum, supplierInfo, database=None: logdir
        try:
            res3 = of.generate_order_files(database=INV, calDatabase="VehicleOrder", localFtp=True,
                                           runDate=runDate, dirOverride={SUPCODE: ftpdir})
        finally:
            of.resolve_logistics_dir = orig_resolve
        kinds = sorted(d["kind"] for s in res3["suppliers"] for d in s["destinations"])
        rep.check("LocalFTP=True -> THREE destinations (supplier, logistics, archive)",
                  kinds == ["archive", "logistics", "supplier"], str(kinds))
        rep.check("3 final .ord files published (one per destination)",
                  len(res3["files"]) == 3, "files=%d" % len(res3["files"]))
        rep.check("a copy landed in the logistics dir",
                  any(p.startswith(logdir) for p in res3["files"]), str(res3["files"]))
        rep.check("a copy landed in <supplierDir>/Archive",
                  any(os.sep + "Archive" + os.sep in p for p in res3["files"]), str(res3["files"]))
        # all 3 copies byte-identical
        contents = set(open(p, "rb").read() for p in res3["files"])
        rep.check("all 3 destination copies are byte-identical", len(contents) == 1, "distinct=%d" % len(contents))
        cleanup()

        # ============ (2b) 'NONE' logistics sentinel SKIPS the logistics copy =========================
        print("\n--- (2b) LocalFTP=True + logistics 'NONE' -> logistics copy SKIPPED (Q8) ---")
        seed_orders(["98R020"], ORDER_QTY)
        of.system = jython_shim._System()
        nonedir = os.path.join(tmproot, "ftp_none")
        of.resolve_logistics_dir = lambda partNum, supplierInfo, database=None: "NONE"
        try:
            res4 = of.generate_order_files(database=INV, calDatabase="VehicleOrder", localFtp=True,
                                           runDate=runDate, dirOverride={SUPCODE: nonedir})
        finally:
            of.resolve_logistics_dir = orig_resolve
        kinds4 = sorted(d["kind"] for s in res4["suppliers"] for d in s["destinations"])
        rep.check("'NONE' logistics -> supplier + archive only (logistics SKIPPED)",
                  kinds4 == ["archive", "supplier"], str(kinds4))
        cleanup()

        # ============ (5) ATOMICITY: DB failure after .tmp write -> NO final .ord, rows un-stamped =====
        print("\n--- (5) atomicity: inject a stamp-tx failure AFTER .tmp write -> NO final .ord (H1 fix) ---")
        seed_orders(["98R030", "98R031"], ORDER_QTY)
        sysobj = jython_shim._System()
        of.system = sysobj
        atomdir = os.path.join(tmproot, "atom")
        real_update = sysobj.db.runPrepUpdate
        spy = {"tmp_existed_before_fail": None}

        def faulty_update(sqltext, args, db=None, getKey=False, tx=None):
            # the stamp runs UPDATE_ORDEROrderDate; force the FIRST stamp to fail. Before failing, confirm
            # the .tmp WAS written (step 2 ran) -> proves a real rollback of a written file, not never-ran.
            # Use RAISERROR (severity 16) on the SAME live tx so the error surfaces as a SqlError ->
            # commitOrders-style rollback, WITHOUT touching any table (a bad INSERT would fire the
            # INV_OPEN_ORDER_INF INSERT trigger / write _HIST, muddying the un-stamped-rows assertion).
            tmps = glob.glob(os.path.join(atomdir, "**", "*.tmp"), recursive=True)
            spy["tmp_existed_before_fail"] = len(tmps)
            return tx.exec_update("RAISERROR ('order-file atomicity test fault', 16, 1)")

        sysobj.db.runPrepUpdate = faulty_update
        threw = False
        try:
            of.generate_order_files(database=INV, calDatabase="VehicleOrder", localFtp=False,
                                    runDate=runDate, dirOverride={SUPCODE: atomdir})
        except Exception as e:
            threw = True
            print("    generate_order_files raised (expected): %s" % str(e)[:70])
        rep.check("driver re-raised the mid-stamp failure", threw)
        rep.check("the .tmp WAS written before the failure (non-vacuous: a real file got rolled back)",
                  spy["tmp_existed_before_fail"] >= 1, "tmps=%r" % spy["tmp_existed_before_fail"])
        final_ords = glob.glob(os.path.join(atomdir, "**", "*.ord"), recursive=True)
        leftover_tmps = glob.glob(os.path.join(atomdir, "**", "*.tmp"), recursive=True)
        rep.check("NO final .ord exists after the rolled-back run (legacy H1 FIXED)",
                  final_ords == [], "ords=%s" % final_ords)
        rep.check("orphan .tmp cleaned up too (no leftover staged file)",
                  leftover_tmps == [], "tmps=%s" % leftover_tmps)
        # Read past the still-closing rollback connection's locks (harness artifact — see sql_uncommitted).
        # The synthetic rows were autocommitted BEFORE the driver, so a dirty read of their order-date is
        # correct: they must persist UN-stamped (the stamp tx rolled back), so a re-run would re-emit them.
        rb_stamp = sql_uncommitted("SELECT VC_RENBAN_NUMBER, VC_ORDER_DATE FROM INV_OPEN_ORDER_INF "
                                   "WHERE VC_ADD='%s'" % ADD_STAMP)
        rep.check("open-order rows PERSIST UN-stamped after rollback (a re-run would re-emit)",
                  all(r[1] in ("", None) for r in rb_stamp) and len(rb_stamp) == 2, str(rb_stamp))
        cleanup()

        # ============ (6) RISK-1: rename #2 of 3 fails POST-COMMIT -> NO silent partial emit ===========
        # The stamp tx has already committed when the .tmp -> final renames run. If destination #2 (the
        # logistics share) fails mid-fan-out, the driver must NOT leave a silent partial (supplier .ord
        # out, logistics/archive lost): it un-renames the already-published copy and raises a LOUD error
        # naming the retained .tmp files. We assert: it raised, NO final .ord remains at ANY destination
        # (all-or-nothing), and every .tmp is retained for operator re-drop.
        print("\n--- (6) RISK-1: rename #2 of 3 fails post-commit -> all-or-nothing, loud, no partial ---")
        seed_orders(["98R040"], ORDER_QTY)
        of.system = jython_shim._System()
        renamedir = os.path.join(tmproot, "rename_fail")
        rlogdir = os.path.join(tmproot, "rename_log")
        of.resolve_logistics_dir = lambda partNum, supplierInfo, database=None: rlogdir
        # Fail the SECOND os.rename (.tmp -> final) — i.e. one of the 3 destinations mid fan-out. The
        # driver does `import os as _os` -> the real os module, so patching os.rename hits it.
        real_rename = os.rename
        rn_calls = {"n": 0}

        def faulty_rename(src, dst):
            rn_calls["n"] += 1
            if rn_calls["n"] == 2:                # the 2nd publish (e.g. logistics) fails
                raise OSError("simulated logistics share down on rename #2")
            return real_rename(src, dst)

        os.rename = faulty_rename
        threw6 = False
        try:
            res6 = of.generate_order_files(database=INV, calDatabase="VehicleOrder", localFtp=True,
                                           runDate=runDate, dirOverride={SUPCODE: renamedir})
        except Exception as e6:
            threw6 = True
            err6 = str(e6)
            print("    generate_order_files raised (expected RISK-1): %s" % err6[:90])
        finally:
            os.rename = real_rename
            of.resolve_logistics_dir = orig_resolve
        rep.check("RISK-1: a failed rename in the fan-out raised LOUDLY (no silent partial)", threw6,
                  "FAIL means generate_order_files did NOT raise on the simulated rename #2 failure")
        rep.check("RISK-1: the error names the retained .tmp files for operator re-drop",
                  threw6 and ".tmp" in err6 and ("STAMPED" in err6 or "stamped" in err6), err6 if threw6 else "")
        # all-or-nothing: NO final .ord at ANY of the 3 destinations (supplier/logistics/archive)
        final6 = (glob.glob(os.path.join(renamedir, "**", "*.ord"), recursive=True) +
                  glob.glob(os.path.join(rlogdir, "**", "*.ord"), recursive=True))
        rep.check("RISK-1: NO final .ord published at ANY destination (all-or-nothing publish)",
                  final6 == [], "leaked .ord=%s" % final6)
        # every .tmp retained (so an operator can re-drop) — supplier + logistics + archive = 3 .tmp
        tmp6 = (glob.glob(os.path.join(renamedir, "**", "*.tmp"), recursive=True) +
                glob.glob(os.path.join(rlogdir, "**", "*.tmp"), recursive=True))
        rep.check("RISK-1: all 3 .tmp staged files retained for re-drop (none deleted)",
                  len(tmp6) == 3, "tmp count=%d files=%s" % (len(tmp6), tmp6))
        cleanup()
    finally:
        cleanup()
        shutil.rmtree(tmproot, ignore_errors=True)

    leftover = int(scalar("SELECT COUNT(*) FROM INV_OPEN_ORDER_INF WHERE VC_ADD='%s'" % ADD_STAMP))
    leftover_hist = int(scalar("SELECT COUNT(*) FROM INV_OPEN_ORDER_INF_HIST WHERE VC_ADD='%s'" % ADD_STAMP))
    rep.check("spike restored as found (no synthetic orders / hist rows remain)",
              leftover == 0 and leftover_hist == 0, "orders=%d hist=%d" % (leftover, leftover_hist))
    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
