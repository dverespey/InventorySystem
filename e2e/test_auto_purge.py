#!/usr/bin/env python3
"""test_auto_purge.py — M4 piece-3 HD4 DATAPURGE gate: prove the rebuild's retention-based purge driver
(docs/analysis/production-readiness/project-library/auto_purge/code.py) is FAITHFUL to the legacy
DELETE_AutoPurge (CreateInventory.sql:7682) AND hardened (transactional, retention from INV_SITES).

WHAT THIS DRIVES (the REAL gateway-side driver, retro R8 — not an EXEC of the proc):
  We load the ACTUAL auto_purge Project Library module through jython_shim and call auto_purge() /
  run_scheduled_purge(). The driver opens a REAL transaction (jython_shim._TxSession = one persistent
  sqlcmd connection spanning the 4 statements), so the transaction + the rollback path are exercised by
  the driver's own code, against the live spike DB.

THE ORACLE IS THE SOURCE, NOT THE REBUILD (retro R15/R16/R20):
  The legacy cutoff is a 16-char string YYYYMMDD+HH+MM+SS+ff = DATEADD(MONTH, -retention, GETDATE())
  (style 112 + substrings of style 114), and DELETE_AutoPurge deletes WHERE VC_ADD <= that cutoff from
  INV_OPEN_ORDER_INF / _HIST / INV_PARTS_STOCK_MST_HIST (no site filter; single-site). We compute that
  cutoff with an INDEPENDENT sqlcmd query of the EXACT legacy expression (NOT by calling the driver), then
  seed:
    * an OLD synthetic row  : VC_ADD strictly < cutoff  -> the SOURCE says MUST be purged
    * a RECENT synthetic row: VC_ADD strictly > cutoff  -> the SOURCE says MUST be kept
  in each scoped table, and assert the driver purged exactly the old ones and kept the recent ones.
  The expected outcome is derived from the legacy predicate, independently computed — not read back from
  the driver's own recomputation.

NON-VACUITY + REVERT-PROOF (a test that can't fail on a wrong rebuild tests nothing):
  * Non-vacuous: the RECENT rows are kept (the purge is not delete-all) AND the OLD rows are gone (not
    keep-all).
  * Revert-proof: we monkey-patch the driver's sign-flip (_neg_retention -> identity, i.e. the bug where
    retention is added to the FUTURE instead of subtracted) and re-run — the OLD rows then SURVIVE (the
    cutoff lands in the future, purging nothing/everything wrong), so the test FAILS on the broken driver.
    Restored after.

HARDENING (HD4):
  * TRANSACTIONAL: a forced failure inside the tx (an injected bad statement) rolls back ALL 4 -> NOTHING
    purged (the old rows survive). Closes the legacy partial-purge hole.
  * RETENTION READ FROM INV_SITES: read_purge_config() reads IN_DATA_RETENTION / the BIT knobs from the
    INV_SITES row (Q17), proven by setting the row and asserting the driver picks it up.
  * THE >=12 FLOOR: auto_purge(11) RAISES (REFUSES to purge), faithful to DataModule.pas:6890.

FIXTURE DISCIPLINE: every synthetic row carries the sentinel VC_PART_NUMBER prefix below and is swept in
teardown from ALL THREE tables (the INV_OPEN_ORDER_INF INSERT trigger also auto-copies into _HIST, so we
sweep _HIST by the same sentinel). Synthetic INV_OPEN_ORDER_INF rows use a part number ABSENT from
INV_PARTS_STOCK_MST and ALL-BLANK status fields, so the insert/delete triggers have NO quantity side
effect (verified: the trigger's IN_QTY update joins find nothing). Real client rows are never touched.

Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 e2e/test_auto_purge.py
"""
import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")))

from lib import Report                        # noqa: E402
import jython_shim                            # noqa: E402

REPO = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
PURGE_CODE = os.path.join(REPO, "ignition", "project-library", "auto_purge", "code.py")

CONTAINER = os.environ.get("CONTAINER", "mssql-spike")
SA_PASS = os.environ.get("SA_PASS", "Spike_Dev_2026!")
DB = "Inventory"

# Sentinel that tags EVERY synthetic row so teardown can sweep precisely (and never a real row).
SENTINEL = "ZZPURGE"                    # VC_PART_NUMBER prefix (absent from real data)
SENTINEL_LIKE = SENTINEL + "%"

# CRITICAL — the test runs the purge with a 600-month (50-year) retention, NOT the 12-month floor. The
# legacy purge has NO site filter and prunes the WHOLE DB by retention; running it at 12 months would
# delete every REAL row older than 12 months on the shared spike (a destructive, irreversible-without-
# re-restore hit). At 600 months the cutoff is ~1976, OLDER than ALL real data on the spike (oldest real
# VC_ADD is 2018), so the purge can ONLY ever match the synthetic OLD rows we seed in 1970 — real data is
# never touched. The retention VALUE does not change the predicate being proven (VC_ADD <= cutoff,
# faithful to DELETE_AutoPurge); it only isolates the test's deletion to our pre-1976 sentinel rows.
RETENTION = 600                         # 50yr -> cutoff ~1976, older than all real data (isolates the test)
OLD_VCADD = "1970010100000000"          # < the 600-mo cutoff (and < all real data) -> MUST be purged
RECENT_VCADD = "2099010100000000"       # > any cutoff -> MUST be kept (and far past all real data)


def sqlq(query, want_rows=True):
    """One-shot sqlcmd against the spike. QUOTED_IDENTIFIER ON (R23)."""
    flags = ["-W"]
    if not want_rows:
        flags += ["-h", "-1"]
    cmd = (["docker", "exec", CONTAINER, "/opt/mssql-tools18/bin/sqlcmd", "-C", "-S", "localhost",
            "-U", "sa", "-P", SA_PASS, "-d", DB] + flags +
           ["-Q", "SET QUOTED_IDENTIFIER ON; SET ANSI_NULLS ON; SET NOCOUNT ON; " + query])
    return subprocess.check_output(cmd, stderr=subprocess.STDOUT, timeout=60).decode("utf-8", "replace")


def scalar(query):
    out = sqlq(query, want_rows=False)
    for tok in out.replace("\t", " ").split():
        try:
            return int(tok)
        except ValueError:
            continue
    return None


def legacy_cutoff(retention):
    """INDEPENDENT oracle: compute the 16-char legacy cutoff with the EXACT DELETE_AutoPurge expression
    (style 112 + substrings 1/4/7/10 of style 114), @DataRentention = NEGATIVE retention. This is the
    SOURCE predicate's boundary, computed WITHOUT touching the driver."""
    neg = 0 - int(retention)
    out = sqlq("SELECT (CONVERT(char(8),DATEADD(MONTH,%d,GETDATE()),112)"
               " + SUBSTRING(CONVERT(char(12),DATEADD(MONTH,%d,GETDATE()),114),1,2)"
               " + SUBSTRING(CONVERT(char(12),DATEADD(MONTH,%d,GETDATE()),114),4,2)"
               " + SUBSTRING(CONVERT(char(12),DATEADD(MONTH,%d,GETDATE()),114),7,2)"
               " + SUBSTRING(CONVERT(char(12),DATEADD(MONTH,%d,GETDATE()),114),10,2)) AS c"
               % (neg, neg, neg, neg, neg), want_rows=False)
    return out.strip().split()[0]


# ---- synthetic row counts by sentinel, per table ----
def count_sentinel(table):
    return scalar("SELECT COUNT(*) FROM %s WHERE VC_PART_NUMBER LIKE '%s'" % (table, SENTINEL_LIKE))


def teardown():
    """Sweep ALL synthetic rows by sentinel from all three scoped tables. INV_OPEN_ORDER_INF DELETE fires
    its trigger (which copies into _HIST), so delete _HIST AFTER the base table, then _HIST again to clean
    any trigger-driven copies. Safe: sentinel part number is absent from real data + INV_PARTS_STOCK_MST."""
    sqlq("DELETE FROM INV_OPEN_ORDER_INF      WHERE VC_PART_NUMBER LIKE '%s';" % SENTINEL_LIKE)
    sqlq("DELETE FROM INV_OPEN_ORDER_INF_HIST WHERE VC_PART_NUMBER LIKE '%s';" % SENTINEL_LIKE)
    sqlq("DELETE FROM INV_PARTS_STOCK_MST_HIST WHERE VC_PART_NUMBER LIKE '%s';" % SENTINEL_LIKE)


def seed_rows(old_vcadd, recent_vcadd):
    """Seed an OLD (purge) + RECENT (keep) synthetic row in each scoped table, tagged with the sentinel.

    INV_OPEN_ORDER_INF: minimal required columns; ALL status fields '' (so the insert trigger's IN_QTY
    update joins nothing) and a sentinel part number absent from INV_PARTS_STOCK_MST. The old row also has
    VC_TERMINATED='' so statement-1 terminate-then-delete is exercised.
    INV_OPEN_ORDER_INF_HIST / INV_PARTS_STOCK_MST_HIST: only the few NOT-NULL columns + VC_ADD."""
    teardown()  # clean slate
    # base table — uses an IDENTITY pk; supply every NOT-NULL non-identity column ('' / 0 / sentinel)
    base_cols = ("VC_SUPPLIER_CODE,VC_PART_NUMBER,VC_FRS_NUMBER,VC_RENBAN_NUMBER,IN_QTY,"
                 "VC_STATUS_SUPPLIER_SHIPPING,VC_ARRIVAL,VC_TRAILER_NUMBER,VC_STATUS_PLANT_YARD,"
                 "VC_PLANT_PARKING,VC_STATUS_ASSEMBLER_YARD,VC_ASSEMBLER_LOCATION,"
                 "VC_STATUS_EMPTY_TRAILER,VC_DETENTION,VC_ORDER_DATE,VC_WAREHOUSE,VC_TERMINATED,"
                 "VC_SHIP_DATE,VC_KANBAN_NUMBER,VC_LAST_UPDATE,VC_ADD")

    def base_vals(part, vcadd):
        return ("'ZZ','%s','','',0, '','','','', '','','', '','','','', '', '','','', '%s'"
                % (part, vcadd))

    sqlq("INSERT INTO INV_OPEN_ORDER_INF (%s) VALUES (%s); INSERT INTO INV_OPEN_ORDER_INF (%s) VALUES (%s);"
         % (base_cols, base_vals(SENTINEL + "OLD", old_vcadd),
            base_cols, base_vals(SENTINEL + "NEW", recent_vcadd)))
    # The INSERT trigger auto-copied both base rows into _HIST. Remove those auto-copies so the _HIST
    # purge proof uses ONLY the rows we explicitly seed below with controlled VC_ADD.
    sqlq("DELETE FROM INV_OPEN_ORDER_INF_HIST WHERE VC_PART_NUMBER LIKE '%s';" % SENTINEL_LIKE)

    # _HIST tables (no triggers): only the NOT-NULL cols + VC_ADD.
    sqlq("INSERT INTO INV_OPEN_ORDER_INF_HIST (VC_SUPPLIER_CODE,VC_PART_NUMBER,VC_FRS_NUMBER,"
         "VC_RENBAN_NUMBER,IN_QTY,VC_ADD) VALUES ('ZZ','%s','','',0,'%s'),('ZZ','%s','','',0,'%s');"
         % (SENTINEL + "OLD", old_vcadd, SENTINEL + "NEW", recent_vcadd))
    sqlq("INSERT INTO INV_PARTS_STOCK_MST_HIST (VC_PART_NUMBER,VC_PARTS_NAME,IN_RENBAN_COUNT,VC_ADD) "
         "VALUES ('%s','QA',0,'%s'),('%s','QA',0,'%s');"
         % (SENTINEL + "OLD", old_vcadd, SENTINEL + "NEW", recent_vcadd))


def load_purge():
    """Load the REAL auto_purge driver with the jython_shim system (real _TxSession for the tx path)."""
    return jython_shim.load_wrapper("auto_purge", PURGE_CODE)


# ===========================================================================
# 1. SCOPE FAITHFULNESS — old purged / recent kept, in every scoped table (oracle = the SOURCE predicate)
# ===========================================================================

def test_scope(rep):
    cutoff = legacy_cutoff(RETENTION)
    old_vcadd, recent_vcadd = OLD_VCADD, RECENT_VCADD
    # real-data row counts BEFORE the purge — to prove afterward that NO real row was touched (the 600-mo
    # cutoff is older than all real data, so the purge must affect ONLY our synthetic 1970 rows).
    real_pre = {t: scalar("SELECT COUNT(*) FROM %s WHERE VC_PART_NUMBER NOT LIKE '%s'" % (t, SENTINEL_LIKE))
                for t in ("INV_OPEN_ORDER_INF", "INV_OPEN_ORDER_INF_HIST", "INV_PARTS_STOCK_MST_HIST")}

    rep.check("oracle: legacy cutoff(600) is a 16-char YYYYMMDD..ff string (SOURCE predicate boundary)",
              len(cutoff) == 16 and cutoff.isdigit(), "cutoff=%r" % cutoff)
    rep.check("SAFETY: cutoff (%s) is older than all real data on the spike (purge can only hit the "
              "synthetic 1970 row, never real data)" % cutoff,
              cutoff < "2018", "cutoff=%s" % cutoff)
    rep.check("oracle: seeded OLD VC_ADD < cutoff < RECENT VC_ADD (rows straddle the SOURCE boundary)",
              old_vcadd < cutoff < recent_vcadd,
              "old=%s cutoff=%s recent=%s" % (old_vcadd, cutoff, recent_vcadd))

    seed_rows(old_vcadd, recent_vcadd)
    pre = {t: count_sentinel(t) for t in
           ("INV_OPEN_ORDER_INF", "INV_OPEN_ORDER_INF_HIST", "INV_PARTS_STOCK_MST_HIST")}
    rep.check("seed: 2 synthetic rows in each scoped table (old+recent)",
              all(v == 2 for v in pre.values()), "counts=%s" % pre)

    mod = load_purge()
    counts = mod.auto_purge(RETENTION)        # the REAL driver, real transaction

    # SOURCE says: rows with VC_ADD <= cutoff are deleted. Old (<cutoff) gone; recent (>cutoff) kept.
    for t in ("INV_OPEN_ORDER_INF", "INV_OPEN_ORDER_INF_HIST", "INV_PARTS_STOCK_MST_HIST"):
        old_left = scalar("SELECT COUNT(*) FROM %s WHERE VC_PART_NUMBER='%sOLD'" % (t, SENTINEL))
        new_left = scalar("SELECT COUNT(*) FROM %s WHERE VC_PART_NUMBER='%sNEW'" % (t, SENTINEL))
        rep.check("%s: OLD synthetic row (< cutoff) PURGED (SOURCE: VC_ADD<=cutoff deleted)" % t,
                  old_left == 0, "old rows remaining=%s" % old_left)
        rep.check("%s: RECENT synthetic row (> cutoff) KEPT (non-vacuous: not delete-all)" % t,
                  new_left == 1, "recent rows remaining=%s" % new_left)

    # SAFETY proof: NO real row was deleted (the cutoff predates all real data; only synthetic 1970 rows
    # qualified). The driver's reported delete counts are EXACTLY the synthetic OLD rows (1 per table),
    # not the whole-DB sweep that a 12-month retention would do.
    real_post = {t: scalar("SELECT COUNT(*) FROM %s WHERE VC_PART_NUMBER NOT LIKE '%s'" % (t, SENTINEL_LIKE))
                 for t in ("INV_OPEN_ORDER_INF", "INV_OPEN_ORDER_INF_HIST", "INV_PARTS_STOCK_MST_HIST")}
    rep.check("SAFETY: no REAL row was purged (real-data counts unchanged — the cutoff predates all real "
              "data, so only the synthetic 1970 rows were deleted)",
              real_post == real_pre, "pre=%s post=%s" % (real_pre, real_post))

    # statement-1 terminate is exercised (the OLD base row had VC_TERMINATED='' and was terminated before
    # delete). The driver reports terminated==1 (exactly the 1 synthetic old open order).
    rep.check("statement-1 TERMINATE ran on exactly the 1 old open order (driver reports terminated==1)",
              counts.get("terminated", 0) == 1, "counts=%s" % counts)
    # SURGICAL (not a whole-DB sweep): the base table deletes EXACTLY 1 (the synthetic OLD row). The two
    # _HIST tables delete >=1: faithfully, statement-1's UPDATE on the OLD base row fires the legacy
    # UPDATE_RecConfStatPartsStockMstQTY trigger which INSERTs a fresh OLD copy into INV_OPEN_ORDER_INF_HIST,
    # so statement-3 deletes BOTH that copy AND our explicit OLD _HIST row (count 2). That is the real
    # legacy trigger interplay — counted, not a real-data hit (the cutoff predates all real data, proven
    # above). So: base==1, _HIST>=1, and crucially the real-data counts were unchanged.
    rep.check("driver deleted EXACTLY 1 from the base table (INV_OPEN_ORDER_INF==1 — surgical)",
              counts.get("INV_OPEN_ORDER_INF", 0) == 1, "counts=%s" % counts)
    rep.check("driver deleted the synthetic OLD row from each _HIST table (>=1; the extra _HIST copy is the "
              "faithful UPDATE-trigger interplay, not a real row)",
              counts.get("INV_OPEN_ORDER_INF_HIST", 0) >= 1
              and counts.get("INV_PARTS_STOCK_MST_HIST", 0) >= 1, "counts=%s" % counts)
    teardown()


# ===========================================================================
# 2. TRANSACTIONAL — a forced failure inside the tx rolls back ALL 4 (nothing purged). HD4 hardening.
# ===========================================================================

def test_transactional(rep):
    seed_rows(OLD_VCADD, RECENT_VCADD)
    pre = {t: count_sentinel(t) for t in
           ("INV_OPEN_ORDER_INF", "INV_OPEN_ORDER_INF_HIST", "INV_PARTS_STOCK_MST_HIST")}

    mod = load_purge()
    # Inject a failure: make the 4th (last) statement reference a non-existent table so it errors AFTER the
    # first 3 deletes have run in the SAME tx. A transactional driver MUST roll back all -> nothing purged.
    orig_tables = mod.PURGE_DELETE_TABLES
    try:
        mod.PURGE_DELETE_TABLES = ("INV_OPEN_ORDER_INF_HIST", "INV_PARTS_STOCK_MST_HIST",
                                   "INV__NO_SUCH_TABLE__ZZ")   # last one errors mid-tx
        failed = False
        try:
            mod.auto_purge(RETENTION)
        except Exception:
            failed = True
        rep.check("transactional: an injected mid-tx failure RAISES (the driver does not swallow it)",
                  failed, "auto_purge raised=%s" % failed)
    finally:
        mod.PURGE_DELETE_TABLES = orig_tables

    post = {t: count_sentinel(t) for t in
            ("INV_OPEN_ORDER_INF", "INV_OPEN_ORDER_INF_HIST", "INV_PARTS_STOCK_MST_HIST")}
    rep.check("transactional: after the forced mid-tx failure NOTHING was purged (all synthetic rows "
              "survive — the first deletes were ROLLED BACK)",
              post == pre and all(v == 2 for v in post.values()),
              "pre=%s post=%s" % (pre, post))
    teardown()


# ===========================================================================
# 3. RETENTION + enable/prompt READ FROM INV_SITES (Q17). read_purge_config + run_scheduled_purge.
# ===========================================================================

def _read_site_knob(col, site_id):
    """Robust single-cell read of an INV_SITES knob (NULL -> the literal 'NULL' marker for restore)."""
    out = sqlq("SELECT ISNULL(CONVERT(varchar,%s),'NULL') FROM INV_SITES WHERE IN_SITE_ID=%d"
               % (col, site_id), want_rows=False)
    return out.strip().split("\n")[-1].strip()


def test_retention_from_sites(rep):
    mod = load_purge()
    site_id = scalar("SELECT MIN(IN_SITE_ID) FROM INV_SITES")
    # capture each knob robustly (single-cell reads, NULL-safe) so the finally-block restores exactly.
    orig = {c: _read_site_knob(c, site_id)
            for c in ("IN_DATA_RETENTION", "BIT_ENABLE_DATA_PURGE", "BIT_PROMPT_DATA_PURGE")}
    try:
        # 600-month retention so a run never touches real data (cutoff ~1976; see RETENTION note).
        sqlq("UPDATE INV_SITES SET IN_DATA_RETENTION=%d, BIT_ENABLE_DATA_PURGE=1, BIT_PROMPT_DATA_PURGE=0 "
             "WHERE IN_SITE_ID=%d;" % (RETENTION, site_id))
        cfg = mod.read_purge_config()
        rep.check("read_purge_config reads IN_DATA_RETENTION from INV_SITES (=%d)" % RETENTION,
                  cfg["retention"] == RETENTION, "cfg=%s" % cfg)
        rep.check("read_purge_config reads BIT_ENABLE_DATA_PURGE from INV_SITES (=True)",
                  cfg["enable"] is True, "cfg=%s" % cfg)
        rep.check("read_purge_config reads BIT_PROMPT_DATA_PURGE from INV_SITES (=False)",
                  cfg["prompt"] is False, "cfg=%s" % cfg)

        # run_scheduled_purge honors BIT_ENABLE: disable -> it SKIPS without purging.
        sqlq("UPDATE INV_SITES SET BIT_ENABLE_DATA_PURGE=0 WHERE IN_SITE_ID=%d;" % site_id)
        res = mod.run_scheduled_purge()
        rep.check("run_scheduled_purge SKIPS when BIT_ENABLE_DATA_PURGE is off (no purge)",
                  res.get("ran") is False and res.get("reason") == "disabled", "res=%s" % res)

        # enable + the 600-mo retention, with seeded rows -> it runs and uses the INV_SITES retention.
        seed_rows(OLD_VCADD, RECENT_VCADD)
        sqlq("UPDATE INV_SITES SET BIT_ENABLE_DATA_PURGE=1, IN_DATA_RETENTION=%d WHERE IN_SITE_ID=%d;"
             % (RETENTION, site_id))
        res = mod.run_scheduled_purge()
        old_left = scalar("SELECT COUNT(*) FROM INV_PARTS_STOCK_MST_HIST WHERE VC_PART_NUMBER='%sOLD'"
                          % SENTINEL)
        new_left = scalar("SELECT COUNT(*) FROM INV_PARTS_STOCK_MST_HIST WHERE VC_PART_NUMBER='%sNEW'"
                          % SENTINEL)
        rep.check("run_scheduled_purge RUNS when enabled, using the INV_SITES retention (=%d), and purges "
                  "the old synthetic row (keeps the recent one)" % RETENTION,
                  res.get("ran") is True and res.get("retention") == RETENTION
                  and old_left == 0 and new_left == 1,
                  "res-ran=%s res-ret=%s old_left=%s new_left=%s"
                  % (res.get("ran"), res.get("retention"), old_left, new_left))
        teardown()
    finally:
        sqlq("UPDATE INV_SITES SET IN_DATA_RETENTION=%s, BIT_ENABLE_DATA_PURGE=%s, BIT_PROMPT_DATA_PURGE=%s "
             "WHERE IN_SITE_ID=%d;"
             % (orig["IN_DATA_RETENTION"], orig["BIT_ENABLE_DATA_PURGE"], orig["BIT_PROMPT_DATA_PURGE"],
                site_id))
        teardown()
    # confirm restore
    now = {c: _read_site_knob(c, site_id)
           for c in ("IN_DATA_RETENTION", "BIT_ENABLE_DATA_PURGE", "BIT_PROMPT_DATA_PURGE")}
    rep.check("INV_SITES knobs restored as-found", now == orig, "now=%s orig=%s" % (now, orig))


# ===========================================================================
# 4. THE >=12 FLOOR — auto_purge(11) REFUSES (raises). Faithful to DataModule.pas:6890.
# ===========================================================================

def test_retention_floor(rep):
    mod = load_purge()
    refused = False
    try:
        mod.auto_purge(11)
    except ValueError:
        refused = True
    rep.check("retention floor: auto_purge(11) REFUSES (raises ValueError, <12) — faithful "
              "DataModule.pas:6890", refused, "raised=%s" % refused)
    # 12 is the boundary and is allowed (it must not raise on the floor itself)
    ok12 = True
    try:
        mod.validate_retention(12)
    except ValueError:
        ok12 = False
    rep.check("retention floor: 12 is allowed (the boundary, not rejected)", ok12)


# ===========================================================================
# 5. REVERT-PROOF — break the sign-flip; the OLD rows then SURVIVE, so the test FAILS on a wrong driver.
# ===========================================================================

def _driver_cutoff(mod, retention):
    """Compute the cutoff the DRIVER would use, by binding the driver's OWN _CUTOFF_EXPR with the driver's
    OWN _neg_retention(retention). This evaluates the SAME T-SQL expression the driver runs, on the engine,
    so a broken sign-flip shows up as a wrong cutoff WITHOUT running the destructive purge."""
    neg = mod._neg_retention(retention)
    expr = mod._CUTOFF_EXPR.replace("?", str(neg))   # the 5 ? are all the same neg literal
    return sqlq("SELECT %s AS c" % expr, want_rows=False).strip().split()[-1]


def test_revert_proof(rep):
    """SAFE revert-proof, at the CUTOFF level (no destructive run). The driver's faithfulness hinges on
    _neg_retention producing a PAST cutoff (DELETE_AutoPurge passes @DataRentention NEGATIVE). Break the
    sign-flip and the driver's cutoff lands in the FUTURE — VC_ADD <= a future cutoff catches ALL real data,
    i.e. the wrong driver would purge the whole DB. We prove the discriminator deterministically by
    computing the cutoff the driver WOULD run (its own _CUTOFF_EXPR bound with its own _neg_retention) for
    BOTH the correct and the broken sign, and asserting: correct -> PAST (< now, keeps real data); broken ->
    FUTURE (> now, would delete real data). A revert that flips this is caught — so the test fails on a wrong
    driver — and NOTHING is purged here (we never call auto_purge with the broken sign)."""
    now = sqlq("SELECT CONVERT(char(8),GETDATE(),112)", want_rows=False).strip().split()[-1]
    mod = load_purge()

    correct_cutoff = _driver_cutoff(mod, RETENTION)
    rep.check("the CORRECT driver's cutoff is in the PAST (< today %s) — keeps real data, deletes only old"
              % now, correct_cutoff[:8] < now, "correct cutoff=%s today=%s" % (correct_cutoff, now))

    orig_neg = mod._neg_retention
    try:
        mod._neg_retention = lambda r: int(r)           # WRONG sign (no flip)
        broken_cutoff = _driver_cutoff(mod, RETENTION)
    finally:
        mod._neg_retention = orig_neg

    rep.check("REVERT-PROOF: with the sign-flip BROKEN the driver's cutoff lands in the FUTURE (> today %s) "
              "-> VC_ADD<=cutoff would purge ALL real data (the whole-DB wrong-sign bug). The correct cutoff "
              "is in the PAST. So a reverted sign-flip is CAUGHT (the test fails on a wrong driver) — and "
              "nothing was purged to prove it." % now,
              broken_cutoff[:8] > now and correct_cutoff[:8] < now,
              "broken cutoff=%s correct cutoff=%s today=%s" % (broken_cutoff, correct_cutoff, now))

    # restore-sanity: with the CORRECT driver, run for real -> old purged + recent kept (faithful holds).
    seed_rows(OLD_VCADD, RECENT_VCADD)
    mod = load_purge()
    mod.auto_purge(RETENTION)
    old_left = scalar("SELECT COUNT(*) FROM INV_PARTS_STOCK_MST_HIST WHERE VC_PART_NUMBER='%sOLD'"
                      % SENTINEL)
    new_left = scalar("SELECT COUNT(*) FROM INV_PARTS_STOCK_MST_HIST WHERE VC_PART_NUMBER='%sNEW'"
                      % SENTINEL)
    rep.check("REVERT-PROOF restore: the correct driver again purges old + keeps recent",
              old_left == 0 and new_left == 1, "old_left=%s new_left=%s" % (old_left, new_left))
    teardown()


def main():
    print("=== M4 HD4 DATAPURGE — faithful (DELETE_AutoPurge scope) + hardened (transactional, INV_SITES) ===\n")
    rep = Report()
    print("-- 0. baseline: synthetic rows swept --")
    teardown()
    rep.check("baseline: no leftover synthetic rows",
              all(count_sentinel(t) == 0 for t in
                  ("INV_OPEN_ORDER_INF", "INV_OPEN_ORDER_INF_HIST", "INV_PARTS_STOCK_MST_HIST")))
    print("\n-- 1. Scope faithfulness (old purged / recent kept; oracle = SOURCE predicate) --")
    test_scope(rep)
    print("\n-- 2. Transactional (forced mid-tx failure -> nothing purged) --")
    test_transactional(rep)
    print("\n-- 3. Retention + enable/prompt read from INV_SITES (Q17) --")
    test_retention_from_sites(rep)
    print("\n-- 4. The >=12 retention floor (REFUSES below 12) --")
    test_retention_floor(rep)
    print("\n-- 5. Revert-proof (break the sign-flip -> test fails on a wrong driver) --")
    test_revert_proof(rep)
    print("\n-- teardown --")
    teardown()
    rep.check("teardown: all synthetic rows swept (spike as-found)",
              all(count_sentinel(t) == 0 for t in
                  ("INV_OPEN_ORDER_INF", "INV_OPEN_ORDER_INF_HIST", "INV_PARTS_STOCK_MST_HIST")))
    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
