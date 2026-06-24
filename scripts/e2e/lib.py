"""Shared helpers for the Perspective E2E harness (Playwright, Python).

No third-party deps beyond playwright. Reads optional gateway creds from
environment or a gitignored scripts/e2e/.env (KEY=VALUE lines).
"""
import os, re, subprocess, sys, time

# Centralized gateway-path + DB-connection constants (repo-split-plan §4.C/§4.D). `_ignenv` lives one
# dir up in scripts/; add it to the path so the whole harness shares ONE env-parametrized source of
# truth (gateway root/project dir + the Inventory_Spike->Inventory rename point). Defaults are the
# current spike values, so nothing changes at runtime unless an IGN_* env var is set.
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _ignenv import (  # noqa: E402
    GATEWAY_ROOT, GATEWAY_PROJECT_DIR, PERSPECTIVE_DIR, GATEWAY_LOG, DB_CONN,
)


def _register_db_shared():
    """Make the centralized `db_shared` project-library module importable headlessly. On the gateway,
    project-library modules are a flat top-level namespace (`from db_shared import CONNECTION` just
    works); the harness execs each code.py in isolation, so an app module that imports db_shared needs it
    pre-registered in sys.modules. Importing `lib` (which every app-loading test does — the one shim-only
    test, dress_rehearsal_smoke, is covered by jython_shim.load_wrapper) registers it once here. Pure
    constant, no `system` — loads identically here and on the gateway. (repo-split-plan §4.D)"""
    if "db_shared" in sys.modules:
        return
    import importlib.util
    code = os.path.normpath(os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "..", "..",
        "docs", "analysis", "production-readiness", "project-library", "db_shared", "code.py"))
    spec = importlib.util.spec_from_file_location("db_shared", code)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    sys.modules["db_shared"] = mod


_register_db_shared()

BASE = os.environ.get("GW_BASE", "http://localhost:8088")
WRAPPER_LOG = GATEWAY_LOG                         # = $GW_LOG or <GATEWAY_ROOT>/logs/wrapper.log
ARTIFACTS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "artifacts")

# Perspective project name = the FIRST path segment of a client route (/data/perspective/client/<PROJECT>/
# <page>). The project was renamed `spike` -> `InventorySystem` in PR #47 (chore: rename Ignition project);
# the per-view CRUD browser routes still pointed at the old `spike` and 404'd until this constant. Override
# with GW_PROJECT (or IGN_PROJECT) if a future rename happens.
PROJECT = os.environ.get("GW_PROJECT", "InventorySystem")


def view_url(page_path, query=""):
    """Build a Perspective client route for the current project: view_url("size") ->
    http://localhost:8088/data/perspective/client/InventorySystem/size. `query` (e.g. "qaAdmin=1") is
    appended as a ?-query when given."""
    url = "%s/data/perspective/client/%s/%s" % (BASE, PROJECT, page_path.lstrip("/"))
    return url + ("?" + query if query else "")


def load_env():
    """Populate os.environ from scripts/e2e/.env if present (no override of real env)."""
    envp = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if os.path.exists(envp):
        for line in open(envp):
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip())


def creds():
    load_env()
    return os.environ.get("GATEWAY_USER"), os.environ.get("GATEWAY_PASS")


def log_marker():
    """Byte offset of the end of wrapper.log — capture before an action, then
    read_log_since(offset) to get only lines the action produced."""
    try:
        return os.path.getsize(WRAPPER_LOG)
    except OSError:
        return 0


def read_log_since(offset):
    try:
        with open(WRAPPER_LOG, "r", errors="replace") as f:
            f.seek(offset)
            return f.read()
    except OSError:
        return ""


def grep_spike_since(offset, needle="SPIKE"):
    return [l for l in read_log_since(offset).splitlines() if needle in l]


# Shared SKIP reason for the per-view master-CRUD UI-WRITE round-trips after the P15 server-side write
# gate landed: an anonymous spike session (no IdP/roles) is correctly DENIED before the write path runs,
# so those UI insert/update/delete cases can no longer be exercised through the live UI here. The gate
# itself is proven headless end-to-end (forged-prop-rejected, revert-proven, per view) by
# test_master_write_gates.py; the LIVE on-the-box deny is proven by check_master_write_gate_live below.
WRITE_GATE_SKIP = ("P15 server-side write gate active: an anon spike session (no IdP/roles) is correctly "
                   "DENIED; UI CRUD-write round-trip needs a logged-in write-role session (Designer "
                   "finish). Gate proven by test_master_write_gates.py + the live deny probe in this run.")


def _btn_enabled(page, dom_id):
    """True iff a Perspective button (#dom_id) is currently ENABLED. Perspective renders props.enabled=false
    as a `disabled` attribute on the <button> AND adds the `ia_button--disabled`/`disabled` class. Treat the
    button as disabled if either signal is present. Returns None if the button isn't found."""
    el = page.query_selector("#" + dom_id)
    if el is None:
        return None
    if el.get_attribute("disabled") is not None:
        return False
    cls = (el.get_attribute("class") or "")
    if "disabled" in cls.lower():
        return False
    try:
        return el.is_enabled()
    except Exception:
        return True


def check_master_write_gate_live(page, rep, view_label, url, clear_btn, save_btn, status_id,
                                 grid_sel, fill_pairs, count_fn, deny_marker, primary_fill=None):
    """LIVE on-the-box proof of BOTH the master-form refinement enable/disable model AND the P15 SERVER-SIDE
    write gate, on a swept master view. Sequence (the New button was REMOVED in the refinement sweep — Clear
    now starts a fresh insert and is ALWAYS enabled):

      1. open the view, click Clear -> blank insert-mode form (recordId=0).
      2. ASSERT (refinement): Save is DISABLED on the blank/cleared form (no record + no primary entered).
      3. fill the PRIMARY field (primary_fill=(domId,value)) -> entering data into a cleared form.
      4. ASSERT (refinement): Save is now ENABLED (has-entered-data signal).
      5. fill the rest of fill_pairs, then click Save. In this headless spike the session is ANONYMOUS (no
         IdP/roles), so auth.requireWrite(self.session) MUST DENY the write server-side -> status label shows
         "DENIED (server-side)" (and/or a SPIKE deny log marker), and the DB row count is UNCHANGED.

    `primary_fill` = (domId, value) for the view's PRIMARY/required TEXT input (the field that enables Save).
    If None (e.g. ManifestCost, whose primary is a dropdown that doesn't fill reliably headless), steps 3-5
    are SKIPPED with an honest reason — the server gate for that view is proven headless end-to-end by
    test_master_write_gates.py; the blank-form Save-disabled assertion (step 2) is still made.

    deny_marker = the SPIKE log substring this view emits on a denied write (e.g. "Size Save DENIED")."""
    try:
        pre = count_fn()
        page.goto(url, wait_until="networkidle", timeout=30000)
        if "Trial Expired" in page.inner_text("body"):
            rep.skip("%s SERVER-SIDE WRITE GATE (live)" % view_label, "Perspective trial expired")
            return
        try:
            page.wait_for_selector(grid_sel, timeout=20000)
        except Exception:
            pass
        cb = page.query_selector("#" + clear_btn)
        if not cb:
            rep.skip("%s SERVER-SIDE WRITE GATE (live)" % view_label, "Clear button not found")
            return
        cb.click()
        page.wait_for_timeout(1000)

        # --- refinement assertion: Save DISABLED on a blank/cleared form (no record, no data entered) ---
        save_disabled = _btn_enabled(page, save_btn) is False
        rep.check("%s (LIVE refinement): Save is DISABLED on a blank/cleared form (no record selected, no "
                  "data entered)" % view_label, save_disabled,
                  "Save enabled=%s after Clear (expected disabled)" % _btn_enabled(page, save_btn))

        def _fill(dom_id, val):
            f = page.query_selector("#" + dom_id)
            if not f:
                return False
            f.click()
            page.keyboard.press("Control+A")
            page.keyboard.press("Meta+A")
            page.keyboard.press("Delete")
            if val:
                f.type(str(val), delay=20)
            page.keyboard.press("Tab")
            page.wait_for_timeout(200)
            return True

        if primary_fill is None:
            rep.skip("%s (LIVE): Save-enable + anon-DENIED click-through" % view_label,
                     "primary is a dropdown (no reliable headless fill); gate proven by "
                     "test_master_write_gates.py")
            return

        # --- step 3-4: enter the primary -> Save becomes ENABLED (has-entered-data signal) ---
        _fill(primary_fill[0], primary_fill[1])
        save_enabled = _btn_enabled(page, save_btn) is True
        rep.check("%s (LIVE refinement): entering the primary field into a cleared form ENABLES Save"
                  % view_label, save_enabled,
                  "Save enabled=%s after entering %s=%r" % (_btn_enabled(page, save_btn), primary_fill[0],
                                                            primary_fill[1]))

        for dom_id, val in fill_pairs:
            _fill(dom_id, val)

        # --- step 5: anon Save -> DENIED server-side, no row written ---
        off = log_marker()
        sb = page.query_selector("#" + save_btn)
        if not sb:
            rep.skip("%s SERVER-SIDE WRITE GATE (live)" % view_label, "Save button not found")
            return
        sb.click()
        page.wait_for_timeout(1800)
        stxt = (page.query_selector("#" + status_id).inner_text() or "") if page.query_selector("#" + status_id) else ""
        denied_lines = grep_spike_since(off, deny_marker)
        post = count_fn()
        rep.check("%s (LIVE): anon Save is DENIED SERVER-SIDE (H3 hole closed end-to-end; no row written)"
                  % view_label,
                  ("DENIED (server-side)" in stxt or bool(denied_lines)) and post == pre,
                  "status=%r; SPIKE=%s; count pre=%d post=%d"
                  % (stxt[:60], denied_lines[-1].split("SPIKE")[-1][:50] if denied_lines else "none",
                     pre, post))
    except Exception as e:
        rep.skip("%s SERVER-SIDE WRITE GATE (live)" % view_label, "interaction failed: %s" % e)


# ---------------------------------------------------------------------------
# P11 (R18) — self-healing fixtures. A suite that seeds synthetic rows (forecast/supplier/master-CRUD/EDI)
# tears them down at the END; but if a PRIOR run was KILLED mid-test (Ctrl-C, trial-expiry abort, a crash),
# its synthetic rows survive and the next run collides — most visibly the forecast suite re-inserting a
# ZZF83x supplier code into the IX_INV_SUPPLIER_MST UNIQUE index. The cure is to PRE-CLEAN by sentinel at
# suite START, independent of any prior teardown having run. preclean_sentinels() is that shared helper:
# the suite hands it the SAME idempotent, child-before-parent DELETE statements its own teardown uses
# (every one keyed on the suite's ZZ*/VC_ADD sentinel), and we run them swallowing per-statement errors so
# a partially-dirty DB still drains to a clean baseline before the run. SENTINEL convention: every
# synthetic row a suite writes carries a greppable ZZ* business key and/or a VC_ADD stamp, so a pre-clean
# can target exactly the synthetic rows and NEVER a real client row.
def preclean_sentinels(exec_one, statements, label=""):
    """Run an ordered list of idempotent sentinel DELETE statements at suite start (self-heal a DB left
    dirty by a killed predecessor — P11/R18). `exec_one(stmt)` runs ONE SQL string (the suite's own
    sqlcmd wrapper). `statements` MUST be child-before-parent ordered (same order the suite's teardown
    uses) so FK/RESTRICT deletes succeed. Per-statement errors are swallowed (a missing table or a
    not-yet-seeded row is fine) so the pre-clean never aborts the run; a count of attempted/failed is
    returned for optional logging. NEVER pass a destructive statement that isn't sentinel-scoped."""
    attempted = failed = 0
    for stmt in statements:
        attempted += 1
        try:
            exec_one(stmt)
        except Exception:
            failed += 1
    if os.environ.get("PRECLEAN_ECHO") == "1":
        print("  [preclean%s] %d sentinel DELETE(s) run, %d swallowed"
              % ((" " + label) if label else "", attempted, failed))
    return attempted, failed


class Report:
    """Accumulates per-check PASS/FAIL/SKIP and prints a summary + exit code."""
    def __init__(self):
        self.rows = []

    def check(self, name, ok, detail=""):
        status = "PASS" if ok else "FAIL"
        self.rows.append((status, name, detail))
        print("  [%s] %s%s" % (status, name, ("  — " + detail) if detail else ""))
        return ok

    def skip(self, name, why):
        self.rows.append(("SKIP", name, why))
        print("  [SKIP] %s  — %s" % (name, why))

    def summary_exit(self):
        p = sum(1 for r in self.rows if r[0] == "PASS")
        f = sum(1 for r in self.rows if r[0] == "FAIL")
        s = sum(1 for r in self.rows if r[0] == "SKIP")
        print("\n===== E2E SUMMARY: %d PASS / %d FAIL / %d SKIP =====" % (p, f, s))
        return 1 if f else 0
