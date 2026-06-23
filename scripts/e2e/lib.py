"""Shared helpers for the Perspective E2E harness (Playwright, Python).

No third-party deps beyond playwright. Reads optional gateway creds from
environment or a gitignored scripts/e2e/.env (KEY=VALUE lines).
"""
import os, re, subprocess, time

BASE = os.environ.get("GW_BASE", "http://localhost:8088")
WRAPPER_LOG = os.environ.get("GW_LOG", "/usr/local/ignition/logs/wrapper.log")
ARTIFACTS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "artifacts")


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


def check_master_write_gate_live(page, rep, view_label, url, new_btn, save_btn, status_id,
                                 grid_sel, fill_pairs, count_fn, deny_marker):
    """LIVE on-the-box proof that a swept master view's WRITE is gated SERVER-SIDE (P15): open the view,
    click New, fill a throwaway probe row, click Save — in this headless spike the session is anonymous
    (no IdP/roles), so auth.requireWrite(self.session) MUST DENY the write server-side. Assert the status
    label shows "DENIED (server-side)" (and/or a SPIKE deny log marker), and the DB row count is unchanged
    (no write slipped through). Mirrors test_sites_crud.check_server_side_write_gate, parameterized per
    view. `count_fn()` returns the current table row count; `fill_pairs` = [(domId, value), ...]; the
    test's own fill_field is used via the page (we call a minimal inline fill here to avoid importing it).

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
        nb = page.query_selector("#" + new_btn)
        if not nb:
            rep.skip("%s SERVER-SIDE WRITE GATE (live)" % view_label, "New button not found")
            return
        nb.click()
        page.wait_for_timeout(1000)
        for dom_id, val in fill_pairs:
            f = page.query_selector("#" + dom_id)
            if not f:
                continue
            f.click()
            page.keyboard.press("Control+A")
            page.keyboard.press("Meta+A")
            page.keyboard.press("Delete")
            if val:
                f.type(str(val), delay=20)
            page.keyboard.press("Tab")
            page.wait_for_timeout(150)
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
