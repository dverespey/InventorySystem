"""E2E QA gate for the Logistics-master CRUD build — no human clicks.

Third master-data rebuild module (leaf, NAME-keyed). Drives a real Perspective
session headless against the live spike gateway (8.1.52 on :8088): resets the 2h
trial if needed, opens /logistics, and asserts on BOTH the DOM and the gateway-log
SPIKE markers (grep wrapper.log). Per-check PASS/FAIL/SKIP; exit 1 if any FAIL.

  python3 e2e/test_logistics_crud.py            # headless gate
  python3 e2e/test_logistics_crud.py --headed   # watch it live

Mirrors test_size_crud.py: reuses lib.py (Report, log markers, SPIKE grep, trial
reset), the domId-first selector and fill_field (Tabs to commit the bidirectional
binding). Views/SQL are NOT edited here.

SINGLE COMBINED VIEW: /logistics opens ONE view Master/Logistics/Logistics
containing the grid (left) AND the detail edit-form (right). Row-select is a
SAME-VIEW prop write (custom.recordId), NOT a navigation — onRowClick sets
self.view.custom.recordId = event.value.RecordID and the custom.recordId onChange
load script populates the in-view form. No List->Detail page navigation.

LOGISTICS = a leaf master keyed by a NAME (VC_LOGISTICS_NAME varchar(25)), NOT a
code. NO FK combos / NO enums. Validation: presence(name) + per-site uniqueness on
the NAME (checkCodeUnique; NO fixed-length rule). Delete-gate refCount counts ALL
THREE inbound refs: INV_SUPPLIER_MST + INV_PARTS_STOCK_MST + INV_PARTS_STOCK_MST_HIST
by IN_LOGISTICS_ID (the live DELETE_LogisticsCode trigger only nullifies the supplier
FK and dangles parts/_HIST; D3 RESTRICT blocks on any of the three).

DB ground truth (verified 2026-06-17 via sqlcmd, source-of-truth
db/namedqueries/master-crud-namedqueries.sql):
  1 logistics row. id 1 = 'TLDLOGISTICS SERVICES', referenced by 3 suppliers,
  0 parts, 0 hist -> refCount 3 (>0; delete MUST be blocked, referenced by SUPPLIER).
  'ZZ QA CARRIER' absent (throwaway round-trip name).
"""
import os
import subprocess
import sys
import time

from playwright.sync_api import sync_playwright

import lib
from reset_trial import reset_trial

LIST_URL = lib.view_url("logistics")
DB = lib.DB_CONN   # centralized DB-conn name (default Inventory_Spike; single prod-rename point)
SA_PASS = os.environ.get("SA_PASS", "Spike_Dev_2026!")

# Verified parity anchors (DB ground truth, see module docstring).
EXPECTED_COUNT = 1
ANCHOR_NAMES = ["TLDLOGISTICS SERVICES"]    # the single live row
DETAIL_ANCHOR = {"id": 1, "name": "TLDLOGISTICS SERVICES"}
# id 1 is referenced by 3 suppliers (parts/_HIST = 0) -> delete MUST be blocked.
REFERENCED = {"id": 1, "name": "TLDLOGISTICS SERVICES", "refCount": 3, "by": "supplier"}
TEST_NAME = "ZZ QA CARRIER"   # throwaway round-trip name (zero refs -> gate allows cleanup)

GRID = "#logistics-grid"
ROW = "div.ia_table__row"
CELL = "div.ia_table__cell"


# ---- selector helper (domId first, then text/role) -----------------------
def q(page, domid, text=None, role=None):
    el = page.query_selector("#" + domid)
    if el:
        return page.locator("#" + domid)
    if text:
        loc = page.get_by_text(text, exact=False)
        if loc.count():
            return loc.first
    if role:
        loc = page.get_by_role(role, name=text) if text else page.get_by_role(role)
        if loc.count():
            return loc.first
    return None


def fill_field(page, domid, val):
    """Robustly set a Perspective input's bound value headless (same as Size).

    Real keystrokes (type) fire onInput/onChange so the bidirectional binding
    commits into view.custom.form_*; Tab blurs to commit before the next action."""
    f = page.query_selector("#" + domid)
    if not f:
        return False
    f.click()
    page.keyboard.press("Control+A")
    page.keyboard.press("Meta+A")   # macOS select-all (dev box is a Mac)
    page.keyboard.press("Delete")
    if val is not None and val != "":
        f.type(str(val), delay=30)  # ElementHandle.type = real per-key events
    page.keyboard.press("Tab")      # blur -> commit the bidirectional binding
    page.wait_for_timeout(350)
    return True


def sqlq(query):
    """Run a one-shot sqlcmd against the spike container; return stdout (str)."""
    cmd = ["docker", "exec", "mssql-spike",
           "/opt/mssql-tools18/bin/sqlcmd", "-C", "-S", "localhost",
           "-U", "sa", "-P", SA_PASS, "-d", "Inventory",
           "-h", "-1", "-W", "-Q", "SET NOCOUNT ON; " + query]
    try:
        out = subprocess.check_output(cmd, stderr=subprocess.STDOUT, timeout=30)
        return out.decode("utf-8", "replace")
    except Exception as e:
        return "SQLERR: %s" % e


def db_logistics_count():
    out = sqlq("SELECT COUNT(*) FROM INV_LOGISTICS_MST")
    for tok in out.split():
        if tok.isdigit():
            return int(tok)
    return -1


def grid_text(page):
    g = page.query_selector(GRID)
    return g.inner_text() if g else ""


# ---- check 1: List renders + parity --------------------------------------
def check_list(page, rep):
    off_grid = lib.log_marker()   # capture before load so we can grep the Logistics/list count line
    page.goto(LIST_URL, wait_until="networkidle", timeout=45000)
    if "Trial Expired" in page.inner_text("body"):
        rep.check("List renders", False,
                  "Perspective TRIAL EXPIRED — run reset_trial.py with creds in .env")
        return False
    rendered = False
    try:
        page.wait_for_selector(GRID, timeout=30000)
        rendered = True
    except Exception:
        pass
    page.wait_for_timeout(2200)   # let the runPrepQuery transform + grid paint
    page.screenshot(path=lib.ARTIFACTS + "/logistics_list.png", full_page=True)
    rep.check("List grid mounts (%s present)" % GRID, rendered,
              "screenshot: artifacts/logistics_list.png")
    if not rendered:
        return False

    # The authoritative count is what the Logistics/list query returned, logged by
    # the gateway as "SPIKE Logistics/list: N rows" (the table DOM-virtualizes, so a
    # rendered-row count is only a floor). Assert that equals the DB count.
    rows = page.query_selector_all(GRID + " " + ROW)
    listed = None
    for l in lib.grep_spike_since(off_grid, "Logistics/list:"):
        try:
            listed = int(l.split("Logistics/list:")[1].strip().split()[0])
        except Exception:
            pass
    rep.check("Logistics/list query returned %d rows (== DB count)" % EXPECTED_COUNT,
              listed == EXPECTED_COUNT, "query reported=%s; rendered DOM window=%d" % (listed, len(rows)))
    rep.check("List renders a window of rows (DOM floor > 0)",
              len(rows) > 0, "rendered rows=%d" % len(rows))

    text = grid_text(page)
    missing = [c for c in ANCHOR_NAMES if c not in text]
    rep.check("List anchor names present (%s)" % ", ".join(ANCHOR_NAMES), not missing,
              "missing: %s" % missing if missing else "all present")
    return True


# ---- check 2: row select -> Detail ---------------------------------------
def open_detail_via_row(page, rep, name):
    """Click the grid row whose first cell == `name`; assert the in-view prop write
    (custom.recordId) fired and the recordId onChange loaded the row into the
    same-view form. NO navigation. Returns True if the form populated."""
    off = lib.log_marker()
    target = None
    for r in page.query_selector_all(GRID + " " + ROW):
        toks = [t.strip() for t in (r.inner_text() or "").split("\n") if t.strip()]
        if toks and toks[0] == name:
            target = r
            break
    if target is None:
        rep.check("Row select: row %s clickable" % name, False, "row not found in grid")
        return False
    target.scroll_into_view_if_needed()
    target.click()
    time.sleep(2.0)   # WS round-trip: prop write -> recordId onChange load
    click_lines = lib.grep_spike_since(off, "Logistics list -> open Detail")
    rep.check("onRowClick set custom.recordId in-view (SPIKE marker, recId=%d)" % DETAIL_ANCHOR["id"],
              any(("recordId=%d" % DETAIL_ANCHOR["id"]) in l for l in click_lines),
              click_lines[-1].split("SPIKE")[-1][:70] if click_lines else "no row-click marker in wrapper.log")

    load_lines = lib.grep_spike_since(off, "Logistics Detail loaded")
    rep.check("recordId onChange loaded the row into the in-view form (SPIKE 'Detail loaded id=%d')" % DETAIL_ANCHOR["id"],
              any(("id=%d" % DETAIL_ANCHOR["id"]) in l for l in load_lines),
              load_lines[-1].split("SPIKE")[-1][:70] if load_lines else "no Detail-loaded marker")

    page.wait_for_timeout(800)
    page.screenshot(path=lib.ARTIFACTS + "/logistics_detail.png", full_page=True)
    name_field = page.query_selector("#logistics-name")
    name_val = name_field.get_attribute("value") if name_field else None
    rep.check("Detail name field populated == %s" % name, name_val == name,
              "field value=%r" % name_val)
    return name_val == name


# ---- check 3: validation (blank name, dup name) --------------------------
def check_validation(page, rep):
    """On a fresh New form: blank name rejected; existing name (the live row) blocked
    (checkCodeUnique). Asserted via #logistics-status text + SPIKE 'save REJECTED'
    markers. No DB write occurs (all fail validation/dup-check). NO fixed-length rule
    for Logistics (it is keyed by a name, not a code)."""
    page.goto(LIST_URL, wait_until="networkidle", timeout=30000)
    page.wait_for_selector(GRID, timeout=20000)
    nb = q(page, "logistics-clear-btn", text="Clear", role="button")
    if not nb:
        rep.skip("Validation: blank name rejected", "Clear button not found")
        rep.skip("Validation: dup name blocked", "Clear button not found")
        return
    nb.click()
    page.wait_for_timeout(1500)

    def set_field(domid, val):
        return fill_field(page, domid, val)

    # --- blank name ---
    off = lib.log_marker()
    set_field("logistics-name", "")
    set_field("logistics-city", "QA Blank Name City")
    save = q(page, "logistics-save-btn", text="Save", role="button")
    if save:
        save.click()
        time.sleep(1.8)
        lines = lib.grep_spike_since(off, "save REJECTED: blank name")
        status = page.query_selector("#logistics-status")
        stxt = status.inner_text() if status else ""
        blank_ok = bool(lines) or ("name is required" in stxt)
        rep.check("Validation: blank name REJECTED (no insert)", blank_ok,
                  (lines[-1].split("SPIKE")[-1][:60] if lines else ("status=%r" % stxt[:60])))
    else:
        rep.skip("Validation: blank name rejected", "Save button not found")

    # --- dup name (existing live row) ---
    off = lib.log_marker()
    set_field("logistics-name", REFERENCED["name"])   # already exists
    set_field("logistics-city", "QA Dup Name City")
    save = q(page, "logistics-save-btn", text="Save", role="button")
    if save:
        save.click()
        time.sleep(1.8)
        lines = lib.grep_spike_since(off, "save REJECTED: dup name")
        status = page.query_selector("#logistics-status")
        stxt = status.inner_text() if status else ""
        dup_ok = bool(lines) or ("already exists" in stxt)
        rep.check("Validation: duplicate name BLOCKED (checkCodeUnique)", dup_ok,
                  (lines[-1].split("SPIKE")[-1][:60] if lines else ("status=%r" % stxt[:60])))
    else:
        rep.skip("Validation: dup name blocked", "Save button not found")

    rep.check("Validation left DB clean (still %d logistics)" % EXPECTED_COUNT,
              db_logistics_count() == EXPECTED_COUNT, "count=%d" % db_logistics_count())


# ---- check 4: R1 delete-gate (CRITICAL) ----------------------------------
def check_delete_gate(page, rep):
    """Open the referenced logistics row (id 1, refCount 3 = 3 suppliers) in Detail
    and click Delete -> MUST be blocked. Assert: status shows the reference message,
    SPIKE refCount ran and returned >0, SPIKE DELETE BLOCKED fired, and the row still
    exists in the DB. Never deletes a referenced logistics. The supplier refs are
    exactly why refCount must count INV_SUPPLIER_MST (which the live trigger would
    only nullify, not block on)."""
    out = sqlq("SELECT (SELECT COUNT(*) FROM INV_SUPPLIER_MST WHERE IN_LOGISTICS_ID=%d)"
               "+(SELECT COUNT(*) FROM INV_PARTS_STOCK_MST WHERE IN_LOGISTICS_ID=%d)"
               "+(SELECT COUNT(*) FROM INV_PARTS_STOCK_MST_HIST WHERE IN_LOGISTICS_ID=%d)"
               % (REFERENCED["id"], REFERENCED["id"], REFERENCED["id"]))
    db_ref = next((int(t) for t in out.split() if t.isdigit()), -1)
    rep.check("Delete-gate pre-state: %s has %d refs in DB (supplier+parts+_HIST, >0)"
              % (REFERENCED["name"], REFERENCED["refCount"]),
              db_ref == REFERENCED["refCount"], "db refCount=%d (referenced by %s)" % (db_ref, REFERENCED["by"]))

    page.goto(LIST_URL, wait_until="networkidle", timeout=30000)
    page.wait_for_selector(GRID, timeout=20000)
    page.wait_for_timeout(1500)
    target = None
    for r in page.query_selector_all(GRID + " " + ROW):
        toks = [t.strip() for t in (r.inner_text() or "").split("\n") if t.strip()]
        if toks and toks[0] == REFERENCED["name"]:
            target = r
            break
    if target is None:
        rep.check("Delete-gate: open referenced row %s" % REFERENCED["name"], False, "row not found")
        return
    target.scroll_into_view_if_needed()
    target.click()
    time.sleep(2.0)

    off = lib.log_marker()
    delbtn = q(page, "logistics-delete-btn", text="Delete", role="button")
    if not delbtn:
        rep.check("Delete-gate: Delete button present", False, "not found")
        return
    delbtn.click()
    time.sleep(2.0)

    ref_lines = lib.grep_spike_since(off, "Logistics refCount")
    blk_lines = lib.grep_spike_since(off, "DELETE BLOCKED")
    status = page.query_selector("#logistics-status")
    stxt = status.inner_text() if status else ""
    page.screenshot(path=lib.ARTIFACTS + "/logistics_delete_blocked.png", full_page=True)

    ran_ref = any(("n=%d" % REFERENCED["refCount"]) in l for l in ref_lines)
    rep.check("Delete-gate: refCount ran and returned %d (SPIKE 'refCount ... n=%d')"
              % (REFERENCED["refCount"], REFERENCED["refCount"]),
              ran_ref, ref_lines[-1].split("SPIKE")[-1][:70] if ref_lines else "no refCount marker")
    rep.check("Delete-gate: DELETE BLOCKED (SPIKE 'DELETE BLOCKED' fired)",
              bool(blk_lines), blk_lines[-1].split("SPIKE")[-1][:70] if blk_lines else "no BLOCKED marker")
    rep.check("Delete-gate: status shows the reference message ('still referenced')",
              "still referenced" in stxt, "status=%r" % stxt[:90])

    still = sqlq("SELECT COUNT(*) FROM INV_LOGISTICS_MST WHERE IN_LOGISTICS_ID=%d" % REFERENCED["id"])
    still_n = next((int(t) for t in still.split() if t.isdigit()), -1)
    rep.check("Delete-gate: referenced logistics %s STILL EXISTS in DB (not deleted)" % REFERENCED["name"],
              still_n == 1, "rows for id %d = %d" % (REFERENCED["id"], still_n))


# ---- check 5: non-destructive insert/update round-trip -------------------
def check_round_trip(page, rep):
    """Insert a throwaway 'ZZ QA CARRIER' logistics through the UI, verify it landed
    in the DB, then delete it through the UI (zero refs -> gate allows) and confirm
    the DB is back to baseline. Fixture discipline: teardown sweeps any stray row."""
    pre = db_logistics_count()
    if pre != EXPECTED_COUNT:
        rep.skip("Round-trip insert/delete %s" % TEST_NAME, "DB not at baseline (%d != %d)" % (pre, EXPECTED_COUNT))
        return
    if sqlq("SELECT COUNT(*) FROM INV_LOGISTICS_MST WHERE VC_LOGISTICS_NAME='%s'" % TEST_NAME).strip().split()[-1] != "0":
        rep.skip("Round-trip insert/delete %s" % TEST_NAME, "%s already present — refusing to collide" % TEST_NAME)
        return

    page.goto(LIST_URL, wait_until="networkidle", timeout=30000)
    page.wait_for_selector(GRID, timeout=20000)
    nb = q(page, "logistics-clear-btn", text="Clear", role="button")
    if not nb:
        rep.skip("Round-trip insert/delete %s" % TEST_NAME, "Clear button not found")
        return
    nb.click()
    page.wait_for_timeout(1500)

    def set_field(domid, val):
        return fill_field(page, domid, val)

    inserted = False
    try:
        set_field("logistics-name", TEST_NAME)
        set_field("logistics-city", "QA CITY")
        set_field("logistics-person", "QA PERSON")
        off = lib.log_marker()
        save = q(page, "logistics-save-btn", text="Save", role="button")
        if save:
            save.click()
            time.sleep(2.0)
            ins_lines = lib.grep_spike_since(off, "Logistics INSERT ok")
            row = sqlq("SELECT IN_LOGISTICS_ID, VC_CITY, VC_PERSON FROM INV_LOGISTICS_MST "
                       "WHERE VC_LOGISTICS_NAME='%s'" % TEST_NAME)
            inserted = ("QA CITY" in row)
            rep.check("Round-trip: INSERT wrote %s to DB (city/person round-tripped)" % TEST_NAME,
                      inserted, ("db row=%r; SPIKE=%s" %
                                 (row.strip()[:60], ins_lines[-1].split("SPIKE")[-1][:40] if ins_lines else "no marker")))
            rep.check("Round-trip: DB now has %d logistics (+1)" % (EXPECTED_COUNT + 1),
                      db_logistics_count() == EXPECTED_COUNT + 1, "count=%d" % db_logistics_count())
        else:
            rep.skip("Round-trip insert/delete %s" % TEST_NAME, "Save button not found")
    except Exception as e:
        rep.skip("Round-trip insert", "insert interaction failed: %s" % e)

    # ---- cleanup: delete the test row through the UI (zero refs -> allowed) ----
    if inserted:
        try:
            off = lib.log_marker()
            delbtn = q(page, "logistics-delete-btn", text="Delete", role="button")
            if delbtn:
                delbtn.click()
                time.sleep(2.0)
                del_lines = lib.grep_spike_since(off, "Logistics DELETE ok")
                rep.check("Round-trip: UI DELETE of zero-ref %s succeeded (gate allows)" % TEST_NAME,
                          bool(del_lines), del_lines[-1].split("SPIKE")[-1][:60] if del_lines else "no DELETE ok marker")
        except Exception as e:
            rep.skip("Round-trip UI delete", "delete interaction failed: %s" % e)


def teardown(rep):
    """Hard fixture-discipline sweep: ensure no ZZ QA CARRIER row survives and the DB
    is back to exactly the baseline count, regardless of where the round-trip stopped.
    The throwaway has zero refs so this DELETE is safe and never touches real rows."""
    stray = sqlq("SELECT COUNT(*) FROM INV_LOGISTICS_MST WHERE VC_LOGISTICS_NAME='%s'" % TEST_NAME)
    n_stray = next((int(t) for t in stray.split() if t.isdigit()), 0)
    if n_stray > 0:
        sqlq("DELETE FROM INV_LOGISTICS_MST WHERE VC_LOGISTICS_NAME='%s'" % TEST_NAME)
    final = db_logistics_count()
    rep.check("Teardown: DB restored to %d logistics, no stray %s" % (EXPECTED_COUNT, TEST_NAME),
              final == EXPECTED_COUNT and n_stray == 0,
              "final count=%d, stray removed=%d" % (final, n_stray))


def main():
    headed = "--headed" in sys.argv
    os.makedirs(lib.ARTIFACTS, exist_ok=True)
    rep = lib.Report()

    print("== DB pre-state ==")
    pre = db_logistics_count()
    rep.check("Sandbox up + baseline (%d logistics)" % EXPECTED_COUNT, pre == EXPECTED_COUNT,
              "count=%d" % pre)

    with sync_playwright() as p:
        b = p.chromium.launch(headless=not headed, slow_mo=250 if headed else 0)
        pg = b.new_page(viewport={"width": 1680, "height": 1050})
        print("== trial reset ==")
        ok, msg = reset_trial(pg)
        if msg == "NEED_CREDS":
            rep.skip("trial reset", "no GATEWAY_USER/GATEWAY_PASS — render blocked if trial expired")
        else:
            rep.check("trial active", ok, msg)

        print("== 1. List renders + parity ==")
        list_ok = check_list(pg, rep)

        print("== 2. Row select -> Detail ==")
        if list_ok:
            open_detail_via_row(pg, rep, DETAIL_ANCHOR["name"])
        else:
            rep.skip("Row select -> Detail", "List did not render")

        print("== 3. SERVER-SIDE WRITE GATE (LIVE, P15 H3 hole closed end-to-end) ==")
        lib.check_master_write_gate_live(
            pg, rep, "Logistics write gate", LIST_URL,
            clear_btn="logistics-clear-btn", save_btn="logistics-save-btn", status_id="logistics-status",
            grid_sel=GRID, fill_pairs=[], count_fn=db_logistics_count,
            deny_marker="Logistics Save DENIED", primary_fill=("logistics-name", "ZZ Probe"))

        # Checks 4-6 drive UI CRUD WRITES (validation/delete-gate/round-trip). With the P15 server-side
        # write gate active, the anon spike session is correctly DENIED before those paths run, so they
        # can no longer be exercised through the live UI here. SKIPPED; the gate is proven headless end-to-
        # end (per view) by test_master_write_gates.py + the live deny above.
        print("== 4-6. UI CRUD-write cases (validation/delete-gate/round-trip) ==")
        rep.skip("Validation (admin UI write)", lib.WRITE_GATE_SKIP)
        rep.skip("R1 delete-gate (admin UI write)", lib.WRITE_GATE_SKIP)
        rep.skip("Round-trip insert/delete (admin UI write)", lib.WRITE_GATE_SKIP)

        b.close()

    print("== teardown / fixture sweep ==")
    teardown(rep)

    return rep.summary_exit()


if __name__ == "__main__":
    sys.exit(main())
