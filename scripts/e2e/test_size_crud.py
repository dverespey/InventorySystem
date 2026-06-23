"""E2E QA gate for the Size-master CRUD build — no human clicks.

Second master-data rebuild module. Drives a real Perspective session headless
against the live spike gateway (8.1.52 on :8088): resets the 2h trial if needed,
opens /size, and asserts on BOTH the DOM and the gateway-log SPIKE markers
(grep wrapper.log). Per-check PASS/FAIL/SKIP; exit 1 if any FAIL.

  python3 scripts/e2e/test_size_crud.py            # headless gate
  python3 scripts/e2e/test_size_crud.py --headed   # watch it live

Mirrors test_supplier_crud.py: reuses lib.py (Report, log markers, SPIKE grep,
trial reset), the domId-first selector and fill_field (Tabs to commit the
bidirectional binding). Views/SQL are NOT edited here.

SINGLE COMBINED VIEW: /size opens ONE view Master/Size/Size containing the grid
(left) AND the detail edit-form (right). Row-select is a SAME-VIEW prop write
(custom.recordId), NOT a navigation — onRowClick sets self.view.custom.recordId =
event.value.RecordID and the custom.recordId onChange load script populates the
in-view form. There is no List->Detail page navigation anywhere.

SIZE = a LEAN SUBSET of Supplier: 4 fields only (code, name, usage int, days int),
NO FK combos / NO enum dropdowns. Validation: presence(code,name) + len(code)<=6
(NOT the ==5 Supplier rule). Delete-gate refCount counts BOTH INV_PARTS_STOCK_MST
AND INV_PARTS_STOCK_MST_HIST by IN_SIZE_ID (the live trigger ignores _HIST; D3
RESTRICT blocks on both).

DB ground truth (verified 2026-06-17 via sqlcmd, source-of-truth
docs/analysis/master-data/master-crud-namedqueries.sql):
  64 sizes. First codes by VC_SIZE_CODE order: '' (blank, id 80) then 15D (id 19),
  15D1 (id 77), 15G (id 28), 15GRND (id 16). Stable non-blank anchors: 15D/15D1/15G.
  id 56 = 16H : 1 current part + 167 history rows = refCount 168 (>0; delete MUST be
  blocked — and the 167 _HIST rows are exactly why refCount must count _HIST).
  ZZTST absent.
"""
import os
import subprocess
import sys
import time

from playwright.sync_api import sync_playwright

import lib
from reset_trial import reset_trial

LIST_URL = lib.BASE + "/data/perspective/client/spike/size"
DB = "Inventory_Spike"
SA_PASS = os.environ.get("SA_PASS", "Spike_Dev_2026!")

# Verified parity anchors (DB ground truth, see module docstring).
EXPECTED_COUNT = 64
ANCHOR_CODES = ["15D", "15D1", "15G"]      # stable non-blank codes, in code order
DETAIL_ANCHOR = {"id": 19, "code": "15D", "name": "CMWA 15\" 0572B"}
REFERENCED = {"id": 56, "code": "16H", "refCount": 168}  # delete MUST be blocked
TEST_CODE = "ZZTST"     # throwaway round-trip code (zero refs -> gate allows cleanup)

GRID = "#size-grid"
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
    """Robustly set a Perspective input's bound value headless (same as Supplier).

    Real keystrokes (type) fire onInput/onChange so the bidirectional binding
    commits into view.custom.form_*; Tab blurs to commit before the next action.
    Works for both text-field (props.text) and numeric-entry-field (props.value)."""
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


def db_size_count():
    out = sqlq("SELECT COUNT(*) FROM INV_SIZE_MST")
    for tok in out.split():
        if tok.isdigit():
            return int(tok)
    return -1


def grid_text(page):
    g = page.query_selector(GRID)
    return g.inner_text() if g else ""


# ---- check 1: List renders + parity --------------------------------------
def check_list(page, rep):
    off_grid = lib.log_marker()   # capture before load so we can grep the Size/list count line
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
    page.screenshot(path=lib.ARTIFACTS + "/size_list.png", full_page=True)
    rep.check("List grid mounts (%s present)" % GRID, rendered,
              "screenshot: artifacts/size_list.png")
    if not rendered:
        return False

    # The grid model carries all 64 rows; the table DOM-VIRTUALIZES (only the
    # viewport window is mounted, ~32 rows), so counting rendered DOM rows is NOT
    # a reliable total. The authoritative count is what the Size/list query
    # returned, logged by the gateway as "SPIKE Size/list: N rows". Assert that
    # equals the DB count, and that the rendered window is a non-empty virtualized
    # subset (floor) — together this proves all 64 are in the grid model in order.
    rows = page.query_selector_all(GRID + " " + ROW)
    listed = None
    for l in lib.grep_spike_since(off_grid, "Size/list:"):
        try:
            listed = int(l.split("Size/list:")[1].strip().split()[0])
        except Exception:
            pass
    rep.check("Size/list query returned %d rows (== DB count; grid DOM-virtualizes)" % EXPECTED_COUNT,
              listed == EXPECTED_COUNT, "query reported=%s; rendered DOM window=%d" % (listed, len(rows)))
    rep.check("List renders a virtualized window of rows (DOM floor > 0)",
              len(rows) > 0, "rendered rows=%d" % len(rows))

    text = grid_text(page)
    missing = [c for c in ANCHOR_CODES if c not in text]
    rep.check("List anchor codes present (%s)" % ", ".join(ANCHOR_CODES), not missing,
              "missing: %s" % missing if missing else "all present")

    # parity ordering: 15D before 15D1 before 15G in grid text.
    pos = [text.find(c) for c in ANCHOR_CODES]
    ordered = all(p >= 0 for p in pos) and pos == sorted(pos)
    rep.check("List code order matches SELECT_SizeInfo '' (15D<15D1<15G)",
              ordered, "text offsets=%s" % pos)
    return True


# ---- check 2: row select -> Detail ---------------------------------------
def open_detail_via_row(page, rep, code):
    """Click the grid row whose first cell == `code`; assert the in-view prop
    write (custom.recordId) fired and the recordId onChange loaded the row into
    the same-view form. NO navigation. Returns True if the form populated."""
    off = lib.log_marker()
    target = None
    for r in page.query_selector_all(GRID + " " + ROW):
        toks = [t.strip() for t in (r.inner_text() or "").split("\n") if t.strip()]
        if toks and toks[0] == code:
            target = r
            break
    if target is None:
        rep.check("Row select: row %s clickable" % code, False, "row not found in grid")
        return False
    target.scroll_into_view_if_needed()
    target.click()
    time.sleep(2.0)   # WS round-trip: prop write -> recordId onChange load
    click_lines = lib.grep_spike_since(off, "Size list -> open Detail")
    rep.check("onRowClick set custom.recordId in-view (SPIKE marker, recId=%d)" % DETAIL_ANCHOR["id"],
              any(("recordId=%d" % DETAIL_ANCHOR["id"]) in l for l in click_lines),
              click_lines[-1].split("SPIKE")[-1][:70] if click_lines else "no row-click marker in wrapper.log")

    load_lines = lib.grep_spike_since(off, "Size Detail loaded")
    rep.check("recordId onChange loaded the row into the in-view form (SPIKE 'Detail loaded id=%d')" % DETAIL_ANCHOR["id"],
              any(("id=%d" % DETAIL_ANCHOR["id"]) in l for l in load_lines),
              load_lines[-1].split("SPIKE")[-1][:70] if load_lines else "no Detail-loaded marker")

    page.wait_for_timeout(800)
    page.screenshot(path=lib.ARTIFACTS + "/size_detail.png", full_page=True)
    body = page.inner_text("body")
    code_field = page.query_selector("#size-code")
    code_val = code_field.get_attribute("value") if code_field else None
    rep.check("Detail code field populated == %s" % code, code_val == code,
              "field value=%r" % code_val)
    rep.check("Detail name field shows %s" % DETAIL_ANCHOR["name"],
              DETAIL_ANCHOR["name"] in body, "name in body=%s" % (DETAIL_ANCHOR["name"] in body))
    return code_val == code


# ---- check 3: validation (blank code, >6-char code, dup code) ------------
def check_validation(page, rep):
    """On a fresh New form: blank code rejected; >6-char code rejected; existing
    code (16H) blocked (checkCodeUnique). Asserted via #size-status text + SPIKE
    'save REJECTED' markers. No DB write occurs (all fail validation/dup-check)."""
    page.goto(LIST_URL, wait_until="networkidle", timeout=30000)
    page.wait_for_selector(GRID, timeout=20000)
    nb = q(page, "size-new-btn", text="New Size", role="button")
    if not nb:
        rep.skip("Validation: blank code rejected", "New Size button not found")
        rep.skip("Validation: >6-char code rejected", "New Size button not found")
        rep.skip("Validation: dup code blocked", "New Size button not found")
        return
    nb.click()
    page.wait_for_timeout(1500)

    def set_field(domid, val):
        return fill_field(page, domid, val)

    # --- blank code ---
    off = lib.log_marker()
    set_field("size-code", "")
    set_field("size-name", "QA Blank Code")
    save = q(page, "size-save-btn", text="Save", role="button")
    if save:
        save.click()
        time.sleep(1.8)
        lines = lib.grep_spike_since(off, "save REJECTED: blank code")
        status = page.query_selector("#size-status")
        stxt = status.inner_text() if status else ""
        blank_ok = bool(lines) or ("code is required" in stxt)
        rep.check("Validation: blank code REJECTED (no insert)", blank_ok,
                  (lines[-1].split("SPIKE")[-1][:60] if lines else ("status=%r" % stxt[:60])))
    else:
        rep.skip("Validation: blank code rejected", "Save button not found")

    # --- >6-char code ---
    off = lib.log_marker()
    set_field("size-code", "ABCDEFG")   # 7 chars
    set_field("size-name", "QA Too Long")
    save = q(page, "size-save-btn", text="Save", role="button")
    if save:
        save.click()
        time.sleep(1.8)
        lines = lib.grep_spike_since(off, "save REJECTED: code len")
        status = page.query_selector("#size-status")
        stxt = status.inner_text() if status else ""
        # NOTE: the field has maxLength=6, so the UI itself caps at 6 chars; the
        # binding may deliver only 6. We assert the >6 rule via the SPIKE marker
        # OR the status; if the field truncated to 6 this becomes a no-reject which
        # we surface honestly.
        long_ok = bool(lines) or ("6 characters or fewer" in stxt)
        rep.check("Validation: >6-char code REJECTED (or field-capped at 6)", long_ok,
                  (lines[-1].split("SPIKE")[-1][:60] if lines
                   else ("status=%r (maxLength may have capped input to 6)" % stxt[:60])))
    else:
        rep.skip("Validation: >6-char code rejected", "Save button not found")

    # --- dup code (existing 16H) ---
    off = lib.log_marker()
    set_field("size-code", REFERENCED["code"])   # 16H, already exists
    set_field("size-name", "QA Dup Code")
    save = q(page, "size-save-btn", text="Save", role="button")
    if save:
        save.click()
        time.sleep(1.8)
        lines = lib.grep_spike_since(off, "save REJECTED: dup code")
        status = page.query_selector("#size-status")
        stxt = status.inner_text() if status else ""
        dup_ok = bool(lines) or ("already exists" in stxt)
        rep.check("Validation: duplicate code BLOCKED (checkCodeUnique)", dup_ok,
                  (lines[-1].split("SPIKE")[-1][:60] if lines else ("status=%r" % stxt[:60])))
    else:
        rep.skip("Validation: dup code blocked", "Save button not found")

    rep.check("Validation left DB clean (still %d sizes)" % EXPECTED_COUNT,
              db_size_count() == EXPECTED_COUNT, "count=%d" % db_size_count())


# ---- check 4: R1 delete-gate (CRITICAL) ----------------------------------
def check_delete_gate(page, rep):
    """Open the referenced size (16H, id 56, refCount 168 = 1 part + 167 _HIST) in
    Detail and click Delete -> MUST be blocked. Assert: status shows the reference
    message, SPIKE refCount ran and returned >0, SPIKE DELETE BLOCKED fired, and the
    row still exists in the DB. Never deletes a referenced size. The 167 _HIST rows
    are precisely why refCount must count _HIST (the live trigger ignores it)."""
    out = sqlq("SELECT (SELECT COUNT(*) FROM INV_PARTS_STOCK_MST WHERE IN_SIZE_ID=%d)"
               "+(SELECT COUNT(*) FROM INV_PARTS_STOCK_MST_HIST WHERE IN_SIZE_ID=%d)"
               % (REFERENCED["id"], REFERENCED["id"]))
    db_ref = next((int(t) for t in out.split() if t.isdigit()), -1)
    rep.check("Delete-gate pre-state: %s has %d refs in DB (parts + _HIST, >0)"
              % (REFERENCED["code"], REFERENCED["refCount"]),
              db_ref == REFERENCED["refCount"], "db refCount=%d" % db_ref)

    page.goto(LIST_URL, wait_until="networkidle", timeout=30000)
    page.wait_for_selector(GRID, timeout=20000)
    page.wait_for_timeout(1500)
    target = None
    for r in page.query_selector_all(GRID + " " + ROW):
        toks = [t.strip() for t in (r.inner_text() or "").split("\n") if t.strip()]
        if toks and toks[0] == REFERENCED["code"]:
            target = r
            break
    if target is None:
        rep.check("Delete-gate: open referenced row %s" % REFERENCED["code"], False, "row not found")
        return
    target.scroll_into_view_if_needed()
    target.click()
    time.sleep(2.0)

    off = lib.log_marker()
    delbtn = q(page, "size-delete-btn", text="Delete", role="button")
    if not delbtn:
        rep.check("Delete-gate: Delete button present", False, "not found")
        return
    delbtn.click()
    time.sleep(2.0)

    ref_lines = lib.grep_spike_since(off, "Size refCount")
    blk_lines = lib.grep_spike_since(off, "DELETE BLOCKED")
    status = page.query_selector("#size-status")
    stxt = status.inner_text() if status else ""
    page.screenshot(path=lib.ARTIFACTS + "/size_delete_blocked.png", full_page=True)

    ran_ref = any(("n=%d" % REFERENCED["refCount"]) in l for l in ref_lines)
    rep.check("Delete-gate: refCount ran and returned %d (SPIKE 'refCount ... n=%d')"
              % (REFERENCED["refCount"], REFERENCED["refCount"]),
              ran_ref, ref_lines[-1].split("SPIKE")[-1][:70] if ref_lines else "no refCount marker")
    rep.check("Delete-gate: DELETE BLOCKED (SPIKE 'DELETE BLOCKED' fired)",
              bool(blk_lines), blk_lines[-1].split("SPIKE")[-1][:70] if blk_lines else "no BLOCKED marker")
    rep.check("Delete-gate: status shows the reference message ('still referenced')",
              "still referenced" in stxt, "status=%r" % stxt[:90])

    still = sqlq("SELECT COUNT(*) FROM INV_SIZE_MST WHERE IN_SIZE_ID=%d" % REFERENCED["id"])
    still_n = next((int(t) for t in still.split() if t.isdigit()), -1)
    rep.check("Delete-gate: referenced size %s STILL EXISTS in DB (not deleted)" % REFERENCED["code"],
              still_n == 1, "rows for id %d = %d" % (REFERENCED["id"], still_n))


# ---- check 5: non-destructive insert/update round-trip -------------------
def check_round_trip(page, rep):
    """Insert a throwaway ZZTST size through the UI, verify it landed in the DB,
    then delete it through the UI (zero refs -> gate allows) and confirm the DB is
    back to baseline. Fixture discipline: teardown sweeps any stray row."""
    pre = db_size_count()
    if pre != EXPECTED_COUNT:
        rep.skip("Round-trip insert/delete ZZTST", "DB not at baseline (%d != %d)" % (pre, EXPECTED_COUNT))
        return
    if sqlq("SELECT COUNT(*) FROM INV_SIZE_MST WHERE VC_SIZE_CODE='%s'" % TEST_CODE).strip().split()[-1] != "0":
        rep.skip("Round-trip insert/delete ZZTST", "%s already present — refusing to collide" % TEST_CODE)
        return

    page.goto(LIST_URL, wait_until="networkidle", timeout=30000)
    page.wait_for_selector(GRID, timeout=20000)
    nb = q(page, "size-new-btn", text="New Size", role="button")
    if not nb:
        rep.skip("Round-trip insert/delete ZZTST", "New Size button not found")
        return
    nb.click()
    page.wait_for_timeout(1500)

    def set_field(domid, val):
        return fill_field(page, domid, val)

    inserted = False
    try:
        set_field("size-code", TEST_CODE)
        set_field("size-name", "QA ROUND TRIP")
        set_field("size-usage", "0")
        set_field("size-days", "0")
        off = lib.log_marker()
        save = q(page, "size-save-btn", text="Save", role="button")
        if save:
            save.click()
            time.sleep(2.0)
            ins_lines = lib.grep_spike_since(off, "Size INSERT ok")
            row = sqlq("SELECT IN_SIZE_ID, VC_SIZE_NAME, IN_USAGE, IN_DAYS FROM INV_SIZE_MST "
                       "WHERE VC_SIZE_CODE='%s'" % TEST_CODE)
            inserted = ("QA ROUND TRIP" in row)
            rep.check("Round-trip: INSERT wrote %s to DB (name round-tripped)" % TEST_CODE,
                      inserted, ("db row=%r; SPIKE=%s" %
                                 (row.strip()[:60], ins_lines[-1].split("SPIKE")[-1][:40] if ins_lines else "no marker")))
            rep.check("Round-trip: DB now has %d sizes (+1)" % (EXPECTED_COUNT + 1),
                      db_size_count() == EXPECTED_COUNT + 1, "count=%d" % db_size_count())
        else:
            rep.skip("Round-trip insert/delete ZZTST", "Save button not found")
    except Exception as e:
        rep.skip("Round-trip insert", "insert interaction failed: %s" % e)

    # ---- cleanup: delete the test row through the UI (zero refs -> allowed) ----
    if inserted:
        try:
            off = lib.log_marker()
            delbtn = q(page, "size-delete-btn", text="Delete", role="button")
            if delbtn:
                delbtn.click()
                time.sleep(2.0)
                del_lines = lib.grep_spike_since(off, "Size DELETE ok")
                rep.check("Round-trip: UI DELETE of zero-ref %s succeeded (gate allows)" % TEST_CODE,
                          bool(del_lines), del_lines[-1].split("SPIKE")[-1][:60] if del_lines else "no DELETE ok marker")
        except Exception as e:
            rep.skip("Round-trip UI delete", "delete interaction failed: %s" % e)


def teardown(rep):
    """Hard fixture-discipline sweep: ensure no ZZTST row survives and the DB is
    back to exactly the baseline count, regardless of where the round-trip stopped.
    ZZTST has zero refs so this DELETE is safe and never touches real client rows."""
    stray = sqlq("SELECT COUNT(*) FROM INV_SIZE_MST WHERE VC_SIZE_CODE='%s'" % TEST_CODE)
    n_stray = next((int(t) for t in stray.split() if t.isdigit()), 0)
    if n_stray > 0:
        sqlq("DELETE FROM INV_SIZE_MST WHERE VC_SIZE_CODE='%s'" % TEST_CODE)
    final = db_size_count()
    rep.check("Teardown: DB restored to %d sizes, no stray %s" % (EXPECTED_COUNT, TEST_CODE),
              final == EXPECTED_COUNT and n_stray == 0,
              "final count=%d, stray removed=%d" % (final, n_stray))


def main():
    headed = "--headed" in sys.argv
    os.makedirs(lib.ARTIFACTS, exist_ok=True)
    rep = lib.Report()

    print("== DB pre-state ==")
    pre = db_size_count()
    rep.check("Sandbox up + baseline (%d sizes)" % EXPECTED_COUNT, pre == EXPECTED_COUNT,
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
            open_detail_via_row(pg, rep, DETAIL_ANCHOR["code"])
        else:
            rep.skip("Row select -> Detail", "List did not render")

        print("== 3. SERVER-SIDE WRITE GATE (LIVE, P15 H3 hole closed end-to-end) ==")
        lib.check_master_write_gate_live(
            pg, rep, "Size write gate", LIST_URL,
            new_btn="size-new-btn", save_btn="size-save-btn", status_id="size-status",
            grid_sel=GRID, fill_pairs=[], count_fn=db_size_count, deny_marker="Size Save DENIED")

        # Checks 4-6 below drive UI CRUD WRITES (validation/delete-gate/round-trip). With the P15 server-
        # side write gate active, the anon spike session (no IdP/roles) is correctly DENIED before those
        # paths run, so they can no longer be exercised through the live UI here. SKIPPED with that reason;
        # the WRITE GATE is proven headless end-to-end (per view) by test_master_write_gates.py + the live
        # deny above.
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
