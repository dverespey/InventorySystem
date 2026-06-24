"""E2E QA gate for the RenbanGroup-master CRUD build — no human clicks.

Fourth master-data rebuild module. Drives a real Perspective session headless
against the live spike gateway (8.1.52 on :8088): resets the 2h trial if needed,
opens /renbangroup, and asserts on BOTH the DOM and the gateway-log SPIKE markers
(grep wrapper.log). Per-check PASS/FAIL/SKIP; exit 1 if any FAIL.

  python3 e2e/test_renbangroup_crud.py            # headless gate
  python3 e2e/test_renbangroup_crud.py --headed   # watch it live

Mirrors test_size_crud.py: reuses lib.py (Report, log markers, SPIKE grep,
trial reset), the domId-first selector and fill_field (Tabs to commit the
bidirectional binding). Views/SQL are NOT edited here.

SINGLE COMBINED VIEW: /renbangroup opens ONE view Master/RenbanGroup/RenbanGroup
containing the grid (left) AND the detail edit-form (right). Row-select is a
SAME-VIEW prop write (custom.recordId), NOT a navigation.

RENBANGROUP fields: code (text, varchar(5)), count (text, 3-char zero-padded
"001".."999"), IN_SHIP_DAYS int, + 6 per-weekday IN_SHIP_DAYS_MONDAY..SATURDAY.
Audit cols VC_LAST_UPDATE (underscore) + VC_ADD (same as Size). Validation:
presence(code) + len(code)<=5 + checkCodeUnique + counter numeric/0-999, padded
to 3 chars on save (the legacy zero-pad). Delete-gate refCount counts BOTH
INV_PARTS_STOCK_MST AND INV_PARTS_STOCK_MST_HIST by IN_RENBAN_ID.

⚠️ SPEC CORRECTION (R1): master-remainders.md §A says "Triggers: none" — WRONG.
The live DB HAS a FOR-DELETE trigger DELETE_RenbanGroupCode (verified body
2026-06-17): it nullifies INV_PARTS_STOCK_MST.IN_RENBAN_ID and IGNORES _HIST.
Same dangling-_HIST pattern as Size. D3 RESTRICT blocks on parts + _HIST.

DB ground truth (verified 2026-06-17 via sqlcmd, source-of-truth
db/namedqueries/master-crud-namedqueries.sql):
  5 renban groups. Codes by VC_RENBAN_GROUP_CODE order: CAP (id 12),
  CMWA (id 11), DICAS (id 8), HCAP (id 9), PACF (id 7).
  id 7 = PACF: parts + _HIST refs (>0; delete MUST be blocked). The exact count
  is computed live each run — _HIST accumulates across test runs (throwaway parts
  leave HIST snapshots), so it is NOT a fixed constant.
  ZZQA absent.
"""
import os
import subprocess
import sys
import time

from playwright.sync_api import sync_playwright

import lib
from reset_trial import reset_trial

LIST_URL = lib.view_url("renbangroup")
DB = lib.DB_CONN   # centralized DB-conn name (default Inventory_Spike; single prod-rename point)
SA_PASS = os.environ.get("SA_PASS", "Spike_Dev_2026!")

# Verified parity anchors (DB ground truth, see module docstring).
EXPECTED_COUNT = 5
ANCHOR_CODES = ["CAP", "CMWA", "DICAS"]      # stable codes, in code order
# DETAIL_ANCHOR.count is computed LIVE (see _resolve_detail_anchor): VC_RENBAN_GROUP_COUNT is the renban
# rollover counter, which ADVANCES every time a renban breakdown commits (P4) — hardcoding it (was '288',
# the value drifted to '297' live) makes the suite go red on counter drift. Read it from the DB at run time.
DETAIL_ANCHOR = {"id": 11, "code": "CMWA", "count": None}   # count resolved live in main()
REFERENCED = {"id": 7, "code": "PACF"}  # delete MUST be blocked; refCount computed live (drifts as _HIST grows)
TEST_CODE = "ZZQA"      # throwaway round-trip code (zero refs -> gate allows cleanup)
TEST_COUNT_IN = "7"     # typed value -> must persist/display as "007"
TEST_COUNT_OUT = "007"

GRID = "#renbangroup-grid"
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


def db_rg_count():
    out = sqlq("SELECT COUNT(*) FROM INV_RENBAN_GROUP_MST")
    for tok in out.split():
        if tok.isdigit():
            return int(tok)
    return -1


def resolve_detail_anchor():
    """Read CMWA's live VC_RENBAN_GROUP_COUNT (the renban rollover counter — advances on every P4
    breakdown commit, so it is NOT a fixed constant). Sets DETAIL_ANCHOR['count'] to the live 3-char value
    the Detail form will display. Returns the value (or None on a DB error)."""
    out = sqlq("SELECT VC_RENBAN_GROUP_COUNT FROM INV_RENBAN_GROUP_MST WHERE VC_RENBAN_GROUP_CODE='%s'"
               % DETAIL_ANCHOR["code"])
    if "SQLERR" in out:
        return None
    val = out.strip().split()[-1] if out.strip() else None
    DETAIL_ANCHOR["count"] = val
    return val


def grid_text(page):
    g = page.query_selector(GRID)
    return g.inner_text() if g else ""


# ---- check 1: List renders + parity --------------------------------------
def check_list(page, rep):
    off_grid = lib.log_marker()   # capture before load so we can grep the list-count line
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
    page.screenshot(path=lib.ARTIFACTS + "/renbangroup_list.png", full_page=True)
    rep.check("List grid mounts (%s present)" % GRID, rendered,
              "screenshot: artifacts/renbangroup_list.png")
    if not rendered:
        return False

    # The authoritative count is what RenbanGroup/list returned, logged by the
    # gateway as "SPIKE RenbanGroup/list: N rows". Assert that equals the DB count.
    rows = page.query_selector_all(GRID + " " + ROW)
    listed = None
    for l in lib.grep_spike_since(off_grid, "RenbanGroup/list:"):
        try:
            listed = int(l.split("RenbanGroup/list:")[1].strip().split()[0])
        except Exception:
            pass
    rep.check("RenbanGroup/list query returned %d rows (== DB count)" % EXPECTED_COUNT,
              listed == EXPECTED_COUNT, "query reported=%s; rendered DOM window=%d" % (listed, len(rows)))
    rep.check("List renders rows (DOM floor > 0)",
              len(rows) > 0, "rendered rows=%d" % len(rows))

    text = grid_text(page)
    missing = [c for c in ANCHOR_CODES if c not in text]
    rep.check("List anchor codes present (%s)" % ", ".join(ANCHOR_CODES), not missing,
              "missing: %s" % missing if missing else "all present")

    # parity ordering: CAP before CMWA before DICAS in grid text.
    pos = [text.find(c) for c in ANCHOR_CODES]
    ordered = all(p >= 0 for p in pos) and pos == sorted(pos)
    rep.check("List code order matches ORDER BY VC_RENBAN_GROUP_CODE (CAP<CMWA<DICAS)",
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
    click_lines = lib.grep_spike_since(off, "RenbanGroup list -> open Detail")
    rep.check("onRowClick set custom.recordId in-view (SPIKE marker, recId=%d)" % DETAIL_ANCHOR["id"],
              any(("recordId=%d" % DETAIL_ANCHOR["id"]) in l for l in click_lines),
              click_lines[-1].split("SPIKE")[-1][:70] if click_lines else "no row-click marker in wrapper.log")

    load_lines = lib.grep_spike_since(off, "RenbanGroup Detail loaded")
    rep.check("recordId onChange loaded the row into the in-view form (SPIKE 'Detail loaded id=%d')" % DETAIL_ANCHOR["id"],
              any(("id=%d" % DETAIL_ANCHOR["id"]) in l for l in load_lines),
              load_lines[-1].split("SPIKE")[-1][:70] if load_lines else "no Detail-loaded marker")

    page.wait_for_timeout(800)
    page.screenshot(path=lib.ARTIFACTS + "/renbangroup_detail.png", full_page=True)
    code_field = page.query_selector("#renbangroup-code")
    code_val = code_field.get_attribute("value") if code_field else None
    rep.check("Detail code field populated == %s" % code, code_val == code,
              "field value=%r" % code_val)
    count_field = page.query_selector("#renbangroup-count")
    count_val = count_field.get_attribute("value") if count_field else None
    rep.check("Detail count field shows %s (3-char)" % DETAIL_ANCHOR["count"],
              count_val == DETAIL_ANCHOR["count"], "count field value=%r" % count_val)
    return code_val == code


# ---- check 3: validation (blank code, >5-char, dup, counter zero-pad) ----
def check_validation(page, rep):
    """On a fresh New form: blank code rejected; >5-char code rejected; existing
    code (PACF) blocked (checkCodeUnique); counter zero-pad verified via the
    round-trip insert (separately). Asserted via #renbangroup-status text + SPIKE
    'save REJECTED' markers. No DB write occurs (all fail validation/dup-check)."""
    page.goto(LIST_URL, wait_until="networkidle", timeout=30000)
    page.wait_for_selector(GRID, timeout=20000)
    nb = q(page, "renbangroup-clear-btn", text="Clear", role="button")
    if not nb:
        rep.skip("Validation: blank code rejected", "New button not found")
        rep.skip("Validation: >5-char code rejected", "New button not found")
        rep.skip("Validation: dup code blocked", "New button not found")
        return
    nb.click()
    page.wait_for_timeout(1500)

    def set_field(domid, val):
        return fill_field(page, domid, val)

    # --- blank code ---
    off = lib.log_marker()
    set_field("renbangroup-code", "")
    set_field("renbangroup-count", "1")
    save = q(page, "renbangroup-save-btn", text="Save", role="button")
    if save:
        save.click()
        time.sleep(1.8)
        lines = lib.grep_spike_since(off, "save REJECTED: blank code")
        status = page.query_selector("#renbangroup-status")
        stxt = status.inner_text() if status else ""
        blank_ok = bool(lines) or ("code is required" in stxt)
        rep.check("Validation: blank code REJECTED (no insert)", blank_ok,
                  (lines[-1].split("SPIKE")[-1][:60] if lines else ("status=%r" % stxt[:60])))
    else:
        rep.skip("Validation: blank code rejected", "Save button not found")

    # --- >5-char code (field maxLength=5 may cap the UI input) ---
    off = lib.log_marker()
    set_field("renbangroup-code", "ABCDEF")   # 6 chars
    set_field("renbangroup-count", "1")
    save = q(page, "renbangroup-save-btn", text="Save", role="button")
    if save:
        save.click()
        time.sleep(1.8)
        lines = lib.grep_spike_since(off, "save REJECTED: code len")
        status = page.query_selector("#renbangroup-status")
        stxt = status.inner_text() if status else ""
        long_ok = bool(lines) or ("5 characters or fewer" in stxt)
        rep.check("Validation: >5-char code REJECTED (or field-capped at 5)", long_ok,
                  (lines[-1].split("SPIKE")[-1][:60] if lines
                   else ("status=%r (maxLength may have capped input to 5)" % stxt[:60])))
    else:
        rep.skip("Validation: >5-char code rejected", "Save button not found")

    # --- non-numeric counter ---
    off = lib.log_marker()
    set_field("renbangroup-code", "ZZNU")
    set_field("renbangroup-count", "AB")
    save = q(page, "renbangroup-save-btn", text="Save", role="button")
    if save:
        save.click()
        time.sleep(1.8)
        lines = lib.grep_spike_since(off, "save REJECTED: non-numeric count")
        status = page.query_selector("#renbangroup-status")
        stxt = status.inner_text() if status else ""
        num_ok = bool(lines) or ("must be numeric" in stxt)
        rep.check("Validation: non-numeric counter REJECTED", num_ok,
                  (lines[-1].split("SPIKE")[-1][:60] if lines else ("status=%r" % stxt[:60])))
    else:
        rep.skip("Validation: non-numeric counter rejected", "Save button not found")

    # --- dup code (existing PACF) ---
    off = lib.log_marker()
    set_field("renbangroup-code", REFERENCED["code"])   # PACF, already exists
    set_field("renbangroup-count", "1")
    save = q(page, "renbangroup-save-btn", text="Save", role="button")
    if save:
        save.click()
        time.sleep(1.8)
        lines = lib.grep_spike_since(off, "save REJECTED: dup code")
        status = page.query_selector("#renbangroup-status")
        stxt = status.inner_text() if status else ""
        dup_ok = bool(lines) or ("already exists" in stxt)
        rep.check("Validation: duplicate code BLOCKED (checkCodeUnique)", dup_ok,
                  (lines[-1].split("SPIKE")[-1][:60] if lines else ("status=%r" % stxt[:60])))
    else:
        rep.skip("Validation: dup code blocked", "Save button not found")

    rep.check("Validation left DB clean (still %d renban groups)" % EXPECTED_COUNT,
              db_rg_count() == EXPECTED_COUNT, "count=%d" % db_rg_count())


# ---- check 4: R1 delete-gate (CRITICAL) ----------------------------------
def check_delete_gate(page, rep):
    """Open the referenced group (PACF, id 7; refCount = parts + _HIST, computed live)
    in Detail and click Delete -> MUST be blocked. Assert: status shows the
    reference message, SPIKE refCount ran and returned >0, SPIKE DELETE BLOCKED
    fired, and the row still exists in the DB. Never deletes a referenced group.
    The _HIST rows are precisely why refCount must count _HIST (the live trigger
    ignores it)."""
    out = sqlq("SELECT (SELECT COUNT(*) FROM INV_PARTS_STOCK_MST WHERE IN_RENBAN_ID=%d)"
               "+(SELECT COUNT(*) FROM INV_PARTS_STOCK_MST_HIST WHERE IN_RENBAN_ID=%d)"
               % (REFERENCED["id"], REFERENCED["id"]))
    db_ref = next((int(t) for t in out.split() if t.isdigit()), -1)
    # NOTE: do NOT hardcode the expected count — _HIST accumulates across test runs (throwaway
    # parts leave HIST snapshots the part-delete never cleans), so PACF's refCount drifts up over
    # time. Assert it's >0 (delete must block) and that the GATE returns this SAME live value below.
    rep.check("Delete-gate pre-state: %s has %d refs in DB (parts + _HIST, >0)"
              % (REFERENCED["code"], db_ref),
              db_ref > 0, "db refCount=%d" % db_ref)

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
    delbtn = q(page, "renbangroup-delete-btn", text="Delete", role="button")
    if not delbtn:
        rep.check("Delete-gate: Delete button present", False, "not found")
        return
    delbtn.click()
    time.sleep(2.0)

    ref_lines = lib.grep_spike_since(off, "RenbanGroup refCount")
    blk_lines = lib.grep_spike_since(off, "DELETE BLOCKED")
    status = page.query_selector("#renbangroup-status")
    stxt = status.inner_text() if status else ""
    page.screenshot(path=lib.ARTIFACTS + "/renbangroup_delete_blocked.png", full_page=True)

    # the gate must report the SAME live count the DB query returned (drift-proof, not a constant)
    ran_ref = any(("n=%d" % db_ref) in l for l in ref_lines)
    rep.check("Delete-gate: refCount ran and returned the live DB value %d (SPIKE 'refCount ... n=%d')"
              % (db_ref, db_ref),
              ran_ref, ref_lines[-1].split("SPIKE")[-1][:70] if ref_lines else "no refCount marker")
    rep.check("Delete-gate: DELETE BLOCKED (SPIKE 'DELETE BLOCKED' fired)",
              bool(blk_lines), blk_lines[-1].split("SPIKE")[-1][:70] if blk_lines else "no BLOCKED marker")
    rep.check("Delete-gate: status shows the reference message ('still referenced')",
              "still referenced" in stxt, "status=%r" % stxt[:90])

    still = sqlq("SELECT COUNT(*) FROM INV_RENBAN_GROUP_MST WHERE IN_RENBAN_ID=%d" % REFERENCED["id"])
    still_n = next((int(t) for t in still.split() if t.isdigit()), -1)
    rep.check("Delete-gate: referenced group %s STILL EXISTS in DB (not deleted)" % REFERENCED["code"],
              still_n == 1, "rows for id %d = %d" % (REFERENCED["id"], still_n))


# ---- check 5: non-destructive insert/update round-trip + zero-pad --------
def check_round_trip(page, rep):
    """Insert a throwaway ZZQA group with count typed as "7" through the UI,
    verify it landed in the DB with VC_RENBAN_GROUP_COUNT == "007" (the legacy
    zero-pad), then delete it through the UI (zero refs -> gate allows) and confirm
    the DB is back to baseline. Fixture discipline: teardown sweeps any stray row."""
    pre = db_rg_count()
    if pre != EXPECTED_COUNT:
        rep.skip("Round-trip insert/delete ZZQA", "DB not at baseline (%d != %d)" % (pre, EXPECTED_COUNT))
        return
    if sqlq("SELECT COUNT(*) FROM INV_RENBAN_GROUP_MST WHERE VC_RENBAN_GROUP_CODE='%s'" % TEST_CODE).strip().split()[-1] != "0":
        rep.skip("Round-trip insert/delete ZZQA", "%s already present — refusing to collide" % TEST_CODE)
        return

    page.goto(LIST_URL, wait_until="networkidle", timeout=30000)
    page.wait_for_selector(GRID, timeout=20000)
    nb = q(page, "renbangroup-clear-btn", text="Clear", role="button")
    if not nb:
        rep.skip("Round-trip insert/delete ZZQA", "New button not found")
        return
    nb.click()
    page.wait_for_timeout(1500)

    def set_field(domid, val):
        return fill_field(page, domid, val)

    inserted = False
    try:
        set_field("renbangroup-code", TEST_CODE)
        set_field("renbangroup-count", TEST_COUNT_IN)   # "7" -> must persist as "007"
        set_field("renbangroup-shipdays", "9")
        # distinct per-weekday sentinels (1..6) so a Mon..Sat arg-transpose is caught by the
        # AUTOMATED gate (not just the manual DB check the reviewer ran).
        for dom, v in (("renbangroup-mon", "1"), ("renbangroup-tue", "2"), ("renbangroup-wed", "3"),
                       ("renbangroup-thu", "4"), ("renbangroup-fri", "5"), ("renbangroup-sat", "6")):
            set_field(dom, v)
        off = lib.log_marker()
        save = q(page, "renbangroup-save-btn", text="Save", role="button")
        if save:
            save.click()
            time.sleep(2.0)
            ins_lines = lib.grep_spike_since(off, "RenbanGroup INSERT ok")
            row = sqlq("SELECT VC_RENBAN_GROUP_COUNT, IN_SHIP_DAYS, IN_SHIP_DAYS_MONDAY "
                       "FROM INV_RENBAN_GROUP_MST WHERE VC_RENBAN_GROUP_CODE='%s'" % TEST_CODE)
            inserted = ("SQLERR" not in row and TEST_COUNT_OUT in row)
            # transpose guard: each weekday must land in its OWN column (shipdays=9, Mon..Sat=1..6)
            wk = sqlq("SELECT CAST(IN_SHIP_DAYS AS varchar)+','+CAST(IN_SHIP_DAYS_MONDAY AS varchar)+','"
                      "+CAST(IN_SHIP_DAYS_TUESDAY AS varchar)+','+CAST(IN_SHIP_DAYS_WEDNESDAY AS varchar)+','"
                      "+CAST(IN_SHIP_DAYS_THURSDAY AS varchar)+','+CAST(IN_SHIP_DAYS_FRIDAY AS varchar)+','"
                      "+CAST(IN_SHIP_DAYS_SATURDAY AS varchar) "
                      "FROM INV_RENBAN_GROUP_MST WHERE VC_RENBAN_GROUP_CODE='%s'" % TEST_CODE).strip()
            rep.check("Round-trip: 6 weekdays each land in their own column (9,1,2,3,4,5,6 — no transpose)",
                      wk == "9,1,2,3,4,5,6", "db ship-days=%r (expected 9,1,2,3,4,5,6)" % wk)
            rep.check("Round-trip: INSERT wrote %s to DB" % TEST_CODE,
                      "SQLERR" not in row and row.strip() != "",
                      "db row=%r; SPIKE=%s" %
                      (row.strip()[:60], ins_lines[-1].split("SPIKE")[-1][:40] if ins_lines else "no marker"))
            # the zero-pad assertion: typed "7" must be stored as "007"
            rep.check("Round-trip: counter zero-padded \"%s\" -> \"%s\" in DB"
                      % (TEST_COUNT_IN, TEST_COUNT_OUT),
                      TEST_COUNT_OUT in row, "db VC_RENBAN_GROUP_COUNT row=%r" % row.strip()[:40])
            # and the UI field reflects the padded value after save
            cf = page.query_selector("#renbangroup-count")
            cval = cf.get_attribute("value") if cf else None
            rep.check("Round-trip: count field displays \"%s\" after save" % TEST_COUNT_OUT,
                      cval == TEST_COUNT_OUT, "field value=%r" % cval)
            rep.check("Round-trip: DB now has %d renban groups (+1)" % (EXPECTED_COUNT + 1),
                      db_rg_count() == EXPECTED_COUNT + 1, "count=%d" % db_rg_count())
        else:
            rep.skip("Round-trip insert/delete ZZQA", "Save button not found")
    except Exception as e:
        rep.skip("Round-trip insert", "insert interaction failed: %s" % e)

    # ---- cleanup: delete the test row through the UI (zero refs -> allowed) ----
    if inserted:
        try:
            off = lib.log_marker()
            delbtn = q(page, "renbangroup-delete-btn", text="Delete", role="button")
            if delbtn:
                delbtn.click()
                time.sleep(2.0)
                del_lines = lib.grep_spike_since(off, "RenbanGroup DELETE ok")
                rep.check("Round-trip: UI DELETE of zero-ref %s succeeded (gate allows)" % TEST_CODE,
                          bool(del_lines), del_lines[-1].split("SPIKE")[-1][:60] if del_lines else "no DELETE ok marker")
        except Exception as e:
            rep.skip("Round-trip UI delete", "delete interaction failed: %s" % e)


def teardown(rep):
    """Hard fixture-discipline sweep: ensure no ZZQA row survives and the DB is
    back to exactly the baseline count, regardless of where the round-trip stopped.
    ZZQA has zero refs so this DELETE is safe and never touches real client rows."""
    stray = sqlq("SELECT COUNT(*) FROM INV_RENBAN_GROUP_MST WHERE VC_RENBAN_GROUP_CODE='%s'" % TEST_CODE)
    n_stray = next((int(t) for t in stray.split() if t.isdigit()), 0)
    if n_stray > 0:
        sqlq("DELETE FROM INV_RENBAN_GROUP_MST WHERE VC_RENBAN_GROUP_CODE='%s'" % TEST_CODE)
    final = db_rg_count()
    rep.check("Teardown: DB restored to %d renban groups, no stray %s" % (EXPECTED_COUNT, TEST_CODE),
              final == EXPECTED_COUNT and n_stray == 0,
              "final count=%d, stray removed=%d" % (final, n_stray))


def main():
    headed = "--headed" in sys.argv
    os.makedirs(lib.ARTIFACTS, exist_ok=True)
    rep = lib.Report()

    print("== DB pre-state ==")
    pre = db_rg_count()
    rep.check("Sandbox up + baseline (%d renban groups)" % EXPECTED_COUNT, pre == EXPECTED_COUNT,
              "count=%d" % pre)
    # Resolve CMWA's LIVE rollover counter (drifts on P4 commits — was a stale hardcoded '288' vs live '297').
    live_count = resolve_detail_anchor()
    rep.check("Detail anchor: CMWA live rollover counter resolved (3-char) for the Detail-count assertion",
              live_count is not None and live_count.isdigit() and len(live_count) == 3,
              "CMWA VC_RENBAN_GROUP_COUNT=%r" % live_count)

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
            pg, rep, "RenbanGroup write gate", LIST_URL,
            clear_btn="renbangroup-clear-btn", save_btn="renbangroup-save-btn",
            status_id="renbangroup-status", grid_sel=GRID, fill_pairs=[], count_fn=db_rg_count,
            deny_marker="RenbanGroup Save DENIED", primary_fill=("renbangroup-code", "ZZR"))

        # Checks 4-6 drive UI CRUD WRITES (validation/delete-gate/round-trip). With the P15 server-side
        # write gate active, the anon spike session is correctly DENIED before those paths run, so they
        # can no longer be exercised through the live UI here. SKIPPED; the gate is proven headless end-to-
        # end (per view) by test_master_write_gates.py + the live deny above.
        print("== 4-6. UI CRUD-write cases (validation/delete-gate/round-trip) ==")
        rep.skip("Validation (admin UI write)", lib.WRITE_GATE_SKIP)
        rep.skip("R1 delete-gate (admin UI write)", lib.WRITE_GATE_SKIP)
        rep.skip("Round-trip insert/delete + zero-pad (admin UI write)", lib.WRITE_GATE_SKIP)

        b.close()

    print("== teardown / fixture sweep ==")
    teardown(rep)

    return rep.summary_exit()


if __name__ == "__main__":
    sys.exit(main())
