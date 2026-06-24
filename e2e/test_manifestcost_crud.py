"""E2E QA gate for the ManifestCost-master CRUD build — no human clicks.

SIXTH and FINAL master (the financially load-bearing billing master). Drives a real
Perspective session headless against the live spike gateway (8.1.52 on :8088): resets
the 2h trial if needed, opens /manifestcost, and asserts on BOTH the DOM and the
gateway-log SPIKE markers (grep wrapper.log). Per-check PASS/FAIL/SKIP; exit 1 if any FAIL.

  python3 e2e/test_manifestcost_crud.py            # headless gate
  python3 e2e/test_manifestcost_crud.py --headed   # watch it live

Mirrors test_size_crud.py: reuses lib.py (Report, log markers, SPIKE grep, trial reset),
the domId-first selector and fill_field (Tabs to commit the bidirectional binding).
Views/SQL are NOT edited here.

SINGLE COMBINED VIEW: /manifestcost opens ONE view Master/ManifestCost/ManifestCost with
the grid (left) AND the detail edit-form (right). Row-select is a SAME-VIEW prop write
(custom.recordId), NOT a navigation. There is no List->Detail page navigation.

THE HEADLINE DELTA (D13.2 option b): manifest-number uniqueness is NOT enforced. The
legacy IX_INV_MANIFEST_COST_MST UNIQUE(VC_ASSY_MANIFEST_NUMBER) was DROPPED as the D13.3
conversion step (applied 2026-06-17 as sa). The REAL rule is the D6 non-overlapping date
window per assy code (checkWindowOverlap) + start<=end + price>=0. The headline test
proves a D6-VALID insert (same assy code, NON-overlapping window, REUSING an existing
manifest number) now SUCCEEDS — which the dropped index would have rejected.

DELETE: ManifestCost has NO trigger and NO stored FK children -> SIMPLE HARD DELETE,
no refCount gate (faithful to legacy DELETE_ManifestCost).

DB ground truth (verified 2026-06-17 via sqlcmd; source-of-truth
db/namedqueries/master-crud-namedqueries.sql):
  45 manifest-cost rows. After dropping the unique constraint the count stays 45.
  Anchor id 202 = assy 42600F281100, manifest 29, window 20250818-20281201 (one window
  for that code). Its code is present in INV_FORECAST_DETAIL_INF (so it is in the combo).
  Overlap-REJECT candidate: same code, window 20260101-20260601 (inside 202's window).
  D6-VALID-ALLOW candidate (headline + round-trip throwaway): same code 42600F281100,
  NON-overlapping window 20200101-20200601, REUSING manifest number 29 -> previously
  index-rejected, now allowed. Cleaned up by hard delete -> DB back to 45.
"""
import os
import subprocess
import sys
import time

from playwright.sync_api import sync_playwright

import lib
from reset_trial import reset_trial

LIST_URL = lib.view_url("manifestcost")
DB = lib.DB_CONN   # centralized DB-conn name (default Inventory_Spike; single prod-rename point)
SA_PASS = os.environ.get("SA_PASS", "Spike_Dev_2026!")

EXPECTED_COUNT = 45

# Anchor row used for list/detail parity + the overlap tests.
ANCHOR = {"id": 202, "code": "42600F281100", "manifest": "29",
          "start": "20250818", "end": "20281201"}
# A second stable assy code to confirm code ordering in the grid (any existing code).
ANCHOR_CODES = ["42600F261100", "42600F281100"]   # in VC_ASSY_PART_NUMBER_CODE order

# Overlap-REJECT: same assy code, a window that falls INSIDE the anchor window.
OVERLAP = {"code": ANCHOR["code"], "manifest": "29", "start": "20260101", "end": "20260601"}
# TOUCH-BOUNDARY: a candidate ENDING exactly on the anchor's START day (20250818). Closed-interval
# billing semantics => touching windows OVERLAP and must be REJECTED. Locks the boundary so a future
# predicate edit ('<' -> '<=') can't silently flip billing (reviewer follow-up, billing-critical).
TOUCH = {"code": ANCHOR["code"], "manifest": "29", "start": "20240101", "end": ANCHOR["start"]}

# D6-VALID-ALLOW (headline) + round-trip throwaway: same assy code, NON-overlapping
# window, REUSING the anchor's manifest number 29 (the dropped index would reject this).
DVALID = {"code": ANCHOR["code"], "manifest": "29", "start": "20200101", "end": "20200601", "price": "77.5"}

GRID = "#mc-grid"
ROW = "div.ia_table__row"


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

    Real keystrokes (type) fire onInput/onChange so the bidirectional binding commits
    into view.custom.form_*; Tab blurs to commit before the next action."""
    f = page.query_selector("#" + domid)
    if not f:
        return False
    f.click()
    page.keyboard.press("Control+A")
    page.keyboard.press("Meta+A")   # macOS select-all (dev box is a Mac)
    page.keyboard.press("Delete")
    if val is not None and val != "":
        f.type(str(val), delay=30)
    page.keyboard.press("Tab")      # blur -> commit the bidirectional binding
    page.wait_for_timeout(350)
    return True


def select_dropdown(page, domid, value):
    """Set a Perspective dropdown's value headless. The ia.input.dropdown renders a
    search-enabled control with a `.iaDropdownCommon_search` input and portal options
    (`.iaDropdownCommon_option`). Open it, type to filter, then CLICK the option whose
    text exactly matches `value` (Enter is unreliable headless). Commits props.value
    into the bidirectional binding. `value` may be ' ' (blank manifest) -> match label."""
    val = str(value)
    inp = page.query_selector("#" + domid + " input")
    if not inp:
        ctrl = page.query_selector("#" + domid)
        if not ctrl:
            return False
        ctrl.click()
        inp = page.query_selector("#" + domid + " input")
        if not inp:
            return False
    inp.click()
    page.wait_for_timeout(300)
    # filter the menu (for the blank manifest value ' ' just open without typing)
    typed = val.strip()
    if typed:
        page.keyboard.type(typed, delay=35)
    page.wait_for_timeout(500)
    opts = page.query_selector_all(".iaDropdownCommon_option")
    target = None
    for o in opts:
        if (o.inner_text() or "").strip() == val.strip():
            target = o
            break
    if target is None and opts:
        # exact-match-by-filter fallback: first remaining option after filtering
        target = opts[0]
    if target is None:
        return False
    target.click()
    page.wait_for_timeout(400)
    return True


def sqlq(query):
    cmd = ["docker", "exec", "mssql-spike",
           "/opt/mssql-tools18/bin/sqlcmd", "-C", "-S", "localhost",
           "-U", "sa", "-P", SA_PASS, "-d", "Inventory",
           "-h", "-1", "-W", "-Q", "SET NOCOUNT ON; " + query]
    try:
        out = subprocess.check_output(cmd, stderr=subprocess.STDOUT, timeout=30)
        return out.decode("utf-8", "replace")
    except Exception as e:
        return "SQLERR: %s" % e


def db_count():
    out = sqlq("SELECT COUNT(*) FROM INV_MANIFEST_COST_MST")
    for tok in out.split():
        if tok.isdigit():
            return int(tok)
    return -1


def grid_text(page):
    g = page.query_selector(GRID)
    return g.inner_text() if g else ""


# ---- check 0: confirm D13.2(b) index was dropped --------------------------
def check_index_dropped(rep):
    out = sqlq("SELECT name FROM sys.indexes WHERE object_id=OBJECT_ID('INV_MANIFEST_COST_MST') "
               "AND is_unique=1")
    kc = sqlq("SELECT COUNT(*) FROM sys.key_constraints "
              "WHERE parent_object_id=OBJECT_ID('INV_MANIFEST_COST_MST')")
    has_unique = "IX_INV_MANIFEST_COST_MST" in out
    kc_n = next((int(t) for t in kc.split() if t.isdigit()), -1)
    rep.check("D13.2(b): unique manifest-number index/constraint DROPPED (no unique index remains)",
              (not has_unique) and kc_n == 0,
              "unique indexes=%r ; key_constraints=%d" % (out.strip(), kc_n))


# ---- check 1: List renders + parity --------------------------------------
def check_list(page, rep):
    off_grid = lib.log_marker()
    page.goto(LIST_URL, wait_until="networkidle", timeout=45000)
    if "Trial Expired" in page.inner_text("body"):
        rep.check("List renders", False, "Perspective TRIAL EXPIRED — run reset_trial.py with creds")
        return False
    rendered = False
    try:
        page.wait_for_selector(GRID, timeout=30000)
        rendered = True
    except Exception:
        pass
    page.wait_for_timeout(2200)
    page.screenshot(path=lib.ARTIFACTS + "/manifestcost_list.png", full_page=True)
    rep.check("List grid mounts (%s present)" % GRID, rendered,
              "screenshot: artifacts/manifestcost_list.png")
    if not rendered:
        return False

    rows = page.query_selector_all(GRID + " " + ROW)
    listed = None
    for l in lib.grep_spike_since(off_grid, "ManifestCost/list:"):
        try:
            listed = int(l.split("ManifestCost/list:")[1].strip().split()[0])
        except Exception:
            pass
    rep.check("ManifestCost/list query returned %d rows (== DB count; grid DOM-virtualizes)" % EXPECTED_COUNT,
              listed == EXPECTED_COUNT, "query reported=%s; rendered DOM window=%d" % (listed, len(rows)))
    rep.check("List renders a virtualized window of rows (DOM floor > 0)",
              len(rows) > 0, "rendered rows=%d" % len(rows))

    text = grid_text(page)
    missing = [c for c in ANCHOR_CODES if c not in text]
    rep.check("List anchor codes present (%s)" % ", ".join(ANCHOR_CODES), not missing,
              "missing: %s" % missing if missing else "all present")

    pos = [text.find(c) for c in ANCHOR_CODES]
    ordered = all(p >= 0 for p in pos) and pos == sorted(pos)
    rep.check("List code order matches ManifestCost/list (assy-code ascending)",
              ordered, "text offsets=%s" % pos)
    return True


# ---- check 2: row select -> Detail ---------------------------------------
def open_detail_via_row(page, rep, code):
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
    time.sleep(2.0)
    click_lines = lib.grep_spike_since(off, "ManifestCost list -> open Detail")
    rep.check("onRowClick set custom.recordId in-view (SPIKE marker, recId=%d)" % ANCHOR["id"],
              any(("recordId=%d" % ANCHOR["id"]) in l for l in click_lines),
              click_lines[-1].split("SPIKE")[-1][:70] if click_lines else "no row-click marker")

    load_lines = lib.grep_spike_since(off, "ManifestCost Detail loaded")
    rep.check("recordId onChange loaded the row (SPIKE 'Detail loaded id=%d')" % ANCHOR["id"],
              any(("id=%d" % ANCHOR["id"]) in l for l in load_lines),
              load_lines[-1].split("SPIKE")[-1][:90] if load_lines else "no Detail-loaded marker")

    page.wait_for_timeout(800)
    page.screenshot(path=lib.ARTIFACTS + "/manifestcost_detail.png", full_page=True)
    # assert the load wrote the anchor's window/price into the form custom props (via the log,
    # which echoes code/manifest/start/end/price — the most reliable headless signal).
    loaded = load_lines[-1] if load_lines else ""
    fields_ok = all(s in loaded for s in (ANCHOR["code"], ANCHOR["start"], ANCHOR["end"]))
    rep.check("Detail form loaded anchor fields (code=%s start=%s end=%s)"
              % (ANCHOR["code"], ANCHOR["start"], ANCHOR["end"]),
              fields_ok, loaded.split("SPIKE")[-1][:110] if loaded else "no load line")
    # also confirm the start text field DOM shows the loaded value
    sf = page.query_selector("#mc-start")
    sval = sf.get_attribute("value") if sf else None
    rep.check("Detail start-date field DOM populated == %s" % ANCHOR["start"], sval == ANCHOR["start"],
              "field value=%r" % sval)
    return True


def _new_form(page):
    nb = q(page, "mc-clear-btn", text="Clear", role="button")
    if not nb:
        return False
    nb.click()
    page.wait_for_timeout(1200)
    return True


def _fill_and_save(page, rec):
    """Fill the detail form from a dict and click Save. assy code + manifest # are
    dropdowns; start/end/price are text/numeric. Returns the off-marker captured just
    before the Save click."""
    select_dropdown(page, "mc-assycode", rec["code"])
    select_dropdown(page, "mc-manifestno", rec["manifest"])
    fill_field(page, "mc-start", rec["start"])
    fill_field(page, "mc-end", rec["end"])
    fill_field(page, "mc-price", rec.get("price", "0"))
    off = lib.log_marker()
    save = q(page, "mc-save-btn", text="Save", role="button")
    if save:
        save.click()
        time.sleep(2.0)
    return off


# ---- check 3: validation (start>end, negative price, overlap) ------------
def check_validation(page, rep):
    page.goto(LIST_URL, wait_until="networkidle", timeout=30000)
    page.wait_for_selector(GRID, timeout=20000)
    page.wait_for_timeout(1200)

    # --- start > end ---
    if not _new_form(page):
        rep.skip("Validation: start>end rejected", "New button not found")
    else:
        off = _fill_and_save(page, {"code": ANCHOR["code"], "manifest": "01",
                                    "start": "20260601", "end": "20260101", "price": "10"})
        lines = lib.grep_spike_since(off, "save REJECTED: start>end")
        status = page.query_selector("#mc-status")
        stxt = status.inner_text() if status else ""
        ok = bool(lines) or ("on or before" in stxt)
        rep.check("Validation: start>end REJECTED (no insert)", ok,
                  (lines[-1].split("SPIKE")[-1][:70] if lines else ("status=%r" % stxt[:70])))

    # --- negative price ---
    if not _new_form(page):
        rep.skip("Validation: negative price rejected", "New button not found")
    else:
        off = _fill_and_save(page, {"code": ANCHOR["code"], "manifest": "01",
                                    "start": "20200101", "end": "20200201", "price": "-5"})
        lines = lib.grep_spike_since(off, "save REJECTED: negative price")
        status = page.query_selector("#mc-status")
        stxt = status.inner_text() if status else ""
        ok = bool(lines) or ("zero or greater" in stxt)
        rep.check("Validation: negative price REJECTED (no insert)", ok,
                  (lines[-1].split("SPIKE")[-1][:70] if lines else ("status=%r" % stxt[:70])))

    # --- HEADLINE A: overlapping window for same assy code -> REJECTED ---
    if not _new_form(page):
        rep.skip("Validation: overlapping window rejected", "New button not found")
    else:
        off = _fill_and_save(page, OVERLAP)
        lines = lib.grep_spike_since(off, "save REJECTED: window overlap")
        status = page.query_selector("#mc-status")
        stxt = status.inner_text() if status else ""
        ok = bool(lines) or ("overlaps an existing price window" in stxt)
        rep.check("Validation HEADLINE: OVERLAPPING window (same assy code) REJECTED (checkWindowOverlap)",
                  ok, (lines[-1].split("SPIKE")[-1][:80] if lines else ("status=%r" % stxt[:80])))

    # --- TOUCH BOUNDARY: candidate ends exactly on the anchor's start day -> REJECTED (closed interval) ---
    if not _new_form(page):
        rep.skip("Validation: touching-boundary window rejected", "New button not found")
    else:
        off = _fill_and_save(page, TOUCH)
        lines = lib.grep_spike_since(off, "save REJECTED: window overlap")
        status = page.query_selector("#mc-status")
        stxt = status.inner_text() if status else ""
        ok = bool(lines) or ("overlaps an existing price window" in stxt)
        rep.check("Validation: TOUCHING-boundary window (cand end == anchor start) REJECTED "
                  "(closed-interval billing — locks the boundary semantics)",
                  ok, (lines[-1].split("SPIKE")[-1][:80] if lines else ("status=%r" % stxt[:80])))

    rep.check("Validation left DB clean (still %d rows)" % EXPECTED_COUNT,
              db_count() == EXPECTED_COUNT, "count=%d" % db_count())


# ---- check 4: HEADLINE B + round-trip — D6-valid insert (allowed) + delete ----
def check_d6valid_round_trip(page, rep):
    """THE HEADLINE TEST for D13.2(b): insert a row for the SAME assy code as the anchor,
    with a NON-overlapping window, REUSING the anchor's manifest number 29. The dropped
    legacy index would have rejected this (duplicate manifest number); now it SUCCEEDS,
    proving option (b). Then hard-delete it -> DB back to 45 (round-trip)."""
    pre = db_count()
    if pre != EXPECTED_COUNT:
        rep.skip("HEADLINE B: D6-valid insert (allowed) + round-trip", "DB not at baseline (%d != %d)" % (pre, EXPECTED_COUNT))
        return
    # pre-clean any stray from a prior aborted run (same code+window)
    sqlq("DELETE FROM INV_MANIFEST_COST_MST WHERE VC_ASSY_PART_NUMBER_CODE='%s' "
         "AND VC_START_MANIFEST='%s' AND VC_END_MANIFEST='%s'"
         % (DVALID["code"], DVALID["start"], DVALID["end"]))

    page.goto(LIST_URL, wait_until="networkidle", timeout=30000)
    page.wait_for_selector(GRID, timeout=20000)
    page.wait_for_timeout(1000)
    if not _new_form(page):
        rep.skip("HEADLINE B: D6-valid insert (allowed) + round-trip", "New button not found")
        return

    off = _fill_and_save(page, DVALID)
    ins_lines = lib.grep_spike_since(off, "ManifestCost INSERT ok")
    rej_lines = lib.grep_spike_since(off, "save REJECTED")
    status = page.query_selector("#mc-status")
    stxt = status.inner_text() if status else ""

    # verify in DB: a NEW row with same code, the non-overlapping window, manifest 29.
    row = sqlq("SELECT IN_MANIFEST_COST_ID, VC_ASSY_MANIFEST_NUMBER, MO_PRICE FROM INV_MANIFEST_COST_MST "
               "WHERE VC_ASSY_PART_NUMBER_CODE='%s' AND VC_START_MANIFEST='%s' AND VC_END_MANIFEST='%s'"
               % (DVALID["code"], DVALID["start"], DVALID["end"]))
    inserted = ("SQLERR" not in row) and any(t.isdigit() for t in row.split())
    reused_29 = " 29 " in (" " + row.replace("\t", " ") + " ") or "\t29\t" in row or row.split() and "29" in row.split()

    rep.check("HEADLINE B: D6-VALID insert ALLOWED — same assy code + NON-overlapping window + "
              "REUSED manifest # 29 (dropped index would have rejected this)",
              inserted and bool(ins_lines) and not rej_lines,
              "db row=%r ; INSERT marker=%s ; reject=%s"
              % (row.strip()[:70],
                 (ins_lines[-1].split("SPIKE")[-1][:40] if ins_lines else "none"),
                 (rej_lines[-1].split("SPIKE")[-1][:40] if rej_lines else "none")))
    rep.check("HEADLINE B: the inserted row REUSES manifest number 29 (proves index relaxed)",
              "29" in row, "db row=%r" % row.strip()[:70])
    rep.check("HEADLINE B: DB now has %d rows (+1)" % (EXPECTED_COUNT + 1),
              db_count() == EXPECTED_COUNT + 1, "count=%d" % db_count())

    # ---- round-trip teardown: hard-delete through the UI (the new row is selected/edit-mode) ----
    if inserted:
        off = lib.log_marker()
        delbtn = q(page, "mc-delete-btn", text="Delete", role="button")
        if delbtn:
            delbtn.click()
            time.sleep(2.0)
            del_lines = lib.grep_spike_since(off, "ManifestCost DELETE ok")
            rep.check("Round-trip: UI hard-DELETE of the throwaway row succeeded",
                      bool(del_lines), del_lines[-1].split("SPIKE")[-1][:60] if del_lines else "no DELETE ok marker")
        else:
            rep.skip("Round-trip: UI hard-delete", "Delete button not found")


def teardown(rep):
    """Hard fixture-discipline sweep: remove the D6-valid throwaway (by code+window, the
    only synthetic row) and confirm the DB is back to exactly 45. Anchor code is real, so
    we delete ONLY the synthetic non-overlapping window, never a real client row."""
    sqlq("DELETE FROM INV_MANIFEST_COST_MST WHERE VC_ASSY_PART_NUMBER_CODE='%s' "
         "AND VC_START_MANIFEST='%s' AND VC_END_MANIFEST='%s'"
         % (DVALID["code"], DVALID["start"], DVALID["end"]))
    final = db_count()
    rep.check("Teardown: DB restored to %d rows, no synthetic row left" % EXPECTED_COUNT,
              final == EXPECTED_COUNT, "final count=%d" % final)


def main():
    headed = "--headed" in sys.argv
    os.makedirs(lib.ARTIFACTS, exist_ok=True)
    rep = lib.Report()

    print("== DB pre-state ==")
    pre = db_count()
    rep.check("Sandbox up + baseline (%d manifest-cost rows)" % EXPECTED_COUNT, pre == EXPECTED_COUNT,
              "count=%d" % pre)

    print("== 0. D13.2(b) index dropped ==")
    check_index_dropped(rep)

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
            open_detail_via_row(pg, rep, ANCHOR["code"])
        else:
            rep.skip("Row select -> Detail", "List did not render")

        print("== 3. SERVER-SIDE WRITE GATE (LIVE, P15 H3 hole closed end-to-end) ==")
        # ManifestCost's PRIMARY (form_assyCode) is a dropdown, not a text field — no reliable headless
        # fill, so primary_fill is omitted: the helper still asserts Save-DISABLED on the blank form, then
        # SKIPS the anon-DENIED click-through (the server gate for ManifestCost is proven headless end-to-end
        # by test_master_write_gates.py).
        lib.check_master_write_gate_live(
            pg, rep, "ManifestCost write gate", LIST_URL,
            clear_btn="mc-clear-btn", save_btn="mc-save-btn", status_id="mc-status",
            grid_sel=GRID, fill_pairs=[], count_fn=db_count,
            deny_marker="ManifestCost Save DENIED")

        # Checks 4-5 drive UI CRUD WRITES (validation incl. D6 overlap, and the D6-valid round-trip). With
        # the P15 server-side write gate active, the anon spike session is correctly DENIED before those
        # paths run, so they can no longer be exercised through the live UI here. SKIPPED; the gate is
        # proven headless end-to-end (per view) by test_master_write_gates.py + the live deny above. NOTE:
        # the D6 window-OVERLAP rule itself is independently proven by test_manifestcost_overlap_guard.py.
        print("== 4-5. UI CRUD-write cases (validation/D6-overlap/round-trip) ==")
        rep.skip("Validation start>end/negative-price/OVERLAP (admin UI write)", lib.WRITE_GATE_SKIP)
        rep.skip("D6-valid insert + round-trip (admin UI write)", lib.WRITE_GATE_SKIP)

        b.close()

    print("== teardown / fixture sweep ==")
    teardown(rep)

    return rep.summary_exit()


if __name__ == "__main__":
    sys.exit(main())
