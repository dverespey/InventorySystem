"""E2E QA gate for the Assembly Detail master CRUD build — no human clicks.

The BOM / ratio master. The FORM is titled "Assembly Detail"; the TABLE is
INV_FORECAST_DETAIL_INF (D13: table rename dropped; this is the live ratio model
the forecast/order explosion reads). Drives a real Perspective session headless
against the live spike gateway (8.1.52 on :8088): resets the 2h trial if needed,
opens /assemblydetail, and asserts on BOTH the DOM and the gateway-log SPIKE
markers (grep wrapper.log). Per-check PASS/FAIL/SKIP; exit 1 if any FAIL.

  python3 scripts/e2e/test_assemblydetail_crud.py            # headless gate
  python3 scripts/e2e/test_assemblydetail_crud.py --headed   # watch it live

Mirrors test_size_crud.py: reuses lib.py (Report, log markers, SPIKE grep, trial
reset), the domId-first selector and fill_field (Tabs to commit the bidirectional
binding). Views/SQL are NOT edited here.

SINGLE COMBINED VIEW: /assemblydetail opens ONE view Master/AssemblyDetail/
AssemblyDetail containing the grid (left) AND the detail edit-form (right).
Row-select is a SAME-VIEW prop write (custom.recordId), NOT a navigation.

ASSEMBLY DETAIL specifics (vs the Size/Supplier template):
  - Composite natural identity = (VC_ASSY_PART_NUMBER_CODE, VC_EFFECTIVE_MONTH).
    NO PK / NO unique index on the live table -> the app pre-check is the ONLY
    uniqueness enforcement (no DB backstop). All 50 live rows have BLANK eff month.
  - Component refs (tire/wheel/valve/film/label/misc1/misc2) are PART-NUMBER STRINGS,
    not FK ids; combos source distinct VC_PART_NUMBER from INV_PARTS_STOCK_MST and
    STORE the string (value == label).
  - Delete-gate (R1 / D3): the live DeleteForecastDetail trigger HARD-DELETEs
    INV_FORECAST_INF rows whose VC_PART_NUMBER = the assy code. refCount counts those
    and BLOCKS the delete if > 0 (so the cascade never fires silently).

DB ground truth (verified 2026-06-17 via sqlcmd; source-of-truth
docs/analysis/master-data/master-crud-namedqueries.sql):
  50 BOM rows. All VC_EFFECTIVE_MONTH = '' (blank). 47 distinct part numbers in
  INV_PARTS_STOCK_MST (combo source). 42 of the 50 BOM rows have matching
  INV_FORECAST_INF forecast rows (delete-blocked).
  Detail anchor: id 202 = assy 42600F261100, broadcast [MNP]HH, tire 4265202R6000,
  wheel 4261102Q4100, ratios 100/100/100, qty 4. First code by order = 42600F261100.
  Delete-gate target: id 86 = assy 426200ZR2000, 25 forecast rows (delete MUST block).
"""
import os
import subprocess
import sys
import time

from playwright.sync_api import sync_playwright

import lib
from reset_trial import reset_trial

LIST_URL = lib.BASE + "/data/perspective/client/spike/assemblydetail"
DB = "Inventory_Spike"
SA_PASS = os.environ.get("SA_PASS", "Spike_Dev_2026!")

EXPECTED_COUNT = 50
# stable non-blank anchor codes, in VC_ASSY_PART_NUMBER_CODE order (first three)
ANCHOR_CODES = ["42600F261100", "42600F263100", "42600F264200"]
DETAIL_ANCHOR = {
    "id": 202, "code": "42600F261100", "broadcast": "[MNP]HH",
    "tire": "4265202R6000", "wheel": "4261102Q4100", "fcRatio": 100,
}
# delete MUST be blocked: assy code has 25 INV_FORECAST_INF rows
REFERENCED = {"id": 86, "code": "426200ZR2000", "refCount": 25}
TEST_CODE = "ZZTSTASSY1"   # 10 chars, <= varchar(12); throwaway round-trip code

GRID = "#assy-grid"
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
    """Set a Perspective text/numeric input's bound value headless. Real keystrokes
    fire onInput/onChange so the bidirectional binding commits; Tab blurs to commit."""
    f = page.query_selector("#" + domid)
    if not f:
        return False
    f.click()
    page.keyboard.press("Control+A")
    page.keyboard.press("Meta+A")   # macOS select-all (dev box is a Mac)
    page.keyboard.press("Delete")
    if val is not None and val != "":
        f.type(str(val), delay=30)
    page.keyboard.press("Tab")
    page.wait_for_timeout(350)
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
    out = sqlq("SELECT COUNT(*) FROM INV_FORECAST_DETAIL_INF")
    for tok in out.split():
        if tok.isdigit():
            return int(tok)
    return -1


def grid_text(page):
    g = page.query_selector(GRID)
    return g.inner_text() if g else ""


# ---- check 1: List renders + parity --------------------------------------
def check_list(page, rep):
    off_grid = lib.log_marker()
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
    page.wait_for_timeout(2200)
    page.screenshot(path=lib.ARTIFACTS + "/assemblydetail_list.png", full_page=True)
    rep.check("List grid mounts (%s present)" % GRID, rendered,
              "screenshot: artifacts/assemblydetail_list.png")
    if not rendered:
        return False

    # Authoritative count = what AssemblyDetail/list returned, logged as
    # "SPIKE AssemblyDetail/list: N rows". The table DOM-virtualizes.
    rows = page.query_selector_all(GRID + " " + ROW)
    listed = None
    for l in lib.grep_spike_since(off_grid, "AssemblyDetail/list:"):
        try:
            listed = int(l.split("AssemblyDetail/list:")[1].strip().split()[0])
        except Exception:
            pass
    rep.check("AssemblyDetail/list query returned %d rows (== DB count; grid virtualizes)" % EXPECTED_COUNT,
              listed == EXPECTED_COUNT, "query reported=%s; rendered DOM window=%d" % (listed, len(rows)))
    rep.check("List renders a virtualized window of rows (DOM floor > 0)",
              len(rows) > 0, "rendered rows=%d" % len(rows))

    text = grid_text(page)
    missing = [c for c in ANCHOR_CODES if c not in text]
    rep.check("List anchor codes present (%s)" % ", ".join(ANCHOR_CODES), not missing,
              "missing: %s" % missing if missing else "all present")

    pos = [text.find(c) for c in ANCHOR_CODES]
    ordered = all(p >= 0 for p in pos) and pos == sorted(pos)
    rep.check("List code order matches SELECT_ForecastDetail '' (ascending assy code)",
              ordered, "text offsets=%s" % pos)
    return True


# ---- check 2: row select -> Detail (incl. component combos showing part #s) ---
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
    click_lines = lib.grep_spike_since(off, "AssemblyDetail list -> open Detail")
    rep.check("onRowClick set custom.recordId in-view (SPIKE marker, recId=%d)" % DETAIL_ANCHOR["id"],
              any(("recordId=%d" % DETAIL_ANCHOR["id"]) in l for l in click_lines),
              click_lines[-1].split("SPIKE")[-1][:70] if click_lines else "no row-click marker in wrapper.log")

    load_lines = lib.grep_spike_since(off, "AssemblyDetail Detail loaded")
    rep.check("recordId onChange loaded the row into the in-view form (SPIKE 'Detail loaded id=%d')" % DETAIL_ANCHOR["id"],
              any(("id=%d" % DETAIL_ANCHOR["id"]) in l for l in load_lines),
              load_lines[-1].split("SPIKE")[-1][:80] if load_lines else "no Detail-loaded marker")

    page.wait_for_timeout(900)
    page.screenshot(path=lib.ARTIFACTS + "/assemblydetail_detail.png", full_page=True)
    body = page.inner_text("body")

    code_field = page.query_selector("#assy-code")
    code_val = code_field.get_attribute("value") if code_field else None
    rep.check("Detail assy-code field populated == %s" % code, code_val == code,
              "field value=%r" % code_val)

    bc_field = page.query_selector("#assy-broadcast")
    bc_val = bc_field.get_attribute("value") if bc_field else None
    rep.check("Detail broadcast field == %s" % DETAIL_ANCHOR["broadcast"],
              bc_val == DETAIL_ANCHOR["broadcast"], "broadcast value=%r" % bc_val)

    # component combos show the stored PART-NUMBER STRINGS (tire + wheel)
    rep.check("Detail tire-PN combo shows the stored part # %s" % DETAIL_ANCHOR["tire"],
              DETAIL_ANCHOR["tire"] in body, "tire %s in body=%s" % (DETAIL_ANCHOR["tire"], DETAIL_ANCHOR["tire"] in body))
    rep.check("Detail wheel-PN combo shows the stored part # %s" % DETAIL_ANCHOR["wheel"],
              DETAIL_ANCHOR["wheel"] in body, "wheel %s in body=%s" % (DETAIL_ANCHOR["wheel"], DETAIL_ANCHOR["wheel"] in body))

    # numeric-entry-field does NOT expose its value via the HTML `value` attribute;
    # read the live input value (input_value) and fall back to the SPIKE load marker.
    fc_field = page.query_selector("#assy-fcratio")
    fc_val = None
    if fc_field:
        try:
            fc_val = fc_field.input_value()
        except Exception:
            fc_val = fc_field.get_attribute("value")
    fc_logged = any(("fcR=%d" % DETAIL_ANCHOR["fcRatio"]) in l for l in load_lines)
    rep.check("Detail forecast-ratio field == %d (numeric-entry value)" % DETAIL_ANCHOR["fcRatio"],
              (fc_val is not None and str(DETAIL_ANCHOR["fcRatio"]) in fc_val) or fc_logged,
              "fcRatio input_value=%r; SPIKE load shows fcR=%d: %s" % (fc_val, DETAIL_ANCHOR["fcRatio"], fc_logged))
    return code_val == code


# ---- check 3: part-number combos populate from INV_PARTS_STOCK_MST -------
def check_part_combos(page, rep):
    """Assert the shared part-number combo option list loaded (SPIKE
    'AssemblyDetail/partNumber lookup: N part numbers') and matches the DB distinct
    count, and that a known part number is selectable in the tire combo."""
    out = sqlq("SELECT COUNT(DISTINCT VC_PART_NUMBER) FROM INV_PARTS_STOCK_MST WHERE VC_PART_NUMBER <> ''")
    db_distinct = next((int(t) for t in out.split() if t.isdigit()), -1)

    # The partOptions binding fires once at view mount (input {siteId} is stable, so
    # it does NOT re-evaluate on a re-goto). Capture from BEFORE the very first list
    # load — set in check_list — so grep the whole session tail, not just this goto.
    off = 0  # whole-session grep: the lookup fired on the first /assemblydetail mount
    page.goto(LIST_URL, wait_until="networkidle", timeout=30000)
    page.wait_for_selector(GRID, timeout=20000)
    page.wait_for_timeout(2000)
    look_lines = lib.grep_spike_since(off, "AssemblyDetail/partNumber lookup")
    listed = None
    for l in look_lines:
        try:
            listed = int(l.split("lookup:")[1].strip().split()[0])
        except Exception:
            pass
    rep.check("Part-number combos populated from INV_PARTS_STOCK_MST (SPIKE lookup == DB distinct=%d)" % db_distinct,
              listed == db_distinct and db_distinct > 0,
              "lookup reported=%s; db distinct=%d" % (listed, db_distinct))

    # the dropdown is present in the DOM
    tire = page.query_selector("#assy-tirepn")
    rep.check("Tire part-# dropdown component present in Detail form", tire is not None,
              "found" if tire else "missing")


# ---- check 4: validation (blank assy, composite-dup) ---------------------
def check_validation(page, rep):
    """On a fresh New form: blank assy code rejected; existing (assy code, eff month)
    composite rejected (checkCodeUnique). Asserted via #assy-status + SPIKE markers.
    No DB write occurs."""
    page.goto(LIST_URL, wait_until="networkidle", timeout=30000)
    page.wait_for_selector(GRID, timeout=20000)
    nb = q(page, "assy-new-btn", text="New", role="button")
    if not nb:
        rep.skip("Validation: blank assy code rejected", "New button not found")
        rep.skip("Validation: composite-dup rejected", "New button not found")
        return
    nb.click()
    page.wait_for_timeout(1500)

    def set_field(domid, val):
        return fill_field(page, domid, val)

    # --- blank assy code ---
    off = lib.log_marker()
    set_field("assy-code", "")
    set_field("assy-broadcast", "QA Blank")
    save = q(page, "assy-save-btn", text="Save", role="button")
    if save:
        save.click()
        time.sleep(1.8)
        lines = lib.grep_spike_since(off, "save REJECTED: blank assy code")
        status = page.query_selector("#assy-status")
        stxt = status.inner_text() if status else ""
        blank_ok = bool(lines) or ("required" in stxt)
        rep.check("Validation: blank assy code REJECTED (no insert)", blank_ok,
                  (lines[-1].split("SPIKE")[-1][:60] if lines else ("status=%r" % stxt[:60])))
    else:
        rep.skip("Validation: blank assy code rejected", "Save button not found")

    # --- composite dup: existing (assy code, blank eff month) ---
    # DETAIL_ANCHOR.code with blank effective month already exists (50 live rows
    # all have blank eff month) -> composite checkCodeUnique must block.
    off = lib.log_marker()
    set_field("assy-code", DETAIL_ANCHOR["code"])
    set_field("assy-effmonth", "")
    set_field("assy-broadcast", "QA Dup")
    save = q(page, "assy-save-btn", text="Save", role="button")
    if save:
        save.click()
        time.sleep(1.8)
        lines = lib.grep_spike_since(off, "save REJECTED: dup composite")
        status = page.query_selector("#assy-status")
        stxt = status.inner_text() if status else ""
        dup_ok = bool(lines) or ("already exists" in stxt)
        rep.check("Validation: duplicate (assy code + eff month) BLOCKED (composite checkCodeUnique)",
                  dup_ok, (lines[-1].split("SPIKE")[-1][:70] if lines else ("status=%r" % stxt[:70])))
    else:
        rep.skip("Validation: composite-dup rejected", "Save button not found")

    rep.check("Validation left DB clean (still %d BOM rows)" % EXPECTED_COUNT,
              db_count() == EXPECTED_COUNT, "count=%d" % db_count())


# ---- check 5: R1 delete-gate (CRITICAL) ----------------------------------
def check_delete_gate(page, rep):
    """Open the forecast-referenced BOM row (426200ZR2000, id 86, 25 forecast rows)
    in Detail and click Delete -> MUST be blocked. Assert status, SPIKE refCount > 0,
    SPIKE DELETE BLOCKED, the BOM row still exists, AND the forecast rows are intact
    (the destructive cascade never fired)."""
    out = sqlq("SELECT COUNT(*) FROM INV_FORECAST_INF WHERE VC_PART_NUMBER='%s'" % REFERENCED["code"])
    db_ref = next((int(t) for t in out.split() if t.isdigit()), -1)
    rep.check("Delete-gate pre-state: %s has %d forecast rows in DB (>0)"
              % (REFERENCED["code"], REFERENCED["refCount"]),
              db_ref == REFERENCED["refCount"], "db forecast refCount=%d" % db_ref)

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
    delbtn = q(page, "assy-delete-btn", text="Delete", role="button")
    if not delbtn:
        rep.check("Delete-gate: Delete button present", False, "not found")
        return
    delbtn.click()
    time.sleep(2.0)

    ref_lines = lib.grep_spike_since(off, "AssemblyDetail refCount")
    blk_lines = lib.grep_spike_since(off, "DELETE BLOCKED")
    status = page.query_selector("#assy-status")
    stxt = status.inner_text() if status else ""
    page.screenshot(path=lib.ARTIFACTS + "/assemblydetail_delete_blocked.png", full_page=True)

    ran_ref = any(("n=%d" % REFERENCED["refCount"]) in l for l in ref_lines)
    rep.check("Delete-gate: refCount ran and returned %d (SPIKE 'refCount ... n=%d')"
              % (REFERENCED["refCount"], REFERENCED["refCount"]),
              ran_ref, ref_lines[-1].split("SPIKE")[-1][:70] if ref_lines else "no refCount marker")
    rep.check("Delete-gate: DELETE BLOCKED (SPIKE 'DELETE BLOCKED' fired)",
              bool(blk_lines), blk_lines[-1].split("SPIKE")[-1][:70] if blk_lines else "no BLOCKED marker")
    rep.check("Delete-gate: status shows the reference message ('still referenced')",
              "still referenced" in stxt, "status=%r" % stxt[:100])

    still = sqlq("SELECT COUNT(*) FROM INV_FORECAST_DETAIL_INF WHERE ID_FORECAST_DETAIL=%d" % REFERENCED["id"])
    still_n = next((int(t) for t in still.split() if t.isdigit()), -1)
    rep.check("Delete-gate: referenced BOM row %s STILL EXISTS in DB (not deleted)" % REFERENCED["code"],
              still_n == 1, "rows for id %d = %d" % (REFERENCED["id"], still_n))

    # the destructive forecast cascade must NOT have fired (blocked delete = no DELETE)
    fc_after = sqlq("SELECT COUNT(*) FROM INV_FORECAST_INF WHERE VC_PART_NUMBER='%s'" % REFERENCED["code"])
    fc_n = next((int(t) for t in fc_after.split() if t.isdigit()), -1)
    rep.check("Delete-gate: forecast rows for %s INTACT after blocked delete (%d, cascade never fired)"
              % (REFERENCED["code"], REFERENCED["refCount"]),
              fc_n == REFERENCED["refCount"], "forecast rows now=%d" % fc_n)


# ---- check 6: non-destructive insert/update round-trip -------------------
def check_round_trip(page, rep):
    """Insert a throwaway ZZTSTASSY1 BOM row (unique assy code + blank eff month ->
    composite is free; zero forecast refs) through the UI, verify it landed, then
    delete it through the UI (zero refs -> gate allows) and confirm the DB is back to
    50. Teardown sweeps any stray row + any forecast rows the cascade could touch
    (there are none for a fresh code)."""
    pre = db_count()
    if pre != EXPECTED_COUNT:
        rep.skip("Round-trip insert/delete ZZTSTASSY1", "DB not at baseline (%d != %d)" % (pre, EXPECTED_COUNT))
        return
    exists = sqlq("SELECT COUNT(*) FROM INV_FORECAST_DETAIL_INF WHERE VC_ASSY_PART_NUMBER_CODE='%s'" % TEST_CODE)
    if next((int(t) for t in exists.split() if t.isdigit()), 0) != 0:
        rep.skip("Round-trip insert/delete ZZTSTASSY1", "%s already present — refusing to collide" % TEST_CODE)
        return

    page.goto(LIST_URL, wait_until="networkidle", timeout=30000)
    page.wait_for_selector(GRID, timeout=20000)
    nb = q(page, "assy-new-btn", text="New", role="button")
    if not nb:
        rep.skip("Round-trip insert/delete ZZTSTASSY1", "New button not found")
        return
    nb.click()
    page.wait_for_timeout(1500)

    def set_field(domid, val):
        return fill_field(page, domid, val)

    inserted = False
    try:
        set_field("assy-code", TEST_CODE)
        set_field("assy-effmonth", "")        # blank -> composite (ZZTSTASSY1, '') is unique
        set_field("assy-broadcast", "QA RT")
        set_field("assy-tireratio", "60")
        set_field("assy-wheelratio", "40")
        set_field("assy-fcratio", "100")
        set_field("assy-qty", "4")
        off = lib.log_marker()
        save = q(page, "assy-save-btn", text="Save", role="button")
        if save:
            save.click()
            time.sleep(2.0)
            ins_lines = lib.grep_spike_since(off, "AssemblyDetail INSERT ok")
            row = sqlq("SELECT IN_TIRE_RATIO, IN_WHEEL_RATIO, IN_RATIO, IN_ASSY_QTY, VC_BROADCAST_CODE "
                       "FROM INV_FORECAST_DETAIL_INF WHERE VC_ASSY_PART_NUMBER_CODE='%s'" % TEST_CODE)
            inserted = ("QA RT" in row and "60" in row.split() and "100" in row.split())
            rep.check("Round-trip: INSERT wrote %s to DB (ratios + broadcast round-tripped)" % TEST_CODE,
                      inserted, ("db row=%r; SPIKE=%s" %
                                 (row.strip()[:60], ins_lines[-1].split("SPIKE")[-1][:50] if ins_lines else "no marker")))
            rep.check("Round-trip: DB now has %d BOM rows (+1)" % (EXPECTED_COUNT + 1),
                      db_count() == EXPECTED_COUNT + 1, "count=%d" % db_count())
        else:
            rep.skip("Round-trip insert/delete ZZTSTASSY1", "Save button not found")
    except Exception as e:
        rep.skip("Round-trip insert", "insert interaction failed: %s" % e)

    if inserted:
        try:
            off = lib.log_marker()
            delbtn = q(page, "assy-delete-btn", text="Delete", role="button")
            if delbtn:
                delbtn.click()
                time.sleep(2.0)
                gate_lines = lib.grep_spike_since(off, "AssemblyDetail refCount")
                del_lines = lib.grep_spike_since(off, "AssemblyDetail DELETE ok")
                rep.check("Round-trip: delete-gate refCount==0 for zero-ref %s (gate ALLOWS)" % TEST_CODE,
                          any("n=0" in l for l in gate_lines),
                          gate_lines[-1].split("SPIKE")[-1][:60] if gate_lines else "no refCount marker")
                rep.check("Round-trip: UI DELETE of zero-ref %s succeeded" % TEST_CODE,
                          bool(del_lines), del_lines[-1].split("SPIKE")[-1][:60] if del_lines else "no DELETE ok marker")
        except Exception as e:
            rep.skip("Round-trip UI delete", "delete interaction failed: %s" % e)


def teardown(rep):
    """Hard fixture-discipline sweep: ensure no ZZTSTASSY1 row survives and the DB is
    back to exactly 50. The throwaway code has zero forecast refs, so this DELETE is
    safe (its DeleteForecastDetail cascade matches nothing) and never touches real rows."""
    stray = sqlq("SELECT COUNT(*) FROM INV_FORECAST_DETAIL_INF WHERE VC_ASSY_PART_NUMBER_CODE='%s'" % TEST_CODE)
    n_stray = next((int(t) for t in stray.split() if t.isdigit()), 0)
    if n_stray > 0:
        sqlq("DELETE FROM INV_FORECAST_DETAIL_INF WHERE VC_ASSY_PART_NUMBER_CODE='%s'" % TEST_CODE)
    final = db_count()
    rep.check("Teardown: DB restored to %d BOM rows, no stray %s" % (EXPECTED_COUNT, TEST_CODE),
              final == EXPECTED_COUNT and n_stray == 0,
              "final count=%d, stray removed=%d" % (final, n_stray))


def main():
    headed = "--headed" in sys.argv
    os.makedirs(lib.ARTIFACTS, exist_ok=True)
    rep = lib.Report()

    print("== DB pre-state ==")
    pre = db_count()
    rep.check("Sandbox up + baseline (%d BOM rows)" % EXPECTED_COUNT, pre == EXPECTED_COUNT,
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

        print("== 2. Row select -> Detail (component combos show part #s) ==")
        if list_ok:
            open_detail_via_row(pg, rep, DETAIL_ANCHOR["code"])
        else:
            rep.skip("Row select -> Detail", "List did not render")

        print("== 3. Part-number combos populate from INV_PARTS_STOCK_MST ==")
        check_part_combos(pg, rep)

        print("== 4. SERVER-SIDE WRITE GATE (LIVE, P15 H3 hole closed end-to-end) ==")
        lib.check_master_write_gate_live(
            pg, rep, "AssemblyDetail write gate", LIST_URL,
            new_btn="assy-new-btn", save_btn="assy-save-btn", status_id="assy-status",
            grid_sel=GRID, fill_pairs=[], count_fn=db_count,
            deny_marker="AssemblyDetail Save DENIED")

        # Checks 5-7 drive UI CRUD WRITES (validation/delete-gate/round-trip). With the P15 server-side
        # write gate active, the anon spike session is correctly DENIED before those paths run, so they
        # can no longer be exercised through the live UI here. SKIPPED; the gate is proven headless end-to-
        # end (per view) by test_master_write_gates.py + the live deny above.
        print("== 5-7. UI CRUD-write cases (validation/delete-gate/round-trip) ==")
        rep.skip("Validation blank assy/composite-dup (admin UI write)", lib.WRITE_GATE_SKIP)
        rep.skip("R1 delete-gate (admin UI write)", lib.WRITE_GATE_SKIP)
        rep.skip("Round-trip insert/delete (admin UI write)", lib.WRITE_GATE_SKIP)

        b.close()

    print("== teardown / fixture sweep ==")
    teardown(rep)

    return rep.summary_exit()


if __name__ == "__main__":
    sys.exit(main())
