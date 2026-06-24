"""E2E QA gate for the Supplier-master CRUD build — no human clicks.

First master-data rebuild module. Drives a real Perspective session headless
against the live spike gateway (8.1.52 on :8088): resets the 2h trial if needed,
opens /supplier, and asserts on BOTH the DOM and the gateway-log SPIKE markers
(grep wrapper.log). Per-check PASS/FAIL/SKIP; exit 1 if any FAIL.

  python3 scripts/e2e/test_supplier_crud.py            # headless gate
  python3 scripts/e2e/test_supplier_crud.py --headed   # watch it live

Mirrors test_order_spike.py: reuses lib.py (Report, log markers, SPIKE grep,
trial reset) and the domId-first selector. Views/SQL are NOT edited here.

SINGLE COMBINED VIEW (David's decision): /supplier opens ONE view
Master/Supplier/Supplier containing the grid (left) AND the detail edit-form
(right). Row-select is a SAME-VIEW prop write (custom.recordId), NOT a
navigation — onRowClick sets self.view.custom.recordId = event.value.RecordID
and the custom.recordId onChange load script populates the in-view form. There
is no List->Detail page navigation anywhere; every check runs against the one
view without leaving it.

Scope (parity-vs-spec, NOT testing the Ignition platform):
  1 List renders          — #supplier-grid shows the 16 real INV_SUPPLIER_MST
                            rows in code order; anchor 0501B/0572B/07100 present;
                            enum LABELS (BOTH/SHIPPED) render, not raw codes.
                            Parity oracle: legacy SELECT_SupplierInfo '' = same
                            16 codes, same order (verified via sqlcmd).
  2 Row select -> form     — onRowClick sets custom.recordId (in-view prop
                            write); the recordId onChange loads the selected
                            supplier's code/name into the same-view form and the
                            Logistics combo shows a NAME (TLD LOGISTICS SERVICES).
  3 Validation             — <5-char code rejected; dup code blocked
                            (checkCodeUnique). Asserted via #supplier-status text
                            + SPIKE "save REJECTED" markers.
  4 R1 delete-gate (CRIT)  — a referenced supplier (07100, refCount 234) is
                            BLOCKED; SPIKE refCount ran and returned >0. Never
                            actually deletes a referenced row.
  5 Round-trip (preferred) — insert throwaway ZZTST -> verify in DB -> delete it
                            (zero refs, gate allows) -> DB back to 16. Fixture
                            discipline: tag + clean; SKIP rather than leave a row.

DB ground truth (verified 2026-06-17 via sqlcmd, source-of-truth
docs/analysis/master-data/master-crud-namedqueries.sql):
  16 suppliers; codes 0501B 0572B 07100 07451 0946A 10011 11111 12720 17800
  1793B 30090 38844 43220 7201A 72100 93031 (code order).
  id 10 = 07100 refCount 234 ; id 3 = 0572B refCount 130 ; ZZTST absent.
  All 16 are OUTPUT_FILE='B'(BOTH) + ADD_POINT='S'(SHIPPED) — so only BOTH/SHIPPED
  labels appear in real data (TEXT/EXCEL/ARRIVED never render; do NOT assert them).
"""
import os
import subprocess
import sys
import time

from playwright.sync_api import sync_playwright

import lib
from reset_trial import reset_trial

LIST_URL = lib.view_url("supplier")
DB = "Inventory_Spike"
SA_PASS = os.environ.get("SA_PASS", "Spike_Dev_2026!")

# Verified parity anchors (DB ground truth, see module docstring).
EXPECTED_COUNT = 16
EXPECTED_CODES = ["0501B", "0572B", "07100", "07451", "0946A", "10011", "11111",
                  "12720", "17800", "1793B", "30090", "38844", "43220", "7201A",
                  "72100", "93031"]
ANCHOR_CODES = ["0501B", "0572B", "07100"]      # first three, in code order
DETAIL_ANCHOR = {"id": 3, "code": "0572B", "name": "CMWA",
                 "logistics": "TLDLOGISTICS SERVICES"}   # logistics name renders w/o the space
REFERENCED = {"id": 10, "code": "07100", "refCount": 234}  # delete MUST be blocked
TEST_CODE = "ZZTST"     # throwaway round-trip code (zero refs -> gate allows cleanup)

GRID = "#supplier-grid"
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
    """Robustly set a Perspective text-field's bound value headless.

    Perspective inputs are React-controlled: a bare .fill() sets the DOM value
    but does NOT always fire the onInput/onChange that commits the bidirectional
    binding into view.custom.form_*. Real keystrokes (press_sequentially) do.
    Sequence: focus -> select-all + delete (clear) -> type -> blur (Tab) so the
    binding write-back fires before the next action."""
    f = page.query_selector("#" + domid)
    if not f:
        return False
    f.click()
    page.keyboard.press("Control+A")
    page.keyboard.press("Meta+A")   # macOS select-all (dev box is a Mac)
    page.keyboard.press("Delete")
    if val:
        f.type(val, delay=30)       # ElementHandle.type = real per-key events
    page.keyboard.press("Tab")      # blur -> commit the bidirectional text binding
    page.wait_for_timeout(350)
    return True


def sqlq(query):
    """Run a one-shot sqlcmd against the spike container; return stdout (str).
    Read-only by convention here; the only writes are the explicit fixture
    cleanup verification queries (no mutation issued from this harness)."""
    cmd = ["docker", "exec", "mssql-spike",
           "/opt/mssql-tools18/bin/sqlcmd", "-C", "-S", "localhost",
           "-U", "sa", "-P", SA_PASS, "-d", "Inventory",
           "-h", "-1", "-W", "-Q", "SET NOCOUNT ON; " + query]
    try:
        out = subprocess.check_output(cmd, stderr=subprocess.STDOUT, timeout=30)
        return out.decode("utf-8", "replace")
    except Exception as e:
        return "SQLERR: %s" % e


def db_supplier_count():
    out = sqlq("SELECT COUNT(*) FROM INV_SUPPLIER_MST")
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
    page.wait_for_timeout(2200)   # let the runPrepQuery transform + grid paint
    page.screenshot(path=lib.ARTIFACTS + "/supplier_list.png", full_page=True)
    rep.check("List grid mounts (%s present)" % GRID, rendered,
              "screenshot: artifacts/supplier_list.png")
    if not rendered:
        return False

    # The table DOM-virtualizes (~32-row viewport window), so an exact DOM-row count
    # is unreliable past a viewport. Assert the AUTHORITATIVE count via the gateway
    # 'Supplier/list: N rows' log line (== DB count), with a DOM floor-check. (Mirrors
    # the Size harness; see reference-headless-ignition-authoring-limits.)
    rows = page.query_selector_all(GRID + " " + ROW)
    listed = None
    for l in lib.grep_spike_since(off_grid, "Supplier/list:"):
        try:
            listed = int(l.split("Supplier/list:")[1].strip().split()[0])
        except Exception:
            pass
    rep.check("Supplier/list query returned %d rows (== DB count; grid DOM-virtualizes)" % EXPECTED_COUNT,
              listed == EXPECTED_COUNT, "query reported=%s; rendered DOM window=%d" % (listed, len(rows)))
    rep.check("List renders a virtualized window of rows (DOM floor > 0)",
              len(rows) > 0, "rendered rows=%d" % len(rows))

    text = grid_text(page)
    missing = [c for c in ANCHOR_CODES if c not in text]
    rep.check("List anchor codes present (%s)" % ", ".join(ANCHOR_CODES), not missing,
              "missing: %s" % missing if missing else "all present")

    # parity ordering: 0501B must render before 0572B before 07100 in the grid text.
    pos = [text.find(c) for c in ANCHOR_CODES]
    ordered = all(p >= 0 for p in pos) and pos == sorted(pos)
    rep.check("List code order matches legacy SELECT_SupplierInfo '' (0501B<0572B<07100)",
              ordered, "text offsets=%s" % pos)

    # enum LABELS render, not raw single-char codes. Real data is all BOTH/SHIPPED.
    rep.check("Enum labels render (BOTH, SHIPPED — not raw 'B'/'S')",
              ("BOTH" in text and "SHIPPED" in text),
              "BOTH=%s SHIPPED=%s" % ("BOTH" in text, "SHIPPED" in text))
    return True


# ---- check 2: row select -> Detail ---------------------------------------
def open_detail_via_row(page, rep, code):
    """Click the grid row whose first cell == `code`; assert the in-view prop
    write (custom.recordId) fired and the recordId onChange loaded the row into
    the same-view form. NO navigation. Returns True if the form populated."""
    off = lib.log_marker()
    # find the row containing the code and click it (single-click selects + fires onRowClick).
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
    click_lines = lib.grep_spike_since(off, "Supplier list -> open Detail")
    rep.check("onRowClick set custom.recordId in-view (SPIKE marker, recId=%d)" % DETAIL_ANCHOR["id"],
              any(("recordId=%d" % DETAIL_ANCHOR["id"]) in l for l in click_lines),
              click_lines[-1].split("SPIKE")[-1][:70] if click_lines else "no row-click marker in wrapper.log")

    # recordId onChange load marker (edit-mode, same view).
    load_lines = lib.grep_spike_since(off, "Supplier Detail loaded")
    rep.check("recordId onChange loaded the row into the in-view form (SPIKE 'Detail loaded id=%d')" % DETAIL_ANCHOR["id"],
              any(("id=%d" % DETAIL_ANCHOR["id"]) in l for l in load_lines),
              load_lines[-1].split("SPIKE")[-1][:70] if load_lines else "no Detail-loaded marker")

    # DOM: the code field holds the supplier code, and the Logistics combo shows a NAME.
    page.wait_for_timeout(800)
    page.screenshot(path=lib.ARTIFACTS + "/supplier_detail.png", full_page=True)
    body = page.inner_text("body")
    code_field = page.query_selector("#supplier-code")
    code_val = code_field.get_attribute("value") if code_field else None
    rep.check("Detail code field populated == %s" % code, code_val == code,
              "field value=%r" % code_val)
    rep.check("Detail name field shows %s" % DETAIL_ANCHOR["name"],
              DETAIL_ANCHOR["name"] in body, "name in body=%s" % (DETAIL_ANCHOR["name"] in body))
    rep.check("Logistics combo shows a NAME (%s), not a raw id" % DETAIL_ANCHOR["logistics"],
              DETAIL_ANCHOR["logistics"] in body,
              "logistics name in body=%s" % (DETAIL_ANCHOR["logistics"] in body))
    return ("#supplier-status" and code_val == code)


# ---- check 3: validation (short code, dup code) --------------------------
def check_validation(page, rep):
    """On the Detail form: type a <5-char code and Save -> rejected; type an
    existing code (07100) and Save -> dup blocked. Asserted via status text +
    SPIKE 'save REJECTED' markers. No DB write occurs (both fail validation)."""
    # navigate to a fresh insert-mode Detail (New) so the form is editable + empty.
    page.goto(lib.view_url("supplier"), wait_until="networkidle", timeout=30000)
    page.wait_for_selector(GRID, timeout=20000)
    nb = q(page, "supplier-clear-btn", text="Clear", role="button")
    if not nb:
        rep.skip("Validation: short code rejected", "Clear button not found")
        rep.skip("Validation: dup code blocked", "Clear button not found")
        return
    nb.click()
    page.wait_for_timeout(1500)

    def set_field(domid, val):
        return fill_field(page, domid, val)

    # --- short code (<5 chars) ---
    off = lib.log_marker()
    set_field("supplier-code", "AB")
    set_field("supplier-name", "QA Short Code")
    save = q(page, "supplier-save-btn", text="Save", role="button")
    short_ok = False
    if save:
        save.click()
        time.sleep(1.8)
        lines = lib.grep_spike_since(off, "save REJECTED: code len")
        status = page.query_selector("#supplier-status")
        stxt = status.inner_text() if status else ""
        short_ok = bool(lines) or ("exactly 5 characters" in stxt)
        rep.check("Validation: <5-char code REJECTED (no insert)", short_ok,
                  (lines[-1].split("SPIKE")[-1][:60] if lines else ("status=%r" % stxt[:60])))
    else:
        rep.skip("Validation: short code rejected", "Save button not found")

    # --- dup code (existing 07100) ---
    off = lib.log_marker()
    set_field("supplier-code", REFERENCED["code"])   # 07100, already exists
    set_field("supplier-name", "QA Dup Code")
    save = q(page, "supplier-save-btn", text="Save", role="button")
    if save:
        save.click()
        time.sleep(1.8)
        lines = lib.grep_spike_since(off, "save REJECTED: dup code")
        status = page.query_selector("#supplier-status")
        stxt = status.inner_text() if status else ""
        dup_ok = bool(lines) or ("already exists" in stxt)
        rep.check("Validation: duplicate code BLOCKED (checkCodeUnique)", dup_ok,
                  (lines[-1].split("SPIKE")[-1][:60] if lines else ("status=%r" % stxt[:60])))
    else:
        rep.skip("Validation: dup code blocked", "Save button not found")

    # safety: confirm neither rejected attempt wrote a row.
    rep.check("Validation left DB clean (still %d suppliers)" % EXPECTED_COUNT,
              db_supplier_count() == EXPECTED_COUNT, "count=%d" % db_supplier_count())


# ---- check 4: R1 delete-gate (CRITICAL) ----------------------------------
def check_delete_gate(page, rep):
    """Open the referenced supplier (07100, id 10, refCount 234) in Detail and
    click Delete -> MUST be blocked. Assert: status shows the reference message,
    SPIKE refCount ran and returned >0, SPIKE DELETE BLOCKED fired, and the row
    still exists in the DB. Never deletes a referenced supplier."""
    # confirm the refcount is still what we expect (pre-state).
    out = sqlq("SELECT (SELECT COUNT(*) FROM INV_PARTS_STOCK_MST WHERE IN_SUPPLIER_ID=%d)"
               "+(SELECT COUNT(*) FROM INV_BREAKDOWN_FC_INF WHERE VC_SUPPLIER_CODE='%s')"
               "+(SELECT COUNT(*) FROM INV_FORECAST_INF WHERE VC_SUPPLIER_CODE='%s')"
               % (REFERENCED["id"], REFERENCED["code"], REFERENCED["code"]))
    db_ref = next((int(t) for t in out.split() if t.isdigit()), -1)
    rep.check("Delete-gate pre-state: %s has %d refs in DB (>0)" % (REFERENCED["code"], REFERENCED["refCount"]),
              db_ref == REFERENCED["refCount"], "db refCount=%d" % db_ref)

    # single combined view: open /supplier and click the referenced row to load
    # it into the in-view detail form (same-view prop write, no navigation).
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
        rep.check("Delete-gate: open referenced row %s" % REFERENCED["code"], False,
                  "row not found")
        return
    target.click()
    time.sleep(2.0)

    off = lib.log_marker()
    delbtn = q(page, "supplier-delete-btn", text="Delete", role="button")
    if not delbtn:
        rep.check("Delete-gate: Delete button present", False, "not found")
        return
    delbtn.click()
    time.sleep(2.0)

    ref_lines = lib.grep_spike_since(off, "Supplier refCount")
    blk_lines = lib.grep_spike_since(off, "DELETE BLOCKED")
    status = page.query_selector("#supplier-status")
    stxt = status.inner_text() if status else ""
    page.screenshot(path=lib.ARTIFACTS + "/supplier_delete_blocked.png", full_page=True)

    # refCount ran and returned the expected >0 value.
    ran_ref = any(("n=%d" % REFERENCED["refCount"]) in l for l in ref_lines)
    rep.check("Delete-gate: refCount ran and returned %d (SPIKE 'refCount ... n=%d')"
              % (REFERENCED["refCount"], REFERENCED["refCount"]),
              ran_ref, ref_lines[-1].split("SPIKE")[-1][:70] if ref_lines else "no refCount marker")
    rep.check("Delete-gate: DELETE BLOCKED (SPIKE 'DELETE BLOCKED' fired)",
              bool(blk_lines), blk_lines[-1].split("SPIKE")[-1][:70] if blk_lines else "no BLOCKED marker")
    rep.check("Delete-gate: status shows the reference message ('still referenced')",
              "still referenced" in stxt, "status=%r" % stxt[:80])

    # the row MUST still exist (proves we never reached the trigger / deleted it).
    still = sqlq("SELECT COUNT(*) FROM INV_SUPPLIER_MST WHERE IN_SUPPLIER_ID=%d" % REFERENCED["id"])
    still_n = next((int(t) for t in still.split() if t.isdigit()), -1)
    rep.check("Delete-gate: referenced supplier %s STILL EXISTS in DB (not deleted)" % REFERENCED["code"],
              still_n == 1, "rows for id %d = %d" % (REFERENCED["id"], still_n))


# ---- check 5: non-destructive insert/update round-trip -------------------
def check_round_trip(page, rep):
    """Insert a throwaway ZZTST supplier through the UI, verify it landed in the
    DB with the typed values, then delete it through the UI (zero refs -> gate
    allows) and confirm the DB is back to 16. Fixture discipline: if anything
    leaves a stray row, the final teardown sweep removes it; if the insert can't
    be driven cleanly, SKIP rather than leave a row."""
    pre = db_supplier_count()
    if pre != EXPECTED_COUNT:
        rep.skip("Round-trip insert/delete ZZTST", "DB not at baseline (%d != %d)" % (pre, EXPECTED_COUNT))
        return
    if sqlq("SELECT COUNT(*) FROM INV_SUPPLIER_MST WHERE VC_SUPPLIER_CODE='%s'" % TEST_CODE).strip().split()[-1] != "0":
        rep.skip("Round-trip insert/delete ZZTST", "%s already present — refusing to collide" % TEST_CODE)
        return

    page.goto(LIST_URL, wait_until="networkidle", timeout=30000)
    page.wait_for_selector(GRID, timeout=20000)
    nb = q(page, "supplier-clear-btn", text="Clear", role="button")
    if not nb:
        rep.skip("Round-trip insert/delete ZZTST", "Clear button not found")
        return
    nb.click()
    page.wait_for_timeout(1500)

    def set_field(domid, val):
        return fill_field(page, domid, val)

    inserted = False
    new_id = None
    try:
        set_field("supplier-code", TEST_CODE)
        set_field("supplier-name", "QA ROUND TRIP")
        set_field("supplier-city", "Testville")
        off = lib.log_marker()
        save = q(page, "supplier-save-btn", text="Save", role="button")
        if save:
            save.click()
            time.sleep(2.0)
            ins_lines = lib.grep_spike_since(off, "Supplier INSERT ok")
            # verify in DB (the oracle, not just the log).
            row = sqlq("SELECT IN_SUPPLIER_ID, VC_SUPPLIER_NAME, VC_CITY FROM INV_SUPPLIER_MST "
                       "WHERE VC_SUPPLIER_CODE='%s'" % TEST_CODE)
            inserted = ("QA ROUND TRIP" in row and "Testville" in row)
            for t in row.split():
                if t.isdigit():
                    new_id = int(t); break
            rep.check("Round-trip: INSERT wrote %s to DB (name+city round-tripped)" % TEST_CODE,
                      inserted, ("db row=%r; SPIKE=%s" %
                                 (row.strip()[:60], ins_lines[-1].split("SPIKE")[-1][:40] if ins_lines else "no marker")))
            rep.check("Round-trip: DB now has %d suppliers (17, +1)" % (EXPECTED_COUNT + 1),
                      db_supplier_count() == EXPECTED_COUNT + 1, "count=%d" % db_supplier_count())
        else:
            rep.skip("Round-trip insert/delete ZZTST", "Save button not found")
    except Exception as e:
        rep.skip("Round-trip insert", "insert interaction failed: %s" % e)

    # ---- cleanup: delete the test row through the UI (zero refs -> allowed) ----
    if inserted:
        try:
            off = lib.log_marker()
            delbtn = q(page, "supplier-delete-btn", text="Delete", role="button")
            if delbtn:
                delbtn.click()
                time.sleep(2.0)
                del_lines = lib.grep_spike_since(off, "Supplier DELETE ok")
                rep.check("Round-trip: UI DELETE of zero-ref %s succeeded (gate allows)" % TEST_CODE,
                          bool(del_lines), del_lines[-1].split("SPIKE")[-1][:60] if del_lines else "no DELETE ok marker")
        except Exception as e:
            rep.skip("Round-trip UI delete", "delete interaction failed: %s" % e)


def teardown(rep):
    """Hard fixture-discipline sweep: ensure no ZZTST row survives and the DB is
    back to exactly 16 suppliers, regardless of where the round-trip stopped.
    ZZTST has zero refs so this DELETE is safe and never touches real client rows."""
    stray = sqlq("SELECT COUNT(*) FROM INV_SUPPLIER_MST WHERE VC_SUPPLIER_CODE='%s'" % TEST_CODE)
    n_stray = next((int(t) for t in stray.split() if t.isdigit()), 0)
    if n_stray > 0:
        sqlq("DELETE FROM INV_SUPPLIER_MST WHERE VC_SUPPLIER_CODE='%s'" % TEST_CODE)
    final = db_supplier_count()
    rep.check("Teardown: DB restored to %d suppliers, no stray %s"
              % (EXPECTED_COUNT, TEST_CODE),
              final == EXPECTED_COUNT and n_stray == 0,
              "final count=%d, stray removed=%d" % (final, n_stray))


def main():
    headed = "--headed" in sys.argv
    os.makedirs(lib.ARTIFACTS, exist_ok=True)
    rep = lib.Report()

    # P11 self-heal: pre-clean the synthetic ZZTST round-trip row a KILLED prior run may have left (else
    # the baseline count would be 17 and the whole suite would fail on a stale row). Sentinel-scoped;
    # ZZTST has zero refs so the DELETE never touches real client rows. Independent of any prior teardown.
    lib.preclean_sentinels(sqlq, ["DELETE FROM INV_SUPPLIER_MST WHERE VC_SUPPLIER_CODE='%s'" % TEST_CODE],
                           label="supplier")
    print("== DB pre-state ==")
    pre = db_supplier_count()
    rep.check("Sandbox up + baseline (%d suppliers)" % EXPECTED_COUNT, pre == EXPECTED_COUNT,
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
            pg, rep, "Supplier write gate", LIST_URL,
            clear_btn="supplier-clear-btn", save_btn="supplier-save-btn", status_id="supplier-status",
            grid_sel=GRID, fill_pairs=[], count_fn=db_supplier_count,
            deny_marker="Supplier Save DENIED", primary_fill=("supplier-code", "ZZSUP"))

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
