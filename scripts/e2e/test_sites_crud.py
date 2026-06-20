"""E2E QA gate for the Sites-master CRUD build — no human clicks.

EIGHTH master-data rebuild module. Drives a real Perspective session headless
against the live spike gateway (8.1.52 on :8088): resets the 2h trial if needed,
opens /sites?qaAdmin=1, and asserts on BOTH the DOM and the gateway-log SPIKE
markers (grep wrapper.log). Per-check PASS/FAIL/SKIP; exit 1 if any FAIL.

  python3 scripts/e2e/test_sites_crud.py            # headless gate
  python3 scripts/e2e/test_sites_crud.py --headed   # watch it live

Mirrors test_size_crud.py: reuses lib.py (Report, log markers, SPIKE grep, trial
reset), the domId-first selector and fill_field (Tabs to commit the bidirectional
binding). Views/SQL are NOT edited here.

SINGLE COMBINED VIEW: /sites opens ONE view Master/Sites/Sites containing the grid
(left) AND the detail edit-form (right). Row-select is a SAME-VIEW prop write
(custom.recordId), NOT navigation.

SITES-MASTER-SPECIFIC (differs from the other masters):
  RULE #1 ADMIN-GATED. The whole view/form is Admin-only. The AUTHORITATIVE gate
    is isAuthorized(false,"Authenticated/Roles/Admin") — but this headless spike has
    NO IdP / no security levels (anonymous sessions), so it fails CLOSED. The view
    carries a SPIKE-ONLY URL escape hatch ?qaAdmin=1 (page.props.urlParams.qaAdmin)
    so the harness can exercise CRUD. This harness asserts BOTH sides:
      (a) admin path: /sites?qaAdmin=1 shows the form + allows CRUD;
      (b) NON-admin path: /sites (no flag) hides the form + the banner shows.
    ⚠️ The enforced role-gate (Designer view permissions + IdP) is OUT OF REACH
    headless — see the build report. This harness proves the in-view guard only.
  RULE #2 NOT SITE-SCOPED (inverted seam): the list shows ALL sites (no site filter).
  RULE #6 delete-gate counts the throwaway INV_PARTS_STOCK_MST.site_id (no FK yet).

DB ground truth (verified 2026-06-19 via sqlcmd, source-of-truth
docs/analysis/master-data/master-crud-namedqueries.sql + spike-inv-sites-table.sql):
  2 seed sites: id1 MAS / Tupelo MS (32 part refs), id2 HERO / San Antonio TX (15 refs).
  Both are referenced -> delete MUST be blocked for both. Anchor = MAS (id1, refCount
  computed LIVE per [[reference-headless-ignition-authoring-limits]] #6).
  ZZTST absent -> throwaway round-trip abbr (0 refs once inserted -> gate allows cleanup).
"""
import os
import subprocess
import sys
import time

from playwright.sync_api import sync_playwright

import lib
from reset_trial import reset_trial

# RULE #1: admin URL = the spike-only qaAdmin escape hatch; non-admin URL = no flag.
ADMIN_URL = lib.BASE + "/data/perspective/client/spike/sites?qaAdmin=1"
NONADMIN_URL = lib.BASE + "/data/perspective/client/spike/sites"
DB = "Inventory_Spike"
SA_PASS = os.environ.get("SA_PASS", "Spike_Dev_2026!")

EXPECTED_COUNT = 2                          # 2 seed sites (MAS, HERO)
ANCHOR_ABBRS = ["HERO", "MAS"]              # ORDER BY VC_SITE_ABBR -> HERO before MAS
DETAIL_ANCHOR = {"id": 1, "abbr": "MAS", "name": "Tupelo MS"}
REFERENCED = {"id": 1, "abbr": "MAS"}      # 32 part refs -> delete MUST be blocked
TEST_ABBR = "ZZTST"                          # throwaway round-trip abbr
NULL_ABBR = "ZZNUL"                          # throwaway abbr for the blank-retention/NULL case (S1)

GRID = "#sites-grid"
ROW = "div.ia_table__row"


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
    """Set a Perspective text/numeric input's bound value headless (Size pattern)."""
    f = page.query_selector("#" + domid)
    if not f:
        return False
    f.click()
    page.keyboard.press("Control+A")
    page.keyboard.press("Meta+A")
    page.keyboard.press("Delete")
    if val is not None and val != "":
        f.type(str(val), delay=25)
    page.keyboard.press("Tab")
    page.wait_for_timeout(300)
    return True


def select_dropdown(page, domid, value):
    """Pick a value in a Perspective dropdown (AUTO/MANUAL)."""
    f = page.query_selector("#" + domid)
    if not f:
        return False
    f.click()
    page.wait_for_timeout(400)
    opt = page.get_by_text(value, exact=True)
    if opt.count():
        opt.first.click()
        page.wait_for_timeout(300)
        return True
    page.keyboard.press("Escape")
    return False


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


def db_int(query):
    out = sqlq(query)
    for tok in out.split():
        try:
            return int(tok)
        except ValueError:
            continue
    return -1


def db_sites_count():
    return db_int("SELECT COUNT(*) FROM INV_SITES")


def db_col_is_null(col, abbr):
    """True iff <col> IS NULL for the row with VC_SITE_ABBR=<abbr>. Distinguishes
    a real NULL from a literal 0 (the whole point of S1) via CASE/IS NULL in SQL."""
    out = sqlq("SELECT CASE WHEN %s IS NULL THEN 'YESNULL' ELSE 'NOTNULL' END "
               "FROM INV_SITES WHERE VC_SITE_ABBR='%s'" % (col, abbr))
    return "YESNULL" in out


def grid_text(page):
    g = page.query_selector(GRID)
    return g.inner_text() if g else ""


# ---- check 0: NON-ADMIN gate (RULE #1) -----------------------------------
def check_nonadmin_gate(page, rep):
    """Open /sites WITHOUT the qaAdmin flag -> form must be HIDDEN and the
    'ADMIN ONLY' banner must show. Proves the in-view Admin guard fails closed."""
    page.goto(NONADMIN_URL, wait_until="networkidle", timeout=45000)
    if "Trial Expired" in page.inner_text("body"):
        rep.skip("Non-admin gate hides the form", "trial expired")
        return
    page.wait_for_timeout(2500)
    page.screenshot(path=lib.ARTIFACTS + "/sites_nonadmin.png", full_page=True)
    form = page.query_selector("#sites-form")
    # form element is either absent or not visible when isAdmin=false
    form_visible = bool(form) and form.is_visible()
    rep.check("RULE #1: NON-admin (/sites no flag) HIDES the detail form", not form_visible,
              "form present=%s visible=%s" % (bool(form), form_visible))
    banner = page.query_selector("#sites-admin-banner")
    banner_visible = bool(banner) and banner.is_visible()
    rep.check("RULE #1: NON-admin shows the 'ADMIN ONLY' banner", banner_visible,
              "banner present=%s visible=%s" % (bool(banner), banner_visible))
    # Save button must also be hidden (action bar gated)
    save = page.query_selector("#sites-save-btn")
    save_visible = bool(save) and save.is_visible()
    rep.check("RULE #1: NON-admin hides Save (no write path)", not save_visible,
              "save present=%s visible=%s" % (bool(save), save_visible))


# ---- check 1: List renders (admin) + NOT-site-scoped parity --------------
def check_list(page, rep):
    off_grid = lib.log_marker()
    page.goto(ADMIN_URL, wait_until="networkidle", timeout=45000)
    if "Trial Expired" in page.inner_text("body"):
        rep.check("List renders", False, "Perspective TRIAL EXPIRED — run reset_trial.py")
        return False
    rendered = False
    try:
        page.wait_for_selector(GRID, timeout=30000)
        rendered = True
    except Exception:
        pass
    page.wait_for_timeout(2200)
    page.screenshot(path=lib.ARTIFACTS + "/sites_list.png", full_page=True)
    rep.check("List grid mounts (%s present, admin)" % GRID, rendered,
              "screenshot: artifacts/sites_list.png")
    if not rendered:
        return False

    # The grid binding logs `SPIKE Sites/list: N rows` once the WS data round-trips,
    # which can land AFTER networkidle — poll the log for up to ~10s (the marker is
    # the authoritative count; the DOM is virtualized). [[reference-headless... #6]]
    # Force the grid binding to re-fire so a FRESH marker lands after off_grid
    # (a cached/restored session may not re-run the binding on plain goto).
    sb = page.query_selector("#sites-search-btn")
    if sb:
        off_grid = lib.log_marker()
        sb.click()
        page.wait_for_timeout(800)
    listed = None
    for _ in range(24):
        for l in lib.grep_spike_since(off_grid, "Sites/list:"):
            try:
                listed = int(l.split("Sites/list:")[1].strip().split()[0])
            except Exception:
                pass
        if listed is not None:
            break
        page.wait_for_timeout(500)
    db_n = db_sites_count()
    rep.check("RULE #2 NOT-site-scoped: Sites/list returned ALL %d sites (== DB count, no site filter)" % db_n,
              listed == db_n, "query reported=%s; DB=%d" % (listed, db_n))

    rows = page.query_selector_all(GRID + " " + ROW)
    rep.check("List renders rows (DOM floor > 0)", len(rows) > 0, "rendered rows=%d" % len(rows))

    text = grid_text(page)
    missing = [c for c in ANCHOR_ABBRS if c not in text]
    rep.check("List anchor abbrs present (%s)" % ", ".join(ANCHOR_ABBRS), not missing,
              "missing: %s" % missing if missing else "all present")
    pos = [text.find(c) for c in ANCHOR_ABBRS]
    ordered = all(p >= 0 for p in pos) and pos == sorted(pos)
    rep.check("List order matches ORDER BY VC_SITE_ABBR (HERO<MAS)", ordered, "offsets=%s" % pos)
    return True


# ---- check 2: row select -> Detail (admin) -------------------------------
def open_detail_via_row(page, rep, abbr):
    off = lib.log_marker()
    target = None
    for r in page.query_selector_all(GRID + " " + ROW):
        toks = [t.strip() for t in (r.inner_text() or "").split("\n") if t.strip()]
        if toks and toks[0] == abbr:
            target = r
            break
    if target is None:
        rep.check("Row select: row %s clickable" % abbr, False, "row not found in grid")
        return False
    target.scroll_into_view_if_needed()
    target.click()
    time.sleep(2.0)
    click_lines = lib.grep_spike_since(off, "Sites list -> open Detail")
    rep.check("onRowClick set custom.recordId in-view (recId=%d)" % DETAIL_ANCHOR["id"],
              any(("recordId=%d" % DETAIL_ANCHOR["id"]) in l for l in click_lines),
              click_lines[-1].split("SPIKE")[-1][:70] if click_lines else "no row-click marker")
    load_lines = lib.grep_spike_since(off, "Sites Detail loaded")
    rep.check("recordId onChange loaded the row (SPIKE 'Detail loaded id=%d')" % DETAIL_ANCHOR["id"],
              any(("id=%d" % DETAIL_ANCHOR["id"]) in l for l in load_lines),
              load_lines[-1].split("SPIKE")[-1][:70] if load_lines else "no Detail-loaded marker")
    page.wait_for_timeout(800)
    page.screenshot(path=lib.ARTIFACTS + "/sites_detail.png", full_page=True)
    abbr_field = page.query_selector("#sites-abbr")
    abbr_val = abbr_field.get_attribute("value") if abbr_field else None
    rep.check("Detail abbr field populated == %s" % abbr, abbr_val == abbr,
              "field value=%r" % abbr_val)
    name_field = page.query_selector("#sites-name")
    name_val = name_field.get_attribute("value") if name_field else None
    rep.check("Detail name field shows %s" % DETAIL_ANCHOR["name"],
              name_val == DETAIL_ANCHOR["name"], "name field=%r" % name_val)
    # RULE #3: read-only system fields shown but disabled. Perspective renders
    # enabled:false as `disabled` on the INNER <input> (the #domId is the wrapper
    # div, which is always enabled) — so drill into the inner input.
    def inner_disabled(domid):
        wrap = page.query_selector("#" + domid)
        if not wrap:
            return None
        inp = wrap.query_selector("input")
        if not inp:
            return None
        return not inp.is_enabled()
    id_dis = inner_disabled("sites-id")
    rep.check("RULE #3: IN_SITE_ID field shown read-only (inner input disabled)", id_dis is True,
              "inner-input disabled=%s" % id_dis)
    ein_dis = inner_disabled("sites-einseq")
    rep.check("RULE #3: IN_EIN_SEQ field shown read-only (inner input disabled)", ein_dis is True,
              "inner-input disabled=%s" % ein_dis)
    return abbr_val == abbr


# ---- check 3: validation (admin) -----------------------------------------
def check_validation(page, rep):
    """Fresh New form: blank name rejected; blank abbr rejected; fill_days>50
    rejected (CK_INV_SITES_FILL_DAYS); retention<12 rejected
    (CK_INV_SITES_DATA_RETENTION). Asserted via SPIKE 'save REJECTED' + #sites-status.
    No DB write occurs (all fail validation)."""
    page.goto(ADMIN_URL, wait_until="networkidle", timeout=30000)
    page.wait_for_selector(GRID, timeout=20000)
    nb = q(page, "sites-new-btn", text="New Site", role="button")
    if not nb:
        for n in ("blank name rejected", "blank abbr rejected", "fill_days>50 rejected", "retention<12 rejected"):
            rep.skip("Validation: " + n, "New Site button not found")
        return
    pre = db_sites_count()

    def fresh():
        nb.click()
        page.wait_for_timeout(1200)

    # blank name
    fresh()
    off = lib.log_marker()
    fill_field(page, "sites-name", "")
    fill_field(page, "sites-abbr", "QABLK")
    save = q(page, "sites-save-btn", text="Save", role="button")
    save.click(); time.sleep(1.6)
    lines = lib.grep_spike_since(off, "save REJECTED: blank name")
    stxt = (page.query_selector("#sites-status").inner_text() or "")
    rep.check("Validation: blank name REJECTED (no insert)",
              bool(lines) or ("name is required" in stxt),
              (lines[-1].split("SPIKE")[-1][:60] if lines else "status=%r" % stxt[:60]))

    # blank abbr
    fresh()
    off = lib.log_marker()
    fill_field(page, "sites-name", "QA No Abbr")
    fill_field(page, "sites-abbr", "")
    save = q(page, "sites-save-btn", text="Save", role="button")
    save.click(); time.sleep(1.6)
    lines = lib.grep_spike_since(off, "save REJECTED: blank abbr")
    stxt = (page.query_selector("#sites-status").inner_text() or "")
    rep.check("Validation: blank abbr REJECTED (no insert)",
              bool(lines) or ("abbr is required" in stxt),
              (lines[-1].split("SPIKE")[-1][:60] if lines else "status=%r" % stxt[:60]))

    # fill_days > 50 (CK_INV_SITES_FILL_DAYS)
    fresh()
    off = lib.log_marker()
    fill_field(page, "sites-name", "QA Fill")
    fill_field(page, "sites-abbr", "QAFIL")
    fill_field(page, "sites-filldays", "99")
    save = q(page, "sites-save-btn", text="Save", role="button")
    save.click(); time.sleep(1.6)
    lines = lib.grep_spike_since(off, "save REJECTED: fillDays")
    stxt = (page.query_selector("#sites-status").inner_text() or "")
    rep.check("Validation: fill_days>50 REJECTED (CK_INV_SITES_FILL_DAYS)",
              bool(lines) or ("50 or fewer" in stxt),
              (lines[-1].split("SPIKE")[-1][:70] if lines else "status=%r" % stxt[:70]))

    # retention < 12 (CK_INV_SITES_DATA_RETENTION)
    fresh()
    off = lib.log_marker()
    fill_field(page, "sites-name", "QA Ret")
    fill_field(page, "sites-abbr", "QARET")
    fill_field(page, "sites-retention", "6")
    save = q(page, "sites-save-btn", text="Save", role="button")
    save.click(); time.sleep(1.6)
    lines = lib.grep_spike_since(off, "save REJECTED: retention")
    stxt = (page.query_selector("#sites-status").inner_text() or "")
    rep.check("Validation: retention<12 REJECTED (CK_INV_SITES_DATA_RETENTION)",
              bool(lines) or (">= 12 months" in stxt),
              (lines[-1].split("SPIKE")[-1][:70] if lines else "status=%r" % stxt[:70]))

    rep.check("Validation left DB clean (still %d sites)" % pre,
              db_sites_count() == pre, "count=%d" % db_sites_count())


# ---- check 4: RESTRICT delete-gate (CRITICAL, rule #6) -------------------
def check_delete_gate(page, rep):
    """Open the referenced site (MAS, id1) and click Delete -> MUST be blocked.
    refCount counted LIVE (no hardcode). Assert: SPIKE refCount ran & returned the
    DB value (>0), SPIKE DELETE BLOCKED fired, status shows the reference message,
    and the row still exists."""
    db_ref = db_int("SELECT COUNT(*) FROM INV_PARTS_STOCK_MST WHERE site_id=%d" % REFERENCED["id"])
    rep.check("Delete-gate pre-state: %s (id %d) has %d part refs in DB (>0)"
              % (REFERENCED["abbr"], REFERENCED["id"], db_ref),
              db_ref > 0, "db refCount=%d" % db_ref)

    page.goto(ADMIN_URL, wait_until="networkidle", timeout=30000)
    page.wait_for_selector(GRID, timeout=20000)
    page.wait_for_timeout(1500)
    target = None
    for r in page.query_selector_all(GRID + " " + ROW):
        toks = [t.strip() for t in (r.inner_text() or "").split("\n") if t.strip()]
        if toks and toks[0] == REFERENCED["abbr"]:
            target = r
            break
    if target is None:
        rep.check("Delete-gate: open referenced row %s" % REFERENCED["abbr"], False, "row not found")
        return
    target.scroll_into_view_if_needed()
    target.click()
    time.sleep(2.0)

    off = lib.log_marker()
    delbtn = q(page, "sites-delete-btn", text="Delete", role="button")
    if not delbtn:
        rep.check("Delete-gate: Delete button present", False, "not found")
        return
    delbtn.click()
    time.sleep(2.0)

    ref_lines = lib.grep_spike_since(off, "Sites refCount")
    blk_lines = lib.grep_spike_since(off, "Sites DELETE BLOCKED")
    stxt = (page.query_selector("#sites-status").inner_text() or "")
    page.screenshot(path=lib.ARTIFACTS + "/sites_delete_blocked.png", full_page=True)

    ran_ref = any(("n=%d" % db_ref) in l for l in ref_lines)
    rep.check("Delete-gate: refCount ran and returned %d (SPIKE 'refCount ... n=%d')" % (db_ref, db_ref),
              ran_ref, ref_lines[-1].split("SPIKE")[-1][:70] if ref_lines else "no refCount marker")
    rep.check("Delete-gate: DELETE BLOCKED (SPIKE 'Sites DELETE BLOCKED' fired)",
              bool(blk_lines), blk_lines[-1].split("SPIKE")[-1][:70] if blk_lines else "no BLOCKED marker")
    rep.check("Delete-gate: status shows the reference message ('still referenced')",
              "still referenced" in stxt, "status=%r" % stxt[:90])

    still = db_int("SELECT COUNT(*) FROM INV_SITES WHERE IN_SITE_ID=%d" % REFERENCED["id"])
    rep.check("Delete-gate: referenced site %s STILL EXISTS in DB (not deleted)" % REFERENCED["abbr"],
              still == 1, "rows for id %d = %d" % (REFERENCED["id"], still))


# ---- check 5: insert/update/delete round-trip (admin) --------------------
def check_round_trip(page, rep):
    """Insert a throwaway ZZTST site via the UI (proves IDENTITY assign + VC_ADD
    audit stamp), verify the row + audit-stamp landed, UPDATE it (proves
    VC_LAST_UPDATE stamp), then DELETE it (0 refs -> gate allows) and confirm the
    DB is back to baseline. Read-only fields (IN_SITE_ID/IN_EIN_SEQ/audit) are
    asserted server-side from the row."""
    pre = db_sites_count()
    if db_int("SELECT COUNT(*) FROM INV_SITES WHERE VC_SITE_ABBR='%s'" % TEST_ABBR) != 0:
        rep.skip("Round-trip insert/update/delete ZZTST", "%s already present" % TEST_ABBR)
        return

    page.goto(ADMIN_URL, wait_until="networkidle", timeout=30000)
    page.wait_for_selector(GRID, timeout=20000)
    nb = q(page, "sites-new-btn", text="New Site", role="button")
    if not nb:
        rep.skip("Round-trip insert/update/delete ZZTST", "New Site button not found")
        return
    nb.click()
    page.wait_for_timeout(1200)

    inserted = False
    new_id = None
    try:
        fill_field(page, "sites-name", "QA Round Trip Site")
        fill_field(page, "sites-abbr", TEST_ABBR)
        fill_field(page, "sites-state", "TN")
        fill_field(page, "sites-duns", "999999999")
        fill_field(page, "sites-filldays", "30")
        fill_field(page, "sites-retention", "12")
        select_dropdown(page, "sites-fcmode", "MANUAL")
        off = lib.log_marker()
        save = q(page, "sites-save-btn", text="Save", role="button")
        save.click(); time.sleep(2.0)
        ins_lines = lib.grep_spike_since(off, "Sites INSERT ok")
        row = sqlq("SELECT IN_SITE_ID, VC_SITE_NAME, IN_EIN_SEQ, VC_FORECAST_IMPORT_MODE, "
                   "VC_ADD, VC_LAST_UPDATE FROM INV_SITES WHERE VC_SITE_ABBR='%s'" % TEST_ABBR)
        inserted = ("QA Round Trip Site" in row)
        new_id = db_int("SELECT IN_SITE_ID FROM INV_SITES WHERE VC_SITE_ABBR='%s'" % TEST_ABBR)
        rep.check("Round-trip: INSERT wrote %s to DB (IDENTITY assigned id %s)" % (TEST_ABBR, new_id),
                  inserted and new_id > 0,
                  "db row=%r; SPIKE=%s" % (row.strip()[:70], ins_lines[-1].split("SPIKE")[-1][:40] if ins_lines else "no marker"))
        # RULE #5: VC_ADD audit stamp present (16-char yyyymmddHHMMSSff)
        addstamp = sqlq("SELECT VC_ADD FROM INV_SITES WHERE VC_SITE_ABBR='%s'" % TEST_ABBR).strip().split()[-1] if inserted else ""
        rep.check("Round-trip: VC_ADD audit stamp written (16-char) on insert",
                  len(addstamp) == 16 and addstamp.isdigit(), "VC_ADD=%r" % addstamp)
        # RULE #3: IN_EIN_SEQ defaulted to 0 (NOT written by form); MANUAL mode round-tripped
        ein = db_int("SELECT IN_EIN_SEQ FROM INV_SITES WHERE VC_SITE_ABBR='%s'" % TEST_ABBR)
        mode = sqlq("SELECT VC_FORECAST_IMPORT_MODE FROM INV_SITES WHERE VC_SITE_ABBR='%s'" % TEST_ABBR).strip().split()[-1] if inserted else ""
        rep.check("Round-trip: IN_EIN_SEQ defaulted to 0 (RULE #3 not form-written)", ein == 0, "IN_EIN_SEQ=%d" % ein)
        rep.check("Round-trip: dropdown mode round-tripped (MANUAL)", mode == "MANUAL", "mode=%r" % mode)
        rep.check("Round-trip: DB now has %d sites (+1)" % (pre + 1),
                  db_sites_count() == pre + 1, "count=%d" % db_sites_count())
    except Exception as e:
        rep.skip("Round-trip insert", "insert interaction failed: %s" % e)

    # ---- UPDATE the row (proves VC_LAST_UPDATE stamp) ----
    if inserted:
        try:
            off = lib.log_marker()
            fill_field(page, "sites-name", "QA Round Trip EDITED")
            save = q(page, "sites-save-btn", text="Save", role="button")
            save.click(); time.sleep(2.0)
            upd_lines = lib.grep_spike_since(off, "Sites UPDATE ok")
            newname = sqlq("SELECT VC_SITE_NAME FROM INV_SITES WHERE VC_SITE_ABBR='%s'" % TEST_ABBR)
            lastupd = sqlq("SELECT VC_LAST_UPDATE FROM INV_SITES WHERE VC_SITE_ABBR='%s'" % TEST_ABBR).strip().split()[-1]
            rep.check("Round-trip: UPDATE changed name in DB", "QA Round Trip EDITED" in newname,
                      "db name=%r; SPIKE=%s" % (newname.strip()[:50], upd_lines[-1].split("SPIKE")[-1][:40] if upd_lines else "no marker"))
            rep.check("Round-trip: VC_LAST_UPDATE audit stamp written (16-char) on update",
                      len(lastupd) == 16 and lastupd.isdigit(), "VC_LAST_UPDATE=%r" % lastupd)
        except Exception as e:
            rep.skip("Round-trip update", "update interaction failed: %s" % e)

    # ---- DELETE the zero-ref row (gate allows) ----
    if inserted:
        try:
            off = lib.log_marker()
            delbtn = q(page, "sites-delete-btn", text="Delete", role="button")
            if delbtn:
                delbtn.click(); time.sleep(2.0)
                del_lines = lib.grep_spike_since(off, "Sites DELETE ok")
                rep.check("Round-trip: UI DELETE of zero-ref %s succeeded (gate allows)" % TEST_ABBR,
                          bool(del_lines), del_lines[-1].split("SPIKE")[-1][:60] if del_lines else "no DELETE ok marker")
        except Exception as e:
            rep.skip("Round-trip UI delete", "delete interaction failed: %s" % e)


# ---- check 6: blank/0 retention -> NULL (B1/N1 fix; S1) -------------------
def check_blank_retention_null(page, rep):
    """B1/N1 regression: a New site saved with Data Retention left at its blank/0
    default MUST succeed and persist IN_DATA_RETENTION **IS NULL** (not literal 0,
    which CK_INV_SITES_DATA_RETENTION rejects). Then editing that NULL-retention row
    re-saves with NO CHECK error (the get coerced NULL->0; the save must re-NULL it).
    Finally a TYPED retention of 6 is still rejected (the real guard stays)."""
    if db_int("SELECT COUNT(*) FROM INV_SITES WHERE VC_SITE_ABBR='%s'" % NULL_ABBR) != 0:
        rep.skip("S1 blank-retention -> NULL round-trip", "%s already present" % NULL_ABBR)
        return
    pre = db_sites_count()

    page.goto(ADMIN_URL, wait_until="networkidle", timeout=30000)
    page.wait_for_selector(GRID, timeout=20000)
    nb = q(page, "sites-new-btn", text="New Site", role="button")
    if not nb:
        rep.skip("S1 blank-retention -> NULL round-trip", "New Site button not found")
        return

    # ---- (a) New site, retention left BLANK/0 -> Save SUCCEEDS, persists NULL ----
    nb.click()
    page.wait_for_timeout(1200)
    saved = False
    try:
        fill_field(page, "sites-name", "QA Null Retention Site")
        fill_field(page, "sites-abbr", NULL_ABBR)
        fill_field(page, "sites-state", "TN")
        # deliberately leave sites-retention untouched (its 0/blank default) AND
        # explicitly blank it to prove blank -> NULL on insert.
        fill_field(page, "sites-retention", "")
        off = lib.log_marker()
        save = q(page, "sites-save-btn", text="Save", role="button")
        save.click(); time.sleep(2.0)
        ins_lines = lib.grep_spike_since(off, "Sites INSERT ok")
        fail_lines = lib.grep_spike_since(off, "Sites save FAILED")
        saved = db_int("SELECT COUNT(*) FROM INV_SITES WHERE VC_SITE_ABBR='%s'" % NULL_ABBR) == 1
        rep.check("S1(a): New site with BLANK retention SAVED (no CHECK error)",
                  saved and bool(ins_lines) and not fail_lines,
                  "inserted=%s; INSERT marker=%s; FAILED=%s"
                  % (saved, bool(ins_lines), fail_lines[-1].split("SPIKE")[-1][:50] if fail_lines else "none"))
        is_null = db_col_is_null("IN_DATA_RETENTION", NULL_ABBR)
        stored = db_int("SELECT ISNULL(IN_DATA_RETENTION,-1) FROM INV_SITES WHERE VC_SITE_ABBR='%s'" % NULL_ABBR)
        rep.check("S1(a): IN_DATA_RETENTION persisted as NULL (NOT literal 0)",
                  is_null, "IS NULL=%s; ISNULL(...,-1)=%d (expect -1 for NULL, NOT 0)" % (is_null, stored))
    except Exception as e:
        rep.skip("S1(a) blank-retention insert", "interaction failed: %s" % e)

    # ---- (b) Edit the NULL-retention row -> re-save with NO CHECK error ----
    if saved:
        try:
            null_id = db_int("SELECT IN_SITE_ID FROM INV_SITES WHERE VC_SITE_ABBR='%s'" % NULL_ABBR)
            page.goto(ADMIN_URL, wait_until="networkidle", timeout=30000)
            page.wait_for_selector(GRID, timeout=20000)
            page.wait_for_timeout(1200)
            target = None
            for r in page.query_selector_all(GRID + " " + ROW):
                toks = [t.strip() for t in (r.inner_text() or "").split("\n") if t.strip()]
                if toks and toks[0] == NULL_ABBR:
                    target = r
                    break
            if target is None:
                rep.check("S1(b): open NULL-retention row %s to edit" % NULL_ABBR, False, "row not found")
            else:
                target.scroll_into_view_if_needed()
                target.click()
                time.sleep(2.0)
                off = lib.log_marker()
                # change only the name; retention loaded NULL->0 in the get, so re-save
                # must re-NULL it rather than bind 0 (which the CHECK would reject).
                fill_field(page, "sites-name", "QA Null Retention EDITED")
                save = q(page, "sites-save-btn", text="Save", role="button")
                save.click(); time.sleep(2.0)
                upd_lines = lib.grep_spike_since(off, "Sites UPDATE ok")
                fail_lines = lib.grep_spike_since(off, "Sites save FAILED")
                newname = sqlq("SELECT VC_SITE_NAME FROM INV_SITES WHERE VC_SITE_ABBR='%s'" % NULL_ABBR)
                rep.check("S1(b): edit of NULL-retention site re-SAVED (no CHECK constraint error)",
                          bool(upd_lines) and not fail_lines and "QA Null Retention EDITED" in newname,
                          "UPDATE marker=%s; FAILED=%s; name=%r"
                          % (bool(upd_lines),
                             fail_lines[-1].split("SPIKE")[-1][:50] if fail_lines else "none",
                             newname.strip()[:40]))
                rep.check("S1(b): IN_DATA_RETENTION STILL NULL after edit (re-NULLed, not 0)",
                          db_col_is_null("IN_DATA_RETENTION", NULL_ABBR),
                          "IS NULL=%s" % db_col_is_null("IN_DATA_RETENTION", NULL_ABBR))
        except Exception as e:
            rep.skip("S1(b) NULL-retention edit", "interaction failed: %s" % e)

    # ---- (c) TYPED retention 6 still REJECTED (real guard intact) ----
    try:
        nb2 = q(page, "sites-new-btn", text="New Site", role="button")
        if nb2:
            nb2.click()
            page.wait_for_timeout(1200)
            off = lib.log_marker()
            fill_field(page, "sites-name", "QA Typed Ret 6")
            fill_field(page, "sites-abbr", "QART6")
            fill_field(page, "sites-retention", "6")
            save = q(page, "sites-save-btn", text="Save", role="button")
            save.click(); time.sleep(1.6)
            rej_lines = lib.grep_spike_since(off, "save REJECTED: retention")
            stxt = (page.query_selector("#sites-status").inner_text() or "")
            no_row = db_int("SELECT COUNT(*) FROM INV_SITES WHERE VC_SITE_ABBR='QART6'") == 0
            rep.check("S1(c): TYPED retention=6 still REJECTED (real guard intact, no insert)",
                      (bool(rej_lines) or ">= 12 months" in stxt) and no_row,
                      (rej_lines[-1].split("SPIKE")[-1][:60] if rej_lines else "status=%r" % stxt[:60])
                      + "; row absent=%s" % no_row)
    except Exception as e:
        rep.skip("S1(c) typed retention rejected", "interaction failed: %s" % e)

    # sweep the ZZNUL row immediately (0 refs -> safe) so teardown count stays clean
    sqlq("DELETE FROM INV_SITES WHERE VC_SITE_ABBR='%s'" % NULL_ABBR)
    rep.check("S1: spike swept back to pre-state after NULL-retention case (%d sites)" % pre,
              db_sites_count() == pre, "count=%d" % db_sites_count())


def teardown(rep, baseline):
    """Hard fixture-discipline sweep: ZZTST has 0 refs so DELETE is safe and never
    touches real client rows. Restore the spike to exactly the 2 seed sites."""
    stray = db_int("SELECT COUNT(*) FROM INV_SITES WHERE VC_SITE_ABBR IN ('%s','%s')"
                   % (TEST_ABBR, NULL_ABBR))
    if stray > 0:
        sqlq("DELETE FROM INV_SITES WHERE VC_SITE_ABBR IN ('%s','%s')" % (TEST_ABBR, NULL_ABBR))
    # also sweep any QA* validation strays (validation should never insert, but be safe)
    sqlq("DELETE FROM INV_SITES WHERE VC_SITE_ABBR IN ('QABLK','QAFIL','QARET','QART6') OR VC_SITE_NAME LIKE 'QA %'")
    final = db_sites_count()
    seeds = db_int("SELECT COUNT(*) FROM INV_SITES WHERE VC_SITE_ABBR IN ('MAS','HERO')")
    rep.check("Teardown: DB restored to %d seed sites (MAS+HERO intact), no stray %s"
              % (baseline, TEST_ABBR),
              final == baseline and seeds == 2 and stray >= 0,
              "final count=%d, seeds=%d, stray swept=%d" % (final, seeds, stray))


def main():
    headed = "--headed" in sys.argv
    os.makedirs(lib.ARTIFACTS, exist_ok=True)
    rep = lib.Report()

    print("== DB pre-state ==")
    pre = db_sites_count()
    rep.check("Sandbox up + baseline (%d seed sites)" % EXPECTED_COUNT, pre == EXPECTED_COUNT,
              "count=%d" % pre)
    baseline = pre

    with sync_playwright() as p:
        b = p.chromium.launch(headless=not headed, slow_mo=200 if headed else 0)
        pg = b.new_page(viewport={"width": 1680, "height": 1050})
        print("== trial reset ==")
        ok, msg = reset_trial(pg)
        if msg == "NEED_CREDS":
            rep.skip("trial reset", "no GATEWAY_USER/GATEWAY_PASS — render blocked if trial expired")
        else:
            rep.check("trial active", ok, msg)

        print("== 0. NON-admin gate (RULE #1) ==")
        check_nonadmin_gate(pg, rep)

        print("== 1. List renders + NOT-site-scoped parity (admin) ==")
        list_ok = check_list(pg, rep)

        print("== 2. Row select -> Detail (admin) ==")
        if list_ok:
            open_detail_via_row(pg, rep, DETAIL_ANCHOR["abbr"])
        else:
            rep.skip("Row select -> Detail", "List did not render")

        print("== 3. Validation (admin) ==")
        check_validation(pg, rep)

        print("== 4. RESTRICT delete-gate (CRITICAL) ==")
        check_delete_gate(pg, rep)

        print("== 5. Round-trip insert/update/delete (non-destructive) ==")
        check_round_trip(pg, rep)

        print("== 6. Blank/0 retention -> NULL (B1/N1 fix; S1) ==")
        check_blank_retention_null(pg, rep)

        b.close()

    print("== teardown / fixture sweep ==")
    teardown(rep, baseline)

    return rep.summary_exit()


if __name__ == "__main__":
    sys.exit(main())
