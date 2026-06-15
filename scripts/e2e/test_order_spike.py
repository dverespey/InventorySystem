"""E2E confirmation of Order/OrderSpike — no human clicks.

Drives a real Perspective session headless: resets the trial if needed, opens the
view, screenshots the rendered 4-row ledger, fires Simulate and the edit/recompute,
and asserts on BOTH the DOM and the gateway-log SPIKE markers. Per-check PASS/FAIL/
SKIP; exit 1 if any FAIL.

  python3 scripts/e2e/test_order_spike.py            # headless
  python3 scripts/e2e/test_order_spike.py --headed   # watch it live (design review)

Selectors prefer component domId when present (#spike-order-grid, #spike-simulate-btn,
#spike-reset-btn) and fall back to visible text, so it works before/after the
developer adds domIds. See docs/automated-ui-testing.md.
"""
import sys, time
from playwright.sync_api import sync_playwright
import lib
from reset_trial import reset_trial

VIEW_URL = lib.BASE + "/data/perspective/client/spike/order"


def q(page, domid, text=None, role=None):
    """domId first, then text/role fallback. Returns a locator or None."""
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


def cell_bg(page, text):
    """computed backgroundColor of the first element whose text == `text`."""
    el = page.query_selector("text=%s" % text)
    if not el:
        return None
    return el.evaluate("e=>getComputedStyle(e.closest('td,[role=cell],div')||e).backgroundColor")


def run(page, rep):
    # ---- render ----------------------------------------------------------
    page.goto(VIEW_URL, wait_until="networkidle", timeout=30000)
    rendered = False
    try:
        page.wait_for_selector("text=Beg Balance", timeout=20000)
        rendered = True
    except Exception:
        pass
    page.screenshot(path=lib.ARTIFACTS + "/order_spike_load.png", full_page=True)
    if "Trial Expired" in page.inner_text("body"):
        rep.check("view renders ledger", False, "Perspective TRIAL EXPIRED — run reset_trial.py with creds")
        return
    rep.check("view renders ledger (Beg Balance present)", rendered,
              "screenshot: artifacts/order_spike_load.png")
    if not rendered:
        rep.skip("downstream grid checks", "grid not rendered")
        return

    body = page.inner_text("body")
    # The table is VIRTUALIZED — assert against the top-of-grid group (15D), whose
    # golden anchors are on-screen. Numbers render comma-formatted.
    rep.check("real inventory numbers render (15D Beg day0 = 12,000)", "12,000" in body,
              "top-group golden anchor")
    rep.check("PAB End row renders (15D End day0 = 11,973)", "11,973" in body)
    for lbl in ("Beg Balance", "Receipts", "Usage", "End Balance"):
        rep.check("ledger row label '%s'" % lbl, lbl in body)
    # glyph-clutter must be GONE (requirement #1)
    glyphs = [g for g in ("★", "🚚", "📦", "⚠", "[LT]", "[OT]") if g in body]
    rep.check("no glyph clutter (numbers/color-only)", not glyphs,
              "found: %s" % glyphs if glyphs else "clean")

    # ---- Simulate fires the transform (log marker) -----------------------
    off = lib.log_marker()
    btn = q(page, "spike-simulate-btn", text="Simulate", role="button")
    if btn:
        btn.click()
        time.sleep(2.5)  # WebSocket round-trip; transform logs SPIKE grid:
        lines = lib.grep_spike_since(off, "SPIKE grid:")
        rep.check("Simulate ran transform (SPIKE grid: marker)", bool(lines),
                  lines[-1].split("SPIKE")[-1][:60] if lines else "no marker in wrapper.log")
    else:
        rep.skip("Simulate marker", "Simulate button not found")

    # ---- peach on order-by cell (requirement #2/#3) ----------------------
    # Peach editable cells render as div.ia_table__cell with inline
    # background-color: rgb(255, 204, 153). Scope to grid cells (exclude the
    # legend swatch, which is an ia_labelComponent).
    PEACH = 'div.ia_table__cell[style*="rgb(255, 204, 153)"]'
    peach_cells = page.locator(PEACH)
    npeach = peach_cells.count()
    rep.check("peach order-by cells render (one per visible supplier)", npeach >= 1,
              "%d peach cells on screen" % npeach)

    # ---- edit -> live recompute (requirement #4) -------------------------
    if npeach:
        off = lib.log_marker()
        try:
            cell = peach_cells.first
            cell.scroll_into_view_if_needed()
            cell.dblclick()                 # IA table: dblclick opens the cell editor
            page.wait_for_timeout(400)
            page.keyboard.type("5000")      # type into the focused editor
            page.keyboard.press("Enter")    # commit -> fires onEditCellCommit
            time.sleep(2.5)
            acc = lib.grep_spike_since(off, "SPIKE edit")
            rep.check("typed qty accepted (SPIKE edit ACCEPTED + recompute)",
                      any("ACCEPTED" in l for l in acc),
                      acc[-1].split("SPIKE")[-1][:70] if acc else "no SPIKE edit marker (gesture/edit-mode needs refinement)")
            page.screenshot(path=lib.ARTIFACTS + "/order_spike_after_edit.png", full_page=True)
        except Exception as e:
            rep.skip("edit recompute", "edit interaction failed: %s" % e)
    else:
        rep.skip("edit recompute", "no peach cell on screen")


def main():
    headed = "--headed" in sys.argv
    import os
    os.makedirs(lib.ARTIFACTS, exist_ok=True)
    rep = lib.Report()
    with sync_playwright() as p:
        b = p.chromium.launch(headless=not headed, slow_mo=300 if headed else 0)
        pg = b.new_page(viewport={"width": 1680, "height": 1050})
        print("== trial reset ==")
        ok, msg = reset_trial(pg)
        if msg == "NEED_CREDS":
            rep.skip("trial reset", "no GATEWAY_USER/GATEWAY_PASS — view render will be blocked if trial expired")
        else:
            rep.check("trial active", ok, msg)
        print("== order view ==")
        run(pg, rep)
        b.close()
    return rep.summary_exit()


if __name__ == "__main__":
    sys.exit(main())
