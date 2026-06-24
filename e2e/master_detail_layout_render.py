"""Render + screenshot the master-detail LAYOUT fixes (David click-through 2026-06-24) so the VISUAL
acceptance can be eyeballed (the headless CRUD/structure tests do NOT catch "compressed"/"misaligned").

Captures, in BOTH themes (tai-light + tai-dark), for Sites (the ~45-field outlier) and Size (a small
master), with a row selected so the detail box is populated:
  * <view>_layout_<theme>.png        full page (list + detail, tops should align)
  * <view>_detail_<theme>.png        the RightPane detail box cropped (Sites readability/scroll)

Theme: the master views have no toggle; we set the session theme via the /home theme-toggle button (a
session-scoped prop that persists across views in the same browser session), then navigate to the view.

  python3 e2e/master_detail_layout_render.py            # headless, writes artifacts/*.png
  python3 e2e/master_detail_layout_render.py --headed   # watch it
"""
import sys
import time

from playwright.sync_api import sync_playwright

import lib

VIEWS = [
    # (route, grid_domId, right_pane_first_row_anchor_text_to_click)
    ("sites", "sites-grid"),
    ("size", "size-grid"),
    ("partsstock", "ps-grid"),   # P21 dark-theme + click-to-fill acceptance (added 2026-06-24)
]

# Per-route detail-form domId to crop (None -> no crop).
DETAIL_FORM = {"sites": "sites-form", "partsstock": "ps-form"}
ROW = "div.ia_table__row"


def _body_is_dark(page):
    """The Perspective theme is a SESSION prop with no queryable html/body theme class; the reliable signal
    is the body background — tai-light ~ rgb(245,247,251), tai-dark ~ rgb(10,17,31). Treat a dark (low-
    luminance) body background as tai-dark."""
    bg = page.evaluate("() => getComputedStyle(document.body).backgroundColor")
    import re
    m = re.findall(r"\d+", bg or "")
    if len(m) >= 3:
        r, g, b = int(m[0]), int(m[1]), int(m[2])
        return (r + g + b) / 3 < 100
    return False


def set_theme(page, theme):
    """Set the Perspective SESSION theme via the /home toggle (session prop persists across views in this
    browser session). The toggle flips light<->dark; detect the current theme from the body background and
    click only if needed."""
    page.goto(lib.view_url("home"), wait_until="networkidle", timeout=45000)
    page.wait_for_timeout(1500)
    btn = page.query_selector("#home-theme-toggle")
    if not btn:
        print("  WARN: #home-theme-toggle not found; cannot set theme=%s deterministically" % theme)
        return
    want_dark = (theme == "tai-dark")
    tries = 0
    while _body_is_dark(page) != want_dark and tries < 2:
        btn.click()
        page.wait_for_timeout(1500)
        tries += 1
    obs = "tai-dark" if _body_is_dark(page) else "tai-light"
    print("  theme requested=%s observed=%s (%d toggle click(s))" % (theme, obs, tries))


def click_first_row(page, grid_id):
    sel = "#%s %s" % (grid_id, ROW)
    rows = page.query_selector_all(sel)
    if not rows:
        return False
    rows[0].scroll_into_view_if_needed()
    rows[0].click()
    page.wait_for_timeout(1500)
    return True


def shoot(page, route, grid_id, theme):
    page.goto(lib.view_url(route), wait_until="networkidle", timeout=45000)
    try:
        page.wait_for_selector("#" + grid_id, timeout=30000)
    except Exception:
        print("  WARN: grid #%s did not mount for %s/%s" % (grid_id, route, theme))
    page.wait_for_timeout(2000)
    # populate the detail by selecting the first grid row
    click_first_row(page, grid_id)
    page.wait_for_timeout(1200)
    full = "%s/%s_layout_%s.png" % (lib.ARTIFACTS, route, theme)
    page.screenshot(path=full, full_page=True)
    print("  wrote %s" % full)
    # crop the RightPane detail box if present
    rp = page.query_selector("div[data-component-path] >> nth=0")
    # try to locate the RightPane by its form domId (sites-form / ps-form) or the Form box
    box = None
    form_id = DETAIL_FORM.get(route)
    if form_id:
        box = page.query_selector("#" + form_id)
    if box is None:
        # fall back: screenshot the right half of the viewport
        box = None
    if box is not None:
        detail = "%s/%s_detail_%s.png" % (lib.ARTIFACTS, route, theme)
        try:
            box.screenshot(path=detail)
            print("  wrote %s" % detail)
        except Exception as e:
            print("  (detail crop skipped: %s)" % e)


def main():
    headed = "--headed" in sys.argv
    with sync_playwright() as p:
        b = p.chromium.launch(headless=not headed)
        ctx = b.new_context(viewport={"width": 1680, "height": 1050})
        page = ctx.new_page()
        body = page
        # quick trial check
        page.goto(lib.view_url("size"), wait_until="networkidle", timeout=45000)
        if "Trial Expired" in page.inner_text("body"):
            print("PERSPECTIVE TRIAL EXPIRED — run e2e/reset_trial.py first")
            b.close()
            sys.exit(2)
        for theme in ("tai-light", "tai-dark"):
            print("=== THEME %s ===" % theme)
            set_theme(page, theme)
            for route, grid_id in VIEWS:
                print(" [%s] %s" % (theme, route))
                shoot(page, route, grid_id, theme)
        b.close()
    print("DONE.")


if __name__ == "__main__":
    main()
