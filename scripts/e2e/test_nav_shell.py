"""Nav-shell minimal-navigation render check (feat/nav-shell-min, PART 1).

Drives a real headless Perspective session against the live 8.1.52 gateway and proves:
  1. The TOP shared dock (Shell/NavBar) RENDERS on multiple pages (domId #navbar-root present)
     — this is the headless proof that the on-disk `sharedDocks.top` entry actually mounts,
     which curl/200-checks cannot show (Perspective delivers docks over the session websocket,
     not in the bootstrap HTML).
  2. The NavBar nav links navigate via system.perspective.navigate() — clicking "Masters ▾"
     opens the flyout and a master link changes the page URL.
  3. The Order landing (/order) renders its cards and a card navigates to /order/renban.
  4. Every mounted route resolves (the page primary view renders a known root domId).

This is a NAVIGATION/render harness only — it does NOT exercise DB writes (the NavBar/landing
have none). Trial is reset first (creds from scripts/e2e/.env).

Usage:
  python3 scripts/e2e/test_nav_shell.py            # headless
  python3 scripts/e2e/test_nav_shell.py --headed   # watch it
"""
import os
import sys

from playwright.sync_api import sync_playwright

import lib
from reset_trial import reset_trial

PROJ = "InventorySystem"
BASE = lib.BASE + "/data/perspective/client/" + PROJ


def url(route):
    return BASE + "/" + route.lstrip("/")


# (route, primary-view root domId we expect to render once the page loads)
ROUTE_ROOTS = [
    ("home", "home-root"),
    ("order", "orderlanding-root"),
    ("order/renban", None),          # RenbanBreakdown (root domId unknown; assert no error + nav dock)
    ("hotcall", None),
    ("edi", "edistub-root"),
    ("reports", "reportsstub-root"),
    ("users", None),
    ("size", None),
    ("supplier", None),
    ("partsstock", None),
    ("manifestcost", None),
    ("renbangroup", None),
    ("assemblydetail", None),
    ("logistics", None),
    ("sites", None),
]


def trial_ok(page):
    try:
        return "Trial Expired" not in page.inner_text("body")
    except Exception:
        return True


def check_dock_on_pages(page, rep):
    """The shared TOP dock must render on every page. Sample a few representative routes."""
    for route in ["home", "order", "size", "users"]:
        page.goto(url(route), wait_until="networkidle", timeout=30000)
        if not trial_ok(page):
            rep.skip("NavBar dock on /%s" % route, "Perspective trial expired")
            continue
        # NavBar root + the title label both come from the docked Shell/NavBar view
        page.wait_for_selector("#navbar-root", timeout=15000)
        has_root = page.query_selector("#navbar-root") is not None
        title = page.query_selector("#navbar-brand")
        title_txt = page.inner_text("#navbar-brand") if title else ""
        rep.check("TOP shared dock (NavBar) renders on /%s" % route,
                  has_root and "Inventory System" in title_txt,
                  "navbar-root=%s title=%r" % (has_root, title_txt))


def check_nav_link(page, rep):
    """Click Masters flyout -> a master link -> URL changes (system.perspective.navigate)."""
    page.goto(url("home"), wait_until="networkidle", timeout=30000)
    if not trial_ok(page):
        rep.skip("NavBar link navigation", "Perspective trial expired")
        return
    page.wait_for_selector("#navbar-link-masters", timeout=15000)
    page.click("#navbar-link-masters")
    page.wait_for_selector("#navbar-masters-flyout", timeout=8000)
    # the flyout becomes visible; click "Size Master"
    page.get_by_text("Size Master", exact=False).first.click()
    page.wait_for_timeout(1500)
    cur = page.url
    rep.check("Masters flyout -> Size navigates", cur.rstrip("/").endswith("/size"),
              "url=%s" % cur)

    # Home link returns
    page.wait_for_selector("#navbar-link-home", timeout=10000)
    page.click("#navbar-link-home")
    page.wait_for_timeout(1500)
    rep.check("Home link navigates back", page.url.rstrip("/").endswith("/home"),
              "url=%s" % page.url)


def check_order_landing(page, rep):
    page.goto(url("order"), wait_until="networkidle", timeout=30000)
    if not trial_ok(page):
        rep.skip("Order landing cards", "Perspective trial expired")
        return
    page.wait_for_selector("#orderlanding-root", timeout=15000)
    has_renban = page.query_selector("#orderlanding-card-renban") is not None
    has_hotcall = page.query_selector("#orderlanding-card-hotcall") is not None
    rep.check("Order landing renders cards", has_renban and has_hotcall,
              "renban=%s hotcall=%s" % (has_renban, has_hotcall))
    # click the Renban Breakdown card -> /order/renban
    page.click("#orderlanding-card-renban")
    page.wait_for_timeout(1800)
    rep.check("Order landing -> Renban card navigates",
              page.url.rstrip("/").endswith("/order/renban"), "url=%s" % page.url)


def check_all_routes(page, rep):
    for route, root in ROUTE_ROOTS:
        page.goto(url(route), wait_until="networkidle", timeout=30000)
        if not trial_ok(page):
            rep.skip("route /%s renders" % route, "Perspective trial expired")
            continue
        body = page.inner_text("body")
        # A Perspective "page not found" / deserialize failure shows an error placeholder.
        bad = ("Page Not Found" in body) or ("Error loading" in body) or ("Unable to" in body)
        if root:
            try:
                page.wait_for_selector("#" + root, timeout=12000)
                ok = page.query_selector("#" + root) is not None
            except Exception:
                ok = False
            rep.check("route /%s renders (#%s)" % (route, root), ok and not bad,
                      "root_present=%s bad_text=%s" % (ok, bad))
        else:
            # no known root domId: assert the dock rendered (page mounted) + no error text
            try:
                page.wait_for_selector("#navbar-root", timeout=12000)
                mounted = True
            except Exception:
                mounted = False
            rep.check("route /%s renders (dock mounted, no error)" % route,
                      mounted and not bad, "dock=%s bad_text=%s" % (mounted, bad))


def main():
    headed = "--headed" in sys.argv
    os.makedirs(lib.ARTIFACTS, exist_ok=True)
    rep = lib.Report()

    with sync_playwright() as p:
        b = p.chromium.launch(headless=not headed, slow_mo=200 if headed else 0)
        pg = b.new_page(viewport={"width": 1680, "height": 1050})
        print("== trial reset ==")
        ok, msg = reset_trial(pg)
        if msg == "NEED_CREDS":
            rep.skip("trial reset", "no GATEWAY_USER/GATEWAY_PASS in .env")
        else:
            rep.check("trial active", ok, msg)

        print("== 1. TOP shared dock renders on pages ==")
        check_dock_on_pages(pg, rep)
        print("== 2. NavBar link navigation ==")
        check_nav_link(pg, rep)
        print("== 3. Order landing ==")
        check_order_landing(pg, rep)
        print("== 4. All mounted routes render ==")
        check_all_routes(pg, rep)

        b.close()

    return rep.summary_exit()


if __name__ == "__main__":
    sys.exit(main())
