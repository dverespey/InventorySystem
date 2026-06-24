"""Nav-shell render check — HUB-AND-SPOKE model with CONTEXT-AWARE back + valid master icons.

David's PART-1 top nav bar (Shell/NavBar shared dock) is REPLACED by a home-as-launcher model:
  * The Home page (/home) IS the launcher — its 8 MODULES cards navigate. There is NO top nav bar.
  * Every NON-home page carries a single minimal back-dock (Shell/BackBar) as a per-page TOP dock; /home
    gets NONE (clean launcher, zero reserved space). The back target is CONTEXT-AWARE (hub-and-spoke
    parent), not always /home: the dock's viewParams {to,label} are read by Shell/BackBar as
    view.params.to (navigate) / view.params.label (link text). Dock viewParams ARE delivered on 8.1.52
    headless (verified) — this is the mechanism (a single param-aware BackBar, NOT per-target variants).
    Per-page parent map (BACK_MAP below): the 9 Master-hub children -> /masters "Master Data";
    /masters,/order,/reports,/coming-soon -> /home "Home"; /order/renban,/hotcall -> /order "Order".
  * /masters is a sub-hub (Master/MasterHub) of 9 cards (8 masters + User Admin), reachable from the
    Home "Master Data" card — this is how /users + /sites stay reachable now the top nav is gone.
  * The 5 not-yet-built modules (Shipping/Receiving/Stocktaking/Forecast/Invoicing) point at a generic
    /coming-soon (Shell/ComingSoon).

This harness drives a real headless Perspective session against the live 8.1.52 gateway and proves the
model end to end. Headless render-level proof is required because the Perspective client is a SPA: curl
returns 200 for ANY path (including a removed mount), so only a rendered session distinguishes a mounted
route from a "Page Not Found" placeholder, and only a session shows whether a dock mounts (docks arrive
over the session websocket, not the bootstrap HTML).

Assertions (the hub-and-spoke contract):
  1. HOME IS CLEAN: /home renders (#home-root) and shows NEITHER the old #navbar-root NOR the new
     #backbar-root dock, and reserves NO top-dock space (content starts at y~0).
  2. CONTEXT-AWARE BACK: each non-home page's back-dock shows its PARENT's LABEL and clicking it
     navigates to that PARENT (per BACK_MAP) — NOT always /home. Spot-checks every distinct target:
     /size->/masters "Master Data", /masters->/home "Home", /order/renban->/order "Order",
     /reports->/home "Home", /hotcall->/order "Order".
  3. THE 8 HOME CARDS NAVIGATE: each #home-card-* click lands on its route.
  4. /masters renders its 9 sub-hub cards and a card (e.g. Sites) navigates to its master.
  4b. THE 4 FIXED MASTER ICONS RENDER: supplier/partsstock/assemblydetail/logistics (+ the swept users
     icon) draw a real glyph — asserted by the svg having >=1 inlined glyph child. NON-VACUOUS: a broken
     icon path (the original material/factory etc., absent from the 8.1.52 sprite) renders an svg with
     ZERO glyph children (verified by reverting one icon on the box: factory -> 0 children, business -> 3).
  5. /coming-soon renders (#comingsoon-root).
  6. THE REMOVED /edi MOUNT NO LONGER RESOLVES (renders Page-Not-Found, not a view).
  7. Every still-mounted route renders its primary view (no deserialize/blank).

NAVIGATION/render only — no DB writes. Trial is reset first (creds from e2e/.env).

Usage:
  python3 e2e/test_nav_shell.py            # headless
  python3 e2e/test_nav_shell.py --headed   # watch it
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


def cur_route(page):
    """Last path segment(s) of the current client URL, normalized to a leading-slash route."""
    u = page.url.split("?")[0].rstrip("/")
    marker = "/client/" + PROJ
    if marker in u:
        tail = u.split(marker, 1)[1]
        return tail if tail else "/"
    return u


def trial_ok(page):
    try:
        return "Trial Expired" not in page.inner_text("body")
    except Exception:
        return True


def page_not_found(page):
    """True when a route is NOT mounted. Perspective renders an empty-page placeholder for an
    unmounted path: 'View Not Found' / 'No view configured for this page' (it does NOT say
    'Page Not Found'). Match those + any legacy strings."""
    try:
        body = page.inner_text("body")
    except Exception:
        return False
    low = body.lower()
    return ("view not found" in low) or ("no view configured" in low) or ("page not found" in low)


# (route, primary-view root domId expected once the page loads). None => assert it mounted via
# the back-dock (#backbar-root) + no error text, since the root domId isn't separately asserted here.
ROUTE_ROOTS = [
    ("home", "home-root"),
    ("order", "orderlanding-root"),
    ("order/renban", None),
    ("hotcall", None),
    ("reports", "reportsstub-root"),
    ("masters", "masterhub-root"),
    ("coming-soon", "comingsoon-root"),
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

# Home MODULES card domId -> route it must navigate to.
HOME_CARDS = [
    ("home-card-order", "/order"),
    ("home-card-masters", "/masters"),
    ("home-card-reports", "/reports"),
    ("home-card-shipping", "/coming-soon"),
    ("home-card-receiving", "/coming-soon"),
    ("home-card-stocktaking", "/coming-soon"),
    ("home-card-forecast", "/coming-soon"),
    ("home-card-invoicing", "/coming-soon"),
]

# CONTEXT-AWARE back map: page route -> (expected back LABEL, expected back TARGET). Mirrors
# scripts/gen_backbar_docks.py BACK_MAP. The spot-checks below cover every DISTINCT target+label pair
# (so a wrong injector map or a viewParams-delivery regression fails here). /home has NO back-dock.
BACK_MAP = {
    "size": ("Master Data", "/masters"),
    "supplier": ("Master Data", "/masters"),
    "partsstock": ("Master Data", "/masters"),
    "manifestcost": ("Master Data", "/masters"),
    "renbangroup": ("Master Data", "/masters"),
    "assemblydetail": ("Master Data", "/masters"),
    "logistics": ("Master Data", "/masters"),
    "sites": ("Master Data", "/masters"),
    "users": ("Master Data", "/masters"),
    "masters": ("Home", "/home"),
    "order": ("Home", "/home"),
    "reports": ("Home", "/home"),
    "coming-soon": ("Home", "/home"),
    "order/renban": ("Order", "/order"),
    "hotcall": ("Order", "/order"),
}

# The 4 master icons this fix repaired (were broken/absent in the 8.1.52 material sprite) + the swept
# users icon. Each card's svg must inline >=1 glyph child to count as RENDERED (non-vacuous: a broken
# path renders an svg with 0 glyph children — proven on the box).
FIXED_ICON_CARDS = [
    "masterhub-card-supplier",
    "masterhub-card-partsstock",
    "masterhub-card-assemblydetail",
    "masterhub-card-logistics",
    "masterhub-card-users",
]
GLYPH_CHILD_SEL = "path,circle,rect,polygon,polyline,g"


def goto_home(page):
    page.goto(url("home"), wait_until="networkidle", timeout=30000)
    page.wait_for_selector("#home-root", timeout=15000)


def check_home_is_clean(page, rep):
    """1. /home renders and shows NO top dock — and reserves NO top dock SPACE.

    The back-dock is per-page and /home is simply NOT given one (its page-config entry has no docks),
    so no BackBar mounts on home and no dock strip is reserved. 'Clean' means two things: (a) neither
    the removed #navbar-root nor #backbar-root render, AND (b) the page content (#home-root) starts at
    the very top (y ~ 0) — i.e. no 40px dock strip is reserved. (b) is the load-bearing check: a merely
    invisible-but-space-reserving dock would still leave an empty bar."""
    page.goto(url("home"), wait_until="networkidle", timeout=30000)
    if not trial_ok(page):
        rep.skip("Home is nav-bar-free", "Perspective trial expired")
        return
    page.wait_for_selector("#home-root", timeout=15000)
    page.wait_for_timeout(1200)  # let docks settle
    has_navbar = page.query_selector("#navbar-root") is not None
    has_backbar = page.query_selector("#backbar-root") is not None
    rep.check("HOME is clean: no top NavBar dock (#navbar-root absent)", not has_navbar,
              "navbar-root present=%s" % has_navbar)
    rep.check("HOME is clean: back-dock absent on /home (#backbar-root absent)", not has_backbar,
              "backbar-root present=%s" % has_backbar)
    # No reserved dock space: #home-root must start at the top of the viewport.
    hr = page.query_selector("#home-root")
    top_y = hr.bounding_box()["y"] if hr else None
    rep.check("HOME is clean: NO reserved top-dock space (content starts at y~0)",
              top_y is not None and top_y <= 2, "home-root top y=%s" % top_y)


def check_context_aware_back(page, rep):
    """2. Each non-home page's back-dock shows its PARENT's LABEL and navigates to that PARENT — the
    CONTEXT-AWARE hub-and-spoke contract, NOT always /home. We spot-check every distinct target+label
    pair in BACK_MAP (one Master child, one Order child, and the three /home-parented top pages):
      /size          -> "Master Data" -> /masters   (Master child)
      /masters       -> "Home"        -> /home      (sub-hub -> home)
      /order/renban  -> "Order"       -> /order     (Order child)
      /reports       -> "Home"        -> /home      (top module -> home)
      /hotcall       -> "Order"       -> /order     (Order child)
    For each: assert the back-dock renders, its label text == expected (proves view.params.label
    delivered), then click and assert it lands on the expected PARENT (proves view.params.to delivered).
    A wrong injector map, an undelivered viewParam (label/target falls back to "Home"/"/home"), or a
    regressed BackBar all fail here."""
    spot = ["size", "masters", "order/renban", "reports", "hotcall"]
    for route in spot:
        want_label, want_to = BACK_MAP[route]
        page.goto(url(route), wait_until="networkidle", timeout=30000)
        if not trial_ok(page):
            rep.skip("context-aware back on /%s" % route, "Perspective trial expired")
            continue
        # back-dock present
        try:
            page.wait_for_selector("#backbar-home", timeout=15000)
            has = page.query_selector("#backbar-root") is not None
        except Exception:
            has = False
        if not has:
            rep.check("/%s shows the back-dock (#backbar-root)" % route, False, "absent")
            continue
        page.wait_for_timeout(600)
        # LABEL == parent label (view.params.label delivered)
        el = page.query_selector("#backbar-home")
        label = (el.inner_text() or "").strip().replace("\n", " ") if el else ""
        rep.check("/%s back-link LABEL is '%s' (parent hub, not always Home)" % (route, want_label),
                  want_label in label, "label=%r want=%r" % (label, want_label))
        # CLICK navigates to parent TARGET (view.params.to delivered)
        try:
            page.click("#backbar-home")
            page.wait_for_timeout(1500)
            landed = cur_route(page)
            nav_ok = landed.rstrip("/").endswith(want_to)
        except Exception as e:
            landed = "click-fail:%s" % e
            nav_ok = False
        rep.check("/%s back-link navigates to PARENT %s" % (route, want_to), nav_ok,
                  "landed=%s want=%s" % (landed, want_to))


def check_home_cards_navigate(page, rep):
    """3. Each of the 8 MODULES cards navigates to its route."""
    for dom_id, route in HOME_CARDS:
        goto_home(page)
        if not trial_ok(page):
            rep.skip("Home card %s -> %s" % (dom_id, route), "Perspective trial expired")
            continue
        sel = "#" + dom_id
        try:
            page.wait_for_selector(sel, timeout=15000)
            page.click(sel)
            page.wait_for_timeout(1500)
            landed = cur_route(page)
            ok = landed.rstrip("/").endswith(route)
        except Exception as e:
            ok = False
            landed = "click-failed: %s" % e
        rep.check("Home card %-22s navigates to %s" % (dom_id, route), ok, "landed=%s" % landed)


def check_masters_hub(page, rep):
    """4. /masters renders 9 cards; a card (Sites) navigates to its master."""
    page.goto(url("masters"), wait_until="networkidle", timeout=30000)
    if not trial_ok(page):
        rep.skip("/masters sub-hub", "Perspective trial expired")
        return
    try:
        page.wait_for_selector("#masterhub-root", timeout=15000)
    except Exception:
        rep.check("/masters renders (#masterhub-root)", False, "root not found")
        return
    # 9 cards: 7 reference-data + Sites + Users
    card_ids = ["masterhub-card-size", "masterhub-card-supplier", "masterhub-card-partsstock",
                "masterhub-card-manifestcost", "masterhub-card-renbangroup",
                "masterhub-card-assemblydetail", "masterhub-card-logistics",
                "masterhub-card-sites", "masterhub-card-users"]
    present = [cid for cid in card_ids if page.query_selector("#" + cid) is not None]
    rep.check("/masters renders all 9 sub-hub cards (8 masters + User Admin)",
              len(present) == 9, "present=%d/9 missing=%s"
              % (len(present), [c for c in card_ids if c not in present]))
    # Sites card -> /sites (proves /sites reachable now the top nav is gone)
    try:
        page.click("#masterhub-card-sites")
        page.wait_for_timeout(1500)
        ok = cur_route(page).rstrip("/").endswith("/sites")
    except Exception:
        ok = False
    rep.check("/masters Sites card navigates to /sites", ok, "route=%s" % cur_route(page))
    # Users card -> /users
    page.goto(url("masters"), wait_until="networkidle", timeout=30000)
    try:
        page.wait_for_selector("#masterhub-card-users", timeout=15000)
        page.click("#masterhub-card-users")
        page.wait_for_timeout(1500)
        ok = cur_route(page).rstrip("/").endswith("/users")
    except Exception:
        ok = False
    rep.check("/masters User Admin card navigates to /users", ok, "route=%s" % cur_route(page))


def check_master_icons_render(page, rep):
    """4b. The 4 fixed master icons (supplier/partsstock/assemblydetail/logistics) + the swept users icon
    actually RENDER on /masters. A Perspective ia.display.icon renders as an inline <svg data-icon=...>
    whose glyph geometry is INLINED as child <path>/<g>/... elements when the icon path resolves in the
    gateway's material sprite, and as an EMPTY svg (zero glyph children) when the path is absent/broken
    (the original material/factory, material/inventory_2, material/precision_manufacturing, material/route
    were all absent from the 8.1.52 sprite). So 'renders' == svg present with >=1 glyph child. This
    discriminator is non-vacuous: verified on the box that the broken material/factory yields 0 glyph
    children while material/business yields 3."""
    page.goto(url("masters"), wait_until="networkidle", timeout=30000)
    if not trial_ok(page):
        rep.skip("master icons render", "Perspective trial expired")
        return
    try:
        page.wait_for_selector("#masterhub-root", timeout=15000)
    except Exception:
        rep.check("/masters renders for icon check", False, "root not found")
        return
    page.wait_for_timeout(1000)
    for card in FIXED_ICON_CARDS:
        sel = "#" + card + " svg[data-icon]"
        svg = page.query_selector(sel)
        if svg is None:
            rep.check("icon RENDERS on %s (svg present + glyph)" % card, False, "no svg[data-icon]")
            continue
        data_icon = svg.get_attribute("data-icon") or "?"
        nchild = page.eval_on_selector(
            sel, "el => el.querySelectorAll('%s').length" % GLYPH_CHILD_SEL)
        rep.check("icon RENDERS on %-30s (%s, %s glyph children > 0)"
                  % (card, data_icon, nchild), nchild and nchild > 0,
                  "data-icon=%s glyph-children=%s" % (data_icon, nchild))


def check_coming_soon(page, rep):
    """5. /coming-soon renders the generic message."""
    page.goto(url("coming-soon"), wait_until="networkidle", timeout=30000)
    if not trial_ok(page):
        rep.skip("/coming-soon renders", "Perspective trial expired")
        return
    try:
        page.wait_for_selector("#comingsoon-root", timeout=15000)
        body = page.inner_text("#comingsoon-root")
        ok = "development" in body.lower()
    except Exception:
        ok = False
        body = ""
    rep.check("/coming-soon renders the in-development message", ok, "text=%r" % body[:60])


def check_edi_removed(page, rep):
    """6. The removed /edi mount no longer resolves (Page Not Found, not a view)."""
    page.goto(url("edi"), wait_until="networkidle", timeout=30000)
    if not trial_ok(page):
        rep.skip("/edi mount removed", "Perspective trial expired")
        return
    page.wait_for_timeout(1200)
    pnf = page_not_found(page)
    # also assert the old EDI stub root is NOT present
    has_stub = page.query_selector("#edistub-root") is not None
    rep.check("removed /edi route no longer resolves (Page Not Found, stub gone)",
              pnf and not has_stub, "page_not_found=%s edistub_present=%s" % (pnf, has_stub))


def check_all_routes(page, rep):
    """7. Every still-mounted route renders its primary view."""
    for route, root in ROUTE_ROOTS:
        page.goto(url(route), wait_until="networkidle", timeout=30000)
        if not trial_ok(page):
            rep.skip("route /%s renders" % route, "Perspective trial expired")
            continue
        if page_not_found(page):
            rep.check("route /%s renders (mounted)" % route, False, "Page Not Found")
            continue
        if root:
            try:
                page.wait_for_selector("#" + root, timeout=12000)
                ok = page.query_selector("#" + root) is not None
            except Exception:
                ok = False
            rep.check("route /%s renders (#%s)" % (route, root), ok, "root_present=%s" % ok)
        else:
            # no known primary-view domId: assert the back-dock mounted (page is a real non-home page)
            try:
                page.wait_for_selector("#backbar-root", timeout=12000)
                mounted = True
            except Exception:
                mounted = False
            rep.check("route /%s renders (back-dock mounted, no error)" % route, mounted,
                      "backdock=%s" % mounted)


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

        print("== 1. HOME is clean (no top nav / no back-dock) ==")
        check_home_is_clean(pg, rep)
        print("== 2. CONTEXT-AWARE back: each page -> its PARENT hub (label + nav) ==")
        check_context_aware_back(pg, rep)
        print("== 3. the 8 Home MODULES cards navigate ==")
        check_home_cards_navigate(pg, rep)
        print("== 4. /masters sub-hub (9 cards) ==")
        check_masters_hub(pg, rep)
        print("== 4b. the 4 fixed master icons RENDER (+ swept users icon) ==")
        check_master_icons_render(pg, rep)
        print("== 5. /coming-soon ==")
        check_coming_soon(pg, rep)
        print("== 6. removed /edi mount ==")
        check_edi_removed(pg, rep)
        print("== 7. all mounted routes render ==")
        check_all_routes(pg, rep)

        # artifact for the human: a clean Home + a module page with the back-dock
        try:
            pg.goto(url("home"), wait_until="networkidle", timeout=20000)
            pg.wait_for_timeout(800)
            pg.screenshot(path=lib.ARTIFACTS + "/nav_home_clean.png", full_page=True)
            pg.goto(url("masters"), wait_until="networkidle", timeout=20000)
            pg.wait_for_timeout(800)
            pg.screenshot(path=lib.ARTIFACTS + "/nav_masters_hub.png", full_page=True)
        except Exception:
            pass

        b.close()

    return rep.summary_exit()


if __name__ == "__main__":
    sys.exit(main())
