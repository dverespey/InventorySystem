#!/usr/bin/env python3
"""Idempotent injector: add the minimal '<- Home' back-dock (Shell/BackBar) as a per-page TOP dock
to EVERY mounted page EXCEPT /home, in the Perspective page-config.

WHY per-page (not a shared dock): a SHARED top dock reserves its `size` (40px) on EVERY page including
/home, leaving an empty bar that violates the "Home is a clean launcher" requirement. An empty per-page
docks override does NOT suppress a shared dock (gateway ignores it), but ADDING a per-page dock works
(verified on 8.1.52). So: no shared dock; inject the back-dock per non-home page. /home gets none -> truly
clean (zero reserved space).

Idempotent: re-running is a no-op (matches on the dock id 'backbar'). `--check` reports drift (exit 1 if
any non-home page is missing the dock or /home has one) without writing. This is the
gen_master_write_gates.py model: a deterministic sweep, not N hand-edits, so the back-dock lands evenly
across every page and stays in sync as pages are added.

Usage:
  python3 scripts/gen_backbar_docks.py            # inject/normalize (writes docs config)
  python3 scripts/gen_backbar_docks.py --check    # report drift only, exit 1 if out of sync
"""
import json, sys, os

CONFIG = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                      "..", "docs", "design", "perspective-views", "page-config", "config.json")
CONFIG = os.path.abspath(CONFIG)

DOCK_ID = "backbar"
HOME_ROUTE = "/home"


def backbar_dock():
    return {
        "id": DOCK_ID,
        "viewPath": "Shell/BackBar",
        "size": 40,
        "anchor": "fixed",
        "content": "push",
        "show": "visible",
        "modal": False,
        "resizable": False,
        "handle": "hide",
        "autoBreakpoint": 0,
        "viewParams": {},
    }


def page_has_backbar(page):
    docks = page.get("docks") or {}
    return any(d.get("id") == DOCK_ID for d in (docks.get("top") or []))


def normalize(cfg, write):
    """Ensure: NO shared top dock; EVERY non-home page has the backbar top dock; /home has NONE.
    Returns list of drift descriptions (empty == in sync)."""
    drift = []

    # 1. shared top dock must be empty (we use per-page docks)
    shared_top = cfg.setdefault("sharedDocks", {}).setdefault("top", [])
    if shared_top:
        drift.append("sharedDocks.top is non-empty (should be [] — back-dock is per-page)")
        if write:
            cfg["sharedDocks"]["top"] = []

    pages = cfg.get("pages", {})
    for route, page in pages.items():
        if route == HOME_ROUTE:
            # /home must NOT carry the back-dock
            if page_has_backbar(page):
                drift.append("%s unexpectedly has the back-dock" % route)
                if write:
                    page["docks"]["top"] = [d for d in page["docks"]["top"] if d.get("id") != DOCK_ID]
                    if not any(page["docks"].get(k) for k in ("top", "bottom", "left", "right")):
                        page.pop("docks", None)
            continue
        # every non-home page MUST carry exactly one back-dock
        if not page_has_backbar(page):
            drift.append("%s missing the back-dock" % route)
            if write:
                docks = page.setdefault("docks", {"top": [], "bottom": [], "left": [], "right": [],
                                                  "cornerPriority": "leftRight"})
                docks.setdefault("top", [])
                docks["top"] = [d for d in docks["top"] if d.get("id") != DOCK_ID]
                docks["top"].insert(0, backbar_dock())
    return drift


def main():
    check = "--check" in sys.argv
    cfg = json.load(open(CONFIG))
    drift = normalize(cfg, write=not check)
    if check:
        if drift:
            print("DRIFT (%d):" % len(drift))
            for d in drift:
                print("  -", d)
            return 1
        print("in sync: every non-home page has the back-dock; /home has none; no shared top dock")
        return 0
    # write
    json.dump(cfg, open(CONFIG, "w"), indent=2)
    open(CONFIG, "a").write("\n")
    npages = sum(1 for r in cfg["pages"] if r != HOME_ROUTE)
    print("normalized: back-dock on %d non-home pages; /home clean; shared top dock empty" % npages)
    if drift:
        print("(applied %d fix(es))" % len(drift))
    return 0


if __name__ == "__main__":
    sys.exit(main())
