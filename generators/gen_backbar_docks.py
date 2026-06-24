#!/usr/bin/env python3
"""Idempotent injector: add a CONTEXT-AWARE back-dock (Shell/BackBar) as a per-page TOP dock to EVERY
mounted page EXCEPT /home, in the Perspective page-config.

WHY per-page (not a shared dock): a SHARED top dock reserves its `size` (40px) on EVERY page including
/home, leaving an empty bar that violates the "Home is a clean launcher" requirement. An empty per-page
docks override does NOT suppress a shared dock (gateway ignores it), but ADDING a per-page dock works
(verified on 8.1.52). So: no shared dock; inject the back-dock per non-home page. /home gets none -> truly
clean (zero reserved space).

CONTEXT-AWARE BACK TARGET (hub-and-spoke parent): the back-dock points at the page's PARENT hub, not
always /home. The target+label travel as the dock's `viewParams` ({to, label}); the Shell/BackBar view
reads `view.params.to` (navigate) and `view.params.label` (link text). See BACK_MAP below for the
per-page parent. Pages not in BACK_MAP fall back to /home / "Home" (safe default — BackBar itself also
defaults to /home / "Home" if no viewParams arrive).

  - the 9 Master-hub children (size/supplier/partsstock/manifestcost/renbangroup/assemblydetail/
    logistics/sites/users) -> /masters  "Master Data"
  - /masters                                                                  -> /home     "Home"
  - the Order children (/order/renban, /hotcall)                              -> /order    "Order"
  - top-level module pages (/order, /reports, /coming-soon)                   -> /home     "Home"
  - /home                                                                     -> no dock

Idempotent: re-running is a no-op (matches on the dock id 'backbar' AND its viewParams). `--check`
reports drift (exit 1 if any non-home page is missing the dock, /home has one, or any dock's
viewParams != its BACK_MAP target) without writing. This is the gen_master_write_gates.py model: a
deterministic sweep, not N hand-edits, so the back-dock lands evenly across every page and the per-page
back target stays in sync as pages are added.

DEPLOY: writes BOTH the committed repo config AND the deployed gateway config (the gateway re-signs the
page-config resource on restart). `gwcmd -r` to load.

Usage:
  python3 scripts/gen_backbar_docks.py            # inject/normalize (writes repo + gateway config)
  python3 scripts/gen_backbar_docks.py --check    # report drift only, exit 1 if out of sync
"""
import json, sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _ignenv import PERSPECTIVE_DIR   # noqa: E402 — centralized gateway path (repo-split-plan §4.C)

_REPO = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
REPO_CONFIG = os.path.join(_REPO, "ignition", "perspective-views", "page-config", "config.json")
GW_CONFIG = os.path.join(PERSPECTIVE_DIR, "page-config", "config.json")

DOCK_ID = "backbar"
HOME_ROUTE = "/home"
DEFAULT_BACK = {"to": "/home", "label": "Home"}

# Per-page PARENT hub (context-aware back target). Any non-home page NOT listed here gets DEFAULT_BACK.
MASTER_CHILDREN = ["/size", "/supplier", "/partsstock", "/manifestcost", "/renbangroup",
                   "/assemblydetail", "/logistics", "/sites", "/users"]
ORDER_CHILDREN = ["/order/renban", "/hotcall"]

BACK_MAP = {}
for r in MASTER_CHILDREN:
    BACK_MAP[r] = {"to": "/masters", "label": "Master Data"}
for r in ORDER_CHILDREN:
    BACK_MAP[r] = {"to": "/order", "label": "Order"}
BACK_MAP["/masters"] = {"to": "/home", "label": "Home"}
BACK_MAP["/order"] = {"to": "/home", "label": "Home"}
BACK_MAP["/reports"] = {"to": "/home", "label": "Home"}
BACK_MAP["/coming-soon"] = {"to": "/home", "label": "Home"}


def back_target(route):
    return BACK_MAP.get(route, DEFAULT_BACK)


def backbar_dock(route):
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
        "viewParams": dict(back_target(route)),
    }


def find_backbar(page):
    docks = page.get("docks") or {}
    for d in (docks.get("top") or []):
        if d.get("id") == DOCK_ID:
            return d
    return None


def normalize(cfg, write):
    """Ensure: NO shared top dock; EVERY non-home page has the backbar top dock with the CORRECT
    context-aware viewParams; /home has NONE. Returns list of drift descriptions (empty == in sync)."""
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
            if find_backbar(page) is not None:
                drift.append("%s unexpectedly has the back-dock" % route)
                if write:
                    page["docks"]["top"] = [d for d in page["docks"]["top"] if d.get("id") != DOCK_ID]
                    if not any(page["docks"].get(k) for k in ("top", "bottom", "left", "right")):
                        page.pop("docks", None)
            continue
        # every non-home page MUST carry exactly one back-dock with the correct viewParams
        want = back_target(route)
        existing = find_backbar(page)
        if existing is None:
            drift.append("%s missing the back-dock" % route)
            if write:
                docks = page.setdefault("docks", {"top": [], "bottom": [], "left": [], "right": [],
                                                  "cornerPriority": "leftRight"})
                docks.setdefault("top", [])
                docks["top"] = [d for d in docks["top"] if d.get("id") != DOCK_ID]
                docks["top"].insert(0, backbar_dock(route))
        elif (existing.get("viewParams") or {}) != want:
            drift.append("%s back-dock target %s != expected %s"
                         % (route, existing.get("viewParams"), want))
            if write:
                existing["viewParams"] = dict(want)
    return drift


def _write_config(path, cfg):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w") as f:
        json.dump(cfg, f, indent=2)
        f.write("\n")


def main():
    check = "--check" in sys.argv
    cfg = json.load(open(REPO_CONFIG))
    drift = normalize(cfg, write=not check)
    if check:
        if drift:
            print("DRIFT (%d):" % len(drift))
            for d in drift:
                print("  -", d)
            return 1
        print("in sync: every non-home page has a context-aware back-dock; /home has none; "
              "no shared top dock")
        return 0
    # write BOTH the committed repo config AND the deployed gateway config
    _write_config(REPO_CONFIG, cfg)
    if os.path.isdir(os.path.dirname(GW_CONFIG)):
        _write_config(GW_CONFIG, cfg)
        deployed = "repo + gateway"
    else:
        deployed = "repo only (gateway path absent)"
    npages = sum(1 for r in cfg["pages"] if r != HOME_ROUTE)
    print("normalized: context-aware back-dock on %d non-home pages; /home clean; "
          "shared top dock empty -> wrote %s" % (npages, deployed))
    if drift:
        print("(applied %d fix(es))" % len(drift))
    return 0


if __name__ == "__main__":
    sys.exit(main())
