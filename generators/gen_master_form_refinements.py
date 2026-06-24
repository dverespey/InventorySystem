#!/usr/bin/env python3
"""gen_master_form_refinements.py — David's master-CRUD form-pattern refinement sweep across ALL 8 master
views (Size, Supplier, PartsStock, ManifestCost, RenbanGroup, AssemblyDetail, Logistics, Sites),
deterministically + idempotently (the gen_master_write_gates.py model: re-run = no-op; --check reports drift).

THE 4 CHANGES (applied identically to all 8; the views are structurally uniform — verified
FilterBar=[SearchWrap, SearchButton, NewButton], ActionBar=[SaveButton, DeleteButton, ClearButton]):

  1. SEARCH BOX -> SINGLE-LINE. The SearchField (ia.input.text-field) had NO explicit height and sits in a
     flex column inside a flex row (FilterBar) whose cross-axis stretch grew it tall. Fix = pin an explicit
     single-line height ("32px") on SearchField props.style. Deterministic single-line regardless of the
     flex stretch.

  2. REMOVE THE NEW BUTTON. NewButton is redundant with Clear (Clear blanks the form -> enter data -> Save
     INSERTs). Delete the NewButton component (and its onActionPerformed reset script) from FilterBar. After
     this the top has ONLY the search; the three action buttons live together in the detail ActionBar.

  3. GROUP SAVE/DELETE/CLEAR IN THE DETAIL SECTION. Already satisfied structurally: Save/Delete/Clear are
     the ActionBar in the RightPane (detail). New was the only action in the top FilterBar; removing it
     (change 2) leaves the three grouped together in detail. This script ASSERTS that invariant (the action
     buttons are in ActionBar and NOT in FilterBar) and fails loudly if a view diverges.

  4. ENABLE/DISABLE THE ACTION BUTTONS (props.enabled expression bindings):
       * Clear  : ALWAYS enabled  -> NO enabled binding (Perspective default enabled=true). Idempotently
                  STRIP any enabled binding off Clear.
       * Delete : enabled ONLY when a record is selected -> enabled = {view.custom.recordId} > 0. Disabled on
                  a blank/cleared form (recordId==0).
       * Save   : enabled when a record is selected OR data has been entered into a cleared form ->
                  enabled = {view.custom.recordId} > 0 || {view.custom.form_<PRIMARY>} != ''. The
                  "has-entered-data" signal is the view's PRIMARY/required input prop being non-blank
                  (PRIMARY_FIELD below) — the SAME field each Save script already REQUIRES (so Save lights up
                  exactly when the required key has been typed). Chosen over a view.custom.dirty flag because
                  it needs NO per-input onChange wiring and NO reset bookkeeping: Clear already blanks the
                  primary -> Save goes disabled; loading a row sets recordId>0 -> Save enabled. Every view's
                  primary blanks to '' (text fields and the ManifestCost assy-code dropdown alike), so the
                  != '' test is uniform across all 8.

DO-NOT-BREAK INVARIANT (P15 server-side write gate): this script touches ONLY button presence/grouping/
enabled-state + the search height. It does NOT edit the Save/Delete onActionPerformed SCRIPTS — the
auth.requireWrite(self.session) gate is untouched. (Idempotency/guards below also refuse to run if a Save/
Delete script is missing.) Re-run test_master_write_gates after = still 86/0.

PROVENANCE (same as gen_master_write_gates): writes BOTH copies so committed == deployed.
  * REPO_BASE = the committed source-of-truth view.json (reviewable/redeployable).
  * GW_BASE   = the DEPLOYED gateway view.json (what the running session serves).
  The committed resource.json's gateway signature is STRIPPED (stale after a view edit; the deployed
  gateway resource.json keeps its attributes and the gateway re-signs on restart — NOT touched here).

Run:  python3 scripts/gen_master_form_refinements.py            # apply to all 8 (repo + gateway)
      python3 scripts/gen_master_form_refinements.py --check    # report only; non-zero exit on any drift
"""
import json
import os
import sys

# All 8 master-CRUD views (each Master/<V>/<V>/view.json). Unlike the P15 write-gate sweep (7 views; Sites
# already carried the gate from its own generator), this refinement applies to ALL 8 including Sites.
VIEWS = ["Size", "Supplier", "PartsStock", "ManifestCost", "RenbanGroup", "AssemblyDetail", "Logistics",
         "Sites"]

# The view's PRIMARY/required input prop — the "has-entered-data" signal for the Save-enable expression.
# This is the SAME field each view's Save script REQUIRES (read from the deployed validation guards), so
# Save lights up exactly when the required key has been typed into a cleared form. Every one blanks to ''.
PRIMARY_FIELD = {
    "Size":           "form_code",
    "Supplier":       "form_code",
    "PartsStock":     "form_partNumber",
    "ManifestCost":   "form_assyCode",   # a dropdown (props.value); its blank value is '' (Clear sets "")
    "RenbanGroup":    "form_code",
    "AssemblyDetail": "form_assyCode",
    "Logistics":      "form_name",
    "Sites":          "form_name",
}

_REPO = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _ignenv import PERSPECTIVE_DIR   # noqa: E402 — centralized gateway path (repo-split-plan §4.C)
REPO_BASE = os.path.join(_REPO, "ignition", "perspective-views", "Master")
GW_BASE = os.path.join(PERSPECTIVE_DIR, "views", "Master")

SEARCH_FIELD = "SearchField"
NEW_BUTTON = "NewButton"
FILTER_BAR = "FilterBar"
ACTION_BAR = "ActionBar"
SAVE_BUTTON = "SaveButton"
DELETE_BUTTON = "DeleteButton"
CLEAR_BUTTON = "ClearButton"

# Single-line input height (px). Perspective ia.input.text-field renders one line at this height; pinning it
# overrides the flex cross-axis stretch that made the search box render tall.
SEARCH_HEIGHT = "32px"


# ---------------------------------------------------------------------------
# tree helpers
# ---------------------------------------------------------------------------

def _find_node(node, name):
    """Depth-first find the component dict whose meta.name == name, else None."""
    if isinstance(node, dict):
        if node.get("meta", {}).get("name") == name:
            return node
        for v in node.values():
            r = _find_node(v, name)
            if r is not None:
                return r
    elif isinstance(node, list):
        for it in node:
            r = _find_node(it, name)
            if r is not None:
                return r
    return None


def _find_parent_with_child(node, child_name):
    """Find the dict node that has a `children` list containing a component named child_name. Returns the
    parent dict (so the caller can mutate parent['children'])."""
    if isinstance(node, dict):
        for ch in node.get("children", []) or []:
            if isinstance(ch, dict) and ch.get("meta", {}).get("name") == child_name:
                return node
        for v in node.values():
            r = _find_parent_with_child(v, child_name)
            if r is not None:
                return r
    elif isinstance(node, list):
        for it in node:
            r = _find_parent_with_child(it, child_name)
            if r is not None:
                return r
    return None


def _save_enable_expr(view_name):
    """enabled = recordId selected OR the primary required field has been typed (has-entered-data)."""
    return "{view.custom.recordId} > 0 || {view.custom.%s} != ''" % PRIMARY_FIELD[view_name]


DELETE_ENABLE_EXPR = "{view.custom.recordId} > 0"


def _enabled_binding(expr):
    return {"binding": {"config": {"expression": expr}, "type": "expr"}}


# ---------------------------------------------------------------------------
# the 4 mutations (each idempotent; returns True if it CHANGED the view)
# ---------------------------------------------------------------------------

def _apply_search_height(view, view_name):
    """Make the SearchField render single-line.

    ROOT CAUSE (verified on the live gateway): SearchWrap is a COLUMN flex, so the SearchField's
    position.basis ("220px"/"240px" — intended as the input WIDTH) is applied to the flex MAIN axis =
    HEIGHT, forcing the input 220/240px TALL. A style.height alone is overridden by that flex-basis.

    FIX: relocate the intended size to the CROSS axis (width) and free the main axis (height):
      * style.width  = the px the basis used to carry (so the search box keeps its width)
      * position.basis = "auto"  (height now follows content/height-style, not the 220px basis)
      * style.height = "32px"    (an explicit single-line height; a definite width also stops the
                                  column's align-items:stretch from re-stretching it)
    Idempotent: a no-op once basis=="auto" and style has the height+width already."""
    f = _find_node(view.get("root", {}), SEARCH_FIELD)
    if f is None:
        raise ValueError("%s: %s not found (irregular structure)" % (view_name, SEARCH_FIELD))
    if f.get("type") != "ia.input.text-field":
        raise ValueError("%s: %s is %r, expected ia.input.text-field" % (view_name, SEARCH_FIELD, f.get("type")))
    props = f.setdefault("props", {})
    style = props.setdefault("style", {})
    pos = f.setdefault("position", {})

    # The width to preserve: whatever the basis carried (px) before; default 220px if it was unset/auto.
    prior_basis = pos.get("basis")
    if isinstance(prior_basis, str) and prior_basis.endswith("px"):
        width = prior_basis
    else:
        width = style.get("width", "220px")

    desired = (style.get("height") == SEARCH_HEIGHT and style.get("width") == width
               and pos.get("basis") == "auto")
    if desired:
        return False
    style["height"] = SEARCH_HEIGHT
    style["width"] = width
    pos["basis"] = "auto"
    return True


def _remove_new_button(view, view_name):
    parent = _find_parent_with_child(view.get("root", {}), NEW_BUTTON)
    if parent is None:
        # already removed (idempotent) — but confirm it isn't lurking elsewhere
        if _find_node(view.get("root", {}), NEW_BUTTON) is not None:
            raise ValueError("%s: %s present but not a direct child of any container" % (view_name, NEW_BUTTON))
        return False
    # Defensive: New must be in the top FilterBar, never the detail ActionBar.
    if parent.get("meta", {}).get("name") != FILTER_BAR:
        raise ValueError("%s: %s parent is %r, expected %s (irregular)"
                         % (view_name, NEW_BUTTON, parent.get("meta", {}).get("name"), FILTER_BAR))
    parent["children"] = [c for c in parent["children"]
                          if not (isinstance(c, dict) and c.get("meta", {}).get("name") == NEW_BUTTON)]
    return True


def _assert_buttons_grouped_in_detail(view, view_name):
    """Change 3 invariant: Save/Delete/Clear live together in the detail ActionBar and NONE leaks into the
    top FilterBar. Raises if violated (so a structural divergence STOPS the sweep rather than shipping)."""
    action = _find_node(view.get("root", {}), ACTION_BAR)
    if action is None:
        raise ValueError("%s: %s not found" % (view_name, ACTION_BAR))
    names = [c.get("meta", {}).get("name") for c in (action.get("children", []) or []) if isinstance(c, dict)]
    for b in (SAVE_BUTTON, DELETE_BUTTON, CLEAR_BUTTON):
        if b not in names:
            raise ValueError("%s: %s not grouped in %s (found %s)" % (view_name, b, ACTION_BAR, names))
    fb = _find_node(view.get("root", {}), FILTER_BAR)
    fb_names = [c.get("meta", {}).get("name") for c in (fb.get("children", []) or []) if isinstance(c, dict)]
    for b in (SAVE_BUTTON, DELETE_BUTTON, CLEAR_BUTTON):
        if b in fb_names:
            raise ValueError("%s: action button %s leaked into the top %s" % (view_name, b, FILTER_BAR))


def _set_enabled(view, view_name, button_name, expr):
    """Set props.enabled to the given expression binding (idempotent). Returns True if changed."""
    btn = _find_node(view.get("root", {}), button_name)
    if btn is None:
        raise ValueError("%s: %s not found" % (view_name, button_name))
    pc = btn.setdefault("propConfig", {})
    want = _enabled_binding(expr)
    if pc.get("props.enabled") == want:
        return False
    pc["props.enabled"] = want
    return True


def _clear_enabled(view, view_name, button_name):
    """Ensure the button has NO props.enabled binding (always enabled). Returns True if it removed one."""
    btn = _find_node(view.get("root", {}), button_name)
    if btn is None:
        raise ValueError("%s: %s not found" % (view_name, button_name))
    pc = btn.get("propConfig", {})
    if "props.enabled" in pc:
        del pc["props.enabled"]
        if not pc:
            btn.pop("propConfig", None)
        return True
    return False


def process_view(view, view_name):
    """Apply all 4 changes to a parsed view.json dict in place. Returns a dict of what changed."""
    changed = {}
    changed["search_height"] = _apply_search_height(view, view_name)
    changed["new_removed"] = _remove_new_button(view, view_name)
    _assert_buttons_grouped_in_detail(view, view_name)        # change 3 invariant (raises on violation)
    changed["save_enabled"] = _set_enabled(view, view_name, SAVE_BUTTON, _save_enable_expr(view_name))
    changed["delete_enabled"] = _set_enabled(view, view_name, DELETE_BUTTON, DELETE_ENABLE_EXPR)
    changed["clear_always_enabled"] = _clear_enabled(view, view_name, CLEAR_BUTTON)
    return changed


# ---------------------------------------------------------------------------
# check (report-only) — true if the view ALREADY matches the desired end-state
# ---------------------------------------------------------------------------

def check_view(view, view_name):
    """Return (ok, states) — ok is True iff all 4 changes are already in place + the P15 gate is intact."""
    states = []
    ok = True

    f = _find_node(view.get("root", {}), SEARCH_FIELD)
    style = (f.get("props", {}).get("style", {}) if f else {})
    pos = (f.get("position", {}) if f else {})
    # single-line == explicit single-line height AND the column-flex basis no longer forces a tall height
    # (basis must be "auto", not a px value, since SearchWrap is a column where basis = height).
    single_line = (bool(f) and style.get("height") == SEARCH_HEIGHT and pos.get("basis") == "auto")
    states.append("search=%s" % ("SINGLE-LINE" if single_line else "UNSIZED"))
    ok = ok and single_line

    new_gone = _find_node(view.get("root", {}), NEW_BUTTON) is None
    states.append("New=%s" % ("REMOVED" if new_gone else "PRESENT"))
    ok = ok and new_gone

    try:
        _assert_buttons_grouped_in_detail(view, view_name)
        grouped = True
    except ValueError:
        grouped = False
    states.append("grouped=%s" % ("YES" if grouped else "NO"))
    ok = ok and grouped

    save = _find_node(view.get("root", {}), SAVE_BUTTON)
    save_enb = (save.get("propConfig", {}).get("props.enabled") if save else None)
    save_ok = save_enb == _enabled_binding(_save_enable_expr(view_name))
    states.append("Save.enabled=%s" % ("OK" if save_ok else "MISSING"))
    ok = ok and save_ok

    dele = _find_node(view.get("root", {}), DELETE_BUTTON)
    del_enb = (dele.get("propConfig", {}).get("props.enabled") if dele else None)
    del_ok = del_enb == _enabled_binding(DELETE_ENABLE_EXPR)
    states.append("Delete.enabled=%s" % ("OK" if del_ok else "MISSING"))
    ok = ok and del_ok

    clear = _find_node(view.get("root", {}), CLEAR_BUTTON)
    clear_always = "props.enabled" not in (clear.get("propConfig", {}) if clear else {})
    states.append("Clear=%s" % ("ALWAYS-ON" if clear_always else "GATED"))
    ok = ok and clear_always

    # P15 gate intact: Save AND Delete onActionPerformed scripts still call the server-side gate.
    for btn in (SAVE_BUTTON, DELETE_BUTTON):
        node = _find_node(view.get("root", {}), btn)
        src = ""
        try:
            src = node["events"]["component"]["onActionPerformed"]["config"]["script"]
        except (KeyError, TypeError):
            pass
        gated = "A.requireWrite(self.session)" in src
        if not gated:
            states.append("%s.GATE=MISSING" % btn)
            ok = False
    return ok, states


# ---------------------------------------------------------------------------
# io
# ---------------------------------------------------------------------------

def _write_view(path, view):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w") as f:
        json.dump(view, f, indent=1, sort_keys=True)
        f.write("\n")


def _normalize_repo_resource(path):
    """Strip the gateway-managed `attributes` (stale signature) from the COMMITTED resource.json after a
    view edit (same convention as gen_master_write_gates._normalize_repo_resource). The DEPLOYED gateway
    resource.json keeps its attributes and is re-signed by the gateway on restart — NOT touched here."""
    if not os.path.exists(path):
        return False
    d = json.load(open(path))
    if "attributes" not in d:
        return False
    d.pop("attributes", None)
    with open(path, "w") as f:
        json.dump(d, f, indent=2, sort_keys=True)
        f.write("\n")
    return True


def run(check_only=False):
    bad = False
    any_change = False
    for V in VIEWS:
        repo_path = os.path.join(REPO_BASE, V, V, "view.json")
        view = json.load(open(repo_path))
        if check_only:
            ok, states = check_view(view, V)
            print("  %-15s %s" % (V, ", ".join(states)))
            if not ok:
                bad = True
            continue

        changed = process_view(view, V)
        changed_keys = [k for k, c in changed.items() if c]
        _write_view(repo_path, view)
        _write_view(os.path.join(GW_BASE, V, V, "view.json"), view)
        res_stripped = _normalize_repo_resource(os.path.join(REPO_BASE, V, V, "resource.json"))
        if changed_keys:
            any_change = True
        print("  %-15s changed=%s repo-resource-stripped=%s -> wrote repo + gateway"
              % (V, changed_keys or "none (idempotent no-op)", res_stripped))

    if check_only:
        return 1 if bad else 0
    print("\nDONE. %s" % ("changes applied" if any_change else "no changes (idempotent no-op)"))
    return 0


def main():
    check_only = "--check" in sys.argv
    print("== master-CRUD form refinement sweep (%s) ==" % ("CHECK" if check_only else "APPLY"))
    sys.exit(run(check_only=check_only))


if __name__ == "__main__":
    main()
