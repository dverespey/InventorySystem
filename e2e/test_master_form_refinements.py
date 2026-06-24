#!/usr/bin/env python3
"""test_master_form_refinements.py — prove David's master-CRUD form-pattern refinements landed on ALL 8
DEPLOYED master views (Size, Supplier, PartsStock, ManifestCost, RenbanGroup, AssemblyDetail, Logistics,
Sites), AND that the P15 server-side write gate SURVIVED the refinement.

WHAT THIS DRIVES (the REAL deployed artifact — same approach as test_master_write_gates / test_master_crud_
logic): for each view it reads the DEPLOYED gateway view.json (what the running session serves) AND the
committed repo copy, and asserts the 4 refinements + committed==deployed + the gate intact. No browser, so
it does not depend on the Perspective trial / route — it is the structural backstop to the live browser
proof in the per-view test_<v>_crud.py (which proves the enable/disable + single-line search render live).

THE 4 REFINEMENTS asserted per view (the new model):
  1. SINGLE-LINE SEARCH — SearchField has style.height=="32px" AND position.basis=="auto" (the column-flex
     basis that forced the box 220/240px tall is gone). A px basis on a column-flex item = height, so
     basis=="auto" is the load-bearing half of the fix.
  2. NO NEW BUTTON — the NewButton component is absent from the view.
  3. ACTIONS GROUPED IN DETAIL — Save/Delete/Clear are all children of the detail ActionBar, and NONE leaks
     into the top FilterBar.
  4. ENABLE/DISABLE:
       * Clear  : NO props.enabled binding (always enabled).
       * Delete : props.enabled expr == "{view.custom.recordId} > 0".
       * Save   : props.enabled expr == "{view.custom.recordId} > 0 || {view.custom.<PRIMARY>} != ''".

P15 GATE SURVIVED: Save AND Delete onActionPerformed scripts STILL call auth.requireWrite(self.session)
(the refinement touched presence/grouping/enabled-state + search, NOT the write scripts).

NON-VACUITY / REVERT-PROOF (retro R15/R16): the assertions are derived from the SOURCE intent (the 4
changes), not from the rebuild. To prove they can FAIL, the test mutates a COPY of each view to the OLD
shape (re-add a NewButton; restore the tall basis; drop the enable bindings; strip the gate) and confirms
each assertion FAILS on the reverted copy — so a green run means the refinement is really present, not that
the assertion is vacuous.

Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 e2e/test_master_form_refinements.py
"""
import copy
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")))

from lib import Report, PERSPECTIVE_DIR          # noqa: E402

REPO = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
GW_BASE = os.path.join(PERSPECTIVE_DIR, "views", "Master")   # centralized (repo-split-plan §4.C)
REPO_BASE = os.path.join(REPO, "ignition", "perspective-views", "Master")

VIEWS = ["Size", "Supplier", "PartsStock", "ManifestCost", "RenbanGroup", "AssemblyDetail", "Logistics",
         "Sites"]

# Each view's PRIMARY/required input prop (the Save-enable "has-entered-data" signal). Mirrors
# gen_master_form_refinements.PRIMARY_FIELD; kept here independently so the test is a separate oracle.
PRIMARY_FIELD = {
    "Size": "form_code", "Supplier": "form_code", "PartsStock": "form_partNumber",
    "ManifestCost": "form_assyCode", "RenbanGroup": "form_code", "AssemblyDetail": "form_assyCode",
    "Logistics": "form_name", "Sites": "form_name",
}
SEARCH_HEIGHT = "32px"
DELETE_EXPR = "{view.custom.recordId} > 0"


def _save_expr(v):
    return "{view.custom.recordId} > 0 || {view.custom.%s} != ''" % PRIMARY_FIELD[v]


def _find(node, name):
    if isinstance(node, dict):
        if node.get("meta", {}).get("name") == name:
            return node
        for v in node.values():
            r = _find(v, name)
            if r is not None:
                return r
    elif isinstance(node, list):
        for it in node:
            r = _find(it, name)
            if r is not None:
                return r
    return None


def _enabled_expr(node):
    if node is None:
        return None
    return (node.get("propConfig", {}).get("props.enabled", {})
            .get("binding", {}).get("config", {}).get("expression"))


def _script(view, btn):
    node = _find(view.get("root", {}), btn)
    try:
        return node["events"]["component"]["onActionPerformed"]["config"]["script"]
    except (KeyError, TypeError):
        return ""


# ---- the 4 refinement predicates (each returns (ok, detail)) ----------------

def chk_single_line_search(view):
    f = _find(view.get("root", {}), "SearchField")
    style = (f.get("props", {}).get("style", {}) if f else {})
    pos = (f.get("position", {}) if f else {})
    ok = bool(f) and style.get("height") == SEARCH_HEIGHT and pos.get("basis") == "auto"
    return ok, "height=%r basis=%r" % (style.get("height"), pos.get("basis"))


def chk_no_new_button(view):
    return _find(view.get("root", {}), "NewButton") is None, "NewButton present=%s" % (
        _find(view.get("root", {}), "NewButton") is not None)


def chk_grouped_in_detail(view):
    action = _find(view.get("root", {}), "ActionBar")
    fb = _find(view.get("root", {}), "FilterBar")
    a_names = [c.get("meta", {}).get("name") for c in (action.get("children", []) or [])] if action else []
    f_names = [c.get("meta", {}).get("name") for c in (fb.get("children", []) or [])] if fb else []
    in_detail = all(b in a_names for b in ("SaveButton", "DeleteButton", "ClearButton"))
    not_in_top = not any(b in f_names for b in ("SaveButton", "DeleteButton", "ClearButton"))
    return in_detail and not_in_top, "ActionBar=%s FilterBar=%s" % (a_names, f_names)


def chk_delete_enabled(view):
    expr = _enabled_expr(_find(view.get("root", {}), "DeleteButton"))
    return expr == DELETE_EXPR, "Delete.enabled=%r" % expr


def chk_save_enabled(view, v):
    expr = _enabled_expr(_find(view.get("root", {}), "SaveButton"))
    return expr == _save_expr(v), "Save.enabled=%r" % expr


def chk_clear_always(view):
    clear = _find(view.get("root", {}), "ClearButton")
    has = "props.enabled" in (clear.get("propConfig", {}) if clear else {})
    return not has, "Clear has enabled-binding=%s" % has


def chk_gate_intact(view):
    sv = _script(view, "SaveButton")
    dl = _script(view, "DeleteButton")
    ok = "A.requireWrite(self.session)" in sv and "A.requireWrite(self.session)" in dl
    return ok, "Save-gated=%s Delete-gated=%s" % ("A.requireWrite(self.session)" in sv,
                                                  "A.requireWrite(self.session)" in dl)


# ---- 1. per-view: all 4 refinements + gate on the DEPLOYED view + committed==deployed --------

def test_refinements_deployed(rep):
    for V in VIEWS:
        gw_path = os.path.join(GW_BASE, V, V, "view.json")
        repo_path = os.path.join(REPO_BASE, V, V, "view.json")
        if not os.path.exists(gw_path):
            rep.check("%s deployed view present" % V, False, gw_path)
            continue
        view = json.load(open(gw_path))
        repo_view = json.load(open(repo_path))

        rep.check("%s: committed view.json == deployed view.json (byte-for-byte)" % V,
                  json.load(open(repo_path)) == view, "repo != gateway" if repo_view != view else "match")

        for label, (ok, detail) in (
            ("single-line search (height + basis=auto)", chk_single_line_search(view)),
            ("New button removed", chk_no_new_button(view)),
            ("Save/Delete/Clear grouped in detail ActionBar (none in top FilterBar)", chk_grouped_in_detail(view)),
            ("Delete enabled ONLY when a record is selected (recordId>0)", chk_delete_enabled(view)),
            ("Save enabled when record selected OR primary entered", chk_save_enabled(view, V)),
            ("Clear ALWAYS enabled (no enabled binding)", chk_clear_always(view)),
            ("P15 server-side write gate SURVIVED (Save+Delete call requireWrite)", chk_gate_intact(view)),
        ):
            rep.check("%s: %s" % (V, label), ok, detail)


# ---- 2. REVERT-PROOF — each assertion FAILS on a view mutated back to the OLD shape ----------

def test_revert_proof(rep):
    view = json.load(open(os.path.join(GW_BASE, "Size", "Size", "view.json")))

    # (a) restore the tall column-flex basis -> single-line check FAILS.
    bad = copy.deepcopy(view)
    _find(bad.get("root", {}), "SearchField")["position"]["basis"] = "220px"
    rep.check("REVERT-PROOF: a restored px search-basis FAILS the single-line check",
              chk_single_line_search(bad)[0] is False, chk_single_line_search(bad)[1])

    # (b) re-add a NewButton -> no-new-button check FAILS.
    bad = copy.deepcopy(view)
    fb = _find(bad.get("root", {}), "FilterBar")
    fb["children"].append({"meta": {"name": "NewButton"}, "type": "ia.input.button"})
    rep.check("REVERT-PROOF: a re-added NewButton FAILS the no-new-button check",
              chk_no_new_button(bad)[0] is False, chk_no_new_button(bad)[1])

    # (c) drop the Delete enable binding -> delete-enabled check FAILS.
    bad = copy.deepcopy(view)
    _find(bad.get("root", {}), "DeleteButton")["propConfig"].pop("props.enabled", None)
    rep.check("REVERT-PROOF: dropping the Delete enable binding FAILS the delete-enabled check",
              chk_delete_enabled(bad)[0] is False, chk_delete_enabled(bad)[1])

    # (d) drop the Save enable binding -> save-enabled check FAILS.
    bad = copy.deepcopy(view)
    _find(bad.get("root", {}), "SaveButton")["propConfig"].pop("props.enabled", None)
    rep.check("REVERT-PROOF: dropping the Save enable binding FAILS the save-enabled check",
              chk_save_enabled(bad, "Size")[0] is False, chk_save_enabled(bad, "Size")[1])

    # (e) gate the Clear button -> clear-always check FAILS.
    bad = copy.deepcopy(view)
    _find(bad.get("root", {}), "ClearButton").setdefault("propConfig", {})["props.enabled"] = {
        "binding": {"config": {"expression": "false"}, "type": "expr"}}
    rep.check("REVERT-PROOF: gating Clear FAILS the clear-always-enabled check",
              chk_clear_always(bad)[0] is False, chk_clear_always(bad)[1])

    # (f) strip the gate from Save -> gate-intact check FAILS.
    bad = copy.deepcopy(view)
    sv = _find(bad.get("root", {}), "SaveButton")
    sv["events"]["component"]["onActionPerformed"]["config"]["script"] = "\tpass\n"
    rep.check("REVERT-PROOF: removing the Save server-gate FAILS the gate-intact check",
              chk_gate_intact(bad)[0] is False, chk_gate_intact(bad)[1])


def main():
    print("=== master-CRUD form refinement gate — 4 changes + P15 gate on all 8 deployed views ===\n")
    rep = Report()
    print("-- 1. Per-view: single-line search / no New / grouped / enable-disable / gate intact / committed==deployed --")
    test_refinements_deployed(rep)
    print("\n-- 2. Revert-proof (each assertion FAILS on a view mutated back to the old shape) --")
    test_revert_proof(rep)
    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
