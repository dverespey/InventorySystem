#!/usr/bin/env python3
"""gen_master_detail_layout.py — David's master-detail LAYOUT sweep (click-through 2026-06-24), applied
deterministically + idempotently across ALL 8 master views (Size, Supplier, PartsStock, ManifestCost,
RenbanGroup, AssemblyDetail, Logistics, Sites). Same model as gen_master_form_refinements.py /
gen_master_write_gates.py: re-run = no-op, --check reports drift, writes BOTH the committed repo copy AND
the deployed gateway copy (committed == deployed), strips the stale signature from the COMMITTED
resource.json only (the gateway resource.json keeps `attributes` and the gateway re-signs on restart).

TWO uniform NITs (NIT 1, the Sites readability fix, is Sites-specific and lives in gen_sites_view.py):

  NIT 2 — REMOVE THE "Detail" HEADING. Every master's RightPane led with a DetailTitle label
          ("…Detail"/"Detail"). David: unnecessary — the master is already identified by the nav/page.
          DELETE the DetailTitle component from all 8 RightPanes. (It carries nothing load-bearing — just a
          static text label; verified before deletion.)

  NIT 3 — ALIGN THE TOPS OF THE LIST AND THE DETAIL BOX. The LeftPane GRID and the RightPane DETAIL BOX
          (the Form, the white bordered box) started at DIFFERENT y. The LeftPane stack is
          Title -> FilterBar -> Grid; the RightPane (after NIT 2) is StatusMsg -> Form, so the Form sat
          HIGHER than the Grid. FIX: give the RightPane a TOP BAND above the Form of the SAME height as the
          LeftPane's Title+FilterBar, so the Form's top edge lines up with the Grid's top edge. The band is:
            * a TitleSpacer label (empty, pinned to the Title line-height) matching the LeftPane Title row, and
            * the StatusMsg relocated INTO an AlignBand of fixed height == the FilterBar height,
          so RightPane top stack height == LeftPane top stack height == Grid offset. The StatusMsg keeps its
          binding + domId + styling (status text still shows, now in the aligned band). Uniform across all 8.

ORDERING (must compose): gen_sites_view.py  ->  gen_master_form_refinements.py  ->  THIS script.
  gen_sites regenerates Sites (re-adds NewButton + DetailTitle, resets the search box); the refinement
  sweep then re-removes New + pins the search; THIS sweep then removes DetailTitle + aligns the tops. After
  running this, re-run `gen_master_form_refinements.py --check` (must still be clean on all 8) to confirm no
  #51 refinement or server-gate was disturbed (this script touches ONLY DetailTitle + the RightPane top
  band; it never edits the FilterBar search, the action buttons, or any Save/Delete onActionPerformed
  script).

Run:  python3 scripts/gen_master_detail_layout.py            # apply to all 8 (repo + gateway)
      python3 scripts/gen_master_detail_layout.py --check    # report only; non-zero exit on any drift
"""
import json
import os
import sys

VIEWS = ["Size", "Supplier", "PartsStock", "ManifestCost", "RenbanGroup", "AssemblyDetail", "Logistics",
         "Sites"]

_REPO = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
REPO_BASE = os.path.join(_REPO, "docs", "analysis", "master-data", "perspective-views", "Master")
GW_BASE = ("/usr/local/ignition/data/projects/InventorySystem/com.inductiveautomation.perspective"
           "/views/Master")

DETAIL_TITLE = "DetailTitle"
RIGHT_PANE = "RightPane"
LEFT_PANE = "LeftPane"
TITLE = "Title"
STATUS_MSG = "StatusMsg"
TITLE_SPACER = "TitleSpacer"   # alignment-only label that MIRRORS the LeftPane Title (invisible text)
ALIGN_BAND = "AlignBand"       # alignment-only band that mirrors the LeftPane FilterBar height (holds Status)

# The AlignBand height == the rendered FilterBar height. The FilterBar is UNIFORM across all 8 (a 11px
# caption + a 32px single-line search field + 10px*2 padding inside the bordered box), so a single constant
# is correct for every view.
ALIGN_BAND_HEIGHT = "70px"     # == rendered FilterBar height (bordered box: caption+32px search+2*10 pad);
                               # calibrated to zero the Grid-top vs Form-top delta (was 74 -> Form sat ~4px low)
RIGHT_PANE_GAP = "10px"        # == LeftPane gap, so the top-band stack height is symmetric with the left

# To align the tops EXACTLY we pin BOTH the LeftPane Title AND the RightPane TitleSpacer to ONE line of a
# fixed height. (Mirroring the Title TEXT into the spacer does NOT work: the Title wraps in the narrow
# LeftPane but the wider RightPane fits the SAME text on one line, so the heights diverge — Sites was 15px
# off. Forcing BOTH to a single fixed-height line guarantees identical height regardless of pane width.) The
# Title gets whiteSpace:nowrap + ellipsis so a long subtitle (Sites) truncates on one line instead of
# wrapping; the full identity is still in the page/nav (NIT-2 rationale) and the grid columns.
TITLE_LINE_HEIGHT = "22px"     # one 15px-bold line; identical on Title and TitleSpacer -> tops align


# ---------------------------------------------------------------------------
# tree helpers (shared shape with gen_master_form_refinements)
# ---------------------------------------------------------------------------

def _find_node(node, name):
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


# ---------------------------------------------------------------------------
# NIT 2 — remove DetailTitle (idempotent)
# ---------------------------------------------------------------------------

def _is_safe_to_drop_detail_title(node):
    """Guard: only delete DetailTitle if it is the inert static heading (a plain label with no binding/
    event/children). If a view ever made it load-bearing, STOP rather than silently drop something."""
    if node.get("type") != "ia.display.label":
        return False
    if node.get("propConfig") or node.get("events") or node.get("children"):
        return False
    return True


def _remove_detail_title(view, view_name):
    parent = _find_parent_with_child(view.get("root", {}), DETAIL_TITLE)
    if parent is None:
        if _find_node(view.get("root", {}), DETAIL_TITLE) is not None:
            raise ValueError("%s: %s present but not a direct child of any container" % (view_name, DETAIL_TITLE))
        return False  # already removed (idempotent)
    if parent.get("meta", {}).get("name") != RIGHT_PANE:
        raise ValueError("%s: %s parent is %r, expected %s" % (view_name, DETAIL_TITLE,
                         parent.get("meta", {}).get("name"), RIGHT_PANE))
    dt = next(c for c in parent["children"]
              if isinstance(c, dict) and c.get("meta", {}).get("name") == DETAIL_TITLE)
    if not _is_safe_to_drop_detail_title(dt):
        raise ValueError("%s: %s carries load-bearing config (binding/event/children) — refusing to drop"
                         % (view_name, DETAIL_TITLE))
    parent["children"] = [c for c in parent["children"]
                          if not (isinstance(c, dict) and c.get("meta", {}).get("name") == DETAIL_TITLE)]
    return True


# ---------------------------------------------------------------------------
# NIT 3 — align the tops: RightPane top band == LeftPane Title+FilterBar height
# ---------------------------------------------------------------------------

def _title_spacer_node():
    """An alignment-only EMPTY label pinned to ONE Title line-height (TITLE_LINE_HEIGHT). Paired with the
    single-line-pinned LeftPane Title (see _pin_title_single_line), this reserves exactly the Title's
    vertical space so the Form top lines up with the Grid top — independent of pane width / text length."""
    return {
        "meta": {"name": TITLE_SPACER},
        "position": {"shrink": 0},
        "props": {"style": {"height": TITLE_LINE_HEIGHT}, "text": ""},
        "type": "ia.display.label",
    }


def _pin_title_single_line(view, view_name):
    """Pin the LeftPane Title to ONE fixed-height line (nowrap + ellipsis), so it can NEVER wrap to 2 lines
    and de-sync from the single-line TitleSpacer. Idempotent. Returns True if it changed the Title."""
    title = _find_node(_find_node(view.get("root", {}), LEFT_PANE) or {}, TITLE)
    if title is None:
        raise ValueError("%s: %s not found in %s" % (view_name, TITLE, LEFT_PANE))
    style = title.setdefault("props", {}).setdefault("style", {})
    want = {"height": TITLE_LINE_HEIGHT, "whiteSpace": "nowrap",
            "overflow": "hidden", "textOverflow": "ellipsis"}
    changed = False
    for k, v in want.items():
        if style.get(k) != v:
            style[k] = v
            changed = True
    # the Title must not shrink in the column flex (keep its line)
    pos = title.setdefault("position", {})
    if pos.get("shrink") != 0:
        pos["shrink"] = 0
        changed = True
    return changed


def _align_band_node(status_node):
    """The fixed-height band (== FilterBar height) that holds the relocated StatusMsg, so the Form top lines
    up with the Grid top. StatusMsg keeps its domId/binding/style; the band centers it vertically."""
    return {
        "meta": {"name": ALIGN_BAND},
        "position": {"shrink": 0},
        "props": {"alignItems": "center", "style": {"height": ALIGN_BAND_HEIGHT}},
        "type": "ia.container.flex",
        "children": [status_node],
    }


def _align_tops(view, view_name):
    """Idempotently ensure the RightPane leads with [TitleSpacer, AlignBand(StatusMsg), Form, ActionBar] so
    the Form top aligns with the Grid top. Returns True if it changed the view."""
    rp = _find_node(view.get("root", {}), RIGHT_PANE)
    if rp is None:
        raise ValueError("%s: %s not found" % (view_name, RIGHT_PANE))

    # Normalize the RightPane vertical gap to the LeftPane gap so the top-band stack height is symmetric with
    # the LeftPane Title+FilterBar stack (the alignment math assumes equal gaps on both panes). All 8 use
    # LeftPane gap 10px; pin the RightPane to the same.
    gap_changed = False
    rp_style = rp.setdefault("props", {}).setdefault("style", {})
    if rp_style.get("gap") != RIGHT_PANE_GAP:
        rp_style["gap"] = RIGHT_PANE_GAP
        gap_changed = True

    want_spacer = _title_spacer_node()

    children = rp.get("children", []) or []
    names = [c.get("meta", {}).get("name") for c in children if isinstance(c, dict)]

    # already aligned? Require the FULL desired top structure: TitleSpacer (matching the single-line spacer
    # props) + AlignBand (matching the CURRENT band height — so a calibration change to ALIGN_BAND_HEIGHT is
    # detected as drift and rebuilt, not skipped) + StatusMsg nested in the band (not a loose direct child).
    if TITLE_SPACER in names and ALIGN_BAND in names and STATUS_MSG not in names:
        cur_spacer = _find_node(rp, TITLE_SPACER)
        cur_band = _find_node(rp, ALIGN_BAND)
        spacer_ok = cur_spacer.get("props") == want_spacer.get("props")
        band_ok = cur_band.get("props", {}).get("style", {}).get("height") == ALIGN_BAND_HEIGHT
        if spacer_ok and band_ok:
            return gap_changed

    # locate the StatusMsg (currently a direct RightPane child) to relocate into the band
    status = None
    for c in children:
        if isinstance(c, dict) and c.get("meta", {}).get("name") == STATUS_MSG:
            status = c
            break
    if status is None:
        # maybe already inside a prior AlignBand — pull it back out to rebuild deterministically
        band = _find_node(rp, ALIGN_BAND)
        if band is not None:
            for c in band.get("children", []) or []:
                if isinstance(c, dict) and c.get("meta", {}).get("name") == STATUS_MSG:
                    status = c
                    break
    if status is None:
        raise ValueError("%s: %s not found in %s (cannot build the align band)" % (view_name, STATUS_MSG, RIGHT_PANE))

    # rebuild the RightPane children deterministically: drop any prior TitleSpacer/AlignBand/StatusMsg, then
    # prepend the fresh [TitleSpacer, AlignBand(StatusMsg)] in front of the remaining children (Form, ActionBar)
    rest = [c for c in children if isinstance(c, dict)
            and c.get("meta", {}).get("name") not in (TITLE_SPACER, ALIGN_BAND, STATUS_MSG)]
    rp["children"] = [want_spacer, _align_band_node(status)] + rest
    return True


def process_view(view, view_name):
    changed = {}
    changed["detail_title_removed"] = _remove_detail_title(view, view_name)
    changed["title_single_line"] = _pin_title_single_line(view, view_name)
    changed["tops_aligned"] = _align_tops(view, view_name)
    return changed


# ---------------------------------------------------------------------------
# check (report-only)
# ---------------------------------------------------------------------------

def check_view(view, view_name):
    states = []
    ok = True

    dt_gone = _find_node(view.get("root", {}), DETAIL_TITLE) is None
    states.append("DetailTitle=%s" % ("REMOVED" if dt_gone else "PRESENT"))
    ok = ok and dt_gone

    rp = _find_node(view.get("root", {}), RIGHT_PANE)
    names = [c.get("meta", {}).get("name") for c in (rp.get("children", []) or []) if isinstance(c, dict)] if rp else []
    ts = _find_node(rp, TITLE_SPACER) if rp else None
    band = _find_node(rp, ALIGN_BAND) if rp else None
    status_in_band = bool(band) and _find_node(band, STATUS_MSG) is not None
    # the band leads the RightPane (alignment requires the spacer+band be the first two children)
    leads = names[:2] == [TITLE_SPACER, ALIGN_BAND]
    # the TitleSpacer is the single-line height spacer
    th_ok = bool(ts) and ts.get("props") == _title_spacer_node().get("props")
    bh_ok = bool(band) and band.get("props", {}).get("style", {}).get("height") == ALIGN_BAND_HEIGHT
    gap_ok = bool(rp) and rp.get("props", {}).get("style", {}).get("gap") == RIGHT_PANE_GAP
    status_not_loose = STATUS_MSG not in names
    # the LeftPane Title is pinned single-line so it can't wrap and de-sync from the spacer
    title = _find_node(_find_node(view.get("root", {}), LEFT_PANE) or {}, TITLE)
    tstyle = (title.get("props", {}).get("style", {}) if title else {})
    title_pinned = (tstyle.get("whiteSpace") == "nowrap" and tstyle.get("height") == TITLE_LINE_HEIGHT
                    and tstyle.get("textOverflow") == "ellipsis" and tstyle.get("overflow") == "hidden")
    aligned = leads and status_in_band and th_ok and bh_ok and gap_ok and status_not_loose and title_pinned
    states.append("tops=%s" % ("ALIGNED" if aligned else "MISALIGNED"))
    ok = ok and aligned

    # StatusMsg binding survived (still bound to view.custom.statusMsg) — alignment must not break status text
    st = _find_node(rp, STATUS_MSG) if rp else None
    st_bound = bool(st) and (st.get("propConfig", {}).get("props.text", {})
                             .get("binding", {}).get("config", {}).get("path") == "view.custom.statusMsg")
    if not st_bound:
        states.append("StatusMsg.binding=MISSING")
        ok = False

    return ok, states


# ---------------------------------------------------------------------------
# io (same convention as gen_master_form_refinements)
# ---------------------------------------------------------------------------

def _write_view(path, view):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w") as f:
        json.dump(view, f, indent=1, sort_keys=True)
        f.write("\n")


def _normalize_repo_resource(path):
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
    print("== master-detail LAYOUT sweep (%s) ==" % ("CHECK" if check_only else "APPLY"))
    sys.exit(run(check_only=check_only))


if __name__ == "__main__":
    main()
