#!/usr/bin/env python3
"""gen_master_theme_tokens.py — replace the hard-coded LIGHT colors in all 8 master CRUD views with the
suite `--tai-*` theme tokens, so both LIGHT and DARK render readable (matching the already-themed
home/nav/landing screens). Deterministic + idempotent (the gen_master_write_gates.py model: re-run =
no-op; --check reports drift).

WHY: the 8 masters hard-code light backgrounds/text (root bg #F5F6F8, card bg #FFFFFF, Title #222222,
borders #D0D4DA, etc.). The suite theme (docs/design/themes/tai/*) swaps --tai-* tokens between
tai-light and tai-dark and maps IA's stock component vars onto them (map.css), so built-in components
(table, dropdowns, inputs) already recolour. But these hand-authored style overrides pin literal light
colors, so in DARK theme the page + form backgrounds stay light and text stays dark -> unreadable.
Replacing the literals with the SAME tokens the home/nav use makes the masters follow the theme.

SCOPE: COLORS -> TOKENS ONLY. No layout/structure changes (keeps #51 refinements + #53 layout intact;
does NOT touch the Save/Delete onActionPerformed scripts -> the P15 server write gates stay
byte-identical). The mapping is keyed on (style-property, literal hex) and is unambiguous across all 8
views (verified: every `color:#FFFFFF` is button text-on-accent; every `backgroundColor:#FFFFFF` is a
container surface; `#1976D2` is the brand accent whether bg or text).

TOKEN MAP (literal light hex -> --tai-* token, by style property):
  backgroundColor #F5F6F8 -> var(--tai-bg)            (root app canvas)
  backgroundColor #FFFFFF -> var(--tai-surface)       (FilterBar / Form cards)
  backgroundColor #1976D2 -> var(--tai-accent)        (Search/Save action buttons)
  backgroundColor #C62828 -> var(--tai-status-danger) (Delete button)
  backgroundColor #607D8B -> var(--tai-neutral-60)    (Clear button — neutral, recolours per theme)
  color           #222222 -> var(--tai-text)          (Title)
  color           #666666 -> var(--tai-text-muted)    (SearchCap)
  color           #666    -> var(--tai-text-muted)    (PartsStock weekday-matrix row labels)
  color           #37474F -> var(--tai-text-muted)    (FooterNote)
  color           #B71C1C -> var(--tai-status-danger) (StatusMsg — error text)
  color           #1976D2 -> var(--tai-accent)        (Sites grp_* section headers)
  color           #FFFFFF -> var(--tai-on-accent)     (button label text on a coloured button)
  <any border*>   #D0D4DA -> var(--tai-border)        (substring within "1px solid #D0D4DA")

The accent buttons (Search/Save) use --tai-on-accent for their label; Delete also uses --tai-on-accent
(white reads on both the light and dark danger reds — ISA-101). Clear is a neutral surface; its label
stays --tai-on-accent (white) which reads on neutral-60 in both themes.

PROVENANCE (same convention as the other master injectors): dual-write BOTH copies so committed ==
deployed. REPO_BASE = committed source-of-truth view.json; GW_BASE = deployed gateway view.json. The
committed resource.json's stale gateway signature (`attributes`) is stripped; the gateway re-signs the
deployed resource on restart (not touched here).

Run:  python3 scripts/gen_master_theme_tokens.py            # apply to all 8 (repo + gateway)
      python3 scripts/gen_master_theme_tokens.py --check    # report only; non-zero exit on any drift
"""
import json
import os
import re
import sys

VIEWS = ["Size", "Supplier", "PartsStock", "ManifestCost", "RenbanGroup", "AssemblyDetail", "Logistics",
         "Sites"]

import sys
_REPO = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _ignenv import PERSPECTIVE_DIR   # noqa: E402 — centralized gateway path (repo-split-plan §4.C)
REPO_BASE = os.path.join(_REPO, "docs", "analysis", "master-data", "perspective-views", "Master")
GW_BASE = os.path.join(PERSPECTIVE_DIR, "views", "Master")

# (style-property, literal hex)  ->  --tai-* token. Property-keyed so the same hex maps to the right
# semantic token (e.g. #FFFFFF is a surface as a backgroundColor but on-accent text as a color).
EXACT_MAP = {
    ("backgroundColor", "#F5F6F8"): "var(--tai-bg)",
    ("backgroundColor", "#FFFFFF"): "var(--tai-surface)",
    ("backgroundColor", "#1976D2"): "var(--tai-accent)",
    ("backgroundColor", "#C62828"): "var(--tai-status-danger)",
    ("backgroundColor", "#607D8B"): "var(--tai-neutral-60)",
    ("color", "#222222"): "var(--tai-text)",
    ("color", "#666666"): "var(--tai-text-muted)",
    ("color", "#666"):    "var(--tai-text-muted)",
    ("color", "#37474F"): "var(--tai-text-muted)",
    ("color", "#B71C1C"): "var(--tai-status-danger)",
    ("color", "#1976D2"): "var(--tai-accent)",
    ("color", "#FFFFFF"): "var(--tai-on-accent)",
}

# For border properties the hex is embedded in a value like "1px solid #D0D4DA" — substring-replace the
# hex -> token. Any style property whose name starts with "border" gets this treatment.
BORDER_HEX_MAP = {
    "#D0D4DA": "var(--tai-border)",
}

_HEX_RE = re.compile(r"#[0-9A-Fa-f]{3}(?:[0-9A-Fa-f]{3})?\b")


def _is_border_prop(name):
    return name.lower().startswith("border")


def _convert_value(prop_name, value):
    """Return (new_value, changed). Handles both exact full-value colors and border substrings.
    Idempotent: a value already containing var(--tai-...) and no literal hex is left untouched."""
    if not isinstance(value, str):
        return value, False
    # already a token, no literal hex left -> no-op
    if "var(--tai-" in value and not _HEX_RE.search(value):
        return value, False

    if _is_border_prop(prop_name):
        new = value
        for hexv, tok in BORDER_HEX_MAP.items():
            new = new.replace(hexv, tok)
        # any OTHER literal hex in a border value we didn't map is a divergence -> surface it
        if _HEX_RE.search(new):
            raise ValueError("unmapped hex in border value %r (prop=%s)" % (value, prop_name))
        return new, (new != value)

    # exact full-value color
    if _HEX_RE.fullmatch(value):
        key = (prop_name, value.upper() if len(value) == 7 else value)
        # normalize: keys store #FFFFFF uppercase (7-char) and #666 as-is (4-char)
        tok = EXACT_MAP.get((prop_name, value)) or EXACT_MAP.get((prop_name, value.upper()))
        if tok is None:
            raise ValueError("unmapped color %r for property %s" % (value, prop_name))
        return tok, (tok != value)
    return value, False


def _walk_styles(node, mutate):
    """Walk the component tree; for every `style` dict, convert color-bearing properties. Returns the
    number of values changed. If mutate is False, only counts (for --check)."""
    changed = 0
    if isinstance(node, dict):
        for k, v in list(node.items()):
            if k == "style" and isinstance(v, dict):
                for prop, val in list(v.items()):
                    if isinstance(val, str) and _HEX_RE.search(val):
                        new, did = _convert_value(prop, val)
                        if did:
                            changed += 1
                            if mutate:
                                v[prop] = new
            else:
                changed += _walk_styles(v, mutate)
    elif isinstance(node, list):
        for it in node:
            changed += _walk_styles(it, mutate)
    return changed


def _count_residual_hex(node):
    """Count literal hex colors remaining inside any `style` dict (the readability metric)."""
    n = 0
    if isinstance(node, dict):
        for k, v in node.items():
            if k == "style" and isinstance(v, dict):
                for val in v.values():
                    if isinstance(val, str):
                        n += len(_HEX_RE.findall(val))
            else:
                n += _count_residual_hex(v)
    elif isinstance(node, list):
        for it in node:
            n += _count_residual_hex(it)
    return n


def process_view(view):
    """Convert all hard-coded style colors to tokens in place. Returns count changed."""
    return _walk_styles(view, mutate=True)


def check_view(view):
    """(ok, state) — ok True iff NO literal hex remains in any style dict."""
    residual = _count_residual_hex(view)
    ok = residual == 0
    return ok, "residual-hex-in-styles=%d (%s)" % (residual, "TOKENIZED" if ok else "HARD-CODED")


def _write_view(path, view):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w") as f:
        json.dump(view, f, indent=1, sort_keys=True)
        f.write("\n")


def _normalize_repo_resource(path):
    """Strip the gateway-managed `attributes` (stale signature) from the COMMITTED resource.json after a
    view edit (same convention as the other master injectors)."""
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
            ok, state = check_view(view)
            print("  %-15s %s" % (V, state))
            if not ok:
                bad = True
            continue
        n = process_view(view)
        _write_view(repo_path, view)
        _write_view(os.path.join(GW_BASE, V, V, "view.json"), view)
        res_stripped = _normalize_repo_resource(os.path.join(REPO_BASE, V, V, "resource.json"))
        if n:
            any_change = True
        print("  %-15s colors->tokens=%d repo-resource-stripped=%s -> wrote repo + gateway"
              % (V, n, res_stripped))
    if check_only:
        return 1 if bad else 0
    print("\nDONE. %s" % ("tokens applied" if any_change else "no changes (idempotent no-op)"))
    return 0


def main():
    check_only = "--check" in sys.argv
    print("== master theme-token sweep (%s) ==" % ("CHECK" if check_only else "APPLY"))
    sys.exit(run(check_only=check_only))


if __name__ == "__main__":
    main()
