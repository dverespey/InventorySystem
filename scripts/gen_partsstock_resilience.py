#!/usr/bin/env python3
"""gen_partsstock_resilience.py — P21 resilience fix for the PartsStock master detail-load.

ROOT CAUSE (David's testing): clicking a PartsStock grid row did NOT fill the detail. The recordId
onChange detail-load builds the FK-combo option lists FIRST (5 Inventory master lookups + the cross-DB
Line list), THEN loads the part's own fields. The Line list is a CROSS-DB read of
`VehicleOrder.dbo.LINE`. The gateway login (ignition_spike) had no user mapping in the VehicleOrder DB,
so that single read raised ("server principal ... not able to access the database VehicleOrder") and,
because it was an UNGUARDED inline assignment, the WHOLE onChange aborted BEFORE populating the form —
so NO field filled on click. One unavailable FK-dropdown source took down the entire detail-load.

THE FIX (2a, the robust one — mirrors the HotCall AD_GetLines graceful pattern, gen_hotcall_view.py):
wrap ONLY the VehicleOrder (Line) read in try/except. On failure -> empty Line dropdown + log.warn, and
the load CONTINUES: the part's own fields STILL load regardless of the VehicleOrder grant. The 5
Inventory-DB master lookups are NOT wrapped (they hit the always-available Inventory_Spike DB; if THOSE
fail it is a real config error that should surface, not be swallowed).

  BEFORE (single unguarded line — any VehicleOrder denial kills the whole onChange):
      c.lineOptions = opts("SELECT DISTINCT LineName AS id, LineName AS label FROM VehicleOrder.dbo.LINE ORDER BY LineName")

  AFTER (degrade gracefully; the rest of the detail-load proceeds):
      c.lineOptions = []
      try:
          c.lineOptions = opts("SELECT DISTINCT LineName AS id, LineName AS label FROM VehicleOrder.dbo.LINE ORDER BY LineName")
      except Exception as e:
          log.warn("SPIKE PartsStock Detail: Line list (cross-DB VehicleOrder) unavailable, dropdown empty: %s" % e)

(2b — the spike GRANT that populates the dropdown for testing — is a one-time SQL action on the spike,
applied separately; the PROD grant at cutover stays on the M5 punch-list. This script is only the code
resilience so the part fields load even WITHOUT that grant.)

PROVENANCE (same convention as the other master injectors): dual-write BOTH copies so committed ==
deployed. REPO_BASE = committed source-of-truth view.json; GW_BASE = the deployed gateway view.json.
The committed resource.json's stale gateway signature (`attributes`) is stripped; the gateway re-signs
the deployed resource on restart (not touched here).

Run:  python3 scripts/gen_partsstock_resilience.py            # apply (repo + gateway)
      python3 scripts/gen_partsstock_resilience.py --check    # report only; non-zero exit on drift
"""
import json
import os
import sys

_REPO = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _ignenv import PERSPECTIVE_DIR   # noqa: E402 — centralized gateway path (repo-split-plan §4.C)
REPO_VIEW = os.path.join(_REPO, "docs", "analysis", "master-data", "perspective-views", "Master",
                         "PartsStock", "PartsStock", "view.json")
REPO_RES = os.path.join(_REPO, "docs", "analysis", "master-data", "perspective-views", "Master",
                        "PartsStock", "PartsStock", "resource.json")
GW_VIEW = os.path.join(PERSPECTIVE_DIR, "views", "Master", "PartsStock", "PartsStock", "view.json")

# The unguarded VehicleOrder read line (exact, tab-indented) that aborts the whole onChange on denial.
OLD_LINE = ('\tc.lineOptions = opts("SELECT DISTINCT LineName AS id, LineName AS label FROM '
            'VehicleOrder.dbo.LINE ORDER BY LineName")')

# Graceful replacement: init empty, attempt the cross-DB read, log.warn + leave [] on failure so the
# detail-load CONTINUES and the part's own fields still populate (mirrors HotCall AD_GetLines).
#
# CATCH TYPE — java.lang.Throwable, NOT plain `except Exception` (learned on the box 2026-06-24):
#   a SQL-Server cross-DB ACCESS-DENIED ("server principal ... not able to access the database
#   VehicleOrder") surfaces from system.db.runPrepQuery as a Java throwable that a Jython `except
#   Exception:` does NOT catch — the exception propagated PAST an `except Exception` wrap and the whole
#   onChange still aborted (verified: detail did NOT fill, no warn logged, traceback line numbers matched
#   THIS script so it was running the wrapped version). Catching java.lang.Throwable catches both the Java
#   DB throwable AND any Python exception, so the load degrades gracefully for real. (The master Save
#   scripts get away with `except Exception` because they catch driver CONSTRAINT errors, a different
#   path; a cross-DB principal-denied is not caught the same way.)
NEW_BLOCK = (
    '\t# IG-RESILIENCE (P21): the Line list is a CROSS-DB read of VehicleOrder.dbo.LINE. If the gateway\n'
    '\t# login lacks VehicleOrder access the read raises; wrap it so the detail-load DEGRADES GRACEFULLY\n'
    '\t# (empty Line dropdown + warn) and the part\'s OWN fields STILL load. Mirrors HotCall AD_GetLines.\n'
    '\t# Catch java.lang.Throwable: a SQL-Server cross-DB ACCESS-DENIED surfaces as a Java throwable that\n'
    '\t# a plain `except Exception` does NOT catch (verified on the box) — Throwable catches both.\n'
    '\tfrom java.lang import Throwable as _JThrowable\n'
    '\tc.lineOptions = []\n'
    '\ttry:\n'
    '\t\tc.lineOptions = opts("SELECT DISTINCT LineName AS id, LineName AS label FROM '
    'VehicleOrder.dbo.LINE ORDER BY LineName")\n'
    '\texcept _JThrowable, e:\n'
    '\t\tlog.warn("SPIKE PartsStock Detail: Line list (cross-DB VehicleOrder) unavailable, '
    'dropdown empty: %s" % e)'
)

# A marker that the AFTER state is present (idempotency probe) — the resilience comment tag.
APPLIED_MARKER = "# IG-RESILIENCE (P21)"

# First-gen (Exception-only) resilience block that did NOT catch the Java DB throwable — migrate it to
# NEW_BLOCK. Matched by the distinguishing `except Exception as e:` inside the P21 block.
FIRSTGEN_BLOCK = (
    '\t# IG-RESILIENCE (P21): the Line list is a CROSS-DB read of VehicleOrder.dbo.LINE. If the gateway\n'
    '\t# login lacks VehicleOrder access the read raises; wrap it so the detail-load DEGRADES GRACEFULLY\n'
    '\t# (empty Line dropdown + warn) and the part\'s OWN fields STILL load. Mirrors HotCall AD_GetLines.\n'
    '\tc.lineOptions = []\n'
    '\ttry:\n'
    '\t\tc.lineOptions = opts("SELECT DISTINCT LineName AS id, LineName AS label FROM '
    'VehicleOrder.dbo.LINE ORDER BY LineName")\n'
    '\texcept Exception as e:\n'
    '\t\tlog.warn("SPIKE PartsStock Detail: Line list (cross-DB VehicleOrder) unavailable, '
    'dropdown empty: %s" % e)'
)


def _onchange_script(view):
    return view["propConfig"]["custom.recordId"]["onChange"]["script"]


def _set_onchange_script(view, script):
    view["propConfig"]["custom.recordId"]["onChange"]["script"] = script


# The distinguishing token of the FINAL (correct) block: catches the Java throwable.
THROWABLE_CATCH = "except _JThrowable, e:"


def process_view(view):
    """Apply (or migrate to) the resilience wrap idempotently. Returns True if it changed the view."""
    script = _onchange_script(view)
    if THROWABLE_CATCH in script:
        return False                                  # already the final block (idempotent no-op)
    # Migrate the first-gen Exception-only block (which did not catch the Java DB throwable) -> NEW_BLOCK.
    if FIRSTGEN_BLOCK in script:
        if script.count(FIRSTGEN_BLOCK) != 1:
            raise ValueError("PartsStock onChange: expected exactly 1 first-gen resilience block, found %d"
                             % script.count(FIRSTGEN_BLOCK))
        _set_onchange_script(view, script.replace(FIRSTGEN_BLOCK, NEW_BLOCK))
        return True
    # Fresh apply from the original unguarded single line.
    if OLD_LINE not in script:
        raise ValueError("PartsStock onChange: neither the final/first-gen resilience block NOR the "
                         "unguarded VehicleOrder Line read line was found — view drifted from the "
                         "expected shape; refusing to guess.")
    if script.count(OLD_LINE) != 1:
        raise ValueError("PartsStock onChange: expected exactly 1 VehicleOrder Line read line, found %d"
                         % script.count(OLD_LINE))
    _set_onchange_script(view, script.replace(OLD_LINE, NEW_BLOCK))
    return True


def check_view(view):
    """(ok, state) — ok True iff the FINAL (Throwable-catching) wrap is present and the part-load body is
    intact. A first-gen Exception-only wrap is reported as drift (it did not catch the Java DB error)."""
    script = _onchange_script(view)
    wrapped = APPLIED_MARKER in script and THROWABLE_CATCH in script
    # The part-load body must still be present (we only wrapped the Line read, nothing else).
    load_intact = "SPIKE PartsStock Detail loaded id=%d" in script and "INV_PARTS_STOCK_MST p WHERE" in script
    ok = wrapped and load_intact
    state = "Line-read=%s, part-load-body=%s" % (
        "WRAPPED(Throwable)" if wrapped else "UNGUARDED/STALE", "INTACT" if load_intact else "MISSING")
    return ok, state


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
    view = json.load(open(REPO_VIEW))
    if check_only:
        ok, state = check_view(view)
        print("  PartsStock      %s" % state)
        return 0 if ok else 1
    changed = process_view(view)
    _write_view(REPO_VIEW, view)
    _write_view(GW_VIEW, view)
    res_stripped = _normalize_repo_resource(REPO_RES)
    print("  PartsStock      changed=%s repo-resource-stripped=%s -> wrote repo + gateway"
          % (changed, res_stripped))
    print("\nDONE. %s" % ("resilience wrap applied" if changed else "no change (idempotent no-op)"))
    return 0


def main():
    check_only = "--check" in sys.argv
    print("== PartsStock detail-load resilience (P21) (%s) ==" % ("CHECK" if check_only else "APPLY"))
    sys.exit(run(check_only=check_only))


if __name__ == "__main__":
    main()
