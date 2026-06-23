#!/usr/bin/env python3
"""gen_master_write_gates.py — P15 CUTOVER-BLOCKER sweep: inject the SERVER-SIDE write gate into the
7 remaining hand-laid master-CRUD Perspective views, deterministically + idempotently.

THE HOLE (same as the Sites write closed, gen_sites_view.py rule #1): each of these 7 views writes to
the DB via inline runPrepUpdate / INSERT / DELETE in its Save/Delete button scripts with NO server-side
role check — exactly the reintroduced legacy H3 hazard (the legacy authz was ONE client-side menu-remove
line, trivially bypassed). A modified client / direct websocket could drive a write.

THE FIX (mirrors the PROVEN Sites pattern): each DB-writing button script (and the New form-reset, for
UI defense-in-depth parity with Sites) calls the reusable server-side gate FIRST, BEFORE any system.db
write:

    import auth as A
    try:
        A.requireWrite(self.session)            # resolves roles FROM THE SESSION (gateway-side), not a prop
    except A.AuthError, e:                        # Jython-2 except syntax (8.1.52 + 8.3, no scripting delta)
        c.statusMsg = "DENIED (server-side): %s" % unicode(e); return

auth.requireWrite (docs/analysis/production-readiness/project-library/auth/code.py) resolves the caller's
roles FROM THE SESSION (session.props.auth.user.roles — gateway-populated, NOT client-writable) and
authorizes ProductionControl|Admin, raising AuthError on deny. So a forged client prop or an anonymous
session CANNOT write: the decision is grounded in the server's view of who is logged in. This is the same
hole-closer the Sites Save/Delete/New already carry.

WHY PROGRAMMATIC (not 28 hand-edits): these 7 are hand-laid JSON with no per-view generator, so a single
deterministic injector is the reliable, auditable, RE-RUNNABLE way to apply the identical gate to all of
them. The anchor is the line `\tc = self.view.custom` — which appears EXACTLY ONCE in every Save/Delete/New
script across all 7 views (verified) — so the insertion point is unambiguous. Idempotent: a script that
already contains `A.requireWrite(` is skipped (re-run = no-op).

SCOPE / what this does NOT touch:
  * Only the New/Save/Delete onActionPerformed button scripts are gated. The custom.recordId onChange
    detail-LOAD script is a read-only SELECT (no INSERT/UPDATE/DELETE/runPrepUpdate) and is left ALONE —
    gating it would break loading a row into the form.
  * No qaAdmin URL hatch exists in any of these write paths (verified); none is added.
  * No mayEdit UI-visibility prop is added (the views have no form-hide to wire it to; the REQUIRED thing
    is the server-side requireWrite, which never trusts a client prop anyway). Adding a UI hide is an
    optional later polish; the security boundary is the server-side gate injected here.

PROVENANCE: writes BOTH copies so the committed (reviewable/redeployable) artifact == runtime.
  * REPO_OUT = the committed source-of-truth view.json (re-deploying from it must NOT drop the gate).
  * GW_OUT   = the DEPLOYED gateway view.json (what the running session serves).
The companion resource.json is NOT rewritten here: it already exists for all 7 (scope/version/files +
gateway-managed attributes), and the gateway re-signs on restart. We only touch view.json content.

Run:  python3 scripts/gen_master_write_gates.py            # inject into all 7 (repo + gateway)
      python3 scripts/gen_master_write_gates.py --check    # report only; non-zero exit if any ungated
"""
import json
import os
import sys

# The 7 hand-laid master-CRUD views to sweep (each Master/<V>/<V>/view.json).
VIEWS = ["Size", "Supplier", "PartsStock", "ManifestCost", "RenbanGroup", "AssemblyDetail", "Logistics"]

_REPO = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
REPO_BASE = os.path.join(_REPO, "docs", "analysis", "master-data", "perspective-views", "Master")
GW_BASE = ("/usr/local/ignition/data/projects/spike/com.inductiveautomation.perspective"
           "/views/Master")

# The button components whose onActionPerformed scripts get the gate. Save + Delete are the DB writes
# (the required security boundary); New is a form-reset gated for UI defense-in-depth parity with Sites.
GATED_BUTTONS = ("SaveButton", "DeleteButton", "NewButton")

# The unambiguous anchor: this exact line occurs ONCE in every gated script. The guard is inserted
# immediately AFTER it (so `c` is already bound for the deny-message statusMsg write).
ANCHOR = "\tc = self.view.custom"

# Marker that the script is already gated (idempotency guard) — exact token, not a comment.
GATED_MARKER = "A.requireWrite("


def guard_block(view_name, op_label):
    """The server-side write-gate block to insert after the anchor. Mirrors the Sites Save/Delete/New
    guard verbatim (import auth as A; try requireWrite; except AuthError -> statusMsg DENIED + return).
    Uses the view's statusMsg prop (every one of the 7 uses `statusMsg`) for the deny message, and the
    Jython-2 `except X, e:` form the deployed scripts use (8.1.52 + 8.3, no scripting delta)."""
    # The deny-log line is built WITHOUT %-formatting the whole block (the block contains literal %s/%%
    # tokens that belong to the GENERATED Jython, not to this generator) — the view/op names are inlined
    # directly, and the runtime %s for unicode(e) is left as a literal in the emitted Jython string.
    log_line = (
        "\t\tsystem.util.getLogger(\"SPIKE\").warn("
        "\"SPIKE %s %s DENIED (server-side gate): \" + unicode(e)); return\n"
        % (view_name, op_label)
    )
    return (
        "\t# ---- SERVER-SIDE WRITE GATE (M4 P15 sweep — the H3 hole-closer) ----\n"
        "\t# auth.requireWrite resolves the caller's roles FROM THE SESSION (gateway-side, NOT a client\n"
        "\t# prop) and authorizes ProductionControl|Admin, raising AuthError on deny — BEFORE any\n"
        "\t# system.db write. A forged client prop or an anonymous session is rejected HERE, not by the\n"
        "\t# UI. This closes the reintroduced legacy H3 hole (the write was previously authorized client-\n"
        "\t# side only). Same gate the Sites Save/Delete/New carry (gen_sites_view.py rule #1).\n"
        "\timport auth as A\n"
        "\ttry:\n"
        "\t\tA.requireWrite(self.session)\n"
        "\texcept A.AuthError, e:\n"
        "\t\tc.statusMsg = \"DENIED (server-side): \" + unicode(e)\n"
        + log_line
    )


def _op_label(button_name):
    return {"SaveButton": "Save", "DeleteButton": "Delete", "NewButton": "New"}[button_name]


def inject_gate(script, view_name, button_name):
    """Insert the guard block immediately after the unique ANCHOR line. Returns (new_script, changed).
    Idempotent: if the script already contains the gate marker, return it unchanged.

    Deterministic + defensive: REQUIRES exactly one ANCHOR line. If 0 or >1 are present (an irregular
    script structure) we raise — the caller STOPS and reports, rather than silently injecting wrong."""
    if GATED_MARKER in script:
        return script, False
    lines = script.split("\n")
    anchor_idx = [i for i, ln in enumerate(lines) if ln == ANCHOR]
    if len(anchor_idx) != 1:
        raise ValueError(
            "IRREGULAR write-script structure in %s/%s: anchor %r found %d times (expected exactly 1) "
            "— STOP and inspect manually." % (view_name, button_name, ANCHOR, len(anchor_idx)))
    i = anchor_idx[0]
    guard = guard_block(view_name, _op_label(button_name)).rstrip("\n").split("\n")
    new_lines = lines[: i + 1] + guard + lines[i + 1:]
    return "\n".join(new_lines), True


def _find_button_node(node, button_name):
    """Return the component dict whose meta.name == button_name (depth-first), else None."""
    if isinstance(node, dict):
        if node.get("meta", {}).get("name") == button_name and "events" in node:
            return node
        for v in node.values():
            r = _find_button_node(v, button_name)
            if r is not None:
                return r
    elif isinstance(node, list):
        for it in node:
            r = _find_button_node(it, button_name)
            if r is not None:
                return r
    return None


def process_view(view, view_name):
    """Mutate `view` (a parsed view.json dict) in place: gate each of GATED_BUTTONS. Returns a list of
    (button_name, changed) results. Raises if a targeted button is missing or its script is irregular."""
    results = []
    for btn in GATED_BUTTONS:
        comp = _find_button_node(view.get("root", {}), btn)
        if comp is None:
            raise ValueError("%s: expected button component %r not found (irregular structure)"
                             % (view_name, btn))
        try:
            cfg = comp["events"]["component"]["onActionPerformed"]["config"]
        except (KeyError, TypeError):
            raise ValueError("%s/%s: no onActionPerformed script (irregular structure)"
                             % (view_name, btn))
        src = cfg.get("script")
        if src is None:
            raise ValueError("%s/%s: onActionPerformed has no script string" % (view_name, btn))
        new_src, changed = inject_gate(src, view_name, btn)
        cfg["script"] = new_src
        results.append((btn, changed))
    return results


def _write_view(path, view):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w") as f:
        json.dump(view, f, indent=1, sort_keys=True)
        f.write("\n")


def _normalize_repo_resource(path):
    """Strip the gateway-managed `attributes` (lastModification + signature) from the COMMITTED
    resource.json, leaving only the load-bearing scope/version/files/restricted/overridable. The
    committed copy must NOT carry a signature: after we change view.json its old signature is stale, and
    a stale committed signature is the cold-start-fault hazard if the repo is ever deployed to a fresh
    gateway (the gateway would validate the committed signature against the changed view.json and fault).
    The DEPLOYED gateway resource.json (which DOES carry attributes) is owned + re-signed by the gateway
    on restart and is deliberately NOT touched here (memory: reference-headless-ignition-authoring-limits
    + the gen_sites_view.py REPO_RESOURCE convention — committed omits attributes, gateway carries them).
    Returns True if the file was changed."""
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
    any_change = False
    any_ungated = False
    for V in VIEWS:
        repo_path = os.path.join(REPO_BASE, V, V, "view.json")
        view = json.load(open(repo_path))
        if check_only:
            # report-only: is every gated button already gated?
            states = []
            for btn in GATED_BUTTONS:
                comp = _find_button_node(view.get("root", {}), btn)
                src = comp["events"]["component"]["onActionPerformed"]["config"]["script"] if comp else ""
                gated = GATED_MARKER in src
                states.append("%s=%s" % (_op_label(btn), "GATED" if gated else "UNGATED"))
                if not gated:
                    any_ungated = True
            print("  %-15s %s" % (V, ", ".join(states)))
            continue

        results = process_view(view, V)
        changed = [b for b, c in results if c]
        skipped = [b for b, c in results if not c]
        # write the gated view to BOTH the committed repo copy AND the deployed gateway copy.
        _write_view(repo_path, view)
        gw_path = os.path.join(GW_BASE, V, V, "view.json")
        _write_view(gw_path, view)
        # committed resource.json: strip the now-stale gateway signature (committed omits attributes;
        # gateway re-signs its own copy on restart). The gateway resource.json is left untouched.
        res_changed = _normalize_repo_resource(os.path.join(REPO_BASE, V, V, "resource.json"))
        if changed:
            any_change = True
        print("  %-15s gated=%s skipped(already)=%s repo-resource-stripped=%s -> wrote repo + gateway"
              % (V, [_op_label(b) for b in changed], [_op_label(b) for b in skipped], res_changed))

    if check_only:
        return 1 if any_ungated else 0
    print("\nDONE. %s" % ("changes applied" if any_change else "no changes (all already gated — idempotent no-op)"))
    return 0


def main():
    check_only = "--check" in sys.argv
    print("== P15 master write-gate sweep (%s) ==" % ("CHECK" if check_only else "INJECT"))
    sys.exit(run(check_only=check_only))


if __name__ == "__main__":
    main()
