#!/usr/bin/env python3
"""test_master_write_gates.py — P15 CUTOVER-BLOCKER gate: prove the SERVER-SIDE write gate on the 7
swept master-CRUD views (Size, Supplier, PartsStock, ManifestCost, RenbanGroup, AssemblyDetail,
Logistics). Companion to test_m4_auth.py's Sites write-gate proof — same approach, all 7 views.

WHAT THIS DRIVES (the REAL gateway-side code, not a reimplementation):
  For EACH of the 7 views we pull the DEPLOYED Save AND Delete onActionPerformed scripts straight from
  the gateway view.json (what the running session serves), compile them (only the Jython-2 `except X, e:`
  -> Py3 `as` syntactic shim), and run them with a fabricated `self` (view.custom form props + session)
  and a `system.db` stub that RECORDS whether the path reached a DB write (INSERT/UPDATE/DELETE). The
  `import auth as A` inside each script resolves to the REAL auth Project Library module (the one the
  gateway loads). So the gate proven here is the gate the gateway runs.

THE BLOCKER PROOF (per view, both Save and Delete):
  * A forged client prop (mayEdit=true) + a session with NO write role (anonymous OR a logged-in viewer)
    is REJECTED SERVER-SIDE — statusMsg "DENIED", NO db write. The SERVER gate (auth.requireWrite reading
    the SESSION roles), not the client prop, is what blocks it.
  * NON-VACUITY: a ProductionControl session (and an Admin session) REACHES the db write — so the gate is
    not block-all, and a green deny is meaningful.
  * THE GATE KEYS OFF THE SESSION, NOT THE PROP: a PC session with mayEdit=FALSE still writes.

REVERT-PROOF (retro R15/R16/R20 — a test that can't fail on a removed gate tests nothing):
  We neuter the SERVER gate (auth.requireWrite -> no-op) and confirm the forged-prop anon SAVE then SLIPS
  THROUGH to the db write on a representative view — proving the gate (not the harness/prop) is what
  blocks it. (One revert-proof is sufficient: the gate code is the SAME module call in all 7; the per-view
  non-vacuity above already proves each view's authorized path reaches the write.)

HONEST SPLIT: the LIVE end-to-end deny through the browser chrome (anon session denied via the running
Perspective UI) is proven for Sites by test_sites_crud.check_server_side_write_gate; for these 7 the same
deny is proven here through the DEPLOYED scripts (forged-prop-rejected, revert-proven, non-vacuous). The
gateway-loads-clean check (no Unable-to-deserialize / FAULTED) is asserted by the deploy step + grep.

Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 scripts/e2e/test_master_write_gates.py
"""
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")))

from lib import Report, PERSPECTIVE_DIR       # noqa: E402
import jython_shim                            # noqa: E402

REPO = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".."))
AUTH_CODE = os.path.join(REPO, "docs", "analysis", "production-readiness", "project-library", "auth", "code.py")
GW_BASE = os.path.join(PERSPECTIVE_DIR, "views", "Master")   # centralized (repo-split-plan §4.C)
REPO_BASE = os.path.join(REPO, "docs", "analysis", "master-data", "perspective-views", "Master")

VIEWS = ["Size", "Supplier", "PartsStock", "ManifestCost", "RenbanGroup", "AssemblyDetail", "Logistics"]

# Sessions (the SAME shapes test_m4_auth uses). Roles come from session.props.auth.user.roles, which the
# auth module reads via sessionRoles — NOT a client prop. The harness dict shape is what jython_shim and
# auth.sessionRoles both accept.
ANON_SESSION = {"auth": {"user": {"userName": None, "roles": []}}}                       # not logged in
VIEWER_SESSION = {"auth": {"user": {"userName": "nobody", "roles": ["SomeOtherRole"]}}}  # logged in, no write role
PC_SESSION = {"auth": {"user": {"userName": "op1", "roles": ["ProductionControl"]}}}
ADMIN_SESSION = {"auth": {"user": {"userName": "boss", "roles": ["Admin"]}}}

# Minimal VALID form fixtures per view (just enough to PASS each view's pre-write validation so the
# authorized non-vacuity case REACHES the db write). Derived from each Save script's own validation guards
# (read directly from the deployed view, see the module docstring); the DB stub returns 0 for every
# dup/overlap check so validation passes. Fields not listed default to "" (their custom-prop default).
VALID_FIXTURE = {
    "Size":           {"form_code": "ZZG", "form_name": "QA Size"},
    "Supplier":       {"form_code": "ZZSUP", "form_name": "QA Supplier", "form_invAddPoint": "S"},
    "PartsStock":     {"form_partNumber": "ZZPART01", "form_partsName": "QA Part"},
    "ManifestCost":   {"form_assyCode": "ZZASSY", "form_start": "20260101", "form_end": "20261231",
                       "form_price": "1.50"},
    "RenbanGroup":    {"form_code": "ZZR", "form_count": "5"},
    "AssemblyDetail": {"form_assyCode": "ZZASSY01", "form_effMonth": "202601"},
    "Logistics":      {"form_name": "QA Logistics"},
}


class _GateDB(object):
    """Records whether a master Save/Delete path reached a DB WRITE (= passed the gate + validation).
    Mirrors test_m4_auth._SitesDB: every dup/overlap/refCount read returns 0 (so validation passes and a
    delete is allowed); only an INSERT/UPDATE/DELETE statement flags a write."""
    def __init__(self):
        self.write_attempted = False

    class _DS(object):
        rowCount = 1
        def getValueAt(self, *a):
            return 0          # dup/overlap count = 0 (passes), refCount = 0 (deletable), SCOPE_IDENTITY = 0

    def runPrepQuery(self, *a, **k):
        sql = (a[0] if a else k.get("sql", "")) or ""
        if "INSERT" in sql or "DELETE" in sql:     # INSERT-with-SCOPE_IDENTITY counts as a write
            self.write_attempted = True
        return self._DS()

    def runPrepUpdate(self, *a, **k):
        self.write_attempted = True                # UPDATE / DELETE path
        return 1


class _GateLogger(object):
    def info(self, *a, **k): pass
    def warn(self, *a, **k): pass
    def error(self, *a, **k): pass


class _GateUtil(object):
    def getLogger(self, *a, **k): return _GateLogger()


class _GateSystem(object):
    def __init__(self, db):
        self.db = db
        self.util = _GateUtil()


class _Custom(object):
    """view.custom stand-in: the form props the script reads, plus recordId/runNonce/statusMsg/mayEdit
    (mayEdit is FORGEABLE here, modelling the client prop). Any attribute the script reads that we didn't
    pre-seed returns "" (the form-prop default) via __getattr__."""
    def __init__(self, fixture, record_id, may_edit):
        self.recordId = record_id
        self.runNonce = 0
        self.statusMsg = ""
        self.siteId = 1
        self.mayEdit = may_edit
        for k, v in fixture.items():
            setattr(self, k, v)

    def __getattr__(self, name):
        # unseen form fields default to "" (matches the custom-prop blank default the views ship)
        if name.startswith("form_"):
            return ""
        raise AttributeError(name)


class _View(object):
    def __init__(self, custom): self.custom = custom


class _Self(object):
    def __init__(self, view, session): self.view = view; self.session = session


def _register_auth_module():
    """Make `import auth as A` (inside the deployed scripts) resolve to the REAL auth module."""
    if "auth" not in sys.modules:
        sys.modules["auth"] = jython_shim.load_wrapper("auth", AUTH_CODE)
    return sys.modules["auth"]


def _deployed_script(view_name, button_name, base=GW_BASE):
    """Pull a button's onActionPerformed script from the DEPLOYED gateway view.json (default) or the repo
    copy. By component meta.name (verified unique in each view)."""
    path = os.path.join(base, view_name, view_name, "view.json")
    if not os.path.exists(path):
        return None
    view = json.load(open(path))

    def walk(node):
        if isinstance(node, dict):
            if node.get("meta", {}).get("name") == button_name and "events" in node:
                try:
                    return node["events"]["component"]["onActionPerformed"]["config"]["script"]
                except (KeyError, TypeError):
                    return None
            for v in node.values():
                r = walk(v)
                if r is not None:
                    return r
        elif isinstance(node, list):
            for it in node:
                r = walk(it)
                if r is not None:
                    return r
        return None
    return walk(view.get("root", {}))


def _compile_script(src):
    """Compile a deployed view button script (Jython-2 fragment) into a CPython3 callable. ONLY the
    syntactic Py2 `except X, e:` -> `except X as e:` shim; NO gate/validation semantics touched."""
    body = re.sub(r"except (\S+), e:", r"except \1 as e:", src)
    full = "def _f(self, system, unicode):\n" + body
    ns = {}
    exec(compile(full, "<master-script>", "exec"), ns)
    return ns["_f"]


def _run(script_src, session, fixture, may_edit, record_id):
    """Drive a deployed master Save/Delete script. Returns (write_attempted, statusMsg)."""
    _register_auth_module()
    fn = _compile_script(script_src)
    db = _GateDB()
    cust = _Custom(fixture, record_id=record_id, may_edit=may_edit)
    fn(_Self(_View(cust), session), _GateSystem(db), str)
    return db.write_attempted, cust.statusMsg


# ---------------------------------------------------------------------------
# 1. PER-VIEW WRITE GATE — forged-prop/anon rejected (Save AND Delete), authorized reaches the write.
# ---------------------------------------------------------------------------

def test_per_view_gate(rep):
    for V in VIEWS:
        fixture = VALID_FIXTURE[V]
        save_src = _deployed_script(V, "SaveButton")
        del_src = _deployed_script(V, "DeleteButton")
        repo_save = _deployed_script(V, "SaveButton", base=REPO_BASE)

        rep.check("%s: deployed Save+Delete scripts found AND deployed==committed (drives the gateway "
                  "view byte-for-byte)" % V,
                  save_src is not None and del_src is not None and save_src == repo_save,
                  "save=%s del=%s deployed==committed=%s"
                  % (save_src is not None, del_src is not None, save_src == repo_save))
        if save_src is None or del_src is None:
            rep.skip("%s write gate" % V, "deployed script(s) not found")
            continue

        # the gate is actually present (the injector did its job in the DEPLOYED script)
        rep.check("%s: deployed Save AND Delete both call the SERVER gate auth.requireWrite(self.session)"
                  % V,
                  "A.requireWrite(self.session)" in save_src and "A.requireWrite(self.session)" in del_src)

        # --- THE BLOCKER PROOF: forged mayEdit=true + a session with NO write role -> REJECTED ---
        for label, sess in (("anonymous (not logged in)", ANON_SESSION),
                            ("logged-in viewer with no write role", VIEWER_SESSION)):
            w, msg = _run(save_src, sess, fixture, may_edit=True, record_id=0)
            rep.check("%s SAVE by a %s is REJECTED server-side EVEN WITH a forged mayEdit=true "
                      "(NO db write)" % (V, label),
                      (w is False) and ("DENIED" in msg), "write=%s msg=%r" % (w, msg))
            w, msg = _run(del_src, sess, fixture, may_edit=True, record_id=7)
            rep.check("%s DELETE by a %s is REJECTED server-side EVEN WITH a forged mayEdit=true "
                      "(NO db write)" % (V, label),
                      (w is False) and ("DENIED" in msg), "write=%s msg=%r" % (w, msg))

        # --- NON-VACUITY: an authorized session (ProductionControl) SAVES + DELETES (gate not block-all) ---
        w, msg = _run(save_src, PC_SESSION, fixture, may_edit=True, record_id=0)
        rep.check("%s SAVE by a ProductionControl session SUCCEEDS (reaches the db write) — non-vacuous"
                  % V, w is True, "write=%s msg=%r" % (w, msg))
        w, msg = _run(del_src, PC_SESSION, fixture, may_edit=True, record_id=7)
        rep.check("%s DELETE by a ProductionControl session SUCCEEDS (reaches the db write) — non-vacuous"
                  % V, w is True, "write=%s msg=%r" % (w, msg))
        # Admin is also a writer (an administrator is never locked out of config edits)
        w, msg = _run(save_src, ADMIN_SESSION, fixture, may_edit=True, record_id=0)
        rep.check("%s SAVE by an Admin session SUCCEEDS (Admin is a write role too)" % V, w is True,
                  "write=%s msg=%r" % (w, msg))

        # --- THE SESSION (not the prop) is the gate: a PC session with mayEdit FALSE still writes ---
        w, msg = _run(save_src, PC_SESSION, fixture, may_edit=False, record_id=0)
        rep.check("%s: the SESSION (not mayEdit) gates — a ProductionControl session with mayEdit=FALSE "
                  "still SAVES" % V, w is True, "write=%s msg=%r" % (w, msg))


# ---------------------------------------------------------------------------
# 2. REVERT-PROOF — neuter the SERVER gate and confirm the forged-prop anon SAVE SLIPS THROUGH.
#    Proven on a representative view (the gate is the SAME module call in all 7).
# ---------------------------------------------------------------------------

def test_revert_proof(rep):
    A = _register_auth_module()
    orig = A.requireWrite
    proven = []
    try:
        A.requireWrite = lambda session: set()   # the hole, re-opened (no server-side check)
        for V in VIEWS:
            save_src = _deployed_script(V, "SaveButton")
            if save_src is None:
                continue
            w, msg = _run(save_src, ANON_SESSION, VALID_FIXTURE[V], may_edit=True, record_id=0)
            proven.append((V, w))
    finally:
        A.requireWrite = orig

    all_slipped = all(w is True for _, w in proven)
    rep.check("REVERT-PROOF: with the server gate removed, the FORGED-prop anon SAVE SLIPS THROUGH to the "
              "db write on ALL 7 views (so the gate, not the harness/prop, is load-bearing)",
              all_slipped and len(proven) == len(VIEWS),
              "results=%s" % ["%s:%s" % (v, w) for v, w in proven])

    # sanity: with the gate RESTORED, the same anon forged SAVE is denied again on a representative view
    w, msg = _run(_deployed_script("Size", "SaveButton"), ANON_SESSION, VALID_FIXTURE["Size"],
                  may_edit=True, record_id=0)
    rep.check("REVERT-PROOF restore: with the gate back, the anon forged SAVE is DENIED again (Size)",
              (w is False) and ("DENIED" in msg), "write=%s msg=%r" % (w, msg))


# ---------------------------------------------------------------------------
# 3. NO qaAdmin / no client-trust in any write path (defense-in-depth audit of the deployed scripts).
# ---------------------------------------------------------------------------

def test_no_client_trust_in_write_path(rep):
    for V in VIEWS:
        for btn in ("SaveButton", "DeleteButton"):
            src = _deployed_script(V, btn) or ""
            # the write must NOT branch on a client prop (mayEdit/qaAdmin) as its security gate; the only
            # authority is auth.requireWrite(self.session). Assert no qaAdmin hatch leaked into the write.
            no_qaadmin = "qaAdmin" not in src
            gates_on_session = "A.requireWrite(self.session)" in src
            rep.check("%s/%s: write path has NO qaAdmin hatch and gates on the SESSION (no client-trust)"
                      % (V, btn), no_qaadmin and gates_on_session,
                      "qaAdmin-absent=%s session-gated=%s" % (no_qaadmin, gates_on_session))


def main():
    print("=== P15 master write-gate sweep — server-side gate on the 7 master CRUD views ===\n")
    rep = Report()
    print("-- 1. Per-view write gate (forged-prop/anon rejected; authorized reaches the write) --")
    test_per_view_gate(rep)
    print("\n-- 2. Revert-proof (neuter the gate -> forged anon write slips through) --")
    test_revert_proof(rep)
    print("\n-- 3. No client-trust in any write path (no qaAdmin; gates on the session) --")
    test_no_client_trust_in_write_path(rep)
    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
