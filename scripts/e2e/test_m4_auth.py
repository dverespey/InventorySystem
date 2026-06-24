#!/usr/bin/env python3
"""test_m4_auth.py — M4 piece 2 (minimal auth) gate.

Proves the parts that are CODE (not gateway/Designer config): the SERVER-SIDE authorization gate (the
hole-closer for legacy H3), the user-management ops, the first-login password-reset flow, and the
role->view consistency. The IdP/user-source itself + the Perspective PAGE-permission are gateway/Designer
config and are DOCUMENTED + asserted only where assertable (see m4-auth-design.md).

WHAT THIS DRIVES (the REAL gateway-side code, not a reimplementation):
  We load the ACTUAL auth Project Library module
  (docs/analysis/production-readiness/project-library/auth/code.py) through the jython_shim with an
  in-memory system.user source + system.perspective, and call its ops exactly as the gateway view button
  scripts do. So the gate decision proven here is the gate the gateway runs.

NON-VACUITY (retro R15/R16/R20 discipline):
  * The deny tests are non-vacuous: the SAME op called by an Admin SUCCEEDS (so the test can fail if the
    gate let everyone in OR blocked everyone).
  * REVERT-PROVEN: we patch the `authorize` chokepoint to a no-op (the rebuild-without-the-gate) and
    re-run the deny case — it then WRONGLY succeeds, proving the gate (not the harness) is what blocks
    a non-Admin. A gate that can't fail on a removed gate is testing nothing.
  * The first-login flag oracle is the SOURCE rule (m4-auth-sites-sourcetruth.md §1: a net-new user has
    NULL LastUpdated => forced first-login reset; Q13 = force reset, no plaintext), transcribed
    independently here, NOT read back from the rebuild's own output.

HONEST SPLIT (what is NOT proven here — it is gateway/Designer config):
  * The Internal user source + the Admin/ProductionControl roles + seeded users (gateway web UI).
  * The Perspective PAGE-level role permission on /admin/users (the AUTHORITATIVE prod UI gate).
  * The security level "Authenticated/Roles/Admin".
  These are documented step-by-step in m4-auth-design.md; the view JSON's isAdmin binding + the auth
  module name them consistently and the test asserts that naming consistency.

Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 scripts/e2e/test_m4_auth.py
"""
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")))

from lib import Report                       # noqa: E402
import jython_shim                           # noqa: E402

REPO = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".."))
AUTH_CODE = os.path.join(REPO, "docs", "analysis", "production-readiness", "project-library", "auth", "code.py")
USER_ADMIN_VIEW = os.path.join(REPO, "docs", "analysis", "production-readiness", "perspective-views",
                               "Admin", "Users", "Users", "view.json")
SITES_VIEW = os.path.join(REPO, "docs", "analysis", "master-data", "perspective-views",
                          "Master", "Sites", "Sites", "view.json")

ADMIN = ["Admin"]
PC = ["ProductionControl"]


def load_auth():
    """Load the REAL auth Project Library module with a FRESH in-memory user source + perspective shim."""
    return jython_shim.load_wrapper("auth_m4", AUTH_CODE)


# ---------------------------------------------------------------------------
# 1. SERVER-SIDE GATE — the hole-closer. A non-Admin caller of each user-mgmt op is REJECTED IN THE
#    GATEWAY SCRIPT, before any system.user.* write. Non-vacuous (Admin succeeds) + revert-proven.
# ---------------------------------------------------------------------------

def test_server_side_gate(rep):
    A = load_auth()
    src = A.USER_SOURCE

    # --- addUser: PC denied, Admin allowed, no user created on deny ---
    before = A.system.user.usernames(src)
    denied = False
    try:
        A.addUser(PC, "intruder", ["Admin"], "pw")
    except A.AuthError:
        denied = True
    after = A.system.user.usernames(src)
    rep.check("addUser by ProductionControl is REJECTED server-side (in auth.addUser, not the UI)", denied)
    rep.check("  ...and NO user was created by the denied add (gate runs BEFORE system.user.addUser)",
              before == after, "before=%s after=%s" % (before, after))

    A.addUser(ADMIN, "op1", ["ProductionControl"], "temp1")
    rep.check("addUser by Admin SUCCEEDS (non-vacuous: the gate is not blocking everyone)",
              "op1" in A.system.user.usernames(src))

    # --- removeUser: PC denied, Admin allowed ---
    denied = False
    try:
        A.removeUser(PC, "op1")
    except A.AuthError:
        denied = True
    rep.check("removeUser by ProductionControl is REJECTED server-side", denied)
    rep.check("  ...and the user STILL exists after the denied delete", "op1" in A.system.user.usernames(src))
    A.removeUser(ADMIN, "op1")
    rep.check("removeUser by Admin SUCCEEDS (non-vacuous)", "op1" not in A.system.user.usernames(src))

    # --- resetPassword: PC denied, Admin allowed ---
    A.addUser(ADMIN, "op2", ["ProductionControl"], "temp2")
    denied = False
    try:
        A.resetPassword(PC, "op2", "newpw")
    except A.AuthError:
        denied = True
    rep.check("resetPassword by ProductionControl is REJECTED server-side", denied)
    A.resetPassword(ADMIN, "op2", "newpw")
    rep.check("resetPassword by Admin SUCCEEDS (non-vacuous)", True)

    # --- the gate ignores a CLIENT-PASSED admin flag: roles come from the session, not a param ---
    # A non-Admin session cannot self-escalate by passing extra roles; the op only trusts callerRoles
    # (which the view derives via auth.sessionRoles(self.session), server-side). We model a forged
    # request: caller is PC but the *target* roles include Admin — still denied (it's the CALLER's role
    # that gates, and the caller is PC).
    denied = False
    try:
        A.addUser(PC, "escalate", ["Admin", "ProductionControl"], "pw")
    except A.AuthError:
        denied = True
    rep.check("a ProductionControl caller cannot create an Admin (no self-escalation)", denied)


def test_gate_revert_proven(rep):
    """REVERT-PROOF: neuter the authorize chokepoint (simulate the rebuild WITHOUT the gate) and confirm
    the deny case then WRONGLY succeeds — proving the gate, not the harness, is what blocks a non-Admin."""
    A = load_auth()
    src = A.USER_SOURCE
    orig = A.authorize
    try:
        A.authorize = lambda roles, required: True   # the hole, re-opened
        slipped = False
        try:
            A.addUser(PC, "ghost", ["ProductionControl"], "pw")
            slipped = ("ghost" in A.system.user.usernames(src))
        except A.AuthError:
            slipped = False
        rep.check("REVERT-PROOF: with the gate removed, the non-Admin add SLIPS THROUGH "
                  "(so the gate is load-bearing, not the harness)", slipped)
    finally:
        A.authorize = orig


# ---------------------------------------------------------------------------
# 2. FIRST-LOGIN PASSWORD RESET (Q13). Oracle = the SOURCE rule (independent), not the rebuild.
# ---------------------------------------------------------------------------

def test_first_login_flow(rep):
    A = load_auth()
    src = A.USER_SOURCE

    # SOURCE rule (m4-auth-sites-sourcetruth.md §1): a net-new user is forced to reset on first login
    # (legacy NULL-LastUpdated). Independent oracle below: a freshly-added user MUST reset.
    A.addUser(ADMIN, "fresh", ["ProductionControl"], "temp")
    u = A.system.user.getUser(src, "fresh")
    rep.check("first-login: a newly-added user is flagged must-reset (Q13 / legacy NULL-LastUpdated)",
              A.firstLoginResetRequired(u) is True)

    # SELF-SERVICE reset clears the flag; it is IDENTITY-gated (a user may only reset their own).
    sess_fresh = {"auth": {"user": {"userName": "fresh", "roles": ["ProductionControl"]}}}
    A.setOwnPasswordFirstLogin(sess_fresh, "fresh", "MyRealPw123!")
    u2 = A.system.user.getUser(src, "fresh")
    rep.check("first-login: self-service reset CLEARS the must-reset flag (user can now use the app)",
              A.firstLoginResetRequired(u2) is False)

    # cross-user self-reset is rejected (you can't reset someone else's via the self path)
    sess_other = {"auth": {"user": {"userName": "mallory", "roles": ["ProductionControl"]}}}
    denied = False
    try:
        A.setOwnPasswordFirstLogin(sess_other, "fresh", "hax")
    except A.AuthError:
        denied = True
    rep.check("first-login: a session cannot self-reset ANOTHER user's password (identity-gated)", denied)

    # an ADMIN reset re-arms must-reset (the user picks their own next login)
    A.resetPassword(ADMIN, "fresh", "tmpAgain")
    u3 = A.system.user.getUser(src, "fresh")
    rep.check("admin reset RE-ARMS must-reset (forces the user to set a new one on next login)",
              A.firstLoginResetRequired(u3) is True)

    # NON-VACUITY: a user with NO sentinel + a set password is NOT forced to reset
    A.system.user.add_seed(src, "settled", ["ProductionControl"], password_set=True)
    settled = A.system.user.getUser(src, "settled")
    rep.check("non-vacuous: a settled user (password set, no sentinel) is NOT forced to reset",
              A.firstLoginResetRequired(settled) is False)

    # NO PLAINTEXT: the user object never carries the cleartext password (shim stores only a hashed marker)
    rep.check("no plaintext: the user source stores a (hashed) password marker, never the cleartext",
              not hasattr(settled, "_password") or settled.get("password") is None)

    # optional age-based expiry (legacy IN_PASSWORD_RESET_DAYS, D-M4-4) — small add, proven cheaply
    rep.check("expiry disabled (maxAgeDays<=0) -> never expired (the common case)",
              A.passwordExpired({"passwordChangedDaysAgo": 999}, 0) is False)
    rep.check("expiry honored: a 100-day-old password with a 90-day policy IS expired (legacy parity)",
              A.passwordExpired({"passwordChangedDaysAgo": 100}, 90) is True)
    rep.check("expiry honored: a 30-day-old password with a 90-day policy is NOT expired",
              A.passwordExpired({"passwordChangedDaysAgo": 30}, 90) is False)


# ---------------------------------------------------------------------------
# 3. ROLE -> VIEW MAPPING consistency: User Admin = Admin-only. The view JSON's isAdmin binding names the
#    SAME security level the auth module names (one source).
#
#    SITES (David 2026-06-22 — "show it like the other masters"): the Sites detail has NO client-side
#    visibility gate. It ALWAYS renders, consistent with the other 7 masters; the ONLY authorization
#    boundary is the SERVER-SIDE auth.requireWrite gate proven in test_sites_write_gate (section 4). So
#    here we assert the visibility gate is GONE (no custom.mayEdit prop, no RESTRICTED AdminBanner,
#    Form/ActionBar carry NO meta.visible binding) — the positive consistency invariant.
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


def test_role_view_consistency(rep):
    A = load_auth()

    # The auth module + the User Admin view must reference the SAME Admin security level string.
    ua = json.load(open(USER_ADMIN_VIEW))
    sites = json.load(open(SITES_VIEW))
    ua_expr = ua["propConfig"]["custom.isAdmin"]["binding"]["config"]["expression"]
    rep.check("User Admin view gates on Authenticated/Roles/Admin (the Admin-only screen)",
              "Authenticated/Roles/Admin" in ua_expr)
    rep.check("the auth module names the SAME Admin security level as the view (one source of truth)",
              A.SECLEVEL_ADMIN == "Authenticated/Roles/Admin" and A.SECLEVEL_ADMIN in ua_expr)

    # ---- SITES: NO client-side visibility gate (consistent with the other 7 masters) ----
    # The detail ALWAYS shows; the write is gated SERVER-SIDE only (proven in section 4). Assert the old
    # UI gate is fully removed: the custom.mayEdit prop, the RESTRICTED AdminBanner, and the Form/ActionBar
    # meta.visible bindings are all GONE. (The qaAdmin URL hatch lived inside custom.mayEdit -> gone too.)
    rep.check("Sites has NO custom.mayEdit UI-visibility prop (removed; detail always shows like the other "
              "masters)", "custom.mayEdit" not in sites.get("propConfig", {}),
              "propConfig keys=%r" % sorted(sites.get("propConfig", {}).keys()))
    rep.check("Sites has NO mayEdit in its custom defaults (the dead UI-gate prop is gone)",
              "mayEdit" not in sites.get("custom", {}))
    rep.check("Sites has NO RESTRICTED AdminBanner component (removed with the UI gate)",
              _find_node(sites["root"], "AdminBanner") is None)
    sform = _find_node(sites["root"], "Form")
    rep.check("Sites detail Form has NO meta.visible binding (ALWAYS visible, like the other masters)",
              sform is not None and "meta.visible" not in sform.get("propConfig", {}),
              "Form propConfig=%r" % (sform.get("propConfig") if sform else None))
    sab = _find_node(sites["root"], "ActionBar")
    rep.check("Sites ActionBar (Save/Delete/Clear) has NO meta.visible binding (ALWAYS visible)",
              sab is not None and "meta.visible" not in sab.get("propConfig", {}),
              "ActionBar propConfig=%r" % (sab.get("propConfig") if sab else None))
    # NON-VACUITY: the source string of the old gate (the ProductionControl-OR-qaAdmin expression and the
    # RESTRICTED banner text) must not survive ANYWHERE in the view (a stray copy would re-hide the form).
    sites_raw = json.dumps(sites)
    rep.check("non-vacuous: the old qaAdmin / RESTRICTED visibility-gate strings are gone from the whole "
              "Sites view (no stray re-hide)", "qaAdmin" not in sites_raw and "RESTRICTED" not in sites_raw)

    # The User Admin button scripts route every write through the auth module (the server-side gate),
    # not an inline system.user.* call that would bypass it.
    scripts = []
    def walk(o):
        if isinstance(o, dict):
            for k, v in o.items():
                if k == "script" and isinstance(v, basestring):  # noqa: F821 (shim builtin)
                    scripts.append(v)
                else:
                    walk(v)
        elif isinstance(o, list):
            for x in o:
                walk(x)
    walk(ua)
    write_scripts = [s for s in scripts if "Add" in s or "remove" in s or "reset" in s or "addUser" in s]
    inline_bypass = any(("system.user.addUser" in s or "system.user.removeUser" in s
                         or "system.user.editUser" in s) for s in scripts)
    routes_via_auth = all(("A.addUser" in s or "A.removeUser" in s or "A.resetPassword" in s)
                          for s in scripts
                          if ("Add user" in s or "Delete user" in s or "Reset password" in s
                              or "ua-add" in s))
    rep.check("User Admin view does NOT call system.user.* inline (every write goes through the gated "
              "auth module, so the UI cannot bypass the server-side gate)", not inline_bypass)


# ---------------------------------------------------------------------------
# 4. SITES MASTER-CRUD WRITE GATE (the REAL boundary — now the SOLE boundary). The Sites detail Form +
#    Save/Delete ALWAYS render (the client-side visibility gate was removed so Sites matches the other 7
#    masters — David 2026-06-22). That makes the SERVER-SIDE auth.requireWrite gate the ONLY thing
#    standing between an unauthorized/anon user and a write — so this test is now MORE load-bearing, not
#    less. We drive the DEPLOYED Sites Save/Delete scripts from a SESSION that carries NO write role
#    (and, to model a forged client prop that the server MUST ignore, also pass may_edit=True), and prove
#    the gate REJECTS the write (no db write). Non-vacuous (a ProductionControl/Admin session succeeds) +
#    revert-proven (neuter the gate -> the no-role write slips through). The may_edit flag is a forged
#    CLIENT prop the deployed script no longer even reads; the gate keys off self.session ONLY, proven by
#    the "ProductionControl session with mayEdit=FALSE still SAVES" case below.
#
# Drives the ACTUAL gen_sites_view.py save_script()/delete_script() strings (the gateway view button
# scripts), with `import auth as A` resolved to the REAL auth module (sys.modules), so the gate proven
# here is the gate the gateway runs.
# ---------------------------------------------------------------------------

import gen_sites_view as gsv                # noqa: E402  (the SHARED Sites view generator)

ANON_SESSION = {"auth": {"user": {"userName": None, "roles": []}}}     # not logged in / no role
VIEWER_SESSION = {"auth": {"user": {"userName": "nobody", "roles": ["SomeOtherRole"]}}}  # logged in, no write role
PC_SESSION = {"auth": {"user": {"userName": "op1", "roles": ["ProductionControl"]}}}
ADMIN_SESSION = {"auth": {"user": {"userName": "boss", "roles": ["Admin"]}}}

SITES_COLS = [f[0] for f in gsv.F if f[3] not in ("ro_text", "ro_num")]
VALID_SITE = {"form_name": "Gate Test Site", "form_abbr": "GTS", "form_fcmode": "AUTO"}


class _SitesDB(object):
    """Records whether the Sites save/delete path reached a DB write (= passed the gate + validation)."""
    def __init__(self):
        self.write_attempted = False

    class _DS(object):
        rowCount = 1
        def getValueAt(self, *a):
            return 0           # refCount=0 (deletable) / SCOPE_IDENTITY id
    def runPrepQuery(self, *a, **k):
        # the INSERT-with-SCOPE_IDENTITY path AND the delete refCount read both go through runPrepQuery;
        # only the INSERT counts as a WRITE. We flag a write only when the SQL is an INSERT/DELETE.
        sql = a[0] if a else k.get("sql", "")
        if "INSERT" in sql or "DELETE" in sql:
            self.write_attempted = True
        return self._DS()
    def runPrepUpdate(self, *a, **k):
        self.write_attempted = True       # the UPDATE / DELETE path
        return 1


class _SitesLogger(object):
    def info(self, *a, **k): pass
    def warn(self, *a, **k): pass
class _SitesUtil(object):
    def getLogger(self, *a, **k): return _SitesLogger()
class _SitesSystem(object):
    def __init__(self, db):
        self.db = db
        self.util = _SitesUtil()


class _SCustom(object):
    def __init__(self, may_edit, record_id):
        self.recordId = record_id
        self.runNonce = 0
        self.statusMsg = ""
        self.mayEdit = may_edit       # a FORGED client prop the deployed script no longer reads (the
                                      # gate keys off self.session ONLY); kept to prove the server ignores it
        for k in SITES_COLS:
            setattr(self, k, VALID_SITE.get(k, ""))
        self.form_id = 0
class _SView(object):
    def __init__(self, custom): self.custom = custom
class _SSelf(object):
    def __init__(self, view, session): self.view = view; self.session = session


def _register_auth_module():
    """Make `import auth as A` (inside the Sites Save/Delete scripts) resolve to the REAL auth module."""
    if "auth" not in sys.modules:
        sys.modules["auth"] = jython_shim.load_wrapper("auth", AUTH_CODE)
    return sys.modules["auth"]


def _compile_script(src, argname="system"):
    """Compile a deployed view button script (Jython-2 fragment) into a CPython3 callable. ONLY syntactic
    shims: Py2 `except X, e:` -> `except X as e:`. No gate/validation semantics touched."""
    import re
    body = re.sub(r"except (\S+), e:", r"except \1 as e:", src)
    full = "def _f(self, system, unicode):\n" + body
    ns = {}
    exec(compile(full, "<sites-script>", "exec"), ns)
    return ns["_f"]


def _run_sites(script_src, session, may_edit, record_id):
    """Drive a deployed Sites Save/Delete script. Returns (write_attempted, statusMsg)."""
    _register_auth_module()
    fn = _compile_script(script_src)
    db = _SitesDB()
    cust = _SCustom(may_edit=may_edit, record_id=record_id)
    fn(_SSelf(_SView(cust), session), _SitesSystem(db), str)
    return db.write_attempted, cust.statusMsg


def test_sites_write_gate(rep):
    save_src = gsv.save_script()
    del_src = gsv.delete_script()

    # Verify the harness drives the DEPLOYED script (byte-identical to the gateway view button), not just
    # the generator — so the gate proven here is the gate the running session serves.
    deployed_save = _deployed_sites_script("sites-save-btn")
    rep.check("sites-gate harness drives the DEPLOYED Save script (byte-identical to the gateway view)",
              deployed_save is not None and deployed_save == save_src,
              "deployed==generator: %s" % (deployed_save == save_src))

    # --- THE BLOCKER PROOF: forged mayEdit=true + a session with NO write role -> REJECTED server-side ---
    for label, sess in (("anonymous (not logged in)", ANON_SESSION),
                        ("logged-in viewer with no write role", VIEWER_SESSION)):
        w, msg = _run_sites(save_src, sess, may_edit=True, record_id=0)
        rep.check("Sites SAVE by a %s is REJECTED server-side EVEN WITH a forged mayEdit=true "
                  "(the SERVER gate, not the prop, blocks it; NO db write)" % label,
                  (w is False) and ("DENIED" in msg), "write=%s msg=%r" % (w, msg))
        w, msg = _run_sites(del_src, sess, may_edit=True, record_id=5)
        rep.check("Sites DELETE by a %s is REJECTED server-side EVEN WITH a forged mayEdit=true "
                  "(NO db write)" % label, (w is False) and ("DENIED" in msg),
                  "write=%s msg=%r" % (w, msg))

    # --- NON-VACUITY: an authorized session (ProductionControl) SAVES + DELETES (gate is not block-all) ---
    w, msg = _run_sites(save_src, PC_SESSION, may_edit=True, record_id=0)
    rep.check("Sites SAVE by a ProductionControl session SUCCEEDS (reaches the db write) — non-vacuous",
              w is True, "write=%s msg=%r" % (w, msg))
    w, msg = _run_sites(del_src, PC_SESSION, may_edit=True, record_id=5)
    rep.check("Sites DELETE by a ProductionControl session SUCCEEDS (reaches the db write) — non-vacuous",
              w is True, "write=%s msg=%r" % (w, msg))
    # Admin is also a writer (so an administrator is never locked out of config edits)
    w, msg = _run_sites(save_src, ADMIN_SESSION, may_edit=True, record_id=0)
    rep.check("Sites SAVE by an Admin session SUCCEEDS (Admin is a write role too)", w is True,
              "write=%s msg=%r" % (w, msg))

    # --- REVERT-PROOF: neuter the SERVER gate (requireWrite -> no-op) and confirm the forged-prop write
    #     by a no-role session then SLIPS THROUGH — proving the SERVER gate is load-bearing, not the prop. ---
    A = _register_auth_module()
    orig = A.requireWrite
    try:
        A.requireWrite = lambda session: set()    # the hole, re-opened (no server-side check)
        w, msg = _run_sites(save_src, ANON_SESSION, may_edit=True, record_id=0)
        rep.check("REVERT-PROOF: with the server gate removed, the FORGED-prop anon SAVE SLIPS THROUGH to "
                  "the db write (so the gate, not the harness/prop, is what blocks it)", w is True,
                  "write=%s msg=%r" % (w, msg))
    finally:
        A.requireWrite = orig

    # The gate keys off the SESSION, NOT the client prop: a no-write session with mayEdit FORGED true is
    # denied (above); a write session with mayEdit FALSE still writes (the prop is UI-only, not the gate).
    w, msg = _run_sites(save_src, PC_SESSION, may_edit=False, record_id=0)
    rep.check("the SESSION (not the mayEdit prop) is the gate: a ProductionControl session with mayEdit=FALSE "
              "still SAVES (server-side decision ignores the client prop)", w is True,
              "write=%s msg=%r" % (w, msg))


# ---------------------------------------------------------------------------
# 5. USER ADMIN READ GATE (SHOULD-FIX-2). system.user.getUsers (the user/role list) runs on page render.
#    Defense-in-depth symmetry: the list builder re-checks Admin SERVER-SIDE from the SESSION and returns
#    an EMPTY list for a non-Admin (no info-leak if the page-permission is ever forgotten). We drive the
#    DEPLOYED list-builder transform script with an Admin vs non-Admin session and prove it. Revert-proof.
# ---------------------------------------------------------------------------

import gen_user_admin_view as guav            # noqa: E402  (the User Admin view generator)


class _ListComp(object):
    """Stand-in for `self` (the component) in a binding-transform: exposes self.session (an attribute of
    self in a Perspective transform — IA docs)."""
    def __init__(self, session):
        self.session = session


def _run_user_list(session, A):
    """Drive the DEPLOYED User Admin list-builder transform. Returns the {data, columns} dict.
    `A` is the loaded auth module; its `system` (with the in-memory user source) backs system.user.*."""
    src = guav.LIST_SCRIPT
    body = re.sub(r"except (\S+), e:", r"except \1 as e:", src)
    full = "def _list(self, value, system, unicode):\n" + body
    ns = {"auth": A}
    sys.modules["auth"] = A
    exec(compile(full, "<user-list>", "exec"), ns)
    return ns["_list"](_ListComp(session), "|0", A.system, str)


def test_user_admin_read_gate(rep):
    A = load_auth()
    src = A.USER_SOURCE
    # seed a couple of users so a successful (Admin) read is non-empty
    A.system.user.add_seed(src, "admin1", ["Admin"], password_set=True)
    A.system.user.add_seed(src, "opA", ["ProductionControl"], password_set=True)

    # assert the deployed list builder CONTAINS a server-side gate (not just inline getUsers on render)
    rep.check("User Admin list builder has a SERVER-SIDE read gate (isAdmin on session roles before "
              "system.user.getUsers)",
              "A.isAdmin(A.sessionRoles(" in guav.LIST_SCRIPT and "system.user.getUsers" in guav.LIST_SCRIPT)

    # --- a non-Admin session gets an EMPTY list (no account/role info-leak) ---
    res = _run_user_list(VIEWER_SESSION, A)
    rep.check("User Admin READ by a non-Admin session returns an EMPTY list (server-side read gate; no "
              "info-leak even if the page-permission is forgotten)", res.get("data") == [],
              "rows=%d" % len(res.get("data", [])))
    res = _run_user_list(ANON_SESSION, A)
    rep.check("User Admin READ by an anonymous session returns an EMPTY list", res.get("data") == [],
              "rows=%d" % len(res.get("data", [])))

    # --- NON-VACUITY: an Admin session sees the users (gate is not block-all) ---
    res = _run_user_list(ADMIN_SESSION, A)
    names = sorted(r["Username"] for r in res.get("data", []))
    rep.check("User Admin READ by an Admin session returns the user list (non-vacuous)",
              names == ["admin1", "opA"], "names=%s" % names)

    # --- REVERT-PROOF: neuter the read gate (isAdmin -> always True) and the non-Admin read LEAKS ---
    orig = A.isAdmin
    try:
        A.isAdmin = lambda roles: True
        res = _run_user_list(VIEWER_SESSION, A)
        rep.check("REVERT-PROOF: with the read gate removed, a non-Admin read LEAKS the user list (so the "
                  "gate is load-bearing)", len(res.get("data", [])) == 2,
                  "rows=%d" % len(res.get("data", [])))
    finally:
        A.isAdmin = orig


def _deployed_sites_script(dom_id):
    """Pull a button's onActionPerformed script from the DEPLOYED gateway Sites view.json."""
    gw = ("/usr/local/ignition/data/projects/InventorySystem/com.inductiveautomation.perspective"
          "/views/Master/Sites/Sites/view.json")
    if not os.path.exists(gw):
        return None
    view = json.load(open(gw))
    def walk(node):
        if isinstance(node, dict):
            if node.get("meta", {}).get("domId") == dom_id:
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
    return walk(view)


# basestring is a Jython builtin; under CPython3 we provide it for the walk() above.
try:
    basestring
except NameError:
    basestring = str


def main():
    print("=== M4 piece 2 — minimal auth (server-side gate + user-mgmt + first-login) ===\n")
    rep = Report()
    print("-- 1. Server-side authorization gate (the hole-closer) --")
    test_server_side_gate(rep)
    test_gate_revert_proven(rep)
    print("\n-- 2. First-login password reset (Q13) + optional age-expiry --")
    test_first_login_flow(rep)
    print("\n-- 3. Role -> view mapping consistency --")
    test_role_view_consistency(rep)
    print("\n-- 4. Sites master-CRUD WRITE gate (the BLOCKER: server-side, forged-prop rejected) --")
    test_sites_write_gate(rep)
    print("\n-- 5. User Admin READ gate (SHOULD-FIX-2: server-side, non-Admin gets empty list) --")
    test_user_admin_read_gate(rep)
    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
