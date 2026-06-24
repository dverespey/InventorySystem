#!/usr/bin/env python3
"""Generate the User Admin Perspective view.json (M4 piece 2 — minimal auth).

The ONE screen Admin can use that ProductionControl cannot: add / delete / reset users in the gateway
Internal user source. Follows the PROVEN canonical master pattern (gen_sites_view.py): single combined
view (list left, action form right), in-view custom.isAdmin gate (defense-in-depth), same-view recordId
prop write (no navigation), SPIKE logging.

THE SECURITY BOUNDARY IS NOT THIS VIEW. Every write button calls the `auth` Project Library module
(auth.addUser / auth.removeUser / auth.resetPassword), which re-checks Admin SERVER-SIDE from the SESSION
roles (session.props.auth.user.roles) before touching system.user.*. Hiding the form (isAdmin) is
defense-in-depth only; the enforced gate is in auth/code.py (the hole-closer for legacy H3). So even a
forged request that reaches a button script is rejected in the gateway by auth.authorize().

DIFFERS from the data masters: there is NO database table here — users live in the gateway Internal user
source, read/written via system.user.* (8.1+ API verified vs IA docs). The list is built from
system.user.getUsers(source); the detail form add/delete/reset call the auth module.

HEADLESS-VS-DESIGNER (honest):
  * BUILDABLE HEADLESS (this file + auth/code.py): the view JSON, the button scripts that call the
    server-side-gated auth ops, the isAdmin gate binding, the first-login reset wiring.
  * GATEWAY-CONFIG (Designer/gateway web UI, NOT headless — documented in m4-auth-design.md): the
    Internal user source itself + the Admin/ProductionControl roles + the seeded users + the Perspective
    PAGE-level role permission on /admin/users (the AUTHORITATIVE prod gate) + the security level
    "Authenticated/Roles/Admin". The qaAdmin URL hatch (below) is a SPIKE-ONLY escape for the harness;
    in prod the page is role-gated so the hatch is unreachable by non-admins. IG83-TODO: drop qaAdmin.

Writes view.json to BOTH the deployed gateway path AND the committed repo path (runtime == reviewable),
plus the committed resource.json.
"""
import json
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _ignenv import PERSPECTIVE_DIR   # noqa: E402 — centralized gateway path (repo-split-plan §4.C)

GW_OUT = os.path.join(PERSPECTIVE_DIR, "views", "Admin", "Users", "Users", "view.json")
_REPO = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
REPO_OUT = os.path.join(_REPO, "ignition", "perspective-views",
                        "Admin", "Users", "Users", "view.json")
REPO_RESOURCE = os.path.join(os.path.dirname(REPO_OUT), "resource.json")
GW_RESOURCE = os.path.join(os.path.dirname(GW_OUT), "resource.json")
# REPO copy omits `attributes` (the gateway re-signs on load). The GATEWAY copy MUST carry an
# `attributes` block: on a COLD start the gateway's ProjectFileTree.toResource() calls setAttributes()
# and a missing/null `attributes` map throws an NPE that FAULTS THE WHOLE GATEWAY (observed 2026-06-22:
# ProjectResourceBuilder.setAttributes NPE during context startup). The signature is a placeholder the
# gateway overwrites + re-signs on first load (logged "Setting LastModification to external").
RESOURCE_JSON = {"scope": "G", "version": 1, "restricted": False, "overridable": True, "files": ["view.json"]}
GW_RESOURCE_JSON = dict(RESOURCE_JSON, attributes={
    "lastModification": {"actor": "external", "timestamp": "2026-06-22T22:52:21Z"},
    "lastModificationSignature": "0" * 64,
})

USER_SOURCE = "Inventory"   # must match auth.USER_SOURCE + the gateway-config user source name

ROLE_OPTIONS = [
    {"value": "ProductionControl", "label": "ProductionControl (everything in the app)"},
    {"value": "Admin", "label": "Admin (user add/delete/reset ONLY)"},
]


# ---------------------------------------------------------------------------
# The list builder: system.user.getUsers(source) -> {data, columns}. Server-side; the list itself is
# readable by the page (which is Admin-gated), and shows username/roles/must-reset. (No DB.)
# ---------------------------------------------------------------------------
LIST_SCRIPT = r'''	# ---- User Admin list — system.user.getUsers (gateway Internal source) ----
	# IG81-COMPAT: system.user.getUsers is 8.1+ and unchanged on 8.3.
	# SHOULD-FIX-2 (server-side READ gate, defense-in-depth symmetry): the user/role list is sensitive
	# (it enumerates accounts + roles), so we re-check Admin SERVER-SIDE here before reading the user
	# source — not only via the page permission. If the page-permission is ever forgotten, this still
	# fails closed: a non-Admin session gets an EMPTY list (no info-leak), the gate logged. The roles
	# come from the SESSION (gateway-side), never a client prop.
	log = system.util.getLogger("SPIKE")
	import auth as A
	# self.session is an attribute of `self` in a Perspective binding transform (it is not placed in the
	# transform's locals() but IS reachable on self — IA docs). Resolve roles SERVER-SIDE from it; fail
	# closed (empty list) if it can't be resolved or the caller is not Admin.
	sess = getattr(self, "session", None)
	if not A.isAdmin(A.sessionRoles(sess)):
		log.warn("SPIKE UserAdmin list DENIED (server-side read gate): non-Admin session")
		return {"data": [], "columns": [
			{"field": "Username", "header": {"title": "Username"}},
			{"field": "Roles", "header": {"title": "Roles"}},
			{"field": "MustReset", "header": {"title": "Must Reset"}}]}
	search = unicode(value).split("|")[0].strip().lower()
	users = system.user.getUsers(A.USER_SOURCE)
	columns = [
		{"field": "Username",  "header": {"title": "Username"}, "sort": "ascending"},
		{"field": "Roles",     "header": {"title": "Roles"}},
		{"field": "MustReset", "header": {"title": "Must Reset"}, "strictWidth": 110},
	]
	rows = []
	for u in users:
		# username: PyUser exposes it as the well-known property (user.get(user.Username) on 8.1) with
		# getUsername() as the convenience accessor; try both for portability.
		try:
			name = u.getUsername()
		except Exception:
			name = u.get("username")
		try:
			roleList = list(u.getRoles())   # 8.1: sequence of role-name strings
		except Exception:
			roleList = []
		mustReset = A.firstLoginResetRequired(u)
		if search and search not in unicode(name).lower():
			continue
		rows.append({"Username": name, "Roles": ", ".join(sorted(roleList)),
		             "MustReset": "YES" if mustReset else ""})
	log.info("SPIKE UserAdmin list: %d users (search='%s')" % (len(rows), search))
	return {"data": sorted(rows, key=lambda r: r["Username"]), "columns": columns}
'''


def add_script():
    return (
        "\t# ---- Add user — SERVER-SIDE Admin gate in auth.addUser ----\n"
        "\t# The boundary is auth.addUser (re-checks Admin from the SESSION); isAdmin here is\n"
        "\t# defense-in-depth UI hiding only. No plaintext stored (Internal source hashes).\n"
        "\tlog = system.util.getLogger(\"SPIKE\")\n"
        "\timport auth as A\n"
        "\tc = self.view.custom\n"
        "\tcallerRoles = A.sessionRoles(self.session)\n"
        "\tusername = (unicode(c.form_username) if c.form_username is not None else u\"\").strip()\n"
        "\tpassword = unicode(c.form_password) if c.form_password is not None else u\"\"\n"
        "\trole = unicode(c.form_role) if c.form_role else u\"ProductionControl\"\n"
        "\tif username == \"\":\n"
        "\t\tc.statusMsg = \"Username is required.\"; return\n"
        "\tif password == \"\":\n"
        "\t\tc.statusMsg = \"A temporary password is required (the user must reset it on first login).\"; return\n"
        "\ttry:\n"
        "\t\tmsg = A.addUser(callerRoles, username, [role], password, forceReset=True)\n"
        "\t\tc.statusMsg = \"Added %s (%s). User must reset password on first login. [%s]\" % (username, role, msg)\n"
        "\t\tc.form_username = \"\"; c.form_password = \"\"\n"
        "\t\tc.runNonce = (c.runNonce or 0) + 1\n"
        "\t\tlog.info(\"SPIKE UserAdmin ADD ok: %s role=%s\" % (username, role))\n"
        "\texcept A.AuthError, e:\n"
        "\t\tc.statusMsg = \"DENIED (server-side): %s\" % unicode(e)\n"
        "\t\tlog.warn(\"SPIKE UserAdmin ADD DENIED: %s\" % unicode(e))\n"
        "\texcept Exception, e:\n"
        "\t\tc.statusMsg = \"Add failed: %s\" % unicode(e)\n"
        "\t\tlog.warn(\"SPIKE UserAdmin ADD FAILED: %s\" % unicode(e))\n"
    )


def delete_script():
    return (
        "\t# ---- Delete user — SERVER-SIDE Admin gate in auth.removeUser ----\n"
        "\tlog = system.util.getLogger(\"SPIKE\")\n"
        "\timport auth as A\n"
        "\tc = self.view.custom\n"
        "\tcallerRoles = A.sessionRoles(self.session)\n"
        "\tusername = (unicode(c.selectedUser) if c.selectedUser is not None else u\"\").strip()\n"
        "\tif username == \"\":\n"
        "\t\tc.statusMsg = \"Select a user to delete.\"; return\n"
        "\tmyName = A.sessionUsername(self.session)\n"
        "\tif myName is not None and unicode(myName) == username:\n"
        "\t\tc.statusMsg = \"You cannot delete your own account while logged in.\"; return\n"
        "\ttry:\n"
        "\t\tmsg = A.removeUser(callerRoles, username)\n"
        "\t\tc.statusMsg = \"Deleted %s. [%s]\" % (username, msg)\n"
        "\t\tc.selectedUser = \"\"\n"
        "\t\tc.runNonce = (c.runNonce or 0) + 1\n"
        "\t\tlog.info(\"SPIKE UserAdmin DELETE ok: %s\" % username)\n"
        "\texcept A.AuthError, e:\n"
        "\t\tc.statusMsg = \"DENIED (server-side): %s\" % unicode(e)\n"
        "\t\tlog.warn(\"SPIKE UserAdmin DELETE DENIED: %s\" % unicode(e))\n"
        "\texcept Exception, e:\n"
        "\t\tc.statusMsg = \"Delete failed: %s\" % unicode(e)\n"
        "\t\tlog.warn(\"SPIKE UserAdmin DELETE FAILED: %s\" % unicode(e))\n"
    )


def reset_script():
    return (
        "\t# ---- Reset password — SERVER-SIDE Admin gate in auth.resetPassword ----\n"
        "\t# Admin reset RE-ARMS must-reset: the user picks their own new password on next login.\n"
        "\tlog = system.util.getLogger(\"SPIKE\")\n"
        "\timport auth as A\n"
        "\tc = self.view.custom\n"
        "\tcallerRoles = A.sessionRoles(self.session)\n"
        "\tusername = (unicode(c.selectedUser) if c.selectedUser is not None else u\"\").strip()\n"
        "\ttmp = unicode(c.form_resetpw) if c.form_resetpw is not None else u\"\"\n"
        "\tif username == \"\":\n"
        "\t\tc.statusMsg = \"Select a user to reset.\"; return\n"
        "\tif tmp == \"\":\n"
        "\t\tc.statusMsg = \"Enter a temporary password for the reset.\"; return\n"
        "\ttry:\n"
        "\t\tmsg = A.resetPassword(callerRoles, username, tmp)\n"
        "\t\tc.statusMsg = \"Reset %s. User must set a new password on next login. [%s]\" % (username, msg)\n"
        "\t\tc.form_resetpw = \"\"\n"
        "\t\tc.runNonce = (c.runNonce or 0) + 1\n"
        "\t\tlog.info(\"SPIKE UserAdmin RESET ok: %s\" % username)\n"
        "\texcept A.AuthError, e:\n"
        "\t\tc.statusMsg = \"DENIED (server-side): %s\" % unicode(e)\n"
        "\t\tlog.warn(\"SPIKE UserAdmin RESET DENIED: %s\" % unicode(e))\n"
        "\texcept Exception, e:\n"
        "\t\tc.statusMsg = \"Reset failed: %s\" % unicode(e)\n"
        "\t\tlog.warn(\"SPIKE UserAdmin RESET FAILED: %s\" % unicode(e))\n"
    )


def _input_row(key, label, comp, vprop, extra_props=None):
    field = {
        "meta": {"domId": "ua-" + key.replace("form_", ""), "name": "fld_" + key},
        "position": {"basis": "300px"},
        "propConfig": {vprop: {"binding": {"config": {"bidirectional": True,
                       "path": "view.custom.%s" % key}, "type": "property"}}},
        "type": comp,
    }
    if extra_props:
        field["props"] = extra_props
    lbl = {"meta": {"name": "lbl_" + key},
           "position": {"basis": "180px"},
           "props": {"style": {"fontSize": "12px"}, "text": label},
           "type": "ia.display.label"}
    return {"meta": {"name": "row_" + key}, "props": {"alignItems": "center", "style": {"gap": "10px"}},
            "type": "ia.container.flex", "children": [lbl, field]}


def build_view():
    role_dropdown = {
        "meta": {"domId": "ua-role", "name": "fld_form_role"},
        "position": {"basis": "300px"},
        "props": {"options": ROLE_OPTIONS},
        "propConfig": {"props.value": {"binding": {"config": {"bidirectional": True,
                       "path": "view.custom.form_role"}, "type": "property"}}},
        "type": "ia.input.dropdown",
    }
    role_row = {"meta": {"name": "row_form_role"},
                "props": {"alignItems": "center", "style": {"gap": "10px"}},
                "type": "ia.container.flex",
                "children": [{"meta": {"name": "lbl_form_role"}, "position": {"basis": "180px"},
                              "props": {"style": {"fontSize": "12px"}, "text": "Role *"},
                              "type": "ia.display.label"}, role_dropdown]}

    add_form_children = [
        {"meta": {"name": "grp_add"},
         "props": {"style": {"color": "#1976D2", "fontSize": "12px", "fontWeight": "bold",
                             "borderBottom": "1px solid #D0D4DA"}, "text": "Add User"},
         "type": "ia.display.label"},
        _input_row("form_username", "Username *", "ia.input.text-field", "props.text", {"maxLength": 30}),
        role_row,
        _input_row("form_password", "Temp Password *", "ia.input.text-field", "props.text",
                   {"maxLength": 50, "type": "password"}),
        {"meta": {"name": "AddBtnRow"}, "props": {"style": {"gap": "10px", "marginTop": "4px"}},
         "type": "ia.container.flex",
         "children": [
             {"events": {"component": {"onActionPerformed": {"config": {"script": add_script()},
                "scope": "G", "type": "script"}}},
              "meta": {"domId": "ua-add-btn", "name": "AddButton"},
              "props": {"style": {"backgroundColor": "#388E3C", "color": "#FFFFFF", "fontWeight": "bold"},
                        "text": "Add User"}, "type": "ia.input.button"},
         ]},
    ]

    manage_children = [
        {"meta": {"name": "grp_manage"},
         "props": {"style": {"color": "#1976D2", "fontSize": "12px", "fontWeight": "bold",
                             "marginTop": "10px", "borderBottom": "1px solid #D0D4DA"},
                   "text": "Manage Selected User"},
         "type": "ia.display.label"},
        {"meta": {"name": "row_selected"}, "props": {"alignItems": "center", "style": {"gap": "10px"}},
         "type": "ia.container.flex",
         "children": [
             {"meta": {"name": "lbl_selected"}, "position": {"basis": "180px"},
              "props": {"style": {"fontSize": "12px"}, "text": "Selected"}, "type": "ia.display.label"},
             {"meta": {"domId": "ua-selected", "name": "SelectedUser"},
              "propConfig": {"props.text": {"binding": {"config": {
                  "path": "view.custom.selectedUser"}, "type": "property"}}},
              "props": {"style": {"fontWeight": "bold", "fontSize": "13px"}}, "type": "ia.display.label"},
         ]},
        _input_row("form_resetpw", "Temp Password", "ia.input.text-field", "props.text",
                   {"maxLength": 50, "type": "password"}),
        {"meta": {"name": "ManageBtnRow"}, "props": {"style": {"gap": "10px", "marginTop": "4px"}},
         "type": "ia.container.flex",
         "children": [
             {"events": {"component": {"onActionPerformed": {"config": {"script": reset_script()},
                "scope": "G", "type": "script"}}},
              "meta": {"domId": "ua-reset-btn", "name": "ResetButton"},
              "props": {"style": {"backgroundColor": "#1976D2", "color": "#FFFFFF", "fontWeight": "bold"},
                        "text": "Reset Password"}, "type": "ia.input.button"},
             {"events": {"component": {"onActionPerformed": {"config": {"script": delete_script()},
                "scope": "G", "type": "script"}}},
              "meta": {"domId": "ua-delete-btn", "name": "DeleteButton"},
              "props": {"style": {"backgroundColor": "#C62828", "color": "#FFFFFF", "fontWeight": "bold"},
                        "text": "Delete User"}, "type": "ia.input.button"},
         ]},
    ]

    view = {
        "custom": {
            "runNonce": 0, "searchTerm": "", "statusMsg": "", "isAdmin": False,
            "selectedUser": "", "form_username": "", "form_password": "",
            "form_role": "ProductionControl", "form_resetpw": "",
        },
        "params": {},
        "propConfig": {
            "custom.userList": {
                "binding": {
                    "config": {"expression": "{view.custom.searchTerm} + '|' + {view.custom.runNonce}"},
                    "transforms": [{"code": LIST_SCRIPT, "type": "script"}],
                    "type": "expr",
                }
            },
            # AUTHORITATIVE prod gate = isAuthorized(false,"Authenticated/Roles/Admin") (needs IdP +
            # the Admin security level — fails CLOSED in this headless spike). qaAdmin is the SPIKE-ONLY
            # harness hatch; in prod the PAGE is role-gated so the hatch is unreachable. IG83-TODO: drop it.
            "custom.isAdmin": {
                "binding": {"config": {
                    "expression": "isAuthorized(false, \"Authenticated/Roles/Admin\") || {page.props.urlParams.qaAdmin} = \"1\""},
                    "type": "expr"}
            },
        },
        "props": {"defaultSize": {"height": 800, "width": 1280}},
        "root": {
            "meta": {"name": "root"},
            "props": {"style": {"backgroundColor": "#F5F6F8", "gap": "12px", "padding": "12px"}},
            "type": "ia.container.flex",
            "children": [
                # LEFT: user list
                {
                    "meta": {"name": "LeftPane"},
                    "position": {"basis": "520px", "grow": 1},
                    "props": {"direction": "column", "style": {"gap": "10px"}},
                    "type": "ia.container.flex",
                    "children": [
                        {"meta": {"name": "Title"},
                         "props": {"style": {"color": "#222222", "fontSize": "15px", "fontWeight": "bold"},
                                   "text": "User Administration  ·  ADMIN ONLY  ·  Internal user source"},
                         "type": "ia.display.label"},
                        {"meta": {"domId": "ua-admin-banner", "name": "AdminBanner"},
                         "propConfig": {
                             "props.text": {"binding": {"config": {
                                 "expression": "if({view.custom.isAdmin}, '', 'ADMIN ONLY — restricted. (Server-side enforcement is in the auth library; this banner is UI-only.)')"},
                                 "type": "expr"}},
                             "meta.visible": {"binding": {"config": {
                                 "expression": "!{view.custom.isAdmin}"}, "type": "expr"}},
                         },
                         "props": {"style": {"backgroundColor": "#FFF3CD", "border": "1px solid #FFC107",
                                             "color": "#7A5C00", "fontSize": "12px", "fontWeight": "bold",
                                             "padding": "6px 8px"}},
                         "type": "ia.display.label"},
                        # filter
                        {"meta": {"name": "FilterBar"}, "position": {"shrink": 0},
                         "props": {"alignItems": "flex-end",
                                   "style": {"backgroundColor": "#FFFFFF", "border": "1px solid #D0D4DA",
                                             "borderRadius": "4px", "gap": "12px", "padding": "10px"}},
                         "type": "ia.container.flex",
                         "children": [
                             {"meta": {"domId": "ua-search", "name": "SearchField"},
                              "position": {"basis": "240px"},
                              "propConfig": {"props.text": {"binding": {"config": {
                                  "bidirectional": True, "path": "view.custom.searchTerm"},
                                  "type": "property"}}},
                              "props": {"placeholder": "Search username"},
                              "type": "ia.input.text-field"},
                             {"events": {"component": {"onActionPerformed": {"config": {
                                 "script": "\tself.view.custom.runNonce = (self.view.custom.runNonce or 0) + 1\n"},
                                 "scope": "G", "type": "script"}}},
                              "meta": {"domId": "ua-search-btn", "name": "SearchButton"},
                              "props": {"style": {"backgroundColor": "#1976D2", "color": "#FFFFFF", "fontWeight": "bold"},
                                        "text": "Refresh"}, "type": "ia.input.button"},
                         ]},
                        # grid
                        {"events": {"component": {"onRowClick": {"config": {
                            "script": "\tlog = system.util.getLogger(\"SPIKE\")\n\ttry:\n\t\tname = event.value[\"Username\"]\n\texcept Exception:\n\t\tname = None\n\tif name is None:\n\t\treturn\n\tself.view.custom.selectedUser = unicode(name)\n\tself.view.custom.statusMsg = \"Selected %s.\" % name\n\tlog.info(\"SPIKE UserAdmin select: %s\" % name)\n"},
                            "scope": "G", "type": "script"}}},
                         "meta": {"domId": "ua-grid", "name": "UsersGrid"},
                         "position": {"basis": "480px", "grow": 1},
                         "propConfig": {
                             "props.columns": {"binding": {"config": {"path": "view.custom.userList.columns"}, "type": "property"}},
                             "props.data": {"binding": {"config": {"path": "view.custom.userList.data"}, "type": "property"}},
                         },
                         "props": {"filtering": {"enabled": False},
                                   "selection": {"enabled": True},
                                   "pager": {"activeOption": 1000, "bottom": False, "options": [1000], "top": False}},
                         "type": "ia.display.table"},
                        {"meta": {"name": "FooterNote"},
                         "props": {"style": {"color": "#37474F", "fontSize": "11px", "fontStyle": "italic"},
                                   "text": "Admin-only. Server-side enforcement lives in the auth Project Library (auth.addUser/removeUser/resetPassword re-check Admin from the SESSION). Passwords are hashed by the Internal user source — never stored plaintext. New users + admin resets force a first-login password reset."},
                         "type": "ia.display.label"},
                    ],
                },
                # RIGHT: actions (hidden from non-admins; defense-in-depth)
                {
                    "meta": {"name": "RightPane"},
                    "position": {"basis": "560px", "grow": 1},
                    "props": {"direction": "column", "style": {"gap": "8px"}},
                    "type": "ia.container.flex",
                    "children": [
                        {"meta": {"name": "DetailTitle"},
                         "props": {"style": {"color": "#222222", "fontSize": "15px", "fontWeight": "bold"}, "text": "Actions"},
                         "type": "ia.display.label"},
                        {"meta": {"domId": "ua-status", "name": "StatusMsg"},
                         "propConfig": {"props.text": {"binding": {"config": {"path": "view.custom.statusMsg"}, "type": "property"}}},
                         "props": {"style": {"color": "#B71C1C", "fontSize": "13px", "fontWeight": "bold", "padding": "6px 8px"}},
                         "type": "ia.display.label"},
                        {"meta": {"domId": "ua-form", "name": "Form"},
                         "propConfig": {"meta.visible": {"binding": {"config": {
                             "expression": "{view.custom.isAdmin}"}, "type": "expr"}}},
                         "props": {"direction": "column",
                                   "style": {"backgroundColor": "#FFFFFF", "border": "1px solid #D0D4DA",
                                             "borderRadius": "4px", "gap": "6px", "padding": "12px"}},
                         "type": "ia.container.flex",
                         "children": add_form_children + manage_children},
                    ],
                },
            ],
        },
    }
    return view


def _write_view(path, view):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w") as f:
        json.dump(view, f, indent=1, sort_keys=True)
        f.write("\n")
    print("wrote", path)


def _repo_resource_json():
    """The repo Admin/Users resource.json is the ONE deliberate exception to the "repo omits attributes"
    convention (repo-split-plan S1/R3): it is the ONLY in-repo seed for this view (no docs/design twin),
    so a cold gateway start that finds a null `attributes` map NPEs in ProjectResourceBuilder.setAttributes
    and FAULTS THE WHOLE GATEWAY. So the repo copy MUST carry a non-null `attributes` block.

    Prefer the live-gateway-signed map already on disk (a real re-signed signature exported per the plan);
    fall back to preserving a valid existing repo signature; only as a last resort emit the all-zeros
    placeholder (still non-null -> no NPE; the gateway re-signs on first load)."""
    def _valid(d):
        a = (d or {}).get("attributes") or {}
        sig = a.get("lastModificationSignature", "")
        return bool(sig) and sig != "0" * 64 and len(sig) == 64
    for src in (GW_RESOURCE, REPO_RESOURCE):   # gateway-signed first, then an already-valid repo copy
        try:
            with open(src) as f:
                d = json.load(f)
            if _valid(d):
                return dict(RESOURCE_JSON, attributes=d["attributes"])
        except (IOError, OSError, ValueError):
            pass
    return GW_RESOURCE_JSON   # all-zeros placeholder attributes: non-null (no NPE), gateway re-signs


def main():
    view = build_view()
    _write_view(GW_OUT, view)
    _write_view(REPO_OUT, view)
    # repo resource.json — MUST carry attributes (S1/R3: the only cold-start seed; null map faults gateway)
    with open(REPO_RESOURCE, "w") as f:
        json.dump(_repo_resource_json(), f, indent=2)
        f.write("\n")
    print("wrote", REPO_RESOURCE)
    # gateway resource.json (WITH attributes — required so a cold start does not NPE/fault the gateway)
    os.makedirs(os.path.dirname(GW_RESOURCE), exist_ok=True)
    with open(GW_RESOURCE, "w") as f:
        json.dump(GW_RESOURCE_JSON, f, indent=2)
        f.write("\n")
    print("wrote", GW_RESOURCE)


if __name__ == "__main__":
    main()
