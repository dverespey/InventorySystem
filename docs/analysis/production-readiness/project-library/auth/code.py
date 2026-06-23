# auth — Project Library service: SERVER-SIDE authorization gate + user-management ops.
#
# M4 piece 2 (minimal auth, single-site). This module is the HOLE-CLOSER for legacy hazard H3
# (m4-auth-sites-sourcetruth.md §1: the legacy authz was ONE client-side line —
# `MenuBar_MainMenu.Items.Remove(Administration_MenuItem)` at MainMenu.pas:457 — trivially
# bypassed by a modified client or direct DB access). Here the Admin/ProductionControl split is
# enforced INSIDE the gateway scope, from the SESSION's roles, never a client-passed flag.
#
# THE ROLE MODEL (David 2026-06-22):
#   * Admin            -> user add / delete / reset ONLY (the one thing ProductionControl cannot do).
#   * ProductionControl-> EVERYTHING ELSE in the app (masters, ASN, order, EDI, reports, Sites config).
#   ~2 users per site; single-site (each deployment = its own gateway = the site; no per-user->site map).
#
# WHY A PROJECT-LIBRARY MODULE (not inline view scripts):
#   1. SERVER-SIDE: the user-mgmt write ops (add/delete/reset) call system.user.* AND re-check the
#      caller's role here, in the gateway, so hiding the UI is NOT the security boundary. A request
#      that reaches addUser/removeUser/setUserPassword is rejected unless the SESSION carries Admin.
#   2. HEADLESS-DRIVABLE: the gate decision (`requireRole`/`authorize`) is pure given the caller's roles,
#      so scripts/e2e/test_m4_auth.py drives the REAL gateway-side decision through the jython_shim
#      (which now provides system.user/system.perspective stand-ins) and proves a non-Admin is denied
#      IN THIS SCRIPT — not in the browser chrome (retro R8: a test that EXECs the proc/UI doesn't
#      exercise the gateway-side driver).
#
# Jython 2.7 (8.1.52 and 8.3 — no scripting delta). IG81-COMPAT: system.user.* / system.perspective.*
# are 8.1+ and unchanged on 8.3. The role NAMES below are the gateway user-source roles seeded per the
# Designer/gateway-config steps documented in docs/analysis/production-readiness/m4-auth-design.md.

USER_SOURCE = "Inventory"   # the gateway Internal user source name (gateway-config; see m4-auth-design.md)

# Canonical role names (must match the seeded user-source roles 1:1; m4-auth-design.md §"Roles").
ROLE_ADMIN = "Admin"
ROLE_PRODUCTION = "ProductionControl"

# The single security-level path the AUTHORITATIVE Perspective gate also checks
# (isAuthorized(false, "Authenticated/Roles/Admin")). Kept here so the view binding + this module
# name the same level (one source of truth).
SECLEVEL_ADMIN = "Authenticated/Roles/Admin"


# ===========================================================================
# PURE decision core (no `system`, no I/O) — CPython-importable; the test drives THIS for the gate.
# ===========================================================================

def _norm_roles(roles):
	"""Normalize a session's roles to a set of unicode strings. Accepts None / a single string /
	any iterable (Perspective hands roles as a list of strings)."""
	if roles is None:
		return set()
	if isinstance(roles, basestring):  # noqa: F821 (Jython 2.7 builtin)
		return set([roles])
	return set(unicode(r) for r in roles)


def hasRole(roles, role):
	"""True iff `role` is present in the caller's roles. Pure."""
	return role in _norm_roles(roles)


def isAdmin(roles):
	"""True iff the caller holds the Admin role. Pure — the user-mgmt gate predicate."""
	return hasRole(roles, ROLE_ADMIN)


def isProductionControl(roles):
	"""True iff the caller holds ProductionControl (everything-else in the app). Pure."""
	return hasRole(roles, ROLE_PRODUCTION)


class AuthError(Exception):
	"""Raised when a server-side authorization check fails. The user-mgmt ops raise this BEFORE any
	system.user.* write, so a denied request never mutates the user source."""
	pass


def authorize(roles, requiredRole):
	"""Pure gate: returns True if `roles` includes `requiredRole`, else raises AuthError. The single
	chokepoint every Admin-only op calls FIRST (server-side), so the deny path is one tested predicate.
	`roles` MUST come from the session/gateway (see sessionRoles), NEVER from a client parameter."""
	if not hasRole(roles, requiredRole):
		raise AuthError("Forbidden: requires role '%s' (caller roles: %s)"
		                % (requiredRole, sorted(_norm_roles(roles))))
	return True


def authorizeAny(roles, requiredRoles):
	"""Pure gate: returns True if `roles` includes ANY of `requiredRoles`, else raises AuthError. Used by
	the master-CRUD write gate (Sites etc.): a write is allowed for ProductionControl (everything-else in
	the app) OR Admin. `roles` MUST come from the session/gateway, NEVER a client parameter."""
	have = _norm_roles(roles)
	want = _norm_roles(requiredRoles)
	if not (have & want):
		raise AuthError("Forbidden: requires one of %s (caller roles: %s)"
		                % (sorted(want), sorted(have)))
	return True


# The roles allowed to perform an app data WRITE (master CRUD, ASN, order, EDI, reports). ProductionControl
# is the everything-else role; Admin is included so an administrator is never locked out of config edits.
WRITE_ROLES = (ROLE_PRODUCTION, ROLE_ADMIN)


def requireWrite(session):
	"""SERVER-SIDE WRITE GATE — the reusable hole-closer every master-CRUD view (Sites, and the swept
	masters) calls BEFORE any system.db write. It resolves the caller's roles FROM THE SESSION (gateway-
	side, via sessionRoles — NOT a client prop) and authorizes ProductionControl|Admin, raising AuthError
	on deny. So a forged client prop (`view.custom.mayEdit=true` via devtools) or an anonymous session
	CANNOT write: the decision is grounded in the server's view of who is logged in, not anything the page
	passed us. This is the fix for the reintroduced legacy H3 hole (master writes were authorized client-
	side only). Returns the resolved roles (set) on success so the caller can log/branch if it wants.

	Usage in a view button script (Save/Delete/New), BEFORE the runPrepUpdate/runPrepQuery:
	    import auth as A
	    A.requireWrite(self.session)   # raises A.AuthError if the SESSION lacks ProductionControl/Admin
	"""
	roles = sessionRoles(session)
	authorizeAny(roles, WRITE_ROLES)
	return roles


def mayWrite(session):
	"""Non-raising companion to requireWrite: True iff the SESSION holds a write role. For UI defense-in-
	depth (hide the form/buttons) ONLY — the server-side gate is requireWrite, which every write MUST call.
	(A view binding may still use isAuthorized(...) for the visual indicator; this is the script-side echo
	of the same SESSION-grounded decision.)"""
	try:
		return bool(sessionRoles(session) & _norm_roles(WRITE_ROLES))
	except (AttributeError, KeyError, TypeError):
		return False


# ===========================================================================
# SESSION role resolution (server-side) — the part the client CANNOT spoof.
# ===========================================================================

def sessionRoles(session):
	"""Resolve the caller's roles from the Perspective SESSION object, server-side.

	In a Perspective component/message-handler script the authenticated identity lives at
	session.props.auth.user.roles — populated by the gateway from the IdP/user-source at login and
	NOT writable by the client. We read it here so the gate decision is grounded in the server's view
	of who is logged in, not a value the page passed us. (m4-auth-sites-sourcetruth.md §2 gating rule.)

	Accepts the live PropertyTreeScriptWrapper (session.props.auth.user.roles) OR a plain dict shape
	(what the headless harness/shim hands us) so the SAME code path is exercised both live and in test."""
	if session is None:
		return set()
	# live Perspective: session.props.auth.user.roles
	try:
		return _norm_roles(session.props.auth.user.roles)
	except (AttributeError, KeyError, TypeError):
		pass
	# dict shape (harness / message-payload echo of the resolved roles): {"auth": {"user": {"roles": [...]}}}
	try:
		return _norm_roles(session["auth"]["user"]["roles"])
	except (KeyError, TypeError):
		return set()


# ===========================================================================
# USER-MANAGEMENT OPS — Admin-gated server-side, then call system.user.*.
# Each op takes `callerRoles` resolved from the SESSION (by the caller, via sessionRoles) and re-checks
# Admin HERE before touching the user source. This is the enforcement boundary, not the UI.
# ===========================================================================

def _uiresponse_text(resp):
	"""Flatten a UIResponse (errors/warns/infos) to a readable string for status + logging."""
	if resp is None:
		return "ok"
	try:
		errs = list(resp.getErrors() or [])
		warns = list(resp.getWarns() or [])
		if errs:
			return "errors: " + "; ".join(unicode(e) for e in errs)
		if warns:
			return "warnings: " + "; ".join(unicode(w) for w in warns)
		return "ok"
	except AttributeError:
		return unicode(resp)


def addUser(callerRoles, username, roles, password, forceReset=True):
	"""ADMIN-ONLY: create a user in the Internal user source. No plaintext is stored — the Internal
	user source hashes the password (HD2 / H1 retired). `forceReset=True` flags the new user must-reset
	on first login (Q13; the faithful successor to the legacy NULL-LastUpdated precedent, §1).

	roles : iterable of role names to grant (validated against the known roles).
	Returns the system.user.addUser UIResponse text. Raises AuthError if the caller is not Admin
	(BEFORE any system.user call — a denied request never creates a user)."""
	authorize(callerRoles, ROLE_ADMIN)
	_validate_username(username)
	grant = _validate_roles(roles)
	log = system.util.getLogger("SPIKE")
	user = system.user.getNewUser(USER_SOURCE, username)
	for r in grant:
		user.addRole(r)
	user.set("password", password)
	# First-login reset flag (Q13): the Internal user-source exposes a schedule/force-change knob;
	# we ALSO record it in a contact-info marker so the first-login flow (firstLoginResetRequired)
	# can detect it regardless of which 8.x exposes the native flag. See markMustReset.
	resp = system.user.addUser(USER_SOURCE, user)
	if forceReset:
		markMustReset(callerRoles, username, True)
	log.info("SPIKE auth addUser: caller=Admin user=%s roles=%s forceReset=%s -> %s"
	         % (username, sorted(grant), forceReset, _uiresponse_text(resp)))
	return _uiresponse_text(resp)


def removeUser(callerRoles, username):
	"""ADMIN-ONLY: delete a user from the Internal user source. Raises AuthError if not Admin (before
	the system.user.removeUser call)."""
	authorize(callerRoles, ROLE_ADMIN)
	_validate_username(username)
	log = system.util.getLogger("SPIKE")
	resp = system.user.removeUser(USER_SOURCE, username)
	log.info("SPIKE auth removeUser: caller=Admin user=%s -> %s" % (username, _uiresponse_text(resp)))
	return _uiresponse_text(resp)


def resetPassword(callerRoles, username, newPassword, clearMustReset=False):
	"""ADMIN-ONLY: set a user's password (the admin-driven reset). No plaintext stored (hashed by the
	source). By default this is an ADMIN reset that re-arms must-reset (the user picks their own on next
	login); the SELF-service first-login flow (setOwnPasswordFirstLogin) clears must-reset instead.
	Raises AuthError if not Admin."""
	authorize(callerRoles, ROLE_ADMIN)
	_validate_username(username)
	log = system.util.getLogger("SPIKE")
	user = system.user.getUser(USER_SOURCE, username)
	if user is None:
		raise AuthError("No such user '%s'" % username)
	user.set("password", newPassword)
	resp = system.user.editUser(USER_SOURCE, user)
	# admin reset -> the user must set their own on next login (unless explicitly clearing)
	markMustReset(callerRoles, username, not clearMustReset)
	log.info("SPIKE auth resetPassword: caller=Admin user=%s mustReset=%s -> %s"
	         % (username, not clearMustReset, _uiresponse_text(resp)))
	return _uiresponse_text(resp)


# ===========================================================================
# FIRST-LOGIN PASSWORD RESET (Q13) — flag + self-service clear. No plaintext.
# ===========================================================================

# The must-reset marker lives in the user's contact-info as a sentinel key. We use contact-info because
# it is the portable, 8.1+-stable PyUser field; the native "schedule adjustment / force password change"
# knob varies across 8.x point releases. The first-login flow reads firstLoginResetRequired(user) which
# treats EITHER the sentinel OR a never-set-password user as must-reset.
MUST_RESET_CONTACT_TYPE = "sms"          # a benign contact slot reused as the sentinel carrier
MUST_RESET_SENTINEL = "MUST_RESET=1"     # presence in any contact value => must reset on next login


def markMustReset(callerRoles, username, mustReset):
	"""ADMIN-ONLY: set/clear the must-reset sentinel on a user. Used by addUser/resetPassword to FORCE a
	first-login reset (the faithful successor to the legacy NULL-LastUpdated, §1)."""
	authorize(callerRoles, ROLE_ADMIN)
	user = system.user.getUser(USER_SOURCE, username)
	if user is None:
		raise AuthError("No such user '%s'" % username)
	_apply_sentinel(user, mustReset)
	system.user.editUser(USER_SOURCE, user)
	return mustReset


def firstLoginResetRequired(user):
	"""PURE-ish read: does this user need a first-login reset? True iff the must-reset sentinel is
	present OR the user has never set a password (last-password-change is null). Called by the login /
	session-startup gate (see m4-auth-design.md "First-login flow"). Takes a PyUser (live) or a dict
	{"contactInfo": [(type, value), ...], "passwordSet": bool} (harness) so the SAME logic is tested."""
	for _t, val in _contact_pairs(user):
		if val and MUST_RESET_SENTINEL in val:
			return True
	# never-set-password fallback (mirrors legacy NULL-LastUpdated => forced reset)
	return not _password_ever_set(user)


def setOwnPasswordFirstLogin(session, username, newPassword):
	"""SELF-SERVICE first-login reset: the logged-in user sets their OWN new password and the must-reset
	sentinel is CLEARED. Authorization here is IDENTITY, not role: the session's own username must equal
	`username` (a user may only reset THEIR OWN password via this path — Admin uses resetPassword). No
	plaintext stored (hashed by the source). Raises AuthError on a username mismatch."""
	who = sessionUsername(session)
	if who is None or unicode(who) != unicode(username):
		raise AuthError("First-login reset is self-service only: session user '%s' != target '%s'"
		                % (who, username))
	log = system.util.getLogger("SPIKE")
	user = system.user.getUser(USER_SOURCE, username)
	if user is None:
		raise AuthError("No such user '%s'" % username)
	user.set("password", newPassword)
	_apply_sentinel(user, False)   # clear the sentinel in the SAME edit
	resp = system.user.editUser(USER_SOURCE, user)
	log.info("SPIKE auth firstLoginReset: user=%s self-set password, must-reset cleared -> %s"
	         % (username, _uiresponse_text(resp)))
	return _uiresponse_text(resp)


# ---- optional age-based expiry (legacy IN_PASSWORD_RESET_DAYS; D-M4-4 small add) ----

def passwordExpired(user, maxAgeDays):
	"""Honor the legacy age-based expiry (IN_PASSWORD_RESET_DAYS). True iff the user's password is older
	than maxAgeDays. maxAgeDays<=0 disables expiry (the common case). PURE given the user's
	last-password-change epoch-day. The Internal source tracks last-change; the live caller passes a
	PyUser, the harness a dict {"passwordChangedDaysAgo": N}."""
	if not maxAgeDays or maxAgeDays <= 0:
		return False
	age = _password_age_days(user)
	if age is None:
		return True   # never set / unknown -> treat as expired (mirrors legacy NULL-LastUpdated)
	return age >= maxAgeDays


# ===========================================================================
# SESSION username (server-side) — for the self-service first-login identity check.
# ===========================================================================

def sessionUsername(session):
	"""The authenticated username from the SESSION (session.props.auth.user.userName), server-side.
	Accepts the live wrapper or the dict shape (harness)."""
	if session is None:
		return None
	try:
		return session.props.auth.user.userName
	except (AttributeError, KeyError, TypeError):
		pass
	try:
		return session["auth"]["user"]["userName"]
	except (KeyError, TypeError):
		return None


# ===========================================================================
# small validators + PyUser/dict accessors (live + harness portable)
# ===========================================================================

def _validate_username(username):
	if not username or not unicode(username).strip():
		raise ValueError("username is required")
	return unicode(username).strip()


def _validate_roles(roles):
	"""Only the two known roles may be granted; reject anything else (defense against a typo'd/escalated
	role string sneaking in)."""
	known = set([ROLE_ADMIN, ROLE_PRODUCTION])
	grant = _norm_roles(roles)
	bad = grant - known
	if bad:
		raise ValueError("Unknown role(s): %s (allowed: %s)" % (sorted(bad), sorted(known)))
	if not grant:
		raise ValueError("at least one role is required")
	return grant


def _contact_pairs(user):
	"""[(contactType, value), ...] from a live PyUser OR a harness dict. 8.1 PyUser.getContactInfo()
	returns ContactInfo objects exposing .contactType / .value (properties). We also tolerate the
	getter shape (getContactType/getValue) for portability. The harness passes a dict
	{"contactInfo": [(type, value), ...]} so firstLoginResetRequired is tested as written."""
	if isinstance(user, dict):
		return list(user.get("contactInfo", []))
	out = []
	try:
		for ci in (user.getContactInfo() or []):
			t = getattr(ci, "contactType", None)
			v = getattr(ci, "value", None)
			if t is None and hasattr(ci, "getContactType"):
				t, v = ci.getContactType(), ci.getValue()
			out.append((t, v))
	except AttributeError:
		pass
	return out


def _apply_sentinel(user, present):
	"""Set (present=True) or clear (present=False) the must-reset sentinel on a user using the DOCUMENTED
	8.1 PyUser contact API (addContactInfo(type, value) / removeContactInfo(type, value)) — NOT a
	setContactInfo replacement (which is not the 8.1 method). Idempotent: clears any existing sentinel
	first, then adds it back only if present. Harness dicts are mutated in the equivalent way."""
	# clear any existing sentinel under our marker type (whatever its value)
	for (t, v) in _contact_pairs(user):
		if v and MUST_RESET_SENTINEL in v:
			_remove_contact(user, t, v)
	if present:
		_add_contact(user, MUST_RESET_CONTACT_TYPE, MUST_RESET_SENTINEL)


def _add_contact(user, ctype, value):
	if isinstance(user, dict):
		user.setdefault("contactInfo", []).append((ctype, value))
		return
	user.addContactInfo(ctype, value)   # 8.1: addContactInfo(contactType, value)


def _remove_contact(user, ctype, value):
	if isinstance(user, dict):
		user["contactInfo"] = [ci for ci in user.get("contactInfo", []) if ci != (ctype, value)]
		return
	user.removeContactInfo(ctype, value)  # 8.1: removeContactInfo(type, value) — exact-match delete


def _password_ever_set(user):
	if isinstance(user, dict):
		return bool(user.get("passwordSet", False))
	try:
		# PyUser exposes the password-set / last-change via getOrDefault; null => never set
		return user.get(user.PasswordLastChanged) is not None
	except (AttributeError, TypeError):
		return True   # can't tell -> assume set (don't force a needless reset on a live source)


def _password_age_days(user):
	if isinstance(user, dict):
		return user.get("passwordChangedDaysAgo")
	return None  # live: the source tracks it; expiry is an optional add (D-M4-4) — left to the source policy
