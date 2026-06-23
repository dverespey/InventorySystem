# M4 piece 2 — Minimal Auth (single-site): role model + gateway-config + the headless-vs-Designer split

This is the design + runbook the auth code points to (`project-library/auth/code.py`,
`scripts/gen_user_admin_view.py`). It records the AS-BUILT role model, what is enforced SERVER-SIDE in
code vs what is gateway/Designer config, the exact gateway-config steps, and the first-login flow.

Source-of-truth design input: `m4-auth-sites-sourcetruth.md` (the legacy auth model + the Ignition
target). Model confirmed by David 2026-06-22.

---

## 1. The role model (as built)

Two roles. ~2 users per site. **Single-site**: each deployment is its OWN gateway + DB, so the gateway
IS the site — there is NO per-user→site mapping.

| Role | What it can do |
|---|---|
| **Admin** | User **add / delete / reset** — ONLY. (The one thing ProductionControl cannot do.) |
| **ProductionControl** | **EVERYTHING ELSE** in the app: all masters, ASN, Order, EDI, reports, **and the Sites config**. |

The faithful legacy precedent: the legacy split was Admin-menu vs everything-else, with the Admin menu
gating User Administration (+ Configuration). M4 narrows "Admin" to exactly user management and folds the
rest (including Sites/site config) into ProductionControl, per David's confirmed model.

### Role → view/page gating matrix

| View / page | Role required | Enforcement |
|---|---|---|
| **User Administration** (`/admin/users`, `Admin/Users/Users`) | **Admin** | Page role-permission (Designer) **+ server-side WRITE gate in `auth.*`** (every add/delete/reset) **+ server-side READ gate** (the list builder re-checks Admin from the session before `system.user.getUsers` → empty list for a non-Admin) |
| **Sites config** (`/sites`, `Master/Sites/Sites`) | **ProductionControl** | Page role-permission (Designer) **+ server-side WRITE gate** (`auth.requireWrite(self.session)` on Save/Delete/New, resolves SESSION roles gateway-side, raises on deny — BEFORE any `system.db` write) **+ PROTECTED `mayEdit` UI prop** (browser cannot forge it; UI defense-in-depth only) |
| All **other masters** (Size, Supplier, PartsStock, ManifestCost, RenbanGroup, AssemblyDetail, Logistics) | **ProductionControl** | Page role-permission (Designer) **+ server-side WRITE gate — TO BE APPLIED (P15):** these 7 views currently write with NO server-side gate (the same H3 hole). The reusable `auth.requireWrite(session)` is built; each is one call away. Tracked as punch-list **P15**. |
| Order / ASN / EDI / reports | **ProductionControl** | Page role-permission (Designer) |

> **AS-BUILT before the M4-auth-write-gate fix (2026-06-22): the Sites/master WRITES had NO server-side
> gate** — they were authorized CLIENT-SIDE only (an advisory `isAdmin`/`isAuthorized()` binding, which IA
> docs say is a visual indicator only, on a forgeable Public prop). That reintroduced the legacy H3 hole.
> The fix routes the **Sites** Save/Delete/New through the SERVER-SIDE `auth.requireWrite(self.session)`
> gate (forged-prop-rejected, revert-proven in `test_m4_auth.py`) and makes the `mayEdit` UI prop
> **PROTECTED**. The Sites UI-visibility prop was renamed `isAdmin → mayEdit` (it now means "may edit
> Sites" = ProductionControl). The **other 7 masters still need the same server-side write gate (P15).**

---

## 2. What is ENFORCED SERVER-SIDE (code) vs GATEWAY-CONFIG (Designer/gateway)

This closes legacy hazard **H3** (the legacy authz was ONE client-side line — a menu removal — trivially
bypassed) for the **user-management ops AND the Sites config write**. Those writes are NOT protected by
hiding the UI; they are protected in the gateway scope.

> **CORRECTION (2026-06-22).** An earlier revision of this doc asserted "closes H3" for the Sites/master
> writes, but **as-built those writes had NO server-side gate** — they were authorized client-side only
> (a forgeable Public `isAdmin` prop + an advisory `isAuthorized()` binding, which IA docs say is a visual
> indicator only). That was the H3 hole reintroduced. It is **now actually closed for Sites** (the
> server-side write gate below); the **7 other masters remain to be gated — punch-list P15.**

### Server-side-enforced (in `project-library/auth/code.py`, proven by `test_m4_auth.py`)
- **Admin gate on every user-mgmt op.** `addUser` / `removeUser` / `resetPassword` each call
  `authorize(callerRoles, ROLE_ADMIN)` FIRST and raise `AuthError` if the caller is not Admin — BEFORE
  any `system.user.*` call. A denied request never mutates the user source.
- **Reusable WRITE gate for app data (master CRUD).** `auth.requireWrite(session)` resolves the caller's
  roles FROM THE SESSION (gateway-side) and authorizes `ProductionControl|Admin` (`authorizeAny`),
  raising `AuthError` on deny. The **Sites** Save/Delete/New call it FIRST, BEFORE any `system.db` write,
  so a forged client prop (`mayEdit=true` via devtools) or an anonymous session is rejected in the
  gateway — proven forged-prop-rejected + revert-proven in `test_m4_auth.py`. The `mayEdit` UI prop is
  **PROTECTED** (the back-end ignores browser writes to it), so the UI gate itself can't be forged either.
  (The 7 other masters get this same one-line gate under P15.)
- **Server-side READ gate on the user list (defense-in-depth symmetry).** The User Admin list builder
  re-checks Admin from the session BEFORE `system.user.getUsers`, returning an EMPTY list to a non-Admin —
  so even if the page-permission is ever forgotten there is no account/role info-leak.
- **Roles come from the SESSION, not the client.** The view scripts resolve `callerRoles` via
  `auth.sessionRoles(self.session)` / `auth.requireWrite(self.session)` (reads
  `session.props.auth.user.roles`, populated by the gateway at login, not writable by the page). No
  client-passed admin flag is ever trusted.
- **No self-escalation.** `_validate_roles` rejects unknown roles, and the gate keys off the CALLER's
  role, so a ProductionControl caller cannot create an Admin.
- **First-login self-service is identity-gated.** `setOwnPasswordFirstLogin` requires the session's own
  username to equal the target (a user may only reset THEIR OWN password via that path).

### Gateway-config (Designer / gateway web UI — NOT headless; do these to finish)
1. **Internal user source** named **`Inventory`** (must match `auth.USER_SOURCE`).
   Gateway → Config → Security → **User Sources** → Create new → **Internal**. Set the password policy
   (min length/complexity — the legacy had none; pick a baseline). Optionally set "Password Expiration"
   to honor the legacy `IN_PASSWORD_RESET_DAYS` (D-M4-4; the code also supports it via
   `auth.passwordExpired`).
2. **Roles** in that source: **`Admin`** and **`ProductionControl`** (exact names — match
   `auth.ROLE_ADMIN` / `auth.ROLE_PRODUCTION`).
3. **Security level** `Authenticated/Roles/Admin` and `Authenticated/Roles/ProductionControl`
   (Gateway → Config → Security → **Security Levels**) so `isAuthorized(...)` resolves. With an Internal
   user source + the default IdP, the `Authenticated/Roles/<roleName>` levels are auto-derived from the
   user's roles; confirm they appear.
4. **Project default IdP** = the IdP backed by the `Inventory` user source (Project Properties →
   Perspective → General, or the gateway IdP config).
5. **Perspective PAGE permissions** (the AUTHORITATIVE UI gate). In the Designer, on the **Page
   Configuration** for:
   - `/admin/users` → require security level **`Authenticated/Roles/Admin`**.
   - `/sites` and every other app page → require **`Authenticated/Roles/ProductionControl`**.
   This makes the page unreachable to a user without the role (the in-view `isAdmin` binding +
   the `qaAdmin` URL hatch are then defense-in-depth/spike-only and unreachable in prod).
6. **Seed users** (NO plaintext migrated — legacy passwords are plaintext, H1):
   create ~2 users; grant `Admin` to the administrator account, `ProductionControl` to operators; set a
   temporary password and leave them **must-reset** (see §4). The legacy `INV_USERS` IDs may be re-used
   as usernames; the `BIT_ADMIN=1` users become `Admin`, the rest `ProductionControl`.

> The spike box has anonymous Perspective sessions + no IdP, so the live `isAuthorized(...)` fails CLOSED
> (false) — the correct prod behavior. The e2e harness reaches the screens via the `?qaAdmin=1` URL
> hatch; **IG83-TODO: drop the `qaAdmin` branch once page permissions are configured in the Designer.**

---

## 3. The User Admin screen (Admin-only)

`Admin/Users/Users` (generated by `scripts/gen_user_admin_view.py`). The ONE screen Admin can use that
ProductionControl cannot. List (left) of users from `system.user.getUsers`; actions (right): **Add**,
**Reset Password**, **Delete**. Every action button calls the `auth` module (server-side-gated) — the
view never calls `system.user.*` inline, so the UI cannot bypass the gate. Passwords are hashed by the
Internal source (never stored plaintext). New users + admin resets force a first-login reset.

8.1 API used (verified vs IA docs): `system.user.getNewUser(source, name)`, `user.addRole(r)`,
`user.set("password", pw)`, `system.user.addUser(source, user)` (returns UIResponse),
`system.user.getUsers(source)`, `user.getRoles()`, `system.user.getUser`, `editUser`, `removeUser`.

---

## 4. First-login password reset flow (Q13)

Faithful successor to the legacy NULL-`LastUpdated` precedent (`m4-auth-sites-sourcetruth.md` §1: a
net-new user has NULL LastUpdated → forced reset on first login). No plaintext stored.

1. **Admin creates** a user with a temporary password (`auth.addUser(..., forceReset=True)`), which sets
   a **must-reset sentinel** on the user (a contact-info marker, portable across 8.x point releases).
2. On **first login**, the session-startup / login gate calls `auth.firstLoginResetRequired(user)` →
   True (sentinel present OR password never set) → the user is routed to a **set-new-password** view
   before the app is usable.
3. The user sets their own password via `auth.setOwnPasswordFirstLogin(session, username, newPw)` —
   identity-gated (own account only), hashed by the source, and the sentinel is CLEARED in the same edit.
4. An **Admin reset** (`auth.resetPassword`) RE-ARMS the sentinel (the user picks a new one next login).
5. **Optional age-expiry** (legacy `IN_PASSWORD_RESET_DAYS`, D-M4-4): `auth.passwordExpired(user, days)`
   honors it; set the policy on the Internal source (§2.1) or call this in the login gate. `days<=0`
   disables it (common case).

> DESIGNER-FINISH for the login routing: a session-startup script (Project Properties → Perspective →
> Session Events → **Startup**) reads the logged-in user, calls `auth.firstLoginResetRequired`, and
> navigates to the set-password view if True. The flag LOGIC + the self-service reset op are built +
> tested; the startup-script wiring + the set-password view chrome are the Designer finish.

---

## 5. Headless-vs-Designer split (honest)

| Built + proven headless (code) | Gateway/Designer config (flagged) |
|---|---|
| `auth` module: the server-side Admin gate, the reusable **`requireWrite(session)` write gate** + `authorizeAny`, addUser/removeUser/resetPassword, first-login flag + self-service reset, age-expiry | The Internal user source + the 2 roles + seeded users |
| User Admin view JSON + button scripts routing through the gate; **server-side READ gate** on the list builder | The Perspective PAGE role-permissions (authoritative UI gate) |
| Sites view: **server-side write gate** (`requireWrite`) on Save/Delete/New + **PROTECTED `mayEdit`** UI prop (renamed from `isAdmin`), gated to ProductionControl | The `Authenticated/Roles/*` security levels + the project IdP |
| `test_m4_auth.py`: user-mgmt gate rejection (revert-proven), **Sites write gate (forged-prop-rejected, revert-proven)**, **User Admin read gate (revert-proven)**, first-login flow, role→view consistency | The session-startup script that routes a must-reset user to the set-password view |
| (P15) the 7 other masters still need the `requireWrite` write gate added | — |

---

## 6. Residual design questions (for David / architects)

- **D-M4-4 (password policy beyond first-login).** Honor the legacy age-based expiry
  (`IN_PASSWORD_RESET_DAYS`)? The code supports it (`passwordExpired`); set the Internal-source policy or
  wire it in the login gate. Also pick a complexity baseline (the legacy had none).
- ~~**Cosmetic:** the Sites view custom prop is still named `isAdmin`~~ — **DONE (2026-06-22):** renamed
  to `mayEdit` (all bindings + banner text updated) and made **PROTECTED** so it cannot be browser-forged.
- **P15 (the systemic write-gate sweep):** apply the reusable `auth.requireWrite(self.session)` gate to
  the New/Save/Delete button scripts of the 7 other master views (Size, Supplier, PartsStock,
  ManifestCost, RenbanGroup, AssemblyDetail, Logistics), which today write with no server-side gate. See
  `cutover-punch-list.md` P15.
