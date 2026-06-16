# Module Analysis: Authentication & User Administration

**Area:** Admin / system / shell  **Status:** ✅ spec complete  **Analyst:** Claude / 2026-06-16

Covers `Logon.pas`, `UserAdmin.pas`, `ConfirmPassword.pas`, `NewPassword.pas`,
`UserInfo.pas`, and the `INV_USERS` proc group. **Security-critical** for the Ignition
rebuild (gateway auth + roles).

---

## 1. Legacy surface
- **Forms (all live in `InventorySystem.dpr`):**
  - `Logon.pas` (+`.dfm`) — login dialog (user id + password).
  - `UserAdmin.pas` (+`.dfm`) — user CRUD (insert/update/delete, admin flag).
  - `ConfirmPassword.pas` (+`.dfm`) — re-prompt current password (dead entry-path, see §5).
  - `NewPassword.pas` (+`.dfm`) — forced password-reset-on-expiry dialog.
  - `UserInfo.pas` — `TUserInfo` component: Win username / machine name / IP (audit context).
- **Entry points (`MainMenu.pas`):**
  - Login: `FormShow` calls `Logon_Form.Execute` at `MainMenu.pas:449-455`; failure
    `Application.Terminate`.
  - Admin menu gating: `MainMenu.pas:456-457` — `If Not gobjUser.AppUserAdmin Then
    MenuBar_MainMenu.Items.Remove(Administration_MenuItem)`.
  - User-admin screen: `Administration_UserAdmin_MenuItemClick` (`MainMenu.pas:548-561`).
- **Purpose:** authenticate the operator against `INV_USERS`, establish the in-memory
  `gobjUser` identity (id / password / admin bit + Windows context), and let an admin manage
  the user list. The single `BIT_ADMIN` flag is the **entire** authorization model.

## 2. Data touched
| Table | Read | Write | Notes |
|-------|:----:|:-----:|-------|
| `INV_USERS` | ✅ | ✅ | id, **plaintext** password, admin bit, `LastUpdated` stamp. |
| `Activity..` (Act DB) | | ✅ | every login / admin action logged via `LogActLog` (cross-DB; see configuration-site.md §DATAPURGE/Activity coupling). |

**`INV_USERS` schema** (`Create Inventory.sql:1773-1778`):
```
VC_USER_ID  varchar(30) NOT NULL
VC_PASSWORD varchar(30) NOT NULL   -- plaintext, max 30 chars
BIT_ADMIN   bit         NULL
LastUpdated varchar(16) NULL       -- yyyymmddHHMMSS+ff stamp, set only by UPDATE_UserInfo
```
**No PRIMARY KEY, no unique index** is declared on the table. Identity is the *composite*
`(VC_USER_ID, VC_PASSWORD)` — every proc keys on **both** id AND password (see §3). There is
no surrogate id (contrast D2 elsewhere).

**Triggers:** none on `INV_USERS`.

## 3. Stored procedures used
| Proc | Op | Business rule (from body) |
|------|----|---------------------------|
| `SELECT_UserInfo` (`:8117`) | SELECT | `@UserID=''` → list all `(VC_USER_ID, VC_PASSWORD, BIT_ADMIN)`; else **`WHERE VC_USER_ID=@UserID AND VC_PASSWORD=@Pass`** — plaintext equality match. Used for both login (validate) and the admin dup-check. |
| `INSERT_UserInfo` (`:3956`) | INSERT | Blind `INSERT (VC_USER_ID, VC_PASSWORD, BIT_ADMIN)`. No dup guard in the proc — dedup is client-side (see below). |
| `UPDATE_UserInfo` (`:9522`) | UPDATE | Keyed on **old** id+password: `WHERE VC_USER_ID=@OldUserID AND VC_PASSWORD=@OldPass`; sets new id/pass/admin + `LastUpdated` = 16-char `yyyymmddHHMMSSff` timestamp (`CONVERT(varchar,112)` [8] + 4×2-char `SUBSTRING(...,114)` slices = **16 chars**, correct count). |
| `DELETE_UserInfo` (`:2579`) | DELETE | `DELETE WHERE VC_USER_ID=@UserID AND VC_PASSWORD=@Pass`. |
| `UPDATE_UserPassword` | UPDATE | **CALLED BUT ABSENT FROM SCHEMA.** `DataModule.pas:6306` calls `dbo.UPDATE_UserPassword;1`, but no `CREATE PROCEDURE UPDATE_UserPassword` exists in `Create Inventory.sql`. **Latent runtime bug** in the password-expiry reset path (see §4, [[reference-schema-snapshot-vs-live]]). Body unverified — the proc does not exist in the snapshot. |

**DataModule wiring:**
- `ValidateUser` (`DataModule.pas:6062`) → `SELECT_UserInfo(@UserID,@Pass)`; if `RecordCount>0`
  it sets `gobjUser.AppUserID/AppUserPass/AppUserAdmin` from the returned row
  (`:6118-6120` / `:6137-6139`) and logs `LOGIN`; else logs `LOGIN ERR`.
- `InsertUser` (`:6218`) → calls `SELECT_UserInfo` first as a **dup-check**; only `If RecordCount=0`
  does it `INSERT_UserInfo` (`:6250`). Dedup is therefore on the *exact id+password pair* —
  same id with a *different* password is NOT a duplicate (see §4).
- `UpdateUserInfo` (`:6338`) → `UPDATE_UserInfo`; `@Admin` passed as `ABS(StrToInt(BoolToStr(...)))`.
- `DeleteUserInfo` (`:6391`) → `DELETE_UserInfo`.
- `UpdateUserPassword` (`:6298`) → `UPDATE_UserPassword` (missing proc).
- `SetComboBoxesWithUserObj` (`:6168`) → raw SQL
  `SELECT VC_USER_ID, VC_PASSWORD, BIT_ADMIN FROM INV_USERS ...`; loads **plaintext passwords
  into the combobox item objects** (`TUserAdminDetail`, `DataModule.pas:69/696`). The UserAdmin
  form reads them back into `Password_Edit` on selection (`UserAdmin.pas:99`) — note the field is
  **masked** (`UserAdmin.dfm:143 PasswordChar='*'`), so it's not readable off-screen, but the
  plaintext is in client process memory + on the wire.

## 4. Business rules & edge cases

### Auth / password model (faithful)
- **Passwords are stored and compared in PLAINTEXT.** `VC_PASSWORD varchar(30)`; login is a
  literal `WHERE VC_PASSWORD = @Pass` (`SELECT_UserInfo:8117`). No hashing, no salt, no
  encryption anywhere. 🔴 **Security finding — top migration concern.**
- **Plaintext is stored, transmitted, and held in client memory.** UserAdmin pre-loads every user's
  password into combobox objects and into `Password_Edit` (`UserAdmin.pas:86,99`, `DataModule.pas:6194`).
  The edit is masked (`PasswordChar='*'`) so it's not visible on screen, but the cleartext sits in
  client process memory and travels the wire in the clear — recoverable via memory/SQL inspection. 🔴
- **Identity = (id + password) composite, not id alone.** Every proc keys on both. Consequences:
  - Two rows with the **same `VC_USER_ID` and different passwords** can coexist — the table has
    no PK/unique constraint and the dup-check (`InsertUser`) tests the *pair*. Login then matches
    whichever pair the user types.
  - `UPDATE`/`DELETE` silently no-op if the supplied old password doesn't match (no rows
    affected, no error surfaced).
- **Admin = single boolean.** `BIT_ADMIN` is the only role. It gates exactly one thing: whether
  the `Administration` top-menu is present (`MainMenu.pas:456`). All other features are open to
  any authenticated user. There is **no per-feature permission model**.
- **Login is case-forced uppercase** — `Logon.EditKeyPress` (`Logon.pas:92`) and
  `UserAdmin.EditKeyPress` (`:81`) uppercase a–z keystrokes. Both id and password are entered
  uppercase; combined with case-insensitive collation (`SQL_Latin1_General_CP1_CI_AS`), the match
  is effectively case-insensitive. 🟠 This shrinks the password keyspace.

### Password expiry / forced reset (partial / broken)
`ValidateUser` (`:6093-6140`) reads two fields off the `SELECT_UserInfo` result —
`UPDATEDIFF` and `IN_PASSWORD_RESET_DAYS` — and if `UPDATEDIFF >= IN_PASSWORD_RESET_DAYS` (or
NULL) it shows "Your password has expired", opens `NewPassword` (`:6101`), and calls
`UpdateUserPassword`.
- **But the snapshot `SELECT_UserInfo` returns only 3 columns** (`VC_USER_ID, VC_PASSWORD,
  BIT_ADMIN`) — no `UPDATEDIFF`, no `IN_PASSWORD_RESET_DAYS`. And `INV_USERS` has no
  reset-days column. So either (a) the live DB has a newer `SELECT_UserInfo`/`INV_USERS` than
  the snapshot, or (b) this whole expiry branch errors/never fires today.
  → Flag as **[[reference-schema-snapshot-vs-live]] mismatch — verify live.** The expiry+reset
  path (`NewPassword`, `UpdateUserPassword`, the reset-days column) is **unverified against a
  live DB**; the snapshot cannot run it (missing columns AND missing `UPDATE_UserPassword` proc).
- `NewPassword` (`NewPassword.pas:52`) only checks new == confirm and non-empty. **No complexity,
  length, or history rules.**

### Other
- `UserAdmin.Insert` requires id + password + confirm all non-blank and password==confirm
  (`UserAdmin.pas:115-132`); on dup (`InsertUser` returns False) shows "already exist".
- `UserAdmin.Update` re-reads the *original* id/pass from the selected combobox object as the
  WHERE key (`UserAdmin.pas:146-147`) — so renaming a user works, but only if that exact
  original pair still matches a row.
- `LogActLog` runs on **every** login/admin action with `gobjUser.AppUserID`, Windows username,
  machine name (audit trail in the Activity DB).

### P12 retry-recursion check (cross-cutting)
The user procs **retry themselves** on exception (`InsertUser`→`InsertUser` `:6287`,
`UpdateUserInfo`→`UpdateUserInfo` `:6381`, `DeleteUserInfo`→`DeleteUserInfo` `:6425`,
`SetComboBoxesWithUserObj`→itself `:6207`). These are **correct-target** retries — **no new
P12 wrong-target bug** in the auth group. (Self-retry is still the at-most-3 transient-retry
pattern; benign for idempotent user CRUD.) Confirmed against
`docs/analysis/cross-cutting/datamodule-retry-target-bugs.md` — none of the user procs are
listed there, consistent with this finding.

## 5. UI / UX notes
- Logon: id + password, Logon/Cancel. Cancel → app terminates.
- `ConfirmPassword.pas` ("re-enter your password to enter User Admin") is **wired but
  commented out** — `Administration_UserAdmin_MenuItemClick` (`MainMenu.pas:551-560`) has the
  whole confirm-password gate commented; UserAdmin opens directly. The form still compiles and
  contains a working `gobjUser.AppUserPass` plaintext compare (`ConfirmPassword.pas:50`) and
  logs `USER ADMIN` / `ADMIN ERR`. Treat as **dormant** (not dead — it ships, just isn't reached).
- UserAdmin: combobox of existing users; selecting fills id/password(visible)/admin;
  Insert/Update/Clear/Delete/Close.

## 6. Target design (Ignition)
- **Replace the entire `INV_USERS`/plaintext model with Ignition gateway security.** Users,
  passwords (hashed by the IdP), and roles live in an Ignition **User Source** (internal,
  or AD/LDAP/OAuth/SAML in prod). Perspective sessions carry the authenticated identity; no
  app-managed password table.
- **Roles, not a single admin bit.** Map `BIT_ADMIN=1` → an `Admin` role; everything else →
  a base `User` role. Use **Security Levels / role-based component security** to gate the
  Administration views (the rebuilt equivalent of removing the menu). This is the natural place
  to introduce the **per-feature permission model** the legacy lacks.
- **D1 multi-site = user↔site binding.** Per [[decisions]] D1, auth **binds each user to a
  site** (the "current site" replaces the single-install INI identity). Implement as a
  per-user site claim/role (e.g. role `Site:NUMMI`) or a `users.site_id` mapping in the User
  Source, enforced by Perspective session scoping. A user sees/operates only their site's data.
- **Audit:** keep the login/admin audit trail — log to the Activity equivalent (gateway audit
  profile or an `audit_log` table) with the Perspective username + client address.
- **Password policy:** enforce length/complexity/expiry at the User Source / IdP, not in app
  code. The legacy expiry intent (reset-days) becomes a gateway policy.
- **User CRUD view:** a Perspective admin view backed by Ignition's user-management API (or
  Named Queries against the new user table) — no plaintext ever shown.

## 7. Migration plan
- [ ] Stage 1 — stand up Ignition User Source; seed it from `INV_USERS` (one-time import;
      passwords cannot be recovered as hashes, so **force a reset on first login**).
- [ ] Stage 2 — map admin bit → roles; gate Administration views by role; bind users to sites.
- [ ] Stage 3 — retire `INV_USERS`, `SELECT/INSERT/UPDATE/DELETE_UserInfo`, the (missing)
      `UPDATE_UserPassword`, and the plaintext UI entirely.

## 8. Open questions for the user
1. **Live `INV_USERS` shape (D-level, blocks the expiry path):** does production have the
   `UPDATEDIFF`/`IN_PASSWORD_RESET_DAYS` columns and a real `UPDATE_UserPassword` proc, or is the
   forced-reset feature effectively dead in the field? (Snapshot can't run it.)
2. **Same-id-different-password rows:** does this ever happen operationally, or should the
   rebuild enforce one row per user id? (Recommend: unique user id, full RESTRICT on delete per D3.)
3. **Roles:** is a single Admin/User split enough, or do you want a finer per-feature permission
   set in Ignition (e.g. EDI-only, receiving-only operators)?
4. **On import we cannot migrate passwords (plaintext→hash is one-way the wrong direction is
   trivial, but we should NOT preserve plaintext):** OK to force every user to set a new password
   at first Ignition login?

## 9. Parity / regression checks
- Login with valid id+password → identity established + `LOGIN` audit row; admin menu present
  iff `BIT_ADMIN`.
- Invalid password → no identity, `LOGIN ERR` audit row.
- UserAdmin insert dup pair → "already exist"; insert new → row added, no plaintext leakage in
  the new UI.
- Update keyed on old pair → row mutated + 16-char `LastUpdated` stamp.
