# M4 source-truth — Auth model, Sites master, cross-DB reads

Source-truth analysis for the non-schema-surgery half of M4 (security / multi-site / hardening).
The `site_id` schema surgery + `_HIST` lockstep is owned by the parallel sql-analyst pass; THIS doc
owns **auth**, the **Sites master (8th master)**, the **site-identity flip**, and **cross-DB reads**.
NO build — analysis to surface the M4 design + the decisions for David.

Confidence: all auth/site claims below were read from the live `.pas`/`.dfm` and proc bodies in
`DB Schema/CreateInventory.sql` (the 2026-06-12 authoritative dump). Bodies verified, not inferred.

---

## 1. The legacy auth/security model — IS IT OPEN?

**Answer: it is a thin, in-app, plaintext-password gate with a binary Admin/User split, and it is
trivially bypassable from the network. There is NO database-level, OS-level, or network-level
access control beyond the SQL Server connection string itself.**

### How a user gets in
- App start → `MainMenu.FormShow` hides the window and runs the **modal Logon form**
  (`MainMenu.pas:448-455`): `Logon_Form.Execute`. If it returns false → `Application.Terminate`.
- `Logon.bitBtnLogonClick` (`Logon.pas:53`) calls `Data_Module.ValidateUser(UserID, Password)`.
- `ValidateUser` (`DataModule.pas:6062`) runs proc `dbo.SELECT_UserInfo;1` with `@UserID/@Pass`.
- **`SELECT_UserInfo` validates by literal plaintext compare** (`CreateInventory.sql:1659-1663`):
  `... FROM INV_USERS, INV_PASSWORD_RESET_DAYS WHERE VC_USER_ID=@UserID AND VC_PASSWORD=@Pass`.
  A returned row = valid login. **Passwords are stored and matched in cleartext** —
  `INV_USERS.VC_PASSWORD varchar(30)` (`CreateInventory.sql:958-963`), no hash, no salt.
- On success the app caches the identity in the process-global `gobjUser` (`TInvUser`):
  `AppUserID`, `AppUserPass` (**plaintext password held in memory for the session**), `AppUserAdmin`
  (`DataModule.pas:6137-6139`).

### What a user can do — the ONLY authorization gate
- The entire authorization model is **one line**: `MainMenu.pas:456-457`
  ```
  If Not Data_Module.gobjUser.AppUserAdmin Then
    MenuBar_MainMenu.Items.Remove(Administration_MenuItem);
  ```
  Non-admins simply have the **Administration menu removed from the menubar**. Everything else —
  Order, Shipping, Receiving, Stocktaking, ASN/EDI create, all masters reachable outside the
  Administration menu — is open to **any authenticated user**. There is no per-feature permission
  matrix, no row-level security, no read-vs-write distinction. It is **all-or-nothing inside the
  Admin menu, fully-open everywhere else.**
- The Admin menu gates: **User Administration** (`UserAdmin`), **Configuration** (`Configuration1`),
  and (via Configuration) the INI-backed settings including DATAPURGE. The `ConfirmPassword` re-prompt
  that once guarded UserAdmin is **commented out / dead** (`MainMenu.pas:551-560`) — the menu removal
  is the sole guard now.

### Admin/User split (Q12 already satisfied in the legacy)
- `INV_USERS.BIT_ADMIN bit` is the only role flag. `UserAdmin` (Admin-only screen) does full CRUD on
  users with an `IsAdmin` checkbox (`UserAdmin.pas:86,121,147,166`) via procs `INSERT_UserInfo`
  / `UPDATE_UserInfo` / `DELETE_UserInfo` / `SELECT_UserInfo`. So the legacy **already is** the
  2-role Admin/User model Q12 settles on.

### First-login / expiry reset — the legacy ALREADY has a primitive version
- `INV_PASSWORD_RESET_DAYS` is a single-row table holding one int `IN_PASSWORD_RESET_DAYS`
  (`CreateInventory.sql:779-781`). `SELECT_UserInfo` cross-joins it and computes `UPDATEDIFF`
  = days since `LastUpdated` (`CreateInventory.sql:1653,1659`).
- In `ValidateUser` (`DataModule.pas:6096-6131`): if `UPDATEDIFF >= IN_PASSWORD_RESET_DAYS` **OR
  `UPDATEDIFF IS NULL`**, the app shows "Your password has expired", pops `NewPasswordDlg`, and on
  confirm calls `UPDATE_UserPassword` (sets `VC_PASSWORD` + `LastUpdated`, `CreateInventory.sql:1587-1589`).
- **Net-new users get a forced first-login reset for free**, because `INSERT_UserInfo`
  (`CreateInventory.sql:1255-1266`) does NOT set `LastUpdated` → it is NULL → `UPDATEDIFF.IsNull`
  → reset forced on the user's first login. So the legacy de-facto already does "Admin creates a user,
  user must set their own password on first login." Q13's first-login reset has a faithful precedent.

### Security posture / hazards (first-class findings)
- **H1 (CRITICAL) — plaintext passwords everywhere.** Stored plaintext in `INV_USERS.VC_PASSWORD`,
  matched plaintext in `SELECT_UserInfo`, displayed plaintext in the UserAdmin combobox
  (`UserAdmin.pas:99` puts the password into `Password_Edit` from the dropdown object), held plaintext
  in `gobjUser.AppUserPass`. Anyone with read access to `INV_USERS` (or the connection string) has
  every credential. Q13 explicitly retires this — must NOT carry forward.
- **H2 — the SQL connection string IS the real perimeter.** The app authenticates to SQL via INI
  connection strings *with passwords* (`fiInventoryConnection/fiActivityConnection/fiALCConnection`,
  loaded raw in `DataModuleCreate` `DataModule.pas:731-733`). Anyone on the LAN with that string (or a
  copy of the INI) bypasses the Logon form entirely and reads/writes the DB directly. The app login is
  cosmetic relative to the DB perimeter.
- **H3 — client-side authorization.** The Admin gate is a menu removal in the Delphi client
  (`MainMenu.pas:457`). A modified client, or direct DB access, ignores it completely. There is no
  server-side enforcement of the Admin/User split.
- **H4 — no lockout / no audit of the password value.** Failed logins are logged to the activity log
  (`LOGIN ERR`, `DataModule.pas:6145`) but there is no lockout, throttle, or complexity rule. The
  only password policy is the age-based reset.
- **H5 (P8/P12 retry bug, confirmed) — wrong-target recursion in the EIN-status path.**
  `UpdateEINStatus`'s exception handler retries by calling **`UpdateRecProdRejInfo`** (a different
  proc) instead of itself (`DataModule.pas:6789`). Same family as the 29 confirmed P8/P12 bugs; the
  rebuild must not port the recursion.

---

## 2. The Ignition auth TARGET (the design input)

Maps the legacy model + Q12/Q13/Q14 onto Ignition. (Target-platform design is the ignition-architect's
to finalize; this is the source-grounded input + the gating matrix.)

### User Source
- **Recommended: an Ignition Internal user source** (gateway-managed users + the 2 roles), seeded from
  `INV_USERS` at cutover, with **passwords NOT migrated** (they're plaintext — see H1; admin re-seeds or
  users self-set on first login). This is the lowest-friction faithful match to the existing 2-role model
  and keeps Q14's single gateway self-contained. A Database user source pointed at a *new, hashed*
  `INV_USERS_AUTH` table is the alternative if David wants users administered in-DB (closer to the legacy
  UserAdmin screen) — but it must hash (bcrypt/PBKDF2), never plaintext. AD/SSO is out of scope (Q12 =
  Admin/User is enough for now; no AD requirement surfaced).
- **Migration note:** seed user IDs + the `BIT_ADMIN→role` mapping only. Force a first-login reset for
  every migrated user (mirrors the legacy NULL-`LastUpdated` behavior, §1).

### The Admin / User role split (Q12) + per-feature gating
Two roles: **Admin** and **User**. The faithful cut is "what the legacy Administration menu gated" → Admin;
everything else → User. Concrete per-feature gating matrix (the `Administration_MenuItem` contents drive it):

| Feature | Legacy location | Role |
|---|---|---|
| User Administration (create/edit/delete users, set Admin) | Admin menu → `UserAdmin` | **Admin only** |
| Configuration / all settings (the INI-backed config) | Admin menu → `Configuration1` | **Admin only** |
| **Sites master** (the new 8th master) | (new, Admin-gated) | **Admin only** |
| DATAPURGE enable/prompt/retention + run | Configuration / startup | **Admin only** |
| Secrets / connection / path config | INI (was machine-level) | **Admin only** (gateway secret store) |
| Daily ASN create / 856 build | ASNSelect, ASNInvoice | **User** (any authenticated) |
| Order / RenbanOrder / OrderFormCreate | Order screens | **User** |
| Shipping / ManualShipping / ModifyShipping | Shipping screens | **User** |
| Receiving (RecConfStat / RecReject) | Receiving screens | **User** |
| Stocktaking / InvMgmt | Inventory screens | **User** |
| EDI upload / forecast import | EDIUpload | **User** |
| HotCall entry | HotCallEntry | **User** |
| Reports (read-only) | Report menu | **User** |
| The other masters (Supplier/Size/AssyRatio/PartsStock/Logistics/Renban/ManifestCost) | reachable outside Admin menu today | **User today** — see DECISION D-M4-2 |

- **Gating mechanism in Ignition:** drive Perspective component `meta.visible`/`enabled` and page-access
  off `session.props.auth.user.roles`, AND **enforce server-side** — the Admin-only writes (user CRUD,
  Sites CRUD, config, purge) must check role inside the Named Query / gateway script, not just hide the
  button. This closes legacy hazard H3 (client-only authz).

### First-login password reset (Q13, no plaintext)
- Faithful to the legacy flow (§1): Admin creates a user with a temporary credential; on first login the
  user is forced to set their own password; the stored value is **hashed** (the user source handles this
  for an Internal source). The legacy age-based expiry (`IN_PASSWORD_RESET_DAYS`) can be reproduced with a
  gateway password policy if David wants it, but the *first-login* reset is the Q13 commitment and is the
  faithful minimum.
- Do NOT port `UPDATE_UserPassword`/`SELECT_UserInfo` as-is — they assume plaintext.

### Session/site scoping
- See §3 — the logged-in session must resolve **which site** it is acting as; that is the multi-site core.

---

## 3. The site identity flip — single-site INI → per-site INV_SITES

### How the legacy gets its single site identity (the "from")
- **`TSiteInfo` (`SiteInfo.pas`) is a read-only property bag and is effectively VESTIGIAL in the live
  path** — it is declared, but it is **never instantiated** in `DataModule.pas` (no `TSiteInfo.Create`,
  no field of that type; grep confirms only `SiteDataSet`/`SiteTMMDUNSDataSet` are used). The doc
  comments in `spike-inv-sites-table.sql` map columns to `TSiteInfo` fields for convenience, but the
  actual runtime site identity does NOT flow through `TSiteInfo`.
- **The live site identity actually comes from TWO sources:**
  1. **The INI, via `TCIniField` properties** read straight into `DataModule` (`fiSupplierCode`,
     `fiEDIOut`, `fiEDIIn`, `fiForecastInputDir`, `fiLogisticsInputDir`, `fiReportsOutputDir`,
     `fiTextShippingFileDir`, `fiFillDays`, `fiForecastUsageCompare`, `fiUseFirstProductionDay`,
     the `fiEnableDataPurge/fiPromptDataPurge/fiDataRetention/...` purge block, the three connection
     strings). These are the *paths + behavioral knobs*. One INI per machine = single site.
  2. **The DB, via `SiteDataset` = proc `AD_GetSite`** (CommandText in `DataModule.dfm:452-454`,
     targets `VehicleOrder.Site` over the ALC connection) and `SiteTMMDUNSDataSet` = `AD_GetSiteTMMDUNS`
     (`DataModule.dfm:691-693`). These supply the **trading identity**: `SiteAbbr`, `SiteDUNS`,
     `SiteSupplierCode`, `SiteTMMDUNS`, `SiteEIN`, `SiteEDIMode`.
- **Proof the trading identity is DB-sourced, not INI-sourced:** the EDI builders read it from
  `SiteDataset` field-by-field at build time, e.g. `EDI810Object.pas:142-162`
  (`SiteAbbr` → ISA06 `%-10s`, `SiteDUNS` → ISA08/GS, `SiteSupplierCode` → ISA06 qualifier,
  `SiteTMMDUNS` → ISA12/GS, `SiteEIN` → control number, `SiteEDIMode` → ISA15 usage indicator).
  `EDI856Object.pas:140+` does the same. So **`AD_GetSite`/`AD_GetSiteTMMDUNS` is the precise relocation
  target into `INV_SITES`** — the rebuild repoints `SiteDataset` from `VehicleOrder.Site` to the
  session's `INV_SITES` row. (This corrects "reads no sites table": it reads it proc-mediated.)

### The EIN counter — the live per-site sequence (Q4)
- Every EDI/ASN producer does the same dance: `SiteDataset.Open` → read `SiteEIN` →
  insert ASN/INV with that EIN → `AD_UpdateEIN` to bump `Site.SiteEIN`. Confirmed call sites:
  `ASNSelect.pas:378-389` and `:446-471`, `HotCallEntry.pas:223-291`, `MainMenu.pas:2616-2636`,
  `Write810File.pas:58-74`, `Reports.pas:298,374`. The HotCall/810 paths use `SiteEIN+1`
  (`HotCallEntry.pas:251,272`; `MainMenu.pas:2619`); ASN reads `SiteEIN` as-is then bumps.
- **`INV_SITES.IN_EIN_SEQ` takes over `Site.SiteEIN` per-site (Q4, already modeled).** The bump
  (`AD_UpdateEIN`) becomes a site-scoped increment of `INV_SITES.IN_EIN_SEQ WHERE IN_SITE_ID=@site`,
  claimed atomically (SERIALIZABLE / `UPDATE ... SET ... OUTPUT`), since it is the EDI control number.

### How a running Ignition session knows WHICH site it is — the multi-site question (options)
The legacy answer is "the machine's INI is the site." The rebuild needs a session-level answer. Options:

- **Option A — Gateway-scoped (one gateway = one site), matches Q14's "single gateway" most literally
  IF each site runs its own gateway.** Site is a gateway System Property / single-row config; no picker;
  every read/write reads the gateway's site. *Simplest, but contradicts a single shared gateway serving
  both Tupelo + Huntsville — only works if Q14 means one-gateway-per-site, which conflicts with
  "single gateway."*  → **Needs David to clarify what Q14 "single gateway" means (see D-M4-1).**
- **Option B — Per-user site (user→site mapping).** Each Ignition user is bound to one site (a user
  attribute or a `INV_USER_SITE` map); on login the session resolves the site from the user. Faithful to
  "an operator works one plant," works on a single shared gateway, no picker. Multi-site admins would
  need either multiple accounts or admin-override.
- **Option C — Session site-picker.** On login the user picks a site (constrained to sites they may
  access); stored in `session.custom.siteId`. Most flexible (one admin spans sites); adds a step and a
  mis-selection risk for operators.
- **Recommended: B for operators + C-style override for Admins**, on ONE shared gateway (the multi-site
  goal). The session resolves `siteId` once at login; **every site-scoped read/write derives site from
  the session, never from a client parameter** (matches the `siteScopedQuery()` rule already baked into
  `spike-inv-sites-table.sql`'s header). This is a DECISION (D-M4-1).
- **Implication for every site-scoped read/write:** once `session.siteId` exists, it must flow into:
  - **the dir/path knobs** — `EDIOut`, `EDIIn`, `ForecastInputDir`, `LogisticsInputDir`,
    `ReportsOutputDir`, `TextShippingFileDir`, template dir — these were INI per-machine; in the rebuild
    they are per-site columns/derived-paths keyed by `siteId` (note: `INV_SITES` as built does NOT yet
    carry the directory columns — see §4 GAP).
  - **the EIN claim** — `IN_EIN_SEQ` per `siteId` (Q4).
  - **the trading identity** — `VC_DUNS`, `VC_SITE_ABBR`, `VC_SUPPLIER_CODE`, `VC_TMM_DUNS`,
    `VC_EDI_MODE`, separators — the `SiteDataset` replacement, keyed by `siteId`.
  - **every site-scoped table read/write** — the `site_id` schema surgery (sql-analyst's half) must use
    the same `session.siteId`.

---

## 4. The Sites master screen (8th master, Admin-gated)

`INV_SITES` is built (`docs/analysis/master-data/spike-inv-sites-table.sql`, 32 logical config columns +
identity/audit). The screen is the CRUD editor over it, following the existing master-screen pattern
documented in `docs/analysis/master-data/` (`IGNITION-master-crud-design.md`,
`master-crud-namedqueries.sql`, the `perspective-views` set). Admin-gated (§2 matrix).

### What it must edit (grouped by the column blocks in the built table)
- **Identity/address** — `VC_SITE_NAME`, `VC_SITE_ABBR`, `VC_STREET/CITY/STATE/COUNTRY/ZIP`.
- **EDI / trading identity** — `VC_DUNS` (indexed routing key), `VC_SUPPLIER_CODE`, `VC_DOCK_CODE`,
  `IN_EIN_SEQ` (the live EDI control counter — edit with care; normally machine-bumped, not hand-edited),
  `VC_EDI_MODE` (single char P/T — ISA15), `VC_SEP_SEGMENT/ELEMENT/SUBELEMENT` (single chars),
  `VC_TMM_NAME/ABBR/DUNS`, `IN_MAX_SEQUENCE`, `BIT_ACCEPT_ANY_ORDER_ASN`, `VC_DELIVERY_METHOD_CODE` (TD5).
- **Order/forecast config** — `IN_FILL_DAYS` (≤50), `IN_FORECAST_USAGE_COMPARE`,
  `BIT_USE_FIRST_PRODUCTION_DAY`, `VC_FORECAST_IMPORT_MODE` (AUTO|MANUAL), `VC_LAST_FORECAST_IMPORT`.
- **Data purge (Q17)** — `BIT_ENABLE_DATA_PURGE`, `BIT_PROMPT_DATA_PURGE`, `IN_DATA_RETENTION` (≥12).
- **Audit** — `VC_LAST_UPDATE`, `VC_ADD` (house 16-char stamps, screen-managed).

### Validation (enforce in the screen + keep the table CHECKs)
- `IN_FILL_DAYS ≤ 50`, `IN_DATA_RETENTION ≥ 12`, `VC_FORECAST_IMPORT_MODE IN (AUTO,MANUAL)` (table
  CHECKs already exist; the screen should validate up front).
- **`VC_EDI_MODE` must be exactly 1 char** — it is emitted verbatim into the positional ISA15
  (`EDI856Object.pas`/`EDI810Object.pas:162`); a 2+ char value produces a malformed ISA. (The seed
  already corrected an earlier 'PROD' → 'P'.) Validate length=1.
- **Separators must be exactly 1 char each** (ISA/GS structural — `char(1)`).
- **`VC_DUNS` / `VC_TMM_DUNS` are real trading identifiers** — validate format (9 or 9+4) but
  **NEVER commit real DUNS/EIN to the repo**; the spike seed uses placeholders (`'000000001'` etc.,
  `spike-inv-sites-table.sql:200-218`). Real values load from the legacy site config at cutover only.
- `VC_SITE_ABBR` is the EDI ISA sender ID (`%-10s`) — uppercase, ≤10.

### GAP to flag (Sites table is missing the directory/path columns)
- The legacy site identity includes the **per-machine directory paths** (`EDIOut`, `EDIIn`,
  `ForecastInputDir`, `LogisticsInputDir`, `ReportsOutputDir`, `TextShippingFileDir`, template dir,
  `LocalFTP`) read from the INI `[DIRECTORIES]`/`[INIT]`. `INV_SITES` **as built does NOT contain these
  columns.** For a true multi-site app, each site's output/input shares differ, so the path resolution
  must be per-site. **DECISION D-M4-3:** add per-site path columns to `INV_SITES` (or a child
  `INV_SITE_PATHS` table), OR resolve paths as `<base>/<site_abbr>/...` conventionally. Either way the
  Sites screen / path-resolution must be site-aware. This is in-scope for the Sites master design.

---

## 5. Cross-DB reads — disposition (stay / move / site-scope)

Per `vehicleorder-sites-verification.md` (post-restore: the real shared table is
`VehicleOrder.dbo.Site`, singular, 2 rows; InventorySystem reads it via the `AD_*` procs):

| Cross-DB read | Mechanism / evidence | Disposition |
|---|---|---|
| **`Site` row (trading identity, EIN)** | `AD_GetSite` (`DataModule.dfm:454`), consumed by EDI builders `EDI810Object.pas:142-162` etc. | **MOVE → `INV_SITES`.** Repoint `SiteDataset` from `VehicleOrder.Site` to the session-site row in Inventory. The single biggest relocation. |
| **`SiteTMMDUNS`** | `AD_GetSiteTMMDUNS` (`DataModule.dfm:693`) | **MOVE → `INV_SITES.VC_TMM_DUNS`** (already a column). Folds into the same site row. |
| **EIN counter bump** | `AD_UpdateEIN` (call sites §3) bumps `Site.SiteEIN` | **MOVE → site-scoped `INV_SITES.IN_EIN_SEQ`** (Q4). Atomic per-site claim. |
| **`LINE` (production lines)** | `VehicleOrder.Line` via ALC conn; heavy order/forecast use (`DataModule.pas:789` `LineName`) | **STAY shared in VehicleOrder (Q9).** It's the shared line master keyed by `LineName`; `INV_SITES` references it for the site↔line mapping. Read-only cross-DB read kept; **may need site-scoping of WHICH lines a site shows** (a site→lines filter), but the table stays put. |
| **`AD_GetSpecialDate` / `AD_GetSpecialDates` (calendar)** | `DataModule.dfm:552`; blocks order forecast-fill | **STAY shared, read-only (Q9).** Shared production calendar keyed by `LineName`. Keep the cross-DB read. |
| **Reports via `VehicleOrderConnection`** | Admin/reporting cross-DB reads | **STAY (read-only)** for the M3 report family; site-scope the *result filtering* by the session site where the report is site-specific. |
| **DUNS reads** | `SiteDUNS`/`SiteTMMDUNS` from `AD_GetSite`/`AD_GetSiteTMMDUNS` → ISA/GS | **MOVE → `INV_SITES.VC_DUNS`/`VC_TMM_DUNS`** with the Site row. DUNS is also the **inbound EDI routing key** (Q7/Q11) — the inbound poller resolves the target site by matching the inbound DUNS against `INV_SITES.VC_DUNS` (indexed `IX_INV_SITES_DUNS`). This is how a single shared gateway routes an inbound 830/856-ack/etc. to the right site. |

**Net:** the **trading-identity reads (`Site`, `SiteTMMDUNS`, EIN, DUNS) MOVE into `INV_SITES`**; the
**shared masters (`LINE`, the calendar) STAY shared in VehicleOrder** read-only (Q9); the **physical
`VehicleOrder.Site` is NOT dropped** (siblings GALC/MES/Admin still read it — cross-system decision per
`vehicleorder-sites-verification.md`). InventorySystem simply stops reading it and reads `INV_SITES`.

---

## 6. Hardening items for the M4 plan (lighter scope)

- **HD1 — Secrets / connection strings → gateway.** The three INI connection strings *with passwords*
  (`fiInventoryConnection/Activity/ALC`, `DataModule.pas:731-733`) become Ignition **DB connections**
  in the gateway config; credentials live in the gateway secret store, never in repo/INI. (Closes H2.)
- **HD2 — Plaintext passwords retired.** `INV_USERS.VC_PASSWORD` plaintext is NOT carried forward;
  user source hashes (§2). Do not migrate password values; force first-login reset. (Closes H1.)
- **HD3 — Server-side authz.** Enforce the Admin/User split inside Named Queries / gateway scripts for
  all Admin-only writes (user CRUD, Sites CRUD, config, purge), not just by hiding UI. (Closes H3.)
- **HD4 — Per-site, transactional DATAPURGE (Q17).** Legacy `DELETE_AutoPurge` (`CreateInventory.sql:7682`)
  purges by `@DataRentention` only with **NO site filter** (it deletes from `INV_OPEN_ORDER_INF`,
  `INV_OPEN_ORDER_INF_HIST`, `INV_PARTS_STOCK_MST_HIST`). In a shared multi-site DB this would purge
  ALL sites. M4: **add a site filter (`WHERE ... AND IN_SITE_ID=@site`)** and wrap transactionally; the
  enable/prompt/retention knobs move from the INI to `INV_SITES` per-site (Q17 — already columns in the
  built table). Also port the `>= 12 months` floor (`DataModule.pas:6890` / table CHECK).
- **HD5 — EIN-status update must be site-scoped (BLOCKER-2).** `UPDATE_EINStatus` (`CreateInventory.sql:1711-1730`)
  updates `INV_ASN_MST`/`INV_INV_MST` `WHERE IN_ASN_EIN=@EIN` / `IN_INV_EIN=@EIN` with **no site column**.
  Because each site has its own `IN_EIN_SEQ` counter, two sites can mint the **same EIN value** → a status
  update would hit the wrong site's row. M4 must re-key it `(IN_SITE_ID, EIN)`. (Also drop the P8 wrong-target
  retry, H5.)
- **HD6 — Path resolution per-site.** The INI directory knobs (EDIOut/EDIIn/forecast/logistics/reports/
  shipping/template) must become per-site (see §4 GAP / D-M4-3) and resolve from the session site.
- **HD7 — Backup runbook + redundancy decision.** Document gateway + DB backup/restore (the spike's
  cutover artifacts give a starting point); redundancy is a David decision (D-M4-5).
- **HD8 — Activity log carries the acting user + site.** Legacy `LogActLog` already stamps `gobjUser.AppUserID`
  / Windows user / IP / machine (`DataModule.pas:6020-6043`). Preserve the audit, and **add the acting
  `siteId`** so the shared multi-site log is attributable per site.

---

## 7. Top DECISIONS for David (auth + site-determination)

- **D-M4-1 (site determination — the multi-site core).** What does Q14 "single gateway" mean for site
  identity? Pick: **(A)** one gateway per site (gateway-scoped site, no picker — but then it's not one
  shared gateway); **(B)** one shared gateway, **per-user site** (operator bound to one plant);
  **(C)** one shared gateway, **session site-picker**. *Recommended: B for operators + admin override,
  on one shared gateway.* Everything site-scoped derives from `session.siteId`.
- **D-M4-2 (gating granularity).** Confirm the Admin/User cut (§2 matrix). Specifically: in the legacy,
  the **non-Admin masters** (Supplier/Size/AssyRatio/PartsStock/Logistics/Renban/ManifestCost) are open
  to any user (reachable outside the Administration menu). Keep them **User-editable** (faithful), or
  promote master edits to **Admin-only** in the rebuild (tighter, but a divergence)? *Faithful = keep as
  User; flag if you want them Admin-gated.*
- **D-M4-3 (per-site directory paths).** `INV_SITES` lacks the directory/path columns the legacy held in
  the INI. Add per-site path columns / an `INV_SITE_PATHS` child, OR derive paths conventionally
  (`<base>/<site_abbr>/`). Needed before EDIOut/forecast/logistics paths can be site-scoped.
- **D-M4-4 (password policy beyond first-login).** Q13 fixes first-login reset + no plaintext. Also port
  the legacy **age-based expiry** (`IN_PASSWORD_RESET_DAYS`), or drop it? And set a complexity policy
  (the legacy has none)?
- **D-M4-5 (user source + redundancy).** Internal Ignition user source (recommended, simplest, Q14)
  vs a Database user source (hashed `INV_USERS_AUTH`, closer to the legacy in-DB UserAdmin)? And the
  gateway/DB redundancy posture (none today) — single instance acceptable for go-live, or HA required?

---

### Evidence index (file:line)
- Legacy auth: `Logon.pas:53`, `MainMenu.pas:448-457,551-560`, `DataModule.pas:6062-6160` (ValidateUser),
  `:6096-6131` (forced reset), `:6298-6336` (UpdateUserPassword), `UserAdmin.pas:86-172`,
  `ConfirmPassword.pas:50` (dead re-prompt), `NewPassword.pas:52-66`, `UserInfo.pas` (TUserInfo),
  `DataModule.pas:42-53` (TInvUser), `:719` (gobjUser create).
- Auth procs/table: `CreateInventory.sql:958-963` (INV_USERS, plaintext), `:779-781`
  (INV_PASSWORD_RESET_DAYS), `:1643-1665` (SELECT_UserInfo), `:1255-1266` (INSERT_UserInfo, no LastUpdated),
  `:1576-1589` (UPDATE_UserPassword), `:1602-1620` (UPDATE_UserInfo).
- Site identity / EDI: `SiteInfo.pas` (vestigial TSiteInfo), `DataModule.dfm:452-454` (AD_GetSite),
  `:691-693` (AD_GetSiteTMMDUNS), `:552` (AD_GetSpecialDates), `EDI810Object.pas:142-162` (ISA from SiteDataset),
  `EDI856Object.pas:140+`, EIN call sites `ASNSelect.pas:378-471`, `HotCallEntry.pas:223-291`,
  `MainMenu.pas:2616-2636`, `Write810File.pas:58-74`, `Reports.pas:298,374`.
- Hardening: `DataModule.pas:731-733` (raw conn strings), `:6753-6794` (UpdateEINStatus + P8 retry bug),
  `CreateInventory.sql:1711-1730` (UPDATE_EINStatus, no site), `:7682-7723` (DELETE_AutoPurge, no site).
- Built artifact: `docs/analysis/master-data/spike-inv-sites-table.sql` (INV_SITES, 32 config cols).
