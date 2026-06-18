# Module Analysis: Configuration, Site Identity, Shell (version gate + DATAPURGE)

**Area:** Admin / system / shell  **Status:** ✅ spec complete  **Analyst:** Claude / 2026-06-16

Covers `Configuration.pas`, `SiteInfo.pas`, the INI-backed config layer (`TCIniField` in
`DataModule`), and the **shell concerns** of `MainMenu.pas`: app startup/connection wiring,
the **version gate**, and the **DATAPURGE** retention mechanism. **This module OWNS the
single-site INI config that D1 (INI → `sites` table) replaces — HIGH VALUE.**

> Out of scope here (already documented): report handlers → `docs/analysis/reporting/`;
> P12 retry bugs / trigger reconciliation → `docs/analysis/cross-cutting/`.

---

## 1. Legacy surface
- **`Configuration.pas` (+`.dfm`)** — the in-app settings dialog (`TConfigurationDlg`), a
  4-tab page-control editing a subset of the INI keys. Reached from MainMenu's `Configure`
  flow / Administration menu.
- **`SiteInfo.pas`** — `TSiteInfo`, a **read-only data holder** (no logic, no INI reads). Its
  fields (name, address, DUNS, supplier code, EIN, EDI separators, TMM identity, max sequence,
  delivery-method code) are populated **from the DB** (`SiteDataSet` / `SiteTMMDUNSDataSet`),
  not from the INI. It is the in-memory "who am I as an EDI trading partner" object.
- **`MainMenu.pas`** — shell: single-instance lock, connection open, version gate, login,
  `Configure`, DATAPURGE scheduler. (Feature orchestration / report handlers excluded.)
- **`DirectorySelect.pas` / `SelectDateRange.pas`** — generic support dialogs (§5).

## 2. Configuration model
The app has **two** configuration stores:

### (A) INI file — single-install, machine-local (the D1 target)
Config is `TCIniField` components on the Data Module; each binds an INI **Section + Key +
Default** and reads/writes the app's `.ini` at runtime (component `Cinifld.pas` in
`Delphi-VCL-Components/NUMMI Version 7/NUMMI Tools/`). Field declarations:
`DataModule.pas:94-150`; Section/Key/Default bindings: `DataModule.dfm:186-690`.

**Full INI inventory** (Section · Key · default · D1 disposition):

| Field | Section | Key | Default | Rebuild home |
|-------|---------|-----|---------|--------------|
| `fiSupplierCode` | SITE | `SupplierCode` | `05680` | **`sites` col** (our supplier code; EDI `delSL[4]` D1 hook) |
| `fiPlantName` | SITE | `PlantName` | `NUMMI` | **`sites` col** |
| `fiAssemblerName` | SITE | `Assembler` | `WQS` | **`sites` col** |
| `fiPOEDISupport` | SITE | `POEDISupport` | `True` | **`sites` col** (feature flag) |
| `fiGenerateEDI` | SITE | `GenerateEDI` | `True` | **`sites` col** (feature flag) |
| `fiFileALC` | SITE | `FileALC` | `True` | **`sites` col** |
| `fiHighSequence` | SITE | `HighSequence` | `True` | **`sites` col** |
| `fiUseBCRatio` | SITE | `UseBCRatio` | `True` | **`sites` col** |
| `fiRevSeqLookup` | SITE | `RevSeqLookup` | `-30` | **`sites` col** (seq lookback days) |
| `fiExcelOrderSheet` | SITE | `ExcelOrderSheet` | `True` | **`sites` col** |
| `fiUseFirstProductionDay` | INIT | `UseFirstProductionDay` | `True` | **`sites` col** — load-bearing forecast/order week-offset flag (already flagged in forecasting/order specs) |
| `fiForecastUsageCompare` | INIT | `ForecastUsageCompare` | `7` | **`sites` col** |
| `fiUsageUpdateCompare` | INIT | `UsageUpdateCompare` | `14` | **`sites` col** |
| `fiHistoricalForecast` | INIT | `HistoricalForecast` | `12` | **`sites` col** |
| `fiFillDays` | INIT | `FillDays` | `23` | **`sites` col** (order fill horizon; UI caps ≤50, `Configuration.pas:206`) |
| `fiConfirmOrderFileCreation` | INIT | `ConfirmOrderFileCreation` | `FALSE` | **`sites` col** |
| `fiCreatePOPriorToClose` | INIT | `CreatePOPriorToClose` | (bool) | **`sites` col** |
| `fiLocalFTP` | INIT | `LocalFTP` | `False` | **`sites` col** |
| `fiHideTerminated` | DISPLAY | `HideTerminated` | `True` | **per-user/site UI pref** |
| `fiBuildOut` | DISPLAY | `BuildOut` | `True` | per-user/site UI pref |
| `fiTruckSeqLength` | DISPLAY | `TruckSeqLength` | `3` | **`sites` col** (seq formatting) |
| `fiPassSeqLength` | DISPLAY | `PassSeqLength` | `4` | **`sites` col** |
| `fiLogisticsInputDir` | DIRECTORIES | `LogisticsInputDir` | `c:\_Inventory_Control\` | **gateway/site path** (file-drop) |
| `fiForecastInputDir` | DIRECTORIES | `ForecastInputDir` | `...\Suppliers\NUMMI` | gateway/site path |
| `fiForecastFilename` | DIRECTORIES | `ForecastFilename` | `nummi.prelftp` | **`sites` col** |
| `fiLogisticsFilename` | DIRECTORIES | `LogisticsFilename` | `nummi.prelftp` | **`sites` col** |
| `fiReportsOutputDir` | DIRECTORIES | `ReportsOutputDir` | `...\Reports` | gateway path |
| `fiTextShippingFileDir` | DIRECTORIES | `TextShippingFileDir` | `D:\Daily Camex Results\` | gateway/site path |
| `fiEDIOut` | DIRECTORIES | `EDIOut` | `...\EDIOut` | gateway/site path |
| `fiEDIIn` | DIRECTORIES | `EDIIn` | `...\EDIIn` | gateway/site path |
| `fiTemplateDir` | DIRECTORIES | `TemplateDir` | (path) | gateway path (or app dir, see `fiUseApplicationDir`) |
| `fiUseApplicationDir` | DIRECTORIES | `UseApplicationDir` | (bool) | gateway setting — if true, templates load from exe dir (`DataModule.pas:711`) |
| `fiDatabaseName` | DATABASE | `DatabaseName` | `INVENTORY` | **gateway DB connection** |
| `fiInventoryConnection` | DATABASE | `InventoryConnection` | conn-string **w/ credentials** | **gateway connection** (secret) |
| `fiActivityConnection` | DATABASE | `ActivityConnection` | conn-string **w/ credentials** | **gateway connection** (secret) |
| `fiALCConnection` | DATABASE | `ALCConnection` | conn-string | **gateway connection** |
| DATAPURGE keys | DATAPURGE | (see §DATAPURGE) | | **gateway scheduled task config** |

🔴 **Security finding:** the `[DATABASE]` connection-string defaults baked into
`DataModule.dfm:406-427` use **`Integrated Security=SSPI` + a `User ID=Inventory` hint, NO password**;
the **live `.ini` overrides these with SQL-auth connection strings carrying real passwords** (git-ignored
per repo guardrails). The plaintext-password exposure is in the live `.ini`, not the DFM default. These must become Ignition
**gateway DB connection profiles** (credentials in the gateway secret store), never per-client
config.

**Three ADO connections** (`Inv_Connection`, `Act_Connection`, `ALC_Connection`) open from
these strings at startup. Note the **cross-DB coupling**: Activity DB (`Act_Connection`) is a
*separate database* used for all `LogActLog` audit writes — relevant to DATAPURGE (below).

### (B) DB-resident site identity (already multi-row-capable)
`TSiteInfo` (DUNS, EIN, EDI ISA/GS separators, TMM trading-partner DUNS, delivery method code,
max sequence, accept-any-order flag) is loaded from `SiteDataSet`/`SiteTMMDUNSDataSet` — i.e.
the **EDI trading-partner identity already lives in the DB**, not the INI. These are exactly the
attributes D1 expects to be per-site. Note the **site DUNS/EIN are the EDI `delSL[4]` D1 hook**
called out in the EDI specs.

### Configuration dialog (`Configuration.pas`)
`Execute` (`:109`) loads ~25 INI fields into editors, `ShowModal`, and on OK-with-changes
writes them back (`:162-211`). Tabs cover SITE/DIRECTORIES, EDI dirs, forecast/usage compares,
and a full **DATAPURGE** tab (`:144-152`, `ShowPurgeRateDetail` `:278`). Directory pickers use
`SelectDirectory` (`:214`). The dialog edits a *subset* of the INI keys; DATABASE connection
strings and most flags are **not** editable here (INI-file-only).

## 3. Shell startup & version gate (`MainMenu.pas`)
Startup sequence (`FormShow`, around `:415-546`):
1. **Single-instance lock** (`LockSignature` in DataModuleCreate, `DataModule.pas:723`) →
   terminate if already running.
2. **Version gate** (`:431-446`): `SELECT_ProgramVersion` (`Create Inventory.sql:7572` →
   `SELECT * FROM inv_program_version`) returns one `Program_Version` string. If it
   `<>` the running exe's `VersionInfo.GetVersion` → ShowMessage "please upgrade. Program
   terminating" + `Application.Terminate`. **Hard gate**: a DB/exe version skew blocks launch.
   - `INV_PROGRAM_VERSION` (`Create Inventory.sql:1673`) = single `PROGRAM_VERSION varchar(50)`
     column. Code reads field alias `Program_Version` (case-insensitive). `About.pas` shows
     `VersionInfo.GetVersion` (`MainMenu.pas:580`).
3. **Login** (`:449-455`) → see auth-users.md.
4. **Admin-menu gating** (`:456-457`).
5. `LogActLog('START', ...)` (`:460`); `Configure` (`:462`); **DATAPURGE** scheduler (`:467`).

## 4. DATAPURGE — data-retention mechanism
**Config (`[DATAPURGE]` INI):**
| Key | Field | Default | Meaning |
|-----|-------|---------|---------|
| `EnableDataPurge` | `fiEnableDataPurge` | `FALSE` | master on/off |
| `PromptDataPurge` | `fiPromptDataPurge` | `FALSE` | ask before running |
| `DataRetention` | `fiDataRetention` | `18` | **months** to keep (must be ≥12, see below) |
| `PurgeRate` | `fiPurgeRate` | `Daily` | Daily / Weekly / Monthly |
| `LastPurge` | `fiLastPurge` | `2008050612000000` | last-run timestamp (NummiTime) — **rewritten to the INI after each run** |
| `PurgeDayWeekly` | `fiPurgeDayWeekly` | `Monday` | weekday for Weekly |
| `PurgeDayMonthly` | `fiPurgeDayMonthly` | `1st` | `1st`/`15th`/`Last` for Monthly |

**Scheduler logic (`MainMenu.pas:467-543`), at each app startup:**
- If disabled → skip.
- Compute days since `LastPurge`. Daily: run if ≥1 day. Weekly: run if ≥7 days **and** today
  matches `PurgeDayWeekly`. Monthly: run if ≥28 days **and** today is the configured day
  (`1st`/`15th`/last-of-month via `MonthDays`/`IsLeapYear`). (Comment "Cheat make it easy".)
- If due and `PromptDataPurge` → confirm dialog; No → `exit` (skips, does **not** reschedule
  for later that day — next attempt is next launch).
- Run `AutoPurge`; on success set `LastPurge := now` and rewrite the INI; show "complete".
  On failure: "will try again on next program start up".

**`AutoPurge` (`DataModule.pas:6885`):**
- **Guard:** `if fiDataRetention.AsInteger < 12` → error "must be greater than 12", abort.
  (So retention is effectively ≥12 months.)
- Calls `DELETE_AutoPurge(@DataRentention := 0 - fiDataRetention.AsInteger)` — passes the
  retention as a **negative month offset**.

**`DELETE_AutoPurge`** — ⚠️ **CORRECTED 2026-06-17 vs the LIVE `CreateInventory.sql` (D9).** An earlier
draft (from the stale `Create Inventory.superseded-2026-06-01.sql`) described a `Purge.PurgeMode` flag
table — **that table and flag DO NOT EXIST in the live DB** (`OBJECT_ID('Purge') IS NULL`; zero `PurgeMode`
references in the live dump). The real live body:
- (1) **Pre-stamps termination:** `UPDATE INV_OPEN_ORDER_INF SET VC_TERMINATED = <cutoff date> WHERE …
  AND VC_TERMINATED = ''` — marks aged un-terminated open orders as terminated.
- (2) Then **deletes rows older than the cutoff** (`DATEADD(MONTH, @DataRentention, getdate())`, a past
  cutoff since the arg is negative) from **three tables**: `INV_OPEN_ORDER_INF`, `INV_OPEN_ORDER_INF_HIST`,
  `INV_PARTS_STOCK_MST_HIST`. Cutoff compares against each table's `VC_ADD` (16-char add-stamp). Each
  statement checks `@@error` and `RETURN`s on failure. **It does NOT delete `INV_REJECT_INF` or
  `INV_PART_SHIPPING_INF`.**
- **The on-hand-drain avoidance mechanism (the real one):** the pre-stamp sets `VC_TERMINATED <> ''`, and
  the `DELETE_RecConfStatPartsStockMstQTY` trigger's qty-subtraction is gated `… AND VC_TERMINATED = ''` —
  so the now-terminated aged rows are **already excluded** from the qty removal when the DELETE fires. There
  is NO PurgeMode flag/bypass; termination-before-delete is what keeps the purge from draining `IN_QTY`.

**Hazards / findings:**
- 🟠 **`DELETE_AutoPurge` is non-transactional** (still a real finding — the PurgeMode framing was wrong,
  the non-atomicity is right). No `BEGIN TRAN`; it pre-stamps `VC_TERMINATED` then deletes 3 tables with the
  `@@error`/`RETURN` pattern. A mid-run failure leaves a **partial delete + partially-stamped termination**,
  not rolled back. The rebuild's purge must be transactional (see D11#6, corrected).
- 🟠 **Cross-DB Activity coupling (commit-path hazard).** Purge only touches the Inventory DB,
  but every operation (including the purge itself, `LogActLog('PURGE',...)` `DataModule.pas:6921`)
  writes to the **separate Activity DB**. The audit and the data live in different databases
  with no distributed transaction — the audit row can persist while the data delete fails, or
  vice versa. This same Inventory↔Activity split is a general commit-atomicity hazard noted
  across the system; the purge inherits it.
- Only 3 tables are purged. Other history/transactional tables (rejects, stocktaking, shipping,
  ASN/invoice, forecast) **grow unbounded** — retention is partial, by design.
- Scheduling is **client-launch-driven**: it runs only when *someone starts the app* on the
  configured day. If no one launches it that day, the window is missed until the next launch.

## 5. Generic support dialogs
- **`DirectorySelect.pas`** — `SelectDirectoryDlg`, a folder-browser wrapper used by
  Configuration's `*_SpeedButton` handlers (`Configuration.pas:214-253`). Support only;
  becomes a server-side path setting in Ignition (no client folder-picker).
- **`SelectDateRange.pas`** — a from/to date-range picker used by report/query screens to
  bound a query. Support only; becomes a Perspective date-range component bound to a Named
  Query parameter.

## 6. Target design (Ignition) — **D1 is centred here**
- **INI → `sites` table + gateway config.** Per [[decisions]] **D1**, "all site info should now
  move from the global INI into the site table." Concretely:
  - **`sites` table columns:** every `[SITE]` and site-scoped `[INIT]`/`[DISPLAY]` flag from
    the inventory above (SupplierCode, PlantName, Assembler, the EDI feature flags
    `POEDISupport`/`GenerateEDI`, `UseFirstProductionDay`, `FillDays`, seq lengths, the
    forecast/usage compares, etc.), **plus** the DB-resident `TSiteInfo` identity (DUNS, EIN,
    ISA/GS separators, TMM DUNS, delivery-method code, max sequence). The two stores merge into
    one per-site row. Every `INV_*` table gains `site_id` (D1).
  - **Gateway config (not per-site rows):** the `[DATABASE]` connection strings → Ignition
    **DB connection profiles** with credentials in the gateway secret store; the
    `[DIRECTORIES]` file paths → gateway-level path settings (the new app does file I/O
    server-side, so client `c:\...` paths disappear); per-user UI prefs (`HideTerminated`,
    `BuildOut`) → Perspective user/session props.
  - **`Configuration.pas` → a Perspective "Site settings" admin view** (role-gated) editing the
    `sites` row, plus a gateway-admin area for connections/paths.
- **Version gate → gateway/project versioning.** The "DB version must equal client version"
  hard gate is obsolete under a single gateway-served Perspective app (one version of truth).
  Keep an optional schema/migration-version check at startup if desired.
- **DATAPURGE → an Ignition Gateway Scheduled Task** (or a DB Agent job) running a window-aware,
  **transactional** purge per `sites.data_retention_months`, scoped by `site_id`. Make the purge
  transactional (the legacy `DELETE_AutoPurge` is non-atomic: pre-stamp `VC_TERMINATED` + 3 deletes
  with `@@error`/`RETURN`, no `BEGIN TRAN` — a mid-run failure partial-deletes; no PurgeMode flag) and decide whether
  to extend retention to the other history tables. The Activity/audit write becomes part of the
  same gateway transaction context or an idempotent audit log.

## 7. Migration plan
- [ ] Stage 1 — model the `sites` table from this INI inventory + `TSiteInfo`; seed one row
      (the current single site) from the live `.ini` + `SiteDataSet`.
- [ ] Stage 2 — move DB connections + file paths to gateway config; build the role-gated Site
      settings view.
- [ ] Stage 3 — reimplement DATAPURGE as a transactional scheduled task scoped by `site_id`;
      retire the version gate.

## 8. Open questions for the user
1. **DATAPURGE retention scope:** today only `INV_OPEN_ORDER_INF`, `..._HIST`, and
   `INV_PARTS_STOCK_MST_HIST` are purged. Should the rebuild keep that exact scope, or extend
   retention to rejects / stocktaking / shipping / ASN / forecast history?
2. **Per-site retention:** is `DataRetention` (months) a per-site setting, or one global policy
   across all sites?
3. ✅ RESOLVED (2026-06-17): there is **no `Purge.PurgeMode` flag** in the live DB (stale-snapshot
   artifact). The legacy avoids draining on-hand during purge by **pre-stamping `VC_TERMINATED`** before
   the delete (the qty trigger is gated `VC_TERMINATED=''`). The transactional rewrite preserves this
   ordering (terminate → delete) atomically; no flag to honor.
4. **DISPLAY prefs (`HideTerminated`, `BuildOut`, seq lengths):** per-site, per-user, or both?
5. **Site identity source of truth:** confirm the `SiteDataSet`/`TSiteInfo` DB values (DUNS,
   EIN, EDI separators, TMM DUNS) are the authoritative trading-partner identity to fold into
   `sites` (vs anything still read from the INI).

## 9. Parity / regression checks
- Each INI key read produces the documented default when absent (verify against
  `DataModule.dfm` defaults).
- Version mismatch → app refuses to launch (legacy) vs gateway single-version (rebuild — N/A).
- DATAPURGE with retention `<12` → rejected; with `18` and a stale `LastPurge` on the scheduled
  day → deletes only the 3 tables older than `now-18mo`, stamps `LastPurge`, leaves other
  tables intact.
