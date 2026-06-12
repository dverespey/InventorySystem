# Module Analysis: Size Master

**Area:** Master data  **Status:** ✅ spec complete  **Analyst:** Claude / 2026-06-04

> Third master analyzed, after Supplier and Logistics — same master-detail CRUD shape.
> "Size" here = a **tire/wheel size category** (e.g. a tire diameter/width class) that each
> part is assigned to via `INV_PARTS_STOCK_MST.IN_SIZE_ID`. The size carries two planning
> numbers — **daily usage** and **safety days** — that feed inventory/forecast math
> elsewhere. The narrowest master so far: only **4 user-visible columns** and **no FK-lookup
> combos**. The headline find is a **copy-paste bug in the insert dup-check** (see §3).

## 1. Legacy surface
- **Form:** `SizeMaster.pas` (~7 KB) + `SizeMaster.dfm` (`TSizeMaster_Form`, Caption
  "Size Master", header label "Size Master"). Author: Aaron Huge, 2002-10-25. Registered live
  in `InventorySystem.dpr` **line 16** (`SizeMaster in 'SizeMaster.pas' {SizeMaster_Form}`).
- **Entry point:** **not reached directly from `MainMenu.pas`** — it is reached through the
  master-maintenance hub. `MainMenu.pas` (line 363-365) opens `MasterMaint_Form`; that hub form
  (`MasterMaint.pas`, dpr line 13) has a `SizeMaster_Button` whose `OnClick`
  (`SizeMaster_ButtonClick`, line 106-113) does `SizeMaster_Form := TSizeMaster_Form.Create(self);
  SizeMaster_Form.Execute; SizeMaster_Form.Free;`. (Supplier and Logistics are reached the same
  way through the same hub.) `Execute` runs `ShowModal`; returns `False` only on `mrCancel` (the
  Close button, `ModalResult = 2 = mrCancel`).
- **Purpose (one paragraph):** Classic master-detail CRUD screen. A read-only-by-convention
  `DBGrid` (`SizeMaster_DBGrid`, `dgRowSelect`) lists all sizes; a small edit panel
  (`SizeMaster_Panel`) shows the selected row's four fields. Buttons: **Insert, Update, Delete,
  Search, Clear, Close**. `FormCreate` clears the dataset filter, calls `GetSizeInfo`, binds
  `Size_DataSource` to `Inv_DataSet`. Selecting a grid row (`OnMouseUp` / `OnKeyUp` /
  `Size_DataSource.OnDataChange` all call `HoldDetails(True)` + `SetDetailBoxes`) copies the row
  into the edits and captures `RecordID` from hidden grid `Fields[4]` (the identity PK).

## 2. Data touched
| Table | Read | Write | Notes |
|-------|:----:|:-----:|-------|
| `INV_SIZE_MST` | ✓ | ✓ | The tire/wheel size master (this module owns it) |
| `INV_PARTS_STOCK_MST` |  | ✓* | *Indirect: DELETE trigger nulls `IN_SIZE_ID` (FK unlink) |

Like Logistics and unlike Supplier, this form pulls in **no FK-lookup combos** — Size is a
**leaf master**: it is *referenced by* parts but references nothing itself. It is also the
**simplest** master: no address/contact block at all, just code + name + two integers.

### `INV_SIZE_MST` columns (authoritative: `DB Schema/Create Inventory.sql`)
| Column | Type | Meaning / notes |
|--------|------|-----------------|
| `IN_SIZE_ID` | `int IDENTITY(1,1) NOT NULL` PK | Surrogate key (`RecordID` in UI, hidden grid `Fields[4]`) |
| `VC_SIZE_CODE` | `varchar(6) NOT NULL` | **Business key — the size code; UNIQUE via `IX_INV_SIZE_MST`** (today global; becomes **per-site composite (site_id, VC_SIZE_CODE)** under D1). Form caps input at `MaxLength=6` (matches DB). `CharCase=ecUpperCase` |
| `VC_SIZE_NAME` | `varchar(50) NOT NULL` | Human-readable size name. Form `MaxLength=50` (matches DB), uppercased |
| `IN_USAGE` | `int NULL` | **Daily usage** — planning quantity per day. Form is a `TMaskEdit` `'99999;1; '`, `MaxLength=5` ⚠️ (form caps at 5 digits / max 99999; DB `int` allows far more) |
| `IN_DAYS` | `int NULL` | **Safety days** — days of safety stock. Same `TMaskEdit` cap of 5 digits ⚠️ |
| `VC_LAST_UPDATE` | `varchar(16) NULL` | **Timestamp as `yyyymmddHHMMSSff` string** (set on UPDATE only); not exposed on form (P2). **Note the column name is `VC_LAST_UPDATE` (with underscore)** — differs from Supplier's `VC_LASTUPDATE` and Logistics' `VC_LASTUPDATE` ⚠️ |
| `VC_ADD` | `varchar(16) NULL` | **Timestamp as `yyyymmddHHMMSSff` string** (set on INSERT only); not exposed on form (P2) |

**Difference vs Supplier / Logistics:**
- **Far fewer columns.** No address/city/state/zip/country, no tel/fax/person/email, **no
  `VC_BREAKDOWN_ORDER_DIRECTORY`** (so **no local-filesystem-path desktop-binding problem here** —
  the one multi-site headache the other two masters carry is simply absent). No order-file /
  enum columns, so **no P4 (coded single-char enums)** and **no FK combos**.
- **Two real planning integers** (`IN_USAGE`, `IN_DAYS`) — the other masters have no numeric
  payload. These are nullable `int` on the table but the form coerces blanks to `0`
  (`TextChange` forces `'0'`, `HoldDetails` does `StrToInt`), so blanks never reach the DB as NULL
  from this form (they arrive as 0).
- **Timestamp column is named `VC_LAST_UPDATE`** (underscore) vs `VC_LASTUPDATE` on Supplier and
  Logistics. A trivial-looking naming drift, but it matters when writing the AR column mapping.
- The audit string is `yyyymmddHHMMSSff` (**16 chars**): `CONVERT(char(8),…,112)` + **four**
  `SUBSTRING(CONVERT(varchar,…,114),p,2)` slices at positions 1/4/7/10 = HH+MM+SS+`ff` (trailing
  `ff` = first 2 digits of the millisecond component). This is the **byte-identical recipe used by
  Supplier *and* Logistics** — verified by reading all three proc bodies; there is **no format
  difference** among the masters (an earlier draft wrongly claimed 14 chars / no `ff`).

**Constraints / indexes:**
- `PK_INV_SIZE_MST` PRIMARY KEY CLUSTERED (`IN_SIZE_ID`).
- `IX_INV_SIZE_MST` **UNIQUE NONCLUSTERED (`VC_SIZE_CODE`)** — a **real DB unique backstop** on
  the code (today **globally** unique; under D1 this becomes **per-site composite (site_id,
  VC_SIZE_CODE)** once the table gains its `site_id` FK). **All three masters carry a real DB UNIQUE
  index** (Size on `VC_SIZE_CODE`, Supplier on `VC_SUPPLIER_CODE` via `IX_INV_SUPPLIER_MST`,
  Logistics on `VC_LOGISTICS_NAME`) — Size's posture **matches** Supplier's, it is not "stronger"
  (an earlier draft wrongly said Supplier's code had no DB uniqueness). No DEFAULT constraints.
- **No declared FOREIGN KEY constraints** out of this table (the whole schema declares only 2 FKs
  total, none involving Size). Inbound reference is **by convention only**:
  `INV_PARTS_STOCK_MST.IN_SIZE_ID` (and the history table `INV_PARTS_STOCK_MST_HIST` carries the
  same column shape) — no declared FK; the relationship is enforced only by the delete trigger.

**Triggers on these tables:**
- `DELETE_SizeCode` (on `INV_SIZE_MST` FOR DELETE; authoritative version in
  `DB Schema/Create Inventory.sql`): runs
  `UPDATE INV_PARTS_STOCK_MST SET IN_SIZE_ID = null FROM INV_PARTS_STOCK_MST a, DELETED d
  WHERE a.IN_SIZE_ID = d.IN_SIZE_ID`.
  **Invariant: deleting a size unlinks (does NOT delete) the parts assigned to it** — the parts
  rows survive with `IN_SIZE_ID = NULL`. Exact analogue of `DELETE_SupplierCode` /
  `DELETE_LogisticsCode` (P5).
- ⚠️ **Two source files disagree (same stale-trigger trap as Logistics §2).**
  `docs/triggers.sql` carries **stale variants** of both a `DELETE_SizeCode` and an
  `UPDATE_SizeCode` that key on a **string column `INV_PARTS_STOCK_MST.VC_SIZE_CODE`** (the
  stale DELETE does `SET VC_SIZE_CODE = ''`; the stale UPDATE propagates a renamed code into
  `VC_SIZE_CODE`). **That column does not exist on the live `INV_PARTS_STOCK_MST`** — parts link
  to size via **`IN_SIZE_ID int NULL`** (verified: the live `CREATE TABLE INV_PARTS_STOCK_MST`
  has `IN_SIZE_ID`, no `VC_SIZE_CODE`). These predate an int-FK refactor; **treat them as
  obsolete.** The **schema file is authoritative** and the live DB has exactly **one** trigger
  here — the `IN_SIZE_ID`-nulling DELETE — and **no UPDATE trigger** (correct: the identity PK is
  immutable, and parts key on the surrogate id, so a code rename needs no propagation).

## 3. Stored procedures used
(Read with `sql.sh proc NAME`. The procs are the behavioral spec.)

| Proc | Op | Business rule (from body) |
|------|----|---------------------------|
| `SELECT_SizeInfo;1 @SizeCode varchar(6) = ''` | SELECT | Single param. If `@SizeCode=''` returns **all** rows; else rows `WHERE VC_SIZE_CODE = @SizeCode`. Both branches select the same **5 UI-aliased columns** (`Size Code, Size Name, Daily Usage, Safety Days, RecordID`) `ORDER BY VC_SIZE_CODE`. No joins; **no name/code→id resolution** (the surrogate is only *returned*). `VC_LAST_UPDATE`/`VC_ADD` are **not** selected. Maps 1:1 to grid `Fields[0..4]` and `HoldDetails(True)`. |
| `INSERT_SizeInfo;1` (4 params) | INSERT | `@SizeCode varchar(6), @SizeName varchar(50), @Usage int, @Safety int`. Computes `VC_ADD` as a `yyyymmddHHMMSSff` string (P2) and inserts `(VC_SIZE_CODE, VC_SIZE_NAME, IN_USAGE, IN_DAYS, VC_ADD)`. `IN_SIZE_ID` is identity (not supplied, **not returned**). **No uniqueness check inside the proc** — the dup guard is meant to be in the app (P1), but it is **broken** (see Call-mechanism bug below). **Param widths match the table** (`@SizeCode varchar(6)` = DB `varchar(6)`, `@SizeName varchar(50)` = DB 50) — unlike Logistics, **no silent proc-level truncation** here. |
| `UPDATE_SizeInfo;1` (5 params, +`@SizeID int`) | UPDATE | Updates one row `WHERE IN_SIZE_ID = @SizeID` (surrogate passed directly; **no code→id resolution**). Sets `VC_LAST_UPDATE` (P2) and rewrites all four data columns **including `VC_SIZE_CODE`** — so the business key **is editable**. Does **not** touch `VC_ADD`. No uniqueness re-check in the proc — does not prevent renaming a code onto an existing one (the DB `IX_INV_SIZE_MST` unique index would, however, reject it with a raw SQL error). Param widths match the table. |
| `DELETE_SizeInfo;1 @SizeID integer` | DELETE | Hard-deletes `WHERE IN_SIZE_ID = @SizeID`. Single surrogate param. No soft-delete flag, no in-use / RI check inside the proc — relies entirely on the `DELETE_SizeCode` trigger to null `INV_PARTS_STOCK_MST.IN_SIZE_ID`. |

**Related, but NOT called by this form (listed for completeness, owned elsewhere):**
| Proc | Op | Where used |
|------|----|------------|
| `SELECT_SizeUsage` (no params) | SELECT | Joins `inv_size_mst s JOIN inv_parts_stock_mst p ON s.IN_SIZE_ID = p.IN_SIZE_ID`, returns `vc_size_code, in_usage, vc_part_number, vc_kanban_number` ordered by size code. **Called by `ForecastBreakdownF.pas` (line 964)**, not SizeMaster. Shows `IN_USAGE` is consumed by the forecast-breakdown math. |
| `UPDATE_SizeUsage @SizeCode varchar(6)='' , @Usage int` | UPDATE | `UPDATE INV_SIZE_MST SET IN_USAGE = @Usage WHERE VC_SIZE_CODE = @SizeCode`. **Called by `ForecastBreakdownF.pas` (lines 981, 1008)** — i.e. the forecast-breakdown screen **writes back daily-usage onto the size master by code**. Important cross-module coupling: the size's `IN_USAGE` is not edited only here; another module updates it (keyed **by code**, so a code rename there would silently miss). Flag for the rebuild. |

### Call mechanism (legacy)
`DataModule.pas` methods `GetSizeInfo / InsertSizeInfo / UpdateSizeInfo / DeleteSizeInfo`
(declarations 531-534; bodies 2487-2699) drive the shared ADO objects (P6). Notable per-method
facts:
- **`GetSizeInfo`** (lines 2487-2529) uses **`Inv_DataSet`** (open result set) with
  `@SizeCode := ''`, times the call (`fBeforeDateTime/fAfterDateTime/fDiffDateTime`), and on
  success logs `LogActLog('GET SIZE', 'SELECTED all SIZE info', 1)`. The other three use
  `Inv_StoredProc` (`ExecProc`).
- ⚠️ **`InsertSizeInfo` (lines 2531-2605) has a copy-paste bug in its dup-check (P1, broken).**
  It is two-step like the other masters, but **STEP 1 sets
  `ProcedureName := 'dbo.SELECT_AssyRatioInfo;1'`** (the **wrong** proc — Assembly-Ratio, not
  Size) while passing `@SizeCode := fSizeCode`. `SELECT_AssyRatioInfo` takes `@BroadCode
  varchar(12)='' `; ADO supplies the value positionally, so the size code is matched against
  `INV_ASSY_RATIO_MST.VC_BROADCAST_CODE`. It then inserts the new size only `If RecordCount = 0`.
  **Consequences:**
  1. For normal size codes (which are not broadcast codes) the assy-ratio query returns 0 rows →
     the insert proceeds — so insert *appears* to work.
  2. **The intended duplicate-size guard never runs.** Inserting an existing `VC_SIZE_CODE`
     is **not** caught in the app — it falls through to `INSERT_SizeInfo`, and the only thing
     that stops it is the **DB `IX_INV_SIZE_MST` unique index**, which raises a raw SQL error
     (`Inv_Connection.Errors.Count > 0` → `ShowMessage('Unable to insert Size data, …')` +
     `LogActLog('ERROR',…)` + `raise EDatabaseError`). The form's own friendlier message
     ("It already exists in the database.") on `If Not InsertSizeInfo` is **effectively dead**
     for the duplicate case, because the function returns `False` only via the `else
     fDescription := '(DUPLICATE)'` branch — which **can only be hit if a size code collides with
     an existing broadcast code**, a coincidental and wrong condition.
  3. **Edge hazard:** if a size code *does* happen to equal an existing broadcast code, the insert
     is **silently suppressed** (no row added) and reported as a duplicate — a false-positive
     "duplicate" with no actual size duplicate. Rare, but real.
  - **Net:** uniqueness for sizes is, in practice, enforced **only by the DB unique index**, not
    by the app dup-check. The rebuild should make this an explicit model validation (and the bug
    is a parity case to *not* reproduce — see §9).
- **`UpdateSizeInfo`** (2607-2656) sets the four fields + `@SizeID := FRecordID`; logs
  `LogActLog('UPDATE SIZ', 'UPDATE Size:'+fSizeCode, 1)`. No app uniqueness re-check (DB index
  backstops a rename collision, surfacing a raw error).
- **`DeleteSizeInfo`** (2658-2699) passes only `@SizeID := fRecordID`; logs
  `LogActLog('DELETE SIZ', 'DELETED Size Info:'+fSizeCode, 1)`.
- All four share the **retry-up-to-3-times-via-recursion** error harness (`fErrorCount < 3` →
  recursive self-call) with a `finally` doing `Inv_StoredProc.Close; fErrorCount := 0` (P8). On a
  hard error each raises a distinct `EDatabaseError` ("Unable to get/insert/update/delete Size
  data") after `ShowMessage` + `LogActLog('ERROR',…)`. The full `GET/INSERT/UPDATE/DELETE SIZ` +
  `ERROR` audit rows are real behavior to preserve as app logging.

**DataModule properties** (lines 436-439): `SizeCode / SizeName / DailyUsage(int) /
SafetyDays(int)`. The record key is the **shared, generic** `RecordID` property (line 337) — *not*
Size-specific; it is reused by Shipping/Invoice/ASN/Supplier/Logistics, so a stale `RecordID` from
another screen is a real (latent) cross-module hazard (P9).

## 4. Business rules & edge cases
- **Identity is `VC_SIZE_CODE`** (`MaxLength=6`, uppercased), backed by a **real DB UNIQUE index**
  (`IX_INV_SIZE_MST`) — today global, **per-site composite (site_id, VC_SIZE_CODE)** under D1
  (§8.1). The surrogate `IN_SIZE_ID` is the actual key for update/delete and the inbound part FK.
- **No form-level validation at all** (like Logistics, unlike Supplier). Lengths are enforced
  *only* by `TEdit.MaxLength`/`TMaskEdit`. There is **no min-length, not-blank, or required-field
  check** in Pascal — **inserting a blank code/name is not blocked by the form**. The DB `NOT
  NULL` columns are satisfied because the form always sends *some* string (possibly empty) and `0`
  for the integers; an empty `VC_SIZE_CODE` would insert once (the unique index permits a single
  `''`). (Contrast Supplier, which rejected `length < 5`.)
- **Uniqueness (P1, but broken in app):** the intended app-side dup check is **defective** (calls
  `SELECT_AssyRatioInfo` — see §3). In practice the **DB unique index is the only real guard**.
  Update/Delete never re-check uniqueness in the app; the DB index backstops a rename collision
  (raw error).
- **Code is editable** — `UPDATE_SizeInfo` rewrites `VC_SIZE_CODE`. Since parts link by
  `IN_SIZE_ID` (not by code) a rename does **not** break part links. But **code-based callers
  do**: `UPDATE_SizeUsage @SizeCode` and `SELECT_SizeUsage` (ForecastBreakdown), and the form's
  own client-side `Search` all key on the code — a name-as-key fragility (§8).
- **Daily Usage / Safety Days:** `TMaskEdit '99999;1; '`, `MaxLength=5` → **0..99999** only at the
  UI ⚠️ (DB `int` allows much more; another module could write larger via `UPDATE_SizeUsage`).
  `TextChange` forces an empty mask edit to `'0'`; `HoldDetails` does `StrToInt(Trim(...))`, so a
  truly blank value would raise on `StrToInt` — but the `'0'` defaulting prevents that. Net: the
  form **always sends 0, never NULL** for these (the column is nullable but the form won't write
  NULL).
- **Timestamps are `yyyymmddHHMMSSff` strings (P2):** `VC_ADD` on insert, `VC_LAST_UPDATE` (note the
  underscore) on update; computed in-proc from `getdate()`. Update preserves the original `VC_ADD`.
  Neither is shown on the form.
- **Delete is soft on parts** (trigger nulls `INV_PARTS_STOCK_MST.IN_SIZE_ID`, P5), **hard** on
  the size row. **Difference vs Logistics:** here the trigger *does* clean up the dependent table
  (parts), whereas the Logistics trigger left part FKs dangling. **But note**: the trigger only
  touches `INV_PARTS_STOCK_MST`; it does **not** address `INV_PARTS_STOCK_MST_HIST.IN_SIZE_ID`
  (if that column is populated, history rows could be left pointing at a gone size — confirm in §8).
- **Stale-`RecordID` hazard (P9):** Update/Delete key off `DataModule.RecordID` (grid `Fields[4]`),
  populated only by `HoldDetails(True)` when a row is selected. If no row was selected first,
  `RecordID` may be 0 or carry a leftover value from another module — no guard exists in the form.
- **Post-action refresh quirks:** after Insert the form re-`GetSizeInfo` and `Locate('Size Code',
  SizeCode_Edit.Text)`; after Delete it clears the code, re-gets, `SearchGrid(SizeCode)` (which
  filters by the now-deleted code → empty), then sets `Filtered := False`. Cosmetic, but worth
  matching only at the UX level.

## 5. UI / UX notes
- Grid + detail-panel pattern; selecting a row syncs the four edits and captures `RecordID`.
- **Search is fully client-side (P7):** `SearchGrid` sets the dataset `Filter := '[Size Code] = '
  + QuotedStr(fSizeCode)` and `Filtered := True` over the **already-loaded** `Inv_DataSet` — it
  does **not** re-query the DB. Match is an **exact equality** on the (uppercased) code, **no
  partial/LIKE**. (Slightly different mechanism from Supplier/Logistics, which looped the dataset
  manually; here it uses the dataset `Filter` property — same net behavior: exact match only.) On
  no match: "No matches were found for your query."
- **Fields (label → control → DB):** Size Code (`MaxLength=6`, uppercased), Size Name
  (`MaxLength=50`, uppercased), Daily Usage (`MaskEdit '99999;1; '`, 0-99999 ⚠️), Safety Days
  (same mask ⚠️). **No address/contact block, no directory picker, no combos** — this is the
  leanest master.
- **No format validation** beyond uppercase + MaxLength + the numeric mask.
- **`SizeCode_EditChange`:** when the code edit is `<= 1` char, it calls
  `Data_Module.ClearControls(SizeMaster_Panel)` — a typing convenience that blanks the panel.
- **Modernize:** standard index/list + new/edit form; **server-side search/sort/pagination** (P7);
  add **inline validation** (presence + uniqueness on code — fixing the legacy's broken dup-check,
  and a presence rule the legacy lacked); allow `IN_USAGE`/`IN_DAYS` beyond 99999 if the domain
  needs it (the DB already does); decide whether 0-vs-NULL matters for the two integers (§8).

## 6. Target design  *(Rails primary)*
- **Model:** `Size` (or `TireSize` to avoid clashing with the Ruby `Object#size` / common
  attribute name — recommend a non-conflicting class name like `TireSize`;
  `self.table_name = 'INV_SIZE_MST'`, `self.primary_key = 'IN_SIZE_ID'`).
  - `belongs_to :site` (D1) — every model belongs to a site with enforced current-site scoping;
    `site_id` is NOT NULL and all queries are scoped to the current site (auth binds users to a
    site). Rows are **per-site**.
  - `has_many :parts_stocks, foreign_key: 'IN_SIZE_ID', dependent: :nullify` — confirmed by the
    inbound `IN_SIZE_ID` column + the `DELETE_SizeCode` trigger (mirrors P5). This is the
    association the trigger actually maintains.
  - Validations: `size_code` **presence** + **uniqueness scoped to `site_id`** (`uniqueness:
    {scope: :site_id, case_sensitive: false}` to match the `SQL_Latin1_General_CP1_CI_AS`
    collation), backed by a **per-site unique index `(site_id, VC_SIZE_CODE)`** (D1 — replaces the
    global `IX_INV_SIZE_MST`); `length: {maximum: 6}`. `size_name` presence + `length: {maximum:
    50}`. (Presence rules are a deliberate improvement over the legacy form, which validated
    nothing.)
  - `usage` / `safety_days` integers — decide NULL vs default-0 policy (§8); add numericality
    bounds only if the domain actually caps them (legacy UI capped at 99999, but the DB and the
    ForecastBreakdown writer do not).
  - **No enums** (P4 N/A), **no FK belongs_to** (Size references nothing).
  - Timestamps: write `vc_add` on create, **`vc_last_update`** (underscore!) on update — keep the
    `yyyymmddHHMMSSff` string format during parallel run (P2); normalize at the Postgres phase.
- **Controller/routes:** RESTful `resources :tire_sizes` (or `:sizes`).
- **Views:** index (server-side searchable/paginated list, replacing the client-side `Filter`,
  P7) + new/edit form. **No combos, no directory picker** (nothing to replace — Size has no path
  column, so the §8 directory question that dogged Supplier/Logistics does **not** arise here).
- **Services:** none needed; pure CRUD. **Stage-1 option:** wrap the four existing procs
  (`SELECT/INSERT/UPDATE/DELETE_SizeInfo`) via `tiny_tds` for guaranteed parity — but **do not**
  reproduce the `SELECT_AssyRatioInfo` dup-check bug; rely on the DB unique index for parity and
  add the real validation in stage 3. **Cross-module note:** the ForecastBreakdown module writes
  `IN_USAGE` via `UPDATE_SizeUsage` (by code); the rebuilt size model must expose that field for
  that writer (or both must move to the surrogate id).
- **Reports:** none specific to this module. (`SELECT_SizeUsage` is a forecast-breakdown read, not
  a Size report.)

## 7. Migration plan for this module
- [ ] **Stage 1 — read-only parity:** `TireSize.all` (or wrap `SELECT_SizeInfo ''`) renders the
      list, ordered by `VC_SIZE_CODE`; columns map 1:1 to grid `Fields[0..4]`. Server-side search
      replaces the in-memory `Filter` (P7). No writes.
- [ ] **Stage 2 — writes via wrapped procs:** call `INSERT/UPDATE/DELETE_SizeInfo` through
      `tiny_tds`, preserving the `DELETE_SizeCode` trigger behavior (parts `IN_SIZE_ID` nulling,
      P5). Confirm `VC_ADD`/`VC_LAST_UPDATE` strings still written. **Replace the broken app-side
      dup-check** with either the DB unique-index error surfaced cleanly, or (preferred) a real
      `SELECT_SizeInfo @SizeCode` pre-check — i.e. fix the `SELECT_AssyRatioInfo` bug rather than
      port it.
- [ ] **Stage 3 — reimplement (Postgres-ready):** ActiveRecord validations (presence + uniqueness
      on code, replacing the defective dup-check) leaning on the unique index;
      `has_many :parts_stocks, dependent: :nullify` replaces `DELETE_SizeCode`; real timestamps
      replace the string audit columns; reconcile `IN_USAGE`/`IN_DAYS` NULL-vs-0; decide the
      `_HIST` table handling (§8). **Multi-site (D1):** the Postgres/DB-modernization phase adds the
      `site_id` (NOT NULL) FK to the new `sites` table and the **per-site unique index
      `(site_id, VC_SIZE_CODE)`** (replacing the global `IX_INV_SIZE_MST`); `belongs_to :site` with
      current-site scoping. The legacy single-site DB is left untouched during the parallel run.

## 8. Open questions for the user (domain expert)
1. ✅ **RESOLVED (D1): Multi-site scope of sizes — per-site.** Per decision D1
   (`docs/analysis/decisions.md`), sites run fully independently with no shared data: the
   tire/wheel size catalog is **per-site**, not shared. `INV_SIZE_MST` gains a `site_id` (NOT NULL)
   FK to the new `sites` table, every query is scoped to the current site, and `VC_SIZE_CODE`
   uniqueness becomes **composite (site_id, VC_SIZE_CODE)** rather than global. (Same answer applies
   consistently to Supplier §8.1 and Logistics §8.1.)
2. ✅ **RESOLVED (D8, Bug 1): confirmed bug — fix it.** Verified against source: `InsertSizeInfo`
   (`DataModule.pas:2531`) runs its dup-check via `SELECT_AssyRatioInfo` (`:2543`), which filters
   `INV_ASSY_RATIO_MST.VC_BROADCAST_CODE = @SizeCode` (broadcast codes, **not** sizes), then inserts
   only `If RecordCount = 0`. So a real duplicate size code is never caught app-side, and a code that
   coincides with a broadcast code is falsely rejected. Per decision D8 (docs/analysis/decisions.md),
   the rebuild checks duplicates against `INV_SIZE_MST.VC_SIZE_CODE` and enforces the DB unique index
   `(site_id, VC_SIZE_CODE)` (per D1/D2).
3. **`IN_USAGE` / `IN_DAYS` semantics & bounds:** Daily Usage and Safety Days are nullable `int`,
   but this form always writes `0` (never NULL) and caps input at 99999, while the
   ForecastBreakdown screen writes `IN_USAGE` via `UPDATE_SizeUsage` with no such cap. (a) Should
   blank mean 0 or NULL? (b) Is the 99999 UI cap meaningful, or just an artifact of a 5-digit mask?
   (c) Is it expected that **two different screens** edit a size's daily usage (here by the master
   form, there by forecast breakdown, keyed by code)?
4. ✅ **RESOLVED (D3): block the delete (RESTRICT) — don't unlink.** Per decision D3
   (docs/analysis/decisions.md), deleting a size that is still referenced by any part (current
   `INV_PARTS_STOCK_MST` **or** `_HIST`) is **blocked**. The rebuild does **not** replicate the
   legacy `DELETE_SizeCode` unlink (which nulled `INV_PARTS_STOCK_MST.IN_SIZE_ID` but left the
   history FK dangling) — no nullifying, no dangling. To remove an in-use size, use the future
   **archival** capability (soft-delete / hide), not delete.
5. ✅ **RESOLVED (D2): standardize all callers on the surrogate `IN_SIZE_ID`; `VC_SIZE_CODE` is an
   editable, non-key attribute.** Per decision D2 (docs/analysis/decisions.md), the surrogate id is
   the sole key. Parts are already safe (linked by id); the callers that legacy-resolved **by code**
   (`UPDATE_SizeUsage` / `SELECT_SizeUsage` and the size form search) must be **reworked to resolve
   by `IN_SIZE_ID`**, treating the code as a display label. Renaming a size code is then **allowed**
   (extremely rare) and **safe with no cascade**. The code stays unique **per-site** (composite
   `(site_id, VC_SIZE_CODE)`, per D1) as an attribute constraint, not a key.

## 9. Test cases / parity checks
- **List all** → row count and ordering match `SELECT_SizeInfo ''` (sorted by `VC_SIZE_CODE`);
  the 5 returned columns map 1:1 to grid `Fields[0..4]`.
- **Insert a new code** → a new row with `VC_ADD` populated as a `yyyymmddHHMMSSff` (16-char) string,
  `IN_SIZE_ID` identity-assigned (legacy does **not** echo the new id back; verify the new app
  persists/returns it). `IN_USAGE`/`IN_DAYS` arrive as the entered integers (0 if left blank).
- **Insert an existing code** → **legacy behavior is the buggy path:** the app dup-check
  (against `SELECT_AssyRatioInfo`) does **not** catch it; the insert is attempted and the **DB
  unique index `IX_INV_SIZE_MST` rejects it** with a raw "Unable to insert Size data" error
  (no row added). **New app:** reject cleanly via uniqueness validation with a friendly message —
  document this as an **intentional divergence** (fixing the legacy bug), not a parity match.
- **Insert a code that equals an existing broadcast code** (legacy-only quirk) → legacy silently
  suppresses the insert and logs "(DUPLICATE)" even though no size duplicate exists. New app must
  **not** reproduce this — assert the size inserts normally.
- **Insert a blank code** → legacy *allows* one blank `VC_SIZE_CODE` (no form guard; unique index
  permits a single `''`). New app: reject via presence validation (document the divergence).
- **Update** → row updated by `IN_SIZE_ID`; `VC_LAST_UPDATE` set (note underscore), `VC_ADD`
  unchanged; all four data fields rewritten including the code.
- **Rename a code onto an existing code** → DB `IX_INV_SIZE_MST` rejects it (legacy `UPDATE` has no
  app re-check → raw error); confirm the new app surfaces a clean validation error.
- **Delete a size referenced by parts** → size row gone; every `INV_PARTS_STOCK_MST` that pointed
  at it now has `IN_SIZE_ID = NULL`, the part rows preserved (trigger parity, P5). Confirm the
  legacy does **not** touch `INV_PARTS_STOCK_MST_HIST` and assert the new app's chosen behavior
  (§8.4).
- **Search** → typing an exact (uppercased) code filters the grid to that one row; a partial or
  non-matching string finds nothing ("No matches were found for your query"). New app server-side
  search should be at least as capable (document any added partial matching).
