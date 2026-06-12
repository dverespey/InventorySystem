# Module Analysis: Logistics Master

**Area:** Master data  **Status:** ✅ spec complete  **Analyst:** Claude / 2026-06-04

> Sibling of the Supplier Master spec — same master-detail CRUD shape, but keyed by
> **name** (no 5-char business code) rather than a code. "Logistics" here = the
> **carriers / logistics providers** suppliers ship through; a supplier points at one
> logistics row via `IN_LOGISTICS_ID`.

## 1. Legacy surface
- **Form:** `LogisticsMaster.pas` + `LogisticsMaster.dfm` (`TLogisticsMaster_Form`,
  Caption "Logistics", header label "Logistics Master"). Registered live in
  `InventorySystem.dpr` line 29.
- **Entry point:** Master-maintenance menu in `MainMenu.pas` → `LogisticsMaster_Form.Execute`
  (`Execute()` returns False on `mrCancel`, i.e. Close).
- **Purpose (one paragraph):** Classic master-detail CRUD screen. A read-only-by-convention
  `DBGrid` (`LogisticsMaster_DBGrid`) lists all logistics records; an edit panel shows the
  selected row's details. Buttons: **Insert, Update, Delete, Search, Clear, Close**, plus a
  glyph-only **Breakdown speedbutton** that runs a Windows directory picker. `FormCreate`
  clears the dataset filter, calls `GetLogisticsInfo`, binds the `DataSource` to `Inv_DataSet`,
  and clears the panel. Selecting a grid row (`OnKeyUp` / `OnMouseUp` / `DataSource.OnDataChange`)
  calls `HoldDetails(True)` + `SetDetailBoxes`, copying the row into the edits and capturing
  `RecordID` from hidden grid `Fields[10]` (the identity PK).

## 2. Data touched
| Table | Read | Write | Notes |
|-------|:----:|:-----:|-------|
| `INV_LOGISTICS_MST` | ✓ | ✓ | The logistics/carrier master (this module owns it) |
| `INV_SUPPLIER_MST` |  | ✓* | *Indirect: DELETE trigger nulls `IN_LOGISTICS_ID` (FK unlink) |
| `INV_PARTS_STOCK_MST` | ✓† |  | †Not by this form; other procs read a part's logistics directory (see §3) |

Unlike Supplier, this form pulls in **no FK-lookup combos** (Supplier reads
`INV_LOGISTICS_MST`, `INV_PART_TYPE_MST`, `INV_ADD_POINT_INF` to populate selects). Logistics
is a leaf master: it is *referenced by* others but references nothing itself.

### `INV_LOGISTICS_MST` columns
| Column | Type | Meaning / notes |
|--------|------|-----------------|
| `IN_LOGISTICS_ID` | int IDENTITY PK | Surrogate key (`RecordID` in UI, hidden grid `Fields[10]`) |
| `VC_LOGISTICS_NAME` | varchar(25) NULL | **Business key — the identity of the row; UNIQUE via `IX_INV_LOGISTICS_MST`** (no 5-char code like Supplier). Per **D1** this uniqueness becomes **per-site composite `(site_id, VC_LOGISTICS_NAME)`**, not global |
| `VC_ADDRESS` | varchar(50) | |
| `VC_CITY` | varchar(50) | Form caps input at 10 ⚠️ (< DB 50) |
| `VC_STATE` | varchar(50) | Form caps input at 10 ⚠️ (< DB 50) |
| `VC_ZIP` | varchar(10) | |
| `VC_COUNTRY` | varchar(50) | **Exists on table but NOT exposed on the form** (same as Supplier's unused COUNTRY) |
| `VC_TEL` | varchar(10) | `fLogTel` → param `@LogPhone` |
| `VC_FAX` | varchar(10) | |
| `VC_PERSON` | varchar(50) | Contact. **Write path narrows to 25** — proc param `@LogPerson varchar(25)` (form caps at 50, DB is 50) ⚠️ |
| `VC_BREAKDOWN_ORDER_DIRECTORY` | varchar(512) | **Local Windows filesystem path** for breakdown-order files ⚠️ desktop-bound. Effective max is **215**, not 512: proc param `@LogDirectory varchar(215)`; the form caps input at 50 ⚠️ (form 50 < proc 215 < DB 512) |
| `VC_EMAIL_ADDRESS` | varchar(255) | |
| `VC_LASTUPDATE` | varchar(16) | **Timestamp as `yyyymmddHHMMSSff` string** (set on UPDATE only); not exposed on form (P2) |
| `VC_ADD` | varchar(16) | **Timestamp as `yyyymmddHHMMSSff` string** (set on INSERT only); not exposed on form (P2) |

**Difference vs Supplier:** the `VC_BREAKDOWN_ORDER_DIRECTORY`, `VC_ADD`, and `VC_LASTUPDATE`
columns are **shared** with `INV_SUPPLIER_MST` (not a Logistics-only trait — verified present on
both `CREATE TABLE`s). The genuine difference is what Logistics **lacks**: no 5-char
`VC_SUPPLIER_CODE` business key, and **none** of the order-file / enum columns `VC_OUTPUT_FILE`,
`BIT_ORDER_FILE_TIMESTAMP`, `BIT_SITE_NUMBER_IN_ORDER`, `VC_CREATE_ORDER_SHEET`, or
`VC_INVENTORY_ADD_POINT`. So **no P4 (coded single-char enums)** apply here, and there are **no
FK-lookup combos**.

**Constraints / indexes:**
- `PK_INV_LOGISTICS_MST` PRIMARY KEY CLUSTERED (`IN_LOGISTICS_ID`).
- `IX_INV_LOGISTICS_MST` **UNIQUE NONCLUSTERED (`VC_LOGISTICS_NAME`)** — a real DB unique
  backstop on the name. (Supplier's 5-char code is **also** DB-unique, via `IX_INV_SUPPLIER_MST` —
  same posture; an earlier draft wrongly called Supplier's uniqueness app-side-only.) No DEFAULT
  constraints. **No declared FK constraints out of this table.**
- Inbound references (by convention, **no declared FK**): `INV_SUPPLIER_MST.IN_LOGISTICS_ID`,
  `INV_PARTS_STOCK_MST.IN_LOGISTICS_ID`, `INV_PARTS_STOCK_MST_HIST.IN_LOGISTICS_ID`.
- **Multi-site (D1):** the live table has **no site column** today. The Postgres-phase rebuild adds a
  `site_id` (NOT NULL) FK → `sites`, rows become per-site, and `IX_INV_LOGISTICS_MST` is replaced by
  a per-site composite unique index `(site_id, VC_LOGISTICS_NAME)`. See §6/§7 and decision D1.

**Triggers on these tables:**
- `DELETE_LogisticsCode` (on `INV_LOGISTICS_MST` FOR DELETE, authoritative version at
  `DB Schema/Create Inventory.sql` line 9636): sets `INV_SUPPLIER_MST.IN_LOGISTICS_ID = NULL`
  for every supplier pointing at the deleted logistics row.
  **Invariant: deleting a logistics record unlinks (does NOT delete) its suppliers.** This is
  the exact analogue of `DELETE_SupplierCode` (P5). Note it only severs the **supplier** FK —
  it does **not** touch `INV_PARTS_STOCK_MST.IN_LOGISTICS_ID`, which is left dangling.
- ⚠️ **Two source files disagree.** `docs/triggers.sql` carries *stale* variants:
  a `DELETE_LogisticsCode` (line 145) and an `UPDATE_LogisticsCode` (line 164) that both key on
  `INV_SUPPLIER_MST.VC_LOGISTICS_CODE` — a string column **that no longer exists** on the live
  supplier table (it links via `IN_LOGISTICS_ID int`; `VC_LOGISTICS_CODE` survives only on the
  legacy `[Results]` table, varchar(5)). These predate the int-FK refactor; **treat them as
  obsolete.** The live schema has exactly one trigger here (the `IN_LOGISTICS_ID`-nulling DELETE)
  and **no UPDATE trigger** — correct, since the identity PK is immutable so there is nothing to
  propagate.

## 3. Stored procedures used
| Proc | Op | Business rule (from body) |
|------|----|---------------------------|
| `SELECT_LogisticsInfo;1 @LogisticsName` | SELECT | Single param. If `@LogisticsName = ''` returns **all** rows; else rows `WHERE VC_LOGISTICS_NAME = @LogisticsName` (filters **by name, not id**). Both branches select the same 11 UI-aliased columns (`Logistics Name, Address, City, State, Zip, Telephone, Fax, Person, Email, Directory, RecordID`) `ORDER BY VC_LOGISTICS_NAME`. No joins; **no name→id resolution** (the surrogate is only *returned*, never used for filtering). `VC_COUNTRY/VC_LASTUPDATE/VC_ADD` are **not** selected. |
| `INSERT_LogisticsInfo;1` (10 params) | INSERT | Computes `VC_ADD` as a `yyyymmddHHMMSSff` string (P2); inserts the 10 data columns + `VC_ADD`. `IN_LOGISTICS_ID` is identity (not supplied, **not returned**). **No uniqueness check inside the proc** — the dup guard is in the app (P1). **Param widths narrower than the table:** `@LogPhone`/`@LogFax varchar(10)`, `@LogPerson varchar(25)` (< table 50), `@LogDirectory varchar(215)` (< table 512) — writes silently truncate to the proc widths. |
| `UPDATE_LogisticsInfo;1` (+`@LogisticsID int`) | UPDATE | Updates one row keyed `WHERE IN_LOGISTICS_ID = @LogisticsID` (surrogate passed directly; **no name→id resolution**). Sets `VC_LASTUPDATE` (P2); rewrites all 10 data columns including `VC_LOGISTICS_NAME` (so the business key **is editable**). Does **not** touch `VC_ADD`. Same narrowed param widths as INSERT (`@LogPerson varchar(25)`, `@LogDirectory varchar(215)`). No uniqueness re-check — does not prevent renaming onto an existing name (the DB `IX_INV_LOGISTICS_MST` would, however, reject it). |
| `DELETE_LogisticsInfo;1 @LogisticsID` | DELETE | Hard-deletes `WHERE IN_LOGISTICS_ID = @LogisticsID`. Single surrogate param. No soft-delete flag, no cascade, **no in-use / RI check inside the proc** — relies entirely on the `DELETE_LogisticsCode` trigger to null supplier FKs; parts FKs are not handled at all. |
| `REPORT_MonthlyLogisticsOrders @StartDate,@EndDate,@Logistics='ALL'` | SELECT (report) | Joins `inv_open_order_inf` × `inv_supplier_mst` × `inv_logistics_mst` (implicit comma joins) on supplier code and `IN_LOGISTICS_ID`. Filters `VC_STATUS_SUPPLIER_SHIPPING BETWEEN @StartDate AND @EndDate`. If `@Logistics<>'ALL'` filters `L.vc_logistics_name=@Logistics` (**resolves by name**). Groups/orders by logistics name + renban. (Consumed by `MonthlyLogiticsOrderReport.pas`; listed for completeness.) |
| `SELECT_PartsStockLogistics @PartNo` | SELECT | Returns one part's `VC_BREAKDOWN_ORDER_DIRECTORY` (aliased `LogisticsDirectory`) via `INV_PARTS_STOCK_MST p JOIN INV_LOGISTICS_MST l ON p.IN_LOGISTICS_ID = l.IN_LOGISTICS_ID WHERE p.VC_Part_Number = @PartNo`. Shows the directory is consumed *per part* elsewhere — relevant to the §8 directory question. |

### Call mechanism (legacy)
`DataModule.pas` methods `GetLogisticsInfo / InsertLogisticsInfo / UpdateLogisticsInfo /
DeleteLogisticsInfo` (declarations 521–524; bodies 802–1039) drive a single shared ADO object
(P6). Notable per-method facts:
- **`GetLogisticsInfo`** uses **`Inv_DataSet`** (open result set) with `@LogCode := ''`, times
  the call (`fBeforeDateTime/…/fDiffDateTime`), and on success logs `LogActLog('GET LOG',…)`.
  The other three use **`Inv_StoredProc`** (`ExecProc`).
- **`InsertLogisticsInfo`** (Boolean) is **two-step (P1)**: STEP 1 calls
  `SELECT_LogisticsInfo;1` but passes **`@LogisticsName := fLogName`** (note: *different param
  name* from `GetLogisticsInfo`'s `@LogCode` against the **same proc** — the proc reads one
  param either way). `If RecordCount = 0` → INSERT and `Result := True`; else it is a **duplicate**:
  sets `fDescription := 'FAILED to … (DUPLICATE)'` and shows `'Unable to insert duplicate
  logistics name(<name>)'`. **The INSERT does not capture the new identity into `fRecordID`.**
- **`UpdateLogisticsInfo`** sets the same 10 fields + `@LogisticsID := fRecordID`. It is the
  **only one of the four with no success `LogActLog` entry**.
- **`DeleteLogisticsInfo`** passes only `@LogisticsID := fRecordID`; success logs
  `'DELETED ' + fLogName`.
- All four share the **retry-up-to-3-times-via-recursion** error pattern (`fErrorCount < 3` →
  recursive self-call) with a `finally` doing `Inv_StoredProc.Close; fErrorCount := 0` (P8). On a
  hard error each raises a distinct `EDatabaseError` ("Unable to get/insert/update/delete
  logistics data") after `ShowMessage` + `LogActLog('ERROR',…)`.

**DataModule properties** (lines 363–372): `LogisticsName / LogisticsAddress / LogisticsCity /
LogisticsState / Logisticszip / LogisticsTelephone / LogisticsFax / LogisticsPerson /
LogisticsEmail / LogisticsDirectory`. The record key is the **shared, generic** `RecordID`
property (line 337) — *not* Logistics-specific; it is reused by Shipping/Invoice/ASN modules,
so a stale `RecordID` from another screen is a real (latent) cross-module hazard (P9).

## 4. Business rules & edge cases
- **Identity is the NAME.** A logistics record is identified by `VC_LOGISTICS_NAME`. **There is
  no 5-char code** (the headline difference from Supplier). The name is `MaxLength=25`, forced
  uppercase (`CharCase=ecUpperCase`).
- **No form-level validation at all.** Lengths are enforced *only* by `TEdit.MaxLength`. There is
  **no min-length, not-blank, or required-field check** — **inserting a blank name is not blocked
  in Pascal** (the DB unique index would still allow a single blank/NULL name). Contrast Supplier,
  which rejected `length < 5`.
- **Uniqueness (P1, but stronger here):** app-side dup check (`InsertLogisticsInfo` STEP 1, by
  name) **plus** a real DB `UNIQUE` index `IX_INV_LOGISTICS_MST`. Update/Delete do **not** re-check
  uniqueness in the app, but the DB index backstops a rename collision.
- **Name is editable** — `UPDATE_LogisticsInfo` rewrites `VC_LOGISTICS_NAME`. Since suppliers link
  by `IN_LOGISTICS_ID` (not by name) a rename does **not** break supplier links. But
  name-based callers (`REPORT_MonthlyLogisticsOrders @Logistics`, supplier-save procs that resolve
  `IN_LOGISTICS_ID` from `VC_LOGISTICS_NAME`, the form's client-side `Search`) **would** be
  affected by a rename — a name-as-key fragility worth flagging (§8).
- **Timestamps are `yyyymmddHHMMSSff` strings (P2):** `VC_ADD` on insert, `VC_LASTUPDATE` on
  update; computed in-proc from `getdate()`. Update preserves the original `VC_ADD`.
- **Delete is soft on suppliers** (trigger nulls `INV_SUPPLIER_MST.IN_LOGISTICS_ID`, P5), **hard**
  on the logistics row. **Gap vs Supplier:** the trigger does **not** unlink
  `INV_PARTS_STOCK_MST.IN_LOGISTICS_ID` (or the `_HIST` table) — those references can be **left
  dangling** after a delete. This asymmetry must be decided in the rebuild (§8).
- **No coded enums (no P4), no FK combos.** Logistics references nothing; the form has no selects.
- **Stale-`RecordID` hazard (P9):** Update/Delete key off `DataModule.RecordID` (grid `Fields[10]`),
  populated only by `HoldDetails(True)` when a row is selected. If no row was selected first,
  `RecordID` may be 0 or carry a leftover value from another module — no guard exists in the form.

## 5. UI / UX notes
- Grid + detail-panel pattern; selecting a row syncs the panel and captures `RecordID`.
- **Search is fully client-side (P7):** `SearchGrid` loops `Inv_DataSet` `First`→`EOF` in memory,
  comparing the typed name to `Trim(FieldByName('LOGISTICS NAME').AsString)`. Match is **exact,
  case/space-sensitive after Trim** (input is uppercased), **no partial/LIKE** — it does **not**
  re-query the DB. On no match: "No matches were found for your query."
- **Fields (label → control → DB):** Logistics Name (25), Address (50), City (10⚠️), State
  (10⚠️), Zip (10), Telephone Number (10, `@LogPhone`), Fax Number (10), Person (50), Email (255),
  **Breakdown Directory** (form 50⚠️) with the folder-picker speedbutton. All edits are
  `CharCase=ecUpperCase`. **No format validation** on Email/Zip/Phone/Fax beyond uppercase +
  MaxLength. (Effective write widths are bounded by the proc params, §3: Person→25, Directory→215.)
- **Breakdown speedbutton** opens `SelectDirectory('Select A Directory','My Computer', dir)`
  (`FileCtrl`), seeded from `Directory_Edit.Text` if `DirectoryExists` else `'c:\'`, and writes the
  chosen path back into the edit. It **only sets a local folder path** — it does **not** launch the
  separate `LogisticsBreakdown.pas` processing form (that is an unrelated inbound-status file
  processor, dpr line 35).
- **Modernize:** standard index/list + new/edit form; server-side search/sort/pagination (P7);
  inline validation (presence + uniqueness on name); **drop the directory picker** (replace per §8);
  widen the truncated fields to DB widths (City/State/Directory). `VC_COUNTRY` is on the table but
  unused by the form — surface it or leave it (§8).

## 6. Target design  *(Rails primary)*
- **Model:** `Logistics` (singular ActiveRecord class; `self.table_name = 'INV_LOGISTICS_MST'`,
  `self.primary_key = 'IN_LOGISTICS_ID'`).
  - `has_many :suppliers, foreign_key: 'IN_LOGISTICS_ID', dependent: :nullify` — confirmed by the
    inbound FK + `DELETE_LogisticsCode` trigger (mirrors P5). The `suppliers` association is the
    one the trigger actually maintains.
  - `has_many :parts_stocks, foreign_key: 'IN_LOGISTICS_ID'` — **deliberately NOT `dependent:`
    by default**, because the legacy trigger leaves part FKs dangling. Make this an explicit §8
    decision (`:nullify` to fix the legacy gap, or `restrict_with_error` to block deleting an
    in-use logistics row); whatever is chosen, document the divergence from legacy.
  - **Multi-site (D1):** `belongs_to :site` with enforced current-site scoping (every query filtered
    to the current site); the carrier/logistics list is **per-site, not shared**. The unique index
    becomes per-site composite **`(site_id, VC_LOGISTICS_NAME)`**.
  - Validations: `logistics_name` **presence** (the legacy form lacked this — a deliberate
    improvement) + `uniqueness` **scoped to `site_id`** (case-insensitive to match the SQL collation)
    backed by the per-site composite `(site_id, VC_LOGISTICS_NAME)` unique index (replacing the global
    `IX_INV_LOGISTICS_MST`). No `length: {is: …}` rule (no fixed code).
  - **No enums** (P4 not applicable). Map column readers/writers to the friendly names
    (`logistics_name`, `breakdown_order_directory`, `tel`, `country`, etc.).
  - Timestamps: `vc_add` on create, `vc_lastupdate` on update — **keep the `yyyymmddHHMMSSff`
    string format during parallel run** (P2), normalize at the Postgres phase.
- **Controller/routes:** RESTful `resources :logistics` (override the default plural inflection
  so the path/route helpers read sensibly, e.g. `resources :logistics, controller: 'logistics'`).
- **Views:** index (server-side searchable/paginated list, replacing client-side `SearchGrid`,
  P7) + new/edit form. **No combos** (this master references nothing). Replace the directory
  picker with the §8 outcome (configured-root + relative subpath, or remove).
- **Services:** none needed; pure CRUD. **Stage-1 option:** wrap the four existing procs
  (`SELECT/INSERT/UPDATE/DELETE_LogisticsInfo`) via `tiny_tds` for guaranteed parity — including
  the app-side dup check (P1) — then switch to ActiveRecord in stage 3.
- **Reports:** `REPORT_MonthlyLogisticsOrders` (monthly logistics order summary,
  `MonthlyLogiticsOrderReport.pas`) — wrap as-is initially; it filters by logistics **name**, so it
  is sensitive to renames. `SELECT_PartsStockLogistics` is a per-part directory lookup consumed by
  other modules, not a logistics report per se.

## 7. Migration plan for this module
- [ ] **Stage 1 — read-only parity:** `Logistics.all` (or wrap `SELECT_LogisticsInfo ''`) renders
      the list, ordered by `VC_LOGISTICS_NAME`. Server-side search replaces the in-memory
      `SearchGrid` (P7). No writes yet.
- [ ] **Stage 2 — writes via wrapped procs:** call `INSERT/UPDATE/DELETE_LogisticsInfo` through
      `tiny_tds`, preserving the app-side dup check (P1) and the `DELETE_LogisticsCode` trigger
      behavior (supplier FK nulling, P5). Confirm `VC_ADD`/`VC_LASTUPDATE` strings still written.
- [ ] **Stage 3 — reimplement (Postgres-ready):** ActiveRecord validations replace the app-side
      dup check (presence + uniqueness on name, now **scoped to `site_id`**); `has_many :suppliers,
      dependent: :nullify` replaces `DELETE_LogisticsCode`; **resolve the
      parts-FK gap** explicitly (nullify or restrict — §8); real timestamps replace the string
      audit columns; widen the form-truncated fields to DB widths.
- [ ] **Multi-site (D1):** the Postgres phase adds the `site_id` (NOT NULL) FK → `sites` and replaces
      `IX_INV_LOGISTICS_MST` with the per-site composite unique index `(site_id, VC_LOGISTICS_NAME)`;
      the model gets `belongs_to :site` with current-site scoping. The legacy single-site SQL Server
      DB is untouched during the parallel run (the new app filters to its one site). See decision D1.

## 8. Open questions for the user (domain expert)
1. ✅ **RESOLVED (D1): per-site** — multi-site scope of logistics/carriers. `INV_LOGISTICS_MST`
   today has **no site/plant column** (single global table, every query returns all rows
   unfiltered). Per **decision D1 (docs/analysis/decisions.md)**, sites run independently with full
   data isolation: the table gains a `site_id` (NOT NULL) FK to the new `sites` table, the
   carrier/logistics list is **per-site (not shared)**, and the unique-name constraint becomes
   composite per-site — `(site_id, VC_LOGISTICS_NAME)`. (Resolved consistently with Supplier §8.1.)
2. **`VC_BREAKDOWN_ORDER_DIRECTORY` replacement:** it is a **local Windows path** chosen via
   `SelectDirectory`, defaulting to `c:\`, used to write breakdown-order files and **read per-part**
   (`SELECT_PartsStockLogistics`). A server-side absolute Windows path is meaningless per browser
   client; and per **D1** site-level paths now live in the `sites` table, not the `[DIRECTORIES]`
   INI, so any output root is configured **per-site** there. What should replace this per-row
   directory — a per-site configured output root (network share / SFTP / object store) with a
   stored relative subpath, an upload target, or is the file-output workflow going away? Also note the form caps this at **50 chars vs proc 215 vs DB 512** —
   real paths almost certainly need the full width (or a different model entirely).
3. ✅ **RESOLVED (D2): standardize on the surrogate `IN_LOGISTICS_ID`; the name is an editable,
   non-key attribute.** Per decision D2 (docs/analysis/decisions.md), the surrogate id is the sole
   key — every FK/join/lookup resolves on `IN_LOGISTICS_ID`, and `VC_LOGISTICS_NAME` becomes a
   display-only, editable label. The name-based lookups that exist today (**supplier-save procs, the
   monthly report's `@Logistics` filter, the form's search**) must be **reworked to resolve by id**.
   Renaming a logistics record is then **allowed** (rare) and **safe with no cascade**. The name
   stays unique **per-site** (composite `(site_id, VC_LOGISTICS_NAME)`, per D1) as an attribute
   constraint, not a key.
4. ✅ **RESOLVED (D3): block the delete (RESTRICT).** Per decision D3 (docs/analysis/decisions.md),
   the rebuild **blocks** deleting a logistics row that is still referenced by any supplier or part —
   it does **not** dangle (today's behavior) and does **not** nullify part links. The legacy
   `DELETE_LogisticsCode` trigger's inconsistency (nulls supplier FKs but leaves
   `INV_PARTS_STOCK_MST.IN_LOGISTICS_ID`/`_HIST` dangling) goes away. To remove an in-use logistics
   record, use the future **archival** capability (soft-delete / hide from view), not delete.
5. **`VC_COUNTRY`** is on the table but never exposed on the form (same as Supplier). Intentional?
   Should the rebuilt form add it, or drop the column?

## 9. Test cases / parity checks
- **List all** → row count and ordering match `SELECT_LogisticsInfo ''` (sorted by
  `VC_LOGISTICS_NAME`); returned columns map 1:1 to grid `Fields[0..10]`.
- **Insert with an existing name** → rejected, no row added; legacy shows "Unable to insert
  duplicate logistics name(<name>)" (app-side P1 dup check). New app: validation error.
- **Insert with a blank name** → legacy *allows* it (no form guard; DB unique index permits one).
  Decide and assert the new-app behavior (recommended: reject via presence validation — document the
  intentional divergence).
- **Insert** → a new row with `VC_ADD` populated as a `yyyymmddHHMMSSff` string and `IN_LOGISTICS_ID`
  identity-assigned (legacy does **not** echo the new id back to the client — verify the new app
  returns/persists it correctly).
- **Update** → row updated by `IN_LOGISTICS_ID`; `VC_LASTUPDATE` set, `VC_ADD` unchanged; all 10
  data fields rewritten including the name.
- **Rename onto an existing name** → DB `IX_INV_LOGISTICS_MST` rejects it (legacy `UPDATE` has no
  app re-check; confirm the new app surfaces a clean validation error rather than a raw DB error).
- **Delete a logistics row referenced by suppliers** → logistics row gone; every
  `INV_SUPPLIER_MST` that pointed at it now has `IN_LOGISTICS_ID = NULL`, supplier rows preserved
  (trigger parity, P5).
- **Delete a logistics row referenced by parts** → confirm the **legacy** leaves
  `INV_PARTS_STOCK_MST.IN_LOGISTICS_ID` **dangling** (no trigger action); assert the new app's chosen
  behavior (nullify or block) per §8.4.
- **Search** → typing an exact (uppercased, trimmed) name selects that row; a partial/lowercase
  string finds nothing in legacy (client-side exact match, P7). New app server-side search should be
  defined to be at least as capable (and documented if it intentionally adds partial matching).
