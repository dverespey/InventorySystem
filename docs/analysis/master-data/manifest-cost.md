# Module Analysis: Manifest Cost Master

**Area:** Master data  **Status:** ✅ spec complete  **Analyst:** Claude / 2026-06-04

> Sibling of the Supplier / Logistics master specs — same master-detail CRUD shape, but
> this master is **financially load-bearing**: `INV_MANIFEST_COST_MST.MO_PRICE` is the
> **unit price the EDI 810 invoice and every invoice/PO report bills the customer at**,
> keyed by **assembly part-number code** (`VC_ASSY_PART_NUMBER_CODE`). "Manifest cost"
> = the agreed per-unit price of an assembly over a date window (Toyota "manifest" =
> the priced shipment manifest). Despite the form/proc/DataModule naming, this is **not**
> the legacy "Monthly PO" master — it is its EDI-era replacement (see §1 entry point).

## 1. Legacy surface
- **Form:** `ManifestCostMaster.pas` (9.2 KB) + `ManifestCostMaster.dfm` (6.9 KB).
  Type `TManifestCostMaster_Form`, Caption "Manifest Cost Master", header label
  "Manifest Cost Master". Registered live in `InventorySystem.dpr` line 49.
- **Entry point:** **Not reached from `MainMenu.pas`** (no reference there at all — a
  difference from Supplier/Logistics). It is launched from the **Master-Maintenance hub**
  `MasterMaint.pas` (dpr line 13), button **"&Monthly PO"** (`MonthlyPO_Button`,
  `MasterMaint.dfm` line 104). The handler `MonthlyPO_ButtonClick` is **feature-flag
  gated**:
  ```pascal
  if Data_Module.fiGenerateEDI.AsBoolean then   // EDI generation ON
    ManifestCostMaster_Form ...Execute          // → THIS module
  else
    MonthlyPOMaster_Form ...Execute             // → legacy MonthlyPOMaster (the older variant)
  ```
  So the same button opens **one of two different masters** depending on the INI/DB EDI
  flag `fiGenerateEDI`. `Execute` does `ShowModal`; returns False only on `mrCancel`
  (Close button, `ModalResult = 2`).
  - ⚠️ The `"&Monthly PO"` caption is the **design-time** value. On EDI sites
    `MasterMaint.pas` (line 73) **rewrites the caption to `"Manifest Cost"` at runtime** when
    `fiGenerateEDI` is true — so the user actually sees a **"Manifest Cost"** button, not
    "Monthly PO". The proc/DataModule code and audit logs still say `PO`/`MonthlyPO` (legacy
    naming carried over from `MonthlyPOMaster`).
- **Purpose (one paragraph):** Classic master-detail CRUD screen for the assembly price
  table that drives EDI 810 invoicing. A `DBGrid` (`MonthlyPOMaster_DBGrid`, bound to the
  shared `Inv_DataSet`) lists all manifest-cost rows (`ID, Assy, Manifest ID, Start
  Manifest, End Manifest, Cost`); an edit panel shows the selected row. Buttons:
  **Insert, Update, Delete, Search, Clear, Close**. Selecting a grid row (or any
  `DataSource.OnDataChange`) calls `HoldDetails(True)` then `SetDetailBoxes`, copying the
  row into the controls and capturing the identity PK into `Data_Module.RecordID`
  (grid `Fields[0]`, which it then hides — `Fields[0].Visible := FALSE`).

## 2. Data touched
| Table | Read | Write | Notes |
|-------|:----:|:-----:|-------|
| `INV_MANIFEST_COST_MST` | ✓ | ✓ | The manifest-cost / assembly-price master (this module owns it) |
| `INV_FORECAST_DETAIL_INF` | ✓ |  | Source of the **Assy Code combo** — `SELECT DISTINCT(VC_ASSY_PART_NUMBER_CODE)` (raw dynamic SQL, no proc) |

Unlike Supplier, this form pulls in **no FK-lookup-by-id combos**. The Assy Code combo is
a flat list of distinct assembly codes seen in forecast detail; the Manifest-ID combo is
**hard-coded in Pascal** (`' '`, then `'01'`..`'99'` — see `FormCreate`). There is **no
DB FK** from manifest-cost to forecast-detail (or anywhere): `VC_ASSY_PART_NUMBER_CODE` is
a free string that *happens* to be validated by the combo at entry time only.

### `INV_MANIFEST_COST_MST` columns (`DB Schema/Create Inventory.sql` line 1464)
| Column | Type | Meaning / notes |
|--------|------|-----------------|
| `IN_MANIFEST_COST_ID` | int IDENTITY(1,1) NOT NULL | Surrogate key (`RecordID`/grid `Fields[0]`). **No PRIMARY KEY constraint is declared** on the table — identity only ⚠️ (see Constraints) |
| `VC_ASSY_PART_NUMBER_CODE` | varchar(12) NOT NULL | **Business key — the assembly part number** the price applies to. Combo `MaxLength` not set; `CharCase=ecUpperCase`. Proc param `@AssyCode varchar(12)` = DB width (no narrowing) |
| `VC_ASSY_MANIFEST_NUMBER` | varchar(2) NOT NULL | 2-char manifest id `'01'..'99'` (and `' '`). Combo `csDropDownList`, hard-coded list. Proc param `@AssyManifestNo varchar(2)` = DB width |
| `VC_START_MANIFEST` | varchar(8) NOT NULL | **Window start, `yyyymmdd` string** (date stored as text). Written from `CostStart_NUMMIBmDateEdit` via `formatdatetime('yyyymmdd',…)`. Proc param `@StartManifest varchar(8)` = DB width |
| `VC_END_MANIFEST` | varchar(8) NOT NULL | **Window end, `yyyymmdd` string.** Same handling. Proc param `@EndManifest varchar(8)` = DB width |
| `MO_PRICE` | **money** NOT NULL | **The unit price** billed on invoices. Edited via `AssyCost_MaskEdit` (`TcurrEdit`, prefix `$`, **4 decimal digits**, `IntDigits=0`). Proc param `@AssyCost money`. ⚠️ form formats with `FormatFloat('$#######0.0000', …)` and parses with `TryStrToFloat` after stripping the `$` — round-trips through a Delphi `Double`, so sub-cent precision beyond `money`'s 4 dp is lost; `money` itself holds 4 dp |
| `VC_LAST_UPDATE` | varchar(16) NULL | **Timestamp as `yyyymmddHHMMSSff` string** (set on UPDATE only). Not exposed on the form (P2). Note the column name is `VC_LAST_UPDATE` (with underscore) — **differs from Supplier/Logistics `VC_LASTUPDATE`** ⚠️ |
| `VC_ADD` | varchar(16) NULL | **Timestamp as `yyyymmddHHMMSSff` string** (set on INSERT only). Not exposed on form (P2) |

**Difference vs Supplier/Logistics:**
- The audit columns differ in **name**: here `VC_ADD` + `VC_LAST_UPDATE`; Supplier/Logistics use
  `VC_ADD` + `VC_LASTUPDATE`. The string **format is identical**, however: `yyyymmddHHMMSSff`
  (**16 chars** — `CONVERT(char(8),…,112)` + four `SUBSTRING(CONVERT(varchar,…,114),p,2)` slices at
  1/4/7/10 = HH+MM+SS+`ff`), the byte-identical recipe Supplier and Logistics use. Both columns are
  `varchar(16)` (the 16-char string fills them exactly — zero slack). *(An earlier draft wrongly
  called this a 14-char `yyyymmddHHMMSS` form differing from Logistics; it does not differ.)*
- **Both `VC_ADD` and `VC_LAST_UPDATE` are set to the same value on INSERT** (the proc
  passes `@AddDate` to both positional columns), so a never-updated row has
  `VC_ADD = VC_LAST_UPDATE`. (Compare Supplier/Logistics, which leave `VC_LASTUPDATE`
  NULL until the first update.)
- The price column is **`money`**, the only numeric-typed master column we have seen so
  far (Supplier/Logistics were all strings/bits). No coded single-char enums (no P4).
- **No `VC_BREAKDOWN_ORDER_DIRECTORY`** and no local-path column — so the desktop-bound
  directory question (Supplier §8.2 / Logistics §8.2) **does not arise here**.

**Constraints / indexes:**
- `IN_MANIFEST_COST_ID` is `IDENTITY(1,1)` but **there is NO declared `PRIMARY KEY`, NO
  unique index, and NO FK** on this table (grep of the schema for `PK_`/`CONSTRAINT`/
  `INDEX` against `INV_MANIFEST_COST_MST` returns nothing — only the `CREATE TABLE`).
  ⚠️ This is weaker than Logistics (which had `PK_INV_LOGISTICS_MST` + a UNIQUE name
  index) and even weaker than Supplier (which at least had an identity PK).
  Consequence: **nothing in the DB enforces one-price-per-assembly** (or even row
  identity beyond the identity column), and **nothing prevents duplicate
  `VC_ASSY_PART_NUMBER_CODE` rows** with overlapping date windows.
- Inbound references (by convention, **no declared FK**): the price is consumed by
  several procs that JOIN on `VC_ASSY_PART_NUMBER_CODE` (see §3) — invoice/EDI810/forecast.

**Triggers on these tables:** **None.** `list triggers` shows no trigger naming
`MANIFEST`, and no trigger body touches `INV_MANIFEST_COST_MST`. Unlike Supplier
(`DELETE_SupplierCode`) and Logistics (`DELETE_LogisticsCode`), deleting a manifest-cost
row fires **no cascade/unlink** — so no P5 here. (The risk is the reverse: an invoice
line that JOINs to a now-deleted price simply disappears from invoice/report output — see
§4.)

## 3. Stored procedures used
(Grep of `ManifestCostMaster.pas` → `DataModule.pas` `GetManifestCostInfo /
InsertManifestCostInfo / UpdateManifestCostInfo / DeleteManifestCostInfo`, plus the raw
`SelectSingleField` helper for the combo.)

| Proc | Op | Business rule (from body) |
|------|----|---------------------------|
| `SELECT_ManifestCost;1 @AssyCode='' , @ProdDate=''` | SELECT | **Three branches.** (a) `@AssyCode=''` → **all** rows. (b) `@AssyCode` set, `@ProdDate=''` → rows `WHERE vc_assy_part_number_code=@AssyCode`. (c) both set → also `AND (@ProdDate > VC_START_MANIFEST AND @ProdDate < VC_END_MANIFEST)` — **strict `>`/`<`, so a date exactly on the window boundary is EXCLUDED** ⚠️. String date comparison works because dates are `yyyymmdd`. Selects 6 UI-aliased columns: `ID, Assy, Manifest ID, Start Manifest, End Manifest, Cost`. **No `ORDER BY`** (grid order is DB-natural, i.e. effectively by identity) — contrast Supplier/Logistics which order by the business key. The form calls this with **no params** (`Parameters.Clear`, defaults), so the screen always loads branch (a) = all rows. Branches (b)/(c) exist for *other* callers / a date-resolution use-case. |
| `INSERT_ManifestCost;1` (5 params) | INSERT | Computes `@AddDate` as a `yyyymmddHHMMSSff` string (P2); does a **positional** `INSERT INTO INV_MANIFEST_COST_MST VALUES(@AssyCode,@AssyManifestNo,@StartManifest,@EndManifest,@AssyCost,@AddDate,@AddDate)` — column order is implied, writing `@AddDate` to **both** `VC_LAST_UPDATE` and `VC_ADD`. ⚠️ Positional `INSERT` with no column list is **schema-order-fragile** (any column reorder/add silently breaks it) — **pattern P10**. **No uniqueness check inside the proc, and — unlike Supplier/Logistics — no app-side dup check either** (P1 is *absent*; `InsertManifestCostInfo` is single-step, no pre-`SELECT`). Param widths **match** the table (no silent truncation). |
| `UPDATE_ManifestCost;1` (+`@RecordID int`) | UPDATE | Recomputes `@AddDate`; updates one row `WHERE IN_MANIFEST_COST_ID = @RecordID`, rewriting all 5 data columns + `VC_LAST_UPDATE`. **Does not touch `VC_ADD`** (so `VC_ADD` preserves the original insert time; after first update `VC_LAST_UPDATE > VC_ADD`). Keys off the surrogate id directly; **the assembly code is editable** (the proc rewrites `VC_ASSY_PART_NUMBER_CODE`). No uniqueness re-check; nothing stops an update that creates a duplicate code/overlap. |
| `DELETE_ManifestCost;1 @RecordID int` | DELETE | Hard-deletes `WHERE IN_MANIFEST_COST_ID = @RecordID`. Single surrogate param. No soft-delete flag, **no in-use / referential check** — and there is no trigger, so a price referenced by historical invoices can be deleted with no guard (the JOINs in consuming procs then just drop those lines). |
| `SelectSingleField(INV_FORECAST_DETAIL_INF, VC_ASSY_PART_NUMBER_CODE, combo)` | SELECT (raw) | **Not a stored proc** — `DataModule.SelectSingleField` builds dynamic SQL `SELECT DISTINCT(VC_ASSY_PART_NUMBER_CODE) FROM INV_FORECAST_DETAIL_INF ORDER BY VC_ASSY_PART_NUMBER_CODE`, prepends a `' '` blank item, and fills the Assy Code combo. (String-concat dynamic SQL → SQL-injection-shaped, though inputs here are constant.) |

**Downstream consumers of `MO_PRICE` (read-only, listed for completeness — NOT called by this form):**
| Proc | Use of the price |
|------|------------------|
| `REPORT_EDI810` / `REPORT_EDI810Recreate` | The **810 invoice line unit price**. JOIN `INV_ASN_DETAIL_MST d … JOIN INV_MANIFEST_COST_MST m ON d.VC_ASSY_PART_NUMBER = m.VC_ASSY_PART_NUMBER_CODE` — **joins on assy code ONLY; the start/end manifest window is ignored.** ⚠️ |
| `SELECT_INVOICEItems` | Invoice line `Unit Price` + computes `Total Price = MO_PRICE * IN_QTY`. Same assy-code-only JOIN. |
| `REPORT_EDI856` | The **856 ASN report** also carries `m.MO_PRICE` as `UnitPrice` via the same `JOIN INV_MANIFEST_COST_MST` on assy code. (Verified in the proc body — omitted from an earlier draft.) |
| `REPORT_INVOICESSummary` / `REPORT_MonthlyINVOICESSummary` | Invoice summaries, same assy-code-only JOIN. ⚠️ `REPORT_MonthlySupplierInvoices` does **not** join manifest cost — it sums pre-stored `INV_INVOICE_INF` amounts, so it is **not** a `MO_PRICE` consumer (an earlier draft listed it here by mistake). |
| `SELECT_ForecastDetailBCASN` | Forecast/BC-ASN view joins manifest cost on assy code (window ignored). |

> **Key cross-module fact:** every billing consumer keys on `VC_ASSY_PART_NUMBER_CODE`
> alone and **ignores the date window**. So if **two** manifest-cost rows share an assy
> code (which nothing prevents — no unique constraint, no dup check), the `JOIN`
> **multiplies every invoice line for that assembly** (one line per matching price row),
> double-billing the customer. The date window is *only* honored when someone calls
> `SELECT_ManifestCost` with `@ProdDate` — and the invoice procs never do. This is the
> single most important rule for the rebuild (§4, §8). The combination of a billing-critical
> table with **no PK/unique/FK/trigger and no app dup check** is **pattern P11**.

### Call mechanism (legacy)
`DataModule.pas` methods (declarations 555–558; bodies 1620–1813) drive the shared ADO
objects (P6):
- **`GetManifestCostInfo`** (line 1620) uses **`Inv_DataSet`** (open result set) with
  `Parameters.Clear` and **no params** → `SELECT_ManifestCost` defaults → all rows; times
  the call; on success logs `LogActLog('GET PO', 'SELECTED all manifest cost info', 1)`.
  The other three use **`Inv_StoredProc`** (`ExecProc`). (Note the persistent **"PO"**
  vocabulary in logs/messages — leftover from the MonthlyPO ancestor.)
- **`InsertManifestCostInfo`** (Boolean, line 1662) is **single-step — NO P1 dup check**.
  It sets 5 params and `ExecProc`s; `Result := True` only if `Inv_Connection.Errors=0`.
  On failure the form shows `'Unable to INSERT <AssyCode>'`. The new identity is **not**
  captured back; instead the form re-`GetManifestCostInfo`s and `Inv_DataSet.Locate`s the
  new row by `('Assy;Start Manifest;End Manifest')` — a **composite natural-key locate**,
  which would land on the *wrong* row if a duplicate (assy+start+end) already existed.
- **`UpdateManifestCostInfo`** (line 1719) sets the 5 fields + `@RecordID := fRecordID`.
- **`DeleteManifestCostInfo`** (line 1772) passes only `@RecordID := fRecordID`.
- All four share the **retry-up-to-3-times-via-recursion** harness (`fErrorCount < 3`)
  with a `finally` doing `Inv_StoredProc.Close; fErrorCount := 0` (P8).
  ⚠️ **Copy-paste bugs in the retry branches** (each retries the *wrong* module's method) — **pattern P12**:
  `GetManifestCostInfo` → recursively calls **`GetSizeInfo`**; `InsertManifestCostInfo` →
  **`InsertSizeInfo`**; `UpdateManifestCostInfo` → **`UpdateSizeInfo`**;
  `DeleteManifestCostInfo` → **`DeleteSupplierInfo`** (which would `DELETE_SupplierInfo
  @SupplierID := fRecordID` — i.e. **delete a *supplier* row whose id equals the shared
  `RecordID`** on a transient manifest-cost DB error!). These wrong-target retries are
  latent landmines; do **not** reproduce them in the rebuild (§8). The `UPDATE PO` success
  log also concatenates unrelated `fPickUp`/`fPONumber` fields (more MonthlyPO leftovers).

**DataModule properties:** `AssyCode` (405), `AssyManifestNo` (495), `POStart` (496),
`POEnd` (497), `AssyCost: Double` (504). The row key is the **shared, generic `RecordID`**
property (337) — the same one Supplier/Logistics/Shipping/Invoice/ASN reuse, so a stale
`RecordID` from another screen is a real latent cross-module hazard (P9), made worse here
by the Delete-retry's wrong-target `DeleteSupplierInfo` call.

## 4. Business rules & edge cases
- **`MO_PRICE` is the billing unit price.** It is keyed by **assembly code**
  (`VC_ASSY_PART_NUMBER_CODE`) and consumed by invoice/810 procs **on the code alone**
  (date window ignored downstream). The window (`VC_START_MANIFEST`/`VC_END_MANIFEST`) is
  intended to make the price time-bounded, but **only `SELECT_ManifestCost @ProdDate`
  honors it** — and no billing caller passes `@ProdDate`. **Effective rule today:
  there must be exactly one manifest-cost row per assembly code, or invoices double-count.**
- **No uniqueness enforcement at any layer** — no DB constraint, no app-side dup check
  (unlike Supplier's app dup check and Logistics' app check + unique index). Inserting a
  second row for the same assy code is fully allowed and silently breaks billing.
- **Dates are `yyyymmdd` strings** (P2-adjacent): written via `formatdatetime('yyyymmdd',
  date)`, read back into the date editor by `copy(...)` slicing month/day/year. Window
  comparison is **string** comparison (valid because `yyyymmdd` sorts chronologically).
  The `@ProdDate` filter is **strict open interval** `> start AND < end` (boundary dates
  excluded) ⚠️.
- **Price precision:** UI is `$` + 4 decimal places; value round-trips through a Delphi
  `Double` (`TryStrToFloat`), DB column is `money` (4 dp, ±922 trillion). Invalid cost
  text → `ShowMessage('Invalid assembly cost')` + refocus, but **`HoldDetails` does not
  abort** — `AssyCost` keeps its previous value and the Insert/Update proceeds. ⚠️ No
  range/positivity check (a negative or zero price is accepted).
- **Audit timestamps** `VC_ADD` / `VC_LAST_UPDATE` are `yyyymmddHHMMSSff` strings (P2);
  INSERT sets both equal; UPDATE rewrites only `VC_LAST_UPDATE`.
- **Assembly code is editable** (UPDATE rewrites it). Because billing joins by code,
  editing a code **re-prices all matching invoice lines** retroactively (there is no
  versioning) — and could create a duplicate-code collision with no guard.
- **Delete is hard, with no trigger and no RI check.** Deleting a price that historical
  invoices reference makes those invoice/report lines **silently drop** (inner JOIN),
  changing already-issued invoice totals. No cascade, no soft-delete.
- **Insert/Update re-locate by composite natural key** (`Assy;Start Manifest;End
  Manifest`) after re-querying — fragile if duplicates exist; selects the first match.
- **Stale-`RecordID` hazard (P9):** Update/Delete key off `Data_Module.RecordID`, set only
  by `HoldDetails(True)` when a grid row is selected. With no row selected, `RecordID` may
  be 0 or a leftover value from another module — no guard. Compounded by the Delete
  retry-path bug (§3) that could target `DELETE_SupplierInfo`.
- **EDI feature-flag gate:** the screen only opens when `fiGenerateEDI` is true; otherwise
  the legacy `MonthlyPOMaster` opens instead. Both are present in the build. (Per decision D1,
  docs/analysis/decisions.md, this flag — today a `[SITE]` INI value — becomes a per-site
  column on the `sites` table.)

## 5. UI / UX notes
- Grid + detail-panel pattern; selecting a row (or any `OnDataChange`) syncs the panel and
  captures `RecordID` from hidden grid `Fields[0]`.
- **Fields (label → control → DB):** ASSY Code (combo `csDropDownList`, uppercase, list
  from forecast-detail) → `VC_ASSY_PART_NUMBER_CODE`; Cost Start Date / Cost End Date
  (`TNUMMIBmDateEdit` calendar pickers) → `VC_START_MANIFEST`/`VC_END_MANIFEST`
  (`yyyymmdd`); Assembly Cost (`TcurrEdit`, `$`, 4 dp) → `MO_PRICE`; Assy Manifest ID
  (combo `csDropDownList`, hard-coded `' '`,`'01'..'99'`) → `VC_ASSY_MANIFEST_NUMBER`.
- **Search is fully client-side (P7):** `SearchGrid` sets `Inv_DataSet.Filter :=
  '[Assy] LIKE ' + QuotedStr(AssyCode)` and `Filtered := True` (an **in-memory dataset
  filter over the already-loaded rows**, not a re-query). It is a `LIKE` on the **aliased
  grid column `Assy`** — so wildcards are possible only if the user types them; the form
  pre-checks `Trim(combo) = ''` and blocks an empty search. On `RecordCount=0`: "No
  matches were found for your query." Because the combo is `csDropDownList`, the search
  term is always a value that exists in forecast-detail, not necessarily in the cost table.
- **No form-level validation beyond the cost parse:** no required-field check on assy
  code/manifest id/dates, no window-ordering check (start ≤ end is *not* enforced), no
  duplicate check. The only guard is `Invalid assembly cost` (non-abortive).
- **Modernize:** standard index/list + new/edit form; server-side search/sort/pagination
  (P7); real date inputs (drop string `yyyymmdd`); decimal/money input with positivity
  range check; **add the missing uniqueness/overlap rule** (§8) — this is the highest-value
  improvement because it directly affects invoice correctness; default the list ordering
  (the proc has none). Source the Assy Code select from the real assembly-code domain
  (today: distinct forecast-detail codes) and the Manifest-ID select from a proper list.

## 6. Target design  *(Rails primary)*
- **Model:** `ManifestCost` (`self.table_name = 'INV_MANIFEST_COST_MST'`,
  `self.primary_key = 'IN_MANIFEST_COST_ID'`).
  - Attributes: `assy_part_number_code` (`VC_ASSY_PART_NUMBER_CODE`), `assy_manifest_number`,
    `start_manifest`, `end_manifest`, `price` (`MO_PRICE`, map to `:decimal`/Money),
    `vc_add`, `vc_last_update`.
  - **Associations:** **no declared FK in legacy** — relationships are by-convention on the
    string `VC_ASSY_PART_NUMBER_CODE`. Model the read side as a lookup
    (`AsnDetail`/`InvoiceItem` resolve price by assy code). Do **not** add a `dependent:`
    cascade (legacy has no trigger); decide delete behavior explicitly (§8).
    - **`belongs_to :site`** (per decision D1, docs/analysis/decisions.md): every row is
      per-site, with current-site scoping applied to all queries (the legacy unfiltered
      `SELECT_ManifestCost ''` becomes site-scoped). Assy-code lookups by the billing read
      side also scope to the current site.
  - **Validations (deliberate improvements over legacy, which had none):** presence of
    `assy_part_number_code`, `start_manifest`, `end_manifest`, `price`; `price`
    numericality (≥ 0 — confirm with domain expert); `start_manifest <= end_manifest`;
    and the **critical uniqueness/overlap rule** — at minimum a unique index on
    `(site_id, assy_part_number_code)` (if one-price-per-assembly) or a **no-overlapping-window**
    validation per (site, assy code) (if time-bounded pricing is real). Per decision D1
    (docs/analysis/decisions.md) the key is unique **per-site**, so the unique index is the
    composite `(site_id, assy_part_number_code)`, not a global one. This is the fix for the
    double-billing hazard (§4, §8).
  - **No enums** (no P4). `assy_manifest_number` is a free 2-char string (`'01'..'99'`).
  - Timestamps: write `vc_add` on create and `vc_last_update` on create+update as
    `yyyymmddHHMMSSff` strings during **parallel run** (P2); normalize at the Postgres phase.
    Note `VC_LAST_UPDATE` underscore naming differs from the other masters — keep the real
    column name.
- **Controller/routes:** RESTful `resources :manifest_costs`.
- **Views:** index (server-side searchable/paginated list, replacing the in-memory `LIKE`
  filter, P7) + new/edit form. Assy Code select sourced from the assembly-code domain;
  Manifest-ID select from a list; real date pickers; money input. Surface the EDI
  feature-flag gate as a route/menu guard (parity with `fiGenerateEDI`) — note that per
  decision D1 (docs/analysis/decisions.md) this flag, formerly read from the `[SITE]` INI
  section, now lives as a per-site column on the `sites` table.
- **Services:** none needed for CRUD. **Stage-1 option:** wrap the four existing procs
  (`SELECT/INSERT/UPDATE/DELETE_ManifestCost`) via `tiny_tds` for guaranteed parity — but
  **do NOT reproduce the wrong-target retry recursion** (§3); use a single connection/retry
  policy (P8). The billing read path (`SELECT_INVOICEItems`, `REPORT_EDI810*`) is owned by
  the Invoice/EDI module — coordinate the "price lookup by assy code, window-aware?"
  decision there.
- **Reports:** none owned by this module; it *feeds* the 810/invoice reports.

## 7. Migration plan for this module
- [ ] **Stage 1 — read-only parity:** `ManifestCost.all` (or wrap `SELECT_ManifestCost ''`)
      renders the list. Match the legacy lack of ordering initially (or add a sensible
      default order and note the divergence). Server-side search replaces the in-memory
      `LIKE` filter (P7). No writes.
- [ ] **Stage 2 — writes via wrapped procs:** call `INSERT/UPDATE/DELETE_ManifestCost`
      through `tiny_tds`. Preserve the `VC_ADD`/`VC_LAST_UPDATE` string writes (both equal
      on insert). **Replace** the per-method recursive retry (and its buggy wrong-target
      calls) with a single retry policy (P8). Add the new uniqueness/overlap validation in
      the app layer even at this stage (the legacy proc won't, and billing depends on it).
- [ ] **Stage 3 — reimplement (Postgres-ready):** ActiveRecord validations (presence,
      money ≥ 0, start ≤ end, uniqueness/overlap on assy code) replace the absent legacy
      checks; add the missing DB constraints (a real PK on `IN_MANIFEST_COST_ID` and a
      unique/exclusion constraint on assy code / window); real `date` and `decimal/money`
      columns replace `yyyymmdd`/string-audit columns; decide and implement delete RI (§8).
      Per decision D1 (docs/analysis/decisions.md), this Postgres phase also adds the
      `site_id` (NOT NULL) FK referencing the new `sites` table and makes the assy-code
      uniqueness rule per-site — i.e. the unique/exclusion constraint becomes composite
      `(site_id, VC_ASSY_PART_NUMBER_CODE[, window])`. The legacy single-site DB is left
      untouched during the parallel run.

## 8. Open questions for the user (domain expert)
1. **One price per assembly, or time-bounded prices?** The schema has start/end windows,
   but **every billing consumer ignores them** (joins on assy code only). Is the intent
   (a) exactly one current price per assembly (then make assy code UNIQUE and drop/repurpose
   the window), or (b) genuinely time-bounded pricing (then the **invoice/810 procs are
   buggy** — they should filter by the ASN production date within the window, and the
   rebuild must add a no-overlapping-window constraint)? This decision directly affects
   invoice correctness and is the most important open question.
2. **Duplicate-code safety / double-billing:** today nothing prevents two manifest-cost
   rows for the same assy code, and the invoice JOINs would then emit duplicate (doubled)
   invoice lines. Has this ever happened? What should the rebuild enforce — unique code, or
   unique non-overlapping window per code?
3. **Delete with referencing invoices:** deleting a price silently removes the matching
   invoice/report lines (inner JOIN, no trigger, no RI). Should deletion be **blocked** when
   the assy code is referenced by any ASN-detail/invoice, **soft-deleted**, or is hard
   delete acceptable because prices are only ever future-dated?
4. **Multi-site scope:** ✅ RESOLVED (D1, docs/analysis/decisions.md): **per-site** —
   assembly prices are NOT shared. `INV_MANIFEST_COST_MST` (today no site/plant column, every
   query returning all rows unfiltered) gains a `site_id` (NOT NULL) FK to the new `sites`
   table; every row belongs to one site and every query is scoped to the current site. The
   billing key `VC_ASSY_PART_NUMBER_CODE` becomes unique **per-site** (composite
   `(site_id, VC_ASSY_PART_NUMBER_CODE)`), not global — so the uniqueness/overlap rule of
   §8.1/§8.2 is enforced within a site. (Consistent with billing being per customer-plant.)
5. **Assembly-code domain:** the Assy Code combo is sourced from `DISTINCT
   VC_ASSY_PART_NUMBER_CODE` in `INV_FORECAST_DETAIL_INF`. Is that the correct/authoritative
   list of billable assemblies, or should it come from a dedicated assembly master?
6. **EDI feature-flag gate:** the screen replaces the legacy `MonthlyPOMaster` only when
   `fiGenerateEDI` is on. Is `MonthlyPOMaster` still in use anywhere (non-EDI sites), or can
   the rebuild drop it and keep only Manifest Cost?
7. **Negative/zero price and `start > end`:** the form accepts them silently. Are these ever
   legitimate, or should the rebuild reject them?
8. **`VC_LAST_UPDATE` (underscore) vs `VC_LASTUPDATE`:** confirm this column naming
   divergence from the other masters is intentional (it is real in the schema), so the
   rebuilt model maps the right column.

## 9. Test cases / parity checks
- **List all** → row count matches `SELECT_ManifestCost ''`; columns map 1:1 to grid
  `Fields[0..5]` (`ID, Assy, Manifest ID, Start Manifest, End Manifest, Cost`). Legacy has
  **no `ORDER BY`** — assert the new app's chosen ordering and document any divergence.
- **Insert a new price** → one new row; `VC_ADD` **and** `VC_LAST_UPDATE` both set to the
  same `yyyymmddHHMMSSff` string; `MO_PRICE` stored with 4-dp precision; identity assigned
  (legacy does not echo the new id back — verify the new app persists/returns it).
- **Insert a second row with the same assy code** → legacy **allows it** (no DB constraint,
  no app dup check). Decide & assert the new-app behavior (recommended: reject — document
  the intentional divergence) and verify it prevents the double-billing JOIN.
- **Invoice line price** → for an ASN detail with assy code X, `SELECT_INVOICEItems` returns
  `Unit Price = MO_PRICE` for X and `Total = MO_PRICE * IN_QTY`, **regardless of the
  manifest window**. With two X rows present, legacy returns **two** lines (doubled) —
  assert the new app's behavior under the §8.1/§8.2 decision.
- **`SELECT_ManifestCost @AssyCode=X, @ProdDate=D`** → returns the row only when
  `Start < D < End` (strict, boundary excluded). Verify a `D` equal to `Start` or `End`
  returns **no** row in legacy.
- **Update** → row updated by `IN_MANIFEST_COST_ID`; `VC_LAST_UPDATE` advanced, `VC_ADD`
  unchanged; all 5 data fields rewritten including the assy code. Editing the code re-prices
  matching invoice lines (no versioning) — verify the new app's behavior matches the §8
  decision.
- **Update with `start > end`** → legacy accepts; assert new-app validation (recommended:
  reject).
- **Invalid cost text on Insert/Update** → legacy shows "Invalid assembly cost" but **does
  not abort** (uses the prior `AssyCost`); confirm the new app rejects/aborts cleanly.
- **Delete a price referenced by historical invoices** → legacy deletes with no guard and
  the referencing invoice/report lines vanish from output; assert the new app's chosen RI
  behavior (block / soft-delete / cascade) per §8.3.
- **Delete with no row selected** → legacy keys off a possibly-stale shared `RecordID`
  (P9); confirm the new app's id comes from `params[:id]` and cannot target a stale row
  (and never falls back to a `DELETE_SupplierInfo` path — the legacy retry bug).
