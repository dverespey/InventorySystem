# Module Analysis: Inventory Management (screen + QuickReport)

**Area:** Inventory / Stock  **Status:** ✅ spec complete  **Analyst:** Claude / 2026-06-05

> Covers **two** live units in one spec: `InvMgmt` (the on-screen inventory listing) and
> `InvMgmtQReport` (its printed QuickReport). Both are read-only over the **parts-stock
> master** `INV_PARTS_STOCK_MST` — the screen is essentially a **read-only grid/report of
> current stock levels**, not an editor. It does **not** call any `REPORT_*` proc; it reuses
> the same parts-stock list proc (`SELECT_PartsStockInfo`) that the Parts Stock master editor
> uses, and renders it both on screen and on paper.
> The headline finds: (1) the entire **search/filter feature is dead code** (commented out);
> (2) the screen's `IN_QTY` column is a **denormalized running balance maintained by triggers
> elsewhere** (receiving / shipping / reject / stocktaking) — this module only *reads* it but
> it is the single most important inventory invariant in the system; (3) the QuickReport binds
> directly to the ADO dataset, so **"Print" and "Print Excel" produce the identical report**.

## 1. Legacy surface
- **Forms (2):**
  - `InvMgmt.pas` (9.0 KB) + `InvMgmt.dfm` (5.6 KB) — `TInvMgmt_Form`, Caption "Inventory
    Management", header label "Inventory Management". Author: Aaron Huge, 2002-10-25.
    Registered live in `InventorySystem.dpr`: `InvMgmt in 'InvMgmt.pas' {InvMgmt_Form}`.
  - `InvMgmtQReport.pas` (3.2 KB) + `InvMgmtQReport.dfm` (19.8 KB) — `TInvMgmtQReport_Form`,
    a **QuickReport (`TQuickRep`)** named `InvMgmt_QuickRep`. Registered live in the dpr:
    `InvMgmtQReport in 'InvMgmtQReport.pas' {InvMgmtQReport_Form}`. Same author/date.
- **Entry point:** `MainMenu.pas` owns an `InvMgmt_GroupBox` + `InvMgmt_Button` (and menu items
  `Window_InvMgmt_InvMgmt_MenuItem` / `Window_InvMgmt_Stocktaking_MenuItem`). The button handler
  `TMainMenu_Form.InvMgmt_ButtonClick` (lines 342-349) is the standard **P14** child-launch
  idiom: `Hide; InvMgmt_Form := TInvMgmt_Form.Create(self); InvMgmt_Form.Execute; InvMgmt_Form.Free;
  Show;` — **no `try..finally`**, so a child exception leaves MainMenu hidden (P14 hazard).
  `Execute` (InvMgmt.pas:70-87) calls `SetCombos; ShowModal;` inside a `try..except` that pops
  "Unable to generate Inventory Management screen." on any error; returns `False` only on
  `mrCancel` (the Close button, `ModalResult = 2 = mrCancel`).
- **Report entry:** from the screen, **both** `Print_Button` **and** `PrintExcel_Button` point at
  the **same** handler `Print_ButtonClick` (InvMgmt.pas:224-231): `Hide;
  InvMgmtQReport_Form := TInvMgmtQReport_Form.Create(self); InvMgmtQReport_Form.InvMgmt_QuickRep.Preview;
  InvMgmtQReport_Form.Free; Show;`. ⚠️ It calls `.Preview` directly (not the form's own `Execute`),
  and **does not branch on Excel vs print** — the "Print Excel" button is mislabeled / does the
  same thing. (QuickReport's preview window offers an "export to Excel" action, so the only way to
  get Excel is manually from the preview dialog — there is no code path that differs.)
- **Purpose (one paragraph):** A read-only **inventory / stock-level browser + printout**. On open
  it loads the **entire parts-stock master** (every part with its current on-hand `IN_QTY` and all
  its planning attributes) into a grid, hiding the engineering/`Report*` columns. Four "search key"
  combos (Supplier, Part Name, Line, Part Type) sit above the grid but the search that would use
  them is **commented out** (§4). "Clear" reloads the full list; "Print"/"Print Excel" preview the
  same QuickReport of the loaded rows. It writes nothing.

## 2. Data touched
| Table | Read | Write | Notes |
|-------|:----:|:-----:|-------|
| `INV_PARTS_STOCK_MST` | ✓ |  | The parts-stock master — the grid + report rows (via `SELECT_PartsStockInfo`). **`IN_QTY` is the current on-hand balance.** This module only reads it. |
| `INV_SUPPLIER_MST` | ✓ |  | LEFT JOIN in the list proc (supplier code/name); also feeds the Supplier combo |
| `INV_LOGISTICS_MST` | ✓ |  | LEFT JOIN (logistics name) |
| `INV_RENBAN_GROUP_MST` | ✓ |  | LEFT JOIN (renban group code) |
| `INV_PART_TYPE_MST` | ✓ |  | LEFT JOIN (part type); also feeds the Part Type combo |
| `INV_SIZE_MST` | ✓ |  | LEFT JOIN (size code) |
| `LINE` (Activity/ALC DB) | ✓ |  | `SelectSingleFieldALC('LINE','LineName',…)` — cross-catalog read for the **Line** combo (Activity/ALC connection, not Inventory) |

This module **persists nothing** — it is the only "Inventory / Stock" surface analyzed so far that
is purely a reader/printer. All the `IN_QTY`-mutating work happens in **other** modules
(receiving, shipping, rejects, stocktaking) via the triggers below; this screen just displays the
result.

### `INV_PARTS_STOCK_MST` columns (authoritative: `DB Schema/Create Inventory.sql`)
The table is wide (33 columns). The ones the list proc returns / this module relies on:
| Column | Type | Meaning / notes |
|--------|------|-----------------|
| `IN_PART_ID` | `int IDENTITY(1,1) NOT NULL` PK | Surrogate key. Returned as `'RecordID'`; captured into the **shared** `Data_Module.RecordID` indirectly? — **NO** (see note). Grid `Fields[?]='RecordID'` is *not* read by `HoldDetails` here. |
| `VC_PART_NUMBER` | `varchar(12) NOT NULL` | **Business key — DB-unique via `IX_INV_PARTS_STOCK_MST` (UNIQUE NONCLUSTERED on `VC_PART_NUMBER`)**. Returned as `'Parts Code'`. ⚠️ Multi-site (D1): uniqueness becomes **composite `(site_id, VC_PART_NUMBER)`**, per-site, not global. |
| `IN_SUPPLIER_ID` | `int NULL` | FK→`INV_SUPPLIER_MST` (by convention; **no declared FK**). Nulled by `DELETE_SupplierCode`. |
| `IN_LOGISTICS_ID` | `int NULL` | FK→`INV_LOGISTICS_MST` (by convention). Nulled by `DELETE_LogisticsCode`. |
| `IN_RENBAN_ID` | `int NULL` | FK→`INV_RENBAN_GROUP_MST` (by convention). Nulled by `DELETE_RenbanGroupCode`. |
| `IN_PART_TYPE_ID` | `int NULL` | FK→`INV_PART_TYPE_MST` (by convention). |
| `IN_SIZE_ID` | `int NULL` | FK→`INV_SIZE_MST` (by convention). Nulled by `DELETE_SizeCode`. |
| `VC_PARTS_NAME` | `varchar(50) NULL` | `'Parts Name'`. |
| `VC_KANBAN_NUMBER` | `varchar(5) NULL` | `'KANBAN'`. |
| `VC_LINE_NAME` | `varchar(10) NULL`, **DEFAULT `'TUNDRA'`** | `'Line Name'`. On the **report** this is mislabeled (the QR field is named `CarTruck_QRDBText` but bound to `'Line Name'`). |
| `IN_1LOTQTY` | `int NULL` | `'1 Lot QTY'`. |
| `BIT_LOT_SIZE_ORDERS` | `bit NULL` | `'Lot Size Orders'`. |
| `IN_LEADTIME`, `IN_LEADTIME_MONDAY..SATURDAY` | `int NULL` ×7 | per-day lead times. |
| `IN_RENBAN_COUNT` | `int NULL` | `'Renban Count'`. |
| `IN_SHIP_DAYS`, `IN_SHIP_DAYS_MONDAY..SATURDAY` | `int NULL` ×7 | per-day ship days. |
| **`IN_QTY`** | **`int NULL`** | **`'QTY'` — the current on-hand stock balance. The core inventory invariant; maintained ONLY by the qty triggers (§2 triggers), never written by this module.** |
| `VC_COMMENTS` | `varchar(300) NULL` | `'Remarks'`. |
| `MO_PART_COST` | `money NULL`, **DEFAULT `0`** | `'Part Cost'`. (Selected by the proc but not shown in the QReport detail band; the report's price columns are commented out.) |
| `VC_LAST_UPDATE` | `varchar(16) NULL` | **Timestamp as `yyyymmddHHMMSSff` string** (P2). Written by the qty triggers (they stamp `i.VC_LAST_UPDATE`), not by this module. |
| `VC_ADD` | `varchar(16) NOT NULL` | **Timestamp string** (P2); set at part creation (elsewhere). Used by the history-purge subquery. |

**Constraints / indexes (authoritative):**
- `PK_INV_PARTS_STOCK_MST` PRIMARY KEY **CLUSTERED** (`IN_PART_ID`).
- `IX_INV_PARTS_STOCK_MST` **UNIQUE NONCLUSTERED (`VC_PART_NUMBER`)** — the part number is
  globally unique at the DB level (a real backstop; this module relies on it indirectly via the
  triggers that join on `VC_PART_NUMBER`). ⚠️ Multi-site (D1): this becomes a **per-site**
  composite unique index **`(site_id, VC_PART_NUMBER)`** in the Postgres phase — part numbers are
  unique within a site, not globally.
- DEFAULTs: `VC_LINE_NAME` → `'TUNDRA'`, `MO_PART_COST` → `0`.
- **No declared FOREIGN KEY constraints** out of this table. The whole schema declares **only 2
  FKs total** (`INV_ASN_DETAIL_MST→INV_ASN_MST`, `INV_PART_SHIPPING_INF→INV_SHIPPING_INF`, both
  `ON DELETE CASCADE`). The five `IN_*_ID` links are **by convention only**, enforced solely by
  the master-delete triggers (P5).
- History twin: `INV_PARTS_STOCK_MST_HIST` — written by `INSERT_PartsStockMST` and
  `UPDATE_PartNumber` triggers (§2 triggers). ⚠️ **It is NOT an exact column-shape twin of the
  master.** Two divergences make the positional copy a landmine: (1) its line-name column is
  **`VC__LINE_NAME` (double underscore)** vs the master's `VC_LINE_NAME`; (2) some HIST columns are
  **`NOT NULL`** (`VC_PARTS_NAME`, `IN_RENBAN_COUNT`) where the master allows **`NULL`**. Both
  trigger writes are a **positional `INSERT INTO …_HIST SELECT * FROM inserted`/`deleted`** (no
  column list), so they rely on column **order** matching and will **fail** the moment a master row
  with a NULL `VC_PARTS_NAME`/`IN_RENBAN_COUNT` is inserted/updated (NOT-NULL violation on the
  history copy). The rebuild must either reproduce this column-for-column (incl. the double-
  underscore name and the NOT-NULL/NULL mismatch) or — better — fix the HIST schema and use an
  explicit column-list insert.

**Triggers on these tables (LIVE — from `DB Schema/Create Inventory.sql`; the authoritative source.
`docs/triggers.sql` is obsolete/stale and must not be used — see
[`trigger-source-reconciliation.md`](../cross-cutting/trigger-source-reconciliation.md)):**

*Triggers ON `INV_PARTS_STOCK_MST` (the table this module reads):*
- **`INSERT_PartsStockMST`** (FOR INSERT) → `INSERT INTO INV_PARTS_STOCK_MST_HIST SELECT * FROM
  inserted`. **Invariant: every new part row is mirrored into the history table.**
- **`UPDATE_PartNumber`** (FOR UPDATE) → (a) `INSERT INTO INV_PARTS_STOCK_MST_HIST SELECT * FROM
  deleted` (snapshots the pre-image); (b) when `IN_QTY` changes, writes a **qty ledger row** into
  `INV_PART_QTY_INF (VC_PART_NUMBER, IN_QTY_CHANGE, IN_QTY, VC_STATUS, VC_ADD)` with
  `IN_QTY_CHANGE = deleted.IN_QTY - inserted.IN_QTY`, `VC_STATUS='U'`; (c) only when **exactly one
  row** changes **and** `update(vc_part_number)`, **propagates the renamed part number** into
  `INV_ASSY_RATIO_MST` (`VC_TIRE_PART_NUMBER1/2_CODE`, `VC_WHEEL_PART_NUMBER1/2_CODE`) and
  `INV_FORECAST_DETAIL_INF` (`VC_TIRE/WHEEL_PART_NUMBER_CODE`). **Invariants: keep history + a
  qty-change ledger; cascade a part-number rename to the assembly-ratio & forecast-detail string
  references.** (Those child tables still reference the part by **string code**, not by `IN_PART_ID`
  — a code rename must be hand-propagated; this trigger is that propagation.)
- **`DELETE_PartNumber`** (FOR DELETE) → writes an Activity audit row, then **blanks the part-code
  string references** in `INV_ASSY_RATIO_MST` (4 columns) and `INV_FORECAST_DETAIL_INF` (2 columns)
  by `SET … = ''` where they matched the deleted `VC_PART_NUMBER`. **Invariant: deleting a part
  unlinks (blanks) its references in assembly-ratio & forecast-detail, does not delete them.**
  (Note: unlike the master-`*Code` triggers which null an int FK, this one blanks **string** code
  columns, because assy-ratio/forecast still key parts by string code.)

*Triggers that MAINTAIN `INV_PARTS_STOCK_MST.IN_QTY` (fired from OTHER tables — the balance this
screen displays). These are the core inventory invariant the rebuild must re-home:*
- **`INSERT_RecConfStatPartsStockMstQTY`** (on `INV_OPEN_ORDER_INF` FOR INSERT) → mirrors the open
  order into `INV_OPEN_ORDER_INF_HIST`, then **adds** the received qty:
  `UPDATE INV_PARTS_STOCK_MST SET IN_QTY = PS.IN_QTY + i.IN_QTY, VC_LAST_UPDATE = i.VC_LAST_UPDATE`
  joined `PS.VC_PART_NUMBER = i.VC_PART_NUMBER` and `PS.IN_SUPPLIER_ID = s.IN_SUPPLIER_ID`. **Two
  branches keyed on the supplier's `VC_INVENTORY_ADD_POINT`:** add when `VC_STATUS_SUPPLIER_SHIPPING
  <> ''` if add-point = **`'S'`** (shipped); add when arrival/plant-yard/assembler-yard/warehouse
  status `<> ''` if add-point = **`'A'`** (arrived). **Invariant: receiving raises on-hand by the
  order qty, but the *moment* it counts (at-ship vs at-arrival) depends on the supplier's add-point
  flag (P4 enum `S/A`).**
- **`UPDATE_RecConfStatPartsStockMstQTY`** (on `INV_OPEN_ORDER_INF` FOR UPDATE) — the symmetric
  delta-adjustment as open-order rows change status/qty (mirror of the insert/delete pair).
  **Invariant: status/qty edits on an open order re-balance `IN_QTY`** — *but* note the dead branch
  below. ⚠️ **DEAD BRANCH:** the "changed to **not** arrived" update (the path meant to *lower*
  stock when an add-point-`'A'` order is un-arrived) is gated on
  `WHERE i.VC_ARRIVAL = '' AND i.VC_ARRIVAL <> ''` — a contradiction that is **always false**, so
  that path **never fires**. Net: an add-point-`'A'` order that is edited back to "not arrived"
  does **not** reverse its earlier on-hand increment (the stock stays raised). Any blanket
  "status/qty edits re-balance `IN_QTY`" claim must be qualified by this gap; the rebuild should
  implement the un-arrive reversal correctly (and not copy the broken predicate).
- **`DELETE_RecConfStatPartsStockMstQTY`** (on `INV_OPEN_ORDER_INF` FOR DELETE) → guarded by
  `SELECT @PurgeMode = PurgeMode FROM Purge` (only runs when `@PurgeMode = 0`, i.e. **a real delete,
  not a data-purge**), then **subtracts**: `IN_QTY = PS.IN_QTY - d.IN_QTY` for the same two
  add-point branches, additionally requiring `VC_STATUS_EMPTY_TRAILER = '' AND VC_TERMINATED = ''`.
  **Invariant: undoing a receipt lowers on-hand by that qty — unless it was an empty-trailer or
  terminated row, or the row is being purged (purge must not move stock).**
- **Reject family** `INSERT/UPDATE/DELETE_RejectParts` (on `INV_REJECT_INF`) → adjust `IN_QTY`
  keyed on `IN_PART_ID`; the DELETE **adds the deleted reject qty back** (`IN_QTY = PS.IN_QTY +
  d.IN_QTY`). **Invariant: rejecting parts lowers on-hand; un-rejecting restores it.**
- **Stocktaking family** `INSERT/UPDATE/DELETE_Stocktaking` (on `INV_STOCKTAKING_INF`) → adjust
  `IN_QTY` by an **additive delta** keyed on `IN_PART_ID` (same shape as the reject family, **not**
  a set/override): `INSERT_Stocktaking` → `IN_QTY = PS.IN_QTY + i.IN_QTY`; `DELETE_Stocktaking` →
  `IN_QTY = PS.IN_QTY - d.IN_QTY`; `UPDATE_Stocktaking` does both (subtract the deleted row, then
  add the inserted row). **Invariant: a stocktaking entry adds/subtracts its counted qty as a delta
  to the running balance — it does *not* overwrite the balance with the counted value.**
- **Shipping family** `InsertPartShipping`, `UpdatePartShipping`, `DeletePartShipping`,
  `DeleteShipDate` (on `INV_PART_SHIPPING_INF` / `INV_SHIPPING_INF`) → ship out lowers `IN_QTY`;
  delete/undo restores it. **Invariant: shipping reduces on-hand.**
- **`INSERT_InvPartQtyInf`** (on `INV_PART_QTY_INF` FOR INSERT) → writes an Activity audit row
  only (no qty math). **Invariant: every qty-ledger insert is audit-logged.**

> ⚠️ **The qty-maintenance triggers use TWO different join keys to `INV_PARTS_STOCK_MST`.**
> Receiving (`*_RecConfStatPartsStockMstQTY`, from `INV_OPEN_ORDER_INF`) and shipping
> (`*PartShipping`, from `INV_PART_SHIPPING_INF`) join on **`VC_PART_NUMBER`** (the `varchar(12)`
> string), while the **reject** and **stocktaking** families join on **`IN_PART_ID`** (the `int`
> surrogate). A part-number **rename mid-flight** (between an open-order/shipping post and the
> matching reversal) could **mis-key the string-joined adjustments** — they'd find no row (or the
> wrong one) and silently fail to re-balance, whereas the int-keyed reject/stocktaking posts are
> rename-immune. The rebuild's stock-ledger should key **every** post on the stable surrogate
> (`IN_PART_ID`), not the mutable part-number string.

> **Net:** `INV_PARTS_STOCK_MST.IN_QTY` is a **denormalized running balance** kept in sync by
> receiving (+), shipping (−), rejects (∓), and stocktaking (∓, additive delta) triggers, with the
> receiving/shipping timing gated by the supplier's `S`/`A` add-point. This module only **reads**
> that balance — but the rebuild must re-home these 12+ triggers as a single transactional
> stock-ledger service so that the displayed `IN_QTY` stays correct (see §6).

## 3. Stored procedures used
(Read with `sql.sh proc NAME`. The procs are the behavioral spec.)

| Proc / call | Op | Business rule (from body) |
|-------------|----|---------------------------|
| `SELECT_PartsStockInfo;1 @PartNum varchar(12) = ''` | SELECT | **The grid + report data source.** Called via `Data_Module.GetInventoryInfo` with `@PartNum=''` → returns **all** parts. If `@PartNum<>''` returns the one part `WHERE VC_PART_NUMBER=@PartNum`. Both branches `SELECT` the **same 30 UI-aliased columns** from `INV_PARTS_STOCK_MST p` with **5 LEFT OUTER JOINs** (supplier, logistics, renban group, part type, size — all by `IN_*_ID`). Returns aliases incl. `'Supplier Code','Parts Code','Logistics Name','Parts Name','Renban Group','Part Type','Line Name','Size Code','KANBAN','1 Lot QTY','Lot Size Orders', 7×'Lead Time *', 'Renban Count', 7×'Ship Days *', 'QTY' (=IN_QTY), 'Remarks', 'Part Cost' (=MO_PART_COST), 'RecordID' (=IN_PART_ID)`. **No `ORDER BY`** (rows arrive in clustered-PK order = `IN_PART_ID`). Because of LEFT JOINs, a part with a null/dangling FK still shows (its joined code blank) — important after a master delete nulls a link (P5). |
| `SELECT_PartsSupplier;1 @VC_SUPPLIER_CODE varchar(5)` | SELECT | Drives the **Supplier combo `OnChange`** via `SelectDependantSingleField('SELECT_PartsSupplier','@VC_SUPPLIER_CODE','VC_PART_NUMBER', <supplier code>, PartNum_ComboBox)` — i.e. picking a supplier repopulates the Part-Name combo with that supplier's part numbers. `@code=''/' '` → all parts joined to suppliers; else `JOIN … AND VC_SUPPLIER_CODE = @code`. ⚠️ **param is `varchar(5)`** (supplier code width) — fine here. |
| *(ad-hoc text)* `SELECT VC_SUPPLIER_CODE, VC_SUPPLIER_NAME FROM INV_SUPPLIER_MST` | SELECT | `SelectMultiField` → fills the 2-column **Supplier combo** (code + name). Dynamic SQL string concatenation (no proc). |
| *(ad-hoc text)* `SELECT DISTINCT(VC_PART_NUMBER) FROM INV_PARTS_STOCK_MST ORDER BY …` | SELECT | `SelectSingleField` → fills the **Part-Name combo**. |
| *(ad-hoc text)* `SELECT DISTINCT(VC_PART_TYPE) FROM INV_PART_TYPE_MST ORDER BY …` | SELECT | `SelectSingleField` → fills the **Part-Type combo**. |
| *(ad-hoc text, ALC/Activity DB)* `SELECT DISTINCT(LineName) FROM LINE ORDER BY …` | SELECT | `SelectSingleFieldALC` → fills the **Line combo** from the Activity/ALC catalog. ⚠️ **this method carries the P12 wrong-target retry bug** (see Call-mechanism). |

**There are NO `REPORT_*` procs in this module.** The QuickReport reuses the same on-screen
`SELECT_PartsStockInfo` result (bound to `Inv_DataSet`). This differs from the report-heavy modules
(Shipping/Invoice/Forecast) that have dedicated `REPORT_*` procs — here "report" = "print the grid".

### Call mechanism (legacy)
- **`GetInventoryInfo`** (`DataModule.pas:1305-1348`): opens `Inv_DataSet` (an ADO `TADODataSet`,
  `CursorType=ctStatic`) with `CommandText='dbo.SELECT_PartsStockInfo;1'`, one param `@PartNum=''`.
  On error logs `'Unable to get Part Master data'` (note: labelled "Part Master", shared with the
  Parts Stock editor). Logs success as `LogActLog('GET PARTS','SELECTED all parts',1)`. Wrapped in
  the **P8** recursive retry harness (`fErrorCount < 3 → GetInventoryInfo`; `finally fErrorCount := 0`).
  This retry target is **correct** (self-recursion) — *not* in the P12 register.
- **Dataset plumbing (`DataModule.dfm`):** `Inv_DataSet` (ADO) → `Inv_DataSetProvider`
  (`TDataSetProvider`, `DataSet=Inv_DataSet`) → `Grid_ClientDataSet` (`TClientDataSet`,
  `ProviderName='Inv_DataSetProvider'`) → `Inv_DataSource.DataSet = Grid_ClientDataSet`.
  **But this form's own `Inventory_DataSource` is bound directly to `Inv_DataSet`** in `FormCreate`
  (`Inventory_DataSource.DataSet := Data_Module.Inv_DataSet`), so the grid reads the **raw ADO
  dataset**, while `Grid_ClientDataSet` is only manipulated for `Filtered := False` toggling.
- **`JustifyColumns(InvMgmt_DBGrid)`** (`DataModule.pas:5951-5961`): right-justifies any column
  named `'Price'`/`'Total Cost'`, and **hides any column whose name starts with `'Report'`**.
  (`SELECT_PartsStockInfo` returns no `Report*` columns, so nothing is hidden here — this is shared
  grid-prep code reused from cost reports.)
- **`SetCombos`** (InvMgmt.pas:213-222): four combo-fill calls (Supplier multi-field, Part single,
  Part-Type single, Line single via **ALC**). ⚠️ **`SelectSingleFieldALC`** (`DataModule.pas:5716-5765`)
  is one of the **P12 LOW-severity** wrong-target bugs: its retry branch calls **`SelectSingleField`**
  (the **Inv**-connection variant) instead of `SelectSingleFieldALC`, so a transient error reading
  the `LINE` table re-reads it via the **wrong server** (Inventory, not Activity); its `finally` also
  closes `Inv_Field_DataSet`, not the `ALC_DataSet` it opened. Read-only, so worst case is a wrong/
  empty Line list — but the rebuild must not reproduce it (see
  [`datamodule-retry-target-bugs.md`](../cross-cutting/datamodule-retry-target-bugs.md)).
- **`SelectDependantSingleField`** (`DataModule.pas:5817-5865`): supplier→parts cascade; its retry
  target is correct (self). It passes `Supplier_NUMMIColumnComboBox.ColumnItems[ItemIndex,0]`
  (column 0 = the supplier **code**) to `SELECT_PartsSupplier`.
- **QuickReport (`InvMgmtQReport`)**: `InvMgmt_QuickRep.DataSet = Data_Module.Inv_DataSet` and
  **every detail `QRDBText.DataSet = Data_Module.Inv_DataSet`** (bound in the `.dfm`, not in code).
  `FormCreate` (lines 81-105) is a `try..except` whose body is **entirely commented out** (it would
  have rebound fields to `Grid_ClientDataSet`); on exception it logs
  `LogActLog('ERROR','InvMgmtReport: '+e.Message+' Class:'+e.ClassName,0)`. So the report simply
  prints the already-open `Inv_DataSet`. `Execute` (ShowModal) exists but is **bypassed** — the
  screen calls `.Preview` directly. **`PrintIfEmpty = True`** (`.dfm`) → it previews even with 0 rows.

## 4. Business rules & edge cases
- **Read-only, no writes.** No insert/update/delete proc; no `RecordID`/`fRecordID` set here.
  `HoldDetails(True)` copies the selected grid row into shared `Data_Module` fields
  (`SupplierCode, PartNum, PartType, LineName, SizeCode, Kanban, Quantity := Fields[26].AsInteger`)
  but **nothing reads those fields back** in this module (`SetDetailBoxes` is fully commented out —
  there are no detail edit boxes on the form). These assignments are vestigial copy-paste from the
  Parts Stock editor and have **no effect** here. (`Fields[26]` = the 27th column = `'QTY'`/`IN_QTY`.)
- **The search feature is DEAD CODE.** `SearchGrid` (the in-memory `Filter`/`Filtered` routine,
  InvMgmt.pas:89-133) and the entire body of `Search_ButtonClick` (246-269) are **commented out**.
  `Search_ButtonClick` therefore does **nothing**; the **Search button is inert**. The Supplier
  combo `OnChange` *does* repopulate the Part combo (live), but the Part/Line/PartType combo
  `OnChange` handlers are **empty stubs** (`//Clear other items and update list`). **Net behavior:
  the screen always shows the full unfiltered parts list**; the "Searching Key" group box is
  decorative. (Contrast Supplier/Size masters, where `SearchGrid` is live.) This is a **P7** pattern
  that was *intended* (client-side filter over the loaded grid) but never wired up.
- **`Clear` reloads everything:** `Clear_ButtonClick` does `Grid_ClientDataSet.Filtered := False;
  GetInventoryInfo; ClearControls(SearchKey_GroupBox); HoldDetails(True);` — i.e. re-fetch the full
  list and reset the (non-functional) combos. `ClearControls` sets each `TComboBox.ItemIndex := 0`
  (the blank first row each `SelectField` inserts).
- **Add-point semantics surface here indirectly:** the `IN_QTY` shown is correct **only if** the
  supplier's `VC_INVENTORY_ADD_POINT` (`S`=shipped / `A`=arrived, P4) is set as intended — the
  receiving triggers add stock at different lifecycle moments based on it. A part whose supplier has
  add-point unset gets **no** receiving increment (neither branch matches), so its on-hand can read
  low. The screen has no way to show or fix this; it is a data-quality dependency to flag.
- **Null/blank joined codes:** because `SELECT_PartsStockInfo` uses LEFT JOINs, a part whose
  `IN_SUPPLIER_ID`/`IN_SIZE_ID`/etc. was nulled by a master-delete trigger (P5) still appears, with
  the corresponding code column blank. The grid/report must tolerate blanks.
- **Timestamps (P2):** `VC_ADD` (insert) / `VC_LAST_UPDATE` (update) are `yyyymmddHHMMSSff` 16-char
  strings; written by the qty triggers/part editor, not by this module. Not displayed.
- **History & ledger are trigger-maintained:** `INSERT_PartsStockMST`/`UPDATE_PartNumber` keep
  `INV_PARTS_STOCK_MST_HIST`; `UPDATE_PartNumber` also appends to the `INV_PART_QTY_INF` qty ledger
  on any `IN_QTY` change. The qty ledger (`IN_QTY_CHANGE`, `VC_STATUS='U'`) is effectively an audit
  trail of stock movements — a natural source for a "stock movements" view in the rebuild.
- **Purge guard:** `DELETE_RecConfStatPartsStockMstQTY` only adjusts stock when `Purge.PurgeMode=0`
  — i.e. the data-retention purge (`DELETE_AutoPurge`, `[DATAPURGE]`) deletes open-order history
  **without** moving stock. The rebuild's purge job must replicate this "don't touch the balance"
  flag, or it will silently corrupt on-hand quantities.

## 5. UI / UX notes
- **Layout:** header label; a "Searching Key" group box with 4 combos (Supplier = a 2-column
  `TNUMMIColumnComboBox` code+name; Part Name = `csDropDownList`, uppercased; Line; Part Type); a
  `TDBGrid` (`dgRowSelect`, read-only by convention) showing the full parts list; a button panel:
  **Print, Print Excel, Search, Clear, Close** (`Close` = `ModalResult=2/mrCancel`).
- **Grid columns:** all non-`Report*` columns from `SELECT_PartsStockInfo` (30 columns: supplier/
  part/logistics/parts-name/renban/part-type/line/size/kanban/lot/lead-times/ship-days/**QTY**/
  remarks/part-cost). Very wide; the on-hand `QTY` is buried among planning attributes.
- **Search/filter:** **non-functional** (dead code, §4). The combos beyond Supplier→Part do nothing.
- **Printing:** **Print and Print Excel are identical** (same handler, same QuickReport). The report
  is a 7-column portrait listing: Supplier, Parts, Size, KANBAN, **Part Type** (QR field misnamed
  `TireWheel_QRLabel`), **Line Name** (QR field misnamed `CarTruck_QRLabel`), QTY. `PrintIfEmpty=True`.
  Date/time in the page header, page number in the footer, report title via `qrsReportTitle` (blank).
- **What to keep vs modernize:**
  - **Keep:** a fast, sortable, *filterable* on-hand stock list + an export.
  - **Fix/modernize:** (1) make search actually work — server-side filter by supplier / part /
    line / part type / size / kanban, plus an "on-hand below X" / "zero stock" filter (P7);
    (2) make `QTY` a first-class, prominent, right-aligned column with totals; (3) split the wide
    grid into a summary view + drill-down; (4) Print vs Excel should be **two real outputs** (HTML/
    PDF print view + a true CSV/XLSX export), not the same QuickReport preview; (5) drop the dead
    `HoldDetails`/`SetDetailBoxes`/`SearchGrid` scaffolding entirely.

## 6. Target design *(Rails primary)*
- **Models (read side):**
  - `PartStock` → `INV_PARTS_STOCK_MST` (`self.table_name`, `self.primary_key = 'IN_PART_ID'`).
    `belongs_to :site` (D1 — current-site scoping; `acts_as_tenant`/`default_scope` so every query
    is filtered to the current site and on-hand is per-site). `belongs_to :supplier, :logistics,
    :renban_group, :part_type, :tire_size` — all `optional: true` (the columns are nullable and the
    joins are LEFT; FKs are by-convention). Validations: `part_number` presence + `uniqueness`
    **scoped to `:site_id`** (case-insensitive; the unique index becomes composite
    **`(site_id, VC_PART_NUMBER)`**, replacing the global `IX_INV_PARTS_STOCK_MST`). Enum on the
    *supplier's* `inventory_add_point {S,A}` (lives on `Supplier`, P4). `in_qty` is **read-mostly**
    here — never written by this controller, and is **per-site**.
  - `PartStockHistory` → `INV_PARTS_STOCK_MST_HIST` (⚠️ **not** an exact twin: line-name column is
    `VC__LINE_NAME` with a double underscore, and `VC_PARTS_NAME`/`IN_RENBAN_COUNT` are NOT NULL
    here while NULLable on the master — see §2; the rebuild should fix the schema and write an
    explicit column list rather than copy the positional `SELECT *` landmine).
    `PartQtyLedger` → `INV_PART_QTY_INF` (`in_qty_change`, `in_qty`, `vc_status`) — surface as a
    "stock movements" feed.
- **Controllers/routes:** `resources :part_stocks, only: [:index, :show]` for this module (read +
  print). Index = the searchable/paginated stock list; `:show` = one part's detail + movement
  history. A `format: :csv`/`:xlsx` on `index` is the real "Excel" export; a print/PDF view
  (`format: :pdf` via wkhtmltopdf/Prawn, or just a print-stylesheet HTML page) replaces the
  QuickReport. **(This module owns no writes — it shares the `PartStock` model with the
  Parts-Stock-editor module, which owns create/update/delete.)**
- **Views/components:** server-side filtered + sorted + paginated index (replacing the dead
  client-side `Filter`, P7) with filters for supplier/part/line/part-type/size/kanban and an
  on-hand threshold; a prominent right-aligned `QTY` with a column total; a printable report
  partial shared by the HTML print view and the PDF/Excel exporters.
- **Services / the stock-ledger transaction (critical):** the displayed `IN_QTY` is the output of
  **12+ triggers**. In the rebuild, re-home them as **one `StockLedger` service / set of model
  callbacks** invoked inside the same DB transaction as the receiving / shipping / reject /
  stocktaking writes:
  - Receiving posts (open-order insert/update/delete) → `+qty` / delta / `−qty`, **timed by the
    supplier add-point (`S` vs `A`)**, with the empty-trailer / terminated / purge-mode exclusions
    of `DELETE_RecConfStatPartsStockMstQTY`.
  - Shipping posts → `−qty` (restore on delete/undo).
  - Reject posts → `−qty` (restore the qty on reject delete).
  - Stocktaking posts → `+qty` / `−qty` **additive delta** (insert adds, delete subtracts, update
    does both) keyed on `IN_PART_ID` — **not** a set/override to the counted value.
  - Every change writes a `PartQtyLedger` row (mirror `UPDATE_PartNumber`'s `INV_PART_QTY_INF`
    insert) and an `*_HIST` snapshot, and updates `VC_LAST_UPDATE`.
  - A part-number rename must cascade to `INV_ASSY_RATIO_MST` / `INV_FORECAST_DETAIL_INF` string
    references (mirror `UPDATE_PartNumber`/`DELETE_PartNumber`) — or, better, migrate those children
    to reference `IN_PART_ID` so no string cascade is needed (Postgres phase).
- **Reports:** **no `REPORT_*` proc to wrap** — reimplement the printout as an HTML print view +
  PDF + a true CSV/XLSX export over the same query. Stage-1 can wrap `SELECT_PartsStockInfo` for
  exact parity; the report is just a formatting of that result set.

## 7. Migration plan for this module
- [ ] **Stage 1 — read-only parity:** `PartStock.all` (or wrap `SELECT_PartsStockInfo ''` via
      `tiny_tds`) renders the index; columns map 1:1 to the 30 proc aliases; **clustered-PK order**
      (no `ORDER BY` in the proc — match it, then add user sort). Reimplement the printout
      (HTML/PDF/CSV) over the same rows. **Wire up real server-side search** (the legacy never did)
      — document this as an intentional improvement, not a parity break.
- [ ] **Stage 2 — (owned by the Parts-Stock editor module, not here):** writes go through the
      shared `PartStock` model; **the `StockLedger` qty triggers stay in the DB** during parallel
      run so both apps see the same `IN_QTY`. This read module just keeps displaying the balance.
- [ ] **Stage 3 — reimplement (Postgres-ready):** move the 12+ qty triggers into the `StockLedger`
      service/callbacks (transactional), with the add-point timing, empty-trailer/terminated/
      purge-mode exclusions, qty-ledger + history writes, and the part-rename cascade preserved.
      Replace the QuickReport with the HTML/PDF/CSV reporters. Real timestamps replace the
      `yyyymmddHHMMSSff` strings. **Multi-site (D1):** add the `site_id` (NOT NULL) FK → `sites` and
      replace the global `IX_INV_PARTS_STOCK_MST` with a **per-site composite unique index
      `(site_id, VC_PART_NUMBER)`**; the qty triggers/`StockLedger` and the ledger/history tables
      become site-scoped (on-hand is per-site). Add the missing declared FKs for the five `IN_*_ID`
      links. (Legacy single-site DB untouched during the parallel run.)

## 8. Open questions for the user (domain expert)
1. ✅ **RESOLVED (D1): Multi-site scope of stock — per-site.** Per **decision D1
   (`docs/analysis/decisions.md`)**, sites run independently with full data isolation. On-hand
   stock is **per-site**: `INV_PARTS_STOCK_MST` gains a `site_id` (NOT NULL) FK → the new `sites`
   table, the qty triggers/ledger become site-scoped, and `VC_PART_NUMBER` uniqueness becomes
   composite **`(site_id, VC_PART_NUMBER)`** rather than global. (Context retained: today the table
   has no site/plant column — one global parts-stock master, one `IN_QTY` per part; this was the
   single biggest multi-site decision in the inventory area.)
2. **"Print Excel" vs "Print":** today both buttons run the identical QuickReport (Excel is only
   reachable from the preview dialog's export). Is a distinct, true Excel/CSV export wanted, and
   what columns should it carry — the on-screen 30 or the report's 7? (The report also drops Part
   Cost, which the proc returns.)
3. **The dead search:** the entire grid search/filter is commented-out code, so the screen always
   shows every part. Confirm the intended search keys (supplier / part / line / part type / size /
   kanban?) and whether an "on-hand below reorder point" / "zero / negative stock" filter is wanted
   — this is the obvious UX win.
4. **Negative on-hand:** `IN_QTY` is a signed `int` maintained by additive triggers with no
   non-negative guard. Can on-hand legitimately go negative (e.g. shipping before a delayed
   receipt), or should the rebuild flag/clamp it? (The legacy neither prevents nor highlights it.)
5. ✅ **RESOLVED (D4): add-point stays supplier-level** (not per-part) — per decision D4
   (docs/analysis/decisions.md). The rule remains on the supplier (`S` add at shipping / `A` add at
   arrival), so the data-quality fix is at the supplier: the rebuild should make the supplier's
   `VC_INVENTORY_ADD_POINT` a **required, valid `S`/`A`** value (recommended), eliminating the
   blank-value case where stock silently never increments on receipt. Historical blank values should
   be remediated (assigned a correct add-point per supplier) during migration.
6. **Part-cost on a stock screen:** `MO_PART_COST` (money, default 0) rides on the stock master and
   is returned by the proc. Is on-screen valuation (qty × cost = inventory value) desired in the
   modern inventory view? (None is shown today.)
7. **History/ledger retention & purge:** `INV_PARTS_STOCK_MST_HIST` + `INV_PART_QTY_INF` grow on
   every qty change; `DELETE_AutoPurge` + the `Purge.PurgeMode` guard control retention. What
   retention is required, and should the modern app expose the qty ledger as a user-facing "stock
   movements" report?

## 9. Test cases / parity checks
- **List all** → row count equals `SELECT_PartsStockInfo ''`; the 30 returned aliases map 1:1 to the
  grid columns; rows arrive in **clustered-PK (`IN_PART_ID`) order** (no proc `ORDER BY`). A part
  with a nulled `IN_SUPPLIER_ID`/`IN_SIZE_ID`/etc. still appears with the joined code blank
  (LEFT-JOIN parity, P5).
- **Print / Print Excel** → both produce the **identical** 7-column report (Supplier, Parts, Size,
  KANBAN, Part Type, Line Name, QTY) over the loaded rows; previews even with 0 rows
  (`PrintIfEmpty=True`). New app: assert the print view and the Excel/CSV export carry the agreed
  columns (document any divergence from the legacy 7).
- **Search button** → legacy does **nothing** (dead code); the full list stays. New app: the search
  actually filters server-side — document as an **intentional divergence** (a fix), not a parity
  match.
- **Supplier combo change** → repopulates the Part-Name combo with that supplier's part numbers
  (`SELECT_PartsSupplier @code`); blank supplier → all parts. (The only live combo cascade.)
- **Stock-qty invariants (the core parity set — exercised via the writer modules, verified by what
  this screen displays):**
  - **Receiving raises on-hand:** insert an open-order row of qty *N* for a part whose supplier
    add-point = `S` and `VC_STATUS_SUPPLIER_SHIPPING<>''` → `IN_QTY` increases by *N*
    (`INSERT_RecConfStatPartsStockMstQTY`). For add-point = `A`, the increase happens only when an
    arrival/plant-yard/assembler-yard/warehouse status is set. For a **blank** add-point → **no
    change** (neither branch matches) — verify the screen shows the un-incremented balance.
  - **Delete-receipt lowers on-hand** by the same qty (`DELETE_RecConfStatPartsStockMstQTY`) —
    **unless** the row was empty-trailer/terminated (no change) **or** `Purge.PurgeMode=1` (purge
    must not move stock). Assert both exclusions.
  - **Reject lowers, un-reject restores:** insert a reject of qty *N* → `IN_QTY` −*N*; delete that
    reject → `IN_QTY` +*N* back (`DELETE_RejectParts`, keyed on `IN_PART_ID`).
  - **Shipping lowers / un-ship restores** (`InsertPartShipping` / `DeletePartShipping`).
  - **Stocktaking adds/subtracts the counted delta:** insert a stocktaking row of qty *N* →
    `IN_QTY` +*N*; delete it → `IN_QTY` −*N* (keyed on `IN_PART_ID`, same shape as the reject
    family — **not** a set/override to the counted value).
  - Every qty change appends an `INV_PART_QTY_INF` ledger row (`IN_QTY_CHANGE = old − new`,
    `VC_STATUS='U'`) and an `INV_PARTS_STOCK_MST_HIST` snapshot, and bumps `VC_LAST_UPDATE`
    (`UPDATE_PartNumber`). Verify the rebuilt `StockLedger` reproduces all three side-effects in one
    transaction.
- **Part-number rename** (via the Parts-Stock editor) → `UPDATE_PartNumber` propagates the new code
  into `INV_ASSY_RATIO_MST` (4 cols) and `INV_FORECAST_DETAIL_INF` (2 cols) **only for a single-row
  change**; on **delete** those references are blanked (`DELETE_PartNumber`). Confirm the rebuild's
  cascade matches (or that children were migrated to `IN_PART_ID`).
