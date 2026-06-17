# Module Analysis: Parts Stock Master (`INV_PARTS_STOCK_MST`)

**Area:** Inventory / Stock  **Status:** ✅ spec complete  **Analyst:** Claude / 2026-06-05

> **The keystone module.** `INV_PARTS_STOCK_MST` is the central part/stock master: every
> part the site stocks is one row. It carries the **planning attributes** (lead-time and
> ship-day matrices, lot/renban sizing, part cost), the **FKs into the four masters**
> (`IN_SUPPLIER_ID`, `IN_LOGISTICS_ID`, `IN_SIZE_ID`, `IN_RENBAN_ID`, plus `IN_PART_TYPE_ID`),
> **and the live on-hand balance `IN_QTY`.** The crucial fact: **this form owns the table but
> does NOT own the qty invariant.** `IN_QTY` is driven by **twelve qty-adjusting triggers
> owned by other modules** (receiving / rejects / stocktaking / shipping). The
> `UPDATE_PartsStockInfo` proc *can* set `IN_QTY`, but the legacy form does **not** expose it for
> editing — `Quantity_MaskEdit` is `ReadOnly=True` (dfm line 458) and never toggled — so Insert
> sends the default/loaded value and Update merely rewrites the value loaded from the grid row.
> This spec also **resolves the dangling-FK-on-delete open questions** left by
> [`supplier.md`](../master-data/supplier.md), [`logistics.md`](../master-data/logistics.md),
> and [`size.md`](../master-data/size.md) by reading the live trigger bodies.

## 1. Legacy surface
- **Form:** `PartsStockMaster.pas` (711 lines / ~24 KB) + `PartsStockMaster.dfm` (802 lines).
  `TPartsStockMaster_Form`, Caption/header label "Parts Stock Master". Author: Aaron Huge,
  2002-10-25 (2002-12-17 edit: `SearchGrid` switched to partial `LIKE 'ABC%'` matching).
  Registered live in `InventorySystem.dpr` **line 18**
  (`PartsStockMaster in 'PartsStockMaster.pas' {PartsStockMaster_Form}`). There is also a
  **dead/legacy `PartsStockMasterNew.pas`** — NOT in the dpr, ignore it (CLAUDE.md guardrail).
- **Entry point:** **not reached directly from `MainMenu.pas`** — per the form code it is reached
  through the master-maintenance hub (P13/P14), exactly like Supplier/Size/Logistics (this routing
  was read from this form's own code, **not** independently re-verified against `MasterMaint.pas`).
  `MainMenu` opens `MasterMaint_Form`; that hub (`MasterMaint.pas`) is described as having a
  `PartsStockMaster_Button` whose `OnClick` (`PartsStockMaster_ButtonClick`, lines 128-134) does the
  standard `PartsStockMaster_Form := TPartsStockMaster_Form.Create(self);
  PartsStockMaster_Form.Execute; PartsStockMaster_Form.Free;` dance (P14). `Execute` runs `ShowModal`
  and returns `False` only on `mrCancel`.
- **Purpose (one paragraph):** The richest master-detail CRUD screen in the app. A `DBGrid`
  (`PartsStockMaster_DBGrid`) lists every part (30 columns); a large edit panel
  (`PartsStockMaster_Panel`) shows the selected part's details across **~40 controls** —
  supplier, logistics, renban group, part type, assembly line, size, kanban, lot qty, on-hand
  qty, comments, part cost, a 7-cell **lead-time** matrix (base + Mon–Sat) and a 7-cell
  **ship-days** matrix (base + Mon–Sat), a `LotSizeOrders` checkbox, and a renban count.
  Buttons: **Insert, Update, Delete, Search, Clear, Close**. `FormCreate` clears the dataset
  filter, calls `GetInventoryInfo`, binds `Parts_DataSource` to `Inv_DataSet`, populates the
  six combos via `SetCombos`, and clears the panel. Selecting a grid row (`OnMouseUp` /
  `OnKeyUp` / `Parts_DataSourceDataChange` all call `HoldDetails(True)` + `SetDetailBoxes`)
  copies the row into the editors and captures `RecordID` from hidden grid `Fields[29]`
  (the identity PK `IN_PART_ID`).

## 2. Data touched
| Table | Read | Write | Notes |
|-------|:----:|:-----:|-------|
| `INV_PARTS_STOCK_MST` | ✓ | ✓ | The part/stock master — **this module owns it** (direct INSERT/UPDATE/DELETE) |
| `INV_PARTS_STOCK_MST_HIST` |  | ✓* | *Indirect: `INSERT_PartsStockMST` + `UPDATE_PartNumber` triggers copy rows here |
| `INV_PART_QTY_INF` |  | ✓* | *Indirect: `UPDATE_PartNumber` writes a qty-ledger row when `IN_QTY` changes |
| `INV_SUPPLIER_MST` | ✓ | | FK-lookup combo (`VC_SUPPLIER_CODE`/`VC_SUPPLIER_NAME` via `Supplier_NUMMIColumnComboBox`); UPDATE/INSERT resolve **code→`IN_SUPPLIER_ID`** (P3) |
| `INV_LOGISTICS_MST` | ✓ | | FK-lookup combo (`VC_LOGISTICS_NAME`); resolve **name→`IN_LOGISTICS_ID`** (P3); blank → NULL |
| `INV_SIZE_MST` | ✓ | | FK-lookup combo (`VC_SIZE_CODE`); resolve **code→`IN_SIZE_ID`** (P3) |
| `INV_RENBAN_GROUP_MST` | ✓ | | FK-lookup combo (`VC_RENBAN_GROUP_CODE`); resolve **code→`IN_RENBAN_ID`** (P3) |
| `INV_PART_TYPE_MST` | ✓ | | FK-lookup combo (`VC_PART_TYPE`); resolve **label→`IN_PART_TYPE_ID`** (P3) |
| `LINE` (Activity/ALC catalog) | ✓ | | `Line_ComboBox` populated via `SelectSingleFieldALC('LINE','LineName',…)` — **different DB connection** (`ALC_Connection`); stored back as the **string** `VC_LINE_NAME` (no FK) |
| `INV_ASSY_RATIO_MST` | | ✓* | *Indirect: `DELETE_PartNumber`/`UPDATE_PartNumber` blank/propagate the part-number **string** code columns |
| `INV_FORECAST_DETAIL_INF` | | ✓* | *Indirect: same two triggers blank/propagate the part-number string code columns |

### `INV_PARTS_STOCK_MST` columns (authoritative: `DB Schema/Create Inventory.sql`)
| Column | Type | Meaning / notes |
|--------|------|-----------------|
| `IN_PART_ID` | `int IDENTITY(1,1) NOT NULL` PK | Surrogate key. `RecordID` in UI (hidden grid `Fields[29]`). `PK_INV_PARTS_STOCK_MST` CLUSTERED |
| `IN_SUPPLIER_ID` | `int NULL` | FK→`INV_SUPPLIER_MST` (by convention; **no declared FK**). Nulled by `DELETE_SupplierCode` |
| `IN_LOGISTICS_ID` | `int NULL` | FK→`INV_LOGISTICS_MST` (by convention). **NOT** nulled on logistics delete — see §2 trigger note ⚠️ |
| `IN_RENBAN_ID` | `int NULL` | FK→`INV_RENBAN_GROUP_MST` (by convention). Nulled by `DELETE_RenbanGroupCode` |
| `IN_PART_TYPE_ID` | `int NULL` | FK→`INV_PART_TYPE_MST` (by convention). **No delete-unlink trigger** on part-type ⚠️ |
| `IN_SIZE_ID` | `int NULL` | FK→`INV_SIZE_MST` (by convention). Nulled by `DELETE_SizeCode` |
| `VC_PART_NUMBER` | `varchar(12) NOT NULL` | **Business key — DB-UNIQUE via `IX_INV_PARTS_STOCK_MST`.** Form `MaxLength=12`, uppercased. Proc params `@PartNum varchar(12)` match |
| `VC_PARTS_NAME` | `varchar(50) NULL` | Part name. Form `MaxLength=50`, uppercased; proc `@PartsName varchar(50)` matches |
| `IN_RENBAN_COUNT` | `int NULL` | Renban count. ⚠️ **proc param is `@RenbanCount varchar(3)`** and the form stores it as a **string** (`RenbanCount: string`) — implicit varchar→int conversion on insert/update; form mask `'999;1; '`, `MaxLength=3` |
| `VC_KANBAN_NUMBER` | `varchar(5) NULL` | Kanban number. Form `MaxLength=5`; proc `@KanbanNumber varchar(5)` matches |
| `IN_LEADTIME` | `int NULL` | Base lead time. Form mask `'99;1; '`, `MaxLength=2` ⚠️ (UI caps 0–99; DB `int` allows more) |
| `IN_LEADTIME_MONDAY`…`_SATURDAY` | `int NULL` (×6) | Per-weekday lead-time overrides. Each `MaxLength=2` ⚠️ |
| `IN_1LOTQTY` | `int NULL` | One-lot quantity. Form `OneLotQty_MaskEdit` mask `'99999;1; '`, `MaxLength=5` ⚠️ (0–99999; DB `int` larger) |
| `IN_QTY` | `int NULL` | **The on-hand stock balance — the core inventory invariant.** Form `Quantity_MaskEdit` mask `'99999;1; '`, `MaxLength=5` ⚠️, **but `ReadOnly=True` (dfm line 458) and never toggled in the `.pas`** — so the UI does **not** let the user edit on-hand. **Maintained by triggers owned by other modules** (§2, §4); the proc can write it, but the form only ever round-trips the loaded value (§4) |
| `BIT_LOT_SIZE_ORDERS` | `bit NULL` | "Lot size orders" flag. ⚠️ **Stored inverted** — form sets `LotSizeOrders := not LotSizeOrders_CheckBox.Checked` and displays `Checked := not LotSizeOrders`. Disables the renban combo when checked |
| `VC_COMMENTS` | `varchar(300) NULL` | Remarks. Form `MaxLength=300`; proc `@Comments varchar(300)` matches |
| `IN_SHIP_DAYS` | `int NULL` | Base ship days. Form `ShipDays_Edit` is a `TEdit` `MaxLength=3` |
| `IN_SHIP_DAYS_MONDAY`…`_SATURDAY` | `int NULL` (×6) | Per-weekday ship-day overrides. Each `MaxLength=2` ⚠️ |
| `VC_LINE_NAME` | `varchar(10) NULL` | Assembly line **as a string** (from the ALC `LINE` catalog, no FK). **DEFAULT `'TUNDRA'`** (`DF_INV_PARTS_STOCK_MST_VC_LINE_NAME`). Proc `@LineName varchar(10)` matches |
| `MO_PART_COST` | `money NULL` | Part unit cost. **DEFAULT `0`** (`DF_…_MO_PART_COST`). Form `PArtCost_MaskEdit` is a `TcurrEdit` formatted `'$#######0.0000'`; proc `@PartCost money` |
| `VC_LAST_UPDATE` | `varchar(16) NULL` | **Timestamp as `yyyymmddHHMMSSff` string** (P2), set on UPDATE by `UPDATE_PartsStockInfo` **and** rewritten by several qty triggers. Not shown on the form |
| `VC_ADD` | `varchar(16) NOT NULL` | **Timestamp as `yyyymmddHHMMSSff` string** (P2), set on INSERT only. **NOT NULL** (the only audit column in the schema that is) |

**Constraints / indexes (authoritative):**
- `PK_INV_PARTS_STOCK_MST` PRIMARY KEY **CLUSTERED** on `IN_PART_ID`.
- `IX_INV_PARTS_STOCK_MST` **UNIQUE NONCLUSTERED on `VC_PART_NUMBER`** — a **real DB uniqueness
  backstop** on the part number (same posture as the other masters' code indexes). **(Multi-site,
  D1:** becomes composite **UNIQUE `(site_id, VC_PART_NUMBER)`** at the Postgres phase — per-site,
  not global. See §8.7.)
- DEFAULTs: `VC_LINE_NAME → 'TUNDRA'`, `MO_PART_COST → 0`.
- **No declared FOREIGN KEY constraints involving this table.** The entire schema declares only
  **2 FKs** (`INV_ASN_DETAIL_MST→INV_ASN_MST`, `INV_PART_SHIPPING_INF→INV_SHIPPING_INF`); none
  touch `INV_PARTS_STOCK_MST`. **Every FK in/out of this table is by convention only**, enforced
  solely by the delete-unlink triggers on the master tables and the in-proc name→id lookups.

**History table `INV_PARTS_STOCK_MST_HIST`:** same column shape but **all columns nullable**
(except `VC_PART_NUMBER`/`VC_PARTS_NAME`/`IN_RENBAN_COUNT` NOT NULL), **no `IN_PART_ID` identity/PK**
(it is a plain audit log, append-only), and a **typo column `VC__LINE_NAME` (double underscore)**
vs the live table's `VC_LINE_NAME` ⚠️. Written to by `INSERT_PartsStockMST` (full inserted row) and
by `UPDATE_PartNumber` (the *deleted*/pre-image row on every update). The `INSERT … SELECT *`
form means a column add/reorder on the base table will silently desync the HIST insert (schema-order
fragile, P10-adjacent).

**Triggers on `INV_PARTS_STOCK_MST` itself (3 — read live bodies):**
- **`INSERT_PartsStockMST`** (FOR INSERT): `INSERT into INV_PARTS_STOCK_MST_HIST SELECT * from
  inserted`. **Invariant: every new part snapshots its initial row into history.** (This is the
  "Initialize stock row on new part" trigger from the inventory; it does **not** touch `IN_QTY` —
  the new row's qty is whatever the INSERT proc supplied.)
- **`UPDATE_PartNumber`** (FOR UPDATE): (1) copies the **deleted** (pre-update) row into
  `INV_PARTS_STOCK_MST_HIST`; (2) **when `IN_QTY` changed**, inserts a ledger row into
  `INV_PART_QTY_INF (VC_PART_NUMBER, IN_QTY_CHANGE = d.IN_QTY - i.IN_QTY, IN_QTY = i.IN_QTY,
  VC_STATUS='U', VC_ADD = i.VC_LAST_UPDATE)` for each row whose qty moved; (3) **only when exactly
  one row is updated AND `VC_PART_NUMBER` itself changed**, propagates the renamed part number into
  the four `INV_ASSY_RATIO_MST.VC_*_PART_NUMBER*_CODE` columns and the two
  `INV_FORECAST_DETAIL_INF.VC_*_PART_NUMBER_CODE` columns. **Invariant: part edits are audited
  (HIST), qty deltas are ledgered (P14-adjacent qty bookkeeping), and a part-number rename cascades
  to the string-code references that still key on the number.**
- **`DELETE_PartNumber`** (FOR DELETE): logs to `Activity.dbo.InsertAct_Log 'INVENTORY','TRIGGER'`,
  then **blanks** (`SET … = ''`) the four `INV_ASSY_RATIO_MST.VC_*_PART_NUMBER*_CODE` and two
  `INV_FORECAST_DETAIL_INF.VC_*_PART_NUMBER_CODE` columns that match the deleted `VC_PART_NUMBER`.
  **Invariant: deleting a part unlinks its number from assembly-ratio and forecast-detail rows
  (string-code blanking, not row deletion).** ⚠️ It does **NOT** delete or adjust the HIST table, the
  qty ledger, or any `IN_QTY` — and it does **not** clean up `INV_OPEN_ORDER_INF` / `INV_REJECT_INF`
  / `INV_STOCKTAKING_INF` / `INV_PART_SHIPPING_INF` rows that reference the part by number.

**Qty-adjusting triggers owned by OTHER modules that write `INV_PARTS_STOCK_MST.IN_QTY` (12):**
These are the inventory-balance invariant. **Three families key on the int `IN_PART_ID`; the
receiving + shipping families key on the string `VC_PART_NUMBER`.** All write `VC_LAST_UPDATE` too.
| Trigger | On table | Effect on `IN_QTY` | Join key | Gate |
|---------|----------|--------------------|----------|------|
| `INSERT_RecConfStatPartsStockMstQTY` | `INV_OPEN_ORDER_INF` | `+= i.IN_QTY` | `VC_PART_NUMBER` | supplier `VC_INVENTORY_ADD_POINT='S'` & `VC_STATUS_SUPPLIER_SHIPPING<>''`, **or** `='A'` & arrival/yard/warehouse status set. Also copies row to `INV_OPEN_ORDER_INF_HIST` |
| `UPDATE_RecConfStatPartsStockMstQTY` | `INV_OPEN_ORDER_INF` | several `±` legs (qty-change at ship/arrival; ship-status flip; arrival-status flip) | `VC_PART_NUMBER` | add-point S/A; copies *deleted* to `_HIST`. The most complex trigger in the DB |
| `DELETE_RecConfStatPartsStockMstQTY` | `INV_OPEN_ORDER_INF` | `-= d.IN_QTY` | `VC_PART_NUMBER` | add-point S/A; **skipped entirely when `Purge.PurgeMode = 1`** (data-purge bypass) ⚠️ |
| `INSERT_RejectParts` | `INV_REJECT_INF` | `-= i.IN_QTY` | **`IN_PART_ID`** | (none) |
| `UPDATE_RejectParts` | `INV_REJECT_INF` | `+= d.IN_QTY` then `-= i.IN_QTY` | **`IN_PART_ID`** | (none) |
| `DELETE_RejectParts` | `INV_REJECT_INF` | `+= d.IN_QTY` (adds the deleted reject back) | **`IN_PART_ID`** | (none) |
| `INSERT_Stocktaking` | `INV_STOCKTAKING_INF` | `+= i.IN_QTY` | **`IN_PART_ID`** | (none) |
| `UPDATE_Stocktaking` | `INV_STOCKTAKING_INF` | `-= d.IN_QTY` then `+= i.IN_QTY` | **`IN_PART_ID`** | (none) |
| `DELETE_Stocktaking` | `INV_STOCKTAKING_INF` | `-= d.IN_QTY` | **`IN_PART_ID`** | (none) |
| `InsertPartShipping` | `INV_PART_SHIPPING_INF` | `-= i.IN_QTY` | `VC_PART_NUMBER` | (none) |
| `UpdatePartShipping` | `INV_PART_SHIPPING_INF` | `+= d.IN_QTY` then `-= i.IN_QTY` | `VC_PART_NUMBER` | (none) |
| `DeletePartShipping` | `INV_PART_SHIPPING_INF` | `+= d.IN_QTY` (returns the shipped qty) | `VC_PART_NUMBER` | (none) |
| (`DeleteShipDate` on `INV_SHIPPING_INF`) | — | indirect: deletes the matching `INV_PART_SHIPPING_INF` rows → fires `DeletePartShipping` | — | (none) |

**Inbound delete-unlink triggers (resolve the master-spec open questions):**
- `DELETE_SupplierCode` (on `INV_SUPPLIER_MST`): `UPDATE INV_PARTS_STOCK_MST SET IN_SUPPLIER_ID =
  null WHERE a.IN_SUPPLIER_ID = d.IN_SUPPLIER_ID` — **deleting a supplier unlinks (NULLs) its
  parts' `IN_SUPPLIER_ID`** (P5). Parts survive.
- `DELETE_SizeCode` (on `INV_SIZE_MST`): `SET IN_SIZE_ID = null …` — **deleting a size unlinks
  parts' `IN_SIZE_ID`** (P5). Parts survive.
- `DELETE_RenbanGroupCode` (on `INV_RENBAN_GROUP_MST`): `SET IN_RENBAN_ID = null …` — **deleting a
  renban group unlinks parts' `IN_RENBAN_ID`** (P5). Parts survive.
- ⚠️ **`DELETE_LogisticsCode` (on `INV_LOGISTICS_MST`) does NOT touch the parts table.** It runs
  `UPDATE INV_SUPPLIER_MST SET IN_LOGISTICS_ID = null …` — it nulls the **supplier's** logistics
  FK, **not** the part's. **Therefore `INV_PARTS_STOCK_MST.IN_LOGISTICS_ID` is left DANGLING** when
  a logistics row is deleted: a part can point at a non-existent `IN_LOGISTICS_ID`. The
  `SELECT_PartsStockInfo` LEFT JOIN tolerates it (shows a blank logistics name), but it is a real
  referential-integrity gap. **This resolves the logistics-spec open question:** parts are NOT
  cleaned up on logistics delete (only suppliers are). See §8.
- **No trigger unlinks `IN_PART_TYPE_ID`** on part-type delete (there is no `DELETE_PartTypeCode`
  trigger), so a part-type delete would also leave a dangling part-type FK on parts. Also resolves
  by absence.

## 3. Stored procedures used
(Read with `sql.sh proc NAME`. The procs are the behavioral spec.)

| Proc | Op | Business rule (from body) |
|------|----|---------------------------|
| `SELECT_PartsStockInfo;1 @PartNum varchar(12) = ''` | SELECT | Single param. `@PartNum=''` → **all** parts (legacy single-site returns every row unfiltered; under D1 this must be **scoped to the current `site_id`** — §8.7); else the one `WHERE VC_PART_NUMBER=@PartNum`. Both branches **LEFT OUTER JOIN** the five masters (Supplier, Logistics, Renban, PartType, Size) so a NULL/dangling FK still returns the part with a blank label. Returns **30 UI-aliased columns** mapped 1:1 to grid `Fields[0..29]`: `Supplier Code, Parts Code, Logistics Name, Parts Name, Renban Group, Part Type, Line Name, Size Code, KANBAN, 1 Lot QTY, Lot Size Orders, Lead Time, Renban Count, Ship Days, Lead Time Mon..Sat, Ship Days Mon..Sat, QTY, Remarks, Part Cost, RecordID(=IN_PART_ID)`. **No name→id resolution** (ids are only returned). No `ORDER BY`. `VC_ADD`/`VC_LAST_UPDATE` not selected. |
| `INSERT_PartsStockInfo;1` (29 params) | INSERT | Computes `VC_ADD` as a `yyyymmddHHMMSSff` string (P2). **Resolves five labels→ids inside the proc** (P3): `@SupCode→IN_SUPPLIER_ID`, `@LogisticsName→IN_LOGISTICS_ID`, `@RenbanCode→IN_RENBAN_ID`, `@PartType→IN_PART_TYPE_ID`, `@SizeCode→IN_SIZE_ID` (unmatched label → NULL id). Then **explicit-column INSERT** (30 cols, names listed — **NOT P10**). `IN_PART_ID` identity not supplied, **not returned**. **No uniqueness check inside the proc** — the app two-step does the dup guard (see Call-mechanism). `IN_QTY` is set directly to `@QTY` — but `@QTY` carries the **default/loaded** value because the form's `Quantity_MaskEdit` is `ReadOnly` (the user can't type a new on-hand). ⚠️ **Param-width mismatches vs table:** `@SizeCode varchar(50)` (table `IN_SIZE_ID` is int, resolved by code; the code col `INV_SIZE_MST.VC_SIZE_CODE` is varchar(6), and the form caps the combo at 6 — proc is wider, harmless); `@RenbanCount varchar(3)` written into `IN_RENBAN_COUNT int`; `@LogisticsName varchar(25)` (matches `VC_LOGISTICS_NAME`). |
| `UPDATE_PartsStockInfo;1` (30 params, +`@PartID int`) | UPDATE | Same five label→id resolutions (P3); sets `VC_LAST_UPDATE` (P2). Updates **by `IN_PART_ID = @PartID`** (the shared `RecordID`, P9). **Rewrites `VC_PART_NUMBER`** (the business key is editable) and **`SET IN_QTY = @QTY`** — but the form feeds `@QTY` the value `SetDetailBoxes` loaded from the selected grid row (the `Quantity_MaskEdit` is `ReadOnly`), so in normal UI operation the written value **equals** the loaded value and the on-hand is unchanged (see §4). Does not touch `VC_ADD`. **No app/proc uniqueness re-check** — a rename collision is caught only by `IX_INV_PARTS_STOCK_MST` (raw SQL error). Note the UPDATE fires `UPDATE_PartNumber`: HIST snapshot + qty-ledger row if `IN_QTY` moved + rename cascade. |
| `DELETE_PartsStockInfo;1 @PartID integer` | DELETE | Hard-`DELETE INV_PARTS_STOCK_MST WHERE IN_PART_ID = @PartID`. Single surrogate param. No soft-delete, no in-use / RI check inside the proc — relies on the `DELETE_PartNumber` trigger to blank the assy-ratio/forecast string-code references. ⚠️ **Does not clean up transactional children** (open orders, rejects, stocktaking, shipping rows referencing the part by number) — see §8. |

**Related parts-stock procs NOT called by this form (owned/used elsewhere — listed for the rebuild):**
| Proc | Op | Where / why it matters |
|------|----|------------------------|
| `SELECT_PartsStockInfoOrder @LineName,@PartType,@SortType` | SELECT | Order-screen part list (`SELECT *`, inner-joins part type/size/supplier). |
| `SELECT_PartsStockLogistics`, `SELECT_PartsStockRenban` | SELECT | Logistics/renban-scoped part reads. |
| **`UPDATE_PartsStockInfoCount @PartNumber varchar(12), @QTY int`** | UPDATE | `SET IN_QTY = IN_QTY - @QTY … WHERE VC_PART_NUMBER = @PartNumber`, sets `VC_LAST_UPDATE`. **A second writer of `IN_QTY` keyed by number** (line-pull / count flow) — confirms `IN_QTY` is mutated from multiple places. |
| `UPDATE_PartsStockRenban` | UPDATE | Renban-count maintenance. |
| `SELECT_PartsDailyLinePull(Count)`, `REPORT_LogicalInventory` | SELECT/REPORT | Daily line-pull and logical-inventory reporting off this table. |
| `INSERT_INVInfo` (`InsertINVInfo`) | INSERT | Writes `INV_INV_MST` (inventory master) — a separate table, not this form. |

### Call mechanism (legacy)
`DataModule.pas` methods drive the shared ADO objects (P6):
- **`GetInventoryInfo`** (lines 1305-1348) uses `Inv_DataSet.Open` with `CommandText :=
  'dbo.SELECT_PartsStockInfo;1'`, `@PartNum := ''` (all parts), times the call, logs
  `LogActLog('GET PARTS','SELECTED all parts',1)`. (Form's grid binds to `Inv_DataSet`.)
- **`InsertPartsStockInfo`** (lines 1350-1473) is **two-step (P1)**: STEP 1 sets
  `ProcedureName := 'dbo.SELECT_PartsStockInfo;1'` with `@PartNum := fPartNum`; only `If
  RecordCount = 0` does it proceed to `INSERT_PartsStockInfo` with all 29 params. ✅ **Unlike
  Size, this dup-check targets the CORRECT proc** (`SELECT_PartsStockInfo`, not a wrong one) — so
  the app-side part-number duplicate guard actually works here, redundant with the DB unique index.
  On `RecordCount > 0` it sets `fDescription := '… (DUPLICATE)'` and returns `False` (the form then
  shows "Unable to INSERT with Parts Code …"). Logs `LogActLog('INS PART', …, 1)`.
- **`UpdatePartsStockInfo`** (1475-1575): 30 params + `@PartID := fRecordID`; logs
  `LogActLog('UPDATE PRT', 'UPDATE S:…P:…Q:…', 1)`. No app uniqueness re-check.
- **`DeletePartsStockInfo`** (1577-1618): only `@PartID := fRecordID`; logs
  `LogActLog('DELETE PRT', …, 1)`.
- All four share the **P8 retry-up-to-3-times-via-recursion** harness (`fErrorCount < 3` →
  recursive **self**-call), `finally` doing `Inv_StoredProc.Close; fErrorCount := 0`, and on hard
  failure `ShowMessage` + `LogActLog('ERROR',…)` (no distinct re-raise here — they swallow after 3
  tries, unlike Supplier/Size which re-raise `EDatabaseError`). ✅ **None of the four PartsStock
  methods appear in the P12 wrong-target-retry register** ([`datamodule-retry-target-bugs.md`])
  — each retry correctly re-invokes its own enclosing method (verified at lines 1339, 1464, 1566,
  1609). So this module is **not** a P12 hazard source. (The `DeleteSupplierInfo` magnet bug,
  however, *cascades into this table* via `DELETE_SupplierCode` — see that register §"magnet".)

**DataModule properties used** (all the shared, generic ones): `SupplierCode/PartNum/LogisticsName/
PartName/RenbanCode/RenbanCount(string)/PartType/LineName/SizeCode/Kanban/LotQty(int)/
Quantity(int→fQTY)/Comments/LotSizeOrders(bool)/LeadTime+6 weekday ints/ShipDays+6 weekday ints/
PartCost(double)` and the record key is the **shared, generic `RecordID`** (line 337) — reused by
Shipping/Invoice/ASN/Supplier/Size/Logistics, so a stale `RecordID` from another screen is a real
latent cross-module hazard (P9): Update/Delete here key off it with no guard.

## 4. Business rules & edge cases
- **Identity is `VC_PART_NUMBER`** (`MaxLength=12`, uppercased), backed by a **real DB UNIQUE
  index** `IX_INV_PARTS_STOCK_MST`. The surrogate `IN_PART_ID` is the actual key for update/delete
  and the inbound FK from every other table. **(Multi-site, per decision D1:** the part row gains a
  `site_id` (NOT NULL) FK and this uniqueness becomes composite `(site_id, VC_PART_NUMBER)` — the
  part number is unique per-site, not globally. See §8.7.)
- **`IN_QTY` (on-hand) is trigger-maintained, and the legacy form does NOT expose it — the headline
  rule.** The `UPDATE_PartsStockInfo`/`INSERT_PartsStockInfo` procs *do* write `IN_QTY` (`SET IN_QTY =
  @QTY`), but the form's `Quantity_MaskEdit` is **`ReadOnly=True` (dfm line 458) and is never set
  back to `False` anywhere in the `.pas`**. `SetDetailBoxes` reloads the box from the selected grid
  row, and `HoldDetails(False)` passes that same loaded value back as `fQTY` — so **the value Update
  writes equals the value it loaded** and on-hand is left unchanged through this UI. On Insert the
  proc gets the default/loaded value, not a hand-keyed one. The "clobber the trigger balance" risk is
  therefore **effectively closed by `ReadOnly`** in normal operation; it would only matter if the
  proc were called with a different `@QTY` from outside the form. The day-to-day balance is moved by
  the 12 qty-triggers (§2): **receiving** adds (open-order posts, gated by the supplier's add-point
  S=at-supplier-shipping vs A=at-arrival), **shipping** subtracts, **rejects** subtract (and add back
  on delete), **stocktaking** adds/reverses. (Note: should `IN_QTY` ever change on an update,
  `UPDATE_PartNumber` ledgers it into `INV_PART_QTY_INF` with status `'U'`, change = old−new — the
  audit path exists even though the form doesn't drive it.) **The rebuild must NOT model `IN_QTY` as a
  freely-editable field**; it must be a balance maintained by the receiving/shipping/reject/
  stocktaking service transactions, with any on-hand correction re-cast as an explicit "adjust /
  correct on-hand" action that writes a ledger entry — never a silent absolute overwrite.
- **Add-point semantics couple this table to the supplier master.** Whether a receiving event
  moves qty at supplier-shipping (`'S'`) or at arrival (`'A'`) is read from
  `INV_SUPPLIER_MST.VC_INVENTORY_ADD_POINT` inside the receiving triggers. So a part's qty behavior
  depends on its supplier's add-point — and a part with `IN_SUPPLIER_ID = NULL` (e.g. after
  `DELETE_SupplierCode`) **silently stops receiving qty updates** (the INNER JOIN to supplier in
  the receiving triggers drops it). Edge case worth a parity test.
- **Five FK labels resolved name/code→id inside the procs (P3).** Blank label → NULL id. Logistics
  blank is explicitly mapped (`if Logistics_ComboBox.Text = ' ' then LogisticsName := ''`), same for
  renban (`if RenbanCode_ComboBox.Text=' '`). Part type / size / line are stored from whatever the
  combo text is.
- **`VC_LINE_NAME` is a string, not an FK**, sourced from the **ALC/Activity `LINE` catalog over a
  different DB connection** and defaulting to `'TUNDRA'`. No referential integrity; a line rename in
  ALC would orphan the stored string. Cross-DB concern (§8.6); and per **decision D1** the part row is
  now per-site (`site_id` FK), so the assembly line resolves within the current site's scope.
- **`BIT_LOT_SIZE_ORDERS` is stored inverted** relative to the checkbox (`not Checked`), and toggling
  the checkbox enables/disables (and clears) the renban combo. Preserve the *meaning*, fix the
  inversion in the rebuild (store the boolean as displayed).
- **Editable business key.** `UPDATE_PartsStockInfo` rewrites `VC_PART_NUMBER`. Parts' own surrogate
  is stable, but a rename **cascades** via `UPDATE_PartNumber` to assy-ratio + forecast-detail
  string-code columns — **but only when exactly one row is updated** (`if @numrows = 1`). And the
  transactional children (open orders, rejects, stocktaking, shipping) key on the **old**
  `VC_PART_NUMBER` and are **not** cascaded — so a rename silently detaches in-flight qty events from
  the part. Significant fragility (§8).
- **Validation is thin.** `Validate` only checks that every numeric field parses as an integer
  (`TryStrToInt`); the **part-number length check is commented out** (the `length < 12` guard is
  disabled in source, lines 288-297). `TextChange` forces empty mask edits to `'0'`; `MaskEditExit`
  strips embedded spaces. Note `PartsNum_Edit` has a **literal default `Text='000000000000'` (twelve
  zeros, NOT blank)** in the dfm — so the "blank part number" edge case below is really "the default
  is twelve zeros." A truly empty number is only reached if the user clears the field; with the
  length check disabled, a **blank (or all-zeros) part number can be submitted** (the DB unique index
  allows a single `''`/`'000000000000'`; the app dup-check would then block a *second* identical one).
  Part cost parses via a `$`-strip + `TryStrToFloat`.
- **UI numeric caps below DB capacity (⚠️):** `IN_QTY` and `IN_1LOTQTY` capped at 99999 (5-digit
  mask), lead-time cells at 99 (2-digit), ship-day cells at 99, renban count at 999 — all narrower
  than the `int` columns; `UPDATE_PartsStockInfoCount` and the triggers can push `IN_QTY` past 99999
  even though this form can't enter it.
- **Delete is hard on the part, soft on string-code references.** `DELETE_PartNumber` blanks the
  assy-ratio/forecast number references; the part row and its HIST/qty-ledger rows are untouched.
  Transactional children are left dangling (§8).
- **Purge-mode bypass:** `DELETE_RecConfStatPartsStockMstQTY` checks `Purge.PurgeMode` and **skips**
  the qty decrement when purge mode is on — so bulk data-purge deletes of open orders don't wrongly
  drain on-hand. The rebuild's purge path must replicate this "don't re-balance during purge" rule.
- **Timestamps are `yyyymmddHHMMSSff` strings (P2):** `VC_ADD` (NOT NULL) on insert,
  `VC_LAST_UPDATE` on update and on every qty-trigger touch. Byte-identical recipe to the other
  masters (`CONVERT(char(8),…,112)` + four `SUBSTRING(…,114,2)` slices).

## 5. UI / UX notes
- Grid + large detail-panel. Selecting a grid row syncs ~40 editors and captures `RecordID` from
  hidden `Fields[29]`. Two 7-cell matrices (lead time, ship days) dominate the panel.
- **Search is client-side over the loaded grid (P7) with partial `LIKE`** (the 2002-12-17 edit):
  `SearchGrid` sets `Inv_DataSet.Filter` to `[Supplier Code] LIKE '<code>'` and/or `[Parts Code]
  LIKE '%<num>%'` and toggles `Filtered`. (Note the supplier filter is **exact** `LIKE` without
  wildcards; the part filter is `%…%` contains.) The `fAssyLine`/`fWheelTire` branch is effectively
  dead (those params arrive uninitialized from `Search_ButtonClick`). On no match:
  "No matches were found for your query."
- **Six combos:** Supplier (two-column code+name `TNUMMIColumnComboBox`), Logistics, Size, Renban
  group, Part type, Line (ALC). **`SizeCode_ComboBox` and `RenbanCode_ComboBox` both have
  `CharCase=ecUpperCase` and `MaxLength=6`** (matching the 6-char `VC_SIZE_CODE` /
  `VC_RENBAN_GROUP_CODE` columns) — the combo text is uppercased and capped before it reaches the
  proc's label→id lookup. **`VendorShare_SpeedButton` (`Visible=False`) and `VendorShare_Edit`
  are dead UI** — declared but never wired (no `OnClick`, no read/write in the `.pas`).
- **`LotSizeOrders` checkbox** gates the renban combo (disables + clears when checked).
- **Modernize:** standard index/list + new/edit form; **server-side search/sort/pagination** (P7,
  replacing the in-memory `Filter`); FK combos → real select inputs posting ids; **keep on-hand
  read-only on this master screen** (the legacy `Quantity_MaskEdit` is `ReadOnly`) and provide any
  correction as a separate explicit "adjust on-hand" action that writes a ledger row; drop the dead
  VendorShare controls; add real part-number presence validation (the legacy length check is
  disabled, and the default is `'000000000000'`); collapse the 14 weekday cells into a clearer
  per-day grid; un-invert the lot-size flag.

## 6. Target design  *(Rails primary)*
- **Model:** `PartStock` → `self.table_name = 'INV_PARTS_STOCK_MST'`, `self.primary_key =
  'IN_PART_ID'`.
  - Associations (all `optional: true`, FK by convention — there are no DB FKs):
    `belongs_to :supplier, foreign_key: 'IN_SUPPLIER_ID'`,
    `belongs_to :logistics, foreign_key: 'IN_LOGISTICS_ID'`,
    `belongs_to :tire_size, foreign_key: 'IN_SIZE_ID'`,
    `belongs_to :renban_group, foreign_key: 'IN_RENBAN_ID'`,
    `belongs_to :part_type, foreign_key: 'IN_PART_TYPE_ID'`.
    `has_many :part_qty_entries` (`INV_PART_QTY_INF`), `has_many :stock_history`
    (`INV_PARTS_STOCK_MST_HIST`). Transactional children (open orders, rejects, stocktaking,
    part-shipping) associate by `VC_PART_NUMBER` (number) — model carefully (§8).
    **`belongs_to :site` (D1, §8.7)** — every part row belongs to one site (`site_id` NOT NULL),
    with enforced **current-site scoping** (`default_scope`/`acts_as_tenant`); auth binds the user
    to a site. On-hand is per-site.
  - Validations: `part_number` **presence** (the legacy length check is disabled — add a real one),
    `length: {maximum: 12}`, **uniqueness scoped to `site_id`** (case-insensitive; backed by the
    composite unique index `(site_id, VC_PART_NUMBER)` that replaces `IX_INV_PARTS_STOCK_MST` — D1).
  - **`in_qty` must NOT be a plain mass-assignable attribute.** Model on-hand as a balance changed
    only through service transactions (receiving/shipping/reject/stocktaking) and an explicit
    manual-adjustment action that writes an `INV_PART_QTY_INF` ledger row (mirrors `UPDATE_PartNumber`).
  - Inverted `lot_size_orders`: store as displayed; map to the legacy inverted `bit` only during
    parallel run.
  - Timestamps: write `vc_add` on create (NOT NULL) and `vc_last_update` on update/qty-change as
    `yyyymmddHHMMSSff` strings during parallel run (P2); normalize at the Postgres phase.
- **The qty triggers → services/callbacks (the core re-homing):** re-implement the 12 qty-adjusting
  triggers as **explicit, atomic service transactions** keyed on `IN_PART_ID` (standardize off the
  string `VC_PART_NUMBER`, see §8): `ReceivingService` (open-order post: +/− gated by the supplier's
  add-point enum S/A), `ShippingService` (−), `RejectService` (− on create, + on reverse),
  `StocktakingService` (count reconcile). Each writes the `INV_PART_QTY_INF` ledger row. Preserve the
  **purge-mode bypass** as a "skip re-balance when purging" guard.
- **Master deletes → associations/callbacks:** `Supplier`/`TireSize`/`RenbanGroup`
  `has_many :part_stocks, dependent: :nullify` (mirrors `DELETE_SupplierCode`/`SizeCode`/
  `RenbanGroupCode`). **Logistics must NOT nullify part FKs** (legacy doesn't) — only the supplier's;
  decide in §8 whether to *fix* the dangling-part-logistics gap. Part-type has no unlink trigger —
  same decision.
- **`DELETE_PartNumber` / `UPDATE_PartNumber` string-code cascade → callbacks** on AssyRatio /
  ForecastDetail (or, preferred, migrate those to int FKs so the cascade disappears).
- **Controllers/routes:** RESTful `resources :part_stocks` + a member action for on-hand adjustment.
- **Views:** index (server-side searchable/paginated, P7) + new/edit form with FK selects; separate
  on-hand adjustment view.
- **Services:** the qty services above (Rails); the EDI/receiving math may live in the Python
  service. **Stage-1 option:** wrap the four existing procs via `tiny_tds` for parity (the working
  app-side dup-check + the live triggers keep balance correct), then reimplement in stage 3.
- **Reports:** `REPORT_LogicalInventory`, `SELECT_PartsDailyLinePull(Count)` (owned by inventory
  reporting, not this CRUD form).

## 7. Migration plan for this module
- [ ] **Stage 1 — read-only parity:** `PartStock` list (or wrap `SELECT_PartsStockInfo ''`); the 30
      returned columns map 1:1 to grid `Fields[0..29]`; LEFT-JOIN masters so dangling FKs render
      blank. Server-side search replaces the in-memory `LIKE` filter (P7).
- [ ] **Stage 2 — writes via wrapped procs:** call `INSERT/UPDATE/DELETE_PartsStockInfo` through
      `tiny_tds`, **keeping all live triggers active** so the qty balance, HIST snapshots, qty
      ledger, FK-unlinks, and rename cascades stay correct. Preserve the working app dup-check.
      **Keep on-hand read-only on the master form** (as the legacy `Quantity_MaskEdit` is) even in
      stage 2 — do not introduce a free qty field; route any correction through a separate explicit
      adjustment action that writes a ledger row.
- [ ] **Stage 3 — reimplement (Postgres-ready):** **add the `site_id` (NOT NULL) FK → `sites` and
      rebuild the part-number unique index as composite `(site_id, VC_PART_NUMBER)` — per-site, not
      global (D1, §8.7); the legacy single-site DB stays untouched during the parallel run, so this
      lands only in the Postgres phase.** ActiveRecord validations (presence + uniqueness on part
      number, **scoped to `site_id`**) backed by that composite index; the 12 qty triggers become the
      four qty services + ledger writes keyed on `IN_PART_ID`; `dependent: :nullify` replaces the
      master-delete unlink triggers (decide the logistics/part-type dangling-FK fix, §8); real FKs
      replace by-convention links; the part-number rename cascade either becomes a callback or
      disappears via int-FK migration of AssyRatio/ForecastDetail; real timestamps replace the string
      audit columns; decide transactional-child cleanup on part delete (§8).

## 8. Open questions for the user (domain expert)
1. **Manual `IN_QTY` override semantics.** The legacy master form does **not** expose on-hand for
   editing (`Quantity_MaskEdit` is `ReadOnly`), so today only `UPDATE_PartsStockInfoCount` and the
   qty-triggers move the balance from outside this screen; if the proc *were* called with a different
   `@QTY` it would overwrite the trigger-maintained balance and the change would be ledgered as a
   `'U'` row in `INV_PART_QTY_INF`. In the rebuild, should the master screen *ever* allow setting
   on-hand, or should every qty change go through a receiving/shipping/reject/stocktaking/adjustment
   transaction? (Recommend: keep it read-only here; explicit adjustment only.)
2. ✅ **RESOLVED (D3): block the delete (RESTRICT) — don't nullify, don't dangle.** Per decision D3
   (docs/analysis/decisions.md), deleting a **logistics** (or **part-type**) record that is still
   referenced by any part is **blocked**. The rebuild does not replicate the legacy
   `DELETE_LogisticsCode` behavior (nulls only `INV_SUPPLIER_MST.IN_LOGISTICS_ID`, leaving
   `INV_PARTS_STOCK_MST.IN_LOGISTICS_ID` dangling) and does not nullify part links; part-type, which
   has no unlink trigger at all today, is likewise blocked while referenced. Removal of an in-use
   master is via the future **archival** capability, not delete.
3. ✅ **RESOLVED (D3): block deleting an in-use part (RESTRICT).** Per decision D3
   (docs/analysis/decisions.md), deleting a part that is still referenced by any open order, reject,
   stocktaking, part-shipping, assy-ratio, or forecast row is **blocked**. This ends the legacy
   `DELETE_PartsStockInfo` hazard (deletes the part, blanks only the assy-ratio/forecast string codes,
   leaves transactional children dangling by number, and strands the qty triggers). To retire an
   in-use part, use the future **archival** capability (soft-delete / hide), not delete.
4. ✅ **RESOLVED (D2): standardize **all** consumers on the surrogate `IN_PART_ID`; `VC_PART_NUMBER`
   is an editable, non-key attribute.** Per decision D2 (docs/analysis/decisions.md), the surrogate
   id is the sole key. The fragile legacy `UPDATE_PartNumber` string-cascade (which only cascaded to
   assy-ratio/forecast string codes, and only when exactly one row matched, and never to the
   transactional children keyed on the old number) **goes away**: every consumer — assy-ratio,
   forecast, open orders, rejects, stocktaking, part-shipping — links by `IN_PART_ID`. Renaming a
   part number is then **allowed** (extremely rare) and **safe with no cascade**, treating the number
   as a display label. The number stays unique **per-site** (composite `(site_id, VC_PART_NUMBER)`,
   per D1) as an attribute constraint, not a key.
5. ✅ **RESOLVED (D4): add-point is supplier-level only — keep the coupling, do NOT move it to the
   part.** Per decision D4 (docs/analysis/decisions.md), `VC_INVENTORY_ADD_POINT` stays on the
   supplier; a part's receiving-qty behavior (add at shipping `S` vs arrival `A`) is read from its
   supplier, and that coupling is **intended**. Add-point is **not** moved onto the part. To remove
   the silent-no-add hazard, the rebuild should require a part to have a supplier and require the
   supplier's add-point to be a valid `S`/`A` (recommended enforcement).
6. **`VC_LINE_NAME` from the ALC `LINE` catalog (cross-DB, string, default `'TUNDRA'`).** Should the
   assembly line become a real FK/lookup, and how does the cross-database `LINE` catalog map in a
   multi-site web app (is `LINE` per-site)?
7. ✅ **RESOLVED (D1): per-site — the part catalog is NOT shared.** Per decision D1
   ([`docs/analysis/decisions.md`](../decisions.md)), sites run independently with full data
   isolation: `INV_PARTS_STOCK_MST` gains a **`site_id` (NOT NULL) FK** → the new `sites` table,
   every read/write is scoped to the current site, on-hand stock is per-site, and `VC_PART_NUMBER`
   uniqueness becomes **composite `(site_id, VC_PART_NUMBER)`** (the `IX_INV_PARTS_STOCK_MST` unique
   index is rebuilt on the pair) — NOT global. Resolved consistently with Supplier/Size/Logistics §8.
   The `HIST` table's typo column `VC__LINE_NAME` and the schema-order-fragile `INSERT … SELECT *`
   HIST writes should also be cleaned up at the Postgres phase.

## 9. Test cases / parity checks
- **List all** → row count matches `SELECT_PartsStockInfo ''`; the 30 columns map 1:1 to grid
  `Fields[0..29]`; a part whose `IN_LOGISTICS_ID`/`IN_SUPPLIER_ID` was nulled (or dangling) still
  appears with a **blank** logistics/supplier label (LEFT JOIN parity).
- **Insert a new part** → row added with `VC_ADD` as a 16-char `yyyymmddHHMMSSff` string, five FK
  ids resolved from the combo labels (unmatched → NULL), `IN_QTY = @QTY` where `@QTY` is the
  default/loaded value (the form's `Quantity_MaskEdit` is `ReadOnly`, so a new part's on-hand is not
  hand-keyed), `BIT_LOT_SIZE_ORDERS` stored **inverted**, and a **HIST row** created by
  `INSERT_PartsStockMST`. `IN_PART_ID` identity assigned (legacy does not echo it back — verify the
  new app persists/returns it).
- **Insert an existing part number** → app dup-check (`SELECT_PartsStockInfo @PartNum`) blocks it,
  "Unable to INSERT with Parts Code …", no row added (this guard **works** here, unlike Size).
- **Insert a blank / all-zeros part number** → the `PartsNum_Edit` default is `'000000000000'` (not
  blank); with the length check disabled the legacy *allows* either a literal `''` (if the user
  clears it) or the default twelve zeros (unique index permits a single instance of each). New app:
  reject via presence/format validation (document the divergence).
- **Update a part through this form** → `IN_QTY` is **unchanged** because `Quantity_MaskEdit` is
  `ReadOnly` and `@QTY` carries the value loaded from the grid row (`IN_QTY_CHANGE = 0`, so
  `UPDATE_PartNumber` writes no qty-ledger row, only a HIST pre-image row); `VC_LAST_UPDATE` set.
  (Parity check: confirm the legacy form cannot move on-hand.) **Qty-ledger path (proc-level, not via
  the form):** if `UPDATE_PartsStockInfo` is invoked with a `@QTY` differing from the stored value,
  `UPDATE_PartNumber` writes an `INV_PART_QTY_INF` row (`IN_QTY_CHANGE = old−new`, `VC_STATUS='U'`)
  plus the HIST pre-image — assert this for the new app's explicit adjustment action.
- **Receiving qty invariant (S add-point):** post an open order (`INV_OPEN_ORDER_INF` insert with
  `VC_STATUS_SUPPLIER_SHIPPING<>''`) for a part whose supplier has `VC_INVENTORY_ADD_POINT='S'` →
  `INV_PARTS_STOCK_MST.IN_QTY` increases by the order qty (`INSERT_RecConfStatPartsStockMstQTY`).
  Delete that open order (purge mode **off**) → `IN_QTY` decreases by the same (`DELETE_…`); with
  purge mode **on**, `IN_QTY` is **unchanged**.
- **Receiving qty invariant (A add-point):** same with `VC_INVENTORY_ADD_POINT='A'` and an arrival/
  yard/warehouse status set; assert add at arrival, not at supplier shipping.
- **Reject invariant:** insert a reject (`INV_REJECT_INF`, keyed `IN_PART_ID`) → `IN_QTY -= reject
  qty`; delete the reject → `IN_QTY += reject qty` (`DELETE_RejectParts` restores it).
- **Stocktaking invariant:** insert a count → `IN_QTY += count`; update → reverse old, apply new;
  delete → `IN_QTY -= count`.
- **Shipping invariant:** insert a part-shipping row → `IN_QTY -= shipped qty`
  (`InsertPartShipping`, keyed by number); delete it → `IN_QTY += shipped qty`; delete the parent
  `INV_SHIPPING_INF` ship-date → cascades to delete the part-shipping rows (`DeleteShipDate`) →
  `IN_QTY` returns.
- **Master-delete unlink parity:** delete the part's supplier → its `IN_SUPPLIER_ID` becomes NULL,
  part survives (`DELETE_SupplierCode`); same for size (`IN_SIZE_ID`) and renban (`IN_RENBAN_ID`).
  Delete the part's **logistics** → the part's `IN_LOGISTICS_ID` is **left dangling** (NOT nulled);
  only the supplier's logistics FK is nulled (`DELETE_LogisticsCode`) — confirm this gap and the
  new app's chosen behavior (§8.2).
- **Part-number rename cascade:** rename one part → assy-ratio/forecast `VC_*_PART_NUMBER*_CODE`
  references update to the new number (`UPDATE_PartNumber`, single-row branch); confirm a multi-row
  update does **not** cascade, and that transactional children keep the old number.
- **Delete a part** → row gone; assy-ratio/forecast string-code references blanked
  (`DELETE_PartNumber`); HIST/qty-ledger rows untouched; transactional children left dangling
  (assert the new app's chosen restrict-vs-allow behavior, §8.3).
