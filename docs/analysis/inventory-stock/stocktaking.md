# Module Analysis: Stocktaking (Physical-Count Adjustment)

**Area:** Inventory / Stock  **Status:** ✅ spec complete  **Analyst:** Claude / 2026-06-05

> First module in the **Inventory / Stock** functional area. "Stocktaking" here is **not**
> an absolute physical-count entry — it is a **signed delta adjustment** to a part's on-hand
> quantity (`INV_PARTS_STOCK_MST.IN_QTY`). Each saved row carries an `IN_QTY` *delta* (positive
> or negative) plus a free-text reason; three triggers on `INV_STOCKTAKING_INF` apply/reverse
> that delta against the stock master. This is the **core inventory invariant** the rebuild must
> re-home as a model callback / service transaction. A **second, headless writer**
> (`InsertAutoScrap`, called by `DailyBuildTotal.pas`) posts negative stocktaking rows for scrap —
> so the stock-qty effect of this table is reached from *two* screens, not one.

## 1. Legacy surface
- **Form:** `Stocktaking.pas` (~362 lines) + `Stocktaking.dfm` (`TStocktaking_Form`, **window
  Caption = "Stocktaking Adjustment"** — note "Adjustment", confirming the delta semantics).
  Author header: Aaron Huge, 10/25/2002. Registered live in **`InventorySystem.dpr` line 12**
  (`Stocktaking in 'Stocktaking.pas' {Stocktaking_Form}`).
- **Entry point:** reached **directly from `MainMenu.pas`** (unlike the masters, which go through
  the `MasterMaint` hub). `MainMenu.pas` `Stocktaking_ButtonClick` (lines 351-358) runs the
  `Hide; Stocktaking_Form := TStocktaking_Form.Create(self); Stocktaking_Form.Execute;
  Stocktaking_Form.Free; Show;` idiom (**P14** — but here it *is* balanced with a trailing `Show`,
  unlike many P14 sites; still no `try..finally`). The button is `&Stocktaking` (MainMenu.dfm
  line 743); a duplicate `Window > InvMgmt > Stocktaking` menu item (`MainMenu.dfm` line 1086)
  fires the same handler. `Execute` runs `ShowModal`; returns `False` only on `mrCancel`.
- **Purpose (one paragraph):** A master-detail screen for recording inventory-count adjustments.
  A `DBGrid` (`Stocktaking_DBGrid`) lists all stocktaking rows (joined to part + supplier); a
  detail panel lets the user pick a **Supplier** (column combo) → which filters the dependent
  **Parts Code** combo, enter a **Quantity** delta (signed) and a free-text **Reason**, with a
  **Date** picker. Buttons: **Insert, Update, Search, Clear, Delete, Close**. Saving a row fires
  the qty-sync trigger that pushes the delta into `INV_PARTS_STOCK_MST.IN_QTY`.

## 2. Data touched
| Table | Read | Write | Notes |
|-------|:----:|:-----:|-------|
| `INV_STOCKTAKING_INF` | ✓ | ✓ | The stocktaking-adjustment ledger (this module owns it) |
| `INV_PARTS_STOCK_MST` | ✓ | ✓* | *Write is **indirect, via the three triggers** — never via a proc. Read: the SELECT joins it for part #/name; the INSERT proc resolves `VC_PART_NUMBER → IN_PART_ID` |
| `INV_SUPPLIER_MST` | ✓ | | FK lookup: the supplier column-combo (`SelectMultiField`) + the SELECT join for the displayed supplier code; `SELECT_DependantPartNumber_Supplier` resolves a supplier code → its parts |

### `INV_STOCKTAKING_INF` columns (authoritative: `DB Schema/Create Inventory.sql` line 1736)
| Column | Type | Meaning / notes |
|--------|------|-----------------|
| `IN_STOCKTAKING_ID` | `int IDENTITY(1,1) NOT NULL` | Surrogate key (`RecordID` in UI, hidden grid `Fields[6]`). **No PK *constraint* declared** (see Constraints) — identity only. |
| `IN_PART_ID` | `int NOT NULL` | FK-by-convention → `INV_PARTS_STOCK_MST.IN_PART_ID`. Set by the INSERT proc from the part number; the triggers key the stock-qty math on **this column**. **No declared FK.** |
| `IN_QTY` | `int NULL` | **The signed adjustment delta** (not an absolute count). Positive = add to stock; negative = subtract (e.g. scrap). The trigger adds/subtracts *this* value to `INV_PARTS_STOCK_MST.IN_QTY`. |
| `VC_REASON` | `varchar(300) NULL` | Free-text reason for the adjustment (form `Reason_Memo`, a multiline `TMemo` — **no MaxLength on the memo**, so >300 chars truncates at the DB; the INSERT/UPDATE proc params are `varchar(300)`, matching the column). |
| `VC_LAST_UPDATE` | `varchar(16) NULL` | **Timestamp as `yyyymmddHHMMSSff` string** (P2). Written on INSERT (= `VC_ADD`) and on UPDATE; the DELETE trigger recomputes a fresh one and stamps it onto the **stock master** (not this row, which is being deleted). |
| `VC_ADD` | `varchar(16) NOT NULL` | **Timestamp as `yyyymmddHHMMSSff` string** (P2), set once on INSERT. **This is the row's "date"** — `SELECT_StockTakingInfo` derives the displayed `Date` (`yyyy/mm/dd`) by `SUBSTRING`-slicing `VC_ADD`, **not** from a real date column ⚠️. |

**Note:** the form has a **Date picker** (`Edit_DateTimePicker`, format `yyyy/MM/dd`), but **the
entered date is never persisted** — `INSERT_StockTakingInfo` ignores it and stamps `VC_ADD`/`VC_LAST_UPDATE`
from `getdate()`. `HoldDetails(False)` computes `EditDate := FormatDateTime('yyyymmdd', picker)` but
the DataModule insert/update never sends it. So the picker is effectively decorative on write; on read
the grid's "Date" comes from `VC_ADD`. ⚠️ A real surprise for parity (see §4, §8).

**Constraints / indexes (verified by reading the schema's ALTER-TABLE block, lines 1824-2000):**
- **NO PRIMARY KEY constraint, NO unique index, NO foreign key, NO default constraint** on
  `INV_STOCKTAKING_INF`. The only structure is the `IN_STOCKTAKING_ID IDENTITY` column. This is a
  **transaction/ledger table with zero declared integrity guards** — comparable posture to the
  least-protected masters (cf. **P11** on `INV_MANIFEST_COST_MST`), but here it is an append-mostly
  ledger so the missing PK is less load-bearing than the missing FK to parts.
- The whole schema declares **exactly two** FOREIGN KEY constraints (ASN-detail→ASN, part-shipping→shipping);
  **neither involves stocktaking**. `IN_PART_ID`'s link to `INV_PARTS_STOCK_MST` is **by convention only**,
  enforced solely by the INSERT proc's lookup and the triggers.

### `INV_PARTS_STOCK_MST` (the table the triggers mutate — relevant columns)
| Column | Type | Relevance |
|--------|------|-----------|
| `IN_PART_ID` | `int IDENTITY PK` (`PK_INV_PARTS_STOCK_MST`, clustered) | The join key the triggers use |
| `VC_PART_NUMBER` | `varchar(12) NOT NULL` | **UNIQUE via `IX_INV_PARTS_STOCK_MST` (nonclustered)** — unique part number (per decision D1 this becomes unique **per-site**, composite `(site_id, VC_PART_NUMBER)`, not global); the INSERT proc resolves it to `IN_PART_ID` |
| `IN_QTY` | `int NULL` | **The on-hand quantity the stocktaking triggers adjust** — the core invariant |
| `VC_LAST_UPDATE` | `varchar(16) NULL` | All three triggers stamp this when they adjust `IN_QTY` |

**Triggers on these tables (LIVE bodies from `DB Schema/Create Inventory.sql`, lines 10383-10464 — the
authoritative source; `docs/triggers.sql` carries an **obsolete** `*_StockTakingPartsStockMstQTY`-named,
string-code-keyed variant that must NOT be used, per
[`trigger-source-reconciliation.md`](../cross-cutting/trigger-source-reconciliation.md)):**

The **three stocktaking triggers fire on `INV_STOCKTAKING_INF`** and adjust `INV_PARTS_STOCK_MST.IN_QTY`,
keyed on **`IN_PART_ID`** (int FK). **Exact qty math (the core inventory invariant):**

- **`INSERT_Stocktaking`** (FOR INSERT) — **adds the new delta to stock:**
  ```sql
  UPDATE INV_PARTS_STOCK_MST
    SET IN_QTY = PS.IN_QTY + i.IN_QTY,
        VC_LAST_UPDATE = i.VC_LAST_UPDATE
  FROM INV_PARTS_STOCK_MST PS, inserted i
  WHERE PS.IN_PART_ID = i.IN_PART_ID
  ```
  Invariant: `IN_QTY := IN_QTY + inserted.IN_QTY`. Since the delta can be negative, an insert can
  decrease stock (scrap). Also copies the new row's `VC_LAST_UPDATE` onto the stock row.

- **`UPDATE_Stocktaking`** (FOR UPDATE) — **reverses the old delta, then applies the new one**
  (two statements, run as one trigger so the net effect is `+ (new − old)`):
  ```sql
  UPDATE INV_PARTS_STOCK_MST                       -- (1) back out the OLD value
    SET IN_QTY = PS.IN_QTY - d.IN_QTY, VC_LAST_UPDATE = d.VC_LAST_UPDATE
  FROM INV_PARTS_STOCK_MST PS, deleted d  WHERE PS.IN_PART_ID = d.IN_PART_ID
  UPDATE INV_PARTS_STOCK_MST                       -- (2) apply the NEW value
    SET IN_QTY = PS.IN_QTY + i.IN_QTY, VC_LAST_UPDATE = i.VC_LAST_UPDATE
  FROM INV_PARTS_STOCK_MST PS, inserted i WHERE PS.IN_PART_ID = i.IN_PART_ID
  ```
  Invariant: `IN_QTY := IN_QTY − deleted.IN_QTY + inserted.IN_QTY` (net change in stock = change in
  the adjustment delta). ⚠️ **If `IN_PART_ID` were changed on update, statement (1) backs the old
  delta off the *old* part and statement (2) adds the new delta to the *new* part** — but the UPDATE
  proc never changes `IN_PART_ID` (it only writes `IN_QTY`, `VC_REASON`, `VC_LAST_UPDATE`), so in
  practice both statements hit the same part. `VC_LAST_UPDATE` on the stock row ends as the *new*
  row's value.

- **`DELETE_Stocktaking`** (FOR DELETE) — **backs the deleted delta out of stock:**
  ```sql
  DECLARE @Deleted varchar(16)
  SET @Deleted = CONVERT(varchar,getdate(),112)
               + SUBSTRING(CONVERT(varchar,getdate(),114),1,2)  -- HH
               + SUBSTRING(CONVERT(varchar,getdate(),114),4,2)  -- MM
               + SUBSTRING(CONVERT(varchar,getdate(),114),7,2)  -- SS
               + SUBSTRING(CONVERT(varchar,getdate(),114),10,2) -- ff
  UPDATE INV_PARTS_STOCK_MST
    SET IN_QTY = PS.IN_QTY - d.IN_QTY, VC_LAST_UPDATE = @Deleted
  FROM INV_PARTS_STOCK_MST PS, deleted d WHERE PS.IN_PART_ID = d.IN_PART_ID
  ```
  Invariant: `IN_QTY := IN_QTY − deleted.IN_QTY` (deleting an adjustment reverses it). Unlike the
  other two, it **computes its own fresh `yyyymmddHHMMSSff` timestamp** (the deleted row's value is
  gone) and stamps it onto the stock master. The full insert→delete cycle is **qty-neutral**
  (`+delta` then `−delta`), which is the property the rebuild must preserve exactly.

> **Re-home target:** these three triggers ARE the inventory invariant. In the rebuild, a
> `StocktakingAdjustment` create/update/destroy must atomically apply `+delta / +(new−old) / −delta`
> to `PartsStock#in_qty` (and bump its `vc_last_update`) **inside one DB transaction** (model callback
> or service object). Do not lose the sign convention or the "delta, not absolute" semantics.

## 3. Stored procedures used
(Read with `sql.sh proc NAME`. The procs are the behavioral spec.)

| Proc | Op | Business rule (from the body) |
|------|----|-------------------------------|
| `SELECT_StockTakingInfo;1` (**no params**) | SELECT | Returns **all** rows (no filter, no paging) — per decision D1 the rebuild scopes this to the current `site_id`. Joins `INV_STOCKTAKING_INF ST → INV_PARTS_STOCK_MST PS ON ST.IN_PART_ID = PS.IN_PART_ID → INV_SUPPLIER_MST SM ON SM.IN_SUPPLIER_ID = PS.IN_SUPPLIER_ID`. **Inner JOINs** ⚠️: a stocktaking row whose part has `IN_SUPPLIER_ID = NULL` (e.g. after a `DELETE_SupplierCode` unlink, P5) is **silently dropped from the list** (and thus uneditable from this screen, though its stock delta already applied). Selects 7 UI-aliased cols mapping 1:1 to grid `Fields[0..6]`: `Date` (sliced from `VC_ADD` → `yyyy/mm/dd`), `Supplier Code`, `Parts Code`, `Parts Name`, `QTY`, `Reason`, `RecordID` (= `IN_STOCKTAKING_ID`). `ORDER BY SM.VC_SUPPLIER_CODE, PS.VC_PART_NUMBER`. |
| `INSERT_StockTakingInfo;1` (**3 params**: `@PartNumber varchar(12), @QTY int, @Reason varchar(300)`) | INSERT | Computes `@Add` = `yyyymmddHHMMSSff` string (P2). **Resolves part number → id:** `SELECT @PartID = IN_PART_ID FROM INV_PARTS_STOCK_MST WHERE VC_PART_NUMBER = @PartNumber` (**P3** — name→id in-proc; note **no supplier code is used in the lookup** — the part number is unique via `IX_INV_PARTS_STOCK_MST`; per decision D1 this lookup must also be **scoped to the current site**, since `VC_PART_NUMBER` becomes unique only per-site `(site_id, VC_PART_NUMBER)`). Inserts `(IN_PART_ID, IN_QTY, VC_REASON, VC_LAST_UPDATE=@Add, VC_ADD=@Add)` — **named column list** (not positional, contrast P10). ⚠️ **If the part number doesn't match any stock row, `@PartID` stays NULL** and the insert violates `IN_PART_ID NOT NULL` → error (no row, no trigger). **No duplicate guard** (none needed; a ledger allows many rows per part). |
| `UPDATE_StockTakingInfo;1` (**5 params**: `@SupCode varchar(5), @PartCode varchar(12), @QTY int, @Reason varchar(300), @StocktakingID int`) | UPDATE | Updates **one row `WHERE IN_STOCKTAKING_ID = @StocktakingID`**, setting **only** `VC_LAST_UPDATE`, `IN_QTY`, `VC_REASON`. ⚠️ **`@SupCode` and `@PartCode` are declared but NEVER used** — the part/supplier of an existing row **cannot be changed** here (only qty + reason), and `IN_PART_ID` is left intact (so the `UPDATE_Stocktaking` trigger's two statements always hit the same part). ⚠️⚠️ **TIMESTAMP BUG:** `@Update` is **read before it is set** — `SET @Update = SUBSTRING(@Update,1,8) + …`. `@Update` is an undeclared-value (NULL) `varchar(16)` at that point, so `SUBSTRING(NULL,1,8)` → NULL and the whole expression → **NULL**. Net: **`VC_LAST_UPDATE` is set to NULL on every update**, and the `UPDATE_Stocktaking` trigger then stamps **NULL into `INV_PARTS_STOCK_MST.VC_LAST_UPDATE`** for that part. (The intended code clearly meant `SUBSTRING(CONVERT(varchar,getdate(),112),1,8)` for the date portion, as the insert/delete versions do.) Capture as a **bug to fix, not replicate** (§8/§9). |
| `DELETE_StocktakingInfo;1` (**1 param**: `@StocktakingID int`) | DELETE | Hard-deletes `WHERE IN_STOCKTAKING_ID = @StocktakingID`. Relies entirely on `DELETE_Stocktaking` to reverse the stock delta. No in-use / RI check. |

⚠️ **Proc-name casing is inconsistent in the schema:** the SELECT/INSERT/UPDATE procs are spelled
`SELECT_StockTakingInfo` / `INSERT_StockTakingInfo` / `UPDATE_StockTakingInfo` (**capital `T`** in
"StockTaking"), but `DELETE_StocktakingInfo` uses a **lowercase `t`**. SQL Server object names are
case-insensitive so all four calls resolve regardless, but the rebuild should **match the schema
casing** (or normalize all four) rather than assume a single convention.

**Cross-module note — also called by the part-number lookup helper:**
| Proc | Op | Where used |
|------|----|------------|
| `SELECT_DependantPartNumber_Supplier @SupplierCode varchar(5)` | SELECT | `SELECT * FROM INV_PARTS_STOCK_MST p JOIN INV_SUPPLIER_MST s ON s.IN_SUPPLIER_ID = p.IN_SUPPLIER_ID WHERE s.VC_SUPPLIER_CODE = @SupplierCode`. Called via `SelectDependantSingleField(…, 'VC_PART_NUMBER', …, PartsCode_ComboBox)` on supplier-combo change — populates the Parts Code combo with that supplier's part numbers. Shared lookup proc (also used by other order/receiving screens). |

### Call mechanism (legacy)
`DataModule.pas` methods `GetStocktakingInfo / InsertStocktakingInfo / UpdateStocktakingInfo /
DeleteStocktakingInfo` (declarations 577-580; bodies 3431-3620) drive the shared ADO objects (**P6**).
Notable per-method facts:
- **`GetStocktakingInfo`** (3431-3470) uses **`Inv_DataSet`** (open result set), no params; logs
  `LogActLog('GET StkTak', 'SELECTED all Stocktaking info', 1)`. The other three use `Inv_StoredProc`
  (`ExecProc`).
- ⚠️ **Param-name mismatch (relies on ADO positional binding):**
  - `InsertStocktakingInfo` (3472-3523) adds params named **`@PartCode`, `@QTY`, `@Reason`** — but the
    proc declares **`@PartNumber`, @QTY, @Reason**. The names differ (`@PartCode` ≠ `@PartNumber`), yet
    the **three values still line up by position** with the proc's three params, so under either
    name- or position-binding the part number can reach `@PartNumber` (name-binding would fail to find
    `@PartCode`, but the values are positionally aligned). It is fragile (a parameter reorder would
    silently corrupt the write, like P10's positional risk). **Contrast `InsertAutoScrap`** (§3,
    "Second writer"), whose params are **not** positionally aligned — it prepends `@SupCode` ahead of
    the part code, so position-binding there would corrupt the write.
  - `UpdateStocktakingInfo` (3525-3577) sends all five params including the two **dead** ones
    (`@SupCode := fSupCode`, `@PartCode := fPartNum`) plus `@StocktakingID := fRecordID`.
- **`DeleteStocktakingInfo`** (3579-3620) sends only `@StocktakingID := fRecordID`; logs
  `LogActLog('DEL StkTak', 'DELETED S: '+fSupCode+' P: '+fPartNum, 1)`.
- **Retry harness (P8) — all four self-recurse correctly (NOT in the P12 wrong-target register):**
  each wraps its ADO call in `fErrorCount := fErrorCount + 1; If fErrorCount < 3 Then <same method>`
  with a `finally` resetting `fErrorCount := 0` (and Insert/Delete also `Inv_StoredProc.Close`). The
  four stocktaking methods **retry into themselves** — verified against
  [`datamodule-retry-target-bugs.md`](../cross-cutting/datamodule-retry-target-bugs.md); they are
  **not** among the 29 wrong-target bugs. **One adjacent bug touches this module, though:** the
  register's LOW-severity `GetBuildHist → GetStocktakingInfo` mis-retry (DataModule.pas **line 4450**)
  means a transient error during `GetBuildHist` retries into `GetStocktakingInfo`, which loads the
  **param-less `SELECT_StockTakingInfo`** into the shared `Inv_DataSet` — a **different-shape read**
  than `GetBuildHist`'s own **4-param `SELECT_AssyBuildHist`**, so the wrong (stocktaking-list) result
  set lands where build history was expected. Harmless in practice (read-only; no persistence — see §8).
- **Record key is the shared, generic `fRecordID`** (P9): `HoldDetails(True)` captures grid
  `Fields[6]` (`IN_STOCKTAKING_ID`) into `Data_Module.RecordID`; Update/Delete key off it. If no row
  was selected, `RecordID` is 0 or a **stale value from another screen** — the same cross-module
  write-to-wrong-row latent hazard as the masters. `Update_ButtonClick` saves `ID := RecordID` before
  the refresh and re-`Locate('RecordID', ID)` afterward.

### Second writer to this table — `InsertAutoScrap` (headless, from DailyBuildTotal)
`DataModule.InsertAutoScrap` (DataModule.pas 4460-4529) is a **separate write path into
`INV_STOCKTAKING_INF` via the same `INSERT_StockTakingInfo` proc**, called from **`DailyBuildTotal.pas`
line 243** (the daily-build-total Excel import). It:
1. Reads the part's current stock via `SELECT_PartsStockInfo @InvMgmtReport='N', @SupCode='', @PartNum=fPartNum`.
2. If `Last Scrap Count <> fScrapCount`, inserts a stocktaking row with
   **`@QTY := 0 - (fScrapCount - LastScrapCount)`** (a **negative delta** = scrap removed from stock),
   `@Reason := 'Auto Scrap Delete on '+now`. ⚠️⚠️ **It sends FIVE params, in order
   `@SupCode, @PartCode, @QTY, @Reason, @AutoScrap`** (it **prepends** `@SupCode` and **appends**
   `@AutoScrap` around the proc's three real params) — but the proc declares only `@PartNumber, @QTY,
   @Reason`. Contrast the manual `InsertStocktakingInfo` (DataModule.pas 3486-3491), which sends only
   `@PartCode, @QTY, @Reason` and therefore aligns positionally with `@PartNumber, @QTY, @Reason`.
   Here the **prepended `@SupCode` shifts every positional slot by one**, so under **position-binding**
   the supplier code would land in `@PartNumber`, the part code in `@QTY`, etc. — corrupting the write
   (and the proc would then fail to resolve a part). Whether this actually corrupts the insert or works
   correctly **depends entirely on whether ADO binds these params by name or by position** — unlike the
   manual path, the names here do *not* line up by position, so this is not a benign reorder. (Both
   paths share the param-name mismatch `@PartCode` ≠ proc's `@PartNumber`; see §3 binding note.) Logs
   `LogActLog('INS BuildHist','Insert auto scrap …')` and `'SCRAP'`.
> **Implication for the rebuild:** the stock-qty effect of stocktaking is reachable from **two**
> screens (manual Stocktaking + automated DailyBuildTotal scrap reconciliation). Both must route
> through the **same** adjustment service so the `IN_QTY` invariant and audit trail stay consistent.

## 4. Business rules & edge cases
- **It is a signed-delta adjustment, not an absolute count.** Entered `IN_QTY` is **added** to
  on-hand stock by `INSERT_Stocktaking` (positive raises stock, negative lowers it). The Qty mask
  `'#######;1; '` (`MaxLength=7`) **permits a leading minus**, so values roughly −999999..9999999
  are enterable ⚠️ (DB `int` allows far more; the form caps at 7 mask chars). `TextChange` forces a
  blank mask to `'0'`, and `HoldDetails(False)` does `StrToInt(Trim(Qty))`, so a blank submits `0`
  (a no-op adjustment), never NULL.
- **Part identity is the 12-char part number.** `HoldDetails(False)` **rejects any Parts Code whose
  trimmed length ≠ 12** (`'Invalid Part Code'`) — the *only* validation in the form. The part number
  is unique (`IX_INV_PARTS_STOCK_MST`; per decision D1 unique **per-site**, composite `(site_id,
  VC_PART_NUMBER)`), so the supplier shown is purely contextual; the INSERT proc resolves the part
  **by number alone** (supplier code is not part of the lookup — but the resolution must be site-scoped).
- **The Date picker is not persisted (⚠️ surprise).** Although the user picks a date, the INSERT proc
  stamps `VC_ADD`/`VC_LAST_UPDATE` from `getdate()` and the displayed grid "Date" is sliced from
  `VC_ADD`. So a backdated count cannot actually be recorded as a past date. The picker only influences
  `SetDetailBoxes` round-tripping. (Confirm desired behavior in §8.)
- **Update changes only qty + reason.** `UPDATE_StockTakingInfo` ignores supplier/part params; you
  cannot re-point an adjustment to a different part from this screen. The trigger re-balances stock by
  `new − old` delta.
- **Update writes a NULL timestamp (bug).** The `SUBSTRING(@Update,1,8)` on an unset `@Update` makes
  `VC_LAST_UPDATE` NULL on the stocktaking row **and** (via the trigger) NULLs the stock master's
  `VC_LAST_UPDATE`. Fix in the rebuild; do not reproduce.
- **Delete reverses the adjustment** (trigger subtracts the deleted delta) — the table is an auditable
  ledger where create and delete are qty-symmetric. There is **no soft-delete flag**.
- **No referential integrity to parts.** `IN_PART_ID` has no FK; if a part stock row were deleted while
  stocktaking rows referenced it, those rows orphan and `SELECT_StockTakingInfo`'s inner join drops
  them. (`DELETE_PartNumber` is a live trigger on `INV_PARTS_STOCK_MST` — out of scope here, but worth
  noting the part side has its own cascade behavior.)
- **Inner-join visibility gap.** A stocktaking row for a part with `IN_SUPPLIER_ID = NULL` (supplier was
  deleted, P5) disappears from the list even though its stock delta already applied — it becomes an
  invisible, uneditable, un-reversible-from-UI adjustment.
- **Timestamps are `yyyymmddHHMMSSff` strings (P2):** `VC_ADD` on insert, `VC_LAST_UPDATE` on
  insert/update; the DELETE trigger computes its own. 16-char recipe `CONVERT(…,112)` + 4×`SUBSTRING(…,114)`
  (HH+MM+SS+ff) — the byte-identical recipe used by every master.
- **Auto-scrap path** posts negative deltas headlessly (see §3) — same trigger, same invariant.

## 5. UI / UX notes
- Grid + detail-panel pattern; selecting a grid row (`OnMouseUp` / `OnKeyUp` /
  `Stocktaking_DataSourceDataChange`, all gated by `fNoChange`) copies the row into the detail
  controls and captures `RecordID` from hidden `Fields[6]`.
- **Supplier combo is a `TNUMMIColumnComboBox`** (two-column: code + name) populated by
  `SelectMultiField('INV_SUPPLIER_MST','VC_SUPPLIER_CODE, VC_SUPPLIER_NAME', …)`. Its `OnChange` calls
  `SELECT_DependantPartNumber_Supplier` to repopulate the **Parts Code combo** (`csDropDownList`) with
  that supplier's part numbers — a cascading dependent select.
- **Search is fully client-side (P7):** `SearchGrid` sets `Inv_DataSet.Filter` to
  `[Supplier Code] = '…' [AND [Parts Code] = '…']` and `Filtered := True` over the **already-loaded**
  dataset — exact equality, no LIKE, no DB re-query. `Search_ButtonClick` only searches when **both**
  supplier and part are chosen.
- **Fields:** Date (`TDateTimePicker`, decorative on write), Supplier (column combo), Parts Code
  (dependent combo, must be 12 chars), Quantity (`TMaskEdit '#######;1; '`, signed, `MaxLength=7`),
  Reason (`TMemo`, no length cap → DB caps at 300).
- **Minor:** `Qty_MaskEdit` carries `CharCase = ecUpperCase` (`Stocktaking.dfm` line 197) — the **only**
  `CharCase` setting on the form. It is **immaterial** for a numeric mask (digits/`-` have no case), but
  worth noting as a stray default if the field is ever repurposed; no need to carry it into the rebuild.
- **Modernize:** index/list with **server-side search/filter/pagination** (P7); a new/edit form whose
  Qty field clearly signals "signed adjustment" (+/− with helper text); a real **date** field if
  backdating should be supported (§8); inline part-number validation (length/existence) replacing the
  single `≠ 12` check; show the *resulting* stock level after the adjustment for operator confidence.

## 6. Target design  *(Rails primary)*
- **Models:**
  - `StocktakingAdjustment` → `self.table_name = 'INV_STOCKTAKING_INF'`,
    `self.primary_key = 'IN_STOCKTAKING_ID'`. `belongs_to :parts_stock, foreign_key: 'IN_PART_ID'`
    (model the link the legacy never declared as a real FK). **`belongs_to :site` (decision D1)** —
    every adjustment carries a `site_id` (NOT NULL), with enforced current-site scoping (default scope
    / query filter on `site_id`) so the unfiltered legacy list becomes per-site. Validations: `in_qty`
    numericality (integer; allow negative — it's a delta), `parts_stock` presence (the part must
    resolve **within the current site** — replaces the proc's silent NULL-`@PartID` failure),
    `vc_reason` length ≤ 300, `site` presence.
  - `PartsStock` → `INV_PARTS_STOCK_MST` (`primary_key 'IN_PART_ID'`); `has_many
    :stocktaking_adjustments, foreign_key: 'IN_PART_ID'`. Also `belongs_to :site` (D1) — on-hand stock
    is per-site, and its unique index on the part number becomes composite **`(site_id,
    VC_PART_NUMBER)`** (was global `IX_INV_PARTS_STOCK_MST`).
  - Timestamps: write `vc_add`/`vc_last_update` as `yyyymmddHHMMSSff` strings during parallel run (P2);
    normalize at the Postgres phase. **Fix the update-timestamp NULL bug** — always compute a real value.
- **The stock-qty invariant → a service / transactional callback (replacing the 3 triggers):**
  - `create` ⇒ `parts_stock.in_qty += delta` (and `vc_last_update`).
  - `update` ⇒ `parts_stock.in_qty += (new_delta − old_delta)` (use `in_qty_was`); update timestamp.
  - `destroy` ⇒ `parts_stock.in_qty -= delta`; stamp a fresh timestamp.
  - All inside **one DB transaction with row locking** on the part (the legacy relied on SQL Server
    trigger atomicity — preserve that). Centralize so the **DailyBuildTotal auto-scrap** writer reuses
    the exact same path (a `negative` delta with reason "Auto Scrap …").
- **Controller/routes:** RESTful `resources :stocktaking_adjustments` (index/new/create/edit/update/destroy).
- **Views:** index (server-side searchable list, replacing the in-memory `Filter`, P7) + new/edit form
  with cascading Supplier→Part selects and a signed-qty input. Decide whether to expose an editable
  adjustment date (§8).
- **Services:** none external; pure DB + the adjustment transaction above. **Stage-1 option:** wrap the
  four stocktaking procs (`SELECT_StockTakingInfo` / `INSERT_StockTakingInfo` / `UPDATE_StockTakingInfo`
  / `DELETE_StocktakingInfo` — mind the `StockTaking` vs `Stocktaking` casing split) via `tiny_tds` for
  parity (keeping the trigger-driven qty math), but
  **do not** reproduce the `UPDATE_StockTakingInfo` NULL-timestamp bug — fix it in the wrapped proc or
  in stage-3 app logic.
- **Reports:** none specific (no `REPORT_*` proc for stocktaking).

## 7. Migration plan for this module
- [ ] **Stage 1 — read-only parity:** `StocktakingAdjustment` joined to part+supplier renders the list,
      ordered by `VC_SUPPLIER_CODE, VC_PART_NUMBER`; 7 columns map 1:1 to grid `Fields[0..6]`; "Date"
      derived from `VC_ADD`. Reproduce the **inner-join visibility gap** only if parity demands it
      (flag it as a known legacy quirk, then prefer a LEFT JOIN so NULL-supplier rows stay visible).
- [ ] **Stage 2 — writes via wrapped procs:** call `INSERT_StockTakingInfo` / `UPDATE_StockTakingInfo` / `DELETE_StocktakingInfo` (note the casing split) through
      `tiny_tds`, preserving the **three triggers'** stock-qty math (the invariant). Route DailyBuildTotal
      auto-scrap through the same wrapper. **Patch the UPDATE NULL-timestamp bug.**
- [ ] **Stage 3 — reimplement (Postgres-ready):** move the +delta / +(new−old) / −delta stock-qty
      adjustment into a **transactional model callback / service** (replacing the 3 triggers); add a
      **real FK** `IN_PART_ID → INV_PARTS_STOCK_MST` and a **PK constraint** on `IN_STOCKTAKING_ID`
      (legacy has neither); real timestamps; presence/numericality validations; server-side search.
      **Multi-site (decision D1):** in this Postgres phase add the `site_id` (NOT NULL) FK →
      `sites` table on `INV_STOCKTAKING_INF` and the **per-site unique index** `(site_id,
      VC_PART_NUMBER)` on `INV_PARTS_STOCK_MST` (replacing the global `IX_INV_PARTS_STOCK_MST`); the
      legacy single-site DB stays untouched during the parallel run.

## 8. Open questions for the user (domain expert)
1. ✅ **RESOLVED (D5): signed adjustment delta.** Per decision D5 (docs/analysis/decisions.md),
   stocktaking `IN_QTY` is a **signed adjustment delta** — the triggers add/subtract it from on-hand
   (entering `100` raises on-hand by 100; `-30` lowers it by 30), **not** an absolute counted total.
   The legacy trigger behavior is the intended behavior and is preserved. The rebuild's UI must label
   the field as an adjustment (+/−) so it's never mistaken for "set on-hand to this count"; a true
   physical-count→set-absolute flow, if ever wanted, is a separate feature that computes the delta.
2. **The unpersisted Date picker:** the form lets users pick a date but it is **never saved** — the row
   timestamp is always "now". Should the rebuild (a) keep "now" semantics and drop the picker, or (b)
   persist a real **adjustment/count date** (enabling backdated physical counts)?
3. **`UPDATE_StockTakingInfo` NULL-timestamp bug:** confirm this is unintended (it almost certainly is)
   so the rebuild fixes it. It currently NULLs `VC_LAST_UPDATE` on both the stocktaking row and the
   affected stock master row on every edit.
4. ✅ **RESOLVED (D3): yes — add `PK_INV_STOCKTAKING_INF` + a real FK, and block deleting a part
   with adjustments (RESTRICT).** Per decision D3 (docs/analysis/decisions.md), the rebuild adds the
   real `PK_INV_STOCKTAKING_INF` on `IN_STOCKTAKING_ID` and the FK `IN_PART_ID →
   INV_PARTS_STOCK_MST`. Adjustments should **not** survive their part's deletion as orphans — the FK
   **blocks** deleting a part that still has stocktaking rows (no orphan/hide). Retiring such a part
   goes through the future **archival** capability, which preserves the adjustment history intact.
5. **Auto-scrap coupling:** `DailyBuildTotal` posts negative stocktaking rows ("Auto Scrap Delete")
   through this same table/trigger. Should auto-scrap remain modeled as a stocktaking adjustment (same
   ledger, distinguishable by reason), or become its own adjustment type/reason-code?
6. **Multi-site:** ✅ RESOLVED (D1): per-site — sites run fully isolated (no shared inventory/data),
   so `INV_STOCKTAKING_INF` gains a `site_id` (NOT NULL) FK → new `sites` table, every adjustment is
   scoped to the current site, and `SELECT_StockTakingInfo` (today returns **all** rows for the whole
   DB unfiltered) becomes site-filtered. `VC_PART_NUMBER` is no longer globally unique but unique
   **per-site** (composite `(site_id, VC_PART_NUMBER)`), and on-hand stock is per-site, so stocktaking
   rows are inherently site-scoped — eliminating the cross-site stock-qty leakage this question raised.
   See decision D1 (docs/analysis/decisions.md).
7. **Reason as free text:** should `VC_REASON` become a **coded reason** (cycle count, damage, scrap,
   correction…) with optional free text, for reporting? Today it's unstructured `varchar(300)`.

## 9. Test cases / parity checks
- **List all** → row count/ordering match `SELECT_StockTakingInfo` (sorted by supplier code, then part
  number); 7 columns map 1:1 to grid `Fields[0..6]`; "Date" = `yyyy/mm/dd` sliced from `VC_ADD`.
  **Edge:** a stocktaking row whose part has `IN_SUPPLIER_ID = NULL` does **not** appear (inner join) —
  assert legacy parity, then verify the new app's chosen LEFT-JOIN behavior keeps it visible.
- **Insert +N (positive delta)** for a part at on-hand `Q` → new row with `VC_ADD`/`VC_LAST_UPDATE` as a
  16-char `yyyymmddHHMMSSff` string; **`INV_PARTS_STOCK_MST.IN_QTY` becomes `Q + N`** and its
  `VC_LAST_UPDATE` = the new row's value (trigger `INSERT_Stocktaking`).
- **Insert −N (negative delta / scrap)** → stock becomes `Q − N` (verify negatives are accepted by the
  mask and the math; the trigger uses `+ i.IN_QTY`).
- **Update an adjustment's qty from `N` to `M`** → stock changes by `M − N` (trigger backs out `N`, adds
  `M`); the stocktaking row's `IN_QTY` = `M`, `VC_REASON` updated, **part/supplier unchanged**.
  **Bug-parity note:** legacy sets `VC_LAST_UPDATE = NULL` on both the row and the stock master — the new
  app must instead write a valid timestamp (document this as an **intentional divergence**, §8.3).
- **Delete an adjustment of delta `D`** → stock becomes `current − D`; stocktaking row gone; stock
  `VC_LAST_UPDATE` = a fresh timestamp computed by `DELETE_Stocktaking`.
- **Round-trip neutrality:** insert `+D` then delete that row → `IN_QTY` returns to its original value
  (the core symmetry property).
- **Insert with a non-existent part number** → INSERT proc leaves `@PartID` NULL → `IN_PART_ID NOT NULL`
  violation → no row inserted, **no stock change, no trigger** (verify the new app rejects with a clean
  validation rather than a raw error).
- **Insert with part code ≠ 12 chars** → legacy form blocks it ("Invalid Part Code"); new app validates
  part-number existence/format and rejects.
- **Auto-scrap path** (`DailyBuildTotal`): when imported scrap count differs from `Last Scrap Count`,
  a stocktaking row with `IN_QTY = -(new − last)` and reason "Auto Scrap Delete on …" is inserted and
  stock drops by that amount — assert the new app routes this through the same adjustment service.
