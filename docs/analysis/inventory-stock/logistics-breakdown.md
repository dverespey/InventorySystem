# Module Analysis: Logistics Breakdown (inbound order-status file processor)

**Area:** Inventory / Stock  **Status:** ✅ spec complete  **Analyst:** Claude / 2026-06-05

> **Not a master-data screen.** Despite the "Logistics" name, this module is **unrelated to
> `LogisticsMaster`** (the carrier master — see [`../master-data/logistics.md`](../master-data/logistics.md) §5).
> `LogisticsBreakdown` is an **inbound batch file processor**: it parses a **fixed-width text file** of
> in-transit / arrival / termination status lines coming back from the carrier (the "logistics" feed),
> and for each line drives a **status UPDATE on the open-order row** keyed by **Renban number**. It is the
> counterpart to the other `*Breakdown` processors (`ForecastBreakdownF`, `InvoiceBreakdown`,
> `ManualForecast`, `DailyBuildTotal`) all launched from the shared `UploadBreakDown` hub. **It writes the
> order's lifecycle status, and — because its target table `INV_OPEN_ORDER_INF` carries the
> `UPDATE_RecConfStatPartsStockMstQTY` trigger — it indirectly adjusts the core inventory invariant
> `INV_PARTS_STOCK_MST.IN_QTY`.** That trigger coupling is the single most important fact in this spec.

## 1. Legacy surface
- **Form:** `LogisticsBreakdown.pas` (~11.6 KB, 255 lines) + `LogisticsBreakdown.dfm` (~1 KB, 48 lines).
  Class `TLogisticsBreakdown_Form`, Caption "Logistics Breakdown". Registered live in
  `InventorySystem.dpr` **line 35** (`LogisticsBreakdown in 'LogisticsBreakdown.pas' {LogisticsBreakdown_Form}`).
  The form is almost UI-less: a full-width read-only **`THistory` log pane** (`Hist`, a third-party VCL
  progress/history control), a hidden **`OK` button** (revealed only when processing finishes), and a
  non-visual **`TCopyFile`** component (`CopyFile`, `MoveFile = True`) for archiving the consumed file.
- **Entry point:** **Not reached directly from `MainMenu.pas`.** `MainMenu` (line 590,
  `UpBreakDown_Form.BreakdownKind := bReceiving`) opens the shared hub `UploadBreakDown.pas`
  (`TUpBreakDown_Form`, dpr listed) in the **`bReceiving`** mode. The hub's Browse dialog filters for
  `*.txt`/`*.log` "Logistics file" seeded from INI `[DIRECTORIES] LogisticsInputDir` (default
  `c:\_Inventory_Control\`) and `[DIRECTORIES] LogisticsFilename`. On **Start**
  (`UploadBreakDown.Start_ButtonClick`, lines 200-206, the `bReceiving` branch):
  `LogisticsBreakdown_Form := TLogisticsBreakdown_Form.Create(self); LogisticsBreakdown_Form.filename :=
  ForecastFilleNameDialog.FileName; LogisticsBreakdown_Form.Execute; LogisticsBreakdown_Form.Free;`
  (the `Hide/Create/Execute/Free` child-launch idiom, **P14**). `Execute` is just `ShowModal`; **all work
  runs in `FormShow`** (fired on display), so the file is processed the instant the form appears — there is
  no Start/confirm step inside this form.
- **Purpose (one paragraph):** Read a carrier status file line by line; for each line extract the
  **Renban** (lot) number and a **status keyword**; verify the Renban exists in the open-order table; then
  call the matching `UPDATE_Order*` proc to stamp the order's shipping / plant-yard / warehouse /
  terminated status (plus trailer number, plant parking/Lot). It is essentially the legacy app's
  **EDI-856-inbound-status / yard-management ingest** done as a flat-file batch rather than a true X12 856.
  Mislabelled "Receiving" in the menu and "Logistics" in the class name; functionally it is **inbound
  order-status reconciliation**.

## 2. Data touched
| Table | Read | Write | Notes |
|-------|:----:|:-----:|-------|
| `INV_OPEN_ORDER_INF` | ✓ | ✓ | The open-purchase-order table. Read via `SELECT_OrderOpenOrderLog` (existence check by Renban); written via the four `UPDATE_Order*` procs (status columns). **This module owns the write path.** |
| `INV_OPEN_ORDER_INF_HIST` |  | ✓* | *Indirect: the live `UPDATE_RecConfStatPartsStockMstQTY` / `INSERT_…` triggers `INSERT … SELECT * from deleted/inserted` into this history table on every order change. |
| `INV_PARTS_STOCK_MST` |  | ✓** | **Indirect: the `UPDATE_RecConfStatPartsStockMstQTY` trigger adjusts `IN_QTY` (and `VC_LAST_UPDATE`) when an order's shipping/arrival status flips — **the core inventory invariant** (see Triggers). |
| `INV_SUPPLIER_MST` | ✓*** | | ***Indirect (read-only inside the trigger): the stock-qty trigger JOINs supplier to read `VC_INVENTORY_ADD_POINT` (`S`/`A`, P4) which decides *whether* a status flip moves stock. |
| `Purge` | ✓*** | | ***Read inside the DELETE trigger only (`PurgeMode`); not relevant to this module's UPDATE path. |

This module reads no master/lookup data of its own and has **no FK-lookup combos** — it is a headless batch
job. The only "input" is the flat file plus INI directory config.

### `INV_OPEN_ORDER_INF` columns (authoritative: `DB Schema/Create Inventory.sql` lines 1484-1509)
| Column | Type | Written by this module? | Meaning / notes |
|--------|------|:----:|-----------------|
| `IN_ORDER_ID` | `int IDENTITY(1,1) NOT NULL` PK | | Surrogate key (`PK_INV_OPEN_ORDER_INF` CLUSTERED). **Not used by this module** — all updates key on Renban, not id. |
| `VC_SUPPLIER_CODE` | `varchar(5) NOT NULL` | | Supplier business code (still a **string** here — this table was *not* int-FK-refactored; it carries `VC_SUPPLIER_CODE`/`VC_PART_NUMBER` strings, unlike `INV_PARTS_STOCK_MST`). |
| `VC_PART_NUMBER` | `varchar(12) NOT NULL` | | The part — **the join key the stock-qty trigger uses** (`PS.VC_PART_NUMBER = i.VC_PART_NUMBER`). |
| `VC_FRS_NUMBER` | `varchar(7) NOT NULL` | | Firm release schedule number. |
| `VC_RENBAN_NUMBER` | `varchar(8) NOT NULL` | (key) | **Toyota lot/sequence number — the business key this whole module updates by.** All four `UPDATE_Order*` procs filter `WHERE VC_RENBAN_NUMBER = @RenbanNumber`. Width 8 = file field width `Renbanl=8`. |
| `IN_QTY` | `int NOT NULL` | | Order quantity. **Not changed here**, but it is the value the trigger adds/subtracts to `INV_PARTS_STOCK_MST.IN_QTY`. |
| `VC_STATUS_SUPPLIER_SHIPPING` | `varchar(8) NOT NULL` DEFAULT `''` | ✓ (InTransit) | Set by `UPDATE_OrderShipping` from the file ship date/DT. Non-blank ⇒ "shipped". |
| `VC_ARRIVAL` | `varchar(8) NOT NULL` DEFAULT `''` | | Arrival date — **the sole column the trigger's arrival-add branch keys on** (`d.VC_ARRIVAL=''→non-blank`). **Set by *other* modules, never by any of this module's four procs** — which is why `ARRIVED*` lines here move no stock (§4). |
| `VC_TRAILER_NUMBER` | `varchar(11) NOT NULL` DEFAULT `''` | ✓ (all 4) | Trailer/equipment number; file field `Equipment` (width 11). |
| `VC_STATUS_PLANT_YARD` | `varchar(8) NOT NULL` DEFAULT `''` | ✓ (Arrived) | Set by `UPDATE_OrderPLANT` for ArrivedNUMMI / ArrivedMANUF. |
| `VC_PLANT_PARKING` | `varchar(10) NOT NULL` DEFAULT `''` | ✓ (Arrived) | "Lot" / parking — set by `UPDATE_OrderPLANT @Lot` (file field `Lot`, width 5). |
| `VC_STATUS_ASSEMBLER_YARD` | `varchar(8) NOT NULL` DEFAULT `''` | | Read by trigger; not written here. |
| `VC_ASSEMBLER_LOCATION` | `varchar(10) NOT NULL` DEFAULT `''` | | |
| `VC_STATUS_EMPTY_TRAILER` | `varchar(8) NOT NULL` DEFAULT `''` | ✓ (Terminated, conditional) | Set by `UPDATE_OrderTerminated`'s second UPDATE branch when previously blank. |
| `VC_DETENTION` | `varchar(50) NOT NULL` DEFAULT `''` | | |
| `VC_ORDER_DATE` | `varchar(8) NOT NULL` DEFAULT `''` | | |
| `VC_WAREHOUSE` | `varchar(8) NOT NULL` DEFAULT `''` | ✓ (ArrivedCONS/Union) | Set by `UPDATE_OrderWarehouse`. |
| `VC_TERMINATED` | `varchar(8) NOT NULL` DEFAULT `''` | ✓ (Terminated) | Set by `UPDATE_OrderTerminated`. |
| `VC_SHIP_DATE` | `varchar(8) NOT NULL` DEFAULT `''` | | |
| `VC_KANBAN_NUMBER` | `varchar(5) NOT NULL` DEFAULT `''` | | (`_HIST` table misnames this `VC__KANBAN_NUMBER` with a double underscore ⚠️ — a latent `SELECT *`/INSERT-by-position hazard between the two tables.) |
| `VC_FRS_DATE` | `varchar(8) NULL` | | Only **NULLable** non-id column. |
| `VC_LAST_UPDATE` | `varchar(16) NOT NULL` DEFAULT `''` | ✓ (all 4) | **Timestamp as `yyyymmddHHMMSSff` string (P2).** Set in every `UPDATE_Order*` proc as `CONVERT(varchar,getdate(),112)` [8 chars] + **four** `SUBSTRING(CONVERT(varchar,getdate(),114),p,2)` at positions 1/4/7/10 = HH+MM+SS+ff = **16 chars (`yyyymmddHHMMSSff`)**, exactly filling the column — **identical to the master procs** (Supplier/Size/Logistics). |
| `VC_ADD` | `varchar(16) NOT NULL` DEFAULT `''` | | Insert timestamp string (P2); set when the order is first created elsewhere, **never touched here**. |

**Constraints / indexes:** `PK_INV_OPEN_ORDER_INF` PRIMARY KEY CLUSTERED (`IN_ORDER_ID`). A long block of
`DEFAULT ('')` constraints (every status/date `varchar` defaults to empty string — this is why the triggers
test `<> ''` rather than `IS NOT NULL`). **No UNIQUE index on `VC_RENBAN_NUMBER`** — so a Renban is **not**
guaranteed unique, and `SELECT_OrderOpenOrderLog`/the `UPDATE_Order*` procs can match/update **multiple rows**
for one Renban (a normal case: one Renban spans many parts). **No declared FOREIGN KEY** out of this table;
`VC_SUPPLIER_CODE`/`VC_PART_NUMBER` link by convention only.

**Triggers on these tables (authoritative bodies — `DB Schema/Create Inventory.sql`):**
`INV_OPEN_ORDER_INF` carries **three** live triggers (the `RecConfStatPartsStockMstQTY` family). Because this
module fires `UPDATE` on the table, the **`UPDATE_` trigger is on its live path every time**.

- **`UPDATE_RecConfStatPartsStockMstQTY`** (FOR UPDATE, schema ~line 9764) — **the core inventory invariant
  this module triggers.** On any order UPDATE it: (1) `INSERT into INV_OPEN_ORDER_INF_HIST SELECT * from
  deleted` (snapshot the prior row), then runs **eight conditional `UPDATE INV_PARTS_STOCK_MST` statements**
  that add/subtract the order qty to/from `IN_QTY`, joining parts↔order on `VC_PART_NUMBER` and parts↔supplier
  on `IN_SUPPLIER_ID`, gated by the supplier's **`VC_INVENTORY_ADD_POINT`** (`'S'`=shipped or `'A'`=arrived,
  P4). The exact invariant set:
  - **Qty changed while already shipped** (`i.IN_QTY <> d.IN_QTY`, supplier `S`, was shipping): subtract old,
    add new (`IN_QTY = IN_QTY − d.IN_QTY` then `+ i.IN_QTY`) — net delta = `(new − old)` qty.
  - **Qty changed while already arrived** (`i.IN_QTY <> d.IN_QTY`, supplier `A`, `i.VC_ARRIVAL<>''`): same
    subtract-old/add-new pair.
  - **Newly shipped** (`VC_STATUS_SUPPLIER_SHIPPING` goes `'' → non-''`, supplier `S`): **`IN_QTY += i.IN_QTY`**.
  - **Un-shipped** (`VC_STATUS_SUPPLIER_SHIPPING` goes `non-'' → ''`, supplier `S`): **`IN_QTY −= i.IN_QTY`**.
  - **Newly arrived** (`VC_ARRIVAL` goes `'' → non-''`, supplier `A`, `WHERE i.VC_ARRIVAL<>d.VC_ARRIVAL AND
    d.VC_ARRIVAL=''`): **`IN_QTY += i.IN_QTY`**. **Keys solely on `VC_ARRIVAL`** — so this branch is dormant
    for `LogisticsBreakdown`, whose four procs never set `VC_ARRIVAL` (see §4). (Cosmetic: this branch joins
    `inserted i ON PS.VC_PART_NUMBER = i.VC_PART_NUMBER`, a slightly different join form than the other
    branches.)
  - **Un-arrived** (`VC_ARRIVAL` non-blank with the contradictory `i.VC_ARRIVAL='' AND i.VC_ARRIVAL<>''`
    guard — a **dead/never-true predicate**, so this final "un-arrived subtract" branch effectively never
    fires ⚠️ likely a legacy bug). It also stamps `INV_PARTS_STOCK_MST.VC_LAST_UPDATE = i.VC_LAST_UPDATE`.
  **Invariant: stock on hand moves the moment an order's shipping/arrival status flips, by the order qty,
  but only for suppliers whose add-point matches the flip (`S` for shipping, `A` for arrival).** This is the
  invariant the rebuild must re-home as a service transaction / model callback.
- **`INSERT_RecConfStatPartsStockMstQTY`** (FOR INSERT, ~9717): `INSERT … SELECT * from inserted` into
  `_HIST`, then `IN_QTY += i.IN_QTY` for newly-shipped (`S`) or newly-arrived (`A`) rows. **Not on this
  module's path** (this module only UPDATEs), but documented because it shares the table and the same math.
- **`DELETE_RecConfStatPartsStockMstQTY`** (FOR DELETE, ~9660): gated by `Purge.PurgeMode = 0`; reverses qty
  (`IN_QTY −= d.IN_QTY`) for shipping (`S`, status non-blank, not terminated/empty) and arrival (`A`) rows.
  **Not on this module's path.**
  > Note: the `UPDATE`/`INSERT` triggers do **not** check `Purge.PurgeMode`; only the `DELETE` trigger does.
  > These three are also where the obsolete `docs/triggers.sql` and the live schema agree on table but the
  > stale file keys on dropped string columns — **trust the schema** (see
  > [`../cross-cutting/trigger-source-reconciliation.md`](../cross-cutting/trigger-source-reconciliation.md)).
  > Unlike the master `*Code` triggers, this table was **not** int-FK-refactored, so its triggers legitimately
  > still join on `VC_PART_NUMBER`/`VC_SUPPLIER_CODE` (strings) — the live bodies above confirm it.

## 3. Stored procedures used
(Grepped from `LogisticsBreakdown.pas`; read each with `sql.sh proc NAME`. The procs are the behavioral spec.
This module calls the procs **inline** — it does **not** go through named `DataModule.pas` CRUD wrappers, so
the P8/P12 retry-recursion register does **not** list these calls; it uses the shared `Inv_StoredProc` object
directly, **P6**.)

| Proc | Op | Business rule (from body) |
|------|----|---------------------------|
| `SELECT_OrderOpenOrderLog;1 @RenbanNumber varchar(8)` | SELECT | `SELECT * FROM INV_OPEN_ORDER_INF WHERE VC_RENBAN_NUMBER = @RenbanNumber`. Pure existence check by Renban. The form only inspects `IsEmpty` — if no row, it logs **"Renban not in database"** and skips the line; if any row(s) exist it proceeds to the matching status update. (Param `varchar(8)` matches the column.) |
| `UPDATE_OrderShipping;1 @RenbanNumber, @Shipping varchar(8), @Trailer varchar(11)=''` | UPDATE | Sets `VC_STATUS_SUPPLIER_SHIPPING = @Shipping`, `VC_TRAILER_NUMBER = @Trailer`, `VC_LAST_UPDATE` (16-char `yyyymmddHHMMSSff` string). Filter: `WHERE VC_RENBAN_NUMBER = @RenbanNumber AND VC_STATUS_SUPPLIER_SHIPPING <> @Shipping` — **idempotent guard:** a no-op if the status is already that value (so re-running the same file won't re-fire the stock-qty trigger). Fires on file status `INTRANSIT`. |
| `UPDATE_OrderPLANT;1 @RenbanNumber, @PLANT varchar(8), @Trailer varchar(11)='', @Lot varchar(4)=''` | UPDATE | Sets `VC_STATUS_PLANT_YARD=@PLANT`, `VC_TRAILER_NUMBER=@Trailer`, `VC_PLANT_PARKING=@Lot`, `VC_LAST_UPDATE`. Filter `… AND VC_STATUS_PLANT_YARD <> @PLANT` (idempotent). Fires on `ARRIVEDMANUF` and `ARRIVEDNUMMI`. ⚠️ **`@Lot varchar(4)` is narrower than the source field width `Lotl=5` and the column `VC_PLANT_PARKING varchar(10)`** — silent truncation of the Lot/parking to 4 chars (P-truncation). |
| `UPDATE_OrderWarehouse;1 @RenbanNumber, @Warehouse varchar(8), @Trailer varchar(11)=''` | UPDATE | Sets `VC_WAREHOUSE=@Warehouse`, `VC_TRAILER_NUMBER=@Trailer`, `VC_LAST_UPDATE`. Filter `… AND VC_WAREHOUSE <> @Warehouse` (idempotent). Fires on `ARRIVEDCONS` and `ARRIVEDUNION` (both route here). |
| `UPDATE_OrderTerminated;1 @RenbanNumber, @Terminated varchar(8), @Trailer varchar(11)=''` | UPDATE | **Two UPDATE statements:** (1) where `VC_STATUS_EMPTY_TRAILER <> ''` sets `VC_TERMINATED`, trailer, `VC_LAST_UPDATE`; (2) where `VC_STATUS_EMPTY_TRAILER = ''` sets `VC_TERMINATED`, **also** `VC_STATUS_EMPTY_TRAILER = @Terminated`, trailer, `VC_LAST_UPDATE`. Both filter `VC_TERMINATED <> @Terminated` (idempotent). Fires on `TERMINATED`. |

**Not called here but on the same table (context):** `INSERT_OpenOrder`, `UPDATE_OrderQty`,
`UPDATE_OrderRenban`, `UPDATE_OrderWarehouse`, `UPDATE_OrderPLANT`, etc. are the Order module's procs; this
processor reuses a subset (the four status setters) to ingest the carrier feed.

### Call mechanism (legacy)
All work is in `TLogisticsBreakdown_Form.FormShow` (lines 61-248), driving the shared
**`Data_Module.Inv_StoredProc`** ADO object directly (**P6** — no per-entity DataModule wrapper):
1. `Hist.Append('Start File Process')` + `LogActLog('LOGISTICS','Start processing file,'+fFileName)`.
2. `AssignFile(fcf, fFileName); Reset(fcf);` then `while not Seekeof(fcf)` read each line `fcl`.
3. Per line, inside `try…except`: set `ProcedureName := 'dbo.SELECT_OrderOpenOrderLog;1'`,
   `@RenbanNumber := copy(fcl, Renban=1, Renbanl=8)`, `Open`. On `Inv_Connection.Errors.Count > 0` →
   `ShowMessage` + `LogActLog('ERROR',…)` + `raise EDatabaseError`.
4. `if not IsEmpty` → `Close`, then branch on `trim(copy(fcl, Status=28, Statusl=12))` against the six status
   constants and assign the matching `UPDATE_Order*` proc + params, then `ExecProc` (with the same post-exec
   error check). `else` (Renban absent) → log "Renban not in database" and skip.
5. The per-line `except on e:exception` swallows the error: `Hist.Append('Unable to update Order record, …')`
   + `LogActLog('ERROR',…)` and a stubbed-out `// Create Error report<<<<`, then **continues to the next
   line** (one bad line does not abort the batch).
6. After EOF: `CloseFile(fcf)`, then **`CopyFile.CopyFrom := filename`** to archive — see the bug below.
7. Outer `except` logs "Unable to load logistics". Always ends with `Hist.Append('End File Process')` +
   `LogActLog` and reveals the `OK` button. `OK_ButtonClick` just `Close`s.

**Fixed-width field layout (1-based start, length — from the unit's `const` block):**
`Renban` @1 len 8 · `Equipment` @9 len 11 · `DT` @20 len 8 · `Status` @28 len 12 · `Lot` @40 len 5.
The status keyword occupies columns 28-39; the **Lot** is only read for `ARRIVED*` lines and only when the
**total line length is exactly 43 or 44** (a magic-number gate — for `ARRIVEDMANUF` both 43 and 44 are
accepted; for `ARRIVEDNUMMI` only 44; otherwise `@Lot := ''`). The `DT` field (cols 20-27) is reused as the
status *value* written to whichever status column the keyword selects (ship date / plant code / warehouse code
/ terminated code).

## 4. Business rules & edge cases
- **Renban is the only key.** Every update is `WHERE VC_RENBAN_NUMBER = @RenbanNumber`. Because the table has
  **no unique index on Renban**, one file line can update **many order rows** (all parts of that lot). This is
  intended (a lot ships as a unit), but means the stock-qty trigger may move qty for **every part in the lot**
  in one statement.
- **Status keyword → proc routing (coded enums, P4):** the file's status field maps to:
  `INTRANSIT → UPDATE_OrderShipping` · `ARRIVEDMANUF → UPDATE_OrderPLANT` (with Lot) ·
  `ARRIVEDNUMMI → UPDATE_OrderPLANT` (with Lot) · `ARRIVEDCONS → UPDATE_OrderWarehouse` ·
  `ARRIVEDUNION → UPDATE_OrderWarehouse` · `TERMINATED → UPDATE_OrderTerminated`. An **unrecognized status
  falls through all branches** → no `ProcedureName` reassigned → it **re-executes whatever proc was last set**
  (still `SELECT_OrderOpenOrderLog` from step 3, harmless `ExecProc` on a SELECT proc) ⚠️ a latent quirk; an
  unknown status is silently ignored rather than logged as "unknown".
- **Idempotency by design.** Every `UPDATE_Order*` filters `AND <statuscol> <> <newvalue>`, so re-processing
  the same file (or a duplicate line) is a no-op and **does not double-move stock** via the trigger. This is the
  legacy's only guard against double-counting — there is **no file-level dedupe, no processed-marker, no
  transaction**.
- **Stock-quantity effect (the headline):** this module never touches `INV_PARTS_STOCK_MST` directly, but its
  UPDATE on `INV_OPEN_ORDER_INF` fires `UPDATE_RecConfStatPartsStockMstQTY`. **Via this module, the *only* path
  that moves `IN_QTY` is the `INTRANSIT` line for an `S`-supplier part:** `UPDATE_OrderShipping` flips
  `VC_STATUS_SUPPLIER_SHIPPING` blank→non-blank, firing the trigger's "newly shipped" branch (add-point `S`),
  which **adds** the order qty to stock (and a flip back to blank **subtracts** it). **The `ARRIVED*` lines do
  NOT move stock through this module.** The trigger's only arrival-**add** branch keys on `VC_ARRIVAL` flipping
  blank→non-blank (`WHERE i.VC_ARRIVAL <> d.VC_ARRIVAL AND d.VC_ARRIVAL = ''`), but **none of the four procs
  this module calls ever sets `VC_ARRIVAL`** — `UPDATE_OrderPLANT` sets `VC_STATUS_PLANT_YARD`,
  `UPDATE_OrderWarehouse` sets `VC_WAREHOUSE`, neither touches `VC_ARRIVAL`. So the add-point `'A'` arrival
  branch is **effectively dormant for this module**, and `ARRIVED*` / `TERMINATED` / warehouse lines set their
  status columns **without moving `IN_QTY`**.
- **`VC_ARRIVAL` is the sole arrival key the trigger watches (load-bearing):** the entire add-point `'A'`
  add/subtract logic in `UPDATE_RecConfStatPartsStockMstQTY` is gated on `VC_ARRIVAL` (blank↔non-blank), **not**
  on `VC_STATUS_PLANT_YARD` or `VC_WAREHOUSE`. Because no proc on this module's path writes `VC_ARRIVAL`, the
  carrier feed's arrival lines are **decoupled from any stock movement** here — arrival-driven stock changes
  must come from some *other* module that actually stamps `VC_ARRIVAL`.
- **Timestamp format (P2):** the `UPDATE_Order*` procs write `VC_LAST_UPDATE` as a **16-char**
  `yyyymmddHHMMSSff` string (`CONVERT(…,112)` [8] + four `SUBSTRING(CONVERT(…,114),p,2)` at 1/4/7/10 =
  HH+MM+SS+ff), exactly filling the `varchar(16)` column — **identical to the master procs**, no width
  mismatch. The trigger then copies this same 16-char value onto `INV_PARTS_STOCK_MST.VC_LAST_UPDATE`.
- **Lot truncation (silent):** `@Lot varchar(4)` < file field `Lotl=5` < column `VC_PLANT_PARKING varchar(10)`
  — the parking/Lot is silently truncated to 4 chars on write.
- **Plant label is from INI, not the file:** the history/log messages use `Data_Module.fiPlantName.AsString`
  (`[SITE] PlantName`, default `NUMMI`) — the same UPDATE proc (`UPDATE_OrderPLANT`) is used for both
  `ARRIVEDMANUF` and `ARRIVEDNUMMI`; only the **log text** differs, driven by the site's configured plant name.
- **Per-line error isolation:** a DB error or parse error on one line is caught, logged, and the batch
  continues. The `// Create Error report<<<<` stub shows an intended-but-unbuilt error report.
- **Archive step is broken (bug):** the form sets `CopyFile.CopyFrom := filename` but **never sets
  `CopyFile.CopyTo`** (contrast `EDIUpload.pas:424-450`, which sets `CopyTo := …\Archive\…`). With
  `MoveFile = True` and no destination, the intended "move the file to Archive" **does not happen** (no-op or
  silent failure) — so processed logistics files are **not archived/removed**, risking **re-processing the same
  file** on the next run. (The idempotent proc guards prevent double stock-movement, but stale files accumulate.)
  Statement also lacks a terminating `;` before the `except` (it compiles because it is the last statement of
  the `try` block).
- **No confirmation / no preview:** processing starts on `FormShow`; the operator only sees the `THistory` log
  scroll and an `OK` button afterward. There is no way to cancel mid-batch or dry-run.

## 5. UI / UX notes
- **Two-screen flow:** the shared `UploadBreakDown` hub (mode `bReceiving`) handles **file selection** (Browse
  dialog, INI-seeded dir/filename, `*.txt`/`*.log` filter); `LogisticsBreakdown` is the **processing window** —
  a scrolling `THistory` log + a post-completion `OK` button. No grids, no editable fields, no search.
- **What to keep:** the line-by-line audit log (`Hist.Append` + `LogActLog('LOGISTICS',…)`) is genuinely
  useful operator feedback — preserve it as a job log / progress stream. The idempotent re-run safety is worth
  keeping as an explicit guard.
- **What to modernize:**
  - Replace flat-file ingest with an **upload endpoint + background job** (the file is a batch; parse server-side).
    A true X12 856 inbound parser is the eventual target (this flat file is a poor-man's 856 status feed).
  - **Fix the archive bug:** move/copy the consumed file to a per-site configured archive target (or mark it
    processed in the DB) — never leave it in place. Replace the local `[DIRECTORIES] LogisticsInputDir`
    Windows path (multi-site lens — see §8).
  - Surface a **summary report** of the batch (the `// Create Error report` the legacy never built): counts of
    updated / "Renban not found" / errored lines, and the resulting stock-qty deltas.
  - Make the fixed-width parse a declarative spec (field offsets/lengths), and **log unrecognized status
    keywords** instead of silently ignoring them.
  - Validate line length / field shape up front rather than via the `length = 43/44` magic numbers.

## 6. Target design  *(Rails primary; Python for the parser)*
- **No new owning model** — this is a *processor*, not a master. It operates on the existing
  `OpenOrder` model (`self.table_name = 'INV_OPEN_ORDER_INF'`, `self.primary_key = 'IN_ORDER_ID'`).
  - `OpenOrder` associations (by convention, no declared FK): `belongs_to :supplier, primary_key:
    'VC_SUPPLIER_CODE', foreign_key: 'VC_SUPPLIER_CODE'`; `belongs_to :part_stock, primary_key:
    'VC_PART_NUMBER', foreign_key: 'VC_PART_NUMBER'` (both string keys — this table was not int-refactored).
  - Status columns become an explicit status/lifecycle concern; the supplier's `VC_INVENTORY_ADD_POINT`
    (`S`/`A`) is a Rails `enum` (P4) consulted by the stock-move service.
- **Service object (the core):** `LogisticsStatusIngestService` (or a `Python` parser feeding it) that, per
  line: parses the fixed-width fields, looks up open-order rows by Renban, and applies the status change **inside
  one transaction per line**, mirroring the idempotent `AND <col> <> <value>` guard. The **stock-qty trigger
  must be re-homed here**, not left in the DB long-term: an `OpenOrder` after-update callback / a
  `StockBalanceService.apply(order, prior_status, new_status)` that adds/subtracts `IN_QTY` on the matching
  `PartStock` per the eight-branch logic above, gated by the supplier add-point. **During parallel run, keep the
  DB trigger and do not double-apply in app code** (stage 1-2); only move the logic into the service when the
  trigger is dropped (stage 3, Postgres).
- **Controller/routes:** a non-RESTful action — `POST /logistics_status_ingests` (upload) enqueues a job;
  `GET /logistics_status_ingests/:id` shows the job log/summary. The actual `UPDATE_Order*` writes go through
  the service, not a CRUD controller.
- **Views:** an upload form + a job-progress / log view (replacing the `THistory` pane) + a batch summary.
- **Background job:** parse + apply runs async (`ActiveJob`/Sidekiq, or a Python worker for the parser); stream
  the per-line log to the job view. Archive the file to object storage / per-site share on success.
- **Reports:** none existing; **add** the batch summary the legacy stubbed.

## 7. Migration plan for this module
- [ ] **Stage 1 — read-only parity / shadow:** parse the same fixed-width file server-side and produce the
      **planned** list of `UPDATE_Order*` calls + the "Renban not found" set, **without writing**. Diff against
      a legacy run on the same file to prove the parser (offsets 1/9/20/28/40, lengths 8/11/8/12/5; the 43/44
      length gate; status→proc routing) is byte-exact.
- [ ] **Stage 2 — writes via the existing procs/trigger:** call `UPDATE_OrderShipping / UPDATE_OrderPLANT /
      UPDATE_OrderWarehouse / UPDATE_OrderTerminated` through `tiny_tds`, **letting the live
      `UPDATE_RecConfStatPartsStockMstQTY` trigger move stock** (do not reimplement the qty math yet). Preserve
      the idempotent guards and per-line error isolation. **Fix the archive bug** (set a real destination /
      mark processed). Keep `LogActLog('LOGISTICS',…)` behavior as app logging.
- [ ] **Stage 3 — reimplement (Postgres-ready):** move the eight-branch stock-qty logic out of the DB trigger
      into `StockBalanceService` (a model callback/service transaction), gated by the supplier add-point enum;
      replace the flat-file parser with a maintained spec (or a real EDI 856 parser); add a unique/dedupe
      strategy for processed files; build the batch summary report. **Drop the dead
      `i.VC_ARRIVAL='' AND i.VC_ARRIVAL<>''` branch** (it never fires) rather than porting it.

## 8. Open questions for the user (domain expert)
1. **What system emits this file, and is it really a substitute 856?** The format (Renban + equipment + DT +
   status keyword + Lot, fixed-width) looks like a carrier/yard status feed. Is it a Toyota/transport-provider
   export, and should the rebuild ingest the **real EDI 856** instead of this flat file (the menu even labels
   this "Receiving")? Confirm the canonical status keyword set — the code recognizes exactly
   `INTRANSIT / ARRIVEDNUMMI / ARRIVEDMANUF / ARRIVEDUNION / ARRIVEDCONS / TERMINATED`; are there others
   (an unknown keyword is currently silently ignored)?
2. **Stock-qty semantics confirmation:** via *this* module, the **only** stock movement is an `INTRANSIT`
   (shipped) line **adding** stock for `VC_INVENTORY_ADD_POINT = 'S'` suppliers (the trigger's shipping-add
   branch). The `ARRIVED*` lines **do not** move stock here, because the procs set `VC_STATUS_PLANT_YARD` /
   `VC_WAREHOUSE` and **never set `VC_ARRIVAL`**, which is the only column the trigger's arrival-add branch
   watches. So where does the `'A'`-supplier *arrival* stock add actually happen — which other module stamps
   `VC_ARRIVAL`? And is it intended that for `'A'`-parts this carrier feed records arrival status **without**
   counting stock (stock only counted by that other arrival path)? We need this pinned before re-homing the
   trigger.
3. **The dead "un-arrived" trigger branch** (`WHERE i.VC_ARRIVAL = '' AND i.VC_ARRIVAL <> ''`) can never be
   true. Was there meant to be a real "arrival reversed → subtract stock" path? Should the rebuild implement
   the intended reversal, or is reversal genuinely not supported?
4. **Renban is not unique** on `INV_OPEN_ORDER_INF` (no unique index); one status line updates every order row
   for that lot, moving stock for **all** its parts. Is that the intended granularity, or should status be
   tracked per part/per Renban+part?
5. **Archive bug:** processed files are not actually moved (no `CopyTo`). Were operators supposed to get an
   archived copy? Where should consumed logistics files go in a multi-site web deployment (per-site share / SFTP
   / object store), and do we need a **processed-file ledger** to prevent re-ingest (the idempotent procs stop
   double stock-movement, but not re-reading a stale file)?
6. **Multi-site:** the input path is `[DIRECTORIES] LogisticsInputDir` (default `c:\_Inventory_Control\`) and
   the plant label is `[SITE] PlantName` (default `NUMMI`) — both single-site INI values. When multi-site, each
   site has its own carrier feed and plant name; confirm the ingest is **per-site** and `INV_OPEN_ORDER_INF`
   needs a `site_id` scope (it has none today). Renban/part keys are global strings — are they
   guaranteed unique across sites?
7. **`Lot`/`VC_PLANT_PARKING` truncation:** the proc caps `@Lot` at `varchar(4)` though the file field is 5 and
   the column is 10. Is 4 chars correct, or is plant-parking being silently truncated?

## 9. Test cases / parity checks
- **`INTRANSIT` line, S-supplier part, qty Q, status currently blank** → `VC_STATUS_SUPPLIER_SHIPPING` set to
  the file DT; trailer set; `VC_LAST_UPDATE` stamped (16-char `yyyymmddHHMMSSff`); and
  **`INV_PARTS_STOCK_MST.IN_QTY` increases by Q** for that part (trigger "newly shipped" branch). A history row
  is written to `INV_OPEN_ORDER_INF_HIST`.
- **Same `INTRANSIT` line re-processed (duplicate / file re-run)** → `UPDATE_OrderShipping`'s
  `AND VC_STATUS_SUPPLIER_SHIPPING <> @Shipping` makes it a **no-op**; **`IN_QTY` does NOT change again**
  (no double-count). Assert exact stock parity vs a single run.
- **`ARRIVEDNUMMI` line, A-supplier part, was not arrived, line length 44** → `UPDATE_OrderPLANT` sets
  `VC_STATUS_PLANT_YARD`, trailer, `VC_PLANT_PARKING = Lot` (≤4 chars), `VC_LAST_UPDATE`; but **`IN_QTY` does
  NOT change.** The proc sets `VC_STATUS_PLANT_YARD`, **not** `VC_ARRIVAL`, and the trigger's "newly arrived"
  add branch keys on `VC_ARRIVAL` flipping blank→non-blank — so no stock moves. With line length 43 for
  `ARRIVEDNUMMI`, `@Lot = ''` (only `ARRIVEDMANUF` accepts 43); verify the Lot-vs-no-Lot length gating.
- **`ARRIVED*` line for *any* supplier (S or A)** → status columns update, but **`IN_QTY` does NOT change** —
  no proc on this path sets `VC_ARRIVAL`, so the trigger's arrival branch never fires here regardless of
  add-point. Conversely an `INTRANSIT` line for an `A`-supplier part updates shipping status but does **not**
  move stock (the shipping-add branch requires add-point `S`). **Only `INTRANSIT` + `S` moves stock via this
  module** — these cases pin the add-point invariant *and* the `VC_ARRIVAL`-decoupling of arrival lines.
- **`TERMINATED` line** → `UPDATE_OrderTerminated` sets `VC_TERMINATED` (and `VC_STATUS_EMPTY_TRAILER` if it was
  blank) via its two-statement body; **no stock movement** (terminated doesn't cross the ship/arrival boundary).
- **Qty-change replay (same Renban, different `IN_QTY` while already shipped, S-supplier)** → trigger subtracts
  old qty and adds new ⇒ `IN_QTY` net delta = `(new − old)`. (Note: this requires `IN_QTY` to differ on the
  UPDATE; the status-only updates this module performs normally keep `IN_QTY` equal, so this branch is exercised
  by Order-qty edits, not the logistics feed — included for trigger completeness.)
- **Renban absent from `INV_OPEN_ORDER_INF`** → line skipped, log "Renban not in database, R:… DT:…", **no
  writes, no stock change**.
- **Malformed line (parse/DB error)** → caught per-line, logged "Unable to update Order record, …", batch
  **continues** to the next line; assert other valid lines in the same file still apply.
- **Unrecognized status keyword** → legacy silently re-`ExecProc`s the last-set (SELECT) proc, effectively a
  no-op with no "unknown status" log. New app must (per §8.1) **log** the unknown keyword; assert it does not
  silently update any column.
- **Archive behavior** → confirm the **legacy does NOT move the file** (no `CopyTo`); the rebuilt job must
  archive/mark the file and must not re-ingest it (or, if re-ingested, must not double-move stock thanks to the
  idempotent guards).
