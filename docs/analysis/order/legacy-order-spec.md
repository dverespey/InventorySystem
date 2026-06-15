# Legacy Behavioral Spec — Order ("What to Order")

Source of truth for the daily ordering tool. Status: **LIVE** — `InventorySystem.dpr:8`
(`Order in 'Order.pas' {Order_Form}`). Form is a thin Select-Order dialog
(`Order.dfm`) driving an Excel-OLE simulation engine (`Order.pas`, ~62 KB) plus
SQL Server stored procs (`DB Schema/Create Inventory.sql`, UTF-16LE) and triggers
(`docs/triggers.sql`, UTF-16LE).

Confidence: HIGH on Pascal flow and on every proc body read below. Items marked
**body unverified** were not read (the proc lives in a different database, or
the rule lives only inside the `.xls` template).

---

## 1. Overview

The dialog (`Order.dfm`) collects: **Line** (`Line_ComboBox`), **Part Type**
(`PartType_ComboBox`: TIRE/WHEEL/VALVE/FILM), **Sort By** (`SortBy_ComboBox`,
hidden unless Line is blank — `Order.pas:1659`), a read-only **Today** date, and
two action buttons: **Start** and **Order**.

- **Start** (`Start_ButtonClick`, `Order.pas:146`) opens the template
  `OrderSimulation.xls` via Excel OLE, builds a time-phased grid (row 5 = dates,
  row 6 = weekday; parts down the rows), fills forecast/leadtime/inventory/
  in-transit/open-order data, lays Excel formulas, and color-codes cells. It does
  **not** touch the database for writes; it only reads.
- **Order** (`ProcessOrder_ButtonClick`, `Order.pas:628`) reads the sheet back
  and **creates real open-order rows** via `INSERT_OpenOrder`, assigning FRS and
  Renban numbers.

Core layout constants (`Order.pas:89-106`): `ForecastCol=20` (col T, day-0 of the
phased grid is `DateWeekCol+1=20`), `DateWeekCol=19` (col S = the "Week/usage/
balance" label column), `DateWeekLetter='T'`, `QtyCol=17` (Q), `LotCol=18` (R),
inventory col K=11, plant L=12, in-transit M=13, open-order N=14, lot-size Q col
15(O), leadtime col 16(P), usage H=8/safety I=9/J=10. After the phased grid, three
summary columns are written at `DateWeekCol+FillDays+1..+3` = **Total Inv / In
Transit / Added Leadtime** (`Order.pas:201-203`).

Window sizing (`Execute`, `Order.pas:109-113`): `FillDays := fiFillDays`
(INI `[INIT] FillDays`, default **23**, `DataModule.dfm:630`).
`DateRangeCount := FillDays*3` (so 69 by default). The phased grid shows exactly
`FillDays` **production** days; `DateRangeCount`/`fDateRange` bounds the calendar
scan that maps those production days back onto real dates. Day arrays are sized
`[0..200]` (`Order.pas:48-51`) — hard cap, no overflow guard. Part rows are
`fpartline[1..200,...]` (`Order.pas:46`) — **max 200 part rows** (silent cap).

---

## 2. Inputs — every proc / dataset called

ADO objects used: `ALC_StoredProc` (Connection = **ALC_Connection**, catalog
`TireOrder` — a *different* database, `DataModule.dfm:532`), `Inv_StoredProc`,
`Inv_DataSet`, `INV_Forecast_DataSet`, and `INV_Order_StoredProc` (all on
**Inv_Connection**). All proc calls use the `;1` group suffix.

### START path (read-only), in call order

| # | Proc / dataset | Object | Signature (params) | Tables read | Exists? |
|---|---|---|---|---|---|
| 1 | `AD_GetSpecialDate` | ALC_StoredProc | `@BeginDate, @EndDate, @LineName` → result set with `DATE`, `Date Status Abrv` | (ALC `TireOrder` DB) | **NOT in this schema** — cross-DB (see Hazards) **body unverified** |
| 2 | `SELECT_PartsStockInfoOrder;1` | Inv_StoredProc | `@LineName varchar(10), @PartType varchar(50), @SortType varchar(50)` | `INV_PARTS_STOCK_MST` JOIN `INV_PART_TYPE_MST`,`INV_SIZE_MST`,`INV_SUPPLIER_MST`, LEFT JOIN `INV_RENBAN_GROUP_MST` | YES `schema:7382` |
| 3 | `SELECT_SupplierInfo;1` | Inv_DataSet | `@SupCode varchar(5)=''` (note: code passes `@SupCode`; proc declares `@SupCode`,`@Logistics`) | `INV_SUPPLIER_MST` LEFT JOIN `INV_LOGISTICS_MST` | YES `schema:7978` |
| 4 | `SELECT_SizeInfo;1` | INV_Forecast_DataSet | `@SizeCode varchar(6)=''` → `Daily Usage`,`Safety Days`,`RecordID`… | `INV_SIZE_MST` | YES `schema:7869` |
| 5 | `SELECT_OrderHistory;1` | INV_Forecast_DataSet | `@PartNumber varchar(12)` → `Qty` | `INV_FORECAST_DETAIL_INF`, `INV_OPEN_ORDER_INF` | YES `schema:6757` |
| 6 | `SELECT_ForecastDetailTWPN;1` | INV_Forecast_DataSet | `@PartNumber, @EffMonth varchar(7), @TireWheel varchar(1), @IncludeZero bit=1` → `IN_TIRE_RATIO`/`IN_WHEEL_RATIO` | `INV_FORECAST_DETAIL_INF` | YES `schema:6228` |
| 7 | `SELECT_UsageDay;1` | INV_Forecast_DataSet | `@Date varchar(8), @PartNo varchar(12)` → `Qty` | `INV_PART_SHIPPING_INF` | YES `schema:8088` |
| 8 | `SELECT_FirstProductionDay;1` | INV_Forecast_DataSet | `@ProdYear varchar(4)=''` → `First Week Number` | `INV_FIRST_PRODUCTION_DAY` | YES `schema:5982` |
| 9 | `SELECT_ForecastPartNumberWeek;1` | INV_Forecast_DataSet | `@WeekNo int, @DayNo int, @PartNo varchar(12)` → `Qty` (returns IN_QTY1..7 by day) | `INV_BREAKDOWN_FC_INF` | YES `schema:6309` |
| 10 | `SELECT_OrderAtassembler;1` | INV_Forecast_DataSet | `@PartNumber` → `Qty` (note proc is `SELECT_OrderAtASSEMBLER`) | `INV_OPEN_ORDER_INF` | YES `schema:6643` |
| 11 | `SELECT_OrderAtPLANT;1` | INV_Forecast_DataSet | `@PartNumber` → `Qty` | `INV_OPEN_ORDER_INF` | YES `schema:6700` |
| 12 | `SELECT_OrderInTransit;1` | INV_Forecast_DataSet | `@PartNumber, @FirstFRS varchar(8)='00000000'` → `Qty` | `INV_OPEN_ORDER_INF` | YES `schema:6816` |
| 13 | `SELECT_OrderInTransitList;1` | INV_Forecast_DataSet | `@PartNumber, @FirstFRS varchar(8)` → rows (`VC_FRS_DATE`,`IN_QTY`) | `INV_OPEN_ORDER_INF` | YES `schema:6850` |
| 14 | `SELECT_OrderOpenOrder;1` | INV_Forecast_DataSet | `@PartNumber` → `Qty` | `INV_OPEN_ORDER_INF` | YES `schema:6955` |
| 15 | `SELECT_OrderOpenOrderList;1` | INV_Forecast_DataSet | `@PartNumber, @FirstFRS varchar(8)` → rows | `INV_OPEN_ORDER_INF` | YES `schema:6985` |

(`SELECT_SizeInfo` is opened **twice** back-to-back at `Order.pas:935-941` — the
second open is dead/no-op; only the first is consumed.)

### ORDER path (writes), in call order — `Order.pas:628`

| # | Proc | Object | Signature | Effect | Exists? |
|---|---|---|---|---|---|
| A | `SELECT_PartsStockRenban;1` | Inv_StoredProc | `@PartNum varchar(12)` → `IN_RENBAN_COUNT` | read current renban counter | YES `schema:7506` |
| B | `UPDATE_PartsStockRenban;1` | Inv_StoredProc | `@PartNum varchar(12), @RenbanCount int` | bumps `INV_PARTS_STOCK_MST.IN_RENBAN_COUNT`, sets `VC_LAST_UPDATE` | YES `schema:9018` |
| C | `INSERT_OpenOrder;1` | INV_Order_StoredProc | `@SupCode varchar(5), @PartNum varchar(12), @KanbanNum varchar(5), @FRSNum varchar(7), @RenbanNum varchar(8), @Qty int` | **inserts into `INV_OPEN_ORDER_INF`**; computes FRS year-roll + FRS sequence suffix server-side | YES `schema:3236` |

`INSERT_OpenOrder` has **no OUTPUT params and no return code** the form reads
(`ExecProc` only). All idempotency/sequence logic is server-side (see §6).
Each insert is wrapped in `Inv_Connection.BeginTrans/CommitTrans` with rollback on
exception (`Order.pas:686/757`, `780/854`).

---

## 3. The simulation algorithm (Start), step by step

### 3.1 Calendar / production-day mapping (`Order.pas:209-339`)
1. Call `AD_GetSpecialDate(@BeginDate=now, @EndDate=now+fDateRange, @LineName)` on
   the ALC DB to get per-day status across the window.
2. Zero `fDates/fOvertimes/fForecast/fNonProduction[0..fDateRange]`
   (`Order.pas:231-237`).
3. Walk forward day by day (`x` = calendar offset from today) until **`fFillDays`
   reaches `FillDays`** (default 23). For each day:
   - If the special-date row matches the day:
     - status **`'O'` (Overtime)** → render the date, record it as a production
       day, AND push its 1-based fill index into `fOvertimes[]`, `INC(fOvertimeCount)`
       (`Order.pas:256-266`).
     - status **`'X'` (Non-Production)** → render the date, count it as a fill day,
       push into `fNonProduction[]`, `INC(fNonProductionCount)` (`:267-277`).
     - (other statuses / `next`) → effectively a **holiday: skipped** (no render,
       fill index not advanced).
   - If no special-date match: render the day only if
     `DayOfTheWeek(now+x) < 6` (Mon–Fri); weekends are skipped (`:282-289`).
   - If the special-date result set is empty (`recordcount=0`): "normal run" —
     pure Mon–Fri loop (`:313-325`).
   `fDates[x]` stores the Excel serial date for each rendered production day;
   note it is indexed by **calendar offset x**, not by fill position.

So "fill days" = number of *production* (worked) days shown; overtime days count
as fill days and additionally extend lead time (see §3.4); X-days count as fill
days but are flagged non-production; holidays consume calendar but not grid
columns.

### 3.2 Part rows + per-row data (`Order.pas:345-578`)
Driven by `SELECT_PartsStockInfoOrder` ordered by size. Rows are grouped by
`VC_SIZE_CODE`; a size header row writes `Beg Balance` and calls `UpdateSizeInfo`
(`Daily Usage`→H, `Safety Days`→I, `J=H*I` safety stock; H:I colored 34). When the
size changes, the prior block is closed with `Usage`/`End Balance` rows, borders,
and `DoFormulas`.

Per part row `i`:
- Supplier name (C), part number (D), `IN_1LOTQTY` (O=15).
- **Lead time selection** (`Order.pas:426-459`): `case DayOfTheWeek(now)` →
  Mon=`IN_LEADTIME_MONDAY`, Tue=`_TUESDAY`, … Sat=`_SATURDAY`; **if the weekday
  column is 0, fall back to `IN_LEADTIME`** (the `else` of the case also uses
  `IN_LEADTIME`). Sunday (7) → `IN_LEADTIME`. Result written to P=16.
- Total inventory `IN_QTY` → `Total Inv` col (`DateWeekCol+FillDays+1`).
- Order share %: `OrderHistory` writes per-part order qty into E; if one part in
  size group → 100 %, else an Excel `=E/(ΣE)` formula across the group
  (`:467-499`).
- `UpdateFRSInfo`: tire/wheel ratio → G (`/100`); also builds a usage-vs-forecast
  compare row using `SELECT_UsageDay` over the last `fiForecastUsageCompare`
  (INI `[INIT] ForecastUsageCompare`, default **7**, `DataModule.dfm:310`) days.
- `ForecastHistory`: prior weeks' forecast via `SELECT_ForecastPartNumberWeek`,
  optionally week-offset by `SELECT_FirstProductionDay` when
  `[INIT] UseFirstProductionDay` is true (`DataModule.dfm:374`).
- `K = TotalInv - (L+M)` available formula; `PutASSEMBLERCount`→K,
  `PutPLANTCount`→L, `PutIntransitCount`→M, `PutOpenOrderCount`→N.
- `FillForecast` writes the daily forecast array `fForecast[0..fFillDays-1]` into
  the phased columns (`Order.pas:1490-1498`). `UpdateForecast`
  (`Order.pas:1135`) accumulates `fForecast[]` per day across the size group,
  using the week/day breakdown table and the same first-production-day offset.

### 3.3 Open-order / in-transit phasing (`PutIntransitCount` 1262, `PutOpenOrderCount` 1391)
Both compute `firstFRS = yyyymmdd of first rendered date`, pull the per-FRS-date
list, and bucket `IN_QTY` into the phased column whose `fDates[i]` matches the
FRS date (`StrToDate(copy(lastFRS,5,2)+'/'+copy(...,7,2)+'/'+copy(...,1,4))`,
`Order.pas:1362,1462`). In-transit numbers are stamped **font color 23**; open
orders **font color 10**. The day's bucket index `z` is recomputed by scanning
`fDates` and counting non-zero (production) days up to the match.

### 3.4 Lead-time / order-point math (`DoLeadTime`, `Order.pas:1570`)
1. `addedleadtime := 0`. For each overtime day `i` (`fOvertimes[]`): if
   `(fOvertimes[i]-1) <= (leadtime+i)` then `INC(addedleadtime)` else `break`
   (`:1576-1582`). I.e. overtime days that fall *inside* the lead-time window
   push the order-by point out by one each. Written to `Added Leadtime`
   (`DateWeekCol+FillDays+3`).
2. **Lead-time zone** = phased columns `T` .. `T+(leadtime-1)+addedleadtime`
   colored **Interior 36** (`Order.pas:1587`). The **order-by column** =
   `T+leadtime+addedleadtime` colored **Interior 40** (`:1588`) — this is the day
   you must place the order to arrive in time.
3. Overtime columns recolored **Interior 3 (red)** (`:1590-1593`);
   non-production columns **Interior 4 (green)** (`:1595-1598`) — these override
   the zone colors at those columns.
4. `fPartLine[line,FRSDate] := date in row 5 at the order-by column`
   (`:1600`) — this is the expected-arrival date used to build the FRS number on
   the Order path. If the order-by cell is empty, it seeds `=Q{row}` (lot qty)
   into it; otherwise it locks Q:R (`:1601-1605`).

### 3.5 End-balance projection (`DoFormulas`, `Order.pas:1510`)
For each phased day, `BegBalance(day+1) = EndBalance(day)` and
`EndBalance = Beg + Receipts - Usage` (the exact term count depends on block
height 3/4/5 rows). End-balance cells get a conditional-format rule:
**if value < `$J$topedge` (the safety-stock cell) → font color 3 (red)**
(`Order.pas:1551-1552, 1557-1558, 1563-1564`). This is the stockout/below-safety
alert, set in Pascal.

---

## 4. Color map (exact, from Order.pas)

Excel `ColorIndex` palette (standard Excel 2000 palette):
3=red, 4=bright green, 10=dark green, 23=dark blue, 34=pale cyan, 36=pale
yellow, 40=pale tan/cream.

| Where set | Color | Target | Business meaning |
|---|---|---|---|
| `Order.pas:523` Interior 36 | pale yellow | P (leadtime cell) | leadtime input box highlight |
| `Order.pas:539-540` Interior 40 | cream | Q,R (lot/qty input) | order-entry cells (the cells you type into) |
| `Order.pas:932` Interior 34 | pale cyan | H:I | size daily-usage / safety-days inputs |
| `Order.pas:1587` Interior 36 | pale yellow | T..T+(leadtime-1)+added | **lead-time zone** (days within lead time) |
| `Order.pas:1588` Interior 40 | cream | T+leadtime+added | **order-by point** (place order this day) |
| `Order.pas:1592` Interior 3 | red | overtime columns | **overtime production day** |
| `Order.pas:1597` Interior 4 | bright green | non-production columns | **non-production ('X') day** |
| `Order.pas:1342,1377` font 23 | dark blue | in-transit qty cells | qty is **in transit** |
| `Order.pas:1449,1476` font 10 | dark green | open-order qty cells | qty is an **open (unshipped) order** |
| `Order.pas:1047,1055` FormatCond font 3 | red | J{row+1}/J{row+2} | usage-vs-forecast **over-produced** (J<0) |
| `Order.pas:1552/1558/1564` FormatCond font 3 | red | end-balance cells | projected balance **below safety stock** (`< $J$top`) |
| `Order.pas:394` Interior `XLColToInt('AV')` | (palette idx of "AV"→ via XlColToInt, =48) | spacer row E..end | block separator shading |

There is **no** use of color indices 40-vs-36 to mean anything beyond the above;
fonts vs interiors are distinct channels (qty-source colors are *font*, zone
colors are *interior*).

---

## 5. Extraction gaps — colors/thresholds that live ONLY in OrderSimulation.xls

These are applied by the template, not by Pascal. Do **not** guess thresholds;
they must be extracted from `<TemplateDir>\OrderSimulation.xls`, worksheet 1.

1. **Base template formatting** — header rows 1–6, column widths, the static
   labels, fonts and fills for the fixed columns (B,C,D, E–R headers) are all in
   the template. Pascal only overwrites values/borders for the dynamic region.
   Gap: full cell formatting of rows 1–6 and columns A–R. (template, sheet 1,
   rows 1-6 / cols A-R)
2. **Any FormatConditions pre-baked in the template** on the phased grid
   (`T5:DD200` is `Locked:=True` at `Order.pas:206` but no conditional formats are
   added there in Pascal). If the template carries conditional-format rules on
   the phased range (e.g. color scales, data bars, threshold fills), they are
   invisible to this code. Gap: conditional-format rules + their numeric
   thresholds on `T5:DD200`, and on any total/summary columns
   (`DateWeekCol+FillDays+1..+3`). (template, sheet 1, range T5:DD200 and the 3
   summary columns)
3. **Number formats** (date format of row 5, %-format of the share column F and
   ratio column G, integer formats) — Pascal writes raw values/serials; display
   formatting is template-driven. Gap: cell number formats. (template, sheet 1)
4. **Palette overrides** — if the workbook redefines the standard color palette,
   the index→RGB mapping in §4 shifts. Gap: verify `Workbook.Colors`/palette in
   the template; if customized, re-derive the RGBs for indices 3,4,10,23,34,36,40.

Until the `.xls` is parsed, treat §4 RGBs as the *standard* Excel-2000 palette and
flag that the live palette is unconfirmed.

---

## 6. Order-creation behavior & idempotency (Order path + INSERT_OpenOrder)

For each `fpartline[i]` with a part number, the form reads `Qty` (Q col=17) and
`Lot` (R col=18) back from the sheet (`Order.pas:654-655`), validates both numeric,
skips zero-qty rows.

- **Lot-size orders** (`BIT_LOT_SIZE_ORDERS = TRUE`, `Order.pas:683`): one
  `INSERT_OpenOrder` with `@Qty = Qty`, FRS suffix `…01`
  (`copy(formatdatetime('yymmdd',FRSDate),2,5)+'01'`, `:702`).
- **Non-lot-size (FRS breakdown)** (`:776`): loop `j := 1..Lot`, one
  `INSERT_OpenOrder` per lot with `@Qty = IN_1LOTQTY` (the 1-lot qty, **not** the
  typed Qty), FRS suffix `'0'+j` / `j` (`:796-799`). i.e. lot count → N trailers.
- **Renban**: if the part is **not** in a renban group (`RenbanGroup=''`), read
  `IN_RENBAN_COUNT` via `SELECT_PartsStockRenban`, form
  `@RenbanNum = Kanban + %.3d(count)`, then `UPDATE_PartsStockRenban` with
  `count+1` (wrap >999 → 1) (`Order.pas:709-745`). If the part **is** in a renban
  group, `@RenbanNum=''` and numbering is deferred to the renban/order grouping
  screen.

### `INSERT_OpenOrder` server-side logic (`schema:3236`)
- `@AddDate` = full `yyyymmddhhmmss` timestamp; `@FRSDate` =
  4-digit year + the 4 date digits of `@FRSNum`, **rolling to next year** if the
  FRS month-digit differs from today's (`schema:3256-3264`).
- **FRS sequence dedup** (`schema:3266-3314`): finds `@MaxFRS = max(VC_FRS_NUMBER)`
  for the same FRS prefix and part. Scope of the "same part" differs by renban:
  - `@RenbanNum <> ''` → match on `VC_PART_NUMBER` only.
  - `@RenbanNum = ''` → match across **all parts sharing `IN_RENBAN_ID`** (renban
    group), via subquery on `INV_PARTS_STOCK_MST`.
  If `@MaxFRS` is null → first order, suffix `'01'`. Else suffix =
  `right('00'+(int(right(@MaxFRS,2))+1),2)` — increment the trailing 2 digits.
  (For renban-group orders with a still-empty last renban it reuses `@MaxFRS`.)
  → This `IF EXISTS`-style max+1 is what makes the FRS series safe/idempotent
  across repeated runs. **Note the Pascal-side suffix (`'01'`, `'0'+j`) is
  overwritten/ignored** for the trailing 2 digits — the proc recomputes them
  (`schema:3284,3294,3311`). The Pascal-supplied `@FRSNum` trailing 2 chars are
  effectively dead.
- Inserts row into `INV_OPEN_ORDER_INF` (`VC_SUPPLIER_CODE, VC_PART_NUMBER,
  VC_KANBAN_NUMBER, VC_FRS_NUMBER, VC_FRS_DATE, VC_RENBAN_NUMBER, IN_QTY,
  VC_ADD`) (`schema:3315-3333`). Status columns left default → new order is
  *not yet shipping*.

### Triggers that fire on insert (`docs/triggers.sql`)
`INSERT_RecConfStatPartsStockMstQTY` (`triggers.sql:214`): FOR INSERT —
1. **Always** copies the inserted rows into `INV_OPEN_ORDER_INF_HIST`
   (`SELECT * from inserted`, `:219-220`).
2. Adds `i.IN_QTY` to `INV_PARTS_STOCK_MST.IN_QTY` **only when
   `i.VC_STATUS_SUPPLIER_SHIPPING <> ''`** (`:221-227`). Since a brand-new order
   from this form has empty shipping status, **stock is NOT bumped at order
   creation** — it is bumped later when the order transitions to shipping
   (handled by the UPDATE trigger `:241`, which net-adjusts on qty change while
   shipping). This is the key inventory-coupling fact for migration.

---

## 7. Multi-site / single-site assumptions

- **Template path**: `Data_Module.TemplateDir + 'OrderSimulation.xls'`
  (`Order.pas:183`). `TemplateDir` (`DataModule.pas:708`): if
  `[DIRECTORIES] UseApplicationDir = True` (default True, `DataModule.dfm:675`) →
  the EXE directory; else `[DIRECTORIES] TemplateDir` (default
  `c:\_Inventory_Control\Templates`, `DataModule.dfm:667`). **Single template per
  install**, no per-site selection. A multi-site rebuild must pick the template
  per site/line.
- **Two databases**: order data on `Inv_Connection`; the working calendar
  (`AD_GetSpecialDate`, overtime/holiday) on `ALC_Connection` → catalog
  `TireOrder` (`DataModule.dfm:532-545`). Both connection strings (with
  passwords) come from INI `[DATABASE]`.
- **Tuning knobs** all single-valued in INI `[INIT]`: `FillDays`=23,
  `ForecastUsageCompare`=7, `UseFirstProductionDay`. No per-line/per-site
  override.
- **Site labels**: `[SITE] PlantName` (default `NUMMI`), `AssemblerName` used only
  in error text (`Order.pas:1228,1255`).
- Hard-coded Windows-only Excel OLE automation (`createOleObject('Excel.Application')`).

---

## 8. Hazards

1. **Cross-DB proc not in this schema (verify, don't assume missing):**
   `AD_GetSpecialDate` is called on `ALC_Connection`/`TireOrder`, not on
   `Inv_Connection`. It is **absent from `Create Inventory.sql`** — expected,
   because it lives in the ALC/Tire-order database (likely the `VehicleOrder.sql`
   / GALC schema family). **Body unverified.** Migration must locate it in the
   ALC schema before re-implementing the overtime/holiday calendar.
2. **Excel automation fragility / orphaned processes:** every error path calls
   `excel.Workbooks.Close; excel.Quit` but exceptions inside the
   open/close sequence can leave a headless `EXCEL.EXE` running; `excel.visible:=
   True` (`Order.pas:169`) means users can edit/lock cells mid-run, and the Order
   path explicitly warns about "active edit on worksheet" (`:890`). Reading values
   back as strings via `mysheet.Cells[].value` is locale/format-sensitive.
3. **Silent caps:** `fpartline[1..200]` (≤200 part rows) and `fDates[0..200]`
   bound the run with no guard (`Order.pas:46-51`); a part type with >200 rows is
   silently truncated.
4. **Dead/duplicate work:** double-open of `SELECT_SizeInfo` (`:920-941`);
   Pascal-computed FRS suffix digits are discarded by the proc (§6); the legacy
   `AQ/AR/AS` column code is commented out in favor of computed
   `DateWeekCol+FillDays+n` (e.g. `:197-199, 463, 1308`). Note: `Order.pas` (this
   live unit) is distinct from dead `Order1.pas`/`Orderold.pas` — confirm any fix
   targets `Order.pas` only.
5. **P9 shared RecordID / shared dataset reuse:** `INV_Forecast_DataSet` is
   reused for ~10 different procs within one row's processing; `Inv_StoredProc` is
   reused for both `SELECT_PartsStockInfoOrder` (the outer driving cursor) **and**
   for `SELECT_PartsStockRenban`/`UPDATE_PartsStockRenban` on the Order path —
   the outer cursor is repurposed mid-loop, so its position is not preserved
   across renban reads (the Order path re-derives everything from `fpartline[]`,
   so OK, but fragile). Any rebuild must not assume a single shared statement
   handle.
6. **No P8/P12 retry-recursion** in this unit (no wrong-target recursive retry
   pattern observed); error handlers `raise` or show-and-abort.
7. **Order math depends on `fDates` being indexed by calendar offset `x`** while
   forecast/phased arrays are indexed by *fill position `j`* — the two index
   spaces are reconciled only via the `fDates[i]<>0` scan
   (`Order.pas:1353-1369,1457-1468`). Off-by-one here misplaces in-transit/open
   buckets. Preserve this mapping exactly.
8. **Inventory coupling is in the trigger, not the proc:** order creation does
   NOT change `INV_PARTS_STOCK_MST.IN_QTY` (status empty → trigger condition
   false). Easy to get wrong in a rebuild that "adds qty on order."
