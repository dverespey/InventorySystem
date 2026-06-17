# SIM_OrderSimulation — Spike Build Contract

Pins the exact legacy semantics the developer must reproduce so they do **not** guess.
Source truth: `legacy-order-spec.md` (cited `spec §N`), `option-a.md` (`opt §N`),
proc bodies in `DB Schema/Create Inventory.sql` (read via UTF-16LE→UTF-8, all line cites
are into that converted file). Sample keys verified against the live spike DB
(`docker exec mssql-spike … -d Inventory`, login `sa`; the gateway connection is
`Inventory_Spike`). **No `@site_id`** — grep confirms zero site columns in the schema
(opt §7); the proc takes NO site param.

Confidence: HIGH on all START proc bodies (read below) and on the sample-data profile
(queried live). The PAB/red threshold + palette are KEY-FACTS-confirmed. `AD_GetSpecialDate`
is cross-DB / body-unverified → stubbed per §3.

---

## 1. `SIM_OrderSimulation` contract

### 1.1 Input params

| Param | Type | Source / legacy origin | Notes |
|---|---|---|---|
| `@LineName` | `varchar(10)` | `Line_ComboBox` (spec §1). `''`/space ⇒ all-lines branch | drives `SELECT_PartsStockInfoOrder` (proc body 7388–7460) |
| `@PartType` | `varchar(50)` | `PartType_ComboBox` TIRE/WHEEL/VALVE/FILM/MISC | matched to `INV_PART_TYPE_MST.VC_PART_TYPE` exactly |
| `@SortType` | `varchar(50)` | `SortBy_ComboBox`, only meaningful when `@LineName=''` | one of `'LINE, RENBAN'` / `'LINE, PART NUMBER'` / `'PART NUMBER'` (proc body 7390/7414/7438). When `@LineName<>''` ⇒ always `ORDER BY s.vc_size_code` |
| `@Today` | `date`/`varchar(8)` | read-only Today field = order date (spec §1; opt §1.1) | the calendar walk anchor; gateway `now()` at simulate time; echoed back in metadata so screen+calc agree |
| `@FillDays` | `int` | INI `[INIT] FillDays` default **23**, max 50 (spec §1, opt §7) | number of **production** day-columns emitted |
| `@ForecastUsageCompare` | `int` | INI `[INIT] ForecastUsageCompare` default **7** (spec §3.2) | length of usage-vs-forecast compare window (drives `SELECT_UsageDay` loop) |
| `@UseFirstProductionDay` | `bit` | INI `[INIT] UseFirstProductionDay` (spec §3.2) | when 1, week-offset forecast lookups via `SELECT_FirstProductionDay` |

> **NO `@site_id`.** Adding scoping is schema surgery (opt §7 OPEN GAP), out of spike scope.

### 1.2 OUTPUT — result set A: day-header (one row per day-column)

Reconciles the two index spaces (§5). The screen builds column headers from this.

| Column | Type | Meaning |
|---|---|---|
| `fill_pos` | int | 0-based **fill position `j`** (0..@FillDays-1); the grid column index |
| `cal_offset` | int | **calendar offset `x`** from `@Today` that produced this column (the `fDates` index, spec §3.1) |
| `serial_date` | date | rendered production date (`fDates[x]`, spec §3.1) |
| `weekday` | tinyint | 1=Mon..7=Sun (`DayOfTheWeek`, spec §3.2) |
| `day_kind` | enum | `NORMAL` \| `OVERTIME` \| `NONPRODUCTION` (spec §3.1: 'O'→OVERTIME, 'X'→NONPRODUCTION, else NORMAL) |

### 1.3 OUTPUT — result set B: grid rows (one row per legacy row)

`row_kind ∈ {SIZE_HEADER, PART, SIZE_FOOTER}` (spec §3.2 block structure). PART rows carry
the identity + order-point + inventory + order columns; SIZE_HEADER carries the H/I/J size
inputs and `Beg Balance`; SIZE_FOOTER carries `Usage`/`End Balance` summary. All columns NULL
where not applicable to the row_kind.

| Column | Legacy col | Source proc / field | Notes |
|---|---|---|---|
| `row_kind` | — | derived | grouping by `VC_SIZE_CODE` (spec §3.2) |
| `size_code` | B | `INV_SIZE_MST.VC_SIZE_CODE` (`SELECT_PartsStockInfoOrder`) | group key |
| `size_name` | — | `SELECT_SizeInfo`→`Size Name` | |
| `supplier_code` | C-key | `INV_SUPPLIER_MST.VC_SUPPLIER_CODE` | |
| `supplier_name` | C | `SELECT_SupplierInfo`→`Supplier Name` | written to C (spec §3.2) |
| `part_number` | D | `VC_PART_NUMBER` | identity |
| `kanban` | — | `VC_KANBAN_NUMBER` | needed by ORDER path renban build (not displayed) |
| `renban_group` | — | `INV_RENBAN_GROUP_MST.VC_RENBAN_GROUP_CODE` (LEFT JOIN; NULL/'' ⇒ not in group) | drives renban path (spec §6) |
| `daily_usage` | **H (8)** | `SELECT_SizeInfo`→`Daily Usage` = `IN_USAGE` | on SIZE_HEADER; interior 34 (spec §4) |
| `safety_days` | **I (9)** | `SELECT_SizeInfo`→`Safety Days` = `IN_DAYS` | on SIZE_HEADER; interior 34 |
| `safety_stock` | **J (10)** | computed `J = H*I` = `IN_USAGE*IN_DAYS` (spec §3.2) | **the BELOW_SAFETY threshold `J7`** (KEY FACT) |
| `total_inv` | Total Inv (`S+FillDays+1`) | `IN_QTY` (spec §3.2 "Total inventory") | summary col |
| `wh_qty` | **K (11)** | `K = TotalInv − (L+M)` available formula (spec §3.2) — reproduce as value | |
| `assembler_qty` | (K input) | `SELECT_OrderAtASSEMBLER`→`Qty` (proc 6643) | feeds K per `PutASSEMBLERCount`→K |
| `plant_qty` | **L (12)** | `SELECT_OrderAtPLANT`→`Qty` (proc 6700) | `PutPLANTCount`→L |
| `in_transit` | **M (13)** | `SELECT_OrderInTransit`→`Qty` (proc 6816) | `PutIntransitCount`→M; **note seed = always 0**, see §4 |
| `open_order` | **N (14)** | `SELECT_OrderOpenOrder`→`Qty` (proc 6955) | `PutOpenOrderCount`→N |
| `lot_qty` | **O (15)** | `IN_1LOTQTY` | order-entry input |
| `lead_time` | **P (16)** | weekday-selected leadtime (spec §3.2) — see §1.4 | interior 36 (input highlight) |
| `qty_default` | **Q (17)** | seeded `=Q{row}` (lot qty) when order-by cell empty (spec §3.4) | editable on screen; cream/40 |
| `lot_default` | **R (18)** | default lot (spec §3.4/§6) | editable; cream/40 |
| `share_pct` | E/F | `SELECT_OrderHistory`→`Qty` per part; share `= E/(ΣE)` over size group, 100% if singleton (spec §3.2) | computed value |
| `tire_wheel_ratio` | G | `SELECT_ForecastDetailTWPN`→`IN_TIRE_RATIO`/`IN_WHEEL_RATIO` /100 (spec §3.2) | |
| `added_leadtime` | `S+FillDays+3` | `DoLeadTime` (spec §3.4) — see §3 | overtime days inside lead window |
| `leadtime_zone_end_index` | — | fill_pos of last lead-time-zone column `= leadtime-1+addedleadtime` (spec §3.4) | for `LEADTIME_ZONE` paint |
| `orderby_col_index` | — | fill_pos of order-by column `= leadtime+addedleadtime` (spec §3.4) | for `ORDER_BY` paint |
| `frs_date` | — | `fPartLine[line,FRSDate]` = serial_date at order-by column (spec §3.4) | ORDER path FRS build |

### 1.4 Lead-time selection (spec §3.2, `Order.pas:426-459`)
`case DayOfTheWeek(@Today)`: Mon→`IN_LEADTIME_MONDAY` … Sat→`IN_LEADTIME_SATURDAY`;
**if that weekday column is 0/NULL ⇒ fall back to `IN_LEADTIME`**; Sun(7)/else ⇒ `IN_LEADTIME`.

### 1.5 OUTPUT — result set C: phased day cells (one row per grid-row × fill_pos)

This is the per-cell payload the PhasedGrid renders. Key = `(part_number, fill_pos)`.

| Column | Type | Meaning |
|---|---|---|
| `part_number` | varchar | join back to result set B |
| `fill_pos` | int | 0-based day column |
| `value` | int/decimal | the emitted cell number (forecast usage, in-transit qty, open-order qty, or projected balance for footer rows) |
| `balance` | int | PAB (Projected Available Balance) for this day on the End-Balance row (§1.6) |
| `signal_enum` | enum | **interior/zone** signal (≤1, with override precedence — §1.7) |
| `source_enum` | enum | **font/qty-source** signal, independent channel (§1.7) |

### 1.6 PAB recursion (KEY FACT — confirmed, spec §3.5)
```
day0.begin = total_inv − in_transit        -- (TotalInv − InTransit)
day_j.begin = day_(j-1).end
day_j.end   = day_j.begin + receipts_j − usage_j
balance (the column above) = day_j.end
```
`receipts_j` = in-transit + open-order qty bucketed into day j (§4/§5). `usage_j` =
phased forecast for day j (`fForecast[j]`, spec §3.2 `FillForecast`/`UpdateForecast`).

### 1.7 Per-cell signal enums

`signal_enum` (interior channel — paint precedence high→low, later overrides earlier per spec §3.4):
| enum | ColorIndex | RGB (KEY FACT) | Rule |
|---|---|---|---|
| `OVERTIME` | 3 | `#FF0000` | column whose `day_kind=OVERTIME` (spec §3.4 step 3 — overrides zone) |
| `NON_PRODUCTION` | 4 | `#00FF00` | column whose `day_kind=NONPRODUCTION` (step 3 — overrides zone) |
| `ORDER_BY` | 40 | `#FFCC99` | `fill_pos == orderby_col_index` (= leadtime+addedleadtime) |
| `LEAD_TIME_ZONE` | 36 | `#FFFF99` | `0 ≤ fill_pos ≤ leadtime_zone_end_index` |
| `NONE` | — | — | otherwise |

`source_enum` (font channel — independent; co-occurs with any interior):
| enum | ColorIndex | RGB | Rule |
|---|---|---|---|
| `IN_TRANSIT` | 23 | `#333399` | cell value came from an in-transit bucket (spec §3.3, font 23) |
| `OPEN_ORDER` | 10 | `#008000` | cell value came from an open-order bucket (spec §3.3, font 10) |
| `NONE` | — | — | forecast/balance cells |

`BELOW_SAFETY` (red font on End-Balance cells — KEY FACT, spec §3.5): emit as a boolean
flag `below_safety = (balance < safety_stock)` i.e. **PAB < J7**. Rendered as red font
(`#FF0000`) + non-color "⚠ below safety" channel (opt §5). Carry it as its own column on the
SIZE_FOOTER/balance rows (it is orthogonal to `signal_enum`/`source_enum`).

---

## 2. Ordered START reads the proc must orchestrate (spec §2)

Calendar first, then per-size header, then per-part fan-out. (`SELECT_SizeInfo` double-open
at `Order.pas:935-941` is dead — call ONCE.)

| Order | Proc (`;1`) | Body | Supplies | → Output col |
|---|---|---|---|---|
| 0 | `AD_GetSpecialDate` (ALC, **STUB §3**) | cross-DB | per-day O/X/holiday status | result-set A `day_kind`, calendar walk |
| 1 | `SELECT_PartsStockInfoOrder` | 7382 | driving cursor: all parts of `@PartType`, ordered by size | identity cols, K/L/M/N keys, P inputs, renban_group |
| 2 | `SELECT_SizeInfo` (once) | 7869 | `IN_USAGE`,`IN_DAYS` per size | H, I, J=H*I |
| 3 | `SELECT_SupplierInfo` | 7978 | `Supplier Name` | C |
| 4 | `SELECT_OrderHistory` | 6757 | per-part historical order Qty | E (→ share %/F) |
| 5 | `SELECT_ForecastDetailTWPN` | 6228 | `IN_TIRE_RATIO`/`IN_WHEEL_RATIO` | G |
| 6 | `SELECT_UsageDay` (×@ForecastUsageCompare) | 8088 | actual shipped qty per recent day | usage-vs-forecast compare row |
| 7 | `SELECT_FirstProductionDay` (if @UseFirstProductionDay) | 5982 | `First Week Number` | week-offset for forecast lookups |
| 8 | `SELECT_ForecastPartNumberWeek` (per week×day) | 6309 | `IN_QTY1..7` by day | `fForecast[j]` → phased usage |
| 9 | `SELECT_OrderAtASSEMBLER` | 6643 | shipping-status assembler-yard qty | K input |
| 10 | `SELECT_OrderAtPLANT` | 6700 | warehouse/plant-yard qty | L |
| 11 | `SELECT_OrderInTransit` | 6816 | in-transit total | M |
| 12 | `SELECT_OrderInTransitList` | 6850 | per-FRS-date in-transit rows | phased buckets, source_enum=IN_TRANSIT |
| 13 | `SELECT_OrderOpenOrder` | 6955 | open-order total | N |
| 14 | `SELECT_OrderOpenOrderList` | 6985 | per-FRS-date open-order rows | phased buckets, source_enum=OPEN_ORDER |

Status-filter semantics that matter (verified bodies): in-transit = `VC_STATUS_SUPPLIER_SHIPPING<>''
AND VC_ARRIVAL='' AND VC_STATUS_PLANT_YARD='' AND VC_STATUS_ASSEMBLER_YARD='' AND VC_WAREHOUSE=''
AND VC_STATUS_EMPTY_TRAILER='' AND VC_TERMINATED='' AND VC_FRS_DATE>=@FirstFRS` (6816/6850).
open-order = `(VC_STATUS_SUPPLIER_SHIPPING IS NULL OR ='')` (6955/6985). These two sets are
**mutually exclusive on `VC_STATUS_SUPPLIER_SHIPPING`** — a row is either in-transit (shipping
set) or open (shipping empty), never both. `@FirstFRS` = `yyyymmdd` of the first rendered date.

---

## 3. `AD_GetSpecialDate` STUB contract

Cross-DB (ALC `TireOrder`), body unverified (spec §8 h1). Stub with a fixture so the calendar
walk (spec §3.1) + added-leadtime break-loop (spec §3.4) run. Mark calendar-derived cells
**fixture-backed**.

### 3.1 Result shape (mirror legacy consumption, spec §3.1)
```
DATE                 (datetime / yyyymmdd)   one row per affected day in [@Today, @Today+@FillDays*3]
[Date Status Abrv]   varchar                 'O' | 'X' | 'H' | (absent row ⇒ normal)
```
Status domain consumed by the walk:
- `'O'` Overtime → render date, counts as a fill day, push 1-based fill index into `fOvertimes[]`, `INC(fOvertimeCount)` (spec §3.1).
- `'X'` Non-Production → render, counts as a fill day, push into `fNonProduction[]` (spec §3.1).
- `'H'`/holiday/any-other matched status → **skipped**: consumes calendar (x advances) but NOT a fill column (spec §3.1 "other → next").
- **No row for a day** → normal: render iff `DayOfTheWeek(@Today+x) < 6` (Mon–Fri); weekends skipped.
- Empty result set ⇒ "normal run": pure Mon–Fri loop (spec §3.1).

### 3.2 Walk consumption
- Walk `x` (calendar offset) forward; advance fill counter only on rendered production days; stop when `fFillDays == @FillDays`.
- `fDates[x]` = serial date of each rendered production day (indexed by `x`, NOT fill_pos — §5).
- `fOvertimes[]` holds 1-based **fill positions** of overtime days; consumed by the added-leadtime loop (spec §3.4):
  `for i in fOvertimes: if (fOvertimes[i]-1) <= (leadtime+i) then INC(addedleadtime) else break`.

### 3.3 Concrete fixture calendar (INCLUDES hazard-7: O/X/holiday between fill positions ⇒ offset≠position)
Anchor `@Today = 2026-06-15` (Mon). Window covers the spike sample dates (open orders 2026-06-15..19+).
```
DATE        STATUS   effect
2026-06-17  H        holiday → skipped (x=2 consumed, NO fill column)   <-- forces offset>position
2026-06-18  X        non-production → fill column, flag NONPRODUCTION
2026-06-20  (Sat)    no row → weekend, skipped
2026-06-21  (Sun)    no row → weekend, skipped
2026-06-23  O        overtime → fill column, flag OVERTIME, push fill idx into fOvertimes
2026-06-25  H        holiday → skipped
2026-07-03  O        overtime (inside a long-leadtime window) → second fOvertimes entry
```
Resulting first fill positions (j) ↔ calendar offset (x from 2026-06-15):
| j (fill_pos) | date | x (cal_offset) | day_kind |
|---|---|---|---|
| 0 | 2026-06-15 Mon | 0 | NORMAL |
| 1 | 2026-06-16 Tue | 1 | NORMAL |
| — | 2026-06-17 Wed | 2 | (H, skipped — no column) |
| 2 | 2026-06-18 Thu | 3 | NONPRODUCTION |
| 3 | 2026-06-19 Fri | 4 | NORMAL |
| — | 06-20/21 | 5,6 | (weekend) |
| 4 | 2026-06-22 Mon | 7 | NORMAL |
| 5 | 2026-06-23 Tue | 8 | OVERTIME (fOvertimes←6, the 1-based fill idx) |
| 6 | 2026-06-24 Wed | 9 | NORMAL |
| — | 2026-06-25 Thu | 10 | (H, skipped) |
| 7 | 2026-06-26 Fri | 11 | NORMAL |
…continue Mon–Fri until j reaches @FillDays-1.

Because of the 06-17 holiday, **fill_pos 2 maps to cal_offset 3** — offset≠position. Any
open-order/in-transit FRS-date that lands on 06-18 must bucket into `fill_pos=2`, found by the
`fDates[i]<>0` scan (§5), not by date-minus-today arithmetic. This is the hazard-7 test case.

---

## 4. Sample-part set (verified against live spike `Inventory` DB)

All parts: `VC_LINE_NAME='COROLLA'` (the only line present), `VC_PART_TYPE` as noted.
Suppliers: 07451=DUNLOP, 17800=MICHELIN, 43220=MAXXIS, 30090=YOKOHAMA, 12720=HANKOOK, 11111(RV).

| Case | Part number | Type / Size | Supplier | Kanban | IN_QTY | 1LotQty | LeadTime | Renban | Why |
|---|---|---|---|---|---|---|---|---|---|
| **(a) singleton size (share=100%)** | `4265202R6000` | TIRE / 15D | 1793B | 15BS | 12000 | 800 | 5 | none | size 15D has exactly 1 part ⇒ share = 100% (no `=E/ΣE`) |
| **(b) shared-size group (=E/ΣE split)** | `4265202S1000` + `4265202S2000` | TIRE / **18DL** | 07451 / 17800 | 18DL / 18M | 23608 / 23608 | 1375 / 1320 | 5 / 6 | size 18DL has 2 parts. Forecast tire ratios 70 / 30 (`INV_FORECAST_DETAIL_INF`) — exercises split + ratio→G |
| **(c) in-transit AND open-order** | **NOT SATISFIABLE in seed** | — | — | — | — | — | — | **GAP** — see below |
| **(d) breaches safety stock (PAB<J7)** | `900804500600` | VALVE / **RV** | 11111 | RV | 14600 | 1000 | 15 | none | RV is the only sampled size with `IN_DAYS>0`: usage 922, days 10 ⇒ **J7=9220**. Start IN_QTY 14600 draws down 922/day ⇒ PAB < 9220 within ~6 production days ⇒ BELOW_SAFETY fires on the End-Balance row. lot_size_orders=1 |
| **(e) hazard-7 fixture calendar + busy phasing** | `4261102Q8000` | WHEEL / M1 | 0572B | M1 | 44418 | 40 | 5 | **CMWA (IN_RENBAN_ID=11)** | singleton size M1 but **8 future open-order rows** dated 2026-06-15..19 — these bucket across the hazard-7 fixture days (incl. the 06-18 X column at fill_pos 2). Also exercises the renban-group path (renban deferred, `@RenbanNum=''`) on the ORDER side |

### Case (c) GAP — minimal seed proposal (REQUIRED to test the in-transit channel)
**Confirmed:** all 4238 `INV_OPEN_ORDER_INF` rows have `VC_STATUS_SUPPLIER_SHIPPING=''`
(distinct values = `['']` only) ⇒ `SELECT_OrderInTransit` returns 0 for every part; the M
column and `source_enum=IN_TRANSIT` path are **never exercised** by current data. Minimal seed:
take an existing open-order row for `4261102Q8000` with a future `VC_FRS_DATE` (e.g. 2026-06-18,
which maps to the X column / fill_pos 2) and set `VC_STATUS_SUPPLIER_SHIPPING='Y'` (any non-empty),
leaving `VC_ARRIVAL='', VC_STATUS_PLANT_YARD='', VC_STATUS_ASSEMBLER_YARD='', VC_WAREHOUSE='',
VC_STATUS_EMPTY_TRAILER='', VC_TERMINATED=''`. That single row then qualifies as in-transit
(font 23) while its sibling rows stay open-order (font 10) — giving case (c) on one part and
simultaneously placing an in-transit bucket on the hazard-7 X day. Apply only in the spike DB.

---

## 5. Hazard-7 index reconciliation rule (spec §8 h7, `Order.pas:1353-1369,1457-1468`)

Two index spaces MUST both be preserved; do not collapse them:

- **`x` = calendar offset** from `@Today`. `fDates[x]` holds the serial date of a rendered
  production day; `fDates[x]=0` for skipped calendar days (weekends, holidays). Indexed 0..@FillDays*3.
- **`j` = fill position** (0..@FillDays-1) = the grid day-column. This is the index of
  `fForecast[]` and of every emitted phased cell.

**Mapping rule (the reconciling scan):** to place an in-transit / open-order qty whose
`VC_FRS_DATE` parses (via `StrToDate(mm/dd/yyyy)` from `copy(frs,5,2)/copy(frs,7,2)/copy(frs,1,4)`,
spec §3.3) to a target serial date `D`:
```
j := -1
for x := 0 to lastCalOffset:
    if fDates[x] <> 0 then INC(j)         -- count only production days
    if fDates[x] = D   then bucket into fill_pos = j; break
```
The proc must emit BOTH `fill_pos` and `cal_offset` in result set A (§1.2) so the bucket
lookup is done by **matching the date through `fDates`**, never by `datediff(@Today, D)`. With
the §3.3 fixture, a 2026-06-18 FRS lands at `fill_pos=2` (not 3), because the 06-17 holiday left
`fDates[2]=0`. The added-leadtime loop (spec §3.4) likewise reads `fOvertimes[]` in **fill-position**
space. A set-based T-SQL rewrite must materialize the `fDates`→`fill_pos` map (e.g. a derived
table `(cal_offset, fill_pos, serial_date)` filtered to production rows) and JOIN FRS dates to it
on `serial_date`. Off-by-one here misplaces every in-transit/open-order bucket by a day.

---

## RETURN SUMMARY

**Result-set column lists:** see §1.2 (A: `fill_pos, cal_offset, serial_date, weekday, day_kind`),
§1.3 (B: row_kind, size_code, size_name, supplier_code/name, part_number, kanban, renban_group,
daily_usage(H), safety_days(I), safety_stock(J=H*I), total_inv, wh_qty(K), assembler_qty,
plant_qty(L), in_transit(M), open_order(N), lot_qty(O), lead_time(P), qty_default(Q), lot_default(R),
share_pct, tire_wheel_ratio, added_leadtime, leadtime_zone_end_index, orderby_col_index, frs_date),
§1.5 (C: part_number, fill_pos, value, balance, signal_enum, source_enum, +below_safety flag).

**Sample-part keys (live spike DB, line=COROLLA):**
- (a) singleton: `4265202R6000` TIRE/15D, DUNLOP-area, kanban 15BS.
- (b) shared group: `4265202S1000`(70%) + `4265202S2000`(30%) TIRE/18DL.
- (c) in-transit+open: **not satisfiable** (all open orders have empty shipping status) — seed
  one shipping row on `4261102Q8000` per §4.
- (d) safety breach: `900804500600` VALVE/RV, J7=9220 (usage 922 × days 10), IN_QTY 14600.
- (e) hazard-7 + busy phasing + renban: `4261102Q8000` WHEEL/M1, CMWA renban, 8 future open orders.

**Stub-calendar fixture:** §3.3 — anchor 2026-06-15; rows 06-17=H, 06-18=X, 06-23=O, 06-25=H,
07-03=O; yielding fill_pos↔cal_offset divergence at the 06-17 holiday (fill_pos 2 = cal_offset 3),
which is the hazard-7 test.
