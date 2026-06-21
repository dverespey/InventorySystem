# Order-File Generator — DATA-Layer Behavioral Spec (M2 unit 2)

Source-truth analysis of the **order-file generator** `OrderFormCreateF.pas`
(`TOrderFormCreate_Form`, the live unit per `InventorySystem.dpr:33`; invoked from
`MainMenu.pas:276-279` with `FileKind:=fText`). This is the tool that, **after** orders
have been placed + renban-grouped, reads the committed open-order rows and emits the
supplier order files (`.ord` text + optional Excel) to each supplier's directory (and,
when FTP is on, to a logistics directory + an Archive copy).

> Scope note. This is a **different unit** from `Order.pas` (the "what to order"
> simulation worksheet) analyzed in `legacy-order-spec.md` / `order-redesign-plan.md`.
> `Order.pas` *computes* the order (lot/renban math) and *commits* `INV_OPEN_ORDER_INF`
> rows via `INSERT_OpenOrder`. `OrderFormCreateF.pas` does **no order math** — it
> serializes already-committed open-order rows into files and stamps them "ordered".
> The qty/renban numbers are decided upstream (`Order.pas` + `RenbanOrder.pas`); see §3.

Companion dead unit: `OrderFormCreate.pas` (class `TOrderFormCreate`) is an older
no-ship-date / no-logistics version, **not** in `InventorySystem.dpr` → ignore for the
rebuild (kept here only to note `.ord` line format drift, §3.4).

Confidence: HIGH — every proc body read from the schema dump AND verified against the
live `Inventory_Live` `OBJECT_DEFINITION`/`sp_helptext`; every data claim proved with a
bounded query (cited inline). `AD_GetSpecialDate` body read from the live `VehicleOrder`
DB.

Verification environment: `mssql-spike`, READ-ONLY against `Inventory_Live` (matched
legacy snapshot) and `VehicleOrder` (real ALC). Schema dump lines cite
`DB Schema/CreateInventory.sql` via `/tmp/inv_utf8.sql`.

---

## 0. The driving loop (what the generator does, in order)

`FormActivate` (`OrderFormCreateF.pas:55-707`) runs inside one
`Inv_Connection.BeginTrans … CommitTrans` (rollback on any exception, `:692`):

1. **Open the order feed** `SELECT_OrderNotOrdered;1` (no params) into `Inv_DataSet`
   (`:90`). If `recordcount = 0` → "No orders to process" and commit (`:680-683`).
2. **Walk the rows** ordered by `(VC_SUPPLIER_CODE, VC_RENBAN_NUMBER)`. On **supplier
   change** (`:101`): close the prior supplier's open file(s)/workbook(s), then for the
   new supplier resolve:
   - **per-part logistics dir** via `SELECT_PartsStockLogistics;1 @PartCode` (§1),
   - **supplier config** via `SELECT_SupplierInfo;1 @SupCode,@Logistics=1` (§4) →
     output-file kind (TEXT/EXCEL/BOTH), `Directory`, `LogisticsDirectory`, timestamp
     flag, site flag, create-order-sheet kind,
   - open the output target(s): a `.ord` text file (always, since `FileKind:=fText`
     forces BOTH-or-TEXT via the supplier's `Output File Type`), an Excel `OrderTemplate.xls`
     if EXCEL/BOTH, and a "special order sheet" (`OrderSheetTemplate{Tire,Wheel}.xls`) if
     the supplier has `VC_CREATE_ORDER_SHEET` set.
3. **Per row**: compute the **ship-date offset** (`SELECT_PartShipDays;1` +
   `GetShip()` calendar scan, §3.3), then write the order line into each open target
   (§3.1/§3.4), then **stamp the order as ordered** via `UPDATE_ORDEROrderDate;1` (§5).
4. After the last row, close the trailing file(s)/workbook(s) and commit.

The **only writes** are `UPDATE_ORDEROrderDate` (sets `VC_ORDER_DATE`, `VC_SHIP_DATE`,
`VC_LAST_UPDATE` on the open-order rows). No `INV_OPEN_ORDER_INF` insert/delete, no
`INV_PARTS_STOCK_MST` change. The files themselves are the product.

---

## 1. `SELECT_PartsStockLogistics` — the per-part skip-logistics resolution

**Body** (`/tmp/inv_utf8.sql:6093-6101`; live `Inventory_Live` matches):

```sql
CREATE PROCEDURE [dbo].[SELECT_PartsStockLogistics]
    @PartNo varchar(12)
AS
    SELECT l.VC_BREAKDOWN_ORDER_DIRECTORY 'LogisticsDirectory'
    FROM INV_PARTS_STOCK_MST p
        JOIN INV_LOGISTICS_MST l ON p.IN_LOGISTICS_ID = l.IN_LOGISTICS_ID
    WHERE p.VC_Part_Number = @PartNo
```

> Param-name trap: the proc declares **`@PartNo`**, but the Pascal adds a parameter
> **named `@PartCode`** (`OrderFormCreateF.pas:184`). ADO positional binding makes this
> work (one positional param), but a Named-Query rebuild must use the proc's real name
> `@PartNo` (or bind positionally). The form's own variable is `lastlogisticsdirectory`.

**What it returns:** a single column `LogisticsDirectory` = the **logistics master's**
breakdown-order directory (`INV_LOGISTICS_MST.VC_BREAKDOWN_ORDER_DIRECTORY`), for the
part's linked logistics provider. The **`JOIN` is INNER** — a part with
`IN_LOGISTICS_ID = NULL` (no logistics provider) yields **zero rows**.

**The NONE / '' / null resolution ladder** (`OrderFormCreateF.pas:188-224`):

| Step | Condition | `lastlogisticsdirectory` becomes |
|---|---|---|
| a | `SELECT_PartsStockLogistics` returns ≥1 row | the part's logistics dir (`:189`) |
| b | returns 0 rows | `''` (empty, `:191`) |
| c | then if it is still `''`: supplier's `LogisticsDirectory` (from `SELECT_SupplierInfo`) is **not null** | the **supplier-level** logistics dir (`:219-220`) |
| d | else (supplier logistics dir IS null) | the literal string **`'NONE'`** (`:222`) |

So the precedence is **part-logistics → supplier-logistics → `'NONE'`**. The string
`'NONE'` is the **skip-logistics sentinel**: every later FTP write guards with
`if lastlogisticsdirectory <> 'NONE'` (`:111,135,147,164,301,332,365,481,579,617,670`).
`'NONE'` ⇒ no logistics file is written for that supplier; the supplier `.ord` + the
Archive copy are still written.

> Edge: step (c) reads `'LogisticsDirectory'` off the **`SELECT_SupplierInfo` recordset**
> (`Inv_StoredProc` still positioned there). That field is `l.VC_BREAKDOWN_ORDER_DIRECTORY`
> from the supplier's logistics LEFT JOIN (§4). So part-level and supplier-level logistics
> both resolve to a `VC_BREAKDOWN_ORDER_DIRECTORY`, just keyed via the part's vs the
> supplier's `IN_LOGISTICS_ID`. The resolution is also **per-supplier, not per-part**:
> it is computed only on the supplier-change boundary using the **first row's** part
> number (`:185`). All later parts of that supplier reuse that one resolution.

### Proof on live data (`Inventory_Live`)

- Distinct `LogisticsDirectory` values reachable through the proc's INNER join:
  **none** — `SELECT … GROUP BY l.VC_BREAKDOWN_ORDER_DIRECTORY` returned **0 groups**.
- Counts: `total_parts = 47`, `parts_with_logistics (INNER-join survivors) = 0`,
  `parts_logistics_id_null = 47`. **Every part has `IN_LOGISTICS_ID = NULL`** → the
  proc returns 0 rows for *every* part → step (b) → `''` for all → resolution falls to
  the supplier level (c)/(d).
- Running `EXEC dbo.SELECT_PartsStockLogistics @PartNo='42602YY05000'` returned **0
  rows** (proof of the empty-result path).
- There IS exactly one logistics master row:
  `IN_LOGISTICS_ID=1, VC_LOGISTICS_NAME='TLDLOGISTICS SERVICES', dir='S:\TLDL'`. But no
  *part* points at it. **3 of 16 suppliers** point at it (`IN_LOGISTICS_ID=1`: suppliers
  `0572B`, `07100`, `38844`); the other 13 have `IN_LOGISTICS_ID = NULL` → for them the
  ladder ends at **`'NONE'`** (skip logistics).

**Net for the rebuild:** on the current production snapshot the part-level logistics
override is dormant (no part is wired); the effective logistics directory is entirely
**supplier-driven**: `S:\TLDL` for the three TLDL suppliers, `'NONE'` (skip) for the
rest. The rebuild must still implement the full part→supplier→NONE ladder because the
part-level column exists and is the documented 2005 feature ("Add Partnumber logistics",
`OrderFormCreateF.pas:12`).

---

## 2. The order data feed — `SELECT_OrderNotOrdered`

**Body** (`/tmp/inv_utf8.sql:6325-6336`; live `Inventory_Live` `sp_helptext` matches
verbatim):

```sql
CREATE PROCEDURE [dbo].[SELECT_OrderNotOrdered] AS
    SELECT *
        FROM INV_OPEN_ORDER_INF i
            JOIN INV_PARTS_STOCK_MST p ON i.VC_PART_NUMBER = p.VC_PART_NUMBER
            JOIN INV_SUPPLIER_MST  s ON p.IN_SUPPLIER_ID = s.IN_SUPPLIER_ID
            LEFT OUTER JOIN INV_RENBAN_GROUP_MST r ON P.IN_RENBAN_ID = r.IN_RENBAN_ID
    WHERE ((i.VC_ORDER_DATE is null) or (i.VC_ORDER_DATE = ''))
      AND i.VC_RENBAN_NUMBER <> ''
    Order by s.VC_SUPPLIER_CODE, i.VC_RENBAN_NUMBER
```

### 2.1 Lineage + cardinality

- **Source of order lines = `INV_OPEN_ORDER_INF`** (the committed open orders) — **NOT**
  `INV_BREAKDOWN_FC_INF`. The generator never reads the breakdown forecast table. The
  breakdown→order link is upstream (the Order/Renban worksheet consumed the forecast and
  wrote the open-order rows); see §6.
- Joins (all proved 1:1, **no fan-out**):
  - `i → p` on `VC_PART_NUMBER` (parts-stock has `UNIQUE (VC_PART_NUMBER)`,
    `CreateInventory.sql:255-258`).
  - `p → s` on `IN_SUPPLIER_ID` (FK to supplier PK).
  - `p → r` LEFT JOIN on `IN_RENBAN_ID` (renban-group PK; 1:1).
  - **Proof:** `all_open_orders = 4284`, `joined_rows = 4284` → **NO FAN-OUT (1:1 joins)**.
    `INV_RENBAN_GROUP_MST GROUP BY IN_RENBAN_ID HAVING COUNT(*)>1` returned **0 rows**
    (the renban-group join cannot fan).
  - One open-order row ⇒ exactly one order-file line.

### 2.2 Filter (the "not yet ordered" selection)

- `(VC_ORDER_DATE IS NULL OR VC_ORDER_DATE = '')` — rows not yet stamped by a prior file
  run (also note `VC_ORDER_DATE` is declared `NOT NULL`, `CreateInventory.sql:4518`, so
  in practice the `=''` branch is what matches; the IS NULL is defensive).
- **`AND VC_RENBAN_NUMBER <> ''`** — **only rows with an assigned renban are emitted.**
  This is the gate that keeps renban-grouped parts out of the file until the RenbanOrder
  grouping step has run (renban-grouped parts are born blank-renban then re-inserted with
  a renban; see `project-order-renban-domain` and §6). Lot-sized parts get their renban at
  `INSERT_OpenOrder` time, so they pass immediately.
- **Proof:** on `Inventory_Live`, `not_ordered (date empty AND renban<>'') = 0` and
  `not_ordered_but_no_renban = 0` (and total open = 4284, all already order-dated). So the
  live snapshot has nothing pending — expected for a matched post-run snapshot. The
  feed/format proofs below use the full open-order set as the representative shape.

### 2.3 Projection (the `SELECT *` 88-column shape)

`SELECT *` over the 4-table join yields **88 columns** with heavy duplication. Proved via
`sys.dm_exec_describe_first_result_set` on `Inventory_Live`. The columns the generator
actually reads, and the ordinal Delphi ADO `fieldbyname()` resolves to (**first match**):

| `fieldbyname(...)` | resolves to ordinal | meaning | the OTHER (shadowed) column |
|---|---|---|---|
| `VC_SUPPLIER_CODE` | **2** (`i.`) | open-order's supplier code | 57 (`s.`), 72-area = same value (proved, §6) |
| `VC_PART_NUMBER` | **3** (`i.`) | part number | 30 (`p.`) |
| `VC_FRS_NUMBER` | 4 (`i.`) | FRS / release number | — |
| `VC_RENBAN_NUMBER` | 5 (`i.`) | renban (trailer) number | — |
| `IN_QTY` | **6** (`i.`) | **order quantity** | **42 (`p.`) = on-hand stock qty** |
| `VC_KANBAN_NUMBER` | **20** (`i.`) | kanban (order-sheet only) | 33 (`p.`) |
| `VC_PARTS_NAME` | 31 (`p.`) | part name (order-sheet) | — |
| `IN_1LOTQTY` | 41 (`p.`) | 1-lot qty (wheel order-sheet) | — |

> **Critical duplicate-column trap (`IN_QTY`).** Ordinal 6 = `INV_OPEN_ORDER_INF.IN_QTY`
> (the order qty, e.g. 1200). Ordinal 42 = `INV_PARTS_STOCK_MST.IN_QTY` (the on-hand
> stock balance, e.g. 28133). `fieldbyname('IN_QTY')` returns the **first** = ordinal 6 =
> the order qty (correct). **Proof:** sampled rows show
> `open_order_qty=1200/400/360` vs `parts_stock_onhand_qty=28133/44418/22451` — two wildly
> different numbers under the one name. A rebuild that re-selects with `SELECT *` and grabs
> "the IN_QTY column" could grab the wrong one. The Named-Query rebuild MUST alias them
> distinctly (e.g. `i.IN_QTY AS order_qty`) and never rely on first-match ordering.
> Same hazard applies to `VC_SUPPLIER_CODE`, `VC_PART_NUMBER`, `VC_KANBAN_NUMBER`,
> `IN_LOGISTICS_ID`, `VC_LAST_UPDATE`, `VC_ADD` (all duplicated across the joined tables).

### 2.4 Per-supplier / per-renban cardinality

- Ordering is `(s.VC_SUPPLIER_CODE, i.VC_RENBAN_NUMBER)`. The generator detects a new
  **supplier** (file boundary, `:101`) and — for WHEEL order-sheets only — a new
  **renban** (a new order-sheet workbook, `:355`). So files are **one set per supplier**;
  the WHEEL special-order-sheet additionally produces **one workbook per renban**
  (plus pagination when a sheet exceeds ~8 line rows, `o>23`, `:470`).
- Multiple renban rows can share one part+FRS (renban-grouped parts): proved e.g.
  `4261102Q4100 / FRS 6052201` has **6 distinct renbans** (6 rows). Each is its own file
  line — but they all get order-stamped together (§5 hazard).

---

## 3. The qty / renban math + the file line format

### 3.1 The generator performs NO order math

The lot-sized / palletized / renban math is done **upstream** and already baked into the
`INV_OPEN_ORDER_INF` rows the generator reads:

- `Order.pas` `ProcessOrder` commits one row per lot via `INSERT_OpenOrder`, branching on
  `BIT_LOT_SIZE_ORDERS` (**stored INVERTED: 0 = lot-sized TRUE, 1 = palletized**, per
  `project-order-renban-domain`). Lot-sized → one row, qty = typed Qty; palletized → one
  row per lot, qty = `IN_1LOTQTY`.
- `RenbanOrder.pas` then deletes the blank-renban placeholders for a renban group and
  re-inserts trailer-grouped rows with real renbans + FRS suffixes.
- By the time `SELECT_OrderNotOrdered` runs, `IN_QTY` per row is the **final per-trailer
  order quantity** and `VC_RENBAN_NUMBER` is assigned. The generator copies `IN_QTY`
  through verbatim; it does **not** apply the lot multiplier, the INVERTED flag, or any
  rounding. **The rebuild's file generator must NOT re-derive qty** — it reads the
  committed `INV_OPEN_ORDER_INF.IN_QTY`.

So "what the `.ord` qty field actually is" = `INV_OPEN_ORDER_INF.IN_QTY` (an `int`,
`CreateInventory.sql:4508`) = the per-trailer/per-renban order quantity decided at order
commit. Not a forecast day-qty, not a lot count.

### 3.2 The `.ord` text line (the live `FileKind:=fText`/BOTH path, `:556-585`)

Built by string concatenation (`tcl`), one line per row:

```
[SiteSupplierCode]  VC_SUPPLIER_CODE  VC_FRS_NUMBER  %8s(VC_RENBAN_NUMBER)
    VC_PART_NUMBER  %.5d(IN_QTY)  yyyymmdd(now+ship)
```

Field by field (concat order = byte order in the file):

| Segment | Source | Format / width | Trap |
|---|---|---|---|
| site supplier prefix | `SiteSupplierCode` **only if `sendsite`** (`:560-565`) | as-is + `VC_SUPPLIER_CODE` | **`SiteSupplierCode` column does NOT exist** — see §4.2 hazard |
| supplier code | `VC_SUPPLIER_CODE` (ordinal 2 = `i.`) | as-is (`varchar(5)`) | — |
| FRS number | `VC_FRS_NUMBER` | as-is (`varchar(7)`) | — |
| renban | `VC_RENBAN_NUMBER` | **`format('%8s',…)`** = right-justified, space-padded to 8 (`:571`) | left-padded with spaces if <8; `varchar(8)` so never truncates |
| part number | `VC_PART_NUMBER` (ordinal 3) | as-is (`varchar(12)`) | — |
| **quantity** | **`IN_QTY` (ordinal 6)** | **`format('%.5d',[IN_QTY])`** = zero-padded to 5 digits (`:573`) | **>99999 prints all digits (no truncation), <0 prints `-NNNN`; 1200 → `01200`** |
| ship date | `now + ship` | `yyyymmdd` (`:574`) | `ship` = working-day offset (§3.3); see §3.3 traps |

`Writeln` to the supplier file `tcf`; if FTP on, also `Writeln` to logistics file `tlf`
(unless `'NONE'`) and Archive file `taf` (`:576-582`). **The same `tcl` line is written
to all targets** — logistics/archive copies are byte-identical to the supplier file.

> `%.5d` uses `IN_QTY.AsInteger`. `IN_QTY` is `NOT NULL int` so no NULL trap. Negative or
> >5-digit values are theoretically possible (no DB check constraint) and would widen the
> fixed-width line — a downstream parser keyed on column positions would misalign. Live
> data: order qtys observed 360-1200 (well within 5 digits).

### 3.3 Ship-date offset — `SELECT_PartShipDays` + `GetShip`

**`SELECT_PartShipDays`** (`/tmp/inv_utf8.sql:4783-4805`): returns 7 ship-day integers
(`Ship`, `ShipM..ShipS`). If the part is in a **renban group** (`IN_RENBAN_ID` not null)
it returns the **renban group's** ship-days (`INV_RENBAN_GROUP_MST.IN_SHIP_DAYS*`); else
the **part's** own (`INV_PARTS_STOCK_MST.IN_SHIP_DAYS*`). The generator picks the column
for **today's weekday** (`DayOfTheWeek(now)`), falling back to `Ship` (the generic
column) when the weekday-specific value is 0 (`:418-463`).

**`GetShip(lead)`** (`:709-770`) converts that lead count into a **calendar-day offset**
by scanning forward from tomorrow, counting only **working days**:
- skips weekends (`DayOfTheWeek(now+1+x) >= 6` → Sat/Sun, `:734`),
- skips **holidays** = a `AD_GetSpecialDate` row whose `Date Status Abrv = 'H'` on that
  date (`:740`),
- increments the valid-day counter `y` until `y = lead`; returns `x` (calendar offset).

The ship date on the file line = `now + GetShip(...)`, formatted `yyyymmdd` (text) /
`mm/dd/yyyy` (Excel) / `mm/dd/yy` (order-sheet).

**`AD_GetSpecialDate`** lives in **`VehicleOrder`** (ALC), not `Inventory` — verified
PRESENT, body read (`OBJECT_DEFINITION`). Called with `@LineName=''` → the "all lines"
branch, filtering `SpecialDate BETWEEN @BeginDate AND @EndDate`, returning columns incl.
`'Date'` (datetime) and `'Date Status Abrv'` (varchar(1)). **Status domain** (proved from
`ProductionStatus`): `N`=NORMAL, `O`=OVERTIME, `H`=HOLIDAY, `W`=WEEKEND, `X`=NON-PRODUCTION.

> **`GetShip` faithfulness traps:**
> - It skips ONLY weekends + `'H'`. It does **NOT** skip `'X'` (non-production) or `'W'`
>   days, and it treats `'O'` (overtime) as a normal working day. So an X/non-prod day
>   inside the window still counts as a shippable day. (Contrast `Order.pas`'s richer
>   O/X/holiday handling.) A rebuild must reproduce this exact, narrower rule.
> - The `@EndDate` window is `now + result + 59` re-evaluated *inside* the loop
>   (`:724`) — `result` starts 0, so effectively `now..now+59`; if `lead` needs >59
>   calendar days the holiday list runs out and holidays beyond day 59 are missed.
>   Live: holidays exist (e.g. 2026-01-01, 2026-07-13 "summer shutdown") — proved.
> - On exception `GetShip` swallows the error and **returns whatever `result` reached so
>   far** (often 0) — a silent wrong ship date, logged to ActLog only (`:763-768`).
> - `GetShip` opens `AD_GetSpecialDate` **once per order row** (inside the row loop) — a
>   cross-DB round-trip per line. The rebuild should cache the holiday set per run.

### 3.4 The Excel + special-order-sheet projections (secondary outputs)

- **`OrderTemplate.xls`** (EXCEL/BOTH, `:543-555`): rows from 10 down — col1=part,
  col2=FRS, col3=renban, col4=`IN_QTY`, col5=`mm/dd/yyyy` ship date. Header rows get
  supplier name/address (`:279-280`).
- **Special order sheet** `OrderSheetTemplate{Tire,Wheel}.xls` (`VC_CREATE_ORDER_SHEET`
  set, `:522-540`): a formatted PO sheet; new workbook per **renban** for WHEEL (`:355`),
  paginated when `o>23` (`:470`). Wheel sheet writes `IN_1LOTQTY` (col5) + `IN_QTY`
  (col7); tire sheet writes renban (col5) + `IN_QTY` (col7). FRS→date reconstruction:
  `year = first-3-of-yyyy + FRS[1]`, date = `FRS[2..3]/FRS[4..5]/year` (`:257-258`) —
  i.e. the FRS encodes a 1-digit year + MMDD.
- **Dead `.ord` format** in `OrderFormCreate.pas:167-171` (the legacy class, NOT shipped):
  no `%8s` renban pad, no site prefix, no ship date. Confirms the live format (§3.2) is
  the authoritative one.

---

## 4. The supplier / logistics master config — `SELECT_SupplierInfo`

**Body** (`/tmp/inv_utf8.sql:5993-6084`). Called by the generator as
`@SupCode=<supplier>, @Logistics=1` → the `@SupCode<>''` branch (single supplier).

### 4.1 Where each config field comes from

| Generator reads (`fieldbyname`) | proc alias | underlying column | role |
|---|---|---|---|
| `Output File Type` | `'Output File Type'` | `CASE VC_OUTPUT_FILE WHEN 'T'→TEXT 'E'→EXCEL 'B'→BOTH` | picks TEXT/EXCEL/BOTH file kind (`:208-213`) |
| `Directory` | `'Directory'` | **`s.VC_BREAKDOWN_ORDER_DIRECTORY`** | the supplier's own output dir (`:216`) |
| `LogisticsDirectory` | `'LogisticsDirectory'` | **`l.VC_BREAKDOWN_ORDER_DIRECTORY`** (via supplier's `IN_LOGISTICS_ID` LEFT JOIN) | supplier-level logistics dir (§1 step c) |
| `Supplier Name` | `'Supplier Name'` | `s.VC_SUPPLIER_NAME` | filename + sheet header (commas/periods stripped, `:227`) |
| `Order File Timestamp` | bit | `BIT_ORDER_FILE_TIMESTAMP` | append `yyyymmddhhmmss00` to filename (`:228`) |
| `Site Number in Order` | bit | `BIT_SITE_NUMBER_IN_ORDER` | gates the `sendsite` site-prefix (`:229`) |
| `Create Order Sheet` | varchar(5) | `VC_CREATE_ORDER_SHEET` | `''`→no sheet, `'TIRE'`/`'WHEEL'`→which template |
| Address/City/State/Zip | | `s.VC_ADDRESS` etc. | sheet header |

So **two independent directories per supplier**: `Directory` (`s.VC_BREAKDOWN_ORDER_DIRECTORY`)
and `LogisticsDirectory` (`l.VC_BREAKDOWN_ORDER_DIRECTORY` of the linked
`INV_LOGISTICS_MST`). Both are `varchar(512)`. Filenames are
`<SupplierName>-<SupCode>[-<timestamp>]` (+`.ord` for text); special-order-sheet files
are `OS<SupplierName>-<SupCode>-<Renban>[<page>]`.

### Proof on live data (`Inventory_Live`, all 16 suppliers)

```
code   supplier_dir         log_id  out  ts  site  create_order_sheet
0501B  [S:\GY]              <NULL>   B    0   1     [TIRE]
0572B  [S:\CMX]             1        B    0   0     [WHEEL]
07100  [S:\DICASTAL]        1        B    0   1     [WHEEL]
07451  [S:\DUNLOP]          <NULL>   B    0   1     [TIRE]
0946A  [S:\RONAL]           <NULL>   B    0   1     [WHEEL]
10011  [S:\EPC]             <NULL>   B    0   1     [MISC]
11111  [S:\PACIFICMFG]      <NULL>   B    0   1     [VALVE]
12720  [S:\HK]              <NULL>   B    0   1     [TIRE]
17800  [S:\MICHELIN]        <NULL>   B    0   1     [TIRE]
1793B  [S:\BSFS]            <NULL>   B    0   1     [TIRE]
30090  [S:\YOKOHAMA]        <NULL>   B    0   1     [TIRE]
38844  [S:\SUPERIOR]        1        B    1   0     [WHEEL]
43220  [S:\MAXXIS]          <NULL>   B    0   1     [TIRE]
7201A  [S:\CONT]            <NULL>   B    0   1     [TIRE]
72100  [S:\INDUSTRYPRODUCTS]<NULL>   B    0   1     [FILM]
93031  [S:\TAI]             <NULL>   B    0   1     [WHEEL]
```

Observations: **all suppliers are `Output File Type = B` (BOTH)** → every supplier
produces BOTH a `.ord` text file and an Excel workbook. Only **3 suppliers link a
logistics provider** (`log_id=1` = TLDL `S:\TLDL`); the rest resolve to `'NONE'`. Only
**1 supplier (`38844 SUPERIOR`) sets `ts=1`** (timestamped filenames). Every supplier has
a `VC_CREATE_ORDER_SHEET` value → every supplier also gets a special order sheet.
**`BIT_SITE_NUMBER_IN_ORDER` is 0 for the TLDL suppliers and 1 for most others** —
meaning the (broken) `sendsite` path WOULD be taken for 11 suppliers (see §4.2).

> `CASE l.VC_LOGISTICS_NAME WHEN null THEN '' ELSE … END` (`:6012-6016`, `:6055-6059`) is
> a **no-op bug**: `CASE x WHEN null` uses `x = NULL` which is never true, so the `''`
> arm never fires and a null logistics name passes through as NULL. Harmless here (the
> generator reads `LogisticsDirectory`/`Directory`, not `Logistics`), but flag it.

### 4.2 Multi-site relevance + the `SiteSupplierCode` drift hazard

- **`BIT_SITE_NUMBER_IN_ORDER`** (the per-supplier "prefix a site supplier code") is the
  one site-aware knob. When true (`sendsite`), the `.ord` line prepends
  `fieldbyname('SiteSupplierCode')` + the supplier code (`:560-565`).
- **HAZARD — `SiteSupplierCode` does not exist.** Proved three ways:
  (1) `grep -i SiteSupplierCode /tmp/inv_utf8.sql` = **0 hits**;
  (2) the live `SELECT_OrderNotOrdered` result set has **88 columns, none named
      `SiteSupplierCode`** (full ordinal list captured);
  (3) `INV_SUPPLIER_MST` on `Inventory_Live` has no column matching `%Site%` other than
      `BIT_SITE_NUMBER_IN_ORDER`.
  So if any supplier has `BIT_SITE_NUMBER_IN_ORDER = 1`, `fieldbyname('SiteSupplierCode')`
  raises (unknown field) → the whole run rolls back. On `Inventory_Live` **11 of 16
  suppliers have `site=1`** → this code path is **broken on this vintage**. Either the
  production DB this binary runs against has a `SiteSupplierCode`-bearing variant of
  `SELECT_OrderNotOrdered`/`INV_SUPPLIER_MST` (vintage drift not captured by the dump), or
  the feature is dead and the order file simply never exercises a `site=1` supplier the
  way the code assumes. **delphi-architect must adjudicate** whether `sendsite` is live in
  production. For the rebuild: the site-supplier-code prefix is a real EDI requirement
  (cf. the EDI layer's site scoping), so model it explicitly as a per-site supplier alias,
  not a phantom `SELECT *` column.
- Per-site **directories**: legacy is single-site — one `Directory`/`LogisticsDirectory`
  per supplier, INI-driven `TemplateDir`/`fiLocalFTP`. A multi-site rebuild must scope the
  output directory + logistics directory + the site-supplier-code per site (same need the
  EDI/Order layers flagged: zero `site_id` columns on these tables in the legacy schema).

---

## 5. The write-back — `UPDATE_ORDEROrderDate` (and its cross-renban hazard)

**Body** (`/tmp/inv_utf8.sql:5972-5984`):

```sql
CREATE PROCEDURE [dbo].[UPDATE_ORDEROrderDate]
    @PartNumber varchar(12), @FRSNumber varchar(7),
    @OrderDate varchar(8), @ShipDate varchar(8)=''
AS
    UPDATE INV_OPEN_ORDER_INF
    SET VC_ORDER_DATE=@OrderDate, VC_SHIP_DATE=@ShipDate,
        VC_LAST_UPDATE=CONVERT(varchar,getdate(),112)+ <hhmmss pieces>
    WHERE VC_PART_NUMBER=@PartNumber AND VC_FRS_NUMBER=@FRSNumber
      AND VC_ORDER_DATE <> @OrderDate
```

Called per row with `@OrderDate=yyyymmdd(now)`, `@ShipDate=yyyymmdd(now+ship)` (`:589-603`).

> **HAZARD — the WHERE has NO renban filter.** It matches on `(VC_PART_NUMBER, VC_FRS_NUMBER)`
> only, so processing **one** row stamps **every** open-order row sharing that part+FRS —
> across all renbans. **Proof:** `4261102Q4100 / FRS 6052201` has **6 distinct-renban
> rows**; one `UPDATE_ORDEROrderDate` call marks all 6 ordered. This is safe-as-written
> because (a) the file generator's recordset is a snapshot opened before any update, so all
> 6 rows are still emitted as 6 file lines, and (b) the purpose of the stamp is only to
> exclude them from the *next* `SELECT_OrderNotOrdered` run. A rebuild that processes
> rows individually and re-queries between rows would **drop** the not-yet-emitted
> siblings (they'd already be order-dated). **The rebuild must emit from a single snapshot,
> then stamp — not interleave per-row re-reads.**
>
> **NULL trap in the guard.** `AND VC_ORDER_DATE <> @OrderDate`: if `VC_ORDER_DATE` were
> NULL this predicate is UNKNOWN → the row would NOT be updated. `VC_ORDER_DATE` is
> `NOT NULL` (default ''), and `'' <> '20260621'` is true, so empty rows DO update.
> But a row already stamped with *today's* date is skipped (idempotent re-run within the
> same day leaves it; cross-day it re-stamps). Reproduce the `<>` guard exactly.

`@FRSNumber` is `varchar(7)` but `INV_OPEN_ORDER_INF.VC_FRS_NUMBER` is `varchar(7)` — OK,
no truncation. `@OrderDate`/`@ShipDate` are `varchar(8)` matching the `yyyymmdd` columns.

---

## 6. Forecast → order consistency (the chain integrity)

The question: does the order generator emit orders keyed to the **same supplier** the M2
forecast importer wrote the breakdown under?

- **The forecast importer** writes `INV_BREAKDOWN_FC_INF` via
  `INSERTUPDATE_BreakdownForecastInfo` (`/tmp/inv_utf8.sql:1215-1242`), keyed
  `(VC_SUPPLIER_CODE, VC_PART_NUMBER, IN_WEEK_NUMBER)` with day-qtys `IN_QTY1..7` and the
  supplier code passed in by the importer (the component's supplier).
- **The order generator** reads `INV_OPEN_ORDER_INF`, **not** the breakdown table. The
  breakdown table feeds the **Order worksheet** (`Order.pas` reads it via
  `SELECT_ForecastPartNumberWeek` to phase the daily forecast), which is where forecast→
  order qty actually happens. The committed open-order rows carry their own
  `VC_SUPPLIER_CODE` (set at `INSERT_OpenOrder` from the part's supplier).
- **Consistency proofs (`Inventory_Live`):**
  - Breakdown supplier vs parts-stock-master supplier: `bd_rows=959`, **matches=959,
    MISMATCH=0**. Every breakdown row's `VC_SUPPLIER_CODE` equals the part's
    `INV_PARTS_STOCK_MST → INV_SUPPLIER_MST` supplier.
  - Open-order supplier vs parts-stock-master supplier: `open_orders=4284`,
    **supplier_matches=4284, MISMATCH=0**. Every open-order row's own
    `VC_SUPPLIER_CODE` equals the part's master supplier — i.e. the column the generator
    *writes* (ordinal 2, `i.`) equals the column it *groups/orders by* (`s.VC_SUPPLIER_CODE`).
  - Transitively: breakdown supplier == parts-stock supplier == open-order supplier for
    every part. **The forecast→order chain is supplier-consistent.** A component the
    importer keyed under supplier S will (a) appear on supplier S's worksheet, (b) commit
    open orders under supplier S, (c) be grouped into supplier S's order file.
- **Where a mismatch COULD arise** (so the rebuild must preserve the invariant): the
  generator groups by `s.VC_SUPPLIER_CODE` (from the part→supplier join) but **writes**
  `i.VC_SUPPLIER_CODE` (the open-order's stored copy) into the file. These are two
  separate columns under the same name. They are equal in all live data, but if a part's
  `IN_SUPPLIER_ID` is repointed after an order is committed, the file grouping (`s.`)
  would diverge from the file's printed supplier code (`i.`). The rebuild should either
  key both off one source or assert the invariant.

**Net:** the forecast importer (`INV_BREAKDOWN_FC_INF`) and the order generator
(`INV_OPEN_ORDER_INF`) are joined through the **part's master supplier**, and on live data
they agree perfectly. The breakdown day-qtys are NOT read by the file generator — they are
consumed earlier by the Order worksheet to size the orders; the file generator only
re-emits the resulting committed orders.

---

## 7. Hazards (what will silently break a rebuild)

1. **`SELECT *` duplicate columns (esp. `IN_QTY`).** 88-column result, `IN_QTY` at
   ordinal 6 (order qty) AND 42 (on-hand stock); `fieldbyname` takes the first. Alias
   every column explicitly in the rebuild; never grab "the IN_QTY column". (§2.3, proved
   1200 vs 28133.)
2. **`SiteSupplierCode` phantom column.** The `sendsite` branch reads a column that does
   not exist in the schema or the live result set, yet 11/16 live suppliers have
   `BIT_SITE_NUMBER_IN_ORDER=1`. Either vintage drift or dead code — adjudicate with
   delphi-architect before trusting `sendsite`. (§4.2.)
3. **`UPDATE_ORDEROrderDate` stamps ALL same part+FRS rows (no renban filter).** Faithful
   only if you emit from one snapshot THEN stamp; per-row re-query drops siblings. NULL
   `<>` guard nuance. (§5, proved 6 renbans / 1 part+FRS.)
4. **Lot-sized flag is INVERTED (0 = lot-sized TRUE).** The generator itself doesn't read
   it, but the *upstream* qty it relies on does — any rebuild of the commit path that the
   generator depends on must invert. (`project-order-renban-domain`.)
5. **Renban gate `VC_RENBAN_NUMBER <> ''`.** Renban-grouped parts are excluded until the
   RenbanOrder step assigns renbans. A rebuild that emits blank-renban rows would ship
   un-grouped placeholder orders. (§2.2.)
6. **`'NONE'` logistics sentinel + 3-level resolution ladder.** part-logistics →
   supplier-logistics → `'NONE'`(skip). On live data the part level is dormant (all
   `IN_LOGISTICS_ID NULL`); effective logistics is supplier-driven, `'NONE'` for 13/16.
   `'NONE'` (string) ≠ NULL ≠ '' — three distinct states drive different file writes.
   (§1.)
7. **Cross-DB `AD_GetSpecialDate` (VehicleOrder).** `GetShip` skips only weekends + `'H'`
   holidays (NOT `X`/`W`/`O`); 59-day window; swallows errors returning a possibly-0
   offset → silent wrong ship date. Cache the holiday set per run. (§3.3.)
8. **Fixed-width `.ord` line (`%8s` renban, `%.5d` qty).** Right-justified 8-char renban,
   zero-padded 5-digit qty; qty >99999 or <0 widens the line and breaks position-based
   parsers. (§3.2.)
9. **Single-site directories.** `Directory`, `LogisticsDirectory`, `TemplateDir`,
   `fiLocalFTP` are all single-valued; no `site_id` on any table the generator touches. A
   multi-site rebuild must scope output dir + logistics dir + site-supplier-code per site
   (same gap as the EDI layer). (§4.2.)
10. **Excel OLE for every supplier (BOTH = all 16 live suppliers).** Orphaned `EXCEL.EXE`
    risk; the rebuild should generate the workbook server-side (no OLE). The `.ord` text
    is the EDI-critical artifact; the Excel/order-sheet are human-readable copies.
11. **Whole run in one transaction; files written mid-transaction.** If the commit fails
    after files are written to disk, the DB rolls back (`VC_ORDER_DATE` not stamped) but
    the **files remain on disk** — a re-run re-emits them (duplicate files, different
    timestamp). Files and DB state can desync. (§0, `:692`.)

---

## 8. What a faithful rebuild's order-file generator MUST read / compute

**Read (in this exact order, one snapshot):**
1. `SELECT_OrderNotOrdered` equivalent: `INV_OPEN_ORDER_INF` JOIN parts-stock JOIN
   supplier LEFT JOIN renban-group, WHERE `VC_ORDER_DATE` empty/null AND
   `VC_RENBAN_NUMBER <> ''`, ORDER BY supplier, renban. Project **aliased** columns
   (order qty = `i.IN_QTY`, supplier = `i.VC_SUPPLIER_CODE`, etc.) — never `SELECT *`.
2. Per supplier (first row's part): `SELECT_PartsStockLogistics(@PartNo)` → part-logistics
   dir or empty; then supplier config (`SELECT_SupplierInfo @SupCode,@Logistics=1`) →
   output kind, `Directory`, `LogisticsDirectory`, timestamp/site flags, create-order-sheet.
   Resolve logistics dir via the **part→supplier→'NONE'** ladder.
3. Per row: `SELECT_PartShipDays(@PartNumber)` (renban-group ship-days override) → weekday
   column → `GetShip` working-day scan against `AD_GetSpecialDate` (skip weekends + `'H'`).

**Compute / emit (no order math):**
- One file line per open-order row (1:1, no fan-out): supplier code, FRS, `%8s` renban,
  part number, `%.5d` order qty (`i.IN_QTY` verbatim — no lot/INVERTED/rounding),
  `yyyymmdd` ship date = `now + GetShip`.
- Files per supplier: `.ord` text (+ logistics copy unless `'NONE'` + Archive copy when
  FTP on); Excel `OrderTemplate.xls` for EXCEL/BOTH; special order sheet per renban for
  suppliers with `VC_CREATE_ORDER_SHEET`.

**Write back:** `UPDATE_ORDEROrderDate(@PartNumber,@FRSNumber,@OrderDate,@ShipDate)` —
stamps order/ship date + last-update on ALL same part+FRS rows; do it AFTER emitting the
snapshot, with the `VC_ORDER_DATE <> @OrderDate` guard.

**Preserve the invariants:** the breakdown→order supplier consistency (part master
supplier == open-order supplier == breakdown supplier, proved 0 mismatches), the renban
gate, the `'NONE'` sentinel semantics, and single-snapshot-then-stamp ordering.

---

## Appendix — cited proofs (all bounded, READ-ONLY, `Inventory_Live` / `VehicleOrder`)

- `SELECT_OrderNotOrdered` body verbatim vs dump: live `sp_helptext` == `/tmp/inv_utf8.sql:6325-6336`.
- 88-column result set, no `SiteSupplierCode`: `sys.dm_exec_describe_first_result_set('EXEC dbo.SELECT_OrderNotOrdered')`.
- No fan-out: `all_open_orders=4284` == `joined_rows=4284`; renban-group GROUP BY IN_RENBAN_ID HAVING COUNT>1 = 0 rows.
- `IN_QTY` dup (order vs on-hand): 1200/400/360 vs 28133/44418/22451.
- Logistics: 47 parts, 0 with non-null `IN_LOGISTICS_ID`; `SELECT_PartsStockLogistics @PartNo='42602YY05000'` = 0 rows; 1 logistics master `TLDL S:\TLDL`; 3 suppliers link it.
- Supplier config table (16 suppliers) with dir/log_id/out/ts/site/create-order-sheet.
- Supplier consistency: breakdown 959/959 match; open-order 4284/4284 match.
- Cross-renban stamp: `4261102Q4100/6052201` = 6 distinct-renban rows.
- `AD_GetSpecialDate` PRESENT in VehicleOrder; body read; status domain N/O/H/W/X; holidays e.g. 2026-01-01, 2026-07-13.
