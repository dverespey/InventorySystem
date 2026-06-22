# Daily Shipping Report — Source-Truth Spec (the FAILING path)

**Phase:** M3 source-truth (2026-06-21) **Analyst:** Claude **Confidence:** HIGH (Pascal handlers +
all 3 proc bodies read from the live dump; table grain verified). **Decision items flagged for David.**

> The M3 plan prioritizes Daily Shipping as the "failing path." This spec gives the exact Excel layout,
> the data procs (verbatim), **why it fails**, and what the server-side rebuild must render WITHOUT Excel.

## 1. The three triggers (all live, in `MainMenu.pas`, Reports menu)

| Variant | Menu item | Handler | Proc | `@param` wiring |
|---------|-----------|---------|------|-----------------|
| **Daily Shipping (T/W)** | `DailyShipping` (`:106`) | `DailyShippingClick` (`:3029`) | `dbo.REPORT_DailyShipping;1` (`:3047`) | `@PDate` = `copy(prodDate,1,4)+copy(,6,2)+copy(,9,2)` → `yyyymmdd` (`:3049-3050`) |
| **Daily Shipping Range (T/W)** | `DailyShippingRangeTireWheelPartNumbers` (`:118`) | `DailyShippingRangeTireWheelPartNumbersClick` (`:2919`) | `dbo.REPORT_DailyShippingRange;1` (`:2939`) | `@BeginPDate`, `@EndPDate` (`:2941-2944`) |
| **Daily Shipping ASN (Assy)** | (Daily ASN) | `DailyASNReportClick` (`:3135`) | `dbo.REPORT_DailyShippingAssy;1` (`:3154`) | `@PDate` + `@Line` (`:3156-3159`) |

Param entry: `TProductionDateSelectDlg` (`ProductionDates.pas`, LIVE). For T/W variants
`INVOICE:=FALSE; ASN:=FALSE` (`:3035-3036`); Range sets `Range:=TRUE` (`:2928`); Assy sets `ASN:=TRUE`
(`:3143`). On `Cancel`, no run. Date format from the dialog is `yyyy/mm/dd` (hence the `copy` re-pack
to `yyyymmdd` for the proc).

## 2. Excel layout — exact `mysheet.cells[r,c]` map

All three: `CreateOleObject('Excel.Application')` → open `TemplateDir+'ReportTemplate.xls'` →
`mysheet := excel.workSheets[1]` → write → `SaveAs(fiReportsOutputDir + '\<name>'+ yyyymmddhhmmss00 +
'.xls')` → optional `MessageDlg('Print this report?')` → `mysheet.PrintOut`. On exception:
`LogActLog('ERROR','Failed on Daily Shipping … '+e.Message)` + `ShowMessage` + Excel cleanup.

### 2a. Daily Shipping (T/W) — `DailyShippingClick` (`MainMenu.pas:3056-3092`), out=`DailyShippingTW`

| Cell | Content | Source |
|------|---------|--------|
| `[1,1]` | `'Daily Shipping (Tire/Wheel Part Numbers)'` | literal title |
| `[2,1]` | production date (`NumberFormat 'yyyy/mm/dd'`) | `ProductionDate` |
| `[2,2]` | `'Start Seq:'+Start+'/End Seq:'+End` | proc cols `Start`, `End` (**first row only**) |
| `[2,3]` | `'Vehicle Count:'+Vehicle Count` | proc col `Vehicle Count` (**first row only**) |
| `[3,1]` | `'Part Number'` (col width 17, fmt `############`) | header |
| `[3,2]` | `'Part Description'` (col width 30) | header |
| `[3,3]` | `'Qty'` (col width 5) | header |
| `[z,1]` z=4..n | row part number | `fieldbyname('Part Number')` |
| `[z,2]` | row description | `fieldbyname('Desc')` |
| `[z,3]` | row qty | `fieldbyname('PQTY')` |

Detail loop `while not eof … next` (`:3080-3087`). **Note:** `Start`/`End`/`Vehicle Count` are read
ONCE before the loop (from the current/first row); they are header values, NOT per-row.

### 2b. Daily Shipping Range (T/W) — `…RangeClick` (`:2950-2986`), out=`DailyShippingRangeTW`

Identical layout EXCEPT header row 2:

| Cell | Content |
|------|---------|
| `[1,1]` | `'Daily Shipping Range(Tire/Wheel Part Numbers)'` |
| `[2,1]` | begin production date (`yyyy/mm/dd`) |
| `[2,2]` | end production date (`yyyy/mm/dd`) |
| `[3,1..3]` / `[z,1..3]` | same `Part Number` / `Part Description` / `Qty` → `Part Number`/`Desc`/`PQTY` |

No Start/End/Vehicle-Count header cells (the Range proc has those columns commented out — see §3).

### 2c. Daily Shipping ASN (Assy) — `DailyASNReportClick` (`:3165-3199`), out=`DailyShippingAssy`

| Cell | Content | Source |
|------|---------|--------|
| `[1,1]` | `'ASN (Assy Part Numbers)'` | title |
| `[2,1]` | production date (`yyyy/mm/dd`) | `ProductionDate` |
| `[2,2]` | `'Start Seq:'+Start+'/End Seq:'+End` | proc `Start`/`End` |
| `[2,3]` | `'Vehicle Count:'+Vehicle Count` | proc `Vehicle Count` |
| `[3,1]` | `'Part Number'` (w17, fmt `############`) | header |
| `[3,2]` | `'Qty'` (w30) | header — **2-col only; `Desc` write commented out (`:3190`)** |
| `[z,1]` | part number | `fieldbyname('Part Number')` (= `d.VC_ASSY_PART_NUMBER`) |
| `[z,2]` | qty | `fieldbyname('PQTY')` |

(`excel.visible := TRUE` here, vs FALSE for the T/W variants — minor.)

## 3. Data procs (verbatim from live dump `DB Schema/CreateInventory.sql`)

### `REPORT_DailyShipping` (single date) — `schema:2875`
```sql
SELECT  s.VC_START_SEQ_NUMBER 'Start', s.VC_END_SEQ_NUMBER 'End', s.IN_QTY 'Vehicle Count',
        m.VC_PART_NUMBER 'Part Number', m.VC_PARTS_NAME 'Desc', SUM(p.IN_QTY) 'PQty'
FROM INV_SHIPPING_INF s
     JOIN INV_PART_SHIPPING_INF p ON s.VC_PRODUCTION_DATE = p.VC_PRODUCTION_DATE   -- (!) date-only join
     JOIN INV_PARTS_STOCK_MST   m ON p.VC_PART_NUMBER     = m.VC_PART_NUMBER
WHERE s.VC_PRODUCTION_DATE = @PDate
GROUP BY m.in_part_type_ID, m.VC_PART_NUMBER, s.VC_START_SEQ_NUMBER, s.VC_END_SEQ_NUMBER, s.IN_QTY, m.VC_PARTS_NAME
ORDER BY m.in_part_type_ID, m.VC_PART_NUMBER
```

### `REPORT_DailyShippingRange` — `schema:2838`
```sql
SELECT  SUM(s.IN_QTY) 'Vehicle Count', m.VC_PART_NUMBER 'Part Number', m.VC_PARTS_NAME 'Desc',
        SUM(p.IN_QTY) 'PQty'
FROM INV_SHIPPING_INF s
     JOIN INV_PART_SHIPPING_INF p ON s.VC_PRODUCTION_DATE = p.VC_PRODUCTION_DATE   -- (!) date-only join
     JOIN INV_PARTS_STOCK_MST   m ON p.VC_PART_NUMBER     = m.VC_PART_NUMBER
WHERE s.VC_PRODUCTION_DATE between @BeginPDate and @EndPDate
GROUP BY m.in_part_type_ID, m.VC_PART_NUMBER, m.VC_PARTS_NAME
ORDER BY m.in_part_type_ID, m.VC_PART_NUMBER
```
(`Start`/`End` cols are commented out in this proc, `schema:2846-2847` — hence §2b has no Start/End cells.)

### `REPORT_DailyShippingAssy` — `schema:3576`
```sql
SELECT  s.VC_START_SEQ_NUMBER 'Start', s.VC_END_SEQ_NUMBER 'End', s.IN_QTY 'Vehicle Count',
        d.VC_ASSY_PART_NUMBER 'Part Number', SUM(d.IN_QTY) 'PQty'
FROM INV_ASN_MST s JOIN INV_ASN_DETAIL_MST d ON s.IN_ASN_ID = d.IN_ASN_ID         -- clean key join
WHERE s.VC_PRODUCTION_DATE = @PDate
GROUP BY d.VC_ASSY_PART_NUMBER, s.VC_START_SEQ_NUMBER, s.VC_END_SEQ_NUMBER, s.IN_QTY, d.IN_QTY
ORDER BY d.VC_ASSY_PART_NUMBER
```
(Declares `@Line` is NOT a parameter of the proc — handler passes `@Line` at `MainMenu.pas:3158` but
the proc signature is `@Pdate` only. Extra named param is benign under ADO positional binding today,
but a **name-binding port will fail** — flag for the rebuild. The Assy proc joins on the real ASN key
`IN_ASN_ID`, so it does NOT have the §4 fan-out bug; its only risk is the `IN_QTY` in GROUP BY, which
splits a part across distinct line-qty values rather than summing — minor.)

## 4. WHY IT FAILS — the root cause (proc-side, NOT Excel)

**The T/W procs (`REPORT_DailyShipping`, `REPORT_DailyShippingRange`) join
`INV_SHIPPING_INF s` to `INV_PART_SHIPPING_INF p` on `VC_PRODUCTION_DATE` ALONE — a cartesian
fan-out — and `INV_SHIPPING_INF` routinely has MULTIPLE rows per production date.**

Evidence the grain is many-rows-per-date:
- `INV_SHIPPING_INF` PK is `IN_SHIPPING_ID IDENTITY` (`schema:17,28`) — **not** keyed on date. Nothing
  enforces one row per date.
- `INSERT_ShippingInfo` (`schema:442-461`) inserts one row per `(@LineName,@StartSeq,@EndSeq,@Date)` —
  i.e. **one row per assembly line and per sequence range**. A normal day ships the **tire line and the
  wheel line** (≥2 rows), and split/continued shipments add more (`IN_CONTINUE_NUMBER` exists for
  exactly this). Multiple `DataModule.pas` callers confirm repeated inserts (`:4237,:4684,:4930`).
- `INV_PART_SHIPPING_INF` is also keyed per `(part, date)` (`schema:270-275`), one set of part rows per
  date.

So with N shipping-header rows for a date and M part rows for that date, the date-only join produces
**N × M rows BEFORE grouping**. The `SUM(p.IN_QTY)` then sums each part's qty **N times** →
**part quantities are inflated by the number of `INV_SHIPPING_INF` rows that day** (typically ×2 for a
tire+wheel day, more on split shipments). The `GROUP BY` in `REPORT_DailyShipping` includes the
shipping-header columns (`VC_START_SEQ_NUMBER, VC_END_SEQ_NUMBER, IN_QTY`), so instead of one row per
part you get **one row per (part × shipping-header)** — i.e. the same part repeated once for each
sequence-range header, each carrying the part's full qty. The Range proc drops the header columns from
GROUP BY but keeps the date-only join, so it **collapses to one row per part but with `PQty` summed
across all the duplicate header joins** → still N× inflated.

### Two distinct observable failures (so "it fails" can mean either)

1. **Wrong numbers (silent, the common case).** When ≥2 shipping headers exist for the date (the normal
   tire+wheel day), part `Qty` is multiplied / parts appear duplicated. The report "runs" but its
   numbers are wrong — which is exactly an M3 gate failure ("report numbers must match legacy" — except
   here legacy itself is wrong, so the rebuild must NOT copy the bug; see §6).
2. **Hard error path.** The handler wraps everything in `try/except` and logs
   `'Failed on Daily Shipping Report, '+e.Message` (`:3114`). Any OLE/template error (missing
   `ReportTemplate.xls` in `TemplateDir`, no client Excel, `SaveAs` to an unwritable
   `fiReportsOutputDir`) lands here. The "Print this report?" + `PrintOut` path also fails if no
   printer. These are the **environment-coupling failures** the M3 Excel retirement removes outright.

**Net:** the proc is the load-bearing defect (data correctness); Excel/OLE is the brittleness. The
single-date `REPORT_DailyShipping` is the worst because its GROUP BY also **explodes rows** per header.

> **Self-flag for the next step (data-dependent):** the exact inflation factor = the count of
> `INV_SHIPPING_INF` rows for the chosen `@PDate`. To confirm against the golden, pick a production date
> and run `SELECT COUNT(*) FROM INV_SHIPPING_INF WHERE VC_PRODUCTION_DATE='<yyyymmdd>'`; if that count
> is K, legacy `PQty` for each part = K × the true part qty (single-date proc repeats each part K times;
> Range proc sums to K× in one row). Verify on one real date before asserting the factor in the rebuild
> parity test. Compare to the truth: `SELECT VC_PART_NUMBER, SUM(IN_QTY) FROM INV_PART_SHIPPING_INF
> WHERE VC_PRODUCTION_DATE='<yyyymmdd>' GROUP BY VC_PART_NUMBER`.

## 5. The corrected grain (what the numbers SHOULD be)

`INV_PART_SHIPPING_INF` already holds the authoritative per-part shipped qty per date (it is what the
`InsertPartShipping`/`DeletePartShipping` triggers decrement/restore stock by, `schema:2904-2925`). The
true Daily Shipping qty is simply:

```sql
SELECT m.VC_PART_NUMBER 'Part Number', m.VC_PARTS_NAME 'Desc', SUM(p.IN_QTY) 'PQty'
FROM INV_PART_SHIPPING_INF p JOIN INV_PARTS_STOCK_MST m ON p.VC_PART_NUMBER = m.VC_PART_NUMBER
WHERE p.VC_PRODUCTION_DATE = @PDate                 -- (Range: BETWEEN @BeginPDate AND @EndPDate)
GROUP BY m.in_part_type_ID, m.VC_PART_NUMBER, m.VC_PARTS_NAME
ORDER BY m.in_part_type_ID, m.VC_PART_NUMBER
```
i.e. **drop `INV_SHIPPING_INF` from the qty join entirely.** The header values (Start/End seq, Vehicle
Count) come from `INV_SHIPPING_INF` separately and should be **aggregated** for the date
(MIN start / MAX end / SUM vehicle count across that date's headers), shown once in the report header —
not joined into the detail.

## 6. Rebuild target — render WITHOUT Excel (server-side)

**Columns to render (the human-visible contract):**
- Header band: Report title; Production date (or Begin/End for Range); **Start Seq** (MIN over date),
  **End Seq** (MAX over date), **Vehicle Count** (SUM `IN_QTY` over date).
- Detail table: `Part Number` | `Part Description` | `Qty` — one row per part, sorted by
  `in_part_type_ID, VC_PART_NUMBER` (tires then wheels, then by number).

**Two Named Queries** (per the per-proc Named-Query practice):
- `daily_shipping_parts` — the **corrected** part-qty query (§5), params `@PDate` (or begin/end). This
  is the detail. Do NOT port the date-only `INV_SHIPPING_INF` join.
- `daily_shipping_header` — `SELECT MIN(VC_START_SEQ_NUMBER), MAX(VC_END_SEQ_NUMBER), SUM(IN_QTY)
  FROM INV_SHIPPING_INF WHERE VC_PRODUCTION_DATE = @PDate` (range: BETWEEN) for the header band.

**Consumer:** Perspective view = production-date picker (single + range toggle) → header labels bound to
`daily_shipping_header` → Table bound to `daily_shipping_parts` → CSV/Excel **export** button (replaces
SaveAs) + optional Reporting-module PDF template for operators who print. No client Excel, no
`ReportTemplate.xls`, no OLE, no PrintOut.

**Assy variant (R3/R4):** keep the existing clean `IN_ASN_ID` join (§3, no fan-out); just remove
`d.IN_QTY` from the GROUP BY so a part sums across its line-qty splits, and render the 2-column
(`Part Number` | `Qty`) layout per §2c. Reconcile to the 856 ASN.

## 7. DIVERGENCE — flag for David (changes a number Toyota/operators see)

The corrected query (§5) makes the Daily Shipping qty **smaller than legacy** (legacy inflates by the
day's `INV_SHIPPING_INF` row count). This is "more correct ≠ equivalent": the rebuild number will NOT
match the legacy Excel for any date with >1 shipping header. Per the divergence rule, this must be an
**explicit David decision recorded in the unit's ledger** (style of `edi810-decisions.md`), because it
changes a reported shipped quantity — even though the legacy value is provably wrong (the fan-out is a
bug, the triggers prove `INV_PART_SHIPPING_INF` is the source of truth). Recommended decision: **adopt
the corrected grain** (parity target = the truth in §4's confirmation query, NOT the legacy Excel
output). Confirm on one real date first (§4 self-flag).

## 8. Confidence / residual checks

- HIGH on the fan-out mechanism: all 3 procs read verbatim; `INV_SHIPPING_INF` PK + `INSERT_ShippingInfo`
  per-line grain verified; triggers confirm `INV_PART_SHIPPING_INF` is the stock-authoritative qty.
- To CONFIRM the daily-log error specifically (the "failing path" report): check the operator's
  `INV_ACTIVITY_LOG` for `ERROR` rows with text `'Failed on Daily Shipping Report'` /
  `'…Range Report'` (the exact `LogActLog` strings, `MainMenu.pas:3114/3008`). If those exist, the
  observed failure is the §4.2 environment path (Excel/OLE/printer); if instead users report "the
  numbers don't add up / parts doubled," it is the §4.1 fan-out. Both are addressed by §6.
