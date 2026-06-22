# M3 Report + Excel-Layer Retirement — INVENTORY (work-list)

**Phase:** M3 source-truth (2026-06-21) **Analyst:** Claude **Scope:** retire the last Excel/OLE
paths; render report families server-side (8.3 Reporting module / 8.1 Perspective-export fallback);
additive + read-only; gate = report numbers match legacy.

> Builds on `reporting.md` (D6 proc catalog — proc bodies/window-blind bugs are specced there; not
> re-derived here). This doc is the **build work-list**: per path, the unit (LIVE-confirmed), trigger,
> data proc(s), Excel/OLE mechanism, output filename/dir, and complexity/priority. The Daily Shipping
> deep-dive (the failing path) is its own file: `daily-shipping-report-spec.md`.

## 0. How to read this

- **TRUE report** = read-only render of a `REPORT_*`/`SELECT_*` proc. M3 retires its Excel.
- **Excel companion of a transactional unit** = the `.ord`/856/810/.frc machine artifact is the real
  output (already built in M0–M2); the **Excel/.xls companion is the human-readable copy**. M3 retires
  the companion only — do NOT touch the machine artifact.
- **Mechanism legend:** `OLE-template` = `CreateOleObject('Excel.Application')` → open
  `TemplateDir+'<T>.xls'` → write `mysheet.cells[r,c]` → `SaveAs`. `OLE-CSV` = same but `SaveAs(...,xlCSV)`.
  `OLE-read` = opens an Excel file the user picked, **reads** cells in (an input path, not output).
  `text` = Pascal `Rewrite`/`Writeln` flat file (no Excel).
- Output dir tokens: `fiReportsOutputDir` (INI `[DIRECTORIES]`), `TemplateDir` (template `.xls` source),
  per-supplier `Directory` (`INV_SUPPLIER_MST.VC_BREAKDOWN_ORDER_DIRECTORY`) / `LogisticsDirectory`.
- Filename suffix `formatdatetime('yyyymmddhhmmss00',now)` = 14 digits + literal `00` (NOT centiseconds;
  see `reporting.md §4.4`).

## 1. Dead-code exclusions (do NOT build)

Confirmed absent from `InventorySystem.dpr`:
- **`Reports.pas`/`.dfm`** — dead QuickReport scaffolding (2 OLE sites). Not a real report.
- **`ForecastBreakDown.pas`** (note: no `F`) — dead twin of live `ForecastBreakdownF.pas` (2 OLE sites).
- **`OrderFormCreate.pas`** (no `F`) — dead twin of live `OrderFormCreateF.pas` (1 OLE site).
- `MonthlySupplerOrderReport` / `MonthlySupplerInvoiceReport` / `MonthlyLogiticsOrderReport` — compiled
  (in `.dpr`) but **every MainMenu call is commented out** (`reporting.md §1b`). DEAD-AT-CALLSITE; use
  the `TQuickRep` `.dfm` only as a **column/header layout reference** for the Monthly report templates.
- Deprecated (D9): `REPORT_ASNWithCost` (`ASNwithCost*.xls`, `MainMenu.pas:3261/3313`),
  `ForecastCamexreport.pas` (`-CFForecast`, CAMEX decommissioned), `REPORT_NUMMILotLocation[W]`.
  Out of rebuild scope.

## 2. THE WORK-LIST

### 2a. TRUE reports — MainMenu Excel/OLE families (the M3 mainstream)

All: `Inv_DataSet` on a `REPORT_*` proc → `OLE-template` (`ReportTemplate.xls`) → `mysheet.cells` →
`SaveAs(fiReportsOutputDir + '\<name>'+ts+'.xls')` → optional `mysheet.PrintOut`. Param entry via the
shared `TDateSelectDlg` (`MonthlyReportSelect.pas`, LIVE) or `TProductionDateSelectDlg`
(`ProductionDates.pas`, LIVE). Trigger = MainMenu Reports menu item.

| # | Report | Handler (`MainMenu.pas:`) | Data proc | Output `<name>` | Cmplx | Prio |
|---|--------|---------------------------|-----------|-----------------|-------|------|
| R1 | **Daily Shipping (T/W)** | `DailyShippingClick` :3029 | `REPORT_DailyShipping` | `DailyShippingTW` | M | **P0 (FAILING)** |
| R2 | **Daily Shipping Range (T/W)** | `DailyShippingRangeTireWheelPartNumbersClick` :2919 | `REPORT_DailyShippingRange` | `DailyShippingRangeTW` | M | **P0 (FAILING)** |
| R3 | Daily Shipping ASN (Assy) | `DailyASNReportClick` :3135 | `REPORT_DailyShippingAssy` | `DailyShippingAssy` | M | P1 |
| R4 | Monthly Shipping ASN (Assy) | (~`MonthlyASNClick` :3356) | `REPORT_MonthlyShippingAssy` | `DailyShippingAssy`* | M | P1 |
| R5 | Daily Supplier Order | (~:1064/:1187) | `REPORT_DailySupplierOrders` / `…Cost` | `DailySupplierOrder` | M | P1 |
| R6 | Monthly Supplier Order (±cost) | `MonthlyOrderSummaryClick` :1403 (+cost ~:1280) | `REPORT_MonthlySupplierOrders` / `…Cost` | `SupplierOrder` | M | P1 |
| R7 | Monthly Logistics Order | `MonthlyLogisticsClick` :1513 | `REPORT_MonthlyLogisticsOrders` | `Logistics` | M | P2 |
| R8 | Monthly Supplier Invoice | `MonthlySupplierInvoice1Click` :1629 | `REPORT_MonthlySupplierInvoices` | `SupplierInvoice` | M | P2 |
| R9 | **INVOICE Summary (D6)** | `INVOICEReportClick` :3592 | `REPORT_INVOICESSummary` | `INVOICESummary` | H | **P1 (D6 window-aware)** |
| R10 | **Monthly INVOICE Summary (D6)** | (~:3487) | `REPORT_MonthlyINVOICESSummary` | `MonthlyINVOICESummary` | H | **P1 (D6 window-aware)** |
| R11 | Logical Inventory | `LogicalInventoryReport1Click` :919 | `REPORT_LogicalInventory` | `LogicalInventory` | M | P2 |
| R12 | Lot Location (PLANT) | `LotLocationClick` :716 (+`…W` 791/829) | `REPORT_PLANTLotLocation`/`…W` | `<PlantName>LotLocation` | M | P2 (print template) |
| R13 | Empty Container | `EmptyContainerClick` :1734 | `REPORT_EmptyContainer` | `EmptyContainer` | M | P3 |
| R14 | Past-Due / Late FRS | (~:2425) | `REPORT_LATEFRS` | `PastDueFRS` | L | P3 |
| R15 | PO Report | `POReportClick` :2773 | `REPORT_PO` | `POReport` | L | P3 |
| R16 | Forecast Parts Summary | (~:1875) | `REPORT_ForecastPartsSummary` | `ForecastPartsSummary` | M | P3 |
| R17 | Forecast Assy Summary | (~:2007) | `REPORT_ForecastSummary` | `ForecastAssySummary` | M | P3 |
| R18 | Forecast Detail | (~:3728/:3799) | `REPORT_ForecastDetail` | `ForecastDetail` | M | P3 |
| R19 | Forecast vs Usage | (~:2381) | (forecast/usage proc — verify) | `ForecastvsUsage` | M | P3 |
| R20 | Unused Tire Part Numbers | tire (:627) | `REPORT_UnusedTirePartNumbers` | `TireWithoutAssembly` | L | P3 |
| R21 | Unused Wheel Part Numbers | wheel (:677) | `REPORT_UnusedWheelPartNumbers` (D11 bug: checks TIRE col) | `WheelWithoutAssembly` | L | P3 |

\* R4 (Monthly ASN) reuses the `DailyShippingAssy` filename at :3424 — output-name collision with R3,
relies on the timestamp suffix to disambiguate. Note for the rebuild (give it a distinct name).

### 2b. The one live on-screen QuickReport

| # | Report | Unit (LIVE) | Trigger | Data source | Mechanism | Prio |
|---|--------|-------------|---------|-------------|-----------|------|
| R22 | InvMgmt QReport | `InvMgmtQReport.pas` | `InvMgmt.pas:227-229` (`.Preview`) | **`Data_Module.Grid_ClientDataSet`** (the InvMgmt grid dataset, NOT a `REPORT_*` proc — `.pas:84`) | `TQuickRep.Preview` (on-screen, no Excel) | P2 |

> Correction to `reporting.md §1b`: the live `.pas` binds the QR bands to `Grid_ClientDataSet`
> (`InvMgmtQReport.pas:84`), i.e. the same client dataset InvMgmt already filled — confirmed in code.
> Rebuild = render InvMgmt's columns as a Perspective table/Reporting template fed by InvMgmt's query.
> No proc to wrap. (No Excel dependency today — lowest-risk of the report set, but no .xls to retire.)

### 2c. Excel COMPANIONS of transactional units (artifact already built; retire the .xls only)

| # | Path | Unit (LIVE) | Trigger | Machine artifact (KEEP) | Excel companion (RETIRE) | Mechanism | Prio |
|---|------|-------------|---------|-------------------------|--------------------------|-----------|------|
| C1 | **Order sheet / breakdown** | `OrderFormCreateF.pas` | from Order/Renban flow | `.ord` text (`:290/321`) | `OS<sup>-<code>-<renban>.xls` (no dir-ts) → per-supplier `Directory`/`LogisticsDirectory`/`Archive` | `OLE-template` `OrderTemplate.xls` + `OrderSheetTemplateTire/Wheel.xls` (`:245/247/274/384/499`) | P2 |
| C2 | **Order simulation** | `Order.pas` | Order screen sim | (on-screen sim) | `OrderSimulation.xls` opened for layout (`:183`) | `OLE-template` `OrderSimulation.xls` | P3 |
| C3 | **Forecast breakdown (.frc)** | `ForecastBreakdownF.pas` | Forecast breakdown run | `.frc` text + Archive (`:450-457`) | `<sup>-<code>-Forecast.xls` → per-supplier `Directory` (`:441/523`) | `OLE-template` `ForecastTemplate.xls` (Excel) **+ `text` (.frc)** | P2 **(has P6 crash — see §3)** |
| C4 | **EDI 856 ASN + 810 INV CSV** | `DailyBuildTotal.pas` | `MainMenu.pas:2570` (fmASN), :2674 (fmINVOICE) | the 856/810 EDI files (M1/M2) | `ASN<ts><n>.csv` (`:319/430`), `INV<ts><n>.csv` (`:499/582`) → `fiReportsOutputDir` | `OLE-CSV` (`SaveAs(...,xlCSV)`) | P2 |
| C5 | **862 Firm Order echo** | `EDIUpload.pas` | EDI inbound import (data='862', `:105`) | (inbound 862 consumed → forecast tables) | `FirmOrder<ts>.xls` → `fiReportsOutputDir` (`:178`) | `OLE-template` `ReportTemplate.xls`, written by **parsing the 862 X12 text directly** (no proc — `:140-174`) | P1 |
| C6 | EDI 861 Receiving Advice echo | `EDIUpload.pas` | inbound import (`:260`) | (inbound consumed) | `ReceivingAdvice<ts>.xls` (`:294`) | `OLE-template` `ReportTemplate.xls` | P3 |
| C7 | EDI 820 Remittance echo | `EDIUpload.pas` | inbound import (`:332`) | (inbound consumed) | `Remittance<ts>.xls` (`:406`) | `OLE-template` `ReportTemplate.xls` | P3 |
| C8 | Forecast CAMEX | `ForecastCamexreport.pas` | `MainMenu.pas:3839` | n/a | `<sup>-<code>-CFForecast.xls` | OLE-template `ForecastCamexTemplate.xls` | **OUT (D9 deprecated)** |

> C5 (862 FirmOrder) and C6/C7 are **read-side echoes**: EDIUpload parses an inbound X12 file and
> writes a human-readable Excel copy. There is **no `REPORT_*` proc** — the "report" content is the
> parsed X12 (fixed-offset `copy(fcl,...)`). The rebuild renders the parsed inbound EDI, not a query.
> C5 is the deferred "862 FirmOrder.xls" from M1/M2.

### 2d. ManualForecast (OLE-READ, an INPUT path — note, don't "retire as report")

| # | Path | Unit (LIVE) | Trigger | Mechanism | Note |
|---|------|-------------|---------|-----------|------|
| — | Manual forecast import | `ManualForecast.pas` | `ManualForecast_ButtonClick` `MainMenu.pas:2519` (button visible only if `fiBuildOut`) | `OLE-READ`: `excel.workbooks.open(fFilename)` (`:70`) reads a user-picked `.xls` IN | This is an **import**, not a report — operator hands it a spreadsheet of forecast qty. Excel-dependency to retire, but the rebuild target is an **upload/parse**, not a render. Flag to M-import work, not M3-report. |

## 3. Latent-crash + hazard flags (first-class findings)

- **P6 — `.frc` `SiteSupplierCode` crash (CONFIRMED COLD).** `ForecastBreakdownF.pas:488`
  reads `fieldbyname('SiteSupplierCode').AsString` in the **text (.frc)** branch whenever
  `sendsite` is true. `sendsite` is set from `'Site Number in Order'` (`:432`), which is the supplier
  flag `BIT_SITE_NUMBER_IN_ORDER` emitted by `SELECT_SupplierInfo` (`CreateInventory.sql:6026/6069`).
  **But the dataset being iterated is `SELECT_ForecastDetail` (`:745/761`), whose result set has NO
  `SiteSupplierCode` column** (`CreateInventory.sql:3154+` emits forecast-detail cols only; grep of
  the whole schema for `SiteSupplierCode` = **0 hits**). → For any supplier with "Site Number in
  Order" ON producing a `.frc`, `fieldbyname('SiteSupplierCode')` raises **"Field not found"**, caught
  by the handler → forecast `.frc` fails for that supplier. **This is the same defect class as the
  order-side `.frc`.** Confidence: HIGH (both sides verified — Pascal call + proc column list + schema
  grep). The rebuild's forecast `.frc` must source the site-supplier code from the supplier master
  (the master's site mapping), not from a non-existent forecast-detail column.
- **D6 window-blind pricing (R9/R10).** `REPORT_INVOICESSummary` / `REPORT_MonthlyINVOICESSummary`
  join `INV_MANIFEST_COST_MST` on assy code **with no start/end-manifest window** vs production date
  (`reporting.md §4.1`). Rebuild MUST use the window-aware manifest-cost lookup (the 856 predicate),
  shared with EDI 810/856. Resolved decision: D11 (make window-aware).
- **R21 Unused-Wheel bug (D11).** `REPORT_UnusedWheelPartNumbers` filters against the **TIRE** code
  column. Rebuild uses the wheel column.
- **R4/R3 filename collision.** Monthly ASN (:3424) and Daily ASN (:3199) both `SaveAs ...DailyShippingAssy`.
- **Param name vs positional binding (whole family).** ADO binds positionally today; a Named-Query
  port that binds by name will break where caller param names ≠ proc param names (`reporting.md §4.2`).
- **R1/R2 — the FAILING Daily Shipping path.** Root cause is a **proc-side join/grain bug** (cartesian
  fan-out across multiple `INV_SHIPPING_INF` rows per date), not Excel. Full analysis in
  `daily-shipping-report-spec.md`.

## 4. Recommended M3 BUILD ORDER

1. **R1 + R2 — Daily Shipping (T/W) & Range.** P0 because the plan flags it as the failing path and it
   is a daily operational report. Fixing it requires a corrected query (see deep-dive), so it is BOTH
   the highest priority AND a clean first proof of the "wrap-proc-as-Named-Query + render server-side"
   pattern (3 columns, no cost). Do these first; they validate the M3 render harness end to end.
2. **R3/R4 — Daily/Monthly ASN (Assy).** Same shipping domain; R3's proc has the **same grain bug** as
   R1 (verify against golden), so build alongside R1 to reuse the corrected pattern.
3. **R9/R10 — INVOICE Summary (D6).** Highest business risk (billing numbers Toyota sees). Build with
   the window-aware manifest-cost lookup; reconcile to the 810/856 path. Flag any divergence in the D6
   ledger.
4. **R5/R6/R7/R8 — Supplier/Logistics order + invoice families.** Straight proc-wrap renders; bulk of
   the table+export work. (D12: monthly order reports should range on **order date**, not ship date —
   confirm with David before changing a number Toyota sees.)
5. **C5 (862 FirmOrder) + C6/C7 EDI echoes.** Render the parsed inbound X12; no proc.
6. **C1/C3 order-sheet + forecast-breakdown companions** (fix P6 in C3 while here).
7. **R11–R21 + R22 (InvMgmt QR), C2/C4.** Lower-frequency / print-template / already-no-Excel.

Out of scope (D9): C8 ForecastCAMEX, R-ASNwithCost, NUMMILotLocation[W], `ManualForecast` (import not
report — route to import work).
