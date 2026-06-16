# Module Analysis: Reporting

**Area:** Reporting  **Status:** ✅ spec complete (pending adversarial verify)  **Analyst:** Claude / 2026-06-16

> Read in full: monthly/daily order/invoice/cost procs + both invoice-summary procs (D6 focus). Other
> `REPORT_*` read and summarized by category. Confidence stated; "body unverified" where noted.
> **Scope correction:** `Reports.pas`+`Reports.dfm` are **DEAD CODE** — not in `InventorySystem.dpr`, and
> `grep -ril "\bReports\b" *.pas` returns only `Reports.pas` itself. The real hub is **`MainMenu.pas`**
> (Reports menu items `MainMenu.pas:80–176`; handlers `:627–3728`), which drives reports as **Excel/OLE
> exports**, not QuickReport.

## 1. Legacy surface

Two live mechanisms + dead scaffolding.

**1a. Excel/OLE reports — the live mainstream (`MainMenu.pas`).** Each handler: opens `Inv_DataSet` on a
`REPORT_*` proc → `CreateOleObject('Excel.Application')` → opens `TemplateDir+'ReportTemplate.xls'` → writes
cells → `SaveAs` to `fiReportsOutputDir + '\<Name>'+formatdatetime('yyyymmddhhmmss00',now)+'.xls'` (literal
trailing `00`, not centiseconds, e.g. `MainMenu.pas:1138`) → optional `mysheet.PrintOut`. Logs
`LogActLog('REPORT'/'ERROR', …)`.

Handler→proc map: `MonthlyOrderSummaryClick`(1403)→`REPORT_MonthlySupplierOrders`; cost variant(~1280)→
`REPORT_MonthlySupplierOrdersCost`; `MonthlyLogisticsClick`(1513)→`REPORT_MonthlyLogisticsOrders`;
`MonthlySupplierInvoice1Click`(1629)→`REPORT_MonthlySupplierInvoices`; daily order(~1187)→
`REPORT_DailySupplierOrders`; daily cost(~1064)→`REPORT_DailySupplierOrdersCost`; `LogicalInventoryReport1Click`
(919)→`REPORT_LogicalInventory`; `LotLocationClick`(716)→`REPORT_PLANTLotLocation`/`…W`(791/829); unused
tire(627)/wheel(677); `EmptyContainerClick`(1734)→`REPORT_EmptyContainer`; forecast parts(~1875)/summary
(~2007)/detail(~3728); late FRS(~2425)→`REPORT_LATEFRS`; `POReportClick`(2773)→`REPORT_PO`; `DailyShippingClick`
(3029); range(2919); `DailyASNReportClick`(3135)→`REPORT_DailyShippingAssy`; `ASNReportWithCostAssyPartNumbers1Click`
(3242)→`REPORT_ASNWithCost` (snapshot-absent); `MonthlyASNClick`(3356)→`REPORT_MonthlyShippingAssy`;
`INVOICEReportClick`(3592)→`REPORT_INVOICESSummary`; monthly invoice(~3487)→`REPORT_MonthlyINVOICESSummary`;
EDI810(2603+`ASNInvoice`); ForecastCAMEX(`MainMenu.pas:3839`, `ForecastBreakdownF.pas:163`)→
`REPORT_ForecastCAMEXReport` (snapshot-absent).

**1b. QuickReport forms — partly live, partly dead.**
- `InvMgmtQReport.pas`/`.dfm` — **LIVE**. A real `TQuickRep` invoked from `InvMgmt.pas:227–229`. The live
  `.dfm` binds its QR bands to `Data_Module.Inv_DataSet` (`.dfm:30,471+`) — a dataset InvMgmt already filled,
  **not** a `REPORT_*` proc. (`Grid_ClientDataSet` appears only in the commented-out `FormCreate` field-binding
  block, `:82–93`.) The only true on-screen QuickReport.
- `MonthlySupplerOrderReport.pas` / `MonthlySupplerInvoiceReport.pas` / `MonthlyLogiticsOrderReport.pas` —
  **DEAD-AT-CALLSITE**. In `.dpr`, each defines a `TQuickRep` + `Execute` (proc + `.Preview`), but **every
  MainMenu call is commented out** (`MainMenu.pas:1413–1417`, `1523–1527`, invoice equiv). Compiled but never
  reached; use only as layout reference.
- `MonthlyReportSelect.pas` (`TDateSelectDlg`) — **LIVE** shared param picker. `FromDate`/`ToDate` + optional
  `Supplier`/`Logistics`/`PartNumber` combos (toggled by `Do*`) + `JustStart`. Combos from `INV_PARTS_STOCK_MST`,
  `INV_SUPPLIER_MST.VC_SUPPLIER_NAME`, `INV_LOGISTICS_MST.VC_LOGISTICS_NAME`, each prepended `'ALL'`; blank
  coerces to `'ALL'` on close. `Cancel` set by button **and** any setup exception.

**1c. Dead scaffolding:** `Reports.pas`/`.dfm` — do not spec.

**Purpose:** read-only operational/financial reporting over inventory + EDI data — monthly/daily supplier-order
ledgers (±cost), logistics roll-ups, supplier invoices, 810/856 billing extracts, inventory/lot-location
snapshots, forecast roll-ups, empty-container, late-FRS, assy/shipping summaries. Output almost entirely Excel
`.xls` (saved + optional print via OLE), plus one live `TQuickRep` and a forecast-CAMEX Excel builder.

## 2. Data touched
All read-only SELECT; **no report proc writes; no triggers participate.** Tables: `INV_OPEN_ORDER_INF`,
`INV_SUPPLIER_MST`, `INV_LOGISTICS_MST`, `INV_PARTS_STOCK_MST` (`MO_PART_COST`), `INV_PART_TYPE_MST`,
`INV_SIZE_MST`, `INV_RENBAN_GROUP_MST`, `INV_INVOICE_INF`, `INV_ASN_MST`/`INV_ASN_DETAIL_MST`, `INV_INV_MST`,
**`INV_MANIFEST_COST_MST` (D6)**, `INV_FORECAST_INF`/`INV_FORECAST_DETAIL_INF`, `INV_BREAKDOWN_FC_INF`,
`INV_SHIPPING_INF`/`INV_PART_SHIPPING_INF`, `INV_ASSY_MONTHLY_PO`.

## 3. `REPORT_*` proc catalog (~29; 27 in snapshot, 2 live callers reference absent procs)
All `schema:` = `DB Schema/Create Inventory.sql`.

**Category A — Supplier/Logistics ORDER reports (over `INV_OPEN_ORDER_INF`):**
- `REPORT_DailySupplierOrders` (4180) `@StartDate,@Supplier='ALL'` — `WHERE VC_ORDER_DATE=@StartDate`; no cost.
- `REPORT_DailySupplierOrdersCost` (4238) — same, ANSI joins, adds `P.MO_PART_COST` + `MO_PART_COST*IN_QTY AS Total`.
- `REPORT_MonthlySupplierOrders` (4997) `@StartDate,@Enddate,@Supplier='ALL'` — **range on
  `VC_STATUS_SUPPLIER_SHIPPING`, NOT order date.**
- `REPORT_MonthlySupplierOrdersCost` (5056) — same shipping-date range (order-date variant commented out
  `:897/917`); + cost/Total.
- `REPORT_MonthlyLogisticsOrders` (4838) `@StartDate,@Enddate,@Logistics='ALL'` — O→S→L joins; shipping-date
  range; group by logistics/shipping/renban.

**Category B — INVOICE / billing (manifest-cost priced) — D6 FOCUS:**
- `REPORT_MonthlySupplierInvoices` (4936) `@StartDate,@Enddate,@Supplier=''` — over `INV_INVOICE_INF`, prices
  from the invoice rows' own `MONEY_UNIT_PRICE`/`MONEY_TOTAL_AMOUNT`. **NOT D6** (no manifest-cost touch).
- **`REPORT_INVOICESSummary` (4681) `@PDate varchar(13)` — D6 WINDOW-BLIND.** Joins `INV_MANIFEST_COST_MST m ON
  d.VC_ASSY_PART_NUMBER=m.VC_ASSY_PART_NUMBER_CODE` (`schema:4694-4695`), `WHERE a.VC_PRODUCTION_DATE=@PDate`,
  emits `m.MO_PRICE`.
- **`REPORT_MonthlyINVOICESSummary` (4802) `@PDate varchar(6)` — D6 WINDOW-BLIND.** Same join (`schema:4814-4815`),
  `WHERE substring(VC_PRODUCTION_DATE,1,6)=@PDate`.
- `REPORT_EDI810` (4303) — D6 window-blind (already in `edi/asn-invoice.md §4.1`; cross-ref only).
- `REPORT_EDI856` (4377) — **window-AWARE** (`m.VC_START_MANIFEST < a.VC_PRODUCTION_DATE AND m.VC_END_MANIFEST >
  a.VC_PRODUCTION_DATE`, `schema:4396-4397`) — the correct pattern.

**Category C — Inventory / lot-location:**
- `REPORT_LogicalInventory` (4747) `@PartNo='ALL'` — `INV_PARTS_STOCK_MST` LEFT-joined supplier/logistics/renban/
  part-type/size.
- `REPORT_PLANTLotLocation`/`…W` (5212/5256) — *body unverified*; `W`=warehouse variant. `REPORT_NUMMILotLocation`/
  `W` (5123/5165) are **dead site-specific twins** (MainMenu calls only PLANT, 791/829).
- `REPORT_UnusedTirePartNumbers` (5329) — tire parts not in any forecast-detail.
- `REPORT_UnusedWheelPartNumbers` (5354) — **BUG (verified):** filters `vc_part_number NOT IN (SELECT
  vc_tire_part_number_code…)` — checks the **TIRE** code column for wheel parts.

**Category D — Shipping / ASN summaries:**
- `REPORT_DailyShipping` (4048) `@Pdate`; `REPORT_DailyShippingRange` (4138) `@BeginPdate,@EndPdate`;
  `REPORT_DailyShippingAssy` (4093) `@Pdate`; `REPORT_MonthlyShippingAssy` (4895) `@Pdate varchar(6)`;
  `REPORT_AvailableProductionDates` (3991) `@Line,@INVOICE,@ASN,@Month=0` (date-picker feed).
- `REPORT_ASNWithCost` — **ABSENT from snapshot**; called `MainMenu.pas:3261`. Name implies manifest-cost join →
  **D6 suspect, verify live.**

**Category E — Forecast:**
- `REPORT_ForecastSummary` (4659); `REPORT_ForecastPartsSummary` (4632) (`IN_QTY1..7` summed, `VC_WEEK_DATE >=
  today(112)`); `REPORT_ForecastDetail` (4611).
- `REPORT_ForecastCAMEXReport` — **ABSENT from snapshot**, `@WeekDate`; called `ForecastCamexreport.pas:104`.

**Category F — Exception / operational:**
- `REPORT_EmptyContainer` (4425) `@StartDate,@Enddate,@Supplier='ALL',@Logistics='ALL'` — 4-way ALL branch matrix;
  `VC_STATUS_EMPTY_TRAILER between … AND <>''`.
- `REPORT_LATEFRS` (4718) `@FRSDate` — open orders `vc_frs_date<@FRSDate` and all downstream status columns blank.
- `REPORT_PO` (5304) `@BeginDate,@EndDate` — `INV_ASSY_MONTHLY_PO`.
- `REPORT_EDI810Recreate` (4338) — also D6 window-blind (EDI doc); not on Reporting menu.

**Verified in full:** all of A & B, DailyShipping*, AvailableProductionDates, PO, Unused tire/wheel, LATEFRS,
EmptyContainer, Forecast*, LogicalInventory. **Body unverified:** `REPORT_PLANTLotLocation`/`W`.

## 4. Business rules & edge cases

**4.1 D6 — invoice-summary reports share the window-blind bug (CONFIRMED, NEW).** `REPORT_INVOICESSummary`
(4681) and `REPORT_MonthlyINVOICESSummary` (4802) price each line via `JOIN INV_MANIFEST_COST_MST m ON
d.VC_ASSY_PART_NUMBER = m.VC_ASSY_PART_NUMBER_CODE` with **no `VC_START_MANIFEST`/`VC_END_MANIFEST` window**
against `a.VC_PRODUCTION_DATE` — the same defect confirmed for `REPORT_EDI810`/`SELECT_INVOICEItems`
(`edi/asn-invoice.md §4.1`). Multiple price rows for one assy code → multiplied rows / wrong unit price.
**New D6 family members, not previously catalogued.** Correct contrast = `REPORT_EDI856` (applies the window).
`REPORT_MonthlySupplierInvoices` is **NOT** D6 (reads stored invoice money). Category-A cost reports price from
`INV_PARTS_STOCK_MST.MO_PART_COST` (part-master cost, not windowed) — outside D6. **`REPORT_ASNWithCost` is a
D6 *risk* pending live verification.**

**4.2 Param-name vs positional binding (verified, benign now).** Callers add params named `@FromDate`/`@ToDate`
(e.g. `MonthlySupplerInvoiceReport.pas`, `MainMenu.pas:1297-1303`) but procs declare `@StartDate`/`@Enddate`.
ADO `CmdStoredProc` binds **positionally** — works today; **a name-binding port (Named Queries) will break**
unless reconciled. Also `REPORT_MonthlySupplierInvoices` defaults `@Supplier=''` (not `'ALL'`) → else-branch
filters `VC_SUPPLIER_NAME=''` → empty; live caller never hits it (dialog coerces blank→`'ALL'`).

**4.3 Snapshot drift (verify-live, not confirmed bug).** `REPORT_ASNWithCost` (`MainMenu.pas:3261`) and
`REPORT_ForecastCAMEXReport` (`ForecastCamexreport.pas:104`) referenced by live code but absent from
`Create Inventory.sql` — consistent with prior EDI/forecast snapshot-lag findings ([[reference-schema-snapshot-vs-live]]).

**4.4 Dates.** `varchar(8)` `yyyymmdd`; reformatted `mm/dd/yy` via `substring(x,5,2)+'/'+substring(x,7,2)+'/'+
substring(x,3,2)` (e.g. `schema:4449`). Monthly *order* reports range on **ship date**
(`VC_STATUS_SUPPLIER_SHIPPING`), not order date. No 16-char audit stamps in this family; only the Delphi
filename suffix `yyyymmddhhmmss00` (14 digits + literal `00`).

**4.5** All read-only; re-running is safe; no dedup needed.

## 5. UI / UX notes
Shared `TDateSelectDlg` param entry; Excel `.xls` written to `fiReportsOutputDir`, optional
`MessageDlg('Print this report?')`; busy `ProcessPanel`. Brittle (requires client Excel + `ReportTemplate.xls`
+ OLE). Modernize: drop client Excel, render server-side.

## 6. Target design (Ignition)
Each report = **one Named Query** (per-proc practice) + a consumer:
- **On-screen/interactive (most: A, C, D, E, F):** Perspective view = param header (date range + lookup-backed
  supplier/logistics/part dropdowns) + Table bound to the Named Query + CSV/Excel **export** button. Replaces
  "Excel + optional print" and kills the client-Excel dependency.
- **Formal print/PDF (lot-location, and operators routinely printing):** Ignition **Reporting module** template
  (title/param-echo/detail/footer) on the same Named Query. The dead `Monthly*Report` QR forms are the **layout
  spec** for column order/headers.
- **`InvMgmtQReport` (live QR):** reports over the InvMgmt client dataset, not a proc → render same columns as a
  Reporting template / table export fed by the InvMgmt query.
- **D6 mandatory:** Named Queries for `REPORT_INVOICESSummary`/`MonthlyINVOICESSummary` (and `ASNWithCost` if
  confirmed) must use the **window-aware** manifest-cost lookup (the 856 predicate), shared with EDI.
- **D1:** add `site_id` predicate + site selector (gap, §8.1). Pure read; no background jobs.

## 7. Migration plan for this module
- [ ] Stage 1 — wrap each `REPORT_*` as a read-only Named Query; render A/C/D/E/F as table+export.
- [ ] Stage 2 — reimplement Category B window-aware (D6) + lot-location/InvMgmt print templates; verify the 2
  absent procs against the live DB.
- [ ] Stage 3 — add `site_id` filter (D1) + site selector; retire Excel-OLE and the dead QR forms.

## 8. Open questions for the user (domain expert)
1. **D1 cross-site (HIGH):** no `REPORT_*` filters by site/plant (the `IN_ASN_EIN`/`IN_INV_EIN` columns are
   output identity, not filters). Should reports be per-site by default on the shared DB, and which fact tables
   carry the authoritative `site_id`?
2. Monthly order range on **ship date** not order date — intended?
3. Confirm `REPORT_INVOICESSummary`/`MonthlyINVOICESSummary` should be made window-aware (D6); have downstream
   numbers been "consistent-but-wrong" vs the 810?
4. `REPORT_UnusedWheelPartNumbers` checks the TIRE column — bug?
5. Supply live bodies for `REPORT_ASNWithCost` + `REPORT_ForecastCAMEXReport`.
6. Confirm `REPORT_NUMMILotLocation[W]` are dead relics; only PLANT variants ship.

## 9. Test cases / parity checks
- D6 multi-window: 2 non-overlapping price windows, run `REPORT_INVOICESSummary @PDate=<window2>` → legacy
  duplicates/wrong price; rebuild = single row at window-2 price (reconcile to 856).
- Range semantics: order shipped month M, ordered M-1 → appears in month M.
- `@Supplier='ALL'` all suppliers; `''` empty (legacy quirk) → rebuild treats blank as ALL.
- Site isolation (post-D1): site-A invoice report excludes site-B ASNs.
