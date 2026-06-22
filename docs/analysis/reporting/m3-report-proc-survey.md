# M3 Report-Proc Data-Surface Survey

**Area:** Reporting (M3 data layer)  **Analyst:** Claude / 2026-06-21
**Goal:** for each report family the M3 build must reproduce, capture the *data surface* — lineage, projection,
aggregation, params, and the data hazards that will silently break "numbers match legacy."

**Sources:** live `Inventory` on `mssql-spike` (bodies via `OBJECT_DEFINITION`) + `/tmp/inv_utf8.sql`.
Builds on the full catalog in [`reporting.md`](reporting.md) (do not re-read; this is the *numbers* lens).
Daily Shipping (the priority path) is specified separately in
[`daily-shipping-data-analysis.md`](daily-shipping-data-analysis.md).

> Line cites are `/tmp/inv_utf8.sql` (live dump). `reporting.md` cites the superseded schema by line.

---

## A. Read-only audit (all 29 `REPORT_*` procs)

Keyword scan of `OBJECT_DEFINITION` for `INSERT/UPDATE/DELETE` across `sys.procedures WHERE name LIKE
'REPORT[_]%'` (live Inventory):

- **27 of 29 are pure read-only.**
- **2 MUTATE ON READ** (hazard): `REPORT_EDI810`, `REPORT_EDI856` — see §C.

---

## B. Data surface by report family

Hazard legend: **ORPHAN** = inner-join silently drops rows; **D6** = window-blind manifest-cost join (wrong
price/fan-out); **WINDOW** = depends on `GetDate()`/relative date (non-reproducible from a fixed snapshot);
**SUM-NULL** = NULL nullifies an additive expression; **SELECT\*** = `select *` exposes ledger `IN_QTY`
(the M2 duplicate-IN_QTY trap) / column-order brittleness; **FANOUT** = join multiplies aggregated rows;
**MUTATE** = writes on read.

| Proc (line) | Reads (lineage) | Projection (key cols) | Aggregation | Params | Hazards |
|---|---|---|---|---|---|
| **REPORT_DailyShipping** (2875) | SHIPPING_INF s ⋈date PART_SHIPPING_INF p ⋈pn PARTS_STOCK_MST m | Start, End, Vehicle Count(=s.IN_QTY), Part Number, Desc, PQty=`SUM(p.IN_QTY)` | GROUP BY type,pn,seq,IN_QTY,name | `@Pdate v8` | ORPHAN (same inner join as Range); seq cols add to grain |
| **REPORT_DailyShippingRange** (2838) | same as above | Vehicle Count=`SUM(s.IN_QTY)`, Part Number, Desc, PQty=`SUM(p.IN_QTY)` | GROUP BY type,pn,name | `@Begin/@EndPdate v8` | ORPHAN (641/886 dropped), FANOUT (benign:1 s/date). See dedicated doc. |
| **REPORT_DailyShippingAssy** (3576) | ASN_MST s ⋈IN_ASN_ID ASN_DETAIL_MST d | Start, End, Vehicle Count, Part Number=d.VC_ASSY_PART_NUMBER, PQty=`SUM(d.IN_QTY)` | GROUP BY assy_pn,seq,IN_QTY | `@Pdate v8` | **clean** (PK⋈FK, no master join → no orphan). The sound shipping path. |
| **REPORT_MonthlyShippingAssy** (3539) | ASN_MST s ⋈IN_ASN_ID ASN_DETAIL_MST d | Vehicle Count, Part Number=d.VC_ASSY_PART_NUMBER, PQty=`SUM(d.IN_QTY)`, PDate | GROUP BY assy_pn,s.IN_QTY,d.IN_QTY,date | `@Pdate v6` | clean; `substring(date,1,6)=@PDate` month filter; d.IN_QTY in GROUP BY makes the SUM a near no-op per (date,qty) |
| **REPORT_AvailableProductionDates** (726) | SHIPPING_INF *or* ASN_MST (branch) | DISTINCT VC_PRODUCTION_DATE (or `substring(,1,6)` for month) 'PDate' | DISTINCT | `@Line v50,@INVOICE int,@ASN int,@Month=0` | date-picker feed for Daily Shipping; SHIPPING branch filters `VC_Line_Name=@Line` |
| **REPORT_LogicalInventory** (6808) | PARTS_STOCK_MST p LEFT⋈ supplier/logistics/renban/parttype/size | `select *` (all p cols + lookups) | none (row-per-part) | `@PartNo='ALL'` | **SELECT\*** → exposes `p.IN_QTY` (ledger column; M2 trap) + column-order brittle; LEFT joins safe (no drop). MainMenu:934,2189 |
| **REPORT_LATEFRS** (6854) | OPEN_ORDER_INF | `select *` | none | `@FRSDate v8` | SELECT\*; predicate = `vc_frs_date<@FRSDate` AND 5 downstream status cols all `=''` (open orders past FRS) |
| **REPORT_EmptyContainer** (6873) | OPEN_ORDER_INF o ⋈pn PARTS_STOCK_MST p ⋈ supplier ⋈ logistics ⋈ parttype | supplier, logistics, renban, plant-parking, kanban, part-type, formatted dates | GROUP BY (4-way ALL/ALL branch matrix) | `@Start,@End v8,@Supplier='ALL',@Logistics='ALL'` | ORPHAN-latent (inner ⋈ master, but 0/4238 dropped today); `VC_STATUS_EMPTY_TRAILER between @Start..@End AND <>''`; dd/mm/yy substring reformat |
| **REPORT_PO** (1287) | INV_ASSY_MONTHLY_PO | `select *` | none | `@Begin,@End v8` | SELECT\*; `VC_PO_MONTH_START>=@Begin AND VC_PO_MONTH_END<=@End` |
| **REPORT_ForecastSummary** (1275) | INV_FORECAST_INF | `select *` | none | none | **WINDOW** (`IN_WEEK_NUMBER >= DATEPART(week,GetDate())`) — result depends on today; SELECT\* |
| **REPORT_ForecastDetail** (3859) | INV_FORECAST_DETAIL_INF | `select *` ordered by assy_pn,eff_month,broadcast | none | none | SELECT\*; whole detail table |
| **REPORT_ForecastPartsSummary** (4034) | BREAKDOWN_FC_INF b ⋈pn PARTS_STOCK_MST p | week#, pn, name, `qty=IN_QTY1+..+IN_QTY7` | none (per breakdown row) | none | ORPHAN-latent (inner ⋈ master, 0/959 dropped today), **SUM-NULL** (any of IN_QTY1..7 NULL → qty NULL; 0 NULL rows today), **WINDOW** (`b.VC_WEEK_DATE >= convert(char8,GetDate(),112)`) |
| **REPORT_INVOICESSummary** (3785) | ASN_MST a ⋈ ASN_DETAIL_MST d ⋈ MANIFEST_COST_MST m ⋈ INV_MST i | Manifest, PartNumber, UnitPrice=m.MO_PRICE, ShipQty=d.IN_QTY, PickUpDate | none | `@PDate v13` | **D6** (no `VC_START/END_MANIFEST` window → wrong price / row mult); `VC_PRODUCTION_DATE=@PDate` |
| **REPORT_MonthlyINVOICESSummary** (3833) | same as above | same | none | `@PDate v6` | **D6** (same); `substring(date,1,6)=@PDate` |
| **REPORT_MonthlySupplierInvoices** (4815) | INV_INVOICE_INF (prices from invoice rows) | invoice ledger cols | per invoice | `@Start,@End v8,@Supplier=''` | NOT D6 (reads stored MONEY cols); `@Supplier=''` default → else-branch empty unless dialog coerces to 'ALL' (reporting.md §4.2) |
| **REPORT_MonthlySupplierOrders** (6763) | OPEN_ORDER_INF (+ supplier) | order ledger | GROUP/range | `@Start,@End v8,@Supplier='ALL'` | ranges on **VC_STATUS_SUPPLIER_SHIPPING** (ship date) not order date — D12 confirmed bug |
| **REPORT_MonthlySupplierOrdersCost** (6706) | OPEN_ORDER_INF ⋈ PARTS_STOCK_MST (cost) | + MO_PART_COST, Total=cost*qty | GROUP/range | `@Start,@End,@Supplier='ALL'` | ship-date range (D12); cost from part-master (not windowed) |
| **REPORT_MonthlyLogisticsOrders** (6659) | OPEN_ORDER ⋈ supplier ⋈ logistics | logistics roll-up | GROUP by logistics/ship/renban | `@Start,@End,@Logistics='ALL'` | ship-date range (D12) |
| **REPORT_DailySupplierOrders** (7109) | OPEN_ORDER_INF | order list | none | `@StartDate,@Supplier='ALL'` | `VC_ORDER_DATE=@StartDate` (correct order-date filter) |
| **REPORT_DailySupplierOrdersCost** (7054) | OPEN_ORDER ⋈ PARTS_STOCK_MST | + cost + Total | none | `@StartDate,@Supplier='ALL'` | adds MO_PART_COST*IN_QTY |
| **REPORT_PLANTLotLocation / …W** (6625/6587) | (lot-location; W=warehouse variant) | lot/location grid | — | none | body not re-decoded (low M3 priority); PLANT variants ship |
| **REPORT_NUMMILotLocation / …W** (1026/988) | — | — | — | — | **DEPRECATED** (NUMMI decommissioned, D9) — out of scope |
| **REPORT_UnusedTirePartNumbers** (4216) | tire parts not in forecast-detail | unused tire pns | — | none | low priority |
| **REPORT_UnusedWheelPartNumbers** (4201) | wheel parts vs forecast-detail | unused wheel pns | — | none | **BUG (D11)**: filters against the TIRE code column for wheel parts |
| **REPORT_EDI810** (3734) | ASN_MST ⋈ ASN_DETAIL ⋈ MANIFEST_COST_MST (+INV_MST in EIN branch) | Manifest, PartNumber, UnitPrice, ShipQty, PickUpDate, ASNid | none | `@EIN int=0` | **D6** (no manifest window) + **MUTATE** (EIN≠0 branch) — see §C |
| **REPORT_EDI810Recreate** (3706) | ASN/detail/manifest | recreate-810 extract | none | none | D6; not on Reporting menu |
| **REPORT_EDI856** (3629) | ASN_MST ⋈ ASN_DETAIL ⋈ MANIFEST_COST_MST ⋈ FORECAST_DETAIL_INF | Manifest, PartNumber, UnitPrice, ShipQty, PickUpDate, Kanban, SiteEIN, StartSeq, LineName | GROUP BY | `@EIN int=0` | EIN=0 branch **WINDOW-AWARE** (correct, D6 contrast) + EIN≠0 branch **MUTATE** — see §C |

---

## C. Mutate-on-read hazards (CRITICAL for M3)

Two "REPORT_" procs are **dual-purpose**: a report SELECT *and* a status-commit UPDATE, branched on `@EIN`.
The default `@EIN=0` branch is a pure SELECT; the `@EIN <> 0` branch runs an UPDATE *after* the SELECT.

**`REPORT_EDI810`** (`/tmp/inv_utf8.sql:3775`):
```sql
UPDATE INV_INV_MST SET VC_INV_STATUS = 'S' WHERE IN_INV_EIN = @EIN
```
→ marks the invoice batch "submitted" as a side effect of running the EIN report.

**`REPORT_EDI856`** (`/tmp/inv_utf8.sql:3695`):
```sql
UPDATE INV_ASN_MST SET VC_ASN_STATUS = 'S' WHERE IN_ASN_EIN = @EIN
```
→ flips ASN status to "submitted." (Note: the SELECT in this branch comments out the manifest window
and hardcodes `IN_ASN_EIN = 6440` — a stale literal; the UPDATE uses `@EIN`.)

**Why this matters for the rebuild:**
- These are NOT pure reports. Re-running with `@EIN ≠ 0` **changes EDI batch state** (idempotent only in
  that re-running re-sets 'S' to 'S', but it transitions a not-yet-submitted batch). A "preview/report"
  button wired to the EIN branch would silently commit the batch.
- A Named-Query port MUST split these: a read-only report query (the `@EIN=0` SELECT) vs an explicit
  "submit" action (the UPDATE). Do **not** fold the UPDATE into a report Named Query.
- These two are already covered as the **D6 EDI path** (`edi/asn-invoice.md`); listed here so M3's report
  inventory does not treat them as plain reports.

All 27 other `REPORT_*` procs (incl. Daily/Monthly Shipping, all Order/Invoice/Forecast/Inventory reports)
are **read-only** — no self-flip, no mutation.

---

## D. Cross-cutting data hazards the M3 build must handle

1. **ORPHAN inner-join to `INV_PARTS_STOCK_MST`** — `REPORT_DailyShipping[Range]`, `REPORT_EmptyContainer`,
   `REPORT_ForecastPartsSummary`. All three inner-join the master, so any source part not registered there
   is silently dropped. **Severity differs by table (verified live):**
   - `REPORT_DailyShipping[Range]`: **actively dropping** — 641/886 (72%) `INV_PART_SHIPPING_INF` rows
     orphaned (tire/wheel parts are not in the master). This is the M3 failing path.
   - `REPORT_EmptyContainer`: latent only — all 4238 `INV_OPEN_ORDER_INF` rows match the master today
     (0 dropped).
   - `REPORT_ForecastPartsSummary`: latent only — all 959 `INV_BREAKDOWN_FC_INF` rows match (0 dropped).

   The Assy variants (`*ShippingAssy`) avoid the hazard entirely by joining ASN PK→FK only — the sound
   pattern. Faithful parity = keep the inner join (and its drop where it bites); corrected behavior = a
   David decision.
2. **D6 window-blind pricing** — `REPORT_INVOICESSummary`, `REPORT_MonthlyINVOICESSummary`, `REPORT_EDI810`,
   `REPORT_EDI810Recreate`. Resolved D11: rebuild uses the window-aware manifest-cost lookup (the 856
   `VC_START_MANIFEST <= date <= VC_END_MANIFEST` predicate). `REPORT_EDI856`(EIN=0) is the correct contrast.
3. **WINDOW / `GetDate()` dependence** — `REPORT_ForecastSummary` (`IN_WEEK_NUMBER >= DATEPART(week,now)`),
   `REPORT_ForecastPartsSummary` (`VC_WEEK_DATE >= today,112`). Their output is **not reproducible from a
   fixed snapshot** without pinning "today" — parity tests must fix the clock or parameterize the date.
4. **SELECT \*** — `REPORT_LogicalInventory`, `REPORT_LATEFRS`, `REPORT_PO`, `REPORT_ForecastSummary`,
   `REPORT_ForecastDetail`. Column order + presence depend on table DDL; `LogicalInventory` exposes the
   ledger `IN_QTY` (the M2 duplicate-IN_QTY trap — `INV_PARTS_STOCK_MST.IN_QTY` is the running balance
   maintained by the InsertPartShipping/DeletePartShipping triggers, not an independent count). A Named
   Query must pin an explicit column list and decide which `IN_QTY` it means.
5. **SUM-NULL** — `REPORT_ForecastPartsSummary` `IN_QTY1+..+IN_QTY7`: any NULL operand → whole `qty` NULL
   (no implicit zero). Use `ISNULL`/`COALESCE` per-operand *only if* matching observed legacy values
   (legacy does NOT coalesce → NULL today; faithful = leave NULL).
6. **Ship-date vs order-date range** (D12) — Monthly order/cost/logistics reports range on
   `VC_STATUS_SUPPLIER_SHIPPING`; resolved bug → rebuild ranges on `VC_ORDER_DATE`. Daily order reports
   already correct.
7. **Positional vs named param binding** (reporting.md §4.2) — ADO binds positionally; some callers add
   differently-named params (e.g. the InvMgmt feed adds `@InvMgmtReport`/`@SupCode` that the live
   `SELECT_PartsStockInfo` doesn't even declare — only its single `@PartNum` binds). A name-bound Named
   Query must reconcile declared param order/names.

---

## E. Notes / drift flagged

- **`SELECT_PartsStockInfo`** (InvMgmtQReport feed, called `DataModule.pas:4460-4476`): live proc declares
  **only `@PartNum varchar(12)=''`** — the caller's extra `@InvMgmtReport='N'` / `@SupCode=''` params bind
  positionally and are effectively ignored (only the first maps; the order means `@InvMgmtReport`'s value
  'N' lands in `@PartNum`). It references `IN_QTY`, is **read-only**, and does **not** emit the
  `'Last Scrap Count'` field the caller reads (`DataModule.pas:4480`) — possible drift in the
  InvMgmt/auto-scrap path (outside the report families; flag for the InvMgmt analysis, not M3 reports).
- `REPORT_PLANTLotLocation[W]` bodies not re-decoded here (lower M3 priority; PLANT variants are the live
  ones, NUMMI twins are deprecated D9).
