# Database Objects Inventory

SQL Server. **41 tables, 179 stored procedures, 24 triggers, 0 views, 0 functions.**
Source: `DB Schema/Create Inventory.sql` (UTF-16LE — use `scripts/sql.sh`).

> The procedures **are the business-logic spec.** When rebuilding a feature, read the
> proc body (`sql.sh proc PROC_NAME`) — the Delphi form mostly just binds to it.

## Tables (41)

### Transaction / operational
- `INV_OPEN_ORDER_INF`, `INV_OPEN_ORDER_INF_HIST` — open purchase orders (+ history)
- `INV_SHIPPING_INF`, `INV_PART_SHIPPING_INF` — shipments and shipped parts
- `INV_ASN_MST`, `INV_ASN_DETAIL_MST` — Advance Ship Notices (EDI 856)
- `INV_INVOICE_INF` — invoices (EDI 810)
- `INV_FORECAST_INF`, `INV_FORECAST_DETAIL_INF`, `INV_BREAKDOWN_FC_INF`, `INV_MANUAL_FORECAST` — demand forecasts (EDI 830 + manual)
- `INV_REJECT_INF` — receiving rejects / discrepancies
- `INV_STOCKTAKING_INF` — physical inventory counts
- `INV_PART_QTY_INF` — part quantity ledger

### Inventory / stock
- `INV_PARTS_STOCK_MST`, `INV_PARTS_STOCK_MST_HIST` — current stock (+ history)
- `INV_INV_MST` — inventory master
- `INV_ADD_POINT_INF` — add/reorder point info

### Assembly
- `INV_ASSY_RATIO`, `INV_ASSY_RATIO_MST`, `INV_BC_RATIO`, `INV_PART_RATIO` — build ratios (broadcast→part)
- `INV_ASSY_BUILD_HIST` — assembly build history
- `INV_ASSY_MONTHLY_PO`, `INV_ASSY_PO_CHARGED` — monthly assembly POs

### Master data
- `INV_SUPPLIER_MST` — suppliers
- `INV_SIZE_MST` — tire/wheel sizes
- `INV_PART_TYPE_MST`, `INV_PART_TYPE_INF` — part types
- `INV_RENBAN_GROUP_MST` — Renban (Toyota lot) groups
- `INV_LOGISTICS_MST` — logistics/carrier master
- `INV_MANIFEST_COST_MST` — manifest costs
- `INV_DOCK_INF` — dock info
- `INV_USERS` — application users/auth

### Calendar
- `INV_FIRST_PRODUCTION_DAY`, `INV_OVERTIME_HOLIDAY` — production calendar

### System / housekeeping
- `INV_PROGRAM_VERSION` — app version gate
- `Purge`, `Results`, `inv_temp`, `tempcount` — scratch/housekeeping (likely droppable)

## Stored procedures by functional domain

Naming convention encodes the operation: `SELECT_` / `INSERT_` / `UPDATE_` /
`DELETE_` / `INSERTUPDATE_` (upsert) / `REPORT_`.

### Ordering (~23)
`SELECT_OrderAtASSEMBLER`, `SELECT_OrderAtNUMMI`, `SELECT_OrderAtPLANT`,
`SELECT_OrderAtWQS`, `SELECT_OrderHistory`, `SELECT_OrderInTransit(List)`,
`SELECT_OrderNoRenban`, `SELECT_OrderNotOrdered`, `SELECT_OrderOpenOrder(List|Log)`,
`SELECT_OrderSplitFRS`, `INSERT_OpenOrder`, `UPDATE_ORDEROrderDate`,
`UPDATE_OrderPLANT`, `UPDATE_OrderQty`, `UPDATE_OrderRenban(Qty)`,
`UPDATE_OrderShipping`, `UPDATE_OrderTerminated`, `UPDATE_OrderWarehouse`,
`DELETE_OrderRenban`. (`AtASSEMBLER/NUMMI/PLANT/WQS` = order status by location stage.)

### Forecasting (~18)
`INSERTUPDATE_ForecastInfo`, `INSERTUPDATE_BreakdownForecastInfo`,
`INSERT_ForecastDetail`, `UPDATE_ForecastDetail`,
`DELETE_ForecastDetail`, `DELETE_ForecastInfoWeekDate(Part|PartOld)`,
`SELECT_ForecastDetail(BC|BCASN|TWPN)`, `SELECT_ForecastPartNumberWeek`,
`SELECT_ForecastSupplier`,
`REPORT_ForecastDetail`, `REPORT_ForecastSummary`, `REPORT_ForecastPartsSummary`,
`REPORT_LATEFRS`.

### Shipping (~13)
`INSERT_ShippingInfo`, `INSERT_ShippingDetail`, `INSERT_ShippingPartInfo`,
`SELECT_ShippingInfo`, `SELECT_ShippingDetail`, `SELECT_ShippingPartInfo`,
`SELECT_ShipMax`, `SELECT_ShipLastSeq`, `UPDATE_Shippingdetail`,
`REPORT_DailyShipping`, `REPORT_DailyShippingAssy`, `REPORT_DailyShippingRange`,
`REPORT_MonthlyShippingAssy`.

### ASN / EDI 856 (~11)
`INSERT_ASNInfo`, `INSERT_ASNDetail`, `SELECT_ASNList`, `SELECT_ASNItems`,
`SELECT_ASNMax`, `SELECT_ASNSeq`, `UPDATE_ASNItem`, `UPDATE_ASNStatus`,
`UPDATE_ASNUnsend`, `DELETE_ASNItem`, `DELETE_ASNList`, `REPORT_EDI856`.

### Invoicing / EDI 810 (~12)
`INSERTUPDATE_Invoice`, `SELECT_INVOICEList`, `SELECT_INVOICEItems`,
`UPDATE_INVItems`, `UPDATE_INVRecreate`, `UPDATE_INVUnsend`, `UPDATE_EINStatus`,
`REPORT_EDI810`, `REPORT_EDI810Recreate`, `REPORT_INVOICESSummary`,
`REPORT_MonthlyINVOICESSummary`, `REPORT_MonthlySupplierInvoices`.

### Parts stock / inventory (~13)
`INSERT_PartsStockInfo`, `SELECT_PartsStockInfo(Order)`, `SELECT_PartsStockLogistics`,
`SELECT_PartsStockRenban`, `UPDATE_PartsStockInfo(Count)`, `UPDATE_PartsStockRenban`,
`DELETE_PartsStockInfo`, `INSERT_INVInfo`, `SELECT_PartsDailyLinePull(Count)`,
`REPORT_LogicalInventory`.

### Assembly (~16)
`INSERT_AssyBuildHist`, `UPDATE_AssyBuildHistINV`, `SELECT_AssyBuildHist`,
`INSERT_AssyMonthlyPO`, `UPDATE_AssyMonthlyPO`, `DELETE_AssyMonthlyPO`,
`SELECT_AssyMonthlyPO(Display)`, `INSERT_AssyPOCharged`, `SELECT_AssyPOInfo`,
`INSERT_AssyRatioInfo`, `UPDATE_AssyRatioInfo`, `DELETE_AssyRatioInfo`,
`SELECT_AssyRatioInfo(Assy|Raw)`.

### Renban groups (~6)
`INSERT_RenbanGroup`, `SELECT_RenbanGroup`, `UPDATE_RenbanGroup(Count)`,
`DELETE_RenbanGroup`.

### Receiving (~9)
`INSERT_RecConfStatInfo`, `SELECT_RecConfStatInfo`,
`UPDATE_RecConfStatInfo`, `UPDATE_RecConfStatRenbanInfo`, `DELETE_RecConfStatInfo`,
`INSERT_RecProdRejInfo`, `SELECT_RecProdRejInfo`, `UPDATE_RecProdRejInfo`,
`DELETE_RecProdRejInfo`. (RecConfStat = receiving confirmation; RecProdRej = rejects.)

### Master data (~25)
- **Supplier:** `INSERT/SELECT/UPDATE/DELETE_SupplierInfo`, `SELECT_PartsSupplier`
- **Size:** `INSERT/SELECT/UPDATE/DELETE_SizeInfo`, `SELECT_SizeUsage`, `UPDATE_SizeUsage`
- **User:** `INSERT/SELECT/UPDATE/DELETE_UserInfo`
- **Logistics:** `INSERT/SELECT/UPDATE/DELETE_LogisticsInfo`, `REPORT_MonthlyLogisticsOrders`
- **ManifestCost:** `INSERT/SELECT/UPDATE/DELETE_ManifestCost`
- **Dependency lookups:** `SELECT_DependantKanbanNumber_PartNumber`,
  `SELECT_DependantPartNumber_PartType`, `SELECT_DependantPartNumber_Supplier`

### Production calendar (~10)
`INSERT/SELECT/DELETE_FirstProductionDay`, `INSERT/SELECT/DELETE_OvertimeHolidayInfo`,
`SELECT_OvertimeHoliday(Date|Week)`, `SELECT_CheckHoliday`, `SELECT_HolidayDate`,
`SELECT_UsageDay`, `SELECT_PartShipDays`, `REPORT_AvailableProductionDates`.

### Stocktaking (~4)
`INSERT/SELECT/UPDATE/DELETE_StockTakingInfo`.

### Lot-location + misc reports
`REPORT_NUMMILotLocation(W)`, `REPORT_PLANTLotLocation(W)`, `REPORT_EmptyContainer`,
`REPORT_PO`, `REPORT_DailySupplierOrders(Cost)`, `REPORT_MonthlySupplierOrders(Cost)`,
`REPORT_UnusedTirePartNumbers`, `REPORT_UnusedWheelPartNumbers`.

### Housekeeping
`DELETE_AutoPurge` (data-retention purge — see `[DATAPURGE]` in INI).

## Triggers (24) — data invariants to preserve

These keep **inventory quantities balanced** as transactions post. In the rebuild they
must become explicit app-layer logic (Rails model callbacks / service objects) — do not
lose them. Grouped by what they guard:

- **Stock-quantity sync on receiving:** `INSERT/UPDATE/DELETE_RecConfStatPartsStockMstQTY`
- **Stock on parts master:** `INSERT_PartsStockMST`, `INSERT_InvPartQtyInf`
- **Rejects adjust stock:** `INSERT/UPDATE/DELETE_RejectParts`
- **Stocktaking adjust stock:** `INSERT/UPDATE/DELETE_Stocktaking`
- **Shipping adjust stock:** `InsertPartShipping`, `UpdatePartShipping`,
  `DeletePartShipping`, `DeleteShipDate`
- **Forecast detail upkeep:** `UPDATE_ForecastDetailInf`, `DeleteForecastDetail`
- **Cascade/maintenance on master edits:** `UPDATE/DELETE_PartNumber`,
  `DELETE_SupplierCode`, `DELETE_SizeCode`, `DELETE_LogisticsCode`,
  `DELETE_RenbanGroupCode`, `UPDATE_AssyRatioMst`

> **Authoritative source = `DB Schema/Create Inventory.sql`** (these 24). `docs/triggers.sql`
> is an **obsolete pre-int-FK-refactor snapshot** — it keys on dropped string columns
> (`VC_SUPPLIER_CODE`, `VC_PART_NUMBER`, `VC_*_CODE`), is missing 5 live triggers, and must
> not be trusted. The live `*Code` DELETE triggers **null the int FK** on the child table
> (e.g. `DELETE_SupplierCode` → `INV_PARTS_STOCK_MST.IN_SUPPLIER_ID = NULL`). Full mapping:
> [`docs/analysis/cross-cutting/trigger-source-reconciliation.md`](../../../docs/analysis/cross-cutting/trigger-source-reconciliation.md).
