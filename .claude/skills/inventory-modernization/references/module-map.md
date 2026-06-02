# Module Map — Delphi forms → functional areas → target web modules

The live forms come from `InventorySystem.dpr` (~45 forms + non-visual units).
Grouped into functional areas below. Each area becomes a web module (Rails resource
group, or Python service for EDI/forecast). "Procs" points into
[database-objects.md](database-objects.md).

Legend: **M** master/CRUD · **T** transactional · **R** reporting · **S** support/dialog
· **X** EDI/integration · **A** admin

## 1. Ordering & Renban  *(T — core)*
| Delphi unit | Role |
|---|---|
| `Order.pas` (62 KB) | Main order entry/management screen (largest business form) |
| `OrderFormCreateF.pas` | Order sheet creation |
| `RenbanOrder.pas` | Group/Renban ordering |
| `RenbanGroupMaster.pas` | Renban group master (M) |
| `MonthlyPOMaster.pas` | Monthly PO master |
| `OrderQty.pas`, `OrderFormCreate.pas` | Helpers |
Procs: Ordering group + Renban group. **Highest-value / highest-risk module.**

## 2. Forecasting  *(T — core)*
| Delphi unit | Role |
|---|---|
| `ForecastBreakdownF.pas` (60 KB) | Forecast breakdown processing (2nd largest form) |
| `ForecastDetail.pas` | Forecast detail view |
| `ManualForecast.pas` | Manual forecast entry |
| `FRSBreakdown.pas` | FRS (firm release schedule) breakdown |
| `UploadBreakDown.pas`, `ForecastUploadBreakDown.pas` | Forecast file upload/breakdown |
| `ForecastCamexreport.pas` | CAMEX forecast report |
| `ForecastBreakDown.pas` | (referenced; legacy breakdown logic) |
Procs: Forecasting group. **Math-heavy → strong candidate for Python service.**

## 3. Receiving  *(T)*
| Delphi unit | Role |
|---|---|
| `RecConfStat.pas` (30 KB) | Receiving confirmation status |
| `RecReject.pas` | Receiving rejects / discrepancies |
Procs: Receiving group. Triggers here adjust stock — preserve carefully.

## 4. Shipping  *(T)*
| Delphi unit | Role |
|---|---|
| `Shipping.pas` | Shipping entry |
| `ManualShipping.pas` | Manual shipping |
| `ModifyShipping.pas` | Edit shipments |
| `DailyBuildTotal.pas` | Daily build totals |
Procs: Shipping group + shipping triggers.

## 5. EDI / Integration  *(X)*
| Delphi unit | Role |
|---|---|
| `EDI810Object.pas`, `Write810File.pas` | 810 invoice build/write |
| `EDI856Object.pas` | 856 ASN build |
| `ASNInvoice.pas` (36 KB), `ASNSelect.pas` | ASN + invoice generation |
| `InvoiceBreakdown.pas` | Invoice breakdown |
| `EDIUpload.pas` | EDI file upload/transfer (FTP) |
| `HotCallEntry.pas` | Hot-call ASN entry |
Procs: ASN + Invoicing groups; `REPORT_EDI810/856`. EDI files in `EDI/`, `EDIIn/`.
**Strong candidate for Python service** (X12 parse/generate, FTP).

## 6. Inventory & Stock  *(T/M)*
| Delphi unit | Role |
|---|---|
| `PartsStockMaster.pas` | Parts stock master |
| `InvMgmt.pas`, `InvMgmtQReport.pas` | Inventory management + its QuickReport |
| `Stocktaking.pas` | Physical inventory counts |
| `LogisticsBreakdown.pas` | Logistics breakdown |
Procs: Parts-stock group; stock-adjusting triggers.

## 7. Assembly  *(T/M)*
| Delphi unit | Role |
|---|---|
| `AssyRatioMaster.pas` (24 KB) | Assembly ratio master (broadcast→part) |
| `BCRatioMaster.pas` | Broadcast-code ratio master |
Procs: Assembly group.

## 8. Master data  *(M)*
| Delphi unit | Role |
|---|---|
| `SupplierMaster.pas` | Suppliers |
| `SizeMaster.pas` | Tire/wheel sizes |
| `LogisticsMaster.pas` | Logistics/carriers |
| `ManifestCostMaster.pas` | Manifest costs |
| `MasterMaint.pas` | Master-maintenance hub |
Procs: Master-data group. Simplest to port first (CRUD scaffolds).

## 9. Production calendar  *(M/S)*
| Delphi unit | Role |
|---|---|
| `FirstProductiionDay.pas` | First production day (sic — typo in filename) |
| `OvertimeHoliday.pas`, `HolidayOvertime.pas` | Overtime/holiday calendar |
| `ProductionDates.pas` | Production date selector (dialog) |
Procs: Production-calendar group. Feeds forecasting/ordering date math.

## 10. Reporting  *(R)*
| Delphi unit | Role |
|---|---|
| `Reports.pas` (41 KB) | Main reporting hub |
| `MonthlyReportSelect.pas` | Report date/param selector |
| `MonthlySupplerOrderReport.pas` | Monthly supplier orders |
| `MonthlySupplerInvoiceReport.pas` | Monthly supplier invoices |
| `MonthlyLogiticsOrderReport.pas` | Monthly logistics orders |
Procs: all `REPORT_*` (29). QuickReport → HTML/PDF in the rebuild.

## 11. Admin / system / shell  *(A/S)*
| Delphi unit | Role |
|---|---|
| `MainMenu.pas` (138 KB) | Main window + menu + orchestration (the shell) |
| `DataModule.pas` (267 KB) | Central data-access layer (all 3 ADO connections) |
| `Logon.pas`, `UserAdmin.pas`, `UserInfo.pas` | Auth + user admin |
| `ConfirmPassword.pas`, `NewPassword.pas` | Password dialogs |
| `Configuration.pas`, `SiteInfo.pas` | App/site configuration (INI-backed) |
| `About.pas`, `VersionInfo.pas` | About / version gate |
| `DirectorySelect.pas`, `SelectDateRange.pas` | Generic dialogs |
Procs: User group; `SELECT_ProgramVersion`. `DataModule.pas` is the map of every
proc call — invaluable when tracing which form uses which proc.

## Suggested rebuild order (see migration-strategy.md)
Masters (8) → Production calendar (9) → Inventory/Stock (6) → Receiving (3) →
Shipping (4) → Ordering (1) → Forecasting (2) → EDI (5) → Assembly (7) →
Reporting (10), with Admin/auth (11) bootstrapped first.
