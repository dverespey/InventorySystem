# Domain Glossary

Automotive parts supply-chain terms used throughout the system (Toyota / TEMA world).
The user is the domain expert — this is a reference for the assistant.

## Trading partners / plants
- **TEMA** — Toyota Engineering & Manufacturing North America (the customer org).
- **CAMEX / NUMMI / TMMTX / WQS / PLANT / ASSEMBLER** — assembly plants / order stages.
  In ordering, `SELECT_OrderAt{ASSEMBLER,NUMMI,PLANT,WQS}` track an order's position
  as it moves through the supply pipeline (supplier → in transit → plant → line).
- **Supplier** — the tire/wheel/parts vendor. The app runs from the supplier's side
  (this is a supplier-side inventory & EDI system). `SupplierCode` set in INI `[SITE]`.

## Parts / build
- **Renban** — Toyota sequencing/lot identifier; groups parts into ordered lots.
  Drives `INV_RENBAN_GROUP_MST` and the Renban ordering screens.
- **Broadcast / Broadcasting Code (BC)** — the build/sequence code for a specific truck
  on the assembly line; expands via ratio tables into the actual parts needed.
  See `INV_BC_RATIO`, `INV_ASSY_RATIO*`, `AssyRatioMaster`.
- **Assy / Assembly ratio** — mapping from a broadcast code / model to component parts
  and quantities (e.g., one truck → 4 tires + 1 spare of size X).
- **Kanban number** — pull signal / part identifier (`SELECT_DependantKanbanNumber_*`).
- **Tire / Wheel / Coil** — the physical goods. Reports distinguish tire vs wheel
  part numbers (`REPORT_UnusedTirePartNumbers`, `...WheelPartNumbers`).
- **Lot location** — where stock physically sits (`REPORT_*LotLocation`).

## Demand / orders
- **Forecast** — projected demand from the customer (EDI 830). Stored in
  `INV_FORECAST_INF` / `_DETAIL_INF`; broken down by part/week.
- **FRS — Firm Release Schedule** — the firm (committed) portion of the forecast vs the
  planning portion. `FRSBreakdown`, `SELECT_OrderSplitFRS`, `REPORT_LATEFRS`.
- **Open order** — a placed-but-not-yet-fulfilled purchase order
  (`INV_OPEN_ORDER_INF`).
- **PO — Purchase Order** — `REPORT_PO`, `MonthlyPOMaster`.
- **Hot call** — an urgent/expedited order or ASN (`HotCallEntry`).
- **Usage / Line pull / Daily line pull** — consumption of parts at the line; basis for
  reorder math (`SELECT_PartsDailyLinePull`, `SELECT_UsageDay`, `SELECT_SizeUsage`).

## Logistics / shipping
- **ASN — Advance Ship Notice (EDI 856)** — notifies the customer of an inbound shipment.
  `INV_ASN_MST` / `_DETAIL_MST`.
- **Manifest** — shipment manifest; `INV_MANIFEST_COST_MST` holds freight/manifest cost.
- **Container / Empty container** — returnable packaging tracking
  (`REPORT_EmptyContainer`).
- **Logistics master** — carriers / lanes (`INV_LOGISTICS_MST`).
- **Dock** — receiving/shipping dock (`INV_DOCK_INF`).

## Billing / EDI
- **EDI X12** — the electronic-data-interchange standard. Transactions used:
  - **830** — Planning Schedule / forecast (inbound).
  - **856** — Advance Ship Notice / ASN (outbound).
  - **810** — Invoice (outbound).
  - **820** — Remittance advice (seen in `docs/481CAMEX820I001_RemitAdvice.TXT`).
  - **824** — Application advice / acknowledgement (`docs/824MessageOut.txt`).
- **EIN** — invoice EDI status flag (`UPDATE_EINStatus`).
- **DUNS** — partner identifier used in EDI envelopes (see `SiteInfo.pas`).
- Files move over **FTP**; see `EDI/`, `EDIIn/`, and `docs/*FTP*`.

## Calendar
- **First production day** — the first build day used as an anchor for date math
  (`INV_FIRST_PRODUCTION_DAY`, INI `UseFirstProductionDay`).
- **Overtime / Holiday** — calendar exceptions affecting build/ship day calculations
  (`INV_OVERTIME_HOLIDAY`).

## Reference documents (in `docs/`)
- `Toyota_EDI_Master_Implementation_Manual_Ver_1 7.pdf` — the authoritative EDI spec.
- `810 Specifications.xls`, `856 Specifications.xls` — transaction layouts.
- `Going_Live_with_EDI_at_TAI_v3.doc`, `TEMA EDI Billing.doc` — process docs.
- `Generic File Format.doc`, `Logistics File Format.doc` — inbound file formats.
