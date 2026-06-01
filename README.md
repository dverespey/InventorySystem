# Inventory Control System

A Windows desktop application for **automotive-parts inventory and supply-chain
management**, built in **Delphi 7 (Object Pascal / VCL)**. It manages
forecasting, ordering, receiving, shipping, stocktaking, and **EDI billing** for
a parts supplier feeding Toyota-family assembly plants.

> Originally created December 2002 by David Verespey (Failproof Manufacturing
> Systems), later maintained under the "TAI" brand.

---

## Domain

The system supports **Toyota / TEMA automotive logistics** as a Tier-1/Tier-2
supplier. You will see this terminology throughout the code and documents:

| Term | Meaning |
|------|---------|
| **CAMEX / NUMMI / TMMTX** | Toyota-related plants / programs the supplier ships to |
| **Renban** | Toyota sequencing/lot identifier used for ordering |
| **Broadcasting Code** | Build/sequence code for a truck on the assembly line |
| **ASN** | Advance Ship Notice (EDI 856) |
| **810 / 856 / 830** | EDI X12 transactions: invoice / ASN / planning forecast |
| **FRS / Forecast Breakdown** | Demand forecast processing |

Primary parts tracked include tires and coils.

---

## Tech stack

- **Language / IDE:** Delphi 7, VCL forms (`.pas` units + `.dfm` form layouts)
- **Database:** Microsoft SQL Server, accessed via **ADO** (`SQLOLEDB` provider).
  Three catalogs are used:
  - `Inventory` — core inventory, parts, orders, forecasts
  - `Activity` — activity/logging
  - `VehicleOrder` — ALC / vehicle order data
- **Reporting:** QuickReport
- **Integration:** EDI X12 (810/856/830) exchanged over FTP

---

## Repository layout

```
InventorySystem.dpr        Program entry point (form/unit manifest)
DataModule.pas             Central data-access layer (ADO connections + queries)
MainMenu.pas               Main UI shell and orchestration
SiteInfo.pas               Per-site configuration object (DUNS, codes, EDI mode)

<feature>.pas / .dfm       One unit+form per business function, e.g.:
  Order, Shipping, RecConfStat (receiving), Stocktaking,
  ForecastBreakdownF, SupplierMaster, RenbanOrder, ASNInvoice, EDIUpload

EDI810Object.pas           EDI 810 (invoice) builder
EDI856Object.pas           EDI 856 (ASN) builder
Write810File.pas           EDI 810 file writer

DB Schema/Create Inventory.sql   Full schema (~40 INV_* tables)
docs/triggers.sql                Database triggers
docs/                            Specs, EDI manuals, style guide, status docs

EDI/, EDIIn/               Sample / working EDI files and FTP scripts
Reports/, Templates/       Report layouts and templates
*.INI                      Per-site runtime configuration (see below)
```

### Key tables (schema)
`INV_PARTS_STOCK_MST`, `INV_OPEN_ORDER_INF`, `INV_FORECAST_INF`,
`INV_SHIPPING_INF`, `INV_ASN_MST`, `INV_INVOICE_INF`, `INV_SUPPLIER_MST`,
`INV_RENBAN_GROUP_MST`, `INV_USERS`, and ~30 more (`INV_*` prefix).

---

## Configuration

Runtime behavior is driven by INI files (e.g. `InventorySystem.INI`, with
site variants `InventorySystemHERO.INI`, `InventorySystemCAMEX.INI`,
`InventorySystemdev.INI`). Notable sections:

- `[DATABASE]` — ADO connection strings for the three catalogs
- `[SITE]` — supplier code, plant name, EDI feature flags
- `[DIRECTORIES]` — forecast/EDI/report input & output folders
- `[INIT]` — forecast/fill-day parameters, order-creation behavior
- `[DATAPURGE]` — automatic data retention/purge settings

> ⚠️ **Secrets:** the INI files contain SQL Server connection strings **with
> passwords**. They are now git-ignored. Keep credentials out of source control
> and rotate any that were previously committed. Maintain a sanitized
> `InventorySystem.INI.example` for onboarding if needed.

---

## Building

1. Open `InventorySystem.dpr` in **Delphi 7**.
2. Ensure required third-party VCL components (QuickReport, ADO data-aware
   controls, custom `NUMMIBmDateEdit`) are installed.
3. Build/compile to produce `InventorySystem.exe`.
4. Provide a valid `InventorySystem.INI` pointing at an accessible SQL Server
   with the `Inventory`, `Activity`, and `VehicleOrder` databases (create them
   from `DB Schema/Create Inventory.sql` + `docs/triggers.sql`).

Compiler options are stored in `InventorySystem.dof` / `InventorySystem.cfg`.

---

## Documentation

The `docs/` folder contains substantial reference material:

- `Toyota_EDI_Master_Implementation_Manual_Ver_1 7.pdf`
- `810 Specifications.xls`, `856 Specifications.xls`
- `Going_Live_with_EDI_at_TAI_v3.doc`
- `Status of Developement.doc`, `Programming Style.doc`
- Various file-format specs (forecast, logistics, generic)

---

## Notes for maintainers

- `DataModule.pas` (~267 KB) and `MainMenu.pas` (~138 KB) are large "god"
  units — most data access and orchestration live here. Change with care.
- The tree historically contained dead/duplicate units (`*old.pas`, `*1.pas`,
  `DataModule1.pas`, `Copy of CAMEX System/`). These are **not** referenced by
  `InventorySystem.dpr` and can be removed once verified.
- Build outputs (`.dcu`, `.exe`), IDE backups (`*.~*`), and OS junk
  (`.DS_Store`, `Thumbs.db`) are now git-ignored; previously-tracked copies may
  need `git rm --cached`.
