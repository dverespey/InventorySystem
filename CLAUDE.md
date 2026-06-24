# CLAUDE.md

Guidance for AI assistants (and new developers) working in this repository.

> ## ⮕ The Ignition rebuild now lives in [`dverespey/inventory`](https://github.com/dverespey/inventory)
> As of 2026-06-24 the **Ignition (Perspective + gateway Jython) rebuild** — the
> gateway project-as-code, the e2e/parity harness, the view generators, and the
> canonical DDL (`db/CreateInventory.sql`) — was **split out into its own private repo
> `dverespey/inventory`**. Do all Ignition build/test/spike work there (`cd ~/Documents/GitHub/inventory`).
> **THIS repo is now the Delphi legacy source + the reverse-engineering analysis** — the
> source-of-truth bridge (`docs/analysis/`, the D-log, the `file:line` cites) that the
> `inventory` repo cross-links back to. Keep that bridge accurate; don't re-add Ignition
> code here.

## What this is

A **Delphi 7 (Object Pascal / VCL)** Windows desktop application for automotive
parts **inventory and EDI supply-chain management**, used by a supplier shipping
to Toyota-family assembly plants (CAMEX / NUMMI / TMMTX). Backend is **Microsoft
SQL Server via ADO**. See `README.md` for the full overview.

There is **no automated build or test in this repo** — it is compiled in the
Delphi 7 IDE from `InventorySystem.dpr`. You cannot build, run, or test it from
the command line here. Do not invent build/test commands.

## Architecture map

- **`InventorySystem.dpr`** — entry point. The authoritative list of *live*
  units/forms. If a `.pas` file is **not** listed here, it is dead/legacy code
  (e.g. `DataModule1.pas`, `Orderold.pas`, `Order1.pas`, `*old.pas`,
  `PartsStockMasterNew.pas`).
- **`DataModule.pas`** (~267 KB) — the data layer. Holds the three
  `TADOConnection`s (`Inv_Connection`, `Act_Connection`, `ALC_Connection`) and
  the bulk of the SQL. Most data changes start here.
- **`MainMenu.pas`** (~138 KB) — main window and feature orchestration.
- **One unit + one `.dfm` per business function** — e.g. `Order`, `Shipping`,
  `RecConfStat` (receiving), `Stocktaking`, `ForecastBreakdownF`,
  `SupplierMaster`, `RenbanOrder`, `ASNInvoice`, `EDIUpload`.
- **EDI:** `EDI810Object.pas` (invoice), `EDI856Object.pas` (ASN),
  `Write810File.pas`. Specs live in `docs/`.
- **Schema:** `DB Schema/CreateInventory.sql` — the **authoritative live server dump
  (2026-06-12)**: 42 `INV_*` tables, 182 procs, 25 triggers. (The older
  `Create Inventory.superseded-2026-06-01.sql` is the prior snapshot that pre-2026-06-16
  analysis specs cite by line; `docs/triggers.sql` is obsolete — do not port from it.)

## Conventions in this codebase

- A VCL form is a pair: `Foo.pas` (logic) + `Foo.dfm` (layout). Edit both
  consistently; never hand-edit `.dcu` (compiled) files.
- Form types/vars use a `T..._Form` / `..._Form` naming pattern
  (e.g. `TMainMenu_Form`, `MainMenu_Form`).
- Database tables use the `INV_*` prefix; columns are SQL-Server style.
- Configuration is read from INI files at runtime (see `[DATABASE]`, `[SITE]`,
  `[DIRECTORIES]`, `[INIT]`, `[DATAPURGE]` sections). Match the existing
  property-backed pattern (see `SiteInfo.pas`) when adding settings.
- Match the surrounding Pascal style — the repo predates modern Delphi idioms;
  keep edits consistent with the existing file rather than modernizing.

## Guardrails

- **Secrets:** INI files contain SQL connection strings *with passwords* and are
  now git-ignored. Never commit credentials or echo them into new files. If you
  add config docs, use placeholder values and an `*.example` file.
- **Generated/backup files:** `.dcu`, `.exe`, `.ddp`, `*.~*` backups,
  `.DS_Store`, and `Thumbs.db` are git-ignored. Do not edit or re-add them.
- **Dead code:** before "fixing" something, confirm the unit is referenced in
  `InventorySystem.dpr`. There are several stale duplicates that look real but
  are not compiled into the product.
- **Encoding/line endings:** these are legacy Windows files (CRLF, possibly with
  IDE markers). Make minimal, targeted edits; avoid reformatting whole files.

## When asked to change behavior

1. Find the live unit via `InventorySystem.dpr`.
2. Trace data access into `DataModule.pas`.
3. Update the `.pas` and its `.dfm` together if the UI is involved.
4. Check `DB Schema/CreateInventory.sql` (the authoritative live dump) if the change
   touches the database.
5. State clearly that the change must be compiled/tested in the Delphi 7 IDE —
   you cannot verify it from this environment.
