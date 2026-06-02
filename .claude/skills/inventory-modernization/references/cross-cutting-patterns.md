# Cross-Cutting Patterns

Patterns that recur across modules, discovered during analysis. Check here before
deep-diving a new module — most masters/transactions reuse these. Add to this file when
a new pattern shows up in ≥2 modules.

## P1 — Client-side duplicate-key guard (not a DB constraint)
Insert flows call `SELECT_<X>Info @key` first and only `INSERT` if `RecordCount = 0`.
The uniqueness rule lives in the **Delphi/DataModule client**, not the database.
- First seen: Supplier (`InsertSupplierInfo` → `SELECT_SupplierInfo` dup check).
- **Rebuild:** replace with a real **DB unique index + model validation**. Don't rely on
  re-implementing the check in app code only — make the constraint real.

## P2 — Timestamps stored as `yyyymmddHHMMSS` strings
Audit columns `VC_ADD` (insert) and `VC_LASTUPDATE` (update) are `varchar(16)` holding a
formatted string, computed in the proc via `CONVERT(...,112)+SUBSTRING(...,114,...)`.
- **Rebuild:** keep writing the string format during **parallel run** (legacy app reads
  the same rows); normalize to real `timestamp`/`datetimeoffset` at the **Postgres phase**.

## P3 — Name→ID resolution inside procs
UI shows a human label (e.g. logistics name) but the proc resolves it to a surrogate id
(`SELECT @LogisticsID = ... WHERE VC_LOGISTICS_NAME = @name`) before insert/update; a
blank label saves `NULL`.
- **Rebuild:** model the real FK association; the form posts the id (select field). No
  in-proc name lookup needed.

## P4 — Coded single-char enums
Flags stored as 1-char codes, mapped to labels in `SELECT_*` via `CASE` (e.g. output file
`T/E/B`→TEXT/EXCEL/BOTH; add point `S/A`→SHIPPED/ARRIVED).
- **Rebuild:** Rails `enum` (keep the stored char values for parallel-run compatibility).

## P5 — Delete unlinks children via trigger (soft-on-children)
`DELETE_<Master>Code` triggers null out FKs in child tables rather than cascading deletes
(e.g. `DELETE_SupplierCode` → `INV_PARTS_STOCK_MST.IN_SUPPLIER_ID = NULL`).
- **Rebuild:** `has_many ..., dependent: :nullify`. Capture each trigger's exact target
  columns in the module spec before removing it.

## P6 — DataModule is one shared ADO `Inv_StoredProc`
Every data method sets `ProcedureName := 'dbo.PROC;1'`, clears params, adds `@params`,
then `Open` (selects) or `ExecProc` (writes). `DataModule.pas` is therefore the index of
**which form calls which proc** — grep it when tracing a module's procs.

## P7 — Client-side search over a loaded grid
Search buttons often loop the already-loaded dataset in memory (e.g. `SearchGrid`) rather
than re-querying.
- **Rebuild:** replace with server-side query + pagination; trivially better.

## Multi-site lens (apply to every module)
The legacy app is single-site (identity in `INI [SITE]`). For each module ask:
**"what here is implicitly single-site?"** Common culprits:
- Tables with no site/tenant column but per-site data → need `site_id` scoping (Postgres phase).
- **Local filesystem paths** (e.g. `VC_BREAKDOWN_ORDER_DIRECTORY`) → replace with per-site
  configured targets (share/SFTP/object storage) or in-app delivery.
- Latent hooks already present, e.g. `BIT_SITE_NUMBER_IN_ORDER` — understand before reusing.
See [[project-multisite]] in memory.
