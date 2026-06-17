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

## P8 — Recursive retry-on-error in DataModule methods
Every `DataModule.pas` data method wraps its ADO call in the same error harness: on an
exception, if a shared `fErrorCount < 3` it **recursively re-calls itself**, with a `finally`
that does `Inv_StoredProc.Close; fErrorCount := 0`; on a hard failure it `ShowMessage`s,
`LogActLog('ERROR',…)`, and raises a distinct `EDatabaseError`. Pervasive — `fErrorCount`
appears ~240× across the unit.
- First seen: Supplier / Logistics (all four CRUD methods each).
- **Rebuild:** this is transport-level retry. Replace with a single connection/retry policy
  (e.g. a `tiny_tds` wrapper or `retriable`), not per-method recursion. The `LogActLog` audit
  trail (`GET/INSERT/UPDATE/DELETE` + ERROR rows) is a real behavior to preserve as app logging.

## P9 — Shared generic `RecordID` property as the row key
Forms don't keep their own selected-id; they write the grid's identity column into one
**shared** `DataModule.RecordID` property (set in `HoldDetails(True)` from a hidden grid field),
then Update/Delete key off it. The same property is reused by Shipping/Invoice/ASN/etc.
- First seen: Supplier, Logistics.
- **Hazard:** if no row was selected first, `RecordID` is `0` or a **stale value from another
  screen** — no guard exists. A cross-module write-to-wrong-row risk.
- **Rebuild:** the id belongs to the request/resource (`params[:id]`), never shared mutable
  state. This whole pattern disappears with RESTful routing.

## P10 — Positional `INSERT` with no column list
Some insert procs do `INSERT INTO <table> VALUES(...)` with **no explicit column list**, relying on
physical column order (and sometimes writing the same value to two columns by position — e.g.
`INSERT_ManifestCost` writes `@AddDate` to both `VC_ADD` and `VC_LAST_UPDATE`). Schema-order-fragile:
any column add/reorder silently corrupts writes.
- First seen: Manifest Cost (`INSERT_ManifestCost`). Contrast: Supplier/Logistics/Size **name** their columns.
- **Rebuild:** always use explicit column lists; with ActiveRecord this disappears (AR names columns).
  During a proc-wrap stage, audit every `VALUES(...)`-without-columns proc as fragile.

## P11 — Financial/master tables with zero integrity guards
A master that is **financially load-bearing** can still have **no PK constraint, no unique index, no
FK, no trigger, and no app-side dup check** (P1 absent). `INV_MANIFEST_COST_MST.MO_PRICE` is the EDI
810 invoice unit price, joined by **assy code alone** (the start/end manifest date window is ignored
by every billing consumer — `REPORT_EDI810/810Recreate/856`, `SELECT_INVOICEItems`,
`REPORT_INVOICESSummary/MonthlyINVOICESSummary`, `SELECT_ForecastDetailBCASN`). So duplicate
assy-code rows silently **double-bill**, and a deleted price silently drops invoice lines.
- First seen: Manifest Cost — the least-protected yet most financially critical master analyzed.
- **Rebuild:** give financial/master tables the **strongest** constraints even when legacy has none
  (real PK, unique-or-no-overlap on the business key, RI on delete). Always check what downstream
  procs actually JOIN on before trusting a table's extra/window columns.

## P12 — Wrong-target copy-paste inside the P8 retry recursion
The recursive retry branch (see P8 above) is copy-pasted boilerplate, and the retry call was
frequently **never renamed** — so it re-invokes a *different* method. A **full fleet audit of
`DataModule.pas` (2026-06-05) found 29 confirmed wrong-target retries, 0 false positives**:
**8 CRITICAL** (wrong-table write/DELETE keyed on the shared `fRecordID`/`fBroadCode`, P9),
8 MODERATE, 13 LOW (wrong SELECT only). The worst: **four `Delete*` methods recurse into
`DeleteSupplierInfo`** (ManifestCost, MonthlyPO, RenbanGroup, OvertimeHoliday) — a transient error
can delete an unrelated supplier by a borrowed id and fire `DELETE_SupplierCode`, blanking
`VC_SUPPLIER_CODE` across `INV_PARTS_STOCK_MST`. These fire only on the retry path, so they are
latent. Full register + per-method severity/line/fix:
[`docs/analysis/cross-cutting/datamodule-retry-target-bugs.md`](../../../docs/analysis/cross-cutting/datamodule-retry-target-bugs.md).
- **Root cause = P8 × P12 × P9 stacked** (per-method recursion × un-renamed retry × shared key).
- **Rebuild:** delete the in-method recursion; use one generic bounded transport-retry that
  re-invokes the *same* op, and pass keys as explicit per-call args (never shared singleton state).
  Removing any one of the three patterns kills the CRITICAL class.

## P13 — Feature-flag-gated navigation hub
A menu/hub form that owns **no data** toggles which child-module entries are visible — and
reroutes/relabels one entry — based on `[SITE]` INI booleans (`POEDISupport`, `GenerateEDI`). The
*same* flags are re-checked independently in `MainMenu.Configure`, so the gating logic is duplicated
across screens. (`MasterMaint` relabels its "&Monthly PO" button to "Manifest Cost" and reroutes it
to a different master when `fiGenerateEDI` is true.)
- First seen: `MasterMaint` (master-data hub).
- **Rebuild:** model the INI `[SITE]` flags as per-site settings (`Site#generates_edi?` etc.) and
  centralize visibility/route decisions in **one** policy object consumed by every nav — don't
  re-check flags per screen. Replace `.Visible` toggles with policy-gated links. See [[project-multisite]].

## P14 — `Hide; Child.Create; Execute; Free; Show` child-launch idiom
Parent forms open children with the uniform `Hide; Child := TChild.Create(self); Child.Execute;
Child.Free; Show;` dance (every `MasterMaint` button, and throughout `MainMenu`). Error safety is
inconsistent — most handlers don't wrap it in `try..finally Show`, so a child exception leaves the
parent hidden and bubbles to an `Application.Terminate` handler.
- First seen: `MasterMaint`, `MainMenu`.
- **Rebuild:** the whole modal dance disappears with RESTful routed pages. Do **not** preserve the
  terminate-on-unhandled-exception behavior — show an error and stay on the page.

## Multi-site lens (apply to every module)
The legacy app is single-site (identity in `INI [SITE]`). For each module ask:
**"what here is implicitly single-site?"** Common culprits:
- Tables with no site/tenant column but per-site data → need `site_id` scoping (Postgres phase).
- **Local filesystem paths** (e.g. `VC_BREAKDOWN_ORDER_DIRECTORY`) → replace with per-site
  configured targets (share/SFTP/object storage) or in-app delivery.
- Latent hooks already present, e.g. `BIT_SITE_NUMBER_IN_ORDER` — understand before reusing.
See [[project-multisite]] in memory.
