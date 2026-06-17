# Module Analysis: Forecast Detail / BOM master + thin forecast forms

**Area:** Forecasting  **Status:** ✅ spec complete  **Analyst:** Claude / 2026-06-15

> **The BOM/ratio master that the breakdown processor reads to explode an assembly into parts.**
> `ForecastDetail` is the CRUD editor over **`INV_FORECAST_DETAIL_INF`** — one row per
> *assembly (broadcast) part* × *effective month*, holding the component part codes (tire / wheel /
> valve / film / label / misc1 / misc2) and the **ratios** (tire %, wheel %, forecast %) that
> `ForecastBreakdownF.UpdateForecast` uses to split a weekly assembly forecast into per-part counts
> (see `forecast-breakdown.md` §4.4). **This screen is configuration, not transaction** — changing a
> ratio here changes every future breakdown run.
>
> This file also folds the **thin** forecasting units: `ManualForecast` (Buildout), `FRSBreakdown`
> (lot-split dialog), `ForecastCamexreport` (CAMEX report object), the `UploadBreakDown` dispatcher,
> and the **dead** `ForecastBreakDown.pas` / `ForecastUploadBreakDown.pas` duplicates.

## 1. Legacy surface
- **Form:** `ForecastDetail.pas` (565 lines / ~17 KB) + `.dfm`. `TForecastDetail_Form`; author
  David Verespey 2003-02-12. Registered live `InventorySystem.dpr:27`.
- **Entry point:** **Not** the `ForecastDetail1Click` menu handler — that one
  (`MainMenu.pas:3716`) is a **report** that runs `REPORT_ForecastDetail` straight to Excel and
  never opens this form. The editable form is opened elsewhere in the menu tree (a `…_Form.Execute`
  call); its `Execute` (`ForecastDetail.pas:98`) just `ShowModal`s. *(Confirm exact menu item; the
  Click handler that constructs `TForecastDetail_Form` is the live entry — body of `Execute` is the
  CRUD screen.)*
- **Purpose:** Maintain the assembly→component BOM with effective-dated ratios. Standard
  insert/update/search/delete grid + detail-panel pattern shared with the master-data forms.

## 2. Data touched
| Table | Read | Write | Notes |
|-------|:----:|:-----:|-------|
| `INV_FORECAST_DETAIL_INF` | ✓ | ✓ | **This module owns it.** PK `ID_FORECAST_DETAIL int IDENTITY` (schema:1390). Natural shape = `(VC_ASSY_PART_NUMBER_CODE, VC_EFFECTIVE_MONTH)` |
| `INV_FORECAST_INF` |  | ✓* | *Indirect via the **`DeleteForecastDetail`** trigger (cascade delete) |
| `INV_PARTS_STOCK_MST` | ✓ | | Component-part dropdowns via `SELECT_DependantPartNumber_PartType` (filtered by part TYPE: TIRE/WHEEL/VALVE/FILM/LABEL/MISC) |

**Triggers on `INV_FORECAST_DETAIL_INF`:**
- **`DeleteForecastDetail`** (schema:9586) FOR DELETE → `DELETE FROM inv_forecast_inf WHERE
  vc_part_number IN (SELECT vc_assy_part_number_code FROM DELETED)`. **Deleting a BOM row also
  deletes the raw forecast rows whose part number equals this assembly code.** Note it matches the
  *assembly* code against `INV_FORECAST_INF.VC_PART_NUMBER` (which stores assembly codes pre-explosion).
- **`UPDATE_ForecastDetailInf`** (schema:9608) FOR UPDATE → effectively a **no-op** (`print` only).

## 3. Stored procedures used
| Proc | Op | Rule |
|------|----|------|
| `SELECT_ForecastDetail` (schema:6014) | SELECT | Grid load (`GetForecastDetailInfo`, `@AssyCode=''`). Returns the full BOM column set aliased (`'Tire Part Number Code'`, `'Tire Ratio'`, `'Forecast Ratio'`, etc). |
| `INSERT_ForecastDetail` (schema:3060) | INSERT | Inserts a BOM row; sets `VC_ADD` = 16-char `yyyymmddHHMMSS` stamp (`CONVERT(char(8),112)` + four 2-char slices of `CONVERT(...,114)` — **verified 16 chars, well-formed**). **⚠️ Snapshot has only 12 params** (no Label/Misc) — see §4. |
| `UPDATE_ForecastDetail` (schema:8412) | UPDATE | Updates by `@RecordID = ID_FORECAST_DETAIL`; sets `VC_LAST_UPDATE` (same 16-char stamp). |
| `DELETE_ForecastDetail` (schema:2184) | DELETE | `DELETE … WHERE ID_FORECAST_DETAIL=@RecordID` → fires the cascade trigger above. |
| `SELECT_DependantPartNumber_PartType` | SELECT | Dropdown source per part type (body unverified; thin lookup). |
| `REPORT_ForecastDetail` | SELECT | Excel report (the `ForecastDetail1Click` path). Body unverified. |

## 4. Business rules & edge cases
- **Effective month** drives the breakdown's date selection (blank = default, `yyyy/mm` = month
  override; see `forecast-breakdown.md` §4.4). `FormCreate` builds a rolling 12-month combo
  (`ForecastDetail.pas:120-138`).
- **Ratio validation** (`HoldDetails`, `:268-307`): tire & wheel ratios clamped 0–100; **forecast
  ratio is intentionally NOT clamped** — the >100 guard is commented out (`:276-281`) to "Allow
  multiple ratios over 100". Assy qty radio maps to 1/2/4/5 (`:222-229`).
- **Shared-state hazard (P9):** `HoldDetails(True)` writes the selected row into
  `Data_Module.RecordID` + ~15 other shared `Data_Module.*` fields (`:193-211`), which the CRUD procs
  then read. Same single-mutable-record pattern as the masters; a concurrent/retry path can act on a
  stale `RecordID`.

> ### ⚠️ Snapshot-vs-live param divergence (NOT a confirmed bug — verify live)
> The live caller `InsertForecastDetailInfo` passes **`@LabelCode`, `@Misc1Code`, `@Misc2Code`**
> (`DataModule.pas:2330-2335`), and the form binds Label/Misc1/Misc2 part combos. But the
> **checked-in snapshot** `INSERT_ForecastDetail` (schema:3060) declares **only 12 params** (no
> Label/Misc) and `INV_FORECAST_DETAIL_INF`'s DDL (schema:1386) has **no label/misc columns**.
> (Corroborating the snapshot-lag: the live caller passes `@LabelCode/@Misc1Code/@Misc2Code`
> (DataModule.pas:2330-2335) AND the breakdown explosion reads `Label/Misc Part Number` from
> `SELECT_ForecastDetail` (ForecastBreakdownF.pas:1265-1286) — both would fail if the columns truly
> didn't exist, so production must have them. See [[reference-schema-snapshot-vs-live]].)
> Per `[[reference-schema-snapshot-vs-live]]`: treat as the snapshot lagging production (extra
> columns + 3 extra params added live), **not** as a runtime mismatch. **Action:** confirm the live
> proc signature + table columns before the rebuild. If the snapshot were accurate, every Insert
> would throw "too many parameters" — since the form is in daily use, the live proc almost certainly
> has them. `UpdateForecastDetailInfo` similarly passes a Label/Misc set (`DataModule.pas:2369+`).

> ### Cross-cutting P12 (already registered — no new finding)
> `GetForecastDetailInfo` retry calls **`GetSizeInfo`** on its `fErrorCount<3` branch
> (`DataModule.pas:2281`). This is the **🟡 LOW** entry already catalogued in
> [`cross-cutting/datamodule-retry-target-bugs.md`](../cross-cutting/datamodule-retry-target-bugs.md)
> (`Get…ForecastDetailInfo→GetSizeInfo`, wrong-SELECT only, loads the wrong dataset into the shared
> `Inv_DataSet`, no persistence). `Insert/Update/DeleteForecastDetailInfo` retries **correctly
> self-call** (verified `DataModule.pas:2358`) — **no new P12 bug in this module.**

## 5. Thin / supporting forms (folded)
- **`UploadBreakDown` (`UploadBreakDown.pas`, live, dpr:7)** — the **dispatcher** form for all six
  `TBreakdownKind`s. Pure file-picker + Start; routes `bForecast`→`ForecastBreakdownF`,
  `bBuildout`→`ManualForecast`, `bInvoice/bReceiving/bDailyBuildT/P`→Invoice/Logistics/DailyBuild
  specs. Per-kind file filters + per-assembler defaults (`WQS`⇒`*.prelftp`, `CAMEX`⇒`*.txt`).
- **`ManualForecast` (Buildout) (`ManualForecast.pas`, live, dpr:43)** — reads an Excel build-out
  plan (`fFilename`), and for each `xxxxx-xxxxx-xx`-formatted part row synthesizes a standard
  `.prelftp` forecast line spread evenly across the date range (`weekvalue = qty div weeks`,
  `ManualForecast.pas:99`), writing `fiForecastInputDir\buildout.prelftp`. **The call to actually run
  the breakdown on that file is commented out** (`:128-135`) — operator re-uploads it via
  `UploadBreakDown`. Date guards require start/end to be **Sundays** and end>start
  (`:153-185`). *(Note: `DayOfTheWeek=7` in `DateUtils` is Sunday only under ISO Mon-start — the
  "must be a Sunday" messages assume that convention; verify against the date-edit control.)*
- **`FRSBreakdown` (`FRSBreakdown.pas`, live, dpr:34)** — a **pure UI dialog**, **no DB**. Splits a
  total qty into N FRS lots: offers lot-counts `2..qty div lotsize` (+1 if remainder), and on change
  builds a comma list `(lots div n)+modt, lots div n, …` (the remainder goes to the **first** lot,
  `:130-138`). Used by the **Order** module, not the breakdown pipeline — included here only because
  the brief scoped it. Returns `BreakdownCount` + `Split` string to its caller.
- **`ForecastCamexreport` (`ForecastCamexreport.pas`, live, dpr:59)** — a `TObject` (not a form), not
  the breakdown writer. Reads **`REPORT_ForecastCAMEXReport(@WeekDate:=today)`** and pivots
  part×week qtys into a per-supplier Excel (`ForecastCamexTemplate.xls`) with SUM formulas, saved as
  `…-CFForecast`. Invoked from `MainMenu.pas:3839` (`CamexTest1Click`) and constructed in
  `Execute`'s scope in the breakdown form's uses. **`REPORT_ForecastCAMEXReport` is absent from the
  snapshot** — verify live (likely a later addition; this unit's header says 2015 vs the 2003 core).

## 6. Dead code (confirmed — do NOT spec as live)
- **`ForecastBreakDown.pas`** (no trailing **F**, 26 KB) — **NOT in `InventorySystem.dpr`** → dead.
  An older/duplicate of `ForecastBreakdownF.pas`. Verified absent from the manifest.
- **`ForecastUploadBreakDown.pas`** (3 KB) — **NOT in the dpr** → dead. The live dispatcher is
  `UploadBreakDown.pas`.
- Do not mine either for behavior; the live processor is `ForecastBreakdownF.pas`.

## 7. Target design (Ignition)
- **`INV_FORECAST_DETAIL_INF` → a Perspective CRUD screen** backed by Named Queries (insert/update/
  delete/select), mirroring the master-data pattern. This *is* screen-shaped (unlike the breakdown
  batch) — small editable grid + detail form.
- **D2:** resolve component parts and the assembly by **surrogate id**, not string code; the BOM rows
  FK to `INV_PARTS_STOCK_MST.IN_PART_ID`. **D3 (RESTRICT):** the `DeleteForecastDetail` cascade into
  `INV_FORECAST_INF` should become an explicit, blocked-if-referenced delete (or an archival flag),
  not a silent trigger cascade.
- **D1:** add `site_id` — BOM ratios are per-site config.
- **Effective-dating:** model the blank-vs-`yyyy/mm` override as a proper effective-date range so the
  breakdown service can resolve it deterministically.
- Fold the **Label/Misc** columns into the schema/model once their live existence is confirmed (§4).
- **CAMEX report** → a gateway export reusing `REPORT_ForecastCAMEXReport` (proc-wrap) once verified.

## 8. Open questions for the user (domain expert)
- **Q1 — confirm live `INSERT_/UPDATE_ForecastDetail` signatures + table columns** (Label/Misc1/Misc2).
  Snapshot lacks them; live caller passes them. (Candidate: fold into schema baseline.)
- **Q2 — BOM delete policy.** Today deleting a BOM row trigger-cascades raw forecast rows. Under D3
  should this be **blocked while referenced** (a forecast exists) instead? Or archived?
- **Q3 — forecast-ratio >100.** Confirmed intentionally unclamped (multi-share). Keep unbounded in the
  rebuild, or validate a max?
- **Q4 — `ForecastDetail` editable-form entry point.** Confirm which menu item opens the editable
  `TForecastDetail_Form` (vs the `ForecastDetail1Click` Excel report). Both should exist in the target.

## 9. Test cases / parity checks
- Insert a BOM `(assy, blank-month, tire=T1@60, wheel=W1@40, forecast=100)` → breakdown of a
  1000-count week yields tire 600 / wheel 400 (ties to `forecast-breakdown.md` §9).
- Add a `2026/07`-dated override row → July weeks use the override, other months use the blank default.
- Delete a BOM row → confirm the cascade removes the matching `INV_FORECAST_INF` rows (legacy) /
  is blocked (rebuilt under D3).
- Round-trip: edit a ratio → re-run breakdown → Order sim qty reflects the new ratio.
