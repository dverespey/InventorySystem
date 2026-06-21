# Module Analysis: Forecast Breakdown (`ForecastBreakdownF` → `INV_BREAKDOWN_FC_INF`)

**Area:** Forecasting  **Status:** ✅ spec complete  **Analyst:** Claude / 2026-06-15

> **The write-side spine of the whole forecasting area, and the 2nd-largest form in the app.**
> `ForecastBreakdownF` ingests a Toyota **EDI 830** material-release (or a legacy fixed-width text
> file), explodes each *assembly* (broadcast) forecast into its component **part numbers** using the
> BOM ratios in `INV_FORECAST_DETAIL_INF`, spreads each week's count across the working days of that
> week, and writes the result into **`INV_BREAKDOWN_FC_INF`** — the table the **Order "what to order"
> sim only READS** (see `order` spec, finding R1). It then emits per-supplier forecast files
> (TEXT `.frc` / Excel) and rolls usage up into `INV_SIZE_MST`. **This module owns the breakdown
> table.** Everything the Order sim consumes is produced here.
>
> **The single most load-bearing fact (reconciled with Order R1):** the breakdown is keyed and
> resolved by **ISO week number, year-blind**. The write side stores `IN_WEEK_NUMBER` =
> the EDI-supplied week (the `checkweeknumber`, *unmodified*); the read side
> (`SELECT_ForecastPartNumberWeek`) matches `IN_WEEK_NUMBER + VC_PART_NUMBER` with **no date/year
> filter**. There is a **first-production-day week offset that is applied asymmetrically** between
> write and read — documented as a hazard in §4.

## 1. Legacy surface
- **Form:** `ForecastBreakdownF.pas` (1489 lines / ~60 KB) + `.dfm`. Type
  `TForecastBreakdown_Form`; author David Verespey, 2003-03-10. Registered live in
  `InventorySystem.dpr:28` (`ForecastBreakdownF in 'ForecastBreakdownF.pas'`). The form itself is
  almost empty UI — a `THistory` log pane + one **OK** button; all behavior is in `Execute`.
- **Entry point:** `MainMenu.pas:248 Forecast_ButtonClick` →
  `UpBreakDown_Form.BreakdownKind:=bForecast; Execute` (`MainMenu.pas:251-254`). The dispatcher
  **`UploadBreakDown`** (`UploadBreakDown.pas:182-192`) creates `TForecastBreakdown_Form`, sets
  `.filename` (from a file-picker, default dir `fiForecastInputDir`) and `.SupplierCode`
  (`= fiSupplierCode`), `Show`s it, then calls `Execute`. **`UploadBreakDown` is the single live
  upload entry for all breakdown kinds** (forecast/invoice/receiving/buildout/dailybuild) — see §5.
- **Purpose:** Parse an external forecast feed → validate parts against the DB → delete the
  to-be-replaced slice of the breakdown table → re-explode assemblies into component parts × ratios
  → day-spread each week → upsert `INV_BREAKDOWN_FC_INF` → recompute `INV_SIZE_MST.IN_USAGE` →
  write per-supplier forecast files for transmission.

## 2. Data touched
| Table | Read | Write | Notes |
|-------|:----:|:-----:|-------|
| `INV_BREAKDOWN_FC_INF` | ✓ | ✓ | **This module owns it.** Deleted by `DELETE_ForecastInfo` (**CONFIRMED PRESENT** at `CreateInventory.sql:2725` — the earlier "missing" flag was snapshot drift; see §3), upserted by `INSERTUPDATE_BreakdownForecastInfo`, read back by `SELECT_ForecastSupplier` |
| `INV_FORECAST_DETAIL_INF` | ✓ | | The BOM/ratio master (see `forecast-detail.md`). `SELECT_ForecastDetail` supplies tire/wheel/valve/film/etc part codes + ratios that drive the explosion |
| `INV_FORECAST_INF` |  | ✓ | The *raw* per-assembly weekly forecast (pre-explosion). Upserted by `INSERTUPDATE_ForecastInfo` |
| `INV_PARTS_STOCK_MST` | ✓ | | `SELECT_PartsStockInfo` reads Line/Supplier/Size for each component part during day-spread |
| `INV_SIZE_MST` | ✓ | ✓ | `SELECT_SizeUsage` (read) + `UPDATE_SizeUsage` (write `IN_USAGE`) in the usage rollup |
| `INV_FIRST_PRODUCTION_DAY` | ✓ | | `SELECT_FirstProductionDay` — the week-offset source (§4) |
| `INV_SUPPLIER_MST` | ✓ | | `SELECT_SupplierInfo` reads output-file-type + directory + "site number in order" flag for file emission |
| *ALC DB* `INV_…` (special dates) | ✓ | | `AD_GetSpecialDateWeek` on the **ALC connection** — holiday/overtime day map for the week (cross-DB) |
| *site DUNS table* | ✓ | | `SiteTMMDUNSDataset` (param `@SiteTMMDUNS`) — EDI trading-partner validation (§4, D1 hook) |

**Triggers on these tables (from the schema):**
- **`DeleteForecastDetail`** on `INV_FORECAST_DETAIL_INF` (schema:9586) → `DELETE FROM inv_forecast_inf
  WHERE vc_part_number IN (SELECT vc_assy_part_number_code FROM DELETED)`. **Deleting a BOM row
  cascades a delete of the raw forecast rows** for that assembly code. (Relevant because this module
  reads that BOM and writes `INV_FORECAST_INF`.)
- **`UPDATE_ForecastDetailInf`** on `INV_FORECAST_DETAIL_INF` (schema:9608) → no-op (just `print`).
- `INV_BREAKDOWN_FC_INF` has **no triggers** — all its invariants live in the procs below.

## 3. Stored procedures used
| Proc | Op | Business rule (from the body) |
|------|----|-------------------------------|
| `INSERTUPDATE_BreakdownForecastInfo` (schema:2608) | UPSERT | **The breakdown writer.** Dedup key = **`(VC_SUPPLIER_CODE, VC_PART_NUMBER, IN_WEEK_NUMBER)`** — *year-blind, date-blind*. On exists → **ADDITIVE** `UPDATE … IN_QTYn = IN_QTYn + @Qtyn` (accumulates, does **not** overwrite). Else INSERT all 7 day-qtys + `VC_WEEK_DATE`, `VC_SIZE_CODE`. **`VC_WEEK_DATE` is written on INSERT only — never refreshed on UPDATE.** |
| `INSERTUPDATE_ForecastInfo` (schema:2655) | UPSERT | Raw per-assembly forecast into `INV_FORECAST_INF`. Dedup key `(Supplier, PartNumber, Kanban, WeekNumber)`; on exists overwrites `IN_COUNT` + `VC_WEEK_DATE`. **Not additive** (unlike the breakdown). |
| `DELETE_ForecastInfo;1` | DELETE | **✅ CONFIRMED PRESENT (correction — this was a snapshot-drift artifact).** In the **authoritative live dump** (`DB Schema/CreateInventory.sql:2725`, byte-identical in `Inventory` and `Inventory_Live`) the proc EXISTS with the exact 3-param `@WeekDate/@HistWeekDate/@PartNumber` signature the form calls (`ForecastBreakdownF.pas:115`). The `;1` is a legacy ADO procedure-group number. The earlier "DOES NOT EXIST" flag was written against the *superseded* 2026-06-01 snapshot; per `reference-schema-snapshot-vs-live`, `CreateInventory.sql` (6/12, no space) is authoritative. **The delete path is LIVE and works.** Semantics (verified): four deletes resolving assembly→its 7 component codes via a `CROSS APPLY (VALUES …)` over `INV_FORECAST_DETAIL_INF` — (A) breakdown forward `VC_WEEK_DATE >= @WeekDate`, (B) breakdown history `<= @HistWeekDate`, (C/D) raw `INV_FORECAST_INF` by assembly on the same two windows. Trims both ends, keeps the middle window. The old `DELETE_ForecastInfoWeekDate*` variants (`:2303/2318/2340`) are dead-but-present. See `docs/analysis/edi/inbound/forecast-import-algorithm.md` §E + `forecast-tables-analysis.md` §3.3. |
| `SELECT_ForecastDetail` (schema:6014) | SELECT | BOM lookup by `@AssyCode` (+ `@ForecastNotZero`, `@EffectiveMonth`). Returns tire/wheel/valve/film part codes, tire/wheel/forecast ratios, broadcast code, kanban, assy qty. With `@ForecastNotZero=1` filters `IN_RATIO<>0`; `@EffectiveMonth` keeps rows `>= month OR blank`. Ordered `assy, effective_month, ratio`. |
| `SELECT_PartsStockInfo` (schema:7270) | SELECT | Per-component part master read (Line/Supplier/Size). Used in day-spread. Body unverified (large proc; only the returned aliases `Line Name`/`Supplier Code`/`Size Code` are consumed). |
| `SELECT_FirstProductionDay` (schema:5982) | SELECT | Returns `INT_FIRST_PRODUCTION_WEEK` as **'First Week Number'** for `@ProdYear` (year prefix of the EDI WeekDate). Drives the week offset. |
| `AD_GetSpecialDateWeek` (ALC DB) | SELECT | **Not in this schema (ALC database).** Returns per-day `Date Status Abrv` ('H'/'X' = non-working) + `Day Number` for `@Week,@Line`. Body unverified (cross-DB). |
| `INSERTUPDATE_BreakdownForecastInfo` consumers: `SELECT_ForecastSupplier` (schema:6378) | SELECT | `SELECT * FROM INV_BREAKDOWN_FC_INF WHERE VC_WEEK_DATE > @WeekDate ORDER BY supplier,part,week_date`. Drives per-supplier file emission. **Filters on `VC_WEEK_DATE` (string compare), not week number.** |
| `SELECT_SizeUsage` (schema:7915) | SELECT | `INV_SIZE_MST ⋈ INV_PARTS_STOCK_MST` on `IN_SIZE_ID`; returns size→part→kanban. Drives usage rollup. |
| `UPDATE_SizeUsage` (schema:1006) | UPDATE | `UPDATE INV_SIZE_MST SET IN_USAGE=@Usage WHERE VC_SIZE_CODE=@SizeCode`. **Keys on `VC_SIZE_CODE` string** (D2: should resolve `IN_SIZE_ID`). Param `@SizeCode varchar(6)` but `VC_SIZE_CODE` is `varchar(10)` — silent truncation risk for >6-char codes. |
| `SELECT_ForecastPartNumberWeek` (schema:6309) | SELECT | **The READ side** (also used by Order). `@WeekNo,@DayNo,@PartNo` → returns `IN_QTY{DayNo}` from the breakdown WHERE `IN_WEEK_NUMBER=@WeekNo AND VC_PART_NUMBER=@PartNo`. **No supplier filter, no date/year filter.** Used here only inside `HistoryForecast` (usage rollup), and by Order for the actual order math. |

## 4. Business rules & edge cases — the breakdown algorithm

### 4.1 File ingest (two formats)
`Execute` (`:149`) opens `fFileName` and sniffs the **first line**:
- **EDI 830** if line 1 contains `ISA` (`:185`). It then:
  1. Splits the ISA on `*`; element **`delSL[4]`** is the **receiver/trading-partner DUNS**. It is
     validated via `SiteTMMDUNSDataset(@SiteTMMDUNS:=delSL[4])` (`:190-201`); **unknown DUNS ⇒ abort
     import** with a log line. *(This is the per-site EDI hook — see §6/D1.)* ⚠️ The abort log line is a
     **copy-pasted, misleading** message ("EDI file type=…expected type=830. Import fail.", `:199`) — it
     is NOT DUNS-specific; the rebuild should log the real reason (unknown trading-partner DUNS), not
     reproduce the legacy string.
  2. Reads to line 3, `copy(fcl,4,3)` must equal **`'830'`** else abort (`:206-213`).
  3. Counts records by counting **`LIN`** segments (`:228`).
  4. Per `LIN` loop: `Supplier := fiSupplierCode` (the *operator's* configured supplier, **not** from
     the file), `Partnumber := delSL[3]` (the assembly/BC part), `Kanban := delSL[5]`. Then for each
     **`FST`** segment: `WeekNumber := StrToInt(copy(delSL[9],3,2))` (chars 3–4 of element 9 — the
     2-digit week), `WeekDate := delSL[4]`, `WeekCount := StrToInt(delSL[1])`. The **first** FST sets
     `fFirstWeekDate/fFirstWeekNumber` (`:283-287`). (`ForecastBreakdownF.pas:262-295`.)
- **Legacy fixed-width text** otherwise: `ScanLine` (`:592`) parses **byte offsets** — Supplier@1(5),
  Part@6(12), Kanban@18(4), then **up to 14 week blocks** at offsets 22,36,…204 each = 2-char week +
  6-char weekdate + 6-char count; `'     -'` count ⇒ 0. **Only lines whose supplier == `fSupplierCode`
  are kept** (`:601`). Week 14 is read **only when `fiAssemblerName='WQS'`** (`:700-708`).

> **Hazard (text path width math):** the offset constants assume a rigid layout; a single shifted
> column corrupts every week. This is the same off-by-one class flagged for GALC frames.

### 4.2 Part validation & reconciliation (`ScanPartnumber`, `:724`)
- For each entry, `SELECT_ForecastDetail(@AssyCode:=Partnumber)`; `recordcount=0` ⇒ mark
  `Skip:=True` (part not a known assembly).
- Cross-checks the other direction (`@AssyCode='', @ForecastNotZero=1, @EffectiveMonth=now`) to list
  DB assemblies **missing from the feed**; prompts the operator (Yes/No) and writes Excel exception
  reports (`ForecastReport`, `ForecastDBError`, `ForecastRecError`) to `fiReportsOutputDir`.
- `fHistDate := now − (fiHistoricalForecast × 7)` days, `yyyymmdd` (`:318`) — passed to the delete proc
  as `@HistWeekDate` (the history-prune cutoff).

### 4.3 Delete-then-rebuild
For every non-skipped entry, `DeleteBreakdown(part)` (`:110`) calls `DELETE_ForecastInfo;1`
(**CONFIRMED PRESENT** at `CreateInventory.sql:2725` — see §3 correction). **Verified** semantics: the
proc resolves the assembly→its 7 component part codes via a `CROSS APPLY (VALUES …)` over
`INV_FORECAST_DETAIL_INF` and deletes breakdown rows for those components in **two windows** — forward
(`VC_WEEK_DATE >= fFirstWeekDate`, the slice about to be rebuilt) AND history
(`VC_WEEK_DATE <= fHistDate`, aged-out weeks) — plus the matching raw `INV_FORECAST_INF` rows by
assembly. It keeps only the middle window. Because the breakdown upsert is **additive** (§3), this delete
running FIRST is what makes a re-import a *replace* not a double-count — proven end-to-end in M2
(`scripts/e2e/test_forecast_import_e2e.py`: re-import the same 830 → qty NOT doubled). See
`docs/analysis/edi/inbound/forecast-import-algorithm.md` §E.

### 4.4 The explosion math (`UpdateForecast` `:1081` + `DoPartNumberForecast` `:1314`)
Per non-skipped entry, per week `j` in `1..count` (`count=14` if `WQS` else `13`, `:1089`):
1. Upsert the **raw** assembly forecast (`INSERTUPDATE_ForecastInfo`) with this week's count.
2. `SELECT_ForecastDetail(@AssyCode, @ForecastNotZero=1)` to get the BOM rows. **Effective-date
   rule** (`:1151-1176`): a row with **blank** `Active Date` is the default; a row whose
   `Active Date` *yy/mm* (`copy(3,2)+copy(6,2)`) equals the week's *yymm* (`copy(WeekDate,1,4)`)
   **overrides and breaks**. If neither matched (`bd=False`) ⇒ log *"No breakdown for part…count will
   be ignored"* and the ratios stay 0 → component counts become 0 (`:1178-1183`).
3. **Ratio math (integer, truncating)** when all three ratios ≠ 0 (`:1186-1203`):
   ```
   tirecount  = (((WeekCount * forecastratio) div 100) * tireratio)  div 100
   wheelcount = (((WeekCount * forecastratio) div 100) * wheelratio) div 100
   ```
   The older 50/50-share variants are **commented out** (`:1206-1225`) — current behavior gives each
   the full `WeekCount` scaled by its own ratio (the "TEMA sends full amount per share line" model,
   `:1191-1199`).
4. For each non-empty component slot — **Tire→tirecount; Wheel/Valve/Film/Label/Misc1/Misc2 →
   `wheelcount`** (`:1230-1287`; note valve/film/label/misc all reuse `wheelcount`, **not** their own
   count — a probable legacy simplification, flag in §8) — call `DoPartNumberForecast(PN, WeekDate,
   count, WeekNumber)`. A slot is "present" iff its part code length `> 2`.

### 4.5 Day-spread (`DoPartNumberForecast`, `:1314`) — the per-day breakdown
1. Default working days Mon–Fri true, Sat/Sun false, `days=5` (`:1321-1328`).
2. `SELECT_PartsStockInfo(PN)` → Line (`'ALL LINES'` if blank), Supplier, Size.
3. **Week-offset (write side):** `checkweeknumber := WeekNumber` is captured first (`:1358`). Then if
   `fiUseFirstProductionDay`: `SELECT_FirstProductionDay(@ProdYear:=copy(WeekDate,1,4))`; if
   `First Week Number ≠ 1`, `WeekNumber := WeekNumber + (First Week Number − 1)` — **ADDS** the offset
   to the *local* `WeekNumber` (`:1371-1375`).
4. `AD_GetSpecialDateWeek(@Week:=WeekNumber, @Line:=Line)` on the **ALC** connection: for each
   returned day, `'H'`/`'X'` ⇒ that day non-working (`DEC days`), else working (`INC days`)
   (`:1394-1410`).
5. Spread: `ratiocount = FCCount div days; leftover = FCCount mod days` (`:1416`). Then for days
   1..7: a working day gets `ratiocount + leftover` and **leftover is then zeroed** — so **all
   remainder lands on the first working day**; non-working days get 0 (`:1424-1434`).
6. Upsert via `INSERTUPDATE_BreakdownForecastInfo` writing **`@WeekNumber := checkweeknumber`**
   (the *unmodified* week, `:1446`) + Supplier + PN + Size + the 7 day-qtys.

> ### ⚠️ R1-RECONCILIATION — the write/read week-number asymmetry (CRITICAL, document, do not "fix")
> - **WRITE side stores the *unmodified* week** (`checkweeknumber`). The offset added at `:1374` is
>   applied to the local `WeekNumber` that is used **only** for the `AD_GetSpecialDateWeek` holiday
>   lookup — it is **NOT** the value written to the breakdown row. So `INV_BREAKDOWN_FC_INF.IN_WEEK_NUMBER`
>   = the raw EDI/text week number, year-blind.
> - **READ side (Order R1, `Order.pas:1145-1196`) SUBTRACTS the offset**: `@WeekNo =
>   WeekOfTheYear(prodDate) − weekoffset` where `weekoffset = First Week Number − 1`
>   (`Order.pas:1162,1175-1178`), with the underflow guard *(if `WeekOfTheYear−weekoffset < 1` then use
>   `WeekOfTheYear` unadjusted, `:1175-1176`)*.
> - **Net effect:** Order computes a *production-week ISO number*, subtracts the first-production-week
>   offset to get a "production-relative" week, and matches it against `IN_WEEK_NUMBER`. For that match
>   to land, the stored `IN_WEEK_NUMBER` must already be **in production-relative terms** — i.e. the
>   raw EDI/`FST` week number from TEMA is *itself* production-relative (TEMA numbers weeks from the
>   plant's production-week-1, not the calendar ISO week). The write side's `+offset` adjustment is
>   confined to the holiday lookup precisely because the stored week is left in the feed's own
>   numbering. **This is internally consistent only if the feed's week numbering and Order's
>   `WeekOfTheYear − offset` produce the same integer for the same physical week** — which is the
>   load-bearing assumption the rebuild must pin with real data (§8 Q1). The year-blindness means a
>   forecast row written in week 30 of 2024 will be read for week 30 of *any* year until overwritten —
>   safe only because the delete-then-rebuild (4.3) clears the forward slice each cycle. (Confirmed
>   matches Order R1.)

### 4.6 Usage rollup (`UpdateUsage` `:951` + `HistoryForecast` `:1031`)
- `SELECT_SizeUsage` walks size→part. For each part, `HistoryForecast` sums the daily breakdown qtys
  over `fiUsageUpdateCompare` weeks × 7 days via `SELECT_ForecastPartNumberWeek(WeekOfTheYear(now+z),
  DayOfTheWeek(now+y), part)` and **averages** (`total div count-of-nonzero-days`, `:1066-1067`).
  ⚠️ The read here uses `WeekOfTheYear(now+z)` **without** the first-production-day offset (unlike both
  Order and the write side) — a third week-number convention in the same file. Flag §8 Q1.
- Per size, `UPDATE_SizeUsage(SizeCode, summed-usage)` (only when usage ≠ 0).

### 4.7 File emission
After the rebuild, `SELECT_ForecastSupplier(@WeekDate:=today)` streams breakdown rows ordered by
supplier; per supplier it reads `SELECT_SupplierInfo` for **Output File Type** (TEXT/EXCEL/BOTH),
**Directory**, and **Site Number in Order** flag; writes a fixed-width `.frc` text line
(`%.2d` week, `%.5d`×7 day-qtys, `:499-506`) and/or an Excel file from `ForecastTemplate.xls`,
optionally archived if `fiLocalFTP`. The `sendsite` flag prefixes `SiteSupplierCode` to the line
(`:486-490`) — a multi-site-aware output path that already exists.

> **Timestamp note (no miscount):** report filenames use `formatdatetime('yyyymmddhhmmss00',now)`
> = 16 chars but **literal `00`** for centiseconds (`:846,853`), not a real ff. The DB add-stamps in
> the forecast-detail procs are a true 16-char `yyyymmdd`+`hhmmss`+2 (CONVERT 112 `char(8)` + four
> 2-char slices of CONVERT 114) — verified well-formed (`forecast-detail.md` §4).

## 5. UI / UX notes
- Effectively head-less: pick a file → watch the `THistory` log → click OK to dismiss. No editable grid.
- `UploadBreakDown` is a thin dispatcher form (file-picker + Start) reused for **6** breakdown kinds
  (`TBreakdownKind`, `UploadBreakDown.pas:23`); `bForecast`→here, `bBuildout`→`ManualForecast`,
  others→Invoice/Logistics/DailyBuild specs. File-type filters and default dirs are per-kind and
  per-assembler (`WQS` ⇒ `*.prelftp`, `CAMEX` ⇒ `*.txt`, `:111-126`).
- **Keep vs modernize:** the whole thing is a batch ETL — in the target it belongs in a **gateway
  service**, not a screen (§6). Keep the exception reports (parts in feed not in DB / in DB not in
  feed) as operator-facing output; they are genuinely useful.

## 6. Target design (Ignition)
This area is **math-heavy batch ETL with cross-DB calls and Excel/file I/O** — the canonical case for
a **gateway Python/Jython service**, not Named-Query-per-screen.

- **Ingest:** a gateway script (file-watch on the forecast input dir, or message-handler) parses the
  830/text feed. Use a real X12 parse, not byte offsets, but **preserve the exact element map**
  (LIN→part/kanban, FST element 9 chars 3–4 = week, element 4 = weekdate, element 1 = count).
- **Breakdown math → Python service**, not a proc: the ratio explosion, effective-date selection,
  day-spread, and remainder-on-first-working-day are imperative integer math that is far clearer (and
  testable) in Python than T-SQL. Reimplement `UpdateForecast`/`DoPartNumberForecast` as a pure
  function `explode(entry, bom_rows, working_days) -> rows`. Unit-test against the parity vectors in §9.
- **Persistence → Named Queries / proc-wraps:** keep `INSERTUPDATE_BreakdownForecastInfo` semantics
  (additive upsert on `supplier+part+week`) and the delete-forward as **idempotent NQs**.
  `DELETE_ForecastInfo` **exists and is reused as-is** (the M2 importer wraps it via a proc-call — the
  "missing proc" §8 Q2 is RESOLVED, see §3 correction); the rebuild runs it FIRST per assembly in one tx
  so the additive upsert behaves as a replace. **Built + proven in M2:**
  `docs/analysis/edi/inbound/project-library/forecast/code.py` (the pure `explode`/`day_spread` +
  `import_830` driver) and `scripts/e2e/test_forecast_import_build.py` / `test_forecast_import_e2e.py`.
- **Working-day calendar:** `AD_GetSpecialDateWeek` lives in the ALC DB today; in Ignition model the
  plant calendar as a first-class table/service the breakdown service queries.
- **Output files:** the per-supplier `.frc`/Excel generation → a gateway report/export step (Perspective
  download or scheduled drop), reusing the supplier output-type/directory config (which under **D1**
  moves into the `sites`/supplier rows).
- **D1 multi-site hook:** the **`delSL[4]` DUNS** validation (`:188-201`) is the per-site EDI filter —
  it already routes a feed to a site by trading-partner. In the rebuild this selects the **`site_id`**;
  the entire breakdown run is then scoped to that site. **This is the concrete forecast-ingest D1 hook.**
- **D2:** `UPDATE_SizeUsage` and the breakdown keys resolve by **string** today (`VC_SIZE_CODE`,
  `VC_PART_NUMBER`, `VC_SUPPLIER_CODE`); rebuild resolves by surrogate id.
- **Reports:** `SELECT_ForecastSupplier` (file feed), plus the Excel exception reports.

## 7. Migration plan
- [ ] Stage 1 — wrap `SELECT_ForecastPartNumberWeek`/`SELECT_ForecastSupplier` read-only; reproduce
      Order parity against the live breakdown table (already partly done via the Order spike).
- [ ] Stage 2 — port ingest + explosion to a gateway service writing through the existing upsert
      proc; run **in parallel** with the Delphi run and diff `INV_BREAKDOWN_FC_INF` row-for-row.
- [ ] Stage 3 — reimplement the upsert/delete in app (Postgres-ready), add `site_id` scoping (D1),
      surrogate-key resolution (D2), and a defined delete predicate.

## 8. Open questions for the user (domain expert)
- **Q1 ✅ RESOLVED (D10) — the canonical week-number convention.** Three week computations coexist: write
  stores the **raw feed week** (`checkweeknumber`); Order reads **`WeekOfTheYear(prod) − (FirstWeek−1)`**;
  the usage rollup reads **`WeekOfTheYear(now)`** with no offset. **A real TMMMS 830 (`EDI/830000008976.EDI`)
  settles it: TEMA's week number is supplied in the FST09 "DO" reference (`copy(delSL[9],3,2)`, e.g. `2624`
  → wk 24) and is PRODUCTION-RELATIVE — exactly `ISO_week(date) − 1` for 2026 across the whole horizon
  (= `INT_FIRST_PRODUCTION_WEEK[2026] − 1 = 1`).** So the stored raw week already equals Order's
  offset-subtracted ISO week → the write/read conventions reconcile, and the shipped R1 fix is validated.
  Rebuild: ingest `IN_WEEK_NUMBER` from FST09 verbatim (don't recompute); read with `ISO(date) − offset`.
  See decision **D10**.
- **Q2 ✅ RESOLVED — `DELETE_ForecastInfo` semantics.** The "absent from the snapshot" flag was
  snapshot drift — the proc is **CONFIRMED PRESENT** at `CreateInventory.sql:2725` (live-identical in both
  DBs) with the exact 3-param signature (§3 correction). It deletes, **per assembly** (resolving
  assembly→its 7 component codes via `CROSS APPLY (VALUES …)` over the recipe): breakdown rows
  `VC_WEEK_DATE >= @WeekDate` (forward) AND `<= @HistWeekDate` (history-prune), plus the raw forecast by
  assembly on the same two windows — keeping only the middle window. **Not a live runtime error.** The M2
  importer reuses it as-is; the delete-then-rebuild is reproduced + proven (`test_forecast_import_e2e.py`).
- **Q3 — year-blindness.** Is forecasting genuinely intended to be year-agnostic (week 30 = week 30
  forever until overwritten), relying on the delete-forward each cycle to avoid stale carryover? Or
  should the rebuild add a year/effective-date dimension to the breakdown key?
- **Q4 ✅ RESOLVED (D11) — valve/film/label/misc qty.** All non-tire components are day-spread using
  **`wheelcount`** (`:1242-1287`), not a per-component count. David confirmed this is a **bug to fix** —
  the rebuild day-spreads each component on its OWN count.
- **Q5 — additive vs replace upsert.** `INSERTUPDATE_BreakdownForecastInfo` is **additive** (`+= qty`).
  Is a re-run within the same delete-window meant to accumulate, or is that a bug masked by the
  delete-first step? (If a delete is skipped/failed, counts double.)
- **Q6 — `ManualForecast`/Buildout is half-wired.** It writes a synthetic `.prelftp` but the call to
  run the breakdown is **commented out** (`ManualForecast.pas:128-135`) — operator must re-upload it
  manually. Keep two-step, or auto-chain in the rebuild?

## 9. Test cases / parity checks
- **Day-spread:** `FCCount=11, days=5` → `[3,2,2,2,2,0,0]` (remainder 1 on day 1). `FCCount=10, days=5`
  → `[2,2,2,2,2,0,0]`. With a Wed holiday (`days=4`, Wed off) `FCCount=11` → `[3,2,0,2,2,0,...]`
  *(remainder on first working day; verify exact index against `AD_GetSpecialDateWeek` day-number map)*.
- **Ratio:** `WeekCount=1000, forecastratio=100, tireratio=60` → `tirecount = ((1000*100 div100)*60)div100 = 600`.
- **Additive upsert:** two explosion passes for the same `(supplier,part,week)` without an intervening
  delete must **sum**; with a delete must **replace**.
- **Round-trip with Order:** a known 830 → run breakdown → run the Order sim → the forecast qty Order
  reads for a given prod-date/part must equal the day-qty this module wrote (modulo the Q1 week math).
- **Effective-date override:** a BOM with a blank-date default and a `yy/mm`-dated override must pick
  the override only in its month, else the default.
