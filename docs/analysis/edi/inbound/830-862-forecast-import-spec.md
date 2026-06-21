# EDI 830 / 862 Forecast Import — behavioral spec (source-of-truth)

> **M2 foundational unit.** Decodes how the forecast that M1's ASN-creation
> (`SELECT_ForecastDetailBCASN`) and the Order both read is *built*. The **830 (DELFOR,
> planning forecast)** is the inbound that populates `INV_FORECAST_DETAIL_INF` consumers and
> `INV_BREAKDOWN_FC_INF` (the week-bucket day-qtys the Order reads). The **862 (DELJIT, firm/JIT)**
> is *not* a forecast writer in this codebase — it is report-only.
>
> **Scope guard:** through M1 the forecast was treated as static. This spec is the WRITE side.
> The READ side (`SELECT_ForecastDetailBCASN`, `SELECT_ForecastPartNumberWeek`) is covered in the
> ASN keystone / Order specs; cited here only where the contract joins.

**Confidence:** HIGH on the parse + table population + procs (all proc bodies read in
`DB Schema/CreateInventory.sql`, the authoritative live dump dated 2026-06-12). Two cross-DB procs
(`AD_GetSpecialDateWeek`, `AD_GetSiteTMMDUNS`) live in the ALC/VehicleOrder DB and are body-unverified
here (noted inline). Data-dependent claims are flagged with the exact cell to confirm against a golden 830.

**Live-unit confirmation (`InventorySystem.dpr`):**
`ForecastBreakdownF in 'ForecastBreakdownF.pas'` (:28) and `EDIUpload in 'EDIUpload.pas'` (:53) — both ship.
`ForecastCamexreport` is `used` by the breakdown unit (:106) but the report path is dormant (the var is declared, never instantiated). `Orderold`/`Order1`/`*old` are dead (not in DPR).

---

## 1. The 830 parse (`ForecastBreakdownF.pas`)

### 1.1 Dispatch & two entry points
The 830 reaches the parser two ways:
1. **Inbound poller** — `EDIUpload.Execute` (`EDIUpload.pas:53`) walks `fiEDIIn`, sniffs `ISA` on line 1,
   resolves the DUNS, reads line 3 `copy(fcl,4,3)`; on `'830'` it creates and runs the breakdown form
   (`EDIUpload.pas:89-104`): sets `filename` + `SupplierCode := fiSupplierCode`, calls `Execute`, frees it,
   logs `EDIIMP / EDI 830 Imported`.
2. **Manual file-picker** — operator-triggered against `ForecastBreakdown_Form` directly (the same
   `Execute`; see `edi-upload.md:169`). Both converge on `TForecastBreakdown_Form.Execute` (`:149`).

### 1.2 File sniff & DUNS guard (`Execute`, :184-217)
- Line 1 must contain `ISA` (`:185`). Split on `*` → element **`delSL[4]`** is the trading-partner DUNS,
  validated via `SiteTMMDUNSDataset(@SiteTMMDUNS := delSL[4])` (`:190-201`). **Unknown DUNS ⇒ `exit`**
  (import aborts). `SiteTMMDUNSDataset` runs `AD_GetSiteTMMDUNS` on **`ALC_Connection`** (the
  VehicleOrder/ALC DB, not the inventory DB) — body unverified (cross-DB).
- Reads to line 3; `data := copy(fcl,4,3)` must equal **`'830'`** else `exit` with a `FORECAST` log
  (`:206-213`). On success sets `EDIfile := TRUE`, then `Reset(fcf)` to re-read from the top.

> ⚠️ **DUNS element-index hazard (shared with M1, see `997-824-inbound-spec.md` §5.2).** `delSL[4]`
> is the **5th** `*`-delimited element = **ISA04** (sender-ID/security slot, authoring-by-code). On a
> file *we* generate ISA04 = our SiteDUNS; the receiver (TMM) DUNS sits at ISA08 = `delSL[8]`. The guard
> matches `delSL[4]` against the **TMM** DUNS column. For an inbound TEMA 830 to ever pass, TEMA's ISA
> must place Toyota's own DUNS at ISA04. **Confirm the exact ISA04 value against the golden
> `EDI/830000008976.EDI` ISA (sender `808369495`, receiver `71930`)** — the green DUNS tests prove guard
> *mechanics only*, not the element index. (D10's golden has sender DUNS = Toyota at ISA06, so re-check
> which slot the real feed uses before trusting the index.)

### 1.3 Record count & the segment walk (X12 DELFOR)
- **Count pass** (`:221-233`): counts `LIN` segments (`copy(fcl,1,3)='LIN'`) → `count`; `SetLength(fEntries,count)`.
- **Per-LIN loop** (`:245-298`): advance to `LIN` (or `CTT` = end-of-transaction → break). For each LIN:
  - `fEntries[].Supplier := fiSupplierCode` — **the operator's configured supplier, NOT from the file** (`:262`).
  - `fEntries[].Partnumber := delSL[3]` — the **assembly / broadcast (BC) part number** (`:263`, LIN03).
  - `fEntries[].KanbanNumber := delSL[5]` (`:264`, LIN05). `Skip := False`.
  - Advance to first `FST` (`:268-272`).
  - **FST loop** (`:276-295`) — one `TWeekData` per FST segment, up to 14 (`Weeks[1..14]`):
    - **`WeekNumber := StrToInt(copy(delSL[9],3,2))`** — chars **3–4 of FST09** (the "DO" reference). See §1.4.
    - **`WeekDate := delSL[4]`** — **FST04** = forecast start date `yyyymmdd` (8 chars).
    - **`WeekCount := StrToInt(delSL[1])`** — **FST01** = the forecast quantity for the bucket.
    - The **first** FST sets module-level `fFirstWeekDate` / `fFirstWeekNumber` (`:283-287`) — used as the
      delete-window anchor (§2.2).

  > Note: the parser keys off `delSL[N]` 1-based after `splitString`. `delSL[0]` = `"FST"`, so `delSL[1]`=FST01
  > (qty), `delSL[4]`=FST04 (start date), `delSL[9]`=FST09 (DO ref). SDQ / N1 / REF segments are **not parsed**
  > by this importer — only LIN + FST. (TEMA's 830 here carries delivery via FST, not SDQ.)

### 1.4 The week-number derivation (D10 — production-relative, ISO − 1)
**`WeekNumber = chars 3–4 of FST09 ("DO" ref).`** For `FST*144*D*W*20260615*20260619***DO*2624`,
`delSL[9]="2624"`, `copy(.,3,2)="24"` → week 24. The DO ref is `2`+`6`(year 2026)+`WW` → "2026, week 24".

- **TEMA's numbering is production-relative, exactly `ISO_week(start) − 1` for 2026** (validated across the
  whole horizon in decisions.md D10:298). The forecast row stores this verbatim — see §1.5.
- This equals `weekoffset = INT_FIRST_PRODUCTION_WEEK[2026] − 1 = 2 − 1 = 1`. The Order READ side computes
  `@WeekNo = ISO(prodDate) − weekoffset = ISO − 1`, so READ and WRITE match on real data (D10 confirms the
  shipped `SIM_OrderSimulation` R1 fix). **The feed being production-relative is load-bearing.**

### 1.5 The FirstProductionDay offset is NOT stored on the row (critical, validated by data)
`DoPartNumberForecast` (`:1314`) keeps two week variables:
- `checkweeknumber := WeekNumber` (`:1358`) — the **raw TEMA week** from FST09.
- If `fiUseFirstProductionDay` (INI `[INIT] UseFirstProductionDay`, default `True`), it calls
  `SELECT_FirstProductionDay(@ProdYear := copy(WeekDate,1,4))` (`:1365-1369`) and, when `First Week Number ≠ 1`,
  does `WeekNumber := WeekNumber + (First Week Number − 1)` (`:1374`) — **mutating only the LOCAL `WeekNumber`**.
- The breakdown INSERT (`:1446`) writes **`@WeekNumber := checkweeknumber`** — the **raw, un-offset** week.
  The offset-adjusted `WeekNumber` is used ONLY for the holiday lookup `AD_GetSpecialDateWeek(@Week := WeekNumber)`
  (`:1390`).

> ⚠️ **Data-adjudicated source disagreement (team retro 2026-06-15, R1).** A naive reading of `:1374` suggests
> the *stored* week is offset by `+(FirstWeek−1)`. **It is not** — `checkweeknumber` (raw) is what lands in
> `IN_WEEK_NUMBER`. The spike data won this twice. **Rebuild MUST store the raw FST09 week unmodified.** The
> exact value to confirm against the golden: for `2624`, `INV_BREAKDOWN_FC_INF.IN_WEEK_NUMBER` must = **24**
> (not 25). If a rebuild ports `:1374` literally onto the stored value, every row drifts by +1 and the Order
> read silently mismatches.

### 1.6 BC / part mapping & the `VC_BROADCAST_CODE` LIKE relationship
The 830 LIN03 (`delSL[3]`) is the **assembly part number** = the BOM key, NOT itself the broadcast code.
- The importer explodes the assembly via `SELECT_ForecastDetail(@AssyCode := Partnumber, @ForecastNotZero:=1)`
  (`:1127-1133`), which returns the BOM row: tire/wheel/valve/film/label/misc1/misc2 part codes + ratios +
  `VC_BROADCAST_CODE` (`SELECT_ForecastDetail`, schema:3154).
- The **`VC_BROADCAST_CODE LIKE` pattern is consumed downstream, not at import.** `SELECT_ForecastDetailBCASN`
  (schema:3011) does `WHERE @BCode LIKE VC_BROADCAST_CODE` — i.e. the stored detail row holds a SQL **pattern**
  (e.g. `42610%`) and the runtime BC is matched against it. The importer only needs `@AssyCode` to match exactly;
  the LIKE/wildcard semantics belong to `INV_FORECAST_DETAIL_INF` config (the BOM master, see
  `forecasting/forecast-detail.md`) and the ASN read. **The 830 import does NOT write `VC_BROADCAST_CODE`** —
  it is pre-existing config keyed by assembly part. (The importer treats an assembly with no BOM row as "skip".)

---

## 2. Tables populated + replace/merge semantics

### 2.1 The three tables & the writer procs

| Table | Written by | Key / semantics |
|---|---|---|
| `INV_FORECAST_INF` | `INSERTUPDATE_ForecastInfo` (schema:1184) | Raw per-assembly forecast. Dedup key **`(VC_SUPPLIER_CODE, VC_PART_NUMBER, VC_KANBAN_NUMBER, IN_WEEK_NUMBER)`** on the EXISTS check; the UPDATE matches on `(Supplier, Part, WeekNumber)` (kanban dropped). On exists → **OVERWRITE** `IN_COUNT`, `VC_WEEK_DATE` (NOT additive). Else INSERT. |
| `INV_BREAKDOWN_FC_INF` | `INSERTUPDATE_BreakdownForecastInfo` (schema:1215) | **This module owns it** — the day-qty buckets the Order reads. Dedup key **`(VC_SUPPLIER_CODE, VC_PART_NUMBER, IN_WEEK_NUMBER)`** — **year-blind / date-blind**. On exists → **ADDITIVE** `IN_QTYn = IN_QTYn + @Qtyn` (accumulates). Else INSERT 7 day-qtys + `VC_WEEK_DATE` + `VC_SIZE_CODE`. **`VC_WEEK_DATE` written on INSERT only, never refreshed on UPDATE.** |
| `INV_FORECAST_DETAIL_INF` | **NOT written by the import** (config master; maintained by `ForecastDetailF` screen) | Read via `SELECT_ForecastDetail` for the assembly→component BOM + ratios + BC pattern. |

Plus a side effect: `INV_SIZE_MST.IN_USAGE` is rolled up via `UPDATE_SizeUsage` (§4).

### 2.2 Replace/merge = **delete-window-then-additive-upsert** (NOT a clean overwrite)
The import is **not** a wholesale "delete the month, re-insert." It is a per-part windowed delete + additive accumulate:

1. **Delete window** — for every non-skipped entry, `DeleteBreakdown(part)` (`:110`, called in the loop at
   `:322-329`) runs **`DELETE_ForecastInfo;1`** (schema:**2725** — see §6 hazard correction) with
   `@WeekDate := fFirstWeekDate`, `@HistWeekDate := fHistDate`, `@PartNumber := <assembly part>`.
   `DELETE_ForecastInfo` deletes, **for each component part of that assembly** (via a `CROSS APPLY (VALUES …)`
   over the 7 component-part columns of `INV_FORECAST_DETAIL_INF` where `VC_ASSY_PART_NUMBER_CODE=@PartNumber`):
   - `INV_BREAKDOWN_FC_INF` rows with `VC_WEEK_DATE >= @WeekDate` (the forecast horizon being refreshed), AND
   - `INV_BREAKDOWN_FC_INF` rows with `VC_WEEK_DATE <= @HistWeekDate` (prune stale history),
   - and the matching `INV_FORECAST_INF` rows by `@PartNumber` on the same two date windows.
   `@WeekDate = fFirstWeekDate` = FST04 of the first FST = the start of the new horizon; `@HistDate =`
   `now − fiHistoricalForecast*7` days (`:318`, INI `[INIT] HistoricalForecast` default `12` → 84 days back).
   So: **wipe the future horizon for these parts, prune old history, leave the gap in between untouched.**

2. **Re-explode + additive insert** — `UpdateForecast` (`:1081`) loops entries × weeks (13 weeks, or **14 if
   `fiAssemblerName='WQS'`**, `:1089-1092`), calls `INSERTUPDATE_ForecastInfo` (raw), reads the BOM ratios
   (`SELECT_ForecastDetail`), computes tire/wheel counts, then `DoPartNumberForecast` per component →
   `INSERTUPDATE_BreakdownForecastInfo` (additive).

> ⚠️ **Replace-correctness depends on the delete actually running first.** Because the breakdown upsert is
> **additive** and keyed year-blind on `(supplier, part, week#)`, re-importing the *same* 830 WITHOUT the
> delete would **double** every qty (and a week-30-2026 row would accumulate onto a leftover week-30-2025 row,
> since the key ignores year — see year-blindness below). The delete window is the idempotency guard. **The
> rebuild MUST run the per-part window delete before the additive accumulate, in the same transaction**, or
> switch the breakdown writer to overwrite-by-key. This is the M2 analogue of GALC's proc-side dedup.

> ⚠️ **Year-blind key.** Both `IN_WEEK_NUMBER` and the dedup key carry no year. Week 30 of 2026 and week 30 of
> 2027 collide. The system is single-year-window in practice (the delete prunes >84 days old), but the rebuild's
> multi-site/multi-year forecast MUST add the year (or the FST04 month) to the key. Confirm: a golden import
> spanning a Dec→Jan rollover (weeks 52→01) is the edge case — see §6.

### 2.3 The ratio explosion (assembly → component qty)
For each BOM row (`:1151-1226`): pick the ratio set by **effective month** —
- blank `Active Date` (`VC_EFFECTIVE_MONTH`) = the **default** ratio (`bd:=TRUE`),
- else match `copy(ActiveDate,3,2)+copy(ActiveDate,6,2)` (yy+mm of the ratio's `yyyy/mm/...`) against
  `copy(WeekDate,1,4)` (yyyy → **note: this compares ratio "yymm" vs weekdate "yyyy"** — a 4-char vs 4-char
  string compare that is *not* obviously aligned; see hazard §6),
- if neither, `bd=FALSE` → log "No breakdown for part number…" and **count is silently ignored** (`:1178-1183`).

If all three ratios non-zero: `tirecount = ((WeekCount * forecastratio /100) * tireratio /100)`,
`wheelcount = (… * wheelratio /100)` (integer div, `:1202-1203`). Each component part code (tire/wheel/valve/
film/label/misc1/misc2, where length>2) gets `DoPartNumberForecast(partcode, weekdate, count, weeknumber)`.
**Valve/film/label/misc all use `wheelcount`** (`:1248-1287`) — not their own ratio. (Faithful to source; flag for rebuild review.)

### 2.4 Day-spread (week qty → IN_QTY1..7)
`DoPartNumberForecast` (`:1314-1480`):
- `SELECT_PartsStockInfo(PN)` → Line (`'ALL LINES'` if blank), Supplier, Size (`:1339-1351`).
- Default workdays Mon–Fri (`workday[1..5]=true`, 6/7 false; `days=5`, `:1321-1328`).
- `AD_GetSpecialDateWeek(@Week := WeekNumber(offset-adjusted), @Line := Line)` on **`ALC_Connection`**
  (`:1387-1393`, cross-DB, body unverified): each row with `Date Status Abrv` ∈ {`H`,`X`} turns a day OFF
  (`DEC(days)`); else ON (`INC(days)`).
- Spread: `ratiocount = FCCount div days`, `leftover = FCCount mod days`; the leftover is dumped on the **first**
  working day (`:1424-1434`).
- Write via `INSERTUPDATE_BreakdownForecastInfo(@WeekNumber := checkweeknumber, @WeekDate, @Supplier,
  @PartNumber, @SizeCode, @Qty1..7)` (`:1443-1469`).

---

## 3. The 862 (DELJIT firm) — report-only, NOT a forecast writer

**The 862 does not touch any forecast table.** It is handled entirely inline in `EDIUpload.Execute`
(`EDIUpload.pas:105-185`) and produces a **report Excel only**:
- Reads the remit/order date `copy(fcl,17,8)` → `yyyy/mm/dd` (`:112-113`).
- Opens `ReportTemplate.xls`, titles it **`'862 Firm Order'`** (`:124`), writes Order Date + headers
  Part Number / Qty / Prod Date (`:126-138`).
- Loop until `CTT` (`:142-174`): part number = `copy(fcl,9,12)`; advance to `SHP`, parse qty
  (`copy` up to next `*`) and prod date (skip 2 `*`, then `yyyymmdd`→`yyyy/mm/dd`); advance to `TD5`.
- Saves `FirmOrder<yyyymmddhhmmss00>.xls` to `fiReportsOutputDir`; logs `EDIIMP / EDI 862 Processed` (`:178-184`).

**Firm-vs-forecast split (Q11/§6):** in the legacy system the firm 862 is a **human-readable Excel** the
operator eyeballs against the planning 830 — there is **no firm/JIT column or table distinction in the DB**.
The 830 alone drives `INV_BREAKDOWN_FC_INF`; the Order reads that forecast. So the firm JIT signal exists only
as a side report today. **Rebuild decision needed (defer to architect):** whether to (a) keep 862 report-only,
or (b) actually persist DELJIT firm qtys (a firm vs planning flag/table) so the Order can prefer firm where present.
This spec records that **today the 862 is non-authoritative** — do not assume the order ever reads it.

---

## 4. Excel / COM dependency — what MUST be replaced server-side

The breakdown import uses Excel via `createOleObject('Excel.Application')` in three places, each a desktop-COM
dependency that does not exist on a gateway:
1. **Per-supplier forecast feed files** (`Execute`, `:435-547`) — `ForecastTemplate.xls` → per-supplier `.frc`
   text and/or `-Forecast` Excel, optionally archived (`fiLocalFTP`). The file format is fixed-width:
   `[SiteSupplierCode]SupplierCode + PartNumber + WeekDate + %.2d week + %.5d×7 day-qtys` (`:486-508`).
2. **Forecast report + DB-mismatch exception reports** (`ScanPartnumber`, `:802-945`) — `ReportTemplate.xls`
   → `ForecastReport*.xls`, `ForecastDBError*.xls`, `ForecastRecError*.xls`.
3. **862 firm-order report** (`EDIUpload.pas:119-182`) — `FirmOrder*.xls`.

**Side-effect usage rollup (not Excel, but server-side ETL to replace):** `UpdateUsage` (`:951-1029`) →
`SELECT_SizeUsage` then per-size `UPDATE_SizeUsage(@SizeCode, @Usage)`. The usage figure comes from
`HistoryForecast(part)` (`:1031-1078`) which averages `SELECT_ForecastPartNumberWeek(@WeekNo := WeekOfTheYear(now+z),
@DayNo := DayOfTheWeek(now+y), @PartNo)` over `fiUsageUpdateCompare` weeks × 7 days. **This read uses
`WeekOfTheYear` (raw ISO) WITHOUT the FirstProductionDay offset** (`:1052`) — inconsistent with the write
side's stored week (which is ISO−1). On real data the usage rollup reads from the *wrong* week bucket by 1.
**Confirm against golden:** `IN_USAGE` for a known size should reconcile to the stored `IN_QTYn` sum ÷ days;
if it's off, this offset bug is why. Flag for rebuild (the rebuild's usage query must use the same ISO−offset week as the stored row).

**"Unable to get month forecast in order" daily-log error → tie-back:** that Order-side failure is the symptom
of a **forecast gap** — either (a) the importer logged "No breakdown for part number(…)" (`:1180`, ratio
lookup failed, count silently dropped) so the week bucket is missing/zero, or (b) the week-number mismatch
above means the Order's `SELECT_ForecastPartNumberWeek`/`SELECT_ForecastDetailBCASN` reads an empty week.
**The rebuild's forecast importer must (1) not silently drop counts on a missing BOM ratio, and (2) write the
week number the Order reads (ISO−1), eliminating the gap that produces that order error.**

Replacement summary: all three Excel emissions → gateway report/export (named-query + report module or
Perspective download); the usage rollup and the whole parse → a server-side import service (no COM).

---

## 5. File source / DUNS routing / trigger (Q11)

- **Inbound drop:** `fiEDIIn` (INI `[DIRECTORIES] EDIIn`, default `c:\_Inventory_Control\EDIIn`, DataModule.dfm:513-516).
  An external TEMA VAN/mailer (not this app) deposits X12 files. EDIUpload **polls on demand** (manual button,
  blocking `while not fclosed … sleep(500)`) — **no scheduler today**.
- **DUNS → site routing:** `delSL[4]` (ISA element) → `AD_GetSiteTMMDUNS` on `ALC_Connection` (§1.2). Same
  allow-list mechanism the M1 inbound poller (997/824) uses — **reuse it** (`997-824-inbound-spec.md` §5.2).
- **Archive / move:** the 830 branch returns to `EDIUpload`, which then runs the shared archive block
  (`EDIUpload.pas:419-435`): `MoveFile := TRUE` into `fiEDIIn\Archive\`. The 830 has no parsed `EIN`, so the
  archive name falls to the `else` branch `<EDIFileNumber><delSL[10]>.EDI` = `830<ISA13>.EDI` (`:428`;
  `delSL[10]` = ISA13 interchange control #). Idempotency on re-poll = the file is moved out of the drop dir.
- **Trigger / Q11 (rebuild):** replace the manual blocking poll with a **per-site scheduled gateway job**
  (auto/manual mode per site), feeding the home-hub **Forecast Import box** and an **8-day staleness alarm**
  (alert if no successful 830 import in 8 days). DUNS no-match → **quarantine** the file (the spike already
  changed legacy's "silently leave in dir / re-scan forever" to quarantine; ignition-spike-log.md:577).
  **Q8 skip-by-config:** the per-config "skip" applies to the import enable/disable per site.

---

## 6. Hazards (first-class findings)

1. **`DELETE_ForecastInfo` — NOT missing (correction to prior analysis).** `forecast-breakdown.md:40,63`
   (written against the *superseded* 2026-06-01 snapshot) flagged this proc as absent → "latent runtime
   failure on every run." **In the authoritative live dump it EXISTS at `CreateInventory.sql:2725`** with the
   exact 3-param signature `@WeekDate, @HistWeekDate, @PartNumber` the form calls. The earlier flag was a
   snapshot-drift artifact. **Resolved: the delete path is live and works.** (Per `reference-schema-snapshot-vs-live`,
   CreateInventory.sql is authoritative; update forecast-breakdown.md §3.)
2. **Stored week MUST be raw FST09 (ISO−1), not offset (§1.5).** The literal-looking `+(FirstWeek−1)` at
   `:1374` does NOT apply to the stored row (`checkweeknumber` is written). Porting it onto the stored value
   drifts every row +1. Golden check: `2624` → stored `IN_WEEK_NUMBER = 24`.
3. **Replace-semantics race / additive double-count (§2.2).** Breakdown writer is additive + year-blind;
   correctness depends on the per-part window delete running first, in the same transaction. Re-import without
   delete doubles qtys; year-blind key collides across years.
4. **Year rollover / week 52→01 (D10 edge).** `IN_WEEK_NUMBER` is year-blind 1–52(/53). A horizon crossing
   the new year reuses low week numbers; the delete window is by `VC_WEEK_DATE` (string `yyyymmdd`, safe across
   years) but the upsert key is not. Confirm with a golden 830 whose horizon spans Dec→Jan.
5. **Usage rollup reads the un-offset ISO week (§4).** `HistoryForecast` uses `WeekOfTheYear` without the
   FirstProductionDay offset, mismatching the stored ISO−1 week → usage averaged from the wrong/empty bucket.
   Confirm `INV_SIZE_MST.IN_USAGE` reconciles to stored day-qtys.
6. **Ratio effective-month compare may be misaligned (§2.3).** `copy(ActiveDate,3,2)+copy(ActiveDate,6,2)`
   (ratio yy+mm) vs `copy(WeekDate,1,4)` (weekdate yyyy) — a 4-vs-4 char compare whose semantics aren't
   obviously equal. Confirm against a golden where the BOM has an effective-dated ratio (does the import pick
   the dated ratio or fall through to default?). Misalignment → wrong ratio silently used.
7. **Silent count drop on missing BOM (§4).** `bd=FALSE` → log + **ignore the count**. The forecast bucket
   ends up zero with no hard error → the Order's "Unable to get month forecast" downstream.
8. **Valve/film/label/misc use `wheelcount`, not own ratio (§2.3).** Faithful to source; verify intended.
9. **Cross-DB coupling.** `AD_GetSpecialDateWeek` + `AD_GetSiteTMMDUNS` live in the ALC/VehicleOrder DB
   (`ALC_Connection`). Bodies unverified here. The forecast day-spread depends on the production calendar in
   the *other* DB — the rebuild must keep that join (or replicate the calendar).
10. **DUNS ISA04 element index (§1.2).** Same hazard as M1 — confirm `delSL[4]` against the golden inbound ISA.

---

## 7. What the rebuild's forecast importer MUST reproduce

1. **Parse:** X12 DELFOR walk — `LIN` → assembly part (LIN03) + kanban (LIN05); `FST` → qty (FST01), start
   date (FST04, `yyyymmdd`), week from **FST09 "DO" chars 3–4** (TEMA-supplied, production-relative). Supplier
   = configured operator supplier, not the file. Stop at `CTT`.
2. **Week number = raw FST09 (ISO−1), stored unmodified.** Apply the FirstProductionDay offset ONLY to the
   holiday calendar lookup, never to `IN_WEEK_NUMBER`.
3. **Tables:** explode each assembly via the `INV_FORECAST_DETAIL_INF` BOM (ratios + component part codes +
   `VC_BROADCAST_CODE` pattern — config, not imported); write raw forecast to `INV_FORECAST_INF` (overwrite key
   supplier/part/week) and day-bucket qtys to `INV_BREAKDOWN_FC_INF` (the table the Order + ASN read).
4. **Replace semantics:** per-part **window delete** (future horizon `>= firstWeekDate`, prune history
   `<= now−HistoricalForecast*7d`) **before** the additive accumulate, in one transaction; add **year** to the
   breakdown key for multi-year safety.
5. **Day-spread:** week qty ÷ working days (Mon–Fri minus H/X calendar days from the production calendar),
   leftover on the first working day → `IN_QTY1..7`.
6. **Usage rollup:** recompute `INV_SIZE_MST.IN_USAGE` reading the **same ISO−offset week** as stored (fix the
   `:1052` mismatch).
7. **862:** keep report-only (or escalate to architect for a firm-vs-planning distinction); the 862 is NOT a
   forecast-table writer today.
8. **No Excel/COM:** all feed files + reports → server-side export; whole parse → import service.
9. **Trigger (Q11):** per-site scheduled poll (auto/manual), home-hub Forecast Import box, 8-day staleness
   alarm, DUNS-no-match quarantine, archive moved file.

---

## Cross-references
- `docs/analysis/decisions.md` D10 (week numbering, validated against golden `EDI/830000008976.EDI`)
- `docs/analysis/forecasting/forecast-breakdown.md` (the breakdown processor — **update §3/§40/§63: `DELETE_ForecastInfo` now confirmed present at schema:2725**)
- `docs/analysis/forecasting/forecast-detail.md` (the BOM/ratio master + `VC_BROADCAST_CODE`)
- `docs/analysis/edi/inbound/997-824-inbound-spec.md` §5 (shared poller / DUNS routing / archive — reuse for 830)
- `docs/analysis/edi/edi-upload.md` (full inbound dispatch overview)
- ASN keystone / Order specs (the READ side: `SELECT_ForecastDetailBCASN`, `SELECT_ForecastPartNumberWeek`)
