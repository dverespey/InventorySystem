# EDI 830 forecast import — the EXACT extracted algorithm (STEP-0, source-of-truth)

> **M2 foundational unit.** This is the line-by-line extraction of the legacy importer
> `ForecastBreakdownF.pas` (LIVE, `InventorySystem.dpr:28`) — the **exact** 830 parse, the
> assembly→component **explode**, the weekly→7-day **day-spread**, and the **write order** the rebuild
> must reproduce. Everything here is quoted to a `.pas` line or a proc body in `/tmp/inv_utf8.sql`
> (= `DB Schema/CreateInventory.sql`, the authoritative 6/12 live dump). Where the source is genuinely
> ambiguous it is flagged `AMBIGUITY` with the exact lines — not guessed.
>
> Companion specs (read alongside): `830-862-forecast-import-spec.md` (behavioral), `forecast-tables-analysis.md`
> (the two tables, proven on live), `decisions.md` D10 (week numbering, golden-validated).

---

## A. The 830 parse (`ForecastBreakdownF.pas:184-298`)

The X12 DELFOR walk. The legacy reads the file line-by-line (`Readln`), one X12 segment per physical
line (TEMA writes it that way — there is **no `~` segment-terminator split**). Elements are obtained by
`splitString('*', line, delSL)` (`:127-147`), which yields `delSL[0]='LIN'/'FST'/'ISA'`, `delSL[1]`=the
first element, etc. (1-based after the segment id).

### A.1 Envelope sniff + DUNS guard (`:184-217`)
- Line 1 must contain `'ISA'` (`pos('ISA',fcl)>0`, `:185`). Split on `*` → **`delSL[4]`** is the
  trading-partner DUNS (`:188`). Validated via `SiteTMMDUNSDataset(@SiteTMMDUNS:=delSL[4])` → cross-DB
  `AD_GetSiteTMMDUNS` on `ALC_Connection`. `RecordCount=0` ⇒ `exit` (`:196-201`).
- Reads to **line 3**, `data:=copy(fcl,4,3)` (chars 4-6 of the ST segment) must equal `'830'` else `exit`
  (`:203-213`). On match `EDIfile:=TRUE`, then `Reset(fcf)` to re-read from the top.
- **delSL[4] element-index hazard:** carried verbatim from M1. On our own outbound ISA delSL[4]=ISA04=
  our SiteDUNS; the real inbound element is PENDING a captured golden TEMA 830 (the D10 golden
  `EDI/830000008976.EDI` is gitignored client data). The rebuild reuses the M1 `_resolve_site_by_duns`
  (matches `INV_SITES.VC_TMM_DUNS`), which is legacy-faithful in COLUMN semantics; the X12 element index
  is the same open gap as 856/810/997/824.

### A.2 Count pass + the per-LIN segment walk (`:221-298`)
- **Count pass** (`:221-233`): count `LIN` segments → `count`; `SetLength(fEntries,count)`.
- **Per-LIN loop** (`:245-298`): advance until `data='LIN'` (or `data='CTT'` = end-of-transaction →
  **break**). For each LIN segment (`splitString` → delSL):
  - `fEntries[].Supplier := Data_Module.fiSupplierCode.AsString` — **the operator's configured supplier,
    NOT from the file** (`:262`). (Rebuild: per-site supplier config, threaded in.)
  - `fEntries[].Partnumber := delSL[3]` — **LIN03 = the assembly / broadcast (BC) part number** (`:263`).
  - `fEntries[].KanbanNumber := delSL[5]` — **LIN05** (`:264`). `Skip := False`.
  - Advance to the first `FST` (`:268-272`).
  - **FST loop** (`:276-295`) — one `TWeekData` per FST segment, accumulated into `Weeks[1..14]`:
    - **`WeekNumber := StrToInt(copy(delSL[9],3,2))`** — **chars 3-4 of FST09** (the "DO" reference).
      For `FST*144*D*W*20260615*20260619***DO*2624` → `delSL[9]="2624"`, `copy(.,3,2)="24"` → week **24**.
    - **`WeekDate := delSL[4]`** — **FST04** = the forecast start date `yyyymmdd` (8 chars).
    - **`WeekCount := StrToInt(delSL[1])`** — **FST01** = the bucket forecast quantity.
    - The **first** FST sets module-level `fFirstWeekDate:=WeekDate` / `fFirstWeekNumber:=WeekNumber`
      (`:283-287`) — the **delete-window anchor** (§D).
  - `INC(counter)` per LIN.

> **Non-EDI fixed-width path (`ScanLine`, :592-721):** an alternate operator file format keyed by
> fixed offsets (`SUPPLIER_OFF=1` … `FORECASTW14_OFF=204`, 14-char week cells of `WW`+`yyyymm`(6)+`count`(6)).
> Used when line 1 is NOT an ISA. The 14th week is read only when `fiAssemblerName='WQS'` (`:700-708`).
> **The rebuild's 830 importer targets the EDI/830 path; the fixed-width path is a legacy operator-feed
> variant (out of M2 scope — note as a follow-on if a `.frc`-style manual feed must be re-supported).**

---

## B. The week-number derivation (D10 — production-relative, ISO−1; store RAW)

**`WeekNumber = chars 3-4 of FST09 ("DO" ref).`** The DO ref is `2`+`6`(year 2026)+`WW`. Measured across
the whole golden horizon, `ISO_week(FST04) − TEMA_DO_week = 1` for every normal production week
(decisions.md D10; `forecast-tables-analysis.md` §4). The stored week is **production-relative = ISO−1**
for 2026 (offset = `INT_FIRST_PRODUCTION_WEEK[2026]−1 = 2−1 = 1`).

**THE R1 CATCH — store the RAW FST09 week unmodified.** In `DoPartNumberForecast` (`:1314-1480`):
- `checkweeknumber := WeekNumber` (`:1358`) — captures the **raw** TEMA week.
- If `fiUseFirstProductionDay` (`:1362`), it reads `SELECT_FirstProductionDay(@ProdYear:=copy(WeekDate,1,4))`
  and, when `First Week Number ≠ 1`, does **`WeekNumber := WeekNumber + FirstWeekNumber − 1`** (`:1374`)
  — **mutating ONLY the LOCAL `WeekNumber`**.
- The breakdown INSERT writes **`@WeekNumber := checkweeknumber`** (`:1446`) — the **raw, un-offset** week.
  The offset-adjusted `WeekNumber` is used ONLY for the holiday lookup `AD_GetSpecialDateWeek(@Week:=WeekNumber)`
  (`:1390`).

> **D10 / R1 (data-adjudicated twice):** the stored `IN_WEEK_NUMBER` MUST be the raw FST09 week. For
> `2624` it is **24**, NOT 25. A rebuild that ports `:1374` onto the stored value drifts every row +1 and
> the Order read silently mismatches ("Unable to get month forecast"). **The offset is for the holiday
> lookup ONLY.** (`INV_FIRST_PRODUCTION_DAY[2026]` confirmed on spike = `2026, 20260105, 2`.)

---

## C. The explode (assembly → component qty) (`UpdateForecast`, :1081-1312)

Outer loop: for each non-skipped entry `i`, for each week `j` in `1..count` where
**`count = 14 if fiAssemblerName='WQS' else 13`** (`:1089-1092`):

1. **Raw forecast write** — `INSERTUPDATE_ForecastInfo(@Supplier, @PartNumber=assembly, @Kanban,
   @WeekNumber=Weeks[j].WeekNumber (raw), @WeekDate=Weeks[j].WeekDate, @Count=Weeks[j].WeekCount)`
   (`:1104-1118`). REPLACE upsert keyed `(supplier, part, week)` — overwrites IN_COUNT + week-date.

2. **Read the BOM recipe** — `SELECT_ForecastDetail(@AssyCode := assembly, @ForecastNotZero := 1)`
   (`:1126-1133`). Returns the recipe row(s): tire/wheel/valve/film/label/misc1/misc2 part codes + the
   three ratios + `VC_BROADCAST_CODE`. Aliases (from the proc body, `/tmp/inv_utf8.sql:3154`):
   `'Tire Ratio'`(=IN_TIRE_RATIO), `'Wheel Ratio'`(=IN_WHEEL_RATIO), `'Forecast Ratio'`(=IN_RATIO),
   `'Tire Part Number Code'`, `'Wheel Part Number Code'`, `'Valve Part Number'`, `'Film Part Number'`,
   `'Label Part Number'`, `'Misc1 Part Number'`, `'Misc2 Part Number'`, `'Active Date'`(=VC_EFFECTIVE_MONTH).

3. **Pick the ratio set by effective month** (`:1151-1176`):
   - `Active Date` blank or `' '` → the **default** ratio (`bd:=TRUE`).
   - else `tm := copy(ActiveDate,3,2)+copy(ActiveDate,6,2)` (ratio yy+mm), `wm := copy(WeekDate,1,4)`
     (weekdate yyyy); `if tm = wm` → match → take that ratio, `bd:=TRUE`, **break**.
   - if neither matched → `bd=FALSE`.
   > **AMBIGUITY (low impact, ratio effective-month compare) — `:1163-1165`.** `tm` is a 4-char yy+mm
   > (`copy(.,3,2)`=chars 3-4 of `yyyy/mm` = `yy`; `copy(.,6,2)`=chars 6-7 = `mm`); `wm` is a 4-char
   > yyyy (`copy(WeekDate,1,4)`). These are **not obviously aligned** (yymm vs yyyy). On live data the
   > comparison is **dead**: ALL 50 recipe rows have `VC_EFFECTIVE_MONTH = ' '` (a single space;
   > `forecast-tables-analysis.md` §1.4), so the default branch ALWAYS fires and the dated branch is
   > never reached. **Faithful behavior: default ratio always.** The dated branch is reproduced for
   > fidelity but exercised only by a synthetic dated-recipe fixture; its real semantics need a golden
   > where a recipe carries a `yyyy/mm` effective month. (Recorded, not guessed.)

4. **Silent drop on `bd=FALSE`** (`:1178-1183`): logs `"No breakdown for part number(…) … count will be
   ignored"` and the count is **dropped** (no components written). **THE FIX (D-Bug-1, §F):** the rebuild
   writes a forecast-gap ALARM row instead of vanishing the count.

5. **Compute counts** (`:1186-1226`):
   ```
   if (forecastratio<>0) and (tireratio<>0) and (wheelratio<>0):
       tirecount  = (((WeekCount * forecastratio) div 100) * tireratio)  div 100
       wheelcount = (((WeekCount * forecastratio) div 100) * wheelratio) div 100
   else: tirecount = wheelcount = 0      # any zero ratio -> zero
   ```
   **Integer division at each `div 100`** (truncating). The commented-out block `:1206-1225` (the old
   share-split math) is dead. This is the **share-fanout-faithful** form: TEMA sends the FULL week count
   on each assembly LIN of a shared BC, and the explode scales each by its own `tireratio/100` — e.g.
   `[KLM]CC` (3 assemblies, tire ratios 40/20/40) yields 40%/20%/40% of the count each, summing 100%
   across the share (proven on spike; recipe comment `:1191-1199`).

6. **Fan to components** (`:1230-1287`): for each component column whose code length > 2,
   `DoPartNumberForecast(code, WeekDate, <count>, WeekNumber)`:
   - **Tire** → `tirecount`.
   - **Wheel / Valve / Film / Label / Misc1 / Misc2** → **all use `wheelcount`** (NOT their own ratio).
   > **D-faithful, flagged (`:1248-1287`):** valve/film/label/misc all carry `wheelcount`. There is no
   > per-component ratio in the math. Faithful to source; verify intended with `delphi-architect` if a
   > rebuild ever wants per-component ratios (out of M2 scope — reproduce as-is).

---

## D. The day-spread (week qty → IN_QTY1..7) (`DoPartNumberForecast`, :1314-1480)

Given `(PN=component, WeekDate, FCCount, WeekNumber)`:

1. **`checkweeknumber := WeekNumber`** (`:1358`) — the raw week, what gets stored (§B).
2. **Part master read** — `SELECT_PartsStockInfo(@PartNum := PN)` (`:1338-1351`):
   `line := 'Line Name'` (or `'ALL LINES'` if blank), `supplier := 'Supplier Code'`, `size := 'Size Code'`.
   So the breakdown row's supplier/size come from the **component's** part master, **not** the feed.
3. **Default workdays Mon-Fri** (`:1321-1328`): `workday[1..5]=true`, `workday[6]=workday[7]=false`,
   **`days := 5`**.
4. **Offset-adjust the holiday-lookup week ONLY** (`:1360-1377`): if `fiUseFirstProductionDay` and
   `FirstWeekNumber ≠ 1`, `WeekNumber := WeekNumber + FirstWeekNumber − 1` (the LOCAL var; never stored).
5. **Apply the production calendar** — `AD_GetSpecialDateWeek(@Week := WeekNumber, @Line := line)` on
   `ALC_Connection` (cross-DB, VehicleOrder; body read on spike). For each returned row (`:1394-1410`):
   - if `trim('Date Status Abrv') in {'H','X'}` → `workday['Day Number'] := False; DEC(days)`
   - **else** → `workday['Day Number'] := True; INC(days)`.
   > **AMBIGUITY — the `INC(days)` else-branch is a LATENT LEGACY BUG (`:1403-1407`).** For a special-date
   > row that is NOT H/X (e.g. an overtime/extra-production 'O'/'P' day, or a Saturday turned ON), the
   > legacy sets `workday[N]:=True` AND `INC(days)` **unconditionally** — so a day that was ALREADY a
   > workday (Mon-Fri, days started at 5) gets **double-counted** in `days`, while a day turned-OFF then
   > back-ON could be mis-tracked. On the spike VehicleOrder `SpecialDate` data the special dates are
   > overwhelmingly H/X (holidays/shutdowns), so the else-branch is rarely hit and the live day-spread is
   > stable; but the arithmetic is not provably correct for an overtime row. **The rebuild reproduces the
   > H/X-turns-off behavior faithfully (the load-bearing, data-confirmed path) and, for the else-branch,
   > sets `workday[N]:=True` only if it was previously off (no double-count) — a documented divergence
   > (§F D-Bug-3), NOT a silent port of the buggy `INC`.** The `'Day Number'` is `DATEPART(DW, date +
   > @@DATEFIRST - 1)` (1=Mon..7=Sun under the proc's @@DATEFIRST math); the rebuild keeps the proc as the
   > authority for the day index (does not recompute it).
6. **Spread** (`:1414-1434`):
   ```
   if days > 0:  ratiocount = FCCount div days;  leftover = FCCount mod days
   else:         ratiocount = 0;                 leftover = 0
   for i in 1..7:
       if workday[i]:  dayforecast[i] = ratiocount + leftover;  leftover = 0   # ALL remainder -> first working day
       else:           dayforecast[i] = 0
   ```
   So: even split across working days, **the entire remainder lands on the first working day**; non-working
   days get 0. (Live-proven: `20260908` part `4265202R6000`, QTY1..7 = `31,27,27,27,27,0,0` → 138 over 5
   days = 27 r3… actually 31=27+4 ⇒ remainder 4 on day 1; `forecast-tables-analysis.md` §2.3.)
7. **Write** — `INSERTUPDATE_BreakdownForecastInfo(@WeekNumber := checkweeknumber (RAW), @WeekDate,
   @Supplier, @PartNumber := PN, @SizeCode := size, @Qty1..7 := dayforecast[1..7])` (`:1443-1469`).
   **ADDITIVE upsert** keyed `(supplier, part, week)`, year-blind: on exists, `IN_QTYn = IN_QTYn + @Qtyn`;
   `VC_WEEK_DATE`/`VC_SIZE_CODE` written on INSERT only (`/tmp/inv_utf8.sql:1215`).

---

## E. The write order (delete FIRST, then additive accumulate)

`Execute` (`:320-335`), after `ScanPartnumber` validates each assembly has a recipe:

1. **DELETE FIRST** (`:322-329`): for **every non-skipped entry**, `DeleteBreakdown(part)` →
   `DELETE_ForecastInfo(@WeekDate := fFirstWeekDate, @HistWeekDate := fHistDate, @PartNumber := assembly)`
   (`:110-125`). `fHistDate := now − fiHistoricalForecast*7 days` (`:318`; INI default 12 → 84 days).
   The proc (`/tmp/inv_utf8.sql:2725`) does **four deletes**, resolving assembly→its 7 component codes via
   a `CROSS APPLY (VALUES …)` over `INV_FORECAST_DETAIL_INF` where `VC_ASSY_PART_NUMBER_CODE=@PartNumber`:
   - (A) `INV_BREAKDOWN_FC_INF` for those components WHERE `VC_WEEK_DATE >= @WeekDate` (forward slice).
   - (B) same components WHERE `VC_WEEK_DATE <= @HistWeekDate` (prune history).
   - (C) `INV_FORECAST_INF` WHERE `VC_PART_NUMBER=@PartNumber AND VC_WEEK_DATE >= @WeekDate`.
   - (D) `INV_FORECAST_INF` WHERE `VC_PART_NUMBER=@PartNumber AND VC_WEEK_DATE <= @HistWeekDate`.
   Boundaries inclusive (`>=`, `<=`); string compare on `yyyymmdd` (lexical=chronological). **Trims both
   ends, keeps only the middle window `(HistWeekDate, WeekDate)`.**
2. **THEN re-explode + additive insert** — `UpdateForecast` (§C/§D).
3. Then `UpdateUsage` (§G) and the per-supplier file emit (Excel/COM — replaced server-side).

> **WHY ORDER IS LOAD-BEARING.** The breakdown upsert is **additive**. Re-importing the SAME 830 without
> the delete would **double** every qty (and a week-30 row would accumulate onto a leftover week-30 of a
> prior year — the key is year-blind). **The delete window IS the idempotency guard.** The rebuild MUST
> run the per-assembly delete BEFORE the additive accumulate, in the SAME transaction (or switch the
> breakdown writer to overwrite-by-key). This is the M2 analogue of GALC's proc-side dedup.
>
> **T1 — delete-by-component crux.** The breakdown is keyed by COMPONENT; the delete is BY ASSEMBLY, so
> `DELETE_ForecastInfo` resolves assembly→components via the recipe. Deleting breakdown by assembly code
> directly removes **0 rows** (0/959 breakdown rows hold an assembly code). The recipe MUST be intact at
> delete time (delete the recipe first and the CROSS APPLY finds nothing).

---

## F. The bugs FIXED in the rebuild (documented divergences from source)

| # | Legacy behavior (source) | Rebuild divergence | Why |
|---|---|---|---|
| **D-Bug-1** | `bd=FALSE` (no BOM ratio match) → log + **silently drop** the count (`:1178-1183`). The week bucket ends up zero → the Order's "Unable to get month forecast". | Write a **forecast-gap alarm** row (`INV_EDI_ALARM_REJ`, type `830_FORECAST_GAP`, manifest=NULL, part=assembly, errorText=the missing-BOM reason) + log; the count is still not exploded (no recipe to explode by) but it is now **visible, not vanished**. | A missing recipe is operationally actionable; silent zero forecast is the root of the daily-log order error (spec §4). |
| **D-Bug-2** | Usage rollup `HistoryForecast` reads `WeekOfTheYear(now)` = **raw ISO** with NO offset (`:1052`), mismatching the stored ISO−1 week → averages from the wrong/empty bucket. | The rebuild's usage rollup reads the **same production-relative week** the row is stored under (`ISO − offset`), so usage reconciles to the stored day-qtys. | Read/write the SAME week consistently — eliminates the "Unable to get month forecast" offset gap (spec §4, hazard #5). |
| **D-Bug-3** | Special-date else-branch (`:1403-1407`) does `workday[N]:=True; INC(days)` **unconditionally** → a NON-H/X special row double-counts an already-on workday in `days`. | Set `workday[N]:=True` only if it was previously OFF (and only then `INC(days)`); H/X turns OFF as in source. | Prevents a spurious `days` inflation that would under-spread the qty. H/X path (the live-dominant case) is byte-faithful. |
| **D-Bug-4 (latent)** | `IN_QTY1..7` are nullable; additive `IN_QTYn = IN_QTYn + @Qtyn` → `NULL + n = NULL` (NULL poison). Live has 0 NULL rows so latent. | The rebuild always writes **0, never NULL** on INSERT (the proc already does), and the delete-then-additive cycle never leaves a partial row. (No code change to the proc; honored by always-non-NULL writes.) | Defends the additive math against a hand-NULLed row (T3). |
| **D-Bug-5 (year-blind, deferred)** | `IN_WEEK_NUMBER` + the upsert key carry **no year** (week 30/2026 collides with 30/2027). Safe today only via delete-forward. | **Deferred to M4 schema** (add year or FST04-month to the breakdown key) — documented, NOT silently dropped. The rebuild keeps the delete-forward discipline (the legacy guard) for now. | Multi-year/multi-site safety needs a key change; in scope of M4 re-key, not M2. |

---

## G. The usage rollup (`UpdateUsage` :951-1029 + `HistoryForecast` :1031-1078) — DATA side

`SELECT_SizeUsage` (`/tmp/inv_utf8.sql:5089`) returns `(vc_size_code, in_usage, vc_part_number,
vc_kanban_number)` joined size→part. Per size, sum `HistoryForecast(part)` over its parts; per size,
`UPDATE_SizeUsage(@SizeCode, @Usage)` (`:5311`) sets `INV_SIZE_MST.IN_USAGE`. `HistoryForecast` averages
`SELECT_ForecastPartNumberWeek(@WeekNo, @DayNo, @PartNo)` over `fiUsageUpdateCompare` weeks × 7 days
(`:1043-1064`), `result = total div count`. **D-Bug-2 applies here** (the `WeekOfTheYear(now)` un-offset
read). The rebuild's usage step reads the production-relative week. (The Excel/COM file emits in `Execute`
are replaced by a server-side export — out of M2 logic scope; noted in spec §4.)

---

## H. What M2 builds (mapping this algorithm to the rebuild)

- **PURE** `forecast/code.py`: `parse_830(text)` (§A), `explode(assemblyRows, recipe)` (§C),
  `day_spread(weeklyQty, weekDate, calendarOffDays, ...)` (§D). No I/O; CPython/Jython-portable.
- **DRIVER** `import_830(...)`: DUNS guard + idempotency ledger (reuse M1); per assembly **DELETE FIRST**
  (§E) then explode→day-spread→`INSERTUPDATE_BreakdownForecastInfo` (additive) +
  `INSERTUPDATE_ForecastInfo` (replace). One tx per file. Fixes D-Bug-1..3 (§F).
- **Q11 data side:** `INV_SITES.VC_FORECAST_IMPORT_MODE` (AUTO/MANUAL — already on the table),
  `VC_LAST_FORECAST_IMPORT` (stamped per successful import), an **8-day staleness** check
  (`stale_sites()`) raising a `830_FORECAST_STALE` alarm. The gateway scheduled poll + the home-hub box
  are prod wiring (follow-ons).
