# Adversary findings — 830 forecast importer vs legacy `ForecastBreakdownF.pas`

**Reviewer stance:** the reimplementation is wrong until proven equivalent. Equivalence is only proven
against the LEGACY computation on the SAME inputs — not against the rebuild's own recomputation.

**Targets reviewed**
- Rebuild: `docs/analysis/edi/inbound/project-library/forecast/code.py` (PURE `parse_830`/`explode`/
  `day_spread`/`pick_ratio` + the `import_830` driver).
- Legacy: `ForecastBreakdownF.pas` — parse `:245-298`, explode `UpdateForecast :1081-1312`, day-spread
  `DoPartNumberForecast :1314-1480`, write-order `Execute :320-335`, `ScanPartnumber :724-797`,
  `DeleteBreakdown :110-125`.
- Procs (live, read from `mssql-spike`): `INSERTUPDATE_BreakdownForecastInfo`, `INSERTUPDATE_ForecastInfo`,
  `DELETE_ForecastInfo`, `SELECT_PartsStockInfo`; cross-DB `VehicleOrder.dbo.AD_GetSpecialDateWeek`.
- Tests: `scripts/e2e/test_forecast_import_build.py` (35/35 PASS), `test_forecast_import_e2e.py`
  (24/24 PASS on spike, restores as-found).

**Honesty gate (carried from the specs, re-affirmed):** there is NO captured golden TEMA 830 on disk, so
byte/offset parity vs a real feed is UNPROVABLE. All findings below are proven against the legacy `.pas`
algorithm + the live proc bodies + live data — NOT against a golden file.

---

## REQUESTED RESULTS (the three numbers + verdict up front)

### 1. EXPLODE-truncation result — MATCHES (no divergence on non-negative inputs)
The rebuild `tire_count = ((wc * fr) // 100) * tr // 100` (code.py:217) is **byte-identical** to the legacy
`tirecount := ((((WeekCount)*forecastratio) div 100)*tireratio) div 100` (`ForecastBreakdownF.pas:1202`)
for ALL non-negative `(wc, fr, tr)` — both truncate at EACH `/100`. Exhaustively checked
`wc∈[0,500), fr∈[1,100], tr∈[1,100]`: identical everywhere. The double-truncation ORDER is preserved: the
rebuild does NOT collapse to the naive `wc*fr*tr//10000` (which DIVERGES — e.g. `wc=3, fr=60, tr=66`:
legacy/rebuild = **0**, naive = **1**). Zero-ratio → 0 guard reproduced. 3-way share 40/20/40 of 1000 →
400/200/400 (sum 100%) reproduced. Valve/film/label/misc all use `wheelcount` — faithful.
**One latent edge (NIT, NEG-DIV):** a NEGATIVE `WeekCount` would diverge — Pascal `div` truncates toward
zero, Python `//` floors. `wc=-7, fr=100, tr=40`: Pascal `-2`, Python `-3`. Requires a negative FST01
(not seen on real feeds). Not weaponizable without a malformed golden.

### 2. DAY-SPREAD result — DIVERGES on real calendar data (weekend overtime)
H/X-only weeks (the live-dominant path): **byte-identical** — `138` over Mon-Fri →
`[30,27,27,27,27,0,0]`; a Wed-`H` → `[36,34,0,34,34,0,0]`; a Mon-`H` → `[0,36,34,34,34,0,0]`. All match.
BUT a NON-H/X special row (status `O`/overtime, status `P`) **diverges**, and the live VehicleOrder
calendar CONTAINS such rows (6 `O` Saturdays on COROLLA — see BLOCKER-1). Counterexample on REAL data
(COROLLA, raw week 22 → lookup week 23 = Saturday 20260606 `O`, qty 138):
- **Legacy** turns Saturday ON (`workday[6]:=True; INC(days)` → days=6) → `[23,23,23,23,23,23,0]`.
- **Rebuild** ignores the `O` row → days=5 → `[30,27,27,27,27,0,0]`.

Every one of the 7 day buckets differs. (Details + the documented-fix contradiction in BLOCKER-1.)

### 3. DELETE-window result — MATCHES (trim both ends, keep middle); idempotency holds
Proven on the spike inside a rolled-back transaction (0 rows leftover after rollback). Component rows at
`VC_WEEK_DATE` `20260101`(history), `20260501`(middle), `20260615`(forward); `DELETE_ForecastInfo
@WeekDate='20260615', @HistWeekDate='20260201', @PartNumber='<assembly>'`:
- history `20260101` (≤ `20260201`): DELETED. forward `20260615` (≥ `20260615`, inclusive): DELETED.
  middle `20260501`: SURVIVES.
The rebuild calls the SAME proc with the SAME params (code.py:545-547), so the window + the
assembly→component CROSS APPLY resolution are faithful. The additive breakdown upsert is a REPLACE only
because the delete clears the forward slice first; re-import-not-doubled confirmed (e2e case 3 + 7).
**One delete-vs-rewrite supplier interaction makes BLOCKER-1 worse** (the delete has no supplier filter;
the re-insert lands under the wrong supplier) — see BLOCKER-1.

### VERDICT (also at the very bottom)
**NOT proven equivalent. Two BLOCKER divergences that change stored numbers on REAL inputs:**
(B1) the breakdown `VC_SUPPLIER_CODE` source, and (B2) the weekend-overtime day-spread. The explode math,
the delete window, D10 raw-week, and idempotency ARE faithful. The remaining items are SHOULD-FIX /
NIT / data-vintage.

---

## BLOCKER-1 — Breakdown rows written under the WRONG supplier (feed supplier, not the component's)

**Claim under test (rebuild):** the breakdown `VC_SUPPLIER_CODE` is the per-site feed supplier
(`import_830(..., supplier=...)`), with `rowSup = rowSupplier or compSupplier` (code.py:583, 590).

**Legacy reality (`DoPartNumberForecast:1349`):** `supplier := FieldByName('Supplier Code').AsString` —
the breakdown row's supplier is read from the **COMPONENT's part master** (`SELECT_PartsStockInfo` →
`INV_SUPPLIER_MST.VC_SUPPLIER_CODE` via `INV_PARTS_STOCK_MST.IN_SUPPLIER_ID`), NOT from the feed. The feed
supplier (`fiSupplierCode`) is used ONLY for the RAW forecast `INSERTUPDATE_ForecastInfo` (`:1107`), never
for the breakdown.

**Counterexample — proven on live `Inventory`:**
```
INV_FORECAST_INF (raw)        : 1 distinct supplier  = 71930 (the feed supplier) across all 1041 rows
INV_BREAKDOWN_FC_INF          : 14 distinct suppliers, e.g. 72100/11111/07100/0572B/07451/07100...
breakdown VC_SUPPLIER_CODE == component part-master supplier : 959 / 959  (0 mismatches)
breakdown rows whose supplier != the feed supplier 71930     : 959 / 959  (ALL of them)
```
So the legacy stores EVERY breakdown row under its component's own supplier; **none** carry the feed
supplier. The rebuild, in the prod path (`_process_one_830` → `supplier = supplierBySite.get(siteId)`,
code.py:767-769), passes a single non-empty site supplier, so `rowSup = rowSupplier` for every row →
**all breakdown rows get the feed supplier, 100% wrong vs legacy.**

**Impact (number-changing):**
- `VC_SUPPLIER_CODE` differs on every breakdown row.
- The upsert dedup key is `(VC_SUPPLIER_CODE, VC_PART_NUMBER, IN_WEEK_NUMBER)`
  (`INSERTUPDATE_BreakdownForecastInfo` EXISTS check, verified live) → a DIFFERENT KEY per row.
- `DELETE_ForecastInfo` has **NO supplier filter** (verified: deletes by component+weekdate only). So on a
  re-import the rebuild's delete removes the rows, then re-inserts them under the feed supplier — silently
  REWRITING the supplier on the whole forward horizon away from the component's true supplier.
- `SELECT_ForecastSupplier` (the per-supplier file emit, ordered by supplier) and any supplier-scoped
  consumer would now group/route every component under one supplier → wrong supplier feed files.
  (The Order read `SELECT_ForecastPartNumberWeek` is supplier-blind, so the day qty itself still reads —
  which is exactly why this would ship silently.)

**Why the tests miss it:** the e2e fixture parts have a NULL part-master supplier and `import_830` is
called with `supplier="ZZF83"`, so `rowSup = "ZZF83" or None = "ZZF83"` — the path where the component HAS
a supplier differing from the feed is never exercised. The PURE test never touches the driver's supplier
decision.

**Classification:** CODE DEFECT. Fix = for the BREAKDOWN write, prefer the component part-master supplier
(`compSupplier`), using the feed supplier only for the RAW forecast (mirror the legacy's two distinct
sources). `file:line` = code.py:583 (`rowSup = rowSupplier or compSupplier` should be
`compSupplier or rowSupplier` for the breakdown) vs `ForecastBreakdownF.pas:1349`.

---

## BLOCKER-2 — Weekend-overtime day-spread diverges on REAL VehicleOrder calendar data

**Claim under test (rebuild, D-Bug-3 §F):** "the rebuild reproduces the H/X-turns-off behavior faithfully
and, for the else-branch, sets `workday[N]:=True` only if it was previously off (no double-count)."

**Actual rebuild code (`_calendar_off_days`, code.py:453-462):** only `H`/`X` rows add to the OFF set; a
NON-H/X row (`O`, `P`, …) is **completely ignored — it never turns any day ON.** The comment at code.py:461
("do NOT add to off (D-Bug-3 — never inflate `days`)") is the actual behavior; the §F prose ("set
`workday[N]:=True` only if previously off") is NOT implemented. These two statements disagree for a day
that was previously OFF (Sat/Sun).

**Legacy (`ForecastBreakdownF.pas:1403-1407`):** a non-H/X row does `workday[N]:=True; INC(days)`
unconditionally — so a Saturday (`workday[6]` started false) is turned ON and `days` goes 5→6.

**Counterexample — proven on REAL `VehicleOrder` data:**
```
ProductionStatusAbv values present on SpecialDate: H (11 rows), O (6 rows)   -- NO 'X' at all
All 6 'O' rows are SATURDAY (Day Number=6) on the COROLLA line:
  20260328 (iso13), 20260411 (iso15), 20260425 (iso17), 20260516 (iso20), 20260606 (iso23), 20260725 (iso30)
AD_GetSpecialDateWeek @Week=23, @LineName='COROLLA'  ->  Day Number=6, Date Status Abrv='O'
```
For an assembly on COROLLA, raw week 22 → lookupWeek 22+offset(1)=23, component count 138:
| | days | IN_QTY1..7 |
|---|---|---|
| **Legacy** (Sat ON) | 6 | `[23,23,23,23,23,23,0]` |
| **Rebuild** (Sat off) | 5 | `[30,27,27,27,27,0,0]` |

All seven buckets differ; the Order would read a different daily forecast for every COROLLA component in
ISO weeks 13/15/17/20/23/30 (and any future `O` Saturday). The day-number mapping is confirmed aligned
(`DATEPART(DW, date + @@DATEFIRST(=7) - 1)`: Mon=1 … Sat=6 … Sun=7 = the rebuild's `workday[1..7]`).

**Also: the rebuild contradicts its OWN documented fix.** Per §F D-Bug-3 ("True only if previously off"),
Saturday WAS off → the documented fix would turn it ON → days=6 → it would MATCH the legacy here. The
shipped code leaves Saturday off (days=5), so it matches NEITHER the legacy NOR the documented intent for
weekend overtime. (For an already-ON weekday `O` row, the legacy double-counts days and the rebuild does
not — there the rebuild matches the documented intent but still diverges from the legacy. Either way the
stored numbers differ from the legacy on a row that exists in live data.)

**Why the tests miss it:** e2e case 5 deliberately uses the COROLLA week-22 **`H`** (Monday turns-OFF) row
— a faithful path — and never the week-23 **`O`** (Saturday turns-ON) row. The PURE `day_spread` test only
feeds an `off_days` set, so it cannot express an `O` "turn-on" at all.

**Classification:** CODE DEFECT (a real divergence on real data) compounded by a DOC/PARITY-METHOD flaw
(the §F text describes a fix the code does not implement, and the suite's holiday case dodges the only
non-H/X status that exists in the calendar). `file:line` = code.py:453-462 vs
`ForecastBreakdownF.pas:1403-1407`. NOTE: whether the legacy's `INC(days)` is itself "correct" is a
separate question — but the review's job is parity, and the rebuild does not reproduce the legacy number.

---

## SHOULD-FIX-1 — No-recipe assembly: rebuild writes a raw `INV_FORECAST_INF` row the legacy never writes

**Legacy:** `ScanPartnumber` sets `Skip:=True` for any assembly whose `SELECT_ForecastDetail` returns 0
rows (`:750-756`). `UpdateForecast` wraps the ENTIRE per-entry body — including the raw
`INSERTUPDATE_ForecastInfo` write — in `if (not fEntries[i].Skip)` (`:1095`). So a no-recipe assembly
writes **nothing** (no raw, no breakdown) — only a "No breakdown for part number" log.

**Rebuild:** for a no-recipe assembly it still executes step 2a — `INSERTUPDATE_ForecastInfo` (REPLACE raw
forecast) — BEFORE the `pick_ratio is None` gap-alarm `continue` (code.py:562-574). The e2e test asserts
this as intended ("4. BOM-miss -> raw forecast STILL written", lines 244-247), calling it "faithful: raw
write precedes explode" — but that ordering only holds for NON-skipped assemblies in the legacy; a skipped
(no-recipe) assembly never reaches the raw write at all.

**Impact:** an extra `INV_FORECAST_INF` row (keyed by assembly) per no-recipe assembly/week that the legacy
would not create. Benign for the Order (which reads the breakdown, not the raw), and the D-Bug-1 gap alarm
is a deliberate, defensible improvement — but the RAW-table row count diverges from the legacy.
**Classification:** CODE DEFECT (minor) + a test that codifies the wrong oracle. Either skip the raw write
for a no-recipe assembly (faithful), or explicitly document the raw write as an intended divergence
alongside D-Bug-1 (it currently is not). `file:line` = code.py:562-566 + e2e:244-247 vs
`ForecastBreakdownF.pas:1095`.

---

## SHOULD-FIX-2 — Fixed 13/14-week legacy loop vs the rebuild's variable per-FST loop (phantom weeks)

**Legacy (`UpdateForecast:1089-1099`):** `count := 14 if fiAssemblerName='WQS' else 13`; then
`for j:=1 to count` over `Weeks[1..14]` REGARDLESS of how many FST segments were parsed. `SetLength` zeroes
the record array, so unparsed slots are `WeekNumber=0, WeekDate='', WeekCount=0`.

**Rebuild (`import_830`):** `for wk in e["weeks"]` — only the FSTs ACTUALLY present.

**Divergence directions:**
- A LIN with **fewer than 13 FSTs** → the legacy writes phantom rows: `INSERTUPDATE_ForecastInfo
  (@WeekNumber=0, @WeekDate='', @Count=0)` and a `DoPartNumberForecast(FCCount=0)` → all-zero day buckets
  at week 0. The rebuild writes none. (Phantom rows would land at `IN_WEEK_NUMBER=0` / blank week-date.)
- A LIN with **more than 13 FSTs** (non-WQS) → the legacy DROPS FST #14+; the rebuild keeps them.

**Evidence (live `Inventory`):** `IN_WEEK_NUMBER=0 OR VC_WEEK_DATE='' / NULL` rows = **0** in both
`INV_FORECAST_INF` and `INV_BREAKDOWN_FC_INF`. So real TEMA feeds appear to send a full FST complement and
the phantom path does not fire today — this is LATENT, but it IS a code difference that would surface on a
short-LIN or >13-FST golden. **Classification:** CODE DEFECT (latent) / UNPROVABLE-without-golden which
direction matters. Cannot adjudicate the production impact without a captured 830 showing a partial LIN.

---

## NIT-1 — `pick_ratio` whitespace handling diverges on a multi-space effective month

**Legacy (`:1153`):** treats `Active Date` as the default ratio only if `em = '' OR em = ' '` (exactly
empty or exactly one space). A TWO-space (or tab) value falls to the dated branch → `tm` = two spaces →
won't match a `yyyy` weekdate → `bd=FALSE` → silent drop / (rebuild: gap alarm).

**Rebuild (`_ratio_matches`, code.py:172):** `em = (effective_month or "").strip(); if em == "": default`.
So `'  '` (two spaces) strips to `''` → treated as DEFAULT → ratio applied.

**Counterexample:** recipe with `effectiveMonth='  '`, weekDate `20260615`: legacy → `None` (gap);
rebuild → applies the ratio. **Latent:** all 50 live recipes have exactly `' '` (single space)
(`forecast-tables-analysis.md §1.4`, re-confirmed: recipe table is a frozen parity baseline). Not
weaponizable on real data. `file:line` = code.py:172 vs `ForecastBreakdownF.pas:1153`.

---

## NIT-2 — `AD_GetSpecialDateWeek` parameter name (`@Line` legacy vs `@LineName` rebuild) — NOT a defect

The legacy ADO call names the param `@Line` (`:1391`); the proc's real param is `@LineName`. ADO binds
stored-proc params positionally, so the legacy worked despite the wrong name. The rebuild uses the correct
named `@LineName = ?` (code.py:447). No divergence — recorded only to pre-empt a false alarm. The `'ALL
LINES'` default for a blank component line is consistent both sides (the proc's `@LineName<>''` branch
UNIONs the `LineID is null` site-wide special dates; `'ALL LINES'` is not a real `Line` row — verified 0
rows — so a blank-line component still picks up site-wide H/X, same as legacy).

---

## CONFIRMED-FAITHFUL (attacked, held up)

- **EXPLODE truncation order** — identical for all non-negative inputs; double-trunc preserved (BLOCKER
  result #1). 3-way share, zero-ratio guard, wheelcount-for-all-non-tire: faithful.
- **DELETE window** — trim-both-ends keep-middle proven on spike (result #3); assembly→component CROSS
  APPLY resolution + both inclusive boundaries faithful.
- **D10 raw week (R1 catch)** — `int(do_ref[2:4])`: DO 2624 → stored 24 (NOT 25). The
  `production_offset` is applied ONLY to `lookupWeek` for the calendar lookup (code.py:580), and the RAW
  `rawWeek` is passed to `INSERTUPDATE_BreakdownForecastInfo @WeekNumber` (code.py:590). No path leaks the
  offset into the stored value. Re-proven; e2e case 2 confirms stored at week 24.
- **Delete-then-accumulate idempotency** — re-import same 830 not doubled (e2e case 3); renamed re-drop is
  a ledger NO-OP (case 7). The additive upsert behaves as REPLACE because the delete clears the slice.
- **`INSERTUPDATE_ForecastInfo` REPLACE** + kanban exists/update asymmetry — proc used verbatim; faithful.
- **Year-blind key (Dec→Jan, DO 2601 and 2701 both → week 1)** — the rebuild keeps the legacy year-blind
  key and the delete-forward guard (D-Bug-5 deferred to M4). Faithful to legacy (not a NEW divergence).
- **NULL-qty poison (T3)** — `day_spread` returns ints (never None); a NEW row is non-NULL. A pre-existing
  hand-NULLed row would still poison the additive UPDATE (same proc), but 0 such rows live → latent and
  equally affects legacy. No divergence.

---

## ATTACK ON THE PARITY METHOD

The PURE test (35/35) is self-consistent but cannot see either BLOCKER: it feeds `day_spread` an
`off_days` set (so the non-H/X turn-ON case is inexpressible) and never drives the driver's supplier
decision. The e2e suite (24/24) runs the REAL driver on the spike — genuinely valuable — but its two
relevant cases dodge the divergences: the holiday case uses the only turns-OFF (`H`) status and skips the
`O` Saturdays that exist in the same calendar, and the supplier path uses NULL-supplier synthetic parts so
`rowSupplier or compSupplier` never has to choose. Both BLOCKERs ship GREEN. Neither test diffs against the
legacy's actual stored values for these two columns/paths — they assert the rebuild agrees with the
rebuild's own expected. To close the gap a parity case must: (a) use a component whose part-master supplier
≠ the feed supplier and assert `VC_SUPPLIER_CODE` matches the COMPONENT supplier; and (b) import an
assembly on COROLLA in ISO week 13/15/17/20/23/30 and assert the legacy Mon-Sat spread.

---

## VERDICT

**The rebuild is NOT proven equivalent to the legacy 830 forecast math/semantics.** Two BLOCKER
divergences change stored numbers on REAL, currently-present inputs:

- **B1 — breakdown supplier:** the rebuild writes every breakdown row under the feed supplier; the legacy
  writes each under the COMPONENT's part-master supplier (proven 959/959 on live). Wrong `VC_SUPPLIER_CODE`
  + wrong upsert key on every breakdown row; the supplier-blind delete silently rewrites supplier on
  re-import. CODE DEFECT.
- **B2 — weekend-overtime day-spread:** for the 6 `O` (overtime) Saturdays on COROLLA in the live
  VehicleOrder calendar, the legacy spreads over Mon-Sat (days=6) and the rebuild over Mon-Fri (days=5) —
  e.g. `[23,23,23,23,23,23,0]` vs `[30,27,27,27,27,0,0]` for qty 138. Every day bucket differs; the
  rebuild also contradicts its own documented D-Bug-3 fix. CODE DEFECT + parity-method/doc flaw.

The EXPLODE truncation math, the DELETE window, the D10 RAW-week store, and delete-then-accumulate
idempotency ARE faithful and proven. SHOULD-FIX-1 (raw write for no-recipe assembly) and SHOULD-FIX-2
(fixed-13-week loop / phantom weeks) are real code differences that are LATENT on today's data;
SHOULD-FIX-2 and the negative-WeekCount NIT are UNPROVABLE in production impact without a captured golden
830 (a partial-LIN / negative-qty feed). NIT-1 (multi-space effmonth) is latent (recipe table all single
space). Modulo the documented golden-830-pending offset/ISA gap, **the day-spread and the breakdown
supplier do NOT reproduce the legacy on inputs that exist in the live data today** — fix B1 and B2 before
cutover.

---

## RE-VERIFY (round 2) — branch `m2-forecast-import-830` (developer claims B1 + B2 fixed)

**Stance unchanged:** wrong-until-proven. This round re-attacks the TWO BLOCKERs against the LEGACY
`.pas` + live procs + live data + an INDEPENDENT re-implementation of the legacy day-spread (not the
rebuild's own oracle). Re-read: `ForecastBreakdownF.pas` (supplier `:1349`, day-spread `:1398-1407`,
explode `:1081-1312`, anchor `:283-287`), `SELECT_PartsStockInfo` / `DELETE_ForecastInfo` /
`INSERTUPDATE_BreakdownForecastInfo` (live bodies on `mssql-spike`), `code.py`,
`test_forecast_import_e2e.py` (now 35 checks), `test_forecast_import_build.py` (now 43 checks).

### B1 — per-component breakdown supplier → **RESOLVED (proven)**

- **Code now correct.** `code.py:680` is `rowSup = compSupplier` (the COMPONENT's part-master supplier
  from `_read_part_master` → `SELECT_PartsStockInfo(@PartNum := compCode)`, code.py:671); it is written to
  `@Supplier` of `INSERTUPDATE_BreakdownForecastInfo` (code.py:687). The feed supplier (`rowSupplier =
  supplier`) is now used ONLY for the RAW `INSERTUPDATE_ForecastInfo` (code.py:649-654). This is exactly
  the legacy's two-source split: `ForecastBreakdownF.pas:1107` (raw = feed `fEntries[].Supplier`) vs
  `:1349-1450` (breakdown = `FieldByName('Supplier Code')` read from `SELECT_PartsStockInfo(@PartNum :=
  PN)` where `PN` is the COMPONENT — confirmed `DoPartNumberForecast` is called per component at
  `:1233/1242/1250/...`). The former defect `rowSupplier or compSupplier` is gone.
- **Legacy reality re-proven on live `Inventory`** (`SELECT_PartsStockInfo` body confirms
  `s.VC_SUPPLIER_CODE` via `INV_PARTS_STOCK_MST.IN_SUPPLIER_ID = INV_SUPPLIER_MST.IN_SUPPLIER_ID`):
  ```
  breakdown rows total                         : 959
  distinct breakdown suppliers                 : 14    (raw forecast: 1, the feed = 71930)
  supplier == component part-master supplier   : 959 / 959
  supplier != component part-master            :   0 / 959
  breakdown supplier == feed supplier 71930    :   0 / 959
  ```
- **Dynamic proof (REAL `import_830` driver, spike, restored as-found):** a 830 whose tire/wheel
  components resolve to DIFFERENT part-master suppliers (`ZZS1A` / `ZZS2B`), feed supplier `ZZF83`:
  - tire breakdown `VC_SUPPLIER_CODE = ZZS1A` (component supplier, ≠ feed), wheel = `ZZS2B` — the two
    breakdown rows carry DIFFERENT per-component suppliers.
  - RAW `INV_FORECAST_INF` row carries the FEED supplier `ZZF83`.
  - re-import (delete-then-accumulate) leaves ONE row per component, qty not doubled — and because the
    breakdown is now keyed by the COMPONENT supplier (the upsert key `(VC_SUPPLIER_CODE, VC_PART_NUMBER,
    IN_WEEK_NUMBER)`), the delete (supplier-blind, confirmed: `DELETE_ForecastInfo` deletes by
    assembly→component + week-date only) removes the SAME rows it re-inserts. The supplier-blind delete is
    CONSISTENT with the per-component write (a re-import no longer rewrites the supplier).
- **Delete/upsert-key alignment confirmed.** `DELETE_ForecastInfo` body (live) has no supplier predicate;
  `INSERTUPDATE_BreakdownForecastInfo` keys on `(VC_SUPPLIER_CODE, VC_PART_NUMBER, IN_WEEK_NUMBER)` and is
  additive on UPDATE. With the per-component supplier the write lands under the same key it deletes →
  idempotent. (Under the OLD code the re-insert under the feed supplier created a divergent key the
  supplier-blind delete still cleared, silently rewriting supplier — that hazard is gone.)

### B2 — 'O' overtime day-spread → **RESOLVED (proven)**

- **Code now byte-faithful to `:1398-1407`.** `day_spread` (code.py:288-301) consumes the FULL
  `AD_GetSpecialDateWeek` row stream (`_calendar_rows`, code.py:498-533 returns every `(day,status)`,
  NOT a pre-reduced off-set) and applies the legacy's UNCONDITIONAL running counter:
  `H`/`X` → `workday[n]=False; days-=1`; else (`O`/any non-H/X) → `workday[n]=True; days+=1`. The former
  D-Bug-3 "True only if previously off" (which ignored `O` and so divided by 5 not 6) is REVERSED.
- **Live calendar re-confirmed independently** (direct `EXEC AD_GetSpecialDateWeek` on `VehicleOrder`):
  week 23 COROLLA → Day Number 6 (Sat 2026-06-06) status `O`; week 22 → Day 1 (Mon 2026-05-25) `H`;
  week 30 → Day 6 (Sat 2026-07-25) `O`. Status universe in `SpecialDate` = only `H` (11) and `O` (6) —
  no `X`/`N`/`W`/`P`. Day-number mapping `DATEPART(DW, date + @@DATEFIRST - 1)` = Mon=1…Sat=6…Sun=7 =
  the rebuild's `workday[1..7]`.
- **Day-buckets proven vs an INDEPENDENT legacy reimplementation** (not the rebuild's oracle) across the
  full status×day matrix — all MATCH:
  | input | rebuild | legacy hand-calc |
  |---|---|---|
  | O-Sat wk23 qty138 | `[23,23,23,23,23,23,0]` | `[23,23,23,23,23,23,0]` |
  | H-Mon wk22 qty138 | `[0,36,34,34,34,0,0]` | `[0,36,34,34,34,0,0]` |
  | no special date qty138 | `[30,27,27,27,27,0,0]` | `[30,27,27,27,27,0,0]` |
  | O-Sat qty99 (poll) | `[19,16,16,16,16,16,0]` | `[19,16,16,16,16,16,0]` |
  | **H on already-OFF Sat** qty100 (under-count) | `[25,25,25,25,25,0,0]` | `[25,25,25,25,25,0,0]` |
  | **O on already-ON Wed** qty100 (double-count) | `[20,16,16,16,16,0,0]` | `[20,16,16,16,16,0,0]` |
  | H-Mon + O-Sat qty138 | `[0,30,27,27,27,27,0]` | `[0,30,27,27,27,27,0]` |
  | two O Sat+Sun qty140 | `[20,20,20,20,20,20,20]` | `[20,20,20,20,20,20,20]` |
  | all-5-weekday H (days=0) qty50 | `[0,0,0,0,0,0,0]` | `[0,0,0,0,0,0,0]` |

  The two latent running-counter edges the doc claims (`:1403-1407` "exactly") ARE reproduced: H-on-an-
  already-off day under-counts `days` (larger per-day qty); O-on-an-already-on weekday double-counts
  `days` (smaller per-day qty). A `count(workday)` derivation would NOT match these — the rebuild does.
  `days=0` → all-zero (legacy `if days>0 else 0`). No status/day combo found where the rebuild diverges.
- **Caveat (not a divergence):** `day_spread` starts `days = sum(workday)` (code.py:287), equal to the
  legacy hardcoded `days := 5` ONLY because the sole call site (code.py:682) always passes the default
  Mon-Fri base (no `work_days` arg). Verified the driver never passes a custom base. Holds.

### Item 3 — no regression → **HOLDS (re-proven)**

- **Explode double-truncation:** 75,000 input combos (`wc∈[0,500)`, `fr∈{1,33,50,60,66,100}`,
  `tr,wr∈{0,20,40,66,100}`) vs an independent legacy-formula reimpl → **0 mismatches**. The naive-collapse
  example still diverges and the rebuild does NOT collapse (`wc=3,fr=60,tr=66`: legacy/rebuild 0, naive 1).
- **D10 raw week:** parse `int("2624"[2:4]) = 24` (NOT 25); year-blind `int("2601"[2:4]) = 1`. The e2e
  driver run stores the breakdown at `IN_WEEK_NUMBER 24`, `wk25 = 0`. Holds.
- **Delete-then-accumulate:** the e2e driver imports the same 830 twice (idempotency bypassed) → qty stays
  `[30,27,27,27,27,0,0]`, one row per component, raw stays 138 (not 276). Holds.
- **PURE 43/43 + e2e 35/35 on a clean DB.** (NOTE — test-harness NIT, not a parity defect: a PRIOR
  aborted run left synthetic `INV_SUPPLIER_MST` rows `ZZS1A/ZZS2B`; on the next run `setup_fixtures`
  re-INSERT hits `IX_INV_SUPPLIER_MST` UNIQUE → the wheel/raw/delete checks fail SPURIOUSLY (tire passes,
  wheel/raw = None). The teardown's per-supplier DELETE goes through the `DELETE_SupplierCode` trigger; if
  a run is killed mid-flight the suppliers persist and the suite is not self-healing. Cleared the leftover
  inside a rolled-back tx; a clean re-run is 35/35. Recommend the suite pre-clean `ZZS%`/`ZZF830%` by
  sentinel at start regardless of prior trigger errors.)

### Item 4 — the anchor (`min(weekDate)` vs legacy last-LIN-first-week) → **NOT a strict no-op; a SAFER, documented deliberate divergence**

- **Mechanics confirmed.** Legacy `:283-287` reassigns `fFirstWeekDate` on the FIRST FST (`weeks=1`) of
  EVERY LIN, inside the per-LIN parse loop → after parsing it holds the **last-parsed LIN's** first-week
  date (order-dependent, ~arbitrary). The rebuild uses `min(weekDate)` over ALL entries/weeks
  (code.py:606-611). BOTH apply ONE global `@WeekDate` floor to EVERY assembly's per-assembly delete
  (legacy single `fFirstWeekDate`; rebuild single `firstWeekDate` at code.py:636). `DELETE_ForecastInfo`
  removes `VC_WEEK_DATE >= @WeekDate`. `min(weekDate) <= legacy anchor` ALWAYS → the rebuild deletes a
  SUPERSET.
- **Reachable divergence (constructed):** a multi-LIN feed whose LINs carry DIFFERENT start weeks AND the
  last-parsed LIN is NOT the earliest. Then forward rows in `[min, legacyAnchor)` are KEPT by legacy,
  DELETED by rebuild. The net stored difference = forward rows in that gap NOT re-inserted by this feed
  (i.e. stale rows from a prior import at weeks absent from the current feed): legacy retains them, the
  rebuild purges them. So it is **NOT** a strict no-op — the stored row SET can differ.
- **Live data does NOT uniformly start at one week.** Per-assembly `MIN(VC_WEEK_DATE)` in `INV_FORECAST_INF`
  = **6 distinct first-week dates** (20220822/20230130/20250602/20250609/20250825/20260323) across 42
  assemblies. (This is the accumulated multi-import table, not a single feed, so it neither proves nor
  refutes the doc's "every LIN starts at the same first week" assumption about a SINGLE 830 — but it shows
  stale forward rows at varied dates DO exist, which is the precondition for the gap to bite.)
- **Direction = SAFER, and matches the DELFOR REPLACE intent.** The divergence only ever DELETES MORE
  forward rows, never fewer; `min(weekDate)` guarantees the delete window covers every week the additive
  upsert re-inserts → the doubling risk the anchor exists to guard cannot bite for ANY LIN ordering. A
  forward week the new feed omits for an assembly is, on a planning forecast, superseded — purging it
  matches the per-assembly forward-horizon REPLACE; the legacy retains it only by accident of the
  order-dependent anchor. **Assessment: not strictly equivalent, but strictly stronger against the
  doubling hazard and consistent with the REPLACE semantics — a defensible, documented (RISK-1)
  divergence, not a silent or number-corrupting one.** It does not mis-bill/mis-count; worst case it
  cleans a stale row the legacy would have left.

### Residual latent items (re-confirmed, not weaponizable today)

- **NULL vs '' breakdown supplier (NEW NIT):** for a component with NO part master the legacy writes
  `FieldByName('Supplier Code').AsString` = `''`; the rebuild's `_read_part_master` returns `None` →
  `NULL`. Live: 0 breakdown rows with `''` or `NULL` supplier, 0 components lacking a part master, 0 with
  NULL `IN_SUPPLIER_ID` — so this cannot fire on current data. Latent NIT only.
- `@Supplier varchar(5)` truncation: source `INV_SUPPLIER_MST.VC_SUPPLIER_CODE` is itself `varchar(5)`,
  so no over-5 value can reach the proc; legacy passes the same. No divergence.
- SHOULD-FIX-1/2, NIT-1, NEG-DIV: unchanged from round 1 (latent / unprovable without a golden 830).

### RE-VERIFY VERDICT

Both prior BLOCKERs are **genuinely RESOLVED** and proven (not merely re-prosed):
- **B1** breakdown `VC_SUPPLIER_CODE` = the COMPONENT's part-master supplier (legacy `:1349`), feed
  supplier on the RAW row only — 959/959 on live + dynamic driver proof (tire≠wheel suppliers, raw=feed),
  delete/upsert-key aligned with the supplier-blind delete.
- **B2** the day-spread reproduces the legacy running DEC/INC counter EXACTLY, incl. the `O`-Saturday
  turn-ON (`[23×6,0]` for qty 138) and both latent under/over-count edges — matched vs an independent
  legacy reimpl across the full status×day matrix; 0 divergences found.

The rebuild now reproduces the legacy 830 forecast math + semantics — **explode double-truncation, the
day-spread incl. `O` overtime, the per-component breakdown supplier, delete-then-accumulate idempotency,
and the D10 raw-week store** — modulo (a) the documented anchor divergence (SAFER, not a no-op, RISK-1),
(b) the latent NULL-vs-'' no-part-master supplier NIT, and (c) the standing golden-830-pending gap (the
ISA routing-element index + byte/offset parity vs a captured TEMA feed remain UNPROVABLE from the data on
disk — same honest caveat as 856/810/997/824). On the inputs that exist in the live data today the
day-spread and the breakdown supplier DO now reproduce the legacy. **B1 and B2 are clear for cutover; the
only equivalence still UNPROVABLE is the golden-feed byte parity.** Spike left as-found (0 synthetic rows,
verified).
