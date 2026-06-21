# Forecast DATA layer — source-truth spec (M2)

**Area:** EDI inbound (830/862) forecast data layer
**Role:** the two tables the 830/862 import populates and that **M1's ASN-create read**
(`SELECT_ForecastDetailBCASN`) and the **Order "what to order" sim** consume.
**Status:** spec complete. **Analyst:** Claude / 2026-06-21.
**Verification:** every claim proven on live `Inventory` + `Inventory_Live` (mssql-spike,
read-only). Proc bodies from the authoritative **6/12 dump** (`/tmp/inv_utf8.sql`, from
`DB Schema/CreateInventory.sql` UTF-16LE) AND `OBJECT_DEFINITION` on live — flagged where they
differ from the older (superseded) snapshot the earlier analysis cited.

> ### TL;DR — the lineage, in one sentence
> An 830/862 carries a per-**assembly** weekly count → it lands raw in **`INV_FORECAST_INF`**
> (keyed by **assembly** code) → the breakdown processor explodes each assembly into its
> **component** part numbers using the BOM/ratio recipe in **`INV_FORECAST_DETAIL_INF`**,
> day-spreads each week, and writes per-**component**, per-week rows into
> **`INV_BREAKDOWN_FC_INF`** (the table the Order sim reads). `INV_FORECAST_DETAIL_INF` is the
> *recipe* (read by both the breakdown AND the M1 ASN fan-out); `INV_BREAKDOWN_FC_INF` is the
> *exploded weekly result* (read by the Order sim). **The importer must populate exactly what
> those two reads expect.**

---

## 0. Two prior-analysis corrections (snapshot-drift artifacts, now RESOLVED)

The earlier `forecasting/` specs were written against the *superseded* snapshot. The 6/12 dump
and the live DB resolve two of their open hazards:

1. **`DELETE_ForecastInfo` is NOT missing.** It exists in the 6/12 dump (`/tmp/inv_utf8.sql:2725`)
   AND live (`OBJECT_DEFINITION` returns 1780 chars, **byte-identical in `Inventory` and
   `Inventory_Live`**), with the exact 3-param `@WeekDate/@HistWeekDate/@PartNumber` signature the
   Delphi calls (`ForecastBreakdownF.pas:115`). It is a real, current proc — §3.3. The `;1` suffix
   in the Delphi call is a legacy ADO procedure-group number, not a different proc.
2. **`INSERT_ForecastDetail` has the Label/Misc columns + params.** Live param count = **15**
   (proven via `sys.parameters`), and `INV_FORECAST_DETAIL_INF` LIVE has the
   `VC_LABEL_PART_NUMBER / VC_MISC1_PART_NUMBER / VC_MISC2_PART_NUMBER` columns (cols 16/17/18).
   The "12-param / no-label" finding was pure snapshot lag. The 6/12 dump's
   `INSERT_ForecastDetail` (`:2591`) already carries all 15.

---

## 1. `INV_FORECAST_DETAIL_INF` — the BOM/ratio RECIPE (read by ASN + breakdown)

### 1.1 Columns (live `sys.columns`, `Inventory`)

| # | Column | Type | Null | Meaning |
|---|---|---|:--:|---|
| 1 | `ID_FORECAST_DETAIL` | int IDENTITY | no | surrogate PK (IDENTITY only; **no PK constraint** — see 1.3) |
| 2 | `VC_ASSY_PART_NUMBER_CODE` | varchar(12) | no | the **assembly** part (the natural row identity; ASN join key) |
| 3 | `VC_EFFECTIVE_MONTH` | varchar(8) | no | effective-month override `yyyy/MM` or `' '` (blank=default) — **dead in data**, §1.4 |
| 4 | `VC_ASSY_KANBAN_NUMBER` | varchar(4) | **yes** | the kanban the 856 emits (`LIN RC`) |
| 5 | `VC_TIRE_PART_NUMBER_CODE` | varchar(12) | no | tire component code |
| 6 | `IN_TIRE_RATIO` | int | **yes** | tire split ratio (the **only** ratio the ASN qty math uses) |
| 7 | `VC_WHEEL_PART_NUMBER_CODE` | varchar(12) | no | wheel component code |
| 8 | `IN_WHEEL_RATIO` | int | **yes** | wheel ratio (= tire ratio in 100% of rows; redundant copy) |
| 9 | `IN_RATIO` | int | **yes** | forecast ratio (breakdown explosion multiplier; **not** the ASN split) |
| 10 | `VC_VALVE_PART_NUMBER` | varchar(12) | yes | valve component |
| 11 | `VC_FILM_PART_NUMBER` | varchar(12) | yes | film component |
| 12 | `IN_ASSY_QTY` | int | **yes** | per-assembly qty multiplier (1..4) |
| 13 | `VC_BROADCAST_CODE` | varchar(20) | no | the **LIKE pattern** the ASN matches against (`[KLM]CC` etc.) |
| 14 | `VC_LAST_UPDATE` | varchar(16) | yes | 16-char `yyyymmddHHmmss` + 2 add-stamp |
| 15 | `VC_ADD` | varchar(16) | yes | 16-char add-stamp |
| 16 | `VC_LABEL_PART_NUMBER` | varchar(12) | yes | label component (LIVE — absent from old snapshot) |
| 17 | `VC_MISC1_PART_NUMBER` | varchar(12) | yes | misc-1 component |
| 18 | `VC_MISC2_PART_NUMBER` | varchar(12) | yes | misc-2 component |

### 1.2 Live state (both DBs identical)

```
rows=50  distinct_BC=29  distinct_assy=50  distinct_effmonth=1 (all ' ')
```

So **1 row per assembly** today (50 rows / 50 distinct assy). One **broadcast code fans to up to 3
assemblies** (13 multi-assy BCs; `[KLM]CC=3`, `[MNP]BB=3`, …; the rest 1:1). The ratios within a
multi-assy BC sum to 100 (e.g. `[KLM]CC` = 20+40+40).

### 1.3 KEY / uniqueness — HEAP, no enforced key (HAZARD)

```
sys.indexes(INV_FORECAST_DETAIL_INF) -> NULL | HEAP | is_unique=0 | is_primary=0   (both DBs)
duplicate (assy, effmonth) key groups = 0
```

**There is no PK constraint, no unique index, no clustered index — it is a pure heap.** The natural
key `(VC_ASSY_PART_NUMBER_CODE, VC_EFFECTIVE_MONTH)` is **enforced only by application logic**:
`INSERT_ForecastDetail` (`:2609`) does `IF NOT EXISTS (... WHERE VC_ASSY_PART_NUMBER_CODE=@AssyCode
AND VC_EFFECTIVE_MONTH=@EffectiveMonth) ... ELSE RETURN -1`. A direct INSERT, a concurrent race, or
any path that bypasses the proc can create duplicate recipe rows; the ASN read (`SELECT *`, no
`DISTINCT`) would then **fan out the ASN line**.

### 1.4 `VC_EFFECTIVE_MONTH` is effectively dead (proven)

```
'['+VC_EFFECTIVE_MONTH+']' for sampled rows -> [ ]   (a single space)
all 50 rows: a single space; ' ' = '' is TRUE in T-SQL (trailing-space trim)
```

The recipe is treated as always-effective. The ASN read's `(VC_EFFECTIVE_MONTH=@EffMonth OR
VC_EFFECTIVE_MONTH='')` passes all rows for any month. The importer should **write `' '`** (or
empty) for the default recipe to match; an actual `yyyy/MM` value would scope the row to that month
only.

### 1.5 How a forecast import lands here — IT DOESN'T (recipe ≠ feed)

**Critical:** `INV_FORECAST_DETAIL_INF` is **NOT** populated by the 830/862 import. It is **master
config** maintained by the `ForecastDetail` CRUD screen (`INSERT_/UPDATE_/DELETE_ForecastDetail`).
The 830/862 import **reads** it (to explode), it never writes it. The importer's job vs this table is
purely: *for every assembly in the feed, a recipe row must already exist* — else the breakdown logs
"No breakdown for part … count will be ignored" and that assembly's count is dropped (zero
components written). The M1 ASN read has a harder rule: a recipe row missing its manifest-cost
mapping **aborts the whole ASN** (see `SELECT_ForecastDetailBCASN-analysis.md` §4).

### 1.6 The ASN-read contract over this table (M1 keystone — cross-ref)

`SELECT_ForecastDetailBCASN(@BCode, @EffMonth)` (`/tmp/inv_utf8.sql:3011`) does
`SELECT * FROM INV_FORECAST_DETAIL_INF f LEFT JOIN INV_MANIFEST_COST_MST c ON
f.VC_ASSY_PART_NUMBER_CODE=c.VC_ASSY_PART_NUMBER_CODE WHERE @BCode LIKE VC_BROADCAST_CODE AND
((VC_EFFECTIVE_MONTH=@EffMonth OR VC_EFFECTIVE_MONTH='') AND IN_TIRE_RATIO<>0 AND IN_WHEEL_RATIO<>0)`.
The load-bearing traps (proven in the keystone analysis) the importer must keep consistent:
- **`@BCode LIKE VC_BROADCAST_CODE`** — the *column* is the pattern. The importer must write BCs as
  the `LIKE` patterns (`[KLM]CC`), not literal codes, or 28/29 BCs return zero rows.
- **NULLable ratios silently drop the row** (`NULL <> 0` = UNKNOWN). Importer/config must never
  leave a ratio NULL on a live recipe row.
- **No `ORDER BY` over a heap** → nondeterministic row order; the single-vehicle ASN branch picks
  "first row" arbitrarily for multi-assy BCs.

### 1.7 Triggers — `DeleteForecastDetail` (archive + cascade)

`DeleteForecastDetail` FOR DELETE (`/tmp/inv_utf8.sql:2661`):
```sql
INSERT INTO INV_FORECAST_DETAIL_INF_HIST SELECT * FROM deleted   -- archives the recipe row
DELETE FROM inv_forecast_inf WHERE vc_part_number IN (SELECT vc_assy_part_number_code FROM DELETED)
```
Deleting a recipe row (a) **archives it to a HIST table** (new finding vs old analysis) and (b)
cascades a delete of the **raw forecast** rows for that assembly. The cascade matches the *assembly*
code against `INV_FORECAST_INF.VC_PART_NUMBER` — valid because the raw forecast IS keyed by assembly
(§2.2). The `UPDATE_ForecastDetailInf` trigger is a no-op (`print` only).

---

## 2. `INV_BREAKDOWN_FC_INF` — the exploded weekly result (read by Order)

### 2.1 Columns (live `sys.columns`)

| # | Column | Type | Null | Meaning |
|---|---|---|:--:|---|
| 1 | `IN_WEEK_NUMBER` | int | no | **production-relative** week = `ISO_week(date) − offset` (§4) — year-blind |
| 2 | `VC_WEEK_DATE` | varchar(8) | no | `yyyymmdd` of the week (FST element-4 date); string-comparable |
| 3 | `VC_SUPPLIER_CODE` | varchar(5) | no | component's supplier (from `INV_PARTS_STOCK_MST`, not the feed) |
| 4 | `VC_PART_NUMBER` | varchar(12) | no | the **COMPONENT** part (tire/wheel/valve/…), NOT the assembly (§2.2) |
| 5 | `VC_SIZE_CODE` | varchar(10) | **yes** | component size; **written on INSERT only, never on the additive UPDATE** (§3.2) |
| 6–12 | `IN_QTY1 … IN_QTY7` | int | **yes** | the 7 **day-of-week** buckets for that week (Mon..Sun day-spread), §2.3 |

### 2.2 `VC_PART_NUMBER` holds COMPONENT codes (proven)

```
breakdown rows where part is a COMPONENT code (tire/wheel/valve/film/label/misc) = 675 / 959
breakdown rows where part is an ASSEMBLY code                                     =   0 / 959
```

Confirmed: the breakdown is keyed by **component** part. (The 284 component rows not in the *current*
50-row recipe are components from older recipe vintages — the breakdown spans 2020-2026; recipes
have churned.) Contrast the raw forecast:

```
INV_FORECAST_INF: rows=1041  part_is_assembly=1016  distinct_parts=42
```

`INV_FORECAST_INF.VC_PART_NUMBER` holds **assembly** codes (1016/1041 match the assy master). So the
explosion is the only place assembly→component happens, and `DELETE_ForecastInfo` (§3.3) has to do
the same assembly→component resolution to delete the right breakdown rows.

### 2.3 The IN_QTY1..7 = **days of a week**, NOT 7 weeks (proven)

`IN_QTY1..7` are the **7 days** of one ISO week (Mon..Sun), produced by the day-spread step. NOT 7
weekly buckets. The read proc `SELECT_ForecastPartNumberWeek(@WeekNo,@DayNo,@PartNo)`
(`/tmp/inv_utf8.sql:2076`) is a 7-way `if @DayNo=N ... SELECT IN_QTY{N}` — it returns one day's qty
for a `(week, day, part)`. Live sample (a Mon-start week):
```
IN_WEEK=36 WEEK_DATE=20260908 part=4265202R6000 size=15D : QTY1..7 = 31,27,27,27,27,0,0
```
QTY6/QTY7 (Sat/Sun) are 0 (non-working). The remainder lands on day 1 (`31 = 27+4`), confirming the
"remainder on first working day" day-spread.

### 2.4 KEY / uniqueness — HEAP again (HAZARD)

```
sys.indexes(INV_BREAKDOWN_FC_INF) -> NULL | HEAP   (both DBs)
duplicate (supplier, part, week) groups = 0
```

Pure heap, no PK/unique. The upsert dedup key **`(VC_SUPPLIER_CODE, VC_PART_NUMBER,
IN_WEEK_NUMBER)`** is currently 1:1 but **not DB-enforced** — only the upsert proc's `IF EXISTS`
check holds it. The READ side does **not** filter by supplier or year (`SELECT_ForecastPartNumberWeek`
matches `IN_WEEK_NUMBER + VC_PART_NUMBER` only) — so a duplicate `(part, week)` across two suppliers
would make the Order read non-deterministic / return two rows.

### 2.5 How the Order sim reads it (cross-ref, R1)

Order computes a production-week number `WeekOfTheYear(prodDate) − weekoffset` (offset = First
Production Week − 1) and calls `SELECT_ForecastPartNumberWeek(@WeekNo, @DayNo, @PartNo)` → `IN_QTY{DayNo}`
WHERE `IN_WEEK_NUMBER=@WeekNo AND VC_PART_NUMBER=@PartNo`. **No date/year/supplier filter** → the
match is year-blind on week-number, keyed by component part. This is why the stored `IN_WEEK_NUMBER`
must already be in production-relative terms (§4) and why the delete-forward each cycle (§3.3) is
load-bearing — it prevents stale year-N rows being read in year N+1.

---

## 3. The forecast write/delete procs + replace-merge semantics

All four write/delete procs target the **transactional** tables (`INV_FORECAST_INF`,
`INV_BREAKDOWN_FC_INF`). The recipe CRUD procs (`INSERT_/UPDATE_/DELETE_ForecastDetail`) target the
*master* and are config, not import — covered in `forecasting/forecast-detail.md`.

### 3.1 `INSERTUPDATE_ForecastInfo` — raw assembly forecast (REPLACE upsert)

`/tmp/inv_utf8.sql:1184`. Dedup key `(@Supplier, @PartNumber, @Kanban, @WeekNumber)` on the EXISTS
check, but the UPDATE/INSERT key on `(@Supplier, @PartNumber, @WeekNumber)` — **kanban is in the
exists-test but not the update-where** (a subtle asymmetry; two kanbans for the same part/week would
upsert the first-found, ignoring kanban on update).
- exists → `UPDATE … SET IN_COUNT=@Count, VC_WEEK_DATE=@WeekDate` → **REPLACE** (overwrite count).
- else → INSERT all 6 columns.
- **Not additive.** A re-run overwrites the raw count.

### 3.2 `INSERTUPDATE_BreakdownForecastInfo` — the breakdown writer (ADDITIVE upsert)

`/tmp/inv_utf8.sql:1215`. Dedup key **`(@Supplier, @PartNumber, @WeekNumber)`**, year-blind,
date-blind.
- exists → `UPDATE … SET IN_QTYn = IN_QTYn + @Qtyn` (n=1..7) → **ADDITIVE** (accumulates; does NOT
  overwrite). `VC_WEEK_DATE` and `VC_SIZE_CODE` are **NOT touched on UPDATE**.
- else → INSERT all 7 day-qtys + `VC_WEEK_DATE` + `VC_SIZE_CODE`.

**Three proven hazards:**
1. **Additive + nullable qty = NULL poison.** `IN_QTY1..7` are nullable; the additive update is
   `IN_QTYn = IN_QTYn + @Qtyn`, and `NULL + n = NULL`. Live has **0 NULL qty rows** today, so it's
   latent — but a single manually-NULLed qty row would silently zero-out (→ NULL) on the next
   accumulate. Importer must always INSERT non-NULL qtys (0, not NULL).
2. **Additive double-count if delete is skipped.** Two explosion passes for the same
   `(supplier, part, week)` without an intervening `DELETE_ForecastInfo` **sum** — the counts double.
   The delete-then-rebuild (§3.3) is what makes a re-run a *replace*. If the delete fails/partially
   runs, the breakdown silently inflates. (Open Q5 in the old analysis.)
3. **Stale `VC_SIZE_CODE`.** Written on INSERT only. Proven: part `4265202R8100` has **2 distinct
   sizes** across its breakdown rows — i.e. an existing row keeps its *first-seen* size even if the
   component's size later changes; only brand-new `(supplier,part,week)` rows pick up the new size.

### 3.3 `DELETE_ForecastInfo` — the delete-then-rebuild (trim-both-ends)

`/tmp/inv_utf8.sql:2725` (live, 1780 chars, identical both DBs). Params
`@WeekDate, @HistWeekDate, @PartNumber` where **`@PartNumber` is the ASSEMBLY code** (Delphi passes
`fEntries[i].PartNumber` = the LIN/assembly part). Four deletes:

```sql
-- (A) breakdown forward slice: resolve assembly -> its 7 component codes via CROSS APPLY VALUES,
--     delete breakdown rows for those components WHERE VC_WEEK_DATE >= @WeekDate
DELETE FROM INV_BREAKDOWN_FC_INF WHERE VC_PART_NUMBER IN (
  SELECT value FROM INV_FORECAST_DETAIL_INF
  CROSS APPLY (VALUES (...tire),(...wheel),(...valve),(...film),(...label),(...misc1),(...misc2)) c(col,value)
  WHERE VC_ASSY_PART_NUMBER_CODE = @PartNumber)
  AND VC_WEEK_DATE >= @WeekDate
-- (B) breakdown historical slice: same component resolution, WHERE VC_WEEK_DATE <= @HistWeekDate
-- (C) raw forward:     DELETE INV_FORECAST_INF WHERE VC_PART_NUMBER=@PartNumber AND VC_WEEK_DATE >= @WeekDate
-- (D) raw historical:  DELETE INV_FORECAST_INF WHERE VC_PART_NUMBER=@PartNumber AND VC_WEEK_DATE <= @HistWeekDate
```

**Semantics (the replace-merge):** the import calls this **per assembly before re-exploding**, with
`@WeekDate = fFirstWeekDate` (the first FST week date in the feed) and `@HistWeekDate = now −
histweeks×7`. It **trims both ends** — deletes everything **≥ the feed's first week** (the slice
about to be rebuilt) AND everything **≤ the history cutoff** (aged-out weeks) — **keeping only the
middle window** `(HistWeekDate, WeekDate)`. Then the additive upsert (§3.2) refills the forward slice.
So the cycle is **replace-forward + purge-history**, with the additive upsert acting as a *replace*
only because (A)/(B) cleared the slice first.

**The assembly→component resolution is the crux:** the breakdown is keyed by component, but the
import deletes by assembly, so the proc must map assembly→components (via the recipe) to find the
right breakdown rows. **An importer that deletes breakdown rows by assembly code directly would
delete NOTHING** (no breakdown row's `VC_PART_NUMBER` is an assembly code — proven §2.2). The rebuild
MUST reproduce this assembly→component-set resolution, and it depends on the recipe being intact at
delete time (delete the recipe first and the cascade can't find the components).

**Boundary inclusivity (exact):** forward `>= @WeekDate` (inclusive of the first feed week);
historical `<= @HistWeekDate` (inclusive). String comparison on `yyyymmdd` (lexical = chronological,
safe).

### 3.4 The legacy `DELETE_ForecastInfoWeekDate*` variants (superseded, but live)

Three older single-window deletes still exist (`:2303/:2318/:2340`):
- `…WeekDatePartOld(@WeekDate,@Part)` → `VC_WEEK_DATE <= @WeekDate AND part` (historical-only).
- `…WeekDatePart(@WeekNumber,@WeekDate,@Part)` → `VC_WEEK_DATE >= @WeekDate AND part` (forward-only;
  the `IN_WEEK_NUMBER>=` arm is commented out). **Note: `@Part` here is matched against the breakdown
  `VC_PART_NUMBER` directly = a COMPONENT code, NOT an assembly** — so these old variants assume the
  caller already resolved to a component. `DELETE_ForecastInfo` (§3.3) supersedes them by doing the
  assembly→component resolution itself. The breakdown form calls **only** `DELETE_ForecastInfo;1`
  (`ForecastBreakdownF.pas:115`); the variants are dead-but-present.

---

## 4. D10 — the week-number mapping (PROVEN on live breakdown data)

### 4.1 `VC_WEEK_DATE` format

```
DATALENGTH=8 for all rows; year-prefix distribution: 2020:75 2022:76 2023:158 2024:50 2025:125 2026:475
```
→ **`yyyymmdd`** (NOT `yyyy/MM`). It is the FST element-4 week date the feed supplies, and the
day-of-week is the **Monday** of the production week in the normal case (proven below). String
compare = chronological (used by the delete windows and `SELECT_ForecastSupplier`).

### 4.2 `IN_WEEK_NUMBER = ISO_week(VC_WEEK_DATE) − 1` (the offset, proven)

For 2026 rows, `DATEPART(ISO_WEEK, weekdate) − IN_WEEK_NUMBER` is **1 across the whole horizon**:
```
WEEK_DATE   IN_WEEK  iso_wk  iso_minus_stored
20260323    12       13      1
20260420    16       17      1
20260601    22       23      1
20260706    27       28      1
20260803    31       32      1   ... (consistent for every Monday week)
```
And `INV_FIRST_PRODUCTION_DAY[2026] = (2026, 2026-01-05, 2)` → `INT_FIRST_PRODUCTION_WEEK = 2` →
**offset = First Week − 1 = 1**. So the stored week is **production-relative = ISO_week − 1** for
2026. This is exactly D10: TEMA's FST09 "DO" week (chars 3–4 of element 9) is production-relative and
already equals `ISO − offset`; the breakdown stores it **verbatim** (`checkweeknumber`), and Order
reads with `ISO(prodDate) − offset`, so the two reconcile.

### 4.3 How the 830 FST09 week maps to the buckets

- **Which week bucket:** `IN_WEEK_NUMBER` = `StrToInt(copy(FST09,3,2))` verbatim from the feed (NOT
  recomputed). It is production-relative; do not add the offset to the stored value. (The Delphi
  write side adds the offset only for the `AD_GetSpecialDateWeek` holiday lookup, never to the stored
  number — R1 reconciliation.)
- **Which day buckets (IN_QTY1..7):** the week's count is day-spread Mon..Sun =
  `IN_QTY1..IN_QTY7`. `ratiocount = FCCount div working_days; leftover = FCCount mod working_days`;
  each working day gets `ratiocount`, and **all leftover lands on the first working day**;
  non-working days (Sat/Sun + ALC special-date H/X) get 0.

### 4.4 The `20260712` anomaly (Sunday week-date — flag)

```
20260706 = Monday   (IN_WEEK 27, iso 28, offset 1)  -- normal
20260712 = Sunday   (IN_WEEK 28, iso 28, offset 0)  -- ANOMALY
20260720 = Monday   (IN_WEEK 29, iso 30, offset 1)  -- normal
20260908 = Tuesday  (IN_WEEK 36, iso ..)            -- not a Monday either
```
Most week-dates are the **Monday** of the production week, but a few are not (`20260712` is a Sunday;
`20260908` a Tuesday). Because ISO_WEEK is computed from the *actual* date, a non-Monday week-date can
shift the ISO week by one and break the clean `offset=1` (the Sunday `20260712` gives `offset=0`).
**Implication:** the offset is reliable when the feed's week-date is the Monday; the rebuild should
**store the feed's FST09 week verbatim and NOT recompute it from the date**, because recomputing
`ISO(weekdate) − offset` would disagree with the feed on these non-Monday rows. (This is precisely
why the legacy stores the raw feed week and only the *read* side recomputes — and why D10 says
"ingest verbatim, read with `ISO(prodDate) − offset`.")

---

## 5. Read-side consistency check (the importer's contract)

Both consumers hit these exact tables; the importer must populate exactly what they read:

| Consumer | Proc | Reads | Key it matches |
|---|---|---|---|
| **M1 ASN create** | `SELECT_ForecastDetailBCASN` | `INV_FORECAST_DETAIL_INF` (recipe) | `@BCode LIKE VC_BROADCAST_CODE`; ratios `<>0`; eff-month |
| **M1 856 report** | `REPORT_EDI856` | `INV_FORECAST_DETAIL_INF` (Kanban) | `d.VC_ASSY_PART_NUMBER = f.VC_ASSY_PART_NUMBER_CODE` |
| **Order sim** | `SELECT_ForecastPartNumberWeek` | `INV_BREAKDOWN_FC_INF` | `IN_WEEK_NUMBER + VC_PART_NUMBER(component)` |
| **File emit** | `SELECT_ForecastSupplier` | `INV_BREAKDOWN_FC_INF` | `VC_WEEK_DATE > @WeekDate` (string), ordered supplier/part/weekdate |

Consistency requirements the importer MUST honor (each proven above):
1. **Recipe before feed.** Every assembly in the feed needs a recipe row (`INV_FORECAST_DETAIL_INF`)
   with the BC as a `LIKE` pattern, non-NULL ratios, and (for the 856) a kanban + a manifest-cost
   mapping — else the ASN read drops or aborts.
2. **Breakdown keyed by COMPONENT + production-relative week.** Order reads
   `(IN_WEEK_NUMBER, VC_PART_NUMBER=component)`; the importer must write the exploded component rows
   with the verbatim feed week (production-relative), or Order reads nothing.
3. **`VC_WEEK_DATE` = `yyyymmdd`** (string-comparable); the file-emit read filters on it lexically.
4. **Delete-then-rebuild must run** (§3.3) so the additive upsert behaves as replace, not double.

---

## 6. Inventory vs Inventory_Live drift (the forecast-recipe vintage)

**Recipe table — NO drift (frozen baseline):**
```
INV_FORECAST_DETAIL_INF CHECKSUM_AGG: Inventory = -294555574 ; Inventory_Live = -294555574  (equal)
both: 50 rows / 29 BC / 50 assy / eff-month=' '
```
Matches the ASN-keystone finding: the recipe is a faithful, frozen parity baseline.

**Breakdown table — DRIFT CONFIRMED (different import vintage):**
```
INV_BREAKDOWN_FC_INF CHECKSUM_AGG: Inventory = 154058948 ; Inventory_Live = 607782123  (DIFFER)
Sept-2026 week-dates:
  Inventory       : 20260908 (19 rows)
  Inventory_Live  : 20260908 (19 rows) + 20260914 (19 rows)   <- 19 EXTRA rows
```
`Inventory_Live` received a **later forecast import** (an additional week, `20260914`) that the
`Inventory` snapshot predates. This is exactly the "4718-4721 ASNs were a different recipe vintage"
artifact the ASN-keystone parity work hit — **but it is a TRANSACTIONAL (breakdown) drift, not a
recipe drift.** The two DBs were breakdown-rebuilt at different points in time.

**Would a fresh import change it?** Yes — a fresh 830 run would (a) `DELETE_ForecastInfo` the forward
slice (`>= fFirstWeekDate`) for each fed assembly, then (b) re-explode and additive-upsert. So a
fresh import **rewrites the forward breakdown horizon** and would reconcile (or further diverge) the
two DBs depending on the feed. **For parity testing the breakdown table must be treated as a *moving*
artifact** — pin it to a known feed + a known run timestamp, and diff after a controlled run; do NOT
assume `Inventory` and `Inventory_Live` agree on the breakdown the way they do on the recipe.

---

## 7. What the rebuild's importer MUST write (and the traps that silently break it)

**MUST reproduce:**
1. **Raw forecast** (`INV_FORECAST_INF`) keyed by **assembly** code, REPLACE-upsert on
   `(supplier, part, week)` (overwrite count + week-date). `IN_WEEK_NUMBER` = FST09 verbatim
   (production-relative); `VC_WEEK_DATE` = `yyyymmdd`.
2. **Breakdown** (`INV_BREAKDOWN_FC_INF`) keyed by **component** code (after assembly→component
   explosion via the recipe), ADDITIVE-upsert on `(supplier, part, week)`, 7 **day-of-week** qty
   buckets, `IN_WEEK_NUMBER` = production-relative (ISO − offset), `VC_WEEK_DATE` = `yyyymmdd`.
3. **Delete-then-rebuild** before each explosion: per assembly, resolve to its 7 components and
   delete breakdown rows `VC_WEEK_DATE >= firstFeedWeek` AND `<= histCutoff`; delete raw
   `INV_FORECAST_INF` by assembly in the same two windows. **Keep only the middle window.** This is
   what turns the additive upsert into a replace.
4. **Recipe contract:** keep BCs as `LIKE` patterns, ratios non-NULL, eff-month `' '` for defaults,
   so the M1 ASN read + 856 read stay correct.

**Traps that will SILENTLY break it:**
- **T1 — delete by assembly, not component.** The breakdown is keyed by component; deleting by
  assembly code directly removes 0 rows. The delete MUST resolve assembly→components via the recipe
  (proven: 0/959 breakdown rows have an assembly code). Recipe must be intact at delete time.
- **T2 — additive double-count.** Skip/fail the delete and the breakdown silently doubles (additive
  `+= qty`). The delete is load-bearing for replace semantics.
- **T3 — NULL qty poison.** Always write 0, never NULL, into `IN_QTY1..7` (additive `NULL+n=NULL`).
- **T4 — stale size.** `VC_SIZE_CODE` is INSERT-only; an additive UPDATE keeps the first-seen size.
  Proven (part `4265202R8100` carries 2 sizes). Decide whether to refresh size on update.
- **T5 — week recompute.** Store FST09 verbatim; do NOT recompute `IN_WEEK_NUMBER` from the date
  (non-Monday week-dates like `20260712` give the wrong offset). Read with `ISO(prodDate) − offset`.
- **T6 — heap, no enforced key.** Both tables are heaps; the `(assy,effmonth)` and
  `(supplier,part,week)` keys are app-enforced only. A bypass path creates duplicates → ASN/Order
  reads fan out or go non-deterministic. The rebuild should add real unique constraints.
- **T7 — LIKE direction / equality rebind.** `:bcode LIKE VC_BROADCAST_CODE` — keep the column as
  the pattern; an equality rebind kills 28/29 BCs.
- **T8 — year-blind reads.** Order matches week-number with no year filter; the delete-forward each
  cycle is what prevents year-N rows leaking into year N+1. The rebuild either keeps the
  delete-forward discipline or adds a year/effective-date dimension to the breakdown key.
- **T9 — breakdown is a moving parity artifact.** `Inventory` vs `Inventory_Live` breakdown DIFFER
  by import vintage (proven: Sept-2026 +19 rows in Live). Pin feed + run-time for parity; only the
  recipe is a frozen baseline.

---

## Appendix — file/line citations

- Tables (live `sys.columns`/`sys.indexes`, `Inventory` + `Inventory_Live`).
- `INSERTUPDATE_ForecastInfo` — `/tmp/inv_utf8.sql:1184`.
- `INSERTUPDATE_BreakdownForecastInfo` — `/tmp/inv_utf8.sql:1215`.
- `DELETE_ForecastInfo` — `/tmp/inv_utf8.sql:2725` (live-identical, 1780 chars, both DBs).
- `DELETE_ForecastInfoWeekDate*` variants — `/tmp/inv_utf8.sql:2303 / 2318 / 2340`.
- `INSERT_ForecastDetail` (15 params, label/misc) — `/tmp/inv_utf8.sql:2591`.
- `DeleteForecastDetail` trigger (archive + cascade) — `/tmp/inv_utf8.sql:2661`.
- `SELECT_ForecastSupplier` — `/tmp/inv_utf8.sql:2058`.
- `SELECT_ForecastPartNumberWeek` — `/tmp/inv_utf8.sql:2076`.
- Delphi caller of `DELETE_ForecastInfo;1` — `ForecastBreakdownF.pas:110-125`.
- ASN read — `docs/analysis/production-readiness/sql/SELECT_ForecastDetailBCASN-analysis.md`.
- 856 read — `docs/analysis/edi/856/report-edi856-data-analysis.md`.
- Prior (snapshot-era) specs — `docs/analysis/forecasting/forecast-detail.md`,
  `docs/analysis/forecasting/forecast-breakdown.md`.
