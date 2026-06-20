# AD_FRSPULL — authoritative SQL-semantics analysis

**Proc:** `dbo.AD_FRSPULL` (DB: **VehicleOrder** — the shared GALC build database; read-only cross-DB from the rebuild).
**Author/created:** David Verespey, 2013/03/06 (per the live header). Description: "FRS with model year info".
**Role (M1 ASN-creation keystone):** the ratio fan-out (`CalculateASNFRS`, Delphi `DataModule.pas:5106`) calls this **once per (line, production-date window)** to obtain, **per broadcast code (BC)**, how many vehicles were built (`VEHICLES`) and an implied order count (`ORDERS`). Those two numbers drive every ASN detail quantity downstream (see §8 and the companion `SELECT_ForecastDetailBCASN-analysis.md`).
**Decoded body:** `docs/analysis/production-readiness/AD_FRSPULL-shared.sql`.
**Verification:** every claim below was proven on the **live VehicleOrder backup** (`mssql-spike`, real GALC, 2.33M `Vehicle` rows). All queries were bounded (single line `COROLLA` / LineID 1, 1-day or ≤5-day `DateCreated` windows, or the small dimension tables). Verified-vs-inferred is stated per item.

---

## 0. Live body vs decoded copy (VERIFIED — no SQL drift)

`OBJECT_DEFINITION(OBJECT_ID('dbo.AD_FRSPULL'))` on the live VehicleOrder is **byte-identical in its SQL** to the decoded body in `AD_FRSPULL-shared.sql`. The only difference is cosmetic: the live object carries the original `-- Author / Create date 2013/03/06 / FRS with model year info` comment header that the decoded copy replaced with rebuild-context prose. **No behavioral drift.**

Signature (5 params; **2 are dead — see §6c**):

```sql
CREATE PROCEDURE [dbo].[AD_FRSPULL]
    @begindate datetime, @enddate datetime,
    @Start int, @Last int,            -- DECLARED BUT NEVER REFERENCED in the body
    @LineName varchar(50)
```

The proc is two `SELECT ... GROUP BY` blocks joined by `UNION`, ending `ORDER BY BC`. It returns a 4-column result: **`BC char(3)`, `ORDERS int`, `VEHICLES int`, and a literal `''`** (the 4th column is always an empty string — a placeholder the Delphi caller ignores).

---

## 1. Lineage — tables, joins, filters, projection

```
Vehicle v                       -- the 2.33M-row fact table (one row per built vehicle)
  JOIN Model m  ON v.ModelID = m.ModelID            -- m.ModelYearCode char(1)
  JOIN Line  l  ON v.LineID  = l.LineID             -- filtered l.LineName = @LineName
  JOIN vehicledata vdX ON v.VehicleID = vdX.VehicleID   -- the EAV value rows (heap)
  JOIN DataItem  iX  ON vdX.DataItemID = iX.DataItemID
                     AND iX.DataItemDescription = '<GROUNDTIRE|GROUNDWHEEL|SPARETIRE>'
WHERE v.DateCreated >= @begindate AND v.DateCreated <= @enddate AND l.LineName = @LineName
GROUP BY <the char(3) BC expression>
```

GALC stores vehicle attributes in an **EAV shape**: `vehicledata(VehicleID, DataItemID, DataValue)` where `DataItemID` resolves through `DataItem` to a textual `DataItemDescription`. The proc pivots three attributes — `GROUNDWHEEL`, `GROUNDTIRE`, `SPARETIRE` — by self-joining `vehicledata` once per attribute and constraining the joined `DataItem` row by description.

Column types that govern the output (VERIFIED via `sys.columns`):

| Column | Type | Nullable | Collation |
|---|---|---|---|
| `Model.ModelYearCode` | **char(1)** | YES (NULL-free in live Model) | `SQL_Latin1_General_CP1_CI_AS` |
| `VehicleData.DataValue` | **varchar(250)** | YES | `SQL_Latin1_General_CP1_CI_AS` |
| `DataItem.DataItemDescription` | varchar(32) | YES | `SQL_Latin1_General_CP1_CI_AS` |
| `Vehicle.DateCreated` | datetime | YES | — |
| `Line.LineName` | varchar(15) | YES | `SQL_Latin1_General_CP1_CI_AS` |

**Live shape note:** there is exactly **one Line — `COROLLA` (LineID 1)** in this backup, so all proofs use it. `Model` holds 15 model-year codes: `A B C D E F G H I J K L M N P`.

---

## 2. The EXACT BC + ORDERS/VEHICLES formulas

### Ground block (first SELECT)

```
BC       = CONVERT(char(3), m.ModelYearCode + vd2.DataValue + vd1.DataValue)
              where vd2 = GROUNDWHEEL value, vd1 = GROUNDTIRE value
VEHICLES = COUNT(*)
ORDERS   = COUNT(*) * 4
```

Concatenation order is **ModelYear, then WHEEL, then TIRE** (`vd2` is GROUNDWHEEL, `vd1` is GROUNDTIRE — the alias numbering is the reverse of the read order; do not assume vd1=first attribute).

### Spare block (second SELECT)

```
BC       = CONVERT(char(3), m.ModelYearCode + vd1.DataValue)
              where vd1 = SPARETIRE value
VEHICLES = COUNT(*)
ORDERS   = COUNT(*)            -- 1:1, no ×4
filter   = ... AND vd1.DataValue <> 'M'   -- "Quick fix for no spare broadcast"
```

The two blocks are `UNION`'d (distinct) and `ORDER BY BC`.

**ORDERS×4 = the 4-corners assumption (ground only).** A vehicle has 4 ground wheel/tire corners; spare is a single unit, so spare ORDERS = VEHICLES. `ORDERS` is **not** a quantity — it is the gate for the downstream "(No Ratio)" small-volume branch (`Orders <= 5`). `VEHICLES` is the qty multiplier. See §8 for how each is consumed.

---

## 3. Edge-case matrix (each PROVEN with a bounded query)

### 3.1 BC composition + char(3) truncation (VERIFIED — fits exactly today, silent-truncation hazard latent)

Ground BC width = `len(ModelYearCode) + len(GROUNDWHEEL) + len(GROUNDTIRE)`. `ModelYearCode` is `char(1)` (always 1). The two `DataValue`s are `varchar(250)` — *width is data-dependent*. Measured over a 5-day COROLLA window:

```
$ SELECT DataItemDescription, MAX(LEN(DataValue)) max_len, MAX(DATALENGTH(DataValue)) max_bytes
    FROM Vehicle .. vehicledata .. DataItem  WHERE LineID=1 AND DateCreated in [06-15,06-20)
    AND DataItemDescription IN ('GROUNDTIRE','GROUNDWHEEL','SPARETIRE') GROUP BY ...
GROUNDTIRE  | 1 | 1
GROUNDWHEEL | 1 | 1
SPARETIRE   | 1 | 1
```

Every `DataValue` is **exactly 1 character**. So:
- **Ground BC = 1+1+1 = 3 chars → `char(3)` holds it with no truncation.**
- **Spare BC = 1+1 = 2 chars → right-padded to `char(3)` as `"XY "` (one trailing space).**

Proven on real proc output (06-18, COROLLA), bracketed + measured:

```
$ INSERT #r EXEC AD_FRSPULL ... ; SELECT '['+BC+']', DATALENGTH(BC), LEN(BC), ORDERS, VEHICLES ...
[NBB] | 3 | 3 | 2028 | 507     <- ground: 3 real chars
[NEE] | 3 | 3 |  776 | 194
[NN ] | 3 | 2 |  798 | 798     <- spare: 2 chars + trailing space (LEN trims it to 2)
[NP ] | 3 | 2 |   20 |  20
[PN ] | 3 | 2 |    1 |   1
```

**THE TRUNCATION TRAP (latent, not firing today):** `CONVERT(char(3), …)` **silently truncates** to 3 characters with no error. It is safe *only because all three components are 1 char in current data*. If GALC ever emits a 2-char `GROUNDWHEEL` or `GROUNDTIRE` value (the column allows 250), the ground BC becomes 4 chars and is **silently chopped to 3**, dropping the last character of the tire code → a **wrong BC that still looks valid** and would match the wrong (or no) forecast recipe. The rebuild must either (a) reproduce the exact `char(3)` truncation for parity, **and** (b) assert/alarm if any component exceeds 1 char, because beyond 1 char the legacy itself is wrong. This is the single highest-severity latent defect in the proc.

### 3.2 The ×4 multiplier and ORDERS vs VEHICLES (VERIFIED)

```
$ -- sums by block over 06-18 COROLLA
GROUND | sum_VEHICLES=819 | sum_ORDERS=3276 (=819*4) | 9 BC rows
SPARE  | sum_VEHICLES=819 | sum_ORDERS= 819 (1:1)     | 3 BC rows
actual_vehicles_in_window = 819
```

`ORDERS = VEHICLES*4` exactly for ground; `ORDERS = VEHICLES` for spare. **Every vehicle is counted once in the ground block and once in the spare block** (819 + 819 = the doubled total) — ground and spare are two independent passes over the same vehicles, not additive across the union for vehicle-count purposes.

Downstream consumption (companion + Delphi):
- `VEHICLES` is the **quantity multiplier**: ASN line qty = `banker_round(VEHICLES * IN_ASSY_QTY * IN_TIRE_RATIO / 100)` (`DataModule.pas:5226`), or `VEHICLES * IN_ASSY_QTY` when ratio=100.
- `ORDERS` is the **No-Ratio gate**: the "(No Ratio)" small-volume branch fires on `Orders <= 5` (`DataModule.pas:5180`). Because ground ORDERS is ×4, a ground BC needs **≤1 vehicle** (1×4=4 ≤ 5) to trip No-Ratio, whereas a spare BC trips it at **≤5 vehicles**. The ×4 therefore changes *which* BCs are treated as small-volume — a real behavioral coupling the rebuild must keep: ground BC No-Ratio threshold ≈ 1 vehicle, spare BC ≈ 5 vehicles.

### 3.3 Spare `DataValue <> 'M'` exclusion (VERIFIED — 'M' = "no spare broadcast")

`M` is the GALC code for **a vehicle with no spare tire** ("Quick fix for no spare broadcast" per the inline comment). Excluding it prevents emitting a meaningless spare BC like `"NM "`.

```
$ -- spare values, 5-day window: only N and P, NO 'M' recently
[N] | 3604     [P] | 103
$ -- 'M' DOES occur historically (LineID 1, SPARETIRE, DataValue='M'), by month:
2020-02 | 5    2019-02 | 9    2019-01 | 5    2018-12 | 4    2018-11 | 2    2018-10 | 9 ...
```

So the exclusion is **real and load-bearing for historical/edge data** (it removed 2–9 spare rows/month in 2018–2020) even though it removes **0 rows in current production**. The rebuild MUST keep `spare value <> 'M'`. Note the filter is on the spare block only; a vehicle whose spare is 'M' still contributes its **ground** BC (it has wheels/tires), it just produces no spare BC — i.e. the union is **not** "drop the whole vehicle", only "drop this vehicle's spare row".

### 3.4 UNION dedup — can ground and spare BCs collide? (VERIFIED — disjoint in practice)

`UNION` (not `UNION ALL`) **deduplicates** identical `(BC, ORDERS, VEHICLES, '')` tuples. A ground row and a spare row are merged only if all four columns match. They cannot in current data because the BC namespaces are structurally disjoint:
- Ground BC = 3 non-space chars (`[NBB]`).
- Spare BC = 2 chars + a trailing space (`[NN ]`).

A collision would require a ground component to be empty/space (making the ground BC also 2-chars-padded). Proven impossible in-window:

```
$ -- count blank/null GROUND values, 5-day window
GROUNDTIRE  | blank_or_null=0 | 3707
GROUNDWHEEL | blank_or_null=0 | 3707
SPARETIRE   | blank_or_null=0 | 3707
```

Zero blank/null ground values → ground BCs are always 3 real chars → **trailing space is the namespace separator**, and the proc output confirms ground vs spare BCs never overlap. **Trap for the rebuild:** the distinction between ground and spare BC is **the char(3) trailing-space padding**. If a reimplementation trims the BC (or stores it in a varchar without padding), `"NN "` becomes `"NN"` and a (hypothetical, but possible if a ground value went blank) ground `"NN"` could then collide and a row would silently vanish under UNION-dedup. Preserve the right-padded `char(3)` semantics, including the trailing space. (Note: the downstream LIKE match — §7 — depends on exactly this padding.)

There is also a benign **same-block dedup** consequence: if two ground rows produced an identical `(BC, ORDERS, VEHICLES)` they could collapse — but they can't within one block because GROUP BY already makes BC unique per block and ORDERS=VEHICLES×4 is a function of VEHICLES. Cross-block, the trailing space keeps them distinct. So today UNION removes **0** rows; it behaves like UNION ALL. The rebuild may use UNION ALL **only if** it independently guarantees the namespaces stay disjoint.

### 3.5 GROUP BY collation/case (VERIFIED — case-insensitive)

`GROUP BY` is on the `char(3)` BC expression, whose components all carry `SQL_Latin1_General_CP1_CI_AS` (**case-insensitive, accent-sensitive**). So a hypothetical lowercase wheel value `b` would group with `B` into the same BC bucket. Current data is all uppercase, so no live effect, but a case-sensitive rebuild collation would split buckets and change counts. Preserve CI collation on the grouping key. (This also aligns with the downstream `LIKE` being CI — companion §1b.)

### 3.6 NULL ModelYearCode / NULL DataValue (VERIFIED — NULL → row dropped, two mechanisms)

- **NULL `DataValue`:** the value lives on a `vehicledata` row reached by an `INNER JOIN`. A vehicle with **no** GROUNDTIRE/GROUNDWHEEL/SPARETIRE row at all is dropped by the inner join (no row to count). If the row exists but `DataValue IS NULL`, the concat `m.ModelYearCode + NULL + …` yields **NULL BC** (string `+` propagates NULL pre-CONCAT), and `GROUP BY` collects all NULL-BC vehicles into a **single NULL bucket** — they are *counted*, but under a NULL BC that matches no forecast recipe downstream (`NULL LIKE pattern` → UNKNOWN → 0 rows), so they silently produce no ASN line. Measured: **0 blank/null ground or spare values in-window** (§3.4), so no live NULL bucket today.
- **NULL `ModelYearCode`:** same NULL-propagation → NULL BC. The live `Model` table has **no NULL year codes** (all 15 are single letters), but the column is `NULLable`, so this is a latent silent-drop.

**Trap:** a NULL in any BC component does not error and does not crash the count — it quietly routes vehicles into a non-matching NULL BC. The rebuild should treat a NULL BC component as a data-integrity alarm, not absorb it.

### 3.7 JOIN cardinality — does vdX fan out and inflate COUNT? (VERIFIED — exactly one row per vehicle per description; NOT enforced, only convention)

This is the critical correctness question, because the join matches `DataItem` **by description**, and many `DataItemID`s share each description (each model variant has its own GROUNDTIRE/GROUNDWHEEL/SPARETIRE `DataItem` row). If a vehicle had VehicleData rows under two GROUNDTIRE `DataItemID`s, `COUNT(*)` would double.

Proven — per-vehicle row counts per description, 1-day window:

```
$ SELECT desc, MAX(cnt) max_per_veh, MIN(cnt) min_per_veh, SUM(cnt>1) veh_with_multi, COUNT(*) vehicles
    FROM (per-vehicle-per-description COUNT(*)) ...  (COROLLA 06-18)
GROUNDTIRE  | 1 | 1 | 0 | 819
GROUNDWHEEL | 1 | 1 | 0 | 819
SPARETIRE   | 1 | 1 | 0 | 819
```

`min = max = 1`, **zero** vehicles with multiple rows, across 819 vehicles. Confirmed stable over a 5-day window and a heavy historical month (Feb 2019, 8742 vehicles — `HAVING COUNT(*) > 1` returned empty). And the description-only join still resolves correctly even though vehicles span model variants:

```
$ -- GROUNDTIRE rows for 06-18 come from TWO DataItemIDs (two models), still 1 per vehicle:
DataItemID 99  (ModelID 14) | 818 vehicles
DataItemID 106 (ModelID 15) |   1 vehicle      (818 + 1 = 819 = vehicle count)
```

So no fan-out today. **BUT `VehicleData` is a HEAP with no unique constraint** (proven via `sys.indexes`: only non-unique helper indexes `idx_VehicleID`, `idxDataItemVID`; no PK/unique). The one-row-per-vehicle-per-description invariant is **data convention, not schema-enforced** — a double-write (e.g. a re-broadcast) would double that vehicle's `COUNT(*)`, doubling both VEHICLES and ORDERS for its BC, inflating every downstream ASN qty for that BC. The rebuild should either reproduce the raw join (accepting the fan-out risk) **or** defensively pivot to one value per (vehicle, attribute) and alarm on duplicates.

### 3.8 Date window inclusivity + the datetime time-component (VERIFIED — the midnight trap)

Filter is `v.DateCreated >= @begindate AND v.DateCreated <= @enddate` — **both bounds inclusive (`>=` / `<=`)**. `DateCreated` carries a **real time-of-day** (proven: 06-18 vehicles run `00:01:19.333` … `23:59:02.653`).

```
$ -- end at the day's midnight (00:00:00) -> entire day lost:
end_at_midnight              | total_veh = 0    | bc_rows = 0
$ -- end at NEXT day's midnight (inclusive) -> full day + the boundary instant:
end_next_midnight_inclusive  | total_veh = 1638 | bc_rows = 12
```

**THE MIDNIGHT TRAP:** calling with `@enddate = '<day> 00:00:00'` returns **0 rows** — because every vehicle is stamped *after* midnight. To capture a full day the caller must pass an end like `'<day> 23:59:59.997'` (datetime's max sub-second tick) or `'<next-day> 00:00:00'` with `<=` knowing the boundary instant leaks in. **Because the upper bound is inclusive (`<=`), a vehicle stamped at exactly the `@enddate` instant is included** — so a half-open caller convention (`< next midnight`) is NOT what this proc implements; two adjacent day-windows that share an endpoint would **double-count** any vehicle on that exact tick. The rebuild must replicate inclusive-both-ends semantics, and the Delphi caller's exact `@enddate` string is part of the contract (verify what `CalculateASNFRS` passes; if it passes a bare date the proc silently returns nothing).

### 3.9 @Start / @Last are inert (VERIFIED)

The body never references `@Start` or `@Last`. Proven by running identical windows with absurd values:

```
$ AD_FRSPULL ... @Start=0,     @Last=0      -> 1638 veh, 12 rows
$ AD_FRSPULL ... @Start=99999, @Last=-99999 -> 1638 veh, 12 rows   (identical)
```

The FRS pull is **purely date+line scoped**. (This corrects the earlier M1-spec inference that S/E were an ASN sequence range.) The rebuild can drop these two params.

---

## 4. UNION ordering / determinism

The final `ORDER BY BC` makes the **row order deterministic** (lexicographic by the char(3) BC, with the trailing-space spare codes sorting among the ground codes — e.g. `NN ` sorts after `NJJ` and before `NP ` because space < letters). Within each block, GROUP BY guarantees one row per BC. So unlike `SELECT_ForecastDetailBCASN` (which is order-nondeterministic over heaps), **AD_FRSPULL's row order is stable**. No determinism trap here, but note the BC sort key includes the trailing space — preserve it if the rebuild relies on output order.

---

## 5. Worked example (real line + 1-day window)

**Call:** `EXEC AD_FRSPULL @begindate='2026-06-18 00:00:00', @enddate='2026-06-18 23:59:59.997', @Start=0, @Last=0, @LineName='COROLLA'`
(819 COROLLA vehicles built that day.)

| BC (braced) | kind | VEHICLES = COUNT(*) | ORDERS | decode |
|---|---|---|---|---|
| `[NBB]` | ground | 507 | 2028 (=507×4) | MY=**N**, WHEEL=**B**, TIRE=**B** |
| `[NEE]` | ground | 194 | 776 (=194×4) | MY=**N**, WHEEL=**E**, TIRE=**E** |
| `[PEE]` | ground | 1 | 4 (=1×4) | MY=**P**, WHEEL=**E**, TIRE=**E** |
| `[NN ]` | spare | 798 | 798 (1:1) | MY=**N**, SPARE=**N** |
| `[NP ]` | spare | 20 | 20 (1:1) | MY=**N**, SPARE=**P** |
| `[PN ]` | spare | 1 | 1 (1:1) | MY=**P**, SPARE=**N** |

(Full output also had ground `NCC`/`NDD`/`NFF`/`NGG`/`NHH`/`NJJ`.)

**Reconciliation of the E-tire vehicles** — note `NEE`=194 while the raw GROUNDTIRE='E' count for the day was 195:

```
$ -- vehicles with WHEEL=E AND TIRE=E for 06-18, by model-year code:
N | 194      P | 1
```

194 of the E/E vehicles are model-year **N** → BC `NEE`; the 195th is model-year **P** → BC `PEE` (which indeed shows VEHICLES=1). The ModelYearCode prefix splits one wheel/tire combination across separate BCs. Block totals reconcile exactly: ground sum_VEHICLES = spare sum_VEHICLES = 819 = actual vehicles built (each vehicle once in ground, once in spare).

**End-to-end transform for one BC** (e.g. `NBB`, VEHICLES=507): the fan-out passes `@BCode='NBB'` to `SELECT_ForecastDetailBCASN`; each matched forecast/recipe row yields an ASN line qty = `banker_round(507 * IN_ASSY_QTY * IN_TIRE_RATIO / 100)` (ratio=100 → exactly `507 * IN_ASSY_QTY`), unless `ORDERS(=2028) <= 5` (false here) which would take the "(No Ratio)" single-assy branch.

---

## 6. How the char(3) BC feeds the downstream LIKE (cross-proc contract)

The companion `SELECT_ForecastDetailBCASN-analysis.md` proved the BC→forecast match is **`@BCode LIKE VC_BROADCAST_CODE`** where the *column* is a character-class pattern (`[KLM]CC`), the inbound BC is the literal left operand, and the match is case-insensitive. Two padding facts must line up:

1. **AD_FRSPULL emits `char(3)`** — ground codes are 3 real chars, **spare codes are 2 chars + a trailing space** (`"NN "`). The Delphi caller reads it as `ALC_StoredProc.FieldByName('BC').AsString` (`DataModule.pas:5151`) and passes it as `@BCode varchar(20)` (companion §1). T-SQL `LIKE` **trims trailing spaces on neither side for pattern semantics but ANSI-pads for `=`**; for `LIKE`, a trailing space in `@BCode` is a *literal space to match*. So a spare BC `"NN "` (with trailing space) tested against a 2-char pattern like `[MN]N` would **fail** unless the stored pattern also accounts for the trailing space or the caller trims it. The companion observed the actual stored ASN BC patterns are right-padded to `char(21)`; the load-bearing point for the rebuild is: **whatever trimming/padding the legacy applies to the BC between AD_FRSPULL and the LIKE must be reproduced exactly, because the spare BC's trailing space is significant under LIKE.** (Confirm in `DataModule.pas` whether `AsString` trims — Delphi `TField.AsString` on a fixed CHAR typically returns the value *with* trailing blanks unless `.AsString` trims; this is the one cross-proc seam to nail in code, flagged for delphi-architect.)
2. The ground/spare disjointness (§3.4) means a spare BC will only ever match a spare-shaped pattern and a ground BC a ground-shaped pattern — the trailing space is part of that routing.

**Inferred (needs the Delphi-side confirmation):** exactly where the `char(3)` is trimmed/repadded to the `char(21)` LIKE operand. Verified facts: the proc emits `char(3)`, spare carries a trailing space, the downstream match is a CI `LIKE` with the column as pattern.

---

## 7. What a faithful reimplementation MUST reproduce (and the traps that silently break it)

1. **The exact BC formulas + concat order.** Ground = `char(3)`-truncated `ModelYearCode + GROUNDWHEEL + GROUNDTIRE`; spare = `char(3)`-truncated `ModelYearCode + SPARETIRE`. WHEEL precedes TIRE; the alias `vd1`=TIRE/`vd2`=WHEEL is *reverse* of read order. (VERIFIED)
2. **char(3) right-padding is semantic.** Ground = 3 real chars; spare = 2 chars **+ trailing space**. The trailing space (a) separates the ground vs spare namespaces and (b) flows into the downstream `LIKE`. Trimming or using an unpadded varchar can both **collide rows under UNION-dedup** and **break the LIKE match**. (VERIFIED)
3. **char(3) silently truncates — latent wrong-BC bug.** Safe only while every component is 1 char. If GALC ever emits a multi-char wheel/tire/spare value, the BC is silently chopped. Reproduce the truncation for parity AND alarm when any component > 1 char. (VERIFIED width=1; truncation is `CONVERT(char(3))` semantics)
4. **Multipliers:** ground `ORDERS = VEHICLES*4` (4 corners), spare `ORDERS = VEHICLES`. `VEHICLES` is the qty multiplier; `ORDERS` is the `<= 5` No-Ratio gate — so the ×4 makes ground No-Ratio fire at ~1 vehicle, spare at ~5. Get the ×4 wrong and the No-Ratio branch fires for the wrong BCs. (VERIFIED)
5. **Spare `<> 'M'` exclusion** — 'M' = no-spare vehicle. Removes 0 rows today but 2–9/month historically; keep it. It drops only the *spare* row, not the whole vehicle. (VERIFIED)
6. **Date window is inclusive both ends with a real time-of-day → the midnight trap.** `@enddate = '<day> 00:00:00'` returns ZERO rows. Must pass an end-of-day/next-midnight value; inclusive `<=` means the boundary instant double-counts across adjacent windows. Replicate inclusive-both-ends and the caller's exact end value. (VERIFIED 0 rows at midnight)
7. **One VehicleData row per vehicle per attribute is convention, not enforced** (VehicleData is a heap, no unique key). No fan-out today, but a duplicate would silently double VEHICLES/ORDERS and every downstream qty for that BC. Defensively pivot + alarm on duplicates. (VERIFIED no fan-out across 5 days + a heavy historical month)
8. **NULL in any BC component → NULL BC** (string-concat NULL propagation), routed into a single non-matching bucket, silently producing no ASN line. ModelYearCode and DataValue are both NULLable. Treat NULL components as data alarms. (VERIFIED no NULLs in-window; mechanism from concat semantics)
9. **Case-insensitive collation** on the GROUP BY / BC key (`CI_AS`). A case-sensitive rebuild would split buckets and change counts. (VERIFIED collation)
10. **UNION dedups (not UNION ALL).** Removes 0 rows today only because ground/spare namespaces are disjoint via the trailing space; UNION ALL is safe only if that disjointness is independently guaranteed. (VERIFIED)
11. **Drop `@Start`/`@Last`** — inert. Output is purely (line, date-window) scoped. Row order is deterministic via `ORDER BY BC`. (VERIFIED inert)

**Top 4 traps most likely to silently corrupt revenue-critical ASN qty:** (a) the **midnight/inclusive date boundary** — wrong `@enddate` → zero or double counts; (b) **char(3) trailing-space padding** of spare BCs feeding the LIKE — trimming breaks the match; (c) the **×4 ground multiplier** mis-gating the No-Ratio branch; (d) the **un-enforced one-row-per-vehicle invariant** doubling counts on a duplicate VehicleData write. None of the four throws an error — they all just emit a wrong number.
