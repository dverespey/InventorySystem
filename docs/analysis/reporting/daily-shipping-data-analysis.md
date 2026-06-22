# M3 Data Spec: Daily Shipping (Range) Report

**Proc:** `REPORT_DailyShippingRange`  **Area:** Reporting / shipping  **Analyst:** Claude / 2026-06-21
**Goal:** decode the exact numbers the Daily Shipping Range report emits so the rebuild reproduces them
(M3 gate = "report numbers match legacy"). This is the failing path the M3 plan prioritizes.

**Sources of truth:** live `Inventory` DB on `mssql-spike` (body verified via `OBJECT_DEFINITION`) +
schema dump `/tmp/inv_utf8.sql` (regen: `iconv -f UTF-16LE -t UTF-8 "DB Schema/CreateInventory.sql"`).
Live body == dump body (no drift). Consumer: `MainMenu.pas:2919` `DailyShippingRangeTireWheelPartNumbersClick`.

> Line citations to the proc body are `/tmp/inv_utf8.sql` (live dump). The older `reporting.md` catalog
> cites the *superseded* schema by line — both describe the same proc.

---

## 1. Purpose (as titled vs as built)

Menu label / Excel title (`MainMenu.pas:2955`): **"Daily Shipping Range(Tire/Wheel Part Numbers)"** — i.e.
"for a production-date range, list the parts shipped and their total quantities."

As **built**, the proc returns a per-part roll-up of `INV_PART_SHIPPING_INF.IN_QTY` over a production-date
range — but it inner-joins to `INV_PARTS_STOCK_MST`, so it only ever returns parts that are *registered in
the parts master*. On the live data that is **3 consumable parts (FILM + VALVE), and ZERO tire/wheel parts**
— the exact opposite of the title. See §5 (failure cause).

It is a **pure SELECT** (read-only; no DML — verified §6).

---

## 2. The SELECT (verbatim, `/tmp/inv_utf8.sql:2838-2862`)

```sql
CREATE PROCEDURE [dbo].[REPORT_DailyShippingRange]
    @BeginPdate varchar(8),
    @EndPdate   varchar(8)
AS
BEGIN
    SET NOCOUNT ON;
    SELECT   SUM(s.IN_QTY) 'Vehicle Count',
             m.VC_PART_NUMBER 'Part Number',
             m.VC_PARTS_NAME 'Desc',
             SUM(p.IN_QTY) 'PQty'
    FROM INV_SHIPPING_INF s
         JOIN INV_PART_SHIPPING_INF p
             ON s.VC_PRODUCTION_DATE = p.VC_PRODUCTION_DATE
         JOIN INV_PARTS_STOCK_MST m
             ON p.VC_PART_NUMBER = m.VC_PART_NUMBER
    WHERE s.VC_PRODUCTION_DATE between @BeginPDate and @EndPDate
    GROUP BY m.in_part_type_ID, m.VC_PART_NUMBER, m.VC_PARTS_NAME
    ORDER BY m.in_part_type_ID, m.VC_PART_NUMBER
END
```

### 2.1 Lineage / joins

| Table | Alias | Role | Join key | Card. (live) |
|---|---|---|---|---|
| `INV_SHIPPING_INF` | `s` | one row per production date = the day's vehicle/seq header | — | 82 rows / 82 distinct dates → **exactly 1 row per date** |
| `INV_PART_SHIPPING_INF` | `p` | parts consumed per production date | `s.VC_PRODUCTION_DATE = p.VC_PRODUCTION_DATE` | 886 rows / 82 dates / **13 distinct part numbers** |
| `INV_PARTS_STOCK_MST` | `m` | parts master (gate + description + type) | `p.VC_PART_NUMBER = m.VC_PART_NUMBER` | 47 rows, `VC_PART_NUMBER` unique |

Join keys are all `varchar` and `NOT NULL` (`VC_PRODUCTION_DATE varchar(8)`, `VC_PART_NUMBER varchar(12)`).
`s.IN_QTY` is `int NULL`; `p.IN_QTY` is `int NOT NULL`; `m.IN_QTY` not selected.

**Cardinality of the s↔p join:** because `s` has exactly 1 row per date, the join fans each `p` row out by
**factor 1** on the live data — there is *no* row multiplication today. (Proven: §3.) This is fragile — see §4.1.

### 2.2 Projection (the 4 result columns)

| Col alias | Expression | Type | Consumed by Delphi? |
|---|---|---|---|
| `Vehicle Count` | `SUM(s.IN_QTY)` | int | **NO** — discarded (`MainMenu.pas:2976-2978` writes only 3 cols) |
| `Part Number` | `m.VC_PART_NUMBER` | varchar(12) | YES → Excel col 1 |
| `Desc` | `m.VC_PARTS_NAME` | varchar(50) | YES → Excel col 2 |
| `PQty` | `SUM(p.IN_QTY)` | int | YES → Excel col 3 (`fieldbyname('PQTY')`) |

**The Excel writer (`MainMenu.pas:2973-2981`):**
```pascal
mysheet.Cells[z,1].value := fieldbyname('Part Number').AsString;
mysheet.Cells[z,2].value := fieldbyname('Desc').AsString;
mysheet.Cells[z,3].value := fieldbyname('PQTY').AsString;
```
So the report's *visible* output is **3 columns: Part Number, Desc, PQty.** The `Vehicle Count` SUM is
computed by the proc but never written — a rebuild must reproduce **Part Number / Desc / PQty** exactly; it
need NOT reproduce `Vehicle Count` for the Excel report (though faithful proc parity should still emit it).

### 2.3 Grouping / aggregation / sort

- `GROUP BY m.in_part_type_ID, m.VC_PART_NUMBER, m.VC_PARTS_NAME` → one output row **per master part**.
- `SUM(p.IN_QTY)` = total parts shipped for that part across the date range.
- `SUM(s.IN_QTY)` = sum of the day's vehicle count once per *surviving (s,p) join row* → a **fan-out
  artifact** (see §4.1), benign only because card(s)=1/date.
- `ORDER BY m.in_part_type_ID, m.VC_PART_NUMBER` — type then part number. `in_part_type_ID` is `int`
  (`1=TIRE,2=WHEEL,3=FILM,4=VALVE,5=MISC`, from `INV_PART_TYPE_MST`), so FILM (3) sorts before VALVE (4).
- No rounding anywhere — all `int`; `SUM(int)` stays `int`. No `money`/`decimal`/`ROUND`.

### 2.4 Params (`@BeginPdate`, `@EndPdate`, both `varchar(8)` `yyyymmdd`)

Delphi builds them from the date-picker by stripping the `/` from `yyyy/mm/dd`
(`MainMenu.pas:2942-2944`): `copy(d,1,4)+copy(d,6,2)+copy(d,9,2)`. The picker is fed by
`REPORT_AvailableProductionDates @Line, @INVOICE=0, @ASN=0` (the `INV_SHIPPING_INF` branch). The WHERE
filters **`s.VC_PRODUCTION_DATE`** only (the header date), not `p`'s — but since the join is on that same
column they coincide.

---

## 3. Live-data proof (numbers)

**Full range `'20111028'`–`'20120312'` — the entire dataset (`EXEC REPORT_DailyShippingRange`):**

```
Vehicle Count   Part Number     Desc                    PQty
14832           478930201000    RED FILM                27796
14829           478930204000    BLUE FILM               31532
14832           900804500600    REGULAR RUBBER VALVE    14832
```
→ **3 rows.** (FILM type 3 before VALVE type 4, then by part number.)

**PQty is CORRECT** — clean per-part `SUM(IN_QTY)` over `INV_PART_SHIPPING_INF` (no shipping join) matches
exactly:

```sql
SELECT p.VC_PART_NUMBER, SUM(p.IN_QTY) FROM INV_PART_SHIPPING_INF p
WHERE p.VC_PART_NUMBER IN ('478930201000','478930204000','900804500600')
  AND p.VC_PRODUCTION_DATE BETWEEN '20111028' AND '20120312'
GROUP BY p.VC_PART_NUMBER;
-- 478930201000 -> 27796   478930204000 -> 31532   900804500600 -> 14832  (== proc PQty)
```

**Single date `'20111101'` (12 part rows that day) — shows the silent drop:**
```
proc output:
124   478930201000  RED FILM               256
124   478930204000  BLUE FILM              240
124   900804500600  REGULAR RUBBER VALVE   124

INV_PART_SHIPPING_INF rows on 20111101 (matched flag):
426070601100=496 DROPPED   426110248000=124 DROPPED   426110288000=252 DROPPED
4261102A0000=4   DROPPED   4261102D4000=240 DROPPED   426250107200=124 DROPPED
4262502F2000=81  DROPPED   4262502F3000=171 DROPPED   4262502F5000=244 DROPPED
478930201000=256 MATCH     478930204000=240 MATCH     900804500600=124 MATCH
INV_SHIPPING_INF on 20111101: seq 0224->0347, IN_QTY=124
```
**9 of 12 part rows dropped** (all the `426...` tire/wheel parts); 3 consumables survive. `Vehicle
Count=124` here equals the true day vehicle count *only because* card(s)=1/date (fan factor 1).

**Boundary inclusivity (BETWEEN is end-inclusive):** `EXEC … '20120312','20120312'` returns the last
date's 3 rows (`1112 / 776 / 472`) → start- and end-inclusive. Lexical `varchar` compare is chronologically
safe for `yyyymmdd` (`'20120312' > '20111028'` = true).

---

## 4. Edge-case matrix (proven)

| # | Hazard | Behavior | Proof |
|---|---|---|---|
| 4.1 | **Fan-out via `SUM(s.IN_QTY)`** | `Vehicle Count` = `s.IN_QTY` summed once per surviving `(s,p)` pair. Benign today (1 `s` row/date → factor 1), but **if any production date ever gets 2+ `INV_SHIPPING_INF` rows, BOTH `Vehicle Count` AND `PQty` inflate** (PQty would double-count each part). | `HAVING COUNT(*)>1` on `INV_SHIPPING_INF` GROUP BY date → empty (every date has exactly 1 row). 82 rows / 82 distinct dates. |
| 4.2 | **Inner-join orphan drop (THE FAILURE)** | The `JOIN INV_PARTS_STOCK_MST` is an inner join. Part rows whose `VC_PART_NUMBER` is absent from the master are silently dropped. | `641 / 886` part-shipping rows orphaned; only **3 of 13** distinct part numbers exist in the master. |
| 4.3 | **Title vs content mismatch** | Report is titled "Tire/Wheel Part Numbers" but the surviving 3 parts are types **3=FILM, 4=VALVE** — no tires/wheels at all. The `426...` tire/wheel parts live in `INV_FORECAST_DETAIL_INF` (as `VC_TIRE/WHEEL_PART_NUMBER_CODE`) but are **not registered in `INV_PARTS_STOCK_MST`**. | `INV_PART_TYPE_MST`: 1=TIRE,2=WHEEL,3=FILM,4=VALVE,5=MISC. Matched parts → types 3,3,4. `426...` codes present in `INV_FORECAST_DETAIL_INF`, absent from master. |
| 4.4 | **`Vehicle Count` discarded by consumer** | Delphi writes only Part Number / Desc / PQty. The fan-out column never reaches Excel. | `MainMenu.pas:2976-2978`. |
| 4.5 | **NULL in aggregate** | `s.IN_QTY` is `int NULL`; `SUM` ignores NULLs (no error, no NULL-poisoning of the sum). `p.IN_QTY` is `NOT NULL`. `VC_PARTS_NAME` is `NULL`-able but is a GROUP BY key, not aggregated. | schema: `INV_SHIPPING_INF.IN_QTY` nullable; `INV_PART_SHIPPING_INF.IN_QTY` not null. |
| 4.6 | **BETWEEN inclusivity** | start- and end-inclusive on `varchar(8)`. | §3 single-boundary test. |
| 4.7 | **Collation / case on join key** | join on `VC_PART_NUMBER` uses DB default collation; all live values are digits so case is moot, but a case-insensitive collation would matter if alpha codes appear. | (noted; not load-bearing on this data) |
| 4.8 | **Empty result → "No daily records"** | If 0 rows, Delphi shows "No daily records" (`MainMenu.pas:3002`). A range with only tire parts shipped (all orphaned) would yield 0 rows and look like "nothing shipped." | proc returns `recordcount=0` when no master parts match. |

**No D6 window-blindness here** (this proc does not touch `INV_MANIFEST_COST_MST`; no price). No mutation.
No self-flip. It is a clean read — its defect is the orphan inner join, not a window bug.

---

## 5. Why the legacy path "fails" (the M3 priority)

The report does **not** error and does **not** depend on a missing proc — `REPORT_DailyShippingRange`
exists, runs, and returns rows. The failure is a **data-fidelity / semantic failure**, with two layers:

1. **Primary — silent orphan drop (§4.2).** The inner join to `INV_PARTS_STOCK_MST` discards **72% of
   part-shipping rows** (641/886) and **10 of 13 part numbers**. Every tire/wheel part shipped is dropped
   because tire/wheel parts are not rows in the parts-stock master. A report titled "Tire/Wheel Part
   Numbers" returns **zero tire/wheel parts** — it returns only RED FILM, BLUE FILM, REGULAR RUBBER VALVE.
   This is almost certainly *the* reason the legacy report is the failing/distrusted path: its numbers don't
   reconcile to what was actually shipped.

2. **Secondary — fan-out fragility (§4.1).** `SUM(s.IN_QTY)` (and, if `s` ever has >1 row/date, `SUM(p.IN_QTY)`
   too) is correct only by accident of the current 1-row-per-date data. A naive reimplementation that
   "fixes" the join could silently introduce or remove this multiplication and diverge.

**Rebuild decision implication (for ignition-architect, not decided here):** a faithful port reproduces 3
rows / the FILM+VALVE numbers above. A *corrected* report (to actually show tire/wheel parts) must change
the lineage — drop the master inner join (or LEFT JOIN it for description, sourcing tire/wheel descriptions
from `INV_FORECAST_DETAIL_INF` / a tire/wheel master). That is a behavior change, not parity, and needs a
David decision. M3 gate "numbers match legacy" = match the 3-row faithful output.

---

## 6. Read-only confirmation

`REPORT_DailyShippingRange` contains no `INSERT`/`UPDATE`/`DELETE` (keyword scan of `OBJECT_DEFINITION`
across all `REPORT_*` procs → `read-only`). No trigger fires on a SELECT. Re-running is idempotent and safe.
(Two *other* REPORT procs DO mutate on read — see the survey doc §"Mutate-on-read", `REPORT_EDI810`/`856`.)

---

## 7. What a faithful reimplementation MUST reproduce (and the traps)

**MUST reproduce:**
- 3 output columns the Excel report consumes: `Part Number`, `Desc`, `PQty` (PQty = `SUM(INV_PART_SHIPPING_INF.IN_QTY)`).
- Only parts present in `INV_PARTS_STOCK_MST` (the inner join) — i.e. the orphan drop is part of legacy behavior.
- `int` arithmetic, no rounding; sort by `IN_PART_TYPE_ID` then `VC_PART_NUMBER`.
- BETWEEN end-inclusive on `yyyymmdd` strings; filter on the production-date (header) field.
- Empty result → caller's "No daily records".

**Traps that will silently break parity:**
- **Replacing the inner join with a LEFT/outer join** → suddenly returns the 10 tire/wheel parts → numbers
  diverge from legacy (more rows, descriptions NULL). Faithful = keep it inner.
- **De-fanning vs re-fanning `IN_QTY`.** If the Named Query joins `INV_SHIPPING_INF` once per part-shipping
  row (faithful) PQty is fine *today*; if the data ever gets >1 `s` row/date, both columns inflate. A
  rebuild that aggregates `p` independently of `s` (cleaner) would *not* inflate — and would silently
  diverge from legacy the day a duplicate `s` row appears. Decide explicitly.
- **Emitting `Vehicle Count` as a "real" vehicle total.** It is a join-fanned sum; on this data it equals
  the day count, but it is not a robust vehicle count and the Excel report ignores it.
- **`@Begin/@EndPdate` param binding.** ADO binds positionally; a Named Query binds by name — keep the two
  params in declared order or rename consistently (general report-family caveat, `reporting.md §4.2`).
- **NULL `s.IN_QTY`** is ignored by SUM (don't coalesce to 0 in a way that changes a count of contributing rows).
