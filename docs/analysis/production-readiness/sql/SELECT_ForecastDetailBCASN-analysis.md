# SELECT_ForecastDetailBCASN — authoritative SQL-semantics analysis

**Proc:** `dbo.SELECT_ForecastDetailBCASN` (DB: `Inventory`)
**Role:** broadcast-code (BC) → ASN parts + ratios lookup, called once per BC row in the
ASN fan-out (`AD_FRSPull` → per-BC → `INSERT_ASNDetail`).
**Dump:** `~/tmp/inv_utf8.sql:3011` (`InventorySystem/DB Schema/CreateInventory.sql`, UTF-16LE).
**Delphi caller:** `DataModule.pas:5149` (the ASN-create fan-out).
**Verification:** every claim below was proven on the live `Inventory` / `Inventory_Live`
containers (`mssql-spike`). Verified-vs-inferred is stated per item.

---

## 0. Live body (verified — live == Inventory_Live == dump, byte-identical)

`OBJECT_DEFINITION` from both `Inventory` and `Inventory_Live` returned the identical body;
it matches the dump. **No proc drift.**

```sql
CREATE PROCEDURE [dbo].[SELECT_ForecastDetailBCASN]
    @BCode    varchar(20),
    @EffMonth varchar(7)
AS
BEGIN
    SET NOCOUNT ON;
    SELECT *
    FROM INV_FORECAST_DETAIL_INF f
        LEFT JOIN INV_MANIFEST_COST_MST c
        ON f.VC_ASSY_PART_NUMBER_CODE = c.VC_ASSY_PART_NUMBER_CODE
    WHERE @BCode LIKE VC_BROADCAST_CODE
    AND ((VC_EFFECTIVE_MONTH = @EffMonth or VC_EFFECTIVE_MONTH = '')
    AND IN_TIRE_RATIO <> 0
    AND IN_WHEEL_RATIO <> 0 )
END
```

Three semantics in this 9-line body dominate the output and **all three are traps**:
`@BCode LIKE VC_BROADCAST_CODE` (column is the *pattern*), no `ORDER BY` over a heap,
and `SELECT *` across a LEFT JOIN producing duplicate column names + NULL cost fields.

---

## 1. Params + how it is keyed by BC

| Param | Type | Notes |
|---|---|---|
| `@BCode` | `varchar(20)` | the inbound broadcast code from `AD_FRSPull` (Delphi passes `ALC_StoredProc.FieldByName('BC').AsString`, `DataModule.pas:5151`). **NOT char(3)** — the task hypothesis is wrong; both param and column are varchar. |
| `@EffMonth` | `varchar(7)` | Delphi builds `YYYY/MM` from production date (`DataModule.pas:5153`). **Dead in current data** — see §1c. |

### 1a. The match is `@BCode LIKE VC_BROADCAST_CODE` — the COLUMN is the pattern (VERIFIED)

This is the single most important semantic. The parameter is the **left** operand (the
string being tested); the stored column `VC_BROADCAST_CODE` is the **right** operand (the
`LIKE` pattern). The stored BCs are deliberately authored as `LIKE` patterns with
character-class wildcards:

```
$ ... -Q "SELECT DISTINCT '['+VC_BROADCAST_CODE+']' FROM INV_FORECAST_DETAIL_INF ..."
[[J]EE]        [[KLM]CC]     [[KLM]JJ]     [[MNP]AA]    [[MNP]JJ]    [[N][F][F]]
[[K][FS][FS]]  [[KLM]DD]     [[KLM]KK]     [[MNP]BB]    [[MNP]KK]    []
[[KL]AA]       [[KLM]EE]     [[KLM]LL]     [[MNP]CC]    [[MNP]LL]
[[KLM][BW][BW]][[KLM]GG]     [[M]FF]       [[MNP]DD]    [[MNP]P]
[[KLM][FR][FR]][[KLM]HH]     [[M]QQ]       [[MNP]EE]    [[MNP][N]]
                                           [[MNP]GG]    [[MNP]HH]
```

`[KLM]CC` is *one char in the class {K,L,M} followed by literal "CC"*. So an inbound
`KCC`, `LCC`, or `MCC` all resolve to the same recipe rows. Proven end-to-end through the proc:

```
$ EXEC SELECT_ForecastDetailBCASN @BCode='KCC', @EffMonth='2026/06';
157|42670FEU0000|...|[KLM]CC|...   (tire 20)
163|42670FET9000|...|[KLM]CC|...   (tire 40)
164|42670FEU1000|...|[KLM]CC|...   (tire 40)
$ SELECT ... WHERE 'NBB' LIKE VC_BROADCAST_CODE;
[[MNP]BB]|42600FEK5000
[[MNP]BB]|42600FEK6000
[[MNP]BB]|42600FEK7000
```

**Direction-of-LIKE trap for the rebuild:** because the wildcard lives in the *trusted*
stored column and the inbound BC is the *literal*, this is safe today. But a naive
Ignition Named Query that "normalizes" this to `VC_BROADCAST_CODE = :bcode` (equality)
**will return zero rows for every patterned BC** — i.e. 28 of 29 stored BCs. The match
MUST stay `:bcode LIKE VC_BROADCAST_CODE`. Equally, if the rebuild ever lets a user-typed
BC reach the *pattern* side, `[`, `%`, `_` become injection. Keep inbound on the left.

### 1b. Collation / case / trailing space on the BC match (VERIFIED)

- `VC_BROADCAST_CODE` collation = `SQL_Latin1_General_CP1_CI_AS` → **case-insensitive**.
  `kcc` would match `[KLM]CC` just as `KCC` does. A case-sensitive rebuild collation would
  break parity.
- Trailing-space: T-SQL `LIKE` / `=` trim trailing spaces on the right operand per ANSI
  padding; inbound `'KCC '` still matches. (Not a live problem — `AD_FRSPull` BCs are clean —
  but a flagged equivalence.)

### 1c. The `@EffMonth` filter is effectively DEAD in current data (VERIFIED)

All 50 forecast rows store a **single space** as the effective month, and `' ' = ''` is
TRUE in T-SQL (trailing-space trim):

```
$ SELECT DATALENGTH(VC_EFFECTIVE_MONTH), COUNT(*) ... GROUP BY DATALENGTH(...);
1|50                                  -- every row is 1 byte = ' '
$ SELECT CASE WHEN ' ' = '' THEN 'space_eq_empty_TRUE' ... ;
space_eq_empty_TRUE
$ SELECT COUNT(*) ... WHERE (VC_EFFECTIVE_MONTH='2026/06' OR VC_EFFECTIVE_MONTH='');
50                                    -- all rows pass regardless of @EffMonth
```

So the `OR VC_EFFECTIVE_MONTH = ''` arm passes **all** rows for **any** `@EffMonth`. The
month parameter currently filters nothing. *Inferred:* the column was designed to date-bound
recipe versions, but no version has ever been month-stamped — the recipe is treated as
always-effective. The rebuild can keep the clause for forward-compat but must NOT assume the
month ever narrows results.

### 1d. Can a BC return ZERO rows → the fan-out's no-fc abort (VERIFIED)

Yes. An inbound BC matching no stored pattern returns 0 rows:

```
$ EXEC SELECT_ForecastDetailBCASN @BCode='ZZZ', @EffMonth='2026/06';  -- (0 rows)
rowcount_ZZZ = 0
```

Delphi guards this: `if Inv_StoredProc.recordcount > 0 then ...` (`DataModule.pas:5155`).
A BC with no recipe is silently skipped (no ASN line, no error). The rebuild must preserve
"zero rows = skip this BC", not "zero rows = error".

---

## 2. Ratio columns — IN_TIRE_RATIO / IN_WHEEL_RATIO / IN_ASSY_QTY + cost fields

### Types (VERIFIED via sys.columns)

| Column | Type | Nullable | Source table |
|---|---|---|---|
| `IN_TIRE_RATIO` | `int` | **YES** | INV_FORECAST_DETAIL_INF |
| `IN_WHEEL_RATIO` | `int` | **YES** | INV_FORECAST_DETAIL_INF |
| `IN_RATIO` | `int` | **YES** | INV_FORECAST_DETAIL_INF |
| `IN_ASSY_QTY` | `int` | **YES** | INV_FORECAST_DETAIL_INF |
| `VC_ASSY_MANIFEST_NUMBER` | `varchar(2)` | NO (in cost tbl) → NULL via LEFT JOIN miss | INV_MANIFEST_COST_MST |
| `IN_MANIFEST_COST_ID` | `int` | NO (PK) → **NULL via LEFT JOIN miss** | INV_MANIFEST_COST_MST |

All ratios are **int**, so there is no fractional ratio and no rounding *inside the proc*;
all rounding happens Delphi-side (§2d).

### 2a. Ranges / NULLs (VERIFIED)

```
$ tire_minmax  = 20 .. 100      wheel_minmax = 20 .. 100      assyqty_minmax = 1 .. 4
$ tire_null=0  wheel_null=0  tire_zero=0  wheel_zero=0
```

No NULL or zero ratios in current data → the `IN_TIRE_RATIO <> 0 AND IN_WHEEL_RATIO <> 0`
filter excludes nothing today. **But the filter is a silent-drop hazard:** a NULL ratio
makes `NULL <> 0` evaluate to UNKNOWN, so that row is *excluded* (3-valued logic). Since
both columns are NULLable, a future bad-data row would vanish from the ASN with no error.
The rebuild should treat a NULL/zero ratio as a data-integrity flag, not silently drop it.

### 2b. When is a ratio 100 (full qty) vs a split (VERIFIED)

Distribution of `(tire, wheel, IN_RATIO)` over all 50 rows:

```
$ SELECT IN_TIRE_RATIO, IN_WHEEL_RATIO, IN_RATIO, COUNT(*) ... GROUP BY ...;
40 | 40 | 100 | 16
100|100 | 100 | 14   <- full qty, no rounding
20 | 20 | 100 |  8
30 | 30 | 100 |  6
70 | 70 | 100 |  5
100|100 | 400 |  1
```

- `ratio = 100` → that assy takes the full vehicle qty (Delphi `count = VEHICLES * IN_ASSY_QTY`,
  no rounding; `DataModule.pas:5220`).
- `ratio < 100` (20/30/40/70) → split; Delphi computes
  `round(VEHICLES * IN_ASSY_QTY * IN_TIRE_RATIO / 100)` (`DataModule.pas:5226`).
- A patterned BC's split rows **sum to 100** within the BC (e.g. `[KLM]CC` = 20+40+40 = 100;
  `[KLM]EE` = 70+30; `[KLM]JJ` = 40+20+40) — i.e. one broadcast code splits across N assy
  part numbers by ratio. `IN_RATIO` is a *separate* multiplier (100 for all but one row,
  400 once) and is **not** used by the ASN fan-out — do not confuse it with the tire/wheel split.

### 2c. Is the wheel ratio ever the divisor, or is only the tire ratio used (VERIFIED — tire only)

**Tire ratio = wheel ratio in 100% of rows:**

```
$ SELECT COUNT(*) ... WHERE IN_TIRE_RATIO <> IN_WHEEL_RATIO;   ->  0
```

The Delphi fan-out uses **only `IN_TIRE_RATIO`** as the numerator
(`... * Inv_StoredProc.FieldByName('IN_TIRE_RATIO').AsInteger) / 100`, `DataModule.pas:5227`)
and never reads `IN_WHEEL_RATIO` for the quantity — it only *gates* on it being `=100` in the
full-qty test (`if tire=100 AND wheel=100`, `DataModule.pas:5218`). Code comment at
`DataModule.pas:5210` confirms intent: *"the share for tire and wheel will be set to the same
quantity … we'll use only one of the values."* So **wheel ratio is never a divisor**; it is a
redundant copy of tire ratio. The task's hypothesis is confirmed. The denominator is always
the literal `100`, never a ratio column.

### 2d. Rounding divergence trap (VERIFIED — relevant if math is ever moved into SQL)

The split math today is Delphi `round()`, which is **banker's rounding** (round-half-to-even):
`Round(4.5)=4`, `Round(2.5)=2`. T-SQL `ROUND()` is **half-away-from-zero**:

```
$ SELECT ROUND(15*1*30/100.0,0), ROUND(5*1*50/100.0,0), ROUND(5*1*30/100.0,0);
5.0 (4.5→5)      3.0 (2.5→3)      2.0 (1.5→2)
```

So a half-integer split (e.g. VEHICLES=15, ASSY_QTY=1, tire=30 → 4.5) yields **4 in Delphi,
5 in T-SQL** — off by one. If the rebuild reimplements the qty math, it must replicate
banker's rounding to preserve parity. Note also Jython 2.x `round()` is half-away (would
*break* parity) while Python 3 `round()` is banker's (preserves it). Pick the runtime
accordingly, or compute explicitly.

### 2e. SELECT * duplicate-column trap (VERIFIED)

`SELECT *` over the LEFT JOIN yields **26 columns**, and three names are emitted twice
(forecast + cost both have them):

```
$ sys.dm_exec_describe_first_result_set('EXEC ... KCC ...'):
 2  VC_ASSY_PART_NUMBER_CODE   ...   20  VC_ASSY_PART_NUMBER_CODE
14  VC_LAST_UPDATE             ...   25  VC_LAST_UPDATE
15  VC_ADD                     ...   26  VC_ADD
19  IN_MANIFEST_COST_ID        21  VC_ASSY_MANIFEST_NUMBER  24  MO_PRICE
```

Delphi reads by ordinal/first-match so it is unaffected. An Ignition Named Query returning a
dataset keyed by column **name** will collide on these duplicates. The rebuild query must use
an explicit projection (the Delphi caller actually only reads: `IN_MANIFEST_COST_ID`,
`VC_ASSY_PART_NUMBER_CODE`, `VC_ASSY_MANIFEST_NUMBER`, `IN_ASSY_QTY`, `IN_TIRE_RATIO`,
`IN_WHEEL_RATIO`).

---

## 3. Row ORDER — nondeterministic (VERIFIED — parity hazard)

The proc has **no `ORDER BY`**, and **both base tables are heaps** (no clustered index, no PK,
no index of any kind):

```
$ fc_indexes=0  fc_heap=1  cost_indexes=0
```

Returned order is therefore heap-scan/allocation order with **no guarantee**. Today it
*happens* to come back in `ID_FORECAST_DETAIL` order for `[KLM]CC` (157,163,164), and forcing
`MAXDOP 1` returns the same — but that is incidental to current allocation, not contractual.
A table reload, page reuse after deletes, statistics change, or parallel plan can reorder it.

This is the **No-Ratio (`Orders <= 5`) branch hazard**: that branch (`DataModule.pas:5180`)
does `Inv_StoredProc.First` then processes the **first row only** and `break`s. With
`[KLM]CC` the three rows carry tire ratios 20 / 40 / 40 and three *different* assy part
numbers (42670FEU0000 / 42670FET9000 / 42670FEU1000) — so "first row" picks a different assy
depending on nondeterministic order. **For any BC that returns multiple recipe rows, the
single-vehicle branch's output assy is nondeterministic.** Multi-row BCs (the at-risk set):

```
$ SELECT VC_BROADCAST_CODE, COUNT(*) ... GROUP BY ... HAVING COUNT(*)>1:
[KLM]JJ=3 [KLM]KK=3 [KLM][BW][BW]=3 [KLM]CC=3 [MNP]BB=3 [MNP]CC=3 [MNP]JJ=3 [MNP]KK=3
[MNP]LL=2 [MNP]EE=2 [KLM]EE=2 [KLM]LL=2 [M]QQ=2
```

**Rebuild action:** add a deterministic `ORDER BY` (e.g. `ID_FORECAST_DETAIL`) so the
single-vehicle "first row" is reproducible, and confirm with David which assy the
single-vehicle case is *supposed* to pick (the legacy result may itself have been arbitrary).

---

## 4. IN_MANIFEST_COST_ID NULL semantics — the pre-loop abort key (VERIFIED)

`IN_MANIFEST_COST_ID` is the PK of `INV_MANIFEST_COST_MST` (int, NOT NULL *in its own table*).
It becomes **NULL only via the LEFT JOIN miss** — i.e. the forecast row's
`VC_ASSY_PART_NUMBER_CODE` has **no matching manifest-cost master row**. Frequency:

```
$ fc_with_no_cost = 5    fc_total = 50      (10% of recipe rows have no cost master)
$ assy_with_multi_cost: (none)              -> LEFT JOIN never fans out / inflates count
```

The 5 cost-less assy parts:

```
[KLM]CC | 42670FEU0000   (also returns FET9000, FEU1000 which DO have cost)
... (5 rows total across the recipe)
```

**Meaning of NULL here = "this assy part is missing its manifest-cost/manifest-number
mapping," which makes the ASN line un-manifestable.** Delphi enforces this as a hard
pre-loop abort: it walks **all** rows for the BC first (`DataModule.pas:5160`), and if **any**
row has `IN_MANIFEST_COST_ID IS NULL` it raises and **fails the entire ASN create**
(`"Missing Manifest Cost Information BCode(...) ASN create failed"`, `DataModule.pas:5170`).
So one cost-less assy aborts the whole shipment, not just that line. Because `VC_ASSY_MANIFEST_NUMBER`
comes from the same joined row, a NULL cost id also means the manifest number is NULL — the
manifest string would otherwise be built from a NULL.

**Rebuild action:** replicate the *abort-the-whole-BC/ASN* behavior on any NULL
`IN_MANIFEST_COST_ID`, surfacing the offending assy part number(s). Do NOT let a NULL
manifest number silently produce a malformed manifest. The check must run over *all* matched
rows before inserting any line (the legacy accumulates `errorstr` across the full result set).

---

## 5. Inventory vs Inventory_Live — proc + data drift (VERIFIED — NO DRIFT)

- **Proc body:** `OBJECT_DEFINITION` identical in both DBs and identical to the dump.
- **Shape:** both DBs = 50 forecast rows / 29 distinct BC / 45 cost rows / single-space eff month.
- **Content:** identical `CHECKSUM_AGG(BINARY_CHECKSUM(...))` and empty set-diff both directions:

```
$ inv_checksum  = -655458872
$ live_checksum = -655458872
$ only_in_inv = 0     only_in_live = 0
```

A 2026/06 COROLLA-family BC (e.g. `KCC` → `[KLM]CC`) returns the identical recipe rows
(157/163/164, assys 42670FEU0000 / 42670FET9000 / 42670FEU1000, ratios 20/40/40) in both DBs.
**There is no forecast-recipe vintage drift between the rebuild target and the legacy
snapshot.** David can treat the recipe data as a faithful, frozen baseline for parity testing.

---

## Summary — the BC→parts/ratios mapping + the traps

**Mapping (how the proc turns a BC into ASN parts):**
One inbound broadcast code (`@BCode`) is matched by `@BCode LIKE VC_BROADCAST_CODE` against
stored `LIKE` *patterns* (character-class wildcards like `[KLM]CC`). Each matched forecast row
is one assy part number (`VC_ASSY_PART_NUMBER_CODE`) with an integer split ratio
(`IN_TIRE_RATIO` = `IN_WHEEL_RATIO`, summing to 100 across the BC's rows) and a per-assy
qty (`IN_ASSY_QTY`); a LEFT JOIN attaches the manifest number + cost id. The Delphi fan-out
then turns each row into an ASN line: full qty when ratio=100, else
`banker_round(VEHICLES * IN_ASSY_QTY * IN_TIRE_RATIO / 100)`.

**Traps the rebuild must honor:**
1. **LIKE direction** — column is the pattern; `:bcode LIKE VC_BROADCAST_CODE`. Equality
   rebinding breaks 28/29 BCs. (VERIFIED)
2. **Case-insensitive collation** on the BC match — preserve `CI_AS`. (VERIFIED)
3. **`@EffMonth` is dead data** (all eff months = `' '`, and `' '=''`) — never assume it
   filters. (VERIFIED)
4. **Zero rows = skip the BC**, not error. (VERIFIED)
5. **NULL/zero ratio silently drops the row** (`<> 0` → UNKNOWN); ratios are NULLable —
   treat as a data flag. (VERIFIED)
6. **Only tire ratio drives qty**; wheel ratio is a redundant copy; denominator is literal
   100. (VERIFIED)
7. **Rounding:** Delphi `round()` = banker's; T-SQL `ROUND()` = half-away (off-by-one on .5);
   Jython2 `round()` = half-away, Python3 = banker's. (VERIFIED)
8. **Nondeterministic order over heaps + no ORDER BY** — the single-vehicle (`Orders<=5`)
   "first row + break" picks a nondeterministic assy for multi-row BCs. Add
   `ORDER BY ID_FORECAST_DETAIL`. (VERIFIED)
9. **`IN_MANIFEST_COST_ID IS NULL` (LEFT JOIN miss, 5/50 rows) aborts the whole ASN** —
   replicate the all-rows pre-check + hard abort surfacing the assy part number. (VERIFIED)
10. **`SELECT *` over the join emits duplicate column names** (VC_ASSY_PART_NUMBER_CODE,
    VC_LAST_UPDATE, VC_ADD) — use an explicit projection in the Named Query. (VERIFIED)
