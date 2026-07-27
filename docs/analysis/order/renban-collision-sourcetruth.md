# Renban Collision — Source-Truth Note (P4)

For the renban rollover fix (clean wrap 999→000 + collision-aware allocator,
warn→guide→fix). Establishes the renban LIFECYCLE, the IN-USE / collision predicate, the
NEXT-FREE computation, and the clean-wrap safety finding on live data. Feeds the
ignition-architect's warn→guide→fix design + the build.

> **⚠ CORRECTION 2026-07-27 — read before trusting any occupancy number below.** Every "live"
> figure in this note (CMWA **863**/1000, DICAS 341, PACF 96, CAP 28, the 24-month resident window,
> 0 `VC_TERMINATED`) was measured on `Inventory.bak.stale-vintage-20260619` — a **`MAS-APPSERVER2`
> status-blind copy**, not the `APPSERVER3` production source. Production (`bak-20260724`) reads
> CMWA **897**/1000, DICAS 358, PACF 100, CAP 30, and **4112 of 4474 rows terminated**. The
> lifecycle, the IN-USE predicate (§2), and the "clean wrap alone is unsafe for CMWA" conclusion all
> still stand — the pressure is real and slightly worse than recorded. Full detail in the boxed
> correction under "Live proof" in §1, and in D14 (`docs/analysis/decisions.md`).

- **Confidence: HIGH.** Lifecycle + predicate derived from proc bodies + `RenbanOrder.pas`;
  every numeric claim proved on `Inventory_Live` (READ-ONLY) or on the writable `Inventory`
  inside a rolled-back transaction (CMWA count restored 297). Verification env `mssql-spike`,
  `SET QUOTED_IDENTIFIER ON`.
- Sources: `RenbanOrder.pas` (esp. `:775-779` renban-number, `:798` count advance, `:426-433`
  persist), `renban-data-analysis.md` (proc bodies), `order-file-data-analysis.md` §2
  (`SELECT_OrderNotOrdered`), `DELETE_AutoPurge` (read live), `DataModule.pas:6885-6929`
  (purge call site).

> The `varchar(3)` count IS a spec requirement (000-999, David confirmed) — this note does
> NOT propose widening it. It nails the predicate + next-free + whether the clean wrap alone
> is safe (it is NOT, for CMWA — the allocator is load-bearing).

---

## 1. The renban lifecycle — where a renban LIVES and what makes it OPEN vs CLOSED

A renban number is the string `<groupCode><000-999>` (e.g. `CMWA297`). It is **assigned at
the renban breakdown** (`RenbanOrder.pas` → `INSERT_OpenOrder @RenbanNum=<assigned>`), flows
to the order file, ships, then ages out. The **only table that holds a live renban is
`INV_OPEN_ORDER_INF`** (column `VC_RENBAN_NUMBER varchar(8)`). There is **no dedicated
"renban status" column and no separate renban ledger** — the renban's life is the life of the
`INV_OPEN_ORDER_INF` rows that carry it. The states, in order:

| State | Predicate on `INV_OPEN_ORDER_INF` | Meaning |
|---|---|---|
| **Placeholder (pre-renban)** | `VC_RENBAN_NUMBER = ''` | Order.pas committed a blank-renban palletized placeholder; awaits this breakdown. Selected by `SELECT_OrderNoRenban`. |
| **Assigned, file-pending (OPEN)** | `VC_RENBAN_NUMBER <> ''` AND (`VC_ORDER_DATE = ''` OR NULL) | Breakdown assigned the renban; not yet emitted to a supplier order file. This is the row `SELECT_OrderNotOrdered` will pick up. |
| **Ordered / shipped (OPEN, resident)** | `VC_RENBAN_NUMBER <> ''` AND `VC_ORDER_DATE <> ''` | Order file generated → `UPDATE_ORDEROrderDate` stamped `VC_ORDER_DATE`/`VC_SHIP_DATE`. Row stays resident; statuses (`VC_STATUS_SUPPLIER_SHIPPING`, `VC_ARRIVAL`, …) advance as the release ships/arrives. |
| **Purged (CLOSED / free)** | row DELETED by `DELETE_AutoPurge` (or no longer present) | `VC_ADD` older than `@DataRentention` months → row removed. The renban number is now reusable. |

**Key facts:**
- There is **no explicit "close" event** — a renban is freed only when its rows leave
  `INV_OPEN_ORDER_INF` via **`DELETE_AutoPurge`** (`@DataRentention` months, ≥12, on `VC_ADD`;
  `DataModule.pas:6890,6903-6904`). A renban is "in use" for as long as ANY row carrying it is
  still resident — i.e. roughly **`@DataRentention` months from issue**, NOT from ship.
- `VC_ORDER_DATE` is the file-generation gate, not a free/close marker. An order-dated renban
  is still resident (and so still occupies its number) until purged.
- `INV_OPEN_ORDER_INF_HIST` (a heap, no PK) accumulates a copy on every insert and is purged
  on the same `VC_ADD` clock — it does NOT gate reuse (the breakdown/file feed never read it).

**Live proof (`Inventory_Live`):** 4284 rows, **all** `VC_RENBAN_NUMBER <> ''` and **all**
`VC_ORDER_DATE` stamped (snapshot is fully post-order-file); `VC_ADD` spans
`20240702…20260619` ≈ **24 months resident**, 0 `VC_TERMINATED`. So the resident window is
~24 months and the purge has not trimmed below it.

> **CORRECTION (2026-07-27) — the Live-proof paragraph above was measured on a STATUS-BLIND copy,
> not on production.** Those exact figures (4284 rows / 0 `VC_TERMINATED`) reproduce only against
> `Inventory.bak.stale-vintage-20260619`, whose backup header names **`MAS-APPSERVER2`** — not the
> `APPSERVER3` production source. This is the `/backup`-mount provenance failure ruled on in
> `dverespey/inventory` #198 (that mount is never a restore source).
>
> Restored side by side and joined row-by-row on (supplier, part, FRS, renban, `VC_ADD`):
> - All 4284 of its rows match a row in the `APPSERVER3` `bak-20260724` vintage — **same site, same
>   lineage, a strict subset** (4284 vs 4474 rows, five weeks apart).
> - **Zero** of its 4284 rows carry any `VC_STATUS_SUPPLIER_SHIPPING` or `VC_ARRIVAL` value, while
>   4112 of the same rows carry both in production. Two years of shipments with no arrival ever
>   recorded is not a plausible production state.
> - `VC_LAST_UPDATE` is **byte-identical** on those 4112 rows across both vintages, so whatever wrote
>   the stamps did not bump it — the signature of the trigger-bypassing external feed of #47. Its
>   maximum across those rows is `2026-05-26 11:04:24`, #47's feed-death date to the day.
>
> **What still stands:** the ring-pressure finding and the conclusion that the clean wrap alone is
> unsafe for CMWA (863/1000 occupied there vs **897**/1000 in production — the pressure is real and
> slightly worse than recorded). The IN-USE predicate (§2) and the allocator design are unaffected.
>
> **What does NOT stand:** the implied reading that the data carries no close/terminate signal. In
> production `VC_TERMINATED` is populated on **4112 of 4474** resident rows (455 distinct stamp dates,
> `20240714`→`20260607`), and the imminent CMWA collisions are pinned exclusively by terminated rows.
> That signal was nevertheless **rejected as a predicate input** — its writer has been dead since
> 2026-05-26 — and the fossils are instead freed by running the existing retention purge at cutover.
> See D14 in `docs/analysis/decisions.md` and `dverespey/inventory` #256.
>
> **Identity ANSWERED (David, 2026-07-27):** *"APPSERVER2 is being decommissioned, all data and
> applications are being moved to APPSERVER3 permanently."* So it is the **predecessor server, mid-
> migration** — not another site and not a standby. `APPSERVER3` is production. This also explains how
> its backup reached the docker `/backup` mount (#198): the mount pointed at the outgoing server's
> backup location.
>
> **CUTOVER HAZARD — direction of the migration copy.** APPSERVER2's `Inventory` looks superficially
> complete and current (same 4284 rows, same `VC_ADD`/`VC_LAST_UPDATE`, backed up 2026-06-19), so a
> decommissioning migration could reasonably treat it as authoritative. **It is not.** Copying it onto
> APPSERVER3 would silently destroy the milestone layer that exists ONLY on APPSERVER3 —
> `VC_STATUS_SUPPLIER_SHIPPING`, `VC_ARRIVAL` and `VC_TERMINATED` on 4112 rows. The row counts would
> still look right afterwards, so the loss would not announce itself. Tracked on `dverespey/inventory`.
>
> **Still unexplained (and now moot for the rebuild):** WHY the APPSERVER2 copy carried the order
> inserts but none of the milestone updates, while staying current to its backup date. The likeliest
> reading is that the trigger-bypassing external feed of #47 wrote milestones to APPSERVER3 only. Not
> worth chasing — the server is going away — but do not treat that copy as a source for ANY behavioural
> claim about production.

> **CAN'T fully pin from source:** the exact configured `@DataRentention` value lives in the
> operator's INI (`[DATAPURGE]`, git-ignored) and the purge cadence is manual/scheduled
> (`AutoPurge` is operator/timer-driven, not transactional). The 24-month resident window is
> what the data shows; treat the retention as a tunable that DIRECTLY governs collision safety
> (§4). Confirm the prod `[DATAPURGE]` retention with David.

---

## 2. The IN-USE / collision predicate (the WARN check)

Given a group `G` (e.g. `'CMWA'`) and a candidate suffix `N` (000-999), the renban string is
`@renban = G + RIGHT('000' + CAST(N AS varchar(10)), 3)`. It is **IN USE** (issuing it would
collide with a still-resident release) iff a resident row already carries it:

```sql
-- IN-USE predicate (resident-rows definition — the safe, strictest one for the allocator)
EXISTS (
  SELECT 1 FROM INV_OPEN_ORDER_INF o
  WHERE o.VC_RENBAN_NUMBER = @renban     -- @renban = G + zero-pad-3(N)
)
```

Two definitions, pick **resident** for the allocator:
- **Resident (recommended):** ANY row with that renban → in use. A purged renban frees up; an
  ordered/shipped-but-not-purged renban is still occupied. This is the only definition that
  prevents a wrap from reusing a number whose old rows are still in the table.
- **File-pending (narrower):** `... AND (o.VC_ORDER_DATE = '' OR o.VC_ORDER_DATE IS NULL)` —
  "in use by a not-yet-ordered release." Too narrow for collision avoidance: an ordered renban
  that hasn't been purged would (wrongly) read as free and collide.

**Equality, not LIKE.** Match `VC_RENBAN_NUMBER = @renban` (exact). A `LIKE G+'%'` over-matches
(`'CMWA'` catches `'CMWA1000'`); only use `LIKE` with a `[0-9][0-9][0-9]` length guard for
group-level scans, never for the per-candidate check.

**Live proof (`Inventory_Live`, CMWA):** `CMWA100`→2 resident rows (IN USE), `CMWA500`→2 (IN
USE), `CMWA999`→3 (IN USE), but `CMWA297`/`CMWA298`/`CMWA300`→0 (FREE). The current count (297)
is genuinely free.

---

## 3. The NEXT-FREE computation (the GUIDE suggestion)

**Scan order: next-after-current, wrapping** (matches the legacy's monotonic count advance —
the breakdown issues `seed3 + TruckNumber` and persists `max+1`, so the natural suggestion is
"the next number at/after the current count that is free"). Lowest-free would diverge from the
legacy's forward-walking number and could hand back a recently-purged low number out of order.

```sql
-- NEXT-FREE: first suffix at/after the current count, wrapping 999->000, with NO resident row.
DECLARE @grp varchar(5) = 'CMWA';
DECLARE @start int = TRY_CAST((SELECT VC_RENBAN_GROUP_COUNT FROM INV_RENBAN_GROUP_MST
                               WHERE VC_RENBAN_GROUP_CODE = @grp) AS int);
;WITH nums AS (
  SELECT TOP (1000) (ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) - 1) AS k
  FROM sys.all_objects
), cand AS (
  SELECT ((@start + k) % 1000) AS n, k AS dist FROM nums   -- dist 0..999 = steps from count
)
SELECT TOP 1
       c.n AS next_free_suffix,
       @grp + RIGHT('000' + CAST(c.n AS varchar(10)), 3) AS next_free_renban,
       c.dist AS steps_from_current
FROM cand c
WHERE NOT EXISTS (SELECT 1 FROM INV_OPEN_ORDER_INF o
                  WHERE o.VC_RENBAN_NUMBER = @grp + RIGHT('000' + CAST(c.n AS varchar(10)), 3))
ORDER BY c.dist;
```

- The candidate set is exactly the 1000 suffixes `(start+0 .. start+999) % 1000` — a full lap,
  so it always considers every number once, in forward-wrap order.
- **Exhaustion (all 1000 in use):** the `WHERE NOT EXISTS` returns **zero rows** → the query
  yields no row. The allocator MUST treat "no row returned" as **space-exhausted → hard
  warn/block** (do not silently wrap into a live number). This is the FIX-step backstop. In
  practice no group is close: occupancy is CMWA 863/1000, DICAS 341, PACF 96, CAP 28, HCAP 0.

**Live proof (CMWA, count 297):** next-free = `CMWA297` at 0 steps (the count is already free);
then 298, 299, 300… contiguous. The simple count-advance and the next-free scan AGREE here —
which is exactly the steady-state the warn step should confirm silently.

> For the warn→guide step the architect can return the *next K free* (TOP K instead of TOP 1)
> to offer the operator a short pick-list, and surface `steps_from_current` so a large jump
> (operator about to skip a big occupied band) is visible.

---

## 4. Clean-wrap safety on live data — the allocator IS load-bearing (CMWA)

**Question:** when the count wraps 999→000, is the 000-block (issued ~1000 renbans / one lap
ago) reliably CLOSED (purged) by then? **Answer: NO for the high-volume CMWA group — the lap
time ≈ the retention window, so the wrap returns to numbers whose rows have NOT yet purged.**

Quantified on `Inventory_Live`:

| Metric (CMWA) | Value | Source |
|---|---|---|
| Issue rate | ~35-44 distinct renbans / month | order-month histogram |
| Lap time (1000 numbers) | ~24-28 months ≈ **~2.4 yr** | 1000 ÷ ~37/mo |
| Resident window (purge) | **~24 months** (`VC_ADD` 2024-07 → 2026-06, 0 terminated) | `DELETE_AutoPurge` on `VC_ADD`, `@DataRentention` ≥12 |
| Suffixes occupied NOW | **863 / 1000** (137 free) | distinct-renban count |
| **Free run AHEAD of count (297)** | **only 35 suffixes** (297-331) | forward-wrap occupancy scan |
| Age of the next occupied block ahead | `CMWA332` dated **2024-07-02** (~2 yr old, the OLDEST resident row) | scan + order-date |

**Interpretation:** lap ≈ retention, so the occupied band wraps almost all the way around. The
count (297) has only **~35 free slots (~1 month of runway)** before it reaches `CMWA332`, a
block that is ~2 years old but **still resident** (not yet purged). So the count is chasing its
own tail with a razor-thin buffer that exists ONLY because purge ≈ lap. **A naive clean wrap
(999→000) would, within ~1 month, advance into still-resident old rows → renban-number reuse =
collision.** The collision-aware allocator (skip-occupied) is therefore **essential, not
optional**, for CMWA.

**The legacy rollover bug has ALREADY FIRED in live data** (the wrap is real, not hypothetical):
- 2026-01-15: `CMWA997`, `CMWA998` issued.
- 2026-01-16: `CMWA999`, `CMWA1000` issued → `next_count` crossed 1000 → persisted count
  LEFT-TRUNCATED `'1002'→'100'` (varchar(3)).
- 2026-01-20: counter re-seeded from 100 → `CMWA100`, `CMWA101` RE-ISSUED, then climbed.
- Today the count is back at 297. So one full operational lap took ~Jul-2024-era 332 → 999 →
  wrap → 100 → 297 ≈ the visible 24-month span.

**Why no DUPLICATE renban survives in the snapshot today:** the only reason `CMWA100` (re-issued
Jan-2026) does NOT collide with an *original* `CMWA100` is that the original aged out / purged
before the wrap returned. That safety is **accidental and retention-tuned**, not designed — and
the §4 free-run-ahead = 35 shows it is now near the edge. (The only renbans serving 2+ FRS today
— `DICAS225`, `DICAS244` — are single-order-date multi-trailer-day runs, NOT wrap reuse; proved
they share one `VC_ORDER_DATE`.)

**Other groups (headroom to 999 + occupancy) — none near rollover NOW, but CMWA is the canary:**

| Group | Count | Headroom to 999 | Occupied/1000 | Risk |
|---|---:|---:|---:|---|
| **PACF** | 634 | 365 | 96 | 2nd-fastest count climb; lots of free space |
| DICAS | 484 | 515 | 341 | moderate |
| **CMWA** | 297 | 702 | **863** | **HIGH — already wrapped once; ~1 mo free runway ahead** |
| HCAP | 088 | 911 | 0 | low volume |
| CAP | 068 | 931 | 28 | low volume |

> **Headroom-to-999 is misleading for CMWA** — its count (297) is low, but it already wrapped,
> so it is mid-lap with 863/1000 occupied. The number that fires the collision is the
> **free-run-ahead (35)**, not the distance to 999. Watch occupancy + free-run-ahead, not the
> raw count.

---

## 5. Edge cases

1. **Seed read `int(str(renban_seed)[-3:])` under a wrapped count** (`code.py:139`,
   `RenbanOrder.pas:775`). `renban_seed = groupCode || count` (e.g. `'CMWA297'`,
   `'DICAS484'`). `RIGHT(seed,3)` = the count's last 3 chars. **Safe ONLY because the count is
   always exactly 3 chars** (varchar(3), zero-padded by `Format('%.3d')`). Proved live: all 5
   counts are LEN=3 and ISNUMERIC=1 (`068/297/484/088/634`). HAZARD: if the count were ever
   <3 chars (e.g. `'5'`), `RIGHT('CMWA5',3)='WA5'` → `int('WA5')` crashes. The legacy
   never writes <3 chars; the rebuild must zero-pad to 3 on write (it does: `("%03d"%N)[:3]`).

2. **A count already past 999 in live data (has the legacy bug already fired?).** YES — see §4:
   `CMWA1000` exists (2026-01-16) and the count collapsed `'1000'→'100'` then re-climbed. No
   group currently holds a >3-char count (all are clean 3-char now), but the WRAP and re-seed
   have demonstrably happened. The post-cutover fix (clean wrap + allocator) prevents the
   re-occurrence; the rebuild as-shipped faithfully REPRODUCES the truncation for parallel-run
   parity (do not "fix" silently in phase-1 — adversary BLOCKER-1, RESOLVED-as-faithful).

3. **NULL / blank renban handling.** `VC_RENBAN_NUMBER` is `NOT NULL` (declared) — the
   "no renban" sentinel is the empty string `''`, not NULL. The in-use predicate uses `= @renban`
   (a non-empty string) so blanks never match a candidate. `VC_ORDER_DATE` is also `NOT NULL`
   (the `IS NULL` arms in the procs are defensive); in practice `=''` is the live branch. The
   next-free scan only emits non-blank `G + 3-digit` candidates, so it can never suggest `''`.

4. **`varchar(3)` count truncation — confirmed live (ROLLED BACK).** On the writable `Inventory`
   inside `BEGIN TRAN … ROLLBACK` (CMWA restored 297):
   `EXEC UPDATE_RenbanGroupCount @RenbanCount='1002'` → stored `'100'`;
   `@RenbanCount='1000'` → `'100'`; `@RenbanCount='999'` → `'999'`. The proc does NO truncation
   logic — the cut is the `@RenbanCount varchar(3)` parameter binding keeping the LEFTMOST 3
   chars. The legacy sends `Format('%.3d',[N])` (min-width-3, never caps), so the exact reduction
   is `("%03d" % N)[:3]` — which the rebuild matches (`code.py:345`). The renban NUMBER itself
   (`VC_RENBAN_NUMBER varchar(8)`) is UNAFFECTED: `CMWA1000` is 8 chars and fits (a 5-char group
   like `DICAS1000` is 9 chars and WOULD truncate at varchar(8) — identically on both sides).

5. **Group-prefix parse ambiguity for scans.** `LEFT(renban, LEN-3)` mis-buckets rollover
   strings (`CMWA1000` parses to prefix `CMWA1`). For per-group occupancy use
   `LIKE G+'[0-9][0-9][0-9]' AND LEN(renban)=LEN(G)+3` (exact 3-digit) to exclude the rollover
   4-digit artifacts, OR `= G + zeropad3(N)` for a specific candidate. The data has a residual
   `CMWA1000` block from the past wrap (2 rows) that must not be counted as group `CMWA1`.

6. **Exhaustion (all 1000 in use).** Not reachable today (max occupancy CMWA 863/1000) but the
   allocator MUST handle "next-free returns no row" as a hard block/warn, never a silent wrap
   into a live number. This is the only correctness backstop if retention is ever raised so far
   that a full lap stays resident (lap < retention → ring fully occupied).

---

## RETURN — answers to the four asks

1. **IN-USE / collision predicate (exact SQL):**
   `EXISTS (SELECT 1 FROM INV_OPEN_ORDER_INF o WHERE o.VC_RENBAN_NUMBER = @grp + RIGHT('000'+CAST(@n AS varchar(10)),3))`
   — **resident-rows** definition (any resident row, regardless of order-date/status), exact
   equality. The renban lives ONLY on `INV_OPEN_ORDER_INF.VC_RENBAN_NUMBER`; it frees only when
   `DELETE_AutoPurge` removes the rows (≥12-month `@DataRentention` on `VC_ADD`).

2. **NEXT-FREE approach:** forward-wrap scan — generate the 1000 suffixes
   `(count+0 .. count+999) % 1000` in distance order, return the first with NO resident row
   (`NOT EXISTS` above), TOP 1 (or TOP K for a pick-list). Exhaustion (no row) = hard block.
   Forward-from-count (not lowest-free) matches the legacy's monotonic count advance.

3. **Is the clean wrap alone safe, or is the allocator essential?** **The allocator is
   ESSENTIAL for CMWA.** Lap time (~2.4 yr) ≈ retention window (~24 mo), so 863/1000 suffixes
   are occupied and the count has only **~35 free slots (~1 month)** before it advances into a
   still-resident ~2-year-old block (`CMWA332`, 2024-07-02). A naive clean wrap would reuse a
   live number within ~1 month → collision. The clean wrap fixes the >999 truncation bug but
   does NOT by itself guarantee a free number; the skip-occupied allocator is load-bearing.

4. **Live group near rollover NOW:** **CMWA** — already wrapped once (the legacy bug fired Jan
   2026: `CMWA999/1000` → count `'100'` → re-issued `CMWA100`), 863/1000 occupied, ~1-month
   free runway ahead. It is the canary; the allocator must be live for it before the next lap
   tightens. (PACF count 634 is the next-fastest climber but has 904 free suffixes — not near.)
