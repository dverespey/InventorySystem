# Adversary Findings — Renban Breakdown rebuild vs legacy `RenbanOrder.pas`

Adversarial parity review. Goal: REFUTE that the rebuilt renban breakdown
(`docs/analysis/order/project-library/renban/code.py`) reproduces the legacy
`RenbanOrder.pas` trailer-distribution + FRS/renban assignment + delete-then-reinsert
write-back. Default stance: the reimplementation is wrong until proven equivalent on the
SAME inputs.

- Legacy bodies read in full: `RenbanOrder.pas` (896 lines), the 4 procs read from the
  LIVE `Inventory` DB via `OBJECT_DEFINITION` (not the spec): `INSERT_OpenOrder`,
  `DELETE_OrderRenban`, `UPDATE_RenbanGroupCount`, `SELECT_OrderNoRenban`.
- Rebuild: `docs/analysis/order/project-library/renban/code.py`.
- Spike: `mssql-spike`, READ-ONLY on `Inventory_Live`, rolled-back trans on `Inventory`;
  spike verified restored as-found (CMWA count 288, on-hand unchanged, 0 stray rows/hist).

The distribution fuzz (`scripts/e2e/test_renban_build.py`) and the e2e write-back
(`scripts/e2e/test_renban_e2e.py`) BOTH pass as-shipped (28/0 and 16/0). The defects below
live in the **rollover tail** and a **partial-lot edge** that those tests do not exercise.

---

## Method note — why the in-repo fuzz is not sufficient on its own

The in-repo "non-vacuous" fuzz (`test_renban_build.py:36-56` `ref_distribute`) is a
**line-for-line copy** of the production `_distribute_part`, INCLUDING the `sizes[0]` /
`trucks[0].size` quirk (`:46` vs `code.py:98`). It can never disagree with production on
anything they share — it proves self-consistency, not legacy-equivalence.

To get a real check I wrote an **independent Pascal-object simulator** (modeling
`TTruck.fCurrentCount`, the `fOrderList` per-part merge, and `TGroupRenban.AddOrder` Phase
A/B from first principles off `RenbanOrder.pas:155-301`, NOT off `code.py`), and compared
production `compute_trailer_breakdown` against it:

- **truck-load totals:** 11,172 scenarios (T 1..6, P 1..8, 1-2 parts incl. over-capacity-
  per-part, non-divisible, exact-fill, forward-spill, lots=0) → **0 mismatches**.
- **full per-(truck,part) read-out** (lots + FRS + renban + qty + next_count) on the named
  hard cases (over-cap 7 lots, A=5/B=3 spill, multi A3/B3/C2, exact fill A=8, zero-lot row,
  carry across 3 trucks, 11 trailers → FRS≥10, renban rollover seed 997) → **0
  divergences**.

So the **core distribution math + the FRS suffix + the in-run renban STRING are PROVEN
faithful** (I could not break them). The two findings below are NOT in the distribution
math — they are in the COUNTER-PERSISTENCE rollover and the WRITE-BACK delete scope.

---

## BLOCKER 1 — group-counter rollover diverges (`Format('%.3d')` min-width + `varchar(3)` left-trunc  vs  `% 1000`)

**Claim refuted:** "the count bump = maxRenban+1 … reproduce the 3-digit zero-pad + wrap"
(spec §12.6 / data-analysis §3.3 / `code.py:320-325`).

**Defect type:** code defect (write-back), reachable on real climbing counters.

**Counterexample (input → legacy vs rebuild):** any breakdown whose advanced counter
`next_count = seed3 + T >= 1000`. With group `CMWA`, seed `997`, 5 trailers, a 5-lot part:

```
renban numbers (in-run, IDENTICAL both sides): CMWA997, CMWA998, CMWA999, CMWA1000, CMWA1001
next_count = 1002
LEGACY persists  : Format('%.3d',[1002]) = '1002'  --> proc @RenbanCount varchar(3) keeps
                   the LEFTMOST 3 chars --> '100'
REBUILD persists : '%03d' % (1002 % 1000) = '002'
```

**Proven on the LIVE proc** (`mssql-spike`, `Inventory`, rolled back, restored to 288):

```
EXEC UPDATE_RenbanGroupCount @RenbanCode='CMWA', @RenbanCount='1002';  -- legacy sends this
  --> stored VC_RENBAN_GROUP_COUNT = '100'
EXEC UPDATE_RenbanGroupCount @RenbanCode='CMWA', @RenbanCount='002';   -- rebuild sends this
  --> stored VC_RENBAN_GROUP_COUNT = '002'
```

**Root cause:** Delphi `Format('%.3d',[N])` treats `.3` as a *minimum* width (zero-pad to
≥3 digits, **never** caps): `1002 -> '1002'`. The proc param `@RenbanCount varchar(3)`
then truncates by keeping the **leftmost** 3 characters (`'1002' -> '100'`). The rebuild
(`code.py:325`) instead does `next_count % 1000` (`1002 -> 2 -> '002'`). These are
different reduction functions for any N≥1000: legacy = `str(N)[:3]`, rebuild = `N % 1000`.

**Why it matters / reachability:** the persisted counter is the SEED for the NEXT
breakdown of that group (`seed3 = int(group_count[-3:])`, `code.py:139`). A wrong seed →
the next run's renban numbers diverge from legacy → renban collisions or mis-broadcast
codes downstream. It is reachable: the counter is `varchar(3)` so it climbs toward 999
over operational time and wraps; live counts today are CAP 068, HCAP 088, CMWA 288/297,
DICAS 480/484, **PACF 633/634** — actively climbing. Any group reaching ~994+ with up to 6
trailers crosses 1000 in `next_count`. (The renban STRING itself is NOT affected — both
`Format('%.3d')` and `'%03d'` render `str(n)` for n≥1000; only the persisted counter
diverges. This narrows blast radius to "the next run after a rollover," but that next run
then mis-numbers every trailer.)

> NB the in-run renban for rcount≥1000 becomes 8 chars for a 4-char group (`CMWA1000`) and
> 9 chars for a 5-char group (`DICAS1000`, truncated by `VC_RENBAN_NUMBER varchar(8)`) —
> but this truncation is IDENTICAL on both sides, so not a divergence; only the counter is.

---

## SHOULD-FIX 2 — write-back deletes a partial-lot (`lots = 0`) placeholder the legacy preserves

**Claim refuted:** "per part `DELETE_OrderRenban(part, FRS='', renban='')` deletes ALL
blank rows for that part; then `INSERT_OpenOrder` per trailer-row" reproduces the legacy
write-back exactly (task item 4; spec §6.1 / §12.8).

**Defect type:** code defect (write-back scope), edge-reachable (sub-lot blank-renban
order).

**Counterexample (input → legacy vs rebuild):** a blank-renban grouped-part order with
`0 < order_qty < lotqty` (so `lots = order_qty div lotqty = 0`), e.g. part `B`, lotqty 40,
order_qty 20, alongside a normal part `A` (qty 200, lotqty 40 → 5 lots), 2 trailers × 10.

```
compute_trailer_breakdown emits rows for: ['A']    (B lots=0 -> lands on no truck -> no row)
total_lots = 5                                       (B contributes 0)
```

- **LEGACY:** `LoadScreen` puts B in the grid with Lots=0; `FRSBreakdown` feeds
  `AddOrder(lots=0)` → Phase A (`0 div T`=0) skipped, Phase B (`0 mod T`=0) skipped → B on
  no truck → read-out (`:746-794`) emits NO grid row for B → `fAvailableCount =
  RowCount-1` excludes B → the commit loop `for i:=1 to fAvailableCount` (`:417`) never
  calls `NewFRSOrder` for B → **`DELETE_OrderRenban` is NEVER called for B**. B's blank
  placeholder **SURVIVES** (correct — it waits for a future breakdown once its qty grows).

- **REBUILD:** `commit_renban_breakdown` builds `parts_seen` from the full loaded `orders`
  (`code.py:299-303`, `for o in orders`), NOT from the emitted `rows`. So it issues
  `DELETE_OrderRenban` for B → B's blank placeholder is **DELETED with no re-insert** (B
  has no qty>0 row to insert) → **the order is silently lost** (and, because the blank
  renban was its only "needs grouping" marker, it never ships and never errors — H1).

**Fix direction (for the architects, not me):** derive `parts_seen` from the EMITTED rows
(parts that actually produced a trailer-row), not from the loaded feed — matching the
legacy's "delete only what the grid emitted."

**Reachability caveat:** requires a blank-renban placeholder with `0 < qty < lotqty`.
Data-analysis proves *already-grouped* CMWA rows are exact lot multiples (1957/1957), and
the normal Order-commit path writes `IN_QTY = R × IN_1LOTQTY` (a multiple). A sub-lot blank
placeholder is therefore off the happy path — but nothing in this stage GUARANTEES it (no
guard rejects qty<lotqty), so it is a latent data-loss divergence, not impossible. Hence
SHOULD-FIX rather than BLOCKER.

---

## Confirmed FAITHFUL (could not refute) — with proof

- **Distribution math (task item 1):** PROVEN equivalent vs an independent Pascal-object
  trace across 11,172 truck-load scenarios + full read-out on every named hard case
  (over-capacity-per-part, non-divisible, exact-fill, forward-spill with leftover carry,
  lots=0, multi-part differing spill). 0 mismatches. The `truck[0].size` quirk
  (`code.py:98`) is faithfully preserved (correct for equal-capacity, the only live case).
  **Could not break it.**

- **FRS / renban assignment (task item 2):** faithful. `_frs_suffix` matches
  `RenbanOrder.pas:763-767` for truck index 0..14 (boundary `TruckNumber > 8`: tn=8→`'09'`,
  tn=9→`'10'`, tn=10→`'11'`). `_renban_number` = group + min-3-width count, matching
  `:775-779`. The renban-string rollover at 999 renders identically to Delphi
  `Format('%.3d')` (both → `str(n)` for n≥1000). seed3 = `rightstr(seed,3)` faithful.

- **FRS-suffix no-op (task item 3):** RE-PROVEN on the LIVE `INSERT_OpenOrder`. With a full
  7-char `@FRSNum` and `@RenbanNum<>''`, the recompute `SET @FRSNum = @FRSNum +
  RIGHT(...+1,2)` produces 9 chars that truncate back to `varchar(7)` = the input. Proven
  on spike (rolled back): inserting two rows with the SAME 7-char FRS `'6090101'` both
  store `'6090101'` (the max+1 append `'609010102'` truncates to `'6090101'`). So the
  proc HONORS Pascal's `TruckNumber+1` suffix; the rebuild correctly does NOT reimplement
  max+1. **Holds regardless of the DELETE-then-INSERT ordering difference** (interleaved in
  Pascal vs batched in the rebuild) because the truncation is order-independent.

- **Write-back delete-then-reinsert (task item 4, happy path):** e2e PASS 16/0 on the live
  spike (restored as-found). NO blank rows left (remaining_blank=0), NO update-in-place
  (the abandoned `:482-539` path is not resurrected), count bump = seed3+T (288→291 for 3
  trailers), stock-neutral (on-hand 13341/28133/44418 unchanged — status-empty rows gate
  the triggers off), idempotent re-run is a clean no-op. Faithful for `next_count < 1000`.

- **Aliased feed (task item 5):** faithful. `_SELECT_NO_RENBAN` (`code.py:239-253`) aliases
  `o.IN_QTY AS order_qty` + `p.IN_1LOTQTY AS lotqty`, dodging the `SELECT *` duplicate
  `IN_QTY` trap (proved live: ord 6 `o.IN_QTY` order-qty 400/360/440 vs ord 42 `p.IN_QTY`
  on-hand 44418). Same blank-renban filter (`o.VC_RENBAN_NUMBER = ''`) and the same 3
  inner joins as the proc. Reads ORDER qty, not on-hand. (The legacy `fieldbyname('IN_QTY')`
  also takes ord 6 — they agree.)

- **R3 sum-all FAITHFUL (task item 6):** faithful. `_Truck.add` merges repeated parts by
  SUMMING lots (`code.py:63-64`), matching `TTruck.AddOrder:160-163`; no renban filter
  added; delete-then-reinsert leaves no blank rows so the downstream sum-all stays correct.
  Unit test "merge: lots SUMMED (5+3=8)" passes.

---

## Distribution-fuzz result

**Could NOT break the distribution.** Independent Pascal-object trace agrees with
production on 11,172 truck-load scenarios + every named hard read-out case (0 mismatches /
0 divergences). The in-repo 10,200-scenario fuzz also passes — but note it is
self-referential (its reference is a copy of the production code) and would not have caught
a shared bug; the independent trace closes that gap. The distribution math, FRS suffix, and
in-run renban string are PROVEN faithful.

## FRS / renban result

Faithful. FRS suffix matches `:763-767` across the `TruckNumber > 8` boundary; renban
matches `:775-779` including seed-tail read and the 999 rollover of the renban STRING
(identical to Delphi `Format('%.3d')`). The FRS-suffix server-side recompute is a proven
varchar(7) truncation no-op on the live proc.

## Write-back result

Mostly faithful (e2e 16/0 on the happy path), with TWO divergences:
1. **BLOCKER 1** — the group COUNTER rollover for `next_count >= 1000` (legacy `str(N)[:3]`
   via `Format('%.3d')`+`varchar(3)` left-trunc vs rebuild `N % 1000`); proven on the live
   proc.
2. **SHOULD-FIX 2** — a partial-lot (`lots=0`) blank placeholder is DELETED by the rebuild
   (deletes from the loaded feed) but PRESERVED by the legacy (deletes only emitted grid
   rows) → silent order loss.

---

## VERDICT

The rebuild is **NOT proven equivalent** to the legacy `RenbanOrder.pas` — there are real
divergences:

- The **trailer distribution, FRS suffix, in-run renban numbers, FRS-suffix no-op, aliased
  feed, R3 sum-all, and the happy-path delete-then-reinsert write-back are PROVEN
  FAITHFUL** (independent trace + live-proc proofs; I could not break them).
- But the **group-counter rollover diverges** (BLOCKER 1: `str(N)[:3]` vs `N % 1000` for
  `next_count >= 1000`, proven on the live proc; reachable as counters climb to 999), and
  the **write-back deletes a partial-lot placeholder the legacy preserves** (SHOULD-FIX 2:
  data-loss on a sub-lot blank order).

So: faithful modulo the golden-pending gap is **FALSE as stated** — there are two concrete
divergences (one BLOCKER in the rollover tail, one SHOULD-FIX at the partial-lot edge),
both outside the range the shipped tests exercise. Bounce the fixes to the
ignition-developer / sql-analyst; the rollover reduction function and the `parts_seen`
derivation are the two lines to change.

---

# RE-VERIFY (round 2) — 2026-06-21

Re-attack of the two round-1 findings after the developer's claimed fixes, plus the
all-lots-0 group edge raised in the dev's scope note. Method unchanged: legacy bodies read
from the LIVE `Inventory` DB, every claim carried by a live-proc proof or a counterexample.
Spike (`mssql-spike`) driven READ-ONLY on references; all probes in rolled-back trans;
restored as-found (E2E final check: CMWA count 288, on-hand 13341/28133/44418 unchanged,
0 leftover rows/hist). A scratch-mutated copy used to prove the regression guard was
restored immediately (`code.py:345` back to the fixed form; verified on disk).

## BLOCKER 1 (rollover) — RESOLVED, with proof

The rebuild now persists `("%03d" % next_count)[:3]` (`code.py:345`). I re-proved the legacy
reduction on the LIVE `UPDATE_RenbanGroupCount` (`@RenbanCount varchar(3)`; the proc does NO
truncation logic — the cut happens at the varchar(3) **parameter binding**, leftmost-3),
inside a rolled-back tran:

```
EXEC UPDATE_RenbanGroupCount @RenbanCode='CMWA', @RenbanCount=...   -> stored:
  '1002'  -> '100'      '1000' -> '100'     '999' -> '999'
  '005'   -> '005'      '634'  -> '634'      '10000'-> '100'   '001' -> '001'
ROLLBACK -> CMWA back to 288.
```

The legacy value sent is `Format('%.3d',[N])` = min-width-3 zero-pad = Python `"%03d" % N`
(NOT bare `str(N)`; for N<100 Delphi pads to `'005'`/`'001'`, which I confirmed round-trip
verbatim on the proc, LEN=3). So the exact legacy reduction is `("%03d" % N)[:3]`, which is
**character-for-character what the rebuild now sends**. Verified across every task case:

| next_count | legacy `("%03d"%N)[:3]` | rebuild `code.py:345` | OLD `%1000` bug |
|-----------:|:-----------------------:|:---------------------:|:---------------:|
| 1002 | `100` | `100` | `002` |
| 1000 | `100` | `100` | `000` |
| 999  | `999` | `999` | `999` |
| 5    | `005` | `005` | `005` |
| 634  | `634` | `634` | `634` |
| 10000| `100` | `100` | `000` (5-digit) |

The renban NUMBER (`VC_RENBAN_NUMBER varchar(8)`) is **unaffected** — the in-run strings
climb cleanly `CMWA997..CMWA1001` (8 chars fit); re-proven live in the E2E
(`got=['CMWA1000','CMWA1001','CMWA997','CMWA998','CMWA999']`). Only the `varchar(3)` COUNT
truncates, identically on both sides.

**Regression guard is real (non-vacuous), proven two ways:**
- Pure test asserts both the new value AND that the OLD `%1000` would give `'002'`
  (`test_renban_build.py:202-203`).
- I drove the **production** `commit_renban_breakdown` through the live-DB E2E with a
  SCRATCH copy reverted to `("%03d" % (next_count % 1000))` → the rollover assertion FAILED
  with `persisted='002'` (E2E 26/1), spike still restored. The fixed code persists `'100'`
  (E2E 27/0, `test_renban_e2e.py:271`). So a revert to `% 1000` is caught by a real
  DB-persisted-value check, not a self-recompute.

**Carry note (unchanged, correct):** the rollover is a LATENT LEGACY BUG — at
`next_count>=1000` BOTH sides collapse the counter to `'100'`, so the next run re-seeds from
`'100'` and renban numbers collide with the earlier `CMWA100x` block. The rebuild faithfully
reproduces this for parallel-run parity (documented `code.py:334-339`, IG83-TODO). This is
the "documented rollover-latent-bug carry," not a new divergence.

## SHOULD-FIX 2 (partial-lot) — RESOLVED, with proof

The delete set is now derived from the EMITTED rows (`parts_seen` built from `rows`,
`code.py:307-311`), matching the legacy commit loop which walks only the grid rows
(`RenbanOrder.pas:417` `for i:=1 to fAvailableCount`, `fAvailableCount = RowCount-1` `:799`,
`NewFRSOrder` deletes by `AvailableGrid.Cells[2,Row]` `:506`).

Re-proven live (E2E, rolled-back fixture): a feed of a sub-lot part (`Q4000` qty 20 <
lotqty 40 → 0 lots) + a normal part (`Q5000` 300/30 → 10 lots), 3 trailers:
- the sub-lot part emits NO trailer row; `total_lots=10`;
- the sub-lot part's BLANK placeholder **SURVIVES** (`zero_blank=1`), qty intact (20),
  exactly 1 untouched row (`zero_total=1`);
- the normal part is grouped (`norm_blank=0`, `norm_grouped=3`);
- the counter advances from the emitted (normal-part) trailers only (291→294).

I then attacked the ONE residual path — a sub-lot row **sharing a part** with a full row.
`DELETE_OrderRenban` (re-read live) with `@FRSNumber='' AND @RenbanNumber=''` deletes
part-wide on blank renban. In that case the part IS in `parts_seen` (it emits via the full
row), so the rebuild's part-wide DELETE sweeps BOTH blank rows. **The legacy does the same**:
`TTruck.AddOrder` merges by part (`:160-163`, folding the sub-lot's `lots=0`), the read-out
emits the shared part, and the commit loop's part-wide `DELETE_OrderRenban` removes both
blank rows. The sub-lot residual qty (20) is dropped IDENTICALLY on both sides (independent
trace: `compute` emits `SAMEP` lots 3+2=5, `parts_seen={SAMEP}`). So there is **no path
where the rebuild deletes a lots=0 placeholder the legacy preserves, and none where it
preserves one the legacy deletes**:
- sub-lot part DISTINCT from all emitted parts → preserved (both sides);
- sub-lot part SHARES with an emitted part → deleted (both sides).

The `parts_seen`-from-emitted fix is genuinely correct.

## All-lots-0 group edge (dev scope note) — REAL divergence; rebuild is the SAFER side; KEEP-SEED, document it

When EVERY order in the group is sub-lot (all `lots=0`):
- **LEGACY (reachable, harmful):** `LoadScreen` enables `FRSBreakdown_Button` whenever
  `recordcount > 0` (`:607/:632`) — it does NOT gate on TotalLots; an all-sub-lot group
  shows `TotalLots=0` but the button is live. The capacity gate `tcount*pcount >= 0`
  (`:709`) passes trivially. The read-out writes no grid row, so `rcount` stays at its init
  `0` (`:705`), and `fNewMaxRenban := rcount + 1 = 1` (`:798`). The commit then sends
  `Format('%.3d',[1]) = '001'` → **CLOBBERS the group counter to 1** (proven live: stored
  `'001'`). For a group with a climbing counter (CMWA 288, PACF 633…) this resets it to 1 →
  the NEXT real breakdown re-issues `CMWA001, CMWA002…`, COLLIDING with the historical
  `CMWA00x` block. (The lone `DELETE_OrderRenban(@PartNumber='')` on the blank cleared grid
  row deletes only rows with an empty part number — benign; the COUNTER clobber is the harm.)
- **REBUILD (safer):** `orders` is non-empty → no early return; `compute` returns `rows=[]`,
  `next_count = seed3 = 288` via the `last_rcount < 0` branch (`code.py:218-224`);
  `parts_seen=[]` → zero DELETEs; the counter write is `("%03d"%288)[:3]='288'` — an
  effective no-op (re-writes the seed). **Keep-seed.**

**Assessment:** this IS a genuine persisted-counter divergence (rebuild keeps 288, legacy
sets 1), and the legacy path IS reachable (button enabled on any non-empty group). But the
legacy behavior is a **destructive bug** (counter reset → renban collisions), and live data
corroborates it is rare-to-never exercised: all 5 counters are mid-range and climbing
(CAP 068, CMWA 288, DICAS 480, HCAP 088, PACF 633) — none parked near 1, as a routine
all-sub-lot clobber would leave them. The rebuild's keep-seed is in the same family as the
min-anchor "SAFER divergence."

**Recommendation:** KEEP the rebuild's keep-seed (do NOT reproduce the clobber-to-1) and
**document it as a deliberate safer divergence** (like the min-anchor / GetShip-calendar
carry) — same class as the rollover note, but here we intentionally DEPART from the legacy
rather than faithfully carry the bug. Add a one-line note in the spec's §12 divergence
ledger and a guard/alert so an operator running a breakdown on an all-sub-lot group is
warned rather than silently no-op'd. This is a documentation/decision gap, not a code defect
to "fix" — flag to the architects for the divergence ledger.

## No-regression re-confirm

- **Distribution math:** pure suite 35/0 incl. the 10,200-scenario cross-check vs the
  INDEPENDENT `.pas` transcription (`ref_distribute`, not the production code) → 0
  mismatches. The two edits touch only `parts_seen` derivation (`:307-311`) and the counter
  reduction (`:345`); `_distribute_part`/`_Truck`/`_frs_suffix`/`_renban_number` are
  untouched.
- **FRS / renban assignment + FRS-suffix no-op:** re-proven live — INSERT_OpenOrder
  (re-read) appends a max+1 2-digit ordinal to a 7-char `@FRSNum` → 9 chars truncating back
  to varchar(7) = input; E2E stored FRS suffixes `01/02/03` verbatim. Honored, not
  reimplemented.
- **Delete-then-reinsert write-back:** E2E 27/0 — all blank placeholders removed, every
  re-inserted row non-blank renban + blank order-date + status-empty (stock-neutral, on-hand
  unchanged), counter = next_count, idempotent re-run a clean no-op, abandoned `:482-539`
  update-in-place NOT resurrected.

## RE-VERIFY VERDICT

Both round-1 findings are **genuinely RESOLVED with proof** (live-proc + DB-persisted-value
checks, regression guards proven non-vacuous by a scratch-revert E2E that FAILS on the old
code). The distribution, FRS, renban, no-op, and write-back remain faithful (no regression).

So: **the rebuild now reproduces the legacy renban semantics MODULO (a) the golden-pending
gap and (b) the documented rollover-latent-bug carry** — as claimed — **with ONE additional
DELIBERATE, SAFER divergence: the all-lots-0 group counter clobber.** The rebuild does NOT
reset a climbing counter to 1; the legacy does. That is reachable but a destructive legacy
bug and apparently unexercised in production (counters all climbing). RECOMMENDATION: keep
the rebuild's keep-seed and record it in the divergence ledger (architects' call); it is the
only place the rebuild intentionally diverges from the legacy, and it diverges in the SAFE
direction. No BLOCKER remains.
