# P4 — Renban Rollover Fix: Clean Wrap + Collision-Aware Allocator (WARN → GUIDE → FIX)

**Status:** BUILT (P4, branch `p4-renban-collision`). The clean wrap + the collision allocator (predicate /
run-of-N / all-path in-tx re-check) + the 2 view shells + the amended spec §12.7 + tests are landed and
green; §6 reflects the SF1/SF2/SF3 corrections as-built. Hand to `sql-adversary` (predicate/run-of-N/wrap) +
`ignition-code-reviewer` (flow/screen/atomicity/TOCTOU/all-path re-check).
**Author:** ignition-architect, 2026-06-23.
**Scope:** the renban breakdown stage (`docs/analysis/order/project-library/renban/code.py`).
**Pairs with (PARALLEL, dependency):** `docs/analysis/order/renban-collision-sourcetruth.md` (sql-analyst)
— the exact **in-use predicate**, the **next-free** SQL, and the **clean-wrap safety** proof. This design
CONSUMES that predicate as a defined input (see §6 "Dependency contract"). We converge in the build.

---

## 0. TL;DR

Two changes to the renban breakdown, both at the *commit* boundary, both solo-dev maintainable, both
reusing proven seams (the `INV_EDI_ALARM_REJ` table, the home-hub alarm surface, the existing one-tx
`commit_renban_breakdown`):

1. **Clean wrap (999→000).** Replace the legacy truncation `('%03d' % next_count)[:3]` with
   `next_count % 1000` (a 3-char zero-padded value). The count rotates the *full* 000-999 ring and wraps to
   the OLDEST block instead of jumping to the recent 100-block. A **deliberate, documented divergence** from
   the legacy — a parallel run differs ONLY at the rollover (and the legacy value there is the buggy one).
   Renban *numbers* stay 3-digit (`CMWA000`, never `CMWA1000`).

2. **Collision-aware allocator (WARN → GUIDE → FIX).** At breakdown-commit, BEFORE the write tx, check each
   candidate renban against the **in-use predicate** (open releases). On a collision: **WARN** (inline
   message + a `RENBAN_COLLISION` row in `INV_EDI_ALARM_REJ`, surfaced on the home hub), **GUIDE** (the
   next-free renban(s) + the colliding open-release details), **FIX** (operator picks: use-next-free /
   override-anyway / cancel; the system re-issues + commits the one tx).

The interaction lives in a **new minimal dialog shell on a new renban-breakdown surface** (today the
breakdown is *driver-only*, invoked from the order flow — there is no screen). The shell is the
headless-authorable minimum; full styling is a Designer follow-on.

**Divergence classification (R17 rule):** the clean wrap CHANGES a number Toyota *could* see (a renban on a
supplier order file at rollover). Per the divergence rule it is **surface-and-document** — but David has
**already decided** it (memory `feedback-warn-guide-fix` P4 names "clean wrap 999→000 PLUS collision-aware
allocator" as the agreed fix). So this is **DECIDE-and-flag**: locked here in this doc's ledger (§9), not a
silent bake, and re-confirmed as open-question Q1 below in case the rollover timing matters for an in-flight
parallel run.

---

## 1. The bug (recap, verified against `renban/code.py:328-345` + spec §12.7)

Legacy advances the group counter and persists `Format('%.3d',[next_count])` (a *minimum*-width pad that
NEVER caps) into `@RenbanCount varchar(3)`, which the proc **left-truncates to 3 chars**:

| next_count | legacy persists | effect |
|---|---|---|
| 634 | `'634'` | fine |
| 1000 | `'1000'` → `'100'` | **collapses to the recent 100-block** |
| 1002 | `'1002'` → `'100'` | same collapse |

So at `next_count ≥ 1000` the group re-seeds from ~100 and the next run's renbans **collide** with the
earlier `CMWA100x` block (degenerate 100-999 cycle + a stray 4-digit boundary renban `CMWA1000` in
`VC_RENBAN_NUMBER varchar(8)`). The rebuild currently REPRODUCES this (`('%03d' % next_count)[:3]`,
`code.py:345`, flagged `# IG83-TODO:` at `:334-339`) for parallel-run parity.

Reachability is live (spec §12.7): PACF 633/634, DICAS 480/484 actively climbing; any group reaching ~994+
with up to 6 trailers crosses 1000 in `next_count`.

---

## 2. Part 1 — The clean wrap

### 2.1 The change point (single line in `commit_renban_breakdown`, step (c))

`renban/code.py:342-345` today:

```python
if next_count is not None:
    system.db.runPrepUpdate(
        "EXEC UPDATE_RenbanGroupCount @RenbanCode=?, @RenbanCount=?",
        [groupCode, ("%03d" % next_count)[:3]], db, tx=tx)
```

Becomes:

```python
if next_count is not None:
    # P4 CLEAN WRAP (divergence D-RNB-1, renban-collision-design.md §2): the count rotates the full
    # 000-999 ring and wraps to the OLDEST block. Replaces the legacy str(N)[:3] truncation (which
    # collapsed 1000->'100' and collided with the recent 100-block). % 1000 is ALWAYS 3 chars after
    # %03d, so it is spec-compliant (varchar(3)) and never emits a 4-digit count.
    persisted_count = "%03d" % (next_count % 1000)
    system.db.runPrepUpdate(
        "EXEC UPDATE_RenbanGroupCount @RenbanCode=?, @RenbanCount=?",
        [groupCode, persisted_count], db, tx=tx)
```

### 2.2 The renban NUMBER must also stay 3-digit (the `CMWA1000` boundary case)

The wrap of the *persisted count* is not sufficient on its own. Within a single breakdown run, the per-truck
renban number is `group_code + ('%03d' % rcount)` where `rcount = seed3 + truck_number`
(`_renban_number`, `code.py:136-141`). If `seed3 = 998` and there are 3 trailers, the emitted rcounts are
998, 999, **1000** → `CMWA1000` (8 chars, fits `varchar(8)`, but it is a 4-digit renban — the same boundary
artifact). The clean wrap must make the renban NUMBER ring-wrap too, so `_renban_number` emits `CMWA998`,
`CMWA999`, `CMWA000`.

`_renban_number` change point (`code.py:136-141`):

```python
def _renban_number(group_code, renban_seed, truck_number):
    seed3 = int(str(renban_seed)[-3:])
    rcount = seed3 + truck_number                    # legacy raw sum (can exceed 999)
    # P4 CLEAN WRAP (D-RNB-1): wrap the renban-number tail to the 000-999 ring so a run that straddles
    # 999 emits CMWA998/CMWA999/CMWA000 (3-digit) instead of CMWA1000 (4-digit). Keeps every renban
    # exactly group_code + 3 digits. rcount (the RAW value) is still returned UNWRAPPED so next_count =
    # max(rcount)+1 carries the true count for the % 1000 persist in step (c) — the two wraps compose.
    wrapped = rcount % 1000
    return group_code + ("%03d" % wrapped), rcount
```

**IMPORTANT — keep `rcount` raw for `next_count`.** `compute_trailer_breakdown` derives
`next_count = last_rcount + 1` from the RAW (unwrapped) rcount (`code.py:208-224`). Leave that logic
untouched: it carries the true running count forward, and the `% 1000` in step (c) does the single
authoritative wrap on persist. Only the *displayed/stored renban string* wraps in `_renban_number`. This
composition (raw count forward, wrap on emit + wrap on persist) is the cleanest: one source of truth for
"how far the group has advanced," two render-time wraps for the two 3-char surfaces.

> **sql-analyst dependency (clean-wrap safety):** the wrap is only safe if the OLDEST block it wraps onto is
> genuinely closed (no open release still carries `CMWA000`...). sql-analyst's "clean-wrap safety" proves the
> ring period (999 renbans) vastly exceeds the open-release window, so a wrap collision is rare — but **the
> collision allocator (Part 2) is the backstop** that catches it if it ever does. The two parts are
> complementary: clean wrap makes collisions rare; the allocator makes the remaining ones safe.

### 2.3 Why divergence, not bug-for-bug parity

Phase-1 parity (`code.py:328-339`) was correct for the *parallel run* (zero side-by-side diff). P4 is the
**post-cutover fix** the IG83-TODO explicitly anticipated ("widen … or alert+block the operator at 999").
We choose neither widen (can't — `varchar(3)` is a SPEC requirement) nor hard-block (dead-end, violates
WARN→GUIDE→FIX). We wrap + guard. A parallel run now differs at exactly one event (the rollover), and the
legacy output there is the defective one. Logged as D-RNB-1 (§9).

---

## 3. Part 2 — The collision-aware allocator (WARN → GUIDE → FIX)

### 3.1 Where it hooks: a pre-commit gate, BEFORE the write tx

The collision check runs **after** `compute_trailer_breakdown` produces `rows` (candidate renbans known)
and **before** `system.db.beginTransaction` (`code.py:313`). It must NOT be inside the tx: the operator
resolution is interactive (it can take seconds-to-minutes) and we never hold a DB transaction open across a
human decision. The atomic one-tx commit (delete-then-reinsert + counter bump) fires only AFTER the
operator's choice is folded into the candidate `rows`.

Flow at the commit boundary:

```
compute_trailer_breakdown(...)            # candidate rows + next_count (PURE, no DB)
        |
        v
check_renban_collisions(rows, db)         # NEW: read-only, consumes the in-use predicate
        |
   any collisions? --- no ---> proceed straight to the existing one-tx commit
        |
       yes
        |
        v
   WARN  (inline message + RENBAN_COLLISION alarm row -> home hub)
   GUIDE (next-free renban(s) + colliding open-release detail)
   FIX   (operator picks: use-next-free | override | cancel)
        |
   re-issue rows with the chosen renban(s)  (PURE re-map, no DB)
   re-run check on the re-issued rows        (the next-free could itself be edge-case taken)
        |
        v
   the existing one-tx commit  (delete-then-reinsert + UPDATE_RenbanGroupCount % 1000)
   + ack/resolve the RENBAN_COLLISION alarm in the SAME tx
```

### 3.2 The collision-detection function (new, in `renban/code.py`)

```python
def check_renban_collisions(rows, database=None):
    """Read-only. For each candidate trailer-row renban, test the IN-USE PREDICATE (sql-analyst's
    renban-collision-sourcetruth.md): does an OPEN release already carry this renban (same group)?
    Returns a list of collision dicts: {renban, candidate_part, open_order_id, open_frs, open_part,
    last_shipped, ...} — empty list = clear to commit. Consumes the predicate as ONE parametrized
    read; we do NOT re-derive 'open' here (sql-analyst owns the exact VC_STATUS_SUPPLIER_SHIPPING set
    + any FRS-date/window cut)."""
```

- Implemented as ONE batched `runPrepQuery` against the predicate over the candidate renbans (not N
  queries), e.g. `... WHERE VC_RENBAN_NUMBER IN (?,?,...) AND <open-predicate>` — exact SQL from
  sql-analyst's source-truth doc. Read-only, outside any tx.
- Returns enough detail to populate GUIDE: the colliding open order's id, FRS, part, and **last-shipped**
  timestamp (so GUIDE can say "last shipped Y").

### 3.3 WARN

Two surfaces, mirroring the M1 inbound pattern exactly (reuse, don't invent):

1. **Inline (the dialog on the breakdown surface):** the breakdown does not silently issue the colliding
   renban and does not hard-fail. The dialog opens stating: *"Renban `CMWA288` would collide with an open
   release (order #4471, FRS 6090102, last shipped 2026-06-19)."*

2. **`RENBAN_COLLISION` alarm row in `INV_EDI_ALARM_REJ`** — the home-hub flag (same table the M1 824/997
   alarms use; the home hub already surfaces `BIT_RESOLVED=0` rows). New `VC_ALARM_TYPE = 'RENBAN_COLLISION'`.
   Column mapping (reuse the existing columns; no schema change):

   | column | renban-collision value |
   |---|---|
   | `IN_SITE_ID` | the session site (M4 marker; same threading as inbound) |
   | `VC_ALARM_TYPE` | `'RENBAN_COLLISION'` |
   | `VC_SUMMARY` | `"Renban CMWA288 collides with open release (order #4471)"` |
   | `VC_MANIFEST_NUMBER` | reuse as the **colliding renban string** (`'CMWA288'`) — varchar(8), exact fit |
   | `VC_ASSY_PART_NUMBER` | the candidate part (varchar(12)) |
   | `VC_ERROR_TEXT` | `"open order #4471 FRS 6090102 last shipped 2026-06-19"` |
   | `IN_EIN` | NULL (824/997-only) |
   | `BIT_RESOLVED` | 0 → flips to 1 when the operator resolves (FIX) |

   Reuse the inbound `_write_alarm(...)` helper shape (`edi_inbound/code.py:424-439`) — a near-identical
   thin INSERT. Add a sibling `_write_renban_alarm(...)` in `renban/code.py` (don't cross-import the
   inbound module; the INSERT is 6 lines and the two services have independent lifecycles).

   > **8.1 note:** the alarm is the *data* row (the proven, headless-testable seam). Wiring a gateway NATIVE
   > `system.alarm` pipeline off these rows is the PROD follow-on (same posture the inbound doc states,
   > `edi_inbound/code.py:426-428`). `# IG83-TODO: wire system.alarm journal off RENBAN_COLLISION rows.`

### 3.4 GUIDE

The dialog presents CHOICES + context, not a dead end:

- **The next-free renban(s)** in the 000-999 ring for this group — from sql-analyst's **next-free** SQL
  (the lowest renban not currently in use by an open release, scanning forward from the candidate). If the
  breakdown needs N sequential renbans (one per trailer), GUIDE shows the next free *run* of N.
- **The colliding open release detail:** which order/ASN, the FRS, the part, and when it was last shipped
  (from `check_renban_collisions`).
- **The three options** (radio/buttons): use-next-free · proceed-anyway · cancel.

### 3.5 FIX

Operator picks; the system applies the choice, then commits the SAME atomic tx:

- **(a) Use the suggested next-free renban.** Re-map the candidate `rows` onto the next-free renban run
  (PURE in-memory re-issue — re-seed `rcount` from the chosen base, re-run `_renban_number` for each truck).
  **Re-run `check_renban_collisions` on the re-mapped rows** before commit (the next-free could itself have
  been claimed by a concurrent breakdown — rare with ~2 operators/site, but cheap to re-check and it closes
  the TOCTOU window). Then commit. Also persist the advanced counter consistent with the chosen base
  (`next_count = chosen_base + max(truck_number) + 1`, then `% 1000`).
- **(b) Override / proceed anyway.** Operator confirms (a second confirm click — they may know the old
  release is closing). Commit with the original candidate renbans. **Log the override** (logger + the alarm
  row stays in history as `BIT_RESOLVED=1` with a note that it was an override, for audit).
- **(c) Cancel.** No write at all. The `RENBAN_COLLISION` alarm row **stays `BIT_RESOLVED=0`** on the home
  hub (a standing reminder to go handle the stale release). The operator goes and closes/ships the stale
  release, then re-runs the breakdown.

On (a) and (b), the `RENBAN_COLLISION` alarm is **acked/resolved (`BIT_RESOLVED=1`) inside the same commit
tx** as the breakdown write — so a rolled-back commit also rolls back the ack (the alarm correctly stays
active if the write fails). On (c), the alarm is intentionally left active.

### 3.6 The commit stays ATOMIC

The existing `commit_renban_breakdown` one-transaction guarantee (`code.py:313-354`: delete-then-reinsert +
`UPDATE_RenbanGroupCount`, rollback on any failure) is UNCHANGED in shape. The collision resolution happens
**before** `beginTransaction`; the only addition INSIDE the tx is the alarm-resolve UPDATE for the (a)/(b)
paths. No partial write is possible: either the operator resolves and the full delete-reinsert-count-ack
commits atomically, or they cancel and nothing is written.

---

## 4. Where the interaction lives (screen vs dialog shell)

**Today: driver-only.** The renban breakdown is a Project Library service (`renban/code.py`) invoked from
the order flow; there is **no Perspective renban screen** (confirmed — no `gen_*renban*` generator; the
landing "Orders & Renban" module card routes into the order area, not a dedicated breakdown view).

**Decision: build the minimal breakdown surface + a collision dialog as part of P4.** The breakdown needs
*some* operator entry point anyway (it takes `groupCode`, `trailers`, `palletsPerTrailer` — the legacy
`RenbanOrder` form's inputs). P4 is the right time to stand up the **minimal headless-authorable shell**;
full styling is a Designer follow-on.

### 4.1 The surface (minimal shell — what the developer builds)

`Order/RenbanBreakdown` Perspective view (single combined view, per our headless limits
`reference-headless-ignition-authoring-limits`):

- **Inputs:** group-code dropdown (Named Query `Order/renbanGroups`), trailers numeric, pallets/trailer
  numeric.
- **Preview button** → calls `compute_trailer_breakdown` (PURE, no DB) via a view script → shows the
  candidate trailer rows in a table.
- **Commit button** → calls the gateway driver path (below). On a collision, the driver returns a
  collision result and the view **opens the collision dialog** (`system.perspective.openPopup`) instead of
  committing.

### 4.2 The collision dialog shell (`Order/RenbanCollisionDialog`)

- **WARN** label (bound to the collision message), a **GUIDE** table (next-free + colliding-release rows),
  three **FIX** buttons (Use next-free / Proceed anyway / Cancel).
- "Proceed anyway" requires a second confirm (a `confirm`-style two-step, or a checkbox-gates-the-button).
- Each button invokes the matching driver function (`commit_renban_breakdown` with a `resolution` arg —
  see §5) and closes the popup. "Use next-free" passes the chosen base; "Proceed anyway" passes
  `override=True`; "Cancel" just closes (no driver call) and the alarm stays active.

### 4.3 Headless-authoring constraints (8.1.52)

Per `reference-headless-ignition-authoring-limits`:

- **Named Query `data.bin`** (the XML) and **page-route mount** require the Designer — list them as the
  Designer hand-off items (the `Order/renbanGroups`, `Order/nextFreeRenban` NQs; the `/order/renban`
  page route; the popup registration). The developer authors the view JSON + driver headlessly, then a
  short Designer pass mounts them.
- **Bidirectional binding** (the group-code dropdown writing back) goes in `binding.config` (NOT a
  top-level prop) and commits on Tab/blur — note it for the dropdown + the numeric inputs.
- Prefer `system.db.runPrepQuery` in view/driver scripts over a Designer-only NQ where it keeps the view
  self-contained and headless-testable (the source-truth predicate + next-free can be `runPrepQuery`
  strings owned by the driver, mirroring how `loadBlankRenbanOrders` already inlines
  `_SELECT_NO_RENBAN` rather than calling the proc — `code.py:239-268`).

---

## 5. Driver API shape (the contract the view calls)

Extend `commit_renban_breakdown` to be collision-aware via a `resolution` parameter (default = the
pre-commit check; explicit = the operator's FIX choice). Keep the existing signature backward-compatible
(the e2e tests `scripts/e2e/test_renban_*.py` pass `groupCode, trailers, palletsPerTrailer` positionally).

```python
# AS-BUILT signature (backward-compatible: positional e2e calls still work; siteId/actor optional kwargs):
def commit_renban_breakdown(groupCode, trailers, palletsPerTrailer, database=None, _orders=None,
                            resolution=None, siteId=None, actor=None):
    """resolution:
         None                      -> compute + check_renban_collisions; if collisions, return WITHOUT
                                      writing: {"status": "COLLISION", "collisions": [...], "rows": [...],
                                      "next_free": <base or None>, "alarm_id": <RENBAN_COLLISION row>}
                                      (WARN already written: the alarm row + the inline payload).
                                      No collisions -> straight commit, status "COMMITTED" (the in-tx
                                      re-check can still flip it to COLLISION on a lost race).
         {"action":"use_next_free", "base": <int>, "alarm_id": <id>}
                                   -> re-map rows onto base, in-tx re-check, commit, ack alarm.
         {"action":"override", "alarm_id": <id>, "acknowledged": [<renban>,...]}
                                   -> commit ORIGINAL rows, in-tx re-check EXCLUDING `acknowledged`
                                      (the seen-colliding set the operator chose to reuse) so only a
                                      NEWLY-taken number aborts; ack alarm with an OVERRIDE audit note.
                                      `acknowledged` defaults to a fresh pre-check if omitted (degraded).
         (cancel is purely client-side: no driver call; alarm stays BIT_RESOLVED=0 / active.)"""
```

- The `None` path WRITES the `RENBAN_COLLISION` alarm (WARN) and returns the collision payload — it does
  NOT open the write tx.
- `use_next_free` / `override` open the existing one-tx commit, fold the resolution in, ack the alarm in-tx.
- This keeps the PURE compute and the human decision OUTSIDE the tx, and the write strictly atomic.

> **Re-check note:** the `use_next_free` path MUST re-run `check_renban_collisions` after re-mapping (the
> TOCTOU window). If the re-check still collides (concurrent breakdown grabbed the next-free), return a
> fresh COLLISION payload — the dialog re-GUIDEs. Bounded: with ~2 operators/site this almost never loops.

---

## 6. Dependency contract with sql-analyst (`renban-collision-sourcetruth.md`)

**RESOLVED + BUILT (SF2/SF3 corrections folded in).** We consumed exactly these three deliverables:

| # | sql-analyst deliverable | how the build consumes it |
|---|---|---|
| P1 | **In-use predicate** — the **RESIDENT-ROWS, STATUS-INDEPENDENT** definition, EXACT equality: `EXISTS (SELECT 1 FROM INV_OPEN_ORDER_INF WHERE VC_RENBAN_NUMBER = @renban)`. NO status filter, NO `LIKE grp+'%'`. | `check_renban_collisions` (`renban/code.py`): ONE batched `runPrepQuery` `… WHERE VC_RENBAN_NUMBER IN (?,?,…)` over the candidate renbans. |
| P2 | **Next-free RUN-of-N** — the lowest base whose `[base..base+N-1] % 1000` are ALL free (a multi-trailer breakdown needs N **contiguous** suffixes), forward-ring scan from the breakdown's starting count; exhaustion → None. | `next_free_run(group, start_count, n, db)` — the gaps-and-islands / forward-ring scan; the GUIDE base + the `use_next_free` base. |
| P3 | **Clean-wrap safety** — the lap ≈ retention finding (CMWA ~863/1000, ~1mo runway): the wrap can land on a still-resident old block, so the allocator is **LOAD-BEARING, not optional**. | the allocator IS the backstop; clean wrap makes the count VALID, the allocator keeps it from REUSING a live number. |

**SF2 — the predicate is SELF-SAFE by construction (corrected framing).** The earlier draft worried about
a breakdown colliding "against the very rows it is about to delete-and-reinsert." That cannot happen: the
breakdown's INPUT rows are the **blank** placeholders (`VC_RENBAN_NUMBER=''`), and the candidates are
**non-blank** (`group + 3-digit`). An exact-equality `= @candidate` never matches a `''` row, so there is
**no self-collision and no own-row exclusion needed** — and we deliberately add **no status filter** (an
ordered/shipped-but-unpurged renban is still occupied; a status filter would UNDER-detect it). The
**resident-rows** definition is the strict/correct form: a number is free only when `DELETE_AutoPurge` has
removed its rows. (The re-inserted rows are status-empty / stock-neutral, `code.py:34`; they become
collision targets for a FUTURE breakdown only once they are resident, which they are immediately after
commit — exactly right.)

**SF1 — TOCTOU on EVERY commit path (built).** The pre-tx check runs BEFORE `beginTransaction` (never hold
a tx across the human decision); the upstream Order commit (`order/code.py:130`) is a concurrent non-blank
renban writer with NO unique-constraint backstop, so `_commit_rows_tx` re-runs the resident-rows predicate
**inside** the tx before the first INSERT on ALL THREE paths (`use_next_free`, `override`, no-collision
straight-commit). On a lost race → rollback + a fresh COLLISION payload (re-WARN). The **override** path
passes an `acknowledged` set (the renbans the operator SAW colliding at the WARN, from the COLLISION
payload) so the in-tx re-check excludes the deliberately-reused numbers but **still aborts on a DIFFERENT
number taken newly since the WARN** — keeping the guard meaningful on the one path that intentionally reuses
in-use numbers. The alarm-ack (`BIT_RESOLVED=1`) is INSIDE the same tx (a rolled-back commit keeps the alarm
active). All proven headless end-to-end (`test_renban_collision_e2e.py` §7a/7b/7c) by seeding a resident row
from a SECOND connection while the driver's tx is open (faithful READ-COMMITTED lost race).

---

## 7. 8.1.52 vs 8.3 notes

- `system.db.beginTransaction / runPrepUpdate / runPrepQuery` — identical on 8.1.52 and 8.3
  (`code.py:43` IG81-COMPAT). The collision path adds no new 8.3-only DB API.
- `system.perspective.openPopup` / `closePopup` — available on 8.1.52 (Perspective 8.1). Fine.
- `RENBAN_COLLISION` alarm is a DATA row (`INV_EDI_ALARM_REJ`) on both versions. The NATIVE
  `system.alarm` journal off these rows is `# IG83-TODO:` (prod), same posture as the inbound alarms.
- No schema change: `INV_EDI_ALARM_REJ` is reused with the existing columns (§3.3 mapping). No
  `_HIST`/site-scoping surgery needed for P4 (the table is rebuild-owned and already carries `IN_SITE_ID`).
- Named Query `data.bin` XML + page-route mount + popup registration require a Designer pass (8.1 + 8.3) —
  the headless hand-off items (§4.3).

---

## 8. Phased build sequence

1. **Clean wrap (Part 1)** — `_renban_number` ring-wrap + step (c) `% 1000`; update the IG83-TODO comment to
   reference D-RNB-1; update the unit test oracle (derive from the SPEC ring-wrap, NOT the rebuild — R15).
   Low-risk, self-contained. **Adversarial review:** number-semantics → full dual-adversary
   (sql-adversary ∥ ignition-code-reviewer) + double re-verify (it changes a renban Toyota could see).
2. **`check_renban_collisions`** — once sql-analyst lands P1/P2; one batched read; unit-tested against a
   seeded open release.
3. **`_write_renban_alarm` + the `resolution` driver path** — WARN write, FIX re-map + ack-in-tx, the
   TOCTOU re-check. e2e against the spike (`scripts/e2e/test_renban_*`).
4. **The Perspective shell** — `Order/RenbanBreakdown` + `Order/RenbanCollisionDialog` view JSON
   (headless) → Designer pass mounts NQs/route/popup. e2e Playwright drive of the dialog (per
   `project-e2e-perspective-testing`).
5. **Home-hub surface** — add `RENBAN_COLLISION` to the active-alarm NQ/rail the hub already shows.

**Review cadence (R19):** Part 1 + the FIX commit path = money/number-adjacent → full dual-adversary +
double re-verify. The dialog shell / WARN surfacing = display-ish → lighter single review.

---

## 9. Divergence ledger (P4)

| id | divergence | class | decision |
|---|---|---|---|
| **D-RNB-1** | Persisted count + renban number wrap `% 1000` (999→000) instead of legacy `str(N)[:3]` truncation | **number Toyota could see** (renban at rollover) | **DECIDE-and-flag.** David pre-decided the clean-wrap fix (memory `feedback-warn-guide-fix` P4). Locked here; re-confirm via Q1 if a parallel run is still live at a rollover. |
| **D-RNB-2** | NEW collision-aware allocator (WARN→GUIDE→FIX) — the legacy silently issued colliding renbans | **safer / behavior-add** (catches a defect the legacy shipped) | **DECIDE-and-flag.** Pre-decided (same memory). Safer-than-legacy; no number changes unless the operator chooses use-next-free, which is then an explicit operator action (logged). |

---

## 10. Open questions / handoffs

**For David:**
- **Q1 (rollover timing):** D-RNB-1 changes the renban at the 999→000 boundary. If any group will be at a
  rollover *during a still-live parallel run* (PACF 633/634 is the nearest), the side-by-side will diff at
  that one event. Confirm: is the parallel run complete for the renban path, or should the clean wrap be
  gated behind a cutover flag until it is? (Default assumption: parallel run is done → ship the wrap.)
- **Q2 (override audit):** on "proceed anyway" (FIX-b), is logger + a resolved alarm row enough audit, or do
  you want a dedicated override note column/row? (Default: reuse `VC_ERROR_TEXT` to stamp "OVERRIDE by
  <user> <ts>"; no schema change.)
- **Q3 (next-free direction):** when no free *run* of N sequential renbans exists below the candidate,
  should next-free scan forward past the candidate (wrapping the ring), or surface "no free run — close a
  release first"? (Default: scan the full ring forward-from-candidate; if truly none free, that's a WARN
  with cancel-only — extremely unlikely given the 999-ring vs open-window.)

**For sql-analyst (dependency, converge in build):**
- P1/P2/P3 per §6. The ONE convergence item: reconcile the in-use "open" status set with the freshly-
  grouped status-empty rows so a breakdown does not collide against the very rows it is about to
  delete-and-reinsert (§6 last paragraph).

**For ignition-developer:** §2 change points (exact), §5 driver API, §4 the shell + headless hand-off list.
**For adversarial-architect-reviewer:** focus on §3.1 (gate outside tx), §3.6 (atomicity), §5 TOCTOU
re-check, and §6 the open/freshly-grouped reconciliation (the subtle correctness seam).
