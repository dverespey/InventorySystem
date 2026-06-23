# Adversary Findings — Renban Collision Rebuild (P4, branch `p4-renban-collision`)

SQL-semantics adversarial review of the clean-wrap ring + collision-aware allocator
(`docs/analysis/order/project-library/renban/code.py`) against the legacy `RenbanOrder.pas`
and the spec oracle (`renban-breakdown-spec.md` §12.7). Goal: find the input where the
rebuild's number / state-change differs from legacy or from its own GUIDE.

- Verification env: `mssql-spike` docker, DB `Inventory` (writable, rolled-back probes only),
  `Inventory_Live` / `VehicleOrder` READ-ONLY. `SET QUOTED_IDENTIFIER ON; SET NOCOUNT ON`.
- All DB probes below were `BEGIN TRAN ... ROLLBACK` (zero leftover rows verified) or pure
  in-memory simulations against the real `code.py`.
- Baselines run GREEN as shipped: `test_renban_build.py` 35/35, `test_renban_e2e.py` 31/31,
  `test_renban_collision_e2e.py` 24/24.

---

## BLOCKER-1 — `use_next_free` writes onto UNVALIDATED suffixes for a wrapped candidate run (silent collision + self-collision next run)

**Claim under test:** `next_free_run` returns a base whose `[base..base+N-1] % 1000` are ALL
free, and `_remap_rows_onto_base` re-seats the trailer rows onto exactly that validated run
(ATTACK #3 contiguity / ATTACK #2 no-over-detect-on-the-FIX-path).

**Defect:** `_remap_rows_onto_base` (code.py:384-407) derives each row's run offset from the
*rendered renban tail* as `offset = (tail - min(tails)) % 1000`. When the candidate run
straddles the 999→000 wrap, the rendered tails are e.g. `[997, 998, 999, 0, 1]`, so
`min(tails) = 0` (the post-wrap low), giving offsets `[997, 998, 999, 0, 1]` instead of the
intended `[0, 1, 2, 3, 4]`. The remap then writes
`(base + offset) % 1000 = [base-3, base-2, base-1, base, base+1]` — a contiguous block shifted
**3 below** the run `next_free_run` actually validated.

**Counterexample (pure, real `code.py`):**
```
rows tails = [997,998,999,0,1];  next_free_run picks base=310 (validated [310,311,312,313,314] free)
_remap_rows_onto_base(rows,'CMWA',310) -> renbans CMWA307, CMWA308, CMWA309, CMWA310, CMWA311
   -> WRITES suffixes [307,308,309,310,311], NONE of {307,308,309} were ever checked for freeness
```

**DB-level proof (Inventory, rolled back, 0 leftover):** seed one resident occupant at
`ZZWR308`; GUIDE validates `[310..314]` (`resident_in_validated = 0` → "free run"); the remap
writes `[307..311]` (`resident_in_written = 1`, collides on `ZZWR308`). The operator is told
"use 310-314"; the system attempts 307-311.

**Second-order damage (persist):** `new_next_count = base + max_offset + 1 = 310 + 999 + 1 =
1310` → persisted `'%03d' % (1310 % 1000) = '310'`, while the highest suffix actually written is
`311`. The NEXT breakdown of the group seeds from `310` and re-issues `CMWA310`/`CMWA311` →
**guaranteed self-collision** with this run's own rows.

**Reachability:** requires (a) the candidate run to straddle 999 (the exact CMWA-canary rollover
the allocator exists for — count near 997 + ≥2 trailers, real per `sourcetruth §4`), (b) a
collision, and (c) the operator picking `use_next_free` (the GUIDE's own recommendation). The
in-tx re-check (BLOCKER-mitigation, runs on the actual `new_rows`) *may* abort if 307-309 are
resident — but that is a backstop, not correctness: when 307-309 are free the bad block commits
silently AND the count persists wrong (self-collision next run); the operator-facing GUIDE
number never matches what the DB receives, breaking WARN→GUIDE→FIX integrity.

**Why the suite is GREEN anyway (test gap):** the only wrapped-run e2e
(`test_renban_e2e.py:285-304`) forces the wrap through the **`override`** path, which commits the
ORIGINAL candidate rows and never calls `_remap_rows_onto_base`. The `use_next_free`
path is only e2e-tested on a NON-wrapped run (`test_renban_collision_e2e.py:182`, base 302). The
wrapped × use_next_free cell is untested.

**Classification:** code defect (the offset/`min`/`next_count` model assumes a non-wrapped run).
The fix belongs to the architect/developer — the run base for offset derivation must be the
run's *forward-lowest* (the candidate's pre-wrap start, `_candidate_base` already holds the right
notion: lowest emitted tail in scan order is the rcount-min, not the numeric min), and
`new_next_count` must be `base + (N-1) + 1`, not `base + max_offset + 1`.

---

## SHOULD-FIX-1 — `next_free_run` / `_all_resident_suffixes` UNDER-detects a trailing-space resident renban (latent GUIDE-onto-occupied)

**Claim under test:** the resident-suffix scan that feeds `next_free_run` registers every
occupied suffix (ATTACK #3 no-under-detect; ATTACK #2 padding trap).

**Defect:** `_all_resident_suffixes` (code.py:347-361) keys occupancy on
`int(RIGHT(VC_RENBAN_NUMBER, 3))`. For a renban stored with a trailing space (`'ZZRB289 '`,
DATALENGTH 8 / LEN 7), the `LIKE G+'[0-9][0-9][0-9]'` + `LEN = len(G)+3` guard **passes**
(verified: `'ZZRB289 ' LIKE 'ZZRB[0-9][0-9][0-9]'` → MATCH, `LEN = 7`), but
`RIGHT('ZZRB289 ', 3) = '89 '` and Python `int('89 ') = 89`. So suffix **289** is recorded as
**089** — `next_free_run` then believes 289 is FREE and can hand the operator a base whose run
includes the actually-occupied 289 → GUIDE-onto-occupied collision.

Note the *equality* predicate `_RENBAN_IN_USE` is NOT affected — SQL Server's trailing-space-
insensitive `=`/`IN` means a candidate `'ZZRB289'` correctly matches the stored `'ZZRB289 '`
(verified: EQUAL, `caught_by_IN = 1`). So `check_renban_collisions` is safe; only the
`RIGHT()`-based GUIDE scan is wrong.

**Reachability:** LATENT. `Inventory_Live` has **0** trailing-space renbans
(`DATALENGTH<>LEN` count = 0) and the rebuild's own writer (`'%03d' % wrapped`) never emits one.
No legacy path currently produces them. Flagged because the GUIDE silently degrades if dirty data
ever appears, and the LEN guard gives false confidence. Fix: `int(LTRIM(RTRIM(suf)))` or scan via
the numeric `VC_RENBAN_NUMBER = G + zeropad3(n)` form instead of `RIGHT()`.

---

## SHOULD-FIX-2 — `override` default-acknowledged auto-acks numbers the operator never saw

**Claim under test (ATTACK #4):** override excludes the operator's SEEN-colliding set from the
abort decision; a DIFFERENT renban taken since the WARN still aborts.

**Defect:** when the dialog does NOT pass `acknowledged`, `commit_renban_breakdown`
(code.py:589-591) recomputes it as `set(c["renban"] for c in check_renban_collisions(rows, db))`
**at override time** (after the WARN). Any renban grabbed by a concurrent writer between the WARN
and the override click is now in-use, so the recompute folds it into `acknowledged` — the in-tx
re-check then EXCLUDES it and does NOT abort. The operator silently reuses a number they never
saw colliding (the opposite of the design intent stated in the code comment 582-588).

**Reachability:** depends on the Perspective dialog ALWAYS passing the WARN's collision set back
(the e2e always passes `acknowledged=` explicitly — code's `None` branch is never exercised by
any test). It is a foot-gun default + an untested branch, not a proven live defect, IF the UI
contract holds. Fix: the `None` default should be the EMPTY set (acknowledge nothing unless the
dialog names it) or the driver should require the explicit set on the override path.

---

## SHOULD-FIX-3 — the "in-tx re-check on the TX connection" claim is false (overclaimed TOCTOU guard)

**Claim under test (ATTACK #5):** the in-tx re-check runs ON THE TX CONNECTION so it sees the
tx's view (code.py comment 472-473).

**Finding:** `_commit_rows_tx` (code.py:477) calls `check_renban_collisions(rows, db)` passing
**`db`, not `tx`** → `runPrepQuery(..., db, tx=None)` → the shim's autocommit path (a separate
connection), NOT the open transaction. Functionally it still detects *committed* concurrent
writers under READ COMMITTED (which is why 7a/7b/7c pass — the seeded residents are committed),
so the guard is not broken, but:
- The re-check holds NO lock on the candidate rows and there is NO unique constraint on
  `VC_RENBAN_NUMBER` (confirmed: only `varchar(8)`, collation `SQL_Latin1_General_CP1_CI_AS`, no
  unique index). The window between the re-check (step 0) and the INSERTs (step b) is narrowed
  but NOT closed — a writer committing in that gap is not caught.
- The e2e cannot exercise the genuine mid-tx race; it seeds the resident row BEFORE the driver
  call, so it only proves "a row committed before step 0 is caught," which the pre-tx check
  already would. The truly concurrent window is untested.

This is strictly safer than the legacy (which had no re-check at all and the same missing
constraint), so it is not a parity regression — but the comment overclaims and the residual race
should be documented (or closed with `WITH (UPDLOCK, HOLDLOCK)` on the re-check, or a filtered
unique index, both of which are architect calls).

---

## Confirmed-CORRECT (attacker could not break these)

- **#1 clean-wrap ring math.** Full persist+reseed cycle simulated on real `code.py`: seed
  `CMWA997` × 5 trailers → renbans `CMWA997/998/999/000/001` (no 4-digit `CMWA1000`),
  `next_count` RAW `1002`, persist `'%03d' % (1002 % 1000) = '002'`; next run seeds `CMWA002`,
  emits `CMWA002/003/004`, persist `'005'`. 999 used exactly once, no skip, no dup, contiguous
  across the boundary. A non-rollover run is byte-identical (seed 288 → CMWA288..292, count 293).
  Oracle is derived from spec §12.7 ring math via a SEPARATE `ring3`/`ring_renban` helper, not the
  rebuild (R15-clean); cross-checked the persisted value against the actual DB write in
  `test_renban_e2e.py`.
- **#1 `next_count` = legacy `fNewMaxRenban`.** Legacy `:798 fNewMaxRenban := rcount+1` uses the
  LAST-emitted (truck,part) rcount; the rebuild uses `max(rcount over emitted incl qty=0 rows)+1`.
  Because `rcount = seed3 + TruckNumber` is monotone in TruckNumber and empty trucks emit no inner
  iteration, last-emitted == max == `seed3 + (highest non-empty truck)`. The rebuild tracks rcount
  BEFORE the qty=0 skip (code.py:220-224), matching the legacy read-out which sets rcount on every
  grid cell. Equal on every probed case (skip-empty, trailing-empty, multi-part). Distribution
  matches an independent .pas transcription across 10,200 scenarios.
- **#2 in-use predicate exactness/self-safety/no-under-detect.** `_RENBAN_IN_USE` uses exact `IN`
  equality, NOT `LIKE G+'%'`: verified the 4-digit artifact `ZZRB2880` does NOT match candidate
  `ZZRB288` (over-match avoided). Self-safe: the breakdown's own input rows are blank
  (`VC_RENBAN_NUMBER=''`), candidates are non-blank, so `= @candidate` never hits them.
  Status-independent: the predicate has no order-date/status filter and the e2e's resident seed is
  order-dated (`VC_ORDER_DATE='20260101'`) yet still detected — an ordered/shipped-but-unpurged
  renban correctly reads in-use (no under-detect). Trailing-space stored renban still matches the
  candidate under `=`/`IN` (SQL Server padding) — safe direction. Live: groups not in the master
  (`SPR`,`16H`,`18M`,`18S`, a bare `072`, `CMWA1000`×2) exist but are never candidates (combobox is
  master-driven), so they cannot false-collide a real group.
- **#3 `next_free_run` run-of-N (non-wrapped).** Interior-occupied avoidance, 999→000 wrap-straddle,
  off-by-one (exactly-N-free returns base; N-1-free returns None), exhaustion → None,
  self-overlap guard (N=1000 with 1 used → None) all correct across targeted probes. Exhaustion
  returns None → hard WARN-cancel (never a silent wrap into a live number). [The wrapped-run failure
  is in the REMAP, BLOCKER-1, not in next_free_run itself.]
- **#4 override acknowledged (explicit-set path).** When the dialog passes the WARN's collision
  set (as the e2e and production dialog do), a deliberately-reused renban commits and a
  newly-taken non-acknowledged renban aborts (e2e 7c: override 301, 302 grabbed → abort). Correct.
  (The hole is only the `None` default — SHOULD-FIX-2.)
- **#5 in-tx re-check fires on a committed lost race on all three paths** (e2e 7a/7b/7c GREEN); no
  path inserts without the step-0 re-check; the alarm ack is in the same tx (rolled back on
  failure). (Caveats: SHOULD-FIX-3.)

---

## VERDICT

**The rebuild is NOT proven equivalent / safe — there is one BLOCKER divergence.**

1. **Clean-wrap spec ring reproduced exactly?** YES — persist `% 1000`, renban tail `% 1000`,
   raw `next_count` carried forward; 999 used once, no skip/dup/4-digit, non-rollover byte-identical,
   reseed resumes correctly. Oracle is spec-derived and non-vacuous.
2. **In-use predicate detects every real collision, no under/over-detect?** YES for the equality
   predicate (`check_renban_collisions`): exact, self-safe, status-independent, padding-safe. The
   `RIGHT()`-based GUIDE scan (`_all_resident_suffixes`) under-detects a trailing-space renban
   (SHOULD-FIX-1) — latent only (0 such rows live).
3. **`next_free_run` run-of-N contiguous, wrap-correct, off-by-one-clean, exhaustion-safe?** YES for
   the scan itself. BUT the FIX that *consumes* it — `_remap_rows_onto_base` — mis-seats a WRAPPED
   candidate run onto suffixes the scan never validated AND persists a wrong `next_count`
   (**BLOCKER-1**): a real GUIDE-onto-occupied collision + a self-collision on the next run, in the
   exact CMWA rollover scenario the allocator exists to protect.
4. **Override slips a real new collision / aborts a legit reuse?** NO on the explicit-acknowledged
   path (correct). The `None` default auto-acks numbers taken since the WARN (SHOULD-FIX-2,
   untested branch, UI-contract-dependent).
5. **In-tx re-check on every path, no constraint-less slip?** Re-check runs on every commit path and
   catches committed races; but it runs on an autocommit connection (not the tx, contra the comment),
   holds no lock, and there is no unique constraint, so the narrow mid-tx window remains open and
   untested (SHOULD-FIX-3) — still strictly safer than the legacy.

**Bottom line:** the clean-wrap ring and the equality-based detection are sound and faithful; the
`use_next_free` GUIDE/FIX path has a wrap-handling defect (BLOCKER-1) that can silently issue a
colliding renban to a sub-supplier and corrupt the rollover counter. Do not ship the
`use_next_free` resolution path until `_remap_rows_onto_base` (offset model + `new_next_count`) is
fixed and a wrapped-run × use_next_free e2e is added (revert-proven non-vacuous).

---

# RE-VERIFICATION OF FIXES (branch `p4-renban-collision`, 2026-06-23)

Re-attack of the fixes to BLOCKER-1 + SF-1/2/3 on `mssql-spike` / DB `Inventory` (writable, all
probes `BEGIN TRAN … ROLLBACK` or autocommit-then-DELETE with 0-leftover verified;
`Inventory_Live`/`VehicleOrder` untouched; `SET QUOTED_IDENTIFIER ON`). Baseline:
`test_renban_collision_e2e.py` now **41/41 PASS**; spike restored as-found after every probe
(0 sentinels / 0 group / 0 alarms / no temp index).

## BLOCKER-1 — FIXED (confirmed, non-vacuous)

`_remap_rows_onto_base` (code.py:409-445) now derives each row's offset from its **0-based rank in
emission order** (distinct renban first-seen), and `new_next_count = base + (N-1) + 1`. Verified
against the REAL `code.py` across every run shape I could construct (`/tmp/probe_remap.py`,
`/tmp/probe_emission.py`):

| run shape (emission tails) | base | written suffixes | next_count | result |
|---|---|---|---|---|
| WRAPPED `[998,999,0]` | 310 | `[310,311,312]` | 313 | EXACTLY the validated block; no shift |
| N=1 `[999]` | 5 | `[5]` | 6 | correct |
| shared-truck `[998,998,999,0]` (2 parts on truck0) | 50 | `[50,50,51,52]` | 53 | shared truck shares suffix; `next_count=base+N` |
| non-contiguous trailer ordinals `[300,302,303]` | 700 | `[700,701,702]` | 703 | ranks compress the gap → contiguous (matches `next_free_run(N=3)`) |
| partial-lot skip (truck0 emits nothing) `[999(t1),000(t2)]` | 700 | `[700,701]` | 702 | first-EMITTED → base; contiguous |

Coupling proven sound: the GUIDE computes `n = len(set(rows.renban))` (code.py:669) and
`_remap_rows_onto_base` writes `N = len(distinct rows.renban)` — **same key over the same `rows`**,
and within one breakdown the rendered tails `(seed3+truck)%1000` are distinct for trailers 1..6, so
`n == N` always. The remap therefore writes EXACTLY the contiguous block `next_free_run` validated,
for wrapped and non-wrapped runs alike. Emission order = truck-outer (code.py:213-227), so rank-0 =
the first-emitted truck → `base`; the persisted count seeds the next run PAST every suffix written
(no self-collision).

**Non-vacuity (revert-proven, `/tmp/probe_revert.py`):** restoring the OLD tail-derived model
(`offset = (tail − min(tails)) % 1000`, `next_count = base + max_offset + 1`) on the wrapped run
`[998,999,0]` base=310 writes `[308,309,310]` (≠ validated `[310,311,312]`) and persists `310`
(≠ 313) → the e2e WRAPPED-9 block-equality AND persisted-count assertions both FAIL. The e2e oracle
(`expected_block` from a standalone `ring3` helper, `expected_persist = ring3(base+3)`,
test lines 389-405) is spec-ring-derived, independent of `_remap_rows_onto_base` (R15-clean).
No residual mis-map found in any probed run shape.

## SHOULD-FIX-1 — FIXED (confirmed on the DB, non-vacuous)

`_all_resident_suffixes` (code.py:352-386) now matches structurally:
`LEFT(.,g)=group` + `SUBSTRING(.,g+1,3) NOT LIKE '%[^0-9]%'` +
`LTRIM(RTRIM(.)) = group + SUBSTRING(.,g+1,3)`, keying occupancy on the exact 3-digit positional
field. DB-verified (rolled back, 0 leftover):
- `'ZZQX289 '` (DATALENGTH 8, trailing space) → suffix **289** (not 089). `int('289')` clean.
- `'ZZQX290'` (clean) → 290 (no legit renban newly missed).
- `'ZZQX1000'` (4-digit artifact) → **EXCLUDED** (`LTRIM(RTRIM('ZZQX1000'))='ZZQX1000' ≠ 'ZZQX'+'100'`).
- On REAL `CMWA` data: `CMWA100` counted as 100, `CMWA1000` artifact excluded, distinct suffix-set =
  **863** (matches the documented ~863/1000 CMWA occupancy → reads true occupancy, no under-detect).

**JDBC-faithful collation check:** I reproduced the dev's reason for abandoning the
`LIKE @grp + N'[0-9][0-9][0-9]'` form. Binding the group param as **nvarchar** (the JDBC string-bind
behavior), `'ZZQX289 ' LIKE N'ZZQX[0-9][0-9][0-9]'` returns **MISS** while the `LEN=7` guard still
reads OK (false confidence) — the exact silent-miss the dev described. The chosen structural form
returns **CAUGHT** on the same row. So the chosen form is correct under the gateway's nvarchar
binding; the abandoned form was genuinely broken.

**Non-vacuity:** the old `int(RIGHT('ZZQX289 ',3))` = `int('89 ')` = 89 → suffix recorded as 089 →
289 looks free → GUIDE-onto-occupied. The new `SUBSTRING(.,5,3)` = `'289'`. Revert breaks the e2e-11
assertion. Reachability remains LATENT (0 trailing-space renbans live; the rebuild's `'%03d'` writer
never emits one) — fix is a forward guard, not a live defect.

## SHOULD-FIX-2 — FIXED (confirmed, fail-closed, non-vacuous)

The override branch (code.py:653-654) now does
`acknowledged = resolution.get("acknowledged"); acknowledged = set() if acknowledged is None else
set(acknowledged)` — **no `check_renban_collisions` recompute** in the branch (verified by reading
the region; the only collision recheck on the override path is the in-tx one inside `_commit_rows_tx`,
which is correct). `_commit_rows_tx` does `ack = set(acknowledged or ())` and the in-tx re-check
`recheck = [c for c in check_renban_collisions(rows, db, tx=tx) if c["renban"] not in ack]`
(code.py:505/519) — so None → empty ack → excludes NOTHING → ANY still-in-use candidate aborts.

**Runtime proof (`/tmp/probe_sf2_runtime.py`, real driver via the shim tx):** override with the
`acknowledged` key entirely ABSENT against a still-resident `ZZRB301` → `status=COLLISION`, **0 rows
written** (fail-closed). The explicit-set path (e2e-12a) still commits the deliberate reuse; a
newly-taken non-acknowledged number still aborts (e2e-7c/12b). **Non-vacuity:** the OLD model would
recompute `acknowledged = {'ZZRB301'}` (still resident) → exclude it → COMMIT — e2e-12b asserts
COLLISION, so reverting the empty-default flips 12b to FAIL.

## SHOULD-FIX-3 — re-check now genuinely on the TX connection; RESIDUAL window real & quantified

**FIXED (the plumbing claim is now TRUE):** `_commit_rows_tx` calls
`check_renban_collisions(rows, db, tx=tx)` (code.py:519); the signature threads `tx` (308) into
`runPrepQuery(..., db, tx=tx)` (331). The shim's `_TxSession` holds ONE persistent sqlcmd connection
across batches (jython_shim.py:215-232, `BEGIN TRANSACTION`), and BOTH `runPrepQuery(tx=tx)`
(line 391-393 → `tx._batch`) and `runPrepUpdate(tx=tx)` (line 403-404 → `tx.exec_update`) route to
THAT connection. So the in-tx re-check and the subsequent INSERTs now share one transaction on one
connection — the previously-overclaimed comment is now accurate (it previously passed `db`, a separate
autocommit connection). The re-check sees committed concurrent writers under READ COMMITTED.

**RESIDUAL window — real, demonstrated, and quantified:**

1. **No constraint / no lock backstop.** Confirmed on `INV_OPEN_ORDER_INF`: the ONLY unique index is
   `PK_INV_OPEN_ORDER_INF (IN_ORDER_ID)` — none on `VC_RENBAN_NUMBER` (`varchar(8)`,
   `SQL_Latin1_General_CP1_CI_AS`). A plain unique index would be **WRONG**: legitimately-multiple
   resident rows share a renban (`DICAS154`×3, `CMWA627`×3, … — the merged multi-part-per-trailer
   case). The re-check `SELECT … IN (…)` under READ COMMITTED holds **0 S-locks after it returns**
   (`sys.dm_tran_locks` = 0 KEY/RID/PAGE S-locks for the SPID post-SELECT), so it places no durable
   guard on the candidates.
2. **Window demonstrated (`/tmp/probe_sf3_midtx.py`).** Open the breakdown tx → step-0 re-check on the
   tx connection sees the candidate FREE → a SEPARATE autocommit connection INSERTs+COMMITs the same
   renban in the gap → the breakdown's INSERT (step b) proceeds → **2 resident rows carry the same
   renban**, nothing stopped it. The gap = the DELETE loop (code.py:528-531) between re-check (519)
   and first INSERT (537): a handful of statement round-trips (sub-ms to low-ms).
3. **Reachability of the concurrent writer (the Order commit, order/code.py:130).** The Order commit
   writes a counter renban `kanban + %03d` (e.g. `16H006`) ONLY for a **non-grouped lot-sized** part;
   a part WITH a renban group is born blank-renban → goes down the breakdown path, not the counter
   path. For the two writers to share a ring a non-grouped part's kanban must EQUAL a group code.
   Group codes are `CAP/CMWA/DICAS/HCAP/PACF`. Kanbans `CAP` and `HCAP` exist and equal group codes —
   BUT both their parts (`42602YY09000`/`42602YY05000`) carry a renban group (`IN_RENBAN_ID` 12/9) and
   `BIT_LOT_SIZE_ORDERS=1` (=palletized, the flag is INVERTED) → breakdown path, NOT counter path. So
   in CURRENT data `KANBAN_EQ_GROUP=0` for non-grouped parts → the cross-writer collision is
   **not currently reachable**; it becomes reachable only on a misconfiguration (a non-grouped
   lot-sized part whose kanban equals a live group code) or two concurrent breakdowns of the same
   group.

**RECOMMENDATION — accept-and-document for the near-single-operator site, with a tripwire; do NOT add
a plain unique index, and do NOT add bare `UPDLOCK,HOLDLOCK` as-is.** Rationale, evidence-backed:
- The window is sub-ms, the only concurrent ring-sharing writer is structurally absent in current
  data (no non-grouped part with a group-code kanban), and the site is ~1-2 operators. The in-tx
  re-check already eliminates the common "committed-before-step-0" race. Residual exposure is a
  genuine but vanishingly-rare lost race.
- A **plain unique index on `VC_RENBAN_NUMBER` is incorrect** (multiple resident rows legitimately
  share a renban — proven above). A *filtered* unique index can't express "one per renban" either,
  for the same reason.
- **`UPDLOCK,HOLDLOCK` on the re-check, as-is, is a footgun**: with NO index on `VC_RENBAN_NUMBER`
  the predicate scans and takes a `RangeS-U` lock on **4286 keys (the whole 4284-row table)** for the
  life of the breakdown tx — measured via `sys.dm_tran_locks`. That serializes the breakdown against
  EVERY writer to `INV_OPEN_ORDER_INF` (every Order commit) and invites blocking/deadlocks — far
  worse than the rare race it closes. It only becomes viable WITH a non-unique covering index on
  `VC_RENBAN_NUMBER` (measured: lock footprint collapses to ~11 candidate KEY locks). That is a
  two-part change (add index + add hints) and an architect call.
- **Disposition:** document the residual window in `commit_renban_breakdown` (the code already does,
  code.py:516-518) and add a cheap post-commit tripwire — a duplicate-renban detector (the home-hub
  alarm surface already exists) so a lost race is *caught after the fact* even if not prevented. If
  concurrency ever rises (multi-site, multiple production-control operators on one group), revisit
  with `WITH (UPDLOCK, HOLDLOCK)` **plus** a non-unique `IX_INV_OPEN_ORDER_INF_RENBAN` (architect +
  sql-analyst). Both are out of this reviewer's scope to design/build.

## RE-VERIFICATION VERDICT

- **BLOCKER-1 — FIXED.** Offset-from-rank writes EXACTLY the validated contiguous block and
  `next_count = base+N` on wrapped, N=1, shared-truck, non-contiguous, and partial-skip runs; no
  residual mis-map; revert-proven non-vacuous; spec-derived oracle.
- **SHOULD-FIX-1 — FIXED.** Structural SUBSTRING scan reads 289 (not 089), rejects the 4-digit
  artifact, misses no legit renban (863 real CMWA suffixes), correct under JDBC nvarchar binding;
  revert-proven. Latent reachability unchanged.
- **SHOULD-FIX-2 — FIXED.** Override-None defaults to the empty set (fail-closed); runtime-proven to
  abort on a still-resident un-acknowledged number; explicit-set path still commits the deliberate
  reuse; revert-proven.
- **SHOULD-FIX-3 — FIXED (plumbing) + RESIDUAL window real (NIT, accept-and-document).** The re-check
  now genuinely runs on the tx connection; the narrow mid-tx race remains open (no lock, no usable
  constraint), is demonstrated, and is structurally unreachable in current data. Recommend
  accept-and-document + a post-commit duplicate tripwire; do NOT add a plain unique index; only adopt
  `UPDLOCK,HOLDLOCK` together with a non-unique covering index (architect call).

**Bottom line:** BLOCKER-1 + SF-1/2/3 are now SOUND. The rebuild detects every real collision
(exact-equality predicate, status-independent, self-safe — confirmed in the prior pass and unchanged)
and issues only validated renbans on the `use_next_free`/override/straight-commit paths (the wrapped
GUIDE→FIX defect is closed). The only residual is the SF-3 mid-tx race window — narrow, currently
unreachable, strictly safer than the legacy (which had no re-check at all) — recommended for
accept-and-document with a post-commit tripwire rather than a heavyweight lock/constraint. The
collision allocator is, on the available data, sound to ship.
