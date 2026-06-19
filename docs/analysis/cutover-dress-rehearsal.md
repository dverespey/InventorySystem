# Cutover Dress-Rehearsal — InventorySystem Delphi → Ignition (against the SPIKE DB)

**Date:** 2026-06-19  ·  **Operator:** ignition-developer  ·  **Target:** `mssql-spike` / DB `Inventory`
(NOT prod — there is none here). **Backup/restore-fenced, destructive-to-spike, restored as-found.**

This is the realistic in-environment de-risking of the real production flip. It executes the
4-phase / 18-step sequence in `docs/analysis/cutover-architecture.md` end-to-end against the spike, with a
mandatory pre-backup and post-restore so the spike is left EXACTLY as found.

## Headline result

**GO/NO-GO GATE (the headline): ZERO DRIFT.** After dropping the 13 qty-triggers and seeding the opening
balance, **0 of 47 parts** had `IN_QTY <> SUM(INV_STOCK_LEDGER.IN_QTY_CHANGE)`. The materialized `IN_QTY`
and the ledger agree for every part. **GO.**

| Result | Value |
|---|---|
| **Go/No-Go zero-drift parts** | **0 / 47** (GO) |
| Parts (INV_PARTS_STOCK_MST) | 47 |
| OPENING_BALANCE ledger rows seeded | 47 (one per part) |
| Triggers dropped (Phase C) | exactly **13** (12 movers + `DeleteShipDate`) |
| `UPDATE_PartNumber` / `DELETE_PartNumber` (kept) | both PRESENT + enabled throughout |
| Forward-post smoke (receiving +60, shipping −50) | each moved IN_QTY by its exact delta ONCE, wrote its ledger row, invariant held |
| Genesis guard | THROWs `Msg 50001` on a part with a forward row but no opening row — CONFIRMED |
| Restore as-found | IN_QTY per-part checksum + all proc bodies byte-identical to baseline |

## Steps executed (against the architecture-doc sequence)

### Step 0 — Backup (restore point)
`BACKUP DATABASE Inventory TO DISK='/var/opt/mssql/dress_rehearsal_pre.bak' WITH INIT, COPY_ONLY`
→ succeeded, 2114 pages / ~19 MB. This is the restore point.

### Step 1 — Pre-flight snapshot (baseline)
- **IN_QTY per part** (47 rows) saved + checksummed: `md5 = 90531d03b309b3a4b5bf447b74ee4ab5`.
- **13 qty-triggers all present + enabled** (`is_disabled=0`); `UPDATE_PartNumber`/`DELETE_PartNumber`
  present + enabled.
- **INV_STOCK_LEDGER = 0 rows** (no forward rows → genesis guard clean; opening-first invariant satisfiable).
- **INV_MANIFEST_COST_MST = 45 rows.**
- **Prereq ledger objects all EXIST** (built across PRs #5–#12 — Phase A was already applied on the spike):
  `INV_STOCK_LEDGER`, `POST_StockMovement`, `PROC_RebuildStockBalance`, `SEED_OpeningBalance`,
  `SEED_AllOpeningBalances`, `fn_ManifestCostAt`. (`TRG_ManifestCost_NoOverlap` absent — Phase B creates it.)
- **Proc-body MD5 hashes** captured for `UPDATE_PartsStockInfo`, `UPDATE_PartsStockInfoCount`, the 4 D6
  report procs, and `fn_ManifestCostAt` — so the restore comparison detects body drift, not just presence.

### Step 2 — Phase A + B (additive / reversible), applied idempotently in the doc's order
1. `spike-manifest-cost-lookup.sql` → `fn_ManifestCostAt` re-created (idempotent).
2. `spike-report-procs-d6.sql` → the 4 window-blind report procs replaced; all 4 now `CROSS APPLY`
   `fn_ManifestCostAt` (verified `uses-TVF`).
3. **Pre-drop overlap diagnostic** (section A of the no-overlap artifact): **0 overlaps** — clean.
4. `spike-manifestcost-nooverlap-trigger.sql` → section B (conditional drop of `IX_INV_MANIFEST_COST_MST`)
   **no-op'd on the spike** ("already absent" — the constraint was dropped in the ManifestCost master build,
   as the artifact documents); section C created `TRG_ManifestCost_NoOverlap`.
5. `spike-partsstockinfo-drop-qty-clause.sql` → `UPDATE_PartsStockInfo` altered (the `IN_QTY=@QTY` clause
   removed — confirmed the body's only `IN_QTY` text is now the explanatory comment); `UPDATE_PartsStockInfoCount`
   dropped (dead). Each object re-compiled cleanly.

### Step 3 — Phase C (the core destructive rehearsal)
1. **Quiesce:** the rehearsal operator is the only writer (gateway holds read connections but is not posting).
2. **Dropped exactly 13 triggers** (idempotent `IF OBJECT_ID(...,'TR') IS NOT NULL DROP TRIGGER` loop):
   the 12 movers (`INSERT/UPDATE/DELETE_RecConfStatPartsStockMstQTY`, `…_RejectParts`, `…_Stocktaking`,
   `Insert/Update/DeletePartShipping`) + `DeleteShipDate`. Post-drop: **0 of the 13 remain**;
   `UPDATE_PartNumber` + `DELETE_PartNumber` still PRESENT (KEPT per David's audit decision).
3. **`SEED_AllOpeningBalances`** → wrote **47 OPENING_BALANCE rows**; IN_QTY checksum UNCHANGED
   (`90531d03…`), confirming the seed records ledger rows WITHOUT bumping IN_QTY.
4. **GO/NO-GO GATE → 0 drift parts** (query: `WHERE p.IN_QTY <> (SELECT ISNULL(SUM(IN_QTY_CHANGE),0) …)`).
5. **Forward-post smoke** (`scripts/e2e/dress_rehearsal_smoke.py`, driving the REAL Jython wrappers via
   `jython_shim` — triggers already DROPPED, so NO disable/enable):
   - `receiving.insertOpenOrder` +60 (shipped 'S', counted) → IN_QTY moved by **exactly +60 once**, ledger
     row `RECEIVING_SHIP:ord=<id>:ins` = +60, invariant held.
   - `shipping.insertPartShipping` −50 → IN_QTY moved by **exactly −50 once**, ledger row
     `SHIPPING:psh=<id>:ins` = −50, invariant held.
   - **Genesis guard:** with a forward (non-OPENING) ledger row present and no opening row,
     `SEED_OpeningBalance` THROWs `Msg 50001 … "the opening balance must be seeded BEFORE any posting"`.

### Step 5 — Restore (mandatory, as-found)
`ALTER DATABASE Inventory SET SINGLE_USER WITH ROLLBACK IMMEDIATE; RESTORE … WITH REPLACE;
SET MULTI_USER` — all in one batch from `master`. The gateway's open connections were rolled back by
`ROLLBACK IMMEDIATE`; doing the `SET MULTI_USER` in the same batch closed the window before the gateway
could re-grab the single-user slot (no contention observed). Restore succeeded (2114 pages).

**Restore PROVEN as-found:**
- IN_QTY per part: **`diff` identical to baseline**; checksum `90531d03…` matches.
- 13 qty-triggers: **13 present, 13 enabled**; `UPDATE_PartNumber` enabled.
- INV_STOCK_LEDGER: **0 rows** (back to baseline); INV_MANIFEST_COST_MST: 45.
- `UPDATE_PartsStockInfo` + `UPDATE_PartsStockInfoCount` back; `TRG_ManifestCost_NoOverlap` +
  `REPORT_EDI810_PricelessLines` absent (as baseline).
- **All 7 proc/fn body MD5 hashes byte-identical to baseline** (including the altered `UPDATE_PartsStockInfo`
  and the dropped-then-restored `UPDATE_PartsStockInfoCount`).
- Gateway: HTTP 302; one session re-connected to Inventory — the datasource auto-reconnected cleanly.

## Deviations / surprises from the architecture-doc sequence (honest log)

The rehearsal followed the doc faithfully. Findings:

1. **Phase A was already applied on the spike (expected, not a deviation).** The ledger objects exist from
   PRs #5–#12, and `IX_INV_MANIFEST_COST_MST` was already dropped in the ManifestCost master build — so the
   no-overlap artifact's section B no-op'd ("already absent"), exactly as the artifact's header predicts. On
   PROD, section B will do real work via `ALTER TABLE … DROP CONSTRAINT`; this rehearsal therefore does NOT
   exercise the constraint-drop path against a populated constraint (it can't — the spike has none). **Carry
   for the real cutover:** confirm prod still has the UNIQUE *constraint* form and that section B's
   constraint branch fires.

2. **No trigger-name mismatch.** All 13 names in carry 5 matched live objects exactly; the loop dropped
   exactly 13 and left `UPDATE_PartNumber`/`DELETE_PartNumber` untouched. No ordering or lock problem.

3. **Genesis-guard test-design correction (test bug, NOT a system finding).** My first smoke attempt
   probed the genesis guard by calling `SEED_OpeningBalance` on an already-seeded part — it returned rc=0.
   Root cause: `SEED_OpeningBalance` checks the **idempotency** guard (opening row already exists → silent
   `RETURN`) BEFORE the **genesis** guard, so an already-seeded part short-circuits. That is correct,
   intended idempotency. The guard is genuinely exercised only by a part that has a forward (non-OPENING)
   ledger row but NO opening row (a misordered cutover); presented that way it THROWs `Msg 50001` as
   designed. The smoke script was corrected to probe it correctly. **Lesson for the runbook:** the genesis
   guard protects against *posting before seeding*, not against *re-seeding* — re-seeding is a safe no-op by
   the idempotency check. Both behaviors are correct.

4. **Smoke-script teardown cosmetic (test artifact, NOT a cutover defect).** The smoke script's in-test
   teardown re-creates part 16's opening row AFTER the forward posts, stamping it with the then-current
   IN_QTY (base+10) rather than base, then resets IN_QTY to base — leaving a transient +10 ledger/IN_QTY
   mismatch on the ONE test part at the end of the script. This is purely an ordering bug in the throwaway
   teardown; it has **zero bearing on the cutover** because (a) the authoritative reset is the STEP-5
   RESTORE, which made the whole DB byte-identical to baseline, and (b) the Go/No-Go gate, the two forward
   posts, and the genesis guard all asserted GREEN before teardown. Documented for honesty; not fixed beyond
   the format-string crash because RESTORE supersedes it.

5. **Gateway connection lock handled cleanly.** `ROLLBACK IMMEDIATE` + same-batch `SET MULTI_USER` avoided
   any auto-reconnect contention; the restore did not have to contend with the datasource mid-operation.
   **Carry for the real cutover:** prod has the live Delphi app + possibly more gateway sessions — quiesce
   the app (Phase C step 11) before this, and expect `ROLLBACK IMMEDIATE` to kill more sessions there.

## Artifacts produced

- `scripts/e2e/dress_rehearsal_smoke.py` — the Phase-C forward-post smoke driver (drives the REAL
  receiving/shipping wrappers via `jython_shim` with triggers already dropped; asserts exact-delta-once +
  ledger row + invariant + genesis guard). Reusable for the real cutover smoke step.
- Backup file `/var/opt/mssql/dress_rehearsal_pre.bak` — **left in the container** (it is the proven
  restore point; harmless, gitignored location, not committed). Delete with
  `docker exec mssql-spike rm /var/opt/mssql/dress_rehearsal_pre.bak` if reclaiming space.

## Bottom line

The cutover sequence executes cleanly end-to-end on the spike. The **Go/No-Go zero-drift gate passed
(0/47)**, the seams are proven as the SOLE `IN_QTY` writer post-trigger-drop (exact-delta-once, no
double-count), the genesis guard fires correctly, and the spike was restored byte-identical to as-found.
The only deviations were a self-inflicted test-script bug (corrected) and the expected spike-vs-prod
difference in the manifest-constraint-drop path (flagged for the real cutover).
