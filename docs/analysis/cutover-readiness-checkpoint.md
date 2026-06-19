# Cutover-Readiness Checkpoint — InventorySystem Delphi → Ignition

**Date:** 2026-06-19  ·  **Status:** 🟢 **Build phase COMPLETE + cutover package TURNKEY.** The whole live
InventorySystem is analyzed, rebuilt on Ignition Perspective, shadow-wired to the stock ledger, and
parity-closed. David's final 4 cutover go-decisions are made and recorded (§4 — all RESOLVED), and the 3
SQL artifacts those decisions newly required are built + spike-validated (branch `cutover-artifacts`). What
remains before production is the **cutover flip** itself (the parallel-run → live switch), whose plan is
fully designed, adversarially reviewed, artifact-backed, and now **dress-rehearsed end-to-end on the spike**.

**DRESS-REHEARSAL (2026-06-19) — GO.** The entire Phase A→D sequence was run against the spike (fenced by a
pre-backup + post-restore, spike proven as-found): **GO/NO-GO zero-drift gate = 0/47 parts** (`IN_QTY ==
SUM(ledger)` after dropping the 13 triggers + `SEED_AllOpeningBalances`); forward-post smoke proved the
seams are the SOLE `IN_QTY` writer with triggers dropped (no double-count); genesis guard fires correctly.
Full write-up: `docs/analysis/cutover-dress-rehearsal.md`. **One residual to verify on PROD:** the
`IX_INV_MANIFEST_COST_MST` constraint-DROP path could NOT be exercised on the spike (already dropped there),
so the `ALTER TABLE … DROP CONSTRAINT` step is the one cutover action untested-in-rehearsal — confirm on prod.

**The 3 new cutover SQL artifacts (built + spike-validated 2026-06-19):**
- `docs/analysis/master-data/spike-partsstockinfo-drop-qty-clause.sql` (Carry 10)
- `docs/analysis/reporting/priceless-lines-diagnostic.sql` (D6 priceless-lines safety net)
- `docs/analysis/master-data/spike-manifestcost-nooverlap-trigger.sql` (manifest index DROP + no-overlap trigger)

This is the single "where are we" page. Detail lives in the linked docs.

---

## 1. What is built (and merged to master)

| Area | State | PRs |
|---|---|---|
| **Full functional analysis** — every live area spec'd + adversarially verified; D1–D13 decisions | ✅ | #2 |
| **Master-data CRUD** — Supplier, Size, Logistics, RenbanGroup, Parts Stock, Assembly Detail, Manifest Cost | ✅ | #3, #4 |
| **Stock-Ledger service** — `INV_STOCK_LEDGER` + `POST_StockMovement` + `PROC_RebuildStockBalance`; ledger = source of truth, `IN_QTY` = materialized `+= delta` | ✅ | #5 |
| **4 producers wired** — receiving(+) / shipping(−) / reject(−) / stocktaking(±), as pure `computePosts` + thin drivers | ✅ | #6 |
| **Write-then-post seams** — all 4 producers' insert/amend/delete write paths call a producer post-service; event-key widened to varchar(100) | ✅ | #7, #8 |
| **Parity closed (two ways)** — controlled full-history from-zero reconstruction + the opening-balance backfill (`SEED_AllOpeningBalances`) | ✅ | #9, #12 |
| **Order commit/write path** — worksheet → `INV_OPEN_ORDER_INF` (no `IN_QTY` move) | ✅ | #11 |
| **D6 end-to-end** — `fn_ManifestCostAt` TVF + no-overlap write guard + the 4 window-blind report procs migrated | ✅ | #10, #14 |
| **Headless Jython seam-runner** — runs the REAL stocktaking/reject **+ shipping/receiving** drivers via the `system.db` shim (R8) | ✅ | #15, **#18 (open)** |
| **Landing hub + suite theme** — Perspective home + `tai-light`/`tai-dark` | ✅ | #16 |
| **Pre-cutover architect pass** — cutover sequence + adversarial review + runbook resolutions | ⏳ | **#17 (open)** |

**Open PRs awaiting David:** **#17** (architect pass docs) and **#18** (seam-runner extension). Both are
docs/test-only, reviewed, and mergeable.

---

## 2. Verification state

- **Full e2e sweep: 26/26 harnesses green** (`scripts/e2e/test_*.py`). The lone
  `test_stock_ledger_parity` SKIP is **by design** — the restored `.bak` is post-`DELETE_AutoPurge`, so the
  from-zero reconstruction can't run (receiving history aged out). That check is the cutover-backfill
  validator; it runs if a pre-purge dump ever surfaces. Parity is instead closed by the opening-balance
  seed + the controlled-history proof (`test_ledger_fullhistory_recon.py`).
- **Seam-runner** now executes the real driver logic for **all producers AND Order** — the last §7 gap is
  **CLOSED.** Order's `commitOrders` spans statements (begin → N× `INSERT_OpenOrder` →
  `UPDATE_PartsStockRenban` → commit/rollback); `jython_shim.py` gained a **persistent-sqlcmd-session**
  extension (`_TxSession`: one long-lived `docker exec -i` connection fed framed batches) so the REAL
  `commitOrders` runs against a real transaction. `scripts/e2e/test_seam_driver_order.py` (13/13 green)
  proves the happy path (2 records committed, counter advanced, IN_QTY unmoved, persists after close) AND
  **mid-transaction rollback atomicity** — the first INSERT is shown visible *inside* the open tx, then a
  forced failure rolls it back so zero rows persist and the counter is unchanged (the thing an autocommit
  shim could never prove). Autocommit producers unregressed (`test_seam_driver.py` 23/23).
- **Infra:** dev gateway Ignition 8.1.52 (`:8088`), docker `mssql-spike` (SQL Server 2019). Prod targets
  8.3 — code is `# IG83-TODO`/`# IG81-COMPAT` annotated.

---

## 3. The cutover plan (designed, not yet executed)

- **`docs/analysis/cutover-architecture.md`** — the **4-phase / 18-step sequence**:
  - **Phase A** (additive, days ahead): deploy ledger objects + seam libs (not yet the live `IN_QTY` writer). Full rollback = drop new objects.
  - **Phase B** (D6, independently reversible): apply the window-aware report procs + the no-overlap constraint; repoint the UI to `fn_ManifestCostAt`.
  - **Phase C** (destructive, **app DOWN / quiesced**): drop the 13 qty-triggers → run `SEED_AllOpeningBalances` → **validate `IN_QTY == SUM(ledger)` = zero drift (the Go/No-Go gate)** → enable the seams as the live writer → smoke-test.
  - **Phase D:** spec sweep; park the Postgres phase.
  - **Rollback point:** through the green-invariant step; after that, easy rollback is gone.
- **`docs/analysis/cutover-runbook.md`** — the **11 carries** (the deferred-to-cutover checklist), each
  flagged in its PR/code so none is lost. Carries 10/11 were added by the architect pass.
- **`docs/analysis/cutover-review-adversarial.md`** — the adversarial refutation: 2 BLOCKERs + 3
  SHOULD-FIX, **all triaged and resolved**. Notable outcomes:
  - **No live 5th `IN_QTY` producer exists.** `UPDATE_PartsStockInfoCount` is dead; `UPDATE_PartsStockInfo`'s
    absolute `IN_QTY=@QTY` is a superseded clobber → drop that clause at cutover (Carry 10).
  - `UPDATE_PartNumber` is **kept** — it audits every parts-stock row-state change by design (David); not
    a qty-mover.
  - The opening-balance seed **freezes the current (possibly buggy) `IN_QTY`** as the genesis row; only
    forward posts are corrected. The 0-drift test proves the seed *mechanism*, not balance correctness —
    stated honestly (fixture-fidelity discipline).

---

## 4. Decisions David made — ✅ ALL 4 RESOLVED (2026-06-19); cutover package is now TURNKEY

David made the final 4 go-decisions on 2026-06-19. Each is recorded in `cutover-architecture.md` (carry
sections + the "Items needing David's decision" list) and `cutover-runbook.md` (carries 10/11), and each
that newly required a SQL artifact has one built + **spike-validated** (evidence below).

1. ✅ **Carry 10 — ACCEPTED.** Drop the `IN_QTY=@QTY` clause from `UPDATE_PartsStockInfo`; on-hand is
   **NEVER** editable on the rebuilt Parts Stock master (all qty change via a ledger transaction — the
   seam/ledger is the sole `IN_QTY` owner). `UPDATE_PartsStockInfoCount` is DEAD → retired. The `@QTY`
   param stays in the signature (now unused) so the rebuilt Save's positional 30-param call is unchanged.
   **Artifact:** `docs/analysis/master-data/spike-partsstockinfo-drop-qty-clause.sql`.
   **Spike-validated:** baseline legacy proc moved IN_QTY (part 12: 7382→99999); after the ALTER, passing
   `@QTY=11111` left IN_QTY at 99999 while VC_COMMENTS + other columns still wrote; `UPDATE_PartsStockInfoCount`
   confirmed dropped; spike restored to as-found (part 12 → 7382, both procs back).
2. ✅ **Carry 11 — CLOSED, NO WORK.** The Order worksheet only emits NEW orders for the calculated FUTURE
   FRS date; an already-shipped/arrived order cannot exist on that future FRS date, so the Order-commit path
   needs NO ledger post. RESOLVED/closed — not a gap. No artifact (no code change).
3. ✅ **D6 priceless lines — KEEP CROSS APPLY (all 4 procs) + ADD a pre-invoice diagnostic.** CROSS APPLY is
   faithful to the legacy inner JOIN and never emits a $0 line to Toyota; the diagnostic surfaces the lines
   CROSS APPLY would drop so gaps are caught BEFORE billing. **Artifact:**
   `docs/analysis/reporting/priceless-lines-diagnostic.sql` (`REPORT_EDI810_PricelessLines`).
   **Spike-validated:** returns 0 on current data; a fabricated out-of-window unbilled line
   (part 42600F261100 @ 20200101, before its 20250404 window) returned EXACTLY that one line; an in-window
   control line returned 0 (predicate is selective); fabrication rolled back, diagnostic proc cleaned up.
4. ✅ **Manifest index — DROP + REPLACE.** Drop `IX_INV_MANIFEST_COST_MST` at cutover (it is a UNIQUE
   *constraint*, not an index — verified, table is a HEAP) so the rebuilt master can hold multiple windows
   per part; integrity enforced by the app `checkWindowOverlap` guard + a NEW DB trigger
   `TRG_ManifestCost_NoOverlap`. **Artifact:**
   `docs/analysis/master-data/spike-manifestcost-nooverlap-trigger.sql` (pre-drop diagnostic + conditional
   drop + trigger). **Spike-validated:** pre-drop diagnostic = 0 overlaps; gap-window INSERT succeeds (part
   holds 2 windows); overlapping AND touching-boundary INSERTs rejected with the THROW; self-update succeeds
   (no false-reject). Confirmed the existing app-guard test (`test_manifestcost_overlap_guard.py`) fails its
   `OUTPUT`-without-`INTO` anchor INSERT with the trigger ENABLED (SQL Server restriction) and passes 10/0
   with the trigger absent — exactly as the artifact's WRITER-COMPAT note documents; trigger cleaned up.

(`UPDATE_PartNumber` keep/audit-by-design — **already decided**.)

**Column-name surprise vs the architecture doc's draft (Artifact 3):** the arch-doc draft's guessed names
were all CORRECT against the live `INV_MANIFEST_COST_MST` — `IN_MANIFEST_COST_ID` (int IDENTITY),
`VC_ASSY_PART_NUMBER_CODE` (varchar 12), `VC_START_MANIFEST`/`VC_END_MANIFEST` (varchar 8, yyyymmdd STRING).
The genuine surprise (already captured in the artifact + arch doc) is structural: there is **NO primary key**
— the table is a HEAP — and `IX_INV_MANIFEST_COST_MST` is a UNIQUE *CONSTRAINT* on `VC_ASSY_MANIFEST_NUMBER`
(varchar **2**, the 2-char manifest #), so it drops via `ALTER TABLE … DROP CONSTRAINT`, not `DROP INDEX`.
The artifact's section B handles both forms defensively.

---

## 5. Genuinely post-cutover (not blocking the flip)

- **Postgres phase (D13 / `# IG83-TODO`):** 16-char string timestamps → `datetime2`; `site_id` NOT NULL
  FKs + per-site scoping; real PKs/FKs on source tables; typed date compares in `fn_ManifestCostAt`.
- **Atomicity (Carry 1):** fold source-write + `POST_StockMovement` into one transactional unit at
  cutover (moot during parallel-run — the ledger is a shadow).
- **D11#7 renban-counter race (Carry 2):** atomic counter allocation in `order.commitOrders`.
- **Order seam-runner (Carry 7 remainder):** the shim's persistent-session extension.
- **Spec-hygiene sweep (Carry 9):** pre-D9 prose vs the live dump.
- **`IN_BALANCE_AFTER`** is computed pre-post (cosmetic; `IN_QTY` itself stays correct).

---

**Bottom line:** the rebuild is feature-complete and parity-proven on the spike; the cutover is a
designed, reviewed, sequence-fenced operation waiting on four go-decisions and the merge of #17/#18.
