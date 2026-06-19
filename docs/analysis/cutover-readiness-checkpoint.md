# Cutover-Readiness Checkpoint — InventorySystem Delphi → Ignition

**Date:** 2026-06-19  ·  **Status:** 🟢 **Build phase COMPLETE.** The whole live InventorySystem is
analyzed, rebuilt on Ignition Perspective, shadow-wired to the stock ledger, and parity-closed. What
remains before production is the **cutover flip** (the parallel-run → live switch), whose plan is fully
designed and adversarially reviewed; it is gated only on a handful of David go-decisions (below), not on
more build.

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
- **Seam-runner** now executes the real driver logic for **all producers except Order** (Order's
  `beginTransaction` spans statements → needs the shim's persistent-sqlcmd-session extension; the one §7
  item still open).
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

## 4. Decisions David still owes before the flip

1. **Carry 10** — accept dropping the `IN_QTY=@QTY` clause from `UPDATE_PartsStockInfo`, and confirm
   on-hand is never editable on the rebuilt Parts Stock master.
2. **Carry 11** — confirm the Order worksheet can never emit an already-shipped/arrived order at commit
   (else the Order commit path needs a ledger post).
3. **D6 report procs** — CROSS vs OUTER APPLY for priceless report lines (recommend CROSS); final sign-off
   on the intended forward-divergence rows + the D6 proc diffs against live data.
4. **`IX_INV_MANIFEST_COST_MST`** — confirm the prod unique index is the one to drop for the no-overlap
   model.

(`UPDATE_PartNumber` keep/audit-by-design — **already decided**.)

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
