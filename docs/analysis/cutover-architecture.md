# Cutover Architecture — decided resolutions for the 9 deferred carries + the end-to-end sequence

**Status:** 🟢 architect pass (pre-cutover). Companion to `docs/analysis/cutover-runbook.md` (the carry
checklist) and `docs/analysis/inventory-stock/IGNITION-stock-ledger-design.md` (the ledger design v2).
**This is a DESIGN pass — no build code changed here.** It turns each runbook carry into a concrete,
decided resolution and orders them into one executable cutover sequence for a solo dev.

**Scope correction up front (read this first — it changes what "flip reads" means).** The ledger design
materializes `IN_QTY` (it is a cached `SUM(qty_delta)`). Every legacy reader —
`SELECT_PartsStockInfo`, `SELECT_PartsStockInfoOrder`, `REPORT_LogicalInventory`, the order-explosion
reads, and the Delphi app — reads `INV_PARTS_STOCK_MST.IN_QTY` directly, and that column STAYS the read
surface after cutover. So **there is no reader rewrite at cutover.** What actually flips is the **writer**:
today 12 triggers own `IN_QTY`; after cutover the producer seams own it. "Flip reads to the ledger" in
the runbook is shorthand for **"flip the writer of `IN_QTY` from triggers to seams"** — the readers are
untouched because the materialized column is the contract on both sides. I use **"writer-flip"** below to
avoid the misleading "read-flip" framing. (A true compute-on-read projection is the Postgres-phase Q1
fork, not a cutover item.)

A second correction: the ledger design §9 mentions a "derive one movement per live source row" backfill;
the runbook §4 and `spike-seed-opening-balance.sql` chose the **opening-balance** backfill instead
(David: "start clean, go forward" — the from-zero reconstruction is impossible, receiving history is
purged). The opening-balance path is the decided one; §9's per-source-row replay survives ONLY as the
parity-harness derivation (a read-only check), never as the cutover writer. I treat that as settled.

---

## Carry 1 — Atomicity (source-write ↔ ledger-post)

**Decision: single stored proc per op (the "write-and-post proc"), NOT a threaded gateway `txId`.**

**Design.** Today each seam does two separate statements: `runPrepUpdate(INSERT/UPDATE/DELETE source)`
then `stockLedger.post(...)` → `POST_StockMovement`. A crash between them orphans the source row from the
ledger. Collapse each op into ONE DB stored proc that does the source write AND the ledger post inside one
`BEGIN TRAN`. The proc internally reuses the exact body of `POST_StockMovement` (same idempotency check,
same additive `IN_QTY += delta`) — factor that body into an internal `#post` block or just inline it, so
there is one definition of "append + bump". Shape for one producer (reject insert):

```
CREATE PROC WRITE_RejectInsert @partId int, @qty int, @division varchar, @reason varchar, @site int=1
AS
BEGIN
  SET NOCOUNT ON;
  BEGIN TRY
    BEGIN TRAN;
      INSERT INV_REJECT_INF (...) VALUES (...);          -- source write
      DECLARE @rejId int = SCOPE_IDENTITY();
      -- inline the POST_StockMovement core: idempotency check on (partId, 'REJECT:rej=<rejId>:ins'),
      -- INSERT INV_STOCK_LEDGER, UPDATE INV_PARTS_STOCK_MST SET IN_QTY += -@qty
    COMMIT TRAN;
    SELECT @rejId AS newId;
  END TRY
  BEGIN CATCH IF @@TRANCOUNT>0 ROLLBACK TRAN; THROW; END CATCH
END
```

The Jython seam becomes a thin `createSProcCall("WRITE_RejectInsert", db)` returning the new id. The
pure `computePosts` logic stays in Jython for the cases where the post is a *function of before/after rows*
(updates, part-changes, header-delete cascades): for those, the seam computes the post list in Jython,
then calls a `WRITE_*` proc that takes the source-write params **plus** the pre-computed `(delta, enum,
eventKey, reason)` and does both writes in one tran. (Insert is the simple single-delta case above;
update/delete pass the computed post(s) in.)

**Why.** (a) The DB owns atomicity natively — exactly what the legacy triggers gave us; no dependency on
Ignition's `system.db.beginTransaction` semantics differing across 8.1.52/8.3 (the ledger design already
chose proc-owned atomicity for `post()` for this reason — §5 gateway-transaction note). (b) A threaded
`txId` would require every seam to open a gateway tran and thread it through `stockLedger.post` — but
`createSProcCall` **cannot join a gateway transaction** (it's why Order had to drop to `EXEC … via
runPrepUpdate(tx=tx)`). Threading `txId` therefore forces rewriting `post()` away from `createSProcCall`
to a raw `EXEC POST_StockMovement` on the shared `tx`, losing the clean proc-call wrapper. The single-proc
pattern keeps the wrapper and is simpler for a solo dev. (c) Idempotency still holds: the `(partId,
eventKey)` UNIQUE backstop and the existence check move inside the same proc — a retried whole op
no-ops on the ledger AND the source insert is inside the same rolled-back tran, so no orphan either way.

**Open-risk.** This adds ~4 `WRITE_*` procs per producer (insert/amend/delete/header-cascade) — ~14 new
procs. That is the right place for the logic (mirrors the legacy trigger atomicity) but it grows the proc
surface. Mitigation: name them `WRITE_<Producer><Op>` so they sort together and obviously pair with the
seam. **Moot during parallel-run** (ledger is a shadow); these are pure cutover artifacts. **No David
decision needed** — recommended pattern is unambiguous.

**Folded note — `IN_BALANCE_AFTER` race (SHOULD-FIX 5, cosmetic).** `POST_StockMovement` computes
`IN_BALANCE_AFTER` from a pre-post `SELECT IN_QTY … + @delta` (a read separate from the additive UPDATE).
Two concurrent posts on the same part can each read the same pre-value and stamp a wrong *running balance*
in that diagnostic column. **`IN_QTY` itself stays correct** — the additive `IN_QTY = IN_QTY + @delta`
serializes on the row lock and is commutative; only the snapshot column is racy. **Disposition: accept
as-is** (optional diagnostic, not the balance). Do NOT verify "monotonic replay" off `IN_BALANCE_AFTER`
under concurrency. The Phase-C quiesce (single writer) removes exposure during the seed/flip; normal
forward operation is the only exposure and it is non-load-bearing. If ever needed, recompute via
`OUTPUT inserted.IN_QTY` inside the additive UPDATE (`# IG83-TODO`, not a cutover gate). *(Cross-ref:
the trigger-drop↔seam-enable double-count/gap hazard SHOULD-FIX 4 raised is already fully fenced by the
Phase C quiesce-window ordering — steps 11–16 + carry 4/5; no additional design needed.)*

---

## Carry 2 — D11#7 renban-counter race (Order commit)

**Decision: atomic `UPDATE … OUTPUT` that reserves the N counter values in one statement, inside the
existing Order transaction. NOT a SQL Server SEQUENCE.**

**Design.** `order.commitOrders` today does: read `IN_RENBAN_COUNT`, compute renbans in Jython, write the
advanced count back. Two specialists race → duplicate renbans. Replace the read with an atomic
reserve-N: the number of counter values a worksheet line consumes is **known before the write** (it is the
lot count for lot-sized parts; 0 for palletized-with-group; 1 for palletized-fallback — `computeOrderRecords`
already computes `newCount`, hence the consumed span). Reserve that span atomically with a new proc:

```
CREATE PROC RESERVE_RenbanCount @partNum varchar, @n int, @startCount int OUTPUT
AS
BEGIN
  SET NOCOUNT ON;
  UPDATE INV_PARTS_STOCK_MST
     SET @startCount = IN_RENBAN_COUNT,                       -- the value the FIRST reserved renban uses
         IN_RENBAN_COUNT = ((IN_RENBAN_COUNT - 1 + @n) % 999) + 1   -- advance by N with the >999->1 rollover
   WHERE VC_PART_NUMBER = @partNum;
END
```

The `UPDATE` takes the row's exclusive lock; the read-and-advance is one statement, so two concurrent
commits serialize and get **disjoint** counter spans. `commitOrders` calls `RESERVE_RenbanCount` FIRST
(getting `startCount`), passes that as the `renbanCount` arg into the pure `computeOrderRecords` (which
already lays N sequential renbans from a start value with the same rollover), then inserts the orders. The
counter write-back at the end is **deleted** — the reserve already advanced it. Do the reserve inside the
existing `beginTransaction` so an order-insert failure rolls back the reservation too (no burned counter
values on rollback).

**Multi-site (D1).** The renban counter lives on `INV_PARTS_STOCK_MST`, which is per-site (the part row
itself is site-scoped). The reserve keys on `VC_PART_NUMBER` within the site's part row, so the counter is
naturally per-(site,part). At the Postgres phase when `site_id` becomes an enforced column, the
`RESERVE_RenbanCount` WHERE gains `AND site_id = @site` — flagged `# IG83-TODO` but the single-site
parallel-run reserve is already correct because each site's part is a distinct row.

**Why not a SEQUENCE.** (a) The renban counter is **per-part with a 1..999 rollover**, not a single global
monotonic counter — a SEQUENCE is global and can't express per-part rollover without one sequence object
per part (thousands of them). (b) The counter value is **persisted on and read from `INV_PARTS_STOCK_MST`
by the legacy app and the worksheet** during parallel run; a sidecar SEQUENCE would diverge from that
column. (c) `UPDATE … SET @out = col, col = f(col)` is a one-statement atomic read-modify-write that needs
no new object and is identical on 8.1.52/8.3. (Note: SQL Server has no row-returning `UPDATE … OUTPUT` of
the *prior* scalar into a variable the way Postgres `RETURNING` does, so the `@out OUTPUT` form above is
the portable expression — `OUTPUT deleted.IN_RENBAN_COUNT` into a table var also works if a multi-row
batch ever reserves for many parts at once.)

**Open-risk.** The rollover arithmetic `((c-1+n) % 999)+1` must exactly match the legacy per-step
`_bumpRenban` over N steps. Validate with a unit test that loops `_bumpRenban` N times from a start and
asserts equality with the closed-form, including a span that crosses 999. **No David decision needed** —
but the rollover-equivalence test is a build gate.

---

## Carry 3 — D6 report-proc cutover apply

**Decision: apply order = fn (already live) → D6 procs → drop the part-blind manifest UNIQUE index (only
if still present) → add a DB-level non-overlap CHECK-via-trigger backstop → repoint the UI lookup. The
non-overlap backstop is a trigger, not a CHECK constraint (a constraint can't see other rows).**

**Design.**
1. **Prereq (verified present):** `dbo.fn_ManifestCostAt` (PR #10, `spike-manifest-cost-lookup.sql`). The
   D6 procs `CROSS APPLY` it, so it must exist first. In the live dump it is not yet applied to prod —
   apply `spike-manifest-cost-lookup.sql` before the D6 procs.
2. **Apply `spike-report-procs-d6.sql`** — this `DROP/CREATE`s `REPORT_INVOICESSummary`,
   `REPORT_MonthlyINVOICESSummary`, `REPORT_EDI810`, `REPORT_EDI856`, replacing the window-blind JOIN with
   the `CROSS APPLY fn_ManifestCostAt(..., a.VC_PRODUCTION_DATE)`. It also drops the hardcoded
   `IN_ASN_EIN=6440` site bug in the EDI856 `@EIN!=0` branch.
3. **The part-blind UNIQUE index — ✅ RESOLVED (David 2026-06-19): DROP + REPLACE.** Live schema
   (verified, line 867) has `IX_INV_MANIFEST_COST_MST UNIQUE (VC_ASSY_MANIFEST_NUMBER)` — unique on the
   **2-char manifest number alone**, NOT on the assy part code. **SCHEMA SURPRISE (verified on the spike):
   this is a UNIQUE NONCLUSTERED *CONSTRAINT* (`CONSTRAINT [IX_INV_MANIFEST_COST_MST] UNIQUE NONCLUSTERED`),
   NOT a standalone `CREATE INDEX`, and the table has NO primary key — it is a HEAP.** So at cutover it must
   be dropped with `ALTER TABLE … DROP CONSTRAINT`, not `DROP INDEX`. On the spike it was already dropped in
   the ManifestCost master build (the table is a pure heap now). At cutover, DROP it (conditionally, handling
   both the constraint and index forms) so the rebuilt master can hold >1 manifest #/window per part; the
   real integrity (non-overlap) moves to the trigger below. **Artifact:
   `docs/analysis/master-data/spike-manifestcost-nooverlap-trigger.sql`** (section B = the conditional drop).
4. **DB-level non-overlap backstop (the app guard's backstop).** The app guard `checkWindowOverlap` (in
   the ManifestCost master Save action, already built + tested) enforces the gap rule at write. The DB
   backstop must catch a write that bypasses the app (Delphi during parallel run, a manual SQL fix, a
   future second client). **A CHECK constraint cannot express this** — overlap is a relationship *between
   rows*, and CHECK sees only the row being written. Implement it as an **INSTEAD OF / AFTER INSERT,UPDATE
   trigger** on `INV_MANIFEST_COST_MST` that rejects a row overlapping any existing window for the same
   part:

   ✅ **RESOLVED (David 2026-06-19) — built + validated. Artifact:
   `docs/analysis/master-data/spike-manifestcost-nooverlap-trigger.sql`** (section C). The drafted column
   names below were GUESSES; **all VERIFIED CORRECT against the live spike**:
   `IN_MANIFEST_COST_ID` (int IDENTITY, no PK — heap), `VC_ASSY_PART_NUMBER_CODE` (varchar 12),
   `VC_START_MANIFEST`/`VC_END_MANIFEST` (varchar 8, yyyymmdd STRING compares). Final trigger:

   ```
   CREATE TRIGGER dbo.TRG_ManifestCost_NoOverlap ON dbo.INV_MANIFEST_COST_MST AFTER INSERT, UPDATE AS
   BEGIN
     SET NOCOUNT ON;
     IF EXISTS (
       SELECT 1 FROM inserted i
       JOIN dbo.INV_MANIFEST_COST_MST m
         ON  m.VC_ASSY_PART_NUMBER_CODE = i.VC_ASSY_PART_NUMBER_CODE
         AND m.IN_MANIFEST_COST_ID     <> i.IN_MANIFEST_COST_ID
         AND NOT (i.VC_END_MANIFEST < m.VC_START_MANIFEST OR i.VC_START_MANIFEST > m.VC_END_MANIFEST))
     BEGIN
       ROLLBACK TRANSACTION;
       THROW 50010, 'Manifest cost window overlaps an existing window for this part.', 1;
     END
   END
   ```

   This is the exact predicate the app guard uses (`NOT (:end < VC_START OR :start > VC_END)`), enforced
   server-side. Because the gap convention means adjacent windows differ by ≥1 day, it never false-rejects
   a legitimate gap window. The **pre-drop overlap diagnostic** is section A of the artifact (run BEFORE the
   drop/trigger; returned ZERO rows on the spike — if it returns rows, existing data already overlaps and
   the trigger would block future edits of those parts; clean them first, David reviews any hits).
   **WRITER-COMPAT NOTE (verified):** SQL Server forbids `INSERT … OUTPUT inserted.col` (no `INTO`) on a
   triggered table. The rebuilt master Save is SAFE (it uses `INSERT …; SELECT SCOPE_IDENTITY()`, no OUTPUT);
   the only OUTPUT-without-INTO writer is the headless `test_manifestcost_overlap_guard.py` anchor INSERT,
   which pins the predicate independently and runs without the trigger present (it still passes, 10/0/0).
   Spike validation proved: gap window INSERT succeeds; overlapping + touching INSERT and overlap-UPDATE are
   rejected with the THROW; a row's own self-update succeeds; a part can hold 2 gap windows.
5. **Repoint the UI lookup.** `SELECT_ManifestCost` (verified live, line 2006) uses STRICT `>`/`<` bounds
   — it drops the first AND last effective day of every window (latent boundary bug) and is window-blind on
   the list branch. It is **SUPERSEDED, not ported.** Repoint the rebuilt master "price on date D" lookup
   to a Named Query that does `SELECT MO_PRICE FROM dbo.fn_ManifestCostAt(@assyCode, @prodDate)` so the UI
   and the reports agree on "in window" (inclusive `<=`/`>=`). The list/browse branches of
   `SELECT_ManifestCost` (the no-date and no-assy cases that just list rows) move to a plain
   `SELECT_ManifestCost`-shaped Named Query over the table (no window logic needed for a list). Per the
   Named-Query practice, name them to mirror the proc/table: `ManifestCost/SelectAll`,
   `ManifestCost/PriceAtDate` (the TVF wrapper).

**Why.** CROSS APPLY = inner-join semantics but window-correct; the TVF is the single-point-of-truth for
the window rule (IA Named-Query practice). The non-overlap trigger is the only DB mechanism that can see
sibling rows. Repointing the UI to the same TVF closes the legacy UI-vs-report disagreement on boundary days.

**Open-risk.** `test_report_procs_d6.py` deploys `_D6`-suffixed copies and diffs them against the legacy
procs — that proves the migration is faithful EXCEPT on the intended-divergence rows (boundary days,
multi-window parts, the dropped 6440 site filter). At cutover those diffs become the production behavior;
**David must review the proc-parity diff once more against live data** and accept the intended divergences
(same discipline as the ledger parity classes).

✅ **RESOLVED (David 2026-06-19) — CROSS vs OUTER APPLY for priceless lines: KEEP CROSS APPLY (all 4 procs)
+ ADD a pre-invoice diagnostic.** CROSS APPLY is faithful to the legacy inner JOIN and NEVER emits a $0
line into the EDI 810/856 to Toyota (a priceless line is dropped, not billed at zero). To prevent a dropped
line becoming a SILENT under-bill, a pre-invoice diagnostic surfaces the lines CROSS APPLY would drop so
gaps are caught BEFORE billing. **Artifact: `docs/analysis/reporting/priceless-lines-diagnostic.sql`**
(OUTER APPLY `fn_ManifestCostAt` + `WHERE m.IN_MANIFEST_COST_ID IS NULL`, scoped to the EDI810 @EIN=0
filter `a.VC_ASN_STATUS='A' AND d.IN_INV_ID IS NULL`; also wrapped as proc
`REPORT_EDI810_PricelessLines`). Validated on the spike: 0 on current data; fabricating one out-of-window
unbilled line made it return exactly that line.

---

## Carry 4 — Opening-balance backfill placement

**Decision: confirmed sound. It sits AFTER the trigger drop and BEFORE the writer-flip goes live — in a
single quiesced maintenance window with NO writes in flight. The genesis guard makes ordering-vs-triggers
moot for correctness, but the quiesce makes it moot for races.**

**Design / critique.** `SEED_AllOpeningBalances` records each part's CURRENT `IN_QTY` as one
`OPENING_BALANCE` ledger row **without bumping `IN_QTY`** (the column already holds the cutover value).
After it, `SUM(ledger) == IN_QTY` by construction; every forward `POST_StockMovement` keeps the invariant
(it bumps ledger + `IN_QTY` in lockstep). The genesis guard in `SEED_OpeningBalance` THROWs if any forward
row exists first (recording an already-bumped `IN_QTY` as the opening would double-count). The set-based
`SEED_AllOpeningBalances` only seeds parts with NO ledger row yet — idempotent and re-runnable.

The critical ordering question is: **does the backfill read `IN_QTY` before or after the last legacy
trigger fire?** The seed captures whatever `IN_QTY` is at the instant it runs. For `SUM(ledger)==IN_QTY`
to hold going forward, the seed must read the **final, settled** `IN_QTY` and **nothing may move stock
between the seed read and the first seam post.** Two ways to guarantee that:

- **(chosen)** Run the seed in a **quiesced window**: the Delphi app is down (legacy triggers can't fire
  because no source writes happen), and the seams are not yet live. Drop the triggers, run the seed, then
  bring the seams up. With no writer active during the seed, `IN_QTY` is frozen and the snapshot is exact.
- The genesis guard is the **safety net**, not the primary mechanism: if a seam somehow posted before the
  seed (misordered cutover), the guard THROWs rather than silently double-counting.

**So the precise placement is: drop triggers → (DB now quiet, no writer owns `IN_QTY`) → run
`seedAllOpeningBalances` → bring seams live (writer-flip).** Putting the seed AFTER the trigger drop is
important: if you seeded while triggers were still live, a concurrent legacy trigger fire between the seed
read and the trigger drop would change `IN_QTY` and break the snapshot. With triggers already gone and
seams not yet up, `IN_QTY` cannot change. (The runbook §4's "before any forward post" is satisfied; this
pins it tighter: also *after* the last legacy mutation.)

**Why this beats the §9 per-source-row replay backfill.** The from-zero replay is impossible
(`DELETE_AutoPurge` aged out the receiving IN-moves), so a replay would reconstruct a WRONG opening
(missing the purged receipts). Opening-balance-from-current-`IN_QTY` is exact by construction and is the
sign-off David chose. The replay survives only as the **parity harness derivation** (read-only).

**Open-risk.** The backfill's correctness depends entirely on the quiesce (no writer touching `IN_QTY`
during the seed). The runbook must make "app down + triggers dropped + seams not yet enabled" an explicit
gate. The genesis guard catches a sequencing mistake but a *concurrent legacy write during the seed* is
NOT caught (it would change `IN_QTY` under the snapshot without a forward ledger row) — hence the quiesce
is mandatory, not optional. **No David decision needed**; this is an execution-discipline item.

---

## Carry 5 — Retire the 12 legacy qty-triggers

**Decision: DROP all 12 (named below) in the same quiesced window, BEFORE the opening-balance seed and
BEFORE the seams go live. One transaction-free DDL batch. This is the single point where double-count or
gap is possible, so it is fenced by the quiesce.**

**The actual triggers (grepped from live `CreateInventory.sql`):**

| Source table | Triggers (the 12 qty movers) |
|---|---|
| `INV_OPEN_ORDER_INF` | `INSERT_RecConfStatPartsStockMstQTY`, `UPDATE_RecConfStatPartsStockMstQTY`, `DELETE_RecConfStatPartsStockMstQTY` |
| `INV_REJECT_INF` | `INSERT_RejectParts`, `UPDATE_RejectParts`, `DELETE_RejectParts` |
| `INV_STOCKTAKING_INF` | `INSERT_Stocktaking`, `UPDATE_Stocktaking`, `DELETE_Stocktaking` |
| `INV_PART_SHIPPING_INF` | `InsertPartShipping`, `UpdatePartShipping`, `DeletePartShipping` |

Plus the **13th, the header cascade**: `DeleteShipDate` on `INV_SHIPPING_INF` (fires the part-shipping
qty restore) — it is re-homed into `shipping.deleteShipmentHeader` / `postHeaderDelete`, so **drop it too**
(otherwise a header delete double-restores). The runbook says "12"; the executable list is **13 DROPs**.

**NOT dropped:** `UPDATE_PartNumber` and `DELETE_PartNumber` on `INV_PARTS_STOCK_MST`. ✅ **RESOLVED
(David decided): `UPDATE_PartNumber` STAYS — KEPT through and after cutover.** Its purpose is to AUDIT every
state change to an `INV_PARTS_STOCK_MST` row: it writes an `INV_PART_QTY_INF` row + an
`INV_PARTS_STOCK_MST_HIST` snapshot whenever the row changes — including when a seam's
`IN_QTY = IN_QTY + @delta` fires it. This is **CORRECT, intended, complementary** behavior: the LEDGER
records the *movement*; the AUDIT trigger records the *row-state change*. They are not duplicative, and any
incidental duplicate audit row is harmless and ignored. `UPDATE_PartNumber` does **not** move stock (it
writes audit tables, never re-mutates `INV_PARTS_STOCK_MST.IN_QTY`), so it cannot double-count. **The prior
"verify the body is audit-only" open item is CLOSED by decision — no body re-verification is a cutover
gate.** The `DISABLE TRIGGER UPDATE_PartNumber` in `test_ledger_opening_balance.py` / `test_seam_driver.py`
is a TEST-ISOLATION measure (keeps audit noise out of the ledger assertions), **NOT a cutover instruction
— do not translate it into a DISABLE/DROP step.**

**The two non-trigger direct `IN_QTY` writers (the 4th/5th-writer hunt — runbook §5b/§10):** beyond the 13
triggers, exactly two procs do a direct `UPDATE INV_PARTS_STOCK_MST SET IN_QTY=…`:
`UPDATE_PartsStockInfo` (absolute `=@QTY`, line 5691 — live caller `DataModule.pas:1482`, latent clobber)
and `UPDATE_PartsStockInfoCount` (additive `-@QTY`, line 4076 — **dead, zero callers**). Neither is a true
LIVE-&-UNWIRED 5th producer: the dead one is retired, the clobber one has its `IN_QTY=@QTY` clause neutered
(it is a master-data edit, not a stock move). See runbook **carry 10** for the actions and the David flags.
(`UPDATE_PartsStockRenban`/`UPDATE_OrderRenbanQty`/`UPDATE_OrderQty`/`UPDATE_Shippingdetail`/
`INSERT_ShippingPartInfo` were checked and do NOT write `INV_PARTS_STOCK_MST.IN_QTY` — they write
`IN_RENBAN_COUNT` or the IN_QTY column of OTHER tables whose own triggers are already producers.)

**Sequence reasoning (why drop BEFORE seed BEFORE seams):**
- If seams go live **while triggers are still on** → IN_QTY double-counts (trigger + seam both move it).
  This is exactly the "CUTOVER write path presumes triggers are gone" warning in every seam header.
- If triggers are dropped but **the seed runs while the app is still up** → a final legacy write would
  have nothing moving `IN_QTY` (gap) and the seed snapshot would miss it. Hence app-down first.
- Drop order within the batch is irrelevant (DDL, app down). Do them as one idempotent script
  (`IF OBJECT_ID(...,'TR') IS NOT NULL DROP TRIGGER ...` ×13) so a re-run is safe.

**Open-risk.** The single irreversible-feeling step. Mitigated by: (a) the drops are scripted and the
trigger bodies are preserved in `CreateInventory.sql`, so a rollback re-creates them; (b) the rollback
point (below) is "before the seams are enabled" — if the writer-flip misbehaves, re-create the 13 triggers
and the legacy path is restored (the opening-balance ledger rows are harmless append-only data). **The
`UPDATE_PartNumber` disposition is now RESOLVED (KEPT, David decided) — no longer an open item.**

---

## Carry 6 — Per-edit version-stamp contracts

**Decision: confirmed sufficient and unambiguous as built. One tightening note on the shipping
header-delete key.**

**Verification against the code:**
- **Shipping amend re-stamps `VC_ADD`:** `amendPartShipping` writes `VC_ADD = <fresh stamp>` on every
  edit, and `_ver(row)` reads `VC_ADD` as the version token. `INV_PART_SHIPPING_INF` has no
  `VC_LAST_UPDATE`, so `VC_ADD` is correctly the amend token. The amend key is
  `SHIPPING:psh=<id>:upd:to=<effect>:v=<VC_ADD>` — collision-safe: two amends to the same target effect
  in the same hundredth-second would share `(stamp,target)`, but a same-target second amend has delta 0
  and posts nothing, so no collision can corrupt the balance. **Sufficient.**
- **Stocktaking non-NULL `VC_LAST_UPDATE`:** the stocktaking edit path writes a fresh 16-char
  `VC_LAST_UPDATE` per edit (fixes legacy Bug2 NULL), and the reject/stocktaking amend keys use
  `:upd:to=<effect>:v=<VC_LAST_UPDATE>`. Non-NULL is guaranteed by the `_STAMP_SQL` recipe. **Sufficient.**
- **The `:upd:to=:v=` key shape:** used uniformly across receiving/shipping/reject/stocktaking amends.
  The `:to=` (target effect) + `:v=` (stamp) pair is collision-safe by the delta-0 argument above.
  Insert/delete keys carry NO version token, which is correct because the surrogate id is IDENTITY
  (inserted/deleted once per lifetime → replay reuses the same key → idempotent). **Sufficient.**

**One tightening note (open-risk, low).** The shipping header-delete (`postHeaderDelete`) emits one
`SHIPPING:psh=<id>:del` per row — those are insert/delete-style keys (no version), correct because each
`IN_PART_SHIPPING_ID` is deleted once. But a header delete followed by a **re-create of the same
production date** would mint NEW `IN_PART_SHIPPING_ID`s (IDENTITY), so new keys — no collision. Confirmed
safe. **No David decision needed.** The contracts are sufficient and unambiguous.

---

## Carry 7 — Headless Jython driver coverage (persistent-sqlcmd-session extension)

**Decision: extend `jython_shim.py` with a persistent `sqlcmd` subprocess that holds an open session so
`beginTransaction`/`commit`/`rollback` span statements — required by Order's `commitOrders`. Shipping +
receiving need NO shim change (mechanical, same autocommit shim).**

**Design.** The shim today spawns a fresh `sqlcmd -Q` per statement (autocommit) and stubs
`beginTransaction` as a no-op. Order's `commitOrders` opens a gateway tran, runs N `EXEC INSERT_OpenOrder`
+ one `EXEC RESERVE_RenbanCount`/`UPDATE_PartsStockRenban` on that `tx`, and commits — a no-op tx shim
would run each in its own autocommit and a mid-loop failure would leave partial orders (the exact thing
the tran protects). Extension:

- Add a `_Session` class that launches **one** long-lived `sqlcmd` process in interactive mode
  (`sqlcmd -S ... -d Inventory` with stdin kept open), writing batches terminated by `GO` and reading
  delimited output. The session is created lazily on the first `beginTransaction`.
- `beginTransaction(db)` → send `SET IMPLICIT_TRANSACTIONS OFF; BEGIN TRAN;` on the session, return a
  token bound to that session. `runPrepUpdate(..., tx=token)` and `runPrepQuery(..., tx=token)` route to
  the session's stdin instead of a fresh `-Q`. `commitTransaction` → `COMMIT TRAN;`; `rollbackTransaction`
  → `ROLLBACK TRAN;`; `closeTransaction` → leave the session open for reuse (or close it).
- Output parsing reuses `_parse_rows`; the session must emit a sentinel after each batch (e.g.
  `PRINT '<<<END>>>'`) so the reader knows where one statement's output stops — the single biggest
  mechanical risk (framing the streamed output). Keep `getKey` working by appending
  `SELECT CAST(SCOPE_IDENTITY() AS int)` in the same batch on the session.
- **Param binding stays inline** (`_bind`/`_esc`) — the shim already has no real `?` binding; that is
  unchanged and acceptable for a test harness.

A new `test_order_seam_driver.py` then drives the REAL `commitOrders` (lot-sized → N records + renban
span; palletized-grouped → 1 blank-renban record; the `RESERVE_RenbanCount` reserve) end-to-end on the
session, asserting the orders + the advanced counter + rollback-on-failure (force a mid-loop error and
assert no partial orders, no counter advance).

**Shipping + receiving** load under the EXISTING shim exactly like stocktaking/reject in
`test_seam_driver.py` (`load_wrapper(... extra_globals={"stockLedger": sl})`, disable the relevant
triggers, drive insert/amend/delete, assert `IN_QTY` + ledger). Header-delete for shipping needs a
multi-row fixture (insert 2 part-shipping rows under one `IN_SHIPPING_ID`, call `deleteShipmentHeader`,
assert both restored — the F3 multi-row case). No shim change.

**Why.** The persistent session is the minimal way to give the shim cross-statement transaction semantics
without a gateway. It is test infra only; it does not touch the production seams. The remaining true gap
(Jython-2.7-vs-CPython runtime quirks, in-gateway Perspective runtime) stays a documented gap — the
wrappers are written 2.7/CPython-portable and reviewer-checked, and Perspective E2E is trial-gated.

**Open-risk.** Streamed-output framing (the sentinel parsing) is fiddly and the one place this can flake.
**No David decision needed** — it is a test-harness build task. Note: this is test coverage, not a cutover
gate per se, but it should pass before the writer-flip so the seams are exercised end-to-end.

---

## Carry 8 — Postgres-phase `# IG83-TODO` triage (cutover vs later)

**Decision: split the `IG83-TODO` set into MUST-AT-CUTOVER vs GENUINELY-LATER. Most are later; three are
cutover-adjacent and must at least be *decided* at cutover even if implemented as-is.**

**MUST be addressed at cutover (the writer-flip happens on SQL Server, pre-Postgres):**
- *(none are blocking)* — the entire ledger + seams + D6 run on SQL Server under 8.1.52 today. The
  cutover does NOT require Postgres. This is the key triage result: **cutover is a SQL-Server-only event;
  no `IG83-TODO` blocks it.**

**Cutover-adjacent — decide/verify at cutover, implement as-is (string types kept):**
- `fn_ManifestCostAt` string-date compare: correct ONLY because dates are zero-padded `yyyymmdd`
  (lexicographic == chronological). **Verify the invariant holds in live data at cutover** (no malformed
  dates); keep strings. Deferred conversion to typed `date` is Postgres-phase.
- `site_id` single-site default (=1): cutover is single-site (parallel-run). The `site_id` column exists
  on `INV_STOCK_LEDGER` with default 1; per-site scoping + NOT NULL FKs land at Postgres. **Verify no
  multi-site assumption leaks into a cutover seam** (the renban reserve, carry 2, is the one to check —
  confirmed per-part-row-correct for single-site).

**GENUINELY LATER (Postgres phase, do NOT touch at cutover):**
- `TS_POSTED`/`VC_ADD`/`VC_LAST_UPDATE` 16-char strings → `datetime2`.
- Real PKs/FKs on the source tables (`INV_OPEN_ORDER_INF`, `INV_REJECT_INF`, etc.).
- The manifest-cost computed-view projection option.
- The materialized-vs-computed `IN_QTY` projection (ledger design Q1) — explicitly Postgres-phase.
- `site_id` NOT NULL FKs + full per-site scoping (D1).

**Why.** Cutover is the writer-flip on the existing SQL Server; the Postgres migration is a separate later
program. Conflating them would block a deliverable cutover on a large DB-platform change. The only
cutover obligations from this set are *verifications* (the string-date invariant, no multi-site leak), not
implementations.

**Open-risk.** None blocking. **No David decision needed** beyond noting that cutover ships on SQL Server.

---

## Carry 9 — Spec-hygiene residue

**Decision: one-pass sweep, LOW priority, NOT a cutover gate.** Fix-on-touch is the standing rule; the
live-truth gate already catches stale claims at build. Known residue: `shipping.md` §2 "No declared FKs"
(live DB HAS `FK_INV_PART_SHIPPING_INF_INV_SHIPPING_INF`); manifest unique-index framing in pre-D9 prose.
Schedule the sweep post-cutover. **No David decision needed.**

---

# END-TO-END CUTOVER SEQUENCE

One ordered runbook. Phases A–B are non-destructive and can be staged ahead of the maintenance window;
phase C is the destructive flip and MUST run in a quiesced window (app down). Each step notes its safety
reasoning and the rollback point.

**Pre-req gate (before scheduling):** all seam end-to-end tests pass under the shim — stocktaking, reject
(done), shipping, receiving (mechanical extension), Order (persistent-session extension, carry 7). The
ledger parity harness (`test_stock_ledger_parity.py`) is GREEN with every EXPECTED-DIVERGENT class tagged
(D8(3), D12#3, F3, F5) and reject/shipping deletes EXPECTED-ZERO. The D6 proc-parity diff
(`test_report_procs_d6.py`) reviewed and intended divergences accepted by David.

### Phase A — non-destructive DB objects (can run days ahead; additive only)
1. **Apply `spike-manifest-cost-lookup.sql`** → `fn_ManifestCostAt`. *Additive; the D6 procs depend on it.
   Rollback: drop the function.*
2. **Apply `spike-stock-ledger-table.sql`** → `INV_STOCK_LEDGER` + UNIQUE `(IN_PART_ID,
   VC_SOURCE_EVENT)` + covering index. *Additive new table; nothing reads it yet. Rollback: drop table.*
3. **Apply `spike-post-stockmovement-proc.sql`** → `POST_StockMovement` + `PROC_RebuildStockBalance`.
   *Additive procs; uncalled until phase C. Rollback: drop procs.*
4. **Apply `spike-seed-opening-balance.sql`** → `SEED_OpeningBalance` + `SEED_AllOpeningBalances`.
   *Additive; uncalled until phase C. Rollback: drop procs.*
5. **Deploy the carry-1 `WRITE_*` write-and-post procs** and the carry-2 `RESERVE_RenbanCount`.
   *Additive; the seams that call them are not yet the live writer. Rollback: drop procs.*
5b. **Neuter the non-trigger `IN_QTY` writers (carry 10) — ✅ RESOLVED (David ACCEPTED 2026-06-19).
   Artifact: `docs/analysis/master-data/spike-partsstockinfo-drop-qty-clause.sql`.** `ALTER PROCEDURE
   UPDATE_PartsStockInfo` to DROP the `IN_QTY=@QTY` clause (master-data edit no longer touches on-hand; the
   `@QTY` param is KEPT in the signature but now unused, so the rebuilt Save's positional 30-param call is
   unchanged — accepted-and-ignored, no Ignition-side change). Retire `UPDATE_PartsStockInfoCount` (dead, no
   caller — `DROP PROCEDURE`, guarded). David confirmed: on-hand is NEVER editable on the rebuilt Parts Stock
   master; all qty change goes through a ledger transaction (the seam/ledger is the sole IN_QTY owner). *Safe
   pre-window: during parallel run the rebuilt master loads and re-writes the same `@QTY` (no-op-in-effect);
   after the edit it simply stops moving on-hand. Rollback: re-create both procs from `CreateInventory.sql`.
   Spike-validated: baseline legacy proc moved IN_QTY (7382→17381); after the ALTER, IN_QTY stays 7382 while
   other columns (comments) still update; spike restored to as-found.*
6. **Deploy the seam Project Libraries** (stockLedger, receiving, shipping, reject, stocktaking, order)
   to the gateway, **but do NOT wire any screen to them as the live writer yet.** *Code present, not on
   the hot path. Rollback: revert the project.*

   > Safety: everything in Phase A is additive. The legacy triggers still own `IN_QTY`; the ledger is an
   > empty shadow. The system runs exactly as before. **Full rollback = drop the new objects.**

### Phase B — D6 report cutover + non-overlap backstop (low-risk, reversible)
7. **Run the pre-drop overlap diagnostic** (section A of
   `docs/analysis/master-data/spike-manifestcost-nooverlap-trigger.sql`, equivalent to the query at the
   bottom of `spike-manifest-cost-lookup.sql`). Expect ZERO rows. If any, David cleans the overlapping
   windows before step 9. *Read-only check (returned 0 on the spike).*
8. **Apply `spike-report-procs-d6.sql`** → replaces the 4 window-blind report procs with the
   `CROSS APPLY fn_ManifestCostAt` versions (+ drops the EDI856 6440 site bug). **CROSS APPLY kept on all 4
   (David 2026-06-19) — never emits a $0 line to Toyota.** Also **apply
   `docs/analysis/reporting/priceless-lines-diagnostic.sql`** (the pre-invoice safety net,
   `REPORT_EDI810_PricelessLines`) and wire it as a pre-billing check (expect 0 before each EDI810/856 run).
   *Rollback: re-create the legacy procs from `CreateInventory.sql`; drop the diagnostic proc.*
9. **Conditionally DROP `IX_INV_MANIFEST_COST_MST`** (it is a UNIQUE *constraint*, not an index — use
   `ALTER TABLE … DROP CONSTRAINT`; the artifact's section B handles both forms) and **create
   `TRG_ManifestCost_NoOverlap`** (carry 3) — both in
   `docs/analysis/master-data/spike-manifestcost-nooverlap-trigger.sql` (sections B + C). *The trigger only
   blocks future overlapping writes; existing data already passed step 7. Rollback: drop the trigger;
   re-create the constraint (`ALTER TABLE … ADD CONSTRAINT IX_INV_MANIFEST_COST_MST UNIQUE NONCLUSTERED
   (VC_ASSY_MANIFEST_NUMBER)`).*
10. **Repoint the ManifestCost UI lookup** Named Queries to `fn_ManifestCostAt` (`PriceAtDate`) + a plain
    list NQ (`SelectAll`). *UI-only; rollback: re-point to `SELECT_ManifestCost`.*

    > Safety: Phase B touches reporting + the manifest master only — orthogonal to the stock-ledger flip.
    > It can ship independently and be rolled back independently. **Rollback point: re-create the 4 legacy
    > report procs + the index, drop the no-overlap trigger.**

### Phase C — the writer-flip (DESTRUCTIVE; quiesced window, app DOWN)
11. **Bring the Delphi app down.** No source writes can occur → no legacy trigger can fire → `IN_QTY` is
    frozen. *This is the quiesce that makes steps 12–14 race-free (carry 4/5).*
12. **DROP the 13 qty triggers** (the 12 movers + `DeleteShipDate`) via the idempotent
    `IF OBJECT_ID(...,'TR') IS NOT NULL DROP TRIGGER` script (carry 5). *After this, NOTHING owns `IN_QTY`
    — but the app is down, so no write is dropped. **KEEP `UPDATE_PartNumber`/`DELETE_PartNumber`** (audit
    triggers, RESOLVED — not movers, do NOT drop, do NOT disable). Do NOT confuse the test-suite's DISABLE
    with a cutover step.*
13. **Run `stockLedger.seedAllOpeningBalances()`** → one `OPENING_BALANCE` ledger row per part recording
    the frozen `IN_QTY`, NO bump. After this `SUM(ledger) == IN_QTY` for every part. *Genesis guard
    THROWs if any forward row pre-exists (sequencing safety). Idempotent — safe to re-run. Must run AFTER
    the drop (no trigger can change `IN_QTY` mid-seed) and BEFORE any seam post (carry 4).*
14. **Validate the invariant:** run the parity/health query `SELECT p.IN_PART_ID FROM INV_PARTS_STOCK_MST
    p WHERE p.IN_QTY <> (SELECT ISNULL(SUM(IN_QTY_CHANGE),0) FROM INV_STOCK_LEDGER l WHERE
    l.IN_PART_ID=p.IN_PART_ID)`. Expect ZERO rows. *Go/No-Go gate. Non-zero ⇒ STOP and roll back.*
15. **Enable the seams as the live writer** (wire the RecConfStat/Shipping/Reject/Stocktaking/Order
    screens to the `WRITE_*` procs / seam libraries). *Now the seams own `IN_QTY` via the ledger; the
    triggers are gone, so no double-count.*
16. **Bring the app/gateway up** for the rebuilt write paths. Smoke-test one of each: a receiving
    confirm, a reject, a stocktaking adjust, a shipment, an order create — assert each moves `IN_QTY` by
    the right delta AND appends exactly one ledger row, and re-run step 14's invariant query (still zero).

    > **ROLLBACK POINT for Phase C:** up to and including step 14, rollback = **re-create the 13 triggers
    > from `CreateInventory.sql` and bring the Delphi app back up.** The `OPENING_BALANCE` ledger rows are
    > harmless append-only data (drop them if desired; they don't affect `IN_QTY`). Once step 15 enables
    > the seams AND real forward posts have happened (step 16), rollback means re-creating the triggers
    > AND reconciling any forward `IN_QTY` deltas the seams applied while live — so **step 14's green
    > invariant is the point of no easy return; treat it as the commit gate.** Schedule the window so 11→16
    > complete in one sitting.

### Phase D — post-cutover (housekeeping, non-blocking)
17. **Schedule the spec-hygiene sweep** (carry 9). *Low priority.*
18. **Park the Postgres-phase `IG83-TODO` set** (carry 8) as the next program — string→datetime2,
    real FKs, `site_id` scoping, computed-view projection. *Not a cutover item.*

---

## Items needing David's decision (flagged) — ✅ ALL RESOLVED (2026-06-19)
1. ✅ **RESOLVED — `UPDATE_PartNumber` disposition** (carry 5): David decided it STAYS (audit trigger,
   complementary to the ledger, does not move stock). No longer an open item; the seam-test DISABLE is test
   isolation, not a cutover step.
2. ✅ **RESOLVED (David 2026-06-19) — CROSS vs OUTER APPLY for priceless report lines** (carry 3): **KEEP
   CROSS APPLY on all 4 procs** (faithful to legacy; never emits a $0 line to Toyota) **+ ADD a pre-invoice
   diagnostic** that surfaces the priceless lines CROSS APPLY drops, so gaps are caught before billing.
   Artifact: `docs/analysis/reporting/priceless-lines-diagnostic.sql` (validated).
3. **Final acceptance of the intended-divergence rows** at cutover — note these are now FORWARD-behavior
   divergences only. Per BLOCKER 2's resolution the opening balance copies legacy `IN_QTY` verbatim, so the
   D8(3)/D12#3/F3/F5 classes are FROZEN into the seed (not corrected). Going forward the seams diverge from
   the old triggers (intended); the D6 proc diffs (boundary days, multi-window parts, dropped 6440 filter)
   become production behavior — review the proc-parity diff once against live data and sign off. *(Standing
   sign-off review, not an open design decision.)*
4. ✅ **RESOLVED (David 2026-06-19) — the `IX_INV_MANIFEST_COST_MST` drop** (carry 3 step 9): **DROP +
   REPLACE.** Drop at cutover (it is a UNIQUE *constraint*, not an index — verified) so the rebuilt master
   can hold multiple windows/part; integrity moves to the app `checkWindowOverlap` guard + the new
   `TRG_ManifestCost_NoOverlap`. Artifact: `docs/analysis/master-data/spike-manifestcost-nooverlap-trigger.sql`
   (validated; pre-drop diagnostic returned 0).
5. ✅ **RESOLVED (David ACCEPTED 2026-06-19) — `UPDATE_PartsStockInfo` qty-leg neuter + on-hand-override
   policy** (carry 10): drop the `IN_QTY=@QTY` clause; on-hand is NEVER editable on the PartsStock master
   (all qty change via a ledger transaction). `UPDATE_PartsStockInfoCount` is dead → retired; not a 5th
   producer. Artifact: `docs/analysis/master-data/spike-partsstockinfo-drop-qty-clause.sql` (validated).
6. ✅ **RESOLVED/CLOSED (David 2026-06-19) — Order-commit ledger "gap"** (carry 11): **NO WORK — not a gap.**
   The Order worksheet only emits NEW orders for the calculated FUTURE FRS date; an already-shipped/arrived
   order cannot exist in the system on that future FRS date, so the Order-commit path needs NO ledger post.
   The effect-0-at-creation assumption holds by construction. Marked closed; no conditional `stockLedger.post()`.
