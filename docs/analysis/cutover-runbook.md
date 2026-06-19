# Cutover Runbook — deferred carries from the stock-ledger / D6 / Order build

**Status:** 🟡 living checklist — the single home for items deliberately DEFERRED to cutover (retro R13,
2026-06-18). Each is flagged in its PR/code; this collects them so none is lost. **Before cutover: an
architect pass (adversarial-architect-reviewer / ignition-architect) over this whole list.**

The build (PRs #8–#14, all on master) shadows the legacy during parallel-run; cutover flips reads to
the ledger + retires the legacy triggers. These are the things that must be resolved at that flip.

## 1. Atomicity — source-write ↔ ledger-post  (all 4 producer seams + Order)
The write-then-post wrappers (`receiving/shipping/reject/stocktaking.*` + `order.commitOrders`) write the
source row and post to the ledger as **separate statements** — not one transaction (except Order, which
wraps its own). A post failure after the source write commits orphans the row from the ledger.
**Fix at cutover:** combine source-write + `POST_StockMovement` into one transactional unit — either a
single stored proc per op, or thread a gateway `txId` through `stockLedger.post`. **Decide the pattern
(architect).** Moot during parallel-run (the ledger is a shadow).

## 2. D11#7 — renban-counter race  (Order commit)
`order.commitOrders` reads `IN_RENBAN_COUNT`, computes renbans, writes the advanced count — a read-then-
write that races under concurrent order-creates (two specialists → duplicate renbans). Mirrors the legacy.
**Fix at cutover:** atomic counter allocation (an `UPDATE … OUTPUT` that reserves N values, or a sequence).

## 3. D6 report-proc cutover apply  — ✅ DECISIONS RESOLVED (David 2026-06-19)
`docs/analysis/reporting/spike-report-procs-d6.sql` is the CUTOVER artifact — it REPLACES the legacy
window-blind `REPORT_INVOICESSummary / MonthlyINVOICESSummary / EDI810 / EDI856`. Not applied to the spike
(legacy kept as the parity baseline). **At cutover:** apply it; confirm the manifest unique-index is
dropped and add the **DB-level non-overlapping-window trigger** as the backstop to the
app guard (§4). `SELECT_ManifestCost` (UI lookup) is SUPERSEDED by `fn_ManifestCostAt` — repoint the UI.

> ✅ **RESOLVED (David 2026-06-19) — priceless lines: KEEP CROSS APPLY (all 4 procs) + ADD a pre-invoice
> diagnostic.** CROSS APPLY is faithful to the legacy inner JOIN and never emits a $0 line into the EDI
> 810/856 to Toyota (a priceless line is DROPPED, not billed at zero). To stop a dropped line becoming a
> silent under-bill, run a pre-invoice diagnostic that surfaces exactly the lines CROSS APPLY would drop.
> **Artifact: `docs/analysis/reporting/priceless-lines-diagnostic.sql`** (`REPORT_EDI810_PricelessLines`;
> OUTER APPLY `fn_ManifestCostAt` + `WHERE m.IN_MANIFEST_COST_ID IS NULL`, scoped to EDI810's
> `VC_ASN_STATUS='A' AND IN_INV_ID IS NULL`). Spike-validated: 0 on current data; a fabricated out-of-window
> unbilled line returned exactly that line; an in-window control returned 0. **Wire it as a pre-billing
> check (expect 0 before each EDI810/856 run).**

> ✅ **RESOLVED (David 2026-06-19) — manifest index: DROP + REPLACE.** Drop `IX_INV_MANIFEST_COST_MST`
> (verified: a UNIQUE *constraint* on `VC_ASSY_MANIFEST_NUMBER`, varchar 2 — the table is a HEAP, no PK —
> so it drops via `ALTER TABLE … DROP CONSTRAINT`, not `DROP INDEX`) so the rebuilt master can hold multiple
> windows per part; integrity moves to the app `checkWindowOverlap` guard + a new DB trigger
> `TRG_ManifestCost_NoOverlap`. A pre-drop overlap diagnostic flags existing overlaps to clean first.
> **Artifact: `docs/analysis/master-data/spike-manifestcost-nooverlap-trigger.sql`** (section A diagnostic,
> section B conditional drop, section C trigger). Spike-validated: pre-drop diagnostic = 0; gap-window INSERT
> succeeds; overlapping + touching-boundary INSERTs rejected with the THROW; self-update succeeds. WRITER
> note: a direct writer using `OUTPUT inserted.*` WITHOUT `INTO` is forbidden on a triggered table (verified
> — the existing `test_manifestcost_overlap_guard.py` anchor INSERT hits this; it runs without the trigger
> present and the rebuilt master Save uses `SCOPE_IDENTITY`, so it is safe).

## 4. Opening-balance backfill  (the parity closure)
Run `stockLedger.seedAllOpeningBalances()` (set-based `SEED_AllOpeningBalances`) ONCE at cutover, **before
any forward post** (the genesis guard THROWs otherwise). After it, `SUM(ledger) == IN_QTY` for every part;
forward posts keep it. The from-zero full reconstruction is impossible (receiving history purged) — this
opening balance + forward parity IS the sign-off. (A real pre-purge dump, if one ever surfaces, would let
`test_stock_ledger_parity.py`'s skipped from-zero check run — but it's not required.)

> ✅ **RESOLVED (BLOCKER 2 — backfill contradiction; David's decided approach).** The shipped
> `SEED_AllOpeningBalances` (copy CURRENT legacy `IN_QTY` verbatim → `SUM(ledger)==IN_QTY` tautologically)
> **is the cutover backfill.** This is David's "start clean, go forward" call: from-zero replay is
> impossible (receiving IN-moves were purged by `DELETE_AutoPurge`), so the opening balance can only be the
> current materialized value. It is **NOT to be redesigned.**
>
> **The stale §9 of `IGNITION-stock-ledger-design.md`** (which describes "derive one movement per live
> source row" → a *corrected* balance, plus a "reconcile the D8(3)/D12#3/F3/F5 corrections at cutover"
> paragraph) **is superseded for the cutover writer.** Annotate §9 in place: the per-source-row derivation
> survives ONLY as the read-only parity-harness derivation, never as the cutover seed. (Architecture doc
> §"Scope correction" para 2 already says this; this is the single source of truth.)
>
> **Accepted tradeoff (document honestly — fixture-fidelity discipline, `feedback-parity-fixture-fidelity`):**
> - Legacy `IN_QTY` bugs of the **D8(3) (arrival-reversal overstate) / D12#3 (yard under-count) /
>   F3 (multi-row under-count) / F5 (part-change no-op)** classes are **FROZEN INTO the opening balance** —
>   the cutover does NOT correct historical on-hand. Only **forward** posts (via the seams) get the
>   corrected behavior.
> - The `test_ledger_opening_balance.py:138` "0 drift after backfill" check proves the seed **MECHANISM**
>   (every part gets exactly one OPENING_BALANCE row equal to its `IN_QTY`, idempotently), **NOT balance
>   correctness.** It is vacuously green by construction and must never be cited as a correctness/parity
>   sign-off. The four EXPECTED-DIVERGENT classes are **retired from cutover scope** as balance corrections;
>   they remain meaningful only as forward-behavior parity (the seams diverge from the old triggers going
>   forward, which is intended).
> - No delta-reconciliation step is needed (none is performed — the copy is exact by construction). The §9
>   "reconcile the delta explicitly" language is dead for the cutover.

> ⚠️ **Minor (folded note — IN_BALANCE_AFTER race, SHOULD-FIX 5).** `POST_StockMovement` computes
> `IN_BALANCE_AFTER` from a pre-post `SELECT IN_QTY ... + @delta` (a separate read before the additive
> UPDATE). Under two concurrent posts on the same part this **diagnostic column** can record a wrong
> running balance for one row (both read the same pre-value). **`IN_QTY` itself stays correct** — the
> additive `IN_QTY = IN_QTY + @delta` UPDATEs serialize on the row lock and are commutative. Only the
> snapshot column is racy. Disposition: **accept as-is** (it is an optional diagnostic, not the balance);
> do NOT verify "monotonic replay" off `IN_BALANCE_AFTER` under concurrency. The Phase-C quiesce window
> (app down, single writer) means there is no concurrency during the seed/flip; the only exposure is
> normal forward operation, which is non-load-bearing. If ever needed, recompute it via
> `OUTPUT inserted.IN_QTY` inside the additive UPDATE (`# IG83-TODO`, not a cutover gate).

## 5. Retire the 12 (really 13) legacy qty-triggers
At cutover, DROP the receiving/shipping/reject/stocktaking qty-triggers — the producer seams replace them.
The seams are CUTOVER paths that presume the triggers are gone (else IN_QTY double-counts). Sequence the
drop with the writer-flip.

> ✅ **RESOLVED (BLOCKER 1 + count — David decided `UPDATE_PartNumber` STAYS).** The executable drop list
> is **13 triggers**, NOT 12 (architecture-doc carry 5 enumerates them): the 12 qty-movers
> (`INSERT/UPDATE/DELETE_RecConfStatPartsStockMstQTY`, `INSERT/UPDATE/DELETE_RejectParts`,
> `INSERT/UPDATE/DELETE_Stocktaking`, `InsertPartShipping`/`UpdatePartShipping`/`DeletePartShipping`)
> PLUS the header cascade `DeleteShipDate` on `INV_SHIPPING_INF`.
>
> **`UPDATE_PartNumber` (and `DELETE_PartNumber`) on `INV_PARTS_STOCK_MST` are NOT in the drop list — they
> are KEPT through and after cutover.** David's decision: `UPDATE_PartNumber`'s purpose is to AUDIT every
> state change to an `INV_PARTS_STOCK_MST` row — it writes an `INV_PART_QTY_INF` row + an
> `INV_PARTS_STOCK_MST_HIST` snapshot whenever the row changes (incl. when a seam's
> `IN_QTY = IN_QTY + @delta` fires it). That is **CORRECT, intended, complementary** behavior: the LEDGER
> records the *movement*, the AUDIT trigger records the *row-state change*. They are not duplicative.
> `UPDATE_PartNumber` does **not** move stock (it writes audit tables, never re-mutates
> `INV_PARTS_STOCK_MST.IN_QTY`) — so it cannot double-count. Any incidental duplicate audit row is
> harmless and ignored. **No body re-verification is a cutover gate** (the prior "confirm audit-only"
> open item is closed by decision).
>
> **The `DISABLE TRIGGER UPDATE_PartNumber` in `test_ledger_opening_balance.py` (and `test_seam_driver.py`)
> is a TEST-ISOLATION measure** (keeps audit noise out of the ledger assertions), **NOT a cutover
> instruction.** Do NOT translate it into a `DISABLE`/`DROP` step at cutover. `UPDATE_PartNumber` stays
> live in production.

### 5b. ✅ RESOLVED — triage of ALL direct `INV_PARTS_STOCK_MST.IN_QTY` writers (the 4th/5th-writer hunt)
Enumerated every proc AND trigger that does a direct `UPDATE INV_PARTS_STOCK_MST … SET IN_QTY = …`
(grep of `/tmp/inv_utf8.sql`, excluding IN_QTY columns on OTHER tables — `INV_ASN_DETAIL_MST`,
`INV_OPEN_ORDER_INF`, `INV_PART_SHIPPING_INF`). Result: **the 12 trigger legs + 2 non-trigger procs.**

| Proc / trigger | Writes `INV_PARTS_STOCK_MST.IN_QTY`? | Verdict | Evidence | Cutover action |
|---|---|---|---|---|
| `INSERT/UPDATE/DELETE_RecConfStatPartsStockMstQTY` | yes (`PS.IN_QTY ± i/d.IN_QTY`) | **(b) ALREADY-A-PRODUCER** | RecConfStat path; lines 7492/5466/7565 | DROP at cutover (in the 13) → replaced by `receiving` seam |
| `INSERT/UPDATE/DELETE_RejectParts` | yes (`PS.IN_QTY ∓ i/d.IN_QTY`) | **(b) ALREADY-A-PRODUCER** | Reject path; lines 4608/5207/4465 | DROP at cutover → `reject` seam |
| `INSERT/UPDATE/DELETE_Stocktaking` | yes | **(b) ALREADY-A-PRODUCER** | Stocktaking path; lines 4701/5290/4384 | DROP at cutover → `stocktaking` seam |
| `InsertPartShipping`/`UpdatePartShipping`/`DeletePartShipping` (+ `DeleteShipDate` cascade) | yes | **(b) ALREADY-A-PRODUCER** | Shipping path; lines 2904/2813/2919 | DROP at cutover → `shipping` seam |
| `UPDATE_PartsStockRenban` (proc, line 4051) | **NO** — only `IN_RENBAN_COUNT` + `VC_LAST_UPDATE` | **out of scope** (not an IN_QTY writer) | line 4065-4067: `SET IN_RENBAN_COUNT=@RenbanCount` only | none for the ledger; it is the legacy renban-counter writer SUPERSEDED by carry 2 `RESERVE_RenbanCount` |
| `UPDATE_OrderRenbanQty` (proc, line 5890) / `UPDATE_OrderQty` (line 5930) | **NO** — write `INV_OPEN_ORDER_INF.IN_QTY` (the order-row qty), not the stock master | **out of scope / (b) feeder** | proc body updates `INV_OPEN_ORDER_INF`; that write FIRES `UPDATE_RecConfStatPartsStockMstQTY` (already a producer). Caller `RenbanOrder.pas:520` | none direct; the stock effect is carried by the receiving seam once triggers are dropped |
| `UPDATE_Shippingdetail` (line 288) / `INSERT_ShippingPartInfo` (line 405) | **NO** — write `INV_PART_SHIPPING_INF.IN_QTY` (shipping-detail row qty) | **out of scope / (b) feeder** | proc bodies update `INV_PART_SHIPPING_INF`; that write fires `Insert/UpdatePartShipping` (already a producer) | none direct; stock effect via `shipping` seam |
| **`UPDATE_PartsStockInfo` (proc, line 5691)** | **yes — absolute `SET … IN_QTY=@QTY …`** (line 5765) | **(c) SUPERSEDED + latent clobber → see carry 10** | Live caller `DataModule.pas:1482` (`'dbo.UPDATE_PartsStockInfo;1'`, in live `DataModule` per `.dpr`). Rebuilt master Save calls the SAME proc with all 30 params incl. `@QTY` (`ignition-spike-log.md:102-114`). Legacy qty box is `ReadOnly` (`parts-stock-master.md:387`) so `@QTY` is the loaded value, not hand-keyed | Mostly superseded by read-only-qty, BUT the read-only is not yet ENFORCED in the rebuild → **carry 10** (neuter the `IN_QTY=@QTY` clause) |
| **`UPDATE_PartsStockInfoCount` (proc, line 4076)** | **yes — additive `SET IN_QTY = IN_QTY-@QTY …`** (line 4086) | **(a) DEAD/legacy — NO caller** | grep `UPDATE_PartsStockInfoCount` across ALL `*.pas` (live + dead) → **zero callers**; whole-repo grep → only spec/review prose. Keyed `VC_PART_NUMBER` | **retire at cutover** (drop the proc, or leave it dead+unwired). No ledger wiring needed because nothing calls it. **carry 10** records the decision |

**Net:** the only two non-trigger direct writers are `UPDATE_PartsStockInfo` (latent clobber, live caller)
and `UPDATE_PartsStockInfoCount` (dead, no caller). Neither is a true LIVE-&-UNWIRED 5th producer that
needs ledger-wiring — `UPDATE_PartsStockInfoCount` has no caller at all, and `UPDATE_PartsStockInfo` is a
master-data edit whose qty leg should be neutered, not ledgered. Both are handled by **carry 10**.

## 5c. NEW — neuter/retire the two non-trigger `IN_QTY` writers (carry 10)
See **carry 10** below. Sequenced in Phase A (additive proc edits, pre-window) per the architecture doc.

## 6. Per-edit version-stamp contracts (idempotency keys)
- **Shipping:** `AmendShipment` must re-stamp `VC_ADD` per edit (INV_PART_SHIPPING_INF has no VC_LAST_UPDATE;
  VC_ADD is the amend version token).
- **Stocktaking:** the edit path must write a NON-NULL `VC_LAST_UPDATE` (fixes the legacy Bug2 NULL).
- All producers' amend keys use `:upd:to=<effect>:v=<stamp>` (collision-safe); keep that on any new write path.

## 7. Headless Jython driver coverage  (retro R8 — ✅ CLOSED 2026-06-19)
`scripts/e2e/jython_shim.py` + `test_seam_driver.py` run the REAL stocktaking + reject + **shipping +
receiving** wrappers end-to-end (autocommit shim) — including shipping's header-delete cascade and
receiving's `resolveAddPoint` add-point gate (asserted with a NON-zero IN_QTY delta on a counted 'S'
order). **Order is now covered too** — the §7 gap is **CLOSED.** The shim gained a
**persistent-sqlcmd-session** transaction extension (`_TxSession`): `beginTransaction` opens ONE
long-lived `docker exec -i sqlcmd` connection and feeds it framed batches (each batch terminated by a
unique `PRINT '<<<EOB:nonce>>>'` + `GO` sentinel read off a background-thread queue), so a `BEGIN TRAN`
spans statements as a real transaction. `runPrepUpdate(...,tx=session)` routes onto that connection;
`commit/rollback/closeTransaction` drive it; the autocommit path (tx None / `"tx-noop"`) is UNCHANGED.
`scripts/e2e/test_seam_driver_order.py` (13/13) drives the REAL `order.commitOrders`: happy-path commit
(2 records + counter advance + IN_QTY unmoved, persisting after close) **and a mid-transaction rollback
proof** (first INSERT shown visible *inside* the open tx, forced failure → 0 rows persist + counter
unchanged = cross-statement atomicity). Regression: `test_seam_driver.py` 23/23.
**Framing caveat:** when a batch raises a T-SQL error, sqlcmd ABORTS the rest of that batch, so the
trailing sentinel PRINT never runs — `_TxSession` watches for `Msg NNNN` lines and raises a `SqlError`
immediately (which is what fires `commitOrders`' except→rollback). `rollback()` guards with
`IF @@TRANCOUNT > 0` so a tx already auto-doomed by a fatal error doesn't raise Msg 3903 and mask the
original error. A true in-gateway runtime test (Perspective/Playwright is trial-gated; no WebDev; no
gateway-event infra) remains a gap — the shim covers the driver LOGIC/SQL + the real transaction lifecycle,
not Jython-2.7-vs-CPython runtime quirks. (Note: extending to shipping/receiving surfaced + fixed a narrow
shim coercion gap — digit-only VARCHAR business keys were being int-coerced on dataset round-trip; see
`jython_shim.py::_PyDataset._coerce`.)

## 8. Postgres-phase (D13) — the `# IG83-TODO` set
`TS_POSTED`/`VC_ADD`/`VC_LAST_UPDATE` 16-char strings → `datetime2`; `site_id` NOT NULL FKs (D1) +
per-site scoping; real PKs/FKs on the source tables; the manifest-cost computed-view projection option;
`fn_ManifestCostAt` string-date compare → typed date. Tracked per `# IG83-TODO` in the code.

## 9. Spec-hygiene residue  (retro R11)
Pre-D9 spec PROSE carries stale claims vs the live dump (e.g. `shipping.md` §2 "No declared FKs" — the live
DB HAS `FK_INV_PART_SHIPPING_INF_INV_SHIPPING_INF`; the manifest unique-index framing). The live-truth gate
catches these at build, but fix-on-touch (or a one-pass sweep) keeps the specs trustworthy.

## 10. ✅ RESOLVED (David ACCEPTED 2026-06-19) — neuter `UPDATE_PartsStockInfo` qty leg + retire `UPDATE_PartsStockInfoCount`  (SHOULD-FIX 3 / 4th-5th-writer hunt)

> ✅ **RESOLVED (David ACCEPTED 2026-06-19). Artifact:
> `docs/analysis/master-data/spike-partsstockinfo-drop-qty-clause.sql` (built + spike-validated).**
> Drop the `IN_QTY=@QTY` clause from `UPDATE_PartsStockInfo` (the `@QTY` param is KEPT in the signature but
> now unused — the rebuilt Save's positional 30-param call is unchanged); on-hand is **NEVER editable** on
> the rebuilt Parts Stock master (all qty change via a ledger transaction). `UPDATE_PartsStockInfoCount` is
> DEAD (zero callers) → **retired (`DROP PROCEDURE`, guarded).** Spike validation: legacy proc moved IN_QTY
> (7382→17381); after the ALTER, IN_QTY stayed 7382 while other columns still updated; spike restored
> to as-found.

After the seams own `IN_QTY` (Phase C), any path still doing a direct, non-ledger `IN_QTY` write to
`INV_PARTS_STOCK_MST` desyncs the materialized balance from `SUM(ledger)` — a lost update with no ledger
row. The 5b triage found exactly two such procs:

- **`UPDATE_PartsStockInfo` (line 5691) — latent clobber. ACTION: drop the `IN_QTY=@QTY` clause** from the
  proc body so the master-data edit no longer touches on-hand (it is a master-data edit, not a stock move;
  every other column stays). This is an **additive Phase-A proc edit** (apply with the other DB objects,
  before the window — it is safe pre-cutover too because during parallel run the rebuilt master loads the
  current `@QTY` and re-writes the same value, a no-op-in-effect, and after the edit it simply stops
  sending it). The rebuilt Perspective Save passes all 30 `sys.parameters` positionally
  (`ignition-spike-log.md:102-114`); dropping the body clause means the param is accepted-and-ignored (no
  Ignition-side change needed — single-point DB edit, Named-Query practice). **Confirm:** the rebuilt
  PartsStock form keeps the qty field read-only so it can never send a *hand-keyed* `@QTY` even before the
  proc edit lands (legacy `Quantity_MaskEdit` is `ReadOnly`, `parts-stock-master.md:387`; the rebuild must
  preserve that). Recommended belt-and-suspenders, **flag to David**: the on-hand-override domain question
  (`parts-stock-master.md` §8 Q1) — recommend on-hand is NEVER editable on the master; all qty change goes
  through a receiving/shipping/reject/stocktaking/adjustment transaction (i.e. the ledger).
- **`UPDATE_PartsStockInfoCount` (line 4076) — DEAD, no caller. ACTION: retire** (leave dead/unwired, or
  `DROP PROCEDURE` for hygiene). Whole-repo grep finds **zero** Delphi callers (live or legacy) and no
  rebuilt-lib reference — it is the additive `IN_QTY = IN_QTY-@QTY` "line-pull/count" writer the
  `parts-stock-master.md:183` note flagged, but nothing invokes it. Because it has no caller it does NOT
  need ledger-wiring as a 5th producer; it simply must not be resurrected. If a future flow needs that
  decrement, it routes through `stockLedger.post()`, not this proc. *(Rollback: re-create from
  `CreateInventory.sql` — but it is dead, so rollback is moot.)*

> **FLAG TO DAVID:** (1) accept dropping the `IN_QTY=@QTY` clause from `UPDATE_PartsStockInfo` (recommended);
> (2) confirm the on-hand-override policy (recommend: never editable on the master). Neither is a 5th
> ledger producer — `UPDATE_PartsStockInfoCount` is dead and `UPDATE_PartsStockInfo`'s qty leg is removed,
> not ledgered.

## 11. ✅ RESOLVED / CLOSED — Order-commit posts nothing to the ledger (NOT a gap)  (NIT 6, David 2026-06-19)
`order.commitOrders` inserts open-order rows (`INSERT_OpenOrder`) + advances the renban counter but **never
calls `stockLedger.post()`**. The concern was: if the worksheet could ever emit an order **already stamped
shipped/arrived** at creation, the legacy `INSERT_RecConfStatPartsStockMstQTY` trigger would move stock but
the rebuilt Order path would not (triggers dropped, no ledger post in this path) — a silent gap.

> ✅ **RESOLVED — CLOSED, NO WORK (David 2026-06-19).** The Order worksheet only emits NEW orders for the
> **calculated FUTURE FRS date**. An already-shipped/arrived order **cannot exist in the system on that
> future FRS date** — so a worksheet-committed order is always effect-0 at creation by construction. The
> Order-commit path therefore needs **NO ledger post**, and no conditional `stockLedger.post()` is added.
> This is **RESOLVED/closed, not a gap.** (`order.commitOrders` correctly posts nothing; the receiving seam
> remains the producer for the later arrival/confirm of that order.)
