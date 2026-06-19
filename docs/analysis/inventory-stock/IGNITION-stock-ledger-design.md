# IGNITION — Stock-Ledger Service Design

**Area:** Inventory / Stock (cross-cutting service)  **Status:** 🟡 design — for adversarial review, then build
**Target:** Ignition 8.3 semantics, runnable on 8.1.52 dev box  **Author:** Ignition architect / 2026-06-17

> **CORRECTION (2026-06-17, D9):** This design was revised after an adversarial review against the **live**
> `DB Schema/CreateInventory.sql`. The earlier draft cited a `Purge.PurgeMode` trigger bypass that **does not
> exist in the live DB** — it was an artifact of the stale `Create Inventory.superseded-2026-06-01.sql`
> (superseded lines 9665-9669). The real purge is `DELETE_AutoPurge` (§3.1). **Verify every trigger/purge claim
> against the live `CreateInventory.sql`, never the spaced superseded file.** Fixes applied: F1 (purge rationale
> re-derived from `DELETE_AutoPurge` + the `VC_TERMINATED=''` DELETE-trigger gate), F2 (D11#9 dropped from the
> divergence set — unreachable), F3 (new 5th divergence class: legacy multi-row-trigger under-count), F4
> (`rebuildBalance` locking), F5 (part-number-changing UPDATE → two-post derivation). The core architecture
> (additive ledger-as-truth, materialized `IN_QTY += delta`, all trigger signs, D8(3)/D12#3 divergences,
> materialized-vs-computed) is CONFIRMED SOUND and unchanged.

> **The service that OWNS the `INV_PARTS_STOCK_MST.IN_QTY` on-hand invariant.** Today twelve
> qty-adjusting triggers — owned by Receiving, Reject, Stocktaking, Shipping — each *directly* mutate
> the on-hand balance, with inconsistent keying (string `VC_PART_NUMBER` vs int `IN_PART_ID`),
> inconsistent gates, one dead branch (D8(3)), and a multi-row-trigger under-count hazard (the
> `FROM PS, inserted i` / `JOIN inserted` shape applies one arbitrary row per part, not the sum), and no
> single audit truth. This design re-homes all twelve into **ONE additive-delta ledger** where the **ledger of
> movements is the source of truth** and on-hand becomes a **derived/materialized balance**.
>
> This is the foundation the receiving/shipping/reject/stocktaking SCREENS will later call. It is a
> **service** — it has no screen of its own (the InvMgmt/Stocktaking/RecConfStat/Shipping screens
> drive it). Source behavior is fixed against the live trigger bodies in
> [`parts-stock-master.md` §2](parts-stock-master.md), [`recconfstat.md`](../receiving/recconfstat.md),
> [`recreject.md`](../receiving/recreject.md), [`shipping.md`](../shipping/shipping.md),
> [`stocktaking.md`](stocktaking.md), with decisions D1/D5/D7/D8(3)/D12#3 baked in. (The source specs
> `parts-stock-master.md` / `recconfstat.md` still cite the non-existent `PurgeMode`; those corrections are
> tracked separately — this design uses the **live** `DELETE_AutoPurge` mechanism, §3.1.)

---

## 1. The inversion — ledger as source of truth, on-hand as derived

**Today (legacy):** `IN_QTY` is the source of truth. Twelve triggers do `IN_QTY = IN_QTY ± x` in place.
`INV_PART_QTY_INF` (7546 rows) is only an *audit side-effect* written by `UPDATE_PartNumber` when
`IN_QTY` happens to change — and the 12 triggers DON'T write it (they only touch `IN_QTY` +
`VC_LAST_UPDATE`). So today's ledger is incomplete and non-authoritative.

**Rebuild:** invert it. The **movements ledger is authoritative**; `IN_QTY` is

```
on_hand(part) = SUM(qty_delta) over all ledger movements for that part
```

Every stock event is an **append-only signed delta posting**. Nothing ever does `SET IN_QTY = <absolute>`
again (the legacy "clobber via UPDATE_PartsStockInfo with a different @QTY" path is gone — D5/PSM§4:
on-hand is never a freely-editable field; a correction is an explicit adjustment *delta* posting).

### Materialized running balance vs computed-on-read — **DECISION: materialized, with the ledger as the rebuild authority**

Use a **materialized running balance**: each post (a) appends a ledger row AND (b) does the additive
`UPDATE INV_PARTS_STOCK_MST SET IN_QTY = IN_QTY + @delta` on the same part, in one transaction. Reasons:

1. **The legacy/read paths need the column populated during parallel run.** `SELECT_PartsStockInfo`,
   `REPORT_LogicalInventory`, the order-explosion reads, and the legacy Delphi app *all read
   `IN_QTY` directly*. A pure compute-on-read (no stored column) would orphan every legacy reader.
   We MUST keep `IN_QTY` current as a materialized column.
2. **It preserves the exact legacy shape** — the additive `IN_QTY += delta` is *byte-identical in effect*
   to what the 12 triggers do today (they are all `IN_QTY = IN_QTY ± x`). This is what makes parallel-run
   parity provable (§6): same additive operation, just funneled through one writer with a ledger row attached.
3. **Compute-on-read SUM over a 7546+ row (and growing) ledger** is fine for one part but a full
   `GROUP BY part` for a 45-screen inventory list would re-sum the entire ledger on every page load —
   needless. The materialized column is O(1) to read.

**The ledger is still the *authority*:** `IN_QTY` is a cache of `SUM(qty_delta)`. The reconciliation
harness (§6) re-derives `SUM(qty_delta)` and asserts it equals the stored `IN_QTY` — that diff IS the
parity/health check. A drift means a bug, not a tolerated state.

> **OPEN QUESTION (for the reviewer/David):** materialized is recommended for the reasons above. The
> alternative — a DB *computed/indexed view* `SUM(qty_delta) GROUP BY part` projected back as `IN_QTY` —
> is cleaner long-term but (a) can't be a plain updatable column the legacy app writes, and (b) needs the
> Postgres phase. Recommend **materialized now**, revisit a computed projection at the Postgres phase
> (`# IG83-TODO`). Flagging as a genuine fork, not guessing.

---

## 2. The ledger data model — `INV_STOCK_LEDGER` (formalizes / replaces `INV_PART_QTY_INF`)

A new movements table. One row = one signed stock movement. Append-only (no UPDATE/DELETE of a posted
movement — corrections are *new* compensating rows, §3 reversal contract).

| Column | Type | Role |
|--------|------|------|
| `IN_LEDGER_ID` | `int IDENTITY` PK | Surrogate id, the movement's own key. |
| `IN_PART_ID` | `int NOT NULL` FK→`INV_PARTS_STOCK_MST` | **The unified key.** Resolve the legacy string `VC_PART_NUMBER` → surrogate **once at the posting boundary** (D2). Ends the string/int inconsistency: receiving/shipping (string-keyed today) and reject/stocktaking (int-keyed today) all post by `IN_PART_ID`. |
| `IN_QTY_CHANGE` | `int NOT NULL` | **Signed delta.** `+` adds on-hand, `−` removes. Same name/semantics as `INV_PART_QTY_INF.IN_QTY_CHANGE` (parallel-run parity). |
| `IN_BALANCE_AFTER` | `int NULL` | Running on-hand AFTER this post (the materialized snapshot at post time). Mirrors `INV_PART_QTY_INF.IN_QTY`. Optional but cheap; lets the harness verify monotonic replay without re-summing. |
| `VC_SOURCE_ENUM` | `varchar(24) NOT NULL` | The movement class (table below). **This is the new analytic axis the legacy lacked** — `INV_PART_QTY_INF` only carried `VC_STATUS` `'U'`. |
| `IN_SOURCE_ROW_ID` | `int NULL` | The surrogate id of the source row that caused the movement (`IN_ORDER_ID` / `IN_REJECT_ID` / `IN_STOCKTAKING_ID` / `IN_PART_SHIPPING_ID`). The reason/ref key — makes a movement traceable and reversible. |
| `VC_SOURCE_EVENT` | `varchar(40) NOT NULL` | The **idempotency key** discriminator: a stable string identifying the *event* that produced this delta (e.g. `RECEIVING_ARRIVAL:ord=8842:v=3`). §4. Unique together with `IN_PART_ID`. |
| `site_id` | `int NOT NULL` FK→`sites` | D1. Every movement is site-scoped; the part it references is already per-site. `# IG83-TODO` (lands Postgres phase; parallel-run is single-site). |
| `VC_REASON` | `varchar(300) NULL` | Free-text/coded reason (stocktaking reason, reject reason, "arrival reversal", "purge — see note"). |
| `TS_POSTED` | `varchar(16) NOT NULL` / `datetime2` | Post timestamp. **Parallel run:** the 16-char `yyyymmddHHMMSSff` string (P2, matches `VC_ADD`/`VC_LAST_UPDATE`). `# IG83-TODO`: real `datetime2` at Postgres phase. |
| `VC_ADD` | `varchar(16) NOT NULL` | Audit insert stamp (mirror `INV_PART_QTY_INF.VC_ADD`). |

**Indexes:** PK `IN_LEDGER_ID`; covering `(site_id, IN_PART_ID)` for the per-part SUM/balance read;
**UNIQUE `(IN_PART_ID, VC_SOURCE_EVENT)`** — the idempotency backstop (§4). FK `IN_PART_ID` RESTRICT
(D3: can't delete a part that has ledger movements).

**Relationship to `INV_PART_QTY_INF`:** the legacy ledger's columns (`VC_PART_NUMBER`, `IN_QTY_CHANGE`,
`IN_QTY`, `VC_STATUS`, `VC_ADD`) map onto ours (`IN_PART_ID` via resolution, `IN_QTY_CHANGE`,
`IN_BALANCE_AFTER`, `VC_SOURCE_ENUM`/`VC_STATUS`, `VC_ADD`). During parallel run we can **derive a
parity view** shaped like `INV_PART_QTY_INF` from `INV_STOCK_LEDGER` to diff against the legacy audit
rows. We do NOT migrate or rely on the 7546 historical `INV_PART_QTY_INF` rows as truth — they're an
incomplete audit. The ledger truth is reconstructed by **replaying the live source tables** (§6).

---

## 3. The 12 trigger effects → posting operations (decisions baked in)

Every legacy trigger becomes one (or a pair of) `postMovement(part_id, delta, source_enum, source_row_id,
event_key, reason)` calls. The **gate** (when to post / what sign) is evaluated in the service, not the DB.

| # | Legacy trigger | Source table | `VC_SOURCE_ENUM` | Sign / delta | Gate (evaluated in service) | Notes & decision |
|---|----------------|--------------|------------------|--------------|------------------------------|------------------|
| 1 | `INSERT_RecConfStatPartsStockMstQTY` `'S'` leg | OPEN_ORDER | `RECEIVING_SHIP` | `+IN_QTY` | supplier add-point `'S'` AND `VC_STATUS_SUPPLIER_SHIPPING<>''` | Add at supplier-shipping. |
| 2 | `INSERT_…` `'A'` leg | OPEN_ORDER | `RECEIVING_ARRIVAL` | `+IN_QTY` | add-point `'A'` AND (arrival **or** plant-yard **or** assembler-yard **or** warehouse set) | **D7**: arrival-add lives in receiving-confirmation. **D12#3**: plant-yard & assembler-yard count as arrival. |
| 3 | `UPDATE_…` legs 1–2 (qty change, shipped/arrived) | OPEN_ORDER | `RECEIVING_SHIP` / `RECEIVING_ARRIVAL` | `+(i.IN_QTY − d.IN_QTY)` *(same part)*; **two posts if part changed (F5)** | qty changed AND still shipped(`'S'`)/arrived(`'A'`) | Post the **net delta** when the part is unchanged. **If `VC_PART_NUMBER` changed** (live legs join `i`/`d`/`PS` on `VC_PART_NUMBER`), `i`/`d` are different part ids → emit `−oldEffect` on old part + `+newEffect` on new part (§4.3 F5). Legacy silently no-ops a part change here. |
| 4 | `UPDATE_…` leg 3 (ship-status set) | OPEN_ORDER | `RECEIVING_SHIP` | `+IN_QTY` | `'S'`, ship-status blank→set | Mirror of #1 on edit. |
| 5 | `UPDATE_…` leg 4 (ship-status cleared) | OPEN_ORDER | `RECEIVING_SHIP` | `−IN_QTY` | `'S'`, ship-status set→blank | Reverse of #1. |
| 6 | `UPDATE_…` leg 5 (arrival set) | OPEN_ORDER | `RECEIVING_ARRIVAL` | `+IN_QTY` | `'A'`, arrival blank→set (**D12#3**: also plant-yard/assembler-yard blank→set) | **D7** add. The only path counting `'A'` stock today. |
| 7 | `UPDATE_…` leg 6 (arrival cleared) | OPEN_ORDER | `RECEIVING_REVERSAL` | `−IN_QTY` | `'A'`, arrival set→blank (**D12#3**: plant-yard/assembler-yard set→blank) | ⚠️ **D8(3): DEAD in legacy** (`i.VC_ARRIVAL='' AND i.VC_ARRIVAL<>''`). **IMPLEMENT the reversal.** **Intentional divergence — "ledger correct, legacy buggy."** |
| 8 | `DELETE_RecConfStatPartsStockMstQTY` `'S'`/`'A'` legs | OPEN_ORDER | `RECEIVING_SHIP`/`_ARRIVAL` | `−d.IN_QTY` | counted (shipped/arrived) AND **`VC_TERMINATED=''`** AND `VC_STATUS_EMPTY_TRAILER=''` AND add-point S/A | Remove on delete of a still-active (non-terminated) order. **No `PurgeMode` exists in the live trigger (D9)** — the live gate is just `@numrows>0` + these status filters. The purge job (`DELETE_AutoPurge`) pre-stamps `VC_TERMINATED`, so purged rows already fail the `VC_TERMINATED=''` gate and never drain on-hand (§3.1). The service posts the `−d.IN_QTY` reversal ONLY for a genuine user delete of a non-terminated order. |
| 9 | `INSERT_RejectParts` | REJECT | `REJECT` | `−i.IN_QTY` | (none) | Every reject subtracts. |
| 10 | `UPDATE_RejectParts` | REJECT | `REJECT` | `+(d.IN_QTY − i.IN_QTY)` net | (none); part is immutable on edit | Re-balance by delta. |
| 11 | `DELETE_RejectParts` | REJECT | `REJECT` | `+d.IN_QTY` | (none) | **D11#9 RETRACTED (F2):** the only reject delete is `DELETE_RecProdRejInfo` (single-row user delete by `@RejectID`). `DELETE_AutoPurge` deletes ONLY open-orders + HIST — it **never** touches `INV_REJECT_INF`. So "purge inflates on-hand via reject delete" is **UNREACHABLE**. A reject delete is always a genuine user un-reject → **always restores on-hand** (faithful legacy behavior). No purge branch here. |
| 12 | `INSERT_Stocktaking` | STOCKTAKING | `STOCKTAKING` | `+i.IN_QTY` (delta may be `−`) | (none) | **D5**: `IN_QTY` is a **signed adjustment delta**, not absolute. Post it verbatim. |
| 13 | `UPDATE_Stocktaking` | STOCKTAKING | `STOCKTAKING` | `+(i.IN_QTY − d.IN_QTY)` net | (none); part immutable | Re-balance by change in the delta. Fix the D8/Bug2 NULL-timestamp by writing a real `TS_POSTED`. |
| 14 | `DELETE_Stocktaking` | STOCKTAKING | `STOCKTAKING` | `−d.IN_QTY` | (none) | Reverse the adjustment. Insert+delete is qty-neutral. |
| 15 | `InsertPartShipping` | PART_SHIPPING | `SHIPPING` | `−i.IN_QTY` | (none) — always subtract, **no add-point** | Stock-OUT at production. `IN_QTY = round(built×ratio/100)`. |
| 16 | `UpdatePartShipping` | PART_SHIPPING | `SHIPPING` | `−(i.IN_QTY − d.IN_QTY)` net | (none) | Detail edit re-balance. |
| 17 | `DeletePartShipping` (and the `DeleteShipDate` cascade) | PART_SHIPPING | `SHIPPING` | `+d.IN_QTY` | (none) — but **scope the header-delete restore to (site, line, production_date)**, fixing the line-blind cascade | Restore on line/shipment delete. `DELETE_AutoPurge` never deletes `INV_PART_SHIPPING_INF` (F2), so a shipping delete is always a genuine user delete → always restores. **Multi-row caution (F3):** the live `InsertPartShipping` trigger uses `FROM PS, inserted i WHERE PS.VC_PART_NUMBER=i.VC_PART_NUMBER` — a header delete spanning >1 same-part row under-counted in legacy; the rebuild restores each row (§6 class 5). |

> The "12 triggers" expand to 17 posting operations because several triggers contain multiple legs.
> The 12 *triggers* are: 3 RecConfStat + 3 Reject + 3 Stocktaking + 3 PartShipping = 12 (the
> `DeleteShipDate` header cascade is a 13th trigger that fires `DeletePartShipping`, folded into #17).

### 3.1 Purge vs an append-only ledger — **the key architectural call** (re-derived from the LIVE purge, F1/D9)

**There is NO `PurgeMode` flag in the live DB.** The earlier draft cited a `Purge.PurgeMode` bypass inside
`DELETE_RecConfStatPartsStockMstQTY`; that belonged to the stale `Create Inventory.superseded-2026-06-01.sql`
(lines 9665-9669) and **does not exist** in the authoritative `CreateInventory.sql`. The live DELETE trigger's
qty-removal is gated only by `@numrows>0` plus the status filters — crucially **`AND VC_TERMINATED=''`** (verified
live, both the `'S'` and `'A'` legs). There is no `Purge` table.

**The real purge is `DELETE_AutoPurge(@DataRetention)`** (verified live). It does two things, in order:

1. **Pre-stamp termination:** `UPDATE INV_OPEN_ORDER_INF SET VC_TERMINATED = <future date> WHERE VC_ADD <= <cutoff> AND VC_TERMINATED=''` — every aged, still-active order is *marked terminated* first.
2. **Delete aged rows:** `DELETE FROM INV_OPEN_ORDER_INF`, `INV_OPEN_ORDER_INF_HIST`, and `INV_PARTS_STOCK_MST_HIST` for `VC_ADD <= <cutoff>`. **It deletes ONLY open-orders + HIST tables — never `INV_REJECT_INF`, never `INV_PART_SHIPPING_INF`.**

**So legacy purge avoids draining on-hand NOT via a flag, but as an emergent property of the gate:** the step-1
pre-stamp sets `VC_TERMINATED <> ''`, so when step-2's `DELETE` fires `DELETE_RecConfStatPartsStockMstQTY`, every
deleted row **fails the trigger's `VC_TERMINATED=''` gate** and contributes no `−IN_QTY`. Termination, not a
PurgeMode, is what spares on-hand.

This re-grounds — but does not change — the append-only design:

- **The append-only ledger never reverses a real movement on a housekeeping delete.** A purge deletes source
  rows; it posts nothing. The historical receipt/shipment already happened; deleting the *record* must not
  un-happen it. The ledger keeps the movement row forever (append-only), so on-hand stays correct AND auditable
  even after the source order is purged — strictly better than legacy, which lost the audit at purge.
- **The rebuild signals this with the context flag `purge=True` on the posting call** (the service-level analog
  of "this DELETE came from `DELETE_AutoPurge`, not a user un-doing an event"). It is **not** a DB column; it
  mirrors the live behavior that purge-driven deletes do not move stock.
- **The faithful equivalent of the live `VC_TERMINATED=''` gate:** a user-deleted order that is **already
  terminated** (`VC_TERMINATED<>''`) also posts no reversal (the live trigger skips it too). The service applies
  the same gate (table row #8), so terminated-then-deleted orders match legacy with `purge=False` as well.
- A *genuine* user-initiated delete/reversal of a **non-terminated** order DOES post the compensating delta
  (#8). Reject delete (#11) and shipping delete (#17) are **always** genuine user deletes — `DELETE_AutoPurge`
  never touches those tables (F2), so there is no purge case to special-case for them.

> **The distinction the service must honor:** *reverse-because-the-event-was-undone* (post a compensating
> delta, for a non-terminated user delete) vs *delete-the-record-for-housekeeping or already-terminated*
> (post nothing). Legacy expressed "post nothing" via the live `VC_TERMINATED=''` gate + the `DELETE_AutoPurge`
> pre-stamp; the rebuild expresses it via the `VC_TERMINATED<>''` check and the `purge=True` flag. Same effect,
> no fictional `PurgeMode`.

---

## 4. Idempotency + concurrency contract

The legacy had real read-then-write races (the renban counter D11#7; concurrent `IN_QTY = IN_QTY ± x`
on the same part). The ledger must be safe under replay and concurrency.

**Posting contract — `postMovement` is the ONLY writer of `IN_QTY`:**

1. **Atomic per post.** One DB transaction does: (a) `INSERT INV_STOCK_LEDGER`, (b)
   `UPDATE INV_PARTS_STOCK_MST SET IN_QTY = IN_QTY + @delta, VC_LAST_UPDATE = @ts WHERE IN_PART_ID=@id`.
   Either both land or neither. This is exactly the atomicity SQL Server gave the legacy triggers — preserved.

2. **Idempotency key = `(IN_PART_ID, VC_SOURCE_EVENT)` UNIQUE.** `VC_SOURCE_EVENT` encodes the *event*
   that caused the delta, including a **version/leg discriminator**, e.g.:
   - `RECEIVING_SHIP:ord=8842:set` (ship-status went set)
   - `RECEIVING_ARRIVAL:ord=8842` / `RECEIVING_REVERSAL:ord=8842`
   - `REJECT:rej=551:ins` / `REJECT:rej=551:del`
   - `STOCKTAKING:stk=903:ins`
   - `SHIPPING:psh=12077:ins`

   A replay (retry after a transport blip, a double-fire from a flaky screen) hits the UNIQUE constraint
   → the second insert is **rejected/no-op**, and crucially the `IN_QTY += delta` does NOT run again.
   **No double-post.** This replaces the legacy's fragile reliance on trigger-once semantics.

3. **UPDATE = reverse-old + post-new, expressed as one net delta — *unless the part identity changed*.**
   A source-row edit *that keeps the same part* (order qty change, stocktaking delta change, shipping detail
   qty change) posts a SINGLE row with `delta = (newEffect − oldEffect)` and a fresh `VC_SOURCE_EVENT`
   (`…:upd:v=N`). One net-delta row keeps the ledger compact and the idempotency key clean.
   *(The legacy triggers used two SQL statements for this; the net effect is identical and we capture it
   as one delta — confirmed equal in §6 parity.)*

   **F5 — part-number-changing UPDATE → TWO posts (two part ids).** `UPDATE_RecConfStatInfo` can re-point an
   order's `VC_PART_NUMBER`. The live `UPDATE_RecConfStatPartsStockMstQTY` qty legs join
   `inserted i ON i.VC_PART_NUMBER = d.VC_PART_NUMBER` and `PS ON PS.VC_PART_NUMBER = d.VC_PART_NUMBER`
   (verified live) — so when the part number changes, `i` and `d` resolve to **different** `IN_PART_ID`s and the
   single-net-delta assumption breaks. The service MUST **detect a part-id change** (`oldPartId != newPartId`)
   and emit the genuine pair: **`post(oldPartId, −oldEffect, …:upd:del-old)`** AND
   **`post(newPartId, +newEffect, …:upd:add-new)`** — two ledger rows, two part ids. (Note: the live trigger's
   own self-join silently does *nothing* on a part-number change, since `i.VC_PART_NUMBER<>d.VC_PART_NUMBER`
   yields no matching row — another legacy gap; the rebuild's two-post is the *correct* behavior and will
   surface as an EXPECTED-DIVERGENT pair, §6.) The `(IN_PART_ID, VC_SOURCE_EVENT)` idempotency key already
   includes the part id, so the two posts carry distinct keys cleanly. Under **D2** (keying moves to
   `IN_PART_ID`) this case mostly evaporates, but **parallel-run legacy keys on the string** — handle it.

4. **DELETE = post the reversal** (compensating `−original`/`+original`), UNLESS `purge=True` (§3.1).

5. **Concurrency: single-writer funnel + row-scoped serialization.** All posts go through the one gateway
   service. Within a post, the `UPDATE … SET IN_QTY = IN_QTY + @delta` is itself atomic and
   **commutative** (additive deltas commute — that's the whole point of the additive model), so two
   concurrent posts to the *same* part serialize at the row lock and both apply correctly regardless of
   order. We do **not** need SERIALIZABLE for the balance math; **READ COMMITTED + the additive UPDATE +
   the UNIQUE idempotency key** is sufficient and avoids the legacy read-then-write hazard entirely
   (we never read `IN_QTY`, compute, and write back — we issue a relative `+= delta`).
   `# IG81-COMPAT`: works identically on 8.1.52 (it's a DB-level guarantee, not an Ignition feature).

> **Why not SERIALIZABLE (for `post()`):** the additive `IN_QTY += delta` needs no snapshot isolation — it
> never reads the prior value into the app. The only thing needing protection is double-posting, handled by
> the UNIQUE event key. This is simpler and faster than the legacy and removes a whole race class. **This
> no-SERIALIZABLE claim is scoped to `post()` only.**

6. **F4 — `rebuildBalance` is the ONE read-then-write; it MUST be serialized against `post()`.** The healing
   command `stockLedger.rebuildBalance(partId)` (§7) re-SUMs the ledger and re-stamps
   `IN_QTY = <absolute SUM>`. That **is** a cross-statement read-then-write: a `post()` delta that commits
   between the SUM and the absolute re-stamp would be **silently clobbered (lost update)**. Invariant:
   `rebuildBalance` must run under one of —
   - **(a)** an explicit row lock taken in the same transaction:
     `SELECT IN_QTY FROM INV_PARTS_STOCK_MST WITH (UPDLOCK, HOLDLOCK) WHERE IN_PART_ID=@id` **before** the
     SUM, holding it through the re-stamp (serializes any concurrent `post()` on that part at the row lock); or
   - **(b)** a `SERIALIZABLE` transaction wrapping `SUM(qty_delta) → SET IN_QTY = @sum`; or
   - **(c)** a quiesced maintenance window where no `post()` runs.

   `post()` itself remains READ COMMITTED + additive (no change). Only `rebuildBalance` carries this stronger
   guarantee, because only it does absolute (read-then-write) arithmetic. This is the single exception to the
   "we never read-compute-write `IN_QTY`" rule, and it is explicitly fenced.

---

## 5. Ignition realization (8.3 semantics, runnable on 8.1.52)

**Mechanism: a DB-side stored procedure `POST_StockMovement`, wrapped by a Project-Library gateway
service `stockLedger.post()`, called via `system.db.createSProcCall`.** Rationale:

- **The atomic insert-ledger + bump-IN_QTY belongs in ONE transactional unit.** A stored proc gives
  that atomicity natively (the legacy did it in triggers; we keep it server-side). `createSProcCall`
  is the architect-standard way to wrap a proc with IN params from a gateway script. The UNIQUE-key
  idempotency and the `IN_QTY += @delta` live in the proc body — minimal logic, maximally testable.
- **A thin Jython service in a Project Library** (`stockLedger`) is the single funnel every module calls:
  `stockLedger.post(partId, delta, sourceEnum, sourceRowId, eventKey, reason, purge=False)`. It resolves
  `VC_PART_NUMBER → IN_PART_ID` at the boundary when a caller only has the string (D2), assembles the
  `VC_SOURCE_EVENT`, and invokes `POST_StockMovement`. The screens (RecConfStat/Shipping/Reject/
  Stocktaking) never touch `IN_QTY` — they call `stockLedger.post()`.
- **Not inline `runPrepQuery` in views** (the masters' pattern): this is a *service*, not a screen.
  It must be one reusable, transactional, idempotent entry point — the opposite of scattered inline SQL.
- **Named Queries** still mirror the source-table CRUD (per `ignition-named-query-crud-practice`):
  `InsertOpenOrder`, `InsertReject`, etc. wrap the existing `INSERT_*` procs. In **parallel run** those
  procs' *legacy triggers stay live* and keep `IN_QTY` correct on the legacy DB. The new
  `stockLedger.post()` path runs **in parallel against the rebuild's own ledger** (§6) — it does not
  replace the triggers during parallel run; it shadows them and is proven equal first.

```
# stockLedger.post(...)  — Project Library, Jython 2.7
#   resolve part (string→id if needed), build event key, call POST_StockMovement
# IG81-COMPAT: createSProcCall + a stored proc — identical on 8.1.52 and 8.3.
# IG83-TODO:  at Postgres phase, TS_POSTED → datetime2; site_id FK enforced; consider
#             a computed-view projection of SUM(qty_delta) instead of the materialized column (§1).
call = system.db.createSProcCall("POST_StockMovement", database)
call.registerInParam("partId", system.db.INTEGER, partId)
call.registerInParam("delta",  system.db.INTEGER, delta)
call.registerInParam("sourceEnum", system.db.VARCHAR, sourceEnum)
call.registerInParam("sourceRowId", system.db.INTEGER, sourceRowId)
call.registerInParam("eventKey", system.db.VARCHAR, eventKey)   # the (part,event) idempotency key
call.registerInParam("reason", system.db.VARCHAR, reason)
call.registerInParam("purge", system.db.BIT, purge)
system.db.execSProcCall(call)   # proc no-ops on duplicate eventKey; otherwise insert+IN_QTY+=delta atomically
```

`POST_StockMovement` body (sketch, lives in the rebuild DB, NOT the legacy DB during parallel run):
1. `IF @purge = 1 RETURN` (purge deletes source rows; never posts — §3.1; mirrors the live `DELETE_AutoPurge`
   pre-stamp + `VC_TERMINATED=''` gate, NOT a `PurgeMode` flag). The caller also passes `purge=False` but the
   service skips the post for an already-terminated order delete (the `VC_TERMINATED<>''` faithful gate, §3.1).
2. `IF EXISTS(SELECT 1 FROM INV_STOCK_LEDGER WHERE IN_PART_ID=@partId AND VC_SOURCE_EVENT=@eventKey) RETURN`
   (idempotency; also backstopped by the UNIQUE index).
3. `BEGIN TRAN` → compute `@ts` (16-char string, P2) → `INSERT INV_STOCK_LEDGER(...)` →
   `UPDATE INV_PARTS_STOCK_MST SET IN_QTY = IN_QTY + @delta, VC_LAST_UPDATE=@ts WHERE IN_PART_ID=@partId` →
   `COMMIT`.

> **Gateway transaction note (8.1 vs 8.3):** keeping the atomic unit *inside the proc* means we don't
> depend on Ignition's `system.db.beginTransaction` semantics differing across 8.1/8.3 — the DB owns
> atomicity. `# IG83-ONLY` paths are avoided here on purpose; the service is version-portable.

---

## 6. Parallel-run PARITY — the reconciliation harness

During parallel run the **legacy 12 triggers still mutate `IN_QTY` on the legacy SQL Server DB.** The
rebuild's ledger must be **provably equivalent** — and, where a decision FIXES a legacy bug, **provably
divergent in exactly the expected way** (the Order-spike "proc-fidelity gap" tagging, applied here as
"ledger correct, legacy buggy").

**The check (a reconciliation harness, scriptable like `scripts/gen_parity_tsv.sh`):**

For a snapshot of the live source tables (`INV_OPEN_ORDER_INF`, `INV_REJECT_INF`, `INV_STOCKTAKING_INF`,
`INV_PART_SHIPPING_INF`) and supplier add-points:

1. **Derive the ledger from source.** Replay every source row through the §3 mapping → produce the set of
   `(IN_PART_ID, qty_delta)` movements the rebuild *would* post. (This is a pure function of the source
   tables + add-point + the gates; no live posting needed for the check.)
2. **Sum per part:** `derived_on_hand[part] = SUM(qty_delta)` over the derived ledger.
3. **Diff vs legacy `IN_QTY`:** `diff[part] = derived_on_hand[part] − INV_PARTS_STOCK_MST.IN_QTY[part]`.
4. **Classify each non-zero diff:**
   - **EXPECTED-ZERO (must be 0):** every part not touched by a fixed bug. A non-zero here is a **real
     parity defect in the rebuild** — fix it.
   - **EXPECTED-DIVERGENT (tagged, must match the predicted sign/magnitude):**
     - **D8(3) arrival-reversal:** parts with an `'A'` order whose arrival was *set then cleared*. Legacy
       left on-hand **overstated** (dead branch); rebuild posts the `−qty` reversal. Predicted
       `derived < legacy` by exactly the cleared arrival qty. **Tag: "ledger correct, legacy buggy (D8(3))."**
     - **D12#3 plant/assembler-yard arrival on edit:** `'A'` orders where arrival-equivalence was stamped
       via plant-yard/assembler-yard *on an edit* (legacy UPDATE leg fired only on `VC_ARRIVAL`, so it
       **under-counted**); rebuild counts them. Predicted `derived > legacy`. **Tag: "ledger correct,
       legacy buggy (D12#3)."**
     - **F3 — legacy multi-row-trigger UNDER-COUNT (NEW 5th class, the most important harness fix):** SQL
       Server's `UPDATE PS … FROM PS, inserted i` / `FROM PS JOIN inserted i ON PS.VC_PART_NUMBER=i.VC_PART_NUMBER`
       shape (verified live in `INSERT_RejectParts`, `InsertPartShipping`, and the RecConfStat legs) does **not
       sum all matching `i` rows** — it applies **ONE arbitrary matching row per PS row**. So when a legacy
       qty-trigger fired on a batch containing **multiple same-part rows in one statement** (a renban batch, a
       multi-line shipment, a multi-row header-delete cascade), the legacy **under-counted** (it moved one row's
       qty, not the sum). The rebuild harness replays *every* source row and SUMs all of them → it will diverge,
       with **`derived > legacy`** for those parts. **This is the ledger being MORE correct than the buggy
       legacy** (frame like the Order spike's proc-fidelity gaps). The harness MUST **predict** this class, not
       treat it as a mystery EXPECTED-ZERO miss — otherwise a real rebuild bug could be excused as "ledger
       correct" (the fixture-fidelity failure mode, `feedback-parity-fixture-fidelity`). **Tag: "ledger correct,
       legacy buggy (F3 multi-row under-count)."** **Detection query (run on the snapshot to enumerate
       candidates):** find parts where a single trigger-firing statement plausibly carried >1 same-part row —
       e.g. `SELECT IN_SHIPPING_ID, VC_PART_NUMBER, COUNT(*) FROM INV_PART_SHIPPING_INF GROUP BY IN_SHIPPING_ID,
       VC_PART_NUMBER HAVING COUNT(*)>1` (confirmed live: **1** such shipment exists), and the analogous
       group-by on reject (`IN_PART_ID`) and open-order (`VC_PART_NUMBER`, **21** multipart groups live). Only
       batches where those same-part rows were inserted/updated/deleted **in one statement** under-counted; the
       harness predicts the magnitude per affected part from the dropped rows.
     - **F5 — part-number-changing UPDATE:** orders whose `VC_PART_NUMBER` was re-pointed by `UPDATE_RecConfStatInfo`.
       Legacy's `i`/`d` self-join on `VC_PART_NUMBER` finds no match → **no qty move at all** (silent no-op);
       the rebuild emits the `−oldPart` + `+newPart` pair (§4.3). Predicted: old part `derived < legacy`, new
       part `derived > legacy`, by exactly the order's effect. **Tag: "ledger correct, legacy buggy (F5 part-change)."**
     - **D8/Bug2 stocktaking NULL-timestamp:** affects `VC_LAST_UPDATE` only, **not the qty** — so it
       must NOT appear as a qty diff. Assert qty parity holds for these parts (timestamp is checked separately).
     - **Reject-delete restore is EXPECTED-ZERO (F2), not divergent.** `DELETE_RecProdRejInfo` is a single-row
       user un-reject that restores on-hand; the rebuild's `+d.IN_QTY` post matches it exactly. The old D11#9
       "reject-delete-during-purge inflates on-hand" class is **REMOVED** — `DELETE_AutoPurge` never deletes
       reject (or shipping) rows, so it is unreachable. Leaving it in the tag set would let the harness
       mis-attribute a genuine reject-delete bug to "ledger correct." Reject/shipping deletes must diff **0**.

**Pass criterion:** every EXPECTED-ZERO part diffs by 0; every EXPECTED-DIVERGENT part diffs by exactly
the predicted amount and is tagged. Any other non-zero diff is a rebuild bug. This is the *exact* discipline
the Order spike used (`VC_ADD`-tagged synthetic rows; separate fixture- vs proc-fidelity — see
`feedback-parity-fixture-fidelity`): **data adjudicates; green-on-buggy-data ≠ faithful.**

**Spike validation deliverable:** a harness (SQL + a small driver, sibling to `scripts/gen_parity_tsv.sh`)
that: dumps the source tables from the restored `Inventory.bak` sandbox → derives the ledger → sums per
part → joins to live `IN_QTY` → emits a TSV of `(part, derived, legacy, diff, classification, tag)`.
The reviewer can eyeball that the only non-zero rows carry a divergence tag.

---

## 7. Scope / sequencing

**Spike builds FIRST (the foundation — this service, no screens):**
1. `INV_STOCK_LEDGER` table + `POST_StockMovement` proc (idempotent, atomic, purge-aware) **in the
   sandbox DB**.
2. `stockLedger.post()` Project-Library service (string→id resolution, event-key assembly,
   `createSProcCall`).
3. The §3 source→post mapping encoded as the *derivation* used by the harness (one function, reused by
   both the live service and the parity check).
4. **The reconciliation-vs-`IN_QTY` parity harness (§6)** against the restored `Inventory.bak`, with the
   **four** EXPECTED-DIVERGENT classes asserted — **D8(3)** arrival-reversal, **D12#3** plant/assembler-yard,
   **F3** multi-row under-count, **F5** part-change — plus reject/shipping deletes asserted EXPECTED-ZERO (F2).
   (D11#9 is retired; D8/Bug2 is timestamp-only, not a qty class.) **This is the spike's GO/NO-GO.**

**Deferred (later, per-module jobs that CALL this service):**
- Wiring the actual **RecConfStat / Shipping / Reject / Stocktaking screens** to post through
  `stockLedger.post()` (each module's Stage-2/Stage-3 in its own spec). During parallel run those screens
  still drive the legacy procs+triggers; the ledger shadows and is proven equal first.
- The **D1 `site_id` FKs**, real PKs/FKs on the source tables, and the timestamp normalization — all
  land in the **Postgres phase** (`# IG83-TODO`), tracked in the D13 DB-conversion script.
- The **computed-view projection** alternative to the materialized `IN_QTY` (§1 open question) — Postgres phase.
- A **compute-on-read fallback / rebuild command** (`stockLedger.rebuildBalance(partId)` = re-SUM the
  ledger and re-stamp `IN_QTY = <absolute>`) for healing a drift the harness finds — cheap to add, build
  alongside #4. **F4 invariant:** unlike `post()`, this is a read-then-write absolute stamp and MUST take an
  exclusive/serializable lock on the part row (`UPDLOCK, HOLDLOCK` before the SUM) or run only in a quiesced
  window — otherwise a concurrent in-flight `post()` delta is lost (§4 point 6). State this in the proc body.

---

## 8. Assumptions & open questions (flag, don't guess)

- **(Q1 — materialized vs computed on-hand)** Recommend materialized `IN_QTY` now (legacy readers need
  the column; additive op = legacy-faithful), computed-view projection deferred to Postgres. §1. **Genuine
  fork — confirm.**
- **(Q2 — purge scope) ✅ RESOLVED (F2, verified live).** The recreject §8.2 open questions (Q2/Q4) — *does
  purge delete reject/shipping rows?* — are now **answered NO** against the live DB. `DELETE_AutoPurge`
  deletes **only** `INV_OPEN_ORDER_INF` + `INV_OPEN_ORDER_INF_HIST` + `INV_PARTS_STOCK_MST_HIST`; it never
  touches `INV_REJECT_INF` or `INV_PART_SHIPPING_INF`. Consequently: purge avoids draining on-hand via the
  `VC_TERMINATED=''` pre-stamp (not a `PurgeMode` flag — which doesn't exist, D9), and the old D11#9
  "reject-delete-during-purge" concern is **unreachable** and retired (§3.1, §6). Reject/shipping deletes are
  always genuine user deletes that restore on-hand.
- **(Q3 — UPDATE as one net-delta row vs two)** Chose one net-delta row per edit (§4.3). Equivalent in
  on-hand effect to the legacy two-statement triggers; differs in *ledger row count* vs a literal
  `INV_PART_QTY_INF` replay. If parity must match the legacy audit row-for-row (not just the balance),
  switch to two rows. Recommend net-delta (balance parity is what matters). **Confirm the parity grain.**
- **(Q4 — over-deduction / negative on-hand)** Legacy triggers do raw subtraction with no floor (a reject
  or ship > on-hand drives `IN_QTY` negative — recreject §8.4). The additive ledger preserves this
  (negative balances are representable and parity-faithful). Whether to *guard* against negative on-hand is
  a **domain policy** beyond this service — recommend the ledger stays faithful (allow negative) and any
  guard is a UI/business-rule layer. **Confirm.**
- **(Q5 — `VC_SOURCE_EVENT` grain for RENBAN batch)** The RecConfStat RENBAN batch update mutates *every*
  order in a renban group in one shot (recconfstat §4), potentially a multi-row qty re-balance. The service
  must post **one movement per affected (part, order)** with distinct event keys — confirm the batch driver
  iterates rows (it should, given per-row idempotency). Multi-row-safety is explicit here and is the *exact*
  hazard the rebuild fixes: the legacy `FROM PS, inserted i` / `JOIN inserted` triggers under-counted on
  multi-same-part batches (F3, §6 class 5 — verified live in `InsertPartShipping`, `INSERT_RejectParts`, and
  the RecConfStat legs). The per-row post makes the rebuild count all rows; the harness predicts the divergence.
- **(Assumption)** `INV_PART_QTY_INF`'s 7546 historical rows are treated as legacy audit only, NOT
  migrated as truth; the rebuild ledger truth is reconstructed by replaying live source tables (§2/§6).

---

## 9. Cutover / DB-conversion notes (feed the D13 conversion script)

- **NEW table `INV_STOCK_LEDGER`** + proc `POST_StockMovement` (+ UNIQUE `(IN_PART_ID, VC_SOURCE_EVENT)`).
- **Retire the 12 qty-triggers** at cutover (they're replaced by `stockLedger.post()` calls from the
  rebuilt screens). Until cutover they stay live on the legacy DB; the rebuild ledger shadows them.
- **`site_id` FK** on `INV_STOCK_LEDGER` + per-site scoping (D1, Postgres phase).
- **Timestamp normalization** `TS_POSTED`/`VC_ADD` 16-char string → `datetime2` (Postgres phase, every
  `# IG83-TODO`).
- **Backfill:** ⚠️ **SUPERSEDED FOR THE CUTOVER WRITER (David decided — see `cutover-runbook.md` §4,
  BLOCKER 2 RESOLVED).** The cutover backfill is the **opening-balance copy**: `SEED_AllOpeningBalances`
  records each part's CURRENT legacy `IN_QTY` as one `OPENING_BALANCE` ledger row (so `SUM(ledger)==IN_QTY`
  by construction), in a quiesced window after the trigger drop. David's "start clean, go forward": a
  from-zero / per-source-row replay is **impossible** because `DELETE_AutoPurge` aged out the receiving
  IN-moves, so a replay reconstructs a WRONG (missing-receipts) opening. **The "derive one movement per
  live source row" derivation below survives ONLY as the read-only §6 parity-harness derivation — it is
  NEVER the cutover seed/writer.** Do not redesign this; the opening-balance copy is the decided path.
- **Bug fixes "at cutover" — SUPERSEDED as *balance corrections*; they are FORWARD-only divergences.**
  ⚠️ The original claim ("the backfill will produce the *corrected* on-hand for D8(3)/D12#3/F3/F5;
  reconcile the delta explicitly") is **dead for the cutover** — the opening-balance copy FREEZES the
  legacy (buggy) `IN_QTY` into the seed; it does NOT correct historical on-hand, and there is **no
  delta-reconciliation step** (the copy is exact by construction). Only **forward** posts (via the seams)
  get the corrected behavior, where the seams intentionally diverge from the old triggers. The four classes
  (D8(3) arrival-reversal, D12#3 plant/assembler-yard, F3 multi-row under-count, F5 part-change) remain
  meaningful ONLY as the §6 forward-behavior parity tags and the read-only harness derivation — they are
  **retired from cutover scope as balance corrections.** **Honest caveat (`feedback-parity-fixture-fidelity`):**
  `test_ledger_opening_balance.py`'s "0 drift after backfill" check is **vacuously green by construction**
  (it copies `IN_QTY` in) — it proves the seed **MECHANISM**, NOT balance correctness; never cite it as a
  correctness/parity sign-off.
```
