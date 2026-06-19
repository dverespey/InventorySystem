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

## 3. D6 report-proc cutover apply
`docs/analysis/reporting/spike-report-procs-d6.sql` is the CUTOVER artifact — it REPLACES the legacy
window-blind `REPORT_INVOICESSummary / MonthlyINVOICESSummary / EDI810 / EDI856`. Not applied to the spike
(legacy kept as the parity baseline). **At cutover:** apply it; confirm the manifest unique-index is
dropped (done in spike) and add the **DB-level non-overlapping-window constraint** as the backstop to the
app guard (§4). `SELECT_ManifestCost` (UI lookup) is SUPERSEDED by `fn_ManifestCostAt` — repoint the UI.

## 4. Opening-balance backfill  (the parity closure)
Run `stockLedger.seedAllOpeningBalances()` (set-based `SEED_AllOpeningBalances`) ONCE at cutover, **before
any forward post** (the genesis guard THROWs otherwise). After it, `SUM(ledger) == IN_QTY` for every part;
forward posts keep it. The from-zero full reconstruction is impossible (receiving history purged) — this
opening balance + forward parity IS the sign-off. (A real pre-purge dump, if one ever surfaces, would let
`test_stock_ledger_parity.py`'s skipped from-zero check run — but it's not required.)

## 5. Retire the 12 legacy qty-triggers
At cutover, DROP the receiving/shipping/reject/stocktaking qty-triggers — the producer seams replace them.
The seams are CUTOVER paths that presume the triggers are gone (else IN_QTY double-counts). Sequence the
drop with the read-flip.

## 6. Per-edit version-stamp contracts (idempotency keys)
- **Shipping:** `AmendShipment` must re-stamp `VC_ADD` per edit (INV_PART_SHIPPING_INF has no VC_LAST_UPDATE;
  VC_ADD is the amend version token).
- **Stocktaking:** the edit path must write a NON-NULL `VC_LAST_UPDATE` (fixes the legacy Bug2 NULL).
- All producers' amend keys use `:upd:to=<effect>:v=<stamp>` (collision-safe); keep that on any new write path.

## 7. Headless Jython driver coverage  (retro R8 — partially done)
`scripts/e2e/jython_shim.py` + `test_seam_driver.py` run the REAL stocktaking + reject + **shipping +
receiving** wrappers end-to-end (autocommit shim) — including shipping's header-delete cascade and
receiving's `resolveAddPoint` add-point gate (asserted with a NON-zero IN_QTY delta on a counted 'S'
order). **Remaining extend:** **Order needs the shim's persistent-sqlcmd-session extension** (its
`beginTransaction` spans statements). A true in-gateway runtime test (Perspective/Playwright is
trial-gated; no WebDev; no gateway-event infra) remains a gap — the shim covers the driver LOGIC/SQL, not
Jython-2.7-vs-CPython runtime quirks. (Note: extending to shipping/receiving surfaced + fixed a narrow
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
