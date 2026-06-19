# Post-Cutover Enhancements — deferred improvements (tracked, NOT in the production build)

**Status:** 🟡 living backlog. **Decision (David 2026-06-19):** the production rebuild is built **faithful
to the legacy** (Option A) — *"build as legacy, don't confuse the users any more than we have to."* The
items here are deliberately deferred to **after** the cutover, each to be evaluated on its own merits with
the operator and validated against real outcomes before adoption. Nothing here ships in M1–M5.

Why deferred: each changes the numbers/behavior the operator sees, which (a) would break parallel-run /
dev-mirror parity (the cutover validation gate, Q16) and (b) adds risk to an already-large rebuild. Keep
the rebuild a faithful reproduction first; optimize second.

---

## Order calculation modernizations (Option-B C2–C6)

Source: `docs/analysis/order/option-b.md` (full citations) + `order-redesign-plan.md`. The Order spike was
built **faithful (Option A)**; SC1 parity = 19/20. **C1 (silent ≤200-row truncation → remove cap + warn) is
a DEFECT FIX and is included in the faithful build** (no order-math change) — it is NOT deferred.

| # | Legacy behavior | Proposed change | Status |
|---|---|---|---|
| **C2** | Lead time selected by **today's** weekday column (`IN_LEADTIME_MONDAY..SATURDAY`), fallback `IN_LEADTIME` (`Order.pas:426-459`) | Select by the **order-by (release) day's** weekday, offset against the working calendar so release lands on a valid working day | DEFERRED — keep-proposed (cited: Oliver Wight, Oracle JDE) |
| **C3** | Added-leadtime: each overtime day in the window pushes order-by +1, `break` on first miss (`Order.pas:1576-1582`) | True working-calendar offset that skips non-production days + absorbs overtime/extra-shift days; replace the break-loop | DEFERRED — keep-proposed (cited: Oracle shop-floor calendar) |
| **C4** | End-balance < fixed `J = usage×days` → red | Running Projected Available Balance vs **parameterized statistical safety stock** (`SS = Z×σ`, per-group Z; King combined formula where demand AND lead time vary) | DEFERRED — keep-proposed (cited: King/MIT, Oracle) |
| **C5** | Order share `= E/(ΣE)` across the size group, 100% singleton | **Net-requirements** as the trigger (gross req − scheduled receipts − projected available − in-transit per period); retain `E/(ΣE)` as downstream allocation; `ROP = LT-demand + SS` fallback | DEFERRED — keep-proposed/revised (cited: Infor TPOP, Netstock) |
| **C6** | Forecast via week/day breakdown table + first-production-day offset | Bucket forecast onto production-calendar working days (one bucket/working day); split firm near-term (862/DELJIT) vs far-horizon forecast (830/DELFOR) | DEFERRED — keep-proposed (cited: Oracle, Orderful) |

**Adoption path (when revisited):** pick one C at a time → quantify the delta vs the faithful baseline on
real data (the dev-mirror harness, Q16) → review the changed recommendations with the operator → sign off →
ship behind a per-site toggle so it can be rolled back. Do not batch-adopt.

---

## Other deferred-to-later items (parked here so they aren't lost)

- **Finer per-feature roles** (Q12) — the Admin/User split ships now; an EDI-only / receiving-only
  permission model is a later refinement if operators specialize.
- **Postgres phase (D13 / `# IG83-TODO`)** — string-timestamp → `datetime2`, typed dates in
  `fn_ManifestCostAt`, real PK/FKs, etc. Tracked in `cutover-runbook.md` §8.

Add future deferred enhancements here as they surface, with a one-line rationale for why they wait.
