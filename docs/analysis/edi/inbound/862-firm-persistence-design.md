# EDI 862 (DELJIT firm) — persistence design (deferred; pivot-ready)

**Decision (David 2026-06-22, punch-list P1): report-only for now** — the rebuild matches the legacy: the
inbound 862 DELJIT "firm" release is parsed into a side report and touches NO forecast/order table. The 830
DELFOR planning forecast alone drives `INV_BREAKDOWN_FC_INF` → the order.

This note exists so we can **pivot to persisting the firm releases later without rework**. It captures what
persistence would look like, so today's report-only build is forward-compatible.

## Why persist later (the value)
The 862 is the *firm/JIT release* — what TEMA actually commits to taking, vs the 830's planning forecast.
Today the operator eyeballs the 862 report. A future enhancement: persist the firm qtys so the system can
(a) flag planning-vs-firm variance, (b) drive the order off the firm release when present (more accurate than
the planning forecast), and (c) alarm when a firm release diverges from what was planned/shipped.

## What persistence would look like (the pivot)
1. **A firm-release store** — a new `INV_BREAKDOWN_FIRM_INF` (mirror `INV_BREAKDOWN_FC_INF`'s shape:
   supplier/part/week/day-qtys + `site_id` from day one) OR a `BIT_IS_FIRM` discriminator + a firm-qty column
   on the existing breakdown. A separate table is cleaner (keeps the planning forecast untouched; no risk to
   the 830 path) and is the recommended pivot.
2. **The 862 parser writes it** — the inbound processor's 862 branch (today report-only) gains a
   `persist_862(refs, site)` that upserts the firm qtys into the firm store (same delete-then-accumulate /
   per-component-supplier discipline as the 830 importer, so the two stay consistent). Reuse the M2 forecast
   importer's proven seam (the explode, the day-spread, the per-component supplier from `SELECT_PartsStockInfo`,
   the DELETE_ForecastInfo-then-additive pattern).
3. **The order optionally reads firm-over-planning** — a per-site config flag (`BIT_USE_FIRM_FOR_ORDER` on
   `INV_SITES`) so a site can opt into ordering off the firm release when a firm row exists for the week,
   falling back to the planning forecast otherwise. Default OFF (= today's behavior).
4. **Variance alarm** — reuse the M1 alarm table (`INV_EDI_ALARM_REJ` / a `FORECAST_FIRM_VARIANCE` type) +
   the home-hub surface to flag planning-vs-firm gaps over a threshold.

## What today's report-only build must NOT do (to stay pivot-ready)
- Do not delete/transform the raw 862 beyond the report (keep the parsed `{manifest, part, qty, proddate}`
  available — the inbound `parse_862` already yields a clean structured result the future `persist_862` can
  consume).
- Keep the 862 branch in the inbound processor (don't fold it into the 830 path) so the firm store is an
  additive, isolated change.
- Carry `site_id` everywhere the firm store would key (the M4 site-scoping makes the firm store multi-site
  from the start when it's built).

## Effort to pivot (estimate)
Small-to-medium: one new table + a `persist_862` driver (mirrors the M2 forecast importer) + an order
read-flag + an alarm type. No change to the 830 path or the order math (firm is additive/opt-in). Build it
when the planning-vs-firm accuracy is worth it; the report-only path serves until then.

Related: `830-862-forecast-import-spec.md` (the 862 source-truth: report-only, no DB writes), the M2 forecast
importer (`project-library/forecast/`), the M1 inbound `parse_862`.
