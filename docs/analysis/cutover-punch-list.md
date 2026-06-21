# Cutover Punch-List

The single tracked home for deferred-but-not-dropped items that must be resolved (or consciously accepted)
before the production cutover. Items are deferred on purpose — they are architect / cutover concerns, not
build blockers — but each one needs an owner-review so none silently rots. Created 2026-06-21 (team retro
R13/R19) to give the accumulating carries a home.

- **Owner:** David.
- **Review cadence:** at each milestone boundary (a milestone closes → walk this list, resolve / re-defer /
  escalate each OPEN item before opening the next milestone).
- **Status legend:** OPEN (carried), DECIDED (resolution recorded — keep until applied), DONE (applied to
  prod / closed).
- Each entry: **source unit** · one-line description · status. Detail lives in the cited per-unit doc /
  decision ledger; this is the index, not the spec.

| # | Source unit | Item | Status |
|---|---|---|---|
| P1 | EDI 862 (release) | 862-firm decision: report-only vs persist the firm releases — behavior not yet chosen. | OPEN |
| P2 | M1 status render | Status-render E/P arms: the inbound 997 AK9 E/P (accept-with-errors / partial) render blank — E and P arms unhandled in the status mapping. | OPEN |
| P3 | EDI 856 / GetShip | GetShip calendar inconsistency: ship-date skips only weekends + H, NOT O/X/W, unlike the forecast day-spread which honors O/X/W — two calendars disagree. | OPEN |
| P4 | Renban order | Renban rollover → collision: the renban counter is varchar(3) and rolls over / collides at 999. | OPEN |
| P5 | Renban order | All-lots-0 keep-safer divergence: rebuild keeps the safer (non-blanking) behavior when all lots are 0 — a decide-and-flag divergence, recorded in the renban divergence ledger (correctly decided, not asked). | OPEN |
| P6 | Forecast (.frc) | The 2nd unfinished SiteSupplierCode crash on the .frc forecast-supplier-feed path — the supplier-feed branch crashes on an incomplete SiteSupplierCode. | OPEN |
| P7 | Golden EDI / .ord | Golden-EDI/.ord confirmations to lock against prod: TEMA byte-parity; real INV_SITES DUNS/EIN/separator values; the delSL[4] inbound element index; the IT104 money scale; the .ord full-width formatting. | OPEN |
| P8 | M4 (multi-site) | site_id = M4 carries (the `-- M4` markers): ASN/INV status UPDATEs, the .ord leading SiteSupplierCode, order/logistics dirs, and UPDATE_EINStatus site-scoping — all deferred to the M4 multi-site milestone. | OPEN |
| P9 | Cutover (manifest cost) | Verify the `IX_INV_MANIFEST_COST_MST` constraint-DROP on PROD (the one residual constraint op from the cutover dress-rehearsal). | OPEN |
| P10 | Order file | Rollback-of-rollback message NIT: the order-file rollback path emits a confusing rollback-of-rollback message — cosmetic, low-risk. | OPEN |
| P11 | Test hygiene (R18) | Fixtures self-healing on a killed run: sentinel-prefix synthetic rows + pre-clean by sentinel at suite start (forecast left synthetic INV_SUPPLIER_MST rows hitting IX_INV_SUPPLIER_MST UNIQUE; S:\CMX path-leak self-cleaned). Routed to ignition-qa. | OPEN |

## Notes

- P5 and P3 are paired with EDI/renban divergence ledgers (renban; `edi810-decisions.md` style D-810-1..5);
  P5 is a *decided* divergence (safer/no-op → decide-and-flag), retained here so it isn't lost at cutover.
- P7's golden-EDI confirmations are byte/value-exact checks that can only be finalized against real prod
  INV_SITES rows — flag, don't fabricate.
- P8 (site_id / M4) is intentionally a whole-milestone bucket; expect it to expand as M4 is scoped, then
  drain into M4 build items.
