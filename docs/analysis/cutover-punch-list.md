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
| P1 | EDI 862 (release) | **DECIDED (David 2026-06-22): report-only for now** (matches legacy — the 862 DELJIT firm stays a side report, touches no forecast/order table). Persistence design captured in `edi/inbound/862-firm-persistence-design.md` so we can pivot to persisting firm releases (planning-vs-firm) later WITHOUT rework. | DECIDED |
| P2 | M1 status render | Status-render E/P arms: the inbound 997 AK9 E/P (accept-with-errors / partial) render blank — E and P arms unhandled in the status mapping. | OPEN |
| P3 | Order / GetShip | **DECIDED (David 2026-06-22): leave as-is.** This is the SUPPLIER ship-date (the order to sub-suppliers); the true supplier schedules are UNKNOWN, so the legacy guess is Mon-Fri + national holidays (`H`). It can't be made certain without real supplier calendars — the M-F+H behavior is the faithful, defensible default. The forecast day-spread honoring `O`/`X`/`W` is a different (TEMA production) calendar; the two legitimately differ. Accepted, not a defect. | DECIDED |
| P4 | Renban order | Renban rollover → collision: the renban counter is varchar(3) and rolls over / collides at 999. | OPEN |
| P5 | Renban order | All-lots-0 keep-safer divergence: rebuild keeps the safer (non-blanking) behavior when all lots are 0 — a decide-and-flag divergence, recorded in the renban divergence ledger (correctly decided, not asked). | OPEN |
| P6 | Forecast (.frc) | The 2nd unfinished SiteSupplierCode crash on the .frc forecast-supplier-feed path — the supplier-feed branch crashes on an incomplete SiteSupplierCode. | OPEN |
| P7 | Golden EDI / .ord | Golden-EDI/.ord confirmations to lock against prod: TEMA byte-parity; real INV_SITES DUNS/EIN/separator values; the delSL[4] inbound element index; the IT104 money scale; the .ord full-width formatting. | OPEN |
| P8 | M4 (multi-site) | **MOSTLY DROPPED (David 2026-06-22): sites run on SEPARATE infra (own gateway+DB each) → single-site deployments, NOT shared-DB multi-tenancy.** The `site_id` schema surgery (site_id on 32 tables, the 7-trigger lockstep, the Class-B index swaps, the EIN-collision scoping of UPDATE_EINStatus/DELETE_AutoPurge) is UNNEEDED — each DB is isolated. The `-- M4` site_id markers in the code become moot. What REMAINS in M4: (1) ✅ DONE the single-site Sites master + INV_SITES path columns (PR #39); (2) minimal auth (Admin=user add/delete; production-control does everything else); (3) hardening (INI secrets→gateway config, backup runbook, single-site DATAPURGE). See `project-multisite` memory (reversed) + `m4-site-id-surgery-map.md` (superseded banner). | RESCOPED |
| P9 | Cutover (manifest cost) | Verify the `IX_INV_MANIFEST_COST_MST` constraint-DROP on PROD (the one residual constraint op from the cutover dress-rehearsal). | OPEN |
| P10 | Order file | Rollback-of-rollback message NIT: the order-file rollback path emits a confusing rollback-of-rollback message — cosmetic, low-risk. | OPEN |
| P11 | Test hygiene (R18) | Fixtures self-healing on a killed run: sentinel-prefix synthetic rows + pre-clean by sentinel at suite start (forecast left synthetic INV_SUPPLIER_MST rows hitting IX_INV_SUPPLIER_MST UNIQUE; S:\CMX path-leak self-cleaned). Routed to ignition-qa. | OPEN |
| P12 | Hot-call (HotCallEntry.pas) | **DONE (PR #37):** the `create_hotcall_asn` driver (reuses the M1 ASN seam, one atomic tx) + the Perspective hot-call entry view (headless shell; site-from-gateway; Designer-finish flagged) are built + dual-adversary-verified. Hot-call ASNs flow through the existing 856 as `8HC…`/M390. | DONE |
| P13 | EDI 856 (hot-call + normal filename) | **RESOLVED in code (2-part fix, operational-sender-faithful) — exact `y` range golden-pending at cutover.** Re-anchored from the RECREATE button (`ASNInvoice.pas:817-825`) to the OPERATIONAL SENDER `MainMenu.ResendMarkedEDIsClick` (`MainMenu.pas:2718/2723`, the live C→S-flip path). `_filename_856` now reproduces BOTH branches: NORMAL = `856 + MMDD(copy(pd,5,4)) + LineName + .txt` (`:2718`); HOT-CALL (`seq='-1'`) = `8HC + Y+MMDD(copy(pd,4,5)) + y + LineName + .txt` (`:2723`). BOTH were wrong before: both omitted `LineName`, and the merged NORMAL (PR #29) also had the wrong date offset. `y` reproduced per-ASN as `1 + count of same-day same-line hot-calls already 'S'`. **Remaining: confirm the exact `y` range/byte against a golden `8HC…` 856 at cutover.** | CODE-FIXED / golden-pending |
| P14 | Hot-call (minors) | **DONE (PR #37):** EIN-at-send applied to the hot-call driver; header `IN_QTY` now = the detail sum (fixed the legacy stale-`@QTY` garbage). | DONE |

## Notes

- P5 and P3 are paired with EDI/renban divergence ledgers (renban; `edi810-decisions.md` style D-810-1..5);
  P5 is a *decided* divergence (safer/no-op → decide-and-flag), retained here so it isn't lost at cutover.
- P7's golden-EDI confirmations are byte/value-exact checks that can only be finalized against real prod
  INV_SITES rows — flag, don't fabricate.
- P8 (site_id / M4) is intentionally a whole-milestone bucket; expect it to expand as M4 is scoped, then
  drain into M4 build items.
