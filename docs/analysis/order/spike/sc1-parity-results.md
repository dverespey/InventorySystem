# SC1 Number-Parity Results — SIM_OrderSimulation vs golden Delphi exports

Run: 2026-06-14. Anchor **Today = 2026-06-15**, **Line = COROLLA**, **FillDays = 25**.
Harness: `docs/analysis/order/spike/parity_diff.py` (re-runnable). Golden:
`DB Schema/OrderSimulationCorolla{Tire,Wheel,Valve,Film}.xls` → `/tmp/golden/*.xlsx`
(LibreOffice headless; gitignored real client data — never commit).

Compares the proc's **per-part** phased output (Section B grid rows + Section C phased
cells) against the golden **4-row ledger** (pooled Beg / per-supplier Receipts / pooled
Usage / pooled End). Golden day column = `20 + fill_pos` (production days; weekends are
already absent as columns). The golden pools Beg/Usage/End at the size-group level, so the
pooled rows are reconstructed as the **SUM of per-part values** by size group — verified
that this reconciles exactly (e.g. 18DL day0: DUNLOP 24491 + MICHELIN 23394 = 47885 = golden
End T41).

## Calendar (requirement #5) — RESOLVED + finding

Retired the fictional fixture (6/17 H, 6/18 X, 6/23 O, 6/25 H, 7/3 O). The golden's own
date-headers are the authoritative production calendar. **Finding: the real calendar is NOT
weekend-only** — within the 6/15→7/27 window it also skips:

- **7/3 (Fri)** — July 4th (Sat) observed holiday.
- **7/13–7/17 (Mon–Fri)** — a full mid-July **shutdown week**.

Encoded both as `H` rows in `SIM_SpecialDate_Fixture` (representing what `AD_GetSpecialDate`
returns). This took End/Usage parity from 0/20 → 14/20 (the rest are the data/proc gaps
below). The original fictional fixture rows are backed up at `/tmp/fixture_backup.sql`.

## Summary (updated 2026-06-15 — R1 + R3 CLOSED)

| Channel | Pass/Total | Status |
|---|---|---|
| Order-by date placement | **22/22** | ✅ exact (req #3 confirmed: order-by = Today + P production days; P = `IN_LEADTIME`-by-weekday, single combined field, NO separate logistics add) |
| Receipts (per supplier) | **21/22** | ✅ R3 closed; the 1 miss is 17D1 = R2 (manual demo edit, excluded) |
| Usage (pooled) | **20/20** | ✅ R1 closed (FILM ×4 now pass) |
| End Balance (pooled) | **19/20** | ✅ R1+R3 closed; the 1 miss is 17D1 = R2 (excluded) |

**19 of 20 size-groups pass cell-for-cell** (was 14). The only remaining miss is WHEEL 17D1 = R2,
David's manual row-23 demo edit, which the proc legitimately cannot reproduce (excluded from scoring) —
so this is effectively 20/20 of the reproducible groups. Passing groups: TIRE 15D, 16D, 16DL, 16G, 16H,
17DL, 18DL, SPARE; WHEEL 15D1, 16D2, 16F, 18D2, **M1**; VALVE RV, TPMSS; **FILM BLACK, BLUE, GREEN, RED**.

### R1 — CLOSED (2026-06-15). FILM forecast week-number mapping.
Legacy resolves the forecast breakdown by **ISO week-number** (year-blind), not absolute `VC_WEEK_DATE`:
`@WeekNo = DATEPART(ISO_WEEK, prodDate) − weekoffset`, `weekoffset = INT_FIRST_PRODUCTION_WEEK[year]−1`
when `@UseFirstProductionDay=1` (the golden client config; 2026 first-prod-week=2 → offset 1), with an
underflow guard. Matching by week-number lets 2026 production days find the 2025-dated FILM rows; TIRE/
WHEEL/VALVE are unaffected (their 2026 breakdown has `IN_WEEK_NUMBER = ISO(VC_WEEK_DATE)−1`, so the same
row is picked). Source: Order.pas:1145-1196, SELECT_ForecastPartNumberWeek (CreateInventory.sql:6309-6355),
guard Order.pas:1175-1178. Implemented in `SIM_OrderSimulation.sql` STEP 4; the OrderSpike view binding now
passes `@UseFirstProductionDay=1` (QA caught a hardcoded `=0` that had defeated the fix at the browser).

### R3 — CLOSED (2026-06-15) as a FIXTURE bug, NOT a proc-math change.
The legacy receipt projection (`SELECT_OrderOpenOrderList`/`PutOpenOrderCount`) **sums ALL rows by
`VC_FRS_DATE`, no renban filter** — so STEP 5's sum-all-rows was already faithful. The overcount came from
`spike-fixtures.sql` injecting 8 synthetic blank-renban "SPIKEFX" rows for M1 (4261102Q8000) ON TOP of the
real 855 renban-grouped CMWA prod rows. In prod, M1 is palletized (BIT_LOT_SIZE_ORDERS=1 = inverted flag =
lot-sized FALSE), renban-grouped, and its orders are created with a placeholder renban that the RenbanOrder
breakdown form overwrites (DELETE_OrderRenban removes placeholders → no blank-renban rows ever reach Order
Start). Fix = delete the 8 SPIKEFX rows so the spike reflects the post-breakdown state; the CMWA rows alone
sum to the golden [440,880,880,880,400,0]. No proc change. See [[project-order-renban-domain]] memory.

## Residual failures — all root-caused

### NOTED FOR FUTURE (proc-fidelity gaps, not view-rebuild scope)

**R1 — FILM ×4 (BLACK/BLUE/GREEN/RED): forecast week-number mapping.**
The spike DB's FILM forecast breakdown (`INV_BREAKDOWN_FC_INF`) is **2025-dated**
(`VC_WEEK_DATE` 20250609…). The golden's FILM usage values (27,24,24,24,24,28…) **exactly
equal those 2025 rows' `IN_QTY` columns** — proving the legacy resolves forecast by
**week-number** (cycling the most recent available year), via the
`SELECT_FirstProductionDay` / `[INIT] UseFirstProductionDay` offset. `SIM_OrderSimulation`
STEP 4 currently matches by **absolute Monday `VC_WEEK_DATE`** (proc comment line 42 assumed
"map by VC_WEEK_DATE so offset is implicit") and finds no 2026 row → usage 0 → flat balance.
Tire/wheel/valve groups pass because their breakdown IS 2026-dated.
- **Fix owner:** delphi-architect to confirm the exact legacy week-number lookup (is it
  `IN_WEEK_NUMBER`, ISO-week mod, or a `SELECT_FirstProductionDay` offset?), then
  ignition-developer to revise STEP 4. This is a fidelity fix (faithful), not a calc change.

**R3 — WHEEL M1 (4261102Q8000): receipt overcount.**
Multiple open-order rows share one `VC_FRS_DATE` (e.g. 6/15 has 440 **and** 500). The proc
sums all of them per day (day0 = 940); the golden shows only 440. The proc's @receipts pulls
both the in-transit (`ssup<>''`) and open-order (`ssup IS NULL/''`) lists with the
`@FirstFRS` filter; the legacy `SELECT_OrderInTransitList` / `SELECT_OrderOpenOrderList` must
apply a tighter status/dedup filter than the proc reproduces.
- **Fix owner:** delphi-architect to confirm the exact two list-proc filters; ignition-developer
  to align @receipts.

### EXPECTED DIVERGENCE (not a bug)

**R2 — WHEEL 17D1 (4261102P8000): day5 receipt gold=25, proc=0.**
This is David's **manual "row-23" demo edit** — he typed 25 into the editable order-by cell
in the live sheet before exporting. The proc legitimately produces 0 (no such open order
exists). Confirms the simulate→adjust mechanic in the golden; **exclude from parity scoring.**

## Verdict

SC1 is **PASS for the faithful calc surface** the view rebuild renders (calendar, order-by
placement, ledger recurrence, share/pooling, receipts) on tire/wheel/valve. The two open
proc-fidelity gaps (R1 FILM forecast-week, R3 M1 receipt-filter) are **tracked for a
follow-up proc pass** and do not block the faithful-layout view rebuild (the view renders
whatever the proc returns; these affect FILM + one wheel group's numbers, not the layout).
