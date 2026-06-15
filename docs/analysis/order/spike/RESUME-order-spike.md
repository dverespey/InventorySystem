# Order Spike — RESUME HERE (checkpoint 2026-06-15)

Single source of truth to continue the **Order** ("what to order") keystone rebuild without re-deriving.
Target stack = **Ignition** (Option A "Faithful-Modern"; Rails retired). Read with: `order-redesign-plan.md`
(decision + spike plan), `legacy-order-spec.md` (source truth), `source-artifacts.md` (incl. §5 CF/palette
gap closure), `option-a.md` (design), `build-spec.md` (proc contract), `sc1-parity-results.md` (parity run),
`../../automated-ui-testing.md` (E2E harness approach).

## STATUS: faithful-layout REBUILD DONE + validated live; SC1 parity run; E2E automation stood up

The faithful 4-row ledger rebuild is built, reviewed, QA'd, and **confirmed rendering+interacting on the
live gateway by automated headless test (12/12)**. The spike's foundation is unchanged and proven:
- **`SIM_OrderSimulation`** T-SQL proc (spike DB `Inventory`) reproduces the legacy day-by-day sim
  (PAB recurrence, weekday lead-time, added-leadtime break-loop, hazard-7 two-index reconciliation,
  `=E/ΣE` share). DDL: `docs/analysis/order/spike/SIM_OrderSimulation.sql`.
- **`Order/OrderSpike`** Perspective view — **REBUILT** to the faithful pooled 4-row ledger
  (Beg / Receipts-per-supplier / Usage / End + gray separator), numbers+color-only (peach editable
  order-by cell, red below-safety End), glyph clutter dropped, order entry locked to each supplier's
  order-by cell, client-side live End recompute (row-23 mechanic). `domId`s added + loaded.
- **SC verdicts:** SC1 = **14/20 size-groups cell-for-cell PASS, order-by 22/22** (see
  `sc1-parity-results.md`; 6 misses are the tracked R1/R2/R3 below); SC2 = PASS (palette/CF known);
  SC3 = PASS, live-confirmed (`ia.display.table`, no Flex-Repeater).

### This session (2026-06-15) — what got done
1. **Req #3 resolved from source:** "Lead (P)" is a single stored field (`IN_LEADTIME` by weekday); legacy
   does NOT add a separate logistics lead time. Order-by = Today + P **production** days. Golden-exact.
2. **Req #5 calendar:** retired the fictional fixture; encoded the REAL calendar from the golden's day
   headers — skips weekends **+ 7/3 (July-4 observed) + 7/13–7/17 (mid-July shutdown week)**. In
   `SIM_SpecialDate_Fixture` as `H` rows. (Original fixture backed up at `/tmp/fixture_backup.sql`.)
3. **SC1 parity harness built:** `parity_diff.py` (re-runnable) diffs proc vs all 4 golden sheets →
   `sc1-parity-results.md`. 14/20 groups + 22/22 order-by.
4. **View rebuilt** via the fleet (ignition-developer → ignition-code-reviewer → ignition-qa). Reviewer's
   two open runtime RISKs (transform-never-ran; `onEditCellCommit` qualified-value access) are now
   **RESOLVED** — the E2E run exercised both live (`SPIKE grid:` + `edit ACCEPTED: 4265202R6000|5`).
5. **E2E automation stood up:** Playwright (Python) headless harness `scripts/e2e/` — auto-resets the 2h
   trial, opens the view, asserts numbers/labels/no-glyphs/peach/edit, greps `SPIKE`, screenshots. **12/12
   PASS, zero human clicks.** Wired into the `ignition-qa` agent as the standard browser-test track.
   Gateway-restart now permitted (`.claude/settings.json`). Creds in gitignored `scripts/e2e/.env`.

## GOLDEN EXPORTS RECEIVED (the big new input)

David exported the live Delphi Order→Start sheets: `DB Schema/OrderSimulationCorolla{Tire,Wheel,Valve,Film}.xls`
(Today=2026-06-15, Line=COROLLA). **These are real client data — gitignored (`DB Schema/OrderSimulation*.xls`),
never commit.** `xlrd` can't read `.xls`; convert first:
`"/Applications/LibreOffice.app/Contents/MacOS/soffice" --headless --convert-to xlsx --outdir /tmp/golden "DB Schema/OrderSimulationCorolla*.xls"`
then read with `openpyxl` (LibreOffice is installed). (`/tmp/golden/*.xlsx` from this session is temp — re-convert.)

### What the golden revealed (decoded, authoritative)
1. **Real layout = a 4-ROW LEDGER per part/group, numbers only** (this is what my spike got wrong — I
   crammed it into 1 glyph-heavy row). Per block, label in col S:
   - **Beg Balance** → **Receipts** → **Usage** → **End Balance** → gray separator (`#969696`).
   - `End Balance = Beg + Receipts − Usage`; carried forward (next day Beg = prior End). PAB.
2. **Shared size group = pooled Beg/Usage/End + ONE Receipts row PER SUPPLIER**, each with its OWN
   editable order-by cell on a DIFFERENT date. Tire 18DL example (rows 37-41):
   - DUNLOP `4265202S1000` lead **5** → editable cell **Y = 06-22**
   - MICHELIN `4265202S2000` lead **6** → editable cell **Z = 06-23**
   - SPARE group: YOKOHAMA lead5→06-22, MAXXIS lead7→06-24.
   - **Order-by column = Today + lead-time counted in PRODUCTION days** (weekends skipped).
3. **Colors (real, minimal):** peach `#FFCC99` = the editable order/receipt cell; **red font** (CF) =
   End Balance `< safety stock` (`$J$` of block start; safety `J = H dailyUsage × I days`); green font
   `#008000` = receipts sourced from open orders; gray `#969696` = separator. CF rule confirmed:
   `End-Balance-row cellIs lessThan $J$<blockStart>`.
4. **Real calendar:** skips weekends ONLY; **6/17 is a normal Wed (NO holiday)** → my fixture was WRONG.
   Retire the `AD_GetSpecialDate` fixture; use the real production calendar (the golden's day headers).
5. **row-23/col-Y demo (Wheel):** David typed **25** into an editable Receipts cell (Y, 06-22); End
   Balance recomputed `Beg(−4)+25−Usage(8)=13` → red shortage cleared. = the simulate→adjust loop.

## REQUIREMENTS (David, 2026-06-14) — ALL DELIVERED ✅ (validated live 2026-06-15)
1. ✅ **Faithful clean ledger** — 4-row ledger, numbers+color-only (peach editable + red below-safety End),
   glyph clutter dropped. (E2E: "no glyph clutter" PASS.)
2. ✅ **Order entry locked to the ONE valid cell** — only each supplier's order-by cell editable; column
   `editable:true` + per-cell `editable:false` + `onEditCellCommit` reject backstop.
3. ✅ **Order date** — resolved: "Lead (P)" = single `IN_LEADTIME`-by-weekday field, NO separate logistics
   add; order-by = Today + P prod days. Golden-exact (order-by 22/22).
4. ✅ **Editable cell → live recompute** — typing a qty recomputes pooled End forward + re-evaluates red.
   (E2E: `edit ACCEPTED: 4265202R6000|5` PASS.)

## NEXT STEPS (do this on resume) — in priority order
1. **E2E scroll-to-row capability.** The harness (`scripts/e2e/test_order_spike.py`) asserts against the
   top-visible group (15D) because `ia.display.table` is **virtualized**. Add a helper that scrolls the
   grid body to bring an off-screen group into view (or use the table filter) so deep groups can be
   asserted — target **18DL** (pooled End day0 must = 47,885; DUNLOP peach @06-22 / fill_pos 5, MICHELIN
   @06-23 / fill_pos 6) and the **SPARE** group. Then extend `test_order_spike.py` with per-group value +
   peach assertions for ≥2 deep groups. (Prefer the loaded `domId` `#spike-order-grid` as the scroll root.)
2. **Close the 2 proc-fidelity gaps (R1, R3 below)** to push SC1 from 14/20 → 20/20: delphi-architect
   confirms the exact legacy rules → ignition-developer revises `SIM_OrderSimulation` → re-run
   `parity_diff.py` + the E2E harness. R1 (FILM forecast week-number mapping) is the bigger one.
3. **Commit path** (still DEFERRED — see below): only after parity is clean. SERIALIZABLE + UPDLOCK +
   commit-claim; needs delphi-architect sign-off on single-writer assumption first.
4. **Export `Order/OrderSpike` view + the NQ SQL into git** (currently gateway-only) once stable.

## NOTED FOR FUTURE — proc-fidelity gaps surfaced by SC1 parity (2026-06-14)
SC1 = **14/20 size-groups pass cell-for-cell**, order-by **22/22** (see `sc1-parity-results.md`).
Two genuine `SIM_OrderSimulation` gaps tracked for a follow-up proc pass (NOT view-rebuild scope;
the view renders whatever the proc returns):
- **R1 — FILM forecast week-number mapping.** Proc STEP 4 matches forecast by absolute Monday
  `VC_WEEK_DATE`; legacy cycles prior-year forecast by week-number (`SELECT_FirstProductionDay` /
  `UseFirstProductionDay` offset). Spike FILM breakdown is 2025-dated → proc returns usage 0 (flat
  balance). Golden FILM usage = those 2025 rows' IN_QTY, confirming week-number lookup. Owner:
  delphi-architect (confirm exact rule) → ignition-developer (revise STEP 4). Faithful fix, not a calc change.
- **R3 — WHEEL M1 receipt overcount.** Multiple open-order rows share one `VC_FRS_DATE`; proc sums
  all, golden shows the legacy `SELECT_OrderInTransitList`/`SELECT_OrderOpenOrderList` filtered set.
  Owner: delphi-architect (exact filters) → ignition-developer (align @receipts).
- (R2 — 17D1 day5=25 is David's manual row-23 demo edit, not a bug; excluded from scoring.)
- Calendar finding: the real production calendar skips **7/3** (July-4 observed) and **7/13–7/17**
  (mid-July shutdown week) in addition to weekends — encoded in `SIM_SpecialDate_Fixture` as `H` rows.

## DEFERRED (later phases, not now)
- **Commit path**: the `INSERT_OpenOrder` / renban read-then-write race needs SERIALIZABLE + UPDLOCK +
  commit-claim; plus a cross-DB `Activity`/`TireOrder` trigger error seen during seeding. Out of spike scope.
- **`AD_GetSpecialDate`** real body/status-domain (real calendar now derivable from golden for parity).
- Export `Order/OrderSpike` view + the NQ SQL into git (currently gateway-only).
- Option B calc changes C1–C6 (all deferred behind David sign-off; faithful calc ships).

## ENV / KEY FACTS (verified)
- Gateway: Ignition **8.1.52**, `/usr/local/ignition`, `:8088`. Project `spike`. Resources load on `gwcmd -r`.
  Logs: `/usr/local/ignition/logs/wrapper.log` (grep `SPIKE` for the view's diagnostics; grep "Unable to
  deserialize" after reload).
- DB: connection NAME `Inventory_Spike` (what Named Queries reference) → database `Inventory` → login
  `ignition_spike`; sa pass = `$SA_PASS` (throwaway dev-container cred — see local `scripts/spike-db.sh`,
  not committed; `export SA_PASS=...` before the sqlcmd calls). Container `mssql-spike` (Colima/docker, localhost:1433).
  `docker exec mssql-spike /opt/mssql-tools18/bin/sqlcmd -C -S localhost -U sa -P "$SA_PASS" -d Inventory -Q "..."`.
- Sample parts (COROLLA): VALVE `900804500600`(RV) + `426070E09000`(TPMS); WHEEL `4261102Q8000`;
  TIRE 18DL `4265202S1000`(DUNLOP)/`4265202S2000`(MICHELIN); FILM (see golden).
- LibreOffice installed: `/Applications/LibreOffice.app/Contents/MacOS/soffice`.
- Agent fleet (`~/.claude/agents/`): delphi-architect → ignition-architect → adversarial-architect-reviewer
  → ignition-developer → ignition-code-reviewer → ignition-qa. All Delphi→Ignition-agnostic.
- SQL files are UTF-16LE → `iconv -f UTF-16LE -t UTF-8` before grep.
