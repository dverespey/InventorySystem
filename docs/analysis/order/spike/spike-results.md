# Order spike — test results (SC1/SC2/SC3 against order-redesign-plan §5.2)

Test plan executed + non-destructive parts run against the live spike env on 2026-06-13.
Oracle = `legacy-order-spec.md` (`spec §N`) as pinned by `build-spec.md` (`bs §N`); the build
under test = `SIM_OrderSimulation.sql` (proc) + `Order/OrderSpike/view.json` (renderer).
This is a **parity** spike: it asserts the build matches the documented legacy behavior, NOT
that the Ignition platform works.

**Environment verified live before testing**
- Gateway 8.1.52 at `:8088`, state `RUNNING` (`StatusPing` = `{"state":"RUNNING"}`).
- DB sandbox `mssql-spike` (SQL Server 2019, `Inventory` DB) up; gateway connection name
  `Inventory_Spike`.
- `dbo.SIM_OrderSimulation` exists (OBJECT_ID 878626173); `SIM_SpecialDate_Fixture` = 5 rows;
  8 `SPIKEFX` open-order seed rows present, exactly 1 flipped to in-transit (the 06-18 / X-day
  bucket). All matches `spike-fixtures.sql`.
- No fixture was mutated by these tests (all reads + a read-only proc); nothing to restore.

---

## SC1 — Number parity (self-consistency) — **SPEC-PARITY PASS; VS-LIVE-LEGACY PENDING GOLDEN**

Run via `docker exec mssql-spike … sqlcmd … -d Inventory` against the proc output for the sample
parts at `@Today='2026-06-15'`. Every assertion below passed and reproduces
`proc-output-sample.md` exactly.

| # | Assertion | Result |
|---|---|---|
| 1 | **Hazard-7 day-header:** `fill_pos 2 → cal_offset 3`, `day_kind=NONPRODUCTION` (06-17 H consumed x=2, no column) | PASS |
| 1b | `fill_pos 5 = OVERTIME` (06-23); full A-set matches sample line-for-line | PASS |
| 2 | **BELOW_SAFETY ⇔ PAB<J7** on all 12 cells of case (d) VALVE/RV, J7=9220 (922×10) | PASS |
| 2b | below_safety FIRES on ≥1 cell (channel exercised, not vacuously 0) | PASS |
| 3 | **PAB recurrence** day0 = TotalInv−InTransit = 44418−440 = 43978; day0 PAB = 44164 (43978+940−754); 12 contiguous non-null balances for case (e) | PASS |
| 4 | **share_pct sums to 100%** per size group across all TIRE sizes (max deviation 0.0000); 15D singleton=100, 18DL split (2 parts)=100 | PASS |
| 5 | **Hazard-7 bucket placement:** the 06-18 in-transit FRS lands on `fill_pos=2` (matched through the calendar, not datediff) | PASS (see SC2 A5) |

**Verdict:** SC1 self-consistency + spec-algorithm reproduction = **PASS**. The PAB recurrence,
the BELOW_SAFETY threshold, share-sum, and the hazard-7 offset≠position reconciliation all hold
arithmetically against the proc output.

**VS-LIVE-LEGACY = PENDING GOLDEN.** No Delphi/Excel here, so byte-for-byte parity of the
*forecast usage values* and any *template-baked conditional-format thresholds* against the live
`OrderSimulation.xls` is unproven. The diff harness (below) is ready to run the moment David
exports a golden.

### Golden-export spec for David (exact)
Run the legacy Delphi **Order → Start** against the **same `Inventory` data the spike DB holds**
(restored from the same `.bak`, plus the `SPIKEFX` seed if testing case c/e — note the seed is a
spike fixture; for a *pure-legacy* golden, omit it and skip case c) and save the populated
`OrderSimulation.xls`, for each of:

- **Order date (Today) = 2026-06-15**, Line = **COROLLA**, and the legacy default FillDays (23).
- Part-type runs: **TIRE** (covers case a `4265202R6000` 15D, case b `4265202S1000`+`4265202S2000`
  18DL), **VALVE** (case d `900804500600` RV), **WHEEL** (case e `4261102Q8000` M1).
- For each, capture **cell values AND each cell's ColorIndex** (interior + font) — the color read
  is what closes the SC2 extraction gap (option (ii) in §6 of the redesign plan).

Caveat to flag on the golden: the spike calendar is **fixture-backed** (`AD_GetSpecialDate`
stubbed with 06-17 H / 06-18 X / 06-23 O / 06-25 H / 07-03 O). The legacy run pulls the real ALC
`TireOrder` calendar, so calendar-derived cells (day_kind, bucket placement) will only match if
the real calendar agrees with the fixture. Either (a) have David also read out the real special
dates for 2026-06-15…07-05 so the fixture can be reconciled, or (b) diff non-calendar cells first
and treat calendar cells as fixture-conditional.

### Diff harness (ready; not yet run — no golden)
The harness is the sqlcmd assertion battery above plus a cell-join: load David's golden into a
staging table `#golden(part_number, col_label, value, color_index)`, then `FULL OUTER JOIN` it to
the proc's B/C output keyed on `(part_number, fill_pos→col_label)` and emit every row where
`value` or `color_index` differs, ordered so any hazard-7 (index-space) mismatch surfaces first.
This is automatable as the **Pytest/pyodbc** track (below).

---

## SC2 — Color/signal parity — **PASS (Pascal-set signals); template-CF parity PENDING golden**

The §5 palette + threshold gap is CLOSED for the **Pascal-set signals** (build-spec §1.7 KEY
FACT). Assertions run against C-set output for case (e); each signal lands on the correct cell
and the renderer attaches ≥2 channels.

| # | Assertion (cell placement, case e WHEEL `4261102Q8000`) | Result |
|---|---|---|
| A1 | `NON_PRODUCTION` on fill_pos 2 (06-18 X day) | PASS |
| A2 | `OVERTIME` on fill_pos 5 (06-23) | PASS |
| A3 | `ORDER_BY` on fill_pos 6 (= orderby_col_index) | PASS |
| A4 | `LEAD_TIME_ZONE` on fp0/fp1 (0..leadtime_zone_end_index) | PASS |
| A5 | **font precedence + hazard-7:** fp2 has 440 in-transit + 880 open co-located → `source_enum=IN_TRANSIT` wins (font 23 over 10), bucketed by calendar match | PASS |
| A6 | day_kind override: OVERTIME at fp5 overrides LEAD_TIME_ZONE inside the lead window | PASS |
| d  | case (d) BELOW_SAFETY emitted as an orthogonal flag on the balance cells | PASS (SC1 #2) |

**Multi-channel render (view, `OrderSpike/view.json` `cellSig`):** every signal carries color
**+** a non-color channel: ORDER_BY `★`, LEAD_TIME_ZONE `[LT]`, OVERTIME `[OT]`, NON_PRODUCTION
`[X]`, IN_TRANSIT `🚚`+font, OPEN_ORDER `📦`+font, BELOW_SAFETY `⚠`+bold. Palette RGBs are the
KEY-FACT legacy ColorIndex values (3/4/10/23/36/40). The **Legend** flex decodes every enum with
its glyph, so the grid is operable with color removed (WCAG 1.4.1). Day-kind tags also appear in
the column header (`[OT]`/`[X]`).

**Verdict:** SC2 = **PASS** for the Pascal-set signals — right enum, right cell, ≥2 channels,
full legend, color-removed-operable. **Still open (not a fail):** any *template-baked
conditional-format* rule the operator relies on that is NOT one of the Pascal-set signals remains
UNREAD (xlrd/openpyxl can't read CF; LibreOffice not installed). It is closed only when David's
golden delivers per-cell ColorIndex (golden spec above) — until then, flagged UNKNOWN, not passed.

---

## SC3 — Renderer + perf — **STOCK PERSPECTIVE TABLE OK (with one perf caveat to confirm in-session)**

The build uses a **stock `ia.display.table` (PhasedGrid)**, not a Flex-Repeater. Evidence the
stock Table carries the dual-channel grid:

- **Per-cell dual-channel:** the model builder emits each day cell as a Perspective
  `{value, style}` object (`cellSig`) — `style.backgroundColor` (interior/zone), `style.color`
  (font source / below-safety), `style.fontWeight`, and a glyph/text prefix in `value`. The stock
  Table renders per-cell objects with `style`, so per-cell color+icon is native — **no
  Flex-Repeater needed for the signaling requirement.**
- **Frozen header + frozen left ID columns:** the 4 identity columns (Status / Size / Brand-Sup /
  Part No.) set `"sticky":"left"`; the Table header is frozen by default. Frozen-column
  requirement met by stock Table.
- **Scale config for ≤200×~26:** `"virtualized": true` + `"pager": {rowsPerPage:50, options:
  [25,50,100]}` — virtualized rows + pager is the documented stock-Table path for ≤1000-row grids
  (research §4). Dynamic columns (one `d{fill_pos}` per day) are built from result-set A, so the
  grid widens with FillDays without code change.
- **Resource integrity on 8.1.52:** `OrderSpike/view.json` was loaded by the gateway
  (`Setting LastModification to "external"`, 06-13 23:51:15) and survived a fresh gateway restart
  (`gwcmd -r` → state RUNNING) with **zero deserialize errors** afterward. The only
  `Unable to deserialize` in `wrapper.log` is a stale 06-12 entry for the unrelated
  `PartsStockMaster/List` view; it did not reappear. **The view parses and loads clean on
  8.1.52** — a necessary (not sufficient) gate, now met.
- **Version discipline:** the binding uses `system.db.runPrepQuery(stmt, args, "Inventory_Spike")`
  — verified Gateway/Designer-scope, 8.1-safe, no 8.3-only API; the proc is compat-100 T-SQL
  (no TRY_CONVERT, ISO-weekday math). No `system.util.getVersion()` guard needed (nothing
  8.3-only present).

**Verdict:** SC3 = **STOCK TABLE OK** — the stock Perspective Table carries the per-cell
dual-channel cells with frozen header + frozen left columns, virtualized+pager, and deserializes
clean on 8.1.52. **Flex-Repeater fallback NOT required.**

**Caveat (necessary, not sufficient — one browser click owed):** clean deserialize proves the
resource is valid, NOT that the model builder runs end-to-end in a live session or that scroll/
sort at full 200×26 is smooth. The binding's `system.util.getLogger("SPIKE").info("SPIKE grid:
A=%d B=%d C=%d rows …")` trace is already in the transform. To confirm render + perf, have David
**open `Order/OrderSpike` once in a Perspective Session** (Designer Preview will NOT log — the
binding runs gateway-scope `runPrepQuery`; use a Session at `:8088`), set Part Type = VALVE,
FillDays = 23, click **Simulate once**, then:
```
grep "SPIKE grid:" /usr/local/ignition/logs/wrapper.log | tail -1
```
Assert the line shows non-zero `A`/`B`/`C` row counts (e.g. `A=23 B≥3 C≥23`) and the grid renders
with no Perspective quality overlays on the PhasedGrid. That single click upgrades SC3 from
"deserializes clean + structurally correct" to "renders live + perf-confirmed."

---

## Bonus — F1 hazard side-check — **PASS**

`site_id` exists on BOTH `INV_PARTS_STOCK_MST` and `INV_PARTS_STOCK_MST_HIST`, so the
`INSERT INTO _HIST SELECT *` history triggers stay column-aligned (the F1 hazard the spike-db
seed handled). `dbo.sites` has 2 rows (MAS/HERO). The sim proc correctly omits `@site_id` (the
START procs don't filter on it; adding the param would give false isolation) — scoping is the
deferred DB-mod-phase surgery, not in spike scope.

---

## Code-review fixes applied (2026-06-14) + OPEN GAPS

Adversarial review resolved. Changelog:

- **F1 (WRONG — lead): PhasedGrid transform never ran (invalid brace-refs).** FIXED in
  `Order/OrderSpike/view.json`. A Perspective **script transform** receives its upstream binding
  result as `value`; `{view.custom.X}` is expression-binding-only syntax and raised
  `NameError: name 'view' is not defined`, so the model builder never executed (zero `SPIKE grid:`
  traces). The expr binding concatenates the five params pipe-delimited
  (`partType|fillDays|today|sortBy|lineName`); the transform now does
  `parts = unicode(value).split("|")` and unpacks them. View re-validated as JSON, gateway
  restarted (`gwcmd -r` → RUNNING), **zero OrderSpike deserialize errors** afterward.
  RESIDUAL: the live end-to-end run (the `SPIKE grid:` trace) still requires one Perspective-session
  click — `runPrepQuery` is gateway-scope and does not fire in Designer Preview, and there is no
  automated session here. The static defect is removed; live confirmation folds into the SC3 click
  owed below.
- **F2 (RISK — value-collapse): result-set C scalar `value` hid the forecast draw on receipt days.**
  ADDRESSED in `SIM_OrderSimulation.sql` (result set C) + renderer. Kept the build-spec §1.5 scalar
  `value` (no regression to SC assertions) and ADDED two faithful channels: `forecast_usage`
  (always the day's forecast draw) and `receipt_qty` (in-transit/open qty bucketed here, NULL if
  none). Renderer now prints `value (f<forecast>)` on receipt days so neither number is lost,
  mirroring the legacy separate-row forecast band (font 23) vs qty cell (font 10), spec §3.2/§3.3.
  Re-ran case (e) `4261102Q8000` live: fp0 now shows `value=940, forecast_usage=754, balance=44164`
  (= 43978+940−754) — the previously-invisible 754 is exposed; PAB recurrence unchanged (SC1 #3
  regression-clean). **OPEN GAP G2: which numbers the LEGACY sheet actually renders on a receipt
  column (forecast AND receipt, or one) is unresolved without David's golden.** The build is now
  information-complete either way (both values carried), but the exact legacy *layout* parity is
  PENDING GOLDEN. Architect decision deferred to `ignition-architect`/`delphi-architect` if the
  golden shows a different arrangement.
- **F3 (RISK — headline stronger than evidence): SC1/SC2 PASS rests on the stubbed calendar.**
  Accepted; the SC1/SC2 verdicts already carry "VS-LIVE-LEGACY PENDING GOLDEN" / "template-CF
  PENDING golden." **OPEN GAP G3:** every hazard-7 / OVERTIME / NON_PRODUCTION / added-leadtime
  result reproduces the **fixture** (06-17 H / 06-18 X / 06-23 O), not a golden. The algorithm-
  faithfulness items are SOUND; the *number parity* is fixture-backed until `AD_GetSpecialDate`'s
  real ALC `TireOrder` calendar is read out and the golden lands. Caveat kept load-bearing in the
  RETURN SUMMARY below.
- **F4 (RISK — artifact tracking): OrderSpike view + NQ resources are not in git.** Documented, not
  silently shipped. **OPEN GAP G4: the gateway is the artifact of record for the renderer half.**
  `view.json`/`resource.json` live only at
  `data/projects/spike/com.inductiveautomation.perspective/views/Order/OrderSpike/`; the NQ bodies
  are SQL-only in `named-queries.sql` (on-disk NQ resource JSON is undocumented and was not
  hand-authored blind — see that file's STATUS NOTE). The proc + NQ SQL + this results doc ARE
  tracked; the view JSON is reproducible from the gateway snapshot. Promoting the view to a
  committed export is a low-blast-radius follow-up for the spike.

SOUND items from the review (PAB recurrence, weekday lead-time + fallback, added-leadtime
break-loop, hazard-7 two-index reconciliation, share split, NQ param order/types, input value-prop
bindings, version discipline, secrets, resource.json) were verified by the reviewer against the
live spike DB and are unchanged by these fixes.

---

## RETURN SUMMARY

- **SC1 Number parity:** SPEC-PARITY **PASS** (hazard-7 day-header, PAB recurrence, BELOW_SAFETY⇔PAB<J7, share-sum 100%, hazard-7 bucket all reproduce the proc output); **VS-LIVE-LEGACY = PENDING GOLDEN** (no Delphi/Excel here).
- **SC2 Color/signal parity:** **PASS** for the Pascal-set signals — correct enum on correct cell, ≥2 channels each, full legend, color-removed-operable; template-baked CF thresholds still UNREAD → PENDING David's golden ColorIndex read.
- **SC3 Renderer + perf:** **STOCK TABLE OK — transform fixed (review F1).** The PhasedGrid model
  builder had invalid `{view.custom.X}` brace-refs (script-transform NameError → never ran); now
  parses the upstream `value` (`split("|")`). Per-cell dual-channel `{value,style}` cells, frozen
  header + sticky-left ID cols, virtualized+pager; deserializes clean and survives restart on
  8.1.52 (re-confirmed 2026-06-14, zero OrderSpike deserialize errors); Flex-Repeater fallback NOT
  required. **One browser click still owed** to confirm live render/perf AND the F1 transform fire
  end-to-end (instrumented `SPIKE grid:` log ready; gateway-scope binding does not run in Designer
  Preview — needs a Session at `:8088`).
- **Renderer decision:** **stock `ia.display.table`** (no Flex-Repeater).
- **Golden-export spec for David:** legacy Order→Start at **Today=2026-06-15, Line=COROLLA, FillDays=23**, against the same `.bak` data, for part-types **TIRE / VALVE / WHEEL** (covers all 5 sample cases); save each populated `OrderSimulation.xls` capturing **cell values AND per-cell interior+font ColorIndex**; also read out the real ALC `TireOrder` special dates for 2026-06-15…07-05 so the fixture calendar can be reconciled (calendar cells are otherwise fixture-conditional).
