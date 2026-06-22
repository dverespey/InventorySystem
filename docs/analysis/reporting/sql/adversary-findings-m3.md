# M3 Reporting — Adversarial SQL Parity Review (branch `m3-reports`)

Reviewer: adversarial T-SQL parity (Delphi→Ignition). Default stance: the rebuild is wrong until proven
equivalent against the LEGACY computation on the SAME inputs. Every finding below carries a file:line or a
counterexample query + its output. All probes were bounded; `Inventory_Live`/`VehicleOrder` read-only; all
mutations on `Inventory` were rolled-back or additive-then-deleted; the spike was left AS-FOUND (verified:
all fixture tags = 0, no leftover `_D6` procs).

Legacy bodies were read from the LIVE DB (`Inventory_Live`, `sys.sql_modules`), not the spec.

---

## SCOPE / METHOD AUDIT (first move: is the parity self-referential?)

`scripts/e2e/test_m3_reports.py` is **NOT circular** and **non-vacuous**:
- Every expectation derives from the SOURCE: the legacy proc (`EXEC REPORT_*`), an INDEPENDENT truth query
  (a hand-written CROSS APPLY for the corrected invoice), or a hand-transcription of the legacy Delphi loop
  (`_forecast_expected_cells`, `_lot_location_expected_cells` = MainMenu.pas:3768-3794 / 716-861), NEVER the
  rebuild's own plan.
- Non-vacuity is PROVEN by a revert step (test_m3_reports.py:849-890): a `+1` qty rebuild FAILS daily-shipping
  parity, and re-introducing the window-blind JOIN over-bills (42 vs 34). I ran it: **48 PASS / 0 FAIL**, and
  the two revert checks confirm the comparisons CAN fail on a wrong rebuild.

---

## JOB 1 — the 4 M3 report numbers

### 1. INVOICE Summary (D6) — the billing one (thorough) — VERDICT: reproduces legacy (modulo the documented D6 over-bill)

**Legacy** (`Inventory_Live` REPORT_INVOICESSummary, OBJECT_DEFINITION): window-blind
`JOIN INV_MANIFEST_COST_MST m ON d.VC_ASSY_PART_NUMBER = m.VC_ASSY_PART_NUMBER_CODE` (NO production-date
window). `invoice_summary_faithful.sql` reproduces this VERBATIM (column order, ORDER BY) — confirmed by
test_m3_reports.py R9 L1a (legacy proc == _faithful, 34 rows on live date 20250121).

**Corrected default** (`invoice_summary.sql:30`): `CROSS APPLY dbo.fn_ManifestCostAt(part, production_date)`
== an INDEPENDENT D6 truth (test R9 L1b: corrected == independent CROSS APPLY truth, 34 == 34).

**Over-bill divergence RE-PROVEN (real counterexample, rolled-back txn):**
- Part `42670FEU2000`, 8 invoiced lines on 20250121, real covering window 20230404–20260901 @ 169.3148.
- Inject a 2nd, non-covering (gap) price window:
  - faithful (window-blind): **34 → 42** (+8 — every line for that part DUPLICATES at the wrong $9.99 → OVER-BILL)
  - corrected (window-aware): **34 → 34** (unchanged — picks the one covering window)
- This is the documented, David-LOCKED D6 divergence (a report is not Toyota-facing/state-changing →
  decide-and-flag, D6 already locked). Faithful is behind the `QUERY_VARIANT='faithful'` seam; the rebuild
  SHIPS the corrected query by default.

**The 3 corrected-query concerns — all CLEAN:**
- **No double-count under OVERLAPPING windows.** Injected a 2nd window that ALSO covers 20250121: corrected
  returned **8 rows, not 16** — `fn_ManifestCostAt`'s `TOP 1 ... ORDER BY VC_START_MANIFEST DESC,
  IN_MANIFEST_COST_ID DESC` deterministically picks ONE row (newest start wins). This is strictly MORE robust
  than faithful (which would double on dirty/overlapping data). (fn body: /tmp/fn_ManifestCostAt.sql; live
  `Inventory.dbo.fn_ManifestCostAt`.)
- **No priceless-line surprise.** On 20250121, 0 lines have a cost row that fails to cover the date (the
  "corrected drops / faithful keeps" case is latent, not present today). When it does occur, corrected DROPS
  the line (window-correct; INNER semantics) and faithful would emit it at a wrong price — the intended D6
  behavior, documented at invoice_summary.sql:12-14. A part with NO cost row at all is dropped by BOTH
  (INNER) — no divergence.
- **No rounding drift.** `IN_QTY` is `int`, `MO_PRICE` is `money(19,4)`. Item Total = int × 4-decimal money is
  exactly representable in float64. The LEGACY itself computes Item Total via an Excel FORMULA `=C*D` and
  INVOICE TOTAL via `=SUM(...)` (MainMenu.pas:3660, 3668) — both IEEE-754 doubles, same precision as the
  rebuild's Python-float product (code.py:107-109). NEITHER path calls ROUND; the `$#,##0.0000` format
  (MainMenu.pas:3646/3650 == report_defs `_MONEY`) rounds the DISPLAY to 4dp identically. The test's grand-sum
  readback `got=519089.35990000004 exp=519089.3599` is float-display noise under the `<1e-4` tolerance, not a
  money divergence. No banker's-vs-away-from-zero issue (no value rounds at storage).

Param note (benign): proc declares `@PDate varchar(13)`; caller/NQ passes 8 chars; `VC_PRODUCTION_DATE` is
`varchar(8)` (all stored values exactly 8 chars). The 8-char `=` matches identically — no truncation/compare
surprise.

### 2. Daily Shipping Assy / Forecast Detail / Lot Location (spot-check) — VERDICT: faithful, numbers match

- **Daily Shipping Assy** (`daily_shipping_assy.sql`): VERBATIM the legacy proc (same cols, IN_ASN_ID join,
  GROUP BY incl. `s.IN_QTY, d.IN_QTY`, ORDER BY). Live 20250121: NQ == legacy proc, **27 == 27 rows,
  `legacy EXCEPT nq` = 0**. Clean key join (no date-only fan-out, no orphan-drop).
- **Forecast Detail** (`forecast_detail.sql`): legacy is `SELECT * ... ORDER BY 3 cols`; NQ pins the 8 read
  columns with the SAME ORDER BY. **50 == 50 rows**, all 8 NQ cols exist in the table (SELECT* superset), and
  **0 tie groups** on the 3 sort keys (render sequence deterministic on current data; any future tie would be
  shared with the legacy proc's identical ORDER BY — not a rebuild divergence).
- **Lot Location** — confirmed the **LIVE PLANT** procs, not the D9 NUMMI twin: the live Delphi handler calls
  `REPORT_PLANTLotLocationW` (MainMenu.pas:791) and `REPORT_PLANTLotLocation` (MainMenu.pas:829);
  `REPORT_NUMMILotLocation` is NOT referenced and its body is fully commented out (`SELECT null`,
  /tmp/REPORT_NUMMILotLocation.sql). NQs are verbatim the PLANT procs; test R12 (seeded) shows NQ == legacy
  (wheels 2==2, tires 2==2) and is non-vacuous (unseeded spike = 0 rows).

`test_m3_reports.py` derives all expectations from the legacy proc / independent truth / .pas transcription
(NOT the rebuild) and is revert-proof (non-vacuous). CONFIRMED.

---

## JOB 2 — the D6 / 856 finding (CONFIRM/REFUTE)

### (a) QI-OFF vacuity of `test_report_procs_d6.py` — CONFIRMED (partially refined)

- The filtered UNIQUE index `UX_INV_ASN_MST_LINE_PDATE_NORMAL` (filter `[VC_START_SEQ_NUMBER]<>'-1'`) exists on
  `INV_ASN_MST` (sys.indexes). `sqlcmd` defaults to `QUOTED_IDENTIFIER = 0` (verified:
  `SESSIONPROPERTY('QUOTED_IDENTIFIER')=0`). DML against a filtered-index table under QI OFF FAILS with
  **Msg 1934** (reproduced); the same INSERT succeeds under QI ON (reproduced, IN_ASN_ID returned).
- `test_report_procs_d6.py`'s `sqlq()` issues `-Q "SET NOCOUNT ON; ..."` with NO `SET QUOTED_IDENTIFIER ON`,
  so its 856 seed INSERT into `INV_ASN_MST` (line 144) hits Msg 1934.
- REFINEMENT of the read: the seed does not produce a silent GREEN — sqlcmd writes Msg 1934 to **stdout** and
  exits **0** (no `-b`), so `check_output` does not raise; `sqlq()` filters the `Msg ` line but KEEPS the
  continuation line, so `asnId` becomes the garbage string `"INSERT"`. Running it today the 856 block FAILS
  loudly (`migrated rows=1`, `legacy=0` — both the fan-out and the parity assert FAIL), AND it **LEAKS fixtures**
  (a mid-batch `IN_ASN_ID='INSERT'` conversion error, Msg 245, aborts the teardown so `ZZ856PART` cost+forecast
  rows survive — I observed `ZZ856PART_cost=1, fcst=2` left behind and cleaned them up). So the net is worse
  than "vacuously green": the 856 sub-test is **non-functional under QI OFF and not fixture-clean.** This is a
  **test/parity-method flaw**, not a code defect.

### (b) Is the divergence a forecast fan-out collapsing to 1 in `REPORT_EDI856_D6`? — REFUTED

Reproduced under QI ON with a proper seed (1 'C' ASN, 1 detail, TWO distinct-kanban forecast rows K1/K2, 1
covering cost window), deploying the `_D6` copies exactly as the test renames them, comparing the legacy
`REPORT_EDI856 @EIN=0` vs `REPORT_EDI856_D6 @EIN=0`, all in a committed-then-deleted fixture:

```
LEGACY REPORT_EDI856    @EIN=0 -> ZZ856PART rows = 2  (K1, K2)
D6     REPORT_EDI856_D6 @EIN=0 -> ZZ856PART rows = 2  (K1, K2)
```

**They MATCH at 2.** There is **NO `legacy=2 vs migrated=1`** in `REPORT_EDI856_D6`. The D6 proc keeps the
INNER `JOIN INV_FORECAST_DETAIL_INF f` and `GROUP BY ... f.VC_ASSY_KANBAN_NUMBER` — confirmed against the live
body: forecast INNER JOIN = YES, TOP-1 forecast collapse = NO, GROUP BY keeps kanban = YES
(spike-report-procs-d6.sql:108-112). The `migrated=1 / legacy=0` the test prints today is an ARTIFACT of the
broken/leaked seed state (no 'C' ASN), NOT a fan-out collapse — in the pure leak-only state both legacy and D6
return 0. The CROSS APPLY in the D6 proc is used ONLY for the price/cost lookup (a per-part-per-date single
covering window — correct), never for the kanban join. **The mechanism stated in the finding ("CROSS APPLY
(SELECT TOP 1 ...) collapse" in REPORT_EDI856_D6) is not present in this proc** — that TOP-1 collapse was the
EARLIER 856-BUILDER bug, already fixed (see (c)).

### (c) Is the shipped 856 builder (PR #29, `spike-edi856-feed.sql`) independent + correct? — CONFIRMED

- INDEPENDENT: the feed is a standalone parameterized SELECT keyed by `@ASNID`; it does NOT call
  `REPORT_EDI856_D6` (or any report proc) and has NO self-flip UPDATE (spike-edi856-feed.sql:1-10, 58-74).
- CORRECT fan-out: it uses the INNER `JOIN INV_FORECAST_DETAIL_INF f ON d.VC_ASSY_PART_NUMBER =
  f.VC_ASSY_PART_NUMBER_CODE` with `GROUP BY ... f.VC_ASSY_KANBAN_NUMBER` keeping each DISTINCT kanban (2
  kanbans → 2 LIN lines), and KEEPS `m.MO_PRICE` in the GROUP BY (so price-distinct windows don't collapse).
  The header explicitly records the fix: it REPLACED the old `CROSS APPLY (SELECT TOP 1 ...)` that had no
  ORDER BY over a heap (nondeterministic) and dropped distinct kanbans (spike-edi856-feed.sql:20-27). The
  builder is therefore UNAFFECTED by anything in `REPORT_EDI856_D6` and is the correct, deterministic feed.

### (d) Disposition of `REPORT_EDI856_D6`

`REPORT_EDI856_D6` is **NOT installed** in `Inventory` or `Inventory_Live` (verified ABSENT in both) — it
exists only as a transient `_D6`-renamed deploy in `test_report_procs_d6.py` derived from
`spike-report-procs-d6.sql`'s `REPORT_EDI856`. For the OUTBOUND 856 wire, the shipped builder
(`spike-edi856-feed.sql`, PR #29) is the production path and is correct and independent; the D6 EDI856 proc is
**superseded** for the wire. It remains useful only as a parallel-run/diff artifact for the legacy on-read
report proc during cutover. Punch-list (NOT an M3 blocker):
1. **Fix the test seed (required to make the 856 sub-test meaningful):** add `SET QUOTED_IDENTIFIER ON` to
   `test_report_procs_d6.py`'s `sqlq()` (and harden the teardown so a failed seed cannot leak — guard on a
   numeric `asnId`). Without this the 856 fan-out assertion is non-functional and leaks fixtures.
2. Either retire `REPORT_EDI856`(/`_D6`) as superseded by the builder for the wire, or keep it explicitly as a
   parallel-run report artifact (it is, by construction, faithful for the @EIN=0 read — proven 2==2 above).
   No fan-out fix is needed in the D6 proc; it already keeps distinct kanbans.

---

## FINDINGS (ranked)

- **BLOCKER:** none.
- **SHOULD-FIX (test/parity-method flaw, punch-list — not an M3 blocker):** `test_report_procs_d6.py` runs its
  856 seed under `QUOTED_IDENTIFIER OFF`, so the `INV_ASN_MST` INSERT hits Msg 1934; the fan-out/parity asserts
  for EDI856 are non-functional AND the run leaks `ZZ856PART` fixtures (teardown aborts on a `IN_ASN_ID='INSERT'`
  conversion error). Add `SET QUOTED_IDENTIFIER ON` and a numeric-`asnId` guard. (Counterexample + repro above;
  spike restored as-found.)
- **NIT (data-vintage, documented):** the INVOICE Summary D6 over-bill and the priceless-line drop are LATENT
  on current data (every part has exactly one covering window on 20250121); a clean live diff cannot exhibit
  them today — both are proven only by injection. This is the expected "self-consistent ≠ equivalent on drifted
  data" situation; the divergence is characterized by injection, not asserted away.
- **NIT:** Forecast Detail's 3-column ORDER BY is not provably unique (0 ties today, data-dependent); shared
  with the legacy proc, so not a rebuild divergence.

---

## VERDICTS

1. **The 4 M3 reports reproduce the legacy numbers** (modulo the documented, David-LOCKED INVOICE-Summary D6
   over-bill, which the rebuild deliberately CORRECTS by default and reproduces faithfully behind the
   `faithful` seam): Daily Shipping Assy 27==27 (0 diff), Forecast Detail 50==50, Lot Location 2==2/2==2
   (PLANT, confirmed not D9 NUMMI), INVOICE Summary corrected == independent D6 truth + faithful == legacy
   verbatim, with no double-count / no priceless surprise / no rounding drift. The harness is non-circular and
   revert-proven non-vacuous (48 PASS / 0 FAIL). **SHIP M3.**

2. **The D6 / 856 finding:** (a) the QI-OFF trap is CONFIRMED — `test_report_procs_d6.py`'s 856 seed is
   non-functional and fixture-leaky under QI OFF (a real test-method flaw). (b) the claimed
   `legacy=2 vs migrated=1` fan-out collapse in `REPORT_EDI856_D6` is **REFUTED** — under QI ON with a proper
   seed both legacy and D6 return 2 (distinct kanbans); the D6 proc has no TOP-1 forecast collapse. (c) the
   shipped 856 builder (`spike-edi856-feed.sql`, PR #29) is **INDEPENDENT and CORRECT** (does not call the D6
   proc; INNER forecast fan-out + GROUP BY keeps distinct kanban; keeps MO_PRICE) and is **UNAFFECTED**.
   Disposition: the D6 EDI856 proc is superseded by the builder for the wire (or kept as a parallel-run
   artifact); the ONLY fix needed is in the TEST (add `SET QUOTED_IDENTIFIER ON` + harden teardown), NOT in
   the builder and NOT in the D6 proc's fan-out. **Punch-list item, not an M3 blocker.**
