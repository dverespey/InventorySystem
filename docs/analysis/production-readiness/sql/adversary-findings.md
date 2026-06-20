# ASN-creation rebuild — adversarial parity review (findings)

**Target under attack:** `computeAsnDetails` + `create_asn` in
`docs/analysis/edi/project-library/asn/code.py`, claimed equivalent to the legacy ASN-creation
chain (`AD_FRSPULL` on `VehicleOrder` → `SELECT_ForecastDetailBCASN` on `Inventory` → Delphi
`CalculateASNFRS` fan-out `DataModule.pas:5106-5318` → `INSERT_ASNInfo` / `INSERT_ASNDetail`).

**Method:** every claim below is proven against the LIVE running procs/data on `mssql-spike`
(`VehicleOrder` real ALC 2.33M rows, `Inventory` rebuild target, `Inventory_Live` legacy ref —
read-only; destructive probes in rolled-back transactions). Posture: "wrong until proven equivalent."

---

## RESOLUTION OF THE #1/#2 COUPLED GAP (the strip + char3/char21 LIKE) — DEFINITIVE

### char(3) vs char(21): RESOLVED — it is **char(3)**. The "char(21)" claim is WRONG.

The equivalence-map (#5) carried an unresolved source contradiction: `AD_FRSPULL-analysis.md` +
`AD_FRSPULL-shared.sql:46/69` say `convert(char(3), …)`; `delphi-fanout-confirmation.md §f` says
`convert(char(21), …)` from "a different VehicleOrder.sql dump." Adjudicated against the LIVE proc:

- `sys.dm_exec_describe_first_result_set('EXEC dbo.AD_FRSPULL …')` → **`BC char(3)` (max_length 3)**.
- `OBJECT_DEFINITION(OBJECT_ID('dbo.AD_FRSPULL'))` dumped via `PRINT` → **two occurrences of
  `convert(char(3),…)`**, zero of `char(21)`.
- `SELECT_ForecastDetailBCASN` param `@BCode` = **`varchar(20)`**; column `VC_BROADCAST_CODE` =
  **`varchar(20)`** (collation `SQL_Latin1_General_CP1_CI_AS`). There is **no `char(21)` anywhere**
  in the live chain.

**Verdict:** `delphi-fanout-confirmation.md §f`'s char(21) reading is a stale/foreign-dump artifact
and must be discarded. The operand width the LIKE matches on is `char(3)` (so spare BCs carry exactly
ONE trailing space, e.g. `'NN '`).

### `bc.strip()` (code.py:278): RESOLVED — **SAFE today**, on both the LIKE match AND collisions.

Two independent live proofs:

**(a) The LIKE match is identical padded vs stripped.** T-SQL `LIKE` ignores trailing spaces on the
*matched (left)* operand. Proven directly:

```
'NN ' LIKE '[MNP][N]'  -> MATCH
'NN'  LIKE '[MNP][N]'  -> MATCH
'NN  'LIKE '[MNP][N]'  -> MATCH      (extra spaces still match)
'NNX' LIKE '[MNP][N]'  -> no         (a real trailing char is NOT ignored — so strip ≠ chop)
```

End-to-end through the proc for every REAL spare BC (`NN `, `NP `, `PN `) and a ground BC (`NBB`), on
BOTH `Inventory` and `Inventory_Live`, the matched-forecast-row count is byte-identical padded vs
stripped (1/1, 1/1, 1/1, 3/3). The earlier analyses' fear — "T-SQL LIKE does NOT trailing-trim the
left operand, so a padded vs stripped operand can diverge" (equivalence-map #5; AD_FRSPULL-analysis
§6) — is **factually wrong**. It does trim the left operand for `LIKE`.

**(b) `.strip()` cannot collide two AD_FRSPULL BC rows into one `forecastByBc` key.** Over a 5-day
COROLLA window, every ground stripped key is 3 chars and every spare stripped key is 2 chars — two
structurally disjoint namespaces; `collisions_after_strip = 0`. AD_FRSPULL's `GROUP BY` + `UNION`
already makes each BC unique, and the rebuild's `if bc in forecastByBc: continue` (code.py:286)
dedups on a key that is unique post-strip.

**Legacy cross-check:** the ASN-create fan-out passes the BC as
`ALC_StoredProc.FieldByName('BC').AsString` **without TRIM** (`DataModule.pas:5152`) — contrast the
inventory-pull path which DOES `TRIM(...)` (`DataModule.pas:4781`). So legacy feeds the char(3)
value; the rebuild feeds the stripped value. Because T-SQL LIKE trims the left operand anyway, the
two produce identical matches. **`bc.strip()` is NOT a defect today.**

**Residual (latent, shared with legacy):** the safety rests on every BC component being exactly 1
char (so ground=3 / spare=2, disjoint). If GALC ever emitted a blank ground component (making a
2-char ground BC), `.strip()` could collide it with a 2-char spare key — but that same blank
component would already produce a wrong/NULL BC in the legacy too. This is the latent
blank/NULL-component hazard (AD_FRSPULL-analysis §3.6), not a strip-specific bug. Tracked under NIT-2.

---

## FINDINGS

### BLOCKER-1 — No DB unique constraint behind the idempotency guard → gateway concurrency double-insert

- **Claim under test:** the `SELECT_ASNSeq` read-back guard prevents a second ASN for the same
  `(line, prodDate)`.
- **Counterexample (live):** the only unique index on `INV_ASN_MST` is `PK_INV_ASN_MST` on the
  IDENTITY column `IN_ASN_ID` (verified via `sys.index_columns`: key_ordinal 1 = `IN_ASN_ID`, and
  only one row in `sys.indexes WHERE is_unique=1`). There is **no unique constraint on
  `(VC_LINE_NAME, VC_PRODUCTION_DATE)`**. `create_asn` (code.py:263-267) does a check-then-act:
  `SELECT_ASNSeq`; if `len(existing)==0`, open a tx and insert — with **no serialization** between
  the read and the insert. Two concurrent gateway `create_asn` calls for the same (line, prodDate)
  both read 0 rows, both proceed, both commit two complete ASNs (header + full fan-out).
- **Legacy vs rebuild:** the legacy guard was a UI form-lock on a single-user desktop
  (`ASNSelect.pas:180-182`) — structurally serial, never concurrency-exposed. The headless gateway
  driver IS concurrency-exposed and inherits the same non-atomic guard with nothing to enforce it.
- **Class:** code/architecture defect introduced by the runtime change (desktop→gateway).
- **Fix owner (not me):** add a unique index on `(VC_LINE_NAME, VC_PRODUCTION_DATE[, IN_SITE_ID])`
  or serialize the create; the read-back guard alone is insufficient under concurrency.

### SHOULD-FIX-1 — GAP #2: No-Ratio assy/manifest pick is nondeterministic, and it FIRES on live data

- **Claim under test:** the rebuild reproduces the legacy fan-out's per-manifest distribution.
- **Counterexample (live, ASN-4721 window 2026-06-18 00:33:49 → 2026-06-19 01:01:51, COROLLA):**
  `AD_FRSPULL` returns `PEE` as a **No-Ratio** BC (ground, VEHICLES=1, ORDERS=4 ≤ 5). `PEE` matches
  the 2-row forecast pattern `[MNP]EE` → two DIFFERENT assy/manifests: `42600FEL1000`/manifest **m36**
  and `42600FEL2000`/manifest **m37**. The No-Ratio branch (`computeAsnDetails` code.py:148-160) emits
  **one** row from `fcRows[0]` and `break`s. `SELECT_ForecastDetailBCASN`'s base tables
  `INV_FORECAST_DETAIL_INF` and `INV_MANIFEST_COST_MST` are **HEAPs with no clustered index and no
  `ORDER BY`** (verified `sys.indexes`), so `fcRows[0]` is heap-scan/allocation order, not contractual.
- **Proven the pick actually swung the result:** decomposing the rebuild's persisted m36=561 →
  NEE-ratio contributes `round(199*4*70/100)=557`, leaving `561-557=4 = V*assy (1*4)` from PEE's
  No-Ratio row → PEE picked `fcRows[0]=FEL1000/m36`. Had the heap returned FEL2000/m37 first, the
  4 units would have landed on m37 instead (m36=557, m37=243). **Which manifest the single-vehicle PEE
  ships is decided purely by nondeterministic order.**
- **Legacy vs rebuild:** the legacy is *equally* nondeterministic (`Inv_StoredProc.First` then first
  row + `break`, `DataModule.pas:5177-5212`); today both DBs happen to return FEL1000/m36 first (stable
  across 6 reruns) — but that is incidental heap allocation. The rebuild did **not** apply the
  analysis-prescribed `ORDER BY ID_FORECAST_DETAIL` (the column exists, col 1 of
  `INV_FORECAST_DETAIL_INF`). So the rebuild faithfully reproduces an *inherited* nondeterminism;
  parity is luck-of-allocation and can silently diverge on a table reload/page-reuse/replan in either
  system independently.
- **Class:** inherited code defect (not rebuild-introduced), but the prescribed hardening is missing
  AND David's "which assy should the single-vehicle case pick?" decision is still open.

### SHOULD-FIX-2 — GAP #6: the parity test's "total-qty invariant (4240==4240)" is partly COINCIDENTAL, not a clean conservation law

- **Claim under test:** `test_create_asn_parity.py:18-21, 237-243` and its NOTE sell
  "TOTAL-QTY INVARIANT … the fan-out conserves the same grand total even when the per-manifest split
  has since changed … a REAL parity signal that survives the vintage drift."
- **Counterexample (live, ASN 4721):** grouping the per-manifest qtys by their source BC trio:

  | BC group | legacy sum | rebuild sum | conserved? |
  |---|---|---|---|
  | NBB (m57/58/59) | 2104 | **2105** | NO (+1) |
  | NEE+PEE (m36/37) | 800 | 800 | yes |
  | NN (m51) | 828 | 828 | yes |
  | NJJ (m63/64/65) | 116 | **115** | NO (−1) |
  | KK (m74/75/76) | 260 | 260 | yes |
  | grand total | **4240** | **4240** | yes (coincidence) |

  The grand total matches **only because the NBB +1 and the NJJ −1 cancel**. Proven mechanism: NBB
  has VEHICLES=526, assy=4, ratios 40/20/40 (sum=100); `round(526*4*40/100)=842`,
  `round(526*4*20/100)=421`, `round(526*4*40/100)=842` → sum 2105 ≠ base 2104 (each .6 rounds up).
  NJJ has VEHICLES=29 → 46+23+46=115 ≠ base 116 (each .4 rounds down). The rebuild's own fan-out does
  **NOT** conserve `VEHICLES*assy` per BC; `sum(round(base*ratio_i/100))` can drift ±1 even when the
  ratios sum to 100.
- **Why it matters:** the test presents the grand-total match as evidence the fan-out is correct
  ("survives the vintage drift"). It is a weak check that can pass by cancellation of opposite
  per-BC rounding residuals; a different window where the residuals did NOT cancel would FAIL a
  total-only gate while the per-BC math was equally faithful, or vice-versa. The honest invariant the
  fan-out actually obeys is **per-BC, not grand-total**, and even per-BC it is `≈ base ± (rounding
  residual)`, not exact.
- **Class:** test/parity-method flaw (the qty math itself is faithful — see PROVEN section). The test
  is NOT vacuous (it does run the real driver and does a legitimate self-consistency diff), but the
  total-qty claim oversells coincidence as a conservation law. Downgrade the NOTE's language and, if a
  total gate is kept, gate per-BC (or document that the grand total can match by cancellation).

### SHOULD-FIX-3 — GAP #4: char(3) BC truncation has NO alarm (latent silent wrong-BC)

- **Claim under test:** the rebuild reproduces AD_FRSPULL's `convert(char(3),…)` AND, per
  AD_FRSPULL-analysis §7.3 ("the single highest-severity latent defect"), alarms when any BC
  component exceeds 1 char.
- **Counterexample:** `CONVERT(char(3),'N'+'BB'+'X')` → `'NBB'` (live, silent chop of the 4th char).
  Every BC component is exactly 1 char today (verified: GROUNDTIRE/GROUNDWHEEL/SPARETIRE
  `MAX(LEN(DataValue))=1`), so truncation does not fire and parity holds. But `create_asn` /
  `computeAsnDetails` have **no width assertion** (grep: the only `len()` guard is on `productionDate`,
  code.py:72; nothing on the BC). If GALC ever emits a 2-char wheel/tire value, both legacy AND
  rebuild silently produce a wrong-but-valid-looking 3-char BC that matches the wrong (or no) recipe.
- **Class:** latent defect, faithful-to-legacy today (both wrap the same proc), missing the
  analysis-mandated alarm. Not a today-divergence.

### NIT-1 — GAP #3: banker's rounding is correctly implemented, but the .5 boundary is UNREACHABLE on current data (double-defense, no action)

- `_bankers_div_round` (code.py:45-61) is correct half-to-even: unit-tested at exact halves
  `50/100→0, 150→2, 250→2, 350→4, 450→4` and negatives `−50→0, −250→−2` — and it **diverges from
  T-SQL `ROUND` (half-away)** at every .5 (`50/100`: bankers 0 vs half-away 1; `250`: 2 vs 3; `450`:
  4 vs 5). It is also correct on Jython 2.7 (whose `round()` is half-away — would break parity) by not
  using the builtin. So if the math were ever moved to SQL via `ROUND`, parity would break by ±1.
- **But it cannot fire today:** every live split row has `IN_ASSY_QTY=4` and `IN_TIRE_RATIO ∈
  {20,30,40,70}`. `gcd(4*tire,100)=20` for all four, and 20 ∤ 50, so `(VEHICLES*4*tire) mod 100`
  can never equal 50 → the value never lands on x.5 for any integer VEHICLES. The fractional parts
  that DO occur are .0/.2/.4/.6/.8, where banker's and half-away agree. **No divergence possible on
  current data; the implementation is correct for the day assy/ratios change.** No action.

### NIT-2 — `.strip()` latent collision shares the legacy's blank-component hazard

- Resolved-safe today (see top section). The only way `.strip()` could merge a ground and spare BC is
  a blank ground component making a 2-char ground BC — the same NULL/blank-component data-integrity
  hazard the legacy also mishandles (AD_FRSPULL-analysis §3.6). No live occurrence (0 blank/null
  ground values in-window). Treat as a data alarm, not a strip-specific fix.

---

## WHAT IS PROVEN EQUIVALENT (no divergence found)

- **Qty math, ratio branch (the revenue-critical path).** `qty = VEHICLES*IN_ASSY_QTY` when both
  ratios=100, else `banker_round(VEHICLES*IN_ASSY_QTY*IN_TIRE_RATIO/100)` — tire-only numerator,
  literal-100 denominator, both-100 gate reads both ratios. Matches `DataModule.pas:5226-5235` exactly;
  verified live (NEE V=199: m36=557, m37=239; NBB V=526: 842/421/842).
- **No-Ratio branch:** `Orders<=5` → one row from `fcRows[0]`, `qty=VEHICLES*IN_ASSY_QTY` (no ratio),
  `break`. Matches `DataModule.pas:5183-5212` (modulo the SHOULD-FIX-1 order hazard, which is inherited).
- **Manifest generation:** `"7"+prodDate[3:8]+assy` = `'7'`+1-digit-year+MM+DD+2-char id. Verified
  `manifest('20260618','57')='76061857'`, `('…','36')='76061836'`. Matches `DataModule.pas:5186/5239`.
- **BC→forecast LIKE handoff:** padded vs stripped operand produce byte-identical matches on both DBs
  (RESOLUTION section). LIKE direction (column=pattern, BC=literal left), CI collation preserved by
  leaving the proc unchanged.
- **Two cost checks, correctly distinguished.** PRE-loop hard-abort on `IN_MANIFEST_COST_ID is None`
  before the tx opens (code.py:139-146, 305) = `DataModule.pas:5160-5175`. POST-loop
  `SELECT_ASNMissingCost` warn-only, try/except-swallowed (code.py:347-358) = `DataModule.pas:5285-5308`.
  Proven live on a REAL expired-cost part `42670FET9000` (cost id 151 non-NULL, window
  20230404..20251227 expired vs prodDate 20260618): PRE-loop PASSES (id non-NULL), POST-loop returns
  "Manifest Cost Entry is out of date" — rebuild WARNs and prices the line (does not abort, does not
  silently drop). Exactly legacy behavior.
- **Header write:** status hard-coded `'C'`, `@ASNID` via OUTPUT/SCOPE_IDENTITY captured on the same
  tx; EIN=0 at create (intended at-SEND divergence — harness correctly ignores `IN_ASN_EIN`).
- **INSERT_ASNDetail accumulate:** re-keyed `(IN_ASN_ID, VC_MANIFEST_NUMBER)` upsert; within-ASN sums
  preserved (the parity run shows several manifests summed from a ratio row + a No-Ratio row, e.g. m36
  = NEE 557 + PEE 4 = 561).

## INTENDED DIVERGENCES (deliberate, not parity defects — confirmed)

- **EIN=0 at create / real EIN at SEND** (removes the legacy out-of-txn `UPDATE Site SET SiteEIN+1`
  no-WHERE multi-site bug + the read-then-bump race). Harness excludes `IN_ASN_EIN`. Correct.
- **Single surrounding transaction** (code.py:309-343) vs legacy per-ExecProc auto-commit. Legacy
  leaves a partial ASN (header + earlier-BC details) on a mid-loop abort; the rebuild is all-or-nothing
  and the pre-loop cost abort runs before the tx opens (code.py:305) so an abort writes nothing. This
  is a fix; failure-path output is intentionally NOT byte-parity. Acknowledge at sign-off.

---

## VERDICT

**The rebuild's per-row qty/manifest MATH is PROVEN equivalent to the legacy** for the
quantity computation (ratio branch, both-100 full-qty, manifest scheme, banker's rounding, the
BC→forecast LIKE handoff including the `.strip()`, and the two cost guards). I could not construct an
input where the *arithmetic* diverges on current data.

**GAP #1/#2 (strip + char3/char21) are RESOLVED:** char width is **char(3)** (the char(21) claim is
wrong); `bc.strip()` is **SAFE** — it neither changes the LIKE match (T-SQL trims the left operand)
nor causes a BC collision (ground/spare namespaces are disjoint 3-vs-2 chars). Not a defect today.

**But equivalence of the persisted ROWS is NOT fully proven, on two axes:**

1. **Row-for-row legacy parity is UNPROVABLE from the available data.** Every reproducible legacy ASN
   (4718-4721; 4722 is a hot-call/manual) was frozen under an older forecast-recipe vintage in
   `Inventory_Live`'s own history (the per-BC qtys imply fractional/mutually-inconsistent vehicle
   counts under today's recipe — e.g. NBB legacy 80/900/1124 vs today's assy=4 tire 40/20/40). The
   per-manifest distribution differs by design of the historical data, and no legacy ASN built under
   the current recipe exists. The test honestly labels this (A) self-consistency + (B) total-qty and
   does NOT sell it as row parity — that discipline is correct. **This part is a data-vintage gap, not
   a code defect.**

2. **Two real, surfaced risks** that are not closed:
   - **BLOCKER-1:** no DB unique constraint behind the idempotency guard → gateway concurrency can
     double-insert a full ASN. New-runtime defect; must be fixed before production.
   - **SHOULD-FIX-1:** the No-Ratio single-vehicle assy/manifest pick is nondeterministic and FIRES on
     live data (PEE→m36-vs-m37); parity is luck-of-heap-allocation, no `ORDER BY` applied, David's
     pick decision still open.
   - **SHOULD-FIX-2:** the parity test's grand-total invariant passes partly by ±1 cancellation across
     BCs — it is not the clean conservation law the test claims (gate per-BC or soften the language).
   - **SHOULD-FIX-3:** char(3) truncation alarm absent (latent silent wrong-BC if GALC emits >1-char).

**Bottom line:** the qty/manifest math is faithful; the `.strip()`/char-width fear is disproven; but
the rebuild is **NOT yet production-equivalent** — BLOCKER-1 (concurrency) must be fixed, and
SHOULD-FIX-1 (No-Ratio nondeterminism, already firing) and SHOULD-FIX-2 (the oversold total invariant)
must be addressed before the parity claim can be signed off. Row-for-row legacy parity remains
UNPROVABLE from the available frozen data (recipe-vintage drift), which the test correctly refuses to
paper over.
