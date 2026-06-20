# ASN-Creation Keystone — Legacy → Rebuild Equivalence Map

**Purpose.** For the M1 ASN-creation E2E architecture doc: a per-behavior equivalence table mapping
each legacy behavior/edge-case to its rebuild location, with a faithfulness status. Feeds the Verify
gate and David's sign-off.

**Sources (legacy behavior — all complete on disk):**
- `AD_FRSPULL-analysis.md` (the GALC vehicle-count pull; proven on live VehicleOrder)
- `SELECT_ForecastDetailBCASN-analysis.md` (BC → parts/ratios lookup)
- `asn-write-chain-analysis.md` (INSERT_ASNInfo / INSERT_ASNDetail / SELECT_ASNSeq / SELECT_ASNMissingCost)
- `delphi-fanout-confirmation.md` (delphi-architect's live-source confirmation of the Pascal fan-out)

**Rebuild artifacts (the "rebuild code location" column points at these):**
- `docs/analysis/edi/project-library/asn/code.py` — pure `computeAsnDetails` (:85-179) + driver `create_asn` (:207-363)
- `docs/analysis/edi/spike-asndetail-rekey.sql` — the Q1 re-key of `INSERT_ASNDetail` / `DELETE_ASNItem`
- `m1-asn-creation-spec.md` — the decoded create chain spec

**Decisions baked in (from the task + spec):** EIN allocated at SEND (create writes 0); Q1 detail key
`(site_id, IN_ASN_ID, manifest)` with `site_id` deferred to M4; banker's rounding to match Delphi `Round`.

**Status legend:** ✅ faithful · ⚠️ intended divergence (rationale required) · ❌ GAP / risk (real, to surface).

---

## Equivalence table

| # | Legacy behavior (cite) | Rebuild location | Status |
|---|---|---|---|
| 1 | **BC composition.** Ground BC = `ModelYearCode + GROUNDWHEEL + GROUNDTIRE` (WHEEL before TIRE; alias vd2=WHEEL/vd1=TIRE is reverse of read order); spare BC = `ModelYearCode + SPARETIRE`, spare filtered `<> 'M'`. (AD_FRSPULL-analysis §2, §7.1) | **Not reimplemented — read from the live `AD_FRSPULL` proc** via `system.db.runPrepQuery("EXEC AD_FRSPULL …")` (code.py:272-274). The composition stays inside the cross-DB proc; the rebuild consumes its `BC`/`ORDERS`/`VEHICLES` output verbatim (code.py:276-279). | ✅ faithful (wrap-the-proc: the BC formula is never re-derived in Ignition, so concat order/spare can't drift) |
| 2 | **char(3) trailing-space padding is semantic** — separates ground (3 real chars) vs spare (2 chars + trailing space) namespaces, and the spare's trailing space flows into the downstream `LIKE`. Trimming can collide rows under UNION-dedup and break the LIKE. (AD_FRSPULL-analysis §3.4, §7.2; §6) | `create_asn` **strips** the BC immediately after reading it: `bc = bc.strip()` (code.py:278), comment "char(3) -> may be space-padded". | ❌ **GAP/RISK.** The strip destroys the trailing-space namespace separator *before* the BC is used as the `LIKE` left operand (code.py:289). Today disjointness holds because ground codes are 3 non-space chars and spare are 2 + space; stripping makes a spare `"NN "` → `"NN"`. Two real risks: (a) `forecastByBc` keys on the stripped BC (code.py:286,300) so a hypothetical ground/spare collision silently merges fan-outs; (b) **see #5** — the stripped BC changes what the proc's `LIKE` matches vs. the legacy padded operand. Must be reconciled with the char(3)-vs-char(21) contradiction (#5) before sign-off. |
| 3 | **char(3) silent truncation** — `CONVERT(char(3), …)` chops a >3-char concat with no error; safe only while every component is 1 char. Latent wrong-BC bug; rebuild must reproduce truncation for parity AND alarm when any component > 1 char. (AD_FRSPULL-analysis §3.1, §7.3) | Truncation lives inside the wrapped proc (faithful by construction). **No alarm / >1-char assertion** anywhere in `create_asn` or `computeAsnDetails`. | ⚠️→❌ Truncation itself is ✅ faithful (proc-resident). But the **alarm the analysis mandates is absent** → ❌ GAP: a future multi-char GALC value silently yields a wrong BC and the rebuild won't flag it. Low live-probability (all values 1 char today) but it is an explicit "MUST … assert/alarm" from §7.3. Surface for David. |
| 4 | **ORDERS = COUNT×4 (ground) / COUNT (spare); VEHICLES = COUNT.** ×4 is the 4-corners assumption and gates only the No-Ratio branch (it is not a qty). (AD_FRSPULL-analysis §2, §3.2) | Proc emits `ORDERS`/`VEHICLES`; rebuild passes them through unchanged (code.py:279) and `computeAsnDetails` uses `orders` only for the `<= 5` gate (code.py:148) and `vehicles` only as the qty multiplier (code.py:151,166). | ✅ faithful (×4 stays in the proc; rebuild never recomputes ORDERS, so the ground≈1-vehicle / spare≈5-vehicle gating coupling is preserved) |
| 5 | **BC→forecast match is a `LIKE` pattern:** `@BCode LIKE VC_BROADCAST_CODE` (column is the pattern, BC is the literal left operand), CI collation. **The two analyses disagree on the BC width AD_FRSPULL emits: `char(3)` (AD_FRSPULL-analysis §1, AD_FRSPULL-shared.sql:46/69) vs `char(21)` (delphi-fanout-confirmation §f, citing a different dump).** Whatever padding the legacy applies between AD_FRSPULL and the LIKE is significant under LIKE and must be reproduced. (AD_FRSPULL-analysis §6; SELECT_ForecastDetailBCASN-analysis §1a) | Rebuild feeds the **stripped** BC (code.py:278) as the `@BCode` left operand into `EXEC SELECT_ForecastDetailBCASN @BCode=?` (code.py:289). LIKE direction is preserved (proc unchanged; column stays the pattern). | ❌ **GAP/RISK — unresolved source contradiction + a possibly-wrong rebuild choice.** (1) The char(3) vs char(21) discrepancy is a **live source conflict** that must be adjudicated against the running proc before parity can be claimed — the two analyses cite different VehicleOrder.sql dumps. (2) The rebuild's `.strip()` means it feeds **neither** `char(3)` nor `char(21)` semantics — it feeds a trimmed varchar. For a ground BC (`NBB`) trim is a no-op so it matches. For a spare BC, legacy feeds `"NN "` (char(3)) or `"NN" + 19 spaces` (char(21)); T-SQL `LIKE` does NOT trailing-trim the *left* operand, so a stored pattern authored to match a padded operand vs. a stripped operand can diverge. **Must verify the spare-BC LIKE match end-to-end against the live proc** (SELECT_ForecastDetailBCASN-analysis §1b flags trailing-space; delphi-fanout-confirmation residual item #3 also lists this as unconfirmed). LIKE-direction itself: ✅. |
| 6 | **SELECT_ForecastDetailBCASN nondeterministic row order** (both base tables heaps, no `ORDER BY`) → the No-Ratio "first row + break" picks a nondeterministic assy for multi-row BCs. Rebuild action: add deterministic `ORDER BY ID_FORECAST_DETAIL` + confirm with David which assy the single-vehicle case should pick. (SELECT_ForecastDetailBCASN-analysis §3) | `computeAsnDetails` takes `fcRows[0]` for the No-Ratio branch (code.py:150), relying on `forecastByBc[bc]` order, which is exactly the proc's returned (heap-scan) order — code.py:288-300 appends rows in proc order with **no re-sort**. The `EXEC SELECT_ForecastDetailBCASN` call is **unchanged** (no `ORDER BY` added). | ❌ **GAP/RISK.** The rebuild inherits the legacy nondeterminism verbatim — "first row" is still allocation-order-dependent. The analysis's prescribed fix (`ORDER BY ID_FORECAST_DETAIL`) is **not applied** and the "which assy should the single-vehicle case pick?" question is **not yet answered by David**. For multi-row No-Ratio BCs the rebuilt qty can byte-match the golden by luck and silently diverge on a table reload/replan. Surface as a parity hazard requiring (a) a deterministic order and (b) David adjudication. |
| 7 | **Orders ≤ 5 No-Ratio:** first fc row only, `qty = VEHICLES * IN_ASSY_QTY` (no ratio), then `break`. (delphi-fanout §a; spec §4A) | `computeAsnDetails`: `if orders <= 5:` → `r = fcRows[0]`, `qty = vehicles * IN_ASSY_QTY` (code.py:148-151), append one row, `continue` (the Pascal `break`, code.py:160). | ✅ faithful (modulo the order-determinism hazard tracked separately in #6) |
| 8 | **Ratio branch qty** = `round(VEHICLES × IN_ASSY_QTY × IN_TIRE_RATIO / 100)` with Delphi **banker's** rounding; full qty (`VEHICLES × IN_ASSY_QTY`, no round) only when **both** ratios = 100; the multiply uses **tire ratio only** — wheel ratio participates *only* in the both-100 gate. (delphi-fanout §b; SELECT_ForecastDetailBCASN-analysis §2c, §2d) | `computeAsnDetails` ratio branch (code.py:163-177): `if tire == 100 and wheel == 100: qty = base` else `qty = _bankers_div_round(base * tire, 100)`; `base = vehicles * IN_ASSY_QTY`. Banker's rounding implemented exactly in integer arithmetic (`_bankers_div_round`, code.py:45-61) to be correct on Jython 2.7 (whose `round()` is half-away) AND Python 3. | ✅ faithful (tire-only numerator ✅; both-100 full-qty gate reads both ratios ✅; banker's-on-exact-rational ✅ — this is the highest-fidelity item) |
| 9 | **Manifest** = `'7' + copy(prodDate,4,5) + VC_ASSY_MANIFEST_NUMBER` = `'7'` + **1-digit year** + MM + DD + 2-char assy id (8 chars). (delphi-fanout §c; spec §6) | `_manifest()` (code.py:64-75): `"7" + productionDate[3:8] + str(assyManifestNumber)`; Python slice `[3:8]` = chars 4..8 = 1-digit-year+MM+DD. Length-8 assertion (`len != 8` → raise on the prodDate input). | ✅ faithful (1-digit-year slice correct; used in both branches: No-Ratio code.py:154, ratio code.py:173) |
| 10 | **INSERT_ASNInfo** — status hard-coded `'C'`; `@ASNID` via **SCOPE_IDENTITY()** OUTPUT (not `@@IDENTITY`/`IDENT_CURRENT`); EIN written at create as `fEIN+1`. (asn-write-chain §1; spec §2) | `create_asn` step 5 (code.py:318-328): `EXEC INSERT_ASNInfo @ASNID=@id OUTPUT … @Ein=?` with the OUTPUT captured via `DECLARE @id … ; EXEC … ; SELECT @id` on the SAME open tx through `runScalarPrepQuery(…, tx)`. Status `'C'` and SCOPE_IDENTITY() are inside the unchanged proc. `@Ein` passed as **0** (code.py:324). | ⚠️ status `'C'` ✅ + identity capture ✅; **EIN=0-at-create is the intended divergence** — see #11. |
| 11 | **EIN handling** — legacy stamps `fEIN+1` onto the header (and every detail) at create and bumps the ALC `Site.SiteEIN` counter inside the create click; send only flips status to `'S'`. (asn-write-chain §1; delphi-fanout §e; spec §2/§8) | Create writes `IN_ASN_EIN = 0` on header (code.py:324) and detail (`@EIN=0`, code.py:335). Docstring (code.py:249-254): real per-site EIN allocated atomically from the site sequence **at SEND** (M1 Rank 2). | ⚠️ **INTENDED DIVERGENCE.** Rationale: at-send EIN removes two real legacy bugs — the EIN-gap when the Inv rollback can't revert the out-of-txn ALC `SiteEIN+1`, and the read-then-bump race / unscoped `UPDATE Site SET SiteEIN+1` (no WHERE → bumps every site; a multi-site BLOCKER). The ASN_DETAIL parity diff must therefore **exclude `IN_ASN_EIN`** (0 here vs `fEIN+1` legacy). Flagged in spec §2/§8 and code.py docstring. Verify the parity harness ignores `IN_ASN_EIN`. |
| 12 | **INSERT_ASNDetail upsert** — `@HotCall=0` accumulate (`IN_QTY += @Qty`); `@HotCall=1` always-insert; **Q1 re-key to `(IN_ASN_ID, VC_MANIFEST_NUMBER)`** (legacy keyed on manifest ALONE → cross-ASN collision); preserve the within-ASN accumulate. (asn-write-chain §2; spike-asndetail-rekey.sql:41-91) | Re-keyed proc in `spike-asndetail-rekey.sql:67-74` (`IF EXISTS … WHERE IN_ASN_ID=@ASNID AND VC_MANIFEST_NUMBER=@Manifest` → UPDATE accumulate else INSERT). Driver calls `EXEC INSERT_ASNDetail @ASNID=?, … @Qty=?` with `@HotCall` defaulted 0 (code.py:333-335). | ✅ faithful (re-key applied; accumulate preserved; supporting composite index added spike-asndetail-rekey.sql:122-123) |
| 13 | **Positional VALUES / IDENTITY-skip fragility** — `INSERT … VALUES(8 values)` with no column list against the 9-col IDENTITY table; any column add/reorder (esp. M4 `IN_SITE_ID`) silently shifts mapping or throws. Rebuild SHOULD convert to explicit column-list INSERT. (asn-write-chain §2c, §6.6) | Re-key keeps the **positional VALUES** (spike-asndetail-rekey.sql:80-81, 87-88) — comments note the IDENTITY-skip and "M4: prepend @Site"; it does **not** convert to an explicit column list. | ⚠️→❌ Faithful to legacy ✅, **but the analysis's recommended hardening (explicit column list) was deliberately not done** → ⚠️ intended (keep parity-minimal for Q1; M4 markers in place) shading to ❌ latent risk: the M4 `IN_SITE_ID` add will break this positional INSERT exactly as warned unless the `@Site` prepend + WHERE edits land together. Surface as a tracked M4 hazard, not a today-blocker. |
| 14 | **SELECT_ASNSeq idempotency guard** — UI-level dedup (form locks + disables Create if a row exists for `(line, prodDate)` with `START_SEQ <> -1`); **no DB unique constraint**. Compare sentinel as string `'-1'` to avoid the legacy implicit varchar→int cast. (asn-write-chain §3; spec §1a) | `create_asn` step 1 (code.py:263-267): `EXEC SELECT_ASNSeq @LineName=?, @PDate=?`; if `len(existing)` → log + return `{"skipped": True}` (no write). Guard runs **before** the transaction opens. | ⚠️ faithful in intent (read-back guard → no-op idempotent skip instead of UI-lock) — reasonable for a headless driver. **Two residual notes:** (a) the proc is unchanged so the `VC_START_SEQ_NUMBER <> -1` implicit-cast lives on (analysis recommended `<> '-1'`); the driver doesn't touch it → ⚠️. (b) Still **no DB unique constraint** → concurrent creates can both pass the read-back and double-insert; the legacy single-user desktop never hit this, the gateway can. Surface as a concurrency risk (#16). |
| 15 | **The TWO cost checks (genuinely different).** PRE-loop (per-BC, pre-insert): hard ABORT on any `IN_MANIFEST_COST_ID IS NULL` (part has no cost master at all; **not date-aware**) → whole create fails. POST-loop (`SELECT_ASNMissingCost`, whole ASN, after inserts): non-aborting WARN, **date-windowed** (`VC_START_MANIFEST <= prodDate <= VC_END_MANIFEST`), CASE-splits "Missing Manifest Cost Entry" (no master at all) vs "out of date" (master exists but prodDate outside window). The out-of-date case is caught ONLY by the post-loop warn. (asn-write-chain §4, §4c; SELECT_ForecastDetailBCASN-analysis §4; delphi-fanout §d) | **PRE-loop** in pure `computeAsnDetails` (code.py:139-146): per BC, scans ALL fc rows; if any `IN_MANIFEST_COST_ID is None` → `raise AsnFanoutError`, **before the transaction opens** (code.py:305 runs the pure fn pre-tx) → nothing written. **POST-loop** in driver (code.py:347-358): `EXEC SELECT_ASNMissingCost @ASNID=?` after commit, WARN-only, wrapped in `try/except` that only logs (matches the Delphi swallow). | ✅ faithful — **both implemented and correctly distinguished.** PRE = NULL-cost-id hard abort, not date-aware, pre-write (code.py docstring :119-121 explicitly says "NOT the post-loop"). POST = the unchanged date-windowed proc, warn-only, never aborts. The two catch-sets are kept separate (the out-of-date population reaches only the post-loop). Minor note: legacy PRE is *per-BC before that BC's inserts* (so earlier BCs' rows exist until the txn rolls back) whereas rebuild runs the whole PRE-check **before any insert** — net all-or-nothing is the same; see #16. |
| 16 | **No surrounding transaction (legacy partial-ASN bug)** — each `ExecProc` auto-commits; a mid-loop abort leaves header + earlier-BC details persisted; the EIN bump runs **outside** the Inv txn on ALC. (asn-write-chain §0, §6.9; spec §0, §8) | `create_asn` wraps INSERT_ASNInfo + all INSERT_ASNDetail in **one** `system.db.beginTransaction` … `commit`/`rollback`/`close` (code.py:309-343); the pure fan-out (incl. PRE-loop abort) runs **before** the tx opens (code.py:305) so an abort writes nothing; EIN allocation moves in-txn at-send (#11). | ⚠️ **INTENDED DIVERGENCE (a fix).** Single-txn changes failure semantics from "partial ASN persists" to all-or-nothing; the at-send in-txn EIN closes the EIN-gap + unscoped-bump bugs. Both are improvements explicitly flagged (spec §7/§8, asn-write-chain §6.9). Sign-off should acknowledge this is **not byte-parity on the failure path** — a deliberately better behavior. (The DB-unique-constraint concurrency gap from #14 remains the one residual risk of the new model.) |
| 17 | **16-char audit stamp** — `VC_ADD`/`VC_LAST_UPDATE` = `yyyymmdd` + HH + mm + ss + **first 2 of the 3 millisecond digits** (style-114 quirk) = 16 chars, NOT 14. (asn-write-chain §1, §6.10) | INSERT_ASNInfo: the unchanged proc builds the stamp (16-char by the live body). INSERT_ASNDetail re-key: `@AddDate varchar(16)` with the same 4-substring CONVERT(112)+114 build (spike-asndetail-rekey.sql:52, 57-61) → emits 16 chars. **But the re-key file's header comment says "14-char yyyymmddHHmmss" (spike-asndetail-rekey.sql:25-26).** | ✅ faithful on the **emitted value** (varchar(16) + the 4-substring build = the 16-char stamp). ⚠️ **DOC DEFECT, not a behavior defect:** the re-key comment (and m1-spec §2) still say "14-char" — asn-write-chain §1 is the correction (it's 16). Fix the stale comments so the parity harness/reviewer isn't misled. Nothing parses the stamp back, so byte-parity is cosmetic-but-checked. |

---

## Consolidated ⚠️ intended-divergences (with rationale)

These are deliberate, defensible deviations from strict legacy byte-parity. Each needs David's explicit OK.

1. **EIN = 0 at create; real EIN at SEND** (#10/#11). Removes the legacy EIN-gap (out-of-txn ALC `SiteEIN+1`
   not reverted on Inv rollback) and the read-then-bump race / unscoped `UPDATE Site SET SiteEIN+1` (no
   WHERE = multi-site BLOCKER). **Consequence the harness must honor: exclude `IN_ASN_EIN` from the
   ASN_DETAIL/header parity diff** (0 here vs `fEIN+1` legacy).
2. **Single surrounding transaction** (#16). Failure semantics change from "partial ASN persists" to
   all-or-nothing; PRE-loop cost abort now runs before any write. A fix, not parity — failure-path output
   is intentionally not byte-equal to legacy.
3. **Idempotency guard = headless no-op skip** instead of the UI form-lock (#14). Same intent (no second
   ASN per line+prodDate); appropriate for a driver with no UI.
4. **Positional VALUES retained for Q1** (#13). Analysis recommended an explicit column list; kept
   positional to stay parity-minimal, with `-- M4:` prepend markers. Acceptable now; becomes a hazard at M4.
5. **16-char stamp is correct in code, but the re-key comment + m1-spec say "14-char"** (#17). The behavior
   is faithful (16); the *documentation* diverges and should be corrected.

## Consolidated ❌ GAPs / risks (real — surface for Verify + sign-off)

Ordered by how directly they threaten ASN-qty / row parity.

1. **char(3) vs char(21) BC-width source contradiction (#5) — UNRESOLVED.** `AD_FRSPULL-analysis.md` +
   `AD_FRSPULL-shared.sql:46/69` prove `convert(char(3), …)`; `delphi-fanout-confirmation.md §f` reads
   `convert(char(21), …)` from a *different* VehicleOrder.sql dump. Both are "verified" against different
   artifacts. **This must be adjudicated against the live running `AD_FRSPULL` before parity can be
   claimed** — the padding width is the operand the downstream `LIKE` matches on. (delphi-architect to
   reconcile the two dumps.)

2. **The `bc.strip()` in `create_asn` (#2, #5).** code.py:278 trims the BC before it is (a) used as the
   `forecastByBc` key and (b) fed as the `@BCode` LIKE left operand. The analysis says the trailing space
   is **semantic** (namespace separator + LIKE input). Stripping feeds neither char(3) nor char(21)
   semantics. **Risk:** spare-BC LIKE matches may diverge from legacy, and a hypothetical ground/spare BC
   collision silently merges fan-outs. Must verify the spare-BC end-to-end match against the live proc
   (delphi-fanout residual #3; SELECT_ForecastDetailBCASN-analysis §1b). Cannot be closed until #1 is.

3. **No deterministic ORDER BY on SELECT_ForecastDetailBCASN (#6).** The No-Ratio branch takes `fcRows[0]`
   (code.py:150) in nondeterministic heap-scan order. For multi-row No-Ratio BCs the chosen assy (and thus
   the shipped qty) is not reproducible. Analysis prescribes `ORDER BY ID_FORECAST_DETAIL` + a David
   decision on which assy the single-vehicle case should pick — **neither is done.** Today's parity may be
   luck-of-allocation.

4. **char(3) truncation alarm absent (#3).** Truncation is faithfully proc-resident, but the
   "assert/alarm if any BC component > 1 char" that §7.3 calls the single highest-severity latent defect is
   not implemented in the rebuild. A future multi-char GALC value silently produces a wrong BC.

5. **No DB unique constraint behind the idempotency guard (#14).** The read-back guard + single-user
   desktop hid this; on the gateway two concurrent `create_asn` calls can both pass `SELECT_ASNSeq` and
   double-insert. Add a unique index on `(line, prodDate[, site_id])` or otherwise serialize, since the
   create is now concurrency-exposed.

6. **M4 positional-INSERT hazard (#13).** Adding `IN_SITE_ID` to `INV_ASN_DETAIL_MST` will break the
   no-column-list positional VALUES unless the `@Site` prepend + WHERE edits land atomically. Tracked, not
   a today-blocker — but it must be on the M4 checklist.

7. **Stale "14-char" stamp comments (#17).** Doc-only, but corrected for the reviewer/harness:
   `spike-asndetail-rekey.sql:25-26` and `m1-asn-creation-spec.md §2` say 14; the true value is 16
   (`asn-write-chain-analysis.md §1`).
