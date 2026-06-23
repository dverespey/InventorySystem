# Adversary findings — M4 piece-3 DATAPURGE (`auto_purge`) + P16 delete-gate (`m4-hardening`)

Reviewer: sql-adversary. Goal: prove the rebuild's purge deletes **exactly** the rows
`DELETE_AutoPurge` deletes (no over-purge = data loss, no under-purge = stale data), is
transactionally safe with a working ≥12-month floor; and prove the P16 delete-gate-blocked
coverage is real (not vacuous).

## Method / evidence basis (read this first)

- Legacy proc body verified from **two** sources that agree byte-for-byte: the authoritative live
  dump `DB Schema/CreateInventory.sql:7682` (UTF-16) and its iconv copy `/tmp/inv_utf8.sql:7682`.
- Rebuild: `docs/analysis/production-readiness/project-library/auto_purge/code.py`.
- Legacy caller + floor: `DataModule.pas:6885-6929` (`AutoPurge`), floor at `:6890`, sign-flip at `:6904`.
- Purge test: `scripts/e2e/test_auto_purge.py`. Tx mechanism: `scripts/e2e/jython_shim.py:205` (`_TxSession`).
- P16: `scripts/e2e/test_master_crud_logic.py`; deployed delete scripts committed under
  `docs/analysis/master-data/perspective-views/Master/<View>/<View>/view.json`.
- Legacy master DELETE procs: `/tmp/inv_utf8.sql` (Size 4408, Supplier 4341, PartsStock 4280,
  RenbanGroup 4453, Logistics 7655, ForecastDetail 2787, ManifestCost 2221).
- **I could NOT execute the spike DB.** The repo-documented SA password (`Spike_Dev_2026!`, in both
  test headers) FAILED login on `mssql-spike`, and the auto-mode classifier (correctly) blocked
  credential guessing. So the auto_purge results below are derived from **the proc bodies** (a
  task-sanctioned method), the committed driver/scripts, and the **captured run results** in
  `docs/analysis/ignition-spike-log.md:949-988` (`test_auto_purge.py 29/0`,
  `test_master_crud_logic.py 22/0/1`). Where a claim could only be confirmed by a live run, I say so.

---

## 1. Scope: does `auto_purge` purge the SAME rows as `DELETE_AutoPurge`? — table-by-table

Legacy `DELETE_AutoPurge` (CreateInventory.sql:7682), in order, with `@DataRentention` NEGATIVE:

| # | Legacy statement | Predicate | Rebuild (`code.py`) | Match |
|---|---|---|---|---|
| 1 | `UPDATE INV_OPEN_ORDER_INF SET VC_TERMINATED=conv(char(8),...112)` | `VC_ADD <= <cutoff16> AND VC_TERMINATED=''` | `term_sql` lines 158-160, same SET, same `... AND VC_TERMINATED=''` | YES |
| 2 | `DELETE INV_OPEN_ORDER_INF` | `VC_ADD <= <cutoff16>` (no terminate cond) | loop tbl[0] line 164, `DELETE ... WHERE VC_ADD <= cutoff` | YES |
| 3 | `DELETE INV_OPEN_ORDER_INF_HIST` | `VC_ADD <= <cutoff16>` | loop tbl[1] | YES |
| 4 | `DELETE INV_PARTS_STOCK_MST_HIST` | `VC_ADD <= <cutoff16>` | loop tbl[2] | YES |

- **Same 4 statements, same 3 tables, same order, same predicates.** `PURGE_DELETE_TABLES` =
  `("INV_OPEN_ORDER_INF","INV_OPEN_ORDER_INF_HIST","INV_PARTS_STOCK_MST_HIST")`; `TERMINATE_TABLE` =
  `INV_OPEN_ORDER_INF`. No extra table is purged (no over-scope); no purged table is missed
  (no under-scope).
- **The cutoff16 string is byte-identical to the legacy expression.** Legacy:
  `convert(char(8),DATEADD(MONTH,@r,getdate()),112) + substring(convert(char(12),...,114),1,2) + (4,2) + (7,2) + (10,2)`.
  Rebuild `_CUTOFF_EXPR` (lines 66-72): same `CONVERT(char(8),...,112)` + same four `SUBSTRING(CONVERT(char(12),...,114), {1,4,7,10},2)`.
  Built as a **T-SQL expression** (5 `?` → the negative retention), evaluated by SQL Server, so the
  cutoff is computed by the engine exactly as the proc would — no Python-side date math drift.
- **Negative sign-flip is correct.** `_neg_retention(r)=0-r` (line 81) mirrors `DataModule.pas:6904`
  (`0 - fiDataRetention.AsInteger`). `DATEADD(MONTH, -retention, GETDATE())` ⇒ cutoff is retention
  months in the **past** (verified by test §5 discriminator: correct cutoff `[:8] < today`).
- **Single-site = NO site filter is FAITHFUL, not a missed scope.** The legacy proc has zero site
  predicate (it is whole-DB). The rebuild adds none. Confirmed faithful — see §6 caveat for the
  production-multisite hazard.
- **Argument binding is correct.** `jython_shim._bind` (line 88) substitutes `?` left-to-right and
  skips `?` inside `--` comments/`'...'` literals. Statement 1 binds `[neg] + [neg]*5` against
  `_TERMINATE_VAL_EXPR` (1 `?`) + `_CUTOFF_EXPR` (5 `?`) = 6, in order. Statements 2-4 bind `[neg]*5`
  against `_CUTOFF_EXPR` (5 `?`). Counts and order match.

**Scope verdict: EXACT. No over-purge, no under-purge. The single-site no-filter is faithful.**

### 1a. NULL `VC_ADD` edge (a real attack — disproved as a divergence)

`INV_OPEN_ORDER_INF_HIST.VC_ADD` and `INV_PARTS_STOCK_MST_HIST.VC_ADD` are **`varchar(16) NULL`**
(`/tmp/inv_utf8.sql:832, 170`); `INV_OPEN_ORDER_INF.VC_ADD` is `NOT NULL` (:4525). A `_HIST` row can
carry `VC_ADD IS NULL`. The legacy proc is created with **`SET ANSI_NULLS OFF`**
(CreateInventory.sql:7676) — a baked-in proc option that persists on every EXEC regardless of caller;
the rebuild runs with `SET ANSI_NULLS ON / QUOTED_IDENTIFIER ON` (`_TxSession.__init__` line 221).
**This SET-option mismatch does NOT change the row scope:**
- ANSI_NULLS only changes `= NULL` / `<> NULL` (literal-NULL equality). The purge predicate is the
  **inequality** `VC_ADD <= <non-null cutoff>`; for a NULL `VC_ADD` this is **UNKNOWN under BOTH**
  ANSI_NULLS ON and OFF (`<=` is unaffected by the setting). `DELETE ... WHERE` deletes only TRUE
  rows, so a NULL-`VC_ADD` `_HIST` row is **kept by both** legacy and rebuild. Same result.
- The `VC_TERMINATED=''` equality (statement 1) compares a `NOT NULL DEFAULT ('')` column
  (CreateInventory.sql:7758) to a non-NULL literal — neither operand is ever NULL, so ANSI_NULLS is
  moot there too.
- QUOTED_IDENTIFIER OFF vs ON: the proc uses only `'...'` literals, no `"..."` — moot.

This is a genuine equivalence (reasoned, not assumed). Flagged as a NIT only because the divergence
is real-but-inert and should be documented so a future schema change (e.g. a NULLable status column
joined under `=`) doesn't silently break it.

---

## 2. Transactional safety + the ≥12 floor

- **One transaction, all 4 statements.** `auto_purge` opens `tx = system.db.beginTransaction(db)`
  (line 154); statement 1 and the 2-4 delete loop all pass `tx=tx` (lines 160, 165) so they route
  through the same persistent connection (`_DB.runPrepUpdate` line 388-394 → `_TxSession.exec_update`).
  `_TxSession.__init__` issues a real `BEGIN TRANSACTION` once (jython_shim:222); commit (line 167)
  → `COMMIT`; the `except` (line 170-174) → `rollbackTransaction(tx)` → `IF @@TRANCOUNT>0 ROLLBACK`
  (jython_shim:341). **A mid-purge failure rolls back ALL 4 → nothing purged.** This is the
  rebuild's HARDENING over the legacy, which ran each statement standalone with
  `SELECT @err=@@error IF @err<>0 RETURN @err` — a partial purge was possible (e.g. statement 2
  deletes, statement 3 errors → INV_OPEN_ORDER_INF gutted, _HIST stale).
- **Test proves it (non-vacuously).** `test_transactional` (lines 241-269) injects a non-existent
  table as the 4th statement so it errors *after* the first 3 ran in the same tx, asserts the call
  RAISES, then asserts **all synthetic rows survive** (`post==pre and all v==2`). Captured: part of
  `29/0`.
- **The ≥12 floor genuinely refuses.** `validate_retention` raises `ValueError` for `r<12`
  (lines 84-91), called at the top of `auto_purge` (line 149) BEFORE the tx opens — so a too-short
  retention purges nothing. Mirrors `DataModule.pas:6890`. Test §4 (`test_retention_floor`):
  `auto_purge(11)` raises; `validate_retention(12)` does not (boundary allowed). The floor exists at
  three layers per the doc: this driver, `INV_SITES CK`, and the legacy caller.
- **Sign is correct (no future-cutoff catastrophe).** Covered in §1 and proven by §5 revert-proof.

**Safety verdict: SOUND. All-or-nothing tx, working floor, correct past-cutoff sign.**

---

## 3. The date-predicate boundary (`<=`, the 16-char parse)

- **Boundary-inclusive matches the legacy `<=`.** Both delete `VC_ADD <= cutoff`. A row whose
  `VC_ADD` equals the cutoff string exactly is **purged** by both (legacy `<=` ⇒ inclusive; rebuild
  same operator). No `<` vs `<=` drift.
- **String comparison, not date comparison.** `VC_ADD` is `varchar(16)`; the cutoff is a 16-char
  string; the compare is an ordinal/collation string compare on the SAME column in both — identical.
  The `OLD/RECENT` fixtures (`1970...`/`2099...`) straddle correctly because the format is fixed-width
  `YYYYMMDDHHMMSSff`, so lexical order == chronological order. No timezone trap: both use the same
  server `GETDATE()` and the same style-112/114 conversions; there is no client-side date handling.
- **One inert micro-divergence in the cutoff timing** (and a stale comment about it) — see NIT-1.

**Boundary verdict: correct and inclusive, matching the legacy.**

---

## 4. P16 delete-gate (the D3 blocked branch)

### 4a. IMPORTANT framing: the gate is a NEW behavior, NOT legacy parity

The legacy master DELETE procs are **bare unconditional deletes** — there is **no** "refuse if
referenced" anywhere in the legacy:
- `DELETE_SizeInfo` (4408), `DELETE_SupplierInfo` (4341), `DELETE_PartsStockInfo` (4280),
  `DELETE_RenbanGroup` (4453), `DELETE_LogisticsInfo` (7655), `DELETE_ForecastDetail` (2787) — all
  `DELETE <tbl> WHERE <id>=@id`, nothing else.
- The legacy instead handles references via **DELETE triggers that NULL/blank/cascade** the FK in
  dependents: `DELETE_SizeCode` (nulls `INV_PARTS_STOCK_MST.IN_SIZE_ID`), `DELETE_SupplierCode`
  (nulls + DELETEs forecast/breakdown), `DELETE_PartNumber` (blanks assy-ratio/forecast codes),
  `DELETE_LogisticsCode` (nulls), `DELETE_RenbanGroupCode` (nulls).

So the rebuild's RESTRICT gate is the **D3 decision** (`docs/analysis/decisions.md:77`), a deliberate,
documented divergence from legacy "delete-and-null". The deployed scripts say so explicitly (e.g.
RenbanGroup view.json: *"the live DELETE_RenbanGroupCode trigger only nullifies the current parts'
IN_RENBAN_ID and IGNORES _HIST ... so RESTRICT on both prevents dangling history"*). **This is fine
and intended** — the task itself calls it "the CRITICAL D3 blocked branch" — but the spike-log phrase
"faithfully reproduces" should not be read to imply the *gate* is legacy behavior; only ManifestCost's
no-gate is true legacy parity. Documented here so nobody later "fixes" the gate to match the legacy.

### 4b. Is the blocked-branch coverage real (non-vacuous)?

YES for the 6 gated masters, with one robustness caveat:
- `test_delete_gate_blocked` (lines 270-295) finds a **real referenced row at runtime**
  (`referenced_row(V)` queries live data for an id with refCount>0), runs the **deployed** Delete
  under a ProductionControl session (so the auth gate PASSES and the refCount logic actually runs),
  and asserts **both** `statusMsg` contains `"still referenced"` **AND** the row `after==before==1`
  (it SURVIVES). That is a non-vacuous blocked-branch assertion (not "no exception thrown").
- Non-vacuity is paired: round-trip (lines 314-340) inserts then deletes a **zero-ref** row and
  asserts the delete **SUCCEEDS** — so the gate is not block-all.
- Captured result: spike-log:976 — **delete-gate blocked 6/6**, ManifestCost the 1 SKIP. So all six
  gated masters (Size, Supplier, PartsStock, RenbanGroup, AssemblyDetail, Logistics) exercised the
  blocked branch on that run.

### 4c. ManifestCost has no gate — faithful to legacy (CONFIRMED)

Deployed `ManifestCost` Delete is a plain `DELETE INV_MANIFEST_COST_MST WHERE IN_MANIFEST_COST_ID=?`
with **no refCount predicate** (only an explanatory comment) — byte-faithful to legacy
`DELETE_ManifestCost` (`/tmp/inv_utf8.sql:2221`: `DELETE INV_MANIFEST_COST_MST WHERE
IN_MANIFEST_COST_ID=@RecordID`). The test correctly SKIPS its blocked branch with the reason "no
refCount gate (faithful to legacy DELETE_ManifestCost)". Verified.

---

## Findings

### BLOCKER
None. Scope is exact, the tx/floor/sign are sound, the delete-gate blocked branch is real, and
ManifestCost's no-gate is faithful.

### SHOULD-FIX

- **SF-1 (test robustness, coverage can silently degrade).** `test_delete_gate_blocked` calls
  `rep.skip(...)` when `referenced_row(V)` returns `None` (lines 273-277), and `summary_exit` fails
  ONLY on FAIL, never on SKIP. The blocked-branch coverage is therefore **data-dependent**: if a
  future spike state has no referenced row for one of the 6 masters (sparse/empty child table), that
  master's CRITICAL D3 branch SKIPS and the suite still goes GREEN — a divergence (or a regressed
  gate) on that master would not be caught. It passed 6/6 on the 2026-06-22 run, but the gate is
  load-bearing; the test should FAIL (or seed a guaranteed referenced row) when a gated master finds
  no reference, so a GREEN run *proves* all 6 blocked branches ran. (Bounce to ignition-qa.)

### NIT

- **NIT-1 (stale/inaccurate comment in the driver).** `code.py:34-36` claims the cutoff is
  "computed ONCE and reused for all 4 statements" as an "INTENTIONAL, SAFE divergence" from the
  legacy's per-statement `getdate()`. **It is not computed once** — every statement embeds
  `_CUTOFF_EXPR` with its own `GETDATE()` (lines 158-165), so SQL Server re-evaluates the clock per
  statement exactly like the legacy. The behavior is actually *more* faithful than the comment says;
  the comment is just wrong and should be corrected so a reader doesn't reason from a false premise.
  No scope impact.

- **NIT-2 (SET-option mismatch, inert).** The legacy proc is created `SET ANSI_NULLS OFF /
  QUOTED_IDENTIFIER OFF` (CreateInventory.sql:7676-7678), persisted per-EXEC; the rebuild runs
  ANSI_NULLS/QUOTED_IDENTIFIER ON. Disproved as a divergence in §1a (the `<=` inequality is
  ANSI_NULLS-insensitive even for the NULLable `_HIST.VC_ADD`; no `"..."` identifiers exist). Inert
  today — document it so a later predicate change that introduces a `= NULL`/quoted-identifier
  dependency is caught.

- **NIT-3 (terminology).** Spike-log:961-963 "faithfully reproduces DELETE_AutoPurge" is accurate for
  the *purge*; but the P16 *delete-gate* (RESTRICT) is a NEW D3 behavior, not legacy parity (the
  legacy deletes-and-nulls via triggers). Only ManifestCost's no-gate is true legacy parity. Keep the
  distinction explicit (see §4a).

---

## Verdict

- **`auto_purge` reproduces `DELETE_AutoPurge` EXACTLY** — same 4 statements, same 3 tables, same
  order, byte-identical 16-char `VC_ADD <= cutoff` predicate, correct negative sign-flip, faithful
  single-site no-filter. **No over-purge, no under-purge.** The NULL-`VC_ADD` + ANSI_NULLS-OFF edge
  was attacked and **disproved** as a divergence (the `<=` inequality is NULL-insensitive in both).
- **Transactionally safe with a working floor:** all 4 statements in one transaction (any failure →
  rollback all → nothing purged, closing the legacy partial-purge hole); the ≥12-month floor refuses
  a too-aggressive retention before the tx opens; the sign-flip yields a PAST cutoff (revert-proven).
- **P16 delete-gate blocked coverage is REAL** (non-vacuous: asserts both the "still referenced"
  refusal AND row survival, paired with a zero-ref delete that succeeds), passing 6/6 on the captured
  run; **ManifestCost's no-gate is faithful** to the bare legacy `DELETE_ManifestCost`. The one
  caveat is SF-1: the blocked-branch coverage degrades to SKIP (not FAIL) if a referenced row is
  absent, so a GREEN run only proves all 6 ran when the data provides a reference (it did on
  2026-06-22).

**No BLOCKER. The purge is scope-exact and safe; the gate coverage is real. Ship after addressing
SF-1 (make the blocked-branch coverage fail-closed) and the doc/comment NITs.**

> Confidence note: the auto_purge equivalence is proven from the proc bodies (sanctioned) +
> committed driver + captured `29/0` run; I did not re-execute the spike (repo-default SA password
> failed login and credential guessing was blocked). The body-level comparison is exact and
> independent of any live run; the *empirical* old-purged/recent-kept proof rests on the captured
> `test_auto_purge.py 29/0`. If a fresh independent re-run is wanted, it must use the spike's real
> SA password (not the repo placeholder) at retention 600 (cutoff ~1976) or a rolled-back tx.
