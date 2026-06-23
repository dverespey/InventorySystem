# Adversary findings — Sites master + INV_SITES path-column ALTER (M4 piece 1)

Branch `m4-sites-master`. Single-site deployment (the 2026-06-22 direction reversal: each site = its own
gateway + DB; no `site_id` surgery; `INV_SITES` holds the ONE deployment's config in one row).

Scope reviewed: `spike-inv-sites-paths.sql` (7-column ALTER), `spike-inv-sites-table.sql` (INV_SITES),
`scripts/gen_sites_view.py` (shared SQL builders + save-path validation), `scripts/e2e/test_sites_master.py`,
the deployed view `…/perspective-views/Master/Sites/Sites/view.json` AND the gateway view, the 856/810
`_read_site` consumers, and the live spike DB (`Inventory` on `mssql-spike`). All probes were bounded,
READ-ONLY on parity refs, rolled-back on `Inventory`; spike left as-found (test reports seed=2 end=2).

---

## 1. ALTER non-breaking — CONFIRMED

**Claim:** adding the 7 path columns to `INV_SITES` cannot shift or break any consumer.

**Proof — every reader uses explicit column projection; ZERO `SELECT *` against INV_SITES anywhere:**
- Grep across `*.py`/`*.sql`/`*.json`: `SELECT * FROM INV_SITES` → **no matches**.
- The EDI trading-identity readers name their columns explicitly:
  - `docs/analysis/edi/856/project-library/edi856/code.py:404-406` — `SELECT VC_SITE_ABBR, VC_DUNS,
    VC_SUPPLIER_CODE, VC_TMM_DUNS, VC_EDI_MODE, VC_DELIVERY_METHOD_CODE, VC_SEP_ELEMENT, VC_SEP_SUBELEMENT,
    VC_SEP_SEGMENT FROM INV_SITES WHERE IN_SITE_ID = ?`
  - `docs/analysis/edi/810/project-library/edi810/code.py:428-431` — same 9 columns + `VC_DOCK_CODE`.
  - EIN allocation (856 `:477`, 810 `:647-ff`) — `UPDATE INV_SITES SET IN_EIN_SEQ=… OUTPUT … WHERE IN_SITE_ID=?` (named).
- The inbound/forecast readers project named columns:
  - `forecast/code.py:359` `SELECT IN_SITE_ID …`, `:380` `SELECT IN_HISTORICAL_FORECAST,
    BIT_USE_FIRST_PRODUCTION_DAY …`, `:754` `SELECT … WHERE VC_FORECAST_IMPORT_MODE='AUTO'`;
    `edi_inbound/code.py:294` `SELECT IN_SITE_ID …`.
- The Sites CRUD itself (list/get/insert/update) uses explicit column lists in `gen_sites_view.py`
  (`GRID_SCRIPT`, `recordid_onchange`, `INS_COLS`, `UPD_SET`, `GET_COLS`).

Adding nullable columns at the end of a table cannot affect any of these — none binds by ordinal, none
expands a star. **Non-breaking confirmed.**

**Idempotency — CONFIRMED on the live DB.** Each ALTER is `IF NOT EXISTS (sys.columns …) ALTER … ADD`;
the placeholder backfill is `COALESCE`-guarded (`WHERE … IS NULL`), so a re-run after a screen edit cannot
clobber a configured value. `test_sites_master.py` runs the DDL **twice** and both report `applied`
(`rc1=0 rc2=0`, PASS). Live state verified: all 7 columns present, `varchar(260) NULL`.

**Column types verified live** (`sys.columns`): the 7 path cols are `varchar(260) NULL` (Windows MAX_PATH
headroom); seeds are neutral `/opt/ignition/...` placeholders (rows 1/2). **No real client paths in the
repo** — confirmed.

---

## 2. Validation load-bearing + correct — CONFIRMED (with a real gap in piece 3)

The deployed gateway view's `save_script` enforces, BEFORE the write:
- `VC_EDI_MODE` exactly 1 char if set (`len(edimode) != 1` → reject);
- the 3 separators exactly 1 char each (`Separator (%s)` loop → reject);
- DUNS / TMM-DUNS = 9 or 13 digits (`_duns_ok`);
- fill-days ≤ 50, retention ≥ 12 (or blank/0 = NULL "unset");
- abbr ≤ 10, name/abbr required.

I verified the deployed gateway view actually contains all of these (`len(edimode)`, `Separator (%s)`,
`_duns_ok`, `_intn` all present in the gateway `view.json`), and that the guard validates the **same
`_str`-coerced value** that is bound to the INSERT (`edimode = _str(c.form_edimode)` and the insert arg is
`_str(c.form_edimode)`) — so there is no validate-one-value/store-another split-brain (a `' P'` strips to
`'P'` for both).

**The guard is genuinely the ONLY backstop for VC_EDI_MODE — PROVEN on the live DB.**
`VC_EDI_MODE` is `varchar(10)`, so the DB silently stores a 2-char value with NO error:
```
BEGIN TRAN; INSERT … VC_EDI_MODE='PP' …;  -- succeeds
  -> STORED_EDIMODE = 'PP', LEN = 2   (no Msg 8152, no CHECK)
ROLLBACK;
```
A malformed 4-char `'PROD'` placeholder is exactly the defect the 856 `_one_char` had to defend against
(`edi856/code.py:88`). On the spike the seed is correctly `'P'` (len 1, both rows). So the app guard is
load-bearing and is enforced in the deployed view.

**The other guards' backstops, confirmed live (rolled back):**
- char(1) separators DO get a DB backstop: 2-char `VC_SEP_ELEMENT='**'` → `Msg 8152` (opaque truncation error).
- fill-days=60 → `Msg 547 CK_INV_SITES_FILL_DAYS`. retention=6 → `Msg 547 CK_INV_SITES_DATA_RETENTION`.
- DUNS has NO DB backstop (`varchar(13)` accepts `'12345'`); the app `_duns_ok` is the only enforcement.

**Field that should be 1-char but is NOT DB-constrained:** `VC_EDI_MODE` (`varchar(10)`). Acceptable as a
schema choice ONLY because the app guard covers it — and it does. No other "should-be-1-char" field is
unguarded (the separators are `char(1)`).

> SHOULD-FIX (test/parity-method): see Finding B — the test never drives a bad EDI-mode/separator/DUNS
> through the **view's** save_script; it checks an independent oracle + the DB-doesn't-backstop fact, but
> the actual deployed guard's rejection is unproven by automated test.

---

## 3. CRUD round-trip — CONFIRMED faithful + non-vacuous

`test_sites_master.py` builds the INSERT/UPDATE/GET from the **shared builders** (`gsv.INS_COLS/INS_VALS/
UPD_SET/GET_COLS/INS_ORDER`) the view emits — same column list/order, only `?` vs inline literal differs —
and runs them in a rolled-back transaction. Live run: **20 PASS / 0 FAIL**. Specifically:
- INSERT stored every field type incl. all 7 paths exactly as set (PASS); the 7 paths spot-checked
  individually (PASS).
- UPDATE stored the edited fields (name/city/edimode/paths/duns/retention) (PASS) AND **non-vacuity** holds
  (`'Roundtrip Test Site'→'Updated Site Name'`, edimode `'T'→'P'`, path/duns changed) — a no-op save would fail.
- Rule #3 honored: `IN_EIN_SEQ` defaulted to literal 0 on insert (never form-written).
- Restore-as-found: `seed=2 end=2`.

The "reverted column → NULL → fails" non-vacuity the brief asks for is structurally present: the round-trip
asserts each column equals the set value, and a dropped column from the INS/UPD list would read back NULL
and fail the assertion.

---

## 4. Single-site correctness — CONFIRMED

- The deployed view's list/get/save carry **NO `site_id`/`IN_SITE_ID` filter** (RULE #2: "deliberately NO
  site_id predicate" present in the grid script; get/update key only on the `IN_SITE_ID = ?` PK). A one-row
  deployment therefore lists/edits/saves its single row correctly.
- The only `site_id = ?` predicate anywhere in the scripts is the delete-gate refCount against
  `INV_PARTS_STOCK_MST.site_id` (counting parts that reference the site being deleted) — correct, not a
  multi-tenant scoping leak.
- The EDI dirs (EDIOut/EDIIn) are documented as a shared-share **config value**, not a schema flag — fine.
- No lingering multi-site assumption that breaks single-site was found in the reviewed artifacts.

---

## Findings

### BLOCKER
None against the deployed/runtime artifacts. (The committed-repo staleness below is provenance, not runtime.)

### SHOULD-FIX

**A. The committed repo `view.json` is STALE — it predates M4 piece 1 entirely (provenance/reviewability).**
`docs/analysis/master-data/perspective-views/Master/Sites/Sites/view.json` (committed in c1c51c2,
2026-06-19) does **NOT match** the current `gen_sites_view.py` output. The committed copy is missing,
in its save/get scripts:
- all 7 path columns (`VC_EDIOUT_DIR…VC_TEMPLATE_DIR` absent from its INSERT/UPDATE/GET — `form_ediout`
  never assigned);
- the load-bearing ISA/DUNS guards (`len(edimode)`, `Separator (%s)`, `_duns_ok` all absent);
- the `_intn` nullable-unset coercion and the retention 1..11 refinement.

The **generator and the GATEWAY-deployed view are CORRECT and current** (I verified the gateway
`view.json` at `/usr/local/ignition/data/projects/spike/…/Master/Sites/Sites/view.json` contains the path
cols, the EDI-mode guard, the separator guard, `_duns_ok`, and `_intn`). The generator only writes to the
gateway path; the repo `perspective-views/` copy was never re-synced after the M4 piece 1 changes. This is a
**code defect in repo hygiene, not a runtime defect** — but per the master-crud header "view = truth, this =
synced spec," a reviewer reading the committed view would wrongly conclude M4 piece 1 is unbuilt, and a
re-deploy FROM the committed copy would ship a Sites screen that silently drops every path column and the
load-bearing ISA validation. Re-run `gen_sites_view.py` and commit the regenerated `view.json` (and have
the generator also write the repo copy, not only the gateway path), so the reviewable artifact == runtime.
This is a **test/parity-method + provenance flaw**, fixable.

**B. The load-bearing guard (EDI-mode/separator/DUNS) is verified by code-inspection, not by test.**
`test_sites_master.py` proves three true-but-insufficient things for the ISA/DUNS cases: (1) the
*independent source-truth oracle* (`source_truth_reject_reason`, transcribed in the test) flags the bad
value; (2) the DB does NOT backstop a 2-char EDI mode; (3) the char(1) separator DB-backstops. It never
runs the **view's actual `save_script`** on a bad EDI-mode/separator/DUNS and asserts rejection. The
browser test `test_sites_crud.py` covers blank-name/blank-abbr/fill_days>50/retention<12 + the delete-gate,
but **not** the ISA/DUNS guards through the live form. So the single most consequential guard in this piece
(operator typo `'PROD'` → malformed ISA15 to TEMA) has no automated proof that the *deployed view rejects
it* — only that the rule is correct and the DB won't save us. Add a browser case (or a Jython-driver case
against `save_script`) that types a 2-char EDI mode / 2-char separator / 5-digit DUNS into the form, clicks
Save, and asserts `SPIKE … save REJECTED: edimode/sep/DUNS` fired + no DB write. (Pairs with the standing
Jython-driver-coverage-gap memory: pure-oracle + DB ≠ the gateway-side guard.) Note this is the same
self-referential-test trap pattern — the test's oracle is independent (good), but the *enforcement path
under test* is not exercised at all.

### NIT

**C. `_duns_ok` uses Python `unicode.isdigit()`, which accepts non-ASCII digits.** A DUNS of 9 Arabic-Indic
or superscript "digits" passes `d.isdigit() and len(d) in (9,13)` yet is not an ASCII DUNS and would emit a
non-ASCII ISA08/GS. Purely theoretical (DUNS is operator-typed ASCII at cutover; real values load from the
legacy config), so NIT. Tighten to an ASCII-digit check (`all('0'<=ch<='9' …)` or a regex) if hardened.

---

## VERDICT

**The INV_SITES path-column ALTER is SOUND and PROVEN non-breaking** (every consumer projects named
columns; zero `SELECT *`; idempotent re-run verified live; types `varchar(260) NULL`; placeholder seeds
only).

**The validation is load-bearing and CORRECTLY ENFORCED in the deployed gateway view** (EDI-mode 1-char,
3 separators 1-char, DUNS 9/13, fill-days ≤50, retention ≥12), and I PROVED on the live DB that the
`varchar(10)` VC_EDI_MODE has no DB backstop — so the app guard is genuinely the only thing preventing a
malformed positional ISA15, and it is present and bound to the same coerced value it validates.

**The CRUD round-trip is faithful and non-vacuous** (20/20 PASS, every field incl. the 7 paths round-trips,
insert/update non-vacuity holds, IN_EIN_SEQ stays system-managed, spike restored as-found).

**Single-site correctness holds** (no site_id predicate on list/get/save; the lone `site_id=?` is the
delete-gate refCount).

For a **single-site deployment the rebuilt Sites master + path-column ALTER is sound — at the RUNTIME
(generator + gateway) layer.** Two non-runtime caveats keep it short of "fully proven from the repo alone":
(A) the **committed repo `view.json` is a stale Jun-19 snapshot** that lacks the path columns AND the
load-bearing ISA validation — fix repo hygiene so the reviewable artifact equals runtime, and never
re-deploy from the stale copy; (B) the **load-bearing ISA/DUNS guard is proven by code-inspection, not by
an automated test that drives the deployed view** — add that test before cutover. Neither is a divergence in
the live behavior; both are gaps in what the repo+tests independently PROVE.
