# InventorySystem Two-Repo Split — Concrete Plan

**Status:** PLAN (review-ready). Not executed. Direction is DECIDED
(memory `project-system-landscape`, 2026-06-23). This document synthesizes four
classification manifests (ignition-appcode, perspective-ui, harness-db,
delphi-analysis-config) into an executable plan.

**The split, in one line:** extract the InventorySystem **Ignition app** (Jython
app code + Perspective views + e2e harness + DDL/fixtures + generators) out of the
current Delphi repo into a **new `inventory` repo**, started **fresh** (no history
rewrite). The current repo **stays** as the legacy Delphi source + reverse-engineering
analysis home (`docs/analysis/**` `.md` specs that cite Delphi `file:line`).

**The central difficulty (all four manifests agree):** the Ignition app code is NOT
in a clean directory. It is physically buried under `docs/analysis/**/project-library/`
and `docs/analysis/**/perspective-views/`, interleaved one-to-two levels beneath the
spec `.md` files — and the e2e/generator tooling reaches into those exact paths by
**hardcoded relative path**. So the cut is a *path lift with tooling-path rewrites*,
not a clean directory move.

---

## Adversarial review — corrections applied (2026-06-24)

This revision folds in the adversary's findings. **Verdict unchanged: READY with
prerequisites** — but the move-set is now *reference-derived* (grep of what the
harness/app/generators actually open/import at runtime), not directory-name-based.

## PREP DONE + S1/R3 WIDENED (2026-06-24, PR #56) — two findings from doing the prep

**PREP 1 (centralize constants) DONE:** the gateway-path + DB-connection are now single-point.
CPython → `scripts/_ignenv.py` (`IGN_DB_CONN`/`IGN_PROJECT`/`GATEWAY_*`, defaults = the spike
values; the PROD-rename edit point is `IGN_DB_CONN=Inventory`). Jython app code → a new
`project-library/db_shared` module (`CONNECTION="Inventory_Spike"`; the one-line prod edit →
`Inventory`). 47 files rewired, behavior-preserving (e2e green, no straggler hardcode), the
two single-points proven (`IGN_DB_CONN=Inventory` flows through). **PREP 2 (Admin/Users
attributes) DONE:** committed the live-gateway-signed `resource.json` as the recovery source.

**★ S1/R3 WIDENED — the cold-start `attributes`-NPE is RESOURCE-TYPE-AGNOSTIC, not view-only.**
`ProjectResourceBuilder.setAttributes()` NPE-bricks the WHOLE gateway on cold start for ANY
seeded `resource.json` with a null `attributes` map — **scripts too**, not just views (proven:
`db_shared/resource.json` faulted the gateway until given an `attributes` block). An audit found
**14 of 32 repo `resource.json` lack `attributes`** — all 8 master views + Home/MasterHub/BackBar/
page-config + the auth/auto_purge script modules. So S1/R3 widens from "the Admin/Users view" to
**EVERY seeded `resource.json` must carry a valid non-null `attributes` in the fresh repo.** The
fix: generalize `gen_user_admin_view.py`'s `_repo_resource_json()` (prefer the live-gateway-signed
map → fall back to a valid sig → last-resort a non-null all-zeros placeholder that still avoids
the NPE) across the seed step. **Add a CUT GATE (§7): a cold-start smoke + a grep proving ZERO
seeded `resource.json` has a null `attributes` before the fresh gateway is started.**

**Stale-generator debt (pre-existing, NOT split-blocking):** `gen_sites_view.py` +
`gen_master_write_gates.py` are stale vs the post-#51/#53 live views (re-running them would
regress the Sites form to hardcoded colors/NewButton + the gate-check raises NewButton-not-found).
The cut MOVES the COMMITTED views (canonical), not regenerates them, so this doesn't block — but
re-sync those generators to the current view structure as a follow-up (punch-list).

- **B1 (BLOCKER, resolved):** the MOVES list was built by directory name and left
  **live runtime dependencies** behind in `docs/analysis`. The authoritative MOVES set
  is now derived by grep (see §2, "Reference-derived MOVES — the grep authority").
  Newly-caught runtime deps: `docs/analysis/reporting/sql/*.sql` (7 files, loaded by
  `report_render/driver.py` + `test_m3_reports.py`), the three `spike-*-feed.sql`
  drift-guard files, `spike-report-procs-d6.sql`, `master-crud-namedqueries.sql`, the
  sites `spike-inv-sites-*.sql`, and the order parity tooling
  (`docs/analysis/order/spike/parity_diff.py` + `scripts/gen_parity_tsv.sh`). The §1
  reorg gate now REQUIRES a grep proving zero surviving cross-boundary path reads.
- **B2 (BLOCKER, resolved):** "all generators are env-overridable" was FALSE. Only **2**
  use `os.environ.get("PROJ_DIR", …)`; **8** hardcode the gateway path with no fallback.
  §4.C + centralize-step now describe per-file env-parametrization surgery on those 8.
- **S1 (resolved):** the `Admin/Users/Users` repo copy has NO `attributes` and there is
  NO `docs/design` twin — the only in-repo seed NPE-faults the whole gateway on cold
  start. The valid re-signed `attributes` map is now a REQUIRED move-set input, sourced
  from a **live gateway export** (§2 + R3 + sequence).
- **S2 (resolved):** R5 undercount (5 → **8** e2e files read `/usr/local/ignition/…`)
  and the path-DEPTH coupling: **27** e2e files climb `"..","..","docs"`, hardcoded to
  the `scripts/e2e/` depth. The reorg step now fixes BOTH the path TAIL and the
  relative HEAD (`../..`) in the same commit.
- **S3 (resolved):** PR #22 (order-seam-runner) was OPEN at draft time; it is now
  **MERGED** (commit f86e03e, 2026-06-24). The prereq gate is kept as a *verify* step.
- **N1/N2 (counts fixed):** `Inventory_Spike` appears in **60 files** (32 live `.py`:
  13 app modules + 15 e2e incl. shim + 4 generators), not "17+". `spike-*.sql` live in
  TWO roots (`scripts/` + `docs/analysis/**`); the sweep covers both. The obsolete
  `scripts/spike-vehicleorder-line-fixture.sql` DROP is kept.

---

## 1. The new `inventory` repo — proposed structure

Mirror the Ignition gateway's project-as-code layout 1:1 so the Designer↔Git
round-trip is direct (no more 3-way snapshot drift). Proposed top-level:

```
inventory/
  CLAUDE.md                      # Ignition-oriented (authored fresh, NOT copied from Delphi)
  README.md                      # Ignition app overview + dev bring-up
  .gitignore                     # Python/Ignition rules (Delphi build section dropped)
  .github/workflows/             # CI: provision DB -> apply migrations -> run e2e -> deploy 8.1.52->8.3

  ignition-project/              # === GATEWAY PROJECT-AS-CODE (Designer<->Git round-trip) ===
    com.inductiveautomation.perspective/
      views/
        Home/Home/
        Shell/{BackBar,ComingSoon}/
        Master/{MasterHub,AssemblyDetail,Logistics,ManifestCost,PartsStock,
                RenbanGroup,Sites,Size,Supplier}/
        Order/{OrderLanding,RenbanBreakdown,RenbanCollisionDialog}/
        HotCall/HotCallEntry/
        Reports/ReportsStub/
        Admin/Users/Users/
      page-config/config.json    # the 16-route AUTHORITATIVE config (design/gateway one)
      session-props/
    ignition/
      named-query/Home/          # landing KPI Named Queries
      script-python/             # === JYTHON APP CODE (project-library modules) ===
        edi810/ edi856/ edi_inbound/ forecast/ asn/ hotcall/
        receiving/ reject/ shipping/ stockLedger/ stocktaking/
        order/ renban/ order_file/ forecast_distribution/
        auth/ auto_purge/ report_render/
        _shared/                 # NEW: centralized DB-connection constant (see 4.D)

  generators/                    # scripts/gen_*.py + build_fonts_css.py, retargeted to ignition-project/
  e2e/                           # scripts/e2e/** harness (lib.py, jython_shim.py, test_*.py, README, .env.example)
  db/
    CreateInventory.sql          # authoritative live DDL baseline (42 tables/182 procs/25 triggers)
    migrations/                  # the 18 spike-*.sql as ordered NN_*.sql (tables->procs/triggers->fixtures)
    namedqueries/                # *-namedqueries.sql (e.g. master-crud, kpi)
    fixtures/                    # gitignored provisioning target (.bak, .xls, CSVs) + README
    README.md                    # provisioning path (out-of-band .bak/CSV restore + migration order)
  themes/                        # tai-light/tai-dark + tai/*.css + fonts/ (SHARED-suite; gateway-deploy)
  design-history/                # landing-mockups/ + OrderSpike snapshot + provenance READMEs (quarantined)
  docs/                          # Ignition-side: copied source-truth pack + cutover playbook (see 4.A)
```

**Resolving the intermixing — EXACTLY how app code + views leave `docs/analysis`:**

Two couplings make a naive path-by-path `git mv`-into-new-repo unsafe (FLAG-F1, I-3,
harness §F1): e2e tests import app code and read spike SQL via hardcoded
`os.path.join(..., "docs","analysis",...)` constants. **Recommended: reorg-first in
the CURRENT repo, then fresh-copy.**

1. **Reorg-in-place PR (on the legacy repo, suite stays green):** `git mv` the
   **reference-derived move-set** (§2, not just the `project-library/` +
   `perspective-views/` subtrees) into a new top-level `ignition/` + `db/` tree shaped
   like the gateway. Critically, the move-set INCLUDES the runtime-loaded `.sql` and
   tooling the directory-name list missed — `reporting/sql/*.sql`, the three
   `spike-*-feed.sql`, `spike-report-procs-d6.sql`, `master-crud-namedqueries.sql`,
   `spike-inv-sites-*.sql`, and `order/spike/parity_diff.py` — because the harness/app
   `open()`s them at runtime (see §2 grep authority). In the SAME commit, fix BOTH
   halves of every harness path constant:
   - the path **TAIL** (`docs/analysis/...` → `ignition/...`/`db/...`), AND
   - the relative **HEAD** — 27 e2e files climb `os.path.join(__file__, "..","..",
     "docs",...)`; the `../..` is hardcoded to the `scripts/e2e/` nesting depth, so if
     the moved tests land at a *different* depth the climb breaks even when the tail
     string is corrected. Re-anchor each constant (e.g. to a single `REPO`/`_ROOT`
     base) so depth and tail are both right.

   Then re-run the full e2e suite to green. After this PR, `docs/analysis/**` is
   **spec-only for the moved areas** — see the precise STAYS/MOVES split in §2 (it is
   NOT "all of docs/analysis is pure spec": several `.md` modules + the `AD_FRSPULL`/
   `spike-asn-unique-guard` *citations* stay, and a few subtrees keep illustrative
   `.sql` that nothing loads).
2. **Fresh-copy to `inventory`:** copy `ignition/` + `generators/` + `e2e/` + `db/` +
   `themes/` + the source-truth doc pack into the new repo. Clean directory cut, no
   history rewrite.

Rationale for reorg-first over path-by-path: a path-by-path extraction leaves the
harness pointing at paths that exist in *neither* repo and forces simultaneous fixups
across two repos. Reorg-first proves the suite green in one place before the cut, and
makes the legacy `docs/analysis` genuinely spec-only. **If the reorg PR is judged too
heavy**, the fallback is a single atomic copy-into-`inventory` commit that rewrites
all `docs/analysis/...` path constants to the new `ignition/`+`db/` layout and re-runs
the suite — but it MUST be one atomic, suite-verified commit, or tests silently
"import nothing" (green-on-zero).

---

## 2. STAYS / MOVES master list

### Reference-derived MOVES — the grep authority (B1)

**The MOVES set below is derived by grepping what the code actually opens/imports at
runtime, NOT by directory name.** Any file the moving harness/app/generators read at
runtime MUST move with it. Grep commands that establish the authority (re-run as the
reorg gate):

```
# every runtime open()/load of a .sql or app code from e2e + app code:
grep -rnE 'open\(|\.sql|__file__|os\.path\.(join|normpath|dirname)' scripts/e2e docs/analysis --include='*.py'
# gateway-path + docs/ write targets in the generators:
grep -nE 'os\.environ\.get|/usr/local/ignition|docs/(design|analysis)' scripts/gen_*.py
```

**Runtime deps the directory-name draft MISSED (now in MOVES):**

| File(s) | Read at runtime by | Why it's a live dep, not spec |
|---|---|---|
| `docs/analysis/reporting/sql/{daily_shipping_assy,daily_shipping_header_assy,forecast_detail,invoice_summary,invoice_summary_faithful,lot_location_tires,lot_location_wheels}.sql` (7) | `report_render/driver.py:31-40` (`_SQL_DIR = ../../sql`, `open(path)`); `test_m3_reports.py:42,92` (`SQL_DIR`, `load_sql`) | The NQ-substitute SQL the report renderer + its parity test LOAD on every run. (`reporting/sql/adversary-findings-m3.md` in the same dir is spec — stays.) |
| `docs/analysis/edi/810/spike-edi810-feed.sql` | `test_edi810_e2e.py:77,186` (`FEED_SQL_FILE`, drift-guard `open`) | The CREATE/RECREATE feed the 810 driver inlines; test asserts byte-identity against this file. |
| `docs/analysis/edi/856/spike-edi856-feed.sql` | `test_edi856_e2e.py:92,176` (`FEED_SQL_FILE`) | Same drift-guard for the 856. |
| `docs/analysis/order/spike-order-file-feed.sql` | `test_order_file_e2e.py:59,127` (`FEED_SQL_FILE`) | Order-file feed drift-guard. |
| `docs/analysis/reporting/spike-report-procs-d6.sql` | `test_report_procs_d6.py:27,60` (`MIGRATION_SQL`, `open().read()`) | The D6 migration the test applies into the live container. |
| `docs/analysis/master-data/master-crud-namedqueries.sql` | the 6 master-CRUD tests cite/apply it; the shared NQ layer | The CRUD NQ source the master views bind through. |
| `docs/analysis/master-data/spike-inv-sites-paths.sql` | `test_sites_master.py:60` (`PATHS_DDL`, applied) | Sites path-columns DDL the test applies + asserts idempotent. |
| `docs/analysis/master-data/spike-inv-sites-table.sql` | `test_sites_crud.py:46` (provisioning) | Sites table the CRUD tests provision against. |
| `docs/analysis/order/spike/parity_diff.py` | `scripts/gen_parity_tsv.sh:32` (`python3 docs/analysis/order/spike/parity_diff.py`) | Order parity TOOLING invoked by the gen script — moves with `generators/`/`db/`. |
| **(input)** live gateway `…/Admin/Users/Users/resource.json` (valid re-signed `attributes`) | gateway cold-start re-sign | NOT in the repo; see S1/R3 — a REQUIRED move-set input from a live export. |

**NOT a runtime dep (citation only — stays as spec):**
`docs/analysis/production-readiness/AD_FRSPULL-shared.sql` is referenced in
`asn/code.py:7,319` and `test_asn_fanout.py:10` as a **NOTE / provenance comment only**
— neither `open()`s it (the proc runs live via `EXEC AD_FRSPULL`). It is spec; it STAYS.
`spike-asn-unique-guard.sql` is likewise cited (apply-by-hand note), not loaded —
treat as a `db/` migration by convention, not as a code-coupled file.

### STAYS — current `InventorySystem` (legacy + analysis) repo

| Subtree / path | Why it stays |
|---|---|
| All 64 `*.pas` + 57 `*.dfm` + `InventorySystem.dpr` + `.cfg/.dof/.res/.ico` | Delphi source — the reason the repo exists |
| `CAMEX Reports/`, `Database/`, `EDI/`, `EDIIn/`, `MAS Files/`, `Reports/`, `Suppliers/`, `Templates/`, `WWW Files/` | Legacy Delphi resource/data dirs |
| `forecast.ini`, `InventorySystem*.INI` | Legacy config (already gitignored; untracked local) |
| `docs/*.xls/.doc/.pdf/.txt/.ppt`, `docs/triggers.sql` | Delphi-era EDI source documents |
| **All `docs/analysis/**/*.md` spec files** | Reverse-engineering bridge — cite Delphi `file:line` |
| `docs/analysis/production-readiness/AD_FRSPULL-shared.sql`, `docs/analysis/edi/spike-asn-unique-guard.sql` | **Citations only** (NOTE in `asn/code.py`/tests — never `open()`ed); spec, not a code dep (B1) |
| `docs/analysis/reporting/sql/adversary-findings-m3.md`, `priceless-lines-diagnostic.sql` | Spec/diagnostic in the reporting tree — NOT loaded by the renderer; stays |
| Cutover/decision specs at analysis root (`cutover-*.md`, `decisions.md`, `ignition-*.md`) | Spec/playbook (copies snapshot into inventory; canonical stays) |
| **Spec-only subtrees** (no app code/views AND no runtime-loaded `.sql`): `admin/`, `assembly/`, `cross-cutting/`, `forecasting/`, `production-calendar/`, `receiving/`, `shipping/` | All `.md` (+ illustrative `.sql` nothing loads) — clean STAYS. NB: this is the *precise* spec-only set — `reporting/`, `edi/`, `order/`, `master-data/`, `production-readiness/` are MIXED (hold the runtime deps above), so "all of docs/analysis is pure spec after the move" is FALSE until the §2 grep gate confirms zero surviving reads |
| `DB Schema/Create Inventory.superseded-2026-06-01.sql` | Citation anchor for pre-2026-06-16 specs only |
| `README.md`, `CLAUDE.md`, `.gitignore` | Repo-meta — **revised in place** (see §5), not moved |

### MOVES — new `inventory` repo

| Subtree / path | Becomes |
|---|---|
| `docs/analysis/edi/{810,856,inbound,}/project-library/**` (edi810, edi856, edi_inbound, forecast, asn, hotcall + `message-handler.py`) | `ignition-project/ignition/script-python/` |
| `docs/analysis/inventory-stock/project-library/**` (receiving, reject, shipping, stockLedger, stocktaking) | script-python/ |
| `docs/analysis/order/project-library/**` (order, renban, order_file, forecast_distribution) | script-python/ |
| `docs/analysis/production-readiness/project-library/**` (auth, auto_purge) | script-python/ (auth flagged SHARED) |
| `docs/analysis/reporting/project-library/report_render/**` (`code.py`,`report_defs.py`,`driver.py`) | script-python/ (drop the stray `code$py.class`) |
| **`docs/analysis/reporting/sql/*.sql` (7 runtime-loaded NQ-substitute files)** | `db/namedqueries/reporting/` (or `script-python/report_render/sql/`) — `driver.py` resolves `../../sql`, so KEEP THE RELATIVE LAYOUT or repoint `_SQL_DIR` (B1) |
| **`docs/analysis/edi/810/spike-edi810-feed.sql`, `edi/856/spike-edi856-feed.sql`, `order/spike-order-file-feed.sql`** | `db/` (drift-guard feed SQL the e2e `open()`s — fix `FEED_SQL_FILE` constants) (B1) |
| **`docs/analysis/reporting/spike-report-procs-d6.sql`** | `db/migrations/` (test applies it; fix `MIGRATION_SQL`) (B1) |
| **`docs/analysis/master-data/master-crud-namedqueries.sql`** | `db/namedqueries/` (the shared CRUD NQ layer the master views bind through) (B1) |
| **`docs/analysis/master-data/{spike-inv-sites-paths,spike-inv-sites-table}.sql`** | `db/migrations/` (applied by `test_sites_master.py`/`test_sites_crud.py`) (B1) |
| **`docs/analysis/order/spike/parity_diff.py`** + `scripts/gen_parity_tsv.sh` | `generators/` (gen script invokes the `.py` by literal path) (B1) |
| `docs/design/perspective-views/**` (Home, Shell/{BackBar,ComingSoon}, Master/MasterHub, Order/*, Reports, HotCall, page-config) | `ignition-project/.../views/` (CANONICAL — the live tree) |
| `docs/analysis/master-data/perspective-views/Master/**` (8 master detail views) | views/Master/ (only in-repo copy of the 8 forms — must merge) |
| `docs/analysis/production-readiness/perspective-views/Admin/Users/**` | views/Admin/Users/ (only in-repo copy — repo `resource.json` has NO `attributes`; see Users-resource input below + R3) |
| **(input, NOT in repo) live gateway `Admin/Users/Users/resource.json` with valid `attributes`** | views/Admin/Users/Users/resource.json — **REQUIRED move-set input** exported from a live gateway (current sig `ea2cb00e…`); the repo copy alone bricks the gateway (S1/R3) |
| `docs/analysis/edi/project-library/hotcall/perspective-views/HotCall/**` | **DEDUP** — byte-identical to design copy, keep one |
| `docs/analysis/order/spike/perspective-views/Order/OrderSpike/**` + README | `design-history/` — SUPERSEDED, not routed live |
| `docs/design/themes/**` (tai-light/dark, tai/*.css, fonts/, README) | `themes/` (SHARED-suite; gateway-deploy artifact) |
| `docs/design/landing-mockups/**` (~1.6 MB PNG/html) | `design-history/` (provenance, quarantined) |
| `scripts/gen_*.py` (12), `scripts/build_fonts_css.py` | `generators/` (retargeted) |
| `scripts/e2e/**` (`.py` harness + README + `.env.example`) | `e2e/` |
| The `spike-*.sql` in **BOTH roots** — `scripts/spike-*.sql` (2: `spike-manifestcost-relax-index.sql`, +obsolete `spike-vehicleorder-line-fixture.sql` DROP) and `docs/analysis/**/spike-*.sql` (17) — plus `*-namedqueries.sql` (`master-crud-namedqueries.sql`, `Home/kpi-namedqueries.sql`) | `db/migrations/` + `db/namedqueries/`. The sweep MUST cover both roots (N2). |
| `DB Schema/CreateInventory.sql` | `db/CreateInventory.sql` — **SHARED** (move canonical + leave pointer stub in legacy) |

### Dual perspective-views consolidation (the views exist in 3 places)

`docs/design/perspective-views`, `docs/analysis/**/perspective-views`, and the live
gateway. Verified byte-identical where they overlap (manifest perspective-ui diff -q).
Consolidation rules:
- **De-dup by identity** — keep ONE gateway-faithful copy per view under
  `ignition-project/.../views/`. HotCall (3-way identical) collapses to one.
- **Authoritative page-config = the 16-route design/gateway one.** DROP the stale
  10-route `docs/analysis/master-data/perspective-views/page-config/config.json`
  (FLAG-V1) — it routes `/order→OrderSpike` and has no docks.
- **Pull the 8 master views + Admin/Users out of `docs/analysis`** — no design twin;
  they are the only in-repo copies.
- **Demote OrderSpike + landing-mockups to `design-history/`** (FLAG-V2/D1).
- **Exclude gateway-only cruft** if seeding from the gateway tree: `Order/OrderPrototype`
  and `Test/` exist only on the gateway, never versioned — do NOT import (FLAG-V4).

### Drop / do-not-move
- `docs/analysis/reporting/project-library/report_render/code$py.class` (build artifact, FLAG-F4).
- `scripts/spike-vehicleorder-line-fixture.sql` — flagged obsolete by
  `test_partsstock_crud.py:81` (harness §F6). Drop, don't carry.
- `__pycache__/`, `*.pyc`, `*$py.class`, `scripts/e2e/artifacts/` — regenerated, never tracked.

---

## 3. Git-history strategy — FRESH START (confirmed)

**Confirm the saved decision: fresh repo, no history rewrite.** Rationale:
- The spike history (134 commits) is a *spike* — exploratory, with the app code living
  in the "wrong" place (`docs/analysis/`). Preserving it via `git filter-repo`/subtree
  would carry forward thousands of paths that no longer exist in the new layout, plus
  the reorg in §1 step 1 already moves everything — a filtered history would show a
  giant rename commit anyway.
- The bridge value (why a proc behaves a certain way, the `file:line` provenance) lives
  in the **legacy repo's `docs/analysis` and the in-code provenance comments**, both of
  which are preserved — not in the spike's commit graph.

**Mechanics (fresh):**
1. Do the §1 reorg PR in the legacy repo (so the move-set is a clean directory).
2. `mkdir inventory && cd inventory && git init` (default branch `main`).
3. Copy the move-set directories in (no `.git`). Author the fresh `CLAUDE.md`/`README`/
   `.gitignore`/CI.
4. Single seed commit: "Initial import: InventorySystem Ignition app (spike snapshot
   <legacy-sha>)". Record the legacy SHA in the commit body for traceability.
5. Push to a new GitHub repo `inventory`.

**If preserve is later argued** (not recommended): `git filter-repo --path docs/analysis
--path scripts --path 'DB Schema'` then path-rename — but only AFTER the §1 reorg, and
accept a noisy history. Not worth it for a spike.

---

## 4. Cross-boundary handling

### A. What `inventory` needs FROM `docs/analysis` (the bridge)
The analysis specs STAY (they cite Delphi `file:line`). But the e2e suite asserts
against parity oracles that should not dangle across a repo boundary.
- **COPY (snapshot into `inventory/docs/`):** the cutover playbook (`cutover-runbook.md`,
  `cutover-punch-list.md` — holds P19/P21, `cutover-readiness-checkpoint.md`),
  `decisions.md`, and the per-module parity oracles the suite asserts against:
  `*-sourcetruth.md`, `*-wire-format.md`, `adversary-findings-*.md`. These are
  *frozen build-time references*; note in both CLAUDE.md files that legacy
  `docs/analysis` is canonical for spec.
- **CROSS-LINK (don't copy):** the bulk form-by-form reverse-engineering `.md` that
  cite `.pas:line` — they belong with the Delphi source. `inventory/CLAUDE.md` points
  back to the legacy repo for "why the proc behaves this way."

### B. Provenance comments (`.pas` / `.md` citations in app code) — KEEP AS-IS
15 app modules carry `# ...Order.pas:628-893`-style comments (X-1). These are the
derive-from-source bridge (R14/R20). After the split they become cross-repo doc links.
**Keep as-is** — cheap, historically accurate, and the cited Delphi source stays put.
Do not rewrite to point at specs.

### C. Generator gateway-path constants (B2 — NOT a one-line flip)
**Grep-verified split (not "all 10 env-overridable"):**
- **Env-overridable — 2:** `gen_landing_view.py`, `gen_renban_views.py` use
  `PROJ_DIR = os.environ.get("PROJ_DIR", "/usr/local/ignition/data/projects/InventorySystem")`.
  For these, flipping the default (or setting `PROJ_DIR` in repo config) is the one-liner.
- **Hardcoded gateway path, NO env fallback — 8:** `gen_backbar_docks.py` (`GW_CONFIG`),
  `gen_master_detail_layout.py` (`GW_BASE`), `gen_master_form_refinements.py` (`GW_BASE`),
  `gen_master_theme_tokens.py` (`GW_BASE`), `gen_master_write_gates.py` (`GW_BASE`/`GW_OUT`),
  `gen_partsstock_resilience.py` (`GW_VIEW`), `gen_sites_view.py` (`GW_OUT`),
  `gen_user_admin_view.py` (`GW_OUT`). Each pins
  `/usr/local/ignition/data/projects/InventorySystem/...` as a string constant with no
  `os.environ`. (The adversary listed 7; the grep finds **8** — `gen_partsstock_resilience.py`
  was also missed.)

So the centralize work is **per-file env-parametrization surgery on those 8** (introduce
the same `PROJ_DIR`/`GW_BASE = os.environ.get(...)` shape, mirroring the 2 that already
have it) — done in the **same sweep** as the centralized DB constant (§D), NOT a default
flip. `gen_hotcall_view.py` is separate: it writes to `docs/analysis/...` (a repo copy),
not the gateway — repoint to the consolidated `ignition-project/` views tree.

Retire `gen_landing_view.py`'s **dual-write** to `docs/design` (`L303`,
`write_view(os.path.join(REPO_ROOT,"docs","design","perspective-views"), view)`) and the
analogous repo-copy writes in `gen_renban_views`/`gen_master_write_gates`/
`gen_master_theme_tokens`/`gen_sites_view` — generators write only to the one tracked
`ignition-project/` tree. (P19/P21 territory; FLAG-G1.)

### D. `Inventory_Spike` DB-connection name (P19/P21) — CENTRALIZE
**Grep-verified counts (N1):** the literal `Inventory_Spike` appears in **60 files**
total; **32 are live `.py`** — **13 app-code project-library modules** (edi810, edi856,
edi_inbound, forecast, asn, hotcall `code.py` + `message-handler.py`, stockLedger,
forecast_distribution, order_file, order, renban, auto_purge) + **15 e2e `.py`** (incl.
the shim's logical→physical map `jython_shim.py:27,35` + ~14 test files) + **4 generators**
(`gen_landing_view.py`, `gen_renban_views.py`, `gen_partsstock_resilience.py`,
`gen_sites_view.py`). (The draft's "17+" was a large undercount.) **Replace all 32 `.py`
copies with a single shared constant** (`ignition/script-python/_shared/db.py` exposing
`DATABASE`), and parametrize the shim/tests/generators via env. Decide the **prod
connection name** at cut time and set it once. The shim already centralizes the map —
fix there + sweep the literals. Do it as an **idempotent injector + `--check`** (the
`gen_master_write_gates.py` model) so the sweep lands evenly across all 32, not via N
hand-edits.

### E. Gitignored `.bak`/CSV/`.xls` — PROVISIONED, never committed
None of this enters git in either repo. `inventory/db/README.md` documents the path:
1. Obtain `Inventory.bak` (+ `VehicleOrder.bak` if order path exercised) out-of-band
   (shared drive / secrets store), restore into docker `mssql-spike` container DB `Inventory`.
2. Apply `db/CreateInventory.sql` baseline, then the 18 `db/migrations/NN_*.sql` in
   dependency order (tables→procs/triggers→fixtures).
3. Client CSVs (`AllLog/ReportLog/DailyWorkLog`) and `OrderSimulation*.xls` (4) →
   provision into the dev DB only, out-of-band, into gitignored `db/fixtures/`.
4. `cp e2e/.env.example e2e/.env`, fill dev gateway creds.
Carry these gitignore rules forward: `*.bak`, `OrderSimulation*.xls`,
`AllLog.csv`/`ReportLog.csv`/`DailyWorkLog*.csv`, `e2e/.env`, `e2e/artifacts/`,
`__pycache__/`, `*.pyc`, `*$py.class`.

### F. Inbound file-drop path (I-4)
`edi_inbound/code.py` references the legacy `EDIIn`/`CAMEX Reports` file-drop dir
(which STAYS in the legacy repo). Repoint to a configured inbound path in `inventory`
(config, not logic).

---

## 5. Per-repo CLAUDE.md / README / .gitignore / CI

### Legacy `InventorySystem` — revise in place
- **CLAUDE.md:** keep the Delphi orientation. ADD a top banner: *"This is the LEGACY +
  ANALYSIS repo. The Ignition rebuild lives in the `inventory` repo. `docs/analysis/**`
  is the reverse-engineering bridge (cites Delphi `file:line`); after the split it no
  longer contains app code or Perspective views."* Relax the modernization-progress
  framing that implies the Ignition build happens here.
- **README.md:** unchanged Delphi overview + one cross-link to `inventory`.
- **.gitignore:** unchanged.

### New `inventory` — author fresh (do NOT copy the Delphi CLAUDE.md)
- **CLAUDE.md:** Ignition 8.1.52→8.3 project-as-code; Designer↔Git round-trip;
  `script-python/` = Jython app modules; `views/` = UI; `e2e/` = Playwright/Jython
  parity harness; `db/CreateInventory.sql` = authoritative DDL. Carry forward
  guardrails that still apply: `Inventory_Spike` connection name (P19/P21, now
  centralized), gateway-path constants, secrets (`.bak`/INI) gitignored, the
  derive-from-source/R14-R20 discipline. Point spec/oracle questions at legacy
  `docs/analysis` + the copied source-truth pack.
- **README.md:** Ignition app overview, dev-env bring-up (gateway + Colima/mssql),
  how to run the e2e suite.
- **.gitignore:** start from legacy, DROP the Delphi-build section
  (`*.dcu/.dcp/.bpl/.exe/__history`), KEEP the Python/Ignition/secret rules (§4.E).
- **CI (`.github/workflows/`):** the new capability the spike never had —
  (1) spin docker `mssql-spike`, restore `.bak` from secrets, apply
  `CreateInventory.sql` + migrations; (2) bring up gateway 8.1.52; (3) run the e2e
  suite; (4) deploy/promotion gate 8.1.52→8.3 (guard the 8.1→8.3 deltas per memory
  `project-ignition-version-constraint`). Designer↔Git: version the gateway project
  on-disk as `ignition-project/` (git working copy or symlink
  `…/data/projects/<name>` → repo), so a Designer save = a Git diff. Exclude churny
  fields (`lastModificationSignature`, Supplier `thumbnail.png`) via `.gitattributes`.

---

## 6. Open questions for David (decide at cut time)

1. **Gateway project name / connection name.** The spike uses project
   `InventorySystem` + connection `Inventory_Spike`. Pick the prod names now
   (P19/P21) so the centralized constant + `PROJ_DIR` default are set once.
   Recommendation: project `inventory`, connection `Inventory` (or `inventory_prod`).
2. **One inheriting Ignition project vs. a shared parent.** mes-suite (ALC/GALC +
   Manifest + Admin) is future. Do we (a) keep one gateway project now and factor a
   shared parent later, or (b) stand up a shared parent project now? Recommendation:
   **(a)** — themes/auth/nav-shell are gateway-scoped and already inherited; defer
   the parent until mes-suite is real.
3. **Shared platform layer location.** Themes (`tai-*`), auth (`auth/code.py` +
   write-gates), nav shell (back-bar docks), harness conventions (`jython_shim.py`)
   are SHARED-suite. Start them in `inventory` (only repo today) under clearly-named
   areas (`themes/`, `script-python/auth/`) and factor out to a `suite-shared`
   location when mes-suite lands? Recommendation: **yes, start-in-inventory, factor later.**
4. **Manifest overlap.** `INV_MANIFEST_COST_MST` (inventory) vs. the MES Manifest
   printer (mes-suite) — who owns the manifest-cost table/lookup at the DB level when
   both exist? Decide before mes-suite is built (not a blocker now).
5. **Create `mes-suite` now (placeholder) or later?** Recommendation: **later** — this
   split is purely inventory-extraction; an empty placeholder adds no value yet.
6. **CreateInventory.sql — one canonical + pointer, or accept a duplicate?**
   Recommendation: canonical to `inventory/db/`, pointer stub in legacy `DB Schema/`.

---

## 7. Risks + recommended step-by-step sequence

### Risks
- **R1 (blocker, highest):** harness imports app code + reads spike/runtime SQL by
  hardcoded `docs/analysis/...` relative paths (F1, I-3, harness §F1), AND the MOVES set
  was directory-name-derived so it missed live runtime deps (B1 — `reporting/sql/*.sql`,
  the feed/migration SQL, parity tooling). A move that doesn't carry those files AND
  rewrite the constants (BOTH tail and `../..` head, S2) in the SAME atomic commit makes
  tests "import nothing"/"file-not-found" — silently green-on-zero. Mitigation: the §2
  grep-derived move-set; reorg-first (§1); verify non-vacuity by reverting a known
  oracle; and the reorg gate's zero-surviving-cross-boundary-read grep.
- **R2 (drift):** the 3-location view fragmentation is identical only by manual
  discipline (FLAG-V3). Any un-exported Designer edit silently diverges. Mitigation:
  adopt direct project-as-code versioning (§5 CI) and retire the dual-write.
- **R3 (auth NPE landmine — brick risk, S1):** the in-repo
  `docs/analysis/.../Admin/Users/Users/` copy omits `attributes` (verified: `view.json`
  has none; `resource.json` is the deliberately attribute-less repo copy) AND there is
  **NO `docs/design` twin** — so a fresh repo's ONLY seed is the attribute-less copy,
  and on a COLD gateway start `ProjectFileTree.toResource().setAttributes()` NPEs on the
  null map and **faults the whole gateway** (observed 2026-06-22; `gen_user_admin_view.py:40-46`
  documents it). **Recovery source is PINNED:** the valid re-signed `attributes` map
  comes ONLY from a **live gateway export** (current `lastModificationSignature` =
  `ea2cb00e…` in the gateway's `resource.json`) — the generator writes an all-zeros
  placeholder the gateway re-signs on first warm load, so neither the generator nor the
  repo holds a usable map. **Make the live-gateway `resource.json` export a REQUIRED
  move-set input** (§2) and a sequence step, not an afterthought.
- **R4 (secrets):** never let `.bak`/CSV/`.xls`/`.env`/INI into the new repo. Seed
  `.gitignore` from day one (§4.E) BEFORE the first copy.
- **R5 (gateway-co-location + path-depth coupling, S2):** **8** (not 5) e2e `.py` read
  `/usr/local/ignition/...` directly (`lib.py` via the `GW_LOG` env default;
  `test_master_crud_logic`, `test_master_write_gates`, `test_master_form_refinements`,
  `test_m4_auth`, `test_sites_master` via hardcoded `GW_BASE`/`gw`; `test_m3_reports`,
  `test_forecast_distribution_e2e` via the `IGN = /usr/local/ignition` bundled-JRE root)
  — they assume the gateway is on the CI host. Parametrize via `GW_PROJECT`/path env.
  **Separately**, the relative-HEAD coupling: **27** e2e files resolve app code via
  `os.path.join(__file__, "..","..","docs",...)` — the `../..` is hardcoded to the
  `scripts/e2e/` depth, so a move to a different nesting depth breaks the climb
  *independent of the tail string*. The reorg commit must fix HEAD + TAIL together (§1).

### Recommended sequence (with verification gates)
1. **Verify open cutover prerequisites are landed** (do not split mid-flight):
   **PR #22 (order-seam-runner) is MERGED** (commit f86e03e, 2026-06-24) — confirm no
   regressions; confirm the P19/P21 connection-name decision (Q1). *(S3: #22 was OPEN at
   draft time and named as the first gate; it is now merged, so this is a verify step,
   not a blocker. FALLBACK if a future prereq stalls: do NOT start the reorg until it
   merges — a half-landed seam runner makes the §1 suite-green gate unreliable.)*
   *Gate: e2e suite green on legacy repo HEAD; `gh pr list --state open` shows no
   split-prereq PR.*
2. **Reorg-in-place PR** (§1 step 1): `git mv` the **§2 reference-derived move-set**
   (project-library/ + perspective-views/ + the runtime-loaded `reporting/sql/*.sql`,
   feed/migration `spike-*.sql` in BOTH roots, `master-crud-namedqueries.sql`,
   `spike-inv-sites-*.sql`, `order/spike/parity_diff.py`) into `ignition/`+`db/`;
   consolidate the 3-way views to one canonical copy each; in the SAME commit rewrite
   every harness/generator path constant — fixing BOTH the path TAIL and the `../..`
   relative HEAD (27 files, S2).
   *Gate: full e2e suite green; revert one oracle to prove non-vacuity; AND
   `grep -rnE 'docs/analysis|/usr/local/ignition' e2e/ <moved-app-code>` returns ZERO
   surviving cross-boundary path reads (B1). The moved `docs/analysis` areas now hold
   only `.md` + illustrative/citation `.sql` per the §2 spec-only set.*
3. **Centralize `Inventory_Spike`** into `_shared/db.py` + parametrize the shim, the
   **32 `.py`** literals (13 app + 15 e2e + 4 generators, §4.D), AND env-parametrize the
   **8 hardcoded generators** (§4.C) in the same idempotent-injector sweep.
   *Gate: suite green with the new constant; `--check` / grep shows no stray hardcodes
   in any of the 32 + no bare `/usr/local/...` default in the 8 generators.*
4. **Export the live-gateway `Admin/Users/Users/resource.json`** (valid re-signed
   `attributes`) and stage it as the seed for that view (S1/R3) — the repo copy alone
   bricks a cold gateway.
   *Gate: the exported `resource.json` contains a non-null `attributes` block with a real
   `lastModificationSignature` (not the all-zeros placeholder); a cold-start smoke shows
   no `FAULTED` in `wrapper.log`.*
5. **Create fresh `inventory` repo** (§3): init, copy move-set (incl. the Users export),
   author CLAUDE/README/`.gitignore`/CI, seed commit recording legacy SHA.
   *Gate: `.gitignore` blocks all secrets BEFORE first add (verify `git status` shows
   no `.bak`/CSV/`.env`).*
6. **Stand up CI in `inventory`** (§5): DB provision → migrations → e2e → deploy gate.
   *Gate: green CI run on a clean runner using only out-of-band-provisioned data.*
7. **Wire Designer↔Git** for the gateway project (version on-disk / symlink) and
   retire the generator dual-write.
   *Gate: a Designer save produces a clean Git diff; re-deploy 8.1.52→8.3 promotion works.*
8. **Revise the legacy repo's CLAUDE.md/README** banners + cross-links (§5).
9. **Tag the cut** in both repos (legacy: "pre-inventory-split"; inventory: seed) and
   record in memory `project-system-landscape`.

---

## 8. Readiness verdict

**READY to split now — with two must-do prerequisites and one careful execution gate.**

The project is in a strong position: build phase complete, cutover designed +
dress-rehearsed, themes/landing/views live-verified, e2e suite green, and the timing
trigger (David's user testing reaching min-nav) is MET. The app *code* is structurally
clean to lift — each `project-library/<module>/` is a self-contained scope-G resource,
and no spec `.md` lives inside a `project-library/` or `perspective-views/` dir; the
spec coupling is one-directional provenance comments, not import dependencies. **But
(B1) the move-set is NOT just the code dirs:** several runtime-loaded `.sql` and one
tooling `.py` live OUTSIDE `project-library/`/`perspective-views/` (notably
`reporting/sql/*.sql` loaded by `report_render/driver.py`, the feed/migration SQL, and
`order/spike/parity_diff.py`) — these must be carried by the §2 reference-derived set or
the moved suite goes red/green-on-zero. The lift is clean once the move-set is
grep-derived, not directory-derived.

**Prerequisites before the cut (do these first, in the legacy repo):**
- **P-A: Settle the connection/project-name decision (Q1)** — so the centralize step
  has a target. (Punch-list P19/P21.)
- **P-B: Do the reorg-first PR (§1 step 1) and prove the suite green** — this is the
  one structural change that de-risks everything else. It is also the answer to the
  central intermixing problem and to R1. The reorg MUST move the §2 grep-derived
  set (not just the code dirs) and fix HEAD+TAIL of all path constants (B1+S2).
- **P-C: Export the live-gateway `Admin/Users/Users/resource.json`** as the seed for
  that view — the repo copy alone bricks a cold gateway (S1/R3). Capture it while the
  spike gateway is live; it cannot be reconstructed from the repo.

**Not blockers, but settle at cut time:** the dual-view consolidation (mechanical,
covered in §2), the gateway-project-versioning model (§5 CI — adopt direct
project-as-code), and the open cutover punch-list residual (verify the
`IX_INV_MANIFEST_COST_MST` constraint-DROP on prod — that's a *prod-flip* concern,
independent of the repo split).

**Execution gate:** the split must be atomic w.r.t. the harness path rewrites
(R1) — reorg-first makes this a single, verifiable, in-place step rather than a
two-repo juggling act. Once the reorg PR is green, the fresh-copy to `inventory` is
low-risk.
