# Ignition + Perspective Feasibility Review — InventorySystem (vs Rails + Python)

*Recorded 2026-06-12. Produced by a 4-agent review workflow* (`ignition-vs-rails-feasibility`):
`ignition-architect` (design) -> `rails-python-architect` (incumbent baseline) ->
`adversarial-architect-reviewer` (refutation) -> synthesis. Run `wf_58326aea-590`.

**Context for this review:** the sibling Delphi system **GALC** is being moved to Inductive
Automation **Ignition**; this evaluates consolidating **InventorySystem** onto Ignition +
Perspective instead of the currently-planned **Ruby on Rails + Python** stack. The domain
decisions **D1-D8** ([`decisions.md`](decisions.md)) are platform-neutral and carry over either way.

**Key framing fact (from the domain expert):** app-side EDI is just **reading/writing X12 text
files in a shared directory**; a *separate, existing* SFTP integration with the VAN handles all
transport. The app does not move files to/from the trading partner.

> **Bottom line: LEAN-GO**, gated on a ~1-2 week vertical-slice spike on `PartsStockMaster`.
> The analysis investment (~100% of D1-D8 + spec sections 1-5/7/8/9) is preserved; only the
> "Target design (Rails primary)" sections would be redone.

---

## 1. Go / No-Go Synthesis

# GO / NO-GO: Switch InventorySystem from Rails+Python to Ignition + Perspective

## 1. VERDICT: **LEAN-GO (conditional)**
Single biggest reason: **platform consolidation with GALC is a real, recurring solo-dev win that Rails structurally cannot answer** — one skillset, one license pool, one infra to maintain for life. Rails' only durable counter-advantage (Postgres-portability of the model layer) is a Phase-2 concern you can defer, not a reason to run two runtimes forever.

## 2. CARRIES OVER vs REDONE (your sunk analysis is safe)
- **~100% carries over:** D1-D8 decisions, spec §1-§5/§7/§8/§9, the cross-cutting registers (29 P12 retry bugs, 24-trigger reconciliation), and the §9 parity tests (reusable verbatim as Ignition acceptance tests). These describe the *source system, procs, and triggers* — which don't change with target platform.
- **REDONE:** only the §6 "Target design (Rails primary)" sections — ActiveRecord/`acts_as_tenant`/ERB → Perspective views + Jython. ~1-2 days/module × ~10 modules ≈ **3 weeks of re-spec**, dwarfed by the build either way.
- **Quantitatively:** the redo is a modest fraction of each ~400-line spec. **The analysis investment is overwhelmingly preserved — switching now does not waste it.** (Switch cost is mostly in the not-yet-written form-by-form UI design, which would be new work under either platform.)

## 3. FIT SCORECARD
| Area | Rating | One line |
|---|---|---|
| Perspective CRUD/forms | **B−** | Maps cleanly (uniform grid+detail) but no scaffold generator — per-field drag/bind; the ~45 screens are the *one place Ignition costs more* than Rails (~+2-4 dev-months). |
| Data access (wrap-the-proc) | **A** | `createSProcCall` + Named Queries ≈ `tiny_tds`; leave qty triggers in DB during parallel run — invariant preserved, no fork. |
| Multi-site (D1) | **C+** | Platform-neutral decision, but enforcement is by-discipline (no `default_scope`); mitigable via server-derived `siteId` + centralized scoped-query wrapper. |
| EDI (transport external) | **B** | Jython handles X12 strings fine — BUT scope is ~2x what the proposal claimed (see risk #2); feasible, not the "free win" advertised. |
| Reporting (29 REPORT_*) | **A−** | Ignition Reporting module beats a Rails PDF stack; matrix reports (LotLocation) cost more than 0.5-1 day — calibrate one. |
| Solo-dev maintainability | **A** | One platform shared with GALC + bundled auth/logging/backups/HA. The decisive lifetime advantage. |

## 4. DECISIVE RISKS
1. **Dual-process EDI shared-dir contention during parallel run** — both legacy `EDIUpload`/`ASNInvoice` and the Ignition poller scan the same dirs: inbound archive-move race (lose a 997 ack) + outbound name-collision / non-atomic `Writeln` (ship a truncated 810 to Toyota's VAN). **FATAL if ignored, trivially MITIGABLE:** single-owner EDI I/O during cutover + atomic temp-then-rename.
2. **EDI scope is ~2x understated** — inbound is 7 doc types (830/862/997+AK1 ack loop/824/820 + `LogActLog` side-effects), not "read 830"; outbound must be re-derived from live `ASNInvoice.pas`, not the **dead** `Write810File.pas`. **Not fatal to platform choice; FATAL to the current estimate** until re-scoped. MITIGABLE by re-reading two units. Note: per-site `delSL[4]` filter lives *inside* the parser — a D1 hook §3 missed.
3. **Multi-site enforcement-by-discipline** — a forgotten `site_id` predicate = cross-site leak. MITIGABLE, but it's the #1 data-integrity risk; gate on a spike.
4. **Perspective UI throughput unproven for you** — the dominant cost line; if per-screen velocity is poor the whole effort balloons. MITIGABLE by calibration spike.
5. **Jython 2.7 ceiling / module licensing** — fine for X12 strings + proc calls; only bites if future logic needs Python 3 libs (not Phase-1). Get the real Perspective+Reporting quote. MITIGABLE / informational.

None are fatal-and-unmitigable. All four killers reduce to "prototype before committing."

## 5. RECOMMENDATION + NEXT STEP
**Proceed (Lean-Go), gated on a single time-boxed vertical-slice spike (~1-2 weeks) with three checks. Don't write the §6 redos until it passes.**

Concrete next step — **one thin vertical slice on PartsStockMaster** (the worst-case screen) that proves:
- **(a) UI velocity:** rebuild PartsStockMaster (~40 controls + 14-cell weekday matrix) in Perspective → calibrate your real per-screen cost, then extrapolate to ~45.
- **(b) Site isolation:** stand up the centralized server-derived `siteScopedQuery()` wrapper and prove no query can leak across `siteId`.
- **(c) EDI re-scope + single-owner I/O:** re-read `EDIUpload.pas` (enumerate all 7 `data='NNN'` branches + side-effects) and `ASNInvoice.pas:820-879`; prototype the inbound poller with atomic archive-move and confirm single-owner ownership during parallel run. Also calibrate one matrix report (`REPORT_NUMMILotLocationW`).

If (a) lands within tolerance and (b)/(c) prove clean → **full GO**. If Perspective form throughput is unacceptable for a solo dev → that's your one real reason to **Stay** with Rails.

---

## 2. Ignition + Perspective Target Architecture  *(ignition-architect)*

# Ignition + Perspective Target Architecture & Fit Assessment — InventorySystem

**Verdict up front: viable and strategically attractive (one platform with GALC), but it is NOT a free ride. The platform-neutral analysis (D1-D8, spec §1-§5/§7) carries over ~100%; the Rails-specific §6 sections (10 modules) must be redone. The UI rebuild — ~45 form-heavy CRUD screens — is where Perspective costs you, not data access.**

## 1. UI: ~45 CRUD forms → Perspective views

The legacy forms are uniformly the **grid + detail-panel master/detail** pattern (confirmed: Supplier §1, PartsStockMaster §1 "richest" with ~40 controls + two 7-cell matrices). This maps cleanly but not cheaply.

**Pattern → reusable Perspective scaffold (build once, reuse ~45×):**
- **List view:** Perspective `Table` bound to a Named Query, server-side search/sort/paging (replaces the in-memory `Filter`/`LIKE` legacy search noted in PartsStockMaster §5 P7, Supplier §5).
- **Detail view:** an **embedded view** parameterized by `recordId`, with input components bound to `view.params`. FK combos → `Dropdown` sourced from master Named Queries posting the surrogate id (D2).
- A **master/detail flex container** template that embeds both, wired to a shared `selectedId` session/view prop.

**Reuse mechanism:** Perspective **Embedded Views + view parameters** are the analog of Rails partials. You get genuine reuse of the list/detail chrome, but **each entity still needs its own field layout** — Perspective has no scaffold-generator. PartsStockMaster's ~40 controls + 14 weekday matrix cells is hand-placement work with no shortcut.

**Honest effort:** This is the single biggest cost line. Rails `scaffold` + ERB partials generate a working CRUD screen in minutes; Perspective is **drag-place-bind per field**. Realistic: 1-3 days per simple master (Supplier, Size, Logistics), 1-2 weeks for the heavy ones (PartsStockMaster, plus the transactional screens Order/Shipping/RecConfStat/ASN). For ~45 screens budget **roughly 2-4 developer-months of UI alone** for a solo dev — materially more than Rails for this form-count. The 29 REPORT_* screens are NOT in this count (see §5).

## 2. Data access: wrap-the-proc parallel run

This is where Ignition is genuinely strong and the strategy maps almost 1:1.

| Need | Ignition mechanism |
|---|---|
| Call existing CRUD procs unchanged (Supplier §3, PartsStock §3) | **`system.db.createSProcCall`** with registerInParam/registerOutParam — direct analog of the planned `tiny_tds` wrap. This is the backbone for Stage-1/Stage-2 parity. |
| Read lists for tables/reports | **Named Queries** (parameterized, cached, the canonical Perspective binding source). |
| Ad-hoc writes once logic is re-homed | `system.db.runPrepUpdate` / Named Queries (Phase 2+). |

**Stock-ledger / qty triggers — leave them in the DB during parallel run.** This is the correct call and it matches the existing plan. The 12 qty-adjusting triggers + 24 total triggers (trigger-reconciliation register: schema is authoritative, all key on int FKs now) **enforce the on-hand invariant**. During parallel run BOTH the live Delphi app and Ignition hit the same SQL Server; the procs+triggers mediate and keep `IN_QTY` correct for everyone. **Do not re-home the ledger into Ignition tags/scripts in Phase 1** — that would fork the invariant. Re-homing (the spec's "12 triggers → service transactions keyed on `IN_PART_ID`", PartsStock §6) is a **Postgres-phase** activity, and in Ignition it becomes **Jython transaction scripts** (or Named Queries inside a `system.db.beginTransaction`/`commitTransaction`), not ActiveRecord callbacks.

**Decisions land in the data layer regardless of platform — all carry over:**
- **D5** (stocktaking = signed delta): trigger behavior preserved as-is in parallel run; the Perspective entry field must be labeled +/− adjustment (D5 §UI note). Platform-neutral.
- **D7** (RecConfStat stamps `VC_ARRIVAL` → `'A'`-supplier arrival-add): the arrival-add re-homes to the **receiving-confirmation action**, which in Ignition is a Jython script on the RecConfStat Perspective view calling `UPDATE_RecConfStatInfo` (DataModule.pas:3346/:3269). Neutral.
- **D8** (3 bug fixes): Bug 1 (size dup-check wrong table), Bug 2 (NULL audit timestamp), Bug 3 (dead arrival-reversal branch → implement compensating −qty) are **DB/logic fixes independent of UI framework** — they apply identically under Ignition.

**One caveat specific to Ignition:** the legacy app uses **three `TADOConnection`s** (`Inv_`, `Act_`, `ALC_` — PartsStock §2 notes `VC_LINE_NAME` comes from `ALC_Connection`, a different DB). Ignition handles this fine (multiple named DB connections), but every cross-DB read (the `LINE` catalog) must be a separate connection — note and plan it.

## 3. Multi-site (D1): per-site isolation in Ignition

D1 (shared schema + `site_id` NOT NULL FK on every table, full isolation, auth binds user→site) is **platform-neutral and carries over entirely** — but the *enforcement mechanism* differs and is **weaker than Rails by default**:

- **Rails plan:** `acts_as_tenant`/`default_scope` — a model-layer guarantee that's hard to bypass.
- **Ignition:** there is no ORM and **no automatic query scoping**. The current site lives in a **Perspective session property** (set at login from the user's roles/IdP claim). Every Named Query and SProc call must **explicitly pass `:siteId`** and every WHERE must filter on it.

**The risk:** scoping is by developer discipline, query-by-query. A forgotten `site_id` predicate = cross-site data leak. Mitigations for a solo dev: (a) centralize all data access in a small set of Named Queries / a Jython `siteScopedQuery()` wrapper that injects `siteId` from the session, never inline SQL in views; (b) derive `siteId` server-side from the authenticated user, never from a client-settable param. During parallel run the legacy SQL Server is single-site (D1 §Phasing), so Phase-1 Ignition simply filters to its one site — the `site_id` columns themselves are a Postgres-phase change, same as Rails.

**User→site binding:** Ignition Security via IdP (SAML/OIDC) or internal users + roles; map a site claim/role to the session's `siteId` in a Perspective session-startup script. This is solid and arguably cleaner than rolling your own auth in Rails.

## 4. EDI as flat-file I/O (using the clarification)

**This is where Ignition's fit IMPROVES dramatically given the clarification.** Scope is only: write outbound X12 (810/856) to a shared dir, read inbound X12 (830) from it. Transport (SFTP/VAN) is **external and already exists**. The legacy units are small — `EDI810Object.pas` 451 lines, `EDI856Object.pas` 463, `Write810File.pas` 128 — i.e. ~1KB of **string concatenation of X12 segments**, not transport, not crypto.

**Ignition design:**
- **Gateway Timer Script** (or file-watch via a scheduled scan) polls the inbound shared dir; `system.file.fileExists`/`readFileAsString`, parse, write via procs, move/archive the file.
- **Outbound:** a Gateway script or button-triggered script builds the X12 string and `system.file.writeFile` to the outbound shared dir. The existing SFTP integration picks it up — Ignition never touches the VAN.

**Jython 2.7 vs external Python 3 helper:** X12 build/parse is **pure string/segment manipulation** (split on element/segment separators, join with `~`/`*`). Jython 2.7 handles this **fine** — it's well within Jython's string capabilities, and the legacy Delphi did it with hand-rolled string ops too. **No external Python 3 helper is needed** for the file I/O scope. (You'd only reach for an external helper if you wanted a real X12 library like `pyx12`/`bots`, which is overkill here and reintroduces the very polyglot complexity Ignition consolidation is meant to remove.) **Note D6:** the 810 billing read path must become **window-aware** (price = manifest-cost row whose `[start,end]` contains the ASN production date) — that's a Named Query fix in `SELECT_INVOICEItems`/`REPORT_EDI810*`, identical work under either platform.

**Vs the Rails+Python plan:** the Rails plan budgeted a separate Python service partly for EDI. **Given transport is external and scope is flat-file string I/O, that Python tier largely evaporates.** Ignition does EDI in-platform with Gateway scripts — this is a real simplification win for the solo-dev/one-skillset driver, and it narrows (does not erase) the overall effort gap from §1.

## 5. Reporting: 29 REPORT_* procs

The **Ignition Reporting module** binds a report's data source directly to a Named Query or stored-proc call, then renders PDF/Excel with scheduled or on-demand delivery. The 29 `REPORT_*` procs (e.g. `REPORT_LogicalInventory`, `REPORT_EDI810*`) become **29 Reporting data sources + layouts**. This is a **clean win over the Rails plan** — Rails would need a PDF/reporting gem and hand-built templates; Ignition Reporting is purpose-built and is part of the consolidation value. Caveat: it's a **separately-licensed module** (factor into the licensing comparison) and each report layout is still hand-built (budget ~0.5-1 day each = ~3-4 weeks for all 29). The D6 window-aware fix applies to the EDI810 report data source.

## 6. Auth / shell / admin

- **Auth:** Ignition IdP (SAML/OIDC) or internal user source + roles; site binding per §3. Equivalent-or-better than building Devise/auth in Rails.
- **Shell/nav:** a Perspective **Page Configuration** + shared header/nav embedded view replaces `MainMenu.pas` (138KB orchestrator). Straightforward.
- **Admin:** Gateway web UI gives you connection management, script status, logs, gateway backups out of the box — less custom admin than Rails. Gateway redundancy available for HA if the GALC consolidation wants it.

## 7. Carryover vs redo (quantitative)

| Artifact | Carries over to Ignition? |
|---|---|
| **D1-D8** (decisions.md) | **100%** — all platform-neutral domain/logic decisions. Zero rework. |
| **Spec §1 Legacy surface, §2 Data touched, §3 Procs, §4 Business rules, §5 UI notes, §8 open Qs, §9 parity tests** | **~100%** — these describe the *source system and the procs/triggers*, which don't change. The §9 parity tests are directly reusable as Ignition acceptance tests. |
| **Cross-cutting registers** (P12 retry bugs — 29 confirmed/8 critical; trigger reconciliation — schema authoritative, 24 triggers) | **100%** — DB/legacy facts. |
| **Spec §6 "Target design (Rails primary)"** | **REDONE** — every §6 is ActiveRecord models, `belongs_to`, `dependent: :nullify`, `acts_as_tenant`, RESTful controllers, ERB views, `tiny_tds`. None of this exists in Ignition. **~10 modules × §6 = the redo cost.** |
| **Spec §7 migration plan (Stage 1 read-only / Stage 2 wrapped procs / Stage 3 reimplement)** | **~80%** — the *staging strategy* is identical (wrap procs, parallel run, Postgres later); only the implementation verbs change (`createSProcCall` not `tiny_tds`; Jython transactions not AR callbacks). |

**Net:** the analysis investment is overwhelmingly preserved. Only the §6 design sections (a modest fraction of each ~400-line spec) and the not-yet-written form-by-form UI/service designs need re-authoring for the Perspective/Jython target. Re-spec'ing §6 is maybe **1-2 days per module = ~3 weeks**, dwarfed by the build itself.

## 8. Honest effort delta vs Rails + top solo-dev risks

**Effort delta (Ignition relative to Rails):**
- **UI: SLOWER (+).** No scaffold generator; ~45 form-heavy screens are per-field drag/bind. This is the dominant cost and the main reason Ignition is *more* work for this app. (Est. +2-4 dev-months over Rails for the form count.)
- **Data access: ~NEUTRAL.** `createSProcCall`/Named Queries ≈ `tiny_tds`/AR for the wrap-the-proc phase.
- **EDI: FASTER (−).** Transport is external; flat-file X12 in Jython kills the separate Python tier. Real saving.
- **Reporting: FASTER (−).** 29 procs → Reporting module beats a Rails PDF stack.
- **Auth/admin/infra: FASTER (−).** Gateway gives auth, logging, backups, HA for free.
- **Strategic: STRONG (−).** One platform/skillset/license shared with GALC — the actual driver, and a genuine long-term solo-maintainability win.

**Rough net:** UI cost pushes Ignition *above* Rails on raw build hours for *this* system in isolation; EDI + Reporting + infra + the GALC consolidation pull it back toward parity and ahead on lifetime maintenance. **Go/no-go pivots on weighting the one-platform consolidation against the heavier UI build.**

**Top risks / unknowns for a solo dev:**
1. **Per-site isolation is enforcement-by-discipline (§3).** No `default_scope` equivalent. Highest data-integrity risk. Mitigate with a centralized site-scoped query wrapper; never inline SQL in views; derive `siteId` server-side. **This is the #1 thing to prototype before committing.**
2. **Perspective form-building throughput is unproven for you.** Build PartsStockMaster (the worst case, ~40 controls + 14 matrix cells) as a spike to calibrate the per-screen estimate before sizing the whole project.
3. **No ORM = more boilerplate per write.** D2's "everything keys on the surrogate id," D3's RESTRICT-on-delete, and the qty-trigger re-homing (Postgres phase) become hand-written Jython transaction scripts rather than declarative associations. More code to get right and test (the §9 parity suites are your safety net).
4. **Cross-DB `LINE`/`ALC_Connection`** (PartsStock §2) — confirm Ignition multi-connection access to the Activity/ALC DB early.
5. **Jython 2.7 ceiling** — fine for X12 strings and proc calls; if any future logic wants Python 3 libraries you're back to an external helper. Not a Phase-1 concern.
6. **Module licensing math** — Perspective + Reporting (+ possibly redundancy) seats vs the GALC-shared licensing; get the actual quote into the go/no-go.

**Recommendation:** Conditional GO, gated on a two-spike de-risk: (a) the site-scoped query wrapper proving clean isolation, and (b) rebuilding PartsStockMaster in Perspective to calibrate UI velocity. If both land acceptably, the GALC consolidation + EDI/Reporting/infra wins justify absorbing the heavier UI build for a solo dev on one platform.

**Key files referenced:** `/Users/apple/Documents/GitHub/InventorySystem/docs/analysis/decisions.md`; `/Users/apple/Documents/GitHub/InventorySystem/docs/analysis/master-data/supplier.md`; `/Users/apple/Documents/GitHub/InventorySystem/docs/analysis/inventory-stock/parts-stock-master.md`; `/Users/apple/Documents/GitHub/InventorySystem/docs/analysis/cross-cutting/datamodule-retry-target-bugs.md`; `/Users/apple/Documents/GitHub/InventorySystem/docs/analysis/cross-cutting/trigger-source-reconciliation.md`; `/Users/apple/Documents/GitHub/InventorySystem/EDI810Object.pas`, `EDI856Object.pas`, `Write810File.pas`; `/Users/apple/Documents/GitHub/InventorySystem/DB Schema/Create Inventory.sql` (179 procs, 29 REPORT_*, 24 triggers).

---

## 3. Rails + Python Baseline -- the incumbent to beat  *(rails-python-architect)*

# Rails + Python Baseline — Shape & Strongest Claims (the bar Ignition/Perspective must clear)

**Shape:** Rails monolith (ActiveRecord on SQL Server via `activerecord-sqlserver-adapter`/`tiny_tds`) as the primary app + UI. Two thin Python edges: (1) X12 build/parse, (2) forecasting/FRS breakdown math, exposed as a small FastAPI service or scheduled jobs. Phase 1 = parallel run wrapping existing procs against the live SQL Server; Phase 2 = reimplement proc/trigger logic into AR + service objects and move to Postgres. Per-site `site_id` scoping modeled from day one, FKs/indexes deferred to the Postgres phase.

### (a) ~45 CRUD screens — *honestly, NO real Rails advantage; arguably Ignition's strongest counter-zone*
These are thin forms over named CRUD procs. Rails scaffolds them cheaply (controller + service-object that calls the proc + ERB/Hotwire), but Perspective is a drag-build form-over-data tool — this is exactly what it's optimized for. Rails gives me cleaner version control, real testing, and full HTML/PDF control for the 29 `REPORT_*` procs (replacing QuickReport). But for plain bound CRUD grids, I'd be hand-writing what Perspective gives by binding. **Net: a wash to slightly-favors-Ignition.** I should not claim CRUD volume as a Rails win.

### (b) Wrap-the-proc parallel run against SQL Server — *genuine, defensible Rails advantage*
ActiveRecord's `find_by_sql` / `connection.exec_query` over an existing SQL Server connection makes proc-wrapping trivial and idiomatic, and service objects give a clean home for the wrap→reimplement transition (P-patterns). Triggers (the stock-quantity invariants) re-home naturally into model callbacks / transactional service objects. Critically, **the Postgres exit (D-phase 2) is a first-class, well-trodden path for Rails**; Ignition is comfortable on any JDBC DB but does not give me the same model-layer to carry trigger/proc logic into a portable codebase. **This is the strongest Rails claim:** lowest-risk parallel run *plus* a credible reimplementation/portability story in one stack.

### (c) Multi-site D1 — *moderate Rails advantage*
Per-site isolation is a solved Rails idiom: a global default scope / `acts_as_tenant`-style pattern, session-bound `current_site`, scoped uniqueness validations (P1). This is enforced uniformly in *one* model layer that both the UI and the Python edges respect. Ignition can do site scoping via security zones/roles + parameterized bindings, but the discipline lives in per-screen bindings and named queries rather than a single enforced data layer — easier to leak. **Rails enforces D1 in one place; Ignition enforces it per-screen.** Real but not decisive.

### (d) EDI — *Rails/Python advantage SHRINKS materially under the new clarification*
With transport owned by the existing SFTP/VAN integration, app-side EDI collapses to: write outbound 810/856 X12 to a shared dir, read inbound 830 from it, plus X12 build/parse. Python (FastAPI job) is genuinely good at the parse/build (clear libraries, real unit tests, the math-heavy 830→forecast path), and Rails schedules/triggers it cleanly. **But "read/write text files in a shared directory" is trivial for any platform, including Ignition** (Jython file I/O + scripted parse, or its own gateway scripts). The former Rails moat — owning FTP transport with bounded retry (P8/P12) — is **gone**; transport is out of scope. So the EDI advantage narrows to "I'd rather write/test X12 parsing in Python than in Jython," which is a developer-ergonomics preference, not an architectural win. **Honest verdict: modest, taste-based edge, not a structural one.**

### Cross-cutting honest notes
- **Solo-maintainability cuts BOTH ways.** Rails+Python is *two* runtimes, a gem/dependency surface, and self-managed hosting for one person. Ignition is one platform with bundled infra/licensing and aligns with the GALC→Ignition consolidation (one skillset, shared infra) — a strategic pull Rails cannot answer.
- **Where Rails genuinely wins:** (b) parallel-run + Postgres portability, (c) single-layer D1 enforcement, and full control of the 29 reports.
- **Where Rails does NOT win:** (a) form-over-data CRUD volume, and now (d) EDI, since it's just shared-dir text I/O with transport already solved elsewhere.

**Bottom line for go/no-go:** the Rails baseline's defensible core is the *wrap-then-reimplement-to-Postgres* migration path and unified multi-site/trigger logic in one model layer. Ignition must be beaten primarily *there*. On CRUD screens, EDI file I/O, and platform consolidation, Rails carries little or no advantage — and consolidation actively favors Ignition.

Relevant source paths: `/Users/apple/Documents/GitHub/InventorySystem/docs/analysis/decisions.md`, `/Users/apple/Documents/GitHub/InventorySystem/docs/analysis/master-data/`, `/Users/apple/Documents/GitHub/InventorySystem/docs/analysis/inventory-stock/`, `/Users/apple/Documents/GitHub/InventorySystem/docs/analysis/cross-cutting/`.

---

## 4. Adversarial Critique -- stress-testing the Ignition fit  *(adversarial-architect-reviewer)*

I now have the ground truth. The active outbound writer is `ASNInvoice.pas` (not `Write810File.pas`, which is dead), and inbound `EDIUpload.pas` handles 7 document types with site-filtering and archive-move disposition. Enough verified.

---

ADVERSARIAL REVIEW — Ignition/Perspective fit for InventorySystem

Counts independently confirmed: 179 procs, 29 `REPORT_*`, 24 triggers (exact). D1-D8 platform-neutrality, the carryover table, the data-access mapping, multi-site enforcement-by-discipline risk, and the reporting-module mapping all survive scrutiny. The strategic core of the proposal is sound. The problem is the EDI section, which is the one part the proposal sold as a "dramatic improvement" — and it is built on an under-read legacy surface.

LEAD FINDING — EDI scope is materially understated. RISK (one claim WRONG).

The proposal says: "Scope is only: write outbound X12 (810/856) to a shared dir, read inbound X12 (830) from it." Verified against the live, dpr-listed `EDIUpload.pas`:

- Inbound is NOT just 830. `EDIUpload.pas` parses SEVEN document types by branch: `830` (:89), `862` (:105), `997` (:186, with an `AK1` acknowledgment loop :192), `824` (:253, with `NTE` handling :281), `820` (:306). 997/824/820 drive EIN accept/reject state updates and `LogActLog` writes — this is functional acknowledgment + application-advice + remittance ingest, not a one-document forecast read. Re-speccing this as "read 830" will silently drop 997 ack reconciliation and 824/820 handling. Settle it: enumerate every `data='NNN'` branch and its proc/log side-effects before sizing the Jython rewrite.

- `Write810File.pas` is DEAD CODE. The proposal cites its 128 lines as evidence of "~1KB of string concatenation." Verified: its entire `Execute` body is inside a `{ }` comment block (`Write810File.pas:33-118`); both the "true EDI" and "TAI EDI" branches are commented out. It IS listed in the dpr, but the body does nothing. The live outbound writer is `ASNInvoice.pas:820-879` (`AssignFile`/`Rewrite`/`Writeln`, with three distinct filename schemes including an `8HC...` variant at :825 and the `810...` at :872). Building the Ignition EDI estimate off `Write810File.pas` is reasoning from a corpse. This is exactly the failure-mode-checklist item 3 (spec built on dead code). Re-base the outbound analysis on `ASNInvoice.pas`.

- Inbound disposition has a real idempotency/atomicity surface the proposal hand-waved. `EDIUpload.pas:418-435`: after read, files are MOVE-archived to `\Archive\` keyed by `EDIFileNumber+EIN+'.EDI'` (with an EIN-empty fallback using `delSL[10]`, :428). And there is per-site filtering: ":not a trading partner file for this site('+delSL[4]+')" (:439). Consequences for the Ignition design:
  - The "poll the dir, parse, move/archive" loop must reproduce the EXACT archive-naming and the EIN-derived key, or re-ingest dedup breaks. Two files with the same EIN collide on the archive name — confirm that's intended before replicating.
  - File LOCKING/partial-write: the legacy reads with `FindFirst`+`Reset` and trusts the SFTP/VAN process not to be mid-write. An Ignition Gateway timer scanning the same dir hits the same race — but now during PARALLEL RUN, BOTH the legacy app's `EDIUpload` AND the Ignition poller scan the same `fiEDIIn` dir. First mover archives the file; the other sees it vanish or reads a half-moved file. That's a parallel-run data hazard the proposal's §2 (which only considered the SQL Server) did not cover. Fatal-if-ignored, trivially mitigable: only ONE of the two apps may own inbound polling at a time. Bounce to whoever owns the parallel-run cutover plan.
  - The per-site `delSL[4]` filter is the D1 multi-site hook living INSIDE the EDI parser, not just in SQL. The proposal's §3 multi-site analysis never mentions that site-scoping is already embedded in the inbound file router. Don't lose it.

Net on EDI: the proposal's conclusion ("Jython 2.7 handles X12 string ops fine; no Python tier needed") is still probably correct — but the evidence offered is wrong (dead file) and the scope is ~2x understated (7 inbound doc types + ack reconciliation + archive idempotency + dual-poller race). The "dramatic improvement / Python tier evaporates" claim is RISK, not SOUND, until re-scoped against `EDIUpload.pas` and `ASNInvoice.pas`.

SECOND — Parallel-run shared-directory hazard generalizes. RISK.

§2 correctly says procs+triggers mediate concurrent SQL writes. But the EDI shared dir is a SECOND shared mutable resource with no trigger to mediate it. Outbound: both apps could write `810NNNN.txt` to `fiEDIOut` with the same date-derived name (`ASNInvoice.pas:872` uses `copy(PickUpDate,5,4)` — a 4-char date stub, NOT an EIN) → name collision / last-writer-wins / the SFTP picks up a half-written file mid-`Writeln` loop (:879, no temp-then-rename atomic write). Legacy got away with it as a single instance; two writers is new. Mitigation direction: atomic write (temp file + rename) and single-owner outbound during parallel run. Bounce to the migration-plan author. Mitigable.

THIRD — Things the proposal got RIGHT (don't manufacture objections):

- SOUND: 179/29/24 counts, exact. The whole carryover table (§7) — D1-D8 and §1-§5/§7 describe the source system and procs/triggers, which don't change; the redo is genuinely confined to §6. This is honest accounting.
- SOUND: "Leave the qty triggers in the DB during parallel run; re-homing is a Postgres-phase activity." This directly avoids forking the on-hand invariant — the single biggest correctness trap, correctly identified.
- SOUND: `system.db.createSProcCall` ≈ `tiny_tds` wrap for the named CRUD procs. The thin-UI-over-procs shape maps to Named Queries + sproc calls cleanly.
- SOUND and well-flagged: multi-site enforcement-by-discipline (no `default_scope` analog) as the #1 data-integrity risk, with the right mitigation (server-derived `siteId`, centralized wrapper, never client-param). The PartsStockMaster spike to calibrate UI velocity is the correct de-risk. Note the §3 caveat that legacy is single-site is reinforced by the `delSL[4]` finding above — site logic is more pervasive than §3 assumed, which only strengthens the "prototype isolation first" recommendation.
- SOUND: three `TADOConnection`s (`Inv_/Act_/ALC_`) → multiple Ignition named connections; calling it out to confirm cross-DB `LINE`/`ALC` access early is the right instinct.

FOURTH — Reporting. RISK (cost, not feasibility).

29 `REPORT_*` → Ignition Reporting data sources is sound mechanically and the names confirm heavy reporting (`REPORT_LogicalInventory`, `REPORT_NUMMILotLocationW`, `REPORT_EDI810Recreate`, etc.). But "0.5-1 day each" is unverified for layouts like the `...LotLocationW`/`PLANTLotLocation` matrix reports, which mirror the PartsStockMaster matrix complexity. The proposal already counts PartsStock as a 1-2 week worst case but assumes reports are uniformly cheap. Calibrate one matrix report (`REPORT_NUMMILotLocationW`) in the same spike. Settle the estimate with one real build.

UNVERIFIED / NOT CHECKED (stated plainly, not claimed as wrong):
- I did not validate Perspective licensing math or the Reporting-module separate-license claim — no source in-repo; get the real quote (proposal already flags this).
- I did not confirm Jython 2.7 cannot do the X12 work — I agree it can for pure string ops; the risk is scope (7 doc types + AK1/997 loop), not language.

KILLER RISKS (go/no-go):
1. Dual-process contention on the EDI shared directory during parallel run (inbound archive-move race + outbound name-collision / non-atomic write). FATAL if not addressed, MITIGABLE by single-owner EDI I/O during parallel run + atomic temp-then-rename. This is the one that loses a 997 ack or ships a truncated 810 to Toyota's VAN.
2. EDI re-scope: 7 inbound doc types incl. functional-ack reconciliation, not "read 830"; outbound spec must be re-derived from `ASNInvoice.pas`, not the dead `Write810File.pas`. NOT fatal to platform choice, but FATAL to the current effort estimate and the "Python tier evaporates" claim until redone. MITIGABLE by re-reading the two live units.
3. Multi-site enforcement-by-discipline. MITIGABLE, correctly flagged, gated on the isolation spike.

Bottom line: the Ignition choice is defensible and the carryover accounting is honest — but the EDI section, the one place the proposal claimed its biggest win, rests on dead code (`Write810File.pas`) and a one-document inbound assumption that the live `EDIUpload.pas` contradicts. Conditional GO stands, but add a third spike gate: re-scope EDI against `EDIUpload.pas` + `ASNInvoice.pas` and prove single-owner shared-dir I/O before sizing.

Files cited: `/Users/apple/Documents/GitHub/InventorySystem/Write810File.pas:33-118` (dead body); `/Users/apple/Documents/GitHub/InventorySystem/ASNInvoice.pas:820-879` (live outbound writer); `/Users/apple/Documents/GitHub/InventorySystem/EDIUpload.pas:53,89,105,186,192,253,306,418-439` (live inbound: 7 doc types, archive-move, site filter); `/Users/apple/Documents/GitHub/InventorySystem/EDI810Object.pas:95,142-442` (TStringList segment building, no file I/O); `/Users/apple/Documents/GitHub/InventorySystem/DB Schema/Create Inventory.sql` (179/29/24 confirmed).
