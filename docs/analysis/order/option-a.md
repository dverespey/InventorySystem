# Order — Ignition Perspective Target Design — OPTION A "FAITHFUL-MODERN"

Status: design proposal. Source truth: `legacy-order-spec.md` (cited as **spec §N**),
`source-artifacts.md` (cited as **artifacts §N**). Locked decisions per task brief
(interactive Perspective grid, single site-scoped screen D1, wrap-the-proc, 8.3-design /
8.1.52-runnable, solo-dev maintainability).

**Thesis of Option A:** the legacy simulation algorithm (spec §3) and its color *semantics*
(spec §4) are trusted daily and reproduced **bit-for-bit**. Researched presentation best
practice is applied **only to how the result is shown** — exception-first ordering,
accessible (never color-only) signaling, frozen header, server-side site scoping. We do not
alter what a number/color *means*; we change only the surface and the substrate (Excel-OLE →
Perspective + DB proc, not a re-derived algorithm).

> ~~OPEN GAP (process gap)~~ **RESOLVED:** `docs/analysis/order/research-best-practice.md` now
> exists on disk (fact-checked; every retained recommendation backed by a live URL). The
> presentation choices below (exception-first ordering, WCAG non-color-only signaling, frozen
> header/sticky first columns, ARIA/text-redundant status, grid performance) are now **cited to that
> research doc** (per-choice citations in §1.2/§5/§9: grid-ux NN/g + Smashing + Perspective Table
> docs; WCAG 1.4.1 + Section 508). They are no longer "reviewer-judgment defaults."
> ONE deliberate **deviation from research, in service of faithfulness:** research §6.7 recommends
> replacing the legacy raw Red-Amber-Green with a CVD-safe palette (Okabe-Ito / IBM). Option A does
> **NOT** adopt the palette swap — it preserves the legacy Excel ColorIndex RGBs bit-for-bit (spec
> §4). Instead it satisfies the governing rule (research §6.6 / WCAG 1.4.1, Level A: status is
> *never* color-only) by adding non-color channels (icon + text + ARIA) on top of the unchanged
> colors. The CVD-safe palette is recorded as an available post-cutover improvement (§5 note), gated
> on David's sign-off, not a faithfulness change. This does not affect Option A's faithfulness
> guarantee, which is anchored to the spec (the color *meanings* in §5 map only to Pascal-set
> signals from spec §4, cross-checked against the color map).

---

## 0. Decision summary (the load-bearing choices)

| Question | Option A answer | Why |
|---|---|---|
| Where does the simulation calc run? | **New DB stored proc `SIM_OrderSimulation`** (T-SQL), not gateway Jython | §2 below — keeps the trusted math next to the data it already reads, one round trip, server-side site scope, no Jython port of a 600-line Pascal algorithm |
| How is the grid rendered? | Perspective **Table** (phased grid) + header **Flex** (Select-Order controls), bound to one Named Query | §3 |
| How are colors carried? | **Per-cell semantic enum** emitted by the proc; Perspective maps enum → {background, **icon/glyph**, **text label**, ARIA} | §5 — preserves spec §4 meaning, adds non-color channel |
| Commit path | `system.db.createSProcCall` wrapping existing `SELECT/UPDATE_PartsStockRenban` + `INSERT_OpenOrder`, **but inside a serializing gateway transaction + DB commit-claim** (the bare procs race — §4.1) | §4, §6 |
| Multi-site | One **screen** (K6/L6 cosmetic). But **data scoping is NOT a wrapper** — there is no `site_id` column in the schema today; scoping requires schema+proc surgery (§7 OPEN GAP) | §7 |

---

## 1. Perspective view structure

Three views, one page. All data access goes through the server-side `siteScopedQuery()`
wrapper (D1); `site_id` is derived from `session.props.auth` / a custom session prop, **never**
a client/page param. **Caveat (see §7):** `siteScopedQuery()` can only *carry* `@site_id` to the
procs — it cannot make a site-blind proc filter by site. Today the `INV_*` tables and the START
procs have **no site column at all**, so until the schema-scoping work in §7 lands, the wrapped
procs are site-blind and the "no leak" guarantee is **not yet in force**.

```
Page: /order
└─ view: Order/OrderPage              (coordinates state, owns custom session-scoped props)
   ├─ view: Order/SelectOrderBar      (the legacy Select-Order dialog, spec §1 / artifacts §1)
   └─ view: Order/PhasedGrid          (the OrderSimulation sheet, spec §3 / artifacts §2-3)
```

### 1.1 `SelectOrderBar` (replaces `Order.dfm`)

Flex row (wraps on mobile). Components and their legacy origin:

| Component | Legacy | Behavior |
|---|---|---|
| **Today** (read-only Label) | `Today` field, read-only (spec §1, artifacts §1) | `now()` on the gateway at simulate time; display only, **is the order date**. Bound to the value the proc actually used (returned in result metadata) so screen and calc never disagree. |
| **Line** (Dropdown) | `Line_ComboBox` | Options from a site-scoped NQ over the lines master. Blank allowed. |
| **Part Type** (Dropdown) | `PartType_ComboBox` TIRE/WHEEL/VALVE/FILM (artifacts §1: file/valves/tires/wheels) | Options from config (site-scoped); see §7 gap note on the type list. |
| **Sort By** (Dropdown) | `SortBy_ComboBox`, **hidden unless Line blank** (`Order.pas:1659`, spec §1) | `visible` bound to `{Line} == ''`. When Line blank → "all parts of a part type," parts first sorted by size to form size/sharing pairs (artifacts §1). |
| **Simulate** button | **Start** (`Order.pas:146`) | Calls `SIM_OrderSimulation` (read-only), populates `PhasedGrid`. Enabled when Part Type set. |
| **Commit Orders** button | **Order** (`Order.pas:628`) | **Disabled until a simulate result exists** (legacy: Order enabled only after Start, artifacts §1). Opens confirm dialog → §4 commit action. |
| **Reset/Exit** | Exit | Clears grid + simulate token. |

State props on `OrderPage` (session/page scoped, not passed as URL params):
`simToken` (uuid of the last simulate run), `simTodayDate`, `simParams` (line/partType/sortBy),
`siteId` (session-derived, read-only mirror for binding convenience — authority stays server-side).

### 1.2 `PhasedGrid` (replaces the OrderSimulation.xls sheet)

A Perspective **Table** (not 200 hand-laid cells). Faithful to spec §1/§3 layout but
presented as rows of structured records, with the phased day-columns as **dynamic columns**
driven by the fill-days config (artifacts §2-3: render N day-columns from fill-days).

Row model (one row per legacy part row, plus size-group header/footer rows — see spec §3.2):

- **Frozen left columns** (cited best practice — NN/g + Pencil&Paper: when a table exceeds the
  screen, freeze the leftmost human-readable identifier columns with a subtle drop shadow on the
  frozen edge; Perspective supports frozen columns natively — research §4 grid-ux):
  Size (B), Brand/Supplier (C), Part No. (D) — the identity columns. Legacy froze nothing;
  this is presentation-only and does not change data.
- **Order-Point block** (spec §3.2): Daily Usage (H), Safety Days (I), Safety Stock (J).
- **Inventory block**: Total Inv (J/Total), `<SITE>` warehouse (K), `<DEST>` plant (L),
  In Transit (M), Open (N). **K/L headers are site labels** (artifacts §3 — config, §7).
- **Order block**: 1Lot Qty (O), Lead Time (P), **Qty (Q, editable)**, **Lot (R, editable)**.
- **Phased day-columns** `day[0..FillDays-1]`: each a date-headed numeric column. Header row
  carries the rendered date + weekday (spec §3.1; OrderSimulationChanged.xls grid, artifacts §3).
- **Summary columns** (spec §1, §3.4): Total Inv / In Transit / Added Leadtime.

Best-practice presentation layered on top (NONE alters legacy numbers). Each is now cited to
`research-best-practice.md` (not reviewer judgment):

- **Frozen header row** (dates/weekday) — sticky while scrolling 200 rows. CITED: NN/g + Pencil&Paper
  — freeze the header row when the table exceeds the screen; Perspective Table supports frozen
  rows/columns natively (research §4 grid-ux).
- **Exception-first ordering toggle**: default sort surfaces rows whose projected balance
  breaches safety stock (spec §3.5 stockout flag) and rows whose order-by day is today/past
  (spec §3.4) to the top. A "show all / faithful order" toggle restores the legacy
  size-grouped order so the trusted reading is always one click away. CITED: operational
  dashboards must surface actionable anomalies first / "steer users toward what matters" (Smashing
  real-time-dashboards), with a sensible default sort and a discoverable "filters active" indicator
  and a clear empty/zero state (NN/g data-tables) — research §4 grid-ux + design implication §9.
  The "faithful order" toggle is Option-A-specific (keeps the trusted legacy reading reachable);
  exception-first as the *default* is the cited recommendation.
- **Grid performance**: enable Perspective Table `virtualized` + `pager` with a fixed row height;
  client-side rendering is fine at the legacy ≤200-row cap (spec §1, hazard 3), well under the
  cited ~1,000-row client/server threshold, so no server-side paging is needed for the simulate
  result. CITED: virtualization renders only the visible slice + buffer with fixed row heights;
  client-side under ~1,000 rows, server-side above; Perspective Table has a `virtualized` property +
  configurable `pager` (research §4 grid-ux). (Note: the research's "scroll feel / query
  re-execution during scroll" forum claim is explicitly unverified there and is NOT relied on.)
- **Two editable columns only** (Q, R) — exactly the legacy cream input cells (spec §4,
  `Order.pas:539-540`). Everything else read-only. Edits are local until Commit. (Faithfulness, not
  a research choice.)

### 1.3 What is deliberately NOT replicated

- Excel OLE, `excel.visible:=True`, orphaned `EXCEL.EXE`, locale-sensitive string read-back,
  cell-lock dance (spec §8 hazard 2) — all eliminated by moving off Excel.
- In-cell Excel formulas (`=E/(ΣE)`, end-balance recurrence) — recomputed in the proc as
  values (§2), because there is no spreadsheet to host formulas. **Faithfulness requirement:**
  the proc must reproduce the same arithmetic (spec §3.2 share %, §3.5 end-balance recurrence
  `EndBalance = Beg + Receipts − Usage`, `BegBalance(day+1)=EndBalance(day)`).

---

## 2. WHERE the simulation calc runs — DB proc, justified

**Decision: a NEW T-SQL stored proc `SIM_OrderSimulation` on Inv_Connection, returning the
fully-computed grid. NOT gateway Jython.**

This is the one place Option A adds DB logic rather than purely wrapping — justified because
the legacy "simulation" was never in a proc; it lived in Pascal driving Excel (spec §3). We
must put it *somewhere*. The choice is DB proc vs gateway Jython.

Why **DB proc** wins for Option A:

1. **Data locality / round-trips.** The Start path fires ~12 distinct procs *per part row*,
   reused on a shared dataset across ≤200 rows (spec §2 START table; §8 hazard 5). Done from
   Jython that is hundreds of `createSProcCall` round trips per simulate — slow and chatty.
   In a proc, the same reads are set-based joins/cursors next to the data.
2. **Faithfulness is easier to *prove* in T-SQL — but it is asserted, not yet shown.** The
   legacy reads are already T-SQL procs (spec §2). Re-expressing the calendar walk (spec §3.1),
   lead-time selection by weekday (spec §3.2), order-point + added-leadtime (spec §3.4) and
   end-balance recurrence (spec §3.5) in T-SQL keeps them in the same language/semantics as the
   source data — fewer type/locale translation bugs than the Pascal→Jython hop (spec §8 hazard 2
   was *caused* by string read-back across the OLE boundary; we don't want to reintroduce a
   boundary). **However, the calc is sound in INTENT, not yet proven** — the proof is the §9.2
   parity diff vs a live legacy Start, which is gated behind the open extraction gaps below.
   **NAMED PARITY HAZARD (do not treat as an assumption) — hazard 7 (spec §8;
   `Order.pas:1353-1369,1457-1468`):** `fDates` is indexed by **calendar offset `x`** while
   `fForecast`/phased arrays are indexed by **fill position `j`**, reconciled only by the
   `fDates[i] <> 0` scan. A set-based T-SQL rewrite that does not preserve this exact
   two-index-space scan will **misplace in-transit / open-order buckets by a day**. The proc must
   reproduce both index spaces and the reconciling scan, and the §9.2 parity diff must include a
   **named test case** that exercises a calendar with non-production days between fill positions
   (so offset≠position) and asserts bucket placement byte-for-byte. This is a parity *test*, not
   a design assumption.
3. **Server-side site scope (D1) — single place to add scoping.** A proc takes `@site_id` and
   every internal read can be scoped in **one body**, vs Jython threading site_id through every
   sub-call. NOTE: this is the *right place to add* scoping; the scoping itself does **not exist
   yet** (no site column in the schema — §7 OPEN GAP). The proc is where that surgery lands.
4. **Single Named Query → single-point maintenance** (memory `ignition-named-query-crud-practice`):
   one NQ `Order/sim/SIM_OrderSimulation` wraps the whole simulate. A schema change is one edit,
   often zero Ignition-side edit.
5. **Solo-dev maintainability.** One proc to test against the legacy (run both, diff the grid)
   beats a sprawling Jython module replicating Excel layout math.

Why **not gateway Jython** (recorded so the choice is defensible): Jython 2.7 is fine for
orchestration but a poor host for a 200-row × 23-day numeric simulation with per-row multi-proc
fan-out; it would multiply round trips and move the trusted math *away* from the data into a
weaker-typed language.

> ⚠️ FAITHFULNESS BOUNDARY / EXTRACTION-GAP DEPENDENCY: `SIM_OrderSimulation` reproduces the
> Pascal algorithm (spec §3) **and only that**. It does **NOT** invent thresholds/colors that
> live in `OrderSimulation.xls` and were unreadable (artifacts §4 gaps 1-3; spec §5).
> Specifically the proc emits *only* the signals Pascal sets in code (spec §4 rows attributed
> to `Order.pas:*`). Any conditional-format rule baked into the template (artifacts §4 gap 1)
> is **out of scope until the .xls CF rules are extracted** — flagged in §5/§8.

`AD_GetSpecialDate` (overtime/holiday calendar, spec §2 START #1) lives in the **ALC `TireOrder`
DB**, NOT this schema (spec §8 hazard 1, **body unverified**).

> **OPEN GAP — BLOCKED (source truth — delphi-architect / ALC schema):** the implementation
> choice is **deliberately NOT made yet**, to avoid hardening a guess. Two faithful candidates
> exist — (a) cross-DB call `TireOrder.dbo.AD_GetSpecialDate` (linked / same-instance), or (b) a
> gateway pre-step that fetches the calendar and passes it as a TVP to `SIM_OrderSimulation` —
> but **both are held** until delphi-architect returns the proc **body, DB/catalog location, its
> result-set shape, and the exact status-code domain** ('O' / 'X' / holiday and any others). The
> entire calendar walk (spec §3.1) and the added-leadtime math (spec §3.4, verified at
> `Order.pas:1576-1582 DoLeadTime`) consume these statuses, so picking cross-DB-vs-TVP before the
> result shape and status domain are confirmed risks baking in a wrong contract. Do **not**
> re-derive the overtime/X/holiday rules — get them from source. (§9 phase 1 resolves this
> gap first.)

---

## 3. Data layer — Named Queries & SProc-calls (mirror the schema/procs)

NQ tree mirrors procs (memory `ignition-named-query-crud-practice`). Every NQ runs through
`siteScopedQuery()`; `@site_id` injected server-side. (Reminder: injection ≠ isolation until the
procs/tables carry a site column — §7 OPEN GAP.)

### 3.1 SIMULATE path (read-only)

| Named Query | Wraps | Type | Notes |
|---|---|---|---|
| `Order/sim/SIM_OrderSimulation` | **NEW** `SIM_OrderSimulation(@site_id,@LineName,@PartType,@SortType,@Today,@FillDays,@ForecastUsageCompare,@UseFirstProductionDay)` | SProc, returns grid result set(s) | The whole Start path (spec §3) collapsed into one server call. Internally calls/inlines the spec §2 START procs #2-#15. |
| `Order/lookup/SelectLines` | site-scoped lines lookup | Query | Line dropdown options |
| `Order/lookup/PartTypes` | site config part-type list | Query | Part-type dropdown (see §7 gap) |

The §2-START procs (#1-#15) are **invoked from inside `SIM_OrderSimulation`**, preserving
their bodies (wrap-the-proc, phase 1 — do not bypass). They are not bound directly from
Perspective; the screen only ever calls the one sim NQ. (`AD_GetSpecialDate` = the cross-DB
exception, §2.)

`SIM_OrderSimulation` return contract (one row per grid row):
`row_kind` (`SIZE_HEADER|PART|SIZE_FOOTER`), `size_code`, `supplier`, `part_number`,
`daily_usage`, `safety_days`, `safety_stock`, `total_inv`, `wh_qty(K)`, `plant_qty(L)`,
`in_transit(M)`, `open_order(N)`, `lot_qty(O)`, `lead_time(P)`, `qty_default(Q)`, `lot_default(R)`,
`share_pct`, `added_leadtime`, `orderby_col_index`, `leadtime_zone_end_index`,
`frs_date`, `kanban`, `renban_group`, plus a **phased payload**:
`day[0..FillDays-1].{value, source_enum, balance, signal_enum}` (JSON column or a child result
set keyed by row+day). Plus a header result set: `day_index → {serial_date, weekday, day_kind}`
where `day_kind ∈ {NORMAL, OVERTIME, NONPRODUCTION}` (spec §3.1).

### 3.2 COMMIT path (writes) — proc bodies unchanged, but wrapped in a serializing tx (spec §2 ORDER, §6)

Called as discrete `system.db.createSProcCall` inside a **single `SERIALIZABLE` gateway
transaction that first takes `UPDLOCK,HOLDLOCK` on the contended rows** (the bare procs race —
§4.1) plus a `simToken` commit-claim insert. This is stricter than the legacy
`BeginTrans/CommitTrans` per insert (spec §2 note): the legacy got away with no locking because
it was single-user; the parallel run cannot.

| Named Query / SProc-call | Wraps | Params | Effect |
|---|---|---|---|
| `Order/renban/SELECT_PartsStockRenban` | `SELECT_PartsStockRenban;1` | `@PartNum` → `IN_RENBAN_COUNT` | read renban counter (spec §6) |
| `Order/renban/UPDATE_PartsStockRenban` | `UPDATE_PartsStockRenban;1` | `@PartNum,@RenbanCount` | bump counter, wrap >999→1 (spec §6) |
| `Order/order/INSERT_OpenOrder` | `INSERT_OpenOrder;1` | `@SupCode,@PartNum,@KanbanNum,@FRSNum,@RenbanNum,@Qty` (all IN; **no OUT/return** — spec §6, confirmed in schema `CREATE PROCEDURE INSERT_OpenOrder`) | inserts `INV_OPEN_ORDER_INF`; **server computes FRS year-roll + FRS sequence suffix** (spec §6). |

**Do NOT touch the qty trigger** (`INSERT_RecConfStatPartsStockMstQTY`, spec §6, `triggers.sql:214`):
new orders have empty shipping status → stock NOT bumped at creation (spec §8 hazard 8). Leave
it in the DB; a rebuild that "adds qty on order" is wrong.

---

## 4. The commit action (gateway-script, serialized + replay-safe under parallel run)

Commit is a **single message-handler / gateway script** (not per-row client calls) so the whole
order set commits atomically per part and the FRS/renban math stays server-side.

Flow (faithful to spec §6, ORDER path `Order.pas:628`):

```
on Commit(simToken, edits[]):
  reject if simToken != OrderPage.simToken      # only commit what was simulated
  reject if any edited row's underlying part changed since simulate (optimistic stamp)
  begin one SERIALIZABLE tx; INSERT commit-claim row keyed by simToken (unique) — see §4.1
  for each part row with a part number:
     qty = edits.Q ; lot = edits.R              # read back the two editable cols (spec §6)
     # FIDELITY GATE (Order.pas:656, confirmed): legacy processes the row ONLY if
     #   qtycount <> '' AND lotcount <> ''  — an EMPTY Lot cell skips the row entirely,
     #   not merely qty==0. Mirror exactly:
     SKIP row unless (Q is non-empty AND R is non-empty)
     validate numeric
     if BIT_LOT_SIZE_ORDERS:                      # lot-size path (spec §6)
        one INSERT_OpenOrder, @Qty = qty, FRS suffix seed '..01' (proc overwrites trailing 2)
     else:                                        # FRS-breakdown path (spec §6)
        for j in 1..lot: INSERT_OpenOrder, @Qty = IN_1LOTQTY (1-lot qty, NOT typed qty)
     renban: if part NOT in renban group:
        # take UPDLOCK,HOLDLOCK on the part's INV_PARTS_STOCK_MST row first (§4.1)
        c = SELECT_PartsStockRenban(@PartNum=...)  # bind by proc's REAL param name @PartNum,
                                                   #   NOT @PartCode — legacy used positional
                                                   #   ADO binding (Order.pas:715) so the name
                                                   #   mismatch was invisible there; named
                                                   #   binding in Ignition MUST use @PartNum.
        @RenbanNum = Kanban + zeropad3(c)
        ... INSERT_OpenOrder ...; UPDATE_PartsStockRenban(@PartNum, c+1, wrap>999→1)
     else: @RenbanNum = '' (numbering deferred to renban screen)
  commit tx (single tx spans claim + all of this part's inserts); on any error rollback all
```

### 4.1 Concurrency & re-submit safety under parallel run (legacy + Ignition both hit one DB)

This is the critical correctness concern. **CORRECTION (was previously mis-stated as
"idempotent / clash-free"):** the bare `INSERT_OpenOrder` proc is **NOT** safe under
concurrency, and Option A no longer claims it is. Verified body (`Create Inventory.sql:3266-3314`):
the proc does an unguarded `SELECT @MaxFRS = max(VC_FRS_NUMBER) ... WHERE ... LIKE @FRSNum+'%'`,
computes `RIGHT(...CAST(RIGHT(@MaxFRS,2))+1...)`, then `INSERT`s — **no `UPDLOCK`/`HOLDLOCK`, no
content unique constraint, no transaction in the proc body.** This is a textbook
read-max-then-insert race. Two concurrent commits on the same part / FRS-prefix scope both read
the same `@MaxFRS` and both compute the **same** trailing-2 suffix → **duplicate FRS numbers**,
not a safe interleave. The legacy is safe **only because it is operationally single-user**, not
because the proc serializes anything. The same hazard applies to the renban read-bump
(`SELECT_/UPDATE_PartsStockRenban`).

Option A's parallel-run window adds multi-session / multi-node surface the single desktop user
never had, so faithfulness here means **adding serialization the legacy got from being
single-user**, not copying the unguarded proc:

- **Serialize the FRS read+insert AND the renban read+bump in ONE transaction with a write lock.**
  Both the `max(VC_FRS_NUMBER)` read and the `INSERT`, and the renban `SELECT`+`UPDATE`, must run
  inside a single `system.db.beginTransaction` at `SERIALIZABLE`, taking an `UPDLOCK, HOLDLOCK`
  on the contended scope (the part's `INV_PARTS_STOCK_MST` row and the FRS-prefix range) so a
  Delphi commit and an Ignition commit on the same part **block** instead of both reading the
  same max. This is a behavior **change vs the bare proc** and requires either (a) wrapping the
  existing proc inside a gateway transaction that first `SELECT ... WITH (UPDLOCK, HOLDLOCK)`s the
  contended rows, or (b) re-authoring the proc to take the lock internally.
  > **OPEN GAP (source truth — delphi-architect):** is commit truly single-writer during the
  > planned parallel run (e.g. only one site/operator commits at a 5am window), or can Delphi and
  > Ignition genuinely commit the same part concurrently? If single-writer is guaranteed
  > operationally, the lock is belt-and-suspenders and exact faithfulness is preserved. If NOT,
  > the serialized FRS+renban write is **mandatory before any commit-phase ships** (§9 phase 4 is
  > gated on this). Either way, do **not** present the bare `INSERT_OpenOrder` as idempotent.
- **Server-side commit-claim (DB-level re-submit guard, not just the button).** The client-side
  `simToken` + disabled button only defends against a double-click in one session. A gateway
  retry after a timeout where the DB **did** commit, two browser tabs, or a session failover
  between gateway nodes all bypass in-memory state and would re-run the whole `edits[]` loop →
  a second full set of open orders. Defense: persist a **commit-claim row** keyed by `simToken`
  (a small new `INV_ORDER_COMMIT_CLAIM` table or equivalent) and `INSERT` it **inside the same
  transaction** as the order inserts, with a unique constraint on `simToken`. A replay then fails
  the unique insert and the whole tx rolls back — the duplicate is rejected at the DB, not the
  UI.
  > **OPEN GAP (source truth — delphi-architect):** confirm there is no existing dedup column or
  > natural key on `INV_OPEN_ORDER_INF` we can reuse before adding a claim table (grep confirms no
  > unique key on (part, FRS, renban, qty); the proc returns **no** OUT/value the caller checks).
  > Adding a claim table is a schema add the legacy did not have — get sign-off that it does not
  > break the parallel Delphi writer (it only INSERTs a new table, so it should not, but confirm).
- **Simulate→commit staleness:** a sim result can age (open orders/stock change between
  Simulate and Commit, possibly via the parallel Delphi app). Option A guards with the
  `simToken` + per-row optimistic check above and a **"re-simulate recommended"** banner if the
  underlying `IN_RENBAN_COUNT` / open-order set changed. The legacy had **no** such guard (it
  read the live sheet); this is a safety add that does not change committed values.
- **Double-click / resubmit (UI layer, secondary):** Commit button disables on click and the
  handler is keyed by `simToken` — but this is now treated as a UX convenience only; the DB-level
  commit-claim above is the authoritative guard. (Store-and-Forward does NOT apply here — these
  are transactional proc calls, not tag history.)

---

## 5. Accessible color/signal model — preserve meaning, never color-only

Legacy uses Excel `ColorIndex` on interior (zone) and font (qty-source) channels (spec §4).
Option A preserves the **meaning** of each (faithful, anchored to spec §4 / `Order.pas:*`) and
renders it with **≥2 channels** (color + icon/glyph + text/ARIA). The multi-channel /
non-color-only rendering is now a **CITED best practice — the governing accessibility rule**, not
reviewer judgment: WCAG 1.4.1 Use of Color (Level A) — "color is not used as the only visual means
of conveying information"; the Section 508 redundant-coding heuristic (mark status in color **and**
a word, supplement with icons + text); the acid test "remove all color, still fully usable"
(research §2 reorder-viz + §6.6). Contrast minimums also apply (research §2): status fill/border and
icon each ≥ 3:1 (WCAG 1.4.11), in-cell text ≥ 4.5:1 (WCAG 1.4.3) — validate the final hex pairs.
The *meanings* below are faithful to spec §4; the *extra channels* implement the cited WCAG/508
requirement on top of the unchanged legacy colors.

**Deliberate deviation from research §6.7 (recorded, not silently dropped):** research recommends
replacing the raw Red-Amber-Green (worst CVD pairing) with a CVD-safe palette — Okabe-Ito (`#009E73`
/ `#E69F00` / `#D55E00`) or IBM (`#648fff` / `#ffb000` / `#dc267f`). Option A's faithfulness
guarantee keeps the legacy Excel ColorIndex RGBs **bit-for-bit** (spec §4: 3/4/10/23/34/36/40), so
the palette swap is **NOT** adopted in the faithful build. Adding the non-color channels (icon +
text + ARIA) already satisfies WCAG 1.4.1 — the governing Level-A rule — without changing the
colors. The CVD-safe palette is logged as a candidate post-cutover improvement requiring David's
sign-off (it changes what the operator sees, even though meanings are preserved), not a faithfulness
change.

`SIM_OrderSimulation` emits **two enums** per relevant cell — `signal_enum` (interior/zone) and
`source_enum` (font/qty-source) — so the meaning is in the data, not in a Perspective style
guess. Perspective maps enum → style class + glyph + accessible label.

| Legacy (spec §4) | Meaning | `enum` | Color (start = standard Excel-2000 palette, see gap) | Non-color channel (added) |
|---|---|---|---|---|
| Interior 36 zone (`:1587`) | within lead-time window | `LEADTIME_ZONE` | pale yellow | left border bracket + cell title "lead-time" |
| Interior 40 (`:1588`) | **order-by point** | `ORDER_BY` | cream | **★ glyph + "ORDER BY" text badge** (most important signal) |
| Interior 3 (`:1592`) | overtime production day | `OVERTIME` | red | hatch pattern + "OT" tag in day header |
| Interior 4 (`:1597`) | non-production 'X' day | `NONPROD` | green | striped pattern + "X" tag in day header |
| Font 23 (`:1342/1377`) | qty is **in transit** | `SRC_INTRANSIT` | dark blue | "in-transit" icon (truck) + tooltip |
| Font 10 (`:1449/1476`) | qty is **open order** | `SRC_OPENORDER` | dark green | "open-order" icon (box) + tooltip |
| FormatCond font 3 (`:1047/1055`) | over-produced (J<0) | `OVERPRODUCED` | red | "▲ over" label |
| FormatCond font 3 (`:1552/1558/1564`) | end-balance **< safety stock** | `BELOW_SAFETY` | red | **"⚠ below safety" text + bold** (stockout alert) |
| Interior 34 (`:932`) | size daily-usage/safety inputs | `SIZE_INPUT` | pale cyan | read-only styling (these are no longer typed) |

Rules:
- **Interior vs font are distinct channels** (spec §4 note) — `signal_enum` and `source_enum`
  are independent and can co-occur on one cell (e.g. an in-transit qty inside the lead-time
  zone): background = `LEADTIME_ZONE`, glyph = truck. Do not collapse them.
- Day-header zone tags (OT/X) move the per-column overtime/non-production signal into the
  **header**, faithful to "overtime/non-production columns" (spec §3.4) but readable without
  scanning every cell.
- A **legend** component decodes every enum (text + swatch + glyph) — CITED: Section 508
  redundant-coding (status in color **and** word/icon) and the "remove all color, still usable"
  acid test both imply a decodable key; this is the thing the trusted Excel sheet never had
  (research §2 reorder-viz).

> ⚠️ EXTRACTION-GAP DEPENDENCY (spec §5; artifacts §4 gaps 1,2,4): the **RGB values** behind
> indices 3/4/10/23/34/36/40 are assumed standard Excel-2000 palette and are **UNCONFIRMED** —
> the workbook may override the palette (spec §5 item 4). AND any **template-baked
> conditional-format rules + their numeric thresholds** on the phased range / summary columns
> are **completely unread** (artifacts §4 gap 1; spec §5 item 2). Option A reproduces only the
> **Pascal-set** signals (the table above — all attributed to `Order.pas:*`). If the .xls
> carries additional CF thresholds, **this grid is missing them.** ACTION: extract the .xls CF
> rules + palette (re-save to .xlsx, read `worksheet.conditional_formatting` via openpyxl, per
> artifacts §4) BEFORE sign-off. Until then the color *exact RGBs* and any extra thresholds are
> flagged UNKNOWN — not guessed.

---

## 6. Simulate → review → commit loop (state machine)

```
EMPTY ──Simulate──▶ SIMULATED(simToken,today,params) ──edit Q/R──▶ DIRTY
   ▲                      │                                          │
   │                      └────────── Re-Simulate (new token, discards edits w/ confirm)
   │                                                                  │
   └────────────── Commit succeeds ◀── COMMITTING ◀───── Commit(simToken,edits) ◀┘
                         (toast: N orders created, FRS list)         │
                                          rollback ▶ SIMULATED + error banner
```

- Simulate is **read-only** (spec §1: Start does no writes). Safe to re-run any time.
- Commit enabled only in SIMULATED/DIRTY with a live token (legacy: Order enabled only after
  Start — artifacts §1).
- On commit success, surface the created FRS/renban numbers (the proc computed them) so the
  user sees what the trusted system produced.

---

## 7. Multi-site (D1) — single screen (cosmetic), but invasive data scoping

**Two separable claims — keep them distinct:**

**(a) SCREEN count = ONE (SOUND).** Template diff is **COSMETIC** (artifacts §3): every
production variant is byte-identical in xlrd cell content except **K6/L6** (source warehouse /
destination plant labels), and the app opens a single fixed `OrderSimulation.xls` with no
per-site branching in `Order.pas`. So ONE site-scoped screen serves all; no per-site fork.

**(b) DATA scoping is NOT a `siteScopedQuery()` wrapper — it is schema + proc surgery (was
overstated).** Verified: grep for `site_id` / `VC_SITE` / `IN_SITE` across
`DB Schema/Create Inventory.sql` returns **zero hits** — the legacy is single-site by INI only.
`@site_id` therefore **cannot be "injected"** into `SELECT_PartsStockInfoOrder`,
`INSERT_OpenOrder`, etc., because those procs have no site column to filter on. A wrapper that
passes `@site_id` to a proc that ignores it provides **no isolation**; on a multi-site DB,
`SELECT_PartsStockInfoOrder` would return all parts of a type regardless of site
(cross-contamination). The previous claim that "no query can leak across sites" from the wrapper
alone is **withdrawn.**

> **OPEN GAP (source truth + schema surgery — delphi-architect):** site scoping requires
> (i) **adding a site column** to the shared tables both apps write — at minimum
> `INV_PARTS_STOCK_MST`, `INV_OPEN_ORDER_INF`, `INV_FORECAST_DETAIL_INF`, `INV_SIZE_MST`
> (delphi-architect to enumerate the full set against the §2 START reads), and (ii) **re-authoring
> every proc body** that reads/writes them — the 15 START procs (spec §2) + the 3 commit procs —
> to filter/stamp by site. **HAZARD (F1, see memory):** adding a `site_id` column breaks any
> `_HIST` trigger that does `INSERT INTO ..._HIST SELECT * FROM inserted` unless the matching
> `_HIST` table also gets the column — confirm every affected `_HIST` table is widened in lockstep
> (`docs/triggers.sql`). This is the dependency that gates §9 phase 1; it must be scoped by
> delphi-architect, **not** assumed solved by the wrapper. The `sites` config table below is
> necessary but **not sufficient** on its own.

Once (i)+(ii) land, ONE site-scoped screen + the `sites` config table below cover all variants:

| Per-site difference (source) | D1 handling |
|---|---|
| K6 = source warehouse label (WQS/HERO/CAMEX/WWW/MAS), L6 = dest plant (NUMMI/TMMTX/DEPOT/SIA/TMMM) (artifacts §3 table) | Columns `sites.warehouse_label`, `sites.plant_label`. `PhasedGrid` K/L header bound to them. Pure labels (artifacts §3). |
| Template path / which `OrderSimulation*.xls` (spec §7) | **N/A** — no Excel; the proc + screen are one, parameterized by `@site_id`. The whole "pick template per site" problem evaporates. |
| INI `[INIT] FillDays`(23, max 50), `ForecastUsageCompare`(7), `UseFirstProductionDay` (spec §7, artifacts §1) | `sites.fill_days`, `sites.forecast_usage_compare`, `sites.use_first_production_day` — passed as params to `SIM_OrderSimulation`. Enforce max 50 (artifacts §1). |
| INI `[SITE] PlantName`(NUMMI), `AssemblerName` (spec §7) | `sites` rows; used in labels/error text. |
| Two databases: Inv_Connection + ALC `TireOrder` for calendar (spec §7, §8 h1) | site → which ALC catalog/connection for `AD_GetSpecialDate`. **Depends on resolving the cross-DB gap (§2/§8).** |

`siteScopedQuery()` injects `@site_id` (session-derived) into every NQ — including
`SIM_OrderSimulation` and the three commit procs — and keeps `site_id` off the client (never a
client param, D1). **But injection alone does not isolate data**: it only achieves leak-proofing
**after** the procs/tables actually filter on that column (the schema surgery above). The wrapper
is the *delivery mechanism* for `@site_id` and the *enforcement point that it comes from the
session, not the client* — it is **not** by itself the isolation guarantee.

> The structural outlier `OrderSimulationChanged.xls` (day-grid, artifacts §3) is **not** a
> per-site fork — it is a newer *version* showing the fill-days grid as columns. Option A's
> `PhasedGrid` already renders N day-columns from `sites.fill_days`, so it natively matches that
> "changed" variant. No second screen.

---

## 8. OPEN GAPS — must resolve before the relevant build phase signs off

Each gap names its **owner** (delphi-architect = source truth, or an extraction step) and the
**phase it blocks**. None of these are hand-waved as solved.

1. **OPEN GAP — commit concurrency (delphi-architect; blocks §9 phase 4).** The bare
   `INSERT_OpenOrder` (`Create Inventory.sql:3266-3314`) and the renban read-bump are unguarded
   read-max-then-insert races; safe in the legacy only because single-user. Confirm whether
   parallel-run commit is operationally single-writer. If not, the serialized FRS+renban write
   (single `SERIALIZABLE` tx + `UPDLOCK,HOLDLOCK`, §4.1) is **mandatory**. Do not ship commit
   until resolved.
2. **OPEN GAP — server-side commit-claim (delphi-architect; blocks §9 phase 4).** No natural
   unique key on `INV_OPEN_ORDER_INF`(part,FRS,renban,qty); proc returns nothing the caller
   checks. Confirm no reusable dedup key exists, then add an `INV_ORDER_COMMIT_CLAIM`
   (unique `simToken`) inserted in the same tx as the orders, so retries/failover/double-tabs are
   rejected at the DB (§4.1). Confirm the new table does not perturb the parallel Delphi writer.
3. **OPEN GAP — multi-site schema surgery (delphi-architect; blocks §9 phase 1).** No `site_id`
   column anywhere (grep = 0). Enumerate every `INV_*` table needing a site column and every START
   proc (15) + commit proc (3) needing re-authoring; widen matching `_HIST` tables in lockstep
   (F1 trigger hazard, `docs/triggers.sql`). `siteScopedQuery()` is delivery, not isolation (§7).
4. **OPEN GAP — `AD_GetSpecialDate` body + catalog + result shape + status domain
   (delphi-architect / ALC schema; blocks §9 phase 1).** Overtime/X/holiday calendar (spec §3.1)
   and added-leadtime (spec §3.4) depend on it. Hold the cross-DB-vs-TVP choice until the
   result-set shape and status-code domain are confirmed (§2). Do NOT re-derive O/X/holiday rules.
5. **OPEN GAP — template conditional-format rules + thresholds UNREAD (extraction; blocks §9
   phase 5 — must close BEFORE the accessibility/signal layer).** xlrd 2.0.2 cannot read BIFF CF
   rules; no LibreOffice; openpyxl can't read .xls (artifacts §4). The §5 table emits **only**
   Pascal-set signals — any template-baked CF threshold the users rely on daily is missing.
   CLOSE PATH: re-save `OrderSimulation.xls` → `.xlsx`, then read
   `worksheet.conditional_formatting` + palette via openpyxl. Gate phase 5 on **signal parity**,
   not just number parity.
6. **OPEN GAP — color palette RGBs UNCONFIRMED (extraction; blocks §9 phase 5).** §5 RGBs are
   standard Excel-2000 placeholders; the workbook may override the palette. Re-derive from
   `Workbook.Colors` / palette via the same .xlsx re-save as #5.
7. **OPEN GAP — in-cell template formulas UNREAD (extraction + parity; blocks §9 phase 2
   sign-off).** Share %, end-balance recurrence reproduced from Pascal spec (§3.2/§3.5), but the
   template's own formulas were unreadable (artifacts §4 gap 3). The §9.2 parity diff vs a live
   legacy Start is the catch-net — it MUST include the **hazard-7 calendar-offset vs fill-position
   index reconciliation** as a named test (§2).
8. **~~OPEN GAP~~ RESOLVED — `research-best-practice.md` now exists (process; no longer blocks §9
   phase 5).** The doc is on disk and fact-checked (every retained recommendation backed by a live
   URL). All presentation choices (exception-first, frozen header, sticky columns, non-color-only
   signaling, grid performance) are now **cited** (§1.2/§5: grid-ux NN/g + Smashing + Perspective
   Table docs; WCAG 1.4.1 + Section 508). One recorded deviation: research §6.7's CVD-safe palette
   swap is deliberately NOT adopted, to keep colors bit-for-bit faithful (§5) — logged as a
   post-cutover, sign-off-gated improvement. Note the research doc itself flags the Perspective
   virtualization scroll/query "feel" as community-forum-only/unverified; §1.2 does not rely on it.
   (Extraction gaps #5/#6 — the .xls CF rules + palette RGBs — remain open and still gate phase 5's
   *signal* parity; those are about the legacy source, not best practice.)
9. **Silent caps** (spec §8 h3): legacy silently truncates >200 part rows / `fDates[0..200]`.
   Faithful (cap at 200) but should **surface a warning** on truncation (reviewer-judgment;
   does not change values). Confirm 200 is acceptable per site (delphi-architect).
10. **Part-type list source** (artifacts §1 file/valves/tires/wheels vs spec §1
    TIRE/WHEEL/VALVE/FILM): confirm the canonical per-site list and where it lives (config row).

## 9. Phased build sequence (parallel-run → cutover)

1. **Schema add (DB phase):** resolve OPEN GAPS #3 (multi-site columns + `_HIST` widening) and
   #4 (calendar) FIRST. Then `sites` table + `site_id` cols (D1) and `SIM_OrderSimulation`
   authored against current SQL Server.
2. **Parity proc:** build `SIM_OrderSimulation`; **diff its grid vs a live legacy Start** across
   **O / X / holiday calendars and singleton-vs-shared-size cases**, including the named hazard-7
   index-reconciliation test (§2/#7). Gate sign-off on **byte-level number parity AND signal
   parity** (the latter after CF/palette gaps #5/#6 closed).
3. **Read-only screen:** `SelectOrderBar` + `PhasedGrid` bound to the sim NQ via
   `siteScopedQuery()`. Ship Simulate-only first (zero write risk during parallel run).
4. **Commit path:** wire the commit procs **inside the serializing tx + commit-claim** (§4.1) —
   NOT bare. **Blocked on OPEN GAPS #1 and #2.** **Parallel run:** Delphi and Ignition both create
   orders against one DB; validate (a) the both-Q-and-R-non-empty skip gate, (b) the
   `SELECT_PartsStockRenban` `@PartNum` binding, and (c) that concurrent same-part commits
   **serialize** (no duplicate FRS, no double-create on replay). The bare proc does NOT mediate —
   the serialization we add does.
5. **Accessibility/exception-first layer:** enums → glyph/text/ARIA + legend (§5), built to the
   now-cited best practice (WCAG 1.4.1 / Section 508 + grid-ux; research-doc gap #8 RESOLVED).
   Still gated on CF/palette **source-extraction** gaps #5/#6 (signal parity needs the .xls CF
   rules + palette RGBs), **not** on best-practice availability. Verify contrast ≥3:1 (status
   fill/icon) / ≥4.5:1 (in-cell text) per research §2.
6. **Cutover:** retire Delphi Order form per site once parity + commit confirmed; keep the qty
   trigger and all wrapped procs in place (no phase-1 logic port).

8.3-design / 8.1.52-runnable notes (greppable):
- `# IG83-TODO:` use Perspective Table native column-freeze / sticky if 8.3 adds richer grid
  styling; on 8.1.52 use the available column config + CSS class styling.
- `# IG81-COMPAT:` confirm `system.db.createSProcCall` multi-result-set handling on 8.1.52 for
  `SIM_OrderSimulation`'s header+grid result sets; if limited, return the header as a second NQ.
- Guard any 8.3-only Perspective component with `system.util.getVersion()`.
