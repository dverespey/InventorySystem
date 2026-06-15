# Order — Ignition Perspective Target Design, OPTION B ("Best-Practice-Forward")

Target design for the Order ("What to Order") rebuild. Companion to the legacy
behavioral spec (`legacy-order-spec.md`) and the template-artifact extraction
(`source-artifacts.md`). Option B differs from a straight 1:1 port (Option A) by
**proposing MRP / time-phased best-practice changes to the simulation calc** — but
*only* where a cited, verified reason exists.

> ## Calc-change governance (read before trusting any calc change in §7)
> `research-best-practice.md` now EXISTS (verified 2026-06-13) and supplies cited MRP /
> safety-stock / lead-time sources. The §7 candidates below are now **resolved against it**:
> each carries a real citation and a status of **KEEP-PROPOSED (cited, NEEDS DAVID SIGN-OFF)**,
> **REVISE**, or **WONTFIX**. The research doc itself frames every math item as a *proposal*
> requiring sign-off and explicitly states it is **not asserting the legacy stored-proc math is
> wrong** (research §5, §7).
>
> **The hard rule is unchanged: NO calc change ships without David's explicit sign-off.** A
> citation makes a candidate *eligible* for sign-off; it does not approve it. Until David signs
> off a given row, **Option B's calc for that row == Option A's calc (the legacy calc)**. The
> presentation/data-layer/loop/color/multi-site design below is independent of §7 and stands on
> its own (grounded in `legacy-order-spec.md` + cited best practice).

---

## 1. Scope and decisions honored

- **Delivery:** interactive Perspective grid only (no Excel). Full loop:
  simulate → review/adjust → COMMIT (wraps the legacy Order-path procs). [LOCKED]
- **Multi-site:** ONE view set, parameterized — not forked views. Per-site differences =
  config/params. **Phase-1 reality (parallel run):** the legacy SQL Server DB has NO `site_id`
  column and the wrapped procs take NO `site_id` param (verified — see §6); a deployed instance
  represents exactly ONE site and isolation is achieved by **which DB/connection the instance is
  pointed at**, not by row filtering. True session-derived multi-tenant scoping (`site_id` as a
  query param) is **deferred to the DB-modernization phase** when the column exists
  (decisions.md:34-36). The `siteScopedQuery()` helper is the eventual mechanism, **not yet
  implemented and unproven (pending spike Check B)** — see §6. [LOCKED intent, decisions.md D1]
- **DB:** wrap-the-proc. Named Queries (named to mirror schema/procs) +
  `system.db.createSProcCall`. Do not bypass procs/triggers in phase 1. [LOCKED]
- **Version split:** design for 8.3, runnable on 8.1.52. Guard 8.3-only paths; greppable
  `# IG83-TODO` / `# IG81-COMPAT` / `# IG83-ONLY` markers. [LOCKED, ignition-version-strategy.md]
- **Solo-dev maintainability:** hard constraint. Prefer simplest Ignition-native option.

Citations below use `spec §N` (legacy-order-spec.md), `schema:NNN`
(`DB Schema/Create Inventory.sql`), `artifacts §N` (source-artifacts.md), `D1`
(decisions.md).

---

## 2. Perspective view structure

One Perspective page, three regions in a root `flex` (column). In **phase 1** the page is
scoped to the single site the instance represents (its DB connection — §6); the eventual
session-derived `siteId` carried on the page model is wired but inert until `site_id` exists
in the DB (DB-mod phase). Site is never bound from a client param.

```
order-page  (Perspective page, route /order)
├─ view: Order/SelectOrder            (the Select-Order control bar — replaces Order.dfm)
├─ view: Order/PhasedGrid             (the time-phased simulation table — replaces the .xls)
└─ view: Order/CommitBar              (review summary + COMMIT action)
```

### 2.1 `Order/SelectOrder` — the control bar (replaces `Order.dfm`)
Maps 1:1 to the legacy dialog (spec §1, artifacts §1):

| Control | Component | Source / behavior |
|---|---|---|
| **Today** | read-only label | `now()` server-side; the order date. Display only (spec §1). |
| **Line** | Dropdown (`ia.input.dropdown`) | options from `Order/dd_lines` NQ (site-scoped). Blank allowed. |
| **Part** | Dropdown | options TIRE/WHEEL/VALVE/FILM. Default list is per-site config, not hardcoded (artifacts §1 lists "file, valves, tires, wheels"; treat as config). |
| **Sort By** | Dropdown | **hidden unless Line is blank** (spec §1, `Order.pas:1659`; artifacts §1). Bind `meta.visible` to `{Line} = ''`. |
| **Start** | Button | `onActionPerformed` → runs the simulate path (§4), writes result to page `view.custom.sim` model. Does NOT write DB (spec §1). |
| **Order** (COMMIT) | Button on CommitBar | **enabled only after a successful Start** (artifacts §1). Binds `enabled` to `{view.custom.simReady}`. |
| **Exit** | Button | navigate away / clear model. |

The Select-Order parameters become a single `selection` object on the page
(`{line, partType, sortBy, today, siteId}`) passed down to the grid view as a param.

### 2.2 `Order/PhasedGrid` — the time-phased simulation (replaces `OrderSimulation.xls`)

The grid is the heart of the screen. **Default target: a stock Perspective Table**
(`ia.display.table`), which in 8.x supports per-cell styling via column `render`/`style`
config and per-cell data objects (a cell value can be an object carrying `{value, style,
icon, ariaLabel}`). The Flex-Repeater-of-row-sub-views is a **fallback, not the default** —
it is a heavy, hand-maintained component for a ~23-day × up-to-200-row grid (spec §1 cap)
and a render-perf risk on the Intel-Mac 8.1.52 dev box at ~200×26 cells. **OPEN ITEM —
bounce to Ignition-spike (see §5):** prototype the per-cell dual-channel encoding (style +
sibling icon/aria) in a stock Table first; only fall to the Flex Repeater if the Table
genuinely cannot carry both channels. Do not commit to the Flex Repeater until the spike
proves the Table can't.

Layout mirrors the legacy column model (spec §1 layout constants, artifacts §2 header rows):

- **Left fixed columns (per part row):** Size (group header), Brand/Supplier (C),
  Part No (D), Sum-of-Order Qty + Actual%/Plan% (E/F/G), Order Point: Daily Usage (H) /
  Safety Days (I) / Safety Stock J=H*I, Inventory: Total (J/IN_QTY) / `<SITE>` (K) /
  `<DEST>` (L) / In-Transit (M) / Open (N), 1Lot Qty (O), Lead Time (P), **Qty (Q,
  editable)**, **Lot (R, editable)**.
- **Phased day columns (T … T+FillDays-1):** one column per *production* day, headed by
  date (row 5) + weekday (row 6). Default FillDays = 23 (spec §1; INI `[INIT] FillDays`,
  **max 50** per artifacts §1). FillDays becomes a per-site config value (§6).
- **Three summary columns** after the phased block: Total Inv / In-Transit / Added
  Leadtime (spec §1, `DateWeekCol+FillDays+1..+3`).

Editable cells: **Qty (Q)** and **Lot (R)** per part row — the cells the user adjusts
during review (spec §6, read back at `Order.pas:654-655`). All other cells are read-only
projections. End-balance / lead-time-zone / order-by-point are computed and colored (§5).

**Cited grid-UX defaults (research §4/§6, not reviewer judgment):** freeze the leftmost
human-readable identifier columns (Size/Part No) and the date header row; time-phased day
columns scroll horizontally; right-align quantities, left-align text
(https://www.nngroup.com/articles/data-tables/,
https://www.pencilandpaper.io/articles/ux-pattern-analysis-enterprise-data-tables — Perspective
supports frozen rows/columns + per-cell conditional styling natively per IA docs,
https://www.docs.inductiveautomation.com/docs/8.1/appendix/components/perspective-components/perspective-display-palette/perspective-table).
**Exception-first**: default-sort/open on the parts that need ordering today, with a visible
"filters active" indicator and a clear "nothing to order today" empty state (research §4/§6.9,
https://www.smashingmagazine.com/2025/09/ux-strategies-real-time-dashboards/). For the
~200×26 grid, enable Perspective's native `virtualized` + `pager` with fixed row height —
client-side rendering is fine under ~1,000 rows, which keeps the stock Table viable and further
argues against the Flex-Repeater fallback (research §4/§6.11; the Infragistics benchmark numbers
are vendor self-reported and the Perspective scroll "feel" is forum-only — both NOT relied on,
research §4 caveats / §8).

`K6`/`L6` header labels (source warehouse / destination plant) come from **site config**,
not the view (artifacts §3 — the ONLY per-site template difference, see §6).

### 2.3 `Order/CommitBar`
- Summary chips: # part rows with non-zero Qty, total order qty, # below-safety rows.
- **Order (COMMIT)** button → §3 commit action. Per cited best practice, this is the **one
  high-emphasis primary CTA** for the view: action-verb copy ("Commit Orders for Today"),
  strong contrast, ≥48px; secondary actions (recalculate/adjust horizon) de-emphasized
  (research §4/§6.12, https://subux.pro/guides/article/button-hierarchy-primary-secondary-tertiary).
- A per-run **client-side commit token** (uuid generated at Start) shown read-only, sent
  with the commit for idempotency tracing (§4.3).

---

## 3. Data layer — Named Queries / SProc calls (wrap-the-proc)

Named Queries are organized to **mirror the schema and the procs they wrap**. Proposed
NQ folder/namespace: `Order/`, one NQ per legacy proc, named after the proc. In phase 1
each NQ targets the instance's single-site DB connection directly — there is no `site_id`
param to add to the wrapped procs (§6), so reads run against the proc signatures verbatim;
the `siteScopedQuery()` indirection is deferred (§6). Writes use `system.db.createSProcCall`
inside the gateway commit script (§3.2) so the existing transaction/trigger semantics
are preserved.

### 3.1 START / simulate path (read-only) — NQ ⇄ proc map (spec §2 START table)

> **BUILD-GATING DEPENDENCY:** `AD_GetSpecialDate` (first row) lives in the **ALC/`TireOrder`
> database**, is absent from `Create Inventory.sql`, and its body is UNVERIFIED. The entire
> simulate path is **blocked** until that proc is located in the ALC schema and its body read —
> `build_production_calendar` (§4) cannot be written without it, and there is no production-day
> axis (and therefore no grid) without the calendar. This is a hard prerequisite, not a footnote
> (see §8.3). Confirmed only as a SELECT-shaped read so far (spec §2); param/return shape and
> per-site keying are unknown.

| Named Query | Wraps proc | Params | Notes |
|---|---|---|---|
| `Order/AD_GetSpecialDate` | `AD_GetSpecialDate` (**ALC/`TireOrder` DB**) | `@BeginDate,@EndDate,@LineName` | **Separate JDBC connection** (ALC catalog). **Proc body UNVERIFIED + cross-DB** (spec §8 hazard 1). **BUILD-GATING — locate in ALC schema before any simulate-path work.** |
| `Order/SELECT_PartsStockInfoOrder` | `SELECT_PartsStockInfoOrder;1` | `@LineName,@PartType,@SortType` | Driving cursor; ordered by size. `schema:7382`. |
| `Order/SELECT_SupplierInfo` | `SELECT_SupplierInfo;1` | `@SupCode` | `schema:7978`. |
| `Order/SELECT_SizeInfo` | `SELECT_SizeInfo;1` | `@SizeCode` | Daily Usage / Safety Days. `schema:7869`. (Legacy double-open is dead; call once — spec §8 hazard 4.) |
| `Order/SELECT_OrderHistory` | `SELECT_OrderHistory;1` | `@PartNumber` | order-share %. `schema:6757`. |
| `Order/SELECT_ForecastDetailTWPN` | `SELECT_ForecastDetailTWPN;1` | `@PartNumber,@EffMonth,@TireWheel,@IncludeZero` | tire/wheel ratio. `schema:6228`. |
| `Order/SELECT_UsageDay` | `SELECT_UsageDay;1` | `@Date,@PartNo` | usage-vs-forecast compare. `schema:8088`. |
| `Order/SELECT_FirstProductionDay` | `SELECT_FirstProductionDay;1` | `@ProdYear` | week offset. `schema:5982`. |
| `Order/SELECT_ForecastPartNumberWeek` | `SELECT_ForecastPartNumberWeek;1` | `@WeekNo,@DayNo,@PartNo` | forecast by day. `schema:6309`. |
| `Order/SELECT_OrderAtASSEMBLER` | `SELECT_OrderAtASSEMBLER;1` | `@PartNumber` | → K. `schema:6643`. |
| `Order/SELECT_OrderAtPLANT` | `SELECT_OrderAtPLANT;1` | `@PartNumber` | → L. `schema:6700`. |
| `Order/SELECT_OrderInTransit` | `SELECT_OrderInTransit;1` | `@PartNumber,@FirstFRS` | → M total. `schema:6816`. |
| `Order/SELECT_OrderInTransitList` | `SELECT_OrderInTransitList;1` | `@PartNumber,@FirstFRS` | per-FRS-date rows for phasing. `schema:6850`. |
| `Order/SELECT_OrderOpenOrder` | `SELECT_OrderOpenOrder;1` | `@PartNumber` | → N total. `schema:6955`. |
| `Order/SELECT_OrderOpenOrderList` | `SELECT_OrderOpenOrderList;1` | `@PartNumber,@FirstFRS` | per-FRS-date rows. `schema:6985`. |

### 3.2 ORDER / commit path (writes) — SProc calls (spec §2 ORDER table, §6)

Run inside a **gateway-side script** that opens an explicit transaction
(`system.db.beginTransaction` / `commitTransaction` / `rollbackTransaction`) so it matches
the legacy per-insert `BeginTrans/CommitTrans` with rollback (spec §2, `Order.pas:686/780`).

| Call | Wraps proc | Params | Effect |
|---|---|---|---|
| `Order/SELECT_PartsStockRenban` | `SELECT_PartsStockRenban;1` | `@PartNum` | read current renban counter. `schema:7506`. |
| `Order/UPDATE_PartsStockRenban` | `UPDATE_PartsStockRenban;1` | `@PartNum,@RenbanCount` | bump counter (wrap >999→1). `schema:9018`. |
| `Order/INSERT_OpenOrder` | `INSERT_OpenOrder;1` | `@SupCode,@PartNum,@KanbanNum,@FRSNum,@RenbanNum,@Qty` | **the order insert.** No OUTPUT/return read. FRS year-roll + sequence suffix computed server-side. `schema:3236`. |

**Commit per-row logic must reproduce spec §6 exactly** (do not reimplement in phase 1):
- Lot-size order (`BIT_LOT_SIZE_ORDERS=TRUE`): one `INSERT_OpenOrder` with `@Qty=Qty`.
- Non-lot-size: loop `j:=1..Lot`, one insert per lot with `@Qty=IN_1LOTQTY` (the 1-lot
  qty, NOT the typed Qty — spec §6, easy to get wrong).
- Renban: if part not in a renban group, read+form+bump counter; if in a group,
  `@RenbanNum=''` and defer numbering. The proc recomputes the FRS trailing-2 digits via
  `max+1`, so the Pascal-supplied suffix is dead (spec §6) — **idempotency lever, §4.3**.
- **Trigger fact:** `INSERT_RecConfStatPartsStockMstQTY` copies to `_HIST` always but only
  bumps `INV_PARTS_STOCK_MST.IN_QTY` when shipping-status ≠ '' — a new order has empty
  status, so **stock is NOT bumped at creation** (spec §6, §8 hazard 8). The Ignition layer
  must NOT add qty on order — the trigger owns inventory coupling. [ASSUMPTION dep:
  trigger behavior verified in spec §6 from `docs/triggers.sql:214-227`; re-confirm `_HIST`
  has `site_id` once D1 lands — `_HIST` `SELECT *` breaks if base gains a column the hist
  table lacks (spike-db.sh:59 hazard F1).]

---

## 4. WHERE the simulation calc runs — DB proc vs gateway Jython

**Decision: the per-day phasing / lead-time / end-balance projection runs in GATEWAY
JYTHON, not in a new DB proc.** The DB procs remain pure data *fetchers* (they already
are — every START proc is a `SELECT`; spec §2). The composition logic
(calendar→production-day mapping, forecast accumulation per size group, in-transit/open
bucketing, `DoLeadTime`, `DoFormulas` end-balance roll) is currently in **Pascal**, not in
SQL (spec §3). So:

**Justification (solo-dev maintainability + wrap-the-proc fidelity):**
1. The legacy calc lives in `Order.pas`, *not* in a stored proc — there is no proc to wrap
   for it. Pushing it into a new T-SQL proc would be a *reimplementation*, which phase-1
   "wrap before reimplement" forbids. Porting Pascal→Jython is a closer, reviewable
   translation than Pascal→T-SQL.
2. The calc is read-only and per-run; it has no transactional/locking need that argues for
   the DB. Keeping it in the gateway keeps the index-space reconciliation
   (`fDates` indexed by calendar offset vs forecast arrays indexed by fill position — spec
   §8 hazard 7) in one readable place, with unit-testable Jython functions.
3. The **commit** path stays in procs (it must — `INSERT_OpenOrder` owns FRS sequencing and
   the trigger owns inventory). Only the *projection* is in Jython.

**Engine shape:** `Order/sim` gateway script module with pure functions:
`build_production_calendar(special_dates, fillDays)` → ordered list of production days with
O/X/holiday flags (spec §3.1); `phase_quantities(part, days, lists)` → in-transit/open
buckets (spec §3.3, preserve the `fDates[i]<>0` scan exactly); `lead_time_zone(part, days,
overtimes)` → leadtime-zone + order-by index + added-leadtime (spec §3.4); `end_balance(part,
days)` → projected balances + below-safety flags (spec §3.5). Start calls these, assembles a
JSON model, returns it to the page (`view.custom.sim`). **No DB write on Start.**

`# IG81-COMPAT:` runs on 8.1.52 as a project/gateway script invoked from the Start button
via `system.util.getGlobals`-free message handler or direct script call.
`# IG83-TODO:` evaluate moving `Order/sim` into an 8.3 script library + verify
`system.db.createSProcCall` OUT-param shape unchanged.

**Parity test (cutover gate):** the Jython phasing functions are the **highest calc-fidelity
risk surface** in the rebuild (a Pascal→Jython port of the time-phased projection). Before
cutover, each function must be diffed against a live legacy run — capture a real Order session's
inputs and its resulting grid from the Delphi app and assert the Jython output matches cell-for-
cell, with special attention to the index-space hazard (hazard 7). No cutover without a green
parity run on a representative line/part set.

### 4.3 Idempotency under parallel run (simulate→review→commit)
- **Start is idempotent by construction** — read-only, no writes (spec §1). Re-running
  Start just rebuilds the model.
- **Commit idempotency rides on `INSERT_OpenOrder`'s server-side `max(FRS)+1`** (spec §6,
  `schema:3266-3314`): re-committing the same run appends *new* FRS sequence numbers rather
  than duplicating, and the Delphi app committing the same line in parallel does the same —
  the proc serializes the sequence. **This means the proc does NOT dedup identical orders;
  it only guarantees unique FRS numbers.** A double-click or double-commit WILL create
  duplicate orders. Mitigation in the Ignition layer — **scoped to what each actually covers:**
  - Disable the Order button on click; re-enable only on commit result. (Covers the common
    double-click within one Ignition session.)
  - Stamp each committed run with the per-run **commit token** (§2.3) and refuse to commit a
    token already committed (gateway-side guard keyed on token in a small `order_commit_log`
    helper table). **Scope is explicit: this de-dups only Ignition's OWN resubmits of the same
    run** (same token). It does **NOT** mediate the cross-app race — the Delphi app never writes
    the token table, so a Delphi commit of the same line is invisible to it and still produces a
    distinct order. The token table is therefore an Ignition-internal idempotency aid, not a
    parallel-run dedup. [ASSUMPTION: helper table is new infra, additive; confirm acceptable vs.
    phase-1 "no schema change beyond `site_id`." If rejected, fall back to button-disable only.]
  - **Cross-app duplicates and the renban read-bump race are NOT solved by the token.** The
    renban path (`SELECT_PartsStockRenban` → `UPDATE_PartsStockRenban`, no lock — §3.2) can
    interleave between the Delphi app and Ignition and produce a skipped/duplicated counter.
    This is a **source-truth + DB-locking decision, not an Ignition design choice**:
    **OPEN GAP — bounce to delphi-architect:** either (a) wrap the read-bump in an `UPDLOCK`/
    serializable transaction inside the proc (a proc edit = reimplementation, out of phase-1
    scope), or (b) accept the existing legacy race (the Delphi app has run this way for years —
    confirm it is in fact tolerated). Until decided, Ignition reproduces the legacy
    read-then-bump verbatim and inherits the legacy race; it does not claim to fix it.
  - During parallel run, the Delphi app and Ignition both hit the same procs/triggers, which
    mediate row *shape* and inventory coupling — **but the procs do not serialize cross-app
    intent** (no dedup; renban unlocked). Do not write open-order rows from Ignition by any path
    other than `INSERT_OpenOrder`.

---

## 5. Accessible color / signal model (never color-only)

The legacy sheet encodes meaning in **two Excel channels** — interior (zone) and font
(qty source) — with the exact palette in spec §4. **Critical gap (artifacts §4):** the
conditional-format rules and any template-baked thresholds are in `OrderSimulation.xls`
and were **NOT extractable** (xlrd can't read CF/formatting). So the *Pascal-set* colors
(spec §4) are known; any *template-set* CF rules are **UNKNOWN — do not guess** (artifacts
§4.1). Below covers only the Pascal-known signals; the template CF gap is flagged in §8.

Each color signal gets a **redundant non-color cue** (icon + text/aria-label). This is **cited
best practice, not a reviewer-judgment default**: WCAG 1.4.1 Use of Color (Level A) governs —
"color is not used as the only visual means of conveying information"
(research §2, https://www.w3.org/WAI/WCAG21/Understanding/use-of-color.html); the Section 508
redundant-coding heuristic (status in color **and** a word, plus warning icon)
(https://www.section508.gov/create/making-color-usage-accessible/); and the acid test "remove all
color and the grid is still operable" (research §2, §6.6). Each status cell carries a unique
{color + icon-shape + text-token} triple readable in grayscale (research §6.6).

| Legacy signal (spec §4) | Meaning | Perspective encoding (color + non-color) |
|---|---|---|
| Interior 36 (pale yellow) on T..T+lead-1+added | lead-time zone | left cell border tint + cell tag/badge "LT" + `aria-label="within lead time"` |
| Interior 40 (cream) order-by column | place order this day | bold left border + a flag/pin icon + header tag "ORDER BY" + aria-label |
| Interior 3 (red) overtime column | overtime production day | red top-border + "OT" superscript on the date header |
| Interior 4 (green) non-production column | 'X' non-production day | hatch pattern + "X" on the date header |
| Font 23 (dark blue) in-transit qty | qty in transit | prefix glyph (arrow) + text tag "in transit" + tooltip (not blue-only) |
| Font 10 (dark green) open-order qty | open (unshipped) order | prefix glyph (open circle) + text tag "open" + tooltip |
| FormatCond font 3 end-balance < safety | below safety stock | warning icon + bold + text "below safety" + tooltip "below safety stock (J=usage×days)"; row also tagged `belowSafety:true` for the CommitBar count |
| FormatCond font 3 J<0 usage-vs-forecast | over-produced | up-arrow icon + text tag "over" |

> **Glyph note (do not ship literal characters):** the marker/glyph names above are
> *descriptions*, not the actual rendered characters — earlier drafts of this table contained
> replacement/mojibake codepoints. Pick concrete icon assets from the Perspective icon library
> (Material icons via the `icon` component / cell render config) and reference them by `library/
> name`; never paste a literal glyph that may not round-trip through the file encoding.

Implementation: the **default attempt is per-cell encoding in a stock Perspective Table** —
each day-cell value is a model object `{value, signal}` and the column's render/style config
maps `signal` to both a `style` and a sibling icon + aria/text string, so color and the
non-color cue bind to the same field. **OPEN ITEM (Ignition-spike):** confirm a stock Table
can carry per-cell style + icon + aria for ~200×26 cells with acceptable render performance on
8.1.52; the claim that a Table "can only style by row" is **unproven and likely overstated**,
so prove or disprove it in a spike before falling back. **Fallback only if the spike fails:** a
Flex Repeater of row sub-views (heavier, hand-maintained, perf-risk at 200×26 — see §2.2).
**Color is never the sole carrier** in either case. Use a CVD-safe palette (cited, research §2/§6.7):
replace raw red/amber/green with **Okabe-Ito** (`#009E73` healthy / `#E69F00` warning /
`#D55E00` shortage; hex confirmed by two independent sources —
https://conceptviz.app/blog/okabe-ito-palette-hex-codes-complete-reference,
https://mk.bcgsc.ca/colorblind/palettes.mhtml) or **IBM Design Language**
(`#648fff`/`#ffb000`/`#dc267f`). Red↔green opposition must never be the sole differentiator
(~8% of men have CVD; research §2) — so the OT vs X pair also differs by border vs hatch + letter.
Meet contrast minimums: status fill/border and icon ≥ 3:1 (WCAG 1.4.11), in-cell text ≥ 4.5:1
(WCAG 1.4.3); validate final hex pairs with a CVD simulator (research §2/§6.7,
https://www.smashingmagazine.com/2024/02/accessibility-standards-empower-better-chart-visual-design/).

`# IG83-TODO:` re-verify `style` binding + `meta` accessibility props (aria) on 8.3 table/flex.

---

## 6. D1 multi-site — one view set, config not forks (phase-1 = one site per instance)

The only real per-site divergence in source is **cosmetic** (artifacts §3 verdict):
templates differ ONLY in **K6** (source warehouse label) and **L6** (destination plant
label). Everything else (columns, layout, formulas) is identical; the `*HERO/*CAMEX/*WWW/
*MAS/*WQS` files are just copies deployed under the plain name per site. So:

- **One `Order/*` view set, parameterized.** No forked views. This is the LOCKED structural
  decision and holds in both phases.

**Site scoping has TWO phases — do not conflate them (this was a prior error in this doc):**

- **Phase 1 (parallel run) = one site per running instance.** The legacy SQL Server DB has
  **no `site_id` column** and the wrapped Order procs take **no `site_id` parameter** — verified:
  `SELECT_PartsStockInfoOrder(@LineName,@PartType,@SortType)`,
  `INSERT_OpenOrder(@SupCode,@PartNum,@KanbanNum,@FRSNum,@RenbanNum,@Qty)`, and
  `SELECT_OrderInTransitList(@PartNumber,@FirstFRS)` have no site param, and a grep for
  `site_id` across the schema returns zero hits. Per decisions.md:34-36 the legacy DB is
  **untouched during parallel run** and "the new app simply filters to the one site it
  represents." Therefore isolation in phase 1 is achieved by **DB/connection selection** — the
  deployed instance is pointed at the one site's database and every NQ runs against the proc
  signatures verbatim. **There is nothing to inject `site_id` into; injecting an extra
  `@site_id` would error or be silently ignored.** `siteScopedQuery()` is effectively a no-op
  pass-through in this phase.
- **Phase 2 (DB-modernization) = true session-derived multi-tenant scoping.** Only after the
  `site_id` FKs land (decisions.md D1, deferred to the Postgres/DB-mod phase) does a real
  `siteId` exist to scope on. At that point `siteScopedQuery(nqPath, params, session)` reads
  `siteId` from the Perspective session (`session.custom.siteId`, set at login from the
  user→site mapping) and constrains the query server-side; the client cannot spoof site.
- **`siteScopedQuery()` is UNPROVEN, pending spike Check B.** It is described in
  `ignition-version-strategy.md` and `README.md` only as a *planned* spike; grep finds no
  implementation in `scripts/`. **Do NOT cite it as established/LOCKED precedent.** Check B
  must run against the actual parallel-run (no-`site_id`) DB and confirm the no-op-then-scope
  two-phase shape. [decisions.md D1 intent LOCKED; the mechanism is not.]
  **OPEN GAP — bounce to delphi-architect/Ignition-spike:** confirm (a) the parallel-run instance
  is single-site-per-deployment as decisions.md implies, and (b) Check B's helper degrades to a
  pass-through when no `site_id` column exists.
- **Per-site config (a `site_config` lookup, site-scoped):**
  - `K6` source-warehouse label, `L6` destination-plant label (artifacts §3).
  - `FillDays` (legacy single INI value 23, max 50 — artifacts §1; per-site now).
  - `ForecastUsageCompare` (legacy 7), `UseFirstProductionDay` flag (spec §3.2/§7).
  - `BIT_LOT_SIZE_ORDERS` behavior toggle (spec §6).
  - Part-type list (TIRE/WHEEL/VALVE/FILM default; per-site config — spec §7/artifacts §1).
  - **ALC/`TireOrder` calendar connection** per site (the overtime/holiday source —
    spec §7, §8 hazard 1). [ASSUMPTION dep: how the ALC calendar is keyed per site is
    UNKNOWN until `AD_GetSpecialDate` body is located — extraction gap, §8.]
- **D1 `_HIST` hazard (F1):** adding `site_id` to base tables breaks
  `INSERT INTO _HIST SELECT * FROM inserted/deleted` triggers unless the `_HIST` table also
  gets `site_id` (spike-db.sh:59). The Order path fires
  `INSERT_RecConfStatPartsStockMstQTY` which does `SELECT * from inserted` into
  `INV_OPEN_ORDER_INF_HIST` (spec §6) — so `INV_OPEN_ORDER_INF_HIST` **must** gain `site_id`
  in lockstep. Flag for the DB-modernization phase.

---

## 7. PROPOSED CALC CHANGES (Option B) — RESOLVED AGAINST research-best-practice.md

> **Every entry here is still a CANDIDATE.** `research-best-practice.md` now exists and each row
> below carries a real citation, so the rows are resolved to one of:
> **KEEP-PROPOSED (cited, NEEDS DAVID SIGN-OFF)** — research supports the change and it is eligible
> for sign-off; **REVISE** — research supports a *different* shape than originally drafted, restated;
> **WONTFIX** — research does NOT support the change as drafted (with reason).
> **A citation does NOT approve anything.** The hard rule stands: **no calc change ships without
> David's explicit sign-off**, and until a given row is signed off, Option B uses the legacy calc
> for that row (== Option A). The research doc itself frames all math as proposals, not as proof the
> legacy proc math is wrong (research §5 preamble, §7 preamble).
>
> **C1 remains tracked SEPARATELY from the MRP candidates (C2–C6).** C1 is a *defect fix* (a silent
> data-loss cap), not a calc-semantics change — it never depended on the research doc. C2–C6 are
> MRP-semantics changes; all six now have citations.

| # | Legacy behavior (cited) | Proposed change (candidate) | Citation | Status |
|---|---|---|---|---|
| C1 | Silent caps: ≤200 part rows (`fpartline[1..200]`), `fDates[0..200]`; >200 rows truncated with no guard (spec §8 hazard 3) | Remove the cap; page/virtualize the part set; surface a warning if a result exceeds a sane bound instead of silently dropping rows | **DEFECT FIX** (independent of MRP); large-grid handling corroborated by research §4/§6.11 (native Perspective `virtualized`+`pager`, fixed row height; client-side under ~1,000 rows) https://www.docs.inductiveautomation.com/docs/8.1/appendix/components/perspective-components/perspective-display-palette/perspective-table, https://www.telerik.com/blogs/best-practices-creating-user-friendly-data-grids | **NEEDS DAVID SIGN-OFF** (defect, not MRP-blocked) |
| C2 | Lead-time chosen by *today's* weekday column (`IN_LEADTIME_MONDAY..SATURDAY`), fallback `IN_LEADTIME` (spec §3.4, `Order.pas:426-459`) | Select lead time by the *order-by (release) day's* weekday and offset it against the working calendar so the release lands on a valid working day | research §3/§6.4 — lead-time offset is counted in working days against a calendar, per supplier/weekday; "single most error-prone mechanic" https://www.oliverwight-americas.com/glossary-terms/lead-time-offset/, https://docs.oracle.com/cd/E16582_01/doc.91/e15139/und_reqs_plng_concepts.htm | **KEEP-PROPOSED (cited, NEEDS DAVID SIGN-OFF)** |
| C3 | Added-leadtime: each overtime day inside the lead window pushes order-by out by 1, `break` on first miss (spec §3.4, `Order.pas:1576-1582`) | Model the lead-time offset as a true working-calendar count that skips non-production days and absorbs overtime/extra-shift days, replacing the `break`-on-first-miss loop | research §3/§6.2/§6.4 — offsets land only on valid working days; non-production days auto-skipped; shift-defined calendars are the mechanism for overtime/extra-shift days https://docs.oracle.com/en/applications/jd-edwards/supply-chain-manufacturing/9.2/eoash/understanding-shop-floor-calendar-setup.html, https://docs.oracle.com/cd/A60725_05/html/comnls/us/mrp/rpper.htm | **KEEP-PROPOSED (cited, NEEDS DAVID SIGN-OFF)** |
| C4 | End-balance = Beg + Receipts − Usage, flat per-day, below-`$J$` (usage×days fixed) → red (spec §3.5) | Express as running Projected Available Balance per working day vs. a **parameterized statistical safety stock** (`SS = Z×σ`, Z per item-group/service level; King combined formula where demand AND lead time vary, NOT the additive form) instead of fixed days×usage | research §1/§5/§6.3/§7(2,3,4) — PAB as running balance; `SS=Z×σ`; King combined formula avoids double-counting; Z per item-group https://web.mit.edu/2.810/www/files/readings/King_SafetyStock.pdf, https://docs.oracle.com/cd/E18727_01/doc.121/e15188/T478564T479029.htm | **KEEP-PROPOSED (cited, NEEDS DAVID SIGN-OFF)** |
| C5 | Order qty / share: `=E/(ΣE)` Excel formula across size group, 100% if singleton (spec §3.2) | Drive the order signal from explicit time-phased **net requirements** (gross req − scheduled receipts − projected available − in-transit per period), keeping `ROP = lead-time demand + SS` as fallback; the across-group share split is then a downstream allocation, not the trigger | research §1/§5/§6.1/§7(5) — standard MRP/TPOP netting and the net-requirements trigger; legacy day-by-day sim is already TPOP-shaped (validates, not contradicts) https://docs.infor.com/ln/2023.x/en-us/lnolh/wholh/whom000316.html, https://www.netstock.com/blog/reorder-point-formula/ | **REVISE → KEEP-PROPOSED (cited, NEEDS DAVID SIGN-OFF)** — restated: net-requirement netting is the *trigger*; the `=E/(ΣE)` share split is retained as allocation, not replaced wholesale |
| C6 | Forecast accumulation uses week/day breakdown table + optional first-production-day offset (spec §3.2) | Normalize forecast bucketing directly onto the production-calendar working days (one bucket per working day), and treat firm near-term (862/DELJIT) demand separately from far-horizon forecast (830/DELFOR) | research §1/§3/§5/§6.2/§6.13/§7(9) — columns = production-calendar working days; firm-vs-forecast EDI cascade drives near-term hard signals vs. softer far-horizon cues https://docs.oracle.com/cd/E26401_01/doc.122/e48795/T478564T478850.htm, https://www.orderful.com/blog/automotive-edi-guide | **KEEP-PROPOSED (cited, NEEDS DAVID SIGN-OFF)** |

**No WONTFIX rows:** research §1/§3/§5/§7 supports every C2–C6 candidate at least in principle, so
none is rejected outright; C5 is REVISED to match what the netting literature actually says (netting
is the trigger, not a replacement for the share allocation). Two research items have **no calc
candidate yet** and are noted for completeness, not added to the build: order-up-to/(R,s,S)
periodic-review overlay (research §5/§7(8)) and Croston for intermittent/lumpy parts (research
§5/§7(7)) — both depend on per-part classification data we do not yet have; capture as future
candidates if/when David raises them.

**Process rule (unchanged in force):** a citation only makes a row *eligible*. Each KEEP-PROPOSED /
REVISE row still requires David's explicit sign-off before it enters the build; before then Option B
runs the legacy calc for that row. Any future calc idea without a real citation may not merge.

**Parity-gate interaction:** even after sign-off, every adopted change must pass the §4 cell-for-cell
parity test against a live legacy run for the *unchanged* portions of the calc, so an approved C-row
does not silently perturb the parts of the projection it was not meant to touch (esp. the index-space
hazard, spec §8 hazard 7).

---

## 8. Open risks / unresolved extraction-gap dependencies

1. **`research-best-practice.md` EXISTS and §7 is resolved against it (no longer a blocker).**
   C1–C6 each carry a real citation with status KEEP-PROPOSED / REVISE / WONTFIX (none WONTFIX;
   C5 REVISED). The remaining gate is **governance, not citation**: no calc row enters the build
   without David's explicit sign-off, and until each row is signed off **Option B == Option A on
   the math for that row** (the legacy calc). Two cited research ideas (periodic-review overlay,
   Croston) have no candidate yet pending per-part classification data.
2. **Template conditional-format rules + thresholds are UNKNOWN** (artifacts §4.1/§4.2;
   xlrd can't read CF). The §5 color model covers only Pascal-set colors (spec §4); any
   template-baked threshold coloring is unmodeled. **Do not guess thresholds** — close by
   converting `OrderSimulation.xls` to `.xlsx` and reading CF via openpyxl (artifacts §4
   "next step").
3. **[BUILD-GATING] `AD_GetSpecialDate` body UNVERIFIED + cross-DB** (spec §8 hazard 1).
   The whole production-calendar mapping (§4 `build_production_calendar`) — and therefore the
   entire simulate path and the phased grid — **cannot be built** until this proc is located in
   the ALC/`TireOrder` schema and its body read. This is the **top build dependency for Option B**,
   not a footnote: no calendar → no production-day axis → no grid → no Start. Per-site keying of
   the calendar (§6) is also unknown until then (annotated OPEN GAP at §3.1 and §6).
4. **`INSERT_OpenOrder` does not dedup identical orders** — only guarantees unique FRS
   numbers (spec §6). Double-commit duplicates orders. Mitigations are **scope-limited** (§4.3):
   button-disable covers in-session double-click; the commit-token + `order_commit_log` helper
   table de-dups **only Ignition's own resubmits** (the Delphi app never writes the token, so it
   does not mediate cross-app dupes); the helper table is new additive infra needing sign-off.
   **OPEN GAP:** the cross-app double-commit and the unlocked renban read-bump race
   (`SELECT_PartsStockRenban`→`UPDATE_PartsStockRenban`) are NOT solved here — they need an
   `UPDLOCK`/serializable proc change (= reimplementation, out of phase-1) or an explicit
   decision to accept the legacy race. Bounce to delphi-architect.
5. **D1 `_HIST` column-add hazard (F1):** `INV_OPEN_ORDER_INF_HIST` must gain `site_id` in
   lockstep with the base table or the insert trigger's `SELECT *` breaks (§6,
   spike-db.sh:59).
6. **Excel locale read-back fragility** is eliminated by dropping Excel (spec §8 hazard 2) —
   a genuine Option-A/B shared win, not a calc change.
7. **Default part-type list** ("file, valves, tires, wheels" artifacts §1 vs TIRE/WHEEL/
   VALVE/FILM spec §1) differs in wording — treat as per-site config and confirm the
   canonical codes against `INV_PART_TYPE_MST`.
8. **Site-scoping mechanism (`siteScopedQuery`) is UNPROVEN** (§6). It is a planned spike
   (Check B), not implemented anywhere in `scripts/`. Phase-1 isolation relies on
   single-site-per-instance DB selection — confirm this against decisions.md with the architect;
   defer `site_id` param scoping to the DB-mod phase. OPEN GAP until Check B runs.
9. **Phased-grid renderer is undecided** (§2.2/§5). Default is a stock Perspective Table with
   per-cell dual-channel encoding; the "Table can't color per-cell" claim is unproven. Spike the
   Table first; the Flex-Repeater fallback is a solo-dev maintenance + render-perf risk at
   ~200×26 cells on 8.1.52. OPEN ITEM — Ignition-spike.
10. **Jython phasing parity** (§4): the Pascal→Jython port of the time-phased calc is the highest
    fidelity-risk surface. Requires a cell-for-cell parity test against a live legacy run as a
    cutover gate (esp. the index-space hazard 7). Not optional.
