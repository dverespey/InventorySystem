# Order Grid Redesign — Cited Best-Practice Research

Consolidated, fact-checked research for redesigning the automotive-parts ordering
tool as an Ignition Perspective time-phased "what to order today" grid.

**Fact-checking rules applied:** every retained recommendation is backed by a live
URL. Claims that could not be sourced to a primary/authoritative reference, or that
were only vendor-self-reported / community-forum / author-flagged judgment, are moved
to **Rejected / unsourced claims** or explicitly labeled. Accessibility rule enforced
throughout: **status is never signaled by color alone** (WCAG 1.4.1, Level A).

---

## 1. MRP / DRP time-phased planning-grid conventions (mrp-grid)

- **Use the standard time-phased record as the row vocabulary**: Gross Requirements →
  Scheduled Receipts → Projected Available Balance (PAB) → Net Requirements → Planned
  Order Receipts → Planned Order Releases. Definitions are APICS-aligned and consistent
  across vendor glossaries.
  https://docs.oracle.com/cd/A60725_05/html/comnls/us/mrp/gls.htm
  https://www.cleverence.com/articles/for-business/mrp-requirement-6732/
- **PAB is a running left-to-right balance**: `PAB = On-Hand + Scheduled Receipts −
  Gross Requirements − Safety Stock` for period 1, carried forward thereafter; a net
  requirement triggers when PAB would fall below safety stock/zero. Real ERP UIs (Oracle
  Planner Workbench horizontal plan, SAP MD04) compute a running available quantity
  per bucket.
  https://docs.oracle.com/cd/E18727_01/doc.121/e15188/T478564T479029.htm
  https://sapinsider.org/expert-insights/advanced-mrp-post-processing-with-new-md06-and-new-md04-part-2/
- **Lead-time offset (receipt vs. release)**: the planned order release is the receipt
  shifted earlier by the item lead time, counted in **working days against a calendar**,
  not naive calendar days; JD Edwards offsets manufactured items by working days,
  purchased items by calendar days.
  https://docs.oracle.com/cd/E16582_01/doc.91/e15139/und_reqs_plng_concepts.htm
- **Buckets + horizon**: columns are time buckets (daily/weekly/monthly); the planning
  horizon must be at least the cumulative lead time so releases land early enough.
  https://docs.oracle.com/cd/E16582_01/doc.91/e15139/und_reqs_plng_concepts.htm
- **Pegging (drill-down)**: link each supply line back to the demand(s) it satisfies;
  Oracle distinguishes single-level from full/graphical pegging, primarily for backward
  shortage tracing.
  https://docs.oracle.com/cd/E18727_01/doc.121/e15188/T478564T479029.htm
- **Exception / action messages drive the grid, not raw cells**: Order, Expedite /
  Reschedule-In, Defer / Reschedule-Out, Cancel, Increase/Decrease, Frozen/Firm,
  Past-due, Below-safety-stock. **Time fences** (freeze / planning / display) gate where
  messages generate.
  https://docs.oracle.com/cd/E16582_01/doc.91/e15139/und_reqs_plng_concepts.htm
  https://docs.infor.com/m3udi/16.x/en-us/m3beud/scplanhs/rps001.html
- **TPOP is the closest analog to this supplier-replenishment tool**: a demand-pull
  system that regulates time-phased supply, checking whether projected inventory drops
  below the reorder point within the order horizon.
  https://docs.infor.com/ln/10.7/en-us/lnolh/help/wh/onlinemanual/000316.html

*Sourcing caveat:* the verbatim APICS *Dictionary* definitions are paywalled; the row
set is corroborated by vendor glossaries and teaching material that agree with each
other. Specific heat-map color palettes are vendor-defined, not a standardized spec.

## 2. Reorder-point / days-of-supply visualization + accessible status signaling (reorder-viz)

- **WCAG 1.4.1 Use of Color (Level A) — the governing rule**: "Color is not used as the
  only visual means of conveying information." Pair color with icon and/or text. This is
  the minimum conformance tier, not optional polish.
  https://www.w3.org/WAI/WCAG21/Understanding/use-of-color.html
- **Section 508 redundant-coding heuristic**: mark status in color **and** a word
  (e.g., "(Correct)"), supplement red error fields with warning icons + text, give chart
  bars different patterns + text labels.
  https://www.section508.gov/create/making-color-usage-accessible/
- **The acid test**: if all color is removed, every user must still understand and
  operate the interface.
  https://wcag.dock.codes/documentation/wcag141/
- **Color-blind-safe palettes replace raw Red-Amber-Green** (red/green is the worst
  pairing; ~8% of men have CVD):
  - **Okabe-Ito** (hex confirmed by two independent sources): Orange `#E69F00`,
    Sky Blue `#56B4E9`, Bluish Green `#009E73`, Yellow `#F0E442`, Blue `#0072B2`,
    Vermillion `#D55E00`, Reddish Purple `#CC79A7`, Black `#000000`.
    https://conceptviz.app/blog/okabe-ito-palette-hex-codes-complete-reference
    https://mk.bcgsc.ca/colorblind/palettes.mhtml
  - **IBM Design Language**: Blue `#648fff`, Purple `#785ef0`, Magenta `#dc267f`,
    Orange `#fe6100`, Yellow `#ffb000`.
    https://www.audioeye.com/post/colorblind-friendly-palettes/
- **BI/dashboard RAG guidance**: prefer high-contrast hues (blue=positive, orange=warning)
  with text labels, or a single-hue sequential scale; always pair with icons/shapes/text.
  https://smart-frames.co.uk/2025/01/23/rethinking-rag-colours-in-business-intelligence-tools/
- **Contrast minimums**: non-text UI components/graphics ≥ 3:1 (WCAG 1.4.11); body text
  ≥ 4.5:1 (1.4.3). Status fill/border and icon must each clear 3:1; in-cell text 4.5:1.
  https://www.smashingmagazine.com/2024/02/accessibility-standards-empower-better-chart-visual-design/
- **Reorder-point visualization**: show projected inventory level, reorder-point line,
  safety-stock floor, and in-transit orders together. `ROP = (avg daily demand × lead
  time) + safety stock`; `safety stock = Z(service level) × σ`.
  https://pyrops.com/best-practices-to-determine-safety-stock-reorder-point-and-reorder-quantity/
  https://abcsupplychain.com/reorder-point-formula/
  https://www.inflowinventory.com/blog/reorder-point-formula-safety-stock/

*Sourcing caveat:* the Carbon Design System status-indicator wording and davidmathlogic
palette specifics did not render to fetch; their *principles* are corroborated by W3C /
Section 508 and two independent palette sources, but the specific Carbon wording is
unverified (see Rejected list).

## 3. Lead-time offset + production calendar in a horizon grid (leadtime-sim)

- **Lead-time offset is the core "order-by date" mechanism**: shift the planned-receipt
  date backward by the item lead time to get the release/order-by date.
  https://www.oliverwight-americas.com/glossary-terms/lead-time-offset/
  https://docs.oracle.com/cd/E26401_01/doc.122/e48795/T478564T478850.htm
- **Time-phasing is what differentiates MRP from simple reorder-point methods**; the
  engine plans to the exact date of demand even when reports bucket by week/month.
  https://docs.oracle.com/cd/E26401_01/doc.122/e48795/T478564T478850.htm
- **Lead time can be quantity-dependent**: `planned order lead time = fixed lead time +
  (order quantity × variable lead time)`. PLT is a tuning lever, not a fixed input.
  https://docs.oracle.com/cd/E26401_01/doc.122/e48795/T478564T478850.htm
  https://www.sciencedirect.com/science/article/abs/pii/S0925527315001656
- **The production/shop-floor calendar drives which dates are valid**: it defines work
  days; holidays are user-defined; calendars can be defined by shift (the mechanism for
  overtime/extra-shift days). Offsets must land on valid working days; planning on a
  non-working day uses the next valid workday.
  https://docs.oracle.com/en/applications/jd-edwards/supply-chain-manufacturing/9.2/eoash/understanding-shop-floor-calendar-setup.html
  https://docs.oracle.com/cd/E26401_01/doc.122/e48795/T478564T478850.htm
- **Work dates vs. calendar dates** is a deliberate choice; lead-time offsetting in
  working days automatically skips non-production days. Manufacturing calendars number
  operating days consecutively to ease delivery-date calculation.
  https://docs.oracle.com/cd/A60725_05/html/comnls/us/mrp/rpper.htm
  https://www.asprova.jp/mrp/glossary/en/index/m/post-504.html
- **Sizing the horizon**: usually expressed in working days; should not exceed the
  accumulated critical lead time (+ some safety time) — too long creates spurious orders,
  too short risks future shortages.
  https://docs.infor.com/m3udi/latest/en-us/m3beud/scplanhs/wok1625195154608.html
- **Lead-time demand & order signal**: `reorder point = demand during lead time + safety
  stock`; lead-time demand = avg daily demand × lead time. A common variability form for
  safety stock is `(max daily demand × max lead time) − (avg daily demand × avg lead time)`.
  https://www.netstock.com/blog/reorder-point-formula/
  https://www.fishbowlinventory.com/blog/calculating-the-safety-stock-formula-6-variations-key-use-cases

*Sourcing note:* "the order-today trigger comes from net requirements going negative
within the lead-time window, not a static reorder point" is the author's inference,
supported by Oracle's MRP-vs-ROP statement but not a verbatim source claim.

## 4. Operational decision data-grid UX (grid-ux)

- **Exception-first surfacing**: operational dashboards fail when nothing highlights what
  matters; steer users toward actionable anomalies. Tables serve four tasks — find,
  compare, view/edit a single row, take action.
  https://www.smashingmagazine.com/2025/09/ux-strategies-real-time-dashboards/
  https://www.nngroup.com/articles/data-tables/
- **Progressive disclosure**: summaries first (~5 groups max), detail on demand via
  drilldown/filter.
  https://www.smashingmagazine.com/2025/09/ux-strategies-real-time-dashboards/
- **Sort/filter**: sorting on by default from the header; sensible default sort by data
  type; filters discoverable, quick, with a clear "filters active" indicator; highlight
  search matches. Ignition Perspective supports single/multi-sort, type-aware sort, and
  contains/equals/starts-with/ends-with filtering natively.
  https://www.nngroup.com/articles/data-tables/
  https://mui.com/x/react-data-grid/sorting/
  https://www.pencilandpaper.io/articles/ux-pattern-analysis-enterprise-data-tables
  https://www.docs.inductiveautomation.com/docs/8.1/appendix/components/perspective-components/perspective-display-palette/perspective-table
- **Frozen header + leftmost identifier column** when the table exceeds the screen; first
  column should be a human-readable identifier; subtle drop shadow on frozen edges; left-
  align text, right-align numbers; light borders + hover highlight. Perspective supports
  frozen rows/columns and per-cell conditional styling natively.
  https://www.nngroup.com/articles/data-tables/
  https://www.pencilandpaper.io/articles/ux-pattern-analysis-enterprise-data-tables
- **Large-grid performance**: virtualization renders only the visible slice + small
  buffer, keeping DOM count flat; use fixed row heights/column widths. Common rule of
  thumb: client-side under ~1,000 rows, server-side above. Perspective Table has a
  `virtualized` property and configurable `pager`.
  https://www.infragistics.com/blogs/best-react-data-grid-for-large-datasets-performance-guide
  https://www.telerik.com/blogs/best-practices-creating-user-friendly-data-grids
  https://www.docs.inductiveautomation.com/docs/8.1/appendix/components/perspective-components/perspective-display-palette/perspective-table
- **One primary CTA**: exactly one high-emphasis action per view, button-styled, ≥48px,
  action-verb copy, secondary actions de-emphasized; provide a clear empty/zero state.
  https://subux.pro/guides/article/button-hierarchy-primary-secondary-tertiary
  https://balsamiq.com/learn/articles/button-design-best-practices/
  https://www.pencilandpaper.io/articles/ux-pattern-analysis-enterprise-data-tables

*Sourcing caveats:* Infragistics 1M-row benchmark numbers are vendor self-reported and
omitted as load-bearing; Perspective virtualization scroll/query behavior is community-
forum only and unverified; zebra-striping is a genuine tradeoff (NN/g pro vs. Pencil &
Paper caution) — see Rejected list.

## 5. Inventory-policy math vs. forecast-driven simulation (calc-bestpractice)

> All math improvements here are **proposals requiring David's sign-off** (see section 7),
> not assertions that the legacy stored-procedure math is wrong.

- **ROP is a single static trigger; the legacy day-by-day sim is closer to TPOP/MRP.**
  TPOP generates multiple suggested orders by checking whether inventory drops below the
  reorder point within the order horizon, netting on-hand + in-transit — matching the
  legacy forward simulation.
  https://en.wikipedia.org/wiki/Reorder_point
  https://docs.infor.com/ln/2023.x/en-us/lnolh/wholh/whom000316.html
- **TPOP = MRP logic applied to independent (forecast) demand** — validates the legacy
  design rather than contradicting it.
  https://docs.infor.com/ln/2023.x/en-us/lnolh/wholh/whom000316.html
- **Statistical safety stock**: `SS = Z × σ`; Z by service level (90%→1.28, 95%→1.65,
  97.5%→1.96, 99%→2.33). The service-level→Z relationship is nonlinear; typical goals
  90–98%; 100% is statistically unattainable.
  https://web.mit.edu/2.810/www/files/readings/King_SafetyStock.pdf
  https://www.netstock.com/blog/safety-stock-meaning-formula-how-to-calculate/
- **Combined demand + lead-time variability (King formula, independent case)**:
  `SS = Z × √( PC·T⁻¹·σ_D² + σ_LT²·D_avg² )`. When both vary independently this gives a
  **lower** total than summing the two; the additive form applies only when the two are
  dependent. Using the additive form for independent variabilities double-counts risk
  and inflates inventory.
  https://web.mit.edu/2.810/www/files/readings/King_SafetyStock.pdf
  https://abcsupplychain.com/safety-stock-formula-calculation/
- **Set Z per item-group, not one global level** (by strategic importance, margin,
  dollar volume).
  https://web.mit.edu/2.810/www/files/readings/King_SafetyStock.pdf
- **(s,S) / order-up-to / periodic (R,s,S)**: periodic review order-up-to
  `S = demand over (R+LT) + SS`; periodic review consolidates shipments and reduces
  replenishment count — relevant where suppliers ship on fixed days.
  https://link.springer.com/article/10.1007/s10479-021-04441-1
  https://towardsdatascience.com/inventory-management-for-retail-periodic-review-policy-4399330ce8b0/
- **Normality breaks down for intermittent/lumpy demand**; Croston's method (forecasts
  demand size and inter-arrival interval separately) is the standard remedy.
  https://medium.com/data-science/croston-forecast-model-for-intermittent-demand-360287a17f5f
  https://juileetalele.medium.com/croston-model-forecasting-intermittent-demand-data-time-series-analysis-6f3a2bb1654b
- **Automotive EDI cascade — forecast vs. firm**: 830/DELFOR (planning forecast, weeks/
  months) drives strategic planning; 862/DELJIT (firm, next 24–48h) drives execution;
  856/DESADV is the ASN. Treat near-term firm demand differently from far-horizon forecast.
  https://www.orderful.com/blog/automotive-edi-guide

---

## 6. Design implications for the Order grid

1. **Adopt the time-phased record as the row vocabulary**, relabeled for the supplier
   domain: Forecast (Gross Requirements) → On-hand + In-transit (Scheduled Receipts) →
   **Projected Available Balance per working day** → Net Requirements → Planned Order
   Receipts → **Planned Order Releases (= what to order today)**. [mrp-grid; calc]
2. **Columns = the production-calendar horizon, one per working day.** Skip/flag non-
   production days from the calendar rather than treating them as plain calendar days;
   cap the horizon at the cumulative critical lead time (+ safety time). [mrp-grid;
   leadtime-sim]
3. **Compute PAB as a running left-to-right balance** (like MD04 / Planner Workbench), so
   each cell reflects the projected position after that day's forecast, receipts, and
   planned orders. [mrp-grid]
4. **Implement lead-time offset against the working calendar**, per supplier/weekday: the
   release column is the receipt column shifted earlier by the supplier's lead time
   counted in production days, landing only on valid working days. This is the single
   most error-prone mechanic. [mrp-grid; leadtime-sim]
5. **Drive the grid by exception/action messages** (Order / Expedite / Defer / Cancel /
   Below-safety-stock), not raw numbers, with a time/freeze fence so the grid does not
   nag inside the committed near-term window. [mrp-grid]
6. **Status signaling is never color-only (WCAG 1.4.1, Level A).** Each status cell
   carries a unique {color + icon-shape + text-token} triple, fully readable in
   grayscale ("remove all color, still usable" test). This overrides the raw red/amber/
   green convention assumed elsewhere in the research. [reorder-viz; WCAG]
7. **Replace raw RAG with a CVD-safe palette** — Okabe-Ito (`#009E73` healthy,
   `#E69F00` warning, `#D55E00` shortage) or IBM (`#648fff` / `#ffb000` / `#dc267f`).
   Avoid red↔green as the only differentiator. Meet contrast: fills/borders/icons ≥ 3:1,
   in-cell text ≥ 4.5:1. Validate final hex pairs with a CVD simulator. [reorder-viz]
8. **Provide pegging-style drill-down**: clicking a planned-order-release cell reveals the
   forecast/demand days and on-hand/in-transit lines that drove it, plus a per-part chart
   of projected on-hand vs. reorder-point line vs. safety-stock floor. [mrp-grid;
   reorder-viz]
9. **Surface exceptions first**: default-sort/open on parts needing ordering today, with
   native Perspective sort + filter and a visible "filters active" indicator; provide a
   clear "nothing to order today" empty state. [grid-ux]
10. **Freeze part number + description (leftmost) and the header row**; time-phased day
    columns scroll horizontally; consider a frozen rightmost "suggested order qty / total"
    column; right-align quantities, left-align text; persist each operator's column
    setup. [grid-ux]
11. **Enable Perspective `virtualized` + `pager` with fixed row height**; client-side is
    fine under ~1,000 rows, push heavier filtering/simulation to the gateway/DB for larger
    sets. [grid-ux]
12. **One high-emphasis primary CTA** ("Commit Orders for Today"), ≥48px, action-verb
    copy, strong contrast; secondary actions (recalculate, export, adjust horizon)
    subordinate. Feed it via multi-interval row selection + editable order-qty cells
    (Enter to commit). [grid-ux]
13. **Treat firm near-term (DELJIT/862) demand separately from far-horizon forecast
    (DELFOR/830)**: firm days drive hard order signals; forecast days drive planning-only
    / softer cues. [calc-bestpractice]

## 7. Candidate calc improvements (NEED DAVID SIGN-OFF)

Each is a proposed algorithm change with its citation. None is presented as a defect in
the current system; all should be re-validated against the legacy stored-procedure math
before adoption.

1. **Document the engine as TPOP/MRP, not ROP**, so future maintainers don't "fix" the
   forward day-by-day net-requirements sim toward a static reorder point.
   https://docs.infor.com/ln/2023.x/en-us/lnolh/wholh/whom000316.html
   https://en.wikipedia.org/wiki/Reorder_point
2. **Add an explicit, parameterized safety-stock layer**: per-part `SS = Z × σ` driven by
   a configurable service level (decide σ source + horizon).
   https://web.mit.edu/2.810/www/files/readings/King_SafetyStock.pdf
3. **Use the King combined formula** for parts where both demand and lead time vary —
   `Z × √(demand-variance term + lead-time-variance term)` — and avoid the additive form
   unless the two are genuinely correlated (validate independence per part).
   https://web.mit.edu/2.810/www/files/readings/King_SafetyStock.pdf
4. **Set service level / Z per item-group** (ABC / strategic importance / volume / margin)
   rather than one global level; expose Z as a grid-configurable parameter.
   https://web.mit.edu/2.810/www/files/readings/King_SafetyStock.pdf
5. **Make the order-today trigger a net-requirements signal**, not just a static reorder
   point: fire when simulated lead-time demand + safety stock exceeds projected available
   + in-transit. Keep `ROP = lead-time demand + safety stock` as the fallback formula.
   https://www.netstock.com/blog/reorder-point-formula/
   https://docs.oracle.com/cd/E26401_01/doc.122/e48795/T478564T478850.htm
6. **Model lead time as fixed + variable** (`fixed + qty × variable`) to support per-
   weekday / quantity-dependent supplier lead times.
   https://docs.oracle.com/cd/E26401_01/doc.122/e48795/T478564T478850.htm
7. **Flag intermittent/lumpy parts** and warn that normal-distribution safety stock is
   unreliable for them; consider Croston or a flat days-of-cover rule (requires
   identifying which parts are intermittent).
   https://medium.com/data-science/croston-forecast-model-for-intermittent-demand-360287a17f5f
8. **Consider an order-up-to (S) / periodic-review (R,s,S) overlay where shipment
   consolidation matters** (supplier ships on fixed days).
   https://towardsdatascience.com/inventory-management-for-retail-periodic-review-policy-4399330ce8b0/
9. **Treat firm (862/DELJIT) vs. forecast (830/DELFOR) demand separately** in the trigger
   logic and column semantics.
   https://www.orderful.com/blog/automotive-edi-guide

*Unsourced-but-recommended (flagged):* keep predictable per-weekday lead times as a
**deterministic production-calendar offset, NOT as a σ_LT term** — author's engineering
judgment, consistent with the King distinction (predictable offsets are not random
variability) but with no citable source. Reserve σ_LT for genuinely random supplier delay.

## 8. Rejected / unsourced claims

- **Carbon Design System status-indicator exact wording** — page did not render to fetch;
  the principle (color + icon + text) is retained via W3C/Section 508, but the specific
  Carbon quote is dropped as unverified.
- **davidmathlogic / David Nichols "Coloring for Colorblindness" palette specifics** —
  page did not render; usable only as a simulator pointer, not as a palette source.
- **Infragistics 1M-row benchmark numbers** (3.38s→0.42s sort, ~400–600 DOM nodes,
  50+ FPS, etc.) — vendor self-reported marketing figures, not independently verified;
  dropped as load-bearing. The general virtualization principle is retained.
- **Perspective virtualization scroll "feel" / query re-execution during scroll** —
  community forum (PerryAJ), explicitly unanswered, no official IA documentation; not
  relied upon.
- **Zebra striping as a scannability aid** — genuine conflict between NN/g (pro) and
  Pencil & Paper (caution against combining with hover/selected/disabled states); left as
  a noted tradeoff, not a recommendation.
- **Mapping color thresholds to cycle-service-level bands** — design proposal with no
  source on color-coding inventory grids specifically; not adopted as cited best practice.
- **Specific MRP heat-map color palettes** — vendor-defined, not a standardized spec;
  treated as implementation choice, superseded by the WCAG-driven CVD-safe palette above.
- **"Per-weekday lead time as deterministic calendar offset vs. σ_LT"** — retained but
  explicitly flagged as author engineering judgment, not a citable source (see section 7).
