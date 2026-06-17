# Order ("What to Order") — Redesign Plan (A vs B synthesis + spike)

Synthesizes `option-a.md` (Faithful-Modern), `option-b.md` (Best-Practice-Forward),
`legacy-order-spec.md`, `research-best-practice.md`, `source-artifacts.md`. Decides the path,
locks the calc-change governance, and defines the next vertical-slice spike.

Citations: `spec §N` = legacy-order-spec.md; `research §N` = research-best-practice.md;
`artifacts §N` = source-artifacts.md; `A §N` / `B §N` = the two option docs.

---

## 1. Executive summary

**Recommendation: Option A — Faithful-Modern, with Option B's *presentation* layer folded in.**

The two options agree on ~90% of the target: one site-scoped Perspective page (SelectOrderBar +
PhasedGrid), wrap-the-proc data layer mirroring the schema, a read-only Simulate → review →
Commit loop, native non-color-only signaling, and `siteScopedQuery()` as the delivery (not
isolation) mechanism. They diverge on **exactly one axis: the simulation math.**

- **Option A** reproduces the legacy `Order.pas` algorithm bit-for-bit (spec §3) and applies
  best practice *only to the surface* (exception-first ordering, frozen header/columns, WCAG
  1.4.1 non-color-only signaling). Calc changes: **zero.**
- **Option B** proposes six MRP/safety-stock calc changes (C1–C6, B §7), each now cited but
  **all gated behind David's sign-off** — meaning until sign-off, *Option B's math IS Option A's
  math.* B is A plus a backlog of eligible-but-unapproved calc proposals.

The decisive facts: (a) the legacy calc is **trusted daily and not yet proven reproducible** —
the parity diff against a live Delphi run is still owed, and the calc carries a known
index-space landmine (hazard 7, spec §8); (b) the research doc itself states it is **not
asserting the legacy proc math is wrong** (research §5/§7); (c) solo-dev maintainability is a
hard constraint and a parity-only build is the smallest provable thing. Adversarial scoring
confirms it: **A = 0 WRONG / 4 RISK, B = 0 WRONG / 1 RISK** — but B's single "risk" is the
whole MRP-rewrite surface deferred behind sign-off, while A's four risks are concrete,
each-named, each-gated open gaps with owners. We adopt A as the build, carry B §7 as a
post-parity calc backlog, and harvest B's presentation refinements (CVD-safe palette as a
sign-off-gated post-cutover option; the "spike the stock Table before the Flex-Repeater
fallback" discipline) into A.

There is **one calc-related decision A must make that is not a "change"**: the legacy commit
procs (`INSERT_OpenOrder` FRS read-max+insert; renban read-bump) are **unguarded
read-then-write races, safe only because the legacy is single-user** (A §4.1, verified
`Create Inventory.sql:3266-3314`). Parallel run adds concurrency the legacy never had, so A
*adds serialization* (SERIALIZABLE tx + UPDLOCK/HOLDLOCK + a commit-claim row). This is
faithfulness-preserving (same committed values) but is correctly NOT on the calc-change list —
it changes locking, not arithmetic.

---

## 2. A vs B — side by side

| Dimension | Option A — Faithful-Modern | Option B — Best-Practice-Forward |
|---|---|---|
| **Fidelity to legacy calc** | Bit-for-bit. Reproduces calendar walk (spec §3.1), weekday lead-time select (§3.2/§3.4), added-leadtime break-loop (§3.4), end-balance recurrence (§3.5), share `=E/(ΣE)` (§3.2). | Identical to A *until* a C-row is signed off; each signed-off row diverges from legacy by design. Default-state == A. |
| **UX / best practice (cited)** | Surface-only: exception-first default sort + faithful-order toggle (research §4 grid-ux); frozen header/left ID cols (NN/g, Pencil&Paper); WCAG 1.4.1 non-color-only (research §2); virtualized+pager <1000 rows (research §4). Keeps legacy RGBs; CVD palette deferred. | Same surface set + adopts CVD-safe Okabe-Ito/IBM palette now (research §6.7); MRP row vocabulary (PAB/net-requirements, research §6.1) *if* C-rows approved; one primary CTA (research §6.12). |
| **Calc changes** | **None.** | **Six candidates** C1–C6 (B §7): C1 200-row cap (defect, MRP-independent); C2 lead-time by release-day weekday; C3 working-calendar offset; C4 statistical SS `Z×σ`/King; C5 net-requirements trigger; C6 calendar-bucketed forecast + firm/forecast split. All KEEP-PROPOSED, none approved. |
| **Build effort** | Lower. One T-SQL proc `SIM_OrderSimulation` to author + diff vs legacy; no new math to validate. | A's effort **plus** per-approved-C-row: new math, new parity reasoning, re-validate it doesn't perturb untouched cells (B §7 parity-gate). Open-ended. |
| **Parallel-run risk** | 4 named RISK gaps (commit race, commit-claim, multi-site schema surgery, `AD_GetSpecialDate`) — all gated with owners (A §8). Same DB, procs mediate row shape. | 1 RISK (the deferred MRP surface) + inherits all of A's same 4 operational gaps (B §8 lists them too). Any approved calc change widens divergence the parallel Delphi app will NOT match → harder parity. |
| **Solo-dev maintainability** | Highest. One proc, one NQ, "run both and diff." No bespoke MRP engine to own. A §2 puts calc in DB proc (set-based, data-local). B's Jython engine is more readable-as-translation but is a hand-owned 4-function port. | Lower if C-rows land: solo dev now owns + tunes safety-stock/service-level/σ parameters and a net-requirements engine the legacy never had. Maintainable only while C-rows stay deferred. |
| **Adversarial status** | **0 WRONG / 4 RISK** | **0 WRONG / 1 RISK** |

Note the A/B calc-engine *placement* also differs: A §2 runs the sim as a **T-SQL proc**
(data-local, set-based, one round trip, easier to diff in the source's own language); B §4 runs
it as **gateway Jython** (closer line-by-line Pascal translation, no proc to "reimplement").
Both are defensible; A's proc placement is the recommendation **because the calc is destined to
stay faithful** — a set-based proc you diff against the live system beats a 200-row × 23-day
per-row fan-out in Jython. If any B calc change is later approved, revisit placement then.

---

## 3. RECOMMENDATION

**Build Option A. Carry Option B §7 as a post-parity calc backlog. Harvest B's presentation
refinements into A.**

Rationale, grounded in spec + cited research:

1. **The legacy calc is trusted but unproven-in-rebuild.** Spec confidence is HIGH on Pascal
   flow, but the rebuild's correctness is established only by the parity diff vs a live Start —
   which is still owed and is gated behind open gaps (A §8 #4/#5/#7). You cannot responsibly
   layer MRP changes on top of an algorithm you have not yet demonstrated you can reproduce.
   First reproduce (A), then improve (B §7).

2. **Hazard 7 makes faithful-first mandatory.** `fDates` is indexed by calendar offset `x`,
   phased arrays by fill position `j`, reconciled only by the `fDates[i]<>0` scan (spec §8 h7,
   `Order.pas:1353-1369,1457-1468`). A set-based rewrite that doesn't preserve both index spaces
   misplaces in-transit/open buckets by a day. The *only* way to know you got this right is
   cell-for-cell parity against legacy — which is A's defining gate. B's calc changes would
   perturb the very surface this hazard lives on, before it's pinned down.

3. **The research doc disclaims itself as a defect finding.** research §5/§7 explicitly: "not
   asserting the legacy stored-proc math is wrong." Every C2–C6 is framed as a *proposal needing
   sign-off*. So B, honestly read, is A + a sign-off-gated backlog — not a different build.

4. **Solo-dev maintainability (hard constraint) favors A.** A is "one proc, one NQ, diff vs
   legacy." B-with-approved-changes makes the solo dev the owner of statistical safety-stock
   tuning (σ source, per-group Z, service levels — research §7) and a net-requirements engine.
   Prefer the simplest Ignition-native option that meets the requirement (operating principle).

5. **What we take from B regardless:** (a) the **CVD-safe palette** (Okabe-Ito `#009E73` /
   `#E69F00` / `#D55E00`, research §6.7) recorded as a sign-off-gated post-cutover improvement —
   A already logs this (A §5); (b) B's **"prove the stock Perspective Table can carry per-cell
   style+icon+aria before falling to a Flex-Repeater"** discipline (B §2.2/§5) — adopt it as a
   spike acceptance criterion rather than committing to a renderer up front; (c) B's documented
   scope-limits on commit idempotency (token de-dups Ignition's own resubmits only; cross-app
   and renban races need DB locking) — already mirrored in A §4.1.

**Front-runner doc: `/Users/apple/Documents/GitHub/InventorySystem/docs/analysis/order/option-a.md`.**

---

## 4. Consolidated CALC-CHANGE SIGN-OFF LIST

Recommendation is the **faithful calc**, so **the approved-change list is EMPTY**. Every entry
below is **DEFERRED** — carried from Option B §7 / research §7 as eligible candidates, none in
the build. Until David signs a row, the build runs the legacy calc for that row.

| ID | Legacy (cited) | Proposed | Citation | Status |
|---|---|---|---|---|
| C1 | Silent ≤200-row / `fDates[0..200]` cap, no guard (spec §8 h3) | Remove cap; virtualize; **surface a warning on truncation** instead of silent drop | DEFECT FIX; large-grid corroborated research §4/§6.11 | [ ] David — **defect, not MRP-gated; recommend approving the *warning* even under faithful build (no value change)** |
| C2 | Lead time by *today's* weekday col, fallback `IN_LEADTIME` (spec §3.4, `Order.pas:426-459`) | Select lead time by *order-by (release) day's* weekday, offset on working calendar | research §3/§6.4 (Oliver Wight, Oracle JDE) | [ ] David — DEFERRED |
| C3 | Added-leadtime: each in-window overtime day pushes order-by +1, `break` on first miss (spec §3.4, `Order.pas:1576-1582`) | True working-calendar offset skipping non-production, absorbing overtime; replace break-loop | research §3/§6.2/§6.4 (Oracle shop-floor calendar) | [ ] David — DEFERRED |
| C4 | End-balance < fixed `J = usage×days` → red (spec §3.5) | Running PAB vs **parameterized statistical SS** (`SS=Z×σ`, King combined when demand+LT vary, per-group Z) | research §1/§5/§6.3/§7 (King/MIT, Oracle) | [ ] David — DEFERRED |
| C5 | Share `=E/(ΣE)` across size group, 100% singleton (spec §3.2) | Net-requirements netting as the *trigger*; retain `=E/(ΣE)` share as downstream *allocation* | research §1/§5/§6.1/§7 (Infor TPOP, Netstock) | [ ] David — DEFERRED (REVISED: netting is trigger, share retained) |
| C6 | Forecast via week/day breakdown + first-prod-day offset (spec §3.2) | Bucket forecast onto production-calendar working days; firm (862/DELJIT) vs forecast (830/DELFOR) split | research §1/§3/§5/§6.2/§6.13/§7 (Oracle, Orderful) | [ ] David — DEFERRED |
| P1 | (presentation, not calc) Legacy Excel RGBs bit-for-bit (spec §4) | Swap to CVD-safe Okabe-Ito/IBM palette | research §6.7 | [ ] David — DEFERRED (post-cutover; changes what operator sees, meanings preserved) |

**Net: faithful calc => approved list is empty / all deferred.** C1's *warning-on-truncation*
and P1 are the only rows that could be approved without altering any computed value; both still
need David's tick.

---

## 5. SPIKE PLAN — the next vertical slice

**Slice: a read-only, single-site PhasedGrid for ONE part type, proven cell-for-cell and
color-for-color against the legacy Excel output.** No commit path in this spike (zero write risk;
commit is gated on A §8 #1/#2 and comes in a later phase). Simulate-only is the smallest slice
that exercises the highest-risk surface: the calc reproduction + the signal/color reproduction.

### 5.1 What it must prove
1. **Calc parity (the point of the whole thing):** a `SIM_OrderSimulation`-shaped read
   reproduces the legacy grid numbers for a sample of parts, including the hazard-7 index
   reconciliation.
2. **Signal/color parity:** the Pascal-set signals (spec §4 — lead-time zone 36, order-by 40,
   overtime 3, non-prod 4, in-transit font 23, open-order font 10, below-safety FormatCond 3)
   render with the right meaning on the right cells, with the non-color channel (icon/text/aria)
   attached (research §2 / WCAG 1.4.1).
3. **Renderer decision:** a **stock Perspective Table** can carry per-cell `{value, signal}` →
   style + sibling icon + aria for the ~200×26 grid on 8.1.52 with acceptable render perf —
   OR it provably cannot, justifying the Flex-Repeater fallback (B §2.2/§5 discipline).

### 5.2 Success criteria (top 3, gating)
- **SC1 — Number parity:** for the sample parts, every grid cell (forecast, K/L/M/N inventory,
  phased day values, end-balance, summary Total Inv / In Transit / Added Leadtime) matches the
  legacy Excel output **exactly**, including ≥1 named hazard-7 case (a calendar with O/X/holiday
  days between fill positions, so offset≠position) asserting bucket placement byte-for-byte.
- **SC2 — Color/signal parity:** lead-time-zone, order-by, overtime, non-production,
  in-transit, open-order, and below-safety signals land on the **same cells** the legacy sheet
  colors, each rendered with ≥2 channels (color + icon/text/aria), legend decodes every enum,
  and the grid is fully operable with color removed.
- **SC3 — Renderer + perf:** the chosen renderer displays the full ≤200-row × ~26-col grid on
  the 8.1.52 dev box with frozen header + frozen left ID columns and acceptable scroll/sort,
  using stock Table virtualized+pager if it carries dual-channel cells; otherwise the
  Flex-Repeater fallback is documented as required, with the perf evidence.

### 5.3 Data wiring vs the spike DB
- Spike DB is the restored **`Inventory`** database (Colima/docker SQL Server 2019, restoring
  `DB Schema/Inventory.bak`, gitignored) via `scripts/spike-db.sh`; app login
  `ignition_spike`, host `localhost:1433`, bundled JDBC.
  > **DISCREPANCY TO RESOLVE:** the task brief names the connection **`Inventory_Spike`** but
  > `spike-db.sh` restores DB **`Inventory`** and seeds login `ignition_spike`. Decide one name
  > and align the Ignition JDBC connection + `spike-db.sh` `DB=` var before wiring — do not let
  > the NQ point at a connection that doesn't exist.
- All START-path reads (spec §2 #2–#15) wrapped as Named Queries mirroring the procs (A §3.1).
  In the spike these run **read-only** against the existing proc signatures verbatim (no
  `@site_id` param exists — A §7 / B §6).
- `SIM_OrderSimulation` authored against the spike DB as the single sim NQ; internally
  calls/inlines the START procs. **Blocked input:** `AD_GetSpecialDate` lives in the ALC
  `TireOrder` DB, not `Inventory` (spec §8 h1) — see §6 gap, must be stubbed-or-sourced before
  the calendar walk runs.
- `siteScopedQuery()` wired but **inert/pass-through** in the spike (no `site_id` column on the
  read procs) — present as the delivery mechanism, proving the shape, not isolation.

### 5.4 Parallel-run parity tests (numbers AND colors vs legacy Excel)
For a fixed **sample set of parts** spanning: a singleton-size part (share=100%), a
shared-size group (the `=E/(ΣE)` split), a part with in-transit AND open-order buckets, a part
that breaches safety stock (below-safety red), and a calendar window containing O / X / holiday
days (hazard-7).

1. **Capture the legacy baseline:** run the Delphi Order Start for each sample part against the
   same data the spike DB holds; capture the populated `OrderSimulation.xls` (cell values +
   colors). This is the golden output.
2. **Number diff (SC1):** diff `SIM_OrderSimulation` output vs the legacy cell values,
   cell-for-cell. Any mismatch is a parity failure — investigate index-space (hazard 7) first.
3. **Color/signal diff (SC2):** compare the Perspective `signal_enum`/`source_enum` per cell vs
   the legacy ColorIndex per cell. **GATED on closing the §6 CF/palette extraction gap** —
   until the .xls conditional-format rules + palette RGBs are read, signal parity can only be
   asserted for the Pascal-set signals (spec §4), and any template-baked CF threshold is
   flagged UNKNOWN, not passed.
4. **Record per-part PASS/FAIL** with the diffing artifact checked in (not the .bak/.xls
   themselves — gitignored).

### 5.5 Multi-site test (acknowledging "no site_id column")
Reality: grep for `site_id`/`VC_SITE`/`IN_SITE` across the schema = **zero hits** (A §7 / B §6);
the read procs have no site param. So the spike does **NOT** prove row-level isolation — it
proves the *two-phase shape*:
- `spike-db.sh` already seeds a `dbo.sites` table (1=MAS/TMMMS, 2=HERO/TMMTX) and has manually
  added `site_id` to `INV_PARTS_STOCK_MST` + its `_HIST` table in lockstep (F1 hazard handled,
  `spike-db.sh:59-64`). The spike uses this **only** to: (a) bind K6/L6 header labels from
  `sites` (the one real per-site difference, artifacts §3); (b) confirm `siteScopedQuery()`
  degrades to a **pass-through** when the read procs lack a site param (Check B shape).
- **Explicitly out of spike scope:** re-authoring the 15 START + 3 commit procs to filter/stamp
  by site (A §8 #3). The spike demonstrates the screen is single (cosmetic K6/L6 only) and the
  scoping *mechanism* threads from session not client — it does not claim data isolation, which
  needs the schema+proc surgery deferred to the DB-mod phase. Verify the `INV_PARTS_STOCK_MST`
  site_id add did NOT break its `_HIST` `SELECT *` trigger (F1) as a spike side-check.

---

## 6. Carried-over EXTRACTION-GAPS to close before/inside the spike

1. **`.xls` conditional-format thresholds + palette RGBs — UNREAD (gates SC2 signal parity).**
   xlrd 2.0.2 cannot read CF rules or `formatting_info`; openpyxl cannot open BIFF; LibreOffice
   not installed (artifacts §4; spec §5). The spike's color/signal layer reproduces **only the
   Pascal-set signals** (spec §4); any template-baked CF threshold the operator relies on daily
   is **missing and flagged UNKNOWN — do not guess** (A §8 #5/#6).
   **CLOSE PATH (pick one, never guess):**
   - **Option (i):** install LibreOffice headless, `soffice --headless` convert
     `OrderSimulation.xls` → `.xlsx`, then read `worksheet.conditional_formatting` + the
     workbook palette via openpyxl.
   - **Option (ii):** David opens a representative populated workbook in Excel and reads out the
     conditional-format rules + thresholds + the actual index→RGB palette directly.
   Until one is done, SC2 can pass only for Pascal-set signals; full signal parity stays open.

2. **In-cell template formulas — UNREAD (gates SC1 sign-off).** Share `=E/(ΣE)` and end-balance
   recurrence are reproduced from the Pascal spec (§3.2/§3.5), but the template's own formulas
   were unreadable (artifacts §4 gap 3). The SC1 number diff vs live legacy is the catch-net and
   **must include the named hazard-7 case** (A §8 #7).

3. **Order-proc race (commit-phase defect — NOT in this spike, blocks the later commit phase).**
   `INSERT_OpenOrder` (`Create Inventory.sql:3266-3314`) and the renban read-bump are unguarded
   read-max-then-insert races, safe only because legacy is single-user (A §4.1). Confirm with
   delphi-architect whether parallel-run commit is operationally single-writer; if not, the
   SERIALIZABLE tx + UPDLOCK/HOLDLOCK + commit-claim row is **mandatory** before any commit
   ships. The simulate-only spike sidesteps this; flag it loudly for the next phase.

4. **200-row silent cap (defect, spec §8 h3).** Faithful behavior caps at 200 with no guard; the
   spike should **surface a truncation warning** (C1, no value change) and confirm 200 is
   acceptable per site with delphi-architect.

5. **`AD_GetSpecialDate` body + catalog + result shape + status domain (blocks the calendar
   walk).** Cross-DB in ALC `TireOrder`, body unverified (spec §8 h1). The calendar walk
   (spec §3.1) and added-leadtime (§3.4) consume its O/X/holiday statuses. Get the body, result
   shape, and exact status domain from delphi-architect before authoring the calendar; do not
   re-derive the overtime/X/holiday rules. In the spike, stub the calendar with a known fixture
   so SC1 can still run, but mark calendar-derived cells as fixture-backed until the real proc
   is sourced.

---

## Return values

- **Recommendation:** **Option A (Faithful-Modern)** — build A, defer all of B §7 behind David's
  sign-off, harvest B's presentation refinements (CVD palette as gated post-cutover; renderer-spike
  discipline).
- **Front-runner doc:** `/Users/apple/Documents/GitHub/InventorySystem/docs/analysis/order/option-a.md`
- **Top 3 spike success criteria:**
  1. **SC1 Number parity** — every grid cell matches the legacy Excel output exactly for the
     sample parts, including a named hazard-7 (offset≠fill-position) case.
  2. **SC2 Color/signal parity** — Pascal-set signals land on the same cells as legacy, each
     with ≥2 channels (color + icon/text/aria) + decoding legend, grid usable with color removed
     (full parity gated on closing the .xls CF/palette extraction gap).
  3. **SC3 Renderer + perf** — stock Perspective Table carries per-cell dual-channel cells at
     ≤200×~26 with frozen header/left columns and acceptable perf on 8.1.52, else Flex-Repeater
     fallback documented with evidence.
