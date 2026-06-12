# Ignition Vertical-Slice Spike — Plan & Exit Criteria

*Scoped 2026-06-12. Gating action for the **LEAN-GO** verdict in
[`ignition-feasibility.md`](ignition-feasibility.md).* This is a **decision spike**, not the start of
the build. Its only job is to convert the three open uncertainties into a hard **GO / STAY** answer
**before** any `§6` "Target design (Rails primary)" section is rewritten for Ignition.

> **Time-box: ~1–2 weeks (solo).** If a check can't be settled in its budget, that *is* the signal —
> stop and record the result. Do not let the spike slide into building the real app.

---

## Objective

Prove (or disprove) the three things the feasibility review could not settle on paper:

| # | Uncertainty | Why it's the gate |
|---|---|---|
| **A** | Perspective **UI build velocity** for a heavy CRUD screen | The ~45 form-heavy screens are the dominant cost line and Ignition's weakest area (no scaffold generator). If per-screen cost is too high for a solo dev, that is the *one* real reason to stay with Rails. |
| **B** | **Multi-site (D1) enforcement** without a `default_scope` analog | A forgotten `site_id` predicate = cross-site data leak. #1 data-integrity risk. Must prove a *structural* guard, not by-discipline. |
| **C** | **EDI as shared-dir file I/O** under parallel run | Two processes (legacy + Ignition) scanning the same dirs can drop a 997 ack or ship a truncated 810. Must prove safe single-owner + atomic I/O, and re-scope the real EDI surface. |

**Exit decision:** A passes within tolerance **and** B/C prove clean → **full GO** (begin §6 redos for
Perspective/Jython). A fails (UI too slow for one dev) → **STAY** with Rails. B or C reveal a fatal,
unmitigable problem → **STAY**.

---

## Prerequisites (Day 0)

- Ignition gateway (trial license is fine for the spike) with **Perspective** + **Reporting** modules.
- A **JDBC connection to the existing SQL Server** (read-only or a restorable copy — **do not** point
  the spike at production write paths).
- A throwaway **`sites`** table stub (2 rows) + a `site_id` column on the one table the slice touches
  (`INV_PARTS_STOCK_MST`), so B can be tested. This is spike scaffolding, **not** the D1 migration.
- Confirm Ignition scripting is **Jython 2.7** in this install (affects C).

---

## Check A — UI velocity on `PartsStockMaster` (the worst-case screen)

**Why this screen:** the feasibility review named it the richest form. Confirmed from
`PartsStockMaster.dfm` (74 objects): **~40 data controls** — 16 `TMaskEdit`, 6 `TEdit`, 5 `TComboBox`,
1 `TcurrEdit`, 1 `TNUMMIColumnComboBox`, 1 `TCheckBox`, 1 `TDBGrid`, 6 `TButton` — plus 32 labels, and
a **12-cell weekday matrix** (`LeadtimeMonday..Saturday_MaskEdit` + `ShipDaysMonday..Saturday_MaskEdit`).
If this one is tolerable, the simpler masters (Supplier/Size/Logistics) are easy.

**Build (in Perspective):**
1. A reusable **list view** — `Table` bound to a Named Query wrapping `SELECT_PartsStockInfo`, with
   server-side search/sort/paging (replaces the legacy in-memory filter, P7).
2. A reusable **detail (embedded) view** parameterized by `recordId`, all ~40 controls bound to
   `view.params`/a form object; FK combos (`TComboBox`) → `Dropdown`s sourced from master Named Queries
   posting the **surrogate id** (D2). Model the weekday matrix as a `FlexRepeater` or a small grid.
3. A **master/detail container** wiring list-selection → detail.
4. Wire **save** through `system.db.createSProcCall` on the existing CRUD procs (read-only or sandbox DB).
   Keep `IN_QTY` **read-only** on this screen (per parts-stock-master §8.1 / D-recommendations).

**Measure:** actual wall-clock hours to (a) build the reusable list+detail scaffold *once*, and
(b) lay out this one heavy screen on top of it. Then extrapolate a blended estimate over ~45 screens
(heavy screens ≈ this; simple masters ≈ a fraction).

**Pass threshold (calibrate, don't guess):** scaffold + this heavy screen land within ~3–4 days, and
the extrapolated ~45-screen total is a UI budget you're willing to carry (review's rough band was
+2–4 dev-months vs Rails). If the *first* screen blows past a week, that's a STAY signal.

---

## Check B — Site isolation without `default_scope`

**Goal:** prove no query can return another site's rows *by construction*, since Ignition has no
ActiveRecord `default_scope`/`acts_as_tenant`.

**Build:**
1. Derive **`siteId` server-side** from the authenticated session (Perspective session prop seeded from
   the IdP/role) — never from a client-settable parameter.
2. A single **`siteScopedQuery()`** gateway helper that is the *only* sanctioned way to run a query: it
   injects `AND site_id = :siteId` (or calls site-scoped Named Queries) and refuses to run a query that
   lacks the predicate. All screens call through it.
3. Apply it to the `PartsStockMaster` list from Check A against the 2-site stub.

**Pass threshold:** with two seeded sites, the slice shows only the current site's parts; a deliberate
attempt to query cross-site (omit the predicate / spoof the param) is **structurally blocked** by the
wrapper, not merely "we remembered to filter." Document the pattern as the D1 enforcement mechanism.

---

## Check C — EDI re-scope + single-owner, atomic file I/O

**First, re-scope on paper (half a day).** The feasibility review found the EDI surface is ~2× the
original assumption. Confirmed from source this session:
- **`EDIUpload.pas` (inbound, 500 lines)** branches on the doc type and handles **`830` / `862` /
  `997` / `824` / `820`**, where `997` runs an **`AK1` acknowledgement loop** accepting/rejecting prior
  **`856` and `810`** (with `LogActLog` side-effects), plus an *unexpected-EDI* and a *NOTEDI* fallback.
- A **per-site filter lives inside the parser** — `delSL[4]` ("Trading Partner Search"). **This is a D1
  multi-site hook our module specs have not yet captured** — fold it into the future EDI/receiving spec.
- Inbound files **are** archived (`CopyFile.Movefile := TRUE`, `CopyTo …\Archive\…`) — so re-ingest
  protection partly exists today; confirm it's atomic.
- **Outbound** 810/856 build lives in **`ASNInvoice.pas`** (~lines 820–879), **not** the **dead**
  `Write810File.pas` (do not port the dead unit).
- Transport is **out of scope** — the existing **SFTP/VAN** integration moves files; the app only
  reads/writes the shared dir (`Data_Module.fiEDIIn`).

**Then prototype (in a gateway timer/event script):**
1. An **inbound poller** that scans the shared dir, dispatches on doc type for **one** representative
   type (suggest `830` — the simplest), parses the X12 in Jython, and writes via a proc call.
2. **Atomic handling:** write/move via **temp-then-rename**, and **archive on success** so a partially
   written or in-flight file is never half-processed.
3. **Single-owner proof:** establish that during parallel run **exactly one** process owns the dir
   (e.g. the legacy app *or* the Ignition poller, never both) — or separate inbound dirs per owner.
   Demonstrate no archive-move race and no double-processing.
4. **Outbound atomicity:** prototype writing one `810` to a temp name then renaming into place, so the
   SFTP picker never grabs a truncated file.
5. **Calibrate one matrix report** (`REPORT_NUMMILotLocationW`) in the Ignition Reporting module to
   confirm the 29 `REPORT_*` procs are tractable (the matrix shape is the hard case).

**Pass threshold:** the `830` round-trips cleanly; temp-then-rename + archive prevents partial/double
processing; a single-owner (or split-dir) scheme is workable during parallel run; X12 string parsing in
Jython 2.7 is comfortable; the matrix report renders without contortions. No fatal blocker surfaces.

---

## Decision rubric

| Check | Pass → | Fail → |
|---|---|---|
| **A** UI velocity | continue | **STAY** with Rails (the one real veto) |
| **B** site isolation | continue | mitigate (it's a known-mitigable pattern); only STAY if no structural guard is possible |
| **C** EDI file I/O | **GO** | STAY only if a *fatal, unmitigable* hazard appears (none expected) |

**All pass → full GO:** start rewriting the `§6` sections for Perspective/named-queries/Jython,
beginning with the masters. Keep **D1–D8** and spec **§1–§5/§7–§9** exactly as they are.

## Explicit non-goals (keep the spike honest)

- **Do NOT** build all ~45 screens, real auth, or the full EDI matrix — one slice each.
- **Do NOT** rewrite any `§6` spec section until the spike passes.
- **Do NOT** run the Ignition EDI poller against the live shared dir alongside the legacy app without
  the single-owner guard — that is the exact parallel-run hazard the spike exists to *prevent*.
- **Do NOT** write to production data — sandbox/restorable DB only.

---

*Source anchors verified 2026-06-12:* `PartsStockMaster.dfm` (74 objects, 12-cell weekday matrix);
`EDIUpload.pas` (830/862/997+AK1/824/820, `delSL[4]` site filter, `Archive` move); outbound in
`ASNInvoice.pas` (not the dead `Write810File.pas`). See [`ignition-feasibility.md`](ignition-feasibility.md).
