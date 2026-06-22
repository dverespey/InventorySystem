# M3 Server-Side Report-Render Architecture

**Phase:** M3 design (2026-06-21) **Author:** ignition-architect **Status:** DESIGN — hand to
ignition-developer to build; §9 flags items for adversarial-architect-reviewer.

> **What this is.** The target architecture that retires the last Excel/OLE report paths (the 22 report
> families catalogued in `m3-report-inventory.md`) and renders them **server-side** on the Ignition
> Gateway — **additive, read-only, runs live alongside the legacy Delphi app** during parallel run. The
> M3 gate is **numbers match legacy** (with a few proven-broken reports getting a *swappable corrected*
> data query — see §2.4 / §4).
>
> This doc decides the **render mechanism** (8.1 spike + 8.3 prod, verified against the gateway and IA
> docs), defines the **reusable Named-Query → render → output harness** so the 22 reports are *config,
> not 22 bespoke builds*, and applies it to **Daily Shipping (the P0 failing path)** as the template.
> It builds on the M1/M2 producer-recipe pattern (`edi810/code.py` pure-module + thin driver +
> `jython_shim.py` headless runner) and the per-proc Named-Query practice.

---

## 0. The legacy pattern we are replacing (one sentence)

`Inv_DataSet` on a `REPORT_*` proc → `CreateOleObject('Excel.Application')` → open
`TemplateDir\ReportTemplate.xls` → write `mysheet.cells[r,c]` → `SaveAs(fiReportsOutputDir +
'\<name>' + yyyymmddhhmmss00 + '.xls')` → optional `MessageDlg('Print this report?')` →
`mysheet.PrintOut`. (`reporting.md §1a`; per-report cell maps in `m3-report-inventory.md`.)

Three brittleness sources to remove outright: **client Excel + OLE**, **the `ReportTemplate.xls` file
dependency**, **client printer coupling**. One *correctness* source to preserve-or-fix: **the
`REPORT_*` proc** (the data — where the real bugs live, e.g. Daily Shipping's fan-out / orphan drop).

---

## 1. THE RENDER MECHANISM (decided + verified)

### 1.1 What's actually on the spike (verified 2026-06-21, this machine)

| Capability | Verified fact (spike `8.1.52`) | Implication |
|---|---|---|
| Gateway version | `gwcmd -i` → `8.1.52 (64-bit)`, RUNNING, :8088 | dev ceiling, per `project-ignition-version-constraint` |
| **Reporting module** | `Reporting-module.modl` INSTALLED; `reporting-gateway-6.1.52.jar` in `data/jar-cache/` | the module *runs* on 8.1.52 |
| `system.report.executeReport` / `executeAndDistribute` | exist on **8.1** (IA docs) — `executeReport` returns `List[Byte]`; `executeAndDistribute` runs save/print/email/FTP actions | usable *if* a report resource exists |
| **Apache POI** | `poi-4.1.2.jar`, `poi-ooxml-4.1.2.jar`, `poi-ooxml-schemas-4.1.2.jar` on `lib/core/common/` (gateway classpath) | **a Jython gateway lib can build a true `.xlsx` server-side** — HSSF (`.xls`) and XSSF (`.xlsx`) both present |
| PDF engine | `icepdf-viewer-6.3.2.jar` is bundled **inside the Reporting module** only; **no standalone iText/PDFBox/FOP on the common classpath** | PDF is reachable **through the Reporting module**, not from a plain Jython lib |
| Headless authoring | Report **execution** is headless (IA: "Reports run on the Gateway … completely headless generation"); report **resource layout** is a Designer-authored project resource with a serialized data file | same wall as Named-Query `data.bin` (`reference-headless-ignition-authoring-limits` #1) |

### 1.2 The decision — two render lanes, one data layer

The render is split by **output format**, because that's where the headless-authoring wall and the
8.1/8.3 delta actually fall:

```
                         ┌─────────────────────────────────────────────┐
   Named Query / NQ      │   report_render (Jython gateway library)     │
   (the ONLY DB touch) ──▶  build_rows(report_key, params) -> dataset   │
                         │         │                                    │
                         │  ┌──────┴───────┐         ┌────────────────┐ │
                         │  │ XLSX lane     │         │ PDF/print lane │ │
                         │  │ Apache POI    │         │ Reporting mod  │ │
                         │  │ (PRIMARY)     │         │ (8.3 / opt)    │ │
                         │  └──────┬───────┘         └───────┬────────┘ │
                         └─────────┼─────────────────────────┼──────────┘
                                   ▼                          ▼
                      fiReportsOutputDir\<name><ts>.xlsx   PDF -> watched dir / printer
                      (tmp-then-rename)                    (executeAndDistribute)
```

**Lane A — XLSX (PRIMARY; the SaveAs replacement). Same on 8.1.52 and 8.3.**
A **Jython gateway library** (`report_render`) builds the workbook with **Apache POI** (XSSF), writing
cells exactly where the legacy `mysheet.cells[r,c]` wrote them (the render def, §2). Output is written
**tmp-then-rename** to `fiReportsOutputDir` (the M2 atomic-write discipline; replaces `SaveAs`). This is
the chosen mechanism because it is:
- **fully headless-authorable** — it's Python + a classpath jar, NOT a Designer-serialized resource, so
  it sidesteps the `data.bin`/report-resource wall (`reference-headless-ignition-authoring-limits`);
- **fully e2e-drivable headless** — invoked via the existing `jython_shim.py` / a gateway message
  handler, the file is read back and asserted (§5);
- **version-neutral** — POI 4.1.2 is on the common classpath on 8.1.52 and ships the same on 8.3;
- **byte-faithful to the human contract** — POI gives per-cell control (column widths, number formats,
  the `############` / `yyyy/mm/dd` formats the legacy set), so the rebuilt `.xlsx` matches the legacy
  `.xls` layout cell-for-cell.

> `.xls` vs `.xlsx`: legacy wrote `.xls` (BIFF). POI's HSSF can emit legacy `.xls`, but **emit `.xlsx`**
> (XSSF) — operators open it in the same Excel, and `.xlsx` is the supported long-term format. Filename
> changes extension only (`<name><ts>.xlsx`). Flag as a no-op divergence (decide-and-flag, §4/§9).

**Lane B — PDF / print (the PrintOut replacement; 8.3-preferred).**
The legacy optional `mysheet.PrintOut` is replaced by a **PDF written server-side to a watched
directory** (operators/print service pick it up) — *not* an interactive client print. The mechanism:
- **8.3 prod (and 8.1 where a report resource exists):** a thin **Reporting-module template** per
  report family bound to the *same Named Query*, driven by `system.report.executeReport(path, project,
  params, "pdf")` (returns bytes → write to dir) or `system.report.executeAndDistribute(..., action,
  actionSettings)` for save/print/FTP. The Reporting module bundles its own PDF engine (icepdf), so no
  extra lib is needed. **The report-template layout is Designer-authored** (the headless wall) — so on
  the **8.1 spike, PDF templates are authored in the Designer once and exported to disk**, or PDF is
  **deferred** (see fallback).
- **8.1 spike fallback (no Designer pass):** **Perspective Request-Print** (`8.1.28+`) on the report
  view, or simply ship **XLSX-only** for the spike and add PDF templates when the Designer is used for
  prod. The spike's job is to prove the **data + XLSX** path headless; PDF is print-formatting, not a
  numbers-parity concern, so it does not gate M3.

**Why not "Reporting module as the primary path"?** The Reporting module *executes* headless, but its
**report resource (the layout) cannot be reliably hand-authored on disk** — it's the same
Designer-serialized-resource wall that blocked Named-Query `data.bin`
(`reference-headless-ignition-authoring-limits` #1). Making it primary would force a Designer pass for
every one of the 22 reports before any of them could be built/tested on the spike, defeating the
headless build loop the whole fleet runs on. POI keeps the **build + parity test** fully headless; the
Reporting module is reserved for the *formatted-PDF/print* surface on 8.3 where a Designer is available.

### 1.3 The single shared data layer (feeds BOTH lanes)

Both lanes consume **the same dataset**, produced by **one place** per report: a **Named Query** wrapping
the `REPORT_*` proc (per the per-proc NQ practice). Because real Named-Query resources can't be
hand-authored headless (`data.bin` wall), the spike uses the **proven Order-spike substitute**: the SQL
lives as a source-of-truth `.sql` doc and is executed via **`system.db.runPrepQuery(sql, args, conn)`**
inside the render library. This is a *promotable* shape — when the Designer is used for prod, each `.sql`
becomes a real Named Query resource and the render lib calls `system.db.runNamedQuery` instead, with no
change to the render def or output. (Same substitution the Order/Supplier spikes used.)

> Net mechanism table:
>
> | | 8.1.52 spike | 8.3 prod |
> |---|---|---|
> | **Data** | `runPrepQuery(report.sql)` (NQ-shaped, promotable) | promoted **Named Query** (`runNamedQuery`) |
> | **XLSX** (primary) | **Apache POI** Jython lib | **Apache POI** Jython lib (identical) |
> | **PDF/print** (opt) | Perspective Request-Print *or* defer | **Reporting module** `executeReport`/`executeAndDistribute` (PDF→dir/printer) |
> | **Interactive view** | Perspective (param picker + table + export button) | Perspective (same) |

---

## 2. THE REUSABLE HARNESS (config, not 22 builds)

A report = **(a) a query**, **(b) a render def**, **(c) an output**. All three are **data** keyed by a
`report_key`; the engine is generic. Mirrors the M1/M2 pure-module + driver split.

### 2.1 Component layout (on disk, headless-buildable)

```
docs/analysis/reporting/sql/
    daily_shipping_parts.sql          # NQ #1 (corrected detail) — source of truth, promotable to NQ
    daily_shipping_parts_faithful.sql # NQ #1-faithful (the legacy fan-out, for parity proof)
    daily_shipping_header.sql         # NQ #2 (header band)
    <one .sql per REPORT_* / proc>    # mirrors schema/procs (NQ CRUD practice)

project-library/report_render/
    code.py        # the GENERIC engine: build_dataset(), render_xlsx(), write_output(), run_report()
    report_defs.py # the CONFIG: REPORTS = { "daily_shipping_tw": {...}, ... } (the 22 entries)
    driver.py      # gateway entry: run_report(report_key, params, do_pdf) -> path  (thin, like edi810 driver)
```

`code.py` is **PURE** where it can be (the render-def → cell-plan mapping is pure list-building, unit-
testable in CPython); it references `system.db` / POI only at call time (the `edi856.send_856` rule), so
the module imports cleanly under `jython_shim.py` for the headless runner.

### 2.2 The render-def schema (mirrors `mysheet.cells`)

Each report's entry in `report_defs.py` is declarative and maps 1:1 onto the legacy cell map:

```python
"daily_shipping_tw": {
    "title":    "Daily Shipping (Tire/Wheel Part Numbers)",   # legacy cell [1,1]
    "out_name": "DailyShippingTW",                            # legacy SaveAs <name>; ts suffix added
    "query":    "daily_shipping_parts",                       # NQ key (detail rows)
    "header_query": "daily_shipping_header",                  # NQ key (header band; None if n/a)
    # HEADER BAND — fixed cells written from header_query's single row + params (legacy rows 1-2)
    "header_cells": [
        # (row, col, source)   source ∈ {"title"} | ("param", name) | ("header", col) | ("lit", text)
        (1, 1, "title"),
        (2, 1, ("param",  "PDate",        "yyyy/mm/dd")),     # production date, formatted
        (2, 2, ("header", "StartEnd")),                       # 'Start Seq:..../End Seq:....'
        (2, 3, ("header", "VehicleCount", "Vehicle Count:")), # label prefix preserved
    ],
    # COLUMN BAND — the detail table header row + the data columns (legacy row 3 + loop)
    "header_row": 3,
    "columns": [
        # (header_text, query_col, excel_width, number_format)
        ("Part Number",      "Part Number", 17, "############"),
        ("Part Description", "Desc",        30, None),
        ("Qty",              "PQty",         5, None),
    ],
    "first_data_row": 4,                                      # legacy z = 4..n
    "totals": None,                                           # optional totals band (other reports)
    "pdf_template": "reports/daily_shipping",                 # 8.3 Reporting-module path (opt); None ok
}
```

The engine reads this and: writes the header band, writes the column-header row with widths/formats,
loops the detail rows into `first_data_row..n`, applies number formats, writes a totals band if present,
then hands the workbook to the output stage. **Adding report #23 = one dict entry + one (or two)
`.sql` files** — no new code. That is the "config not bespoke" lever the 45-screen CRUD generation
strategy uses, applied to reports.

### 2.3 The output stage (replaces `SaveAs` + `PrintOut`)

`write_output(workbook, out_name, do_pdf, params)`:
1. `ts = formatdatetime('yyyymmddhhmmss00', now)` — **reproduce the legacy 14-digit + literal `00`
   suffix** (`m3-report-inventory.md §0`; not centiseconds) so filenames are continuous with legacy.
2. Write XLSX to `fiReportsOutputDir\<out_name><ts>.xlsx.tmp`, `flush`, **rename** to `.xlsx` (atomic;
   the M2 temp-then-rename discipline — never expose a half-written report). `fiReportsOutputDir` comes
   from gateway config, not a client INI.
3. `LogActLog('REPORT', ...)` equivalent — log a `SPIKE report/<key>: <rows> rows -> <path>` line (the
   e2e count anchor, §5) and an Ignition audit/log entry.
4. If `do_pdf` and a `pdf_template` is set and the Reporting module path is available
   (`system.util.getVersion()` ≥ 8.3 **or** a resource exists): `executeReport(template, project,
   params, "pdf")` → write the PDF beside the XLSX (or `executeAndDistribute` to a printer). Otherwise
   skip PDF (XLSX is the gate-bearing artifact).

### 2.4 The faithful ↔ corrected seam (one-line swap)

The proven-broken reports (Daily Shipping fan-out/orphan; D6 invoice-summary; D12 ship-date range; D11
wheel/tire) need a **swappable corrected query** while a faithful render runs in parallel. The seam is
**the `query` key**: each report def names a query *key*, and the resolver maps the key to a `.sql`
file. A faithful build and a corrected build differ by **which `.sql` the key resolves to** — one line:

```python
# report_defs.py — the ONLY place faithful-vs-corrected is chosen, per report
QUERY_VARIANT = "corrected"   # or "faithful"  (config / per-report override)

def resolve_query(key):
    if QUERY_VARIANT == "faithful":
        return key + "_faithful" if exists(key + "_faithful") else key
    return key            # the corrected/default .sql
```

So `daily_shipping_parts` (corrected, drops the orphan inner-join + the `INV_SHIPPING_INF` fan-out) and
`daily_shipping_parts_faithful` (the legacy proc, reproduces the inflated numbers) are **both** built;
flipping `QUERY_VARIANT` (or a per-report override) swaps which one a given output uses, with **zero**
render-def or engine change. This is what lets M3 (a) prove parity against the faithful query, and
(b) ship the corrected one once David locks the divergence — without a rebuild. The same seam serves the
D6 reports (faithful = window-blind JOIN; corrected = the already-built `fn_ManifestCostAt` CROSS APPLY
from `spike-report-procs-d6.sql`) and D12/D11.

---

## 3. APPLY TO DAILY SHIPPING (P0)

The failing path. Root cause is **proc-side**, not Excel: the T/W procs join `INV_SHIPPING_INF` to
`INV_PART_SHIPPING_INF` on **date alone** (fan-out) AND inner-join `INV_PARTS_STOCK_MST` (orphan drop of
72% of part rows — every tire/wheel part). Full analysis: `daily-shipping-report-spec.md §4`,
`daily-shipping-data-analysis.md §4-5`. The render contract (3 columns + a header band) is in
`daily-shipping-report-spec.md §6`.

### 3.1 The Named Queries (the data layer; `docs/analysis/reporting/sql/`)

**`daily_shipping_parts.sql` — CORRECTED detail** (the default the rebuild ships). Drops
`INV_SHIPPING_INF` from the qty join entirely (`INV_PART_SHIPPING_INF` is the stock-authoritative
per-part qty the triggers maintain). The orphan inner-join to the master is a **separate** decision —
see the variant note:

```sql
-- daily_shipping_parts.sql  (params: @PDate varchar(8))   [Range: swap = to BETWEEN @BeginPDate AND @EndPDate]
SELECT  m.VC_PART_NUMBER  'Part Number',
        m.VC_PARTS_NAME   'Desc',
        SUM(p.IN_QTY)     'PQty'
FROM    INV_PART_SHIPPING_INF p
JOIN    INV_PARTS_STOCK_MST   m ON p.VC_PART_NUMBER = m.VC_PART_NUMBER   -- (see §3.3 orphan decision)
WHERE   p.VC_PRODUCTION_DATE = @PDate
GROUP BY m.in_part_type_ID, m.VC_PART_NUMBER, m.VC_PARTS_NAME
ORDER BY m.in_part_type_ID, m.VC_PART_NUMBER;
```

**`daily_shipping_parts_faithful.sql` — FAITHFUL detail** (the legacy proc verbatim, for the parity
proof; the date-only `INV_SHIPPING_INF` join that inflates ×K):

```sql
-- daily_shipping_parts_faithful.sql  — reproduces REPORT_DailyShipping(Range) EXACTLY (the fan-out)
SELECT  m.VC_PART_NUMBER 'Part Number', m.VC_PARTS_NAME 'Desc', SUM(p.IN_QTY) 'PQty'
FROM    INV_SHIPPING_INF s
JOIN    INV_PART_SHIPPING_INF p ON s.VC_PRODUCTION_DATE = p.VC_PRODUCTION_DATE   -- (!) date-only fan-out
JOIN    INV_PARTS_STOCK_MST   m ON p.VC_PART_NUMBER     = m.VC_PART_NUMBER
WHERE   s.VC_PRODUCTION_DATE = @PDate
GROUP BY m.in_part_type_ID, m.VC_PART_NUMBER, m.VC_PARTS_NAME
ORDER BY m.in_part_type_ID, m.VC_PART_NUMBER;
```
> (Single-date `REPORT_DailyShipping` also keeps the seq cols in GROUP BY → one row per (part×header).
> For the *faithful* numbers parity test we match the **Range** grain (one row/part, PQty = K×truth),
> which is the cleaner, fully-characterized inflation; the single-date per-header row explosion is noted
> as a second faithful variant if a byte-for-byte single-date `.xls` ever needs reproducing.)

**`daily_shipping_header.sql` — header band** (aggregated over the date's shipping headers, NOT joined
into detail — fixes the fan-out at the source of Start/End/Vehicle-Count too):

```sql
-- daily_shipping_header.sql  (params: @PDate)   [Range: BETWEEN]
SELECT  MIN(VC_START_SEQ_NUMBER) 'StartSeq',
        MAX(VC_END_SEQ_NUMBER)   'EndSeq',
        SUM(IN_QTY)              'VehicleCount',
        'Start Seq:' + MIN(VC_START_SEQ_NUMBER) + '/End Seq:' + MAX(VC_END_SEQ_NUMBER) 'StartEnd'
FROM    INV_SHIPPING_INF
WHERE   VC_PRODUCTION_DATE = @PDate;
```

### 3.2 The render def (the 3 cols + header band)

The `daily_shipping_tw` dict in §2.2 *is* the Daily Shipping render def — it reproduces
`daily-shipping-report-spec.md §2a` cell-for-cell: title at `[1,1]`, date at `[2,1]`, Start/End at
`[2,2]`, Vehicle Count at `[2,3]`, the `Part Number/Part Description/Qty` header at row 3 (widths
17/30/5, `############` format on col 1), detail from row 4. The **Range** variant
(`daily_shipping_range_tw`) reuses the same def with: title changed, `header_cells` = begin/end date at
`[2,1]/[2,2]` and **no** Start/End/Vehicle-Count cells (the Range proc omits them — spec §2b), and
`query`/`header_query` params switched to `@BeginPDate/@EndPDate`. The **Assy** variant (R3/R4) keeps
the clean `IN_ASN_ID` join (no fan-out), drops `d.IN_QTY` from GROUP BY, and renders the 2-column
`Part Number | Qty` layout (spec §2c) — same engine, different def.

### 3.3 The faithful ↔ corrected swap point (Daily Shipping)

Two independent corrections, both behind the §2.4 seam:

1. **Fan-out (the §4 root cause):** corrected `daily_shipping_parts` drops `INV_SHIPPING_INF` from the
   qty join → PQty = true per-part shipped qty; faithful keeps it → PQty = K×truth (K = that date's
   `INV_SHIPPING_INF` row count). **This changes a number operators see** → per the divergence rule it
   is an **explicit David decision** in the Daily Shipping ledger (style of `edi810-decisions.md`).
   Recommended: adopt corrected; parity target = the truth query, not the legacy `.xls`.
2. **Orphan drop (the §4.2 inner-join):** on the live spike the inner join to `INV_PARTS_STOCK_MST`
   drops **641/886 part rows (72%)** — every tire/wheel part, because tire/wheel parts aren't in the
   parts-stock master. A LEFT join would suddenly surface 10 more parts with NULL descriptions →
   **also a number/row-count change Toyota/operators see**, and arguably the bigger fidelity issue (a
   report titled "Tire/Wheel" returns zero tire/wheel parts). This is a **SECOND, separate David
   decision** — do NOT silently flip it. Recommended default = **keep the inner join** in
   `daily_shipping_parts` (faithful on *that* axis) so the corrected query changes only the fan-out;
   offer a `daily_shipping_parts_tirewheel.sql` variant (LEFT join + tire/wheel description source) as
   the corrected-orphan option David can elect. (Flagged §9 — this is the higher-stakes of the two.)

### 3.4 Output

`write_output("DailyShippingTW", do_pdf=False, params)` → `fiReportsOutputDir\DailyShippingTW<ts>.xlsx`
via tmp-then-rename. No client Excel, no `ReportTemplate.xls`, no OLE, no PrintOut. Operators who print
get the 8.3 Reporting-module PDF (`pdf_template = "reports/daily_shipping"`) when that path is enabled;
on the spike, XLSX only.

---

## 4. 8.1 → 8.3 DELTAS (greppable markers for the build)

| Concern | 8.1.52 (spike) | 8.3 (prod) | Marker in code |
|---|---|---|---|
| **Data access** | `system.db.runPrepQuery(sql, args, conn)` (NQ `data.bin` can't be hand-authored) | promote each `.sql` to a real **Named Query** → `system.db.runNamedQuery(path, params)` | `# IG83-TODO: promote <key>.sql to Named Query; swap runPrepQuery -> runNamedQuery` |
| **XLSX render** | Apache POI 4.1.2 (common classpath) | Apache POI (same; verify version on the 8.3 box) | `# IG81-COMPAT: POI on common classpath both versions` |
| **PDF / print** | Reporting template needs a Designer-authored resource → **defer PDF or use Perspective Request-Print (8.1.28+)** | **Reporting module** `system.report.executeReport`/`executeAndDistribute` (PDF→dir/printer) | `# IG83-ONLY: report-module PDF path (guard on getVersion()/resource-exists)` |
| **Interactive view** | Perspective table + export button (`system.dataset.toCSV` + `system.perspective.download`) — `data.bin`-free | same | `# IG81-COMPAT: Perspective export, no report resource` |
| **Print action** | `system.perspective.print` / Request-Print Action (8.1.28+) | same + Reporting distribute | `# IG83-TODO: prefer report-module distribute for formatted print` |
| **Version guard** | `system.util.getVersion()` exists 8.1+ | same | gate the PDF branch on it |

**Rule:** the **numbers-gate path (data + XLSX) is 100% version-neutral and 100% headless** — it's the
only thing M3 parity depends on. Everything 8.3-specific is **PDF/print presentation**, guarded and
non-gating. This keeps the whole M3 build runnable + testable on the 8.1.52 spike.

---

## 5. PARITY-TEST APPROACH (numbers match legacy, headless, on 8.1.52)

Three layers, mirroring the M1/M2 + D6 test discipline (`test_report_procs_d6.py`,
`feedback-parity-fixture-fidelity`). All run headless against the live spike DB.

**Layer 1 — query parity (the gate; pure SQL diff).** For each report, run the **faithful `.sql`** and
diff it against the **legacy proc** (`EXEC REPORT_DailyShipping @PDate=...`) on a sample of real
production dates → must be **identical** (proves the faithful query reproduces legacy exactly). Then run
the **corrected `.sql`** and assert it equals the **independent truth query** (e.g. for Daily Shipping,
`SELECT VC_PART_NUMBER, SUM(IN_QTY) FROM INV_PART_SHIPPING_INF WHERE VC_PRODUCTION_DATE=@PDate GROUP BY
...`). This is exactly the shape `test_report_procs_d6.py` already uses (legacy vs migrated vs
independent CROSS-APPLY truth) — reuse `lib.Report` and the `docker exec sqlcmd` harness. Pick the
sample dates **live** (don't hardcode), e.g. dates with `COUNT(*)>1` in `INV_SHIPPING_INF` so the
fan-out actually bites — `daily-shipping-report-spec.md §4` self-flag.

> **Daily Shipping characterization assertion:** for a date with K shipping headers, assert
> `faithful.PQty == K * corrected.PQty` per part (proves we've *characterized* the bug, not just
> diffed). On the current spike data K=1 for every date (`daily-shipping-data-analysis.md §4.1`), so the
> e2e must **seed a 2nd `INV_SHIPPING_INF` row** for a test date (fixture-tagged `VC_ADD`, torn down) to
> exercise K≥2 — otherwise the fan-out parity is vacuously green (the fixture-fidelity lesson).

**Layer 2 — render parity (the XLSX is faithful to the data).** Drive `report_render.run_report(
"daily_shipping_tw", params)` through the gateway via `jython_shim.py` (the R8 headless runner) → it
writes the `.xlsx` to a temp dir → the e2e **reads the workbook back with openpyxl** (CPython side) and
asserts: title cell `[1,1]`, header-band cells, column headers + widths/formats at row 3, and that the
detail rows (`[z,1..3]`) equal the Named-Query rows in order. This proves the **render def → cells**
mapping reproduces `daily-shipping-report-spec.md §2a` without a Designer or a client.

**Layer 3 — count/log anchor (drift-proof).** The driver logs `SPIKE report/daily_shipping_tw: N rows
-> <path>`; the e2e asserts the workbook detail-row count == N == the Named-Query row count, all
computed **live** (never hardcoded — the `_HIST`/virtualization drift lesson,
`reference-headless-ignition-authoring-limits` #6).

New e2e file: `scripts/e2e/test_daily_shipping_render.py` (Layer 1+3 pure-SQL/driver, like
`test_report_procs_d6.py`) + a render-readback assertion (Layer 2). The faithful↔corrected seam is
tested by running both `QUERY_VARIANT`s in one pass and asserting faithful==legacy and
corrected==truth.

---

## 6. BUILD SEQUENCE (hand to ignition-developer)

1. **Engine + Daily Shipping (P0).** Build `report_render/{code.py, report_defs.py, driver.py}` with
   the generic engine and the `daily_shipping_tw` def; the 4 `.sql` files (parts corrected/faithful,
   header, range); `test_daily_shipping_render.py` (all 3 parity layers). Proves the harness end-to-end.
2. **Range + Assy (R2/R3/R4).** New defs only (engine unchanged); R3 reuses the clean ASN path; verify
   the R4 filename collision is fixed (distinct `out_name`).
3. **D6 invoice-summary (R9/R10).** Wire the render to the **already-built** `fn_ManifestCostAt`
   corrected `.sql` (`spike-report-procs-d6.sql`) as the corrected variant; faithful = window-blind.
   Highest business risk → full dual-adversary review (money/EDI).
4. **Order/Invoice/Logistics families (R5–R8), D12 seam.** Straight proc-wrap defs.
5. **C5/C6/C7 EDI echoes** (render parsed inbound X12, no proc) + **C1/C3** order/forecast companions
   (fix the C3 P6 `.frc` crash while there).
6. **R11–R21 + R22 (InvMgmt), C2/C4** — lower-frequency / print-template / already-no-Excel.

Each report after #1 is **config** (a def + `.sql`), not a new build. Right-size review per
`feedback`: full dual-adversary + double-re-verify for the money/EDI/number-semantics reports
(D6/810/856, Daily Shipping divergence); a lighter single pass for display-only/low-risk reports.

---

## 7. WHAT'S DECIDED vs OPEN

**Decided here (decide-and-flag, no number Toyota sees changes):**
- Render mechanism: **POI Jython lib for XLSX (primary, both versions)** + Reporting module for 8.3 PDF.
- Data layer: one **Named Query per proc** (`runPrepQuery` on the spike, promotable), faithful↔corrected
  behind the `query`-key seam.
- Output: tmp-then-rename `.xlsx` to `fiReportsOutputDir`, legacy `yyyymmddhhmmss00` suffix; **`.xls →
  .xlsx` extension change** (no-op divergence, logged).

**Open — needs David (each changes a number/row-count someone sees → divergence rule, surface it):**
1. **Daily Shipping fan-out fix** (§3.3 #1): adopt corrected (smaller, true PQty) vs faithful (×K
   inflated)? Recommended: corrected; record in the Daily Shipping ledger.
2. **Daily Shipping orphan-drop fix** (§3.3 #2): keep the inner join (report stays "no tire/wheel
   parts") vs LEFT-join to surface tire/wheel parts (needs a tire/wheel description source)? Higher
   stakes — the report's title vs content contradiction. **Default keep-inner unless David elects the
   fix.**
3. **PDF on the spike:** defer PDF until the Designer is used for prod, or author the Daily Shipping
   report template in the Designer now (one-time) so the 8.1 spike can demo the PDF lane too?

---

## 8. ASSUMPTIONS DEPENDENT ON UNVERIFIED SOURCE/ENV BEHAVIOR (flag)

- **POI on the 8.3 prod box** is assumed present on the common classpath (verified on 8.1.52; standard
  Ignition bundling, but confirm on the actual 8.3 install). `# IG83-TODO`.
- **`fiReportsOutputDir`** is assumed reachable + writable by the gateway service account as a server-side
  path (legacy was a client INI `[DIRECTORIES]` token). Multi-site (D1) may need per-site dirs — out of
  M3 scope but note it.
- The **faithful single-date `REPORT_DailyShipping` per-header row explosion** (vs the Range one-row
  grain) is characterized but the parity test matches the Range grain; if a byte-for-byte single-date
  `.xls` reproduction is ever required, add the per-header faithful variant.
- The exact **inflation factor K** per date is data-dependent — the e2e computes it live; do not assert
  a fixed K.

---

## 9. FOR adversarial-architect-reviewer

- **§3.3 #2 (orphan-drop)** is the load-bearing call: it's a 72%-row-difference, title-vs-content
  defect, and it's a *separate* divergence from the fan-out. Confirm the design correctly treats it as a
  distinct David decision (not folded into the fan-out fix) and that the recommended default
  (keep-inner) is the right parity stance.
- **POI-vs-Reporting-module** as primary: confirm the headless-authoring rationale (Reporting resource =
  Designer-serialized, same wall as `data.bin`) justifies a custom Jython render lib over the native
  module — the solo-dev-maintainability bar. (Counter-argument to weigh: the Reporting module is the
  Ignition-native option; is a one-time Designer authoring pass per report cheaper long-term than a POI
  render lib? The design bets headless-build-loop > native, because all 22 must be build/test-able on
  the spike before prod.)
- **Faithful↔corrected seam** is a global `QUERY_VARIANT` + per-report override; confirm that's the
  right granularity (vs per-call) for parallel-run, where faithful and corrected may need to run
  simultaneously for the *same* report (parity proof) — the per-report override covers this, verify.
- The **`.xls → .xlsx`** extension change: confirm it's a true no-op for downstream consumers (no
  machine parses these report `.xls` files — they're human-readable copies per
  `m3-report-inventory.md §0`; the machine artifacts are the EDI files, untouched).

---

## 10. BUILD RECORD (ignition-developer, 2026-06-21) — the 4 LIVE reports + the engine

**What shipped** (`project-library/report_render/{code.py, report_defs.py, driver.py}` + `sql/*.sql` +
`scripts/e2e/test_m3_reports.py`): the reusable engine + **the 4 reports that actually ran in a year**
(data-driven prune — `m3-report-usage-prune.md`: 18 of 22 never ran; these 4 are the entire live set):

| key | R# | proc(s) | variant | render | layout source |
|---|---|---|---|---|---|
| `daily_shipping_assy` | R3 | `REPORT_DailyShippingAssy` | FAITHFUL | declarative | `MainMenu.pas:3165-3199` |
| `invoice_summary` | R9 | `REPORT_INVOICESSummary` | **D6 CORRECTED** (default) + `_faithful` | declarative + totals | `MainMenu.pas:3624-3667` |
| `forecast_detail` | R18 | `REPORT_ForecastDetail` | FAITHFUL | custom (group-break) | `MainMenu.pas:3744-3794` |
| `lot_location` | R12 | `REPORT_PLANTLotLocation[W]` | FAITHFUL | custom (two-proc) | `MainMenu.pas:716-861` |

**Render mechanism — POI confirmed runnable HEADLESS on 8.1.52.** `code.render_xlsx` (Apache POI XSSF) was
driven headless via the gateway's **bundled JRE + jython-ia-2.7.3.3 + POI 4.1.2** (no system Java needed):
`/usr/local/ignition/lib/runtime/jre-mac/bin/java -Dpython.path=/usr/local/ignition/user-lib/pylib -cp
jython-ia.jar:lib/core/common/* org.python.util.jython -S <script>`. This is the exact code path that runs
on the gateway; the test reads the produced `.xlsx` back with openpyxl and asserts cells/widths/formats.
(Confirms §1.2's POI bet end-to-end on the spike — the data+XLSX numbers-gate path is headless + version-
neutral, as designed.)

**Two render lanes, honestly applied.** 2 of 4 are pure **declarative config** (Daily Shipping Assy, INVOICE
Summary). The other 2 are **NOT flat tables** — their legacy render has group-break rows + blank separators
the declarative model (§2.2) cannot express faithfully, so each got a **small named PURE plan function**
(`build_forecast_detail_plan`, `build_lot_location_plan`) rather than an invented generic group-break
feature. Forecast Detail: a blank separator row before each part group + part#/eff-month only on the group's
first row (`MainMenu.pas:3768-3794`, hand-transcribed in the test as the independent oracle). Lot Location:
WHEELS proc then TIRES proc into ONE sheet, with `Car Wheels`/`Truck Wheels` (by line name) and size-code
group headers + conditional `assembler_location`-or-`plant_parking` cells + `yyyy/mm/dd` date reformat. The
two-`.sql` dispatch lives in `driver.run_report`. **Adding a flat-table report is still one def + one .sql;**
the two custom plans are the genuinely-non-declarative exceptions, cited per `.pas` line.

**Lot Location IS the live one (CONFIRMED, not a SKIP).** The live `LotLocationClick` handler calls
`REPORT_PLANTLotLocationW` + `REPORT_PLANTLotLocation` (the **PLANT** procs). The D9-deprecated
`REPORT_NUMMILotLocation[W]` is the twin the inventory flagged — it is **NOT** referenced by the live
handler. So Lot Location is a BUILD. (On the current spike data every `vc_status_PLANT_yard=''` → the procs
return 0 rows; the test SEEDS a few rows additive-then-restore so parity is non-vacuous.)

**INVOICE Summary D6 note (the divergence, decide-and-flag).** The legacy `REPORT_INVOICESSummary` is
window-blind (JOIN `INV_MANIFEST_COST_MST` on the assy code with no production-date window → over-bills a
part with >1 price window). The rebuild ships the **corrected** D6 window-aware query (`invoice_summary.sql`
via `fn_ManifestCostAt`, consistent with the locked 810/856 D6 decision) as the **default**; the
window-blind legacy (`invoice_summary_faithful.sql`) sits behind the `QUERY_VARIANT` seam for the parity
proof and the parallel run. A report is not Toyota-facing/state-changing → decide-and-flag (D6 is already
David-locked), documented in `invoice_summary.sql` + `report_defs.py`. The test proves the divergence
NON-VACUOUSLY: injecting a 2nd (gap) price window makes the faithful/legacy output go 34→**42** (+8 = the
over-bill) while the corrected stays **34** (the one covering window) — and a REVERT (CROSS APPLY → the
window-blind JOIN) FAILS the corrected-vs-truth check.

**Test results:** `scripts/e2e/test_m3_reports.py` → **48 PASS / 0 FAIL**. Layer 1 (numbers parity, every
expectation from the legacy proc / an independent truth query, NEVER the rebuild) + Layer 2 (real POI render
read back with openpyxl) + an output-stage unit test (legacy `yyyymmddhhmmss00` filename, `.xlsx`,
tmp-then-rename atomicity) + REVERT non-vacuity proofs. HONEST: numbers are faithful to the legacy proc on
live data; the exact `.xls` VISUAL CHROME (the `ReportTemplate.xls`, borders, page breaks, fonts) is **not**
reproduced — we render the same CELLS, not the template chrome (Lot Location's page-breaks/borders/alignment
are print presentation, non-gating, §4).

### 10.1 SKIP LIST — the 18 dead reports (deprecate-by-deletion at legacy retirement; NOT ported)

Per `m3-report-usage-prune.md`, 18 of the 22 report families have NO usage signal (the one real prod day
ran exactly one report — Daily Shipping Assy). They are **NOT built**; their legacy Excel/OLE paths are
retired by deletion at the legacy-retirement step (not here). Full table (R#, proc, retired Excel path,
why) lives in `report_defs.py::SKIP_REPORTS`. Summary:

| R# | report | legacy proc | retired Excel path |
|---|---|---|---|
| R1 | Daily Shipping (T/W) | `REPORT_DailyShipping` | OLE-template → `DailyShippingTW<ts>.xls` |
| R2 | Daily Shipping Range (T/W) | `REPORT_DailyShippingRange` | OLE-template → `DailyShippingRangeTW<ts>.xls` |
| R4 | Monthly Shipping ASN (Assy) | `REPORT_MonthlyShippingAssy` | OLE-template → `DailyShippingAssy<ts>.xls` |
| R5 | Daily Supplier Order | `REPORT_DailySupplierOrders/Cost` | OLE-template → `DailySupplierOrder<ts>.xls` |
| R6 | Monthly Supplier Order (±cost) | `REPORT_MonthlySupplierOrders/Cost` | OLE-template → `SupplierOrder<ts>.xls` |
| R7 | Monthly Logistics Order | `REPORT_MonthlyLogisticsOrders` | OLE-template → `Logistics<ts>.xls` |
| R8 | Monthly Supplier Invoice | `REPORT_MonthlySupplierInvoices` | OLE-template → `SupplierInvoice<ts>.xls` |
| R10 | Monthly INVOICE Summary | `REPORT_MonthlyINVOICESSummary` | OLE-template → `MonthlyINVOICESummary<ts>.xls` |
| R11 | Logical Inventory | `REPORT_LogicalInventory` | OLE-template → `LogicalInventory<ts>.xls` |
| R13 | Empty Container | `REPORT_EmptyContainer` | OLE-template → `EmptyContainer<ts>.xls` |
| R14 | Past-Due / Late FRS | `REPORT_LATEFRS` | OLE-template → `PastDueFRS<ts>.xls` |
| R15 | PO Report | `REPORT_PO` | OLE-template → `POReport<ts>.xls` |
| R16 | Forecast Parts Summary | `REPORT_ForecastPartsSummary` | OLE-template → `ForecastPartsSummary<ts>.xls` |
| R17 | Forecast Assy Summary | `REPORT_ForecastSummary` | OLE-template → `ForecastAssySummary<ts>.xls` |
| R19 | Forecast vs Usage | (forecast/usage proc) | OLE-template → `ForecastvsUsage<ts>.xls` |
| R20 | Unused Tire Part Numbers | `REPORT_UnusedTirePartNumbers` | OLE-template → `TireWithoutAssembly<ts>.xls` |
| R21 | Unused Wheel Part Numbers | `REPORT_UnusedWheelPartNumbers` | OLE-template → `WheelWithoutAssembly<ts>.xls` |
| R22 | InvMgmt QReport | (none — `Grid_ClientDataSet`) | `TQuickRep.Preview` (on-screen, no Excel) |

Already-decided D9 deprecations (out of rebuild scope entirely; `report_defs.py::SKIP_DEPRECATED_D9`):
`REPORT_ASNWithCost`, `REPORT_NUMMILotLocation[W]` (the deprecated Lot-Location twin — NOT the live R12),
`ForecastCamexreport` (C8).

### 10.2 Pre-existing finding flagged for sql-adversary (NOT M3 scope)

While running the regression suite, `scripts/e2e/test_report_procs_d6.py` showed **9 PASS / 2 FAIL** on this
spike. Root cause is pre-existing + **independent of M3**: a filtered UNIQUE index
(`UX_INV_ASN_MST_LINE_PDATE_NORMAL`, added by the later ASN-unique-guard PR) requires `QUOTED_IDENTIFIER ON`
for any DML on `INV_ASN_MST` (Msg 1934), but that test's `sqlq()` runs with sqlcmd's `-Q` default
(QUOTED_IDENTIFIER OFF). So its EDI856-fan-out **seed INSERT silently fails** → the check was **vacuously
green** before the index existed. Adding the session SET options (the fix the `jython_shim` already applies)
un-masks a **real `legacy=2` vs `migrated=1` divergence** in `REPORT_EDI856_D6`'s forecast fan-out vs the
legacy windowed JOIN when 2 forecast kanbans exist — a D6-migration concern, **owned by the D6/sql-adversary
review, not M3**. I reverted my test edit (left the test at its master baseline) and surface this finding.
