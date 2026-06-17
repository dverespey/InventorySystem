---
name: inventory-modernization
description: >
  Analyze the legacy Delphi 7 InventorySystem and drive its rebuild into a modern
  Ruby/Python web application. Use when working on the modernization: understanding a
  legacy form, mapping stored procedures to app code, planning a module migration, or
  recording analysis findings. Holds the domain glossary, database-object inventory,
  module map, migration strategy, and a repeatable per-module analysis methodology.
---

# Inventory System Modernization

Skill for reverse-engineering the legacy **Delphi 7 + SQL Server** InventorySystem
(automotive parts / Toyota-TEMA EDI logistics) and rebuilding it as a modern
**web app**. SQL Server is kept as-is initially; Postgres is a later phase.

> **Target stack is under reconsideration (2026-06).** The working plan was **Ruby on Rails**
> (primary) + **Python** (EDI/forecasting). **Inductive Automation Ignition + Perspective** is now
> being evaluated as an alternative, to consolidate with the sibling GALC→Ignition migration —
> verdict so far **LEAN-GO**, gated on a vertical-slice spike (see below). **The analysis is
> stack-neutral and survives either choice** — only each module spec's §6 "Target design" section is
> platform-specific. Don't assume Rails.

## When to use this skill

- Analyzing a legacy Delphi form/unit and writing its migration spec
- Mapping stored procedures / triggers to modern app code
- Planning the rebuild of a functional module (Ordering, Shipping, EDI, etc.)
- Looking up domain terms, table meanings, or the proc/trigger inventory
- Recording new findings so they accumulate across sessions

## Core insight (read this first)

**The business logic lives in the database, not the Delphi code.** There are
**179 stored procedures, 24 triggers, 41 tables, 0 views**. Delphi forms (~45) are
thin UI over named CRUD procs. So:

- The **stored procedures are the behavioral spec.** Read the proc, not just the form.
- The **triggers enforce data invariants** (especially inventory-quantity balance).
  Each must be deliberately re-homed in the app layer or DB.
- A Delphi form ≈ one web "module" = (UI screen) + (the procs it calls) + (the
  triggers on its tables).

## Reference material

| File | What's in it |
|------|--------------|
| [references/domain-glossary.md](references/domain-glossary.md) | Toyota/automotive/EDI terms (Renban, ASN, CAMEX, broadcasting code, 810/856/830) |
| [references/database-objects.md](references/database-objects.md) | All 41 tables, 179 procs (grouped), 24 triggers, and what they govern |
| [references/module-map.md](references/module-map.md) | The ~45 live forms grouped into functional areas → target web modules |
| [references/migration-strategy.md](references/migration-strategy.md) | Phased roadmap, stack rationale, proc→code mapping pattern, parallel-run plan |
| [references/cross-cutting-patterns.md](references/cross-cutting-patterns.md) | Recurring patterns (dup-check, string timestamps, enums, triggers) + multi-site lens — **check before each module** |
| [templates/module-analysis-template.md](templates/module-analysis-template.md) | Fill-in template for deep-diving one module |
| [scripts/sql.sh](scripts/sql.sh) | Helper: read the UTF-16 SQL files (grep them safely) |

### Project decision artifacts (live in `docs/analysis/`, not the skill)

| File | What's in it |
|------|--------------|
| [`docs/analysis/decisions.md`](../../../docs/analysis/decisions.md) | **Domain-decisions log** (`D1`, `D2`, …) — the domain expert's answers that close the §8 "open questions." Each spec references the `D#` that resolves its question. **Read before designing any module.** |
| [`docs/analysis/ignition-feasibility.md`](../../../docs/analysis/ignition-feasibility.md) | Multi-agent go/no-go on the Rails→Ignition target switch (LEAN-GO). |
| [`docs/analysis/ignition-spike-plan.md`](../../../docs/analysis/ignition-spike-plan.md) | The gating **vertical-slice spike** that converts LEAN-GO → GO/STAY. |

## Methodology — analyzing one module

1. **Identify the form** in `module-map.md` and its `.pas`/`.dfm` pair.
2. **List the procs it calls** — grep the `.pas` for `.ProcedureName :=` / command text,
   cross-reference `database-objects.md`.
3. **List the tables + their triggers** the module touches.
4. **Extract the business rules** from the proc/trigger bodies (this is the real spec —
   the Delphi code is mostly data binding).
5. **Fill the template** (`templates/module-analysis-template.md`) → save to
   `docs/analysis/<area>/<module>.md`.
6. **Map to target**: the module spec's §6 "Target design" (Rails model/controller/view or Python
   service — or Ignition Perspective view + named query, depending on the stack decision).
7. **Update this skill's references** with anything newly learned.

## Methodology — recording a domain decision (the `D#` workflow)

Each module spec ends with **§8 "Open questions for the user (domain expert)."** Those questions get
answered over time; capture each answer so it propagates and never has to be re-asked.

1. **Collect & dedupe.** Re-extract every spec's §8 (`awk '/^## 8\./{f=1}/^## 9\./{f=0}f' <spec>`).
   Questions that recur across specs become **cross-cutting decisions** (e.g. multi-site, key strategy,
   delete policy) — answer those once, for all specs.
2. **Verify before recording.** Check the answer against the actual proc/trigger/Delphi source first —
   especially "confirm-and-fix bug" questions. Never record a claim from a spec's own citation; the
   verify pass routinely catches errors (miscounted 16-char `yyyymmddHHMMSSff` timestamps; over-claimed
   editability). Cite `file:line` / `schema:line` in the decision.
3. **Record as `D#`** in [`docs/analysis/decisions.md`](../../../docs/analysis/decisions.md) — newest at
   the bottom: a title + date, the **verbatim intent** of the expert's answer, and a "**What this means
   for the rebuild**" section. List the specs/§-numbers it resolves.
4. **Propagate.** In every affected spec, mark the §8 item **✅ RESOLVED (D#)** with a one-paragraph
   summary, and update the relevant **§2/§6/§7** sections (design, schema, migration) to match. One
   answer can resolve several §8 items at once and reshape sub-questions (e.g. "time-bounded pricing"
   also settles the overlap-constraint and `start>end` questions).
5. **Commit per decision-batch** (`Record decision D# … and propagate …`).

**Decisions are platform-neutral** — they describe what the system must *do*, so they survive a
target-stack change. (D1–D8 are recorded; see the log.)

## Methodology — de-risking with a vertical-slice spike

Before committing to a high-uncertainty architecture choice (a target platform, a tricky integration),
run a **time-boxed decision spike** — *not* the start of the build.

1. **Pick the worst-case slice** (the richest screen / hardest path), so a pass generalizes.
2. **Define 2–4 checks, each with an explicit pass threshold**, and a **decision rubric** naming which
   check is a veto (e.g. UI velocity) vs merely mitigable.
3. **Anchor every claim in source** — verify control counts, proc branches, file formats *before*
   scoping, so the spike measures reality (`PartsStockMaster.dfm` = 74 objects; `EDIUpload.pas` =
   830/862/997/824/820 + `delSL[4]` site filter; etc.).
4. **State non-goals** so it can't slide into the real build (don't build all screens; don't rewrite
   any §6 spec until the spike passes).
5. **For a platform go/no-go,** a useful pattern is a multi-agent review — *target-architect* designs,
   *incumbent-architect* states the baseline to beat, *adversarial-reviewer* tries to refute, then
   synthesize a verdict. (Worked example: the Ignition feasibility review + spike plan in
   `docs/analysis/`.)

## Hard-won gotchas

- **SQL files are UTF-16LE.** `grep` returns nothing silently. Always pipe through
  `iconv -f UTF-16LE -t UTF-8`, or use `scripts/sql.sh`.
- **`InventorySystem.dpr` is the authoritative list of live units.** A `.pas` not
  listed there is dead code (several `*old`/`*1`/`New` duplicates were already removed).
- **A form = `.pas` (logic) + `.dfm` (layout).** Read both; `.dcu` is compiled, ignore.
- **Config is INI-driven** (`InventorySystem.INI`) — DB connection strings (3 catalogs:
  Inventory, Activity, VehicleOrder), site/EDI feature flags. Contains plaintext passwords.
- **Cannot build/run the Delphi app here** — no CLI toolchain. Analysis is static.

## Status

Analysis is incremental. See `docs/analysis/` for completed module specs,
[`docs/analysis/decisions.md`](../../../docs/analysis/decisions.md) for the locked domain decisions
(`D1`–`D8`), and `references/migration-strategy.md` for the phase checklist / progress. Two areas are
fully analyzed (Master-data, Inventory/Stock); the target stack is at a **GO/STAY gate** pending the
[Ignition spike](../../../docs/analysis/ignition-spike-plan.md).

**Worked example:** [`docs/analysis/master-data/supplier.md`](../../../docs/analysis/master-data/supplier.md)
is the first fully analyzed module — use it as the reference for the depth/shape a
completed module spec should have.
