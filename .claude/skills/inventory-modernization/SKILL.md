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
**web app** (Ruby on Rails primary, Python for EDI + forecasting math). SQL Server is
kept as-is initially; Postgres is a later phase.

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
| [templates/module-analysis-template.md](templates/module-analysis-template.md) | Fill-in template for deep-diving one module |
| [scripts/sql.sh](scripts/sql.sh) | Helper: read the UTF-16 SQL files (grep them safely) |

## Methodology — analyzing one module

1. **Identify the form** in `module-map.md` and its `.pas`/`.dfm` pair.
2. **List the procs it calls** — grep the `.pas` for `.ProcedureName :=` / command text,
   cross-reference `database-objects.md`.
3. **List the tables + their triggers** the module touches.
4. **Extract the business rules** from the proc/trigger bodies (this is the real spec —
   the Delphi code is mostly data binding).
5. **Fill the template** (`templates/module-analysis-template.md`) → save to
   `docs/analysis/<area>/<module>.md`.
6. **Map to target**: Rails model(s)/controller/view, or Python service for EDI/forecast.
7. **Update this skill's references** with anything newly learned.

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

Analysis is incremental. See `docs/analysis/` for completed module specs and
`references/migration-strategy.md` for the phase checklist / progress.
