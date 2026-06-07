# Modernization Analysis

Deep-dive analysis for rebuilding the legacy Delphi InventorySystem as a modern
web app. The **methodology, domain knowledge, and inventories** live in the
`inventory-modernization` skill at
[`.claude/skills/inventory-modernization/`](../../.claude/skills/inventory-modernization/SKILL.md).

This folder holds **per-module specs** produced as we work through the system,
one file per module, organized by functional area:

```
docs/analysis/
  <area>/<module>.md     e.g. ordering/order.md, edi/asn-856.md
```

Use the template:
[`.claude/skills/inventory-modernization/templates/module-analysis-template.md`](../../.claude/skills/inventory-modernization/templates/module-analysis-template.md)

## Headline findings (from the first pass)
- **The business logic is in the database, not the Delphi code:** 41 tables,
  **179 stored procedures**, **24 triggers**, 0 views. Delphi forms (~45) are thin UI
  over named CRUD procs + 29 `REPORT_*` procs.
- **Procs = the behavioral spec.** Migration of a feature = understand its procs/triggers,
  then re-express them in app code (wrap first, reimplement later).
- **Triggers enforce inventory-quantity invariants** — must be deliberately re-homed.
- Clean functional decomposition into ~11 areas (see the skill's `module-map.md`).

## Functional areas (rebuild units)
Ordering & Renban · Forecasting/FRS · Receiving · Shipping · EDI (810/856/830) ·
Inventory/Stock · Assembly · Master data · Production calendar · Reporting ·
Admin/auth/shell.

## Progress
| Area | Spec | Rebuilt |
|------|:----:|:-------:|
| Master data | ✅ [Supplier](master-data/supplier.md) · [Logistics](master-data/logistics.md) · [Size](master-data/size.md) · [Manifest cost](master-data/manifest-cost.md) · [Master-maint hub](master-data/master-maint.md) | ⬜ |
| Production calendar | ⬜ | ⬜ |
| Inventory / Stock | ✅ [Parts-stock master](inventory-stock/parts-stock-master.md) · [Stocktaking](inventory-stock/stocktaking.md) · [Inv-mgmt](inventory-stock/inv-mgmt.md) · [Logistics breakdown](inventory-stock/logistics-breakdown.md) | ⬜ |
| Receiving | ⬜ | ⬜ |
| Shipping | ⬜ | ⬜ |
| Ordering & Renban | ⬜ | ⬜ |
| Forecasting / FRS | ⬜ | ⬜ |
| EDI / billing | ⬜ | ⬜ |
| Assembly | ⬜ | ⬜ |
| Reporting | ⬜ | ⬜ |
| Admin / auth | ⬜ | ⬜ |

## Cross-cutting findings
Findings that span modules (live legacy defects, recurring hazards) live in `cross-cutting/`:
- [`datamodule-retry-target-bugs.md`](cross-cutting/datamodule-retry-target-bugs.md) — **29 confirmed
  wrong-target retry-recursion bugs in `DataModule.pas`** (8 CRITICAL, incl. 4 `Delete*` methods that
  can silently delete an unrelated supplier on a transient error). Legacy-hotfix candidates +
  rebuild fix. See pattern **P12** in the skill's `cross-cutting-patterns.md`.
- [`trigger-source-reconciliation.md`](cross-cutting/trigger-source-reconciliation.md) — the **24 live
  triggers** (`DB Schema/Create Inventory.sql` is authoritative); **`docs/triggers.sql` is an obsolete
  pre-int-FK-refactor snapshot** (keys on dropped string columns, missing 5 triggers). Don't port from it.
