# CreateInventory.sql moved → `db/CreateInventory.sql`

The authoritative live server DDL dump (42 `INV_*` tables, 182 procs, 25 triggers;
2026-06-12) was relocated to **`db/CreateInventory.sql`** in the repo reorg
(`chore/repo-reorg-ignition`), alongside the Ignition app code (`ignition/`),
the e2e harness (`e2e/`), and the migrations/named-queries (`db/migrations`,
`db/namedqueries`).

- **Canonical now:** `db/CreateInventory.sql`
- **Why moved:** the Ignition rebuild's DB provisioning (`db/`) owns the schema
  baseline; the legacy Delphi source stays here, but the schema dump is a
  build-time DB artifact, not Delphi source.
- Analysis specs under `docs/analysis/**` that cite `DB Schema/CreateInventory.sql`
  by `file:line` remain accurate against this same dump content (only the path
  changed); they are the reverse-engineering bridge and are not rewritten.

The prior-snapshot citation anchor `Create Inventory.superseded-2026-06-01.sql`
(used by pre-2026-06-16 specs) STAYS in this directory unchanged.
