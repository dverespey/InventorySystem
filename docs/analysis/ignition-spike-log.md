# Ignition Vertical-Slice Spike — Running Log

*Live record of the GO/STAY spike defined in [`ignition-spike-plan.md`](ignition-spike-plan.md). Append
findings as they happen; this is the artifact the GO/STAY decision is written from.*

Dev env: Ignition **8.1.52** (dev ceiling; prod target 8.3 — see
[`ignition-version-strategy.md`](ignition-version-strategy.md)). DB sandbox: Colima + SQL Server 2019.

---

## Day 0 — environment & DB sandbox (2026-06-13) ✅

| Prereq | Status |
|---|---|
| Gateway 8.1.52, Perspective + Reporting | ✅ live on `:8088` |
| Jython | ✅ 2.7.3.3 |
| Container runtime | ✅ Colima 0.10.3 + docker 29.5.3 (Intel/amd64 native) |
| SQL Server | ✅ `mssql-spike` container, SQL 2019, port 1433 (host-reachable) |
| DB restore | ✅ `Inventory` restored from `DB Schema/Inventory.bak` (**SQL 2008R2** source → compat level **100**) |
| Schema sanity | ✅ **40** `INV_*` tables, **182** procs, **25** triggers; `SELECT_PartsStockInfo` + `INV_PARTS_STOCK_MST` (47 rows) present |
| App login | ✅ least-priv `ignition_spike` (datareader/writer + EXECUTE) — not SA |
| Check-B scaffolding | ✅ `sites` (2 rows) + `site_id` on `INV_PARTS_STOCK_MST` → site 1 = 32 rows, site 2 = 15 rows |
| Ignition JDBC connection | ⏳ **pending** — needs gateway admin web UI (manual). Params below. |

Reproducible bring-up: [`scripts/spike-db.sh`](../../scripts/spike-db.sh).

**Ignition JDBC connection to create** (Config → Databases → Connections): driver *Microsoft SQL Server*
(bundled `mssql-jdbc-9.4.0`), host `localhost`, port `1433`, database `Inventory`, user `ignition_spike`.
Connection string note: with JDBC driver 9.x + a self-signed dev cert, set
`;encrypt=true;trustServerCertificate=true` (or `;encrypt=false`) or the handshake fails.

## Finding F1 (D1 hazard) — `SELECT *` history triggers break when a column is added

Adding `site_id` to `INV_PARTS_STOCK_MST` immediately broke an **UPDATE**: trigger
`UPDATE_PartNumber` runs `INSERT INTO INV_PARTS_STOCK_MST_HIST SELECT * FROM deleted` (positional
`SELECT *`). With `site_id` on the base table but not the history table, column counts mismatch →
*"Column name or number of supplied values does not match table definition."*

- **Implication for D1:** the multi-site migration cannot just add `site_id` to live tables — every
  paired `_HIST` (and any `INSERT … SELECT *` target) must get the column too, **and every such trigger
  must be re-reviewed.** This is a concrete, schema-wide D1 task, not a per-screen detail.
- **Spike workaround applied:** mirrored `site_id` onto `INV_PARTS_STOCK_MST_HIST`; UPDATE + history
  insert then succeed. Captured in `scripts/spike-db.sh`.
- **Action:** fold "audit all `SELECT *` triggers for column-add safety" into the D1 decision record
  when §6 redo begins. (Cross-ref `decisions.md` D1.)

---

## Check A — UI velocity — ⏳ not started (unblocks once JDBC connection exists)
## Check B — siteScopedQuery() — ⏳ scaffolding ready (sites + site_id seeded)
## Check C — EDI re-scope + atomic I/O — ⏳ not started (no DB dependency for the paper re-scope)
