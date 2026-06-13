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

## Check A — UI velocity — 🔄 in progress

**JDBC connection live:** gateway connection **`Inventory_Spike`** → `localhost:1433/Inventory`,
status *Valid* (connect URL `jdbc:sqlserver://localhost:1433;databaseName=Inventory;encrypt=true;trustServerCertificate=true`).

### F2 — Perspective views load from hand-authored on-disk JSON (codegen path PROVEN)

Created `data/projects/spike/` by hand (project.json + one Perspective view as
`com.inductiveautomation.perspective/views/Test/{view.json,resource.json}`), restarted the gateway, and
it **imported the project, validated the view, and re-signed `resource.json`** (stamped
`lastModificationSignature`). Gateway log: `Starting project: spike` + `Setting LastModification to
"external" on spike/Test`. **No Designer GUI required to create Perspective views.**

- **Why this is the headline Check-A result:** Perspective's weakness (no scaffold generator → per-field
  drag/bind × ~45 screens) is the whole reason Check A is the gating veto. If view JSON can be
  *generated from the table schema*, that cost line collapses from "hours of drag-drop per screen" to
  "run a generator." This directly attacks the one veto.
- **Caveat (don't over-claim yet):** must still prove a *generated* heavy screen (the ~40-control
  PartsStockMaster) actually renders + round-trips data, not just a one-label view. In progress.
- **Format notes (8.1.52):** view type `ia.container.coord` / `ia.display.label`; resource folder
  `com.inductiveautomation.perspective`; `resource.json` needs `scope/version/files`, gateway fills
  `attributes`. On-disk projects load on **gateway restart** (no live FS watch observed).
### F3 — A heavy CRUD screen GENERATED from schema loads in Perspective

`scripts/gen_perspective_view.py` reads `INV_PARTS_STOCK_MST`'s columns from the live DB and emits the
**PartsStockMaster Detail view (32 fields incl. the 12-cell weekday matrix) + a List view** (Table bound
to `SELECT_PartsStockInfo`). After a gateway restart both views **deserialize and load clean** (gateway
re-signs them; no `Unable to deserialize` warnings). Type→component mapping: int→numeric-entry, bit→
checkbox, long varchar→text-area, else text-field; FK ids grouped; audit cols read-only; `site_id`
hidden. Data binding is a self-contained Perspective **script transform** calling
`system.db.runPrepQuery(..., "Inventory_Spike")` — no Named-Query resource needed to prove the screen.

**Preliminary Check-A read (trending GO, not yet final):**
- The veto premise — "no scaffold generator → per-field drag/bind × ~45 screens" — is **directly
  undercut**: a generator emits any table's CRUD screen in seconds. One-time cost = building the
  generator + format discovery (done this session); per-screen cost ≈ run generator + targeted polish.
- **Still to prove before GO is final (honest gaps):** (1) the screen **visually renders + is usable**
  (open in Designer/browser); (2) the **save/write round-trip** via `system.db.createSProcCall` on the
  CRUD proc; (3) **FK combos** generated as real `ia.input.dropdown` (currently numeric stubs);
  (4) validation/formatting + the weekday-matrix layout look; (5) a **manual Designer build** of one
  screen timed as the baseline to quantify codegen-vs-hand.
- **8.1.52 caveat:** generated on the dev ceiling; component prop schemas may differ slightly on 8.3
  (`# IG83-TODO` in the generator). Nothing 8.1-only used so far.

**Next for Check A:** David eyeballs the generated views in the Designer (fidelity + the matrix), and
times a hand-built screen for the baseline; then I add FK dropdowns + the createSProcCall save path.

## Check B — siteScopedQuery() — ⏳ scaffolding ready (sites + site_id seeded)
## Check C — EDI re-scope + atomic I/O — ⏳ not started (no DB dependency for the paper re-scope)
