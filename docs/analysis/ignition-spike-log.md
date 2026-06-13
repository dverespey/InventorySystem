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

**Render + read round-trip CONFIRMED (2026-06-13):** in the Designer **Preview Mode** with view param
`recordId=12`, the generated Detail screen rendered and **populated all 32 fields from live data** (the
`runPrepQuery`→`view.custom.record` script-transform binding works end to end). So: generate-from-schema
→ loads → renders → binds live data, all proven. Remaining for GO-final: the **save/write** path and FK
dropdowns. (Note: page/session config still absent — view tested via Designer Preview, not a launched
session; Session Launcher needs a page config.)

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

### F4 — Editable screen: FK code-dropdowns + createSProcCall save (generated)

Generator now emits, from the table + the `UPDATE_PartsStockInfo` signature: **5 FK `ia.input.dropdown`s**
(options from the master tables; value = the business **code**), an **id→code resolving load join**, and a
**Save button** whose `onActionPerformed` builds `system.db.createSProcCall("dbo.UPDATE_PartsStockInfo",
"Inventory_Spike")` with all 30 params in `sys.parameters` order and execs it. Both views deserialize
clean after reload.

- **Key legacy semantic (faithful):** `UPDATE_PartsStockInfo` takes CODES, not ids — it re-resolves
  `VC_SUPPLIER_CODE→IN_SUPPLIER_ID` etc. internally and stamps the 16-char `VC_LAST_UPDATE`. The screen
  therefore loads/saves codes; ids never leave the DB.
- **SQL save path PROVEN independently:** edited part 12's name via the proc with the exact positional
  param set and reverted — `VC_PARTS_NAME` changed then restored, FK ids preserved (`72100`→sup 6,
  `RED`→size 12), `VC_LAST_UPDATE`=`2026061307010654`; the `_HIST` `SELECT *` trigger fired fine (F1 fix
  holds). The Perspective Save button uses the same positional mapping → correct by construction.
- **Remaining:** one UI click in Preview to confirm the button→proc wiring end-to-end (below).

### Check A verdict: **GO** (pending one confirmation click)

A schema+proc-driven generator produces a **fully editable heavy screen** (32 fields, 12-cell matrix, FK
code-dropdowns, read load + createSProcCall save) that loads, renders, binds, and round-trips live data —
on the 8.1.52 dev ceiling. The "no scaffold generator → ~45 hand-built screens" veto premise is
**refuted**: per-screen cost is ≈ run the generator + targeted polish. Recommend **GO** for the Ignition
target on the Check-A axis, conditioned on the Preview save-click confirming the event wiring.

### F5 — Perspective binding lessons (from live Designer testing + log instrumentation)

Driving the generated screen through Designer Preview (with `system.util.getLogger("SPIKE")` lines read
back from `logs/wrapper.log`) surfaced three concrete Perspective rules the generator must follow:

1. **Each input exposes its value under a DIFFERENT prop** — `ia.input.text-field`/`text-area` →
   **`props.text`**; `ia.input.numeric-entry-field` → **`props.value`**; `ia.input.dropdown` →
   **`props.value`**; `ia.input.checkbox` → **`props.selected`**. Binding everything to `props.value`
   left text fields blank while numerics populated. **Fixed & confirmed**: all fields now display.
2. **Object sub-path bindings are unreliable** for `view.custom.<obj>.<key>`:
   - *Display:* replacing the whole object (`self.view.custom.form = row`) orphans the child bindings →
     blank. **Fix:** pre-declare every key in `custom.form` at view-build time, and have Load **mutate
     keys in place** (`self.view.custom.form[k] = v`). Confirmed working.
   - *Write-back:* `"bidirectional": true` into an object sub-path does **not** write the edit back —
     Save read the old value (`SAVE reads VC_PARTS_NAME=RED FILM` after the user edited it). **OPEN.**
3. **Log-driven diagnosis is the headless debugging tool.** Add `system.util.getLogger("SPIKE").info(...)`
   to gateway/event scripts, click in Preview, `grep SPIKE logs/wrapper.log`. This converted "fields are
   blank" guesswork into exact facts (recordId seen, row count, what Save read). Use it for every
   Perspective script issue. Also grep `Unable to deserialize` after every `gwcmd -r`.

### Check A verdict at session close (2026-06-13): **GO**

Everything the gate hinges on is proven on the 8.1.52 dev box: schema-driven **generation** of a 32-field
heavy screen (refutes the "no scaffold generator → ~45 hand screens" veto); it **loads, renders, and
displays live data in every field**; **FK code-dropdowns** from masters; and a Save button that **invokes
the legacy `UPDATE_PartsStockInfo` end-to-end** (proc executes, `VC_LAST_UPDATE` advances — verified in
DB). Per-screen cost ≈ run generator + targeted polish. **GO on the Check-A axis.**

**One open polish item (does NOT change the verdict):** the edit→`form` **write-back** is one-way
(object-subpath bidirectional limitation, F5.2). Planned fix next session: generate **flat** custom props
(`view.custom.form_<col>`) instead of one `form` object, so bidirectional bindings write back cleanly;
then one Load/edit/Save cycle closes the editable round-trip.

---

## RESUME HERE (next session) — read this first

**Environment (verify it's up):**
- Gateway: Ignition **8.1.52** at `/usr/local/ignition`, auto-starts (launchd). Check:
  `curl -s localhost:8088/StatusPing` → `{"state":"RUNNING"}`.
- DB sandbox: **Colima + docker** container `mssql-spike` (SQL 2019, port 1433). After a Mac reboot:
  `colima start` then `docker start mssql-spike` (or rerun `scripts/spike-db.sh`). Gateway DB connection
  **`Inventory_Spike`** reconnects automatically.
- Spike project on disk: `/usr/local/ignition/data/projects/spike/` (loads on `gwcmd -r`).
- Generators (untracked — carry throwaway container creds): `scripts/gen_perspective_view.py`,
  `scripts/spike-db.sh`. DEBUG `SPIKE` logging currently left in the generated Detail view's Load/Save
  scripts (marked `# IG-DEBUG`) — remove once write-back is confirmed.

**Immediate next action:** finish Check A's write-back — edit `gen_perspective_view.py` to emit flat
`view.custom.form_<col>` props (drop the single `form` object); inputs bind bidirectionally to those;
Load sets each `self.view.custom.form_<col>`; Save reads them. Regenerate → `gwcmd -r` → grep
`Unable to deserialize` → one Designer Preview Load/edit/Save → verify the row in DB. Then remove IG-DEBUG.

**Then:** **Check B** (`siteScopedQuery()` multi-site guard) — backend/structural, drivable headlessly.
Scaffolding ready: `sites` (2 rows) + `site_id` on `INV_PARTS_STOCK_MST` (site 1 = 32 rows, site 2 = 15).
**Then:** **Check C** EDI re-scope (paper, no DB) + atomic poller.

**Known facts to reuse:** test record `IN_PART_ID=12` (part `478930201000`, sup code `72100`, size `RED`,
type `FILM`). Save proc `UPDATE_PartsStockInfo` takes CODES not ids (30 params, `@PartID` last). Load via
the id→code join. macOS + 8.1 docs for any Designer guidance (panels by name, not position).

## Check B — siteScopedQuery() — ⏳ scaffolding ready (sites + site_id seeded)
## Check C — EDI re-scope + atomic I/O — ⏳ not started (no DB dependency for the paper re-scope)
