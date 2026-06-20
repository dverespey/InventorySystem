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

## Order spike (SC2/SC3) — read-only wired PhasedGrid — ✅ loads + wired + verified
- **Proc:** `docs/analysis/order/spike/SIM_OrderSimulation.sql` gained a `@Section char(1)` param
  (A/B/C; NULL=all). IG81-COMPAT: Ignition reads only the FIRST result set of a proc, so the view
  calls the proc ONCE PER SECTION. Each section IF-guarded so it emits exactly one result set
  (an empty `WHERE` would still emit an empty set and become the "first" set). Default (no @Section)
  still emits all three → sqlcmd/parity runs unchanged. Re-applied to `mssql-spike` (verified: A→only A,
  C→only C with no row_kind leak, default→3 sets).
- **NQ source-of-truth:** `docs/analysis/order/spike/named-queries.sql` — three section queries on
  connection `Inventory_Spike` (the DB-access-layer artifacts). NOT promoted to on-disk NQ *resources*:
  the on-disk NQ metadata-JSON schema is undocumented and hand-authoring it blind risks a
  non-deserializing resource + wasted restart; wired instead via the proven `system.db.runPrepQuery`
  script-transform pattern (same as `PartsStockMaster/List`). Promotion = mechanical paste into the
  Designer Named Query workspace. **Follow-up.**
- **View:** `data/projects/spike/.../views/Order/OrderSpike/{view.json,resource.json}`. SelectOrderBar
  (read-only Today, Line/PartType/SortBy dropdowns, FillDays, Simulate, Commit DISABLED "later phase")
  + PhasedGrid (`ia.display.table`, virtualized + pager, frozen-left identity cols, dynamic per-day
  cols from result set A) + decoding Legend. One `custom.gridModel` script-transform runs A/B/C, pivots
  C per (part,fill_pos), emits **per-cell `{value, style}`** dual-channel cells (8.0+ Table feature):
  background = legacy palette (#FF0000/#00FF00/#FFCC99/#FFFF99), font = source (#333399/#008000),
  below-safety = red bold + ⚠ word; glyph/text prefix ([OT]/[X]/★/[LT]/🚚/📦) is the non-color channel.
  Exception-first default sort + Faithful (size-group) toggle. SC3 renderer = `ia.display.table`.
- **Verify:** `gwcmd -r`; `grep "Unable to deserialize"` (new log lines) = **0**; log shows
  `Setting LastModification to "external" on spike/Order/OrderSpike` + gateway re-signed resource.json
  → parsed/imported OK. `Inventory_Spike` datasource restarted successfully. Dual-channel proof: WHEEL
  4261102Q8000 fill_pos 2 = NON_PRODUCTION bg + IN_TRANSIT font simultaneously (hazard-7 06-18 X-day).
- **Markers:** `IG81-COMPAT` (section-per-call), `IG83-TODO` candidates: native column-freeze if 8.3
  adds richer grid styling. **Fixture-backed:** calendar OT/X/holiday cells (AD_GetSpecialDate stubbed).
  **PENDING golden:** byte-for-byte vs live legacy Excel (SC1 — no Delphi/Excel here).

### Update 2026-06-15 — golden received; faithful-layout REBUILD + SC1 parity + E2E automation
- **Golden in hand:** `DB Schema/OrderSimulationCorolla{Tire,Wheel,Valve,Film}.xls` (real client data,
  gitignored). Decoded → real layout = **pooled 4-row ledger** (Beg / Receipts-per-supplier / Usage /
  End). View **REBUILT** to it: numbers+color-only (peach editable + red below-safety), glyphs dropped,
  order entry locked to each supplier's order-by cell, client-side live End recompute. Fleet:
  developer → code-reviewer → qa.
- **Calendar:** retired fictional fixture; encoded REAL calendar from golden (weekends + **7/3** July-4
  observed + **7/13–7/17** shutdown week). Req#3 resolved: "Lead (P)" = single `IN_LEADTIME`/weekday,
  order-by = Today + P prod days.
- **SC1 parity:** `parity_diff.py` vs all 4 golden → **14/20 groups cell-for-cell, order-by 22/22**
  (`sc1-parity-results.md`). Open: R1 FILM forecast week-number mapping; R3 WHEEL M1 receipt filter.
- **E2E automation:** Playwright headless harness `scripts/e2e/` — auto-resets trial, asserts UI +
  `SPIKE` markers + edit, screenshots. **12/12 PASS, zero clicks.** `domId`s added+loaded; gateway
  restart now permitted (`.claude/settings.json`). Wired into the `ignition-qa` agent. Resolved the
  code-reviewer's "transform-never-ran" + qualified-value RISKs live. **Resume:** `RESUME-order-spike.md`.

## Check B — siteScopedQuery() — ⏳ scaffolding ready (sites + site_id seeded)
## Check C — EDI re-scope + atomic I/O — ⏳ not started (no DB dependency for the paper re-scope)

## Stock-Ledger Service — ✅ foundation built; ⚠️ reconciliation blocked by purge horizon (2026-06-17)
- **Built (all verified on the box):** `INV_STOCK_LEDGER` table + UNIQUE `(IN_PART_ID,VC_SOURCE_EVENT)`
  + FK→parts (`docs/analysis/inventory-stock/spike-stock-ledger-table.sql`); `POST_StockMovement` proc
  (atomic insert-ledger + additive `IN_QTY+=delta`, idempotent, purge-aware) + `PROC_RebuildStockBalance`
  (F4 read-then-write under `UPDLOCK,HOLDLOCK`/SERIALIZABLE) (`…/spike-post-stockmovement-proc.sql`);
  `stockLedger` Project-Library service `post()/rebuildBalance()/resolvePartId()` via `createSProcCall`
  (`data/projects/spike/ignition/script-python/stockLedger/code.py` — loads clean, gateway re-signed it,
  no script errors after `gwcmd -r`).
- **Proc self-test (6/6 PASS):** post bumps IN_QTY by delta; ledger row + balance_after correct; replay
  same event key does NOT double-post (idempotent); purge=1 posts nothing; rebuildBalance re-stamps
  absolute SUM; DB restored as found.
- **⚠️ GO/NO-GO reconciliation — BLOCKED by the snapshot, NOT a rebuild bug.** The restored `Inventory.bak`
  is **post-`DELETE_AutoPurge`**: **0** of 4238 open orders carry a counting status (all
  `VC_STATUS_SUPPLIER_SHIPPING`/`VC_ARRIVAL` blank), and all 16 suppliers are add-point **'S'** (no 'A').
  The receiving IN-movements that BUILT the legacy `IN_QTY` have been aged out (§3.1 purge horizon). A
  from-zero replay of this snapshot therefore CANNOT reconstruct `IN_QTY` — 36/47 parts diff negative
  (derived<legacy = purged receipts), 11 reconcile exact, **0 UNEXPLAINED** (no derived>legacy). The §3
  derivation logic is exercised + correct; the data is the limit. Of the 4 predicted FIX-divergence
  classes, only **F3** is even reachable here (1 shipping multi-row group), and its part's receiving
  history is purged too, so it can't show the over-count signature in this snapshot. **D8(3)/D12#3 need
  'A' add-point (none); F5 needs a re-pointed part-number (not in a static snapshot).**
- **What a real GO/NO-GO needs:** a **pre-purge event dump** (complete `INV_OPEN_ORDER_INF` history) or
  the **live cutover backfill window** — the harness IS the backfill validator (design §9). Harness +
  TSV (`scripts/e2e/artifacts/stock_ledger_parity.tsv`) ready to re-run against that data unchanged.
- **Verdict:** the SERVICE foundation is GO (mechanics proven). The PARITY adjudication is **deferred to
  pre-purge/cutover data** — green-on-post-purge-data would be a fixture-fidelity false pass, so the
  harness SKIPs (not PASSes) that check and reports the horizon explicitly.

## 2026-06-19 — FULL CUTOVER DRESS-REHEARSAL on the spike (backup→flip→restore) — GO

Ran the entire 4-phase cutover sequence (`docs/analysis/cutover-architecture.md`) end-to-end against the
spike, fenced by a pre-backup + post-restore. Full write-up: `docs/analysis/cutover-dress-rehearsal.md`.

- **GO/NO-GO ZERO-DRIFT GATE: 0 / 47 parts drift** after dropping the 13 qty-triggers and running
  `SEED_AllOpeningBalances` (47 OPENING_BALANCE rows; IN_QTY checksum unchanged by the seed). `IN_QTY ==
  SUM(ledger)` for every part. **GO.**
- **Phase A/B applied idempotently**: fn_ManifestCostAt → D6 procs (all 4 now CROSS APPLY the TVF) →
  pre-drop overlap diagnostic = **0** → no-overlap trigger created (IX drop no-op'd: already absent on
  spike) → UPDATE_PartsStockInfo qty-leg dropped + UPDATE_PartsStockInfoCount retired.
- **Phase C**: exactly **13** triggers dropped, `UPDATE_PartNumber`/`DELETE_PartNumber` KEPT. Forward-post
  smoke (REAL receiving +60 / shipping −50 wrappers via jython_shim, triggers already gone) → each moved
  IN_QTY by its exact delta ONCE + wrote its ledger row + invariant held (seams are the SOLE writer, no
  double-count). **Genesis guard** THROWs `Msg 50001` on a part with a forward row + no opening row.
- **Restore as-found PROVEN**: IN_QTY per-part `diff`-identical to baseline; 13 triggers back + enabled;
  ledger back to 0 rows; all 7 touched proc/fn **body MD5 hashes byte-identical** to baseline; gateway
  re-connected cleanly (handled via SINGLE_USER WITH ROLLBACK IMMEDIATE + same-batch MULTI_USER).
- **Deviations (honest):** (1) spike already had Phase-A objects + the manifest constraint already dropped,
  so section B no-op'd — the constraint-DROP path is NOT exercised here; **confirm on prod**. (2) A
  self-inflicted smoke-script teardown/genesis-probe bug (corrected) — NOT a cutover defect; RESTORE is the
  authoritative reset. No trigger-name mismatch, no ordering/lock problem.
- **Verdict:** the cutover sequence is executable and the zero-drift gate is GREEN on the spike. Backup
  `.bak` left in the container as the proven restore point.

## 2026-06-19 — M1 build #1: ASN-creation fan-out (pure) + Q1 re-keyed ASN-detail procs

Built the fully-testable-now pieces of the M1 ASN-creation keystone. The gateway `create_asn` driver +
live end-to-end parity remain **DEFERRED** to the live VehicleOrder backup (AD_FRSPull's GALC tables —
Vehicle/Model/vehicledata/DataItem — are NOT on the spike; spike VehicleOrder is a Line-only stub).

- **PART A — pure fan-out** `docs/analysis/edi/project-library/asn/code.py` (`computeAsnDetails`): the
  producer-recipe reimplementation of the hand-written Delphi `CalculateASNFRS` (DataModule.pas:5180-5268).
  No-Ratio branch (`Orders<=5`: one row from the FIRST fc row, `qty=Vehicles*IN_ASSY_QTY`, break) + ratio
  branch (`round(Vehicles*IN_ASSY_QTY*tire/100)`, both-100→full qty) + manifest `'7'+1-digit-year+MM+DD+id`
  + the two aborts (no fc rows; NULL manifest cost). **Banker's rounding (round-half-to-even)** implemented
  explicitly in exact integer math — Jython-2.7's `round()` is half-away-from-zero, so the built-in can't
  be trusted; this is the known parity trap.
- **PART B — unit test** `scripts/e2e/test_asn_fanout.py`: **34 PASS / 0 FAIL** (CPython, no DB). Covers
  ground/spare ratio splits, the No-Ratio single-row branch, ratio=100 full qty, the .5 banker's boundary
  (2.5→2, 3.5→4, 4.5→4), manifest `76061857` shape, multi-BC mix, and both aborts.
- **PART C — Q1 re-key** `docs/analysis/edi/spike-asndetail-rekey.sql`: re-keys `INSERT_ASNDetail`'s
  upsert existence-check AND `DELETE_ASNItem` to `(IN_ASN_ID, manifest)`; `DELETE_ASNItem` now takes
  `@ASNID`+`@ManifestNumber`. site_id intent for `(site_id, IN_ASN_ID, manifest)` left as explicit `-- M4:`
  markers (site_id not on the table yet). Added `IX_INV_ASN_DETAIL_MST_ASN_MANIFEST`. **APPLIED + VERIFIED
  on the spike**: accumulate within one ASN (10+15→25); same manifest in a DIFFERENT ASN gets its own row
  (no collision, 7+3→10); @HotCall=1 still always-inserts; `DELETE_ASNItem` scoped to one ASN (other ASN's
  same-manifest row survives). All test ASN headers/details swept. **Left the re-keyed procs + index in
  place** (cutover-correct; no other spike test/seam/library references either proc — verified by grep).

## 2026-06-19 — M1 build #2: `create_asn` gateway driver + live end-to-end parity

The keystone seam now runs end-to-end. The DEFERRAL from build #1 is **lifted** — the spike now holds
the REAL `VehicleOrder` (2.3M-row Vehicle/Model/VehicleData/DataItem/Line + `AD_FRSPULL`), a matched
LEGACY snapshot `Inventory_Live` (max ASN 4722; the daily-log COROLLA ASN 4721, P:20260618), and the
working rebuild `Inventory` (re-keyed ASN-detail procs applied).

- **PART D — driver** `docs/analysis/edi/project-library/asn/code.py` (`create_asn`, alongside the pure
  `computeAsnDetails`, the order/code.py pattern): SELECT_ASNSeq idempotency guard → `AD_FRSPULL` on the
  **VehicleOrder** datasource (READ) → per-BC `SELECT_ForecastDetailBCASN` on Inventory (READ) → pure
  fan-out (aborts pre-transaction on missing cost / no-fc) → **ONE Inventory transaction**: `INSERT_ASNInfo`
  (status 'C', OUTPUT @ASNID captured via `runScalarPrepQuery(...,tx)`, **EIN=0 — allocated at SEND**, the
  intended divergence, spec §2/§8) + per-detail `INSERT_ASNDetail` (re-keyed accumulate upsert); commit /
  rollback+raise. Post-loop `SELECT_ASNMissingCost` audit = WARN, never aborts (faithful to the Delphi
  pre-loop-abort vs post-loop-warn split). `@Start/@Last` passed to AD_FRSPULL but unused (verified).
- **Shim extension** `scripts/e2e/jython_shim.py` (minimal, backward-compatible): logical-DB routing map
  (`VehicleOrder` → cross-DB read for AD_FRSPULL; everything else → Inventory, unchanged); `runScalarPrepQuery`
  (real 8.1+ API) on both autocommit and the persistent-tx session (captures the OUTPUT @id inside the
  BEGIN TRAN — `createSProcCall` can't join a gateway tx); `beginTransaction` resolves the logical name;
  fixed a latent `getLogger` bug (`_System._Util` was a method-local, never callable). All prior seam/order
  tests still **green** (seam_driver 23, seam_driver_order 13, order_commit_integration 6, asn_fanout 34).
- **PART E — end-to-end parity** `scripts/e2e/test_create_asn_parity.py`: **9 PASS / 0 FAIL.** Runs the
  REAL `create_asn` in Inventory for COROLLA/20260618 over ASN 4721's window, then:
  - **Driver-correctness (the PASS gate): 17/17 persisted detail rows == `computeAsnDetails` over the
    spike's OWN `AD_FRSPULL`+forecast** — proves the wiring (cross-DB read, banker's round, accumulate
    upsert, header+OUTPUT capture, single transaction) is faithful. Header status 'C', EIN 0, qty mirrors
    the Check count.
  - **Legacy parity vs frozen ASN 4721 (informational): 6/17 manifests reproduce the REAL legacy rows
    EXACTLY** — all the small/stable BCs (NDD/NFF/NGG/NHH ground veh 2/4/14/8 + spare NP) plus 805/826/828/829.
    The 10 large-BC mismatches are **GALC fixture drift, NOT a driver fault**: the spike `VehicleOrder` is a
    full historical RELOAD, not the point-in-time build snapshot that froze 4721 — the implied per-row vehicle
    counts in 4721 are mutually inconsistent under current forecast ratios (e.g. NBB's three rows imply
    50/1125/702 vehicles), and the No-Ratio BC set shifted (4721's No-Ratio manifest 76061836 vs the spike's
    PEE/PN). Forecast detail is IDENTICAL between Inventory and Inventory_Live (ratios are not the cause).
    Reported, **not forced green** (fixture-fidelity discipline).
  - Spike restored as-found (test ASN header+details swept; max ASN back to 4715; Inventory_Live & VehicleOrder
    never written; re-keyed procs retained).
- **GO/STAY impact:** the revenue-keystone create chain is now a working, headlessly-driven Ignition seam
  with a clean self-consistency proof and a real (explained) legacy diff. Open item for FULL legacy row-for-row
  parity: a point-in-time GALC snapshot contemporaneous with ASN 4721 (the reload can't reproduce frozen
  build composition). Not a blocker for the driver — the seam is correct against its inputs.

## 2026-06-20 — M1 keystone: two SQL-review MUST/SHOULD-FIXes on `create_asn` (BLOCKER closed)

Closed the two SQL-review fixes on the ASN-creation keystone (architecture doc §6, adversary-findings).
Branch `m1-asn-creation`.

- **FIX 1 (BLOCKER) — concurrency double-insert.** `create_asn`'s `SELECT_ASNSeq` guard is non-atomic;
  two concurrent gateway sessions could both pass it and both commit a full ASN (a risk the gateway
  CREATES vs the single-user Delphi desktop). **Real-data cardinality check FIRST** (the gating
  question): on `Inventory_Live`, normal ASNs (`VC_START_SEQ_NUMBER <> -1`) have **ZERO** duplicate
  `(VC_LINE_NAME, VC_PRODUCTION_DATE)` groups (`GROUP BY … HAVING COUNT(*)>1` = 0); including hot-calls
  there are 206 dup groups — i.e. (line, prodDate) IS unique for normal ASNs and the `-1` exclusion is
  exactly what hot-calls legitimately repeat on. **So a filtered unique index is the correct key, not
  wrong.** Artifact `docs/analysis/edi/spike-asn-unique-guard.sql` = filtered `UNIQUE` index
  `UX_INV_ASN_MST_LINE_PDATE_NORMAL (VC_LINE_NAME, VC_PRODUCTION_DATE) WHERE VC_START_SEQ_NUMBER <> '-1'`
  (idempotent; `-- M4:` marker to lead with `IN_SITE_ID`). **APPLIED + VERIFIED on Inventory**: CREATE
  succeeded against all 2550 live rows (no existing violation); a 2nd NORMAL insert for an existing
  (line,prodDate) → **rejected, Msg 2601**; a hot-call (`'-1'`) insert for the same key → **allowed**
  (and a 2nd hot-call too). `create_asn` is now **race-safe**: the header insert catches the
  unique-violation (`_isUniqueViolation`, matches the index name / SQL 2601 + "duplicate key") and
  returns the idempotent `{'skipped': True}` instead of erroring; any other tx error still rolls back +
  re-raises.
- **FIX 2 (SHOULD-FIX) — No-Ratio nondeterministic pick.** `SELECT_ForecastDetailBCASN` has no ORDER BY
  over two HEAPs; the No-Ratio branch took `fcRows[0]` = luck-of-allocation (FIRES on live BC `PEE`:
  id 189 `42600FEL1000`/m36 vs id 190 `42600FEL2000`/m37). Did **not** ALTER the shared legacy proc;
  instead the REBUILD driver sorts each BC's rows by `ID_FORECAST_DETAIL` ascending (the table's
  `IDENTITY(1,1)` PK = first-configured) before the fan-out, and `computeAsnDetails` documents the
  caller-side order CONTRACT. ⚠️ **DAVID-CONFIRM flagged in code + return:** deterministic = LOWEST
  `ID_FORECAST_DETAIL` wins (a domain choice; strictly better than nondeterministic either way).
- **Harness fidelity finding (worth keeping):** once a FILTERED index exists on a table, **every DML
  against it requires `SET QUOTED_IDENTIFIER ON` + `ANSI_NULLS ON`** (else **Msg 1934**). The gateway's
  JDBC connection already runs these ON (production unaffected), but `sqlcmd -Q` defaults
  QUOTED_IDENTIFIER **OFF** — so the headless shim + the parity test had to prepend the SET options to
  faithfully mirror the gateway. Also hardened `jython_shim._TxSession` to recognize `SqlState NNNNN`
  error lines (not just `Msg NNNN`) so a unique-violation inside a tx batch — and the transient
  `SqlState 24000 Invalid cursor state` it leaves on the next batch — surface as a catchable error
  instead of hanging; rollback now swallows that follow-on.
- **Tests:** `scripts/e2e/test_asn_fixes.py` NEW (**22 PASS**: index reject/allow, the race-catch via a
  guard-blinded `create_asn`, `_isUniqueViolation` specificity, the deterministic PEE pick + end-to-end
  persisted m36→FEL1000). Regression green: `test_asn_fanout` **34**, `test_create_asn_parity` **10**,
  seam_driver **23**, seam_driver_order **13**, order_commit_integration **6**.
- **Spike restored as-found:** Inventory back to 2550 rows (315 hot-calls; all test ASNs swept);
  `Inventory_Live` (2557) + `VehicleOrder` (2.33M) never written; the unique index retained, matching
  the existing convention for cutover schema artifacts (the rekey procs/index are likewise left applied).
- **GO/STAY impact:** the keystone's BLOCKER (concurrency double-insert) is **CLOSED** at the DB level +
  the driver; the No-Ratio nondeterminism is removed (pending David's tiebreak confirm). Remaining
  keystone gate is unchanged: a true row-parity oracle (recipe-vintage drift, §7.5).
