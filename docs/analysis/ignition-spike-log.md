# Ignition Vertical-Slice Spike — Running Log

*Live record of the GO/STAY spike defined in [`ignition-spike-plan.md`](ignition-spike-plan.md). Append
findings as they happen; this is the artifact the GO/STAY decision is written from.*

Dev env: Ignition **8.1.52** (dev ceiling; prod target 8.3 — see
[`ignition-version-strategy.md`](ignition-version-strategy.md)). DB sandbox: Colima + SQL Server 2019.

> **Project rename (2026-06-23):** the Ignition project was renamed `spike` → **`InventorySystem`**
> (display title "Inventory System"). Gateway folder is now
> `data/projects/InventorySystem/`; the generators' and e2e harness' gateway-path constants were
> updated to match. References to `data/projects/spike/…` **below** are kept as accurate spike-phase
> history (the path was `spike` at the time). The `Inventory_Spike` DB connection, the `mssql-spike`
> container, and the `SPIKE` logger name are unchanged.

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

## 2026-06-20 — M1 Rank-2: outbound EDI 856 ASN builder (pure + feed SQL + driver) — BUILT
- **What:** new gateway code porting `EDI856Object.pas` (T856EDI), byte-faithful to the TEMA-accepted
  legacy 856 (David locked byte parity). Three artifacts:
  `docs/analysis/edi/856/project-library/edi856/code.py` = PURE `build_856(...)` (ordered segment list)
  + DRIVER `send_856(...)`; `docs/analysis/edi/856/spike-edi856-feed.sql` = the side-effect-free feed
  SELECT (the driver inlines it as `_FEED_SQL`). Mirrors the asn keystone pattern (pure logic + thin
  driver + jython_shim headless seam).
- **Locked decisions realized:** A=KEEP cost INNER join; B=forecast INNER via `CROSS APPLY (SELECT TOP 1)`
  (same INNER drop, no fan-out); C=GROUP-BY collapse (no sum); D=CRLF-only (segment terminator read but
  NEVER emitted); E=filename `856<copy(prodDate,4,5)>.txt`. EIN per-site from `INV_SITES.IN_EIN_SEQ`
  allocated **at send** (atomic `UPDATE … OUTPUT INSERTED.IN_EIN_SEQ WHERE IN_SITE_ID=?`), `%09d` in all
  7 control positions + BSN02=`yyyymmdd`+`%09d`; NOT the `6440` literal; NOT EIN-at-create. Status flip
  per-ASN `C→S` at send, DECOUPLED — never `REPORT_EDI856`'s self-flip, never the blanket
  `UPDATE_ASNStatus`.
- **Tests:** `scripts/e2e/test_edi856_build.py` (pure, **46 PASS** — full byte-exact segment list for a
  multi-manifest/multi-item fixture, EIN in 7 positions, control-number consistency, SE01==actual ST..SE
  count, CTT01==HL count, ISA widths + ISA09 yymmdd, S→O→I chain, PRF/LIN/SN1, CRLF, filename, Trap-6
  truncation). `scripts/e2e/test_edi856_e2e.py` (**27 PASS** — real `send_856` via the shim on a copy of
  Inventory_Live ASN 4721: feed-row parity vs the legacy `REPORT_EDI856 @EIN=0` SELECT 16/16; byte-exact
  STRUCTURE of the emitted file; EIN from `INV_SITES.IN_EIN_SEQ` 9100→9101 stamped on header; per-ASN flip
  with a decoy 'C' ASN left untouched). Regression green: asn_fanout 34, create_asn_parity 10,
  order_commit_integration 6, seam_driver_order 13, receiving_writepost 12.
- **Honest verification (fixture-fidelity):** NO golden 856 exists → byte-for-byte legacy parity is
  UNPROVABLE and NOT claimed. PROVEN: feed-row parity vs the legacy SELECT + byte-exact STRUCTURE +
  self-consistency + EIN provenance + the decoupled flip. Caveat: the spike `Inventory.INV_SITES` holds
  PLACEHOLDER site identity (DUNS `000000001`, supplier `MAS`, EDI mode `PROD`) — the real legacy values
  live in `VehicleOrder.Site` (DUNS `969009112`, supplier `71930`, TMM DUNS `808369495`, EDI mode `P`,
  sub-elem `#`), relocated to INV_SITES at the sites-CRUD step but not yet loaded with real values — so
  the ISA/GS BYTES emitted on the spike reflect placeholders, not the TEMA wire. Structure assertions are
  site-value-independent; exact wire bytes await the real INV_SITES values at cutover.
- **Byte-exact trap hit:** the driver opens a tx, stamps the EIN on the ASN header (X-lock), then must
  READ that same header (prodDate + feed). An autocommit read on a separate connection BLOCKS on the
  uncommitted row lock → hang. Fix: route the in-tx reads through `runPrepQuery(..., tx)` (8.1+ accepts a
  tx arg) so they share the connection. Hardened `jython_shim._DB.runPrepQuery` to honor a `tx` (route to
  the persistent session); added minimal `system.file.writeFile` + `system.date` shims. All additive /
  backward-compatible (regressions unchanged).
- **Spike restored as-found:** Inventory 2550 ASNs (test+decoy swept), `INV_SITES.IN_EIN_SEQ` back to 0,
  Inventory_Live + VehicleOrder never written.
- **GO/STAY impact:** the second M1 keystone (outbound 856) builds clean on the spike with the decoupled
  EIN/flip model — no new blocker. Open item for cutover: load the real site identity into INV_SITES
  before claiming wire-byte parity (a data task, not a code one); a golden 856 would upgrade the e2e from
  structure-parity to byte-parity.

---

## M1 Rank-4 — outbound EDI 810 INVOICE builder (2026-06-20) ✅

- **What:** new gateway code (NOT a wrap of the self-flipping `REPORT_EDI810`), byte-faithful to the
  legacy 810 wire format (`EDI810Object.pas`, the LIVE `T810EDI`) with **clean/correct money** (David
  locked) + **D6 window-aware pricing** + **Carry-5 in-place unsend**. Twin of the just-shipped 856.
- **Where:**
  - PURE builder `docs/analysis/edi/810/project-library/edi810/code.py` → `build_810(invoiceRows, site,
    ein, fileTime)` returns the ordered segment list (CRLF-joined by the driver). 3 drivers in the same
    module: `create_invoice` (daily invoicing), `recreate_invoice` (re-emit, EIN reused), `unsend_invoice`
    (Carry-5 in-place). Each driver = ONE Inventory transaction + `.tmp`→rename-after-commit atomicity.
  - Feed SQL `docs/analysis/edi/810/spike-edi810-feed.sql` — both REPORT_EDI810 branches (`@EIN=0` create
    / `@EIN<>0` recreate), `CROSS APPLY fn_ManifestCostAt`, **side-effect-free (NO self-flip)**. Canonical
    copy of the two inline `_*_FEED_SQL` strings; a drift guard in the e2e asserts byte-identity.
  - Tests `scripts/e2e/test_edi810_build.py` (pure) + `scripts/e2e/test_edi810_e2e.py` (headless).
- **Segment fidelity (vs `EDI810Object.pas`, the 856 lesson — expected bytes derived FROM the .pas):**
  GS01=`IN` (856 uses `SH`); BIG02=SiteSupplierCode; IT101 `M391` if manifest starts `'7'` else `M390`;
  IT1 ends on the dock code (the audit's "CLEAN — zero trailing seps" VERIFIED line-by-line — no `**`,
  no trailing sep anywhere); interior REF=`MK`+PREVIOUS manifest / DTM=`050`+pickup, trailing REF=FINAL
  manifest; dates off **NOW** (ISA09 yymmdd, GS04/BIG01 yyyymmdd) not the pickup date; EIN `%09d` in all
  7 control positions; SE01/CTT01 COUNTED from the emitted segments.
- **CLEAN MONEY (David locked — legacy bugs DELIBERATELY not reproduced):**
  - TDS01 = `round(total*10000)` integer, implied-decimal scale-4. Legacy hand-rolled this WITH BUGS: a
    1-digit fraction was NOT padded (`1234.5` → `12345`, an implied `1.2345` = **off by 10000×**) and a
    whole-dollar total emitted a malformed string. The test asserts the CORRECT values (incl. `1234.5` →
    `12345000`) AND that the off-by-10000 value (`22924` for the fixture total `2292.40`) is ABSENT.
  - IT104 = a fixed scale-4 decimal (`12.5000`), locale-independent — NOT `FloatToStr` (locale separator
    + variable fraction). Exact IT104 scale (4 vs trimmed) **flagged pending a golden 810**.
- **EIN:** allocated per-site from `INV_SITES.IN_EIN_SEQ` at invoice-CREATE (the SAME shared sequence the
  856 uses — faithful, interleaved control numbers), **reused** at recreate (no re-alloc, so the 997-ack
  `UPDATE_EINStatus WHERE IN_INV_EIN=@EIN` still lands). NEVER the self-flip, NEVER `UPDATE_INVRecreate`.
- **Carry-5 unsend (in-place):** reverts `INV_INV_MST.VC_INV_STATUS` to `'C'` (the unsent/recreate state)
  + re-pools detail (`IN_INV_ID=null`), **keeping the header + EIN + audit**. NOT the legacy
  `UPDATE_INVUnsend` HARD-DELETE (whose commented-out line shows the original intent was this status
  revert).
- **Tests:** `test_edi810_build.py` **61 PASS** (full byte-exact segment list vs the .pas derivation;
  GS01/BIG02; M391/M390; IT1 clean/no-trailing-sep; %09d EIN ×7; SE01==actual ST..SE; CTT01==#IT1; dates
  off NOW; CLEAN TDS incl. the 1-digit-fraction case + the legacy bug ABSENT; IT104 scale-4; empty-feed
  guard; single-item edge; ISA truncation). `test_edi810_e2e.py` **53 PASS** (real drivers via the shim:
  CREATE feed-row parity 3/3 vs the legacy `REPORT_EDI810 @EIN=0` SELECT on synthesized unbilled lines;
  EIN 9200→9201 from `INV_SITES.IN_EIN_SEQ`; `INSERT_INVInfo` header status 'S'; detail linked
  (`UPDATE_INVItems`); byte-exact file STRUCTURE + CLEAN money; RECREATE reuses the EIN, seq NOT bumped,
  byte-identical wire; ATOMICITY — commit-fault rolls back the EIN bump + header + link AND leaves no
  final 810 (only a swept `.tmp`); UNSEND in-place — header+EIN survive, detail re-pooled). Regressions
  green: asn_fanout 34, create_asn_parity 10, edi856_build 49, edi856_e2e 51, seam_driver 23,
  seam_driver_order 13, order_commit_integration 6.
- **Honest verification (fixture-fidelity):** NO golden 810 → byte-for-byte TEMA parity UNPROVABLE and
  NOT claimed. The spike `INV_SITES` holds PLACEHOLDER site identity (DUNS `000000001`, supplier `MAS`,
  dock `D01`, EDI mode `P`), so ISA/GS/BIG BYTES reflect placeholders, not the TEMA wire. PROVEN: feed
  STRUCTURE + parity vs the legacy SELECT + byte-exact STRUCTURE & self-consistency & CLEAN money on real
  data. Pending cutover: real INV_SITES values + a golden 810 (→ true byte-parity) + the exact IT104
  scale.
- **Spike restored as-found:** sentinel ASNs/detail swept; driver-created invoices swept by the
  test-owned EIN range (seeded 9200, above the snapshot max 9058); `INV_SITES.IN_EIN_SEQ` back to 0;
  counts back to baseline (INV_INV_MST 2934, INV_ASN_MST 2550, detail 39707); idempotent across reruns.
  Inventory_Live + VehicleOrder never written.
- **GO/STAY impact:** the third M1 keystone (outbound 810) builds clean — the clean-money + D6-pricing +
  in-place-unsend model holds on real data, no new blocker. Same two cutover open items as the 856 (real
  site identity + a golden file), plus: confirm the exact IT104 scale against a golden 810.

---

## 2026-06-21 — 810 CREATE-path adversary fixes: per-pickup-date file split (FIX 1 done; FIX 2 STOP)

Two CREATE-path divergences the sql-adversary found (`docs/analysis/edi/810/sql/adversary-findings-810.md`
SHOULD-FIX 1/2). Byte-faithful (reproduce-the-legacy) stance.

- **Legacy grouping CONFIRMED = per DISTINCT PickUpDate (NOT per-(date,line)).** `CreateINVOICEClick`
  (`MainMenu.pas:2613-2654`) walks the `REPORT_EDI810 @EIN=0` cursor (ordered `BY VC_MANIFEST_NUMBER`);
  each `EDI810.Execute` consumes one pickup-date run — the IT1 loop BREAKS only on a pickup-date change
  (`EDI810Object.pas:263-266`, the manifest break at :268 is an *interior* REF/DTM, not a file break) —
  then writes a SEPARATE file + allocates a NEW EIN (`SiteEIN+1`/`AD_UpdateEIN`) + `INSERT_INVInfo`. There
  is NO line dimension: the `@EIN=0` feed returns **6 columns** (Manifest/Part/Price/Qty/PickUp/ASNid, NO
  `LineName`) — verified in BOTH `CreateInventory.sql:3739-3745` (authoritative live dump) and
  `/tmp/inv_utf8.sql:3739-3745`. So: **N distinct pickup dates → N invoices / N sequential EINs / N files.**
- **FIX 1 DONE — `create_invoice` now splits per pickup date.** Reworked `code.py`: read the `@EIN=0` feed
  once → `_group_by_pickup_date` → per group `_create_one_invoice` (its own EIN + `INSERT_INVInfo` header +
  date-scoped detail link + `build_810` + `.tmp`→rename-after-commit), each its OWN transaction
  (commit-per-invoice, like the legacy per-`Execute` `InsertINVInfo`). Detail link is **date-scoped**
  (`UPDATE d … JOIN INV_ASN_MST a … WHERE d.IN_INV_ID IS NULL AND d.IN_ASN_ID=? AND a.VC_PRODUCTION_DATE=?`)
  — a correctness-preserving tightening of the legacy per-ASN `UPDATE_INVItems` (`VC_PRODUCTION_DATE` lives
  on the ASN header, one date per ASN, so the two coincide on real data). Return shape now
  `{invoiceCount, invoices[…], eins[…], rowCount}`. Atomicity/no-self-flip/per-invoice-status preserved.
- **FIX 2 STOP (escalated) — the create filename's `LineName` has NO source in the 810 create feed.**
  `MainMenu.pas:2623` interpolates `EDI810DataSet.FieldByName('LineName')`, but `REPORT_EDI810 @EIN=0`
  does NOT select a `LineName` column (proven above) and `EDI810DataSet` has no persistent field defs
  (`DataModule.dfm:460-467`) — so in Delphi/ADO that `FieldByName` would raise `EDatabaseError`. By
  contrast the **856** create feed DOES select `a.VC_LINE_NAME 'LineName'` (`CreateInventory.sql` REPORT_EDI856),
  and the **810 recreate** filename (`ASNInvoice.pas:872`) carries NO `LineName` — which the rebuild's
  `_filename_810` already matches. Cannot byte-faithfully add a `LineName` the create feed never produces;
  left `_filename_810` as `810<mmdd>.txt` pending delphi-architect / source confirmation of the `LineName`
  source (or a decision that the legacy create-with-LineName branch was dead/broken).
- **Tests:** `test_edi810_build.py` **61 PASS** (pure builder unchanged). `test_edi810_e2e.py` **72 PASS**
  (was 53; +new section (7) MULTI-DATE: 2 dates → 2 invoices / 2 distinct sequential EINs from
  `IN_EIN_SEQ` / 2 files `8100610.txt`+`8100611.txt`, each carrying ONLY its date's IT1 + DTM, detail-link
  date-scoped with NO cross-date leak; single-date still → exactly 1 invoice). Regressions green:
  asn_fanout 34, create_asn_parity 10, edi856_build 49, edi856_e2e 51, seam_driver 23, seam_driver_order
  13, order_commit_integration 6. Spike restored as-found (sentinel + test-EIN-range sweep; `IN_EIN_SEQ`
  back to 0; `unbilled_A`=0).

---

## M1 inbound (997/824) — the loop-closer for the 856/810 outbound (2026-06-21)

- **BUILT — the inbound 997/824 processor closes the EDI loop.** New gateway code, producer-recipe
  pattern (pure parsers + driver, the twin of `edi856`/`edi810`):
  `docs/analysis/edi/inbound/project-library/edi_inbound/code.py`. **GO unchanged / strengthened** — the
  inbound ack path runs end-to-end against the spike through the same headless shim as the outbound, so
  the full send→ack→reject cycle is now demonstrable without the Designer.
- **997 functional ack:** `parse_997` walks the AK1 loop (one ack per functional group, multi-group
  supported), reading the functional id (`copy(fcl,5,2)` chars 5-6) + EIN (`copy(fcl,8,9)` chars 8-16)
  by the legacy `EDIUpload.pas:194-195` offsets (the parity oracle) — but **fixes the legacy fragility**:
  it scans forward to the REAL `AK9`, **tolerating AK2/AK3/AK4** detail between AK1 and AK9 (legacy read
  char 5 of the *next line* blindly, `:196-197`). `ak9_to_status` maps **A/E/P/R distinct** + M/W/X/
  garbage→R (Q6 — a strict superset of the legacy binary A-vs-rest collapse). The driver wraps
  `UPDATE_EINStatus` in-line (SH→`INV_ASN_MST`, else→`INV_INV_MST`) so it can **CHECK @@ROWCOUNT** — an
  unknown EIN (0 rows) raises an **alarm row**, NOT a silent success (legacy bug fixed).
- **824 application advice:** `parse_824` reads each `NTE` by the legacy fixed offsets (errorText 9-58,
  manifest 60-67, part 69-80; `EDIUpload.pas:283-285`). Per **Q10** the driver flips the referenced ASN
  to `'R'` (matched by manifest on **`INV_ASN_DETAIL_MST` → parent `INV_ASN_MST`** — the manifest lives
  on the DETAIL row, confirmed via `sys.columns`; legacy flipped NOTHING → the 6 stuck-`S` invoices on
  Live) + writes one **alarm row per reject line** (manifest/part/errorText) for the main-screen alarm +
  click-to-detail.
- **Two NEW rebuild tables** (`docs/analysis/edi/inbound/spike-edi-inbound-tables.sql`, applied to the
  spike): `INV_EDI_INBOUND_LOG` (the **DB-tracked processed-files ledger** — kills the legacy re-ingest
  hazard; a re-drop of the same name+content-hash is a no-op) and `INV_EDI_ALARM_REJ` (the 824-reject /
  997-unknown-EIN **alarm records**, `BIT_RESOLVED=0` surfaced by the home hub; the native
  `system.alarm` wiring off these rows is the prod follow-on). Both carry `IN_SITE_ID` from day one.
- **Bugs deliberately NOT carried** (vs `EDIUpload.pas`): EIN-carryover (parse each file fresh), no-
  archive re-ingest (DB ledger is authoritative), the P12 #8 wrong-target retry, the no-@@ROWCOUNT
  silent-success. **Site-scoping** is a `-- M4` marker on the legacy-table UPDATEs (`INV_ASN_MST`/
  `INV_INV_MST` have no `IN_SITE_ID` yet); the resolved site (by ISA `delSL[4]` → `VC_TMM_DUNS`) is
  threaded so the M4 add is one line. **DUNS guard** (Q11): a no-match file is **quarantined**, not
  silently dropped (legacy `:439`).
- **Required M1 follow-on (status display):** the rebuild status-render CASEs/NQs must add `E`→
  "Accepted/errors" and `P`→"Partial" arms or those statuses render BLANK (same gap the legacy already
  has). No rebuild status-render NQ exists on disk yet to extend — recorded here as a required follow-on.
- **Tests:** `scripts/e2e/test_edi_inbound_build.py` **42 PASS** (pure: parse_997 multi-AK1 + AK2/3/4
  tolerance + A/E/P/R, parse_824 multi-NTE offsets, ak9_to_status, envelope helpers).
  `scripts/e2e/test_edi_inbound_e2e.py` **29 PASS** (headless via the shim, REAL driver: SH-accept→ASN
  `'A'` / IN-reject→INV `'R'`; unknown-EIN→@@ROWCOUNT-0 alarm; 824 manifest match→ASN `'R'` + per-line
  alarm detail; DUNS-guard quarantine; idempotent re-process; `process_inbound` poll). **Shim extended:**
  `jython_shim` `runPrepUpdate` now returns the genuine `@@ROWCOUNT` (was a constant `1`) so the driver's
  0-row alarm path is exercised for real — non-breaking, verified by the full regression suite.
- **Honest verification:** NO golden TEMA 997/824 exists — byte/offset parity vs a real file is NOT
  claimed. Fixtures are built to the legacy `EDIUpload.pas` copy() offsets (the oracle) + the outbound
  856 ISA shape (TEMA-as-sender inbound mirror). What IS proven: the status flow, the @@ROWCOUNT alarm,
  the 824 reject flag+detail, the DUNS quarantine, idempotency — all through the real driver + DB.
  **Pending:** confirm the AK1/NTE/ISA `delSL[4]` offsets against a real inbound sample (spec §2.3 #1,
  §3.1, §5.2).
- **Regressions green:** edi810_build 61, edi810_e2e 72, edi856_build 49, edi856_e2e 51, asn_fanout 34,
  create_asn_parity 10, seam_driver 23, seam_driver_order 13, order_commit_integration 6. **Spike
  restored as-found** (zero leftover rows; the two new tables empty as intended; `IN_EIN_SEQ` untouched).

---

## M2 foundational unit — EDI 830 forecast importer (2026-06-21) ✅

**What/where.** The 830 (DELFOR planning forecast) importer that populates the forecast M1's ASN-create
+ the Order read. PURE parse/explode/day-spread + a gateway driver, twin of the M1 inbound producer.
- STEP-0 extracted algorithm: `docs/analysis/edi/inbound/forecast-import-algorithm.md` (line-by-line from
  `ForecastBreakdownF.pas` + the proc bodies — the parse offsets, the explode int-div math, the
  day-spread, the delete-then-accumulate write order, the bugs-fixed table).
- PURE + DRIVER: `docs/analysis/edi/inbound/project-library/forecast/code.py`
  (`parse_830`/`pick_ratio`/`explode`/`day_spread`/`production_offset` pure; `import_830` +
  `process_forecast_dir` + `stale_sites` driver). **Reuses M1** `INV_EDI_INBOUND_LOG` (idempotency) +
  `INV_EDI_ALARM_REJ` (alarms) + the DUNS guard.
- DDL addendum: `docs/analysis/edi/inbound/spike-forecast-import-tables.sql` (one new column
  `INV_SITES.IN_HISTORICAL_FORECAST` = legacy `[INIT] HistoricalForecast`, default 12 weeks; the other
  Q11 columns — `VC_FORECAST_IMPORT_MODE`/`VC_LAST_FORECAST_IMPORT`/`BIT_USE_FIRST_PRODUCTION_DAY` —
  already existed on `INV_SITES` from the master-data sites work).

**Key invariants PROVEN end-to-end** (REAL `import_830` via the shim, against the spike):
- **D10 raw week:** `DO 2624 → stored IN_WEEK_NUMBER 24` (NOT 25). The FirstProductionDay offset (+1 for
  2026) is applied ONLY to the holiday-calendar lookup, never to the stored row (the R1 catch).
- **Delete-then-accumulate:** the additive `INSERTUPDATE_BreakdownForecastInfo` is a *replace* only
  because `DELETE_ForecastInfo` runs FIRST per assembly. Re-importing the SAME 830 → qty **NOT doubled**
  (still `[30,27,27,27,27,0,0]`), one row per component — the load-bearing M2 invariant.
- **Holiday day-spread off the REAL production calendar:** raw week 21 → lookupWeek 22 (offset +1); week
  22/COROLLA has an H on Mon → `[0,36,34,34,34,0,0]` (138 over 4 working days, remainder on the first
  working day). Proven against `AD_GetSpecialDateWeek` in `VehicleOrder` (cross-datasource RO read,
  autocommit — NOT in the import tx, since it's a separate datasource).
- **DUNS guard / idempotency:** no-match DUNS → QUARANTINED; a renamed re-drop → `SKIPPED_ALREADY_PROCESSED`
  (content-hash dedup), no double breakdown.

**Bugs FIXED (documented divergences — algorithm §F):**
- **D-Bug-1:** a missing BOM-ratio match writes a `830_FORECAST_GAP` alarm (visible) instead of the legacy
  **silent count drop** (`ForecastBreakdownF.pas:1178-1183`) — the root of the "Unable to get month
  forecast" order error.
- **D-Bug-2:** the usage rollup reads the SAME production-relative week the row is stored under (not the
  legacy raw-ISO `WeekOfTheYear(now)`).
- **D-Bug-3 (REVERSED 2026-06-21 — see the adversary-fix entry below):** the day-spread now reproduces the
  legacy running `days` counter EXACTLY (H/X → off + DEC; 'O' overtime → ON + INC, `:1398-1407`). The former
  "True only if previously off" divergence ignored 'O' and so diverged from the legacy on the live COROLLA
  O-Saturdays (BLOCKER-2). Parity wins.

**Correction logged:** `forecast-breakdown.md` flagged `DELETE_ForecastInfo` as "absent from the snapshot
→ latent runtime failure." It is **CONFIRMED PRESENT** at `CreateInventory.sql:2725` (snapshot drift;
authoritative dump is the 6/12 file). Updated §2/§3/§4.3/§7/§8-Q2 there.

**Tests:** `scripts/e2e/test_forecast_import_build.py` **35 PASS** (pure: parse_830 multi-LIN/FST +
D10 raw week, pick_ratio default/dated, explode int-div + 3-way share summing 100, day_spread with/without
holiday). `scripts/e2e/test_forecast_import_e2e.py` **24 PASS** (headless, REAL driver: raw replace +
breakdown explode, D10 raw week, delete-then-accumulate not-doubled, BOM-miss alarm, holiday spread off
the real calendar, DUNS quarantine, idempotency, Q11 last-import stamp + 8-day staleness alarm).

**Honest verification:** NO golden TEMA 830 on disk (the D10 sample is gitignored client data) → byte/
offset parity vs a real file is NOT claimed. Fixtures are built to the `ForecastBreakdownF.pas` copy()
offsets (the oracle). Pending a golden: the exact ISA `delSL[4]` element index + the real day-spread on a
captured feed (same stance as 856/810/997/824). Year-blind key (D-Bug-5) deferred to the M4 re-key.

**Regressions green:** forecast_import_build 35, forecast_import_e2e 24, edi_inbound_build 42,
edi_inbound_e2e 41, edi810_build 61, edi810_e2e 72, edi856_build 49, edi856_e2e 51, asn_fanout 34,
create_asn_parity 10, seam_driver 23, seam_driver_order 13, order_commit_integration 6. **Spike restored
as-found** (0 synthetic rows; real `4265202R6000` wk-24 breakdown untouched; site stamps NULL as found).

---

## M2 forecast importer — adversary BLOCKER/RISK fixes (2026-06-21) ✅

Closed the 2 BLOCKERs + 2 RISKs the sql-adversary/code-reviewer found
(`docs/analysis/edi/inbound/sql/adversary-findings-forecast.md`). All are **reproduce-the-legacy (parity)**
fixes against `ForecastBreakdownF.pas`. Files: `forecast/code.py`, both forecast tests, algorithm §D/§E/§F/§H.

- **BLOCKER-1 (per-component supplier).** Each `INV_BREAKDOWN_FC_INF` row is written/keyed/deleted under the
  **COMPONENT's part-master supplier** (`:1349` `supplier := FieldByName('Supplier Code')` from
  `SELECT_PartsStockInfo(@PartNum := the component)`), NOT the feed supplier. The feed supplier stays on the
  RAW `INV_FORECAST_INF` row only (`:1107`). Was `rowSup = rowSupplier or compSupplier` (every row → feed
  supplier: wrong `VC_SUPPLIER_CODE` + wrong additive-upsert key `(supplier, part, week)`; and the
  supplier-blind `DELETE_ForecastInfo` would silently rewrite supplier on re-import). Now `rowSup =
  compSupplier`. **Proven live:** breakdown = 14 distinct suppliers, 959/959 == component part-master
  supplier, 0/959 == feed. New e2e: two components → two DIFFERENT synthetic suppliers; the row carries the
  component's, the raw carries the feed's.
- **BLOCKER-2 (overtime 'O' day-spread).** `AD_GetSpecialDateWeek` 'Date Status': **H/X turn a day OFF, 'O'
  (overtime) turns a day ON** (a worked Saturday). `day_spread` now takes the FULL calendar rows and
  reproduces the legacy running `days` counter (`:1398-1407`): H/X → off + DEC, 'O'/non-H/X → ON + INC,
  incl. the latent already-on double-count. **Proven live:** 6 'O' Saturdays on COROLLA (ISO
  13/15/17/20/23/30); wk23 qty 138 → legacy `[23,23,23,23,23,23,0]` (days=6), was `[30,27,27,27,27,0,0]`
  (days=5). New PURE + e2e cases assert the O-spread + the parity edges.
- **RISK-1 (delete anchor).** Anchor = `min(weekDate)` across all entries (documented divergence from the
  legacy "last LIN's first-week-date", which is arbitrary). Equal on real data; strictly stronger against
  the additive-doubling risk the anchor guards (the delete window covers every week we re-insert for any LIN
  ordering / start-week mix).
- **RISK-2 (per-site config wired).** `_process_one_830` now reads `IN_HISTORICAL_FORECAST` +
  `BIT_USE_FIRST_PRODUCTION_DAY` from `INV_SITES` (`_read_site_forecast_config`, mirrors `supplierBySite`)
  and threads them into `import_830` — the columns the M2 schema added are LIVE, not dead. e2e case 7 pins
  `BIT_USE_FIRST_PRODUCTION_DAY=0` for site 1 → offset 0 → lookupWeek 30 = COROLLA O-Saturday →
  `[19,16,16,16,16,16,0]` (and restores the flag as-found).

**Tests:** forecast_import_build **43 PASS** (+8: O-Saturday spread, X==H, mixed H+O week, the two running-
counter parity edges), forecast_import_e2e **35 PASS** (+11: per-component supplier B1, O-Saturday 5b,
RISK-2 poll honors the per-site flag). **Regressions all green** (edi_inbound_build 42, edi_inbound_e2e 41,
edi810_build 61, edi810_e2e 72, edi856_build 49, edi856_e2e 51, asn_fanout 34, create_asn_parity 10,
seam_driver 23, seam_driver_order 13, order_commit 16, order_commit_integration 6). **Spike restored
as-found** (0 synthetic rows + suppliers; site config + stamps unchanged).

**Honest verification:** still NO golden TEMA 830 on disk — the supplier/'O'-day/anchor fixes are proven
against the legacy `.pas` algorithm + the live proc bodies + live `Inventory`/`VehicleOrder` data, NOT a
golden file. The ISA `delSL[4]` element index remains the standing pending-golden caveat.

---

## M2 unit-2 — Order-FILE generator (`.ord` emitter) BUILT + e2e-proven (2026-06-21)

New gateway producer: serializes committed, renban-assigned, not-yet-ordered open orders to the
sub-supplier `.ord` files. SEPARATE unit from the Order COMMIT path
(`project-library/order/code.py`) — shares no code, does no order math. Source:
`OrderFormCreateF.pas:55-707` + the 5 procs (proved on `mssql-spike`). Files:
`docs/analysis/order/project-library/order_file/{code.py,excel.py,resource.json}`,
`docs/analysis/order/spike-order-file-feed.sql`, `scripts/e2e/test_order_file_{build,e2e}.py`.

**`.ord` format (byte-faithful to OrderFormCreateF.pas:556-585):** one positional, delimiter-less line
per row = `VC_SUPPLIER_CODE`(raw) + `VC_FRS_NUMBER`(raw) + `%8s`(renban, right-just min-8) +
`VC_PART_NUMBER`(raw) + `%05d`(IN_QTY, zero-pad min-5) + ship-date(`yyyymmdd`), CRLF after every line.
Only renban+qty are padded; supplier/FRS/part are RAW (legacy H3 — a short value shifts following
fields; reproduced faithfully, NOT silently padded). The phantom `SiteSupplierCode` leading field
(legacy H4, 14/16 suppliers, BIT_SITE_NUMBER_IN_ORDER=1) is a LATENT CRASH like the 810 — delphi-
architect adjudicated this session: built the working **non-sendsite 6-field line for ALL suppliers**,
did NOT reproduce the crash, did NOT invent a leading field. The M4 multi-site leading field =
`INV_SITES.VC_SUPPLIER_CODE` (marked `_M4`, width pending golden).

**3-destination fan-out + skip-logistics:** `[INIT] LocalFTP` (default False) = supplier dir ONLY;
True = supplier + logistics + `<supplierDir>/Archive`. Logistics resolved per-supplier via the
part→supplier→`'NONE'` ladder (`SELECT_PartsStockLogistics` then `SELECT_SupplierInfo`); the `'NONE'`
string sentinel SKIPS the logistics copy (the three states `'NONE'`/`''`/NULL are distinct — `''` slips
past the guard, legacy H8, flagged). All 3 copies are byte-identical (proven).

**Bugs FIXED vs legacy (all e2e-proven):**
- **H1 atomicity** (files-not-transactional): legacy writes `.ord` MID-tx, a rollback leaves the final
  file → duplicate re-emit. FIXED with the 856 temp-then-rename idiom: write every `.ord` to `.tmp`, run
  the ONE stamp tx, rename to final ONLY after commit; on any failure delete every `.tmp`. Proven: a
  stamp-tx failure after the `.tmp` write leaves NO final `.ord`, no orphan `.tmp`, rows un-stamped
  (a re-run re-emits).
- **H10 column aliasing**: legacy `SELECT *` (88 cols) duplicates `IN_QTY`/part/supplier/kanban;
  `fieldbyname` grabs the FIRST. The aliased `_FEED_SQL` (canonical `spike-order-file-feed.sql`) reads
  every column from `i.` (open-order) — the emitted qty is the ORDER `IN_QTY` (1200), NOT the
  parts-stock on-hand (proven 1200 emitted, on-hand 0 absent).
- **H11 emit-then-stamp**: `UPDATE_ORDEROrderDate` keys on part+FRS ONLY (no renban filter) → one stamp
  marks all same part+FRS rows. Emit from a single SNAPSHOT, then stamp the distinct (part,FRS) pairs
  once each — proven both same-part+FRS renban rows stamped, a 2nd run emits NOTHING (re-emit guard).

**GetShip CALENDAR-INCONSISTENCY FLAG (FOR REVIEW):** ship-date = `now + GetShip(lead)` where the
working-day scan (`compute_ship_offset`) skips ONLY weekends + `'H'` holidays — it does NOT skip `'X'`
(non-prod) or `'W'`, and treats `'O'` (overtime) as a normal day. This is UNLIKE the M2 forecast
day-spread (which accounts for O/X/W). Reproduced FAITHFULLY (weekend + 'H' only); flagged as a possible
carry-forward bug for delphi/ignition-architect to adjudicate at cutover. e2e proved the renban-group
ship-day override path: part in renban group 7 (Mon lead 13) → 18-calendar-day offset off the REAL
VehicleOrder holiday calendar.

**Excel order-forms (secondary):** data model built byte-faithful to the Pascal `Cells[r,c]` writes
(Excel order + Wheel/Tire order-sheet, page-fill at o>23, per-renban WHEEL workbooks, FRS→date cell);
`render_xlsx` is a STUB — the EXACT visual layout of the 3 `.xls` templates is PENDING the template
files (not confirmed on disk), like the forecast Excel reports. The `.ord` (EDI-critical) is fully
built + byte-tested.

**Tests:** order_file_build **44 PASS** (exact-byte `.ord`, %8s/%05d/raw-shift edges, GetShip scan,
LocalFTP fan-out, Excel data model), order_file_e2e **25 PASS** (REAL `generate_order_files` via the
shim's persistent-session tx: exact bytes, aliasing, 3-dest fan-out + NONE skip, emit-then-stamp,
H1 atomicity rollback). **All regressions green** (forecast 43/35, edi_inbound 42/41, edi810 61/72,
edi856 49/51, asn_fanout 34, create_asn_parity 10, seam 23/13, order_commit 16/6). **Spike restored
as-found** (4238 open orders, 0 not-ordered, 0 synthetic rows/hist, part on-hand unchanged).

**Honest verification / pending golden:** the `.ord` is byte-faithful to the `.pas`; PENDING a golden
`.ord` are (a) whether the receiving parser expects always-full-width supplier/FRS/part (the raw-concat
H3); (b) the M4 leading-field width; (c) the Excel template visuals. Ship date uses the REAL
AD_GetSpecialDate calendar.

---

## M2 renban breakdown (RenbanOrder.pas / GroupRenbanOrder_Form) — ✅ pure + driver built + verified

The MIDDLE order stage: splits blank-renban placeholder orders across N trailers, assigning each a
unique FRS# + renban, via DELETE-then-reINSERT. New gateway code: pure `compute_trailer_breakdown` +
driver `commit_renban_breakdown` in `docs/analysis/order/project-library/renban/code.py`. STEP-0
algorithm extraction appended to `renban-breakdown-spec.md` §12 (verbatim from RenbanOrder.pas:255-301
distribution + :746-799 read-out). Does NOT touch the GO/STAY picture — confirms the Order chain's
middle stage is rebuildable as pure-logic + a thin tx driver, same shape as Order commit / order-file.

- **Trailer distribution (`:255-301`):** per part `lots = order_qty div lotqty`; Phase A even base
  share `lots div T` (+ forward-spill of overflow); Phase B `lots mod T` remainder dribbled one-lot
  round-robin; merge repeated parts per truck (SUM lots — R3 sum-all, FAITHFUL). **Cross-checked vs an
  INDEPENDENT transcription of the .pas across 10,200 scenarios → 0 mismatches** (incl. the multi-part
  forward-spill else-branch). KEY fidelity catch: `TTruck.AddOrder` bumps `CurrentCount` INLINE
  (`:179`), so later same-part iterations see the running counts — a delta-then-add-later refactor
  diverged (1830/10200 mismatches); fixed by mutating the truck in place, matching the interleaved
  Pascal.
- **FRS suffix (`:763-767`):** `copy(frs,1,5)` + 2-digit trailer ordinal (`TruckNumber+1`, zero-pad
  <10, raw digits ≥10). **Renban (`:775-779`):** `groupCode + %.3d(seed3 + TruckNumber)`, seed3 = the
  3-digit tail of `groupCode||groupCount`; never blank.
- **FRS-SUFFIX NO-OP CONFIRMED (the assigned-task gate):** `INSERT_OpenOrder` re-derives the trailing
  2 FRS digits server-side, BUT `@FRSNum` is `varchar(7)` and we send the full 7-char FRS, so
  `@FRSNum + <suffix>` → 9 chars → silently truncated back to 7 = our value. RE-PROVED 2026-06-21 both
  branches: `'6090102'+'01'` → `'6090102'`, `'6090103'+(max+1)` → `'6090103'` (both len 7). The proc
  HONORS Pascal's `TruckNumber+1` suffix; no override → no STOP needed. (The recompute is only live for
  the original 5-char Order path.) E2E confirms persisted FRS `9123101/02/03` == compute output.
- **Driver (one tx):** per part `DELETE_OrderRenban(@FRS='',@Renban='')` (clears ALL still-blank rows
  of the part) → `INSERT_OpenOrder` per trailer-row (renban NON-blank, qty>0, skip qty=0) → once
  `UPDATE_RenbanGroupCount(next_count)`. Delete-then-reinsert ONLY (the commented-out update-in-place
  `:482-539` is NOT resurrected). Read uses an ALIASED SELECT (not the proc) to dodge the
  `SELECT_OrderNoRenban` duplicate-`IN_QTY` trap (o.IN_QTY order vs p.IN_QTY on-hand).
- **Hazards handled:** H7 div-by-zero (lotqty 0/NULL → raise, no crash); varchar(3) counter rollover
  at 999 (the persisted count = `('%03d' % next_count)[:3]` = legacy `str(N)[:3]` left-trunc — see the
  2026-06-21 adversary-fix entry below; the original `% 1000` was a BLOCKER, now fixed); Phase-B infinite
  loop bounded (raise vs hang); inverted lot-flag confirmed UPSTREAM-only (not read here). H1 (silent
  stuck blank-renban orders) noted as a follow-on alert (not built this pass).

**Tests (initial build; later raised to 35/27 by the 2026-06-21 adversary fixes below):**
`test_renban_build.py` **28 PASS** (hand-traced base case, FRS/renban derivation, merge
sum-all, skip-qty=0, capacity gate, div-zero guard, the 10,200-scenario independent-reference fuzz,
conservation). `test_renban_e2e.py` **16 PASS** (REAL `commit_renban_breakdown` via the shim's
persistent-session tx: seed 3 blank CMWA placeholders → delete-then-reinsert → 9 trailer rows all
NON-blank renban + blank order-date (emit-eligible), per-trailer FRS/renban/qty == compute, counter
288→291, STOCK-NEUTRAL on-hand unchanged, idempotent re-run no-op). **All regressions green**
(order_file 53/34, forecast 43/35, edi_inbound 42/41, edi810 61/72, edi856 49/51, asn_fanout 34,
create_asn_parity 10, seam 23/13, order_commit 16/6). **Spike restored as-found** (0 tagged rows/hist,
CMWA count back to 288, on-hand unchanged).

**Honest verification / pending golden:** faithful to RenbanOrder.pas; the distribution proven vs an
independent .pas transcription + the FRS-suffix no-op proven on the live proc. PENDING a golden renban
breakdown are the exact persisted FRS/renban cells for a real operator run (the spec's H8 names the
cells: a 3-trailer CMWA breakdown → `01/02/03` FRS + `CMWA000/001/002` renban) — the E2E reproduces
this shape on synthetic-but-real-keyed placeholders.

---

## M2 renban breakdown — sql-adversary parity fixes (2026-06-21) ✅

Closed the 2 divergences the sql-adversary found
(`docs/analysis/order/sql/adversary-findings-renban.md`); both reproduce-the-legacy (parity-first).
The distribution math / FRS suffix / in-run renban string were PROVEN faithful (could not be broken);
these two live in the counter-PERSISTENCE rollover and the write-back DELETE scope. Does NOT change the
GO/STAY picture — sharpens parallel-run fidelity at two previously-untested edges.

- **BLOCKER 1 — counter-rollover parity (`code.py` step (c)).** The persisted group count for
  `next_count >= 1000` was `'%03d' % (next_count % 1000)` (`1002 → '002'`). The legacy persists
  `Format('%.3d',[next_count])` = `'1002'` (min-width, never caps) into `@RenbanCount varchar(3)`, which
  the proc LEFT-TRUNCATES to the first 3 chars → `'100'`. Reduction is `str(N)[:3]`, NOT `% 1000`. FIX:
  `('%03d' % next_count)[:3]`. **PROVEN on the live proc** (mssql-spike, rolled back, CMWA restored 288):
  `@RenbanCount='1002' → '100'`; `@RenbanCount='002' → '002'` (the old-bug value); `'634'→'634'`,
  `'005'→'005'`. The renban NUMBER itself is unaffected (`CMWA1000` is 8 chars, fits varchar(8)). The
  rollover is itself a **latent legacy bug** (`1000→'100'` collides the next run's renban block with the
  earlier `CMWA100x`) — faithfully reproduced for parity, carried as a POST-CUTOVER fix (widen the count
  / alert+block at 999), tagged `# IG83-TODO:` in code + spec §12.7 (rollover-latent-bug carry).
- **SHOULD-FIX 2 — partial-lot delete scope (`code.py` `parts_seen`).** The delete set was built from
  the full loaded feed, so a part with `0 < qty < lotqty` (`lots=0`, emits no trailer row) had its
  blank placeholder DELETED with no re-insert → silent order loss. The legacy commit loop iterates only
  the EMITTED grid rows (`:417`/`fAvailableCount`=`:799`) and deletes by the emitted row's part
  (`:506`), so a `lots=0` part is NEVER deleted. FIX: derive `parts_seen` from the emitted `rows`
  (compute output), not `orders`. Verified live: a sub-lot part (qty 20 < lotqty 40) + a normal part →
  after commit the sub-lot placeholder SURVIVES with qty 20 intact, the normal part is grouped.

**Tests:** `test_renban_build.py` **35 PASS** (+7 rollover: next_count crosses 1000, persisted `'100'`,
regression-guard that the old `% 1000` would give `'002'`, the 8-char renban numbers, sub-1000
unchanged). `test_renban_e2e.py` **27 PASS** (+11: partial-lot survival incl. qty intact + normal-part
grouped + counter-from-emitted-only; rollover persists `'100'` in the DB + the DB-side CMWA1000/1001
renban numbers). **All regressions green** (order_file 53/34, forecast 43/35, edi_inbound 42/41,
edi810 61/72, edi856 49/51, asn_fanout 34, create_asn_parity 10, seam 23/13, order_commit 16/6).
**Spike restored as-found** (0 tagged rows/hist, CMWA 288, on-hand 13341 unchanged).

---

## M3 — server-side report-render harness + the 4 live reports (2026-06-21)

**Built:** the reusable POI render engine (`docs/analysis/reporting/project-library/report_render/`
`{code.py, report_defs.py, driver.py}`) + the **4 reports that actually ran in a year** (data-driven
prune — 18 of 22 never ran): `daily_shipping_assy` (R3, faithful), `invoice_summary` (R9, D6 corrected
default + faithful behind the seam), `forecast_detail` (R18, faithful), `lot_location` (R12, faithful).
One `.sql` per proc in `docs/analysis/reporting/sql/`; SKIP-list of the 18 dead reports in
`report_defs.py::SKIP_REPORTS` + `m3-render-architecture.md §10.1`.

**GO-relevant finding — POI render is HEADLESS-runnable on 8.1.52.** The render mechanism the architect
bet on (Apache POI XSSF Jython lib, §1.2) is **proven end-to-end headless** using the gateway's bundled
JRE + `jython-ia-2.7.3.3` + `poi-4.1.2` (no system Java, no running gateway):
`lib/runtime/jre-mac/bin/java -Dpython.path=user-lib/pylib -cp jython-ia.jar:lib/core/common/*
org.python.util.jython -S <script>`. This is the exact code path that runs on the gateway; the e2e reads
the `.xlsx` back with openpyxl. Strengthens GO: the entire report surface (data + XLSX numbers gate) is
build/test-able on the 8.1.52 spike with no Designer and no PDF/print dependency (PDF is the only
8.3-specific, non-gating surface).

**Lot Location confirmed LIVE (not the D9 NUMMI twin):** the live `LotLocationClick` calls
`REPORT_PLANTLotLocation[W]` (PLANT procs); `REPORT_NUMMILotLocation[W]` is the deprecated twin and is
NOT referenced. → BUILD, not SKIP.

**INVOICE Summary D6 (billing):** ships the window-aware corrected query (`fn_ManifestCostAt`, consistent
with locked 810/856 D6) as default; window-blind legacy behind the `QUERY_VARIANT` seam. Divergence proven
NON-VACUOUSLY (inject a 2nd gap window → legacy 34→42 over-bill, corrected stays 34; revert FAILS).

**Tests:** `scripts/e2e/test_m3_reports.py` **48 PASS / 0 FAIL** — Layer 1 numbers parity (every
expectation from the legacy proc / an independent truth query, never the rebuild), Layer 2 real-POI render
read back with openpyxl, output-stage (legacy `yyyymmddhhmmss00` filename + `.xlsx` + tmp-then-rename) +
revert non-vacuity. HONEST: numbers faithful to legacy procs on live data; the `.xls` template chrome
(borders/page-breaks/fonts) is NOT reproduced — same cells, not the template. **Regressions green**
(renban 35/27, order_file 53/34, forecast 43/35, edi_inbound 42/41, edi810 61/72, edi856 49/51,
asn_fanout 34, create_asn_parity 10, seam 23/13, order_commit 16/6). **Spike restored as-found**
(yard-status 0, no injected windows, no temp procs).

**Flagged (NOT M3): `test_report_procs_d6.py` is 9/2 on this spike** — pre-existing + independent: a
filtered unique index later added to `INV_ASN_MST` (`UX_INV_ASN_MST_LINE_PDATE_NORMAL`) needs
`QUOTED_IDENTIFIER ON` (Msg 1934), but that test's sqlcmd uses the `-Q` default OFF, so its EDI856 seed
INSERT silently fails → its fan-out check was vacuously green. Fixing the SET options un-masks a real
`legacy=2` vs `migrated=1` divergence in `REPORT_EDI856_D6`'s forecast fan-out. → for D6/sql-adversary.

---

## M4 piece 1 — Sites master + path columns (2026-06-22) ✅

**Direction reversal (David 2026-06-22):** each site runs on its OWN gateway + DB (single-site
deployments, NOT shared-DB multi-tenancy). So NO `site_id` surgery — `INV_SITES` holds the ONE site's
config per deployment (one row). Sites config stays in the DB (David's call). This piece adds the
per-deployment directory columns + the Sites config CRUD; the role-gate (Admin + production-control user)
lands in M4 piece 2.

**Path columns added (`spike-inv-sites-paths.sql`, idempotent ALTER addendum — M2 forecast-config style):**
7 varchar(260) columns relocating the legacy INI `[DIRECTORIES]` set (verified field-for-field in
`DataModule.dfm`): `VC_EDIOUT_DIR`←EDIOut, `VC_EDIIN_DIR`←EDIIn, `VC_FORECAST_DIR`←ForecastInputDir,
`VC_LOGISTICS_DIR`←LogisticsInputDir, `VC_REPORTS_DIR`←ReportsOutputDir, `VC_SHIPPING_DIR`←TextShippingFileDir,
`VC_TEMPLATE_DIR`←TemplateDir. (`LocalFTP` is an `[INIT]` boolean, not a dir → excluded.) EDIOut/EDIIn are
a SHARED EDI share for now (deployment-config fact, no schema flag); the rest are per-deployment. Placeholder
seeds only (NEVER real client paths). Applied to spike, re-run-safe.

**Sites CRUD:** the existing 8th-master combined view (`gen_sites_view.py`, route `/sites`) extended with
the 7 path fields + the load-bearing positional-ISA validation the source-truth (§4) required and the prior
build lacked: `VC_EDI_MODE` exactly 1 char, the 3 separators exactly 1 char each, DUNS 9-or-13-digit format.
NQ SQL (`master-crud-namedqueries.sql` Sites/get/insert/update) mirrored to carry the path cols. View
deserialized clean on `gwcmd -r` (gateway jvm 40, 2026/06/22 15:59:53; zero today-dated deserialize/error).

**Driver wiring (deliverable #3) confirmed:** the 856/810/order/forecast/report drivers read site config
from `INV_SITES` by COLUMN PROJECTION (`SELECT VC_SITE_ABBR, VC_DUNS, …`) keyed by `IN_SITE_ID`, and take
their output dir as a parameter today (test temp dirs). The path columns are now the source for those dirs
(deployment config). Adding columns is non-breaking — confirmed by the regression suites below.

**Tests:** `scripts/e2e/test_sites_master.py` (shim/DB, headless) **20 PASS / 0 FAIL** — DDL idempotency;
full CRUD round-trip running the IDENTICAL insert/update/get SQL the view emits (imported from the shared
builders in `gen_sites_view.py`, no drift), every field type incl. the 7 paths read back as set; validation
with an INDEPENDENT oracle (fill-days>50 / retention 1..11 rejected by the table CHECK; 1-char sep/EDI-mode
+ DUNS by the source-truth predicate). NON-VACUITY proven: a reverted rebuild that drops a path column
stores NULL → the round-trip assertion fails; and `VC_EDI_MODE` is varchar(10) so the DB SILENTLY stores a
2-char value (no backstop) → the save-path guard is genuinely load-bearing. **Regressions green** (edi856
52, edi810 72, forecast_import 35, order_file 34, m3_reports 50, edi_inbound 41) — the INV_SITES path-column
ALTER broke no 856/810/order/forecast/report/inbound consumer. **Spike restored as-found** (2 seed rows
MAS/HERO, all CRUD rolled back; the path columns are an intended permanent addition, left in place).

---

## M4 hardening piece-3 + P16 coverage (2026-06-22) ✅

The final M4 piece (security/multi-site hardening, single-site) + the P16 test-coverage follow-on.

- **HD1 secrets → gateway** (`m4-hardening-secrets.md`). The legacy holds SQL conn strings WITH passwords
  in the INI (`[DATABASE]`, `DataModule.pas:731-733`) / `DataModule.dfm`. The rebuild uses the **named
  gateway DB connection** (`Inventory_Spike`; cross-DB reads use `VehicleOrder`), creds encrypted gateway-
  side. **Grep-CONFIRMED: ZERO hardcoded JDBC/passwords/INI-conn-strings in any project-library driver**
  (incl. the new `auto_purge`) — every driver references the connection by logical NAME only.
- **HD7 backup runbook** (`m4-backup-runbook.md`): the two tiers (SQL `.bak` nightly+log / Ignition
  `.gwbk` nightly+pre-change), schedule/retention (≥12-mo to match the purge floor)/offsite, the paired
  restore drill, and the redundancy posture (Q14: invisible to the app — note, don't build HA).
- **HD4 DATAPURGE** — BUILT + verified. Driver
  `docs/analysis/production-readiness/project-library/auto_purge/code.py` faithfully reproduces
  `DELETE_AutoPurge` (`CreateInventory.sql:7682`) SCOPE: the 16-char `VC_ADD<=cutoff` predicate
  (style112+114 substrings) over the SAME 4 statements (terminate-then-delete `INV_OPEN_ORDER_INF`,
  delete `_HIST` + `INV_PARTS_STOCK_MST_HIST`), single-site = NO site filter (faithful), the `@DataRentention`
  NEGATIVE sign-flip, the ≥12 floor (`DataModule.pas:6890`). HARDENED: all 4 in ONE transaction (any
  failure → rollback ALL → nothing purged, closing the legacy partial-purge hole); retention + enable/prompt
  read from `INV_SITES` (Q17); schedulable via a gateway Timer/Scheduled script (8.1-safe).
  Test `scripts/e2e/test_auto_purge.py` (29/0): oracle = the SOURCE predicate (independent SQL), old purged
  / recent kept in every scoped table, transactional rollback, INV_SITES read, ≥12 floor, revert-proven at
  the cutoff level. **Runs at retention 600 (cutoff ~1976, older than ALL real data) so it only ever
  deletes synthetic 1970 rows — real data untouched (proven: real-row counts unchanged).**
- **P16 coverage restoration** — `scripts/e2e/test_master_crud_logic.py` (22/0/1) re-covers what the P15
  write gate caused the live per-view tests to SKIP: drives the DEPLOYED Save/Delete scripts under a
  seeded **ProductionControl** session against the real DB → validation rejects (7/7), the **D3 delete-gate
  BLOCKED branch** refuses on a referenced row (6/6; ManifestCost skipped — no refCount gate, faithful), and
  insert+zero-ref-delete round-trip (4/4). Shim gap fixed: `_bind` now ignores `?` in `--` comments/literals
  (JDBC-faithful) so deployed views with a trailing `-- IG-SITE: ... ?` drive correctly.
- **NIT — Clear/New consistency**: New was GATED while Clear was ungated; both are no-DB-write resets.
  Fixed to gate NEITHER (the WRITES Save/Delete stay gated) across all 8 master views + the Sites generator.
  Gateway restart: views deserialize clean (0 `Unable to deserialize`, 0 FAULTED); committed==deployed.
- **NIT — PartsStock line dropdown**: `lineOptions` binds `SELECT DISTINCT LineName FROM VehicleOrder.dbo.LINE`;
  the spike's VehicleOrder has only COROLLA. NOTED as a test-fixture/data-seed item (production VehicleOrder
  carries all lines) — NOT a code defect; VehicleOrder not written.
- **Regressions green**: m4_auth 39, sites_master 28, master_write_gates 86, edi856_e2e 52, m3_reports 50;
  full headless suite 42/42 green. **Spike restored as-found** (max ASN 4715, INV_SITES 12/0/1, report procs
  window-blind original, IX_INV_MANIFEST_COST_MST dropped + no-overlap trigger absent, no sentinels).

## P4 — renban rollover: clean wrap (999→000) + collision-aware allocator (2026-06-23) ✅

The renban-breakdown counter could exceed 999; the legacy `Format('%.3d')`+`varchar(3)` left-TRUNCATED
`1002→'100'` (PROVEN live, mssql-spike), re-seeding the group into the recent `CMWA100x` block → renban
**collision** (the bug ALREADY FIRED Jan-2026 on CMWA, source-truth §4). The spike-1 rebuild reproduced it
for parallel-run parity. **P4 ships the post-cutover fix** (David pre-decided; parallel run complete):
- **Clean wrap (D-RNB-1):** persist `next_count % 1000` and ring-wrap the renban-number tail
  (`group + '%03d' % (rcount % 1000)`) — a straddle emits `…998/…999/…000/…001` (NO 4-digit `CMWA1000`),
  `999` used once. `next_count` keeps the RAW count (the `% 1000` is the single authoritative persist wrap).
  A non-rollover run is byte-for-byte unchanged. Oracle derived from the AMENDED spec §12.7 ring math (R15),
  non-vacuity proven by REVERTING both wraps (renban string flips red in `test_renban_build.py`; persist
  count flips red — DB stores `'100'` — in `test_renban_e2e.py`).
- **Collision allocator (D-RNB-2, WARN→GUIDE→FIX):** `check_renban_collisions` (resident-rows,
  status-independent, EXACT-equality predicate — self-safe since inputs are blank `''`), `next_free_run`
  (forward-ring RUN-of-N: first base whose `[base..base+N-1] % 1000` are all free; exhaustion→None→hard
  WARN-cancel). WARN = a `RENBAN_COLLISION` row in `INV_EDI_ALARM_REJ` (colliding renban in
  `VC_MANIFEST_NUMBER varchar(8)`, exact fit) surfaced on the **home-hub attention rail** (a new
  `RenbanCollisions` count in `Home/kpiSummary`). FIX = `commit_renban_breakdown(..., resolution=)`:
  `None`→check+WARN (no write), `use_next_free`→re-map+commit, `override`→commit acknowledged rows+audit
  note; cancel = client-side (alarm stays active). **TOCTOU re-check on ALL THREE commit paths** inside the
  one tx (the concurrent Order commit `order/code.py:130` is a non-blank-renban writer with NO unique
  constraint); the alarm-ack is in-tx (a rolled-back commit keeps the alarm active). Override acknowledges
  only the SEEN-colliding set so a NEWLY-taken number still aborts.
- **Views (driver-only today → minimal shells):** `Order/RenbanBreakdown` + `Order/RenbanCollisionDialog`
  (gen_renban_views.py) load CLEAN on `gwcmd -r` (0 `Unable to deserialize`, 0 FAULTED; gateway re-signed
  both). Designer-finish FLAGGED: NQ data.bin, `/order/renban` route, popup registration, bidirectional
  binding.config, the server-side write gate.
- **Tests green:** renban_build 35/0 (clean-wrap oracle, non-vacuous), renban_e2e 31/0 (clean-wrap persist
  `'002'` + ring-wrapped strings, non-vacuous), **renban_collision_e2e 24/0** (detect / WARN+home-hub /
  run-of-N=302 skipping in-use 301 / use_next_free+override commit+ack / cancel-leaves-alarm / TOCTOU fires
  on all 3 paths via a 2nd-connection lost race / exhaustion→None). Regressions: order_file_e2e 38/0,
  edi856_e2e 52/0. **Spike restored as-found** (CMWA count 297, no sentinel orders/groups/alarms,
  re-pointed part renban-id restored). *Pre-existing, unrelated:* `test_renbangroup_crud.py` has 1 FAIL — a
  stale hard-coded `count:'288'` anchor vs the live `297` (a fixture-data drift, not a P4 regression).

## P6 — forecast-distribution feed (`.frc` text + `.xls` Excel outbound) (2026-06-23) ✅
**GO unchanged.** The OUTBOUND complement of the M2 830 import — forwards TEMA's imported forecast DOWN
to each sub-supplier. Cloned the two proven seams (the `.ord` order-file emit + the M3 POI report-render
lane); no architect fork. `docs/analysis/order/project-library/forecast_distribution/code.py`
(`emit_forecast_distribution`).
- **Driving read:** `SELECT_ForecastSupplier @WeekDate=<today>` → `SELECT * FROM INV_BREAKDOWN_FC_INF
  WHERE VC_WEEK_DATE > @WeekDate` (future weeks only, lexicographic `>` on yyyymmdd) — the SAME table the
  830 import writes. Supplier-sorted → the `VC_SUPPLIER_CODE` break drives one file set per supplier.
- **`.frc` byte format (spec §3, byte-faithful to ForecastBreakdownF.pas:484-508):** positional concat
  `[siteSupplierCode if sendsite] + VC_SUPPLIER_CODE(raw) + VC_PART_NUMBER(raw) + VC_WEEK_DATE(raw) +
  %.2d(IN_WEEK_NUMBER) + %.5d(IN_QTY1..7)`, CRLF after EVERY line incl. last (Writeln). The Pascal `%.Nd`
  is MIN-WIDTH (`-5`→`-00005` not `-0005`; `100000`→`100000` widened not truncated; NULL→`00000`) — proven
  on a negative + an >99999 overflow. Filename `<dir>\<name '/'-stripped>-<code>.frc` (H-CLEAN: strips ONLY
  `/`, NOT the `.ord`'s broader clean). Archive `<dir>\Archive\<name>-<code><yyyymmdd>.frc` (no separator
  before date, H-ARCH-SEP) when LocalFTP.
- **`.xls` Excel (spec §4, POI fresh, no template):** the M3 `report_render` declarative lane — header
  band row 1 + the 11-col data band row 2+ (`VC_SIZE_CODE`, part, weekdate, weeknum, qty1..7) ALL AsString.
  **Excel INCLUDES `VC_SIZE_CODE` (col 1) which the `.frc` OMITS.** SaveAs `<name>-<code>-Forecast.xlsx`.
  Header row-1 titles GOLDEN-PENDING (binary template, P7).
- **SiteSupplierCode (the P6 crash fix):** the legacy `:488` read a PHANTOM `SiteSupplierCode` column
  (exists in no table → crash for the 11/16 sendsite suppliers); the ORIGINAL commented-out line was
  `tcl:=fiSupplierCode` (= `TSiteInfo.SiteSupplierCode` = **`INV_SITES.VC_SUPPLIER_CODE`**, per the spike
  INV_SITES note `VC_SUPPLIER_CODE <- TSiteInfo.SiteSupplierCode`). So sendsite (`BIT_SITE_NUMBER_IN_ORDER=1`)
  prepends `INV_SITES.VC_SUPPLIER_CODE` ('MAS' on the spike), non-sendsite has no lead — NOT the crash, NOT
  dropped-for-all (which would break the sendsite positional parse). Exact byte/width GOLDEN-PENDING (P7),
  same `_M4` stance as the `.ord` leading field. **NOTE: this DIVERGES from the source-truth spec §5** (which
  said "non-sendsite for all, M4-deferred") — the brief + the commented-out `fiSupplierCode` line + the spike
  INV_SITES mapping all point to the concrete `INV_SITES.VC_SUPPLIER_CODE`; built that (matching the `.ord`).
- **Atomicity (H1 fix):** per-supplier temp-then-rename — stage main+archive(+xlsx) to `.tmp`, publish
  (rename) only after the supplier's file(s) complete; a mid-emit failure deletes that supplier's staged
  `.tmp` (no partial final). H2 (NULL Output File Type → BOTH) reproduced via `_file_kind`.
- **Trigger wiring:** `import_830(..., emitFeed=True)` calls `emit_forecast_distribution` AFTER its own-tx
  commit (reads the freshly-committed breakdown rows); `process_forecast_dir(..., emitFeed=True)` does the
  same post-commit per processed 830 (the operational EDIUpload path). The emit is a SEPARATE callable
  (testable independently) + non-fatal (a feed failure is logged, the import stays committed — legacy
  Execute's try/except logs + returns FALSE without un-doing the import).
- **Tests green:** `test_forecast_distribution_e2e.py` 35/0 — REAL `emit_forecast_distribution` via the
  shim against the spike DB, Excel via the REAL POI lane (bundled jython) read back with openpyxl. Covers
  the exact `.frc` byte format (widths + negative + >99999 overflow + NULL→0 + CRLF), TEXT/EXCEL/BOTH/
  NULL→BOTH, the sendsite `INV_SITES.VC_SUPPLIER_CODE` prefix vs non-sendsite (byte-derived from the
  SOURCE spec §3, R15), the per-supplier file break, future-weeks-only, atomicity (mid-emit render fail →
  no partial), the archive copy, the Excel cols 1-11 + `VC_SIZE_CODE`, and the import trigger end-to-end.
  Non-vacuity proven by reverting `_format_qty` to Python `'%05d'` (`-0005 != -00005` → check fails).
  **Found + faithfully emitted the spike's 290 REAL future-dated breakdown rows for 11 live suppliers** (the
  feed is correct against real committed data, not only synthetic). Regressions: forecast_import_e2e 35/0,
  order_file_e2e 38/0. **Spike restored as-found** (0 synthetic rows; INV_SITES.VC_SUPPLIER_CODE untouched).
- **Hand off** to `sql-adversary` (the SELECT_ForecastSupplier/SELECT_SupplierInfo query + `.frc` format
  parity) + `ignition-code-reviewer` (the emit/atomicity/Excel/trigger wiring).
