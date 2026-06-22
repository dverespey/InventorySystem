# M4 source-truth — the `site_id` schema-surgery map

> **⚠️ SUPERSEDED (David 2026-06-22) — this surgery is DROPPED.** The customer direction reversed: each site
> runs on its OWN gateway + DB (single-site deployments, NOT shared-DB multi-tenancy). With isolated per-site
> DBs, no `site_id` is needed — none of the surgery below (the 32-table `site_id` add, the 7-trigger lockstep,
> the Class-B unique-index swaps, the EIN-collision scoping) applies. Kept for reference only. See the
> `project-multisite` memory (reversed) + punch-list P8. The auth/Sites-master/hardening half of M4
> (`m4-auth-sites-sourcetruth.md`) still applies, simplified to single-site.

Source-truth analysis for the **foundational, riskiest half of M4**: putting `site_id` (the
canonical `IN_SITE_ID` FK, D2) onto the schema end-to-end. This doc is the deliberate companion to
`m4-auth-sites-sourcetruth.md`, which states *"the `site_id` schema surgery + `_HIST` lockstep is
owned by the parallel sql-analyst pass; THIS doc owns auth/Sites-master/cross-DB."* So **THIS doc
owns**: the per-table `site_id` verdict, the 3 history tables + the 7-trigger `SELECT *` lockstep,
the EIN-keyed-update site-scoping (BLOCKER-2), the hardcoded `6440` literal, and the
staged-additive rollout sequence + its sequencing hazards.

**NO DDL is executed here.** Every step is *described*, not run. All claims are proven on the live
spike (`mssql-spike`, working `Inventory` DB) and cross-checked against the authoritative dump
`DB Schema/CreateInventory.sql` (UTF-8 mirror `/tmp/inv_utf8.sql`, 2026-06-12). Where the live DB
and the dump agree I say so; no drift was found in this surface.

Verification stance: live DB wins. The working `Inventory` is the restored legacy `.bak`; the
matched legacy snapshot `Inventory_Live` and the dump were used to confirm the EDI procs still
carry their original bodies.

---

## 0. What is ALREADY on the schema (the starting state — don't re-do it)

Three groups of `site_id`/`IN_SITE_ID` columns already exist on the spike and must be reconciled,
NOT re-added:

| Column found (live `sys.columns`) | On table | Origin | M4 disposition |
|---|---|---|---|
| `IN_SITE_ID int` | `INV_SITES` (PK) | Built (`spike-inv-sites-table.sql`) | The canonical site master — the FK target. |
| `IN_SITE_ID int` | `INV_EDI_INBOUND_LOG`, `INV_EDI_ALARM_REJ` | M1/M2 new tables, built site-aware day-one | Keep; just add the FK constraint to `INV_SITES`. |
| `site_id int` | `INV_STOCK_LEDGER` | M1 new table, built site-aware day-one | Keep; **rename to `IN_SITE_ID` for naming consistency** (see Hazard H6). |
| `site_id int` | `INV_PARTS_STOCK_MST` **and** `INV_PARTS_STOCK_MST_HIST` | Throwaway Check-B scaffold (col 33 on both) | **Reconcile**: rename `site_id`→`IN_SITE_ID`, add FK. The HIST mirror is the spike's F1 workaround and is already correct. |

> Trap: there are **two naming conventions in flight** — the canonical `IN_SITE_ID` (INV_SITES,
> EDI logs) vs the throwaway lowercase `site_id` (PARTS_STOCK, STOCK_LEDGER). D2 mandates
> `IN_SITE_ID` as the sole canonical column name. **Settle on `IN_SITE_ID` everywhere in M4**;
> every `-- M4` marker and `siteScopedQuery()` must target the same column name or the structural
> guard leaks. (Hazard H6.)

`INV_SUPPLIER_MST.BIT_SITE_NUMBER_IN_ORDER` is a **bit flag**, NOT a site key — ignore it for the
surgery.

---

## 1. Per-table `site_id` verdict (43 INV_ tables, live row counts)

Criterion (D1, verbatim intent: *"sites run independently with no shared inventory or data… all
the current tables foreign-keyed by the site"*): **does a row belong to one site, such that two
sites' rows would collide or need separating?** If yes → needs `IN_SITE_ID`. The only exceptions
are pure global lookup/enum/app-config tables that hold the SAME content for every site.

### Verdict = NEEDS `IN_SITE_ID` (per-site business data) — **29 base tables**

| Table | rows (live) | Why per-site |
|---|---|---|
| `INV_ASN_MST` | 2550 | ASN per site; status flip target (BLOCKER-2) — has **no site col yet** |
| `INV_ASN_DETAIL_MST` | 39707 | manifest lines per site; the `(site_id,IN_ASN_ID,manifest)` re-key (Q1) |
| `INV_INV_MST` | 2934 | invoice/EIN per site; status flip target (BLOCKER-2) — **no site col yet** |
| `INV_INVOICE_INF` | 3 | invoice detail per site |
| `INV_OPEN_ORDER_INF` | 4238 | open orders per site (**HIST lockstep table**) |
| `INV_PARTS_STOCK_MST` | 47 | on-hand stock per site (Q1) — has throwaway `site_id` (**HIST lockstep**) |
| `INV_PART_QTY_INF` | 7582 | qty-change audit per site |
| `INV_PART_SHIPPING_INF` | 886 | shipping detail per site |
| `INV_PART_RATIO` | 0 | part ratio per site |
| `INV_FORECAST_DETAIL_INF` | 50 | forecast per site (**HIST lockstep table**) |
| `INV_FORECAST_INF` | 1041 | forecast rollup per site |
| `INV_BREAKDOWN_FC_INF` | 959 | breakdown forecast per site |
| `INV_MANUAL_FORECAST` | 0 | manual forecast per site |
| `INV_MANIFEST_COST_MST` | 45 | per-site manifest cost windows (the `IX_INV_MANIFEST_COST_MST` constraint, cutover residual) |
| `INV_SUPPLIER_MST` | 16 | suppliers per site (D1 Q2) — UQ key must become `(site,code)` |
| `INV_LOGISTICS_MST` | 1 | carriers per site (D1 Q2) — UQ key must become `(site,name)` |
| `INV_SIZE_MST` | 64 | size catalog per site (D1 Q2) — UQ key must become `(site,code)` |
| `INV_RENBAN_GROUP_MST` | 5 | renban grouping per site — UQ key must become `(site,code)` |
| `INV_ASSY_RATIO_MST` | 0 | assembly ratio per site |
| `INV_ASSY_RATIO` | 0 | assembly ratio per site |
| `INV_ASSY_BUILD_HIST` | 0 | assembly build history per site (NOTE: a `_HIST` table but **NOT** a `SELECT *` trigger target — explicit VALUES; see §2) |
| `INV_ASSY_MONTHLY_PO` | 0 | monthly PO per site |
| `INV_ASSY_PO_CHARGED` | 0 | PO charged per site |
| `INV_SHIPPING_INF` | 82 | shipping records per site |
| `INV_STOCKTAKING_INF` | 13 | stocktake per site |
| `INV_REJECT_INF` | 0 | rejects per site |
| `INV_RENBAN_GROUP_MST` | (above) | — |
| `INV_FIRST_PRODUCTION_DAY` | 9 | first-prod-day calendar per site |
| `INV_OVERTIME_HOLIDAY` | 0 | per-site overtime/holiday calendar |
| `INV_STOCKTAKING_INF` | (above) | — |

Plus the **3 already-site-aware new tables** (reconcile, don't add): `INV_STOCK_LEDGER` (0),
`INV_EDI_INBOUND_LOG` (0), `INV_EDI_ALARM_REJ` (0).

**Count needing `IN_SITE_ID`: 29 base tables that lack it + 3 already-built site-aware = 32
per-site tables total.** (`INV_PARTS_STOCK_MST` is counted as needing-reconcile because its
column is the throwaway `site_id`, not the canonical FK.)

### Verdict = GLOBAL / SHARED (do NOT add `IN_SITE_ID`) — **8 tables**

| Table | rows | Why global |
|---|---|---|
| `INV_ADD_POINT_INF` | 2 | enum lookup: `A=ARRIVED / S=SHIPPED` — identical for all sites |
| `INV_PART_TYPE_MST` | 5 | enum: TIRE/WHEEL/FILM/VALVE/MISC — domain-wide |
| `INV_PART_TYPE_INF` | 0 | child of the global part-type enum |
| `INV_BC_RATIO` | 0 | broadcast-code ratio constants (domain math, not site data) |
| `INV_PASSWORD_RESET_DAYS` | 1 | single global app-policy value (90) |
| `INV_PROGRAM_VERSION` | 1 | single global app version (2.9.4.1) |
| `INV_SITES` | 2 | **the site master itself** — IS the site dimension, not scoped by it |
| `inv_temp` | 0 | scratch/staging table — out of scope |

> **Decision point for David (DP-1):** `INV_DOCK_INF` (8 rows: `1AN01…`) and `INV_USERS` (4) are
> JUDGMENT CALLS. Dock codes look plant-specific (each TMM plant has its own docks) → **lean
> per-site**, but legacy treats them as a flat list. `INV_USERS` is the legacy auth table; the
> rebuild replaces it with Ignition's User Source (`m4-auth-sites-sourcetruth.md §2`), so the
> DB column may be moot — **but if any login-era user row is retained, a user belongs to a site
> (auth binds user→site, D1)**. Recommend: dock = per-site; users = retired (Ignition owns it),
> so neither needs a column on the legacy table during parallel-run. Confirm with David.

---

## 2. The 3 HISTORY tables + the F1-safe `SELECT *` lockstep (the riskiest part)

### The 3 lockstep `_HIST` tables (proven live)

A re-audit of **all 25 triggers** (live `sys.triggers` ⋈ `sys.sql_modules`, matching
`SELECT * FROM inserted|deleted`) returns **exactly 7 triggers across exactly 3 base→history
pairs** — confirming the plan's count with zero others hiding:

| `_HIST` table | base table | base cols | hist cols | aligned? |
|---|---|---|---|---|
| `INV_FORECAST_DETAIL_INF_HIST` | `INV_FORECAST_DETAIL_INF` | 18 | 18 | ✅ exact |
| `INV_PARTS_STOCK_MST_HIST` | `INV_PARTS_STOCK_MST` | 33 | 33 | ✅ (already mirrors the throwaway `site_id` at col 33 — the spike's Check-B workaround) |
| `INV_OPEN_ORDER_INF_HIST` | `INV_OPEN_ORDER_INF` | 23 | 23 | ✅ exact |

`INV_ASSY_BUILD_HIST` is the **decoy** — it is a `_HIST`-named table but is populated by an
explicit-column `INSERT … VALUES`, **not** a `SELECT *` trigger, so it is *safe* and not part of
the lockstep. The re-audit found NO other `SELECT *` history trigger.

### WHY this is the lockstep hazard (proven on the spike already — ignition-spike-log.md:34)

Each of the 7 triggers does a bare `INSERT INTO <hist> SELECT * FROM inserted|deleted`. A bare
`SELECT *` into a column-less `INSERT` matches **by ordinal position and column count**, NOT by
name. So:

- Add `IN_SITE_ID` to the **base** table → `inserted`/`deleted` now has N+1 columns → the
  `SELECT *` yields N+1 values into an N-column `_HIST` → **`Msg 213 / column count mismatch`,
  the trigger throws, and because it is an AFTER trigger inside the DML transaction, the base
  INSERT/UPDATE/DELETE itself is ROLLED BACK.** This was reproduced live when the throwaway
  `site_id` was first added to `INV_PARTS_STOCK_MST` alone (ignition-spike-log.md:34-42) — an
  UPDATE immediately broke. The fix was to mirror `site_id` onto `INV_PARTS_STOCK_MST_HIST`.
- Add to the **history** but not the base → the `SELECT *` yields N values into an N+1 `_HIST`
  → same mismatch, same rollback.

**The only safe sequence is to add `IN_SITE_ID` to base and history TOGETHER (same migration,
same transaction window).** Two traps make this subtler than "just add a column":

1. **`SELECT *` matches by POSITION, so put `IN_SITE_ID` in the SAME ordinal on base and hist.**
   The clean choice is *append as the last column on both*. Proven safe because all 3 `_HIST`
   tables have **NO identity column and NO primary key** (live `sys.indexes`/`sys.columns`), so a
   trailing column needs no column-list rewrite of the `INSERT`. The throwaway `site_id` already
   demonstrates this: it sits at col 33 on both PARTS_STOCK base and hist, and the `SELECT *`
   trigger still works.
2. **Name typos in `_HIST` are HARMLESS to `SELECT *` but a TRAP for a human "fix".** The live
   schema already has position-aligned but name-divergent columns:
   - `INV_PARTS_STOCK_MST.VC_LINE_NAME` (col 29) ↔ `INV_PARTS_STOCK_MST_HIST.VC__LINE_NAME`
     (double underscore)
   - `INV_OPEN_ORDER_INF.VC_KANBAN_NUMBER` (col 20) ↔ `INV_OPEN_ORDER_INF_HIST.VC__KANBAN_NUMBER`
   These work today because `SELECT *` is positional. **Do NOT "tidy" these names during the M4
   surgery** unless you also rebuild every dependent — the alignment that matters is *count +
   position*, not name. A reimplementation that rebuilds `_HIST` from the base column list would
   silently rename them and (harmlessly) diverge from the legacy dump; flag it, don't fix it
   mid-surgery.

### The 7 triggers decoded — what each does + the EXACT lockstep change

| # | Trigger (dump line) | On base | Fires | The base→hist copy | Other side-effects (must NOT break) | Lockstep change |
|---|---|---|---|---|---|---|
| 1 | `DeleteForecastDetail` (`:2664`) | FORECAST_DETAIL | FOR DELETE | `INSERT INV_FORECAST_DETAIL_INF_HIST SELECT * FROM deleted` | also cascades `DELETE FROM inv_forecast_inf WHERE vc_part_number IN (SELECT vc_assy_part_number_code FROM DELETED)` | add `IN_SITE_ID` to base+hist together; cascade delete should ALSO gain `AND site=…` once both tables are scoped (else it deletes another site's forecast) — **DP-2** |
| 2 | `UPDATE_ForecastDetailInf` (`:3440`) | FORECAST_DETAIL | FOR UPDATE | `IF @@rowcount>0 INSERT … SELECT * FROM inserted` | none | base+hist together |
| 3 | `INSERTForecastDetail` (`:3617`) | FORECAST_DETAIL | AFTER INSERT | `INSERT … SELECT * FROM inserted` | none | base+hist together |
| 4 | `UPDATE_PartNumber` (`:4107`) | PARTS_STOCK | FOR UPDATE | `IF @numrows>0 INSERT INV_PARTS_STOCK_MST_HIST SELECT * FROM deleted` | writes `INV_PART_QTY_INF` qty-delta rows; on single-row key change, cascades part-number rename into `INV_ASSY_RATIO_MST` | base+hist ALREADY both have throwaway `site_id`@33 → only the **rename `site_id`→`IN_SITE_ID` + FK** remains; verify the qty-delta INSERT (explicit column list) gains `IN_SITE_ID` |
| 5 | `INSERT_PartsStockMST` (`:4269`) | PARTS_STOCK | FOR INSERT | `IF @numrows>0 INSERT INV_PARTS_STOCK_MST_HIST SELECT * FROM inserted` | a `print()` only | as #4 — rename + FK |
| 6 | `UPDATE_RecConfStatPartsStockMstQTY` (`:5475`) | OPEN_ORDER | FOR UPDATE | `IF @numrows>0 INSERT INV_OPEN_ORDER_INF_HIST SELECT * FROM deleted` | **decrements `INV_PARTS_STOCK_MST.IN_QTY` by `d.IN_QTY`** via a 4-way join (PARTS_STOCK⋈deleted⋈inserted⋈SUPPLIER) — the supplier-shipping stock move | base+hist together; the stock-move join keys on `VC_PART_NUMBER` only → **must add `AND site` to the join** once PARTS_STOCK is scoped, or it moves the wrong site's stock — **DP-3** |
| 7 | `INSERT_RecConfStatPartsStockMstQTY` (`:7496`) | OPEN_ORDER | FOR INSERT | `INSERT INV_OPEN_ORDER_INF_HIST SELECT * FROM inserted` | **increments `INV_PARTS_STOCK_MST.IN_QTY`** via two PARTS_STOCK⋈inserted⋈SUPPLIER joins (the receiving/arrival stock adds) | base+hist together; same join-key hazard as #6 — **DP-3** |

> **The lockstep change is two layers, not one.** Layer A (count alignment) keeps the `SELECT *`
> from throwing — add `IN_SITE_ID` to base+hist together. Layer B (correctness) is the
> *side-effect joins/cascades* inside triggers 1, 6, 7 that key on `VC_PART_NUMBER` /
> `vc_assy_part_number_code` **alone**: once two sites can hold the same part number (D1), these
> joins fan across sites and move/delete the wrong site's data. Layer B is the same class of bug
> as BLOCKER-2 (§3), just inside triggers. **Layer A is mandatory at the schema add; Layer B must
> land before per-site data coexists** (i.e. before a second site's rows are loaded), not
> necessarily at the column add. Note: in the parallel-run single-site phase Layer B is inert
> (only one site exists), so it can be staged after Layer A — but it MUST be done before cutover
> to true multi-site.

---

## 3. EIN-keyed updates to site-scope (BLOCKER-2)

The 997/824 ack path flips a row's status keyed on **EIN alone**. With per-site EIN sequences
(`INV_SITES.IN_EIN_SEQ`, Q4 — already built), two sites can hold the same EIN, so an unscoped flip
acks the wrong site. The live targets `INV_ASN_MST` and `INV_INV_MST` confirmed to have **no site
column yet** (live `sys.columns`) — they are the BLOCKER-2 schema targets.

### The canonical offender — `UPDATE_EINStatus` (live body, verified)

```
if @EINType = 'SH'
    UPDATE INV_ASN_MST SET VC_ASN_STATUS = @EINStatus WHERE IN_ASN_EIN = @EIN   -- 997/824 ACK
else
    UPDATE INV_INV_MST SET VC_INV_STATUS = @EINStatus WHERE IN_INV_EIN = @EIN   -- invoice ack
```

Both branches key on `IN_*_EIN = @EIN` **alone**. **Fix:** add `IN_SITE_ID` to both target tables
and to the WHERE → `WHERE IN_SITE_ID=@SiteId AND IN_ASN_EIN=@EIN`. The inbound 997/824 already
resolves a site by DUNS (ISA `delSL[4]` → `VC_TMM_DUNS`, Q7/Q11), so `@SiteId` is in hand at flip
time (it is already threaded as a `-- M4` marker in the inbound edi_inbound code per
ignition-spike-log.md:575). The Ignition rebuild already re-implements this flip in-line so it can
`CHECK @@ROWCOUNT` (silent-success guard, ignition-spike-log.md:560) — the M4 add is one predicate.

### The audit of EVERY EIN-keyed write (the full BLOCKER-2 set)

Live audit of all modules referencing `IN_ASN_EIN`/`IN_INV_EIN`, classified:

| Object | Touches EIN how | Site-scope action |
|---|---|---|
| `UPDATE_EINStatus` | **UPDATE … WHERE EIN** (both branches) | **Add `IN_SITE_ID` to WHERE** (the primary BLOCKER-2 fix) |
| `REPORT_EDI856` | embedded self-flip `UPDATE INV_ASN_MST … WHERE EIN` in the `@EIN<>0` branch (BLOCKER-1) + the `6440` literal | **already removed** in the Ignition rebuild (pure-SELECT NQ feeds the builder); the flip is re-done in-line site-scoped at send. Legacy still carries it (`Inventory_Live` + working `Inventory` both `HAS_6440 +EMBEDDED_UPDATE`). |
| `REPORT_EDI810` / `REPORT_EDI810Recreate` | **read** EIN as a projected column (`IN_INV_EIN 'EIN'`); status predicate on `VC_INV_STATUS`/`VC_ASN_STATUS`, not EIN | when the invoice/ASN selects are site-scoped (general §1), these need `AND IN_SITE_ID=@Site` on the FROM — read-scoping, not a flip |
| `UPDATE_INVRecreate` | `UPDATE INV_INV_MST SET VC_INV_STATUS='S' WHERE VC_INV_STATUS='C'` — keyed on **status, not EIN, no WHERE-EIN** | **un-scoped MASS update** — flips EVERY site's 'C' invoices to 'S'. **Must add `AND IN_SITE_ID=@Site`** or one site's recreate wipes another's invoice state. **Same class as BLOCKER-2** (status-wide, not EIN-keyed, but cross-site). Add to the BLOCKER-2 list. |
| `INSERT_ASNDetail` | references `IN_ASN_EIN` only as a column read (no UPDATE on EIN) | covered by the Q1 `(site_id,IN_ASN_ID,manifest)` re-key, not a BLOCKER-2 item |
| `SELECT_ASNList` / `SELECT_INVOICEList` | **read/list** by EIN | site-scope the SELECT (`AND IN_SITE_ID=@Site`) — read-scoping |

**BLOCKER-2 write-scoping set (must add `IN_SITE_ID` to the WHERE before per-site EIN goes live):**
1. `UPDATE_EINStatus` (both branches) — the 997/824 ack.
2. `UPDATE_INVRecreate` — the un-WHERE'd status-wide invoice reset (newly surfaced; previously
   not on the BLOCKER-2 list).
3. `REPORT_EDI856`'s embedded `@EIN<>0` self-flip — already excised in the rebuild but listed for
   completeness (it is the same bug class the plan flagged).

---

## 4. Hardcoded single-site literals to parameterize

| Literal | Location | What it is | Fix |
|---|---|---|---|
| `6440` | `REPORT_EDI856` `@EIN<>0` branch — **dump `:3683`** `WHERE a.IN_ASN_EIN = 6440` (live: present in working `Inventory` AND `Inventory_Live`, no drift) | a **baked-in single-site EIN** — the legacy "recreate this site's 856" path filters to one hardcoded site EIN | parameterize → `WHERE IN_SITE_ID=@Site AND a.IN_ASN_EIN=@EIN`. The rebuild's pure-SELECT path already dropped this branch; the literal is a legacy-only artifact (cutover-architecture.md:161/230). |
| `'6440'`-style site assumptions in the `.ord` leading field | order_file (P8 punch-list) | the supplier-code lead field is single-site | already marked `_M4` → `INV_SITES.VC_SUPPLIER_CODE` (width pending golden, ignition-spike-log.md:723) |

No other hardcoded site EIN/DUNS literal was found in any proc/trigger (`sys.sql_modules LIKE
'%6440%'` returns ONLY `REPORT_EDI856`). The other single-site assumptions are INI-sourced (paths,
DUNS, separators) and relocate into `INV_SITES` columns per `m4-auth-sites-sourcetruth.md`.

---

## 5. The staged-additive rollout (described, NOT executed)

The plan's approach — **add `IN_SITE_ID` NULLABLE with a default = the current single site during
parallel-run, backfill, enforce NOT-NULL at cutover** — is FEASIBLE for most tables, but **NOT for
the tables whose business-key UNIQUE index must change** (those need a key change, not an additive
column). The unique-index audit (live `sys.indexes`) splits the tables into two rollout classes.

### Class A — pure additive-nullable is SAFE (no unique-key change)

These per-site tables have only a surrogate PK + non-unique indexes, so adding a nullable
`IN_SITE_ID` defaulted to the current site breaks nothing: `INV_ASN_MST` (its UQ is separate — see
Class B), `INV_INV_MST`, `INV_OPEN_ORDER_INF`, `INV_PART_QTY_INF`, `INV_PART_SHIPPING_INF`,
`INV_FORECAST_INF`, `INV_BREAKDOWN_FC_INF`, `INV_MANIFEST_COST_MST` (but see the cutover residual),
`INV_SHIPPING_INF`, `INV_STOCKTAKING_INF`, `INV_REJECT_INF`, the `_HIST` trio (no PK/identity), and
the already-site-aware new tables.

Ordered DDL per Class-A table (DESCRIBED):
1. `ALTER TABLE … ADD IN_SITE_ID INT NULL CONSTRAINT DF_… DEFAULT (<current_site_id>)` — additive,
   reversible (`DROP CONSTRAINT` + `DROP COLUMN`).
2. (for the 3 base tables with `SELECT *` `_HIST` triggers) **add the column to the `_HIST` table
   in the SAME step/transaction** — append as last column on both (Layer A, §2).
3. Backfill existing rows: `UPDATE … SET IN_SITE_ID=<current_site_id> WHERE IN_SITE_ID IS NULL`
   (the DEFAULT covers new rows; backfill covers the historical ones).
4. Add the FK: `ADD CONSTRAINT FK_…_SITE FOREIGN KEY (IN_SITE_ID) REFERENCES INV_SITES(IN_SITE_ID)`.
5. **At cutover only:** `ALTER COLUMN IN_SITE_ID INT NOT NULL` (+ drop the now-redundant DEFAULT,
   or keep it as the session default). Reversible up to here.

### Class B — additive-nullable is NOT safe; the UNIQUE key must change (a key change, not additive)

These carry a **single-column business-key UNIQUE index** that D1 says must become
`(IN_SITE_ID, <key>)`. You cannot just add a nullable column — the existing global unique would
reject a second site reusing a part/supplier/size code. (Live `sys.indexes`.)

| Table | existing UNIQUE | must become | hazard |
|---|---|---|---|
| `INV_PARTS_STOCK_MST` | `IX_INV_PARTS_STOCK_MST (VC_PART_NUMBER)` | `(IN_SITE_ID, VC_PART_NUMBER)` | two sites can't share a part number under the global UQ |
| `INV_SUPPLIER_MST` | `IX_INV_SUPPLIER_MST (VC_SUPPLIER_CODE)` | `(IN_SITE_ID, VC_SUPPLIER_CODE)` | per-site supplier codes (D1 Q2) |
| `INV_SIZE_MST` | `IX_INV_SIZE_MST (VC_SIZE_CODE)` | `(IN_SITE_ID, VC_SIZE_CODE)` | per-site size catalog |
| `INV_LOGISTICS_MST` | `IX_INV_LOGISTICS_MST (VC_LOGISTICS_NAME)` | `(IN_SITE_ID, VC_LOGISTICS_NAME)` | per-site carriers |
| `INV_RENBAN_GROUP_MST` | `IX_INV_RENBAN_GROUP_MST (VC_RENBAN_GROUP_CODE)` | `(IN_SITE_ID, VC_RENBAN_GROUP_CODE)` | per-site renban groups |
| `INV_ASN_MST` | `UX_INV_ASN_MST_LINE_PDATE_NORMAL (VC_LINE_NAME, VC_PRODUCTION_DATE)` **FILTERED** `WHERE VC_START_SEQ_NUMBER <> '-1'` | `(IN_SITE_ID, VC_LINE_NAME, VC_PRODUCTION_DATE)` same filter | **R23 filtered-index gate** — the filtered unique guard must be REBUILT with `IN_SITE_ID` as the leading key; a key change requires DROP+CREATE, not `ALTER` |
| `INV_PART_TYPE_MST` | `IX_INV_PART_TYPE_MST (VC_PART_TYPE)` | **stays global** (§1 verdict) — do NOT add site | (global enum; do not change) |
| `INV_STOCK_LEDGER` | `UQ_INV_STOCK_LEDGER_EVENT (IN_PART_ID, VC_SOURCE_EVENT)` | `(IN_SITE_ID, IN_PART_ID, VC_SOURCE_EVENT)` | `IN_PART_ID` is already per-site so this is implicitly scoped, but add `IN_SITE_ID` for an explicit guard + index seek |

Ordered DDL per Class-B table (DESCRIBED — the EXTRA steps vs Class A):
1. Add the nullable `IN_SITE_ID` column (+ DEFAULT current site) — as Class A step 1.
2. Backfill — as Class A step 3.
3. **DROP the single-column UNIQUE index, CREATE the composite `(IN_SITE_ID, <key>)`** (for
   `INV_ASN_MST` re-create the FILTERED predicate verbatim). This is the **non-reversible-by-ALTER**
   step — it must be a DROP+CREATE, and it CANNOT happen while `IN_SITE_ID` is still NULL on any
   row (the composite UQ would treat NULL site as a distinct group and could admit a legacy
   duplicate, or fail if duplicates already span the to-be-NULL site). So: **backfill to NOT-NULL
   value first (step 2), THEN swap the index (step 3).**
4. Add the FK to `INV_SITES`.
5. At cutover: `ALTER COLUMN … NOT NULL`.

> **The Class-B sequencing is the second-biggest hazard after the trigger lockstep**: you must
> backfill BEFORE swapping the unique index, and the index swap is a DROP+CREATE (the old
> single-column UQ has to go before the composite can guarantee per-site uniqueness). During
> single-site parallel-run the composite `(1, code)` is behaviorally identical to the old
> `(code)`, so the swap is safe to do early — but it is irreversible-by-ALTER, so script the
> rollback as the inverse DROP+CREATE.

### The FK direction / order across all classes
Add columns + backfill on ALL tables first; add the FKs to `INV_SITES` last (after every table has
a valid non-NULL site value), so no FK insert order problem. `INV_SITES` is already populated
(2 rows) so the FK target exists from day one.

---

## 6. Top sequencing hazards + decisions for David

**Hazards (ordered by blast radius):**

- **H1 — The 7-trigger lockstep is a HARD-FAIL, not a warning.** Add `IN_SITE_ID` to any of
  FORECAST_DETAIL / PARTS_STOCK / OPEN_ORDER **without** the matching `_HIST` add and the very
  next INSERT/UPDATE/DELETE on that base table **rolls back** (column-count mismatch in the
  `SELECT *` trigger). Proven live (ignition-spike-log.md:34). Mitigation: base+hist in one
  migration step, append-as-last-column, and re-run a smoke DML on each of the 3 tables after.
- **H2 — Trigger side-effect joins fan across sites (Layer B).** Triggers 1/6/7 join/cascade on
  `VC_PART_NUMBER` / `vc_assy_part_number_code` alone; once two sites share a part number they
  move/delete the wrong site's stock/forecast. Inert during single-site parallel-run; **must be
  fixed before a 2nd site's data coexists.**
- **H3 — Class-B unique-index swaps are irreversible-by-ALTER + order-sensitive.** Backfill to a
  non-NULL site BEFORE the DROP+CREATE of the composite UQ; the `INV_ASN_MST` one is a FILTERED
  index (R23) that must re-create its predicate verbatim.
- **H4 — `UPDATE_INVRecreate` is an un-WHERE'd cross-site status reset** (newly surfaced). It
  flips EVERY site's 'C' invoices to 'S'. Add `AND IN_SITE_ID=@Site` — same class as BLOCKER-2,
  add it to that fix list.
- **H5 — `UPDATE_EINStatus` is the canonical BLOCKER-2** — one predicate add per branch, but it
  is the 997/824 ack so getting the `@SiteId` thread-through right (from the DUNS-resolved inbound
  file) is the integration point.
- **H6 — Two column-name conventions in flight** (`IN_SITE_ID` canonical vs throwaway `site_id` on
  PARTS_STOCK/STOCK_LEDGER). `siteScopedQuery()` injects ONE name; a mismatch = a silent
  un-scoped query = cross-site leak (the #1 D1 risk). Settle on `IN_SITE_ID` everywhere in M4 and
  rename the two throwaway columns as part of the surgery.

**Decisions needed from David:**

- **DP-1 — `INV_DOCK_INF` and `INV_USERS`: per-site or global?** Recommend dock = per-site, users
  = retired (Ignition User Source owns auth, so no legacy column needed during parallel-run).
- **DP-2 — `DeleteForecastDetail` cascade:** once FORECAST tables are scoped, should the
  `DELETE FROM inv_forecast_inf … IN (SELECT … FROM DELETED)` cascade gain `AND IN_SITE_ID`? (Yes,
  to avoid deleting another site's forecast — but it is inert single-site.)
- **DP-3 — The receiving/shipping stock-move joins (triggers 6/7):** confirm the stock
  add/decrement must be site-scoped (`AND PS.IN_SITE_ID = i.IN_SITE_ID`). Recommend yes; inert
  single-site, mandatory before multi-site data coexists.
- **DP-4 — Cutover residual carried in (not new to M4):** verify the `IX_INV_MANIFEST_COST_MST`
  constraint-drop on prod (cutover-readiness-checkpoint) — `INV_MANIFEST_COST_MST` is per-site and
  its constraint interacts with the multi-window per-`(site,assy)` rule (D1).

---

## Evidence index (queries run on `mssql-spike`, working `Inventory`, read-only)

- 7-trigger / 3-table re-audit: `sys.triggers ⋈ sys.sql_modules WHERE definition LIKE '%SELECT *
  FROM inserted|deleted%'` → exactly 7 rows, 3 base tables; no others.
- base↔hist column counts: 18/18 (FORECAST_DETAIL), 33/33 (PARTS_STOCK, incl. throwaway `site_id`
  @33 on both), 23/23 (OPEN_ORDER).
- `_HIST` have NO identity / NO PK (3/3): `sys.columns is_identity` + `sys.indexes is_primary_key`.
- name-divergence (positional, harmless): `VC_LINE_NAME`↔`VC__LINE_NAME` (PARTS_STOCK col 29),
  `VC_KANBAN_NUMBER`↔`VC__KANBAN_NUMBER` (OPEN_ORDER col 20).
- `UPDATE_EINStatus` body: both branches `WHERE IN_*_EIN = @EIN` alone (live OBJECT_DEFINITION).
- `INV_ASN_MST` / `INV_INV_MST` have NO site column (live `sys.columns`) — BLOCKER-2 targets.
- `6440`: ONLY `REPORT_EDI856`, `@EIN<>0` branch, dump `:3683`; present in both `Inventory` and
  `Inventory_Live` (no drift).
- unique/PK index inventory: `sys.indexes WHERE is_primary_key OR is_unique` — the 6 single-column
  business-key UQs (Class B) + the FILTERED `UX_INV_ASN_MST_LINE_PDATE_NORMAL`
  (`filter_definition = ([VC_START_SEQ_NUMBER]<>'-1')`).
- already-site-aware columns: `INV_SITES.IN_SITE_ID`, `INV_EDI_INBOUND_LOG.IN_SITE_ID`,
  `INV_EDI_ALARM_REJ.IN_SITE_ID`, `INV_STOCK_LEDGER.site_id`, `INV_PARTS_STOCK_MST(.site_id)`+HIST.
- 43 INV_ tables with live row counts (§1 table).
