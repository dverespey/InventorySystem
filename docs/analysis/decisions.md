# Domain Decisions Log

Decisions from the domain expert (David) that resolve the §8 "Open questions" raised
across the module specs. Each decision has an ID (`D#`); specs reference it when their
open question is closed. Newest decisions appended at the bottom.

---

## D1 — Multi-site: independent, fully-isolated sites  *(2026-06-07)*

**Resolves:** Group 1, Q1–Q3 — the recurring multi-site §8 question in
`supplier`, `logistics`, `size`, `manifest-cost`, `master-maint`, `parts-stock-master`,
`stocktaking`, `inv-mgmt`, and `logistics-breakdown`.

**Decision (verbatim intent):** *"The sites are run independently with no shared inventory
or data. All the current tables would be foreign-keyed by the site. All site info should now
move from the global INI into the site table."*

**What this means for the rebuild:**
- **Tenancy = shared schema, `site_id` FK on every table.** Add a new **`sites`** table; every
  existing `INV_*` table gains a `site_id` (NOT NULL) FK. **Full data isolation** — no site sees
  another's rows; every query is scoped to the current site. (Answers Q1: on-hand stock is
  **per-site**; Q2: suppliers, logistics/carriers, sizes, the part catalog, and assembly prices
  are all **per-site**, not shared.)
- **All `[SITE]` INI config becomes `sites` rows** (Q3). PlantName, Assembler/SupplierCode, DUNS,
  the EDI feature flags (`POEDISupport`, `GenerateEDI`), directory paths, etc. move out of the
  single-install INI and into per-site columns. `SiteInfo.pas`/INI reads → a `Site` model.
- **Uniqueness becomes per-site.** Every previously "globally unique" business key
  (`VC_SUPPLIER_CODE`, `VC_SIZE_CODE`, `VC_LOGISTICS_NAME`, `VC_PART_NUMBER`, the manifest assy
  code, …) becomes unique **within a site**: composite unique `(site_id, <key>)`, not global.
- **App-layer pattern:** every ActiveRecord model `belongs_to :site` with enforced current-site
  scoping (e.g. `acts_as_tenant`/`default_scope`); **auth binds each user to a site** (the
  "current site" replaces the single-install INI identity).
- **Phasing:** the `site_id` FKs + per-site unique indexes land in the **Postgres / DB-modernization
  phase**. During the parallel-run phase the legacy single-site SQL Server DB is untouched and the
  new app simply filters to the one site it represents.

> Closes the "today the table has no site column; key is globally unique — shared or per-site?"
> note that recurred in nearly every spec. The answer is uniformly **per-site, fully isolated**.

See [[project-multisite]] in the modernization notes.

---

## D2 — Surrogate int IDs are the only key; business codes/names are editable attributes  *(2026-06-12)*

**Resolves:** the recurring "name/code as key, is renaming expected?" §8 question in
`supplier` (§8.3), `logistics` (§8.3), `size` (§8.5), and `parts-stock-master` (§8.4).

**Decision (verbatim intent):** *"Yes, all keys should be done through the surrogate. A part
number or supplier code should be editable, not a key — although a change there is an extremely
rare event."*

**What this means for the rebuild:**
- **The surrogate integer id is the sole key.** Every FK, join, and lookup resolves on the
  surrogate id (`IN_SUPPLIER_ID`, `IN_LOGISTICS_ID`, `IN_SIZE_ID`, `IN_PART_ID`, …) — never on the
  business string. This is already how most parts/FKs behave; the decision makes it **uniform**.
- **Business codes/names are plain editable attributes** (`VC_SUPPLIER_CODE`, `VC_LOGISTICS_NAME`,
  `VC_SIZE_CODE`, `VC_PART_NUMBER`, etc.). They are **not** keys and carry no referential weight.
- **Renames are allowed but extremely rare,** and are **safe with no cascade** precisely because
  nothing references the string — a rename is a single-row attribute UPDATE.
- **Legacy string-keyed callers must be reworked to resolve by id.** Concretely, the rebuild must
  fix the paths that legacy code resolved by string rather than id: the supplier-save procs and the
  monthly report `@Logistics` filter (resolve logistics by `IN_LOGISTICS_ID`), `UPDATE_SizeUsage` /
  `SELECT_SizeUsage` and the size form search (resolve by `IN_SIZE_ID`), the `UPDATE_PartNumber`
  string-cascade and the transactional children that key on `VC_PART_NUMBER` (link by `IN_PART_ID`).
  These string-cascade/partial-cascade behaviors disappear once everything keys on the id.
- **Interaction with D1:** the code stays a **unique attribute per-site** — composite unique
  `(site_id, <code>)` from D1 still holds — but uniqueness is now a *constraint on an attribute*,
  not a key. A rename simply has to keep the code unique within its site.

> Closes the "should we standardize on the surrogate id / are renames expected?" question that
> recurred across the masters. Answer: **yes, surrogate id everywhere; codes are editable, rename-safe.**

---

## D3 — Block deletes that are still referenced (RESTRICT); archival is a separate future capability  *(2026-06-12)*

**Resolves:** the recurring "delete when referenced: block / nullify / dangle?" §8 question in
`logistics` (§8.4), `size` (§8.4), `manifest-cost` (§8.3), `parts-stock-master` (§8.2 + §8.3),
and `stocktaking` (§8.4).

**Decision (verbatim intent):** *"Block the delete when referenced. There should in the future be
an archival function to remove the data from view and/or the primary database."*

**What this means for the rebuild:**
- **RESTRICT, uniformly.** Deleting any record that is still referenced by another row is
  **blocked** with a clear error. This replaces every inconsistent legacy behavior — the
  null-one-FK-but-dangle-the-rest triggers, the silent inner-JOIN line loss, and the orphaned
  transactional children. No more dangling FKs, no more nulled-out links on delete.
  - Master deletes are blocked while referenced: a **logistics** row referenced by any supplier or
    part; a **size** referenced by any part (current or `_HIST`); a **part** referenced by any open
    order / reject / stocktaking / part-shipping / assy-ratio / forecast row; a **manifest-cost**
    price referenced by any ASN-detail / invoice line.
  - **Transactional children** (stocktaking adjustments, orders, rejects, shipping lines) likewise
    are not orphaned — they keep a **real FK** to their parent (e.g. add `PK_INV_STOCKTAKING_INF`
    on `IN_STOCKTAKING_ID` and FK `IN_PART_ID → INV_PARTS_STOCK_MST`), so the parent cannot be
    hard-deleted out from under them.
- **Archival is a SEPARATE, FUTURE capability — not delete.** "Getting rid of" a record that is
  referenced is done by **archival**, not deletion: an archival function that **removes the data
  from view and/or moves it out of the primary database** (soft-delete / status flag that hides it
  from pickers and default queries, and eventually relocates aged data to an archive store).
  Archival is explicitly **out of scope for the initial rebuild** — design models so it can be
  added later (e.g. a nullable `archived_at` / status column, queries that default to active rows),
  but the first cut only needs RESTRICT-on-delete.
- **Supersedes the spec sub-options.** Wherever a spec offered "nullify part links too" or "hard
  delete is acceptable because prices are future-dated," the answer is now **block instead**.

> Closes the delete-policy question across the masters and the stock ledger. Answer: **block
> (RESTRICT) when referenced; never dangle or null; archival/soft-delete is a later, separate feature.**

---

## D4 — Inventory add-point is supplier-level only (not per-part)  *(2026-06-12)*

**Resolves:** the "add-point coupling" §8 question in `parts-stock-master` (§8.5) and
`inv-mgmt` (§8.5).

**Decision (verbatim intent):** *"The add point is supplier based only."*

**What this means for the rebuild:**
- **`VC_INVENTORY_ADD_POINT` stays an attribute of the supplier**, not the part. All of a
  supplier's parts share one add rule — `S` = add stock at shipping, `A` = add at arrival. It is
  **not** moved onto the part; parts do not carry their own add-point. The legacy coupling (a
  part's receiving-qty behavior is read from its supplier) is therefore **intended and preserved**.
- **Implication:** a part with a NULL/blank `IN_SUPPLIER_ID`, or a supplier with a blank/invalid
  add-point, has no add rule and so **stock silently never increments on receipt**. Since the rule
  lives only on the supplier, the rebuild should make the supplier's add-point a **required, valid
  value (`S`/`A`)** and require a part to have a supplier, so this can't silently happen.
  *(Placement is decided; the require-valid-value enforcement is the recommended implementation —
  confirm during the supplier-model build.)*

> Closes the add-point question. Answer: **supplier-level only; not per-part.**

---

## D5 — Stocktaking quantity is a signed adjustment delta (not an absolute count)  *(2026-06-12)*

**Resolves:** `stocktaking` §8.1 — the "single most important domain check."

**Decision:** the stocktaking `IN_QTY` an operator enters is a **signed adjustment delta** — the
triggers **add/subtract** it from on-hand. It is **not** an absolute counted total. Entering `100`
**raises** on-hand by 100; entering `-30` lowers it by 30. The legacy trigger behavior is the
**intended** behavior and is preserved.

**What this means for the rebuild:**
- The stock-ledger service applies stocktaking rows as **deltas** to on-hand (`IN_QTY` += entered
  value), consistent with `DailyBuildTotal`'s negative "Auto Scrap Delete" rows.
- **UI must make "delta, not total" unambiguous** — label the field as an adjustment (+/−), so an
  operator never mistakes it for "set on-hand to this counted number." (If a true *physical-count →
  set absolute* workflow is ever wanted, that is a **separate** feature that computes the delta for
  the operator; it is not what stocktaking does today.)

> Closes the delta-vs-absolute check. Answer: **signed adjustment delta.**

---

## D6 — Manifest-cost pricing is genuinely time-bounded; the legacy invoice/810 procs are buggy  *(2026-06-12)*

**Resolves:** `manifest-cost` §8.1 (chooses option **b**), and consequently §8.2 (duplicate/overlap)
and the `start > end` half of §8.7.

**Decision:** assembly prices are **genuinely time-bounded** — the `start_manifest`/`end_manifest`
window is real and meaningful. Because every current billing consumer **ignores** the window and
joins on assy code only, **the legacy invoice/810 procs are confirmed buggy** and must be fixed in
the rebuild.

**What this means for the rebuild:**
- **Billing must be window-aware.** The price for an invoice/810 line is the manifest-cost row whose
  `[start_manifest, end_manifest]` window **contains the ASN production date** (not just any row with
  the matching assy code). This is the fix for invoice correctness. The billing read path
  (`SELECT_INVOICEItems`, `REPORT_EDI810*`), owned by the Invoice/EDI module, must implement this
  window filter — flag it when that module is analyzed.
- **No-overlapping-window constraint (resolves §8.2).** The rebuild enforces **unique
  non-overlapping windows per `(site_id, VC_ASSY_PART_NUMBER_CODE)`** — NOT a single unique code.
  Two prices for the same assy code are allowed *only* if their windows don't overlap; this prevents
  the doubled-invoice-line hazard while supporting price changes over time.
- **Reject `start > end` (resolves the §8.7 window half).** With windows real, a row where
  `start_manifest > end_manifest` is invalid and must be **rejected** (legacy accepted it silently).
  *(The negative/zero-price half of §8.7 is separate and still open.)*
- **Interacts with D3:** blocking delete of a referenced price still holds; superseding an old price
  is done by adding a new non-overlapping window (and/or archival), not by editing/deleting in place.

> Closes the most important billing question. Answer: **time-bounded is real; fix the window-blind
> billing; enforce non-overlapping windows per (site, assy code); reject start > end.**

---

## D7 — The `'A'`-supplier arrival stock-add happens in Receiving Confirmation (RecConfStat)  *(2026-06-12)*

**Resolves:** `logistics-breakdown` §8.2 — "where does the `'A'`-supplier arrival stock-add happen?"

**Decision (confirmed against code):** the **Receiving Confirmation (`RecConfStat`)** screen is the
arrival path. Its arrival-date field (`RecConfStat.pas`) feeds `Arrival` into
`UPDATE_RecConfStatInfo` / `UPDATE_RecConfStatRenbanInfo` (`DataModule.pas:3346` / `:3269`), which
`SET VC_ARRIVAL = @Arrival` on `INV_OPEN_ORDER_INF` — and that stamp is the **only** thing that fires
the qty-trigger's arrival-add branch for `VC_INVENTORY_ADD_POINT = 'A'` suppliers.

**What this means for the rebuild:**
- **Two distinct stock-moving events, two modules:** the carrier/logistics feed sets `INTRANSIT`
  → **shipping-add** for `'S'` suppliers; **RecConfStat** sets the arrival date → **arrival-add**
  for `'A'` suppliers. (Guard confirmed at `RecConfStat.pas:818`: "Order must be marked In Transit
  when arrival is set" — INTRANSIT precedes arrival.)
- For `'A'` parts, the carrier feed (logistics-breakdown) records arrival **status only** and does
  **not** count stock — by design. Stock for `'A'` parts is counted exclusively by the RecConfStat
  arrival stamp.
- **Re-homing the trigger:** the arrival-add belongs to the **receiving-confirmation** action in the
  rebuilt stock-ledger service, keyed off the confirmed arrival — not the carrier-feed ingest.
  (Captured here so the future Receiving-module analysis owns it.)

> Closes the arrival-path question. Answer: **RecConfStat stamps `VC_ARRIVAL`; that is the
> `'A'`-supplier arrival-add path; the carrier feed only records arrival status for `'A'` parts.**

---

## D8 — Three confirmed legacy bugs the rebuild fixes (do not preserve)  *(2026-06-12)*

**Resolves:** the "confirm-and-fix" §8 questions — `size` §8.2, `stocktaking` §8.3,
`logistics-breakdown` §8.3. All three verified against source this session; the rebuild fixes them.

**Bug 1 — Size duplicate-check queries the wrong table.** `DataModule.pas:2531` `InsertSizeInfo`
runs its app-side duplicate check via `SELECT_AssyRatioInfo` (`DataModule.pas:2543`), which filters
`INV_ASSY_RATIO_MST.VC_BROADCAST_CODE = @SizeCode` — the **assy-ratio broadcast codes, not the size
master** — then inserts only `If RecordCount = 0`. Effects: (a) a genuine duplicate size code is
**never** caught app-side (only a DB unique index would); (b) a size code that coincides with an
existing broadcast code is **falsely rejected** as a duplicate.
*Fix:* check duplicates against `INV_SIZE_MST.VC_SIZE_CODE`, and enforce a DB unique index
`(site_id, VC_SIZE_CODE)` (per D1/D2). Confirmed bug.

**Bug 2 — Stocktaking edit blanks the audit timestamp.** `UPDATE_StockTakingInfo` (schema:9407)
builds `@Update` from `SUBSTRING(@Update, 1, 8)` where `@Update` was just `DECLARE`d and **never
initialized** → NULL; NULL string-concat makes the whole value NULL, so it writes
`VC_LAST_UPDATE = NULL` on the stocktaking row. The `UPDATE_Stocktaking` trigger (schema:10441) then
copies that NULL onto the affected `INV_PARTS_STOCK_MST.VC_LAST_UPDATE` too. The date portion should
have been `CONVERT(char(8), getdate(), 112)` (exactly how `INSERT_SizeInfo` does it correctly),
yielding a 16-char `yyyymmddHHMMSSff` stamp.
*Fix:* write a correct 16-char timestamp on every stocktaking edit (and let the re-balance carry it).
Confirmed bug.

**Bug 3 — Arrival-reversal branch is dead code; the rebuild IMPLEMENTS the reversal.**
`UPDATE_RecConfStatPartsStockMstQTY` (schema:9764) "changed to not arrived" branch has
`WHERE i.VC_ARRIVAL = '' AND i.VC_ARRIVAL <> ''` — both clauses on `i`, a contradiction that is
**always false**, so the branch never runs. The arrival-add branch above it is
`i.VC_ARRIVAL <> d.VC_ARRIVAL AND d.VC_ARRIVAL = ''`; the correct mirror is
`i.VC_ARRIVAL <> d.VC_ARRIVAL AND i.VC_ARRIVAL = ''`. Today, clearing a set arrival for an `'A'`
supplier does **not** reverse the previously-added stock (on-hand overstated).
**DECISION (David):** **implement the reversal** — when an `'A'`-supplier arrival is cleared, the
rebuild's stock-ledger posts a compensating **−qty** (the corrected mirror). Per **D7**, this lives
in the **receiving-confirmation** action alongside the arrival-add, not the carrier-feed ingest.

> Closes the confirm-and-fix batch. All three are real defects; the rebuild fixes them (Bug 3 by
> implementing the intended arrival-reversal).

---

## D9 — The live server dump `CreateInventory.sql` (2026-06-12) is the authoritative schema; resolution of the "snapshot-drift" findings

*Recorded 2026-06-16.*

**Verbatim intent (David):** "The dump is live." `DB Schema/CreateInventory.sql` (no space, dated
2026-06-12) is the **current live-server schema** — 182 procs / 42 tables / 25 triggers. The older
`Create Inventory.sql` (spaced, 2026-06-01) was a stale snapshot, now renamed
`Create Inventory.superseded-2026-06-01.sql` (retained only so pre-2026-06-16 specs' `schema:NNNN`
line cites still resolve). Also: **CAMEX is a decommissioned site** (as is NUMMI); their reports are
deprecated relics, out of rebuild scope.

**What this means for the rebuild.** Many analysis specs flagged proc-signature mismatches / missing
procs as "vs the checked-in snapshot — verify live" ([[reference-schema-snapshot-vs-live]]). Re-verified
against the LIVE dump, each finding now has a concrete verdict:

| Finding | Spec | Verdict vs LIVE dump |
|---|---|---|
| EDI `@EIN` on `REPORT_EDI810` + `REPORT_EDI856` (D1 cross-site-bleed risk) | edi/asn-invoice | **RESOLVED** — both procs declare `@EIN INTEGER=0` and use it (`IF @EIN=0`→all, else `WHERE …_EIN=@EIN`). Passing `@EIN` correctly scopes to one site. The D1-blocking concern is GONE. |
| `DELETE_ForecastInfo` (was "missing") | forecasting/forecast-breakdown | **RESOLVED** — exists, 3 params (`@WeekDate,@HistWeekDate,@PartNumber`) matching the caller exactly. |
| `INV_FORECAST_DETAIL_INF` label/misc columns | forecasting/forecast-detail | **RESOLVED** — live table HAS `VC_LABEL_PART_NUMBER`/`VC_MISC1_PART_NUMBER`/`VC_MISC2_PART_NUMBER`; CRUD procs reference them. |
| `UPDATE_UserPassword` (was "missing") | admin/auth-users | **RESOLVED on existence** — exists, 2 params (`@UserID,@NewPass`). **But NEW REAL mismatch:** caller passes `@Password` (DataModule.pas:6310) ≠ proc `@NewPass` → by-name ADO bind fails. |
| Shipping M1 `INSERT_ShippingDetail` | shipping/shipping | **REAL** — live declares 4 params (`@PartShipID,@PartNumber,@Productiondate,@Qty`); caller passes 5 (diff names). |
| Shipping M2 `INSERT_StockTakingInfo` | shipping/dailybuildtotal | **REAL** — live 3 params (`@PartNumber,@QTY,@Reason`); caller passes 5. |
| Shipping M3 `INSERT_ShippingInfo` | shipping/shipping | **REAL** — live 9 params (incl. `@ShippingID OUTPUT,@DTStartSeq,@DTEndSeq`); caller passes 6. |
| `SELECT_PartsStockInfo` (auto-scrap) | shipping/dailybuildtotal | **REAL** — live still 1 param (`@PartNum`); caller passes 3 and reads a `'Last Scrap Count'` column the proc doesn't return → `FieldByName` raises. |
| `REPORT_ASNWithCost`, `REPORT_ForecastCAMEXReport` | reporting, forecasting | **DEPRECATED** — absent from live; CAMEX decommissioned. Not bugs. |
| `REPORT_NUMMILotLocation[W]`, `ForecastCamexreport.pas` | reporting | **DEPRECATED RELIC** — NUMMI/CAMEX decommissioned sites; out of scope. |

**NEW findings surfaced by the live dump (real, for the build):**
- **Hardcoded site EIN in `REPORT_EDI856`:** one branch has `WHERE a.IN_ASN_EIN = 6440` (live `:3683`) —
  a literal site EIN baked into the proc. A D1 hazard (pins one site); the rebuild must parameterize it.
- **`UPDATE_UserPassword` param-name mismatch:** caller `@Password` vs proc `@NewPass` — reconcile in the
  auth rebuild (moot once auth moves to the Ignition User Source).

**Net:** the four scariest "snapshot-drift" items (EDI cross-site, the three missing procs) are RESOLVED;
the genuine residue is the **Shipping signature mismatches (M1/M2/M3 + SELECT_PartsStockInfo)** — these are
REAL vs the live schema and confirm the ManualShipping / daily-pull / auto-scrap paths are broken in the
deployed code (a real defect to fix in the rebuild, not a snapshot artifact). Resolves the
"verify-live / snapshot-drift" §8 items across shipping, edi, forecasting, admin specs.

---

## D10 — Forecast week numbering is TEMA-supplied and production-relative (ISO − 1 for 2026); the Order R1 fix is validated against a real 830

*Recorded 2026-06-16. Resolves forecast-breakdown.md §8 Q1 (the highest-risk forecast question) and
confirms the shipped `SIM_OrderSimulation` R1 fix.*

**Evidence:** David provided a real TMMMS 830 forecast feed (`EDI/830000008976.EDI`, gitignored client
data — sender Toyota DUNS `808369495`, receiver Magnolia `71930`, horizon 5/08–7/31, generated 2026-04-27).
Its FST (Forecast Schedule) segments carry the week number in the **FST09 "DO" reference number**, e.g.:

```
FST*144*D*W*20260615*20260619***DO*2624   → production week of 6/15, DO ref 2624
```

The legacy parser reads `week = copy(delSL[9], 3, 2)` = chars 3-4 of the DO ref → for `2624` → **`24`**.
The DO ref is structured `2` + `6` (year 2026) + `WW` (week), so `2624` = "2026, week 24".

**Measured across the whole horizon, `ISO_week(start_date) − TEMA_DO_week = 1` for every normal
production week** (the lone exception, the 7/12 single-day `FST*0…*2628`, is the mid-July shutdown stub —
zero qty, off-cycle):

| FST start | DO ref | TEMA wk | ISO wk | ISO − TEMA |
|---|---|---|---|---|
| 2026-06-08 | 2623 | 23 | 24 | 1 |
| 2026-06-15 | 2624 | 24 | 25 | 1 |
| 2026-06-22 | 2625 | 25 | 26 | 1 |

(…holds for all 12 normal weeks 18→30.)

**What this means for the rebuild.**
1. **The forecast week number is TEMA-supplied, NOT app-computed.** It is parsed verbatim from the 830's
   FST09 DO reference (chars 3-4) and stored as `INV_BREAKDOWN_FC_INF.IN_WEEK_NUMBER`. The forecast WRITE
   side stores it unmodified (`checkweeknumber`); the FirstProductionDay offset is applied only to a local
   holiday-lookup variable, never to the row (per the Forecasting spec).
2. **TEMA's numbering is production-relative, running exactly `ISO − 1` for 2026** — which equals
   `weekoffset = INT_FIRST_PRODUCTION_WEEK[2026] − 1 = 2 − 1 = 1` (the value `INV_FIRST_PRODUCTION_DAY`
   carries; see [[project-order-renban-domain]] / the Production-calendar spec).
3. **The shipped Order R1 fix (`SIM_OrderSimulation` STEP 4) is VALIDATED.** The READ side computes
   `@WeekNo = ISO_week(prodDate) − weekoffset = ISO − 1`, and the stored `IN_WEEK_NUMBER` = the TEMA DO
   week = `ISO − 1`. They match exactly on real data. The "three week-number conventions coexist; consistency
   hinges on the 830 being production-relative" risk (Forecasting §8 Q1) is **confirmed safe** — the feed IS
   production-relative.
4. **Rebuild rule:** ingest `IN_WEEK_NUMBER` from FST09 chars 3-4 verbatim; do not recompute from the date.
   When the read side needs a week for a production date, use `ISO_week(date) − (INT_FIRST_PRODUCTION_WEEK[year] − 1)`.
   The "year + week" DO encoding (`26` + `WW`) is the durable key; carry both year and week, not just week,
   to disambiguate across year boundaries (the legacy stores week-only and relies on delete-forward each cycle).

---

## D11 — Confirmed-bug batch: the rebuild FIXES these (David, 2026-06-16)

*Resolves the "confirmed bug" §8 items across the Receiving/Shipping/Forecasting/EDI/Reporting/
Production-calendar/Admin specs. The D8 pattern: each is a verified-against-source defect; the rebuild
implements the corrected behavior (the legacy is NOT hotfixed — see the P12 NO-legacy-hotfix policy).*

**Decision (verbatim intent):** *"Confirm Group B, fix in the rebuild."*

**The batch (all confirmed against the live `CreateInventory.sql` and the `.pas`):**

1. **D6 window-aware pricing — everywhere.** Manifest-cost pricing must pick the price whose
   `VC_START_MANIFEST`/`VC_END_MANIFEST` window contains the ASN/production date, in **all** instances:
   the EDI 810 build (`REPORT_EDI810`/`SELECT_INVOICEItems`) and the Reporting invoice summaries
   (`REPORT_INVOICESSummary`, `REPORT_MonthlyINVOICESSummary`). Copy the correct `REPORT_EDI856` predicate.
   Enforce non-overlapping windows per (site, assy code); reject `start>end`. (Extends D6 to the Reporting
   instances found 2026-06-16.)
2. **`REPORT_UnusedWheelPartNumbers`** queries the TIRE part-number column for wheel parts
   (schema `…UnusedWheelPartNumbers` `NOT IN (SELECT vc_tire_part_number_code…)`) → use the WHEEL column.
3. **Forecast day-spread** for valve/film/label/misc uses `wheelcount` (`ForecastBreakdownF.pas:1252-1285`)
   instead of each component's own count → spread each component on its own count.
4. **Shipping proc-signature mismatches (M1/M2/M3 + `SELECT_PartsStockInfo`)** — REAL vs the live schema
   (D9): the ManualShipping / daily-ALC-pull / auto-scrap paths are broken in deployed code. The rebuild
   uses correct, reconciled signatures (one canonical Named Query per op).
5. **Hardcoded `WHERE a.IN_ASN_EIN = 6440`** in `REPORT_EDI856` (live `CreateInventory.sql:3683`) →
   parameterize by the current site's EIN (D1).
6. **DATAPURGE non-transactional `PurgeMode`** (`DELETE_AutoPurge`) — a mid-run error leaves `PurgeMode=1`
   + a partial delete. The rebuild wraps the purge in a transaction (set/clear the flag atomically) and
   re-homes the cross-DB `Activity` audit coupling.
7. **RenbanGroup counter read-then-write race** — `UPDATE_RenbanGroupCount` + the client-side
   `Format('%.3d',…)` count (RenbanOrder/RenbanGroupMaster) → atomic, by-id increment in the rebuild.
8. **`INV_FIRST_PRODUCTION_DAY`** has no PK and `INSERT_FirstProductionDay` never dedups (the form's
   "already exists" message is fiction) → real PK on `(site_id, production_year)` + a true upsert.
9. **Reject-delete inflates on-hand** — `DELETE_RejectParts` has no purge bypass (unlike RecConfStat), so
   purging a reject row adds its qty back. The rebuild's single stock-ledger service handles reject
   reversal correctly (a reject delete is not a stock movement).

> All nine are fixed structurally in the rebuild. Where a fix lands in the re-homed **stock-ledger**
> service (4 partially, 9), it composes with the additive-delta ledger model. D6 (1) shares one
> window-aware manifest-cost lookup across EDI + Reporting.

---

## D12 — Group C domain-judgment answers (David, 2026-06-16)

*Resolves the remaining genuine-domain §8 questions across Assembly, EDI, Receiving, and Reporting.
This closes the §8 decisions pass for the InventorySystem analysis (D1–D12).*

**1. Drop `INV_ASSY_RATIO_MST` from the rebuild.** *(Assembly — assy-ratio-master §8.1/§8.2)*
Verbatim intent: *"INV_ASSY_RATIO_MST failed conversion thought, drop for rebuild."* The table + its
AssyRatioMaster screen were an abandoned design (the screen is hidden "not used yet"; no forecast/order
proc reads the table — the live explosion uses `INV_FORECAST_DETAIL_INF`). **Do NOT port** `INV_ASSY_RATIO_MST`,
`AssyRatioMaster`, or `BCRatioMaster` (already dead). The broadcast→part ratio model lives entirely in
`INV_FORECAST_DETAIL_INF`.

**2. EDI 820 remittance stays report-only.** *(EDI — edi-upload §8)*
Verbatim intent: *"Keep report only now, site isn't using it."* The legacy 820 path's "Store in Table"
TODO stays unimplemented; the rebuild renders 820 as a report (no payment-application persistence) for now.
Revisit if a site starts using remittance reconciliation. (Still fix the latent parse bugs — `SE*` EOF
loop, the `TStringList` leak — if/when the report is rebuilt.)

**3. Plant-yard AND assembler-yard count as arrival on edit.** *(Receiving — recconfstat §8.5)*
Verbatim intent: *"Plant/yard both count as arrival."* The legacy UPDATE arrival-add leg fires only on
`VC_ARRIVAL`, so stamping plant-yard/assembler-yard on an EDIT of an `'A'` order under-counts vs the
INSERT/DELETE legs (which treat plant-yard/assembler-yard/warehouse as arrival-equivalents). The rebuild's
receiving action treats **plant-yard and assembler-yard as arrival-equivalent on edit too** — symmetric
with insert — so the `'A'`-supplier stock-add fires consistently however arrival status is recorded. (Folds
into the stock-ledger service alongside D7 + the D8(3) reversal.)

**4. Monthly order reports range on ORDER date.** *(Reporting — reporting §8.2)*
Verbatim intent: *"Should range on order date."* The legacy monthly supplier/logistics ORDER reports
(`REPORT_MonthlySupplierOrders`/`…Cost`/`REPORT_MonthlyLogisticsOrders`) filter on
`VC_STATUS_SUPPLIER_SHIPPING` (ship date), which is wrong — an order placed in month M-1 and shipped in M
shows in M. The rebuild ranges these on **`VC_ORDER_DATE`** instead. (Confirmed bug; behaves like a D11
fix. The daily order reports already use `VC_ORDER_DATE` correctly. Invoice reports are unaffected.)

> Closes the InventorySystem §8 decisions pass. D1–D12 cover the cross-cutting decisions + every spec's
> confirmed-bug and domain-judgment open items. Remaining spec §8 entries are narrow "verify during build"
> notes, not decisions.

---

## D13 — Assembly Detail master = INV_FORECAST_DETAIL_INF; ManifestCost relaxes the manifest-number unique quirk (option b); rollout needs a DB-diff conversion script

*Recorded 2026-06-17. Two build decisions for the remaining master forms + a cross-cutting rollout note.*

**1. "Assembly Detail" = the FORM name over `INV_FORECAST_DETAIL_INF` (not a table rename).**
Verbatim intent (David): *"I'm thinking in the name of the form not name of the table. The name was
changed on the form post go-live, the table change was dropped."* So the live BOM/ratio master — assy
part code → tire/wheel/valve/film/label/misc part codes + ratios, effective month, broadcast code — is
maintained by a form now titled **"Assembly Detail"** but still backed by `INV_FORECAST_DETAIL_INF`
(table never renamed). This is the master the forecast/order explosion actually reads (the separate
`INV_ASSY_RATIO_MST` was DROPPED per D12). The rebuild builds an "Assembly Detail" CRUD over
`INV_FORECAST_DETAIL_INF`. Spec: `docs/analysis/forecasting/forecast-detail.md`.

**2. ManifestCost — option (b): correct the legacy manifest-number-uniqueness quirk.**
Verbatim intent (David): *"use b, correct a legacy quirk from a developed-on-the-fly feature that mostly
works."* The live `IX_INV_MANIFEST_COST_MST` UNIQUE on `VC_ASSY_MANIFEST_NUMBER` (global manifest-number
uniqueness) **conflicts with D6's "multiple cost windows per assy code"** and is an artifact of an on-the-fly
feature. The rebuild **drops/relaxes** that global unique (→ allow multiple windows per assy code; the real
constraint is **D6 non-overlapping `VC_START/END_MANIFEST` windows per (site, assy code)** + `start<=end`).
So the ManifestCost master uses **`checkWindowOverlap`**, NOT `checkManifestNumberUnique`. (Supersedes the
"default: honor the live index" placeholder in IGNITION-master-crud-design.md §C / R2.)

**3. ROLLOUT (cross-cutting) — a solid DB-diff conversion script is required.**
Verbatim intent (David): *"ensure we have notes for rollout that there will need to be a solid conversion
script to support these DB diffs."* The rebuild intentionally diverges from the live legacy schema in
several places that a production cutover must reconcile with a migration/conversion script, NOT silently:
- **ManifestCost (D13.2):** drop the `IX_INV_MANIFEST_COST_MST` global-unique index; add the D6 per-(site,
  assy) non-overlapping-window constraint/check.
- **D1 multi-site:** add `site_id` (NOT NULL FK) to every master table; replace each single-column UNIQUE
  (`IX_INV_SUPPLIER_MST`/`_SIZE_`/`_LOGISTICS_`/`_RENBAN_GROUP_`/`_PARTS_STOCK_`) with a per-site composite
  `(site_id, <code>)`; flip every `-- IG-SITE:` predicate on in the same ordered migration (R3).
- **Audit columns:** the rebuild keeps the 16-char `yyyymmddHHMMSSff` string form during parallel run;
  the Postgres phase converts to real datetime DEFAULT/triggers (every `# IG83-TODO`).
- Assembly Detail / `INV_FORECAST_DETAIL_INF` (confirmed during build, 2026-06-17): the live table has
  **NO PK and NO unique index** — the rebuild's composite app-check `(VC_ASSY_PART_NUMBER_CODE,
  VC_EFFECTIVE_MONTH)` is the ONLY uniqueness today (a concurrent double-insert could slip a dup). Rollout
  must (i) add a real PK on `ID_FORECAST_DETAIL`, (ii) add `UNIQUE(VC_ASSY_PART_NUMBER_CODE,
  VC_EFFECTIVE_MONTH)` (per-site composite under D1) as the backstop, and (iii) **pin + validate a canonical
  `VC_EFFECTIVE_MONTH` format** (e.g. `yyyy/mm`) — all 50 live rows are blank today, so `2026/01` vs `202601`
  would read as distinct composites and both insert, giving the unique index false confidence until the
  format is enforced.
**Action:** maintain a running "DB conversion / cutover script" deliverable enumerating every such diff so the
production migration is deterministic and reviewable. Track it as the masters + later modules land.
