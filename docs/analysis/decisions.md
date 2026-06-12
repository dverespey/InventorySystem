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
