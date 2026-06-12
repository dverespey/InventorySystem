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
