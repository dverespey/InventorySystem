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
