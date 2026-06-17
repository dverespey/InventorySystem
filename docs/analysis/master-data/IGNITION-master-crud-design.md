# Ignition Architecture — Master-Data CRUD Area

**Status:** architecture set, ready for `ignition-developer` build
**Author:** ignition-architect / 2026-06-16
**Supersedes** the Rails-flavored §6 in `supplier.md` / `size.md` / `logistics.md` / `manifest-cost.md`
(those predate the Ignition decision — see `project-ignition-eval`). This is the Ignition §6 for the
master-data area.

This is the first real rebuild module after the Order keystone spike and the PartsStockMaster codegen
proof. It delivers **(§A)** a reusable master-CRUD pattern, **(§B)** a concrete Supplier build spec
(the worked example / depth reference), and **(§C)** the deltas each later master applies to the template.

Ground truth was read off the live spike DB (`Inventory` on `mssql-spike`), not assumed:
- `INV_SUPPLIER_MST` has **no `site_id` column** (`COL_LENGTH(...,'site_id')` → NULL).
- `sites` table exists, 2 rows: `(1,'MAS','TMMMS')`, `(2,'HERO','TMMTX')` (seeded by `spike-db.sh`).
- `IX_INV_SUPPLIER_MST` is a **single-column** UNIQUE index on `VC_SUPPLIER_CODE` (not yet composite).
- The five procs (`SELECT/INSERT/UPDATE/DELETE_SupplierInfo`, `SELECT_PartsSupplier`) all exist.
- **`DELETE_SupplierCode` trigger (live body, verified 2026-06-16) does THREE things, not one:**
  (1) `UPDATE INV_PARTS_STOCK_MST SET IN_SUPPLIER_ID=null WHERE ... = DELETED` (the nullify we knew about),
  (2) `DELETE FROM INV_BREAKDOWN_FC_INF WHERE VC_SUPPLIER_CODE=(SELECT VC_SUPPLIER_CODE from DELETED)`,
  (3) `DELETE FROM INV_FORECAST_INF WHERE VC_SUPPLIER_CODE=(SELECT VC_SUPPLIER_CODE from DELETED)`.
  Both forecast tables are populated and keyed by supplier **CODE** (not id): `INV_BREAKDOWN_FC_INF` ≈959
  rows, `INV_FORECAST_INF` ≈1066 rows. **The supplier spec §2 under-documented this trigger** (it described
  only the nullify branch) — *spec correction needed: §2 must list the two forecast hard-deletes.* This is
  load-bearing for the delete gate (A.5 / B.5): a refCount that counts only `INV_PARTS_STOCK_MST` is **blind**
  to the forecast cascade.
- `INSERT_SupplierInfo` resolves logistics **by name**, narrows `@SupPerson`→25 / `@SupDirectory`→215,
  computes the 16-char `yyyymmddHHMMSSff` timestamp inline, and **takes no `site_id`**.
- FK lookup tables confirmed: `INV_LOGISTICS_MST(IN_LOGISTICS_ID,VC_LOGISTICS_NAME)` (genuine surrogate FK),
  `INV_PART_TYPE_MST(VC_PART_TYPE,...)` (Create-Order-Sheet combo, code-valued). `INV_ADD_POINT_INF` exists
  but is **not** used — Inventory-Add-Point is the static S/A enum (D4 / R6), no NQ.

---

## §A. The reusable master-CRUD pattern

### A.0 The three architecture decisions (answers up front)

**1. site_id timing — build single-site-NOW, design D1-shaped, with site as a real binding param.**
Do **not** add `site_id` to `INV_SUPPLIER_MST` (or the other master tables) in the spike DB. Rationale:

- *Parallel-run safety (the hard constraint).* During parallel run the legacy Delphi app and the rebuild
  hit the **same** SQL Server. The legacy supplier procs `SELECT *`-shape and positionally `INSERT` into
  `INV_SUPPLIER_MST`; the legacy `DELETE_SupplierCode` trigger does `... FROM INV_PARTS_STOCK_MST a, DELETED d`.
  Adding `site_id` to a live master is the exact `_HIST`-trigger hazard already flagged in `spike-db.sh`
  (F1) — it can break legacy `INSERT INTO _HIST SELECT * FROM deleted` paths unless every dependent
  table mirrors the column. D1 itself phases `site_id` to the **Postgres/DB-modernization phase**, *after*
  parallel run, precisely for this reason. So the spike/parallel-run build stays on the un-`site_id`'d schema.
- *But the D1 shape must be real in the design, not deferred.* The lever that makes this work: **a
  gateway-resolved "current site" that is a real bound parameter on every Named Query and every view**,
  even while there is physically one site of data. Concretely:
  - The `sites` table already exists in the spike DB (`spike-db.sh`). Keep it. It is the seam.
  - Every list/get/insert/update NQ takes a `site_id` param. For the **current** un-`site_id`'d tables,
    the param is **accepted and validated but not yet applied to a WHERE** — guarded by a single
    greppable marker so the Postgres phase turns it on in one place per NQ:
    ```sql
    -- IG-SITE: when INV_SUPPLIER_MST gains site_id (D1 / Postgres phase), uncomment:
    -- AND s.site_id = :site_id
    ```
  - The **uniqueness check NQ** (A.4) is written composite-ready the same way: today it checks
    `VC_SUPPLIER_CODE` alone (matching the live single-column `IX_INV_SUPPLIER_MST`); the `:site_id`
    predicate is the commented `IG-SITE` line, flipped on with the composite index in the Postgres phase.
  - **R3 — flip ordering is ONE atomic, ordered migration step (uncomment AFTER index is composite).**
    Uncommenting the `-- AND s.site_id = :site_id` predicate on `checkCodeUnique` *before*
    `IX_INV_SUPPLIER_MST` is rebuilt as composite `(site_id, VC_SUPPLIER_CODE)` would let the SAME code be
    inserted under two sites with the single-column unique index still rejecting it — the validator and the
    index would fight. **Required order within the single Postgres-phase migration:** (a) add `site_id`
    column + backfill; (b) DROP single-column `IX_INV_SUPPLIER_MST`, CREATE composite
    `(site_id, VC_SUPPLIER_CODE)` UNIQUE; (c) **only then** uncomment every `-- IG-SITE:` predicate
    (list/get/update/delete/checkCodeUnique) in one pass. (a)→(b)→(c), atomic, never (c) before (b).
  - **`site_id` is sourced server-side, never from the client.** It comes from
    `session.custom.siteId`. **R4 — where it is set:** a **gateway login / authentication event** writes
    `session.custom.siteId` from the authenticated user's site binding on the Ignition User Source (`D1`:
    "auth binds each user to a site"). It is a **session property set server-side at auth time** — it is
    **never written from a client binding, view script, or NQ param**, so it cannot be spoofed by a client.
    A view passes `{session.custom.siteId}` into NQ params; the NQ never trusts a client-supplied site. This
    is the InventorySystem analogue of the GALC `siteScopedQuery()` rule, and it is the actual tenancy
    boundary. **This auth→session wiring must be implemented and verified before any `-- IG-SITE:` predicate
    is flipped on** — flipping the predicate on while `siteId` is still a defaulted/client-writable value
    would be a false isolation boundary.
  - For the spike, `session.custom.siteId` defaults to `1` (TMMMS). Switching it to `2` exercises the
    multi-site code path against one-site data — proving the param *plumbs through*, not isolation.

  Net (R5 — honest scope): **one site of data; the multi-site SHAPE is real** (site flows session → view →
  NQ param; turning it on is a per-NQ uncomment + composite-index migration per R3, not a redesign). What is
  **NOT** yet proven is real data isolation — the spike only smoke-tests that the `siteId` param plumbs
  end-to-end; true isolation (predicate enforced against site-tagged rows, verified server-side auth source)
  is **deferred to the Postgres phase**. One honest partial today: `spike-db.sh` already adds `site_id` to
  `INV_PARTS_STOCK_MST` (and its `_HIST`) and seeds 15 site-2 rows, so a **parts-level** isolation test is
  possible now; **supplier-level** isolation is not (no `site_id` on `INV_SUPPLIER_MST`). Don't overclaim
  "multi-site is structurally proven" — say "param plumbs; parts-level partial isolation; supplier isolation
  deferred."

**2. Stage-1 wrap-the-proc vs Stage-3 direct CRUD — for masters, write DIRECT parameterized Named
Queries against the tables; do NOT wrap the legacy `*_SupplierInfo` procs.** This is a deliberate
divergence from the GALC "wrap the proc" default and from the Order spike's `runPrepQuery` proc-wraps,
and it is justified by what the master procs actually are:

- The master procs are **thin, app-flavored CRUD** (a positional INSERT, a keyed UPDATE, a keyed DELETE,
  a label-mapping SELECT) — not load-bearing business logic like the order/stock-ledger procs. There is
  almost nothing to preserve by wrapping them.
- They actively **fight three required decisions**:
  - **D2 (resolve by id):** `INSERT_SupplierInfo` resolves logistics *by name* (`WHERE VC_LOGISTICS_NAME
    = @SupLogistics`). D2 requires resolving by `IN_LOGISTICS_ID`. Wrapping forces a name round-trip we
    are explicitly told to remove.
  - **D1 (site_id param):** the procs take no `site_id`, so a wrap can't carry the current-site param —
    the D1 seam can't even exist through them.
  - **Truncation:** `@SupPerson varchar(25)` / `@SupDirectory varchar(215)` silently truncate below the
    table widths (50 / 512). A direct NQ writes the full column widths.
- The legacy **dup-check lives in the CLIENT** (the two-step insert: `SELECT_SupplierInfo @SupCode` then
  insert only `If RecordCount=0`). There is nothing in the proc to wrap for uniqueness anyway — we must
  reimplement it regardless (A.4). The real DB unique index `IX_INV_SUPPLIER_MST` is the backstop.

  > **Parity caveat (call out for the reviewer):** "direct NQ, not proc-wrap" means the rebuilt masters
  > are *behaviorally* parity-checked against the legacy (row counts, ordering, enum labels, dup
  > rejection, delete semantics) but do **not** route through the same proc objects. This is safe for the
  > masters specifically because they own no cross-module side effects beyond the delete trigger (handled
  > in A.5). It would NOT be safe for the order/stock-ledger procs — those still wrap. The boundary is:
  > **leaf-master CRUD = direct NQ; anything touching the stock ledger / EDI = wrap first.**

- **One thing we still preserve via the proc-shaped recipe:** the 16-char `yyyymmddHHMMSSff` audit
  timestamp. We inline the *same* T-SQL expression in the insert/update NQs (A.2) so legacy readers of
  `VC_ADD`/`VC_LASTUPDATE` during parallel run see byte-identical stamps. We do **not** normalize to real
  `datetime` until the Postgres phase (per every master spec's P2 note).

**3. The reusable shape: List view + Detail view, driven by a per-master Named Query set, combos from
FK-lookup NQs.** Detailed in A.1–A.6 below; it is a direct generalization of the proven
`PartsStockMaster` spike (List grid + Detail form, `gen_perspective_view.py`), with the data path moved
off inline `runPrepQuery`/`createSProcCall` and onto Named Queries (the `ignition-named-query-crud-practice`
memory: organize NQs to mirror schema/procs → schema change = single-point edit).

---

### A.1 View structure (the template)

Two Perspective views per master, under `views/<Master>/` (mirrors the spike's `PartsStockMaster/List`
+ `PartsStockMaster/Detail`):

```
views/<Master>/List      — searchable / sortable / paginated grid (ia.display.table)
views/<Master>/Detail    — form (one labelled input per editable column; combos for FKs)
```

**List view**
- Root: `ia.container.flex` (column). Children: a filter row (search `ia.input.text-field` + New button),
  the `ia.display.table`, a pager.
- `table.props.data` ← binding on the **list NQ** (`<Master>/list`), params `{searchTerm, siteId, sort}`.
  Server-side search/sort/pagination replaces the legacy client-side `SearchGrid`/`Filter` (every master
  spec's P7). The list NQ does the `WHERE ... LIKE` and `ORDER BY`, not the component.
- `table.props.selection` → on row select, navigate to `<Master>/Detail` passing
  `params.recordId = <selected RecordID>` (`system.perspective.navigate`). RecordID is the surrogate id
  (`IN_SUPPLIER_ID`), per D2 — never the business code.
- New button → navigate to `<Master>/Detail` with `recordId = 0` (new-record sentinel; mirrors the
  spike's `params.recordId` default of 0).

**Detail view** (the form)
- `view.params.recordId` (int, default 0) — `0` = insert mode, `>0` = edit mode.
- `view.custom.form` — a **plain unbound object**, one key per editable column + one per FK code/id
  alias. Pre-declared at startup (so input sub-path bindings resolve), then seeded by the **get NQ**.
  This is the spike's proven anti-blank pattern: bind inputs *bidirectionally into an unbound custom
  object*; do **not** bind inputs into a query-bound object (that orphans write-backs — the lesson the
  spike's comments call out explicitly).
- Each editable column → one labelled `ia.container.flex` row holding an input, by type:
  - `bit` → `ia.input.checkbox` (`props.selected`)
  - numeric (`int`/`money`/`decimal`/…) → `ia.input.numeric-entry-field` (`props.value`)
  - long text (len ≥ 300 or max) → `ia.input.text-area` (`props.text`)
  - else → `ia.input.text-field` (`props.text`)
  - single-char coded enum (`VC_OUTPUT_FILE`, `VC_INVENTORY_ADD_POINT`) → `ia.input.dropdown` with a
    **static** option list mapping code→label (A.6)
  - FK column → `ia.input.dropdown` sourced from an **FK-lookup NQ** (A.6), value = the surrogate id (D2)
  - audit columns (`VC_ADD`/`VC_LASTUPDATE`/`VC_LAST_UPDATE`) → **not rendered** (proc-maintained)
- Action row: **Save** (insert-or-update), **Delete** (edit mode only), **Cancel** (back to List).
  Save/Delete call the write NQs via `system.db.runNamedQuery` inside an `onActionPerformed` script,
  with validation (A.4) run first.

> The above is exactly what `gen_perspective_view.py` already emits, with three swaps the developer makes
> to lift it from spike to product: (a) FK dropdown options + load + save move from inline
> `runPrepQuery`/`createSProcCall` to **Named Queries**; (b) Save runs **validation** before the write;
> (c) every data path carries the `siteId` param. The codegen script is the starting point, not thrown away.

### A.2 Named Query set + naming (mirror the schema/procs)

> ⚠️ **MECHANISM CORRECTION (2026-06-16, build finding — overrides the "formal on-disk Named Query
> resources" assumption below).** A headless build CANNOT reliably author on-disk Named Query resources:
> the gateway parses a NQ's `data.bin` as **Ignition XML serialization**, and a hand-authored file fails
> with `SAXParseException: Content is not allowed in prolog` (proven — the first hand-built `lookups/partType`
> NQ would not deserialize; this is the exact "undocumented on-disk NQ format, do not hand-author blind"
> hazard called out earlier). Formal NQ resources require the **Designer** (interactive), which the headless
> fleet build doesn't have. **DELIVERY DECISION:** keep this §A.2 set as the **SQL source-of-truth** in a
> `.sql` doc (`docs/analysis/master-data/master-crud-namedqueries.sql`, mirroring the Order spike's
> `named-queries.sql`), and **execute each via inline `system.db.runPrepQuery(sql, args, "Inventory_Spike")`
> in the Perspective view's binding/script transforms** — the pattern PROVEN on this gateway by the Order
> spike. The folder/naming/param/site-seam design below is unchanged; only the runtime delivery is inline
> `runPrepQuery` instead of a `system.db.runNamedQuery` call against an on-disk resource. (When the project
> is later opened in the Designer, these can be promoted to true NQ resources for the single-point-edit
> benefit — a Designer task, not a headless one.) The broken `lookups/partType` NQ resource must be removed.
>
> *Folder/naming below is the logical organization of that SQL (NQ-style), not literal on-disk NQ resources.*

Per the `ignition-named-query-crud-practice` memory, organize the SQL in a structure that **mirrors the
table** and name each for the op it performs, so "which query touches this table" is a single lookup and a
schema change is a single-point edit. Logical set per master:

```
Named Queries/
  <Master>/                         e.g. Supplier/  (mirrors INV_SUPPLIER_MST)
    list        — grid rows; params: searchTerm, siteId, sort/page
    get         — one row by id (resolves FK ids AND codes for the form); param: recordId, siteId
    insert      — INSERT; returns new identity (SCOPE_IDENTITY()); all data params + siteId
    update      — UPDATE ... WHERE id = :recordId; all data params + siteId
    delete      — DELETE ... WHERE id = :recordId (after the RESTRICT pre-check, A.5)
    checkCodeUnique  — uniqueness pre-check (A.4); params: code, excludeId, siteId
    refCount    — D3 RESTRICT pre-check: count EVERY row the legacy delete-trigger would touch
                  (per master — for Supplier that is parts + both forecast tables); params: recordId, code  (A.5)
  lookups/
    logistics   — IN_LOGISTICS_ID + VC_LOGISTICS_NAME, site-scoped  (shared FK combo source)
    partType    — VC_PART_TYPE (Create-Order-Sheet combo; code-valued)
    -- R6: NO addPoint NQ. Inventory-Add-Point is the static S/A enum (D4 / A.6), not an NQ-driven
    --     combo, so the lookups/addPoint NQ is DROPPED as dead.
```

Conventions baked in for every NQ:
- **Every NQ that reads/writes a master table declares a `:siteId` param**, even where the predicate is
  still the commented `-- IG-SITE:` line (decision 1). This is the single-point seam.
- **All keys are the surrogate id** (`get`/`update`/`delete`/`refCount` key on `IN_*_ID`), per D2. The
  business code is a payload column, never a key, never a join target.
- `insert` returns the new id via `SELECT SCOPE_IDENTITY()` (the legacy procs never echoed it — a real
  improvement; the spike re-located by natural key, which is fragile).
- Audit columns are written **inside** insert/update with the preserved 16-char recipe:
  ```sql
  -- VC_ADD on insert; VC_LASTUPDATE on update. Byte-identical to the legacy procs (P2).
  CONVERT(char(8), GETDATE(), 112)
    + SUBSTRING(CONVERT(varchar, GETDATE(), 114), 1, 2)
    + SUBSTRING(CONVERT(varchar, GETDATE(), 114), 4, 2)
    + SUBSTRING(CONVERT(varchar, GETDATE(), 114), 7, 2)
    + SUBSTRING(CONVERT(varchar, GETDATE(), 114), 10,2)
  -- IG83-TODO: at the Postgres phase replace with a real datetime DEFAULT/trigger; drop the string form.
  ```
- NQ type: `list`/`get`/`lookups/*`/`checkCodeUnique`/`refCount` are **Query**; `insert`/`update`/`delete`
  are **Update** queries (so `runNamedQuery` returns the rowcount / the insert returns identity via a
  Query that ends in `SELECT SCOPE_IDENTITY()`).

### A.3 Param / binding flow

```
List:   table.data  ← NQ Supplier/list   params {searchTerm: TextField.text,
                                                  siteId: session.custom.siteId}
        rowSelect   → navigate Detail (recordId = row.RecordID)

Detail load (onStartup or recordId change):
        if recordId > 0:  rows = runNamedQuery("Supplier/get", {recordId, siteId})
                          for each column: view.custom.form[col] = rows[0][col]   (mutate in place)
        else:             leave form at its pre-declared null/default seed (insert mode)
        FK dropdowns: options ← runNamedQuery("lookups/logistics", {siteId}) (etc.)

Detail Save (onActionPerformed):
        rec = view.custom.form
        errs = validate(rec)                    # A.4 — client-side fast-fail
        if errs: show + return
        params = { ...map rec → NQ params..., siteId: session.custom.siteId }
        if recordId == 0:  newId = runNamedQuery("Supplier/insert", params)  # returns SCOPE_IDENTITY
                          view.params.recordId = newId                       # flip to edit mode
        else:              runNamedQuery("Supplier/update", {...params, recordId})
        # DB unique index is the backstop: catch the duplicate-key SQLException and surface
        # the same friendly message validate() would (A.4) — covers the race the pre-check can't.
        navigate back to List

Detail Delete:
        # A.5 D3 RESTRICT — refCount counts EVERY table the live trigger touches
        # (parts by id + both forecast tables by code), not just parts.
        n = runNamedQuery("Supplier/refCount", {recordId, code: form['VC_SUPPLIER_CODE']})
        if n > 0: show "Cannot delete — still referenced by N parts/forecast rows"; return
        runNamedQuery("Supplier/delete", {recordId}); navigate List
```

FK-combo value semantics (D2): the dropdown's **value is the surrogate id**, its **label is the code/
name**. The `get` NQ returns *both* the id (to set the combo value) and the code (for display elsewhere).
This is the one place the Supplier build diverges from the legacy proc, which returned only the logistics
*name* — so `get` must `LEFT JOIN` logistics to return `IN_LOGISTICS_ID` for the combo (B.2).

### A.4 Validation approach

Validation runs **client-side in the Save script first** (fast, friendly), with the **DB unique index as
the authoritative backstop** for the dup race. Rules, per master, declared as data so the template can
carry them:

| Rule | Where enforced | Notes |
|---|---|---|
| **5-char code** (Supplier) / fixed-length code | Save script: `len(code) == 5` (Supplier); ≤6 (Size) | Legacy form rejected `<5`; we reject `!= 5` (still ≤ varchar(5)). Size has no min in legacy — we **add** presence + ≤6 (a deliberate improvement, documented divergence). |
| **presence** (code, name) | Save script | Legacy Size/Logistics had no presence check; we add it (documented divergence). |
| **per-site uniqueness** of the code | `checkCodeUnique` NQ pre-check **+** `IX_*` unique index backstop | Pre-check: `SELECT COUNT(*) ... WHERE code=:code AND id<>:excludeId [AND site_id=:siteId]`. The `site_id` predicate is the commented `IG-SITE` line until the Postgres phase. Replaces the legacy client two-step insert. `excludeId` lets an UPDATE keep its own code (rename support, D2). |
| **enum domain** (`T/E/B`, `S/A`) | dropdown is `csDropDownList`-equivalent (static options) | Can't enter an out-of-domain value; no extra check needed. |
| **duplicate-key race** | catch SQLException on insert/update | The pre-check + index together replace the legacy app dup-check; the catch turns the raw DB error into the friendly message. |

> **Do NOT reproduce the Size dup-check bug (D8 Bug 1).** Legacy Size checked dups against
> `SELECT_AssyRatioInfo` (broadcast codes), so real size dups were never caught app-side. The template's
> `checkCodeUnique` checks the master's **own** code column. This is an intentional divergence (parity
> check asserts the *correct* behavior, not the legacy bug) — see Size delta in §C.

### A.5 D1 / D2 / D3 application (template-wide)

- **D1 (multi-site):** site flows session → view → NQ `:siteId` param; predicate is the `-- IG-SITE:`
  seam; uniqueness is composite-ready. One site of data in the spike, multi-site-shaped in the design.
- **D2 (surrogate id is the sole key):** RecordID = `IN_*_ID` everywhere; FK combos store ids, display
  codes; `get`/`update`/`delete` key on id; **renames are plain attribute updates** (the code is a
  non-key payload column with a per-site uniqueness constraint). The legacy name→id resolution in
  `INSERT_SupplierInfo` is replaced by passing the id directly.
- **D3 (RESTRICT on delete) vs the legacy `DELETE_SupplierCode` trigger — the reconciliation:**
  The live trigger does THREE things (verified 2026-06-16, see the ground-truth note above): it
  **NULLIFIES** `INV_PARTS_STOCK_MST.IN_SUPPLIER_ID` for the supplier's parts, **and hard-DELETEs** every
  `INV_BREAKDOWN_FC_INF` and `INV_FORECAST_INF` row whose `VC_SUPPLIER_CODE` matches the deleted supplier.
  **D3 overrides the legacy semantics: the rebuild BLOCKS the delete while ANY of those three tables still
  references the supplier.** The gate cannot just protect parts — it must see everything the trigger touches,
  or a supplier with zero parts but live forecast rows would pass the gate and the trigger would silently
  hard-delete hundreds of forecast rows (the R1 data-loss path). Reconciliation, concretely:
  1. Delete is **gated by a single `refCount` NQ that counts ALL THREE referencing sets** — parts by
     `IN_SUPPLIER_ID`, plus both forecast tables by `VC_SUPPLIER_CODE` (B.1 `Supplier/refCount`). The gate
     takes the supplier's **code** as well as its id, because the forecast tables key on code. Any non-zero
     total → block with a clear message; never call `delete`.
  2. Because the live trigger still exists in the parallel-run DB and would fire (nullify parts **and**
     cascade-delete both forecast tables) if a `DELETE` *did* reach the table, the RESTRICT gate ensures we
     **only ever DELETE rows with zero references in all three tables** — so the trigger has nothing to act
     on and is effectively inert. P12 forbids editing the live trigger during parallel run (the legacy app
     may still rely on it), so **the gate is the only lever and it must mirror the full trigger body.** The
     trigger is formally retired in the Postgres phase when D3 becomes real FK `ON DELETE NO ACTION`.
  3. This is a behavioral **divergence from legacy**, mandated by D3 — document it in the Supplier
     parity checks: legacy delete-with-references nullifies parts and cascade-deletes forecast rows;
     rebuild delete-with-references is **blocked**, and a *blocked* delete must leave both forecast tables
     untouched (B.6 must assert this).
  > **Open question for the reviewer / domain expert (flagged, not guessed):** the supplier spec §9 parity
  > check still reads "delete supplier that has parts → parts survive with `IN_SUPPLIER_ID=NULL`." That
  > parity check is *pre-D3* and now contradicts D3. I am following **D3 (block)** as the authoritative,
  > newer decision and flagging the stale §9 line — but if the intent for Supplier specifically is to keep
  > the nullify-unlink (because a supplier genuinely going away while parts persist is a real workflow),
  > that needs an explicit confirmation. Default taken: **RESTRICT, per D3.**

### A.6 FK + enum combo sourcing

- **FK combos** (Supplier only, among these four masters): `lookups/logistics` (genuine surrogate FK, value
  = `IN_LOGISTICS_ID`) and `lookups/partType` (code-valued). Site-scoped via `:siteId` (`IG-SITE` seam).
  Blank/empty selection → save `NULL` (the legacy "empty string bug" workaround in `HoldDetails` — an empty
  logistics saves `IN_LOGISTICS_ID = NULL`).
- **R6 — Inventory-Add-Point is the static S/A enum, NOT an NQ-driven combo.** Per D4 it is required and
  must be `S`/`A`, so it is a static dropdown (below); the `lookups/addPoint` NQ is **dropped as dead** —
  it was never wired to anything once D4 picked the static enum.
  > **Logistics name-resolution parity nuance (R6):** the legacy `INSERT_SupplierInfo` resolved logistics
  > *by name* (`WHERE VC_LOGISTICS_NAME=@SupLogistics`), so it depended on logistics names being unique. The
  > rebuild resolves by `IN_LOGISTICS_ID` (D2), which removes that dependency — but the `lookups/logistics`
  > combo still **displays** the name as its label, so a parity check should confirm the label set matches
  > the legacy name list (and flag any duplicate logistics names that the legacy name-resolution would have
  > silently mis-resolved).
- **Enum combos** are static (no NQ): options declared in the view.
  - `VC_OUTPUT_FILE`: `{value:'T',label:'TEXT'}, {'E','EXCEL'}, {'B','BOTH'}`
  - `VC_INVENTORY_ADD_POINT`: `{value:'S',label:'SHIPPED'}, {value:'A',label:'ARRIVED'}` — per **D4**,
    this is **required + must be S or A** on Supplier (no blank; stock-add depends on it).

### A.7 8.1 ↔ 8.3 notes

- Build on **8.1.52** component ids (`ia.input.*`, `ia.container.flex`, `ia.display.table`) — same set
  the PartsStockMaster spike used and the codegen emits. `# IG81-COMPAT:` markers carry over.
- `# IG83-TODO:` flags: (a) revisit Perspective component prop schemas + event shapes on 8.3;
  (b) replace the `yyyymmddHHMMSSff` audit-string recipe with a real `datetime` default;
  (c) flip every `-- IG-SITE:` predicate on and add the composite `(site_id, code)` unique index.
- Guard any 8.3-only path with `system.util.getVersion()` (none required for this module today — masters
  are plain NQ CRUD that runs identically on both).
- Named Queries themselves are version-portable; nothing in this module needs an 8.3-only API.

---

## §B. Supplier build spec (the worked example)

Master: `INV_SUPPLIER_MST` (PK `IN_SUPPLIER_ID`, business key `VC_SUPPLIER_CODE` varchar(5), unique via
`IX_INV_SUPPLIER_MST`). FK combos: Logistics, Part Type (Create-Order-Sheet), Add Point
(Inventory-Add-Point). This is the deepest master — Size/Logistics/ManifestCost are strict subsets (§C).

### B.1 Named Queries (exact, parameterized)

All in `Named Queries/Supplier/` unless noted. SQL targets SQL Server (`Inventory_Spike` connection).
Param syntax is Ignition NQ `:paramName`. The `-- IG-SITE:` lines are the D1 seam (commented now,
uncommented at the Postgres phase with the composite index).

**`Supplier/list`** (Query) — params: `searchTerm` (String, default `''`), `siteId` (Int4)
```sql
SELECT  s.IN_SUPPLIER_ID        AS "RecordID",
        s.VC_SUPPLIER_CODE      AS "Supplier Code",
        s.VC_SUPPLIER_NAME      AS "Supplier Name",
        s.VC_CITY               AS "City",
        s.VC_STATE              AS "State",
        l.VC_LOGISTICS_NAME     AS "Logistics",
        CASE s.VC_OUTPUT_FILE WHEN 'T' THEN 'TEXT' WHEN 'E' THEN 'EXCEL' WHEN 'B' THEN 'BOTH' END AS "Output File Type",
        CASE s.VC_INVENTORY_ADD_POINT WHEN 'S' THEN 'SHIPPED' WHEN 'A' THEN 'ARRIVED' END AS "Inventory Add Point"
FROM    INV_SUPPLIER_MST s
        LEFT OUTER JOIN INV_LOGISTICS_MST l ON s.IN_LOGISTICS_ID = l.IN_LOGISTICS_ID
WHERE  (:searchTerm = '' OR s.VC_SUPPLIER_CODE LIKE '%' + :searchTerm + '%'
                        OR s.VC_SUPPLIER_NAME LIKE '%' + :searchTerm + '%')
-- IG-SITE:  AND s.site_id = :siteId
ORDER BY s.VC_SUPPLIER_CODE
```
> Parity: with `searchTerm=''` this matches `SELECT_SupplierInfo ''` ordering/contents (enum labels
> identical). Server-side `LIKE` search **improves on** the legacy exact-match client filter (documented
> divergence). Grid shows a useful subset; the full column set lives on Detail.

**`Supplier/get`** (Query) — params: `recordId` (Int4), `siteId` (Int4)
```sql
SELECT  s.IN_SUPPLIER_ID,  s.VC_SUPPLIER_CODE, s.VC_SUPPLIER_NAME,
        s.VC_ADDRESS, s.VC_CITY, s.VC_STATE, s.VC_ZIP, s.VC_COUNTRY,
        s.VC_TEL, s.VC_FAX, s.VC_PERSON, s.VC_EMAIL_ADDRESS,
        s.VC_BREAKDOWN_ORDER_DIRECTORY,
        s.IN_LOGISTICS_ID,                      -- combo value (D2: id, not name)
        l.VC_LOGISTICS_NAME,                    -- combo display
        s.VC_OUTPUT_FILE, s.BIT_ORDER_FILE_TIMESTAMP, s.BIT_SITE_NUMBER_IN_ORDER,
        s.VC_CREATE_ORDER_SHEET, s.VC_INVENTORY_ADD_POINT
FROM    INV_SUPPLIER_MST s
        LEFT OUTER JOIN INV_LOGISTICS_MST l ON s.IN_LOGISTICS_ID = l.IN_LOGISTICS_ID
WHERE   s.IN_SUPPLIER_ID = :recordId
-- IG-SITE:  AND s.site_id = :siteId
```
> Note vs legacy: `SELECT_SupplierInfo` returned the logistics **name** only. `get` returns
> `IN_LOGISTICS_ID` too, so the FK combo binds by id (D2). `VC_CREATE_ORDER_SHEET` (part type) and
> `VC_INVENTORY_ADD_POINT` are returned as their stored codes — the combos map code→label client-side.

**`Supplier/insert`** (Update Query, returns identity) — params: all data fields + `siteId`
```sql
INSERT INTO INV_SUPPLIER_MST
    (VC_SUPPLIER_CODE, VC_SUPPLIER_NAME, VC_ADDRESS, VC_CITY, VC_STATE, VC_ZIP, VC_COUNTRY,
     VC_TEL, VC_FAX, VC_PERSON, VC_EMAIL_ADDRESS, VC_BREAKDOWN_ORDER_DIRECTORY,
     IN_LOGISTICS_ID, VC_OUTPUT_FILE, BIT_ORDER_FILE_TIMESTAMP, BIT_SITE_NUMBER_IN_ORDER,
     VC_CREATE_ORDER_SHEET, VC_INVENTORY_ADD_POINT, VC_ADD
     /* IG-SITE: , site_id */)
VALUES
    (:code, :name, :address, :city, :state, :zip, :country,
     :tel, :fax, :person, :email, :directory,
     :logisticsId, :outputFile, :orderFileTimestamp, :siteNumberInOrder,
     :createOrderSheet, :invAddPoint,
     CONVERT(char(8),GETDATE(),112)
       + SUBSTRING(CONVERT(varchar,GETDATE(),114),1,2) + SUBSTRING(CONVERT(varchar,GETDATE(),114),4,2)
       + SUBSTRING(CONVERT(varchar,GETDATE(),114),7,2) + SUBSTRING(CONVERT(varchar,GETDATE(),114),10,2)
     /* IG-SITE: , :siteId */);
SELECT CAST(SCOPE_IDENTITY() AS int) AS newId;
```
> Improvements over `INSERT_SupplierInfo`: passes `:logisticsId` directly (D2 — no name→id lookup);
> writes full `VC_PERSON`(50)/`VC_BREAKDOWN_ORDER_DIRECTORY`(512) widths (no proc truncation); returns the
> new id. Audit `VC_ADD` is byte-identical to legacy. Explicit column list (not positional — kills the
> ManifestCost P10 fragility class).

**`Supplier/update`** (Update Query) — params: data fields + `recordId` + `siteId`
```sql
UPDATE INV_SUPPLIER_MST SET
    VC_SUPPLIER_CODE=:code, VC_SUPPLIER_NAME=:name, VC_ADDRESS=:address, VC_CITY=:city,
    VC_STATE=:state, VC_ZIP=:zip, VC_COUNTRY=:country, VC_TEL=:tel, VC_FAX=:fax,
    VC_PERSON=:person, VC_EMAIL_ADDRESS=:email, VC_BREAKDOWN_ORDER_DIRECTORY=:directory,
    IN_LOGISTICS_ID=:logisticsId, VC_OUTPUT_FILE=:outputFile,
    BIT_ORDER_FILE_TIMESTAMP=:orderFileTimestamp, BIT_SITE_NUMBER_IN_ORDER=:siteNumberInOrder,
    VC_CREATE_ORDER_SHEET=:createOrderSheet, VC_INVENTORY_ADD_POINT=:invAddPoint,
    VC_LASTUPDATE = CONVERT(char(8),GETDATE(),112)
       + SUBSTRING(CONVERT(varchar,GETDATE(),114),1,2) + SUBSTRING(CONVERT(varchar,GETDATE(),114),4,2)
       + SUBSTRING(CONVERT(varchar,GETDATE(),114),7,2) + SUBSTRING(CONVERT(varchar,GETDATE(),114),10,2)
WHERE IN_SUPPLIER_ID = :recordId
-- IG-SITE:  AND site_id = :siteId
```
> Keys on the surrogate id (D2); rewrites the code (rename-safe, D2). `VC_ADD` untouched.

**`Supplier/checkCodeUnique`** (Query) — params: `code` (String), `excludeId` (Int4, default 0), `siteId`
```sql
SELECT COUNT(*) AS n
FROM   INV_SUPPLIER_MST
WHERE  VC_SUPPLIER_CODE = :code
  AND  IN_SUPPLIER_ID  <> :excludeId
-- IG-SITE:  AND site_id = :siteId
```
> `excludeId=0` for insert; `excludeId=recordId` for update (lets a row keep its own code). Replaces the
> legacy client two-step. The `IX_INV_SUPPLIER_MST` unique index is the race backstop.

**`Supplier/refCount`** (Query) — params: `recordId` (Int4), `code` (String) — the D3 RESTRICT gate
```sql
-- Counts EVERY table the live DELETE_SupplierCode trigger would touch, so the gate blocks any
-- delete that would fire the trigger's forecast cascade (R1). Parts key on the surrogate id;
-- BOTH forecast tables key on the supplier CODE (that is how the trigger matches them).
SELECT
    (SELECT COUNT(*) FROM INV_PARTS_STOCK_MST  WHERE IN_SUPPLIER_ID   = :recordId)
  + (SELECT COUNT(*) FROM INV_BREAKDOWN_FC_INF WHERE VC_SUPPLIER_CODE = :code)
  + (SELECT COUNT(*) FROM INV_FORECAST_INF     WHERE VC_SUPPLIER_CODE = :code)
    AS n
```
> **R1 fix (data-loss).** The original gate counted only `INV_PARTS_STOCK_MST.IN_SUPPLIER_ID` and was
> **blind** to the trigger's two forecast hard-deletes. A supplier with zero parts but live forecast rows
> would pass the old gate; the rebuild would call `delete`; the live trigger would silently hard-delete
> hundreds of `INV_BREAKDOWN_FC_INF` (~959) + `INV_FORECAST_INF` (~1066) rows. This NQ now mirrors the full
> trigger body so any forecast cascade is blocked at the gate. The supplier `code` is sourced from the
> already-loaded form (it is the deleted row's `VC_SUPPLIER_CODE`); never trust a client-supplied code for
> a different row.

**`Supplier/delete`** (Update Query) — param: `recordId` (Int4)
```sql
DELETE FROM INV_SUPPLIER_MST WHERE IN_SUPPLIER_ID = :recordId
-- IG-SITE:  AND site_id = :siteId
```
> Only reached after `refCount = 0` (A.5). The live `DELETE_SupplierCode` trigger then has no parts to
> nullify (inert by construction).

**FK-combo lookups** (in `Named Queries/lookups/`), each Query with param `siteId`:
```sql
-- lookups/logistics
SELECT IN_LOGISTICS_ID AS id, VC_LOGISTICS_NAME AS label
FROM INV_LOGISTICS_MST  /* IG-SITE: WHERE site_id = :siteId */  ORDER BY VC_LOGISTICS_NAME;

-- lookups/partType   (Create-Order-Sheet combo; VC_CREATE_ORDER_SHEET stores the part-type code)
SELECT VC_PART_TYPE AS id, VC_PART_TYPE AS label
FROM INV_PART_TYPE_MST  /* IG-SITE: WHERE site_id = :siteId */  ORDER BY VC_PART_TYPE;

-- R6: NO lookups/addPoint NQ. Inventory-Add-Point is the static S/A enum (D4 / A.6 / B.2); the
--     previously-listed INV_ADD_POINT_INF query is dropped as dead.
```
> **Note (flagged):** `VC_CREATE_ORDER_SHEET` (Create-Order-Sheet) is stored as a **code**, not an id — it
> is not a true surrogate FK (the legacy schema has no FK here; the combo validates-at-entry only), so its
> dropdown value is the **code string**, not an id. Only `IN_LOGISTICS_ID` is a genuine surrogate FK
> (value = id, per D2). `VC_INVENTORY_ADD_POINT` is the **static S/A enum** (D4), not a lookup combo at all
> (R6 — `lookups/addPoint` dropped). If the rebuild later promotes Part Type to a real surrogate FK, that is
> a separate schema change.

### B.2 List + Detail component layout

**List** (`Supplier/List`)
- Filter row: `text-field` (Search code/name) + `Search` button (re-runs list NQ) + `New` button.
- `ia.display.table` bound to `Supplier/list`. Columns: Supplier Code, Supplier Name, City, State,
  Logistics, Output File Type, Inventory Add Point. Hidden: `RecordID` (used for navigation).
- Row select → navigate `Supplier/Detail` with `recordId = RecordID`.

**Detail** (`Supplier/Detail`) — rows top-to-bottom, label → control → form key:
| Label | Control | form key / column |
|---|---|---|
| Supplier Code | text-field (maxLength 5) | `VC_SUPPLIER_CODE` |
| Supplier Name | text-field (25) | `VC_SUPPLIER_NAME` |
| Address | text-field (50) | `VC_ADDRESS` |
| City | text-field (50) | `VC_CITY` |
| State | text-field (50) | `VC_STATE` |
| Zip | text-field (10) | `VC_ZIP` |
| Country | text-field (50) | `VC_COUNTRY`  *(legacy form omitted it; surface it — spec §8.4 open, harmless to show)* |
| Telephone | text-field (10) | `VC_TEL` |
| Fax | text-field (10) | `VC_FAX` |
| Person | text-field (50) | `VC_PERSON`  *(full 50 — no proc truncation)* |
| Email | text-field (255) | `VC_EMAIL_ADDRESS` |
| Breakdown Order Directory | text-field (512) | `VC_BREAKDOWN_ORDER_DIRECTORY`  *(see B.3 note)* |
| Logistics | **dropdown** (NQ `lookups/logistics`) | value `IN_LOGISTICS_ID`, label name |
| Output File Type | **dropdown** (static T/E/B) | `VC_OUTPUT_FILE` |
| Order File Timestamp | checkbox | `BIT_ORDER_FILE_TIMESTAMP` |
| Site Number in Order | checkbox | `BIT_SITE_NUMBER_IN_ORDER` |
| Create Order Sheet | dropdown (NQ `lookups/partType`, code-valued) | `VC_CREATE_ORDER_SHEET` |
| Inventory Add Point | **dropdown** (static S/A, required) | `VC_INVENTORY_ADD_POINT` |
| *(action row)* | Save · Delete · Cancel | |

Audit columns `VC_ADD` / `VC_LASTUPDATE` are not rendered.

### B.3 Binding / transform flow (Supplier specifics)

- **Load (edit mode):** `runNamedQuery("Supplier/get", {recordId, siteId})` → mutate `view.custom.form`
  keys in place. Set the Logistics combo value from `IN_LOGISTICS_ID` (null → blank). Output File / Add
  Point combos take the stored code directly (static option lists map to labels).
- **FK options:** Logistics combo options ← `runNamedQuery("lookups/logistics", {siteId})` on startup.
- **Save:** build the param map from `view.custom.form`; coerce empties:
  - empty Logistics selection → `logisticsId = None` (NULL) — preserves the legacy "blank logistics saves
    NULL" behavior.
  - bits → 0/1 from checkbox `selected`.
  - run validation (B.4) → on pass, insert or update; on the dup-key SQLException, surface the friendly
    "Supplier code already exists" message.
- **`VC_BREAKDOWN_ORDER_DIRECTORY` (flagged):** it is a legacy local-Windows path chosen via a directory
  picker — meaningless in a web/multi-site client. The supplier spec §8.2 is still **open** (per-site
  output root vs SFTP vs going away). For this first build: **render it as a plain editable text field at
  full 512 width** (parity: the column keeps working for the parallel-run legacy app), and **flag** that
  the picker is dropped and the eventual per-site output-root model (D1: directories move into `sites`) is
  unresolved. Do not build a directory picker.

### B.4 Validation enforcement (Supplier)

In the Save `onActionPerformed` script, before any write:
1. **5-char code:** `len(form['VC_SUPPLIER_CODE'].strip()) == 5` else error "Supplier code must be
   exactly 5 characters." (Legacy rejected `<5`; we reject `!=5`.)
2. **presence:** code and `VC_SUPPLIER_NAME` non-blank.
3. **Inventory Add Point required + S/A** (D4): non-blank and in `{'S','A'}`.
4. **per-site uniqueness:** `runNamedQuery("Supplier/checkCodeUnique", {code, excludeId, siteId})`;
   `n > 0` → error "Supplier code already exists." `excludeId = recordId` on update, `0` on insert.
5. **race backstop:** wrap the insert/update in try/except; on SQLException whose message indicates a
   unique-index violation (`IX_INV_SUPPLIER_MST`), show the same friendly uniqueness message.

This set **replaces** the legacy client two-step dup-check and the redundant app/index duplication; the
index is the authoritative backstop (spec §4).

### B.5 Delete — D3 reconciliation (Supplier)

1. On Delete: `n = runNamedQuery("Supplier/refCount", {recordId, code: form['VC_SUPPLIER_CODE']})`.
   refCount counts **all three** tables the live trigger touches: parts (by id) + `INV_BREAKDOWN_FC_INF`
   + `INV_FORECAST_INF` (both by code).
2. `n > 0` → block: "Cannot delete supplier — still referenced by N part(s) / forecast row(s). Reassign
   or archive those rows first." (Archival is the future D3 path; out of scope now.)
3. `n == 0` → `runNamedQuery("Supplier/delete", {recordId})` → back to List.

**R1 (data-loss) — why the gate must see the forecast tables:** the live `DELETE_SupplierCode` trigger does
not only nullify parts; it **hard-DELETEs** matching `INV_BREAKDOWN_FC_INF` + `INV_FORECAST_INF` rows by
supplier code. A parts-only refCount would let a parts-free-but-forecast-bearing supplier through and the
trigger would silently destroy hundreds of forecast rows. The gate now mirrors the full trigger body, so the
**only** `DELETE` we ever issue is against a supplier with zero references in all three tables — making the
trigger inert by construction. P12 forbids touching the live trigger during parallel run, so the gate is the
sole lever and must remain in lockstep with the trigger body (re-verify if the trigger changes).

**Divergence from legacy (document in parity checks):** legacy `DELETE_SupplierCode` *unlinked* parts
(`IN_SUPPLIER_ID=NULL`) **and cascade-deleted both forecast tables**, then deleted the supplier; the rebuild
**blocks** the delete while parts or forecast rows reference it (D3). **This contradicts the stale supplier §9
parity line** ("parts survive with `IN_SUPPLIER_ID=NULL`") — see the A.5 flag; default taken is D3 (block).

### B.6 Parity / divergence checklist (for `ignition-qa`)

| Check | Expected (rebuild) | vs legacy |
|---|---|---|
| List all (`searchTerm=''`) | rows + order match `SELECT_SupplierInfo ''`; enum labels identical | parity |
| Search "ABC" | server-side `LIKE` on code+name | **divergence** (legacy = client exact match) — documented |
| Insert dup 5-char code | rejected (pre-check + index) with friendly msg | parity in outcome; mechanism differs |
| Insert `<5` or `!=5` code | rejected | parity (legacy rejected `<5`) |
| Insert blank Add Point | rejected (D4) | **divergence** (legacy allowed) — documented |
| Save blank Logistics | `IN_LOGISTICS_ID = NULL` | parity |
| Rename code to a free code | succeeds (D2) | parity (legacy allowed) |
| Rename code onto existing | rejected (pre-check/index) | parity in outcome |
| New row returns id | `SCOPE_IDENTITY()` echoed; form flips to edit | **improvement** (legacy never echoed) |
| Delete supplier w/ parts | **blocked** (D3) | **divergence** (legacy unlinked) — documented; see A.5 flag |
| Delete supplier w/ **forecast rows but zero parts** | **blocked** (D3) — refCount sees `INV_BREAKDOWN_FC_INF`+`INV_FORECAST_INF` by code | **R1 regression guard** — old parts-only gate let this through and the trigger hard-deleted forecast rows |
| Blocked delete leaves forecast tables intact | row counts of `INV_BREAKDOWN_FC_INF` + `INV_FORECAST_INF` for that supplier code are **unchanged** after a blocked delete (no `DELETE` ever issued) | **R1 assertion** — proves the cascade never fired |
| Delete supplier w/o **any** references (parts + both forecast tables = 0) | row deleted | parity |
| `VC_ADD`/`VC_LASTUPDATE` | byte-identical 16-char strings | parity (P2) |

---

## §C. Later masters — deltas vs the Supplier template

Each later master instantiates §A; below is only what changes. All three are **strict simplifications**
of Supplier except ManifestCost's overlap rule.

### Size (`INV_SIZE_MST`) — the leanest
- **NQ folder** `Size/`; key `IN_SIZE_ID`, code `VC_SIZE_CODE` varchar(6), index `IX_INV_SIZE_MST`.
- **No FK combos, no enums** — leaf master. Drop all `lookups/*` for this view.
- Columns: `VC_SIZE_CODE`, `VC_SIZE_NAME`, `IN_USAGE` (int), `IN_DAYS` (int) — two numeric-entry-fields.
- **Audit column is `VC_LAST_UPDATE` (underscore)** — map the right column in `Size/update` (Supplier
  uses `VC_LASTUPDATE`, no underscore). Easy to get wrong.
- **Validation:** add presence + `len(code) ≤ 6` (legacy had **none** — documented improvement).
  `checkCodeUnique` checks `INV_SIZE_MST.VC_SIZE_CODE` — **NOT** `SELECT_AssyRatioInfo` (D8 Bug 1; do not
  reproduce). `excludeId` supports rename (D2).
- **Cross-module coupling (flag):** ForecastBreakdown writes `IN_USAGE` via `UPDATE_SizeUsage` keyed *by
  code*; per D2 that caller must be reworked to key by `IN_SIZE_ID` when that module is built. Not this
  module's job, but the Size NQ set should expose an id-keyed usage update for it to adopt.
- **Delete:** `refCount` on `INV_PARTS_STOCK_MST.IN_SIZE_ID` (+ `_HIST` per D3); block if referenced.
- **`0` vs `NULL` for usage/days** (spec §8.3) still open — default: write what the field holds (0 if
  blank, matching legacy); flag for confirmation.

### Logistics (`INV_LOGISTICS_MST`) — leaf, name-keyed
- **NQ folder** `Logistics/`; key `IN_LOGISTICS_ID`, business key `VC_LOGISTICS_NAME` varchar(25) (no
  5-char code), index `IX_INV_LOGISTICS_MST`.
- **No FK combos, no enums.** Columns: name + address block + tel/fax/person/email + directory.
- Validation: presence + per-site uniqueness on **name** (`checkCodeUnique` checks `VC_LOGISTICS_NAME`);
  **no fixed-length rule** (no code). Legacy had a working app dup-check **and** index — we keep the
  pre-check + index pattern. `excludeId` supports rename (D2). Widen the legacy proc-truncated fields
  (Person 25→50, Directory 215→512) since direct NQ writes full width.
- **`VC_BREAKDOWN_ORDER_DIRECTORY`** — same flagged open question as Supplier B.3 (plain text field at
  full width for now; picker dropped).
- **Delete (D3):** `refCount` must count **both** `INV_SUPPLIER_MST.IN_LOGISTICS_ID` **and**
  `INV_PARTS_STOCK_MST.IN_LOGISTICS_ID` (+ `_HIST`). Legacy nulled only supplier FKs and **dangled** part
  FKs; D3 blocks if *either* references it. This is the cleanest D3 win (kills the legacy dangle).

### ManifestCost (`INV_MANIFEST_COST_MST`) — financially load-bearing, the outlier
- **NQ folder** `ManifestCost/`; key `IN_MANIFEST_COST_ID`, assembly business key `VC_ASSY_PART_NUMBER_CODE`
  varchar(12), manifest id `VC_ASSY_MANIFEST_NUMBER` varchar(2).
- **⚠️ R2 — live-schema correction (verified 2026-06-16):** the manifest-cost spec §6 and this design
  previously asserted "**no PK / no unique index / no FK / no trigger** (P11)". That is **stale** against the
  LIVE spike DB. The live DB **HAS a UNIQUE index `IX_INV_MANIFEST_COST_MST` on `VC_ASSY_MANIFEST_NUMBER`**
  (the 2-char manifest id) — *not* on the assembly code, *not* composite. So a live INSERT with a fresh assy
  code and a fresh, non-overlapping window but a **duplicate manifest number** is **rejected by the index**.
  There is still no PK and no FK; the rebuild adds a surrogate PK. *Spec correction needed: §6 "no unique
  index" must be replaced with the live `IX_INV_MANIFEST_COST_MST` UNIQUE(`VC_ASSY_MANIFEST_NUMBER`).*
- **The constraint set the rebuild ENFORCES (re-derived from live schema + spec §8 + D6):**
  1. **`start_manifest <= end_manifest`** — reject `start > end` (D6, spec §8.7). Validation rule.
  2. **No-overlapping-window per `(site, VC_ASSY_PART_NUMBER_CODE)`** — D6's real rule: two price rows for the
     **same assy code** are allowed only when their date windows don't overlap. Enforced by
     `checkWindowOverlap` (below). This is on the **assembly code**, orthogonal to the manifest number.
  3. **Global-unique `VC_ASSY_MANIFEST_NUMBER`** — *this is the live `IX_INV_MANIFEST_COST_MST` index*, NOT
     something the design invented, and **NOT** the same constraint as (2). It forbids two rows sharing a
     2-char manifest number **regardless of assy code or window**. The rebuild must pre-check it (a
     `checkManifestNumberUnique` NQ) AND catch the index violation as the backstop — otherwise a D6-valid
     insert (same manifest number, different assy code, non-overlapping window) sails past `checkWindowOverlap`
     and dies on the live index with a raw SQLException the try/except wasn't written to expect. **(3) is the
     R2 fix: it was entirely missing from the previous design.**
- Columns: assy code (combo, code-valued, sourced from distinct `INV_FORECAST_DETAIL_INF` codes — keep
  but flag spec §8.5: is that the authoritative assembly domain?); manifest number (static `' ','01'..'99'`
  dropdown); start/end manifest (`yyyymmdd` strings → real date pickers, format on save); `MO_PRICE`
  (money, numeric-entry-field, `$`/4dp).
- **Both `VC_ADD` and `VC_LAST_UPDATE` are set equal on insert** (legacy positional INSERT did this) —
  the `insert` NQ writes the same 16-char string to both. `VC_LAST_UPDATE` is **underscored** (like Size).
- **Validation is the headline delta (D6):**
  - presence on assy code / both dates / price;
  - `start_manifest <= end_manifest` (reject `start > end`, D6);
  - `MO_PRICE` numericality (≥ 0 half of §8.7 still open — flag);
  - **no-overlapping-window per `(site, VC_ASSY_PART_NUMBER_CODE)`** — `checkCodeUnique` becomes
    `checkWindowOverlap`: `SELECT COUNT(*) ... WHERE VC_ASSY_PART_NUMBER_CODE=:code AND
    IN_MANIFEST_COST_ID<>:excludeId AND NOT (:endManifest < VC_START_MANIFEST OR :startManifest >
    VC_END_MANIFEST) [AND site_id=:siteId]`; `n>0` → reject. Two prices for one assy code are allowed only
    when windows don't overlap (D6). **This is on the ASSEMBLY code, and by itself does NOT satisfy the
    live DB** (see next bullet).
  - **`checkManifestNumberUnique` (R2 — was missing): `SELECT COUNT(*) ... WHERE
    VC_ASSY_MANIFEST_NUMBER=:manifestNo AND IN_MANIFEST_COST_ID<>:excludeId`.** This pre-check mirrors the
    **live `IX_INV_MANIFEST_COST_MST` UNIQUE index** so a D6-valid-but-manifest-number-duplicate insert is
    rejected with a friendly message instead of dying on a raw index SQLException. The validation order is:
    presence → `start<=end` → `checkWindowOverlap` → `checkManifestNumberUnique`; the index is the race
    backstop (catch the unique-violation SQLException and surface the friendly manifest-number message).
    > **R2 reconciliation — FLAGGED FOR DAVID (do not guess):** `checkWindowOverlap` (per-assy, D6) and the
    > live global-unique manifest number are **two different constraints**. In the current live data they
    > coincide (45 rows = 45 distinct assy codes = 45 distinct manifest numbers — one window per assy, one
    > manifest number per row), so the conflict is latent. D6 explicitly contemplates **multiple windows per
    > assy code**, but the live index forbids reusing a manifest number for the second window — so the moment
    > a second window is added for an assy code it needs a *different* manifest number, and manifest numbers
    > are globally scarce (`' '`,`'01'..'99'` = 100 values). **Question for David:** is global manifest-number
    > uniqueness *intended* (manifest number is a real global slot/identifier), or is it a **legacy quirk** —
    > the index should arguably be `(VC_ASSY_PART_NUMBER_CODE, VC_ASSY_MANIFEST_NUMBER)` or dropped in favor
    > of the D6 window rule? Until David rules: the rebuild **honors the live index as-is** (pre-check +
    > backstop), because the parallel-run legacy app shares the table and the index is live. The developer
    > builds `checkManifestNumberUnique` against the real index, NOT a phantom "no index" baseline.
- **Delete (D3):** `refCount` against ASN-detail / invoice-line references on the assy code; block if
  referenced. Ends the legacy silent invoice-line-loss (no trigger today).
- **Do NOT reproduce** the wrong-target retry recursion (P12) or the positional INSERT (P10) — direct NQ
  with an explicit column list and a single try/except handles both.
- **Billing window-bug (D6/D11) is NOT fixed here** — it lives in the EDI/Invoice read path
  (`REPORT_EDI810`/`SELECT_INVOICEItems`). This master only *feeds* the price + enforces non-overlap; flag
  the window-aware billing fix for that module.

---

## Open items flagged for the reviewer / domain expert (not guessed)

1. **Supplier delete policy (A.5/B.5):** D3 says block; the stale supplier §9 parity line says unlink.
   Default taken = **D3 block**. Confirm Supplier shouldn't keep nullify-unlink.
2. **`VC_BREAKDOWN_ORDER_DIRECTORY` model (Supplier/Logistics §8.2 — still open):** rendered as plain
   text now; the per-site output-root model (D1: dirs into `sites`) is undecided. Picker dropped.
3. **Part Type (Create-Order-Sheet) as a code-valued combo, not a surrogate FK (B.1 note):** preserved as
   legacy (no FK exists); confirm we are not expected to promote it to a real FK id now. Add Point is the
   static S/A enum per D4 (R6 — `lookups/addPoint` NQ dropped).
4. **Size `0` vs `NULL`** for `IN_USAGE`/`IN_DAYS` (§8.3): default = write field value (0 if blank).
5. **ManifestCost negative/zero price** (§8.7 half still open) and the **assembly-code domain** (§8.5):
   flagged; not blocking the CRUD build.
6. **ManifestCost manifest-number global-uniqueness (R2) — NEEDS DAVID.** The live
   `IX_INV_MANIFEST_COST_MST` is UNIQUE on `VC_ASSY_MANIFEST_NUMBER` (verified 2026-06-16; spec §6 stale).
   It is a *different* constraint from D6's per-assy non-overlap rule and conflicts with D6's "multiple
   windows per assy code" the moment a second window is added. Is global manifest-number uniqueness intended
   or a legacy quirk (→ composite index or drop)? Default until ruled: **honor the live index** (pre-check +
   backstop). See §C ManifestCost reconciliation block.
