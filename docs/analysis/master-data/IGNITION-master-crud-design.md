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
- `INSERT_SupplierInfo` resolves logistics **by name**, narrows `@SupPerson`→25 / `@SupDirectory`→215,
  computes the 16-char `yyyymmddHHMMSSff` timestamp inline, and **takes no `site_id`**.
- FK lookup tables confirmed: `INV_LOGISTICS_MST(IN_LOGISTICS_ID,VC_LOGISTICS_NAME)`,
  `INV_PART_TYPE_MST(IN_PART_TYPE_ID,VC_PART_TYPE,VC_PART_TYPE_DESCRIPTION)`,
  `INV_ADD_POINT_INF(IN_ADD_POINT_ID,VC_ADD_POINT_ABRV,VC_ADD_POINT)`.

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
  - **`site_id` is sourced server-side, never from the client.** It comes from
    `session.custom.siteId` (set at login from the Ignition User Source's site binding — `D1`: "auth binds
    each user to a site"). A view passes `{session.custom.siteId}` into NQ params; the NQ never trusts a
    client-supplied site. This is the InventorySystem analogue of the GALC `siteScopedQuery()` rule.
  - For the spike, `session.custom.siteId` defaults to `1` (TMMMS). Switching it to `2` exercises the
    multi-site code path against one-site data — proving the *shape* without needing the column yet.

  Net: **one site of data, but multi-site is structurally real** — site flows session → view → NQ param,
  and turning it on is a per-NQ uncomment + a composite-index migration, not a redesign.

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

Per the `ignition-named-query-crud-practice` memory, organize NQs in a folder that **mirrors the table**
and name each for the op it performs, so "which NQ touches this table" is a single lookup and a schema
change is a single-point edit. Folder per master under the project's Named Queries:

```
Named Queries/
  <Master>/                         e.g. Supplier/  (mirrors INV_SUPPLIER_MST)
    list        — grid rows; params: searchTerm, siteId, sort/page
    get         — one row by id (resolves FK ids AND codes for the form); param: recordId, siteId
    insert      — INSERT; returns new identity (SCOPE_IDENTITY()); all data params + siteId
    update      — UPDATE ... WHERE id = :recordId; all data params + siteId
    delete      — DELETE ... WHERE id = :recordId (after the RESTRICT pre-check, A.5)
    checkCodeUnique  — uniqueness pre-check (A.4); params: code, excludeId, siteId
    refCount    — D3 RESTRICT pre-check: count referencing rows; param: recordId   (A.5)
  lookups/
    logistics   — IN_LOGISTICS_ID + VC_LOGISTICS_NAME, site-scoped  (shared FK combo source)
    partType    — IN_PART_TYPE_ID + VC_PART_TYPE
    addPoint    — IN_ADD_POINT_ID + VC_ADD_POINT_ABRV + VC_ADD_POINT
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
        n = runNamedQuery("Supplier/refCount", {recordId})    # A.5 D3 RESTRICT
        if n > 0: show "Cannot delete — still referenced by N parts"; return
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
- **D3 (RESTRICT on delete) vs the legacy nullify trigger — the reconciliation:**
  The legacy `DELETE_SupplierCode` trigger **NULLIFIES** `INV_PARTS_STOCK_MST.IN_SUPPLIER_ID` for the
  deleted supplier's parts (unlink, don't block). **D3 overrides this:** the rebuild **blocks** the
  delete while any part still references the supplier. Reconciliation, concretely:
  1. Delete is **gated by a `refCount` NQ** (`SELECT COUNT(*) FROM INV_PARTS_STOCK_MST WHERE
     IN_SUPPLIER_ID = :recordId`). `> 0` → block with a clear message; never call `delete`.
  2. Because the live `DELETE_SupplierCode` trigger still exists in the parallel-run DB and would fire
     (and nullify) if a `DELETE` *did* reach the table, the rebuild's RESTRICT gate ensures we **only
     ever DELETE rows with zero references** — so the trigger's nullify branch has nothing to act on and
     is effectively inert. We do **not** drop or edit the legacy trigger during parallel run (P12
     no-legacy-hotfix; the legacy app may still rely on it). The trigger is formally retired in the
     Postgres phase when D3 becomes a real FK `ON DELETE NO ACTION`.
  3. This is a behavioral **divergence from legacy**, mandated by D3 — document it in the Supplier
     parity checks: legacy delete-with-parts unlinks; rebuild delete-with-parts is **blocked**.
  > **Open question for the reviewer / domain expert (flagged, not guessed):** the supplier spec §9 parity
  > check still reads "delete supplier that has parts → parts survive with `IN_SUPPLIER_ID=NULL`." That
  > parity check is *pre-D3* and now contradicts D3. I am following **D3 (block)** as the authoritative,
  > newer decision and flagging the stale §9 line — but if the intent for Supplier specifically is to keep
  > the nullify-unlink (because a supplier genuinely going away while parts persist is a real workflow),
  > that needs an explicit confirmation. Default taken: **RESTRICT, per D3.**

### A.6 FK + enum combo sourcing

- **FK combos** (Supplier only, among these four masters): `lookups/logistics`, `lookups/partType`,
  `lookups/addPoint` NQs return `(id, displayCode)`; dropdown value=id, label=displayCode; site-scoped
  via `:siteId` (`IG-SITE` seam). Blank/empty selection → save `NULL` (the legacy "empty string bug"
  workaround in `HoldDetails` — an empty logistics saves `IN_LOGISTICS_ID = NULL`).
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

**`Supplier/refCount`** (Query) — param: `recordId` (Int4)  — the D3 RESTRICT gate
```sql
SELECT COUNT(*) AS n FROM INV_PARTS_STOCK_MST WHERE IN_SUPPLIER_ID = :recordId
```

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

-- lookups/addPoint   (static enum S/A is preferred per D4; this NQ is only if the live table drives it)
SELECT VC_ADD_POINT_ABRV AS id, VC_ADD_POINT AS label
FROM INV_ADD_POINT_INF  ORDER BY VC_ADD_POINT_ABRV;
```
> **Note (flagged):** `VC_CREATE_ORDER_SHEET` and `VC_INVENTORY_ADD_POINT` are stored as **codes**, not
> ids — they are not true surrogate FKs (the legacy schema has no FK here; the combos validate-at-entry
> only). So for these two the dropdown value is the **code string**, not an id. Only `IN_LOGISTICS_ID` is
> a genuine surrogate FK (value = id, per D2). I'm preserving the legacy code-valued behavior for the two
> non-FK combos because nothing references them by id; if the rebuild later promotes Part Type / Add Point
> to real surrogate FKs, that is a separate schema change. (Per D4, Inventory-Add-Point is better as the
> static `S/A` enum than an NQ-driven combo — recommend static.)

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

1. On Delete: `n = runNamedQuery("Supplier/refCount", {recordId})`.
2. `n > 0` → block: "Cannot delete supplier — still referenced by N part(s). Reassign or archive those
   parts first." (Archival is the future D3 path; out of scope now.)
3. `n == 0` → `runNamedQuery("Supplier/delete", {recordId})` → back to List.

**Divergence from legacy (document in parity checks):** legacy `DELETE_SupplierCode` would *unlink*
(null the parts' `IN_SUPPLIER_ID`) and delete the supplier; the rebuild **blocks** the delete while parts
reference it (D3). The live trigger remains in the parallel-run DB but is inert because we only ever delete
zero-reference rows. **This contradicts the stale supplier §9 parity line** ("parts survive with
`IN_SUPPLIER_ID=NULL`) — see the A.5 flag; default taken is D3 (block).

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
| Delete supplier w/o parts | row deleted | parity |
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
- **NQ folder** `ManifestCost/`; key `IN_MANIFEST_COST_ID`, business key `VC_ASSY_PART_NUMBER_CODE`
  varchar(12). **No PK / no unique index / no trigger exist in the legacy schema** (P11) — the rebuild
  **adds** them (real PK; the overlap constraint below).
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
  - **no-overlapping-window per `(site, VC_ASSY_PART_NUMBER_CODE)`** — replaces a simple unique-code
    check. `checkCodeUnique` becomes `checkWindowOverlap`: `SELECT COUNT(*) ... WHERE
    VC_ASSY_PART_NUMBER_CODE=:code AND IN_MANIFEST_COST_ID<>:excludeId AND NOT
    (:endManifest < VC_START_MANIFEST OR :startManifest > VC_END_MANIFEST) [AND site_id=:siteId]`;
    `n>0` → reject. Two prices for one assy code are allowed only when windows don't overlap (D6).
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
3. **Part Type / Add Point as code-valued combos, not surrogate FKs (B.1 note):** preserved as legacy
   (no FK exists); confirm we are not expected to promote them to real FK ids now. Recommend static S/A
   enum for Add Point per D4.
4. **Size `0` vs `NULL`** for `IN_USAGE`/`IN_DAYS` (§8.3): default = write field value (0 if blank).
5. **ManifestCost negative/zero price** (§8.7 half still open) and the **assembly-code domain** (§8.5):
   flagged; not blocking the CRUD build.
