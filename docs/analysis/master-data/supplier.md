# Module Analysis: Supplier Master

**Area:** Master data  **Status:** ✅ spec complete  **Analyst:** Claude / 2026-06-01

> First fully-analyzed module — also validates the migration methodology end to end.
> "Supplier" here = the **vendors the site orders parts from** (not the site's own
> identity, which lives in `InventorySystem.INI [SITE]`).

## 1. Legacy surface
- **Form:** `SupplierMaster.pas` (12.7 KB) + `SupplierMaster.dfm`. Author: Aaron Huge, 2002.
- **Entry point:** Master-maintenance menu in `MainMenu.pas` → `SupplierMaster_Form.Execute`.
- **Purpose:** Classic master-detail CRUD screen. A `DBGrid` lists all suppliers; an edit
  panel shows the selected supplier's details. Buttons: **Insert, Update, Search, Clear,
  Delete, Close**. Selecting a grid row (click / key / datasource change) populates the
  detail panel via `HoldDetails(True)` → `SetDetailBoxes`.

## 2. Data touched
| Table | Read | Write | Notes |
|-------|:----:|:-----:|-------|
| `INV_SUPPLIER_MST` | ✓ | ✓ | The supplier/vendor master (this module owns it) |
| `INV_LOGISTICS_MST` | ✓ |  | FK lookup; combo resolves **name → `IN_LOGISTICS_ID`** |
| `INV_PART_TYPE_MST` | ✓ |  | Populates "Create Order Sheet" combo (`VC_PART_TYPE`) |
| `INV_ADD_POINT_INF` | ✓ |  | Populates "Inventory Add Point" combo (`VC_ADD_POINT`) |
| `INV_PARTS_STOCK_MST` |  | ✓* | *Indirect: via DELETE trigger (FK unlink) |

### `INV_SUPPLIER_MST` columns
| Column | Type | Meaning / notes |
|--------|------|-----------------|
| `IN_SUPPLIER_ID` | int IDENTITY PK | Surrogate key (`RecordID` in UI) |
| `VC_SUPPLIER_CODE` | varchar(5) NOT NULL | **Business key — exactly 5 chars; DB-unique via `IX_INV_SUPPLIER_MST`** |
| `VC_SUPPLIER_NAME` | varchar(25) | |
| `VC_ADDRESS`,`VC_CITY`,`VC_STATE`,`VC_ZIP`,`VC_COUNTRY` | varchar | Address (COUNTRY unused by form) |
| `VC_TEL`,`VC_FAX` | varchar(10) | |
| `VC_PERSON` | varchar(50) | Contact (form caps input at 25) |
| `VC_EMAIL_ADDRESS` | varchar(255) | |
| `VC_BREAKDOWN_ORDER_DIRECTORY` | varchar(512) | **Local filesystem path** for order files ⚠️ desktop-bound |
| `IN_LOGISTICS_ID` | int FK | → `INV_LOGISTICS_MST` |
| `VC_OUTPUT_FILE` | varchar(1) | `T`=TEXT, `E`=EXCEL, `B`=BOTH |
| `BIT_ORDER_FILE_TIMESTAMP` | bit | Append timestamp to order file? |
| `BIT_SITE_NUMBER_IN_ORDER` | bit | **Include site number in order file** — latent multi-site hook |
| `VC_CREATE_ORDER_SHEET` | varchar(5) | Part type that triggers order-sheet creation |
| `VC_INVENTORY_ADD_POINT` | varchar(1) | `S`=SHIPPED, `A`=ARRIVED (when stock is counted as on-hand) |
| `VC_ADD`,`VC_LASTUPDATE` | varchar(16) | **Timestamps stored as `yyyymmddHHMMSSff` strings** (16 chars: `CONVERT(…,112)` date + 4×`SUBSTRING(…,114)` = HH+MM+SS+`ff`), not datetime ⚠️ |

**Triggers on these tables:**
- `DELETE_SupplierCode` (on `INV_SUPPLIER_MST` FOR DELETE): sets
  `INV_PARTS_STOCK_MST.IN_SUPPLIER_ID = NULL` for any parts pointing at the deleted
  supplier. **Invariant: deleting a supplier unlinks (does NOT delete) its parts.**

## 3. Stored procedures used
| Proc | Op | Business rule (from body) |
|------|----|---------------------------|
| `SELECT_SupplierInfo @SupCode='' , @Logistics=0` | SELECT | If `@SupCode=''` returns **all** suppliers (LEFT JOIN logistics) ordered by code; else the one matching `@SupCode`. Maps stored codes to display labels (`T/E/B`→TEXT/…, `S/A`→SHIPPED/ARRIVED). Returns `RecordID` + `LogisticsDirectory`. |
| `INSERT_SupplierInfo` (17 params) | INSERT | Resolves logistics **name→id**; computes `VC_ADD` timestamp string; inserts. **No uniqueness check inside the proc** — the dup-check is in the app (see below). |
| `UPDATE_SupplierInfo` (+`@SupplierID`) | UPDATE | Resolves logistics name→id; sets `VC_LASTUPDATE`; updates **by `IN_SUPPLIER_ID`**. (Note: it *does* rewrite `VC_SUPPLIER_CODE`, so the business key is editable.) |
| `DELETE_SupplierInfo @SupplierID` | DELETE | Deletes by `IN_SUPPLIER_ID`; relies on the trigger to unlink parts. |
| `SELECT_PartsSupplier @VC_SUPPLIER_CODE` | SELECT | Parts joined to supplier; all parts if code blank. (Used elsewhere, e.g. parts/order screens — listed here for completeness.) |

### Call mechanism (legacy)
`DataModule.pas` methods `GetSupplierInfo / InsertSupplierInfo / UpdateSupplierInfo /
DeleteSupplierInfo` use a single ADO `Inv_StoredProc`, setting `ProcedureName :=
'dbo.PROC;1'` and adding `@params`. **Insert is two-step:** it first calls
`SELECT_SupplierInfo @SupCode` and only inserts `If RecordCount = 0` — i.e. the
**duplicate-code guard lives in the client, not the DB.** (Reusable pattern across masters.)

## 4. Business rules & edge cases
- **Supplier code = exactly 5 chars** (form `Validate`: `length < 5` rejected; field is
  varchar(5) so >5 truncates).
- **Code must be unique** — enforced **both** by the app-side dup check (the two-step Insert) **and**
  by a real DB UNIQUE index `IX_INV_SUPPLIER_MST` on `VC_SUPPLIER_CODE` (verified in the schema). The
  app check is redundant with the index; the index is the true backstop against a rename collision.
  → In the rebuild, a model `uniqueness` validation backed by that **existing** index replaces the app check.
- **Logistics is referenced by name** in the UI but stored as id; an empty/blank combo
  saves `IN_LOGISTICS_ID = NULL` (explicit "empty string bug" workaround in `HoldDetails`).
- **Timestamps are `yyyymmddHHMMSSff` strings** (16 chars; `VC_ADD` on insert, `VC_LASTUPDATE` on
  update). Preserve format if the legacy app keeps reading the same rows during parallel
  run; normalize to real timestamps only at the Postgres phase.
- **Delete is soft on parts** (unlink via trigger), hard on the supplier row.
- Coded single-char enums: OutputFile `T/E/B`, AddPoint `S/A`.
- `VC_BREAKDOWN_ORDER_DIRECTORY` is a Windows path chosen via a directory picker —
  meaningless in a web/multi-site context (see §8).

## 5. UI / UX notes
- Grid + detail-panel pattern; selection syncs panel. Search is **client-side** over the
  already-loaded grid (`SearchGrid` loops the dataset) — fine to replace with a server query.
- Combos: Logistics, Create-Order-Sheet (part type), Inventory-Add-Point.
- Radio: Output file type. Checkboxes: Order-file timestamp, Site-number-in-order.
- Modernize: standard index/list + form (new/edit), server-side search/sort/pagination,
  inline validation. Drop the directory-picker (replace per §8).

## 6. Target design  *(Rails primary)*
- **Model:** `Supplier` → table `INV_SUPPLIER_MST` (`self.table_name`, custom PK
  `IN_SUPPLIER_ID`). `belongs_to :logistics, optional: true` (FK `IN_LOGISTICS_ID`).
  `has_many :parts_stocks` with `dependent: :nullify` (mirrors the trigger).
  - Validations: `supplier_code` presence, `length: {is: 5}`, `uniqueness` (+ DB unique index).
  - Enums: `output_file {T,E,B}`, `inventory_add_point {S,A}`.
  - Callback/columns: set `vc_add`/`vc_lastupdate` — keep string format during parallel run.
- **Controller/routes:** RESTful `resources :suppliers`.
- **Views:** index (searchable/paginated list) + new/edit form; combos → selects sourced
  from Logistics / PartType / AddPoint.
- **Services:** none needed; pure CRUD. **Stage-1 option:** wrap the existing procs (call
  `SELECT/INSERT/UPDATE/DELETE_SupplierInfo` via `tiny_tds`) for guaranteed parity, then
  switch to ActiveRecord in stage 3.
- **Reports:** none specific.

## 7. Migration plan for this module
- [ ] Stage 1 — read-only: `Supplier.all` (or wrap `SELECT_SupplierInfo`) renders the list.
- [ ] Stage 2 — writes via wrapped procs (keeps app-side dup check + trigger behavior).
- [ ] Stage 3 — reimplement: ActiveRecord validations replace the dup check; a
      `dependent: :nullify` association replaces `DELETE_SupplierCode`; real timestamps;
      carry the **existing** `IX_INV_SUPPLIER_MST` unique index across. Postgres-ready.

## 8. Open questions for the user (domain expert)
1. **Multi-site & suppliers:** when the app goes multi-site, is the supplier/vendor list
   **shared across all sites** or **per-site**? This decides whether `INV_SUPPLIER_MST`
   gets a `site_id` scope. (Today it has none; `VC_SUPPLIER_CODE` is globally unique.)
2. **`VC_BREAKDOWN_ORDER_DIRECTORY`** is a local Windows path for writing order files.
   In a web/multi-site world, what should replace it — a per-site configured output
   target (network share / SFTP / object storage), or is file output going away in favor
   of in-app delivery? `BIT_SITE_NUMBER_IN_ORDER` suggests order files already encode a
   site number — worth understanding that convention.
3. Is editing `VC_SUPPLIER_CODE` after creation actually used/desired? (UPDATE allows it.)
4. Is `VC_COUNTRY` intentionally unused by the form?

## 9. Test cases / parity checks
- List all → row count and ordering match `SELECT_SupplierInfo ''`.
- Insert with existing 5-char code → rejected (no row added), matching dup-check.
- Insert with <5-char code → rejected by validation.
- Update logistics to blank → `IN_LOGISTICS_ID` becomes NULL.
- Delete supplier that has parts → supplier gone, its `INV_PARTS_STOCK_MST` rows remain
  with `IN_SUPPLIER_ID = NULL` (verify trigger parity).
