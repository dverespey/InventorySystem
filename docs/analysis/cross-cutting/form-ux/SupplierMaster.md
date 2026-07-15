# Form-UX semantics: `TSupplierMaster_Form` — `SupplierMaster.pas` / `SupplierMaster.dfm`

CRUD editor over `INV_SUPPLIER_MST` (code/name/address/contact/logistics/output-file/add-point
config) with grid `SupplierMaster_DBGrid` bound via `Supplier_DataSource`.

## Dialogs & confirmations
- **Insert failure** — `SupplierMaster.pas:158-159`: `MessageDlg('Unable to INSERT ' +
  Data_Module.SupplierName + '(' + Data_Module.SupplierCode + ')', mtInformation, [mbOk], 0)` if
  `InsertSupplierInfo` returns `False`. Note: unlike `SizeMaster`/`LogisticsMaster`/
  `RenbanGroupMaster`, this message does **not** say "It already exists in the database" — same
  failure semantics (duplicate key), different wording.
- **Delete confirmation (armed, two-step)** — `:206-208`: `MessageDlg('Are you sure you wish to
  delete' + #13 + Data_Module.SupplierName + ' (' + Data_Module.SupplierCode + ') from the
  database?', mtWarning, [mbYes, mbNo], 0) = mrYes` gates the actual `DeleteSupplierInfo` call.
- **Validation failures (`Validate`, `:169-183`)** — only ONE rule: `ShowMessage('Supplier Code must
  be 5 characters')` if `Length(SupplierCode_Edit.Text) < 5`, then `SetFocus` back to the code field
  and abort (both `Insert_ButtonClick` and `Update_ButtonClick` call `Validate` first and skip the
  whole operation on failure — `:153-167`, `:185-201`).
- **Search "not found"** — `:233`: `ShowMessage('No matches were found for your query.')`.
- **Search/SetDetailBoxes error** — `:272-273` and `:309-311`: both wrapped in `try/except` that
  `ShowMessage('Error in Search'/'Error in SetDetailBoxes' + #13 + e.Message)`.
- Grid carries `dgConfirmDelete` (`SupplierMaster.dfm:424`) — same VCL-native secondary confirm path
  as other masters, independent of the `Delete_Button` handler.
- **Directory picker** — `Breakdown_SpeedButtonClick` (`:382-394`) uses `SelectDirectory` (native
  Windows folder browser), defaulting to the current `Directory_Edit.Text` if it exists on disk,
  else `'c:\'`. Not a MessageDlg but a modal OS dialog worth noting for a web port (no browser
  equivalent — needs a design decision).

## Field clear / repopulate
- **`FormCreate`** (`:140-151`): loads three combo lookups (`Logistics_ComboBox`,
  `CreateOrderSheet_ComboBox`, `InventoryAddPoint_ComboBox`) via `SelectSingleField`, then
  `GetSupplierInfo`, binds `Supplier_DataSource`, sets `SupplierCode_Edit.Text := ''`, and
  **`Data_Module.ClearControls(SupplierMaster_Panel)`** — unlike `SizeMaster`, this form DOES clear
  the whole panel on create.
- **`FormShow`** (`:419-424`): `SupplierCode_Edit.SetFocus` THEN `HoldDetails(True)` THEN
  `SetDetailBoxes` — i.e. **on every show, it reads the grid's CURRENT row (whatever the DBGrid
  defaults to — normally its first record after binding) into `Data_Module` and repopulates the
  detail panel from it**, overwriting the blank state `FormCreate` just set up. This is the #135-class
  hazard in its purest form here: if the grid has ANY rows, the panel is force-populated with the
  first row's data on open (not blank), even though `FormCreate` explicitly blanked it moments
  earlier. If the grid is EMPTY (no supplier rows), `HoldDetails(True)`'s `with
  SupplierMaster_DBGrid.DataSource.DataSet do Fields[0].AsString` etc. would run against a dataset
  with no active record — **not guarded**; behavior on an empty table is unverified (likely raises
  or returns blank/'0'-per-field depending on ADO cursor state — body unverified beyond this file).
- **`SetDetailBoxes`** (`:278-313`) special-cases the empty-logistics-combo case in `HoldDetails`
  (not `SetDetailBoxes`) — see below — but itself has no empty-guard; wrapped in `try/except` that
  silently `ShowMessage`s on any field-access error (no partial-fill guarantee: if field N throws,
  fields 1..N-1 already got written into their edits, field N+1.. do not).
  - `OutputFileType_RadioGroup` decode: `'TEXT'→0`, `'EXCEL'→1`, else→`2` ("Both") (`:296-301`) —
    note this compares against the FULL words `'TEXT'`/`'EXCEL'`, but `HoldDetails` (write path)
    encodes back as single letters `'T'`/`'E'`/`'B'` (`:365-370`) — **an asymmetric round-trip**: the
    grid's raw `SupplierOutputFileType` field (from `Fields[12].AsString`) is single-letter-coded per
    the DB, so `SetDetailBoxes`'s `'TEXT'`/`'EXCEL'` string comparison against a single-letter DB value
    will NEVER match `'TEXT'` or `'EXCEL'` and will always fall through to the `else` branch
    (`ItemIndex := 2`, "Both") when populating from a grid-sourced value. This looks like a live bug:
    confirm against the actual stored `VC_...OUTPUT_FILE_TYPE` values before treating "Both" as the
    silently-defaulted display for every existing supplier row. *(Confidence: pattern read directly
    from source; DB column content not independently queried — flag before relying on it.)*
- **Empty-combo-value normalization** — `HoldDetails(fFromGrid=False)` (`:359-363`): `// dose this
  due to empty string bug` — `if (Logistics_ComboBox.Text=' ') or (Logistics_ComboBox.Text='') then
  LogisticsName := '' else LogisticsName := Logistics_ComboBox.text` — an explicit legacy workaround
  for a combo-box holding a single-space sentinel string as its "no selection" value. Any port must
  reproduce the `' '` (single space) vs `''` (empty) distinction the combo population code
  (`SelectSingleField`) apparently seeds as a blank-row placeholder.
- **On Insert** (`:153-167`): `HoldDetails(False)` → insert → `GetSupplierInfo` (requery) →
  `Locate('Supplier Code', SupplierCode_Edit.Text, [])` → `SetDetailBoxes` (echoes inserted row) →
  `SupplierCode_Edit.SelectAll` + `SetFocus`.
- **On Update** (`:185-201`): same pattern but `Locate(...,[lopartialkey])` (partial-key match, unlike
  Insert's exact match) — a locate-option asymmetry between Insert and Update worth flagging if a
  rebuild needs exact parity on partial vs exact key matching.
- **On Delete** (`:203-219`): `SupplierCode_Edit.Text := ''` cleared, `GetSupplierInfo` (requery),
  then `SearchGrid(Data_Module.SupplierCode)` — same stale-code-search pattern as `SizeMaster`
  (searches for the just-deleted code, will not find it, so `SetDetailBoxes` isn't reached via
  `SearchGrid`'s success path) — **panel fields other than the code box are left stale after delete**
  (same hazard class as `SizeMaster.pas:123-129`).
- **`Clear_Button`** (`:240-245`): blanks code field, sets focus, `ClearControls(SupplierMaster_Panel)`
  — a full, correct blank (this is the one button that reliably clears everything).

## Focus & keyboard
- `FormShow` (`:419`) sets focus to `SupplierCode_Edit` first, but the subsequent `HoldDetails(True)`
  + `SetDetailBoxes` calls don't move focus again — so the field-clear-then-repopulate sequence
  above happens **after** focus is already placed, not before.
- No VCL `Default`/`Cancel` button flags; `Close_Button` has `ModalResult = 2` only
  (`SupplierMaster.dfm:84`). All navigation via mnemonics.
- Grid `OnKeyUp`/`OnMouseUp` (`:396-410`) both call `HoldDetails(True)` + `SetDetailBoxes` — same
  click-or-keystroke re-sync pattern as `SizeMaster`.
- `Supplier_DataSourceDataChange` (`:412-417`) mirrors the grid handlers on any dataset navigation.

## Enable/disable state machine
- **None found** — no `Enabled` toggling anywhere on this form; all buttons always clickable
  regardless of grid/selection state, same as `SizeMaster`.

## Error surfacing
- `SetDetailBoxes` and `SearchGrid` both wrap in `try/except` → `ShowMessage`. `Insert_/Update_/
  Delete_ButtonClick` have no try/except of their own (same gap as `SizeMaster`) — a DB-layer
  exception from `InsertSupplierInfo`/`UpdateSupplierInfo`/`DeleteSupplierInfo` is unguarded here.

## Cross-refs
- `docs/analysis/master-data/supplier.md` (proc/data spec).
