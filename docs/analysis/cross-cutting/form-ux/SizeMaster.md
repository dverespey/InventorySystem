# Form-UX semantics: `TSizeMaster_Form` — `SizeMaster.pas` / `SizeMaster.dfm`

CRUD editor over `INV_SIZE_MST` (size code, size name, daily usage, safety days) with a live
`TDBGrid` bound to `Data_Module.Inv_DataSet` via `Size_DataSource`.

## Dialogs & confirmations
- **Insert failure** — `SizeMaster.pas:93-95`: if `Data_Module.InsertSizeInfo` returns `False`,
  `MessageDlg('Unable to INSERT ' + Data_Module.SizeCode + #13 + 'It already exists in the
  database.', mtInformation, [mbOk], 0)`. Single-OK information dialog; duplicate-key is the only
  documented cause.
- **Delete confirmation (armed, two-step)** — `SizeMaster.pas:118-129`:
  `MessageDlg('Are you sure you wish to delete' + #13 + Data_Module.SizeName + ' (' +
  Data_Module.SizeCode + ') from the database?', mtWarning, [mbYes, mbNo], 0) = mrYes`. Only on
  `mrYes` does it call `Data_Module.DeleteSizeInfo`, clear the code field, refresh, and re-search.
  On `mrNo` (or dialog dismissed) nothing happens except focus moves to `SizeCode_Edit`.
- **Search "not found"** — `SizeMaster.pas:146`: plain `ShowMessage('No matches were found for your
  query.')`, single OK.
- **Search-time error** — `SizeMaster.pas:184-187`: `ShowMessage('Error in Search' + #13 +
  e.Message)` inside `SearchGrid`'s except block.
- Grid itself carries `dgConfirmDelete` (`SizeMaster.dfm:192`) — the DBGrid's own built-in
  Ctrl+Delete/grid-delete confirmation is ALSO active independent of the `Delete_Button` handler
  above (a second, VCL-native confirm path on the grid row itself, not wired to any custom handler
  here — standard VCL "Delete this record?" if triggered via grid key).

## Field clear / repopulate
- `SizeCode_EditChange` (`:160-165`): **as soon as the operator types down to ≤1 character** in the
  code field, `Data_Module.ClearControls(SizeMaster_Panel)` wipes the whole detail panel — a
  live-typing side effect, not just on blur/search.
  - `TextChange` (`:237-243`) is wired (via `.dfm`, see below) to the two `TMaskEdit` fields
    (`DailyUsage_MaskEdit`, `SafetyDays_MaskEdit`): if the trimmed text is empty, it's forced back
    to `'0'` — **empty numeric input auto-defaults to zero on every keystroke**, never left blank.
  - `HoldDetails(fFromGrid=False)` (`:216-224`) reads the edits into `Data_Module` fields with a raw
    `StrToInt(Trim(...))` — NOT `TryStrToInt` — so if a mask edit somehow held a non-numeric value at
    Insert/Update time, this raises an unhandled exception (no `Validate` function exists on this
    form at all, unlike `PartsStockMaster`/`RenbanGroupMaster`).
- **On Insert** (`:90-102`): `HoldDetails(False)` → insert → `GetSizeInfo` (full requery) →
  `Inv_DataSet.Locate(...)` on the code just entered → `SetDetailBoxes` (repopulates from
  `Data_Module`, i.e. echoes back what was inserted) → `SizeCode_Edit.SelectAll` + `SetFocus`. Fields
  are NOT blanked after insert — they're refreshed to the just-inserted row's values.
- **On Delete** (`:123-127`): `SizeCode_Edit.Text := ''` is explicitly cleared, then
  `Data_Module.GetSizeInfo` + `SearchGrid(Data_Module.SizeCode)` — note `Data_Module.SizeCode` at
  this point still holds the just-deleted row's code (set by the earlier `HoldDetails(False)` at
  `:118`), so `SearchGrid` searches for the code that no longer exists → `RecordCount` will be 0 →
  `SetDetailBoxes` is NOT called (no `SearchGrid` success) → **the detail panel is left showing the
  deleted record's stale values** except for `SizeCode_Edit` which was blanked. This is the
  empty-repopulate-after-delete hazard class (siblings of #135): the grid re-renders correctly, but
  `SizeName_Edit`/`DailyUsage_MaskEdit`/`SafetyDays_MaskEdit` do NOT get re-cleared here — only
  `Clear_Button` or a new grid selection will refresh them.
- **`FormCreate`** (`:227-235`): `Inv_DataSet.Filter := ''`, `Filtered := FALSE`, `GetSizeInfo`,
  bind `Size_DataSource`, `Filtered := False` again, `SizeCode_Edit.Text := ''`. Note: does NOT call
  `ClearControls` here (unlike `SupplierMaster`/`RenbanGroupMaster`/`PartsStockMaster` `FormCreate`s)
  — `SizeName_Edit`/mask-edits retain whatever the `.dfm` design-time defaults are until `FormShow`
  runs.
- **`FormShow`** (`:267-271`): unconditionally calls `SetDetailBoxes` then `SizeCode_Edit.SetFocus`
  — `SetDetailBoxes` reads current `Data_Module` property values (which are **whatever was left over
  from the previous invocation of this form in the same session**, since `Data_Module` fields are
  module-level, not form-owned) into the edits. **If this is the session's first open and no prior
  master screen touched these `Data_Module` properties, the fields show Delphi's default
  empty-string/zero values; if a PRIOR SizeMaster session left stale values in `Data_Module`, they
  reappear here** — this is the #135-class hazard: FormShow repopulates from shared mutable state,
  not from "no selection = blank."

## Focus & keyboard
- No `ActiveControl` set in `.dfm`; `FormShow` explicitly sets focus to `SizeCode_Edit`
  (`:270`), overriding the natural tab-order default of `Insert_Button` (`TabOrder=0` in
  `ManagementButtons_Panel`) since `SizeCode_Edit` sits in a separate panel with its own `TabOrder=0`.
- `Close_Button` has `ModalResult = 2` (`mrCancel`, `SizeMaster.dfm:84`) but no `Cancel=True`/
  `Default=True` flags set on any button — Esc/Enter are not VCL-wired; only accelerators
  (`&Insert`, `&Update`, `&Search`, `Cl&ear`, `&Close`, `&Delete`) work as shortcuts.
- Grid mouse-up (`SizeMaster_DBGridMouseUp`, `:245-250`) and grid key-up
  (`SizeMaster_DBGridKeyUp`, `:253-258`) both call `HoldDetails(True)` + `SetDetailBoxes` — so
  **both a click AND any keystroke** (including arrow-key navigation) on the grid re-syncs the detail
  panel from the grid's current row.
- `Size_DataSourceDataChange` (`:260-265`) does the same on any dataset navigation (e.g. programmatic
  `Locate`) — this is the mechanism by which Insert/Update/Delete's `Locate` calls repopulate the
  panel, **triggered by the DataSource, not by explicit form code** calling `SetDetailBoxes` a second
  time (though the Insert/Update handlers also call `SetDetailBoxes` explicitly right after `Locate`,
  making it redundant-but-harmless here).
- `SizeCode_Edit` has `CharCase = ecUpperCase`, `MaxLength = 6` (`.dfm:147-148`).

## Enable/disable state machine
- **None found.** No button or edit has its `Enabled` toggled anywhere in `SizeMaster.pas` — Insert/
  Update/Delete/Search/Clear are always enabled regardless of grid state (e.g. even with an empty
  grid, all buttons remain clickable; `Update`/`Delete` on a non-existent record depend entirely on
  whatever stale `Data_Module` values are in memory, since there's no "is a row selected" gate).

## Error surfacing
- Two `try/except` blocks (`SearchGrid` and the outer `Execute`) surface errors via `ShowMessage`/
  `showMessage` dialogs (see Dialogs above). `Insert_ButtonClick`, `Update_ButtonClick`,
  `Delete_ButtonClick` have **no try/except of their own** — an exception inside
  `Data_Module.InsertSizeInfo`/`UpdateSizeInfo`/`DeleteSizeInfo` (e.g. a SQL error) propagates
  unguarded up to the VCL message loop's default `Application.HandleException` (a generic Delphi
  error box), not a form-authored message.

## Cross-refs
- `docs/analysis/master-data/size.md` (proc/data spec).
