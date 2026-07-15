# Form-UX: `TRecConfStat_Form` — `RecConfStat.pas` + `RecConfStat.dfm`

> Wide editable grid over `INV_OPEN_ORDER_INF` (Receiving Confirmation Status): insert/edit/delete
> open orders, stamp logistics milestones, batch-edit a RENBAN group. Confidence: high — both files
> read in full (904 + 809 lines).
> Business/proc spec: [`../../receiving/recconfstat.md`](../../receiving/recconfstat.md).

## Dialogs & confirmations
- **`Delete_ButtonClick`** (`:147-160`) — **armed confirm:**
  `MessageDlg('Are you sure you wish to delete'+#13+FRSNo+' ('+PartNum+') from the database?',
  mtWarning, [mbYes, mbNo], 0) = mrYes`. Only on `mrYes` calls `DeleteRecConfStatInfo`. `mtWarning`
  icon (not `mtConfirmation`) — the only form in this family whose delete-confirm uses the Warning
  icon rather than Confirmation.
- **`Insert_ButtonClick`** (`:162-190`) — no confirmation; on `InsertRecConfStatInfo` returning
  `False` (the P1 in-proc dup-guard rejecting an existing 4-tuple), shows a detailed
  `MessageDlg` (not `ShowMessage`) with **`mtInformation`**:
  `'Unable to INSERT Supplier Code: '+SupplierCode+#13+'Parts Code: '+PartNum+#13+'FRS Number: '+
  FRSNo+#13+'RENBAN: '+Renban+#13+'It already exists in the database.'` (`:172-176`) — the most
  detailed error dialog in this family (echoes back all four key values).
- **`Update_ButtonClick`** (`:192-214`) — no confirmation, including for the **RENBAN batch
  update** path (`RenbanUpdate_CheckBox.Checked` → `UpdateRecConfStatRenbanInfo`, which per the
  business spec can re-balance stock on *every* row sharing the RENBAN) — a batch operation with
  wide blast radius triggered with the same unconfirmed click as a single-row edit.
- **`Search_ButtonClick`** (`:216-269`) — validation-only, not a confirm: if every search field is
  blank and "No Order Search" is unchecked,
  `ShowMessage('At least one of the following fields'+#13+'must contain data before searching:'+#13+
  '    Supplier Code'+#13+'    Parts Code'+#13+'    Kanban Code'+#13+'    FRS Number'+#13+
  '    Order Date'+#13+'    Ship Date'+#13+'    RENBAN Number'+#13+'OR the No Order Search must be
  selected')` (`:223-232`); a search yielding zero rows shows `'No matches were found for your
  query.'` (`:265`).
- **`InTransit_NUMMIBmDateEditExit`** (`:667-687`) — date-format validation, not a confirm:
  `MessageDlg('Please enter a valid date.', mtError, [mbOK], 0)` (`:679`) if the typed text parses
  to a zero date; refocuses the same field. Its own exception handler is **empty**
  (`except on e:exception do end`, `:682-684`) — any error inside this validator is silently
  swallowed.
- **`Validate`** (`:782-848`, called by both Insert and Update before `HoldDetails`) — five
  sequential blocking `ShowMessage`s, first-failure-wins, each followed by `SetFocus` to the
  offending control and `exit`: `'Supplier Code must be 5 characters'` (`:790`), `'Parts Code must
  be 12 characters'` (`:797`), `'FRS number must be 7 characters'` (`:804`), `'Quantity must be a
  numeric'` (`:811`), and the **in-transit precondition** (four variants, one per dependent field):
  `'Order must be marked In Transit when arrival is set'` (`:820`), `'... when warehouse is set'`
  (`:827`), `'... when plant yard is set'` (`:834`), `'... when assembler yard is set'` (`:841`).
- **No confirmation for `HideTerminated_CheckBoxClick`, `SortBy_ComboBoxChange`, or the RENBAN
  checkbox itself** — these are pure filter/mode toggles, immediate effect, no dialog.

## Field clear / repopulate
- **`SetDetailBoxes`** (`:436-547`) is the row→controls repopulate, called after every grid
  selection, Insert, Update, and Search-hit. For each of the ~10 date-stamp fields
  (`InTransit`/`Arrival`/`PlantYard`/`AssemblerYard`/`EmptyTrailer`/`Warehouse`/`Order`/`Ship`/
  `Terminated`), the pattern is uniform: **if the 8-char `yyyymmdd` string is non-blank, reformat
  and set the date picker's `.Date`; else set the picker's `.Text:=''`** — i.e. **an empty incoming
  value maps to a genuinely empty date-edit, not today's date or a placeholder** (contrast the
  Clear-button path below, which uses `SetTodaysDate`). This is the #135-relevant class: confirm the
  rebuild's date-picker component renders "no value" identically for a truly-empty string vs. an
  unset/null field.
- **Every call to `SetDetailBoxes` also unconditionally resets** `RenbanUpdate_CheckBox.Enabled:=
  False` + `.Checked:=False` and `Unordered_Box.Enabled:=False` + `.Checked:=False` (`:538-542`) —
  selecting/loading any row **always** clears the RENBAN-batch-update opt-in, forcing the operator
  to re-trigger it (via `TrailerNo_EditChange`, see below) for every edit.
- **`Clear_ButtonClick`** (`:271-292`) — the "new order" reset: `SetTodaysDate(RecConfStat_Panel)`
  (walks the panel and sets **every** `TDateTimePicker` to `Date` — **not** the `NUMMIBmDateEdit`
  milestone fields, which are a different control type not touched by this DataModule helper),
  then `ClearControls(RecConfStat_Panel)` + `ClearControls(SearchKey_GroupBox)` (the shared
  DataModule walker — blanks `TEdit`/`TMaskEdit`/`TMemo`/`TNUMMIBmDateEdit` to `''`, `TCheckBox` to
  `False`, **`TComboBox.ItemIndex:=0`** — note combos reset to index 0, i.e. whatever `SelectSingle
  Field` populated as the first item, conventionally the literal single-space placeholder, not a
  true "nothing selected" state). Then explicitly re-sets `SupplierCode_ComboBox`/`PartsCode_
  ComboBox`/`KanbanCode_ComboBox`/`AssemblerLocation_ComboBox.ItemIndex:=-1` (`:283-286`) —
  **overriding** `ClearControls`'s `ItemIndex:=0` with **`-1` (no selection at all)** for these four
  specifically. **This double-reset (0 then -1) is a real divergence hazard**: a rebuild that only
  implements the generic `ClearControls` behavior (index 0 / blank-placeholder) without this
  form-specific `-1` override would show a stale first-item selection instead of a truly empty combo.
- **`FormCreate`** (`:294-321`) calls `SetTodaysDate` + `GetRecConfStatInfo` then immediately
  `ClearControls` on both panels — the grid loads fully populated, but the detail/search panels
  start blank (no row auto-selected on open).
- **`SearchGrid`'s no-match branch** (`:626-639`) explicitly re-clears both panels and the three
  cascading combos to `-1`, plus re-arms `Unordered_Box.Enabled:=TRUE` — mirrors `Clear_
  ButtonClick`'s reset almost exactly (duplicated logic, not shared).
- **`HoldDetails(False)`** (manual-entry capture, `:358-433`) treats a literal `' '` (single space)
  in `SupplierCode_ComboBox`/`PartsCode_ComboBox`/`KanbanCode_ComboBox` as equivalent to blank
  (`if X.Text = ' ' then Y := '' else Y := X.Text`, `:362-375`) — this is the DataModule
  placeholder-string convention (see Cross-refs) made explicit at the call site.

## Focus & keyboard
- **No `FormShow`-driven `SetFocus`... actually it does:** `FormShow` (`:763-780`) calls
  `SearchGrid`, resets the two batch-checkboxes, clears both panels, **then `Renban_Edit.SetFocus`**
  (`:772`), sets the three dynamic labels (Assembler/Plant names from site config), and finally
  calls `Clear_ButtonClick(self)` — which itself ends in `Renban_Edit.SetFocus` again (`:287`) — so
  initial focus is `Renban_Edit`, set twice.
- **Post-Insert/-Update/-Search focus:** `Renban_Edit.SetFocus` in all three (`:188`, `:212`, `:268`)
  — the operator is always returned to the RENBAN field to start the next search/entry, regardless
  of which action just ran.
- **`ASSEMBLERLocation_ComboBoxChange`** (`:850-854`) — selecting an assembler location blanks
  `ParkingSpot_Edit.Text:=''` and **auto-stamps** `AssemblerYard_NUMMIBmDateEdit.Date :=
  GetLastProductionDate` — a side-effecting `OnChange`, not a pure display update: choosing a
  location silently sets/overwrites the assembler-yard date.
- **`TrailerNo_EditChange`** (`:856-873`) — shared handler across the trailer-number edit **and**
  the assembler-location combo (it dispatches by `Sender is TComboBox` + name check, `:859-862`).
  Its real job: **`RenbanUpdate_CheckBox.Enabled:=True`** (`:858`) — the RENBAN-batch checkbox only
  becomes available for the operator to check **after** a trailer-number change (per
  `recconfstat.md` §5) — a specific, easy-to-miss precondition for a bulk-affecting action.
- **No `Default`/`Cancel` flags, no `KeyPreview`** anywhere in `RecConfStat.dfm` (grep-verified) —
  Enter does not trigger Insert/Update/Delete/Search from anywhere on this form.
- **All ~10 date-picker child controls carry `Options = [doButtonTabStop, doCanClear, doCanPopup,
  doIsMasked, doShowCancel, doShowToday]`** (`RecConfStat.dfm`, e.g. `:353`) — `doCanClear` means the
  operator can explicitly blank a milestone date via the picker's own UI (right-click/clear
  button), independent of any app-level "clear" action — a component-level escape hatch the
  rebuild's date-input equivalent should preserve (clearing a milestone is how the D8(3) arrival-
  reversal gets triggered in practice).

## Enable/disable state machine
- **`RenbanUpdate_CheckBox`**: starts `Enabled:=False`/`Checked:=False` (`.dfm`; also reset by
  every `SetDetailBoxes`/`Clear_ButtonClick`/`SearchGrid`-no-match call); becomes `Enabled:=True`
  only after `TrailerNo_EditChange` fires (a trailer-number **or** assembler-location edit).
  Checking it changes `Update_ButtonClick`'s dispatch target from `UpdateRecConfStatInfo` (single
  row) to `UpdateRecConfStatRenbanInfo` (whole-RENBAN batch) — **the state machine, not a
  confirmation dialog, is the only gate on this high-blast-radius batch update.**
- **`Unordered_Box` ("No Order Date" / "No Order Search")**: starts `Enabled:=True`/`Checked:=False`;
  disabled (`Enabled:=False`) after a successful search or field-load (`SetDetailBoxes`,
  `SearchGrid`'s match branch is silent on it but the no-match branch re-enables it, `:638`) — so it
  is only checkable when the panel is in a "fresh search" state, not while viewing a loaded/selected
  row.
- **`HideTerminated_CheckBox`**: gated by the private `fCheck` flag (`:878` — set `True` only after
  `FormCreate` finishes, `:313`) so the checkbox's own `OnCreate`-time default assignment
  (`HideTerminated_CheckBox.Checked:=fiHideTerminated.AsBoolean`, `:311`) doesn't re-trigger a
  premature `SearchGrid` before the form is ready.
- **No Insert/Update/Delete button ever disables based on row-selection state** — all three are
  always clickable; `Insert`/`Update` differ only by which key-tuple `HoldDetails` captured last
  (from the grid vs. the manual-entry fields), not by button enablement.

## Error surfacing
- **Mix of `ShowMessage` (info-only) and `MessageDlg` (icon + button-set)** — this is the only form
  in the family that varies the dialog *icon* meaningfully: `mtWarning` for delete-confirm,
  `mtInformation` for the duplicate-key insert failure, `mtError` for the invalid-date validator,
  `mtConfirmation` is never actually used here (the delete confirm uses `mtWarning` instead, despite
  being a Yes/No choice — a naming/semantics mismatch worth normalizing in the rebuild).
  `SearchGrid`'s try/except (`:641-644`) and the top-level `Execute` catch (`:135-140`) both surface
  as plain `ShowMessage`.
- **No inline per-field error text or status label anywhere** — all validation surfaces via
  `Validate`'s sequential `ShowMessage`+`SetFocus` chain.

## Cross-refs
- Business rules / procs / triggers (D7 arrival-add, D8(3) dead arrival-reversal, RENBAN batch
  blast radius, supplier-blind delete key):
  [`../../receiving/recconfstat.md`](../../receiving/recconfstat.md).
- Shared DataModule field-clear/combo-populate ancestor: `DataModule.pas:5976` `ClearControls`
  (walks a Panel/GroupBox, blanks `TEdit`/`TMaskEdit`/`TMemo`/`TNUMMIBmDateEdit`, `TCheckBox→False`,
  **`TComboBox.ItemIndex:=0`**, honoring a `Tag=1` opt-out per control) and `:5767`
  `SelectSingleField` (always inserts a literal `' '` as combo item 0 before the real values — the
  "blank" placeholder every combo in this codebase relies on).
