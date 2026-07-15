# Form-UX semantics: `TRenbanGroupMaster_Form` — `RenbanGroupMaster.pas` / `RenbanGroupMaster.dfm`

CRUD editor over the renban-group table (group code, renban count, ship-days override +
per-weekday ship-days). Grid `RenbanGroupMaster_DBGrid` bound via `Renban_DataSource`.

## Dialogs & confirmations
- **Insert failure** — `RenbanGroupMaster.pas:230-232`: `MessageDlg('Unable to INSERT ' +
  Data_Module.RenbanCode + #13 + 'It already exists in the database.', mtInformation, [mbOk], 0)` if
  `InsertRenbanGroupInfo` fails — same "already exists" wording family as `SizeMaster`.
- **Delete confirmation (armed, two-step)** — `:300-305`: `MessageDlg('Are you sure you wish to
  delete' + #13 + Data_Module.RenbanCode + ' from the database?', mtWarning, [mbYes, mbNo], 0) =
  mrYes` gates `DeleteRenbanGroupInfo`. Note: **no requery/re-search or field-clear at all in this
  handler** — after a successful delete it only calls `Data_Module.GetRenbanGroupInfo` (requery) and
  then `RenbanGroupCode_Edit.SelectAll` + `SetFocus` (`:306-311`) — the detail panel (count/ship-days
  fields) is left showing whatever was there before the delete; there is no `SetDetailBoxes` or
  `ClearControls` call in this handler at all, the strongest stale-panel-after-delete case in the
  family (SizeMaster/SupplierMaster at least attempt a `SearchGrid` afterward).
- **`Validate` (`:241-283`)** — 7 checks, all sharing ONE accumulator string `fErrMsg` (unlike
  `PartsStockMaster`'s fail-fast-on-first-field style): `RenbanGroupCount_Edit`,
  `ShipDays_Edit`, and each weekday `ShipDays*_MaskEdit` (Mon–Sat) are checked via `TryStrToInt`; **each
  failing check simply OVERWRITES `fErrMsg`** (`fErrMsg := #13 + 'Invalid Monday Ship Days'` etc,
  `:254-273`) rather than appending, so **only the LAST failing field's message survives** to be
  shown — if both `ShipDays_Edit` and `ShipDaysSaturday_MaskEdit` are bad, the dialog reports only
  "Invalid Saturday Ship Days" even though the real cause could be the Monday-equivalent
  `ShipDays_Edit` failing first. The combined message is prefixed `'The following fields need to be
  corrected:'` (`:277`) but in practice only ever names one field. Focus always lands on
  `RenbanGroupCount_Edit` (`:279`) regardless of which check actually failed — **not
  field-specific focus at all**, unlike `PartsStockMaster`'s (mostly) per-field `SetFocus`.
- **Search "not found"** — `:325`: `ShowMessage('No matches were found for your query.')`.
- **Search error** — `:205-206`: `ShowMessage('Error in Search' + #13 + e.Message)`.
- Grid carries `dgConfirmDelete` (`RenbanGroupMaster.dfm:229`) — VCL-native secondary confirm.

## Field clear / repopulate
- **`RenbanCount` zero-padding on write** (`HoldDetails(False)`, `:143-148`): pads
  `RenbanGroupCount_Edit.Text` to 3 digits with leading zeros depending on its current length (1→
  `'00'+x`, 2→`'0'+x`, 3→as-is) — **no `else` for length 0 or >3**: if the edit is empty (length 0),
  none of the three `if`/`else if` branches match, and `Data_Module.RenbanCount` is **left holding
  whatever its previous value was** (not reset to `''` or `'000'`) — a silent no-op on empty input
  that could let a stale `RenbanCount` slip into an Insert/Update after a `Clear_Button` press
  followed immediately by an Insert without retyping the count (Clear does set it to `'0'`,
  `:336`, which is length 1 → padded to `'00'`, so this specific edge only bites if the field is
  made truly empty by some other path, e.g. select-all+delete).
- **Numeric ship-day fields tolerate non-numeric input silently** (`HoldDetails(False)`,
  `:150-183`): unlike `PartsStockMaster`'s raw `StrToInt` (which raises), every ship-days field here
  uses `TryStrToInt(...) ... Else ShipDaysX := 0` — a non-numeric ship-days value is silently
  coerced to `0` rather than raising OR blocking the write (this is a DIFFERENT, more permissive
  strategy than `Validate`'s own numeric check on the same fields — `Validate` would already have
  caught and blocked a non-numeric entry before `HoldDetails` runs, on the Insert/Update paths; but
  `HoldDetails(false)` is ALSO called by nothing else that bypasses `Validate`, so this fallback
  path is presently unreachable through the UI as far as this file shows — flag as defensive
  dead code, not a live divergence).
- **`RenbanGroupCode_EditChange`** (`:361-366`): typing the code down to ≤1 character clears the
  whole panel (`ClearControls(RenbanGroupMaster_Panel)`) — same live-typing clear-as-you-type pattern
  as `SizeMaster.SizeCode_EditChange`.
- **`FormCreate`** (`:347-359`): unfilter, `GetRenbanGroupInfo`, bind `Renban_DataSource`,
  `ClearControls(RenbanGroupMaster_Panel)`, re-unfilter, `RenbanGroupCode_Edit.Text := ''`.
- **`FormShow`** (`:375-379`): `RenbanGroupCode_Edit.SetFocus` THEN `Data_Module.Inv_DataSet.First`
  — **no explicit `SetDetailBoxes` call**; like `ManifestCostMaster`, this form relies entirely on
  `Inv_DataSet.First` firing `Renban_DataSourceDataChange` (`:368-373`, which does call
  `HoldDetails(True)` + `SetDetailBoxes`) to populate the panel on open. If the table is empty,
  `.First` is a no-op and the panel stays however `FormCreate`'s `ClearControls` left it (correctly
  blank in that case) — same DB-dataset-driven initial-populate pattern as `ManifestCostMaster`, and
  the ONE OTHER form in this family besides it that behaves this way (not form-code-driven for
  initial show).
- **On Insert** (`:225-239`): `Validate` → `HoldDetails(False)` → insert → `GetRenbanGroupInfo`
  (requery) → re-unfilter → `RenbanGroupCode_Edit.Text := ''` (explicit blank — this form DOES clear
  the key field after insert, unlike its detail fields) → `SetFocus`. **No `SetDetailBoxes` call
  after insert** — unlike every other master form, the just-inserted row's values are NOT echoed
  back into the panel; the code edit is simply blanked, ready for the next entry.
- **On Update** (`:285-298`): `Validate` → `HoldDetails(False)` → requery → `Locate('Renban Group
  Code', renban, [])` → `SetDetailBoxes` (echoes updated row) — Update DOES repopulate, Insert does
  NOT; an asymmetry unique to this form in the family.
- **`Clear_Button`** (`:332-345`): blanks code field, resets `RenbanGroupCount_Edit`/`ShipDays_Edit`/
  all 6 weekday ship-days edits to `'0'` individually (same manual-reset style as
  `PartsStockMaster.Clear_ButtonClick`, not relying solely on `ClearControls`), focus to code field.
  Note: does NOT call `Data_Module.ClearControls` at all here (unlike `FormCreate`/
  `RenbanGroupCode_EditChange`) — relies entirely on the manual per-field resets.

## Focus & keyboard
- `FormShow` sets focus to `RenbanGroupCode_Edit` (`:377`) BEFORE `Inv_DataSet.First` runs — focus
  is placed, then the panel populates via the `OnDataChange` cascade afterward (order doesn't affect
  which control is focused, since none of that cascade calls `SetFocus`).
- No VCL `Default`/`Cancel` flags; `Close_Button` has `ModalResult = 2` only
  (`RenbanGroupMaster.dfm:289`). Accelerators only.
- Grid `OnMouseUp`/`OnKeyUp` (`:210-223`) both call `HoldDetails(True)` + `SetDetailBoxes` — standard
  click-or-keystroke re-sync; `Renban_DataSourceDataChange` (`:368-373`) mirrors on dataset
  navigation and is, per above, the form's actual initial-populate mechanism.
- Every field on this form's detail panel (`RenbanGroupCode_Edit`, `RenbanGroupCount_Edit`,
  `ShipDays_Edit`, all 6 weekday ship-days) has `Tag = 1` set in the `.dfm`
  (`RenbanGroupMaster.dfm:120,131,141,151,163,175,187,199,211`) — a design-time marker with no
  handler in this `.pas` reading `Tag` anywhere; likely a vestige of a shared/generic field-
  iteration helper used elsewhere in the app (e.g. `ClearControls` may filter on `Tag`) —
  **`Data_Module.ClearControls`'s exact `Tag`-driven behavior is in `DataModule.pas`, not this
  form; treat as [UNVERIFIED — confirm before use] if a rebuild wants to replicate WHICH fields
  `ClearControls` actually touches.**

## Enable/disable state machine
- **None found** — no `Enabled` toggling; all buttons always clickable.

## Error surfacing
- Only `SearchGrid` wraps in `try/except` → `ShowMessage`. `Insert_/Update_/Delete_ButtonClick` have
  no try/except of their own.

## Cross-refs
- No `docs/analysis/master-data/renban-group.md` found at time of writing (only `logistics.md`,
  `manifest-cost.md`, `master-maint.md`, `size.md`, `supplier.md` exist in that directory) — this is
  a placeholder cross-ref pending that proc/data spec being written. See also
  `docs/analysis/master-data/IGNITION-master-crud-design.md` for any cross-form CRUD design notes
  that may already cover Renban Group.
