# Form-UX semantics: `TLogisticsMaster_Form` — `LogisticsMaster.pas` / `LogisticsMaster.dfm`

CRUD editor over `INV_LOGISTICS_MST` (logistics-company name/address/contact/breakdown directory).
Grid `LogisticsMaster_DBGrid` bound via `LogisticsMAster_DataSource` (note the source-name typo
`LogisticsMAster_DataSource`, carried verbatim from the `.dfm`/`.pas`).

## Dialogs & confirmations
- **Insert failure** — `LogisticsMaster.pas:257-258`: `MessageDlg('Unable to INSERT ' +
  Data_Module.LogisticsName, mtInformation, [mbOk], 0)` if `InsertLogisticsInfo` fails. Shortest of
  the family's insert-failure messages (no "already exists" suffix, no code echoed).
- **Delete confirmation (armed, two-step)** — `:232-234`: `MessageDlg('Are you sure you wish to
  delete' + #13 + Data_Module.LogisticsName + ' from the database?', mtWarning, [mbYes, mbNo], 0) =
  mrYes` gates `DeleteLogisticsInfo`.
- **Search "not found"** — `:184`: `ShowMessage('No matches were found for your query.')`.
- **Search error** — `:216-217`: `ShowMessage('Error in Search' + #13 + e.Message)` in `SearchGrid`'s
  except block.
- Grid carries `dgConfirmDelete` (`LogisticsMaster.dfm:313`) — same VCL-native secondary confirm.
- **Directory picker** — `Breakdown_SpeedButtonClick` (`:267-279`), identical pattern to
  `SupplierMaster`'s (`SelectDirectory`, defaults to current text if it's a real directory else
  `'c:\'`).

## Field clear / repopulate
- **`FormCreate`** (`:106-114`): unfilter dataset, `GetLogisticsInfo`, bind data source,
  `LogisticsName_Edit.Text := ''`, `Data_Module.ClearControls(LogisticsMaster_Panel)` — full blank,
  same as `SupplierMaster`/`RenbanGroupMaster`/`PartsStockMaster`.
- **`FormShow`** (`:304-308`): `SetDetailBoxes` then `LogisticsName_Edit.SetFocus` — **repopulates
  the panel from whatever `Data_Module.LogisticsName`/`LogisticsAddress`/etc. currently hold** (module-
  level shared state, not form-scoped) BEFORE moving focus. Unlike `SupplierMaster`'s `FormShow`
  (which calls `HoldDetails(True)` first to pull the grid's row into `Data_Module`), this form's
  `FormShow` does NOT re-read the grid — it trusts whatever is already in `Data_Module` from a prior
  screen/session. **This is a stronger version of the #135 hazard**: if `Data_Module.LogisticsName`
  etc. were left populated by an entirely different master screen earlier in the session (they are
  shared module-level properties, not Logistics-specific), this form's very first repaint on open
  could show a foreign master's leftover values, until the operator clicks the grid or types a search.
  *(Confidence: read directly off `FormCreate`/`FormShow`; whether `Data_Module` actually reuses
  the SAME string properties across different master forms is a `DataModule.pas` cross-cutting fact
  — flag before relying on it for a specific field name collision.)*
- **`SearchGrid`** (`:191-220`) always calls `Data_Module.ClearControls(LogisticsMaster_Panel)` at
  entry (`:196`) before attempting the linear `First`/`Next` scan for a name match — so a failed
  search DOES leave the panel blank (unlike `SizeMaster`'s post-delete stale-panel case), because
  `SearchGrid` clears unconditionally up front regardless of outcome.
- **On Insert** (`:254-265`): `HoldDetails(False)` → insert → `GetLogisticsInfo` (requery) →
  `Locate('LOGISTICS NAME', LogisticsName_Edit.Text, [])` → `SetDetailBoxes` (echoes inserted row) →
  `SelectAll` + `SetFocus`.
- **On Update** (`:245-252`): same requery/`Locate`/`SetDetailBoxes` pattern, no `SelectAll`.
- **On Delete** (`:229-243`): `LogisticsName_Edit.Text := ''` cleared, `GetLogisticsInfo` (requery),
  `SearchGrid(Data_Module.LogisticsName)` (still holds the deleted name from the prior
  `HoldDetails(False)` at `:231`) — will not find it, but since `SearchGrid` clears the panel
  unconditionally at its own entry (`:196`), the net effect here IS a full blank, unlike the
  `SizeMaster`/`SupplierMaster` delete paths which can leave stale non-code fields. This form's
  delete-then-clear is more complete than its siblings — note the difference explicitly if a
  rebuild wants one unified "post-delete" behavior across all master screens.
- **`Clear_Button`** (`:222-227`): blanks name field, sets focus, `ClearControls(...)` — full blank.
- No standalone `Validate` function on this form (unlike `SupplierMaster`/`RenbanGroupMaster`/
  `PartsStockMaster`) — Insert/Update proceed with whatever is in the edits, no numeric/length
  pre-check before calling the DataModule insert/update.

## Focus & keyboard
- `FormShow` sets focus to `LogisticsName_Edit` (`:307`) AFTER `SetDetailBoxes` runs (order matters
  only for which control paints focused, not for data — same ordering nuance as `SizeMaster`).
- No VCL `Default`/`Cancel` flags; `Close_Button` has `ModalResult = 2` only
  (`LogisticsMaster.dfm:294`). All navigation via mnemonics (`&Insert`, `&Update`, `&Search`,
  `Cl&ear`, `&Close`, `&Delete`).
- Grid `OnKeyUp`/`OnMouseUp` (`:282-295`) both call `HoldDetails(True)` + `SetDetailBoxes` — standard
  click-or-keystroke re-sync.
- `LogisticsMAster_DataSourceDataChange` (`:297-302`) mirrors on any dataset navigation.
- `BorderStyle = bsSingle` (`LogisticsMaster.dfm:5`) and `Position = poScreenCenter` (`:16`) — unlike
  every sibling master form in this family, which uses `bsDialog` + `poDesktopCenter`. Minor but a
  literal behavioral difference (this form is resizable/has a maximize-eligible border style,
  whereas the others are fixed dialog boxes) worth flagging for pixel/window-chrome parity, though
  likely irrelevant to a Perspective port.

## Enable/disable state machine
- **None found** — no `Enabled` toggling anywhere; all buttons always clickable.

## Error surfacing
- Only `SearchGrid` wraps in `try/except` → `ShowMessage`. `Insert_/Update_/Delete_ButtonClick` have
  no try/except of their own — same gap as `SizeMaster`/`SupplierMaster`.

## Cross-refs
- `docs/analysis/master-data/logistics.md` (proc/data spec).
