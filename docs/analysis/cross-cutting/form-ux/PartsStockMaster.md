# Form-UX semantics: `TPartsStockMaster_Form` — `PartsStockMaster.pas` / `PartsStockMaster.dfm`

CRUD editor over `INV_PARTS_STOCK_MST` — the largest master form in the family (30 detail fields:
part/supplier/logistics/renban/type/line/size/kanban/lot-qty/lead-time-per-weekday/ship-days-per-
weekday/quantity/comments/cost). Grid `PartsStockMaster_DBGrid` bound via `Parts_DataSource`.

## Dialogs & confirmations
- **Insert failure** — `PartsStockMaster.pas:170-172`: `MessageDlg('Unable to INSERT ' + #13 +
  'with Parts Code ' + PartNum, mtInformation, [mbOk], 0)` if `InsertPartsStockInfo` fails.
- **Delete confirmation (armed, two-step)** — `:203-208`: `MessageDlg('Are you sure you wish to
  delete' + #13 + 'supplier code' + Data_Module.SupplierCode + #13 + 'and part number ' +
  Data_Module.PartNum + #13 + 'from the database?', mtWarning, [mbYes, mbNo], 0) = mrYes` gates
  `DeletePartsStockInfo`.
- **`Validate` (`:282-422`) — the family's longest validation chain**, 14 sequential numeric checks,
  each independently gating with `ShowMessage` + `SetFocus` + `exit` (so only the FIRST failing
  field is ever reported per click — an operator must fix-and-resubmit repeatedly to walk through
  multiple bad fields, one at a time): `Quantity_MaskEdit` ("Quantity must be a numeric"),
  `OneLotQty_MaskEdit` ("Lot Qty must be a numeric"), `LeadTime_MaskEdit` ("Lead Time must be a
  numeric"), then Monday–Saturday `LeadTime*_MaskEdit` (each "Lead Time <Day> must be a numeric" —
  **but note all six weekday lead-time failures call `LeadTime_MaskEdit.SetFocus`, NOT the specific
  day's own edit control** — `:326,334,341,348,355,362` all read `LeadTime_MaskEdit.SetFocus` even
  though the failing field is e.g. `LeadTimeTuesday_MaskEdit`; this is a focus-target bug — the
  message names the right day but focus lands on the Monday/overall lead-time box instead**), then
  `ShipDays_Edit` ("Ship Days must be a numeric", focus correctly on itself), then Monday–Saturday
  `ShipDays*_MaskEdit` (each correctly self-targeting `SetFocus`, `:375-421` — the ship-days block
  does NOT have the same bug as the lead-time block above). `Insert_ButtonClick` and
  `Update_ButtonClick` both call `Validate` first and skip the whole operation on `False`
  (`:164-179`, `:186-197`).
- **Invalid part cost (non-blocking)** — `HoldDetails(False)` (`:539-548`): if the currency text
  (post `$`-strip) fails `TryStrToFloat`, `ShowMessage('Invalid part cost')` +
  `PartCost_MaskEdit.SetFocus` (note: references `PartCost_MaskEdit`, but the actual `.dfm` control
  is named `PArtCost_MaskEdit` — this line compiles only because Delphi identifiers are
  case-insensitive; **same non-blocking hazard as `ManifestCostMaster`**: no failure flag propagates
  to `Insert_ButtonClick`/`Update_ButtonClick`, which is called AFTER `Validate` already passed, so a
  bad cost value does not stop the write — `PartCost` simply keeps its previous value).
- **Search "not found"** — `:231`: `ShowMessage('No matches were found for your query.')`.
- **Search error** — `:636-637`: `ShowMessage('Error in Search' + #13 + e.Message)`.
- Grid carries `dgConfirmDelete` (`PartsStockMaster.dfm:787`) — VCL-native secondary confirm.

## Field clear / repopulate
- **`LotSizeOrders` is an INVERTED flag** (matches the documented project-wide inversion hazard):
  - Write (`HoldDetails(False)`, `:502`): `LotSizeOrders := not LotSizeOrders_CheckBox.Checked`.
  - Read (`SetDetailBoxes`, `:573`): `LotSizeOrders_CheckBox.Checked := not LotSizeOrders`.
  - The checkbox's on-screen semantics are the LOGICAL NEGATION of the stored/DB field — a rebuild
    that binds the checkbox directly to the raw column will show/save the opposite of the legacy
    checked-state. **Load-bearing: [UNVERIFIED against a live DB value] — confirm the actual stored
    bit for a known-good row before wiring a Perspective checkbox to this column.**
  - `LotSizeOrders_CheckBoxClick` (`:696-709`): when checked, `RenbanCode_ComboBox.Enabled := FALSE`
    and forced to `ItemIndex := 0` (blank); when unchecked, re-enabled and re-populated via
    `Data_Module.SearchCombo(RenbanCode_ComboBox, Data_Module.RenbanCode)`. This is a genuine
    enable/disable state transition (see below) coupled to the inverted flag.
- **Empty-combo normalization** (`HoldDetails(False)`, `:480-490`): `Logistics_ComboBox.Text = ' '`
  → stored as `''`; same for `RenbanCode_ComboBox.Text = ' '` → `''` — the same single-space
  sentinel workaround documented in `SupplierMaster.md`.
- **`FormCreate`** (`:265-280`): unfilter, `GetInventoryInfo`, bind `Parts_DataSource`,
  `JustifyColumns(PartsStockMaster_DBGrid)` (grid column-width normalization — DataModule-owned
  helper), `SetCombos` (loads Supplier/Logistics/Size/RenbanCode/PartType/Line lookups — six combo
  populations in one call, `:147-158`), re-unfilter, `ClearControls(PartsStockMaster_Panel)`,
  `PartsNum_Edit.Text := ''`.
- **`Clear_Button`** (`:236-263`): unfilters, `ClearControls(...)`, resets
  `Supplier_NUMMIColumnComboBox.ItemIndex`/`Logistics_ComboBox.ItemIndex` to `0`, AND explicitly
  re-zeroes all 13 numeric mask-edits (`LeadTime_MaskEdit` + 6 weekday lead-times + `ShipDays_Edit` +
  6 weekday ship-days) to `'0'` one by one (`:247-261`) — a longer manual reset than
  `Data_Module.ClearControls` alone provides for THIS form (the mask-edits apparently aren't fully
  reset by the generic helper, or the author didn't trust it to).
- **On Insert** (`:160-180`): `Validate` → `HoldDetails(False)` → insert → `GetInventoryInfo`
  (requery) → `Locate('Parts Code', part, [])` → `SetDetailBoxes` (echoes inserted row) →
  `PartsNum_Edit.SetFocus` (no `SelectAll` here, unlike `SizeMaster`/`SupplierMaster`/
  `LogisticsMaster`'s post-insert focus).
- **On Update** (`:182-198`): same pattern, no `SelectAll`.
- **On Delete** (`:200-215`): NO explicit field-clear of `PartsNum_Edit` here (unlike `SizeMaster`/
  `SupplierMaster`/`LogisticsMaster`, which blank their key field before requery) — just
  `GetInventoryInfo` (requery) then `SetDetailBoxes` directly (no `SearchGrid` re-search step
  either). `SetDetailBoxes` at this point reads `Data_Module`'s in-memory values, which still hold
  the JUST-DELETED row's data (`HoldDetails(False)` ran at `:203` before the confirm dialog, holding
  the current edit-box contents, not clearing them) — **the panel and `PartsNum_Edit` are left
  showing the deleted part's values verbatim after a successful delete**, a stronger/more visible
  version of the stale-panel hazard seen elsewhere (here even the code field itself is stale, not
  just the secondary fields).
- **`TextChange`** (`:641-647`): any `TMaskEdit` whose trimmed text becomes empty is forced to `'0'`
  — same "never blank, defaults to zero" behavior as `SizeMaster`, applied here to all the lead-
  time/ship-days mask edits (wired via `.dfm` `OnChange = TextChange` on 13 controls).
- **`MaskEditExit`** (`:649-666`): strips ALL internal spaces from any `TMaskEdit`'s text on exit
  (not just trim/pad) — reconstructs the string character-by-character removing every `' '` — wired
  via `.dfm` `OnExit = MaskEditExit` on the same mask-edit set (`OneLotQty`, `Quantity`, `LeadTime`,
  `RenbanCount`, and all 12 weekday lead-time/ship-days edits).

## Focus & keyboard
- `FormShow` (`:690-694`): `SetDetailBoxes` then `PartsNum_Edit.SetFocus` — same repopulate-from-
  shared-`Data_Module`-state-then-focus pattern as `SizeMaster`/`SupplierMaster` (same #135-class
  hazard: on first open with a non-empty grid, or with leftover state from a prior session, the
  panel is NOT guaranteed blank).
- No VCL `Default`/`Cancel` flags; `Close_Button` has `ModalResult = 2` only
  (`PartsStockMaster.dfm:84`). Accelerators only (`&Insert`, `&Update`, `&Search`, `Cl&ear`,
  `&Close`, `&Delete`).
- Grid `OnKeyUp`/`OnMouseUp` (`:668-681`) both call `HoldDetails(True)` + `SetDetailBoxes` — standard
  click-or-keystroke re-sync; `Parts_DataSourceDataChange` (`:683-688`) mirrors on dataset
  navigation.
- `PartsNum_Edit` has `MaxLength = 12`, `CharCase = ecUpperCase`, and a literal `.dfm` design-time
  default `Text = '000000000000'` (`PartsStockMaster.dfm:390`) — **the field is NOT blank at design
  time**; `FormCreate` explicitly overrides this to `''` (`:278`) so the design-time default never
  actually reaches the operator in practice, but it's the only field in the family with a non-blank
  literal default baked into the `.dfm` (worth flagging in case any other code path constructs this
  form without going through the normal `FormCreate`/`Execute` flow).
- Tab order runs 0 (`PartsNum_Edit`) → 1 (`KanbanNum_Edit`) → 2 (`PartsName_Edit`) → 3 (Supplier
  combo) → 4 (Logistics) → ... → 29 (`Remarks_Edit`) per the `.dfm` `TabOrder` values
  (`PartsStockMaster.dfm:389,426,435,...,417`); `VendorShare_Edit`/`VendorShare_SpeedButton`/
  `Label21` are all `Visible = False` design-time (`:359-360,380,778`) — dead/hidden UI still present
  in the tab chain (`TabOrder = 13` on the hidden `VendorShare_Edit`, `.dfm:777`) — a rebuild
  shouldn't reproduce a tab-stop on an invisible control.

## Enable/disable state machine
- **The one real enable/disable transition in this family**: `LotSizeOrders_CheckBoxClick`
  (`:696-709`) toggles `RenbanCode_ComboBox.Enabled` based on the (inverted) checkbox state — see
  Field clear/repopulate above. This is the ONLY control-enable state machine found across all
  eight master forms in this sweep.
- All buttons (Insert/Update/Delete/Search/Clear) are otherwise always enabled regardless of grid
  state.

## Error surfacing
- `SearchGrid` wraps in `try/except` → `ShowMessage`. `Insert_/Update_/Delete_ButtonClick` have no
  try/except of their own — a DB-layer exception from `InsertPartsStockInfo`/`UpdatePartsStockInfo`/
  `DeletePartsStockInfo` is unguarded, same gap as every sibling form.

## Cross-refs
- `docs/analysis/master-data/parts-stock.md` (proc/data spec) — **not found in
  `docs/analysis/master-data/` at time of writing** (only `logistics.md`, `manifest-cost.md`,
  `master-maint.md`, `size.md`, `supplier.md` exist there); if this file doesn't yet exist, this is
  the placeholder cross-ref pending that spec being written.
