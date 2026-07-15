# Form-UX semantics: `TManifestCostMaster_Form` — `ManifestCostMaster.pas` / `ManifestCostMaster.dfm`

CRUD editor over the assembly-manifest-cost table (Assy code + Assy Manifest No + cost-start/end
date range + cost), keyed on `(Assy; Start Manifest; End Manifest)`. Reached from `MasterMaint`'s
`MonthlyPO_Button` ONLY when `Data_Module.fiGenerateEDI.AsBoolean` is true (button caption becomes
"Manifest Cost" in that mode — see `MasterMaint.md`). Grid `MonthlyPOMaster_DBGrid` bound via
`MAnifestCost_DataSource` (source-name typo carried verbatim).

## Dialogs & confirmations
- **Insert failure** — `ManifestCostMaster.pas:304`: `MessageDlg('Unable to INSERT ' +
  Data_Module.AssyCode, mtInformation, [mbOk], 0)` if `InsertManifestCostInfo` fails.
- **Delete confirmation (armed, two-step)** — `:262-266`: `MessageDlg('Are you sure you wish to
  delete Assy Number ' + Data_Module.AssyCode + #13 + 'From ' + Data_Module.POStart + #13 + 'To ' +
  Data_Module.POEnd + #13 + ' from the database?', mtWarning, [mbYes, mbNo], 0) = mrYes` — the
  richest confirmation text in the family (echoes the full composite key: assy + date range).
- **Empty-search guard** — `:242-243`: if `Trim(AssyCode_ComboBox.Text) = ''`, `ShowMessage('Please
  enter a assembly code before searching.')` and the search is skipped entirely (does not even call
  `SearchGrid`) — a pre-condition dialog not present on the simpler code-based masters (which just
  search on whatever's in the box, blank or not).
- **Search "not found"** — `:253`: `ShowMessage('No matches were found for your query.')`.
- **Search error** — `:148-150`: `ShowMessage('Error in Search' + #13 + e.Message)`.
- **Combo-population failure** — `GetParts` (`:153-173`): `MessageDlg('Unable to get a list of
  parts.', mtError, [mbOK], 0)` if `SelectSingleField` raises. Used both for `AssyCode_ComboBox`
  (called from `FormShow`, `:224`) — note this is the ONLY master form in the family that populates
  one of its key combos from `FormShow` rather than `FormCreate`.
- **Inline invalid-cost message (not a dialog, a `ShowMessage`)** — `HoldDetails(False)` (`:118-121`):
  if the currency text (after stripping the `$` prefix) doesn't parse via `TryStrToFloat`,
  `ShowMessage('Invalid assembly cost')` + `AssyCost_MaskEdit.SetFocus` — but **execution continues
  past this point without setting a failure flag**; `HoldDetails` has no return value consumed by
  the caller for this condition (its siblings' `Validate` functions return `False` and abort; this
  one just complains and moves on) — Insert/Update proceed regardless of whether `AssyCost` got a
  valid new value (it silently keeps its PREVIOUS `Data_Module.AssyCost` value in that branch, since
  the `Else AssyCost := fTempDouble` branch is skipped). **Hazard: a rejected cost entry does not
  block the Insert/Update it was meant to gate.**
- Grid carries `dgConfirmDelete` (`ManifestCostMaster.dfm:39`) — VCL-native secondary confirm path.

## Field clear / repopulate
- **`FormCreate`** (`:195-219`): populates `AssyManifestID_ComboBox` programmatically with `' '`
  (blank sentinel) + `'01'..'99'` (zero-padded) — the ONLY combo in the family built by a hardcoded
  loop rather than a DB lookup. Then unfilters dataset, `GetManifestCostInfo`, binds data source,
  `ClearControls(ManifestCost_Panel)`, re-unfilters. **Does NOT clear `AssyCode_ComboBox`** here
  (that combo isn't populated until `FormShow`).
- **`FormShow`** (`:221-226`): `AssyCode_ComboBox.SetFocus`, THEN `GetParts(...)` populates
  `AssyCode_ComboBox` from `INV_FORECAST_DETAIL_INF`, THEN `Data_Module.Inv_DataSet.First` — no
  `SetDetailBoxes` call here at all (unlike every sibling form's `FormShow`). The panel's actual
  population on open therefore depends entirely on `Inv_DataSet.First` firing
  `MAnifestCost_DataSourceDataChange` (`:319-324`), which DOES call `HoldDetails(True)` +
  `SetDetailBoxes` — **this form's initial-populate is 100% DB/dataset-driven (via the
  `TDataSource.OnDataChange` event), not form-code-driven** — the only member of this family where
  that's true for `FormShow` specifically. If the underlying dataset is empty, `.First` is a no-op on
  an empty set and `OnDataChange` may not fire at all → the panel would show whatever was left over
  from `FormCreate`'s `ClearControls` (i.e., correctly blank in that specific empty-table case,
  unlike the stale-state hazards on the code-keyed masters).
- **`SetDetailBoxes`** (`:62-84`) wraps its entire body in a `try/except` with an **EMPTY except
  block** (`:78-82`, `on e:exception do begin end`) — any error while formatting dates/currency into
  the edits is **completely silent**, no ShowMessage, no log. This is the family's only fully-silent
  error path in a detail-repopulate routine. Also note the date parsing:
  `StrTodate(copy(POStart,5,2)+'/'+copy(POStart,7,2)+'/'+copy(POStart,1,4))` — expects `POStart`/
  `POEnd` as `YYYYMMDD` 8-char strings; a malformed/short string here is exactly the kind of input
  the empty except-block would swallow.
- **On Insert/Update** (`:277-317`): both capture `Assy`/`SMan`(start)/`EMan`(end) from
  `Data_Module` right after `HoldDetails(False)`, requery via `GetManifestCostInfo`, re-unfilter,
  then `Locate('Assy;Start Manifest;End Manifest', VarArrayOf([Assy,SMan,EMan]), [])` — a genuine
  **composite-key locate**, unlike the single-field `Locate`s on the other masters. `SetDetailBoxes`
  runs after, echoing the just-written row.
- **On Delete** (`:259-275`): `GetManifestCostInfo` (requery), `ClearControls(ManifestCost_Panel)`
  (full blank — this form DOES clear on delete, unlike `SizeMaster`/`SupplierMaster`), re-unfilter,
  `Inv_DataSet.First` (which will fire `OnDataChange` → repopulate from whatever the new first row
  is, if any exist — so the "blank" from `ClearControls` may be immediately overwritten by the first
  remaining row, not left blank, if the table is non-empty post-delete).
- **`Clear_Button`** (`:228-236`): unfilters, `ClearControls(...)`, focus to `AssyCode_ComboBox` —
  standard full blank.

## Focus & keyboard
- `FormShow` sets focus to `AssyCode_ComboBox` FIRST (`:223`), before the combo is even populated
  (`GetParts` runs after `SetFocus` in source order, `:223-224`) — cosmetically harmless (focus
  persists once the list is populated) but notable ordering inversion vs. every sibling form (which
  populate first, then focus).
- Every button-click handler that finishes also explicitly `SetFocus`s `AssyCode_ComboBox`
  (`Clear_ButtonClick :235`, `Search_ButtonClick :256`, `Delete_ButtonClick :274`,
  `Update_ButtonClick :293`, `Insert_ButtonClick :316`) — a **uniform final-focus target**
  regardless of action, unlike the code-edit-focused pattern on `SizeMaster`/`SupplierMaster`/
  `LogisticsMaster` (which focus the code/name field they just wrote).
- No VCL `Default`/`Cancel` flags; `Close_Button` has `ModalResult = 2` only
  (`ManifestCostMaster.dfm:97`). Accelerators only.
- `MAnifestCost_DataSourceDataChange` (`:319-324`) is the ONLY dataset-change hook on this form —
  there is **no grid `OnKeyUp`/`OnMouseUp` handler** wired at all (`ManifestCostMaster.dfm`'s
  `MonthlyPOMaster_DBGrid` object has no `OnKeyUp =` / `OnMouseUp =` lines, `:33-46`), unlike every
  other master form in this family. Grid row selection still reaches the panel because DBGrid
  navigation moves the underlying `TDataSet`'s current record, which fires `TDataSource.OnDataChange`
  regardless of how the move happened (click, arrow keys, or code) — so behavior is equivalent in
  practice, just implemented one layer lower (DB-dataset-driven, not form-code-driven, for BOTH
  grid interaction and initial `FormShow` here).

## Enable/disable state machine
- **None found** — no `Enabled` toggling; all buttons always clickable.

## Error surfacing
- `SearchGrid` (`try/except` → `ShowMessage`), `GetParts` (`try/except` → `MessageDlg`), and
  `SetDetailBoxes` (`try/except` → **silent, no message**) are the three guarded blocks.
  `Insert_/Update_/Delete_ButtonClick` have no try/except of their own.

## Cross-refs
- `docs/analysis/master-data/manifest-cost.md` (proc/data spec).
