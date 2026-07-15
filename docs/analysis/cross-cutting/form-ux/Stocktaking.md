# Form-UX: `TStocktaking_Form` — `Stocktaking.pas` + `Stocktaking.dfm`

> "Stocktaking Adjustment" — signed-delta inventory adjustment ledger. Confidence: high — both
> files read in full (362 + 259 lines). Nearly structurally identical to `RecReject.pas` (same
> author/era; grid + detail-panel + supplier/part cascading combo pattern).
> Business/proc spec: [`../../inventory-stock/stocktaking.md`](../../inventory-stock/stocktaking.md).

## Dialogs & confirmations
- **`Delete_ButtonClick`** (`:284-297`) — **armed confirm:**
  `MessageDlg('Are you sure you wish to delete'+#13+'part code '+PartNum+#13+' from the database?',
  mtWarning, [mbYes, mbNo], 0) = mrYes` → only then `DeleteStocktakingInfo`. Same `mtWarning` icon
  convention as `RecConfStat`/`RecReject`'s delete confirms.
- **`Insert_ButtonClick`** (`:253-264`) — no confirmation; on `InsertStocktakingInfo` returning
  `False`:
  `MessageDlg('Unable to INSERT '+PartNum+'('+SupplierCode+')', mtInformation, [mbOk], 0)` (`:259`)
  — a single-button `mbOk` info dialog (not Yes/No), the only failure-dialog shape in this form.
- **`Update_ButtonClick`** (`:266-282`) — **no confirmation and no failure-path check at all**: the
  return value of `UpdateStocktakingInfo` is discarded (`:275`) — identical silent-failure gap to
  `RecReject.pas`'s Update. Per the business spec, this same proc also has the confirmed
  `VC_LAST_UPDATE = NULL` timestamp bug (D8/Bug 2) — so a failed *or* successful-but-buggy update
  both surface identically to the operator: nothing.
- **`HoldDetails(False)`'s blank/invalid-part-code guard is also silently non-blocking**
  (`:167-169`): `if Length(trim(PartsCode_ComboBox.Text)) <> 12 then fErrMsg := #13+'Invalid Part
  Code'` — but **`Insert_ButtonClick` calls `HoldDetails(False)` and discards the return value**
  (`:255`), so this message is built and never shown; the operator only learns of a bad part code
  via the DB `IN_PART_ID NOT NULL` failure surfacing as the generic `'Unable to INSERT...'` dialog
  above (which doesn't mention "invalid part code" at all — a misleading disconnect between the
  built-but-unused message and the actual dialog shown).
- **No confirmation on `Search_ButtonClick`/`Clear_ButtonClick`** — pure filter actions.
- **Top-level `Execute`'s catch** (`:84-89`): `ShowMessage('Unable to generate Stocktaking
  screen.'+#13+'ERROR:'+#13+E.Message)`.

## Field clear / repopulate
- **`SetDetailBoxes`** (`:182-203`) — **no `try…except` wrapper at all** (contrast `RecReject.pas`'s
  equivalent, which swallows errors silently; here an exception would propagate uncaught to the
  caller). If `EditDate <> ''`, reformats into the date picker; **else `Edit_DateTimePicker.Date :=
  now`** (`:195`) — **this form defaults the picker to "now" on an empty incoming date, unlike
  `RecConfStat`'s milestone fields (which go to a true-empty `.Text:=''`) and unlike `RecReject`'s
  (which leaves the picker's prior value untouched on a bad/empty string).** Three different
  empty-date behaviors across three sibling forms in this family — **a concrete #135-class
  discrepancy to reconcile explicitly in the rebuild's date-input component**, since "empty → now"
  vs. "empty → blank" vs. "empty → keep prior" are materially different UX outcomes from the same
  underlying empty string.
  `SearchMultiCombo`/`SearchCombo` reposition the supplier/part combos (shared placeholder
  convention, see Cross-refs); `Qty_MaskEdit.Text:=IntToStr(Quantity)`;
  `Reason_Memo.Text:=Comments`.
- **`Clear_ButtonClick`** (`:307-316`): `SetTodaysDate(Stocktaking_Panel)` (resets the date picker to
  today) + `ClearControls(Stocktaking_Panel)` (shared DataModule walker: blanks
  `TEdit`/`TMaskEdit`/`TMemo`, `TComboBox.ItemIndex:=0`) — **no explicit `-1` override for the
  Supplier/Parts combos** (contrast `RecConfStat.Clear_ButtonClick`, which explicitly forces `-1`
  after `ClearControls`'s `0`) — so on this form, Clear leaves the Supplier/Parts combos at
  **index 0 (the blank-placeholder first item), not `-1`**. Functionally similar in appearance
  (both show blank-looking text) but a different underlying `ItemIndex` value — worth noting if the
  rebuild's "no selection" state is keyed off `ItemIndex = -1` specifically.
- **`FormCreate`** (`:102-117`): `fNochange:=True` guards initial combo population from re-firing
  `Stocktaking_DataSourceDataChange`; `SetTodaysDate` + `ClearControls` on the panel before the grid
  is wired — same pattern as `RecReject.pas`.
- **`HoldDetails(True)` (from-grid capture, `:144-157`)** — straightforward field-index copy, no
  derived-label parsing (contrast `RecReject`'s division-label lookup) since Stocktaking has no
  classifier column.

## Focus & keyboard
- **Initial focus (`FormShow`, `:342-347`)** — **unlike every other form in this family, `FormShow`
  here calls `HoldDetails(True)` THEN `SetDetailBoxes` before `SetFocus`** (`:344-345`): it
  re-captures whatever the DBGrid's *currently selected* row already is (not a fresh blank state)
  into the detail panel, then focuses `Supplier_NUMMIColumnComboBox` (`:346`). Since
  `Stocktaking_DataSource.DataSet` was wired in `FormCreate` and the grid auto-selects its first row
  on data-bind, **the form opens with the first ledger row's values already loaded into the detail
  panel** — a materially different initial state than `RecReject`'s "blank panel on open" (per
  `RecReject.pas:154-169`, `ClearControls` runs and nothing re-populates it before `FormShow`).
  **Flag this cross-form inconsistency explicitly** — two structurally near-identical forms differ
  on whether opening the screen starts blank or pre-loaded with the first row.
- **Post-Insert/-Update/-Delete/-Search focus:** `Supplier_NUMMIColumnComboBox.SetFocus` in all four
  (`:263`, `:281`, `:296`, `:304`) — identical pattern to `RecReject.pas`.
- **`Supplier_NUMMIColumnComboBoxChange`** (`:349-360`) — repopulates the dependent Parts combo and
  **`SetFocus`es `PartsCode_ComboBox`** (`:358`) — cascading auto-advance, identical to `RecReject`.
- **`PartsCode_ComboBoxChange`** (`:248-251`) — **`SetFocus`es `Qty_MaskEdit`** (`:250`) — continues
  the chain: Supplier → Part → Qty, same keyboard/scanner-friendly shape as `RecReject`.
- **No `Default`/`Cancel` flags, no `KeyPreview`** anywhere in `Stocktaking.dfm` (grep-verified) —
  Enter does not trigger Insert/Update/Delete/Search.
- **`Qty_MaskEdit` carries `CharCase = ecUpperCase`** (`Stocktaking.dfm:197`) — immaterial for a
  numeric mask (digits/minus have no case) but is the **only** `CharCase` setting on the form; a
  stray default, not meaningful behavior (per the business spec's own note on this).

## Enable/disable state machine
- **No button ever disables based on state** — identical to `RecReject.pas`: Insert/Update/Search/
  Clear/Delete are always clickable; there is no visible distinction between "editing an existing
  row" and "entering a new one" beyond which fields `HoldDetails` last captured.
- **`Stocktaking_DBGrid`'s `dgConfirmDelete`** (`Stocktaking.dfm:41`) — same inert-unless-proven
  caveat as the other grids in this family (the explicit `Delete_Button` + `MessageDlg` is the real
  delete path; the grid's own Delete-key confirm is unverified/likely unreachable given the
  stored-proc-backed, non-keyset dataset).

## Error surfacing
- **Update's silent failure is the standout gap** (identical to `RecReject.pas`): no return-value
  check, no dialog, no log call on failure — combined with the business spec's confirmed
  `VC_LAST_UPDATE = NULL` bug on this exact proc, an operator has **zero** signal that an update
  just corrupted the timestamp.
- **`SetDetailBoxes` has NO exception handler** (unlike `RecReject`'s silent-swallow) — an error
  here would propagate as an unhandled-exception dialog from the VCL runtime itself, not a
  controlled app message; different failure mode from its sibling form despite near-identical code
  shape.
- Everything else surfaces via plain `ShowMessage`/`MessageDlg`; no inline field-error text.

## Cross-refs
- Business rules / procs / triggers (signed-delta semantics D5, unpersisted date picker on write,
  `UPDATE_StockTakingInfo` NULL-timestamp bug D8/Bug 2, second writer `InsertAutoScrap` from
  `DailyBuildTotal`): [`../../inventory-stock/stocktaking.md`](../../inventory-stock/stocktaking.md).
- Sibling form (near-identical structure, several UX divergences noted above):
  [`RecReject.md`](RecReject.md).
- Shared DataModule field-clear/combo-populate ancestor: `DataModule.pas:5976` `ClearControls`,
  `:5916` `SearchMultiCombo`, `:5935` `SearchCombo`, `:5964` `SetTodaysDate` — see `RecConfStat.md`
  Cross-refs for the blank-placeholder (`' '` literal, `ItemIndex:=0`) convention shared by every
  combo in this family.
