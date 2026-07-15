# Form-UX: `TModifyShipping_Form` — `ModifyShipping.pas` + `ModifyShipping.dfm`

> "Amend an already-posted shipment's part lines" sub-form, opened only from `Shipping.pas`'s
> `UpdateShipping_Button`. Confidence: high — both `.pas`/`.dfm` read in full.
> Business/proc spec: [`../../shipping/shipping.md`](../../shipping/shipping.md).

## Dialogs & confirmations
- **Zero `MessageDlg`/`ShowMessage` calls of any kind in this unit** (grep-verified — the string
  `'ShowMessage'`/`'MessageDlg'` does not appear in `ModifyShipping.pas`). All validation failures
  are `SetFocus`-only (see Error surfacing). This is the **quietest** form in the family — no
  success confirmation after Insert/Update either.
- **`Execute`'s exception handler is empty** (`:166-170`: `on e:exception do begin end;`) — any
  error during form setup (e.g. `GetShippingInfo` failure) is **completely swallowed, no dialog, no
  log call**. A genuine silent-failure hazard: the operator sees a form that just didn't populate,
  with no indication anything went wrong.
- **No delete confirmation because there is no wired Delete** — `Delete_Button` exists on the
  `.dfm` (`ModifyShipping.dfm:89-97`) with `Visible = False` and **no `OnClick` handler at all**
  (confirmed: no `Delete_ButtonClick` procedure exists in `ModifyShipping.pas`). Matches
  `shipping.md`'s finding that this form has no line-delete/restore path.
- **`ModifyShipping_DBGrid` carries `dgConfirmDelete`** (`ModifyShipping.dfm:192`) — the VCL
  built-in "Delete this record?" grid confirm exists at the component-option level, but since the
  underlying `Inv_DataSet` (a stored-proc SELECT result) is not app-wired for in-grid delete and no
  Delete button drives it, this is very likely **inert** in practice — flag as unverified, not a
  real operator-facing confirm; do not carry the built-in-grid-delete assumption into the rebuild.
- **`Search_Button`** is also `Visible = False` with no `OnClick` (`ModifyShipping.dfm:60-67`) —
  dead control, same pattern as Delete.

## Field clear / repopulate
- **`SetDetailBoxes`** (`:124-131`) is trivial: `SearchCombo(PartNumber_ComboBox, PartNum)` (the
  shared DataModule combo-search helper — see `RecConfStat.md` Cross-refs for its "blank → literal
  single-space list item, ItemIndex 0" convention) + `Qty_Edit.Text:=IntToStr(Quantity)`.
- **`Clear_ButtonClick`** (`:219-227`) — the "new line" reset: `PartNumber_ComboBox.ItemIndex:=0`
  (the first combo item, **not** necessarily blank — depends on what `SelectSingleField` populated
  as item 0; per the DataModule convention this is the literal `' '` placeholder, so ItemIndex 0
  **is** effectively "blank," consistent) and `Qty_Edit.Text:=''` — **note the asymmetry: the combo
  resets to its blank placeholder, but `Qty_Edit` resets to a truly empty string, not `'0'`** (no
  `TextChange`/`OnChange` handler forces it back to a numeric default here, unlike the mask-edit
  forms). A rebuild that assumes "all numeric fields default to 0" would diverge from this form.
- **`ModifyShipping_DataSourceDataChange`** (`:201-212`, fires on grid row selection) — when
  `not InUpdate`: calls `HoldDetails(True)` (copies grid row into `PartNum`/`Quantity`) then
  `SetDetailBoxes`, and **flips to "edit an existing line" mode**: `Update_Button.Visible:=TRUE`,
  `Insert_button.Visible:=FALSE`, `Clear_button.Visible:=TRUE`. Selecting a grid row always
  re-populates the qty/part fields from that row — no stale carry-over between rows.
- **`InUpdate` re-entrancy flag:** set `TRUE` at the top of `Update_ButtonClick`/`Insert_ButtonClick`
  before calling `GetShippingInfoDetail` (which re-opens the grid dataset and would otherwise
  re-trigger `ModifyShipping_DataSourceDataChange` recursively) — set back `FALSE` immediately after
  the refetch, **before** `SetDetailBoxes` is called at the end of each handler. This guards against
  a field-clear feedback loop, not a business rule.
- **`HoldDetails(False)`** (validate+capture from the manual-entry controls, `:79-122`) — on a
  blank/space `PartNumber_ComboBox.Text` (`' '` or `''`), shows nothing (see Error surfacing) and
  returns `False` **without clearing any field** — the operator's bad entry stays on screen for
  correction, contrast `ManualShipping`'s blank-on-error behavior.

## Focus & keyboard
- **Initial focus (`FormShow`, `:214-217`):** always `Qty_Edit.SetFocus`.
- **Post-Update/-Insert focus:** `Qty_Edit.SetFocus` (`:197`, `:244`) in both handlers — the operator
  is dropped back into the qty field ready for the next entry (grid-and-edit-panel rhythm).
- **`Update_ButtonClick`** (`:178-199`) re-locates the grid to the row just edited:
  `Data_Module.Inv_DataSet.Locate('IN_PART_SHIPPING_ID', ID, [])` (`:194`) after refetching — so the
  just-edited row stays visually selected post-save (a UX nicety the rebuild should preserve).
- **No `Default`/`Cancel` button flags** anywhere in `ModifyShipping.dfm` (grep-verified) and **no
  `KeyPreview`** — Enter does not trigger Insert/Update/Delete from this form; only explicit click.

## Enable/disable state machine
- **Three-state button visibility, driven entirely by grid-selection vs Clear:**
  1. **Fresh/no selection (initial, and post-`Clear_ButtonClick`):** `Insert_Button.Visible:=TRUE`,
     `Update_Button.Visible:=FALSE`, `Clear_Button.Visible:=FALSE` (`:221-223`).
  2. **A grid row is selected** (`ModifyShipping_DataSourceDataChange`, `:208-210`):
     `Update_Button.Visible:=TRUE`, `Insert_button.Visible:=FALSE`, `Clear_button.Visible:=TRUE`.
  3. There is **no third "new line, unsaved" indicator** distinct from state 1 — `Insert_Button`
     covers both "form just opened" and "operator clicked Clear to start a new line."
- **`GroupBox1` (production-date/line/seq display)** is `Enabled = False` in the `.dfm`
  (`ModifyShipping.dfm:105`) — permanently a **read-only display block**, never toggled by code;
  the line/date/sequence values shown here are informational only, not editable from this sub-form.
- **`Search_Button` / `Delete_Button`** are permanently `Visible = False` with no runtime code path
  to show them (dead controls, see Dialogs section).

## Error surfacing
- **Validation failures are `SetFocus`-only, no dialog, no message, no visible error text**:
  - Blank/space part number (`HoldDetails:103-109`): `PartNumber_ComboBox.SetFocus`, `result:=FALSE`
    — the operator gets no textual cue at all why nothing happened.
  - Non-numeric qty (`:112-118`): `Qty_Edit.SetFocus`, `result:=FALSE` — same silent block.
- **Form-setup exceptions are fully swallowed** (`Execute`'s empty `except` block, `:166-170`) — the
  single worst error-surfacing gap in this family; the rebuild must NOT reproduce this (log at
  minimum, surface a toast/banner).

## Cross-refs
- Business rules / procs / triggers: [`../../shipping/shipping.md`](../../shipping/shipping.md)
  (§5 "no delete button is wired," the only stock-restore path being the un-triggered
  `DeleteShipDate` header cascade).
- Shared combo-search helper conventions: see `RecConfStat.md` Cross-refs
  (`DataModule.pas:5935` `SearchCombo`).
