# Form-UX: `TASNInvoice_Form` — `ASNInvoice.pas` + `ASNInvoice.dfm`

> The ASN/Invoice browser-editor hub: list ASNs or Invoices by status, drill into a header's line
> items, edit/insert/delete a line, delete/unsend a whole ASN/Invoice, recreate the 856/810 file.
> Confidence: high — both files read in full (989 + 468 lines).
> Business/proc spec: [`../../edi/asn-invoice.md`](../../edi/asn-invoice.md).

## Dialogs & confirmations
- **`Delete_ButtonClick`** (line-item delete, `:613-642`) — **armed confirm:**
  `MessageDlg('Delete ASN item with Manifest Number('+ManifestNumber_Edit.Text+')', mtConfirmation,
  [mbYes, mbNo], 0) = mrYes`. Only on `mrYes` does it call `dbo.DELETE_ASNItem`. Default button of
  a `[mbYes,mbNo],0` dialog is **Yes** (index 0).
- **`DeleteASN_ButtonClick`** (whole-ASN delete, `:697-720`) — **armed confirm:**
  `MessageDlg('Delete complete ASN ('+ASNList_DataSet.fieldByName('Production Date').AsString+')',
  mtConfirmation, [mbYes, mbNo], 0) = mrYes` → `dbo.DELETE_ASNList`. Same default-Yes shape.
- **`UnsendASN_ButtonClick`** (`:722-778`) — **armed confirm, two message variants depending on
  mode:**
  - ASN mode: `'Unsend ASN ('+ProductionDate+')'` (`:729`) → `dbo.UPDATE_ASNUnsend`.
  - Invoice mode: `'Unsend INV ('+EIN Number+')'` (`:756`) → `dbo.UPDATE_INVUnsend`.
  Both `mtConfirmation, [mbYes,mbNo], 0`, both gated additionally by `RecordCount > 0` (silently
  no-ops if the list is empty — no dialog at all in that case).
- **No confirmation on Insert or Update of a line item** (`Insert_ButtonClick:555-611`,
  `Update_ButtonClick:644-695`) — both post directly on click. Validation-only dialogs:
  `'Must have numberic value for Qty'` (`:562`), `'Quantity must be a numeric'` (`:652`),
  `'Quantity must be greater than 0'` (`:660` — **⚠️ the check is `i<0`, so `i=0` PASSES**; the
  message text says "greater than 0" but the code allows exactly 0 through).
- **No confirmation on `RecreateFile_ButtonClick`** (`:792-905`) — regenerating and overwriting the
  856/810 output file on disk is a single click, no "this will overwrite/resend" warning. Outcome
  dialogs: `'Recreate 856 file, complete'` (`:850`), `'Recreate 810 file, complete'` (`:889`),
  `'Unable to Recreate EDI856 for (...)'` (`:841`), `'Unable to recreate 856 file'` (`:854`),
  `'Unable to recreate 810 file'` (`:893`). **This is the same "RecreateFile" button referenced in
  the R20 multi-path-oracle lesson** (see Cross-refs) — confirm any test/oracle anchors to the
  *operational* 856/810 sender, not this manual recreate path.
- **Generic exception dialogs** throughout (`ShowMessage('Exception: unable to …, '+e.Message)`)
  in `GetASNs`/`GetINVOICEs` (`:205-212`, `:331-336` — the latter also `Close`s the whole form on
  error), `ASNListDataSourceDataChange` (`:399-404`, also closes the form), `Invoice_
  DataSourceDataChange` (`:457-463`, also closes), `Insert_ButtonClick`'s catch (`:594-609`, does
  **not** close — instead re-syncs the button-visibility state and re-populates the edit fields
  from the dataset, an unusually thorough error-recovery path for this codebase).
- **`SpeedButton1Click`** (search, `:907-987`) — validation-only: `'Incorrect search data, can
  search for Production Date or Manifest Number'` (`:983`) if the input doesn't start with `'2'` or
  `'7'`; `'Production Date not found'`, `'Manifest date not found'`, `'Manifest number not found'`
  are plain not-found info dialogs, not confirmations.

## Field clear / repopulate
- **Status-combo-driven list reload** (`GetASNs`/`GetINVOICEs`, `:89-338`) — switching
  `ASNStatus_ComboBox` (ALL/NOT CREATED/SENT/ACCEPTED/REJECTED) re-queries `SELECT_ASNList` /
  `SELECT_INVOICEList` with a `@List` code and re-derives the **entire button-visibility matrix**
  per status (see Enable/disable). **Item 1 ("NOT CREATED") self-corrects**:
  `GetINVOICEs`'s case `1` (`:115-135`) forcibly resets `ASNStatus_ComboBox.ItemIndex:=0` mid-handler
  — selecting "NOT CREATED" in Invoice mode silently snaps back to "ALL" (a real UX surprise: the
  combo visually shows "ALL" after the operator picked "NOT CREATED", with no message explaining
  why). `GetASNs`'s parallel case `1` does **not** self-correct (ASN mode's "NOT CREATED" *does*
  load the `'C'` list normally) — **an asymmetry between ASN and Invoice modes for the same combo
  index.**
- **`ASNorInvoice_ComboBoxChange`** (`:419-445`) — switching ASN↔Invoice mode fully rebuilds
  `ASNStatus_ComboBox.Items` (Clear + re-Add the same 5 strings) and resets `ItemIndex:=0`, discarding
  whatever status filter was previously selected — no attempt to preserve the operator's filter
  choice across the mode switch.
- **`Clear_ButtonClick`** (`:488-540`, the "new ASN item" reset) — populates
  `AssemblyPartNumber_Combo` from `AssyManifest_DataSet` **minus** parts already itemized on this
  ASN (an explicit `IndexOf`+`Delete` loop, `:513-522` — the combo only offers not-yet-added parts).
  Then `AssemblyPartNumber_Combo.ItemIndex:=0` and fires `AssemblyPartNumber_ComboChange` to
  regenerate the manifest number. **`ManifestNumber_Edit.Text:=''`** momentarily then immediately
  overwritten by the ComboChange call; **`Qty_Edit.Text:='0'`** (not blank).
- **`ASNItemsDataSourceDataChange`** (`:466-486`) — only repopulates
  `ManifestNumber_Edit`/`AssemblyPartNumber_Edit`/`Qty_Edit` from the grid row **when
  `ASNStatus_ComboBox.ItemIndex` is 1 or 4** (NOT CREATED / REJECTED) **and** `not Insert_Button.
  Visible` — i.e. only while in "editing an existing line" mode for those two statuses; for SENT/
  ACCEPTED/ALL the item-detail fields are simply **not refreshed on grid-row selection** (they stay
  whatever they last held) — **potential stale-value display (#135-adjacent)** if an operator
  switches grid rows under a status where this guard doesn't fire; verify the fields are hidden or
  irrelevant in those statuses before assuming this is safe (per `ASNItem_Box.Visible` toggling in
  `GetASNs`/`GetINVOICEs`, the box **is** hidden for statuses 0/2/3, so this is likely benign, but
  status-1/4-only-refresh plus a manual grid-DataSource wire-up is a fragile combination to
  reproduce faithfully).
- **`Insert_ButtonClick`'s error path** (`:594-609`) explicitly **re-populates** the edit fields
  from `ASNItems_DataSet` (`ManifestNumber_Edit`, `AssemblyPartNumber_Edit`, `Qty_Edit`) after a
  failed insert — an intentional revert-to-last-known-good, not a stale leftover.
- **`AssemblyPartNumber_ComboChange`** (`:542-553`) — recomputes `ManifestNumber_Edit.Text` as a
  **derived, non-editable composite**: `'7' + <year-digit> + <MM> + <DD> + <ManifestID>` sliced from
  `ASNList_DataSet`'s Production Date + the located `AssyManifest_DataSet` row's Manifest ID
  (`:546-550`) — this field is **always computed, never typed by the operator** in the insert path.

## Focus & keyboard
- **No `FormShow`-driven initial-focus `SetFocus` call** — `FormShow` (`:356-364`) only dispatches
  to `GetASNs`/`GetINVOICEs`; whichever control is first in tab order gets default VCL focus.
- **`AssemblyPartNumber_ComboChange`** ends with `Qty_Edit.SetFocus` (`:551`) — after picking a part
  for a new line, focus lands on Qty (scanner/keyboard-friendly: pick part → type qty → click
  Insert).
- **`Clear_ButtonClick`** ends with `AssemblyPartNumber_Combo.SetFocus` (`:539`).
- **`Qty_Edit.SetFocus`** is also called on **every** numeric-validation failure in
  `Insert_ButtonClick`/`Update_ButtonClick` (`:563`, `:654`, `:662`) — consistent re-focus-on-error.
- **No `Default`/`Cancel` button flags, no `KeyPreview`** anywhere in `ASNInvoice.dfm`
  (grep-verified) — Enter does not submit Insert/Update/Delete from anywhere on this form; only
  explicit click (or the DBGrid's own default row-delete key, see below).
- **Both list/item `TDBGrid`s carry `dgConfirmDelete` + `dgCancelOnExit`**
  (`ASNInvoice.dfm:88`, `:127`) — the VCL built-in in-grid Delete-key confirm exists at the
  component-option level on `ListDBGrid`/`ItemsDBGrid`; whether it is reachable depends on whether
  the bound `TADODataSet`s are grid-editable (stored-proc result sets are typically not
  update-in-place without a keyset cursor) — **flag as unverified**, do not assume the rebuild needs
  to reproduce a working in-grid delete-key path distinct from the explicit `Delete_Button`.

## Enable/disable state machine
- **The single largest state machine in this family** — `GetASNs`/`GetINVOICEs` set **nine**
  button/control visibilities per status-combo selection (`Insert_Button`, `Update_Button`,
  `Delete_Button`, `DeleteASN_Button`, `RecreateFile_Button`, `UnsendASN_Button` [+ its caption text],
  `Clear_Button`, `Cancel_Button`, `AssemblyPartNumber_Combo` vs `_Edit`), duplicated near-verbatim
  across 5 case branches × 2 methods (`GetASNs`, `GetINVOICEs`) — **10 copies of substantially the
  same block**, a maintenance/parity hazard the rebuild should collapse into one status→visibility
  table rather than replicate as-is.
- **Status → controls matrix (ASN mode, `GetASNs`):**
  | `ASNStatus_ComboBox` index | List label | Item box | Insert | Update/Delete/DeleteASN | RecreateFile | UnsendASN |
  |---|---|---|---|---|---|---|
  | 0 ALL | — | hidden | hidden | hidden | hidden | hidden |
  | 1 NOT CREATED | — | **shown** | hidden | **shown** | hidden | hidden |
  | 2 SENT | — | hidden | hidden | hidden | hidden | **shown** ("Unsend ASN") |
  | 3 ACCEPTED | — | hidden | hidden | hidden | **shown** | **shown** ("Unsend ASN") |
  | 4 REJECTED | — | **shown** | hidden | **shown** | hidden | hidden |
  (Invoice mode mirrors this with `'X'`/`'S'`/`'A'`/`'R'` `@List` codes and analogous button sets —
  see `:94-213`; note case 0 and case 1 are **byte-identical** in `GetINVOICEs` except for the
  forced `ItemIndex:=0` snap-back described above.)
- **`ASNListDataSourceDataChange`** (grid row select, `:366-406`) additionally forces the item-edit
  panel visible **only when `ASNStatus_ComboBox.ItemIndex = 1`** (`:378-390`), and separately toggles
  `Clear_Button.Visible` based on whether `'Start Seq' = -1` (`:394-397`) — a **third** partial
  visibility path layered on top of the status-driven matrix above, evaluated only in ASN mode (no
  Invoice-mode equivalent exists for this specific nuance).
- **`Insert_ButtonClick` success** (`:586-593`) flips to "editing" mode (`Update`/`Delete`/
  `DeleteASN`/`Clear` visible, `Insert` hidden, combo→edit swap) — the mirror image of
  `Clear_ButtonClick`'s "new line" mode.

## Error surfacing
- **Overwhelmingly `ShowMessage`/`MessageDlg` dialogs**; no inline field-level error text or status
  bar anywhere in this unit.
- **Three handlers close the entire form on a caught exception** (not just show a dialog):
  `GetINVOICEs` (`:210`), `GetASNs` (`:335`), `ASNListDataSourceDataChange` (`:403`),
  `Invoice_DataSourceDataChange` (`:462`) — a transient DB error while browsing silently ejects the
  operator from the whole ASN/Invoice screen, losing their place (status filter, selected row) with
  no "would you like to retry" option. **Rebuild must not reproduce blind form-close-on-error.**
- **`LogActLog` calls accompany most successful writes** (Insert/Update/Delete/DeleteASN/UnsendASN/
  RecreateFile all log) but **errors are inconsistently logged** — some catch blocks call
  `Data_Module.LogActLog('ERROR', …)` (`GetASNs` does not; `DeleteASN_ButtonClick`,
  `UnsendASN_ButtonClick`, `RecreateFile_ButtonClick` do) alongside the dialog.

## Cross-refs
- Business rules / procs / X12 wire format:
  [`../../edi/asn-invoice.md`](../../edi/asn-invoice.md),
  [`../../edi/856/edi856-wire-format.md`](../../edi/856/edi856-wire-format.md),
  [`../../edi/810/edi810-wire-format.md`](../../edi/810/edi810-wire-format.md).
- **The R20 multi-path-oracle lesson names `RecreateFile_ButtonClick` explicitly**: the 856 filename
  can be produced by more than one legacy path (this manual recreate button vs. the operational
  ASN-creation sender in `ASNSelect.pas`) — when testing/oracle-anchoring the 856 filename, confirm
  which path the rebuild's behavior actually reproduces; don't anchor only to this button because
  it's the easiest to find.
- Sibling create-side form: [`ASNSelect.md`](ASNSelect.md).
