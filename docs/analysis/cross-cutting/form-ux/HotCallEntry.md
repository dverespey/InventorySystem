# Form-UX Semantics — `THotCallEntryForm` (HotCallEntry.pas / HotCallEntry.dfm)

Caption "One Cycle Entry" (`HotCallEntry.dfm:5`). Manual hot-call ASN entry — 12 part/qty pairs, one
Line, one manifest number, one date. LIVE — `InventorySystem.dpr:55`.

## Dialogs & confirmations

| Trigger | Text (verbatim) | Buttons | Consequence |
|---|---|---|---|
| Line blank on Add | `'Select a line'` — `HotCallEntry.pas:145` | OK | `Line_ComboBox.SetFocus`, `exit` (`:146-147`). |
| Manifest number < 8 chars | `'Manifest number must 8 characters'` **[verbatim — missing "be"]** — `:159` | OK | `ManifestNumber_Edit.SetFocus`, `exit` (`:160-161`). Note the commented-out numeric-manifest check at `:150-155` is **dead** — manifest is only length-checked, never validated as numeric, despite the field being called "Manifest **Number**". |
| A filled part-number row's qty is non-numeric | `'Qty must be numeric'` — `:186` | OK | The offending qty `TEdit.Text` is force-reset to `'0'` (`:187`) **then** `SetFocus` (`:188`), `exit` — so the operator sees a corrected `0` in the field, not their bad input. |
| A filled part-number row's qty is `<= 0` | `'Qty must be greater than 0'` — `:193` | OK | Same reset-to-`'0'` + `SetFocus` + `exit` pattern (`:194-196`). |
| A qty is entered but its paired part-number combo is blank | `'Part number required'` — `:212` | OK | `TComboBox` (the paired part-number field) `.SetFocus` (`:213`), `exit`. |
| Failure loading Line list on open | `'Failed on ASN select, '+<ADO error description>'` — `:106`, or `'Failed on ASN select, '+e.Message'` — `:122` | OK | `exit`/handled in `except`; form never opens the Line list correctly but `ShowModal` is only reached if `Open` succeeds without the `Errors.Count>0` short-circuit (`:104-109`) — i.e. this is a **soft-fail without a hard `exit` from `Execute`**, `[UNVERIFIED — confirm before use: whether `ShowModal` still runs after the `exit` at line 108 inside the `with...do` block — Delphi `exit` inside a `with` still exits the whole procedure, so `ClearEntries`/`ShowModal` at `:115-116` are skipped]`. |
| Successful Add commit | `'Added HotCall manifest(<manifest>)'` — `:298` | OK | Fires **after** `CommitTrans` (`:295`) and the matching `LogActLog` call (`:297`); then `ClearEntries` + `Line_ComboBox.SetFocus` (`:301-302`) — see field-clear section. |
| Any exception during Add's DB work | `'Failed on HotCall ASN add, '+e.Message'` — `:309` | OK | `RollbackTrans` if `InTransaction` (`:306-307`), logged (`:308`), dialog shown; entered field values are **left in place** (no `ClearEntries` on this path) so the operator can correct and retry. |

**No confirmation dialogs at all in this form** — Add is a single-click commit with no "are you sure"
step (contrast with `Order.pas`'s Start/ProcessOrder/Cancel, which are all confirmed). This form's
"safety" comes entirely from field-level validation, not from an armed/two-step commit.

## Field clear / repopulate

- `ClearEntries` (`HotCallEntry.pas:72-90`) — the single clear routine, called from three places
  (`Execute:115`, `Clear_ButtonClick:128`, post-commit `Add_ButtonClick:301`):
  - `Line_ComboBox.ItemIndex:=0` (resets to the FIRST line, not blank).
  - `ASN_DateTimePicker.DateTime:=now` (always resets to today, discarding any date the operator
    picked).
  - `ManifestNumber_Edit.Text:=''`.
  - Every `TComboBox` in `ASNItems_GroupBox` is **re-populated from the DB**
    (`Data_Module.SelectSingleField('INV_FORECAST_DETAIL_INF', 'VC_ASSY_PART_NUMBER_CODE', ...)`,
    `:83`) rather than merely blanked — i.e. clearing this form re-queries and rebuilds the 12
    part-number dropdown lists every time.
  - Every `TEdit` in `ASNItems_GroupBox` (the 12 qty fields) is set to `''` (`:87`).
- `Execute` (`:92-125`): `Line_ComboBox.Items.Clear` then refilled from `AD_GetLines` (`:98-114`)
  **once per form open**, followed by `ClearEntries` and `ShowModal`.
- Post-commit reset happens **unconditionally after success**, discarding all 12 rows even if only
  one part number was filled — the operator must re-enter the Line and manifest from scratch for the
  next hot-call entry, even though `Line_ComboBox.ItemIndex:=0` (first item) rather than blank means
  a wrong line could silently ship if the operator doesn't re-check it (**#135-class risk**: a reset
  to index 0 is a NON-empty stale default, not an explicit "nothing selected" state).
- Empty-value handling for the qty/part pairs is intertwined validation logic, not simple clearing:
  a comboqty pair where BOTH are blank is silently skipped/ignored (no error, not even attempted) —
  see the loop at `:178-200`, which only complains when a combo has a NON-blank qty pair with an
  invalid/zero value, or (separately, `:204-218`) when an edit has a value but its paired combo is
  blank. **A row where only the combo is filled and the qty is left blank is silently dropped** — no
  error message for that specific combination; it simply isn't picked up in the write loop at
  `:258-285` because the write loop only iterates `TEdit`s with non-empty `.text` (`:262`).

## Focus & keyboard

- `FormShow` (`HotCallEntry.pas:131-134`): `Line_ComboBox.SetFocus` — always the initial focus target
  on show (also functions as the `ActiveControl`, since no `ActiveControl` property is set in the
  `.dfm`).
- Post-validation-failure focus jumps (all in `Add_ButtonClick`): `Line_ComboBox.SetFocus` (`:146`),
  `ManifestNumber_Edit.SetFocus` (`:160`), the failing qty `TEdit.SetFocus` (`:188`,`:195`), the
  failing part-number `TComboBox.SetFocus` (`:213`) — every validation failure moves focus to the
  exact offending control.
- Post-success: `Line_ComboBox.SetFocus` (`:302`) — focus returns to the top of the form, mirroring
  the initial-show focus target, ready for the next hot-call entry.
- No `Default`/`Cancel` button flags in `HotCallEntry.dfm` (confirmed via grep) **except**
  `Close_Button` has `ModalResult = 2` (`mrCancel`) at `HotCallEntry.dfm:33` — this makes `Close_Button`
  respond to the VCL's Escape-key-triggers-mrCancel-button convention (a `TForm`'s default `Escape`
  behavior fires whichever button has `ModalResult=mrCancel`), even though `Cancel:=True` is not
  explicitly set — `[UNVERIFIED — confirm before use: whether `ModalResult=2` alone is sufficient for
  Escape-to-close without an explicit `Cancel=True`; Delphi's default form-level Escape handling
  keys off `ModalResult`, but this should be confirmed against the live form if Escape behavior
  matters to the rebuild]`. No button has `Default=True`, so Enter does not auto-trigger Add.
- No `KeyPreview`/`OnKeyPress` handler on the form.

## Enable/disable state machine

There is **no enable/disable state machine on this form** — `Add_Button`, `Clear_Button`, and
`Close_Button` are all permanently enabled (no `.Enabled:=` assignment anywhere in
`HotCallEntry.pas`). All gating is done via validation-then-`exit`, not via disabling controls
pre-emptively.

## Error surfacing

100% `ShowMessage` dialogs for both validation and DB/COM exceptions (table above); DB-write failures
additionally log via `Data_Module.LogActLog('ERROR', ...)` with matching text (`:308`/`:309`). The
Line-list load failure surfaces via `Data_Module.LogActLog('ERROR', ...)` (`:107`) alongside its
dialog (`:106`). No status-label surfacing anywhere in this form.

## Cross-refs

- `docs/analysis/edi/hotcall-coverage-analysis.md` — full proc/write-seam spec for this form
  (`INSERT_ASNInfo`/`INSERT_ASNDetail`/`AD_UpdateEIN`, EIN model, live-DB row verification, and the
  M1 ASN write-seam reuse this form shares with the automated 856 path).
