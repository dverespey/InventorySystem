# Form-UX semantics: `TEDIUpload_Form` — `EDIUpload.pas` / `EDIUpload.dfm`

**Confirmed LIVE** (`InventorySystem.dpr:53`). The inbound EDI funnel: polls a drop directory,
sniffs/dispatches 830/862/997/824/820 documents. UI is a `THistory` log pane + one initially-hidden
`OKButton` + a non-visual `TCopyFile` component — effectively headless, same shape as
`ForecastBreakdownF`. Full data/proc/wire-format spec: `docs/analysis/edi/edi-upload.md`.

## Dialogs & confirmations
- **None.** A full-text search of `EDIUpload.pas` for `ShowMessage`/`MessageDlg`/
  `Application.MessageBox` returns **zero hits**. Every condition this form can encounter — file
  found, trading-partner match/no-match, transaction-type dispatch, 997 accept/reject, copy
  failure — surfaces **only** through `Hist.Append` (the log pane) and
  `Data_Module.LogActLog` (the Activity-DB audit trail). There is no confirmation of any kind
  before archiving/moving a file, and no armed step before the 830 hand-off to
  `ForecastBreakdown_Form` (`:94-99`) — dispatch is unconditional once the DUNS + transaction-type
  checks pass.
- **The one thing resembling an interactive gate** is the busy-wait at the very end
  (`:465-469`, see Focus & keyboard) — the operator's only action on this form is to click OK once
  the whole batch (all files in the drop dir) has been processed; there is nothing to approve or
  reject along the way.

## Field clear / repopulate
- **No editable fields exist on this form** — nothing to clear or repopulate. The `Hist` log is
  append-only per `Execute` call; each invocation is a fresh `TEDIUpload_Form.Create`/`Free`
  (`MainMenu.pas:2904-2908`), so the log always starts empty.
- **`fClosed` is reset at the top of every `Execute`** (`:48`), and the module-level shared state
  `Data_Module.EIN`/`EINStatus`/`EINType` is overwritten fresh for **each** 997 AK1 loop iteration
  (`:199-201`) — but `EIN` itself (the local `string`, not `Data_Module.EIN`) is only initialized
  once at the very top of `Execute` (`:47 EIN:=''`), **not per file** in the outer file loop. This
  is the archive-mislabel hazard already flagged in the data-side spec (§4.8 of
  `edi-upload.md`): a 997 sets `EIN`, and if a later file in the same batch doesn't hit the 997
  branch, the stale `EIN` from the previous file leaks into that later file's archive name
  (`:425-428`). Recorded here because it's a **shared-mutable-state carry-over across "records"
  processed in one form-lifetime**, the same class of bug as a field that doesn't get cleared
  between records in a data-entry form (P9-adjacent, though this is a local `var`, not a
  `Data_Module` property).

## Focus & keyboard
- No `ActiveControl` set in `.dfm`; default focus is `Hist` (`TabOrder=0`), a non-interactive log
  control — same pattern as `ForecastBreakdownF`/`ManualForecast`.
- **No `SetFocus` calls anywhere in this unit.**
- **`OKButton` starts hidden** (`EDIUpload.dfm:34 Visible = False`) and has no `Default`/`Cancel`
  flag — Enter/Escape do nothing; the operator must click it once it appears.
- **The form hides itself during the 830 sub-run**: `Hide; Application.ProcessMessages;` (`:92-93`)
  before creating/showing `ForecastBreakdown_Form`, then `Show; Application.ProcessMessages;`
  (`:100-101`) immediately after that sub-form is freed — so during an 830's processing, only
  `ForecastBreakdown_Form`'s own window (and its own log pane) is visible; `EDIUpload_Form`
  reappears (with its log unchanged, the hide/show doesn't clear it) once the sub-run completes.
- **Busy-wait blocks the UI thread until OK is clicked**, identical pattern to
  `ForecastBreakdownF`: `while not fclosed do begin application.ProcessMessages;
  sleep(500); end` (`:465-469`), inside the outer `finally` so it always runs — even if the
  `FindFirst`/dispatch loop raised an exception, the operator still gets a chance to read the log
  and dismiss. `OKButtonClick` (`:495-498`) just flips `fClosed:=TRUE`.

## Enable/disable state machine
- **None beyond `OKButton.Visible`** (hidden until the whole directory scan completes, `:464`,
  inside the `finally`). No other control's `Enabled`/`Visible` is touched in this unit.
- **The button that opens this form is itself gated**, one level up: `EDIUploadBox` (the group box
  holding the launch button) is shown only when `Data_Module.fiGenerateEDI.AsBoolean` is true
  (`MainMenu.pas:2911-2916`, `FormShow`) — a per-site config flag, not form-local state, but the
  effective enable/disable gate for reaching this screen at all.

## Error surfacing
- **Single channel: `Hist.Append` + `LogActLog`.** Every failure path in this unit — unknown DUNS
  (`:198-199`), non-830/862/997/824/820 transaction type (`:414-417`, logged as an ERROR-tagged
  line, not a raised exception), 824 Excel-open failure (`:302-304`, its own nested
  `try/except`), file-copy/move failure (`:430-434`) — is caught locally and only ever appended to
  the log pane / written to the Activity DB. **No exception from any per-file branch escapes to
  cause a modal error dialog**; the closest thing to a hard stop is the file *staying in the drop
  directory* to be re-scanned next run (an operational retry, not a UI error state).
- This is a genuinely different error-surfacing posture from `UserAdmin`/`NewPassword`/
  `ConfirmPassword` (all of which use `MessageDlg` for validation errors) and even from
  `ForecastBreakdownF` (which does raise two `ShowMessage`s for DB/Excel access failures,
  `edi_upload`'s sibling breakdown form) — **EDIUpload never blocks the batch with a dialog**, it
  only logs and continues to the next file.

## Cross-refs
- `docs/analysis/edi/edi-upload.md` — full proc/wire-format/business-rule spec (§4.1–4.8) this
  file's UX wraps; §4.8 is the data-side statement of the `EDIFileNumber`/`EIN` carry-over bug
  cited above under Field clear/repopulate.
- `docs/analysis/cross-cutting/form-ux/ForecastBreakdownF.md` — the form this screen `Hide`s
  itself for and hands the 830 file to.
