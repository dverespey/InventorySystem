# Form-UX semantics: `TManualForecast_Form` — `ManualForecast.pas` / `ManualForecast.dfm`

**Confirmed LIVE** (`InventorySystem.dpr:43`). "Buildout" forecast entry: operator supplies an
Excel workbook + a Start/End date range; the form writes a synthetic `buildout.prelftp` text feed
(one line per part) in the same fixed-width shape `ForecastBreakdownF.ScanLine` parses, then
**should** hand off to the breakdown importer — but that hand-off is commented out (see below).
This is the **"Q6" half-wired path** flagged in `docs/analysis/forecasting/forecast-breakdown.md`
§8 Q6 — this file adds the UX-layer evidence for that finding.

## Dialogs & confirmations
- **No armed/two-step confirmation anywhere on this form.** Writing the `.prelftp` file and running
  the Excel walk both fire directly off `StartButtonClick` (`ManualForecast.pas:58-151`) with no
  "are you sure" gate — contrast the delete-confirmations elsewhere in the app.
- **Date-order validation is delivered via plain `ShowMessage`, not `MessageDlg`** (single implicit
  OK, no button-set choice):
  - `'End date must be a Sunday'` (`:158`) — end date entered but not a Sunday.
  - `'End date must be greater than start date'` (`:165`) — only shown if `startdate.Text <> ''`
    (i.e. suppressed on the very first keystroke into Enddate before Startdate has a value).
  - `'Start date must be a Sunday'` (`:175`).
  - `'Start date must be less than end date'` (`:182`) — same suppression condition, mirrored
    (`enddate.Text <> ''`).
- **Runtime/Excel failure** — `StartButtonClick`'s `except` block (`:140-150`) does **not** show
  any dialog. It logs via `Data_Module.LogActLog('ERROR', 'Failed on Manual Forecast, '+e.Message)`
  and appends the same text to the `Hist` log pane (`:143-144`) — errors here are **silent to a
  modal dialog**, visible only in the on-screen log and the Activity DB. This differs from
  `ForecastBreakdownF`'s pattern of at least logging comparable failures the same way (both forms
  favor the log pane over dialogs for processing errors), but differs from `UserAdmin`/`SizeMaster`
  which use `MessageDlg` for validation errors.

## Field clear / repopulate
- **The breakdown hand-off is dead code, commented out in-place** (`ManualForecast.pas:128-135`):
  ```
  {breakdown:=TForecastBreakdown_Form.Create(self);
  breakdown.filename:=Data_Module.fiForecastInputDir.AsString+'\buildout.prelftp';
  breakdown.SupplierCode:=data_Module.fiSupplierCode.AsString;
  breakdown.Show;
  if not breakdown.Execute then
    Hist.Append('Forecast has not been added, please retry');
  breakdown.Free;
  close;}
  ```
  So today the operator must **manually** re-run the Forecast breakdown screen (or EDIUpload) on
  the generated `buildout.prelftp` file — the button click only produces the file and stops.
- **After a successful run** (`:136-139`): `StartButton.Visible:=False`; `startdate.Clear`;
  `enddate.Clear` — the Start button is hidden and both dates are wiped, forcing the operator to
  re-enter both dates (each triggering re-validation via `OnChange`) before `StartButton` can
  reappear. There is **no** explicit re-clear of these fields in the `except` branch — if the run
  fails partway, the dates are left populated and the Start button remains visible (its `Visible`
  flip only happens in the success path, after `CloseFile`).
- **`Execute`** (`:46-51`) is a near-no-op: just `ShowModal`. No field seeding/reset happens here —
  whatever the dates showed from a previous invocation persists, but the caller (`UploadBreakDown`,
  see Cross-refs) `Create`s a fresh `TManualForecast_Form` and `Free`s it after `Execute` returns
  (`UploadBreakDown.pas:209-213`), so in practice each invocation starts from the `.dfm` design-time
  defaults (both date edits blank).
- **`StartButton` starts hidden** (`ManualForecast.dfm:59 Visible = False`) and is only revealed by
  the date `OnChange` handlers once both dates validate (Sunday-to-Sunday, end > start).

## Focus & keyboard
- No `ActiveControl` set in `.dfm` — default focus goes to the first tab-order control, `Hist`
  (`TabOrder=0`, a log/list control, not an edit) since Startdate/Enddate are `TabOrder=3/4`.
- No `SetFocus` calls anywhere in `ManualForecast.pas`.
- `Button1` (Close) has no `Default`/`Cancel` property set (`.dfm:43-51`) — Enter/Escape are not
  VCL-wired to any button on this form.
- Startdate/Enddate are `TNUMMIBmDateEdit` with `Options = [doButtonTabStop, doCanClear, doCanPopup,
  doIsMasked, doShowCancel, doShowToday]` (`.dfm:94,130`) — the shared date-picker control's own
  keyboard/popup semantics apply (out of scope here; component-level, not form-level).

## Enable/disable state machine
- **Only one state toggle exists**: `StartButton.Visible` — hidden by default, shown only when
  BOTH date fields independently pass validation (each `OnChange` handler only sets
  `Visible:=TRUE` when ITS OWN date checks out; there is no re-validation of the *other* field when
  one changes, so entering a valid Enddate first, then an invalid Startdate, correctly hides the
  button via `StartdateChange`'s own `else` branch — but the interaction is two independent
  one-way checks, not a joint validity check recomputed from both fields together).
- No other Enabled/Visible toggling in this unit.

## Error surfacing
- Date-order/day-of-week violations → `ShowMessage` (single OK, no severity icon distinction from
  `MessageDlg`).
- Excel/file-I/O failures inside `StartButtonClick` → **log pane + Activity log only**, no dialog
  (`:140-150`).
- No validation exists for the uploaded Excel file's structure — a malformed sheet raises inside
  the `try` and is caught by the same generic handler above (no field-specific error path).

## Cross-refs
- **Entry point (confirmed live):** `UploadBreakDown.pas:207-214`, the `bBuildout` branch of the
  shared 6-way breakdown-kind dispatcher (`UploadBreakDown.pas` — see `forecast-breakdown.md` §5
  for the dispatcher itself). Note the dispatcher's own `ManualForecast_Form.Show` call is
  **commented out** (`UploadBreakDown.pas:211`) — harmless since `Execute` calls `ShowModal`
  directly, but worth noting as another commented-out line in this same code path (alongside the
  breakdown hand-off below).
- `docs/analysis/forecasting/forecast-breakdown.md` §8 Q6 (the same commented-out hand-off, called
  out from the data/proc side — this file is the paired UX-side citation).
- `docs/analysis/forecasting/forecast-detail.md` (BOM/ratio master this feed ultimately explodes
  against, once re-imported through `ForecastBreakdownF`).
