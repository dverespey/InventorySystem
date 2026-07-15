# Form-UX semantics: `TForecastBreakdown_Form` — `ForecastBreakdownF.pas` / `ForecastBreakdownF.dfm`

**Confirmed LIVE** (`InventorySystem.dpr:28`). Effectively headless: parse a forecast feed →
explode → write. The visible UI is a `THistory` log pane (`Hist`) + one `OKButton`; almost the
entire 1489-line unit is batch logic (specced in `docs/analysis/forecasting/forecast-breakdown.md`).
This file is the UX-only layer around that batch run: what the operator sees, what they're asked,
and where the run blocks for their input.

## Dialogs & confirmations
- **Two mid-run Yes/No confirmations, both gate whether processing continues** (`mtConfirmation`,
  `[mbYes, mbNo]`, no explicit default in code — VCL defaults focus/Enter to the first button,
  `mbYes`):
  - *DB assemblies missing from the feed* — `'There are '+IntToStr(dbMissing.count)+', in the
    database and not in the forecast. Continue processing?'` (`ForecastBreakdownF.pas:865-866`).
    `mrNo` → `result:=False; exit` — **aborts the entire import immediately**, before any delete/
    upsert has run for this batch (this check runs inside `ScanPartnumber`, which is called
    *before* `DeleteBreakdown`/`UpdateForecast` in `Execute`, `:320-335`). `mrYes` → continues and
    additionally writes an Excel exception report (`ForecastDBError<ts>.xls`).
  - *Feed part numbers missing from the DB* — `'There are '+IntToStr(skip)+', forecast records not
    in the database. Continue processing?'` (`:907-908`). Same `mrNo` → abort-before-writes
    semantics; `mrYes` continues (the `skip`-flagged entries are excluded from the explosion per
    `ScanPartnumber`'s earlier per-entry `Skip:=True` marking, `:755`) and writes
    `ForecastRecError<ts>.xls`.
  - **Both prompts fire only if their respective count is nonzero** (`if dbmissing.Count > 0`
    `:863`; `if skip <> 0` `:905`) — a clean feed with full part-number reconciliation shows
    neither dialog and the run proceeds silently.
- **Table-access error during reconciliation** — `ShowMessage('Error on
  INV_FORECAST_DETAIL_INF table access, '+e.Message)` (`:792`, inside `ScanPartnumber`'s except)
  and again verbatim inside `UpdateForecast`'s per-entry except (`:1292`) — both single-OK, no
  severity distinction, and both ALSO log to the Activity DB (`LogActLog('ERROR', …)`).
- **Excel template-save failure** — `Showmessage('Cannot save excel report template(…), '+
  e.Message)` (`:853`, `:896`) for the two exception-report writers; same log-pairing pattern.
- **No confirmation before the delete-then-rebuild itself** (`DeleteBreakdown` per non-skipped
  entry, `:326-329`) — the two Yes/No prompts above are the *only* gates on the whole run; once
  past them, deletes/upserts run unconditionally per entry.
- **Unrecognized/malformed file:** no dedicated dialog — file-open failure, bad EDI-830 tag
  (not `'830'`, `:206-213`), and unknown-DUNS abort (`:196-201`) all surface as `Hist.Append` log
  lines only (see Error surfacing below), not a modal.

## Field clear / repopulate
- This form has **no editable input fields** — `filename`/`SupplierCode` are published properties
  set by the caller (`EDIUpload.pas:95-96` or `ForecastUploadBreakDown.pas:101-102`/
  `UploadBreakDown.pas:185-186`) before `Show`; there is nothing on the form itself to clear or
  repopulate. The `Hist` log pane is append-only for the life of one `Execute` call (no `Clear`
  call anywhere in this unit) — re-running the form (a fresh `Create`/`Free` per invocation, per
  every caller) starts with an empty log because it's a new instance, not because anything resets
  it.
- **`OKButton.Visible`** is the only mutable UI state: hidden at the start of `Execute`
  (`:171 OKButton.Visible:=False`), shown only after the entire batch (parse → explode → file
  emission) completes or the outer `except` fires (`:582 OKButton.Visible:=True`, reached via the
  `finally` that also runs after an exception). So **even a failed run still reveals OK** — the
  operator is never left with no way to dismiss the form.

## Focus & keyboard
- No `ActiveControl` set in `.dfm` (only two controls exist: `Hist` `TabOrder=0`, `OKButton`
  `TabOrder=1` — default focus goes to `Hist`, a non-interactive log control).
- **No `SetFocus` calls anywhere in this unit.**
- `OKButton` has no `Default`/`Cancel` flag in `.dfm` (`ForecastBreakdownF.dfm:29-37`) — Enter does
  not trigger OK; the operator must click it (or Tab to it and press Space/Enter — standard VCL
  button activation once focused, not a form-authored behavior).
- **Busy-wait blocks the UI thread until OK is clicked**: `while not fclosed do begin
  application.ProcessMessages; sleep(500); end` (`:583-587`); `OKButtonClick` (`:1482-1484`) just
  sets `fclosed:=true`. The form remains responsive to repaint/click via `ProcessMessages`, but no
  other action is possible while this loop runs (it *is* the entire post-processing state — no
  Cancel/Abort button exists once a run starts).

## Enable/disable state machine
- **None.** The only visibility state is `OKButton.Visible` (see above); there is no `Enabled`
  toggling anywhere in `ForecastBreakdownF.pas`. There is nothing else on the form to gate.

## Error surfacing
- Two channels, used inconsistently by section:
  1. **`Hist.Append` (log pane) + `Data_Module.LogActLog`** for expected/handled conditions:
     unknown trading-partner DUNS (`:198-199`), wrong EDI transaction type (`:210-211`), the
     top-level `except` around the whole parse/explode/emit pipeline (`:571-577` — *"Unable to
     load forecast, "+e.Message*), and the day-to-day progress narration ("Trading Partner
     Search:", "EDI Count=", "N total records to process", "Forecast processing complete, Press OK
     to continue" `:567`).
  2. **`ShowMessage` (modal, single OK) + `LogActLog`** for the two table-access/Excel-save
     failure points cited above (`:792`, `:853`, `:896`, `:1292`) — these interrupt the batch with
     a blocking dialog mid-run, unlike the channel-1 conditions which only ever log and continue
     (or `exit` silently on the DUNS/830-type checks).
- **No distinct "row-level error" surface** — a bad `FST`/`LIN` line inside the EDI parse loop
  (`:245-298`) has no per-line try/except; a malformed line raises up into the outer `except`
  (`:571`), aborting the **entire file's** processing, not just the offending assembly.

## Cross-refs
- `docs/analysis/forecasting/forecast-breakdown.md` — the full data/proc/business-rule spec this
  file's UX wraps (§4 for the parse/explosion/day-spread algorithm; §5 for the "effectively
  headless" UI summary this doc expands on).
- `docs/analysis/edi/edi-upload.md` §4.3 — the 830 dispatch boundary: `EDIUpload_Form` `Hide`s
  itself, creates/`Show`s/`Execute`s this form, then `Show`s itself again
  (`EDIUpload.pas:92-100`) — so from the 830-ingest path this form is the *only* visible window
  during breakdown processing.
- `docs/analysis/cross-cutting/form-ux/ForecastUploadBreakDown.md` — a **second, dead** entry path
  that would have called this same form (with a property, `FileKind`, that doesn't exist on the
  published interface — see that file for why it's dead).
