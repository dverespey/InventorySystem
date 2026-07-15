# Form-UX semantics: `TForeUpBreakDown_Form` — `ForecastUploadBreakDown.pas` / `.dfm`

## Live-vs-dead verdict: **DEAD CODE — confirmed, and would not compile if reinstated as-is**

1. **Not in the `.dpr` manifest.** `InventorySystem.dpr` has no `ForecastUploadBreakDown in
   'ForecastUploadBreakDown.pas'` line at all (searched the full file — absent). The **live**
   forecast-upload dispatcher is a *different* unit, `UploadBreakDown.pas`
   (`InventorySystem.dpr:7`), a 6-way `TBreakdownKind` dispatcher documented in
   `docs/analysis/forecasting/forecast-breakdown.md` §5. This file (`ForecastUploadBreakDown.pas`)
   is a same-purpose but forecast-only, single-kind predecessor form — its `.dfm` caption
   ("Forecast Information Upload & Break Down") and `.txt`-only file filter (`:67`) confirm it
   pre-dates the generalized dispatcher.
2. **References a property that doesn't exist on its target form — a latent compile error.**
   `Start_ButtonClick` sets `breakdown.FileKind:=fText;` (`ForecastUploadBreakDown.pas:103`), but
   `TForecastBreakdown_Form`'s published interface (`ForecastBreakdownF.pas:65-72`) exposes only
   `filename` and `SupplierCode` — there is **no `FileKind` property** (only a private field
   `fFileKind: TFileKind`, `ForecastBreakdownF.pas:50`, with no accessor). If this unit were added
   back to the `.dpr` it would fail to compile against the current `ForecastBreakdownF.pas`. This
   is strong independent confirmation the two units drifted apart after one was retired.
3. **Zero references from any other compiled unit** (same `grep` sweep as
   `ForecastBreakDown.pas`'s verdict — no hits outside this file itself).

## Form-UX semantics (recorded for completeness — DO NOT PORT; superseded by `UploadBreakDown`)
Since the unit is dead, the following is descriptive only, not a spec to build against.

### Dialogs & confirmations
- **Wrong extension** — `MessageDlg('Not a file with the .TXT extension', mtError, [mbOK], 0)`
  (`:75-76`) if the file dialog's `Options` flags `ofExtensionDifferent`.
- **Uncaught exception in the modal run** — `showMessage('Unable to generate Forecast Info. Upload
  & Break Down screen.' + #13 + 'ERROR:' + #13 + E.Message)` (`:51-54`), from `Execute`'s own
  `try/except`.
- **No file selected yet, Start clicked** — `ShowMessage('Please select a valid file first')`
  (`:111`).
- **Breakdown sub-form reports failure** — `Showmessage('Forecast has not been added, please
  retry')` (`:106`) if `breakdown.Execute` returns `False`.
- **Generic exception in Start** — `ShowMessage('Err: ' + e.Message)` (`:114`).
- No armed/two-step confirmation on Start itself — clicking Start with a file selected runs
  immediately.

### Field clear / repopulate
- `FormCreate` (`:86-91`) seeds `FileName_Edit.Text` to a placeholder string, `'[Type file path and
  name here or click Browse]'`, held in `fDefaultFilePathLabel`/`fFilePathLabel` — this placeholder
  is itself valid `Text`, not blank, so a naive "is it empty" check would miss it (the form instead
  tracks readiness via a separate `fFileSelected` boolean, flipped `True` by **either** picking a
  file via Browse (`:81`) **or** merely editing the text box (`FileName_EditChange`, `:118-121`,
  fires on any keystroke) — so typing garbage into the box also arms Start).
- No reset-on-new/selection-change logic exists beyond the one-time `FormCreate` seed.

### Focus & keyboard
- No `ActiveControl` in `.dfm`; default focus is `FileName_Edit` (`TabOrder=0`).
- No `SetFocus` calls in the unit.
- No `Default`/`Cancel` flags on any button in `.dfm` — only accelerator keys (`&Browse`, `&Start`,
  `&Close`) work.

### Enable/disable state machine
- **None found** — Start/Browse/Close are always enabled; "not ready" is caught reactively (via the
  `ShowMessage`s above), not by disabling controls.

### Error surfacing
- All via `ShowMessage`/`MessageDlg` (see Dialogs above) — no log-pane equivalent to `THistory`
  exists on this form (unlike its live successor `UploadBreakDown`/`ForecastBreakdownF`, which
  narrate progress into a `Hist` pane).

## Cross-refs
- `docs/analysis/forecasting/forecast-breakdown.md` §5 — the live `UploadBreakDown` dispatcher
  (6-way `TBreakdownKind`) that replaced this single-purpose form.
- `docs/analysis/cross-cutting/form-ux/ForecastBreakdownF.md` — the target form this dead unit
  would have driven (with a since-removed `FileKind` property).
