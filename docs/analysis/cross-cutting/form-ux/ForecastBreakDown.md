# Form-UX semantics: `TForecastBreakDown` — `ForecastBreakDown.pas` (NO `.dfm`)

## Live-vs-dead verdict: **DEAD CODE — confirmed by three independent signals**

1. **Not in the `.dpr` manifest.** `InventorySystem.dpr` has no `ForecastBreakDown in
   'ForecastBreakDown.pas'` line (only `ForecastBreakdownF in 'ForecastBreakdownF.pas'` at
   `InventorySystem.dpr:28` — note the trailing **`F`**, a different unit). Per this repo's
   methodology, a `.pas` absent from the `.dpr` does not ship.
2. **No `.dfm` file exists** (`ls ForecastBreakDown.dfm` → not found). The type declared,
   `TForecastBreakDown = class(TObject)` (`ForecastBreakDown.pas:23`), is a **plain object, not a
   `TForm`** — it was never a visual form to begin with, so there is no UX layer to extract. This
   looks like the **pre-form predecessor** to `ForecastBreakdownF.pas` (same `TWeekData`/`TEntryRec`
   record shapes, same field names, same 2002-2003 change-history block) that was later rewritten
   as a `TForm` descendant and renamed with the `F` suffix.
3. **Zero references from any other compiled unit.** `grep -l "ForecastBreakDown" *.pas` (excluding
   `ForecastBreakDown.pas`/`ForecastBreakdownF.pas`/`ForecastUploadBreakDown.pas` themselves)
   returns nothing — no unit `uses` it, no code creates a `TForecastBreakDown`.

## Form-UX semantics
**N/A — not a form.** There are no dialogs, confirmations, focus rules, or enable/disable state to
extract; the class exposes methods (`Execute`, parse/explode helpers per its own `interface`
section) but no VCL controls. Do not port any UX from this file — its logic-shaped body was
superseded by `ForecastBreakdownF.pas`, which is the file to read for the equivalent (and now
form-based) behavior. See `docs/analysis/cross-cutting/form-ux/ForecastBreakdownF.md`.

## Cross-refs
- `docs/analysis/cross-cutting/form-ux/ForecastBreakdownF.md` — the live successor.
- `docs/analysis/forecasting/forecast-breakdown.md` — confirms `ForecastBreakdownF.pas` (not this
  file) as the live breakdown module (§1: "Registered live in `InventorySystem.dpr:28`").
