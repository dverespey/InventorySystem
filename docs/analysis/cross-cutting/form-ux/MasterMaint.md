# Form-UX semantics: `TMasterMaint_Form` — `MasterMaint.pas` / `MasterMaint.dfm`

Hub/menu form. Hosts no data itself — it modally `Create`s/`Free`s each master-data child form
one at a time and re-`Show`s itself when the child closes. No DB dataset of its own.

## Dialogs & confirmations
- **Error dialog on launch failure** — `MasterMaint.pas:82-85`, inside `Execute`'s `except`:
  `showMessage('Unable to generate Master Date Maintenance screen.' + #13 + 'ERROR:' + #13 + E.Message)`.
  Plain `ShowMessage` (single OK, no icon/button choice). Only fires if `ShowModal` itself raises —
  none of the button-click handlers are guarded, so a child form's own `Create`/`Execute` exception
  is NOT caught here (propagates to the child's own `Execute` try/except instead, see each child doc).
- No confirmation dialogs on this hub itself (no delete/close-with-unsaved-changes — it has no
  editable fields).

## Field clear / repopulate
- N/A — no data-bound fields on this form.

## Focus & keyboard
- No `ActiveControl` set in the `.dfm`; first tab-order control is `SupMaster_Button` (`TabOrder = 0`,
  `MasterMaint.dfm:38`).
- `Close_Button` has `ModalResult = 2` (`mrCancel`) (`MasterMaint.dfm:74`) — this is the form's only
  `ModalResult`-bearing control, making it the de facto Cancel/Esc-equivalent action, but **no
  `Cancel = True` property is set** on it, and no button has `Default = True` — so there is no
  keyboard-Enter/Esc shortcut wired via VCL defaults; all navigation is by mnemonic accelerator
  (`&Supplier Master`, `Si&ze Master`, etc., every caption carries an `&` accelerator letter).
- Each child-launch handler follows the same **Hide → Create/Execute/Free child → Show** pattern
  (`MasterMaint.pas:93-196`) — the hub form visually disappears while a child is modal, then
  reappears. `SupMaster_ButtonClick` and `ASNINVOIVE_ButtonClick` wrap this in `try/finally` so `Show`
  always runs even if the child raises (`:95-102`, `:188-195`); the other seven handlers
  (`SizeMaster_ButtonClick`, `AssyRatioMaster_ButtonClick`, `PartsStockMaster_ButtonClick`,
  `ForecastDetail_ButtonClick`, `Button1Click` [Logistics], `RenbanGroupMaster_ButtonClick`,
  `MonthlyPO_ButtonClick`) have **no try/finally** — if `Create` or `Execute` on the child raises,
  the hub form is left `Hide`-den with no code path to `Show` it again (`:106-183`). **Hazard:** an
  unhandled exception in seven of nine child launches would strand the operator on an invisible
  hub window.

## Enable/disable state machine
- Button **visibility**, not enable/disable, is the state machine, computed once in `Execute`
  (`MasterMaint.pas:60-90`), before `ShowModal`, driven by `Data_Module` config flags — this is
  **DataModule/DB-driven**, not form-internal state:
  - `MonthlyPO_Button.Visible := Data_Module.fiPOEDISupport.AsBoolean` (`:66-69`).
  - If `Data_Module.fiGenerateEDI.AsBoolean`: `MonthlyPO_Button.Caption := 'Manifest Cost'`,
    `AssyRatioMaster_Button.Visible := FALSE`, `ASNINVOIVE_Button.Visible := TRUE` (`:71-76`).
  - **`AssyRatioMaster_Button.Visible := FALSE` unconditionally, `// not used yet`** (`:78`) — this
    runs regardless of the branch above, so the ASSY/Ratio Master button is **always hidden**. The
    child form (`AssyRatioMaster.pas`) is fully compiled (`InventorySystem.dpr:17`) and functional but
    **unreachable from this menu in the shipping app** — see
    `docs/analysis/assembly/assy-ratio-master.md` §1.1 for the confirmed dead-path finding.
  - `MonthlyPO_ButtonClick` re-checks `Data_Module.fiGenerateEDI.AsBoolean` at click time (`:167`) to
    decide whether to open `ManifestCostMaster` or `MonthlyPOMaster` — the SAME flag gates both the
    caption/label swap at open-time and the destination form at click-time, so they can't disagree
    within one session.
  - `ASNINVOIVE_Button` defaults `Visible = False` in the `.dfm` (`MasterMaint.dfm:120`) and is only
    flipped `TRUE` in the `fiGenerateEDI` branch.

## Error surfacing
- Only one surfaced error path: the `Execute` wrapper `ShowMessage` above. Each child form has its
  own independent error dialog (see child docs) — errors inside a child's own DB calls never reach
  this hub's dialog.

## Cross-refs
- `docs/analysis/master-data/master-maint.md` (proc/data spec — none; this hub has no dataset).
- `docs/analysis/assembly/assy-ratio-master.md` §1.1 (confirms the dead `AssyRatioMaster_Button`
  finding independently).
