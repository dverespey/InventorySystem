# Form-UX semantics: `TForecastDetail_Form` — `ForecastDetail.pas` / `ForecastDetail.dfm`

CRUD editor over **`INV_FORECAST_DETAIL_INF`** (the assembly→component BOM master: tire/wheel/
valve/film/label/misc1/misc2 part codes + tire/wheel/forecast ratios + assy quantity + broadcast
code, keyed on assembly-part-number × effective-month). Grid `ForecastDetail_DBGrid` bound to
`Data_Module.Inv_DataSet` via `ForecastDetail_DataSource`. Caption is **"Assembly Detail Master"**
(`ForecastDetail.dfm:6`) — cosmetically indistinguishable from `AssyRatioMaster`'s domain, which is
the root of the original mis-mapping this file corrects (see Cross-refs).

**Reachability (confirmed live, unlike `AssyRatioMaster`):** constructed from
`MasterMaint.ForecastDetail_ButtonClick` (`MasterMaint.pas:137-143`) behind `ForecastDetail_Button`,
which — unlike `AssyRatioMaster_Button` (`MasterMaint.pas:78`, unconditionally hidden) — is **never
hidden anywhere in `MasterMaint.Execute`** (`MasterMaint.pas:60-91`). Compiled
(`InventorySystem.dpr:27`) and reachable. This is a live, in-use screen.

## Dialogs & confirmations
- **Insert failure** — `ForecastDetail.pas:486-487`: `MessageDlg('Unable to INSERT ' +
  Data_Module.AssyCode, mtInformation, [mbOk], 0)` if `InsertForecastDetailInfo` returns `False`.
  Single-OK information dialog.
- **Delete confirmation (armed, two-step)** — `:446-450`: `MessageDlg('Are you sure you wish to
  delete' + #13 + Data_Module.AssyCode + ' (' + Data_Module.BeginDatestr + ') from the database?',
  mtWarning, [mbYes, mbNo], 0) = mrYes` gates the call to `DeleteForecastDetailInfo`. On `mrNo` (or
  dismissed) nothing is deleted, but see the stale-key hazard below.
  - **Dialog-text/action mismatch hazard:** `Delete_ButtonClick` calls `HoldDetails(False)`
    (`:446`, the UI→`Data_Module` direction) **before** building the confirm text — this refreshes
    `Data_Module.AssyCode`/`BeginDatestr` from whatever is *currently typed* in
    `ForecastPartsCode_ComboBox`/`EffectiveMonth_ComboBox`, but does **not** touch
    `Data_Module.RecordID` (only the `fFromGrid=True` branch sets `RecordID`,
    `HoldDetails.pas:193`). The actual delete, however, is keyed **only** on `@RecordID`
    (`DataModule.pas:2451-2452`, `DELETE_ForecastDetail;1`). So if an operator selects a grid row
    (arming `RecordID`) then edits the Assy/Month fields without re-selecting, the confirm dialog
    shows the **freshly typed** values while the delete silently acts on the **originally selected**
    row's `RecordID` — text and action can disagree. `Update_ButtonClick` has the analogous
    `@RecordID`-keyed write (`DataModule.pas:2408-2409`) but that's coherent "edit this row" semantics
    since Update has no confirm-text step to go stale.
- **Search "no criteria" guard — does NOT short-circuit** — `:381-385`: if
  `ForecastPartsCode_ComboBox`, `BroadcastCode_Edit`, and `EffectiveMonth_ComboBox` are all blank,
  `ShowMessage('Search on Forecast(Assembly) Part Number, Broadcast Code and/or Effective Month')` +
  `ForecastPartsCode_ComboBox.SetFocus` fires — but this is only one branch of an `if/else if` chain
  with **no `exit`/`Result:=False` afterward**; execution falls through unconditionally to
  `Data_Module.ClearControls(...)` / `Filtered := True` / `RecordCount` check (`:388-395`) using
  whatever `Inv_DataSet.Filter` string was left from the **previous** search (never reset to `''`
  here). A blank-criteria search can therefore silently re-apply a stale prior filter instead of
  showing "no criteria" and stopping.
- **Search "not found"** — `:415`: `ShowMessage('No matches were found for your query.')`.
- **Search-time error** — `:398-399`: `ShowMessage('Error in Search' + #13 + e.Message)`.
- **Combo-population failure** — `GetAssy` (`:156-172`): `MessageDlg('Unable to get a list of
  parts.', mtError, [mbOK], 0)` if `SelectSingleField` raises. Called from `FormCreate` (`:151`) and
  again after a successful Insert (`:492`, to pick up a newly-added assembly code). **Only this one
  combo's population is guarded** — the other 7 `SelectDependantSingleField` calls for
  Tire/Wheel/Valve/Film/Label/Misc1/Misc2 part-number combos (`:144-150`) have **no individual
  try/except**; an exception there propagates unguarded (see Error surfacing).
- **Ratio validation is warn-only, never blocking** — see Enable/disable & Error surfacing below;
  this is the form's most consequential UX-vs-data-integrity gap.
- Grid carries `dgConfirmDelete` (`ForecastDetail.dfm:112`) — VCL-native secondary confirm on the
  grid itself, independent of `Delete_Button`.

## Field clear / repopulate
- **Ratio/qty validation in `HoldDetails(fFromGrid=False)` (`:213-310`, the UI→`Data_Module` write
  path) is entirely warn-only — confirmed, does NOT gate the write:**
  - `ForecastRatio_Edit`: `TryStrToInt` failure → `ShowMessage('Invalid ratio')` + `SetFocus`
    (`:268-272`); the **0–100 range check for this field is commented out entirely**
    (`:276-281`, comment "Allow multiple ratios over 100" — a deliberate legacy business decision,
    not dead code left by accident).
  - `TireForecastRatio_Edit` / `WheelForecastRatio_Edit`: same `TryStrToInt`-failure ShowMessage
    pattern (`:283-289`, `:296-302`), **plus** a live (not commented out) range check —
    `if (TireRatio1 > 100) or (TireRatio1 < 0) then ShowMessage('Ratio must be between 0 and 100')`
    (`:290-294`) and the identical check for `WheelRatio1` (`:303-307`). **Neither check sets any
    result/abort flag** — `HoldDetails` is a `procedure`, not a `function`, has no return value, and
    nothing downstream inspects one. `Insert_ButtonClick`/`Update_ButtonClick` call `HoldDetails(False)`
    unconditionally after their only gate, `Validate` (`:481-483`, `:462-464`), and `Validate` itself
    (`:502-513`) **unconditionally `Result := True`** — its one real check (Forecast Part Code length)
    is commented out (`:507-512`). **Net: there is no field-level validation anywhere in this form
    that can prevent an Insert or Update from proceeding** — a value of `150` typed into
    `TireForecastRatio_Edit` pops the warning dialog and is written to the database exactly as typed.
    **[VERIFIED by direct read: this confirms the reviewer's claim (b) — the 0–100 ShowMessage does
    NOT block the write.]**
  - Both stacked-dialog and focus-clobber effects from `AssyRatioMaster`'s `Verify` do **not** apply
    here: since neither check has an `exit`, if BOTH tire and wheel ratios are out of range, BOTH
    `ShowMessage`s fire in sequence (same "two stacked dialogs" pattern as `AssyRatioMaster.Verify`,
    `AssyRatioMaster.md:26-35`), and `WheelForecastRatio_Edit.SetFocus` (the later call) wins over
    `TireForecastRatio_Edit.SetFocus` — but it's moot for gating since the write happens regardless.
- **`AssyQty_RadioGroup` domain and round-trip — CONFIRMS claim (a) on the READ side, but reveals an
  ASYMMETRIC WRITE side the reviewer's claim doesn't mention:**
  - **Read (`Data_Module.Quantity` → UI), `SetDetailBoxes` `:332-339`:**
    ```
    case Quantity of
      0: AssyQty_RadioGroup.ItemIndex := 0;
      2: AssyQty_RadioGroup.ItemIndex := 1;
      4: AssyQty_RadioGroup.ItemIndex := 2;
      5: AssyQty_RadioGroup.ItemIndex := 3;
      Else
        AssyQty_RadioGroup.ItemIndex := 0;
    End;
    ```
    Domain `{0,2,4,5}` with `Else → 0` (i.e. same visual slot as an actual `Quantity=0`). **Confirmed
    — matches claim (a).** (The component itself is declared at `ForecastDetail.pas:47`
    `AssyQty_RadioGroup: TRadioGroup;` — that line is only the field declaration, not the domain; the
    domain data lives at `:332-339` per above and the `.dfm` `Items.Strings` below.)
  - **Write (UI → `Data_Module.Quantity`), `HoldDetails` `:222-229`:**
    ```
    case AssyQty_RadioGroup.ItemIndex of
      0: Data_Module.Quantity := 1;
      1: Data_Module.Quantity := 2;
      2: Data_Module.Quantity := 4;
      3: Data_Module.Quantity := 5;
      Else
         Data_Module.Quantity := 1;
    end;
    ```
    Domain `{1,2,4,5}` with `Else → 1` — **NOT the same domain as the read side.** ItemIndex 0 reads
    back as `1` on write, not `0`.
  - **`.dfm` confirms the visible labels are `'1','2','4','5'`, `ItemIndex = 0` default**
    (`ForecastDetail.dfm:335-348`) — so index 0's on-screen label is literally `"1"`, yet
    `SetDetailBoxes` also lands a **stored `Quantity = 0`** on that same index-0/"1" slot. **Net
    effect: a row whose DB `Quantity` is `0` displays as radio-button `"1"` (indistinguishable on
    screen from a row whose `Quantity` genuinely is `1`), and if that row is then re-saved (Update)
    without the operator touching the radio group, its `Quantity` is silently coerced from `0` to
    `1`** — a real, demonstrable round-trip data-mutation bug hiding in the read/write asymmetry.
    This is additive to claim (a), which only describes the read-side default-to-0 behavior; the
    write-side default-to-1 (and the general read⇄write domain mismatch) is a **separate, more
    consequential** finding a rebuild must reproduce-or-flag (per the project's divergence rule — a
    silent `0→1` coercion changes a number Toyota's forecast explosion consumes downstream via
    `ForecastBreakdownF`, so this is a David-decision candidate, not a quiet "fix").
- **`ClearControls(ForecastDetail_Panel)` does NOT reach `AssyQty_RadioGroup` at all** —
  `TData_Module.ClearControls` (`DataModule.pas:5976-6017`, the `TPanel` branch) only special-cases
  `TEdit`/`TMaskEdit`/`TCheckBox`/`TMemo`/`TComboBox`/`TNUMMIBmDateEdit`; there is **no `TRadioGroup`
  branch**. So `Clear_Button` (`:174-182`) and `Delete_ButtonClick`'s post-delete clear (`:454`) leave
  `AssyQty_RadioGroup`'s `ItemIndex` exactly as it was — the radio group is never blanked/reset by
  either path, only ever changed by `SetDetailBoxes` (grid selection/search/Insert/Update echo) or
  direct operator click.
- **`Misc2PartNum_ComboBox` is structurally OUTSIDE `ForecastDetail_Panel` in the `.dfm` — a second,
  independent field-clear gap, distinct from the radio-group one above.** `ForecastDetail_Panel`
  opens at `ForecastDetail.dfm:122` and its matching `end` is at `:408`; `Misc2PartNum_ComboBox` is
  declared at `:409-417`, **after** that closing `end` — i.e. it is a sibling of the panel at the
  **form** level, not a VCL child of it, even though its screen position (`Left=440, Top=214`,
  `:410-411`) places it visually inside the panel's on-screen rectangle (panel spans
  `Left=21..606, Top=31..297`). Its neighbor label `Label10` ("Label Part Number" — a duplicate
  caption of `Label8`, almost certainly a copy/paste-and-move artifact that never got renamed,
  `:34-42` vs `:238-246`) is likewise form-level, while every OTHER Misc/Label control
  (`Label11`/`Misc1PartNum_ComboBox`, `:256-264`/`:398-407`) is properly panel-nested. Consequence:
  **`ClearControls(ForecastDetail_Panel)` never touches `Misc2PartNum_ComboBox`** — on `Clear_Button`,
  on `Delete_ButtonClick`'s post-delete clear, and on a **failed search** (`SearchGrid`'s
  `Data_Module.ClearControls(ForecastDetail_Panel)` at `:388` runs before the `RecordCount` check,
  regardless of outcome), every other panel field blanks/resets but `Misc2PartNum_ComboBox` visibly
  retains its previous value — a concrete, reproducible field-clear divergence in the `#135` class,
  scoped to exactly one control. (`SetDetailBoxes`'s `SearchCombo(Misc2PartNum_ComboBox, ...)` still
  repopulates it correctly on a **successful** selection/search/Insert/Update, since `SearchCombo`
  is called directly on the control reference and doesn't care about VCL parentage — only the
  `ClearControls`-driven blank path is affected.)
- **On Insert** (`:477-500`): `Validate` (no-op, always `True`) → `HoldDetails(False)` → insert →
  on success, capture `assy`/`eff`, `GetAssy(ForecastPartsCode_ComboBox)` (refresh combo — may add
  the just-inserted assembly code), `GetForecastDetailInfo` (full requery),
  `Inv_DataSet.Locate('Assembly Part Number Code; Active Date', VarArrayOf([assy,eff]), [])` →
  **`SetDetailBoxes` runs unconditionally after the `if/else`, `:497`**, i.e. even on Insert
  **failure** (mirrors `AssyRatioMaster`'s finding, `AssyRatioMaster.md:110-117`) — on failure it
  repopulates from whatever `Data_Module` still holds (the rejected attempted values), not a
  requeried/reverted row.
- **On Update** (`:458-475`): `Validate` → `HoldDetails(False)` → `Id := RecordID` (captured before
  any requery) → `UpdateForecastDetailInfo` (no boolean return checked — always "succeeds" from the
  form's perspective) → `GetForecastDetailInfo` (requery) → `Inv_DataSet.Locate('ID', ID, [])` →
  `SetDetailBoxes` (echo).
- **On Delete** (`:444-456`): confirm (see Dialogs) → on `mrYes`: `DeleteForecastDetailInfo` +
  `GetForecastDetailInfo` (requery) → **unconditionally** (outside the `if`)
  `Data_Module.ClearControls(ForecastDetail_Panel)` (`:454`) — full blank of the panel's covered
  controls regardless of whether the delete was confirmed or declined. Per the two gaps above,
  `AssyQty_RadioGroup` and `Misc2PartNum_ComboBox` are **not** actually blanked by this call.
- **`FormCreate`** (`:116-153`): builds the rolling 12-month `EffectiveMonth_ComboBox` (`:120-138`,
  current month −1 to +10), opens `Inv_DataSet` unfiltered, binds `ForecastDetail_DataSource`,
  populates 7 part-type combos + the assembly-code combo via `GetAssy`. **Does not call
  `ClearControls`** — panel fields retain `.dfm` design-time defaults until the first
  `FormShow`/grid-navigation event fires `SetDetailBoxes`.
- **`FormShow`** (`:522-526`): unconditionally `SetDetailBoxes` then
  `ForecastPartsCode_ComboBox.SetFocus` — same shared-`Data_Module`-state hazard as `SizeMaster`
  (`SizeMaster.md:54-61`): the panel repopulates from whatever `Data_Module` fields were left by
  **any** previous form/screen in the session that touched the same module-level fields (`AssyCode`,
  `Quantity`/`fQTY`, `TireRatio1`, etc. — `fQTY` in particular is shared across **many** other
  screens, e.g. Shipping detail inserts/updates also read/write `fQTY`, `DataModule.pas:1397,3993,
  4040` among others — this is the P9 shared-mutable-field pattern, here extended to `Quantity` too,
  not just `RecordID`).
- **`ForecastPartsCode_ComboBoxSelect`** (`:528-561`): on picking an assembly code from the dropdown
  (only fires on a list selection, not free-typed text — the combo has no `Style` set, i.e. default
  `csDropDown`, editable), does a manual `FindFirst`/`FindNext` scan of the (still-filtered-or-not)
  dataset for a matching `Fields[1]` (Assembly code) value, `HoldDetails(True)` on match, then
  `SetDetailBoxes` regardless of whether a match was found. If the combo is reset to blank (`' '`,
  the sentinel first item), it instead unfilters and `ClearControls(ForecastDetail_Panel)`s — subject
  to the same two clear-gaps (radio group, Misc2 combo) noted above.
- **The `ForecastPartsCode_EditChange` handler is dead** (`:437-442`): its entire body is commented
  out (`// If Length(...) <= 1 Then // Data_Module.ClearControls(...)`), yet it remains wired as the
  `OnChange` handler for **four** unrelated edit fields — `ForecastRatio_Edit`, `Kanban_Edit`,
  `TireForecastRatio_Edit`, `WheelForecastRatio_Edit` (`ForecastDetail.dfm:293,303,313,323`) — a
  misleadingly-named leftover (likely copy/paste from a sibling form like `SizeMaster`'s live
  `SizeCode_EditChange`, `SizeMaster.md:26-31`) that fires on every keystroke in those four fields but
  does nothing. A rebuild should NOT infer any live-typing clear/validate behavior on these fields
  from the handler's name or wiring — there is none.

## Focus & keyboard
- No `ActiveControl` set anywhere in the `.dfm`; `FormShow` explicitly sets focus to
  `ForecastPartsCode_ComboBox` (`:525`), overriding the natural top-level tab order in which
  `ManagementButtons_Panel` (`TabOrder=0`, `.dfm:50`) would otherwise receive focus first.
- `Close_Button` has `ModalResult = 2` (`mrCancel`, `.dfm:93`) but **no `Default`/`Cancel` flag** on
  any button in this form — Enter/Esc are not VCL-wired to any button; only accelerators (`&Insert`,
  `&Update`, `&Delete`, `&Search`, `Cl&ear`, `&Close`) work as keyboard shortcuts.
- Grid `OnKeyUp`/`OnMouseUp` (`:422-435`) both call `HoldDetails(True)` + `SetDetailBoxes` — any
  click or keystroke (including arrow-key row navigation) on the grid re-syncs the detail panel.
  `ForecastDetail_DataSourceDataChange` (`:515-520`) mirrors this on any dataset navigation
  (e.g. the programmatic `Locate` calls from Insert/Update), same redundant-but-harmless double-call
  pattern as `SizeMaster` (`SizeMaster.md:74-78`).
- Combo `CharCase` is `ecUpperCase` on every combo/edit in the panel (`.dfm`, e.g. `:271,281,290,
  300,310,320,331,354,363,374,385,394,404,415`) — consistent forced-uppercase input across the form.
- `BroadcastCode_Edit` has `MaxLength = 20` (`.dfm:355`); `ForecastRatio_Edit`/`Kanban_Edit`/
  `TireForecastRatio_Edit`/`WheelForecastRatio_Edit` all have `MaxLength = 12` (`.dfm:291,301,311,
  321`) despite holding what are semantically small integers/ratios — no numeric-range enforcement
  via `MaxLength`, consistent with the "no real validation" finding above.
- Tab order within `ForecastDetail_Panel` runs `ForecastRatio_Edit`(0) → `EffectiveMonth_ComboBox`(1)
  → `Kanban_Edit`(2) → `TirePartNum_ComboBox`(3) → `WheelPartNum_ComboBox`(4) →
  `TireForecastRatio_Edit`(5) → `WheelForecastRatio_Edit`(6) → `AssyQty_RadioGroup`(7) →
  `BroadcastCode_Edit`(8) → `ForecastPartsCode_ComboBox`(9) → `ValvePartNum_ComboBox`(10) →
  `FilmPArtNum_ComboBox`(11) → `LabelPartNum_ComboBox`(12) → `Misc1PartNum_ComboBox`(13)
  (`.dfm:285-407`) — a non-obvious order that does NOT match the on-screen left-to-right/top-to-
  bottom visual layout (e.g. `ForecastPartsCode_ComboBox`, visually the top-left-most field, is tab
  stop 9, near the end). `Misc2PartNum_ComboBox` is form-level `TabOrder=3` (`.dfm:417`) which — since
  the panel itself is the form's tab stop 2 — still lands immediately after the panel's own tab chain
  is exhausted, so **despite the structural clear-gap above, the perceived tab flow still lands on
  Misc2 right after Misc1**, coincidentally matching visual expectation.

## Enable/disable state machine
- **None found.** No control's `Enabled` property is ever set in `ForecastDetail.pas` — Insert/
  Update/Delete/Search/Clear remain clickable regardless of grid selection state or whether any
  criteria have been entered, identical to `SizeMaster` (`SizeMaster.md:82-85`). The `AssyQty_
  RadioGroup` domain restriction ({1,2,4,5} on write, {0,2,4,5} on read) is enforced entirely by the
  fixed 4-item `.dfm` `Items.Strings` list (`'1','2','4','5'`, `:342-346`) — the operator physically
  cannot select a 5th value; there is no `Enabled`/visibility gating tied to any other field's state
  (unlike `AssyRatioMaster`'s tire/wheel-slot fill-order snap-back, which has none here either — this
  form has no analogous cascading-combo business logic; each part-type combo is independent).

## Error surfacing
- `Execute`'s outer `try/except` (`:101-109`) catches only exceptions raised **during `ShowModal`**
  (i.e., during the form's already-running event loop) → `showMessage('Unable to generate Forecast
  Detail screen.' + #13 + 'ERROR:' + #13 + E.Message)`.
- **`FormCreate` exceptions are NOT covered by `Execute`'s try/except** — `FormCreate` runs inside
  `TForecastDetail_Form.Create` (`MasterMaint.pas:140`), which executes **before** `Execute` is ever
  called (`MasterMaint.pas:141`). Of the 8 lookup calls in `FormCreate` (`:144-151`), only the
  assembly-code combo (`GetAssy`) has its own try/except (`:160-169`); the other 7
  `SelectDependantSingleField` calls (Tire/Wheel/Valve/Film/Label/Misc1/Misc2 part-type combos,
  `:144-150`) are unguarded. An exception there propagates out of `Create` straight to
  `MasterMaint.ForecastDetail_ButtonClick` (`MasterMaint.pas:137-145`), which — unlike its sibling
  handlers `SupMaster_ButtonClick`/`SizeMaster_ButtonClick` (`MasterMaint.pas:93-117`, both wrap
  `Hide`/`Show` in `try/finally`) — has **no try/finally at all**. Net: a `FormCreate`-time DB error
  here leaves `MasterMaint_Form` permanently `Hide`d (called at `:139` before `Create`) with no code
  path left to call `Show` again — the operator is stuck on a blank/closed main menu until the
  unhandled exception reaches Delphi's default `Application.HandleException` (generic error box) and,
  even then, `MasterMaint_Form.Show` is never re-invoked. This is a distinct, form-specific hazard not
  present in `SizeMaster`'s or `AssyRatioMaster`'s entry points.
- `SearchGrid`'s `try/except` (`:397-400`) and `GetAssy`'s (`:164-169`) surface via `ShowMessage`/
  `MessageDlg` as noted in Dialogs. `Insert_ButtonClick`/`Update_ButtonClick`/`Delete_ButtonClick`
  have no try/except of their own — an exception inside `InsertForecastDetailInfo`/
  `UpdateForecastDetailInfo`/`DeleteForecastDetailInfo` would normally propagate unguarded, **but**
  note those three `Data_Module` procs each carry their **own** internal try/except with
  `ShowMessage` + `LogActLog` + a retry-on-error loop (e.g. `UpdateForecastDetailInfo`,
  `DataModule.pas:2426-2434`, retries itself up to 3 times) — so in practice DB-layer errors here are
  caught and logged one level down, inside the DataModule, not by the form.

## Cross-refs
- **`docs/analysis/forecasting/forecast-detail.md`** — the authoritative proc/data/trigger spec for
  `INV_FORECAST_DETAIL_INF` (SELECT/INSERT/UPDATE/DELETE proc bodies, the `DeleteForecastDetail`
  cascade trigger into `INV_FORECAST_INF`, the P9/P12 cross-cutting notes, and the open
  snapshot-vs-live param-count question). That doc's §4 already notes the write-side
  `AssyQty_RadioGroup` mapping (`:59`, "Assy qty radio maps to 1/2/4/5") and the tire/wheel 0–100
  clamp with the forecast-ratio clamp intentionally disabled (`:57-58`) — **this file adds the
  read-side `{0,2,4,5}`-domain mapping, the resulting read⇄write asymmetry/round-trip bug, and the
  two independent field-clear gaps (radio group, Misc2 combo) that a pure proc/data reading would
  not surface.**
- That same spec's §1 flags an open **Q4** ("confirm which menu item opens the editable
  `TForecastDetail_Form`", `forecast-detail.md:141-142`) — **resolved here**:
  `MasterMaint.ForecastDetail_ButtonClick` (`MasterMaint.pas:137-143`), button never hidden, is the
  live, sole entry point.
- **Correction to the original sweep:** issue #141 originated from a mis-mapping of this rebuild
  screen (Assembly Detail) to the DEAD `AssyRatioMaster.pas` — that form's constructing button
  (`AssyRatioMaster_Button`) is unconditionally hidden (`MasterMaint.pas:78`) and is unreachable in
  production (`docs/analysis/cross-cutting/form-ux/AssyRatioMaster.md`, confirmed against
  `docs/analysis/assembly/assy-ratio-master.md`). The correct live source for the "Assembly Detail
  Master" screen (same on-screen caption, `ForecastDetail.dfm:6` vs `AssyRatioMaster.dfm`'s own
  caption) is **this file's subject, `ForecastDetail.pas`/`.dfm`**, over `INV_FORECAST_DETAIL_INF`,
  not `AssyRatioMaster`'s broadcast-code/ratio-explosion table. Any rebuild spec, test oracle, or
  code comment that cites `AssyRatioMaster` for the AssemblyDetail/AssyQty/forecast-ratio screen
  should be corrected to cite `ForecastDetail` instead.
