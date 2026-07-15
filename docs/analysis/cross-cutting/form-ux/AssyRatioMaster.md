# Form-UX semantics: `TAssyRatioMaster_Form` — `AssyRatioMaster.pas` / `AssyRatioMaster.dfm`

CRUD editor over the broadcast-code → tire/wheel ratio-explosion table (2-char broadcast code, assy
code/name, tire qty + up to 3 tire part/ratio pairs, wheel qty + up to 3 wheel part/ratio pairs,
spare tire/wheel qty+parts). Grid `ASSYRatioMaster_DBGrid` bound via `AssyRatio_DataSource`.

**Reachability finding (load-bearing — read this before anything else in this file):** the ONLY
constructor of this form is `MasterMaint.AssyRatioMaster_ButtonClick`
(`MasterMaint.pas:119-126`), behind `AssyRatioMaster_Button`. `MasterMaint.Execute` sets
`AssyRatioMaster_Button.Visible := FALSE // not used yet` **unconditionally**
(`MasterMaint.pas:78`, confirmed by an independent second hide inside the `GenerateEDI` branch at
`:74`) — **the button is permanently hidden and this form is unreachable from the shipping menu**,
even though it is fully compiled (`InventorySystem.dpr:17`) and functionally complete. This matches
the confirmed finding already on record in `docs/analysis/assembly/assy-ratio-master.md` §1.1. The
UX semantics below describe fully-functional, compiled, LIVE CODE that no operator can currently
reach through the standard UI — treat this as the spec for what the data-maintenance behavior WOULD
be if re-exposed, not as a currently-operating screen.

## Dialogs & confirmations
- **Insert failure** — `AssyRatioMaster.pas:168-169`: `MessageDlg('Unable to INSERT ' +
  Data_Module.AssyCode, mtInformation, [mbOk], 0)` if `InsertAssyRatioInfo` fails.
- **Delete confirmation (armed, two-step)** — `:206-208`: `MessageDlg('Are you sure you wish to
  delete' + #13 + Data_Module.BroadcastCode + ' from the database?', mtWarning, [mbYes, mbNo], 0) =
  mrYes` gates `DeleteAssyRatioInfo`.
- **`Verify` (`:122-139`) — a ratio-sum gate, NOT a per-field type check** (contrast with
  `PartsStockMaster`/`RenbanGroupMaster`'s numeric-parse validators): checks
  `StrToInt(TotalTireRatio_Edit.Text) <> 100` → `ShowMessage('Tire ratio must total to 100')` +
  `TireRatio1_MaskEdit.SetFocus` + `result := FALSE`, and independently
  `StrToInt(TotalWheelRatio_Edit.Text) <> 100` → `ShowMessage('Wheel ratio must total to 100')` +
  `WheelRatio1_MaskEdit.SetFocus` + `result := FALSE` — **both checks always run (no `exit` between
  them)**, so if BOTH totals are wrong, BOTH ShowMessage dialogs fire in sequence (the operator must
  dismiss two stacked dialogs), and focus ends up on `WheelRatio1_MaskEdit` (the second check's
  `SetFocus` wins, overwriting the first's). Both `Insert_ButtonClick` (`:163`) and
  `Update_ButtonClick` (`:187`) call `Verify` first and skip the whole operation if it returns
  `False`.
  - `TotalTireRatio_Edit`/`TotalWheelRatio_Edit` are `ReadOnly = True` display-only fields
    (`AssyRatioMaster.dfm:537,548`) computed live by the ratio-`MaskEdit` `OnChange` handlers (see
    below) — the operator never types into them directly; `StrToInt` on them will only fail to
    parse if the live-sum computation itself produced a non-numeric string, which the source does
    not appear to allow (always builds via `IntToStr`), so a `Verify`-time parse EXCEPTION (as
    opposed to a wrong-total ShowMessage) is unlikely but unguarded (`StrToInt`, not `TryStrToInt`,
    at `:126,133`).
- **Empty-search guard** — `:221-222`: if `Trim(BroadcastCode_Edit.Text) = ''`,
  `ShowMessage('Please enter a broadcast code before searching.')`, search skipped — same
  pre-condition pattern as `ManifestCostMaster.Search_ButtonClick`.
- **Search "not found"** — `:232`: `ShowMessage('No matches were found for your query.')`.
- **Search error** — `:283-284`: `ShowMessage('Error in Search' + #13 + e.Message)`.
- **Combo-population failure** — `GetParts` (`:458-478`), identical shape to
  `ManifestCostMaster.GetParts`: `MessageDlg('Unable to get a list of parts.', mtError, [mbOK], 0)`.
  Called SEVEN times in `FormCreate` (`:441-447`) — three tire-part combos, three wheel-part combos,
  one assy-code combo — any one of the seven can independently pop this dialog if its
  `SelectSingleField` call raises.
- Grid carries `dgConfirmDelete` (`AssyRatioMaster.dfm:560`) — VCL-native secondary confirm.

## Field clear / repopulate
- **Cascading combo-reset chain on tire/wheel part-number change** — the busiest inter-field
  cascade logic in the whole master-data family (`:516-608`):
  - `TirePartNum1_ComboBoxChange` (`:550-561`): if slot 1 is set to blank (`ItemIndex = 0`), forces
    slot 3 blank + `TireRatio3 := '0'`, slot 2 blank + `TireRatio2 := '0'`, AND `TireRatio1 := '0'` —
    i.e. clearing the FIRST tire part cascades to blank/zero ALL THREE tire slots.
  - `TirePartNum2_ComboBoxChange` (`:532-548`): if slot 1 is still blank, forces slot 2 back to
    blank and `SetFocus`s slot 1 (**prevents filling slot 2 before slot 1** — an implicit
    fill-in-order requirement enforced by combo `OnChange`, not by any `Enabled` gate) and `exit`s;
    otherwise, if slot 2 is itself blank, cascades to blank slot 3 + zero ratios 2 and 3.
  - `TirePartNum3_ComboBoxChange` (`:516-530`): same fill-in-order guard against slot 2 being blank
    (forces slot 3 blank, focuses slot 2, exits); otherwise if slot 3 is blank, zeroes ratio 3 only.
  - The wheel-part combos (`WheelPartNum1/2/3_ComboBoxChange`, `:563-608`) mirror this EXACT
    cascade/fill-order-guard structure for the wheel side.
  - **Net effect: a rebuild must reproduce a strict left-to-right fill-order enforcement (tire slot 2
    cannot be set while slot 1 is blank; slot 3 cannot be set while slot 2 is blank; same for wheel)
    plus a cascading blank/zero-out when an earlier slot is cleared** — this is real business logic
    living in form-code `OnChange` handlers, not in a stored proc; a straightforward proc-wrapping
    rebuild would MISS this entirely since it's client-side-only in the legacy app.
  - `TireRatio1/2/3_MaskEditChange` and `WheelRatio1/2/3_MaskEditChange` (`:610-768`, six near-
    identical handlers): (a) if the corresponding part-number combo is blank, force this ratio's
    text to `'0'`; (b) if the ratio's own text is empty, force it to `'0'` too; (c) recompute
    `TotalTireRatio_Edit`/`TotalWheelRatio_Edit` as the running sum of ratio1 (+ratio2 if non-empty
    (+ratio3 if non-empty)) — wrapped in a `try/except` with an **empty except block** (`:632-633`
    etc., silent on any `StrToInt` parse failure during the running-sum computation).
  - `TotalTireRatio_EditChange` (`:508-514`): sets font color black if the (read-only, programmatically
    updated) total is exactly `100` or `0`, else **red** — a live, non-blocking visual cue (not a
    dialog) that the ratio doesn't sum to 100 yet, distinct from the blocking `Verify` check that
    runs at Insert/Update time. `TotalWheelRatio_Edit`'s `OnChange` is ALSO wired to
    `TotalTireRatio_EditChange` (`AssyRatioMaster.dfm:551` — reuses the SAME handler for both total
    fields, which works because the handler operates on `TEdit(Sender)` generically), so the wheel
    total gets the identical black/red visual rule.
- **`SetDetailBoxes`** (`:288-333`): decodes `TireQty`/`WheelQty` (0/4/5 and 1/4/5 respectively) into
  `RadioGroup.ItemIndex` via `case` with an `Else` defaulting to index 0 (`:298-304`,`:321-327`) —
  **any TireQty/WheelQty value other than the three named ones silently maps to the FIRST radio
  option** rather than erroring (e.g. a legacy/dirty value of `2` or `3` in the DB would silently
  display as if TireQty were its `0`-mapped choice).
- **`FormCreate`** (`:435-456`): loads 7 combos via `GetParts` (see above), sets
  `AssyCode_ComboBox.Text := ''`, `ClearControls(AssyRatioMaster_Panel)`, THEN manually re-zeroes all
  6 ratio mask-edits to `'0'` individually (`:450-455`) — same manual-reset-on-top-of-ClearControls
  pattern as `PartsStockMaster`/`RenbanGroupMaster`.
- **`FormShow`** (`:502-505`): `BroadcastCode_Edit.SetFocus` ONLY — **no `SetDetailBoxes` call, no
  dataset `.First` call either** (unlike `ManifestCostMaster`/`RenbanGroupMaster`, which at least
  call `Inv_DataSet.First` to trigger the `OnDataChange` cascade). This form's panel is whatever
  `FormCreate` last left it (blanked, per above) UNLESS the grid already has a current record from
  binding `AssyRatio_DataSource.DataSet` in `FormCreate` (`:440`) — **if binding a `TDataSource` to
  an already-positioned `TDataSet` fires `OnDataChange` as a side effect of the assignment itself**,
  the panel could populate from the FIRST grid row before `FormShow` even runs (Delphi/VCL commonly
  fires `OnDataChange` on `DataSet :=` if the dataset is already active and positioned) — **this
  specific ordering/side-effect claim is [UNVERIFIED — confirm before use]**; it depends on ADO/VCL
  dataset-binding semantics not confirmed by reading this file alone.
- **`SearchGrid`** (`:254-286`): unconditionally re-zeroes all 6 ratio mask-edits AND calls
  `ClearControls(AssyRatioMaster_Panel)` at entry (`:257-263`), before attempting the filter — so a
  failed search (no match) leaves a genuinely blank panel, same clean-failure behavior as
  `LogisticsMaster.SearchGrid`.
- **On Insert** (`:159-181`): `Verify` → `HoldDetails(False)` → insert → (on success) capture
  `broad := BroadcastCode`, requery `GetAssyRatioInfo`, re-unfilter, `Locate('Broadcast Code', broad,
  [])` → `SetDetailBoxes` (echoes inserted row, called OUTSIDE the `with Data_Module` block,
  `:178`, so it runs regardless of whether the insert branch or the `else` failure branch executed —
  **on Insert FAILURE, `SetDetailBoxes` still runs and repopulates from whatever `Data_Module`
  currently holds**, which at that point is the attempted-but-rejected new values, not a requeried
  row — the panel shows the rejected insert's data as if it were live, not blanked or reverted) →
  `BroadcastCode_Edit.SetFocus`.
- **On Update** (`:183-201`): same requery/`Locate`/`SetDetailBoxes` pattern, always runs regardless
  since there's no failure branch for Update (`UpdateAssyRatioInfo` has no boolean return checked
  here, unlike Insert's `InsertAssyRatioInfo`).
- **On Delete** (`:203-215`): `GetAssyRatioInfo` (requery) + `ClearControls(AssyRatioMaster_Panel)` —
  full blank (this form DOES clear on delete). Note: does **not** re-zero the 6 ratio mask-edits
  individually here (unlike `Clear_Button`/`SearchGrid`/`FormCreate`) — relies solely on
  `ClearControls` for this path, so whether the ratio fields actually end up at `'0'` vs blank after
  a delete depends entirely on what `ClearControls` does to `TMaskEdit` controls specifically
  ([UNVERIFIED — confirm against `DataModule.ClearControls`'s actual per-control-type behavior]).
- **`Clear_Button`** (`:238-252`): unfilters, `ClearControls(...)`, manually re-zeroes all 6 ratio
  mask-edits, focus to `BroadcastCode_Edit`.

## Focus & keyboard
- `FormShow` focuses `BroadcastCode_Edit` only (`:504`) — the simplest `FormShow` in the family (no
  repopulate call at all).
- No VCL `Default`/`Cancel` flags; `Close_Button` has `ModalResult = 2` only
  (`AssyRatioMaster.dfm:84`). Accelerators only.
- Grid `OnKeyUp`/`OnMouseUp` (`:480-493`) both call `HoldDetails(True)` + `SetDetailBoxes`;
  `AssyRatio_DataSourceDataChange` (`:495-500`) mirrors on dataset navigation.
- `SpareTirePartsCode_Label`/`_Edit`, `SpareTireQty_Label`/`_RadioGroup`,
  `SpareWheelPartsCode_Label`/`_Edit` are all `Visible = False` at design time
  (`AssyRatioMaster.dfm:176-177,230-231,241-242,316-317,326-327,362-363`) yet remain in the tab
  chain (`TabOrder = 17,18,19` respectively, `:316,326,362`) with live read/write code in
  `HoldDetails`/`SetDetailBoxes` (`SpareTireQty`, `SpareTirePartNum`, `SpareWheelPartNum`,
  `:328-331,428-430`) — **hidden-but-functional fields**, same class of issue as
  `PartsStockMaster`'s hidden `VendorShare_Edit`, but here there are THREE hidden controls plus
  their labels, all still wired into the read/write data path. A rebuild must decide whether to
  surface these (they clearly carry live spare-tire/wheel data per the DB proc calls) or confirm
  they're genuinely unused before dropping them.

## Enable/disable state machine
- **None via `.Enabled` found** — the fill-order enforcement described above (tire/wheel slot 2
  cannot be set while slot 1 is blank, etc.) is implemented via `OnChange` snap-back
  (forcing `ItemIndex := 0` + `SetFocus` + `exit`) rather than via disabling the later combo boxes —
  a rebuild could reasonably choose to implement this as true `Enabled` gating instead, which would
  be a UX IMPROVEMENT but is a deliberate divergence from the legacy interaction (the legacy combo
  is always clickable/open-able; it just snaps back after the fact) — flag per the project's
  divergence rule if changed.

## Error surfacing
- `SearchGrid` and `GetParts` both wrap in `try/except` → `ShowMessage`/`MessageDlg`. The six
  ratio-`MaskEditChange` handlers wrap their running-sum computation in `try/except` with an EMPTY
  except (silent). `Insert_/Update_/Delete_ButtonClick` have no try/except of their own.

## Cross-refs
- `docs/analysis/assembly/assy-ratio-master.md` — the authoritative proc/data spec AND the source
  of the confirmed dead-button/unreachable-form finding echoed at the top of this file. That doc
  also covers the dead-code stub sibling `BCRatioMaster.pas` (never compiled into the product,
  `NOT in InventorySystem.dpr`) — not in scope for this form-UX sweep since it never ships.
