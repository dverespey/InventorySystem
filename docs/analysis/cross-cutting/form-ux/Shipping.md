# Form-UX: `TShipping_Form` — `Shipping.pas` + `Shipping.dfm`

> GALC-sequence shipping post. Confidence: high — both `.pas`/`.dfm` read in full.
> Business/proc spec: [`../../shipping/shipping.md`](../../shipping/shipping.md).

## Dialogs & confirmations
- **No confirmation dialogs at all for Insert/Post.** `Insert_ButtonClick` (`Shipping.pas:140-162`)
  posts directly — no "are you sure" before subtracting stock. Only two outcome `ShowMessage`s:
  - Success: `'Shipping info update is complete.'` (`:149`)
  - Failure: `'Failed to update shipping info.'` (`:147`)
  - Gate message (not a confirm, a block): `'Please run check first'` (`:160`) if `fCheck=False`.
- **`Check_ButtonClick`** (`:290-312`) has no dialog at all — silently fills qty/cont-no from the
  ALC preview and sets `fCheck:=true`.
- **Sequence-lookup failures** are plain info dialogs, not confirmations: `'Sequence not found'`
  (`:404`), `'Unable to access sequence number date/time, '+e.Message` (`:409`).
- **`StartSeqNo_EditChange`** auto-advances the End sequence combo and, if the chosen start is after
  every candidate end time, shows `'Start date is after end date, changing to previous date for
  starting sequence.'` (`:397`) then **silently** moves `StartBox.ItemIndex` back one and re-derives —
  no operator confirmation, an automatic correction.
- **`UpdateShipping_ButtonClick`** (`:433-448`) — opening the amend sub-form (`ModifyShipping`) has
  no confirm; only a catch-all error dialog `'Unable to update shipping, '+e.Message` (`:444`).
- **`Refill_ButtonClick`** (`:334-351`) — **dead/unreachable**: the handler exists but no button is
  declared in `Shipping.dfm` (confirmed — no `Refill_Button` object). If ever wired, it re-runs
  `CalculateFRS` with no confirmation either.
- **No delete anywhere on this form** — there is no delete-shipment action here at all (see
  `shipping.md` §5: the only restore path is `ModifyShipping`'s underlying `DeleteShipDate`
  trigger cascade, not exposed on any button in either form).

## Field clear / repopulate
- **`SetDetailBoxes`** (`:186-262`) is the single field-repopulate routine, called on line change,
  date change, and post-Insert. Branches on whether `ProductionDate` already has a shipment:
  - **Already-shipped branch** (`:199-225`): fills Start/Last seq edit + qty/continuation from the
    DB row, **and** rebuilds `StartBox`/`EndBox` as **single-item** combos (`Items.Clear` then one
    `Items.Add` + `ItemIndex:=0`) from the stored `DT_START_SEQ_NUMBER`/`DT_END_SEQ_NUMBER` — the
    picklist collapses to just the recorded value. Sets `StartSeqNo_Edit.ReadOnly:=True` etc.
    (see Enable/disable below).
  - **Not-yet-shipped branch** (`:227-260`): if a prior end-seq exists, computes `next = end+1`
    (wraps `>999 → 0`) into `StartSeqNo_Edit`; **else blanks `StartSeqNo_Edit.Text:=''`** and clears
    `StartBox.Items`. **Always** blanks `LastSeqNo_Edit`, `ShipQty_MaskEdit`, `ContNo_MaskEdit` to
    `''` and clears `EndBox.Items` — **no stale values survive a date/line change.**
- **`TextChange`** (`:264-269`) — any `TMaskEdit` that goes empty (`Length(Trim(Text))<1`) is forced
  back to `'0'` — the two authentication mask edits (`ShipQty_MaskEdit`, `ContNo_MaskEdit`) can
  never actually be blank on screen; an empty entry silently becomes the literal `'0'`.
- **`Insert_ButtonClick`** post-success: does **not** clear the panel (the `ClearControls` calls are
  commented out, `:151-152`) — instead calls `SetDetailBoxes` which re-fetches and re-locks the
  now-already-shipped day (see above). `StartSeqNo_Edit.SelectAll` (`:154`) highlights the (now
  read-only, locked) start-seq text.
- **Empty-value class (#135-adjacent):** `StartBox`/`EndBox` are `TNUMMIComboBox`, not the
  DataModule combo-populate helpers — they're populated ad hoc here and via
  `StartSeqNo_EditChange`, so the "blank → literal space item" DataModule convention
  (`SelectSingleField`, `Shipping.pas` doesn't call it directly) does **not** apply to these two;
  they go fully empty (`Items.Clear`) with no placeholder — verify a rebuild list component treats
  an empty options list as "no selection," not as "select first."

## Focus & keyboard
- **Initial focus (`FormShow`, `:415-421`):** `StartSeqNo_Edit` if blank, else `LastSeqNo_Edit`.
- **Post-Insert focus:** `StartSeqNo_Edit.SetFocus` (`:156`), guarded by `if Shipping_Form.Visible`
  (defends against a `SetFocus` on a form mid-hide — a Delphi VCL gotcha, not app logic).
- **`StartSeqNo_EditChange`** (`:358-413`) fires **on every keystroke** (`OnChange`, not `OnExit`)
  but only *acts* once `Length(Text) = MaxLength` (3 chars) — then looks up the ALC sequence
  date/time and **auto-moves focus to `LastSeqNo_Edit`** (`:380`) once the Start box is resolved
  (scanner-friendly: type 3 digits, Enter/next-char isn't needed, focus jumps automatically).
- **No `KeyPreview`, no Enter-as-tab, no Default/Cancel button flags anywhere in `Shipping.dfm`**
  (grep-verified: zero `Default = True` / `Cancel = True` objects) — Enter does not trigger
  Insert/Check; only mouse-click (or Alt+accelerator, e.g. `&Insert`... actually `Insert_Button`
  caption is `'&Update Inventory'`, `Check_Button` is `'C&heck'`) invokes them.
- **`Line_ComboBoxChange`** (`:424-431`) and **`Production_DateTimePickerChange`** (`:314-332`, guarded
  by `fNoDateTimeUpdate` re-entrancy flag) both trigger a full `SetDetailBoxes` refresh — no
  confirmation of losing in-progress (unsaved) sequence entry.

## Enable/disable state machine
- **Already-shipped (locked) state** (`SetDetailBoxes:216-222`): `StartSeqNo_Edit.ReadOnly:=True`,
  `LastSeqNo_Edit.ReadOnly:=True`, `StartBox.ReadOnly:=True`, `EndBox.ReadOnly:=True`,
  `Insert_Button.Enabled:=False`, `Check_Button.Enabled:=False`,
  `UpdateShipping_Button.Visible:=True` (only surfaces in this state).
- **Not-yet-shipped (open) state** (`:253-259`): all four `ReadOnly:=False`,
  `Insert_Button.Enabled:=True`, `Check_Button.Enabled:=True`, `UpdateShipping_Button.Visible:=False`.
- **`fCheck` gate:** `Insert_Button` is always visually `Enabled` in the open state, but
  `Insert_ButtonClick` (`:142`) refuses (dialog, not disable) unless `fCheck=True` — so the button
  is enabled-but-blocked until Check has run. `fCheck` resets to `False` on
  `Production_DateTimePickerChange` (`:329`) forcing a re-Check after any date edit.
- **`UpdateShipping_Button`** starts `Visible:=False` in the `.dfm` (`Shipping.dfm:253`) and is only
  ever shown by `SetDetailBoxes`'s locked branch — never hidden again except by re-entering the
  open branch on a date/line change.

## Error surfacing
- **All errors are `ShowMessage`/`showMessage` dialogs**, not inline status labels — includes the
  top-level `Execute` catch-all (`'Error on get information shipping screen.'+#13+'ERROR:'+#13+
  E.Message`, `:130-132`) and every per-action catch. No field-level red-highlight or status bar.
- **`CalculateFRS`/`InsertShippingInfo` failures** surface as the single generic
  `'Failed to update shipping info.'` — the operator gets no detail on *which* part/broadcast-code
  failed (see `shipping.md` §4 for the underlying missing-T/W-slot hard-fail this masks).

## Cross-refs
- Business rules / procs / triggers: [`../../shipping/shipping.md`](../../shipping/shipping.md)
  (§4 "already processed" lock, `fCheck` gate, `CalculateFRS` explosion).
- Shared field-clear/combo ancestor conventions used elsewhere in this family (not by this form
  directly — `Shipping` builds its own combos): `DataModule.pas:5767` `SelectSingleField` /
  `:5976` `ClearControls` (see `RecConfStat.md`/`Stocktaking.md` for the forms that do use them).
