# Form-UX: `TRecRej_Form` — `RecReject.pas` + `RecReject.dfm`

> Reused for both "Receiving Reject" and "Production Reject" (division set by caller's button
> `Tag`). Confidence: high — both files read in full (394 + 278 lines).
> Business/proc spec: [`../../receiving/recreject.md`](../../receiving/recreject.md).

## Dialogs & confirmations
- **`Delete_ButtonClick`** (`:324-339`) — **armed confirm:**
  `MessageDlg('Are you sure you wish to delete'+#13+'supplier code '+SupplierCode+#13+'and part
  code '+PartNum+#13+' from the database?', mtWarning, [mbYes, mbNo], 0) = mrYes` → only then
  `DeleteRecProdRejInfo`. Same `mtWarning`-not-`mtConfirmation` icon choice as `RecConfStat`'s
  delete confirm.
- **`Insert_ButtonClick`** (`:301-309`) — no confirmation; on `InsertRecProdRejInfo` returning
  `False`, a bare `ShowMessage('Unable to INSERT Reject Data')` (`:305`) — **the least detailed
  failure message in this family** (no supplier/part/qty echoed back, unlike `RecConfStat`'s
  equivalent).
- **`Update_ButtonClick`** (`:311-322`) — no confirmation, no failure-path check at all: the return
  value of `Data_Module.UpdateRecProdRejInfo` is **not** tested (`:317` — the call's result is
  discarded), so a failed update shows **no dialog whatsoever**; the operator only notices via the
  grid/detail panel not reflecting the intended change.
- **`HoldDetails(False)`'s blank-part guard is silently non-blocking** (`:213-217`): sets
  `fErrMsg:='Part Number must not be blank'` but **`Insert_ButtonClick` never inspects `fErrMsg`**
  (`:303` calls `HoldDetails(False)` and discards the returned string) — so this "error message"
  never actually reaches a dialog or blocks the insert; the only real gate is the DB's `IN_PART_ID
  NOT NULL` constraint failing (per `recreject.md` §4). **This confirms the same "built error text,
  never surfaced" pattern flagged in the business spec — from the UX side, there is truly zero
  operator-visible feedback for a blank part on Insert.**
- **No confirmation on `Search_ButtonClick`/`Clear_ButtonClick`** — pure filter actions, no dialog.
- **Top-level `Execute`'s catch** (`:96-101`): `ShowMessage('Unable to generate Receiving Reject
  screen.'+#13+'ERROR:'+#13+E.Message)`.

## Field clear / repopulate
- **`SetDetailBoxes`** (`:228-256`) — wrapped in its own `try…except` that **swallows all errors
  silently** (`:254-255` — bare `except end`, no log, no message) — any exception while
  reformatting a date or searching a combo is invisible to the operator. Reformats `EditDate`
  (8-char `yyyymmdd`) into the date picker only via `TryStrToDate` (`:238` — if the string doesn't
  parse, the picker **keeps its prior value**, no explicit clear); sets
  `DiscnDiv_RadioGroup.ItemIndex` from `Division` (blank `Division` → index 0, i.e. "Receiving");
  `SearchMultiCombo`/`SearchCombo` for supplier/part (see Cross-refs for the shared blank-placeholder
  convention); `RejQty_MaskEdit.Text:=IntToStr(Quantity)`; `Reason_Memo.Text:=Comments`.
- **⚠️ `FormShow` resets the division radio group AFTER `Execute` already set it — a real
  divergence hazard.** `Execute` (`:80-106`) sets `DiscnDiv_RadioGroup.ItemIndex` from
  `Data_Module.Division` **before** calling `ShowModal` (`:88-94`); but `ShowModal` triggers
  `FormShow` (`:367-375`) as part of making the form visible, and `FormShow` **unconditionally**
  does `DiscnDiv_RadioGroup.Items.Clear` + re-`Add`s the three division labels + **`ItemIndex:=0`**
  (`:370-374`) — **overwriting whatever division-appropriate index `Execute` just set.** Net effect:
  opening this form for "Production Reject" (division 2 or 3, launched via `MainMenu.pas:292`'s
  `Tag`-based dispatch) **visually resets to "Receiving Reject" (index 0) on open**, even though the
  window `Caption` correctly still says "Production Reject" (`Execute:92`, set before `ShowModal`
  and not touched by `FormShow`). **Confirm with the domain expert whether this is a known/accepted
  quirk or a genuine bug** — it means the division radio and the window caption can disagree at
  first paint for the Production Reject entry point.
- **`Clear_ButtonClick`** (`:349-355`): `Supplier_NUMMIColumnComboBox.ItemIndex:=0` (first item,
  the placeholder — this control uses `SelectMultiField`, a two-column variant of the same
  blank-placeholder convention), `ClearControls(RecProdRej_Panel)` (shared DataModule walker),
  `SetTodaysDate(RecProdRej_Panel)` (resets the date picker to **today**, not blank — divergence
  from `RecConfStat`'s milestone-date fields, which go to true-empty on clear).
- **`FormCreate`** (`:154-169`) — `fNoChange:=True` guards the initial `SelectMultiField` combo
  population from re-firing `RecRej_DataSourceDataChange`; `SetTodaysDate` + `ClearControls` on the
  detail panel before the grid's `DataSet` is even wired, so the form opens with a blank detail
  panel regardless of what the grid shows.
- **`HoldDetails(True)` (from-grid capture, `:177-198`)** re-derives `Division` from the **grid's
  displayed text** (`'Receiving'`/`'Assembler'`/`'Plant'` → `'1'`/`'2'`/`'3'`, defaulting to `'1'`
  for anything else, `:182-189`) rather than a raw code column — a string-match dependency on the
  proc's `CASE`-derived display label (`SELECT_RecProdRejInfo`, per the business spec) that would
  break silently if that label text ever changed.

## Focus & keyboard
- **Initial focus (`FormShow`, `:367`):** `Supplier_NUMMIColumnComboBox.SetFocus` — called **before**
  the division-radio reset later in the same handler (order: SetFocus first, then Items.Clear/Add/
  ItemIndex, `:367-374`) — the radio-group reset does not re-steal focus since it's not a focus-
  taking operation itself.
- **Post-Insert/-Update/-Delete/-Search focus:** `Supplier_NUMMIColumnComboBox.SetFocus` in **all
  four** (`:308`, `:321`, `:338`, `:345`) — every action returns focus to the top of the entry chain.
- **`Supplier_NUMMIColumnComboBoxChange`** (`:377-387`) — selecting a supplier repopulates the
  dependent Parts combo (`SELECT_DependantPartNumber_Supplier`) and **`SetFocus`es
  `PartsCode_ComboBox`** (`:386`) — cascading-combo auto-advance.
- **`PartsCode_ComboBoxChange`** (`:389-392`) — selecting a part **`SetFocus`es `RejQty_MaskEdit`**
  (`:391`) — continues the auto-advance chain: Supplier → Part → Qty, keyboard/scanner-friendly.
- **No `Default`/`Cancel` flags, no `KeyPreview`** anywhere in `RecReject.dfm` (grep-verified) —
  Enter does not trigger Insert/Update/Delete/Search.
- **`TextChange`/`MaskEditExit`** (`:122-146`) — the shared mask-edit hygiene pair used across this
  whole form family: `TextChange` forces an emptied `TMaskEdit` back to `'0'` on every keystroke;
  `MaskEditExit` strips embedded spaces from the mask's fixed-width padding on focus-loss. Applies to
  `RejQty_MaskEdit` only on this form.

## Enable/disable state machine
- **No button ever disables based on state** on this form — Insert/Update/Search/Clear/Delete are
  always clickable regardless of whether a row is selected, unlike `ModifyShipping`'s Insert↔Update
  visibility swap. The only "state" distinguishing an edit-in-progress from a fresh entry is which
  fields `HoldDetails` most recently captured (grid vs. manual panel) — invisible to the operator.
- **`RecProdRej_DBGrid`'s `dgConfirmDelete`** (`RecReject.dfm:41`) is present at the component-option
  level but, as elsewhere in this family, is likely inert given the app's own explicit `Delete_
  Button` + confirm dialog is the real delete path — flag as unverified rather than assume a working
  second delete route exists via the grid's Delete key.

## Error surfacing
- **Update's silent failure is the standout gap**: `Update_ButtonClick` does not check
  `UpdateRecProdRejInfo`'s return value at all — no dialog, no log call visible in this unit — a
  failed update is completely invisible to the operator.
- **`SetDetailBoxes`'s blanket `except end`** (no log, no message) is the second-worst gap — any
  reformatting/combo-search exception during row-load is silent.
- Everything else surfaces via plain `ShowMessage`/`MessageDlg`; no inline field-error text.

## Cross-refs
- Business rules / procs / triggers (division-as-label-only, part-immutable-on-update, keying on
  `IN_PART_ID`, no purge bypass on delete):
  [`../../receiving/recreject.md`](../../receiving/recreject.md).
- Shared DataModule field-clear/combo-populate ancestor: `DataModule.pas:5976` `ClearControls`,
  `:5916` `SearchMultiCombo`, `:5935` `SearchCombo` — see `RecConfStat.md` Cross-refs for the
  blank-placeholder (`' '` literal, `ItemIndex:=0`) convention shared by every combo in this family.
