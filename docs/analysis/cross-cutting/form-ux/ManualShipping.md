# Form-UX: `TManualShipping_Form` — `ManualShipping.pas` + `ManualShipping.dfm`

> No-ALC manual daily-build-count shipping post. Confidence: high — both `.pas`/`.dfm` read in full.
> Business/proc spec: [`../../shipping/shipping.md`](../../shipping/shipping.md) (M1/M3 signature
> mismatches — this form's Post/IrregularShip paths call procs with the wrong param names).

## Dialogs & confirmations
- **`FormCloseQuery`** (`:462-472`) — **the one real "unsaved changes" confirm in this family:**
  if `fChanged` (an Update was made to the grid since last post), closing prompts
  `MessageDlg('Close and cancel changes?', mtConfirmation, [mbYes, mbNo], 0)`; **`mrNo` blocks the
  close** (`CanClose:=FALSE`). Default button of a `MessageDlg([mbYes,mbNo],0)` in VCL is **Yes**
  (the `0`-index button gets the default focus) — so Enter/click-through defaults to closing.
- **`Post_ButtonClick`** (`:303-378`) has **no confirmation** before posting (unlike the Close-guard
  above, committing the day's shipment is a single click with no "are you sure"). Info dialogs only:
  - `'Starting sequence invalid'` (`:313`), `'Last sequence invalid'` (`:321`) — validation blocks,
    then `SetFocus` back to the offending edit and **blanks it** (`StartSeqNo_Edit.Text:=''` /
    `LastSeqNo_Edit.Text:=''`, `:314`/`:322`) before the operator can retype.
  - `'Count is 0, no data to post'` (`:370`) — blocks post when `DailyTotal_Edit` sums to 0, and
    resets `fchanged:=FALSE` (so a subsequent Close won't re-prompt).
  - `'No updated entries'` (`:376`) — Post clicked with `fchanged=False` (nothing edited since load).
  - `'Failed to update shipping info.'` on any `InsertShippingInfoManual`/`InsertShippingDetailManual`
    failure (`:341`/`:354`), each preceded by `Inv_Connection.RollbackTrans` — **the whole day's post
    is transactional** (`BeginTrans` at `:329`), but the operator only learns "failed," no per-part
    detail (mirrors Shipping.pas's `CalculateFRS` opacity).
  - Success: `'Shipping info update is complete.'` (`:360`).
- **`Update_ButtonClick`** (grid entry commit, `:404-448`): `'Invalid count'` (`:443`) if the typed
  count isn't an integer — blanks `Count_Edit` and refocuses; `'End of parts list'` (`:426`) is an
  informational dead-end, not an error, when Update is clicked past the last grid row.
- **`IrregularShip_ButtonClick`** (`:474-506`) — the off-grid one-off adjustment — **no confirmation**
  despite bypassing the normal per-part grid entirely: `'Must select a partnumber'` (`:482`),
  `'Irregular ship count invalid'` (`:487`), `'Failed to update shipping info.'` (`:499`) are the only
  dialogs. Per `shipping.md` finding **M1**, this call is evidenced to fail against the checked-in
  proc signature — the operator would see the generic failure dialog with no hint why.

## Field clear / repopulate
- **`SetDetailBoxes`** (`:167-301`) — same two-branch shape as `Shipping.pas`, but drives a
  `TStringGrid` instead of combos:
  - **Already-shipped branch** (`:181-232`): loads header fields, then re-populates
    `Parts_StringGrid` from `GetPartsListCount` (per-part **posted** qty). If zero rows, sets
    `RowCount:=2` and **blanks row 1's four cells explicitly** (`:212-215`) rather than leaving a
    stale prior grid; `PartNumber_Edit.Text:=''` / `Count_Edit.Text:=''` (`:226-227`) also cleared.
    Locks edits (`ReadOnly:=True`), disables Update/Post, **shows** the Irregular-Ship controls
    (`:229-231`) — irregular ship is only reachable *after* the day is already posted.
  - **Not-yet-shipped branch** (`:234-297`): computes next start-seq (or `'000'` if none), reloads
    `Parts_StringGrid` from `GetPartsList` with **every count defaulted to `'0'`** (`:272` —
    `Cells[3,i]:='0'`, not blank), sets `PartNumber_Edit.Text` to the **first** grid part
    (`Cells[0,1]`, `:277`), blanks `Count_Edit`, `LastSeqNo_Edit`, `DailyTotal_Edit`. Unlocks edits,
    enables Update/Post, **hides** the Irregular-Ship controls.
  - **Ends with `fChanged:=FALSE`** (`:300`) — repopulating the boxes always clears the dirty flag,
    even mid-session on a line/date change; any grid edits not yet Posted are silently forgotten by
    the *dirty-tracking* (though the visible grid cells retain their typed values until the next
    `SetDetailBoxes` call actually overwrites them).
- **Empty-value class:** counts default to the string `'0'`, never blank — a part with no prior
  scrap/build shows `0`, not an empty cell. Consistent, no #135-class gap observed here.
- **`Execute`'s `fUpdateShipping` (redo) branch** (`:100-118`): when re-invoked for an already-known
  line/date, hides the whole Statistics/Irregular/Post UI (`Statistics_GroupBox.Enabled:=FALSE`,
  Irregular controls + `Post_Button.Visible:=FALSE`) — this path exists in code but **no caller in
  this repo sets `fUpdateShipping:=True`** (grep found no external setter); treat as latent/unused
  unless a caller is found elsewhere.

## Focus & keyboard
- **Initial focus (`FormShow`, `:450-453`):** always `Count_Edit.SetFocus` — regardless of whether
  the day is locked (already-shipped) or open; on the locked branch `Count_Edit.ReadOnly` is already
  `True` (`SetDetailBoxes:221`), so focus lands on a read-only field.
- **Grid row selection (`Parts_StringGridSelectCell`, `:455-460`)** — on every cell/row select,
  copies `PartNumber_Edit.Text`/`Count_Edit.Text` from the grid row (no explicit `SetFocus` call
  here; selection is user-driven).
- **`Update_ButtonClick`** (`:404-448`, the per-part grid commit) is **scanner/keyboard-friendly by
  design:** `Count_Edit.SetFocus` (`:417`) at entry, on valid count: writes the cell, **advances the
  grid row** (`Parts_StringGrid.Row:=Row+1`, `:424`), blanks `Count_Edit`, and **re-`SetFocus`es
  `Count_Edit`** (`:429`) — an operator can key count→Enter... but note **`Update_Button` has
  `Default = True`** (`ManualShipping.dfm:78`), so **pressing Enter while focus is in `Count_Edit`
  triggers `Update_ButtonClick`** (VCL routes Enter to the form's Default button when no control
  consumes it) — this is the scanner-driven entry path: type count, Enter, next part is ready.
- **⚠️ Two `Default = True` buttons on one form:** both `Update_Button` (`:78`) **and**
  `IrregularShip_Button` (`:207`) carry `Default = True` in `ManualShipping.dfm`. Delphi VCL applies
  the *last-created* `Default=True` button as the form's actual active default at runtime when both
  are simultaneously visible — **but `IrregularShip_Button` is only `Visible` in the locked
  (already-shipped) branch** (`SetDetailBoxes:231`) while `Update_Button` is only meaningfully used
  in the open branch, so in practice they don't fight over the same visible state. Still: **flag for
  the rebuild** — confirm which control genuinely receives Enter in each state; don't assume "the
  first Default found" is authoritative.
- **No `KeyPreview`** set on the form (not present in `.dfm`); no other Enter-as-tab tricks.

## Enable/disable state machine
- **Locked (already-shipped) state:** `StartSeqNo_Edit`/`LastSeqNo_Edit`/`Count_Edit`
  `ReadOnly:=True`; `Update_Button.Enabled:=FALSE`; `Post_Button.Enabled:=FALSE`;
  `IrregularShipCount_Label/Edit`/`IrregularShip_Button` **`Visible:=TRUE`**.
- **Open (not-yet-shipped) state:** all three `ReadOnly:=False`; `Update_Button.Enabled:=TRUE`;
  `Post_Button.Enabled:=TRUE`; Irregular-ship controls **`Visible:=FALSE`**.
- **`Insert_ButtonClick`/insert-on-fail branch** (`AssemblyPartNumber` fields, ASNInvoice-style
  toggling) — **not present on this form**; ManualShipping has no per-line insert/edit toggle beyond
  the grid.

## Error surfacing
- **100% `ShowMessage` dialogs**, no inline validation labels or status bar. The top-level `Execute`
  wraps everything in one catch-all (`'Error on get information manual shipping screen.'`, `:154-158`).
- **Rollback-on-failure is silent to the operator beyond the generic dialog** — `Inv_Connection.
  RollbackTrans` (`:340`/`:353`) happens before the `ShowMessage`, so the DB state is consistent, but
  the message never says which part/line failed.

## Cross-refs
- Business rules / procs / M1 (`InsertShippingDetailManual` ↔ `INSERT_ShippingDetail` mismatch) /
  M3 (`InsertShippingInfoManual` ↔ `INSERT_ShippingInfo` mismatch):
  [`../../shipping/shipping.md`](../../shipping/shipping.md) §3–4.
