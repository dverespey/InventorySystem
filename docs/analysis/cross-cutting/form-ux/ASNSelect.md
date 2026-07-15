# Form-UX: `TASNSelect_Form` — `ASNSelect.pas` + `ASNSelect.dfm`

> "ASN Create" screen — GALC-sequence based EDI 856 (ASN) generation, sibling of `Shipping.pas` but
> for the ASN/856 side rather than the stock-OUT ship post. Confidence: high — both files read in
> full. Business/proc spec: [`../../edi/asn-invoice.md`](../../edi/asn-invoice.md).

## Dialogs & confirmations
- **No confirmations before either create action.** `CreateASNEntries_ButtonClick` and
  `CreateASN_ButtonClick` (`:369-525`) both post directly — no "are you sure you want to create the
  ASN / write the 856 file" prompt, even though the second writes a file to disk and flips
  `ASNStatus:='S'` (`:474-475`).
- **Success/failure are plain `ShowMessage` info dialogs**, not confirmations:
  - `'Failed on ASN select, '+…Errors…` (`:110`), `'Failed to create ASN entries'` (`:416`),
    `'Unable to create EDI856'` (`:501`/`:511`), `'Sequence not found'` (`:260`/`:301`),
    `'Unable to access sequence number date/time, '+e.Message` (`:264`/`:306`).
  - No explicit "success" `ShowMessage` for either create action — success is implied by the form
    resetting (`LoadASNDates`, sequence boxes cleared) with no dialog at all.
- **Every write path wraps `BeginTrans`/`CommitTrans`/`RollbackTrans`** (`:372-429`, `:440-524`) and
  on **any** exception (not just a validation failure) calls `RollbackTrans` if
  `Data_Module.Inv_Connection.InTransaction` — but this is silent to the operator beyond the generic
  `ShowMessage`; no distinct "rolled back" wording.

## Field clear / repopulate
- **`LoadASNDates`** (`:105-135`) is the master repopulate, called on line change and after each
  successful create: fetches `GetNextASNDate`, sets the date picker, and — **critically** —
  **always blanks `StartSeqNo_Edit.Text:=''` and clears `StartBox.Items`** (`:114-115`) *before*
  conditionally repopulating from `Data_Module.StartSeq` if it isn't `'-1'`. So a line with no
  pending start sequence shows a **fully empty** Start box (no placeholder), while `LastSeqNo_Edit`/
  `EndBox` are untouched by this routine (they get cleared by `LoadSeqNumbers` or the
  `StartSeqNo_EditChange` cascade instead) — a genuine two-routine split for what look like a single
  logical "clear the sequence range" operation. **Empty-value class (#135-adjacent):** if
  `Data_Module.StartSeq = '-1'` (no next start found), the Start controls stay blank with **no
  visual cue that this differs from "not yet checked"** — same blank appearance either way.
- **`LoadSeqNumbers`** (`:137-211`) — the ASN-specific per-date sequence lookup (distinct from
  `LoadASNDates`): if `SELECT_ASNSeq` returns a row, locks Start/End edits+boxes to that row's
  values (`ReadOnly:=TRUE` on all four); **else** explicitly blanks all four
  (`StartSeqNo_Edit.Text:=''`, `StartBox.Text:=''`, `LastSeqNo_Edit.Text:=''`, `EndBox.Text:=''`,
  `ShipQty_MaskEdit.Text:='0'`) and unlocks them. Uses the `fInit` reentrancy flag to avoid firing
  `ASN_DateTimePickerChange` recursively while it repopulates the picker's own change event.
- **`ClearCheck`** (`:222-229`) — the "the Check result is now stale" reset: blanks
  `ShipQty_MaskEdit.Text:='0'` (note: `'0'`, not empty — differs from the Start/End edits which go
  fully empty) and disables Check/CreateASN/CreateASNEntries. Called from `LastSeqNo_EditChange` and
  `StartSeqNo_EditChange` whenever the sequence range changes, so any prior Check is invalidated the
  moment either sequence edit changes.
- **Post-create reset** (`CreateASNEntries_ButtonClick:393-409`, `CreateASN_ButtonClick:480-497`):
  `fInit:=FALSE` → blank `StartSeqNo_Edit`/`StartBox` → `LoadASNDates` (which re-derives the next
  window) → `fInit:=TRUE` → blank `LastSeqNo_Edit`/`EndBox` → disable Check/CreateASN/
  CreateASNEntries, `ShipQty_MaskEdit.Text:='0'` → `LastSeqNo_Edit.SetFocus`. A precise multi-step
  choreography; **any rebuild of this reset sequence must reproduce the `fInit` toggling** or risk
  re-firing the date-change handler mid-reset.

## Focus & keyboard
- **Initial focus (`FormShow`, `:324-332`):** `StartSeqNo_Edit` if blank, else `LastSeqNo_Edit` if
  blank, else `ASN_DateTimePicker`.
- **`StartBoxChange`** (`:527-532`) — selecting a Start-time candidate **blanks `LastSeqNo_Edit` and
  clears `EndBox`, then `SetFocus`es `LastSeqNo_Edit`** unconditionally — every Start-box change
  forces the operator to re-pick an End, even if they were just re-confirming the same Start.
  (Contrast `Shipping.pas`'s `StartBoxChange`, which delegates to the same `StartSeqNo_EditChange`
  cascade rather than blindly clearing — a **behavioral divergence between the two sibling forms**
  worth flagging if the rebuild unifies them into one component.)
- **`StartSeqNo_EditChange`/`LastSeqNo_EditChange`** (`:277-322`, `:231-275`) — same "act only once
  `MaxLength` reached" pattern as `Shipping.pas`, gated additionally by the `fInit` flag (so
  programmatic repopulation from `LoadASNDates`/`LoadSeqNumbers` doesn't re-trigger the lookup
  cascade). On a resolved Start, auto-clears End + `ClearCheck`; no explicit focus-move to End here
  (unlike `Shipping.pas`'s equivalent, which does `LastSeqNo_Edit.SetFocus`) — **this form relies on
  natural tab order / `StartBoxChange`'s forced focus instead.**
- **No `Default`/`Cancel` flags, no `KeyPreview`** anywhere in `ASNSelect.dfm` (grep-verified) —
  Enter does not trigger Check/CreateASN.

## Enable/disable state machine
- **Three action buttons gated together, always moved as a set:** `Check_Button`,
  `CreateASN_Button`, `CreateASNEntries_Button`.
  - **Fresh/no-pending-sequence:** Check enabled, both Create buttons disabled
    (`LoadASNDates:125-127`).
  - **Sequence already recorded for this date** (`LoadSeqNumbers`'s found branch, `:180-182`): all
    three **disabled** — the day is already ASN'd, nothing left to do from this screen for that date.
  - **A fresh range typed and resolved** (`StartSeqNo_EditChange`/`LastSeqNo_EditChange`): Check
    re-enabled via `ClearCheck`, both Create buttons stay disabled until Check runs.
  - **After `Check_ButtonClick` succeeds** (`:354-357`): `Check_Button.Enabled:=FALSE`, **both**
    `CreateASN_Button`/`CreateASNEntries_Button` **enabled** — the operator now chooses file-based
    or entries-only creation; there is no way to re-Check without changing the sequence range again.
- **Start/End edit `ReadOnly`** flips in lockstep with the found/not-found branch in
  `LoadSeqNumbers` (`:166-197`) — mirrors `Shipping.pas`'s locked-day pattern but at the
  Start/End-edit level rather than the whole panel.

## Error surfacing
- **100% `ShowMessage` dialogs** for both validation and DB/EDI failures — no inline field errors, no
  status label. `Data_Module.LogActLog('ERROR', …)` is called alongside most dialogs (unlike several
  other forms in this family, which only log on success) — this form is comparatively well-logged.
- **`CreateASN_ButtonClick`'s EDI-write failure path** (`:498-505`) shows
  `'Unable to create EDI856'` — the **same message** used for the upstream `InsertASNInfo` failure
  (`:511`), so the operator cannot distinguish "the DB insert failed" from "the file write failed"
  from the dialog text alone.

## Cross-refs
- Business rules / procs / X12 wire format: [`../../edi/asn-invoice.md`](../../edi/asn-invoice.md),
  [`../../edi/856/edi856-wire-format.md`](../../edi/856/edi856-wire-format.md).
- Sibling GALC-sequence form: [`Shipping.md`](Shipping.md) (same `Check→Post` two-step gate shape,
  same `StartSeqNo_EditChange`/ALC-lookup pattern, diverges on `StartBoxChange` clearing behavior).
