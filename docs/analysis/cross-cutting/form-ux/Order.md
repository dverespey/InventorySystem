# Form-UX Semantics — `TOrder_Form` (Order.pas / Order.dfm)

Caption "Order" (`Order.dfm:6`). Order worksheet select-and-run dialog that drives an Excel-OLE
simulation engine. LIVE — `InventorySystem.dpr:8`.

## Dialogs & confirmations

| Trigger | Text (verbatim) | Buttons | Default | Consequence |
|---|---|---|---|---|
| Start with blank Part Type | `'Select a valid part type'` — `Order.pas:156` | OK | — | `ShowMessage`, `exit`; nothing else happens. The commented-out sibling check for blank Line (`Order.pas:153,155`) is **dead** — Line is never actually validated. |
| Cancel while a worksheet is open (`Start_Button.Enabled = False`) | `'Cancel current worksheet, all changes will be lost?'` — `Order.pas:600` | Yes/No, `mtConfirmation` | none (`0`) | **Yes**: closes the live Excel workbook without saving, resets buttons, logs `'Exit Order without save'` (`:609`). **No**: does nothing (dialog just closes, worksheet remains open). |
| Cancel with no worksheet open | *(no dialog)* | — | — | `Close` immediately (`Order.pas:616`) — form exits straight away. This is the **armed/unarmed split**: the same button is unconfirmed when idle, confirmed-destructive when a worksheet is live. |
| Process Order | `'Create order(s) from worksheet?'` — `Order.pas:634` | Yes/No, `mtConfirmation` | none (`0`) | **Yes**: loops all 200 worksheet rows and calls `INSERT_OpenOrder` per lot (real DB writes, wrapped `BeginTrans`/`CommitTrans` per row — see hazard below). **No**: no-op, worksheet stays open. |
| Invalid qty column value | `'Invalid data in order qty column Line Number(<n>), must be numeric'` — `:664` | OK | — | `exit`s the whole `ProcessOrder_ButtonClick` — **any single bad row aborts the entire order run**, but earlier rows already committed (see hazard). |
| Invalid lot column value | `'Invalid data in order lot column Line Number(<n>), must be numeric'` — `:674` | OK | — | Same abort-mid-loop behavior. |
| Excel/DB failure at any of several points | `'Unable to create order records, '+e.Message'` (`:762`,`:859`), `'Unable to export infomation in order, '+e.Message'` (`:584`), `'Unable to get holiday/overtime infomation in order, '+e.Message'` (`:332`), `'Unable to save orders, check for active edit on worksheet and retry, '+e.Message'` (`:890`), `'Unable to exit, check Excel form for active edit and retry, '+e.Message'` (`:622`) | OK | — | All are bare `ShowMessage` in an `except` handler; the underlying Excel COM object is force-closed (`excel.Workbooks.Close; excel.Quit`) on the export/create-order failure paths, discarding the in-progress worksheet. |

**Hazard — partial-commit on validation failure:** `ProcessOrder_ButtonClick` (`Order.pas:628-893`)
processes rows 1..200 in a single Delphi `for` loop but opens/commits a **separate SQL transaction
per lot** (`BeginTrans`/`CommitTrans` inside the loop body, e.g. `:686,757`). A numeric-validation
failure on row *k* (`:664`,`:674`) `exit`s the procedure — rows `1..k-1` are **already committed** to
the DB, row `k` onward are silently dropped with no rollback of the earlier rows and no indication to
the operator which rows made it. `[UNVERIFIED — confirm before use: exact row count actually
committed on a mid-loop abort — depends on where in the 200-row scan the bad value sits]`.

## Field clear / repopulate

- `FormCreate` (`Order.pas:138-144`): `Date_DateTimePicker` seeded to today (`MinDate:=Date-30`,
  `MaxDate:=Date`, `Date:=today`); `Cancel_Button.Caption:='E&xit'`.
- `FormShow` (`Order.pas:1649-1657`): **every time the form is shown** — `ProcessOrder_Button.Enabled
  := False`; `Line_ComboBox.ItemIndex:=1`; `PartType_ComboBox.ItemIndex:=1` (defaults to the SECOND
  combo item, not the first); `SortBy_Label`/`SortBy_ComboBox.Visible:=FALSE`;
  `SortBy_ComboBox.ItemIndex:=0`. No re-fetch of combo contents happens here — that's done once in
  `Execute` (`:120-121`) before `ShowModal`.
- Cancel-with-confirm (`:608-611`) resets `Cancel_Button.Caption` back to `'E&xit'` and re-enables
  `Start_Button` / disables `ProcessOrder_Button` — returning the form to its pre-Start state, but
  does **not** clear `Line_ComboBox`/`PartType_ComboBox`/`SortBy_ComboBox` selections.
- After a successful order commit (`:882-884`) the same reset happens (`ProcessOrder_Button.Enabled
  :=False; Start_Button.Enabled:=True; Cancel_Button.Caption:='E&xit'`) — the combo selections are
  **left as-is**, so the operator can immediately re-Start on the same Line/PartType.
- `Line_ComboBoxChange` (`Order.pas:1659-1671`): toggles `SortBy_Label`/`SortBy_ComboBox.Visible`
  based on whether `Line_ComboBox.Text` is empty — an **empty Line value shows the sort-by picker**;
  a non-empty value hides it. This is the only "empty incoming value" branch in this form.

## Focus & keyboard

- No `ActiveControl` set in the `.dfm`; Delphi's default (first tab-order control,
  `Start_Button` at `TabOrder=0`, `Order.dfm:75`) receives initial focus.
- No `Default`/`Cancel` button flags set anywhere in `Order.dfm` (confirmed via grep) — Enter does
  **not** trigger any button by VCL convention on this form; no `KeyPreview`/`OnKeyPress` handler
  exists either.
- No `SetFocus` calls anywhere in `Order.pas` — no scripted focus jumps on error or after action.

## Enable/disable state machine

| State | Start_Button | ProcessOrder_Button | Cancel_Button caption |
|---|---|---|---|
| Initial show (`FormShow`) | enabled (default) | `False` | `'E&xit'` |
| After `Start_ButtonClick` succeeds (`:163-164`) | `False` | `True` | `'&Cancel'` (`:160`) |
| Export/create-order exception mid-`Start` (`:589-590`) | `True` | `False` | unchanged |
| Cancel confirmed while worksheet open (`:610-611`) | `True` | `False` | `'E&xit'` |
| Any create-order exception (`:767-768`, `:864-865`) | `True` | `False` | unchanged |
| Create-order success (`:882-884`) | `True` | `False` | `'E&xit'` |

`ProcessOrder_ButtonClick` has no early-exit toggle of `ProcessOrder_Button.Enabled` on the two
numeric-validation `ShowMessage`+`exit` paths (`:664`,`:674`) — the button stays enabled, so the
operator can immediately retry Process Order against the (partially-committed, see hazard) worksheet.

## Error surfacing

100% `ShowMessage` dialogs (never a status label) for both validation and DB/COM exceptions — see
table above. Every DB-write failure path also logs to `Data_Module.LogActLog('ERROR', ...)` with the
same message text before/alongside the dialog (e.g. `:761`/`:762`, `:858`/`:859`).

## Cross-refs

- `docs/analysis/order/legacy-order-spec.md` — proc/data-flow spec for this same form (order
  worksheet, Excel-OLE engine, `INSERT_OpenOrder`/renban-count logic).
