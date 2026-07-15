# Form-UX Semantics — `TOrderFormCreate_Form` (OrderFormCreateF.pas / OrderFormCreateF.dfm)

Caption "Order Form Create" (`OrderFormCreateF.dfm:5`). Batch order-file emitter — runs entirely on
`FormActivate` with no operator input; the only interactive element is a terminal OK button. LIVE —
`InventorySystem.dpr:33` (`OrderFormCreateF in 'OrderFormCreateF.pas' {OrderFormCreate_Form}`),
instantiated from `MainMenu.pas:276-279`. **`OrderFormCreate.pas`** (no trailing `F`, class
`TOrderFormCreate = class(TObject)` — not even a `TForm`) is **absent from `InventorySystem.dpr`**
and has zero callers in the tree — confirmed **dead code**, never compiled into the shipping app.

## Dialogs & confirmations

**None.** This form has no `MessageDlg`/`ShowMessage`/`Application.MessageBox` calls anywhere in
`OrderFormCreateF.pas`. It is a fully unattended batch run — the operator triggers it from
`MainMenu.pas` and cannot cancel, confirm, or be warned mid-run. All progress/outcome is written only
to the embedded `THistory` log control (`Hist: THistory`, `OrderFormCreateF.dfm:30-39`) and to
`Data_Module.LogActLog` (e.g. `OrderFormCreateF.pas:81,250-251,276-277,293-294,687-688,694-696`).

## Field clear / repopulate

N/A — the form has no data-entry fields; it is a live activity/history log view driven entirely by
`FormActivate` (`OrderFormCreateF.pas:55-707`), which begins a single DB transaction
(`Data_Module.Inv_Connection.BeginTrans`, `:82`) spanning the whole multi-supplier file-generation
loop and commits once at the end (`:685`) or rolls back on any exception (`:692`).

## Focus & keyboard

- `Default = True` on `OK_Button` (`OrderFormCreateF.dfm:25`) — Enter triggers OK, but the button
  starts `Visible = False` (`:27`) and is not made visible until `FormActivate` finishes (success:
  `OK_Button.Visible:=True` at `OrderFormCreateF.pas:686`; failure: same at `:693`) — so Enter has no
  effect while the batch is running.
- No `ActiveControl`/`SetFocus` calls anywhere in this unit.
- No `KeyPreview`/`OnKeyPress` handler.

## Enable/disable state machine

Single-state: `OK_Button` is hidden for the entire duration of `FormActivate`'s work and becomes
visible **only on completion**, whether that completion is success (`:686-688`, "Order file
generation complete") or failure (`:693-696`, "Failed on order file creation"). There is no
intermediate cancel affordance — once started, the batch cannot be interrupted from the UI.
`OK_ButtonClick` (`:772-775`) simply `Close`s the form.

## Error surfacing

Exceptions during file generation are caught by a single top-level `except` in `FormActivate`
(`:689-706`): DB transaction is rolled back (`RollbackTrans`, `:692`), `OK_Button` is revealed
(`:693`) so the operator can dismiss the form, and the failure is written to **both** the on-screen
`Hist` log (`Hist.Append('Unable to create order output files, '+e.message)`, `:695`) and
`Data_Module.LogActLog('ERROR', ...)` (`:694`) — **no modal dialog is ever shown**; the operator must
read the history pane to learn what happened. If an Excel COM object or open text file handle was
mid-use it is force-closed (`:697-704`) without any further prompt.

## Cross-refs

- `docs/analysis/order/order-file-generation-spec.md` — proc/data-flow spec for this same form (the
  supplier-file OUTPUT stage: Excel/text emission, `UPDATE_ORDEROrderDate`, FTP/archive/logistics
  directory fan-out).
