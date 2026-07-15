# Form-UX Semantics — `TOrderFormCreate` (OrderFormCreate.pas) — DEAD CODE, no `.dfm`

## Live-vs-dead verdict

**DEAD CODE — never compiled into the shipping app.**

- Not present in `InventorySystem.dpr` at all (`grep -c "OrderFormCreate\b" InventorySystem.dpr` →
  `0`, confirmed — only `OrderFormCreateF` appears, at `InventorySystem.dpr:33`).
- `TOrderFormCreate = class(TObject)` (`OrderFormCreate.pas:20`) — this is **not a `TForm`**, it has
  no visual surface at all, and correspondingly there is **no `OrderFormCreate.dfm`** in the tree
  (confirmed: `find` for that filename returns nothing).
- No other unit's `uses` clause references `OrderFormCreate` (searched every `.pas` in the repo).
  `MainMenu.pas:196,276-279` — the only caller of this feature — imports and instantiates
  `TOrderFormCreate_Form` from `OrderFormCreateF` (note the trailing `F`), a genuine `TForm` with an
  `Execute`/`ShowModal` + `Hist: THistory` UI (see `docs/analysis/cross-cutting/form-ux/OrderFormCreateF.md`).
- The two units are near-duplicate implementations of the same order-file-emission logic
  (`SELECT_OrderNotOrdered` → per-supplier Excel/text file → `UPDATE_ORDEROrderDate`), but
  `OrderFormCreate.pas` is missing the logistics-directory/Order-Sheet/renban-group/text-timestamp
  features that `OrderFormCreateF.pas` has (compare `OrderFormCreate.pas:33-221` vs
  `OrderFormCreateF.pas:55-707`) — it reads as an **earlier, superseded revision** left in the tree.

## Dialogs & confirmations

Since this unit is never invoked in the live app, its dialogs are **not part of the shipping UX**.
For completeness (in case a fork or future archaeology needs it), the bare `ShowMessage` calls it
contains are: `'There are (<n>) records to process'` (`:63,134`), `'Blank Supplier, invalid order'`
(`:177`), `'No orders to process'` (`:203`) — none gate an action; all are informational OK-only.

## Field clear / repopulate

N/A — no form, no fields.

## Focus & keyboard

N/A — no form, no `.dfm`, no `TForm` ancestor.

## Enable/disable state machine

N/A.

## Error surfacing

If it ran, exceptions log to `Data_Module.LogActLog('ERROR', 'Unable to create order output files, '
+e.message)` (`OrderFormCreate.pas:210`) and set `result:=false`; no dialog on the exception path (the
three `ShowMessage`s above are the only user-visible surface, and only on the non-exception success
paths). This is moot for the rebuild — do not port this unit's behavior.

## Cross-refs

- `docs/analysis/cross-cutting/form-ux/OrderFormCreateF.md` — the LIVE order-file-generation form
  this unit was superseded by.
- `docs/analysis/order/order-file-generation-spec.md` — proc/data-flow spec (specs `OrderFormCreateF.pas`
  only; this file exists so a future reader doesn't re-discover `OrderFormCreate.pas` and wonder
  whether it should be specced too).
