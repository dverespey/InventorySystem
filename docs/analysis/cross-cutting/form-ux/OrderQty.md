# Form-UX Semantics — `TForm1` (OrderQty.pas / OrderQty.dfm) — DEAD CODE, empty stub

## Live-vs-dead verdict

**DEAD CODE — never compiled into the shipping app, and never functional to begin with.**

- Absent from `InventorySystem.dpr` entirely (grep for `OrderQty` in the `.dpr` returns nothing).
- No other `.pas` in the tree references `OrderQty` in a `uses` clause or instantiates `Form1`/
  `TForm1` from it (searched the whole tree; only self-reference is the unit's own declaration at
  `OrderQty.pas:1,10,19`).
- The class name (`TForm1`) and global variable (`Form1`) are the **unrenamed Delphi IDE defaults** —
  a strong signal this was a scratch/experiment form that was never wired up, not a feature that was
  later retired.
- Body has **zero event handlers** (`OrderQty.pas:9-16`: the class declares only the dropped-in
  `ProfGrid1: TProfGrid` component, no `private`/`public` methods, no `implementation` beyond
  `{$R *.dfm}`). It cannot do anything even if instantiated — there is no code path that would ever
  populate, validate, or act on the grid.
- `.dfm` (`OrderQty.dfm`) shows a bare `TProfGrid` (a 3rd-party grid component, `About = 'v2.26
  [Trial-Run]'` at `:21` — note this is a **trial/unlicensed component build**) with generic Delphi
  default caption `'Form1'` (`:6`) and no other controls, confirming it was never dressed for use.

## Dialogs & confirmations

None — no code exists to raise any.

## Field clear / repopulate

N/A — no data-binding code of any kind.

## Focus & keyboard

- No `ActiveControl` set in `OrderQty.dfm`.
- No `Default`/`Cancel` button flags (there are no buttons at all).
- No `OnKeyPress`/`KeyPreview`.

## Enable/disable state machine

N/A — single static component, no interaction code.

## Error surfacing

N/A — no code paths that could raise/catch an error.

## Cross-refs

None applicable — this form has no live counterpart or spec elsewhere in `docs/analysis/order/`.
The actual "how much to order" quantity UX lives in `Order.pas`'s worksheet (qty/lot columns on the
Excel-OLE grid) — see `docs/analysis/cross-cutting/form-ux/Order.md` and
`docs/analysis/order/legacy-order-spec.md`.
