# Form-UX Semantics — `TGroupRenbanOrder_Form` (RenbanOrder.pas / RenbanOrder.dfm)

Caption "Renban Group Order" (`RenbanOrder.dfm:6`). Middle-stage trailer-grouping / renban-assignment
tool. LIVE — `InventorySystem.dpr:32` (`RenbanOrder in 'RenbanOrder.pas' {GroupRenbanOrder_Form}`).

## Dialogs & confirmations

| Trigger | Text (verbatim) | Buttons | Default | Consequence |
|---|---|---|---|---|
| Create Renban with a breakdown pending | `'Update these records?'` — `RenbanOrder.pas:413` | Yes/No, `mtConfirmation` | none (`0`) | **Yes**: `BeginTrans`, loops `fAvailableCount` rows calling `NewFRSOrder` (delete-then-reinsert `INSERT_OpenOrder`/`DELETE_OrderRenban`), bumps the renban-group counter (`UPDATE_RenbanGroupCount`), `CommitTrans`, then clears the grid + resets every button back to the pre-breakdown state (`:438-466`). **No**: no branch taken at all — dialog dismissal (either button) other than Yes is a pure no-op; the breakdown stays pending. |
| Closing the form (`FormCloseQuery`) with a breakdown pending | `'There is an unprocessed breakdown waiting, close anyway?'` — `:687` | Yes/No, `mtConfirmation` | none (`0`) | **No** → `CanClose:=False` (form stays open). **Yes** → `FreeList` (frees the in-memory `GroupRenban` truck/order tree) then `CanClose:=True` — the pending breakdown is **discarded**, not saved. |
| Clear Breakdown with a breakdown pending | `'There is an unprocessed breakdown waiting, clear anyway?'` — `:839` | Yes/No, `mtConfirmation` | none (`0`) | **No** → `exit`, nothing cleared. **Yes** (or if nothing pending) → grid wiped, `FreeList`, listbox cleared, `LoadScreen` re-run to reload the un-broken-down renban-group orders, logged `'Clear Renban breakdown'` (`:857`). |
| Trailer/pallet count invalid or insufficient on "Create FRS Breakdown" | `'The total lots will not fit on the selected number of trailers max(<max>):current(<total>)'` — `:814` | OK | — | Informational only; no state change, breakdown not created. |
| Trailers combo not a number | `'Please select trailer count'` — `:821` | OK | — | `Trailers_ComboBox.SetFocus` (`:822`). |
| Trailer pallet count not a number | `'Please select trailer pallet count'` — `:826` | OK | — | `TrailerPalletCount_Edit.SetFocus` (`:827`). |
| No renban-group orders found on load | `'No records for this Renban Group'` — `:648` | OK | — | Grid forced to 1 blank row (`RowCount:=2`, cells cleared `:638-646`), `LoadScreen` returns `False`. |
| Any exception in `Execute`/`CreateOrder`/`LoadScreen` | `'Unable to get Renban Group Orders, '+e.message'` (`:400`,`:657`), `'Unable to update these records,'+e.Message'` (`:472`) | OK | — | Bare `ShowMessage`, logged to `Data_Module.LogActLog('ERROR', ...)`; on the create-order path, if a transaction is open it's rolled back (`:473-474`). |

## Field clear / repopulate

- `Execute` (`RenbanOrder.pas:348-404`): on EVERY open of this form — `RenbanGroups_ComboBox` items
  cleared and refilled from `SELECT_RenbanGroup`, `.Text:=''` (empty selection forced); grid headers
  re-seeded, `RowCount:=2` (i.e. one blank data row); `fBreakdownWaiting:=FALSE`;
  `TrailerPalletCount_Edit.Text:=''`; `TotalLots_Edit.Text:=''`; all action buttons disabled except
  `RenbanGroups_ComboBox` (enabled) and `TrailerCounts_ListBox` hidden.
- `RenbanGroups_ComboBoxChange` → `LoadScreen` (`:575-661`): **every column of every existing row is
  blanked first** (`:595-605`) before repopulating from `SELECT_OrderNoRenban`, so a renban-group with
  zero rows correctly leaves the grid empty (`:636-650`) rather than showing stale data from the
  previously selected group. `TotalLots_Edit` is recomputed live during the fill (`:625`).
- After a successful Create Renban commit (`:438-466`): grid wiped, `TrailerPalletCount_Edit.Text`,
  `TotalLots_Edit.Text`, `RenbanGroups_ComboBox.Text`, `Trailers_ComboBox.Text` all forced back to
  `''` — a full reset requiring the operator to re-pick a renban group from scratch.
- `Trailers_ComboBoxChange` (`:867-879`) auto-computes `TrailerPalletCount_Edit.Text` from
  `TotalLots_Edit.Text div Trailers_ComboBox.Text` whenever the trailer count changes and is
  non-blank; conversely `TrailerPalletCount_EditChange` (`:881-894`) auto-computes
  `Trailers_ComboBox.ItemIndex` from the pallet count — **each field's OnChange writes the other**,
  guarded by a `fTrailerChange` flag (`:873,877,887`) to prevent infinite re-entrant updates. If
  either field is blank, `TryStrToInt` fails silently and the cross-field write is simply skipped
  (no error message) — this is the empty-value class: a blank `Trailers_ComboBox` produces no cascade
  and no complaint until the operator clicks "Create FRS Breakdown" and hits the `Please select
  trailer count` dialog instead.

## Focus & keyboard

- No `ActiveControl` in `RenbanOrder.dfm`; default is `RenbanGroups_ComboBox` (`TabOrder=0`,
  `RenbanOrder.dfm:55`).
- `RenbanGroups_ComboBoxChange` → on successful `LoadScreen`, `Trailers_ComboBox.SetFocus` (`:668`) —
  focus jumps straight to the trailer-count picker once a renban group's orders are loaded.
- Validation-failure focus jumps in `FRSBreakdown_ButtonClick`: `Trailers_ComboBox.SetFocus` (`:822`)
  or `TrailerPalletCount_Edit.SetFocus` (`:827`) depending on which field failed `TryStrToInt`.
- No `Default`/`Cancel` button flags anywhere in `RenbanOrder.dfm` (confirmed via grep) — Enter does
  not trigger OK/Create by VCL convention.
- No `KeyPreview`/`OnKeyPress` handler.

## Enable/disable state machine

| State | ClearBreakdown | CreateOrder | FRSBreakdown | RenbanGroups combo | TrailerCounts listbox |
|---|---|---|---|---|---|
| `Execute` initial (`:389-393`) | `False` | `False` | `False` | `True` | hidden |
| Renban group selected, rows loaded (`LoadScreen:632-634`) | `False` | `False` | `True` | `True` | hidden |
| Renban group selected, zero rows (`:648-650`) | (unchanged from before) | (unchanged) | (unchanged) | (unchanged) | (unchanged) — only the grid is reset; button states are NOT explicitly touched on the "no records" branch, so whatever was enabled from the prior selection **stays enabled** — `[UNVERIFIED — confirm before use: whether this is an intentional carry-over or a missed reset]`. |
| FRS Breakdown computed (`:802-807`) | `True` | `True` | `False` | `False` | `True` |
| Create Order committed (`:454-459`) | `False` | `False` | `False` | `True` | hidden |
| Clear Breakdown (`:837-857`) | (falls through to `LoadScreen`, which re-applies the "rows loaded" state above) | | | | |

## Error surfacing

100% `ShowMessage` dialogs for both validation and DB/COM exceptions (table above); every DB-write
failure also logs via `Data_Module.LogActLog('ERROR', ...)` with matching text (e.g. `:471`/`:472`,
`:569` re-raises after logging so the caller's `except` in `CreateOrder_ButtonClick` shows the
dialog). No status-label surfacing anywhere in this form.

## Cross-refs

- `docs/analysis/order/renban-breakdown-spec.md` — proc/data-flow spec for this same form (trailer
  packing algorithm, `INSERT_OpenOrder`/`DELETE_OrderRenban`/`UPDATE_RenbanGroupCount` bodies).
