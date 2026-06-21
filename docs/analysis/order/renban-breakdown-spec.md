# Legacy Behavioral Spec — Renban Breakdown (the trailer-grouping / renban-assignment stage)

Source of truth for the **middle stage** of the Order domain. Status: **LIVE** —
`InventorySystem.dpr:32` (`RenbanOrder in 'RenbanOrder.pas' {GroupRenbanOrder_Form}`).
Confirmed compiled; not dead code.

This is the operator-driven step that takes the **blank-renban open orders** created
by the Order worksheet commit (for **palletized / renban-grouped** parts only) and
**splits them across trailers**, assigning each trailer a unique **FRS number** and a
unique **renban number**. After this stage runs, every grouped order has a non-blank
renban and becomes eligible for the order-file emitter.

Pipeline position:

```
Order.pas worksheet commit  ──►  RenbanOrder.pas (THIS)  ──►  OrderFormCreateF.pas
(INSERT_OpenOrder; grouped       (split to trailers,           (emits .ord / Excel;
 parts born BLANK renban)         assign renban + FRS)          filters VC_RENBAN_NUMBER <> '')
   legacy-order-spec.md §6        renban-breakdown-spec.md      order-file-generation-spec.md
   project-library/order/code.py  ← NEW BUILD SCOPE (this doc)
```

**Scope boundary (do not re-spec):** the worksheet build, the lot-sized-vs-palletized
branch at commit, and the `INSERT_OpenOrder` server logic are covered in
`legacy-order-spec.md` §6 and built in `project-library/order/code.py`. The emitter is
`order-file-generation-spec.md`. **This stage shares no code with either** and is
**entirely new build scope** — there is no `renban_breakdown` project-library service yet.

Confidence: **HIGH** on the Pascal flow (`RenbanOrder.pas` read in full) and on all 5
proc bodies (read from the authoritative live dump `DB Schema/CreateInventory.sql`,
UTF-16LE → UTF-8). Two data-dependent claims are flagged with the exact cell to confirm
against the golden.

---

## 1. Where this fits the col R→Q→FRS domain facts (verified)

The `project-order-renban-domain` memo states: worksheet **col R = number of lots**
(operator types an integer), **col Q = Q×O = order qty** (`O = IN_1LOTQTY`), and the
**lot-sized flag is INVERTED (`BIT_LOT_SIZE_ORDERS`: 0 = lot-sized TRUE, 1 =
palletized)**. Verified against source as follows:

- **`BIT_LOT_SIZE_ORDERS` is NOT read in `RenbanOrder.pas` at all.** The lot-sized /
  palletized branch lives entirely **upstream** in `Order.pas` (the worksheet commit).
  By the time orders reach RenbanOrder they are ALREADY the palletized class — they are
  the **blank-renban** rows that the commit deliberately left ungrouped. So the inverted
  flag is **not a hazard in this unit**; it is consumed before this stage. (The inverted
  read still matters for the rebuild's commit service — see `project-library/order/code.py:80`
  — but RenbanOrder never touches the column.) Confirmed: no `BIT_LOT_SIZE_ORDERS` /
  `LotSize` / `LotSizeOrders` reference anywhere in `RenbanOrder.pas`.
- **col R / col Q reappear here as grid columns, re-derived from the DB, not the
  worksheet.** RenbanOrder rebuilds its own 8-column `AvailableGrid` from
  `SELECT_OrderNoRenban` (not from the Excel sheet). The grid's **"Lots" column (grid
  col 6)** is the renban-stage analogue of worksheet col R, and **"Order Qty" (grid
  col 4)** is the analogue of col Q. The mapping is computed (`RenbanOrder.pas:620-623`):
  - grid[4] "Order Qty" = `IN_QTY` (the order row's qty)
  - grid[5] "Lot Qty"   = `IN_1LOTQTY` (= the worksheet's O)
  - grid[6] "Lots"      = `IN_QTY div IN_1LOTQTY` (integer division — **truncates**)
  So **Lots = floor(IN_QTY / IN_1LOTQTY)**. This is the inverse of the worksheet's
  `Q = R × O`: here `R(lots) = Q(qty) / O(lotqty)`. If `IN_1LOTQTY` is 0/NULL this is a
  **divide-by-zero** (`StrToInt('')` or `div 0`) — see Hazard H7.

So **R→Q→FRS in the renban stage** = the operator works in **lots** (col R / grid
"Lots"), the lots map to **qty** (col Q / grid "Order Qty" via `lots × lotqty`), and the
trailer split produces a **new FRS number per trailer** (grid col 3) — confirmed below.

---

## 2. The form & UI — what is interactive vs automatic

`GroupRenbanOrder_Form` (`RenbanOrder.dfm`) — Caption "Renban Group Order". A
**StringGrid-driven** operator screen (NO Excel/COM in this unit — see §8). Controls:

| Control | Type | Role |
|---|---|---|
| `RenbanGroups_ComboBox` | combo | pick the renban group code (e.g. `CMWA`). Populated from `SELECT_RenbanGroup` (`:359`). `CharCase=ecUpperCase`. |
| `AvailableGrid` | 8-col StringGrid | the working set: Kanban, Supplier, Part Number, FRS Number, Order Qty, Lot Qty, Lots, Renban (`:372-379`). Both the **input** (un-grouped orders) and the **output** (post-breakdown trailer rows) are shown here. |
| `Trailers_ComboBox` | combo (1..6) | number of trailers. |
| `TrailerPalletCount_Edit` | edit (max 3 chars) | pallets (= lots) per trailer = trailer capacity. |
| `TotalLots_Edit` | read-only edit | running sum of all lots in the group (auto-filled). |
| `TRailerCounts_ListBox` | listbox | post-breakdown per-trailer load summary (e.g. "Trailer1 count =23"). |
| `FRSBreakdown_Button` | "Create FRS Breakdown" | runs the in-memory trailer split (`:699`). |
| `ClearBreakdown_BitBtn` | "Clear Breakdown" | discards the in-memory split, reloads (`:832`). |
| `CreateOrder_Button` | "Create Renban" | **commits** the split to the DB (`:406`). |
| `OKButton` | OK | closes (`:677`). |

**Interactive (operator):** pick group → pick trailer count + pallets/trailer → click
"Create FRS Breakdown" (preview the split in the grid) → click "Create Renban" (commit).
**Automatic:** the load-balancing distribution math, the FRS-per-trailer numbering, the
renban-per-trailer numbering, and the DB delete+reinsert are all computed.

**Auto-calc convenience (two-way):**
- `Trailers_ComboBoxChange` (`:867`): sets `TrailerPalletCount := TotalLots div Trailers`
  (`:874`). Guarded by `fTrailerChange` to avoid recursion with the next handler.
- `TrailerPalletCount_EditChange` (`:881`): sets `Trailers.ItemIndex := TotalLots div
  palletCount` (`:891`). **`ItemIndex`, not value** — the combo has items `'1'..'6'`, so
  `ItemIndex=k` selects display value `k+1`. **Off-by-one / out-of-range hazard** — see H6.

The two buttons form a **two-phase** flow: phase 1 (`fBreakdownWaiting=FALSE→TRUE`)
previews; phase 2 commits and resets `fBreakdownWaiting→FALSE`. `FormCloseQuery`
(`:682`) blocks close if a breakdown is waiting unprocessed.

---

## 3. The input feed — `SELECT_OrderNoRenban` (the blank-renban selector)

`LoadScreen` (`:575`) drives the grid from `SELECT_OrderNoRenban;1` with
`@RenbanGroupCode = RenbanGroups_ComboBox.Text` (`:588-592`). Proc body
(`CreateInventory.sql:6345`):

```sql
CREATE PROCEDURE [dbo].[SELECT_OrderNoRenban]
    @RenbanGroupCode varchar(5)
AS
    SELECT *
      FROM INV_OPEN_ORDER_INF o
        JOIN INV_PARTS_STOCK_MST p   ON o.VC_PART_NUMBER = p.VC_PART_NUMBER
        JOIN INV_RENBAN_GROUP_MST r  ON r.IN_RENBAN_ID = p.IN_RENBAN_ID
                                    AND r.VC_RENBAN_GROUP_CODE = @RenbanGroupCode
    WHERE o.VC_RENBAN_NUMBER = ''
```

Key behaviors:
- **The selection key is `VC_RENBAN_NUMBER = ''`** — the **blank renban IS the
  "needs grouping" flag**. This is the same blank that the Order commit deliberately
  leaves on grouped parts (`legacy-order-spec.md` §6; domain memo: "born blank-renban").
  It is the exact mirror of the emitter's `VC_RENBAN_NUMBER <> ''` gate
  (`order-file-generation-spec.md` §8) — RenbanOrder converts `''`→non-blank so the
  emitter can pick the rows up.
- Joined to `INV_RENBAN_GROUP_MST` on `IN_RENBAN_ID = p.IN_RENBAN_ID` **and** the chosen
  group code → only this group's parts. **INNER JOIN on the group**, so a blank-renban
  order whose part is NOT in any renban group never appears here. (No `ORDER BY` →
  result order is engine-dependent; the grid just numbers rows in arrival order.)

Per row the grid is filled (`:615-624`):

| Grid col | Header | Source | Notes |
|---|---|---|---|
| 0 | Kanban | `VC_KANBAN_NUMBER` | |
| 1 | Supplier | `VC_SUPPLIER_CODE` | |
| 2 | Part Number | `VC_PART_NUMBER` | |
| 3 | FRS Number | `VC_FRS_NUMBER` | the placeholder FRS from commit |
| 4 | Order Qty | `IN_QTY` | |
| 5 | Lot Qty | `IN_1LOTQTY` | = worksheet col O |
| 6 | Lots | `IN_QTY div IN_1LOTQTY` | **integer truncation** (H7) |
| 7 | Renban | `VC_RENBAN_GROUP_CODE` + `VC_RENBAN_GROUP_COUNT` | the **seed** renban (group code + 3-char count), e.g. `CMWA000` |

`TotalLots_Edit` accumulates `Σ (IN_QTY div IN_1LOTQTY)` across all rows (`:623-625`).
`fAvailableCount := recordcount` (`:609`). Empty result → "No records for this Renban
Group" (`:648`).

> Confirm against golden: the **seed renban** in grid col 7 is
> `VC_RENBAN_GROUP_CODE || VC_RENBAN_GROUP_COUNT` (e.g. group `CMWA`, count `000` →
> `CMWA000`). `VC_RENBAN_GROUP_COUNT` is `varchar(3)`; the memo confirms `000` is valid
> (it is a trailer-uniqueness seed that rolls over, not a quantity). The breakdown reads
> the **3-digit tail** of this back out (`:775`) as the renban base — verify the seed for
> group `CMWA` is exactly `000` in the golden, since the final renbans are `seed +
> trailerNumber` (§5).

---

## 4. The grouping algorithm — load-balancing lots across trailers

The grouping is **by trailer (truck)**, NOT by part or by FRS. All of a renban group's
un-grouped order lines are pooled and **distributed across N equal-capacity trailers** by
a round-robin fill. The in-memory model (`RenbanOrder.pas:20-92`):

- `TOrder` — one part's order on one trailer: kanban, supplier, partnumber, frsnumber,
  lotqty, **lots** (count on this trailer), renban.
- `TTruck` — one trailer: a fixed `Size` (= pallets/trailer), a `CurrentCount` (lots
  loaded so far), and an `fOrderList` keyed **by part number** (so multiple lots of the
  same part on the same trailer **merge** into one `TOrder`, summing `lots` — `:160-164`).
- `TGroupRenban` — the set of trailers; `SetTrucks(N, Size)` creates N `TTruck`s named
  `Trailer1..TrailerN`, each with capacity `Size` (`:239-253`).

### 4.1 Gate (`FRSBreakdown_ButtonClick`, `:699`)
Reads `tcount` = trailer count, `pcount` = pallets/trailer. **Capacity check**:
`tcount * pcount >= TotalLots` (`:709`) — else error "will not fit". Then
`GroupRenban.SetTrucks(tcount, pcount)` and feeds every grid row via `AddOrder` (`:719-729`).

### 4.2 The distribution math (`TGroupRenban.AddOrder`, `:255-301`)
For each part line with `lots` lots, distribute across `T = trailer count` trucks:

**Phase A — even share (`:262-277`):** if `lots div T <> 0`, give each truck a base
share of `lots div T` (+ any `leftover` carried from a truck that overflowed):
```
for each truck i:
  share = (lots div T) + leftover
  if truck[i].CurrentCount + share <= truck[i].Size:
      add `share` lots of this part to truck[i];  leftover := 0
  else:                                    # truck full — spill the excess forward
      leftover := share - (truck[0].Size - truck[i].CurrentCount)   # NOTE: truck[0].Size
      add (truck[i].Size - truck[i].CurrentCount) lots to truck[i]  # top it off
```
**Phase B — remainder (`:279-300`):** the `lots mod T` (+ leftover) units are dribbled
**one lot at a time, round-robin** across trucks that still have room, until exhausted
(`:285-298`).

So a part's lots are spread as evenly as possible across all trucks, with overflow
spilling to later trucks and a final one-at-a-time top-up. Each truck merges repeated
parts (`TTruck.AddOrder` `:155-182`). **Result: per (truck, part) one `TOrder` with a
`lots` count; `CurrentCount` per truck ≈ capacity.**

> The leftover formula at `:273` uses `truck[0].Size` (truck **zero**'s size) where it
> arguably means `truck[i].Size`. Because **all trucks are created with the same `Size`**
> (`SetTrucks` `:249`), `truck[0].Size == truck[i].Size` always, so it is **behaviorally
> correct today** — but a rebuild with variable-capacity trucks would break here. Hazard
> H4 (faithful-but-fragile; do NOT "fix" by changing the distribution outcome).

### 4.3 Read-out → grid + FRS + renban (`:746-795`)
Walk every truck, every order on it (`GroupRenban.First/Next`, `Truck.First/Next`), and
write a grid row per (truck, part):

| Grid col | Value | Source line |
|---|---|---|
| 0 Kanban | `Order.kanban` | `:758` |
| 1 Supplier | `Order.supplier` | `:759` |
| 2 Part | `Order.partnumber` | `:760` |
| 3 **FRS** | `copy(frsnumber,1,5)` + trailer suffix | `:763-767` (see §4.4) |
| 4 **Order Qty** | `lots × lotqty` | `:770` |
| 5 Lot Qty | `Order.lotqty` | `:771` |
| 6 Lots | `Order.lots` | `:772` |
| 7 **Renban** | `RenbanGroups_ComboBox.Text` + `%.3d(rcount)` | `:779` (see §5) |

`fAvailableCount := AvailableGrid.RowCount-1` (`:799`) — the working set is now the
**trailer rows**, not the original orders. The per-truck loads are listed in
`TRailerCounts_ListBox` (`:792`).

### 4.4 FRS-number-per-trailer (`:763-767`)
The new FRS = the **first 5 chars of the original FRS** (the `YMMDD` order-by prefix) +
a 2-char trailer suffix = **`TruckNumber + 1`** (1-based trailer index):
```pascal
AvailableGrid.Cells[3,x] := copy(Order.frsnumber,1,5);
if GroupRenban.TruckNumber > 8 then
    Cells[3,x] := Cells[3,x] + IntToStr(TruckNumber + 1)      // 2-digit, e.g. 10,11…
else
    Cells[3,x] := Cells[3,x] + '0' + IntToStr(TruckNumber + 1); // zero-padded, 01..09
```
`TruckNumber` is the 0-based `fTruckOffset` (`:90-91`), so trailer 0 → FRS suffix `01`,
trailer 1 → `02`, … trailer 9 (offset 9) → `10`. **The FRS trailer suffix = trailer
ordinal** (each trailer is one FRS "trip" of the day). This is consistent with the domain
fact "FRS = the sequence of use in the day."

> Note: this FRS suffix is what the Pascal computes for the grid, but `INSERT_OpenOrder`
> **re-derives the trailing 2 digits server-side** (`legacy-order-spec.md` §6;
> `CreateInventory.sql:7388-7436`). For grouped parts (renban now non-blank at insert
> time), the proc matches `max(VC_FRS_NUMBER)` for the **part** (`@RenbanNum<>''` branch,
> `:7390-7393`) and increments. So the **proc owns the final 2 FRS digits**; the grid
> suffix seeds the prefix and the proc's max+1 sequences the trailers as they insert. See
> §6 + Hazard H8 for the interaction.

---

## 5. Renban-number assignment (the never-blank guarantee)

The renban for trailer row `x` (`:775-779`):
```pascal
rcount := StrToInt(rightstr(Order.renban, 3));   // the 3-digit seed count, e.g. '000' -> 0
rcount := rcount + GroupRenban.TruckNumber;       // + 0-based trailer index
Cells[7,x] := RenbanGroups_ComboBox.Text + format('%.3d', [rcount]);  // group code + 3 digits
```
So **renban = `<groupCode><seedCount + trailerOrdinal, zero-padded to 3>`**. With group
`CMWA`, seed `000`: trailer 0 → `CMWA000`, trailer 1 → `CMWA001`, trailer 2 → `CMWA002`.
**Every trailer in the group gets a distinct renban**; all parts riding the same trailer
share that trailer's renban (they ship together — the domain definition of a renban
group). The format is **group code (≤5 chars) + 3-digit zero-padded counter**.

**Counter advance:** after the read-out, `fNewMaxRenban := rcount + 1` (`:798`) — i.e. the
**highest trailer renban + 1**, the next free count for the group. On commit (§6) this is
written back via `UPDATE_RenbanGroupCount` so the **next** breakdown of this group starts
where this one ended (the count rolls forward, rolls over per the memo). `Format('%.3d',
[fNewMaxRenban])` is `varchar(3)`.

**Never-blank invariant — how it is enforced:**
1. The selector pulls **only** `VC_RENBAN_NUMBER = ''` rows (§3). Those rows are then
   **deleted and reinserted** with the computed non-blank renban (§6). So the breakdown
   **eliminates** the blanks it consumes.
2. The downstream emitter filters `VC_RENBAN_NUMBER <> ''` (`order-file-generation-spec.md`
   §8). Any blank-renban order that is **never run through a breakdown** (e.g. its renban
   group was never selected, or the operator never grouped it) **silently never ships** —
   it sits forever in `INV_OPEN_ORDER_INF` with a blank renban, invisible to the emitter.
   There is **no error, no default** — the enforcement is "the emitter ignores you."
   Hazard H1 (the never-blank invariant is a *filter at both ends*, not a guard;
   un-grouped grouped-parts are a silent stuck-order class).

---

## 6. The write-back — DELETE then re-INSERT (the supersede lifecycle)

`CreateOrder_ButtonClick` (`:406`), guarded by `fBreakdownWaiting` + a "Update these
records?" confirm (`:413`), runs in ONE transaction (`Inv_Connection.BeginTrans` `:416`
… `CommitTrans` `:436`; `RollbackTrans` on exception `:473-474`). For each trailer row it
calls `NewFRSOrder` (`:417-421`), then bumps the group counter (`:423-434`), then clears
the grid.

### 6.1 `NewFRSOrder` (`:481`) — delete the placeholder, insert the grouped order
The active code path (most of the proc body is **commented out** — a prior buggy
update-in-place approach; see Hazard H3) does **two** things per grid row:

**(a) DELETE the original blank-renban placeholder** (`DELETE_OrderRenban;1`, `:502-511`):
```
@PartNumber  = grid[2] (part)
@FRSNumber   = ''        <-- forced blank (NOT the grid FRS)
@RenbanNumber= ''        <-- forced blank
```
Proc (`CreateInventory.sql:7629`): with `@FRSNumber=''`, deletes **by part + blank
renban**:
```sql
IF @FRSNumber = ''
   DELETE FROM INV_OPEN_ORDER_INF
   WHERE VC_PART_NUMBER = @PartNumber AND VC_RENBAN_NUMBER = @RenbanNumber  -- (= '')
```
So it removes **ALL still-blank-renban rows for that part** — the original placeholder(s)
from commit. (It does NOT delete already-grouped rows; those have a non-blank renban.)
Note: the **first** trailer row for a part deletes the placeholder; subsequent trailer
rows for the same part call DELETE again but it is a no-op (the blank is gone) — benign.

**(b) INSERT the grouped order** (`INSERT_OpenOrder;1`, `:540-563`), **only if grid Order
Qty > 0** (`:540` — "small runner don't send 0 qty orders"):
```
@SupCode  = grid[1]   @PartNum = grid[2]   @KanbanNum = grid[0]
@FRSNumber= grid[3]   (the 5-char prefix + trailer suffix, §4.4)
@RenbanNum= grid[7]   (group code + 3-digit, §5 — NON-BLANK)
@Qty      = grid[4]   (lots × lotqty)
```
`INSERT_OpenOrder` (`CreateInventory.sql:7358`): because `@RenbanNum <> ''`, it takes the
**part-scoped** max-FRS branch (`:7388-7393`), re-derives the trailing 2 FRS digits
(max+1, or `01` if first), computes `VC_FRS_DATE` (with year-rollover, `:7378-7386`), and
inserts into `INV_OPEN_ORDER_INF` (`VC_SUPPLIER_CODE, VC_PART_NUMBER, VC_KANBAN_NUMBER,
VC_FRS_NUMBER, VC_FRS_DATE, VC_RENBAN_NUMBER, IN_QTY, VC_ADD`). The `INSERT_RecConfStat-
PartsStockMstQTY` trigger fires (copies to `_HIST`, but does **not** bump
`INV_PARTS_STOCK_MST.IN_QTY` because shipping status is empty — same as commit,
`legacy-order-spec.md` §6).

> **This is the "delete-supersede" lifecycle behind R3** (memo): the breakdown DELETEs the
> blank placeholder and re-INSERTs grouped rows, so by Order-Start (post-grouping) NO
> blank rows remain — which is why the downstream receipt projection's sum-all-rows looks
> renban-only and is **faithful, not a bug**. Do NOT "fix" the sum-all behavior.

### 6.2 Bump the group counter (`UPDATE_RenbanGroupCount;1`, `:423-434`)
```
@RenbanCode  = RenbanGroups_ComboBox.Text
@RenbanCount = Format('%.3d', [fNewMaxRenban])   -- next free count (§5)
```
Proc (`CreateInventory.sql:5142`): sets `INV_RENBAN_GROUP_MST.VC_RENBAN_GROUP_COUNT =
@RenbanCount` and `VC_LAST_UPDATE` for the group code. This **persists the renban counter
forward** so the next breakdown of the group continues the sequence (vs the per-part
`IN_RENBAN_COUNT` that lot-sized commit uses). The grid seed in §3 reads this value back
on the next run.

### 6.3 The relationship to the emitter gate + `UPDATE_ORDEROrderDate`
After commit, the grouped orders have **non-blank renban** and **null/blank
`VC_ORDER_DATE`** (a fresh INSERT). They therefore satisfy BOTH emitter predicates:
`(VC_ORDER_DATE is null or '')` AND `VC_RENBAN_NUMBER <> ''`
(`order-file-generation-spec.md` §8). The emitter picks them up, emits the `.ord`/Excel,
then stamps `VC_ORDER_DATE` via `UPDATE_ORDEROrderDate` so they drop out. RenbanOrder
does **not** stamp `VC_ORDER_DATE` itself — it only makes the rows *eligible*. The two
idempotency guards are layered: blank-renban filter (RenbanOrder's input) → non-blank +
no-order-date (emitter's input) → order-date stamp (emitter's output).

---

## 7. The transient FRSBreakdown.pas dialog (uses-clause; effectively dead here)

`RenbanOrder.pas:138` `uses ... FRSBreakdown`, but **`TFRSBreakdownDlg` is never
instantiated** in `RenbanOrder.pas`. The "Create FRS Breakdown" button (`:699`) does the
trailer split **inline** with `TGroupRenban`/`TTruck`, not via the dialog. `FRSBreakdown`
(`InventorySystem.dpr:34`, LIVE) is a separate single-part split dialog (split one part's
lots across `n` FRS trucks, comma-list preview) used elsewhere — its `uses` here is a
**stale import** (dead dependency in this unit). Do not port a `FRSBreakdownDlg`
interaction into the renban-breakdown rebuild; the trailer split is the inline algorithm
of §4. (FRSBreakdown body: `FRSBreakdown.pas:78-141` — pure UI, no DB; computes
`flots div count (+remainder on first)` — not used by RenbanOrder.)

---

## 8. Excel / COM dependency — NONE

Unlike `Order.pas` (worksheet) and `OrderFormCreateF.pas` (emitter), **RenbanOrder.pas
uses no Excel/COM** — no `createOleObject`, no `Cells[r,c].value` on a workbook. All
"Cells" references are the **VCL `TStringGrid`** (`AvailableGrid.Cells[col,row]`), an
in-memory UI grid. The rebuild needs **no Excel replacement** for this stage — it is a
pure DB + in-memory-distribution + grid screen. This is the **lowest-COM-risk** of the
three Order stages.

---

## 9. What the rebuild's renban breakdown MUST reproduce

1. **Input selector:** pull `INV_OPEN_ORDER_INF` rows where `VC_RENBAN_NUMBER = ''`,
   joined to the chosen renban group (inner join on `IN_RENBAN_ID` + group code). Blank
   renban IS the selection flag. (`SELECT_OrderNoRenban`.)
2. **Lots = `floor(IN_QTY / IN_1LOTQTY)`** per row; **TotalLots = Σ lots** across the
   group. Guard divide-by-zero (H7).
3. **Capacity gate:** require `trailers × palletsPerTrailer >= TotalLots`; else refuse.
4. **Distribution:** spread each part's lots across `N` equal-capacity trailers — even
   `lots div N` base share with forward-spill of overflow, then round-robin the
   `lots mod N` remainder one lot at a time; merge repeated parts per trailer (sum lots).
   Faithful outcome: trailers fill to capacity in order, last trailer holds the remainder.
5. **Per trailer-row outputs:**
   - **FRS** = `<original FRS [0:5]>` + 2-digit trailer ordinal (`trailerIndex+1`,
     zero-padded <10). The server-side `INSERT_OpenOrder` then owns the final 2 digits
     (max+1 per part) — replicate that ownership (§6.1b / H8).
   - **Order Qty** = `lots × lotqty` per trailer-row.
   - **Renban** = `<groupCode>` + 3-digit (`seedCount + trailerIndex`), zero-padded.
     **Never blank.** seedCount = the 3-digit tail of `VC_RENBAN_GROUP_CODE ||
     VC_RENBAN_GROUP_COUNT`.
6. **Commit = DELETE-then-INSERT, one transaction:** per part, `DELETE_OrderRenban(part,
   FRS='', renban='')` removes the blank placeholder(s); then `INSERT_OpenOrder(...,
   renban=non-blank, qty>0)` per trailer-row (skip qty=0). Then `UPDATE_RenbanGroupCount`
   persists `maxRenban+1` for the group. Atomic; rollback on any failure.
7. **Eligibility contract:** the committed rows must satisfy the emitter's
   `(VC_ORDER_DATE null/'') AND VC_RENBAN_NUMBER <> ''` so they flow to the order files.
   RenbanOrder does NOT stamp `VC_ORDER_DATE`.
8. **No Excel.** Pure DB + in-memory split + a grid/table UI.

A clean rebuild split: a **pure** `computeTrailerBreakdown(rows, trailers, palletsPer)`
→ trailer-rows (FRS/qty/renban, unit-testable like `computeOrderRecords`), and a
**driver** `commitBreakdown(group, trailerRows)` that does the delete/insert/count-bump in
one transaction. Named Queries: `SELECT_OrderNoRenban`, `DELETE_OrderRenban`,
`INSERT_OpenOrder` (shared with commit), `UPDATE_RenbanGroupCount`, `SELECT_RenbanGroup`.

---

## 10. Hazards (first-class findings)

- **H1 — Silent stuck-order class for ungrouped grouped-parts.** The never-blank
  invariant is a *filter at both ends* (RenbanOrder pulls `=''`; emitter pulls `<>''`),
  **not a guard**. A blank-renban order whose group is never run through the breakdown
  **never ships and never errors** — it sits in `INV_OPEN_ORDER_INF` forever, invisible to
  the emitter. The rebuild should surface "blank-renban orders awaiting grouping" as an
  operational alert, not rely on the operator remembering. (§5)
- **H2 — Files/DB are atomic *within* the commit, but the breakdown is fully reversible
  only if not yet committed.** The "Create FRS Breakdown" preview is in-memory; only
  "Create Renban" writes. Good. But note the delete-then-insert: if the transaction is
  interrupted between DELETE and the matching INSERT (it isn't, both are in one tx —
  `:416`/`:436`), the placeholder would be lost. Preserve the single-transaction
  delete+insert+count-bump. (§6)
- **H3 — Large commented-out update-in-place path (`:482-539`).** The original design
  tried `UPDATE_OrderRenbanQty` to update the *first* trailer's order in place and only
  delete/insert the rest. It was abandoned for **delete-all-then-reinsert** because (per
  the code comment `:489-493`) "an error can occur where only one item order can be put on
  any FRS truck — if it's not the first truck the initial order is left without a Renban
  number." **Do NOT resurrect the update-in-place path**; the live behavior is
  delete-all-blank + reinsert-all. `UPDATE_OrderRenbanQty` (`CreateInventory.sql:5890`)
  and `UPDATE_OrderRenban` (`:5910`) exist but are **not called** by this unit.
- **H4 — `truck[0].Size` in the leftover formula (`:273`).** Uses truck **zero**'s size
  where it means truck **i**'s. Correct **only** because all trucks share one `Size`.
  Faithful today; a variable-capacity rebuild must use the per-truck size. Do not change
  the distribution *outcome* for the equal-capacity case.
- **H5 — `TGroupRenban.Next` omits `fEof:=FALSE`/`fEof:=TRUE` symmetry.** `First`
  (`:303`) sets `fEof:=FALSE`; `Next` (`:326`) sets `fEof:=TRUE` only at exhaustion and
  never re-sets FALSE — harmless because `fEof` is already FALSE after `First` until
  exhausted. Preserve the iteration *contract*, not the literal field-setting.
- **H6 — `TrailerPalletCount_EditChange` sets `Trailers.ItemIndex` (not value), with no
  range guard (`:891`).** `ItemIndex := TotalLots div palletCount` — for `TotalLots=20,
  palletCount=10 → ItemIndex=2 → display '3'` (off-by-one vs intent), and any index ≥6 or
  the `-1`/overflow cases leave the combo unselected. It is a UX convenience, not the
  authoritative split (the authoritative numbers are read at `FRSBreakdown_ButtonClick`),
  so it is cosmetic — but a faithful rebuild should reproduce the auto-fill *intent*
  (`trailers ≈ totalLots / palletsPerTrailer`) cleanly, not the buggy `ItemIndex` math.
- **H7 — Divide-by-zero on `IN_1LOTQTY` = 0/NULL (`:622`, `:623`).** `IN_QTY div
  IN_1LOTQTY` with a 0/NULL lot qty raises (`div 0` or `StrToInt('')`). The whole
  LoadScreen aborts (caught `:654`, logged, `result:=False`). A grouped part with no lot
  qty stalls the screen. Guard in the rebuild.
- **H8 — FRS-suffix double-ownership.** The grid computes a trailer-ordinal FRS suffix
  (§4.4) but `INSERT_OpenOrder` **re-derives** the trailing 2 digits server-side
  (`CreateInventory.sql:7388-7436`, `@RenbanNum<>''` → part-scoped `max+1`). For the FIRST
  trailer of a part the proc emits `01`; for later inserts it does `max+1`. Because the
  trailers insert in order and each is a new max, the proc's sequence *usually* matches the
  grid ordinals — **but they are computed independently.** A rebuild that trusts the grid
  FRS without replicating the proc's max+1 (or vice-versa) can drift. Confirm against a
  golden: for a 3-trailer `CMWA` breakdown, that the persisted `VC_FRS_NUMBER` trailing 2
  digits are `01/02/03` AND the renbans are `CMWA000/001/002`. Name the exact cells: pick a
  group, run the breakdown, read back `INV_OPEN_ORDER_INF.VC_FRS_NUMBER` and
  `VC_RENBAN_NUMBER` for those parts and compare to the grid.
- **H9 — `qty=0` trailer-rows are silently dropped at INSERT (`:540`)** ("small runner
  don't send 0 qty orders") but the **DELETE still fired** for the part (placeholder
  removed). If a part's lots all land on trailers as 0-qty rows (cannot happen with the
  current math, since a part with ≥1 lot always lands ≥1 lot somewhere), the placeholder
  would be deleted with no replacement. Faithful but verify no real input produces it.
- **H10 — No `ORDER BY` in `SELECT_OrderNoRenban` (`:6345`).** Grid row order is
  engine-dependent. The breakdown outcome is order-independent for the *trailer fill*
  (each part is distributed independently), so this is benign — but a rebuild that, e.g.,
  numbers FRS by grid position must not assume a stable input order.

---

## 11. Multi-site notes (defer design to ignition-architect)

- This stage has **no INI/template/path dependency** (no Excel, no filesystem) — it is
  pure DB + UI. The only single-site coupling is **implicit**: `INV_RENBAN_GROUP_MST` and
  `INV_OPEN_ORDER_INF` are single-site tables today; a multi-site rebuild scoping these by
  site (per `project-multisite`) automatically scopes the breakdown. The renban **counter**
  (`VC_RENBAN_GROUP_COUNT`) is per-group, single-valued — multi-site needs it per
  (site, group) or the trailer-uniqueness seed collides across sites.
- The connection is `Inv_Connection` only (no cross-DB ALC call in this unit — unlike the
  worksheet and the emitter, which need `AD_GetSpecialDate` on `TireOrder`). So the
  breakdown stage is **self-contained within the Inventory DB.**

---

## 12. REBUILD ALGORITHM (STEP-0 extraction — the EXACT trailer distribution from `RenbanOrder.pas`)

This section is the literal, line-cited transcription of the in-memory distribution + the
FRS/renban assignment that the rebuild's `compute_trailer_breakdown` must reproduce
**bit-for-bit**. It is the authoritative target for `project-library/renban/code.py` and
`scripts/e2e/test_renban_build.py`. Every claim is from `RenbanOrder.pas` (read in full).

### 12.0 Inputs (per group, from `SELECT_OrderNoRenban` + the two operator entries)

Per blank-renban order row (aliased — never `SELECT *`, §3 / data-analysis §2.2):
`kanban (o.VC_KANBAN_NUMBER)`, `supplier (o.VC_SUPPLIER_CODE)`, `part (o.VC_PART_NUMBER)`,
`frs (o.VC_FRS_NUMBER)`, `order_qty (o.IN_QTY)`, `lotqty (p.IN_1LOTQTY)`,
`group_code (r.VC_RENBAN_GROUP_CODE)`, `group_count (r.VC_RENBAN_GROUP_COUNT)`.
Per row: `lots = order_qty div lotqty` (integer truncation, `:622`). Operator: `T` =
trailer count, `P` = pallets/trailer (= each truck's capacity `Size`).

### 12.1 Capacity gate (`FRSBreakdown_ButtonClick`, `:709`)

```
TotalLots = Σ (order_qty div lotqty)          # over every group row, :623
require  T * P >= TotalLots                    # else "will not fit" — abort, :709
```

### 12.2 Per-part distribution across trucks (`TGroupRenban.AddOrder`, :255-301) — VERBATIM

For each part line (its own `lots`), `T = number of trucks`, trucks indexed `0..T-1`, each
with `.Size = P` and a running `.CurrentCount` (lots already loaded across ALL prior parts).
`leftover` starts at 0 **per part call**.

```
leftover = 0

# --- Phase A: even base share (only entered when lots div T <> 0) ------------  :262-277
if (lots div T) != 0:
    for i in 0 .. T-1:
        share = (lots div T) + leftover
        if truck[i].current + share <= truck[i].size:        # fits
            truck[i].add(part, lots=share)                    # :268
            leftover = 0
        else:                                                  # truck full -> spill forward
            leftover = ((lots div T) + leftover) - (truck[0].size - truck[i].current)   # :273  NB truck[0]
            truck[i].add(part, lots=(truck[i].size - truck[i].current))                 # :274  top it off

# --- Phase B: remainder dribbled one lot at a time, round-robin -------------  :279-300
if (lots mod T) != 0:
    remainder = (lots mod T) + leftover                        # :282  (carry the Phase-A overflow too)
    while remainder != 0:                                       # :285
        for i in 0 .. T-1:                                      # :287
            if truck[i].current + 1 <= truck[i].size:           # :289
                truck[i].add(part, lots=1)                      # :291
                remainder -= 1
            if remainder == 0: break                            # :295-296
```

**FAITHFULNESS NOTES (must replicate exactly, do NOT "fix"):**
- **`truck[0].size` at `:273`** (NOT `truck[i].size`). Behaviorally identical *only* because
  `SetTrucks` (`:249`) gives every truck the same `Size = P`. The rebuild keeps the
  equal-capacity outcome; it must use the **common P** here, matching the Pascal. (Spec H4.)
- **Phase B is an INFINITE LOOP if `remainder` cannot be placed** (all trucks full). The
  capacity gate (12.1) is what prevents it: `T*P >= TotalLots` guarantees room. A rebuild MUST
  keep the gate (and additionally bound the loop to fail loudly rather than hang — see 12.6).
- **`leftover` carries from Phase A into Phase B** (`:282`). When a truck overflows in Phase A,
  the spilled remainder is added to the mod-remainder and round-robined.

### 12.3 Merge repeated parts on a truck (`TTruck.AddOrder`, :155-182) — the R3 sum-all

```
on truck.add(part, lots):
    if part already on this truck:  existing.lots += lots          # :160-163  (SUM, do not split)
    else:                            create one order line for part  # :166-177
    truck.CurrentCount += lots                                       # :179
```

One truck carries **at most one line per part** (lots summed). This is the trailer-level R3
sum-all faithfulness (data-analysis §5.3 / memory `project-order-renban-domain`). Do NOT emit
per-lot duplicate lines.

### 12.4 Read-out → per-trailer rows (`:746-799`)

Walk trucks outer (`GroupRenban.First/Next`), orders inner (`Truck.First/Next`); emit one row
per (truck, part) that has a line. `TruckNumber` = the truck's 0-based index (`fTruckOffset`).
A truck with NO orders contributes NO rows (and no listbox spill effect on the others).

| Field | Formula | Lines |
|---|---|---|
| kanban / supplier / part | from the merged `TOrder` | `:758-760` |
| **FRS** | `copy(frs,1,5)` + 2-digit trailer ordinal (see 12.5) | `:763-767` |
| **Order Qty** | `lots × lotqty` | `:770` |
| Lot Qty / Lots | `lotqty` / merged `lots` | `:771-772` |
| **Renban** | `group_code` + `%.3d(seed3 + TruckNumber)` (see 12.5) | `:775-779` |

### 12.5 FRS suffix + renban number (the per-trailer identity)

```
# FRS (:763-767): first 5 chars of the original FRS + the trailer ordinal (TruckNumber+1)
frs_prefix = frs[0:5]
ordinal    = TruckNumber + 1                       # 1-based
frs_out    = frs_prefix + (str(ordinal)  if TruckNumber > 8 else "0" + str(ordinal))
#   truck 0 -> '01', truck 7 -> '08', truck 8 -> '09', truck 9 (TruckNumber>8) -> '10'
#   -> 2 raw digits once the ordinal reaches 10 (TruckNumber>8 means ordinal>=10).

# Renban (:775-779): group code + 3-digit (seed count + trailer index), zero-padded
seed3      = int( renban_seed[-3:] )               # the 3-digit tail of group_code||group_count, e.g. 'CMWA288' -> 288
rcount     = seed3 + TruckNumber                   # 0-based trailer index
renban_out = group_code + "%03d" % rcount          # e.g. CMWA288, CMWA289, CMWA290 for trucks 0,1,2
```

The renban_seed fed in is `group_code || group_count` (built at load, `:624`), e.g.
`CMWA` + `288` = `CMWA288`; `rightstr(...,3)` reads `288` back out. **Renban is NEVER blank**
(the never-blank invariant, §5).

### 12.6 Counter advance (`fNewMaxRenban`, `:798`) — read carefully

`fNewMaxRenban := rcount + 1` is set from `rcount` of the **LAST (truck,part) row emitted**
in the read-out loop. Because the loop walks trucks outer / parts inner, the last row is the
last part on the **last non-empty truck**. For the normal case (every truck gets ≥1 part), the
last truck is `T-1`, so `rcount = seed3 + (T-1)` and `fNewMaxRenban = seed3 + T`. This is the
value persisted by `UPDATE_RenbanGroupCount` (`%.3d`, `:431`) so the **next** breakdown of the
group seeds from `seed3 + T`. EDGE: if the last truck is empty (no part landed on it — cannot
happen under the current capacity-filling math for T trucks all ≤ TotalLots, but a rebuild must
not assume it), `fNewMaxRenban` reflects the highest-numbered NON-empty truck +1. The rebuild
computes `next_count = max(rcount over emitted rows) + 1`, which equals `seed3 + (highest
non-empty TruckNumber) + 1` — faithful to the Pascal's "last emitted rcount + 1" for the
trucks-fill-in-order case, and robust to a trailing-empty truck.

### 12.7 Hazards the rebuild must handle (beyond faithful replication)

- **H7 div-by-zero:** `lotqty` 0/NULL → `lots = qty div 0` crashes the legacy `LoadScreen`
  (`:622`). Rebuild: **skip the row + raise/alert** (do NOT crash the whole group). (12.0)
- **varchar(3) counter rollover at 999 — EXACT reduction function (FIXED 2026-06-21,
  sql-adversary BLOCKER 1):** when `next_count >= 1000` the persisted group counter is NOT
  `next_count % 1000`. The legacy persists `Format('%.3d',[next_count])`, where `.3` is a
  *minimum* width that NEVER caps (`1002 → '1002'`, 4 chars), into `@RenbanCount varchar(3)`,
  which the proc **LEFT-TRUNCATES to the first 3 chars** (`'1002' → '100'`). So the reduction is
  **`str(N)[:3]`**, not `N % 1000`. Rebuild persists **`('%03d' % next_count)[:3]`**
  (`1002→'100'`, exact `1000→'100'`, `634→'634'`, `5→'005'`). **PROVEN on the live proc**
  (mssql-spike, rolled back): `EXEC UPDATE_RenbanGroupCount @RenbanCount='1002'` stored `'100'`;
  `@RenbanCount='002'` stored `'002'` (the old-bug value). For parallel-run parity the rebuild
  must persist what the legacy persists, so a side-by-side run shows zero diff.
  - **The renban NUMBER itself is UNAFFECTED:** `group_code + '%03d' % rcount` for rcount≥1000
    renders the full string (`CMWA1000`, `CMWA1001`) and `VC_RENBAN_NUMBER varchar(8)` holds 8
    chars — `CMWA1000` fits. Only the persisted `varchar(3)` COUNT truncates. (A 5-char group like
    `DICAS1000` is 9 chars and WOULD truncate at varchar(8), but identically on both sides — not a
    divergence.)
  - **ROLLOVER-LATENT-BUG CARRY (do NOT fix here — post-cutover, like the GetShip-calendar carry):**
    the wrap is itself a latent legacy bug. At `next_count >= 1000` the persisted count COLLAPSES
    (`1000→'100'`, `1002→'100'`), so the NEXT run of that group re-seeds from ~`100` and its renban
    numbers **COLLIDE** with the earlier `CMWA100x` block. We faithfully reproduce this for
    parallel-run parity; we must NOT silently "fix" the wrap in phase-1, or the rebuild would diverge
    from the legacy. **Post-cutover fix:** widen the count column/param (so it doesn't truncate), or
    **alert + block the operator at 999** before the counter can roll. Reachability is live: counts
    climb over operational time — today CMWA 288, DICAS 480/484, **PACF 633/634** actively climbing;
    any group reaching ~994+ with up to 6 trailers crosses 1000 in `next_count`. Tracked as an
    `# IG83-TODO:` in `renban/code.py` step (c).
- **Phase-B infinite loop:** bound it — if a full round-robin pass places nothing while
- **Phase-B infinite loop:** bound it — if a full round-robin pass places nothing while
  `remainder > 0`, raise (capacity gate should preclude this). (12.2)

### 12.8 Write-back (driver, unchanged from §6 / data-analysis §3) — for completeness

One transaction: per **EMITTED-row part** `DELETE_OrderRenban(part, @FRSNumber='',
@RenbanNumber='')` (deletes ALL still-blank rows of the part); then per **trailer-row with
qty>0** `INSERT_OpenOrder(@SupCode,@PartNum,@KanbanNum,@FRSNum=<7-char FRS>,@RenbanNum=<assigned>,
@Qty=<lots×lotqty>)` (skip qty=0, `:540`); then once `UPDATE_RenbanGroupCount(@RenbanCode,
@RenbanCount=('%03d' % next_count)[:3])` (see §12.7 — `str(N)[:3]`, NOT `% 1000`). The 7-char FRS
suffix is honored by the proc (recompute is a varchar(7) truncation no-op — PROVED, data-analysis
§3.2 + re-proved 2026-06-21: `'6090102'+'01'` → `'6090102'`, `'6090103'+max+1` → `'6090103'`,
both len 7). Do NOT resurrect the commented-out update-in-place path (`:482-539`).

- **DELETE SCOPE = the EMITTED parts, NOT the loaded feed (FIXED 2026-06-21, sql-adversary
  SHOULD-FIX 2):** the legacy commit loop iterates ONLY the emitted grid rows
  (`for i:=1 to fAvailableCount`, `:417`; `fAvailableCount = AvailableGrid.RowCount-1`, `:799`) and
  `NewFRSOrder` deletes by the EMITTED row's part (`@PartNumber := AvailableGrid.Cells[2,Row]`,
  `:506`). So a part with **`0 < qty < lotqty` → `lots = qty div lotqty = 0`** lands on no truck,
  emits no grid row, and is **NEVER deleted** — its blank-renban placeholder SURVIVES for a future
  breakdown (once its qty grows) or manual handling. The rebuild therefore derives the distinct
  delete-parts from the **emitted `rows`** (`compute` output), NOT the full loaded `orders`. Deriving
  from the loaded feed would `DELETE_OrderRenban` the `lots=0` part with no re-insert → **silent
  order loss** (and, because the blank renban was its only "needs grouping" marker, it would never
  ship and never error — H1). Verified live: a feed of a sub-lot part (qty 20 < lotqty 40) + a normal
  part → after commit the sub-lot placeholder STILL EXISTS with its qty intact, the normal part is
  grouped (`scripts/e2e/test_renban_e2e.py` partial-lot case).
