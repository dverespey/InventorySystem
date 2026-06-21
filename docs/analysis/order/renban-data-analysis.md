# Renban Breakdown — DATA-Layer Behavioral Spec (M2 unit 3, final)

Source-truth analysis of the **renban-grouping breakdown** `RenbanOrder.pas`
(`TGroupRenbanOrder_Form`, the live unit per `InventorySystem.dpr:32`). This is the
**middle** of the daily order chain: `Order.pas` commits blank-renban placeholder orders
for palletized parts, then **this** form breaks those placeholders into trailer-grouped
orders with assigned renban numbers, after which `OrderFormCreateF.pas` (the order-file
generator, see `order-file-data-analysis.md`) serializes them to supplier files.

> Read first: `legacy-order-spec.md` (Order worksheet + `INSERT_OpenOrder` server logic),
> `order-file-data-analysis.md` (the downstream `SELECT_OrderNotOrdered` feed + the
> `VC_RENBAN_NUMBER <> ''` gate), and the memory `project-order-renban-domain`
> (the INVERTED lot flag, blank-renban transient, R3 sum-all fixture bug).

Confidence: HIGH — every proc body read from `/tmp/inv_utf8.sql` (UTF-8 of the authoritative
`DB Schema/CreateInventory.sql`, 2026-06-12 dump); every data claim proved with a bounded
query against `Inventory_Live` (matched legacy snapshot, READ-ONLY); the full write-back
sequence executed end-to-end on the writable `Inventory` target inside a rolled-back
transaction. Pascal flow read from `RenbanOrder.pas` (all 896 lines).

Verification env: `mssql-spike`, READ-ONLY on `Inventory_Live`, rolled-back trans on `Inventory`.

---

## 0. The driving flow (what the form does, in order)

`RenbanOrder.pas` is a modal worksheet (`Execute`, `:348`). Workflow:

1. **`Execute` (:348)** fills `RenbanGroups_ComboBox` from `SELECT_RenbanGroup;1 @RebanCode=''`
   (all groups; reads `'Renban Group Code'`, `:367`).
2. **Pick a group** → `RenbanGroups_ComboBoxChange` → `LoadScreen` (:575) calls
   `SELECT_OrderNoRenban;1 @RenbanGroupCode` and fills the `AvailableGrid` with the
   **blank-renban** orders of that group (the placeholders `Order.pas` left), computing
   `Lots = IN_QTY div IN_1LOTQTY` per row and a `TotalLots` sum (:622-625).
3. **Enter # trailers + pallets/trailer**, click **FRS Breakdown** (`FRSBreakdown_ButtonClick`,
   :699) → in-memory `TGroupRenban`/`TTruck` objects distribute the lots across trucks
   (§4), then re-paint the grid with the **broken-down** rows: new FRS suffix + new renban
   number per truck (:746-795). Nothing is written to the DB yet.
4. **Click Create Order** (`CreateOrder_ButtonClick`, :406): inside one
   `Inv_Connection.BeginTrans…CommitTrans` (rollback on exception, :473), for each grid row
   run `NewFRSOrder` (:481) = **`DELETE_OrderRenban`** (kill the blank placeholder) +
   **`INSERT_OpenOrder`** (re-insert with the assigned renban + FRS-with-trailer); then bump
   the group's roll-over counter via **`UPDATE_RenbanGroupCount`** (:426-433).

The **only DB writes** are: `DELETE_OrderRenban` (delete blank placeholders), `INSERT_OpenOrder`
(re-insert assigned rows), `UPDATE_RenbanGroupCount` (advance the group counter). The grid +
the `TGroupRenban` truck objects are the in-memory product that decides the renban/FRS/qty.

Procs called (all on `Inv_StoredProc`, `Inv_Connection`):

| # | Proc | Where | Params (Pascal sends) | Real proc param names |
|---|---|---|---|---|
| 1 | `SELECT_RenbanGroup;1` | `:359` | `@RebanCode` | `@RenbanCode` (Pascal typo "Reban", positional bind) |
| 2 | `SELECT_OrderNoRenban;1` | `:588` | `@RenbanGroupCode` | `@RenbanGroupCode` |
| 3 | `DELETE_OrderRenban;1` | `:503` | `@PartNumber,@FRSNumber='',@RenbanNumber=''` | same |
| 4 | `INSERT_OpenOrder;1` | `:543` | `@SupCode,@PartNum,@KanbanNum,@FRSNumber,@RenbanNum,@Qty` | `@FRSNum` (Pascal sends `@FRSNumber`, positional bind) |
| 5 | `UPDATE_RenbanGroupCount;1` | `:426` | `@RenbanCode,@RenbanCount` | same |

> Param-name traps (positional binding saves them, a Named-Query rebuild must use the real
> names): `SELECT_RenbanGroup` proc declares **`@RenbanCode`** but Pascal adds **`@RebanCode`**
> (`:361`, missing the "n"); `INSERT_OpenOrder` proc declares **`@FRSNum`** but Pascal adds
> **`@FRSNumber`** (`:551`). Proved: `EXEC INSERT_OpenOrder @FRSNumber=...` fails
> ("expects parameter '@FRSNum'"); `@FRSNum=...` succeeds.

---

## 1. `INV_RENBAN_GROUP_MST` — the renban group master

**DDL** (`/tmp/inv_utf8.sql:3870-3891`):

```sql
CREATE TABLE [dbo].[INV_RENBAN_GROUP_MST](
    [IN_RENBAN_ID]            int IDENTITY(1,1) NOT NULL,   -- PK
    [VC_RENBAN_GROUP_CODE]    varchar(5)  NOT NULL,         -- UNIQUE (IX_INV_RENBAN_GROUP_MST)
    [VC_RENBAN_GROUP_COUNT]   varchar(3)  NULL,             -- the roll-over trailer counter
    [IN_SHIP_DAYS]            int NULL,                     -- generic ship-day lead
    [IN_SHIP_DAYS_MONDAY..SATURDAY] int NULL,               -- per-weekday ship-day lead
    [VC_LAST_UPDATE]          varchar(16) NULL,
    [VC_ADD]                  varchar(16) NULL,
 PK CLUSTERED (IN_RENBAN_ID),
 UNIQUE NONCLUSTERED (VC_RENBAN_GROUP_CODE) )
```

- **PK = `IN_RENBAN_ID`** (identity). **Unique key = `VC_RENBAN_GROUP_CODE`** (the business
  key the parts-stock join + the renban-number prefix use).
- **There is NO `IN_SHIP_DAYS_SUNDAY` column** — Saturday is the last per-weekday column.
  (Same shape as `INV_PARTS_STOCK_MST`; `SELECT_PartShipDays` only returns Mon..Sat + generic.)
- `INV_PARTS_STOCK_MST` links a part to a group via **`IN_RENBAN_ID`** (FK to the PK).
  **There is NO `VC_RENBAN_GROUP_CODE`/`_COUNT` column on the parts table** — the group code
  and count come from the `INV_RENBAN_GROUP_MST r` join in `SELECT_OrderNoRenban`
  (`:6352-6354`). (Earlier dead-code comments reference `p.VC_RENBAN_GROUP_CODE`, `:6361`, but
  that join form is commented out and the column does not exist; the live join is on
  `IN_RENBAN_ID`.)

### Proof on live data (`Inventory_Live`)

5 groups, code is unique (0 dup rows). Full master:

```
IN_RENBAN_ID  CODE   COUNT  Ship  Mon Tue Wed Thu Fri Sat
12            CAP    068    0     10  10  10  10  10  0
11            CMWA   297    4     4   4   4   4   4   4
8             DICAS  484    5     5   5   5   5   5   5
9             HCAP   088    10    10  10  10  10  10  10
7             PACF   634    0     13  13  13  13  13  0
```

- `VC_RENBAN_GROUP_COUNT` is a **3-char zero-padded string** (`'068'`,`'297'`,`'484'`,`'088'`,
  `'634'`), all `ISNUMERIC=1`. (`Inventory` target shows the same groups, counts 068/288/480/088/633.)
- `RenbanGroupMaster.pas` (`TRenbanGroupMaster_Form`, `dpr:31`) is the CRUD master for these
  rows (count + ship-days editor). Per memory `project-order-renban-domain`: count `000` is
  valid (count exists only for trailer-uniqueness/roll-over, not a quantity).

### 1.1 The ship-day override is REAL and consumed by the order-file gen — CONFIRMED

`SELECT_PartShipDays` (`/tmp/inv_utf8.sql:4783-4805`, called by the order-file generator,
see `order-file-data-analysis.md` §3.3) branches on the part's `IN_RENBAN_ID`:

```sql
SELECT @gc = IN_RENBAN_ID FROM INV_PARTS_STOCK_MST WHERE VC_PART_NUMBER = @PartNumber
if @gc is null  -> return the PART's own IN_SHIP_DAYS* (INV_PARTS_STOCK_MST)
else            -> return the GROUP's IN_SHIP_DAYS* (INV_RENBAN_GROUP_MST g JOIN ... p)
```

So **a part in a renban group inherits the GROUP's ship-day lead, overriding its own**
(`:4798-4805`). E.g. every CMWA part ships on the group's `4/4/4/4/4/4` lead, not the part's.
A rebuild's ship-date calc for a grouped part must read the group's ship-days, not the part's.

---

## 2. `SELECT_OrderNoRenban` — the load feed (which orders are "available to group")

**Body** (`/tmp/inv_utf8.sql:6345-6363`):

```sql
CREATE PROCEDURE [dbo].[SELECT_OrderNoRenban] @RenbanGroupCode varchar(5) AS
    SELECT *
    FROM INV_OPEN_ORDER_INF o
        JOIN INV_PARTS_STOCK_MST p   ON o.VC_PART_NUMBER = p.VC_PART_NUMBER
        JOIN INV_RENBAN_GROUP_MST r  ON r.IN_RENBAN_ID = p.IN_RENBAN_ID
                                    AND r.VC_RENBAN_GROUP_CODE = @RenbanGroupCode
    WHERE o.VC_RENBAN_NUMBER = ''
```

### 2.1 Selection = blank-renban orders of one group

- **`WHERE o.VC_RENBAN_NUMBER = ''`** — the blank renban is the **selection flag**. Only the
  transient placeholders `Order.pas` left (palletized parts → `@RenbanNum=''` at commit, per
  `legacy-order-spec.md` §6 / memory) are pulled. Lot-sized parts (born with a renban) never appear.
- The group filter is on the **renban-group join** (`r.VC_RENBAN_GROUP_CODE = @RenbanGroupCode`),
  reached through `p.IN_RENBAN_ID = r.IN_RENBAN_ID`. All three joins are INNER, so a part with
  `IN_RENBAN_ID = NULL` (no group) can never appear — correct, ungrouped parts are not grouped here.
- **Proof:** on `Inventory_Live`, `EXEC SELECT_OrderNoRenban @RenbanGroupCode='CMWA'` returns
  **0 rows** — the snapshot is post-breakdown (every CMWA order already has an assigned renban).
  This is the expected steady state (matches the order-file gen's "nothing pending" proof).

### 2.2 Cardinality + the `SELECT *` duplicate-column trap (same as order-file gen)

`SELECT *` over the 3-table join. The columns the form reads via `fieldbyname` (= **first
match**), with the ordinal proved by `sys.dm_exec_describe_first_result_set`:

| `fieldbyname(...)` (`:615-624`) | ordinal | source | the shadowed dup |
|---|---|---|---|
| `VC_KANBAN_NUMBER` | **20** (`o.`) | open-order kanban | 33 (`p.`) |
| `VC_SUPPLIER_CODE` | **2** (`o.`) | open-order supplier | (p has none) |
| `VC_PART_NUMBER` | **3** (`o.`) | part number | 30 (`p.`) |
| `VC_FRS_NUMBER` | 4 (`o.`) | FRS number | — |
| **`IN_QTY`** | **6 (`o.`) = ORDER qty** | the placeholder order qty | **42 (`p.`) = on-hand stock** |
| `IN_1LOTQTY` | 41 (`p.`) | 1-lot/pallet qty | — (only on p) |
| `VC_RENBAN_GROUP_CODE` | 57 (`r.`) | group code (for the renban prefix) | — (only on r) |
| `VC_RENBAN_GROUP_COUNT` | 58 (`r.`) | group counter (renban suffix seed) | — (only on r) |

> **Critical `IN_QTY` dup trap (identical to the order-file gen).** Ordinal 6 =
> `INV_OPEN_ORDER_INF.IN_QTY` (the placeholder order qty); ordinal 42 =
> `INV_PARTS_STOCK_MST.IN_QTY` (on-hand stock). `fieldbyname('IN_QTY')` takes ordinal 6
> (correct). A rebuild MUST alias explicitly (`o.IN_QTY AS order_qty`) and never grab "the
> IN_QTY column" — on-hand stock for these parts is 13341/28133/44418 vs order qtys 40/400/1200.
> Same hazard for `VC_KANBAN_NUMBER`/`VC_PART_NUMBER` (duplicated p vs o).

### 2.3 The Lots derivation (`Lots = IN_QTY div IN_1LOTQTY`)

Per loaded row (`:620-625`):
- grid col 4 = `IN_QTY` (order qty), col 5 = `IN_1LOTQTY` (lot/pallet qty),
- **grid col 6 = `Lots = StrToInt(IN_QTY) div StrToInt(IN_1LOTQTY)`** (integer division),
- col 7 = `VC_RENBAN_GROUP_CODE + VC_RENBAN_GROUP_COUNT` (e.g. `CMWA297` — the seed renban),
- `TotalLots_Edit` = running sum of Lots.

> **Div-by-zero risk:** `IN_QTY div IN_1LOTQTY` raises if `IN_1LOTQTY` is 0/NULL. **Proved
> safe on live data** (0 grouped parts have `IN_1LOTQTY = 0 or NULL`). A rebuild should guard
> it anyway (a misconfigured palletized part with no lot qty would crash the load).

---

## 3. The write-back — DELETE_OrderRenban + INSERT_OpenOrder + UPDATE_RenbanGroupCount

The write-back IS the renban assignment. It runs once per broken-down grid row (`NewFRSOrder`)
plus one counter bump, all in one transaction (`CreateOrder_ButtonClick`).

### 3.1 `DELETE_OrderRenban` — kill the blank placeholder

**Body** (`/tmp/inv_utf8.sql:7629-7647`):

```sql
CREATE PROCEDURE [dbo].[DELETE_OrderRenban] @PartNumber varchar(12),@FRSNumber varchar(7),@RenbanNumber varchar(8) AS
    IF @FRSNumber = ''
        DELETE FROM INV_OPEN_ORDER_INF WHERE VC_PART_NUMBER=@PartNumber AND VC_RENBAN_NUMBER=@RenbanNumber
    ELSE
        DELETE FROM INV_OPEN_ORDER_INF WHERE VC_PART_NUMBER=@PartNumber AND VC_FRS_NUMBER=@FRSNumber AND VC_RENBAN_NUMBER=@RenbanNumber
```

`NewFRSOrder` always calls it with **`@FRSNumber=''`, `@RenbanNumber=''`** (`:508-510`) → the
`IF @FRSNumber=''` branch → **deletes ALL blank-renban rows for that part** (across all FRS).
This is the "delete-supersede" lifecycle: it wipes every placeholder for the part, then the
truck-distributed re-inserts rebuild that part's rows with assigned renbans.

> **HAZARD — delete scope is part-wide (all blank-renban rows of the part), not per-FRS.**
> `@FRSNumber=''` deliberately ignores FRS so a part whose placeholders span several FRS days
> gets all of them cleared. Faithful because the FRS Breakdown step has already captured every
> lot of that part into the in-memory trucks before any delete. A rebuild that deletes
> per-(part,FRS) would orphan placeholders the truck objects already consumed. The dead
> alternate code (`:516-536`, `UPDATE_OrderRenbanQty`) is commented out — the live path is
> always delete-then-insert (the comment `:489-492` explains why: a single-item non-first
> truck could leave the initial order without a renban, so they delete everything and re-add).

**Stock neutrality — PROVED.** The DELETE trigger `DELETE_RecConfStatPartsStockMstQTY`
(`Inventory`, body read) only de-bumps `INV_PARTS_STOCK_MST.IN_QTY` when the deleted row had
`VC_STATUS_SUPPLIER_SHIPPING <> ''` (or arrival/yard/warehouse) AND the supplier's
`VC_INVENTORY_ADD_POINT` matches. A blank placeholder has all-empty statuses → **no stock
change on delete**. The INSERT trigger `INSERT_RecConfStatPartsStockMstQTY`
(`/tmp/inv_utf8.sql:7492-7521`) similarly only bumps stock when the inserted row has
`VC_STATUS_SUPPLIER_SHIPPING <> ''` → a freshly grouped order (empty shipping) → **no stock
bump on insert** (it DOES always copy `inserted` into the `INV_OPEN_ORDER_INF_HIST` heap).
Net: **the entire delete-and-reinsert renban breakdown is inventory-neutral.** Proved by the
end-to-end test (§3.4): on-hand 44418/28133 unchanged before vs after the full sequence.

### 3.2 `INSERT_OpenOrder` — re-insert with the assigned renban (the `@RenbanNum <> ''` branch)

Body at `/tmp/inv_utf8.sql:7358-7456` (full FRS/renban server logic in `legacy-order-spec.md`
§6). For the renban breakdown, Pascal sends a **non-empty `@RenbanNum`** (the truck's
`GroupCode + %.3d`, `:553-554`) → the **`@RenbanNum <> ''` branch** (`:7388`):

- `@FRSDate` = 4-digit year (rolled to next year if the FRS month-digit ≠ today's) + the FRS
  MMDD digits (`:7378-7386`).
- MaxFRS scope = **this part only**: `WHERE VC_FRS_NUMBER LIKE @FRSNum+'%' AND VC_PART_NUMBER=@PartNum`
  (`:7390-7393`). (Contrast the `@RenbanNum=''` branch which scopes across the whole renban
  group — that branch is NOT taken here.)
- Suffix: if MaxFRS null → keep `@FRSNum`+`'01'`; else `@FRSNum + (right(MaxFRS,2)+1)`.

> **CRITICAL — the proc's FRS-suffix recompute is a NO-OP for the renban breakdown (varchar(7)
> truncation).** Pascal sends the **full 7-char** `@FRSNum` (5-char prefix + 2-digit trailer,
> e.g. `'6090102'`, built at `:763-767`). `@FRSNum` is declared `varchar(7)`. The recompute
> `SET @FRSNum = @FRSNum + '02'` produces a 9-char string that is **silently truncated back to
> 7 chars** = the original value. So **the renban breakdown's FRS suffix is exactly what Pascal
> supplied** (`TruckNumber+1`), not the proc's MaxFRS+1. PROVED:
> `DECLARE @F varchar(7)='6090201'; SET @F=@F+'02';` → `@F='6090201'` (len 7). With a 5-char
> seed (`'60902'`) the same recompute yields `'6090202'` (len 7) — so the recompute is ONLY
> live for the original 5-char Order path. A rebuild MUST honor Pascal's trailer suffix for the
> grouped re-insert; do NOT reimplement "max+1" here or you will diverge from the truck layout.

The FRS trailer suffix Pascal builds (`:763-767`): `copy(frs,1,5)` + (`TruckNumber > 8` ?
`IntToStr(TruckNumber+1)` : `'0'+IntToStr(TruckNumber+1)`). TruckNumber is 0-based; truck 0 →
`'01'`, truck 7 → `'08'`, truck 8 → `'09'`, truck 9 → `'10'`. The renban suffix
(`:775-779`): `rcount = StrToInt(right(seedRenban,3)) + TruckNumber`; renban =
`RenbanGroups_ComboBox.Text + %.3d(rcount)`.

### 3.3 `UPDATE_RenbanGroupCount` — advance the roll-over counter

**Body** (`/tmp/inv_utf8.sql:5142-5157`):

```sql
CREATE PROCEDURE [dbo].[UPDATE_RenbanGroupCount] @RenbanCode varchar(5),@RenbanCount varchar(3) AS
    ... @Updated = yyyymmddhhmmss ...
    UPDATE INV_RENBAN_GROUP_MST SET VC_RENBAN_GROUP_COUNT=@RenbanCount, VC_LAST_UPDATE=@Updated
    WHERE VC_RENBAN_GROUP_CODE = @RenbanCode
```

Called once after all rows (`:426-433`) with `@RenbanCount = Format('%.3d',[fNewMaxRenban])`,
where `fNewMaxRenban = rcount + 1` (`:798`) = the last-used renban suffix + 1. So the next
breakdown run seeds renbans from where this one stopped. `@RenbanCount` is `varchar(3)` ⇒ a
value ≥1000 truncates → this is the **roll-over at 999** (counter wraps within 3 digits). The
count is purely for trailer-renban uniqueness, not a quantity (memory: `000` valid).

### 3.4 End-to-end write-back — PROVED on the live `Inventory` target (rolled back)

Executed the exact form sequence on `Inventory` inside `BEGIN TRAN…ROLLBACK`:
seed two blank CMWA placeholders (Q8000 qty 400, Q5100 qty 1200, FRS `6090101`, renban `''`) →
`DELETE_OrderRenban @FRS='' @Renban=''` (×2) → `INSERT_OpenOrder @RenbanNum='CMWA289'
@FRSNum='6090101'` (×2) → `UPDATE_RenbanGroupCount @RenbanCode='CMWA' @RenbanCount='290'`.

Result (then rolled back, group count restored 288→288):
```
final_rows        4261102Q5100  6090101  20260901  CMWA289  1200
final_rows        4261102Q8000  6090101  20260901  CMWA289  400
remaining_blank   0                                         -- placeholders gone
stock_after       4261102Q5100  28133                       -- == stock_before (NEUTRAL)
stock_after       4261102Q8000  44418                       -- == stock_before (NEUTRAL)
group_count_after 290                                       -- counter advanced
```
And the eligibility gate for the downstream order-file gen:
```
eligible_for_orderfile  1   -- VC_ORDER_DATE='' AND VC_RENBAN_NUMBER<>''  → passes SELECT_OrderNotOrdered
```

So the write-back: (a) removes the blank placeholders, (b) creates assigned-renban rows that
**immediately satisfy `SELECT_OrderNotOrdered`'s `VC_RENBAN_NUMBER <> ''` gate**
(`order-file-data-analysis.md` §2.2), (c) advances the group counter, (d) **leaves stock
untouched** (both triggers are status-gated and the placeholders are status-empty).

---

## 4. The lot distribution algorithm (`TGroupRenban`/`TTruck`) — the in-memory breakdown

The actual "which lots go on which truck" logic is pure Pascal (no DB). Documented because a
faithful rebuild must reproduce its quirks.

- **`SetTrucks(NumberOfTrucks, SizeOfTruck)`** (`:239`): creates N `TTruck`, each `.Size =
  pallets/trailer` (`TrailerPalletCount_Edit`). Trailers/pallet auto-fill each other
  (`:867-893`): `pallets = TotalLots div trailers` and vice-versa (integer division).
- **`TGroupRenban.AddOrder`** (`:255-301`), per part: spread `lots` across trucks.
  - **Even part:** `lots div trucks` to each truck, with a `leftover` carry when a truck is
    full (`CurrentCount + share > Size`); the overflow rolls to the next truck (`:262-277`).
  - **Remainder:** `lots mod trucks` (+ leftover) distributed one-lot-at-a-time round-robin
    over trucks with room (`:279-300`).
- **`TTruck.AddOrder`** (`:155-182`): **consolidates lots per part within a truck** —
  `fOrderList.IndexOf(partnumber)`; if the part already exists on that truck, it **adds the
  lots to the existing entry** (`:160-163`) rather than creating a second line. So one truck
  carries at most one line per part (qty summed). This is the trailer-level analog of the **R3
  "sum-all" faithfulness** — the breakdown SUMS lots per part, it does not split a part into
  duplicate lines on the same truck. Do NOT "fix" this into per-lot lines.
- **Final qty per grid row** (`:770`): `IN_QTY = lots × lotqty`. So the written `IN_QTY` is
  **always an exact lot multiple** — PROVED on live: all 1957 CMWA order rows have
  `IN_QTY % IN_1LOTQTY = 0` (1957/1957 evenly divisible, 0 exceptions). No rounding ever
  applies to the order qty; the lot multiple IS the rounding unit (set upstream at Order time).
- **Capacity gate** (`:709`): the breakdown only proceeds if `trucks × pallets ≥ TotalLots`,
  else "will not fit" and aborts. So a trailer is never over-filled past `Size` pallets.

> **Known bug noted in the code (`:489-492`):** with the new breakdown, "only one item order
> can be put on any FRS truck — if it's not the first truck the initial order is left without a
> renban." The fix they shipped is the unconditional delete-then-reinsert (§3.1). A rebuild
> that re-implements an in-place UPDATE path (the commented-out `:516-536`) would reintroduce
> this missing-renban bug. Keep delete-then-insert.

---

## 5. Lot-sizing — the INVERTED flag (proved) + the lot multiple + sum-all

`INV_PARTS_STOCK_MST` lot/pallet columns (`/tmp/inv_utf8.sql:236-238`):
`IN_1LOTQTY int` (the 1-lot / 1-pallet qty), `IN_QTY int` (on-hand stock — NOT order qty),
`BIT_LOT_SIZE_ORDERS bit`, `IN_RENBAN_COUNT int` (per-part renban counter, used by the
non-grouped Order path), `IN_RENBAN_ID int` (FK to the group).

### 5.1 The INVERTED flag — PROVED on live data

`BIT_LOT_SIZE_ORDERS` is stored **inverted**: **0 = lot-sized TRUE** (one lot fills one
trailer = one renban; Order form emits the renban directly), **1 = palletized = lot-sized
FALSE** (one lot is a portion of a trailer; needs THIS renban breakdown). Proof
(`Inventory_Live`, distribution of the flag by group membership):

```
in-group   BIT_LOT_SIZE_ORDERS=1   21 parts   -- ALL grouped parts are flag=1 (palletized)
no-group   BIT_LOT_SIZE_ORDERS=0   11 parts   -- lot-sized-TRUE parts, no group
no-group   BIT_LOT_SIZE_ORDERS=1   15 parts   -- palletized-but-not-yet-grouped / other
```

**Every part with a renban group has flag=1, and 0 grouped parts have flag=0** — i.e. the
palletized (flag=1) parts are exactly the ones that get renban-grouped, confirming the inversion.
A rebuild reading `BIT_LOT_SIZE_ORDERS` MUST treat `0` as "lot-sized" and `1` as "palletized"
(invert before any "is this lot-sized?" test). Any naive `if BIT_LOT_SIZE_ORDERS then …`
inverts the behavior of the entire order path.

### 5.2 The lot multiple + rounding

- The lot/pallet qty is **`IN_1LOTQTY`** (the worksheet's column O; there is **no separate
  "pallet qty" column** — for palletized parts `IN_1LOTQTY` IS the per-pallet qty).
- Order qty is **always a whole number of lots × `IN_1LOTQTY`** — the "rounding to the lot" is
  done at Order-worksheet entry time (the specialist types a lot count into col R, which
  computes Q = R × O; see memory + `legacy-order-spec.md` §6). By the time this breakdown runs,
  `IN_QTY` is already an exact lot multiple. PROVED: 1957/1957 CMWA order rows divisible.
- This breakdown does **NOT** re-round: it only redistributes whole lots across trucks
  (`lots = IN_QTY div IN_1LOTQTY`) and recomputes `IN_QTY = lots × IN_1LOTQTY` per truck.

### 5.3 The R3 "sum-all" learning (faithful — do not "fix")

Per memory `project-order-renban-domain` / `feedback-parity-fixture-fidelity`: the receipt
projection (`SELECT_OrderOpenOrderList`/`PutOpenOrderCount`) SUMS ALL rows by FRS-date with no
renban filter, and that is **already faithful** — the golden looks "renban-only" only because
this breakdown DELETES the placeholder rows and re-inserts grouped rows, so by Order-Start only
the grouped (CMWA) rows exist. The R3 overcount was a **fixture bug** (8 synthetic `SPIKEFX`
blank-renban rows injected on top of the real CMWA rows), NOT a proc-math bug. The sum-all
appears here too: `TTruck.AddOrder` sums lots per part on a truck (§4). A rebuild must
reproduce the delete-supersede lifecycle (so no stray blank-renban rows survive to be summed)
rather than adding a renban filter to the receipt math.

---

## 6. Palletization — how palletized parts group into renban by pallet

There is **no dedicated pallet table or pallet-config column** beyond `IN_1LOTQTY`. The
palletization model is entirely: **`IN_1LOTQTY` = 1 pallet of the part; the renban group
(`IN_RENBAN_ID` → `INV_RENBAN_GROUP_MST`) = "these parts ship together on one trailer"; the
trailer's pallet capacity is a per-run user input** (`TrailerPalletCount_Edit`), not stored.

So palletized parts group into renban as: all parts sharing one `IN_RENBAN_ID` are pooled, the
user says "N trailers × P pallets each", and `TGroupRenban` packs whole lots (= pallets) into
the N trucks (§4). Each truck becomes one renban number (`GroupCode + %.3d`) and one FRS
(`prefix + trailer suffix`); a truck carries multiple part numbers.

### Proof on live data (`Inventory_Live`)

- Parts per group (all flag=1 / palletized), with their pallet qty (`IN_1LOTQTY`):
  CAP→1 part (lotqty 390), CMWA→5 parts (40/40/30/30/40), DICAS→6 parts (25-30),
  HCAP→1 (125), PACF→8 (150-10000). 21 grouped parts total.
- **One renban = one trailer carrying multiple parts.** E.g. renban `CMWA334` (FRS `4071001`)
  carries 3 distinct parts: `4261102Q4100` qty 40 (1×40), `4261102Q5100` qty 1200 (40×30),
  `4261102Q8000` qty 360 (9×40). All on the same FRS, all exact lot multiples.
- **One renban group = exactly one supplier** (no cross-supplier trailers): CMWA→`0572B`,
  DICAS→`07100`, PACF→`11111`, CAP+HCAP→`10011` (proved: `COUNT(DISTINCT IN_SUPPLIER_ID)=1`
  per group). So a renban trailer is always single-supplier.
- Renban-number format = **`<GROUP_CODE><%.3d count>`**, e.g. `CMWA100`..`CMWA999`,
  `DICAS143`..`DICAS483`. `DICAS` (5-char code) + 3 digits = 8 chars = full `varchar(8)`.
  (Lot-sized non-grouped parts instead use `<Kanban><%.3d>` and `IN_RENBAN_ID=NULL`, e.g.
  renban `072` on `4265202R8000` flag=0 — a different, shorter format.)

---

## 7. col R → Q → FRS mapping (proved on real rows)

The worksheet columns (memory `project-order-renban-domain`, `legacy-order-spec.md` §1/§6):
- **R (LotCol=18)** = the integer the specialist types = **number of lots** to order.
- **Q (QtyCol=17)** = computed **order qty = R × O** where **O = `IN_1LOTQTY`** (lot/pallet qty).
- **FRS** = the release/trailer identifier written to **`INV_OPEN_ORDER_INF.VC_FRS_NUMBER`**.

The transform across the chain (all proved):

| Stage | R (lots) | Q (`IN_QTY`) | FRS (`VC_FRS_NUMBER`) | Renban |
|---|---|---|---|---|
| Order worksheet | user types lots into R | `Q = R × IN_1LOTQTY` | (leadtime+1)-day cell | — |
| Order commit (`INSERT_OpenOrder`) | — | `IN_QTY = Q` (palletized: one row) | `copy(yymmdd,2,5)` + `01/02…` (5+2 chars) | `''` (BLANK for grouped parts) |
| **THIS breakdown** | `lots = IN_QTY div IN_1LOTQTY`, redistributed across trucks | `IN_QTY = lots(per truck) × IN_1LOTQTY` | `copy(frs,1,5)` + `(TruckNumber+1)` 2-digit | `GroupCode + %.3d(count+TruckNumber)` |
| Order-file gen | — | copies `IN_QTY` verbatim (`%.5d`) | copies FRS (5+2) | copies renban (`%8s`) |

- **FRS format = 7 chars = `Y MMDD TT`**: 1-digit year + MMDD + 2-digit trailer suffix.
  PROVED: CMWA rows `4070901`,`4071001`,… = `4`(year)+`0709`/`0710`(MMDD)+`01`(trailer).
  Trailer suffix distribution for CMWA: `01`×1153, `02`×798, `03`×6 (most days 1-2 trailers).
- **Q = R × O, exact.** PROVED: every grouped order row's `IN_QTY` is divisible by
  `IN_1LOTQTY` (e.g. `CMWA334`: 40=1×40, 1200=40×30, 360=9×40).
- **R → Q is multiply-by-lot; Q → FRS is the trailer-suffix increment.** The breakdown takes
  the placeholder's (R,Q,FRS) and re-derives per-truck (R',Q',FRS') by repacking lots.

---

## 8. Forecast → order → renban chain consistency (proved)

The breakdown reads what Order committed and writes what the order-file gen reads. Supplier/part
keys align across the whole chain.

- **Reads** the blank-renban `INV_OPEN_ORDER_INF` rows `Order.pas` produced (via
  `INSERT_OpenOrder @RenbanNum=''` for palletized parts) — joined through the part's
  `IN_RENBAN_ID` to the group. (The breakdown never touches `INV_BREAKDOWN_FC_INF`; the
  forecast→order qty sizing happened earlier in the Order worksheet.)
- **Writes** assigned-renban rows that pass `SELECT_OrderNotOrdered`'s
  `VC_RENBAN_NUMBER <> '' AND VC_ORDER_DATE empty` gate (PROVED, §3.4) → the order-file gen
  picks them up.
- **Supplier consistency for grouped parts (PROVED, `Inventory_Live`):** all 3111 grouped
  open-order rows have `open-order.VC_SUPPLIER_CODE == part-master supplier` (3111/3111, **0
  mismatch**). Each renban group maps to one supplier (§6). So a part the forecast importer
  keyed under supplier S → committed under S → grouped into S's single-supplier trailer →
  emitted in S's order file. The chain is supplier-consistent end to end (matches the
  order-file gen's 4284/4284 proof).

---

## 9. Hazards (what will silently break a renban-breakdown rebuild)

1. **INVERTED `BIT_LOT_SIZE_ORDERS` (0 = lot-sized TRUE, 1 = palletized).** Proved: all 21
   grouped parts are flag=1, 0 grouped parts flag=0. A rebuild that reads the flag literally
   inverts which parts get renban-grouped vs renban-at-commit. Always invert. (§5.1)
2. **The blank renban (`VC_RENBAN_NUMBER = ''`) is the selection flag AND the eligibility
   gate.** `SELECT_OrderNoRenban` pulls only `=''`; `SELECT_OrderNotOrdered` (downstream)
   emits only `<>''`. The breakdown's job is to flip blank → assigned. Emit a non-blank renban
   or the order never ships; leave it blank and it never groups. (§2.1, §3.4)
3. **R3 sum-all is FAITHFUL — do not add a renban filter.** `TTruck.AddOrder` sums lots per
   part on a truck; the receipt projection sums all rows. The "renban-only" look comes from the
   delete-supersede lifecycle, not a filter. Reproduce delete-then-reinsert; don't "fix" the
   sum. (§4, §5.3)
4. **`DELETE_OrderRenban @FRS='' @Renban=''` deletes ALL blank rows of the part (part-wide, not
   per-FRS).** Faithful only if the FRS Breakdown has already captured every lot into the
   in-memory trucks before the delete. A per-(part,FRS) delete would orphan placeholders. (§3.1)
5. **FRS-suffix recompute is a NO-OP for the breakdown (varchar(7) truncation).** Pascal sends
   the full 7-char FRS; the proc's `@FRSNum + max+1` truncates back to 7 chars = Pascal's value.
   Honor Pascal's `TruckNumber+1` suffix; do NOT reimplement "max+1" for grouped re-inserts (it
   IS live for the 5-char original Order path). PROVED both ways. (§3.2)
6. **`SELECT *` duplicate `IN_QTY` (ord 6 order-qty vs ord 42 on-hand stock).** `fieldbyname`
   takes ord 6 (correct); a rebuild must alias `o.IN_QTY AS order_qty`, never grab "the IN_QTY
   column". On-hand for these parts is 13341/28133/44418 vs order qtys 40/400/1200. Same for
   duplicated `VC_KANBAN_NUMBER`/`VC_PART_NUMBER`. (§2.2)
7. **Stock neutrality depends on status-empty placeholders.** Both INSERT and DELETE triggers
   are gated on `VC_STATUS_SUPPLIER_SHIPPING <> ''` (+ `VC_INVENTORY_ADD_POINT`). The breakdown
   is inventory-neutral ONLY because its rows are status-empty (PROVED). A rebuild that sets any
   shipping/arrival status on the re-inserted rows would (wrongly) bump stock. The INSERT
   trigger ALSO always writes to the `INV_OPEN_ORDER_INF_HIST` heap (no PK) — placeholder
   deletes leave their HIST rows behind (HIST accumulates). (§3.1)
8. **Renban-number length / count roll-over.** `<GROUP_CODE(≤5)> + %.3d` = up to 8 chars =
   exactly `varchar(8)` (DICAS+3 = 8, no headroom). `VC_RENBAN_GROUP_COUNT` is `varchar(3)` so
   it rolls over at 999 (truncation). The count is a uniqueness seed, not a quantity; 000 is
   valid. Reproduce the 3-digit zero-pad + wrap. (§1, §3.3, §6)
9. **Div-by-zero on `IN_QTY div IN_1LOTQTY`.** Currently safe (no grouped part has lotqty
   0/NULL) but a misconfigured palletized part with no lot qty crashes `LoadScreen` (`:622`).
   Guard it. (§2.3)
10. **Ship-day override lives on the GROUP.** `SELECT_PartShipDays` returns the group's
    ship-days for grouped parts (overriding the part's own). A rebuild's ship-date calc for a
    grouped part must read `INV_RENBAN_GROUP_MST.IN_SHIP_DAYS*`, not the part's. There is no
    Sunday column (Mon..Sat + generic only). (§1.1)
11. **Multi-site gap.** No `site_id` on `INV_RENBAN_GROUP_MST`, `INV_PARTS_STOCK_MST`, or
    `INV_OPEN_ORDER_INF`. The renban group code is global; a multi-site rebuild must scope the
    group + its counter per site (same gap as Order/EDI/order-file layers).
12. **Whole write-back in one transaction; the FRS Breakdown is in-memory only.** If the commit
    fails, the grid's broken-down state is lost (the `TGroupRenban` is freed on close); the user
    must re-enter trailers/pallets and re-break. No partial-write resume. (§0, :416-475)

---

## 10. What a faithful renban-breakdown rebuild MUST read / write

**Read:**
1. `SELECT_RenbanGroup` equiv → list groups (code + count + ship-days).
2. `SELECT_OrderNoRenban(@RenbanGroupCode)` equiv: `INV_OPEN_ORDER_INF o` JOIN parts-stock JOIN
   renban-group (on `IN_RENBAN_ID`), `WHERE o.VC_RENBAN_NUMBER = ''`. Project **aliased**:
   `o.VC_KANBAN_NUMBER, o.VC_SUPPLIER_CODE, o.VC_PART_NUMBER, o.VC_FRS_NUMBER, o.IN_QTY AS
   order_qty, p.IN_1LOTQTY, r.VC_RENBAN_GROUP_CODE, r.VC_RENBAN_GROUP_COUNT` — never `SELECT *`.
   Compute `lots = order_qty div IN_1LOTQTY` (guard /0).

**Compute (in-memory):** distribute whole lots across `trucks × pallets` (capacity-gated,
remainder round-robin, lots-per-part-per-truck SUMMED). Per truck: `IN_QTY = lots × IN_1LOTQTY`;
`FRS = first5(FRS) + zeropad2(TruckNumber+1)` (truck>8 → 2 raw digits); `renban = GroupCode +
%.3d(GroupCount + TruckNumber)`; `next_count = last_renban_suffix + 1`.

**Write (one transaction, per part then once):**
1. `DELETE_OrderRenban(@PartNumber, @FRSNumber='', @RenbanNumber='')` — clear ALL blank
   placeholders of the part.
2. `INSERT_OpenOrder(@SupCode,@PartNum,@KanbanNum,@FRSNum=<7-char FRS>,@RenbanNum=<assigned>,
   @Qty=<lots×lotqty>)` per truck-line with qty>0 (zero-qty lines skipped, `:540`). Keep the
   full 7-char FRS (suffix recompute is a truncation no-op). Re-inserted rows are status-empty
   (stock-neutral) and pass the order-file gen's `VC_RENBAN_NUMBER <> ''` gate.
3. `UPDATE_RenbanGroupCount(@RenbanCode, @RenbanCount=%.3d(next_count))` — advance the group's
   roll-over counter (3-digit, wraps at 999).

**Preserve the invariants:** inverted lot flag (0=lot-sized), blank↔assigned renban lifecycle,
delete-then-reinsert (not in-place update), exact lot-multiple qtys (no rounding), sum-all per
part, one-supplier-per-group / one-renban-per-trailer, group-level ship-day override, and stock
neutrality via status-empty rows.

---

## Appendix — cited proofs (all bounded; READ-ONLY `Inventory_Live`, rolled-back `Inventory`)

- `INV_RENBAN_GROUP_MST`: 5 rows, code UNIQUE (0 dup); counts 068/297/484/088/634 (all numeric);
  ship-days per weekday; no Sunday column. Bodies: table `/tmp/inv_utf8.sql:3870-3891`.
- Proc bodies verbatim: `SELECT_OrderNoRenban:6345`, `DELETE_OrderRenban:7629`,
  `INSERT_OpenOrder:7358`, `UPDATE_RenbanGroupCount:5142`, `SELECT_RenbanGroup:4952`,
  `SELECT_PartShipDays:4783`; triggers `INSERT_RecConfStatPartsStockMstQTY:7492` +
  `DELETE_RecConfStatPartsStockMstQTY` (body read live).
- Inverted flag: in-group flag=1 ×21 / 0 grouped flag=0; no-group flag=0 ×11.
- Lot multiples: 1957/1957 CMWA order rows `IN_QTY % IN_1LOTQTY = 0`.
- Trailer grouping: renban `CMWA334` / FRS `4071001` carries 3 parts (40/1200/360, all exact
  lot multiples); FRS suffix dist 01×1153/02×798/03×6.
- Renban format: `CMWA100`..`CMWA999`; len 3..8; DICAS(5)+3=8=varchar(8) cap.
- Supplier consistency: 3111/3111 grouped open-order supplier == part-master supplier (0
  mismatch); 1 distinct supplier per group.
- `SELECT_OrderNoRenban` describe: `IN_QTY` ord 6 (order) vs 42 (on-hand); group code ord 57,
  count ord 58; kanban ord 20 (o) vs 33 (p).
- FRS-suffix truncation: `DECLARE @F varchar(7)='6090201'; SET @F=@F+'02'` → `'6090201'` (len
  7); 5-char seed → `'6090202'`.
- End-to-end write-back on `Inventory` (rolled back): placeholders→assigned `CMWA289`/FRS
  `6090101`, remaining_blank=0, stock 44418/28133 unchanged, count 288→290→(rollback)288,
  eligible_for_orderfile=1.
- Stock neutrality: both triggers gated on `VC_STATUS_SUPPLIER_SHIPPING <> ''` +
  `VC_INVENTORY_ADD_POINT`; placeholders status-empty → no bump/de-bump.
- Div-by-zero: 0 grouped parts with `IN_1LOTQTY` 0/NULL. HIST is a HEAP (no PK).
