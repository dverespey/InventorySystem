# Module Analysis: Receiving Confirmation Status (`RecConfStat` → `INV_OPEN_ORDER_INF`)

**Area:** Receiving  **Status:** ✅ spec complete  **Analyst:** Claude / 2026-06-15

> **The arrival/receiving "status board" — and the second-most stock-moving screen in the app.**
> `RecConfStat` is a wide editable grid over **`INV_OPEN_ORDER_INF`** (one row per open
> order = supplier+part+FRS+RENBAN). It is **not** a clean "receiving" form: it lets an operator
> create, status-stamp, and delete open-order rows, and it owns the **logistics milestone stamps**
> (in-transit, plant yard, assembler yard, warehouse, arrival, terminated). The crucial fact:
> **this form does NOT contain the stock logic.** The on-hand balance `INV_PARTS_STOCK_MST.IN_QTY`
> is moved entirely by **three triggers on `INV_OPEN_ORDER_INF`** (`INSERT_/UPDATE_/DELETE_
> RecConfStatPartsStockMstQTY`), gated by the supplier's `VC_INVENTORY_ADD_POINT` (`'S'` =
> add-at-supplier-shipping, `'A'` = add-at-arrival). Per **decision D7**, the `'A'`-supplier
> **arrival** stock-add fires *here* (the `VC_ARRIVAL` stamp written by `UPDATE_RecConfStatInfo` /
> `UPDATE_RecConfStatRenbanInfo`), and per **decision D8(3)** the matching arrival-**reversal**
> branch is **dead code that the rebuild must implement**. The procs are thin pass-throughs; the
> triggers (read below) are the behavioral spec.

## 1. Legacy surface
- **Form:** `RecConfStat.pas` (904 lines / ~30 KB) + `RecConfStat.dfm` (809 lines).
  `TRecConfStat_Form`, caption "Receiving Confirmation Status". Author: Aaron Huge, 2002-10-25
  (12/17/2002: `IN_QTY`/`Quantity` added — confirmed in the DataModule header and the `--AH`
  comments at `DataModule.pas:3196/3273/3356`). Registered live in `InventorySystem.dpr` **line 9**
  (`RecConfStat in 'RecConfStat.pas' {RecConfStat_Form}`).
- **Entry point:** **`MainMenu.pas:282` `RecStatMgmt_ButtonClick`** — `Hide;
  RecConfStat_Form := TRecConfStat_Form.Create(self); if RecConfStat_Form.Execute then
  ShowMessage('Order file generation completed'); RecConfStat_Form.Free; Show;` (the
  "Order file generation completed" message is vestigial — `Execute` returns `True` unless cancelled;
  no file is generated). `Execute` (`RecConfStat.pas:125`) pre-loads the assembler-location combo
  from `INV_DOCK_INF.VC_DOCK_NAME` then `ShowModal`.
- **Purpose (one paragraph):** Operators view every open order and **stamp logistics milestones** as
  a shipment moves from supplier → in-transit → plant yard → assembler yard / warehouse → arrival →
  terminated, plus trailer/parking/detention free-text. The same screen can **insert** a brand-new
  open order, **edit** one order, **batch-edit every order in a RENBAN group** (the
  `RenbanUpdate_CheckBox` path), and **delete** an order. Every milestone field is a `yyyymmdd`
  date string (8 chars) held in an `INV_OPEN_ORDER_INF.VC_*` column; "set" vs "blank" on a few of
  those columns is what drives the qty triggers.

## 2. Data touched
| Table | Read | Write | Notes |
|-------|:----:|:-----:|-------|
| `INV_OPEN_ORDER_INF` | ✓ | ✓ | **This module owns it** — INSERT/UPDATE/DELETE via the four procs below |
| `INV_OPEN_ORDER_INF_HIST` |  | ✓* | *Indirect: INSERT trigger snapshots `inserted`; UPDATE trigger snapshots `deleted` (`INSERT … SELECT *` — schema-order fragile, P10-adjacent; note HIST col `VC__KANBAN_NUMBER` has a **double-underscore typo** vs live `VC_KANBAN_NUMBER`, schema:1532) |
| `INV_PARTS_STOCK_MST` |  | ✓* | *Indirect: the three qty triggers add/subtract `IN_QTY` and rewrite `VC_LAST_UPDATE`. **The entire reason this module matters to inventory.** |
| `INV_SUPPLIER_MST` | ✓* | | *Joined inside every qty trigger to read `VC_INVENTORY_ADD_POINT` (`'S'`/`'A'`); also the supplier-code search combo (`SelectSingleField`) |
| `INV_DOCK_INF` | ✓ | | Assembler-location combo source (`Execute`, `RecConfStat.pas:130`) |
| `INV_PART_STOCK_MST` *(sic)* | ✓ | | ⚠️ Two combo-reload calls (`RecConfStat.pas:695-696`) name **`INV_PART_STOCK_MST`** (missing the `S` in `PARTS`) — a **non-existent table** → those reloads silently fail/raise inside a dead branch (only reached when supplier text is blanked). Latent bug; the working path uses `SELECT_DependantPartNumber_Supplier` |

### `INV_OPEN_ORDER_INF` columns (authoritative: `DB Schema/Create Inventory.sql:1484`)
| Column | Type | Meaning / role |
|--------|------|----------------|
| `IN_ORDER_ID` | `int IDENTITY` **NOT NULL** | Surrogate PK. ⚠️ **Never used by this module** — not selected, not a proc param; all proc keys use the natural 4-tuple (see §4) |
| `VC_SUPPLIER_CODE` | `varchar(5) NOT NULL` | Supplier business code (string, **not** an FK to `IN_SUPPLIER_ID`). The qty triggers re-join supplier through `INV_PARTS_STOCK_MST.IN_SUPPLIER_ID`, **not** this column |
| `VC_PART_NUMBER` | `varchar(12) NOT NULL` | Part number (string). **The join key the qty triggers use** into `INV_PARTS_STOCK_MST.VC_PART_NUMBER` |
| `VC_FRS_NUMBER` | `varchar(7) NOT NULL` | "FRS" order number (1-char year prefix + 4 — drives `VC_FRS_DATE`, §4) |
| `VC_RENBAN_NUMBER` | `varchar(8) NOT NULL` | RENBAN sequence. **The batch-update key** (`UPDATE_RecConfStatRenbanInfo` keys on this alone) |
| `IN_QTY` | `int NOT NULL` | Order quantity. **This is the amount added to / subtracted from on-hand by the triggers** — distinct from the part-master on-hand balance |
| `VC_STATUS_SUPPLIER_SHIPPING` | `varchar(8) NOT NULL` | In-transit date stamp. **Non-blank ⇒ shipped.** Drives the `'S'` add-point add/remove |
| `VC_ARRIVAL` | `varchar(8) NOT NULL` | Arrival date stamp. **Non-blank ⇒ arrived.** Drives the `'A'` add-point add (D7) |
| `VC_TRAILER_NUMBER` | `varchar(11) NOT NULL` | Trailer free-text |
| `VC_STATUS_PLANT_YARD` | `varchar(8) NOT NULL` | Plant-yard date stamp. One of the `'A'` arrival-status alternatives (INSERT/DELETE triggers) |
| `VC_PLANT_PARKING` | `varchar(10) NOT NULL` | Parking-spot free-text |
| `VC_STATUS_ASSEMBLER_YARD` | `varchar(8) NOT NULL` | Assembler-yard date stamp. `'A'` arrival-status alternative |
| `VC_ASSEMBLER_LOCATION` | `varchar(10) NOT NULL` | Dock/location (from `INV_DOCK_INF`) |
| `VC_STATUS_EMPTY_TRAILER` | `varchar(8) NOT NULL` | Empty-trailer date. ⚠️ **Gate in the DELETE trigger only** (`= ''` required to reverse) |
| `VC_DETENTION` | `varchar(50) NOT NULL` | Detention free-text |
| `VC_ORDER_DATE` | `varchar(8) NOT NULL` | Order date stamp |
| `VC_WAREHOUSE` | `varchar(8) NOT NULL` | Warehouse date. `'A'` arrival-status alternative |
| `VC_TERMINATED` | `varchar(8) NOT NULL` | Terminated date. ⚠️ **Gate** (`= ''` required for INSERT/DELETE-trigger qty moves); also the `HideTerminated` filter |
| `VC_SHIP_DATE` | `varchar(8) NOT NULL` | Ship date (set by Update, **not** Insert — Insert proc omits it) |
| `VC_KANBAN_NUMBER` | `varchar(5) NOT NULL` | Kanban. ⚠️ **Both UPDATE procs leave it untouched** — `UPDATE_RecConfStatInfo` writes it, but `UPDATE_RecConfStatRenbanInfo`'s `VC_KANBAN_NUMBER = @Kanban` line is **commented out** (schema:9169); Insert proc never sets it |
| `VC_FRS_DATE` | `varchar(8) NULL` | **Derived** in `UPDATE_RecConfStatInfo` from `@FRSNo` with a year-rollover rule (§4). Only nullable column; Insert never sets it |
| `VC_LAST_UPDATE` | `varchar(16) NOT NULL` | **16-char `yyyymmddHHMMSSff` timestamp** (P2), set by both UPDATE procs and copied onto `INV_PARTS_STOCK_MST.VC_LAST_UPDATE` by the qty triggers |
| `VC_ADD` | `varchar(16) NOT NULL` | ⚠️ **Set by Insert as an 8-char `CONVERT(varchar(8),@Now,112)` value** (`yyyymmdd` only — schema:3552), even though the column is `varchar(16)`. **Not** the 16-char recipe used elsewhere (P2 deviation; the time portion is dropped on insert) |

**Constraints / indexes:** `IN_ORDER_ID IDENTITY` is the only PK-style column, but **there is no
declared PRIMARY KEY or UNIQUE index on `INV_OPEN_ORDER_INF`** (verify against the constraints
block — none found in the table DDL). The natural key `(VC_SUPPLIER_CODE, VC_PART_NUMBER,
VC_FRS_NUMBER, VC_RENBAN_NUMBER)` is treated as unique **only by application convention** (the
app dup-check + the proc `WHERE` clauses), **not** enforced by the DB. **No declared FOREIGN KEYs**
touch this table — supplier/part links are by string convention, resolved inside the triggers.

**Triggers on `INV_OPEN_ORDER_INF` (3 — read live bodies; these ARE the stock-ledger logic):**
All three key on the **string `VC_PART_NUMBER`** and re-derive the supplier (and thus add-point)
through `INV_PARTS_STOCK_MST.IN_SUPPLIER_ID → INV_SUPPLIER_MST`. A part with a NULL/dangling
`IN_SUPPLIER_ID` (e.g. after `DELETE_SupplierCode`) drops out of the INNER JOIN → **stock silently
never moves** for that part (edge case; ties to D4).

- **`INSERT_RecConfStatPartsStockMstQTY`** (FOR INSERT, schema:9717): (1) `INSERT INTO
  INV_OPEN_ORDER_INF_HIST SELECT * FROM inserted`; (2) **`'S'` leg** — `IN_QTY += i.IN_QTY` when
  `i.VC_STATUS_SUPPLIER_SHIPPING <> ''` AND supplier add-point `= 'S'`; (3) **`'A'` leg** —
  `IN_QTY += i.IN_QTY` when (`VC_ARRIVAL <> ''` **OR** `VC_STATUS_PLANT_YARD <> ''` **OR**
  `VC_STATUS_ASSEMBLER_YARD <> ''` **OR** `VC_WAREHOUSE <> ''`) AND add-point `= 'A'`.
  Also writes `VC_LAST_UPDATE = i.VC_LAST_UPDATE` onto the part row.
  **Invariant: inserting an already-shipped (`'S'`) or already-arrived (`'A'`) open order adds its
  qty to on-hand immediately.** ⚠️ Note: an INSERT has **no `VC_TERMINATED`/`VC_STATUS_EMPTY_TRAILER`
  gate** (unlike DELETE) — the gates are asymmetric between the INSERT and DELETE triggers.
- **`UPDATE_RecConfStatPartsStockMstQTY`** (FOR UPDATE, schema:9764) — **the most complex trigger in
  the DB.** Snapshots `deleted` to HIST, then runs **six logical `IN_QTY` legs** (= **8 raw `UPDATE
  INV_PARTS_STOCK_MST` statements** — legs 1 & 2 below are each a remove-old/add-new pair of statements),
  all gated `i.IN_QTY <> d.IN_QTY` or a status flip, joined on `VC_PART_NUMBER`:
  1. **Qty-change while shipped, `'S'`** (schema:9782-9808): a `−d.IN_QTY` leg **and** a `+i.IN_QTY`
     leg with **identical** `WHERE` (qty changed, `d.VC_STATUS_SUPPLIER_SHIPPING <> ''`, `'S'`) →
     net **`+= (i.IN_QTY − d.IN_QTY)`** (a correct delta — the two legs are a remove-old/add-new pair,
     NOT a double-count).
  2. **Qty-change while arrived, `'A'`** (schema:9813-9839): same delta pair gated `i.VC_ARRIVAL <>
     ''` + `'A'` → net **`+= (i.IN_QTY − d.IN_QTY)`**.
  3. **Ship-status set, `'S'`** (schema:9845): `+= i.IN_QTY` when `VC_STATUS_SUPPLIER_SHIPPING`
     went blank→set, `'S'`. (Mirror of the INSERT `'S'` leg for an edit that marks shipped.)
  4. **Ship-status cleared, `'S'`** (schema:9860): `−= i.IN_QTY` when set→blank, `'S'`.
  5. **Arrival set, `'A'`** (schema:9877): `+= i.IN_QTY` when `i.VC_ARRIVAL <> d.VC_ARRIVAL` AND
     `d.VC_ARRIVAL = ''`, `'A'`. ⚠️ **This is the D7 arrival-add** — the only path that counts stock
     for an `'A'` supplier, fired by stamping `VC_ARRIVAL`.
  6. **Arrival cleared, `'A'`** (schema:9894): intended `−= i.IN_QTY` when an arrival is unset — but
     the `WHERE` is **`i.VC_ARRIVAL = '' AND i.VC_ARRIVAL <> ''`** (both clauses on `i`, an
     always-false contradiction). ⚠️⚠️ **DEAD CODE — the arrival-reversal NEVER fires (D8 Bug 3).**
     The correct mirror is `i.VC_ARRIVAL <> d.VC_ARRIVAL AND i.VC_ARRIVAL = ''`. **Per D8(3) the
     rebuild MUST implement this reversal** (post a compensating `−qty` when an `'A'` arrival is
     cleared), homed in the receiving-confirmation action.
  **Invariant: editing an open order re-balances on-hand by qty-delta and by status-flip, gated on
  the supplier's add-point — EXCEPT clearing an arrival, which today leaves on-hand overstated.**
- **`DELETE_RecConfStatPartsStockMstQTY`** (FOR DELETE, schema:9660): **skipped entirely when
  `Purge.PurgeMode = 1`** (purge bypass — so bulk-purge of open orders doesn't drain on-hand).
  Otherwise: **`'S'` leg** `IN_QTY −= d.IN_QTY` when `d.VC_STATUS_SUPPLIER_SHIPPING <> ''` AND
  `VC_STATUS_EMPTY_TRAILER = ''` AND `VC_TERMINATED = ''` AND add-point `'S'`; **`'A'` leg**
  `IN_QTY −= d.IN_QTY` when (arrival/plant-yard/assembler-yard/warehouse set) AND empty-trailer
  `= ''` AND terminated `= ''` AND add-point `'A'`. ⚠️ Does **NOT** rewrite `VC_LAST_UPDATE` and does
  **NOT** snapshot to HIST. **Invariant: deleting a still-active counted order removes its qty from
  on-hand, unless it was already empty-trailered, terminated, or we're purging.**

## 3. Stored procedures used
(Read from `DB Schema/Create Inventory.sql`. All bodies verified.)

| Proc | Op | Business rule (from body) |
|------|----|---------------------------|
| `SELECT_RecConfStatInfo;1 (@SupCode='',@PartCode='',@FrsNo='',@Renban='')` | SELECT | schema:7593. `@SupCode=''` → **all** rows `ORDER BY supplier,part,FRS,RENBAN`; else the one row `WHERE` all four match. Returns **19 UI-aliased columns** mapped 1:1 to grid `Fields[0..18]`: `Supplier, Parts, FRSNo, RENBAN, Quantity, Order, In Transit, Arrival Date, Trailer No., PLANT Yard, Parking Spot, ASSEMBLER Yard, ASSEMBLER Location, Empty Trailer, Detention, Warehouse, Terminated, Kanban, Shipped`. **No site filter** (legacy single-site; under D1 must be scoped to `site_id`). Note `IN_QTY` aliased `'Quantity'` here is the **order** qty, not on-hand. |
| `INSERT_RecConfStatInfo;1` (17 params) | INSERT | schema:3525. Plain `INSERT INTO INV_OPEN_ORDER_INF` of the supplier-code/part/FRS/renban/qty + all status stamps + `VC_ADD = CONVERT(varchar(8),@Now,112)` (⚠️ **8-char add stamp**). **Does NOT set** `VC_KANBAN_NUMBER`, `VC_SHIP_DATE`, `VC_FRS_DATE`, `VC_LAST_UPDATE` (the NOT-NULL columns must rely on table DEFAULTs — verify defaults exist, else insert fails). The INSERT fires `INSERT_RecConfStatPartsStockMstQTY` (the qty add). No id returned. |
| `UPDATE_RecConfStatInfo;1` (23 params, incl. 4 `*Prev`) | UPDATE | schema:9055. Keys on the **previous** 4-tuple `(@SupCodePrev,@PartNumPrev,@FRSNoPrev,@RenbanCodePrev)` (the natural key is **editable** — P9-style "use the old key to find the row"). Rewrites all status columns incl. **`VC_ARRIVAL = @Arrival`** (D7 arrival stamp) and **`VC_KANBAN_NUMBER = @Kanban`**, sets `VC_LAST_UPDATE` (16-char). **Derives `VC_FRS_DATE`** (§4). Fires `UPDATE_RecConfStatPartsStockMstQTY` (the 6-leg re-balance). |
| `UPDATE_RecConfStatRenbanInfo;1` (16 params) | UPDATE | schema:9136. ⚠️ **Batch update — `WHERE VC_RENBAN_NUMBER = @Renban` only.** Updates **every open-order row in the RENBAN group at once** (no part/FRS key), rewriting status stamps + `VC_ORDER_DATE` + `VC_SHIP_DATE` + 16-char `VC_LAST_UPDATE`. **Does NOT touch** `VC_SUPPLIER_CODE/VC_PART_NUMBER/VC_FRS_NUMBER/IN_QTY/VC_FRS_DATE`, and **`VC_KANBAN_NUMBER` is commented out** (schema:9169). Because it can flip arrival/ship status on many rows simultaneously, it can trigger a **multi-row qty re-balance** in one shot. Reached via `RenbanUpdate_CheckBox` (`RecConfStat.pas:204`). |
| `DELETE_RecConfStatInfo;1 (@PartCode,@FrsNo,@Renban)` | DELETE | schema:2429. `DELETE FROM INV_OPEN_ORDER_INF WHERE VC_PART_NUMBER=@PartCode AND VC_FRS_NUMBER=@FrsNo AND VC_RENBAN_NUMBER=@Renban`. ⚠️ **Keys on only 3 of the 4 natural-key columns — supplier code is NOT in the WHERE** → if two suppliers share a part/FRS/RENBAN this deletes **both**. Fires `DELETE_RecConfStatPartsStockMstQTY` (qty removal, purge-gated). |
| `SELECT_DependantPartNumber_Supplier (@SupplierCode)` | SELECT | schema:5956. Cascading combo — parts/kanbans for the chosen supplier (`RecConfStat.pas:700`). Body unverified in detail (combo-population only; not stock-bearing). |
| `SELECT_DependantKanbanNumber_PartNumber (@PartNumber)` | SELECT | schema:5906. Cascading combo — kanban for the chosen part (`RecConfStat.pas:720`). Body unverified in detail (combo-population only). |

### Call mechanism (legacy — `DataModule.pas`)
- **`GetRecConfStatInfo`** (3106): opens `Inv_DataSet` on `SELECT_RecConfStatInfo` with all-blank
  params (load all). P8 retry → correctly re-calls **itself** (3145). ✅ not a P12 hazard.
- **`InsertRecConfStatInfo`** (3159): **two-step dup-guard (P1)** — STEP 1 `SELECT_RecConfStatInfo`
  with the 4-tuple; only `If RecordCount = 0` does STEP 2 `INSERT_RecConfStatInfo` (17 params). On
  `RecordCount > 0` returns `False` ("It already exists"). ✅ dup-check targets the **correct** proc.
  P8 retry → re-calls **itself** (3248). ✅ not P12.
- **`UpdateRecConfStatInfo`** (3339): 23 params incl. the four `*Prev` (captured into `Data_Module`
  in `HoldDetails(True)` at `RecConfStat.pas:335-338`). P8 retry → re-calls **itself** (3417). ✅.
- **`UpdateRecConfStatRenbanInfo`** (3262): 16 params. ⚠️ **P12 WRONG-TARGET RETRY BUG** — on
  exception (`fErrorCount < 3`) it calls **`UpdateRecConfStatInfo`** (`DataModule.pas:3325`), **not
  itself.** A transient failure of the RENBAN batch-update silently retries as a *single-row* update
  with whatever `Data_Module` fields are currently set — a different, likely-wrong DB write. Add to
  the P12 register.
- **`DeleteRecConfStatInfo`** (3055): 3 params. ⚠️ **P12 WRONG-TARGET RETRY BUG** — on exception it
  calls **`DeleteAssyRatioInfo`** (`DataModule.pas:3091`; the method is even mis-commented
  `//DeleteAssyRatioInfo` at 3103 and the log says "FAILED DELETE ASSY"). A failed RecConfStat delete
  retries by deleting **assembly-ratio** data. Egregious copy-paste; add to the P12 register.
- All four share the P8 `fErrorCount < 3` recursive-retry harness with `finally` closing the proc
  and resetting `fErrorCount := 0`. `RecordID` (the shared P9 field) is **not** used here — the
  module keys on the natural 4-tuple, so this screen is **not** a P9 `RecordID` hazard (it has its
  own `fSupplierCodePrev/PartNumPrev/FRSNoPrev/RenbanCodePrev` shadow fields instead).

## 4. Business rules & edge cases
- **Add-point is the master switch (D4/D7).** Whether an open-order event moves on-hand, and *when*,
  is read from `INV_SUPPLIER_MST.VC_INVENTORY_ADD_POINT` **inside the triggers** (via the part's
  supplier): `'S'` adds at supplier-shipping (`VC_STATUS_SUPPLIER_SHIPPING` set), `'A'` adds at
  arrival (`VC_ARRIVAL`/plant-yard/assembler-yard/warehouse set). Any other value, or a part with no
  resolvable supplier, moves **no** stock.
- **D7 — the `'A'` arrival-add lives here.** For `'A'` suppliers, the *only* event that counts stock
  is stamping `VC_ARRIVAL` (or one of the three alternative arrival statuses on INSERT). That stamp
  is written by `UPDATE_RecConfStatInfo`/`UPDATE_RecConfStatRenbanInfo` from the form's
  `ArrivalDate_NUMMIBmDateEdit` (`RecConfStat.pas:384`). The carrier/logistics feed records arrival
  *status* for `'A'` parts but does **not** count stock (see logistics-breakdown spec). **Rebuild:
  re-home the arrival-add into the receiving-confirmation action, keyed off the confirmed arrival.**
- **D8(3) — implement the arrival reversal.** The "changed to not arrived" trigger leg (schema:9894)
  is permanently dead (`i.VC_ARRIVAL='' AND i.VC_ARRIVAL<>''`). Today, **clearing a set arrival on an
  `'A'` order does not give back the previously-added stock → on-hand is overstated.** The rebuild's
  stock-ledger must post the compensating `−qty` (the corrected mirror
  `i.VC_ARRIVAL <> d.VC_ARRIVAL AND i.VC_ARRIVAL = ''`) in the receiving-confirmation path.
- **In-transit precedes everything (UI guard).** `Validate` (`RecConfStat.pas:782`): supplier code
  must be ≥5 chars, part ≥12, FRS ≥7, qty numeric; and **if `InTransit` is blank, none of arrival /
  warehouse / plant-yard / assembler-yard may be set** (`RecConfStat.pas:816-845`: "Order must be
  marked In Transit when arrival is set"). This is a UI-only invariant — the DB does not enforce it.
- **FRS-date year rollover.** `UPDATE_RecConfStatInfo` derives `VC_FRS_DATE` from `@FRSNo`
  (schema:9082-9090): if the 4th digit of today's `yyyymmdd` ≠ the 1st char of the FRS number, the
  order is treated as **next year** (`DATEADD(year,1,…)`); else current year. Builds an 8-char
  `yyyy + SUBSTRING(@FRSNo,2,4)`. Off-by-one risk: this compares the **4th** char of `yyyymmdd`
  (the last digit of the year) against a single-digit FRS year prefix. `INSERT_RecConfStatInfo` and
  the RENBAN update do **not** set `VC_FRS_DATE`, so it stays NULL until a (non-RENBAN) edit.
- **Natural key is editable; updates use the previous tuple.** `UPDATE_RecConfStatInfo` finds the
  row by `(@*Prev)` and may rewrite all four key columns — so an operator can re-point an order to a
  different supplier/part/FRS/RENBAN. ⚠️ Because the qty triggers join `inserted`/`deleted` on
  `VC_PART_NUMBER`, **changing the part number in one update has both `i` and `d` part numbers
  joining to (potentially different) part-master rows** — the qty re-balance math assumes the part
  number is stable across the edit (the qty-change legs join `i` and `d` on the *same* part number).
  A simultaneous part-number + qty change is an untested corner.
- **RENBAN batch edit is a blunt instrument.** `UPDATE_RecConfStatRenbanInfo` rewrites **every** row
  sharing the RENBAN — handy for "the whole renban arrived," but it overwrites each row's
  supplier-shipping/arrival/etc. with the single set of form values, and can fire a **multi-row** qty
  re-balance in one transaction. The `RenbanUpdate_CheckBox` is only enabled after a trailer-number
  change (`RecConfStat.pas:858`).
- **Delete drops the supplier from its key.** `DELETE_RecConfStatInfo` omits `VC_SUPPLIER_CODE` from
  its `WHERE` (schema:2436-2439) → a part/FRS/RENBAN shared across suppliers deletes more than the
  selected row. Combined with the purge-gated qty trigger, a delete normally returns the qty to
  on-hand (unless terminated / empty-trailered / purging).
- **Timestamp encodings are inconsistent (P2).** `VC_LAST_UPDATE` is the full **16-char**
  `yyyymmddHHMMSSff` (`CONVERT(varchar,getdate(),112)` + four 2-char `,114` slices). But
  `INSERT_RecConfStatInfo` writes `VC_ADD` as only **8 chars** (`CONVERT(varchar(8),@Now,112)`) —
  the add stamp loses its time portion. (Counting check: the `,114` recipe yields HH+MM+SS+ff = 8
  digits → 8+8 = **16** chars total for `VC_LAST_UPDATE`; the 8-char `VC_ADD` is date-only.)
- **Hide-terminated is a persisted user preference** (`fiHideTerminated`, a `TField` config flag,
  `RecConfStat.pas:311/880`), applied as a client-side `Filter [Terminated] = ''`.

## 5. UI / UX notes
- Wide `DBGrid` over all open orders + a large detail panel of ~18 date/text editors
  (`TNUMMIBmDateEdit` for every milestone). Selecting a row (`MouseUp`/`KeyUp`/`DataChange`) calls
  `HoldDetails(True)` → captures the row + the four `*Prev` key fields → `SetDetailBoxes` reformats
  each 8-char `yyyymmdd` into `mm/dd/yyyy` for the date pickers.
- **Search is entirely client-side (P7)** over the loaded grid via `Inv_Dataset.Filter` with `LIKE
  '%…%'` on Supplier/Parts/Kanban/FRSNo/RENBAN/Order/Shipped, plus a "No Order Search" mode
  (`Unordered_Box` → `[Order] LIKE ''`). Requires ≥1 criterion. `SortBy_ComboBox` sets
  `Inv_Dataset.Sort`.
- **Cascading combos:** Supplier → parts/kanban (`SELECT_DependantPartNumber_Supplier`); part →
  kanban (`SELECT_DependantKanbanNumber_PartNumber`).
- **`Quantity_Edit` is a plain editable `TEdit`** (`RecConfStat.pas:77`) — unlike the part-master
  screen, on-hand-affecting **order qty IS operator-editable here**, and an edit flows into the
  6-leg qty re-balance. (Distinct from `INV_PARTS_STOCK_MST.IN_QTY`, which this screen never edits
  directly.)
- **Modernize:** server-side search/sort/pagination (P7); make the milestone progression an explicit
  status workflow rather than free date stamps; surface add-point so operators see *why* a stamp did
  or didn't move stock; replace the all-or-nothing RENBAN batch update with an explicit multi-select;
  fix the supplier-blind delete key.

## 6. Target design (Ignition — Perspective + Named Queries + gateway stock-ledger)
- **Perspective views:**
  - `Receiving/OpenOrders` — a Perspective **Table** bound to a `SELECT_RecConfStatInfo` Named Query
    (param `site_id`, plus optional supplier/part/FRS/RENBAN/kanban/date filters pushed **server-side**,
    replacing the in-memory `LIKE`). Columns map 1:1 to the legacy 19 aliases.
  - `Receiving/OpenOrderEditor` — detail form: supplier/part cascading dropdowns (Named Queries
    `SELECT_DependantPartNumber_Supplier`, `SELECT_DependantKanbanNumber_PartNumber`), milestone date
    fields, order qty, and a clear **status timeline** component. Enforce the "In Transit first" guard
    in a binding/validation script (port of `RecConfStat.pas:816-845`).
- **Named Queries (mirror the procs, one NQ per proc — IA practice):**
  `SelectOpenOrders`, `InsertOpenOrder`, `UpdateOpenOrder`, `UpdateOpenOrderRenban`, `DeleteOpenOrder`.
  During parallel-run, wrap the existing procs verbatim via `system.db.createSProcCall` so the live
  triggers keep on-hand correct; add `site_id` to every WHERE.
- **Gateway stock-ledger service (the core re-homing).** Re-implement the three
  `…RecConfStatPartsStockMstQTY` triggers as **one explicit, atomic stock-ledger posting** in the
  shared `StockLedger` gateway script (the same service the shipping/reject/stocktaking modules feed):
  - On open-order **insert/update**: post `+qty` at the add-point event (`'S'`→shipping set,
    `'A'`→arrival set), `−qty` on the reverse (ship-status cleared, qty decrease, delete).
  - **Implement the D8(3) arrival-reversal** (`−qty` when an `'A'` arrival is cleared) — the one leg
    that is dead today.
  - Key the ledger on **`IN_PART_ID`** (resolve the part-number string to the surrogate once at the
    boundary — standardize off the inconsistent string/int keying noted in the inventory invariant).
  - Preserve the **purge-mode bypass** as a "skip re-balance when purging" flag.
  - Each posting writes a ledger row (analogous to `INV_PART_QTY_INF`) and stamps `VC_LAST_UPDATE`.
- **Fixes baked in:** correct the `DELETE_RecConfStatInfo` supplier-blind key (include `site_id` +
  supplier); fix the two **P12 wrong-target retries** (RENBAN→single, Delete→AssyRatio); decide
  whether RENBAN-batch and KANBAN-on-RENBAN are intended (§8).
- **Reports:** none owned here (receiving reporting lives in the breakdown/inventory modules).

## 7. Migration plan for this module
- [ ] **Stage 1 — read-only parity:** `SelectOpenOrders` NQ → Perspective table; 19 columns map 1:1
      to grid `Fields[0..18]`; add `site_id` filter (D1). Server-side search replaces the `LIKE` filter.
- [ ] **Stage 2 — writes via wrapped procs:** call `INSERT/UPDATE/UPDATE…Renban/DELETE_RecConfStatInfo`
      through `system.db.createSProcCall`, **keeping all three qty triggers live** so on-hand,
      add-point gating, HIST snapshots, and the purge bypass stay correct. Fix the two P12 retries in
      the Ignition wrapper (no recursive wrong-target retry). Re-key Delete on (site, supplier, part,
      FRS, RENBAN).
- [ ] **Stage 3 — reimplement (Postgres-ready):** add `site_id` NOT NULL FK; give the table a real PK
      (`IN_ORDER_ID`) and a `(site_id, supplier, part, FRS, RENBAN)` unique index; move the three qty
      triggers into the `StockLedger` gateway service keyed on `IN_PART_ID`; **implement the D8(3)
      arrival reversal**; normalize the 8-char `VC_ADD` / 16-char `VC_LAST_UPDATE` strings to real
      timestamps; fix the HIST `VC__KANBAN_NUMBER` typo and the schema-order-fragile `SELECT *` HIST
      inserts; decide RENBAN-batch + KANBAN-on-RENBAN behavior (§8).

## 8. Open questions for the user (domain expert)
1. **RENBAN batch update — intended scope?** `UpdateRecConfStatRenbanInfo` rewrites the milestone
   stamps on **every** open-order row sharing a RENBAN (no part/FRS key). Is "stamp the whole renban
   arrived in one click" the intended behavior, and should it really overwrite each row with a single
   set of values (and possibly mass-move stock)? Or should the rebuild make it an explicit multi-select?
2. **KANBAN on the RENBAN update.** `UPDATE_RecConfStatRenbanInfo` has `VC_KANBAN_NUMBER = @Kanban`
   **commented out** (schema:9169) and `INSERT_RecConfStatInfo` never sets kanban either, so a renban
   edit silently never changes kanban. Is that intentional, or a latent bug to fix?
3. **Asymmetric INSERT vs DELETE qty gates.** The INSERT qty add has **no** terminated/empty-trailer
   gate, but the DELETE removal requires `VC_TERMINATED='' AND VC_STATUS_EMPTY_TRAILER=''`. So an
   order inserted-then-terminated-then-deleted **adds** stock but never removes it (asymmetry → drift).
   Should the rebuild make insert/delete gates symmetric?
4. **Supplier-blind delete key.** `DELETE_RecConfStatInfo` omits supplier code — two suppliers sharing
   a part/FRS/RENBAN would both be deleted. Confirm the rebuild should key the delete on the full
   tuple (incl. supplier + site).
5. ✅ RESOLVED (D12) — **plant-yard AND assembler-yard count as arrival on edit too.** David:
   *"Plant/yard both count as arrival."* The legacy UPDATE `'A'` add-leg fires only on `VC_ARRIVAL`, so
   plant-yard/assembler-yard on an edit under-counts vs insert/delete. The rebuild's receiving action
   treats plant-yard and assembler-yard as arrival-equivalent on edit (symmetric with insert) so the
   `'A'`-supplier stock-add fires consistently — folds into the stock-ledger service with D7 + D8(3).
6. **`VC_FRS_DATE` derivation.** Confirm the year-rollover rule (compare last digit of current year
   vs FRS prefix) is correct, and whether a 2030s prefix collision (single-digit) is a concern.
7. ✅ **RESOLVED (D7):** the `'A'`-supplier arrival stock-add happens here, via the `VC_ARRIVAL`
   stamp — confirmed against the UPDATE trigger leg (schema:9877). Re-home to the receiving-confirmation
   stock-ledger action.
8. ✅ **RESOLVED (D8 Bug 3):** the arrival-reversal branch (schema:9894) is dead; **implement** the
   `−qty` reversal in the rebuild's receiving-confirmation stock-ledger.
9. ✅ **RESOLVED (D1):** per-site isolation — `INV_OPEN_ORDER_INF` gains a `site_id` NOT NULL FK;
   `SELECT_RecConfStatInfo` is scoped to the current site; the natural-key uniqueness becomes
   `(site_id, supplier, part, FRS, RENBAN)`.
10. ✅ **RESOLVED (D4):** add-point stays supplier-level; the trigger coupling (a part's qty behavior
    read from its supplier) is intended.

## 9. Test cases / parity checks
- **List all** → row count = `SELECT_RecConfStatInfo '' '' '' ''`; 19 columns map to `Fields[0..18]`.
- **Insert a shipped `'S'` order** (`VC_STATUS_SUPPLIER_SHIPPING` set, supplier add-point `'S'`) →
  `INV_PARTS_STOCK_MST.IN_QTY += order qty` (`INSERT_RecConfStatPartsStockMstQTY` `'S'` leg);
  `VC_ADD` is 8 chars (`yyyymmdd`).
- **Insert an arrived `'A'` order** (`VC_ARRIVAL` set, add-point `'A'`) → on-hand `+= qty` via the
  `'A'` leg; insert with only `VC_STATUS_PLANT_YARD` set (no arrival) on `'A'` → still `+= qty`.
- **Edit order qty (shipped/`'S'`)** from Q1→Q2 → on-hand `+= (Q2−Q1)` (delta pair, not double-count).
- **Stamp arrival on an `'A'` order** (blank→set) → on-hand `+= qty` (D7 add-leg, schema:9877).
- **Clear a set arrival on an `'A'` order** → **legacy: on-hand UNCHANGED** (dead branch, D8 Bug 3) /
  **rebuild: on-hand `−= qty`** (implemented reversal). Assert the divergence.
- **Mark `'S'` order not-shipped** (set→blank) → on-hand `−= qty` (schema:9860).
- **RENBAN batch update** marking a whole renban arrived (`'A'` parts) → on-hand of **each** part in
  the renban `+= its qty`; `VC_KANBAN_NUMBER` unchanged on all (commented-out).
- **Delete an active shipped order** (not terminated, not empty-trailered, purge **off**) → on-hand
  `−= qty`; with purge **on** → on-hand unchanged; if `VC_TERMINATED` set → on-hand unchanged.
- **Delete with a shared part/FRS/RENBAN across two suppliers** → legacy deletes **both** rows
  (supplier-blind key); assert the rebuild deletes only the selected (site+supplier) row.
- **P12 retry parity:** force a transient failure on the RENBAN update / delete → assert the rebuild
  does **not** retry into `UpdateRecConfStatInfo` / `DeleteAssyRatioInfo` (the legacy bug).
