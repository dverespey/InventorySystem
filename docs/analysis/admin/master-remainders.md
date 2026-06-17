# Module Analysis: Phase-1 Master Remainders (RenbanGroup, MonthlyPO, Part-type)

**Area:** Master data (Phase-1 leftovers)  **Status:** ✅ spec complete  **Analyst:** Claude / 2026-06-16

The three master-maintenance items not yet specced: **`RenbanGroupMaster`** (important — the
renban-grouping master behind the Order palletized-parts pipeline), **`MonthlyPOMaster`**
(CAMEX monthly PO master), and the **Part-type** question (resolved: not a distinct form).

All three forms confirmed live in `InventorySystem.dpr` (RenbanGroupMaster `:31`,
MonthlyPOMaster `:45`). Depth follows `docs/analysis/master-data/supplier.md`.

---

## A. RenbanGroupMaster (`RenbanGroupMaster.pas` + `.dfm`)

### A.1 Legacy surface
- **Form:** `TRenbanGroupMaster_Form` — grid + detail panel; Insert/Update/Search/Clear/
  Delete/Close. Entry: Administration/master-maintenance menu in `MainMenu.pas`.
- **Purpose:** maintains the **renban group master** `INV_RENBAN_GROUP_MST`. A *renban group*
  is the grouping unit for palletized parts that ship together under a shared renban-group code
  (e.g. "CMWA"). Each group carries a **renban counter** (`VC_RENBAN_GROUP_COUNT`, a 3-char
  zero-padded sequence) and a per-weekday ship-days schedule. The Order/RenbanOrder pipeline
  groups palletized parts by this code and uses the counter to assign renban numbers.

### A.2 Data touched
| Table | Read | Write | Notes |
|-------|:----:|:-----:|-------|
| `INV_RENBAN_GROUP_MST` | ✅ | ✅ | group code, 3-char count, ship-days (overall + per weekday). |

**`INV_RENBAN_GROUP_MST`** (`Create Inventory.sql:1789+`):
```
IN_RENBAN_ID            int IDENTITY PK (surrogate)
VC_RENBAN_GROUP_CODE    varchar(5)  NOT NULL    -- e.g. CMWA
VC_RENBAN_GROUP_COUNT   varchar(3)  NULL        -- zero-padded counter "001".."999"
IN_SHIP_DAYS            int NULL
IN_SHIP_DAYS_MONDAY..SATURDAY int NULL          -- per-weekday ship-day offsets
VC_ADD / VC_LAST_UPDATE varchar(16)             -- 16-char yyyymmddHHMMSSff stamps
```
**Triggers:** none.

### A.3 Stored procedures
| Proc | Op | Rule (from body) |
|------|----|------------------|
| `SELECT_RenbanGroup` (`:7679`) | SELECT | `@RenbanCode=''` → all rows ordered by code; else filtered. Aliases include `IN_RENBAN_ID 'RecordID'` (surrogate carried to UI). |
| `INSERT_RenbanGroup` (`:3618`) | INSERT | Inserts code, count, ship-days (overall + per-day), and `VC_ADD` = 16-char stamp. **No dup guard in the proc.** |
| `UPDATE_RenbanGroup` (`:9229`) | UPDATE | Keyed on `IN_RENBAN_ID = @RenbanID` (surrogate). `@ShipDays = -1` → updates **only** code+count+stamp (counter-only edit); otherwise updates the full ship-day set too. Sets `VC_LAST_UPDATE` = 16-char stamp. |
| `UPDATE_RenbanGroupCount` (`:9281`) | UPDATE | `SET VC_RENBAN_GROUP_COUNT=@RenbanCount, VC_LAST_UPDATE=stamp WHERE VC_RENBAN_GROUP_CODE=@RenbanCode`. **Keyed on the code string, not the id.** The dedicated counter-bump proc. |
| `DELETE_RenbanGroup` (`:2481`) | DELETE | `DELETE WHERE IN_RENBAN_ID = @RenbanID`. Unconditional — no referenced-by check. |

**DataModule wiring:** `GetRenbanGroupInfo` (`:2032`), `InsertRenbanGroupInfo` (`:2076`),
`UpdateRenbanGroupInfo` (`:2141`), `DeleteRenbanGroupInfo` (`:2202`).

### A.4 Renban-counter maintenance (the load-bearing part)
There are **two** writers of `VC_RENBAN_GROUP_COUNT`:
1. **Manual (this form):** the operator types the count into `RenbanGroupCount_Edit`;
   `HoldDetails(False)` (`RenbanGroupMaster.pas:143-148`) **left-zero-pads it to exactly 3
   chars** (1→"001", 2→"012", 3-char passthrough) before sending it to `INSERT_RenbanGroup` /
   `UPDATE_RenbanGroup`. So the master count is always a 3-char string. The form's `Validate`
   (`:251`) requires it to be numeric but does **not** range-check 0–999.
2. **Automatic (Order pipeline):** `UPDATE_RenbanGroupCount` is the dedicated bump proc that
   advances the counter for a group **by code**. **It is NOT called from `DataModule.pas` or
   from RenbanGroupMaster** — confirmed by grep across the live `.pas` set; the only matches in
   `RenbanGroupMaster.pas` are the UI edit/pad. The Order/RenbanOrder write pipeline is the
   caller (it increments the group's counter when assigning the next renban). **Body of the
   caller belongs to the Order domain — cross-reference `docs/analysis/order/`** for the
   increment logic and the read-current-count → assign → write-incremented-count sequence.
   *(Caller verified absent from this module; increment arithmetic in the order pipeline is
   out of scope here — flag for the order spec.)*

🟠 **Concurrency hazard (counter):** `UPDATE_RenbanGroupCount` is a bare `UPDATE ... SET count =
@new` with the new value computed **client-side** (read-then-write), not `SET count = count + 1`
in-proc. Two concurrent order runs for the same group could read the same count and assign a
duplicate renban. Faithful to legacy single-operator use; the rebuild should make the bump
atomic/server-side.

🟠 **Code-keyed bump vs id-keyed everything-else:** the counter proc keys on the **code string**
while CRUD keys on the surrogate `IN_RENBAN_ID`. Under D2 (surrogate everywhere) the rebuild
should bump by id; the code stays a unique-per-site attribute (D1).

### A.5 P12 retry-recursion (cross-cutting — already documented, NOT new)
`InsertRenbanGroupInfo` retries `InsertSizeInfo` on exception (`DataModule.pas:2129`);
`UpdateRenbanGroupInfo` retries `UpdateSizeInfo` (`:2186`); `DeleteRenbanGroupInfo` retries
`DeleteSupplierInfo` (`:2234`). The log strings also mislabel as `fAssyCode`. **All three are
already catalogued** in `docs/analysis/cross-cutting/datamodule-retry-target-bugs.md`
(entries: CRITICAL #3 `DeleteRenbanGroupInfo`→`DeleteSupplierInfo` `:2234`; CRITICAL #7
`UpdateRenbanGroupInfo`→`UpdateSizeInfo`; MODERATE `InsertRenbanGroupInfo`→`InsertSizeInfo`
`:2130`; LOW `GetRenbanGroupInfo`→`GetSizeInfo`). **No new P12 bug discovered here** — the
Delete→DeleteSupplier path keyed on the shared `fRecordID` (a renban id reinterpreted as a
supplier id) is the most dangerous and is already a CRITICAL entry.

---

## B. MonthlyPOMaster (`MonthlyPOMaster.pas` + `.dfm`)

### B.1 Legacy surface
- **Form:** `TMonthlyPOMaster_Form` — grid + detail (assy code combo, PO start/end dates,
  pickup date+time, PO number, assy cost (currEdit), PO qty, PO-charged). Entry:
  Administration/master menu (`MainMenu.pas`; live `.dpr:45`).
- **Purpose:** the **CAMEX monthly purchase-order master** — registers a monthly PO window for
  an assembly code: PO number, validity window (start/end), scheduled pickup, unit cost, ordered
  qty, and a running **charged** counter. Drives PO-based billing/ordering for CAMEX sites
  (`fiPOEDISupport`/`fiCreatePOPriorToClose` gate the PO workflow — see configuration-site.md).

### B.2 Data touched
| Table | Read | Write | Notes |
|-------|:----:|:-----:|-------|
| `INV_ASSY_MONTHLY_PO` | ✅ | ✅ | PO window, pickup, cost, qty, `IN_PO_CHARGED`, add/update stamps. |
| `INV_ASSY_PO_CHARGED` | | ✅ | charged-against-PO ledger (written by `INSERT_AssyPOCharged`, not this form — Order/close pipeline). |

### B.3 Stored procedures
| Proc | Op | Rule (from body) |
|------|----|------------------|
| `SELECT_AssyMonthlyPO` (`:5664`) / `SELECT_AssyMonthlyPODisplay` (`:5702`) | SELECT | grid/display reads. |
| `INSERT_AssyMonthlyPO` (`:2895`) | INSERT | `@AssyCode,@POStart(8),@POEnd(8),@PickUp(8),@PickUpTime(4),@PONumber(10),@AssyCost money,@POQty int`. Inserts with `IN_PO_CHARGED = 0` and `VC_ADD = VC_LAST_UPDATE` = 16-char stamp. 🟠 **Uses positional `INSERT ... VALUES(...)` with NO column list** — fragile to any `INV_ASSY_MONTHLY_PO` column reorder/add. |
| `UPDATE_AssyMonthlyPO` (`:8268`) | UPDATE | edits the PO row (body unverified beyond signature; assume keyed on PO/assy). |
| `DELETE_AssyMonthlyPO` (`:2052`) | DELETE | removes a PO row (body unverified). |
| `INSERT_AssyPOCharged` (`:2929`) | INSERT+UPDATE | (Order/close pipeline, not this form) inserts a charged row AND `UPDATE INV_ASSY_MONTHLY_PO SET IN_PO_CHARGED = IN_PO_CHARGED + @qty WHERE assy AND ponumber` — the running-charged maintenance. Listed here because it mutates this master's `IN_PO_CHARGED`. |

**DataModule wiring:** `GetMonthlyPOInfo` (`:1816`), `InsertMonthlyPOInfo` (`:1858`),
`UpdateMonthlyPOInfo` (`:1921`), `DeleteMonthlyPOInfo` (`:1984`).

### B.4 P12 (already documented, NOT new)
`InsertMonthlyPOInfo`→`InsertSizeInfo` (`:1910`), `UpdateMonthlyPOInfo`→`UpdateSizeInfo`
(`:1973`), `GetMonthlyPOInfo`→`GetSizeInfo` — all already in the cross-cutting P12 catalogue
(MODERATE/LOW). **No new finding.**

### B.5 Notes
- `IN_PO_CHARGED` is incremented by `INSERT_AssyPOCharged` (Order/close), giving a
  PO-vs-charged comparison; this master holds the authorized qty/cost and the running charged.
- Date fields are 8-char (`yyyymmdd`) and pickup-time 4-char (`HHmm`), consistent with the rest
  of the system's string-date convention.

---

## C. Part-type master — resolved: NOT a distinct form
There is **no standalone Part-type master form** in `InventorySystem.dpr`. "Part type"
(`fPartType`, `DataModule.pas:198`) is a plain **attribute on the parts/stock master**
(`PartsStockMaster`, already specced in `docs/analysis/master-data/parts-stock-master.md`) and
appears as fields/combos within other masters (BCRatio, AssyRatio, Order, InvMgmt — grep matches
are field references, not a dedicated screen). **No separate spec needed**; the attribute is
owned by parts-stock-master.

---

## 6. Target design (Ignition) — applies to A & B
- **D1:** both `INV_RENBAN_GROUP_MST` and `INV_ASSY_MONTHLY_PO` gain `site_id`; the group code
  and the (assy code, PO number) become **unique per-site** attributes (D1/D2). Renban-group
  CRUD and Monthly-PO CRUD become role-gated Perspective master views backed by Named Queries
  mirroring these procs.
- **D2:** key all CRUD on the surrogate (`IN_RENBAN_ID`); change `UPDATE_RenbanGroupCount` to
  bump **by id**, atomically (`SET count = count + 1` server-side, or a sequence) to close the
  read-then-write race. Code stays an editable unique-per-site attribute.
- **D3:** block delete of a renban group referenced by any order/part, and of a monthly PO
  referenced by any `INV_ASSY_PO_CHARGED` row, instead of the current unconditional delete.
- **Counter ownership:** the renban counter belongs to the **Order/RenbanOrder** service (it
  bumps on renban assignment); the master view edits the *current* value but the runtime
  increment lives with order processing — reuse one atomic counter service.
- **MonthlyPO:** fix `INSERT_AssyMonthlyPO`'s positional VALUES by using an explicit column list
  in the Named Query. `IN_PO_CHARGED` maintenance (the `+= qty` in `INSERT_AssyPOCharged`)
  becomes part of the order-close service, transactional with the charged-ledger insert.

## 7. Migration plan
- [ ] Stage 1 — wrap `SELECT_*` for read-only renban-group + monthly-PO views.
- [ ] Stage 2 — writes via the (corrected) procs; atomic counter bump; explicit column lists.
- [ ] Stage 3 — reimplement with `site_id`, surrogate keys, RESTRICT-on-delete.

## 8. Open questions for the user
1. **Renban counter semantics:** confirm `VC_RENBAN_GROUP_COUNT` is a rolling assignment
   sequence that wraps at 999→001 (3-char), and that the Order pipeline is the sole automatic
   bumper. Should the rebuild make it a true atomic sequence per (site, group)?
2. **Range check:** the form doesn't bound the count 0–999 — is "999 then wrap" the intended
   behavior, or should it hard-stop / roll the group?
3. **Monthly PO uniqueness:** is `(site, assy code, PO number)` the intended unique key, and can
   two open PO windows for the same assy code overlap (cf. D6 for manifest-cost windows)?
4. **PO-charged reconciliation:** should `IN_PO_CHARGED` ever exceed `IN_PO_QTY` (over-charge),
   or be blocked?

## 9. Parity / regression checks
- RenbanGroup insert "5" → master row count `"005"`; counter-only edit (`@ShipDays=-1`) leaves
  ship-days untouched; full edit updates the weekday set; delete by `IN_RENBAN_ID`.
- MonthlyPO insert → row with `IN_PO_CHARGED=0` + 16-char add stamp; an `INSERT_AssyPOCharged`
  for that (assy, PO) raises `IN_PO_CHARGED` by qty.
- Verify the order-pipeline renban bump advances the *correct group's* counter (regression for
  the code-keyed `UPDATE_RenbanGroupCount`).
