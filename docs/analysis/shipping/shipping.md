# Module Analysis: Shipping (Daily build pull → stock-OUT) — `Shipping` / `ManualShipping` / `ModifyShipping`

**Area:** Shipping  **Status:** ✅ spec complete  **Analyst:** Claude / 2026-06-15

> **The stock-OUT side of the ledger — the mirror image of Receiving.** Where Receiving
> (`RecConfStat`, see [`../receiving/recconfstat.md`](../receiving/recconfstat.md)) *adds* parts to
> on-hand at supplier-shipping / arrival, **Shipping consumes them at production**: an operator
> declares a production date + a GALC sequence-number range (or, in manual mode, raw counts), the app
> explodes each built assembly into its tire/wheel **part numbers × ratios**, and writes one
> `INV_PART_SHIPPING_INF` row per part — and the trigger **`InsertPartShipping` SUBTRACTS that qty
> from `INV_PARTS_STOCK_MST.IN_QTY`** (schema:10152). That trigger is the entire behavioral spec; the
> Pascal forms and the `INSERT_Shipping*` procs are thin plumbing. **Three forms, one stock effect.**
>
> **Three big findings up front (all evidenced below):**
> 1. **The shipping qty triggers key on the `VC_PART_NUMBER` *string*** (schema:10152/10128/10177),
>    exactly like Receiving's `…RecConfStatPartsStockMstQTY` triggers — confirming the cross-module
>    **stock-ledger keying inconsistency**: shipping + receiving key on the string; reject + stocktaking
>    key on `IN_PART_ID` int. **And shipping has NO add-point gate at all** — it *always* subtracts,
>    unlike receiving's `'S'`/`'A'` gated add.
> 2. **Two confirmed proc-signature mismatches** in `ManualShipping`'s DataModule writers
>    (`InsertShippingDetailManual` calls `INSERT_ShippingDetail` with 5 wrong-named params vs the
>    schema's 4; `InsertAutoScrap` calls `INSERT_StocktakingInfo` with 5 params vs the schema's 3) →
>    latent runtime failures. New findings, not previously catalogued.
> 3. **`InsertShippingInfo` is the one transactional stock-OUT path** (`BeginTrans` →
>    `INSERT_ShippingInfo` → `CalculateFRS` ratio-explosion → commit/rollback). `CalculateFRS` is the
>    real engine; `InsertShippingInfo` itself only writes the header.

---

## 1. Legacy surface

All four units are **live** (confirmed in `InventorySystem.dpr`): `Shipping` (line 14),
`DailyBuildTotal` (44), `ManualShipping` (46), `ModifyShipping` (54). This file covers the three
operator forms; **`DailyBuildTotal` has its own spec** ([`dailybuildtotal.md`](dailybuildtotal.md))
because it is a substantial 3-mode batch processor (daily ALC pull + ASN + invoice export).

| Form | File | Size | Entry point |
|------|------|------|-------------|
| **`Shipping`** | `Shipping.pas` + `.dfm` | 450 lines / ~14 KB | `MainMenu.pas:314` `CarOrTruckShip_ButtonClick` → **the `else` branch** (`Shipping_Form`) when `fiFileALC = False` (`MainMenu.pas:334-337`). |
| **`ManualShipping`** | `ManualShipping.pas` + `.dfm` | 508 lines / ~14 KB | **same button**, the `if fiFileALC` branch (`MainMenu.pas:322-325`). The INI `[INIT] FileALC` flag picks Shipping vs ManualShipping at runtime — **they are alternatives, not both reachable in one install.** |
| **`ModifyShipping`** | `ModifyShipping.pas` + `.dfm` | 248 lines / ~6 KB | **Not** on the main menu — opened only from inside `Shipping` via `UpdateShipping_Button` (`Shipping.pas:433-440`), which is shown only when the chosen production date has *already* been shipped (`SetDetailBoxes`, `Shipping.pas:222`). It is the "edit an already-posted shipment's part lines" sub-form. |

- **`fiFileALC` (the mode switch):** `Shipping` is the **GALC-sequence** path — operator gives a
  start/end sequence number, the app queries the ALC database (`AD_ProductionSeq`, `AD_FRSPull`) for
  the vehicles built in that range and explodes them by broadcast-code ratios. `ManualShipping` is the
  **no-ALC** path — operator types a per-part daily count into a grid; no sequence lookup, no ratio
  explosion. Same downstream table (`INV_PART_SHIPPING_INF`) and the same stock-OUT trigger.

- **Purpose (one paragraph):** Record that a day's vehicle build consumed inventory, and decrement
  on-hand accordingly. `Shipping` derives the consumed parts from the GALC build sequence; the
  `Check_Button` first *previews* the vehicle count for a sequence range (`AD_ProductionSeq`), then
  `Insert_Button` posts it — writing the `INV_SHIPPING_INF` header and, via `CalculateFRS`, one
  `INV_PART_SHIPPING_INF` line per exploded part (each firing the stock-OUT trigger). `ManualShipping`
  posts the same header + lines from hand-entered counts and supports an "irregular ship" one-off
  adjustment. `ModifyShipping` edits the part lines of a shipment already posted for that date/line.

## 2. Data touched

| Table | Read | Write | Notes |
|-------|:----:|:-----:|-------|
| `INV_SHIPPING_INF` | ✓ | ✓ | **Shipment header**, one per (line, production date). PK `IN_SHIPPING_ID IDENTITY` (DDL schema:1708). Written by `INSERT_ShippingInfo`; read by `SELECT_ShipLastSeq` / `SELECT_ShipMax`. |
| `INV_PART_SHIPPING_INF` | ✓ | ✓ | **The stock-OUT ledger rows** — one per exploded part. **This module owns it.** PK `IN_PART_SHIPPING_ID IDENTITY`, FK-by-convention `IN_SHIPPING_ID`, `VC_PART_NUMBER varchar(12)`, `IN_QTY`, 16-char `VC_ADD` (DDL schema:1644). |
| `INV_PARTS_STOCK_MST` |  | ✓* | *Indirect — **the three `*PartShipping` triggers subtract/restore `IN_QTY` and stamp `VC_LAST_UPDATE`.** The reason this module matters to inventory. |
| `INV_SUPPLIER_MST` | ✓ |  | Joined in `SELECT_PartsDailyLinePull(Count)` for supplier-name display only. ⚠️ **Note: the shipping qty triggers do NOT read add-point** (unlike Receiving) — they always subtract. |
| `INV_ASSY_RATIO_MST` | ✓ |  | Read by `SELECT_AssyRatioInfoAssy` / `SELECT_ForecastDetailBC` to get tire/wheel part numbers + ratios for the explosion. |
| `INV_STOCKTAKING_INF` |  | ✓* | *Only via `InsertAutoScrap` (DailyBuildTotal path) — see §3 / `dailybuildtotal.md`. Fires `INSERT_Stocktaking` (a *separate* stock adjustment keyed on `IN_PART_ID`). |

### `INV_PART_SHIPPING_INF` columns (authoritative DDL `DB Schema/Create Inventory.sql:1644`)
| Column | Type | Role |
|--------|------|------|
| `IN_PART_SHIPPING_ID` | `int IDENTITY NOT NULL` | Detail PK. **This is the key `UPDATE_Shippingdetail` updates on** (schema:9320) — distinct from the parent `IN_SHIPPING_ID`. |
| `IN_SHIPPING_ID` | `int NOT NULL` | FK-by-convention to the `INV_SHIPPING_INF` header. No declared FK. |
| `VC_PART_NUMBER` | `varchar(12) NOT NULL` | **The join key all three qty triggers use** into `INV_PARTS_STOCK_MST.VC_PART_NUMBER` (string). |
| `VC_PRODUCTION_DATE` | `varchar(8) NOT NULL` | `yyyymmdd`. **The key `DeleteShipDate` cascades on** (schema:10337). |
| `IN_QTY` | `int NOT NULL` | The exploded part qty. **The amount subtracted from on-hand** by `InsertPartShipping`. |
| `VC_ADD` | `varchar(16) NOT NULL` | **16-char `yyyymmddHHMMSSff`** (set by `INSERT_ShippingDetail`/`INSERT_ShippingPartInfo`, schema:3674/3744). **Copied onto `INV_PARTS_STOCK_MST.VC_LAST_UPDATE` by the INSERT/UPDATE triggers.** |

### `INV_SHIPPING_INF` columns (DDL schema:1708)
`IN_SHIPPING_ID` (PK IDENTITY), `VC_LINE_NAME varchar(15)`, `VC_START_SEQ_NUMBER varchar(4)`,
`DT_START_SEQ_NUMBER datetime NULL`, `VC_END_SEQ_NUMBER varchar(4)`, `DT_END_SEQ_NUMBER datetime NULL`,
`IN_CONTINUE_NUMBER int NULL`, `IN_QTY int NULL`, `VC_PRODUCTION_DATE varchar(8) NOT NULL`,
`VC_LAST_UPDATE varchar(16) NULL`, **`VC_ADD varchar(50) NOT NULL`** (note: 50, not 16 — but
`INSERT_ShippingInfo` writes only the 16-char recipe into it, schema:3718). **No declared PK/UNIQUE
beyond the IDENTITY; no declared FKs.** Natural uniqueness `(VC_LINE_NAME, VC_PRODUCTION_DATE)` is
**app-convention only** (enforced by `SetDetailBoxes`' "already processed" check, not the DB).

### Triggers — THE STOCK-OUT SPEC (read live bodies)

All three `*PartShipping` triggers fire on `INV_PART_SHIPPING_INF` and key on the **`VC_PART_NUMBER`
string**. **None reads `VC_INVENTORY_ADD_POINT`** — shipping is unconditional stock-OUT.

- **`InsertPartShipping`** (FOR INSERT, schema:10152) — **the core stock-OUT leg:**
  ```sql
  UPDATE INV_PARTS_STOCK_MST
  SET IN_QTY = PS.IN_QTY - i.IN_QTY, VC_LAST_UPDATE = i.VC_ADD
  FROM INV_PARTS_STOCK_MST PS, inserted i
  WHERE PS.VC_PART_NUMBER = i.VC_PART_NUMBER
  ```
  **Invariant: inserting a part-shipping row SUBTRACTS its `IN_QTY` from on-hand** and stamps the
  16-char `VC_ADD` onto `VC_LAST_UPDATE`. ⚠️ **Old-style `FROM a,b WHERE` cross-join update** — if a
  part number appears N times in `inserted` (multi-row insert), SQL Server's non-deterministic single
  assignment means only **one** decrement is applied per `PS` row (classic "trigger doesn't handle
  multi-row" hazard). In practice each `INSERT_ShippingPartInfo`/`INSERT_ShippingDetail` call inserts
  one row at a time (`DoPartNumberInventory` loops), so it fires once per part — but the trigger is
  **not multi-row-safe** and the rebuild must apply this as a per-row additive delta.
- **`DeletePartShipping`** (FOR DELETE, schema:10128): `IN_QTY = PS.IN_QTY + d.IN_QTY` on
  `VC_PART_NUMBER` match. **Invariant: deleting a part-shipping row RESTORES (adds back) its qty.**
  ⚠️ Does **not** rewrite `VC_LAST_UPDATE`. **No purge-mode bypass** (unlike Receiving's DELETE
  trigger) — so a bulk purge of `INV_PART_SHIPPING_INF` *would* re-inflate on-hand.
- **`UpdatePartShipping`** (FOR UPDATE, schema:10177): a remove-old (`+ d.IN_QTY`) **then** add-new
  (`- i.IN_QTY`) pair → **net `−= (i.IN_QTY − d.IN_QTY)`**, a correct delta. Re-stamps `VC_LAST_UPDATE`
  from `i.VC_ADD`. This fires from `ModifyShipping`'s `UPDATE_Shippingdetail`. Same string key.

- **`DeleteShipDate`** (FOR DELETE on `INV_SHIPPING_INF`, schema:10334) — **cascade:**
  ```sql
  DELETE FROM INV_PART_SHIPPING_INF
  WHERE VC_PRODUCTION_DATE in (SELECT VC_PRODUCTION_DATE FROM deleted)
  ```
  Deleting a shipment header **deletes all part-shipping rows for that production date** — which fires
  `DeletePartShipping` and **restores on-hand**. ⚠️⚠️ **Keyed on `VC_PRODUCTION_DATE` ONLY — no line
  scope.** If two lines (Car / Truck) shipped on the same production date, deleting *one* header
  **wipes both lines' part-shipping rows** and restores stock for both. Confirmed cross-line hazard;
  mirror of the supplier-blind delete found in Receiving (recconfstat.md §4). No proc in any of these
  three forms deletes an `INV_SHIPPING_INF` header (no `DELETE_ShippingInfo` is called by them — the
  cascade exists for purge / other callers), so this is a latent path, not an everyday operation.

## 3. Stored procedures used
(All bodies read from `DB Schema/Create Inventory.sql`. Confidence noted per row.)

| Proc | Op | Business rule (from body) |
|------|----|---------------------------|
| `SELECT_ShipMax;1 (@LineName)` | SELECT | schema:7770. `SELECT max(vc_production_date) prod FROM INV_SHIPPING_INF WHERE VC_LINE_NAME=@LineName`. Drives `GetNextProductionDate` (next un-shipped day). **No site filter.** |
| `SELECT_ShipLastSeq;1 (@LineName,@Date)` | SELECT | schema:7734. If exactly one shipment exists for (line, date) → returns it (the "already processed" lock path); else returns the row with `max(VC_ADD)` for the line (the latest, to seed the next start-seq). ⚠️ The `else` branch **ignores `@Date`** — returns the line's latest regardless of date. **No site filter.** |
| `SELECT_ShippingDetail;1 (@ShipID)` | SELECT | schema:7793. `SELECT * FROM INV_PART_SHIPPING_INF WHERE IN_SHIPPING_ID=@ShipID`. ModifyShipping's grid source. (Pascal passes `@ShipID := fRecordID = parent IN_SHIPPING_ID` — correct here.) |
| `INSERT_ShippingInfo;1` | INSERT | schema:3696. Two live overloads exist — the **9-param OUTPUT form** (`@ShippingID OUTPUT, @LineName, @StartSeq, @DTStartSeq, @EndSeq, @DTEndSeq, @QTY, @Continue, @Date`) is the one the schema declares (returns `SCOPE_IDENTITY()`). Inserts the header with a 16-char `VC_ADD`. ⚠️ **`InsertShippingInfoManual`/`InsertExcelShippingEndInfo` call `INSERT_ShippingInfo` with a *different 6-param set* (`@AssyLine,@StartSeq,@LastSeq,@QTY,@Continue,@Date` — no OUTPUT, no DT params)** — a **signature mismatch** against the schema's 9-param proc (see §4 finding M3). |
| `INSERT_ShippingPartInfo;1 (@ShippingID,@Part,@QTY,@Date)` | INSERT/UPDATE | schema:3735. **Idempotent (P1 in-proc):** `SELECT * … WHERE IN_SHIPPING_ID=@ShippingID AND VC_PART_NUMBER=@Part`; if found → `UPDATE … SET IN_QTY = IN_QTY + @QTY` (the UPDATE fires `UpdatePartShipping` → net delta), else `INSERT` (fires `InsertPartShipping` → subtract). **This is the dedup that makes re-posting the same part safe.** Called by `DoPartNumberInventory`. |
| `INSERT_ShippingDetail;1 (@PartShipID,@PartNumber,@Productiondate,@Qty)` | INSERT | schema:3662. Plain `INSERT INTO INV_PART_SHIPPING_INF VALUES(...)` with a 16-char `VC_ADD` → fires `InsertPartShipping` (subtract). Called correctly by `InsertShippingInfoDetail` (ModifyShipping insert). ⚠️ **`InsertShippingDetailManual` calls this with 5 wrong-named params** → fails (finding M1). |
| `UPDATE_Shippingdetail;1 (@PartShipID,@PartNumber,@Qty)` | UPDATE | schema:9315. `UPDATE INV_PART_SHIPPING_INF SET IN_QTY=@Qty, VC_PART_NUMBER=@PartNumber WHERE IN_PART_SHIPPING_ID=@PartShipID`. Fires `UpdatePartShipping` (delta re-balance). **Keys on the DETAIL PK** `IN_PART_SHIPPING_ID`. |
| `SELECT_PartsDailyLinePull;1 (@AssyLine)` | SELECT | schema:7206. Parts for a line (`INV_PARTS_STOCK_MST` ⋈ supplier ⋈ part-type) for ManualShipping's grid. `@AssyLine varchar(10)`. |
| `SELECT_PartsDailyLinePullCount;1 (@AssyLine,@Date)` | SELECT | schema:7237. `SUM(s.in_qty)` of `INV_PART_SHIPPING_INF` per part for a date. ⚠️ **`@AssyLine` declared `varchar(1)`** but Pascal passes the full line name (`fAssyCode`); ⚠️ the `s.vc_production_date=@Date` predicate is in the **JOIN ON**, not WHERE — works but fragile. Read-only display. |
| `SELECT_AssyRatioInfoAssy;1 (@AssyCode)` | SELECT | schema:5808. `SELECT * FROM INV_ASSY_RATIO_MST WHERE VC_ASSY_PART_NUMBER_CODE=@AssyCode`. The tire/wheel part numbers + ratios for the explosion (DailyBuildTotal path). |
| `SELECT_ForecastDetailBC;1 (@BCode,@EffMonth,@TireWheel)` | SELECT | Used by `CalculateFRS` to map a broadcast code → part numbers + ratios per tire/wheel slot, by effective month. **Body unverified** (combo/forecast lookup; not stock-bearing itself — it feeds `DoPartNumberInventory`). |
| `AD_GetLines;1`, `AD_ProductionSeq;1`, `AD_GetLastPrint;1`, `AD_FRSPull;1` | SELECT | **ALC database** procs (cross-connection, `ALC_StoredProc`/`ALC_DataSet`). Line list, vehicle-count preview, sequence date/time lookup, and the FRS broadcast-code pull. **Bodies live in the ALC DB, not `Create Inventory.sql`** — out of scope for this schema; treat as external read-only feeds. |
| `INSERT_AssyBuildHist`, `INSERT_AssyPOCharged`, `UPDATE_AssyBuildHistINV`, `SELECT_AssyBuildHist`, `INSERT_StocktakingInfo`, `SELECT_PartsStockInfo` | mixed | **DailyBuildTotal-owned** — bodies verified there; see [`dailybuildtotal.md`](dailybuildtotal.md). |

### `DoPartNumberInventory` — the stock-OUT call site (`DataModule.pas:6717`)
The single funnel that turns a built-assembly count into a stock decrement:
```pascal
update := round((qty*ratio) / 100);           // 6722 — ratio is a percentage
... ProcedureName := 'INSERT_ShippingPartInfo;1';  // 6733
    @ShippingID := fRecordID; @part := partnumber; @QTY := update; @Date := fProductionDate;
```
**Business rule:** consumed qty per part = `round(builtCount × ratio / 100)`. Integer rounding here is
load-bearing for parity (banker's vs half-up — Delphi `round` is **banker's rounding**; the rebuild
must match or accept per-part ±1 drift). `INSERT_ShippingPartInfo`'s in-proc upsert means repeated
calls for the same (shipping, part) **accumulate** (`IN_QTY += @QTY`) rather than duplicate.

### Call mechanism & P12 retry audit (`DataModule.pas`)
The Shipping writers mostly do **not** use the recursive `If fErrorCount < 3` retry harness (the core
posters `InsertShippingInfo` 4914 / `InsertShippingInfoManual` 4668 / `InsertExcelShippingEndInfo`
4231 / `InsertPOCharged` 4366 / `InsertBuildHist` 4531 / `DoPartNumberInventory` 6717 just log/`raise`
on error — no wrong-target recursion). The retry-harness methods here:

- **`GetPartsListCount` (3622) → retries into `GetRecConfStatInfo` (3657)** — ⚠️ **P12 wrong-target
  (LOW)**. Already catalogued in [`../cross-cutting/datamodule-retry-target-bugs.md`](../cross-cutting/datamodule-retry-target-bugs.md)
  (the 🟡 LOW list, `GetPartsListCount→GetRecConfStatInfo`). Read-only; no new entry.
- **`GetPartsList` (3668) → retries into `GetRecConfStatInfo` (3701)** — ⚠️ **P12 (LOW)**, also in the
  existing 🟡 list (`GetPartsList→GetRecConfStatInfo`). No new entry.
- **`UpdateINVDone` (4317) → retries into `UpdateAssyRatioInfo` (4351)** — already catalogued
  (MODERATE list). DailyBuildTotal-invoice path; cited in `dailybuildtotal.md`.
- **`GetBuildHist` (4410) → retries into `GetStocktakingInfo` (4450)** — already in the 🟡 LOW list
  (`GetBuildHist→GetStocktakingInfo`). DailyBuildTotal path.
- **`GetShippingInfo` (4116), `GetShippingInfoDetail` (4073), `InsertShippingInfoDetail` (3977),
  `UpdateShippingInfoDetail` (4026)** — retry into **themselves**. ✅ Correct; **not** P12.

> **No NEW P12 entries from Shipping** — all wrong-target retries here were already found by the
> cross-cutting audit. Cite, don't re-report. (The genuinely new defects in this area are the two
> **signature mismatches** below, which are a different bug class.)

## 4. Business rules & edge cases

- **Shipping always subtracts; there is no add-point gate.** Unlike Receiving (where
  `VC_INVENTORY_ADD_POINT` decides whether/when stock moves), the `*PartShipping` triggers
  unconditionally decrement on insert and restore on delete. Stock-OUT is universal at production.
- **Stock-OUT feeds the ONE re-homed stock-ledger as NEGATIVE deltas.** Per the Receiving spec's
  additive-delta model: a shipping post = a **`−qty` ledger posting** per part
  (`round(built×ratio/100)`); a delete/restore = **`+qty`**; an edit = the signed delta. This is the
  exact inverse of Receiving's `+qty` postings and the same shape as `DailyBuildTotal`'s negative
  "Auto Scrap Delete" rows (D5) — so all four (receiving add, shipping subtract, stocktaking delta,
  reject) collapse into one additive ledger keyed on `IN_PART_ID`.
- **The "already processed" lock.** `SetDetailBoxes` (`Shipping.pas:199` / `ManualShipping.pas:180`)
  compares the form's production date to `SELECT_ShipLastSeq`'s returned `VC_PRODUCTION_DATE`. If equal
  → the day is already shipped → controls lock, Insert/Check disabled, and (Shipping only) the
  `UpdateShipping_Button` appears to open `ModifyShipping`. **This is the only dup-guard** — there is
  no DB unique constraint on (line, date), so a concurrent second post would double-subtract.
- **`Check` then `Insert` sequencing (Shipping).** `Insert_Button` refuses unless `fCheck` is true
  (`Shipping.pas:142`) — the operator must run `Check_ButtonClick` (vehicle-count preview via
  `AD_ProductionSeq`) first. `fCheck` is reset to false on any production-date change
  (`Shipping.pas:329`). Pure UI gate.
- **`CalculateFRS` is the explosion engine** (`DataModule.pas:4728`). For each broadcast code in the
  ALC `AD_FRSPull` result, for each of the 4 tire/wheel slots `['T','W','V','F']`, it reads
  `SELECT_ForecastDetailBC` (part numbers + ratios for the production month) and calls
  `DoPartNumberInventory(partNo, ratio, orders)`. **Special case:** if `Orders ∈ {4,5}` it forces
  `ratio=100` for the first item and `break`s (one-wheel-set rule, `DataModule.pas:4792-4799`).
  **A missing broadcast code for slot index `< 2` (T or W) raises and rolls back the whole post**
  (`:4812-4817`); a missing V/F slot is silently skipped. The disabled `AD_WQSFRS` `else` branch
  (`:4824-4899`, commented out) is **dead code** — do not spec it.
- **The transaction boundary.** `InsertShippingInfo` (4914) wraps `BeginTrans → INSERT_ShippingInfo
  (header, captures `@ShippingID` OUTPUT into `fRecordID`) → CalculateFRS → Commit/Rollback`. If
  `CalculateFRS` returns false, the header insert is rolled back too — **the post is atomic**.
  `ManualShipping`'s `Post_ButtonClick` (`ManualShipping.pas:329`) does its own `BeginTrans` around
  `InsertShippingInfoManual` + a per-grid-row `InsertShippingDetailManual` loop.
- **Timestamp encoding (P2) — count check.** Every shipping `VC_ADD` uses
  `CONVERT(varchar, getdate(), 112)` (**8 chars** `yyyymmdd`) `+ SUBSTRING(...,114,1,2)+(4,2)+(7,2)+(10,2)`
  (**4×2 = 8 chars** `HHMMSSff`) = **16 chars total, `yyyymmddHHMMSSff`** (verified char-by-char at
  `INSERT_ShippingDetail` schema:3674, `INSERT_ShippingInfo` schema:3718, `INSERT_ShippingPartInfo`
  schema:3744). **Not 14, not 8.** `INV_SHIPPING_INF.VC_ADD` is declared `varchar(50)` but only the
  16-char value is written. The trigger copies this 16-char string onto `INV_PARTS_STOCK_MST.VC_LAST_UPDATE`.

### 🐞 NEW findings — proc signature mismatches (latent runtime failures)

> These are **new** (not in the existing cross-cutting registers). They are a different class from
> P12: the *call* is to the right method, but the **parameter set doesn't match the proc the schema
> declares**, so ADO's by-name parameter binding sends params the proc never declares (or omits ones
> it requires) → the call fails at runtime (or silently no-ops). All confirmed against schema bodies.

- **M1 — `InsertShippingDetailManual` ↔ `INSERT_ShippingDetail` (5 vs 4, wrong names).**
  Pascal (`DataModule.pas:4631-4642`) calls `INSERT_ShippingDetail;1` with
  `@part, @QTY, @Date, @AssyLine, @IrregularQty`. The **only** `INSERT_ShippingDetail` in the schema
  (schema:3662) declares `@PartShipID, @PartNumber, @Productiondate, @Qty`. **None of the names match.**
  → `ManualShipping`'s per-part grid post (`Post_ButtonClick`) and the **irregular-ship** button
  (`IrregularShip_ButtonClick`, `ManualShipping.pas:497`) both call this → **both fail at runtime.**
  Confidence: **high** (both proc bodies read; no second overload exists — grep schema:339/3662 are the
  only `INSERT_ShippingDetail` objects). This implies `ManualShipping`'s detail/irregular path is
  **effectively broken in production** (or relies on an undeployed proc variant — flag to domain expert).
- **M2 — `InsertAutoScrap` ↔ `INSERT_StocktakingInfo` (5 vs 3, wrong names).**
  Pascal (`DataModule.pas:4485-4496`) calls `INSERT_StocktakingInfo;1` with
  `@SupCode, @PartCode, @QTY, @Reason, @AutoScrap`. The schema proc (schema:3821) declares only
  `@PartNumber, @QTY, @Reason`. → the DailyBuildTotal **auto-scrap** path (negative stocktaking delta)
  would fail. Confidence: **high** (body read; resolves the part via `IN_PART_ID` internally, no
  `@SupCode`/`@AutoScrap`). *(Caveat: the live SQL Server may carry a newer `INSERT_StocktakingInfo`
  overload than the checked-in script — `Create Inventory.sql` is a snapshot. Treat as "schema/code
  drift to confirm," but the checked-in schema is the authoritative artifact and it does not match.)*
- **M3 — `InsertShippingInfoManual` / `InsertExcelShippingEndInfo` ↔ `INSERT_ShippingInfo` (6 vs 9).**
  These two (`DataModule.pas:4686` / `4239`) call `INSERT_ShippingInfo;1` with
  `@AssyLine, @StartSeq, @LastSeq, @QTY, @Continue, @Date` — but the schema proc (schema:3696) declares
  9 params incl. the **OUTPUT `@ShippingID`** and `@DTStartSeq/@DTEndSeq` and uses `@LineName`/`@EndSeq`
  (not `@AssyLine`/`@LastSeq`). **Names and arity both differ.** Confidence: **high** for the mismatch;
  **medium** on impact (same drift caveat as M2 — there may be a legacy 6-param overload not in this
  snapshot). The GALC path (`InsertShippingInfo`, 4914) uses the correct 9-param OUTPUT form.

> **Net:** the **GALC/`Shipping` post path is internally consistent** (correct procs); the
> **`ManualShipping` and auto-scrap paths show 1:1 code↔schema mismatches** in the checked-in artifact.
> The rebuild resolves all of these by defining one canonical Named Query per operation (§6).

## 5. UI / UX notes
- **`Shipping`:** line combo (from `AD_GetLines`), production date picker (auto-advanced past
  holidays/overtime by `GetNextProductionDate`), start/end sequence edits with date-time pick boxes
  (`StartBox`/`EndBox` populated from `AD_GetLastPrint`), ship-qty + continuation mask edits.
  `Check` → preview count; `Insert` → post; `UpdateShipping` (conditional) → `ModifyShipping`. A
  `Refill_Button` handler exists (`Shipping.pas:334`, calls `CalculateFRS` with `Refill:=True`) but
  **no `Refill_Button` control is declared on the form** — likely a dead/removed button (the `.dfm`
  has no such control); treat the handler as latent/dead.
- **`ManualShipping`:** a `TStringGrid` of parts (Part / Supplier / Description / Daily Build Count);
  operator types a count per part, `Update` accumulates a running `DailyTotal`, `Post` writes header +
  detail rows in a transaction; `IrregularShip` posts a single off-grid adjustment (path is broken per
  M1).
- **`ModifyShipping`:** `DBGrid` over `SELECT_ShippingDetail` (the posted part lines); selecting a row
  loads part+qty into editors; `Update` calls `UPDATE_Shippingdetail` (qty/part-number edit → on-hand
  re-balance via `UpdatePartShipping`); `Insert` adds a new detail line (`INSERT_ShippingDetail`, which
  subtracts). **No delete button is wired** — you cannot remove a posted line here (so the only stock
  *restore* path is the `DeleteShipDate` header-cascade, which no form here triggers).
- **Modernize:** replace the (line, date) app-convention lock with a DB unique constraint + an explicit
  "shipment already posted — view / amend" flow; surface the ratio explosion (show which parts × how
  many will be consumed) before commit; make irregular-ship a first-class adjustment; scope the
  `DeleteShipDate` cascade to (line, date); server-side everything.

## 6. Target design (Ignition — Perspective + Named Queries + gateway stock-ledger)
- **Perspective views:**
  - `Shipping/PostBuild` — line dropdown, production-date picker, sequence range (GALC mode) **or** a
    parts-count table (manual mode); a **"preview consumption"** panel showing each exploded part ×
    `round(built×ratio/100)` before commit; a single **Post** action.
  - `Shipping/AmendShipment` — replacement for `ModifyShipping`: a table of the posted part lines
    (Named Query `SelectShippingDetail`) with inline qty/part edit and an explicit **remove line**
    (currently impossible) that posts the restoring `+qty` ledger delta.
- **Named Queries (one per proc, IA practice — single-point schema edits):** `SelectShipMax`,
  `SelectShipLastSeq`, `InsertShippingHeader` (the 9-param OUTPUT form), `InsertShippingPartInfo`
  (the idempotent upsert — keep its in-proc dedup), `InsertShippingDetail`, `UpdateShippingDetail`,
  `SelectPartsDailyLinePull(+Count)`, `SelectAssyRatioInfoAssy`, `SelectForecastDetailBC`. The ALC-DB
  procs (`AD_*`) become Named Queries against the ALC datasource (cross-connection preserved).
  **Define ONE canonical signature per operation** — this is where M1/M2/M3 get fixed by construction.
- **Gateway stock-ledger service (the core re-homing — shared with Receiving/Reject/Stocktaking).**
  Re-implement the three `*PartShipping` triggers + the `DeleteShipDate` cascade as explicit, atomic,
  **multi-row-safe** ledger postings:
  - Post **`−round(built×ratio/100)`** per exploded part at shipment post (was `InsertPartShipping`).
  - Post **`+qty`** when a line is removed / a header is deleted (was `DeletePartShipping` /
    `DeleteShipDate`) — and **scope the header-delete restore to (line, production date)**, fixing the
    line-blind cascade.
  - Post the signed **delta** on a detail edit (was `UpdatePartShipping`).
  - **Key the ledger on `IN_PART_ID`** — resolve `VC_PART_NUMBER` → surrogate **once at the boundary**,
    standardizing off the string/int keying inconsistency (D2). This is the same boundary resolution
    Receiving's re-home does; both modules then post to the identical additive-delta ledger.
  - Match Delphi `round` (banker's) or explicitly choose half-up and document the parity delta.
  - Each posting stamps a 16-char-equivalent real timestamp on `VC_LAST_UPDATE` (normalize the string).
- **Reports:** none owned by these three forms (ASN/invoice CSV export lives in `DailyBuildTotal`).

## 7. Migration plan for this module
- [ ] **Stage 1 — read-only parity:** `SelectShipMax` / `SelectShipLastSeq` / `SelectShippingDetail`
      → Perspective; reproduce the "already-processed" lock; add `site_id` (D1) to every query.
- [ ] **Stage 2 — writes via wrapped procs:** post through `InsertShippingHeader` (OUTPUT id) +
      `InsertShippingPartInfo` (keep the idempotent upsert) **with the live `*PartShipping` triggers
      intact** so on-hand stays correct during parallel-run. **Fix M1/M2/M3 in the wrapper** by calling
      the one correct proc signature (no recursive wrong-target retry; bounded transport retry only).
      Wrap the GALC `Check→Post` and the manual grid post in a DB transaction matching the legacy.
- [ ] **Stage 3 — reimplement (Postgres-ready):** add `site_id` NOT NULL FK; give `INV_SHIPPING_INF`
      a real unique `(site_id, line, production_date)`; give `INV_PART_SHIPPING_INF` a real FK to the
      header (`IN_SHIPPING_ID`) and to the part (`IN_PART_ID`); move the three triggers + `DeleteShipDate`
      cascade into the `StockLedger` gateway service (keyed on `IN_PART_ID`, multi-row-safe, additive
      `−qty` deltas); scope the header-delete restore to (line, date); normalize the 16-char `VC_ADD`/
      `VC_LAST_UPDATE` to real timestamps; decide rounding policy.

## 8. Open questions for the domain expert (candidate decisions)
1. **(candidate D#) — `ManualShipping` detail/irregular path appears broken (M1).** The
   checked-in `INSERT_ShippingDetail` (4 params) does not match `InsertShippingDetailManual`'s call
   (5 differently-named params). Is `ManualShipping` actually used in production today, or has the
   live DB a different `INSERT_ShippingDetail` overload than the checked-in script? (Determines whether
   this is a real outage or schema-snapshot drift.) Same question for `InsertAutoScrap` ↔
   `INSERT_StocktakingInfo` (M2) and the manual `INSERT_ShippingInfo` 6-vs-9 form (M3).
2. **(candidate D#) — `DeleteShipDate` line-blind cascade.** Deleting a shipment header restores stock
   for **every** part-shipping row on that `VC_PRODUCTION_DATE` regardless of line (Car/Truck). Confirm
   the rebuild should scope the restore to (site, line, production date). (Mirror of the Receiving
   supplier-blind delete already raised in recconfstat.md §8.4.)
3. **Ratio rounding policy.** Consumed qty = `round(built × ratio / 100)` with Delphi **banker's
   rounding**. Is per-part ±1 drift acceptable, or must the rebuild reproduce banker's exactly? Is
   there a reconciliation step that catches rounding drift today?
4. **Re-posting / amend semantics.** `INSERT_ShippingPartInfo` *accumulates* (`IN_QTY += @QTY`) on a
   repeat (shipping, part). Is re-running a day's post meant to add to, or replace, the prior counts?
   And should `ModifyShipping` gain a **delete-line** (currently the only restore path is the
   header-cascade, which no form triggers)?
5. **Missing-broadcast-code hard-fail.** A missing T or W broadcast code rolls back the entire post; a
   missing V/F slot is silently skipped (`CalculateFRS` `i<2` test). Is the asymmetry intended?
6. ✅ **RESOLVED (D1):** per-site isolation — `INV_SHIPPING_INF` / `INV_PART_SHIPPING_INF` gain
   `site_id` NOT NULL FK; the (line, date) lock and `SELECT_Ship*` queries scope to site.
7. ✅ **RESOLVED (D2):** stock-ledger postings key on the surrogate `IN_PART_ID`; the legacy
   `VC_PART_NUMBER`-string trigger key is resolved to the surrogate at the boundary.

## 9. Test cases / parity checks
- **Next production date** = first un-shipped day after `SELECT_ShipMax` for the line (holiday/overtime
  advanced). Re-opening a shipped day → controls lock, `UpdateShipping` button shows.
- **GALC post (`Shipping.Insert`)** for a sequence range → one `INV_SHIPPING_INF` header (16-char
  `VC_ADD`), and one `INV_PART_SHIPPING_INF` row per exploded part with
  `IN_QTY = round(built×ratio/100)`; **`INV_PARTS_STOCK_MST.IN_QTY −= Σ part qty`** per part
  (`InsertPartShipping`); `VC_LAST_UPDATE` = the 16-char `VC_ADD`.
- **Re-post the same part** (same shipping id) → `INSERT_ShippingPartInfo` UPDATE branch
  → on-hand subtracts the *additional* qty only (accumulation), not a duplicate row.
- **Edit a posted detail qty (`ModifyShipping`)** Q1→Q2 → on-hand `+= (Q1−Q2)` net (`UpdatePartShipping`
  delta pair); part-number edit re-points the row.
- **Delete a shipment header** (e.g. via purge) with two lines on one date → **legacy: both lines'
  part rows deleted, on-hand restored for both** (line-blind `DeleteShipDate`); assert the rebuild
  restores only the targeted (line, date).
- **Manual post (`ManualShipping`)** of a parts grid → header + one detail per non-zero part; on-hand
  subtracts each. **Assert M1**: against the checked-in schema this call fails (no matching
  `INSERT_ShippingDetail` overload) — parity test must record whether the live DB diverges.
- **Timestamp**: every shipping `VC_ADD` is exactly **16 chars** (`yyyymmddHHMMSSff`); assert not 8/14.
- **P12 retry parity:** force a transient failure on `GetPartsList`/`GetPartsListCount`/`GetBuildHist`
  → assert the rebuild does **not** retry into `GetRecConfStatInfo`/`GetStocktakingInfo` (legacy P12).
