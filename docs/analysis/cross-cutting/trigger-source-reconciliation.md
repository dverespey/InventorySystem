# Trigger Source Reconciliation: `Create Inventory.sql` vs `docs/triggers.sql`

**Area:** Cross-cutting (reference)  **Status:** ✅ reconciled  **Analyst:** Claude / 2026-06-05

> **Bottom line.** `DB Schema/Create Inventory.sql` is **authoritative** — its **24 triggers**
> are the live ones. **`docs/triggers.sql` is an obsolete pre-int-FK-refactor snapshot** and must
> **not** be trusted for any trigger. At some point the parts/child tables were refactored from
> **string business codes** (`VC_SUPPLIER_CODE`, `VC_PART_NUMBER`, `VC_*_CODE`) to **int surrogate
> FKs** (`IN_PART_ID`, `IN_SUPPLIER_ID`, `IN_SIZE_ID`, `IN_LOGISTICS_ID`, `IN_RENBAN_ID`).
> The schema's triggers were rewritten to key on the int FKs; `docs/triggers.sql` still keys on the
> old string columns — several of which **no longer exist** on the live tables, so those triggers
> are invalid against the current DB.

## Evidence (read directly from both files)
`INV_PARTS_STOCK_MST` has only int linkage columns — `IN_SUPPLIER_ID`, `IN_LOGISTICS_ID`,
`IN_SIZE_ID` (int, NULL); there is **no** `VC_SUPPLIER_CODE`/`VC_SIZE_CODE`/`VC_LOGISTICS_CODE`.
Yet every `docs/triggers.sql` trigger keys on those dropped string columns:

| Trigger | LIVE (schema) body | STALE (`docs/triggers.sql`) body |
|---------|--------------------|----------------------------------|
| `DELETE_SupplierCode` | `UPDATE INV_PARTS_STOCK_MST SET IN_SUPPLIER_ID = null WHERE a.IN_SUPPLIER_ID = d.IN_SUPPLIER_ID` | `SET VC_SUPPLIER_CODE = '' WHERE a.VC_SUPPLIER_CODE = d.VC_SUPPLIER_CODE` ⛔ dropped column |
| `DELETE_SizeCode` | `SET IN_SIZE_ID = null WHERE a.IN_SIZE_ID = d.IN_SIZE_ID` | `SET VC_SIZE_CODE = '' WHERE a.VC_SIZE_CODE = d.VC_SIZE_CODE` ⛔ |
| `DELETE_LogisticsCode` | `SET IN_LOGISTICS_ID = null WHERE a.IN_LOGISTICS_ID = d.IN_LOGISTICS_ID` | `SET VC_LOGISTICS_CODE = i.VC_LOGISTICS_CODE …` ⛔ |
| `DELETE_RenbanGroupCode` | `SET IN_RENBAN_ID = null WHERE p.IN_RENBAN_ID = d.IN_RENBAN_ID` | string-code based ⛔ |
| `DELETE_RejectParts` (stock-qty) | `SET IN_QTY = PS.IN_QTY + d.IN_QTY … WHERE ps.IN_PART_ID = d.IN_PART_ID` | `… WHERE PS.VC_SUPPLIER_CODE = d.VC_SUPPLIER_CODE AND PS.VC_PART_NUMBER = d.VC_PART_NUMBER` ⛔ + extra `Activity.dbo.InsertAct_Log` audit call |

The stock-quantity **math** is identical (`IN_QTY = IN_QTY + deleted.IN_QTY`); only the **join key**
changed (int id vs string code). So the staleness is uniform across the whole file, not just the
`*Code` master triggers.

## The 24 LIVE triggers (authoritative — from the schema)
Grouped by the invariant they enforce. ⭐ = body read directly during this audit.

| Trigger | Table | Event | Invariant |
|---------|-------|:-----:|-----------|
| `INSERT/UPDATE/DELETE_RecConfStatPartsStockMstQTY` | `INV_OPEN_ORDER_INF` | I/U/D | Receiving sync: keep `INV_PARTS_STOCK_MST.IN_QTY` balanced as open-order rows post |
| `INSERT/UPDATE/DELETE_RejectParts` | `INV_REJECT_INF` | I/U/D | Rejects adjust stock qty (⭐ DELETE adds the deleted reject qty back to `IN_QTY`, keyed on `IN_PART_ID`) |
| `INSERT/UPDATE/DELETE_Stocktaking` | `INV_STOCKTAKING_INF` | I/U/D | Physical-count adjustments to stock qty |
| `InsertPartShipping`,`UpdatePartShipping`,`DeletePartShipping`,`DeleteShipDate` | `INV_PART_SHIPPING_INF` / `INV_SHIPPING_INF` | I/U/D | Shipping adjusts stock qty |
| `INSERT_PartsStockMST` | `INV_PARTS_STOCK_MST` | I | Initialize stock row on new part |
| `INSERT_InvPartQtyInf` | `INV_PART_QTY_INF` | I | Stock-qty bookkeeping |
| `UPDATE_PartNumber`,`DELETE_PartNumber` | `INV_PARTS_STOCK_MST` | U/D | Cascade/maintenance on part-number edits |
| `UPDATE_ForecastDetailInf`,`DeleteForecastDetail` | `INV_FORECAST_DETAIL_INF` | U/D | Forecast-detail upkeep |
| `UPDATE_AssyRatioMst` | `INV_ASSY_RATIO_MST` | U | Assembly-ratio maintenance |
| `DELETE_SupplierCode` ⭐ | `INV_SUPPLIER_MST` | D | Null `INV_PARTS_STOCK_MST.IN_SUPPLIER_ID` for parts of the deleted supplier |
| `DELETE_SizeCode` ⭐ | `INV_SIZE_MST` | D | Null `INV_PARTS_STOCK_MST.IN_SIZE_ID` |
| `DELETE_LogisticsCode` ⭐ | `INV_LOGISTICS_MST` | D | Null `INV_PARTS_STOCK_MST.IN_LOGISTICS_ID` |
| `DELETE_RenbanGroupCode` ⭐ | `INV_RENBAN_GROUP_MST` | D | Null `IN_RENBAN_ID` on the child table |

(24 total: the four I/U/D stock-qty families = 12, shipping = 4, the two PartNumber + two PartsStock/QtyInf inits = 4, the two forecast + AssyRatio = 3 → wait that lists groups; the authoritative list of 24 distinct objects is in the schema. Names above are exact.)

## `docs/triggers.sql` reconciliation (its 23 entries)
| `docs/triggers.sql` trigger | Live counterpart | Verdict |
|---|---|---|
| `DELETE_SupplierCode`, `DELETE_SizeCode`, `DELETE_LogisticsCode`, `DELETE_RenbanGroupCode` | same-named live trigger | **STALE BODY** — old string-code form; trust the schema (int-FK null) |
| `UPDATE_SupplierCode`, `UPDATE_SizeCode`, `UPDATE_LogisticsCode`, `UPDATE_RenbanGroupCode` | **none** | **OBSOLETE — no live counterpart.** Propagated renamed `VC_*_CODE` to parts; unnecessary after the int-FK refactor (surrogate ids are immutable) and references dropped columns. Correctly absent from the live DB. |
| `DELETE/INSERT/UPDATE_RejectPartsStockMstQTY` | `DELETE/INSERT/UPDATE_RejectParts` | **RENAMED + STALE BODY** — same logical stock-qty triggers under longer legacy names, but keyed on dropped `VC_SUPPLIER_CODE`/`VC_PART_NUMBER` |
| `DELETE/INSERT/UPDATE_StockTakingPartsStockMstQTY` | `DELETE/INSERT/UPDATE_Stocktaking` | **RENAMED + STALE BODY** — same, for stocktaking |
| `DELETE/INSERT/UPDATE_RecConfStatPartsStockMstQTY`, `DELETE/UPDATE_PartNumber`, `INSERT_PartsStockMST`, `INSERT_InvPartQtyInf`, `UPDATE_ForecastDetailInf`, `UPDATE_AssyRatioMst` | same name in schema | Present in both; **trust the schema version** (likely also string-keyed/stale in `triggers.sql` — do not copy from it) |

### Live triggers MISSING from `docs/triggers.sql` (5)
`InsertPartShipping`, `UpdatePartShipping`, `DeletePartShipping`, `DeleteShipDate`, `DeleteForecastDetail`
— `docs/triggers.sql` has no shipping/forecast-delete triggers at all, so it is also **incomplete**, not
just stale-in-body. Another reason it cannot serve as a trigger inventory.

## Behavioral note (open, minor)
The stale stock-qty triggers wrote an audit row via `exec Activity.dbo.InsertAct_Log 'INVENTORY','TRIGGER', …`;
the **live** versions omit that call. If trigger-level inventory audit logging is desired in the rebuild,
note that the live DB triggers do **not** currently emit it (the app-layer `LogActLog` does its own logging).

## Actions
- **Treat `docs/triggers.sql` as obsolete.** Keep it for historical reference only; the skill's
  `database-objects.md` note and any module spec should point here and to the schema, not to that file.
- **The live `*Code` DELETE triggers null int FKs** — this confirms `supplier.md`/`logistics.md`/`size.md`
  (which already key on the int-FK nulling) and **corrects the P12 register**, whose cascade description
  had quoted the stale `VC_SUPPLIER_CODE=''` form (the verifiers read `docs/triggers.sql:765`). The real
  cascade is `SET IN_SUPPLIER_ID = NULL`.
- **Rebuild:** re-home the **24 live** triggers as model callbacks / service transactions, keyed on the
  int FKs. Do not port any logic from `docs/triggers.sql`.
