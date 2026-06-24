/* =====================================================================
   spike-order-file-feed.sql — the order-FILE-GENERATOR feed (the aliased
   replacement for SELECT_OrderNotOrdered's `SELECT *`).

   This is the canonical, reviewable copy of order_file/code.py's _FEED_SQL.
   The test (test_order_file_e2e.py) enforces that _FEED_SQL equals this
   SELECT body with the bind sites identical, so the two cannot drift.

   WHY AN ALIASED FEED (legacy H10 / order-file-data-analysis §2.3):
   The legacy proc SELECT_OrderNotOrdered is `SELECT *` over a 4-table join
   (88 columns) in which IN_QTY, VC_PART_NUMBER, VC_KANBAN_NUMBER and
   VC_SUPPLIER_CODE are DUPLICATED across the joined tables. Delphi
   `fieldbyname` silently resolves to the FIRST occurrence (the open-order
   copy). The order quantity MUST be the OPEN-ORDER IN_QTY (i.IN_QTY, e.g.
   1200), NOT the parts-stock on-hand IN_QTY (p.IN_QTY, e.g. 28133). Every
   read column below is aliased EXPLICITLY off the `i.` (open-order) table so
   the wrong shadow column can never be picked.

   ORDER BY is byte-identical to the proc — the supplier/renban sort is
   load-bearing (the generator's supplier-break detection relies on it).

   WHERE — RISK-2: the legacy proc filtered `((VC_ORDER_DATE IS NULL) OR ='')`,
   but the stamp UPDATE_ORDEROrderDate clears via `VC_ORDER_DATE <> @OrderDate`,
   and under ANSI_NULLS ON `NULL <> '...'` is UNKNOWN -> a NULL row is NEVER
   stamped, so the feed would re-emit it forever. VC_ORDER_DATE is NOT NULL
   DEFAULT '' on the live schema (verified), so the IS NULL arm is unreachable.
   We drop it: the feed selects ONLY rows the stamp can clear (= '').

   Connection (Ignition): Inventory_Spike.
   IG81-COMPAT: plain T-SQL; identical on 8.1.52 and 8.3.
   ===================================================================== */

SELECT i.VC_SUPPLIER_CODE AS supCode,        -- the OPEN-ORDER supplier code (i., not s./p.)
       i.VC_FRS_NUMBER    AS frsNum,
       i.VC_RENBAN_NUMBER AS renban,
       i.VC_PART_NUMBER   AS partNum,
       i.IN_QTY           AS orderQty,        -- the OPEN-ORDER qty (NOT p.IN_QTY on-hand) — H10
       i.VC_KANBAN_NUMBER AS kanban,          -- the OPEN-ORDER kanban (i., not p.)
       p.VC_PARTS_NAME    AS partsName,       -- parts-master (Excel order-sheet only)
       p.IN_1LOTQTY       AS oneLotQty        -- parts-master (wheel order-sheet only)
FROM INV_OPEN_ORDER_INF i
     JOIN INV_PARTS_STOCK_MST p ON i.VC_PART_NUMBER = p.VC_PART_NUMBER
     JOIN INV_SUPPLIER_MST   s ON p.IN_SUPPLIER_ID = s.IN_SUPPLIER_ID
     LEFT OUTER JOIN INV_RENBAN_GROUP_MST r ON p.IN_RENBAN_ID = r.IN_RENBAN_ID
WHERE i.VC_ORDER_DATE = ''                   -- RISK-2: only what the stamp can clear (no IS NULL arm)
  AND i.VC_RENBAN_NUMBER <> ''
ORDER BY s.VC_SUPPLIER_CODE, i.VC_RENBAN_NUMBER;
