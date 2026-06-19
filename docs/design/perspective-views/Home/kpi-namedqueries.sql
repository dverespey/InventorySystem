-- =============================================================================
-- Home hub KPI Named Queries  (Inventory_Spike)
--
-- The headless build loads KPIs with inline system.db.runPrepQuery in the
-- Home/Home view binding (see scripts/gen_landing_view.py). Named Queries can't
-- be authored headlessly (the NQ data.bin is Designer-only), so this file is the
-- ready-to-paste SQL for the eventual migration. In the Designer:
--
--   1. Create Named Query  Home/kpiSummary   (Type: Query, DB: Inventory_Spike)
--      -> paste KPI_SUMMARY below.  Returns ONE row, all counts.
--   2. Create Named Query  Home/kpiTopShort  (Type: Query, DB: Inventory_Spike)
--      -> paste KPI_TOP_SHORT below.  Returns the single worst-short part.
--   3. In Home/Home, replace the custom.kpi script transform's runPrepQuery calls
--      with: data = system.db.runNamedQuery("Home/kpiSummary", {})
--             top  = system.db.runNamedQuery("Home/kpiTopShort", {})
--      then build the same display strings from data.getValueAt(0,"<Col>").
--
-- IG-SITE: single-site spike — no site predicate yet. When multi-site lands, add a
--          :siteId NQ param and  AND site_id = :siteId  to each subquery (the
--          INV_PARTS_STOCK_MST.site_id column already exists; others need plumbing).
-- =============================================================================

-- ---- Home/kpiSummary ---------------------------------------------------------
SELECT
    (SELECT COUNT(*)                       FROM INV_OPEN_ORDER_INF WHERE ISNULL(VC_TERMINATED,'') = '') AS OpenOrders,
    (SELECT COUNT(DISTINCT VC_SUPPLIER_CODE) FROM INV_OPEN_ORDER_INF WHERE ISNULL(VC_TERMINATED,'') = '') AS OpenSuppliers,
    (SELECT COUNT(*)                       FROM INV_ASN_MST WHERE VC_ASN_STATUS = 'A')                  AS ActiveAsns,
    (SELECT COUNT(*)                       FROM INV_ASN_DETAIL_MST)                                     AS AsnLines,
    (SELECT COUNT(*)                       FROM INV_SHIPPING_INF)                                       AS Shipments,
    (SELECT COUNT(*)                       FROM INV_PARTS_STOCK_MST)                                    AS PartsTotal,
    (SELECT COUNT(*)                       FROM INV_PARTS_STOCK_MST WHERE IN_QTY <= 0)                  AS PartsShort,
    (SELECT COUNT(*)                       FROM INV_PARTS_STOCK_MST WHERE IN_QTY < IN_1LOTQTY)          AS PartsBelowLot,
    (SELECT COUNT(*)                       FROM INV_STOCKTAKING_INF)                                    AS StocktakeEntries,
    (SELECT COUNT(*)                       FROM INV_SUPPLIER_MST)                                       AS Suppliers;

-- ---- Home/kpiTopShort --------------------------------------------------------
SELECT TOP 1 VC_PARTS_NAME AS PartName, IN_QTY AS Qty
FROM   INV_PARTS_STOCK_MST
WHERE  IN_QTY <= 0
ORDER  BY IN_QTY ASC;
