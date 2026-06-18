-- spike-report-procs-d6.sql
-- D6 / D11#1 MIGRATION of the window-blind manifest-cost report procs to the shared window-aware
-- lookup dbo.fn_ManifestCostAt (docs/analysis/master-data/spike-manifest-cost-lookup.sql).
--
-- Each proc JOINed INV_MANIFEST_COST_MST on the assy code ALONE (no date window), so a part with >1
-- price window multiplies invoice/ASN lines and over-bills (latent until a price change adds a window).
-- The fix replaces every window-blind JOIN with:
--     CROSS APPLY dbo.fn_ManifestCostAt(d.VC_ASSY_PART_NUMBER, a.VC_PRODUCTION_DATE) m
-- CROSS APPLY = inner-join semantics, but window-CORRECT: a line whose production date has no covering
-- window now DROPS OUT (the window-blind legacy JOIN wrongly KEPT it — it had no date filter at all),
-- and a part with >1 window yields ONE row (the covering one) instead of legacy's N. The 2nd arg is
-- ALWAYS the per-row 8-char VC_PRODUCTION_DATE (never a 6-char @PDate month — see
-- REPORT_MonthlyINVOICESSummary). Boundary is inclusive (D6 decision). (REPORT_EDI856 @EIN=0 was the one
-- branch already window-aware; routing it through the TVF too is a faithful no-op for consistency.)
--
-- This is the CUTOVER artifact (replaces the legacy procs). During parallel run the legacy procs stay
-- live; the migration is proven by scripts/e2e/test_report_procs_d6.py (which deploys _D6-suffixed copies
-- side-by-side and diffs them against the legacy procs). Requires fn_ManifestCostAt (PR #10) to exist.
--
-- IG81-COMPAT: CROSS APPLY + an inline TVF run identically on 8.1.52 and 8.3.

-- =============================================================================================
IF OBJECT_ID('dbo.REPORT_INVOICESSummary', 'P') IS NOT NULL DROP PROCEDURE dbo.REPORT_INVOICESSummary;
GO
CREATE PROCEDURE dbo.REPORT_INVOICESSummary
    @PDate varchar(13)
AS
    SELECT  d.VC_MANIFEST_NUMBER 'Manifest', d.VC_ASSY_PART_NUMBER 'PartNumber',
            m.MO_PRICE 'UnitPrice', d.IN_QTY 'ShipQty', a.VC_PRODUCTION_DATE 'PickUpDate'
    FROM    INV_ASN_MST a
    JOIN    INV_ASN_DETAIL_MST d ON a.IN_ASN_ID = d.IN_ASN_ID
    CROSS APPLY dbo.fn_ManifestCostAt(d.VC_ASSY_PART_NUMBER, a.VC_PRODUCTION_DATE) m   -- was window-blind JOIN
    JOIN    INV_INV_MST i ON i.IN_INV_ID = d.IN_INV_ID
    WHERE   a.VC_PRODUCTION_DATE = @PDate
    ORDER BY d.VC_MANIFEST_NUMBER;
GO
PRINT 'REPORT_INVOICESSummary migrated (D6).';
GO

-- =============================================================================================
IF OBJECT_ID('dbo.REPORT_MonthlyINVOICESSummary', 'P') IS NOT NULL DROP PROCEDURE dbo.REPORT_MonthlyINVOICESSummary;
GO
CREATE PROCEDURE dbo.REPORT_MonthlyINVOICESSummary
    @PDate varchar(6)
AS
    -- @PDate is a 6-char MONTH (the WHERE filter); the TVF still takes the PER-ROW 8-char date.
    SELECT  d.VC_MANIFEST_NUMBER 'Manifest', d.VC_ASSY_PART_NUMBER 'PartNumber',
            m.MO_PRICE 'UnitPrice', d.IN_QTY 'ShipQty', a.VC_PRODUCTION_DATE 'PickUpDate'
    FROM    INV_ASN_MST a
    JOIN    INV_ASN_DETAIL_MST d ON a.IN_ASN_ID = d.IN_ASN_ID
    CROSS APPLY dbo.fn_ManifestCostAt(d.VC_ASSY_PART_NUMBER, a.VC_PRODUCTION_DATE) m   -- per-row date, NOT @PDate
    JOIN    INV_INV_MST i ON i.IN_INV_ID = d.IN_INV_ID
    WHERE   substring(a.VC_PRODUCTION_DATE, 1, 6) = @PDate
    ORDER BY d.VC_MANIFEST_NUMBER;
GO
PRINT 'REPORT_MonthlyINVOICESSummary migrated (D6).';
GO

-- =============================================================================================
IF OBJECT_ID('dbo.REPORT_EDI810', 'P') IS NOT NULL DROP PROCEDURE dbo.REPORT_EDI810;
GO
CREATE PROCEDURE dbo.REPORT_EDI810
    @EIN INTEGER = 0
AS
    IF @EIN = 0
    BEGIN
        SELECT  d.VC_MANIFEST_NUMBER 'Manifest', d.VC_ASSY_PART_NUMBER 'PartNumber',
                m.MO_PRICE 'UnitPrice', d.IN_QTY 'ShipQty', a.VC_PRODUCTION_DATE 'PickUpDate',
                d.IN_ASN_ID 'ASNid'
        FROM    INV_ASN_MST a
        JOIN    INV_ASN_DETAIL_MST d
                ON a.IN_ASN_ID = d.IN_ASN_ID AND a.VC_ASN_STATUS = 'A' AND d.IN_INV_ID is null
        CROSS APPLY dbo.fn_ManifestCostAt(d.VC_ASSY_PART_NUMBER, a.VC_PRODUCTION_DATE) m
        ORDER BY d.VC_MANIFEST_NUMBER;
    END
    ELSE
    BEGIN
        SELECT  d.VC_MANIFEST_NUMBER 'Manifest', d.VC_ASSY_PART_NUMBER 'PartNumber',
                m.MO_PRICE 'UnitPrice', d.IN_QTY 'ShipQty', a.VC_PRODUCTION_DATE 'PickUpDate',
                d.IN_ASN_ID 'ASNid', iim.IN_INV_EIN 'Site EIN'
        FROM    INV_ASN_MST a
        JOIN    INV_ASN_DETAIL_MST d ON a.IN_ASN_ID = d.IN_ASN_ID
        JOIN    INV_INV_MST iim ON d.IN_INV_ID = iim.IN_INV_ID
        CROSS APPLY dbo.fn_ManifestCostAt(d.VC_ASSY_PART_NUMBER, a.VC_PRODUCTION_DATE) m
        WHERE   iim.IN_INV_EIN = @EIN
        ORDER BY d.VC_MANIFEST_NUMBER;

        UPDATE INV_INV_MST SET VC_INV_STATUS = 'S' WHERE IN_INV_EIN = @EIN;
    END
GO
PRINT 'REPORT_EDI810 migrated (D6).';
GO

-- =============================================================================================
IF OBJECT_ID('dbo.REPORT_EDI856', 'P') IS NOT NULL DROP PROCEDURE dbo.REPORT_EDI856;
GO
CREATE PROCEDURE dbo.REPORT_EDI856
    @EIN INTEGER = 0
AS
    IF @EIN = 0
    BEGIN
        SELECT  d.VC_MANIFEST_NUMBER 'Manifest', d.VC_ASSY_PART_NUMBER 'PartNumber',
                m.MO_PRICE 'UnitPrice', d.IN_QTY 'ShipQty', a.VC_PRODUCTION_DATE 'PickUpDate',
                f.VC_ASSY_KANBAN_NUMBER 'Kanban', a.IN_ASN_EIN 'SiteEIN',
                VC_START_SEQ_NUMBER 'StartSeq', a.VC_LINE_NAME 'LineName'
        FROM    INV_ASN_MST a
        JOIN    INV_ASN_DETAIL_MST d ON a.IN_ASN_ID = d.IN_ASN_ID
        CROSS APPLY dbo.fn_ManifestCostAt(d.VC_ASSY_PART_NUMBER, a.VC_PRODUCTION_DATE) m   -- was windowed JOIN
        JOIN    INV_FORECAST_DETAIL_INF f ON d.VC_ASSY_PART_NUMBER = f.VC_ASSY_PART_NUMBER_CODE
        WHERE   a.VC_ASN_STATUS = 'C'
        GROUP BY a.IN_ASN_EIN, d.VC_MANIFEST_NUMBER, d.VC_ASSY_PART_NUMBER, m.MO_PRICE, d.IN_QTY,
                 a.VC_PRODUCTION_DATE, f.VC_ASSY_KANBAN_NUMBER, VC_START_SEQ_NUMBER, a.VC_LINE_NAME;
    END
    ELSE
    BEGIN
        SELECT  d.VC_MANIFEST_NUMBER 'Manifest', d.VC_ASSY_PART_NUMBER 'PartNumber',
                m.MO_PRICE 'UnitPrice', d.IN_QTY 'ShipQty', a.VC_PRODUCTION_DATE 'PickUpDate',
                f.VC_ASSY_KANBAN_NUMBER 'Kanban', a.IN_ASN_EIN 'SiteEIN', VC_START_SEQ_NUMBER 'StartSeq'
        FROM    INV_ASN_MST a
        JOIN    INV_ASN_DETAIL_MST d ON a.IN_ASN_ID = d.IN_ASN_ID
        CROSS APPLY dbo.fn_ManifestCostAt(d.VC_ASSY_PART_NUMBER, a.VC_PRODUCTION_DATE) m
        JOIN    INV_FORECAST_DETAIL_INF f ON d.VC_ASSY_PART_NUMBER = f.VC_ASSY_PART_NUMBER_CODE
        WHERE   a.IN_ASN_EIN = @EIN                          -- D6 fix: was a hardcoded site EIN that mismatched the UPDATE's @EIN
        GROUP BY a.IN_ASN_EIN, d.VC_MANIFEST_NUMBER, d.VC_ASSY_PART_NUMBER, m.MO_PRICE, d.IN_QTY,
                 a.VC_PRODUCTION_DATE, f.VC_ASSY_KANBAN_NUMBER, VC_START_SEQ_NUMBER;

        UPDATE INV_ASN_MST SET VC_ASN_STATUS = 'S' WHERE IN_ASN_EIN = @EIN;
    END
GO
PRINT 'REPORT_EDI856 migrated (D6).';
GO
