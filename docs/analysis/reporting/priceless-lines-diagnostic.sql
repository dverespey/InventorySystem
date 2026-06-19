-- priceless-lines-diagnostic.sql
-- CUTOVER artifact — D6 priceless-lines decision (David 2026-06-19).
--
-- DECISION (verbatim-faithful): "D6 priceless lines — CROSS APPLY (keep, all 4 procs) + ADD a pre-invoice
-- diagnostic. Keep CROSS APPLY (faithful to legacy; never emits a $0 line into the EDI 810/856 to Toyota).
-- Build a diagnostic that surfaces priceless lines so gaps are caught BEFORE billing."
--
-- WHY. The migrated D6 report procs (spike-report-procs-d6.sql) use
--     CROSS APPLY dbo.fn_ManifestCostAt(d.VC_ASSY_PART_NUMBER, a.VC_PRODUCTION_DATE) m
-- CROSS APPLY = inner-join semantics: a shipped detail line whose production date has NO covering manifest-
-- cost window is DROPPED from the report (no $0 line ever goes to Toyota). That is the faithful, correct
-- behavior — but a dropped line is a SILENT under-bill. This diagnostic is the safety net: it surfaces
-- EXACTLY the lines CROSS APPLY would drop, so a missing/expired manifest-cost window is caught and fixed
-- BEFORE the EDI 810/856 run, not discovered as a billing gap afterward.
--
-- HOW. It mirrors the EXACT join each D6 proc uses, but flips CROSS APPLY -> OUTER APPLY and keeps only the
-- rows where the lookup found nothing (m.IN_MANIFEST_COST_ID IS NULL == "no covering window"). There are
-- TWO billing feeds with MATERIALLY DIFFERENT shipped filters, so this file carries TWO diagnostics — one
-- per feed — because a single 810-shaped query silently leaves the 856 feed unguarded:
--
--   EDI810 (@EIN=0) scope:  a.VC_ASN_STATUS = 'A'  AND  d.IN_INV_ID IS NULL
--     (the unbilled flag lives on the ASN DETAIL row — INV_ASN_DETAIL_MST.IN_INV_ID — there is no IN_INV_ID
--      on the ASN header. This is the EDI810 join's own predicate, verified against the live schema.)
--
--   EDI856 (@EIN=0) scope:  a.VC_ASN_STATUS = 'C'  + an INNER JOIN INV_FORECAST_DETAIL_INF f
--                           ON d.VC_ASSY_PART_NUMBER = f.VC_ASSY_PART_NUMBER_CODE, with the same GROUP BY
--     fan-out the 856 proc uses. The forecast JOIN is kept as an INNER join EXACTLY as 856 has it: a status-C
--     line whose part has NO forecast-detail row is dropped by 856 for a DIFFERENT reason (no forecast match,
--     not a missing price window) and is NOT a priceless line — surfacing it here would be a false positive.
--     So we reproduce 856's full pre-cost row set (status-C ∩ forecast match), then flag the subset whose
--     manifest-cost lookup misses. Verified against the LIVE REPORT_EDI856 body in /tmp/inv_utf8.sql
--     (lines 3632-3664): WHERE a.VC_ASN_STATUS='C', the forecast inner JOIN, and the legacy window predicate
--     m.VC_START_MANIFEST<=date AND m.VC_END_MANIFEST>=date that D6 replaced with CROSS APPLY fn_ManifestCostAt.
--     (The pre-D6 856 was already window-aware via that inline predicate — fn_ManifestCostAt is the faithful
--      route for it; the same OUTER APPLY ... m.IN_MANIFEST_COST_ID IS NULL surfaces what CROSS APPLY drops.)
--
-- COLUMN VERIFICATION (live spike, INV_ASN_MST / INV_ASN_DETAIL_MST):
--   INV_ASN_MST:        IN_ASN_ID, VC_ASN_STATUS(varchar 1), VC_PRODUCTION_DATE(varchar 8)
--   INV_ASN_DETAIL_MST: IN_ASN_ID, IN_INV_ID, VC_MANIFEST_NUMBER(varchar 8), VC_ASSY_PART_NUMBER(varchar 12), IN_QTY
--   fn_ManifestCostAt returns: IN_MANIFEST_COST_ID, VC_ASSY_MANIFEST_NUMBER, MO_PRICE
--
-- USAGE: run BOTH diagnostics below before the EDI 810/856 billing run. Expect ZERO rows from each. Any row
-- = a shipped line that the matching feed will DROP (under-billed) — add/extend the part's manifest-cost
-- window, then re-run to zero. The 810 query guards the status-'A'/unbilled invoice feed; the 856 query
-- guards the status-'C' ASN feed (the bigger one). Run the one(s) matching the feed you are about to send,
-- or both before a combined run.
--
-- IG81-COMPAT: OUTER APPLY + an inline TVF run identically on 8.1.52 and 8.3.
-- IG83-TODO: the string-date compare inside fn_ManifestCostAt is correct only while VC_PRODUCTION_DATE /
--            VC_START/END_MANIFEST stay zero-padded yyyymmdd (lexicographic == chronological).

-- =============================================================================================
-- (A1) EDI810 priceless-lines diagnostic QUERY — mirrors REPORT_EDI810 (@EIN=0). Run before the 810 feed.
--      Label: '810' result set.
-- =============================================================================================
SELECT  '810'                          AS Feed,
        d.VC_ASSY_PART_NUMBER          AS PartNumber,
        d.VC_MANIFEST_NUMBER           AS Manifest,
        a.VC_PRODUCTION_DATE           AS ProductionDate,
        d.IN_QTY                       AS ShipQty,
        d.IN_ASN_ID                    AS ASNid,
        d.IN_ASN_DETAIL_ID             AS ASNdetailId
FROM    INV_ASN_MST a
JOIN    INV_ASN_DETAIL_MST d
        ON a.IN_ASN_ID = d.IN_ASN_ID
       AND a.VC_ASN_STATUS = 'A'        -- same shipped filter as REPORT_EDI810 (@EIN=0)
       AND d.IN_INV_ID IS NULL          -- same unbilled filter (detail-level)
OUTER APPLY dbo.fn_ManifestCostAt(d.VC_ASSY_PART_NUMBER, a.VC_PRODUCTION_DATE) m
WHERE   m.IN_MANIFEST_COST_ID IS NULL   -- NO covering manifest-cost window -> CROSS APPLY would DROP this line
ORDER BY d.VC_ASSY_PART_NUMBER, a.VC_PRODUCTION_DATE;
GO

-- =============================================================================================
-- (A2) EDI856 priceless-lines diagnostic QUERY — mirrors REPORT_EDI856 (@EIN=0) EXACTLY. Run before the
--      856 feed. Label: '856' result set.
--      Shape copied verbatim from the live REPORT_EDI856 @EIN=0 body (/tmp/inv_utf8.sql:3632-3664):
--        - WHERE a.VC_ASN_STATUS = 'C'                                  (NOT 'A' — the 856 feed's filter)
--        - INNER JOIN INV_FORECAST_DETAIL_INF f ON d.VC_ASSY_PART_NUMBER = f.VC_ASSY_PART_NUMBER_CODE
--          (kept INNER, exactly as 856 — a part with no forecast row is dropped by 856 anyway; not priceless)
--        - the SAME GROUP BY fan-out (so multi-forecast-row parts fan out IDENTICALLY to the 856 report,
--          surfacing precisely the rows 856's CROSS APPLY would drop — no false +/-).
--      The legacy window predicate (m.VC_START_MANIFEST<=date AND m.VC_END_MANIFEST>=date) that D6 replaced
--      with CROSS APPLY is reproduced as OUTER APPLY fn_ManifestCostAt ... WHERE m.IN_MANIFEST_COST_ID IS NULL.
-- =============================================================================================
SELECT  '856'                          AS Feed,
        d.VC_ASSY_PART_NUMBER          AS PartNumber,
        d.VC_MANIFEST_NUMBER           AS Manifest,
        a.VC_PRODUCTION_DATE           AS ProductionDate,
        d.IN_QTY                       AS ShipQty,
        f.VC_ASSY_KANBAN_NUMBER        AS Kanban,
        a.IN_ASN_EIN                   AS SiteEIN,
        a.VC_START_SEQ_NUMBER          AS StartSeq,
        a.VC_LINE_NAME                 AS LineName,
        d.IN_ASN_ID                    AS ASNid
FROM    INV_ASN_MST a
JOIN    INV_ASN_DETAIL_MST d ON a.IN_ASN_ID = d.IN_ASN_ID
JOIN    INV_FORECAST_DETAIL_INF f ON d.VC_ASSY_PART_NUMBER = f.VC_ASSY_PART_NUMBER_CODE
OUTER APPLY dbo.fn_ManifestCostAt(d.VC_ASSY_PART_NUMBER, a.VC_PRODUCTION_DATE) m
WHERE   a.VC_ASN_STATUS = 'C'           -- same shipped filter as REPORT_EDI856 (@EIN=0)
  AND   m.IN_MANIFEST_COST_ID IS NULL   -- NO covering manifest-cost window -> CROSS APPLY would DROP this line
GROUP BY d.VC_ASSY_PART_NUMBER, d.VC_MANIFEST_NUMBER, a.VC_PRODUCTION_DATE, d.IN_QTY,
         f.VC_ASSY_KANBAN_NUMBER, a.IN_ASN_EIN, a.VC_START_SEQ_NUMBER, a.VC_LINE_NAME, d.IN_ASN_ID
ORDER BY d.VC_ASSY_PART_NUMBER, a.VC_PRODUCTION_DATE;
GO

-- =============================================================================================
-- (B1) Optional wrapped proc for the 810 feed (same logic as A1) — for a Named Query / scheduled check.
--      Named to mirror the report family it guards (IA Named-Query practice).
-- =============================================================================================
IF OBJECT_ID('dbo.REPORT_EDI810_PricelessLines', 'P') IS NOT NULL
    DROP PROCEDURE dbo.REPORT_EDI810_PricelessLines;
GO
CREATE PROCEDURE dbo.REPORT_EDI810_PricelessLines
AS
    SELECT  d.VC_ASSY_PART_NUMBER          AS PartNumber,
            d.VC_MANIFEST_NUMBER           AS Manifest,
            a.VC_PRODUCTION_DATE           AS ProductionDate,
            d.IN_QTY                       AS ShipQty,
            d.IN_ASN_ID                    AS ASNid,
            d.IN_ASN_DETAIL_ID             AS ASNdetailId
    FROM    INV_ASN_MST a
    JOIN    INV_ASN_DETAIL_MST d
            ON a.IN_ASN_ID = d.IN_ASN_ID
           AND a.VC_ASN_STATUS = 'A'
           AND d.IN_INV_ID IS NULL
    OUTER APPLY dbo.fn_ManifestCostAt(d.VC_ASSY_PART_NUMBER, a.VC_PRODUCTION_DATE) m
    WHERE   m.IN_MANIFEST_COST_ID IS NULL
    ORDER BY d.VC_ASSY_PART_NUMBER, a.VC_PRODUCTION_DATE;
GO
PRINT 'REPORT_EDI810_PricelessLines created (pre-invoice priceless-line diagnostic, 810 feed).';
GO

-- =============================================================================================
-- (B2) Optional wrapped proc for the 856 feed (same logic as A2) — mirrors REPORT_EDI856 (@EIN=0).
-- =============================================================================================
IF OBJECT_ID('dbo.REPORT_EDI856_PricelessLines', 'P') IS NOT NULL
    DROP PROCEDURE dbo.REPORT_EDI856_PricelessLines;
GO
CREATE PROCEDURE dbo.REPORT_EDI856_PricelessLines
AS
    SELECT  d.VC_ASSY_PART_NUMBER          AS PartNumber,
            d.VC_MANIFEST_NUMBER           AS Manifest,
            a.VC_PRODUCTION_DATE           AS ProductionDate,
            d.IN_QTY                       AS ShipQty,
            f.VC_ASSY_KANBAN_NUMBER        AS Kanban,
            a.IN_ASN_EIN                   AS SiteEIN,
            a.VC_START_SEQ_NUMBER          AS StartSeq,
            a.VC_LINE_NAME                 AS LineName,
            d.IN_ASN_ID                    AS ASNid
    FROM    INV_ASN_MST a
    JOIN    INV_ASN_DETAIL_MST d ON a.IN_ASN_ID = d.IN_ASN_ID
    JOIN    INV_FORECAST_DETAIL_INF f ON d.VC_ASSY_PART_NUMBER = f.VC_ASSY_PART_NUMBER_CODE
    OUTER APPLY dbo.fn_ManifestCostAt(d.VC_ASSY_PART_NUMBER, a.VC_PRODUCTION_DATE) m
    WHERE   a.VC_ASN_STATUS = 'C'
      AND   m.IN_MANIFEST_COST_ID IS NULL
    GROUP BY d.VC_ASSY_PART_NUMBER, d.VC_MANIFEST_NUMBER, a.VC_PRODUCTION_DATE, d.IN_QTY,
             f.VC_ASSY_KANBAN_NUMBER, a.IN_ASN_EIN, a.VC_START_SEQ_NUMBER, a.VC_LINE_NAME, d.IN_ASN_ID
    ORDER BY d.VC_ASSY_PART_NUMBER, a.VC_PRODUCTION_DATE;
GO
PRINT 'REPORT_EDI856_PricelessLines created (pre-invoice priceless-line diagnostic, 856 feed).';
GO
