-- invoice_summary.sql  (Named Query: report_render/invoice_summary)  — CORRECTED (DEFAULT)
-- Wraps the D6 window-AWARE INVOICE Summary (M3 report R9). This is the corrected variant the rebuild
-- SHIPS by default (faithful↔corrected seam, m3-render-architecture.md §2.4; invoice_summary_faithful.sql
-- carries the legacy window-blind version behind the swap).
-- params: @PDate varchar(8)  (yyyymmdd — the caller fills only yyyymmdd; the hhmm tail is commented out
--         in MainMenu.pas:3617, so @PDate is effectively 8 chars even though the proc declares varchar(13))
--
-- D6 DIVERGENCE (DOCUMENTED, David-LOCKED): the legacy REPORT_INVOICESSummary JOINs INV_MANIFEST_COST_MST
-- on the assy code ALONE with NO production-date window, so a part with >1 price window MULTIPLIES the
-- invoice lines and OVER-BILLS (latent until a price change adds a 2nd window). This is the same window-
-- blind defect fixed for the 810/856. Per the LOCKED D6 decision this rebuild is window-aware via
-- dbo.fn_ManifestCostAt(part, production_date) (CROSS APPLY = inner-join semantics but window-correct: a
-- part with >1 window yields the ONE covering row; a line whose date has no covering window DROPS — the
-- window-blind legacy wrongly kept it). A report is NOT Toyota-facing/state-changing, so per the divergence
-- rule this is decide-and-flag (not a new David decision) — D6 is already locked. See
-- docs/analysis/reporting/spike-report-procs-d6.sql (the cutover proc) + report_defs.py note + the M3
-- finding text. INVOICE Summary = billing -> the NUMBERS matter; this is the deliberately-corrected one.
--
-- READ-ONLY pure SELECT. Requires dbo.fn_ManifestCostAt (PR #10) to exist.
-- IG81-COMPAT: CROSS APPLY + inline TVF run identically on 8.1.52 and 8.3.
-- IG83-TODO: promote invoice_summary.sql to a Named Query; swap runPrepQuery -> runNamedQuery.

SELECT  d.VC_ASSY_PART_NUMBER  'PartNumber',
        d.VC_MANIFEST_NUMBER   'Manifest',
        d.IN_QTY               'ShipQty',
        m.MO_PRICE             'UnitPrice',
        a.VC_PRODUCTION_DATE   'PickUpDate'
FROM    INV_ASN_MST        a
JOIN    INV_ASN_DETAIL_MST d ON a.IN_ASN_ID = d.IN_ASN_ID
CROSS APPLY dbo.fn_ManifestCostAt(d.VC_ASSY_PART_NUMBER, a.VC_PRODUCTION_DATE) m   -- D6: was window-blind JOIN
JOIN    INV_INV_MST        i ON i.IN_INV_ID = d.IN_INV_ID
WHERE   a.VC_PRODUCTION_DATE = ?            -- @PDate
ORDER BY d.VC_MANIFEST_NUMBER
