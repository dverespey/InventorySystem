-- spike-manifest-cost-lookup.sql
-- fn_ManifestCostAt — the ONE shared, window-aware manifest-cost lookup (D6 / D11#1).
-- Design refs: docs/analysis/master-data/manifest-cost.md; docs/analysis/edi/asn-invoice.md;
--              docs/analysis/reporting/reporting.md.
--
-- INV_MANIFEST_COST_MST prices a part for a DATE WINDOW: (VC_ASSY_PART_NUMBER_CODE,
-- VC_START_MANIFEST..VC_END_MANIFEST] -> MO_PRICE. The CORRECT lookup picks the row whose window
-- COVERS the production/pickup date — exactly what REPORT_EDI856 (@EIN=0 branch) does:
--     AND m.VC_START_MANIFEST <= a.VC_PRODUCTION_DATE
--     AND m.VC_END_MANIFEST   >= a.VC_PRODUCTION_DATE
-- But REPORT_INVOICESSummary, REPORT_MonthlyINVOICESSummary, the REPORT_EDI810 path, and the
-- REPORT_EDI856 @EIN!=0 branch JOIN the cost master on part code ALONE (no window filter) — they are
-- WINDOW-BLIND. Today that's latent (the live data has exactly one cost row per part). It becomes a
-- bug the moment a part has >1 row with a covering window. NOTE on reachability: the table's UNIQUE
-- index IX_INV_MANIFEST_COST_MST is on VC_ASSY_MANIFEST_NUMBER (the 2-char manifest #), NOT on the
-- assy part code — so a part CAN already carry multiple windows under DISTINCT manifest numbers (the
-- natural way a price change is recorded), independent of that index. (The checked-in dump still has
-- the index; the spike DB dropped it in the ManifestCost master build — either way the part-level
-- multi-window the bug needs is reachable.) When it is, every window-blind query returns BOTH rows
-- -> row multiplication + a wrong (often doubled) invoice/ASN total.
--
-- This function centralizes the window logic so the fix is a single-point edit (the IA Named-Query /
-- one-lookup-mirrors-the-rule practice): every report NQ CROSS APPLYs it instead of re-deriving the
-- (often missing) window filter inline.
--
-- BOUNDARY + OVERLAP — RESOLVED (David 2026-06-18):
--   (Q-boundary) INCLUSIVE (<= / >=). VC_END_MANIFEST is the LAST day the price is effective (the live
--     data confirms it: end dates like 20280630/20270627 are real last-effective days, not next-start
--     days). So a production date ON a window's start OR end is IN that window and IS priced. This TVF's
--     <= / >= is correct as-is. The legacy SELECT_ManifestCost;1 (the master/UI date lookup) used STRICT
--     (> / <), which silently DROPPED the first AND last effective day of every window (no price on a
--     price-change day) — a latent bug. SELECT_ManifestCost is therefore SUPERSEDED by this TVF, not
--     ported: the rebuilt master "what price on date D" lookup SELECTs fn_ManifestCostAt so the UI and
--     the reports agree on "in window" (they disagreed in legacy).
--   (Q-overlap) GAP convention (David): adjacent windows are entered with a 1-day gap (…1231 then 0101…),
--     so they NEVER share a boundary day -> inclusive-both never double-matches -> windows are
--     non-overlapping by data-entry rule. The TOP 1 below is therefore dead-defense (it never
--     disambiguates a real overlap). ENFORCE the gap at WRITE in the ManifestCost master (reject a new
--     window that overlaps an existing one for the part) so the rule can't be violated; the overlap
--     diagnostic at the bottom is the interim/cutover check (expect ZERO rows). [Master no-overlap guard:
--     follow-up on the ManifestCost master CRUD module.]
-- IG81-COMPAT: a plain inline TVF, identical on 8.1.52 and 8.3; report Named Queries CROSS APPLY it.
-- IG83-TODO: at the Postgres phase, VC_START/VC_END_MANIFEST (yyyymmdd strings) -> date; the string
--            comparison here is correct ONLY because the stamps are zero-padded yyyymmdd (lexicographic
--            == chronological). Preserve that invariant or switch to a typed date column.

IF OBJECT_ID('dbo.fn_ManifestCostAt', 'IF') IS NOT NULL
    DROP FUNCTION dbo.fn_ManifestCostAt;
GO

CREATE FUNCTION dbo.fn_ManifestCostAt (@partCode varchar(12), @date varchar(8))
RETURNS TABLE
AS RETURN
    -- The single cost row whose manifest window covers @date. TOP 1 + ORDER BY makes the result
    -- deterministic if a data error ever leaves OVERLAPPING windows for a part (newest start wins);
    -- with clean (non-overlapping) windows there is at most one match and the ordering is moot.
    SELECT TOP 1
           m.IN_MANIFEST_COST_ID,
           m.VC_ASSY_MANIFEST_NUMBER,
           m.MO_PRICE
    FROM dbo.INV_MANIFEST_COST_MST m
    WHERE m.VC_ASSY_PART_NUMBER_CODE = @partCode
      AND m.VC_START_MANIFEST <= @date
      AND m.VC_END_MANIFEST   >= @date
    ORDER BY m.VC_START_MANIFEST DESC, m.IN_MANIFEST_COST_ID DESC;
GO

PRINT 'fn_ManifestCostAt created.';
GO

-- ---------------------------------------------------------------------------------------------
-- CALLER MIGRATION (window-blind -> window-aware). Replace, in REPORT_INVOICESSummary /
-- REPORT_MonthlyINVOICESSummary / REPORT_EDI810 / REPORT_EDI856(@EIN!=0):
--
--     JOIN INV_MANIFEST_COST_MST m ON d.VC_ASSY_PART_NUMBER = m.VC_ASSY_PART_NUMBER_CODE
--   (no window filter)  -->
--     CROSS APPLY dbo.fn_ManifestCostAt(d.VC_ASSY_PART_NUMBER, a.VC_PRODUCTION_DATE) m
--
-- CROSS APPLY keeps inner-join semantics (a detail line with no covering window drops out, matching
-- the legacy inner JOIN's behavior for a date outside every window). Use OUTER APPLY only if a
-- report must keep priceless lines (a domain call — flag to David). Also drop the hardcoded
-- IN_ASN_EIN=6440 in the REPORT_EDI856 @EIN!=0 branch (a separate D1/site bug) when migrating it.
--
-- SELECT_ManifestCost;1 (the master/UI "price on date D" lookup) is also SUPERSEDED by this TVF — its
-- legacy STRICT (> / <) boundary dropped the first/last effective day (Q-boundary, resolved inclusive).
-- The rebuilt UI lookup becomes:  SELECT MO_PRICE FROM dbo.fn_ManifestCostAt(@assyCode, @prodDate);
-- so the master screen and the reports return the SAME price for a boundary day (they disagreed before).
--
-- ⚠️ Per-caller date ARG (the lookup's 2nd arg must be the PER-ROW 8-char yyyymmdd production date,
-- never a report filter param):
--   REPORT_EDI856 / REPORT_EDI810 / REPORT_INVOICESSummary : pass a.VC_PRODUCTION_DATE (8-char).
--   REPORT_MonthlyINVOICESSummary : its @PDate is a 6-char MONTH; do NOT pass @PDate — pass the
--     per-row a.VC_PRODUCTION_DATE (8-char) or the string compare breaks (lexicographic != chrono).
--
-- STATUS: this is the shared LOOKUP only. The four window-blind callers are NOT migrated yet — D6 is
-- "tool built", not "fixed in production". The caller rewrites (+ dropping IN_ASN_EIN=6440, + a
-- proc-level parity harness over the real ASN/detail JOIN+GROUP BY) are the follow-up where faithfulness
-- is decided.

-- ---------------------------------------------------------------------------------------------
-- OVERLAP DIAGNOSTIC (Q-overlap): lists any part whose windows overlap — these are the only rows where
-- this lookup's TOP 1 could pick differently than legacy EDI856's GROUP BY. Expect ZERO once the
-- ManifestCost master enforces non-overlap (D6). Run before trusting invoice/ASN totals.
--   SELECT a.VC_ASSY_PART_NUMBER_CODE, a.VC_ASSY_MANIFEST_NUMBER, b.VC_ASSY_MANIFEST_NUMBER
--   FROM INV_MANIFEST_COST_MST a JOIN INV_MANIFEST_COST_MST b
--     ON a.VC_ASSY_PART_NUMBER_CODE = b.VC_ASSY_PART_NUMBER_CODE
--    AND a.IN_MANIFEST_COST_ID < b.IN_MANIFEST_COST_ID
--    AND a.VC_START_MANIFEST <= b.VC_END_MANIFEST AND b.VC_START_MANIFEST <= a.VC_END_MANIFEST;
