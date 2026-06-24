-- spike-edi856-feed.sql — the side-effect-free SELECT that feeds the outbound 856 builder.
--
-- Reproduces REPORT_EDI856's read (report-edi856-data-analysis.md) WITHOUT the proc's hazards:
--   * NO self-flip (the proc's @EIN<>0 branch runs UPDATE INV_ASN_MST SET VC_ASN_STATUS='S' — reading
--     the feed mutates state; this SELECT is pure. The C->S flip is done SEPARATELY, per-ASN, at send,
--     by the driver edi856.send_856).
--   * Parameterized by @ASNID, NOT the literal WHERE IN_ASN_EIN=6440 (the proc's @EIN<>0 branch pins one
--     baked-in EIN and ignores the EIN passed). The rebuild sends ONE specific ASN by its id.
--   * @EIN passed for traceability only (the feed itself does not filter on it — the EIN is allocated at
--     send and stamped on the header; the feed is keyed by @ASNID).
--
-- LOCKED DECISIONS (byte-faithful — match the TEMA-accepted legacy output):
--   A  KEEP the cost INNER join (filters to cost-covered lines, like legacy REPORT_EDI856). A line whose
--      part has no covering manifest-cost window silently vanishes from the 856 — that is the legacy
--      behaviour, reproduced.
--   B  INNER forecast join for Kanban — the SAME `JOIN INV_FORECAST_DETAIL_INF f ON d.VC_ASSY_PART_NUMBER
--      = f.VC_ASSY_PART_NUMBER_CODE` the legacy proc uses. A part with >1 forecast row FANS OUT to
--      multiple rows; the GROUP BY then keeps each DISTINCT kanban (so 2 distinct-kanban forecast rows ->
--      2 LIN lines, exactly like legacy). A part with NO forecast row drops the detail line (INNER drop).
--      NOTE: this REPLACES the earlier `CROSS APPLY (SELECT TOP 1 ...)` — that TOP 1 had NO ORDER BY over
--      a HEAP (nondeterministic) AND dropped legacy lines (kept ONE kanban where legacy keeps every
--      distinct one). The legacy fan-out is NOT pure inflation: GROUP BY collapses same-kanban dupes and
--      keeps distinct kanbans, which is what the wire requires.
--   C  GROUP-BY collapse (poor-man's DISTINCT), NO sum. Distinct detail rows that share all GROUP-BY
--      columns merge into one — matching legacy. The GROUP BY KEEPS m.MO_PRICE (as legacy does): a part
--      with 2 price-distinct overlapping cost windows emits 2 detail rows. MO_PRICE is NOT in the SELECT
--      (the 856 carries no price) but it IS a GROUP-BY key — dropping it collapsed legacy's 2 rows to 1.
--   inclusive cost window  (<= / >=) — the LIVE proc uses inclusive compares (supersedes asn-invoice.md
--      §4.1's strict-compare note; confirmed against the live REPORT_EDI856 body). String compares on
--      varchar(8) yyyymmdd are correct because that format sorts lexicographically.
--
-- PROJECTION: the builder (edi856.build_856) consumes only Manifest / PartNumber / ShipQty / Kanban, plus
-- the header's PickUpDate (for the dates + filename). The legacy 9-col feed also returned UnitPrice (NOT
-- emitted — the 856 carries no price), SiteEIN, StartSeq, LineName (unused by the segment build). We
-- project the 5 columns the build needs but KEEP MO_PRICE in the GROUP BY (it changes the collapse
-- cardinality even though it never reaches the wire). The header-constant legacy GROUP-BY columns
-- (a.IN_ASN_EIN, VC_START_SEQ_NUMBER, a.VC_LINE_NAME) are per-ASN constant under our `WHERE IN_ASN_ID=?`
-- scope, so they cannot change the row count and are omitted; MO_PRICE comes from the cost join and CAN.
--
-- ORDER BY Manifest, PartNumber — the S->O->I HL grouping needs the rows grouped by manifest. The legacy
-- proc's GROUP BY emitted in an engine-defined order; the rebuild pins a DETERMINISTIC order so the
-- Order-HL breaks (and thus the HL ids / SE01 / CTT01 counts) are reproducible. PartNumber is the stable
-- tiebreak within a manifest.
--
-- CANONICAL COPY + DRIFT GUARD: the driver (edi856.code.py:_FEED_SQL) inlines the SELECT body between the
-- two markers below. The two are BYTE-IDENTICAL except the single bind: this file declares @ASNID (so it
-- runs standalone) and the inline uses runPrepQuery's `?`. test_edi856_e2e.py asserts
--   _FEED_SQL == (the marked block, whitespace-collapsed, with '@ASNID' -> '?')
-- so they cannot drift. If you edit one, edit both; the test fails if you forget.
-- Run standalone:  sqlcmd -v ASNID=4721 EIN=0  (or set the vars below for an ad-hoc preview).

-- :setvar ASNID 4721      -- the ASN to read (the driver binds this as ?)
-- :setvar EIN 0           -- traceability only (NOT a filter; the feed is keyed by @ASNID)

DECLARE @ASNID int = $(ASNID);
DECLARE @EIN   int = $(EIN);     -- carried for parity with the legacy signature; the feed does not filter on it

-- >>> BEGIN FEED SQL (canonical; byte-identical to code.py:_FEED_SQL modulo @ASNID<->?) >>>
SELECT d.VC_MANIFEST_NUMBER AS Manifest,
       d.VC_ASSY_PART_NUMBER AS PartNumber,
       d.IN_QTY AS ShipQty,
       a.VC_PRODUCTION_DATE AS PickUpDate,
       f.VC_ASSY_KANBAN_NUMBER AS Kanban
FROM INV_ASN_MST a
JOIN INV_ASN_DETAIL_MST d ON a.IN_ASN_ID = d.IN_ASN_ID
JOIN INV_MANIFEST_COST_MST m ON d.VC_ASSY_PART_NUMBER = m.VC_ASSY_PART_NUMBER_CODE
JOIN INV_FORECAST_DETAIL_INF f ON d.VC_ASSY_PART_NUMBER = f.VC_ASSY_PART_NUMBER_CODE
WHERE a.IN_ASN_ID = @ASNID
AND m.VC_START_MANIFEST <= a.VC_PRODUCTION_DATE
AND m.VC_END_MANIFEST   >= a.VC_PRODUCTION_DATE
GROUP BY d.VC_MANIFEST_NUMBER, d.VC_ASSY_PART_NUMBER, m.MO_PRICE, d.IN_QTY,
       a.VC_PRODUCTION_DATE, f.VC_ASSY_KANBAN_NUMBER
ORDER BY d.VC_MANIFEST_NUMBER, d.VC_ASSY_PART_NUMBER;
-- <<< END FEED SQL <<<
