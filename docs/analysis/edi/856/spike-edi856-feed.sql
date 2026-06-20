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
--   B  INNER forecast join for Kanban, capped to ONE row via CROSS APPLY (SELECT TOP 1 ...). CROSS APPLY
--      (not OUTER APPLY) keeps the INNER drop: a part with NO forecast row produces no APPLY row, so the
--      detail line drops — exactly like the legacy INNER JOIN INV_FORECAST_DETAIL_INF. The TOP 1 prevents
--      the fan-out the legacy INNER join would suffer if a part had >1 forecast row (no uniqueness on the
--      forecast part code). So: same DROP as legacy, but no DUPLICATION (the latent legacy fan-out is the
--      one thing we tighten — it could only ever have inflated the 856, never matched it).
--   C  GROUP-BY collapse (poor-man's DISTINCT), NO sum. Distinct detail rows that share all projected
--      columns merge into one — matching legacy. (A "correct" ASN might SUM IN_QTY over merged lines, but
--      byte-parity locks the legacy DROP/collapse.)
--   inclusive cost window  (<= / >=) — the LIVE proc uses inclusive compares (supersedes asn-invoice.md
--      §4.1's strict-compare note; confirmed against the live REPORT_EDI856 body). String compares on
--      varchar(8) yyyymmdd are correct because that format sorts lexicographically.
--
-- PROJECTION: the builder (edi856.build_856) consumes only Manifest / PartNumber / ShipQty / Kanban, plus
-- the header's PickUpDate (for the dates + filename). The legacy 9-col feed also returned UnitPrice (NOT
-- emitted — the 856 carries no price), SiteEIN, StartSeq, LineName (unused by the segment build). We
-- project the 5 columns the build actually needs; the driver reads PickUpDate from the header row too.
-- (Keeping all 9 would be harmless but adds GROUP BY columns that change the collapse cardinality, so we
-- project exactly what the wire needs.)
--
-- ORDER BY Manifest, PartNumber — the S->O->I HL grouping needs the rows grouped by manifest. The legacy
-- proc's GROUP BY emitted in an engine-defined order; the rebuild pins a DETERMINISTIC order so the
-- Order-HL breaks (and thus the HL ids / SE01 / CTT01 counts) are reproducible. PartNumber is the stable
-- tiebreak within a manifest.
--
-- USAGE: the driver (edi856.code.py:_FEED_SQL) inlines this SELECT body verbatim with a ? for @ASNID.
-- This .sql file is the reviewable canonical copy; if you edit one, edit both (a code comment flags it).
-- Run standalone:  sqlcmd -v ASNID=4721 EIN=0  (or set the vars below for an ad-hoc preview).

-- :setvar ASNID 4721      -- the ASN to read (the driver binds this as ?)
-- :setvar EIN 0           -- traceability only (NOT a filter; the feed is keyed by @ASNID)

DECLARE @ASNID int = $(ASNID);
DECLARE @EIN   int = $(EIN);     -- carried for parity with the legacy signature; the feed does not filter on it

SELECT  d.VC_MANIFEST_NUMBER    AS Manifest,    -- -> PRF (Order HL)            d.VC_MANIFEST_NUMBER
        d.VC_ASSY_PART_NUMBER   AS PartNumber,  -- -> LIN BP                    d.VC_ASSY_PART_NUMBER
        d.IN_QTY                AS ShipQty,     -- -> SN1                       d.IN_QTY (int)
        a.VC_PRODUCTION_DATE    AS PickUpDate,  -- -> BSN/DTM/filename          a.VC_PRODUCTION_DATE yyyymmdd
        f.VC_ASSY_KANBAN_NUMBER AS Kanban       -- -> LIN RC                    f.VC_ASSY_KANBAN_NUMBER
FROM            INV_ASN_MST          a
JOIN            INV_ASN_DETAIL_MST   d  ON a.IN_ASN_ID = d.IN_ASN_ID
JOIN            INV_MANIFEST_COST_MST m  ON d.VC_ASSY_PART_NUMBER = m.VC_ASSY_PART_NUMBER_CODE   -- A: INNER cost
CROSS APPLY (   SELECT TOP 1 f1.VC_ASSY_KANBAN_NUMBER                                            -- B: INNER + TOP 1
                FROM  INV_FORECAST_DETAIL_INF f1
                WHERE f1.VC_ASSY_PART_NUMBER_CODE = d.VC_ASSY_PART_NUMBER ) f
WHERE   a.IN_ASN_ID = @ASNID                                                                     -- @ASNID, NOT 6440
AND     m.VC_START_MANIFEST <= a.VC_PRODUCTION_DATE                                              -- inclusive window
AND     m.VC_END_MANIFEST   >= a.VC_PRODUCTION_DATE
GROUP BY        d.VC_MANIFEST_NUMBER,                                                            -- C: collapse, no sum
                d.VC_ASSY_PART_NUMBER,
                d.IN_QTY,
                a.VC_PRODUCTION_DATE,
                f.VC_ASSY_KANBAN_NUMBER
ORDER BY        d.VC_MANIFEST_NUMBER, d.VC_ASSY_PART_NUMBER;                                      -- deterministic HL order
