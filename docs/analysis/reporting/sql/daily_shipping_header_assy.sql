-- daily_shipping_header_assy.sql  (Named Query: report_render/daily_shipping_header_assy)
-- Header band for the Daily Shipping Assy report (R3): the Start Seq / End Seq / Vehicle Count cells
-- ([2,2] = 'Start Seq:..../End Seq:....', [2,3] = 'Vehicle Count:<n>'). params: @PDate varchar(8).
--
-- FAITHFUL rebuild of the LEGACY header read. The Delphi handler prints the header band from the
-- FIRST DETAIL ROW of REPORT_DailyShippingAssy, BEFORE the z:=4 detail loop:
--     mysheet.cells[2,2].value:='Start Seq:'+fieldbyname('Start').AsString+'/End Seq:'+fieldbyname('End').AsString;
--     mysheet.cells[2,3].value:='Vehicle Count:'+fieldbyname('Vehicle Count').AsString;
--   (MainMenu.pas:3173-3174, in TMainMenu_Form.DailyASNReportClick — DataSet still on row 1).
-- The detail proc carries Start/End/Vehicle Count per (part x ASN header) and is ordered
-- `ORDER BY d.VC_ASSY_PART_NUMBER` (live REPORT_DailyShippingAssy body, CreateInventory.sql:3576),
-- so "row 1" = the values of whichever ASN header owns the alphabetically-first VC_ASSY_PART_NUMBER.
-- The legacy therefore prints ONE header's range, NOT an aggregate.
--
-- DO NOT use MIN/MAX/SUM here (the prior rebuild did, and it diverged): on a multi-header day a
-- throwaway `-1/-1` sentinel header (qty 1) co-exists with the real header. The string MIN(start) makes
-- the varchar sentinel `-1` win (prints "Start Seq: -1" -- nonsensical) and SUM(IN_QTY) inflates the
-- vehicle count (e.g. live 20250121: MIN/MAX/SUM -> -1 / 0914 / 831, but the legacy proc's row 1 prints
-- 0095 / 0914 / 820). Reachable on 205 production dates (>1 distinct start-seq); the regression is the
-- `-1` sentinel + inflated count.
--
-- This NQ reproduces the legacy "first detail row" deterministically with TOP 1 under the proc's
-- ORDER BY d.VC_ASSY_PART_NUMBER, plus tiebreakers that pin "row 1" so it is reproducible AND can never
-- regress to the sentinel:
--   1) d.VC_ASSY_PART_NUMBER ASC                          -- the proc's own ORDER BY -> "row 1" = first part
--   2) CASE WHEN VC_START_SEQ_NUMBER='-1' THEN 1 ELSE 0   -- the `-1` sentinel header can never be row 1
--   3) VC_START_SEQ_NUMBER, VC_END_SEQ_NUMBER, IN_QTY DESC -- fully deterministic order
-- On the 3 live dates where the first-sorted part is tied across a sentinel + a real header
-- (20180726/20190413/20211201) tiebreaker (2) selects the REAL header (536/444/788), never the `-1`/1
-- placeholder -- matching what the legacy meaningfully printed. On single-header dates and on 20250121
-- the TOP-1 row equals the live proc's row 1 exactly (verified against EXEC REPORT_DailyShippingAssy).
--
-- StartEnd is pre-concatenated server-side so the header_cell ("header","StartEnd") is a single bind,
-- matching the legacy 'Start Seq:'+Start+'/End Seq:'+End string.
-- READ-ONLY pure SELECT. IG83-TODO: promote to a Named Query; swap runPrepQuery -> runNamedQuery.

SELECT TOP 1
        s.VC_START_SEQ_NUMBER 'StartSeq',
        s.VC_END_SEQ_NUMBER   'EndSeq',
        s.IN_QTY              'Vehicle Count',
        'Start Seq:' + s.VC_START_SEQ_NUMBER + '/End Seq:' + s.VC_END_SEQ_NUMBER 'StartEnd'
FROM    INV_ASN_MST        s
JOIN    INV_ASN_DETAIL_MST d ON s.IN_ASN_ID = d.IN_ASN_ID
WHERE   s.VC_PRODUCTION_DATE = ?            -- @PDate
GROUP BY d.VC_ASSY_PART_NUMBER, s.VC_START_SEQ_NUMBER, s.VC_END_SEQ_NUMBER, s.IN_QTY, d.IN_QTY
ORDER BY d.VC_ASSY_PART_NUMBER ASC,
         CASE WHEN s.VC_START_SEQ_NUMBER = '-1' THEN 1 ELSE 0 END ASC,
         s.VC_START_SEQ_NUMBER ASC,
         s.VC_END_SEQ_NUMBER ASC,
         s.IN_QTY DESC
