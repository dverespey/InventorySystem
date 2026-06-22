-- daily_shipping_header_assy.sql  (Named Query: report_render/daily_shipping_header_assy)
-- Header band for the Daily Shipping Assy report (R3): the Start Seq / End Seq / Vehicle Count cells
-- ([2,2] = 'Start Seq:..../End Seq:....', [2,3] = 'Vehicle Count:<n>'). params: @PDate varchar(8).
--
-- The legacy handler reads Start/End/Vehicle Count from the FIRST detail row of REPORT_DailyShippingAssy
-- (MainMenu.pas:3174-3175 — fieldbyname('Start'/'End'/'Vehicle Count') before the loop), i.e. the values
-- of whichever ASN header sorted first. The detail proc carries these per (part × header) so they vary
-- across rows; reading "the first row" is the legacy behavior but is grain-fragile when a date has >1 ASN.
-- This header NQ instead AGGREGATES the date's ASN headers (MIN start / MAX end / SUM qty) into ONE row —
-- the same de-fan fix daily_shipping_header.sql applies for the T/W report (daily-shipping-report-spec.md
-- §5/§6: header values shown ONCE, aggregated for the date, not joined into detail). On the spike a date
-- typically has 1 ASN header so MIN/MAX/SUM == the single header (faithful); when a date has >1 ASN this
-- is the de-fanned correct header (and avoids the "first row only" fragility).
--
-- StartEnd is pre-concatenated server-side so the header_cell ("header","StartEnd") is a single bind.
-- READ-ONLY pure SELECT. IG83-TODO: promote to a Named Query; swap runPrepQuery -> runNamedQuery.

SELECT  MIN(VC_START_SEQ_NUMBER) 'StartSeq',
        MAX(VC_END_SEQ_NUMBER)   'EndSeq',
        SUM(IN_QTY)              'Vehicle Count',
        'Start Seq:' + MIN(VC_START_SEQ_NUMBER) + '/End Seq:' + MAX(VC_END_SEQ_NUMBER) 'StartEnd'
FROM    INV_ASN_MST
WHERE   VC_PRODUCTION_DATE = ?            -- @PDate
