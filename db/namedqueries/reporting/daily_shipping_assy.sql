-- daily_shipping_assy.sql  (Named Query: report_render/daily_shipping_assy)
-- Wraps dbo.REPORT_DailyShippingAssy (M3 report R3, the 285-runs/yr live shipping report).
-- params: @PDate varchar(8)  (yyyymmdd)
--
-- FAITHFUL rebuild: this query is the proc body VERBATIM (CreateInventory.sql:3576 / live
-- OBJECT_DEFINITION). The Assy proc joins ASN_MST -> ASN_DETAIL_MST on the real key IN_ASN_ID, so it has
-- NO date-only fan-out (contrast the T/W Daily Shipping procs) and NO orphan-drop master inner join.
-- It is the SOUND shipping path (m3-report-proc-survey.md row REPORT_DailyShippingAssy = "clean").
--
-- READ-ONLY: no INSERT/UPDATE/DELETE (verified — pure SELECT). Not an EDI mutate-on-read proc.
--
-- PORT NOTE (m3-report-inventory.md / daily-shipping-report-spec.md §3): the Delphi caller
-- (MainMenu.pas:3158) passes an extra @Line param the proc does NOT declare. ADO binds positionally so
-- it is benign today, but a NAME-binding Named Query would fail on an undeclared param -> @Line is DROPPED
-- here. Only @PDate is bound. (daily-shipping-report-spec.md §3 self-flag.)
--
-- The s.IN_QTY in GROUP BY (and the Start/End/Vehicle Count header cols) means a part is split across
-- distinct (seq-range, line-qty) headers rather than fully summed — this is FAITHFUL to the proc; the
-- Excel render only consumes Part Number + PQty for the detail (Start/End/Vehicle Count -> header band).
-- IG83-TODO: promote daily_shipping_assy.sql to a real Named Query; swap runPrepQuery -> runNamedQuery.

SELECT  s.VC_START_SEQ_NUMBER  'Start',
        s.VC_END_SEQ_NUMBER    'End',
        s.IN_QTY               'Vehicle Count',
        d.VC_ASSY_PART_NUMBER  'Part Number',
        SUM(d.IN_QTY)          'PQty'
FROM    INV_ASN_MST        s
JOIN    INV_ASN_DETAIL_MST d ON s.IN_ASN_ID = d.IN_ASN_ID
WHERE   s.VC_PRODUCTION_DATE = ?            -- @PDate  (Named Query: the marker becomes :PDate)
GROUP BY d.VC_ASSY_PART_NUMBER, s.VC_START_SEQ_NUMBER, s.VC_END_SEQ_NUMBER, s.IN_QTY, d.IN_QTY
ORDER BY d.VC_ASSY_PART_NUMBER
