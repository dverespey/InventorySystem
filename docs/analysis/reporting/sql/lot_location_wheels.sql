-- lot_location_wheels.sql  (Named Query: report_render/lot_location_wheels)
-- Wraps dbo.REPORT_PLANTLotLocationW (the WHEELS half of the LIVE "Do Lot Location" report, R12).
-- CONFIRMED LIVE, NOT the deprecated NUMMI variant: MainMenu.pas:LotLocationClick calls
-- REPORT_PLANTLotLocationW then REPORT_PLANTLotLocation (the PLANT procs); REPORT_NUMMILotLocation[W]
-- is the D9-deprecated twin and is NOT referenced by the live handler. (Confirmed by reading the handler
-- body + grepping the "Do Lot Location" LogActLog string -> MainMenu.pas:882.)
-- params: NONE. FAITHFUL rebuild. READ-ONLY pure SELECT (the W proc GROUPs but does not mutate).
--
-- Body VERBATIM (live OBJECT_DEFINITION). The wheels proc emits NO size/part-number col (commented out
-- in the proc); the report's wheel section is grouped by line name into "Car Wheels" (Passenger) /
-- "Truck Wheels" group-break header rows (MainMenu.pas:792-803). NOTE: returns 0 rows on the current spike
-- (every INV_OPEN_ORDER_INF.vc_status_PLANT_yard = '' -> the WHERE drops all) — parity must SEED fixture
-- rows to be non-vacuous (fixture-fidelity lesson).
-- IG83-TODO: promote to a Named Query; swap runPrepQuery -> runNamedQuery.

SELECT  o.vc_PLANT_parking,
        o.vc_ASSEMBLER_location,
        o.vc_status_PLANT_yard,
        o.vc_renban_number,
        o.vc_trailer_number,
        p.vc_LINE_NAME
FROM    inv_open_order_inf o
JOIN    inv_parts_stock_mst p ON o.vc_part_number = p.vc_part_number
JOIN    INV_PART_TYPE_MST    x ON p.IN_PART_TYPE_ID = x.IN_PART_TYPE_ID
WHERE   o.vc_status_PLANT_yard <> ''
  AND   x.VC_PART_TYPE = 'WHEEL'
  AND  (o.vc_status_empty_trailer = '' AND o.vc_terminated = '')
GROUP BY o.vc_renban_number, o.vc_trailer_number, o.vc_status_PLANT_yard,
         o.vc_PLANT_parking, o.vc_ASSEMBLER_location, p.vc_LINE_NAME
ORDER BY p.vc_LINE_NAME, o.vc_status_PLANT_yard, o.vc_renban_number
