-- lot_location_tires.sql  (Named Query: report_render/lot_location_tires)
-- Wraps dbo.REPORT_PLANTLotLocation (the TIRES half of the LIVE "Do Lot Location" report, R12).
-- CONFIRMED LIVE (PLANT proc, not the D9-deprecated NUMMI twin — see lot_location_wheels.sql).
-- params: NONE. FAITHFUL rebuild. READ-ONLY pure SELECT.
--
-- Body VERBATIM (live OBJECT_DEFINITION). The tires proc emits vc_size_code + line_name; the report's tire
-- section is grouped by size_code into a size-code header row ("<NO SIZE>" when blank), MainMenu.pas:836-861.
-- Same 0-row-on-spike caveat as the wheels query (vc_status_PLANT_yard='' everywhere) -> SEED for parity.
-- IG83-TODO: promote to a Named Query; swap runPrepQuery -> runNamedQuery.

SELECT  s.vc_size_code,
        p.vc_part_number,
        o.vc_PLANT_parking,
        o.vc_ASSEMBLER_location,
        o.vc_status_PLANT_yard,
        o.vc_renban_number,
        o.vc_trailer_number,
        x.VC_PART_TYPE,
        p.vc_LINE_NAME
FROM    inv_open_order_inf o
JOIN    inv_parts_stock_mst p ON o.vc_part_number = p.vc_part_number
JOIN    INV_PART_TYPE_MST    x ON p.IN_PART_TYPE_ID = x.IN_PART_TYPE_ID
JOIN    INV_SIZE_MST         s ON p.IN_SIZE_ID = s.IN_SIZE_ID
WHERE   o.vc_status_PLANT_yard <> ''
  AND   x.VC_PART_TYPE = 'TIRE'
  AND  (o.vc_status_empty_trailer = '' AND o.vc_terminated = '')
ORDER BY x.VC_PART_TYPE DESC, p.vc_LINE_NAME, s.vc_size_code,
         o.vc_status_PLANT_yard, o.vc_renban_number
