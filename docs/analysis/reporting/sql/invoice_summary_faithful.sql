-- invoice_summary_faithful.sql  (Named Query: report_render/invoice_summary_faithful) — FAITHFUL (legacy)
-- The legacy window-BLIND REPORT_INVOICESSummary VERBATIM (live OBJECT_DEFINITION; CreateInventory.sql:3785).
-- Sits behind the faithful↔corrected seam (QUERY_VARIANT='faithful') so the parity test can prove the
-- corrected query diverges ONLY by the D6 window fix. NOT the default ship variant (default = corrected,
-- invoice_summary.sql). Reproduces the over-bill: JOIN INV_MANIFEST_COST_MST on the assy code ALONE (no
-- VC_START/END_MANIFEST window vs production date) -> a part with >1 price window multiplies its lines.
-- params: @PDate varchar(13) (caller fills yyyymmdd only). READ-ONLY pure SELECT.

SELECT  d.VC_MANIFEST_NUMBER   'Manifest',
        d.VC_ASSY_PART_NUMBER  'PartNumber',
        m.MO_PRICE             'UnitPrice',
        d.IN_QTY               'ShipQty',
        a.VC_PRODUCTION_DATE   'PickUpDate'
FROM    INV_ASN_MST           a
JOIN    INV_ASN_DETAIL_MST    d ON a.IN_ASN_ID = d.IN_ASN_ID
JOIN    INV_MANIFEST_COST_MST m ON d.VC_ASSY_PART_NUMBER = m.VC_ASSY_PART_NUMBER_CODE   -- (!) window-blind
JOIN    INV_INV_MST           i ON i.IN_INV_ID = d.IN_INV_ID
WHERE   a.VC_PRODUCTION_DATE = ?            -- @PDate
ORDER BY d.VC_MANIFEST_NUMBER
