-- spike-edi810-feed.sql — the side-effect-free SELECTs that feed the outbound 810 invoice builder.
--
-- Reproduces REPORT_EDI810's TWO read branches (report-edi810-data-analysis.md; the D6 migration
-- spike-report-procs-d6.sql:60-92) WITHOUT the proc's hazards:
--   * NO self-flip. The proc's @EIN<>0 branch runs `UPDATE INV_INV_MST SET VC_INV_STATUS='S' WHERE
--     IN_INV_EIN=@EIN` (:88) — reading the recreate feed MUTATES state. These SELECTs are PURE; the
--     per-invoice status flip is done SEPARATELY by the drivers (create inserts status 'S' via
--     INSERT_INVInfo; unsend reverts in-place to 'C'). NEVER UPDATE_INVRecreate (blanket 'C'->'S').
--   * D6 WINDOW-AWARE pricing. Both branches use `CROSS APPLY dbo.fn_ManifestCostAt(part, productionDate)`
--     (inclusive <= / >=, TOP-1 newest-start) instead of the legacy window-blind `JOIN
--     INV_MANIFEST_COST_MST` on part code ALONE. PROVEN over-bill on real data: EIN 5692 legacy
--     $61,783.75 (2 lines, wrong window) vs D6 $58,093.75 (1 line) — an accepted billing-correctness
--     divergence (report-edi810-data-analysis.md). The price IS emitted on the 810 (IT104 + the TDS
--     total), so a wrong/duplicate price mis-bills — unlike the 856, which carries no price.
--   * NO GROUP BY (unlike the 856 feed). One row per (header x surviving detail). CROSS APPLY TOP 1
--     prevents fan-out on a clean cost master.
--
-- TWO BRANCHES (the proc's @EIN signature):
--   CREATE (@EIN=0): the daily-invoicing read. Unbilled detail only — the unbilled flag is on the DETAIL
--     (d.IN_INV_ID IS NULL) with the header VC_ASN_STATUS='A'. 6 cols
--     (Manifest/Part/UnitPrice/Qty/PickUp/ASNid). The driver allocates a NEW EIN, inserts the header,
--     links this detail. (Snapshot edge: ALL detail is already billed -> this returns 0 rows on the live
--     snapshot; test_edi810_e2e.py synthesizes unbilled lines.)
--   RECREATE (@EIN<>0): the existing-invoice read. + JOIN INV_INV_MST iim ON d.IN_INV_ID=iim.IN_INV_ID,
--     WHERE iim.IN_INV_EIN=@EIN. 7 cols (+InvEIN). The driver REUSES the EIN (no re-alloc), re-emits the
--     same wire. The proc's @EIN<>0 self-flip is OMITTED here.
--
-- PROJECTION: the builder (edi810.build_810) consumes Manifest / PartNumber / ShipQty / UnitPrice (money/
-- scale-4) / PickUpDate. The create branch also returns d.IN_ASN_ID (ASNid) so the driver can link the
-- detail per ASN (UPDATE_INVItems @ASNID). The recreate branch returns iim.IN_INV_EIN (InvEIN) for
-- traceability. ORDER BY VC_MANIFEST_NUMBER — the IT1 manifest-break (REF*MK / DTM*050) needs the rows
-- grouped by manifest; the legacy proc ordered the same way. (No PartNumber tiebreak in the legacy proc;
-- kept faithful — the manifest order is what drives the breaks.)
--
-- CANONICAL COPY + DRIFT GUARD: the drivers (edi810/code.py:_CREATE_FEED_SQL / _RECREATE_FEED_SQL) inline
-- the SELECT bodies between the markers below. Each is BYTE-IDENTICAL to its inline modulo the single bind:
-- this file uses @EIN, the inline uses runPrepQuery's `?` (create takes NO bind; recreate binds the EIN).
-- test_edi810_e2e.py asserts _CREATE_FEED_SQL/_RECREATE_FEED_SQL == the marked blocks (whitespace-
-- collapsed, '@EIN' -> '?') so they cannot drift. If you edit one, edit both; the test fails if you forget.
-- Run standalone:  sqlcmd -v EIN=5692  (recreate)  /  -v EIN=0  (create).

-- :setvar EIN 0      -- 0 = the CREATE/unbilled branch; non-zero = the RECREATE branch for that invoice EIN

DECLARE @EIN int = $(EIN);

-- >>> BEGIN CREATE FEED SQL (canonical; byte-identical to code.py:_CREATE_FEED_SQL; takes NO bind) >>>
SELECT d.VC_MANIFEST_NUMBER AS Manifest,
       d.VC_ASSY_PART_NUMBER AS PartNumber,
       m.MO_PRICE AS UnitPrice,
       d.IN_QTY AS ShipQty,
       a.VC_PRODUCTION_DATE AS PickUpDate,
       d.IN_ASN_ID AS ASNid
FROM INV_ASN_MST a
JOIN INV_ASN_DETAIL_MST d
ON a.IN_ASN_ID = d.IN_ASN_ID AND a.VC_ASN_STATUS = 'A' AND d.IN_INV_ID IS NULL
CROSS APPLY dbo.fn_ManifestCostAt(d.VC_ASSY_PART_NUMBER, a.VC_PRODUCTION_DATE) m
ORDER BY d.VC_MANIFEST_NUMBER;
-- <<< END CREATE FEED SQL <<<

-- >>> BEGIN RECREATE FEED SQL (canonical; byte-identical to code.py:_RECREATE_FEED_SQL modulo @EIN<->?) >>>
SELECT d.VC_MANIFEST_NUMBER AS Manifest,
       d.VC_ASSY_PART_NUMBER AS PartNumber,
       m.MO_PRICE AS UnitPrice,
       d.IN_QTY AS ShipQty,
       a.VC_PRODUCTION_DATE AS PickUpDate,
       d.IN_ASN_ID AS ASNid,
       iim.IN_INV_EIN AS InvEIN
FROM INV_ASN_MST a
JOIN INV_ASN_DETAIL_MST d ON a.IN_ASN_ID = d.IN_ASN_ID
JOIN INV_INV_MST iim ON d.IN_INV_ID = iim.IN_INV_ID
CROSS APPLY dbo.fn_ManifestCostAt(d.VC_ASSY_PART_NUMBER, a.VC_PRODUCTION_DATE) m
WHERE iim.IN_INV_EIN = @EIN
ORDER BY d.VC_MANIFEST_NUMBER;
-- <<< END RECREATE FEED SQL <<<
