/* =====================================================================
   spike-fixtures.sql  —  Order simulation spike, fixture/seed data
   ---------------------------------------------------------------------
   SPIKE-ONLY. Apply to the docker "Inventory" spike DB. Never to prod.
   Idempotent (drop/recreate + delete-by-marker re-seed).

   Two fixtures:
   (1) SIM_SpecialDate_Fixture  — STUB for the cross-DB ALC proc
       AD_GetSpecialDate (legacy-order-spec.md §8 h1, build-spec §3).
       Body of the real proc is UNVERIFIED (lives in ALC TireOrder DB),
       so the calendar walk + added-leadtime loop are driven from this
       fixture. Calendar-derived cells are therefore "fixture-backed".
       Rows reproduce build-spec §3.3 exactly (anchor 2026-06-15 Mon):
         06-17 H (holiday, skipped -> forces offset>position),
         06-18 X (non-production),
         06-23 O (overtime, 1st fOvertimes entry),
         06-25 H (holiday),
         07-03 O (overtime, 2nd entry inside a long leadtime window).

   (2) Future open-order / in-transit seed for case (e)+(c) part
       4261102Q8000 (WHEEL/M1). The REAL open-order rows for this part
       are all dated 2024-07 (VC_FRS_DATE < @FirstFRS=20260615) so they
       are filtered out by SELECT_OrderInTransitList/OpenOrderList
       (VC_FRS_DATE >= @FirstFRS). To exercise the phasing + the
       in-transit font channel we seed 8 future rows dated 2026-06-15..19,
       one of which (the 06-18 row, which buckets onto the X column at
       fill_pos 2 — the hazard-7 day) is flipped to shipping status so it
       qualifies as IN-TRANSIT (build-spec §4 case (c) seed proposal).
       All others stay open-order (empty shipping status).
       Marked VC_ADD='SPIKEFX' so the re-seed can delete just these.
   ===================================================================== */

SET NOCOUNT ON;
GO

/* ---------- (1) calendar stub ---------- */
IF OBJECT_ID('dbo.SIM_SpecialDate_Fixture','U') IS NOT NULL
    DROP TABLE dbo.SIM_SpecialDate_Fixture;
GO
CREATE TABLE dbo.SIM_SpecialDate_Fixture (
    DT_DATE        date         NOT NULL,   -- the affected day
    VC_STATUS_ABRV varchar(1)   NOT NULL,   -- 'O' overtime | 'X' non-production | 'H' holiday/skip
    VC_LINE_NAME   varchar(10)  NOT NULL DEFAULT('')  -- '' = all lines (fixture is line-agnostic)
);
GO
INSERT INTO dbo.SIM_SpecialDate_Fixture (DT_DATE, VC_STATUS_ABRV) VALUES
    ('2026-06-17','H'),   -- holiday  -> calendar consumed, NO fill column (offset>position from here)
    ('2026-06-18','X'),   -- non-production -> fill column, flag NONPRODUCTION
    ('2026-06-23','O'),   -- overtime -> fill column, flag OVERTIME, push fill idx into fOvertimes
    ('2026-06-25','H'),   -- holiday  -> skipped
    ('2026-07-03','O');   -- overtime inside a long-leadtime window -> 2nd fOvertimes entry
GO

/* ---------- (2) future open-order + in-transit seed for 4261102Q8000 ---------- */
-- clean any prior spike seed
DELETE FROM dbo.INV_OPEN_ORDER_INF WHERE VC_ADD = 'SPIKEFX';
GO
-- The FOR-INSERT trigger INSERT_RecConfStatPartsStockMstQTY bumps stock for
-- shipping-status rows by calling INTO the cross-DB "Activity" catalog, which
-- does NOT exist in the spike container -> the insert rolls back. SIM_OrderSimulation
-- is READ-ONLY (spec §1: Start does no writes) so the trigger's coupling is
-- irrelevant to the sim; disable it ONLY around the spike seed, then re-enable so
-- the real commit-path semantics (spec §6 hazard 8) are left intact for later phases.
-- SPIKE-ONLY.
DISABLE TRIGGER dbo.INSERT_RecConfStatPartsStockMstQTY ON dbo.INV_OPEN_ORDER_INF;
GO
-- 8 future open orders for the case (e) part, dated across the fixture window.
-- 06-15..06-19 (Mon..Fri of anchor week) + 06-22/06-23 next week to exercise
-- the overtime-day bucket. Status columns all empty => OPEN-ORDER (font 10),
-- EXCEPT the 06-18 row which we flip to shipping => IN-TRANSIT (font 23, case c)
-- and which lands on the X / fill_pos 2 hazard-7 day.
INSERT INTO dbo.INV_OPEN_ORDER_INF
    (VC_SUPPLIER_CODE, VC_PART_NUMBER, VC_FRS_NUMBER, VC_RENBAN_NUMBER, IN_QTY,
     VC_STATUS_SUPPLIER_SHIPPING, VC_ARRIVAL, VC_TRAILER_NUMBER, VC_STATUS_PLANT_YARD,
     VC_PLANT_PARKING, VC_STATUS_ASSEMBLER_YARD, VC_ASSEMBLER_LOCATION,
     VC_STATUS_EMPTY_TRAILER, VC_DETENTION, VC_ORDER_DATE, VC_WAREHOUSE, VC_TERMINATED,
     VC_SHIP_DATE, VC_KANBAN_NUMBER, VC_FRS_DATE, VC_LAST_UPDATE, VC_ADD)
VALUES
    ('0572B','4261102Q8000','6061501','',         500, '', '','','','','','','','','20260613','','','','M1','20260615','20260613','SPIKEFX'),
    ('0572B','4261102Q8000','6061601','',         480, '', '','','','','','','','','20260613','','','','M1','20260616','20260613','SPIKEFX'),
    ('0572B','4261102Q8000','6061701','',         460, '', '','','','','','','','','20260613','','','','M1','20260617','20260613','SPIKEFX'),
    -- the 06-18 (X-day / fill_pos 2) row flipped to SHIPPING => IN-TRANSIT (case c):
    ('0572B','4261102Q8000','6061801','',         440, 'Y','','','','','','','','','20260613','','','','M1','20260618','20260613','SPIKEFX'),
    ('0572B','4261102Q8000','6061901','',         420, '', '','','','','','','','','20260613','','','','M1','20260619','20260613','SPIKEFX'),
    ('0572B','4261102Q8000','6062201','',         400, '', '','','','','','','','','20260613','','','','M1','20260622','20260613','SPIKEFX'),
    -- 06-23 lands on the OVERTIME day (fill_pos 5):
    ('0572B','4261102Q8000','6062301','',         380, '', '','','','','','','','','20260613','','','','M1','20260623','20260613','SPIKEFX'),
    ('0572B','4261102Q8000','6062401','',         360, '', '','','','','','','','','20260613','','','','M1','20260624','20260613','SPIKEFX');
GO

ENABLE TRIGGER dbo.INSERT_RecConfStatPartsStockMstQTY ON dbo.INV_OPEN_ORDER_INF;
GO

PRINT 'spike-fixtures.sql applied: SIM_SpecialDate_Fixture + 8 SPIKEFX open orders (1 in-transit).';
GO
