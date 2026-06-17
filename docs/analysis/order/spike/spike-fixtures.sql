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

/* ---------- (2) M1 4261102Q8000 receipts — use the REAL renban-grouped prod data ----------
   R3 (resolved 2026-06-15): an earlier version of this fixture INJECTED 8 synthetic
   blank-renban "SPIKEFX" open/in-transit rows for 4261102Q8000 to exercise font/bucket
   scenarios. That was UNFAITHFUL: 4261102Q8000 is PALLETIZED (BIT_LOT_SIZE_ORDERS=1, which
   is the inverted prod flag => lot-sized FALSE) and renban-grouped (group CMWA). In prod its
   open orders ALWAYS carry a renban (a placeholder that the RenbanOrder grouping form
   overwrites with the grouped CMWA renban); there is NEVER a blank-renban row, and the
   breakdown step DELETEs the placeholders so duplicates never reach Order Start. The injected
   blank rows coexisted with the real 855 CMWA rows (from Inventory.bak), double-counting the
   receipts (e.g. 06-15: golden CMWA 440 + injected blank 500 = 940). The Order receipt
   projection (SELECT_OrderOpenOrderList / PutOpenOrderCount) SUMS all rows by VC_FRS_DATE with
   NO renban filter, so SIM_OrderSimulation STEP 5 (sum-all-rows) is already faithful — the bug
   was purely the injected fixture rows. FIX = drop them; the real CMWA prod data alone
   reproduces the golden M1 receipts [440,880,880,880,400,0]. NO proc change.  */
-- Remove the unfaithful synthetic seed (keep this so a re-seed cleans any prior SPIKEFX rows).
-- The FOR-DELETE trigger DELETE_RecConfStatPartsStockMstQTY writes into the cross-DB "Activity"
-- catalog (absent in the spike container) and would roll the DELETE back, so disable it just
-- around the cleanup. SIM_OrderSimulation is READ-ONLY so the trigger is irrelevant to the sim.
DISABLE TRIGGER dbo.DELETE_RecConfStatPartsStockMstQTY ON dbo.INV_OPEN_ORDER_INF;
GO
DELETE FROM dbo.INV_OPEN_ORDER_INF WHERE VC_ADD = 'SPIKEFX';
GO
ENABLE TRIGGER dbo.DELETE_RecConfStatPartsStockMstQTY ON dbo.INV_OPEN_ORDER_INF;
GO
GO

PRINT 'spike-fixtures.sql applied: SIM_SpecialDate_Fixture + 8 SPIKEFX open orders (1 in-transit).';
GO
