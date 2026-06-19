-- spike-partsstockinfo-drop-qty-clause.sql
-- CUTOVER artifact — Carry 10 (David ACCEPTED 2026-06-19).
--
-- DECISION (verbatim-faithful): "Carry 10 — ACCEPTED. Drop the IN_QTY=@QTY clause from
-- UPDATE_PartsStockInfo (~line 5747 in the live schema; on-hand is never editable on the rebuilt Parts
-- Stock master). Also: UPDATE_PartsStockInfoCount is DEAD — retire it."
--
-- WHY. After the writer-flip (Phase C), the stock LEDGER + the producer seams own INV_PARTS_STOCK_MST.IN_QTY
-- (it is a materialized SUM(ledger)). Any path still doing a direct, non-ledger IN_QTY write desyncs the
-- materialized balance from SUM(ledger) — a lost update with no ledger row. The 4th/5th-writer hunt
-- (cutover-runbook §5b) found exactly two non-trigger procs that write INV_PARTS_STOCK_MST.IN_QTY directly:
--   1. UPDATE_PartsStockInfo — absolute `IN_QTY = @QTY` (live caller DataModule.pas:1482; the rebuilt
--      Parts Stock master Save calls the SAME proc with all 30 params incl. @QTY). It is a MASTER-DATA
--      edit, not a stock move — the qty leg is a latent clobber. -> drop the `IN_QTY=@QTY` clause.
--   2. UPDATE_PartsStockInfoCount — additive `IN_QTY = IN_QTY-@QTY`. DEAD: whole-repo grep finds ZERO
--      callers (live or legacy Delphi) and no rebuilt-lib reference. -> retire it.
--
-- @QTY PARAM IS KEPT IN THE SIGNATURE (now unused). The rebuilt Perspective master Save passes all 30
-- sys.parameters POSITIONALLY (ignition-spike-log.md:102-114); removing @QTY would shift every following
-- arg and break the call. Dropping only the body assignment means the param is accepted-and-ignored — a
-- single-point DB edit, NO Ignition-side change (Named-Query practice). The rebuilt master's IN_QTY field
-- is ReadOnly (legacy Quantity_MaskEdit was ReadOnly, parts-stock-master.md:387), so even before this edit
-- it can only re-send the loaded value (no-op-in-effect); after this edit the server-side drop makes the
-- ReadOnly TRULY authoritative — the seam/ledger is the sole owner of IN_QTY. (Belt-and-suspenders.)
--
-- ON-HAND-OVERRIDE POLICY (David, confirmed): on-hand is NEVER editable on the Parts Stock master. All qty
-- change goes through a receiving / shipping / reject / stocktaking / adjustment transaction (i.e. the
-- ledger). This proc edit enforces that server-side.
--
-- PLACEMENT: Phase A (additive proc edit, days ahead — safe pre-window: during parallel run the rebuilt
-- master loads the current @QTY and re-writes the same value, a no-op-in-effect; after the edit it simply
-- stops moving on-hand). ROLLBACK: re-create both procs verbatim from DB Schema/CreateInventory.sql.
--
-- IG81-COMPAT: a plain ALTER PROCEDURE + DROP PROCEDURE, identical on 8.1.52 and 8.3.

-- =============================================================================================
-- (1) UPDATE_PartsStockInfo — drop ONLY the `IN_QTY=@QTY,` assignment; every other column write and the
--     @QTY param are preserved. Body is otherwise verbatim from CreateInventory.sql (line 5691).
-- =============================================================================================
SET ANSI_NULLS ON;
GO
SET QUOTED_IDENTIFIER ON;
GO
ALTER PROCEDURE [dbo].[UPDATE_PartsStockInfo]
	@SupCode 			varchar(5),
	@LogisticsName		varchar(25),
	@PartNum			varchar(12),
	@PartsName			varchar(50),
	@SizeCode			varchar(50),
	@LineName 			varchar(10),
	@KanbanNumber 		varchar(5),
	@LotQty 			 int,
	@QTY 				int,            -- CUTOVER (Carry 10): now UNUSED — on-hand is owned by the ledger.
	                                    -- Kept in the signature so the rebuilt Save's positional 30-param call
	                                    -- is unchanged; the value is accepted-and-ignored.
	@Comments			varchar(300),
	@RenbanCode 		varchar(5),
	@LotSizeOrders 		bit,
	@LeadTime 			int,
	@LeadTimeMonday		int,
	@LeadTimeTuesday		int,
	@LeadTimeWednesday	int,
	@LeadTimeThursday		int,
	@LeadTimeFriday		int,
	@LeadTimeSaturday		int,
	@ShipDays			int,
	@ShipDaysMonday		int,
	@ShipDaysTuesday		int,
	@ShipDaysWednesday	int,
	@ShipDaysThursday		int,
	@ShipDaysFriday		int,
	@ShipDaysSaturday		int,
	@RenbanCount 		varchar(3),
	@PartType			varchar(10),
	@PartCost			money,
	@PartID			int
AS
DECLARE
	@AddDate	varchar(16),
	@Now		datetime,
	@SupplierID	integer,
	@LogisticsID	integer,
	@RenbanID	integer,
	@PartTypeID	integer,
	@SizeID	integer
	SET @Now = getdate()
	SET @AddDate = 	CONVERT(char(8), @Now, 112) +
				SUBSTRING(CONVERT(varchar, @Now, 114), 1, 2) +
				SUBSTRING(CONVERT(varchar, @Now, 114), 4, 2) +
				SUBSTRING(CONVERT(varchar, @Now, 114), 7, 2) +
				SUBSTRING(CONVERT(varchar, @Now, 114), 10, 2)
	SELECT @SupplierID = IN_SUPPLIER_ID FROM INV_SUPPLIER_MST
	WHERE VC_SUPPLIER_CODE = @SupCode
	SELECT @LogisticsID = IN_LOGISTICS_ID FROM INV_LOGISTICS_MST
	WHERE VC_LOGISTICS_NAME = @LogisticsName
	SELECT @RenbanID = IN_RENBAN_ID fROM INV_RENBAN_GROUP_MST
	WHERE VC_RENBAN_GROUP_CODE = @RenbanCode
	SELECT @PartTypeID = IN_PART_TYPE_ID FROM INV_PART_TYPE_MST
	WHERE VC_PART_TYPE = @PartType
	SELECT @SizeID = IN_SIZE_ID FROM INV_SIZE_MST
	WHERE VC_SIZE_CODE = @SizeCode
UPDATE INV_PARTS_STOCK_MST SET
		IN_SUPPLIER_ID=@SupplierID,
		IN_LOGISTICS_ID=@LogisticsID,
		IN_RENBAN_ID=@RenbanID,
		IN_PART_TYPE_ID=@PartTypeID,
		IN_SIZE_ID=@SizeID,
		VC_PART_NUMBER=@PartNum,
		VC_PARTS_NAME=@PartsName,
		IN_RENBAN_COUNT=@RenbanCount,
		VC_KANBAN_NUMBER=@KanbanNumber,
		 IN_LEADTIME=@LeadTime,
		IN_LEADTIME_MONDAY=@LeadTimeMonday,
		IN_LEADTIME_TUESDAY=@LeadTimeTuesday,
		IN_LEADTIME_WEDNESDAY=@LeadTimeWednesday,
		IN_LEADTIME_THURSDAY=@LeadTimeThursday,
		IN_LEADTIME_FRIDAY=@LeadTimeFriday,
		IN_LEADTIME_SATURDAY=@LeadTimeSaturday,
		IN_1LOTQTY=@LotQty,
		-- CUTOVER (Carry 10, David ACCEPTED): IN_QTY=@QTY clause REMOVED. On-hand is never editable on the
		-- rebuilt Parts Stock master; the stock ledger / producer seams are the sole writer of IN_QTY.
		BIT_LOT_SIZE_ORDERS=@LotSizeOrders,
		VC_COMMENTS=@Comments,
		IN_SHIP_DAYS=@ShipDays,
		IN_SHIP_DAYS_MONDAY=@ShipDaysMonday,
		IN_SHIP_DAYS_TUESDAY=@ShipDaysTuesday,
		IN_SHIP_DAYS_WEDNESDAY=@ShipDaysWednesday,
		IN_SHIP_DAYS_THURSDAY=@ShipDaysThursday,
		IN_SHIP_DAYS_FRIDAY=@ShipDaysFriday,
		IN_SHIP_DAYS_SATURDAY=@ShipDaysSaturday,
		MO_PART_COST=@PartCost,
	 	VC_LINE_NAME=@LineName,
		VC_LAST_UPDATE=@AddDate
WHERE IN_PART_ID  = @PartID
GO
PRINT 'UPDATE_PartsStockInfo altered (Carry 10): IN_QTY=@QTY clause dropped; @QTY now unused.';
GO

-- =============================================================================================
-- (2) UPDATE_PartsStockInfoCount — DEAD (zero callers). Retire it.
--     Guarded DROP so the script is idempotent / safe to re-run.
-- =============================================================================================
IF OBJECT_ID('dbo.UPDATE_PartsStockInfoCount', 'P') IS NOT NULL
BEGIN
    DROP PROCEDURE dbo.UPDATE_PartsStockInfoCount;
    PRINT 'UPDATE_PartsStockInfoCount dropped (Carry 10): dead, zero callers.';
END
ELSE
    PRINT 'UPDATE_PartsStockInfoCount already absent.';
GO
