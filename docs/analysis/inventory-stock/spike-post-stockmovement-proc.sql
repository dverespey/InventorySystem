-- spike-post-stockmovement-proc.sql
-- POST_StockMovement — the ONLY writer of INV_PARTS_STOCK_MST.IN_QTY.
-- Design: docs/analysis/inventory-stock/IGNITION-stock-ledger-design.md §5.
--
-- Atomic per post: INSERT INV_STOCK_LEDGER + additive UPDATE IN_QTY = IN_QTY + @delta,
-- in one transaction. Idempotent on (IN_PART_ID, VC_SOURCE_EVENT). Purge-aware.
--
-- §4 concurrency argument: the additive `IN_QTY += @delta` is commutative and never reads
-- the prior value into the app, so post() needs NO SERIALIZABLE — READ COMMITTED + the
-- additive UPDATE + the UNIQUE event key is sufficient. (rebuildBalance, the one read-then-write
-- absolute re-stamp, DOES take an exclusive lock — see PROC_RebuildStockBalance below.)
--
-- IG81-COMPAT: plain stored proc invoked via createSProcCall — identical on 8.1.52 and 8.3.
-- IG83-TODO: TS_POSTED/VC_ADD -> datetime2 at the Postgres phase.

SET NOCOUNT ON;
GO

IF OBJECT_ID('dbo.POST_StockMovement', 'P') IS NOT NULL
    DROP PROCEDURE dbo.POST_StockMovement;
GO

CREATE PROCEDURE dbo.POST_StockMovement
    @partId      int,
    @delta       int,
    @sourceEnum  varchar(24),
    @sourceRowId int          = NULL,
    @eventKey    varchar(100),
    @reason      varchar(300) = NULL,
    @site        int          = 1,
    @purge       bit          = 0
AS
BEGIN
    SET NOCOUNT ON;

    -- 1. Purge / housekeeping delete posts nothing (§3.1: mirrors DELETE_AutoPurge pre-stamp +
    --    the live VC_TERMINATED='' gate, NOT a fictional PurgeMode flag). The caller decides
    --    purge=1 for a housekeeping/already-terminated delete.
    IF @purge = 1
        RETURN;

    -- 2. Idempotency (§4): a replay of the same (part, event) is a no-op; IN_QTY += delta must
    --    NOT run again. Backstopped by UQ_INV_STOCK_LEDGER_EVENT.
    IF EXISTS (SELECT 1 FROM dbo.INV_STOCK_LEDGER
               WHERE IN_PART_ID = @partId AND VC_SOURCE_EVENT = @eventKey)
        RETURN;

    DECLARE @ts varchar(16) =
        CONVERT(varchar, GETDATE(), 112) +                         -- yyyymmdd
        SUBSTRING(CONVERT(varchar, GETDATE(), 114), 1, 2) +        -- HH
        SUBSTRING(CONVERT(varchar, GETDATE(), 114), 4, 2) +        -- MM
        SUBSTRING(CONVERT(varchar, GETDATE(), 114), 7, 2) +        -- SS
        SUBSTRING(CONVERT(varchar, GETDATE(), 114), 10, 2);        -- ff (hundredths)

    BEGIN TRY
        BEGIN TRAN;

            -- (a) append the signed movement, capturing the balance after the bump.
            DECLARE @balanceAfter int =
                (SELECT IN_QTY FROM dbo.INV_PARTS_STOCK_MST WHERE IN_PART_ID = @partId) + @delta;

            INSERT INTO dbo.INV_STOCK_LEDGER
                (IN_PART_ID, IN_QTY_CHANGE, IN_BALANCE_AFTER, VC_SOURCE_ENUM,
                 IN_SOURCE_ROW_ID, VC_SOURCE_EVENT, site_id, VC_REASON, TS_POSTED, VC_ADD)
            VALUES
                (@partId, @delta, @balanceAfter, @sourceEnum,
                 @sourceRowId, @eventKey, @site, @reason, @ts, @ts);

            -- (b) the additive materialized bump — byte-identical in effect to the legacy triggers
            --     (all 12 do IN_QTY = IN_QTY +/- x). Single relative statement: commutative, no race.
            UPDATE dbo.INV_PARTS_STOCK_MST
                SET IN_QTY = IN_QTY + @delta,
                    VC_LAST_UPDATE = @ts
                WHERE IN_PART_ID = @partId;

        COMMIT TRAN;
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK TRAN;
        -- A UNIQUE-violation here means a concurrent post of the same event won the race;
        -- treat it as the idempotent no-op (error 2627/2601), otherwise re-raise.
        IF ERROR_NUMBER() NOT IN (2627, 2601)
            THROW;
    END CATCH
END
GO

PRINT 'POST_StockMovement created.';
GO

-- ---------------------------------------------------------------------------------------------
-- PROC_RebuildStockBalance — the healing command behind stockLedger.rebuildBalance(partId).
-- F4 invariant (design §4 point 6 / §7): this is the ONE read-then-write absolute re-stamp.
-- It re-SUMs the ledger and stamps IN_QTY = <absolute SUM>. A post() delta committing between
-- the SUM and the re-stamp would be SILENTLY CLOBBERED (lost update), so it MUST hold an
-- exclusive row lock across the whole read-then-write. We take it with (UPDLOCK, HOLDLOCK) on
-- the part row BEFORE the SUM and hold it through the re-stamp inside one SERIALIZABLE tran.
-- IG81-COMPAT: DB-level locking — identical on 8.1.52 and 8.3.
IF OBJECT_ID('dbo.PROC_RebuildStockBalance', 'P') IS NOT NULL
    DROP PROCEDURE dbo.PROC_RebuildStockBalance;
GO

CREATE PROCEDURE dbo.PROC_RebuildStockBalance
    @partId int
AS
BEGIN
    SET NOCOUNT ON;
    SET TRANSACTION ISOLATION LEVEL SERIALIZABLE;
    BEGIN TRY
        BEGIN TRAN;
            -- Take + hold the exclusive lock on the part row first (fences any concurrent post()).
            DECLARE @dummy int =
                (SELECT IN_QTY FROM dbo.INV_PARTS_STOCK_MST WITH (UPDLOCK, HOLDLOCK)
                 WHERE IN_PART_ID = @partId);

            DECLARE @sum int =
                (SELECT ISNULL(SUM(IN_QTY_CHANGE), 0) FROM dbo.INV_STOCK_LEDGER
                 WHERE IN_PART_ID = @partId);

            UPDATE dbo.INV_PARTS_STOCK_MST
                SET IN_QTY = @sum
                WHERE IN_PART_ID = @partId;
        COMMIT TRAN;
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK TRAN;
        THROW;
    END CATCH
END
GO

PRINT 'PROC_RebuildStockBalance created.';
GO
