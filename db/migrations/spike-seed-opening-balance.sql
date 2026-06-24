-- spike-seed-opening-balance.sql
-- SEED_OpeningBalance — the CUTOVER opening-balance backfill for the stock ledger (thread b, option 2).
-- Design: docs/analysis/inventory-stock/IGNITION-stock-ledger-design.md §1/§9 (backfill); the from-zero
-- reconstruction is impossible on the purged snapshot (DELETE_AutoPurge aged out the receiving IN-moves,
-- F2), so instead of reconstructing IN_QTY we SEED the ledger with each part's CURRENT IN_QTY as a single
-- opening movement and run forward from there. After seeding, SUM(IN_QTY_CHANGE) == IN_QTY by construction;
-- every forward POST_StockMovement keeps the invariant (it bumps ledger + IN_QTY atomically in lockstep).
--
-- CRITICAL: the opening row does NOT bump IN_QTY — the materialized column ALREADY holds the cutover value;
-- we are only recording the matching ledger row. (Contrast POST_StockMovement, which += delta.) Idempotent
-- on (IN_PART_ID, 'OPENING:part=<id>') so a re-run seeds nothing twice.
--
-- IG81-COMPAT: plain proc via createSProcCall; identical on 8.1.52 and 8.3.

IF OBJECT_ID('dbo.SEED_OpeningBalance', 'P') IS NOT NULL
    DROP PROCEDURE dbo.SEED_OpeningBalance;
GO

CREATE PROCEDURE dbo.SEED_OpeningBalance
    @partId int,
    @site   int = 1
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @evt varchar(100) = 'OPENING:part=' + CONVERT(varchar, @partId);

    -- idempotent: one opening movement per part, ever.
    IF EXISTS (SELECT 1 FROM dbo.INV_STOCK_LEDGER
               WHERE IN_PART_ID = @partId AND VC_SOURCE_EVENT = @evt)
        RETURN;

    -- GENESIS GUARD: the opening balance must be the FIRST movement. The no-bump logic is correct ONLY
    -- because IN_QTY still equals the cutover value — if any forward movement already posted (which
    -- bumped IN_QTY), recording that already-bumped IN_QTY as the opening would double-count. Fail loud
    -- rather than silently drift. (Seeding is a one-shot pre-cutover step; nothing should post first.)
    IF EXISTS (SELECT 1 FROM dbo.INV_STOCK_LEDGER
               WHERE IN_PART_ID = @partId AND VC_SOURCE_ENUM <> 'OPENING_BALANCE')
        THROW 50001, 'SEED_OpeningBalance: part already has forward ledger movements; the opening balance must be seeded BEFORE any posting.', 1;

    DECLARE @qty int = (SELECT IN_QTY FROM dbo.INV_PARTS_STOCK_MST WHERE IN_PART_ID = @partId);
    IF @qty IS NULL RETURN;   -- unknown part: nothing to seed

    DECLARE @ts varchar(16) =
        CONVERT(varchar, GETDATE(), 112) +
        SUBSTRING(CONVERT(varchar, GETDATE(), 114), 1, 2) +
        SUBSTRING(CONVERT(varchar, GETDATE(), 114), 4, 2) +
        SUBSTRING(CONVERT(varchar, GETDATE(), 114), 7, 2) +
        SUBSTRING(CONVERT(varchar, GETDATE(), 114), 10, 2);

    -- record the opening balance as the movement; balance-after == the opening qty. NO IN_QTY bump.
    INSERT INTO dbo.INV_STOCK_LEDGER
        (IN_PART_ID, IN_QTY_CHANGE, IN_BALANCE_AFTER, VC_SOURCE_ENUM,
         IN_SOURCE_ROW_ID, VC_SOURCE_EVENT, site_id, VC_REASON, TS_POSTED, VC_ADD)
    VALUES
        (@partId, @qty, @qty, 'OPENING_BALANCE',
         NULL, @evt, @site, 'cutover opening balance', @ts, @ts);
END
GO

PRINT 'SEED_OpeningBalance created.';
GO

-- ---------------------------------------------------------------------------------------------
-- SEED_AllOpeningBalances — the one-shot cutover backfill for EVERY part, SET-BASED (one statement,
-- one transaction) so it scales to thousands of parts. Genesis-only + idempotent: seeds only parts
-- that have NO ledger row yet, so a re-run (or a part added later) is safe and a part with any forward
-- movement is left alone (its opening should already exist from the pre-cutover seed). This is the
-- proc stockLedger.seedAllOpeningBalances() wraps and the validator exercises — shipped == tested.
IF OBJECT_ID('dbo.SEED_AllOpeningBalances', 'P') IS NOT NULL
    DROP PROCEDURE dbo.SEED_AllOpeningBalances;
GO

CREATE PROCEDURE dbo.SEED_AllOpeningBalances
    @site int = 1
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @ts varchar(16) =
        CONVERT(varchar, GETDATE(), 112) +
        SUBSTRING(CONVERT(varchar, GETDATE(), 114), 1, 2) +
        SUBSTRING(CONVERT(varchar, GETDATE(), 114), 4, 2) +
        SUBSTRING(CONVERT(varchar, GETDATE(), 114), 7, 2) +
        SUBSTRING(CONVERT(varchar, GETDATE(), 114), 10, 2);

    INSERT INTO dbo.INV_STOCK_LEDGER
        (IN_PART_ID, IN_QTY_CHANGE, IN_BALANCE_AFTER, VC_SOURCE_ENUM,
         IN_SOURCE_ROW_ID, VC_SOURCE_EVENT, site_id, VC_REASON, TS_POSTED, VC_ADD)
    SELECT p.IN_PART_ID, p.IN_QTY, p.IN_QTY, 'OPENING_BALANCE',
           NULL, 'OPENING:part=' + CONVERT(varchar, p.IN_PART_ID), @site, 'cutover opening balance', @ts, @ts
    FROM dbo.INV_PARTS_STOCK_MST p
    WHERE NOT EXISTS (SELECT 1 FROM dbo.INV_STOCK_LEDGER l WHERE l.IN_PART_ID = p.IN_PART_ID);
END
GO

PRINT 'SEED_AllOpeningBalances created.';
GO
