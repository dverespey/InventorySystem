-- spike-asndetail-rekey.sql — Q1 cutover proc artifact: re-key INSERT_ASNDetail's upsert
-- existence-check AND DELETE_ASNItem to include IN_ASN_ID, so the per-manifest accumulate is scoped
-- PER ASN instead of per manifest GLOBALLY.
--
-- WHY (m1-asn-creation-spec §5 + Q1 RESOLVED): both procs today key on VC_MANIFEST_NUMBER ALONE.
--   * INSERT_ASNDetail (@HotCall=0): SELECT-by-manifest -> if absent INSERT else IN_QTY += @Qty.
--     A LATER ASN (different production day / reused manifest id) accumulates into the OLD ASN's row.
--     Within one create the accumulate is DESIRED (one manifest hit by several BC/parts sums) — the
--     re-key keeps that and removes the cross-ASN collision.
--   * DELETE_ASNItem(@ManifestNumber): deleting one ASN's manifest line wipes that manifest from
--     EVERY ASN. The re-key scopes the delete to one ASN.
--
-- KEY: (IN_ASN_ID, manifest). The Delphi already passes @ASNID into INSERT_ASNDetail (fRecordID,
--   DataModule.pas:5194/:5246) — IN_ASN_ID is on INV_ASN_DETAIL_MST today — so INSERT_ASNDetail needs
--   NO new param. DELETE_ASNItem gains @ASNID (the caller must now pass which ASN's line to delete).
--
-- M4 SITE INTENT (do NOT implement now): at the M4 multi-site schema surgery the full Q1 key becomes
--   (site_id, IN_ASN_ID, manifest). site_id is NOT on INV_ASN_DETAIL_MST yet, so it is NOT added here.
--   Every WHERE below is laid out so `site_id = @site AND` slots in at the FRONT of the predicate with
--   one line added — see the `-- M4:` markers. The supporting index is likewise widened at M4.
--
-- IDEMPOTENT: guarded DROP+CREATE (re-runnable). Apply on the spike (DB Inventory) to verify the
--   accumulate-within-one-ASN + per-ASN scoping; the rebuild applies this at cutover anyway.
--
-- VERIFIED behaviour preserved: @HotCall=0 accumulate; @HotCall=1 always-INSERT; the 16-char
--   yyyymmddHHmmssff VC_ADD/VC_LAST_UPDATE stamp (style-114: HHmmss + first 2 ms digits — NOT 14);
--   positional VALUES (col2 = IN_ASN_ID).

USE [Inventory];
GO
SET ANSI_NULLS ON;
GO
SET QUOTED_IDENTIFIER ON;
GO

-- ============================================================================================
-- INSERT_ASNDetail — re-keyed upsert: existence-check scoped to (IN_ASN_ID, VC_MANIFEST_NUMBER).
-- ============================================================================================
IF OBJECT_ID('dbo.INSERT_ASNDetail', 'P') IS NOT NULL
    DROP PROCEDURE dbo.INSERT_ASNDetail;
GO
CREATE PROCEDURE [dbo].[INSERT_ASNDetail]
    @ASNID      integer,
    @EIN        integer,
    @Manifest   varchar(8),
    @PartNumber varchar(12),
    @Qty        integer,
    @HotCall    bit = 0
    -- M4: add  @Site integer  here (joins the key as the leading column).
AS
BEGIN
    DECLARE
        @AddDate varchar(16),
        @Now     datetime;

    SET NOCOUNT ON;
    SET @Now = getdate();
    SET @AddDate = CONVERT(char(8), @Now, 112)
                 + SUBSTRING(CONVERT(varchar, @Now, 114), 1, 2)
                 + SUBSTRING(CONVERT(varchar, @Now, 114), 4, 2)
                 + SUBSTRING(CONVERT(varchar, @Now, 114), 7, 2)
                 + SUBSTRING(CONVERT(varchar, @Now, 114), 10, 2);

    IF @HotCall = 0
    BEGIN
        -- Existence check re-keyed: same manifest in a DIFFERENT ASN no longer collides.
        -- M4: add  IN_SITE_ID = @Site AND  to the front of this WHERE.
        IF EXISTS (SELECT 1 FROM INV_ASN_DETAIL_MST
                   WHERE IN_ASN_ID = @ASNID AND VC_MANIFEST_NUMBER = @Manifest)
        BEGIN
            -- accumulate WITHIN this ASN (the intentional, load-bearing sum, spec §5).
            -- M4: add  IN_SITE_ID = @Site AND  to the front of this WHERE.
            UPDATE INV_ASN_DETAIL_MST
               SET IN_QTY = IN_QTY + @Qty
             WHERE IN_ASN_ID = @ASNID AND VC_MANIFEST_NUMBER = @Manifest;
        END
        ELSE
        BEGIN
            -- positional VALUES: (IN_ASN_ID, IN_INV_ID=null, IN_ASN_EIN, manifest, part, qty, last, add)
            -- M4: prepend @Site once INV_ASN_DETAIL_MST has IN_SITE_ID.
            INSERT INTO INV_ASN_DETAIL_MST
                VALUES(@ASNID, null, @EIN, @Manifest, @PartNumber, @Qty, @AddDate, @AddDate);
        END
    END
    ELSE
    BEGIN
        -- @HotCall=1: hot calls never dedup -> always INSERT (unchanged).
        INSERT INTO INV_ASN_DETAIL_MST
            VALUES(@ASNID, null, @EIN, @Manifest, @PartNumber, @Qty, @AddDate, @AddDate);
    END
END
GO

-- ============================================================================================
-- DELETE_ASNItem — re-keyed: deletes only THIS ASN's manifest line (was: every ASN's).
-- ============================================================================================
IF OBJECT_ID('dbo.DELETE_ASNItem', 'P') IS NOT NULL
    DROP PROCEDURE dbo.DELETE_ASNItem;
GO
CREATE PROCEDURE [dbo].[DELETE_ASNItem]
    @ASNID          integer,
    @ManifestNumber varchar(8)
    -- M4: add  @Site integer  (leading key column).
AS
BEGIN
    SET NOCOUNT ON;
    -- M4: add  IN_SITE_ID = @Site AND  to the front of this WHERE.
    DELETE FROM INV_ASN_DETAIL_MST
     WHERE IN_ASN_ID = @ASNID AND VC_MANIFEST_NUMBER = @ManifestNumber;
END
GO

-- ============================================================================================
-- Supporting index. The legacy IX_INV_ASN_DETAIL_MST is on VC_MANIFEST_NUMBER only
-- (m1-asn-creation-spec §5). Widen to the composite key so the re-keyed existence-check + delete
-- seek. Guarded; re-runnable. M4: lead with IN_SITE_ID -> (IN_SITE_ID, IN_ASN_ID, VC_MANIFEST_NUMBER).
-- ============================================================================================
IF EXISTS (SELECT 1 FROM sys.indexes
           WHERE name = 'IX_INV_ASN_DETAIL_MST_ASN_MANIFEST'
             AND object_id = OBJECT_ID('dbo.INV_ASN_DETAIL_MST'))
    DROP INDEX IX_INV_ASN_DETAIL_MST_ASN_MANIFEST ON dbo.INV_ASN_DETAIL_MST;
GO
CREATE NONCLUSTERED INDEX IX_INV_ASN_DETAIL_MST_ASN_MANIFEST
    ON dbo.INV_ASN_DETAIL_MST (IN_ASN_ID, VC_MANIFEST_NUMBER);
GO
