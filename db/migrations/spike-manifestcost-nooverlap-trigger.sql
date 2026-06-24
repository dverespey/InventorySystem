-- spike-manifestcost-nooverlap-trigger.sql
-- CUTOVER artifact — Carry 3 / D6 manifest-index decision (David 2026-06-19).
--
-- DECISION (verbatim-faithful): "Manifest index — DROP + REPLACE. At cutover drop
-- IX_INV_MANIFEST_COST_MST (if prod still has it) so the rebuilt master can hold multiple windows per
-- part; integrity enforced by the existing app checkWindowOverlap guard + a NEW DB trigger
-- TRG_ManifestCost_NoOverlap. A pre-drop overlap diagnostic flags existing overlaps to clean first."
--
-- WHY. The live IX_INV_MANIFEST_COST_MST is UNIQUE on VC_ASSY_MANIFEST_NUMBER (the 2-char manifest #),
-- NOT on the assy part code. It does not directly block a part from carrying >1 price window, but the
-- rebuilt ManifestCost master relies on it being gone (it lets a part carry >1 manifest #/window). What it
-- CANNOT give us is the real integrity rule — non-OVERLAPPING windows per part — because that is a
-- relationship BETWEEN rows, which neither a UNIQUE index nor a CHECK constraint can see. So: drop the
-- old part-blind unique, and replace the integrity with a trigger that sees sibling rows.
--
-- *** SCHEMA SURPRISE vs the architecture-doc draft (READ THIS) ***
-- The arch doc (carry 3) drafted the trigger with GUESSED names. Verified against the live spike:
--   COLUMNS (CONFIRMED, the guesses were CORRECT):
--     IN_MANIFEST_COST_ID     int IDENTITY  (the surrogate id; used for the "<> self" exclusion)
--     VC_ASSY_PART_NUMBER_CODE varchar(12)  (the part key)
--     VC_START_MANIFEST       varchar(8)    (window start, yyyymmdd STRING)
--     VC_END_MANIFEST         varchar(8)    (window end,   yyyymmdd STRING)
--   So window comparisons are STRING compares — correct ONLY because the stamps are zero-padded yyyymmdd
--   (lexicographic == chronological). This matches fn_ManifestCostAt and the app guard. (IG83-TODO: typed
--   date at the Postgres phase.)
--   PK / INDEX (SURPRISE): there is NO primary key on this table — it is a HEAP. IX_INV_MANIFEST_COST_MST
--     is a UNIQUE NONCLUSTERED *CONSTRAINT* (CONSTRAINT [IX_INV_MANIFEST_COST_MST] UNIQUE NONCLUSTERED
--     (VC_ASSY_MANIFEST_NUMBER), schema dump line 867), NOT a standalone CREATE INDEX. Dropping it
--     therefore needs ALTER TABLE ... DROP CONSTRAINT, not DROP INDEX. The conditional block below handles
--     BOTH forms (constraint OR index) defensively so it works whatever prod actually has.
--   On the SPIKE the constraint was ALREADY dropped during the ManifestCost master build (the table is a
--     pure heap now) — so step (B) is a no-op on the spike and only does real work against prod.
--
-- INTEGRITY MODEL: app guard checkWindowOverlap (built + tested) rejects overlaps at the UI; this trigger
-- is the DB-level BACKSTOP that catches any write bypassing the app (legacy Delphi during parallel run, a
-- manual SQL fix, a future second client). The predicate is IDENTICAL to the app guard.
--
-- *** WRITER COMPATIBILITY NOTE (verified) — OUTPUT-without-INTO ***
-- SQL Server forbids `INSERT ... OUTPUT inserted.col` (no INTO clause) on a table that has an ENABLED
-- trigger. The REBUILT ManifestCost master Save is SAFE: it uses
--     INSERT ...; SELECT CAST(SCOPE_IDENTITY() AS int) AS newId
-- (view.json Save action) — no OUTPUT clause — so the trigger does not affect it. The only writer that
-- uses OUTPUT-without-INTO is the headless test scripts/e2e/test_manifestcost_overlap_guard.py (its anchor
-- INSERT). That test pins the PREDICATE independently of the DB trigger (it COUNTs the guard expression),
-- so it is run with the trigger NOT present; it does not need the trigger and would need an `OUTPUT ... INTO
-- @tbl` rewrite to run with the trigger enabled. CUTOVER takeaway: any future direct writer using
-- OUTPUT-without-INTO against INV_MANIFEST_COST_MST must switch to `OUTPUT ... INTO @tbl` (or SCOPE_IDENTITY).
--
-- IG81-COMPAT: AFTER trigger + ALTER TABLE DROP CONSTRAINT are identical on 8.1.52 and 8.3.
-- ROLLBACK: DROP TRIGGER dbo.TRG_ManifestCost_NoOverlap; re-create the unique constraint
--   (ALTER TABLE dbo.INV_MANIFEST_COST_MST ADD CONSTRAINT IX_INV_MANIFEST_COST_MST
--    UNIQUE NONCLUSTERED (VC_ASSY_MANIFEST_NUMBER)).

-- =============================================================================================
-- (A) PRE-DROP OVERLAP DIAGNOSTIC — run FIRST. Expect ZERO rows. Any hit = existing data already overlaps
--     and would make the trigger block future edits of that part; David cleans those windows before (C).
--     (Same predicate as fn_ManifestCostAt's overlap diagnostic + the app guard: inclusive both ends.)
-- =============================================================================================
SELECT  a.VC_ASSY_PART_NUMBER_CODE                         AS PartCode,
        a.IN_MANIFEST_COST_ID AS Id_A, a.VC_ASSY_MANIFEST_NUMBER AS Manifest_A,
        a.VC_START_MANIFEST AS Start_A, a.VC_END_MANIFEST AS End_A,
        b.IN_MANIFEST_COST_ID AS Id_B, b.VC_ASSY_MANIFEST_NUMBER AS Manifest_B,
        b.VC_START_MANIFEST AS Start_B, b.VC_END_MANIFEST AS End_B
FROM    INV_MANIFEST_COST_MST a
JOIN    INV_MANIFEST_COST_MST b
        ON a.VC_ASSY_PART_NUMBER_CODE = b.VC_ASSY_PART_NUMBER_CODE
       AND a.IN_MANIFEST_COST_ID < b.IN_MANIFEST_COST_ID
       AND a.VC_START_MANIFEST <= b.VC_END_MANIFEST
       AND b.VC_START_MANIFEST <= a.VC_END_MANIFEST
ORDER BY a.VC_ASSY_PART_NUMBER_CODE, a.IN_MANIFEST_COST_ID;
GO

-- =============================================================================================
-- (B) CONDITIONAL DROP of IX_INV_MANIFEST_COST_MST — handles BOTH the constraint form (live/prod) and the
--     plain-index form, guarded by existence checks. No-op if already gone (spike).
-- =============================================================================================
IF EXISTS (SELECT 1 FROM sys.key_constraints
           WHERE name='IX_INV_MANIFEST_COST_MST' AND parent_object_id=OBJECT_ID('dbo.INV_MANIFEST_COST_MST'))
BEGIN
    ALTER TABLE dbo.INV_MANIFEST_COST_MST DROP CONSTRAINT IX_INV_MANIFEST_COST_MST;
    PRINT 'Dropped UNIQUE CONSTRAINT IX_INV_MANIFEST_COST_MST (part-blind manifest-# unique).';
END
ELSE IF EXISTS (SELECT 1 FROM sys.indexes
                WHERE name='IX_INV_MANIFEST_COST_MST' AND object_id=OBJECT_ID('dbo.INV_MANIFEST_COST_MST'))
BEGIN
    DROP INDEX IX_INV_MANIFEST_COST_MST ON dbo.INV_MANIFEST_COST_MST;
    PRINT 'Dropped INDEX IX_INV_MANIFEST_COST_MST.';
END
ELSE
    PRINT 'IX_INV_MANIFEST_COST_MST already absent (spike / already migrated) — no-op.';
GO

-- =============================================================================================
-- (C) THE NO-OVERLAP BACKSTOP TRIGGER. Rejects any INSERTed/UPDATEd window that overlaps an existing
--     window for the SAME part. Predicate is the app guard's: two windows overlap unless one ends strictly
--     before the other starts (inclusive boundary -> touching windows DO overlap -> rejected; gap windows
--     do not). The self-exclusion (m.IN_MANIFEST_COST_ID <> i.IN_MANIFEST_COST_ID) lets a row's own update
--     pass against itself.
--     It ALSO rejects a malformed window first (THROW 50011): a blank/empty VC_START_MANIFEST or
--     VC_END_MANIFEST, or an inverted window (start > end). A blank window is a garbage-row hazard — string
--     compare '' < anydate would silently ACCEPT it past the overlap check — so we close it cheaply here.
-- =============================================================================================
IF OBJECT_ID('dbo.TRG_ManifestCost_NoOverlap', 'TR') IS NOT NULL
    DROP TRIGGER dbo.TRG_ManifestCost_NoOverlap;
GO
CREATE TRIGGER dbo.TRG_ManifestCost_NoOverlap
ON dbo.INV_MANIFEST_COST_MST
AFTER INSERT, UPDATE
AS
BEGIN
    SET NOCOUNT ON;

    -- (0) WELL-FORMED-WINDOW guard. A blank/empty start or end, or an inverted window
    --     (start > end), is a garbage row: it cannot legitimately overlap-check (string compare
    --     '' < anydate silently ACCEPTS a blank window) and has no meaning. Reject it outright.
    --     Empty test covers NULL and '' (NULLIF(...,'') IS NULL); inverted test is the same
    --     yyyymmdd string compare the overlap predicate uses.
    IF EXISTS (
        SELECT 1
        FROM   inserted i
        WHERE  NULLIF(i.VC_START_MANIFEST, '') IS NULL
            OR NULLIF(i.VC_END_MANIFEST,   '') IS NULL
            OR i.VC_START_MANIFEST > i.VC_END_MANIFEST
    )
    BEGIN
        ROLLBACK TRANSACTION;
        THROW 50011, 'Manifest cost window is malformed (blank start/end or start after end).', 1;
    END

    IF EXISTS (
        SELECT 1
        FROM   inserted i
        JOIN   dbo.INV_MANIFEST_COST_MST m
               ON  m.VC_ASSY_PART_NUMBER_CODE = i.VC_ASSY_PART_NUMBER_CODE
               AND m.IN_MANIFEST_COST_ID     <> i.IN_MANIFEST_COST_ID
               -- overlap == NOT (one ends before the other starts); inclusive both ends.
               AND NOT (i.VC_END_MANIFEST < m.VC_START_MANIFEST OR i.VC_START_MANIFEST > m.VC_END_MANIFEST)
    )
    BEGIN
        ROLLBACK TRANSACTION;
        THROW 50010, 'Manifest cost window overlaps an existing window for this part.', 1;
    END
END
GO
PRINT 'TRG_ManifestCost_NoOverlap created (DB-level no-overlap backstop).';
GO
