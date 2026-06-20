-- spike-asn-unique-guard.sql — FIX 1 (BLOCKER): the DB-level backstop behind the ASN idempotency
-- guard. A FILTERED UNIQUE index on INV_ASN_MST (VC_LINE_NAME, VC_PRODUCTION_DATE) that excludes
-- the hot-call / manual placeholder rows (VC_START_SEQ_NUMBER = '-1').
--
-- WHY (m1-asn-creation-architecture §6 BLOCKER + adversary-findings BLOCKER-1):
--   create_asn (docs/analysis/edi/project-library/asn/code.py) does a NON-ATOMIC check-then-act:
--     SELECT_ASNSeq guard  ->  (no serialization)  ->  INSERT_ASNInfo.
--   On the single-user Delphi desktop the guard was a UI form-lock (structurally serial). The
--   headless multi-session GATEWAY can have two concurrent create_asn calls for the same
--   (line, prodDate) both read 0 rows from the guard and both commit a FULL ASN (header + complete
--   fan-out) -> a duplicate shipment. This is a concurrency risk the REBUILD introduces (desktop ->
--   gateway). The read-back guard alone cannot enforce uniqueness under concurrency; only the DB can.
--
-- KEY = (VC_LINE_NAME, VC_PRODUCTION_DATE) WHERE VC_START_SEQ_NUMBER <> '-1'.
--   * VC_LINE_NAME       varchar(50)  — the assembly line (e.g. 'COROLLA').        [INV_ASN_MST col 4]
--   * VC_PRODUCTION_DATE varchar(8)   — yyyymmdd production date.                   [INV_ASN_MST col 11]
--   * VC_START_SEQ_NUMBER varchar(4)  — the sequence column hot-calls/manual set to '-1'. [col 6]
--   These are the EXACT live column names (verified: sys.columns on Inventory.INV_ASN_MST; the
--   SELECT_ASNSeq proc body keys on the same three and excludes VC_START_SEQ_NUMBER <> -1).
--
-- THE -1 EXCLUSION IS LOAD-BEARING — it mirrors SELECT_ASNSeq's own `AND VC_START_SEQ_NUMBER <> -1`.
--   Hot-call / manual ASNs are stored with VC_START_SEQ_NUMBER = '-1' and LEGITIMATELY repeat per
--   (line, prodDate). Verified on the real data:
--     * Inventory_Live: normal ASNs (seq <> -1) have ZERO duplicate (line, prodDate) groups
--       (GROUP BY ... HAVING COUNT(*)>1 returned 0). So (line, prodDate) IS unique for normal ASNs.
--     * Including hot-calls (no seq filter): 206 duplicate (line, prodDate) groups — these are the
--       -1 rows that MUST be excluded or the constraint would (wrongly) reject valid hot-calls.
--     * The string filter <> '-1' is equivalent to the proc's int <> -1 on the real data (stored
--       value is exactly '-1', 2 chars; SUM(seq<>'-1') == SUM(seq<>-1) == 2235; no value
--       int-converts to -1 without being the literal '-1'). We use the STRING literal '-1' because a
--       filtered-index predicate must be deterministic on the column's own (varchar) type — no
--       implicit int conversion in the filter.
--
-- This index is the BACKSTOP; the driver's read-back guard stays as the fast common-path no-op, and
-- create_asn now ALSO catches the unique-violation race (Msg 2601) and treats it as the idempotent
-- skip (asn/code.py — see the race-handling diff). Belt + suspenders: guard for the 99% case, the
-- unique index for the concurrent-create race the guard cannot see.
--
-- M4 MULTI-SITE: at the M4 schema surgery, IN_SITE_ID joins the key as the LEADING column ->
--   (IN_SITE_ID, VC_LINE_NAME, VC_PRODUCTION_DATE) WHERE VC_START_SEQ_NUMBER <> '-1'. INV_ASN_MST
--   has no IN_SITE_ID column yet, so it is NOT added here. See the `-- M4:` marker on the CREATE.
--
-- IDEMPOTENT: guarded DROP + CREATE (re-runnable). SET options are required for a FILTERED index
--   create (ANSI_NULLS / QUOTED_IDENTIFIER ON, etc.) — set explicitly so the artifact is portable.
--
-- VERIFY (applied + checked on Inventory; see the e2e test):
--   * existing rows do NOT violate it (CREATE succeeds against the live data);
--   * a 2nd NORMAL insert for the same (line, prodDate) is REJECTED (Msg 2601);
--   * a hot-call insert (VC_START_SEQ_NUMBER='-1') for the same (line, prodDate) is ALLOWED.

USE [Inventory];
GO

-- A filtered index requires these SET options ON at CREATE time (and for any later DML on the table).
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
SET ANSI_PADDING ON;
SET ANSI_WARNINGS ON;
SET ARITHABORT ON;
SET CONCAT_NULL_YIELDS_NULL ON;
SET NUMERIC_ROUNDABORT OFF;
GO

IF EXISTS (SELECT 1 FROM sys.indexes
           WHERE name = 'UX_INV_ASN_MST_LINE_PDATE_NORMAL'
             AND object_id = OBJECT_ID('dbo.INV_ASN_MST'))
    DROP INDEX UX_INV_ASN_MST_LINE_PDATE_NORMAL ON dbo.INV_ASN_MST;
GO

-- M4: prepend IN_SITE_ID as the leading key column once INV_ASN_MST has it:
--   ON dbo.INV_ASN_MST (IN_SITE_ID, VC_LINE_NAME, VC_PRODUCTION_DATE) WHERE VC_START_SEQ_NUMBER <> '-1'
CREATE UNIQUE NONCLUSTERED INDEX UX_INV_ASN_MST_LINE_PDATE_NORMAL
    ON dbo.INV_ASN_MST (VC_LINE_NAME, VC_PRODUCTION_DATE)
    WHERE VC_START_SEQ_NUMBER <> '-1';
GO
