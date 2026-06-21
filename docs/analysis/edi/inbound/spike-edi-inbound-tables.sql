-- spike-edi-inbound-tables.sql — M1 inbound (997/824) rebuild artifacts: the two NEW tables the
-- inbound processor (edi_inbound/code.py) writes. These do NOT exist in the legacy schema (the legacy
-- 997 only UPDATEs status; the 824 was Excel-only, never persisted). They are rebuild-owned tables.
--
--   INV_EDI_INBOUND_LOG  — the idempotency / processed-files LEDGER. Kills the legacy re-ingest hazard
--                          (EDIUpload.pas :430-434 — a failed archive move re-processes the file next
--                          run; the 824 re-emits a duplicate Excel report). CONTENT IS THE IDENTITY: a
--                          file is deduped by VC_FILE_HASH ALONE, so a re-drop of the same content is a
--                          NO-OP EVEN UNDER A NEW NAME (a renamed VAN/mailer redelivery — common — no
--                          longer re-processes; closes adversary SHOULD-FIX-1: a renamed 824 redelivery
--                          would otherwise DOUBLE-flip + write DUPLICATE alarms). A same-name/new-content
--                          file IS reprocessed (new hash = new event). VC_FILE_NAME is recorded for audit
--                          / archive traceability but is NOT part of the dedup key. Written INSIDE the
--                          per-file tx so the 'PROCESSED' row commits atomically with the status flips
--                          (a rolled-back file is never recorded processed). Also records QUARANTINED
--                          (no-DUNS-match / not-EDI / out-of-M1-scope) files (Q11 — NOT silently dropped).
--
--   INV_EDI_ALARM_REJ    — the 824 reject / 997-unknown-EIN ALARM records (Q10). The DATA/flag side of
--                          the main-screen alarm + click-to-detail. One row per 824 reject line
--                          (manifest/part/errorText) + one row per 997 unknown-EIN ack. BIT_RESOLVED=0
--                          = active/unacknowledged -> the home hub surfaces it as the main-screen alarm.
--                          The real gateway NATIVE alarm pipeline (system.alarm / an alarm-journal tag
--                          off these rows) is the PROD follow-on; the DATA side here is what makes the
--                          behavior testable headlessly.
--
-- M4 SITE INTENT (do NOT implement now): both tables carry IN_SITE_ID from day one (these are NEW
--   rebuild tables, so unlike the legacy INV_ASN_MST / INV_INV_MST — which lack IN_SITE_ID until the
--   M4 schema surgery — there is no migration cost to including it). The status-flip UPDATEs in
--   edi_inbound/code.py against INV_ASN_MST / INV_INV_MST stay -- M4 markers (those legacy tables have
--   no IN_SITE_ID yet); the resolved siteId is threaded + stored HERE so the alarm/ledger are already
--   site-scoped and the M4 add to the legacy-table WHEREs is one line.
--
-- IDEMPOTENT: guarded CREATEs (re-runnable; existing rows preserved). Apply on the spike (DB Inventory)
--   to exercise the inbound processor; the rebuild applies this at cutover.

USE [Inventory];
GO
SET ANSI_NULLS ON;
GO
SET QUOTED_IDENTIFIER ON;
GO

-- ============================================================================================
-- INV_EDI_INBOUND_LOG — processed-files ledger (idempotency).
-- ============================================================================================
IF OBJECT_ID('dbo.INV_EDI_INBOUND_LOG', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.INV_EDI_INBOUND_LOG (
        IN_LOG_ID     int IDENTITY(1,1) NOT NULL CONSTRAINT PK_INV_EDI_INBOUND_LOG PRIMARY KEY,
        VC_FILE_NAME  varchar(255)  NOT NULL,
        VC_FILE_HASH  varchar(64)   NOT NULL,   -- sha256 hex of the file body (content dedupe key)
        VC_EDI_TYPE   varchar(4)    NULL,        -- '997' | '824' | '??' (not-EDI/unknown)
        IN_SITE_ID    int           NULL,        -- the resolved site (NULL for a quarantined no-DUNS file)
        VC_STATUS     varchar(16)   NOT NULL,    -- 'PROCESSED' | 'QUARANTINED'
        VC_DETAIL     varchar(400)  NULL,        -- a short human summary (counts / reason)
        DT_PROCESSED  datetime      NOT NULL CONSTRAINT DF_INV_EDI_INBOUND_LOG_DT DEFAULT (GETDATE())
    );
    -- the dedupe seek: VC_FILE_HASH (the content identity) where the file is PROCESSED.
    -- _already_processed keys on the HASH ALONE (content = identity; a renamed redelivery is a no-op),
    -- so the index LEADS on (VC_FILE_HASH, VC_STATUS). A non-unique index (a file CAN be re-attempted
    -- after a quarantine fix, producing a second row) — _already_processed filters to
    -- VC_STATUS='PROCESSED'. VC_FILE_NAME is INCLUDEd (covering, for audit lookups), not a key column.
    CREATE NONCLUSTERED INDEX IX_INV_EDI_INBOUND_LOG_HASH
        ON dbo.INV_EDI_INBOUND_LOG (VC_FILE_HASH, VC_STATUS)
        INCLUDE (VC_FILE_NAME);
END;
GO

-- ============================================================================================
-- INV_EDI_ALARM_REJ — 824 reject + 997 unknown-EIN alarm records (the home-hub flag + detail).
-- ============================================================================================
IF OBJECT_ID('dbo.INV_EDI_ALARM_REJ', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.INV_EDI_ALARM_REJ (
        IN_ALARM_ID         int IDENTITY(1,1) NOT NULL CONSTRAINT PK_INV_EDI_ALARM_REJ PRIMARY KEY,
        IN_SITE_ID          int           NULL,    -- the resolved site (NULL only for a siteless alarm)
        VC_ALARM_TYPE       varchar(32)   NOT NULL, -- '824_ASN_REJECT' | '997_UNKNOWN_EIN'
        VC_SUMMARY          varchar(200)  NULL,
        VC_MANIFEST_NUMBER  varchar(8)    NULL,    -- the 824 manifest (the ASN join key)
        VC_ASSY_PART_NUMBER varchar(12)   NULL,    -- the 824 part number
        VC_ERROR_TEXT       varchar(200)  NULL,    -- the 824 free-text reason
        IN_EIN              int           NULL,    -- the 997 unknown EIN (NULL for 824 rows)
        BIT_RESOLVED        bit           NOT NULL CONSTRAINT DF_INV_EDI_ALARM_REJ_RES DEFAULT (0),
        DT_RAISED           datetime      NOT NULL CONSTRAINT DF_INV_EDI_ALARM_REJ_DT  DEFAULT (GETDATE())
    );
    -- the home-hub "active alarms" surface: unresolved rows by site, newest first.
    CREATE NONCLUSTERED INDEX IX_INV_EDI_ALARM_REJ_ACTIVE
        ON dbo.INV_EDI_ALARM_REJ (BIT_RESOLVED, IN_SITE_ID, DT_RAISED);
END;
GO
