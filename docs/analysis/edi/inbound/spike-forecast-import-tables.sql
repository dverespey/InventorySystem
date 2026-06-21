-- spike-forecast-import-tables.sql — M2 EDI 830 forecast-import rebuild DATA-side artifacts.
--
-- The M2 importer (forecast/code.py) REUSES the M1 inbound tables (spike-edi-inbound-tables.sql):
--   INV_EDI_INBOUND_LOG   — the idempotency / processed-files ledger (VC_EDI_TYPE='830'). The
--                           delete-then-accumulate is only correct ONCE per content hash; a re-drop of
--                           the same 830 is a NO-OP (proves the additive breakdown is not doubled by a
--                           re-poll).
--   INV_EDI_ALARM_REJ     — alarms. M2 adds two new VC_ALARM_TYPE values (the column is varchar(32), no
--                           DDL change needed):
--                             '830_FORECAST_GAP'   — D-Bug-1: a missing BOM recipe / ratio match. The
--                                                    assembly goes in VC_ASSY_PART_NUMBER; the week/qty in
--                                                    VC_SUMMARY/VC_ERROR_TEXT. Replaces the legacy SILENT
--                                                    drop (ForecastBreakdownF.pas:1178-1183).
--                             '830_FORECAST_STALE' — Q11: no successful 830 import in >= 8 days for an
--                                                    AUTO-mode site (stale_sites()).
--
-- The Q11 per-site config columns ALREADY EXIST on INV_SITES (added in the master-data sites work,
-- spike-inv-sites-table.sql): VC_FORECAST_IMPORT_MODE ('AUTO'|'MANUAL'), VC_LAST_FORECAST_IMPORT
-- (16-char stamp), IN_FORECAST_USAGE_COMPARE (the usage-rollup compare weeks), BIT_USE_FIRST_PRODUCTION_DAY.
--
-- THE ONE NEW COLUMN this file adds: IN_HISTORICAL_FORECAST — the legacy [INIT] HistoricalForecast value
--   (DataModule.fiHistoricalForecast, default 12 weeks) that drives DELETE_ForecastInfo's @HistWeekDate
--   prune window (now - HistoricalForecast*7 days; ForecastBreakdownF.pas:318). It was NOT on INV_SITES.
--   Adding it makes the prune window per-site config (faithful to the single-site INI today; needed once
--   sites diverge). The driver defaults to 12 when the column/value is absent (DEFAULT_HISTORICAL_WEEKS).
--
-- IDEMPOTENT: guarded ALTER + UPDATE (re-runnable). Apply on the spike (DB Inventory). Restores nothing
--   on its own — the e2e test cleans its synthetic rows; this DDL is additive config and is left in place
--   (the column is part of the rebuild schema, not test scaffolding).

USE [Inventory];
GO
SET ANSI_NULLS ON;
GO
SET QUOTED_IDENTIFIER ON;
GO

-- IN_HISTORICAL_FORECAST: per-site history-prune window in WEEKS (legacy [INIT] HistoricalForecast).
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('dbo.INV_SITES') AND name = 'IN_HISTORICAL_FORECAST')
BEGIN
    ALTER TABLE dbo.INV_SITES
        ADD IN_HISTORICAL_FORECAST INT NULL
            CONSTRAINT DF_INV_SITES_HIST_FC DEFAULT (12);   -- 12 weeks = 84-day prune (legacy default)
END;
GO

-- backfill any existing NULLs to the legacy default so the prune window is defined for every site.
UPDATE dbo.INV_SITES SET IN_HISTORICAL_FORECAST = 12 WHERE IN_HISTORICAL_FORECAST IS NULL;
GO
