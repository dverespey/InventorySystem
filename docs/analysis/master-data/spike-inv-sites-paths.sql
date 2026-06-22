/* ============================================================================
   INV_SITES — PATH COLUMNS ADDENDUM  (M4 piece 1, David 2026-06-22)
   ----------------------------------------------------------------------------
   Adds the per-deployment directory columns the legacy Delphi app held in the
   INI [DIRECTORIES] section (read into DataModule TCIniField properties). This is
   an ADDITIVE, IDEMPOTENT addendum to spike-inv-sites-table.sql — same shape as
   the M2 forecast-config addendum (spike-forecast-import-tables.sql's
   IN_HISTORICAL_FORECAST ALTER): guarded `IF NOT EXISTS` ALTERs that re-run safe.

   ----------------------------------------------------------------------------
   DIRECTION (David 2026-06-22 — the direction reversal):
     Each site runs on its OWN gateway + DB (single-site deployments, NOT a shared
     multi-tenant DB). So there is NO site_id surgery. INV_SITES holds the ONE
     site's config per deployment (one row). These path columns are simply that
     deployment's directory configuration, relocated out of the per-machine INI
     into the DB (David's call: keep sites config in the DB).

   ----------------------------------------------------------------------------
   THE LEGACY [DIRECTORIES] SET (verified in DataModule.dfm, the TCIniField defs):
     Key                  Field (DataModule)        Default (legacy)
     -------------------  ------------------------  --------------------------------
     EDIOut               fiEDIOut                  c:\_Inventory_Control\EDIOut
     EDIIn                fiEDIIn                   c:\_Inventory_Control\EDIIn
     ForecastInputDir     fiForecastInputDir        c:\_Inventory_Control\Suppliers\NUMMI
     LogisticsInputDir    fiLogisticsInputDir       c:\_Inventory_Control\
     ReportsOutputDir     fiReportsOutputDir        c:\_Inventory_Control\Reports
     TextShippingFileDir  fiTextShippingFileDir     D:\Daily Camex Results\
     TemplateDir          fiTemplateDir             c:\_Inventory_Control\Templates
   (`LocalFTP` lives in [INIT], is a BOOLEAN flag — NOT a directory — so it is NOT
    a path column. The other [DIRECTORIES] keys above are the complete dir set.)

   COLUMN -> LEGACY KEY MAPPING (this addendum):
     VC_EDIOUT_DIR    <- [DIRECTORIES] EDIOut               (856/810 outbound write dir)
     VC_EDIIN_DIR     <- [DIRECTORIES] EDIIn                (inbound 997/824/830/862 poll dir)
     VC_FORECAST_DIR  <- [DIRECTORIES] ForecastInputDir     (forecast .frc import dir)
     VC_LOGISTICS_DIR <- [DIRECTORIES] LogisticsInputDir    (logistics import dir)
     VC_REPORTS_DIR   <- [DIRECTORIES] ReportsOutputDir     (report output dir)
     VC_SHIPPING_DIR  <- [DIRECTORIES] TextShippingFileDir  (daily shipping-result text dir)
     VC_TEMPLATE_DIR  <- [DIRECTORIES] TemplateDir          (order/report template dir)

   SHARED vs PER-DEPLOYMENT (a deployment-config fact, NOT a schema flag):
     The EDI dirs (EDIOut/EDIIn) are a SHARED network path for now — both sites'
     gateways point at the SAME EDI share until it's split. All other dirs are
     per-deployment / local. There is no special "shared" column or flag: these
     are all just config values here, and "shared" is just what value the two
     deployments happen to be given. (Note recorded so a future split is a
     value change, not a schema change.)

   WIDTH: varchar(260) — Windows MAX_PATH (260) headroom; legacy paths are short
     (the longest default is ~38 chars) but a UNC share or a deep tree can be long.

   SECRETS / SAFETY: SEED PLACEHOLDER example paths ONLY. NEVER commit real client
     paths (the legacy 'c:\_Inventory_Control\...' / 'D:\Daily Camex Results\' are
     the legacy DEFAULTS, not client-specific, but the real deployed values load at
     cutover from the deployment's config — they are NOT in this repo).

   8.1 <-> 8.3:
     # IG81-COMPAT: plain ALTER/UPDATE; runs identically on 8.1.52 and 8.3.
     # IG83-TODO: at the Postgres phase, paths may move to a gateway-config /
                  named-path resolution scheme; the columns stay as the source.

   IDEMPOTENT: guarded ALTER + a guarded placeholder backfill (re-runnable). Apply
     on the spike (DB Inventory). This DDL is ADDITIVE rebuild schema (like the M2
     forecast-config columns) — left in place; it is NOT test scaffolding.
   ============================================================================ */

USE [Inventory];
GO
SET ANSI_NULLS ON;
GO
SET QUOTED_IDENTIFIER ON;
GO
SET NOCOUNT ON;
GO

/* --- EDIOut (shared EDI share for now) --- */
IF NOT EXISTS (SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('dbo.INV_SITES') AND name = 'VC_EDIOUT_DIR')
    ALTER TABLE dbo.INV_SITES ADD VC_EDIOUT_DIR VARCHAR(260) NULL;
GO

/* --- EDIIn (shared EDI share for now) --- */
IF NOT EXISTS (SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('dbo.INV_SITES') AND name = 'VC_EDIIN_DIR')
    ALTER TABLE dbo.INV_SITES ADD VC_EDIIN_DIR VARCHAR(260) NULL;
GO

/* --- ForecastInputDir (per-deployment) --- */
IF NOT EXISTS (SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('dbo.INV_SITES') AND name = 'VC_FORECAST_DIR')
    ALTER TABLE dbo.INV_SITES ADD VC_FORECAST_DIR VARCHAR(260) NULL;
GO

/* --- LogisticsInputDir (per-deployment) --- */
IF NOT EXISTS (SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('dbo.INV_SITES') AND name = 'VC_LOGISTICS_DIR')
    ALTER TABLE dbo.INV_SITES ADD VC_LOGISTICS_DIR VARCHAR(260) NULL;
GO

/* --- ReportsOutputDir (per-deployment) --- */
IF NOT EXISTS (SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('dbo.INV_SITES') AND name = 'VC_REPORTS_DIR')
    ALTER TABLE dbo.INV_SITES ADD VC_REPORTS_DIR VARCHAR(260) NULL;
GO

/* --- TextShippingFileDir (per-deployment) --- */
IF NOT EXISTS (SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('dbo.INV_SITES') AND name = 'VC_SHIPPING_DIR')
    ALTER TABLE dbo.INV_SITES ADD VC_SHIPPING_DIR VARCHAR(260) NULL;
GO

/* --- TemplateDir (per-deployment) — completes the legacy [DIRECTORIES] set --- */
IF NOT EXISTS (SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('dbo.INV_SITES') AND name = 'VC_TEMPLATE_DIR')
    ALTER TABLE dbo.INV_SITES ADD VC_TEMPLATE_DIR VARCHAR(260) NULL;
GO

/* ----------------------------------------------------------------------------
   PLACEHOLDER backfill — example paths only (NEVER real client paths). Only
   touches rows where the path is still NULL, so a re-run after a Sites-screen
   edit does NOT clobber a configured value. Placeholders use a neutral
   /opt/ignition/edi-style scheme (clearly example, not a real client path).
   ---------------------------------------------------------------------------- */
UPDATE dbo.INV_SITES
SET VC_EDIOUT_DIR    = COALESCE(VC_EDIOUT_DIR,    '/opt/ignition/shared/edi/out'),
    VC_EDIIN_DIR     = COALESCE(VC_EDIIN_DIR,     '/opt/ignition/shared/edi/in'),
    VC_FORECAST_DIR  = COALESCE(VC_FORECAST_DIR,  '/opt/ignition/site/forecast/in'),
    VC_LOGISTICS_DIR = COALESCE(VC_LOGISTICS_DIR, '/opt/ignition/site/logistics/in'),
    VC_REPORTS_DIR   = COALESCE(VC_REPORTS_DIR,   '/opt/ignition/site/reports/out'),
    VC_SHIPPING_DIR  = COALESCE(VC_SHIPPING_DIR,  '/opt/ignition/site/shipping/out'),
    VC_TEMPLATE_DIR  = COALESCE(VC_TEMPLATE_DIR,  '/opt/ignition/site/templates')
WHERE VC_EDIOUT_DIR IS NULL OR VC_EDIIN_DIR IS NULL OR VC_FORECAST_DIR IS NULL
   OR VC_LOGISTICS_DIR IS NULL OR VC_REPORTS_DIR IS NULL OR VC_SHIPPING_DIR IS NULL
   OR VC_TEMPLATE_DIR IS NULL;
GO

PRINT 'spike-inv-sites-paths.sql applied (idempotent): 7 path columns present + placeholders backfilled.';
GO
