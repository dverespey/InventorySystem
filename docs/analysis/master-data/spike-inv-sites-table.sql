/* ============================================================================
   INV_SITES  --  Multi-site configuration master table
   ----------------------------------------------------------------------------
   First production-build schema artifact (David, 2026-06-19).

   PURPOSE
     This is the canonical multi-site configuration table for the Ignition
     rebuild. It RELOCATES into the Inventory database what the legacy Delphi app
     held PER-INSTALL in an INI file (read into TSiteInfo / SiteInfo.pas and the
     DataModule [INIT] / [DATAPURGE] CIniField properties). The legacy product is
     single-site (one INI per machine); decision D1 makes the rebuild multi-site
     with full per-site isolation, so the per-install INI config becomes one ROW
     per site here.

     Conceptually relocated from the VehicleOrder area into Inventory: Inventory
     owns the site master going forward.

   CANONICAL MULTI-SITE KEY (D2)
     IN_SITE_ID (surrogate INT IDENTITY) is the SOLE primary key and the canonical
     site FK column for the WHOLE rebuild. Every multi-site table gets an
     IN_SITE_ID FK referencing INV_SITES.IN_SITE_ID (D1). All server-side queries
     must be site-scoped via siteScopedQuery(); site id derives from the session,
     never a client parameter.

   THROWAWAY dbo.sites
     This table SUPERSEDES the throwaway spike table dbo.sites
     (site_id int / site_code varchar / site_name varchar; 2 rows
      1=MAS/TMMMS, 2=HERO/TMMTX; NOTHING FKs it). dbo.sites is LEFT IN PLACE for
     now (the spike Order screens may still read it) and should be DROPPED once
     those readers are repointed at INV_SITES.

     Likewise the throwaway scaffold column INV_PARTS_STOCK_MST.site_id (Check B,
     no FK constraint) is NOT touched here. It will be reconciled into an
     IN_SITE_ID FK referencing INV_SITES.IN_SITE_ID during M4 schema surgery.
     Do NOT alter that throwaway column in this artifact.

   SECRETS
     NO password / credential columns. Legacy SiteInfo.SitePassword and any EDI
     passwords belong in the Ignition gateway secret store (§4), not in this row.

   COLUMN -> LEGACY FIELD MAPPING (TSiteInfo / DataModule INI fields)
     VC_SITE_NAME                <- TSiteInfo.SiteName
     VC_SITE_ABBR                <- TSiteInfo.SiteAbbr      (EDI ISA sender ID, %-10s; e.g. MAS/HERO)
     VC_STREET                   <- TSiteInfo.SiteStreet
     VC_CITY                     <- TSiteInfo.SiteCity
     VC_STATE                    <- TSiteInfo.SiteState
     VC_COUNTRY                  <- TSiteInfo.SiteCountry
     VC_ZIP                      <- TSiteInfo.SiteZip
     VC_DUNS                     <- TSiteInfo.SiteDUNS      (Q7/Q11 inbound EDI routing key; INDEXED)
     VC_SUPPLIER_CODE            <- TSiteInfo.SiteSupplierCode
     VC_DOCK_CODE                <- TSiteInfo.SiteDockCode
     IN_EIN_SEQ                  <- TSiteInfo.SiteEIN       (Q4 per-site EIN sequence/counter)
     VC_EDI_MODE                 <- TSiteInfo.SiteEDIMode
     VC_SEP_SEGMENT              <- TSiteInfo.SiteSepSegment    (ISA/GS segment terminator)
     VC_SEP_ELEMENT              <- TSiteInfo.SiteSepElement    (ISA/GS element separator)
     VC_SEP_SUBELEMENT           <- TSiteInfo.SiteSepSubElement (ISA/GS sub-element separator)
     VC_TMM_NAME                 <- TSiteInfo.SiteTMMName
     VC_TMM_ABBR                 <- TSiteInfo.SiteTMMAbbr
     VC_TMM_DUNS                 <- TSiteInfo.SiteTMMDUNS
     IN_MAX_SEQUENCE             <- TSiteInfo.SiteMaxSequence
     BIT_ACCEPT_ANY_ORDER_ASN    <- TSiteInfo.AcceptAnyOrderASN (legacy integer flag)
     VC_DELIVERY_METHOD_CODE     <- TSiteInfo.SiteDeliveryMethodCode (EDI TD5, e.g. 'CV')
     IN_FILL_DAYS                <- DataModule fiFillDays           ([INIT]; CHECK <= 50)
     IN_FORECAST_USAGE_COMPARE   <- DataModule fiForecastUsageCompare ([INIT])
     BIT_USE_FIRST_PRODUCTION_DAY<- DataModule fiUseFirstProductionDay ([INIT])
     VC_FORECAST_IMPORT_MODE     <- Q11 per-site forecast import mode ('AUTO'|'MANUAL')
     VC_LAST_FORECAST_IMPORT     <- last forecast import stamp (16-char yyyymmddHHMMSSff)
     BIT_ENABLE_DATA_PURGE       <- DataModule fiEnableDataPurge    ([DATAPURGE])
     BIT_PROMPT_DATA_PURGE       <- DataModule fiPromptDataPurge    ([DATAPURGE])
     IN_DATA_RETENTION           <- DataModule fiDataRetention      ([DATAPURGE]; months;
                                       CHECK >= 12 per DataModule.pas:6890 which rejects < 12)
     VC_LAST_UPDATE / VC_ADD     <- INV_* house-style 16-char audit stamps (yyyymmddHHMMSSff)

   WIDTH CHOICES (verified against CreateInventory.sql live dump where analogous;
                  flagged where chosen):
     VC_SITE_NAME   varchar(50)  -- legacy VC_*_NAME range 25..50; chose 50 headroom  [CHOICE]
     VC_SITE_ABBR   varchar(10)  -- EDI ISA sender ID is %-10s (EDI810/856Object.pas)
     VC_STREET      varchar(50)  -- matches legacy VC_ADDRESS varchar(50)
     VC_CITY        varchar(50)  -- matches legacy VC_CITY varchar(50)
     VC_STATE       char(2)      -- 2-letter US/CA state code (spec)
     VC_COUNTRY     varchar(3)   -- ISO-ish short code (spec)                          [CHOICE]
     VC_ZIP         varchar(10)  -- matches legacy VC_ZIP varchar(10)
     VC_DUNS        varchar(13)  -- DUNS is 9 (or 9+4); 13 fits DUNS+4               [CHOICE]
     VC_SUPPLIER_CODE varchar(5) -- matches legacy VC_SUPPLIER_CODE varchar(5)
     VC_DOCK_CODE   varchar(10)  -- aligns to legacy VC_DOCK_NAME varchar(10)
     VC_EDI_MODE    varchar(10)  -- short mode token                                  [CHOICE]
     VC_SEP_*       char(1)      -- ISA/GS separators are single chars (EDI*Object.pas)
     VC_TMM_NAME    varchar(50)  -- mirrors VC_SITE_NAME                              [CHOICE]
     VC_TMM_ABBR    varchar(10)  -- mirrors VC_SITE_ABBR
     VC_TMM_DUNS    varchar(13)  -- mirrors VC_DUNS
     VC_DELIVERY_METHOD_CODE varchar(5) -- EDI TD5 method code, 2 chars typ.          [CHOICE]
     VC_FORECAST_IMPORT_MODE varchar(6) -- 'AUTO'/'MANUAL' (spec)
     VC_LAST_FORECAST_IMPORT varchar(16)-- 16-char audit-style stamp (spec)
     audit cols     varchar(16)  -- INV_* house style

   IDEMPOTENT: guarded DROP + CREATE so this re-runs cleanly on the spike.
   ============================================================================ */

SET NOCOUNT ON;
GO

IF OBJECT_ID('dbo.INV_SITES', 'U') IS NOT NULL
    DROP TABLE dbo.INV_SITES;
GO

CREATE TABLE dbo.INV_SITES
(
    IN_SITE_ID                   INT IDENTITY(1,1) NOT NULL,

    /* --- identity / address --- */
    VC_SITE_NAME                 VARCHAR(50)  NOT NULL,
    VC_SITE_ABBR                 VARCHAR(10)  NOT NULL,
    VC_STREET                    VARCHAR(50)  NULL,
    VC_CITY                      VARCHAR(50)  NULL,
    VC_STATE                     CHAR(2)      NULL,
    VC_COUNTRY                   VARCHAR(3)   NULL,
    VC_ZIP                       VARCHAR(10)  NULL,

    /* --- EDI / trading identity --- */
    VC_DUNS                      VARCHAR(13)  NULL,
    VC_SUPPLIER_CODE             VARCHAR(5)   NULL,
    VC_DOCK_CODE                 VARCHAR(10)  NULL,
    IN_EIN_SEQ                   INT          NOT NULL CONSTRAINT DF_INV_SITES_EIN_SEQ DEFAULT (0),
    VC_EDI_MODE                  VARCHAR(10)  NULL,
    VC_SEP_SEGMENT               CHAR(1)      NULL,
    VC_SEP_ELEMENT               CHAR(1)      NULL,
    VC_SEP_SUBELEMENT            CHAR(1)      NULL,
    VC_TMM_NAME                  VARCHAR(50)  NULL,
    VC_TMM_ABBR                  VARCHAR(10)  NULL,
    VC_TMM_DUNS                  VARCHAR(13)  NULL,
    IN_MAX_SEQUENCE              INT          NULL,
    BIT_ACCEPT_ANY_ORDER_ASN     BIT          NOT NULL CONSTRAINT DF_INV_SITES_ACCEPT_ASN DEFAULT (0),
    VC_DELIVERY_METHOD_CODE      VARCHAR(5)   NULL,

    /* --- order / forecast config ([INIT]) --- */
    IN_FILL_DAYS                 INT          NULL,
    IN_FORECAST_USAGE_COMPARE    INT          NULL,
    BIT_USE_FIRST_PRODUCTION_DAY BIT          NOT NULL CONSTRAINT DF_INV_SITES_USE_FPD DEFAULT (0),

    /* --- forecast import (Q11) --- */
    VC_FORECAST_IMPORT_MODE      VARCHAR(6)   NOT NULL CONSTRAINT DF_INV_SITES_FC_MODE DEFAULT ('AUTO'),
    VC_LAST_FORECAST_IMPORT      VARCHAR(16)  NULL,

    /* --- data purge (Q17, per-site) --- */
    BIT_ENABLE_DATA_PURGE        BIT          NOT NULL CONSTRAINT DF_INV_SITES_ENABLE_PURGE DEFAULT (0),
    BIT_PROMPT_DATA_PURGE        BIT          NOT NULL CONSTRAINT DF_INV_SITES_PROMPT_PURGE DEFAULT (1),
    IN_DATA_RETENTION            INT          NULL,

    /* --- audit --- */
    VC_LAST_UPDATE               VARCHAR(16)  NULL,
    VC_ADD                       VARCHAR(16)  NOT NULL,

    CONSTRAINT PK_INV_SITES PRIMARY KEY CLUSTERED (IN_SITE_ID),

    /* IN_FILL_DAYS capped at 50 ([INIT] business rule) */
    CONSTRAINT CK_INV_SITES_FILL_DAYS
        CHECK (IN_FILL_DAYS IS NULL OR IN_FILL_DAYS <= 50),

    /* retention must be >= 12 months (DataModule.pas:6890 rejects < 12) */
    CONSTRAINT CK_INV_SITES_DATA_RETENTION
        CHECK (IN_DATA_RETENTION IS NULL OR IN_DATA_RETENTION >= 12),

    /* forecast import mode is AUTO or MANUAL (Q11) */
    CONSTRAINT CK_INV_SITES_FC_IMPORT_MODE
        CHECK (VC_FORECAST_IMPORT_MODE IN ('AUTO', 'MANUAL'))
);
GO

/* DUNS is the inbound EDI routing key (Q7/Q11) -> index it */
CREATE NONCLUSTERED INDEX IX_INV_SITES_DUNS
    ON dbo.INV_SITES (VC_DUNS);
GO

/* ----------------------------------------------------------------------------
   SEED -- migrate the 2 throwaway dbo.sites rows.
   dbo.sites mapping observed on spike:
     site_code -> site ABBR (MAS / HERO);  site_name -> TMM plant (TMMMS / TMMTX)
   Non-identity/non-abbr values below are PER-SITE PLACEHOLDERS [PLACEHOLDER];
   real values are loaded from the legacy INI at cutover. IDENTITY columns are
   NOT seeded (let IDENTITY assign 1, 2 to preserve the throwaway ids).
   VC_ADD stamp is the build date in the 16-char yyyymmddHHMMSSff house format.
   ---------------------------------------------------------------------------- */
INSERT INTO dbo.INV_SITES
(
    VC_SITE_NAME, VC_SITE_ABBR, VC_STREET, VC_CITY, VC_STATE, VC_COUNTRY, VC_ZIP,
    VC_DUNS, VC_SUPPLIER_CODE, VC_DOCK_CODE, IN_EIN_SEQ, VC_EDI_MODE,
    VC_SEP_SEGMENT, VC_SEP_ELEMENT, VC_SEP_SUBELEMENT,
    VC_TMM_NAME, VC_TMM_ABBR, VC_TMM_DUNS, IN_MAX_SEQUENCE, BIT_ACCEPT_ANY_ORDER_ASN,
    VC_DELIVERY_METHOD_CODE,
    IN_FILL_DAYS, IN_FORECAST_USAGE_COMPARE, BIT_USE_FIRST_PRODUCTION_DAY,
    VC_FORECAST_IMPORT_MODE, VC_LAST_FORECAST_IMPORT,
    BIT_ENABLE_DATA_PURGE, BIT_PROMPT_DATA_PURGE, IN_DATA_RETENTION,
    VC_LAST_UPDATE, VC_ADD
)
VALUES
-- site 1: MAS / TMMMS (Tupelo MS)  -- placeholders pending legacy INI
-- VC_EDI_MODE is the X12 ISA15 usage indicator: a SINGLE char 'P'(roduction)/'T'(est) (legacy
-- Site.SiteEDIMode). It must be 1 char (EDI856Object.pas:161 emits it verbatim into the positional ISA);
-- the earlier 'PROD' placeholder would emit a malformed 4-char ISA15. Seeded 'P'.
('Tupelo MS',     'MAS',  NULL, 'Tupelo',     'MS', 'US', NULL,
 '000000001',     'MAS', 'D01', 0, 'P',
 '~', '*', '>',
 'Toyota Motor Mfg Mississippi',  'TMMMS', '000000011', 999, 0,
 'CV',
 30, 0, 0,
 'AUTO', NULL,
 0, 1, 12,
 NULL, '2026061900000000'),
-- site 2: HERO / TMMTX (San Antonio TX)  -- placeholders pending legacy INI
('San Antonio TX','HERO', NULL, 'San Antonio','TX', 'US', NULL,
 '000000002',     'HERO','D01', 0, 'P',
 '~', '*', '>',
 'Toyota Motor Mfg Texas',        'TMMTX', '000000012', 999, 0,
 'CV',
 30, 0, 0,
 'AUTO', NULL,
 0, 1, 12,
 NULL, '2026061900000000');
GO

/* ----------------------------------------------------------------------------
   NOTE ON dbo.sites: superseded by this table. Left in place intentionally so
   existing spike Order screens that still read dbo.sites keep working. DROP it
   once those readers are repointed at INV_SITES.
   ---------------------------------------------------------------------------- */
