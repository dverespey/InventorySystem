-- spike-stock-ledger-table.sql
-- INV_STOCK_LEDGER — the additive movements ledger that OWNS INV_PARTS_STOCK_MST.IN_QTY.
-- Design: docs/analysis/inventory-stock/IGNITION-stock-ledger-design.md §2.
-- Reproducible spike DDL (apply via sa into the restored Inventory sandbox DB).
--
-- IG83-TODO: at the Postgres phase TS_POSTED/VC_ADD become real datetime2; site_id
--            becomes an enforced FK; consider a computed-view projection of SUM(qty_delta)
--            instead of the materialized INV_PARTS_STOCK_MST.IN_QTY (design §1 open Q1).
-- IG81-COMPAT: pure T-SQL DDL — identical on SQL Server under 8.1.52 and 8.3 gateways.

SET NOCOUNT ON;
GO

IF OBJECT_ID('dbo.INV_STOCK_LEDGER', 'U') IS NOT NULL
    DROP TABLE dbo.INV_STOCK_LEDGER;
GO

CREATE TABLE dbo.INV_STOCK_LEDGER (
    IN_LEDGER_ID      int IDENTITY(1,1) NOT NULL,           -- surrogate PK, the movement's own key
    IN_PART_ID        int           NOT NULL,               -- the unified key (D2); FK -> INV_PARTS_STOCK_MST
    IN_QTY_CHANGE     int           NOT NULL,               -- signed delta: + adds on-hand, - removes
    IN_BALANCE_AFTER  int           NULL,                   -- running on-hand AFTER this post (materialized snapshot)
    VC_SOURCE_ENUM    varchar(24)   NOT NULL,               -- movement class: RECEIVING_SHIP / _ARRIVAL / _REVERSAL / REJECT / STOCKTAKING / SHIPPING
    IN_SOURCE_ROW_ID  int           NULL,                   -- surrogate id of the source row (order/reject/stk/shipping)
    VC_SOURCE_EVENT   varchar(100)  NOT NULL,               -- idempotency discriminator, e.g. 'SHIPPING:psh=12077:ins'
                                                            -- (100 not 40: amend keys carry ':upd:to=<effect>:v=<16-char stamp>' and the
                                                            --  longest ENUM is RECEIVING_REVERSAL -> ~64 chars; varchar(40) silently truncated,
                                                            --  dropping the per-edit stamp and re-introducing the amend-collision it guards against)
    site_id           int           NOT NULL DEFAULT (1),   -- D1; parallel-run is single-site (default 1)
    VC_REASON         varchar(300)  NULL,                   -- coded/free-text reason
    TS_POSTED         varchar(16)   NOT NULL,               -- 16-char yyyymmddHHMMSSff (P2; matches VC_ADD/VC_LAST_UPDATE)
    VC_ADD            varchar(16)   NOT NULL,               -- audit insert stamp (mirror INV_PART_QTY_INF.VC_ADD)
    CONSTRAINT PK_INV_STOCK_LEDGER PRIMARY KEY CLUSTERED (IN_LEDGER_ID),
    -- §4 idempotency backstop: one (part, event) can be posted at most once.
    CONSTRAINT UQ_INV_STOCK_LEDGER_EVENT UNIQUE (IN_PART_ID, VC_SOURCE_EVENT),
    -- D3: cannot delete a part that has ledger movements (RESTRICT / NO ACTION).
    CONSTRAINT FK_INV_STOCK_LEDGER_PART FOREIGN KEY (IN_PART_ID)
        REFERENCES dbo.INV_PARTS_STOCK_MST (IN_PART_ID)
);
GO

-- Covering index for the per-part SUM / balance read (design §2 indexes).
CREATE NONCLUSTERED INDEX IX_INV_STOCK_LEDGER_PART
    ON dbo.INV_STOCK_LEDGER (site_id, IN_PART_ID)
    INCLUDE (IN_QTY_CHANGE);
GO

PRINT 'INV_STOCK_LEDGER created.';
GO
