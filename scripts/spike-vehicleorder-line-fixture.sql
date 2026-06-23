/* ============================================================================
   spike-vehicleorder-line-fixture.sql — VehicleOrder.LINE stub for the spike.
   ----------------------------------------------------------------------------
   ⚠️ SUPERSEDED (2026-06-23, P17.3): VehicleOrder is now a FULLY-RESTORED REAL
   database in mssql-spike (165 objects, the real `Line` table with 10 columns,
   real client data — READ-ONLY). This stub PREDATES that restore and would
   CORRUPT the real `Line` table: SQL Server is case-insensitive, so the
   OBJECT_ID('dbo.LINE') guard resolves to the real `Line`, the CREATE is skipped,
   but the DELETE + INSERT below would WIPE/overwrite the real single COROLLA row.
   DO NOT RUN. Kept only for history. The PartsStock Line dropdown now reads the
   real restored VehicleOrder.dbo.Line (1 line: COROLLA); multi-line coverage is a
   data-seed carry needing a real multi-line VehicleOrder restore, not this stub.
   The guard below hard-aborts if the real restored Line schema is detected.
   ----------------------------------------------------------------------------
   The Parts-Stock-Master "Line" dropdown reads, in PRODUCTION:
     SELECT DISTINCT LineName FROM LINE   (on ALC_Connection → Initial Catalog=VehicleOrder)
   i.e. a CROSS-DATABASE read of VehicleOrder.dbo.LINE (a flat assembly-line list;
   only LineName is needed — David 2026-06-17). VC_LINE_NAME is stored back as a
   plain string (no FK; default 'TUNDRA').

   The spike container only restores the `Inventory` DB, so VehicleOrder/LINE don't
   exist (same cross-DB gap the Order spike hit with the Activity AD_GetSpecialDate).
   This fixture creates a real second database `VehicleOrder` with `dbo.LINE` in the
   SAME mssql-spike instance, so the Parts view's lookup can reference it cross-DB via
   the 3-part name `VehicleOrder.dbo.LINE` over the existing Inventory_Spike connection
   — faithful to the production cross-DB read, headless-buildable.

   Run (sa):  export SA_PASS=...; docker exec -i mssql-spike /opt/mssql-tools18/bin/sqlcmd \
                -C -S localhost -U sa -P "$SA_PASS" -i /dev/stdin < this file
   ============================================================================ */
IF DB_ID('VehicleOrder') IS NULL CREATE DATABASE VehicleOrder;
GO
USE VehicleOrder;
GO
-- HARD GUARD (P17.3): refuse to run against the real restored VehicleOrder. If the real multi-column
-- `Line` schema is present, abort BEFORE the destructive DELETE/INSERT below would wipe real client data.
IF OBJECT_ID('dbo.Line','U') IS NOT NULL
   AND EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.Line') AND name = 'LineID')
BEGIN
    RAISERROR('SUPERSEDED: VehicleOrder.dbo.Line is the REAL restored table — this stub would corrupt it. Aborting.', 16, 1);
    SET NOEXEC ON;   -- skip the remaining destructive batches
END
GO
IF OBJECT_ID('dbo.LINE','U') IS NULL
    CREATE TABLE dbo.LINE (LineName varchar(10) NOT NULL);
GO
-- Seed the real assembly lines (LineName only — nothing else is needed here).
-- COROLLA is the spike's working line (matches INV_PARTS_STOCK_MST); TUNDRA is the
-- table default; the others give the dropdown a realistic multi-value set.
DELETE FROM dbo.LINE;
INSERT INTO dbo.LINE (LineName) VALUES ('COROLLA'), ('TUNDRA'), ('CAMRY'), ('TACOMA'), ('HIGHLANDER');
GO
-- Let the gateway's least-priv app login read it cross-DB (mirrors the Inventory grant).
IF SUSER_ID(N'ignition_spike') IS NOT NULL
BEGIN
    IF USER_ID(N'ignition_spike') IS NULL CREATE USER ignition_spike FOR LOGIN ignition_spike;
    ALTER ROLE db_datareader ADD MEMBER ignition_spike;
END
GO
SELECT 'VehicleOrder.LINE seeded' AS status, COUNT(*) AS lines FROM dbo.LINE;
GO
