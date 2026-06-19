-- AD_GetSpecialDate — the SHARED production calendar proc (David-provided body, 2026-06-19).
-- Lives in the **VehicleOrder** DB and is SHARED across all applications (Inventory, GALC, MES, …).
-- DECISION (Q9): the calendar STAYS shared in VehicleOrder — do NOT relocate it into Inventory
-- (contrast the `sites` table, which IS relocated into Inventory). The Ignition rebuild reads it
-- read-only via a Named Query / cross-DB proc wrapper. The status domain (O/X/holiday/production
-- statuses) is DATA-DRIVEN via the ProductionStatus table — read it, don't hardcode.
-- Tables (all VehicleOrder): SpecialDate(SpecialDateID, SpecialDate, SpecialDateName, LineID,
-- ProductionStatusID), Line(LineID, LineName), ProductionStatus(ProductionStatusID, ProductionStatus,
-- ProductionStatusAbv). Helper: VehicleOrder.dbo.F_ISO_WEEK_OF_YEAR (ISO week number).
-- Line resolution: a specific @LineName returns its line-specific dates UNION the all-lines
-- (s.LineID IS NULL) global dates; blank/space @LineName returns every date in the window.
-- NOTE for D1/Q4: the Inventory `sites`/site-line config must reference these shared VehicleOrder
-- Line **LineName** values (the calendar is keyed by LineName) — reconcile the mapping in M2.

USE [VehicleOrder]
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

ALTER PROCEDURE [dbo].[AD_GetSpecialDate]
    @BeginDate    DateTime,
    @EndDate    DateTime,
    @LineName    varchar(10)
AS
if @LineName = ''  or @lineName = ' '
begin
    SELECT     SpecialDateID,
            SpecialDate AS 'Date',
            SpecialDateName AS 'Description',
            CASE
                WHEN l.LineName is null THEN 'ALL LINES'
                ElSE l.LineName
            END AS 'Line Name',
            p.ProductionStatus AS 'Date Status',
            VehicleOrder.dbo.F_ISO_WEEK_OF_YEAR(SpecialDate)  'Week Number',
            DATEPART(DW, SpecialDate + @@DATEFIRST - 1)  'Day Number',
            p.ProductionStatusAbv AS 'Date Status Abrv'
    FROM SpecialDate s
    LEFT JOIN Line l
    ON l.lineID = s.LineID
    JOIN ProductionStatus p
    ON s.ProductionStatusID = p.ProductionStatusID
    WHERE SpecialDate between @BeginDate and @EndDate
    Order By SpecialDate
end
else
begin
    SELECT     SpecialDateID,
            SpecialDate AS 'Date',
            SpecialDateName AS 'Description',
            l.LineName AS 'Line Name',
            p.ProductionStatus AS 'Date Status',
            VehicleOrder.dbo.F_ISO_WEEK_OF_YEAR(SpecialDate)  'Week Number',
            DATEPART(DW, SpecialDate + @@DATEFIRST - 1)  'Day Number',
            p.ProductionStatusAbv AS 'Date Status Abrv'
    FROM SpecialDate s
    JOIN Line l
    ON l.lineID = s.LineID
    AND l.LineName = @LineName
    JOIN ProductionStatus p
    ON s.ProductionStatusID = p.ProductionStatusID
    WHERE SpecialDate between @BeginDate and @EndDate
    UNION
    SELECT     SpecialDateID,
            SpecialDate AS 'Date',
            SpecialDateName AS 'Description',
            CASE
                WHEN l.LineName is null THEN 'ALL LINES'
                ElSE l.LineName
            END AS 'Line Name',
            p.ProductionStatus AS 'Date Status',
            VehicleOrder.dbo.F_ISO_WEEK_OF_YEAR(SpecialDate)  'Week Number',
            DATEPART(DW, SpecialDate + @@DATEFIRST - 1)  'Day Number',
            p.ProductionStatusAbv AS 'Date Status Abrv'
    FROM SpecialDate s
    LEFT JOIN Line l
    ON l.lineID = s.LineID
    JOIN ProductionStatus p
    ON s.ProductionStatusID = p.ProductionStatusID
    WHERE SpecialDate between @BeginDate and @EndDate
    AND s.LineID is null
    Order By SpecialDate
end
GO
