-- AD_FRSPULL — the GALC broadcast vehicle-count pull feeding ASN creation
-- (David-provided body, 2026-06-19). Lives in the SHARED VehicleOrder DB (ALC side),
-- like AD_GetSpecialDate (Q9) — read-only cross-DB from the rebuild; NOT relocated.
--
-- ROLE (M1 ASN-creation keystone): the ratio fan-out (CalculateASNFRS, hand-written
-- Delphi DataModule.pas:5106) calls this per (line, production-date window) to get,
-- PER BROADCAST CODE (BC), how many vehicles were built and the implied order count.
-- Those drive the per-manifest ASN detail qty: round(VEHICLES * IN_ASSY_QTY * ratio/100),
-- and the "(No Ratio)" branch = Orders <= 5 (small-volume BC).
--
-- DECODE:
--   Returns rows of (BC char(3), ORDERS int, VEHICLES int, '' ) — 4th col always blank.
--   Query 1 (GROUND): BC = ModelYearCode + GROUNDWHEEL value + GROUNDTIRE value;
--     VEHICLES = COUNT(*), ORDERS = COUNT(*) * 4  (4 corners per vehicle).
--   Query 2 (SPARE): BC = ModelYearCode + SPARETIRE value;
--     VEHICLES = ORDERS = COUNT(*)  (1 spare per vehicle); excludes DataValue='M' (no spare).
--   UNION'd, ORDER BY BC.
--   Source tables (GALC build data, VehicleOrder DB): Vehicle, Model, Line, vehicledata,
--     DataItem (DataItemDescription in GROUNDTIRE/GROUNDWHEEL/SPARETIRE). NOT on the spike
--     (spike VehicleOrder is a LINE-only stub) -> needs the live VehicleOrder backup to test.
--
-- NOTE: @Start / @Last are declared but UNUSED in the body — only @begindate/@enddate +
--   @LineName filter. (Corrects the M1 spec's inference that S/E were the ASN sequence range;
--   the FRS pull is purely date+line scoped.)
--
-- BUILD: the rebuild CALLS this (a read on the shared ALC DB), as it does AD_GetSpecialDate;
--   the fan-out LOGIC around it (branch + round + BC->parts via SELECT_ForecastDetailBCASN +
--   manifest gen) is reimplemented as pure Jython (computeAsnDetails) + unit-tested with
--   synthetic FRS/forecast rows; live parity re-runs against the VehicleOrder backup.

USE [VehicleOrder]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[AD_FRSPULL]
    @begindate    datetime,
    @enddate    datetime,
    @Start        int,
    @Last        int,
    @LineName    varchar(50)
AS

SELECT convert(char(3),m.ModelYearCode+vd2.DataValue+vd1.DataValue) AS BC,
                       COUNT(*)*4 AS ORDERS,COUNT(*) AS VEHICLES,''
    FROM Vehicle v with(NOLOCK)
    JOIN Model m
    ON v.ModelID = m.ModelID
    JOIN Line l with(NOLOCK)
        on v.LineID = l.LineID
    JOIN vehicledata vd1 with(NOLOCK)
        on v.VehicleID = vd1.VehicleID
    JOIN DataItem i1 with(NOLOCK)
        on vd1.DataItemID = i1.DataItemID
        AND i1.DataItemDescription = 'GROUNDTIRE'
    JOIN vehicledata vd2 with(NOLOCK)
        on v.VehicleID = vd2.VehicleID
    JOIN DataItem i2 with(NOLOCK)
        on vd2.DataItemID = i2.DataItemID
        AND i2.DataItemDescription = 'GROUNDWHEEL'
    WHERE v.DateCreated >=  @begindate
    and v.DateCreated <= @enddate
    AND l.LineName =@LineName
    GROUP BY  convert(char(3),m.ModelYearCode+vd2.DataValue+vd1.DataValue)
    --ORDER By BC
UNION
    SELECT convert(char(3),m.ModelYearCode+vd1.DataValue) AS BC,
                       COUNT(*) AS ORDERS, COUNT(*) AS VEHICLES,''
    FROM Vehicle v with(NOLOCK)
    JOIN Model m
    ON v.ModelID = m.ModelID
    JOIN Line l with(NOLOCK)
        on v.LineID = l.LineID
    JOIN vehicledata vd1 with(NOLOCK)
        on v.VehicleID = vd1.VehicleID
    JOIN DataItem i1 with(NOLOCK)
        on vd1.DataItemID = i1.DataItemID
        AND i1.DataItemDescription = 'SPARETIRE'
    WHERE v.DateCreated >= @begindate
    and v.DateCreated <= @enddate
    AND l.LineName = @LineName
    AND vd1.DataValue <> 'M' --Quick fix for no spare broadcast
    GROUP BY  convert(char(3),m.ModelYearCode+vd1.DataValue)
    ORDER By BC
GO
