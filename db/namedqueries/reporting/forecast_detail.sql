-- forecast_detail.sql  (Named Query: report_render/forecast_detail)
-- Wraps dbo.REPORT_ForecastDetail (M3 report R18, 1 run/yr). FAITHFUL rebuild.
-- params: NONE (the proc + the Delphi caller MainMenu.pas:3727 pass no params — whole detail table).
--
-- The live proc body (OBJECT_DEFINITION) is `SELECT * FROM inv_forecast_detail_inf ORDER BY ...`. A
-- `SELECT *` is brittle (column-order/presence depends on DDL — m3-report-proc-survey.md hazard SELECT*),
-- so this Named Query pins the EXACT 8 columns the Excel render reads (MainMenu.pas:3768-3801) in the
-- report's column order. Faithful = the same rows, same ORDER BY. READ-ONLY pure SELECT.
--
-- The 8 columns the Delphi render consumes (fieldbyname, MainMenu.pas:3768-3801):
--   VC_ASSY_PART_NUMBER_CODE, VC_EFFECTIVE_MONTH, VC_BROADCAST_CODE, VC_TIRE_PART_NUMBER_CODE,
--   IN_TIRE_RATIO, VC_WHEEL_PART_NUMBER_CODE, IN_WHEEL_RATIO, IN_RATIO
-- IG83-TODO: promote forecast_detail.sql to a Named Query; swap runPrepQuery -> runNamedQuery.

SELECT  VC_ASSY_PART_NUMBER_CODE,
        VC_EFFECTIVE_MONTH,
        VC_BROADCAST_CODE,
        VC_TIRE_PART_NUMBER_CODE,
        IN_TIRE_RATIO,
        VC_WHEEL_PART_NUMBER_CODE,
        IN_WHEEL_RATIO,
        IN_RATIO
FROM    inv_forecast_detail_inf
ORDER BY vc_assy_part_number_code, vc_effective_month, vc_broadcast_code
