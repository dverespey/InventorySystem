/* =====================================================================
   Named-Query definitions — Order simulation (spike "Inventory" DB)
   Connection (Ignition):  Inventory_Spike
   These are the canonical DB-access-layer artifacts (one findable NQ per
   proc/section — memory `ignition-named-query-crud-practice`). They wrap the
   ONE proc SIM_OrderSimulation; a schema change is one edit here (or zero,
   if it lives in the proc).

   IG81-COMPAT (option-a §"IG81-COMPAT", build-spec §"IG81-COMPAT"):
   Ignition reads only the FIRST result set of a stored proc, so the proc
   takes an @Section selector and we expose ONE Named Query per section.
   All three call the same proc with identical params + a different @Section.

   STATUS NOTE: delivered here as the SQL source-of-truth + wired into the
   Perspective view via the proven system.db script-transform pattern (same as
   PartsStockMaster/List). Promoting these to formal on-disk NQ *resources*
   (data/projects/spike/ignition/named-query/...) is a mechanical paste into
   the Designer Named Query workspace — the on-disk NQ metadata-JSON schema is
   undocumented and was NOT hand-authored blind (would risk a non-deserializing
   resource + a wasted restart cycle). The SQL below IS the NQ body.

   Parameters (same for all three; types as Ignition NQ "Value" params):
     :LineName              String   default 'COROLLA'
     :PartType              String   (TIRE|WHEEL|VALVE|FILM|MISC)
     :SortType              String   default 'PART NUMBER'
     :Today                 Date     (gateway now() at simulate time)
     :FillDays              Integer  default 23  (max 50, clamped in proc)
     :ForecastUsageCompare  Integer  default 7
     :UseFirstProductionDay Integer  default 1   (bit; R1: golden client config = TRUE,
                                                  forecast resolved by ISO week# − weekoffset.
                                                  The OrderSpike view binding passes 1.)
   ===================================================================== */

/* ---- NQ: Order/sim/SIM_OrderSimulation_DayHeaders  (result set A) ---- */
EXEC dbo.SIM_OrderSimulation
     @LineName              = :LineName,
     @PartType              = :PartType,
     @SortType              = :SortType,
     @Today                 = :Today,
     @FillDays              = :FillDays,
     @ForecastUsageCompare  = :ForecastUsageCompare,
     @UseFirstProductionDay = :UseFirstProductionDay,
     @Section               = 'A';
-- returns: fill_pos, cal_offset, serial_date, weekday, day_kind

/* ---- NQ: Order/sim/SIM_OrderSimulation_GridRows    (result set B) ---- */
EXEC dbo.SIM_OrderSimulation
     @LineName              = :LineName,
     @PartType              = :PartType,
     @SortType              = :SortType,
     @Today                 = :Today,
     @FillDays              = :FillDays,
     @ForecastUsageCompare  = :ForecastUsageCompare,
     @UseFirstProductionDay = :UseFirstProductionDay,
     @Section               = 'B';
-- returns: row_kind, sort_seq, sub_seq, size_code, size_name, supplier_code,
--          supplier_name, part_number, kanban, renban_group, daily_usage,
--          safety_days, safety_stock, total_inv, wh_qty, assembler_qty,
--          plant_qty, in_transit, open_order, lot_qty, lead_time, qty_default,
--          lot_default, share_pct, tire_wheel_ratio, added_leadtime,
--          leadtime_zone_end_index, orderby_col_index, frs_date

/* ---- NQ: Order/sim/SIM_OrderSimulation_PhasedCells (result set C) ---- */
EXEC dbo.SIM_OrderSimulation
     @LineName              = :LineName,
     @PartType              = :PartType,
     @SortType              = :SortType,
     @Today                 = :Today,
     @FillDays              = :FillDays,
     @ForecastUsageCompare  = :ForecastUsageCompare,
     @UseFirstProductionDay = :UseFirstProductionDay,
     @Section               = 'C';
-- returns: part_number, fill_pos, value, forecast_usage, receipt_qty,
--          balance, signal_enum, source_enum, below_safety
--   value          = build-spec §1.5 single-scalar cell (receipt if bucketed else forecast)
--   forecast_usage = the day's forecast draw (always present) — REVIEW-FIX F2
--   receipt_qty    = in-transit/open-order qty bucketed here (NULL if none) — REVIEW-FIX F2
