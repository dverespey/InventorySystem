/* =====================================================================
   SIM_OrderSimulation  —  Order ("What to Order") day-by-day simulation
   Faithful T-SQL reproduction of the legacy Excel-OLE Start path.
   Sources: legacy-order-spec.md §3 (algorithm), option-a.md §2 (decision),
            build-spec.md §1-§5 (the pinned contract). All `spec §N` refer to
            legacy-order-spec.md; `bs §N` to build-spec.md.

   SPIKE-ONLY object on the docker "Inventory" DB. READ-ONLY (Start does no
   writes — spec §1). Wraps the existing START procs' LOGIC (bodies inlined as
   set-based reads — phase-1 wrap-the-proc; identical WHERE/status semantics,
   bs §2). The cross-DB AD_GetSpecialDate is STUBBED by SIM_SpecialDate_Fixture
   (bs §3); calendar-derived cells are fixture-backed.

   NO @site_id param (bs §1.1: grep = zero site columns drive the procs; site
   scoping is schema surgery, out of spike scope). A site_id column exists on
   INV_PARTS_STOCK_MST from a prior spike (F1) but the START procs do not filter
   on it, so adding the param here would give false isolation — deliberately omitted.

   Returns THREE result sets:
     A. day headers      (one row per fill column)   — bs §1.2
     B. grid rows        (SIZE_HEADER|PART|SIZE_FOOTER) — bs §1.3
     C. phased day cells (one row per grid-row x fill_pos) — bs §1.5
   8.1-COMPAT: multi-result-set SProc call works on 8.1.52; if the gateway
   binding chokes, split C into its own NQ (option-a §"IG81-COMPAT").
   ===================================================================== */
SET ANSI_NULLS ON;
GO
SET QUOTED_IDENTIFIER ON;
GO

IF OBJECT_ID('dbo.SIM_OrderSimulation','P') IS NOT NULL
    DROP PROCEDURE dbo.SIM_OrderSimulation;
GO

CREATE PROCEDURE dbo.SIM_OrderSimulation
    @LineName               varchar(10)  = '',        -- '' => all-lines branch (bs §1.1)
    @PartType               varchar(50),               -- TIRE/WHEEL/VALVE/FILM/MISC
    @SortType               varchar(50)  = 'PART NUMBER',
    @Today                  date         = NULL,        -- order/anchor date; NULL => getdate()
    @FillDays               int          = 23,          -- production day-columns (INI default 23, max 50)
    @ForecastUsageCompare   int          = 7,           -- usage-vs-forecast window (unused in cell calc; reserved)
    @UseFirstProductionDay  bit          = 0,           -- week-offset path; we map by VC_WEEK_DATE so offset is implicit
    @Section                char(1)      = NULL          -- IG81-COMPAT (bs §"IG81-COMPAT"): Ignition reads only the
                                                         -- FIRST result set, so the Perspective NQs call this proc once
                                                         -- per section: 'A' day-headers | 'B' grid rows | 'C' phased cells.
                                                         -- NULL (default) => emit ALL THREE sets (sqlcmd / parity tests).
AS
BEGIN
    SET NOCOUNT ON;

    IF @Today IS NULL SET @Today = CAST(getdate() AS date);
    IF @FillDays IS NULL OR @FillDays < 1 SET @FillDays = 23;
    IF @FillDays > 50 SET @FillDays = 50;                 -- INI max (spec §7, bs §1.1)

    DECLARE @lastCalOffset int = @FillDays * 3;           -- DateRangeCount window (spec §1)

    /* =================================================================
       STEP 0 + 1 — CALENDAR WALK (spec §3.1, bs §3)
       Build the (cal_offset x, fill_pos j, serial_date, weekday, day_kind)
       map. fDates[x] is keyed by calendar offset; only PRODUCTION days get a
       fill_pos. Holidays ('H'/other) consume x but NOT j. This is the two-
       index-space reconciliation the hazard-7 buckets depend on (bs §5).
       ================================================================= */
    DECLARE @cal TABLE (
        cal_offset  int        NOT NULL,
        serial_date date       NOT NULL,
        weekday     tinyint     NOT NULL,              -- 1=Mon..7=Sun
        status_abrv varchar(1)  NOT NULL,              -- 'O'|'X'|'H'|'' (none)
        is_prod     bit         NOT NULL,              -- counts as a fill day?
        day_kind    varchar(13) NOT NULL,
        fill_pos    int         NULL                   -- assigned to production days only
    );

    DECLARE @x int = 0, @j int = 0;
    DECLARE @d date, @wd tinyint, @st varchar(1), @prod bit, @kind varchar(13);

    WHILE (@x <= @lastCalOffset AND @j < @FillDays)
    BEGIN
        SET @d  = DATEADD(day, @x, @Today);
        -- ISO weekday 1=Mon..7=Sun, independent of @@DATEFIRST:
        SET @wd = ((DATEDIFF(day, '19000101', @d) % 7) + 1);   -- 1900-01-01 = Monday
        SET @st = ISNULL((SELECT TOP 1 VC_STATUS_ABRV
                          FROM dbo.SIM_SpecialDate_Fixture
                          WHERE DT_DATE = @d
                            AND (VC_LINE_NAME = '' OR VC_LINE_NAME = @LineName)), '');

        SET @prod = 0; SET @kind = 'NORMAL';
        IF @st = 'O'      BEGIN SET @prod = 1; SET @kind = 'OVERTIME';      END
        ELSE IF @st = 'X' BEGIN SET @prod = 1; SET @kind = 'NONPRODUCTION'; END
        ELSE IF @st = ''  BEGIN
            -- no special row: render iff Mon-Fri (weekday<6), weekends skip (spec §3.1)
            IF @wd < 6 BEGIN SET @prod = 1; SET @kind = 'NORMAL'; END
            ELSE        BEGIN SET @prod = 0; SET @kind = 'NORMAL'; END
        END
        -- else 'H' / any other matched status: skipped (prod=0)

        INSERT INTO @cal (cal_offset, serial_date, weekday, status_abrv, is_prod, day_kind, fill_pos)
        VALUES (@x, @d, @wd, @st, @prod,
                CASE WHEN @prod = 0 AND @st IN ('H','') THEN 'NORMAL' ELSE @kind END,
                CASE WHEN @prod = 1 THEN @j ELSE NULL END);

        IF @prod = 1 SET @j = @j + 1;
        SET @x = @x + 1;
    END

    -- fOvertimes[] = 1-based FILL positions of overtime days, in order (spec §3.1, bs §3.2)
    DECLARE @overtimes TABLE (ord int IDENTITY(1,1), fill_idx1 int);  -- fill_idx1 = fill_pos+1
    INSERT INTO @overtimes (fill_idx1)
        SELECT fill_pos + 1 FROM @cal
        WHERE is_prod = 1 AND day_kind = 'OVERTIME'
        ORDER BY fill_pos;

    -- @FirstFRS = yyyymmdd of the first rendered production date (spec §3.3) for the
    -- in-transit / open-order list filters (VC_FRS_DATE >= @FirstFRS).
    DECLARE @FirstFRS varchar(8) =
        (SELECT CONVERT(varchar(8), MIN(serial_date), 112) FROM @cal WHERE is_prod = 1);
    IF @FirstFRS IS NULL SET @FirstFRS = '00000000';

    /* =================================================================
       RESULT SET A — day headers (bs §1.2)
       Gated by @Section so Ignition can fetch this single set (IG81-COMPAT).
       ================================================================= */
    IF @Section IS NULL OR @Section = 'A'
    SELECT
        fill_pos,
        cal_offset,
        serial_date,
        weekday,
        day_kind
    FROM @cal
    WHERE is_prod = 1
    ORDER BY fill_pos;

    /* =================================================================
       STEP 2 — driving cursor: parts of @PartType ordered by size
       (SELECT_PartsStockInfoOrder logic, schema:7382). Faithful WHERE/ORDER.
       ================================================================= */
    DECLARE @parts TABLE (
        seq            int IDENTITY(1,1),
        part_number    varchar(12),
        kanban         varchar(5),
        supplier_code  varchar(5),
        supplier_name  varchar(60),
        size_code      varchar(6),
        size_name      varchar(60),
        renban_group   varchar(20),
        daily_usage    int,           -- H
        safety_days    int,           -- I
        safety_stock   int,           -- J = H*I
        total_inv      int,           -- IN_QTY (Total Inv)
        lot_qty        int,           -- O (IN_1LOTQTY)
        lead_time      int,           -- P (weekday-selected)
        lot_size_ord   bit,
        in_leadtime    int,
        in_lt_mon int, in_lt_tue int, in_lt_wed int, in_lt_thu int, in_lt_fri int, in_lt_sat int
    );

    INSERT INTO @parts (part_number,kanban,supplier_code,supplier_name,size_code,size_name,
                        renban_group,daily_usage,safety_days,safety_stock,total_inv,lot_qty,
                        lot_size_ord,in_leadtime,in_lt_mon,in_lt_tue,in_lt_wed,in_lt_thu,in_lt_fri,in_lt_sat)
    SELECT
        p.VC_PART_NUMBER,
        p.VC_KANBAN_NUMBER,
        m.VC_SUPPLIER_CODE,
        m.VC_SUPPLIER_NAME,
        s.VC_SIZE_CODE,
        s.VC_SIZE_NAME,
        ISNULL(g.VC_RENBAN_GROUP_CODE,''),
        ISNULL(s.IN_USAGE,0),
        ISNULL(s.IN_DAYS,0),
        ISNULL(s.IN_USAGE,0) * ISNULL(s.IN_DAYS,0),           -- J = H*I (spec §3.2)
        ISNULL(p.IN_QTY,0),
        ISNULL(p.IN_1LOTQTY,0),
        ISNULL(p.BIT_LOT_SIZE_ORDERS,0),
        ISNULL(p.IN_LEADTIME,0),
        ISNULL(p.IN_LEADTIME_MONDAY,0),  ISNULL(p.IN_LEADTIME_TUESDAY,0),
        ISNULL(p.IN_LEADTIME_WEDNESDAY,0),ISNULL(p.IN_LEADTIME_THURSDAY,0),
        ISNULL(p.IN_LEADTIME_FRIDAY,0),  ISNULL(p.IN_LEADTIME_SATURDAY,0)
    FROM INV_PARTS_STOCK_MST p
        JOIN INV_PART_TYPE_MST t ON p.IN_PART_TYPE_ID = t.IN_PART_TYPE_ID
        JOIN INV_SIZE_MST      s ON p.IN_SIZE_ID      = s.IN_SIZE_ID
        JOIN INV_SUPPLIER_MST  m ON p.IN_SUPPLIER_ID  = m.IN_SUPPLIER_ID
        LEFT OUTER JOIN INV_RENBAN_GROUP_MST g ON p.IN_RENBAN_ID = g.IN_RENBAN_ID
    WHERE p.IN_SIZE_ID IS NOT NULL
      AND t.VC_PART_TYPE = @PartType
      AND ( (@LineName = '' OR @LineName = ' ')        -- all-lines branch
            OR p.VC_LINE_NAME = @LineName )
    ORDER BY
        s.VC_SIZE_CODE,
        CASE WHEN @LineName='' AND @SortType='LINE, RENBAN'      THEN p.VC_LINE_NAME END,
        CASE WHEN @LineName='' AND @SortType='LINE, RENBAN'      THEN g.VC_RENBAN_GROUP_CODE END,
        CASE WHEN @LineName='' AND @SortType='LINE, PART NUMBER' THEN p.VC_LINE_NAME END,
        CASE WHEN @LineName='' AND @SortType='LINE, PART NUMBER' THEN p.VC_PART_NUMBER END,
        CASE WHEN @LineName='' AND @SortType='PART NUMBER'       THEN p.VC_PART_NUMBER END,
        p.VC_PART_NUMBER;   -- stable tiebreak

    /* ---- lead-time selection by @Today weekday (spec §3.2 / bs §1.4) ----
       Mon->_MONDAY ... Sat->_SATURDAY; if that weekday column is 0/NULL fall
       back to IN_LEADTIME; Sun(7)/else -> IN_LEADTIME. */
    DECLARE @todayWd tinyint = ((DATEDIFF(day,'19000101',@Today) % 7) + 1);
    UPDATE @parts
       SET lead_time =
           CASE
             WHEN @todayWd = 1 AND ISNULL(NULLIF(in_lt_mon,0),0) <> 0 THEN in_lt_mon
             WHEN @todayWd = 2 AND ISNULL(NULLIF(in_lt_tue,0),0) <> 0 THEN in_lt_tue
             WHEN @todayWd = 3 AND ISNULL(NULLIF(in_lt_wed,0),0) <> 0 THEN in_lt_wed
             WHEN @todayWd = 4 AND ISNULL(NULLIF(in_lt_thu,0),0) <> 0 THEN in_lt_thu
             WHEN @todayWd = 5 AND ISNULL(NULLIF(in_lt_fri,0),0) <> 0 THEN in_lt_fri
             WHEN @todayWd = 6 AND ISNULL(NULLIF(in_lt_sat,0),0) <> 0 THEN in_lt_sat
             ELSE in_leadtime
           END;

    /* ---- per-part inventory/order columns (bs §1.3) ----
       K/L/M/N + assembler, faithful to the four SELECT_Order* status filters
       (bs §2). M (in-transit) and the lists are FRS>=@FirstFRS scoped. */
    DECLARE @inv TABLE (
        part_number   varchar(12) PRIMARY KEY,
        assembler_qty int, plant_qty int, in_transit int, open_order int, wh_qty int,
        share_e int, tire_wheel_ratio decimal(9,4)
    );
    INSERT INTO @inv (part_number, assembler_qty, plant_qty, in_transit, open_order, wh_qty, share_e, tire_wheel_ratio)
    SELECT
        pr.part_number,
        a.aq, l.lq, m.mq, n.nq,
        pr.total_inv - (ISNULL(l.lq,0) + ISNULL(m.mq,0)),    -- K = TotalInv - (L+M) (bs §1.3)
        ISNULL(e.eq,0),                                      -- E = order history qty
        ISNULL(tw.ratio,0)
    FROM @parts pr
        OUTER APPLY (SELECT SUM(IN_QTY) aq FROM INV_OPEN_ORDER_INF o
                     WHERE o.VC_PART_NUMBER=pr.part_number
                       AND o.VC_STATUS_ASSEMBLER_YARD<>'' AND o.VC_STATUS_EMPTY_TRAILER='' AND o.VC_TERMINATED='') a
        OUTER APPLY (SELECT SUM(IN_QTY) lq FROM INV_OPEN_ORDER_INF o
                     WHERE o.VC_PART_NUMBER=pr.part_number
                       AND o.VC_STATUS_EMPTY_TRAILER='' AND o.VC_TERMINATED=''
                       AND (o.VC_WAREHOUSE<>'' OR o.VC_STATUS_PLANT_YARD<>'')) l
        OUTER APPLY (SELECT SUM(IN_QTY) mq FROM INV_OPEN_ORDER_INF o
                     WHERE o.VC_PART_NUMBER=pr.part_number
                       AND o.VC_ARRIVAL='' AND o.VC_STATUS_SUPPLIER_SHIPPING<>''
                       AND o.VC_STATUS_PLANT_YARD='' AND o.VC_STATUS_ASSEMBLER_YARD=''
                       AND o.VC_WAREHOUSE='' AND o.VC_STATUS_EMPTY_TRAILER='' AND o.VC_TERMINATED=''
                       AND o.VC_FRS_DATE>=@FirstFRS) m
        OUTER APPLY (SELECT SUM(IN_QTY) nq FROM INV_OPEN_ORDER_INF o
                     WHERE o.VC_PART_NUMBER=pr.part_number
                       AND (o.VC_STATUS_SUPPLIER_SHIPPING IS NULL OR o.VC_STATUS_SUPPLIER_SHIPPING='')) n
        OUTER APPLY (SELECT SUM(IN_QTY) eq FROM INV_OPEN_ORDER_INF o
                     WHERE o.VC_PART_NUMBER=pr.part_number) e          -- SELECT_OrderHistory (complete-history branch)
        OUTER APPLY (SELECT TOP 1
                        CASE WHEN fd.VC_TIRE_PART_NUMBER_CODE=pr.part_number
                             THEN fd.IN_TIRE_RATIO ELSE fd.IN_WHEEL_RATIO END / 100.0 ratio
                     FROM INV_FORECAST_DETAIL_INF fd
                     WHERE (fd.VC_TIRE_PART_NUMBER_CODE=pr.part_number OR fd.VC_WHEEL_PART_NUMBER_CODE=pr.part_number)
                       AND (fd.IN_TIRE_RATIO<>0 OR fd.IN_WHEEL_RATIO<>0)) tw;

    -- share E sum per size group, share_pct = E/(SUM E) (100% if singleton) (spec §3.2)
    DECLARE @share TABLE (part_number varchar(12) PRIMARY KEY, share_pct decimal(9,4));
    ;WITH grp AS (
        SELECT p.part_number, p.size_code, i.share_e,
               SUM(i.share_e) OVER (PARTITION BY p.size_code) grp_e,
               COUNT(*)       OVER (PARTITION BY p.size_code) grp_n
        FROM @parts p JOIN @inv i ON i.part_number=p.part_number
    )
    INSERT INTO @share (part_number, share_pct)
    SELECT part_number,
           CASE WHEN grp_n=1 THEN 100.0
                WHEN grp_e=0  THEN 0.0
                ELSE CAST(share_e AS decimal(18,6))*100.0/grp_e END
    FROM grp;

    /* =================================================================
       STEP 3 — added-leadtime + zone indices (spec §3.4, bs §1.3)
       addedleadtime: for each overtime day i (1-based ord): if
       (fOvertimes[i]-1) <= (leadtime + i) then INC else break (the BREAK
       makes it order-dependent — reproduce with a running stop flag).
       ================================================================= */
    DECLARE @lt TABLE (
        part_number varchar(12) PRIMARY KEY,
        lead_time int, added_leadtime int,
        leadtime_zone_end_index int, orderby_col_index int, frs_date date
    );
    ;WITH ot AS (   -- evaluate the break-loop per part: count overtime entries that
                    -- satisfy the condition with NO earlier failure (spec §3.4 break)
        SELECT pr.part_number, pr.lead_time,
               -- number of leading overtime entries (ordered) where (fill_idx1-1) <= (lead_time + ord)
               (SELECT COUNT(*) FROM @overtimes o
                WHERE o.ord <= ISNULL((SELECT MIN(o2.ord)-1 FROM @overtimes o2
                                       WHERE (o2.fill_idx1-1) > (pr.lead_time + o2.ord)),
                                      (SELECT MAX(o3.ord) FROM @overtimes o3))
                  AND (o.fill_idx1-1) <= (pr.lead_time + o.ord)) AS added
        FROM @parts pr
    )
    INSERT INTO @lt (part_number, lead_time, added_leadtime, leadtime_zone_end_index, orderby_col_index, frs_date)
    SELECT
        ot.part_number, ot.lead_time, ot.added,
        ot.lead_time - 1 + ot.added            AS leadtime_zone_end_index,   -- spec §3.4 step2
        ot.lead_time + ot.added                AS orderby_col_index,
        (SELECT serial_date FROM @cal WHERE is_prod=1 AND fill_pos = ot.lead_time + ot.added) AS frs_date
    FROM ot;

    /* =================================================================
       STEP 4 — phased forecast usage per (part, fill_pos)  (spec §3.2)
       fForecast[j] = INV_BREAKDOWN_FC_INF.IN_QTY{weekday} for the production
       day's week. VC_WEEK_DATE = Monday of the week (yyyymmdd); IN_QTYn n=ISO
       weekday. This subsumes the UseFirstProductionDay week-number offset.
       ================================================================= */
    DECLARE @usage TABLE (part_number varchar(12), fill_pos int, qty int,
                          PRIMARY KEY (part_number, fill_pos));
    INSERT INTO @usage (part_number, fill_pos, qty)
    SELECT pr.part_number, c.fill_pos,
           ISNULL(CASE c.weekday
                    WHEN 1 THEN b.IN_QTY1 WHEN 2 THEN b.IN_QTY2 WHEN 3 THEN b.IN_QTY3
                    WHEN 4 THEN b.IN_QTY4 WHEN 5 THEN b.IN_QTY5 WHEN 6 THEN b.IN_QTY6
                    WHEN 7 THEN b.IN_QTY7 END, 0)
    FROM @parts pr
    CROSS JOIN (SELECT cal_offset, fill_pos, serial_date, weekday FROM @cal WHERE is_prod=1) c
    OUTER APPLY (
        SELECT TOP 1 b.IN_QTY1,b.IN_QTY2,b.IN_QTY3,b.IN_QTY4,b.IN_QTY5,b.IN_QTY6,b.IN_QTY7
        FROM INV_BREAKDOWN_FC_INF b
        WHERE b.VC_PART_NUMBER = pr.part_number
          AND b.VC_WEEK_DATE = CONVERT(varchar(8),
                  DATEADD(day, -((DATEDIFF(day,'19000101',c.serial_date))%7), c.serial_date), 112)
    ) b;

    /* =================================================================
       STEP 5 — phased receipts buckets (in-transit + open-order)
       (spec §3.3 / bs §5). Bucket each list row's VC_FRS_DATE onto the
       fill_pos by MATCHING the serial_date through @cal (NOT datediff) — the
       hazard-7 reconciliation. Carries source_enum (IN_TRANSIT / OPEN_ORDER).
       ================================================================= */
    DECLARE @receipts TABLE (part_number varchar(12), fill_pos int, qty int, source_enum varchar(10),
                             PRIMARY KEY (part_number, fill_pos, source_enum));
    ;WITH frs AS (
        -- in-transit list rows (font 23) : same filter as SELECT_OrderInTransitList
        SELECT o.VC_PART_NUMBER pn, o.VC_FRS_DATE frs, o.IN_QTY qty, 'IN_TRANSIT' src
        FROM INV_OPEN_ORDER_INF o
        WHERE o.VC_ARRIVAL='' AND o.VC_STATUS_SUPPLIER_SHIPPING<>''
          AND o.VC_STATUS_PLANT_YARD='' AND o.VC_STATUS_ASSEMBLER_YARD=''
          AND o.VC_WAREHOUSE='' AND o.VC_STATUS_EMPTY_TRAILER='' AND o.VC_TERMINATED=''
          AND o.VC_FRS_DATE >= @FirstFRS
        UNION ALL
        -- open-order list rows (font 10) : SELECT_OrderOpenOrderList filter
        SELECT o.VC_PART_NUMBER, o.VC_FRS_DATE, o.IN_QTY, 'OPEN_ORDER'
        FROM INV_OPEN_ORDER_INF o
        WHERE (o.VC_STATUS_SUPPLIER_SHIPPING IS NULL OR o.VC_STATUS_SUPPLIER_SHIPPING='')
          AND o.VC_FRS_DATE >= @FirstFRS
    )
    INSERT INTO @receipts (part_number, fill_pos, qty, source_enum)
    SELECT f.pn, c.fill_pos, SUM(f.qty), f.src
    FROM frs f
        JOIN @parts pr ON pr.part_number = f.pn
        JOIN @cal c ON c.is_prod=1
                   -- match FRS date THROUGH @cal by yyyymmdd string equality (bs §5).
                   -- compat-level-100 safe (no TRY_CONVERT); VC_FRS_DATE is yyyymmdd.
                   AND f.frs = CONVERT(varchar(8), c.serial_date, 112)
    GROUP BY f.pn, c.fill_pos, f.src;

    /* =================================================================
       STEP 6 — PAB recursion (spec §3.5 / bs §1.6)
       day0.begin = total_inv - in_transit ; day_j.begin = day_(j-1).end
       day_j.end  = begin + receipts_j - usage_j   (receipts = in-transit+open)
       below_safety = (end < safety_stock = J7).
       Iterative per part (fill positions are few; <=50).
       ================================================================= */
    DECLARE @bal TABLE (part_number varchar(12), fill_pos int, balance int, below_safety bit,
                        PRIMARY KEY (part_number, fill_pos));

    DECLARE @pn varchar(12), @tot int, @intr int, @ss int;
    DECLARE part_cur CURSOR LOCAL FAST_FORWARD FOR
        SELECT pr.part_number, pr.total_inv, ISNULL(i.in_transit,0), pr.safety_stock
        FROM @parts pr LEFT JOIN @inv i ON i.part_number=pr.part_number;
    OPEN part_cur; FETCH NEXT FROM part_cur INTO @pn, @tot, @intr, @ss;
    WHILE @@FETCH_STATUS = 0
    BEGIN
        DECLARE @begin int = @tot - @intr;   -- day0 begin = TotalInv - InTransit
        DECLARE @fp int = 0, @maxfp int = (SELECT MAX(fill_pos) FROM @cal WHERE is_prod=1);
        WHILE @fp <= @maxfp
        BEGIN
            DECLARE @recv int = ISNULL((SELECT SUM(qty) FROM @receipts WHERE part_number=@pn AND fill_pos=@fp),0);
            DECLARE @use  int = ISNULL((SELECT qty FROM @usage WHERE part_number=@pn AND fill_pos=@fp),0);
            DECLARE @end int = @begin + @recv - @use;
            INSERT INTO @bal (part_number, fill_pos, balance, below_safety)
            VALUES (@pn, @fp, @end, CASE WHEN @end < @ss THEN 1 ELSE 0 END);
            SET @begin = @end;          -- next day's begin = this day's end
            SET @fp = @fp + 1;
        END
        FETCH NEXT FROM part_cur INTO @pn, @tot, @intr, @ss;
    END
    CLOSE part_cur; DEALLOCATE part_cur;

    /* =================================================================
       RESULT SET B — grid rows (bs §1.3)
       Per part: SIZE_HEADER (first of size group), PART, SIZE_FOOTER (last).
       SIZE_HEADER carries H/I/J; PART carries identity+inventory+order;
       SIZE_FOOTER is the End-Balance summary row. We emit a PART row per part
       plus one SIZE_HEADER and one SIZE_FOOTER per size group.
       IG81-COMPAT: IF-guarded so @Section suppresses the statement entirely
       (an empty WHERE would still emit an empty result set, and Ignition reads
       the FIRST result set — the gate must skip, not return zero rows).
       ================================================================= */
    IF @Section IS NULL OR @Section = 'B'
    BEGIN
    ;WITH grp AS (
        SELECT p.*, i.assembler_qty, i.plant_qty, i.in_transit, i.open_order, i.wh_qty,
               i.tire_wheel_ratio, sh.share_pct,
               lt.added_leadtime, lt.leadtime_zone_end_index, lt.orderby_col_index, lt.frs_date,
               ROW_NUMBER() OVER (PARTITION BY p.size_code ORDER BY p.seq) rn_first,
               ROW_NUMBER() OVER (PARTITION BY p.size_code ORDER BY p.seq DESC) rn_last,
               MIN(p.seq) OVER (PARTITION BY p.size_code) grp_seq
        FROM @parts p
        JOIN @inv   i  ON i.part_number=p.part_number
        JOIN @share sh ON sh.part_number=p.part_number
        JOIN @lt    lt ON lt.part_number=p.part_number
    )
    SELECT row_kind, sort_seq, sub_seq,
           size_code, size_name, supplier_code, supplier_name, part_number, kanban, renban_group,
           daily_usage, safety_days, safety_stock, total_inv, wh_qty, assembler_qty, plant_qty,
           in_transit, open_order, lot_qty, lead_time, qty_default, lot_default,
           share_pct, tire_wheel_ratio, added_leadtime, leadtime_zone_end_index, orderby_col_index, frs_date
    FROM (
        -- SIZE_HEADER
        SELECT 'SIZE_HEADER' row_kind, grp_seq sort_seq, 0 sub_seq,
               size_code, size_name, CAST(NULL AS varchar(5)) supplier_code, CAST(NULL AS varchar(60)) supplier_name,
               CAST(NULL AS varchar(12)) part_number, CAST(NULL AS varchar(5)) kanban, CAST(NULL AS varchar(20)) renban_group,
               daily_usage, safety_days, safety_stock,
               CAST(NULL AS int) total_inv, CAST(NULL AS int) wh_qty, CAST(NULL AS int) assembler_qty,
               CAST(NULL AS int) plant_qty, CAST(NULL AS int) in_transit, CAST(NULL AS int) open_order,
               CAST(NULL AS int) lot_qty, CAST(NULL AS int) lead_time,
               CAST(NULL AS int) qty_default, CAST(NULL AS int) lot_default,
               CAST(NULL AS decimal(9,4)) share_pct, CAST(NULL AS decimal(9,4)) tire_wheel_ratio,
               CAST(NULL AS int) added_leadtime, CAST(NULL AS int) leadtime_zone_end_index,
               CAST(NULL AS int) orderby_col_index, CAST(NULL AS date) frs_date
        FROM grp WHERE rn_first = 1

        UNION ALL
        -- PART rows
        SELECT 'PART', grp_seq, seq,
               size_code, size_name, supplier_code, supplier_name, part_number, kanban, renban_group,
               daily_usage, safety_days, safety_stock, total_inv, wh_qty, assembler_qty, plant_qty,
               in_transit, open_order, lot_qty, lead_time,
               lot_qty AS qty_default,            -- Q seeded = lot qty (spec §3.4); editable on screen
               lot_qty AS lot_default,            -- R default (spec §3.4/§6)
               share_pct, tire_wheel_ratio, added_leadtime, leadtime_zone_end_index, orderby_col_index, frs_date
        FROM grp

        UNION ALL
        -- SIZE_FOOTER (End-Balance summary anchor)
        SELECT 'SIZE_FOOTER', grp_seq, 999999,
               size_code, size_name, NULL,NULL, NULL,NULL,NULL,
               daily_usage, safety_days, safety_stock,
               NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,
               NULL,NULL,NULL,NULL,NULL,NULL
        FROM grp WHERE rn_last = 1
    ) z
    ORDER BY sort_seq, sub_seq;
    END

    /* =================================================================
       RESULT SET C — phased day cells (bs §1.5)
       One row per (part_number, fill_pos). value = forecast usage (the cell
       number the legacy writes in the forecast band); balance = PAB end;
       signal_enum (interior/zone, single, precedence); source_enum (font);
       below_safety flag (orthogonal). Receipts (in-transit/open) are emitted
       as their own value+source on their bucketed day.
       Signal precedence (high->low, bs §1.7): OVERTIME / NON_PRODUCTION
       (day_kind) override ORDER_BY override LEAD_TIME_ZONE override NONE.
       IG81-COMPAT: IF-guarded (see B) so @Section='C' emits ONLY this set.
       ================================================================= */
    IF @Section IS NULL OR @Section = 'C'
    SELECT
        u.part_number,
        u.fill_pos,
        -- value: receipts qty if a bucket lands here, else forecast usage
        -- (build-spec §1.5 single-scalar cell). REVIEW-FIX (Finding 2): the
        -- scalar alone hides the forecast draw on a receipt day, so we ALSO
        -- emit forecast_usage + receipt_qty as separate channels. Legacy renders
        -- the forecast band (font 23) and the qty cells (font 10) in SEPARATE
        -- rows of the size block (spec §3.2 vs §3.3); the renderer shows both.
        -- value parity vs the legacy sheet is OPEN until David's golden (see
        -- spike-results.md OPEN GAP G2).
        ISNULL(r.qty, u.qty)                       AS value,
        u.qty                                      AS forecast_usage,
        r.qty                                      AS receipt_qty,
        b.balance                                  AS balance,
        -- signal_enum (interior): day_kind override first, then order-by, then zone
        CASE
            WHEN c.day_kind='OVERTIME'      THEN 'OVERTIME'
            WHEN c.day_kind='NONPRODUCTION' THEN 'NON_PRODUCTION'
            WHEN u.fill_pos = lt.orderby_col_index               THEN 'ORDER_BY'
            WHEN u.fill_pos <= lt.leadtime_zone_end_index        THEN 'LEAD_TIME_ZONE'
            ELSE 'NONE'
        END                                        AS signal_enum,
        ISNULL(r.source_enum,'NONE')               AS source_enum,
        b.below_safety                             AS below_safety
    FROM @usage u
        JOIN @cal c ON c.is_prod=1 AND c.fill_pos=u.fill_pos
        JOIN @lt  lt ON lt.part_number=u.part_number
        LEFT JOIN @bal b ON b.part_number=u.part_number AND b.fill_pos=u.fill_pos
        LEFT JOIN (   -- collapse in-transit+open per cell; in-transit takes the font if both
            -- co-occur (bs §1.7 font precedence). MIN over {'IN_TRANSIT','OPEN_ORDER'}
            -- yields 'IN_TRANSIT' (lexically smaller) so in-transit wins, faithful to
            -- spec §3.3 (in-transit font 23 stamped over open-order font 10).
            SELECT part_number, fill_pos, SUM(qty) qty,
                   MIN(source_enum) source_enum
            FROM @receipts GROUP BY part_number, fill_pos
        ) r ON r.part_number=u.part_number AND r.fill_pos=u.fill_pos
    ORDER BY u.part_number, u.fill_pos;
END
GO
