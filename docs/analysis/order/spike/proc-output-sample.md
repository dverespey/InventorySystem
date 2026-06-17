# SIM_OrderSimulation — sample-part run output (spike "Inventory" DB)

Captured from the live docker spike DB (`mssql-spike`, login `sa`), proc applied from
`SIM_OrderSimulation.sql`, fixtures from `spike-fixtures.sql`. Anchor `@Today='2026-06-15'`
(the build-spec §3.3 calendar fixture anchor). Calendar-derived cells are **fixture-backed**
(AD_GetSpecialDate stubbed by `SIM_SpecialDate_Fixture`). All runs `@LineName='COROLLA'`
(only line present).

Run command (per part-type):
```
docker exec mssql-spike /opt/mssql-tools18/bin/sqlcmd -C -S localhost -U sa -P '<spike-sa>' \
  -d Inventory -s '|' -W -Q "EXEC SIM_OrderSimulation \
    @LineName='COROLLA', @PartType='<TIRE|WHEEL|VALVE>', @Today='2026-06-15', @FillDays=<N>"
```

Three result sets per call: **A** day-headers, **B** grid rows (SIZE_HEADER/PART/SIZE_FOOTER),
**C** phased cells `(part_number, fill_pos, value, balance, signal_enum, source_enum, below_safety)`.

---

## Result set A — day headers (FillDays=12, anchor 2026-06-15)  [HAZARD-7 PROOF]

`fill_pos | cal_offset | serial_date | weekday | day_kind`
```
0 |0 |2026-06-15 |1 |NORMAL
1 |1 |2026-06-16 |2 |NORMAL
2 |3 |2026-06-18 |4 |NONPRODUCTION   <-- cal_offset 3, NOT 2: 06-17 (H) consumed x=2, no column
3 |4 |2026-06-19 |5 |NORMAL
4 |7 |2026-06-22 |1 |NORMAL          <-- 06-20/21 weekend skipped (x=5,6)
5 |8 |2026-06-23 |2 |OVERTIME        <-- fOvertimes <- fill_idx1 6
6 |9 |2026-06-24 |3 |NORMAL
7 |11|2026-06-26 |5 |NORMAL          <-- 06-25 (H) skipped (x=10)
8 |14|2026-06-29 |1 |NORMAL
9 |15|2026-06-30 |2 |NORMAL
10|16|2026-07-01 |3 |NORMAL
11|17|2026-07-02 |4 |NORMAL
```
**fill_pos 2 = cal_offset 3** — the offset≠position divergence forced by the 06-17 holiday.
At FillDays=23 a 2nd overtime appears at fill_pos 12 = 2026-07-03 (fill_idx1 13).

---

## Case (a) singleton 4265202R6000 — TIRE / 15D (share = 100%)

B (PART row, key cols): supplier BRIDGESTONE, kanban 15BS, H=27 I=0 J=0, total_inv 12000,
wh_qty 12000, open_order 8800, lot_qty 800, lead_time 5, **share_pct 100.0000**,
tire_wheel_ratio 1.0000, added_leadtime 1, leadtime_zone_end_index 5, orderby_col_index 6,
frs_date 2026-06-24.

C (FillDays=8):
```
fp value balance signal          source below_safety
0  27    11973   LEAD_TIME_ZONE  NONE   0
1  27    11946   LEAD_TIME_ZONE  NONE   0
2  27    11919   NON_PRODUCTION  NONE   0
3  27    11892   LEAD_TIME_ZONE  NONE   0
4  28    11864   LEAD_TIME_ZONE  NONE   0
5  28    11836   OVERTIME        NONE   0
6  28    11808   ORDER_BY        NONE   0
7  28    11780   NONE            NONE   0
```
Singleton ⇒ no `=E/ΣE` split, share=100. J7=0 ⇒ never below safety. PAB draws down by forecast.

## Case (b) shared-size 18DL — 4265202S1000 (70%) + 4265202S2000 (30%)

B PART rows:
```
part         supplier   lead_time share_pct  ratio  orderby_col_index
4265202S1000 DUNLOP     5         70.1599    .7000  6
4265202S2000 MICHELIN   6         29.8401    .3000  7
```
**share_pct sums to 100.0000** across the size group (E/ΣE from order history, not the
tire ratio). Per-part weekday lead-time selection: S2000 picks 6 (its Tuesday/IN_LEADTIME),
shifting its order-by column to 7.

C (FillDays=8, abridged): S1000 fp0 value 1375 (open-order receipt) balance 24491;
S2000 fp2 value 1320 balance 24294 source OPEN_ORDER. No safety breach (J7=0).

## Case (c) in-transit + open-order — seeded on 4261102Q8000 (see case e)

The §4 seed flips the 2026-06-18 (X-day / fill_pos 2) row of 4261102Q8000 to
`VC_STATUS_SUPPLIER_SHIPPING='Y'` ⇒ qualifies as **in-transit**. At fill_pos 2 an in-transit
bucket (440) co-occurs with open-order buckets (400+480) → `value=1320`, `source_enum=IN_TRANSIT`
(in-transit wins font precedence per bs §1.7, fixed via MIN over source enums). See case (e).

## Case (d) safety breach 900804500600 — VALVE / RV (J7 = 922×10 = 9220)  [BELOW_SAFETY PROOF]

B SIZE_HEADER: H=922 I=10 **J=9220**. PART: total_inv 14600, lot_qty 1000, lead_time 15,
share 100, added_leadtime 2 (at FillDays=23), orderby_col_index 17, frs_date 2026-07-10.

C (FillDays=12):
```
fp value balance signal          source     below_safety
0  915   13685   LEAD_TIME_ZONE  NONE       0
1  898   12787   LEAD_TIME_ZONE  NONE       0
2  898   11889   NON_PRODUCTION  NONE       0
3  898   10991   LEAD_TIME_ZONE  NONE       0
4  945   10046   LEAD_TIME_ZONE  NONE       0
5  930    9116   OVERTIME        NONE       1   <-- 9116 < 9220 = J7 -> BELOW_SAFETY fires
6  5000  13186   LEAD_TIME_ZONE  OPEN_ORDER 0   <-- open-order receipt 5000 -> recovers above J7
7  930   12256   LEAD_TIME_ZONE  NONE       0
...
```
`below_safety = (PAB < J7)` fires exactly when projected balance dips below 9220 and clears
when a receipt lifts it back. At FillDays=23 a second below_safety run appears at fp17+ as
the balance trends down again (fp17 balance 8305 < 9220 → 1). lead_time 15 > FillDays-1 at
N=12 ⇒ order-by column off-window (frs_date NULL); at N=23 frs_date = 2026-07-10.

## Case (e) hazard-7 + busy phasing + renban — 4261102Q8000 — WHEEL / M1

B PART row: supplier CMWA, kanban M1, **renban_group CMWA** (in a group ⇒ ORDER-path renban
deferred), total_inv 44418, in_transit 440, open_order 369560, wh_qty 43978 (= 44418−440),
lot_qty 40, lead_time 5, share 100, added_leadtime 1, orderby_col_index 6, frs_date 2026-06-24.

C (FillDays=12):
```
fp value balance signal          source      below_safety
0  940   44164   LEAD_TIME_ZONE  OPEN_ORDER  0
1  1360  44775   LEAD_TIME_ZONE  OPEN_ORDER  0
2  1320  45346   NON_PRODUCTION  IN_TRANSIT  0   <-- 440 in-transit + 880 open on 06-18 (X day);
                                                      in-transit wins font; bucketed by fDates match
3  820   45417   LEAD_TIME_ZONE  OPEN_ORDER  0
4  400   45035   LEAD_TIME_ZONE  OPEN_ORDER  0
5  380   44638   OVERTIME        OPEN_ORDER  0   <-- 06-23 overtime-day bucket
6  360   44221   ORDER_BY        OPEN_ORDER  0
7  777   43444   NONE            NONE        0
...
```
The 06-18 FRS lands on **fill_pos 2** (cal_offset 3) by matching serial_date through @cal,
NOT by datediff — the hazard-7 reconciliation. PAB: day0 begin = 44418−440 = 43978;
fp0 = 43978 + 940 − 754(forecast) = 44164; recurrence holds across all days.

---

## Self-validation summary  (see returned message for PASS/FAIL)

- PAB recurrence (begin=prior end; end=begin+receipts−usage; day0 begin=TotalInv−InTransit): **PASS** (verified arithmetically on cases d & e).
- BELOW_SAFETY fires iff PAB < J7: **PASS** (case d, fp5 9116<9220 → 1, fp6 13186 → 0).
- share_pct sums to 100% per size group: **PASS** (18DL 70.1599+29.8401; SPARE 97.1460+2.8540; singletons 100).
- Hazard-7 fill_pos vs cal_offset (06-18 → fill_pos 2, not 3) and FRS bucket placement by fDates match: **PASS**.
- Added-leadtime break-loop (order-dependent cumulative): **PASS** (N=12 added=1; N=23 added=2 as 07-03 overtime enters window+leadtime window).
- in-transit font precedence when co-located with open-order: **PASS** (fp2 case e = IN_TRANSIT).

## NOT validatable without a golden (David's live legacy Excel export)
- Byte-for-byte parity of forecast usage values vs the live `OrderSimulation.xls` run (SC1: no Delphi/Excel here).
- Any template-baked conditional-format thresholds / palette RGBs (extraction gaps — option-a §8 #5/#6).
- The real AD_GetSpecialDate status domain/body (stubbed; calendar cells fixture-backed).
</content>
</invoke>
