# Module Analysis: Production Calendar

> Area: Production calendar (module-map §9). One file covers all three live units.
> Status: ✅ spec complete · Analyst: Claude (2026-06-16)

**Area:** Production calendar  **Status:** ✅  **Analyst:** Claude / 2026-06-16

This area is the **date-math source** that the forecast and order simulations depend on.
It has two distinct calendars that live in **two different databases**:

1. **First production day** — `INV_FIRST_PRODUCTION_DAY`, in the **InventorySystem DB**.
   Maintained by `FirstProductiionDay.pas`. Supplies the **weekoffset** used in the
   forecast/order week-number math.
2. **Special dates** (overtime `O` / non-production `X` / holiday `H`) — lives in the
   **ALC (Activity / GALC) database**, reached via the `ALC_Connection` and the `AD_*`
   procs. Maintained by `OvertimeHoliday.pas`. Read by `ForecastBreakdownF` and the
   `Order` calendar walk via `AD_GetSpecialDate` / `AD_GetSpecialDateWeek`.

A third unit, `ProductionDates.pas`, is a **date-range selector dialog** that does not
maintain a calendar at all — it derives available production dates from shipping/ASN
history and feeds report date params.

---

## 0. Live-vs-dead confirmation

| Unit | In `InventorySystem.dpr`? | Verdict |
|------|---------------------------|---------|
| `FirstProductiionDay.pas` (typo, double-i) | yes — `InventorySystem.dpr:42` | **LIVE** |
| `OvertimeHoliday.pas` | yes — `InventorySystem.dpr:26` | **LIVE** |
| `ProductionDates.pas` | yes — `InventorySystem.dpr:50` | **LIVE** |
| `HolidayOvertime.pas` | **not listed** | **DEAD** (2006 rewrite attempt) |

`HolidayOvertime.pas` is the dead duplicate. Beyond not being in the `.dpr`, it would not
even compile: its method bodies reference identifiers that are never declared on the form
class — `ProductionDateStatusComboBox` (`HolidayOvertime.pas:78`), `SpecialDatesDataSet`
(`:100`), `SpecialDatesCommand` (`:147`), `Event_DBGrid` (`:263`), and a bare `Refresh`
(`:160`). It also declares a free-floating `procedure Execute;` (`:67`) shadowing the
method. Confidence: high — it is an abandoned draft, do not spec it as shipping.

> Useful artifact only: the dead form names the intended ALC maintenance API —
> `AD_InsertSpecialDate` (`:148`), `AD_UpdateSpecialDate` (`:200`), `AD_DeleteSpecialDate`
> (`:237`). The **live** path (below) uses `AD_InsertSpecialDate` + `AD_DeleteSpecialDate`
> only — there is **no Update** in the live form.

---

## 1. Legacy surface

### 1a. First production day — `FirstProductiionDay.pas` + `.dfm` (~5.8 KB, simple)
- **Entry point:** `MainMenu.pas:2508` — `FirstProductionDay_Form := TFirstProductionDay_Form.Create(self); .Execute (ShowModal); .Free`.
- **Purpose:** master-data maintenance for the one-row-per-year table that records, for each
  production year, the calendar date of the first production day and the ISO-style week
  number that day falls in. `INT_FIRST_PRODUCTION_WEEK − 1` is the **weekoffset** that the
  forecast/order simulations subtract to renumber weeks relative to the plant's production
  start (see §4).

### 1b. Overtime / holiday (special dates) — `OvertimeHoliday.pas` + `.dfm` (~10 KB)
- **Entry point:** `MainMenu.pas:563` `OvertimeHoliday1Click` → `Create / Execute / Free`
  (menu item `OvertimeHoliday1`, `MainMenu.pas:70`).
- **Purpose:** maintain the plant's calendar of special production days per line — overtime
  (`O`), non-production (`X`), holiday (`H`) — plus a free-text description. **These rows
  live in the ALC database, not InventorySystem** (see §3/§4). This is the calendar the
  forecast day-spread and the order Christmas-shutdown walk read.

### 1c. Production-date selector dialog — `ProductionDates.pas` + `.dfm` (~6.5 KB)
- **Entry point:** instantiated inline by many `MainMenu` report handlers, e.g.
  `MainMenu.pas:2925, 3034, 3141, 3248, 3362, …`. Caller sets `INVOICE`/`ASN`/`Range`/`Month`
  properties, calls `Execute` (ShowModal), then reads back `ProductionDate` / `ToProductionDate`
  / `Line` / `Cancel`.
- **Purpose:** a modal picker that lists available production dates (or months, or a
  from/to range) for a chosen line, sourced from shipping/ASN history. It feeds the date
  params of report procs (PO report, daily shipping, ASN reports, invoice reports). It is
  **not** a calendar editor.

---

## 2. Data touched

| Table | Read | Write | DB | Notes |
|-------|:----:|:-----:|----|-------|
| `INV_FIRST_PRODUCTION_DAY` | ✔ | ✔ | Inventory | year → first-day date + first-week number; cols `VC_PRODUCTION_YEAR varchar(4) NOT NULL`, `DT_FIRST_PRODUCTION_DAY_OF_YEAR datetime NOT NULL`, `INT_FIRST_PRODUCTION_WEEK int NULL` (`Create Inventory.sql:1378`). **No PK / no unique index** declared. |
| `INV_SHIPPING_INF` | ✔ | | Inventory | `REPORT_AvailableProductionDates` reads distinct `VC_PRODUCTION_DATE` here for the non-EDI branch. |
| `INV_ASN_MST` | ✔ | | Inventory | same proc, EDI (ASN/INVOICE) branch. |
| ALC special-dates table (name unknown; `SpecialDateID`, `LineName`, `ProductionStatus`, date, description) | ✔ | ✔ | **ALC** | maintained + consumed via `AD_*` procs on `ALC_Connection`. Table DDL is **not in this repo** — body/schema unverified. |
| ALC `LINE` / `ProductionStatus` ref tables | ✔ | | **ALC** | combo population (`SelectSingleFieldALC` / `SelectMultiFieldALC`, `OvertimeHoliday.pas:105,111`). `ProductionStatus` filtered `SpecialDateUse = 1`. |

**Triggers:** none on `INV_FIRST_PRODUCTION_DAY` (confirmed absent from the schema; it is a
plain 3-column lookup with no IDENTITY and no FK). ALC-side triggers are out of repo
(unverified). `docs/triggers.sql` is OBSOLETE — schema is authoritative.

---

## 3. Stored procedures used

### InventorySystem DB (verified — bodies read in `Create Inventory.sql`)

| Proc | Op | Schema:line | Business rule (from body) |
|------|----|-------------|---------------------------|
| `SELECT_FirstProductionDay` | SELECT | 5982 | `@ProdYear varchar(4)=''`. If blank → all rows; else `WHERE vc_production_year=@ProdYear`. Returns 3 aliased cols: `'Production Year'`, `'First Day'`, **`'First Week Number'`** (= `INT_FIRST_PRODUCTION_WEEK`). This last column is the weekoffset source. |
| `INSERT_FirstProductionDay` | INSERT | 3032 | `(@ProdYear varchar(4), @ProdDate datetime, @WeekNumber int)`. Body is a bare positional `INSERT INV_FIRST_PRODUCTION_DAY VALUES(@ProdYear,@ProdDate,@WeekNumber)`. **No `IF EXISTS` / no dedup** — see hazard H1. |
| `DELETE_FirstProductionDay` | DELETE | 2162 | `@ProdYear varchar(4)`; `DELETE … WHERE vc_production_year=@ProdYear`. Deletes the whole year (and would delete all dup rows for that year). |
| `REPORT_AvailableProductionDates` | SELECT | 3991 | `(@Line varchar(50), @INVOICE int, @ASN int, @Month int=0)`. `@ASN=0 AND @INVOICE=0` → distinct `VC_PRODUCTION_DATE` from `INV_SHIPPING_INF WHERE VC_Line_Name=@Line` desc. Else from `INV_ASN_MST`; if `@Month=1`, distinct `substring(VC_PRODUCTION_DATE,1,6)` (yyyymm). **Note:** in the non-EDI branch `@Month` and `@Line`-vs-ASN are ignored; ASN/invoice branch ignores `@Line`. |

### ALC DB (external — bodies NOT in this repo; `AD_*` = Activity/GALC database, body unverified)

| Proc | Op | Called from | Param wiring (from call site) |
|------|----|-------------|-------------------------------|
| `AD_GetSpecialDates` | SELECT | `DataModule.GetOvertimeHolidayInfo` (`DataModule.pas:6585`) on `ALC_DataSet` | no params; grid result columns consumed positionally by `OvertimeHoliday.HoldDetails` (see §4). |
| `AD_InsertSpecialDate` | INSERT | `DataModule.InsertOvertimeHolidayInfo` (`DataModule.pas:6674`) on `ALC_StoredProc` | `@LineName=fLineName`, `@ProductionStatus=fEventType`, `@SpecialDateName=fEventDescription`, `@SpecialDate=fEventDate`. |
| `AD_DeleteSpecialDate` | DELETE | `DataModule.DeleteOvertimeHolidayInfo` (`DataModule.pas:6626`) on `ALC_StoredProc` | `@SpecialDateID=fRecordID`. |
| `AD_GetSpecialDate` | SELECT | order/forecast consumers (Order.pas:216, OrderFormCreateF.pas:719, MainMenu.pas:2114, DataModule.pas:3768/3895) | `@BeginDate`, `@EndDate`, `@LineName`. Returns the special dates in a window for the calendar walk. ⚠️ **Call sites disagree on the line param name** — `DataModule.pas:3768` sends `@LineName`, `:3895` sends `@Line`; one is wrong vs the real ALC proc signature (a param the proc ignores silently widens the window to all lines). Verify the live ALC proc signature. |
| `AD_GetSpecialDateWeek` | SELECT | `ForecastBreakdownF.pas:1387` | week-keyed variant used by the forecast day-spread; statuses `'H'`/`'X'`. |
| `AD_GetLines` | SELECT | `ProductionDates.Execute` (`ProductionDates.pas:173`) on `ALC_DataSet` | no params; populates line combo from ALC `LINE` master. |

> **Boundary resolution (the key question):** the **LIVE** special-date calendar lives in the **ALC
> database**, not InventorySystem — **zero `AD_*` procs in `Create Inventory.sql`** (grep confirmed), and
> every live consumer + the live OvertimeHoliday maintenance uses the ALC connection's `AD_*` procs.
> ⚠️ **Caveat (don't trip on this):** a legacy InventorySystem-side calendar DOES exist in the schema —
> table **`INV_OVERTIME_HOLIDAY`** (`Create Inventory.sql:1539`) + procs `INSERT_/DELETE_OvertimeHolidayInfo`,
> `SELECT_OvertimeHoliday`/`…Date`/`…Week`, `SELECT_CheckHoliday`, `SELECT_HolidayDate` — but it is **DEAD**:
> the only caller is `ForecastBreakDown.pas:599` (`SELECT_OvertimeHolidayWeek`), and `ForecastBreakDown.pas`
> is NOT in the `.dpr` (the live unit is `ForecastBreakdownF.pas`). So the rebuild must NOT wire to
> `INV_OVERTIME_HOLIDAY` thinking it's the source of truth — the live calendar is ALC. The OvertimeHoliday
> FORM both **maintains** (`AD_Insert/DeleteSpecialDate`, no update) and forecast/order **consume**
> (`AD_GetSpecialDate`/`…Week`) the ALC calendar. InventorySystem owns only the *first-production-day* table.

---

## 4. Business rules & edge cases

### First production day → forecast/order weekoffset (the load-bearing rule)
- The form computes defaults from the entered date: `Year_Edit = yyyy(date)` and
  `WeekNumber_Edit = WeekoftheYear(date)` (`FirstProductiionDay.pas:206-207`). The user can
  override the week number before insert.
- Persisted row = (`year`, `firstDate`, `weekNumber`). The consumers read it back via
  `SELECT_FirstProductionDay(@ProdYear)` and use **`'First Week Number' − 1` as `weekoffset`**:
  - `Order.pas:1091` `weekoffset := FieldByName('First Week Number').AsInteger-1;` (also 1162)
  - `MainMenu.pas:2287` `weekoffset := FieldByName('First Week Number').AsInteger-1;`
  - `ForecastBreakdownF.pas:1374` `WeekNumber := WeekNumber + FieldByName('First Week Number').AsInteger - 1;` (adds the offset back when labeling)
  - `OvertimeHoliday.pas:320` (the form's own "calculated week" display) subtracts `(First Week Number − 1)` from `WeekoftheYear`.
- **Confirms the Order R1 / Forecasting finding:** for 2026, week 2 → offset 1. The whole
  chain is gated on the INI flag `fiUseFirstProductionDay` (`DataModule.pas:114`, a
  `TCIniField` boolean). When **false**, consumers use raw `WeekoftheYear` and this table is
  never read. The OvertimeHoliday form hides its "calculated week" column when the flag is
  off (`OvertimeHoliday.pas:126-135`).
- **Year-boundary guard** (consumers + form): when `WeekoftheYear(date) >= 52` and the date's
  year ≠ current year, they query the *current* year's first-production row instead
  (`OvertimeHoliday.pas:311-314`) — a hand-rolled fix for the week-52/53 rollover.

### First-production-day form rules
- Insert validation (`HoldDetails`, `FirstProductiionDay.pas:163`): date must be ≥ `01/01/2002`,
  else `'Invalid Date'`. No upper bound, no uniqueness check in the form.
- **The Insert success/failure message is misleading (H1):** `Insert_ButtonClick`
  (`:121-134`) treats `InsertFirstProductionDayInfo = False` as "already exists", but the
  proc never checks existence and never returns false for a duplicate — it just inserts a
  second row for the same year. (False only ever comes back on an ADO error.) With no PK,
  duplicate years silently accumulate, and `SELECT_FirstProductionDay(@year)` would then
  return multiple rows → consumers read `FieldByName('First Week Number')` from the **first**
  row only. Confidence: high (proc body read).
- Delete is whole-year (`DELETE_FirstProductionDay` by year), so it also cleans up dups.

### Overtime/holiday form rules
- Grid columns consumed positionally in `HoldDetails(True)` (`OvertimeHoliday.pas:150-162`):
  `0=SpecialDateID(RecordID)`, `1=EventDate`, `2=Description`, `3=LineName`, `4=EventType`,
  `5=WeekNumber`, `6=DayNumber`, `7=EventTypeAbv`. The grid column **order is a hard contract**
  with the unverified `AD_GetSpecialDates` result set — any reordering there breaks the form.
- Insert validation (`:170-192`): date must be ≥ `now-1` (no past dates); a line must be
  chosen (combo index > 0); an event type must be chosen. `EventType` resolves to two values
  from the column-combo: abbreviation (`EventTypeAbv`, e.g. `O`/`X`/`H`) and full text.
- The form stores **`fEventType` (full ProductionStatus text)** into `@ProductionStatus` on
  insert (`DataModule.pas:6679`) — not the abbreviation. The ALC proc presumably maps it;
  unverified.
- "ALL LINES" pseudo-option is appended if >2 real lines (`OvertimeHoliday.pas:107-110`).
- **No Update** path — to change a special date the user deletes and re-inserts (the live
  form has Insert/Delete/Clear/Search/Close only; Search shows "Search not available",
  `:284-287`). The dead `HolidayOvertime.pas` had Update via `AD_UpdateSpecialDate`; that
  capability is not shipping.
- Day/week defaults: `DayNumber = DayofTheWeek(date)`, `WeekNumber = WeekoftheYear(date)`
  on date change (`:228-229`), plus the calculated (offset-adjusted) week if the flag is on.

### Production-date selector rules
- `Execute` first loads lines from ALC (`AD_GetLines`) then calls `GetProductionDates`
  (`REPORT_AvailableProductionDates`) and `ShowModal`.
- Month mode formats `yyyy/mm`; date mode `yyyy/mm/dd`; range mode shows the To combo.
- Range guard (`:219-220, 229-230`): the From index is forced to `To+1` to keep From after
  To in the desc-sorted list (newest first). Output strings are sliced back to `yyyymmdd` by
  the callers (e.g. `MainMenu.pas:2942`).
- **Connection/error mismatch (H3):** `GetProductionDates` runs on `Inv_DataSet` (Inventory),
  correct. But `Execute` runs `AD_GetLines` on `ALC_DataSet` (`:168-173`) yet checks
  `Inv_Connection.Errors` (`:176`) — wrong connection's error collection. An ALC failure
  here would not be detected by that guard. Confidence: high (call site read).

---

## 5. UI / UX notes
- All three are simple modal dialogs. First-production-day and overtime/holiday are classic
  master-CRUD grids (grid + detail panel + Insert/Delete/Clear/Close, Search disabled).
- Keep: the first-production-day year/date/week capture and the special-date line/type/date
  capture. Modernize: add a real uniqueness constraint + edit (update) on first-production-day;
  surface the ALC vs Inventory split explicitly; make the production-date picker a parameter
  control rather than a separate modal.
- The `O`/`X`/`H` status set and the `SpecialDateUse=1` filter on `ProductionStatus` come from
  ALC reference data — preserve as a lookup.

---

## 6. Target design (Ignition)

**Two calendars, two data sources — model them distinctly.**

### First production day (Inventory DB — owned here)
- **View:** a master-CRUD Perspective view (mirror `master-data/supplier` pattern): a table
  bound to a Named Query listing all years, plus a row editor (year / first date / week number).
- **Named Queries** (one per proc, organized to mirror the procs — per the team Named-Query
  CRUD practice):
  - `FirstProductionDay/Select` (wrap `SELECT_FirstProductionDay`, optional `@ProdYear`).
  - `FirstProductionDay/Insert` — **fix H1:** make it an UPSERT keyed on year (add the unique
    constraint), so the legacy "already exists" UX becomes real.
  - `FirstProductionDay/Delete` (by year).
- The **weekoffset** (`firstWeek − 1`) is shared date logic. Implement once in a Project-Library
  script (e.g. `calendar.weekOffset(year)`), gated by the migrated `useFirstProductionDay`
  site setting, and reuse from forecast/order — do not re-derive at each call site.

### Special dates (ALC DB — cross-DB dependency)
- **Cross-DB call-out:** these `AD_*` procs are in the **ALC/Activity (GALC-side) database**.
  In Ignition this is a **separate database connection** (the GALC modernization target). Two
  options, defer the choice to the architect:
  1. **Keep the cross-DB boundary:** a second Named-Query source pointed at the ALC connection,
    wrapping `AD_GetSpecialDates` / `AD_InsertSpecialDate` / `AD_DeleteSpecialDate` (+ the
    read-only `AD_GetSpecialDate` / `AD_GetSpecialDateWeek` consumed by forecast/order). This is
    the lowest-risk parity path while GALC is still on SQL Server.
  2. **Co-locate** the special-date calendar into the unified target DB once both InventorySystem
    and GALC are migrated. The order spike already proved a fixture stand-in
    (`SIM_SpecialDate_Fixture` stubbing `AD_GetSpecialDate`) — that fixture is the contract to
    reimplement against.
- **Add Update** in the rebuild (the dead form intended it; `AD_UpdateSpecialDate` exists ALC-side).

### Production-date picker
- A reusable Perspective parameter component (line dropdown + date/month/range mode) backed by
  a Named Query wrapping `REPORT_AvailableProductionDates` (Inventory) and `AD_GetLines` (ALC).
  Feeds report views' date params.

---

## 7. Migration plan for this module
- [ ] Stage 1 — wrap `SELECT_FirstProductionDay`, `REPORT_AvailableProductionDates`, and the ALC
  `AD_Get*` reads as read-only Named Queries; render both calendars; verify parity vs live DB.
- [ ] Stage 2 — enable writes: first-production-day Insert/Delete (Inventory); special-date
  Insert/Delete (ALC connection). Add the missing year-uniqueness + special-date Update.
- [ ] Stage 3 — reimplement weekoffset + day-spread logic in Project-Library scripts; decide
  ALC co-location vs cross-DB; retire the `SIM_SpecialDate_Fixture` stub for the real source.

---

## 8. Open questions for the user (domain expert)

1. **D1 multi-site / first-production-day.** `INV_FIRST_PRODUCTION_DAY` is keyed only by
   `VC_PRODUCTION_YEAR` (no site/line/plant column) — it is **single-site today**, a real D1
   gap. Different plants (CAMEX/NUMMI/TMMTX) start production on different dates. Should the
   rebuilt table be `(site_id, year)`? (D1 says per-site, fully isolated — confirm this applies
   and that the weekoffset is resolved per current site.)

2. **D1 multi-site / special dates.** ALC special dates are keyed per **line** (`@LineName`,
   plus an "ALL LINES" option). Is "line" already the per-site discriminator (one line per
   plant), or do we still need a `site_id`? This depends on how lines map to sites in the ALC
   DB — needs confirmation since the table is out of repo.

3. **ALC ownership after migration.** The special-date calendar lives in the GALC-side DB. In
   the unified target, who owns it — InventorySystem-side, GALC-side, or a shared calendar
   service? (Determines Stage-3 co-locate vs cross-DB.)

4. **First-week semantics across the 52/53 boundary.** Confirm the intended weekoffset behavior
   for dates the engine pushes into week ≥52 of a year whose first-production-row belongs to the
   next year — today there's a hand-rolled "use current year" guard. Is that the desired rule?

5. **Duplicate-year tolerance (H1).** Is it acceptable / expected that a year can have more than
   one first-production row today (no PK)? The rebuild should add a unique constraint and a real
   upsert; confirm no downstream relies on multiple rows per year.

6. **Special-date `@ProductionStatus` value.** Insert sends the full status text, not the
   abbreviation. Confirm the ALC proc keys on text (vs `O`/`X`/`H` abbrev) so the rebuild sends
   the right value.

---

## 9. Test cases / parity checks
- `SELECT_FirstProductionDay('2026')` returns the row whose `'First Week Number'` drives
  weekoffset = value−1; assert order/forecast week labeling matches old for the same DB.
- Insert a second 2026 row (legacy) → confirm two rows exist (documents H1); rebuild upsert →
  one row.
- `REPORT_AvailableProductionDates(@Line, INVOICE=0, ASN=0)` = distinct shipping dates desc for
  that line; with `ASN=1, Month=1` = distinct `yyyymm` from `INV_ASN_MST`.
- OvertimeHoliday round-trip: insert `(line, 'H'/holiday-text, date>now)` via `AD_InsertSpecialDate`,
  confirm it appears in `AD_GetSpecialDates` grid columns 0–7 in the expected order, then delete
  by `SpecialDateID`.
- Forecast day-spread: a date marked `'H'`/`'X'` in ALC special dates is excluded from
  production-day spreading (parity vs `AD_GetSpecialDateWeek`).

---

## Cross-cutting findings (P12 retry-recursion)

All wrong-target retries in this area are **already logged** in
`docs/analysis/cross-cutting/datamodule-retry-target-bugs.md` — citing, not re-filing:
- `DeleteOvertimeHolidayInfo` → `DeleteSupplierInfo` (entry #4, `DataModule.pas:6651`) —
  **crosses ALC→Inv connection**; shared `fRecordID` is a special-date id.
- `InsertFirstProductionDayInfo` → `InsertSizeInfo` (Inserts list, `DataModule.pas:6564`).
- `InsertOvertimeHolidayInfo` → `InsertSupplierInfo` (Inserts list, `DataModule.pas:6706`) —
  **crosses ALC→Inv connection**.
- `GetFirstProductionDayInfo` → `GetOvertimeHolidayInfo` (Gets list, `DataModule.pas:6472`) —
  on retry it switches from the Inventory SELECT to the ALC `AD_GetSpecialDates`, a wrong-DB
  fallback; also note `GetFirstProductionDayInfo`'s `finally` sets `Inv_DataSource.DataSet`
  even though it filled `Inv_DataSet` — benign here.

**New (not previously filed):**
- **N1 — wrong error-connection in `ProductionDates.Execute`:** `AD_GetLines` runs on
  `ALC_DataSet` but the failure guard checks `Inv_Connection.Errors`
  (`ProductionDates.pas:168-176`). An ALC failure is invisible to that check. This is a
  *connection/error-object* mismatch, not a retry-target bug, so it is not a P12 entry —
  flagging it as a NEW cross-cutting "ALC call checked against Inv_Connection" pattern (the
  same shape appears at `DataModule.pas:6590` where `AD_GetSpecialDates` on `ALC_DataSet`
  checks `Inv_Connection.Errors`, and `DataModule.pas:3777` where `AD_GetSpecialDate` checks
  `Inv_Connection.Errors`). Recommend a dedicated cross-cutting note if it recurs elsewhere.

---

## Hazards summary
- **H1** `INSERT_FirstProductionDay` has no dedup and the table has no PK → duplicate-year rows;
  the form's "already exists" message is fiction (proc never returns false for a dup).
  (`Create Inventory.sql:3032`, `FirstProductiionDay.pas:121-134`)
- **H2** Cross-DB dependency: special-date calendar lives in the ALC database; all `AD_*` procs
  are external (body unverified). Forecast/order correctness depends on a DB this repo doesn't
  contain. (`DataModule.pas:6585,6626,6674`)
- **H3 / N1** ALC calls guarded against `Inv_Connection.Errors` (wrong connection object) in
  `ProductionDates.pas:176`, `DataModule.pas:6590`, `DataModule.pas:3777`.
- **H4 (D1 gap)** First-production-day table is single-site (year-only key); special dates are
  per-line. Multi-plant calendars are not modeled today.
- **H5** OvertimeHoliday grid binds to `AD_GetSpecialDates` columns by ordinal 0–7 — a fragile
  positional contract with an out-of-repo result set.

### Confidence
- High: all three live `.pas` forms, the four InventorySystem procs (bodies read), the table
  DDL, the weekoffset wiring, and the dead-code verdict.
- **Body unverified (expected):** all `AD_*` procs and the ALC special-date table schema —
  they live in the ALC/GALC database, outside this repo. Behavior described from call sites
  only; do not treat ALC proc internals as confirmed.
- Per [[reference-schema-snapshot-vs-live]]: treat any `AD_*` signature here as "verify live";
  call-site param names are what the Delphi sends, not a proven proc signature.
