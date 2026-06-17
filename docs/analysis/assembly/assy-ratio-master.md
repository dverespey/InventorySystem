# Module Analysis: Assembly Ratio Master (broadcast-code → tire/wheel parts)

**Area:** Assembly  **Status:** ✅ spec complete  **Analyst:** Claude / 2026-06-16

> **The broadcast-code → tire/wheel-part explosion master.** `AssyRatioMaster` is the CRUD editor
> over **`INV_ASSY_RATIO_MST`** — one row per Toyota assembly **broadcast code** (`VC_BROADCAST_CODE`,
> a 2-char code), mapping it to an assembly part number, up to **three tire part numbers + three
> tire ratios**, up to **three wheel part numbers + three wheel ratios**, tire/wheel quantities, and
> a spare tire/wheel. Tire ratios must sum to 100; wheel ratios must sum to 100. This is the
> data layer **behind** the per-build tire/wheel split, *parallel to* the forecast-detail BOM, **not
> the same table** (see §4.5 boundary). This single file covers both Assembly forms because the
> second one is dead.

**Combined-file note (matches established split style):** this file documents `AssyRatioMaster`
(live) and folds the **dead, stub** `BCRatioMaster` (§1.2). They are tightly coupled (BCRatioMaster
is a never-finished re-design of this exact screen and even `uses AssyRatioMaster`).

---

## 1. Legacy surface

### 1.1 `AssyRatioMaster` — LIVE in dpr, but UNREACHABLE in the running app
- **Form:** `AssyRatioMaster.pas` (772 lines / ~24 KB) + `AssyRatioMaster.dfm` (~13 KB).
  `TAssyRatioMaster_Form`; author Aaron Huge, 2002-10-25; revised 2002-11-14 to support
  "two tire part numbers and two ratios" (the screen actually carries **three** of each — the grid
  `Fields[0..19]` and procs confirm 3 tire + 3 wheel slots).
- **Registered live:** `InventorySystem.dpr:17` — `AssyRatioMaster in 'AssyRatioMaster.pas'`. It
  **is** compiled into the product.
- **Entry point — DEAD UI PATH (high-confidence finding):** the only constructor is
  `MasterMaint.pas:119` `AssyRatioMaster_ButtonClick`, behind `AssyRatioMaster_Button`. But
  `MasterMaint.Execute` runs, **unconditionally and after the conditional block**,
  `AssyRatioMaster_Button.Visible := FALSE; // not used yet` (`MasterMaint.pas:78`; a second
  identical hide sits inside the `GenerateEDI` branch at `MasterMaint.pas:74`). **The button is
  always hidden**, so an operator cannot open this form. It is *live code on a dead path* — compiled,
  fully functional, but not reachable through the menu. Treat its behavior as the spec for the data
  it would maintain, but flag that the current production system does **not** edit this table through
  the UI (the rows are maintained by some other means — see §4.6 / §8).
  *Confidence: high (read both hides in `MasterMaint.Execute`).*
- **Purpose:** Maintain, per broadcast code, the set of tire/wheel part numbers and the percentage
  split among them, plus quantities and spares — i.e. how one assembly's wheels-and-tires order
  explodes across multiple component part numbers.

### 1.2 `BCRatioMaster` — DEAD CODE + UNFINISHED STUB (do not spec as shipping)
- **Files:** `BCRatioMaster.pas` (135 lines / 3.6 KB) + `.dfm` (8 KB). `TBCRatioMaster_Form`.
- **NOT in `InventorySystem.dpr`** (verified: no match for `BCRatio`/`BC_Ratio` in the dpr) → **dead
  code, never compiled into the product.** Its only referrer is itself; `MasterMaint` does not
  mention it.
- **It is also a stub even as source:** `SetDetailBoxes`, `HoldDetails`, and `SearchGrid` are empty
  bodies (`BCRatioMaster.pas:100–111`); `SearchGrid` just `Result := True`. It has no
  insert/update/delete handlers at all — only `FormCreate` + `Execute`.
- It references `Data_Module.GetBCRatioInfo` (`:89`) and a second dataset `Inv_Field_DataSet` (`:91`),
  implying an intended **two-grid** redesign (broadcast→ratio header + part-number→ratio detail,
  i.e. an N-part generalization of the fixed 3-tire/3-wheel layout). That redesign was never finished.
  `GetBCRatioInfo` **does exist** in `DataModule.pas:566/2701` but its retry branch calls
  `GetAssyRatioInfo` (a documented LOW P12 mis-target — see §7).
- **Conclusion:** ignore `BCRatioMaster` for migration except as evidence of intended direction
  (generalize the fixed tire1/2/3 + wheel1/2/3 columns into a child ratio table — see §6).

---

## 2. Data touched

| Table | Read | Write | Notes |
|-------|:----:|:-----:|-------|
| `INV_ASSY_RATIO_MST` | ✅ | ✅ | The master. CRUD via the 4 procs in §3. **Effective key = `VC_BROADCAST_CODE`** (all of SELECT-dup-check/UPDATE/DELETE filter on it). |
| `INV_PARTS_STOCK_MST` | ✅ | — | Combo-box sources: tire part numbers (`WHERE VC_TIRE_WHEEL='T'`), wheel part numbers (`='W'`), via `GetParts`/`SelectSingleField` (`AssyRatioMaster.pas:441–446`). |
| `INV_FORECAST_DETAIL_INF` | ✅ | — | Combo source for the **AssyCode** dropdown: `SelectSingleField('INV_FORECAST_DETAIL_INF', 'VC_ASSY_PART_NUMBER_CODE', …)` (`AssyRatioMaster.pas:447`). Read-only; this is where the screen's notion of valid assembly codes comes from — **a cross-link to the forecast-detail BOM** (see §4.5). |

**Columns of `INV_ASSY_RATIO_MST`** (DDL, `DB Schema/Create Inventory.sql`):
`VC_ASSY_PART_NUMBER_CODE varchar(12) NOT NULL`, `VC_ASSY_NAME varchar(50) NULL`,
`VC_BROADCAST_CODE varchar(2) NOT NULL`, `IN_TIRE_QTY int NULL`,
`VC_TIRE_PART_NUMBER{1,2,3}_CODE varchar(12)`, `IN_TIRE_RATIO1 int NOT NULL`, `IN_TIRE_RATIO{2,3} int NULL`,
`VC_WHEEL_PART_NUMBER{1,2,3}_CODE varchar(12)`, `IN_WHEEL_RATIO1 int NOT NULL`, `IN_WHEEL_RATIO{2,3} int NULL`,
`IN_WHEEL_QTY int NULL`, `IN_SPARE_TIRE_QTY int NULL`,
`VC_SPARE_TIRE_PART_NUMBER_CODE varchar(12)`, `VC_SPARE_WHEEL_PART_NUMBER_CODE varchar(12)`,
`VC_BLANKET_PO varchar(8) NOT NULL`, `MO_ASSEMBLY_COST money NOT NULL`,
`VC_LAST_UPDATE varchar(16) NULL`, `VC_ADD varchar(16) NULL`.

**No declared PRIMARY KEY / UNIQUE / DEFAULT / FK constraints** appear for this table in the schema
(grep for `CONSTRAINT|PRIMARY|DEFAULT|ALTER TABLE` against `INV_ASSY_RATIO_MST` returned nothing).
Uniqueness of broadcast code is enforced **only** by the app-side `If RecordCount = 0` dup-check
inside `InsertAssyRatioInfo`. *Confidence: high.*

**Triggers on `INV_ASSY_RATIO_MST`:**
- `UPDATE_AssyRatioMst` (`FOR UPDATE`) — **no-op**: `if @numrows>0 begin SET NOCOUNT ON; print('UPDATE AssyRatioMst') end`. Enforces nothing. *Confidence: high (read body).*

**Triggers on OTHER tables that WRITE `INV_ASSY_RATIO_MST`** (part-master rename/delete cascade — this
is the only real referential glue, since there are no FKs):
- `DELETE_PartNumber` (on `INV_PARTS_STOCK_MST`, `FOR DELETE`): when a part is deleted, **blanks**
  `VC_TIRE_PART_NUMBER1_CODE`, `…2_CODE`, `VC_WHEEL_PART_NUMBER1_CODE`, `…2_CODE` where they matched
  the deleted part number (`SET = ''`). **Gap: it does NOT touch the `_3_CODE` slots or the spare
  tire/wheel codes** — deleting a part that is referenced only as tire3/wheel3/spare leaves a dangling
  code in the ratio master. *Confidence: high (read body).*
- `UPDATE_PartNumber` (on `INV_PARTS_STOCK_MST`, `FOR UPDATE`, only when `@numrows = 1` **and**
  `update(vc_part_number)`): propagates a part-number **rename** into
  `VC_TIRE_PART_NUMBER{1,2}_CODE` and `VC_WHEEL_PART_NUMBER{1,2}_CODE` (set to the new code). **Same
  gap: `_3_CODE` slots and spares are not renamed**, and the cascade only fires for a single-row
  rename. *Confidence: high (read body).*

(Both also cascade into `INV_FORECAST_DETAIL_INF.VC_TIRE_PART_NUMBER_CODE` / `VC_WHEEL_PART_NUMBER_CODE`
— relevant to the boundary in §4.5.)

---

## 3. Stored procedures used

All four CRUD procs wrap `INV_ASSY_RATIO_MST` and are invoked via `DataModule` methods of the same
name. Bodies read from `DB Schema/Create Inventory.sql` (authoritative; `docs/triggers.sql` obsolete).

| Proc | Op | Business rule (from body) |
|------|----|---------------------------|
| `SELECT_AssyRatioInfo` | SELECT | `@BroadCode varchar(12)=''`. If empty → all rows; else `WHERE VC_BROADCAST_CODE = @BroadCode`. Returns the 20 display columns with **friendly aliases** (`'Broadcast Code'`, `'ASSY Code'`, … `'Spare Wheel Parts Code'`) consumed positionally by the grid (`HoldDetails` reads `Fields[0..19]`). `ORDER BY VC_BROADCAST_CODE`. **Also (mis)used as the dup-check for `INSERT_SizeInfo` — see §7 / D8 Bug 1.** *Body verified.* |
| `INSERT_AssyRatioInfo` | INSERT | 21 params (assy code/name, broad code, tire qty, tire1/2/3 + ratios, wheel qty, wheel1/2/3 + ratios, spare tire qty + part, spare wheel part). Computes `@AddDate` and writes it to `VC_ADD`. **Inserts only 21 columns — does NOT set `VC_BLANKET_PO` or `MO_ASSEMBLY_COST`, both `NOT NULL` with no schema default, nor `VC_LAST_UPDATE`.** *Body verified.* |
| `UPDATE_AssyRatioInfo` | UPDATE | Same 21 value params **+ `@BroadCodePrev varchar(2)`**. `UPDATE … WHERE VC_BROADCAST_CODE = @BroadCodePrev`. Writes `VC_LAST_UPDATE = @Updated` (computed). **Does NOT touch `VC_ADD`, `VC_BLANKET_PO`, or `MO_ASSEMBLY_COST`.** Because the WHERE is on the *previous* broadcast code, the broadcast code itself is editable (rename). *Body verified.* |
| `DELETE_AssyRatioInfo` | DELETE | `@BroadCode varchar(12)`. `DELETE … WHERE VC_BROADCAST_CODE = @BroadCode`. No cascade, no guard. *Body verified.* |
| `SELECT_AssyRatioInfoAssy` | SELECT | `@AssyCode varchar(12)`; `SELECT * … WHERE VC_ASSY_PART_NUMBER_CODE = @AssyCode`. **Not called by this form**; wired at `DataModule.pas:4576` (assembly-cost / build path). *Body verified.* |
| `SELECT_AssyRatioInfoRaw` | SELECT | `@BroadcastCode varchar(2)='', @TirePartNum varchar(12)='', @WheelPartNum varchar(12)=''`. Branchy lookup by broadcast code, OR tire1/tire2, OR wheel1/wheel2, else all. **Not called by this form**; wired at `DataModule.pas:4854`. Note it searches only `_1`/`_2` tire/wheel slots, never `_3`. *Body verified.* |

**`@AddDate`/`@Updated` timestamp math.** Both INSERT and UPDATE build:
`CONVERT(char(8), @Now, 112)` (= `yyyymmdd`, 8 chars) `+` **four** `SUBSTRING(CONVERT(varchar,@Now,114), n, 2)`
slices at offsets 1/4/7/10 = HH, MM, SS, **and the first two of the milliseconds field (= `ff`)**. That is
**8 + 2 + 2 + 2 + 2 = 16 characters** → format `yyyymmddHHMMSSff`, stored in the `varchar(16)` column.
This is the **identical 16-char stamp used by every other proc** (receiving/shipping/EDI/etc.) — NOT a
14-char outlier; the 4th slice (offset 10 of format 114, `hh:mi:ss:mmm`) IS the `ff`. D8 Bug 2's "16-char"
phrasing is correct. *(Verified char-by-char during the adversarial pass — the schema has exactly this
4-slice recipe in all 54 timestamp procs; no 3-slice/14-char variant exists.)*

---

## 4. Business rules & edge cases

### 4.1 Effective key = broadcast code (D2 / surrogate-key implications)
There is no surrogate id. The dup-check (`InsertAssyRatioInfo`, `DataModule.pas:2835` `If RecordCount=0`),
UPDATE (`WHERE VC_BROADCAST_CODE = @BroadCodePrev`), and DELETE all key on `VC_BROADCAST_CODE`. The form
also carries `BroadcastCode`/`BroadcastCodePrev` (`HoldDetails`, `:342–343`) so the code can be renamed
in place. Under **D2** the rebuild keeps a surrogate PK and treats broadcast code as an editable,
**site-unique** non-key attribute (per D1, unique on `(site_id, VC_BROADCAST_CODE)`).

### 4.2 Ratio-sum validation (the core business invariant)
`Verify` (`AssyRatioMaster.pas:122–139`) blocks save unless `TotalTireRatio_Edit.Text = 100` **and**
`TotalWheelRatio_Edit.Text = 100`. The totals are recomputed on every ratio edit (`TireRatioN_MaskEditChange`,
`WheelRatioN_MaskEditChange`) by summing the up-to-3 slots; the total turns **red** when ≠100/≠0
(`TotalTireRatio_EditChange`, `:508–514`). So: **tire ratios across the active tire slots must sum to
exactly 100; wheel ratios likewise.** This is the percentage-split semantics the explosion math relies
on (e.g. forecast-breakdown's `… *tireratio div 100`). *Confidence: high.*
- **Bug in the wheel total:** `WheelRatioN_MaskEditChange` (`:706, 733, 760`) gates the 3rd-slot add on
  `trim(TireRatio3_MaskEdit.Text) <> ''` — it checks the **tire** field, not the wheel field, when
  deciding whether to add `WheelRatio3`. With the maskedits defaulted to `'0'` (never empty) the third
  wheel ratio is in practice always included, so the visible total is usually right; but the condition
  is logically wrong (copy-paste from the tire handler). *Confidence: high — flag, don't preserve.*

### 4.3 Cascading combo-clear rules (data hygiene the UI enforces)
The `…ComboBoxChange`/`…MaskEditChange` handlers enforce **left-to-right slot ordering**: clearing
tire slot 1 (ItemIndex 0) zeroes slots 2+3 and their ratios; clearing slot 2 zeroes slot 3; selecting
slot 3 while slot 2 is empty bounces focus back. Same for wheels. So a saved row never has a "gap"
(slot 3 filled while slot 2 empty). The rebuild should reproduce this as a validation rule rather than
focus-juggling. *Confidence: high.*

### 4.4 Quantity radio-group encodings (non-obvious)
- `TireQty_RadioGroup`: index 0→**1** (note: `SetDetailBoxes` maps stored `TireQty=0`→index 0, but
  `HoldDetails` writes index 0→**1** — an asymmetry: a stored 0 displays as "index 0" yet re-saving
  persists 1), index 1→4, index 2→5 (`HoldDetails` `:377–383`, `SetDetailBoxes` `:298–304`).
- `WheelQty_RadioGroup`: 0→1, 1→4, 2→5 (`:421–427` / `:321–327`).
- `SpareTireQty_RadioGroup`: stored value **is** the index directly (`SpareTireQty := …ItemIndex`,
  `:428`). Capture the radio captions from the `.dfm` to map indices to real quantities; the integers
  above (1/4/5) are the persisted values. *Confidence: high for the int mappings; the asymmetry at
  TireQty index 0 is a real edge case.*

### 4.5 BOUNDARY with `forecast-detail.md` — RESOLVED: two different tables, parallel concepts
This was the open question in the task. The evidence:
- **`INV_ASSY_RATIO_MST`** (this module) is keyed by **broadcast code** and holds **up to 3 tire + 3
  wheel** part codes with **integer percentage ratios that sum to 100 per category**, plus quantities
  and spares.
- **`INV_FORECAST_DETAIL_INF`** (forecast-detail.md) is keyed by **assembly part × effective month**
  and holds **one tire + one wheel** part code (plus valve/film/label/misc) with the ratios
  (`tire %`, `wheel %`, `forecast %`) that `ForecastBreakdownF.UpdateForecast` actually reads to do
  the documented explosion `tirecount = ((WeekCount*forecastratio div 100)*tireratio) div 100`.
- **The forecast/order explosion reads `INV_FORECAST_DETAIL_INF`, NOT `INV_ASSY_RATIO_MST`.** Verified
  two ways: (a) `forecast-detail.md` and `forecast-breakdown.md` cite `INV_FORECAST_DETAIL_INF` as the
  BOM the breakdown reads; (b) grepping every `FROM`/`JOIN` of `INV_ASSY_RATIO_MST` across the whole
  schema shows the table is touched **only** by its own CRUD/raw SELECT procs, the part-master cascade
  triggers, and **one assembly-cost proc** that joins it to `INV_ASSY_MONTHLY_PO` (§4.6) — **no
  forecast or order-simulation proc reads it.** *Confidence: high.*
- **So they are parallel, partially-overlapping masters, not the same data.** Both encode
  broadcast/assembly → tire/wheel-part-with-ratio. `INV_FORECAST_DETAIL_INF` is the *single-component,
  effective-dated* form the live explosion uses; `INV_ASSY_RATIO_MST` is the *multi-component (3-way
  split), non-dated* form whose editor is currently UI-disabled (§1.1). The link between them is
  one-directional and weak: this form's AssyCode dropdown is *populated from*
  `INV_FORECAST_DETAIL_INF.VC_ASSY_PART_NUMBER_CODE` (`:447`), and the part-master rename/delete
  triggers cascade into **both** tables. The task's prompt referenced an order `tire_wheel_ratio` /
  `SIM_OrderSimulation` reading "the `INV_FORECAST_DETAIL_INF` / assy-ratio read" — confirmed that the
  ratio the order path consumes is the **forecast-detail** one, not `INV_ASSY_RATIO_MST`.
  *(If the order spec named `SIM_OrderSimulation` reading assy-ratio, verify against the order module's
  own analysis; from the schema side `INV_ASSY_RATIO_MST` has no order-sim reader.)*

### 4.6 Other consumer: assembly cost (separate Assembly sub-area, not explosion)
The only non-trigger, non-CRUD reader of the table is a proc that does
`select * from inv_assy_ratio_mst a , inv_assy_monthly_po p WHERE a.vc_assy_part_number_code =
p.vc_assy_part_number_code AND a.vc_assy_part_number_code = @Assycode AND (@ProdDate BETWEEN
VC_PO_MONTH_START AND VC_PO_MONTH_END)`. This joins the ratio master to `INV_ASSY_MONTHLY_PO` by
**assembly part number** (not broadcast code) to resolve an assembly's monthly PO/cost for a
production date — the `VC_BLANKET_PO` / `MO_ASSEMBLY_COST` columns (which the editor never populates)
belong to this **assembly-cost / monthly-PO** feature, suggesting those columns are maintained by a
different screen (MonthlyPO / Manifest Cost). The sibling tables `INV_ASSY_MONTHLY_PO`,
`INV_ASSY_PO_CHARGED`, `INV_ASSY_BUILD_HIST` form that cost/build sub-area and are out of scope here.
*Confidence: medium — the join proc body is verified; the ownership of `VC_BLANKET_PO`/`MO_ASSEMBLY_COST`
is inferred and should be confirmed (§8).* 

### 4.7 GALC-broadcast linkage (note, not deep analysis)
`VC_BROADCAST_CODE` is the Toyota **GALC assembly broadcast code** — the same broadcast identifier the
sibling **GALC broadcast receiver** (`GALCComm.exe`) consumes off the wire. In this system the broadcast
code is the *join key* between an assembly build signal and the parts it explodes into. Two observations:
(1) the column is only `varchar(2)`, so InventorySystem models the broadcast code as a **2-character**
value (the GALC wire may carry a longer code that is truncated/mapped to these 2 chars — confirm against
the GALC spec); (2) nothing in InventorySystem ingests the broadcast directly into this table — rows are
maintained as configuration (today, given §1.1, *not even via this form*). The runtime tie to actual
GALC build counts happens downstream through the **forecast** path (`INV_FORECAST_DETAIL_INF`), not here.
*Confidence: medium — boundary noted per scope; GALC wire detail deferred to the GALC analysis.*

### 4.8 Insert will fail on NOT-NULL columns the proc never sets (latent runtime bug)
`INSERT_AssyRatioInfo` omits `VC_BLANKET_PO` (`varchar(8) NOT NULL`) and `MO_ASSEMBLY_COST`
(`money NOT NULL`); the table DDL shows **no default constraints**. As written, a fresh INSERT would
violate NOT NULL and fail — unless the **live** database has DEFAULT constraints not present in the
committed schema snapshot (see [[reference-schema-snapshot-vs-live]] — signature/DDL mismatches are
common; treat as **verify-live**). This is consistent with §1.1 (the form is disabled, so the failing
INSERT path is never exercised in production). *Confidence: medium — DDL says it should fail; live
defaults may mask it. Flag for verification, do not assume.*

---

## 5. UI / UX notes
- Standard master-data pattern: top **detail panel** (broadcast code edit, assy-code combo, assy-name,
  3 tire combos + 3 ratio mask-edits + tire-qty radio, 3 wheel combos + 3 ratio mask-edits + wheel-qty
  radio, spare tire qty radio + spare tire/wheel edits, live tire-total + wheel-total boxes) over a
  read-only **grid** (`ASSYRatioMaster_DBGrid`) bound to `Inv_DataSet`. Buttons: Insert / Update /
  Search / Clear / Delete / Close.
- Selecting a grid row (`KeyUp`/`MouseUp`/`DataChange` → `HoldDetails(True)` → `SetDetailBoxes`) loads
  the detail panel from `Fields[0..19]` (positional — fragile if the SELECT column order changes).
- Search filters the in-memory dataset client-side (`Filter := '[Broadcast Code] = ' + QuotedStr(...)`),
  not a server round-trip.
- **Keep:** ratio-sums-to-100 validation; left-to-right slot ordering; broadcast-code uniqueness.
  **Modernize:** the disabled menu path (decide whether this screen should exist at all, or whether the
  3-way split should be merged into forecast-detail); the positional grid coupling; the fixed
  tire1/2/3 + wheel1/2/3 columns → a child ratio table (the direction `BCRatioMaster` was reaching for).

---

## 6. Target design (Ignition Perspective + Named Queries)
- **Perspective view:** a master-CRUD view `assembly/assy-ratio-master` mirroring the master-data
  views (list table + edit form), gated behind a role — and **explicitly flagged as currently
  UI-disabled in the legacy app** (§1.1): confirm with the domain expert before surfacing it
  (§8.1). Dropdowns:
  - Tire part numbers ← Named Query `parts/list_tire_parts` (`INV_PARTS_STOCK_MST WHERE VC_TIRE_WHEEL='T'`).
  - Wheel part numbers ← `parts/list_wheel_parts` (`='W'`).
  - Assembly codes ← `forecast/list_assy_codes` (`DISTINCT VC_ASSY_PART_NUMBER_CODE FROM INV_FORECAST_DETAIL_INF`).
- **Named Queries** (one per proc, mirroring schema per [[ignition-named-query-crud-practice]]):
  `assembly/assy_ratio_list` (= `SELECT_AssyRatioInfo`), `assy_ratio_insert`, `assy_ratio_update`,
  `assy_ratio_delete`. Reuse the existing procs read-only in Stage 1; reimplement in Stage 3.
- **Validation in the view's transform/script, not focus-juggling:** tire-ratio slots sum = 100,
  wheel-ratio slots sum = 100; no slot gaps (slot N filled ⇒ slot N-1 filled); broadcast code unique
  per site.
- **Schema fixes baked into the rebuild (do NOT preserve legacy bugs):**
  - Add surrogate PK + unique `(site_id, VC_BROADCAST_CODE)` (D1/D2) so uniqueness isn't app-only.
  - Normalize the fixed 3-tire/3-wheel columns into a child `assy_ratio_component` table
    `(assy_ratio_id, kind 'T'/'W', slot, part_number, ratio)` — removes the slot-3 trigger/raw-SELECT
    gaps (§2 cascade, §3 `…Raw`) for free and is the generalization `BCRatioMaster` intended.
  - Fix the wheel-total handler's tire-field condition (§4.2) and the TireQty index-0 asymmetry (§4.4).
  - Make `VC_BLANKET_PO` / `MO_ASSEMBLY_COST` nullable-with-default or own them in the assembly-cost
    module (§4.6) so inserts can't fail (§4.8).
  - Cascade part renames/deletes via real FKs (RESTRICT per **D3**, or blank-on-delete if the domain
    requires it) covering **all** slots + spares, not just slots 1–2.
- **No reports** specific to this module.

## 7. Cross-cutting findings (P-patterns)
- **D8 Bug 1 (size dup-check on the wrong table) — this module's proc is the wrong target.**
  `InsertSizeInfo` (`DataModule.pas:2531`) runs its duplicate check via **`SELECT_AssyRatioInfo`**
  (`DataModule.pas:2543`; the caller adds a param it calls `@SizeCode` that lands positionally in the
  proc's real `@BroadCode` filter) → filters `INV_ASSY_RATIO_MST.VC_BROADCAST_CODE`,
  so a size code colliding with a broadcast code is falsely rejected and genuine size dups are never
  caught. The proc itself is correct; the **misuse** is in the Size module. Already captured in
  `docs/analysis/decisions.md` D8 Bug 1 and `docs/analysis/master-data/size.md`; cross-referenced here.
- **P12 wrong-target retry-recursion — all AssyRatio involvements ALREADY documented; NO new bugs.**
  Verified the four AssyRatio CRUD methods recurse to **themselves** (correct target):
  `GetAssyRatioInfo`→`GetAssyRatioInfo` (`:2796`), `InsertAssyRatioInfo`→`InsertAssyRatioInfo` (`:2907`),
  `UpdateAssyRatioInfo`→`UpdateAssyRatioInfo` (`:2994`), `DeleteAssyRatioInfo`→`DeleteAssyRatioInfo`
  (`:3040`) — these are pure P8 self-retry, not P12. The cases where an AssyRatio proc is the *wrong*
  target are already in `docs/analysis/cross-cutting/datamodule-retry-target-bugs.md`:
  **CRITICAL #5** `DeleteRecConfStatInfo`→`DeleteAssyRatioInfo` (`:3091`, keyed on stale `fBroadCode`,
  "breaks forecast/breakdown math"); **MODERATE** `UpdateINVDone`→`UpdateAssyRatioInfo` (`:4351`);
  **LOW** `GetBCRatioInfo`→`GetAssyRatioInfo`. **No new cross-cutting bug found in this module.**
- **P8 (per-method recursive retry boilerplate) + P9 (shared mutable fields):** present in all four
  CRUD methods (`fBroadCode`, `fAssyCode`, `fTire*`, `fWheel*` are unit-level singletons). The rebuild's
  generic transport-retry wrapper + explicit per-call args dissolves this class (per the cross-cutting
  doc's "real fix").
- **DEAD-CODE / DEAD-PATH discipline:** `BCRatioMaster` = dead unit (not in dpr) + stub; `AssyRatioMaster`
  = live unit on a dead UI path (`MasterMaint.pas:78`). Both flagged so neither is over-claimed as a
  shipping feature.

## 8. Open questions for the user (domain expert)
1. ✅ RESOLVED (D12) — **DROP for the rebuild.** David: *"INV_ASSY_RATIO_MST failed conversion thought,
   drop for rebuild."* The screen + table were an abandoned design (hidden "not used yet"; no forecast/order
   proc reads the table). Do NOT port `INV_ASSY_RATIO_MST`, `AssyRatioMaster`, or `BCRatioMaster`.
2. ✅ RESOLVED (D12) — moot: the table is dropped (Q1). The live broadcast→part ratio model lives entirely
   in `INV_FORECAST_DETAIL_INF`, which the rebuild carries.
3. **Who owns `VC_BLANKET_PO` and `MO_ASSEMBLY_COST`?** The editor never sets them; an assembly-cost
   proc joins them to `INV_ASSY_MONTHLY_PO`. Confirm these are maintained by the MonthlyPO/Manifest-Cost
   module and verify whether the live DB has DEFAULT constraints (otherwise `INSERT_AssyRatioInfo` would
   fail — §4.8).
4. **Broadcast-code width vs GALC.** Is the 2-char `VC_BROADCAST_CODE` the full GALC broadcast code, or a
   truncation/mapping of a longer wire code? Needed to align with the GALC receiver.
5. **Ratio semantics confirmation:** tire ratios sum to 100 across active tire slots and wheel ratios sum
   to 100 across active wheel slots — confirm this is "percent of this assembly's tires that are
   part X", consistent with the forecast-breakdown `… div 100` math.

## 9. Test cases / parity checks
- **Dup guard:** insert broadcast code `'AA'`, re-insert `'AA'` → second logs `FAILED … (DUPLICATE)`,
  no second row (matches `If RecordCount=0`).
- **Ratio validation:** tire ratios 60/40/0 (=100) save; 60/30/0 (=90) blocked with "Tire ratio must
  total to 100"; same for wheels.
- **Rename via UPDATE:** edit broadcast code `'AA'`→`'AB'`; UPDATE keys on `@BroadCodePrev='AA'`, row now
  `'AB'`, `VC_LAST_UPDATE` = a 16-char `yyyymmddHHMMSSff` stamp (assert exactly 16 chars, same as all procs).
- **Timestamp length:** after insert, `len(VC_ADD)=14`; after update, `len(VC_LAST_UPDATE)=14`.
- **Part-rename cascade:** rename a tire part used as `_1_CODE` → ratio row's `VC_TIRE_PART_NUMBER1_CODE`
  follows (via `UPDATE_PartNumber`); rename one used only as `_3_CODE` → **does not follow** (documents
  the slot-3 gap; the rebuild should fix it).
- **Part-delete cascade:** delete a part used as `_2_CODE` → blanked to `''`; used only as spare → **not**
  blanked (slot gap).
- **Boundary check:** confirm no order/forecast proc reads `INV_ASSY_RATIO_MST` (grep `FROM/JOIN`):
  only CRUD/raw SELECTs, the two part-master cascade triggers, and the `INV_ASSY_MONTHLY_PO` cost join.
