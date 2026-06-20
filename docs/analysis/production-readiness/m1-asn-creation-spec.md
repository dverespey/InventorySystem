# M1 — ASN Entry Creation: Source-Truth Behavioral Spec

**Keystone:** the morning revenue starter (Rank 1). "ASN sequence check" → "Create ASN entries
only" → header + N manifest-detail rows + the "(No Ratio)" remainders, all under one ASN id, in one
transaction.

**Author:** delphi-architect · **Date:** 2026-06-19
**Live caller:** `ASNSelect.pas` (confirmed live in `InventorySystem.dpr`).
**Procs (Inventory DB):** `DB Schema/CreateInventory.sql` (UTF-16LE; UTF-8 at `/tmp/inv_utf8.sql`).
**Procs (ALC / VehicleOrder DB):** `/Users/apple/Documents/FP docs/SQL/VehicleOrder.sql` (UTF-16LE;
UTF-8 at `/tmp/vo_utf8.sql`).

> **The single biggest correction to the plan:** there is **no `CalculateASNFRS` stored proc, and no
> `INV_ASSY_RATIO_MST` lookup.** The split-by-ratio fan-out is **hand-written Delphi** in
> `DataModule.pas:5106 CalculateASNFRS`, driven by **two source datasets** — `AD_FRSPull` (ALC DB,
> **body NOT in any dump available here** — the one true source gap) and `SELECT_ForecastDetailBCASN`
> (Inventory DB, body read). The ratio split is **live and active**, done entirely in Pascal over the
> forecast-detail ratios; D12's dropped `INV_ASSY_RATIO_MST` was a different/abandoned mechanism and
> is **not** what the live fan-out uses. See §4.

---

## 0. The ordered create-ASN chain (what one click does)

Trigger: operator clicks **Create ASN/Entries** (`CreateASNEntries_Button`,
`ASNSelect.pas:369 CreateASNEntries_ButtonClick`). The full ordered chain, with the DB each step
hits:

| # | Step | Code | DB | Effect |
|---|---|---|---|---|
| 0a | `BeginTrans` on **Inv_Connection** | `ASNSelect.pas:372` | Inv | opens the single transaction |
| 0b | Read site EIN: `SiteDataset.Open` (`AD_GetSite`, `@LineName=''`) | `ASNSelect.pas:378-380` | **ALC** | `fEIN := Site.SiteEIN` (first/only Site row — single-site, §2) |
| 1 | `INSERT_ASNInfo` (OUTPUT `@ASNID`) | `DataModule.pas:5321 InsertASNInfo` → proc `/tmp/inv_utf8.sql:2529` | Inv | inserts `INV_ASN_MST` header, status `'C'`, `@Ein=fEIN+1`; `fRecordID := SCOPE_IDENTITY()` |
| 2 | `CalculateASNFRS` (Delphi, **not a proc**) | `DataModule.pas:5106`, called at `:5381` | both | the fan-out: pull build data + forecast ratios, INSERT N detail rows (§4) |
| 2a | per BC: `AD_FRSPull` | `DataModule.pas:5125` | **ALC** | **body unverified** — returns one row per broadcast code with `BC`, `Orders`, `VEHICLES` (§3/§4) |
| 2b | per BC: `SELECT_ForecastDetailBCASN` | `/tmp/inv_utf8.sql:3011` | Inv | forecast-detail rows for that BC (part numbers + tire/wheel ratios + assy qty) |
| 2c | per detail row: `INSERT_ASNDetail` | `DataModule.pas:5191/5243` → proc `/tmp/inv_utf8.sql:2682` | Inv | upsert one `INV_ASN_DETAIL_MST` row per manifest (accumulates) |
| 2d | post-loop: `SELECT_ASNMissingCost` | `/tmp/inv_utf8.sql:3504` | Inv | read-only audit: logs any detail with no in-window manifest cost (does **not** abort) |
| 3 | `AD_UpdateEIN` | `ASNSelect.pas:387-389` → proc `/tmp/vo_utf8.sql:623` | **ALC** | `UPDATE Site SET SiteEIN = SiteEIN+1` (no WHERE — bumps every site row) |
| 4 | `CommitTrans` on **Inv_Connection** | `ASNSelect.pas:391` | Inv | commits header + details |

**Transaction boundary hazard (call out for the rebuild):** the `BeginTrans`/`CommitTrans` is on
**`Inv_Connection` only** (`ASNSelect.pas:372,391`). Steps 0b (`AD_GetSite`), 2a (`AD_FRSPull`), and
**3 (`AD_UpdateEIN`)** all run on the **ALC_Connection**, which is **outside the transaction**. So:
- A rollback of the Inventory side does **not** roll back the `SiteEIN` increment. The EIN counter
  can advance even when the ASN insert fails → EIN gaps. (Latent legacy bug; the single-user desktop
  rarely hit it.)
- `INSERT_ASNDetail` (`INV_ShippingStoredProc`) and `INSERT_ASNInfo` (`Inv_StoredProc`) are both on
  Inv_Connection → correctly inside the transaction.

There is a **second create path** — `CreateASN_Button` / `CreateASN/Files`
(`ASNSelect.pas:432 CreateASN_ButtonClick`) — that additionally builds the EDI 856 file inline and
flips status to `'S'` in the same transaction. **The daily log uses "Create ASN entries only"
(rows 4-25), not "Create ASN/Files".** M1 reproduces the **entries-only** path (§7); the 856 build +
status flip is M1 Rank 2, decoupled per Q2.

---

## 1. ASN sequence check (`SELECT_ASNSeq` + the ALC seq pickers)

The log line `ASN Sequence number check, P:.. L:.. S:909 E:756 Q:848` is emitted at
`ASNSelect.pas:358` after the **Check** button runs. The S/E/Q values do **not** come from
`SELECT_ASNSeq` — that proc only guards against re-creating an existing ASN. The real flow:

**1a. `SELECT_ASNSeq` — the "already exists?" guard** (`/tmp/inv_utf8.sql:1517`):
```sql
CREATE PROCEDURE [dbo].[SELECT_ASNSeq] @LineName varchar(50), @PDate varchar(8) AS
  SELECT * FROM INV_ASN_MST
  WHERE VC_PRODUCTION_DATE = @PDate AND VC_LINE_NAME = @LineName
    AND VC_START_SEQ_NUMBER <> -1
```
Called from `ASNSelect.pas:137 LoadSeqNumbers` on date change. **If a row exists** (an ASN already
created for this line+production-date), the screen **locks** the seq fields read-only and disables
Check/Create (`:162-182`). **If none exists**, the operator may enter sequence numbers and Check is
enabled (`:184-201`). So `SELECT_ASNSeq` is a **per-line + per-production-date idempotency lock**, not
the source of the range. **Verified.** Multi-site note: it keys on `(VC_PRODUCTION_DATE, VC_LINE_NAME)`
with no site — re-key to `(site_id, line, prod_date)`.

**1b. Where S / E / Q actually come from (the GALC-fed Vehicle table, ALC DB):**
- **Start seq (S)** default: `GetNextASNDate` (`DataModule.pas:3712`) calls `SELECT_ASNMax` (Inv,
  last ASN's production date) then **`AD_GetNextASN`** (`/tmp/vo_utf8.sql:1917`) — the first Vehicle
  *after* the last ASN's end time on this line. `fStartSeq := Vehicle.ASN`. (S is the broadcast
  sequence number stamped on each built vehicle.)
- **Start/End times** (the `StartBox`/`EndBox` datetime combos): `GetLastSeqDate`
  (`DataModule.pas:4161`) → **`AD_GetLastPrint`** (`/tmp/vo_utf8.sql:1813`): for a given `@ASN`
  (3-4-digit seq, which **wraps** — note S:909→E:756) and line, returns `max(DateCreated)` (the most
  recent build at that wrapped sequence). The operator picks the specific timestamp; this disambiguates
  the wrap. `@RevCount` (= `fiRevSeqLookup`, INI) bounds the lookback window.
- **Qty (Q)** = the vehicle count: **Check** (`ASNSelect.pas:334 Check_ButtonClick`) calls
  `CheckShippingInfo` (`DataModule.pas:4195`) → **`AD_ProductionSeq`** (`/tmp/vo_utf8.sql:2396`):
  ```sql
  SELECT v.ASN, VehicleID FROM Vehicle v JOIN Line l ON v.lineID=l.lineID AND l.LineName=@LineName
  WHERE DateCreated between @begindate and @enddate
  ```
  `Q := ALC_StoredProc.recordCount` (`ASNSelect.pas:351`) = **number of vehicles built between the
  chosen start-time and end-time on this line.** This is the shipped vehicle quantity (848). The seq
  bounds in the proc are **commented out** — the active query filters **purely by the datetime range**,
  not by ASN number. **Verified.**

**So the range comes from the GALC broadcast (the ALC `Vehicle` table), not an Inventory table.** The
operator picks it by: line + production date (auto-defaulted) → enters/confirms start & last seq →
picks the matching start-time & end-time from the datetime combos → Check counts the vehicles. **Q is
a row count, not a stored value.**

**Verified:** `SELECT_ASNSeq`, `AD_GetNextASN`, `AD_GetLastPrint`, `AD_ProductionSeq` bodies all read.

---

## 2. `INSERT_ASNInfo` — the header insert

**Proc body** (`/tmp/inv_utf8.sql:2529`):
```sql
CREATE PROCEDURE [dbo].[INSERT_ASNInfo]
  @ASNID int OUTPUT, @LineName varchar(50), @AssyLine varchar(1),
  @StartSeq varchar(4), @DTStartSeq datetime, @EndSeq varchar(4), @DTEndSeq datetime,
  @Qty integer, @PDate varchar(8), @Ein integer AS
BEGIN
  ... @AddDate = <yyyymmdd + HHmmss as 14-char string> ...
  INSERT INTO INV_ASN_MST
  VALUES( @Ein, 'C', @LineName, @AssyLine, @StartSeq, @DTStartSeq, @EndSeq, @DTEndSeq,
          @Qty, @PDate, @AddDate, @AddDate)
  SET @ASNID = SCOPE_IDENTITY()
END
```
Confirmed against the plan's §2 Rank 1 assumptions — **all correct**:
- Status hard-coded **`'C'`** (created). ✔
- `@ASNID` returned via **OUTPUT = `SCOPE_IDENTITY()`**; caller stores it in `fRecordID`
  (`DataModule.pas:5364`). ✔
- Columns set (positional VALUES, matches `INV_ASN_MST` `/tmp/inv_utf8.sql:692`):
  `IN_ASN_EIN, VC_ASN_STATUS='C', VC_LINE_NAME, VC_ASSEMBLY_LINE, VC_START_SEQ_NUMBER, DT_START_SEQ,
  VC_END_SEQ_NUMBER, DT_END_SEQ, IN_QTY, VC_PRODUCTION_DATE, VC_LAST_UPDATE, VC_ADD`.
- `@AddDate` is a **14-char `yyyymmddHHmmss` string** built via `CONVERT(112)+114` substrings (the
  repo-wide date convention) — written to **both** `VC_LAST_UPDATE` and `VC_ADD`.

**EIN handling at create (ties Q4) — verified, and it contradicts the plan's "allocated at send, not
create":**
- `fEIN := SiteDataset['SiteEIN']` (read from ALC `Site`, `ASNSelect.pas:380`).
- The header is inserted with **`@Ein = fEIN+1`** (`DataModule.pas:5358-5359`). So the **next** EIN is
  written onto the ASN header **at create time**, status `'C'`.
- Detail rows are also inserted with `@EIN = fEIN+1` (`:5196`, `:5248`).
- **`AD_UpdateEIN` then bumps the counter at the very end of the create transaction**
  (`ASNSelect.pas:387`), so the *next* ASN gets the next number.

> **Correction to the plan (and to Q4):** the legacy allocates the EIN **at CREATE**, not at send.
> The 856 send (`CreateASN/Files` path) re-reads the same `fEIN+1` and just flips status to `'S'`; it
> does **not** allocate a fresh EIN. `AD_UpdateEIN` (`UPDATE Site SET SiteEIN = SiteEIN+1`, no WHERE,
> `/tmp/vo_utf8.sql:623`) is the only place the counter advances, and it runs **inside the
> entries-only create** click. **Build decision for M1:** the plan's "EIN allocated atomically per
> site at send" is a *deliberate improvement*, not parity — flag it. If David wants strict parity the
> EIN is allocated at create; if he wants the cleaner model, allocate at send. Either way the
> per-site atomic sequence replaces `Site.SiteEIN+1` + the unscoped `AD_UpdateEIN` (which today bumps
> **every** site row — a real multi-site bug, BLOCKER-class).

---

## 3. `CalculateASNFRS` — the flagged "cross-module proc" (BIGGEST UNKNOWN)

**It is NOT a stored proc. It is a Delphi method** — `DataModule.pas:5106 CalculateASNFRS`. The
implementation plan and §6 list it as a body to "confirm with delphi-architect"; here is the truth.

**What it actually does** (verified from the Pascal body, `DataModule.pas:5106-5319`):
1. Calls **`AD_FRSPull`** (ALC DB, `ALC_StoredProc`, `:5125`) with `@begindate,@enddate,@Start,@Last,
   @LineName` (the checked range). Returns a result set the code iterates; per row it reads fields
   **`BC`** (broadcast code), **`Orders`** (vehicles-per-broadcast count), **`VEHICLES`** (vehicle
   count for the multiply). One row per distinct broadcast code shipped in the range.
2. For each BC, calls **`SELECT_ForecastDetailBCASN`** (Inventory DB, §4) to get the assembly parts +
   ratios for that broadcast code.
3. Computes per-part ship qty and calls **`INSERT_ASNDetail`** (§5) — this is the detail fan-out.
4. Post-loop, runs **`SELECT_ASNMissingCost`** (`:5287`) as a **read-only warning** (logs, does not
   raise).

**It is NOT a stock decrement and NOT a renban/FRS-date calc.** Despite the name "FRS", it does **no
inventory decrement** and **no FRS-date math** in the verified body. The comment at `:5378-5380`
("Calculate the number of assemblys built ... and remove inventory") is **aspirational/stale — there
is no decrement in the code.** (Inventory is decremented elsewhere, via the part-shipping trigger
path `UpdatePartShipping`, `/tmp/inv_utf8.sql` — not here.) **The method's real job is: turn the
build data + forecast ratios into ASN detail rows.** Mark the plan's "(stock decrement; cross-module)"
annotation as **WRONG** — no decrement happens in `CalculateASNFRS`.

**What is verifiable vs inferred:**
- **VERIFIED (read the Pascal):** the orchestration, the two source procs called, the ratio math, the
  `INSERT_ASNDetail` calls, the missing-cost warning, the error/abort conditions.
- **INFERRED (body unverified — THE source gap):** **`AD_FRSPull`** is **not present in either
  `VehicleOrder.sql` dump** available here (`/Users/apple/Documents/FP docs/SQL/VehicleOrder.sql` nor
  `.../Database Scripts/VehicleOrder.sql`). From the call site we know only its **contract**:
  - **Params:** `@begindate, @enddate` (the chosen start/end timestamps), `@Start, @Last` (seq
    numbers), `@LineName`.
  - **Returns** (columns the Delphi reads): `BC` (broadcast code, string), `Orders` (int — the
    "single vehicle ≤5" branch trigger), `VEHICLES` (int — the multiplier for qty).
  - **Almost certainly** a roll-up of the ALC `Vehicle` rows in the range, grouped by broadcast code,
    counting vehicles per BC. (Inferred from how `VEHICLES` and `Orders` are used; the grouping by
    `BC` mirrors `AD_ProductionSeq`'s per-vehicle select.)
  - **DEVELOPER ACTION (blocks M1 parity):** obtain `AD_FRSPull`'s body from the live ALC/VehicleOrder
    DB before porting. The fan-out **cannot be reproduced faithfully without it** — the qty per
    manifest depends on `VEHICLES` and `Orders` per BC, which only `AD_FRSPull` produces. This is the
    one genuinely missing piece. (`AD_GetNextASN`, `AD_GetLastPrint`, `AD_ProductionSeq`,
    `AD_GetSite`, `AD_UpdateEIN` were all found and read; only `AD_FRSPull` is absent.)

**Abort conditions inside `CalculateASNFRS` (must reproduce):**
- Missing manifest-cost on any assy part of a BC → `raise`, whole create fails (`:5170-5175`).
- A BC with **no** forecast-detail rows → `raise` "Missing Broadcast Code Information" (`:5273-5275`).
- On any exception the function returns `False` (`:5316`) → `InsertASNInfo` returns False →
  `CreateASNEntries_ButtonClick` rolls back (`ASNSelect.pas:417`).

---

## 4. The split-by-ratio detail fan-out (~20 rows + 2 "(No Ratio)")

**Driver:** `CalculateASNFRS` Pascal loop (`DataModule.pas:5180-5268`), **not** `INV_ASSY_RATIO_MST`
(D12 — that table was a *different*, dropped conversion; it is **not referenced anywhere in the live
create path**, confirmed by grep). The live ratio source is **`INV_FORECAST_DETAIL_INF`** via
`SELECT_ForecastDetailBCASN`. **The ratio split is LIVE and ACTIVE.**

**`SELECT_ForecastDetailBCASN`** (`/tmp/inv_utf8.sql:3011`, verified):
```sql
CREATE PROCEDURE [dbo].[SELECT_ForecastDetailBCASN] @BCode varchar(20), @EffMonth varchar(7) AS
  SELECT * FROM INV_FORECAST_DETAIL_INF f
    LEFT JOIN INV_MANIFEST_COST_MST c ON f.VC_ASSY_PART_NUMBER_CODE = c.VC_ASSY_PART_NUMBER_CODE
  WHERE @BCode LIKE VC_BROADCAST_CODE
    AND ((VC_EFFECTIVE_MONTH = @EffMonth or VC_EFFECTIVE_MONTH = '')
    AND IN_TIRE_RATIO <> 0 AND IN_WHEEL_RATIO <> 0 )
```
Returns the assembly part rows for the broadcast code (matched by `LIKE`, effective-month-filtered),
each with `VC_ASSY_PART_NUMBER_CODE`, `VC_ASSY_MANIFEST_NUMBER` (the 2-char manifest id, §6),
`IN_ASSY_QTY`, `IN_TIRE_RATIO`, `IN_WHEEL_RATIO`, `IN_MANIFEST_COST_ID`. The manifest-cost join is
used only to detect missing cost (`IN_MANIFEST_COST_ID IS NULL` → abort, `:5163`).

**The two branches (this is exactly what produces the 20 + 2 rows):**

**(A) "No Ratio" branch** — `if AD_FRSPull.Orders <= 5` (`DataModule.pas:5183`):
```
manifest := '7' + copy(fProductionDate,4,5) + VC_ASSY_MANIFEST_NUMBER
count    := VEHICLES * IN_ASSY_QTY            // NO ratio applied
INSERT_ASNDetail(@ASNID=fRecordID, @EIN=fEIN+1, @Manifest, @PartNumber, @Qty=count)
break                                          // only the FIRST forecast-detail row, then stop this BC
```
Logged as `INSERT ASN entry(No Ratio) ASNID(..) Manifest(..) Qty(..)` (`:5211`). The "(No Ratio)"
rows in the daily log (manifests 76061836, 76061851) are **broadcast codes with ≤5 vehicles** — too
few to split by ratio, so the full assy qty ships against one part with **no ratio multiply** and the
loop `break`s after one row.

**(B) Ratio branch** — `else` (`>5` vehicles, `:5214-5265`), per forecast-detail row:
```
if IN_TIRE_RATIO = 100 AND IN_WHEEL_RATIO = 100:
    count := VEHICLES * IN_ASSY_QTY                              // both 100% → no split
else:
    count := round( VEHICLES * IN_ASSY_QTY * IN_TIRE_RATIO / 100 )   // tire-ratio split
manifest := '7' + copy(fProductionDate,4,5) + VC_ASSY_MANIFEST_NUMBER
INSERT_ASNDetail(...)   // one row per forecast-detail row (no break)
```
Logged as `INSERT ASN entry ASNID(..) Manifest(..) Qty(..)` (`:5263`). These are the ~20 normal rows.

**Key fidelity points (must reproduce exactly):**
- The split uses **`IN_TIRE_RATIO` only** as the divisor-percentage (`/100`), with `round()`
  (banker's? — Delphi `round` is round-half-to-even; **flag** — confirm against golden qty). The code
  comment (`:5222`) says tire and wheel share are set equal for forecast, so one ratio suffices.
- The "(No Ratio)" trigger is **`Orders <= 5`** (a per-BC vehicle/order count from `AD_FRSPull`), NOT
  a "ratio doesn't divide evenly" remainder as `daily-workflow-usage.md §2` row-64 inferred. **Correct
  the workflow doc:** "(No Ratio)" = small-volume broadcast code (≤5), shipped whole, not an
  even-division remainder.
- One detail row per `(BC × forecast-detail-part)` in branch B; **exactly one** row per BC in branch A.

**Data-dependent claims to confirm against the golden (per output discipline):**
1. The `Orders <= 5` threshold producing the 2 "(No Ratio)" rows (manifests 76061836, 76061851) — confirm
   those two BCs had `AD_FRSPull.Orders <= 5` and the other ~20 had `> 5`.
2. The `round(VEHICLES * IN_ASSY_QTY * IN_TIRE_RATIO / 100)` qty for a ratio-split row vs the legacy
   `IN_QTY` actually written (round-half-to-even vs away-from-zero matters at .5).
3. That `INV_ASSY_RATIO_MST` is truly dead — confirmed no reference in the create path here, but verify
   no OTHER live caller exists before declaring it removable.

---

## 5. `INSERT_ASNDetail` — the per-manifest upsert

**Proc body** (`/tmp/inv_utf8.sql:2682`, verified — matches Q1's description exactly):
```sql
CREATE PROCEDURE [dbo].[INSERT_ASNDetail]
  @ASNID integer, @EIN integer, @Manifest varchar(8), @PartNumber varchar(12),
  @Qty integer, @HotCall bit=0 AS
BEGIN
  ... @AddDate = <14-char timestamp> ...
  if @HotCall = 0
  begin
    SELECT * FROM INV_ASN_DETAIL_MST WHERE VC_MANIFEST_NUMBER = @Manifest
    if @@rowcount = 0
      INSERT INTO INV_ASN_DETAIL_MST VALUES(@ASNID, null, @EIN, @Manifest, @PartNumber, @Qty, @AddDate, @AddDate)
    else
      UPDATE INV_ASN_DETAIL_MST SET IN_QTY = IN_QTY + @Qty WHERE VC_MANIFEST_NUMBER = @Manifest
  end
  else
    INSERT INTO INV_ASN_DETAIL_MST VALUES(@ASNID, null, @EIN, @Manifest, @PartNumber, @Qty, @AddDate, @AddDate)
END
```
- **`@HotCall=0` (normal):** SELECT-by-manifest → if absent INSERT, else **`IN_QTY += @Qty`
  accumulate**. The accumulate is intentional and load-bearing (Q1 RESOLVED): the same manifest can be
  hit by multiple BCs/parts in one create and the qty sums. ✔
- **`@HotCall=1`:** always INSERT (hot calls never dedup). ✔ (Used by `HotCallEntry.pas`, not the
  morning create — the morning path never passes `@HotCall`, so it defaults 0.)
- `IN_INV_ID` inserted **NULL** (set later when the invoice is created). `VC_ADD` = `VC_LAST_UPDATE` =
  14-char timestamp. Positional VALUES match `INV_ASN_DETAIL_MST` (`/tmp/inv_utf8.sql:2179`).

**The Q1 re-key to `(site_id, IN_ASN_ID, manifest)` maps onto this in two places (both verified to key
on manifest ALONE today — the bug Q1 fixes):**
1. The existence check: `WHERE VC_MANIFEST_NUMBER = @Manifest` (`:upsert SELECT` + the `UPDATE` WHERE)
   → re-key to `WHERE site_id=@site AND IN_ASN_ID=@ASNID AND VC_MANIFEST_NUMBER=@Manifest`.
2. **`DELETE_ASNItem`** (`/tmp/inv_utf8.sql:2800`, verified):
   ```sql
   CREATE PROCEDURE [dbo].[DELETE_ASNItem] @ManifestNumber varchar(8) AS
     DELETE FROM INV_ASN_DETAIL_MST WHERE (VC_MANIFEST_NUMBER = @ManifestNumber)
   ```
   Takes **only `@ManifestNumber`** → today deleting one ASN's manifest line wipes that manifest from
   **every** ASN (and every site). Re-key to `(site_id, IN_ASN_ID, manifest)` per Q1. The supporting
   index `IX_INV_ASN_DETAIL_MST` is on `VC_MANIFEST_NUMBER` only (`/tmp/inv_utf8.sql:2199`) — extend it
   to the composite key.

> **Cross-ASN accumulation bug (preserve the intent, fix the scope):** because the upsert keys on
> manifest alone, a *later* ASN (different production day) that reuses a manifest number would
> accumulate into the *old* ASN's detail row. Within one create that's the desired accumulate; across
> ASNs it's data corruption. The `(site_id, IN_ASN_ID, manifest)` re-key keeps the within-ASN
> accumulate and removes the cross-ASN/cross-site collision. ✔ matches Q1 RESOLVED.

---

## 6. Manifest number generation

**Scheme:** `'7' + copy(fProductionDate, 4, 5) + VC_ASSY_MANIFEST_NUMBER`
(`DataModule.pas:5186` and `:5239`, identical in both branches).
- `fProductionDate` is `yyyymmdd` (e.g. `20260618`). `copy(...,4,5)` = **chars 4-8** = `60618` =
  the last digit of year (`6`) + `MM` (`06`) + `DD` (`18`). So the prefix is `'7' + '6' + '06' + '18'`
  = `7` + 1-digit-year + MM + DD.
- `VC_ASSY_MANIFEST_NUMBER` is the **2-char id from `INV_FORECAST_DETAIL_INF`** (per assembly part).
- Result: `7` + `60618` + `57` = `76061857` — **8 chars, exactly matching the daily log manifests
  `76061857`…`76061805`.** ✔ Confirms `asn-invoice.md §4.6`'s `'7'+yy+MM+DD+2-char id` (note `yy` is
  really the **single last digit** of the year via `copy(,4,5)`, not 2-digit year — **correct the spec
  to "1-digit year".** Confirm against golden: a 2027 production date would yield `7` + `70618` + id.)
- **When/where generated:** at fan-out time in `CalculateASNFRS` (the broadcast-fed morning path), one
  per assembly part per BC. The **manual/hot-call** path (`HotCallEntry.pas`) supplies the manifest
  differently (e.g. log row 150 `manifest(52089913)` — a vehicle/kanban-derived number, not this
  `7`-prefixed scheme), so the `7`+date+id generator is **specific to the broadcast-fed create**, not
  hot calls. The plan's `asn/Manifest` combo for manual adds is a separate concern.

---

## 7. What the Ignition build MUST reproduce (entries-only / "Create ASN entries" path)

A single Gateway transaction (`create_asn`) ordered exactly as §0:

1. **Idempotency guard:** run `SELECT_ASNSeq(line, prod_date)` re-keyed `(site_id, line, prod_date)`;
   if a row exists, block (the legacy locks the UI). 
2. **Resolve EIN** from the per-site sequence (replaces ALC `Site.SiteEIN`). **Decide parity vs
   improvement on allocation timing (§2):** legacy allocates `SiteEIN+1` at **create**; plan proposes
   at **send**. Pick one explicitly with David.
3. **`INSERT_ASNInfo`** — status `'C'`, OUTPUT `@ASNID = SCOPE_IDENTITY()`, `@Ein = <next EIN>`, the
   14-char `yyyymmddHHmmss` timestamp into `VC_ADD`/`VC_LAST_UPDATE`. Add `site_id`.
4. **The fan-out (= `CalculateASNFRS`, reimplemented as Gateway Jython, NOT a wrap):**
   - per BC from **`AD_FRSPull`** (cross-DB read — **MUST obtain its body first; the one source gap**):
     get `BC, Orders, VEHICLES`.
   - per BC: `SELECT_ForecastDetailBCASN(BC, effMonth)` → parts + `IN_ASSY_QTY` + `IN_TIRE_RATIO`/
     `IN_WHEEL_RATIO` + `VC_ASSY_MANIFEST_NUMBER`; abort the whole create if any part has no in-window
     manifest cost (`IN_MANIFEST_COST_ID IS NULL`) or the BC has no forecast detail.
   - **branch on `Orders <= 5`:** "(No Ratio)" → one row, `qty = VEHICLES * IN_ASSY_QTY`, break;
     else per-row `qty = (both ratios 100 ? VEHICLES*IN_ASSY_QTY : round(VEHICLES*IN_ASSY_QTY*IN_TIRE_RATIO/100))`.
   - manifest = `'7' + prodDate[4..8] + VC_ASSY_MANIFEST_NUMBER` (1-digit year + MM + DD + 2-char id).
   - `INSERT_ASNDetail(@ASNID, @EIN=nextEIN, manifest, part, qty, @HotCall=0)` — **preserve the
     accumulate-on-repeat upsert**, re-keyed `(site_id, IN_ASN_ID, manifest)`.
5. **Post-loop missing-cost audit** (`SELECT_ASNMissingCost`) — log/surface, do **not** abort (the
   abort already happened pre-insert in step 4).
6. **Advance the EIN counter** (replaces `AD_UpdateEIN`; the legacy's unscoped
   `UPDATE Site SET SiteEIN+1` is a multi-site bug — make it an atomic per-site sequence claim).
7. **Commit.** Put the **whole chain (including EIN allocation/advance)** in one transaction — fixing
   the legacy split where `AD_UpdateEIN` ran on the un-transacted ALC_Connection (EIN-gap bug, §0).

**Parity test:** the 20-detail + 2-No-Ratio fan-out (log rows 4-25, ASNID 4721) must reproduce
row-for-row on the same seq range — requires `AD_FRSPull`'s real output.

---

## 8. Where the live source CONTRADICTS the plan's §2 Rank 1 assumptions

| Plan §2 Rank 1 / §6 assumption | Live source truth | Action |
|---|---|---|
| `CalculateASNFRS` = "stock decrement; cross-module proc, body unverified" | **Delphi method, no decrement.** It is the **ratio fan-out** (build-data × forecast-ratio → detail INSERTs). No FRS-date calc, no inventory decrement in the verified body. | Re-label in the plan. The "stock decrement" line is wrong. |
| EIN "allocated at send, not create" (Q4) | Legacy allocates `SiteEIN+1` onto header **at CREATE**; `AD_UpdateEIN` bumps the counter **inside the create transaction**. Send only flips status. | Decide parity (create) vs improvement (send) with David. |
| "(No Ratio)" = even-division remainder (`daily-workflow §2`) | "(No Ratio)" = **broadcast code with `Orders <= 5`** (small volume), shipped whole, loop `break`. | Correct the workflow doc. |
| Ratio split source possibly `INV_ASSY_RATIO_MST` (D12 dropped) | Split uses **`INV_FORECAST_DETAIL_INF.IN_TIRE_RATIO`** via `SELECT_ForecastDetailBCASN`. `INV_ASSY_RATIO_MST` is **not in the live create path**. | Fan-out is live; D12's table is unrelated/dead. Confirm no other live caller of `INV_ASSY_RATIO_MST`. |
| Manifest = `'7'+yy(2-digit)+MM+DD+id` | `copy(prodDate,4,5)` = **1-digit year** + MM + DD. | Correct to 1-digit year. |
| Transaction wraps the create | Wraps **Inv_Connection only**; `AD_GetSite`/`AD_FRSPull`/`AD_UpdateEIN` run on **ALC_Connection outside it** → EIN-gap on rollback. | Single-transaction orchestration in the rebuild fixes this (a parity *improvement*). |
| `INSERT_ASNDetail`/`DELETE_ASNItem` dedup scope | Both key on **manifest ALONE** (verified). | Re-key `(site_id, IN_ASN_ID, manifest)` per Q1 (already RESOLVED). |

---

## 9. Remaining unknowns the developer MUST resolve before locking M1

1. **`AD_FRSPull` body (BLOCKER for parity).** Not in any dump available here. Pull it from the live
   ALC/VehicleOrder DB. Need: its exact grouping, the definitions of `Orders` and `VEHICLES`, and any
   filtering by `@Start/@Last` seq (the `@begindate/@enddate` range is the active filter elsewhere).
   The entire detail fan-out qty depends on it. **Confidence on the fan-out is HIGH for the Delphi
   logic, but the qty inputs are UNVERIFIED until `AD_FRSPull` is read.**
2. **EIN allocation timing decision** (create vs send) — David, per §2/§8.
3. **`round()` semantics** for the ratio split qty (Delphi round-half-to-even) — confirm against a
   golden ratio-split row so the rebuilt qty matches at the .5 boundary.
4. **The `effMonth` derivation** passed to `SELECT_ForecastDetailBCASN` =
   `copy(fproductiondate,1,4)+'/'+copy(fproductiondate,5,2)` (`DataModule.pas:5154`) = `yyyy/MM`
   (7-char) — confirm forecast-detail rows are stored with that exact `VC_EFFECTIVE_MONTH` format (or
   `''`), else the BC pull returns nothing and the create aborts.
5. **`INV_ASSY_RATIO_MST` truly dead?** Confirmed absent from the create path; verify no other live
   caller before treating D12 as fully closed for ASN.

**Confidence summary:** §1, §2, §4 (Inventory side), §5, §6, §7 are **VERIFIED from read proc/Pascal
bodies**. §3's `AD_FRSPull` is **INFERRED from the call site only** (body genuinely missing) — the one
hard gap. All ALC procs except `AD_FRSPull` were located and read in
`/Users/apple/Documents/FP docs/SQL/VehicleOrder.sql`.
