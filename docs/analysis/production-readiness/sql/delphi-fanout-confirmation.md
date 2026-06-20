# ASN-Creation Fan-Out — Delphi Source-Truth Confirmation (for the E2E architecture doc)

**Author:** delphi-architect · **Date:** 2026-06-19
**Verifies/corrects:** `docs/analysis/production-readiness/m1-asn-creation-spec.md`
**Live source read this pass:**
- `DataModule.pas:5106-5402` (`CalculateASNFRS` + `InsertASNInfo`)
- `ASNSelect.pas:369-430` (`CreateASNEntries_ButtonClick` orchestration + Inv_Connection txn + rollback)
- `DataModule.dfm:452-473` (`SiteDataSet`/`AD_GetSite`, `UpdateReportCommand` — both on **ALC_Connection**)
- Procs (Inventory live dump `DB Schema/CreateInventory.sql`, UTF-8 `/tmp/inv_utf8.sql`):
  `INSERT_ASNInfo:2529`, `INSERT_ASNDetail:2682`, `SELECT_ForecastDetailBCASN:3011`, `SELECT_ASNMissingCost:3504`
- Procs (ALC `VehicleOrder.sql`, UTF-8 `/tmp/vo_utf8.sql`):
  `AD_GetSite:592`, `AD_UpdateEIN:623`, **`AD_FRSPULL:4714`**

> **HEADLINE CORRECTION:** the existing spec's "one true source gap / BLOCKER for parity" —
> `AD_FRSPull` body unverified — is **WRONG**. `AD_FRSPULL` is fully present at
> `/Users/apple/Documents/FP docs/SQL/VehicleOrder.sql` (UTF-8 `/tmp/vo_utf8.sql:4714`). Its body is
> read and confirmed below (§F). The fan-out is now **fully verified end-to-end**; there is no
> remaining blocker, and the definitions of `Orders`/`VEHICLES`/`BC` are no longer inferred.

---

## Confirmed fan-out chain (one "Create ASN entries only" click)

All on `ASNSelect.pas:369 CreateASNEntries_ButtonClick`:

1. `Inv_Connection.BeginTrans` — `ASNSelect.pas:372`. **Inv_Connection only.**
2. Stash run params: `BeginDatestr := StartBox.Text`, `EndDateStr := EndBox.Text`,
   `ProductionDate := formatdatetime('yyyymmdd', ASN_DateTimePicker.DateTime)` — `:375-377`.
   (Properties are pure pass-through fields, `DataModule.pas:481-486`; **no transform** of the date strings.)
3. `SiteDataSet.Close/Open` → `EIN := SiteDataSet['SiteEIN'].AsInteger` — `:378-380`.
   `SiteDataSet` = `AD_GetSite` on **ALC_Connection**, `Parameters = <>` so `@LineName` defaults `''`
   → `SELECT * FROM Site` (first/only row). **Outside the Inv txn.**
4. `LineName := Line_ComboBox.Text`; `Quantity := StrToInt(trim(ShipQty_MaskEdit.Text))` — `:381-382`.
5. `InsertASNInfo` (`DataModule.pas:5321`):
   a. `INSERT_ASNInfo;1` on `Inv_StoredProc` (**Inv_Connection**, in txn), `@EIN = fEIN+1`,
      `@ASNID` OUTPUT → `fRecordID := @ASNID` (`:5364`). Header status `'C'`.
   b. on success calls **`CalculateASNFRS`** (`:5381`) — the fan-out.
6. `CalculateASNFRS` (`DataModule.pas:5106`): per BC from `AD_FRSPULL` → `SELECT_ForecastDetailBCASN`
   → per detail row `INSERT_ASNDetail;1` (**Inv_Connection**, in txn) → post-loop `SELECT_ASNMissingCost` warn.
7. Back in `CreateASNEntries_ButtonClick`, **only if `InsertASNInfo` returned True**:
   `UpdateReportCommand.CommandText := 'AD_UpdateEIN'; .Execute` — `:387-389`.
   `UpdateReportCommand` is on **ALC_Connection** (`DataModule.dfm:469`) → **outside the Inv txn.**
8. `Inv_Connection.CommitTrans` — `:391`.
9. On `InsertASNInfo=False` or any exception: `if InTransaction then RollbackTrans` — `:417-418, :426-427`.

---

## (a) Orders<=5 "No Ratio" branch — CONFIRMED

`DataModule.pas:5183-5213`. Inside the per-detail-row `while not Inv_StoredProc.eof` loop, the **very
first** check is `if ALC_StoredProc.FieldByName('Orders').AsInteger <= 5 then`. The body:
- `manifest := '7'+copy(fProductionDate,4,5)+VC_ASSY_MANIFEST_NUMBER` (`:5186`)
- `count := FieldByName('VEHICLES').AsInteger * IN_ASSY_QTY` — **no ratio** (`:5187`)
- one `INSERT_ASNDetail` (`:5189-5203`), log `INSERT ASN entry(No Ratio)` (`:5211`)
- **`break;`** (`:5212`) — exits the detail-row loop after the FIRST forecast-detail row.

`break` is inside `while not Inv_StoredProc.eof`, so it emits **exactly one** detail row for that BC
from its FIRST forecast-detail row, then advances to the next BC (`next;` at `:5277`). **Spec §4(A) correct.**

> **NEW, load-bearing nuance the spec misses — `Orders` is branch-dependent in `AD_FRSPULL` (§F):**
> in the GROUND tire/wheel half of the UNION `ORDERS = COUNT(*)*4`; in the SPARE tire half
> `ORDERS = COUNT(*)`. So the `Orders<=5` test does **not** mean "≤5 vehicles" uniformly:
> a ground-tire BC trips No-Ratio only when `COUNT(*)*4 <= 5`, i.e. **`VEHICLES = 1`** (4≤5; 2→8>5);
> a spare BC trips it when `VEHICLES <= 5`. The existing spec's prose "broadcast code with ≤5 vehicles"
> is only true for the spare branch. **Data-dependent — confirm against the golden:** the two No-Ratio
> manifests (`76061836`, `76061851`) should be either ground-tire BCs with exactly 1 vehicle or spare BCs
> with ≤5 vehicles. Check `AD_FRSPULL`'s `ORDERS`/`VEHICLES` for those two BCs on the golden seq range.

## (b) Ratio branch qty — CONFIRMED

`DataModule.pas:5214-5265` (the `else`, i.e. `Orders > 5`), per forecast-detail row:
- both ratios 100: `count := FieldByName('VEHICLES').AsInteger * IN_ASSY_QTY` (`:5226-5230`) — full qty.
- else: `count := round((VEHICLES * IN_ASSY_QTY * IN_TIRE_RATIO) / 100)` (`:5234-5235`).
  Uses **`IN_TIRE_RATIO` only**; `IN_WHEEL_RATIO` participates only in the `=100` guard, never the multiply
  (code comment `:5222-5223`: tire/wheel share set equal for forecast). One `INSERT_ASNDetail` per row,
  **no break** → ~20 rows. **Spec §4(B) correct.**
- **`round` is Delphi banker's rounding (round-half-to-even).** Still a real fidelity risk at `.5`.
  **Data-dependent — confirm against golden:** any ratio-split row whose
  `VEHICLES*IN_ASSY_QTY*IN_TIRE_RATIO/100` lands on exactly `x.5` — verify the legacy `IN_QTY` matches
  half-to-even (e.g. 2.5→2, 3.5→4), not away-from-zero.

## (c) Manifest = '7' + copy(prodDate,4,5) + VC_ASSY_MANIFEST_NUMBER — CONFIRMED

`DataModule.pas:5186` (No-Ratio) and `:5239` (ratio) — byte-identical. `fProductionDate` = `yyyymmdd`;
`copy(s,4,5)` = chars 4..8 = **last digit of year + MM + DD** (1-digit year). For `20260618` →
`'7'+'60618'+id`. **Spec §6 correct; the "1-digit year" correction stands.** Data-dependent: a 2027 date
yields `'7'+'70618'+id` (year digit `7`), so an FY-rollover collision with a 2017 same-MMDD manifest is
latent — note for multi-year retention but not in scope here.

## (d) Pre-insert RAISE vs post-loop SELECT_ASNMissingCost warn — CONFIRMED (with a scope refinement)

Two distinct mechanisms, both verified:

- **Pre-insert hard abort (per BC, NOT once-per-whole-run):** `DataModule.pas:5160-5175`. After
  `SELECT_ForecastDetailBCASN` returns rows for a BC, a first pass scans **all** detail rows of *that BC*
  and accumulates `errorstr` from any with `IN_MANIFEST_COST_ID IS NULL` (the LEFT JOIN miss in
  `SELECT_ForecastDetailBCASN`). If `errorstr <> ''` → `raise EDatabaseError` "Missing Manifest Cost
  Information BCode(..)" (`:5170-5174`) **before any `INSERT_ASNDetail` for that BC**. The raise unwinds to
  the outer `try` (`:5312`), `CalculateASNFRS` returns **False**, `InsertASNInfo` returns False, and
  `CreateASNEntries_ButtonClick` **rolls back the whole Inv txn** (`ASNSelect.pas:417`). So one missing
  cost on any assy part aborts the **entire** create.
  > **Refinement to spec §3:** describe this as **per-BC, pre-insert**, evaluated each iteration of the
  > BC loop — not "pre-loop" (it runs after the first BC may already have inserted rows; those inserts are
  > undone only by the transaction rollback, not by the raise itself). The net effect is still all-or-nothing
  > because everything is one Inv txn.

- **Post-loop read-only warn:** `DataModule.pas:5285-5308`. After all BCs, `SELECT_ASNMissingCost(@ASNID)`
  (`/tmp/inv_utf8.sql:3504`) re-checks every inserted detail row against an **in-window** manifest cost
  (`m.VC_START_MANIFEST <= prodDate <= m.VC_END_MANIFEST`), distinguishing "Missing Manifest Cost Entry"
  (no cost at all) vs "out of date" (cost exists but window doesn't cover prodDate). It only
  `ShowMessage`/`LogActLog('ERROR',...)` and is wrapped in its own `try/except` that swallows exceptions
  (`:5303-5307`) — **never aborts, never rolls back.** **Spec §3/§4(2d) correct.**
  Note the asymmetry: the pre-insert raise keys on `IN_MANIFEST_COST_ID IS NULL` (cost row exists at all),
  the post-loop warn keys on the date-window match — so a cost that exists but is out-of-window passes the
  abort and only triggers the warn. Faithful, intentional.

## (e) EIN read (AD_GetSite.SiteEIN) + AD_UpdateEIN bump OUTSIDE the Inv txn — CONFIRMED (collision/gap bug)

- Read: `EIN := SiteDataSet['SiteEIN']` via `AD_GetSite` on **ALC_Connection** (`ASNSelect.pas:380`,
  `DataModule.dfm:453`). Header + all details written with `@EIN = fEIN+1`
  (`DataModule.pas:5359, 5196, 5248`). `AD_GetSite` body (`/tmp/vo_utf8.sql:592`): `@LineName=''` →
  `SELECT * FROM Site` (all rows; single-site today returns one).
- Bump: `AD_UpdateEIN` via `UpdateReportCommand` on **ALC_Connection** (`ASNSelect.pas:387-389`,
  `DataModule.dfm:469`). Body (`/tmp/vo_utf8.sql:623`): `UPDATE Site SET SiteEIN = SiteEIN+1` —
  **no WHERE** → bumps **every** Site row.
- **Both the read and the bump run on ALC_Connection, OUTSIDE the `Inv_Connection` BeginTrans/CommitTrans.**
  The bump executes between `InsertASNInfo=True` and `Inv_Connection.CommitTrans` (`:387` then `:391`).

  Two real bugs the rebuild fixes:
  1. **EIN gap / no rollback coupling:** an `Inv_Connection.RollbackTrans` (`:417`) does **not** revert the
     ALC-side `SiteEIN+1`. But note the *ordering* limits this: `AD_UpdateEIN` only fires after
     `InsertASNInfo` succeeded, so a fan-out failure rolls back the Inv side **without** bumping EIN (good).
     The genuine gap is the reverse window: if `CommitTrans` (`:391`) fails *after* `AD_UpdateEIN` ran, the
     counter advanced but no ASN persisted → EIN gap. Plus the unscoped UPDATE.
  2. **Read-then-bump race (no atomicity):** `AD_GetSite` read and `AD_UpdateEIN` bump are two separate,
     unscoped, un-transacted ALC round-trips with the Inv work between them. Two concurrent creates can read
     the same `SiteEIN` and both ship `fEIN+1` → duplicate EIN (collision). The unscoped UPDATE also makes
     the value meaningless per-site once multi-site lands.
  > **Rebuild fix (parity-improving):** allocate the EIN with an **atomic, per-site, in-transaction**
  > sequence claim (single UPDATE...OUTPUT or sequence) **at send**, inside the same Gateway transaction as
  > the inserts — removing both the gap and the read-then-bump race, and scoping to `site_id`. The existing
  > spec §2/§8 already flags the create-vs-send timing decision; this confirms the *mechanism* is a
  > two-round-trip non-atomic ALC read+UPDATE, which is the concrete thing being replaced.

## (f) AD_FRSPULL begindate/enddate window derivation — CONFIRMED (and the proc body, formerly "missing")

> ⚠️ **CORRECTION (orchestrator-adjudicated against the LIVE DB):** the `char(21)` BC + `ModelYearCode +
> tire + wheel` concat reading below was taken from the STALE on-disk dump `/Users/apple/Documents/FP docs/
> SQL/VehicleOrder.sql` (Script Date 06/10/2026). The **LIVE running proc** on David's restored backup (the
> DB the rebuild reads) is **`char(3)`, ground concat `ModelYearCode + WHEEL(vd2) + TIRE(vd1)`, WITH the spare
> `<> 'M'` filter** — see `AD_FRSPULL-analysis.md` and the authoritative `m1-asn-creation-architecture.md`
> version-conflict box. "Live wins over dump." Treat the char(21)/reversed-concat reading in this section as
> the stale dump's, NOT the production proc.

- `@begindate := fbegindatestr` and `@enddate := fenddatestr` (`DataModule.pas:5128, 5130`).
- `fbegindatestr := StartBox.Text`, `fenddatestr := EndBox.Text` (`ASNSelect.pas:375-376`), via the
  pass-through `BeginDatestr`/`EndDateStr` properties (`DataModule.pas:485-486`). **No transformation.**
  These are the **operator-selected start-time and end-time** strings from the `StartBox`/`EndBox` datetime
  combos (the same window already used by `AD_ProductionSeq` to count Q during Check). So the
  `AD_FRSPULL` window is **purely the chosen [start-time, end-time] datetime range on the line** — `@Start`
  and `@Last` (seq numbers) are passed but, per the proc body, **unused** (see below).
- **`AD_FRSPULL` body (`/tmp/vo_utf8.sql:4714`) — verified:** a UNION of two grouped selects over
  `Vehicle v JOIN Model m JOIN Line l JOIN VehicleData (×2) JOIN DataItem (×2)`, filtered
  `v.DateCreated >= @begindate AND v.DateCreated <= @enddate AND l.LineName = @LineName`:
  - **Half 1 (ground):** `DataItem='GROUNDTIRE'` + `DataItem='GROUNDWHEEL'`;
    `BC = convert(char(21), ModelYearCode + groundtire.DataValue + groundwheel.DataValue)`;
    `ORDERS = COUNT(*)*4`, `VEHICLES = COUNT(*)`, grouped by that BC.
  - **Half 2 (spare):** `DataItem='SPARETIRE'` (the spare-wheel join is commented out);
    `BC = convert(char(21), ModelYearCode + sparetire.DataValue)`;
    `ORDERS = COUNT(*)`, `VEHICLES = COUNT(*)`, grouped by that BC. `ORDER BY BC`.
  - `@Start`/`@Last` (seq) are **accepted but not referenced** — the active filter is the datetime range +
    line (same pattern as `AD_ProductionSeq`, whose seq bounds are also commented out). **DateCreated bounds
    are inclusive on both ends** (`>=` and `<=`).

  Implications for the rebuild fan-out (now fully specifiable, no gap):
  - `BC` is a **21-char fixed-width** string (`convert(char(21), ...)`) = ModelYearCode + tire/wheel
    DataValue codes; it's matched in `SELECT_ForecastDetailBCASN` via `@BCode LIKE VC_BROADCAST_CODE`
    (the forecast row holds the pattern). Trailing-space behavior of `char(21)` matters for the `LIKE` —
    **data-dependent: confirm `VC_BROADCAST_CODE` patterns account for the char(21) right-padding** (else
    a BC that should match returns no forecast rows → "Missing Broadcast Code Information" abort `:5273`).
  - `VEHICLES` (the qty multiplier) = vehicle count per BC in the window. `ORDERS` is the No-Ratio trigger
    and is `COUNT(*)*4` for ground BCs, `COUNT(*)` for spare BCs (see §a).

---

## Spec corrections (things in `m1-asn-creation-spec.md` that are WRONG vs live source)

| Spec location | Claim | Live truth | Severity |
|---|---|---|---|
| §0/§3/§7/§9 (and the §"single biggest correction" callout) | `AD_FRSPull` body **not in any dump** — "the one true source gap," **BLOCKER for M1 parity**, qty inputs UNVERIFIED | `AD_FRSPULL` is in `/Users/apple/Documents/FP docs/SQL/VehicleOrder.sql` (`/tmp/vo_utf8.sql:4714`); body read & confirmed. **No gap, no blocker.** | **HIGH — remove the blocker; the spec's central caveat is false** |
| §4 / §8 row "(No Ratio)" | "(No Ratio)" = "broadcast code with ≤5 **vehicles**" | `Orders<=5` where `Orders` is branch-dependent: ground `=VEHICLES*4` (so ≤5 ⇒ exactly **1** vehicle), spare `=VEHICLES`. "≤5 vehicles" is only the spare case. | **MED — the threshold semantics are mis-stated for ground BCs** |
| §3 / §4(missing-cost) | pre-insert manifest-cost raise framed as "pre-loop" | It is **per-BC, pre-insert** (inside the BC loop, before that BC's detail INSERTs). Earlier BCs may have inserted; only the txn rollback undoes them. | **LOW — net all-or-nothing holds; wording refinement** |
| §9 #1 "Confidence ... qty inputs UNVERIFIED until AD_FRSPull is read" | listed as remaining unknown / DEVELOPER ACTION blocking M1 | Resolved — `AD_FRSPULL` read; `Orders`/`VEHICLES`/`BC` definitions are now VERIFIED facts. | **HIGH — close the open item** |

**Everything else in the spec is confirmed accurate** against the live source this pass: the ordered
chain (§0), `INSERT_ASNInfo` header insert + status `'C'` + SCOPE_IDENTITY (§2), the ratio math (§4),
`INSERT_ASNDetail` manifest-keyed accumulate upsert + `@HotCall` (§5), the manifest scheme (§6, 1-digit
year), the txn-boundary hazard / EIN-gap analysis (§0/§8), and the dedup re-key targets (§5). The
`effMonth` passed to `SELECT_ForecastDetailBCASN` is `copy(fproductiondate,1,4)+'/'+copy(...,5,2)` =
`yyyy/MM` (`DataModule.pas:5154`), confirmed.

## Residual data-dependent items to confirm against the golden (not source gaps)
1. The two No-Ratio manifests (`76061836`, `76061851`): are they ground BCs with `VEHICLES=1` or spare BCs
   with `VEHICLES<=5`? (Check `AD_FRSPULL.ORDERS/VEHICLES` for those BCs on the golden seq range.)
2. A ratio-split row landing on exactly `x.5` → confirm legacy `IN_QTY` follows Delphi round-half-to-even.
3. `char(21)` right-padding of `BC` vs the `VC_BROADCAST_CODE` LIKE patterns in `INV_FORECAST_DETAIL_INF`.
