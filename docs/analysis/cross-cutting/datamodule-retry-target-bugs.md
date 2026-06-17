# Defect Register: `DataModule.pas` wrong-target retry recursion (pattern P12)

**Area:** Cross-cutting (legacy defect)  **Status:** ✅ audited & verified  **Analyst:** Claude / 2026-06-05

> **What this is.** Every CRUD method in `DataModule.pas` wraps its ADO call in an
> identical copy-pasted error harness (pattern **P8**): on an exception it does
> `fErrorCount := fErrorCount + 1; If fErrorCount < 3 Then <retry> Else LogActLog('ERROR',…)`.
> The retry is meant to re-call the **same** method. The harness was pasted between
> methods and the retry call was frequently **never renamed**, so it re-invokes a
> *different* method (pattern **P12**). Because many of those wrong targets key on the
> **shared `fRecordID`/`fBroadCode`** singleton state (pattern **P9**), a transient DB
> error can make a retry **read, write, or DELETE the wrong table**.

## How this was found
1. Deterministic scan of `DataModule.pas`: pair every method with the call in its
   `If fErrorCount < 3 Then` branch (CR-stripped, case-insensitive). 79 retry branches;
   **29** call a genuinely different method.
2. Adversarial verification workflow (7 agents, one per method-family): each candidate
   was confirmed against source and the target proc/trigger bodies, then classified by
   blast radius. **Result: 29/29 confirmed real, 0 false positives.**

## Severity tally
| Severity | Count | Meaning |
|---|---|---|
| 🔴 **CRITICAL** | **8** | Wrong-table INSERT/UPDATE/DELETE keyed on shared/stale state → can corrupt or destroy *unrelated* data |
| 🟠 MODERATE | 8 | Wrong write, but keyed on its own (stale) fields and/or dup-guarded, or confined to the same table family → bogus/junk rows, lost intended write |
| 🟡 LOW | 13 | Wrong **SELECT** only → loads the wrong dataset into the shared `Inv_DataSet`; no persistence |

> **Important caveat on realism:** these fire **only on the retry path** — i.e. only when
> the original operation throws a transient exception **and** the wrong target's stale key
> happens to match a real row in the victim table. That rarity is why they have survived in
> production. They are latent landmines, not everyday failures — but the CRITICAL ones are
> silent data loss when they do fire.

## 🔴 The 8 CRITICAL bugs (legacy-hotfix candidates)
Line = the retry-branch call line in `DataModule.pas` (the one to change). "Fix" = repoint the
retry at the **enclosing** method (or remove the in-method recursion).

| # | Method | Line | Wrong target | Victim table | Keyed on | Cascade |
|---|--------|-----:|--------------|--------------|----------|---------|
| 1 | `DeleteManifestCostInfo` | 1804 | `DeleteSupplierInfo` | `INV_SUPPLIER_MST` | shared `fRecordID` (a manifest-cost id) | ✅ live `DELETE_SupplierCode` nulls `INV_PARTS_STOCK_MST.IN_SUPPLIER_ID` |
| 2 | `DeleteMonthlyPOInfo` | 2020 | `DeleteSupplierInfo` | `INV_SUPPLIER_MST` | shared `fRecordID` (stale; method never sets it) | ✅ same trigger cascade |
| 3 | `DeleteRenbanGroupInfo` | 2234 | `DeleteSupplierInfo` | `INV_SUPPLIER_MST` | shared `fRecordID` (a renban id) | ✅ same trigger cascade |
| 4 | `DeleteOvertimeHolidayInfo` | 6651 | `DeleteSupplierInfo` | `INV_SUPPLIER_MST` | shared `fRecordID` (a special-date id) | ✅ same trigger cascade **+ crosses ALC→Inv connection** |
| 5 | `DeleteRecConfStatInfo` | 3091 | `DeleteAssyRatioInfo` | `INV_ASSY_RATIO_MST` | shared `fBroadCode` (never set by this method) | breaks forecast/breakdown math |
| 6 | `UpdateManifestCostInfo` | 1761 | `UpdateSizeInfo` | `INV_SIZE_MST` | shared `fRecordID` reinterpreted as `IN_SIZE_ID` | clobbers a size row's code/usage/safety |
| 7 | `UpdateRenbanGroupInfo` | 2191 | `UpdateSizeInfo` | `INV_SIZE_MST` | shared `fRecordID` (a renban id) as `IN_SIZE_ID` | clobbers a size master row |
| 8 | `UpdateEINStatus` | 6789 | `UpdateRecProdRejInfo` | `INV_REJECT_INF` | shared `fRecordID` as `IN_REJECT_ID` | overwrites a production-reject row; EDI doc left in wrong status |

**The `DeleteSupplierInfo` magnet:** 4 of the 5 wrong-target DELETEs land on
`DeleteSupplierInfo` (it sits where the copy-paste source was). Worst case for each: a
transient error during an unrelated delete → **a real supplier row is deleted** (by an id
borrowed from another table) → the live `DELETE_SupplierCode` trigger then **nulls
`IN_SUPPLIER_ID` on every part of that supplier** in `INV_PARTS_STOCK_MST`, orphaning them
from their supplier. (The trigger's exact live behavior — nulling the int FK, not blanking a
string code — is confirmed in [`trigger-source-reconciliation.md`](trigger-source-reconciliation.md);
the obsolete `docs/triggers.sql` form blanks `VC_SUPPLIER_CODE`, a dropped column.) This is the
single highest-impact path in the unit.

## 🟠 MODERATE (8) — wrong write, constrained
`InsertManifestCostInfo`→`InsertSizeInfo` (1708) · `InsertMonthlyPOInfo`→`InsertSizeInfo` (1910)
· `UpdateMonthlyPOInfo`→`UpdateSizeInfo` (1973, keyed on shared `fRecordID` but a single Size row)
· `InsertRenbanGroupInfo`→`InsertSizeInfo` (2130) · `UpdateRecConfStatRenbanInfo`→`UpdateRecConfStatInfo`
(3325, **same** table `INV_OPEN_ORDER_INF` — broader key-mutating update than intended) ·
`UpdateINVDone`→`UpdateAssyRatioInfo` (4351) · `InsertFirstProductionDayInfo`→`InsertSizeInfo` (6564)
· `InsertOvertimeHolidayInfo`→`InsertSupplierInfo` (6706, **crosses ALC→Inv connection**).
The `Insert*` cases key on their own (stale) `fSize*`/`fSup*` fields and are dup-guarded, so the
realistic damage is a **junk/duplicate master row + a silently lost intended write**, not destruction.

## 🟡 LOW (13) — wrong SELECT, read-only
All load the wrong result set into the shared `Inv_DataSet`/`Inv_StoredProc` on retry; no persistence.
`Get{ManifestCost,MonthlyPO,RenbanGroup,ForecastDetail}Info`→`GetSizeInfo` ·
`GetBCRatioInfo`→`GetAssyRatioInfo` · `Get{PartsListCount,PartsList,NextProductionDate,ExcelPO}Info`→`GetRecConfStatInfo`
· `GetBuildHist`→`GetStocktakingInfo` · `Get{FirstProductionDay,LastProductionDate}Info`→`GetOvertimeHolidayInfo`
· `SelectSingleFieldALC`→`SelectSingleField` (reads via the **Inv** connection instead of **ALC** — wrong
server if they differ; also a `finally` closes the wrong dataset).

## Root cause & the real fix
- **Root cause = three stacked patterns:** P8 (per-method recursive retry boilerplate) ×
  P12 (retry call not renamed on paste) × P9 (a single shared `fRecordID`/`fBroadCode`
  reused as the key for every entity). Remove any one and the CRITICAL class disappears.
- **Rebuild fix (preferred):** delete the in-method recursion entirely. Use **one generic,
  bounded transport-retry wrapper** that re-invokes the *same* operation, and pass record
  keys as **explicit per-call arguments** (never shared mutable singleton state). With
  RESTful routing + ActiveRecord this whole class is structurally impossible.
- **Legacy hotfix (if patching the running Delphi app):** repoint each of the 8 CRITICAL
  retry lines above at its enclosing method (or just drop the `If fErrorCount < 3 Then …`
  retry). The 5 wrong-DELETEs (lines 1804, 2020, 2234, 6651, 3091) are the urgent set.
  Compile/test in the Delphi 7 IDE — cannot be verified from this environment.

## Secondary findings surfaced during the audit
- **`InsertSizeInfo`'s own dup-check is itself broken:** it calls `SELECT_AssyRatioInfo`
  (param `@BroadCode`) with `@SizeCode` — so even *direct* callers of `InsertSizeInfo` have no
  working duplicate guard (already captured in [`../master-data/size.md`](../master-data/size.md)).
- **Cross-connection hazard:** the overtime/holiday writers use `ALC_StoredProc`/`ALC_Connection`,
  but their wrong Supplier/Size targets use `Inv_StoredProc`/`Inv_Connection` — a retry crosses
  databases (#4 and the `InsertOvertimeHoliday` MODERATE case).
- **`SelectSingleFieldALC` `finally` mismatch:** closes `Inv_Field_DataSet`, not the `ALC_DataSet`
  it opened — a pre-existing cleanup bug independent of the retry mis-target.

## Parity / regression checks for the rebuild
- The rebuild must have **no** path where an operation's failure handler invokes a *different*
  entity's persistence. A retry that fails N times must touch only the original target row.
- Assert that no record-id value from one table can be used as another table's key (the absence
  of shared mutable `RecordID` makes this structurally true).
