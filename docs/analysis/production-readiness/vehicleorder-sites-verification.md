# Verify-before-relocate: VehicleOrder.sites readers (2026-06-19)

The plan's M4 pre-task (from the stress-test SHOULD-FIX) required verifying the real `VehicleOrder.sites`
table and its readers **before** anything relocates/retires it. Result below.

## What was checked
- **InventorySystem Delphi source** (`*.pas`/`*.dfm`) — every reference to a `sites` table, to `VehicleOrder`,
  and to the ALC cross-DB connection + `AD_*` procs.
- **The spike SQL Server** — the databases present and the `VehicleOrder` tables/procs.

## Findings

1. **InventorySystem does NOT read any `sites` table — anywhere.** Zero `FROM/JOIN/INTO sites` in the source;
   zero 3-part `VehicleOrder..sites` references. InventorySystem's own site config comes from the **INI**
   (`SiteInfo.pas` / `TSiteInfo`), not a DB table. (Confirms the stress-test reviewer's observation.)

2. **InventorySystem's REAL cross-DB (VehicleOrder/ALC) dependencies are:**
   - **`LINE`** (production-line list: COROLLA/TUNDRA/CAMRY/TACOMA/HIGHLANDER) — heavy use across order/forecast.
   - **`AD_GetSpecialDate` / `AD_GetSpecialDates`** — the production calendar (**stays shared in VehicleOrder**, Q9).
   - **`AD_UpdateEIN`** — the EIN counter (**becomes per-site `INV_SITES.IN_EIN_SEQ`**, Q4).
   - **Reports** via `VehicleOrderConnection` (Admin/reporting side).
   The ALC connection (`fiALCConnection`) targets `TireOrder`; a separate `VehicleOrderConnection` targets
   `VehicleOrder` (both are the ALC/Tire-order cross-app area). Connection strings in `DataModule.dfm` are
   dev-machine artifacts (catalogs `InventoryH`/`Inventory2`/`Activity`/`VehicleOrder`/`TireOrder`).

3. **The real `VehicleOrder.sites` cannot be inspected from this environment.** The spike server's
   `VehicleOrder` DB is a **`LINE`-only stub** (the Order-spike fixture, `scripts/spike-vehicleorder-line-fixture.sql`);
   the production VehicleOrder (with `sites`, `SpecialDate`, `ProductionStatus`, `Line`, the `AD_*` procs) is
   not restored here, and no VehicleOrder schema dump exists in the repo. The sibling apps that likely read
   `VehicleOrder.sites` (GALC, MES, Admin, the Tire-order system) are separate codebases not in this repo.

## Conclusion

- ✅ **Adopting `INV_SITES` is SAFE on the InventorySystem side.** InventorySystem never read `VehicleOrder.sites`,
  so pointing the rebuild at `INV_SITES` breaks no existing InventorySystem reader — it is a **net-new
  authoritative source**, not a migration of an in-use dependency. The "relocate" half of the M4 pre-task is
  effectively already satisfied for InventorySystem.
- ⚠️ **Do NOT retire/drop the production `VehicleOrder.sites`** as part of InventorySystem work. Its real
  external readers (GALC / MES / Admin / Tire-order) are unverifiable here. Physically retiring the shared
  table is a **cross-system decision** for when/if those siblings also migrate — out of scope for the
  InventorySystem rebuild. INV_SITES simply becomes InventorySystem's own site source.
- 🔗 **Cross-DB reconciliation (Q9):** `LINE` and the calendar stay shared in VehicleOrder. `INV_SITES`
  references the shared `VehicleOrder.Line.LineName` for its site↔line mapping (the calendar is keyed by
  LineName) — so a site row in Inventory still points at the shared line list, not a copy.

## Open for David (only if retirement is ever in scope)
To verify the production `VehicleOrder.sites` structure + its GALC/MES/Admin readers, we'd need the **real
VehicleOrder schema dump** or David's confirmation of which apps read it. Until then, treat `VehicleOrder.sites`
as a shared external table that stays put.
