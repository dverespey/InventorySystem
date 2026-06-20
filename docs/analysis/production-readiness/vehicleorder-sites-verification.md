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

---

## POST-RESTORE CONFIRMATION (2026-06-19, real VehicleOrder backup restored — corrects Finding 1)

David provided the live `VehicleOrder.bak` (1.7G) + a matched live `Inventory.bak`, restored to the spike
(VehicleOrder = real; Inventory.bak → separate `Inventory_Live` for parity, working `Inventory` untouched).

- **The real table is `VehicleOrder.dbo.Site`** (capital-S, **singular**; **2 rows**), NOT `sites`. Its columns
  are **exactly the `TSiteInfo` fields**: `SiteID, SiteName, SiteAbbr, SiteStreet, SiteCity, SiteState,
  SiteCountry, SiteZip, SiteDUNS, SiteSupplierCode, SiteDockCode, SiteEIN, …`.
- **Correction to Finding 1:** InventorySystem DOES read it — **indirectly via the `AD_*` procs**: `AD_GetSite`
  returns the `Site` row (the ASN-create path reads `SiteEIN` from it), and `AD_UpdateEIN` bumps `Site.SiteEIN`.
  (The earlier "reads no sites table" was literally true for a `FROM sites` grep, but missed the proc-mediated
  read.) So **`INV_SITES` is precisely the relocation of `VehicleOrder.Site` into Inventory** — same columns,
  and `INV_SITES.IN_EIN_SEQ` takes over `Site.SiteEIN` (Q4).
- **Readers of `VehicleOrder.Site`:** InventorySystem (via `AD_GetSite`/`AD_UpdateEIN`) **and** the shared
  siblings (GALC/MES/Admin — VehicleOrder is the shared ALC DB per Q9). So the retire-the-shared-copy
  conclusion STANDS: don't drop `VehicleOrder.Site`; InventorySystem repoints to `INV_SITES`, the shared
  table stays for siblings (cross-system decision, the OTHER repo).
- **Unblocked:** `AD_FRSPULL` + the GALC tables (Vehicle 2.33M rows, Model, VehicleData, DataItem, Line,
  SpecialDate, ProductionStatus) are all present → the M1 `create_asn` driver + **end-to-end parity** can now
  run against the matched live pair (VehicleOrder + `Inventory_Live`, max ASN id 4722 ≈ the daily-log 4721).
- **Real DUNS/EIN values:** the real `Site` rows carry actual plant DUNS/EIN — load into the spike `INV_SITES`
  for testing, but KEEP the committed `spike-inv-sites-table.sql` seed as PLACEHOLDERS (don't commit real
  trading identifiers).
