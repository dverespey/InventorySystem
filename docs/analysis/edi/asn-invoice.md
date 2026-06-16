# Module Analysis: EDI Outbound — ASN (856) + Invoice (810)

> Area §5 EDI/Integration. Covers the outbound X12 generation surface:
> `ASNSelect`, `ASNInvoice`, `EDI856Object`, `EDI810Object`, `HotCallEntry`,
> `InvoiceBreakdown`, and the dead `Write810File`.

**Area:** EDI/Integration  **Status:** ✅ spec complete  **Analyst:** Claude / 2026-06-16

---

## 1. Legacy surface

- **Forms / units (all live in `InventorySystem.dpr`):**
  - `ASNSelect.pas` + `.dfm` (16 KB) — pick line + sequence range → create the ASN (856) + file.
  - `ASNInvoice.pas` + `.dfm` (36 KB) — the ASN/Invoice browser/editor; recreate 810/856 files;
    unsend; delete; manual ASN-detail edit. The hub screen.
  - `EDI856Object.pas` (12.6 KB) — `T856EDI` builder: emits the 856 ASN X12 segment stream.
  - `EDI810Object.pas` (12.6 KB) — `T810EDI` builder: emits the 810 invoice X12 segment stream.
  - `HotCallEntry.pas` + `.dfm` (10 KB) — manual non-broadcast ("hot call") ASN entry.
  - `InvoiceBreakdown.pas` + `.dfm` (9 KB) — fixed-width supplier-invoice text → `INV_INVOICE_INF`.
- **DEAD CODE — `Write810File.pas`:** listed in `InventorySystem.dpr:57`, so it *compiles*, but the
  entire body of `T810EDIFile.Execute` is commented out — a `{` opens at `Write810File.pas:35` and the
  `}` closes at `:118`. The compiled method runs only the empty `else` branch / `try…except` shell. It
  performs **no work**. Treat as dead: the live 810 file-write path is `ASNInvoice.RecreateFile_ButtonClick`
  (`ASNInvoice.pas:857-895`), not this unit. Confidence: high (read in full).

- **Entry points (from `MainMenu.pas`):** the EDI menu opens `ASNSelect` (create ASN/856),
  `ASNInvoice` (browse/recreate/unsend 810 & 856), and `HotCallEntry`. `InvoiceBreakdown` is invoked
  by the breakdown dispatcher for a fixed-width supplier-invoice file (not an X12 path).

- **Purpose:** Generate and ship the two outbound Toyota-TEMA EDI documents. The **856 ASN** declares
  what is on the truck for a production-date / sequence window; the **810 invoice** bills for the parts
  on accepted ASNs. Both are written as X12 text files into the EDI-out directory; an external mailer
  (out of this app) transmits them. The 997 ACK flows back through `EDIUpload` (see `edi-upload.md`).

---

## 2. Data touched

| Table | Read | Write | Notes |
|-------|:----:|:-----:|-------|
| `INV_ASN_MST` | ✓ | ✓ | ASN header. `VC_ASN_STATUS` lifecycle `C`→`S`→`A`/`R`. `IN_ASN_EIN`, `VC_PRODUCTION_DATE`, seq range. |
| `INV_ASN_DETAIL_MST` | ✓ | ✓ | ASN line = manifest+assy-part+qty. `IN_INV_ID` (NULL until invoiced) links it to an 810. |
| `INV_INV_MST` | ✓ | ✓ | Invoice (810) header. `VC_INV_STATUS` `S`→`A`/`R` (created `'S'`); `IN_INV_EIN`. |
| `INV_MANIFEST_COST_MST` | ✓ |  | Assembly price. **Joined on assy code; window honored by 856, IGNORED by 810/SELECT_INVOICEItems — see §4 / D6.** |
| `INV_FORECAST_DETAIL_INF` | ✓ |  | Supplies `VC_ASSY_KANBAN_NUMBER` for the 856 LIN segment; feeds HotCall assy-code combos. |
| `INV_INVOICE_INF` |  | ✓ | **Separate** supplier-invoice ledger written ONLY by `InvoiceBreakdown` (fixed-width text). NOT the 810 path. |
| `AD_GetSite` / `AD_GetSiteTMMDUNS` result | ✓ |  | **On `ALC_Connection`, not the inventory DB.** Source of EIN, DUNS, separators, dock code, EDI mode. See §3 + §4 multi-site. |

**Triggers:** none of these tables carry triggers relevant to EDI generation (the stock-consumption
side-effect happens via `CalculateASNFRS`, owned by the Stock/Forecasting area, not a trigger). Body
of `CalculateASNFRS` unverified here — flagged as a cross-module boundary in §4.

---

## 3. Stored procedures used

| Proc | Operation | Business rule (from body) |
|------|-----------|---------------------------|
| `INSERT_ASNInfo` (`schema:2822`) | INSERT (OUTPUT `@ASNID`) | Inserts `INV_ASN_MST` with status `'C'` (created). 16-char `yyyymmddHHMMSSff` stamp. Returns `SCOPE_IDENTITY()`. |
| `INSERT_ASNDetail` (`schema:2762`) | INSERT/UPDATE | **Dedup keyed on `VC_MANIFEST_NUMBER` ALONE, globally (no ASN scoping).** If `@HotCall=0` and a row with that manifest exists anywhere, it **adds to `IN_QTY`** instead of inserting (`schema:2784-2792`). If `@HotCall=1` it always inserts (`:2796`). See §4 hazard. |
| `UPDATE_ASNItem` (`schema:8158`) | UPDATE | Sets `IN_QTY=@Qty` on one detail row by `IN_ASN_DETAIL_ID`. |
| `DELETE_ASNItem` (`schema:2006`) | DELETE | **Deletes by `VC_MANIFEST_NUMBER` globally** — same un-scoped key as the dedup. Can hit detail rows under other ASNs that share the manifest string. |
| `DELETE_ASNList` (`schema:2030`) | DELETE | Deletes the `INV_ASN_MST` header by id **only — leaves `INV_ASN_DETAIL_MST` children orphaned** (no cascade). D3 concern. |
| `UPDATE_ASNStatus` (`schema:8186`) | UPDATE | If `@ASNStatus='S'`, flips **ALL** `INV_ASN_MST` rows currently `'C'` → `'S'` (no id filter). Used at send time. |
| `UPDATE_ASNUnsend` (`schema:8223`) | UPDATE | Resets one ASN `'S'`→`'C'` by id (un-send). |
| `INSERT_INVInfo` (`schema:3129`) | INSERT (OUTPUT `@INVID`) | Inserts `INV_INV_MST` status `'S'`; then the wrapper calls `UPDATE_INVItems`. |
| `UPDATE_INVItems` (`schema:8469`) | UPDATE | Stamps `IN_INV_ID=@INVID` onto every `INV_ASN_DETAIL_MST` row of `@ASNID` where `IN_INV_ID IS NULL` — links the ASN's lines to the new invoice. |
| `UPDATE_INVUnsend` (`schema:8526`) | UPDATE+DELETE | Nulls `IN_INV_ID` on the detail rows, then **DELETEs the `INV_INV_MST` header**. (The `UPDATE…SET 'C'` line is commented out — it hard-deletes the invoice instead.) |
| `UPDATE_EINStatus` (`schema:8374`) | UPDATE | 997-ACK landing: `@EINType='SH'` → `INV_ASN_MST.VC_ASN_STATUS=@status WHERE IN_ASN_EIN=@EIN`; else `INV_INV_MST.VC_INV_STATUS WHERE IN_INV_EIN=@EIN`. Called from `EDIUpload`. |
| `REPORT_EDI810` (`schema:4303`) | SELECT | 810 line source. **Window-blind** (see §4 / D6). Filters `a.VC_ASN_STATUS='A' AND d.IN_INV_ID IS NULL`. |
| `REPORT_EDI810Recreate` (`schema:4338`) | SELECT | Recreate variant; joins `INV_INV_MST` where status `'C'`. **Also window-blind** (assy-code join only, `:4356-4357`). *(Note: caller `ASNInvoice` actually uses `REPORT_EDI810` with `@EIN`, not this proc — see §4.)* |
| `REPORT_EDI856` (`schema:4377`) | SELECT | 856 line source. **DOES filter the price window** (`schema:4396-4397`) but with strict `<`/`>` string compares — see §4. Filters `a.VC_ASN_STATUS='C'`. |
| `SELECT_INVOICEItems` (`schema:6429`) | SELECT | Invoice-detail display + line total `m.MO_PRICE*d.IN_QTY`. **Window-blind** (assy-code join only, `:6442-6443`). |
| `SELECT_ManifestCost` | SELECT | Feeds `ASNInvoice.AssyManifest_DataSet` (manifest-number generation combo). Body unverified. |
| `INSERTUPDATE_Invoice` (`schema:2697`) | INSERT/UPDATE | `InvoiceBreakdown` path → `INV_INVOICE_INF`. Dedup on `(VC_SUPPLIER_CODE, VC_PART_NUMBER, VC_FRS_NUMBER, VC_RENBAN_NUMBER)`. **Separate subsystem from the 810.** |
| `AD_GetSite`, `AD_GetSiteTMMDUNS` | SELECT | **Not in the inventory schema** (run on `ALC_Connection`); supply EIN/DUNS/separators/dock-code/EDI-mode. Bodies unverified (live in the ALC DB). |
| `AD_UpdateEIN` | EXEC | Called from `ASNSelect.pas:388`, `HotCallEntry.pas:291`, `Write810File.pas:74`. **NOT present in `Create Inventory.sql`** — runs on `ALC_Connection` (`UpdateReportCommand.Connection=ALC_Connection`), so it lives in the ALC DB. Body unverified; treat like GALC's `VO_DeleteVehicleMove` — confirm it exists in the ALC schema before relying on it. |

> **Snapshot-drift note** (cf. [[reference-schema-snapshot-vs-live]]): `AD_GetSite`,
> `AD_GetSiteTMMDUNS`, `AD_UpdateEIN` are referenced by live code but absent from the checked-in
> inventory snapshot **because they live in the separate ALC database**, not because they're missing in
> prod. Verify against the live ALC schema; do not report as a prod bug.

---

## 4. Business rules & edge cases

### 4.1 The 810 line price is window-blind — confirmed buggy (D6)
The 810 line price comes from `REPORT_EDI810`:
```
JOIN INV_MANIFEST_COST_MST m
ON d.VC_ASSY_PART_NUMBER = m.VC_ASSY_PART_NUMBER_CODE   -- schema:4317-4318
```
There is **no `start_manifest`/`end_manifest` window predicate** — the join matches on assy code only.
`SELECT_INVOICEItems` (the invoice-display + total) has the identical defect (`schema:6442-6443`) and
computes the billed total as `m.MO_PRICE*d.IN_QTY` (`schema:6437`). `REPORT_EDI810Recreate` is window-blind
too (`schema:4356-4357`).

Per **D6**: assembly prices are **genuinely time-bounded** and these procs are **confirmed buggy**.
Effect: if `INV_MANIFEST_COST_MST` holds >1 price row for an assy code (a price change over time), the
assy-code-only join either (a) **picks an arbitrary/wrong-window price**, or (b) **duplicates the 810/invoice
line** (one per matching cost row), doubling the bill. The rebuild must select the cost row whose
`[VC_START_MANIFEST, VC_END_MANIFEST]` window **contains `a.VC_PRODUCTION_DATE`** (the ASN production date),
and enforce D6's non-overlapping-window constraint per `(site, assy code)` so exactly one row matches.

> **Contrast — the 856 is NOT window-blind:** `REPORT_EDI856` *does* filter the window
> (`schema:4396-4397`):
> ```
> AND m.VC_START_MANIFEST < a.VC_PRODUCTION_DATE
> AND m.VC_END_MANIFEST   > a.VC_PRODUCTION_DATE
> ```
> Two caveats to preserve/fix: (1) these are **strict** `<`/`>`, so a production date **equal to** a window
> boundary is **excluded** → the row drops out of the 856 (an ASN line silently vanishes). (2) These are
> **string** comparisons on `varchar(8)` `yyyymmdd` — correct only because the format sorts
> lexicographically. The rebuild should use a `>=`/`<=` (inclusive) date-typed window, and apply the *same*
> window logic to the 810/invoice path so 810 and 856 agree on price. (D6 candidate detail: confirm whether
> boundaries are inclusive.)

### 4.2 856 ASN X12 build (`EDI856Object.T856EDI`) — segment map
Driven row-by-row over `EDI856DataSet` (= `REPORT_EDI856`). Element separator = `SiteSepElement` (`*`),
sub-element = `SiteSepSubElement`, all from `AD_GetSite`. Segments emitted in `Execute` order
(`EDI856Object.pas:116-126`):
- **ISA** (`:140`): `ISA*00*<SiteAbbr,%-10s>*01*<SiteDUNS,%-10s>*ZZ*<DUNS-SupplierCode,%-15s>*01*<SiteTMMDUNS,%-15s>*<copy(PickupDate,3,6)>*<hhmm>*U*00400*<%9.9d EIN>*0*<SiteEDIMode>*<subElem>`.
  Date field at `:154` uses `copy(fPickupDate,3,6)` (drops century → `yymmdd`).
- **GS** (`:174`): `GS*SH*<SiteDUNS>*<SiteTMMDUNS>*<fPickupDate full>*<hhmm>*<%9.9d EIN>*X*004010`.
- **ST** (`:199`): `ST*856*<%9.9d EIN>` (transaction-set control = the EIN; not a 1-based counter).
- **BSN** (`:219`): `BSN*00*<fPickupDate + %9.9d EIN  (shipment id)>*<fPickupDate>*<hhmm>`.
- **DTM** (`:242`): `DTM*011*<fPickupDate>*<hhmm>*ET`.
- **HL loop** (`:265`): a Shipment HL (`S`), then per distinct **manifest** an Order HL (`O`) + **PRF**
  (`PRF*<Manifest>-<Manifest>`), then per detail an Item HL (`I`) + **LIN** (`LIN**BP*<PartNumber>*RC*<Kanban>`)
  + **SN1** (`SN1**<ShipQty>*PC`). The static `TD5` carrier uses `SiteDeliveryMethodCode`; `TD3` truck id is
  the literal `'1234567890'` (`:308` — hardcoded). HL parent id is hardcoded `'1'` (`:322`), not `fParent`.
- **CTT** (`:380`): `CTT*<fHLCount>`. **SE** (`:400`): `SE*<fSegCount>*<%9.9d EIN>`. **GE** (`:421`):
  `GE*1*<%9.9d EIN>`. **IEA** (`:441`): `IEA*1*<%9.9d EIN>`.
- **Segment-count caveat:** `fSegCount` (the SE count) is incremented inconsistently — ST/BSN/DTM each
  `INC`, the HL loop `INC`s only inside the final emit loop (`:367`), CTT/SE `INC`. Verify the SE01 count
  matches the true segment count in a real file; an off-by-N here is a TEMA reject.

### 4.3 810 invoice X12 build (`EDI810Object.T810EDI`) — segment map
Driven over `EDI810DataSet` (= `REPORT_EDI810`). Segments (`EDI810Object.pas:112-120`):
- **ISA** (`:135`): same shape as 856 ISA but ISA09 date = `formatdatetime('yymmdd', now)` (`:149`),
  not the pickup date. EIN from the passed-in `fein` (the caller sets `EDI810.EIN := <EIN Number>`).
- **GS** (`:175`): `GS*IN*<SiteDUNS>*<SiteTMMDUNS>*<yyyymmdd now>*<hhmm>*<%9.9d EIN>*X*004010`.
- **ST** (`:200`): `ST*810*<%9.9d EIN>`.
- **BIG** (`:220`): `BIG*<yyyymmdd now>*<SiteSupplierCode>` (invoice header / date+number).
- **IT1 loop** (`:241`): per pickup-date break (new pickup date ⇒ new invoice file, `:263`), per
  **manifest** break emit **REF** (`REF*MK*<lastManifest>`) + **DTM** (`DTM*050*<PickUpDate>`), then the
  **IT1** line: `IT1*<M391 if manifest starts with '7' else M390>*<ShipQty>*EA*<UnitPrice>*QT*PN*<PartNumber>*PK*1*ZZ*<SiteDockCode>`.
  `M391` = broadcast part, `M390` = non-broadcast / hot-call (`:291-294`). Running `total` accumulates
  `UnitPrice*ShipQty`.
- **TDS** (`:337`): invoice total. **Hand-rolled money formatting** (`:339-350`): splits `FloatToStr(total)`
  on `'.'`, right-pads the fraction to 4 digits, concatenates as `<int><4-digit-frac>` with **no decimal
  point** (X12 implied-decimal). Locale-sensitive (`FloatToStr` uses the OS decimal separator) and breaks
  if the fraction has >4 digits — a fragility to replace with explicit integer-cents math.
- **CTT** (`:371`): `CTT*<fID1Count>` (line count). **SE** (`:391`): `SE*<fSECount>*<%9.9d EIN>`.
  **GE/IEA** as 856. Same `fSECount` correctness caveat as 4.2.

### 4.4 ASN creation / send / invoice lifecycle
- **Create ASN** (`ASNSelect.CreateASN_ButtonClick`, `:432`): inside an `Inv_Connection` transaction —
  `InsertASNInfo` (status `'C'`) → `CalculateASNFRS` (decrements stock — **cross-module, body unverified,
  owned by Stock/Forecasting**) → build 856 via `T856EDI` → write file `856<copy(ProductionDate,4,5)>.txt`
  → `AD_UpdateEIN` → `UpdateASNStatus('S')` → commit. `CreateASNEntries_Button` is the same minus the file.
- **Status model:** ASN `'C'` created → `'S'` sent → `'A'` accepted / `'R'` rejected (set by 997 ACK via
  `UPDATE_EINStatus`). Invoice (`INV_INV_MST`) created `'S'`, then `'A'`/`'R'` via ACK; `'C'` appears only as
  a recreate filter (`REPORT_EDI810Recreate`). The `ASNStatus_ComboBox` index→`@List` letter map
  (`ASNInvoice.pas`): 0=`X`(all) / 1=`C`(not-created/editable) / 2=`S`(sent) / 3=`A`(accepted) / 4=`R`(rejected).
- **810 generation** is **EIN-driven on recreate** (`ASNInvoice.pas:857-895`): selects `REPORT_EDI810` by
  `@EIN`, builds via `T810EDI`, writes `810<copy(PickUpDate,5,4)>.txt`.
  ✅ **RESOLVED (D9) vs the live dump:** both `REPORT_EDI810` AND `REPORT_EDI856` **DO** declare `@EIN INTEGER=0`
  and use it — `IF @EIN=0` returns all rows, else `WHERE …_EIN=@EIN` scopes to that site's EIN. Passing `@EIN`
  is correct; the earlier "param-less, cross-site-bleed" concern was a **stale-snapshot artifact** (the 6/1
  snapshot showed these procs param-less; the live 6/12 dump has them EIN-aware). NOT a D1 blocker.
  ⚠️ **NEW finding (D9) — hardcoded site EIN in `REPORT_EDI856`:** one branch has `WHERE a.IN_ASN_EIN = 6440`
  (live `CreateInventory.sql:3683`) — a literal site EIN baked into the proc. The rebuild MUST parameterize
  this (it pins one site; a D1 hazard).
- **Unsend:** ASN `UPDATE_ASNUnsend` (`'S'`→`'C'`); Invoice `UPDATE_INVUnsend` **hard-deletes** the
  `INV_INV_MST` row and nulls the detail links (not a soft "C" status, despite the screen label "Unsend").

### 4.5 Hot-call ASN (`HotCallEntry`)
Manual entry of up to 12 assy-part/qty pairs against a manifest + line + date. Flow (`:221-295`): begin
transaction → `INSERT_ASNInfo` with `@StartSeq=-1`, `@EndSeq=-1`, `@EIN = SiteEIN+1` (start/end seq `-1`
is the **hot-call sentinel** — `ASNSelect.LoadASNDates:117` and the 856 filename branch
`ASNInvoice.pas:817-825` test `StartSeq='-1'` to emit `8HC…` instead of `856…`) → per filled row
`INSERT_ASNDetail` with `@HotCall=1` (forces insert, bypasses the manifest dedup) → `AD_UpdateEIN` →
commit. Manifest must be 8 chars (`:157`). Hot-call manifests do **not** start with `'7'`, so the 810
emits `M390` for them (4.3).

### 4.6 Manifest number generation
`ASNInvoice.AssemblyPartNumber_ComboChange` (`:542`) builds the manifest as
`'7' + copy(ProdDate,4,1) + copy(ProdDate,6,2) + copy(ProdDate,9,2) + <Manifest ID>` — i.e.
`7` + last-digit-of-year + MM + DD + 2-char assy manifest id (8 chars). Leading `'7'` marks a broadcast
manifest (→ `M391` in the 810). Off-by-one sensitive to the `ProductionDate` display format `yyyy/mm/dd`.

### 4.7 Timestamp format (16-char `yyyymmddHHMMSSff`) — count is correct here
Every INSERT/UPDATE proc in this module builds `VC_*` stamps as
`CONVERT(char(8),@Now,112) + SUBSTRING(114,1,2)+SUBSTRING(114,4,2)+SUBSTRING(114,7,2)+SUBSTRING(114,10,2)`
= 8 (date) + 2+2+2+2 = **16 chars** (`hh`,`mm`,`ss`,`ff`). Verified in `INSERT_ASNInfo:2842`,
`INSERT_ASNDetail:2780`, `INSERT_INVInfo:3141`, `UPDATE_INVItems:8482`, `UPDATE_ASNStatus:8199`. These are
**correct** (unlike the stocktaking D8-Bug-2 NULL-stamp). No new timestamp bug found in EDI.

---

## 5. UI / UX notes
- `ASNInvoice` is a dual-mode browser (ASN vs Invoice via `ASNorInvoice_ComboBox`) with a status filter and
  a `@Range` date window. The grid show/hide logic is heavily duplicated across the 5 status cases — collapse
  to one parameterized query in the rebuild.
- `SpeedButton1Click` search (`:907`) infers entity by first char: `'2'` ⇒ production-date search,
  `'7'` ⇒ manifest (then derives the date as `<3-yr-digits>+copy(text,2,1)+'/'+MM+'/'+DD`). Keep the UX
  (search by manifest or date) but back it with real queries.
- `HotCallEntry` walks `GroupBox.Controls[i]`/`[i±1]` pairing combos to edits by control order — fragile
  positional coupling; rebuild as a proper line grid.

## 6. Target design (Ignition)

> **This is the single strongest candidate in the whole system for a Gateway Python service.** X12
> generation is string/segment assembly with implied-decimal money math and TEMA-specific quirks — exactly
> what does not belong in Perspective bindings or Named Queries.

- **Gateway service (`edi_outbound.py` Project Library / scripting):**
  - `build_856(asn_ein)` and `build_810(inv_ein)` that read the line data (via Named Queries wrapping
    window-aware replacements for `REPORT_EDI856`/`REPORT_EDI810`) and emit the X12 stream. Port the segment
    maps in §4.2/§4.3 **verbatim** (separators, element order, the `M390/M391` rule, the `7…`-broadcast
    manifest rule), but replace the hand-rolled TDS money formatting with explicit integer-cents.
  - **Fix the 810 price window (D6):** the line-price Named Query selects the `INV_MANIFEST_COST_MST` row
    whose window contains the ASN production date; reuse the *same* window predicate the 856 uses (made
    inclusive). Make 810 and 856 share one pricing function.
  - File write goes to the configured EDI-out path; transmission stays external (as today).
- **Named Queries** mirror the procs (per the team's NQ-CRUD practice): `asn/list`, `asn/items`,
  `asn/insert_detail` (re-scope the dedup to `(asn_id, manifest)` — see §8), `asn/delete_item`
  (scope to `IN_ASN_DETAIL_ID`, not global manifest), `asn/delete_list` (RESTRICT or cascade detail — D3),
  `inv/create` (`INSERT_INVInfo`+`UPDATE_INVItems`), `inv/unsend`, `ein/ack` (`UPDATE_EINStatus`).
- **Multi-site (D1):** the `AD_GetSite`/`AD_GetSiteTMMDUNS` ALC-DB lookups (EIN, DUNS, separators, dock
  code, EDI mode, supplier code) become per-site columns on the `sites` table; `build_856`/`build_810` take
  a `site` and read its EDI identity. `AD_UpdateEIN` (the EIN counter increment) becomes a per-site sequence.
- **Reports:** `REPORT_EDI856` (window-aware), `REPORT_EDI810`/`REPORT_EDI810Recreate` (window-aware fix),
  `SELECT_INVOICEItems` (window-aware fix) → Named Queries feeding the build service and the browser screen.
- **`InvoiceBreakdown`** (fixed-width supplier-invoice import → `INV_INVOICE_INF`) is a separate small
  importer; see its byte map in §7 of `edi-upload.md`-adjacent notes below — keep it as a tiny parse service.

## 7. Migration plan for this module
- [ ] Stage 1 — wrap `REPORT_EDI856`/`REPORT_EDI810`/`SELECT_INVOICEItems`/list procs as read-only Named
      Queries; render the ASN/Invoice browser; diff generated X12 vs legacy files byte-for-byte.
- [ ] Stage 2 — enable create/send/unsend/ACK writes through Named Queries + the build service; keep the
      legacy EIN counter source until the per-site sequence is cut over.
- [ ] Stage 3 — reimplement window-aware pricing (D6), re-scoped dedup/delete keys, and the per-site EDI
      identity; retire the ALC-DB `AD_*` procs.

## 8. Open questions (candidate D# decisions)
1. **(→ D6 detail) Window boundary inclusivity.** The 856 uses strict `<`/`>`, excluding a production date
   that equals a window boundary. Should the window be **inclusive** (`>=`/`<=`)? (Recommend inclusive; a
   boundary-day ASN currently drops silently from the 856.) Also confirm 810 must use the **same** window
   as 856 so the two documents price identically.
2. **ASN-detail dedup scope.** `INSERT_ASNDetail` dedups (and `DELETE_ASNItem` deletes) by
   `VC_MANIFEST_NUMBER` **globally**. Is a manifest number guaranteed unique across all ASNs/sites, or
   should dedup/delete be scoped to `(IN_ASN_ID, manifest)` (and per-site under D1)? (Recommend scope it.)
3. **ASN delete cascade.** `DELETE_ASNList` orphans detail rows. Under **D3 (RESTRICT)**: block deleting an
   ASN that still has detail lines / is invoiced, or cascade-delete its details? Confirm.
4. **`UPDATE_ASNStatus('S')` flips ALL `'C'` ASNs.** Send currently marks every created ASN as sent (no id
   filter). Is "send" intentionally a batch ("send everything staged") or should it be per-ASN? (Likely a
   single-operator single-batch assumption that breaks multi-user / multi-site.)
5. **Unsend semantics.** `UPDATE_INVUnsend` hard-deletes the invoice header (despite the "Unsend" label).
   Should unsend be a status revert (recoverable) instead, per D3's no-hard-delete direction?
6. **`AD_GetSite`/`AD_GetSiteTMMDUNS`/`AD_UpdateEIN`** live in the ALC DB. Confirm their bodies and that the
   EIN counter is the authoritative outbound control number to replicate as a per-site sequence.

## 9. Test cases / parity checks
- Same `INV_ASN_MST`/`INV_ASN_DETAIL_MST` + a multi-window `INV_MANIFEST_COST_MST` → generate 810: legacy
  picks an arbitrary/duplicated price; rebuild picks the production-date window price. Diff line price + total.
- Production date exactly on a window boundary → confirm 856 drops the line today; rebuild keeps it (inclusive).
- Broadcast manifest (`7…`) vs hot-call manifest → IT1 emits `M391` vs `M390`.
- Byte-diff a generated 856/810 against a known-good legacy file (segment order, separators, SE/CTT counts).
