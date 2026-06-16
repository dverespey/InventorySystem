# Module Analysis: Receiving / Production Reject (`RecReject` → `INV_REJECT_INF`)

**Area:** Receiving  **Status:** ✅ spec complete  **Analyst:** Claude / 2026-06-15

> **A defect/scrap log that subtracts from on-hand.** `RecReject` records a rejected quantity for a
> part (by "division": Receiving / Assembler-production / Plant-production) and a free-text reason.
> The same form is reused for both **Receiving Reject** and **Production Reject** — `MainMenu` sets
> `Data_Module.Division` by the button's `Tag` before opening it. The stock effect is **not** in the
> form or the procs: **three triggers on `INV_REJECT_INF` move `INV_PARTS_STOCK_MST.IN_QTY`**,
> keyed on the **int `IN_PART_ID`** (unlike RecConfStat/shipping, which key on the part-number
> string — this is the keying inconsistency the rebuild's single stock-ledger must reconcile).
> A reject **subtracts** on insert, **adds back** on delete, and re-balances by delta on update.

## 1. Legacy surface
- **Form:** `RecReject.pas` (394 lines / ~13 KB) + `RecReject.dfm` (278 lines). `TRecRej_Form`,
  caption set at runtime to "Receiving Reject" or "Production Reject" (`RecReject.pas:80-94`).
  Author: Aaron Huge, 2002-10-25. Registered live in `InventorySystem.dpr` **line 10**
  (`RecReject in 'RecReject.pas' {RecRej_Form}`).
- **Entry point:** **`MainMenu.pas:292` `RecProdRej_ButtonClick`** — `Case TControl(Sender).Tag of
  1: Division:='1'; 2: Division:='2' …` then `Hide; RecRej_Form := TRecRej_Form.Create(self);
  RecRej_Form.Execute; RecRej_Form.Free; Show;`. **Tag 1 = Receiving Reject button**, **Tag 2 =
  the (assembler) Production Reject button** (`RecProdRej`/`ProdRej`). `Execute` (`RecReject.pas:80`)
  reads `Data_Module.Division` to set the caption + the division radio default.
- **Purpose (one paragraph):** Operators log a rejected/scrapped quantity against a part: pick
  supplier → part (cascading combo), choose the **division** (Receiving / *Assembler* Production /
  *Plant* Production), enter quantity + reason + date, and Insert. Existing rejects list in a grid
  with Update/Delete. Every reject **removes** its quantity from on-hand; deleting a reject **restores**
  it. Division is purely descriptive (a label/classifier) — it does **not** change the stock math.

## 2. Data touched
| Table | Read | Write | Notes |
|-------|:----:|:-----:|-------|
| `INV_REJECT_INF` | ✓ | ✓ | **This module owns it** — INSERT/UPDATE/DELETE via the four procs |
| `INV_PARTS_STOCK_MST` | ✓* | ✓* | *Read: `INSERT_RecProdRejInfo` resolves `@PartNumber→IN_PART_ID`; the `SELECT` joins for code/name. *Write: the three reject triggers add/subtract `IN_QTY` and rewrite `VC_LAST_UPDATE` |
| `INV_SUPPLIER_MST` | ✓ | | `SELECT_RecProdRejInfo` **INNER JOINs** supplier (via part's `IN_SUPPLIER_ID`) for the supplier code; also the supplier search combo (`SelectMultiField`) |

### `INV_REJECT_INF` columns (authoritative: `DB Schema/Create Inventory.sql:1679`)
| Column | Type | Meaning / role |
|--------|------|----------------|
| `IN_REJECT_ID` | `int IDENTITY` **NOT NULL** | Surrogate PK. **The `RecordID` (P9)** used by Update/Delete. Aliased `'RecordID'` in the SELECT, captured from grid `Fields[7]` (`RecReject.pas:196`) |
| `VC_DIVISION` | `char(1)` **NOT NULL** | `'1'`=Receiving, `'2'`=Assembler, `'3'`=Plant. Display-only classifier |
| `IN_PART_ID` | `int NOT NULL` | **FK to `INV_PARTS_STOCK_MST` (by convention) — the trigger join key.** Resolved from the part number inside the INSERT proc |
| `IN_QTY` | `int NOT NULL` | Rejected quantity. **Subtracted from on-hand** by the triggers |
| `VC_REASON` | `varchar(300) NULL` | Free-text reason (the `Reason_Memo`) |
| `VC_LAST_UPDATE` | `varchar(16) NULL` | **16-char `yyyymmddHHMMSSff`** (P2), set by INSERT/UPDATE proc + trigger touches |
| `VC_ADD` | `varchar(16)` **NOT NULL** | **16-char `yyyymmddHHMMSSff`** (P2), set on INSERT only. ⚠️ Note `INSERT_RecProdRejInfo` builds it with `CONVERT(varchar, getdate(), 112)` (no width) + four `,114` 2-char slices — date(8)+HH+MM+SS+ff = **16 chars** (counting check: the `,114` slices are 2 chars each ×4 = 8, +8 date = 16). This is the *correct* full recipe (unlike RecConfStat's 8-char `VC_ADD`) |

**Constraints / indexes:** `IN_REJECT_ID IDENTITY` is the only PK-style column; **no declared
PRIMARY KEY / UNIQUE / FOREIGN KEY** on `INV_REJECT_INF` (verify — none in the table DDL). The
`IN_PART_ID` link is by convention only (RESTRICT-on-delete to be added per D3).

**Triggers on `INV_REJECT_INF` (3 — read live bodies; these ARE the stock effect):**
All key on the **int `IN_PART_ID`** (contrast: RecConfStat/shipping key on the `VC_PART_NUMBER`
string — the inconsistency the rebuild's single ledger must unify).
- **`INSERT_RejectParts`** (FOR INSERT, schema:10239): `UPDATE INV_PARTS_STOCK_MST SET IN_QTY =
  PS.IN_QTY − i.IN_QTY, VC_LAST_UPDATE = i.VC_LAST_UPDATE WHERE PS.IN_PART_ID = i.IN_PART_ID`.
  **Invariant: logging a reject removes its qty from on-hand immediately.** No add-point gate, no
  division gate — *every* reject subtracts.
- **`UPDATE_RejectParts`** (FOR UPDATE, schema:10272): two unconditional legs — `+= d.IN_QTY` (add
  back the pre-update qty) then `−= i.IN_QTY` (subtract the new qty), both `WHERE IN_PART_ID = …`.
  Since `IN_PART_ID` is **not changeable through the form** (see §4), this nets to
  **`IN_QTY += (d.IN_QTY − i.IN_QTY)`** for the one part. ⚠️ The "add-back" leg uses a freshly
  computed `@Deleted` timestamp for `VC_LAST_UPDATE`; the "subtract" leg uses `i.VC_LAST_UPDATE`.
  **Invariant: editing a reject re-balances on-hand by the qty delta.**
- **`DELETE_RejectParts`** (FOR DELETE, schema:10206): `IN_QTY += d.IN_QTY, VC_LAST_UPDATE =
  @Deleted WHERE IN_PART_ID = d.IN_PART_ID`. **Invariant: deleting a reject restores its qty to
  on-hand.** ⚠️ **No purge-mode bypass** (unlike the RecConfStat delete trigger) — a data-purge that
  deletes reject rows **will add their qty back to on-hand**. Possible purge-time drift; flag.

## 3. Stored procedures used
(Read from `DB Schema/Create Inventory.sql`. All bodies verified.)

| Proc | Op | Business rule (from body) |
|------|----|---------------------------|
| `SELECT_RecProdRejInfo;1` (no params) | SELECT | schema:7635. Lists **all** rejects (no filter — legacy single-site; D1 must scope by `site_id`). **INNER JOIN** part (`r.IN_PART_ID = p.IN_PART_ID`) **and** supplier (`p.IN_SUPPLIER_ID = s.IN_SUPPLIER_ID`). ⚠️ The inner supplier join means **a reject whose part has a NULL/dangling `IN_SUPPLIER_ID` is silently dropped from the list** (P5-adjacent: the reject still affects stock but vanishes from the grid). Returns 8 cols → grid `Fields[0..7]`: `Date` (formatted `yyyy/mm/dd` from `VC_ADD`), `Division` (CASE 1→Receiving/2→Assembler/3→Plant), `Supplier Code`, `Parts Code`, `Parts Name`, `QTY`, `Reason`, `RecordID(=IN_REJECT_ID)`. `ORDER BY supplier, part`. |
| `INSERT_RecProdRejInfo;1` (`@Division,@PartNumber,@QTY,@Reason`) | INSERT | schema:3571. Builds a 16-char `@Add`, **resolves `@PartNumber→@PartID` via `SELECT IN_PART_ID FROM INV_PARTS_STOCK_MST WHERE VC_PART_NUMBER=@PartNumber`** (P3), then inserts `(VC_DIVISION, IN_PART_ID, IN_QTY, VC_REASON, VC_LAST_UPDATE=@Add, VC_ADD=@Add)`. ⚠️ If the part number doesn't match, `@PartID` is **NULL** → insert fails (`IN_PART_ID` NOT NULL). Fires `INSERT_RejectParts` (the subtract). |
| `UPDATE_RecProdRejInfo;1` (`@Division,@SupCode,@PartCode,@QTY,@Reason,@RejectID`) | UPDATE | schema:9191. `UPDATE INV_REJECT_INF SET VC_LAST_UPDATE=@Update, VC_DIVISION=@Division, IN_QTY=@QTY, VC_REASON=@Reason WHERE IN_REJECT_ID=@RejectID`. ⚠️⚠️ **`@SupCode` and `@PartCode` are accepted but NEVER USED** — the proc does **not** update `IN_PART_ID`. **So an operator cannot change a reject's part via Update; only qty/division/reason change.** Editing the part combo and saving silently keeps the old part. (And because `IN_PART_ID` can't change, `UPDATE_RejectParts` correctly re-balances the *original* part.) Fires `UPDATE_RejectParts`. |
| `DELETE_RecProdRejInfo;1` (`@RejectID int`) | DELETE | schema:2458. `DELETE INV_REJECT_INF WHERE IN_REJECT_ID = @RejectID`. Fires `DELETE_RejectParts` (the add-back). |
| `SELECT_DependantPartNumber_Supplier (@SupplierCode)` | SELECT | schema:5956. Cascading part combo for the chosen supplier (`RecReject.pas:380`). Body unverified in detail (combo-population only). |

### Call mechanism (legacy — `DataModule.pas`)
- **`GetRecProdRejInfo`** (5404): opens `Inv_DataSet` on `SELECT_RecProdRejInfo`. P8 retry → re-calls
  **itself** (5436). ✅ not P12.
- **`InsertRecProdRejInfo`** (5446): **no app-side dup-check** (unlike most other inserts) — calls
  `INSERT_RecProdRejInfo` directly with `@Division,@PartCode,@QTY,@Reason`. ⚠️ **Param-name note:**
  the DataModule adds a parameter literally named **`@PartCode`** (`DataModule.pas:5462`) but the proc
  declares **`@PartNumber`** (schema:3573). ADO `TADOStoredProc` binds these positionally at exec time
  (the names are local labels), so position 2 still carries the part number — **but verify** during the
  rebuild that the binding is positional and not by-name (a by-name bind would leave `@PartNumber`
  NULL → insert failure). Flag as a fragility, not a confirmed runtime break. P8 retry → re-calls
  **itself** (5490). ✅ not P12.
- **`UpdateRecProdRejInfo`** (5501): passes `@Division,@SupCode,@PartCode,@QTY,@Reason,@RejectID`
  (`@RejectID := fRecordID`, the shared P9 field). Two of the six params (`@SupCode`,`@PartCode`)
  are dead-ends in the proc (above). P8 retry → re-calls **itself** (5545). ✅ not P12. ⚠️ Note its
  `finally` does **not** `Inv_StoredProc.Close` (only resets `fErrorCount`) — minor leak vs the others.
- **`DeleteRecProdRejInfo`** (5555): `@RejectID := fRecordID`. P8 retry → re-calls **itself** (5587).
  ✅ not P12.
- **P9 `RecordID` hazard:** Update/Delete key off the **shared, generic `Data_Module.RecordID`**
  (line 196 sets it from grid `Fields[7]`; 316/5523/5565 read it). A stale `RecordID` left by another
  screen would mis-target — the standard P9 risk. The form mitigates by re-`HoldDetails(True)` on
  every grid selection, but the field is global.

## 4. Business rules & edge cases
- **Every reject subtracts; delete restores; update re-balances by delta.** No add-point gate, no
  division gate — division is purely a label. This is simpler than RecConfStat (no `'S'`/`'A'` logic).
- **Part is fixed after creation.** `UPDATE_RecProdRejInfo` ignores `@SupCode`/`@PartCode` and never
  touches `IN_PART_ID`, so a reject's part cannot be re-pointed via Update. To move a reject to a
  different part the operator must delete and re-insert. **Rebuild decision needed (§8):** allow part
  change on edit (and re-balance both old and new parts), or keep it immutable.
- **Insert requires a resolvable part number.** `INSERT_RecProdRejInfo` looks up `IN_PART_ID` by
  `VC_PART_NUMBER`; an unmatched/blank part → NULL `IN_PART_ID` → insert fails (NOT NULL). The form's
  only guard is `HoldDetails` setting `fErrMsg='Part Number must not be blank'` **but that message is
  built and never actually blocks the insert** (`RecReject.pas:213-216/304` — `Insert_ButtonClick`
  ignores the returned `fErrMsg`). So a blank part is caught only by the DB NOT-NULL failure.
- **Keying inconsistency vs the rest of receiving.** Reject triggers key on **`IN_PART_ID` (int)**;
  RecConfStat and shipping triggers key on **`VC_PART_NUMBER` (string)**. The rebuild's single
  stock-ledger service must standardize on `IN_PART_ID` and resolve the string at the boundary.
- **No purge bypass on reject delete.** `DELETE_RejectParts` always adds the qty back, even under
  `Purge.PurgeMode = 1` (the RecConfStat delete trigger checks `Purge`). A purge of old reject rows
  would inflate on-hand — confirm whether purge ever deletes rejects (§8).
- **Inner-join supplier hides "supplierless" rejects.** A reject on a part whose supplier was deleted
  (`IN_SUPPLIER_ID` NULLed by `DELETE_SupplierCode`) disappears from the grid while still affecting
  stock — an invisible-but-active row.
- **Timestamps (P2):** both `VC_ADD` and `VC_LAST_UPDATE` use the full **16-char**
  `yyyymmddHHMMSSff` recipe. The grid `Date` column reformats `VC_ADD`'s first 8 chars to
  `yyyy/mm/dd`. (This module gets the 16-char recipe right, unlike RecConfStat's 8-char `VC_ADD`.)
- **Division mapping is duplicated** between the proc (`CASE`/numeric) and the form (radio index ↔
  `'1'/'2'/'3'`, `RecReject.pas:182-211/240-243`). Labels for divisions 2/3 are site names
  (`fiAssemblerName`/`fiPlantName`, `RecReject.pas:372-373`) — per D1 these come from the site row.

## 5. UI / UX notes
- Grid of all rejects + a small detail panel: supplier combo (`TNUMMIColumnComboBox`), cascading
  part combo, a 3-option division `RadioGroup`, qty `TMaskEdit`, reason `TMemo`, date picker.
- **Search is client-side (P7):** `Filter [Supplier Code] = '<exact>' AND [Parts Code] = '<exact>'`
  (exact match, both required — `RecReject.pas:341-346`).
- **Modernize:** server-side search/pagination (P7); make blank-part a real blocking validation;
  decide editable-part-on-update; show "supplierless" rejects instead of hiding them (LEFT JOIN);
  surface the running on-hand impact so an operator sees the deduction.

## 6. Target design (Ignition — Perspective + Named Queries + gateway stock-ledger)
- **Perspective views:**
  - `Receiving/Rejects` — a Table bound to a `SELECT_RecProdRejInfo` Named Query (param `site_id`),
    with server-side supplier/part filters. **Use a LEFT JOIN** so supplierless rejects still show
    (fixing the inner-join hide).
  - `Receiving/RejectEditor` — supplier + cascading part dropdowns (NQ
    `SELECT_DependantPartNumber_Supplier`), division radio, qty, reason. Real blocking validation on
    blank/unresolvable part.
- **Named Queries (one per proc):** `SelectRejects`, `InsertReject`, `UpdateReject`, `DeleteReject`.
  Parallel-run: wrap the procs via `system.db.createSProcCall` so the three triggers keep on-hand
  correct; add `site_id`.
- **Gateway stock-ledger service:** fold the three reject triggers into the **same** `StockLedger`
  service used by RecConfStat/shipping/stocktaking: reject insert → post `−qty`; delete → `+qty`;
  update → delta. Key on `IN_PART_ID`. Each posting writes a ledger row + stamps the part's
  last-update. Decide the **purge** policy explicitly (legacy adds back on delete with no bypass).
- **Fixes baked in:** drop the dead `@SupCode`/`@PartCode` proc params or wire them to re-point the
  part (and re-balance both parts) per the §8 decision; make blank-part blocking; resolve the
  `@PartCode`/`@PartNumber` param-name mismatch.
- **Reports:** none owned here.

## 7. Migration plan for this module
- [ ] **Stage 1 — read-only parity:** `SelectRejects` NQ → Table; 8 columns map to grid `Fields[0..7]`;
      add `site_id`; switch the supplier join to LEFT (note the parity divergence for supplierless rows).
- [ ] **Stage 2 — writes via wrapped procs:** call `INSERT/UPDATE/DELETE_RecProdRejInfo` through
      `system.db.createSProcCall`, **keeping the three reject triggers live**. Add real blank-part
      validation. Decide the purge policy (whether reject deletes re-balance on-hand during purge).
- [ ] **Stage 3 — reimplement (Postgres-ready):** add `site_id` NOT NULL FK; add a real PK
      (`IN_REJECT_ID`) and a **FK `IN_PART_ID → INV_PARTS_STOCK_MST` with RESTRICT** (D3 — block
      deleting a part that still has rejects); move the three triggers into the `StockLedger` gateway
      service keyed on `IN_PART_ID`; normalize the timestamp strings; resolve editable-part-on-update;
      normalize division to an enum.

## 8. Open questions for the user (domain expert)
1. **Editable part on reject Update?** Today `UPDATE_RecProdRejInfo` silently ignores the part — a
   reject's part is immutable after insert (change requires delete+re-insert). Should the rebuild let
   an operator re-point a reject to a different part on edit, re-balancing on-hand for **both** the old
   and new part? (Recommend: allow it, with a two-part ledger adjustment.)
2. **Purge policy for rejects.** `DELETE_RejectParts` has **no** purge bypass (unlike the RecConfStat
   delete), so purging old reject rows would **add their qty back to on-hand**. Does the data-purge
   job ever delete `INV_REJECT_INF` rows? If so, the rebuild must skip the re-balance during purge
   (mirror RecConfStat); if not, this is moot.
3. **Division semantics.** `VC_DIVISION` (Receiving / Assembler / Plant) is currently display-only and
   doesn't affect stock. Is it meant to be purely a classifier, or should some divisions route to a
   different account / not deduct on-hand (e.g. a "found, not scrapped" case)?
4. **Negative / over-reject quantities.** A reject qty larger than on-hand drives `IN_QTY` negative
   (the trigger does raw subtraction, no floor). Is negative on-hand acceptable, or should the rebuild
   reject (pun intended) an over-deduction?
5. ✅ **RESOLVED (D1):** per-site — `INV_REJECT_INF` gains a `site_id` NOT NULL FK; `SELECT_RecProdRejInfo`
   scoped to the current site.
6. ✅ **RESOLVED (D2):** standardize the `IN_PART_ID` surrogate as the sole link; the dead
   `@SupCode`/`@PartCode` string params disappear.
7. ✅ **RESOLVED (D3):** block deleting a part that still has reject rows (add the `IN_PART_ID` FK with
   RESTRICT); use archival, not delete, to retire a referenced part.

## 9. Test cases / parity checks
- **List all** → row count = `SELECT_RecProdRejInfo` (note: rows whose part has a NULL supplier are
  **hidden** by the legacy inner join — assert the rebuild's LEFT-JOIN shows them, documenting the
  divergence).
- **Insert a reject** (part resolvable, qty Q) → `INV_REJECT_INF` row added with 16-char `VC_ADD`,
  `IN_PART_ID` resolved from the part number; `INV_PARTS_STOCK_MST.IN_QTY −= Q` (`INSERT_RejectParts`).
- **Insert with a blank/unresolvable part** → legacy: DB NOT-NULL failure (no row, on-hand unchanged);
  rebuild: blocked by validation before the DB.
- **Update a reject qty** Q1→Q2 → on-hand `+= (Q1−Q2)` (`UPDATE_RejectParts` delta); change the part
  combo and save → legacy: part **unchanged**, on-hand re-balances the original part only; rebuild:
  per §8.1 decision.
- **Delete a reject** (purge off) → `INV_REJECT_INF` row gone; on-hand `+= Q` (`DELETE_RejectParts`).
  ⚠️ With purge on, legacy **still** adds back (no bypass) — assert the rebuild's chosen purge behavior.
- **Over-reject** (Q > on-hand) → on-hand goes negative in legacy; assert the rebuild's floor/guard (§8.4).
- **P9 stale-`RecordID` parity:** open Reject after another screen set `RecordID`, Update without
  re-selecting a grid row → confirm the rebuild scopes the edit to the form's own selected row id.
