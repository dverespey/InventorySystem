# Module Analysis: Daily Build Total (`DailyBuildTotal` → ALC pull / ASN / Invoice export)

**Area:** Shipping  **Status:** ✅ spec complete  **Analyst:** Claude / 2026-06-15

> **A single timer-driven batch unit with three modes** (`fmDaily`, `fmASN`, `fmINVOICE`) selected by
> the caller before `Execute`. It is the **automated** sibling of the interactive `Shipping` form:
> `fmDaily` reads an Excel build sheet, decrements inventory the same way `Shipping` does (via the
> stock-OUT triggers), records build history + auto-scrap; `fmASN` and `fmINVOICE` consume that build
> history to emit Toyota-format **ASN** and **invoice** CSV files and mark rows charged/invoiced. The
> work happens in `RunTimerTimer` (`DailyBuildTotal.pas:90`), a 526-line `case fFormMode of`.
>
> **Where it moves stock:** only `fmDaily`, via `InsertExcelShippingInfo → DoPartNumberInventory →
> INSERT_ShippingPartInfo`, which fires **`InsertPartShipping`** (`−IN_QTY` on `INV_PARTS_STOCK_MST`,
> keyed on `VC_PART_NUMBER` string — see [`shipping.md`](shipping.md) §2 for the trigger body), plus
> **auto-scrap** as a negative `INSERT_Stocktaking` delta (keyed on `IN_PART_ID`). `fmASN`/`fmINVOICE`
> are **read + status-flag only** — they do not touch on-hand.

## 1. Legacy surface
- **Form:** `DailyBuildTotal.pas` (618 lines / ~27 KB) + `DailyBuildTotal.dfm` (a near-empty form: a
  `History` log panel, a `DoneButton`, a `RunTimer`). `TDailyBuildtotalForm`. **Live** in
  `InventorySystem.dpr:44`.
- **Entry points (all in `MainMenu.pas`):**
  - `fmDaily` — the ALC daily pull (caller sets `FileName` + `Line`, picks the Excel build sheet).
    Reached from the daily-shipping menu items.
  - `fmASN` — `MainMenu.pas:2570-2575` (`DailyBuildTotalForm.FormMode:=fmASN`; `FromDate`/`ToDate` from
    a date dialog) → emits ASN CSVs.
  - `fmINVOICE` — `MainMenu.pas:2674-2678` (`fmINVOICE`) → emits invoice CSVs.
- **Purpose:** Automate the daily build→inventory→billing pipeline. `fmDaily` is the production
  consume + history capture; `fmASN`/`fmINVOICE` are downstream document generation off
  `INV_ASSY_BUILD_HIST` / `INV_ASSY_PO_CHARGED`.

## 2. Data touched
| Table | Read | Write | Notes |
|-------|:----:|:-----:|-------|
| `INV_SHIPPING_INF` | ✓ | ✓ | `fmDaily`: dup-check via `GetShippingInfo` (`SELECT_ShipLastSeq`); header written by `InsertExcelShippingEndInfo` (`INSERT_ShippingInfo`). |
| `INV_PART_SHIPPING_INF` |  | ✓* | `fmDaily`: `DoPartNumberInventory` → `INSERT_ShippingPartInfo` → fires `InsertPartShipping` (**stock-OUT**). |
| `INV_PARTS_STOCK_MST` |  | ✓* | *Indirect via `InsertPartShipping` (subtract) **and** `INSERT_Stocktaking` (auto-scrap negative delta). |
| `INV_ASSY_RATIO_MST` | ✓ |  | `SELECT_AssyRatioInfoAssy` — tire/wheel part numbers + ratios for the explosion. |
| `INV_ASSY_BUILD_HIST` | ✓ | ✓ | `fmDaily`: `INSERT_AssyBuildHist` per assy. `fmASN`/`fmINVOICE`: read via `SELECT_AssyBuildHist`. DDL schema:1270 (`VC_ASSY_PART_NUMBER_CODE`, `VC_PRODUCTION_DATE`, `IN_QTY`, 16-char `VC_ADD` NULL). **No PK/IDENTITY** — append-only log. |
| `INV_ASSY_MONTHLY_PO` | ✓ | ✓* | `INSERT_AssyPOCharged` also `UPDATE … SET IN_PO_CHARGED += @qty`; `SELECT_AssyBuildHist` joins it for price/availability. |
| `INV_ASSY_PO_CHARGED` | ✓ | ✓ | `fmASN`: `INSERT_AssyPOCharged` (charge a PO). `fmINVOICE`: `SELECT_AssyBuildHist` reads it, `UpdateINVDone` sets `BI_INVOICED=1`. DDL schema:1295 (`IN_CHARGE_ID IDENTITY`, `BI_INVOICED bit`, 16-char `VC_ADD`). |
| `INV_STOCKTAKING_INF` |  | ✓* | `fmDaily` auto-scrap → `INSERT_StocktakingInfo` → fires `INSERT_Stocktaking` (negative on-hand delta, keyed `IN_PART_ID`). |

**Triggers that fire (bodies in [`shipping.md`](shipping.md) §2 and `../inventory-stock`):**
- `InsertPartShipping` (schema:10152) — `−IN_QTY` on `INV_PARTS_STOCK_MST`, string-keyed. **Stock-OUT.**
- `INSERT_Stocktaking` (schema:10416) — `+IN_QTY` on `INV_PARTS_STOCK_MST`, `IN_PART_ID`-keyed; auto-scrap
  passes a **negative** `@QTY` (`0-(new−old)`, `DataModule.pas:4492`) so it nets a decrement (D5 delta model).
- `INSERT_AssyPOCharged`'s in-proc `UPDATE INV_ASSY_MONTHLY_PO SET IN_PO_CHARGED += @qty` (schema:2954) —
  not a trigger but a side-effect to capture.

## 3. Stored procedures used
| Proc | Op | Business rule (from body) |
|------|----|---------------------------|
| `SELECT_AssyRatioInfoAssy;1 (@AssyCode)` | SELECT | schema:5808. `SELECT * FROM INV_ASSY_RATIO_MST WHERE VC_ASSY_PART_NUMBER_CODE=@AssyCode`. Source of the 4 tire/wheel part numbers + ratios in `InsertExcelShippingInfo`. |
| `INSERT_ShippingPartInfo;1` | INSERT/UPD | schema:3735. Idempotent upsert (`IN_QTY += @QTY` if exists); fires the stock-OUT trigger. (Same proc `Shipping` uses.) |
| `INSERT_ShippingInfo;1` | INSERT | schema:3696. `InsertExcelShippingEndInfo` (`DataModule.pas:4231`) writes the daily header. ⚠️ **Calls the 6-param form vs the schema's 9-param OUTPUT proc — same M3 signature mismatch flagged in [`shipping.md`](shipping.md) §4.** |
| `INSERT_AssyBuildHist;1 (@AssyCode,@ProdDate,@Qty)` | INSERT | schema:2865. Append a build-history row with 16-char `VC_ADD`. No dedup. |
| `INSERT_StocktakingInfo;1` | INSERT | schema:3821. ⚠️ **Auto-scrap calls it with 5 params (`@SupCode,@PartCode,@QTY,@Reason,@AutoScrap`) vs the schema's 3 (`@PartNumber,@QTY,@Reason`) — M2 signature mismatch (see [`shipping.md`](shipping.md) §4).** The schema proc resolves `IN_PART_ID` internally and inserts a stocktaking row → fires `INSERT_Stocktaking`. |
| `SELECT_PartsStockInfo;1` | SELECT | schema:7269. ⚠️ Declares **one** param `@PartNum varchar(12)`, but `InsertAutoScrap` (`DataModule.pas:4468`) calls it with **three** (`@InvMgmtReport,@SupCode,@PartNum`) **and reads a field `'Last Scrap Count'` that this proc's SELECT column list does NOT return** (it returns `'Supplier Code'`, `'Parts Code'`, …, `IN_PART_ID 'RecordID'`, but no scrap-count column). **NEW finding: code↔schema drift — auto-scrap's count comparison reads a non-existent field.** Delphi ADO `FieldByName('Last Scrap Count')` on a column absent from the result set **raises `EDatabaseError` ("Field 'Last Scrap Count' not found")** — it does NOT return 0/null. So `InsertAutoScrap` **throws before the comparison runs**; against the checked-in schema auto-scrap is hard-broken (same failure mode as M2), not silently over-scrapping. (Like M1–M3, the live proc may differ from this snapshot — verify.) Confidence high (both bodies read). |
| `SELECT_AssyBuildHist;1 (@BeginDate,@EndDate,@ASN,@INVOICE)` | SELECT | schema:5616. `@ASN<>0` → build-hist ⋈ monthly-PO **on assy code AND `production_date BETWEEN po_month_start AND po_month_end`**, only where `IN_PO_QTY−IN_PO_CHARGED > 0` (PO not yet exhausted). `@INVOICE<>0` → charged ⋈ monthly-PO on PO+assy, `BI_INVOICED=0`, with `IN_QTY_CHARGED*MO_ASSEMBLY_COST totalcost`. **The ASN branch IS window-aware (uses the PO month window)** — contrast D6, where the *invoice/810* read path is window-blind; this build-hist read joins on the window. |
| `INSERT_AssyPOCharged;1 (@AssyCode,@ProdDate,@PickUp,@Qty,@PONumber)` | INSERT+UPD | schema:2930. Inserts a charged row (16-char `VC_ADD`) **and** `UPDATE INV_ASSY_MONTHLY_PO SET IN_PO_CHARGED += @qty WHERE assy AND PO`. Drives PO consumption. |
| `UPDATE_AssyBuildHistINV;1 (@ChargedID)` | UPDATE | schema:8245. `UPDATE INV_ASSY_PO_CHARGED SET BI_INVOICED=1 WHERE IN_CHARGE_ID=@ChargedID`. Marks a charge invoiced. Called by `UpdateINVDone`. |

### P12 retry audit (cited, not re-reported — see [`../cross-cutting/datamodule-retry-target-bugs.md`](../cross-cutting/datamodule-retry-target-bugs.md))
- **`GetBuildHist` (4410) → retries into `GetStocktakingInfo` (4450)** — 🟡 **LOW** P12 (already in the
  register's LOW list). Read-only.
- **`UpdateINVDone` (4317) → retries into `UpdateAssyRatioInfo` (4351)** — 🟠 **MODERATE** P12 (already
  in the register). On a transient invoice-status update failure, the retry calls `UpdateAssyRatioInfo`
  (keyed on shared `fBroadCode`/`fRecordID`) → can clobber an **assembly-ratio** row instead of
  re-flagging the charge. Confirmed existing entry; **cite only**.
- `InsertExcelShippingEndInfo` (4231), `InsertBuildHist` (4531), `InsertPOCharged` (4366),
  `InsertAutoScrap` (4460), `InsertExcelShippingInfo` (4571) — **no recursive retry harness** (they
  log/`raise`/`result:=FALSE`). Not P12.

## 4. Business rules & edge cases
- **`fmDaily` dup-guard.** Before processing it runs `GetShippingInfo` and compares the sheet's
  production date to the stored one (`DailyBuildTotal.pas:131`); if equal → "Already processed", skip.
  Same app-convention lock as `Shipping` (no DB constraint).
- **`fmDaily` is one transaction.** `BeginTrans` (`:140`) wraps the whole assy loop (inventory pull +
  build-hist) and scrap loop and `InsertExcelShippingEndInfo`; any failure rolls back the lot
  (`:170-180`, `:264`). Auto-scrap failures log but do **not** abort (negative `z` just logs, `:249`).
- **Excel coordinate coupling.** `fmDaily` reads fixed cells: production date at `[3,2]`, build rows
  start at row 8 col 1/8, **scrap block at rows 31..** col 1/10 (`:222-256`). The scrap part number is
  reconstructed as `copy(cell,1,5)+copy(cell,7,5)+'00'` (`:240`) — a 12-char part number from a sliced
  assy code. **Brittle, layout-locked** — the rebuild should ingest a defined schema, not cell offsets.
- **Auto-scrap is a delta, not absolute (D5).** `InsertAutoScrap` compares the sheet's scrap count to a
  prior "Last Scrap Count" and posts `0-(new-old)` as a **negative** stocktaking qty (`DataModule.pas:4492`)
  → on-hand drops by the *increase* in cumulative scrap. ⚠️ Tied to the `SELECT_PartsStockInfo`
  drift above — the prior count field isn't returned, so (against the checked-in schema) the
  `FieldByName` call **raises and aborts the post** rather than comparing against 0. Whether auto-scrap
  works at all depends on the live proc matching the snapshot — see §8.1.
- **`fmASN` charges POs; `fmINVOICE` invoices charges.** `fmASN` walks `SELECT_AssyBuildHist`(ASN),
  emits one CSV per pickup-date group, and calls `InsertPOCharged` per line (consuming PO qty).
  `fmINVOICE` walks `SELECT_AssyBuildHist`(INVOICE), emits one CSV per pickup group with
  `IN_QTY_CHARGED*MO_ASSEMBLY_COST`, and calls `UpdateINVDone` (`BI_INVOICED=1`) per charge. Both wrap
  the loop in a transaction; CSVs go to `fiReportsOutputDir`.
- **Invoice price = `MO_ASSEMBLY_COST` from the joined monthly-PO row** (`:534`). The `SELECT_AssyBuildHist`
  invoice branch joins charged↔monthly-PO on **PO number + assy code only** (no date window,
  schema:5638) — consistent with **D6**'s finding that the billing read path is window-blind and
  buggy; the fix (window-aware pricing) is owned by the Invoice/EDI module per D6.
- **Timestamps (P2).** `INSERT_AssyBuildHist`/`INSERT_AssyPOCharged`/`INSERT_ShippingInfo` all use the
  **16-char `yyyymmddHHMMSSff`** recipe (`CONVERT(...,112)` 8 + four `,114` 2-char slices = 16;
  verified schema:2874/2945/3718). CSV file names use `formatdatetime('yyyymmddhhmmss0'/'yyyymmddhhmmss',now)`
  (Delphi-side, 14–15 chars) — a different, non-DB stamp.

## 5. UI / UX notes
- Essentially headless: a `History` scroll log + a `Done` button that appears when the run finishes or
  errors. Driven entirely by `RunTimer` firing once on `FormShow`. Excel is automated via OLE
  (`createOleObject('Excel.Application')`) — a **hard Windows/Excel dependency**.
- **Modernize:** replace Excel-OLE ingest with a defined file (CSV/JSON) parser + a gateway batch job;
  replace the three CSV exporters with server-side report generation; show a real run log + per-line
  outcomes; make the daily pull idempotent against a DB constraint, not a string compare.

## 6. Target design (Ignition)
- **Gateway scripts (not Perspective forms — these are batch jobs):**
  - `dailyBuildPull(site_id, line, file)` — parse the build file; per assy, explode via
    `SelectAssyRatioInfoAssy` and post **`−round(built×ratio/100)`** per part to the shared
    **StockLedger** service (same negative-delta path as `Shipping`); append build-history; post
    auto-scrap as a negative ledger delta; write the daily header — all in one transaction. **Fix M2/M3
    and the `SELECT_PartsStockInfo` drift** by defining correct canonical Named Queries.
  - `generateASN(site_id, fromDate, toDate)` / `generateInvoice(...)` — read via window-correct Named
    Queries, emit the Toyota CSVs to the gateway file area, flag charged/invoiced. **Apply D6**: make
    the invoice price lookup **window-aware** (PO month window contains the production date).
- **Named Queries:** `SelectAssyBuildHist`, `InsertAssyBuildHist`, `InsertAssyPOCharged`,
  `UpdateAssyBuildHistInvoiced`, `SelectPartsStockInfo` (fixed to return the scrap-count field the
  caller needs, or recompute it).
- **Reports:** ASN CSV (`H/O/I` qualifier rows), invoice CSV (`H/I/T` rows) — port the exact column
  layout from `DailyBuildTotal.pas` `fmASN`/`fmINVOICE` (load-bearing field positions).

## 7. Migration plan for this module
- [ ] **Stage 1 — read-only parity:** reproduce `SELECT_AssyBuildHist` (ASN + invoice) outputs;
      validate CSV byte-parity against legacy files for a known date range.
- [ ] **Stage 2 — writes via wrapped procs:** run `fmDaily` through wrapped `INSERT_ShippingPartInfo` /
      `INSERT_AssyBuildHist` / `INSERT_AssyPOCharged` with live triggers intact; **fix M2/M3 and the
      `SELECT_PartsStockInfo` param/field drift** in the wrapper; keep the single-transaction boundary.
- [ ] **Stage 3 — reimplement (Postgres-ready):** replace Excel-OLE with a defined-format ingest; move
      the stock-OUT + auto-scrap into the StockLedger gateway service (additive deltas, `IN_PART_ID`
      keyed); apply D6 window-aware invoice pricing; add `site_id` (D1) to every table/query; give
      `INV_ASSY_BUILD_HIST` a real PK and a unique daily key.

## 8. Open questions for the domain expert (candidate decisions)
1. **(candidate D#) — HARD GATE: live-proc vs checked-in-snapshot diff.** `InsertAutoScrap` reads a
   `'Last Scrap Count'` field the proc doesn't return and passes 3 params to a 1-param proc — against the
   checked-in schema this **raises and auto-scrap can't run at all** (not silent over-scrap). Same shape as
   the M1/M2/M3 signature mismatches (shipping.md). Since this path presumably runs daily in production,
   the live SQL Server almost certainly has richer/divergent procs. **Before the rebuild "fixes" any of
   these signatures, dump the LIVE `SELECT_PartsStockInfo`, `INSERT_ShippingInfo`, `INSERT_ShippingDetail`,
   `INSERT_StockTakingInfo` and diff against `DB Schema/Create Inventory.sql`** — do not reimplement a
   signature production depends on. Is auto-scrap actually working today, and against which proc shape?
2. **Excel layout contract.** Are the fixed cell offsets (date `[3,2]`, builds from row 8, scrap from
   row 31, scrap-part slice `1..5 + 7..11 + '00'`) a stable contract, or do they vary by line/plant?
   The rebuild needs a defined ingest schema.
3. **Invoice pricing window (cross-ref D6).** Confirm the rebuild's invoice price must use the
   PO-month-window-correct `MO_ASSEMBLY_COST` (D6), since the legacy invoice join is window-blind.
4. **Auto-scrap abort policy.** Today a scrap failure logs but does not roll back the daily pull
   (unlike a build/inventory failure). Intended, or should scrap failure abort the day?
5. ✅ **RESOLVED (D5):** auto-scrap qty is a **signed delta** (`0-(new-old)`) applied via
   `INSERT_Stocktaking`; the negative-delta behavior is intended.
6. ✅ **RESOLVED (D1):** per-site — all five tables gain `site_id`; ASN/invoice ranges scope to site.

## 9. Test cases / parity checks
- **`fmDaily` on a fresh date** → per assy: `INV_PARTS_STOCK_MST.IN_QTY −= round(built×ratio/100)` per
  exploded part; one `INV_ASSY_BUILD_HIST` row per assy; one `INV_SHIPPING_INF` header; scrap rows →
  negative `INSERT_Stocktaking` deltas. Re-running the same date → "Already processed", no change.
- **`fmASN`** for a range → one CSV per pickup-date group with `H/O/I` rows; `INSERT_AssyPOCharged`
  consumes PO qty (`IN_PO_CHARGED +=`); only POs with remaining qty (`IN_PO_QTY−IN_PO_CHARGED>0`) appear.
- **`fmINVOICE`** for a range → one CSV per pickup group with `H/I/T` rows; `totalcost = Σ
  IN_QTY_CHARGED×MO_ASSEMBLY_COST`; each charge flips `BI_INVOICED=1`; already-invoiced rows excluded.
- **Auto-scrap delta** → entering cumulative scrap `new` with prior `old` lowers on-hand by `new−old`
  (assert against the `SELECT_PartsStockInfo` field issue — record legacy vs rebuild divergence).
- **Timestamp** = 16-char `yyyymmddHHMMSSff` on every DB `VC_ADD`; CSV filename stamp is the separate
  Delphi 14–15 char form.
- **P12 parity:** force transient failure on `GetBuildHist` / `UpdateINVDone` → assert the rebuild does
  NOT retry into `GetStocktakingInfo` / `UpdateAssyRatioInfo` (legacy P12).
