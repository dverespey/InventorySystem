# Daily-Workflow Usage Model + Spike-Coverage Gap Map

**Source of truth:** `DB Schema/DailyWorkLog.csv` — one full production day (2026-06-19,
07:52–14:57) of the production-control operator (KTURNER, user `lktur`; one admin session by
GHEATH/`Administrator` 11:15–11:26) driving the LEGACY Delphi `InventorySystem.exe` (Ver 2.9.4.1).
174 log rows = `app, action-category, description, user, timestamp`, emitted by
`Data_Module.LogActLog(category, description)`.

**Purpose:** decode what the operator actually does in a day, map each action to the live Delphi
unit + stored proc + file/EDI artifact, diagnose the ERROR rows, and produce a coverage-gap table
vs the Ignition spike. This is the input the ignition-architect turns into a production plan.

**Headline finding:** the cutover-readiness checkpoint
(`docs/analysis/cutover-readiness-checkpoint.md`) calls the rebuild "feature-complete." Measured
against a real operator-day, that is **true only of the data/ledger/master-data foundation + the
Order worksheet DISPLAY.** The actual revenue-critical daily loop — **ASN entry → EDI 856 out →
997/862/810 import → ASN invoicing → order-file generation to supplier/logistics/archive** — is
**almost entirely NOT BUILT** in the spike. 86 of the 174 log rows (≈49%) are ASN/EDI/EDIIMP/order-
file operations that have no Perspective view today.

---

## 1. Action-category frequency (parsed from all 174 rows)

| Category | Count | Meaning | Emitting unit |
|---|---:|---|---|
| SELECT SFT | 32 | combobox/lookup audit ("SELECTED `<col>` from `<table>`") | `DataModule.pas:5750,5800` |
| ORDERF | 27 | order-FILE generation (text/excel for supplier/logistics/archive) | `OrderFormCreateF.pas` |
| ASN | 23 | ASN sequence check + create ASN entries (detail inserts) | `ASNSelect.pas`, `DataModule.pas` |
| RENBAN BD | 19 | renban-group breakdown (FRS lots → renban numbers) | `RenbanOrder.pas` |
| ORDER | 13 | order worksheet open/create/exit | `Order.pas` |
| EDIIMP | 11 | inbound EDI import (810/856 ack via 997, 862 firm order) | `EDIUpload.pas` |
| ASNINV | 11 | ASN-invoice qty edits (update/delete on ASN detail) | `ASNInvoice.pas` |
| ORDERS | 9 | per-supplier Order Sheet (Excel) within order-file gen | `OrderFormCreateF.pas` |
| UPD EINSta | 5 | EIN status update from 997 ack | `DataModule.pas:6783` |
| ERROR | 5 | runtime failures (see §4) | various |
| EDI | 4 | outbound EDI 856 file creation | `MainMenu.pas`, `ASNInvoice.pas` |
| START / LOGIN | 3 / 3 | program start / login | `Logon.pas` / startup |
| UPD ASNSta | 2 | ASN status C→S after 856 sent | `DataModule.pas:5095` |
| STOP | 2 | program stop | shutdown |
| INS ASN | 2 | single ASN-detail insert (ASNInvoice screen) | `DataModule.pas:5040` |
| HOTCALL | 2 | urgent hot-call ASN add | `HotCallEntry.pas` |
| REPORT | 1 | Daily Shipping Assy Report completed | `MainMenu.pas:3210` |

(Live-unit confirmation: every unit above is listed in `InventorySystem.dpr`. None are dead-code
dupes.)

---

## 2. Chronological daily workflow (the real operator loop)

The day is **not** linear feature-by-feature; it is an **EDI cadence** the operator services in
bursts as files arrive, interleaved with the once-or-twice-a-day order run. Reconstructed phases:

### Phase 1 — Morning ASN build for prior production day (07:52–08:09, lktur)
1. **Login + start** (rows 1–2).
2. **ASN sequence check** (row 3): `ASN Sequence number check, P:06/18/2026 L:COROLLA S:909 E:756 Q:848`
   — operator confirms the GALC-fed broadcast sequence range (Start 0909 → Last 0756, wraps) and
   shipped qty 848 for the COROLLA line, production date 06/18.
3. **Create ASN entries** (rows 4–25): one click fans out into **20 detail inserts** keyed to
   `ASNID(4721)` — one per manifest (`76061857`…`76061805`), plus **2 "(No Ratio)"** rows
   (manifests 76061836, 76061851) — the system splits qty by assembly-ratio and emits a remainder
   row when the ratio doesn't divide evenly. Then an **ASN information** header row (row 24:
   `Production Date 20260618, Line COROLLA, Start 0909, Last 0756`) and `Finish create ASN entries`.
   This is one transaction (`Inv_Connection.BeginTrans` … `CommitTrans` in
   `ASNSelect.pas:372–412`).

### Phase 2 — ASN invoice reconciliation (08:06–08:08, lktur)
4. **ASNINV edits** (rows 26–36): operator opens the ASN-invoice screen and **corrects per-manifest
   qty against actual shipped** — 9 Updates + 1 Delete (e.g. manifest 76061836 842→800, 76061837
   deleted, 76061857 842→80). This is the human reconciliation step between auto-split and reality.

### Phase 3 — EDI 856 (ASN) outbound + status flip (08:09)
5. **Create EDI 856** (rows 37–38): writes `X:\EDIOut\85620260618COROLLA.txt` for production date
   20260618.
6. **UPD ASNSta** (row 39): `Update ASN Status=S` — flips the ASN master from C(reated) to
   S(ent).

### Phase 4 — Daily Shipping Assy Report (08:34–08:48) — FAILS 3×, then succeeds
7. Three **ERROR** rows (40–42): `Failed on Daily Shipping Assy Report, PrintOut method of
   Worksheet class failed` — the Excel print-out fails repeatedly (see §4).
8. A detour into **Order** (rows 43–47, see Phase 5) produces 2 more errors, then
9. **REPORT** (row 48): `Do Daily Shipping Assy Report` finally succeeds (the operator likely
   declined the print prompt, or printer recovered).

### Phase 5 — Order worksheet (08:47, lktur attempt; then 11:15–11:26 admin completes)
10. lktur opens **Start Excel Order Line(COROLLA) Part Type(WHEEL)** (row 45) but it **fails
    immediately** (rows 46–47): `Unable to get month forecast in order` then `Unable to export
    infomation in order` — Excel/forecast failure aborted the worksheet (see §4). lktur abandons it.
11. **GHEATH (Administrator) logs in 11:15** and runs the full order cycle successfully:
    - WHEEL order: creates 4 order lines (rows 62–65: pn/kanban/qty, **renban blank** = assigned
      later), Exit Order.
    - **RENBAN BD** breakdown (rows 67–74): `Total Lots:50, Trailer Pallet Count:50, Trailers:1`;
      new breakdown items per part; deletes temp FRS group orders; inserts records with renban
      `CMWA296`; advances the group counter `Renban=CMWA, Number=297`.
    - Repeats for a second supplier group (rows 75–90: renban `DICAS483`, counter→484).
    - **ORDERF / ORDERS** order-file generation (rows 91–117): per supplier — Order Sheet Excel +
      **3 text files**: supplier (`S:\CMX\CMWA-0572B.ord`), logistics (`S:\TLDL\CMWA-0572B.ord`),
      archive (`S:\CMX\Archive\CMWA-0572B<timestamp>.ord`). Then DICASTAL (07100) the same way.
    - A second order run for **Part Type(TIRE)** (rows 118–131): one group-renban line
      (4265202R8000, renban group `16DL`, qty 1620) → DUNLOP order files (note: **TIRE writes
      supplier + archive but NO logistics file** — rows 124–127 have no `TLDL` line).
12. Admin **STOP** (row 132).

### Phase 6 — Inbound EDI acks throughout the day (lktur)
Serviced in three bursts as Toyota's mailbox delivers:
- **09:50** (rows 49–55): 810 EIN 9068 Accepted (997 `M6I29098`), **862 firm order** `M6J29108`
  processed, 856 EIN 9069 Accepted (997 `M6J29114`).
- **11:26** (rows 133–135): 810 EIN 9070 Accepted (997 `M6J29120`).
- **13:43** (rows 167–169): 856 EIN 9071 Accepted (997 `M6J29122`).
- **14:57** (rows 171–173): 810 EIN 9072 Accepted (997 `M6J29126`).
Each 997 ack triggers a **UPD EINSta** (rows 49,53,133,167,171) that flips the matching ASN/INV
master record to A(ccepted).

### Phase 7 — Afternoon hot-call (13:23–13:24, lktur)
13. lktur logs back in (rows 136–137), the forecast-detail combobox fires **24 SELECT SFT** lookups
    (rows 138–163 — populating the assembly-part dropdown twice), then a **HOTCALL** (rows
    150–151): `Added PartNumber(42600FEL1000) Qty(1) HotCall manifest(52089913)` — an urgent
    single-part add outside the normal ASN cycle.
14. **EDI 856 for the hot call** (rows 164–166): filename `X:\EDIOut\8HC606191COROLLA.txt` (the
    `8HC` prefix = hot-call 856, `StartSeq = -1` branch in `MainMenu.pas:2722`), then
    `UPD ASNSta = S`.
15. **INS ASN** (row 170, also row 56): a single manual ASN-detail insert (`Insert , 1`).
16. **STOP** (row 174).

**Cadence summary:** ASN build + 856 out happens **once per production day per line** (morning).
EDI inbound import is **polled in bursts ~4×/day**. Order worksheet + renban + order-file gen runs
**once or twice** (here: WHEEL then TIRE). Hot-call is **ad hoc** (1×). The Daily Shipping Assy
Report is **once** (after a fight). SELECT SFT is constant background noise (combobox population).

---

## 3. Action → legacy code map (unit + proc + artifact)

| Category | Live unit (`.pas`) | Stored proc(s) | Reads / Writes |
|---|---|---|---|
| **SELECT SFT** | `DataModule.pas:5750,5800` (`LogActLog('SELECT SFT', 'SELECTED '+fFieldName+' from '+fTableName)`) | generic dataset opens (e.g. `LINE`, `INV_PART_TYPE_MST`, `INV_FORECAST_DETAIL_INF`) | read-only combobox population; **pure audit noise**, no business effect |
| **ASN** (seq check) | `ASNSelect.pas:358` | `SELECT_ASNSeq` (`CreateInventory.sql:1517`) | reads broadcast seq range from GALC-fed data |
| **ASN** (create entries) | `ASNSelect.pas:369–419` → `DataModule.InsertASNInfo`; detail rows `DataModule.pas:5211,5263` | `INSERT_ASNInfo` (`:2529`, OUTPUT `@ASNID = SCOPE_IDENTITY`, status `'C'`), `INSERT_ASNDetail` (`:2682`), `AD_UpdateEIN` | writes `INV_ASN_MST` (header) + `INV_ASN_DTL` (one row/manifest; "(No Ratio)" remainder rows) |
| **ASNINV** | `ASNInvoice.pas` | `UPDATE_ASNItem` (`:3374`), `DELETE_ASNItem` (`:2800`), `SELECT_ASNItems` (`:3332`), `SELECT_ASNMissingCost` (`:3504`) | edits `INV_ASN_DTL` qty before invoicing |
| **INS ASN** | `DataModule.pas:5040` (`LogActLog('INS ASN','Insert , 1')`) | `INSERT_ASNDetail` | single manual ASN-detail insert |
| **EDI** (856 out) | `MainMenu.pas:2717,2722,2736` (`ResendMarkedEDIs`); also `ASNInvoice.pas:837`; builder `EDI856Object.pas` | `REPORT_EDI856` (`:3629`) feeds `T856EDI` | writes `X:\EDIOut\856<date><line>.txt` (normal) or `8HC<...>` (hot-call, StartSeq=-1) |
| **UPD ASNSta** | `DataModule.pas:5095` (`UpdateASNStatus`) | `UPDATE_ASNStatus` (`:1679`) — `WHERE VC_ASN_STATUS='C'` → `'S'` | flips `INV_ASN_MST` C→S |
| **EDIIMP** (997/862/856/810/830/824) | `EDIUpload.pas` (`:103,184,210,215,251,300`) | `UPDATE_EINStatus` (`:1711`) per AK1 segment | reads `<EDIIn>\*.txt`; 862 writes `FirmOrder<ts>.xls`; 824 writes report sheet; 830 → `ForecastBreakdown_Form` |
| **UPD EINSta** | `DataModule.pas:6783` (`UpdateEINStatus`) | `UPDATE_EINStatus` (`:1711`): `SH`→`INV_ASN_MST.VC_ASN_STATUS`, else→`INV_INV_MST.VC_INV_STATUS` by EIN | flips ASN/INV master to A/R per 997 ack |
| **ORDER** | `Order.pas:162,755,852,617` | `SELECT_PartsStockInfoOrder` (`:4865`), `SELECT_ForecastPartNumberWeek` (`:2076`), `SELECT_FirstProductionDay`, `INSERT_OpenOrder`, `UPDATE_PartsStockRenban` | builds Excel worksheet; writes `INV_OPEN_ORDER_INF` on commit |
| **RENBAN BD** | `RenbanOrder.pas:433,512,558,712,781` | `SELECT_RenbanGroup` (`:4952`), `UPDATE_RenbanGroupCount` (`:5142`), FRS-renban insert/delete procs | FRS lots → renban numbers; advances group counter |
| **ORDERF / ORDERS** | `OrderFormCreateF.pas` (`:80,251,277,294,307,314,388,524,687`) | reads `Inv_StoredProc` order result (`Create Order Sheet` field) | writes Order Sheet `.xls` + 3 `.ord` text files (supplier / logistics / archive) per supplier |
| **HOTCALL** | `HotCallEntry.pas:282,297` | ASN-detail insert procs (hot-call manifest) | writes `INV_ASN_DTL` for urgent manifest; then 856 `8HC` file |
| **REPORT** | `MainMenu.pas:3210` | `REPORT_DailyShippingAssy` (`:3576`) | Excel report `DailyShippingAssy<ts>.xls`, optional PrintOut |

**File-path hazards (single-site, hard-coded UNC):** `X:\EDIOut\` (outbound EDI), the inbound
`<EDIIn>` dir, and per-supplier roots `S:\CMX\`, `S:\DICASTAL\`, `S:\DUNLOP\`, `S:\TLDL\` (logistics),
each with an `\Archive\` subdir. These come from INI `[DIRECTORIES]` + supplier-master path columns.
The `.ord` text layout is **fixed-width positional** (e.g. row 99:
`0572B6062601 CMWA2964261102Q51000120020260625` = supplier+forecast+renban+part+qty(7,zero-pad)+
FRS-date(8)). Any Ignition replacement must reproduce these byte-exact filenames + layouts because
**downstream supplier/logistics systems parse them.**

---

## 4. ERROR-row diagnoses (production pain — 5 rows)

All five ERROR rows trace to **Excel COM/OLE automation**, the single most fragile dependency in
the app. Each Delphi report/worksheet drives a hidden `Excel.Application` via `CreateOleObject`.

| Rows | Message | Code path | Diagnosis |
|---|---|---|---|
| 40,41,42 | `Failed on Daily Shipping Assy Report, PrintOut method of Worksheet class failed` | `MainMenu.pas:3203` (`mysheet.PrintOut(...)`) inside try at `:3219` | **Not a code bug — an Excel/printer-driver failure.** The proc + data succeeded (the report later completes at 08:48); only `Worksheet.PrintOut` threw, 3× across 13 min. Indicates a flaky/offline default printer or an Excel automation-print quirk. Production-hardening: render server-side (no Excel, no client printer) so reporting can't be blocked by a printer. |
| 46 | `Unable to get month forecast in order, Arguments are of the wrong type, are out of acceptable range, or are in conflict with one another` | `Order.pas:1125` in `ForecastHistory` — writes `mysheet.Cells[line+1,8].value` (`:1120`) | Classic Excel-OLE `DISP_E_*` "wrong type / out of range" error from a Cells assignment (or a NULL/locale-mangled value pushed into a cell). The forecast SELECT procs (`SELECT_FirstProductionDay`, `SELECT_ForecastPartNumberWeek`) likely returned fine; the **Excel cell write failed**, and because `ForecastHistory` does `raise` (`:1127`), it propagates up and aborts the whole worksheet. |
| 47 | `Unable to export infomation in order, <same Excel arg error>` | `Order.pas:583` — the outer `except` in `ExportToExcel`; the re-raised forecast error lands here | **Same root cause, just the outer handler.** One Excel failure (row 46) bubbled to abort the export (row 47). The worksheet was abandoned by lktur and later redone by the admin session — i.e. **the order run is brittle: a single Excel hiccup loses the worksheet.** |

**Pattern (production-critical):** the operational reports + the Order worksheet are **client-side
Excel-automation-bound**, so they fail on printer/Excel/locale conditions the server can't control.
This is the strongest argument for the Ignition rebuild to render reports + worksheets
**server-side** (the D6 report-proc migration already moves the *data* server-side; the
*presentation/print* path is still the unhardened part).

---

## 5. COVERAGE GAP MAP — daily capability vs Ignition spike

Spike inventory confirmed by enumerating built Perspective views (`find . -name view.json`): only
**Home** (landing hub), **Order/OrderSpike** (worksheet DISPLAY), and **7 master-data CRUD** views
(Supplier, Size, Logistics, RenbanGroup, PartsStock, AssemblyDetail, ManifestCost) exist. Plus the
non-UI foundation: stock-ledger + 4 producers + Order commit path + D6 report-proc migration
(`cutover-readiness-checkpoint.md` §1).

| # | Daily capability (from log) | Legacy unit/proc | Spike status | Evidence / note |
|---|---|---|---|---|
| 1 | **ASN sequence check** | `ASNSelect.pas` / `SELECT_ASNSeq` | **NOT BUILT** | no ASN view; spec only (`docs/analysis/edi/asn-invoice.md`) |
| 2 | **ASN entry creation** (split-by-ratio, No-Ratio remainder, header+detail) | `ASNSelect.pas` / `INSERT_ASNInfo`+`INSERT_ASNDetail` | **NOT BUILT** | the morning revenue keystone — zero spike UI/logic |
| 3 | **ASN invoice reconciliation** (qty update/delete) | `ASNInvoice.pas` / `UPDATE_ASNItem`,`DELETE_ASNItem` | **NOT BUILT** | spec only (`edi/asn-invoice.md`) |
| 4 | **EDI 856 (ASN) outbound** incl. hot-call `8HC` | `MainMenu.pas`/`ASNInvoice.pas`+`EDI856Object.pas` / `REPORT_EDI856` | **NOT BUILT** | no file-writer/seam; `EDI856Object` builder not ported |
| 5 | **EDI 810 (invoice) outbound** | `EDI810Object.pas`,`Write810File.pas` / `REPORT_EDI810` | **NOT BUILT** | not exercised on 6/19, but the daily-revenue pair to #4; spec `reporting/reporting.md` + D6 `REPORT_EDI810` migrated (data only) |
| 6 | **EDI inbound import** (997 ack, 862 firm, 830 fc, 824) | `EDIUpload.pas` / `UPDATE_EINStatus` | **NOT BUILT** | 11 EDIIMP rows/day; no Ignition file-poll/parse; spec `edi/edi-upload.md` |
| 7 | **EIN / ASN status update** from 997 | `DataModule` / `UPDATE_EINStatus`,`UPDATE_ASNStatus` | **PARTIAL** | procs exist live; **no Ignition caller/UI** wires them |
| 8 | **Order worksheet (compute + display)** | `Order.pas` / forecast+partstock procs | **PARTIAL (display only)** | `Order/OrderSpike` view + commit path (PR #11) built; the **forecast-fill / Excel-export** path is not reproduced |
| 9 | **Order FILE generation** (supplier/logistics/archive `.ord` + Order Sheet xls) | `OrderFormCreateF.pas` | **NOT BUILT** | 27 ORDERF + 9 ORDERS rows; fixed-width files parsed downstream — **zero spike coverage** |
| 10 | **Renban breakdown** (FRS lots→renban, counter advance) | `RenbanOrder.pas` / `UPDATE_RenbanGroupCount` | **PARTIAL** | RenbanGroup CRUD master built; the **breakdown algorithm** (lots→trailer→renban + counter race, Carry 2) is NOT built |
| 11 | **Hot call** (urgent ASN add + `8HC` 856) | `HotCallEntry.pas` | **NOT BUILT** | depends on #2 + #4 |
| 12 | **Daily Shipping Assy Report** | `MainMenu.pas` / `REPORT_DailyShippingAssy` | **PARTIAL** | proc migrated under D6 (data); **no Perspective report/print view**; legacy Excel-print path is the failing one (§4) |
| 13 | **Other daily reports** (FirmOrder xls, 824, monthly) | `EDIUpload.pas`, monthly report units | **NOT BUILT** | data procs partly migrated; no views |
| 14 | **Shift/part-type/forecast selection (SELECT SFT)** | `DataModule.pas` comboboxes | **PARTIAL** | master-data views provide some lookups; not the order/ASN-context selectors |
| 15 | **Login / audit log** | `Logon.pas` / `LogActLog`→activity log | **PARTIAL** | landing hub + gateway auth exist; the **`LogActLog` activity-trail equivalent** (every action audited) is not reproduced |
| 16 | **Master-data CRUD** | 7 master units | **BUILT** | PR #3/#4 — the one fully-built operator-facing area |
| 17 | **Stock ledger / on-hand** | producers + `POST_StockMovement` | **BUILT** | PR #5–#9,#12 — foundation, but **not in this day's log at all** (operator never touched receiving/stocktaking on 6/19) |

**Net:** of the 17 daily capabilities, **2 BUILT**, **6 PARTIAL**, **9 NOT BUILT**. The two BUILT
items (master-data, stock ledger) are exactly the two the operator **barely or never used** on a
normal day. Everything the operator actually did all day (ASN, EDI, order files) is PARTIAL/NOT
BUILT.

---

## 6. Production-critical ranking (what the operator cannot work without)

Ranked by log frequency × dependency-chain criticality (revenue/compliance impact). The daily
revenue loop is **ASN → 856 out → 997 import → status flip → 810 invoice → 997 import**; break any
link and Toyota shipments/payments stop.

| Rank | Capability | Why critical | Spike gap |
|---|---|---|---|
| **1** | **ASN entry creation (#2)** | Every shipment to Toyota starts here; 23 ASN rows/day; nothing downstream exists without it | NOT BUILT |
| **2** | **EDI 856 outbound (#4)** + status flip (#7) | The ASN *is* the shipment notice; Toyota won't receive without it; daily, per line + hot-call | NOT BUILT |
| **3** | **EDI inbound import / 997 ack (#6,#7)** | 11 rows/day; closes the loop (Accept/Reject); unacked 856/810 = compliance/payment risk | NOT BUILT (procs exist, no caller) |
| **4** | **ASN invoice reconciliation (#3)** + **EDI 810 out (#5)** | Billing accuracy → payment; qty corrections feed the invoice | NOT BUILT |
| **5** | **Order worksheet + FILE generation (#8,#9)** + renban breakdown (#10) | The replenishment side; supplier/logistics receive the `.ord` files daily; 36 ORDERF/ORDERS rows | PARTIAL (display) / NOT BUILT (files) |
| **6** | **Hot call (#11)** | Ad-hoc but urgent (line-down avoidance) | NOT BUILT |
| **7** | **Daily Shipping Assy + other reports (#12,#13)** | Operational visibility; currently the FAILING path (§4) | PARTIAL |
| **8** | **Master-data CRUD (#16)** + stock ledger (#17) | Necessary but low-frequency; edited occasionally, not the daily loop | BUILT |

**Bottom line for the architect:** the spike correctly built the *foundation* (ledger + master data
+ Order display + report-proc data layer), but the **operational EDI/ASN/order-file loop that
constitutes ~half the operator's day and 100% of the Toyota revenue path is not yet built.**
Production readiness requires, in priority order: (1) ASN create/invoice screens + procs wired,
(2) the EDI 856/810 **file writers** (byte-exact, to `X:\EDIOut\`) + outbound seam, (3) the EDI
inbound **file-poll/parser** (997/862/830/824) wiring `UPDATE_EINStatus`, (4) the order-file
generator (`.ord` fixed-width + Order Sheet) to supplier/logistics/archive, and (5) server-side
report/worksheet rendering to retire the fragile client-Excel print path that produced every ERROR
row in the log.

---

## Confidence notes

- Every action→unit mapping is grounded by grepping the literal `LogActLog` string in the live
  `.pas` (cited file:line); all units confirmed live via `InventorySystem.dpr`.
- Proc bodies READ + quoted for `INSERT_ASNInfo`, `UPDATE_ASNStatus`, `UPDATE_EINStatus`. Other
  procs (`REPORT_EDI856`, `REPORT_EDI810`, `SELECT_ASNSeq`, FRS-renban inserts) confirmed to EXIST
  in `CreateInventory.sql` by name/line but **bodies not fully quoted here** — behavior inferred
  from call sites; verify before porting.
- Spike-built inventory is from enumerating `view.json` files + `cutover-readiness-checkpoint.md`
  §1; if a view was authored but not committed under `docs/`, re-confirm.
- **Data-dependent claim to confirm against golden:** the TIRE order run (rows 124–131) wrote
  supplier+archive but **no logistics (`TLDL`) file**, whereas WHEEL wrote all three. Confirm
  whether TIRE suppliers are configured `logistics=none` in supplier-master (intended) vs a code
  branch skipping logistics for TIRE part-type — check `OrderFormCreateF.pas:306` logistics
  condition against the supplier-master `VC_LOGISTICS_*` value for supplier `07451` (DUNLOP).
