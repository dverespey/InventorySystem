# Module Analysis: EDI Inbound + Transport — `EDIUpload`

> Area §5 EDI/Integration. Covers the INBOUND X12 ingest + the file-transport layer:
> `EDIUpload.pas`/`.dfm` (the single `Execute` method that polls the EDIIn directory, sniffs the
> ISA, validates the receiver DUNS, and dispatches per transaction set: **830 / 862 / 997 / 824 /
> 820**). Pairs with the OUTBOUND spec `docs/analysis/edi/asn-invoice.md` (810/856 generation +
> ASNSelect/ASNInvoice/InvoiceBreakdown/HotCallEntry). The 830 *parse* is owned by
> `docs/analysis/forecasting/forecast-breakdown.md`; this spec owns its **dispatch/transport**.

**Area:** EDI/Integration  **Status:** ✅ spec complete  **Analyst:** Claude / 2026-06-16

---

## 1. Legacy surface

- **Form / unit (live):** `EDIUpload.pas` + `EDIUpload.dfm`. Confirmed live in
  `InventorySystem.dpr:53` (`EDIUpload in 'EDIUpload.pas' {EDIUpload_Form}`). The form is trivial:
  a `THistory` log pane (`Hist`), one `OKButton`, and a `TCopyFile` (`CopyFile`) component. All
  behavior is in **one method**, `TEDIUpload_Form.Execute` (`EDIUpload.pas:33-471`).
- **Entry point:** `MainMenu.EDIUpload_ButtonClick` (`MainMenu.pas:2897-2909`) — creates the form,
  `Show`, `Execute`, `Free`. The button (`EDIUploadBox`) is shown only when
  `Data_Module.fiGenerateEDI.AsBoolean` is true (`MainMenu.pas:2911-2916`, `FormShow`). So the EDI
  ingest UI is gated by the same per-site `GenerateEDI` flag as the outbound side.
- **Manual, operator-triggered, blocking.** There is **no scheduler**. An operator clicks the
  button; `Execute` runs synchronously, then **blocks the UI thread** waiting for the operator to
  click OK: `while not fclosed do begin Application.ProcessMessages; sleep(500); end`
  (`EDIUpload.pas:465-469`); `OKButtonClick` sets `fClosed:=TRUE` (`:495-498`).
- **`Write810File.pas` dead-code verdict** — *already adjudicated in the outbound spec*
  (`asn-invoice.md:21-25`): listed in `InventorySystem.dpr:57` so it compiles, but the entire body
  of `T810EDIFile.Execute` is commented out (`Write810File.pas:35`–`:118`). It performs no work.
  **DEAD.** Not re-analyzed here; recorded for completeness only.
- **Purpose:** the single inbound EDI funnel. Toyota/TEMA drops X12 files into a directory; the
  operator runs this screen to ingest them: route 830 release schedules to forecasting, turn
  862 firm orders / 824 application advice / 820 remittance into Excel reports, and land the 997
  functional ack as ASN/invoice accept/reject status.

---

## 2. Data touched

| Table | Read | Write | Notes |
|-------|:----:|:-----:|-------|
| `INV_ASN_MST` |  | ✓ | 997 ack → `VC_ASN_STATUS = @status WHERE IN_ASN_EIN=@EIN` when `@EINType='SH'` (via `UPDATE_EINStatus`, schema:8374). |
| `INV_INV_MST` |  | ✓ | 997 ack → `VC_INV_STATUS = @status WHERE IN_INV_EIN=@EIN` when `@EINType<>'SH'` (same proc). |
| `AD_GetSiteTMMDUNS` result | ✓ |  | **On `ALC_Connection` (the VehicleOrder DB), NOT the inventory DB** (`DataModule.dfm:691-697`). The receiver-DUNS allow-list / trading-partner validation. Param `@SiteTMMDUNS := delSL[4]`. See §4 D1. |
| `INV_FORECAST_DETAIL_INF` | (✓) | (✓) | **Only via the 830 hand-off** — written by `ForecastBreakdown_Form`, not by EDIUpload. See `forecast-breakdown.md` §2/§3. |
| Activity log (`Act_Connection`) |  | ✓ | `LogActLog('EDIIMP'/'ERROR', …)` on every transaction (`EDIUpload.pas` throughout). |

- **830 / 862 / 824 / 820 write NOTHING to the inventory DB from EDIUpload itself.** 830 is
  delegated (§4.3). 862, 824, 820 are **report-only** — they drive an Excel `OLEObject` to write
  `.xls` files into `fiReportsOutputDir`; the source comment is explicit ("Currently only a report
  is available … Store in Table" is a *TODO*, `EDIUpload.pas:310-314`). **820 remittance is NOT
  persisted to any payment table today.** Only the **997** has a DB side-effect.
- **Triggers:** none fire from this module. `UPDATE_EINStatus` is a plain UPDATE on
  `INV_ASN_MST` / `INV_INV_MST`; neither carries an EDI-relevant trigger.

---

## 3. Stored procedures / datasets used

| Proc / dataset | Conn | Operation | Business rule (from body / wiring) |
|----------------|------|-----------|------------------------------------|
| `AD_GetSiteTMMDUNS` (`SiteTMMDUNSDataset`) | **ALC_Connection** | SELECT | Receiver-DUNS allow-list. Called with `@SiteTMMDUNS := delSL[4]`; if `RecordCount > 0` the file is "for this site" and processing continues; else the file is **ignored** (`EDIUpload.pas:81-84, 437-440`). **Body unverified — lives in the VehicleOrder/ALC DB, absent from `Create Inventory.sql`.** Same proc the outbound side reads for site identity (`asn-invoice.md:48,77`). |
| `UPDATE_EINStatus` (`schema:8374`) | Inv_Connection | UPDATE | The 997-ack landing. `@EINType='SH'` → `INV_ASN_MST.VC_ASN_STATUS=@status WHERE IN_ASN_EIN=@EIN`; **else** → `INV_INV_MST.VC_INV_STATUS=@status WHERE IN_INV_EIN=@EIN`. Invoked as `dbo.UPDATE_EINStatus;1` via `DataModule.UpdateEINStatus` (`DataModule.pas:6753-6794`). Body cross-confirmed by the outbound spec (`asn-invoice.md:70`). **Body verified by cross-reference; not re-read this session — schema:8374, confidence high via outbound spec.** |

> **Snapshot-drift note** (cf. [[reference-schema-snapshot-vs-live]]): `AD_GetSiteTMMDUNS` is
> referenced by live code but absent from the checked-in inventory snapshot **because it lives in
> the separate ALC/VehicleOrder database** (`DataModule.dfm:692` `Connection = ALC_Connection`),
> not because it is missing in prod. Verify against the live ALC schema; do not report as a prod bug.

---

## 4. Business rules & edge cases

### 4.1 Transport model — local-filesystem polling, NOT FTP

**There is no FTP client in EDIUpload, and `TCopyFile` is a pure local byte-copy/move component**
(`Delphi-VCL-Components/Downloaded Version 7/Copy File/CopyFile.pas` — `CopyNow`,
`CopyFile.pas:118-228`, is a Win32 `FileOpen`/`FileRead`/`FileWrite` chunk copy with an optional
`DeleteFile(FCopyFrom)` when `MoveFile=True`; no sockets, no FTP). The "transport" is:

1. **Inbound drop dir:** `Data_Module.fiEDIIn` — INI `[DIRECTORIES] EDIIn`, default
   `c:\_Inventory_Control\EDIIn` (`DataModule.dfm:513-520`). An **external** process (TEMA mailer /
   VAN client, outside this app) deposits X12 files here. EDIUpload **polls it on demand**:
   `FindFirst(fiEDIIn + '\*.*', faAnyFile, …)` (`EDIUpload.pas:53`), iterating every non-directory
   entry (`sr.Attr <> 16` excludes folders, `:56`).
2. **Per file:** open as text (`Reset`), read line 1; if `pos('ISA', line1) > 0` it is treated as
   X12 (`:64-65`), else logged as "not a valid EDI document" and moved to `…\Archive\NOTEDI<name>`
   (`:443-452`).
3. **Archive on success:** after dispatch, `CopyFile.MoveFile := TRUE` and the file is **moved**
   into `fiEDIIn\Archive\` (`:419-435`). Archive name = `<EDIFileNumber><EIN>.EDI` when an EIN was
   parsed (997 path), else `<EDIFileNumber><delSL[10]>.EDI` (`:425-428`). `delSL[10]` is the ISA
   interchange control number (ISA13). **`EDIFileNumber` is only set inside the per-type branches**
   — an 830 sets it (`:91`) but the file is moved *after* the forecast hand-off returns; 824 sets
   it late (`:274`); the unknown-type `else` (`:414-417`) leaves `EDIFileNumber` from the previous
   loop iteration → **wrong archive name** (latent, §4.7).
4. **Outbound FTP is also external** — consistent with the outbound spec (`asn-invoice.md:33-34`:
   "an external mailer (out of this app) transmits them"). **The INI `[INIT] LocalFTP` flag
   (`fiLocalFTP`, `DataModule.dfm:294-300`) is a forecast/report-output concern, not read by
   EDIUpload** (forecast spec `forecast-breakdown.md:185`); EDIUpload never references it.

> **Transport verdict:** the whole transport layer is **filesystem move semantics over a
> shared/mapped directory**, with the real FTP/VAN leg done by a separate component. Success/failure
> handling is best-effort: copy errors are caught and logged (`:430-434`) but the source is **not**
> re-tried and **not** quarantined — a failed move leaves the file in place to be **re-ingested on
> the next run** (idempotency hazard, §4.7).

### 4.2 The per-site `delSL[4]` D1 hook — receiver-DUNS routing (central under D1)

This is the multi-site routing key. After confirming `ISA` on line 1, `Execute` splits it on `*`
into `delSL` via the local `splitString` helper (`EDIUpload.pas:71`, helper at `:473-493`) and
reads **`delSL[4]`** as the trading-partner / receiver DUNS:

```
splitString('*',fcl,delSL);                       // :71
Hist.Append('Trading Partner Search:'+delSL[4]);  // :73
SiteTMMDUNSDataset.Parameters['@SiteTMMDUNS'] := delSL[4];  // :78
... if RecordCount > 0 then  (process)  else  (IGNORE, :437-440)
```

- **Index semantics (watch off-by-one):** `splitString` (`:473-493`) emits the substring **before**
  each `*`, starting with the text before the first `*`. So for a standard ISA
  `ISA*00*          *00* ... *ZZ*RECEIVERID     *...`, `delSL[0]='ISA'`, `delSL[1]='00'`, … and
  **`delSL[4]`** is the **5th** element = ISA05 (sender qualifier) / ISA06 region depending on
  spacing. *(The X12 element this lands on is the receiver-DUNS used for trading-partner match;
  treat the index as authoritative-by-code, not by X12 spec name — confirm against a real inbound
  ISA during parity testing.)* `delSL[10]` (archive name, §4.1) is the 11th element = ISA13
  interchange control number.
- **Per-site filter:** an inbound file whose `delSL[4]` is **not** in `AD_GetSiteTMMDUNS` is
  **silently ignored** ("IGNORE: File(…) is not a trading partner file for this site", `:439`) and
  **not even archived** — it is left in the drop dir. Under multi-site, multiple sites polling the
  **same** drop directory would each ignore the others' files (the intended routing), but a file
  for site A left un-archived will be re-encountered by site A every run until processed.
- **Same hook as the 830 ingest.** `forecast-breakdown.md` (`§4`, lines 77-81, §6/D1 at 223-225)
  documents the **identical `delSL[4]` DUNS validation** inside `ForecastBreakdownF`. **The 830
  hand-off therefore validates DUNS twice:** once here (`:78`) before dispatch, and again inside
  `ForecastBreakdown.Execute` (`forecast-breakdown.md:79`). Not a bug — redundant defense — but the
  rebuild should validate **once** at the gateway and pass the resolved `site` down.

> **D1 (locked):** `delSL[4]` is the concrete multi-site receiver-routing key for **all** inbound
> EDI. In the rebuild it resolves the inbound file to a `sites` row (the DUNS allow-list becomes a
> per-site attribute), and **every** downstream write (997 status, 830 forecast, future 820/824
> persistence) is scoped to that site. The legacy `AD_GetSiteTMMDUNS` ALC-DB lookup collapses into
> the `sites` table (cf. `asn-invoice.md:228-230`).

### 4.3 830 ownership boundary — EDIUpload owns TRANSPORT/DISPATCH, ForecastBreakdownF owns PARSE

Resolved precisely. After the DUNS check, `Execute` reads to line 3, takes `copy(fcl,4,3)` as the
transaction-set id (`:86-88`), and on `'830'` it **delegates the entire parse**:

```
if data='830' then begin
  EDIFileNumber:=data; Hide;
  ForecastBreakdown_Form := TForecastBreakdown_Form.Create(self);
  ForecastBreakdown_Form.filename    := fiEDIIn + '\' + sr.Name;   // :95
  ForecastBreakdown_Form.SupplierCode:= Data_Module.fiSupplierCode.AsString;  // :96
  ForecastBreakdown_Form.Show; ForecastBreakdown_Form.Execute; ForecastBreakdown_Form.Free;  // :97-99
  ...LogActLog('EDIIMP','EDI 830 Imported: '+sr.Name);
end
```

**Boundary:** EDIUpload performs **detection + file routing only** (sniff ISA, validate DUNS, peek
at the `'830'` tag, then hand the *filename* to ForecastBreakdownF and archive afterward). **All
830 X12 element parsing — the LIN/FST loop, `delSL[3]`/`delSL[5]` part/kanban, the week math, the
`INV_FORECAST_DETAIL_INF` upsert — lives entirely in `ForecastBreakdownF.UploadBreakDown`** and is
specified in `forecast-breakdown.md` (§4, lines 77-93). ForecastBreakdownF re-opens the same file
from `.filename` and re-runs its own ISA/`delSL[4]` validation (the double-validation in §4.2).
**Note:** `ForecastBreakdownF` is *also* reachable directly from its own menu entry with a
file-picker (forecast spec §1, lines 29-30) — so the forecast importer has **two entry points**
(EDIUpload dispatch + direct picker); the rebuild should keep one parse service callable from both.

### 4.4 997 functional ack — parse → ack → status-update (the only DB-writing inbound path)

On `data='997'` (`:186-252`). The 997 carries one or more **AK1** loops; each AK1 names the
acknowledged group and the following segment carries the accept/reject code:

```
Readln; data:=copy(fcl,1,3);                 // :190-191
while data = 'AK1' do begin
  EDIType := copy(fcl,5,2);                   // :194  AK1*<grp> -> chars 5-6 = group functional id ('SH','IN', etc.)
  EIN     := copy(fcl,8,9);                   // :195  chars 8-16 (9 chars) = the group control number == our EIN
  Readln; Status := copy(fcl,5,1);            // :196-197  next segment (AK9), char 5 = accept/reject code ('A'/...)
  Data_Module.EIN := StrToInt(EIN);
  Data_Module.EINStatus := Status;            // 'A' accepted, anything else => reject branch
  Data_Module.EINType   := EDIType;           // 'SH' => ASN(856), else => invoice(810)
  if Data_Module.UpdateEINStatus then (log accept/reject) else (log 'Unable to update');  // :203-244
  Readln; data:=copy(fcl,1,3);                // :246-247  advance; loop while next is AK1
end;
```

- **Element offsets (watch off-by-one).** `EDIType := copy(fcl,5,2)` reads chars **5–6** of the AK1
  segment (after `AK1*`, positions 1–4). `EIN := copy(fcl,8,9)` reads chars **8–16** (9 chars) —
  i.e. it assumes a **single-char** AK1 element-1 then `*` then a 9-char control number, fixed
  position. **`Status := copy(fcl,5,1)`** is read from the **segment immediately after AK1** — the
  code does **not** check it is `AK9`; it blindly takes char 5 of whatever the next line is. If the
  997 interleaves `AK2`/`AK3`/`AK4` detail segments between `AK1` and `AK9`, **`Status` is parsed
  from the wrong segment** → wrong accept/reject. This is a real fragility (TEMA 997s for accepted
  groups are typically `AK1`+`AK9` with no detail, which is why it works in practice). Confidence:
  parse-by-fixed-offset confirmed in code; **assumption that next-segment == AK9 is unvalidated.**
- **Status → table routing.** `UpdateEINStatus` (`DataModule.pas:6753`) calls `UPDATE_EINStatus`
  (`schema:8374`) with `@EIN/@EINStatus/@EINType`. `@EINType='SH'` → updates `INV_ASN_MST`
  (the 856 ack); otherwise → `INV_INV_MST` (the 810 ack). This **confirms the outbound spec's
  claim** (`asn-invoice.md:70`) from the EDIUpload side: the 997 is the mechanism that flips an ASN
  `'S'`→`'A'/'R'` and an invoice `'S'`→`'A'/'R'`. The accept/reject branch only affects the **log
  text** (`:205-244`); the DB write happens regardless of `'A'` vs reject because `Status` is passed
  straight through as `@EINStatus`.
- **AK9 nuance.** The code keys the document status off the **AK1-group control number (= EIN)** and
  a **single status char**, not off AK9's group-level accept/partial/reject code distinctly. There
  is no handling of AK9 "partially accepted" (`P`) vs "rejected" (`R`) vs "accepted with errors"
  (`E`) — whatever char is at position 5 of the post-AK1 segment becomes the stored status verbatim.
- **EIN-keyed archive.** Because the 997 sets `EIN`, the file is archived as `997<EIN>.EDI`
  (`:425-426`) — useful traceability the other types lack.

### 4.5 862 firm order — report-only

On `data='862'` (`:105-185`). Reads the remittance/order date from `copy(fcl,17,8)` of a header
line (`:112`, reformatted `yyyy/mm/dd` at `:113`), skips two lines, then **drives Excel** from
`TemplateDir+'ReportTemplate.xls'` (`:122`) writing Part Number / Qty / Prod Date rows. Parse loop:
- Outer `while data <> 'CTT'` (`:142`) — one block per part; part number from `copy(fcl,9,12)`
  (`:144`).
- Inner `while data <> 'SHP'` advances to the ship segment (`:146-150`), then peels Qty by `*`
  splitting (`copy(fcl,8,…)` then to first `*`, `:152-155`) and prod-date from the next-but-one
  `*` field `copy(…,1,8)` reformatted (`:157-163`).
- Inner `while data <> 'TD5'` advances to the carrier segment (`:166-170`), then one more `Readln`.
- Output: `fiReportsOutputDir\FirmOrder<yyyymmddhhmmss00>.xls` (`:178`). **No DB write.**
- **Timestamp note:** `formatdatetime('yyyymmddhhmmss00',now)` = 8 date + 6 time + literal `00` =
  **16 chars** — correct, not a miscount (same pattern the forecast spec flags clean,
  `forecast-breakdown.md:188`).

### 4.6 824 application advice — error report only

On `data='824'` (`:253-305`). Wrapped in its own `try/except` (`:256/302`). Drives Excel
(Manifest / Part / Error-text columns), loops `while data <> 'SE*'` (`:279`) collecting **`NTE`**
segments by **fixed byte offsets**:

```
if data = 'NTE' then begin
  mysheet.cells[x,3] := copy(fcl,9,50);   // :283  error text, chars 9-58
  mysheet.cells[x,1] := copy(fcl,60,8);   // :284  manifest, chars 60-67
  mysheet.cells[x,2] := copy(fcl,69,12);  // :285  part number, chars 69-80
  INC(x);
end;
```

- **Loop terminator off-by-one trap:** `while data <> 'SE*'` but `data := copy(fcl,1,3)` — a real
  trailer is `SE*<count>*<control>`, so `copy(fcl,1,3)='SE*'` matches and the loop ends correctly.
  **But** if the file uses a different element separator, or the SE is the last line with no
  trailing data, `copy(fcl,1,3)` could be `'SE'`+EOL → **never matches `'SE*'` → infinite loop on
  EOF** (Readln past EOF returns empty strings forever). Fragile.
- **Fixed-offset NTE parse** (`:283-285`) assumes the 824 NTE is a fixed-width packed segment, not a
  `*`-delimited one — i.e. it relies on TEMA emitting the error text/manifest/part at exact
  columns 9/60/69. Any width drift silently mis-columns the report.
- Output: `fiReportsOutputDir\ReceivingAdvice<ts>.xls`; error count logged as `x-3` (`:301`).
  **No DB write — errors are NOT flagged on any ASN/invoice row** (the menu comment
  "compile error report for printing", `MainMenu.pas:2901`, is the whole behavior).

### 4.7 820 remittance — report only, NOT persisted (TODO in source)

On `data='820'` (`:306-413`). Source comment is explicit: *"Store in Table and print report …
Currently only a report is available"* (`:310-314`) — **so 820 writes nothing to the DB.** It:
- Skips to the `BPR` header (`:316-320`), splits it via a `TStringList` with `Delimiter:='*'`
  (`:322-324`) and reads `remittotal := StrToFloat(sl[2])` (BPR02 = payment amount) and
  `remitdate := sl[16]` (BPR16 = payment effective date) (`:326-327`).
- Loops `while header <> 'SE'` (`:368`), branching by segment:
  - `RMR` → `manifest := copy(fcl,8,8)` (chars 8-15), `total := StrToFloat(copy(fcl,18,…))`
    (amount paid for the remittance ref) (`:374-378`).
  - `IT1` → walks `*` fields: Qty, then skip 4 chars (`pos('*')+4`) to unit cost, then skip 7
    (`+7`) to a 12-char part number (`:380-390`). The `+4`/`+7` hard skips assume fixed
    intervening element widths — fragile.
  - `DTM` → `proddate := copy(fcl,9,8)` and **emits one Excel row** (manifest/part/proddate/qty/
    cost/total) (`:391-404`). Note `total` here is the **RMR remittance total**, repeated on every
    IT1 line under that RMR — the per-line "Item Total" column is the manifest total, not
    cost×qty (a reporting quirk worth confirming).
- `sl` (`TStringList`) is **created but never freed** (`:322`) → a small memory leak per 820 file.
- Output: `fiReportsOutputDir\Remittance<ts>.xls`. **No payment/error status table is updated.**

> **Both 820 and 824 are read-only reporting today.** If the rebuild must reconcile payments
> (820) or flag rejected ASNs from application advice (824), that is **new behavior** (the source
> has the 820 "Store in Table" as an unfinished TODO) — flag as §8 questions, not a port.

### 4.8 Idempotency / re-ingest hazard

There is **no client-side dup guard** (no P1 pattern) and **no proc-side `IF EXISTS` dedup** on the
inbound path — `UPDATE_EINStatus` is an unconditional UPDATE keyed on EIN. Re-running EDIUpload over
the same drop dir:
- **Processed files** were moved to `Archive\` so they are gone (safe) — **unless** the
  `CopyFile.CopyNow` move failed (caught + logged only, `:430-434`): the file stays and is
  **re-processed next run**. For 997 that re-applies the same status (idempotent); for 830 it
  re-runs the forecast upsert (idempotent per forecast spec); for 862/824/820 it produces a
  **duplicate report file** (different timestamp). Acceptable but noisy.
- **Ignored (wrong-site) files** are never moved (`:439`), so they are re-scanned every run — by
  design under shared-drop multi-site, but they accumulate.
- **`EDIFileNumber` carry-over bug:** the unknown-type `else` (`:414-417`) and any branch that
  doesn't set `EDIFileNumber` will archive using the **previous** file's `EDIFileNumber` (and
  `EIN`, which is only reset to `''` once at the top of `Execute`, `:47`, **not per file**). So a
  997 followed by a non-997 file can archive the second file as `…<staleEIN>.EDI`. Latent
  mislabel; confirm in parity testing.

---

## 5. UI / UX notes

- Single modal-ish screen: a scrolling `THistory` log + an OK button that only becomes visible
  (`:464`) after the batch completes, and whose click ends the busy-wait. Operator watches the log,
  clicks OK. No progress bar (the `TCopyFile` progress form flashes per archived file).
- **Keep:** the human-readable per-file audit trail (Found / Trading Partner / type / accept-reject
  / archive). **Modernize:** drop the blocking `sleep(500)` busy-wait; make ingest a Gateway timer
  job with a results table + a Perspective status view; make 820/824 actual data, not throwaway
  Excel.
- **Excel/OLE dependency** (`createOleObject('Excel.Application')`, `:119`,`:257`,`:329`) is a hard
  Windows + installed-Excel coupling for 862/824/820 reports — must be replaced (server-side
  XLSX/report).

## 6. Target design (Ignition) — strongest gateway-Python-service candidate

> **This module is the single strongest candidate in the whole system for a Gateway Python
> service** (peer to the outbound builder in `asn-invoice.md:209-235`). X12 parse → ack →
> status-update, plus directory polling, is exactly what does **not** belong in Perspective
> bindings or Named Queries. Pair it with the outbound `edi_outbound.py` into one `edi/` package.

- **Gateway service `edi_inbound.py` (Project Library) + a Gateway Timer/Scheduled script:**
  - `poll_edi_in(site)` replaces the manual button + filesystem scan: list the configured inbound
    location, for each file sniff `ISA`, parse the interchange header, resolve **`delSL[4]` → site**
    (D1), dispatch by transaction-set id, then move to archive (or a DB-tracked processed state).
    Make the loop **idempotent and crash-safe**: mark a file processed only after the DB side-effect
    commits; never re-emit a duplicate report.
  - **Use a real X12 parser** (segment/element split honoring ISA-declared separators), **not**
    `copy(,n,len)` byte offsets — but **preserve the exact element map** documented in §4.4–§4.7 as
    parity oracles (AK1 group-id at element 1, control-number == EIN; NTE error/manifest/part;
    820 BPR02/BPR16, RMR/IT1/DTM).
  - **FTP/VAN transport stays a separate concern** (as today an external mailer drops/pulls files).
    If the rebuild owns transport too, add an `edi_transport.py` (SFTP via a Gateway library) that
    writes into the same inbound location the parser polls — keep parse and transport decoupled.
- **997 ack** → a single Named Query `ein/ack` wrapping `UPDATE_EINStatus` (also used by the
  outbound spec, `asn-invoice.md:227`): `@EINType='SH'` updates the ASN status, else the invoice
  status, scoped to `site_id` under D1. Add **AK9-aware** logic: map AK9 group accept/reject
  (`A`/`E`/`P`/`R`) deliberately rather than blindly reading char 5 of the next segment.
- **830** → call the **shared forecast-ingest service** (forecast spec §6) with the resolved
  `site` + supplier; **validate DUNS once** at the gateway, not twice.
- **820 / 824** → **decide intent (§8):** if persistence is wanted, add real tables
  (`edi_remittance` / `edi_application_advice`) + Named Queries and reconcile against
  `INV_INV_MST` / `INV_ASN_MST`; reports become server-side renders, not Excel/OLE.
- **Multi-site (D1):** the `AD_GetSiteTMMDUNS` ALC-DB allow-list becomes a per-site DUNS attribute
  on `sites`; one shared inbound drop is fanned out by receiver DUNS to each site. **(D2)** all
  status updates resolve the ASN/invoice by its surrogate id once the EIN is matched. **(D3)** no
  hard deletes are involved on the inbound path; archival of processed EDI files is a logging/move
  concern, not a DB delete.

## 7. Migration plan for this module
- [ ] Stage 1 — stand up `edi_inbound.py` parse-only against captured sample files; **byte/element
      parity** vs the legacy parse map (§4.4–§4.7); render the ingest log in Perspective. Read-only
      DUNS→site resolution (D1).
- [ ] Stage 2 — enable the 997 status write through the `ein/ack` Named Query; wire the 830
      hand-off to the shared forecast service; move-to-archive becomes a DB-tracked processed flag
      (idempotent, fixes §4.8 carry-over + re-ingest).
- [ ] Stage 3 — replace Excel/OLE reports with server-side renders; **(decision-gated)** persist
      820 remittance + 824 advice to real tables and reconcile; optionally absorb SFTP transport.

## 8. Open questions (candidate D# decisions)
1. **`delSL[4]` exact X12 element + DUNS provenance.** Confirm against a real inbound ISA which X12
   element `delSL[4]` lands on (the code's index, not the spec name, is authoritative) and that
   `AD_GetSiteTMMDUNS` is the canonical receiver-DUNS allow-list to migrate to a per-site `sites`
   attribute (D1). (The outbound side reads the same proc, `asn-invoice.md:259`.)
2. **997 AK9 semantics.** Today the stored status is a single char read from the segment *after*
   AK1 (assumed AK9), passed verbatim into `VC_ASN_STATUS`/`VC_INV_STATUS`. Should the rebuild map
   AK9 group codes (`A` accept / `E` accept-with-errors / `P` partial / `R` reject) to distinct
   statuses, and must it tolerate `AK2/AK3/AK4` detail segments between AK1 and AK9? (Recommend:
   yes, parse AK9 explicitly.)
3. **820 remittance persistence.** Source has an unfinished "Store in Table" TODO; today 820 is
   report-only. Should the rebuild persist remittance (payment date, amount, per-manifest paid) and
   reconcile it against invoices (`INV_INV_MST`)? What is the payment-status model?
4. **824 application-advice action.** Today 824 produces an Excel error list and updates nothing.
   Should a 824 NTE error **flag/reject the named ASN** (`INV_ASN_MST` by manifest) automatically,
   or stay an operator report? Confirm the manifest/part columns are stable fixed-width offsets.
5. **Ingest trigger + concurrency.** Manual button today (no scheduler). Confirm a Gateway
   **scheduled poll** is the desired model, the cadence, and that a single shared inbound directory
   fanned out by DUNS is correct for multi-site (vs per-site drop dirs).
6. **Unexpected/duplicate-file handling.** Define behavior for: a file that fails to move
   (re-ingest hazard §4.8), an unknown transaction set (`else`, `:414`), and a wrong-site file
   (currently left un-archived). Should these quarantine instead of looping?

## 9. Test cases / parity checks
- Feed a captured **997** with AK1(`SH`)+AK9(`A`) → assert `INV_ASN_MST.VC_ASN_STATUS='A'` for that
  `IN_ASN_EIN`; AK1(`IN`)+AK9(reject) → `INV_INV_MST.VC_INV_STATUS` set to the reject char. Confirm
  the legacy proc routes SH→ASN, else→invoice (`UPDATE_EINStatus` schema:8374).
- 997 with detail segments (`AK2`/`AK3`) between AK1 and AK9 → legacy reads the wrong status char;
  rebuild reads AK9 correctly. (Adversarial parity.)
- **830** dropped in EDIIn → confirm EDIUpload hands the *filename* to the forecast service and the
  forecast outcome matches a **direct** ForecastBreakdownF picker run on the same file (single-parse
  parity; DUNS validated once in rebuild vs twice in legacy).
- Wrong-site file (`delSL[4]` not in allow-list) → ignored, **not** archived (legacy); rebuild
  routes/skips per site without leaving it to be re-scanned.
- 862/824/820 → byte-compare the extracted fields (offsets in §4.5–§4.7) against the generated
  report; verify the `SE*` terminator and the `+4`/`+7`/fixed-column assumptions on a real TEMA file.

---

## Cross-cutting findings

- **P12 (retry-recursion, wrong target) — ALREADY LOGGED, no new entry.**
  `DataModule.UpdateEINStatus` (`DataModule.pas:6789`) is **entry #8** in
  `docs/analysis/cross-cutting/datamodule-retry-target-bugs.md`: on exception it retries by calling
  **`UpdateRecProdRejInfo`** (a production-reject UPDATE on `INV_REJECT_INF`) instead of itself, and
  the shared `fRecordID`/`fEIN` state means the retry can stamp an **unrelated reject row** while the
  EDI document is **left in the wrong ACK status**. Confirmed reachable **from this module** — it is
  the proc behind the 997 ack (`EDIUpload.pas:203`). Cited, not re-filed. **No NEW P12 found in
  EDIUpload** (the form itself contains no DB-retry recursion).
- **No P1/P9 here.** EDIUpload performs no client-side dup guard and uses no shared `RecordID`
  itself; the only shared mutable state it touches is `Data_Module.fEIN/fEINStatus/fEINType` set
  immediately before `UpdateEINStatus` (`:199-201`) — which is exactly what feeds the P12 #8 path.
- **New (module-local) findings, not yet in any cross-cutting doc:** the `EDIFileNumber`/`EIN`
  carry-over archive mislabel (§4.8), the `820` `TStringList` leak (`:322`), the `824` `SE*`
  potential infinite-loop on EOF (§4.6), and the 997 "next segment == AK9" unvalidated assumption
  (§4.4). These are **single-module** parser fragilities (not the cross-DataModule retry family), so
  recorded here in §4 rather than promoted to the cross-cutting log.
