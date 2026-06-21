# EDI Inbound (M1): 997 Functional Ack + 824 Application Advice — Behavioral Spec

> **Scope (M1, minimal — the loop-closer for the 856/810 outbound).** This spec covers ONLY the two
> inbound transactions that close the outbound loop:
> - **997 functional ack** → match echoed control number to our EIN → flip the ASN(856)/invoice(810) status.
> - **824 application advice** → surface the TEMA reject (today: Excel report) → rebuild auto-flags + alarms.
>
> The other inbound transactions parsed by the same dispatcher — **830** release schedule,
> **862** firm order, **820** remittance, and any **824-other** — are **M2**. They are noted in
> §7 (Out-of-scope) but NOT specified here. The transport/dispatch shell that hosts all of them is
> already specified in `docs/analysis/edi/edi-upload.md` (the full module analysis); this doc is the
> M1 narrow drill-down on the two ack paths for the rebuild's inbound processor.

**Area:** EDI/Integration (inbound)  **Status:** spec complete  **Analyst:** Claude / 2026-06-21
**Source unit:** `EDIUpload.pas` (live), proc `UPDATE_EINStatus`, wrapper `DataModule.UpdateEINStatus`.

---

## 0. Liveness (confirmed)

- `EDIUpload.pas` is **LIVE**: `InventorySystem.dpr:53` — `EDIUpload in 'EDIUpload.pas' {EDIUpload_Form}`.
- All inbound behavior is one method, `TEDIUpload_Form.Execute` (`EDIUpload.pas:33-471`).
- Entry: `MainMenu.EDIUpload_ButtonClick` (per `edi-upload.md:20`), gated by per-site
  `Data_Module.fiGenerateEDI`. **Manual, operator-triggered, blocking** — no scheduler today.
- The 997-ack proc `UPDATE_EINStatus` is **LIVE and verified** in the authoritative dump
  `DB Schema/CreateInventory.sql:1711-1730` (body read this session, quoted in §3).

---

## 1. How the loop ties back: EIN == the echoed control number

The outbound builders write **one integer, the EIN**, into every X12 control-number position of the
856/810 (`%9.9d`, zero-padded):

- **856**: ISA13, GS06, ST02, BSN02-suffix, SE02, GE02, IEA02 = the EIN
  (`edi856-wire-format.md:21-46,53-54`). GS01=`SH`.
- **810**: same 7 positions = the **invoice** EIN `INV_INV_MST.IN_INV_EIN`
  (`edi810-wire-format.md:9,19-30,48-54`). GS01=`IN`.

TEMA's **997 echoes the control numbers back** (X12 rule: AK1 carries the acknowledged
functional-group control number = our **GS06**; AK2 would carry ST02). The legacy keys the status
flip off the **AK1 group control number**, which equals our **EIN**. That is the entire match
mechanism — **no manifest, no date, no site** is used to match; just the integer EIN.

> **EIN allocation is per-site SHARED across 856 and 810** (`AD_UpdateEIN`/`Site.SiteEIN`, one
> counter for both — `edi810-wire-format.md:53-54`). So 856 and 810 EINs **interleave** in one
> numeric space. The 997 disambiguates ASN-vs-invoice **not by the number** but by the **AK1 group
> functional id** (`SH` vs `IN`) → see §2 and the `@EINType` routing in §3. This is load-bearing:
> the EIN integer alone is not unique to a table; the SH/IN tag is what routes the write.

---

## 2. The 997 parse — the X12 segment walk (`EDIUpload.pas:186-252`)

### 2.1 Envelope detection (shared with all inbound)

Before any per-type branch (`EDIUpload.pas:64-88`):
1. `Readln` line 1; require `pos('ISA', fcl) > 0` (`:65`) — else "not a valid EDI document",
   moved to `Archive\NOTEDI<name>` (`:443-452`).
2. `splitString('*', fcl, delSL)` (`:71`) splits the ISA on `*`.
3. **DUNS/site routing on `delSL[4]`** (`:73-79`) — see §5.
4. Skip to line 3, `data := copy(fcl, 4, 3)` (`:86-88`) = the **transaction-set id** (`ST*<id>`,
   chars 4-6). `'997'` enters the ack branch (`:186`).

### 2.2 The AK1 loop (`:190-248`)

```
Readln; data := copy(fcl,1,3);          // :190-191  first segment after ST
while data = 'AK1' do begin             // :192      one iteration per functional group acked
  EDIType := copy(fcl,5,2);             // :194  AK1 chars 5-6  = group functional id: 'SH' or 'IN'
  EIN     := copy(fcl,8,9);             // :195  AK1 chars 8-16 (9 chars) = group control # == our EIN
  Readln; Status := copy(fcl,5,1);      // :196-197  NEXT segment, char 5 = the ack code ('A'/'E'/'P'/'R')
  Data_Module.EIN       := StrToInt(EIN);
  Data_Module.EINStatus := Status;      // :199-201  staged into shared DataModule fields
  Data_Module.EINType   := EDIType;
  if Data_Module.UpdateEINStatus then   // :203  DB write happens here regardless of A/E/P/R
    (log Accepted if Status='A' else log Rejected, by EDIType)   // :205-230
  else
    (log 'Unable to update 856/810 EIN(...)');                   // :232-244
  Readln; data := copy(fcl,1,3);        // :246-247  advance; loop continues while next seg is AK1
end;
Hist.Append('EDI 997 Processed: '+sr.Name);  // :250
```

**Multiple functional groups per 997 file: YES.** The `while data='AK1'` loop processes each AK1
in sequence, one `UPDATE_EINStatus` per AK1, each matched to its own EIN. A single 997 file can
therefore flip several ASNs/invoices.

### 2.3 Element offsets — exact byte map (parity oracle)

The legacy parses by **fixed `copy(s, start, len)` offsets**, NOT by `*`-splitting (unlike the ISA
line). Assumed AK1 layout: `AK1*<1-char fnid>*<9-digit control#>`:

| Field | Code | 1-based chars | Meaning | Assumption |
|-------|------|---------------|---------|------------|
| Group functional id | `EDIType := copy(fcl,5,2)` (`:194`) | 5–6 | `SH`=856/ASN, `IN`(or anything ≠ SH)=810/invoice | AK1 element-1 starts at char 5 (after `AK1*`) |
| Group control # = EIN | `EIN := copy(fcl,8,9)` (`:195`) | 8–16 (9 chars) | matches outbound GS06 `%9.9d` | element-1 is **exactly 1 char + `*`** so element-2 starts at char 8, fixed-width 9 |
| Ack code | `Status := copy(fcl,5,1)` (`:197`) | char 5 of the **next** segment | `A`/`E`/`P`/`R` | next segment **is assumed to be AK9**; not checked |

> **R-style data-confirm flags (name the exact cell to check against a real TEMA 997):**
> - **Offset assumption (EIN at chars 8–16):** `copy(fcl,8,9)` only lands on GS06 if AK1's
>   functional-id element is **exactly one character** (e.g. `SH`→ but wait — `EDIType:=copy(fcl,5,2)`
>   reads **2** chars at 5-6, yet `EIN:=copy(fcl,8,9)` starts at 8, implying a `*` at char 7, i.e. a
>   **1-char** element-1). The 2-char read at 5-6 vs the 9-char read starting at 8 is internally
>   inconsistent unless TEMA writes `AK1*XX*<ctrl>` where element-1 is 2 chars (`SH`) and the `*` is
>   at char 7. **Confirm against a captured 997:** the actual char positions of the `*` separators
>   and the control number. If TEMA uses a variable-width control number or a different element-1
>   length, both reads drift. **This is the #1 thing to verify against the golden inbound 997.**
> - **AK9 assumption:** `Status` is read from char 5 of *whatever line follows AK1* with no
>   `if copy(...)='AK9'` guard. If TEMA interleaves `AK2`/`AK3`/`AK4` transaction-set detail between
>   AK1 and AK9, `Status` is parsed from the **wrong segment** → wrong accept/reject. Works in
>   practice because accepted-group 997s are typically `AK1`+`AK9` with no detail. **Confirm whether
>   TEMA emits AK2/AK3 for rejects.**

### 2.4 AK2/AK3/AK4 detail: SKIPPED (not read)

The legacy **does not read** AK2 (transaction-set response), AK3 (data-segment note), or AK4
(data-element note). It reads exactly two lines per group — AK1 then the immediately-following
segment (assumed AK9) — then advances looking for the next AK1. So **per-transaction-set and
per-segment error detail is discarded** for the 997. Only the group-level (AK1→AK9) result lands.

### 2.5 AK9 — codes and what the legacy does with them

The code reads a single char (`Status`) and passes it **verbatim** into `UPDATE_EINStatus` as
`@EINStatus`. The branch logic (`:205-230`) ONLY affects **log text**: `Status='A'` logs
"Accepted", **anything else** logs "Rejected". The **DB write is identical** for all codes — the
char is stored as-is into `VC_ASN_STATUS` / `VC_INV_STATUS` (`varchar(1)`).

**Standard X12 AK901 codes and their fate in the legacy:**

| AK901 | X12 meaning | Legacy log | Stored char | Displays as (downstream CASE) |
|-------|-------------|------------|-------------|-------------------------------|
| `A` | Accepted | "Accepted" | `A` | **Accepted** |
| `E` | Accepted, but errors noted | "Rejected" | `E` | **(blank — no CASE match)** |
| `P` | Partially accepted | "Rejected" | `P` | **(blank — no CASE match)** |
| `R` | Rejected | "Rejected" | `R` | **Rejected** |
| `M`/`W`/`X` | (auth/security failures) | "Rejected" | that char | **(blank)** |

> **Keystone data-dependent finding (name the cell):** The downstream status-display proc only
> recognizes **four** chars. `SELECT_ASNStatus` `CASE VC_ASN_STATUS` (`CreateInventory.sql:1955-1959`
> and `:1979-1983`) maps **only** `A`→Accepted, `S`→Sent, `C`→Create File, `R`→Rejected.
> Invoice side `CASE i.VC_INV_STATUS` (`CreateInventory.sql:3254,3278`) is the same family.
> **Therefore an AK9 of `E` or `P` stores fine (varchar(1)) but renders as BLANK in the ASN/invoice
> status screens** — the operator sees an empty status, not "errors" or "partial". This is a real
> fidelity gap to confirm against the golden: **check what AK901 value TEMA actually returns for an
> accepted-with-corrections ASN** — if TEMA only ever sends `A` or `R` to this supplier, the gap is
> latent; if it sends `E`/`P`, the legacy already silently mis-displays them today.
>
> **Rebuild MUST:** map AK9 codes deliberately (`A`→accepted, `R`→rejected, `E`/`P`→a distinct
> "accepted-with-errors"/"partial" status that the status view renders), NOT store the raw char.

### 2.6 The status flip — `UpdateEINStatus` → `UPDATE_EINStatus`

Wrapper `DataModule.UpdateEINStatus` (`DataModule.pas:6753-6797`) calls `dbo.UPDATE_EINStatus;1`
with `@EIN` (int, `fEIN`), `@EINStatus` (`fEINStatus`), `@EINType` (`fEINType`). On error it
**ShowMessages + raises**, then in the `except` does the P12 retry (§6). Returns `boolean` —
`EDIUpload` only uses it to choose accept-vs-fail log text; it always returns TRUE unless the proc
itself throws.

Proc body — **verified** (`CreateInventory.sql:1711-1730`):

```sql
CREATE PROCEDURE [dbo].[UPDATE_EINStatus]
    @EIN integer, @EINStatus varchar(1), @EINType varchar(2)
AS BEGIN
    SET NOCOUNT ON;
    if @EINType = 'SH'
        UPDATE INV_ASN_MST SET VC_ASN_STATUS = @EINStatus WHERE IN_ASN_EIN = @EIN
    else
        UPDATE INV_INV_MST SET VC_INV_STATUS = @EINStatus WHERE IN_INV_EIN = @EIN
END
```

**Behavior, exactly:**
- `@EINType='SH'` → `INV_ASN_MST.VC_ASN_STATUS = @EINStatus WHERE IN_ASN_EIN = @EIN` (the 856/ASN ack).
- **else** (any value ≠ `'SH'`, including `'IN'`) → `INV_INV_MST.VC_INV_STATUS = @EINStatus WHERE
  IN_INV_EIN = @EIN` (the 810/invoice ack).
- **Unconditional UPDATE**, keyed solely on EIN. No `IF EXISTS`, no rowcount check, no OUTPUT param,
  no return code. If no row matches `@EIN` (wrong/unknown EIN), it **silently updates 0 rows** and
  the proc still succeeds → the wrapper logs "UPDated EIN(...)" and EDIUpload logs "Accepted"/
  "Rejected" **even though nothing changed**. (Idempotency hazard, §8.)
- **NO site scoping.** `WHERE IN_ASN_EIN=@EIN` only. Under multi-site with a shared EIN space this
  is a collision risk — see §5/§8.
- **Does NOT touch `VC_LAST_UPDATE`** (unlike sibling procs at `:1695`,`:1568` that set it). So the
  "Status Updated" date shown in `SELECT_ASNStatus` (`:1960`) does **not** advance on a 997 ack —
  it stays at the value from the original send. Confirm if that matters for the rebuild's audit.

Column types (`CreateInventory.sql`): `VC_ASN_STATUS varchar(1) NOT NULL` (`:695`),
`VC_INV_STATUS varchar(1) NOT NULL` (`:535`); `IN_ASN_EIN int NOT NULL` (`:694`),
`IN_INV_EIN int NOT NULL` (`:534`).

---

## 3. The 824 parse — application advice (`EDIUpload.pas:253-305`)

The 824 carries TEMA's **application-level** reject of a prior ASN (e.g. receiving discrepancy).
Today it is **report-only** — it writes **NOTHING** to the DB and flags **nothing** on any
ASN/invoice row.

```
try
  excel := createOleObject('Excel.Application');          // :257  Excel/OLE dependency
  excel.workbooks.open(TemplateDir+'ReportTemplate.xls'); // :260
  mysheet.cells[3,1]:='Manifest Number';                  // :264  column headers
  mysheet.cells[3,2]:='Part Number';                      // :267
  mysheet.cells[3,3]:='Error Text';                       // :270
  Readln; data:=copy(fcl,1,3); x:=4;                      // :276-278
  while data <> 'SE*' do begin                            // :279  loop to trailer
    if data = 'NTE' then begin                            // :281  each NTE = one error line
      mysheet.cells[x,3] := copy(fcl,9,50);               // :283  Error Text   = NTE chars 9-58
      mysheet.cells[x,1] := copy(fcl,60,8);               // :284  Manifest #   = chars 60-67
      mysheet.cells[x,2] := copy(fcl,69,12);              // :285  Part Number  = chars 69-80
      INC(x);
    end;
    Readln; data:=copy(fcl,1,3);                          // :289-290
  end;
  excel.SaveAs(fiReportsOutputDir+'\ReceivingAdvice<ts>.xls');  // :294
  Hist.Append('ASN ERRORS: '+IntToStr(x-3)+' total ...');       // :301  error count = x-3
except
  Hist.Append('EDI 824 Failed to open excel for proecessing');  // :303  (sic)
end;
```

### 3.1 What the 824 extracts (fixed-width NTE, parity oracle)

Each `NTE` segment is parsed by **fixed byte offsets** (NOT `*`-split) — TEMA packs the fields:

| Excel col | Field | NTE 1-based chars | `copy` |
|-----------|-------|-------------------|--------|
| 3 | Error Text | 9–58 (50 chars) | `copy(fcl,9,50)` `:283` |
| 1 | Manifest Number | 60–67 (8 chars) | `copy(fcl,60,8)` `:284` |
| 2 | Part Number | 69–80 (12 chars) | `copy(fcl,69,12)` `:285` |

- **Manifest (chars 60-67, 8 chars)** is the key that ties the 824 back to the shipment — it is the
  **manifest number**, the same key the 820 RMR uses (`edi-upload.md:265`) and the manifest carried
  on the ASN. This is the join the rebuild's auto-flag needs. **Confirm against a golden 824** that
  the manifest truly sits at chars 60-67 and that it matches an `INV_ASN_MST`/manifest column —
  these fixed offsets are the single most fragile thing in the 824 path (any width drift silently
  mis-columns every field).
- **Error count** logged as `x-3` (`:301`); `x` starts at 4, so `x-3` = number of NTE rows.

### 3.2 What the Excel report DOES (the info the rebuild must reproduce)

The Excel sheet IS the entire 824 behavior — it is how the **reject reaches a human**. It surfaces,
per error line: **Manifest #, Part #, Error Text**. There is no DB flag, no alarm, no automatic ASN
status change. The operator must open `fiReportsOutputDir\ReceivingAdvice<ts>.xls` and read it
(the `Hist` log line `:301` even tells them to). So the operator-facing payload to preserve is:
**(manifest, part, free-text reason) × N error lines + a total count.**

### 3.3 824 — file handling note

The 824 branch is the ONLY type wrapped in its own `try/except` (`:256/302`). It does **not** set
`EDIFileNumber` until `:274` (mid-branch), and it never parses/sets `EIN`, so on archive (§5) it
falls to the `else` name `<EDIFileNumber><delSL[10]>.EDI`. The `824` literal is captured into
`EDIFileNumber` at `:274` via `EDIFileNumber:=data` (where `data` is still `'824'` from `:88`).

### 3.4 824 auto-flag SCOPE — flag EVERY ASN carrying the rejected manifest (Q10, REBUILD decision)

The rebuild's auto-flag (Q10) replaces the Excel report. A 824 reject manifest can legitimately
appear on **more than one ASN's detail rows** — **proven on Live (read-only)**: a manifest maps to
up to **3 distinct `IN_ASN_ID`** (e.g. manifest `52066074` → 7 detail rows across **3** ASN IDs;
`MAX(distinct ASN per manifest) = 3`). **Decided scope (the rebuild flips ALL of them):** for each
rejected manifest, set `VC_ASN_STATUS='R'` on **every** `INV_ASN_MST` whose
`INV_ASN_DETAIL_MST.VC_MANIFEST_NUMBER` matches (a single set-based join `UPDATE`, so
`asnsFlagged` = the total `@@ROWCOUNT`, which may be 2 or 3).

**Rationale:** a reject is **recoverable** (D3); **over-flagging is the safe failure mode**. Flagging
every shipment that carried the rejected manifest can at worst mark a recoverable ASN that did not
strictly need it — vs. the *unsafe* alternative of missing an ASN that genuinely carried the rejected
manifest and leaving it looking accepted. The per-line alarm rows (manifest/part/reason) are written
regardless of how many ASNs matched, so the operator-facing detail is complete either way. **No
unrelated ASN is touched** (only those whose detail carries the exact rejected manifest). Site
scoping (`-- M4`) will further narrow the flip to the resolving site's ASNs once `INV_ASN_MST`
carries `IN_SITE_ID`. (Legacy flagged **nothing** — Excel-only — so this is new behavior, not a
divergence; adversary SHOULD-FIX-2.)

---

## 4. (M2 cross-reference only) the other inbound types

NOT specified here — see `edi-upload.md` §4.3/§4.5/§4.7. For M1 awareness only:
- **830** release schedule (`:89-104`) → delegated whole to `ForecastBreakdown_Form.Execute`.
- **862** firm order (`:105-185`) → Excel report only.
- **820** remittance (`:306-413`) → Excel report only, **D12 RESOLVED: stays report-only**
  (`edi-upload.md:369`).
- Unknown type (`else`, `:414-417`) → logged, not processed.

---

## 5. Inbound file source + DUNS/site routing (shared by 997 & 824)

### 5.1 Source dir

- Inbound drop dir: `Data_Module.fiEDIIn` — INI `[DIRECTORIES] EDIIn`, default
  `c:\_Inventory_Control\EDIIn` (`edi-upload.md:82`). An **external** TEMA mailer/VAN client (NOT
  this app) deposits the X12 files. EDIUpload **polls on demand**: `FindFirst(fiEDIIn+'\*.*',
  faAnyFile, sr)` (`EDIUpload.pas:53`), iterating every non-folder entry (`sr.Attr <> 16`, `:56`).
- **No FTP in this app.** `TCopyFile.CopyNow` is a pure local Win32 file move (`edi-upload.md:77-80`).

### 5.2 DUNS → site resolution (`delSL[4]`)

```
splitString('*', fcl, delSL);                             // :71  split ISA on '*'
SiteTMMDUNSDataset.Parameters['@SiteTMMDUNS'] := delSL[4];// :77-78
SiteTMMDUNSDataset.Open;
if RecordCount > 0 then (process) else (IGNORE, :437-440);// :81 / :439
```

- `delSL[4]` is the **5th** `*`-delimited element of the ISA (the `splitString` helper at `:473-493`
  emits the text **before** each `*`, starting with `delSL[0]='ISA'`). It is the **receiver/trading-
  partner DUNS** used to decide "is this file for this site?" Dataset `SiteTMMDUNSDataset` runs
  `AD_GetSiteTMMDUNS` on **`ALC_Connection`** (the VehicleOrder/ALC DB, NOT the inventory DB —
  `edi-upload.md:45,63`); proc body lives in the ALC schema, not `CreateInventory.sql`.
- **No match → file SILENTLY IGNORED and NOT archived** (`:439`) — left in the drop dir, re-scanned
  every run (by design for shared-drop multi-site, but accumulates).
- **Same DUNS index as the 830 path** — and the 830 hand-off re-validates `delSL[4]` again inside
  `ForecastBreakdownF` (double-validation, `edi-upload.md:134-138`). The rebuild validates **once**.

> **Data-confirm (name the cell) — PENDING A GOLDEN INBOUND ISA (BLOCKER for the routing claim):**
> `delSL[4]`'s exact X12 element is **authoritative-by-code, not by spec name** — confirm against a
> captured inbound ISA which element (ISA05/ISA06/receiver-ID region) actually lands at index 4 given
> TEMA's real ISA spacing. The index, not the X12 label, is what the rebuild's DUNS guard must
> replicate.
>
> **Heads-up from OUR OWN OUTBOUND layout (verified, `EDI856Object.pas:145-153` /
> `EDI810Object.pas:140-148`):** on a file *we* generate, `delSL[4]` = **ISA04** = **our** `SiteDUNS`,
> and the **TMM** DUNS sits at `delSL[8]` = **ISA08** (receiver ID). The legacy guard (and the
> rebuild, faithfully) matches `delSL[4]` against the **TMM** column (`Site.SiteTMMDUNS` /
> `INV_SITES.VC_TMM_DUNS`) — i.e. against the Toyota DUNS, at the ISA04-security slot. For an inbound
> TEMA file to ever match, TEMA's inbound ISA must place its OWN (Toyota) DUNS at the ISA04 slot — a
> **non-standard** layout (ISA04 = "Security Information," normally blank/zeros). If TEMA's inbound
> ISA is X12-standard, **both legacy and rebuild quarantine every TEMA file.** The rebuild is
> **legacy-faithful on the column** (not a regression), but the **routing correctness is UNPROVABLE
> until a real inbound TEMA 997/824 ISA is captured** to pin the element — same honest stance as the
> 856/810 byte-parity-pending-golden. The test fixtures hard-code the DUNS at `delSL[4]`, so the green
> DUNS tests prove the guard **mechanics only**, NOT the element index.

### 5.3 Poll/trigger + ACK-back

- **Trigger = manual button**, blocking (`while not fclosed ... sleep(500)`, `:465-469`). No timer.
- **No ACK-back.** EDIUpload sends **nothing** to TEMA on inbound — it does not generate a 997 in
  response to inbound docs, and it does not acknowledge the 824/997. It is purely a consumer.
  (The outbound 856/810 are built elsewhere; the 997 we *receive* is TEMA acking *us*.)

### 5.4 Archive / idempotency (file handling)

After a type branch returns (`:418-435`):
- `CopyFile.MoveFile := TRUE`; **moves** the file to `fiEDIIn\Archive\`.
- Archive name: `<EDIFileNumber><EIN>.EDI` if an EIN was parsed (997 path, `:425-426`), else
  `<EDIFileNumber><delSL[10]>.EDI` (`:428`; `delSL[10]` = ISA13 interchange control #). So a 997
  archives as `997<EIN>.EDI` (per-EIN traceable); 824 archives as `824<ISA13>.EDI`.
- **Re-ingest hazard:** if the move **fails** (caught + logged only, `:430-434`), the file stays and
  is **re-processed next run**. For 997 the re-apply is idempotent (`UPDATE_EINStatus` is an
  unconditional UPDATE to the same status). For 824 it regenerates a **duplicate** Excel report
  (new timestamp). (`edi-upload.md:281-297`.)
- **`EIN`/`EDIFileNumber` carry-over bug:** `EIN` is reset to `''` **once** at `:47` (top of
  `Execute`), **not per file**. So after a 997 sets `EIN`, a following **824** (which never sets
  `EIN`) archives using the **stale 997 EIN** in the `if EIN<>''` branch (`:425-426`) →
  **mislabeled archive name** (`997-EIN` reused on the 824). Latent; confirm in parity testing.

---

## 6. Cross-cutting hazard — P12 (retry-recursion, wrong target)

`DataModule.UpdateEINStatus` (`DataModule.pas:6786-6794`) is **entry #8** in the cross-cutting
retry-target-bug log. On exception it does NOT retry itself — it calls **`UpdateRecProdRejInfo`**
(a production-reject UPDATE on `INV_REJECT_INF`):

```pascal
except on e:exception do begin
  fErrorCount := fErrorCount + 1;
  If fErrorCount < 3 Then UpdateRecProdRejInfo   // <-- WRONG target; not UpdateEINStatus
  Else LogActLog('ERROR', 'FAILED UPDated EIN(...)');
end;
```

Because `fEIN`/`fEINStatus`/`fRecordID` are **shared mutable DataModule state**, a transient failure
on the 997 ack can fire an **unrelated reject-row UPDATE** while the EDI document is **left in the
wrong/unchanged ack status**. Confirmed reachable from this module (`EDIUpload.pas:203`). Already
logged — do NOT re-file. **Rebuild MUST NOT carry this**: a 997 ack failure retries the *ack*, not a
reject write, and uses no shared mutable id.

---

## 7. Out of scope for M1 (M2 backlog — noted, not specified)

- 830 / 862 / 820 / unknown-type dispatch (see `edi-upload.md`).
- 824 detail beyond NTE manifest/part/text (AK-style segment detail, multiple ST loops in one 824).
- 997 AK2/AK3/AK4 per-segment/per-element error capture (legacy discards it; M2 may want it).
- 820 remittance persistence — D12: stays report-only.
- The Excel/OLE report replacement strategy for the non-ack types.

---

## 8. What the rebuild's inbound processor MUST reproduce (M1)

**997 (faithful behavior + the fixes the rebuild owes):**
1. Poll the inbound location (gateway scheduled job, Q11) → for each file, sniff `ISA`, split, and
   **resolve `delSL[4]` → site** via the per-site DUNS attribute (the `AD_GetSiteTMMDUNS` allow-list
   collapses into a `sites` DUNS column, D1). Wrong-site → skip without re-scan churn.
2. Walk the **AK1 loop** — one status flip per AK1 (multiple groups per file). Match the **AK1 group
   control number = our EIN** (an integer). Use a **real X12 parser** (honor ISA separators), but
   keep the legacy element map (§2.3) as the **parity oracle**.
3. Route by **AK1 group functional id**: `SH` → ASN, `IN`(or ≠SH) → invoice — exactly
   `UPDATE_EINStatus`'s `@EINType` branch. **Add site scoping** to the WHERE (legacy keys on EIN
   only; under one shared DB + interleaved per-site EINs this can collide — match
   `IN_ASN_EIN=@EIN AND site_id=@site`).
4. **Map AK9 deliberately** (`A`/`E`/`P`/`R`) — do NOT store the raw char. Legacy stores it verbatim,
   and `E`/`P` then render **blank** in the status views (§2.5); the rebuild must give `E`/`P` a real
   status. **Tolerate AK2/AK3/AK4** between AK1 and AK9 (legacy reads char 5 of the next line
   blindly — fragile).
5. **Idempotent + crash-safe:** only mark the file processed after the status write commits; a
   0-row match (unknown EIN) must be surfaced, not silently logged as success. Drop the P12 retry
   (§6).

**824 (faithful info + the new auto-flag/alarm per Q10):**
6. Parse each `NTE` (loop to `SE`): extract **(Error Text, Manifest #, Part #)** per line + a total
   count (the §3.1 map is the oracle; confirm offsets on a golden 824).
7. Replace the Excel report with: **auto-flag the named ASN as rejected** (match on **manifest #**
   → `INV_ASN_MST`, scoped to site) + a **main-screen alarm** + **click-to-detail** showing the
   (manifest, part, reason) lines (decision Q10). This is **new behavior** — the legacy flagged
   nothing — but it must capture the **same operator-facing payload** the Excel sheet showed.
8. No ACK-back to TEMA (consumer-only). Archive/track processed files in a DB-tracked processed
   state (kills the re-ingest + EIN-carry-over bugs, §5.4).

### 8.1 REQUIRED M1 FOLLOW-ON — the E/P status-render arm (the Q6 fix is half-done without it)

`ak9_to_status` (`edi_inbound/code.py`) now STORES `E` (accepted-with-errors) and `P` (partial) as
**distinct** status chars, per Q6 — the deliberate fix for the legacy "store the raw char →
renders blank" bug (§2.5). **But there is no status-render decode on disk yet that recognizes `E`/`P`**
(a repo-wide search found only the legacy `A/S/C/R`-only CASEs at `CreateInventory.sql:1955-1959` /
`:3254-3257`, and no rebuild status-render NQ/view). So **a stored `E`/`P` currently renders BLANK
downstream — the exact bug Q6 set out to fix.** The fix is therefore **incomplete until the render arm
exists** (adversary SHOULD-FIX-3):

> **REQUIRED:** when the ASN/invoice status-DISPLAY Named Query / view is built, it **MUST** add
> decode arms `E → "Accepted w/ errors"`, `P → "Partial"` (and `R → "Rejected"`) alongside the
> existing `A`/`S`/`C` arms (`einstatus-status-flow-analysis.md §3`). There is no render NQ to extend
> *yet* — this is a hard M1 follow-on, not a silent drop. Confirm against a golden whether TEMA sends
> `E`/`P` to this supplier (Live is 100% `A`, so the gap is latent on parity data but real the moment
> TEMA sends one).

---

## 9. M1 parity checks

- **997 accept:** feed a captured 997 with `AK1*SH*<ein>` + `AK9*A*...` → assert
  `INV_ASN_MST.VC_ASN_STATUS='A'` for that `IN_ASN_EIN`. `AK1*IN*<ein>`+`AK9*R` →
  `INV_INV_MST.VC_INV_STATUS='R'`. (Routes SH→ASN, else→invoice, `UPDATE_EINStatus`
  `CreateInventory.sql:1722-1728`.)
- **997 offset oracle (R-style, confirm against golden):** verify the real `*` positions in
  `AK1` — that `copy(fcl,5,2)`=`SH`/`IN` and `copy(fcl,8,9)`=the 9-digit EIN on a captured TEMA 997.
  If TEMA's AK1 element-1 width or control-number width differs, both reads drift.
- **997 AK9 E/P:** feed `AK9*E` → confirm whether the status view renders blank (it will, per
  `CASE` at `:1955`) — quantify the fidelity gap; check what AK901 TEMA actually returns for this
  supplier (golden).
- **997 adversarial:** AK2/AK3 between AK1 and AK9 → legacy reads the wrong status char; rebuild must
  read AK9 correctly.
- **997 multi-group:** a 997 with two AK1 loops → two status flips, correct EIN each.
- **824 NTE oracle:** byte-compare extracted (manifest@60-67, part@69-80, text@9-58) against a golden
  824; confirm the manifest joins to an `INV_ASN_MST` row for the auto-flag.
- **Wrong-site file:** `delSL[4]` not in allow-list → ignored, not archived (legacy);
  rebuild routes/skips per site without re-scan churn.

---

## Confidence

- **High / body verified:** `UPDATE_EINStatus` body (`CreateInventory.sql:1711-1730`, read this
  session); `UpdateEINStatus` wrapper + P12 retry (`DataModule.pas:6753-6797`, read this session);
  997 AK1-loop + offsets and 824 NTE offsets (`EDIUpload.pas`, read this session); downstream status
  CASE (`CreateInventory.sql:1955-1959,1979-1983`); EIN-as-control-number tie-back (856/810 wire
  docs).
- **Body unverified:** `AD_GetSiteTMMDUNS` (lives in the ALC/VehicleOrder DB, not in
  `CreateInventory.sql`) — DUNS allow-list behavior inferred from call wiring + `edi-upload.md`.
- **Data-dependent (verify vs golden):** AK1 `*`/control-number byte positions (§2.3 #1);
  AK901 values TEMA actually returns (§2.5); `delSL[4]` exact ISA element (§5.2); 824 NTE fixed
  offsets + manifest join key (§3.1).
