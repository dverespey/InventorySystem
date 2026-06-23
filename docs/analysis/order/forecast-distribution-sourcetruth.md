# Forecast-Distribution Feed — Source-Truth Spec

**Legacy unit:** `ForecastBreakdownF.pas` (LIVE — `InventorySystem.dpr:28`).
**Feed name:** the per-supplier **forecast distribution** — forwards TEMA's forecast DOWN to each
sub-supplier as a `.frc` text file and/or an Excel workbook.
**Scope of THIS spec:** the OUTBOUND emitter — the `SELECT_ForecastSupplier` supplier loop at
`ForecastBreakdownF.pas:340-568`. The inbound 830 import (the part that POPULATES
`INV_BREAKDOWN_FC_INF`) is already specced/built at `docs/analysis/edi/inbound/830-862-forecast-import-spec.md`
(M2); this feed READS what M2 wrote.
**Confidence:** HIGH. All four procs, both tables, and the trigger callers read from source. The only
data-dependent items flagged inline ("confirm against golden").

---

## 0. TL;DR for the developer (the reuse map)

This is the **outbound complement of the 830 import** and is structurally the **same shape as the
already-built `.ord` order-file emit**. Build it by cloning that seam:

- **Driver/atomicity/supplier-loop/filename/`SELECT_SupplierInfo`/`_coerce_bit` pattern** → reuse
  `docs/analysis/order/project-library/order_file/code.py` verbatim in spirit. Same supplier break-detection,
  same `_clean_name`, same temp-then-rename publish, same `Output File Type` → channels mapping, same
  `_supplier_destinations` (supplier dir + Archive on LocalFTP).
- **Excel** → reuse the **M3 POI render harness** `docs/analysis/reporting/project-library/report_render/`
  (declarative lane: header band off i=1, 11-column flat table off i=2). **No `.xls` template needed** — POI
  builds the sheet fresh (the legacy template `ForecastTemplate.xls` carries only header formatting; the
  data rows are written cell-by-cell).
- **SiteSupplierCode (P6 crash)** → apply the **identical H4 adjudication the `.ord` rebuild used**: build the
  working **non-sendsite** line for ALL suppliers, do NOT reproduce the crash, defer the real multi-site
  leading field to M4 (it does not exist single-site).

**No separate architect pass is required** — there is no design fork beyond what the `.ord` emit already
resolved. The developer can build directly from this doc + those two seams.

---

## 1. The driving query/proc

The supplier-emit loop opens **`dbo.SELECT_ForecastSupplier;1`** (`ForecastBreakdownF.pas:350`), one param:

```
@WeekDate := formatdatetime('yyyymmdd', now)   -- today, yyyymmdd  (:353)
```

Proc body (`CreateInventory.sql:2058`):

```sql
CREATE PROCEDURE [dbo].[SELECT_ForecastSupplier]
	@WeekDate		varchar(8)
 AS
	SELECT * FROM INV_BREAKDOWN_FC_INF
		WHERE  VC_WEEK_DATE > @WeekDate
	order by vc_supplier_code, vc_part_number, vc_week_date
```

**It is `SELECT *` over a single table — `INV_BREAKDOWN_FC_INF`.** This is the SAME table the 830 import
owns and populates (M2 spec line 118: *"This module owns it — the day-qty buckets the Order reads"*). It is
NOT a derived view and NOT a join — there is no ratio/breakdown math here. The emit just **serializes the
already-day-spread week-buckets** that M2 wrote.

**Filter:** `VC_WEEK_DATE > @WeekDate` (strict `>`), i.e. only week-buckets whose week date is **strictly
after today** — "the next week out and beyond." `VC_WEEK_DATE` is `varchar(8)` (`yyyymmdd`); the comparison
is lexicographic, which is correct for that format.

**ORDER BY `vc_supplier_code, vc_part_number, vc_week_date`** — the **`vc_supplier_code` lead drives the
new-file-per-supplier break detection** (§2).

### Every yielded column (`SELECT *`) and its source

All 12 columns come from the one table `INV_BREAKDOWN_FC_INF` (`CreateInventory.sql:578`):

| Column | Type | Read at | Used for |
|---|---|---|---|
| `IN_WEEK_NUMBER` | `int` | `:471` (Excel col 4), `:499` (text `%.2d`) | week number |
| `VC_WEEK_DATE` | `varchar(8)` | `:470` (Excel col 3), `:498` (text raw) | week date `yyyymmdd` + the loop filter |
| `VC_SUPPLIER_CODE` | `varchar(5)` | `:366`/`:412` (break key), `:490`/`:494` (text) | supplier break + text body |
| `VC_PART_NUMBER` | `varchar(12)` | `:469` (Excel col 2), `:497` (text raw) | part |
| `VC_SIZE_CODE` | `varchar(10)` NULL | `:468` (Excel col 1) | size (Excel only; NOT in text) |
| `IN_QTY1` | `int` NULL | `:472` (Excel col 5), `:500` (text `%.5d`) | day-1 qty |
| `IN_QTY2` | `int` NULL | `:473` / `:501` | day-2 qty |
| `IN_QTY3` | `int` NULL | `:474` / `:502` | day-3 qty |
| `IN_QTY4` | `int` NULL | `:475` / `:503` | day-4 qty |
| `IN_QTY5` | `int` NULL | `:476` / `:504` | day-5 qty |
| `IN_QTY6` | `int` NULL | `:477` / `:505` | day-6 qty |
| `IN_QTY7` | `int` NULL | `:478` / `:506` | day-7 qty |

**`SiteSupplierCode` is read at `:488` but is NOT a column of this table** (or any table — see §5). That is
the P6 crash.

> The columns the task brief listed under "main loop dataset" mix TWO datasets: `VC_SIZE_CODE` … `IN_QTY1..7`
> come from `SELECT_ForecastSupplier` (this table). `SiteSupplierCode`, `Supplier Name`, `Supplier Code`,
> `Directory`, `Output File Type`, `Site Number in Order` come from a SEPARATE per-supplier
> `SELECT_SupplierInfo` lookup (§2). The Pascal reads them off `Data_Module.Inv_StoredProc`, a different
> dataset than the `Inv_DataSet` driving the row loop. Do not conflate them.

---

## 2. The supplier loop + new-file logic

The loop (`:364-515`) walks the `SELECT_ForecastSupplier` cursor (sorted by supplier). `lastsupplier`
(init `''`, `:169`) holds the supplier of the file currently open.

**Break detection (`:366`):** when `VC_SUPPLIER_CODE <> lastsupplier`:
1. **Close the previous supplier's outputs** (`:368-410`): SaveAs+close the Excel workbook if open
   (`:374-386`), CloseFile the `.frc` text file + the Archive `.frc` if `LocalFTP` (`:388-396`). Wrapped in
   try/except — a close failure is logged, not fatal.
2. `lastsupplier := VC_SUPPLIER_CODE` (`:412`).
3. **Per-supplier config lookup** via `dbo.SELECT_SupplierInfo;1` (`:419`), one param
   `@SupCode := lastsupplier`. (Note: `@Logistics` defaults 0 — this feed does NOT pass it, unlike the
   `.ord` emit which passes `@Logistics=1`.)
4. From that row:
   - `Output File Type` → `fFileKind`: `'TEXT'`→`fText`, `'EXCEL'`→`fExcel`, **else** `fBoth` (`:425-430`).
     **NULL/empty silently → BOTH** (same as `.ord` H2).
   - `Site Number in Order` (`BIT_SITE_NUMBER_IN_ORDER`) → `sendsite` boolean (`:432`).
5. **Open the new file(s)** for this supplier (`:435-463`):
   - if `fExcel`/`fBoth`: COM-open `TemplateDir+'ForecastTemplate.xls'`, sheet 1, set row cursor `i:=2`
     (`:437-447`).
   - if `fText`/`fBoth`: `Rewrite` a `.frc` at the supplier directory (`:450-453`), and if `LocalFTP`
     also `Rewrite` an Archive `.frc` (`:454-459`).

After the loop ends (EOF), the final supplier's open files are SaveAs'd/closed (`:517-544`).

### `SELECT_SupplierInfo` (the per-supplier config)

`CreateInventory.sql:5993`. Signature `(@SupCode varchar(5) = '', @Logistics bit = 0)`. With a non-empty
`@SupCode` it returns ONE row from `INV_SUPPLIER_MST s LEFT JOIN INV_LOGISTICS_MST l`. The aliases this feed
consumes:

| Alias | Source column | Used at |
|---|---|---|
| `Supplier Code` | `s.VC_SUPPLIER_CODE` | filenames (`:374`, `:450`, …) |
| `Supplier Name` | `s.VC_SUPPLIER_NAME` (varchar 25) | filenames (`ANSIReplaceStr(..,'/','')`) |
| `Directory` | `s.VC_BREAKDOWN_ORDER_DIRECTORY` (varchar 512) | output dir |
| `Output File Type` | `CASE VC_OUTPUT_FILE` `'T'`→TEXT / `'E'`→EXCEL / `'B'`→BOTH (no ELSE) | channel select |
| `Site Number in Order` | `BIT_SITE_NUMBER_IN_ORDER` (bit) | `sendsite` flag |

---

## 3. The EXACT `.frc` text format

Built at `:484-508`, written `Writeln(tcf, tcl)` (`:508`) — i.e. **CRLF after every line, including the
last** (trailing CRLF). It is **positional / delimiter-less** (concatenation), exactly like the `.ord`.

Field order, in byte order:

| # | Field | Source | Format | Width |
|---|---|---|---|---|
| 0 | **SiteSupplierCode** | `fieldbyname('SiteSupplierCode')` | as-is | **ONLY when `sendsite`=TRUE** — phantom column, see §5. Single-site: ABSENT. |
| 1 | **Supplier code** | `VC_SUPPLIER_CODE` | as-is | raw `varchar(5)`, NO pad (`:490`/`:494`). When `sendsite`, appended AFTER field 0. |
| 2 | **Part number** | `VC_PART_NUMBER` | as-is | raw `varchar(12)`, NO pad (`:497`) |
| 3 | **Week date** | `VC_WEEK_DATE` | as-is | raw `varchar(8)` = `yyyymmdd` (`:498`) |
| 4 | **Week number** | `IN_WEEK_NUMBER` | `Format('%.2d', [n])` | min 2 digits, zero-pad (`:499`) |
| 5 | **Qty day 1** | `IN_QTY1` | `Format('%.5d', [n])` | min 5 digits, zero-pad (`:500`) |
| 6 | **Qty day 2** | `IN_QTY2` | `Format('%.5d', [n])` | (`:501`) |
| 7 | **Qty day 3** | `IN_QTY3` | `Format('%.5d', [n])` | (`:502`) |
| 8 | **Qty day 4** | `IN_QTY4` | `Format('%.5d', [n])` | (`:503`) |
| 9 | **Qty day 5** | `IN_QTY5` | `Format('%.5d', [n])` | (`:504`) |
| 10 | **Qty day 6** | `IN_QTY6` | `Format('%.5d', [n])` | (`:505`) |
| 11 | **Qty day 7** | `IN_QTY7` | `Format('%.5d', [n])` | (`:506`) |

A normal (single-site, non-sendsite) line is therefore exactly:

```
<VC_SUPPLIER_CODE><VC_PART_NUMBER><VC_WEEK_DATE><WW><QQQQQ><QQQQQ><QQQQQ><QQQQQ><QQQQQ><QQQQQ><QQQQQ>\r\n
```

**Format-string semantics (CRITICAL — reuse `order_file/_format_qty`):**
- `Format('%.2d', [n])` and `Format('%.5d', [n])` are **PRECISION (minimum DIGIT count)**, NOT field width.
  The sign is prepended OUTSIDE the zero-pad. So `Format('%.5d', [-5]) = '-00005'` (6 chars), whereas Python
  `'%05d' % -5 = '-0005'` (5 chars) — **DIVERGENT for negatives**. Use the existing
  `order_file.code._format_qty` (sign + abs zero-padded to min 5) for the qty fields and the analogous
  min-2 logic for the week number. Forecast qtys are normally non-negative, but reproduce the Pascal shape
  to be safe.
- A qty **> 99999** is NOT truncated — `%.5d` prints all digits and **widens the line** (shifts nothing
  before it, since qty is the trailing block, but a positional parser reading fixed 5-char qty windows would
  mis-read). See Hazard H-OVF (§7). Likely benign in practice (day-buckets rarely ≥100k) but confirm vs golden.
- `IN_QTYn` are `int` NULL; `AsInteger` on a NULL field returns 0 → `%.5d` → `00000`. Safe.

### Filename + directory + archive

- **Filename (`:450`):**
  `<Directory>\<SupplierName with '/' removed>-<SupplierCode>.frc`
  i.e. `ANSIReplaceStr(SupplierName,'/','') + '-' + SupplierCode + '.frc'`.
  **NOTE the cleaning differs from `.ord`**: this feed strips ONLY `'/'` (`:374`/`:450`). The `.ord` emit
  strips `','`, `'.'`, AND `'\'` (`order_file._clean_name`). Reproduce the **forecast-specific cleaning**
  (`'/'` only) — do NOT reuse `_clean_name` unchanged. Flagged H-CLEAN (§7).
- **Archive filename (`:456`, only when `LocalFTP`=TRUE):**
  `<Directory>\Archive\<SupplierName w/o '/'>-<SupplierCode><yyyymmdd>.frc`
  — the date is **concatenated with no separator** before `.frc` (same no-separator quirk as `.ord` H7).
- **Directory** = `SELECT_SupplierInfo.Directory` = `VC_BREAKDOWN_ORDER_DIRECTORY` per supplier. (Single
  site: this is a configured absolute share. M4 relates these to the `INV_SITES` path columns — but in this
  single-site rebuild, treat as the per-supplier configured directory exactly as the `.ord` emit does.)
- The `.frc` is written line-by-line during the loop and `taf` (archive) is a second `Writeln` of the same
  `tcl` (`:509-512`). **Not transactional** — same H1 hazard as `.ord`; use the temp-then-rename publish.

---

## 4. The Excel layout

Built at `:437-447` (open) and `:466-481` (write). Template `ForecastTemplate.xls` from `TemplateDir`,
worksheet 1, data starts at **row `i:=2`** (`:446`; row 1 reserved for the template header). One row per
cursor record, `i` increments (`:480`).

| Excel col (1-based) | Source | Read at |
|---|---|---|
| 1 | `VC_SIZE_CODE` | `:468` |
| 2 | `VC_PART_NUMBER` | `:469` |
| 3 | `VC_WEEK_DATE` | `:470` |
| 4 | `IN_WEEK_NUMBER` | `:471` |
| 5 | `IN_QTY1` | `:472` |
| 6 | `IN_QTY2` | `:473` |
| 7 | `IN_QTY3` | `:474` |
| 8 | `IN_QTY4` | `:475` |
| 9 | `IN_QTY5` | `:476` |
| 10 | `IN_QTY6` | `:477` |
| 11 | `IN_QTY7` | `:478` |

All written `.value := fieldbyname(...).AsString` (so the Excel cells are **strings**, NOT numbers — the
qtys/week-num are NOT `%.Nd`-formatted in the Excel path, just the raw string value). Note Excel includes
`VC_SIZE_CODE` (col 1) which the `.frc` text does NOT.

### Excel filename + archive

- **SaveAs (`:376`/`:523`):**
  `<Directory>/<SupplierName w/o '/'>-<SupplierCode>-Forecast`
  (literal suffix `-Forecast`, **no `.xls` extension in the SaveAs string** — Excel/COM appends it; the
  rebuild should emit `<...>-Forecast.xlsx`). Pre-deletes an existing file first (`:374-375`/`:521-522`).
- **Archive (`:527-530`, only when `LocalFTP`=TRUE):**
  `<Directory>\Archive\<SupplierName w/o '/'>-<SupplierCode>-Forecast` (SaveAs again).

### Template dependency — POI builds it fresh

`ForecastTemplate.xls` only supplies the **header row** styling (row 1; the data starts at row 2). The data
itself is written cell-by-cell. **The rebuild does NOT need the `.xls` template** — use the M3 POI
declarative lane: a header band (row 1: the 11 column titles — recover exact header text from a sample
`ForecastTemplate.xls` or a golden output if header fidelity matters; otherwise label them
Size/Part/WeekDate/WeekNo/Qty1..7) + a column band (row 2+: the 11 columns above). This is exactly the
`report_render` declarative model (`report_render/code.py` header-band + column-band).

---

## 5. The `SiteSupplierCode` mod + the P6 crash + single-site fix

**The mechanism (proven):**

At `:486-490`:
```pascal
if sendsite then                              // Mod to add site supplier
begin
  tcl:=fieldbyname('SiteSupplierCode').AsString;   // <-- crashes
  //tcl:=Data_Module.fiSupplierCode.AsString;       // (old line, commented out)
  tcl:=tcl+fieldbyname('VC_SUPPLIER_CODE').AsString;
end
```
`sendsite` = `SELECT_SupplierInfo.'Site Number in Order'` = `BIT_SITE_NUMBER_IN_ORDER` (`:432`).

`SELECT_ForecastSupplier` is `SELECT * FROM INV_BREAKDOWN_FC_INF`. The table
(`CreateInventory.sql:578-590`) has exactly 12 columns: `IN_WEEK_NUMBER, VC_WEEK_DATE, VC_SUPPLIER_CODE,
VC_PART_NUMBER, VC_SIZE_CODE, IN_QTY1..7`. **There is NO `SiteSupplierCode` column.** A grep of the entire
live dump for `SiteSupplierCode` / `INV_SITES` / `VC_SITE_SUPPLIER` returns **zero hits** — the column
exists in no table, and there is no `INV_SITES` table in this snapshot. So when `sendsite=TRUE`,
`fieldbyname('SiteSupplierCode')` raises *"field not found"*, caught by the outer try/except at `:550-563`
→ `LogActLog('ERROR','Unable to create forecast output files…')`, `result:=FALSE`, the whole run aborts for
that supplier. This is the **same latent crash as the `.ord` emit's H4** (`order_file/code.py:31-37`).

**The mod was never finished.** The old working line (`tcl := fiSupplierCode`) was commented out and
replaced with the phantom-column read; `INSERTUPDATE_BreakdownForecastInfo` (`CreateInventory.sql:1215`)
INSERTs only the 12 real columns — it never writes a `SiteSupplierCode` — so nothing populates it.

**Live exposure:** on `Inventory_Live`, **11 of 16 suppliers have `BIT_SITE_NUMBER_IN_ORDER = 1`**
(`order-file-data-analysis.md:382-407`). So the broken branch WOULD be taken for most suppliers — which is
strong evidence the feed is currently run only in a non-sendsite-effective way, OR the operator tolerates
the per-supplier error. (Data-dependent — confirm with David whether forecast `.frc` files are actually
produced for the sendsite suppliers today.)

**Single-site fix (apply the `.ord` H4 adjudication verbatim):**
- **BUILD the working non-sendsite line for ALL suppliers** — emit field 1 (`VC_SUPPLIER_CODE`) and the rest,
  with NO leading SiteSupplierCode.
- **Do NOT reproduce the crash** and do NOT invent a leading field.
- The real multi-site leading field is an **M4** concern: the `.ord` rebuild defined it as
  `INV_SITES.VC_SUPPLIER_CODE` (the per-site supplier code), width/justification **pending a golden file**
  (`order_file/code.py:111-114`, `_M4` note). The forecast feed's leading field, when M4 lands, mirrors that
  — but per memory `project-multisite`, the deployment reversed to single-site, so this branch stays
  unbuilt. Cross-reference: the `.ord` `BIT_SITE_NUMBER_IN_ORDER` path and this `.frc` `sendsite` path are
  the **same flag from the same column** and should resolve identically.

---

## 6. The trigger + directory + usage markers

### Trigger
`TForecastBreakdown_Form.Execute` (`:149`) is the entry. It does TWO things in one call: (a) IMPORTS the
inbound 830/legacy file named by the `filename` property (parse → `DeleteBreakdown` → `UpdateForecast` →
`UpdateUsage`, the M2 path), THEN (b) runs the OUTBOUND supplier emit (this spec, `:340-568`). The emit only
runs if `ScanPartnumber` returned TRUE (`:320`).

Live callers (set `.filename` + `.SupplierCode` then `.Execute`):
- **`EDIUpload.pas:94-99`** (LIVE, `dpr:53`) — automated EDI inbound dir scan:
  `filename := fiEDIIn + '\' + sr.Name`, then `.Show; .Execute; .Free`. **This is the operational sender.**
- **`UploadBreakDown.pas:184-186`** (LIVE, `dpr:7`) — manual upload dialog (`ForecastFilleNameDialog`).
- `ForecastUploadBreakDown.pas:100-102` — **NOT in dpr → DEAD CODE**, ignore.
- `ManualForecast.pas:128-130` — the whole block is **commented out**, ignore.

So the feed fires whenever an 830 forecast file is imported (either auto-scanned by `EDIUpload` or
hand-uploaded). It is **not a standalone "send forecasts" button** — emit is a side effect of import. The
form's only UI is a `THistory` log + an OK button that the operator clicks to dismiss after
`fclosed` (`:582-587`, `OKButtonClick:1482`).

> **Discriminator for the rebuild (multi-path rule):** the BEHAVIOR a rebuild reproduces is "after a
> successful 830 import, regenerate every affected supplier's `.frc`/Excel." Anchor the oracle on the
> `EDIUpload` operational path (the auto-scan), not the manual dialog — they call the identical `Execute`,
> so the emit is the same, but the operational trigger is the inbound-dir scan.

### Directory
Per-supplier `SELECT_SupplierInfo.Directory` (`VC_BREAKDOWN_ORDER_DIRECTORY`). Single-site: the configured
absolute share per supplier (M4 would relate to `INV_SITES` path columns; not built single-site). `LocalFTP`
(`fiLocalFTP`, INI `[INIT] LocalFTP`) gates the `\Archive\` copy.

### LogActLog markers (usage signal)
`LogActLog('FORECAST', …)` fires throughout. The emit-stage markers (useful to detect real usage / volume):
- `'Create text file :<path>'` (`:452`), `'Create archive text file :<path>'` (`:457`)
- `'Create text file for supplier, <code>'` (`:462`), `'Create excel file for supplier, <code>'` (`:444`)
- `'Forecast processing complete'` (`:566`)

Plus the import-stage markers (`'EDI Count='`, `'… total records to process'`, etc.). **Ask David for a real
production Activity-log export filtered to `VC_TYPE='FORECAST'` to (a) confirm the feed is actively run and
(b) see whether the `'Create text file for supplier'` markers fire for the sendsite suppliers** (which would
contradict the P6-crash reading) — the spike's synthetic log can't answer this.

---

## 7. Hazards (first-class findings)

- **H-SITE (the P6 crash / unfinished mod).** `:488` reads the phantom `SiteSupplierCode` column (absent
  from `INV_BREAKDOWN_FC_INF` and every table). When `BIT_SITE_NUMBER_IN_ORDER=1` (11/16 live suppliers) the
  run aborts for that supplier. **Fix: build non-sendsite for all suppliers (§5).** Same as `.ord` H4.
- **H1 (not transactional).** `Rewrite`/`Writeln`/Excel `SaveAs` write to disk immediately; the import-stage
  DB writes happen earlier and aren't rolled back with the files. A failure mid-run can leave a partial
  `.frc`. **Fix: temp-then-rename publish (reuse `order_file` step 4).**
- **H2 (NULL output type → BOTH).** `SELECT_SupplierInfo`'s CASE has no ELSE; NULL `VC_OUTPUT_FILE` →
  `Output File Type` NULL → Pascal `else` → `fBoth` (`:429-430`). A supplier with unset output type silently
  gets both. Reuse `order_file._file_kind` (already reproduces this).
- **H-OVF (qty > 99999 widens the line).** `%.5d` is min-width, not truncating; a day-bucket ≥ 100000 prints
  6+ digits and shifts a fixed-position parser's qty windows. Confirm against a golden `.frc` whether qtys
  ever reach 5+ digits and whether the receiving parser expects fixed 5-char windows.
- **H-CLEAN (filename cleaning differs from `.ord`).** This feed strips ONLY `'/'` from the supplier name
  (`:374`/`:450`); the `.ord` emit strips `','`,`'.'`,`'\'`. Do NOT reuse `order_file._clean_name` unchanged
  — replicate the `'/'`-only cleaning. A supplier name with a comma/period/backslash would land in a
  different filename than `.ord` for the same supplier.
- **H-ARCH-SEP (no separator before archive date).** Archive `.frc` name concatenates `<code>yyyymmdd.frc`
  with no `-` (`:456`) — same no-separator quirk as `.ord` H7. Two runs the same day overwrite the archive
  (date has day resolution only). Reproduce faithfully unless the golden shows otherwise.
- **H-NEG (`%.Nd` sign placement).** Negative qty (shouldn't occur in a forecast, but) formats as
  `-00005`, not `-0005` — use `order_file._format_qty`, not Python `'%05d'`.
- **Excel COM dependency — MOOT in Ignition.** The legacy `CreateOleObject('Excel.Application')` is replaced
  by the M3 POI render (no Excel install, no template `.xls` needed).
- **Empty supplier set.** `SELECT_ForecastSupplier` can return 0 rows (no future-dated buckets) → the
  `while not eof` body never runs, no files emitted, `excel` stays `Unassigned`, `fopen` stays FALSE → the
  EOF cleanup (`:517-544`) no-ops. Benign. The rebuild should likewise emit nothing (return a 0-file
  summary), not error.
- **Single-site assumption.** The per-supplier `Directory` and the (broken) site prefix are the only
  site-aware knobs. Single-site rebuild: treat `Directory` as the configured per-supplier share; the site
  prefix is deferred to M4 (and is moot under the single-site deployment per `project-multisite`).

---

## 8. Items to confirm against the golden / David (data-dependent)

1. **Is the feed actually producing `.frc` files for the sendsite (`BIT_SITE_NUMBER_IN_ORDER=1`)
   suppliers today?** If yes, the P6-crash reading is wrong (a `SiteSupplierCode` column would have to exist
   in the deployed DB beyond this dump) — adjudicate before shipping. Check via a real `VC_TYPE='FORECAST'`
   Activity-log export for `'Create text file for supplier, <sendsite-code>'` markers.
2. **Golden `.frc`** — to confirm field widths/order, whether the qty ever exceeds 5 digits, and whether the
   receiving sub-supplier parser expects always-full-width positional columns (the unpadded supplier/part
   fields, H3-style, would shift a fixed-position parser).
3. **`ForecastTemplate.xls` header row** — recover the exact row-1 column titles if header fidelity matters;
   otherwise POI can emit reasonable labels.
4. **Exact week-number value** (cross-check with M2): for a TEMA `2624` line, `IN_WEEK_NUMBER` must serialize
   as `24` (the `%.2d` of the stored week number), consistent with the M2 import's
   `INV_BREAKDOWN_FC_INF.IN_WEEK_NUMBER = 24` (forecast-import-spec line 93).

---

## 9. Ignition build mapping (no separate architect pass needed)

- **`.frc` text + driver + atomicity + supplier loop + filename + `SELECT_SupplierInfo` + `_coerce_bit`** →
  clone `docs/analysis/order/project-library/order_file/code.py`. Swap the feed SQL to
  `EXEC SELECT_ForecastSupplier @WeekDate=?` (today `yyyymmdd`), the break key stays `VC_SUPPLIER_CODE`, the
  line formatter becomes the §3 field list (use `_format_qty` for the 7 qtys + a min-2 for the week number),
  the filename uses `'/'`-only cleaning + `.frc`, and the destinations are supplier dir (+ `\Archive\` on
  `LocalFTP`).
- **Excel** → M3 `report_render` declarative lane: header band (row 1) + 11-column column band (row 2+) per
  §4. No template file.
- **SiteSupplierCode** → non-sendsite for all suppliers, M4-deferred multi-site field (§5).
- **Trigger** → fire the emit as the post-import step of the 830 forecast-import service (the inbound EDI
  scan), matching `EDIUpload.pas` — not a standalone button.

The only design fork is the §8 golden/sendsite confirmation; everything else is a faithful clone of two
already-built seams.
