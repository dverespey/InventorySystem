# Legacy Behavioral Spec — Order **File Generation** (the OUTPUT stage)

Source of truth for the order-file emitter. Status: **LIVE** —
`InventorySystem.dpr:33` (`OrderFormCreateF in 'OrderFormCreateF.pas'
{OrderFormCreate_Form}`). Confirmed compiled; not dead code.

This is the **output stage** of the Order domain, downstream of:
- `Order.pas` worksheet build + commit (`legacy-order-spec.md`) which writes
  open-order rows into `INV_OPEN_ORDER_INF`, and
- the renban grouping form (RenbanOrder) which stamps the renban number onto
  blank-renban (palletized/grouped) orders.

It reads back the **not-yet-ordered** open orders and emits supplier order files
(an Excel order-form workbook and/or a machine-readable `.ord` text file) to up
to three destinations, then stamps each order with its order date so it is not
re-emitted.

Boundary note (do not re-spec): the worksheet, the FRS/renban assignment, lot-
sized vs palletized splitting, and the `INSERT_OpenOrder` server logic are all
covered in `legacy-order-spec.md` §6 and built in
`docs/analysis/order/project-library/order/code.py`. **None of that is in scope
here.** This stage is **entirely new build scope** — the project-library `order`
service is the commit path, not the emitter.

Confidence: **HIGH** on Pascal flow and on all 5 proc bodies (read below from the
authoritative live dump `DB Schema/CreateInventory.sql`, UTF-16LE). The Excel
**cell coordinates** below are HIGH (read directly from the Pascal `Cells[r,c]`
writes); the visual layout/labels of the three `.xls` templates
(`OrderTemplate.xls`, `OrderSheetTemplateWheel.xls`, `OrderSheetTemplateTire.xls`)
are **body unverified** — those template files were not parsed in this pass (the
prior `source-artifacts.md` extraction covered `OrderSimulation.xls`, a different
template for a different stage). See Gaps §10.

---

## 1. Overview & control flow

Entry: `MainMenu.pas:276-279` — creates the form, sets `FileKind:=fText`
(immediately overwritten per-supplier, see §2), and `Execute` → `ShowModal` →
**all work happens in `FormActivate`** (`OrderFormCreateF.pas:55-707`). The form
itself (`OrderFormCreate_Form.dfm`) is a pure progress shell: a `THistory` log
pane + an `OK` button hidden until completion. No user input.

Gated upstream by `[INIT] ConfirmOrderFileCreation` (`MainMenu.pas:269-274`) — a
"Would you like to create the order files now?" confirmation.

The whole run is **one transaction** on `Inv_Connection`
(`BeginTrans` :82 … `CommitTrans` :685; `RollbackTrans` on any exception :692).
Excel/`.ord` file writes happen **inside** that transaction window — see Hazard H1
(files are NOT rolled back; only the DB order-date stamps are).

Algorithm (single forward cursor over `SELECT_OrderNotOrdered`):

```
open SELECT_OrderNotOrdered           -- all un-ordered orders, sorted by supplier, renban
if recordcount = 0 -> log "No orders to process", done
while not eof:
  if VC_SUPPLIER_CODE != lastsupplier:        -- SUPPLIER BREAK (:101)
      close out previous supplier's open files (Excel order-sheet, Excel order, .ord)
      lastsupplier := this supplier
      resolve per-part logistics dir   (SELECT_PartsStockLogistics)   (:179)
      resolve supplier config          (SELECT_SupplierInfo)          (:197)
      open this supplier's files (order-sheet wkbk / Excel wkbk / .ord file(s))
  if renban changed AND order-sheet = 'WHEEL':  -- RENBAN BREAK (:355)  WHEEL ONLY
      close + reopen a fresh order-sheet workbook (one wkbk per renban)
  compute ship-date offset (SELECT_PartShipDays + GetShip weekend/holiday walk)
  if order-sheet is full (o>23):                -- PAGE BREAK (:469)
      save + reopen another order-sheet page
  write this part's line into: order-sheet (if open), Excel order (if EXCEL/BOTH),
                               .ord (if TEXT/BOTH)
  UPDATE_ORDEROrderDate  (stamp this order so it won't re-emit)        (:589)
  next
close out the final supplier's files                                  (:609-676)
commit
```

`lastsupplier` starts `''` (:71) so the very first row always triggers the
supplier-break setup; the close-previous blocks are guarded by `VarIsEmpty`/`fopen`
so they no-op on the first iteration.

---

## 2. The four file-output channels (per supplier)

`SELECT_SupplierInfo` (§5) returns per supplier:
- **`Output File Type`** (`VC_OUTPUT_FILE`: `'T'→TEXT`, `'E'→EXCEL`, `'B'→BOTH`)
  → drives `fFileKind` (`fText`/`fExcel`/`fBoth`, `OrderFormCreateF.pas:208-213`).
  **NOTE:** if `VC_OUTPUT_FILE` is anything else (NULL, `''`, `'B'`) the alias
  computes NULL/`''` and the final `else` makes it `fBoth` (:212-213). So a NULL
  output-file column **defaults to BOTH** (Excel + text). Hazard H2.
- **`Create Order Sheet`** (`VC_CREATE_ORDER_SHEET`: `''`/`'WHEEL'`/`'TIRE'`/…) →
  whether to ALSO produce the human-readable Excel **order sheet** (`excelOrder`),
  in addition to the Excel order/`.ord` machine files.

So a single supplier can emit up to **three artifact families** simultaneously:

| Channel | Var | When | Template | Filename stem |
|---|---|---|---|---|
| **Excel order sheet** (the formatted PO the supplier reads) | `excelOrder` | `Create Order Sheet` ≠ `''`/`' '` | `OrderSheetTemplateWheel.xls` if `'WHEEL'` else `OrderSheetTemplateTire.xls` | `OS<name>-<sup>-<renban>` |
| **Excel order** (column dump) | `excel` | `fFileKind` ∈ {fExcel, fBoth} | `OrderTemplate.xls` | `<name>-<sup>-<timestamp>` or `<name>-<sup>` |
| **`.ord` text** (machine-readable) | `tcf`/`tlf`/`taf` | `fFileKind` ∈ {fText, fBoth} | none (raw text) | `<name>-<sup><timestamp>.ord` or `<name>-<sup>.ord` |

`<name>` = `lastsuppliername` = `VC_SUPPLIER_NAME` with `,` and `.` stripped
(:227) then `\` stripped at filename time (`ANSIReplaceStr(...,'\','')`). `<sup>` =
`VC_SUPPLIER_CODE`. The order-sheet stem is prefixed `OS`.

---

## 3. The `.ord` text format (field-by-field) — THE machine-readable order

Built at `OrderFormCreateF.pas:556-585`, one **fixed-position concatenated line
per order row**, written with `Writeln` (so each line is terminated CRLF on
Windows). There are **no delimiters** — fields are positional, widths come from
the source column widths and the two `format()` pads.

Line = concatenation, in order:

| # | Field | Source (column / expr) | Width / format | Notes |
|---|---|---|---|---|
| 1a | **SiteSupplierCode** | `fieldbyname('SiteSupplierCode')` | native (varies) | **ONLY when `sendsite`=TRUE.** See H4 — this column does NOT exist → runtime error. (:564) |
| 1b | **Supplier code** | `VC_SUPPLIER_CODE` | as-is, `varchar(5)` | always present; when `sendsite` it is appended AFTER SiteSupplierCode (:565) |
| 2 | **FRS number** | `VC_FRS_NUMBER` | as-is, `varchar(7)` | the FRS number assigned at commit (:570) |
| 3 | **Renban number** | `VC_RENBAN_NUMBER` | `format('%8s', …)` → **right-justified to width 8** (space-padded on the LEFT) | (:571) `varchar(8)`; never blank here (proc filters `<> ''`) |
| 4 | **Part number** | `VC_PART_NUMBER` | as-is, `varchar(12)` | (:572) |
| 5 | **Quantity** | `IN_QTY` | `format('%.5d', …)` → **5-digit zero-padded integer** | (:573) qty >99999 overflows the width (no truncation, just wider) |
| 6 | **Ship date** | `now + ship` | `FormatDateTime('yyyymmdd', …)` → 8 chars | (:574) the computed ship date (§6), NOT the order date |

So a normal (non-sendsite) line is:
`<sup:5><frs:7><renban:right-just-8><part:12><qty:%.5d:5+><shipdate:yyyymmdd:8>`
e.g. supplier `ABCDE`, FRS `1234567`, renban `H006` (→ `"    H006"`), part
`90210ABCDEF1`, qty 240, ship 2026-06-25:
`ABCDE1234567    H00690210ABCDEF1002402026 0625` → exact bytes:
`ABCDE1234567` + `    H006` + `90210ABCDEF1` + `00240` + `20260625`.

**Width hazards:** only fields 3 and 5 are explicitly padded. Fields 1b/2/4/6 rely
on the underlying columns being exactly their declared widths. `VC_SUPPLIER_CODE`
is `varchar(5)` but **not** space-padded by the code — a short supplier code
shifts every following field left (positional parser at the sub-supplier breaks).
Hazard H3. Confirm against a golden `.ord`: whether the receiving sub-supplier
parser expects fixed 5-char supplier / 7-char FRS / 12-char part, i.e. whether
those columns are always full-width in practice.

The same `tcl` line is written to **all open text destinations** in lockstep
(:576-582): supplier file (`tcf`) always; logistics (`tlf`) and archive (`taf`)
only when `fiLocalFTP` AND (for `tlf`) logistics ≠ `'NONE'` — see §4.

---

## 4. The three destinations + skip-logistics (Q8) + the `fiLocalFTP` mode switch

There are **TWO emit modes**, switched by `Data_Module.fiLocalFTP.AsBoolean`
(INI `[INIT] LocalFTP`, **default `False`** — `DataModule.dfm:294-300`):

- **`LocalFTP = False` (default):** write **ONE copy**, to the **supplier
  directory only** (`lastdirectory`). No logistics copy, no archive copy. This is
  the plain "drop the file where the FTP/dispatch process picks it up" mode.
- **`LocalFTP = True`:** write **three copies** — supplier dir, logistics dir,
  and `<supplier dir>\Archive`. The comment calls this "FTP writes auto archive
  and logistics file" (:297-298). So `LocalFTP` truly means "this box also owns
  the archive + logistics fan-out locally" rather than "use FTP".

The three destinations:

| # | Destination | Source | Skip condition |
|---|---|---|---|
| (a) | **Supplier dir** | `lastdirectory` = `SELECT_SupplierInfo.Directory` = `VC_BREAKDOWN_ORDER_DIRECTORY` | never skipped (always written) |
| (b) | **Logistics dir** | `lastlogisticsdirectory` (resolved §5) | written only if `LocalFTP=True` **AND** `lastlogisticsdirectory <> 'NONE'` |
| (c) | **Archive** | `lastdirectory + '\Archive\'` | written only if `LocalFTP=True` |

**Q8 "no logistics, supplier delivers on their own" = the `'NONE'` sentinel.**
When the resolved logistics directory is the literal string `'NONE'`, the
logistics copy (b) is skipped at every save/write/close site (the `if
lastlogisticsdirectory <> 'NONE'` guards at :111, :135, :147, :164, :301, :332,
:365, :481, :579, :617, :640, :652, :670). The supplier dir (a) and archive (c)
are still written (when `LocalFTP=True`). So `'NONE'` = "this supplier has no
3PL/logistics drop; they collect directly." Distinguish from `''` and NULL — see
§5 for how the sentinel is set.

### 4.1 Filename patterns (exact)

`fts` = `formatdatetime('yyyymmddhhmmss00', now)` (14 digits + literal `00` =
16 chars). `<name>` = `ANSIReplaceStr(lastsuppliername,'\','')`.

**Excel order sheet (`excelOrder`)** — `SaveAs` writes `.xls`:
- supplier: `lastdirectory\OS<name>-<sup>-<renban>` (:107)
- logistics: `lastlogisticsdirectory\OS<name>-<sup>-<renban>` (:112) [LocalFTP & ≠NONE]
- archive: `lastdirectory\Archive\OS<name>-<sup>-<renban>` (:114) [LocalFTP]
- on a **page break** (`o>23`) the supplier+archive names get `+IntToStr(page)`
  appended (`…-<renban><page>`, :477/:484), **but the logistics name does NOT get
  the page suffix** (:482) — Hazard H6: multi-page wheel sheets collide their
  logistics copies (each page overwrites the prior at the same logistics path).

**Excel order (`excel`)** — `SaveAs`:
- if `lasttimestamp` (`BIT_ORDER_FILE_TIMESTAMP`): `…\<name>-<sup>-<fts>` (:131)
  (note: the supplier-break close uses `<name>-<sup>-<fts>` WITH a `-` before the
  timestamp at :636/:131, while the in-loop `.ord` name has **no** `-` before
  the timestamp — see H7 inconsistency)
- if NOT `lasttimestamp`: `…\<name>-<sup>` (:143) — **no timestamp → overwrites
  on every run** (intended: stable filename per supplier). BUT the archive copy
  of a non-timestamped Excel order is FORCED to carry a timestamp
  (`…-<fts>`, :150) so the archive keeps history.

**`.ord` text** — `AssignFile`+`Rewrite`:
- if `lasttimestamp`: `lastdirectory\<name>-<sup><fts>.ord` (:290) — **no `-`
  between `<sup>` and the timestamp** (concatenated)
- if NOT `lasttimestamp`: `lastdirectory\<name>-<sup>.ord` (:321)
- logistics `.ord`: same stem in `lastlogisticsdirectory` (:303 / :334)
- archive `.ord`: `lastdirectory\Archive\<name>-<sup><fts>.ord` — **archive
  always timestamped** even in the no-timestamp branch (:310 / :341), so archives
  never overwrite.

> Each `SaveAs` with no extension lets Excel append the workbook default
> (`.xls`). The `.ord` paths carry an explicit `.ord` extension.

### 4.2 Where each artifact goes (summary)

Both the Excel order-sheet AND the Excel order AND the `.ord` go to the **same
set** of destinations (supplier always; logistics+archive only when `LocalFTP`).
They are NOT routed to different destinations — every produced artifact follows the
same a/b/c fan-out. Which **artifacts** are produced is per-supplier config
(`Output File Type`, `Create Order Sheet`); which **destinations** receive them is
the `LocalFTP` + `NONE` logic above.

---

## 5. `SELECT_PartsStockLogistics` + logistics-directory resolution (per supplier)

The logistics directory is resolved **once per supplier** (at the supplier break),
from TWO sources with a precedence + sentinel rule
(`OrderFormCreateF.pas:179-224`):

**Step 1 — per-PART logistics** (`SELECT_PartsStockLogistics;1`, schema:6093),
keyed on the supplier's FIRST part in the cursor (`VC_PART_NUMBER`):
```sql
SELECT l.VC_BREAKDOWN_ORDER_DIRECTORY 'LogisticsDirectory'
FROM INV_PARTS_STOCK_MST p JOIN INV_LOGISTICS_MST l ON p.IN_LOGISTICS_ID = l.IN_LOGISTICS_ID
WHERE p.VC_Part_Number = @PartNo
```
- It is an **INNER JOIN** on `IN_LOGISTICS_ID`. If the part has no logistics link
  (`IN_LOGISTICS_ID` NULL / no matching logistics row) → **0 rows** →
  `lastlogisticsdirectory := ''` (:191). Else → the part's logistics directory.
- **The part's logistics dir is taken from whichever part happens to be FIRST in
  the cursor for that supplier** — it is NOT per-line. All of a supplier's order
  lines share one logistics directory (the first part's). Hazard H5: a supplier
  shipping parts with mixed logistics IDs uses only the first part's dir.

**Step 2 — supplier-level fallback** (`SELECT_SupplierInfo;1`, schema:5993),
keyed on `@SupCode=lastsupplier`, `@Logistics=1`. Returns the supplier's own
`Directory` (`VC_BREAKDOWN_ORDER_DIRECTORY` from `INV_SUPPLIER_MST`) and a
`LogisticsDirectory` (`l.VC_BREAKDOWN_ORDER_DIRECTORY` from the supplier's
`INV_LOGISTICS_MST` via `s.IN_LOGISTICS_ID` LEFT JOIN).
```
lastdirectory := Directory                       -- supplier's own order dir (dest a)
if lastlogisticsdirectory = '' then              -- part-level gave nothing
   if not LogisticsDirectory.IsNull
      lastlogisticsdirectory := LogisticsDirectory     -- supplier-level logistics dir
   else
      lastlogisticsdirectory := 'NONE'           -- the Q8 skip sentinel
```

**Resolution precedence + the three sentinel states:**
1. Part has a logistics link → use the **part's** logistics dir (non-empty).
2. Else supplier has a logistics link (`s.IN_LOGISTICS_ID` → a row, dir not NULL)
   → use the **supplier's** logistics dir.
3. Else (no part link AND supplier's logistics dir is NULL) → `'NONE'` → **skip
   the logistics copy** (Q8 case).

> Subtle: if the supplier-level `LogisticsDirectory` is present but an **empty
> string** `''` (not NULL), step-2 sets `lastlogisticsdirectory := ''` (not
> `'NONE'`) — and then `'' <> 'NONE'` is TRUE, so the code would try to write to
> an **empty directory path** (`'\OS…'`, a root-relative path). Hazard H8 — the
> `'NONE'` sentinel only protects against NULL, not against an empty-string
> logistics dir. Confirm against data: whether any `INV_LOGISTICS_MST` row has an
> empty `VC_BREAKDOWN_ORDER_DIRECTORY`.

---

## 6. Ship-date computation (per line)

Each order line stamps a **ship date** = `now + ship` where `ship` is a
business-day offset (`OrderFormCreateF.pas:407-464`, `GetShip` :709-770):

1. `SELECT_PartShipDays;1` (schema:4783) returns the part's ship-day lead values.
   It first looks up the part's `IN_RENBAN_ID`; if the part is in a renban group
   it returns the **group's** ship days (`INV_RENBAN_GROUP_MST`), else the part's
   own (`INV_PARTS_STOCK_MST`). Columns: `Ship` (base `IN_SHIP_DAYS`) plus per-
   weekday overrides `ShipM/ShipT/ShipW/ShipTh/ShipF/ShipS`.
2. By **today's weekday** (`DayOfTheWeek(now)`, 1=Mon…7=Sun): if that weekday's
   override column is non-zero use it (`GetShip(ShipX)`), else fall back to the
   base `Ship`. Sunday/other → base `Ship` (:461-462).
3. `GetShip(lead)` (:709) converts the **business-day lead** into a **calendar-day
   offset** by walking forward from tomorrow, skipping weekends
   (`DayOfTheWeek < 6`) and holidays (status `'H'` from the ALC special-date
   calendar), counting only valid days until `lead` valid days have elapsed; the
   calendar offset `x` is returned.

`AD_GetSpecialDate` runs on **`ALC_StoredProc` / `ALC_Connection`** (the
`TireOrder` DB) — **NOT in `CreateInventory.sql`** (cross-DB; same proc as
`Order.pas`). Body unverified here; see `legacy-order-spec.md` Hazard 1.

The ship date feeds: the `.ord` field 6, the Excel order col 5, the order-sheet
cell `[8,6]`, and `UPDATE_ORDEROrderDate.@ShipDate`. The order date itself is
`FormatDateTime('yyyymmdd', now)` (today).

---

## 7. The Excel artifacts (cell layout, exact coordinates)

> Coordinates are HIGH confidence (read from `Cells[row,col]` writes). Template
> visual layout (labels/merges/formats baked into the 3 `.xls` files) is **body
> unverified** — Gap §10.1.

### 7.1 Excel **order** (`excel`, `OrderTemplate.xls`) — the column dump

Header, written once per supplier at open (:279-280):
- `Cells[1,6]` (row 1, col F) = Supplier Name
- `Cells[2,6]` (row 2, col F) = Address

Body rows start at `i:=10` (:282), one row per order line (:545-554):

| Cell | Col | Value | Source |
|---|---|---|---|
| `[i,1]` | A | Part number | `VC_PART_NUMBER` |
| `[i,2]` | B | FRS number | `VC_FRS_NUMBER` |
| `[i,3]` | C | Renban number | `VC_RENBAN_NUMBER` |
| `[i,4]` | D | Quantity | `IN_QTY` |
| `[i,5]` | E | Ship date | `FormatDateTime('mm/dd/yyyy', now+ship)` |

`INC(i)` per line. No page/overflow handling for the Excel order (unbounded rows).

### 7.2 Excel **order sheet** (`excelOrder`, Wheel/Tire template) — the formatted PO

Header, written once per workbook (at supplier break :253-261, renban break
:390-397, page break :506-513):

| Cell | (row,col) | Value | Source |
|---|---|---|---|
| `[11,1]` | K… | Supplier Name | `SELECT_SupplierInfo.'Supplier Name'` |
| `[11,2]` | | Supplier Code | `'Supplier Code'` |
| `[11,3]` | | Renban number | `VC_RENBAN_NUMBER` (**WHEEL only** at supplier break, :255-256; always at renban/page break) |
| `[8,8]` | | FRS date `mm/dd/yyyy` | derived from `VC_FRS_NUMBER` (see below) |
| `[11,5]` | | Page number | `page` var |
| `[13,2]` | | Address | `'Address'` |
| `[14,2]` | | City, State, Zip | `'City'+', '+'State'+', '+'Zip'` |

FRS-date cell `[8,8]` (:257-258): `year := copy(yyyy,1,3) + copy(VC_FRS_NUMBER,1,1)`
(century from `now` + the FRS's leading year-digit), then displayed as
`copy(FRS,2,2)+'/'+copy(FRS,4,2)+'/'+year` (the FRS number encodes `Y MM DD` in
its first 5 chars → `MM/DD/YYY?`). Confirm against golden: this reconstructs the
order-by date from the FRS prefix.

Body rows start at `o:=16` (:265), one row per line (:522-540), with `[8,6]` =
ship date `mm/dd/yy` rewritten each line:

| Cell | Col | Value (WHEEL) | Value (TIRE/other) |
|---|---|---|---|
| `[o,1]` | A | `VC_PART_NUMBER` | same |
| `[o,2]` | B | `VC_PARTS_NAME` | same |
| `[o,4]` | D | `VC_KANBAN_NUMBER` | same |
| `[o,5]` | E | `IN_1LOTQTY` | `VC_RENBAN_NUMBER` |
| `[o,7]` | G | `IN_QTY` | `IN_QTY` |

`INC(o)` per line. **Page fill at `o>23`** (:469-517): save current sheet (with
`+page` suffix on supplier/archive names), open a fresh workbook, reset `o:=16`,
`inc(page)`. Note the page-break **condition is `(A<>'') or (A<>' ')` which is
ALWAYS TRUE** (:469-470) — so the page logic runs whenever `o>23` regardless of
whether an order sheet was requested; guarded only by the inner `VarIsEmpty`
checks. Hazard H9 (dead-ish tautology, but the `VarIsEmpty(excelOrder)` guard
saves it from acting when no order sheet is open).

### 7.3 Order-sheet workbook granularity
- **One order-sheet workbook per renban** for WHEEL (renban break :355 closes +
  reopens). TIRE order sheets are NOT re-broken on renban change (the renban-break
  `if` requires `'WHEEL'`), so a tire supplier gets one order-sheet workbook per
  **supplier** (until page-fill). Confirm: whether tire suppliers ever use the
  order sheet in practice.

---

## 8. The data feed — `SELECT_OrderNotOrdered` and the loop driver

Single driving cursor, `Inv_DataSet`, proc `SELECT_OrderNotOrdered;1`
(schema:6325):
```sql
SELECT *
FROM INV_OPEN_ORDER_INF i
  JOIN INV_PARTS_STOCK_MST p   ON i.VC_PART_NUMBER = p.VC_PART_NUMBER
  JOIN INV_SUPPLIER_MST   s    ON p.IN_SUPPLIER_ID = s.IN_SUPPLIER_ID
  LEFT OUTER JOIN INV_RENBAN_GROUP_MST r ON p.IN_RENBAN_ID = r.IN_RENBAN_ID
WHERE ((i.VC_ORDER_DATE is null) or (i.VC_ORDER_DATE = ''))
  AND i.VC_RENBAN_NUMBER <> ''
ORDER BY s.VC_SUPPLIER_CODE, i.VC_RENBAN_NUMBER
```

Key behaviors:
- **"Not ordered" = `VC_ORDER_DATE` null or `''`.** Once a line is emitted,
  `UPDATE_ORDEROrderDate` stamps `VC_ORDER_DATE` so it drops out of this set on the
  next run. **This is the idempotency / re-run guard** — re-running the emitter
  produces files only for still-unstamped orders. (Distinct from `INSERT_OpenOrder`
  dedup; this is the emit-once guard.)
- **`VC_RENBAN_NUMBER <> ''` (never-blank-renban filter).** Blank-renban orders
  (palletized/grouped parts born blank-renban at commit, per
  `project-order-renban-domain` memo) are **NOT picked up here** — they must first
  pass through the renban grouping form (RenbanOrder), which assigns the renban.
  This enforces the memo's "never blank renban" rule **at the output boundary**:
  the emitter never emits a blank-renban order. Confirm against data: that no
  un-grouped palletized order leaks into the supplier files.
- **`SELECT *` field provenance** — `fieldbyname` resolves against the joined row:
  - From `i` (`INV_OPEN_ORDER_INF`): `VC_SUPPLIER_CODE`, `VC_FRS_NUMBER`,
    `VC_RENBAN_NUMBER`, `IN_QTY`, `VC_PART_NUMBER`, `VC_KANBAN_NUMBER`.
  - From `p` (`INV_PARTS_STOCK_MST`): `VC_PARTS_NAME`, `IN_1LOTQTY`. **Note
    `VC_PART_NUMBER` and `VC_KANBAN_NUMBER` exist in BOTH `i` and `p`** — in a
    `SELECT *`, ADO sees duplicate column names; `fieldbyname` returns the FIRST
    (the `INV_OPEN_ORDER_INF` copy). Behaviorally the part numbers are equal (join
    key) but the kanban could differ between order row and master. Hazard H10:
    rely on the **order row's** kanban (`i.VC_KANBAN_NUMBER`) to match legacy.
  - `r` (`INV_RENBAN_GROUP_MST`) is LEFT-joined but **no `r.*` field is read** by
    this stage (renban-group columns unused here).
- **`SiteSupplierCode` is read at :564 but exists in NO table** in the live schema
  (verified across `INV_SUPPLIER_MST`, `INV_OPEN_ORDER_INF`, `INV_PARTS_STOCK_MST`
  and the full dump). → Hazard H4.

Loop ordering: rows are ordered `supplier, renban`. The code's break detection
relies on this exact ordering: **supplier break** on `VC_SUPPLIER_CODE` change
(outer), **renban break** on `VC_RENBAN_NUMBER` change (inner, WHEEL only). If the
proc's `ORDER BY` were ever removed/changed, the breaks would mis-fire and split a
supplier across multiple file sets. Preserve the sort.

Per-line write: each row writes to whichever channels are open, then
`UPDATE_ORDEROrderDate;1` (schema:5972) stamps `VC_ORDER_DATE`, `VC_SHIP_DATE`,
`VC_LAST_UPDATE` **WHERE part+FRS match AND `VC_ORDER_DATE <> @OrderDate`** (so a
same-day re-stamp is a no-op). Note this UPDATE keys on **part + FRS only** (not
renban) — if two renban rows share a part+FRS, the FIRST stamp marks BOTH ordered;
the second's row was already emitted in the same run so this is benign, but a
rebuild must replicate the part+FRS scope. Hazard H11.

---

## 9. The Excel/COM dependency to replace (what the rebuild must reproduce WITHOUT Excel)

Every COM touchpoint (`createOleObject('Excel.Application')`, `workbooks.open`,
`worksheets[1]`, `Cells[r,c].value`, `ActiveWorkbook.SaveAs`, `Workbooks.Close`,
`Quit`) at :240-247, :270-275, :380-384, :494-501 (opens); :104-119, :125-157,
:358-373, :474-491, :522-540, :545-554, :610-662 (writes/saves/closes).

The rebuilt generator must reproduce, **without Excel**:

1. **The `.ord` text file** — fully specified in §3. Pure string formatting + file
   write. No template needed. This is the highest-fidelity, lowest-risk artifact.
   Reproduce the exact field widths/pads (esp. `%8s` renban, `%.5d` qty,
   `yyyymmdd` ship date) and the `sendsite` prefix (after resolving H4).
2. **The Excel order** (`OrderTemplate.xls` + the 5-col body, §7.1) — regenerate
   as `.xlsx` (or whatever the supplier consumes) from a server-side writer. Layout
   is just header (Supplier Name @F1, Address @F2) + a 5-column table from row 10.
   Mirror the **856/862/forecast** Excel replacements already done in the suite
   (server-side workbook writer, no COM).
3. **The Excel order sheet** (`OrderSheetTemplateWheel.xls` /
   `OrderSheetTemplateTire.xls` + §7.2 cells) — the formatted PO. Needs the two
   **template visual layouts** recovered (Gap §10.1) since the supplier reads this
   as a document; reproduce header block (rows 8/11/13/14) + body table from row
   16 with the WHEEL/TIRE column variant, the per-renban / per-page workbook
   splitting, and page-fill at row 23.
4. **The SaveAs fan-out** — replace `SaveAs` to {supplier, logistics, archive} with
   the destination logic of §4 (the `LocalFTP` mode + `NONE` skip), writing the
   bytes once and copying to the (1 or 3) destinations.

Build-scope verdict: **the entire emitter is new.** The project-library `order`
service builds the **commit** path only. The emitter shares NO code with it. It
needs: a `.ord` formatter (pure, unit-testable like `computeOrderRecords`), an
Excel-order writer, an order-sheet writer, the per-supplier/renban/page loop
driver, and the destination fan-out. The data feed (`SELECT_OrderNotOrdered`) and
the stamp (`UPDATE_ORDEROrderDate`) become Named Queries; `SELECT_SupplierInfo` /
`SELECT_PartsStockLogistics` / `SELECT_PartShipDays` likewise.

---

## 10. Gaps (what was NOT read / must be confirmed)

1. **The 3 order templates** `OrderTemplate.xls`, `OrderSheetTemplateWheel.xls`,
   `OrderSheetTemplateTire.xls` were **not parsed** in this pass. The §7 cell
   coordinates are from Pascal; the **static labels, merges, column widths, fonts,
   and any header text** baked into those files are unknown. (`source-artifacts.md`
   parsed `OrderSimulation.xls`, the worksheet template — a different file/stage.)
   To close: locate these 3 templates in `TemplateDir`, convert with LibreOffice,
   read with openpyxl (same method as `source-artifacts.md` §5).
2. **`AD_GetSpecialDate`** (ship-date holiday calendar) lives in the ALC/`TireOrder`
   DB — body unverified (same as `legacy-order-spec.md` Hazard 1).
3. **`SiteSupplierCode`** column — confirmed ABSENT from the live `Inventory`
   schema. Whether it exists in some deployed-but-not-in-dump column, or whether
   `BIT_SITE_NUMBER_IN_ORDER` is simply never set TRUE in production, must be
   confirmed against the live DB / golden. If never TRUE, the `sendsite` branch is
   dead; if ever TRUE it is a hard runtime error (H4).
4. **Golden `.ord` file** — to confirm §3 field widths/order (esp. the unpadded
   supplier/FRS/part columns, the `sendsite` prefix shape, and whether qty ever
   exceeds 5 digits), diff one real production `.ord` against the §3 format.

---

## 11. Hazards (first-class findings)

- **H1 — Files are NOT transactional.** `BeginTrans`/`CommitTrans` (:82/:685) wrap
  the run, but Excel `SaveAs` and `.ord` `Rewrite/Writeln` write to the filesystem
  immediately. On an exception mid-run, `RollbackTrans` reverts the
  `UPDATE_ORDEROrderDate` stamps **but the already-written files remain**. Next run
  re-emits those orders (now-unstamped) → **duplicate supplier files** (with new
  timestamps) for the partially-processed batch. A rebuild should stage files and
  publish only on commit, or make file writes idempotent per order.
- **H2 — NULL `VC_OUTPUT_FILE` defaults to BOTH.** `SELECT_SupplierInfo`'s CASE has
  no ELSE; NULL → `Output File Type` NULL → the Pascal `else` → `fBoth` (:212).
  A supplier with an unset output type silently gets both Excel and `.ord`.
- **H3 — Positional `.ord` fields are unpadded except renban & qty.** Supplier
  (`varchar(5)`), FRS (`varchar(7)`), part (`varchar(12)`) are written raw; a
  short value shifts all following fields. Fixed-position parsers at the
  sub-supplier break. Confirm columns are always full-width.
- **H4 — `SiteSupplierCode` runtime error.** When `BIT_SITE_NUMBER_IN_ORDER=1`,
  `fieldbyname('SiteSupplierCode')` (:564) targets a column that **does not exist**
  in the result set → ADO raises → the whole run rolls back. Latent bug / dead
  path depending on whether the bit is ever set. (Memory pattern: "code calls
  things that don't exist.")
- **H5 — Per-supplier logistics dir from the FIRST part only.** `SELECT_PartsStock-
  Logistics` is called once per supplier with the cursor's first part; all the
  supplier's lines share that one logistics dir even if other parts have different
  `IN_LOGISTICS_ID`. (:179-186)
- **H6 — Multi-page wheel order sheet: logistics copy collides.** Page break adds
  `+page` to supplier/archive names but NOT to the logistics name (:482 vs
  :477/:484), so pages 2..n overwrite the page-1 logistics copy.
- **H7 — Timestamp-in-filename inconsistency & collision risk.** `fts` is
  `yyyymmddhhmmss00` (1-second resolution + `00`). Two runs within the same second,
  or two supplier breaks in the same second (same `now`), collide. Also the
  separator before the timestamp is inconsistent: Excel-order uses `-<fts>`
  (:131/:636) while `.ord` uses `<fts>` with no separator (:290) — a rebuild's
  filename parser must handle both.
- **H8 — Empty-string logistics dir bypasses the `NONE` guard.** §5: a non-NULL
  but `''` logistics directory becomes `''` (not `'NONE'`), passes the `<> 'NONE'`
  guard, and produces a root-relative `\OS…` / `\…ord` path. Confirm no
  `INV_LOGISTICS_MST` row has an empty directory.
- **H9 — Page-break condition is a tautology** `(A<>'') or (A<>' ')` (:469-470)
  always TRUE; the `VarIsEmpty(excelOrder)` inner guard is what actually gates it.
  Preserve the **intent** (page only when an order sheet is open and full), not the
  literal condition.
- **H10 — `SELECT *` duplicate columns** (`VC_PART_NUMBER`, `VC_KANBAN_NUMBER` in
  both `i` and `p`). `fieldbyname` resolves to the first (the order-row copy). A
  rebuild's query should alias these explicitly to avoid ambiguity and to lock the
  kanban source to `i.VC_KANBAN_NUMBER`.
- **H11 — `UPDATE_ORDEROrderDate` keys on part+FRS only**, not renban (:5979-5981).
  Two renban rows sharing part+FRS are both stamped by the first emit. Benign
  within one run (both already emitted) but a rebuild that batches/parallelizes
  must replicate this scope or it will mis-mark.
- **H12 — Domain quirks from `project-order-renban-domain` at the line level:**
  the order-sheet column 5 for WHEEL is `IN_1LOTQTY` (per-lot), for TIRE is
  `VC_RENBAN_NUMBER`; `IN_QTY` is the whole order qty in both the `.ord` and Excel.
  Lot-sized vs palletized splitting already happened at commit (each lot is its own
  `INV_OPEN_ORDER_INF` row) — the emitter just emits whatever rows exist, one
  `.ord`/Excel line per row. The "never blank renban" rule is enforced by the
  `VC_RENBAN_NUMBER <> ''` proc filter (§8), not by the emitter.

---

## 12. Multi-site notes (defer design to ignition-architect)

- All three template paths are `Data_Module.TemplateDir + '<fixed name>.xls'`
  (:245/:247/:274/:384/:499). `TemplateDir` (`DataModule.pas:708`) = the EXE dir if
  `[DIRECTORIES] UseApplicationDir` else `[DIRECTORIES] TemplateDir`. Single
  template set per install; a multi-site rebuild must pick per site.
- Destinations are absolute filesystem paths from `INV_SUPPLIER_MST` /
  `INV_LOGISTICS_MST` (`VC_BREAKDOWN_ORDER_DIRECTORY`, `varchar(512)`). These are
  Windows local/UNC paths today → in Ignition they become a configured
  output/share location per supplier/site.
- `[INIT] LocalFTP` (default False) is a **single global flag** controlling the
  archive+logistics fan-out for the whole box — no per-supplier override. A
  multi-site rebuild likely wants this per site or always-on with explicit
  per-supplier destinations.
- `BIT_SITE_NUMBER_IN_ORDER` (the `sendsite` flag) + the missing `SiteSupplierCode`
  is the only place "site identity in the wire output" appears — and it is broken
  (H4). The rebuild's multi-site emitter must define what the site-supplier prefix
  actually is.
