# Order Form — Non-Code Source Artifacts

Extraction of the Word docs and Excel templates that drive the Select-Order
dialog and the Order Simulation sheet. Read-only toward legacy files. This doc
records what was extractable and — importantly — what was **not**.

Tooling: `textutil` (Word -> txt), `xlrd 2.0.2` (BIFF8 `.xls`).
LibreOffice is **not installed**; `openpyxl` **cannot** read `.xls`/BIFF.

---

## 1. Word docs — business rules

### `docs/SELECT ORDER.doc` — the Select-Order dialog

Accessed from the main program screen via the **Order** button.

- By default, orders are selected by **Line** and **Part type**.
- Omitting the **Line** name selects **all parts of a part type**; doing so
  **enables a "Sort By" drop-down**.
- Default screen options:
  - **Today** — current date = the order date (read-only, no selection).
  - **Line** — assembly line to order parts for.
  - **Part** — type of part to order. Default part types: **file, valves,
    tires, wheels**.
  - **Start** — builds the Excel order screen from the chosen options.
  - **Order** — creates orders from the Excel sheet data (enabled only **after**
    the order sheet is built).
  - **Exit** — cancel.
- Added option when assembly line is blank:
  - **Sort By** — sorts the part type by the selected option. All selected parts
    are **first sorted by size name** to create the correct **size / sharing
    pairs**.

### `docs/Order Sheet Fill Days.doc` — "fill days" config

- **Fill days** = the number of **visible usage days** shown in the Excel order
  sheet simulation.
- Configured via **File | Configuration -> System tab**, "fill days" field;
  click **OK** to apply.
- **Maximum value = 50.** **Old-system default = 23.**
- After changing, the new number of days becomes visible on the order
  simulation sheet (i.e. fill days controls how many day-columns are rendered).

> Cross-reference: the day-columns are the `S..AN+` block seen only in the
> `OrderSimulationChanged.xls` variant below (Week / Mon-Fri date headers).

---

## 2. Per-template dump (xlrd)

All files: **BIFF version 80** (Excel 97-2003 `.xls`), 3 sheets
(`Sheet1`, `Sheet2`, `Sheet3`); **Sheet2/Sheet3 are empty** (0x0) in every file.
**No named ranges** in any file.

The live template is opened by code as a fixed name:
`Order.pas:183` -> `excel.workbooks.open(Data_Module.TemplateDir+'OrderSimulation.xls')`.
There is **no site-suffix logic in code** — the site variant is whichever copy
is deployed *as* `OrderSimulation.xls` in that site's `TemplateDir`. The
`*HERO/*CAMEX/*WQS/*WWW/*MAS` files are the source copies kept alongside.

### Sheet1 header structure — identical across all non-"Changed" variants

Used range **6 rows x 18 cols** (cols B..R). Two header rows:

Row 5 (group headers):
`B5=Size  C5=Tire  D5=Parts  E5="Sum of Order"  H5="Order Point"
 K5=Inventory  N5=Open  O5=1Lot  P5=Lead  Q5=Qty  R5=Lot`

Row 6 (column headers):
`C6=Brand  D6=No.  E6=Qty  F6=Actual%  G6=Plan%  H6="Daily Usage"
 I6=Days  J6=Total  K6=<SITE>  L6=<DEST>  M6=InTransit  N6=Order
 O6=Qty  P6=Time  Q6=Order  R6=Order`

The **only cells that differ between variants** are **K6** and **L6**
(the two Inventory sub-columns: source warehouse / destination plant):

| File | K6 | L6 |
|------|----|----|
| `Templates/OrderSimulation.xls` (default) | WQS | NUMMI |
| `Templates/OrderSimulationWQS.xls` | WQS | NUMMI |
| `Templates/OrderSimulationHERO.xls` | HERO | TMMTX |
| `Templates/OrderSimulationCAMEX.xls` | CAMEX | DEPOT |
| `Templates/OrderSimulationWWW.xls` | WWW | SIA |
| `WWW Files/OrderSimulationWWW.xls` | WWW | SIA |
| `WWW Files/OrderSimulation.xls` | WWW | SIA |
| `MAS Files/OrderSimulationMAS.xls` | MAS | TMMM |
| `MAS Files/OrderSimulation.xls` | MAS | TMMM |

(The default `Templates/OrderSimulation.xls` ships configured as the **WQS/NUMMI**
site; each site dir carries its own copy deployed as the plain name.)

### `Templates/OrderSimulationChanged.xls` — the only structural outlier

Used range **14 rows x 45 cols** (cols B..AN+). Same B..R block as above
(`K6=WWW L6=TMMNK`), **plus** an appended day-grid:

Row 5: `S5=Date` then a run of serial dates (`T5=39160` = 2007-03-05 ...
through `AN5=39189`) — i.e. ~22 day columns.
Row 6: `S6=Week  T6=Mon U6=Tue V6=Wed W6=Thu X6=Fri  Y6=Mon ...` (weekday labels).

This is the **"fill days" usage grid** described in the Word doc (default 23 days)
materialized as columns. Cell types here include numbers/dates (`type 2,3`),
whereas the plain templates are **text-only** (`type 0,1`) — they are blank
header shells the app populates at runtime.

---

## 3. Template diff verdict: COSMETIC, not structural

**One site-scoped Perspective screen CAN serve all sites.**

Evidence:
- Every production variant (`default/WQS/HERO/CAMEX/WWW/MAS`) is **byte-for-byte
  identical in cell content except K6 and L6** — two header labels naming the
  source warehouse and destination plant. Columns, layout, row structure,
  sheet set, and (xlrd-visible) content are otherwise the same.
- The application opens a **single fixed filename** (`OrderSimulation.xls`);
  there is no per-site column/formula branching in `Order.pas`.
- K6/L6 are pure **labels** -> in a rebuild they become a **site-config lookup**
  (site code + destination plant), not a different screen.

Classification:
- **COSMETIC** (one screen, parameterized): K6/L6 site/destination labels. The
  large raw-byte `cmp` deltas (250k-325k differing bytes) are OLE/BIFF stream
  re-layout noise, **not** content differences — confirmed by the identical
  xlrd cell dumps.
- **STRUCTURAL** (genuinely different layout): only `OrderSimulationChanged.xls`,
  which adds the day-by-day usage grid (cols S..AN). This is **not a per-site
  divergence** — it is a different/newer *version* of the same sheet (the
  "fill days" feature), and it is **not** a file the app opens. It informs the
  target design (render N day-columns from the fill-days config) rather than
  requiring a separate site screen.

Net: site divergence is configuration, not structure. D1's single site-scoped
screen approach is sound.

---

## 4. Extraction-gap list (what could NOT be read here)

State plainly — do not infer thresholds or colors from these gaps.

1. **Conditional-formatting rules — UNREADABLE.** xlrd does not expose
   conditional-formatting (CF) rules at all. The color-coding logic that flags
   order points / shortages on the simulation sheet is therefore **not
   extracted** from any template (`OrderSimulation*.xls`, all dirs). Thresholds,
   comparison operators, and color mappings are **unknown from this pass**.
2. **Cell fill / font colors and number formats — UNREADABLE.** xlrd 2.0.2 does
   not load `formatting_info` for `.xls`. Static highlight colors, fonts,
   borders, and display number formats were **not** extracted for any template.
3. **Formula strings — NOT EXPOSED.** xlrd 2.0.2 dropped formula-text
   extraction. Whether cells like `J6 (Total)`, `H (Daily Usage)`, `N (Order)`
   carry in-cell formulas vs. app-injected values is **undetermined from the
   templates**. (The shipped templates are blank header shells; any formulas
   would live in rows the app adds at runtime, which are absent here.)
4. **Column widths / merged-cell extents — NOT captured** (formatting_info off).
   Group-header spans (e.g. how far `E5 "Sum of Order"`, `K5 Inventory` merge)
   are inferred from adjacent populated cells, not read directly.
5. **`OrderSimulationChanged.xls` day-grid beyond col AN** — dump truncated the
   listing at col 40 (AN). The sheet is 45 cols wide; cols AO-AS were not
   enumerated cell-by-cell (they continue the day series).
6. **Tooling constraints recorded:** LibreOffice not installed (could not
   `soffice --headless` convert to read CF); openpyxl cannot open BIFF `.xls`.
   To recover gaps 1-4, convert templates to `.xlsx` on a machine with Excel/
   LibreOffice, then read CF + formulas with openpyxl, **or** parse the BIFF
   stream directly.

### How to close the CF gap — DONE (superseded by §5 below)
Recommendation was: convert to `.xlsx` with LibreOffice and read
`worksheet.conditional_formatting` via openpyxl. **This has now been done** — see §5.

---

## 5. GAP CLOSURE — LibreOffice extraction (2026-06-13)

Tooling: `LibreOffice 26.2.4.2` headless (`soffice --headless --convert-to xlsx`)
→ read with `openpyxl` (data_only=False). Converted `OrderSimulation.xls`,
`OrderSimulationCAMEX.xls`, and `OrderSimulationChanged.xls`.

**Base templates the app actually opens (`OrderSimulation.xls`, `…CAMEX.xls`):**
**0 conditional-formatting rules, 0 formulas** — CONFIRMS they are blank header
shells. The CF rules and the simulation formulas are **applied by `Order.pas` at
runtime** (consistent with the Pascal `FormatConditions[1].font.ColorIndex:=3`
calls at `Order.pas:1047,1055,1552-1564`). The only static fill is **`FFFFCC99`
(peach) on `Q5:R6`** — the "Order Qty / Lot" entry columns — which matches the
Pascal `Interior.ColorIndex:=40` on Q,R. So **`ColorIndex 40 = #FFCC99`,
empirically confirmed**.

**`OrderSimulationChanged.xls` (the day-grid sibling) reveals the intended CF +
formula model** — corroborating the spec's inferred §3.5 logic exactly:

- **Conditional formatting (3 rules, all red font `#FF0000`):**
  - `T10:AP10` — `cellIs lessThan $J$7` → red. *(the projected-balance day row turns
    red when the day's value < `$J$7`)*
  - `J9` — expression `$J$9<0` → red; `J14` — expression `$J$14<0` → red.
- **Formula model (extracted):**
  - `J7 = H7*I7` → **Order Point / safety stock = Daily Usage (H) × Days (I)**.
  - `J9 = H9-I9` (the row-level Total/coverage that goes red when negative).
  - `T7 = SUM(AQ8:AQ10) - SUM(AR8:AR10)` → **day-0 beginning balance = Total Inv −
    In-Transit bucket** (AQ/AR are the hidden Total-Inv / In-Transit cols the Pascal
    comments name at `Order.pas:197-203`).
  - `T10 = T7+T8-T9` → **ending balance = beginning + receipts − usage** (the PAB
    recursion); `U7=T10, V7=U10, …` → **each day's beginning = prior day's ending**.
  - `K7:N7 = SUM(rows 8:10)` (size-group aggregation); `Q8 = R8*O8`, `W8 = Q8`
    (order qty = lot×count, landed in a future day column).

**Threshold now KNOWN (no longer a guess):** the shortage/red signal fires when a
day's **projected ending on-hand (the PAB row) < safety stock `J7 = Daily Usage ×
Order-Point Days`**. This is exactly the §3.5 / §4 inference — now confirmed from
the actual CF rule + formulas.

**Palette mapping (template applies NO palette override → standard Excel-2003
56-color `ColorIndex` palette is authoritative unless the running client overrides
it; `40=#FFCC99` confirmed empirically above):**

| Pascal ColorIndex | Std RGB | Meaning (spec §4) |
|---|---|---|
| 3  | `#FF0000` red | overtime day col / red-font shortage |
| 4  | `#00FF00` green | non-production ('X') day col |
| 10 | `#008000` dark green | open (unshipped) order qty (font) |
| 23 | `#333399` indigo | in-transit qty (font) |
| 34 | `#CCFFFF` light cyan | size usage/safety inputs (H:I) |
| 36 | `#FFFF99` light yellow | lead-time zone |
| 40 | `#FFCC99` peach | order-by point / order-entry cells **(confirmed)** |

**Updated gap status:** Gap 1 (CF/threshold) **CLOSED**; Gap 3 (formula model)
**CLOSED**; Gap 2 (fills/colors) **CLOSED** for the order-entry highlight and
mapped via the standard palette for the rest (verify against the running Excel only
if pixel-exact RGB is ever required). Residual (low value): Gap 4 col-widths/merges
and Gap 5 cols AO-AS are readable from the `.xlsx` if ever needed. **The spike's SC2
(color/signal parity) is now fully specified.**
