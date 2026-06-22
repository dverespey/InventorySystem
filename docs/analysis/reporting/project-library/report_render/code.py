# -*- coding: utf-8 -*-
# report_render — Project Library: the M3 server-side report-render engine.
#
# Replaces the legacy CreateOleObject('Excel.Application') -> open ReportTemplate.xls -> write
# mysheet.cells[r,c] -> SaveAs(<name><ts>.xls) -> optional PrintOut path (the FAILING Excel/printer
# coupling) with a HEADLESS Apache-POI render run on the Ignition Gateway. Design:
#   docs/analysis/reporting/m3-render-architecture.md (§1 mechanism, §2 harness, §4 8.1->8.3 deltas).
#
# A report = (a) a Named-Query result + (b) a render def -> an .xlsx, mirroring the legacy
# `mysheet.cells[r,c]`. Two render lanes share ONE engine + ONE output stage:
#   * DECLARATIVE (header band + column band + optional totals) for reports whose layout is a flat table
#     (Daily Shipping Assy, INVOICE Summary). Adding one = ONE def entry + ONE .sql, no code.
#   * CUSTOM-PLAN for the two reports whose legacy render is NOT a flat table — it has group-break rows
#     and blank separators that the declarative model cannot express faithfully (Forecast Detail's
#     part#/eff-month group breaks with a blank separator row; Lot Location's two-proc union + line/size
#     group headers + conditional cells). Each is a PURE plan function registered by key. This is the
#     architect's "config not bespoke" lever applied honestly: 2 of 4 are pure config; the 2 genuinely
#     non-declarative ones get a small named pure function rather than an invented generic group-break
#     feature the design (m3-render-architecture §2.2) does not specify. (STOP-and-flag: see the M3
#     finding — these two exceeded the declarative spec; built faithfully from the .pas, cited per fn.)
#
# SPLIT (the M1/M2 producer-recipe pattern — pure logic vs gateway I/O, so the module imports cleanly
# under scripts/e2e/jython_shim.py for headless drive):
#   * build_*_plan(...)  — PURE: render def/rows -> an ordered (row, col, value, num_fmt) cell list +
#                          a {col: width} map. Imports nothing, no gateway global, no POI. CPython-testable.
#   * build_dataset(...) — the ONLY DB touch: system.db.runPrepQuery (NQ-shaped, promotable -> runNamedQuery).
#   * render_xlsx(...)   — POI (XSSF) bytes from a cell plan. Imports org.apache.poi at CALL time only.
#   * write_output(...)  — tmp-then-rename to fiReportsOutputDir, legacy yyyymmddhhmmss00 filename. NO PrintOut.
#   * run_report(...)    — the thin driver entry (driver.py): dataset -> plan -> xlsx -> write.
#
# Jython 2.7 (8.1 and 8.3, no delta). IG81-COMPAT: POI 4.1.2 is on lib/core/common/ on BOTH versions
# (m3-render-architecture §1.1) — the data+XLSX numbers-gate path is version-neutral + headless. PDF/print
# is the only 8.3-specific surface (guarded, non-gating; not built here).

import os as _os


# ============================================================================================
# PURE CELL-PLAN BUILDERS  — render def + rows -> cells. No POI, no system.db, no I/O.
# A "cell" is a 4-tuple (row, col, value, num_fmt), 1-based row/col EXACTLY like legacy mysheet.cells[r,c]
# (so a layout extracted from the .pas maps 1:1). num_fmt is an Excel format string or None.
# ============================================================================================

def _resolve_header_value(source, params, header_row, report_def):
    """Resolve one header_cell source spec to (value, num_fmt). Forms:
        "title"                       -> the report title literal
        ("param", name)               -> params[name]
        ("param", name, num_fmt)      -> params[name], with an Excel number-format
        ("param", name, prefix, "p")  -> prefix + str(params[name])  (label-prefixed string)
        ("header", col)               -> header_row[col]             (single header-query row)
        ("header", col, prefix)       -> prefix + str(header_row[col])
        ("lit", text)                 -> text literal
    Pure dict/string access — no gateway calls."""
    if source == "title":
        return report_def.get("title", ""), None
    kind = source[0]
    if kind == "param":
        name, val = source[1], (params or {}).get(source[1])
        if len(source) == 4 and source[3] == "p":
            return source[2] + ("" if val is None else str(val)), None
        return val, (source[2] if len(source) >= 3 else None)
    if kind == "header":
        col = source[1]
        val = header_row[col] if header_row is not None else None
        if len(source) >= 3:
            return source[2] + ("" if val is None else str(val)), None
        return val, None
    if kind == "lit":
        return source[1], None
    raise ValueError("report_render: unknown header_cell source %r" % (source,))


def build_declarative_plan(report_def, detail_rows, header_row=None, params=None):
    """PURE: a flat header-band + column-band + optional totals report -> (cells, widths).

    report_def keys: title, header_cells [(r,c,source)], header_row (int col-title row), columns
    [(header_text, query_col, width, num_fmt)], first_data_row (z start), totals (optional, see below).
    Used by Daily Shipping Assy + INVOICE Summary. detail_rows are dict-like (PyDataset row[colName])."""
    cells, widths = [], {}

    # header band (legacy rows 1-2)
    for (r, c, source) in report_def.get("header_cells", []):
        val, num_fmt = _resolve_header_value(source, params, header_row, report_def)
        cells.append((r, c, val, num_fmt))

    # column band: column-title row (legacy row 3) + widths
    columns = report_def.get("columns", [])
    hdr_row = report_def.get("header_row")
    for ci, coldef in enumerate(columns, start=1):
        if hdr_row is not None:
            cells.append((hdr_row, ci, coldef[0], None))
        if len(coldef) >= 3 and coldef[2] is not None:
            widths[ci] = coldef[2]

    # detail rows (legacy z := first_data_row .. n) + optional per-row computed item-total
    z0 = report_def.get("first_data_row", (hdr_row + 1) if hdr_row else 1)
    totals = report_def.get("totals")
    item = totals.get("item") if totals else None      # {"col", "factors": (a,b), "num_fmt"}
    z = z0
    grand = 0.0
    for row in detail_rows:
        for ci, coldef in enumerate(columns, start=1):
            qcol = coldef[1]
            num_fmt = coldef[3] if len(coldef) >= 4 else None
            cells.append((z, ci, (row[qcol] if qcol is not None else None), num_fmt))
        if item:
            prod = _num(row[item["factors"][0]]) * _num(row[item["factors"][1]])
            grand += prod
            cells.append((z, item["col"], prod, item.get("num_fmt")))
        z += 1

    # grand-total band (legacy: INC(z);INC(z); 'INVOICE TOTAL'; =SUM(...))
    if totals:
        gap = totals.get("gap", 2)
        tz = (z - 1) + gap + 1                          # two blank rows past the last detail row
        for (col, src) in totals.get("cells", []):
            if src == "__grand__":
                cells.append((tz, col, grand, totals.get("grand_num_fmt")))
            else:
                cells.append((tz, col, src, None))      # a literal label

    return cells, widths


def build_forecast_detail_plan(report_def, detail_rows, header_row=None, params=None):
    """PURE custom plan for REPORT_ForecastDetail (R18). Faithful to MainMenu.pas:3768-3794:
      z := 4; lastpn := ''
      for each row:
        if lastpn <> part#:  INC(z); [z,1]=part#; [z,2]=eff_month; lastpn=part#; lastmn=eff_month
        if lastmn <> eff_month:  [z,2]=eff_month; lastmn=eff_month
        [z,3]=bc [z,4]=tire# [z,5]=tireRatio [z,6]=wheel# [z,7]=wheelRatio [z,8]=ratio
        INC(z)
    => one BLANK separator row before each new part group (the INC(z) on part-change), part#/eff-month
    printed only at a group's first row, ratios are integers. Header band + col titles are declarative
    (rows 1 + 3); only the detail loop is custom."""
    cells, widths = [], {}
    cells.append((1, 1, report_def.get("title", ""), None))
    columns = report_def["columns"]
    hdr_row = report_def.get("header_row", 3)
    for ci, coldef in enumerate(columns, start=1):
        cells.append((hdr_row, ci, coldef[0], None))
        if len(coldef) >= 3 and coldef[2] is not None:
            widths[ci] = coldef[2]
        # (legacy sets Columns[4/6].NumberFormat := '############' on the tire#/wheel# columns, but those
        #  cells hold AsString part-number codes -> the format is a visual no-op; we render them as strings.)
    z = report_def.get("first_data_row", 4)
    lastpn, lastmn = "", ""
    PART, MONTH = "VC_ASSY_PART_NUMBER_CODE", "VC_EFFECTIVE_MONTH"
    for row in detail_rows:
        pn = _s(row[PART]); mn = _s(row[MONTH])
        if lastpn != pn:
            z += 1
            cells.append((z, 1, pn, None))
            cells.append((z, 2, mn, None))
            lastpn, lastmn = pn, mn
        if lastmn != mn:
            cells.append((z, 2, mn, None))
            lastmn = mn
        cells.append((z, 3, _s(row["VC_BROADCAST_CODE"]), None))
        cells.append((z, 4, _s(row["VC_TIRE_PART_NUMBER_CODE"]), None))
        cells.append((z, 5, _int(row["IN_TIRE_RATIO"]), None))
        cells.append((z, 6, _s(row["VC_WHEEL_PART_NUMBER_CODE"]), None))
        cells.append((z, 7, _int(row["IN_WHEEL_RATIO"]), None))
        cells.append((z, 8, _int(row["IN_RATIO"]), None))
        z += 1
    return cells, widths


def build_lot_location_plan(report_def, datasets, header_row=None, params=None):
    """PURE custom plan for the LIVE Lot Location report (R12). Faithful to MainMenu.pas:716-861.
    `datasets` is a dict {"wheels": <rows>, "tires": <rows>} (the two PLANT procs — the W/wheels proc
    runs FIRST, then the tires proc, into one sheet). Reproduces:
      * title [1,1] = '<Plant> Lot Location'; col titles row 3 (Size/Parking/Trailer/Renban/Arrived/
        <Assembler> Loc/<Plant> Loc), 7 cols with the legacy widths.
      * WHEELS (z starts at 5): line-name group break -> 'Car Wheels' (Passenger) at current z, else
        INC(z) then 'Truck Wheels'. Per row: col2 = assembler_location or (if '') plant_parking;
        col3 trailer; col4 renban; col5 = vc_status_PLANT_yard reformatted yyyy/mm/dd (copy 1-4/5-2/7-2).
      * TIRES (continues z): size-code group break -> INC(z); size header ('<NO SIZE>' if blank) at col1.
        Per row: same col2/3/4/5 mapping. (Page-break + border/alignment chrome is NOT reproduced — it is
        print presentation, not a number; XLSX is the gate-bearing artifact, m3-render-architecture §4.)
    Header/col-title text uses params['PlantName'] / params['AssemblerName'] (gateway config, not a client
    INI). Widths: 11/11/14/11/10/12/12 (MainMenu.pas:752-779)."""
    p = params or {}
    plant = p.get("PlantName", "")
    assembler = p.get("AssemblerName", "")
    cells = []
    cells.append((1, 1, plant + " Lot Location", None))
    titles = ["Size", "Parking", "Trailer", "Renban", "Arrived", assembler + " Loc", plant + " Loc"]
    for ci, t in enumerate(titles, start=1):
        cells.append((3, ci, t, None))
    widths = {1: 11, 2: 11, 3: 14, 4: 11, 5: 10, 6: 12, 7: 12}

    z = 5
    # --- wheels (REPORT_PLANTLotLocationW) ---
    lastwheel = ""
    for row in datasets.get("wheels", []):
        ln = _s(row["vc_LINE_NAME"])
        if lastwheel != ln:
            if ln == "Passenger":
                cells.append((z, 1, "Car Wheels", None))
            else:
                z += 1
                cells.append((z, 1, "Truck Wheels", None))
            lastwheel = ln
        cells.append((z, 2, _parking(row), None))
        cells.append((z, 3, _s(row["vc_trailer_number"]), None))
        cells.append((z, 4, _s(row["vc_renban_number"]), None))
        cells.append((z, 5, _yard_date(row["vc_status_PLANT_yard"]), None))
        z += 1
    # --- tires (REPORT_PLANTLotLocation) ---
    lastsize = "X"
    for row in datasets.get("tires", []):
        sz = _s(row["vc_size_code"])
        if sz != lastsize:
            z += 1
            cells.append((z, 1, sz if sz != "" else "<NO SIZE>", None))
            lastsize = sz
        cells.append((z, 2, _parking(row), None))
        cells.append((z, 3, _s(row["vc_trailer_number"]), None))
        cells.append((z, 4, _s(row["vc_renban_number"]), None))
        cells.append((z, 5, _yard_date(row["vc_status_PLANT_yard"]), None))
        z += 1
    return cells, widths


def _parking(row):
    """col2: vc_ASSEMBLER_location, or vc_PLANT_parking when assembler is '' (MainMenu.pas:809/848)."""
    a = _s(row["vc_ASSEMBLER_location"])
    return a if a != "" else _s(row["vc_PLANT_parking"])


def _yard_date(v):
    """vc_status_PLANT_yard reformatted to yyyy/mm/dd via copy(1,4)/copy(5,2)/copy(7,2) (MainMenu.pas:817)."""
    s = _s(v)
    return s[0:4] + "/" + s[4:6] + "/" + s[6:8]


# ---- pure coercion helpers (no gateway) ----
def _num(v):
    if v is None or v == "":
        return 0.0
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _int(v):
    """Faithful to fieldbyname(...).AsInteger: NULL/'' -> 0, else int."""
    if v is None or v == "":
        return 0
    try:
        return int(v)
    except (TypeError, ValueError):
        return int(_num(v))


def _s(v):
    return "" if v is None else _u(v)


def _u(v):
    try:
        return unicode(v)            # Jython 2.7
    except NameError:
        return str(v)                # CPython 3 (shim/unit-test side)


# ============================================================================================
# DATA LAYER  — the ONLY system.db touch.
# ============================================================================================
def build_dataset(sql_text, args, database=None):
    """Run a report's resolved .sql via system.db.runPrepQuery -> PyDataset. `system` is a gateway global
    at call time (never imported) so the module imports cleanly under the shim. args binds the ? markers.
    IG83-TODO: when each .sql becomes a Named Query, swap for system.db.runNamedQuery(path, paramsDict)."""
    return system.db.runPrepQuery(sql_text, args, database)   # noqa: F821 (gateway global)


# ============================================================================================
# POI RENDER  — cell plan -> .xlsx bytes. POI imported at CALL time only.
# ============================================================================================
def render_xlsx(cells, widths):
    """Build an .xlsx (XSSF) from a cell plan and return the bytes. POI is on the gateway common classpath
    on 8.1.52 AND 8.3 (IG81-COMPAT). Imported INSIDE the function so the module body never needs the
    classpath (CPython-importable for the pure unit test; the test drives the real render via the gateway/
    shim, never by importing POI here).

    Numeric value -> numeric cell (+ a cached data-format style if num_fmt); else a string cell. Column
    widths are the legacy Columns[c].ColumnWidth (Excel char units) * 256 (POI's 1/256-char unit)."""
    from org.apache.poi.xssf.usermodel import XSSFWorkbook   # noqa: E402 (call-time, gateway classpath)
    from java.io import ByteArrayOutputStream                 # noqa: E402

    wb = XSSFWorkbook()
    sheet = wb.createSheet("Report")
    fmt_helper = wb.getCreationHelper().createDataFormat()
    style_cache = {}

    def style_for(num_fmt):
        st = style_cache.get(num_fmt)
        if st is None:
            st = wb.createCellStyle()
            st.setDataFormat(fmt_helper.getFormat(num_fmt))
            style_cache[num_fmt] = st
        return st

    for (r, c, value, num_fmt) in cells:
        row = sheet.getRow(r - 1) or sheet.createRow(r - 1)
        cell = row.getCell(c - 1) or row.createCell(c - 1)
        if _is_number(value):
            cell.setCellValue(float(value))
            if num_fmt:
                cell.setCellStyle(style_for(num_fmt))
        else:
            cell.setCellValue("" if value is None else _u(value))

    for (c, width) in widths.items():
        sheet.setColumnWidth(c - 1, int(width) * 256)

    bos = ByteArrayOutputStream()
    wb.write(bos)
    wb.close()
    return bos.toByteArray()


def _is_number(v):
    if isinstance(v, bool):
        return False
    return isinstance(v, (int, float))


# ============================================================================================
# OUTPUT  — tmp-then-rename to fiReportsOutputDir, legacy yyyymmddhhmmss00 filename. NO PrintOut.
# ============================================================================================
def legacy_timestamp(now=None):
    """Reproduce formatdatetime('yyyymmddhhmmss00', now): 14 digits + the LITERAL '00' suffix (NOT
    centiseconds — m3-report-inventory.md §0). Filenames stay continuous with the legacy ones."""
    if now is None:
        now = system.date.now()                           # noqa: F821 (gateway global / shim)
    return system.date.format(now, "yyyyMMddHHmmss") + "00"   # noqa: F821


def write_output(xlsx_bytes, out_name, out_dir, now=None):
    """Write the .xlsx bytes to out_dir/<out_name><ts>.xlsx via tmp-then-rename (the M2 atomic-write
    discipline — never expose a half-written report). `.xls -> .xlsx`: legacy wrote .xls (BIFF); we emit
    .xlsx (XSSF) — same Excel opens it; no-op divergence (m3-render-architecture §7). Returns the path."""
    filename = out_name + legacy_timestamp(now) + ".xlsx"
    path = _os.path.join(out_dir, filename)
    tmp = path + ".tmp"
    system.file.writeFile(tmp, xlsx_bytes)                 # noqa: F821 (8.1+ system.file; shim provides)
    if _os.path.exists(path):
        _os.remove(path)
    _os.rename(tmp, path)
    return path


def _log(msg):
    """SPIKE log anchor (the e2e count anchor + an audit line). Quiet if no gateway logger."""
    try:
        system.util.getLogger("SPIKE").info(msg)           # noqa: F821
    except Exception:
        pass
