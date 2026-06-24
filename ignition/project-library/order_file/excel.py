# order_file.excel — the EXCEL ORDER-FORM channel of the order-file generator (SECONDARY to the .ord).
#
# The legacy emits, alongside the machine-readable .ord, up to TWO Excel artifacts per supplier
# (OrderFormCreateF.pas):
#   * the "Excel order"  (OrderTemplate.xls)            — a 5-column dump (EXCEL/BOTH suppliers; all 16 live).
#   * the "Excel order sheet" (OrderSheetTemplate{Wheel,Tire}.xls) — the formatted human-readable PO
#     (suppliers with VC_CREATE_ORDER_SHEET set; all 16 live), one workbook per renban for WHEEL,
#     paginated at o>23.
#
# SCOPE + HONESTY (per the M2 brief): this module builds the DATA MODEL only — the exact rows/headers the
# Excel forms need, faithful to the Pascal Cells[row,col] writes (those coordinates are HIGH confidence,
# read directly from the .pas). The EXACT VISUAL LAYOUT of the 3 .xls templates (static labels, merges,
# column widths, fonts) is **template-visual-UNVERIFIED** — the template files were NOT confirmed on disk
# (order-file-generation-spec §10.1). So the cell PLACEMENTS below are byte-faithful to the Pascal; the
# look/labels baked into the templates are PENDING the template files (same posture as the forecast Excel
# reports). The .ord is the EDI-critical artifact and is fully built + byte-tested in code.py; THIS is the
# human-readable copy and is intentionally not blocked on the templates.
#
# A server-side renderer (render_xlsx) is a STUB: it documents the writer contract and (if an xlsx library
# is available) can emit a placeholder workbook from the data model. The real renderer is deferred until
# the templates are recovered, so the placement is wired to confirmed coordinates and nothing is invented.
#
# Pure Jython 2.7 / CPython-importable (the data-model builders import nothing). render_xlsx is a no-COM
# server-side writer (the suite's 856/862/forecast Excel-replacement pattern).


# =================================================================================================
# A. EXCEL "ORDER" (OrderTemplate.xls) — the 5-column dump. OrderFormCreateF.pas:279-280, :543-554.
# =================================================================================================
def build_excel_order(supplierInfo, ordRows):
    """The Excel-order data model for ONE supplier (the column dump). `ordRows` are the same per-row dicts
    the .ord uses (supCode/frsNum/renban/partNum/qty/shipDate) PLUS the ship date pre-formatted.

    Layout (OrderFormCreateF.pas, exact cell coords — Cells[row,col], 1-based):
        header (written once at supplier open, :279-280):
            [1,6] (row 1, col F) = Supplier Name
            [2,6] (row 2, col F) = Address
        body (one row per order line, starting i=10, :545-549):
            [i,1] A = VC_PART_NUMBER
            [i,2] B = VC_FRS_NUMBER
            [i,3] C = VC_RENBAN_NUMBER
            [i,4] D = IN_QTY
            [i,5] E = ship date, FormatDateTime('mm/dd/yyyy', now+ship)
        i increments per line; no page/overflow handling (unbounded rows).

    Ship date here is mm/dd/yyyy (the Excel format), DISTINCT from the .ord's yyyymmdd — the caller passes
    shipDateMdy per row. Returns {'header': {cell: value}, 'body': [ {cell: value}, ... ], 'startRow': 10}.
    """
    header = {
        "F1": supplierInfo.get("Supplier Name"),     # Cells[1,6]
        "F2": supplierInfo.get("Address"),           # Cells[2,6]
    }
    body = []
    i = 10
    for r in ordRows:
        body.append({
            "A%d" % i: r["partNum"],                 # Cells[i,1]
            "B%d" % i: r["frsNum"],                  # Cells[i,2]
            "C%d" % i: r["renban"],                  # Cells[i,3]
            "D%d" % i: r["qty"],                     # Cells[i,4]
            "E%d" % i: r.get("shipDateMdy", ""),     # Cells[i,5] = mm/dd/yyyy
        })
        i += 1
    return {"header": header, "body": body, "startRow": 10}


# =================================================================================================
# B. EXCEL "ORDER SHEET" (OrderSheetTemplate{Wheel,Tire}.xls) — the formatted PO. :253-261, :522-540.
# =================================================================================================
PAGE_FILL_ROW = 23                                   # o>23 -> page break (OrderFormCreateF.pas:470)
BODY_START_ROW = 16                                  # o:=16 (OrderFormCreateF.pas:265)


def _frs_to_date(frsNum, runYear):
    """Reconstruct the order-by date displayed in cell [8,8] from the FRS prefix (OrderFormCreateF
    pas:257-258). year = first-3-of-runYear + FRS[0]; date = FRS[1:3]/FRS[3:5]/year (the FRS encodes
    a 1-digit year + MMDD in its first 5 chars). runYear is 'yyyy' (the run's now-year)."""
    frs = str(frsNum or "")
    if len(frs) < 5:
        return ""
    year = str(runYear)[:3] + frs[0:1]               # century-3 + FRS leading year digit
    return "%s/%s/%s" % (frs[1:3], frs[3:5], year)   # copy(FRS,2,2)/copy(FRS,4,2)/year


def build_order_sheet(supplierInfo, ordRows, kind, runYear, shipDateMdyByRenban=None):
    """The order-sheet data model for ONE renban (WHEEL) / supplier (TIRE). `kind` is 'WHEEL' or 'TIRE'
    (VC_CREATE_ORDER_SHEET). Splits into PAGES of (PAGE_FILL_ROW - BODY_START_ROW + 1) body rows.

    Header per workbook page (Cells, exact coords — :253-261 / :390-397 / :506-513):
        [11,1] = Supplier Name        [11,2] = Supplier Code
        [11,3] = renban (WHEEL pages always; supplier-break only for WHEEL)
        [8,8]  = FRS-derived order-by date mm/dd/YYY? (_frs_to_date)
        [11,5] = page number
        [13,2] = Address              [14,2] = City, State, Zip
        [8,6]  = ship date mm/dd/yy (rewritten each line; we carry the LAST per page, :525)
    Body per line (o from 16, :522-538):
        [o,1] = VC_PART_NUMBER  [o,2] = VC_PARTS_NAME  [o,4] = VC_KANBAN_NUMBER
        WHEEL: [o,5] = IN_1LOTQTY    [o,7] = IN_QTY
        TIRE : [o,5] = VC_RENBAN_NUMBER  [o,7] = IN_QTY

    ordRows carry: partNum, partsName, kanban, oneLotQty, renban, qty, shipDateMdyShort (mm/dd/yy).
    Returns a list of page dicts: [{'page': n, 'renban': r, 'header': {...}, 'body': [...]}].

    LAYOUT-VISUAL CAVEAT: the Cells coordinates are byte-faithful to the Pascal; the static template
    labels/merges are PENDING the .xls files (see module docstring)."""
    isWheel = (str(kind or "").strip().upper() == "WHEEL")
    perPage = PAGE_FILL_ROW - BODY_START_ROW + 1     # body rows per page before the o>23 break
    pages = []
    name = supplierInfo.get("Supplier Name")
    code = supplierInfo.get("Supplier Code")
    addr = supplierInfo.get("Address")
    city = "%s, %s, %s" % (supplierInfo.get("City"), supplierInfo.get("State"), supplierInfo.get("Zip"))

    # WHEEL = one workbook PER renban (the renban break, :355); TIRE = one per supplier (no renban break).
    if isWheel:
        groups = []
        cur = None
        for r in ordRows:
            if cur is None or r["renban"] != cur[0]:
                cur = (r["renban"], [])
                groups.append(cur)
            cur[1].append(r)
    else:
        groups = [(None, ordRows)]

    for renban, rows in groups:
        pageNo = 1
        for start in range(0, max(1, len(rows)), perPage):
            chunk = rows[start:start + perPage]
            lastShip = chunk[-1].get("shipDateMdyShort", "") if chunk else ""
            firstFrs = chunk[0]["frsNum"] if chunk else (rows[0]["frsNum"] if rows else "")
            header = {
                "A11": name, "B11": code,                       # [11,1] [11,2]
                "H8": _frs_to_date(firstFrs, runYear),          # [8,8]
                "E11": pageNo,                                  # [11,5]
                "B13": addr, "B14": city,                       # [13,2] [14,2]
                "F8": lastShip,                                 # [8,6] (rewritten per line; last per page)
            }
            if isWheel and renban is not None:
                header["C11"] = renban                          # [11,3] WHEEL only
            body = []
            o = BODY_START_ROW
            for r in chunk:
                cell = {
                    "A%d" % o: r["partNum"],                    # [o,1]
                    "B%d" % o: r.get("partsName"),              # [o,2]
                    "D%d" % o: r.get("kanban"),                 # [o,4]
                    "G%d" % o: r["qty"],                        # [o,7] IN_QTY (both variants)
                }
                if isWheel:
                    cell["E%d" % o] = r.get("oneLotQty")        # [o,5] IN_1LOTQTY (WHEEL)
                else:
                    cell["E%d" % o] = r["renban"]               # [o,5] renban (TIRE)
                body.append(cell)
                o += 1
            pages.append({"page": pageNo, "renban": renban, "kind": "WHEEL" if isWheel else "TIRE",
                          "header": header, "body": body})
            pageNo += 1
    return pages


# =================================================================================================
# C. SERVER-SIDE RENDER STUB — no-COM xlsx writer. Deferred until the 3 templates are recovered.
# =================================================================================================
def render_xlsx(dataModel, outPath):
    """Server-side workbook writer for an order-form data model (the suite's no-COM Excel pattern).

    STUB — intentionally not a full renderer. The exact cell PLACEMENT contract is fixed (build_excel_order
    / build_order_sheet emit A1-style cell->value maps faithful to the Pascal Cells[r,c]); what is PENDING
    is the VISUAL template (labels/merges/fonts) baked into OrderTemplate.xls /
    OrderSheetTemplate{Wheel,Tire}.xls — those files were not confirmed on disk. Building a renderer now
    would invent a layout. When the templates are recovered, this becomes: load the template workbook,
    write each (cell -> value) from the data model, SaveAs outPath (xlsx) — mirroring the 856/862/forecast
    Excel-replacement writers already in the suite.

    Raising NotImplementedError keeps a caller from silently shipping a blank/invented order form; the
    .ord (the EDI-critical artifact) is fully built and tested in code.py regardless."""
    raise NotImplementedError(
        "order_file.excel.render_xlsx is a stub: the OrderTemplate / OrderSheetTemplate{Wheel,Tire}.xls "
        "VISUAL layouts are unverified (not confirmed on disk). The data model (build_excel_order / "
        "build_order_sheet) is byte-faithful to the Pascal Cells[r,c]; wire a no-COM template-fill writer "
        "once the templates are recovered. The .ord channel is fully built + tested in code.py.")
