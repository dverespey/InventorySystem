# forecast_distribution — Project Library service for the FORECAST-DISTRIBUTION feed (the OUTBOUND
# emitter that forwards TEMA's imported forecast DOWN to each sub-supplier as a `.frc` text file and/or
# an Excel workbook). Legacy: ForecastBreakdownF.pas:340-568 (the SELECT_ForecastSupplier supplier loop).
#
# Source-truth spec: docs/analysis/order/forecast-distribution-sourcetruth.md (THE spec — read first).
# This is the OUTBOUND complement of the M2 830 import (docs/analysis/edi/inbound/project-library/
# forecast/code.py) and structurally the SAME shape as the already-built `.ord` order-file emit
# (docs/analysis/order/project-library/order_file/code.py) — built by cloning that seam.
#
# SCOPE — this is the EMITTER. It READS the freshly-written future-week buckets out of
# INV_BREAKDOWN_FC_INF (the SAME table the 830 import owns + populates) and serializes them; it does NO
# breakdown/ratio/day-spread math (M2 already did that). It shares NO code with the import path. It fires
# as the POST-IMPORT side-effect of import_830 (the legacy runs the emit after the EDI-dir scan import),
# AND is a separate callable so it is testable independently (matches the legacy Execute flow).
#
# KEY DIVERGENCES / FIXES vs the legacy (all flagged inline + cited):
#   * SiteSupplierCode (the P6 crash / unfinished mod, legacy H-SITE): at ForecastBreakdownF.pas:488 the
#     sendsite branch reads a PHANTOM `SiteSupplierCode` column that exists in NO table -> "field not
#     found" -> the run aborts for that supplier (11/16 live suppliers have BIT_SITE_NUMBER_IN_ORDER=1).
#     The ORIGINAL line was `tcl := Data_Module.fiSupplierCode.AsString` (commented out at :489) — i.e.
#     the SITE's own supplier code (TSiteInfo.SiteSupplierCode). In the rebuild that is
#     INV_SITES.VC_SUPPLIER_CODE (the spike INV_SITES note maps VC_SUPPLIER_CODE <- TSiteInfo.
#     SiteSupplierCode), EXACTLY as the `.ord` rebuild's BIT_SITE_NUMBER_IN_ORDER leading field resolves
#     it (_M4). We DO NOT reproduce the crash, and we DO NOT "drop the prefix for all" (that would break
#     the positional parse for the sendsite suppliers). We prepend INV_SITES.VC_SUPPLIER_CODE when
#     `sendsite` (BIT_SITE_NUMBER_IN_ORDER=1), else the plain VC_SUPPLIER_CODE line. The exact byte/width
#     of the leading field is GOLDEN-PENDING (P7 — same stance as the `.ord` leading field).
#   * ATOMICITY (legacy H1): the legacy Rewrite/Writeln writes to disk immediately + is non-transactional;
#     a failure mid-emit leaves a partial `.frc`. FIXED with the order-file/856 temp-then-rename idiom:
#     stage every file (main + archive) for a supplier to `<final>.tmp`, and publish (rename .tmp ->
#     final) ALL-OR-NOTHING PER SUPPLIER only after that supplier's file(s) complete; on any failure
#     mid-supplier delete every staged .tmp for that supplier (no partial final file).
#   * NULL/blank Output File Type -> BOTH (legacy H2): SELECT_SupplierInfo's CASE has no ELSE, so a NULL
#     VC_OUTPUT_FILE yields a NULL "Output File Type" -> the Pascal `else` -> fBoth. Reproduced via
#     _file_kind (cloned from order_file).
#   * H-CLEAN (filename cleaning differs from `.ord`): this feed strips ONLY '/' from the supplier name
#     (ForecastBreakdownF.pas:374/450), NOT order_file._clean_name's broader ','/'.'/'\\'. _clean_frc_name
#     reproduces the '/'-only cleaning.
#
# Jython 2.7 (8.1 and 8.3, no delta). The pure formatters (format_frc_line / format_frc_file /
# _format_qty / build_forecast_excel_plan) import nothing gateway-y -> CPython-importable for the unit
# test; only emit_forecast_distribution + _supplier_info + render_xlsx touch the gateway globals (system,
# org.apache.poi). IG81-COMPAT: runPrepQuery / system.file / getLogger + POI 4.1.2 are identical on
# 8.1.52 and 8.3 (the report_render M3 lane proved POI is version-neutral + headless). IG83-TODO: on 8.3
# prefer system.file move/rename + an Event Stream for the import trigger; here we use the 8.1-safe
# os.rename (atomic on one filesystem).

from db_shared import CONNECTION as DATABASE   # centralized DB-conn name (default Inventory_Spike; single prod-rename point)


class ForecastDistributionError(Exception):
    pass


# =================================================================================================
# 1. PURE .frc LINE FORMATTER — the EXACT positional, delimiter-less line. Derived field-by-field from
#    ForecastBreakdownF.pas:484-508 (the live fText/fBoth path). Pure / portable / CPython-importable.
# =================================================================================================
def _format_qty(q, width=5):
    """Pascal Format('%.Nd', [q]) — the '.N' is a PRECISION (a MINIMUM DIGIT count), NOT a field width:
    the sign is prepended OUTSIDE the zero-pad. Python '%0Nd' would let the sign consume one pad slot, so
    a negative < 10^(N-1) diverges: Format('%.5d',[-5]) = '-00005' (6 chars) vs Python '%05d' % -5 =
    '-0005' (5 chars). A value > 10^N-1 prints all its digits (widens the line, no truncation). This is
    the SAME formatter the order-file emit uses (order_file._format_qty); duplicated here so this module
    is self-contained + CPython-importable without an order_file dependency.
        240   -> '00240'     0    -> '00000'    123456 -> '123456'
        -5    -> '-00005'    -50  -> '-00050'   -12345 -> '-12345'
    width=2 reproduces the week-number Format('%.2d', ...): 24 -> '24', 5 -> '05', 100 -> '100'."""
    n = int(q)
    fmt = "%0" + str(int(width)) + "d"
    if n < 0:
        return "-" + (fmt % -n)                      # sign OUTSIDE the zero-pad (Pascal %.Nd)
    return fmt % n


def format_frc_line(row, siteSupplierCode=None):
    """One INV_BREAKDOWN_FC_INF row -> the exact .frc line (NO trailing newline; the caller terminates
    CRLF). Reproduces ForecastBreakdownF.pas:484-508 byte-for-byte.

    `row` is a dict with the SELECT_ForecastSupplier (= INV_BREAKDOWN_FC_INF) column values:
        supCode    VC_SUPPLIER_CODE  (raw varchar(5), NO pad)
        partNum    VC_PART_NUMBER    (raw varchar(12), NO pad)
        weekDate   VC_WEEK_DATE      (raw varchar(8) yyyymmdd)
        weekNumber IN_WEEK_NUMBER    (Format('%.2d', ...) min-2 zero-pad)
        qty1..qty7 IN_QTY1..7        (Format('%.5d', ...) min-5 zero-pad each; NULL -> 0 -> '00000')

    `siteSupplierCode` is the INV_SITES.VC_SUPPLIER_CODE leading field, prepended ONLY when the supplier
    is `sendsite` (BIT_SITE_NUMBER_IN_ORDER=1). Pass None for the non-sendsite line. See H-SITE above.

    Positional layout, in byte order (ForecastBreakdownF.pas line cited per field):
        [SiteSupplierCode] when sendsite                                  (:488 — the P6 fix, GOLDEN-PENDING)
        VC_SUPPLIER_CODE   raw, no pad                                    (:490 / :494)
        VC_PART_NUMBER     raw, no pad                                    (:497)
        VC_WEEK_DATE       raw, no pad (yyyymmdd)                         (:498)
        IN_WEEK_NUMBER     Format('%.2d', ...)  min-2 zero-pad           (:499)
        IN_QTY1..7         Format('%.5d', ...)  min-5 zero-pad each       (:500-:506)

    Width hazards (legacy H-OVF / H3 / H-NEG): only the week-number (min-2) + qtys (min-5) are padded.
    Supplier/part/week-date are RAW — a short value shifts every following field (a positional sub-supplier
    parser keys on column positions). We reproduce the legacy raw-concat EXACTLY (faithfulness over
    correctness; pending a golden .frc). A qty > 99999 prints all digits (widens, no truncation); a
    negative qty (shouldn't occur in a forecast) formats as '-00005', not '-0005' (Pascal %.5d)."""
    parts = []
    if siteSupplierCode is not None:
        parts.append(_s(siteSupplierCode))           # the P6 site-supplier-code lead (sendsite only)
    parts.append(_s(row["supCode"]))                 # raw, no pad (:490/:494)
    parts.append(_s(row["partNum"]))                 # raw, no pad (:497)
    parts.append(_s(row["weekDate"]))                # raw yyyymmdd (:498)
    parts.append(_format_qty(_as_int(row.get("weekNumber")), 2))             # %.2d (:499)
    for k in ("qty1", "qty2", "qty3", "qty4", "qty5", "qty6", "qty7"):       # %.5d each (:500-:506)
        parts.append(_format_qty(_as_int(row.get(k)), 5))
    return "".join(parts)


def format_frc_file(rows, siteSupplierCode=None, lineSep="\r\n"):
    """Join formatted lines into the full .frc byte payload. Each line (incl. the LAST) is CRLF-terminated
    — the legacy `Writeln(tcf, tcl)` (:508) writes a CRLF after EVERY line, so the file ends with a
    trailing CRLF. Pure. `rows` is the list of dicts for ONE supplier."""
    if not rows:
        return ""
    return lineSep.join(format_frc_line(r, siteSupplierCode) for r in rows) + lineSep


# =================================================================================================
# 2. PURE EXCEL CELL-PLAN — the 11-column data table. Faithful to ForecastBreakdownF.pas:466-481.
#    Reuses the M3 report_render declarative model (header band row 1 + column band row 2+). No template.
# =================================================================================================
# The legacy ForecastTemplate.xls supplied only the row-1 HEADER styling; the data is written cell-by-cell
# (mysheet.Cells[i,c]). POI builds the sheet fresh — no template needed (m3 report_render lane). The
# row-1 header titles are NOT recoverable from the .pas (they live in the binary template); we emit
# reasonable labels (GOLDEN-PENDING for exact header fidelity — see spec §8.3). NOTE the Excel INCLUDES
# VC_SIZE_CODE (col 1) which the .frc text OMITS.
_EXCEL_HEADERS = ["Size", "Part Number", "Week Date", "Week No",
                  "Qty1", "Qty2", "Qty3", "Qty4", "Qty5", "Qty6", "Qty7"]


def build_forecast_excel_plan(rows, headers=None):
    """PURE: one supplier's INV_BREAKDOWN_FC_INF rows -> a (cells, widths) cell-plan for report_render's
    POI render lane. A "cell" is a 4-tuple (row, col, value, num_fmt), 1-based EXACTLY like the legacy
    mysheet.Cells[i,c]. Faithful to ForecastBreakdownF.pas:466-481:

        row 1 (header band): the 11 column titles (the legacy template's row-1 header; labels GOLDEN-PENDING)
        row i>=2 (data): one row per record, cols 1-11 (ForecastBreakdownF.pas:468-478):
            1  VC_SIZE_CODE        (:468) -- Excel-ONLY; the .frc OMITS it
            2  VC_PART_NUMBER      (:469)
            3  VC_WEEK_DATE        (:470)
            4  IN_WEEK_NUMBER      (:471)
            5-11 IN_QTY1..7        (:472-:478)

    ALL cells are written as STRINGS (the legacy `.value := fieldbyname(...).AsString`, :468-478) — the
    qtys/week-num are NOT %.Nd-formatted in the Excel path (unlike the .frc text), just the raw AsString
    value (NULL -> '' via _s). So a numeric-looking cell is still emitted as text (num_fmt=None + a string
    value -> report_render.render_xlsx writes a string cell). Returns (cells, widths)."""
    hdrs = headers if headers is not None else _EXCEL_HEADERS
    cells = []
    for ci, title in enumerate(hdrs, start=1):
        cells.append((1, ci, title, None))           # row 1 header band
    i = 2                                             # legacy i := 2 (:446); row 1 is the header
    for r in rows:
        cells.append((i, 1, _s(r.get("sizeCode")), None))    # VC_SIZE_CODE (:468) — Excel only
        cells.append((i, 2, _s(r["partNum"]), None))         # VC_PART_NUMBER (:469)
        cells.append((i, 3, _s(r["weekDate"]), None))        # VC_WEEK_DATE (:470)
        cells.append((i, 4, _s(r.get("weekNumber")), None))  # IN_WEEK_NUMBER (:471) AsString (raw, NOT %.2d)
        for off, k in enumerate(("qty1", "qty2", "qty3", "qty4", "qty5", "qty6", "qty7")):
            cells.append((i, 5 + off, _s(r.get(k)), None))   # IN_QTY1..7 (:472-:478) AsString
        i += 1
    widths = {}                                       # no template widths to reproduce (binary template)
    return cells, widths


# ---- pure coercion helpers (no gateway) -----------------------------------------------------------
def _s(v):
    """fieldbyname(...).AsString: NULL -> '' (the legacy AsString of a NULL field is the empty string)."""
    if v is None:
        return ""
    try:
        return unicode(v)                            # Jython 2.7
    except NameError:
        return str(v)                                # CPython 3 (shim/unit-test side)


def _as_int(v):
    """fieldbyname(...).AsInteger: NULL/'' -> 0 (a NULL int field's AsInteger is 0), else int."""
    if v is None:
        return 0
    s = _s(v).strip()
    if s == "":
        return 0
    try:
        return int(s)
    except (ValueError, TypeError):
        try:
            return int(float(s))
        except (ValueError, TypeError):
            return 0


def _clean_frc_name(name):
    """Supplier name with ONLY '/' stripped (ForecastBreakdownF.pas:374/450:
    ANSIReplaceStr(SupplierName,'/','')). H-CLEAN: this is DELIBERATELY NARROWER than the `.ord` emit's
    order_file._clean_name (which also strips ','/'.'/'\\') — do NOT reuse that here. A supplier name with
    a comma/period/backslash lands in a DIFFERENT filename than the .ord for the same supplier."""
    return _s(name).replace("/", "")


def _file_kind(outputFileType):
    """Map SELECT_SupplierInfo 'Output File Type' to the channels produced. The proc's CASE emits
    'TEXT'/'EXCEL'/'BOTH'; anything else (NULL/'') falls through to BOTH (legacy H2 — a NULL output type
    silently gets both; ForecastBreakdownF.pas:425-430 has no ELSE). Returns (wantText, wantExcel).
    Identical semantics to order_file._file_kind."""
    t = _s(outputFileType).strip().upper()
    if t == "TEXT":
        return True, False
    if t == "EXCEL":
        return False, True
    return True, True                                 # 'BOTH' and the NULL/'' default (legacy H2)


def _coerce_bit(v):
    """Coerce a SQL `bit` to a bool the way Delphi TField.AsBoolean does (0->False, 1->True), robust
    across BOTH transports (gateway JDBC Boolean/Integer + the spike sqlcmd shim's STRING '0'/'1'). NB
    `bool('0')` is True in Python (any non-empty string is truthy) — so a plain bool() would wrongly read
    the no-sendsite bit (0) as True. Mirrors order_file._coerce_bit / TField.AsBoolean = (bit <> 0)."""
    if v is None:
        return False
    if isinstance(v, bool):
        return v
    if not isinstance(v, str) and not hasattr(v, "lower"):
        try:
            return int(v) != 0
        except (ValueError, TypeError):
            return bool(v)
    s = _s(v).strip().lower()
    if s in ("", "0", "false", "no", "n"):
        return False
    if s in ("1", "true", "yes", "y", "-1"):
        return True
    try:
        return int(s) != 0
    except (ValueError, TypeError):
        return False


# =================================================================================================
# 3. THE DB LAYER — the feed read + per-supplier config. The ONLY system.db touch.
# =================================================================================================
# SELECT_ForecastSupplier;1 @WeekDate -> SELECT * FROM INV_BREAKDOWN_FC_INF WHERE VC_WEEK_DATE > @WeekDate
# ORDER BY vc_supplier_code, vc_part_number, vc_week_date. We wrap the proc with an inline EXEC so the NQ
# convention (one NQ per proc) is honored at cutover; the ORDER BY on vc_supplier_code drives the
# new-file-per-supplier break (§2). Future-weeks-only (strict >, lexicographic on yyyymmdd = chronological).
_FEED_SQL = "EXEC dbo.SELECT_ForecastSupplier @WeekDate = ?"


def _read_feed(db, weekDate):
    """Read the future-week buckets via SELECT_ForecastSupplier (sorted supplier,part,weekdate) into a
    snapshot of dicts. weekDate = today yyyymmdd (the strict-> filter floor)."""
    feed = system.db.runPrepQuery(_FEED_SQL, [weekDate], db)             # noqa: F821 (gateway global)
    snapshot = []
    for r in feed:
        snapshot.append({
            "supCode": r["VC_SUPPLIER_CODE"],
            "partNum": r["VC_PART_NUMBER"],
            "weekDate": r["VC_WEEK_DATE"],
            "weekNumber": r["IN_WEEK_NUMBER"],
            "sizeCode": r["VC_SIZE_CODE"],
            "qty1": r["IN_QTY1"], "qty2": r["IN_QTY2"], "qty3": r["IN_QTY3"], "qty4": r["IN_QTY4"],
            "qty5": r["IN_QTY5"], "qty6": r["IN_QTY6"], "qty7": r["IN_QTY7"],
        })
    return snapshot


def _supplier_info(supCode, database):
    """SELECT_SupplierInfo(@SupCode, @Logistics=0) -> the single-supplier config row as a dict. NOTE this
    feed does NOT pass @Logistics=1 (unlike the .ord emit) — ForecastBreakdownF.pas:419 passes only
    @SupCode (so @Logistics defaults 0). The aliases this feed consumes (proc body CreateInventory.sql:
    5993): Supplier Code / Supplier Name / Directory / Output File Type / Site Number in Order."""
    rows = system.db.runPrepQuery(                                       # noqa: F821 (gateway global)
        "EXEC dbo.SELECT_SupplierInfo @SupCode = ?, @Logistics = ?", [supCode, 0], database)
    if not len(rows):
        raise ForecastDistributionError(
            "emit_forecast_distribution: no SELECT_SupplierInfo row for supplier %r" % (supCode,))
    r = rows[0]
    keys = ("Supplier Code", "Supplier Name", "Directory", "Output File Type", "Site Number in Order")
    return dict((k, r[k]) for k in keys)


def _site_supplier_code(database):
    """The single site's own supplier code = INV_SITES.VC_SUPPLIER_CODE (the spike note maps
    VC_SUPPLIER_CODE <- TSiteInfo.SiteSupplierCode, i.e. the legacy Data_Module.fiSupplierCode the
    commented-out :489 line read). This is the leading field prepended for sendsite suppliers (the P6
    fix). Returns '' if no site row or a NULL code (degrade gracefully — the sendsite line then carries an
    empty lead, flagged in the summary).

    # IG-SITE (single-site-correct, NOT a bug): the TOP 1 ... ORDER BY IN_SITE_ID always returns the ONE
    # INV_SITES row (e.g. 'MAS') for EVERY sendsite supplier — and that is CORRECT for our committed
    # deployment model: single-site, exactly one INV_SITES row per gateway/DB (the M4 multi-site reversal;
    # the separate-infra decision says multi-site won't return). The sendsite leading field is the SITE's
    # own supplier code, which is genuinely site-global, so a single SELECT is right. If multi-site ever DID
    # come back, this leading field would need PER-SUPPLIER site scoping — but there is NO site FK on
    # INV_BREAKDOWN_FC_INF or INV_SUPPLIER_MST today, so that resolution does not exist to do per-supplier
    # and would be a schema change, not a code tweak here. Do not "fix" this to a join."""
    rows = system.db.runPrepQuery(                                       # noqa: F821 (gateway global)
        "SELECT TOP 1 VC_SUPPLIER_CODE FROM INV_SITES ORDER BY IN_SITE_ID", [], database)
    if not len(rows):
        return ""
    return _s(rows[0]["VC_SUPPLIER_CODE"])


# =================================================================================================
# 4. THE DRIVER — emit_forecast_distribution. Per-supplier emit-to-.tmp then publish ALL-OR-NOTHING.
# =================================================================================================
def _cleanup_tmps(tmpPaths):
    """Best-effort delete of every staged .tmp for a supplier whose emit failed (the order-file/856
    idiom). Never raises (the caller is mid-propagation of the REAL failure). A leftover .tmp is harmless
    (no dispatcher reads .tmp); the invariant is that NO FINAL file exists for a failed supplier."""
    import os as _os
    for p in tmpPaths:
        try:
            if _os.path.exists(p):
                _os.remove(p)
        except Exception:
            pass


def _frc_paths(supplierDir, cleanName, supCode, localFtp, weekDate, outRoot=None):
    """Compute the (1 or 2) .frc destinations for a supplier. Returns a list of (kind, finalPath):
        ('main',    <Directory>\\<cleanName>-<supCode>.frc)                        (ForecastBreakdownF.pas:450)
        ('archive', <Directory>\\Archive\\<cleanName>-<supCode><yyyymmdd>.frc)     (:456, LocalFTP only)
    The archive date is concatenated with NO separator before '.frc' (legacy H-ARCH-SEP, :456) — two runs
    the same day overwrite the archive (day resolution only). Reproduced faithfully. outRoot sandboxes the
    configured absolute share under a temp dir for the spike test."""
    import os as _os
    base = supplierDir
    if outRoot is not None:
        rel = _s(supplierDir).lstrip("\\/").replace(":", "")
        base = _os.path.join(outRoot, rel) if rel else outRoot
    main = _os.path.join(base, "%s-%s.frc" % (cleanName, supCode))
    dests = [("main", main)]
    if localFtp:
        archive = _os.path.join(base, "Archive", "%s-%s%s.frc" % (cleanName, supCode, weekDate))
        dests.append(("archive", archive))
    return dests


def _excel_path(supplierDir, cleanName, supCode, outRoot=None):
    """The Excel SaveAs path. Legacy SaveAs is `<Directory>/<cleanName>-<supCode>-Forecast` with NO
    extension (ForecastBreakdownF.pas:376/523 — Excel/COM appended .xls). The POI rebuild emits .xlsx (the
    report_render lane), so we append '.xlsx' (same Excel opens it; a no-op divergence, m3 §7)."""
    import os as _os
    base = supplierDir
    if outRoot is not None:
        rel = _s(supplierDir).lstrip("\\/").replace(":", "")
        base = _os.path.join(outRoot, rel) if rel else outRoot
    return _os.path.join(base, "%s-%s-Forecast.xlsx" % (cleanName, supCode))


def emit_forecast_distribution(database=None, localFtp=False, runDate=None, weekDate=None,
                               outRoot=None, dirOverride=None, renderModule=None):
    """Read the future-week forecast buckets (INV_BREAKDOWN_FC_INF via SELECT_ForecastSupplier), and per
    supplier emit a `.frc` text file and/or an Excel workbook, publishing each supplier's file(s)
    atomically (stage to .tmp, then rename to final ALL-OR-NOTHING per supplier).

    Reproduces ForecastBreakdownF.pas:340-568 WITHOUT Excel/COM (POI replaces it, via the M3
    report_render lane). The supplier loop walks the supplier-sorted feed; a VC_SUPPLIER_CODE break starts
    a new supplier's file set. Fixes legacy H1 (atomicity), H-SITE (the SiteSupplierCode crash ->
    INV_SITES.VC_SUPPLIER_CODE prefix for sendsite), H2 (NULL output type -> BOTH), H-CLEAN ('/'-only
    name clean), H-NEG/H-OVF (Pascal %.Nd qty shape).

    Parameters
    ----------
    database    : Inventory connection name (defaults to DATABASE).
    localFtp    : INI [INIT] LocalFTP (default False = supplier dir only; True = + the \\Archive\\ copy).
    runDate     : the run 'now' (a python date); defaults to today. Drives weekDate (the >today filter)
                  + the archive filename's date suffix.
    weekDate    : override the SELECT_ForecastSupplier @WeekDate floor (yyyymmdd). Defaults to
                  runDate.strftime('%Y%m%d') (= today, ForecastBreakdownF.pas:353). The spike test pins it
                  so the synthetic future-dated rows are deterministically in/out of the window.
    outRoot     : if set, every supplier 'Directory' is treated as RELATIVE to this root (the spike test
                  uses a temp dir so we never write to the real configured shares). On the gateway leave
                  None (the supplier dirs are absolute configured shares).
    dirOverride : optional dict supCode->directory to override the configured supplier dir (spike test).
    renderModule: the report_render module (for the Excel POI render). The gateway/shim injects the real
                  module; required only when a supplier wants Excel. If None and Excel is wanted, the
                  module is imported by name ('report_render') at call time.

    Returns a summary dict: {'rowCount', 'supplierCount', 'files'(published final paths), 'suppliers'
    (per-supplier detail), 'siteSupplierCode'}.

    Atomicity / order of operations (PER SUPPLIER — no cross-supplier transaction; the emit is read-only
    on the DB, so each supplier is an independent all-or-nothing file unit, matching the legacy
    per-supplier file lifecycle):
      1. read the WHOLE feed into a snapshot, resolve the site supplier code ONCE.
      2. group by VC_SUPPLIER_CODE (the feed is supplier-sorted -> a break = a new file set).
      3. per supplier: resolve config; build the .frc text (+ the Excel bytes if wanted); stage each
         file to <final>.tmp; then rename every .tmp -> final. On ANY failure mid-supplier, delete that
         supplier's staged .tmp (no partial final) and re-raise.
    """
    import os as _os
    import datetime as _dt
    db = database if database is not None else DATABASE
    if runDate is None:
        runDate = _dt.date.today()
    if weekDate is None:
        weekDate = runDate.strftime("%Y%m%d")        # today yyyymmdd (ForecastBreakdownF.pas:353)
    archiveDate = runDate.strftime("%Y%m%d")         # the archive filename date suffix (:456)
    log = system.util.getLogger("SPIKE.forecast_distribution")          # noqa: F821 (gateway global)

    # --- 1. read the feed snapshot + the site supplier code ---------------------------------------
    snapshot = _read_feed(db, weekDate)
    siteSupCode = _site_supplier_code(db)            # the sendsite leading field (INV_SITES.VC_SUPPLIER_CODE)
    if not snapshot:
        log.info("emit_forecast_distribution: no future-dated forecast buckets (> %s) — nothing to emit"
                 % weekDate)
        return {"rowCount": 0, "supplierCount": 0, "files": [], "suppliers": [],
                "siteSupplierCode": siteSupCode}

    # --- 2. group by supplier, preserving the feed order (already sorted supplier,part,weekdate) ----
    published = []
    suppliers = []

    def _flush_supplier(supCode, rows):
        if not rows:
            return
        info = _supplier_info(supCode, db)
        wantText, wantExcel = _file_kind(info["Output File Type"])       # H2: NULL -> BOTH
        cleanName = _clean_frc_name(info["Supplier Name"])               # H-CLEAN: '/'-only
        sendsite = _coerce_bit(info["Site Number in Order"])             # BIT_SITE_NUMBER_IN_ORDER
        supplierDir = (dirOverride or {}).get(supCode, info["Directory"])
        # H-SITE (the P6 fix): when sendsite, the .frc lines lead with INV_SITES.VC_SUPPLIER_CODE (the
        # site's own supplier code, the legacy fiSupplierCode); else no lead. NOT the phantom column.
        lead = siteSupCode if sendsite else None

        detail = {"supCode": supCode, "wantText": wantText, "wantExcel": wantExcel,
                  "sendsite": sendsite, "lineCount": len(rows), "files": []}

        plannedTmps = []                             # (tmpPath, finalPath) staged for THIS supplier
        try:
            # --- .frc text (+ archive) ---
            if wantText:
                text = format_frc_file(rows, lead)   # CRLF-terminated incl. last line (:508)
                for kind, finalPath in _frc_paths(supplierDir, cleanName, supCode, localFtp,
                                                  archiveDate, outRoot):
                    destDir = _os.path.dirname(finalPath)
                    if not _os.path.isdir(destDir):
                        _os.makedirs(destDir)
                    tmpPath = finalPath + ".tmp"
                    # X12/ASCII text with explicit \r\n — write bytes so no newline translation occurs.
                    system.file.writeFile(tmpPath, text)                 # noqa: F821 (8.1+ system.file)
                    plannedTmps.append((tmpPath, finalPath))
                    detail["files"].append({"kind": "frc-" + kind, "path": finalPath})

            # --- Excel (.xlsx via POI) ---
            if wantExcel:
                rm = renderModule
                if rm is None:
                    import report_render as rm        # gateway project-library module (M3)
                cells, widths = build_forecast_excel_plan(rows)
                xlsx = rm.render_xlsx(cells, widths)
                finalPath = _excel_path(supplierDir, cleanName, supCode, outRoot)
                destDir = _os.path.dirname(finalPath)
                if not _os.path.isdir(destDir):
                    _os.makedirs(destDir)
                tmpPath = finalPath + ".tmp"
                system.file.writeFile(tmpPath, xlsx)                     # noqa: F821 (bytes -> binary)
                plannedTmps.append((tmpPath, finalPath))
                detail["files"].append({"kind": "xlsx", "path": finalPath})
                # NOTE: the legacy ALSO writes an Archive copy of the Excel on LocalFTP
                # (ForecastBreakdownF.pas:527-530). Reproduced for parity with the .frc archive.
                if localFtp:
                    archDir = _os.path.join(destDir, "Archive")
                    if not _os.path.isdir(archDir):
                        _os.makedirs(archDir)
                    archFinal = _os.path.join(archDir, _os.path.basename(finalPath))
                    archTmp = archFinal + ".tmp"
                    system.file.writeFile(archTmp, xlsx)                 # noqa: F821
                    plannedTmps.append((archTmp, archFinal))
                    detail["files"].append({"kind": "xlsx-archive", "path": archFinal})

        except Exception:
            # H1 fix: any failure WHILE STAGING (render/write to .tmp, before the publish phase) -> delete
            # this supplier's staged .tmp (no partial final; nothing was renamed yet). The publish phase is
            # handled separately BELOW (outside this try) so its retained-rollback .tmp are NOT cleaned away
            # by this branch — they are the operator re-drop payload the publish error promises.
            _cleanup_tmps([t for t, _ in plannedTmps])
            log.error("emit_forecast_distribution: FAILED staging supplier %s — staged .tmp cleaned up"
                      % supCode)
            raise

        # --- publish: rename every .tmp -> final, ALL-OR-NOTHING for THIS supplier ----------------------
        # P6 (the code-reviewer RISK; mirrors order_file step-4 / P10/R26): a supplier's .frc + .xlsx (+
        # archives) are INDEPENDENT files renamed in sequence. A naive loop that renames the .frc, then
        # THROWS on the .xlsx rename, would leave the supplier HALF-EMITTED — a published .frc with no .xlsx
        # (a partial supplier the downstream parser would mis-handle). So we make the rename phase all-or-
        # nothing PER SUPPLIER: on any mid-sequence failure, UN-rename the finals already published back to
        # .tmp (best-effort) so NO final is left for this supplier, then report the TRUE on-disk state (which
        # finals are STUCK as final, which .tmp remain — all FS-verified, never a blanket claim). One
        # supplier's failure never touches another supplier's already-completed files (published[] from
        # earlier _flush_supplier calls is untouched). This phase is OUTSIDE the staging try/except so the
        # retained-rollback .tmp survive (the H1 cleanup above only fires for a STAGING failure).
        supPublished = []
        try:
            for tmpPath, finalPath in plannedTmps:
                # IG83-TODO: on 8.3 prefer system.file move/rename; os.rename is atomic on one filesystem.
                if _os.path.exists(finalPath):
                    _os.remove(finalPath)            # legacy pre-deletes before SaveAs (:374/:521)
                _os.rename(tmpPath, finalPath)
                supPublished.append(finalPath)
        except Exception as renameExc:
            # Un-rename what already published so THIS supplier's file set is all-or-nothing (no half emit).
            # The un-rename is itself best-effort: it CAN fail (e.g. the share vanished), leaving a final
            # stuck on disk. Track which un-renames succeeded vs failed so the error reports the ACTUAL
            # state, never the old blanket "staged .tmp cleaned up" (which over-claimed when a final had
            # already been published — the P10 over-claim lesson, R20).
            rolled_back = []     # finals successfully un-renamed back to .tmp (no longer a final on disk)
            still_final = []     # finals whose un-rename FAILED -> STILL published on disk (half emit!)
            for finalPath in supPublished:
                try:
                    _os.rename(finalPath, finalPath + ".tmp")
                    rolled_back.append(finalPath + ".tmp")
                except Exception:
                    still_final.append(finalPath)
            # .tmp actually retained on disk now = the rolled-back finals (now .tmp) + every dest that NEVER
            # published (its .tmp was never renamed away). FS-verify so we never list a phantom .tmp.
            never_published = [t for (t, f) in plannedTmps if f not in supPublished]
            retained_tmp = [p for p in (rolled_back + never_published) if _os.path.exists(p)]
            if still_final:
                log.error("emit_forecast_distribution: PUBLISH FAILED for supplier %s and the file "
                          "rollback was INCOMPLETE — %d final(s) still on disk (could not un-rename): %s | "
                          ".tmp retained: %s"
                          % (supCode, len(still_final), still_final, retained_tmp))
                raise ForecastDistributionError(
                    "emit_forecast_distribution: publish failed for supplier %s AND the file rollback was "
                    "INCOMPLETE. The supplier is PARTIALLY emitted: %d final file(s) could NOT be rolled "
                    "back and STILL HOLD a final on disk: %s. The rest are retained as .tmp: %s. Underlying "
                    "error: %s" % (supCode, len(still_final), still_final, retained_tmp, renameExc))
            # Clean rollback: every already-published final for this supplier was un-renamed.
            log.error("emit_forecast_distribution: PUBLISH FAILED for supplier %s — rolled back the file "
                      "publish (un-renamed %d already-published final(s); NO final left on disk for this "
                      "supplier). Retained .tmp: %s" % (supCode, len(rolled_back), retained_tmp))
            raise ForecastDistributionError(
                "emit_forecast_distribution: publish failed for supplier %s (all-or-nothing per supplier). "
                "No final was left on disk for this supplier (the %d already-published final(s) were rolled "
                "back to .tmp); every staged file is retained as .tmp: %s. Underlying error: %s"
                % (supCode, len(rolled_back), retained_tmp, renameExc))
        published.extend(supPublished)

        suppliers.append(detail)

    curSup = None
    supRows = []
    for row in snapshot:
        if row["supCode"] != curSup:
            _flush_supplier(curSup, supRows)
            curSup = row["supCode"]
            supRows = []
        supRows.append(row)
    _flush_supplier(curSup, supRows)

    log.info("emit_forecast_distribution: %d rows, %d suppliers, %d files published (siteSupCode=%r)"
             % (len(snapshot), len(suppliers), len(published), siteSupCode))
    return {"rowCount": len(snapshot), "supplierCount": len(suppliers), "files": published,
            "suppliers": suppliers, "siteSupplierCode": siteSupCode}
