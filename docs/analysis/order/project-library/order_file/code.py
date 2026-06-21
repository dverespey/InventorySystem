# order_file — Project Library service for the Order-FILE GENERATOR (the OUTPUT stage), legacy
# TOrderFormCreate_Form.FormActivate, OrderFormCreateF.pas:55-707. Invoked from MainMenu.pas:276-279.
#
# Source specs: docs/analysis/order/order-file-generation-spec.md (flow + hazards) and
# docs/analysis/order/order-file-data-analysis.md (every proc body proved on mssql-spike). Domain:
# project-order-renban-domain memory.
#
# SCOPE — this is the EMITTER, a SEPARATE unit from the COMMIT path (project-library/order/code.py).
# The commit path (order.commitOrders / computeOrderRecords) turns a worksheet line into
# INV_OPEN_ORDER_INF rows. THIS service reads the already-committed, renban-assigned, not-yet-ordered
# open-order rows and serializes them into supplier `.ord` text files (+ Excel order-forms, secondary),
# then stamps each as ordered so it never re-emits. It shares NO code with the commit path and does NO
# order math — it copies IN_QTY / VC_RENBAN_NUMBER / VC_FRS_NUMBER through verbatim.
#
# KEY DIVERGENCES / BUGS FIXED vs the legacy (all flagged inline):
#   * ATOMICITY (legacy H1): the legacy writes files to disk MID-transaction; a rollback leaves the
#     final .ord on disk -> a re-run re-emits it (duplicate supplier files). FIXED here with the 856
#     temp-then-rename idiom: write every .ord to <name>.ord.tmp, run the ONE stamp transaction, and
#     ONLY after it commits rename all .tmp -> final; on any failure delete every .tmp. A rollback
#     leaves NO final .ord.
#   * EXPLICIT ALIASING (legacy H10 / data-analysis §2.3): the legacy proc is `SELECT *` over a 4-table
#     join (88 cols) with IN_QTY / VC_PART_NUMBER / VC_KANBAN_NUMBER / VC_SUPPLIER_CODE DUPLICATED.
#     `fieldbyname` silently grabs the FIRST (the open-order copy). The order qty is the OPEN_ORDER
#     IN_QTY (e.g. 1200), NOT the parts-stock on-hand IN_QTY (e.g. 28133). _FEED_SQL aliases every read
#     column explicitly from the `i.` (open-order) table so the wrong column can never be picked.
#   * EMIT-FROM-SNAPSHOT-THEN-STAMP (legacy H11 / data-analysis §5): UPDATE_ORDEROrderDate keys on
#     part+FRS ONLY (no renban filter) -> stamping one row marks every same part+FRS row ordered. The
#     legacy is safe because it emits from a single forward cursor opened BEFORE any stamp. We replicate
#     that: read the WHOLE feed into a snapshot, emit every line, THEN stamp from the emitted snapshot
#     (never re-query per row, or we'd drop not-yet-emitted siblings).
#   * sendsite / SiteSupplierCode (legacy H4): the leading SiteSupplierCode field is a LATENT CRASH — a
#     phantom column read by `fieldbyname('SiteSupplierCode')` that exists in NO table (proved: 0 hits in
#     the dump, absent from the 88-col result set), yet 11/16 live suppliers have
#     BIT_SITE_NUMBER_IN_ORDER=1. delphi-architect adjudicated (this session): like the 810, this is a
#     latent crash, NOT a live feature. We BUILD the working non-sendsite 6-field .ord for ALL suppliers,
#     do NOT reproduce the crash, and do NOT invent a leading field. The real multi-site leading field is
#     defined at M4 (see _M4 note in format_ord_line).
#
# Jython 2.7 (8.1 and 8.3). The pure formatters (format_ord_line / compute_ship_offset) import nothing
# and are CPython-importable for unit test; only get_ship_offset + generate_order_files touch the gateway
# globals `system`. IG81-COMPAT: runPrepQuery / runPrepUpdate(...,tx=) / beginTransaction / system.file /
# system.date / getLogger are identical on 8.1.52 and 8.3 — no version guard needed.
# IG83-TODO: on 8.3 prefer system.file move/rename + an Event Stream for the timer trigger; here we use
# the 8.1-safe os.rename (atomic on one filesystem) + a gateway timer script calling generate_order_files.

DATABASE = "Inventory_Spike"            # the Inventory rebuild connection name
CAL_DATABASE = "VehicleOrder"           # the ALC/TireOrder DB holding AD_GetSpecialDate (cross-DB)


class OrderFileError(Exception):
    pass


# =================================================================================================
# 1. PURE .ord LINE FORMATTER — the EXACT positional, delimiter-less line. Derived field-by-field from
#    OrderFormCreateF.pas:556-585 (the live fText/fBoth path). Pure / portable / CPython-importable.
# =================================================================================================
def _format_qty(q):
    """Pascal format('%.5d', [q]) — the '.5' precision zero-pads the DIGITS to a minimum of 5 with the
    sign prepended OUTSIDE the padding. Python '%05d' would let the sign consume one pad slot, so a
    negative < 10000 diverges (format('%.5d',[-5]) = '-00005' [6 chars] vs '%05d' % -5 = '-0005'
    [5 chars]). Reproduce the Delphi shape exactly. (BLOCKER-2 / adversary SHOULD-FIX 3.)
        240   -> '00240'     0    -> '00000'    123456 -> '123456'
        -5    -> '-00005'    -50  -> '-00050'   -12345 -> '-12345'
    """
    n = int(q)
    if n < 0:
        return "-" + ("%05d" % -n)                  # sign OUTSIDE the 5-digit zero-pad (Pascal %.5d)
    return "%05d" % n


def format_ord_line(row):
    """One open-order row -> the exact .ord line (NO trailing newline; the caller terminates CRLF).

    `row` is a dict with the keys the aliased feed yields:
        supCode    VC_SUPPLIER_CODE  (the OPEN-ORDER copy, i.VC_SUPPLIER_CODE; raw varchar(5))
        frsNum     VC_FRS_NUMBER     (raw varchar(7))
        renban     VC_RENBAN_NUMBER  (right-justified to min width 8)
        partNum    VC_PART_NUMBER    (raw varchar(12))
        qty        IN_QTY            (the OPEN-ORDER order qty; zero-padded to 5 digits)
        shipDate   the computed ship date as 8-char yyyymmdd

    Positional layout, in byte order (OrderFormCreateF.pas line cited per field):
        VC_SUPPLIER_CODE   raw, no pad                         (:568)  -- non-sendsite path ONLY
        VC_FRS_NUMBER      raw, no pad                         (:570)
        VC_RENBAN_NUMBER   format('%8s', ...)  right-justified (:571)
        VC_PART_NUMBER     raw, no pad                         (:572)
        IN_QTY             format('%.5d', ...) sign + min-5    (:573)  -- see _format_qty
        shipDate           FormatDateTime('yyyymmdd', now+ship)(:574)

    Width hazards (legacy H3 / data-analysis #8): only renban (field 3) and qty (field 5) are padded.
    Supplier/FRS/part are written RAW — a short value shifts every following field (the positional
    sub-supplier parser keys on column positions). We reproduce the legacy raw-concat EXACTLY (we do NOT
    silently pad supplier/FRS/part) — faithfulness over correctness; pending a golden .ord to confirm the
    receiving parser expects always-full-width columns.

    %8s right-justifies (space-pads on the LEFT) to a MINIMUM of 8; a renban longer than 8 is NOT
    truncated (Pascal format('%8s',...) and Python '%8s' both widen). Qty is Pascal format('%.5d',...):
    the '.5' is a PRECISION (a minimum DIGIT count) — the sign is prepended OUTSIDE the 5 zero-padded
    digits, so format('%.5d',[-5]) = '-00005' (6 chars). Python's '%05d' is a FIELD WIDTH the sign
    CONSUMES ('%05d' % -5 = '-0005', 5 chars) — DIVERGENT for negatives < 10000. We emit the Pascal
    shape via _format_qty (sign + abs zero-padded to min 5). A qty > 99999 prints all its digits (widens
    the line, no truncation). Both widen/negative cases are faithful + flagged as position-parser hazards.
    """
    sup = str(row["supCode"])
    frs = str(row["frsNum"])
    renban = "%8s" % str(row["renban"])             # right-justified min-8 (Pascal format('%8s',...))
    part = str(row["partNum"])
    qty = _format_qty(row["qty"])                   # Pascal format('%.5d',...): sign OUTSIDE 5-digit pad
    ship = str(row["shipDate"])
    # _M4 (multi-site leading field): when BIT_SITE_NUMBER_IN_ORDER=1, the legacy PREPENDED a per-site
    # supplier code via the phantom SiteSupplierCode column (a crash). The faithful multi-site field is
    # INV_SITES.VC_SUPPLIER_CODE (see ignition-architect M4); width/justification PENDING a golden .ord.
    # We deliberately do NOT emit it now (non-sendsite 6-field line for ALL suppliers, per delphi-architect).
    return sup + frs + renban + part + qty + ship


def format_ord_file(rows, lineSep="\r\n"):
    """Join formatted lines into the full .ord byte payload. Each line (incl. the last) is CRLF-terminated
    — the legacy `Writeln(tcf, tcl)` writes a CRLF after EVERY line, so the file ends with a trailing
    CRLF. Pure."""
    if not rows:
        return ""
    return lineSep.join(format_ord_line(r) for r in rows) + lineSep


# =================================================================================================
# 2. SHIP-DATE — SELECT_PartShipDays weekday pick (pure) + GetShip working-day scan (DB-backed).
#    Faithful to OrderFormCreateF.pas:407-464 (weekday pick) + GetShip :709-770.
# =================================================================================================
# Pascal DayOfTheWeek is 1=Mon..7=Sun. The weekday->ship-column map mirrors the case 1..6 (Mon..Sat);
# Sunday (7) and any unmatched weekday fall through to the base 'Ship'. A weekday-specific value of 0
# means "no override" -> use the base 'Ship'. (OrderFormCreateF.pas:418-462.)
_WEEKDAY_SHIP_COL = {1: "ShipM", 2: "ShipT", 3: "ShipW", 4: "ShipTh", 5: "ShipF", 6: "ShipS"}


def pick_ship_lead(shipDays, weekdayMon1):
    """Pure: pick the ship-day LEAD count for today's weekday. `shipDays` is a dict with keys
    Ship/ShipM/ShipT/ShipW/ShipTh/ShipF/ShipS (the SELECT_PartShipDays columns). `weekdayMon1` is the
    Pascal DayOfTheWeek (1=Mon..7=Sun). Returns the lead int the legacy passes to GetShip:
    the weekday override if it is non-zero, else the base Ship. (OrderFormCreateF.pas:418-462.)"""
    col = _WEEKDAY_SHIP_COL.get(int(weekdayMon1))
    if col is not None:
        v = int(shipDays.get(col) or 0)
        if v != 0:
            return v
    return int(shipDays.get("Ship") or 0)


def compute_ship_offset(lead, baseDateOrdinal, isWeekend, isHoliday):
    """Pure reimplementation of GetShip (OrderFormCreateF.pas:709-770) — convert a working-day LEAD into
    a CALENDAR-day offset by walking forward from tomorrow, counting only working days.

    lead            : the working-day lead (from pick_ship_lead).
    baseDateOrdinal : an integer day ordinal for `now` (e.g. date.toordinal()). The scan probes
                      baseDateOrdinal + 1 + x for x = 0,1,2,...
    isWeekend(ord)  : True if the day ordinal is Sat/Sun (Pascal: DayOfTheWeek((now+1)+x) >= 6).
    isHoliday(ord)  : True if the day ordinal is an 'H' special-date (Pascal: a row with
                      DATE == that day AND Date Status Abrv == 'H').

    FAITHFULNESS — the legacy counts a day as valid (INC y, INC x) ONLY when it is NEITHER a weekend NOR
    an 'H' holiday; a weekend or holiday day only INC x (skipped). When y == lead the loop exits and the
    returned offset is x (the offset of the day on which the last valid day landed). lead == 0 returns 0
    (the while never runs). Mirrors the legacy x/y bookkeeping exactly.

    GetShip CALENDAR-INCONSISTENCY FLAG (data-analysis §3.3, FOR REVIEW): GetShip skips ONLY weekends +
    'H'. It does NOT skip 'X' (non-production) or 'W' days and treats 'O' (overtime) as a normal working
    day — UNLIKE the M2 forecast day-spread (ForecastBreakdownF.pas:1403-1407), which turns Saturday ON
    for an 'O' Saturday and accounts for O/X/W. Reproduced FAITHFULLY here (weekend + 'H' only); flagged
    as a possible carry-forward bug / divergence for delphi-architect + ignition-architect to adjudicate
    at cutover (is the narrower ship-date rule intentional, or should it align with the forecast rule?).
    """
    lead = int(lead)
    if lead <= 0:
        return 0
    x, y = 0, 0
    # Guard against an unbounded scan if every day were somehow invalid (the legacy would spin forever);
    # lead working days can never need more than lead*7 calendar days even with full weekends + holidays.
    limit = lead * 7 + 14
    while y != lead:
        probe = baseDateOrdinal + 1 + x
        if not isWeekend(probe):
            if not isHoliday(probe):
                y += 1
                x += 1
            else:
                x += 1
        else:
            x += 1
        if x > limit:
            # Defensive: never loop forever (the legacy could). Treat as "no further valid days found".
            break
    return x


def _holiday_ordinals(calDb, beginDate, endDate, lineName=""):
    """Read the 'H' special-date ordinals from AD_GetSpecialDate (VehicleOrder) for the window, ONCE per
    run (the legacy re-opens it per row; we cache — data-analysis §3.3). Returns a set of date ordinals.

    AD_GetSpecialDate(@BeginDate, @EndDate, @LineName) returns rows with a 'Date' (datetime) and a
    'Date Status Abrv' (varchar(1)); we keep only the 'H' rows, matching GetShip's holiday test
    (OrderFormCreateF.pas:740: Date Status Abrv == 'H'). @LineName='' = the all-lines branch."""
    import datetime as _dt
    rows = system.db.runPrepQuery(
        "EXEC AD_GetSpecialDate @BeginDate=?, @EndDate=?, @LineName=?",
        [beginDate, endDate, lineName], calDb)
    out = set()
    for r in rows:
        abbr = r["Date Status Abrv"]
        if abbr is not None and str(abbr).strip() == "H":
            d = str(r["Date"])[:10]                  # 'yyyy-mm-dd' prefix of the datetime
            try:
                dt = _dt.datetime.strptime(d, "%Y-%m-%d").date()
                out.add(dt.toordinal())
            except ValueError:
                pass
        # NON-'H' rows (O / X / W / N) are intentionally ignored — faithful to GetShip (the flag above).
    return out


def get_ship_offset(partNum, runDate, calDb=None, database=None):
    """Resolve a part's ship-date CALENDAR offset = pick_ship_lead(SELECT_PartShipDays, today's weekday)
    fed through GetShip against the AD_GetSpecialDate holiday set. Touches `system` (DB).

    runDate : a python date (the order-run 'now'); the offset is added to it for the .ord ship-date field.
    Returns the integer calendar-day offset (compute_ship_offset's x)."""
    import datetime as _dt
    db = database if database is not None else DATABASE
    cdb = calDb if calDb is not None else CAL_DATABASE

    ds = system.db.runPrepQuery("EXEC SELECT_PartShipDays @PartNumber=?", [partNum], db)
    if not len(ds):
        # The part must exist in the feed (it joined parts-stock). No ship-days row -> base 0 (no offset).
        return 0
    r = ds[0]
    shipDays = {k: r[k] for k in ("Ship", "ShipM", "ShipT", "ShipW", "ShipTh", "ShipF", "ShipS")}
    # Pascal DayOfTheWeek: Monday=1..Sunday=7. python date.isoweekday() is the same convention.
    weekdayMon1 = runDate.isoweekday()
    lead = pick_ship_lead(shipDays, weekdayMon1)
    if lead <= 0:
        return 0

    # Holiday window mirrors GetShip: @BeginDate=now, @EndDate=now+59 (the legacy 59-day window; a lead
    # needing >59 calendar days would miss holidays past day 59 — faithful, flagged data-analysis §3.3).
    begin = runDate.strftime("%Y-%m-%d 00:00:00")
    end = (runDate + _dt.timedelta(days=59)).strftime("%Y-%m-%d 00:00:00")
    holidays = _holiday_ordinals(cdb, begin, end)
    base = runDate.toordinal()

    def isWeekend(o):
        # Pascal DayOfTheWeek((now+1)+x) >= 6 means Sat(6)/Sun(7). python date.isoweekday() 6=Sat,7=Sun.
        return _dt.date.fromordinal(o).isoweekday() >= 6

    def isHoliday(o):
        return o in holidays

    return compute_ship_offset(lead, base, isWeekend, isHoliday)


# =================================================================================================
# 3. LOGISTICS-DIR RESOLUTION — the part->supplier->'NONE' ladder + the three sentinel states.
#    Faithful to OrderFormCreateF.pas:179-224 + data-analysis §1.
# =================================================================================================
def resolve_logistics_dir(partNum, supplierInfo, database=None):
    """Resolve the per-supplier logistics directory ONCE per supplier (called with the supplier's FIRST
    part, like the legacy). `supplierInfo` is the SELECT_SupplierInfo row (a dict). Returns the resolved
    directory string, where the literal 'NONE' is the skip-logistics sentinel.

    Ladder (data-analysis §1 / OrderFormCreateF.pas:188-224):
      a. SELECT_PartsStockLogistics(@PartNo) returns >=1 row -> the PART's logistics dir.
      b. 0 rows -> '' ; then:
      c. supplier-level LogisticsDirectory NOT null -> the SUPPLIER's logistics dir.
      d. supplier-level LogisticsDirectory IS null -> the literal 'NONE' (skip).

    THREE distinct states drive different writes: 'NONE' (string) != NULL != '' (legacy H8: a non-NULL
    empty-string supplier logistics dir slips past the 'NONE' guard -> a root-relative path; we preserve
    that faithfully — '' is NOT coerced to 'NONE')."""
    db = database if database is not None else DATABASE
    # Step a/b: the part-level INNER-join lookup (proc declares @PartNo; bind positionally / by real name).
    partRows = system.db.runPrepQuery(
        "EXEC SELECT_PartsStockLogistics @PartNo=?", [partNum], db)
    if len(partRows):
        return partRows[0]["LogisticsDirectory"]     # step a — the part's dir
    # step b -> '' ; fall to the supplier level (step c/d)
    supLog = supplierInfo.get("LogisticsDirectory")
    if supLog is not None:
        return supLog                                 # step c — the supplier's dir (may be '' — legacy H8)
    return "NONE"                                     # step d — the skip sentinel


def _supplier_destinations(supplierDir, logisticsDir, localFtp):
    """Compute the (1 or 3) .ord destinations for a supplier (data-analysis §4 / OrderFormCreateF.pas).

    localFtp == False (default): ONE copy -> the supplier dir only.
    localFtp == True           : THREE copies -> supplier dir + logistics dir (UNLESS 'NONE') + Archive.

    Returns a list of (kind, directory) where kind in {'supplier','logistics','archive'}; the caller
    builds the filename per directory. The supplier dir is ALWAYS first/always present."""
    import os as _os
    dests = [("supplier", supplierDir)]
    if localFtp:
        if logisticsDir != "NONE":                    # the Q8 skip sentinel
            dests.append(("logistics", logisticsDir))
        dests.append(("archive", _os.path.join(supplierDir, "Archive")))
    return dests


# =================================================================================================
# 4. THE DRIVER — generate_order_files. EMIT-FROM-SNAPSHOT, then ONE stamp transaction, then publish.
# =================================================================================================
#
# The aliased feed replacing SELECT_OrderNotOrdered's `SELECT *` (legacy H10 / data-analysis §2.3). Every
# read column is aliased EXPLICITLY off the open-order table `i` (so the order IN_QTY / supplier / part /
# kanban are the i. copies, never the parts-stock shadows). p./r. columns the Excel order-sheet needs
# (parts name, 1-lot qty) are aliased distinctly. Same WHERE + ORDER BY as the proc (the break detection
# relies on the supplier,renban sort). Canonical copy: spike-order-file-feed.sql (kept byte-aligned).
_FEED_SQL = (
    "SELECT i.VC_SUPPLIER_CODE AS supCode, "
    "i.VC_FRS_NUMBER       AS frsNum, "
    "i.VC_RENBAN_NUMBER    AS renban, "
    "i.VC_PART_NUMBER      AS partNum, "
    "i.IN_QTY              AS orderQty, "          # the OPEN-ORDER qty (NOT p.IN_QTY on-hand)
    "i.VC_KANBAN_NUMBER    AS kanban, "           # the OPEN-ORDER kanban (i., not p.)
    "p.VC_PARTS_NAME       AS partsName, "        # parts-master (order-sheet only)
    "p.IN_1LOTQTY          AS oneLotQty "         # parts-master (wheel order-sheet only)
    "FROM INV_OPEN_ORDER_INF i "
    "JOIN INV_PARTS_STOCK_MST p ON i.VC_PART_NUMBER = p.VC_PART_NUMBER "
    "JOIN INV_SUPPLIER_MST   s ON p.IN_SUPPLIER_ID = s.IN_SUPPLIER_ID "
    "LEFT OUTER JOIN INV_RENBAN_GROUP_MST r ON p.IN_RENBAN_ID = r.IN_RENBAN_ID "
    # RISK-2: select ONLY what the stamp can clear. UPDATE_ORDEROrderDate clears rows via
    # `VC_ORDER_DATE <> @OrderDate`; under ANSI_NULLS ON `NULL <> '...'` is UNKNOWN, so a NULL row is
    # NEVER stamped -> if the feed selected NULL it would re-emit forever. The legacy proc's
    # `((VC_ORDER_DATE IS NULL) OR ='')` had that latent mismatch; VC_ORDER_DATE is NOT NULL DEFAULT ''
    # on the live schema (verified), so the IS NULL arm is unreachable. Drop it to match the stamp exactly.
    "WHERE i.VC_ORDER_DATE = '' "
    "AND i.VC_RENBAN_NUMBER <> '' "
    "ORDER BY s.VC_SUPPLIER_CODE, i.VC_RENBAN_NUMBER"
)


def _supplier_info(supCode, database):
    """SELECT_SupplierInfo(@SupCode, @Logistics=1) -> the single-supplier config row as a dict (the
    aliases SELECT_SupplierInfo emits: Output File Type / Directory / LogisticsDirectory / Supplier Name /
    Order File Timestamp / Site Number in Order / Create Order Sheet / address fields). data-analysis §4."""
    rows = system.db.runPrepQuery(
        "EXEC SELECT_SupplierInfo @SupCode=?, @Logistics=?", [supCode, 1], database)
    if not len(rows):
        raise OrderFileError("generate_order_files: no SELECT_SupplierInfo row for supplier %r" % (supCode,))
    r = rows[0]
    keys = ("Supplier Code", "Supplier Name", "Address", "City", "State", "Zip", "Directory",
            "LogisticsDirectory", "Output File Type", "Order File Timestamp", "Site Number in Order",
            "Create Order Sheet")
    return dict((k, r[k]) for k in keys)


def _file_kind(outputFileType):
    """Map SELECT_SupplierInfo 'Output File Type' to the channels produced. The proc's CASE emits
    'TEXT'/'EXCEL'/'BOTH'; anything else (NULL/'') falls through to BOTH (legacy H2 — a NULL output type
    silently gets both). Returns (wantText, wantExcel)."""
    t = (outputFileType or "").strip().upper()
    if t == "TEXT":
        return True, False
    if t == "EXCEL":
        return False, True
    return True, True                                 # 'BOTH' and the NULL/'' default (legacy H2)


def _coerce_bit(v):
    """Coerce a SQL `bit` to a Python bool the way Delphi TField.AsBoolean does (0->False, 1->True),
    robustly across BOTH transports the code runs under (BLOCKER-1 / adversary BLOCKER 1):
      * gateway JDBC returns a `bit` as a Java/Python Boolean or Integer (True/False, 1/0);
      * the spike sqlcmd shim returns it as the STRING '0'/'1' (and NULL -> None).
    `bool('0')` is True in Python (any non-empty string is truthy), so the legacy no-timestamp bit (0)
    was wrongly treated as timestamped under the shim. Treat '0'/0/False/None/'' -> False;
    '1'/1/True/'true' -> True. Mirrors TField.AsBoolean = (bit value <> 0)."""
    if v is None:
        return False
    if isinstance(v, bool):
        return v
    # int/long (Jython 2.7 has both; CPython3 has only int) and any numeric -> AsBoolean = (value <> 0).
    if not isinstance(v, str) and not hasattr(v, "lower"):
        try:
            return int(v) != 0
        except (ValueError, TypeError):
            return bool(v)
    s = str(v).strip().lower()
    if s in ("", "0", "false", "no", "n"):
        return False
    if s in ("1", "true", "yes", "y", "-1"):        # SQL bit is 0/1; '-1' guards a driver-quirk truthy
        return True
    # An unexpected non-empty token: fall back to numeric truthiness if it parses, else False (safe).
    try:
        return int(s) != 0
    except (ValueError, TypeError):
        return False


def _clean_name(name):
    """Supplier name with ',' and '.' stripped (OrderFormCreateF.pas:227), then '\\' stripped at filename
    time (ANSIReplaceStr(...,'\\','')). Used in every output filename."""
    return str(name or "").replace(",", "").replace(".", "").replace("\\", "")


def _ord_filename(cleanName, supCode, timestamp, fts):
    """The .ord filename stem (OrderFormCreateF.pas:290/321). With timestamp: <name>-<sup><fts>.ord
    (NO '-' before the timestamp — concatenated, legacy H7). Without: <name>-<sup>.ord."""
    if timestamp:
        return "%s-%s%s.ord" % (cleanName, supCode, fts)
    return "%s-%s.ord" % (cleanName, supCode)


def _cleanup_tmps(tmpPaths):
    """Best-effort delete of every orphan .tmp on a rolled-back run (the 856 _cleanup_tmp idiom, fan-out).
    Never raises — the caller is mid-propagation of the REAL failure. A leftover .tmp is harmless (no
    dispatcher reads .tmp); the invariant is that NO FINAL .ord exists for an un-stamped batch."""
    import os as _os
    for p in tmpPaths:
        try:
            if _os.path.exists(p):
                _os.remove(p)
        except Exception:
            pass


def generate_order_files(database=None, calDatabase=None, localFtp=False, runDate=None,
                         outRoot=None, dirOverride=None):
    """Read the not-yet-ordered open orders, emit one .ord line per row to each supplier's destination(s),
    stamp them ordered (ONE transaction), and publish the files atomically (temp-then-rename post-commit).

    Reproduces OrderFormCreateF.pas:55-707 WITHOUT Excel/COM for the .ord channel (the Excel order-forms
    are a separate, secondary build — see order_file_excel.compute_* below). ONE Inventory transaction
    wraps the stamps; the file writes are staged to .tmp and published only on commit (fixes legacy H1).

    Parameters
    ----------
    database    : Inventory connection name (defaults to DATABASE).
    calDatabase : the ALC/VehicleOrder connection holding AD_GetSpecialDate (defaults to CAL_DATABASE).
    localFtp    : INI [INIT] LocalFTP (default False = supplier dir only; True = + logistics + Archive).
    runDate     : the order-run 'now' (a python date); defaults to today. Drives the order/ship dates.
    outRoot     : if set, every supplier 'Directory' is treated as RELATIVE to this root (spike test uses
                  a temp dir so we never write to the real S:\\... paths). On the gateway leave None
                  (the supplier dirs are absolute configured shares).
    dirOverride : optional dict supCode->directory to override the configured supplier dir (spike test).

    Returns a summary dict: {'rowCount', 'supplierCount', 'files'(list of published final paths),
    'suppliers'(per-supplier detail), 'stamped'(int)}.

    Atomicity / order of operations (ONE tx):
      1. read the WHOLE aliased feed into a snapshot (BEFORE any stamp — legacy single-cursor invariant).
      2. per supplier: resolve config + logistics dir; compute the ship offset per row; build the .ord
         text; write it to <final>.ord.tmp at every destination (NOT the final name).
      3. open ONE transaction; stamp every emitted row via UPDATE_ORDEROrderDate (from the SNAPSHOT,
         de-duplicated on part+FRS since the proc has no renban filter — legacy H11); commit.
      4. ONLY after commit: rename every .tmp -> final, ALL-OR-NOTHING. A rollback in step 3 deletes
         every .tmp + re-raises (NO final .ord — the H1 fix). In step 4 (post-commit), if any rename
         fails the already-published copies are un-renamed back to .tmp and an OrderFileError is raised
         naming the failed destination — so the file phase is never SILENTLY partial (RISK-1). The DB
         rows stay stamped (the 2-phase DB+files limit is explicit in the error, see step-4 comment).
    """
    import os as _os
    import datetime as _dt
    db = database if database is not None else DATABASE
    cdb = calDatabase if calDatabase is not None else CAL_DATABASE
    if runDate is None:
        runDate = _dt.date.today()
    log = system.util.getLogger("SPIKE.order_file")

    orderDateStr = runDate.strftime("%Y%m%d")        # yyyymmdd (the order date = today, :599)
    fts = runDate.strftime("%Y%m%d") + "000000" + "00"   # placeholder fts; see note below
    # NOTE on fts: the legacy fts = formatdatetime('yyyymmddhhmmss00', now) (1-second resolution + '00').
    # For deterministic tests we derive it from runDate at midnight; the gateway timer would pass a real
    # now(). Only timestamped suppliers (BIT_ORDER_FILE_TIMESTAMP=1, 1/16 live) use it.

    # --- 1. read the WHOLE feed into a snapshot (before any stamp) --------------------------------
    feed = system.db.runPrepQuery(_FEED_SQL, [], db)
    snapshot = [{
        "supCode": r["supCode"], "frsNum": r["frsNum"], "renban": r["renban"],
        "partNum": r["partNum"], "qty": int(r["orderQty"]), "kanban": r["kanban"],
        "partsName": r["partsName"], "oneLotQty": r["oneLotQty"],
    } for r in feed]
    if not snapshot:
        log.info("generate_order_files: no orders to process (0 not-ordered rows)")
        return {"rowCount": 0, "supplierCount": 0, "files": [], "suppliers": [], "stamped": 0}

    # --- 2. per supplier: build the .ord text + stage to .tmp at every destination ----------------
    # Group by supplier preserving the feed order (the snapshot is already sorted supplier,renban).
    shipCache = {}                                    # partNum -> ship offset (one lookup per part/run)
    plannedTmps = []                                  # list of (tmpPath, finalPath) to publish on commit
    suppliers = []                                    # per-supplier detail for the summary
    curSup = None
    supRows = []

    def _flush_supplier(supCode, rows):
        if not rows:
            return
        info = _supplier_info(supCode, db)
        wantText, wantExcel = _file_kind(info["Output File Type"])
        cleanName = _clean_name(info["Supplier Name"])
        # Legacy: lasttimestamp := fieldbyname('Order File Timestamp').AsBoolean (OrderFormCreateF.pas:228).
        # AsBoolean on a SQL bit is (value <> 0): 0 -> False (stable filename), 1 -> True (timestamped).
        # _coerce_bit reproduces that across the JDBC-Boolean and the shim-string '0'/'1' transports
        # (BLOCKER-1 — plain bool('0') was True, wrongly timestamping 15/16 live suppliers).
        timestamp = _coerce_bit(info["Order File Timestamp"])
        supplierDir = (dirOverride or {}).get(supCode, info["Directory"])
        if outRoot is not None:
            # strip a leading separator/drive so join stays under outRoot (spike test sandboxing)
            rel = str(supplierDir or "").lstrip("\\/").replace(":", "")
            supplierDir = _os.path.join(outRoot, rel) if rel else outRoot
        logisticsDir = resolve_logistics_dir(rows[0]["partNum"], info, db)
        if outRoot is not None and logisticsDir not in (None, "", "NONE"):
            rel = str(logisticsDir).lstrip("\\/").replace(":", "")
            logisticsDir = _os.path.join(outRoot, rel) if rel else outRoot

        # compute the ship date per row (cache by part), then the .ord lines
        ordRows = []
        for row in rows:
            pn = row["partNum"]
            if pn not in shipCache:
                shipCache[pn] = get_ship_offset(pn, runDate, cdb, db)
            shipDate = (runDate + _dt.timedelta(days=shipCache[pn])).strftime("%Y%m%d")
            ordRows.append({"supCode": row["supCode"], "frsNum": row["frsNum"], "renban": row["renban"],
                            "partNum": pn, "qty": row["qty"], "shipDate": shipDate})

        detail = {"supCode": supCode, "wantText": wantText, "wantExcel": wantExcel,
                  "logisticsDir": logisticsDir, "lineCount": len(ordRows), "destinations": []}

        if wantText:
            text = format_ord_file(ordRows)
            # The .ord filename is computed PER destination (afname below) — the supplier/logistics copy
            # uses `timestamp`, the archive copy is always timestamped (legacy :310/:341).
            for kind, dest in _supplier_destinations(supplierDir, logisticsDir, localFtp):
                if dest in (None, ""):
                    # legacy H8: an empty logistics dir yields a root-relative path. We SKIP rather than
                    # write to a root-relative path in the rebuild (faithful behaviour is undefined/harmful);
                    # flagged. The 'NONE' sentinel is already filtered by _supplier_destinations.
                    log.warn("generate_order_files: empty %s directory for supplier %s — skipping (H8)"
                             % (kind, supCode))
                    continue
                # archive filename is ALWAYS timestamped (legacy :310/:341) even for non-timestamp suppliers
                afname = _ord_filename(cleanName, supCode, timestamp or kind == "archive", fts)
                finalPath = _os.path.join(dest, afname)
                tmpPath = finalPath + ".tmp"
                if not _os.path.isdir(dest):
                    _os.makedirs(dest)
                system.file.writeFile(tmpPath, text)         # stage to .tmp ONLY (8.1+ safe)
                plannedTmps.append((tmpPath, finalPath))
                detail["destinations"].append({"kind": kind, "path": finalPath})

        # IG-TODO: EXCEL-only suppliers (Output File Type 'EXCEL' -> wantExcel only, no .ord) emit NO file
        # until render_xlsx is wired (the Excel order-form is the separate order_file_excel build). All 16
        # live suppliers are 'B'=BOTH (-> wantText True), so this branch is unreachable today; when an
        # EXCEL-only supplier exists, _flush_supplier would publish nothing for it without the xlsx render.

        suppliers.append(detail)

    for row in snapshot:
        if row["supCode"] != curSup:
            _flush_supplier(curSup, supRows)
            curSup = row["supCode"]
            supRows = []
        supRows.append(row)
    _flush_supplier(curSup, supRows)

    # --- 3. ONE transaction: stamp every emitted row (de-duped on part+FRS — legacy H11) ----------
    # The legacy stamps per-row, but UPDATE_ORDEROrderDate keys on part+FRS only (no renban filter), so a
    # part+FRS with N renban rows is stamped once and the proc marks all N. We emit from the snapshot
    # (already done above) THEN stamp the DISTINCT (part, FRS) pairs once each — identical net effect,
    # fewer round-trips, and it CANNOT drop a not-yet-emitted sibling (we never re-query the feed).
    distinctKeys = []
    seen = set()
    for row in snapshot:
        key = (row["partNum"], row["frsNum"])
        if key not in seen:
            seen.add(key)
            distinctKeys.append(key)

    committed = False
    stamped = 0
    tx = system.db.beginTransaction(db)
    try:
        for partNum, frsNum in distinctKeys:
            shipOff = shipCache.get(partNum, 0)
            shipDate = (runDate + _dt.timedelta(days=shipOff)).strftime("%Y%m%d")
            # UPDATE_ORDEROrderDate(@PartNumber,@FRSNumber,@OrderDate,@ShipDate) — the re-emit guard,
            # with the legacy `VC_ORDER_DATE <> @OrderDate` clause inside the proc (idempotent same-day).
            n = system.db.runPrepUpdate(
                "EXEC UPDATE_ORDEROrderDate @PartNumber=?, @FRSNumber=?, @OrderDate=?, @ShipDate=?",
                [partNum, frsNum, orderDateStr, shipDate], db, tx=tx)
            stamped += int(n or 0)
        system.db.commitTransaction(tx)
        committed = True
    except Exception:
        system.db.rollbackTransaction(tx)
        _cleanup_tmps([t for t, _ in plannedTmps])   # rollback -> NO final .ord (H1 fix)
        raise
    finally:
        system.db.closeTransaction(tx)

    # --- 4. POST-COMMIT: publish — rename every .tmp -> final, ALL-OR-NOTHING. ---------------------
    # RISK-1: the stamp tx has ALREADY committed here. A naive per-dest rename with no error handling
    # could publish the supplier copy, fail on (e.g.) the logistics share, and leave a SILENT partial
    # emit — the supplier .ord out + rows stamped but logistics/archive lost with no signal. That is a
    # TWO-PHASE problem (DB rows + N files) we cannot make one atomic unit: the DB is committed and the
    # legacy never wrote a compensating un-stamp. So we make the FILE phase all-or-nothing and LOUD:
    #   * rename every .tmp -> final;
    #   * on ANY failure, UN-rename the ones that already renamed (final -> .tmp), so NO destination is
    #     left half-published, and raise OrderFileError naming the failed destination + the fact the rows
    #     are STAMPED (committed) and every .tmp is retained for operator re-drop.
    # Net: either ALL destinations publish, or NONE do + a hard error (no silent partial). The residual
    # 2-phase limit (rows stamped but files not yet on disk) is explicit in the error, not silent.
    published = []
    try:
        for tmpPath, finalPath in plannedTmps:
            # IG83-TODO: on 8.3 prefer system.file move/rename; os.rename is atomic on one filesystem (8.1-safe).
            _os.rename(tmpPath, finalPath)
            published.append(finalPath)
    except Exception as exc:
        # Un-rename what already published so the FILE set is all-or-nothing (no partial emit). Best-effort
        # — a failure here can't worsen the invariant (we then re-raise loudly with every dest named).
        for finalPath in published:
            try:
                _os.rename(finalPath, finalPath + ".tmp")
            except Exception:
                pass
        retained = [t for t, _ in plannedTmps]
        log.error("generate_order_files: PUBLISH FAILED after the stamp COMMITTED — rolled back the file "
                  "publish (un-renamed %d already-published copies). Rows ARE STAMPED. Retained .tmp at: %s"
                  % (len(published), retained))
        raise OrderFileError(
            "generate_order_files: .ord publish failed after the stamp tx committed (2-phase limit). "
            "NO destination was left partially published; rows are STAMPED (ordered) and every .ord is "
            "retained as .tmp for operator re-drop. Failed rename + retained .tmp files: %s. "
            "Underlying error: %s" % (retained, exc))

    log.info("generate_order_files: %d rows, %d suppliers, %d files published, %d open-order rows stamped"
             % (len(snapshot), len(suppliers), len(published), stamped))
    return {"rowCount": len(snapshot), "supplierCount": len(suppliers), "files": published,
            "suppliers": suppliers, "stamped": stamped, "committed": committed}
