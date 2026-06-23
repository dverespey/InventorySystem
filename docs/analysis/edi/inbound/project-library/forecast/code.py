# forecast — Project Library: the M2 EDI 830 (DELFOR planning forecast) IMPORTER. Populates the forecast
# that M1's ASN-create (SELECT_ForecastDetailBCASN) + the Order sim (SELECT_ForecastPartNumberWeek) read.
# Producer-recipe pattern, the twin of edi_inbound/edi856/edi810 code.py: PURE Jython-2.7/CPython parse +
# explode + day-spread, plus a thin gateway DRIVER (import_830) that reuses the M1 DUNS guard +
# idempotency ledger + alarm table.
#
# SOURCE TRUTH (all on disk; line cites are the EXTRACTED algorithm — see forecast-import-algorithm.md):
#   ForecastBreakdownF.pas (LIVE, InventorySystem.dpr:28) — the legacy importer:
#       parse        :184-298  (LIN03 assembly :263, LIN05 kanban :264, FST01 qty :288,
#                               FST04 weekdate :282, FST09 week copy(delSL[9],3,2) :281)
#       explode      :1081-1312 (ratio pick :1151-1176, counts :1186-1226, fan to components :1230-1287)
#       day-spread   :1314-1480 (workdays :1321-1328, calendar H/X :1394-1410, spread :1414-1434,
#                               store checkweeknumber RAW :1446)
#       write order  :320-335   (DeleteBreakdown FIRST :322-329, then UpdateForecast :332)
#   DELETE_ForecastInfo            /tmp/inv_utf8.sql:2725  (assembly->component CROSS APPLY, trim both ends)
#   INSERTUPDATE_BreakdownForecastInfo :1215 (ADDITIVE), INSERTUPDATE_ForecastInfo :1184 (REPLACE)
#   SELECT_ForecastDetail :3154, SELECT_PartsStockInfo :6143, SELECT_FirstProductionDay :1884,
#   AD_GetSpecialDateWeek (VehicleOrder DB) — body read on spike
#   docs/analysis/edi/inbound/forecast-import-algorithm.md  — the line-by-line extraction (READ FIRST)
#   docs/analysis/decisions.md D10 — week numbering (golden-validated)
#
# LOCKED DECISIONS realized here:
#   D10  store the RAW FST09 week (checkweeknumber) in IN_WEEK_NUMBER. The FirstProductionDay offset is
#        applied ONLY to the holiday-calendar lookup week, NEVER to the stored row. (The R1 catch — a
#        literal port of ForecastBreakdownF.pas:1374 onto the stored value drifts every row +1.)
#   E    DELETE_ForecastInfo runs FIRST per assembly (the additive breakdown upsert is only a *replace*
#        after the delete clears the forward slice), in ONE tx per file.
#   Q11  per-site AUTO/MANUAL import mode + last-import stamp + 8-day staleness alarm (DATA side; the
#        gateway scheduled poll + the home-hub box are prod wiring).
#
# BUGS WE FIX (documented divergences — forecast-import-algorithm.md §F):
#   D-Bug-1  NO silent-drop on a BOM-ratio-lookup miss -> a 830_FORECAST_GAP alarm row (visible, not
#            vanished). Legacy logged + dropped the count (ForecastBreakdownF.pas:1178-1183).
#   D-Bug-2  usage rollup reads the SAME production-relative week the row is stored under (ISO-offset),
#            not the legacy raw-ISO WeekOfTheYear(now) (:1052) -> usage reconciles to the stored qtys.
#   D-Bug-3  REVERSED to PARITY (BLOCKER-2): the day-spread now reproduces the legacy running days counter
#            EXACTLY (H/X -> off + DEC; 'O' overtime / any non-H/X -> ON + INC, :1398-1407), incl. the
#            latent already-on double-count. The former "True only if previously off" divergence IGNORED
#            'O' and so diverged from the legacy on the live COROLLA O-Saturdays. See day_spread + §D.5.
# PARITY FIXES (this branch — adversary BLOCKER/RISK fixes, all reproduce-the-legacy):
#   BLOCKER-1  breakdown VC_SUPPLIER_CODE = the COMPONENT's part-master supplier (:1349), NOT the feed
#              supplier (the feed supplier is the RAW row only, :1107). See import_830 step 2d.
#   BLOCKER-2  'O' (overtime) days turn ON in the day-spread (above). See day_spread / _calendar_rows.
#   RISK-1     delete anchor = min(weekDate) across all entries (a documented divergence safer than the
#              legacy "last LIN's first week"; equal on real data). See import_830.
#   RISK-2     _process_one_830 reads per-site IN_HISTORICAL_FORECAST + BIT_USE_FIRST_PRODUCTION_DAY from
#              INV_SITES (the columns are live, not dead). See _read_site_forecast_config.
#
# Jython 2.7 (8.1 and 8.3, no delta — pure string slicing + system.db; no Py3-only libs).
#
# This module holds BOTH:
#   * parse_830 / explode / day_spread + helpers — PURE. Reference no gateway global at import time ->
#     CPython-importable for the unit test (scripts/e2e/test_forecast_import_build.py). No DB, no I/O.
#   * import_830 / the driver helpers — the GATEWAY DRIVER. References `system` only at CALL time, so
#     importing the module never needs the gateway runtime; driven headlessly via jython_shim.


# =================================================================================================
# PURE: parse_830, explode, day_spread + the recipe model. Import nothing gateway-y.
# =================================================================================================

class ForecastImportError(Exception):
    """Raised on a structural parse/processing error. The driver rolls back the per-file tx on it (the
    file is NOT marked processed, so a fixed re-drop can re-run it)."""
    pass


def _seg_id(line):
    """The 3-char segment id = the legacy copy(fcl,1,3) (ForecastBreakdownF.pas:227,250,271,292). Pascal
    copy is 1-based start 1 len 3 -> Python line[0:3]. '' for a short/blank line so a trailing blank line
    ends a loop cleanly."""
    return (line or "")[0:3]


def _lines(text):
    """Split an EDI file body into segment lines (one X12 segment per physical line — TEMA writes it that
    way; the legacy Readln walk relies on it, never on a '~' terminator). Strip \\r so CRLF and LF both
    work; drop nothing else (a blank line just yields '')."""
    if text is None:
        return []
    return [ln.rstrip("\r") for ln in text.split("\n")]


def _split_star(line):
    """splitString('*', line) the way ForecastBreakdownF.pas:127-147 does: one token per '*', the text
    BEFORE each separator, starting delSL[0]=the segment id. Python str.split('*') yields exactly that
    (plus the trailing token after the last '*'). So delSL[N] == split('*')[N], 1-based after the seg id."""
    return (line or "").split("*")


def _elem(parts, idx):
    """delSL[idx] with a guard: '' if the element is missing (a short segment). The legacy would raise on
    a missing element; we return '' so a malformed FST is skipped rather than aborting the whole file
    (the driver's per-assembly try still records what parsed)."""
    return parts[idx] if (parts is not None and len(parts) > idx) else ""


def parse_830(text):
    """PURE 830 (DELFOR) parser. Returns a list of per-assembly entries, in file order:
        {'assembly': <LIN03 str>, 'kanban': <LIN05 str>,
         'weeks': [{'weekNumber': int, 'weekDate': '<yyyymmdd>', 'weekCount': int}, ...]}

    Reproduces the legacy LIN/FST walk (ForecastBreakdownF.pas:245-298) by the SAME element offsets (the
    parity oracle, spec §1.3 / algorithm §A):
      LIN: assembly = delSL[3] (LIN03)            -> the assembly / broadcast (BC) part = the BOM key
           kanban   = delSL[5] (LIN05)
      FST: weekNumber = int(delSL[9][2:4])        -> copy(FST09,3,2) chars 3-4 of the "DO" ref (D10).
                                                     STORED RAW (production-relative ISO-1); see day_spread.
           weekDate   = delSL[4] (FST04)          -> yyyymmdd forecast start date
           weekCount  = int(delSL[1]) (FST01)     -> the bucket forecast qty
    Walk: advance to a LIN (or stop at CTT = end-of-transaction); per LIN, collect every following FST
    until the next LIN/CTT/EOF. The SUPPLIER is NOT in the file (legacy reads fiSupplierCode :262) — the
    driver threads in the per-site supplier; parse_830 does not invent one.

    The element offsets assume TEMA's LIN = 'LIN**BP*<assy>*..*<kanban>' (delSL[3]=assy, delSL[5]=kanban)
    and FST = 'FST*<qty>*..*..*<yyyymmdd>*..*..*..*..*<DOref>'. Data-dependent vs a real golden 830 (the
    D10 golden EDI/830000008976.EDI is gitignored client data; the parse is offset-faithful to the .pas).

    Parameters
    ----------
    text : the full 830 file body (str). Already DUNS-guarded + envelope-checked by the driver.

    Returns the list of entry dicts. An empty list means no LIN segments were found (envelope-only /
    non-830) — the driver treats that as nothing-to-do, not an error. A FST whose qty/week is non-numeric
    raises ForecastImportError (a corrupt feed; the driver rolls back rather than write garbage)."""
    lines = _lines(text)
    entries = []
    i, n = 0, len(lines)
    while i < n:
        if _seg_id(lines[i]) == "CTT":
            break                                       # end-of-transaction (algorithm §A.2)
        if _seg_id(lines[i]) != "LIN":
            i += 1
            continue
        lin = _split_star(lines[i])
        entry = {"assembly": _elem(lin, 3).strip(),     # LIN03 (kept as-is below; strip for the key)
                 "kanban": _elem(lin, 5).strip(),
                 "weeks": []}
        i += 1
        # collect FSTs until the next LIN / CTT / EOF.
        while i < n and _seg_id(lines[i]) not in ("LIN", "CTT"):
            if _seg_id(lines[i]) == "FST":
                fst = _split_star(lines[i])
                do_ref = _elem(fst, 9)
                wk_raw = do_ref[2:4]                     # copy(delSL[9],3,2) chars 3-4
                week_date = _elem(fst, 4).strip()        # FST04 yyyymmdd
                qty_raw = _elem(fst, 1).strip()          # FST01 qty
                try:
                    week_number = int(wk_raw)
                    week_count = int(qty_raw)
                except (ValueError, TypeError):
                    raise ForecastImportError(
                        "parse_830: non-numeric FST for assembly %r: DOref=%r qty=%r (segment %r)"
                        % (entry["assembly"], do_ref, qty_raw, lines[i]))
                entry["weeks"].append({
                    "weekNumber": week_number,           # RAW FST09 week (D10 — stored unmodified)
                    "weekDate": week_date,
                    "weekCount": week_count,
                })
            i += 1
        entries.append(entry)
    return entries


# A recipe row model (mirrors INV_FORECAST_DETAIL_INF via SELECT_ForecastDetail aliases). The driver reads
# these from the DB; the PURE explode takes them as plain dicts so the unit test can build them inline.
#   keys: tireRatio, wheelRatio, forecastRatio (ints; the IN_TIRE_RATIO/IN_WHEEL_RATIO/IN_RATIO),
#         tire, wheel, valve, film, label, misc1, misc2 (component part-code strs; '' if none),
#         effectiveMonth ('' / ' ' = default, else 'yyyy/mm').

_COMPONENT_KEYS = ("tire", "wheel", "valve", "film", "label", "misc1", "misc2")


def _ratio_matches(effective_month, week_date):
    """The effective-month ratio pick (ForecastBreakdownF.pas:1151-1174 / algorithm §C.3):
      * blank / ' ' effective month  -> the DEFAULT ratio (always applies).
      * else tm = effMonth[2:4]+effMonth[5:7] (yy+mm of 'yyyy/mm') vs wm = weekDate[0:4] (yyyy); tm==wm.
    Returns (isMatch, isDefault). NB (AMBIGUITY, algorithm §C.3): the dated compare is a 4-char yymm vs
    4-char yyyy that is NOT obviously aligned, and is DEAD on live data (all 50 recipes have effmonth=' ').
    Reproduced for fidelity; exercised only by a synthetic dated-recipe fixture."""
    em = (effective_month or "").strip()
    if em == "":
        return (True, True)                              # default ratio
    tm = (effective_month[2:4] + effective_month[5:7])   # ratio yy + mm
    wm = (week_date or "")[0:4]                           # weekdate yyyy
    return (tm == wm, False)


def pick_ratio(recipes, week_date):
    """From the recipe row(s) for ONE assembly, pick the ratio set for this week's date. Mirrors the
    legacy while-loop (:1151-1176): take the default (blank effmonth) ratio; a dated row whose tm==wm
    OVERRIDES and breaks. Returns the chosen recipe dict, or None if NEITHER a default nor a matching
    dated row exists (-> bd=FALSE -> the D-Bug-1 forecast-gap alarm in the driver; legacy silently
    dropped the count). `recipes` is the list of recipe dicts for the assembly."""
    chosen = None
    for r in (recipes or []):
        is_match, is_default = _ratio_matches(r.get("effectiveMonth", ""), week_date)
        if is_default:
            chosen = r                                   # default; keep scanning for a dated override
        elif is_match:
            return r                                     # dated match wins, break (legacy :1172)
    return chosen


def explode(week_count, recipe):
    """PURE assembly->component explosion for ONE (assembly, week) (ForecastBreakdownF.pas:1186-1287 /
    algorithm §C.5-6). Given the week's assembly count + the chosen recipe, returns the list of
    (componentPartCode, componentCount) to day-spread. Reproduces EXACTLY:
        if forecastRatio and tireRatio and wheelRatio all <> 0:
            tireCount  = ((weekCount * forecastRatio) // 100) * tireRatio  // 100   # int div at each /100
            wheelCount = ((weekCount * forecastRatio) // 100) * wheelRatio // 100
        else: tireCount = wheelCount = 0          # any zero ratio -> zero (the legacy guard :1186-1188)
    Fan-out (component code must have length > 2 to count, legacy :1230 etc.):
        tire   -> tireCount
        wheel/valve/film/label/misc1/misc2 -> wheelCount   (ALL use wheelCount; faithful, algorithm §C.6)
    Returns [] if recipe is None (no BOM -> the driver raises the gap alarm). A component whose code is
    blank/<=2 chars is skipped (legacy length>2 guard). Note the SHARE-FANOUT fidelity: TEMA sends the
    FULL count on each assembly LIN of a shared BC; scaling by this assembly's own tireRatio/100 yields
    its share (e.g. [KLM]CC 40/20/40 -> 40%/20%/40% each, summing 100%)."""
    if recipe is None:
        return []
    tr = int(recipe.get("tireRatio") or 0)
    wr = int(recipe.get("wheelRatio") or 0)
    fr = int(recipe.get("forecastRatio") or 0)
    wc = int(week_count or 0)
    if fr != 0 and tr != 0 and wr != 0:
        tire_count = ((wc * fr) // 100) * tr // 100
        wheel_count = ((wc * fr) // 100) * wr // 100
    else:
        tire_count = 0
        wheel_count = 0
    out = []
    tire_code = (recipe.get("tire") or "")
    if len(tire_code.strip()) > 2:
        out.append((tire_code.strip(), tire_count))
    for key in ("wheel", "valve", "film", "label", "misc1", "misc2"):
        code = (recipe.get(key) or "")
        if len(code.strip()) > 2:
            out.append((code.strip(), wheel_count))      # ALL non-tire components use wheel_count
    return out


def day_spread(fc_count, calendar_rows, work_days=None):
    """PURE weekly-qty -> 7 day-buckets (Mon..Sun = IN_QTY1..7) (ForecastBreakdownF.pas:1321-1434 /
    algorithm §D.3-6).

    BYTE-FAITHFUL to the legacy DoPartNumberForecast day-spread, INCLUDING the running `days` counter:

      * The legacy STARTS from a fixed Mon-Fri working week (workday[1..5]=true, workday[6]=workday[7]=
        false; days := 5; :1321-1328).
      * Then, for EACH AD_GetSpecialDateWeek row (:1394-1410), it applies the row's status to that
        'Day Number' AND adjusts the running `days` counter UNCONDITIONALLY:
            H / X (HOLIDAY / NON-PRODUCTION) -> workday[N] := False; DEC(days)   (:1398-1401)
            else  (OVERTIME 'O', any non-H/X) -> workday[N] := True;  INC(days)  (:1403-1407)
        'O' (OVERTIME — e.g. a worked Saturday) turns a day ON; this is the BLOCKER-2 fix. The live
        VehicleOrder calendar carries 6 'O' Saturdays on COROLLA (ISO wks 13/15/17/20/23/30); without
        this branch an O-Saturday week divides qty by 5 instead of 6 and every bucket is wrong.
      * The remaining defined statuses (N=NORMAL, W=WEEKEND) never appear in SpecialDate data (verified
        on spike), so only H/X and O are exercised; any OTHER status falls to the else-branch (turn ON),
        matching the legacy's binary `if H or X .. else` exactly.

    PARITY OVER THE D-Bug-3 "fix": the legacy DEC/INC is a RUNNING COUNTER, NOT count(workday). It can
    diverge from |working| in two latent edges and we reproduce the legacy NUMBER in both (the review's
    job is parity, ForecastBreakdownF.pas:1403-1407 "exactly"):
      - H/X on an already-OFF day (e.g. an H Saturday) -> DEC(days) below 5 working days, even though the
        workday set already had that day off (the legacy under-counts days here -> a LARGER per-day qty).
      - 'O' on an already-ON weekday (Mon-Fri) -> INC(days) above the workday count (the legacy DOUBLE-
        COUNTS days here -> a SMALLER per-day qty). LATENT: all 6 live 'O' rows are Saturdays (start OFF),
        so the double-count is not hit on real data today; reproduced anyway for strict parity. (This
        REVERSES the former D-Bug-3 divergence, which avoided the double-count and therefore did NOT match
        the legacy on an O-Saturday — see forecast-import-algorithm.md §F.)

    Spread (:1414-1434): days = the running counter; ratiocount = fc_count // days; leftover = fc_count %
    days; each working day (in 1..7 order) gets ratiocount; the ENTIRE remainder lands on the FIRST
    working day (legacy: leftover added to the first working day then zeroed, :1428-1430). Non-working
    days get 0. If days <= 0 -> all zeros (legacy :1419-1422).

    `calendar_rows` = the AD_GetSpecialDateWeek rows for (lookupWeek, line) as a list of
    (day_number:int, status_abbr:str) tuples (1=Mon..7=Sun; the driver resolves them via
    _calendar_rows()). Pass [] for a no-special-date week. A caller that needs a different base week can
    pass `work_days` explicitly; the driver passes the default (the legacy Mon-Fri base).

    Returns a 7-list of ints [qty1..qty7]. D-Bug-4 defense: returns 0 (never None) for every day -> the
    additive upsert never sees NULL."""
    base = set((1, 2, 3, 4, 5)) if work_days is None else set(work_days)
    workday = [(d in base) for d in (1, 2, 3, 4, 5, 6, 7)]
    days = sum(1 for w in workday if w)                    # legacy: days := 5 (the Mon-Fri base count)
    for (day_no, status) in (calendar_rows or []):
        try:
            n = int(day_no)
        except (ValueError, TypeError):
            continue
        if not (1 <= n <= 7):
            continue
        abbr = (status or "").strip().upper()
        if abbr in ("H", "X"):
            workday[n - 1] = False                         # legacy :1400 workday[N] := False
            days -= 1                                       # legacy :1401 DEC(days) (unconditional)
        else:
            workday[n - 1] = True                          # legacy :1405 workday[N] := True ('O' = ON)
            days += 1                                       # legacy :1406 INC(days) (unconditional)
    fc = int(fc_count or 0)
    if days > 0:
        ratiocount = fc // days
        leftover = fc % days
    else:
        ratiocount = 0
        leftover = 0
    out = []
    for i in range(7):
        if workday[i]:
            out.append(ratiocount + leftover)            # all remainder -> first working day
            leftover = 0
        else:
            out.append(0)
    return out


def production_offset(first_week_number):
    """The holiday-lookup week offset (ForecastBreakdownF.pas:1371-1375 / algorithm §D.4): if the year's
    first production week != 1, the lookup week = rawWeek + (firstWeek - 1). Applied ONLY to the calendar
    lookup, NEVER to the stored IN_WEEK_NUMBER (D10 / the R1 catch). Returns the additive offset
    (firstWeek-1), or 0 when firstWeek is 1 or unknown."""
    try:
        fw = int(first_week_number)
    except (ValueError, TypeError):
        return 0
    return (fw - 1) if fw != 1 else 0


# =================================================================================================
# THE GATEWAY DRIVER — import_830. ONE transaction per file. Reuses the M1 DUNS guard / ledger / alarm.
# =================================================================================================
#
# IG81-COMPAT: every gateway API used here (runPrepQuery / runPrepUpdate(...,tx=) / beginTransaction /
# getLogger) is identical on 8.1.52 and 8.3 — no version guard needed. The SCHEDULED POLL host (Q11) is a
# gateway TIMER/event script on 8.1 (NOT 8.3 Event Streams) — IG83-TODO at process_forecast_dir below.

DATABASE = "Inventory_Spike"            # the Inventory rebuild connection name
CAL_DATABASE = "VehicleOrder"           # the ALC/VehicleOrder DB holding AD_GetSpecialDateWeek (cross-DB)
DEFAULT_HISTORICAL_WEEKS = 12           # legacy [INIT] HistoricalForecast default (84-day prune window)


def _hash_text(text):
    """Stable content hash = the ledger dedupe key (reused from M1 edi_inbound). CONTENT is the file
    identity: the same content re-dropped is a no-op even under a new name. Jython-2.7-safe (hashlib is in
    the stdlib on both 8.1/8.3). 64-char hex."""
    import hashlib
    return hashlib.sha256((text or "").encode("utf-8")).hexdigest()


def _resolve_site_by_duns(db, duns, tx):
    """DUNS guard (Q11) — resolve the inbound ISA delSL[4] to a site by VC_TMM_DUNS. Reused from M1
    (edi_inbound._resolve_site_by_duns); the element-index gap (which X12 element a real inbound TEMA ISA
    carries the routing DUNS on) is the same pending-golden caveat. Returns IN_SITE_ID or None."""
    if duns is None:
        return None
    rows = system.db.runPrepQuery(
        "SELECT IN_SITE_ID FROM INV_SITES WHERE VC_TMM_DUNS = ?", [str(duns).strip()], db, tx)
    if not len(rows):
        return None
    return int(rows[0]["IN_SITE_ID"])


def _read_site_forecast_config(db, siteId, tx):
    """RISK-2 — resolve the per-site forecast config from INV_SITES (the columns the M2 schema added):
      * IN_HISTORICAL_FORECAST  -> the history-prune window in WEEKS (legacy [INIT] HistoricalForecast;
        used for DELETE_ForecastInfo's @HistWeekDate = now - weeks*7). Falls back to
        DEFAULT_HISTORICAL_WEEKS (12) if NULL/missing.
      * BIT_USE_FIRST_PRODUCTION_DAY -> the per-site equivalent of fiUseFirstProductionDay
        (ForecastBreakdownF.pas:1362): when on, the holiday-lookup week is offset by firstWeek-1 (the
        stored week stays RAW — D10). Falls back to True (the legacy default-on) if NULL/missing.
    Returns (historicalWeeks:int, useFirstProductionDay:bool). Read inside the import tx (Inventory DB).
    Mirrors the supplierBySite per-site lookup so the per-site config is live, not hardcoded."""
    histWeeks = DEFAULT_HISTORICAL_WEEKS
    useFPD = True
    if siteId is None:
        return (histWeeks, useFPD)
    rows = system.db.runPrepQuery(
        "SELECT IN_HISTORICAL_FORECAST, BIT_USE_FIRST_PRODUCTION_DAY FROM INV_SITES WHERE IN_SITE_ID = ?",
        [int(siteId)], db, tx)
    if not len(rows):
        return (histWeeks, useFPD)
    r = rows[0]
    hv = r["IN_HISTORICAL_FORECAST"]
    if hv is not None and _as_str(hv).strip() != "":
        histWeeks = _as_int(hv) or DEFAULT_HISTORICAL_WEEKS
    fv = r["BIT_USE_FIRST_PRODUCTION_DAY"]
    if fv is not None and _as_str(fv).strip() != "":
        useFPD = _as_int(fv) != 0                          # BIT 1 -> True, 0 -> False
    return (histWeeks, useFPD)


def _already_processed(db, fileHash, tx):
    """Idempotency ledger check (reused M1 INV_EDI_INBOUND_LOG). Content-hash keyed; PROCESSED rows only.
    A second drop of the same 830 is a NO-OP -> proves the delete-then-accumulate is not double-applied by
    a re-poll (the file never re-runs). Returns True if already processed."""
    rows = system.db.runPrepQuery(
        "SELECT 1 FROM INV_EDI_INBOUND_LOG WHERE VC_FILE_HASH = ? AND VC_STATUS = 'PROCESSED'",
        [str(fileHash)], db, tx)
    return len(rows) > 0


def _record_processed(db, fileName, fileHash, siteId, status, detail, tx):
    """Write the processed-files ledger row INSIDE the per-file tx (reused M1). VC_EDI_TYPE = '830'."""
    system.db.runPrepUpdate(
        "INSERT INTO INV_EDI_INBOUND_LOG "
        "(VC_FILE_NAME, VC_FILE_HASH, VC_EDI_TYPE, IN_SITE_ID, VC_STATUS, VC_DETAIL, DT_PROCESSED) "
        "VALUES (?, ?, '830', ?, ?, ?, GETDATE())",
        [str(fileName), str(fileHash), siteId, status, (detail or "")[:400]], db, tx=tx)


def _write_gap_alarm(db, siteId, assembly, weekDate, weekCount, reason, tx):
    """D-Bug-1 — a forecast-gap alarm row (INV_EDI_ALARM_REJ, type '830_FORECAST_GAP') so a missing BOM
    recipe is VISIBLE, not silently dropped (legacy ForecastBreakdownF.pas:1178-1183 logged + dropped).
    The assembly goes in VC_ASSY_PART_NUMBER; the reason + week in VC_ERROR_TEXT/VC_SUMMARY. Reuses the
    M1 alarm table so the home hub surfaces it identically to a 824 reject."""
    system.db.runPrepUpdate(
        "INSERT INTO INV_EDI_ALARM_REJ "
        "(IN_SITE_ID, VC_ALARM_TYPE, VC_SUMMARY, VC_MANIFEST_NUMBER, VC_ASSY_PART_NUMBER, "
        " VC_ERROR_TEXT, IN_EIN, BIT_RESOLVED, DT_RAISED) "
        "VALUES (?, '830_FORECAST_GAP', ?, NULL, ?, ?, NULL, 0, GETDATE())",
        [siteId,
         ("830 forecast gap: assembly %s week %s qty %d" % (assembly, weekDate, int(weekCount or 0)))[:200],
         (assembly or "")[:12], (reason or "")[:200]],
        db, tx=tx)


def _write_stale_alarm(db, siteId, lastImport, days, tx):
    """Q11 8-day staleness alarm (INV_EDI_ALARM_REJ, type '830_FORECAST_STALE'): no successful 830 import
    in `days` days for this site. Raised by stale_sites()."""
    system.db.runPrepUpdate(
        "INSERT INTO INV_EDI_ALARM_REJ "
        "(IN_SITE_ID, VC_ALARM_TYPE, VC_SUMMARY, VC_MANIFEST_NUMBER, VC_ASSY_PART_NUMBER, "
        " VC_ERROR_TEXT, IN_EIN, BIT_RESOLVED, DT_RAISED) "
        "VALUES (?, '830_FORECAST_STALE', ?, NULL, NULL, ?, NULL, 0, GETDATE())",
        [siteId,
         ("830 forecast import stale: %s days since last import" % (days,))[:200],
         ("last import stamp: %s" % (lastImport or "<never>",))[:200]],
        db, tx=tx)


def _read_recipes(db, assembly, tx):
    """Read the BOM recipe row(s) for an assembly via SELECT_ForecastDetail(@AssyCode, @ForecastNotZero=1)
    (ForecastBreakdownF.pas:1126-1133). Returns a list of recipe dicts in the PURE explode's shape. Empty
    list = no recipe (the assembly's count gets the D-Bug-1 gap alarm). Uses the proc's column ALIASES
    (proc body /tmp/inv_utf8.sql:3154): 'Tire Ratio'/'Wheel Ratio'/'Forecast Ratio'/'Tire Part Number
    Code'/... — we wrap the proc with an inline EXEC so the NQ convention (one NQ per proc) is honored at
    cutover; here the driver calls it directly through the shim."""
    rows = system.db.runPrepQuery(
        "EXEC dbo.SELECT_ForecastDetail @AssyCode = ?, @ForecastNotZero = 1", [str(assembly)], db, tx)
    out = []
    for r in rows:
        out.append({
            "tireRatio": _as_int(r["Tire Ratio"]),
            "wheelRatio": _as_int(r["Wheel Ratio"]),
            "forecastRatio": _as_int(r["Forecast Ratio"]),
            "tire": _as_str(r["Tire Part Number Code"]),
            "wheel": _as_str(r["Wheel Part Number Code"]),
            "valve": _as_str(r["Valve Part Number"]),
            "film": _as_str(r["Film Part Number"]),
            "label": _as_str(r["Label Part Number"]),
            "misc1": _as_str(r["Misc1 Part Number"]),
            "misc2": _as_str(r["Misc2 Part Number"]),
            "effectiveMonth": _as_str(r["Active Date"]),
        })
    return out


def _read_part_master(db, partCode, tx):
    """Resolve a COMPONENT part's supplier/line/size via SELECT_PartsStockInfo(@PartNum)
    (ForecastBreakdownF.pas:1338-1351). Returns (supplier, line, size); line defaults to 'ALL LINES' when
    blank (legacy :1344-1347). Returns (None, 'ALL LINES', None) if the part has no master row (the
    breakdown row would then carry NULL supplier/size — the legacy behaved the same; we still write the
    row so the qty is not lost)."""
    rows = system.db.runPrepQuery(
        "EXEC dbo.SELECT_PartsStockInfo @PartNum = ?", [str(partCode)], db, tx)
    if not len(rows):
        return (None, "ALL LINES", None)
    r = rows[0]
    line = _as_str(r["Line Name"]).strip() or "ALL LINES"
    supplier = _as_str(r["Supplier Code"]).strip() or None
    size = _as_str(r["Size Code"]).strip() or None
    return (supplier, line, size)


def _read_first_week(db, prodYear, tx):
    """SELECT_FirstProductionDay(@ProdYear) -> 'First Week Number' (INT_FIRST_PRODUCTION_WEEK), used ONLY
    for the holiday-lookup offset (NEVER the stored week — D10). Returns the int first-week, or 1 if no
    row (offset 0)."""
    rows = system.db.runPrepQuery(
        "EXEC dbo.SELECT_FirstProductionDay @ProdYear = ?", [str(prodYear)], db, tx)
    if not len(rows):
        return 1
    return _as_int(rows[0]["First Week Number"]) or 1


def _calendar_rows(calDb, lookupWeek, line):
    """Resolve the production-calendar special-date rows for (week, line) via AD_GetSpecialDateWeek on the
    VehicleOrder DB (ForecastBreakdownF.pas:1383-1412 / algorithm §D.5). Returns a list of
    (day_number:int, status_abbr:str) tuples for day_spread to apply — the FULL legacy row stream, NOT a
    pre-reduced OFF set, so day_spread can reproduce the legacy's running DEC/INC counter exactly:
        H / X 'Date Status Abrv' -> that 'Day Number' OFF (DEC days)
        O (OVERTIME, e.g. a worked Saturday) / any non-H/X -> that 'Day Number' ON (INC days)  [BLOCKER-2]
    (See day_spread for the parity rationale — we reproduce the legacy INC(days) incl. the latent
    already-on double-count, REVERSING the former D-Bug-3 divergence.)

    CROSS-DATASOURCE READ — autocommit, NOT inside the import tx (this is the load-bearing detail). The
    calendar lives on a SEPARATE datasource (VehicleOrder/ALC), so a read against it is a DIFFERENT JDBC
    connection that is NOT part of the Inventory write transaction — passing the Inventory tx here would be
    wrong (the proc isn't even in that DB) and is not transactional with it anyway. The calendar is RO
    reference data (legacy used the separate ALC_Connection), so an autocommit read on CAL_DATABASE is both
    correct and what the gateway does. Returns [] on any calendar error (logged) so a calendar outage
    degrades to the default Mon-Fri spread rather than aborting the import."""
    log = system.util.getLogger("SPIKE.import_830")
    out = []
    try:
        rows = system.db.runPrepQuery(
            "EXEC dbo.AD_GetSpecialDateWeek @Week = ?, @LineName = ?",
            [int(lookupWeek), str(line)], calDb)              # NO tx — separate RO datasource (autocommit)
    except Exception as e:
        log.warn("calendar lookup failed for week=%s line=%s: %s -> default Mon-Fri"
                 % (lookupWeek, line, e))
        return out
    for r in rows:
        abbr = _as_str(r["Date Status Abrv"]).strip().upper()
        try:
            day_no = int(r["Day Number"])
        except (ValueError, TypeError):
            continue
        if 1 <= day_no <= 7:
            out.append((day_no, abbr))
    return out


def _as_int(v):
    if v is None:
        return 0
    try:
        return int(str(v).strip())
    except (ValueError, TypeError):
        return 0


def _as_str(v):
    return "" if v is None else str(v)


def import_830(text, fileName, database=None, calDatabase=None, fileHash=None, siteId=None,
               supplier=None, historicalWeeks=None, useFirstProductionDay=True, tx=None,
               emitFeed=False, emitFeedKwargs=None):
    """Process ONE 830 file body end-to-end. ONE transaction per file (the caller may pass tx; else we
    open + commit our own). Reproduces the legacy Execute flow (ForecastBreakdownF.pas:320-335) with the
    fixes (§F):

      0. parse_830 -> per-assembly weekly entries.
      1. per NON-SKIPPED assembly (an assembly WITH a recipe): DELETE_ForecastInfo FIRST
         (@WeekDate=firstWeekDate, @HistWeekDate=now-historicalWeeks*7, @PartNumber=assembly) — the
         delete-then-accumulate (§E). An assembly with NO recipe is skipped for the delete (legacy
         ScanPartnumber Skip) BUT still raises a per-week gap alarm in step 2 (D-Bug-1).
      2. per assembly x week:
           a. INSERTUPDATE_ForecastInfo (REPLACE raw assembly forecast).
           b. read the recipe; pick_ratio by week date. NO recipe / NO ratio match -> 830_FORECAST_GAP
              alarm (D-Bug-1, NOT a silent drop) and the count is not exploded (no BOM to explode by).
           c. explode -> component (code, count) pairs.
           d. per component: resolve supplier/line/size (part master); resolve the calendar special-date
              rows for the OFFSET holiday-lookup week (D10 — stored week stays RAW); day_spread; then
              INSERTUPDATE_BreakdownForecastInfo (ADDITIVE) with @WeekNumber = the RAW week.
      3. record the ledger row + stamp VC_LAST_FORECAST_IMPORT on the site (Q11). Commit.

    TWO DISTINCT SUPPLIER SOURCES (legacy-faithful — BLOCKER-1):
      * RAW INV_FORECAST_INF: the per-site FEED supplier (legacy fiSupplierCode, :1107) — passed in as
        `supplier`.
      * BREAKDOWN INV_BREAKDOWN_FC_INF: the COMPONENT's part-master supplier (:1349), resolved per
        component from SELECT_PartsStockInfo. NOT the feed supplier. This is also the additive-upsert key
        (VC_SUPPLIER_CODE, VC_PART_NUMBER, IN_WEEK_NUMBER), so it must be the component supplier or the
        delete-then-accumulate idempotency lands under the wrong key.

    Returns a summary dict: {'entries': N, 'assembliesDeleted': N, 'breakdownWrites': N, 'rawWrites': N,
    'gapAlarms': N, 'fileHash': fileHash, 'feed': <emit summary or None>}.

    POST-IMPORT FORECAST-DISTRIBUTION FEED (P6 — the trigger wiring):
      The legacy Execute IMPORTS the 830 then runs the OUTBOUND supplier emit in the SAME call
      (ForecastBreakdownF.pas:149/320 then :340-568). We reproduce that side-effect: when emitFeed=True
      AND we own the transaction (so the breakdown rows are COMMITTED before the emit reads them), call
      forecast_distribution.emit_forecast_distribution at the very end — AFTER the commit, reading the
      freshly-written INV_BREAKDOWN_FC_INF future-week buckets. The emit is a SEPARATE module/callable
      (testable independently — and matching the legacy two-phase Execute); a failure in the emit is
      LOGGED, not fatal to the (already-committed) import — the legacy emit's outer try/except likewise
      logs + returns FALSE without un-doing the import. When called WITH a caller tx (the poll path passes
      its own tx), the feed is the caller's responsibility AFTER its commit (see _process_one_830) — we do
      NOT emit mid-tx (the rows are not yet committed). emitFeedKwargs threads through to the feed driver
      (database/localFtp/runDate/outRoot/dirOverride/renderModule for the spike sandbox).

    !!! HONEST: there is NO captured golden 830 on disk (the D10 sample is gitignored client data), so
    BYTE/OFFSET parity vs a real TEMA 830 is UNPROVABLE — not claimed. The parse/explode/spread are
    faithful to the .pas algorithm + the proven breakdown table shape; the real day-spread vs a holiday
    week and the exact ISA element index are PENDING a captured golden (same stance as 856/810/997/824)."""
    db = database if database is not None else DATABASE
    calDb = calDatabase if calDatabase is not None else CAL_DATABASE
    histWeeks = DEFAULT_HISTORICAL_WEEKS if historicalWeeks is None else int(historicalWeeks)
    log = system.util.getLogger("SPIKE.import_830")
    if fileHash is None:
        fileHash = _hash_text(text)

    entries = parse_830(text)
    # RISK-1 — the delete-window anchor (the @WeekDate floor for DELETE_ForecastInfo's forward slice).
    #   Legacy: fFirstWeekDate is reassigned on the FIRST FST of EVERY LIN (ForecastBreakdownF.pas:283-287
    #   inside the per-LIN parse loop), so after parsing it holds the LAST assembly's first-week-date — an
    #   essentially arbitrary value. On real TEMA feeds every LIN starts at the same first week, so legacy
    #   == FIRST-entry-first-week == min-across-all; the three agree on live data.
    #   DELIBERATE DIVERGENCE (documented): we anchor on min(weekDate) across ALL entries/weeks instead.
    #   WHY SAFER: DELETE_ForecastInfo deletes the forward slice WHERE VC_WEEK_DATE >= @WeekDate, then the
    #   breakdown upsert is ADDITIVE. If the anchor were LATER than some assembly's earliest forecast week
    #   (which the legacy's "last LIN's first week" can be, if LINs are mis-ordered or carry different
    #   start weeks), those earlier weeks would NOT be deleted but WOULD be re-inserted -> DOUBLED qty on
    #   re-import. min(weekDate) guarantees the delete window covers every week we will re-insert, so the
    #   delete-then-accumulate idempotency holds for ANY LIN ordering / start-week mix. Equal to legacy on
    #   real data; strictly stronger against the doubling risk the anchor exists to guard.
    firstWeekDate = None
    for e in entries:
        for wk in e["weeks"]:
            wd = wk["weekDate"]
            if wd and (firstWeekDate is None or wd < firstWeekDate):
                firstWeekDate = wd                          # min yyyymmdd (lexical == chronological)
    histWeekDate = _hist_week_date(db, histWeeks, tx)

    ownTx = tx is None
    if ownTx:
        if siteId is None:
            raise ForecastImportError("import_830: siteId is required when called without a tx")
        tx = system.db.beginTransaction(db)
    summary = {"entries": len(entries), "assembliesDeleted": 0, "breakdownWrites": 0,
               "rawWrites": 0, "gapAlarms": 0, "fileHash": fileHash, "feed": None}
    try:
        # 1. DELETE FIRST — per assembly WITH a recipe (the delete resolves assembly->components via the
        #    recipe; an assembly with no recipe has nothing to delete, and gets gap alarms below).
        recipe_cache = {}
        for e in entries:
            assembly = e["assembly"]
            if not assembly:
                continue
            recipes = recipe_cache.get(assembly)
            if recipes is None:
                recipes = _read_recipes(db, assembly, tx)
                recipe_cache[assembly] = recipes
            if recipes and firstWeekDate is not None:
                system.db.runPrepUpdate(
                    "EXEC dbo.DELETE_ForecastInfo @WeekDate = ?, @HistWeekDate = ?, @PartNumber = ?",
                    [firstWeekDate, histWeekDate, assembly], db, tx=tx)
                summary["assembliesDeleted"] += 1

        # 2. re-explode + additive accumulate.
        for e in entries:
            assembly = e["assembly"]
            if not assembly:
                continue
            recipes = recipe_cache.get(assembly, [])
            for wk in e["weeks"]:
                rawWeek = wk["weekNumber"]
                weekDate = wk["weekDate"]
                weekCount = wk["weekCount"]
                rowSupplier = supplier
                # a. REPLACE raw assembly forecast.
                system.db.runPrepUpdate(
                    "EXEC dbo.INSERTUPDATE_ForecastInfo @Supplier = ?, @PartNumber = ?, @Kanban = ?, "
                    "@WeekNumber = ?, @WeekDate = ?, @Count = ?",
                    [rowSupplier, assembly, e["kanban"], rawWeek, weekDate, weekCount], db, tx=tx)
                summary["rawWrites"] += 1
                # b. pick the ratio.
                recipe = pick_ratio(recipes, weekDate)
                if recipe is None:
                    # D-Bug-1: NO silent drop — a visible forecast-gap alarm.
                    _write_gap_alarm(db, siteId, assembly, weekDate, weekCount,
                                     "no BOM recipe / ratio match for assembly on this week date", tx)
                    summary["gapAlarms"] += 1
                    continue
                # c. explode.
                components = explode(weekCount, recipe)
                # d. day-spread each component into the breakdown.
                prodYear = (weekDate or "")[0:4]
                firstWeek = _read_first_week(db, prodYear, tx) if useFirstProductionDay else 1
                lookupWeek = rawWeek + production_offset(firstWeek)
                for (compCode, compCount) in components:
                    compSupplier, line, size = _read_part_master(db, compCode, tx)
                    # BLOCKER-1 (parity): the breakdown row's supplier is the COMPONENT's part-master
                    # supplier (ForecastBreakdownF.pas:1349 supplier := FieldByName('Supplier Code') read
                    # from SELECT_PartsStockInfo(@PartNum := the COMPONENT) — proven 959/959 on live, 14
                    # distinct suppliers). The feed supplier (rowSupplier) is used ONLY for the RAW
                    # INV_FORECAST_INF write above (:1107). NOT `rowSupplier or compSupplier` — that wrote
                    # the feed supplier on every row (wrong VC_SUPPLIER_CODE + wrong additive-upsert key;
                    # and DELETE_ForecastInfo has no supplier filter, so a re-import would silently rewrite
                    # the supplier). compSupplier may be None (no part master) -> the legacy wrote NULL too.
                    rowSup = compSupplier
                    calRows = _calendar_rows(calDb, lookupWeek, line)   # cross-DS RO read (no tx)
                    qtys = day_spread(compCount, calRows)
                    system.db.runPrepUpdate(
                        "EXEC dbo.INSERTUPDATE_BreakdownForecastInfo @WeekNumber = ?, @WeekDate = ?, "
                        "@Supplier = ?, @PartNumber = ?, @SizeCode = ?, @Qty1 = ?, @Qty2 = ?, @Qty3 = ?, "
                        "@Qty4 = ?, @Qty5 = ?, @Qty6 = ?, @Qty7 = ?",
                        [rawWeek, weekDate, rowSup, compCode, size,
                         qtys[0], qtys[1], qtys[2], qtys[3], qtys[4], qtys[5], qtys[6]], db, tx=tx)
                    summary["breakdownWrites"] += 1

        # 3. ledger + last-import stamp (Q11).
        if ownTx:
            detail = ("%d entries, %d assys deleted, %d raw, %d breakdown, %d gaps"
                      % (summary["entries"], summary["assembliesDeleted"], summary["rawWrites"],
                         summary["breakdownWrites"], summary["gapAlarms"]))
            _record_processed(db, fileName, fileHash, siteId, "PROCESSED", detail, tx)
            _stamp_last_import(db, siteId, tx)
            system.db.commitTransaction(tx)
    except Exception:
        if ownTx:
            system.db.rollbackTransaction(tx)
        raise
    finally:
        if ownTx:
            system.db.closeTransaction(tx)

    # POST-IMPORT (P6): the outbound forecast-distribution emit, ONLY after our own commit (so the
    # breakdown rows the feed reads are committed). Side-effect of a successful import (legacy Execute
    # :320 import -> :340 emit). A feed failure is LOGGED, not fatal (the import is already committed; the
    # legacy emit's try/except logs + returns FALSE without un-doing the import).
    if ownTx and emitFeed:
        summary["feed"] = _emit_forecast_feed(db, emitFeedKwargs, log)

    log.info("import_830 file=%s site=%s -> %d entries, %d assys deleted, %d raw, %d breakdown, %d gaps"
             % (fileName, siteId, summary["entries"], summary["assembliesDeleted"],
                summary["rawWrites"], summary["breakdownWrites"], summary["gapAlarms"]))
    return summary


def _emit_forecast_feed(db, emitFeedKwargs, log):
    """Run the outbound forecast-distribution feed (P6) as the post-import side-effect. Imports the
    forecast_distribution project-library module at CALL time (so this module imports cleanly without it),
    and threads emitFeedKwargs (database/localFtp/runDate/weekDate/outRoot/dirOverride/renderModule). A
    failure is caught + logged (the import is already committed; the legacy emit is likewise non-fatal to
    the import). Returns the feed summary dict, or {'error': ...} on a caught failure."""
    kw = dict(emitFeedKwargs or {})
    kw.setdefault("database", db)
    try:
        # Prefer a module injected into this module's globals (the headless shim / a test injects
        # `forecast_distribution` via extra_globals so no real import path is needed); fall back to the
        # real gateway project-library import. On the gateway both resolve to the same project-library.
        _fd = globals().get("forecast_distribution")
        if _fd is None:
            import forecast_distribution as _fd       # gateway project-library module (P6)
        return _fd.emit_forecast_distribution(**kw)
    except Exception as e:
        log.error("import_830: forecast-distribution feed FAILED post-import (import IS committed): %s"
                  % (e,))
        return {"error": str(e)}


def _hist_week_date(db, histWeeks, tx):
    """The history-prune cutoff = now - histWeeks*7 days, as yyyymmdd (ForecastBreakdownF.pas:318:
    formatdatetime('yyyymmdd', now - HistoricalForecast*7)). Computed in SQL Server so the date math
    matches the gateway/DB clock (not the harness host). Returns the 8-char string."""
    rows = system.db.runPrepQuery(
        "SELECT CONVERT(varchar(8), DATEADD(day, ?, GETDATE()), 112) AS d", [-7 * int(histWeeks)], db, tx)
    return rows[0]["d"] if len(rows) else None


def _stamp_last_import(db, siteId, tx):
    """Q11 — stamp INV_SITES.VC_LAST_FORECAST_IMPORT with a 16-char yyyymmddHHmmssff audit stamp on a
    successful import (the home-hub Forecast Import box + the staleness check read it)."""
    if siteId is None:
        return
    system.db.runPrepUpdate(
        "UPDATE INV_SITES SET VC_LAST_FORECAST_IMPORT = "
        "LEFT(CONVERT(varchar(20), GETDATE(), 112) + REPLACE(CONVERT(varchar(12), GETDATE(), 114), ':', ''), 16) "
        "WHERE IN_SITE_ID = ?", [int(siteId)], db, tx=tx)


def stale_sites(database=None, staleDays=8, raiseAlarms=True, tx=None):
    """Q11 8-day staleness check (DATA side). For each site in AUTO import mode, compare
    VC_LAST_FORECAST_IMPORT (yyyymmddHHmmss...) to now; if >= staleDays days (or NEVER imported), it is
    stale. Optionally raises a 830_FORECAST_STALE alarm per stale site (deduped: only if no UNRESOLVED
    stale alarm already exists for the site). Returns the list of stale {siteId, lastImport, days}.

    The gateway SCHEDULED job calls this daily (prod wiring — IG83-TODO: a gateway timer script on 8.1,
    NOT 8.3 Event Streams). MANUAL-mode sites are excluded (operator owns the cadence)."""
    db = database if database is not None else DATABASE
    log = system.util.getLogger("SPIKE.import_830")
    ownTx = tx is None
    if ownTx:
        tx = system.db.beginTransaction(db)
    stale = []
    try:
        rows = system.db.runPrepQuery(
            "SELECT IN_SITE_ID, VC_LAST_FORECAST_IMPORT, "
            "  CASE WHEN VC_LAST_FORECAST_IMPORT IS NULL OR LEN(VC_LAST_FORECAST_IMPORT) < 8 "
            "       THEN 99999 "
            "       ELSE DATEDIFF(day, CONVERT(datetime, LEFT(VC_LAST_FORECAST_IMPORT, 8), 112), GETDATE()) "
            "  END AS daysSince "
            "FROM INV_SITES WHERE VC_FORECAST_IMPORT_MODE = 'AUTO'", [], db, tx)
        for r in rows:
            days = _as_int(r["daysSince"])
            if days >= int(staleDays):
                siteId = int(r["IN_SITE_ID"])
                lastImport = _as_str(r["VC_LAST_FORECAST_IMPORT"]) or None
                stale.append({"siteId": siteId, "lastImport": lastImport, "days": days})
                if raiseAlarms:
                    existing = system.db.runPrepQuery(
                        "SELECT 1 FROM INV_EDI_ALARM_REJ WHERE IN_SITE_ID = ? "
                        "AND VC_ALARM_TYPE = '830_FORECAST_STALE' AND BIT_RESOLVED = 0",
                        [siteId], db, tx)
                    if not len(existing):
                        _write_stale_alarm(db, siteId, lastImport, days, tx)
        if ownTx:
            system.db.commitTransaction(tx)
    except Exception:
        if ownTx:
            system.db.rollbackTransaction(tx)
        raise
    finally:
        if ownTx:
            system.db.closeTransaction(tx)
    log.info("stale_sites: %d AUTO-mode sites stale (>= %d days)" % (len(stale), staleDays))
    return stale


def process_forecast_dir(inDir, database=None, calDatabase=None, archiveDir=None, quarantineDir=None,
                         supplierBySite=None, emitFeed=False, emitFeedKwargs=None):
    """The SCHEDULED POLL entry (Q11 — replaces the legacy manual button). List inDir; for each FILE,
    envelope-check (ISA on line 1) + DUNS-guard + idempotency, then dispatch to import_830 for an '830'
    body. Mirrors the M1 edi_inbound.process_inbound shape (a non-830 / no-DUNS file is QUARANTINED, not
    silently dropped). The per-site supplier is looked up from supplierBySite (a {siteId: supplierCode}
    map the caller supplies from INV_SITES.VC_SUPPLIER_CODE).

    P6 TRIGGER: emitFeed=True runs the outbound forecast-distribution feed
    (forecast_distribution.emit_forecast_distribution) once per successfully-PROCESSED 830, AFTER that
    file's import commits — matching the legacy EDIUpload operational path (import the inbound 830 then
    emit each affected supplier's .frc/Excel; ForecastBreakdownF.pas via EDIUpload.pas:94-99). The emit is
    a non-fatal side-effect (a feed failure is logged, the import stays committed). emitFeedKwargs threads
    the feed driver's localFtp/runDate/outRoot/dirOverride/renderModule.

    IG83-TODO: on 8.3 prefer an Event Stream + system.file move/archive; this 8.1-safe shape uses a
    gateway timer script + os-level listing/rename. IG81-COMPAT: the timer script + system.db APIs are
    identical on 8.1/8.3. This is the PROD wiring stub; the headless tests drive import_830 directly."""
    db = database if database is not None else DATABASE
    calDb = calDatabase if calDatabase is not None else CAL_DATABASE
    log = system.util.getLogger("SPIKE.import_830")
    import os as _os
    if archiveDir is None:
        archiveDir = _os.path.join(inDir, "Archive")
    if quarantineDir is None:
        quarantineDir = _os.path.join(inDir, "Quarantine")
    supplierBySite = supplierBySite or {}

    results = []
    try:
        names = sorted(_os.listdir(inDir))
    except OSError as e:
        log.error("process_forecast_dir: cannot list inDir %s: %s" % (inDir, e))
        return results
    for name in names:
        srcPath = _os.path.join(inDir, name)
        if not _os.path.isfile(srcPath):
            continue
        try:
            with open(srcPath, "rb") as fh:
                text = fh.read().decode("utf-8", "replace")
        except Exception as e:
            log.error("process_forecast_dir: cannot read %s: %s" % (srcPath, e))
            results.append({"fileName": name, "outcome": "READ_ERROR", "detail": str(e)})
            continue
        try:
            r = _process_one_830(text, name, db, calDb, srcPath, archiveDir, quarantineDir,
                                 supplierBySite, emitFeed=emitFeed, emitFeedKwargs=emitFeedKwargs)
            r["fileName"] = name
            results.append(r)
        except Exception as e:
            log.error("process_forecast_dir: FAILED processing %s: %s" % (name, e))
            results.append({"fileName": name, "outcome": "ERROR", "detail": str(e)})
    return results


def _process_one_830(text, fileName, db, calDb, srcPath, archiveDir, quarantineDir, supplierBySite,
                     emitFeed=False, emitFeedKwargs=None):
    """One file, one tx: envelope + DUNS + idempotency, then import_830. Quarantine a non-830 / no-DUNS
    file (Q11). Mirrors edi_inbound.process_one_file. Returns an outcome dict.

    RISK-2: reads the per-site forecast config (IN_HISTORICAL_FORECAST + BIT_USE_FIRST_PRODUCTION_DAY) from
    INV_SITES via _read_site_forecast_config and threads it into import_830, so the poll path honors the
    per-site config instead of always defaulting (the columns the M2 schema added are LIVE here)."""
    log = system.util.getLogger("SPIKE.import_830")
    fileHash = _hash_text(text)
    lines = _lines(text)
    line1 = lines[0] if lines else ""
    if "ISA" not in line1:
        _move(srcPath, quarantineDir, log)
        return {"outcome": "QUARANTINED_NOT_EDI", "detail": "no ISA on line 1"}
    duns = _elem(_split_star(line1), 4) or None
    typ = _detect_type(lines)

    tx = system.db.beginTransaction(db)
    committed = False
    try:
        siteId = _resolve_site_by_duns(db, duns, tx)
        if siteId is None:
            _record_processed(db, fileName, fileHash, None, "QUARANTINED",
                              "no site for DUNS %r" % (duns,), tx)
            system.db.commitTransaction(tx)
            committed = True
            _move(srcPath, quarantineDir, log)
            return {"outcome": "QUARANTINED_NO_DUNS", "detail": "no site for DUNS %r" % (duns,)}
        if _already_processed(db, fileHash, tx):
            system.db.rollbackTransaction(tx)
            committed = True
            return {"outcome": "SKIPPED_ALREADY_PROCESSED", "siteId": siteId}
        if typ != "830":
            _record_processed(db, fileName, fileHash, siteId, "QUARANTINED",
                              "type %r not 830" % (typ,), tx)
            system.db.commitTransaction(tx)
            committed = True
            _move(srcPath, quarantineDir, log)
            return {"outcome": "QUARANTINED_NOT_830", "detail": "type %r" % (typ,), "siteId": siteId}
        supplier = supplierBySite.get(siteId)
        # RISK-2 — read the per-site forecast config live (mirrors the supplierBySite lookup pattern) so
        # the columns the schema added are NOT dead: IN_HISTORICAL_FORECAST (the history-prune window in
        # weeks, legacy [INIT] HistoricalForecast) and BIT_USE_FIRST_PRODUCTION_DAY (the per-site
        # equivalent of fiUseFirstProductionDay — drives the holiday-lookup offset). Without this the poll
        # path always defaulted (12 weeks / use-FPD on) regardless of the site row.
        histWeeks, useFPD = _read_site_forecast_config(db, siteId, tx)
        summary = import_830(text, fileName, database=db, calDatabase=calDb, fileHash=fileHash,
                             siteId=siteId, supplier=supplier, historicalWeeks=histWeeks,
                             useFirstProductionDay=useFPD, tx=tx)
        detail = ("%d entries, %d breakdown writes, %d gaps"
                  % (summary["entries"], summary["breakdownWrites"], summary["gapAlarms"]))
        _record_processed(db, fileName, fileHash, siteId, "PROCESSED", detail, tx)
        _stamp_last_import(db, siteId, tx)
        system.db.commitTransaction(tx)
        committed = True
    except Exception:
        if not committed:
            system.db.rollbackTransaction(tx)
        raise
    finally:
        system.db.closeTransaction(tx)
    _move(srcPath, archiveDir, log)
    # P6: outbound forecast-distribution emit AFTER this file's import committed (the feed reads the
    # freshly-committed INV_BREAKDOWN_FC_INF future-week buckets). Non-fatal side-effect (logged on error).
    feed = _emit_forecast_feed(db, emitFeedKwargs, log) if emitFeed else None
    return {"outcome": "PROCESSED", "siteId": siteId, "summary": summary, "feed": feed}


def _detect_type(lines):
    """The transaction-set id = legacy copy(fcl,4,3) on the ST segment (ForecastBreakdownF.pas:205).
    line[3:6] of 'ST*830'. Scan for the first ST segment (robust to envelope ordering)."""
    for ln in lines:
        if ln[0:2] == "ST" and (len(ln) < 3 or ln[2] in ("*", " ", "\t")):
            return ln[3:6]
    return ""


def _move(srcPath, destDir, log):
    """Move a file aside (archive on success / quarantine on reject). Best-effort; the DB ledger is the
    authoritative idempotency guard so a move failure never risks a double-process. IG83-TODO: prefer
    system.file move on 8.3."""
    if srcPath is None or destDir is None:
        return None
    import os as _os
    try:
        if not _os.path.isdir(destDir):
            _os.makedirs(destDir)
        dest = _os.path.join(destDir, _os.path.basename(srcPath))
        _os.rename(srcPath, dest)
        return dest
    except Exception as e:
        log.warn("move: failed %s -> %s: %s" % (srcPath, destDir, e))
        return None
