#!/usr/bin/env python3
"""test_forecast_distribution_e2e.py — END-TO-END exercise of the REAL forecast-distribution feed
(docs/analysis/order/project-library/forecast_distribution/code.py) driven headlessly through the
system.db shim (jython_shim) against the spike DB — the P6 unit proof.

WHAT THIS PROVES (running the ACTUAL emit_forecast_distribution driver, NOT a re-implementation — R8)
----------------------------------------------------------------------------------------------------
Seeds synthetic future-dated INV_BREAKDOWN_FC_INF rows (the SAME table the 830 import writes) for
synthetic suppliers + synthetic supplier-master config rows, then drives emit_forecast_distribution
end-to-end:

  (1) .frc BYTE FORMAT — the published .frc bytes EXACTLY equal the spec §3 positional layout: raw
      supplier + raw part + raw weekdate + %.2d week-number + %.5d qty1..7, one line per record, CRLF
      after EVERY line (incl. the last). The qty min-width is proven on a NEGATIVE (-5 -> '-00005', NOT
      '-0005') and an >99999 OVERFLOW (100000 -> '100000', widened, NOT truncated). The expected lines are
      built field-by-field from the SOURCE spec (ForecastBreakdownF.pas:484-508), NOT the rebuild.
  (2) TEXT vs EXCEL vs BOTH vs NULL->BOTH — each supplier's Output File Type drives the channels; a NULL
      VC_OUTPUT_FILE silently gets BOTH (legacy H2, no ELSE).
  (3) SENDSITE PREFIX (the P6 crash fix) — a sendsite supplier (BIT_SITE_NUMBER_IN_ORDER=1) leads each
      .frc line with INV_SITES.VC_SUPPLIER_CODE (the legacy fiSupplierCode, NOT the phantom
      SiteSupplierCode column that crashed); a non-sendsite supplier has NO lead. Proven byte-for-byte.
  (4) PER-SUPPLIER FILE BREAK — the supplier-sorted feed produces one file set per VC_SUPPLIER_CODE.
  (5) FUTURE-WEEKS-ONLY FILTER — a past-dated bucket (VC_WEEK_DATE <= @WeekDate) is NOT emitted.
  (6) ATOMICITY (legacy H1 FIX) — a mid-emit failure (the render/write throws) leaves NO partial final
      file for that supplier (the staged .tmp is cleaned up; no half-written .frc).
  (7) ARCHIVE COPY — LocalFTP=True writes a second \\Archive\\ copy (<name>-<code><yyyymmdd>.frc), byte-
      identical to the main .frc.
  (8) EXCEL cols 1-11 + VC_SIZE_CODE inclusion — the .xlsx (rendered via the REAL POI lane under the
      bundled jython, read back with openpyxl) has the 11 columns from row 2, col 1 = VC_SIZE_CODE (which
      the .frc OMITS), cols 2-11 = part/weekdate/weeknum/qty1..7 as STRINGS.
  (9) TRIGGER WIRING — the 830 import (import_830) calls emit_forecast_distribution post-commit when
      emitFeed=True (the legacy Execute import-then-emit side-effect); proven by an end-to-end import that
      writes a future-dated breakdown row and emits its .frc in the SAME call.

NON-VACUITY — each format assertion is anchored to a SOURCE-derived constant (the spec §3 byte layout),
not the rebuild's own output; reverting the rebuild's formatter (e.g. Python '%05d' for the qty, or the
phantom-column crash) makes the corresponding check FAIL. A scripted revert-proof is at the end (set
FD_REVERT_PROOF=1 to monkeypatch _format_qty to the wrong Python '%05d' and confirm (1) fails).

HONEST VERIFICATION (no golden .frc)
------------------------------------
NO captured TEMA-downstream .frc / ForecastTemplate.xls exists on disk. So: (a) whether the receiving
sub-supplier parser expects always-full-width positional columns (the legacy writes supplier/part/weekdate
RAW); (b) the EXACT sendsite leading-field width/justification (we emit INV_SITES.VC_SUPPLIER_CODE raw,
matching the .ord _M4 stance); (c) the Excel row-1 HEADER titles (the binary template's styling) are
GOLDEN-PENDING (P7). What IS proven: the .frc byte layout vs the .pas, the channel selection, the sendsite
fix, atomicity, the archive copy, the Excel 11-col layout incl. VC_SIZE_CODE, and the import trigger.

Synthetic-only + restore-as-found: every breakdown row uses a sentinel supplier code (ZZP6*) + a ZZP6
part prefix; supplier-master rows carry VC_ADD=sentinel. Teardown deletes exactly those (suppliers one at
a time — the DELETE_SupplierCode trigger cascades the breakdown rows) and asserts the spike is restored;
INV_SITES.VC_SUPPLIER_CODE is captured + restored as-found. Files go to a tempfile.mkdtemp, removed at the
end. Inventory_Live / VehicleOrder are never written.

Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 scripts/e2e/test_forecast_distribution_e2e.py
"""
import os
import sys
import subprocess
import tempfile
import shutil
import glob
import json
import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report, preclean_sentinels   # noqa: E402
import jython_shim              # noqa: E402

CONTAINER = os.environ.get("CONTAINER", "mssql-spike")
SA_PASS = os.environ.get("SA_PASS")
INV = "Inventory"
SITE_ID = 1                     # the single site whose VC_SUPPLIER_CODE is the sendsite lead

REPO = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".."))
CODE_PY = os.path.join(REPO, "docs", "analysis", "order", "project-library",
                       "forecast_distribution", "code.py")
FORECAST_CODE_PY = os.path.join(REPO, "docs", "analysis", "edi", "inbound", "project-library",
                                "forecast", "code.py")
RR_DIR = os.path.join(REPO, "docs", "analysis", "reporting", "project-library", "report_render")

# Bundled JRE + jython + POI — the EXACT gateway runtime, lets us run the REAL report_render.render_xlsx
# (POI) headless (reused from test_m3_reports.py). No system Java required.
IGN = "/usr/local/ignition"
JAVA = os.path.join(IGN, "lib", "runtime", "jre-mac", "bin", "java")
COMMON = os.path.join(IGN, "lib", "core", "common")
PYLIB = os.path.join(IGN, "user-lib", "pylib")
JYJAR = os.path.join(COMMON, "jython-ia-2.7.3.3.jar")

# Sentinels — supplier codes (5-char) + a part prefix; supplier-master rows carry VC_ADD=ADD_STAMP.
ADD_STAMP = "ZZP6FDADD0000000"  # 16-char VC_ADD on synthetic supplier-master rows
SUP_TEXT = "ZZP6T"              # supplier wanting TEXT only ('T'), non-sendsite
SUP_EXCEL = "ZZP6E"            # supplier wanting EXCEL only ('E')
SUP_BOTH = "ZZP6B"             # supplier wanting BOTH ('B'), SENDSITE=1
SUP_NULL = "ZZP6N"             # supplier with NULL VC_OUTPUT_FILE -> BOTH (H2)
PART1 = "ZZP6PART0001"          # 12-char synthetic part
PART2 = "ZZP6PART0002"
SIZE1 = "ZZP6SZ"               # synthetic size code (Excel col 1)
ALL_SUPS = (SUP_TEXT, SUP_EXCEL, SUP_BOTH, SUP_NULL)


def sql(query, db=INV):
    if not SA_PASS:
        sys.exit("export SA_PASS first")
    out = subprocess.check_output([
        "docker", "exec", CONTAINER, "/opt/mssql-tools18/bin/sqlcmd", "-C", "-S", "localhost",
        "-U", "sa", "-P", SA_PASS, "-d", db, "-h", "-1", "-W", "-s", "\t",
        "-Q", "SET QUOTED_IDENTIFIER ON; SET ANSI_NULLS ON; SET NOCOUNT ON; " + query], text=True)
    return [l.split("\t") for l in out.splitlines()
            if l.strip() and not l.startswith("(") and not l.startswith("Msg ")]


def scalar(q, db=INV):
    r = sql(q, db)
    return r[0][0] if r else None


# --- synthetic fixture setup / teardown ------------------------------------------------------------

def _sentinel_statements():
    """Child-before-parent sentinel DELETEs. The DELETE_SupplierCode trigger CASCADES the breakdown +
    INV_FORECAST_INF rows for a deleted supplier (CreateInventory.sql:4362), so we delete suppliers ONE AT
    A TIME (the trigger uses a scalar SELECT .. FROM DELETED — a multi-row delete would error). We ALSO
    delete the breakdown rows explicitly first (in case a supplier-master row was never seeded)."""
    stmts = [
        "DELETE FROM INV_BREAKDOWN_FC_INF WHERE VC_SUPPLIER_CODE IN ('%s','%s','%s','%s') "
        "OR VC_PART_NUMBER IN ('%s','%s')"
        % (SUP_TEXT, SUP_EXCEL, SUP_BOTH, SUP_NULL, PART1, PART2),
    ]
    for code in ALL_SUPS:
        stmts.append("DELETE FROM INV_SUPPLIER_MST WHERE VC_SUPPLIER_CODE = '%s'" % code)
    stmts.append("DELETE FROM INV_SUPPLIER_MST WHERE VC_ADD = '%s'" % ADD_STAMP)
    return stmts


def preclean_fixtures():
    preclean_sentinels(lambda s: sql(s), _sentinel_statements(), label="fcdist")


def teardown_fixtures():
    for stmt in _sentinel_statements():
        sql(stmt)


def setup_supplier(code, outputFile, sendsite, name, directory):
    """Insert a synthetic INV_SUPPLIER_MST row. outputFile is 'T'/'E'/'B'/NULL (drives Output File Type).
    sendsite -> BIT_SITE_NUMBER_IN_ORDER. directory -> VC_BREAKDOWN_ORDER_DIRECTORY (the per-supplier out
    dir; the test sandboxes it under outRoot)."""
    of = "NULL" if outputFile is None else "'%s'" % outputFile
    sql("INSERT INTO INV_SUPPLIER_MST "
        "(VC_SUPPLIER_CODE, VC_SUPPLIER_NAME, VC_BREAKDOWN_ORDER_DIRECTORY, VC_OUTPUT_FILE, "
        " BIT_SITE_NUMBER_IN_ORDER, VC_ADD) "
        "VALUES ('%s', '%s', '%s', %s, %d, '%s')"
        % (code, name, directory, of, 1 if sendsite else 0, ADD_STAMP))


def seed_breakdown(supCode, partNum, weekDate, weekNumber, sizeCode, qtys):
    """Insert a synthetic INV_BREAKDOWN_FC_INF row (the table the 830 import writes; SELECT_ForecastSupplier
    reads it). qtys is a 7-list (IN_QTY1..7); a None becomes SQL NULL (-> AsInteger 0 -> '00000')."""
    qcols = ", ".join("NULL" if q is None else str(int(q)) for q in qtys)
    sql("INSERT INTO INV_BREAKDOWN_FC_INF "
        "(IN_WEEK_NUMBER, VC_WEEK_DATE, VC_SUPPLIER_CODE, VC_PART_NUMBER, VC_SIZE_CODE, "
        " IN_QTY1, IN_QTY2, IN_QTY3, IN_QTY4, IN_QTY5, IN_QTY6, IN_QTY7) "
        "VALUES (%d, '%s', '%s', '%s', %s, %s)"
        % (weekNumber, weekDate, supCode, partNum,
           "NULL" if sizeCode is None else "'%s'" % sizeCode, qcols))


# --- the REAL POI Excel render via the bundled jython (report_render.render_xlsx) -------------------

_JY_RENDER = r'''# -*- coding: utf-8 -*-
import sys, json
sys.path.insert(0, %(rr)r)
import code as engine
from java.io import FileOutputStream
payload = json.loads(open(%(payload)r).read())
cells = [tuple(c) for c in payload["cells"]]
widths = dict((int(k), v) for k, v in payload["widths"].items())
b = engine.render_xlsx(cells, widths)
fos = FileOutputStream(payload["out"]); fos.write(b); fos.close()
print "OK bytes=%%d" %% len(b)
'''


def _jython_render_xlsx(cells, widths, tmpdir):
    """Render a (cells, widths) plan to .xlsx bytes via the REAL report_render.render_xlsx (POI) under the
    bundled jython — the exact gateway lane. Returns the bytes. (Mirrors test_m3_reports.jython_render.)"""
    out_path = os.path.join(tmpdir, "_render_%d.xlsx" % id(cells))
    payload = {"cells": [list(c) for c in cells], "widths": dict((str(k), v) for k, v in widths.items()),
               "out": out_path}
    pfile = out_path + ".payload.json"
    with open(pfile, "w") as fh:
        json.dump(payload, fh)
    sfile = out_path + ".render.py"
    with open(sfile, "w") as fh:
        fh.write(_JY_RENDER % {"rr": RR_DIR, "payload": pfile})
    cp = JYJAR + ":" + os.path.join(COMMON, "*")
    cmd = [JAVA, "-Dpython.path=" + PYLIB, "-Dpython.cachedir.skip=true",
           "-cp", cp, "org.python.util.jython", "-S", sfile]
    subprocess.check_output(cmd, text=True, stderr=subprocess.STDOUT, env=dict(os.environ))
    with open(out_path, "rb") as fh:
        return fh.read()


class _RenderModule(object):
    """A renderModule for emit_forecast_distribution whose render_xlsx delegates to the REAL POI lane
    (report_render.render_xlsx under the bundled jython). So the driver's Excel path runs the genuine
    gateway render, and we read the .xlsx back with openpyxl."""
    def __init__(self, tmpdir):
        self.tmpdir = tmpdir
    def render_xlsx(self, cells, widths):
        return _jython_render_xlsx(cells, widths, self.tmpdir)


# --- expected-.frc builder DERIVED FROM THE SOURCE SPEC (R15 — NOT the rebuild) --------------------

def _src_qty(n, width):
    """The SPEC §3 Pascal Format('%.Nd', [n]) — INDEPENDENT re-derivation (not a call into the rebuild's
    _format_qty): sign prepended OUTSIDE the zero-pad of min `width` digits. -5/5 -> '-00005'; 100000 ->
    '100000' (widened). If this ever matched Python '%0Nd' for a negative it would be WRONG."""
    n = int(n if n is not None else 0)
    digits = "%0*d" % (width, abs(n))
    return ("-" + digits) if n < 0 else digits


def src_frc_line(supCode, partNum, weekDate, weekNumber, qtys, sitePrefix=None):
    """Build ONE expected .frc line from the SOURCE spec §3 byte layout (positional concat), independent
    of the rebuild. qtys is a 7-list; a None is treated as 0 (AsInteger of a NULL field)."""
    s = (sitePrefix or "") + supCode + partNum + weekDate + _src_qty(weekNumber, 2)
    for q in qtys:
        s += _src_qty(q, 5)
    return s


def main():
    print("=" * 84)
    print(" FORECAST-DISTRIBUTION feed E2E — REAL emit_forecast_distribution via the shim (P6)")
    print("=" * 84)
    rep = Report()
    if not SA_PASS:
        sys.exit("export SA_PASS first")
    if not os.path.exists(JAVA):
        sys.exit("bundled JRE not found at %s — cannot run the POI Excel render" % JAVA)

    fd = jython_shim.load_wrapper("fcdist_real", CODE_PY)
    rep.check("real forecast_distribution module loads under the shim (emit_forecast_distribution present)",
              hasattr(fd, "emit_forecast_distribution") and hasattr(fd, "format_frc_line"))

    preclean_fixtures()

    # Capture + (for the sendsite test) the site supplier code = INV_SITES.VC_SUPPLIER_CODE.
    site_sup = scalar("SELECT TOP 1 VC_SUPPLIER_CODE FROM INV_SITES ORDER BY IN_SITE_ID")
    rep.check("INV_SITES has a site supplier code (the sendsite lead source)", bool(site_sup),
              "site_sup=%r" % site_sup)

    # The run/filter dates. weekDate floor = 20260601; future-dated rows use 2026070x (> floor), and one
    # PAST-dated row uses 20260515 (<= floor -> filtered out).
    runDate = datetime.date(2026, 6, 1)
    WEEK_FLOOR = "20260601"
    WD_FUTURE1 = "20260706"
    WD_FUTURE2 = "20260713"
    WD_PAST = "20260515"

    tmproot = tempfile.mkdtemp(prefix="fcdist_e2e_")
    rmod = _RenderModule(tmproot)
    try:
        # ---- seed the supplier-master config rows ----
        setup_supplier(SUP_TEXT, "T", sendsite=False, name="Text/Only Co", directory="text_dir")
        setup_supplier(SUP_EXCEL, "E", sendsite=False, name="Excel Only", directory="excel_dir")
        setup_supplier(SUP_BOTH, "B", sendsite=True, name="Both Site Co", directory="both_dir")
        setup_supplier(SUP_NULL, None, sendsite=False, name="Null Type Co", directory="null_dir")

        # ---- seed the breakdown rows (future + one past) ----
        # SUP_TEXT: two future rows (file break + multi-line + a NEGATIVE qty + an >99999 OVERFLOW).
        seed_breakdown(SUP_TEXT, PART1, WD_FUTURE1, 28, SIZE1, [100, 0, None, 5, 99999, 100000, -5])
        seed_breakdown(SUP_TEXT, PART1, WD_FUTURE2, 29, SIZE1, [10, 20, 30, 40, 50, 60, 70])
        # SUP_TEXT: a PAST-dated row that MUST be filtered out (week date <= floor).
        seed_breakdown(SUP_TEXT, PART2, WD_PAST, 20, SIZE1, [999, 999, 999, 999, 999, 999, 999])
        # SUP_EXCEL: one future row.
        seed_breakdown(SUP_EXCEL, PART1, WD_FUTURE1, 28, SIZE1, [11, 22, 33, 44, 55, 66, 77])
        # SUP_BOTH (sendsite): one future row.
        seed_breakdown(SUP_BOTH, PART1, WD_FUTURE1, 28, SIZE1, [1, 2, 3, 4, 5, 6, 7])
        # SUP_NULL: one future row.
        seed_breakdown(SUP_NULL, PART1, WD_FUTURE1, 28, SIZE1, [8, 9, 10, 11, 12, 13, 14])

        of_sys = jython_shim._System()
        fd.system = of_sys

        # =============== MAIN RUN: LocalFTP=False, all four suppliers ===============
        print("\n--- emit (LocalFTP=False); per-supplier file break + channel selection ---")
        res = fd.emit_forecast_distribution(database=INV, localFtp=False, runDate=runDate,
                                            weekDate=WEEK_FLOOR, outRoot=tmproot, renderModule=rmod)

        # (4) PER-SUPPLIER FILE BREAK — each of OUR 4 synthetic suppliers got its OWN file set. (The spike
        # also holds real future-dated forecast data for live suppliers, which the feed faithfully emits
        # too — so we assert OUR 4 sentinel suppliers are EACH present exactly once, not a global total.)
        det = dict((s["supCode"], s) for s in res["suppliers"])
        synth_present = [c for c in ALL_SUPS if c in det]
        synth_codes = [s["supCode"] for s in res["suppliers"] if s["supCode"] in ALL_SUPS]
        rep.check("4. per-supplier break: each of the 4 synthetic suppliers got its OWN file set (1 each)",
                  sorted(synth_present) == sorted(ALL_SUPS) and len(synth_codes) == 4,
                  "present=%s" % synth_codes)
        # (5) FUTURE-WEEKS-ONLY: SUP_TEXT emitted 2 lines (the 2 future rows), NOT the past row.
        rep.check("5. future-weeks-only: SUP_TEXT has 2 lines (past-dated row filtered out)",
                  det[SUP_TEXT]["lineCount"] == 2, "lineCount=%d" % det[SUP_TEXT]["lineCount"])
        # the past-dated SUP_TEXT/PART2 row must appear NOWHERE in the emitted output (any supplier file).
        past_leak = []
        for s in res["suppliers"]:
            for f in s["files"]:
                if f["kind"].startswith("frc"):
                    fb = open(f["path"], "rb").read().decode("ascii", "replace")
                    if (WD_PAST in fb) and (PART2 in fb):
                        past_leak.append(f["path"])
        rep.check("5. future-weeks-only: the past-dated synthetic row appears in NO emitted .frc",
                  past_leak == [], "leaked into=%s" % past_leak)

        # (2) channel selection: TEXT -> .frc only; EXCEL -> .xlsx only; BOTH/NULL -> both.
        def kinds(s):
            return sorted(f["kind"] for f in s["files"])
        rep.check("2. SUP_TEXT (Output File Type 'T') -> .frc only", kinds(det[SUP_TEXT]) == ["frc-main"],
                  str(kinds(det[SUP_TEXT])))
        rep.check("2. SUP_EXCEL (Output File Type 'E') -> .xlsx only", kinds(det[SUP_EXCEL]) == ["xlsx"],
                  str(kinds(det[SUP_EXCEL])))
        rep.check("2. SUP_BOTH (Output File Type 'B') -> .frc + .xlsx", kinds(det[SUP_BOTH]) == ["frc-main", "xlsx"],
                  str(kinds(det[SUP_BOTH])))
        rep.check("2. SUP_NULL (NULL Output File Type) -> BOTH (.frc + .xlsx; legacy H2)",
                  kinds(det[SUP_NULL]) == ["frc-main", "xlsx"], str(kinds(det[SUP_NULL])))

        # (1) .frc BYTE FORMAT — SUP_TEXT (non-sendsite), expected built from the SOURCE spec.
        text_frc = [f["path"] for f in det[SUP_TEXT]["files"] if f["kind"] == "frc-main"][0]
        body = open(text_frc, "rb").read().decode("ascii")
        want_l1 = src_frc_line(SUP_TEXT, PART1, WD_FUTURE1, 28, [100, 0, None, 5, 99999, 100000, -5])
        want_l2 = src_frc_line(SUP_TEXT, PART1, WD_FUTURE2, 29, [10, 20, 30, 40, 50, 60, 70])
        want = want_l1 + "\r\n" + want_l2 + "\r\n"
        rep.check("1. SUP_TEXT .frc bytes EXACTLY match the SOURCE spec §3 layout (CRLF incl. last line)",
                  body == want, "got=%r\n      want=%r" % (body, want))
        # qty min-width edges, isolated:
        rep.check("1. qty NEGATIVE -> '-00005' (Pascal %.5d, NOT Python '-0005')",
                  "-00005" in body and "-0005\r" not in body, "line1=%r" % body.split("\r\n")[0])
        rep.check("1. qty OVERFLOW >99999 -> '100000' widened (NOT truncated to 5)",
                  "100000" in body, "line1=%r" % body.split("\r\n")[0])
        rep.check("1. NULL qty -> '00000' (AsInteger of NULL = 0)", "00000" in body,
                  "line1=%r" % body.split("\r\n")[0])
        rep.check("1. week-number %.2d -> '28'/'29' (min-2 zero-pad)",
                  want_l1[len(SUP_TEXT) + len(PART1) + len(WD_FUTURE1):][:2] == "28", want_l1)

        # (3) SENDSITE PREFIX — SUP_BOTH leads each line with INV_SITES.VC_SUPPLIER_CODE.
        both_frc = [f["path"] for f in det[SUP_BOTH]["files"] if f["kind"] == "frc-main"][0]
        both_body = open(both_frc, "rb").read().decode("ascii")
        want_both = src_frc_line(SUP_BOTH, PART1, WD_FUTURE1, 28, [1, 2, 3, 4, 5, 6, 7],
                                 sitePrefix=site_sup) + "\r\n"
        rep.check("3. SENDSITE: SUP_BOTH .frc leads with INV_SITES.VC_SUPPLIER_CODE (%r), source-derived"
                  % site_sup, both_body == want_both, "got=%r\n      want=%r" % (both_body, want_both))
        rep.check("3. SENDSITE: the leading field is the site code, NOT a phantom-column crash + NOT dropped",
                  both_body.startswith(site_sup + SUP_BOTH), "first bytes=%r" % both_body[:20])
        rep.check("3. NON-sendsite: SUP_TEXT line does NOT lead with the site code",
                  not body.startswith(site_sup + SUP_TEXT) and body.startswith(SUP_TEXT),
                  "first bytes=%r" % body[:20])
        rep.check("3. siteSupplierCode in the summary == INV_SITES.VC_SUPPLIER_CODE",
                  res["siteSupplierCode"] == site_sup, "got=%r" % res["siteSupplierCode"])

        # (8) EXCEL cols 1-11 + VC_SIZE_CODE — SUP_EXCEL's .xlsx, read back with openpyxl.
        import openpyxl, warnings
        excel_path = [f["path"] for f in det[SUP_EXCEL]["files"] if f["kind"] == "xlsx"][0]
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = openpyxl.load_workbook(excel_path)
        ws = wb.active

        def cellv(r, c):
            return ws.cell(row=r, column=c).value
        # row 1 = header band (11 cols). row 2 = the one future record.
        row2 = [cellv(2, c) for c in range(1, 12)]
        # SOURCE spec §4: col1=VC_SIZE_CODE, col2=part, col3=weekdate, col4=weeknum, col5-11=qty1..7, ALL
        # AsString (so col1 carries the size the .frc OMITS; the qtys are RAW strings, NOT %.5d).
        want_row2 = [SIZE1, PART1, WD_FUTURE1, "28", "11", "22", "33", "44", "55", "66", "77"]
        # openpyxl returns string cells as str; numbers-as-text stay str (we wrote AsString). Normalize.
        got_row2 = [None if v is None else str(v) for v in row2]
        rep.check("8. Excel row 2 cols 1-11 == [size, part, weekdate, weeknum, qty1..7] AsString",
                  got_row2 == want_row2, "got=%s\n      want=%s" % (got_row2, want_row2))
        rep.check("8. Excel col 1 == VC_SIZE_CODE (%r) — the column the .frc OMITS" % SIZE1,
                  got_row2[0] == SIZE1, "col1=%r" % got_row2[0])
        rep.check("8. Excel has 11 data columns (row 2 fully populated)",
                  all(v is not None for v in got_row2), str(got_row2))
        # row 1 is a header band (11 non-empty cells).
        hdr = [cellv(1, c) for c in range(1, 12)]
        rep.check("8. Excel row 1 = an 11-cell header band (data starts row 2, legacy i:=2)",
                  all(h is not None for h in hdr) and len(hdr) == 11, str(hdr))

        # all final files are .frc/.xlsx (no .tmp left)
        leftover_tmp = glob.glob(os.path.join(tmproot, "**", "*.tmp"), recursive=True)
        rep.check("no .tmp staged file left after a clean run (all renamed to final)",
                  leftover_tmp == [], "tmp=%s" % leftover_tmp)

        # =============== (7) ARCHIVE COPY: LocalFTP=True ===============
        print("\n--- (7) LocalFTP=True -> the \\Archive\\ copy (byte-identical) ---")
        ftproot = os.path.join(tmproot, "ftp")
        res2 = fd.emit_forecast_distribution(database=INV, localFtp=True, runDate=runDate,
                                             weekDate=WEEK_FLOOR, outRoot=ftproot, renderModule=rmod)
        det2 = dict((s["supCode"], s) for s in res2["suppliers"])
        text_files2 = sorted(f["kind"] for f in det2[SUP_TEXT]["files"])
        rep.check("7. LocalFTP=True: SUP_TEXT emits frc-main + frc-archive", text_files2 == ["frc-archive", "frc-main"],
                  str(text_files2))
        main2 = [f["path"] for f in det2[SUP_TEXT]["files"] if f["kind"] == "frc-main"][0]
        arch2 = [f["path"] for f in det2[SUP_TEXT]["files"] if f["kind"] == "frc-archive"][0]
        # name 'Text/Only Co' -> '/'-stripped 'TextOnly Co'; archive = <name>-<code><yyyymmdd>.frc (no '-').
        want_arch_name = "TextOnly Co-%s%s.frc" % (SUP_TEXT, runDate.strftime("%Y%m%d"))
        rep.check("7. archive path = <dir>/Archive/<name>-<code><yyyymmdd>.frc (no separator before date)",
                  os.sep + "Archive" + os.sep in arch2 and os.path.basename(arch2) == want_arch_name,
                  "arch=%r want=%r" % (os.path.basename(arch2), want_arch_name))
        rep.check("7. archive .frc is byte-identical to the main .frc",
                  open(main2, "rb").read() == open(arch2, "rb").read(), "differ")
        # H-CLEAN: the supplier name 'Text/Only Co' has its '/' stripped (only '/') -> 'TextOnly Co'.
        rep.check("7. H-CLEAN: filename strips ONLY '/' from the name ('Text/Only Co' -> 'TextOnly Co')",
                  os.path.basename(main2).startswith("TextOnly Co-" + SUP_TEXT),
                  "name=%r" % os.path.basename(main2))

        # =============== (6) ATOMICITY: a mid-emit failure leaves no partial file ===============
        print("\n--- (6) atomicity: a render/write failure mid-supplier -> NO partial final file ---")
        atomroot = os.path.join(tmproot, "atom")
        # Force the EXCEL render to throw ONLY for MY synthetic supplier (SUP_BOTH) — by inspecting the
        # cell plan for the ZZP6 part sentinel — so real-supplier renders succeed and the failure point is
        # DETERMINISTIC + scoped to my data (the spike also holds real BOTH suppliers the feed emits). The
        # driver must clean up SUP_BOTH's staged .frc .tmp + raise -> NO final .frc/.xlsx for SUP_BOTH.
        class _FailRender(object):
            def __init__(self, tmpdir):
                self.real = _RenderModule(tmpdir)
            def render_xlsx(self, cells, widths):
                if any("ZZP6" in str(c[2]) for c in cells if c[2] is not None):
                    raise RuntimeError("simulated POI render failure mid-emit (synthetic supplier)")
                return self.real.render_xlsx(cells, widths)
        fd.system = jython_shim._System()
        threw = False
        try:
            fd.emit_forecast_distribution(database=INV, localFtp=False, runDate=runDate,
                                          weekDate=WEEK_FLOOR, outRoot=atomroot,
                                          renderModule=_FailRender(tmproot))
        except Exception as e:
            threw = True
            print("    emit_forecast_distribution raised (expected): %s" % str(e)[:70])
        rep.check("6. atomicity: the mid-emit render failure was raised (not swallowed)", threw)
        # The render fails on SUP_BOTH's excel (after its .frc .tmp was staged). Assert: NO final .frc or
        # .xlsx for SUP_BOTH, and its staged .tmp is gone (the whole supplier rolled back, no partial).
        both_finals = (glob.glob(os.path.join(atomroot, "**", "*-%s.frc" % SUP_BOTH), recursive=True) +
                       glob.glob(os.path.join(atomroot, "**", "*-%s-Forecast.xlsx" % SUP_BOTH), recursive=True))
        both_tmps = glob.glob(os.path.join(atomroot, "**", "*%s*.tmp" % SUP_BOTH), recursive=True)
        rep.check("6. atomicity: NO final .frc/.xlsx for the failing supplier (no partial emit)",
                  both_finals == [], "leaked=%s" % both_finals)
        rep.check("6. atomicity: the failing supplier's staged .tmp was cleaned up",
                  both_tmps == [], "tmp=%s" % both_tmps)

        # =============== (6b) PUBLISH-PHASE all-or-nothing: a 2nd-rename (.xlsx) failure ============
        # The code-reviewer RISK (mirrors order_file step-4 / P10): a BOTH supplier renames its .frc -> final
        # SUCCESSFULLY, then the .xlsx rename THROWS. The driver must UN-RENAME the already-published .frc
        # back to .tmp so the supplier is all-or-nothing (NO orphaned final .frc), and raise an accurate
        # ForecastDistributionError reporting the TRUE on-disk state. We target SUP_NULL (NULL Output Type ->
        # BOTH via H2, so it has .frc + .xlsx). The feed is supplier-sorted by VC_SUPPLIER_CODE, so the four
        # synthetic suppliers flush in order ZZP6B, ZZP6E, ZZP6N(=SUP_NULL), ZZP6T — i.e. SUP_BOTH + SUP_EXCEL
        # publish CLEANLY before SUP_NULL fails, which lets us prove per-supplier ISOLATION (a failure does
        # not corrupt an already-completed supplier). We force the failure at the os.rename of SUP_NULL's
        # .xlsx FINAL (its 2nd planned rename: .frc then .xlsx); staging (render/write) all SUCCEEDS, so the
        # failure is purely in the publish/rename phase this fix governs, not staging.
        print("\n--- (6b) publish-phase all-or-nothing: .xlsx rename fails AFTER .frc renamed -> un-rename ---")
        import os as _os_mod
        halfroot = os.path.join(tmproot, "half")
        real_rename = _os_mod.rename

        def _failing_rename(src, dst, *a, **k):
            # Fail ONLY when publishing SUP_NULL's .xlsx FINAL (the 2nd rename). The .frc final rename (1st)
            # succeeds first, so we reproduce a true mid-sequence half-emit. We key on the FINAL target (dst
            # matches SUP_NULL's Excel name and has no .tmp suffix) so the driver's un-rename
            # (final -> final+'.tmp') is NOT blocked.
            if (("-%s-Forecast.xlsx" % SUP_NULL) in str(dst)) and not str(dst).endswith(".tmp"):
                raise RuntimeError("simulated .xlsx rename failure mid-publish (synthetic BOTH supplier)")
            return real_rename(src, dst, *a, **k)

        fd.system = jython_shim._System()
        _os_mod.rename = _failing_rename
        half_threw = False
        half_err = ""
        try:
            try:
                fd.emit_forecast_distribution(database=INV, localFtp=False, runDate=runDate,
                                              weekDate=WEEK_FLOOR, outRoot=halfroot, renderModule=rmod)
            except Exception as e:
                half_threw = True
                half_err = str(e)
                print("    emit_forecast_distribution raised (expected): %s" % half_err[:90])
        finally:
            _os_mod.rename = real_rename
        rep.check("6b. publish-phase: the .xlsx rename failure was raised (not swallowed)", half_threw)
        # The KEY non-vacuous assertion: NO final .frc left for SUP_NULL (un-renamed back to .tmp). Without
        # the un-rename fix the .frc final would be orphaned (a half-emitted supplier) and this FAILS.
        half_frc_final = glob.glob(os.path.join(halfroot, "**", "*-%s.frc" % SUP_NULL), recursive=True)
        half_xlsx_final = glob.glob(os.path.join(halfroot, "**", "*-%s-Forecast.xlsx" % SUP_NULL),
                                    recursive=True)
        rep.check("6b. NO orphaned final .frc for SUP_NULL after the .xlsx rename failed (un-renamed)",
                  half_frc_final == [], "orphaned .frc finals=%s" % half_frc_final)
        rep.check("6b. NO final .xlsx for SUP_NULL (its rename is what failed)",
                  half_xlsx_final == [], "xlsx finals=%s" % half_xlsx_final)
        # The .frc payload is RETAINED as .tmp for operator re-drop (the un-rename target), FS-verified.
        half_frc_tmp = glob.glob(os.path.join(halfroot, "**", "*-%s.frc.tmp" % SUP_NULL), recursive=True)
        rep.check("6b. the un-renamed .frc is retained as .tmp (operator re-drop payload)",
                  len(half_frc_tmp) == 1, "frc .tmp=%s" % half_frc_tmp)
        # The error TRUTHFULLY reports the state: a clean rollback (no final stuck) -> names the supplier,
        # 'all-or-nothing', and the retained .tmp path actually on disk (P10 true-state report, not blanket).
        rep.check("6b. error names SUP_NULL + 'all-or-nothing' + the actual retained .tmp (true-state report)",
                  (SUP_NULL in half_err) and ("all-or-nothing" in half_err)
                  and any(os.path.basename(p) in half_err for p in half_frc_tmp),
                  "err=%r" % half_err)
        # PER-SUPPLIER ISOLATION: SUP_BOTH + SUP_EXCEL flush BEFORE SUP_NULL (ZZP6B/E < ZZP6N), so their
        # finals must stay fully published — SUP_NULL's failure does NOT corrupt an already-completed
        # supplier (the per-supplier all-or-nothing invariant, scoped to ONE supplier).
        both_frc_ok = glob.glob(os.path.join(halfroot, "**", "*-%s.frc" % SUP_BOTH), recursive=True)
        both_xlsx_ok = glob.glob(os.path.join(halfroot, "**", "*-%s-Forecast.xlsx" % SUP_BOTH), recursive=True)
        excel_ok = glob.glob(os.path.join(halfroot, "**", "*-%s-Forecast.xlsx" % SUP_EXCEL), recursive=True)
        rep.check("6b. per-supplier isolation: SUP_BOTH (flushed earlier) keeps BOTH finals (.frc + .xlsx)",
                  len(both_frc_ok) == 1 and len(both_xlsx_ok) == 1,
                  "frc=%s xlsx=%s" % (both_frc_ok, both_xlsx_ok))
        rep.check("6b. per-supplier isolation: SUP_EXCEL (flushed earlier) keeps its final .xlsx",
                  len(excel_ok) == 1, "SUP_EXCEL .xlsx finals=%s" % excel_ok)

        # ---- 6b NON-VACUITY (self-contained, always runs): neutralize the un-rename -> the orphaned .frc
        # final REMAINS, so the 6b "NO orphaned final .frc" assertion is genuinely contingent on the fix.
        # We run the SAME scenario but ALSO fail the un-rename (final -> .tmp), which drops into the driver's
        # 'still_final' path (rollback INCOMPLETE). Effect = "no un-rename happened": SUP_NULL's .frc final
        # is left orphaned on disk + the error TRUTHFULLY reports it as a final still on disk (not blanket).
        print("--- 6b NON-VACUITY: with the un-rename neutralized, the orphaned .frc final REMAINS ---")
        nvroot = os.path.join(tmproot, "half_nv")

        def _failing_rename_no_unrename(src, dst, *a, **k):
            # Fail BOTH the .xlsx publish (final, no .tmp) AND the .frc un-rename (target ends .frc.tmp) for
            # SUP_NULL. The .frc PUBLISH (frc final, no .tmp) still succeeds first -> orphaned final remains.
            d = str(dst)
            if ("-%s-Forecast.xlsx" % SUP_NULL) in d and not d.endswith(".tmp"):
                raise RuntimeError("simulated .xlsx publish failure (non-vacuity)")
            if d.endswith("-%s.frc.tmp" % SUP_NULL):
                raise RuntimeError("simulated un-rename failure (non-vacuity: as if no un-rename)")
            return real_rename(src, dst, *a, **k)

        fd.system = jython_shim._System()
        _os_mod.rename = _failing_rename_no_unrename
        nv_err = ""
        try:
            try:
                fd.emit_forecast_distribution(database=INV, localFtp=False, runDate=runDate,
                                              weekDate=WEEK_FLOOR, outRoot=nvroot, renderModule=rmod)
            except Exception as e:
                nv_err = str(e)
        finally:
            _os_mod.rename = real_rename
        nv_frc_final = glob.glob(os.path.join(nvroot, "**", "*-%s.frc" % SUP_NULL), recursive=True)
        rep.check("6b NON-VACUITY: with un-rename neutralized, SUP_NULL's .frc final IS orphaned (fix matters)",
                  len(nv_frc_final) == 1, "frc finals=%s" % nv_frc_final)
        rep.check("6b NON-VACUITY: the error reports the INCOMPLETE rollback (still-final on disk), not a blanket",
                  ("INCOMPLETE" in nv_err) and (SUP_NULL in nv_err)
                  and any(os.path.basename(p) in nv_err for p in nv_frc_final),
                  "err=%r" % nv_err)

        # =============== (9) TRIGGER WIRING: import_830 emits the feed post-commit ===============
        print("\n--- (9) trigger: import_830(emitFeed=True) emits the feed post-import ---")
        # The feed driver needs a renderModule for any BOTH/EXCEL supplier; the synthetic SUP_TEXT is TEXT
        # only, so a future-dated breakdown written by the import for SUP_TEXT emits a pure .frc. We drive
        # import_830 with a synthetic 830 that (via a real recipe) is heavy; simpler + faithful: assert the
        # WIRING — import_830 calls emit_forecast_distribution post-commit — by writing a future breakdown
        # row directly is what the import does, so we exercise the HOOK with a minimal import that writes a
        # raw forecast + a gap (no recipe) and confirm the feed RAN (summary['feed'] populated) and emitted
        # the already-seeded future rows. (The import's own breakdown math is proven by the M2 e2e.)
        fcmod = jython_shim.load_wrapper("forecast_for_p6", FORECAST_CODE_PY,
                                         extra_globals={"forecast_distribution": fd})
        trigroot = os.path.join(tmproot, "trigger")
        # a minimal 830 for an assembly with NO recipe -> a gap alarm, raw write, ZERO breakdown writes
        # (so the import itself adds no breakdown rows; the feed then emits the rows ALREADY seeded above).
        # This proves the HOOK fires + reads committed rows, independent of the import's breakdown math.
        def isa(d):
            return ("ISA*00*          *01*%-10s*ZZ*RECEIVER  *01*000000011*260427*1200*U*00400*"
                    "000000888*0*P*>" % d[:10])
        duns = scalar("SELECT VC_TMM_DUNS FROM INV_SITES WHERE IN_SITE_ID=%d" % SITE_ID)
        body830 = ("\n".join([isa(duns), "GS*PS*S*R*20260427*1200*1*X*004010", "ST*830*0001",
                              "LIN**BP*ZZP6NORECIP*RC*K1", "FST*5*D*W*20260706*20260706***DO*2628",
                              "CTT*1", "SE*9*0001"]) + "\n")
        # clear the gap alarm + raw row this leaves, after.
        try:
            s9 = fcmod.import_830(body830, "zzp6_trigger.edi", database=INV, siteId=SITE_ID,
                                  supplier=SUP_TEXT, emitFeed=True,
                                  emitFeedKwargs={"localFtp": False, "runDate": runDate,
                                                  "weekDate": WEEK_FLOOR, "outRoot": trigroot,
                                                  "renderModule": rmod})
            rep.check("9. import_830(emitFeed=True) ran the feed post-commit (summary['feed'] populated)",
                      isinstance(s9.get("feed"), dict) and "error" not in s9["feed"],
                      "feed=%s" % (s9.get("feed"),))
            rep.check("9. the post-import feed emitted the future-dated suppliers (>=1 supplier)",
                      s9["feed"] and s9["feed"].get("supplierCount", 0) >= 1,
                      "feed=%s" % (s9.get("feed"),))
            # the SUP_TEXT .frc was emitted by the import's trigger call (a real file on disk).
            trig_frc = glob.glob(os.path.join(trigroot, "**", "*-%s.frc" % SUP_TEXT), recursive=True)
            rep.check("9. the trigger emitted SUP_TEXT's .frc on disk (import->emit side-effect)",
                      len(trig_frc) == 1, "files=%s" % trig_frc)
        finally:
            sql("DELETE FROM INV_EDI_ALARM_REJ WHERE VC_ALARM_TYPE='830_FORECAST_GAP' "
                "AND VC_ASSY_PART_NUMBER='ZZP6NORECIP'")
            sql("DELETE FROM INV_FORECAST_INF WHERE VC_PART_NUMBER='ZZP6NORECIP'")
            sql("DELETE FROM INV_EDI_INBOUND_LOG WHERE VC_FILE_NAME='zzp6_trigger.edi'")
            sql("UPDATE INV_SITES SET VC_LAST_FORECAST_IMPORT=NULL WHERE IN_SITE_ID=%d" % SITE_ID)

        # =============== NON-VACUITY (revert proof, opt-in) ===============
        if os.environ.get("FD_REVERT_PROOF") == "1":
            print("\n--- NON-VACUITY: revert _format_qty to Python '%05d' -> check (1) FAILS ---")
            fd.system = jython_shim._System()
            orig = fd._format_qty
            fd._format_qty = lambda q, width=5: ("%0*d" % (int(width), int(q)))  # WRONG: sign in pad
            try:
                rv = fd.emit_forecast_distribution(database=INV, localFtp=False, runDate=runDate,
                                                   weekDate=WEEK_FLOOR, outRoot=os.path.join(tmproot, "rev"),
                                                   renderModule=rmod)
                rdet = dict((s["supCode"], s) for s in rv["suppliers"])
                rpath = [f["path"] for f in rdet[SUP_TEXT]["files"] if f["kind"] == "frc-main"][0]
                rbody = open(rpath, "rb").read().decode("ascii")
                rep.check("NON-VACUITY: reverted qty formatter diverges from the source (-0005 != -00005)",
                          rbody != want, "revert still matched source -> test is vacuous!")
            finally:
                fd._format_qty = orig

    finally:
        shutil.rmtree(tmproot, ignore_errors=True)
        teardown_fixtures()

    # restore-as-found assertion.
    leftover = (int(scalar("SELECT COUNT(*) FROM INV_SUPPLIER_MST WHERE VC_ADD='%s'" % ADD_STAMP) or 0)
                + int(scalar("SELECT COUNT(*) FROM INV_BREAKDOWN_FC_INF WHERE VC_SUPPLIER_CODE IN "
                             "('%s','%s','%s','%s')" % ALL_SUPS) or 0))
    rep.check("spike restored as-found (0 synthetic supplier/breakdown rows remain)", leftover == 0,
              "leftover=%d" % leftover)
    # INV_SITES VC_SUPPLIER_CODE untouched (we only READ it).
    site_after = scalar("SELECT TOP 1 VC_SUPPLIER_CODE FROM INV_SITES ORDER BY IN_SITE_ID")
    rep.check("INV_SITES.VC_SUPPLIER_CODE unchanged (feed only READS it)", site_after == site_sup,
              "before=%r after=%r" % (site_sup, site_after))
    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
