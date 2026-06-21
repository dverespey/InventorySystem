#!/usr/bin/env python3
"""test_order_file_build.py — PURE unit test for the order-FILE generator's formatters.

Imports the pure builders straight from the on-disk project library (no DB, no env, no `system`):
  * format_ord_line / format_ord_file     — the EXACT positional .ord line (OrderFormCreateF.pas:556-585)
  * pick_ship_lead                        — the weekday ship-column pick (OrderFormCreateF.pas:418-462)
  * compute_ship_offset                   — the GetShip working-day scan (OrderFormCreateF.pas:709-770)
  * _supplier_destinations                — the LocalFTP 1-vs-3 fan-out + 'NONE' skip (data-analysis §4)
  * excel.build_excel_order / build_order_sheet / _frs_to_date — the Excel data model (Cells coords)

EVERY .ord assertion derives its expected bytes from the cited .pas formats:
  supplier(raw) + frs(raw) + %8s(renban) + part(raw) + %05d(qty) + yyyymmdd(ship), CRLF-terminated.
Edge cases are NON-VACUOUS: short supplier/FRS field-shift (the legacy raw-concat), a >8 renban (no
truncation), a >5-digit qty (widens, no truncation), a negative qty.

Run:  python3 scripts/e2e/test_order_file_build.py
"""
import os, sys, importlib.util, datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report          # noqa: E402

OFDIR = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                                     "docs", "analysis", "order", "project-library", "order_file"))


def _load(name, fn):
    spec = importlib.util.spec_from_file_location(name, os.path.join(OFDIR, fn))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


of = _load("order_file_code", "code.py")
ofx = _load("order_file_excel", "excel.py")


def main():
    print("=" * 80)
    print(" ORDER-FILE generator — PURE formatter unit test (OrderFormCreateF.pas:556-585, :709-770)")
    print("=" * 80)
    rep = Report()

    # ---- 1. format_ord_line: the canonical example from the generation spec §3 ----------------------
    print("\n--- 1. format_ord_line: canonical line (spec §3 example) ---")
    # supplier ABCDE, FRS 1234567, renban H006 (->'    H006'), part 90210ABCDEF1, qty 240, ship 20260625
    row = {"supCode": "ABCDE", "frsNum": "1234567", "renban": "H006",
           "partNum": "90210ABCDEF1", "qty": 240, "shipDate": "20260625"}
    line = of.format_ord_line(row)
    # field by field per :568/:570/:571/:572/:573/:574
    want = "ABCDE" + "1234567" + "    H006" + "90210ABCDEF1" + "00240" + "20260625"
    rep.check("canonical .ord line exact bytes", line == want, "got=%r want=%r" % (line, want))
    rep.check("no delimiters (length = sum of field widths)",
              len(line) == 5 + 7 + 8 + 12 + 5 + 8, "len=%d" % len(line))

    # ---- 2. each format primitive, isolated --------------------------------------------------------
    print("\n--- 2. field formats (%8s renban, %05d qty, raw fields) ---")
    rep.check("renban right-justified to 8 ('H006' -> '    H006')",
              of.format_ord_line(dict(row, renban="H006"))[12:20] == "    H006",
              of.format_ord_line(dict(row, renban="H006"))[12:20])
    rep.check("renban already-8 unchanged ('16GY1234')",
              of.format_ord_line(dict(row, renban="16GY1234"))[12:20] == "16GY1234")
    rep.check("renban >8 NOT truncated ('16GY12345' stays 9 wide)",
              of.format_ord_line(dict(row, renban="16GY12345")).find("16GY12345") >= 0,
              of.format_ord_line(dict(row, renban="16GY12345")))
    rep.check("qty zero-padded to 5 (240 -> '00240')",
              of.format_ord_line(dict(row, qty=240)).find("00240") >= 0)
    rep.check("qty 1200 -> '01200'", of.format_ord_line(dict(row, qty=1200)).find("01200") >= 0)
    rep.check("qty >99999 widens, no truncation (123456 -> '123456')",
              of.format_ord_line(dict(row, qty=123456)).find("123456") >= 0)
    # BLOCKER-2: Pascal format('%.5d',[-5]) = '-00005' (sign OUTSIDE the 5-digit zero-pad), NOT Python
    # '%05d' % -5 = '-0005' (sign consumes a pad slot). The .ord must carry the Delphi shape.
    rep.check("qty -5 -> '-00005' (Pascal %.5d: sign OUTSIDE the 5-digit pad, NOT Python %05d '-0005')",
              of.format_ord_line(dict(row, qty=-5)).find("-00005") >= 0 and
              of.format_ord_line(dict(row, qty=-5)).find("-0005" + "2") < 0,
              of.format_ord_line(dict(row, qty=-5)))
    rep.check("_format_qty matches Pascal format('%.5d') field-by-field (+ve and -ve)",
              of._format_qty(240) == "00240" and of._format_qty(0) == "00000" and
              of._format_qty(123456) == "123456" and of._format_qty(-5) == "-00005" and
              of._format_qty(-50) == "-00050" and of._format_qty(-12345) == "-12345",
              "qty(-5)=%r qty(240)=%r" % (of._format_qty(-5), of._format_qty(240)))

    # ---- 3. raw-field SHIFT (legacy H3): short supplier/FRS shifts following fields -----------------
    print("\n--- 3. raw-concat field shift (legacy H3 — supplier/FRS/part NOT padded) ---")
    short = of.format_ord_line({"supCode": "AB", "frsNum": "12345", "renban": "R",
                                "partNum": "PART", "qty": 1, "shipDate": "20260101"})
    # AB(2) + 12345(5) + '       R'(8) + PART(4) + 00001(5) + 20260101(8). The renban now starts at byte 7
    # (not 12) — every following field shifts LEFT. We assert the documented raw-concat, NOT a padded line.
    rep.check("short supplier+FRS produce a SHORTER line (fields shift left — H3 faithful)",
              short == "AB" + "12345" + "       R" + "PART" + "00001" + "20260101",
              "got=%r" % short)
    rep.check("the rebuild does NOT silently pad supplier/FRS/part",
              short.startswith("AB12345"), "got=%r" % short)

    # ---- 4. format_ord_file: CRLF-terminated, every line incl. last ---------------------------------
    print("\n--- 4. format_ord_file: CRLF after every line (legacy Writeln) ---")
    payload = of.format_ord_file([row, dict(row, renban="H007", qty=120)])
    rep.check("two lines joined CRLF + trailing CRLF",
              payload == want + "\r\n" + of.format_ord_line(dict(row, renban="H007", qty=120)) + "\r\n",
              "repr=%r" % payload)
    rep.check("empty rows -> empty payload", of.format_ord_file([]) == "")
    rep.check("payload ends with CRLF (trailing, like Writeln)", payload.endswith("\r\n"))

    # ---- 5. pick_ship_lead: weekday override else base Ship ----------------------------------------
    print("\n--- 5. pick_ship_lead (weekday override else base 'Ship') ---")
    sd = {"Ship": 3, "ShipM": 0, "ShipT": 5, "ShipW": 0, "ShipTh": 0, "ShipF": 2, "ShipS": 0}
    rep.check("Monday (1), ShipM=0 -> base Ship=3", of.pick_ship_lead(sd, 1) == 3, str(of.pick_ship_lead(sd, 1)))
    rep.check("Tuesday (2), ShipT=5 -> override 5", of.pick_ship_lead(sd, 2) == 5)
    rep.check("Friday (5), ShipF=2 -> override 2", of.pick_ship_lead(sd, 5) == 2)
    rep.check("Sunday (7) -> base Ship=3 (fallthrough)", of.pick_ship_lead(sd, 7) == 3)
    rep.check("Saturday (6), ShipS=0 -> base Ship=3", of.pick_ship_lead(sd, 6) == 3)

    # ---- 6. compute_ship_offset: the GetShip working-day scan --------------------------------------
    print("\n--- 6. compute_ship_offset (GetShip: skip weekends + 'H' holidays) ---")
    # base = a Monday. With NO weekends/holidays in range, lead N -> offset N (each valid day is +1).
    monday = datetime.date(2026, 6, 22)              # 2026-06-22 is a Monday (isoweekday()==1)
    rep.check("base IS a Monday (fixture sanity)", monday.isoweekday() == 1)
    base = monday.toordinal()
    no_we = lambda o: datetime.date.fromordinal(o).isoweekday() >= 6
    no_hol = lambda o: False
    # lead 0 -> 0 (the while never runs)
    rep.check("lead 0 -> offset 0", of.compute_ship_offset(0, base, no_we, no_hol) == 0)
    # lead 1 from Monday: probe Tue(+1) valid -> x=1
    rep.check("lead 1 from Mon -> offset 1 (Tue)", of.compute_ship_offset(1, base, no_we, no_hol) == 1)
    # lead 4 from Monday: Tue,Wed,Thu,Fri valid -> x=4
    rep.check("lead 4 from Mon -> offset 4 (Fri)", of.compute_ship_offset(4, base, no_we, no_hol) == 4)
    # lead 5 from Monday: Tue..Fri (4 valid) then Sat/Sun skipped (x=5,6) then next Mon valid (x=7)
    rep.check("lead 5 from Mon -> offset 7 (skips the weekend)",
              of.compute_ship_offset(5, base, no_we, no_hol) == 7,
              "got=%d" % of.compute_ship_offset(5, base, no_we, no_hol))
    # holiday on the Wed (+2) pushes the count: lead 3 from Mon with Wed an 'H' -> Tue(x1),Wed skip(x2),
    #   Thu(x3),Fri... -> need 3 valid: Tue,Thu,Fri -> x=4
    wed = (monday + datetime.timedelta(days=2)).toordinal()
    hol_wed = lambda o: o == wed
    rep.check("lead 3 from Mon, Wed is 'H' -> offset 4 (Tue,Thu,Fri; Wed skipped)",
              of.compute_ship_offset(3, base, no_we, hol_wed) == 4,
              "got=%d" % of.compute_ship_offset(3, base, no_we, hol_wed))
    rep.check("'O'/'X'/'W' days are NOT skipped (only weekend + 'H') — the GetShip flag",
              of.compute_ship_offset(3, base, no_we, no_hol) == 3,
              "got=%d" % of.compute_ship_offset(3, base, no_we, no_hol))

    # ---- 7. _supplier_destinations: LocalFTP fan-out + 'NONE' skip ---------------------------------
    print("\n--- 7. destinations (LocalFTP off=1 / on=3; 'NONE' logistics skipped) ---")
    off = of._supplier_destinations("S:\\GY", "NONE", localFtp=False)
    rep.check("LocalFTP off -> ONE destination (supplier dir only)",
              [k for k, _ in off] == ["supplier"], str(off))
    on_none = of._supplier_destinations("S:\\GY", "NONE", localFtp=True)
    rep.check("LocalFTP on + 'NONE' -> supplier + archive (logistics SKIPPED)",
              [k for k, _ in on_none] == ["supplier", "archive"], str(on_none))
    on_log = of._supplier_destinations("S:\\GY", "S:\\TLDL", localFtp=True)
    rep.check("LocalFTP on + real logistics -> supplier + logistics + archive (3)",
              [k for k, _ in on_log] == ["supplier", "logistics", "archive"], str(on_log))
    rep.check("archive dir is <supplierDir>/Archive",
              on_log[2][1].endswith(os.path.join("S:\\GY", "Archive")) or
              on_log[2][1].endswith("Archive"), on_log[2][1])

    # ---- 8. _file_kind + _ord_filename + _clean_name -----------------------------------------------
    print("\n--- 8. file-kind mapping, filename, name cleaning ---")
    rep.check("'TEXT' -> (text, no excel)", of._file_kind("TEXT") == (True, False))
    rep.check("'EXCEL' -> (no text, excel)", of._file_kind("EXCEL") == (False, True))
    rep.check("'BOTH' -> (text, excel)", of._file_kind("BOTH") == (True, True))
    rep.check("NULL output type -> BOTH (legacy H2 default)", of._file_kind(None) == (True, True))
    rep.check("'' output type -> BOTH (legacy H2 default)", of._file_kind("") == (True, True))
    # strips ',' and '.' (OrderFormCreateF.pas:227) then '\' at filename time (ANSIReplaceStr) — the
    # chars are REMOVED, not replaced with a space. 'ACME, Inc.\X' -> 'ACME Inc' + 'X' (no space) -> 'ACME IncX'.
    rep.check("clean name strips , . and backslash (removed, not spaced)",
              of._clean_name("ACME, Inc.\\X") == "ACME IncX", of._clean_name("ACME, Inc.\\X"))
    rep.check("no-timestamp filename = <name>-<sup>.ord",
              of._ord_filename("ACME", "0501B", False, "20260621000000" + "00") == "ACME-0501B.ord")
    rep.check("timestamp filename = <name>-<sup><fts>.ord (NO '-' before fts — legacy H7)",
              of._ord_filename("ACME", "0501B", True, "2026062100000000") == "ACME-0501B2026062100000000.ord",
              of._ord_filename("ACME", "0501B", True, "2026062100000000"))

    # ---- 8b. BLOCKER-1: the bit-coercion that drives the timestamp-in-filename decision -------------
    # Legacy: lasttimestamp := fieldbyname('Order File Timestamp').AsBoolean (OrderFormCreateF.pas:228);
    # AsBoolean on a SQL bit is (value <> 0). The bug was bool('0') == True (any non-empty str is truthy),
    # which wrongly TIMESTAMPED 15/16 live suppliers (bit 0) under the spike sqlcmd-string transport.
    print("\n--- 8b. _coerce_bit (BLOCKER-1: SQL bit -> AsBoolean, string '0'/'1' AND JDBC bool/int) ---")
    fts16 = "2026062100000000"
    rep.check("_coerce_bit('0') is False (the shim's string transport — the bug input)",
              of._coerce_bit("0") is False, repr(of._coerce_bit("0")))
    rep.check("_coerce_bit('1') is True", of._coerce_bit("1") is True)
    rep.check("_coerce_bit(0/1) is False/True (JDBC Integer transport)",
              of._coerce_bit(0) is False and of._coerce_bit(1) is True)
    rep.check("_coerce_bit(False/True) is False/True (JDBC Boolean transport)",
              of._coerce_bit(False) is False and of._coerce_bit(True) is True)
    rep.check("_coerce_bit(None/'') is False (NULL / empty -> AsBoolean False)",
              of._coerce_bit(None) is False and of._coerce_bit("") is False)
    rep.check("_coerce_bit('true'/'false') is True/False (defensive string variants)",
              of._coerce_bit("true") is True and of._coerce_bit("false") is False)
    # The legacy-anchored end-to-end filename claim: bit 0 -> NO timestamp; bit 1 -> timestamp.
    # PACIFIC_MFG (sup 11111, bit 0) -> 'PACIFIC_MFG-11111.ord' (legacy :321, NOT the timestamped form).
    name0 = of._ord_filename(of._clean_name("PACIFIC MFG"), "11111", of._coerce_bit("0"), fts16)
    rep.check("bit '0' supplier (PACIFIC_MFG/11111) -> STABLE 'PACIFIC MFG-11111.ord' (legacy :321, no fts)",
              name0 == "PACIFIC MFG-11111.ord", "got=%r" % name0)
    name1 = of._ord_filename(of._clean_name("SUPERIOR"), "38844", of._coerce_bit("1"), fts16)
    rep.check("bit '1' supplier (SUPERIOR/38844) -> TIMESTAMPED 'SUPERIOR-38844<fts>.ord' (legacy :290)",
              name1 == "SUPERIOR-38844" + fts16 + ".ord", "got=%r" % name1)

    # ---- 9. Excel data model (Cells coordinates — order-file-generation-spec §7) --------------------
    print("\n--- 9. Excel ORDER data model (OrderTemplate.xls cells, :279/:545) ---")
    sup = {"Supplier Name": "ACME", "Address": "1 Main St", "Supplier Code": "0501B",
           "City": "Tupelo", "State": "MS", "Zip": "38801"}
    eo = ofx.build_excel_order(sup, [dict(row, shipDateMdy="06/25/2026")])
    rep.check("Excel-order header: F1=name, F2=address (Cells[1,6]/[2,6])",
              eo["header"]["F1"] == "ACME" and eo["header"]["F2"] == "1 Main St", str(eo["header"]))
    rep.check("Excel-order body starts row 10 with A/B/C/D/E (part/frs/renban/qty/ship)",
              eo["startRow"] == 10 and eo["body"][0]["A10"] == "90210ABCDEF1" and
              eo["body"][0]["D10"] == 240 and eo["body"][0]["E10"] == "06/25/2026", str(eo["body"][0]))

    print("\n--- 9b. Excel ORDER SHEET data model (WHEEL/TIRE variants, :253/:522) ---")
    rows = [{"partNum": "P1", "partsName": "TIRE A", "kanban": "16GY", "oneLotQty": 40,
             "renban": "16GY118", "qty": 1200, "frsNum": "4072201", "shipDateMdyShort": "06/25/26"}]
    wheel = ofx.build_order_sheet(sup, rows, "WHEEL", "2026")
    rep.check("WHEEL order-sheet: header A11=name, B11=code, C11=renban, body G16=qty, E16=1LotQty",
              wheel[0]["header"]["A11"] == "ACME" and wheel[0]["header"]["C11"] == "16GY118" and
              wheel[0]["body"][0]["G16"] == 1200 and wheel[0]["body"][0]["E16"] == 40,
              str(wheel[0]))
    tire = ofx.build_order_sheet(sup, rows, "TIRE", "2026")
    rep.check("TIRE order-sheet: E16=renban (NOT 1LotQty), G16=qty, no C11 renban-in-header",
              tire[0]["body"][0]["E16"] == "16GY118" and tire[0]["body"][0]["G16"] == 1200 and
              "C11" not in tire[0]["header"], str(tire[0]))
    rep.check("FRS->date cell H8 (FRS 4072201 -> '07/22/' + 2026[:3]+'4' = 07/22/2024)",
              wheel[0]["header"]["H8"] == "07/22/2024", wheel[0]["header"]["H8"])

    # WHEEL pages: one workbook PER renban; >perPage body rows paginate
    many = [dict(rows[0], renban="R1") for _ in range(10)] + [dict(rows[0], renban="R2") for _ in range(2)]
    pages = ofx.build_order_sheet(sup, many, "WHEEL", "2026")
    rep.check("WHEEL -> one workbook page-group per renban (R1, R2)",
              sorted(set(p["renban"] for p in pages)) == ["R1", "R2"], str([p["renban"] for p in pages]))

    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
