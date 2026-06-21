#!/usr/bin/env python3
"""test_forecast_import_build.py — PURE unit test for the M2 EDI 830 forecast importer's I/O-free core
(docs/analysis/edi/inbound/project-library/forecast/code.py): parse_830, pick_ratio, explode, day_spread,
production_offset.

WHAT THIS PROVES
----------------
The PURE logic is FAITHFUL to the legacy ForecastBreakdownF.pas algorithm (the parity oracle — see
forecast-import-algorithm.md), field by field:

  * parse_830 reads the LIN/FST walk by the EXACT legacy element offsets: assembly=LIN03 (delSL[3]),
    kanban=LIN05 (delSL[5]), FST qty=FST01 (delSL[1]), weekDate=FST04 (delSL[4]), week=copy(FST09,3,2)
    (chars 3-4 of the DO ref). Multi-LIN + multi-FST. Stops at CTT.
  * D10 RAW WEEK: DO ref 2624 -> stored weekNumber 24 (NOT 25). The FirstProductionDay offset is NEVER
    applied to the stored week (production_offset feeds ONLY the holiday-calendar lookup).
  * explode reproduces the integer-div ratio math: tireCount=((wc*fr)//100)*tr//100 etc.; a 3-way share
    BC (ratios 40/20/40) scales the FULL count on each assembly LIN to its share, summing 100%. Valve/
    film/label/misc all use wheelCount (faithful).
  * pick_ratio: blank/' ' effective-month = the default ratio; a dated row whose yymm==weekdate-yyyy wins.
  * day_spread: weekly qty // working-days, with the WHOLE remainder on the first working day; a holiday
    (an off day) reduces the working-day count and re-spreads.

HONEST VERIFICATION (no golden 830)
-----------------------------------
There is NO captured TEMA 830 on disk (the D10 golden EDI/830000008976.EDI is gitignored client data).
So this does NOT claim byte-parity with a real TEMA file — it asserts the PURE logic is faithful to the
ForecastBreakdownF.pas copy()-offsets + the integer-div explode + the day-spread the legacy itself uses,
on fixtures built field-by-field at those exact offsets. The exact ISA element index + the real holiday
day-spread are PENDING a golden (same stance as 856/810/997/824).

Non-vacuous: every assertion compares parsed/computed output to an INDEPENDENTLY-constructed expected
value (the 830 fixture is assembled at named element positions; the explode/spread expecteds are computed
by hand from the legacy formulas), so a wrong slice/formula in code.py fails here.

Run:  python3 scripts/e2e/test_forecast_import_build.py
"""
import os
import sys
import importlib.util

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report          # noqa: E402

CODE_PY = os.path.normpath(os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "..", "..",
    "docs", "analysis", "edi", "inbound", "project-library", "forecast", "code.py"))


def load_pure():
    """Import forecast/code.py as a pure module (no `system` injected — the pure functions reference no
    gateway global at import or call time)."""
    spec = importlib.util.spec_from_file_location("forecast_pure", CODE_PY)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


# --- 830 fixture builders (legacy element offsets) --------------------------------------------------

def lin(assembly, kanban):
    """A LIN segment placing assembly at delSL[3] and kanban at delSL[5] (ForecastBreakdownF.pas:263-264).
    'LIN' delSL[0], '' delSL[1], 'BP' delSL[2], assembly delSL[3], 'RC' delSL[4], kanban delSL[5]."""
    return "LIN**BP*%s*RC*%s" % (assembly, kanban)


def fst(qty, weekDate, doRef):
    """An FST segment: qty at delSL[1] (FST01), weekDate at delSL[4] (FST04), DO ref at delSL[9] (FST09).
    'FST' delSL[0], qty delSL[1], 'D' delSL[2], 'W' delSL[3], weekDate delSL[4], endDate delSL[5], two
    empties delSL[6..7], 'DO' delSL[8], doRef delSL[9]. So copy(delSL[9],3,2) = doRef chars 3-4."""
    return "FST*%s*D*W*%s*%s***DO*%s" % (qty, weekDate, weekDate, doRef)


def build_830(entries):
    """entries = [(assembly, kanban, [(qty, weekDate, doRef), ...]), ...]. Wrap with a minimal ISA/GS/ST
    envelope + a CTT trailer (the parse stops at CTT). Returns the file body str."""
    out = ["ISA*00*          *01*SENDER    *ZZ*RECEIVER  *01*000000011*260427*1200*U*00400*000000001*0*P*>",
           "GS*PS*SENDER*RECEIVER*20260427*1200*1*X*004010",
           "ST*830*0001"]
    for (assembly, kanban, weeks) in entries:
        out.append(lin(assembly, kanban))
        for (qty, weekDate, doRef) in weeks:
            out.append(fst(qty, weekDate, doRef))
    out.append("CTT*%d" % len(entries))
    out.append("SE*99*0001")
    return "\n".join(out) + "\n"


def main():
    rep = Report()
    print("M2 forecast-import PURE unit test (parse_830 / explode / day_spread / D10 raw-week)\n")
    m = load_pure()

    # ---------------------------------------------------------------------------------------------
    # 1. parse_830 — multi-LIN, multi-FST, exact offsets + D10 raw week.
    # ---------------------------------------------------------------------------------------------
    body = build_830([
        ("42600F261100", "RC01", [(144, "20260615", "2624"), (200, "20260622", "2625")]),
        ("42670FET9000", "RC02", [(1000, "20260615", "2624")]),
    ])
    entries = m.parse_830(body)
    rep.check("parse_830: 2 LIN entries", len(entries) == 2, "got %d" % len(entries))
    rep.check("parse_830: assembly LIN03", entries[0]["assembly"] == "42600F261100",
              repr(entries[0]["assembly"]))
    rep.check("parse_830: kanban LIN05", entries[0]["kanban"] == "RC01", repr(entries[0]["kanban"]))
    rep.check("parse_830: 2 FST weeks on entry 0", len(entries[0]["weeks"]) == 2,
              "got %d" % len(entries[0]["weeks"]))
    w0 = entries[0]["weeks"][0]
    rep.check("parse_830: FST01 qty -> weekCount=144", w0["weekCount"] == 144, str(w0["weekCount"]))
    rep.check("parse_830: FST04 weekDate=20260615", w0["weekDate"] == "20260615", w0["weekDate"])
    # D10 — THE RAW-WEEK PROOF: DO 2624 -> chars 3-4 = '24' -> stored week 24 (NOT 25).
    rep.check("D10 RAW WEEK: DO 2624 -> weekNumber 24 (NOT 25)", w0["weekNumber"] == 24,
              "got %d" % w0["weekNumber"])
    rep.check("parse_830: 2nd FST DO 2625 -> week 25", entries[0]["weeks"][1]["weekNumber"] == 25,
              "got %d" % entries[0]["weeks"][1]["weekNumber"])
    rep.check("parse_830: entry 1 single FST qty 1000",
              len(entries[1]["weeks"]) == 1 and entries[1]["weeks"][0]["weekCount"] == 1000,
              str(entries[1]["weeks"]))
    # stops at CTT (no phantom entry from the SE trailer).
    rep.check("parse_830: stops at CTT (exactly 2 entries)", len(entries) == 2)
    # empty / non-830 body -> [].
    rep.check("parse_830: envelope-only -> []", m.parse_830("ISA*..\nGS*..\nST*830\nCTT*0\nSE*1\n") == [])

    # non-numeric FST -> ForecastImportError (corrupt feed, not silent garbage).
    bad = build_830([("42600F261100", "RC01", [("ABC", "20260615", "2624")])])
    try:
        m.parse_830(bad)
        rep.check("parse_830: non-numeric qty raises", False, "no exception")
    except m.ForecastImportError:
        rep.check("parse_830: non-numeric qty raises ForecastImportError", True)

    # ---------------------------------------------------------------------------------------------
    # 2. production_offset — D10 / R1: offset is firstWeek-1, used ONLY for the holiday lookup.
    # ---------------------------------------------------------------------------------------------
    rep.check("production_offset: firstWeek 2 (2026) -> +1", m.production_offset(2) == 1)
    rep.check("production_offset: firstWeek 1 -> 0", m.production_offset(1) == 0)
    rep.check("production_offset: garbage -> 0", m.production_offset("x") == 0)

    # ---------------------------------------------------------------------------------------------
    # 3. pick_ratio — default (blank effmonth) vs a dated override.
    # ---------------------------------------------------------------------------------------------
    default_recipe = {"effectiveMonth": " ", "tireRatio": 100, "wheelRatio": 100, "forecastRatio": 100,
                      "tire": "TIRE00000001", "wheel": "WHEEL0000001", "valve": "", "film": "",
                      "label": "", "misc1": "", "misc2": ""}
    rep.check("pick_ratio: blank effmonth -> default",
              m.pick_ratio([default_recipe], "20260615") is default_recipe)
    # Realistic effmonth '2026/06' -> tm = effmonth[2:4]+effmonth[5:7] = '26'+'06' = '2606'.
    # A real weekDate '20260615' -> wm = weekDate[0:4] = '2026'. '2606' != '2026' -> NO match. This is the
    # AMBIGUITY / dead-branch finding (algorithm §C.3): on real yyyymmdd weekdates the dated branch NEVER
    # fires, so the default ratio always wins (matches live: all 50 recipes have effmonth=' ').
    dated = dict(default_recipe, effectiveMonth="2026/06", tireRatio=50)
    rep.check("pick_ratio: dated 2026/06 vs real weekdate 2026 (tm='2606'!=wm='2026') -> default kept",
              m.pick_ratio([default_recipe, dated], "20260615") is default_recipe)
    # The dated branch DOES fire only when wm (weekDate[0:4]) == tm. For tm '2606' that needs a contrived
    # weekDate beginning '2606' — proving the branch is reachable but mis-aligned on real yyyymmdd data.
    rep.check("pick_ratio: dated 2026/06 matches a contrived weekdate '2606....' (branch reachable)",
              m.pick_ratio([default_recipe, dated], "26060101") is dated)
    rep.check("pick_ratio: no default + no match -> None",
              m.pick_ratio([dict(default_recipe, effectiveMonth="2026/06")], "20260615") is None)

    # ---------------------------------------------------------------------------------------------
    # 4. explode — the integer-div ratio math + the 3-way share fanout (ratios summing 100).
    # ---------------------------------------------------------------------------------------------
    # 4a. single assembly, all ratios 100: tireCount=wheelCount=weekCount.
    comps = m.explode(144, default_recipe)
    rep.check("explode: 100/100/100 -> tire 144, wheel 144",
              comps == [("TIRE00000001", 144), ("WHEEL0000001", 144)], str(comps))

    # 4b. a 3-way share BC (1 BC -> 3 assemblies; tire ratios 40/20/40, forecast 100). TEMA sends the FULL
    #     count (1000) on EACH assembly LIN; the explode scales each by its own tireRatio/100. Sum == 100%.
    share = [
        {"effectiveMonth": " ", "tireRatio": 40, "wheelRatio": 40, "forecastRatio": 100,
         "tire": "TIRE0000000A", "wheel": "WHEELSHARE00", "valve": "", "film": "", "label": "",
         "misc1": "", "misc2": ""},
        {"effectiveMonth": " ", "tireRatio": 20, "wheelRatio": 20, "forecastRatio": 100,
         "tire": "TIRE0000000B", "wheel": "WHEELSHARE00", "valve": "", "film": "", "label": "",
         "misc1": "", "misc2": ""},
        {"effectiveMonth": " ", "tireRatio": 40, "wheelRatio": 40, "forecastRatio": 100,
         "tire": "TIRE0000000C", "wheel": "WHEELSHARE00", "valve": "", "film": "", "label": "",
         "misc1": "", "misc2": ""},
    ]
    tire_total = 0
    for r in share:
        cs = m.explode(1000, r)
        tire_total += dict(cs)["TIRE" + r["tire"][4:]] if False else dict(cs)[r["tire"]]
    rep.check("explode: 3-way share 40/20/40 of 1000 -> 400+200+400 = 1000 (sums to 100%)",
              tire_total == 1000, "got %d" % tire_total)
    # one share leg explicitly: 40% of 1000 = 400.
    leg = dict(m.explode(1000, share[0]))
    rep.check("explode: 40% leg of 1000 -> tire 400", leg["TIRE0000000A"] == 400, str(leg))

    # 4c. integer truncation: weekCount 145, forecast 100, tire 40 -> ((145*100)//100)*40//100 = 5800//100 = 58.
    trunc = dict(m.explode(145, share[0]))
    rep.check("explode: int-div truncation 145 @ 40% -> 58", trunc["TIRE0000000A"] == 58, str(trunc))

    # 4d. a zero ratio -> zero count (legacy guard).
    zero = dict(default_recipe, tireRatio=0)
    rep.check("explode: tireRatio 0 -> all zero", m.explode(144, zero) == [("TIRE00000001", 0),
                                                                            ("WHEEL0000001", 0)])
    # 4e. valve/film/label/misc all use wheelCount (faithful).
    full = {"effectiveMonth": " ", "tireRatio": 50, "wheelRatio": 30, "forecastRatio": 100,
            "tire": "TIRE00000001", "wheel": "WHEEL0000001", "valve": "VALVE0000001",
            "film": "FILM00000001", "label": "LABEL0000001", "misc1": "MISC10000001", "misc2": ""}
    cf = dict(m.explode(1000, full))
    # tire = 500, wheel/valve/film/label/misc1 = 300 each.
    rep.check("explode: tire uses tireCount(500)", cf["TIRE00000001"] == 500, str(cf))
    rep.check("explode: valve/film/label/misc1 all use wheelCount(300)",
              cf["VALVE0000001"] == 300 and cf["FILM00000001"] == 300 and cf["LABEL0000001"] == 300
              and cf["MISC10000001"] == 300, str(cf))
    rep.check("explode: misc2 blank -> omitted", "" not in cf and len(cf) == 6, str(cf))
    # short component code (<=2 chars) is skipped.
    short = dict(default_recipe, valve="V")
    rep.check("explode: <=2-char component code skipped",
              all(c != "V" for c, _ in m.explode(100, short)))

    # ---------------------------------------------------------------------------------------------
    # 5. day_spread — no-special-date week, holiday (H) week, and an OVERTIME ('O') Saturday week.
    #    day_spread now takes the FULL AD_GetSpecialDateWeek rows as (dayNumber, statusAbbr) tuples (so it
    #    can reproduce the legacy running DEC/INC days counter exactly — H/X off+DEC, O on+INC). [] = no
    #    special dates. (BLOCKER-2 fix.)
    # ---------------------------------------------------------------------------------------------
    # 5a. NO special date: 138 over Mon-Fri (5 days). 138//5=27 r3 -> day1=27+3=30, days2-5=27, Sat/Sun=0.
    spread = m.day_spread(138, [])        # Sat(6)/Sun(7) off automatically (the Mon-Fri base)
    rep.check("day_spread: 138 / Mon-Fri (no special date) -> [30,27,27,27,27,0,0]",
              spread == [30, 27, 27, 27, 27, 0, 0], str(spread))
    rep.check("day_spread: sum preserved (138)", sum(spread) == 138, str(sum(spread)))

    # 5b. WITH a holiday (H) on Wednesday (day 3 off): 138 over 4 working days (Mon,Tue,Thu,Fri).
    #     138//4=34 r2 -> day1=34+2=36, day2=34, day3=0(holiday), day4=34, day5=34, Sat/Sun=0.
    spread_h = m.day_spread(138, [(3, "H")])
    rep.check("day_spread: 138 / Wed H-holiday -> [36,34,0,34,34,0,0]",
              spread_h == [36, 34, 0, 34, 34, 0, 0], str(spread_h))
    rep.check("day_spread: holiday week sum preserved (138)", sum(spread_h) == 138, str(sum(spread_h)))
    # 5b'. X (NON-PRODUCTION) behaves identically to H (both turn the day OFF + DEC days).
    rep.check("day_spread: X turns the day OFF same as H ([36,34,0,34,34,0,0])",
              m.day_spread(138, [(3, "X")]) == [36, 34, 0, 34, 34, 0, 0],
              str(m.day_spread(138, [(3, "X")])))

    # 5c. OVERTIME 'O' SATURDAY (BLOCKER-2 — the legacy parity case proven on live COROLLA wk23, qty 138).
    #     The legacy turns Saturday (day 6, started OFF) ON and INC(days) 5->6 -> 138//6=23 r0 -> all six
    #     worked days (Mon-Sat) get 23; Sun=0. This is [23,23,23,23,23,23,0] — NOT the old (Sat-ignored)
    #     [30,27,27,27,27,0,0]. Every bucket differs from the pre-fix behavior.
    spread_o = m.day_spread(138, [(6, "O")])
    rep.check("BLOCKER-2: O-Saturday turns Sat ON (days 5->6) -> [23,23,23,23,23,23,0] (legacy parity)",
              spread_o == [23, 23, 23, 23, 23, 23, 0], str(spread_o))
    rep.check("BLOCKER-2: O-Saturday week sum preserved (138)", sum(spread_o) == 138, str(sum(spread_o)))
    rep.check("BLOCKER-2: O-Saturday is NOT the pre-fix Sat-ignored [30,27,27,27,27,0,0]",
              spread_o != [30, 27, 27, 27, 27, 0, 0], str(spread_o))
    # 5c'. an O-Saturday that does NOT divide evenly: qty 140 / 6 = 23 r2 -> day1(Mon)=23+2=25, days2-6=23.
    spread_o2 = m.day_spread(140, [(6, "O")])
    rep.check("day_spread: O-Saturday 140/6 -> [25,23,23,23,23,23,0] (remainder on first working day)",
              spread_o2 == [25, 23, 23, 23, 23, 23, 0], str(spread_o2))

    # 5d. MIXED week — an H on Mon (day1 off, DEC) AND an O on Sat (day6 on, INC) -> still 5 working days
    #     (Tue,Wed,Thu,Fri,Sat). 138//5=27 r3 -> first WORKING day (Tue=idx2) gets 27+3=30 ->
    #     [0,30,27,27,27,27,0]. Proves both branches compose via the running counter.
    spread_mix = m.day_spread(138, [(1, "H"), (6, "O")])
    rep.check("day_spread: Mon-H + Sat-O -> 5 days (Tue..Sat) -> [0,30,27,27,27,27,0]",
              spread_mix == [0, 30, 27, 27, 27, 27, 0], str(spread_mix))

    # 5e. PARITY of the legacy RUNNING COUNTER (not count(workday)) — the latent edges we deliberately
    #     reproduce (ForecastBreakdownF.pas:1403-1407 "exactly"):
    #   * 'O' on an already-ON weekday (Mon) double-counts days: 5+1=6 (NOT 5). 120//6=20 -> all Mon-Fri
    #     get 20, Sat/Sun=0; sum=100, NOT 120 (the legacy LOSES the remainder differently). Assert the
    #     legacy NUMBER: days=6 -> 120//6=20 r0 -> [20,20,20,20,20,0,0].
    rep.check("day_spread (parity edge): O on already-ON Mon double-counts days (6) -> [20,20,20,20,20,0,0]",
              m.day_spread(120, [(1, "O")]) == [20, 20, 20, 20, 20, 0, 0],
              str(m.day_spread(120, [(1, "O")])))
    #   * H/X on an already-OFF day (Sat) under-counts days: 5-1=4 (NOT 5). 120//4=30 -> Mon-Thu get 30,
    #     Fri... wait Fri is still ON: working = Mon,Tue,Wed,Thu,Fri = 5 set-on but days counter = 4.
    #     ratiocount=120//4=30 r0; spread over the 5 ON days gives 30 each -> sum 150 (the legacy OVER-
    #     spreads). Assert the legacy NUMBER: [30,30,30,30,30,0,0].
    rep.check("day_spread (parity edge): H on already-OFF Sat under-counts days (4) -> [30,30,30,30,30,0,0]",
              m.day_spread(120, [(6, "H")]) == [30, 30, 30, 30, 30, 0, 0],
              str(m.day_spread(120, [(6, "H")])))

    # 5f. all-week holiday (Mon-Fri all off) -> all zeros (days==0).
    rep.check("day_spread: all working days H-off -> zeros",
              m.day_spread(138, [(1, "H"), (2, "H"), (3, "H"), (4, "H"), (5, "H")]) == [0, 0, 0, 0, 0, 0, 0])

    # 5g. zero qty -> zeros (no NULL poison — D-Bug-4).
    rep.check("day_spread: 0 qty -> all-zero (never None)", m.day_spread(0, []) == [0] * 7)
    rep.check("day_spread: every bucket is an int 0, not None",
              all(isinstance(x, int) for x in m.day_spread(0, [])))

    return rep.summary_exit()


if __name__ == "__main__":
    sys.exit(main())
