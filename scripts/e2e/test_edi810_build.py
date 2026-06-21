#!/usr/bin/env python3
"""test_edi810_build.py — unit test for the PURE outbound-810 INVOICE segment builder (edi810.build_810).

Exercises the byte-faithful X12 segment assembly that PORTS the legacy EDI810Object.T810EDI
(EDI810Object.pas) with NO database: edi810/code.py is CPython-importable straight from the on-disk project
library. Twin of test_edi856_build.py.

Source truth:
  docs/analysis/edi/810/edi810-wire-format.md          — the byte-exact segment spec (THE target)
  docs/analysis/edi/810/report-edi810-data-analysis.md — the data feed + D6 pricing
  EDI810Object.pas                                     — the original builder (EXPECTED bytes derived here)

THE BYTES ARE DERIVED FROM EDI810Object.pas (the 856 lesson) — the `expected` array below is hand-computed
by reading the .pas Write/format calls, NOT copied from the rebuild's sep.join, so the test catches the
rebuild DIVERGING from legacy, never blesses it. EXCEPT the MONEY fields (IT104 + TDS): those assert the
CLEAN values (DECISION 810-1/810-2, David locked) — the legacy buggy values are documented as a DELIBERATE
divergence and asserted to be ABSENT (anti-regression).

NON-VACUOUS COVERAGE (synthetic multi-manifest, multi-item invoice):
  * the FULL ordered segment list asserted BYTE-EXACT vs the .pas derivation (ISA..IEA, the IT1 loop, the
    interior + trailing REF*MK / DTM*050 breaks). GS01='IN', BIG02=SupplierCode, IT101 M391/M390.
  * the %09d EIN in ALL 7 control positions (ISA13/GS06/ST02/SE02/GE02/IEA02) + consistency.
  * SE01 == the ACTUAL ST..SE segment count (recomputed independently); CTT01 == #IT1 (NOT a magic offset).
  * ISA fixed widths + ISA09 yymmdd while GS04/BIG01 = yyyymmdd, dated off NOW (NOT pickup date).
  * IT1 has NO trailing separator (ends on the dock code) — the 810 is CLEAN (verified vs the .pas).
  * CLEAN MONEY: TDS01 = round(total*10000) integer scale-4 INCLUDING the 1-digit-fraction case legacy
    got wrong (asserts the CORRECT value, and that the legacy off-by-10000 value is ABSENT); IT104 fixed
    scale-4.
  * edge: single manifest / single item (minimal envelope) counts still self-consistent.

Run:  python3 scripts/e2e/test_edi810_build.py     (no DB, no env needed)
"""
import os, sys, importlib.util
from decimal import Decimal

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report  # noqa: E402

PLIB = os.path.normpath(os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "..", "..",
    "docs", "analysis", "edi", "810", "project-library"))


def _load(name, rel):
    spec = importlib.util.spec_from_file_location(name, os.path.join(PLIB, rel, "code.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


edi810 = _load("edi810_lib", "edi810")

# --- a synthetic site row (the dict build_810 takes). element sep '*', sub-element '>', terminator '~'
#     (read-but-never-emitted). dockCode is the 810-specific IT111 field. ---------------------------------
SITE = {
    "abbr": "MAS",                 # ISA02 (%-10 -> 'MAS       ')
    "duns": "969009112",           # ISA04 (%-10), GS02, ISA06 prefix
    "supplierCode": "71930",       # ISA06 suffix -> '969009112-71930' (exactly 15); BIG02
    "tmmDuns": "808369495",        # ISA08 (%-15), GS03
    "ediMode": "P",                # ISA15
    "dockCode": "D01",             # IT111 (ZZ D01)
    "sepElement": "*",
    "sepSubElement": ">",
    "sepSegment": "~",             # read-but-never-emitted (CRLF only)
}

# the file time = NOW, captured once (legacy f810Time:=now). The 810 dates OFF NOW, not the pickup date.
FTIME = {"yyyymmdd": "20260620", "hhmm": "1430"}     # -> ISA09 '260620'; GS04/BIG01 '20260620'
PICKUP = "20260610"                 # the shipments' production date -> DTM*050 + filename (a DIFFERENT date)
EIN = 5692                          # -> %09d '000005692'
EINS = "000005692"


def main():
    print("=" * 88)
    print(" EDI 810 INVOICE outbound — pure build_810 unit test (PORT of EDI810Object.pas T810EDI)")
    print(" (segment bytes DERIVED FROM the .pas; MONEY asserts the CLEAN values — DECISION 810-1/810-2)")
    print("=" * 88)
    rep = Report()

    # ============================================================================================
    # MULTI-MANIFEST / MULTI-ITEM — the non-vacuous byte-exact case.
    #   Manifest 76061857 (starts '7' -> M391 broadcast) : 2 items (parts A, B)
    #   Manifest 80061900 (starts '8' -> M390 hot-call)  : 1 item  (part C)
    # All rows share PICKUP (one invoice file = one pickup date in legacy).
    # ============================================================================================
    detail = [
        {"Manifest": "76061857", "PartNumber": "42600FEK5000", "ShipQty": 80,
         "UnitPrice": Decimal("12.5500"), "PickUpDate": PICKUP},
        {"Manifest": "76061857", "PartNumber": "42600FEK5001", "ShipQty": 12,
         "UnitPrice": Decimal("3.2000"), "PickUpDate": PICKUP},
        {"Manifest": "80061900", "PartNumber": "42600FEK6000", "ShipQty": 100,
         "UnitPrice": Decimal("12.5000"), "PickUpDate": PICKUP},
    ]
    segs = edi810.build_810(detail, SITE, EIN, FTIME)

    # --- the CLEAN money, computed INDEPENDENTLY here (not from the builder) ---------------------
    # total = 80*12.55 + 12*3.20 + 100*12.50 = 1004.00 + 38.40 + 1250.00 = 2292.40
    total = Decimal("80") * Decimal("12.5500") + Decimal("12") * Decimal("3.2000") \
        + Decimal("100") * Decimal("12.5000")
    assert total == Decimal("2292.4000"), total
    # CLEAN TDS01 = round(total*10000) = 22924000 (implied 2292.4000). LEGACY would emit (1-digit-frac
    # path): floattostr(2292.4)='2292.4' -> temp='2292', frac='4' (len 1, NOT padded) -> '22924' = an
    # implied 2.2924, OFF BY 10000x. We assert the CLEAN value AND that the legacy bug value is ABSENT.
    CLEAN_TDS = "22924000"
    LEGACY_BUGGY_TDS = "22924"          # the off-by-10000 the legacy 1-digit-fraction path emitted

    # --- BYTE-EXACT expected segment list — DERIVED FROM EDI810Object.pas (the LEGACY bytes), EXCEPT
    #     IT104/TDS which carry the CLEAN money. Per-segment .pas derivation:
    #   ISA  :139-163 — 16 elements each +fSepElement EXCEPT ISA16(:163 fSepSubElement, no trailing sep).
    #                   ISA09(:149)=formatdatetime('yymmdd',NOW)=yymmdd of FTIME; ISA13(:160)=%9.9d EIN;
    #                   ISA15(:162)=ediMode 'P'; ISA16(:163)='>'.
    #   GS   :180-188 — GS01(:181)='IN'; GS04(:184)=yyyymmdd of NOW; ends +'004010', NO trailing sep.
    #   ST   :205-207 — 'ST*810*'+%9.9d EIN, NO trailing sep.
    #   BIG  :226-228 — 'BIG*'+yyyymmdd NOW+'*'+SiteSupplierCode, NO trailing sep (ends on supplier code).
    #   IT1 LOOP (:241-369):
    #     IT1(A,M1) :289-305 — 'IT1*M391*80*EA*<price>*QT*PN*42600FEK5000*PK*1*ZZ*D01' (NO trailing sep;
    #                          M391 because manifest 76061857 starts '7'). IT104 = CLEAN '12.5500'.
    #     IT1(B,M1) — 'IT1*M391*12*EA*3.2000*QT*PN*42600FEK5001*PK*1*ZZ*D01'.
    #     manifest CHANGES at row C (:268) -> interior break BEFORE IT1(C):
    #       REF :271-273 — 'REF*MK*76061857'   (REF carries the PREVIOUS manifest, lastManifest).
    #       DTM :281-283 — 'DTM*050*20260610'   (the current row's PickUpDate).
    #     IT1(C,M2) — 'IT1*M390*100*EA*12.5000*QT*PN*42600FEK6000*PK*1*ZZ*D01' (M390 because 80061900
    #                  does NOT start '7').
    #     trailing REF :320-322 — 'REF*MK*80061900' (the FINAL manifest).
    #     trailing DTM :328-330 — 'DTM*050*20260610' (lastPickUp).
    #     TDS :337-350 — CLEAN 'TDS*22924000' (DECISION 810-1; legacy bug NOT reproduced).
    #   CTT  :377-378 — 'CTT*3' (#IT1 = 3, fID1Count).
    #   SE   :398-400 — 'SE*<count ST..SE>*'+%9.9d EIN.
    #   GE   :418-420 — 'GE*1*'+%9.9d EIN.  IEA :438-440 — 'IEA*1*'+%9.9d EIN.
    # ST..SE: ST,BIG (2) + IT1(A),IT1(B) (2) + REF,DTM (interior, 2) + IT1(C) (1) + REF,DTM (trailing, 2)
    #         + TDS (1) + CTT (1) + SE (1) = 12 -> SE01.
    expected = [
        "ISA*00*MAS       *01*969009112 *ZZ*969009112-71930*01*808369495      *260620*1430*U*00400*000005692*0*P*>",
        "GS*IN*969009112*808369495*20260620*1430*000005692*X*004010",
        "ST*810*000005692",
        "BIG*20260620*71930",
        "IT1*M391*80*EA*12.5500*QT*PN*42600FEK5000*PK*1*ZZ*D01",
        "IT1*M391*12*EA*3.2000*QT*PN*42600FEK5001*PK*1*ZZ*D01",
        "REF*MK*76061857",                 # interior break — PREVIOUS manifest (.pas:271-273)
        "DTM*050*20260610",                # the pickup date (.pas:281-283)
        "IT1*M390*100*EA*12.5000*QT*PN*42600FEK6000*PK*1*ZZ*D01",
        "REF*MK*80061900",                 # trailing — FINAL manifest (.pas:320-322)
        "DTM*050*20260610",                # trailing pickup (.pas:328-330)
        "TDS*22924000",                    # CLEAN scale-4 integer (DECISION 810-1)
        "CTT*3",
        "SE*12*000005692",
        "GE*1*000005692",
        "IEA*1*000005692",
    ]

    print("\n--- byte-exact full segment list (multi-manifest, multi-item) ---")
    rep.check("segment COUNT matches expected (%d)" % len(expected), len(segs) == len(expected),
              "got %d, expected %d" % (len(segs), len(expected)))
    mismatch = None
    for i in range(min(len(segs), len(expected))):
        if segs[i] != expected[i]:
            mismatch = i
            break
    if mismatch is None and len(segs) == len(expected):
        rep.check("every segment byte-exact vs .pas-derived expected (CLEAN money)", True,
                  "%d segments identical" % len(segs))
    else:
        i = mismatch if mismatch is not None else min(len(segs), len(expected))
        rep.check("every segment byte-exact vs .pas-derived expected (CLEAN money)", False,
                  "first diff @ seg %d:\n      got=%r\n      exp=%r"
                  % (i, segs[i] if i < len(segs) else "<none>",
                     expected[i] if i < len(expected) else "<none>"))

    # --- GS01='IN' + BIG02=SupplierCode (the 810 markers, vs 856's GS01='SH') -------------------
    print("\n--- GS01='IN' (the INVOICE marker) + BIG02 = SupplierCode ---")
    gs = segs[1].split("*")
    big = segs[3].split("*")
    rep.check("GS01 == 'IN' (invoice; NOT the 856's 'SH')", gs[1] == "IN", gs[1])
    rep.check("BIG01 == yyyymmdd of NOW '20260620'", big[1] == FTIME["yyyymmdd"], big[1])
    rep.check("BIG02 == SiteSupplierCode '71930' (the invoice number)", big[2] == "71930", big[2])
    rep.check("BIG has no trailing element (ends on the supplier code)", segs[3] == "BIG*20260620*71930",
              segs[3])

    # --- IT101 M391 (broadcast, manifest '7') vs M390 (hot-call) --------------------------------
    print("\n--- IT101 M391 (manifest starts '7') vs M390 (else) ---")
    it1s = [s for s in segs if s.startswith("IT1*")]
    rep.check("3 IT1 lines", len(it1s) == 3, "%d" % len(it1s))
    rep.check("M-manifest 76061857 ('7' prefix) -> IT101 M391 (broadcast)",
              it1s[0].split("*")[1] == "M391" and it1s[1].split("*")[1] == "M391",
              "%s / %s" % (it1s[0].split("*")[1], it1s[1].split("*")[1]))
    rep.check("manifest 80061900 (not '7') -> IT101 M390 (hot-call)",
              it1s[2].split("*")[1] == "M390", it1s[2].split("*")[1])

    # --- IT1 NO trailing separator (the 810 is CLEAN — ends on the dock code) -------------------
    print("\n--- IT1 ends on the dock code (NO trailing sep — the 810 is CLEAN) ---")
    rep.check("no IT1 ends with a trailing separator", all(not s.endswith("*") for s in it1s),
              "all end on a value")
    rep.check("IT1 ends on the dock code 'D01' (IT111)", all(s.endswith("*D01") for s in it1s),
              "all end *D01")
    rep.check("IT1 has 12 elements (IT1..IT111, no empty interior)",
              all(len(s.split("*")) == 12 and "" not in s.split("*") for s in it1s),
              "12 populated elements")
    # guard the 856 LIN trailing-sep quirk did NOT bleed in.
    rep.check("no '**' (empty interior element) anywhere in the file",
              all("**" not in s for s in segs), "clean — no empty interior elements")

    # --- CLEAN MONEY: TDS01 (incl. the 1-digit-fraction case legacy got wrong) ------------------
    print("\n--- CLEAN MONEY: TDS01 = round(total*10000) scale-4 integer (DECISION 810-1) ---")
    tds = next(s for s in segs if s.startswith("TDS*")).split("*")
    rep.check("TDS01 == CLEAN '22924000' (implied 2292.4000)", tds[1] == CLEAN_TDS, tds[1])
    rep.check("TDS01 != the LEGACY off-by-10000 bug value '22924' (1-digit-frac unpadded)",
              tds[1] != LEGACY_BUGGY_TDS, "legacy buggy value absent")
    rep.check("TDS01 has no decimal point (implied-decimal scale-4)", "." not in tds[1], tds[1])

    # the headline 1-digit-fraction case as a STANDALONE assertion on _tds_total: 1234.5 -> 12345000
    # (CLEAN) vs the legacy 12345 (off by 10000x). This is the bug the build spec calls out explicitly.
    rep.check("_tds_total(1234.5) == '12345000' (CLEAN) — legacy emitted '12345' (off by 10000x)",
              edi810._tds_total(Decimal("1234.5")) == "12345000", edi810._tds_total(Decimal("1234.5")))
    rep.check("_tds_total(12.55) == '125500' (2-digit frac)",
              edi810._tds_total(Decimal("12.55")) == "125500", edi810._tds_total(Decimal("12.55")))
    rep.check("_tds_total(100) == '1000000' (whole dollar) — legacy emitted a malformed string",
              edi810._tds_total(Decimal("100")) == "1000000", edi810._tds_total(Decimal("100")))
    rep.check("_tds_total(0.0001) == '1' (one tick)",
              edi810._tds_total(Decimal("0.0001")) == "1", edi810._tds_total(Decimal("0.0001")))
    rep.check("_tds_total(58093.75) == '580937500' (the D6 EIN-5692 corrected total)",
              edi810._tds_total(Decimal("58093.75")) == "580937500",
              edi810._tds_total(Decimal("58093.75")))

    # --- CLEAN MONEY: IT104 fixed scale-4 (locale-independent) ----------------------------------
    print("\n--- CLEAN MONEY: IT104 = fixed scale-4 decimal (DECISION 810-2) ---")
    rep.check("IT104 for 12.55 == '12.5500' (fixed scale-4, NOT FloatToStr '12.55')",
              it1s[0].split("*")[4] == "12.5500", it1s[0].split("*")[4])
    rep.check("IT104 for 3.20 == '3.2000' (fixed scale-4, NOT '3.2')",
              it1s[1].split("*")[4] == "3.2000", it1s[1].split("*")[4])
    rep.check("IT104 for 12.50 == '12.5000' (fixed scale-4, NOT '12.5')",
              it1s[2].split("*")[4] == "12.5000", it1s[2].split("*")[4])
    rep.check("_it104(Decimal('0')) == '0.0000' (zero scaled, locale-independent)",
              edi810._it104(Decimal("0")) == "0.0000", edi810._it104(Decimal("0")))
    rep.check("_it104(12) == '12.0000' (int -> scale-4)", edi810._it104(12) == "12.0000",
              edi810._it104(12))

    # --- %09d EIN in ALL 7 positions + consistency ----------------------------------------------
    print("\n--- the %09d EIN appears in all 7 control positions + is consistent ---")
    isa = segs[0].split("*"); st = segs[2].split("*")
    se = segs[-3].split("*"); ge = segs[-2].split("*"); iea = segs[-1].split("*")
    rep.check("ISA13 == %09d EIN", isa[13] == EINS, isa[13])
    rep.check("GS06  == %09d EIN", gs[6] == EINS, gs[6])
    rep.check("ST02  == %09d EIN", st[2] == EINS, st[2])
    rep.check("SE02  == %09d EIN", se[2] == EINS, se[2])
    rep.check("GE02  == %09d EIN", ge[2] == EINS, ge[2])
    rep.check("IEA02 == %09d EIN", iea[2] == EINS, iea[2])
    rep.check("all 7 control numbers agree (the TEMA-reject one)",
              len({isa[13], gs[6], st[2], se[2], ge[2], iea[2]}) == 1, isa[13])

    # --- SE01 == actual ST..SE count (recomputed INDEPENDENTLY) ; CTT01 == #IT1 ------------------
    print("\n--- SE01 = ST..SE segment count (independent) ; CTT01 = #IT1 ---")
    st_idx = next(i for i, s in enumerate(segs) if s.startswith("ST*"))
    se_idx = next(i for i, s in enumerate(segs) if s.startswith("SE*"))
    actual_se01 = (se_idx - st_idx) + 1
    rep.check("SE01 == actual ST..SE segment count", int(se[1]) == actual_se01,
              "SE01=%s actual=%d" % (se[1], actual_se01))
    rep.check("SE01 == 12 for this fixture (sanity)", int(se[1]) == 12, se[1])
    ctt = segs[se_idx - 1].split("*")
    rep.check("CTT01 == #IT1 lines (3) — NOT a segment count", int(ctt[1]) == len(it1s) == 3,
              "CTT01=%s #IT1=%d" % (ctt[1], len(it1s)))

    # --- ISA fixed widths + dates off NOW (yymmdd vs yyyymmdd) -----------------------------------
    print("\n--- ISA fixed widths + dates off NOW (ISA09 yymmdd ; GS04/BIG01 yyyymmdd) ---")
    rep.check("ISA02 (abbr) padded to exactly 10", len(isa[2]) == 10 and isa[2] == "MAS       ", isa[2])
    rep.check("ISA06 (DUNS-supplier) padded to exactly 15", len(isa[6]) == 15 and isa[6] == "969009112-71930",
              isa[6])
    rep.check("ISA08 (TMM DUNS) padded to exactly 15", len(isa[8]) == 15, isa[8])
    rep.check("ISA09 == yymmdd of NOW '260620' (century dropped; off NOW not pickup)",
              isa[9] == "260620", isa[9])
    rep.check("GS04 == full yyyymmdd of NOW '20260620'", gs[4] == FTIME["yyyymmdd"], gs[4])
    rep.check("BIG01 == full yyyymmdd of NOW '20260620'", big[1] == FTIME["yyyymmdd"], big[1])
    rep.check("DTM*050 carries the PICKUP date '20260610' (NOT NOW) — the shipment received date",
              all(s == "DTM*050*20260610" for s in segs if s.startswith("DTM*")), "all DTM = pickup")
    rep.check("ISA15 == ediMode 'P' (single char)", isa[15] == "P", isa[15])
    rep.check("ISA16 == sub-element separator '>'", isa[16] == ">", isa[16])

    # --- REF interior-vs-trailing manifest (interior=PREVIOUS, trailing=FINAL) -------------------
    print("\n--- REF*MK manifests: interior carries the PREVIOUS, trailing the FINAL ---")
    refs = [s for s in segs if s.startswith("REF*")]
    rep.check("2 REF*MK segments (1 interior break + 1 trailing)", len(refs) == 2, str(refs))
    rep.check("interior REF carries the PREVIOUS manifest '76061857' (.pas:271-273)",
              refs[0] == "REF*MK*76061857", refs[0])
    rep.check("trailing REF carries the FINAL manifest '80061900' (.pas:320-322)",
              refs[1] == "REF*MK*80061900", refs[1])

    # --- decision: no segment terminator inside any segment (CRLF only) --------------------------
    print("\n--- CRLF-only: no '~' segment terminator inside any segment ---")
    rep.check("no segment contains the segment terminator '~'", all("~" not in s for s in segs),
              "terminator never emitted")
    joined = "\r\n".join(segs) + "\r\n"
    rep.check("CRLF-join + split round-trips to the same segments",
              joined.split("\r\n")[:-1] == segs, "round-trip clean")
    rep.check("the file text ends with a trailing CRLF (legacy Writeln)", joined.endswith("\r\n"))

    # --- filename pattern (recreate: 810 + copy(PickUpDate,5,4) + .txt) --------------------------
    print("\n--- filename ('810' + mmdd of PickUpDate + '.txt') ---")
    fn = edi810._filename_810(PICKUP)
    rep.check("filename '20260610' -> '8100610.txt' (810 + '0610')", fn == "8100610.txt", fn)
    rep.check("filename for 20271225 -> '8101225.txt'", edi810._filename_810("20271225") == "8101225.txt",
              edi810._filename_810("20271225"))

    # ============================================================================================
    # EDGE: single manifest / single item — the minimal envelope, counts still self-consistent.
    # ============================================================================================
    print("\n--- edge: single manifest / single item (minimal envelope) ---")
    one = [{"Manifest": "76061805", "PartNumber": "42600F261100", "ShipQty": 32,
            "UnitPrice": Decimal("9.9900"), "PickUpDate": PICKUP}]
    s1 = edi810.build_810(one, SITE, EIN, FTIME)
    se1 = s1[-3].split("*")
    ctt1 = s1[-4].split("*")
    st1 = next(i for i, s in enumerate(s1) if s.startswith("ST*"))
    se1i = next(i for i, s in enumerate(s1) if s.startswith("SE*"))
    rep.check("single-item SE01 == actual ST..SE count", int(se1[1]) == (se1i - st1) + 1,
              "SE01=%s actual=%d" % (se1[1], (se1i - st1) + 1))
    # ST..SE: ST,BIG(2) + IT1(1) + REF,DTM trailing(2) + TDS(1) + CTT(1) + SE(1) = 8
    rep.check("single-item SE01 == 8", int(se1[1]) == 8, se1[1])
    rep.check("single-item CTT01 == 1 (one IT1)", int(ctt1[1]) == 1, ctt1[1])
    # single-manifest: NO interior break -> exactly 1 REF + 1 DTM (the trailing pair).
    rep.check("single-manifest emits exactly 1 REF + 1 DTM (no interior break)",
              sum(1 for s in s1 if s.startswith("REF*")) == 1
              and sum(1 for s in s1 if s.startswith("DTM*")) == 1, "1 REF, 1 DTM")
    # single-item TDS = 32 * 9.99 = 319.68 -> round(319.68*10000) = 3196800
    rep.check("single-item TDS01 == '3196800' (32 * 9.99 = 319.68, scale-4)",
              next(s for s in s1 if s.startswith("TDS*")) == "TDS*3196800",
              next(s for s in s1 if s.startswith("TDS*")))
    rep.check("single-item EIN still in all positions", s1[0].split("*")[13] == EINS
              and s1[-1].split("*")[2] == EINS)

    # ============================================================================================
    # GUARD: empty feed -> Edi810BuildError (legacy 'No records to process').
    # ============================================================================================
    print("\n--- guard: empty feed raises (legacy 'No records to process') ---")
    raised = False
    try:
        edi810.build_810([], SITE, EIN, FTIME)
    except edi810.Edi810BuildError:
        raised = True
    rep.check("build_810([]) raises Edi810BuildError (no empty envelope emitted)", raised)

    # ============================================================================================
    # ISA over-width hard-truncation (pathological; never fires on real data).
    # ============================================================================================
    print("\n--- ISA over-width hard-truncation (fix-only) ---")
    wide = dict(SITE)
    wide["abbr"] = "ABCDEFGHIJKLMNOP"     # 16 chars -> must truncate to 10
    sw = edi810.build_810(one, wide, EIN, FTIME)
    isaw = sw[0].split("*")
    rep.check("over-wide ISA02 hard-truncated to exactly 10", isaw[2] == "ABCDEFGHIJ" and len(isaw[2]) == 10,
              "%r" % isaw[2])

    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
