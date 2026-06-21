#!/usr/bin/env python3
"""test_edi856_build.py — unit test for the PURE outbound-856 segment builder (edi856.build_856).

Exercises the byte-faithful X12 segment assembly that PORTS the legacy EDI856Object.T856EDI
(EDI856Object.pas) with NO database: edi856/code.py imports nothing and is CPython-importable straight
from the on-disk project library. Twin of test_asn_fanout.py (producer-recipe pattern).

Source truth:
  docs/analysis/edi/856/edi856-wire-format.md          — the byte-exact segment spec (THE target)
  docs/analysis/edi/856/report-edi856-data-analysis.md — the data feed + decisions A-E
  EDI856Object.pas                                     — the original builder

NON-VACUOUS COVERAGE (synthetic header + multi-manifest/multi-item detail):
  * the FULL ordered segment list asserted BYTE-EXACT (every segment string compared to a hand-computed
    expected), proving segment order + content (ISA..IEA, the HL S->O->I chain, PRF/LIN/SN1).
  * the %09d EIN in ALL 7 positions (ISA13, GS06, ST02, SE02, GE02, IEA02 + the BSN02 yyyymmdd+%09d suffix).
  * control-number consistency: ISA13 == GE02 == IEA02 (== GS06 == ST02 == SE02).
  * SE01 == the ACTUAL count of segments ST..SE inclusive (recomputed independently, not the builder's).
  * CTT01 == the HL-segment count only (recomputed independently).
  * ISA fixed widths (ISA02/04 = 10, ISA06/08 = 15) + ISA09 = yymmdd while GS04/BSN03/DTM02 = yyyymmdd.
  * CRLF joining is the DRIVER's job (build_856 returns a list); we assert no segment carries a separator
    terminator and the join is clean.
  * the HL parent chain: S(id1,no parent) -> O(id,parent '1') -> I(id,parent=order id).
  * trailerId default '1234567890' AND a parameterized override.
  * the filename pattern (decision E): 856 + copy(prodDate,4,5) + .txt.
  * edge: single manifest / single item (minimal envelope) counts still self-consistent.

Run:  python3 scripts/e2e/test_edi856_build.py     (no DB, no env needed)
"""
import os, sys, importlib.util

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report  # noqa: E402

PLIB = os.path.normpath(os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "..", "..",
    "docs", "analysis", "edi", "856", "project-library"))


def _load(name, rel):
    spec = importlib.util.spec_from_file_location(name, os.path.join(PLIB, rel, "code.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


edi856 = _load("edi856_lib", "edi856")

# --- a synthetic site row (the dict build_856 takes). Values chosen to exercise ISA widths + the
#     duns-supplier ISA06 join. element sep '*', sub-element '>', segment terminator '~' (NEVER emitted).
SITE = {
    "abbr": "MAS",                 # ISA02 (%-10 -> 'MAS       ')
    "duns": "969009112",           # ISA04 (%-10), GS02, ISA06 prefix
    "supplierCode": "71930",       # ISA06 suffix -> '969009112-71930' (exactly 15)
    "tmmDuns": "808369495",        # ISA08 (%-15), GS03
    "ediMode": "P",                # ISA15
    "deliveryMethodCode": "CV",    # TD5 carrier
    "sepElement": "*",
    "sepSubElement": ">",
    "sepSegment": "~",             # read-but-never-emitted (decision D)
}

PDATE = "20260618"                  # -> ISA09 '260618'; GS04/BSN03/DTM02 '20260618'; filename 856 60618
HHMM = "1430"
EIN = 9069                          # -> %09d '000009069'
EINS = "000009069"


def main():
    print("=" * 84)
    print(" EDI 856 outbound — pure build_856 unit test (PORT of EDI856Object.pas T856EDI)")
    print("=" * 84)
    rep = Report()

    # ============================================================================================
    # MULTI-MANIFEST / MULTI-ITEM — the non-vacuous byte-exact case.
    #   Manifest 76061857 : 2 items (parts A, B)
    #   Manifest 76061858 : 1 item  (part C)
    # ============================================================================================
    detail = [
        {"Manifest": "76061857", "PartNumber": "42600FEK5000", "ShipQty": 80,  "Kanban": "JZUX"},
        {"Manifest": "76061857", "PartNumber": "42600FEK5001", "ShipQty": 12,  "Kanban": "JZUW"},
        {"Manifest": "76061858", "PartNumber": "42600FEK6000", "ShipQty": 900, "Kanban": "JZUY"},
    ]
    segs = edi856.build_856({"prodDate": PDATE}, detail, SITE, EIN, HHMM)

    # --- BYTE-EXACT expected segment list — RE-DERIVED FROM EDI856Object.pas (the LEGACY bytes) -----
    # This array is computed BY READING THE .pas Write/format calls, NOT from the rebuild's sep.join — so
    # the test catches the rebuild DIVERGING from legacy, never blesses it. Per-segment .pas derivation:
    #   ISA  :145-162 — 16 elements each +fSepElement EXCEPT ISA16(:162 fSepSubElement, no trailing sep).
    #                   ISA09(:154) = copy(prodDate,3,6)=yymmdd; ISA15(:161)=ediMode 'P'; ISA16(:162)='>'.
    #   GS   :179-187 — ends +'004010', NO trailing sep.
    #   ST   :204-206 — ends +Format('%9.9d'), NO trailing sep.
    #   BSN  :225-229 — :229 hhmm with '//+fSepElement' commented -> NO trailing sep.
    #   DTM  :248-252 — ends +'ET', NO trailing sep.
    #   HL(S):285-290 — :286 IntToStr(fHID)+fSepElement+fSepElement -> HL02 empty -> 'HL*1**S*1'.
    #   TD1  :294-296 — 'TD1'+fSepElement(:294) THEN +fSepElement(:295), nothing after -> 'TD1**'
    #                   (TWO empty trailing elements). NOT 'TD1*'.
    #   TD5  :298-302 — ends +SiteDeliveryMethodCode('CV'), NO trailing sep.
    #   TD3  :305-308 — ends +'1234567890', NO trailing sep.
    #   HL(O):320-324 — :322 HL02 hardcoded '1' (always parents shipment); ends +'1', NO trailing sep.
    #   PRF  :330-331 — 'PRF'+fSepElement+Manifest+'-'+Manifest, NO trailing sep.
    #   HL(I):338-343 — ends +'0', NO trailing sep.
    #   LIN  :347-352 — :352 appends Kanban +fSepElement -> 'LIN**BP*<part>*RC*<kanban>*' (TRAILING sep,
    #                   one empty trailing element). NOT '...*RC*<kanban>'.
    #   SN1  :355-358 — ends +'PC', NO trailing sep -> 'SN1**<qty>*PC'.
    #   CTT  :386-387 / SE :406-409 / GE :427-429 / IEA :447-449 — all end on the value, NO trailing sep.
    # HL ids: S=1, O(57)=2, I(A)=3, I(B)=4, O(58)=5, I(C)=6.  HL count = 6 -> CTT01.
    # ST..SE: ST,BSN,DTM (3) + [HL(S),TD1,TD5,TD3]=4 + [HL O57,PRF]=2 + 2*[HL I,LIN,SN1]=6
    #         + [HL O58,PRF]=2 + 1*[HL I,LIN,SN1]=3 + CTT(1) + SE(1) = 22 -> SE01.
    expected = [
        # ISA — positional, fixed widths. abbr 'MAS'->10, duns->10, '969009112-71930'->15, tmmDuns->15.
        "ISA*00*MAS       *01*969009112 *ZZ*969009112-71930*01*808369495      *260618*1430*U*00400*000009069*0*P*>",
        "GS*SH*969009112*808369495*20260618*1430*000009069*X*004010",
        "ST*856*000009069",
        "BSN*00*20260618000009069*20260618*1430",
        "DTM*011*20260618*1430*ET",
        # HL hierarchy
        "HL*1**S*1",
        "TD1**",                                       # :294-296 — TWO empty elements (legacy)
        "TD5*B*25*00000*CV",
        "TD3*TL**1234567890",
        "HL*2*1*O*1",
        "PRF*76061857-76061857",
        "HL*3*2*I*0",
        "LIN**BP*42600FEK5000*RC*JZUX*",               # :352 — TRAILING sep after kanban (legacy)
        "SN1**80*PC",
        "HL*4*2*I*0",
        "LIN**BP*42600FEK5001*RC*JZUW*",               # :352 — TRAILING sep after kanban (legacy)
        "SN1**12*PC",
        "HL*5*1*O*1",
        "PRF*76061858-76061858",
        "HL*6*5*I*0",
        "LIN**BP*42600FEK6000*RC*JZUY*",               # :352 — TRAILING sep after kanban (legacy)
        "SN1**900*PC",
        "CTT*6",
        "SE*22*000009069",
        "GE*1*000009069",
        "IEA*1*000009069",
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
        rep.check("every segment byte-exact vs hand-computed expected", True,
                  "%d segments identical" % len(segs))
    else:
        i = mismatch if mismatch is not None else min(len(segs), len(expected))
        rep.check("every segment byte-exact vs hand-computed expected", False,
                  "first diff @ seg %d:\n      got=%r\n      exp=%r"
                  % (i, segs[i] if i < len(segs) else "<none>",
                     expected[i] if i < len(expected) else "<none>"))

    # --- %09d EIN in ALL 7 positions ------------------------------------------------------------
    print("\n--- the %09d EIN appears in all 7 control positions ---")
    isa = segs[0].split("*")
    gs = segs[1].split("*")
    st = segs[2].split("*")
    bsn = segs[3].split("*")
    se = segs[-3].split("*")
    ge = segs[-2].split("*")
    iea = segs[-1].split("*")
    rep.check("ISA13 == %09d EIN", isa[13] == EINS, isa[13])
    rep.check("GS06  == %09d EIN", gs[6] == EINS, gs[6])
    rep.check("ST02  == %09d EIN", st[2] == EINS, st[2])
    rep.check("SE02  == %09d EIN", se[2] == EINS, se[2])
    rep.check("GE02  == %09d EIN", ge[2] == EINS, ge[2])
    rep.check("IEA02 == %09d EIN", iea[2] == EINS, iea[2])
    rep.check("BSN02 == yyyymmdd + %09d EIN (17 chars)",
              bsn[2] == PDATE + EINS and len(bsn[2]) == 17, "%r len=%d" % (bsn[2], len(bsn[2])))

    # --- control-number consistency (ISA13 == GE02 == IEA02, and the rest echo it) --------------
    print("\n--- control-number consistency (the TEMA-reject one) ---")
    rep.check("ISA13 == GE02 == IEA02 (interchange/group/interchange-trailer agree)",
              isa[13] == ge[2] == iea[2], "%s / %s / %s" % (isa[13], ge[2], iea[2]))
    rep.check("GS06 == GE02 (group control number echoes its trailer)", gs[6] == ge[2],
              "%s == %s" % (gs[6], ge[2]))
    rep.check("ST02 == SE02 (transaction-set control number echoes its trailer)", st[2] == se[2],
              "%s == %s" % (st[2], se[2]))

    # --- SE01 == actual ST..SE segment count (recomputed INDEPENDENTLY of the builder) ----------
    print("\n--- SE01 = count of segments ST..SE inclusive (independently recomputed) ---")
    st_idx = next(i for i, s in enumerate(segs) if s.startswith("ST*"))
    se_idx = next(i for i, s in enumerate(segs) if s.startswith("SE*"))
    actual_se01 = (se_idx - st_idx) + 1     # ST..SE inclusive (SE counts itself)
    rep.check("SE01 == actual ST..SE segment count", int(se[1]) == actual_se01,
              "SE01=%s actual=%d" % (se[1], actual_se01))
    rep.check("SE01 == 22 for this fixture (sanity)", int(se[1]) == 22, se[1])

    # --- CTT01 == HL count only (independently recomputed) --------------------------------------
    print("\n--- CTT01 = HL-segment count only (independently recomputed) ---")
    ctt = segs[se_idx - 1].split("*")
    hl_count = sum(1 for s in segs if s.startswith("HL*"))
    rep.check("CTT01 == count of HL segments only (S + orders + items)", int(ctt[1]) == hl_count,
              "CTT01=%s HLcount=%d" % (ctt[1], hl_count))
    rep.check("CTT01 == 6 for this fixture (1 S + 2 O + 3 I)", int(ctt[1]) == 6, ctt[1])
    # CTT01 is NOT the line/detail count (which is 3) — guard against the easy mistake.
    rep.check("CTT01 (6) != detail-line count (3) — HL count, not line count", int(ctt[1]) != 3)

    # --- ISA fixed widths + the yymmdd vs yyyymmdd date split -----------------------------------
    print("\n--- ISA fixed widths + ISA09 yymmdd while GS04/BSN03/DTM02 yyyymmdd ---")
    rep.check("ISA02 (abbr) padded to exactly 10", len(isa[2]) == 10 and isa[2] == "MAS       ",
              "%r" % isa[2])
    rep.check("ISA04 (DUNS) padded to exactly 10", len(isa[4]) == 10 and isa[4] == "969009112 ",
              "%r" % isa[4])
    rep.check("ISA06 (DUNS-supplier) padded to exactly 15", len(isa[6]) == 15 and isa[6] == "969009112-71930",
              "%r" % isa[6])
    rep.check("ISA08 (TMM DUNS) padded to exactly 15", len(isa[8]) == 15 and isa[8] == "808369495      ",
              "%r" % isa[8])
    rep.check("ISA09 == yymmdd (century dropped) '260618'", isa[9] == "260618", isa[9])
    rep.check("GS04 == full yyyymmdd '20260618'", gs[4] == PDATE, gs[4])
    rep.check("BSN03 == full yyyymmdd '20260618'", bsn[3] == PDATE, bsn[3])
    rep.check("DTM02 == full yyyymmdd '20260618'", segs[4].split("*")[2] == PDATE, segs[4])
    rep.check("ISA16 == sub-element separator '>'", isa[16] == ">", isa[16])
    rep.check("ISA15 == ediMode 'P'", isa[15] == "P", isa[15])

    # --- the HL parent chain S -> O -> I --------------------------------------------------------
    print("\n--- HL S->O->I parent chain ---")
    hls = [s.split("*") for s in segs if s.startswith("HL*")]
    s_hl = hls[0]
    rep.check("Shipment HL = HL*1**S*1 (id 1, NO parent, level S)",
              s_hl[1] == "1" and s_hl[2] == "" and s_hl[3] == "S", str(s_hl))
    orders = [h for h in hls if h[3] == "O"]
    items = [h for h in hls if h[3] == "I"]
    rep.check("2 Order HLs, 3 Item HLs", len(orders) == 2 and len(items) == 3,
              "O=%d I=%d" % (len(orders), len(items)))
    rep.check("every Order HL parents the shipment (HL02 hardcoded '1', Trap 3)",
              all(h[2] == "1" for h in orders), str([h[2] for h in orders]))
    # items 3,4 -> parent 2 (order 57); item 6 -> parent 5 (order 58)
    item_parents = {h[1]: h[2] for h in items}
    rep.check("Item HLs parent their OWN order (3,4->2 ; 6->5)",
              item_parents == {"3": "2", "4": "2", "6": "5"}, str(item_parents))

    # --- PRF / LIN / SN1 content ----------------------------------------------------------------
    print("\n--- PRF / LIN / SN1 content ---")
    prfs = [s for s in segs if s.startswith("PRF*")]
    rep.check("PRF = 'PRF*<manifest>-<manifest>' (the '-' is data)",
              prfs == ["PRF*76061857-76061857", "PRF*76061858-76061858"], str(prfs))
    lins = [s for s in segs if s.startswith("LIN*")]
    # legacy LIN (.pas:347-352) ends in a TRAILING element separator after the kanban.
    rep.check("LIN = 'LIN**BP*<part>*RC*<kanban>*' (LIN01 empty + TRAILING sep, .pas:352)",
              lins[0] == "LIN**BP*42600FEK5000*RC*JZUX*", lins[0])
    rep.check("every LIN carries the trailing element separator (.pas:352)",
              all(s.endswith("*") for s in lins), str(lins))
    sn1s = [s for s in segs if s.startswith("SN1*")]
    rep.check("SN1 = 'SN1**<qty>*PC' (SN101 empty; qty no leading zeros; NO trailing sep, .pas:358)",
              sn1s == ["SN1**80*PC", "SN1**12*PC", "SN1**900*PC"], str(sn1s))

    # --- ANTI-REGRESSION: these are the LEGACY bytes, NOT the rebuild's earlier (wrong) ones --------
    # The prior build emitted 'TD1*' (one empty element) and 'LIN**BP*...*RC*<kanban>' (no trailing sep);
    # both diverged from EDI856Object.pas by one byte on EVERY 856. Assert the legacy bytes explicitly
    # AND that the old wrong bytes are NOT present, so a regression to the rebuild's model fails loudly.
    print("\n--- anti-regression: TD1**/LIN trailing-sep are the LEGACY bytes (.pas), not the rebuild's ---")
    td1 = next(s for s in segs if s.startswith("TD1*"))
    rep.check("TD1 == 'TD1**' (two empty elements, .pas:294-296) — NOT the rebuild's 'TD1*'",
              td1 == "TD1**" and td1 != "TD1*", td1)
    rep.check("no LIN equals the rebuild's old trailing-sep-LESS form",
              all(s != "LIN**BP*42600FEK5000*RC*JZUX" for s in lins),
              "old wrong form absent")

    # --- no segment carries a segment terminator (decision D — CRLF only) -----------------------
    print("\n--- decision D: no '~' segment terminator inside any segment ---")
    rep.check("no segment contains the segment terminator '~' (CRLF-only emission)",
              all("~" not in s for s in segs), "terminator never emitted")
    # the driver joins with CRLF; assert a clean join round-trips back to the same list.
    joined = "\r\n".join(segs) + "\r\n"
    rep.check("CRLF-join + split round-trips to the same segments",
              joined.split("\r\n")[:-1] == segs, "round-trip clean")
    rep.check("the file text ends with a trailing CRLF (legacy Writeln on the last segment)",
              joined.endswith("\r\n"))

    # --- trailerId: default + parameterized override --------------------------------------------
    print("\n--- TD3 trailer id: default literal + parameterized override ---")
    td3 = next(s for s in segs if s.startswith("TD3*"))
    rep.check("default TD3 trailer id == legacy literal '1234567890'", td3 == "TD3*TL**1234567890", td3)
    segs2 = edi856.build_856({"prodDate": PDATE}, detail, SITE, EIN, HHMM, trailerId="9998887770")
    td3b = next(s for s in segs2 if s.startswith("TD3*"))
    rep.check("parameterized trailer id overrides the literal", td3b == "TD3*TL**9998887770", td3b)

    # --- filename pattern (decision E) ----------------------------------------------------------
    print("\n--- filename (decision E: CREATE pattern 856 + copy(prodDate,4,5) + .txt) ---")
    fn = edi856._filename_856(PDATE)
    rep.check("filename '20260618' -> '85660618.txt' (856 + '60618')", fn == "85660618.txt", fn)
    rep.check("filename for 20271225 -> '85671225.txt' (1-digit year)",
              edi856._filename_856("20271225") == "85671225.txt", edi856._filename_856("20271225"))

    # ============================================================================================
    # EDGE: single manifest / single item — the minimal envelope, counts still self-consistent.
    # ============================================================================================
    print("\n--- edge: single manifest / single item (minimal envelope) ---")
    one = [{"Manifest": "76061805", "PartNumber": "42600F261100", "ShipQty": 32, "Kanban": "JZV9"}]
    s1 = edi856.build_856({"prodDate": PDATE}, one, SITE, EIN, HHMM)
    se1 = s1[-3].split("*")
    ctt1 = s1[-4].split("*")
    st1 = next(i for i, s in enumerate(s1) if s.startswith("ST*"))
    se1i = next(i for i, s in enumerate(s1) if s.startswith("SE*"))
    rep.check("single-item SE01 == actual ST..SE count", int(se1[1]) == (se1i - st1) + 1,
              "SE01=%s actual=%d" % (se1[1], (se1i - st1) + 1))
    # HL: S(1) + O(2) + I(3) = 3 ; ST..SE = ST,BSN,DTM(3)+4(HL/TD)+2(O,PRF)+3(I,LIN,SN1)+CTT+SE = 14
    rep.check("single-item CTT01 == 3 (S + 1 O + 1 I)", int(ctt1[1]) == 3, ctt1[1])
    rep.check("single-item SE01 == 14", int(se1[1]) == 14, se1[1])
    rep.check("single-item EIN still in all positions", s1[0].split("*")[13] == EINS
              and s1[-1].split("*")[2] == EINS)

    # ============================================================================================
    # ISA over-width hard-truncation (Trap 6) — pathological input; never fires on real data.
    # ============================================================================================
    print("\n--- ISA over-width hard-truncation (Trap 6, fix-only) ---")
    wide = dict(SITE)
    wide["abbr"] = "ABCDEFGHIJKLMNOP"     # 16 chars -> must truncate to 10
    sw = edi856.build_856({"prodDate": PDATE}, one, wide, EIN, HHMM)
    isaw = sw[0].split("*")
    rep.check("over-wide ISA02 hard-truncated to exactly 10", isaw[2] == "ABCDEFGHIJ" and len(isaw[2]) == 10,
              "%r" % isaw[2])

    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
