#!/usr/bin/env python3
"""test_edi_inbound_build.py — PURE-PARSER unit test for the M1 inbound 997/824 processor
(docs/analysis/edi/inbound/project-library/edi_inbound/code.py).

WHAT THIS PROVES
----------------
The PURE, I/O-free parsers — parse_997, parse_824, ak9_to_status, _functional_to_eintype, _split_isa,
_detect_type — are FAITHFUL to the legacy EDIUpload.pas element offsets (the parity oracle) AND apply the
deliberate M1 FIXES on top:

  * parse_997 reads the AK1 functional id (copy(fcl,5,2) chars 5-6) + the EIN (copy(fcl,8,9) chars 8-16)
    and the AK9 status — but, unlike the legacy "char 5 of the NEXT line" (EDIUpload.pas:196-197), it
    scans forward to the REAL AK9, TOLERATING interleaved AK2/AK3/AK4 detail (Q6). Multi-AK1 (multiple
    functional groups per 997) -> multiple acks.
  * ak9_to_status maps A/E/P/R distinctly + M/W/X/garbage/empty -> R (Q6 — a strict SUPERSET of the
    legacy binary A-vs-rest collapse, which stored the raw char so E/P render blank downstream).
  * parse_824 reads each NTE by the legacy FIXED offsets: errorText copy(fcl,9,50) chars 9-58, manifest
    copy(fcl,60,8) chars 60-67, part copy(fcl,69,12) chars 69-80 (spec §3.1). Multi-NTE -> multiple refs.

HONEST VERIFICATION (no golden inbound sample)
----------------------------------------------
There is NO captured real TEMA 997/824 on disk yet (the data-dependent offset confirms in spec §2.3
#1 / §3.1 remain pending a golden). So this test does NOT claim byte-parity with a real TEMA file — it
asserts the parse is FAITHFUL TO THE LEGACY EDIUpload.pas COPY-OFFSETS (the oracle the legacy itself
uses) on fixtures BUILT to those exact offsets, plus the structural fixes. The fixtures are derived from
EDIUpload.pas:194-197 (997) and :283-285 (824) + the outbound 856 ISA shape (the inbound ISA mirrors it
with TEMA as sender). When a real inbound sample arrives, re-confirm the offsets against it.

Non-vacuous: every assertion compares parsed output to an INDEPENDENTLY-constructed expected value (the
fixtures are built field-by-field at named char positions, so a wrong slice in code.py fails here).

Run:  python3 scripts/e2e/test_edi_inbound_build.py
"""
import os
import sys
import importlib.util

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report          # noqa: E402

CODE_PY = os.path.normpath(os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "..", "..",
    "docs", "analysis", "edi", "inbound", "project-library", "edi_inbound", "code.py"))


def load_pure():
    """Import edi_inbound/code.py as a pure module (no `system` injected — the pure parsers reference no
    gateway global at import or call time, which is the whole point)."""
    spec = importlib.util.spec_from_file_location("edi_inbound_pure", CODE_PY)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


# --- fixture builders: construct each segment at the EXACT legacy char offsets ----------------------

def build_ak1(fnId, ein):
    """AK1 at the legacy offsets: 'AK1*' (4 chars) + fnId at chars 5-6 + '*' at char 7 + ein %09d at
    chars 8-16. So copy(fcl,5,2)=fnId, copy(fcl,8,9)=ein (EDIUpload.pas:194-195)."""
    return "AK1*" + ("%-2s" % fnId)[:2] + "*" + ("%09d" % int(ein))


def build_ak9(code):
    """AK9 at the legacy offset: 'AK9*' (4 chars) + code at char 5. copy(fcl,5,1)=code (:197)."""
    return "AK9*" + (str(code) if code else " ") + "*1*1*0"


def build_nte(errText, manifest, part):
    """NTE at the legacy FIXED offsets (regardless of '*' positions — the legacy slices fixed widths):
      chars 1-8  = 'NTE*GEN*' (seg + a 2-element prefix, 8 chars)
      chars 9-58 = errText left-justified to 50  (copy(fcl,9,50))
      char  59   = '*'
      chars 60-67= manifest left-justified to 8  (copy(fcl,60,8))
      char  68   = '*'
      chars 69-80= part left-justified to 12     (copy(fcl,69,12))
    """
    line = "NTE*GEN*"                               # 8 chars (1-8)
    line += ("%-50s" % errText)[:50]                # 9-58
    line += "*"                                     # 59
    line += ("%-8s" % manifest)[:8]                 # 60-67
    line += "*"                                     # 68
    line += ("%-12s" % part)[:12]                   # 69-80
    return line


def build_inbound_isa(tmmDuns):
    """The inbound ISA — TEMA as SENDER. We PLACE the trading-partner DUNS at split index 4 (delSL[4])
    to match how the code reads it (EDIUpload.pas:73-78 reads delSL[4]).

    !!! HONESTY CAVEAT — UNPROVABLE WITHOUT A GOLDEN INBOUND ISA (adversary BLOCKER-1 / KEY FINDING #2):
    this fixture HARD-CODES the TMM DUNS into the delSL[4] slot, so the DUNS-guard assertions below are
    SELF-CONSISTENT, NOT independent proof of which X12 element a REAL inbound TEMA ISA uses. They prove
    the guard MECHANICS (a value at delSL[4] matching VC_TMM_DUNS routes; a non-matching one quarantines)
    — they do NOT, and cannot, prove that delSL[4] is the right element. On OUR OWN OUTBOUND ISA
    (EDI856Object/EDI810Object) delSL[4]=ISA04 carries OUR SiteDUNS and the TMM DUNS sits at delSL[8]/
    ISA08; so the inbound element MUST be confirmed against a REAL captured TEMA file before the routing
    is trusted (same byte-parity-pending-golden stance as the 856/810). The CODE is legacy-faithful (it
    matches VC_TMM_DUNS exactly as AD_GetSiteTMMDUNS does); the ELEMENT INDEX is the open golden gap.

    Shape: ISA*00*<10>*01*<tmmDuns 10>*ZZ*<15>*01*<15>*<yymmdd>*<hhmm>*U*00400*<9>*0*P*>."""
    return ("ISA*00*          *01*" + ("%-10s" % tmmDuns)[:10] +
            "*ZZ*               *01*               *260618*1430*U*00400*000000777*0*P*>")


def main():
    print("=" * 92)
    print(" edi_inbound PURE parsers — faithful to EDIUpload.pas offsets (oracle) + the Q6 fixes")
    print(" (no golden TEMA sample; offsets are the legacy copy() map — see docstring)")
    print("=" * 92)
    rep = Report()
    m = load_pure()

    # ============================================================================================
    # (A) ak9_to_status — A/E/P/R distinct; M/W/X/garbage/empty -> R (Q6 superset)
    # ============================================================================================
    print("\n--- (A) ak9_to_status (Q6: A/E/P/R distinct, else -> R) ---")
    cases = [("A", "A"), ("E", "E"), ("P", "P"), ("R", "R"),
             ("M", "R"), ("W", "R"), ("X", "R"), ("Z", "R"),
             ("", "R"), (None, "R"), ("a", "A"), ("e", "E"), (" A", "A")]
    for code, exp in cases:
        got = m.ak9_to_status(code)
        rep.check("ak9_to_status(%r) == %r" % (code, exp), got == exp, "got %r" % got)
    rep.check("E and P are DISTINCT statuses (not collapsed to R like legacy)",
              m.ak9_to_status("E") == "E" and m.ak9_to_status("P") == "P",
              "E->%s P->%s" % (m.ak9_to_status("E"), m.ak9_to_status("P")))

    # ============================================================================================
    # (B) parse_997 — AK1 offsets, multi-group, AK2/AK3/AK4 tolerance, A/E/P/R
    # ============================================================================================
    print("\n--- (B) parse_997 (AK1 chars 5-6/8-16; AK9 char 5; AK2/3/4 tolerated; multi-group) ---")

    # (B1) single AK1*SH accept, no detail between AK1 and AK9 (the simple legacy case).
    t = "\r\n".join([build_inbound_isa("000000011"), "GS*FA*X*Y*20260618*1430*9101*X*004010",
                     "ST*997*0001", build_ak1("SH", 9101), build_ak9("A"),
                     "SE*4*0001", "GE*1*9101", "IEA*1*000000777"])
    acks = m.parse_997(t)
    rep.check("single AK1*SH*9101 + AK9*A -> one ack {SH, 9101, 'A'}",
              acks == [{"functionalCode": "SH", "ein": 9101, "ak9": "A"}], repr(acks))

    # (B2) single AK1*IN reject (routes to invoice via the else branch).
    t = "\r\n".join([build_inbound_isa("000000011"), "ST*997*0001",
                     build_ak1("IN", 9102), build_ak9("R"), "SE*3*0001"])
    acks = m.parse_997(t)
    rep.check("AK1*IN*9102 + AK9*R -> {IN, 9102, 'R'}",
              acks == [{"functionalCode": "IN", "ein": 9102, "ak9": "R"}], repr(acks))

    # (B3) AK2/AK3/AK4 detail BETWEEN AK1 and AK9 -> the legacy would read AK2's char 5; we read the
    #      REAL AK9 (the deliberate fix). Build AK2/AK3/AK4 whose char-5 is NOT the true ack code.
    t = "\r\n".join([build_inbound_isa("000000011"), "ST*997*0001",
                     build_ak1("SH", 9105),
                     "AK2*856*0001",            # char 5 = '8' (would be mis-read by legacy)
                     "AK3*HL*7**8",             # char 5 = 'H'
                     "AK4*1*350*1",             # char 5 = '1'
                     build_ak9("E"),            # the REAL ack code = 'E'
                     "SE*7*0001"])
    acks = m.parse_997(t)
    rep.check("AK2/AK3/AK4 between AK1 and AK9 -> reads the REAL AK9 ('E'), not the next line "
              "(the legacy fragility fix)",
              acks == [{"functionalCode": "SH", "ein": 9105, "ak9": "E"}], repr(acks))
    rep.check("  ...and that 'E' maps to a DISTINCT status (not blank/R)",
              m.ak9_to_status(acks[0]["ak9"]) == "E", m.ak9_to_status(acks[0]["ak9"]))

    # (B4) MULTI-GROUP: two AK1 loops in one 997 -> two acks, correct EIN + code each (spec §2.2).
    t = "\r\n".join([build_inbound_isa("000000011"), "ST*997*0001",
                     build_ak1("SH", 9101), build_ak9("A"),
                     build_ak1("IN", 9102), build_ak9("R"),
                     build_ak1("SH", 9103),
                     "AK2*856*0002", build_ak9("P"),     # detail + partial on the 3rd group
                     "SE*9*0001"])
    acks = m.parse_997(t)
    rep.check("multi-group 997 -> 3 acks in file order with correct EIN/code each",
              acks == [{"functionalCode": "SH", "ein": 9101, "ak9": "A"},
                       {"functionalCode": "IN", "ein": 9102, "ak9": "R"},
                       {"functionalCode": "SH", "ein": 9103, "ak9": "P"}], repr(acks))
    rep.check("  ...the P (partial) group maps to 'P' (distinct), survives the AK2 between AK1 and AK9",
              m.ak9_to_status(acks[2]["ak9"]) == "P", repr(acks[2]))

    # (B5) the EIN is parsed at chars 8-16 EXACTLY (a 9-digit %09d), as an int.
    t = "\r\n".join(["ST*997*0001", build_ak1("SH", 123456789), build_ak9("A")])
    acks = m.parse_997(t)
    rep.check("EIN read from chars 8-16 (copy(fcl,8,9)) as int 123456789",
              acks == [{"functionalCode": "SH", "ein": 123456789, "ak9": "A"}], repr(acks))

    # (B6) no AK1 groups (envelope-only body) -> empty list, NOT an error.
    rep.check("a 997 with no AK1 groups -> [] (nothing to do, not an error)",
              m.parse_997("ISA*00*x\r\nST*997*0001\r\nSE*1*0001") == [], "[]")

    # (B7) a non-integer EIN raises EdiInboundError (a malformed AK1, not silently coerced).
    bad = "ST*997*0001\r\nAK1*SH*BADCTRLNO\r\n" + build_ak9("A")
    try:
        m.parse_997(bad)
        rep.check("malformed AK1 control # raises EdiInboundError", False, "did NOT raise")
    except m.EdiInboundError:
        rep.check("malformed AK1 control # raises EdiInboundError", True, "raised")

    # (B8) functional-id routing: 'SH'->SH (ASN); anything else (incl 'IN','XX','') -> 'IN' (invoice),
    #      matching UPDATE_EINStatus's `if @EINType='SH' ... else` (the else is the invoice default).
    rep.check("_functional_to_eintype: SH->SH, IN->IN, XX->IN, ''->IN (the legacy else-default)",
              (m._functional_to_eintype("SH"), m._functional_to_eintype("IN"),
               m._functional_to_eintype("XX"), m._functional_to_eintype("")) == ("SH", "IN", "IN", "IN"),
              "SH/IN/XX/'' routing")

    # ============================================================================================
    # (C) parse_824 — NTE fixed offsets, multi-NTE, SE-trailer termination
    # ============================================================================================
    print("\n--- (C) parse_824 (NTE chars 9-58 / 60-67 / 69-80; multi-NTE; SE ends) ---")

    refs_in = [("SHORT SHIP QTY", "00123456", "478930223000"),
               ("PART NOT ON ASN", "00123457", "999999999999"),
               ("QTY OVER MANIFEST", "00123458", "ABC123XYZ000")]
    t = "\r\n".join([build_inbound_isa("000000011"), "GS*AG*X*Y*20260618*1430*777*X*004010",
                     "ST*824*0001", "BGN*11*REF*20260618"] +
                    [build_nte(e, mf, p) for (e, mf, p) in refs_in] +
                    ["SE*8*0001", "GE*1*777", "IEA*1*000000777"])
    parsed = m.parse_824(t)
    rep.check("parse_824 count == number of NTE lines (3)", parsed["count"] == 3,
              "count=%d" % parsed["count"])
    expected_refs = [{"errorText": e, "manifest": mf, "part": p} for (e, mf, p) in refs_in]
    rep.check("parse_824 refs == the field-by-field expected (errorText@9-58, manifest@60-67, part@69-80)",
              parsed["refs"] == expected_refs, repr(parsed["refs"]))
    for idx, (e, mf, p) in enumerate(refs_in):
        r = parsed["refs"][idx]
        rep.check("  NTE[%d] manifest exactly %r (chars 60-67)" % (idx, mf), r["manifest"] == mf,
                  r["manifest"])
        rep.check("  NTE[%d] part exactly %r (chars 69-80)" % (idx, p), r["part"] == p, r["part"])
        rep.check("  NTE[%d] errorText exactly %r (chars 9-58)" % (idx, e), r["errorText"] == e,
                  r["errorText"])

    # (C2) the SE trailer ends the NTE walk (a NTE-looking line AFTER SE is NOT collected).
    t = "\r\n".join(["ST*824*0001", build_nte("REAL ERR", "00111111", "PARTAAA00000"),
                     "SE*2*0001", build_nte("AFTER SE", "00222222", "PARTBBB00000")])
    parsed = m.parse_824(t)
    rep.check("parse_824 stops at the SE trailer (post-SE NTE ignored)",
              parsed["count"] == 1 and parsed["refs"][0]["manifest"] == "00111111",
              "count=%d" % parsed["count"])

    # (C3) an 824 with zero NTE -> {refs:[], count:0} (driver has nothing to flag; not an error).
    parsed = m.parse_824("ST*824*0001\r\nBGN*11*REF*20260618\r\nSE*2*0001")
    rep.check("an 824 with no NTE -> count 0, empty refs (not an error)",
              parsed == {"refs": [], "count": 0}, repr(parsed))

    # ============================================================================================
    # (D) envelope helpers — _detect_type (ST copy(fcl,4,3)) + _split_isa (delSL[4])
    # ============================================================================================
    print("\n--- (D) envelope: _detect_type (ST chars 4-6) + _split_isa delSL[4] ---")
    rep.check("_detect_type finds '997' from the ST segment (copy(fcl,4,3))",
              m._detect_type(m._lines("ISA*00*x\r\nGS*FA*x\r\nST*997*0001")) == "997", "997")
    rep.check("_detect_type finds '824'",
              m._detect_type(m._lines("ISA*00*x\r\nGS*AG*x\r\nST*824*0001")) == "824", "824")
    rep.check("_detect_type finds '830' (an M2 type — detected so the driver can quarantine it)",
              m._detect_type(m._lines("ISA*00*x\r\nGS*PS*x\r\nST*830*0001")) == "830", "830")
    isa = build_inbound_isa("000000011")
    delSL = m._split_isa(isa)
    rep.check("_split_isa: delSL[0]=='ISA' (the splitString helper shape, EDIUpload.pas:473-493)",
              delSL[0] == "ISA", delSL[0])
    # MECHANICS ONLY: this confirms the split reads index 4 of the SAME slot the fixture wrote — it does
    # NOT prove a real inbound TEMA ISA puts the routing DUNS at index 4 (see build_inbound_isa caveat /
    # adversary BLOCKER-1). The element index is golden-pending; only the split+match plumbing is proven.
    rep.check("_split_isa: delSL[4] (stripped) == the value the fixture placed there '000000011' "
              "(MECHANICS; element index pending a golden inbound ISA — NOT proof of the X12 element)",
              delSL[4].strip() == "000000011", repr(delSL[4]))

    print("")
    return rep.summary_exit()


if __name__ == "__main__":
    sys.exit(main())
