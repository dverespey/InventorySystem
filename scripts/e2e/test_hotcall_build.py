#!/usr/bin/env python3
"""test_hotcall_build.py — PURE unit test for the hot-call ASN producer (the P12/P13/P14 punch-list
follow-on): the 8HC-vs-856 filename branch (edi856._filename_856) + the create_hotcall_asn pure bits
(hotcall.computeHotcallDetails + the manifest/qty validators). NO database — both modules import nothing
and are CPython-importable straight from the on-disk project library.

Source truth (every EXPECTED value is derived FROM the legacy .pas, never from the rebuild):
  MainMenu.pas:2691-2771  — the OPERATIONAL 856 sender (ResendMarkedEDIsClick, the C->S-flip send path):
                            the 8HC vs 856 filename switch (:2715/2718/2722-2724) + the per-batch y counter
                            (init 1 :2702, INC only in the hot-call branch :2724). THE P13 target bytes —
                            re-anchored from the recreate button (ASNInvoice) to the OPERATIONAL sender.
  HotCallEntry.pas        — the "One Cycle Entry" form: manifest >=8 (:157), qty numeric>0 (:184-197),
                            part-required (:204-218), per-row detail (:258-285), the @QTY garbage (:246-247)
  docs/analysis/edi/hotcall-coverage-analysis.md — the P12/P14 analysis

SELF-REFERENTIAL-TEST DISCIPLINE (memory feedback-self-referential-test-discipline, agent retro R15/R16):
  * The 8HC EXPECTED filename string ('8HC606181COROLLA.txt') is HAND-TRANSCRIBED from MainMenu.pas:2723's
    literal — computed in this test by an INDEPENDENT reimplementation (_legacy_8hc_from_pas) that mirrors
    the Pascal `'8HC'+copy(PickupDate,4,5)+IntToStr(y)+LineName+'.txt'` directly — NOT by calling the
    rebuild's _filename_856. The test then asserts the REBUILD's _filename_856 equals that independent
    legacy derivation. A test that fed the rebuild's own output back as the expectation would prove only
    self-consistency.
  * NON-VACUITY is proven explicitly: a deliberately-WRONG reimplementation of the filename (the pre-fix
    rebuild that omitted LineName / used the wrong offset) is run through the SAME assertion and MUST fail
    it (so we know the assertion can fail on a wrong rebuild). See the 'non-vacuity' block.

Run:  python3 scripts/e2e/test_hotcall_build.py     (no DB, no env needed)
"""
import os, sys, importlib.util

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report  # noqa: E402

ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".."))
EDI856_PY = os.path.join(ROOT, "docs", "analysis", "edi", "856", "project-library", "edi856", "code.py")
HOTCALL_PY = os.path.join(ROOT, "docs", "analysis", "edi", "project-library", "hotcall", "code.py")


def _load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


edi856 = _load("edi856_lib", EDI856_PY)
hotcall = _load("hotcall_lib", HOTCALL_PY)


# =================================================================================================
# INDEPENDENT legacy derivations — transcribed DIRECTLY from the .pas, NOT from the rebuild.
# =================================================================================================
def _legacy_856_from_pas(prodDate, lineName):
    """The legacy NORMAL filename, transcribed DIRECTLY from the OPERATIONAL sender MainMenu.pas:2718:

        AssignFile(fcf, ...+'\\856'+copy(EDI856.PickupDate,5,4)+EDI856.LineName+'.txt');

    Breaking the Pascal literal down:
      '856'                          -> the constant prefix
      + copy(EDI856.PickupDate,5,4)  -> chars 5..8 of yyyymmdd = MM+DD (4 chars) = s[4:8] 0-based
      + EDI856.LineName              -> the ASN's line name (e.g. 'COROLLA')
      + '.txt'

    So for '20260618'/'COROLLA': '856' + '0618' + 'COROLLA' + '.txt' = '8560618COROLLA.txt'. This is a hand
    reimplementation of the Pascal, independent of edi856._filename_856. NB the OPERATIONAL sender uses the
    [4:8] (MMDD) offset AND appends LineName — the recreate button (ASNInvoice:820) used a different shape;
    the operational sender is the live daily path the rebuild reproduces."""
    return "856" + prodDate[4:8] + lineName + ".txt"


def _legacy_8hc_from_pas(prodDate, lineName, y=1):
    """The legacy HOT-CALL filename, transcribed DIRECTLY from the OPERATIONAL sender MainMenu.pas:2723:

        AssignFile(fcf, ...+'\\8HC'+copy(EDI856.PickupDate,4,5)+InttoStr(y)+EDI856.LineName+'.txt');

    Breaking the Pascal literal down:
      '8HC'                          -> the constant prefix
      + copy(EDI856.PickupDate,4,5)  -> chars 4..8 of yyyymmdd = last-year-digit+MM+DD (5 chars) = s[3:8]
      + IntToStr(y)                  -> the per-batch hot-call counter (NOT a literal '1'): y init 1 (:2702),
                                        INC only in the hot-call branch (:2724) — single send -> y=1
      + EDI856.LineName              -> the ASN's line name (e.g. 'COROLLA')
      + '.txt'

    So for '20260618'/'COROLLA'/y=1: '8HC' + '60618' + '1' + 'COROLLA' + '.txt' = '8HC606181COROLLA.txt'.
    This is an INDEPENDENT hand-port of the Pascal — it never calls the rebuild's _filename_856.
    Note the hot-call [3:8] offset DIFFERS from the normal [4:8] — a legacy asymmetry preserved verbatim."""
    return "8HC" + prodDate[3:8] + str(y) + lineName + ".txt"


def main():
    print("=" * 88)
    print(" Hot-call ASN producer — PURE build test (8HC filename P13 + create_hotcall_asn P12/P14)")
    print("=" * 88)
    rep = Report()

    # ============================================================================================
    # PART 1 — the 8HC vs 856 filename branch (P13). EXPECTED derived FROM the OPERATIONAL sender
    #          MainMenu.pas:2718/2723 (the C->S-flip send path), NOT the ASNInvoice recreate button.
    # ============================================================================================
    print("\n--- (P13) 8HC vs 856 filename — derived from the operational sender MainMenu.pas:2718/2723 ---")

    # Real production dates from live hot-call ASNs (hotcall-coverage-analysis.md / live query) + a line
    # name; the operational sender appends LineName to BOTH branches.
    LINE = "COROLLA"
    for pd in ("20260618", "20260606", "20260529", "20271225"):
        # NORMAL ASN (startSeq != '-1') -> '856'+MMDD+LineName. Compare REBUILD to the INDEPENDENT .pas port.
        legNormal = _legacy_856_from_pas(pd, LINE)
        rebNormal = edi856._filename_856(pd, LINE, "0001")   # a real seq -> normal branch
        rep.check("normal ASN %s/%s -> rebuild '%s' == legacy-from-.pas '%s'"
                  % (pd, LINE, rebNormal, legNormal),
                  rebNormal == legNormal, "rebuild=%s legacy=%s" % (rebNormal, legNormal))

        # HOT-CALL ASN (startSeq == '-1', y=1) -> '8HC'+Y+MMDD+'1'+LineName. Compare to the .pas port.
        legHot = _legacy_8hc_from_pas(pd, LINE, 1)
        rebHot = edi856._filename_856(pd, LINE, "-1", 1)
        rep.check("hot-call ASN %s/%s (y=1) -> rebuild '%s' == legacy-from-.pas '%s'"
                  % (pd, LINE, rebHot, legHot),
                  rebHot == legHot, "rebuild=%s legacy=%s" % (rebHot, legHot))

    # the concrete bytes for prodDate 20260618 / COROLLA, spelled out so a reviewer can eyeball them.
    rep.check("hot-call '8HC606181COROLLA.txt' (8HC + copy('20260618',4,5)='60618' + y='1' + 'COROLLA')",
              edi856._filename_856("20260618", "COROLLA", "-1", 1) == "8HC606181COROLLA.txt",
              edi856._filename_856("20260618", "COROLLA", "-1", 1))
    rep.check("the same date NORMAL == '8560618COROLLA.txt' (856 + MMDD '0618' + LineName) — branch diverges",
              edi856._filename_856("20260618", "COROLLA", "0001") == "8560618COROLLA.txt",
              edi856._filename_856("20260618", "COROLLA", "0001"))

    # the y counter advances the hot-call name (2nd hot-call of the day for a line -> y=2, no collision).
    print("\n--- (P13) hot-call y counter advances the filename (MainMenu.pas:2702/2724) ---")
    rep.check("hot-call y=2 -> '8HC606182COROLLA.txt' (the 2nd hot-call of the day; no collision with y=1)",
              edi856._filename_856("20260618", "COROLLA", "-1", 2) == "8HC606182COROLLA.txt",
              edi856._filename_856("20260618", "COROLLA", "-1", 2))
    rep.check("y=1 and y=2 names DIFFER (the counter is in the name -> no same-day collision)",
              edi856._filename_856("20260618", "COROLLA", "-1", 1)
              != edi856._filename_856("20260618", "COROLLA", "-1", 2), "y=1 != y=2")

    # back-compat: no startSeq arg -> normal pattern (callers that only build normal ASNs are unaffected).
    rep.check("back-compat: _filename_856(pd, line) (no startSeq) -> normal '856' pattern",
              edi856._filename_856("20260618", "COROLLA") == "8560618COROLLA.txt",
              edi856._filename_856("20260618", "COROLLA"))
    # whitespace-padded sentinel ('-1 ') from a char column still routes to 8HC after strip in send_856 —
    # but _filename_856 itself compares exact '-1'; the driver strips. Assert the exact contract here.
    rep.check("_filename_856 keys on EXACT '-1' (driver strips the char-column padding before calling)",
              edi856._filename_856("20260618", "COROLLA", "-1", 1) == "8HC606181COROLLA.txt"
              and edi856._filename_856("20260618", "COROLLA", "-1 ") == "8560618COROLLA.txt",
              "exact '-1' -> 8HC; ' -1 ' (unstripped) -> 856")

    # --- NON-VACUITY: a WRONG rebuild (the pre-fix recreate-anchored _filename_856 — omits LineName AND
    #     uses the wrong [3:8] offset on the normal branch / literal '1' on hot-call) must FAIL the
    #     operational-sender assertion. Simulate the reverted rebuild inline and prove the assertion would
    #     have caught it (so the green above is meaningful, not vacuous). -----------------------------
    print("\n--- (P13) non-vacuity: the reverted (recreate-anchored, no-LineName) rebuild MUST fail ---")
    def _reverted_normal(prodDate, startSeq=None):
        # the PRE-FIX (PR #29) normal: '856' + [3:8] + '.txt' — wrong offset, NO LineName.
        return "856" + prodDate[3:8] + ".txt"
    def _reverted_hot(prodDate):
        # the PRE-FIX hot-call: '8HC' + [3:8] + literal '1' + '.txt' — NO LineName, no y counter.
        return "8HC" + prodDate[3:8] + "1.txt"
    revNormal = _reverted_normal("20260618", "0001")
    legNormal = _legacy_856_from_pas("20260618", "COROLLA")
    rep.check("REVERTED normal ('85660618.txt') diverges from the operational .pas byte (CAN fail)",
              revNormal != legNormal, "reverted='%s' legacy='%s' -> the NORMAL assertion would FAIL"
              % (revNormal, legNormal))
    revHot = _reverted_hot("20260618")
    legHot = _legacy_8hc_from_pas("20260618", "COROLLA", 1)
    rep.check("REVERTED hot-call ('8HC606181.txt', no LineName) diverges from the operational .pas (CAN fail)",
              revHot != legHot, "reverted='%s' legacy='%s' -> the HOT-CALL assertion would FAIL"
              % (revHot, legHot))

    # ============================================================================================
    # PART 2 — create_hotcall_asn pure bits: computeHotcallDetails + the validators.
    #          EXPECTED derived FROM HotCallEntry.pas (manifest>=8/:157, qty>0/:191, part-req/:212,
    #          per-row detail/:258-285, header IN_QTY = sum fixing @QTY garbage/:246-247).
    # ============================================================================================
    print("\n--- (P12/P14) computeHotcallDetails — manual parts -> detail rows ---")

    # manual entry: 3 parts, 2 sharing the manifest (the @HotCall=1 always-INSERT keeps them DISTINCT).
    MAN = "52089698"            # a real live hot-call manifest (ASN 4712), 8 chars, non-'7'
    items = [("42600FEL2000", 1), ("42607fek5000", 2), ("42600FEL2000", 3)]   # note lower-case middle part
    details = hotcall.computeHotcallDetails(items, MAN)

    rep.check("3 manual rows -> 3 detail rows (no dedup; @HotCall=1 always-INSERT, :279-280)",
              len(details) == 3, "%d rows" % len(details))
    rep.check("every detail carries the operator-typed manifest (:273-274)",
              all(d["manifest"] == MAN for d in details), MAN)
    rep.check("every detail @HotCall=1 (:279-280, the Q1 always-INSERT branch)",
              all(d["hotCall"] == 1 for d in details), "all hotCall=1")
    rep.check("part numbers upper-cased (legacy AssyPartsCodeN CharCase=ecUpperCase)",
              details[1]["partNumber"] == "42607FEK5000", details[1]["partNumber"])
    rep.check("two rows can share a manifest AND be DISTINCT rows (not accumulated)",
              details[0]["partNumber"] == "42600FEL2000" and details[2]["partNumber"] == "42600FEL2000"
              and details[0]["qty"] == 1 and details[2]["qty"] == 3,
              "row0 qty=1, row2 qty=3 — both kept")

    # header IN_QTY = SUM of detail qtys (P14, fixes the @QTY stale-loop-var garbage at :246-247).
    rep.check("header IN_QTY = SUM(detail qtys) = 1+2+3 = 6 (P14, fixes the legacy @QTY garbage)",
              hotcall._detailQtySum(details) == 6, "%d" % hotcall._detailQtySum(details))

    # blank rows are skipped exactly like the legacy loop ignores an empty qty Edit (:262).
    print("\n--- (P12) blank-row skip (legacy ignores an empty qty Edit, :262) ---")
    mixed = [("", ""), ("42600FEL2000", 5), (None, None), ("42600FEK6000", 7)]
    md = hotcall.computeHotcallDetails(mixed, MAN)
    rep.check("blank rows skipped -> only the 2 filled rows emit",
              len(md) == 2 and hotcall._detailQtySum(md) == 12, "%d rows, qty=%d"
              % (len(md), hotcall._detailQtySum(md)))

    # --- validators: manifest length / broadcast / qty / part — all from HotCallEntry.pas -----------
    print("\n--- (P12) validation — from HotCallEntry.pas guards ---")
    def _raises(fn, *a):
        try:
            fn(*a)
            return False
        except hotcall.HotcallError:
            return True

    rep.check("manifest < 8 chars -> HotcallError (HotCallEntry.pas:157, 'must 8 characters')",
              _raises(hotcall._validate_manifest, "1234567"), "len 7 rejected")
    rep.check("manifest exactly 8 chars -> OK (the .dfm MaxLength=8 case; >=8 passes)",
              hotcall._validate_manifest("52089698") == "52089698", "8-char accepted")
    rep.check("manifest starting '7' -> HotcallError (broadcast/M391; a hot-call is M390)",
              _raises(hotcall._validate_manifest, "76061857"), "'7'-prefix rejected")
    rep.check("qty <= 0 -> HotcallError (HotCallEntry.pas:191-197)",
              _raises(hotcall.computeHotcallDetails, [("42600FEL2000", 0)], MAN), "qty 0 rejected")
    rep.check("non-numeric qty -> HotcallError (HotCallEntry.pas:184-190)",
              _raises(hotcall.computeHotcallDetails, [("42600FEL2000", "abc")], MAN), "non-numeric rejected")
    rep.check("qty present but part blank -> HotcallError 'Part number required' (HotCallEntry.pas:204-218)",
              _raises(hotcall.computeHotcallDetails, [("", 5)], MAN), "missing part rejected")
    rep.check("zero items at all -> HotcallError (a hot-call must ship at least one part)",
              _raises(hotcall.computeHotcallDetails, [], MAN), "empty entry rejected")

    # --- NON-VACUITY for the qty-sum (P14): a reverted header-qty (the legacy stale @QTY = last row's qty)
    #     differs from the correct SUM, so the P14 fix is observable / the assertion can fail. -----------
    print("\n--- (P14) non-vacuity: the legacy @QTY garbage (last-row qty) != the correct SUM ---")
    legacyGarbageQty = int(items[-1][1])    # the stale loop var = the LAST validated qty (3), :246-247/:264
    correctSum = hotcall._detailQtySum(details)    # 6
    rep.check("legacy @QTY garbage (last-row qty=3) != correct SUM (6) -> the P14 fix is observable",
              legacyGarbageQty != correctSum, "garbage=%d sum=%d" % (legacyGarbageQty, correctSum))

    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
