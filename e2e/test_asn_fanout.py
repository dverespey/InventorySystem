#!/usr/bin/env python3
"""test_asn_fanout.py — unit test for the PURE ASN-creation ratio fan-out (asn.computeAsnDetails).

Exercises the pure decision logic that reproduces the hand-written Delphi CalculateASNFRS
(DataModule.pas:5180-5268) with NO database: asn/code.py imports nothing and is CPython-importable
straight from the on-disk project library. Twin of test_shipping_posts.py (producer-recipe pattern).

Source truth:
  docs/analysis/production-readiness/m1-asn-creation-spec.md  §4 (the two branches), §6 (manifest)
  docs/analysis/production-readiness/AD_FRSPULL-shared.sql    (Orders/Vehicles definitions)
  DataModule.pas:5180-5268                                    (exact branch/round/manifest logic)

Covers (with synthetic FRS + forecast rows of the verified shapes):
  * GROUND BC ratio split across tire+wheel forecast rows (Orders = Vehicles*4, ratio path)
  * SPARE  BC ratio path (Orders = Vehicles)
  * Orders <= 5 -> single No-Ratio row (qty = Vehicles*IN_ASSY_QTY, no split, first fc row only)
  * both ratios = 100 -> full qty, no rounding
  * round-half-to-EVEN (banker's) boundary cases (2.5->2, 3.5->4, 4.5->4) — the parity trap
  * manifest format: '7' + 1-digit-year + MM + DD + 2-char id (asserts the 7###### shape)
  * abort: BC with no forecast detail; a part with NULL IN_MANIFEST_COST_ID

Run:  python3 e2e/test_asn_fanout.py     (no DB, no env needed)
"""
import os, sys, importlib.util

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report  # noqa: E402

PLIB = os.path.normpath(os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "..",
    "ignition", "project-library"))


def _load(name, rel):
    spec = importlib.util.spec_from_file_location(name, os.path.join(PLIB, rel, "code.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


asn = _load("asn_lib", "asn")


def fc(part, manifestId, assyQty, tire, wheel, costId=1):
    """A SELECT_ForecastDetailBCASN row (the columns computeAsnDetails reads)."""
    return {
        "VC_ASSY_PART_NUMBER_CODE": part,
        "VC_ASSY_MANIFEST_NUMBER": manifestId,
        "IN_ASSY_QTY": assyQty,
        "IN_TIRE_RATIO": tire,
        "IN_WHEEL_RATIO": wheel,
        "IN_MANIFEST_COST_ID": costId,
    }


def frs(bc, orders, vehicles):
    """An AD_FRSPull row."""
    return {"BC": bc, "Orders": orders, "VEHICLES": vehicles}


PDATE = "20260618"   # -> manifest prefix '7' + '60618'
ASNID = 4721


def main():
    print("=" * 80)
    print(" ASN ratio fan-out — pure computeAsnDetails unit test (DataModule.pas:5180-5268)")
    print("=" * 80)
    rep = Report()
    cad = asn.computeAsnDetails

    # --- banker's-rounding helper directly (the parity trap, spec §9.3) -------------------------
    print("\n--- banker's rounding (round-half-to-EVEN), exact integer math ---")
    br = asn._bankers_div_round
    rep.check("2.5 -> 2 (even; round-half-up would give 3)", br(250, 100) == 2, str(br(250, 100)))
    rep.check("3.5 -> 4 (even)", br(350, 100) == 4, str(br(350, 100)))
    rep.check("4.5 -> 4 (even; round-half-up would give 5)", br(450, 100) == 4, str(br(450, 100)))
    rep.check("5.5 -> 6 (even)", br(550, 100) == 6, str(br(550, 100)))
    rep.check("2.4 -> 2 (below half, down)", br(240, 100) == 2, str(br(240, 100)))
    rep.check("2.6 -> 3 (above half, up)", br(260, 100) == 3, str(br(260, 100)))
    rep.check("exact 300/100 -> 3 (no rounding)", br(300, 100) == 3, str(br(300, 100)))

    # --- manifest format (spec §6: '7' + 1-digit year + MM + DD + 2-char id) --------------------
    print("\n--- manifest generation ---")
    m = asn._manifest(PDATE, "57")
    rep.check("manifest 20260618 + '57' == '76061857' (matches daily log)", m == "76061857", m)
    rep.check("manifest starts with '7' and is 8 chars", m[0] == "7" and len(m) == 8, m)
    m2 = asn._manifest("20270618", "05")   # 2027 -> 1-digit year '7'
    rep.check("2027 production date -> '77061805' (1-digit year)", m2 == "77061805", m2)

    # --- GROUND BC: Orders = Vehicles*4 (>5) ratio split across two forecast rows ----------------
    print("\n--- GROUND BC ratio split (tire + wheel forecast rows) ---")
    # 10 vehicles, ground -> Orders = 40. Two fc rows: a 60/40 tire/wheel-ish split (ratio applied
    # to tire only per the Delphi). assyQty 1 each. tire ratios 60 and 40.
    ground = {"GAB": [
        fc("TIRE-PART", "10", assyQty=1, tire=60, wheel=60),   # qty = round(10*1*60/100)=6
        fc("WHEEL-PART", "11", assyQty=1, tire=40, wheel=40),  # qty = round(10*1*40/100)=4
    ]}
    out = cad([frs("GAB", orders=40, vehicles=10)], ground, PDATE, ASNID)
    byPart = dict((d["partNumber"], d) for d in out)
    rep.check("ground BC emits one row per forecast row (2)", len(out) == 2,
              str([(d["partNumber"], d["qty"]) for d in out]))
    rep.check("tire row qty = round(10*1*60/100) = 6", byPart["TIRE-PART"]["qty"] == 6,
              str(byPart["TIRE-PART"]["qty"]))
    rep.check("wheel row qty = round(10*1*40/100) = 4", byPart["WHEEL-PART"]["qty"] == 4,
              str(byPart["WHEEL-PART"]["qty"]))
    rep.check("ratio rows flagged noRatio=False", all(not d["noRatio"] for d in out))
    rep.check("ratio row manifest = '7'+'60618'+id", byPart["TIRE-PART"]["manifest"] == "76061810",
              byPart["TIRE-PART"]["manifest"])
    rep.check("every detail carries the asnId", all(d["asnId"] == ASNID for d in out))

    # --- SPARE BC: Orders = Vehicles (>5), ratio path -------------------------------------------
    print("\n--- SPARE BC ratio path (Orders = Vehicles) ---")
    spare = {"SPB": [fc("SPARE-PART", "20", assyQty=1, tire=100, wheel=100)]}
    out = cad([frs("SPB", orders=8, vehicles=8)], spare, PDATE, ASNID)
    rep.check("spare BC single fc row -> 1 detail", len(out) == 1, str(out))
    rep.check("spare qty (both ratios 100) = 8*1 = 8 (no rounding)", out[0]["qty"] == 8, str(out[0]["qty"]))

    # --- both ratios = 100 -> full qty, no rounding ---------------------------------------------
    print("\n--- ratio = 100 -> full qty ---")
    full = {"FUL": [fc("FULL-PART", "30", assyQty=3, tire=100, wheel=100)]}
    out = cad([frs("FUL", orders=12, vehicles=7)], full, PDATE, ASNID)
    rep.check("both-100 -> qty = Vehicles*IN_ASSY_QTY = 7*3 = 21 (full, no round)",
              out[0]["qty"] == 21, str(out[0]["qty"]))

    # --- Orders <= 5 -> single No-Ratio row (no split, first fc row only, break) -----------------
    print("\n--- No-Ratio branch (Orders <= 5) ---")
    small = {"SML": [
        fc("FIRST-PART", "40", assyQty=2, tire=60, wheel=60),   # only THIS one emitted
        fc("SECOND-PART", "41", assyQty=9, tire=40, wheel=40),  # must be SKIPPED (break)
    ]}
    out = cad([frs("SML", orders=5, vehicles=3)], small, PDATE, ASNID)
    rep.check("Orders<=5 emits exactly ONE row (break after first fc row)", len(out) == 1, str(out))
    rep.check("No-Ratio qty = Vehicles*IN_ASSY_QTY = 3*2 = 6 (NO ratio applied)",
              out[0]["qty"] == 6, str(out[0]["qty"]))
    rep.check("No-Ratio row is the FIRST fc part (FIRST-PART), not the second",
              out[0]["partNumber"] == "FIRST-PART", out[0]["partNumber"])
    rep.check("No-Ratio row flagged noRatio=True", out[0]["noRatio"] is True)
    rep.check("No-Ratio manifest = '7'+'60618'+'40' = '76061840'",
              out[0]["manifest"] == "76061840", out[0]["manifest"])
    # boundary: Orders == 5 is No-Ratio; Orders == 6 is ratio.
    out6 = cad([frs("SML", orders=6, vehicles=3)], small, PDATE, ASNID)
    rep.check("Orders == 6 -> ratio branch (both fc rows, no break)", len(out6) == 2,
              str([(d["partNumber"], d["qty"]) for d in out6]))

    # --- round-half-to-even through the full fan-out (parity trap end-to-end) --------------------
    print("\n--- banker's rounding through the fan-out (2.5->2, 4.5->4) ---")
    # Vehicles=5, assyQty=1, tire=50 -> base*tire = 250 -> 2.5 -> banker's 2 (round-half-up=3).
    even1 = {"E1": [fc("EVEN1", "50", assyQty=1, tire=50, wheel=50)]}
    out = cad([frs("E1", orders=20, vehicles=5)], even1, PDATE, ASNID)
    rep.check("fan-out 5*1*50/100 = 2.5 -> 2 (banker's; round-half-up would give 3)",
              out[0]["qty"] == 2, str(out[0]["qty"]))
    # Vehicles=9, assyQty=1, tire=50 -> 450 -> 4.5 -> banker's 4 (round-half-up=5).
    even2 = {"E2": [fc("EVEN2", "51", assyQty=1, tire=50, wheel=50)]}
    out = cad([frs("E2", orders=36, vehicles=9)], even2, PDATE, ASNID)
    rep.check("fan-out 9*1*50/100 = 4.5 -> 4 (banker's; round-half-up would give 5)",
              out[0]["qty"] == 4, str(out[0]["qty"]))
    # Vehicles=7, assyQty=1, tire=50 -> 350 -> 3.5 -> banker's 4.
    odd = {"OD": [fc("ODD", "52", assyQty=1, tire=50, wheel=50)]}
    out = cad([frs("OD", orders=28, vehicles=7)], odd, PDATE, ASNID)
    rep.check("fan-out 7*1*50/100 = 3.5 -> 4 (banker's, rounds to even)",
              out[0]["qty"] == 4, str(out[0]["qty"]))

    # --- multi-BC in one create: ratios + a No-Ratio in the same call ---------------------------
    print("\n--- multi-BC create (ratio + No-Ratio mixed) ---")
    frows = [frs("GAB", 40, 10), frs("SML", 5, 3)]
    fmap = {}
    fmap.update(ground)
    fmap.update(small)
    out = cad(frows, fmap, PDATE, ASNID)
    rep.check("2 BCs (ground=2 rows + small=1 row) -> 3 details total", len(out) == 3,
              str([(d["partNumber"], d["qty"], d["noRatio"]) for d in out]))
    rep.check("exactly one No-Ratio row in the mix",
              sum(1 for d in out if d["noRatio"]) == 1)

    # --- abort conditions -----------------------------------------------------------------------
    print("\n--- abort conditions (whole create fails) ---")
    try:
        cad([frs("NOFC", 40, 10)], {}, PDATE, ASNID)   # BC has no forecast detail
        rep.check("BC with no forecast detail -> AsnFanoutError", False, "no error raised")
    except asn.AsnFanoutError as e:
        rep.check("BC with no forecast detail -> AsnFanoutError",
                  "Missing Broadcast Code Information" in str(e), str(e))
    try:
        bad = {"BADC": [fc("NOCOST", "60", 1, 60, 60, costId=None)]}
        cad([frs("BADC", 40, 10)], bad, PDATE, ASNID)  # NULL manifest cost
        rep.check("part with NULL IN_MANIFEST_COST_ID -> AsnFanoutError", False, "no error raised")
    except asn.AsnFanoutError as e:
        rep.check("part with NULL IN_MANIFEST_COST_ID -> AsnFanoutError",
                  "Missing Manifest Cost" in str(e), str(e))
    # the cost abort fires BEFORE any detail of that BC is emitted (mixed-row BC).
    try:
        mixed = {"MXC": [fc("OKPART", "61", 1, 60, 60, costId=1),
                         fc("BADPART", "62", 1, 40, 40, costId=None)]}
        cad([frs("MXC", 40, 10)], mixed, PDATE, ASNID)
        rep.check("cost abort pre-empts the whole BC (no partial emit)", False, "no error raised")
    except asn.AsnFanoutError as e:
        rep.check("cost abort pre-empts the whole BC (no partial emit)",
                  "BADPART" in str(e), str(e))

    print("\n--- empty input ---")
    rep.check("no FRS rows -> no details", cad([], {}, PDATE, ASNID) == [])

    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
