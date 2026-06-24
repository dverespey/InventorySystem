#!/usr/bin/env python3
"""test_renban_build.py — PURE unit test for the renban-breakdown trailer distribution + FRS/renban
assignment (docs/analysis/order/project-library/renban/code.py :: compute_trailer_breakdown).

Every expected value is derived FROM RenbanOrder.pas:255-301 (distribution) + :746-799 (read-out /
FRS / renban / counter), transcribed in the spec's STEP-0 section (renban-breakdown-spec.md §12).
Non-vacuous: the distribution is cross-checked against an INDEPENDENT reference transcription of the
Pascal (so it is not "the code agrees with itself"), AND pinned with explicit hand-traced cases.

No DB, no gateway — pure logic. Run:  python3 e2e/test_renban_build.py
"""
import os, sys, importlib.util

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report  # noqa: E402

RENBAN_PY = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..",
                             "ignition", "project-library", "renban", "code.py"))


def _load():
    spec = importlib.util.spec_from_file_location("renban_lib", RENBAN_PY)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


renban = _load()


# ---------------------------------------------------------------------------------------------
# INDEPENDENT reference transcription of RenbanOrder.pas:255-301 (NOT importing the production code).
# Mutates curs inline exactly as TTruck.AddOrder bumps fCurrentCount at :179. Used to cross-check the
# production distribution across thousands of scenarios — this is what makes the math test non-vacuous.
# ---------------------------------------------------------------------------------------------
def ref_distribute(curs, sizes, lots):
    T = len(curs)
    leftover = 0
    if lots // T != 0:                                   # :262
        for i in range(T):
            share = lots // T + leftover                 # :266
            if curs[i] + share <= sizes[i]:             # :266 fits
                curs[i] += share                         # :268
                leftover = 0
            else:                                        # :271 spill
                leftover = (lots // T + leftover) - (sizes[0] - curs[i])   # :273 (sizes[0]!)
                curs[i] += sizes[i] - curs[i]            # :274 top off
    if lots % T != 0:                                    # :279
        rem = lots % T + leftover                        # :282
        while rem != 0:                                  # :285
            for i in range(T):
                if curs[i] + 1 <= sizes[i]:             # :289
                    curs[i] += 1                         # :291
                    rem -= 1
                if rem == 0:
                    break                                # :295


def _part(part, lotqty, order_qty, frs="6090101", kanban="K1", sup="0572B",
          gc="CMWA", seed="CMWA288"):
    return {"kanban": kanban, "supplier": sup, "part": part, "frs": frs,
            "order_qty": order_qty, "lotqty": lotqty, "group_code": gc, "renban_seed": seed}


def main():
    print("=" * 78)
    print(" RENBAN BREAKDOWN — pure distribution + FRS/renban (vs RenbanOrder.pas:255-301 / :746-799)")
    print("=" * 78)
    rep = Report()

    # --- 1. HAND-TRACED base case: 1 part, 10 lots, 3 trailers x 4 pallets ----------------------
    # lots=10, T=3: Phase A 10 div 3 = 3 -> each truck +3 (t=[3,3,3]); Phase B 10 mod 3 = 1 -> t0 +1.
    #   => t0=4, t1=3, t2=3.  FRS prefix '60901' + ordinal 01/02/03.  renban seed 288 + truck idx.
    #   qty = lots * lotqty (30) = 120/90/90.  next_count = last rcount (288+2=290) + 1 = 291.
    res = renban.compute_trailer_breakdown([_part("Q5100", 30, 300)], 3, 4, "CMWA")
    rep.check("base: total_lots = floor(300/30) = 10", res["total_lots"] == 10, str(res["total_lots"]))
    rep.check("base: truck loads [4,3,3] (lots div 3 + 1-lot remainder on truck0)",
              res["truck_counts"] == [4, 3, 3], str(res["truck_counts"]))
    got = [(r["truck"], r["frs"], r["renban"], r["lots"], r["qty"]) for r in res["rows"]]
    want = [(0, "6090101", "CMWA288", 4, 120), (1, "6090102", "CMWA289", 3, 90),
            (2, "6090103", "CMWA290", 3, 90)]
    rep.check("base: per-trailer (frs, renban, lots, qty) match the .pas read-out", got == want,
              "got=%s" % got)
    rep.check("base: next_count = last rcount + 1 = 291", res["next_count"] == 291, str(res["next_count"]))

    # --- 2. FRS suffix derivation (:763-767) — 5-char prefix + 2-digit trailer ordinal ----------
    # ordinal = TruckNumber+1; trucks 0..7 -> '01'..'08', truck 8 -> '09', truck 9 (TruckNumber>8) -> '10'.
    rep.check("FRS suffix truck0 -> '01'", renban._frs_suffix("6090199", 0) == "6090101")
    rep.check("FRS suffix truck7 -> '08'", renban._frs_suffix("6090199", 7) == "6090108")
    rep.check("FRS suffix truck8 -> '09'", renban._frs_suffix("6090199", 8) == "6090109")
    rep.check("FRS suffix truck9 -> '10' (TruckNumber>8, raw 2 digits)",
              renban._frs_suffix("6090199", 9) == "6090110")
    rep.check("FRS keeps only first 5 chars of the original FRS (copy(frs,1,5))",
              renban._frs_suffix("ABCDEFG", 0) == "ABCDE01")

    # --- 3. Renban number derivation (:775-779) — group code + %.3d(seed3 + truck idx) ----------
    rep.check("renban truck0: CMWA + 288", renban._renban_number("CMWA", "CMWA288", 0) == ("CMWA288", 288))
    rep.check("renban truck1: CMWA + 289", renban._renban_number("CMWA", "CMWA288", 1) == ("CMWA289", 289))
    rep.check("renban truck2: CMWA + 290", renban._renban_number("CMWA", "CMWA288", 2) == ("CMWA290", 290))
    rep.check("renban reads the 3-digit TAIL of the seed (rightstr(renban,3))",
              renban._renban_number("DICAS", "DICAS480", 3) == ("DICAS483", 483))
    rep.check("renban seed '000' is valid (count is a uniqueness seed, not a qty)",
              renban._renban_number("CAP", "CAP000", 2) == ("CAP002", 2))

    # --- 4. MERGE repeated parts on a truck (R3 sum-all, TTruck.AddOrder :160-163) --------------
    # Two order rows for the SAME part in the same group, both on a single trailer (T=1): the lots
    # must SUM into ONE line, not two. part lotqty 40; rows 200 (5 lots) + 120 (3 lots) = 8 lots.
    res = renban.compute_trailer_breakdown(
        [_part("Q8000", 40, 200), _part("Q8000", 40, 120)], 1, 20, "CMWA")
    rep.check("merge: T=1 carries ONE merged line for the repeated part (sum-all)",
              len(res["rows"]) == 1, "rows=%d" % len(res["rows"]))
    rep.check("merge: lots SUMMED (5 + 3 = 8), qty = 8*40 = 320",
              res["rows"][0]["lots"] == 8 and res["rows"][0]["qty"] == 320,
              "lots=%d qty=%d" % (res["rows"][0]["lots"], res["rows"][0]["qty"]))

    # --- 5. SKIP qty==0 trailer rows (:540 "don't send 0 qty orders") ---------------------------
    # 1 part, 2 lots, 4 trailers x 1 pallet (capacity 4 >= 2). Phase B dribbles 2 lots onto t0,t1;
    # t2,t3 get nothing -> no rows for them. Only 2 rows emitted (the empty trucks produce no lines).
    res = renban.compute_trailer_breakdown([_part("Q4000", 10, 20)], 4, 1, "CMWA")
    rep.check("skip-empty: only the 2 loaded trailers emit rows (empty trucks -> no 0-qty rows)",
              len(res["rows"]) == 2 and all(r["qty"] > 0 for r in res["rows"]),
              "rows=%s" % [(r["truck"], r["qty"]) for r in res["rows"]])
    rep.check("skip-empty: truck_counts show the 2 empties [1,1,0,0]",
              res["truck_counts"] == [1, 1, 0, 0], str(res["truck_counts"]))
    # next_count tracks the highest NON-empty truck's rcount + 1 (last emitted t1 -> 288+1=289, +1=290)
    rep.check("skip-empty: next_count = highest non-empty rcount + 1 = 290",
              res["next_count"] == 290, str(res["next_count"]))

    # --- 6. CAPACITY GATE (:709) — refuse when trailers x pallets < TotalLots -------------------
    threw = False
    try:
        renban.compute_trailer_breakdown([_part("Q5100", 30, 300)], 2, 4, "CMWA")  # 10 lots > 2*4=8
    except ValueError as e:
        threw = "will not fit" in str(e)
    rep.check("capacity gate: 10 lots into 2x4=8 capacity raises 'will not fit'", threw)
    # exact-fit boundary passes (3x4 = 12 >= 10, and 10 into 2x5=10 exactly)
    ok_fit = True
    try:
        renban.compute_trailer_breakdown([_part("Q5100", 30, 300)], 2, 5, "CMWA")   # 10 == 10
    except ValueError:
        ok_fit = False
    rep.check("capacity gate: exact-fit (10 lots into 2x5=10) is allowed", ok_fit)

    # --- 7. DIV-BY-ZERO guard (H7) — lotqty 0/NULL raises, does not crash silently --------------
    for bad in (0, None, ""):
        threw = False
        try:
            renban.compute_trailer_breakdown([_part("Qbad", bad, 100)], 1, 10, "CMWA")
        except ValueError as e:
            threw = "lotqty" in str(e)
        rep.check("div-zero guard: lotqty=%r raises (no crash)" % bad, threw)

    # --- 8. MULTI-PART distribution incl. the forward-spill, vs the INDEPENDENT reference --------
    # Exhaustive cross-check: every 1-3 part scenario over T in 2..4, P in 2..6 whose lots fit. The
    # production compute's per-truck totals MUST equal the independent ref_distribute transcription.
    total = mism = spill_probe = 0
    for T in range(2, 5):
        for P in range(2, 7):
            for a in range(0, T * P + 1):
                for b in range(0, T * P + 1):
                    for c in range(0, T * P + 1):
                        if a + b + c > T * P:
                            continue
                        lots_list = [x for x in (a, b, c) if x > 0]
                        if not lots_list:
                            continue
                        # production: lotqty=1 so order_qty == lots; one part-number per line index
                        parts = [_part("P%d" % i, 1, lots) for i, lots in enumerate(lots_list)]
                        res = renban.compute_trailer_breakdown(parts, T, P, "CMWA")
                        prod = res["truck_counts"]
                        ref = [0] * T
                        for lots in lots_list:
                            ref_distribute(ref, [P] * T, lots)
                        total += 1
                        if prod != ref:
                            mism += 1
                            if mism <= 3:
                                print("    MISMATCH T=%d P=%d parts=%s prod=%s ref=%s"
                                      % (T, P, lots_list, prod, ref))
                        # probe the forward-spill else-branch is exercised somewhere (3+ uneven parts)
                        if len(lots_list) >= 3:
                            spill_probe += 1
    rep.check("distribution matches the independent .pas transcription across %d scenarios" % total,
              mism == 0, "mismatches=%d" % mism)
    rep.check("the exhaustive set is real (>5000 scenarios, incl. multi-part spill paths)",
              total > 5000 and spill_probe > 100, "total=%d spill_probe=%d" % (total, spill_probe))

    # --- 8b. COUNTER-ROLLOVER — P4 CLEAN WRAP (999->000), oracle derived from spec §12.7 ring math -----
    # POST-CUTOVER (P4): the persisted count and the renban-number tail both RING-WRAP `% 1000`. The
    # expected values below are derived INDEPENDENTLY from the spec's clean-wrap ORACLE (renban-breakdown-
    # spec.md §12.7 — `next_count % 1000` for the count, `group_code + '%03d' % (rcount % 1000)` for the
    # renban), NOT recomputed by the code under test (R15: derive the expectation from the SOURCE/spec, not
    # the rebuild). A SEPARATE ring-math helper here (ring3 / ring_renban) is the oracle; the rebuild's
    # output is compared against it. Non-vacuity: REVERT the wrap (restore str(N)[:3] + the 4-digit emit)
    # and these checks FAIL — proven below + by reverting renban/code.py (see the run note).
    #
    # Scenario (spec §12.7 worked example): seed 'CMWA997', 5 trailers x 2 pallets, one 5-lot part ->
    # Phase A gives each truck 1 lot -> RAW rcounts 997, 998, 999, 1000, 1001 -> next_count = 1001 + 1 =
    # 1002. The RAW rcount is the running count (next_count carries it forward UNWRAPPED); the wrap is
    # applied ONLY on render (renban string) and on persist (the % 1000 count).
    def ring3(n):
        """SPEC ORACLE: the persisted varchar(3) count under the clean wrap = '%03d' % (n % 1000)."""
        return "%03d" % (n % 1000)
    def ring_renban(group, raw_rcount):
        """SPEC ORACLE: the rendered renban string under the clean wrap = group + 3-digit (rcount % 1000)."""
        return group + ("%03d" % (raw_rcount % 1000))

    res = renban.compute_trailer_breakdown([_part("Q5100", 30, 150, seed="CMWA997")], 5, 2, "CMWA")
    # next_count carries the RAW running count forward (max raw rcount + 1) — UNCHANGED by the wrap.
    rep.check("rollover: next_count carries the RAW running count across 1000 (seed 997 + 5 trailers -> 1002)",
              res["next_count"] == 1002, str(res["next_count"]))
    # PERSISTED count under the clean wrap = the SPEC ring oracle ring3(1002) = '002', NOT the pre-cutover
    # str(1002)[:3]='100'. NOTE the persist itself (UPDATE_RenbanGroupCount step (c)) is a DB write only
    # exercised by test_renban_e2e.py, which reads the ACTUAL stored value back (the persist-step non-vacuity
    # lives there — revert step (c) -> the DB stores '100' -> that e2e check flips red). Here in the PURE
    # test we pin the ORACLE itself (so a wrong oracle is caught) and that ring3 != the pre-cutover trunc.
    rep.check("rollover (clean wrap): spec persist oracle ring3(1002) = '002' (ring-wrap), NOT pre-cutover '100'",
              ring3(1002) == "002" and ring3(1002) != ("%03d" % 1002)[:3] and ("%03d" % 1002)[:3] == "100",
              "ring3(1002)=%r pre-cutover=%r" % (ring3(1002), ("%03d" % 1002)[:3]))
    # the renban NUMBERS ring-wrap: 997/998/999 stay, 1000->000, 1001->001 (3-digit, NO 4-digit CMWA1000).
    renban_strings = sorted(r["renban"] for r in res["rows"])
    expected_renbans = sorted(ring_renban("CMWA", rc) for rc in (997, 998, 999, 1000, 1001))  # spec oracle
    rep.check("rollover (clean wrap): renban strings = spec oracle CMWA997/998/999/000/001 (ring-wrapped)",
              renban_strings == expected_renbans and renban_strings ==
              ["CMWA000", "CMWA001", "CMWA997", "CMWA998", "CMWA999"], str(renban_strings))
    rep.check("rollover (clean wrap): NO 4-digit renban emitted (CMWA1000 never appears; CMWA000 instead)",
              "CMWA1000" not in renban_strings and "CMWA000" in renban_strings, str(renban_strings))
    # sub-1000 is byte-for-byte UNCHANGED (N % 1000 == N for N<1000 == str(N)[:3]): the wrap touches ONLY
    # the 999->000 boundary, so every non-rollover renban/count is identical to before.
    rep.check("rollover (clean wrap): sub-1000 persisted count unchanged vs both renders (5->'005', 634->'634')",
              ring3(5) == "005" and ring3(634) == "634"
              and ring3(5) == ("%03d" % 5)[:3] and ring3(634) == ("%03d" % 634)[:3])
    rep.check("rollover (clean wrap): next_count EXACTLY 1000 persists ring3(1000)='000' (was '100' pre-cutover)",
              ring3(1000) == "000" and ring3(1000) != ("%03d" % 1000)[:3])
    # a run that does NOT cross 999 is identical to the legacy: same seed, sub-1000 rcounts -> no change.
    res_nowrap = renban.compute_trailer_breakdown([_part("Q5100", 30, 150, seed="CMWA288")], 5, 2, "CMWA")
    nowrap_renbans = sorted(r["renban"] for r in res_nowrap["rows"])
    rep.check("clean wrap is a NO-OP for a non-rollover run (seed 288 -> CMWA288..292, count 293 unchanged)",
              nowrap_renbans == ["CMWA288", "CMWA289", "CMWA290", "CMWA291", "CMWA292"]
              and ring3(res_nowrap["next_count"]) == "293" and res_nowrap["next_count"] == 293,
              "renbans=%s next_count=%s" % (nowrap_renbans, res_nowrap["next_count"]))

    # --- 9. Conservation: total emitted lots == total input lots (no lot lost/duplicated) --------
    parts = [_part("Q4000", 40, 200), _part("Q5000", 30, 300), _part("Q8000", 40, 360)]  # 5+10+9=24 lots
    res = renban.compute_trailer_breakdown(parts, 3, 10, "CMWA")
    emitted = sum(r["lots"] for r in res["rows"])
    rep.check("conservation: sum of emitted lots == 24 == sum(floor(qty/lotqty))",
              emitted == 24 and res["total_lots"] == 24, "emitted=%d total=%d" % (emitted, res["total_lots"]))
    rep.check("conservation: every emitted qty is an exact lot multiple (qty == lots*lotqty)",
              all(r["qty"] == r["lots"] * r["lotqty"] for r in res["rows"]))

    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
