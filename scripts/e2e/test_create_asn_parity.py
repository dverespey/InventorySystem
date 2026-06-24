#!/usr/bin/env python3
"""test_create_asn_parity.py — END-TO-END exercise of the REAL `asn.create_asn` gateway driver,
driven headlessly through the system.db shim (jython_shim) — the M1 ASN-creation keystone proof.

WHAT THIS PROVES (and what it HONESTLY does NOT)
------------------------------------------------
This runs the ACTUAL driver code (the AD_FRSPULL cross-DB read on VehicleOrder, the per-BC
SELECT_ForecastDetailBCASN read on Inventory, the INSERT_ASNInfo header + OUTPUT capture, the
INSERT_ASNDetail accumulate upsert, all inside ONE Inventory transaction) — not a proc-EXEC, which
would never touch the driver/seam (retro R8 discipline). It asserts two things that ARE real:

  (A) DRIVER SELF-CONSISTENCY — every detail row the driver PERSISTED equals what the reviewed pure
      fan-out (asn.computeAsnDetails) produces over the SAME inputs (the spike's AD_FRSPULL +
      forecast). This proves the wiring (cross-DB read -> banker's round -> accumulate upsert ->
      header/transaction) is faithful. It is self-consistency, NOT legacy parity.

  (B) TOTAL-QTY INVARIANT — the rebuild's total detail qty equals the legacy ASN's total detail qty
      (4240 == 4240 for ASN 4721). This holds INDEPENDENT of per-manifest distribution and is a real
      cross-check against the frozen legacy data: the fan-out conserves the same grand total even
      when the recipe vintage that shaped the per-manifest split has since changed.

It does NOT claim row-for-row legacy parity, because — as documented in the NOTE below — no currently
reproducible legacy ASN matches the current DB's forecast recipe (see "WHY NOT ROW PARITY"). Per the
fixture-fidelity discipline (memory feedback-parity-fixture-fidelity): we do NOT sell self-consistency
as parity. The per-manifest legacy diff is REPORTED for transparency, not gated.

WHY NOT ROW PARITY (the corrected diagnosis — the old "VehicleOrder fixture drift" note was WRONG)
-------------------------------------------------------------------------------------------------
The earlier note blamed a "full historical VehicleOrder reload (2.3M vehicles)". That is FALSE and the
data disproves it three ways:

  1. THE WINDOW IS THE FILTER. AD_FRSPULL bounds purely by [begindate, enddate] + line. For ASN 4721's
     window the spare (1-per-vehicle) BC rows sum to exactly 848 = the legacy ASN 4721 header qty. No
     "2.3M reload" leaks in — the date window already isolates the build snapshot's 848 vehicles.
     (Verified: sum of the 1-per-vehicle BC rows in the 4721 window == 848.)

  2. TOTAL DETAIL QTY IS IDENTICAL. Legacy 4721 detail sums to 4240; the rebuild over the spike's own
     AD_FRSPULL + today's forecast also sums to 4240. A vehicle-count drift would move the TOTAL; it
     does not. Only the per-manifest DISTRIBUTION differs.

  3. THE PER-MANIFEST DIFFS ARE MATHEMATICALLY IMPOSSIBLE FROM VEHICLE DRIFT. The forecast detail the
     driver consumes (SELECT_ForecastDetailBCASN: parts / ratios / assy / manifest ids) is
     BYTE-IDENTICAL between Inventory and Inventory_Live for the current 2026/06 vintage (verified per
     BC). Yet within a single BC the legacy qtys imply inconsistent vehicle counts under today's
     recipe. NBB today is assy=4, tire 40/20/40 -> manifests 76061857/58/59. The legacy trio
     80/900/1124 would require VEHICLES = 50, 1125 and 702.5 respectively — fractional and mutually
     inconsistent. No single integer vehicle count can produce all three. (Same shape for NJJ, NCC.)

==> ASN 4721 (and 4718-4720) were FROZEN under a DIFFERENT forecast-recipe vintage in Inventory_Live's
    OWN history — the ratios / assy / manifest-mapping in effect when the ASN was built differ from
    today's DB. This is recipe-vintage drift in the legacy reference's history, NOT a VehicleOrder
    reload and NOT a driver fault.

ATTEMPT AT A TRUE POINT-IN-TIME ORACLE (and why none was reproducible)
---------------------------------------------------------------------
Per the task, we first looked for the most-recent NORMAL multi-detail legacy ASN whose recipe vintage
is most likely UNCHANGED vs the current DB, hoping for a ROW-FOR-ROW match:
  * max id 4722 is a hot-call/manual ASN (qty 1, seq -1/-1, no FRS fan-out) — not usable.
  * 4721/4720/4719/4718 are normal multi-detail ASNs, but ALL drift per-manifest exactly as above
    (each 6/17 manifests match; totals 4240/4000/4080/4280 reproduce, distribution does not).
No currently reproducible legacy ASN matches the current forecast recipe, so there is no honest
row-for-row parity assertion to make. The test is therefore (A)+(B), labeled as such. If a future
legacy ASN is created under the current recipe vintage it WILL match row-for-row and this test should
be upgraded to gate on full parity (see the IG-TODO at the gate).

METHOD
------
  1. Read the legacy header (window + seq) and detail rows for LEGACY_ASN from Inventory_Live.
  2. Run the REAL create_asn in Inventory for (LINE, PDATE, that window).
  3. Read back the created INV_ASN_DETAIL_MST rows.
  4. (A) Recompute computeAsnDetails over the SAME inputs and diff vs the persisted rows.
     (B) Assert rebuild total detail qty == legacy total detail qty.
     Report the per-manifest legacy diff (informational, not gated).
  5. Restore Inventory as-found: delete the test ASN header + its details.

Honest fixture-fidelity: self-consistency + the conserved total are asserted; row-parity is reported,
not forced green.

Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 scripts/e2e/test_create_asn_parity.py
"""
import os, subprocess, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report, DB_CONN          # noqa: E402
import jython_shim              # noqa: E402

CONTAINER = os.environ.get("CONTAINER", "mssql-spike")
SA_PASS = os.environ.get("SA_PASS")
INV = "Inventory"
LIVE = "Inventory_Live"

LINE = "COROLLA"
PDATE = "20260618"
LEGACY_ASN = 4721

ASN_PY = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                          "docs", "analysis", "edi", "project-library", "asn", "code.py"))


def sql(db, query):
    if not SA_PASS:
        sys.exit("export SA_PASS first")
    # SET QUOTED_IDENTIFIER/ANSI_NULLS ON: the gateway JDBC session runs with these ON; once
    # INV_ASN_MST carries the filtered unique index (spike-asn-unique-guard.sql) ANY DML against it
    # (e.g. the cleanup DELETE below) REQUIRES them ON or SQL Server raises Msg 1934. sqlcmd's -Q
    # default has QUOTED_IDENTIFIER OFF, so prepend them to mirror the gateway.
    out = subprocess.check_output([
        "docker", "exec", CONTAINER, "/opt/mssql-tools18/bin/sqlcmd", "-C", "-S", "localhost",
        "-U", "sa", "-P", SA_PASS, "-d", db, "-h", "-1", "-W", "-s", "\t",
        "-Q", "SET QUOTED_IDENTIFIER ON; SET ANSI_NULLS ON; SET NOCOUNT ON; " + query], text=True)
    return [l.split("\t") for l in out.splitlines()
            if l.strip() and not l.startswith("(") and not l.startswith("Msg ")]


def scalar(db, q):
    rows = sql(db, q)
    return rows[0][0] if rows else None


def details_map(db, asnId):
    """{manifest: (part, qty)} for one ASN's detail rows."""
    rows = sql(db, "SELECT VC_MANIFEST_NUMBER, VC_ASSY_PART_NUMBER, IN_QTY "
                   "FROM INV_ASN_DETAIL_MST WHERE IN_ASN_ID=%d" % asnId)
    return dict((r[0], (r[1], int(r[2]))) for r in rows)


def _total_qty(dmap):
    return sum(q for (_, q) in dmap.values())


def main():
    print("=" * 80)
    print(" create_asn END-TO-END — driver self-consistency + total-qty invariant "
          "(legacy ASN %d)" % LEGACY_ASN)
    print(" (NOT row parity — legacy ref frozen under an older forecast-recipe vintage; see NOTE)")
    print("=" * 80)
    rep = Report()

    # --- load the REAL driver under the shim --------------------------------------------------
    asn = jython_shim.load_wrapper("asn_driver", ASN_PY)

    # --- 1. legacy ground truth ---------------------------------------------------------------
    hdr = sql(LIVE, "SELECT VC_START_SEQ_NUMBER, CONVERT(varchar,DT_START_SEQ,121), "
                    "VC_END_SEQ_NUMBER, CONVERT(varchar,DT_END_SEQ,121), IN_QTY, VC_PRODUCTION_DATE, "
                    "VC_LINE_NAME FROM INV_ASN_MST WHERE IN_ASN_ID=%d" % LEGACY_ASN)
    if not hdr:
        rep.check("legacy ASN %d present in Inventory_Live" % LEGACY_ASN, False, "no header row")
        sys.exit(rep.summary_exit())
    seqStart, dtStart, seqLast, dtEnd, legQty, legPDate, legLine = hdr[0]
    legQty = int(legQty)
    rep.check("legacy target = COROLLA / 20260618", legLine == LINE and legPDate == PDATE,
              "line=%s pdate=%s" % (legLine, legPDate))
    legacy = details_map(LIVE, LEGACY_ASN)
    legTotal = _total_qty(legacy)
    rep.check("legacy ASN %d has detail rows" % LEGACY_ASN, len(legacy) > 0, "%d rows" % len(legacy))
    print("    legacy window: %s .. %s  seq %s..%s  hdrQty=%d  detailRows=%d  detailQtySum=%d"
          % (dtStart, dtEnd, seqStart, seqLast, legQty, len(legacy), legTotal))

    # The window-is-the-filter cross-check: AD_FRSPULL's 1-per-vehicle (spare) BC rows in this window
    # sum to the legacy header qty (848) — the date window isolates the build snapshot, no reload leak.
    frsCheck = asn.system.db.runPrepQuery(
        "EXEC AD_FRSPULL @begindate=?, @enddate=?, @Start=?, @Last=?, @LineName=?",
        [dtStart, dtEnd, int(seqStart), int(seqLast), LINE], "VehicleOrder")
    spareVehicles = sum(int(r["VEHICLES"]) for r in frsCheck if int(r["ORDERS"]) == int(r["VEHICLES"]))
    rep.check("AD_FRSPULL window bounds to the legacy header vehicle count (window IS the filter, "
              "no historical-reload leak)", spareVehicles == legQty,
              "spare/1-per-vehicle BC rows sum=%d vs legacy hdrQty=%d" % (spareVehicles, legQty))

    # --- guard: Inventory must be clean for (COROLLA, 20260618) so the idempotency guard lets us run.
    pre = scalar(INV, "SELECT COUNT(*) FROM INV_ASN_MST WHERE VC_LINE_NAME='%s' "
                      "AND VC_PRODUCTION_DATE='%s' AND VC_START_SEQ_NUMBER<>-1" % (LINE, PDATE))
    if int(pre) != 0:
        rep.check("Inventory clean for COROLLA/20260618 before run", False,
                  "%s pre-existing ASN(s) — clean up first" % pre)
        sys.exit(rep.summary_exit())

    # --- 2. run the REAL create_asn -----------------------------------------------------------
    created = None
    try:
        result = asn.create_asn(
            line=LINE, prodDate=PDATE, seqStart=seqStart, seqLast=seqLast,
            beginDate=dtStart, endDate=dtEnd, shipQty=legQty,
            database=DB_CONN, alcDatabase="VehicleOrder")
        created = result["asnId"]
        rep.check("create_asn returned a new ASN id (not skipped)",
                  created is not None and not result["skipped"], "asnId=%s" % created)
        print("    rebuild ASN id=%s  detailRows=%d  qty=%d  missingCostAudit=%d"
              % (created, len(result["details"]), result["qty"], len(result["missingCost"])))

        # --- header sanity (status 'C', EIN 0 = at-SEND decision, qty mirrors caller) ---------
        chdr = sql(INV, "SELECT VC_ASN_STATUS, IN_ASN_EIN, IN_QTY, VC_PRODUCTION_DATE "
                        "FROM INV_ASN_MST WHERE IN_ASN_ID=%d" % created)[0]
        rep.check("rebuild header status = 'C' (created)", chdr[0] == "C", chdr[0])
        rep.check("rebuild header EIN = 0 (EIN allocated at SEND, intended divergence)",
                  int(chdr[1]) == 0, chdr[1])
        rep.check("rebuild header prodDate = 20260618", chdr[3] == PDATE, chdr[3])

        # --- 3. read back the created rows ----------------------------------------------------
        rebuilt = details_map(INV, created)
        rebTotal = _total_qty(rebuilt)

        # --- (A) DRIVER SELF-CONSISTENCY (persisted == computeAsnDetails over the SAME inputs) ----
        # The part the driver controls: given the spike's AD_FRSPULL + forecast, does every row the
        # driver WROTE equal what computeAsnDetails (reviewed pure logic: branch + banker's round +
        # manifest gen) produces? Re-fetch the inputs through the shim and recompute INDEPENDENTLY of
        # the write path, then diff. A clean pass proves the wiring is faithful — decoupled from the
        # legacy recipe-vintage drift. This is SELF-CONSISTENCY, explicitly NOT legacy parity.
        print("\n--- (A) driver self-consistency (persisted rows vs computeAsnDetails over spike inputs) ---")
        frsRows = [{"BC": (r["BC"] or "").strip(), "Orders": int(r["ORDERS"]),
                    "VEHICLES": int(r["VEHICLES"])} for r in frsCheck]
        effMonth = PDATE[0:4] + "/" + PDATE[4:6]
        fmap = {}
        for fr in frsRows:
            if fr["BC"] in fmap:
                continue
            fcDs = asn.system.db.runPrepQuery(
                "EXEC SELECT_ForecastDetailBCASN @BCode=?, @EffMonth=?", [fr["BC"], effMonth], INV)
            fmap[fr["BC"]] = [{
                "VC_ASSY_PART_NUMBER_CODE": r["VC_ASSY_PART_NUMBER_CODE"],
                "VC_ASSY_MANIFEST_NUMBER": r["VC_ASSY_MANIFEST_NUMBER"],
                "IN_ASSY_QTY": r["IN_ASSY_QTY"], "IN_TIRE_RATIO": r["IN_TIRE_RATIO"],
                "IN_WHEEL_RATIO": r["IN_WHEEL_RATIO"], "IN_MANIFEST_COST_ID": r["IN_MANIFEST_COST_ID"],
            } for r in fcDs]
        expected = asn.computeAsnDetails(frsRows, fmap, PDATE, created)
        # accumulate expected per manifest the way INSERT_ASNDetail does (same manifest sums within ASN)
        exp_map = {}
        for d in expected:
            p, q = exp_map.get(d["manifest"], (d["partNumber"], 0))
            exp_map[d["manifest"]] = (d["partNumber"], q + int(d["qty"]))
        sc_ok = (exp_map == rebuilt)
        for m in sorted(set(exp_map) | set(rebuilt)):
            tag = "OK " if exp_map.get(m) == rebuilt.get(m) else "DIF"
            print("    [%s] %s  expected=%s persisted=%s" % (tag, m, exp_map.get(m), rebuilt.get(m)))
        rep.check("driver persisted rows == computeAsnDetails over the spike's own AD_FRSPULL+forecast "
                  "(SELF-CONSISTENCY: wiring faithful — cross-DB read, banker's round, accumulate "
                  "upsert; NOT legacy row parity)", sc_ok,
                  "%d/%d manifests self-consistent" % (
                      sum(1 for m in (set(exp_map) | set(rebuilt)) if exp_map.get(m) == rebuilt.get(m)),
                      len(set(exp_map) | set(rebuilt))))

        # --- (B) TOTAL-QTY INVARIANT (rebuild total == legacy total) ------------------------------
        # Distribution-independent cross-check against the frozen legacy data: the fan-out conserves
        # the same grand total (4240) even though the per-manifest split shifted with the recipe
        # vintage. A REAL parity signal that survives the vintage drift.
        rep.check("rebuild total detail qty == legacy total detail qty (distribution-independent "
                  "invariant — %d == %d)" % (rebTotal, legTotal), rebTotal == legTotal,
                  "rebuild=%d legacy=%d" % (rebTotal, legTotal))

        # --- INFORMATIONAL: per-manifest legacy diff (REPORTED, NOT gated — recipe-vintage drift) ---
        print("\n--- (informational) per-manifest rebuild vs frozen Inventory_Live ASN %d ---" % LEGACY_ASN)
        all_manifests = sorted(set(legacy) | set(rebuilt))
        matches = qtymis = legonly = rebonly = 0
        for m in all_manifests:
            lp = legacy.get(m)
            rp = rebuilt.get(m)
            if lp and rp:
                if lp[0] == rp[0] and lp[1] == rp[1]:
                    matches += 1
                    print("    [OK ] %s  part=%s qty=%d" % (m, lp[0], lp[1]))
                else:
                    qtymis += 1
                    print("    [DIF] %s  legacy(part=%s qty=%d)  rebuild(part=%s qty=%d)"
                          % (m, lp[0], lp[1], rp[0], rp[1]))
            elif lp and not rp:
                legonly += 1
                print("    [LEG-ONLY] %s  legacy(part=%s qty=%d)" % (m, lp[0], lp[1]))
            else:
                rebonly += 1
                print("    [REB-ONLY] %s  rebuild(part=%s qty=%d)" % (m, rp[0], rp[1]))

        total = len(all_manifests)
        rate = (100.0 * matches / total) if total else 0.0
        print("\n    per-manifest: %d matched / %d qty-mismatch / %d legacy-only / %d rebuild-only "
              "(of %d) = %.1f%%  [REPORTED ONLY]" % (matches, qtymis, legonly, rebonly, total, rate))
        print("    NOTE — this per-manifest diff is NOT a driver fault and NOT VehicleOrder drift:")
        print("      * The AD_FRSPULL date window already bounds the spike VehicleOrder to %d vehicles"
              % legQty)
        print("        (= the legacy header qty). The window IS the filter — no historical-reload leak.")
        print("      * Total detail qty is IDENTICAL (legacy %d == rebuild %d); only the per-manifest"
              % (legTotal, rebTotal))
        print("        distribution differs — inconsistent with any vehicle-count drift.")
        print("      * SELECT_ForecastDetailBCASN (parts/ratios/assy/manifest ids) is BYTE-IDENTICAL")
        print("        between Inventory and Inventory_Live for 2026/06 (ratios ruled out).")
        print("      * The diffs are mathematically IMPOSSIBLE from vehicle drift: e.g. NBB today is")
        print("        assy=4 tire 40/20/40 -> 76061857/58/59; the legacy trio 80/900/1124 would need")
        print("        VEHICLES = 50, 1125 and 702.5 (fractional/inconsistent) — no single integer")
        print("        count yields all three. Same shape for NJJ/NCC.")
        print("      ==> Legacy ASN %d was FROZEN under a DIFFERENT forecast-recipe vintage in" % LEGACY_ASN)
        print("          Inventory_Live's OWN history; the ratios/assy/manifest-mapping then differ from")
        print("          today's DB. No currently reproducible legacy ASN (incl. the most-recent normal")
        print("          ones 4718-4721; 4722 is a hot-call/manual ASN) matches the current recipe, so")
        print("          there is no honest ROW-FOR-ROW parity assertion to make today.")
        # IG-TODO: if a future legacy ASN is created under the CURRENT forecast recipe (run this probe
        # against max(IN_ASN_ID) normal ASN), it will match row-for-row — upgrade this informational
        # block into a gated `rebuilt == legacy` parity assertion at that point.

    finally:
        # --- 5. restore Inventory as-found ----------------------------------------------------
        if created is not None:
            sql(INV, "DELETE FROM INV_ASN_DETAIL_MST WHERE IN_ASN_ID=%d; "
                     "DELETE FROM INV_ASN_MST WHERE IN_ASN_ID=%d" % (created, created))
        post = scalar(INV, "SELECT COUNT(*) FROM INV_ASN_MST WHERE VC_LINE_NAME='%s' "
                           "AND VC_PRODUCTION_DATE='%s' AND VC_START_SEQ_NUMBER<>-1" % (LINE, PDATE))
        rep.check("Inventory restored as-found (test ASN swept)", int(post) == 0,
                  "%s remaining" % post)

    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
