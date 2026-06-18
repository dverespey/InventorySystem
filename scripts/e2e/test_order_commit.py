#!/usr/bin/env python3
"""test_order_commit.py — unit test for the Order commit/write path (computeOrderRecords).

Pure: imports `order.computeOrderRecords` straight from the on-disk project library (no DB, no env).
Verifies the lot-sized vs palletized branching, FRS# derivation, renban assignment + counter (with
999 rollover), against the legacy Order.pas:683-849 logic and the project-order-renban domain model.

Reminder: BIT_LOT_SIZE_ORDERS is INVERTED — 0 = lot-sized (loop per lot), 1 = palletized (one record).

Run:  python3 scripts/e2e/test_order_commit.py
"""
import os, sys, importlib.util

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report  # noqa: E402

ORDER = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                                      "docs", "analysis", "order", "project-library", "order", "code.py"))


def _load():
    spec = importlib.util.spec_from_file_location("order_lib", ORDER)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


order = _load()
cp = order.computeOrderRecords
FRSD = "20260615"   # order-by date -> FRS prefix '60615' (yyyymmdd[3:8]) -> FRS '6061501'...


def main():
    print("=" * 78)
    print(" ORDER commit/write path — computeOrderRecords unit test (Order.pas:683-849)")
    print("=" * 78)
    rep = Report()

    # --- LOT-SIZED (bit 0): one record per lot, sequential FRS + sequential renban ---------------
    print("\n--- LOT-SIZED (bit 0) ---")
    recs, newc = cp("0572B", "4265202R9000", "16H", 0, "", lotCount=2, qtyQ=80, lotQty=40,
                    frsDateYmd=FRSD, renbanCount=6)
    rep.check("lot-sized 2 lots -> 2 records", len(recs) == 2, str([r["frsNum"] for r in recs]))
    rep.check("sequential FRS '6061501','6061502'",
              [r["frsNum"] for r in recs] == ["6061501", "6061502"], str([r["frsNum"] for r in recs]))
    rep.check("each record qty = lotQty (not Q)", all(r["qty"] == 40 for r in recs), str([r["qty"] for r in recs]))
    rep.check("sequential renban 16H006/16H007 (kanban + 3-digit counter)",
              [r["renbanNum"] for r in recs] == ["16H006", "16H007"], str([r["renbanNum"] for r in recs]))
    rep.check("renban counter advanced by 2 (6 -> 8)", newc == 8, "newCount=%d" % newc)

    recs1, _ = cp("0572B", "P", "16H", 0, "", lotCount=1, qtyQ=40, lotQty=40, frsDateYmd=FRSD, renbanCount=6)
    rep.check("lot-sized 1 lot -> 1 record FRS '...01'", len(recs1) == 1 and recs1[0]["frsNum"] == "6061501")

    recs11, _ = cp("0572B", "P", "16H", 0, "", lotCount=11, qtyQ=440, lotQty=40, frsDateYmd=FRSD, renbanCount=1)
    rep.check("lot-sized 11 lots -> 2-char suffixes incl '10','11'",
              recs11[9]["frsNum"] == "6061510" and recs11[10]["frsNum"] == "6061511",
              "%s / %s" % (recs11[9]["frsNum"], recs11[10]["frsNum"]))

    # --- PALLETIZED (bit 1): ONE record, qty = Q, renban BLANK when grouped --------------------
    print("\n--- PALLETIZED (bit 1) ---")
    recs, newc = cp("0572B", "4261102Q8000", "M1", 1, "CMWA", lotCount=3, qtyQ=120, lotQty=40,
                    frsDateYmd=FRSD, renbanCount=6)
    rep.check("palletized -> exactly ONE record (regardless of lotCount)", len(recs) == 1, str(len(recs)))
    rep.check("palletized record qty = Q (whole order)", recs[0]["qty"] == 120, str(recs[0]["qty"]))
    rep.check("palletized FRS suffix '01'", recs[0]["frsNum"] == "6061501", recs[0]["frsNum"])
    rep.check("renban-grouped -> BLANK renban (assigned later by RenbanOrder)", recs[0]["renbanNum"] == "")
    rep.check("renban counter UNTOUCHED for grouped part (6)", newc == 6, "newCount=%d" % newc)

    # palletized WITHOUT a renban group -> falls back to the counter (defensive legacy path)
    recsng, newcng = cp("0572B", "P", "M1", 1, "", lotCount=3, qtyQ=120, lotQty=40, frsDateYmd=FRSD, renbanCount=6)
    rep.check("palletized + NO group -> 1 record, counter renban used",
              len(recsng) == 1 and recsng[0]["renbanNum"] == "M1006" and newcng == 7, str(recsng))

    # --- renban counter rollover (>999 -> 1) -----------------------------------------------------
    print("\n--- renban counter rollover ---")
    recr, newcr = cp("0572B", "P", "16H", 0, "", lotCount=2, qtyQ=80, lotQty=40, frsDateYmd=FRSD, renbanCount=999)
    rep.check("renban uses 999 then rolls to 1 (16H999, 16H001)",
              [r["renbanNum"] for r in recr] == ["16H999", "16H001"], str([r["renbanNum"] for r in recr]))
    rep.check("counter after rollover = 2", newcr == 2, "newCount=%d" % newcr)

    # --- FRS prefix derivation (copy(yymmdd,2,5)) ------------------------------------------------
    print("\n--- FRS prefix ---")
    r, _ = cp("S", "P", "K", 1, "", lotCount=1, qtyQ=1, lotQty=1, frsDateYmd="20251231", renbanCount=0)
    rep.check("frsDate 20251231 -> FRS '5123101' (yyyymmdd[3:8] + '01')", r[0]["frsNum"] == "5123101", r[0]["frsNum"])

    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
