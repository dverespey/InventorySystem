#!/usr/bin/env python3
"""
SC1 number-parity harness — SIM_OrderSimulation vs the golden Delphi exports.

Compares the proc's per-part phased output (Section B + C dumps) against the
golden 4-row ledger (Beg / Receipts-per-supplier / Usage / End) decoded from
DB Schema/OrderSimulationCorolla{Tire,Wheel,Valve,Film}.xls (converted to .xlsx
in /tmp/golden). The golden pools Beg/Usage/End at the size-group level; the proc
emits per-part rows, so the pooled rows are reconstructed as the SUM of per-part
values by size group. Golden day column index = 20 + fill_pos (production days,
weekends already absent as columns).

Inputs (regenerate with the dump loop in the session):
  /tmp/secB_<pt>.tsv  Section B grid rows
  /tmp/secC_<pt>.tsv  Section C phased cells
  /tmp/golden/OrderSimulationCorolla<Pt>.xlsx

Usage: python3 parity_diff.py
Exit 0 if all groups PASS, 1 otherwise.
"""
import sys, os, openpyxl
from collections import defaultdict

GOLDEN_DIR = "/tmp/golden"
TSV_DIR = "/tmp"
FIRST_DAY_COL = 20          # col T
NDAYS = 25                  # FillDays
PART_TYPES = [("TIRE", "tire", "Tire"), ("WHEEL", "wheel", "Wheel"),
              ("VALVE", "valve", "Valve"), ("FILM", "film", "Film")]

def num(v):
    if v is None or v == "":
        return 0
    try:
        return int(round(float(v)))
    except (ValueError, TypeError):
        return 0

# ---- proc dump parsers -------------------------------------------------------
# Section B columns (0-based): 0 row_kind,3 size_code,5 sup_code,6 sup_name,
#   7 part,20 lead_time,27 orderby_col_index,28 frs_date
def parse_secB(path):
    parts = {}   # part_number -> {size_code, supplier, lead, orderby, frs}
    with open(path) as f:
        for line in f:
            c = line.rstrip("\n").split("\t")
            if len(c) < 29 or c[0] != "PART":
                continue
            parts[c[7]] = dict(size_code=c[3].strip(), supplier=c[6], lead=num(c[20]),
                               orderby=num(c[27]), frs=c[28])
    return parts

# Section C columns: 0 part,1 fill_pos,2 value,3 forecast_usage,4 receipt_qty,
#   5 balance,6 signal_enum,7 source_enum,8 below_safety
def parse_secC(path):
    bal = defaultdict(dict)   # part -> {fill_pos: balance}
    fc  = defaultdict(dict)   # part -> {fill_pos: forecast_usage}
    rcp = defaultdict(dict)   # part -> {fill_pos: receipt_qty}
    with open(path) as f:
        for line in f:
            c = line.rstrip("\n").split("\t")
            if len(c) < 9:
                continue
            p, fp = c[0], num(c[1])
            fc[p][fp]  = num(c[3])
            rcp[p][fp] = num(c[4])
            bal[p][fp] = num(c[5])
    return bal, fc, rcp

# ---- golden parser -----------------------------------------------------------
def parse_golden(path):
    """Return list of blocks: {size_code, suppliers:[(part,lead)], beg/usage/end:[25],
       receipts:{part:[25]}}."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    def daycols(r):
        return [num(ws.cell(row=r, column=FIRST_DAY_COL + k).value) for k in range(NDAYS)]
    blocks, cur = [], None
    for r in range(7, ws.max_row + 1):
        s = ws.cell(row=r, column=19).value   # col S label
        if s is None:
            continue
        s = str(s).strip()
        if s == "Beg Balance":
            if cur:
                blocks.append(cur)
            cur = dict(size_code=str(ws.cell(row=r, column=2).value).strip(),
                       suppliers=[], receipts={}, beg=daycols(r), usage=None, end=None)
        elif s == "Receipts" and cur is not None:
            part = ws.cell(row=r, column=4).value          # col D
            lead = num(ws.cell(row=r, column=16).value)    # col P
            if part:
                part = str(part).strip()
                cur["suppliers"].append((part, lead))
                cur["receipts"][part] = daycols(r)
        elif s == "Usage" and cur is not None:
            cur["usage"] = daycols(r)
        elif s == "End Balance" and cur is not None:
            cur["end"] = daycols(r)
    if cur:
        blocks.append(cur)
    return blocks

# ---- comparison --------------------------------------------------------------
def cmp_series(label, gold, got, fails):
    bad = [(k, gold[k], got[k]) for k in range(NDAYS) if gold[k] != got[k]]
    if bad:
        fails.append((label, bad))
    return not bad

def run_pt(pt, lc, Pt):
    gold_path = os.path.join(GOLDEN_DIR, f"OrderSimulationCorolla{Pt}.xlsx")
    parts = parse_secB(os.path.join(TSV_DIR, f"secB_{lc}.tsv"))
    bal, fc, rcp = parse_secC(os.path.join(TSV_DIR, f"secC_{lc}.tsv"))
    blocks = parse_golden(gold_path)

    # group proc parts by size_code
    by_size = defaultdict(list)
    for p, meta in parts.items():
        by_size[meta["size_code"]].append(p)

    total = {"end": [0, 0], "usage": [0, 0], "receipts": [0, 0], "orderby": [0, 0]}
    detail = []
    for blk in blocks:
        sz = blk["size_code"]
        grp_parts = by_size.get(sz, [])
        fails = []
        # pooled End = sum per-part balance ; pooled Usage = sum per-part forecast
        sum_bal = [sum(bal[p].get(k, 0) for p in grp_parts) for k in range(NDAYS)]
        sum_fc  = [sum(fc[p].get(k, 0) for p in grp_parts) for k in range(NDAYS)]
        ok_end = cmp_series("End", blk["end"], sum_bal, fails) if blk["end"] else False
        ok_use = cmp_series("Usage", blk["usage"], sum_fc, fails) if blk["usage"] else False
        total["end"][0] += ok_end; total["end"][1] += 1
        total["usage"][0] += ok_use; total["usage"][1] += 1
        # per-supplier receipts + order-by
        for part, lead in blk["suppliers"]:
            grcp = blk["receipts"].get(part, [0] * NDAYS)
            prcp = [rcp[part].get(k, 0) for k in range(NDAYS)]
            ok_r = cmp_series(f"Receipts[{part}]", grcp, prcp, fails)
            total["receipts"][0] += ok_r; total["receipts"][1] += 1
            # order-by: proc orderby fill_pos -> golden col; golden order-by date =
            # the day the supplier's lead lands on (today + lead prod days = fill_pos lead).
            meta = parts.get(part, {})
            ok_o = (meta.get("orderby") == lead)
            total["orderby"][0] += ok_o; total["orderby"][1] += 1
            if not ok_o:
                fails.append((f"OrderBy[{part}]",
                              [(0, f"lead={lead}", f"orderby_fp={meta.get('orderby')}")]))
        detail.append((sz, [p for p, _ in blk["suppliers"]], fails))
    return total, detail

def main():
    grand = {"end": [0, 0], "usage": [0, 0], "receipts": [0, 0], "orderby": [0, 0]}
    any_fail = False
    for pt, lc, Pt in PART_TYPES:
        print(f"\n===== {pt} =====")
        total, detail = run_pt(pt, lc, Pt)
        for sz, parts, fails in detail:
            status = "PASS" if not fails else "FAIL"
            if fails:
                any_fail = True
            print(f"  [{status}] {sz:8s} suppliers={parts}")
            for label, bad in fails:
                shown = ", ".join(f"day{k}: gold={g} got={t}" for k, g, t in bad[:6])
                more = "" if len(bad) <= 6 else f" (+{len(bad)-6} more)"
                print(f"          {label}: {shown}{more}")
        for k in grand:
            grand[k][0] += total[k][0]; grand[k][1] += total[k][1]
    print("\n===== SC1 SUMMARY (passed/total) =====")
    for k in ("end", "usage", "receipts", "orderby"):
        p, t = grand[k]
        print(f"  {k:9s}: {p}/{t}  {'OK' if p == t else 'MISMATCH'}")
    sys.exit(1 if any_fail else 0)

if __name__ == "__main__":
    main()
