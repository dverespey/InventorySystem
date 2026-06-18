#!/usr/bin/env python3
"""test_manifest_cost_lookup.py — the shared window-aware manifest-cost lookup (D6 / D11#1).

Proves dbo.fn_ManifestCostAt(partCode, date) picks the cost row whose manifest window COVERS the date,
and that this fixes the window-blind row-multiplication the legacy REPORT_INVOICESSummary /
REPORT_MonthlyINVOICESSummary / REPORT_EDI810 paths exhibit once a part has more than one window.

The live data has exactly one window per part (45/45), so the bug is LATENT. This test inserts a
synthetic part with TWO non-overlapping windows (a price change) and asserts:
  * the lookup returns the OLD price for a date in the old window, the NEW price for the new window,
    and NO row for a date outside every window;
  * a window-BLIND join (part code only — the legacy shape) returns BOTH rows for one date = the bug;
  * the window-aware lookup returns exactly ONE row for that same date = the fix.
Fixture-disciplined (deletes only its synthetic rows).

Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 scripts/e2e/test_manifest_cost_lookup.py
"""
import os, subprocess, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report  # noqa: E402

CONTAINER = os.environ.get("CONTAINER", "mssql-spike")
SA_PASS = os.environ.get("SA_PASS")
DB = "Inventory"
PARTC = "ZZTESTCOST1"           # synthetic part: two NON-overlapping windows (the normal price-change)
PARTC2 = "ZZTESTCOST2"          # synthetic part: two OVERLAPPING windows (a data error - Q-overlap)
M_OLD, M_NEW = "Z1", "Z2"       # synthetic manifest numbers
P_OLD, P_NEW = "100.0000", "150.0000"
# two NON-overlapping windows: old 2024, new 2025+
W_OLD = ("20240101", "20241231")
W_NEW = ("20250101", "20281231")
# two OVERLAPPING windows for PARTC2 (overlap across 2025): exercises the TVF-vs-legacy divergence
W_OV1 = ("20240101", "20251231")
W_OV2 = ("20250101", "20281231")


def sql(query):
    if not SA_PASS:
        sys.exit("export SA_PASS first")
    out = subprocess.check_output([
        "docker", "exec", CONTAINER, "/opt/mssql-tools18/bin/sqlcmd", "-C", "-S", "localhost",
        "-U", "sa", "-P", SA_PASS, "-d", DB, "-h", "-1", "-W", "-s", "\t",
        "-Q", "SET NOCOUNT ON; " + query], text=True)
    return [l.split("\t") for l in out.splitlines()
            if l.strip() and not l.startswith("(") and not l.startswith("Msg ")]


def scalar(query):
    r = sql(query)
    return r[0][0] if r else None


def price_at(date):
    return scalar("SELECT MO_PRICE FROM dbo.fn_ManifestCostAt('%s', '%s')" % (PARTC, date))


def rowcount_at(date):
    return int(scalar("SELECT COUNT(*) FROM dbo.fn_ManifestCostAt('%s', '%s')" % (PARTC, date)) or 0)


def cleanup():
    sql("DELETE FROM INV_MANIFEST_COST_MST WHERE VC_ASSY_PART_NUMBER_CODE IN ('%s','%s')" % (PARTC, PARTC2))


def legacy_windowed_count(part, date):
    """The legacy REPORT_EDI856 window-filtered JOIN shape (part + window), as a row count."""
    return int(scalar("SELECT COUNT(*) FROM INV_MANIFEST_COST_MST WHERE VC_ASSY_PART_NUMBER_CODE='%s' "
                      "AND VC_START_MANIFEST <= '%s' AND VC_END_MANIFEST >= '%s'" % (part, date, date)) or 0)


def legacy_windowed_price(part, date):
    return scalar("SELECT MO_PRICE FROM INV_MANIFEST_COST_MST WHERE VC_ASSY_PART_NUMBER_CODE='%s' "
                  "AND VC_START_MANIFEST <= '%s' AND VC_END_MANIFEST >= '%s'" % (part, date, date))


def main():
    print("=" * 78)
    print(" MANIFEST-COST window-aware lookup — fn_ManifestCostAt (D6/D11#1)")
    print("=" * 78)
    rep = Report()
    cleanup()  # in case a prior run aborted
    try:
        for mfst, win, price in ((M_OLD, W_OLD, P_OLD), (M_NEW, W_NEW, P_NEW)):
            sql("INSERT INTO INV_MANIFEST_COST_MST (VC_ASSY_PART_NUMBER_CODE, VC_ASSY_MANIFEST_NUMBER, "
                "VC_START_MANIFEST, VC_END_MANIFEST, MO_PRICE) VALUES ('%s','%s','%s','%s',%s)"
                % (PARTC, mfst, win[0], win[1], price))

        # window-aware: the right price for a date in each window
        rep.check("date in OLD window -> old price",
                  price_at("20240615") == P_OLD, "got %s (want %s)" % (price_at("20240615"), P_OLD))
        rep.check("date in NEW window -> new price",
                  price_at("20260615") == P_NEW, "got %s (want %s)" % (price_at("20260615"), P_NEW))
        rep.check("boundary dates are inclusive (<=, >=)",
                  price_at("20241231") == P_OLD and price_at("20250101") == P_NEW,
                  "%s / %s" % (price_at("20241231"), price_at("20250101")))
        rep.check("date OUTSIDE every window -> no row",
                  rowcount_at("20230101") == 0 and rowcount_at("20290101") == 0,
                  "before=%d after=%d" % (rowcount_at("20230101"), rowcount_at("20290101")))
        rep.check("lookup returns exactly ONE row for an in-window date (no multiplication)",
                  rowcount_at("20260615") == 1, "rows=%d" % rowcount_at("20260615"))

        # the bug it fixes: a window-BLIND join (legacy shape, part code only) double-counts
        blind = int(scalar("SELECT COUNT(*) FROM INV_MANIFEST_COST_MST WHERE VC_ASSY_PART_NUMBER_CODE='%s'" % PARTC))
        rep.check("window-BLIND join (part code only) returns BOTH windows = the legacy bug",
                  blind == 2, "blind rows=%d" % blind)
        rep.check("window-aware lookup returns 1 where blind returns 2 (D6 fix proven)",
                  rowcount_at("20260615") == 1 and blind == 2)

        # NON-OVERLAP parity: over a clean (non-overlapping) part the TVF == the legacy WINDOWED JOIN
        # (the EDI856 @EIN=0 shape) — proves the CROSS APPLY migration preserves report output where
        # windows are well-formed (the normal price-change case).
        rep.check("non-overlap: TVF price == legacy windowed-JOIN price (old window)",
                  price_at("20240615") == legacy_windowed_price(PARTC, "20240615") and
                  legacy_windowed_count(PARTC, "20240615") == 1, "TVF/legacy agree, 1 row")
        rep.check("non-overlap: TVF price == legacy windowed-JOIN price (new window)",
                  price_at("20260615") == legacy_windowed_price(PARTC, "20260615") and
                  legacy_windowed_count(PARTC, "20260615") == 1)

        # OVERLAP divergence (Q-overlap, made EXPLICIT not hidden): on overlapping windows the legacy
        # windowed JOIN (and EDI856's GROUP-BY-on-price) returns BOTH rows = row-multiply; the TVF's
        # TOP 1 returns ONE. This is the intended dedup PENDING David's overlap policy — surfaced here.
        for mfst, win in (("Y1", W_OV1), ("Y2", W_OV2)):
            sql("INSERT INTO INV_MANIFEST_COST_MST (VC_ASSY_PART_NUMBER_CODE, VC_ASSY_MANIFEST_NUMBER, "
                "VC_START_MANIFEST, VC_END_MANIFEST, MO_PRICE) VALUES ('%s','%s','%s','%s',%s)"
                % (PARTC2, mfst, win[0], win[1], P_OLD))
        OVDATE = "20250615"   # in BOTH overlapping windows
        rep.check("overlap: legacy windowed JOIN returns 2 rows (multiply)",
                  legacy_windowed_count(PARTC2, OVDATE) == 2, "legacy=%d" % legacy_windowed_count(PARTC2, OVDATE))
        rep.check("overlap: TVF returns exactly 1 (deterministic dedup; DIVERGES from legacy — Q-overlap)",
                  int(scalar("SELECT COUNT(*) FROM dbo.fn_ManifestCostAt('%s','%s')" % (PARTC2, OVDATE))) == 1)
    finally:
        cleanup()

    leftover = int(scalar("SELECT COUNT(*) FROM INV_MANIFEST_COST_MST WHERE VC_ASSY_PART_NUMBER_CODE='%s'" % PARTC) or 0)
    rep.check("fixture clean (synthetic cost rows removed)", leftover == 0, "leftover=%d" % leftover)
    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
