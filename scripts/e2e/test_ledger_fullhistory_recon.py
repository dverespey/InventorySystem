#!/usr/bin/env python3
"""test_ledger_fullhistory_recon.py — close the from-zero reconciliation GO/NO-GO that the purge
horizon left SKIPped in test_stock_ledger_parity.py.

THE PROBLEM (design §6 / the parity harness banner): the restored Inventory.bak is post-DELETE_AutoPurge
— 0 counting open-orders survive, so a from-zero replay of the live snapshot cannot reconstruct the
legacy IN_QTY (the receiving IN-movements that built it were aged out). The reconciliation DERIVATION
was therefore exercised but never proven to equal IN_QTY on a real balance. There is no pre-purge dump
in this environment (only the same post-purge .bak), so the production sign-off number must wait for a
real pre-purge backup / the cutover-window snapshot (the harness IS that validator).

WHAT THIS CLOSES: the derivation LOGIC, on a CONTROLLED COMPLETE history immune to the purge horizon.
For one part we create a known, marker-isolated set of movements across ALL FOUR producers via the live
triggers (so legacy IN_QTY moves by exactly their sum), then run the §3 from-zero derivation over ONLY
those marked rows and assert it reconstructs the exact IN_QTY delta. That is the GO/NO-GO ("does the
from-zero replay == IN_QTY when history is complete?") answered YES on data we fully control — the
purge horizon is the only thing between this and a full-snapshot PASS, and that's a DATA gap, not a
logic gap. Fixture-disciplined (restores the part + removes all marked rows + HIST).

Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 scripts/e2e/test_ledger_fullhistory_recon.py
"""
import os, subprocess, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report  # noqa: E402

CONTAINER = os.environ.get("CONTAINER", "mssql-spike")
SA_PASS = os.environ.get("SA_PASS")
DB = "Inventory"
PART = 16
PARTNUM = "478930223000"        # 'S' add-point
RENBAN = "FHRECON1"             # open-order marker
REASON = "FHRECON"              # reject + stocktaking marker
PDATE = "29991228"             # shipping marker (production date)
LINE = "FHRLINE"                # shipment-header marker (INV_PART_SHIPPING_INF has a real FK to the header)
SHIP = "20260618"               # non-blank ship status => counts for 'S'
# Known complete history (signed contributions to on-hand):
Q_OO, Q_REJ, Q_STK, Q_SHP = 100, 30, 10, 20   # +100 (recv) -30 (rej) +10 (stk) -20 (shp) = +60
EXPECT = Q_OO - Q_REJ + Q_STK - Q_SHP

STAMP = ("CONVERT(varchar,getdate(),112)+SUBSTRING(CONVERT(varchar,getdate(),114),1,2)"
         "+SUBSTRING(CONVERT(varchar,getdate(),114),4,2)+SUBSTRING(CONVERT(varchar,getdate(),114),7,2)"
         "+SUBSTRING(CONVERT(varchar,getdate(),114),10,2)")
OO_INS = ("INSERT INTO INV_OPEN_ORDER_INF (VC_SUPPLIER_CODE, VC_PART_NUMBER, VC_FRS_NUMBER, "
          "VC_RENBAN_NUMBER, IN_QTY, VC_STATUS_SUPPLIER_SHIPPING, VC_ARRIVAL, VC_TRAILER_NUMBER, "
          "VC_STATUS_PLANT_YARD, VC_PLANT_PARKING, VC_STATUS_ASSEMBLER_YARD, VC_ASSEMBLER_LOCATION, "
          "VC_STATUS_EMPTY_TRAILER, VC_DETENTION, VC_ORDER_DATE, VC_WAREHOUSE, VC_TERMINATED, "
          "VC_SHIP_DATE, VC_KANBAN_NUMBER, VC_LAST_UPDATE, VC_ADD) VALUES "
          "('FHR', '%s', 'FHR0001', '%s', %d, '%s', '', '', '', '', '', '', '', '', '', '', '', '', '', %s, %s)")

# §3 from-zero derivation over ONLY the marked rows, summed per the producer signs (mirrors the
# parity harness DERIVE_SQL, scoped to this part's complete controlled history).
DERIVE = """
SELECT
 (SELECT ISNULL(SUM(CASE WHEN s.VC_INVENTORY_ADD_POINT='S' AND oo.VC_STATUS_SUPPLIER_SHIPPING<>'' THEN oo.IN_QTY
                         WHEN s.VC_INVENTORY_ADD_POINT='A' AND (oo.VC_ARRIVAL<>'' OR oo.VC_STATUS_PLANT_YARD<>''
                              OR oo.VC_STATUS_ASSEMBLER_YARD<>'' OR oo.VC_WAREHOUSE<>'') THEN oo.IN_QTY ELSE 0 END),0)
  FROM INV_OPEN_ORDER_INF oo JOIN INV_PARTS_STOCK_MST ps ON ps.VC_PART_NUMBER=oo.VC_PART_NUMBER
       JOIN INV_SUPPLIER_MST s ON ps.IN_SUPPLIER_ID=s.IN_SUPPLIER_ID
  WHERE oo.VC_RENBAN_NUMBER='%s')
 - (SELECT ISNULL(SUM(IN_QTY),0) FROM INV_REJECT_INF WHERE IN_PART_ID=%d AND VC_REASON='%s')
 + (SELECT ISNULL(SUM(IN_QTY),0) FROM INV_STOCKTAKING_INF WHERE IN_PART_ID=%d AND VC_REASON='%s')
 - (SELECT ISNULL(SUM(s.IN_QTY),0) FROM INV_PART_SHIPPING_INF s
    JOIN INV_PARTS_STOCK_MST ps ON ps.VC_PART_NUMBER=s.VC_PART_NUMBER
    WHERE s.VC_PRODUCTION_DATE='%s' AND ps.IN_PART_ID=%d)
"""


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


def qty():
    return int(scalar("SELECT IN_QTY FROM INV_PARTS_STOCK_MST WHERE IN_PART_ID=%d" % PART))


def cleanup(base):
    # remove marked rows with triggers OFF (so the deletes don't reverse IN_QTY), then restore + re-enable.
    sql("DISABLE TRIGGER DELETE_RejectParts ON INV_REJECT_INF; "
        "DISABLE TRIGGER DELETE_Stocktaking ON INV_STOCKTAKING_INF; "
        "DISABLE TRIGGER DeletePartShipping ON INV_PART_SHIPPING_INF; "
        "DISABLE TRIGGER DeleteShipDate ON INV_SHIPPING_INF; "
        "DISABLE TRIGGER DELETE_RecConfStatPartsStockMstQTY ON INV_OPEN_ORDER_INF; "
        "DELETE FROM INV_OPEN_ORDER_INF WHERE VC_RENBAN_NUMBER='%s'; "
        "DELETE FROM INV_OPEN_ORDER_INF_HIST WHERE VC_RENBAN_NUMBER='%s'; "
        "DELETE FROM INV_REJECT_INF WHERE VC_REASON='%s'; "
        "DELETE FROM INV_STOCKTAKING_INF WHERE VC_REASON='%s'; "
        "DELETE FROM INV_PART_SHIPPING_INF WHERE VC_PRODUCTION_DATE='%s'; "
        "DELETE FROM INV_SHIPPING_INF WHERE VC_LINE_NAME='%s'; "
        "UPDATE INV_PARTS_STOCK_MST SET IN_QTY=%d WHERE IN_PART_ID=%d; "
        "ENABLE TRIGGER DELETE_RejectParts ON INV_REJECT_INF; "
        "ENABLE TRIGGER DELETE_Stocktaking ON INV_STOCKTAKING_INF; "
        "ENABLE TRIGGER DeletePartShipping ON INV_PART_SHIPPING_INF; "
        "ENABLE TRIGGER DeleteShipDate ON INV_SHIPPING_INF; "
        "ENABLE TRIGGER DELETE_RecConfStatPartsStockMstQTY ON INV_OPEN_ORDER_INF; "
        "ENABLE TRIGGER UPDATE_PartNumber ON INV_PARTS_STOCK_MST"
        % (RENBAN, RENBAN, REASON, REASON, PDATE, LINE, base, PART))


def main():
    print("=" * 78)
    print(" FULL-HISTORY RECONCILIATION — from-zero §3 derivation == IN_QTY on complete controlled history")
    print("=" * 78)
    rep = Report()
    base = qty()
    print("  part %d base IN_QTY = %d ; building a complete %d-movement history (expect delta %+d)"
          % (PART, base, 4, EXPECT))
    sql("DISABLE TRIGGER UPDATE_PartNumber ON INV_PARTS_STOCK_MST")   # suppress audit-row noise
    try:
        # Build a KNOWN complete history across all four producers via the live triggers.
        sql(OO_INS % (PARTNUM, RENBAN, Q_OO, SHIP, STAMP, STAMP))                       # +Q_OO  (recv 'S')
        sql("INSERT INTO INV_REJECT_INF (VC_DIVISION, IN_PART_ID, IN_QTY, VC_REASON, VC_LAST_UPDATE, VC_ADD) "
            "VALUES ('T', %d, %d, '%s', %s, %s)" % (PART, Q_REJ, REASON, STAMP, STAMP))  # -Q_REJ
        sql("INSERT INTO INV_STOCKTAKING_INF (IN_PART_ID, IN_QTY, VC_REASON, VC_LAST_UPDATE, VC_ADD) "
            "VALUES (%d, %d, '%s', %s, %s)" % (PART, Q_STK, REASON, STAMP, STAMP))        # +Q_STK
        shipId = scalar("INSERT INTO INV_SHIPPING_INF (VC_LINE_NAME, VC_START_SEQ_NUMBER, VC_END_SEQ_NUMBER, "
                        "VC_PRODUCTION_DATE, VC_ADD) OUTPUT INSERTED.IN_SHIPPING_ID VALUES "
                        "('%s', '0', '0', '%s', %s)" % (LINE, PDATE, STAMP))             # real header (FK target)
        sql("INSERT INTO INV_PART_SHIPPING_INF (IN_SHIPPING_ID, VC_PART_NUMBER, VC_PRODUCTION_DATE, IN_QTY, VC_ADD) "
            "VALUES (%s, '%s', '%s', %d, %s)" % (shipId, PARTNUM, PDATE, Q_SHP, STAMP))   # -Q_SHP

        legacy_delta = qty() - base
        derived = int(scalar(DERIVE % (RENBAN, PART, REASON, PART, REASON, PDATE, PART)))
        rep.check("legacy triggers moved IN_QTY by the expected complete-history sum",
                  legacy_delta == EXPECT, "delta %+d (expect %+d)" % (legacy_delta, EXPECT))
        rep.check("from-zero §3 derivation reconstructs the IN_QTY delta EXACTLY (GO/NO-GO logic)",
                  derived == legacy_delta, "derived %+d == legacy %+d" % (derived, legacy_delta))
        rep.check("reconciliation is non-vacuous (covers recv+reject+stk+shipping, delta != 0)",
                  legacy_delta == EXPECT and EXPECT != 0, "delta %+d" % legacy_delta)
    finally:
        cleanup(base)

    restored = qty()
    leftover = (int(scalar("SELECT COUNT(*) FROM INV_OPEN_ORDER_INF WHERE VC_RENBAN_NUMBER='%s'" % RENBAN))
                + int(scalar("SELECT COUNT(*) FROM INV_REJECT_INF WHERE VC_REASON='%s'" % REASON))
                + int(scalar("SELECT COUNT(*) FROM INV_STOCKTAKING_INF WHERE VC_REASON='%s'" % REASON))
                + int(scalar("SELECT COUNT(*) FROM INV_PART_SHIPPING_INF WHERE VC_PRODUCTION_DATE='%s'" % PDATE)))
    rep.check("DB restored as found (IN_QTY, no marked rows)", restored == base and leftover == 0,
              "IN_QTY=%d leftover=%d" % (restored, leftover))

    print("\n  NOTE: this proves the reconciliation LOGIC on complete history. The FULL-SNAPSHOT")
    print("  sign-off (derive == live IN_QTY across all parts) still needs a real PRE-PURGE dump or")
    print("  the cutover-window snapshot — unavailable in this env (post-DELETE_AutoPurge). When that")
    print("  data is provided, test_stock_ledger_parity.py's skipped from-zero check becomes the gate.")
    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
