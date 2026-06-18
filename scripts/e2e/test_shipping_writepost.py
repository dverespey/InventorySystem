#!/usr/bin/env python3
"""test_shipping_writepost.py — row-level PARITY integration test for the shipping write path.

Proves, against the LIVE spike DB (fixture-disciplined), that the REBUILD write-then-post path moves
INV_PARTS_STOCK_MST.IN_QTY IDENTICALLY to the LEGACY shipping triggers. Shipping is stock-OUT
(effect = -IN_QTY): a part-shipping insert CONSUMES qty, a delete RESTORES it. Two scenarios:

  CORE  — a part-shipping line insert/amend/delete vs InsertPartShipping/UpdatePartShipping/
          DeletePartShipping.
  CASCADE — a shipment header with 2 SAME-PART rows; header delete should restore both. This is a
          DIVERGENCE, empirically demonstrated here: the legacy DeleteShipDate cascade deletes both
          rows in one statement, but DeletePartShipping's `FROM PS, deleted d` cross-join restores
          only ONE arbitrary row (the F3 multi-row under-count, design §6 class 5) -> legacy leaves
          on-hand short by the other row. The rebuild's postHeaderDelete restores BOTH (multi-row-safe).
          "Ledger correct, legacy buggy (F3)."

Two arms per scenario (rebuild path + legacy triggers both write IN_QTY -> can't co-run):
  legacy  — *PartShipping + DeleteShipDate triggers ENABLED; direct DML fires them.
  rebuild — those triggers DISABLED; only POST_StockMovement moves IN_QTY.
UPDATE_PartNumber disabled throughout (no audit residue). Everything restored in finally.

Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 scripts/e2e/test_shipping_writepost.py
"""
import os, subprocess, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report  # noqa: E402

CONTAINER = os.environ.get("CONTAINER", "mssql-spike")
SA_PASS = os.environ.get("SA_PASS")
DB = "Inventory"
PART = 16
PARTNUM = "478930223000"
DCORE, DCASC = "29991229", "29991230"   # unique production dates -> no collision with real rows
LINE = "WPTLINE"
LREASON = "WPTEST_SHIP"
Q1, Q2 = 137, 200
CA, CB = 30, 20                          # cascade: two part rows

STAMP = ("CONVERT(varchar,getdate(),112)+SUBSTRING(CONVERT(varchar,getdate(),114),1,2)"
         "+SUBSTRING(CONVERT(varchar,getdate(),114),4,2)+SUBSTRING(CONVERT(varchar,getdate(),114),7,2)"
         "+SUBSTRING(CONVERT(varchar,getdate(),114),10,2)")


def sql(query):
    if not SA_PASS:
        sys.exit("export SA_PASS first (see scripts/spike-db.sh)")
    out = subprocess.check_output([
        "docker", "exec", CONTAINER, "/opt/mssql-tools18/bin/sqlcmd", "-C", "-S", "localhost",
        "-U", "sa", "-P", SA_PASS, "-d", DB, "-h", "-1", "-W", "-s", "\t",
        "-Q", "SET NOCOUNT ON; " + query], text=True)
    rows = []
    for line in out.splitlines():
        line = line.rstrip("\n")
        if not line or line.startswith("(") or line.startswith("Msg "):
            continue
        rows.append(line.split("\t"))
    return rows


def scalar(query):
    r = sql(query)
    return r[0][0] if r else None


def qty():
    return int(scalar("SELECT IN_QTY FROM INV_PARTS_STOCK_MST WHERE IN_PART_ID=%d" % PART))


def ins_partship(shipId, date, q):
    sql("INSERT INTO INV_PART_SHIPPING_INF (IN_SHIPPING_ID, VC_PART_NUMBER, VC_PRODUCTION_DATE, IN_QTY, VC_ADD) "
        "VALUES (%s, '%s', '%s', %d, %s)" % (shipId, PARTNUM, date, q, STAMP))


def ins_header(date):
    return scalar("INSERT INTO INV_SHIPPING_INF (VC_LINE_NAME, VC_START_SEQ_NUMBER, VC_END_SEQ_NUMBER, "
                  "VC_PRODUCTION_DATE, VC_ADD) OUTPUT INSERTED.IN_SHIPPING_ID VALUES "
                  "('%s', '0', '0', '%s', %s)" % (LINE, date, STAMP))


def post(part, delta, evt):
    sql("EXEC POST_StockMovement @partId=%d, @delta=%d, @sourceEnum='SHIPPING', @sourceRowId=0, "
        "@eventKey='%s', @reason='%s', @site=1, @purge=0" % (part, delta, evt, LREASON))


def cleanup(base):
    sql("DELETE FROM INV_STOCK_LEDGER WHERE VC_REASON='%s'; "
        "DELETE FROM INV_PART_SHIPPING_INF WHERE VC_PRODUCTION_DATE IN ('%s','%s'); "
        "DELETE FROM INV_SHIPPING_INF WHERE VC_LINE_NAME='%s'; "
        "UPDATE INV_PARTS_STOCK_MST SET IN_QTY=%d WHERE IN_PART_ID=%d; "
        "ENABLE TRIGGER InsertPartShipping ON INV_PART_SHIPPING_INF; "
        "ENABLE TRIGGER UpdatePartShipping ON INV_PART_SHIPPING_INF; "
        "ENABLE TRIGGER DeletePartShipping ON INV_PART_SHIPPING_INF; "
        "ENABLE TRIGGER DeleteShipDate ON INV_SHIPPING_INF; "
        "ENABLE TRIGGER UPDATE_PartNumber ON INV_PARTS_STOCK_MST"
        % (LREASON, DCORE, DCASC, LINE, base, PART))


def disable_ship_triggers():
    sql("DISABLE TRIGGER InsertPartShipping ON INV_PART_SHIPPING_INF; "
        "DISABLE TRIGGER UpdatePartShipping ON INV_PART_SHIPPING_INF; "
        "DISABLE TRIGGER DeletePartShipping ON INV_PART_SHIPPING_INF; "
        "DISABLE TRIGGER DeleteShipDate ON INV_SHIPPING_INF")


def core(rep, base, arm, post_fn):
    print("\n--- CORE part-line insert/amend/delete (%s) ---" % arm)
    shipId = ins_header(DCORE)
    ins_partship(shipId, DCORE, Q1)
    psid = scalar("SELECT MAX(IN_PART_SHIPPING_ID) FROM INV_PART_SHIPPING_INF WHERE VC_PRODUCTION_DATE='%s'" % DCORE)
    if post_fn:
        post(PART, -Q1, "SHIPPING:psh=%s:ins" % psid)
    i = qty() - base
    sql("UPDATE INV_PART_SHIPPING_INF SET IN_QTY=%d, VC_ADD=%s WHERE IN_PART_SHIPPING_ID=%s" % (Q2, STAMP, psid))
    if post_fn:
        post(PART, -(Q2 - Q1), "SHIPPING:psh=%s:upd" % psid)
    a = qty() - base - (-Q1)
    if post_fn:
        post(PART, Q2, "SHIPPING:psh=%s:del" % psid)
    sql("DELETE FROM INV_PART_SHIPPING_INF WHERE IN_PART_SHIPPING_ID=%s" % psid)
    d = qty() - base
    rep.check("%s: insert consumes -Q1" % arm, i == -Q1, "%d" % i)
    rep.check("%s: amend -(Q2-Q1)" % arm, a == -(Q2 - Q1), "%d" % a)
    rep.check("%s: delete restores to base" % arm, d == 0, "delta %d" % d)
    return {"i": i, "a": a, "d": d}


def cascade(rep, base, arm, rebuild):
    print("\n--- CASCADE header+2 rows, header delete restores BOTH (%s) ---" % arm)
    shipId = ins_header(DCASC)
    ins_partship(shipId, DCASC, CA)
    ins_partship(shipId, DCASC, CB)
    if rebuild:
        post(PART, -CA, "SHIPPING:psh=cascA:ins")
        post(PART, -CB, "SHIPPING:psh=cascB:ins")
    consumed = qty() - base
    if rebuild:
        # rebuild header delete: restore each row (scoped to IN_SHIPPING_ID), then physically delete.
        for q in (CA, CB):
            post(PART, q, "SHIPPING:psh=casc%d:del" % q)
        sql("DELETE FROM INV_PART_SHIPPING_INF WHERE IN_SHIPPING_ID=%s" % shipId)
        sql("DELETE FROM INV_SHIPPING_INF WHERE IN_SHIPPING_ID=%s" % shipId)
    else:
        # legacy DeleteShipDate cascade: delete the header -> cascade deletes part rows -> restores.
        sql("DELETE FROM INV_SHIPPING_INF WHERE IN_SHIPPING_ID=%s" % shipId)
    restored = qty() - base
    rep.check("%s: 2 rows consumed -(CA+CB)" % arm, consumed == -(CA + CB), "%d" % consumed)
    return {"consumed": consumed, "restored": restored}


def main():
    print("=" * 78)
    print(" SHIPPING write-then-post — row-level PARITY integration test (rebuild == legacy)")
    print("=" * 78)
    rep = Report()
    base = qty()
    print("  part %d base IN_QTY = %d" % (PART, base))
    sql("DISABLE TRIGGER UPDATE_PartNumber ON INV_PARTS_STOCK_MST")
    try:
        legc = core(rep, base, "LEGACY", post_fn=False)
        sql("UPDATE INV_PARTS_STOCK_MST SET IN_QTY=%d WHERE IN_PART_ID=%d" % (base, PART))
        legx = cascade(rep, base, "LEGACY", rebuild=False)
        sql("UPDATE INV_PARTS_STOCK_MST SET IN_QTY=%d WHERE IN_PART_ID=%d" % (base, PART))
        disable_ship_triggers()
        rebc = core(rep, base, "REBUILD", post_fn=True)
        sql("UPDATE INV_PARTS_STOCK_MST SET IN_QTY=%d WHERE IN_PART_ID=%d" % (base, PART))
        rebx = cascade(rep, base, "REBUILD", rebuild=True)
        print("\n--- PARITY ---")
        rep.check("core insert parity", rebc["i"] == legc["i"], "rebuild %d vs legacy %d" % (rebc["i"], legc["i"]))
        rep.check("core amend parity", rebc["a"] == legc["a"], "rebuild %d vs legacy %d" % (rebc["a"], legc["a"]))
        rep.check("core delete parity", rebc["d"] == legc["d"] == 0)
        rep.check("cascade consume parity", rebx["consumed"] == legx["consumed"],
                  "rebuild %d vs legacy %d" % (rebx["consumed"], legx["consumed"]))
        # The cascade restore is a DIVERGENCE, not parity: the legacy DeletePartShipping trigger's
        # `FROM PS, deleted d` cross-join applies ONE arbitrary same-part row per PS row, so deleting
        # two same-part rows in one cascade restores only ONE of them (the F3 multi-row under-count,
        # design §6 class 5). The rebuild's postHeaderDelete restores BOTH (multi-row-safe).
        rep.check("legacy cascade UNDER-restores (F3 multi-row trigger bug): only one row given back",
                  legx["restored"] in (-CA, -CB), "legacy restored delta %d (expected -%d or -%d)"
                  % (legx["restored"], CA, CB))
        rep.check("rebuild cascade restores BOTH rows (F3 fix -> base)", rebx["restored"] == 0,
                  "delta %d" % rebx["restored"])
        rep.check("DIVERGENCE confirmed: rebuild restores MORE than legacy (ledger correct, legacy buggy F3)",
                  rebx["restored"] > legx["restored"], "rebuild %d > legacy %d" % (rebx["restored"], legx["restored"]))
    finally:
        cleanup(base)

    restored = qty()
    ps = int(scalar("SELECT COUNT(*) FROM INV_PART_SHIPPING_INF WHERE VC_PRODUCTION_DATE IN ('%s','%s')" % (DCORE, DCASC)))
    hd = int(scalar("SELECT COUNT(*) FROM INV_SHIPPING_INF WHERE VC_LINE_NAME='%s'" % LINE))
    led = int(scalar("SELECT COUNT(*) FROM INV_STOCK_LEDGER WHERE VC_REASON='%s'" % LREASON))
    trig_off = int(scalar("SELECT COUNT(*) FROM sys.triggers WHERE is_disabled=1 AND name IN "
                          "('InsertPartShipping','UpdatePartShipping','DeletePartShipping','DeleteShipDate','UPDATE_PartNumber')"))
    rep.check("DB restored as found (IN_QTY, no test rows/ledger, triggers re-enabled)",
              restored == base and ps == 0 and hd == 0 and led == 0 and trig_off == 0,
              "IN_QTY=%d ps=%d hd=%d led=%d trig_off=%d" % (restored, ps, hd, led, trig_off))
    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
