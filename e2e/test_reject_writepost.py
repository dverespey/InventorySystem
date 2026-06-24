#!/usr/bin/env python3
"""test_reject_writepost.py — row-level PARITY integration test for the reject write path.

Proves, against the LIVE spike DB (fixture-disciplined — leaves the DB exactly as found), that the
REBUILD write-then-post path (source-row write + POST_StockMovement) moves INV_PARTS_STOCK_MST.IN_QTY
IDENTICALLY to the LEGACY reject triggers (INSERT_/UPDATE_/DELETE_RejectParts), for insert/amend/delete.
Twin of test_stocktaking_writepost.py, but reject is stock-OUT (effect = -IN_QTY): an insert REMOVES
qty, a delete (un-reject) RESTORES it.

Two arms (the rebuild path + legacy triggers both write IN_QTY, so they can't co-run):
  legacy  — reject triggers ENABLED; direct INV_REJECT_INF DML fires them (the procs are thin; the
            triggers ARE the legacy stock logic).
  rebuild — reject triggers DISABLED; only POST_StockMovement moves IN_QTY.
UPDATE_PartNumber is disabled for both arms (no INV_PART_QTY_INF audit residue). All restored in finally.

Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 e2e/test_reject_writepost.py
"""
import os, subprocess, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report  # noqa: E402

CONTAINER = os.environ.get("CONTAINER", "mssql-spike")
SA_PASS = os.environ.get("SA_PASS")
DB = "Inventory"
PART = 16
REASON = "WRITEPOST_TEST"
DIV = "T"
Q1, Q2 = 137, 200               # reject qty, then amend-to qty (both REMOVE from on-hand)

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


def cleanup(base):
    sql("DELETE FROM INV_STOCK_LEDGER WHERE VC_SOURCE_EVENT LIKE 'REJECT:rej=%%' AND VC_REASON='%s'; "
        "DELETE FROM INV_REJECT_INF WHERE VC_REASON='%s'; "
        "UPDATE INV_PARTS_STOCK_MST SET IN_QTY=%d WHERE IN_PART_ID=%d; "
        "ENABLE TRIGGER INSERT_RejectParts ON INV_REJECT_INF; "
        "ENABLE TRIGGER UPDATE_RejectParts ON INV_REJECT_INF; "
        "ENABLE TRIGGER DELETE_RejectParts ON INV_REJECT_INF; "
        "ENABLE TRIGGER UPDATE_PartNumber ON INV_PARTS_STOCK_MST"
        % (REASON, REASON, base, PART))


def legacy_arm(rep, base):
    print("\n--- ARM 1: LEGACY (reject triggers ENABLED) ---")
    sql("INSERT INTO INV_REJECT_INF (VC_DIVISION, IN_PART_ID, IN_QTY, VC_REASON, VC_LAST_UPDATE, VC_ADD) "
        "VALUES ('%s', %d, %d, '%s', %s, %s)" % (DIV, PART, Q1, REASON, STAMP, STAMP))
    ins = qty() - base
    rej = scalar("SELECT MAX(IN_REJECT_ID) FROM INV_REJECT_INF WHERE VC_REASON='%s'" % REASON)
    sql("UPDATE INV_REJECT_INF SET IN_QTY=%d, VC_LAST_UPDATE=%s WHERE IN_REJECT_ID=%s" % (Q2, STAMP, rej))
    amend = qty() - base - (-Q1)          # change from the post-insert state
    sql("DELETE FROM INV_REJECT_INF WHERE IN_REJECT_ID=%s" % rej)
    dele = qty() - base
    rep.check("legacy insert delta == -Q1 (reject removes)", ins == -Q1, "%d" % ins)
    rep.check("legacy amend delta == -(Q2-Q1) (reject more)", amend == -(Q2 - Q1), "%d" % amend)
    rep.check("legacy after delete back to base", dele == 0, "delta %d" % dele)
    return {"insert": ins, "amend": amend, "delete_to": qty()}


def rebuild_arm(rep, base):
    print("\n--- ARM 2: REBUILD (reject triggers DISABLED; only POST_StockMovement moves IN_QTY) ---")
    sql("DISABLE TRIGGER INSERT_RejectParts ON INV_REJECT_INF; "
        "DISABLE TRIGGER UPDATE_RejectParts ON INV_REJECT_INF; "
        "DISABLE TRIGGER DELETE_RejectParts ON INV_REJECT_INF")
    # INSERT source row (no trigger) + post -Q1.
    sql("INSERT INTO INV_REJECT_INF (VC_DIVISION, IN_PART_ID, IN_QTY, VC_REASON, VC_LAST_UPDATE, VC_ADD) "
        "VALUES ('%s', %d, %d, '%s', %s, %s)" % (DIV, PART, Q1, REASON, STAMP, STAMP))
    rej = scalar("SELECT MAX(IN_REJECT_ID) FROM INV_REJECT_INF WHERE VC_REASON='%s'" % REASON)
    sql("EXEC POST_StockMovement @partId=%d, @delta=%d, @sourceEnum='REJECT', @sourceRowId=%s, "
        "@eventKey='REJECT:rej=%s:ins', @reason='%s', @site=1, @purge=0" % (PART, -Q1, rej, rej, REASON))
    ins = qty() - base
    led_ins = int(scalar("SELECT IN_QTY_CHANGE FROM INV_STOCK_LEDGER WHERE VC_SOURCE_EVENT='REJECT:rej=%s:ins'" % rej))
    # AMEND Q1->Q2 + post net (old-new = -(Q2-Q1)).
    sql("UPDATE INV_REJECT_INF SET IN_QTY=%d, VC_LAST_UPDATE=%s WHERE IN_REJECT_ID=%s" % (Q2, STAMP, rej))
    sql("EXEC POST_StockMovement @partId=%d, @delta=%d, @sourceEnum='REJECT', @sourceRowId=%s, "
        "@eventKey='REJECT:rej=%s:upd', @reason='%s', @site=1, @purge=0" % (PART, -(Q2 - Q1), rej, rej, REASON))
    amend = qty() - base - (-Q1)
    # DELETE: post reversal +Q2, then remove the source row (no trigger).
    sql("EXEC POST_StockMovement @partId=%d, @delta=%d, @sourceEnum='REJECT', @sourceRowId=%s, "
        "@eventKey='REJECT:rej=%s:del', @reason='%s', @site=1, @purge=0" % (PART, Q2, rej, rej, REASON))
    sql("DELETE FROM INV_REJECT_INF WHERE IN_REJECT_ID=%s" % rej)
    dele = qty() - base
    led_sum = int(scalar("SELECT ISNULL(SUM(IN_QTY_CHANGE),0) FROM INV_STOCK_LEDGER "
                         "WHERE VC_SOURCE_EVENT LIKE 'REJECT:rej=%s:%%'" % rej))
    rep.check("rebuild insert delta == -Q1", ins == -Q1, "%d" % ins)
    rep.check("rebuild ledger row recorded -Q1", led_ins == -Q1, "%d" % led_ins)
    rep.check("rebuild amend delta == -(Q2-Q1)", amend == -(Q2 - Q1), "%d" % amend)
    rep.check("rebuild after delete back to base", dele == 0, "delta %d" % dele)
    rep.check("rebuild ledger nets to 0 over insert+amend+delete", led_sum == 0, "%d" % led_sum)
    return {"insert": ins, "amend": amend, "delete_to": qty()}


def main():
    print("=" * 78)
    print(" REJECT write-then-post — row-level PARITY integration test (rebuild == legacy)")
    print("=" * 78)
    rep = Report()
    base = qty()
    print("  part %d base IN_QTY = %d" % (PART, base))
    sql("DISABLE TRIGGER UPDATE_PartNumber ON INV_PARTS_STOCK_MST")
    try:
        leg = legacy_arm(rep, base)
        sql("UPDATE INV_PARTS_STOCK_MST SET IN_QTY=%d WHERE IN_PART_ID=%d" % (base, PART))
        reb = rebuild_arm(rep, base)
        print("\n--- PARITY: rebuild deltas vs legacy deltas ---")
        rep.check("insert parity (rebuild == legacy)", reb["insert"] == leg["insert"],
                  "rebuild %d vs legacy %d" % (reb["insert"], leg["insert"]))
        rep.check("amend parity (rebuild == legacy)", reb["amend"] == leg["amend"],
                  "rebuild %d vs legacy %d" % (reb["amend"], leg["amend"]))
        rep.check("delete parity (both restore to base)", reb["delete_to"] == leg["delete_to"] == base,
                  "rebuild %d / legacy %d / base %d" % (reb["delete_to"], leg["delete_to"], base))
    finally:
        cleanup(base)

    restored = qty()
    leftover_rows = int(scalar("SELECT COUNT(*) FROM INV_REJECT_INF WHERE VC_REASON='%s'" % REASON))
    leftover_led = int(scalar("SELECT COUNT(*) FROM INV_STOCK_LEDGER WHERE VC_REASON='%s'" % REASON))
    trig_off = int(scalar("SELECT COUNT(*) FROM sys.triggers WHERE is_disabled=1 AND name IN "
                          "('INSERT_RejectParts','UPDATE_RejectParts','DELETE_RejectParts','UPDATE_PartNumber')"))
    rep.check("DB restored as found (IN_QTY, no test rows/ledger, triggers re-enabled)",
              restored == base and leftover_rows == 0 and leftover_led == 0 and trig_off == 0,
              "IN_QTY=%d rows=%d led=%d trig_off=%d" % (restored, leftover_rows, leftover_led, trig_off))
    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
