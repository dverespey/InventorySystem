#!/usr/bin/env python3
"""test_asn_fixes.py — the two SQL-review keystone fixes on create_asn (m1-asn-creation branch).

FIX 1 (BLOCKER) — the DB unique-index backstop behind the non-atomic idempotency guard, and the
                  race-safe create_asn that treats the concurrent-create unique-violation as the
                  idempotent skip. Hot-calls (VC_START_SEQ_NUMBER='-1') must STILL be allowed to repeat.
FIX 2 (SHOULD-FIX) — the No-Ratio pick is now deterministic: create_asn sorts each BC's forecast rows
                  by ID_FORECAST_DETAIL ascending before the fan-out, so the small-volume PEE BC
                  (FEL1000/m36 id=189 vs FEL2000/m37 id=190) deterministically picks the LOWEST id
                  (FEL1000/m36) across runs — instead of luck-of-heap-allocation.

Drives the REAL asn.create_asn through the system.db shim (retro R8 discipline — exercises the actual
gateway-side driver/seam, not a proc-EXEC). All writes go to Inventory and are swept as-found.

Source truth:
  docs/analysis/edi/spike-asn-unique-guard.sql                       (FIX 1 filtered unique index)
  docs/analysis/edi/project-library/asn/code.py                      (race-safe create_asn + FIX2 sort)
  docs/analysis/production-readiness/m1-asn-creation-architecture.md §6 (MUST-FIX / SHOULD-FIX)
  docs/analysis/production-readiness/sql/adversary-findings.md       (BLOCKER-1 / SHOULD-FIX-1)

Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 scripts/e2e/test_asn_fixes.py
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
PDATE = "20260618"          # same window as the parity test (PEE No-Ratio fires here)
LEGACY_ASN = 4721           # source of the legacy window/seq for a real create

INDEX_NAME = "UX_INV_ASN_MST_LINE_PDATE_NORMAL"

ASN_PY = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..",
                          "docs", "analysis", "edi", "project-library", "asn", "code.py"))

# The gateway JDBC session runs QUOTED_IDENTIFIER/ANSI_NULLS ON; required for DML on the filtered-index
# table (else Msg 1934). Mirror it in this test's own sqlcmd helper.
_SET = "SET QUOTED_IDENTIFIER ON; SET ANSI_NULLS ON; SET NOCOUNT ON; "


def sql(db, query, allow_err=False):
    if not SA_PASS:
        sys.exit("export SA_PASS first")
    out = subprocess.check_output([
        "docker", "exec", CONTAINER, "/opt/mssql-tools18/bin/sqlcmd", "-C", "-S", "localhost",
        "-U", "sa", "-P", SA_PASS, "-d", db, "-h", "-1", "-W", "-s", "\t",
        "-Q", _SET + query], text=True)
    if not allow_err:
        out = "\n".join(l for l in out.splitlines() if not l.startswith("Msg "))
    return [l.split("\t") for l in out.splitlines()
            if l.strip() and not l.startswith("(")]


def scalar(db, q):
    rows = sql(db, q)
    return rows[0][0] if rows else None


def sweep(asnIds):
    for a in asnIds:
        if a is not None:
            sql(INV, "DELETE FROM INV_ASN_DETAIL_MST WHERE IN_ASN_ID=%d; "
                     "DELETE FROM INV_ASN_MST WHERE IN_ASN_ID=%d" % (int(a), int(a)))


def normal_count():
    return int(scalar(INV, "SELECT COUNT(*) FROM INV_ASN_MST WHERE VC_LINE_NAME='%s' "
                           "AND VC_PRODUCTION_DATE='%s' AND VC_START_SEQ_NUMBER<>'-1'" % (LINE, PDATE)))


def main():
    print("=" * 80)
    print(" create_asn keystone fixes — FIX 1 (unique-index race backstop) + FIX 2 (deterministic "
          "No-Ratio)")
    print("=" * 80)
    rep = Report()

    asn = jython_shim.load_wrapper("asn_driver_fixes", ASN_PY)

    # ---- preconditions ---------------------------------------------------------------------------
    idx = sql(INV, "SELECT name, has_filter, filter_definition FROM sys.indexes WHERE name='%s' "
                   "AND object_id=OBJECT_ID('dbo.INV_ASN_MST')" % INDEX_NAME)
    rep.check("FIX1: filtered unique index %s exists on INV_ASN_MST" % INDEX_NAME,
              len(idx) == 1 and idx[0][1] == "1", str(idx))
    if not idx:
        print("    (apply docs/analysis/edi/spike-asn-unique-guard.sql first)")
        sys.exit(rep.summary_exit())
    rep.check("FIX1: index filter excludes hot-calls (VC_START_SEQ_NUMBER <> '-1')",
              "-1" in idx[0][2], idx[0][2])

    if normal_count() != 0:
        rep.check("Inventory clean for COROLLA/20260618 before run", False,
                  "%d pre-existing normal ASN(s) — clean up first" % normal_count())
        sys.exit(rep.summary_exit())

    # legacy window/seq for a real create
    hdr = sql(LIVE, "SELECT VC_START_SEQ_NUMBER, CONVERT(varchar,DT_START_SEQ,121), "
                    "VC_END_SEQ_NUMBER, CONVERT(varchar,DT_END_SEQ,121), IN_QTY "
                    "FROM INV_ASN_MST WHERE IN_ASN_ID=%d" % LEGACY_ASN)[0]
    seqStart, dtStart, seqLast, dtEnd, legQty = hdr[0], hdr[1], hdr[2], hdr[3], int(hdr[4])

    created = []
    try:
        # =========================================================================================
        # FIX 1a — the read-back guard no-ops a second create (the fast common path)
        # =========================================================================================
        print("\n--- FIX 1a: second create_asn for same (line,prodDate) -> idempotent skip (guard) ---")
        r1 = asn.create_asn(line=LINE, prodDate=PDATE, seqStart=seqStart, seqLast=seqLast,
                            beginDate=dtStart, endDate=dtEnd, shipQty=legQty,
                            database=DB_CONN, alcDatabase="VehicleOrder")
        created.append(r1["asnId"])
        rep.check("first create_asn made a normal ASN", r1["asnId"] is not None and not r1["skipped"],
                  "asnId=%s" % r1["asnId"])
        r2 = asn.create_asn(line=LINE, prodDate=PDATE, seqStart=seqStart, seqLast=seqLast,
                            beginDate=dtStart, endDate=dtEnd, shipQty=legQty,
                            database=DB_CONN, alcDatabase="VehicleOrder")
        rep.check("second create_asn -> skipped (read-back guard saw the existing ASN)",
                  r2["skipped"] and r2["asnId"] is None, str(r2))
        rep.check("second create_asn wrote NOTHING (still exactly 1 normal ASN)", normal_count() == 1,
                  "%d normal ASN(s)" % normal_count())

        # =========================================================================================
        # FIX 1b — the DB unique-index REJECTS a concurrent second normal insert (the race the
        # read-back guard cannot see). Simulate the loser of the race: a session whose SELECT_ASNSeq
        # guard returned 0 rows (the competitor hadn't committed yet) but whose INSERT lands AFTER the
        # competitor's row exists. We reproduce it by forcing the guard to see 0 rows (monkeypatch the
        # SELECT_ASNSeq read) while a normal ASN for (line,prodDate) is already committed — so the
        # driver's own INSERT_ASNInfo collides with the unique index and create_asn must catch it and
        # return the idempotent skip (writing nothing extra), NOT raise.
        # =========================================================================================
        print("\n--- FIX 1b: concurrent-create RACE -> create_asn catches the unique violation, skips ---")
        real_runPrepQuery = asn.system.db.runPrepQuery

        def guard_blind_runPrepQuery(query, args, db=None):
            # Make ONLY the SELECT_ASNSeq guard return empty (simulate the racing session that read
            # before the competitor committed); every other read (AD_FRSPULL, forecast) is untouched.
            if "SELECT_ASNSeq" in query:
                class _Empty(object):
                    def __len__(self): return 0
                    def __iter__(self): return iter(())
                return _Empty()
            return real_runPrepQuery(query, args, db)

        asn.system.db.runPrepQuery = guard_blind_runPrepQuery
        try:
            r3 = asn.create_asn(line=LINE, prodDate=PDATE, seqStart=seqStart, seqLast=seqLast,
                                beginDate=dtStart, endDate=dtEnd, shipQty=legQty,
                                database=DB_CONN, alcDatabase="VehicleOrder")
            raced_ok = r3["skipped"] and r3["asnId"] is None
            rep.check("create_asn with guard bypassed (race) -> caught unique violation, returned "
                      "skipped (NOT an error)", raced_ok, str(r3))
        except Exception as e:
            rep.check("create_asn with guard bypassed (race) -> caught unique violation, returned "
                      "skipped (NOT an error)", False, "raised instead of skipping: %r" % e)
        finally:
            asn.system.db.runPrepQuery = real_runPrepQuery
        rep.check("the race wrote NOTHING extra (still exactly 1 normal ASN)", normal_count() == 1,
                  "%d normal ASN(s)" % normal_count())

        # =========================================================================================
        # FIX 1c — a raw duplicate normal INSERT is rejected by the index (the backstop directly);
        # and a hot-call (VC_START_SEQ_NUMBER='-1') for the SAME key is STILL allowed.
        # =========================================================================================
        print("\n--- FIX 1c: index rejects raw duplicate normal insert; allows hot-call ---")
        dup = sql(INV,
                  "BEGIN TRY "
                  " INSERT INTO INV_ASN_MST (IN_ASN_EIN,VC_ASN_STATUS,VC_LINE_NAME,VC_ASSEMBLY_LINE,"
                  "VC_START_SEQ_NUMBER,DT_START_SEQ,VC_END_SEQ_NUMBER,DT_END_SEQ,IN_QTY,"
                  "VC_PRODUCTION_DATE,VC_LAST_UPDATE,VC_ADD) "
                  " VALUES(0,'C','%s','','0001',GETDATE(),'0002',GETDATE(),5,'%s',"
                  "'2026062012000000','2026062012000000'); "
                  " SELECT 'INSERTED' "
                  "END TRY BEGIN CATCH SELECT 'REJECTED:' + CAST(ERROR_NUMBER() AS varchar) END CATCH"
                  % (LINE, PDATE), allow_err=True)
        dupres = dup[0][0] if dup else ""
        rep.check("raw duplicate NORMAL insert for (line,prodDate) REJECTED by the index (Msg 2601)",
                  dupres.startswith("REJECTED:2601"), dupres)

        hot = sql(INV,
                  "BEGIN TRY "
                  " INSERT INTO INV_ASN_MST (IN_ASN_EIN,VC_ASN_STATUS,VC_LINE_NAME,VC_ASSEMBLY_LINE,"
                  "VC_START_SEQ_NUMBER,DT_START_SEQ,VC_END_SEQ_NUMBER,DT_END_SEQ,IN_QTY,"
                  "VC_PRODUCTION_DATE,VC_LAST_UPDATE,VC_ADD) "
                  " OUTPUT INSERTED.IN_ASN_ID "
                  " VALUES(0,'C','%s','','-1',GETDATE(),'-1',GETDATE(),1,'%s',"
                  "'2026062012000000','2026062012000000') "
                  "END TRY BEGIN CATCH SELECT 'REJECTED:' + CAST(ERROR_NUMBER() AS varchar) END CATCH"
                  % (LINE, PDATE), allow_err=True)
        hotres = hot[0][0] if hot else ""
        hot_allowed = hotres.isdigit()
        rep.check("hot-call (VC_START_SEQ_NUMBER='-1') for the SAME (line,prodDate) is ALLOWED",
                  hot_allowed, hotres)
        if hot_allowed:
            created.append(int(hotres))   # sweep it too
        # a SECOND hot-call for the same key is also allowed (hot-calls legitimately repeat)
        hot2 = sql(INV,
                   "BEGIN TRY "
                   " INSERT INTO INV_ASN_MST (IN_ASN_EIN,VC_ASN_STATUS,VC_LINE_NAME,VC_ASSEMBLY_LINE,"
                   "VC_START_SEQ_NUMBER,DT_START_SEQ,VC_END_SEQ_NUMBER,DT_END_SEQ,IN_QTY,"
                   "VC_PRODUCTION_DATE,VC_LAST_UPDATE,VC_ADD) "
                   " OUTPUT INSERTED.IN_ASN_ID "
                   " VALUES(0,'C','%s','','-1',GETDATE(),'-1',GETDATE(),1,'%s',"
                   "'2026062012000000','2026062012000000') "
                   "END TRY BEGIN CATCH SELECT 'REJECTED:' + CAST(ERROR_NUMBER() AS varchar) END CATCH"
                   % (LINE, PDATE), allow_err=True)
        hot2res = hot2[0][0] if hot2 else ""
        rep.check("a SECOND hot-call for the same (line,prodDate) is ALSO allowed (repeat is legit)",
                  hot2res.isdigit(), hot2res)
        if hot2res.isdigit():
            created.append(int(hot2res))

        # =========================================================================================
        # FIX 1d — the violation detector is specific: it identifies 2601/the index, NOT unrelated errors.
        # =========================================================================================
        print("\n--- FIX 1d: _isUniqueViolation is specific (no false skips) ---")
        rep.check("_isUniqueViolation True for the ASN index name in the message",
                  asn._isUniqueViolation(Exception(
                      "Cannot insert duplicate key row in object 'dbo.INV_ASN_MST' with unique index "
                      "'%s'. The duplicate key value is (COROLLA, 20260618)." % INDEX_NAME)))
        rep.check("_isUniqueViolation True for a Msg-2601 duplicate-key text",
                  asn._isUniqueViolation(Exception(
                      "Msg 2601, Level 14, State 1, Cannot insert duplicate key row ...")))
        rep.check("_isUniqueViolation False for an unrelated error (deadlock) -> would re-raise",
                  not asn._isUniqueViolation(Exception(
                      "Msg 1205, Transaction was deadlocked on lock resources")))
        rep.check("_isUniqueViolation False for a different FK/constraint error",
                  not asn._isUniqueViolation(Exception(
                      "Msg 547, The INSERT statement conflicted with the FOREIGN KEY constraint")))

        # =========================================================================================
        # FIX 2 — the No-Ratio PEE pick is deterministic = lowest ID_FORECAST_DETAIL (FEL1000/m36).
        # =========================================================================================
        print("\n--- FIX 2: No-Ratio PEE picks lowest ID_FORECAST_DETAIL deterministically ---")
        # PEE matches two forecast rows: id 189 -> 42600FEL1000/m36, id 190 -> 42600FEL2000/m37.
        # Read them through the SAME proc create_asn uses (SELECT_ForecastDetailBCASN), so the test sees
        # exactly the columns the driver consumes (VC_ASSY_MANIFEST_NUMBER comes from the proc's JOIN to
        # INV_MANIFEST_COST_MST — it is NOT a column on INV_FORECAST_DETAIL_INF).
        effMonth = PDATE[0:4] + "/" + PDATE[4:6]
        fcDs = asn.system.db.runPrepQuery(
            "EXEC SELECT_ForecastDetailBCASN @BCode=?, @EffMonth=?", ["PEE", effMonth], INV)
        fcRows = [{
            "ID_FORECAST_DETAIL": int(r["ID_FORECAST_DETAIL"]),
            "VC_ASSY_PART_NUMBER_CODE": r["VC_ASSY_PART_NUMBER_CODE"],
            "VC_ASSY_MANIFEST_NUMBER": r["VC_ASSY_MANIFEST_NUMBER"],
            "IN_ASSY_QTY": r["IN_ASSY_QTY"], "IN_TIRE_RATIO": r["IN_TIRE_RATIO"],
            "IN_WHEEL_RATIO": r["IN_WHEEL_RATIO"], "IN_MANIFEST_COST_ID": r["IN_MANIFEST_COST_ID"],
        } for r in fcDs]
        peeSorted = sorted(fcRows, key=lambda x: x["ID_FORECAST_DETAIL"])
        rep.check("PEE matches >=2 forecast rows (the nondeterminism trigger is present)",
                  len(peeSorted) >= 2,
                  str([(p["ID_FORECAST_DETAIL"], p["VC_ASSY_PART_NUMBER_CODE"]) for p in peeSorted]))
        lowestId = peeSorted[0]["ID_FORECAST_DETAIL"]
        lowestPart = peeSorted[0]["VC_ASSY_PART_NUMBER_CODE"]
        lowestManId = peeSorted[0]["VC_ASSY_MANIFEST_NUMBER"]
        expManifest = asn._manifest(PDATE, lowestManId)   # '7'+1-digit-yr+MM+DD+id
        print("    PEE rows (sorted): %s ; lowest-id pick = id %d -> part %s manifest-id %s -> %s"
              % ([(p["ID_FORECAST_DETAIL"], p["VC_ASSY_PART_NUMBER_CODE"]) for p in peeSorted],
                 lowestId, lowestPart, lowestManId, expManifest))
        frs = [{"BC": "PEE", "Orders": 4, "VEHICLES": 1}]   # No-Ratio (ORDERS<=5), the live PEE shape

        # (i) The sort is LOAD-BEARING: pass the rows DESCENDING (the adversarial heap order) straight to
        #     the pure computeAsnDetails with NO sort -> it would pick the WRONG (highest-id) row. This
        #     reproduces the pre-fix nondeterminism and proves the pick depends on input order.
        desc = sorted(fcRows, key=lambda r: -r["ID_FORECAST_DETAIL"])
        out_unsorted = asn.computeAsnDetails(frs, {"PEE": desc}, PDATE, 1)
        rep.check("UNSORTED (descending) input -> pure fan-out picks the HIGHEST-id row (the pre-fix "
                  "nondeterminism the sort removes)",
                  out_unsorted[0]["partNumber"] == peeSorted[-1]["VC_ASSY_PART_NUMBER_CODE"],
                  "picked %s (highest-id=%s)" % (out_unsorted[0]["partNumber"],
                                                 peeSorted[-1]["VC_ASSY_PART_NUMBER_CODE"]))

        # (ii) The DRIVER's sort (ID_FORECAST_DETAIL ascending) -> the pick is the lowest-id row,
        #      INDEPENDENT of the heap order the proc returned. Run computeAsnDetails over the
        #      driver-sorted rows from BOTH an ascending and a descending starting order -> identical.
        out_asc = asn.computeAsnDetails(
            frs, {"PEE": sorted(fcRows, key=lambda r: r["ID_FORECAST_DETAIL"])}, PDATE, 1)
        out_from_desc = asn.computeAsnDetails(
            frs, {"PEE": sorted(desc, key=lambda r: r["ID_FORECAST_DETAIL"])}, PDATE, 1)
        rep.check("No-Ratio PEE emits exactly ONE row (Orders<=5, break)", len(out_asc) == 1,
                  str(out_asc))
        rep.check("DRIVER-sorted -> No-Ratio PEE picks the LOWEST-id part (%s) deterministically"
                  % lowestPart, out_asc[0]["partNumber"] == lowestPart, out_asc[0]["partNumber"])
        rep.check("DRIVER-sorted -> No-Ratio PEE manifest = lowest-id manifest (%s)" % expManifest,
                  out_asc[0]["manifest"] == expManifest, out_asc[0]["manifest"])
        rep.check("pick is STABLE regardless of the heap's starting order (driver sort makes asc==desc)",
                  out_asc[0]["partNumber"] == out_from_desc[0]["partNumber"]
                  and out_asc[0]["manifest"] == out_from_desc[0]["manifest"],
                  "asc=%s fromDesc=%s" % (out_asc[0]["partNumber"], out_from_desc[0]["partNumber"]))

        # End-to-end through the persisted ASN: the PEE-sourced manifest (m36) must be present in the
        # created ASN's details and carry FEL1000 (proves the live driver applied the sort on write).
        peeManifest = asn._manifest(PDATE, "36")
        det = sql(INV, "SELECT VC_ASSY_PART_NUMBER, IN_QTY FROM INV_ASN_DETAIL_MST "
                       "WHERE IN_ASN_ID=%d AND VC_MANIFEST_NUMBER='%s'" % (int(r1["asnId"]), peeManifest))
        rep.check("persisted ASN has the PEE-target manifest %s with FEL1000 (driver sort applied on "
                  "the real write)" % peeManifest,
                  len(det) == 1 and det[0][0] == "42600FEL1000", str(det))

    finally:
        sweep(created)
        rep.check("Inventory restored as-found (no normal ASN left for COROLLA/20260618)",
                  normal_count() == 0, "%d remaining" % normal_count())

    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
