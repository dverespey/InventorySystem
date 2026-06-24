#!/usr/bin/env python3
"""test_manifestcost_overlap_guard.py — headless lock on the ManifestCost no-overlap WRITE guard
against the confirmed D6 decision (boundary INCLUSIVE + GAP convention, David 2026-06-18).

The guard already exists: the ManifestCost master Save action (view.json) runs, before every
insert/update, the predicate
    NOT (:end < VC_START_MANIFEST OR :start > VC_END_MANIFEST)   -- params [code, excludeId, end, start]
and rejects when any row matches (a window overlap for the same assy code). The Playwright CRUD test
(test_manifestcost_crud.py) exercises it through the real view — but that needs a live, non-expired
Perspective trial. This test pins the *predicate* deterministically + headlessly, locking BOTH sides of
the inclusive+gap boundary so a future edit can't silently loosen it:
  * touching windows (share a boundary day) -> OVERLAP -> REJECT   (inclusive: end-day belongs to the window)
  * windows separated by >= 1 day            -> no overlap -> ALLOW (gap convention)
  * a window inside/over the anchor          -> REJECT
  * the row's own update (excludeId)         -> never conflicts with itself
Fixture-disciplined (removes only its synthetic rows).

Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 e2e/test_manifestcost_overlap_guard.py
"""
import os, subprocess, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lib import Report  # noqa: E402

CONTAINER = os.environ.get("CONTAINER", "mssql-spike")
SA_PASS = os.environ.get("SA_PASS")
DB = "Inventory"
CODE = "ZZGUARD1"               # synthetic assy code
ANCHOR = ("20250101", "20251231")   # the existing price window [start, end]


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


def overlaps(start, end, excludeId=0):
    """The EXACT guard predicate from the ManifestCost Save action — returns True if the candidate
    [start,end] would be REJECTED (n>0) for assy code CODE."""
    n = int(scalar(
        "SELECT COUNT(*) FROM INV_MANIFEST_COST_MST WHERE VC_ASSY_PART_NUMBER_CODE='%s' "
        "AND IN_MANIFEST_COST_ID <> %d AND NOT ('%s' < VC_START_MANIFEST OR '%s' > VC_END_MANIFEST)"
        % (CODE, excludeId, end, start)) or 0)
    return n > 0


def cleanup():
    sql("DELETE FROM INV_MANIFEST_COST_MST WHERE VC_ASSY_PART_NUMBER_CODE='%s'" % CODE)


def main():
    print("=" * 78)
    print(" MANIFESTCOST no-overlap WRITE guard — inclusive + gap boundary lock (D6, David 2026-06-18)")
    print("=" * 78)
    rep = Report()
    cleanup()
    try:
        anchorId = int(scalar(
            "INSERT INTO INV_MANIFEST_COST_MST (VC_ASSY_PART_NUMBER_CODE, VC_ASSY_MANIFEST_NUMBER, "
            "VC_START_MANIFEST, VC_END_MANIFEST, MO_PRICE) OUTPUT INSERTED.IN_MANIFEST_COST_ID "
            "VALUES ('%s','Z1','%s','%s',100)" % (CODE, ANCHOR[0], ANCHOR[1])))
        s, e = ANCHOR
        print("  anchor window %s-%s (id %d)" % (s, e, anchorId))

        # REJECT cases
        rep.check("REJECT: window INSIDE the anchor (20250601-20250801)", overlaps("20250601", "20250801"))
        rep.check("REJECT: window straddling the anchor (20241201-20260201)", overlaps("20241201", "20260201"))
        rep.check("REJECT: touching at anchor END (cand start == anchor end %s)" % e, overlaps(e, "20260601"))
        rep.check("REJECT: touching at anchor START (cand end == anchor start %s)" % s, overlaps("20240101", s))

        # ALLOW cases — the gap convention's TIGHTEST valid windows (exactly 1 day clear)
        rep.check("ALLOW: 1-day gap BEFORE (cand end 20241231 = anchor start - 1)",
                  not overlaps("20240101", "20241231"))
        rep.check("ALLOW: 1-day gap AFTER (cand start 20260101 = anchor end + 1)",
                  not overlaps("20260101", "20260601"))
        rep.check("ALLOW: far-separated window (20200101-20200601)", not overlaps("20200101", "20200601"))

        # self-exclusion: editing the anchor's OWN row must not conflict with itself
        rep.check("self-edit (excludeId=anchor) never conflicts with itself",
                  not overlaps(s, e, excludeId=anchorId))
        rep.check("...but WITHOUT excludeId the identical window DOES match (proves the predicate fires)",
                  overlaps(s, e, excludeId=0))
    finally:
        cleanup()
    rep.check("fixture clean (synthetic rows removed)",
              int(scalar("SELECT COUNT(*) FROM INV_MANIFEST_COST_MST WHERE VC_ASSY_PART_NUMBER_CODE='%s'" % CODE) or 0) == 0)
    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
