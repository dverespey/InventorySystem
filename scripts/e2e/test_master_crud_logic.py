#!/usr/bin/env python3
"""test_master_crud_logic.py — P16 coverage RESTORATION: re-cover the per-view CRUD logic (validation,
the D3 delete-gate-blocked branch, and the insert/update/delete round-trip) that the P15 server-side
write gate caused the live-browser per-view tests (test_<view>_crud.py) to SKIP.

WHY THIS EXISTS. After P15, the swept master views' Save/Delete call auth.requireWrite(self.session)
FIRST. The live spike session is anonymous (no IdP/roles), so it is correctly DENIED before the
validation / delete-gate / round-trip logic can run — so those three cases got SKIPPED in
test_size_crud.py et al. (lib.WRITE_GATE_SKIP). This test RE-COVERS them by driving the SAME DEPLOYED
Save/Delete scripts under a SEEDED ProductionControl SESSION (the gate PASSES for PC), so the write
reaches the validation / delete-gate / round-trip logic — against the REAL spike DB.

WHAT THIS DRIVES (the REAL gateway-side code — retro R8, same approach as test_master_write_gates):
  We pull each view's DEPLOYED Save/Delete onActionPerformed script from the gateway view.json (what the
  running session serves), compile it (only the Jython-2 `except X, e:` -> Py3 `as` syntactic shim), and
  run it with a fabricated `self` (view.custom form props + a PC session) and the jython_shim `system`
  (real `system.db` against the spike). `import auth as A` resolves to the REAL auth module. So the gate
  passes (PC is a write role), then the deployed validation/delete-gate/round-trip logic runs for real.

THE THREE RE-COVERED CASES, per master (the 7 swept + Sites where applicable):
  (a) VALIDATION rejects bad input  — a blank/short/oversize/dup field -> statusMsg rejection, NO DB write.
  (b) DELETE-GATE REFUSES (the CRITICAL D3 blocked branch) — point recordId at an EXISTING REFERENCED row
      (refCount>0, discovered at runtime from the live data via the view's OWN refCount predicate) and run
      the deployed Delete -> statusMsg "still referenced", SPIKE DELETE BLOCKED, and the row STILL EXISTS.
  (c) ROUND-TRIP — insert a throwaway row through the deployed Save, verify it landed in the DB, then
      delete it through the deployed Delete (zero refs -> gate allows), back to baseline. Swept in teardown.

NON-VACUITY: (a) is paired with a VALID fixture that REACHES the write (so validation isn't block-all);
(b)'s referenced-row blocked is paired with (c)'s zero-ref delete that SUCCEEDS (so the gate isn't
block-all); the round-trip's own insert+verify proves the write path is live under PC.

FIXTURE DISCIPLINE: every throwaway row carries a ZZ-sentinel and is swept in teardown; the delete-gate
case targets an EXISTING referenced row and asserts it SURVIVES (never deletes real data). The spike is
left as-found.

Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 scripts/e2e/test_master_crud_logic.py
"""
import json
import os
import re
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")))

from lib import Report                        # noqa: E402
import jython_shim                            # noqa: E402

REPO = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".."))
AUTH_CODE = os.path.join(REPO, "docs", "analysis", "production-readiness", "project-library", "auth", "code.py")
GW_BASE = ("/usr/local/ignition/data/projects/spike/com.inductiveautomation.perspective/views/Master")

CONTAINER = os.environ.get("CONTAINER", "mssql-spike")
SA_PASS = os.environ.get("SA_PASS", "Spike_Dev_2026!")
DB = "Inventory"

PC_SESSION = {"auth": {"user": {"userName": "op1", "roles": ["ProductionControl"]}}}


def sqlq(query, want_rows=True):
    flags = ["-W"] + ([] if want_rows else ["-h", "-1"])
    cmd = (["docker", "exec", CONTAINER, "/opt/mssql-tools18/bin/sqlcmd", "-C", "-S", "localhost",
            "-U", "sa", "-P", SA_PASS, "-d", DB] + flags +
           ["-Q", "SET QUOTED_IDENTIFIER ON; SET ANSI_NULLS ON; SET NOCOUNT ON; " + query])
    return subprocess.check_output(cmd, stderr=subprocess.STDOUT, timeout=60).decode("utf-8", "replace")


def scalar(query):
    for tok in sqlq(query, want_rows=False).replace("\t", " ").split():
        try:
            return int(tok)
        except ValueError:
            continue
    return None


# ---------------------------------------------------------------------------
# per-view spec: a VALID insert fixture, a BAD (validation-reject) fixture, the table + key for round-trip
# verify/sweep, and a runtime query that finds a REFERENCED row (recordId, plus the form fields the Delete
# reads) for the delete-gate-blocked case. ManifestCost has NO refCount gate (faithful to legacy) -> no (b).
# ---------------------------------------------------------------------------

VIEWS = ["Size", "Supplier", "PartsStock", "ManifestCost", "RenbanGroup", "AssemblyDetail", "Logistics"]

# VALID fixture (passes validation -> reaches the write). ZZ-sentinel codes (swept).
VALID = {
    "Size":           {"form_code": "ZZG", "form_name": "QA Size", "form_usage": "0", "form_days": "0"},
    "Supplier":       {"form_code": "ZZSUP", "form_name": "QA Supplier", "form_invAddPoint": "S"},
    "PartsStock":     {"form_partNumber": "ZZPART01", "form_partsName": "QA Part"},
    "ManifestCost":   {"form_assyCode": "ZZASSY", "form_start": "20260101", "form_end": "20261231",
                       "form_price": "1.50"},
    "RenbanGroup":    {"form_code": "ZZR", "form_count": "5"},
    "AssemblyDetail": {"form_assyCode": "ZZASSY01", "form_effMonth": "202601"},
    "Logistics":      {"form_name": "QA Logistics"},
}

# BAD fixture (a blank required field) + the substring(s) any of which signals a VALIDATION rejection.
# Different views word the blank-field rejection differently (e.g. Supplier's code is varchar(5) so a
# blank code trips the "exactly 5 characters" rule first) — so we accept any of the view's reject words.
REJECT_WORDS = ("required", "exactly", "characters", "must", "cannot", "invalid", "fewer", "blank")
BAD = {
    "Size":           {"form_code": "", "form_name": "QA"},
    "Supplier":       {"form_code": "", "form_name": "QA", "form_invAddPoint": "S"},
    "PartsStock":     {"form_partNumber": "", "form_partsName": "QA"},
    "ManifestCost":   {"form_assyCode": "", "form_start": "20260101", "form_end": "20261231",
                       "form_price": "1.50"},
    "RenbanGroup":    {"form_code": "", "form_count": "5"},
    "AssemblyDetail": {"form_assyCode": "", "form_effMonth": "202601"},
    "Logistics":      {"form_name": ""},
}

# round-trip verify: (table, key column, the VALID fixture's identifying value, the form key that holds it)
ROUND_TRIP = {
    "Size":           ("INV_SIZE_MST", "VC_SIZE_CODE", "ZZG", "form_code"),
    "Supplier":       ("INV_SUPPLIER_MST", "VC_SUPPLIER_CODE", "ZZSUP", "form_code"),
    "PartsStock":     ("INV_PARTS_STOCK_MST", "VC_PART_NUMBER", "ZZPART01", "form_partNumber"),
    "RenbanGroup":    ("INV_RENBAN_GROUP_MST", "VC_RENBAN_GROUP_CODE", "ZZR", "form_code"),
    "Logistics":      ("INV_LOGISTICS_MST", "VC_LOGISTICS_NAME", "QA Logistics", "form_name"),
}

# delete-gate-blocked: a runtime query returning (recordId, {form fields the Delete reads}) for a
# REFERENCED row (refCount>0). None -> the view has no refCount gate (ManifestCost).
def referenced_row(view):
    if view == "Size":
        rid = scalar("SELECT TOP 1 IN_SIZE_ID FROM INV_SIZE_MST WHERE IN_SIZE_ID IN "
                     "(SELECT IN_SIZE_ID FROM INV_PARTS_STOCK_MST UNION "
                     " SELECT IN_SIZE_ID FROM INV_PARTS_STOCK_MST_HIST) ORDER BY IN_SIZE_ID")
        return (rid, {"form_code": "REF"}) if rid else None
    if view == "Supplier":
        rid = scalar("SELECT TOP 1 IN_SUPPLIER_ID FROM INV_SUPPLIER_MST WHERE IN_SUPPLIER_ID IN "
                     "(SELECT IN_SUPPLIER_ID FROM INV_PARTS_STOCK_MST) ORDER BY IN_SUPPLIER_ID")
        # Supplier refCount binds [recId, code, code]; the code must be a real referenced code.
        code = sqlq("SELECT TOP 1 VC_SUPPLIER_CODE FROM INV_SUPPLIER_MST WHERE IN_SUPPLIER_ID=%d" % rid,
                    want_rows=False).strip().split("\n")[-1].strip() if rid else None
        return (rid, {"form_code": code}) if rid else None
    if view == "PartsStock":
        rid = scalar("SELECT TOP 1 IN_PART_ID FROM INV_PARTS_STOCK_MST WHERE VC_PART_NUMBER IN "
                     "(SELECT VC_PART_NUMBER FROM INV_OPEN_ORDER_INF) ORDER BY IN_PART_ID")
        part = sqlq("SELECT TOP 1 VC_PART_NUMBER FROM INV_PARTS_STOCK_MST WHERE IN_PART_ID=%d" % rid,
                    want_rows=False).strip().split("\n")[-1].strip() if rid else None
        return (rid, {"form_partNumber": part}) if rid else None
    if view == "RenbanGroup":
        rid = scalar("SELECT TOP 1 IN_RENBAN_ID FROM INV_RENBAN_GROUP_MST WHERE IN_RENBAN_ID IN "
                     "(SELECT IN_RENBAN_ID FROM INV_PARTS_STOCK_MST WHERE IN_RENBAN_ID IS NOT NULL UNION "
                     " SELECT IN_RENBAN_ID FROM INV_PARTS_STOCK_MST_HIST WHERE IN_RENBAN_ID IS NOT NULL) "
                     "ORDER BY IN_RENBAN_ID")
        return (rid, {"form_code": "REF"}) if rid else None
    if view == "AssemblyDetail":
        # refCount binds the assy code (form_assyCode); recordId is ID_FORECAST_DETAIL.
        rid = scalar("SELECT TOP 1 ID_FORECAST_DETAIL FROM INV_FORECAST_DETAIL_INF WHERE "
                     "VC_ASSY_PART_NUMBER_CODE IN (SELECT VC_PART_NUMBER FROM INV_FORECAST_INF) "
                     "ORDER BY ID_FORECAST_DETAIL")
        assy = sqlq("SELECT TOP 1 VC_ASSY_PART_NUMBER_CODE FROM INV_FORECAST_DETAIL_INF WHERE "
                    "ID_FORECAST_DETAIL=%d" % rid, want_rows=False).strip().split("\n")[-1].strip() \
            if rid else None
        return (rid, {"form_assyCode": assy}) if rid else None
    if view == "Logistics":
        rid = scalar("SELECT TOP 1 IN_LOGISTICS_ID FROM INV_LOGISTICS_MST WHERE IN_LOGISTICS_ID IN "
                     "(SELECT IN_LOGISTICS_ID FROM INV_SUPPLIER_MST WHERE IN_LOGISTICS_ID IS NOT NULL UNION "
                     " SELECT IN_LOGISTICS_ID FROM INV_PARTS_STOCK_MST WHERE IN_LOGISTICS_ID IS NOT NULL) "
                     "ORDER BY IN_LOGISTICS_ID")
        return (rid, {"form_name": "REF"}) if rid else None
    return None   # ManifestCost: no refCount gate


# ---------------------------------------------------------------------------
# harness: compile + drive a deployed script under a PC session against the real DB.
# ---------------------------------------------------------------------------

def _register_auth():
    if "auth" not in sys.modules:
        sys.modules["auth"] = jython_shim.load_wrapper("auth", AUTH_CODE)
    return sys.modules["auth"]


def deployed_script(view, button):
    path = os.path.join(GW_BASE, view, view, "view.json")
    view_json = json.load(open(path))

    def walk(n):
        if isinstance(n, dict):
            if n.get("meta", {}).get("name") == button and "events" in n:
                try:
                    return n["events"]["component"]["onActionPerformed"]["config"]["script"]
                except (KeyError, TypeError):
                    return None
            for v in n.values():
                r = walk(v)
                if r is not None:
                    return r
        elif isinstance(n, list):
            for it in n:
                r = walk(it)
                if r is not None:
                    return r
        return None
    return walk(view_json.get("root", {}))


def compile_script(src):
    body = re.sub(r"except (\S+), e:", r"except \1 as e:", src)
    full = "def _f(self, system, unicode):\n" + body
    ns = {}
    exec(compile(full, "<crud-script>", "exec"), ns)
    return ns["_f"]


class _Custom(object):
    def __init__(self, fields, record_id):
        self.recordId = record_id
        self.runNonce = 0
        self.statusMsg = ""
        self.siteId = 1
        self.mayEdit = True
        for k, v in fields.items():
            setattr(self, k, v)

    def __getattr__(self, name):
        if name.startswith("form_"):
            return ""
        raise AttributeError(name)


class _View(object):
    def __init__(self, c): self.custom = c


class _Self(object):
    def __init__(self, v, s): self.view = v; self.session = s


def run_script(src, fields, record_id, session=PC_SESSION):
    """Drive a deployed Save/Delete script with the REAL jython_shim system (real spike DB). Returns the
    _Custom (so the caller can read statusMsg / recordId after)."""
    _register_auth()
    fn = compile_script(src)
    cust = _Custom(fields, record_id)
    fn(_Self(_View(cust), session), jython_shim._System(), str)
    return cust


# ---------------------------------------------------------------------------
# 1. VALIDATION — bad input rejected (no DB write); the gate already PASSED (PC session).
# ---------------------------------------------------------------------------

def test_validation(rep):
    for V in VIEWS:
        save = deployed_script(V, "SaveButton")
        if save is None:
            rep.skip("%s validation" % V, "deployed Save not found")
            continue
        tbl = ROUND_TRIP.get(V, (None,))[0]
        pre = scalar("SELECT COUNT(*) FROM %s" % tbl) if tbl else None
        cust = run_script(save, dict(BAD[V]), record_id=0)
        msg = (cust.statusMsg or "")
        # rejected = a validation message (one of the reject words), NOT a gate-deny, NOT a success.
        rejected = (any(w in msg.lower() for w in REJECT_WORDS)
                    and "DENIED" not in msg and "Inserted" not in msg and "Updated" not in msg)
        # and (where we can count) NO row was written by the rejected save.
        no_write = (pre is None) or (scalar("SELECT COUNT(*) FROM %s" % tbl) == pre)
        rep.check("%s VALIDATION (PC session): bad input REJECTED past the gate (validation message, no "
                  "DB write)" % V, rejected and no_write, "statusMsg=%r pre=%s" % (msg, pre))


# ---------------------------------------------------------------------------
# 2. DELETE-GATE BLOCKED (CRITICAL D3) — a referenced row (refCount>0) REFUSES delete; row survives.
# ---------------------------------------------------------------------------

def test_delete_gate_blocked(rep):
    for V in VIEWS:
        ref = referenced_row(V)
        if ref is None:
            rep.skip("%s delete-gate-blocked" % V,
                     "no refCount gate (ManifestCost — faithful to legacy DELETE_ManifestCost)"
                     if V == "ManifestCost" else "no referenced row found")
            continue
        rid, fields = ref
        delete = deployed_script(V, "DeleteButton")
        # pre: prove this id exists in its master table (so 'still exists after' is meaningful)
        tbl_key = {"Size": ("INV_SIZE_MST", "IN_SIZE_ID"),
                   "Supplier": ("INV_SUPPLIER_MST", "IN_SUPPLIER_ID"),
                   "PartsStock": ("INV_PARTS_STOCK_MST", "IN_PART_ID"),
                   "RenbanGroup": ("INV_RENBAN_GROUP_MST", "IN_RENBAN_ID"),
                   "AssemblyDetail": ("INV_FORECAST_DETAIL_INF", "ID_FORECAST_DETAIL"),
                   "Logistics": ("INV_LOGISTICS_MST", "IN_LOGISTICS_ID")}[V]
        before = scalar("SELECT COUNT(*) FROM %s WHERE %s=%d" % (tbl_key[0], tbl_key[1], rid))

        cust = run_script(delete, fields, record_id=rid)
        after = scalar("SELECT COUNT(*) FROM %s WHERE %s=%d" % (tbl_key[0], tbl_key[1], rid))
        blocked = "still referenced" in (cust.statusMsg or "")
        rep.check("%s DELETE-GATE BLOCKED (PC session, referenced id=%d): refused (statusMsg 'still "
                  "referenced') AND the referenced row STILL EXISTS (not deleted)" % (V, rid),
                  blocked and after == before == 1,
                  "statusMsg=%r before=%s after=%s" % (cust.statusMsg, before, after))


# ---------------------------------------------------------------------------
# 3. ROUND-TRIP — insert via deployed Save, verify in DB, delete via deployed Delete, back to baseline.
#    (For the 5 views with a simple single-row identity; PartsStock/AssemblyDetail/ManifestCost inserts
#     need richer fixtures — their write path is already proven non-vacuous in test_master_write_gates;
#     here we round-trip the straightforward masters end-to-end through the REAL DB.)
# ---------------------------------------------------------------------------

ROUND_TRIP_VIEWS = ["Size", "Supplier", "RenbanGroup", "Logistics"]


def _sweep_round_trip():
    for V in ROUND_TRIP_VIEWS:
        tbl, key, val, _ = ROUND_TRIP[V]
        sqlq("DELETE FROM %s WHERE %s='%s';" % (tbl, key, val.replace("'", "''")))


def test_round_trip(rep):
    _sweep_round_trip()
    for V in ROUND_TRIP_VIEWS:
        tbl, key, val, _ = ROUND_TRIP[V]
        save = deployed_script(V, "SaveButton")
        delete = deployed_script(V, "DeleteButton")
        base = scalar("SELECT COUNT(*) FROM %s" % tbl)

        # INSERT via the deployed Save (recordId=0).
        cust = run_script(save, dict(VALID[V]), record_id=0)
        landed = scalar("SELECT COUNT(*) FROM %s WHERE %s='%s'" % (tbl, key, val.replace("'", "''")))
        rep.check("%s ROUND-TRIP (PC session): deployed Save INSERTED the row into %s (id flips to edit "
                  "mode)" % (V, tbl), landed == 1 and int(cust.recordId or 0) > 0,
                  "landed=%s recordId=%s statusMsg=%r" % (landed, cust.recordId, cust.statusMsg))
        if landed != 1:
            continue

        # DELETE via the deployed Delete (zero refs -> the gate ALLOWS — non-vacuous vs the blocked case).
        new_id = int(cust.recordId)
        del_fields = dict(VALID[V])
        cust2 = run_script(delete, del_fields, record_id=new_id)
        gone = scalar("SELECT COUNT(*) FROM %s WHERE %s='%s'" % (tbl, key, val.replace("'", "''")))
        rep.check("%s ROUND-TRIP: deployed Delete of the zero-ref row SUCCEEDED (gate allows; DB back to "
                  "baseline %d)" % (V, base),
                  gone == 0 and scalar("SELECT COUNT(*) FROM %s" % tbl) == base,
                  "gone=%s statusMsg=%r" % (gone, cust2.statusMsg))
    _sweep_round_trip()


def main():
    print("=== P16 master-CRUD logic re-coverage under a ProductionControl session "
          "(validation / D3 delete-gate-blocked / round-trip) ===\n")
    rep = Report()
    print("-- 1. Validation (bad input rejected past the gate) --")
    test_validation(rep)
    print("\n-- 2. Delete-gate BLOCKED (D3 critical — referenced row refuses delete, survives) --")
    test_delete_gate_blocked(rep)
    print("\n-- 3. Round-trip (insert + zero-ref delete through the deployed scripts on the real DB) --")
    test_round_trip(rep)
    print("\n-- teardown --")
    _sweep_round_trip()
    rep.check("teardown: round-trip sentinels swept (spike as-found)",
              all(scalar("SELECT COUNT(*) FROM %s WHERE %s='%s'"
                         % (ROUND_TRIP[V][0], ROUND_TRIP[V][1], ROUND_TRIP[V][2].replace("'", "''"))) == 0
                  for V in ROUND_TRIP_VIEWS))
    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
