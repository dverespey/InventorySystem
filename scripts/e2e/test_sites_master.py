#!/usr/bin/env python3
"""test_sites_master.py — M4 piece 1 DB/save-logic gate for the Sites master + the path columns.

NOT a browser test (that is test_sites_crud.py, which drives the live Perspective chrome). This proves the
DATA + VALIDATION + ROUND-TRIP layer of the Sites master headless, via the spike DB:

  1. PATH-COLUMNS DDL (spike-inv-sites-paths.sql) applied + IDEMPOTENT (re-run safe; 7 cols present).
  2. INV_SITES CRUD ROUND-TRIP: run the EXACT insert/update/get SQL the view runs (imported from the
     shared builders in scripts/gen_sites_view.py — single source of truth, no drift) against the DB in a
     ROLLED-BACK transaction; edit every field TYPE incl. the new path columns; re-read; assert the save
     wrote what was set.
  3. VALIDATION: a 2-char separator / 2-char EDI-mode REJECTED; a bad DUNS REJECTED; fill-days>50 and
     retention 1..11 REJECTED. The oracle is INDEPENDENT of the rebuild:
       * fill-days / retention — the TABLE CHECK constraints (CK_INV_SITES_FILL_DAYS / _DATA_RETENTION),
         in the table artifact, NOT the view: we attempt the insert and assert the DB itself rejects it.
       * 1-char separators / EDI-mode / DUNS — the source-truth rule (m4-auth-sites-sourcetruth.md §4:
         "VC_EDI_MODE + separators EXACTLY 1 char — load-bearing positional ISA"; DUNS 9 or 9+4). We
         transcribe that rule as an INDEPENDENT predicate here and assert it flags the bad values. We also
         show WHY the save-path guard is load-bearing for VC_EDI_MODE: it is varchar(10), so the DB
         SILENTLY STORES a 2-char value (NO error) -> a malformed positional ISA15 ships unless the save
         path catches it. (The char(1) separators happen to get a DB backstop, Msg 8152, but it is opaque.)

  4. SAVE-PATH GUARD (the load-bearing enforcement — FIX 2 from the dual-adversary review): drive bad
     values through the DEPLOYED view's ACTUAL save_script() validation, not just the independent oracle.
     We execute the EXACT save_script string the gateway view contains (gsv.save_script(), verified
     byte-identical to the deployed sites-save-btn onActionPerformed), with a stub system.db that RECORDS
     whether any write was attempted. Assert: a 2-char VC_EDI_MODE, a 2-char separator, and a bad DUNS are
     each REJECTED BY THE SAVE PATH ITSELF (early return, statusMsg set, NO db write) — the guard that
     prevents a malformed positional ISA15 reaching TEMA. NON-VACUITY: a fully-valid value PASSES the guard
     (reaches the db write); and reverting the guard (we patch the guard line out of the save source) lets
     the bad value SLIP THROUGH to the db write — proving the guard, not the harness, is what blocks it.

HONEST: the screen's exact Perspective chrome (the visual form, the buttons, the admin gate) is
Designer-finish + is exercised separately by test_sites_crud.py (browser). THIS gate proves the
data/validation/round-trip. The role-gate enforcement (Admin + production-control user) lands in M4 piece 2.

NON-VACUITY: each round-trip assertion reads back a value DIFFERENT from the seed, so a no-op save would
fail; each validation assertion would FAIL if the rule were dropped (a 2-char value would round-trip).

Spike-as-found: the round-trip + validation rows live INSIDE a transaction that is ALWAYS rolled back, so
INV_SITES ends exactly as found (2 seed rows). The path columns are an intended PERMANENT addition (like
the M2 forecast-config columns) — left in place, no test rows.

Run:  export SA_PASS='Spike_Dev_2026!'  &&  python3 scripts/e2e/test_sites_master.py
"""
import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")))

from lib import Report                      # noqa: E402
import gen_sites_view as gsv                # noqa: E402  (the SHARED SQL builders the view uses)

CONTAINER = os.environ.get("CONTAINER", "mssql-spike")
SA_PASS = os.environ.get("SA_PASS")
DB = "Inventory"
REPO = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".."))
PATHS_DDL = os.path.join(REPO, "docs", "analysis", "master-data", "spike-inv-sites-paths.sql")

# The 7 path columns this M4 piece adds (legacy INI [DIRECTORIES] set).
PATH_COLS = ["VC_EDIOUT_DIR", "VC_EDIIN_DIR", "VC_FORECAST_DIR", "VC_LOGISTICS_DIR",
             "VC_REPORTS_DIR", "VC_SHIPPING_DIR", "VC_TEMPLATE_DIR"]

# The 16-char yyyymmddHHMMSSff audit recipe (same as the view + the table artifact).
AUDIT_SQL = ("CONVERT(char(8),GETDATE(),112) "
             "+ SUBSTRING(CONVERT(varchar,GETDATE(),114),1,2) + SUBSTRING(CONVERT(varchar,GETDATE(),114),4,2) "
             "+ SUBSTRING(CONVERT(varchar,GETDATE(),114),7,2) + SUBSTRING(CONVERT(varchar,GETDATE(),114),10,2)")


def _sqlcmd(sql_text, db=DB):
    """One batch via sqlcmd. R23 gate: QUOTED_IDENTIFIER ON (filtered-index/DML safety + parity with the
    gateway JDBC session). Returns (returncode, stdout)."""
    if not SA_PASS:
        sys.exit("export SA_PASS first")
    p = subprocess.run(
        ["docker", "exec", CONTAINER, "/opt/mssql-tools18/bin/sqlcmd", "-C", "-S", "localhost",
         "-U", "sa", "-P", SA_PASS, "-d", db, "-W", "-s", "\t", "-h", "-1",
         "-Q", "SET QUOTED_IDENTIFIER ON; SET ANSI_NULLS ON; SET NOCOUNT ON; " + sql_text],
        capture_output=True, text=True)
    return p.returncode, (p.stdout or "") + (p.stderr or "")


def sql_rows(sql_text, db=DB):
    rc, out = _sqlcmd(sql_text, db)
    return [l.split("\t") for l in out.splitlines()
            if l.strip() and not l.startswith("(") and not l.startswith("Msg ")
            and not set(l) <= set("- \t")]


def scalar(sql_text, db=DB):
    r = sql_rows(sql_text, db)
    return r[0][0] if r else None


def lit(v):
    """T-SQL literal for a bound value (the harness inlines; the view uses real ? binding)."""
    if v is None:
        return "NULL"
    if isinstance(v, bool):
        return "1" if v else "0"
    if isinstance(v, int):
        return str(v)
    return "N'" + str(v).replace("'", "''") + "'"


# ---------------------------------------------------------------------------
# Build the INSERT/UPDATE the VIEW runs, from the SHARED builders (gsv.*), with
# the test's value map. This is the IDENTICAL column list/order the view emits;
# only the placeholder binding differs (inline literals vs ?).
# ---------------------------------------------------------------------------

def _arg_value(form_key, kind, values):
    """Mirror the view save-script coercions for the harness (the view uses _str/_int/_intn/_bit/_mode)."""
    raw = values[form_key]
    if kind == "bit":
        return 1 if raw in (1, True, "1", "true", "True") else 0
    if kind == "num":
        if raw in (None, ""):
            return None
        n = int(raw)
        if form_key in gsv.NULLABLE_UNSET:            # blank/0 -> NULL (B1/N1 rule)
            return n if n != 0 else None
        return n
    if kind == "mode":
        m = (str(raw) if raw not in (None, "") else "AUTO").upper()
        return m if m in ("AUTO", "MANUAL") else "AUTO"
    return str(raw) if raw is not None else ""


def insert_sql(values):
    # gsv.INS_VALS is the placeholder run for the 34 ? cols PLUS a trailing ", " (IN_EIN_SEQ is the
    # literal 0 baked in). The generator appends the AUDIT expr (VC_ADD) AFTER it; we do the same here.
    args = iter([_arg_value(k, kd, values) for (k, kd) in gsv.INS_ORDER])
    out = []
    for ch in gsv.INS_VALS:
        if ch == "?":
            out.append(lit(next(args)))
        else:
            out.append(ch)
    body = "".join(out) + AUDIT_SQL          # VC_ADD = the 16-char audit stamp (last column)
    return ("INSERT INTO INV_SITES (" + gsv.INS_COLS + ") VALUES (" + body + ");"
            " SELECT CAST(SCOPE_IDENTITY() AS int) AS newId")


def update_sql(values, rec_id):
    setclause = gsv.UPD_SET + "VC_LAST_UPDATE = " + AUDIT_SQL
    args = iter([_arg_value(k, kd, values) for (k, kd) in gsv.INS_ORDER])
    out = []
    for ch in setclause:
        if ch == "?":
            out.append(lit(next(args)))
        else:
            out.append(ch)
    return "UPDATE INV_SITES SET " + "".join(out) + " WHERE IN_SITE_ID = " + str(rec_id)


def get_sql(rec_id):
    return "SELECT " + gsv.GET_COLS + " FROM INV_SITES WHERE IN_SITE_ID = " + str(rec_id)


# ---------------------------------------------------------------------------
# INDEPENDENT validation oracle — transcribed from m4-auth-sites-sourcetruth.md §4
# (NOT from the view save-script). A field is VALID iff it passes this predicate.
# ---------------------------------------------------------------------------

def source_truth_reject_reason(values):
    """Return the FIRST rejection reason per the source-truth rules, or None if all pass. This is the
    independent oracle the test asserts the rebuild's save-path must match."""
    def s(k):
        v = values.get(k)
        return str(v).strip() if v is not None else ""
    # load-bearing positional ISA: EDI mode + 3 separators EXACTLY 1 char if set.
    if s("form_edimode") != "" and len(s("form_edimode")) != 1:
        return "EDI mode must be exactly 1 char (ISA15)"
    for k, lbl in (("form_sepseg", "Segment"), ("form_sepelem", "Element"), ("form_sepsub", "Sub-element")):
        if s(k) != "" and len(s(k)) != 1:
            return "Separator %s must be exactly 1 char" % lbl
    # DUNS 9 or 9+4 digits if set.
    for k, lbl in (("form_duns", "DUNS"), ("form_tmmduns", "TMM DUNS")):
        d = s(k)
        if d != "" and not (d.isdigit() and len(d) in (9, 13)):
            return "%s must be 9 or 13 digits" % lbl
    # CHECK-backed: fill-days <= 50; retention >= 12 (typed 1..11 invalid).
    fd = values.get("form_filldays")
    if fd not in (None, "") and int(fd) > 50:
        return "Fill Days must be <= 50"
    rt = values.get("form_retention")
    if rt not in (None, "") and 0 < int(rt) < 12:
        return "Data Retention must be >= 12"
    return None


# A complete, ALL-VALID value map (every field type incl. the 7 paths). Distinct from the seed so a no-op
# save would not satisfy the round-trip assertions (non-vacuity).
def valid_values(abbr="ZZTST"):
    return {
        "form_name": "Roundtrip Test Site", "form_abbr": abbr, "form_street": "1 Test Way",
        "form_city": "Testville", "form_state": "MS", "form_country": "US", "form_zip": "38801",
        "form_duns": "123456789", "form_supcode": "ZZ", "form_dock": "D9", "form_edimode": "T",
        "form_sepseg": "~", "form_sepelem": "*", "form_sepsub": ">",
        "form_tmmname": "Test TMM", "form_tmmabbr": "TTMM", "form_tmmduns": "987654321",
        "form_maxseq": 888, "form_acceptasn": 1, "form_delivery": "CV",
        "form_filldays": 45, "form_fcusage": 7, "form_usefpd": 1,
        "form_fcmode": "MANUAL", "form_enpurge": 1, "form_prpurge": 0, "form_retention": 24,
        # the 7 path columns (placeholder example values — NEVER real client paths):
        "form_ediout": "/opt/test/edi/out", "form_ediin": "/opt/test/edi/in",
        "form_forecastd": "/opt/test/forecast", "form_logisticsd": "/opt/test/logistics",
        "form_reportsd": "/opt/test/reports", "form_shippingd": "/opt/test/shipping",
        "form_templated": "/opt/test/templates",
    }


# expected DB values after a save (mirror the view coercions: bit->0/1, nullable-unset INT, mode upper).
def expected_db(values):
    exp = {}
    for k, kd in gsv.INS_ORDER:
        exp[k] = _arg_value(k, kd, values)
    return exp


# form_key -> DB column (from the generator's field model F).
FORM_TO_COL = {key: dbcol for (key, _, _, _, dbcol, _, _, _) in gsv.F}


# ---------------------------------------------------------------------------
# SAVE-PATH HARNESS (FIX 2). Drives the DEPLOYED view's ACTUAL save_script through
# its validation, with a stub system.db that records whether a write was attempted.
#
# We execute the EXACT save_script() string the generator emits (verified byte-identical
# to the deployed sites-save-btn onActionPerformed in main()), NOT a reimplementation.
# The body is a Jython-2 method fragment (one leading tab/line); to run it under CPython3
# we apply ONLY syntactic shims that do NOT touch validation semantics:
#   * `except Exception, e:` -> `except Exception as e:`  (Py3 except syntax)
#   * inject unicode=str into the namespace (u"" literals are already valid Py3 str)
# len()/.strip()/.isdigit()/.upper() behave identically, so the guards run unchanged.
# ---------------------------------------------------------------------------

class _Custom(object):
    """Attribute-style stand-in for self.view.custom; seeded from a value map."""
    def __init__(self, values, is_admin=True, record_id=0):
        self.recordId = record_id
        self.runNonce = 0
        self.searchTerm = ""
        self.statusMsg = ""
        self.isAdmin = is_admin
        for k in FORM_TO_COL:
            setattr(self, k, values.get(k, ""))


class _View(object):
    def __init__(self, custom):
        self.custom = custom


class _Self(object):
    def __init__(self, view):
        self.view = view


class _StubDB(object):
    """Records whether the save path reached a DB write (= passed all validation)."""
    def __init__(self):
        self.write_attempted = False

    def _ds(self):
        class _DS(object):
            rowCount = 1
            def getValueAt(self, *a):
                return 1
        return _DS()

    def runPrepQuery(self, *a, **k):
        self.write_attempted = True       # the INSERT-with-SCOPE_IDENTITY path
        return self._ds()

    def runPrepUpdate(self, *a, **k):
        self.write_attempted = True       # the UPDATE path
        return 1


class _StubLogger(object):
    def info(self, *a, **k):
        pass
    def warn(self, *a, **k):
        pass


class _StubUtil(object):
    def getLogger(self, *a, **k):
        return _StubLogger()


class _StubSystem(object):
    def __init__(self, db):
        self.db = db
        self.util = _StubUtil()


def _compile_save(save_src):
    """Wrap the save_script body into a callable under CPython3 (syntactic shims only)."""
    body = save_src.replace("except Exception, e:", "except Exception as e:")
    src = "def _save(self, system, unicode):\n" + body
    ns = {}
    exec(compile(src, "<sites-save_script>", "exec"), ns)
    return ns["_save"]


def run_save_path(values, save_src=None, record_id=0):
    """Run the deployed save_script over `values`. Returns (write_attempted, statusMsg).
    write_attempted=False -> the save path REJECTED the input (early return, no db write)."""
    if save_src is None:
        save_src = gsv.save_script()
    fn = _compile_save(save_src)
    db = _StubDB()
    cust = _Custom(values, is_admin=True, record_id=record_id)
    fn(_Self(_View(cust)), _StubSystem(db), str)
    return db.write_attempted, cust.statusMsg


def _deployed_save_script():
    """Pull the sites-save-btn onActionPerformed script from the DEPLOYED gateway view.json, so we can
    assert the harness drives the SAME source the running session serves (not just the generator)."""
    import json
    gw = os.path.join("/usr/local/ignition/data/projects/spike/com.inductiveautomation.perspective"
                      "/views/Master/Sites/Sites/view.json")
    if not os.path.exists(gw):
        return None
    with open(gw) as f:
        view = json.load(f)

    def walk(node):
        if isinstance(node, dict):
            if node.get("meta", {}).get("domId") == "sites-save-btn":
                try:
                    return node["events"]["component"]["onActionPerformed"]["config"]["script"]
                except (KeyError, TypeError):
                    return None
            for v in node.values():
                r = walk(v)
                if r is not None:
                    return r
        elif isinstance(node, list):
            for it in node:
                r = walk(it)
                if r is not None:
                    return r
        return None

    return walk(view)


def main():
    print("=== Sites master — DDL + CRUD round-trip + validation (M4 piece 1) ===")
    rep = Report()

    # ---- 1. PATH-COLUMNS DDL applied + idempotent ----
    docker_cp = subprocess.run(["docker", "cp", PATHS_DDL, "%s:/tmp/paths_test.sql" % CONTAINER],
                               capture_output=True, text=True)
    run1 = subprocess.run(["docker", "exec", CONTAINER, "/opt/mssql-tools18/bin/sqlcmd", "-C", "-S",
                           "localhost", "-U", "sa", "-P", SA_PASS, "-i", "/tmp/paths_test.sql"],
                          capture_output=True, text=True)
    run2 = subprocess.run(["docker", "exec", CONTAINER, "/opt/mssql-tools18/bin/sqlcmd", "-C", "-S",
                           "localhost", "-U", "sa", "-P", SA_PASS, "-i", "/tmp/paths_test.sql"],
                          capture_output=True, text=True)
    idem_ok = (docker_cp.returncode == 0 and run1.returncode == 0 and run2.returncode == 0
               and "applied" in (run1.stdout or "") and "applied" in (run2.stdout or ""))
    rep.check("path-columns DDL applies + re-runs safe (idempotent)", idem_ok,
              "rc1=%s rc2=%s" % (run1.returncode, run2.returncode))

    present = sql_rows("SELECT name FROM sys.columns WHERE object_id=OBJECT_ID('dbo.INV_SITES') "
                       "AND name IN (%s) ORDER BY name" % ", ".join("'%s'" % c for c in PATH_COLS))
    got = sorted(r[0] for r in present)
    rep.check("all 7 path columns present on INV_SITES", got == sorted(PATH_COLS),
              "got=%s" % got)

    # ---- snapshot the seed for restore-as-found verification ----
    seed_count = int(scalar("SELECT COUNT(*) FROM INV_SITES"))

    # ---- 2. CRUD round-trip (rolled-back transaction) ----
    # We open ONE batch that BEGIN TRANs, inserts via the REAL insert SQL, reads back, updates via the
    # REAL update SQL, reads back, and SELECTs both snapshots — then ROLLBACK. Parse the two snapshots.
    vals_ins = valid_values("ZZRT1")
    vals_upd = dict(vals_ins)
    vals_upd.update({"form_name": "Updated Site Name", "form_city": "Updatedville",
                     "form_edimode": "P", "form_fcmode": "AUTO", "form_retention": 36,
                     "form_filldays": 12, "form_acceptasn": 0, "form_prpurge": 1,
                     "form_ediout": "/opt/test/edi/out2", "form_templated": "/opt/test/templates2",
                     "form_duns": "111222333"})

    # Build a single scripted batch: insert -> capture id -> read row1 -> update -> read row2 -> rollback.
    ins = insert_sql(vals_ins)
    cols = gsv.GET_COLS
    batch = (
        "BEGIN TRAN;\n"
        "DECLARE @id int;\n"
        + ins.replace("SELECT CAST(SCOPE_IDENTITY() AS int) AS newId",
                      "SET @id = CAST(SCOPE_IDENTITY() AS int)") + ";\n"
        "SELECT 'AFTER_INSERT' AS phase, " + cols + " FROM INV_SITES WHERE IN_SITE_ID=@id;\n"
        + update_sql(vals_upd, "@id") + ";\n"
        "SELECT 'AFTER_UPDATE' AS phase, " + cols + " FROM INV_SITES WHERE IN_SITE_ID=@id;\n"
        "ROLLBACK TRAN;\n"
    )
    rc, out = _sqlcmd(batch)
    # parse the two phase rows (header suppressed by -h -1; we tagged each row with a phase literal)
    rows = [l.split("\t") for l in out.splitlines() if l.startswith("AFTER_")]
    after_ins = next((r for r in rows if r[0] == "AFTER_INSERT"), None)
    after_upd = next((r for r in rows if r[0] == "AFTER_UPDATE"), None)
    rep.check("round-trip batch ran (insert + update inside one tx)", rc == 0 and after_ins and after_upd,
              "rc=%s rows=%d" % (rc, len(rows)))

    if after_ins and after_upd:
        # map: col name (GET_COLS order) -> value cell. after_*[0] is the phase tag, [1:] are the columns.
        colnames = cols.split(", ")
        ins_map = dict(zip(colnames, after_ins[1:]))
        upd_map = dict(zip(colnames, after_upd[1:]))

        def db_cell(m, col):
            v = m.get(col)
            return None if v in (None, "NULL") else v

        # assert the INSERT wrote every field type incl. the 7 paths
        exp = expected_db(vals_ins)
        mism = []
        for fk, ev in exp.items():
            col = FORM_TO_COL[fk]
            got = db_cell(ins_map, col)
            want = None if ev is None else str(ev)
            if got != want:
                mism.append("%s: got=%r want=%r" % (col, got, want))
        rep.check("INSERT round-trip: every field (incl. 7 paths) stored as set", not mism,
                  "; ".join(mism[:6]))
        # spot the path columns explicitly (the M4 deliverable)
        path_ok = all(db_cell(ins_map, c) == vals_ins[fk]
                      for fk, c in FORM_TO_COL.items() if c in PATH_COLS)
        rep.check("INSERT: the 7 path columns round-trip exactly", path_ok,
                  ", ".join("%s=%s" % (c, db_cell(ins_map, c)) for c in PATH_COLS))

        # assert the UPDATE changed the edited fields
        exp_u = expected_db(vals_upd)
        mism_u = []
        for fk, ev in exp_u.items():
            col = FORM_TO_COL[fk]
            got = db_cell(upd_map, col)
            want = None if ev is None else str(ev)
            if got != want:
                mism_u.append("%s: got=%r want=%r" % (col, got, want))
        rep.check("UPDATE round-trip: edited fields (name/city/edimode/paths/duns/retention) stored", not mism_u,
                  "; ".join(mism_u[:6]))
        # non-vacuity: the update must have CHANGED at least the values we changed
        changed = (db_cell(upd_map, "VC_SITE_NAME") != db_cell(ins_map, "VC_SITE_NAME")
                   and db_cell(upd_map, "VC_EDIOUT_DIR") != db_cell(ins_map, "VC_EDIOUT_DIR")
                   and db_cell(upd_map, "VC_EDI_MODE") != db_cell(ins_map, "VC_EDI_MODE"))
        rep.check("UPDATE non-vacuity: changed values actually differ from the insert snapshot", changed,
                  "name %r->%r, edimode %r->%r" % (db_cell(ins_map, "VC_SITE_NAME"),
                  db_cell(upd_map, "VC_SITE_NAME"), db_cell(ins_map, "VC_EDI_MODE"),
                  db_cell(upd_map, "VC_EDI_MODE")))
        # IN_EIN_SEQ must be the literal-0 default on insert (rule #3 — never written by the form)
        rep.check("rule #3: IN_EIN_SEQ defaulted to 0 on insert (not form-written)",
                  db_cell(ins_map, "IN_EIN_SEQ") == "0", db_cell(ins_map, "IN_EIN_SEQ"))

    # ---- 3a. VALIDATION — CHECK-backed (DB is the independent oracle) ----
    # fill-days > 50: attempt the REAL insert; assert the DB CHECK rejects it.
    bad_fd = valid_values("ZZBFD"); bad_fd["form_filldays"] = 60
    rc_fd, out_fd = _sqlcmd("BEGIN TRAN; " + insert_sql(bad_fd) + "; ROLLBACK TRAN;")
    fd_rejected = "CK_INV_SITES_FILL_DAYS" in out_fd or ("conflict" in out_fd.lower() and rc_fd != 0)
    rep.check("validation: fill-days=60 (>50) REJECTED by the table CHECK", fd_rejected,
              "rc=%s last=%s" % (rc_fd, (out_fd.strip().splitlines() or [""])[-1]))
    # source-truth oracle agrees it should be rejected
    rep.check("  (oracle) source-truth flags fill-days=60", source_truth_reject_reason(bad_fd) is not None,
              source_truth_reject_reason(bad_fd))

    # retention typed 1..11: DB CHECK rejects. (Bind the typed value directly — NOT the nullable-unset 0.)
    bad_rt = valid_values("ZZBRT"); bad_rt["form_retention"] = 6
    rc_rt, out_rt = _sqlcmd("BEGIN TRAN; " + insert_sql(bad_rt) + "; ROLLBACK TRAN;")
    rt_rejected = "CK_INV_SITES_DATA_RETENTION" in out_rt or ("conflict" in out_rt.lower() and rc_rt != 0)
    rep.check("validation: retention=6 (1..11) REJECTED by the table CHECK", rt_rejected,
              "rc=%s" % rc_rt)
    rep.check("  (oracle) source-truth flags retention=6", source_truth_reject_reason(bad_rt) is not None,
              source_truth_reject_reason(bad_rt))

    # ---- 3b. VALIDATION — save-path-only (the source-truth oracle + the DB-doesn't-catch-it proof) ----
    # 2-char separator: source-truth REJECTS. The char(1) DB column happens to BACKSTOP it (Msg 8152
    # "String or binary data would be truncated"), but that error is opaque and the EDI-mode case (below)
    # has NO DB backstop — so the save-path guard is the real, user-facing enforcement.
    bad_sep = valid_values("ZZBSP"); bad_sep["form_sepelem"] = "**"
    rep.check("validation: 2-char separator REJECTED by the source-truth oracle",
              source_truth_reject_reason(bad_sep) == "Separator Element must be exactly 1 char",
              source_truth_reject_reason(bad_sep))
    rc_sp, out_sp = _sqlcmd("BEGIN TRAN; " + insert_sql(bad_sep) + "; ROLLBACK TRAN;")
    rep.check("  (DB backstop) char(1) separator: DB errors on a 2-char value (Msg 8152)",
              "8152" in out_sp or "truncat" in out_sp.lower(),
              (out_sp.strip().splitlines() or [""])[0])

    # 2-char EDI mode: source-truth REJECTS. THE LOAD-BEARING case — VC_EDI_MODE is varchar(10), so the
    # DB SILENTLY STORES a 2-char value with NO error -> a malformed positional ISA15 ships unless the
    # save-path guard catches it. This proves the guard is load-bearing (the DB will NOT save us).
    bad_em = valid_values("ZZBEM"); bad_em["form_edimode"] = "PP"
    rep.check("validation: 2-char EDI mode REJECTED by the source-truth oracle",
              source_truth_reject_reason(bad_em) == "EDI mode must be exactly 1 char (ISA15)",
              source_truth_reject_reason(bad_em))
    rc_em, out_em = _sqlcmd("BEGIN TRAN; " + insert_sql(bad_em)
                            + "; SELECT 'EMLEN' AS p, LEN(VC_EDI_MODE) AS n FROM INV_SITES "
                            "WHERE IN_SITE_ID = SCOPE_IDENTITY(); ROLLBACK TRAN;")
    emrow = next((l.split("\t") for l in out_em.splitlines() if l.startswith("EMLEN")), None)
    db_stored_2char = (rc_em == 0 and emrow is not None and emrow[1] == "2" and "8152" not in out_em)
    rep.check("validation: varchar(10) VC_EDI_MODE SILENTLY stores a 2-char value (NO DB backstop -> "
              "the save-guard is load-bearing)", db_stored_2char, "emrow=%s" % emrow)

    # bad DUNS (non-9/13, or non-digit): source-truth REJECTS.
    bad_duns = valid_values("ZZBDN"); bad_duns["form_duns"] = "12345"
    rep.check("validation: bad DUNS (5 digits) REJECTED by the source-truth oracle",
              source_truth_reject_reason(bad_duns) == "DUNS must be 9 or 13 digits",
              source_truth_reject_reason(bad_duns))
    bad_duns2 = valid_values("ZZBD2"); bad_duns2["form_duns"] = "12345678X"
    rep.check("validation: non-digit DUNS REJECTED by the source-truth oracle",
              source_truth_reject_reason(bad_duns2) == "DUNS must be 9 or 13 digits",
              source_truth_reject_reason(bad_duns2))

    # non-vacuity for the oracle: an ALL-VALID map passes (so the predicate isn't trivially always-reject)
    valid_9 = valid_values()                                  # 9-digit DUNS
    valid_13 = dict(valid_values(), form_duns="1234567890123")  # 9+4 = 13-digit DUNS
    rep.check("oracle non-vacuity: fully-valid value maps pass (9-digit AND 9+4 DUNS)",
              source_truth_reject_reason(valid_9) is None and source_truth_reject_reason(valid_13) is None,
              "9=%r 13=%r" % (source_truth_reject_reason(valid_9), source_truth_reject_reason(valid_13)))

    # ---- 4. SAVE-PATH GUARD — drive bad values through the DEPLOYED view's ACTUAL save_script ----
    # (FIX 2) Not the independent oracle: we run gsv.save_script() — verified below byte-identical to
    # the deployed sites-save-btn — with a stub system.db that records whether a DB write was reached.
    # write_attempted=False => the save path REJECTED the input (early return, NO write). This is the
    # guard that stops a malformed positional ISA15 (or bad DUNS) from ever reaching TEMA.
    save_src = gsv.save_script()
    deployed_save = _deployed_save_script()
    rep.check("save-path harness drives the DEPLOYED save_script (byte-identical to gateway view)",
              deployed_save is not None and deployed_save == save_src,
              "deployed==generator: %s" % (deployed_save == save_src))

    # 4a. a fully-valid value PASSES the guard (reaches the db write) — non-vacuity for the harness.
    w_ok, msg_ok = run_save_path(valid_values("ZZSPV"))
    rep.check("save-path: a fully-valid value PASSES the guard (reaches the db write)",
              w_ok is True, "write_attempted=%s msg=%r" % (w_ok, msg_ok))

    # 4b. a 2-char VC_EDI_MODE is REJECTED by the save path itself (no db write).
    bad_em_sp = valid_values("ZZSPE"); bad_em_sp["form_edimode"] = "PP"
    w_em, msg_em = run_save_path(bad_em_sp)
    rep.check("save-path: 2-char VC_EDI_MODE REJECTED by the save_script (no db write)",
              w_em is False and "exactly 1 character" in msg_em,
              "write_attempted=%s msg=%r" % (w_em, msg_em))

    # 4c. a 2-char separator is REJECTED by the save path itself (no db write).
    bad_sep_sp = valid_values("ZZSPS"); bad_sep_sp["form_sepelem"] = "**"
    w_sp, msg_sp = run_save_path(bad_sep_sp)
    rep.check("save-path: 2-char separator REJECTED by the save_script (no db write)",
              w_sp is False and "exactly 1 character" in msg_sp,
              "write_attempted=%s msg=%r" % (w_sp, msg_sp))

    # 4d. a bad DUNS (5 digits) is REJECTED by the save path itself (no db write).
    bad_dn_sp = valid_values("ZZSPD"); bad_dn_sp["form_duns"] = "12345"
    w_dn, msg_dn = run_save_path(bad_dn_sp)
    rep.check("save-path: bad DUNS (5 digits) REJECTED by the save_script (no db write)",
              w_dn is False and "DUNS must be" in msg_dn,
              "write_attempted=%s msg=%r" % (w_dn, msg_dn))

    # 4e. REVERT-PROOF (the test must be able to fail on a wrong rebuild): neutralize each guard's
    # condition in the save source and confirm the SAME bad value then SLIPS THROUGH to the db write.
    # This proves the GUARD — not the harness — is what blocks the malformed value.
    rev_em = save_src.replace('\tif edimode != "" and len(edimode) != 1:', '\tif False:')
    w_rev_em, _ = run_save_path(bad_em_sp, save_src=rev_em)
    rep.check("save-path REVERT-PROOF: removing the EDI-mode guard lets the 2-char value slip to db write",
              rev_em != save_src and w_rev_em is True, "write_attempted=%s" % w_rev_em)

    rev_sep = save_src.replace('\t\tif _v != "" and len(_v) != 1:', '\t\tif False:')
    w_rev_sep, _ = run_save_path(bad_sep_sp, save_src=rev_sep)
    rep.check("save-path REVERT-PROOF: removing the separator guard lets the 2-char sep slip to db write",
              rev_sep != save_src and w_rev_sep is True, "write_attempted=%s" % w_rev_sep)

    rev_dn = save_src.replace('\tif not _duns_ok(duns):', '\tif False:')
    w_rev_dn, _ = run_save_path(bad_dn_sp, save_src=rev_dn)
    rep.check("save-path REVERT-PROOF: removing the DUNS guard lets the bad DUNS slip to db write",
              rev_dn != save_src and w_rev_dn is True, "write_attempted=%s" % w_rev_dn)

    # ---- restore-as-found check ----
    end_count = int(scalar("SELECT COUNT(*) FROM INV_SITES"))
    rep.check("spike restored as-found (INV_SITES row count unchanged — all CRUD rolled back)",
              end_count == seed_count, "seed=%d end=%d" % (seed_count, end_count))

    sys.exit(rep.summary_exit())


if __name__ == "__main__":
    main()
