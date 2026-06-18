"""jython_shim.py — a headless `system.db` shim so the REAL gateway-side Jython wrappers can be driven
end-to-end against the spike DB without the gateway runtime / Perspective / a live trial (retro R8).

The producer write-then-post seams (stocktaking/reject/shipping/receiving) and the Order commit path are
pure Jython that only touch the gateway globals `system` and (for posts) the `stockLedger` library. This
shim emulates the slice of `system.db` they use — runPrepQuery / runPrepUpdate(getKey) / createSProcCall
+ registerInParam + execSProcCall + the type constants — backed by `sqlcmd` in the mssql-spike container.
Loading a wrapper module with this `system` injected runs its ACTUAL code (the dynamic SQL assembly,
getKey round-trip, _readRow re-read, branch selection, the stockLedger.post funnel) — closing the gap
that the SQL-reimplementing integration tests left open.

Scope: autocommit only (the producer seams don't open transactions; Order's beginTransaction path is a
documented extension — see notes). Jython 2.7 vs CPython differences aren't covered (the wrappers are
written 2.7/CPython-portable, verified by the reviewers); this catches the LOGIC/SQL the proc-EXEC tests
couldn't.
"""
import os, subprocess, sys, importlib.util

CONTAINER = os.environ.get("CONTAINER", "mssql-spike")
SA_PASS = os.environ.get("SA_PASS")
SQL_DB = "Inventory"   # the SQL Server DB; the wrappers' logical name "Inventory_Spike" maps here


def _run(sql_text, want_rows):
    """Execute one batch via sqlcmd. want_rows=True keeps the header row (for column-name access)."""
    if not SA_PASS:
        sys.exit("export SA_PASS first")
    args = ["docker", "exec", CONTAINER, "/opt/mssql-tools18/bin/sqlcmd", "-C", "-S", "localhost",
            "-U", "sa", "-P", SA_PASS, "-d", SQL_DB, "-W", "-s", "\t"]
    if not want_rows:
        args += ["-h", "-1"]
    out = subprocess.check_output(args + ["-Q", "SET NOCOUNT ON; " + sql_text], text=True)
    return out


def _esc(v):
    """Inline a bound param as a T-SQL literal (the shim has no real ? binding)."""
    if v is None:
        return "NULL"
    if isinstance(v, bool):
        return "1" if v else "0"
    if isinstance(v, (int, float)):
        return repr(v)
    return "N'" + unicode(v).replace("'", "''") + "'" if sys.version_info[0] < 3 else "N'" + str(v).replace("'", "''") + "'"


def _bind(sql, args):
    """Replace each ? with the next escaped arg (left to right)."""
    out, i = [], 0
    for ch in sql:
        if ch == "?":
            out.append(_esc(args[i])); i += 1
        else:
            out.append(ch)
    return "".join(out)


class _PyDataset(object):
    """Minimal Ignition PyDataset stand-in: len(), [i] -> row, row[colName], iteration, .rowCount,
    .getValueAt(r, colNameOrIdx)."""
    def __init__(self, columns, rows):
        self.columns = columns
        self.rows = rows                      # list of lists (string cells)
        self.rowCount = len(rows)

    def _coerce(self, s):
        if s == "NULL":
            return None
        try:
            return int(s)
        except ValueError:
            try:
                return float(s)
            except ValueError:
                return s

    class _Row(object):
        def __init__(self, ds, cells):
            self._ds, self._cells = ds, cells
        def __getitem__(self, key):
            idx = key if isinstance(key, int) else self._ds.columns.index(key)
            return self._ds._coerce(self._cells[idx])

    def __len__(self):
        return self.rowCount

    def __getitem__(self, i):
        return _PyDataset._Row(self, self.rows[i])

    def __iter__(self):
        return (self[i] for i in range(self.rowCount))

    def getValueAt(self, r, col):
        idx = col if isinstance(col, int) else self.columns.index(col)
        return self._coerce(self.rows[r][idx])


def _parse_rows(out):
    """Parse sqlcmd -W output WITH the header row into (columns, rows). sqlcmd emits: header, a dashes
    separator, then data; messages/row-count lines are filtered."""
    lines = [l.rstrip("\n") for l in out.splitlines()
             if l.strip() and not l.startswith("(") and not l.startswith("Msg ")]
    if not lines:
        return [], []
    columns = lines[0].split("\t")
    data = [ln.split("\t") for ln in lines[1:] if set(ln) - set("- \t")]  # drop the dashes separator
    return columns, data


class _SProcCall(object):
    def __init__(self, name):
        self.name, self.params = name, []
    def registerInParam(self, name, _type, value):
        self.params.append((name, value))


class _DB(object):
    # type constants the wrappers pass to registerInParam (values irrelevant — shim inlines literals)
    INTEGER = "INTEGER"; VARCHAR = "VARCHAR"; BIT = "BIT"

    def runPrepQuery(self, sql, args, db=None):
        return _PyDataset(*_parse_rows(_run(_bind(sql, args), want_rows=True)))

    def runPrepUpdate(self, sql, args, db=None, getKey=False, tx=None):
        stmt = _bind(sql, args)
        if getKey:
            out = _run(stmt + "; SELECT CAST(SCOPE_IDENTITY() AS int) AS k", want_rows=True)
            cols, rows = _parse_rows(out)
            return int(rows[-1][0]) if rows else None
        _run(stmt, want_rows=False)
        return 1

    def createSProcCall(self, name, db=None):
        return _SProcCall(name)

    def execSProcCall(self, call):
        names = [n for n, _ in call.params]
        vals = [v for _, v in call.params]
        exec_sql = "EXEC %s %s" % (call.name, ", ".join("@%s=?" % n for n in names))
        self.runPrepUpdate(exec_sql, vals, getKey=False)

    # autocommit shim: transactions are no-ops (producer seams don't use them; Order's path is a
    # documented extension — it would need a persistent sqlcmd session to span statements).
    def beginTransaction(self, db=None, **k):
        return "tx-noop"
    def commitTransaction(self, tx):
        pass
    def rollbackTransaction(self, tx):
        pass
    def closeTransaction(self, tx):
        pass


class _System(object):
    def __init__(self):
        self.db = _DB()
        class _Util(object):
            class _Logger(object):
                def info(self, *a): pass
                def warn(self, *a): pass
            def getLogger(self, name):
                return _System._Util._Logger()
        self.util = _Util()


def load_wrapper(name, code_path, extra_globals=None):
    """Import a project-library code.py with `system` (+ any extra globals like `stockLedger`) injected,
    so its functions run for real under this shim. Returns the module."""
    spec = importlib.util.spec_from_file_location(name, code_path)
    mod = importlib.util.module_from_spec(spec)
    mod.system = _System()
    if extra_globals:
        for k, v in extra_globals.items():
            setattr(mod, k, v)
    spec.loader.exec_module(mod)
    # the module body may reference `system` only inside functions; ensure it's bound post-exec too
    mod.system = mod.system if hasattr(mod, "system") else _System()
    return mod
