"""jython_shim.py — a headless `system.db` shim so the REAL gateway-side Jython wrappers can be driven
end-to-end against the spike DB without the gateway runtime / Perspective / a live trial (retro R8).

The producer write-then-post seams (stocktaking/reject/shipping/receiving) and the Order commit path are
pure Jython that only touch the gateway globals `system` and (for posts) the `stockLedger` library. This
shim emulates the slice of `system.db` they use — runPrepQuery / runPrepUpdate(getKey) / createSProcCall
+ registerInParam + execSProcCall + the type constants — backed by `sqlcmd` in the mssql-spike container.
Loading a wrapper module with this `system` injected runs its ACTUAL code (the dynamic SQL assembly,
getKey round-trip, _readRow re-read, branch selection, the stockLedger.post funnel) — closing the gap
that the SQL-reimplementing integration tests left open.

Scope: the producer seams run autocommit (a fresh `docker exec -Q` per call); Order's `commitOrders`
opens a transaction that spans statements, served by the persistent-sqlcmd-session extension below
(beginTransaction -> a long-lived `docker exec -i` connection; see `_TxSession`). Jython 2.7 vs CPython
differences aren't covered (the wrappers are
written 2.7/CPython-portable, verified by the reviewers); this catches the LOGIC/SQL the proc-EXEC tests
couldn't.
"""
import os, subprocess, sys, importlib.util, threading, time
try:
    import queue          # Py3
except ImportError:       # pragma: no cover - Jython/Py2 fallback (shim runs under CPython3 here)
    import Queue as queue

CONTAINER = os.environ.get("CONTAINER", "mssql-spike")
SA_PASS = os.environ.get("SA_PASS")
SQL_DB = "Inventory"   # the SQL Server DB; the wrappers' logical name "Inventory_Spike" maps here

# Logical Ignition connection name -> physical spike DB. The producer/Order seams only ever touch the
# default Inventory DB; the ASN create driver also READS AD_FRSPULL on the shared ALC datasource, so a
# query/proc-call can be routed to VehicleOrder by passing db="VehicleOrder" (or its connection name).
# Unknown names fall through to SQL_DB (the historical behaviour — db was previously ignored entirely).
_DB_MAP = {
    None: SQL_DB,
    "Inventory_Spike": SQL_DB,
    "Inventory": SQL_DB,
    "VehicleOrder": "VehicleOrder",
}


def _resolve_db(db):
    """Map a logical connection name to the physical spike DB; default to SQL_DB for unknown names so
    the autocommit Inventory path is unchanged for every existing caller."""
    return _DB_MAP.get(db, SQL_DB)


def _sqlcmd_args(db, interactive=False):
    """The shared sqlcmd invocation flags. `interactive=True` adds `-i` stdin streaming for the
    persistent transaction session (so multiple batches share ONE connection)."""
    if not SA_PASS:
        sys.exit("export SA_PASS first")
    base = ["docker", "exec"] + (["-i"] if interactive else []) + [
        CONTAINER, "/opt/mssql-tools18/bin/sqlcmd", "-C", "-S", "localhost",
        "-U", "sa", "-P", SA_PASS, "-d", db, "-W", "-s", "\t"]
    return base


# The ANSI SET options the real gateway JDBC connection runs with. sqlcmd's `-Q` default has
# QUOTED_IDENTIFIER OFF, but the Ignition JDBC driver opens connections with QUOTED_IDENTIFIER ON +
# ANSI_NULLS ON (the JDBC/SQL-standard default). This matters once a FILTERED INDEX (or an indexed
# view / computed-column index) exists on a table: SQL Server REQUIRES these options ON for any DML
# against that table (else Msg 1934). spike-asn-unique-guard.sql adds a filtered unique index to
# INV_ASN_MST, so the shim must mirror the gateway's session options to faithfully exercise the driver
# (and to not falsely fail on the harness's own sqlcmd defaults). Prepended to every batch.
_SESSION_SET = "SET QUOTED_IDENTIFIER ON; SET ANSI_NULLS ON; SET NOCOUNT ON; "


def _run(sql_text, want_rows, db=SQL_DB):
    """Execute one batch via a fresh (autocommit) sqlcmd. want_rows=True keeps the header row."""
    args = _sqlcmd_args(db)
    if not want_rows:
        args += ["-h", "-1"]
    out = subprocess.check_output(args + ["-Q", _SESSION_SET + sql_text], text=True)
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
    """Replace each ? PLACEHOLDER with the next escaped arg (left to right). A `?` inside a `--` line
    comment or inside a '...' string literal is NOT a placeholder and is left untouched — exactly as the
    real JDBC prepared-statement parser treats it. (Several deployed views carry a trailing
    `-- IG-SITE: AND site_id = ?` comment on a query whose bound args count only the REAL placeholders;
    a naive per-`?` bind would over-count and IndexError. This mirrors JDBC so the shim drives those
    views faithfully.)"""
    out, i = [], 0
    n = len(sql)
    k = 0
    in_str = False
    while k < n:
        ch = sql[k]
        if in_str:
            out.append(ch)
            if ch == "'":
                # doubled '' is an escaped quote inside the literal
                if k + 1 < n and sql[k + 1] == "'":
                    out.append("'"); k += 2; continue
                in_str = False
            k += 1
            continue
        if ch == "'":
            in_str = True; out.append(ch); k += 1; continue
        if ch == "-" and k + 1 < n and sql[k + 1] == "-":
            # line comment: copy verbatim to end-of-line (any ? in here is NOT a placeholder)
            eol = sql.find("\n", k)
            if eol == -1:
                out.append(sql[k:]); k = n
            else:
                out.append(sql[k:eol]); k = eol
            continue
        if ch == "?":
            out.append(_esc(args[i])); i += 1; k += 1; continue
        out.append(ch); k += 1
    return "".join(out)


class _PyDataset(object):
    """Minimal Ignition PyDataset stand-in: len(), [i] -> row, row[colName], iteration, .rowCount,
    .getValueAt(r, colNameOrIdx)."""
    def __init__(self, columns, rows):
        self.columns = columns
        self.rows = rows                      # list of lists (string cells)
        self.rowCount = len(rows)

    def _coerce(self, s):
        # NULL -> None; everything else stays a STRING. A real Ignition PyDataset preserves the column's
        # declared type, so a VARCHAR business key (e.g. a digit-only VC_PART_NUMBER like '478930223000')
        # comes back as a string and re-binds as a quoted literal. The shim parses untyped text output, so
        # earlier it greedily int-coerced any all-digit cell — which silently corrupted the digit-only part
        # number on the round-trip _readRow -> resolvePartId/resolveAddPoint bind (emitted UNQUOTED ->
        # "Error converting data type varchar to numeric" -> no part resolved -> no post). The int-keyed
        # producers (stocktaking/reject) hold the int id directly and never hit this; only the part-number-
        # resolving shipping/receiving drivers did. Every numeric the drivers read is explicitly int()-cast
        # by their own code, so keeping cells as strings is both correct and sufficient. (Genuine narrow
        # shim coercion gap — NOT the documented persistent-session/transaction extension.)
        return None if s == "NULL" else s

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


# ----------------------------------------------------------------------------------------------------
# Persistent-session transaction support (the documented extension — Order's commitOrders).
#
# The autocommit path (_run) is a fresh `docker exec sqlcmd -Q` per statement: a separate connection
# each time, so a `BEGIN TRAN` in one would NOT carry to the next. To make `beginTransaction` actually
# span statements we hold ONE long-lived `docker exec -i sqlcmd` subprocess open and feed it batches on
# stdin; all statements then run on the SAME connection/transaction until COMMIT/ROLLBACK.
#
# THE FIDDLY PART — batch framing. sqlcmd in stdin mode only executes a batch when it reads a `GO`. To
# know when a batch has FINISHED (and to capture its rows) we append a unique sentinel
# `PRINT '<<<EOB:nonce>>>'` and read stdout until that line appears. ERROR CAVEAT: when a batch raises a
# T-SQL error, sqlcmd ABORTS the rest of that batch — so the trailing PRINT sentinel NEVER executes and
# would hang the reader forever. We therefore watch for `Msg NNNN` error lines and, on the first one,
# drain the immediately-available error text and RAISE — which is exactly what lets `commitOrders`'
# try/except fire `rollbackTransaction`. (Verified: after such an error the session is still alive and
# accepts ROLLBACK on the same connection.) stdout is drained by a background thread into a queue so a
# missing sentinel times out instead of deadlocking.
#
# IG81-COMPAT: this is TEST-HARNESS plumbing only — the gateway provides real JDBC transactions via
# system.db.beginTransaction; nothing here ships. 8.1.52 and 8.3 expose the same beginTransaction API.
# ----------------------------------------------------------------------------------------------------
class SqlError(Exception):
    pass


class _TxSession(object):
    """One open sqlcmd connection holding a transaction across batches."""
    _seq = 0

    def __init__(self, db):
        self.db = db
        self.proc = subprocess.Popen(_sqlcmd_args(db, interactive=True), stdin=subprocess.PIPE,
                                     stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                     text=True, bufsize=1)
        self._q = queue.Queue()
        self._closed = False
        t = threading.Thread(target=self._pump)
        t.daemon = True
        t.start()
        # Mirror the gateway JDBC session options (QUOTED_IDENTIFIER/ANSI_NULLS ON) BEFORE the tx opens
        # so DML on the filtered-index INV_ASN_MST succeeds (see _SESSION_SET note; else Msg 1934).
        self._batch("SET QUOTED_IDENTIFIER ON; SET ANSI_NULLS ON; SET NOCOUNT ON;", want_rows=False)
        self._batch("BEGIN TRANSACTION;", want_rows=False)

    def _pump(self):
        for line in self.proc.stdout:
            self._q.put(line)
        self._q.put(None)               # EOF sentinel

    def _send(self, text):
        self.proc.stdin.write(text + "\n")
        self.proc.stdin.flush()

    @staticmethod
    def _is_err_line(line):
        """True for a sqlcmd error header. Two shapes occur:
          * `Msg <n>, Level <n>, State <n>, ...`  — the usual T-SQL error (e.g. Msg 2601 duplicate key).
          * `SqlState NNNNN, ...`                 — a driver/cursor-level error (e.g.
            `SqlState 24000, Invalid cursor state`) that sqlcmd emits when a batch runs after a
            statement-terminated error left a pending result. This line is NOT a `Msg` header, so the
            original detector missed it and the next batch hung waiting for a sentinel that never came.
        Recognizing both makes a unique-violation (FIX 1) — and the cursor-state fallout it leaves on the
        very next batch — surface as a catchable SqlError instead of a timeout."""
        if line.startswith("Msg ") and ", Level " in line and ", State " in line:
            return True
        if line.startswith("SqlState "):
            return True
        return False

    def _batch(self, sql_text, want_rows, timeout=30):
        """Run one batch on the session; return (columns, rows) when want_rows else None.
        Raises SqlError on any sqlcmd error header (so the caller's rollback path runs)."""
        _TxSession._seq += 1
        nonce = "EOB%d" % _TxSession._seq
        sentinel = "<<<%s>>>" % nonce
        self._send(sql_text)
        self._send("PRINT '%s';" % sentinel)   # framing marker (skipped by sqlcmd if the batch errored)
        self._send("GO")
        captured, err, end = [], None, time.time() + timeout
        while True:
            try:
                line = self._q.get(timeout=max(0.1, end - time.time()))
            except queue.Empty:
                if err is not None:
                    raise SqlError("; ".join(err))
                raise SqlError("tx batch timed out waiting for sentinel; partial=%r" % captured[:8])
            if line is None:
                raise SqlError("tx session closed unexpectedly during batch")
            line = line.rstrip("\n")
            if sentinel in line:
                break
            if self._is_err_line(line):
                # error -> the sentinel PRINT may or may not run; grab the immediate error text up to
                # (and consuming) the sentinel if it appears, then raise. Draining to the sentinel here
                # is what keeps the SESSION clean for the next batch (the rollback) — otherwise leftover
                # lines from this batch would be misread as the next batch's output.
                err = [line]
                while True:
                    try:
                        more = self._q.get(timeout=2)
                    except queue.Empty:
                        break
                    if more is None:
                        break
                    more = more.rstrip("\n")
                    if sentinel in more:
                        break
                    err.append(more)
                raise SqlError("; ".join(err))
            captured.append(line)
        if err is not None:
            raise SqlError("; ".join(err))
        if not want_rows:
            return None
        # parse the framed rows the same way _parse_rows handles -Q output (header, dashes, data)
        body = [l for l in captured if l.strip() and not l.startswith("(")]
        if not body:
            return [], []
        columns = body[0].split("\t")
        data = [ln.split("\t") for ln in body[1:] if set(ln) - set("- \t")]
        return columns, data

    def exec_update(self, stmt, get_key=False):
        if get_key:
            cols, rows = self._batch(stmt + "; SELECT CAST(SCOPE_IDENTITY() AS int) AS k", want_rows=True)
            return int(rows[-1][0]) if rows and rows[-1][0] != "NULL" else None
        # The real system.db.runPrepUpdate returns the AFFECTED-ROW COUNT (@@ROWCOUNT), not a constant.
        # The inbound 997/824 driver (edi_inbound/code.py) keys its unknown-EIN / unmatched-manifest
        # ALARM on a 0-row update (the legacy no-@@ROWCOUNT silent-success fix) — so the shim must return
        # the genuine @@ROWCOUNT, captured by a trailing SELECT on the SAME tx batch. (NB: NOCOUNT is ON
        # for the session, which suppresses the "N rows affected" message but NOT @@ROWCOUNT itself.)
        cols, rows = self._batch(stmt + "; SELECT @@ROWCOUNT AS rc", want_rows=True)
        return int(rows[-1][0]) if rows and rows[-1][0] not in ("NULL", "") else 0

    def query_scalar(self, stmt):
        """Run a batch on the live tx session and return the first cell of the LAST result row (the
        ASN-create captures INSERT_ASNInfo's OUTPUT @id via a trailing SELECT @id on the same tx).
        Mirrors system.db.runScalarPrepQuery(..., tx=). None when the cell is NULL / no rows."""
        cols, rows = self._batch(stmt, want_rows=True)
        if not rows:
            return None
        v = rows[-1][0]
        return None if v == "NULL" else v

    def commit(self):
        self._batch("COMMIT TRANSACTION;", want_rows=False)
        self.close()

    def rollback(self):
        # Guard with @@TRANCOUNT: a fatal statement error (e.g. a type-conversion failure) can DOOM and
        # auto-roll-back the transaction before our explicit ROLLBACK runs; an unconditional ROLLBACK
        # would then raise Msg 3903 ("no corresponding BEGIN") and MASK the original SqlError. The IF
        # makes rollback a no-op when SQL Server already unwound the tx — so commitOrders' real failure
        # propagates. (Real JDBC connection.rollback() is likewise safe on an already-rolled-back tx.)
        if not self._closed:
            # Best-effort: a statement-terminated error in the PRIOR batch (e.g. a unique violation,
            # FIX 1) can leave a transient `SqlState 24000, Invalid cursor state` on THIS rollback batch.
            # Rollback is cleanup of an already-failed tx — swallow such a follow-on error rather than
            # masking the caller's original exception (which is mid-propagation). Real JDBC rollback()
            # likewise does not throw the prior statement's error again.
            try:
                self._batch("IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;", want_rows=False)
            except SqlError as rbexc:
                # Only swallow the transient cursor-state follow-on; a GENUINE rollback failure must
                # surface (NIT-1) — else a real failed ROLLBACK would be masked as clean cleanup.
                msg = str(rbexc)
                if "24000" not in msg and "Invalid cursor state" not in msg:
                    raise

    def close(self):
        if self._closed:
            return
        self._closed = True
        try:
            self.proc.stdin.close()
        except Exception:
            pass
        try:
            self.proc.wait(timeout=10)
        except Exception:
            self.proc.kill()


class _SProcCall(object):
    def __init__(self, name):
        self.name, self.params = name, []
    def registerInParam(self, name, _type, value):
        self.params.append((name, value))


class _DB(object):
    # type constants the wrappers pass to registerInParam (values irrelevant — shim inlines literals)
    INTEGER = "INTEGER"; VARCHAR = "VARCHAR"; BIT = "BIT"

    def runPrepQuery(self, sql, args, db=None, tx=None):
        # Real 8.1+ API: runPrepQuery(query, args, [database], [tx]). Inside a live transaction, route
        # the SELECT to THAT session's connection so it sees the tx's own uncommitted writes (e.g. the
        # 856 driver reads the ASN header AFTER stamping the EIN on it in the same tx — an autocommit
        # read on a separate connection would BLOCK on the uncommitted row lock and hang). Returns a
        # PyDataset either way.
        stmt = _bind(sql, args)
        if isinstance(tx, _TxSession):
            cols, rows = tx._batch(stmt, want_rows=True)
            return _PyDataset(cols, rows)
        # Route to the named DB (VehicleOrder for AD_FRSPULL); defaults to Inventory for every other
        # caller. The Inventory autocommit path is byte-for-byte unchanged (db resolves to SQL_DB).
        return _PyDataset(*_parse_rows(_run(stmt, want_rows=True, db=_resolve_db(db))))

    def runPrepUpdate(self, sql, args, db=None, getKey=False, tx=None):
        stmt = _bind(sql, args)
        # Inside a live transaction, route the statement to THAT session's connection (so it shares the
        # open BEGIN TRAN). getKey is supported on the tx path too (SCOPE_IDENTITY in the framed output)
        # even though Order doesn't use it — keep it correct for future tx writers.
        if isinstance(tx, _TxSession):
            return tx.exec_update(stmt, get_key=getKey)
        # autocommit path (tx is None or the "tx-noop" sentinel) — UNCHANGED behaviour for db=None;
        # db now resolves through the same map so an explicit name routes correctly.
        rdb = _resolve_db(db)
        if getKey:
            out = _run(stmt + "; SELECT CAST(SCOPE_IDENTITY() AS int) AS k", want_rows=True, db=rdb)
            cols, rows = _parse_rows(out)
            return int(rows[-1][0]) if rows else None
        # Return the genuine @@ROWCOUNT (see _TxSession.exec_update) so a 0-row update is observable to
        # callers like the inbound 997/824 driver's unknown-EIN / unmatched-manifest alarm path.
        out = _run(stmt + "; SELECT @@ROWCOUNT AS rc", want_rows=True, db=rdb)
        cols, rows = _parse_rows(out)
        return int(rows[-1][0]) if rows and rows[-1][0] not in ("NULL", "") else 0

    def runScalarPrepQuery(self, sql, args, db=None, tx=None):
        """Real 8.1+ API: run a prepared statement, return the first row/first column (None if no rows).
        Supports the tx id so the ASN create can read INSERT_ASNInfo's OUTPUT @id inside its BEGIN TRAN.
        Autocommit path routes through the same db map as runPrepQuery."""
        stmt = _bind(sql, args)
        if isinstance(tx, _TxSession):
            return tx.query_scalar(stmt)
        cols, rows = _parse_rows(_run(stmt, want_rows=True, db=_resolve_db(db)))
        if not rows:
            return None
        v = rows[-1][0]
        return None if v == "NULL" else v

    def createSProcCall(self, name, db=None):
        return _SProcCall(name)

    def execSProcCall(self, call):
        names = [n for n, _ in call.params]
        vals = [v for _, v in call.params]
        exec_sql = "EXEC %s %s" % (call.name, ", ".join("@%s=?" % n for n in names))
        self.runPrepUpdate(exec_sql, vals, getKey=False)

    # Transactions: open a persistent sqlcmd session that spans statements (Order's commitOrders).
    # The producer seams pass no tx; they keep hitting the autocommit _run path above untouched.
    def beginTransaction(self, db=None, **k):
        # Resolve the logical connection name (e.g. "Inventory_Spike") to the physical spike DB so the
        # persistent sqlcmd -d target is real. Order passed no db (-> SQL_DB); create_asn passes the
        # Inventory connection name.
        return _TxSession(_resolve_db(db))
    def commitTransaction(self, tx):
        if isinstance(tx, _TxSession):
            tx.commit()
    def rollbackTransaction(self, tx):
        if isinstance(tx, _TxSession):
            tx.rollback()
    def closeTransaction(self, tx):
        # idempotent — commit/rollback may already have closed the session.
        if isinstance(tx, _TxSession):
            tx.close()


class _Logger(object):
    """A quiet stand-in for the gateway's named logger. Swallows info/warn/error/debug/trace so the
    REAL driver's logging calls run for real (any arg shape) without noise. (Set SHIM_LOG_ECHO=1 to
    echo to stderr when debugging a driver.)"""
    _echo = os.environ.get("SHIM_LOG_ECHO") == "1"

    def __init__(self, name):
        self.name = name

    def _emit(self, level, args):
        if self._echo:
            sys.stderr.write("[%s] %s %s\n" % (level, self.name, " ".join(str(a) for a in args)))

    def info(self, *a): self._emit("INFO", a)
    def warn(self, *a): self._emit("WARN", a)
    def error(self, *a): self._emit("ERROR", a)
    def debug(self, *a): self._emit("DEBUG", a)
    def trace(self, *a): self._emit("TRACE", a)


class _Util(object):
    def getLogger(self, name):
        return _Logger(name)


class _File(object):
    """Minimal stand-in for system.file (8.1+). The 856 driver writes the EDI-out file with
    writeFile(path, text). Mirrors the gateway signature (path, data[, append]); data may be a str or
    bytes. Test-harness only — the gateway provides the real system.file."""
    def writeFile(self, path, data, append=False):
        mode = "ab" if append else "wb"
        if isinstance(data, bytes):
            payload = data
        else:
            # the gateway writes the platform default; the 856 driver hands us a str of ASCII X12 with
            # explicit \r\n. Encode ASCII (X12 is 7-bit) and DO NOT let open() translate newlines (wb).
            payload = data.encode("ascii")
        with open(path, mode) as fh:
            fh.write(payload)

    def readFileAsString(self, path, encoding="UTF-8"):
        with open(path, "rb") as fh:
            return fh.read().decode(encoding)


class _Date(object):
    """Minimal stand-in for system.date (8.1+). The 856 driver calls now() + format(d, 'HHmm') for the
    one shared file timestamp. Java SimpleDateFormat patterns are translated to the strftime equivalents
    the driver actually uses (HHmm). Test-harness only."""
    import datetime as _dt

    _PATTERN = {"HHmm": "%H%M", "yyyyMMdd": "%Y%m%d", "yyyy-MM-dd HH:mm:ss": "%Y-%m-%d %H:%M:%S"}

    def now(self):
        return self._dt.datetime.now()

    def format(self, d, pattern):
        py = self._PATTERN.get(pattern)
        if py is None:
            raise ValueError("jython_shim._Date.format: unmapped pattern %r" % (pattern,))
        return d.strftime(py)


# ---------------------------------------------------------------------------
# system.user / system.perspective stand-ins (M4 auth). The auth project-library module
# (docs/analysis/production-readiness/project-library/auth/code.py) calls system.user.* (the user-mgmt
# ops) and resolves the caller's roles from the SESSION (server-side). These stand-ins back those calls
# with an in-memory user source so test_m4_auth.py drives the REAL gateway-side gate + user ops headless
# WITHOUT mutating the live gateway user source. Mirrors the 8.1 PyUser / UIResponse / addUser shape.
# ---------------------------------------------------------------------------

class _UIResponse(object):
    """Minimal system.user.* UIResponse: getErrors()/getWarns()/getInfos()."""
    def __init__(self, errors=None, warns=None, infos=None):
        self._e, self._w, self._i = list(errors or []), list(warns or []), list(infos or [])

    def getErrors(self): return self._e
    def getWarns(self): return self._w
    def getInfos(self): return self._i


class _PyUser(object):
    """Minimal PyUser: username, roles, a 'password' marker (HASHED — we store a sentinel, NEVER the
    plaintext, mirroring the Internal source which hashes), and contact info (type,value) pairs.
    The auth module uses addRole / set('password', ..) / getContactInfo / setContactInfo."""
    PasswordLastChanged = "PasswordLastChanged"

    def __init__(self, username):
        self.username = username
        self._roles = []
        self._contacts = []                 # list of (type, value)
        self._password_set = False
        self._password_changed = None       # epoch-ish marker; the harness only checks set/never-set

    # --- the API the auth module + the User Admin list builder use ---
    def getUsername(self):
        return self.username

    def getRoles(self):
        return list(self._roles)

    def addRole(self, r):
        if r not in self._roles:
            self._roles.append(r)

    def set(self, key, value):
        if key == "password":
            # CRITICAL: never store the plaintext. Record only that a (hashed) password now exists.
            self._password_set = True
            self._password_changed = time.time()
        else:
            setattr(self, "_" + key, value)

    def get(self, key):
        if key == self.PasswordLastChanged:
            return self._password_changed
        return getattr(self, "_" + str(key), None)

    def getContactInfo(self):
        # 8.1: returns ContactInfo objects exposing .contactType / .value
        return [_ContactInfo(t, v) for (t, v) in self._contacts]

    def addContactInfo(self, ctype, value):
        # 8.1: addContactInfo(contactType, value)
        self._contacts.append((ctype, value))

    def removeContactInfo(self, ctype, value):
        # 8.1: removeContactInfo(type, value) — exact-match delete
        self._contacts = [c for c in self._contacts if c != (ctype, value)]

    # --- harness introspection ---
    def roles(self): return list(self._roles)
    def password_set(self): return self._password_set


class _ContactInfo(object):
    """8.1 ContactInfo: exposes .contactType / .value (properties) AND the getter shape for portability."""
    def __init__(self, t, v):
        self.contactType, self.value = t, v
    def getContactType(self): return self.contactType
    def getValue(self): return self.value


class _UserModule(object):
    """In-memory stand-in for system.user (per-source dicts of _PyUser). Seed test users via add_seed()."""
    def __init__(self):
        self._sources = {}                   # source name -> {username: _PyUser}

    def _src(self, source):
        return self._sources.setdefault(source, {})

    def getNewUser(self, source, username):
        return _PyUser(username)

    def addUser(self, source, user):
        s = self._src(source)
        if user.username in s:
            return _UIResponse(errors=["User '%s' already exists" % user.username])
        s[user.username] = user
        return _UIResponse(infos=["added %s" % user.username])

    def editUser(self, source, user):
        s = self._src(source)
        s[user.username] = user
        return _UIResponse(infos=["edited %s" % user.username])

    def removeUser(self, source, username):
        s = self._src(source)
        if username in s:
            del s[username]
            return _UIResponse(infos=["removed %s" % username])
        return _UIResponse(errors=["No such user '%s'" % username])

    def getUser(self, source, username):
        return self._src(source).get(username)

    def getUsers(self, source):
        return list(self._src(source).values())

    # --- harness helpers ---
    def add_seed(self, source, username, roles, password_set=True):
        u = _PyUser(username)
        for r in roles:
            u.addRole(r)
        if password_set:
            u.set("password", "x")
        self._src(source)[username] = u
        return u

    def usernames(self, source):
        return sorted(self._src(source).keys())


class _PerspectiveModule(object):
    """Minimal system.perspective for the gate. isAuthorized(isAllOf, securityLevels, sessionId) maps a
    session's roles -> the Authenticated/Roles/<Role> security levels (the convention the auth module +
    the Sites view both use). The harness registers a session's roles via register_session()."""
    def __init__(self):
        self._sessions = {}                  # sessionId -> set(roles)

    def register_session(self, sessionId, roles):
        self._sessions[sessionId] = set(roles or [])

    def isAuthorized(self, isAllOf, securityLevels, sessionId=None):
        roles = self._sessions.get(sessionId, set())
        have = set("Authenticated/Roles/" + r for r in roles)
        if not roles:
            return False
        want = set(securityLevels)
        return want <= have if isAllOf else bool(want & have)


class _System(object):
    def __init__(self):
        self.db = _DB()
        self.util = _Util()
        self.file = _File()
        self.date = _Date()
        self.user = _UserModule()
        self.perspective = _PerspectiveModule()


def _py2_builtins():
    """Jython 2.7 builtins the gateway provides but CPython3 does not, so a project-library module
    written for the gateway (which legitimately uses unicode/basestring/xrange) loads + runs under the
    CPython3 harness exactly as the gateway runs it. Injected into the loaded module's namespace."""
    return {
        "unicode": str,
        "basestring": str,
        "xrange": range,
        "unichr": chr,
    }


def load_wrapper(name, code_path, extra_globals=None):
    """Import a project-library code.py with `system` (+ any extra globals like `stockLedger`) injected,
    so its functions run for real under this shim. Returns the module."""
    spec = importlib.util.spec_from_file_location(name, code_path)
    mod = importlib.util.module_from_spec(spec)
    mod.system = _System()
    for k, v in _py2_builtins().items():
        setattr(mod, k, v)
    if extra_globals:
        for k, v in extra_globals.items():
            setattr(mod, k, v)
    spec.loader.exec_module(mod)
    # the module body may reference `system` only inside functions; ensure it's bound post-exec too
    mod.system = mod.system if hasattr(mod, "system") else _System()
    return mod
