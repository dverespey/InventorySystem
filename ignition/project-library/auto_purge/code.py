# auto_purge — Project Library service: retention-based DATAPURGE, TRANSACTIONAL, gateway-schedulable.
#
# M4 piece-3 hardening HD4 (m4-auth-sites-sourcetruth.md §6), SINGLE-SITE faithful rebuild of the legacy
# DELETE_AutoPurge (CreateInventory.sql:7682), driven the legacy way from DataModule.AutoPurge
# (DataModule.pas:6885-6929) with the [DATAPURGE] knobs (fiEnableDataPurge / fiPromptDataPurge /
# fiDataRetention).
#
# WHAT THE LEGACY DOES (decoded byte-for-byte from the proc + its caller):
#   * DataModule.AutoPurge (DataModule.pas:6885):
#       - rejects retention < 12 months UP FRONT (line 6890) -> the >=12 floor (also INV_SITES CHECK).
#       - passes @DataRentention = 0 - fiDataRetention.AsInteger  (i.e. NEGATIVE; e.g. -12) (line 6904).
#       - EXEC dbo.DELETE_AutoPurge;1 @DataRentention=<negative>.
#   * DELETE_AutoPurge (CreateInventory.sql:7682) — the cutoff is a 16-char string
#       YYYYMMDD + HH + MM + SS + ff  (style 112 (8) + substrings 1/4/7/10 of style 114), matching
#       VC_ADD varchar(16). With @DataRentention NEGATIVE, DATEADD(MONTH,@r,getdate()) is RETENTION
#       MONTHS IN THE PAST -> the cutoff. Then, in order:
#         1) UPDATE INV_OPEN_ORDER_INF SET VC_TERMINATED=<8-char date> WHERE VC_ADD <= <cutoff16> AND
#            VC_TERMINATED = ''   (terminate the still-open old orders BEFORE deleting them)
#         2) DELETE  INV_OPEN_ORDER_INF        WHERE VC_ADD <= <cutoff16>
#         3) DELETE  INV_OPEN_ORDER_INF_HIST   WHERE VC_ADD <= <cutoff16>
#         4) DELETE  INV_PARTS_STOCK_MST_HIST  WHERE VC_ADD <= <cutoff16>
#       NO site filter (single-site legacy; faithful here — single-site rebuild, so NO site predicate).
#       NOT transactional (each statement RETURN @err on failure -> a partial purge was possible).
#
# WHAT THIS REBUILD DOES (faithful scope, hardened where required by HD4):
#   * SAME 4 statements, SAME tables, SAME VC_ADD <= cutoff16 semantics, SAME >=12 floor, NO site filter
#     (single-site, exactly like the legacy). This is the SCOPE faithfulness sql-adversary checks.
#   * TRANSACTIONAL: all 4 run in ONE system.db transaction (BEGIN/COMMIT/ROLLBACK). A forced failure in
#     ANY statement rolls back ALL of them -> nothing purged (closes the legacy partial-purge hole).
#   * RETENTION READ FROM INV_SITES.IN_DATA_RETENTION (the [DATAPURGE] knob's relocation, Q17), not an
#     INI. The enable/prompt knobs (BIT_ENABLE_DATA_PURGE / BIT_PROMPT_DATA_PURGE) are also read from
#     INV_SITES (a scheduled gateway job honors BIT_ENABLE; PROMPT is a UI concern, surfaced for the
#     manual run).
#   * FAITHFUL to the legacy on the cutoff: `_CUTOFF_EXPR` (with its own GETDATE()) is inlined into EACH
#     statement, so SQL Server re-evaluates the cutoff per-statement exactly as the legacy proc does (any
#     sub-second drift between statements is identical legacy behavior). No divergence here.
#
# HEADLESS-DRIVABLE (retro R8): the SQL assembly + the transaction control here are the gateway-side
# driver; scripts/e2e/test_auto_purge.py drives THIS module through jython_shim (real _TxSession) so the
# transaction + rollback are exercised, not just an EXEC of the proc. The retention READ is a real query.
#
# Jython 2.7 (8.1.52 + 8.3, no scripting delta). IG81-COMPAT: system.db.beginTransaction / runPrepUpdate /
# runScalarPrepQuery are 8.1+ and unchanged on 8.3. Schedule via a gateway Timer/Scheduled script
# (8.1-safe) — NOT an 8.3 Event Stream — calling run_scheduled_purge() (see docstring there).

from db_shared import CONNECTION as DATABASE   # centralized DB-conn (default Inventory_Spike; single prod-rename pt; NO creds — see m4-hardening-secrets.md)

# The retention floor (legacy DataModule.pas:6890 rejects < 12; INV_SITES CK_INV_SITES_DATA_RETENTION).
MIN_RETENTION_MONTHS = 12

# The purge scope, in legacy order. (table, has_terminate_update). Only INV_OPEN_ORDER_INF gets the
# pre-delete VC_TERMINATED stamp (statement 1); the other two are delete-only. Kept as data so the scope
# is auditable in one place and sql-adversary can diff it against DELETE_AutoPurge.
PURGE_DELETE_TABLES = ("INV_OPEN_ORDER_INF", "INV_OPEN_ORDER_INF_HIST", "INV_PARTS_STOCK_MST_HIST")
TERMINATE_TABLE = "INV_OPEN_ORDER_INF"


# ===========================================================================
# PURE-ish core: build the legacy cutoff string + the statements. No transaction here (so the SQL the
# transaction runs is inspectable + testable in isolation).
# ===========================================================================

# The 16-char cutoff = style 112 (YYYYMMDD) + substrings 1/4/7/10 of style 114 (HH MM SS ff), built on
# the SERVER from DATEADD(MONTH, <-retention>, GETDATE()). Built as a T-SQL expression (not a Python
# string) so the cutoff is computed by SQL Server at run time, byte-identical to the legacy proc.
_CUTOFF_EXPR = (
	"(CONVERT(char(8),DATEADD(MONTH,?,GETDATE()),112)"
	" + SUBSTRING(CONVERT(char(12),DATEADD(MONTH,?,GETDATE()),114),1,2)"
	" + SUBSTRING(CONVERT(char(12),DATEADD(MONTH,?,GETDATE()),114),4,2)"
	" + SUBSTRING(CONVERT(char(12),DATEADD(MONTH,?,GETDATE()),114),7,2)"
	" + SUBSTRING(CONVERT(char(12),DATEADD(MONTH,?,GETDATE()),114),10,2))"
)

# The 8-char terminate date (style 112) the legacy writes into VC_TERMINATED (statement 1's SET value).
_TERMINATE_VAL_EXPR = "CONVERT(char(8),DATEADD(MONTH,?,GETDATE()),112)"


def _neg_retention(retentionMonths):
	"""The legacy passes @DataRentention = 0 - retention (NEGATIVE), so DATEADD adds a negative month
	count -> the cutoff is retention months in the PAST. Mirror that sign here (DataModule.pas:6904)."""
	return 0 - int(retentionMonths)


def validate_retention(retentionMonths):
	"""The >=12 floor (legacy DataModule.pas:6890; INV_SITES CHECK). Raises ValueError below 12 — the
	driver REFUSES to purge with too-short a retention (faithful: the legacy aborts with a message)."""
	r = int(retentionMonths)
	if r < MIN_RETENTION_MONTHS:
		raise ValueError("Invalid data retention %d — must be >= %d months (legacy floor)."
		                 % (r, MIN_RETENTION_MONTHS))
	return r


# ===========================================================================
# RETENTION + the enable/prompt knobs from INV_SITES (the [DATAPURGE] relocation, Q17).
# ===========================================================================

def read_purge_config(db=DATABASE):
	"""Read the single-site [DATAPURGE] config from INV_SITES (single-site: the one site row). Returns a
	dict {retention, enable, prompt}. The retention is IN_DATA_RETENTION; enable/prompt are the BIT knobs.

	Single-site: there is ONE operative site row. We read MIN(IN_SITE_ID) deterministically (the seeded
	site). (Multi-site is a later M4 concern — then the scheduled job would loop sites; HD4's site-filter
	note. Single-site here = NO site loop, NO site filter, faithful to the legacy single-INI install.)"""
	sql = ("SELECT TOP 1 IN_DATA_RETENTION, BIT_ENABLE_DATA_PURGE, BIT_PROMPT_DATA_PURGE "
	       "FROM INV_SITES ORDER BY IN_SITE_ID")
	ds = system.db.runPrepQuery(sql, [], db)
	if ds is None or ds.rowCount == 0:
		raise ValueError("No INV_SITES row — cannot resolve DATAPURGE retention.")
	ret = ds.getValueAt(0, 0)
	enable = ds.getValueAt(0, 1)
	prompt = ds.getValueAt(0, 2)
	return {
		"retention": int(ret) if ret not in (None, "") else None,
		"enable": _as_bool(enable),
		"prompt": _as_bool(prompt),
	}


def _as_bool(v):
	if isinstance(v, bool):
		return v
	if v in (None, ""):
		return False
	try:
		return int(v) != 0
	except (TypeError, ValueError):
		return unicode(v).strip().lower() in ("1", "true", "y", "yes")


# ===========================================================================
# THE PURGE — transactional. The HD4 hardening: all 4 statements in ONE tx; any failure -> rollback all.
# ===========================================================================

def auto_purge(retentionMonths, db=DATABASE):
	"""Run the retention-based purge TRANSACTIONALLY. Faithful scope of DELETE_AutoPurge (single-site, NO
	site filter): terminate-then-delete old INV_OPEN_ORDER_INF, delete old _HIST + INV_PARTS_STOCK_MST_HIST,
	all keyed VC_ADD <= the legacy 16-char cutoff (retention months in the past).

	retentionMonths : months to keep (POSITIVE; the legacy sign-flip to negative happens here). Must be
	                  >= MIN_RETENTION_MONTHS (raises ValueError otherwise — REFUSES to purge, faithful).
	db              : the named gateway connection (default the Inventory rebuild conn).

	Returns a dict of affected-row counts {terminated, INV_OPEN_ORDER_INF, INV_OPEN_ORDER_INF_HIST,
	INV_PARTS_STOCK_MST_HIST}. On ANY statement failure: rolls back the WHOLE tx (NOTHING purged) and
	re-raises — closing the legacy partial-purge hole (the legacy ran each statement separately).
	"""
	log = system.util.getLogger("SPIKE")
	r = validate_retention(retentionMonths)        # the >=12 floor; REFUSE below it
	neg = _neg_retention(r)
	counts = {"terminated": 0, "INV_OPEN_ORDER_INF": 0,
	          "INV_OPEN_ORDER_INF_HIST": 0, "INV_PARTS_STOCK_MST_HIST": 0}

	tx = system.db.beginTransaction(db)
	try:
		# --- statement 1: terminate the still-open old orders (VC_TERMINATED='' AND VC_ADD <= cutoff) ---
		# args order: [terminate-val retention] + [5 cutoff retentions] (one per DATEADD in the expr).
		term_sql = ("UPDATE " + TERMINATE_TABLE + " SET VC_TERMINATED = " + _TERMINATE_VAL_EXPR + " "
		            "WHERE VC_ADD <= " + _CUTOFF_EXPR + " AND VC_TERMINATED = ''")
		counts["terminated"] = system.db.runPrepUpdate(term_sql, [neg] + [neg] * 5, db, tx=tx)

		# --- statements 2-4: delete old rows from each table, VC_ADD <= cutoff ---
		for tbl in PURGE_DELETE_TABLES:
			del_sql = ("DELETE FROM " + tbl + " WHERE VC_ADD <= " + _CUTOFF_EXPR)
			counts[tbl] = system.db.runPrepUpdate(del_sql, [neg] * 5, db, tx=tx)

		system.db.commitTransaction(tx)
		log.info("SPIKE auto_purge COMMIT retention=%d (neg=%d) counts=%s" % (r, neg, counts))
		return counts
	except Exception as e:
		# HD4: any failure rolls back ALL 4 -> nothing purged. (Closes the legacy partial-purge hole.)
		system.db.rollbackTransaction(tx)
		log.warn("SPIKE auto_purge ROLLBACK retention=%d err=%s (nothing purged)" % (r, unicode(e)))
		raise
	finally:
		system.db.closeTransaction(tx)


def run_scheduled_purge(db=DATABASE):
	"""The entry point a GATEWAY SCHEDULED/TIMER script calls (8.1-safe; NOT an 8.3 Event Stream). Reads
	the [DATAPURGE] config from INV_SITES (Q17) and runs auto_purge IFF BIT_ENABLE_DATA_PURGE is set.
	Single-site: one site row, no site loop. Returns a dict describing what happened so the schedule log /
	the manual-run UI can report it. Does NOT prompt (a scheduled job is unattended); the BIT_PROMPT knob
	is surfaced in the return for the manual/attended path to honor.

	Wire-up (gateway): Config -> Gateway Events -> Scheduled, cron e.g. nightly off-hours:
	    import auto_purge
	    auto_purge.run_scheduled_purge()
	"""
	log = system.util.getLogger("SPIKE")
	cfg = read_purge_config(db)
	if not cfg["enable"]:
		log.info("SPIKE auto_purge SKIPPED: BIT_ENABLE_DATA_PURGE is off")
		return {"ran": False, "reason": "disabled", "config": cfg}
	if cfg["retention"] is None:
		log.warn("SPIKE auto_purge SKIPPED: IN_DATA_RETENTION is NULL")
		return {"ran": False, "reason": "no-retention", "config": cfg}
	counts = auto_purge(cfg["retention"], db)
	return {"ran": True, "retention": cfg["retention"], "prompt": cfg["prompt"], "counts": counts}
