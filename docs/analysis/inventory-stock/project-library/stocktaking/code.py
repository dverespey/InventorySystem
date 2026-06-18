# stocktaking — Project Library service that re-homes the three Stocktaking qty triggers
# (INSERT_/UPDATE_/DELETE_Stocktaking on INV_STOCKTAKING_INF) into stockLedger.post() calls.
#
# Design: docs/analysis/inventory-stock/IGNITION-stock-ledger-design.md §3 rows #12-#14;
#         source spec: docs/analysis/inventory-stock/stocktaking.md.
#
# A stocktaking row is a SIGNED ADJUSTMENT delta (D5): IN_QTY is NOT an absolute count, it is the
# +/- correction to on-hand, posted VERBATIM. So effect(row) = +IN_QTY (a -5 row removes 5). Like
# reject it is INT-KEYED on IN_PART_ID (no string->id resolution). No add-point gate.
#
# Verified against the LIVE trigger bodies (CreateInventory.sql), all joined on IN_PART_ID:
#   INSERT_Stocktaking : IN_QTY = PS.IN_QTY + i.IN_QTY            -> delta = +new.IN_QTY
#   DELETE_Stocktaking : IN_QTY = PS.IN_QTY - d.IN_QTY            -> delta = -old.IN_QTY (reverse adj.)
#   UPDATE_Stocktaking : -d.IN_QTY (back out old, joined on d.IN_PART_ID) THEN
#                        +i.IN_QTY (apply new, joined on i.IN_PART_ID)   -> same-part net = new-old
#       Independent d/i joins -> a (domain-forbidden) part change is a faithful two-post.
#
# Two legacy quirks the rebuild corrects, both NON-qty (stock parity is unaffected):
#   * D8/Bug2 NULL-timestamp: UPDATE_StockTakingInfo builds VC_LAST_UPDATE from an UNINITIALIZED
#     @Update (SUBSTRING(@Update,1,8) on NULL) -> writes NULL. The rebuild's stocktaking-edit write
#     path writes a real fresh 16-char VC_LAST_UPDATE (design §3 #13); this module's amend key uses it.
#   * IN_QTY is NULL-able (DDL). A NULL adjustment is treated as a no-op (effect 0) rather than
#     replicating the legacy NULL-poisoning of IN_QTY (PS.IN_QTY + NULL = NULL).
#
# Jython 2.7 (8.1 and 8.3). Pure logic imports nothing; CPython-importable for unit test
# (scripts/e2e/test_stocktaking_posts.py). Only the driver touches `system`/`stockLedger`.

ENUM_STOCKTAKING = "STOCKTAKING"   # local mirror of stockLedger.ENUM_STOCKTAKING; test asserts equality.


def _effect(row):
	"""Signed on-hand contribution of a stocktaking row: +IN_QTY (the adjustment, posted verbatim,
	D5). IN_QTY is NULL-able -> a NULL/blank adjustment contributes 0 (no-op), not a NULL-poison."""
	v = row.get("IN_QTY")
	if v is None or str(v).strip() == "":
		return 0
	return int(v)


def _stkId(row):
	v = row.get("IN_STOCKTAKING_ID")
	return None if v is None or str(v).strip() == "" else int(v)


def _partId(row):
	return int(row["IN_PART_ID"])


def _ver(row):
	"""Per-edit version token for amend keys = the row's VC_LAST_UPDATE. The rebuild's stocktaking-edit
	write path writes a fresh NON-NULL stamp (fixing the legacy Bug2 NULL), so successive amends are
	distinct events and a replay reuses the key. Insert/delete keys need no version (identity-stable)."""
	return str(row.get("VC_LAST_UPDATE") or "")


def _key(row, suffix):
	"""'STOCKTAKING:stk={IN_STOCKTAKING_ID}:{suffix}' — e.g. STOCKTAKING:stk=903:ins (design §4)."""
	return "%s:stk=%s:%s" % (ENUM_STOCKTAKING, _stkId(row), suffix)


def _post(partId, delta, eventKey, row, reason):
	return {
		"partId": int(partId),
		"delta": int(delta),
		"sourceEnum": ENUM_STOCKTAKING,
		"eventKey": eventKey,
		"sourceRowId": _stkId(row),
		"reason": reason,
	}


def computePosts(op, old=None, new=None):
	"""Translate one INV_STOCKTAKING_INF event into the signed stock movements to post.
	op: 'insert' | 'update' | 'delete'. old/new: dict rows (keys: IN_PART_ID, IN_QTY, IN_STOCKTAKING_ID,
	VC_LAST_UPDATE). Returns 0/1/2 post dicts."""
	if op == "insert":
		delta = _effect(new)                       # +new.IN_QTY (the signed adjustment)
		if delta == 0:
			return []
		return [_post(_partId(new), delta, _key(new, "ins"), new, "stocktaking: signed adjustment")]

	if op == "delete":
		delta = -_effect(old)                      # -old.IN_QTY (reverse the adjustment)
		if delta == 0:
			return []
		return [_post(_partId(old), delta, _key(old, "del"), old, "stocktaking: reverse adjustment")]

	if op == "update":
		if _partId(old) == _partId(new):
			delta = _effect(new) - _effect(old)    # new.IN_QTY - old.IN_QTY
			if delta == 0:
				return []
			# Amend key includes the TARGET adjustment (:to=) as well as the version stamp, so it is
			# collision-safe even at the stamp's hundredth-second resolution: two posting amends of the
			# same row cannot share (stamp, target) — if they targeted the same value the second has
			# delta 0 and posts nothing. (Without :to=, two sub-hundredth amends collide and the 2nd is
			# silently swallowed by the (IN_PART_ID, VC_SOURCE_EVENT) idempotency backstop.)
			return [_post(_partId(new), delta, _key(new, "upd:to=%d:v=%s" % (_effect(new), _ver(new))),
			              new, "stocktaking: amend adjustment net delta")]
		# part-id change (domain-forbidden; faithful two-post via the independent-join trigger):
		posts = []
		oldBackout = -_effect(old)                 # remove the old part's adjustment
		if oldBackout != 0:
			posts.append(_post(_partId(old), oldBackout, _key(new, "upd:del-old:v=%s" % _ver(new)),
			                   old, "stocktaking amend part-change: back out old part"))
		newApply = _effect(new)                    # apply the new part's adjustment
		if newApply != 0:
			posts.append(_post(_partId(new), newApply, _key(new, "upd:add-new:v=%s" % _ver(new)),
			                   new, "stocktaking amend part-change: apply new part"))
		return posts

	raise ValueError("stocktaking.computePosts: unknown op %r" % (op,))


# ---------------------------------------------------------------------------------------------
# Gateway driver — post through the stockLedger funnel (NO part resolution; the row holds IN_PART_ID).
# ---------------------------------------------------------------------------------------------
def postStocktakingEvent(op, old=None, new=None, site=1, database=None):
	"""Wire one INV_STOCKTAKING_INF insert/update/delete into the stock ledger. Call AFTER the
	stocktaking row is written. Insert+delete of the same adjustment is qty-neutral."""
	db = database if database is not None else stockLedger.DATABASE
	for p in computePosts(op, old, new):
		stockLedger.post(p["partId"], p["delta"], p["sourceEnum"], p["eventKey"],
		                 sourceRowId=p["sourceRowId"], reason=p["reason"],
		                 site=site, purge=False, database=db)


# =============================================================================================
# WRITE-THEN-POST — the amend/write seam the Stocktaking screen calls (the Stage-3 reimplement).
# Each function writes the INV_STOCKTAKING_INF source row, re-reads it, then posts the matching
# ledger movement through postStocktakingEvent. This is the CUTOVER write path: it owns IN_QTY via
# the ledger, so it presumes the three legacy stocktaking triggers are GONE (during parallel run the
# legacy triggers stay live and the ledger is a shadow built by replay — do NOT run this path against
# a DB that still has those triggers, or IN_QTY double-counts).
#
# The source row is written with a fresh, non-NULL 16-char VC_LAST_UPDATE via the live SQL stamp
# recipe — this BOTH fixes the legacy Bug2 NULL-timestamp (UPDATE_StockTakingInfo built it off an
# uninitialized @Update) AND gives the amend its idempotency version token. We do NOT call the buggy
# UPDATE_StockTakingInfo proc; the amend is a correct UPDATE here.
#
# ATOMICITY (flagged for cutover, not solved here): the source-row write and the ledger post are two
# statements on (potentially) two connections, so they are NOT yet one atomic unit — a post failure
# after the source write commits would orphan the source row from the ledger. During parallel run this
# is moot (the ledger is a shadow). At cutover, make them atomic — combine into one stored proc, or
# thread a txId through stockLedger.post. Escalate the choice to the architect before cutover.

# The live 16-char yyyymmddHHMMSSff stamp recipe (matches INSERT_StockTakingInfo / the legacy triggers).
_STAMP_SQL = ("CONVERT(varchar, getdate(), 112) + SUBSTRING(CONVERT(varchar, getdate(), 114), 1, 2) "
              "+ SUBSTRING(CONVERT(varchar, getdate(), 114), 4, 2) "
              "+ SUBSTRING(CONVERT(varchar, getdate(), 114), 7, 2) "
              "+ SUBSTRING(CONVERT(varchar, getdate(), 114), 10, 2)")


def _readRow(stkId, db):
	"""Re-read a stocktaking row as a plain dict (the source of truth the post is built from)."""
	rows = system.db.runPrepQuery(
		"SELECT IN_STOCKTAKING_ID, IN_PART_ID, IN_QTY, VC_REASON, VC_LAST_UPDATE, VC_ADD "
		"FROM INV_STOCKTAKING_INF WHERE IN_STOCKTAKING_ID = ?", [stkId], db)
	if len(rows) == 0:
		return None
	r = rows[0]
	return {"IN_STOCKTAKING_ID": r["IN_STOCKTAKING_ID"], "IN_PART_ID": r["IN_PART_ID"],
	        "IN_QTY": r["IN_QTY"], "VC_REASON": r["VC_REASON"],
	        "VC_LAST_UPDATE": r["VC_LAST_UPDATE"], "VC_ADD": r["VC_ADD"]}


def insertAdjustment(partId, qty, reason=None, site=1, database=None):
	"""Record a new stocktaking adjustment for partId (signed qty, D5) and post +qty to the ledger.
	Returns the new IN_STOCKTAKING_ID."""
	db = database if database is not None else stockLedger.DATABASE
	newId = system.db.runPrepUpdate(
		"INSERT INTO INV_STOCKTAKING_INF (IN_PART_ID, IN_QTY, VC_REASON, VC_LAST_UPDATE, VC_ADD) "
		"VALUES (?, ?, ?, " + _STAMP_SQL + ", " + _STAMP_SQL + ")",
		[partId, qty, reason], db, getKey=True)
	postStocktakingEvent("insert", new=_readRow(newId, db), site=site, database=db)
	return newId


def amendAdjustment(stkId, newQty, newReason=None, site=1, database=None):
	"""Edit an existing adjustment's qty/reason and post the net delta. Writes a fresh non-NULL
	VC_LAST_UPDATE (fixes Bug2; serves as the amend version token)."""
	db = database if database is not None else stockLedger.DATABASE
	old = _readRow(stkId, db)
	if old is None:
		raise ValueError("stocktaking.amendAdjustment: no row IN_STOCKTAKING_ID=%r" % (stkId,))
	system.db.runPrepUpdate(
		"UPDATE INV_STOCKTAKING_INF SET IN_QTY = ?, VC_REASON = ?, VC_LAST_UPDATE = " + _STAMP_SQL +
		" WHERE IN_STOCKTAKING_ID = ?", [newQty, newReason, stkId], db)
	postStocktakingEvent("update", old=old, new=_readRow(stkId, db), site=site, database=db)


def deleteAdjustment(stkId, site=1, database=None):
	"""Remove an adjustment and post the reversal (-its qty)."""
	db = database if database is not None else stockLedger.DATABASE
	old = _readRow(stkId, db)
	if old is None:
		return
	system.db.runPrepUpdate("DELETE FROM INV_STOCKTAKING_INF WHERE IN_STOCKTAKING_ID = ?", [stkId], db)
	postStocktakingEvent("delete", old=old, site=site, database=db)
