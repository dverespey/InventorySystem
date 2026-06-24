# reject — Project Library service that re-homes the three RejectParts qty triggers
# (INSERT_/UPDATE_/DELETE_RejectParts on INV_REJECT_INF) into stockLedger.post() calls.
#
# Design: docs/analysis/inventory-stock/IGNITION-stock-ledger-design.md §3 rows #9-#11;
#         source spec: docs/analysis/receiving/recreject.md.
#
# A reject removes parts from on-hand. Like shipping it is unconditional stock-OUT (effect=-IN_QTY,
# no add-point gate) — but it is already INT-KEYED on IN_PART_ID (the row carries the surrogate, so
# NO string->id resolution at the boundary, unlike receiving/shipping). The reject delete is ALWAYS a
# genuine user un-reject that restores on-hand: DELETE_AutoPurge never touches INV_REJECT_INF (design
# F2), so there is NO purge case here (the old D11#9 "purge inflates on-hand" class is retired).
#
# Verified against the LIVE trigger bodies (CreateInventory.sql), all joined on IN_PART_ID:
#   INSERT_RejectParts : IN_QTY = PS.IN_QTY - i.IN_QTY            -> delta = -new.IN_QTY
#   DELETE_RejectParts : IN_QTY = PS.IN_QTY + d.IN_QTY            -> delta = +old.IN_QTY
#   UPDATE_RejectParts : +d.IN_QTY (restore old, joined on d.IN_PART_ID) THEN
#                        -i.IN_QTY (subtract new, joined on i.IN_PART_ID)   -> same-part net = old-new
#       The two legs join PS INDEPENDENTLY on d/i IN_PART_ID, so a (domain-forbidden) part change is
#       still handled correctly as a two-post — same shape as shipping, parity with the trigger.
#
# Jython 2.7 (8.1 and 8.3). Pure logic (computePosts) imports nothing; CPython-importable for unit
# test (scripts/e2e/test_reject_posts.py). Only the driver touches `system`/`stockLedger`.
# IG81-COMPAT: createSProcCall identical on 8.1.52 and 8.3.

ENUM_REJECT = "REJECT"   # local mirror of stockLedger.ENUM_REJECT (the wire contract); test asserts equality.


def _effect(row):
	"""Signed on-hand contribution of a reject row: ALWAYS -IN_QTY (a reject removes stock; no gate).
	INV_REJECT_INF.IN_QTY is NOT NULL (DDL), so no None guard needed here."""
	return -int(row["IN_QTY"])


def _rejectId(row):
	v = row.get("IN_REJECT_ID")
	return None if v is None or str(v).strip() == "" else int(v)


def _partId(row):
	return int(row["IN_PART_ID"])


def _ver(row):
	"""Per-edit version token for amend keys = the row's VC_LAST_UPDATE (the rebuild's reject-edit
	write path writes a fresh 16-char stamp per edit, as the UPDATE trigger reads i.VC_LAST_UPDATE).
	Insert/delete keys need no version (IN_REJECT_ID is inserted/deleted once — identity-stable)."""
	return str(row.get("VC_LAST_UPDATE") or "")


def _key(row, suffix):
	"""'REJECT:rej={IN_REJECT_ID}:{suffix}' — e.g. REJECT:rej=551:ins (design §4)."""
	return "%s:rej=%s:%s" % (ENUM_REJECT, _rejectId(row), suffix)


def _post(partId, delta, eventKey, row, reason):
	return {
		"partId": int(partId),
		"delta": int(delta),
		"sourceEnum": ENUM_REJECT,
		"eventKey": eventKey,
		"sourceRowId": _rejectId(row),
		"reason": reason,
	}


def computePosts(op, old=None, new=None):
	"""Translate one INV_REJECT_INF event into the signed stock movements to post.
	op: 'insert' | 'update' | 'delete'.  old/new: dict rows (keys: IN_PART_ID, IN_QTY, IN_REJECT_ID,
	VC_LAST_UPDATE). Returns 0/1/2 post dicts (two only if IN_PART_ID changed on an edit — forbidden by
	the domain but handled faithfully)."""
	if op == "insert":
		delta = _effect(new)                       # -new.IN_QTY
		if delta == 0:
			return []
		return [_post(_partId(new), delta, _key(new, "ins"), new, "reject: remove from on-hand")]

	if op == "delete":
		delta = -_effect(old)                      # +old.IN_QTY (always restores — F2, no purge case)
		if delta == 0:
			return []
		return [_post(_partId(old), delta, _key(old, "del"), old, "reject: un-reject restores on-hand")]

	if op == "update":
		if _partId(old) == _partId(new):
			delta = _effect(new) - _effect(old)    # old.IN_QTY - new.IN_QTY
			if delta == 0:
				return []
			# Amend key carries the TARGET effect (:to=) plus the version stamp -> collision-safe even
			# at the stamp's hundredth-second resolution (two posting amends of one row can't share
			# (stamp, target): a same-target second amend has delta 0 and posts nothing). Same fix as
			# the stocktaking twin; see IGNITION-stock-ledger-design §4.
			return [_post(_partId(new), delta, _key(new, "upd:to=%d:v=%s" % (_effect(new), _ver(new))),
			              new, "reject: amend qty net delta")]
		# part-id change (domain-forbidden; faithful two-post via the independent-join trigger):
		posts = []
		oldRestore = -_effect(old)                 # +old.IN_QTY back to the old part
		if oldRestore != 0:
			posts.append(_post(_partId(old), oldRestore, _key(new, "upd:del-old:v=%s" % _ver(new)),
			                   old, "reject amend part-change: restore old part"))
		newRemove = _effect(new)                   # -new.IN_QTY from the new part
		if newRemove != 0:
			posts.append(_post(_partId(new), newRemove, _key(new, "upd:add-new:v=%s" % _ver(new)),
			                   new, "reject amend part-change: remove from new part"))
		return posts

	raise ValueError("reject.computePosts: unknown op %r" % (op,))


# ---------------------------------------------------------------------------------------------
# Gateway driver — post through the stockLedger funnel (NO part resolution; the row holds IN_PART_ID).
# ---------------------------------------------------------------------------------------------
def postRejectEvent(op, old=None, new=None, site=1, database=None):
	"""Wire one INV_REJECT_INF insert/update/delete into the stock ledger. Call AFTER the reject row
	is written. No purge flag: a reject delete is always a genuine user un-reject (design F2)."""
	db = database if database is not None else stockLedger.DATABASE
	for p in computePosts(op, old, new):
		stockLedger.post(p["partId"], p["delta"], p["sourceEnum"], p["eventKey"],
		                 sourceRowId=p["sourceRowId"], reason=p["reason"],
		                 site=site, purge=False, database=db)


# =============================================================================================
# WRITE-THEN-POST — the reject screen's amend/write seam (the Stage-3 reimplement), twin of the
# stocktaking seam. Each function writes the INV_REJECT_INF source row, re-reads it, then posts the
# matching ledger movement via postRejectEvent. CUTOVER write path: owns IN_QTY via the ledger, so it
# presumes the three legacy reject triggers are GONE (during parallel run the legacy triggers stay
# live and the ledger is a shadow built by replay — do NOT run this against a trigger-live DB or
# IN_QTY double-counts). Source-write/ledger-post ATOMICITY is a flagged cutover TODO (combine into a
# proc or thread a txId through stockLedger.post — escalate to the architect). VC_LAST_UPDATE is
# written fresh per edit (the amend version token).

# The live 16-char yyyymmddHHMMSSff stamp recipe (matches the legacy reject triggers / INSERT procs).
_STAMP_SQL = ("CONVERT(varchar, getdate(), 112) + SUBSTRING(CONVERT(varchar, getdate(), 114), 1, 2) "
              "+ SUBSTRING(CONVERT(varchar, getdate(), 114), 4, 2) "
              "+ SUBSTRING(CONVERT(varchar, getdate(), 114), 7, 2) "
              "+ SUBSTRING(CONVERT(varchar, getdate(), 114), 10, 2)")


def _readRow(rejId, db):
	"""Re-read a reject row as a plain dict (the source of truth the post is built from)."""
	rows = system.db.runPrepQuery(
		"SELECT IN_REJECT_ID, VC_DIVISION, IN_PART_ID, IN_QTY, VC_REASON, VC_LAST_UPDATE, VC_ADD "
		"FROM INV_REJECT_INF WHERE IN_REJECT_ID = ?", [rejId], db)
	if len(rows) == 0:
		return None
	r = rows[0]
	return {"IN_REJECT_ID": r["IN_REJECT_ID"], "VC_DIVISION": r["VC_DIVISION"],
	        "IN_PART_ID": r["IN_PART_ID"], "IN_QTY": r["IN_QTY"], "VC_REASON": r["VC_REASON"],
	        "VC_LAST_UPDATE": r["VC_LAST_UPDATE"], "VC_ADD": r["VC_ADD"]}


def insertReject(partId, qty, division, reason=None, site=1, database=None):
	"""Record a new reject for partId (qty removed from on-hand) and post -qty to the ledger.
	Returns the new IN_REJECT_ID."""
	db = database if database is not None else stockLedger.DATABASE
	newId = system.db.runPrepUpdate(
		"INSERT INTO INV_REJECT_INF (VC_DIVISION, IN_PART_ID, IN_QTY, VC_REASON, VC_LAST_UPDATE, VC_ADD) "
		"VALUES (?, ?, ?, ?, " + _STAMP_SQL + ", " + _STAMP_SQL + ")",
		[division, partId, qty, reason], db, getKey=True)
	postRejectEvent("insert", new=_readRow(newId, db), site=site, database=db)
	return newId


def amendReject(rejId, newQty, newReason=None, site=1, database=None):
	"""Edit an existing reject's qty/reason and post the net delta. Writes a fresh VC_LAST_UPDATE
	(the amend version token)."""
	db = database if database is not None else stockLedger.DATABASE
	old = _readRow(rejId, db)
	if old is None:
		raise ValueError("reject.amendReject: no row IN_REJECT_ID=%r" % (rejId,))
	system.db.runPrepUpdate(
		"UPDATE INV_REJECT_INF SET IN_QTY = ?, VC_REASON = ?, VC_LAST_UPDATE = " + _STAMP_SQL +
		" WHERE IN_REJECT_ID = ?", [newQty, newReason, rejId], db)
	postRejectEvent("update", old=old, new=_readRow(rejId, db), site=site, database=db)


def deleteReject(rejId, site=1, database=None):
	"""Un-reject (remove the reject row) and post the restoring +qty."""
	db = database if database is not None else stockLedger.DATABASE
	old = _readRow(rejId, db)
	if old is None:
		return
	system.db.runPrepUpdate("DELETE FROM INV_REJECT_INF WHERE IN_REJECT_ID = ?", [rejId], db)
	postRejectEvent("delete", old=old, site=site, database=db)
