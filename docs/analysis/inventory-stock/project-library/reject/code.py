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
			return [_post(_partId(new), delta, _key(new, "upd:v=%s" % _ver(new)), new,
			              "reject: amend qty net delta")]
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
