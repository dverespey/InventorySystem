# receiving — Project Library service that re-homes the three RecConfStat qty triggers
# (INSERT_/UPDATE_/DELETE_RecConfStatPartsStockMstQTY) into stockLedger.post() calls.
#
# Design: docs/analysis/inventory-stock/IGNITION-stock-ledger-design.md §3 rows #1-#8;
#         source spec: docs/analysis/receiving/recconfstat.md §2-§4 (D7 / D8(3) / D12#3 / F5).
#
# The RecConfStat screen NEVER touches IN_QTY. It writes INV_OPEN_ORDER_INF (via the wrapped
# INSERT/UPDATE/UPDATE..Renban/DELETE procs) and then hands the before/after order row(s) to
# this service, which evaluates the add-point GATE (in the service, not the DB — design §3) and
# posts the signed stock movement through the single stockLedger funnel.
#
# Jython 2.7 (8.1 and 8.3 — no scripting delta). The pure decision logic (computePosts and its
# helpers) imports nothing and is CPython-importable for unit test (scripts/e2e/test_receiving_posts.py).
# Only postOrderEvent / resolveAddPoint touch the gateway globals `system` and `stockLedger`.
#
# IG81-COMPAT: createSProcCall + runPrepQuery behave identically on 8.1.52 and 8.3.
# IG83-TODO:  at the Postgres phase, add-point resolution stays a part->supplier join; site scoping
#             becomes an enforced FK (D1). No version delta here.

# VC_SOURCE_ENUM movement classes for the receiving paths (design §3). Kept as local mirrors of
# stockLedger.ENUM_* so this module's PURE logic carries no import; they ARE the wire contract
# (the VC_SOURCE_ENUM strings written to INV_STOCK_LEDGER) and must equal stockLedger's — the unit
# test asserts that equality so the two can never drift silently.
ENUM_RECEIVING_SHIP     = "RECEIVING_SHIP"
ENUM_RECEIVING_ARRIVAL  = "RECEIVING_ARRIVAL"
ENUM_RECEIVING_REVERSAL = "RECEIVING_REVERSAL"


# ---------------------------------------------------------------------------------------------
# Status helpers — what makes an open-order row "counted" toward on-hand for each add-point.
# ---------------------------------------------------------------------------------------------
def _filled(v):
	"""A date/status stamp is 'set' when it is a non-blank string (legacy uses <> '')."""
	return v is not None and str(v).strip() != ""


def _shipped(row):
	"""'S' add-point counts at supplier-shipping: VC_STATUS_SUPPLIER_SHIPPING set (recconfstat §2)."""
	return _filled(row.get("VC_STATUS_SUPPLIER_SHIPPING"))


def _arrived(row):
	"""'A' add-point counts at arrival. D7 + D12#3: arrival is set when ANY of arrival / plant-yard /
	assembler-yard / warehouse is stamped. (Legacy's INSERT/DELETE triggers already OR the four; the
	UPDATE arrival-leg gated only on VC_ARRIVAL and so UNDER-counted yard-on-edit — D12#3 fixes that
	here by treating the yards as arrival-equivalent on every op, an intentional 'ledger correct,
	legacy buggy' divergence the parity harness predicts.)"""
	return (_filled(row.get("VC_ARRIVAL"))
	        or _filled(row.get("VC_STATUS_PLANT_YARD"))
	        or _filled(row.get("VC_STATUS_ASSEMBLER_YARD"))
	        or _filled(row.get("VC_WAREHOUSE")))


def _counted(row, addPoint):
	"""Is this row contributing its order qty to on-hand, given the supplier's add-point?
	Any add-point other than 'S'/'A' (or an unresolved supplier) moves NO stock — faithful to the
	legacy INNER JOIN through the part's supplier dropping such parts (recconfstat §2)."""
	if addPoint == "S":
		return _shipped(row)
	if addPoint == "A":
		return _arrived(row)
	return False


def _effect(row, addPoint):
	"""The signed on-hand contribution of a single order row: its order qty when counted, else 0.
	The whole §3 mapping reduces to delta = effect(new) - effect(old)."""
	return int(row["IN_QTY"]) if _counted(row, addPoint) else 0


def _arrivalAddPoint(addPoint):
	return ENUM_RECEIVING_SHIP if addPoint == "S" else ENUM_RECEIVING_ARRIVAL


def _updateEnum(addPoint, old, new):
	"""Classify the movement class of a same-part edit. 'S' is always SHIP. For 'A', an arrival that
	was set and is now cleared is the D8(3) REVERSAL (dead in the legacy trigger; implemented here);
	everything else (arrival set, or a qty change while still arrived) is ARRIVAL."""
	if addPoint == "S":
		return ENUM_RECEIVING_SHIP
	if _arrived(old) and not _arrived(new):
		return ENUM_RECEIVING_REVERSAL
	return ENUM_RECEIVING_ARRIVAL


# ---------------------------------------------------------------------------------------------
# Event-key assembly (design §4: (IN_PART_ID, VC_SOURCE_EVENT) is the idempotency key).
# ---------------------------------------------------------------------------------------------
def _orderId(row):
	v = row.get("IN_ORDER_ID")
	return None if v is None or str(v).strip() == "" else int(v)


def _ver(row):
	"""Per-edit version discriminator so successive edits of the same order are distinct events
	(but a replay of the SAME edit is idempotent). VC_LAST_UPDATE is the 16-char yyyymmddHHMMSSff
	stamp the UPDATE proc writes — monotonic and unique per edit."""
	return str(row.get("VC_LAST_UPDATE") or "")


def _key(enum, row, suffix):
	"""'{ENUM}:ord={IN_ORDER_ID}:{suffix}' — e.g. RECEIVING_ARRIVAL:ord=8842:upd:v=2026...
	Insert ('ins') and delete ('del') keys carry NO version token; that is safe because
	IN_ORDER_ID is IDENTITY (recconfstat §2) — an order is inserted/deleted exactly once per id,
	so a replay reuses the same key (idempotent). Only EDITS need the per-edit version (a row is
	editable many times), so 'upd' suffixes carry :v=<VC_LAST_UPDATE>."""
	return "%s:ord=%s:%s" % (enum, _orderId(row), suffix)


def _post(partNumber, delta, sourceEnum, eventKey, row, reason):
	"""One post-instruction (pure data — the driver resolves partNumber->IN_PART_ID and calls
	stockLedger.post). sourceRowId is the open-order surrogate IN_ORDER_ID."""
	return {
		"partNumber": partNumber,
		"delta": int(delta),
		"sourceEnum": sourceEnum,
		"eventKey": eventKey,
		"sourceRowId": _orderId(row),
		"reason": reason,
	}


# ---------------------------------------------------------------------------------------------
# computePosts — the PURE §3 mapping. Open-order event -> list of post-instructions. No DB.
# ---------------------------------------------------------------------------------------------
def computePosts(op, old=None, new=None, addPoint=None, addPointOld=None, addPointNew=None):
	"""Translate one INV_OPEN_ORDER_INF event into the signed stock movements to post.

	op           : 'insert' | 'update' | 'delete'
	old / new    : the before / after order rows as dicts (keys = INV_OPEN_ORDER_INF columns:
	               VC_PART_NUMBER, IN_QTY, VC_STATUS_SUPPLIER_SHIPPING, VC_ARRIVAL,
	               VC_STATUS_PLANT_YARD, VC_STATUS_ASSEMBLER_YARD, VC_WAREHOUSE,
	               VC_TERMINATED, VC_STATUS_EMPTY_TRAILER, VC_LAST_UPDATE, IN_ORDER_ID).
	               'insert' uses new; 'delete' uses old; 'update' uses both.
	addPoint     : the supplier add-point ('S'/'A'/other) for insert/delete and a same-part update.
	addPointOld/addPointNew : add-points of the old/new part for a part-RE-POINTING update (F5);
	               each part carries its own supplier and hence its own add-point.

	Returns a list of post dicts (0, 1, or 2 entries). Purge / already-terminated suppression is
	the driver's job for purge (§3.1); the DELETE branch here applies the live VC_TERMINATED='' /
	VC_STATUS_EMPTY_TRAILER='' gate.
	"""
	if op == "insert":
		# §3 #1/#2 — inserting an already-shipped('S') / already-arrived('A') order adds its qty.
		# Faithful to the legacy INSERT trigger, which has NO terminated/empty-trailer gate (the
		# INSERT vs DELETE gate asymmetry, recconfstat §8 open-Q3 — left faithful; flagged there).
		eff = _effect(new, addPoint)
		if eff == 0:
			return []
		enum = _arrivalAddPoint(addPoint)
		return [_post(new["VC_PART_NUMBER"], eff, enum, _key(enum, new, "ins"), new,
		              "receiving insert: %s-counted order added" % addPoint)]

	if op == "delete":
		# §3 #8 — remove the qty of a still-active counted order. Live gate: VC_TERMINATED='' AND
		# VC_STATUS_EMPTY_TRAILER='' (terminated/empty-trailered rows post nothing — that is also
		# how DELETE_AutoPurge spares on-hand, §3.1: it pre-stamps VC_TERMINATED before deleting).
		if _filled(old.get("VC_TERMINATED")) or _filled(old.get("VC_STATUS_EMPTY_TRAILER")):
			return []
		eff = _effect(old, addPoint)
		if eff == 0:
			return []
		enum = _arrivalAddPoint(addPoint)
		return [_post(old["VC_PART_NUMBER"], -eff, enum, _key(enum, old, "del"), old,
		              "receiving delete: reverse %s-counted order" % addPoint)]

	if op == "update":
		oldPart = old["VC_PART_NUMBER"]
		newPart = new["VC_PART_NUMBER"]

		if oldPart == newPart:
			# §3 #3-#7 same-part edit collapses to ONE net-delta post (design §4 point 3):
			#   delta = effect(new) - effect(old)
			# captures qty-change, ship/arrival set, ship/arrival clear (incl. the D8(3) reversal).
			delta = _effect(new, addPoint) - _effect(old, addPoint)
			if delta == 0:
				return []
			enum = _updateEnum(addPoint, old, new)
			# Amend key carries the TARGET effect (:to=) + version stamp -> collision-safe at the stamp's
			# hundredth-second resolution (two posting amends of one order can't share (stamp, target): a
			# same-target second amend has delta 0 and posts nothing). Twin of stocktaking/reject; §4.
			return [_post(newPart, delta, enum,
			              _key(enum, new, "upd:to=%d:v=%s" % (_effect(new, addPoint), _ver(new))), new,
			              "receiving edit: net on-hand delta")]

		# F5 — the edit RE-POINTS the part number (UPDATE_RecConfStatInfo can rewrite the key).
		# i/d resolve to DIFFERENT part ids, so the single-net-delta assumption breaks: emit the
		# genuine pair (remove the old part's effect, add the new part's effect). Each part uses its
		# own supplier add-point. (Legacy's self-join on VC_PART_NUMBER silently no-ops this — the
		# two-post is the correct 'ledger correct, legacy buggy (F5)' divergence the harness predicts.)
		apOld = addPointOld if addPointOld is not None else addPoint
		apNew = addPointNew if addPointNew is not None else addPoint
		posts = []
		oldEff = _effect(old, apOld)
		if oldEff != 0:
			enumO = _arrivalAddPoint(apOld)
			posts.append(_post(oldPart, -oldEff, enumO,
			                   _key(enumO, old, "upd:del-old:v=%s" % _ver(new)), old,
			                   "receiving edit part-change (F5): remove from old part"))
		newEff = _effect(new, apNew)
		if newEff != 0:
			enumN = _arrivalAddPoint(apNew)
			posts.append(_post(newPart, newEff, enumN,
			                   _key(enumN, new, "upd:add-new:v=%s" % _ver(new)), new,
			                   "receiving edit part-change (F5): add to new part"))
		return posts

	raise ValueError("receiving.computePosts: unknown op %r" % (op,))


# ---------------------------------------------------------------------------------------------
# Gateway driver — resolve add-point + part id, post through the stockLedger funnel.
# (Touches the gateway globals `system` and `stockLedger`; not exercised by the pure unit test.)
# ---------------------------------------------------------------------------------------------
def resolveAddPoint(partNumber, database=None):
	"""Resolve a part's supplier add-point the SAME way the legacy qty triggers do — through the
	part's IN_SUPPLIER_ID, NOT the order's VC_SUPPLIER_CODE string (recconfstat §2). Returns 'S' /
	'A' / other, or None when the part has no resolvable supplier (-> moves no stock, faithful)."""
	db = database if database is not None else stockLedger.DATABASE
	rows = system.db.runPrepQuery(
		"SELECT s.VC_INVENTORY_ADD_POINT "
		"FROM INV_PARTS_STOCK_MST ps "
		"JOIN INV_SUPPLIER_MST s ON ps.IN_SUPPLIER_ID = s.IN_SUPPLIER_ID "
		"WHERE ps.VC_PART_NUMBER = ?",
		[partNumber], db)
	if len(rows) == 0:
		return None
	return rows[0]["VC_INVENTORY_ADD_POINT"]


def postOrderEvent(op, old=None, new=None, site=1, purge=False, database=None):
	"""Wire an INV_OPEN_ORDER_INF insert/update/delete into the stock ledger. Call AFTER the
	open-order row has been written by the wrapped INSERT/UPDATE/DELETE proc.

	purge=True (a DELETE_AutoPurge housekeeping delete) posts NOTHING (§3.1) — the historical
	receipt already happened; deleting the aged record must not un-happen it.
	"""
	if purge:
		return
	db = database if database is not None else stockLedger.DATABASE

	if op == "update" and old["VC_PART_NUMBER"] != new["VC_PART_NUMBER"]:
		posts = computePosts(op, old, new,
		                     addPointOld=resolveAddPoint(old["VC_PART_NUMBER"], db),
		                     addPointNew=resolveAddPoint(new["VC_PART_NUMBER"], db))
	else:
		keyRow = old if op == "delete" else new
		posts = computePosts(op, old, new,
		                     addPoint=resolveAddPoint(keyRow["VC_PART_NUMBER"], db))

	for p in posts:
		partId = stockLedger.resolvePartId(p["partNumber"], db)
		if partId is None:
			# No master row for this part number -> legacy join drops it, no stock moves. Faithful.
			continue
		stockLedger.post(partId, p["delta"], p["sourceEnum"], p["eventKey"],
		                 sourceRowId=p["sourceRowId"], reason=p["reason"],
		                 site=site, purge=False, database=db)


# =============================================================================================
# WRITE-THEN-POST — the RecConfStat screen's amend/write seam (the Stage-3 reimplement). Writes the
# INV_OPEN_ORDER_INF row, re-reads it, then posts via postOrderEvent. CUTOVER write path: owns IN_QTY
# via the ledger, presumes the three legacy RecConfStat qty triggers are GONE (parallel run keeps the
# triggers live + the ledger a shadow — do NOT run this against a trigger-live DB or IN_QTY
# double-counts). Source-write/ledger-post ATOMICITY is a flagged cutover TODO (combine into a proc or
# thread a txId — escalate to the architect).
#
# Stage-3 fixes baked in by keying writes on the SURROGATE IN_ORDER_ID (not the editable natural
# 4-tuple): this sidesteps BOTH the supplier-blind DELETE_RecConfStatInfo bug (recconfstat §4 — it
# omitted VC_SUPPLIER_CODE) AND the prev-tuple UPDATE_RecConfStatInfo complexity. VC_LAST_UPDATE is
# re-stamped fresh per write (the amend version token). VC_FRS_DATE (the only nullable column) is left
# to the screen's derivation step and not touched here.
#
# SCOPE: this is the stock-path seam. The full screen (FRS-date year-rollover derivation, the RENBAN
# batch update, server-side search) is the dedicated RecConfStat module build, not this seam.

# The live 16-char yyyymmddHHMMSSff stamp recipe (matches the legacy RecConfStat procs/triggers).
_STAMP_SQL = ("CONVERT(varchar, getdate(), 112) + SUBSTRING(CONVERT(varchar, getdate(), 114), 1, 2) "
              "+ SUBSTRING(CONVERT(varchar, getdate(), 114), 4, 2) "
              "+ SUBSTRING(CONVERT(varchar, getdate(), 114), 7, 2) "
              "+ SUBSTRING(CONVERT(varchar, getdate(), 114), 10, 2)")

_OO_VARCHARS = ("VC_STATUS_SUPPLIER_SHIPPING", "VC_ARRIVAL", "VC_TRAILER_NUMBER", "VC_STATUS_PLANT_YARD",
                "VC_PLANT_PARKING", "VC_STATUS_ASSEMBLER_YARD", "VC_ASSEMBLER_LOCATION",
                "VC_STATUS_EMPTY_TRAILER", "VC_DETENTION", "VC_ORDER_DATE", "VC_WAREHOUSE",
                "VC_TERMINATED", "VC_SHIP_DATE", "VC_KANBAN_NUMBER")  # NOT-NULL varchars (default '')


def _readOrder(orderId, db):
	"""Re-read the stock-relevant columns of an open-order row as a plain dict (what computePosts reads)."""
	rows = system.db.runPrepQuery(
		"SELECT IN_ORDER_ID, VC_PART_NUMBER, IN_QTY, VC_STATUS_SUPPLIER_SHIPPING, VC_ARRIVAL, "
		"VC_STATUS_PLANT_YARD, VC_STATUS_ASSEMBLER_YARD, VC_WAREHOUSE, VC_TERMINATED, "
		"VC_STATUS_EMPTY_TRAILER, VC_LAST_UPDATE FROM INV_OPEN_ORDER_INF WHERE IN_ORDER_ID = ?",
		[orderId], db)
	if len(rows) == 0:
		return None
	r = rows[0]
	return dict((c, r[c]) for c in
	            ("IN_ORDER_ID", "VC_PART_NUMBER", "IN_QTY", "VC_STATUS_SUPPLIER_SHIPPING", "VC_ARRIVAL",
	             "VC_STATUS_PLANT_YARD", "VC_STATUS_ASSEMBLER_YARD", "VC_WAREHOUSE", "VC_TERMINATED",
	             "VC_STATUS_EMPTY_TRAILER", "VC_LAST_UPDATE"))


def insertOpenOrder(oo, site=1, database=None):
	"""Insert an open order (oo = dict of supplier/part/FRS/renban/qty + any milestone stamps; missing
	NOT-NULL varchars default to '') and post the add-point-gated movement. Returns the new IN_ORDER_ID."""
	db = database if database is not None else stockLedger.DATABASE
	cols = ["VC_SUPPLIER_CODE", "VC_PART_NUMBER", "VC_FRS_NUMBER", "VC_RENBAN_NUMBER", "IN_QTY"] + list(_OO_VARCHARS)
	vals = [oo.get(c, "") for c in cols]
	vals[4] = int(oo["IN_QTY"])
	placeholders = ", ".join(["?"] * len(cols))
	newId = system.db.runPrepUpdate(
		"INSERT INTO INV_OPEN_ORDER_INF (" + ", ".join(cols) + ", VC_LAST_UPDATE, VC_ADD) "
		"VALUES (" + placeholders + ", " + _STAMP_SQL + ", " + _STAMP_SQL + ")",
		vals, db, getKey=True)
	postOrderEvent("insert", new=_readOrder(newId, db), site=site, database=db)
	return newId


def updateOpenOrder(orderId, oo, site=1, database=None):
	"""Edit an open order by surrogate id (oo = the fields to set) and post the net movement. Re-stamps
	VC_LAST_UPDATE. Keying on IN_ORDER_ID fixes the legacy editable-key / prev-tuple hazards."""
	db = database if database is not None else stockLedger.DATABASE
	old = _readOrder(orderId, db)
	if old is None:
		raise ValueError("receiving.updateOpenOrder: no row IN_ORDER_ID=%r" % (orderId,))
	cols = ["VC_SUPPLIER_CODE", "VC_PART_NUMBER", "VC_FRS_NUMBER", "VC_RENBAN_NUMBER", "IN_QTY"] + list(_OO_VARCHARS)
	setters, vals = [], []
	for c in cols:
		if c in oo:
			setters.append("%s = ?" % c)
			vals.append(int(oo[c]) if c == "IN_QTY" else oo[c])
	setters.append("VC_LAST_UPDATE = " + _STAMP_SQL)
	vals.append(orderId)
	system.db.runPrepUpdate("UPDATE INV_OPEN_ORDER_INF SET " + ", ".join(setters) +
	                        " WHERE IN_ORDER_ID = ?", vals, db)
	postOrderEvent("update", old=old, new=_readOrder(orderId, db), site=site, database=db)


def deleteOpenOrder(orderId, site=1, purge=False, database=None):
	"""Delete an open order by surrogate id and post the reversal (gated by VC_TERMINATED=''/empty-
	trailer in computePosts; purge=True posts nothing, §3.1). Keying on IN_ORDER_ID fixes the legacy
	supplier-blind delete."""
	db = database if database is not None else stockLedger.DATABASE
	old = _readOrder(orderId, db)
	if old is None:
		return
	system.db.runPrepUpdate("DELETE FROM INV_OPEN_ORDER_INF WHERE IN_ORDER_ID = ?", [orderId], db)
	postOrderEvent("delete", old=old, site=site, purge=purge, database=db)
