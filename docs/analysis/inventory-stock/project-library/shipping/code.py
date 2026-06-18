# shipping — Project Library service that re-homes the three *PartShipping qty triggers
# (InsertPartShipping / UpdatePartShipping / DeletePartShipping) and the DeleteShipDate header
# cascade into stockLedger.post() calls.
#
# Design: docs/analysis/inventory-stock/IGNITION-stock-ledger-design.md §3 rows #15-#17;
#         source spec: docs/analysis/shipping/shipping.md §2 (the four trigger bodies), §4, §6.
#
# Shipping is the stock-OUT mirror of receiving: production CONSUMES on-hand. The two are the same
# additive-delta shape, with two differences from receiving:
#   * NO add-point gate — shipping ALWAYS moves stock (the *PartShipping triggers never read
#     VC_INVENTORY_ADD_POINT). So effect(row) = -IN_QTY unconditionally (negative = consumption).
#   * a header-delete CASCADE (DeleteShipDate) that the rebuild scopes to the deleted header's own
#     rows (IN_SHIPPING_ID) — fixing the legacy line-blind `WHERE VC_PRODUCTION_DATE IN (...)` cascade
#     (shipping.md §2) — and restores EVERY row (multi-row-safe; legacy's `FROM PS, deleted` cross-join
#     under-restored a multi-row same-part cascade, the F3 hazard, design §6 class 5).
#
# Verified against the LIVE trigger bodies (CreateInventory.sql):
#   InsertPartShipping : IN_QTY = PS.IN_QTY - i.IN_QTY                          -> delta = -new.IN_QTY
#   DeletePartShipping : IN_QTY = PS.IN_QTY + d.IN_QTY                          -> delta = +old.IN_QTY
#   UpdatePartShipping : +d.IN_QTY (restore old, joined on d) THEN -i.IN_QTY (subtract new, joined on i)
#       Same-part edit  -> net delta = old.IN_QTY - new.IN_QTY.
#       Part-number edit-> the two legs join PS INDEPENDENTLY on d/i part numbers, so legacy ALREADY
#       restores the old part and subtracts the new part correctly. The two-post below therefore
#       MATCHES legacy (parity) — it is NOT a "ledger correct, legacy buggy" divergence (contrast the
#       receiving F5, where the legs cross-joined and legacy silently no-op'd a part change).
#
# Jython 2.7 (8.1 and 8.3). The pure logic (computePosts / computeHeaderDeletePosts) imports nothing
# and is CPython-importable for unit test (scripts/e2e/test_shipping_posts.py). Only the drivers touch
# the gateway globals `system` and `stockLedger`.
#
# IG81-COMPAT: createSProcCall + runPrepQuery behave identically on 8.1.52 and 8.3.
# IG83-TODO:  at the Postgres phase, scope the cascade via a real IN_SHIPPING_ID FK; normalize VC_ADD.

# VC_SOURCE_ENUM class (design §3). Local mirror of stockLedger.ENUM_SHIPPING (the wire contract);
# the unit test asserts equality so the two can never drift silently.
ENUM_SHIPPING = "SHIPPING"


def _effect(row):
	"""The signed on-hand contribution of a part-shipping row: ALWAYS -IN_QTY (consumption at
	production; no add-point gate). The whole §3 mapping reduces to delta = effect(new) - effect(old)."""
	return -int(row["IN_QTY"])


def _partShipId(row):
	v = row.get("IN_PART_SHIPPING_ID")
	return None if v is None or str(v).strip() == "" else int(v)


def _ver(row):
	"""Per-edit version discriminator for amend (update) event keys. INV_PART_SHIPPING_INF carries only
	VC_ADD; the rebuild's AmendShipment write path writes a FRESH 16-char VC_ADD per edit (design §6
	timestamp normalization — the legacy UPDATE_Shippingdetail did NOT re-stamp it, but VC_ADD is
	audit-only so stock parity is unaffected). That makes successive amends distinct events while a
	replay of the same amend reuses the key (idempotent). Insert/delete keys need no version — an
	IN_PART_SHIPPING_ID is inserted/deleted exactly once per lifetime (identity-stable)."""
	return str(row.get("VC_ADD") or "")


def _key(row, suffix):
	"""'SHIPPING:psh={IN_PART_SHIPPING_ID}:{suffix}' — e.g. SHIPPING:psh=12077:ins (design §4)."""
	return "%s:psh=%s:%s" % (ENUM_SHIPPING, _partShipId(row), suffix)


def _post(partNumber, delta, eventKey, row, reason):
	"""One post-instruction (pure data — the driver resolves partNumber->IN_PART_ID and calls
	stockLedger.post). sourceRowId is the part-shipping surrogate IN_PART_SHIPPING_ID."""
	return {
		"partNumber": partNumber,
		"delta": int(delta),
		"sourceEnum": ENUM_SHIPPING,
		"eventKey": eventKey,
		"sourceRowId": _partShipId(row),
		"reason": reason,
	}


# ---------------------------------------------------------------------------------------------
# computePosts — the PURE §3 #15-#17 mapping. Part-shipping row event -> post-instructions. No DB.
# ---------------------------------------------------------------------------------------------
def computePosts(op, old=None, new=None):
	"""Translate one INV_PART_SHIPPING_INF event into the signed stock movements to post.

	op        : 'insert' | 'update' | 'delete'
	old / new : the before / after part-shipping rows as dicts (keys = INV_PART_SHIPPING_INF columns:
	            VC_PART_NUMBER, IN_QTY, IN_PART_SHIPPING_ID, IN_SHIPPING_ID, VC_ADD).
	            'insert' uses new (the final accumulated row — a build post may accumulate several
	            DoPartNumberInventory calls into one row via INSERT_ShippingPartInfo's upsert; the
	            ledger sees the committed row once). 'delete' uses old. 'update' uses both.

	Returns a list of post dicts (0, 1, or 2 entries — two when an amend re-points VC_PART_NUMBER).
	"""
	if op == "insert":
		# §3 #15 — a part-shipping row consumes its qty. effect(new) - 0 = -new.IN_QTY.
		delta = _effect(new)
		if delta == 0:
			return []
		return [_post(new["VC_PART_NUMBER"], delta, _key(new, "ins"), new,
		              "shipping: consume at production")]

	if op == "delete":
		# §3 #17 — removing a part-shipping row restores its qty. 0 - effect(old) = +old.IN_QTY.
		delta = -_effect(old)
		if delta == 0:
			return []
		return [_post(old["VC_PART_NUMBER"], delta, _key(old, "del"), old,
		              "shipping: restore on line/shipment delete")]

	if op == "update":
		oldPart = old["VC_PART_NUMBER"]
		newPart = new["VC_PART_NUMBER"]
		if oldPart == newPart:
			# §3 #16 same-part amend -> one net-delta post: effect(new) - effect(old) = old - new.
			delta = _effect(new) - _effect(old)
			if delta == 0:
				return []
			return [_post(newPart, delta, _key(new, "upd:v=%s" % _ver(new)), new,
			              "shipping: amend qty net delta")]

		# Part-number amend -> two posts (restore old part, consume new part). Verified to MATCH the
		# live UpdatePartShipping (independent d/i joins) — parity, not a divergence (see header note).
		posts = []
		oldRestore = -_effect(old)   # +old.IN_QTY
		if oldRestore != 0:
			posts.append(_post(oldPart, oldRestore, _key(new, "upd:del-old:v=%s" % _ver(new)), old,
			                   "shipping amend part-change: restore old part"))
		newConsume = _effect(new)    # -new.IN_QTY
		if newConsume != 0:
			posts.append(_post(newPart, newConsume, _key(new, "upd:add-new:v=%s" % _ver(new)), new,
			                   "shipping amend part-change: consume new part"))
		return posts

	raise ValueError("shipping.computePosts: unknown op %r" % (op,))


def computeHeaderDeletePosts(rows):
	"""The DeleteShipDate header-cascade as posts: restore EVERY part-shipping row of the deleted
	shipment (multi-row-safe — fixes the legacy `FROM PS, deleted` single-assignment under-restore,
	F3). `rows` MUST already be scoped to the deleted header's IN_SHIPPING_ID (the driver does that),
	which fixes the legacy line-blind `WHERE VC_PRODUCTION_DATE IN (...)` cascade (shipping.md §2).
	Pure: one '+qty' delete post per row."""
	posts = []
	for r in rows:
		posts.extend(computePosts("delete", old=r))
	return posts


# ---------------------------------------------------------------------------------------------
# Gateway drivers — resolve part id, post through the stockLedger funnel.
# (Touch the gateway globals `system` and `stockLedger`; not exercised by the pure unit test.)
# ---------------------------------------------------------------------------------------------
def _postAll(posts, site, database):
	db = database if database is not None else stockLedger.DATABASE
	for p in posts:
		partId = stockLedger.resolvePartId(p["partNumber"], db)
		if partId is None:
			# No master row for this part number -> legacy join matches nothing, no stock moves. Faithful.
			continue
		stockLedger.post(partId, p["delta"], p["sourceEnum"], p["eventKey"],
		                 sourceRowId=p["sourceRowId"], reason=p["reason"],
		                 site=site, purge=False, database=db)


def postShippingEvent(op, old=None, new=None, site=1, database=None):
	"""Wire one INV_PART_SHIPPING_INF insert/update/delete into the stock ledger. Call AFTER the row
	has been written by the wrapped INSERT_ShippingPartInfo / UPDATE_Shippingdetail / delete path
	(for a build post, call once per committed part-shipping row)."""
	_postAll(computePosts(op, old, new), site, database)


def postHeaderDelete(shippingId, site=1, database=None):
	"""Wire a DeleteShipDate header delete into the ledger: restore every part-shipping row belonging
	to THIS shipment header (scoped to IN_SHIPPING_ID -> (line, production_date), fixing the legacy
	line-blind cascade), each as a '+qty' post (multi-row-safe). Call BEFORE the cascade physically
	deletes the rows (it reads them to know the qtys to restore)."""
	db = database if database is not None else stockLedger.DATABASE
	rows = system.db.runPrepQuery(
		"SELECT IN_PART_SHIPPING_ID, IN_SHIPPING_ID, VC_PART_NUMBER, IN_QTY, VC_ADD "
		"FROM INV_PART_SHIPPING_INF WHERE IN_SHIPPING_ID = ?",
		[shippingId], db)
	# runPrepQuery returns a dataset; normalize each row to a plain dict for the pure mapping.
	plain = [{"IN_PART_SHIPPING_ID": r["IN_PART_SHIPPING_ID"], "IN_SHIPPING_ID": r["IN_SHIPPING_ID"],
	          "VC_PART_NUMBER": r["VC_PART_NUMBER"], "IN_QTY": r["IN_QTY"], "VC_ADD": r["VC_ADD"]}
	         for r in rows]
	_postAll(computeHeaderDeletePosts(plain), site, database)
