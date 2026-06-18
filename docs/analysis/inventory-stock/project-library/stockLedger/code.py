# stockLedger — Project Library service that OWNS INV_PARTS_STOCK_MST.IN_QTY.
# Design: docs/analysis/inventory-stock/IGNITION-stock-ledger-design.md §5.
#
# The single funnel every stock module (RecConfStat / Shipping / Reject / Stocktaking)
# calls. Screens NEVER touch IN_QTY directly; they call stockLedger.post(). post() wraps
# the DB-side POST_StockMovement (atomic insert-ledger + additive IN_QTY += delta, idempotent
# on (partId, eventKey)). rebuildBalance() wraps PROC_RebuildStockBalance (the F4-fenced
# read-then-write absolute re-stamp; takes an exclusive lock — never co-runs with post()).
#
# Jython 2.7 (8.1 and 8.3 — no scripting delta).
# IG81-COMPAT: createSProcCall + a stored proc behaves identically on 8.1.52 and 8.3.
# IG83-TODO: at the Postgres phase TS_POSTED -> datetime2; site_id FK enforced; consider a
#            computed-view projection of SUM(qty_delta) instead of the materialized column.

DATABASE = "Inventory_Spike"

# VC_SOURCE_ENUM movement classes (design §3).
ENUM_RECEIVING_SHIP     = "RECEIVING_SHIP"
ENUM_RECEIVING_ARRIVAL  = "RECEIVING_ARRIVAL"
ENUM_RECEIVING_REVERSAL = "RECEIVING_REVERSAL"
ENUM_REJECT             = "REJECT"
ENUM_STOCKTAKING        = "STOCKTAKING"
ENUM_SHIPPING           = "SHIPPING"
ENUM_OPENING_BALANCE    = "OPENING_BALANCE"   # cutover backfill seed (thread b, option 2)


def resolvePartId(partNumber, database=DATABASE):
	"""Resolve the legacy string VC_PART_NUMBER -> surrogate IN_PART_ID at the posting
	boundary (D2). Receiving/shipping callers hold the string; reject/stocktaking already
	hold the id. Returns the int id, or None if the part number is unknown."""
	rows = system.db.runPrepQuery(
		"SELECT IN_PART_ID FROM INV_PARTS_STOCK_MST WHERE VC_PART_NUMBER = ?",
		[partNumber], database)
	if len(rows) == 0:
		return None
	return rows[0]["IN_PART_ID"]


def post(partId, delta, sourceEnum, eventKey, sourceRowId=None,
         reason=None, site=1, purge=False, database=DATABASE):
	"""Post ONE signed stock movement. The only writer of IN_QTY.

	Atomic + idempotent (design §4): the DB proc appends a ledger row and bumps
	IN_QTY += delta in one transaction; a replay of the same (partId, eventKey) is a
	no-op (UNIQUE backstop). purge=True posts nothing (§3.1 housekeeping/terminated delete).

	partId      : int  surrogate IN_PART_ID (resolve a string via resolvePartId first)
	delta       : int  signed change (+ adds on-hand, - removes)
	sourceEnum  : str  one of the ENUM_* classes
	eventKey    : str  stable idempotency discriminator, e.g. 'SHIPPING:psh=12077:ins'
	sourceRowId : int  source-row surrogate id (order/reject/stk/shipping), optional
	reason      : str  coded/free-text reason, optional
	site        : int  D1 site scope (default 1 for parallel run)
	purge       : bool True => post nothing (delete-the-record / already-terminated)
	"""
	call = system.db.createSProcCall("POST_StockMovement", database)
	call.registerInParam("partId",      system.db.INTEGER, partId)
	call.registerInParam("delta",       system.db.INTEGER, delta)
	call.registerInParam("sourceEnum",  system.db.VARCHAR, sourceEnum)
	call.registerInParam("sourceRowId", system.db.INTEGER, sourceRowId)
	call.registerInParam("eventKey",    system.db.VARCHAR, eventKey)
	call.registerInParam("reason",      system.db.VARCHAR, reason)
	call.registerInParam("site",        system.db.INTEGER, site)
	call.registerInParam("purge",       system.db.BIT,     1 if purge else 0)
	system.db.execSProcCall(call)


def rebuildBalance(partId, database=DATABASE):
	"""Healing command (design §7 / F4): re-SUM the ledger and re-stamp IN_QTY = <absolute SUM>
	for one part. This is the ONE read-then-write absolute path; the wrapped proc takes an
	exclusive (UPDLOCK,HOLDLOCK) lock under SERIALIZABLE so a concurrent post() delta is never
	lost. Use to heal a drift the reconciliation harness finds — NOT in the hot path."""
	call = system.db.createSProcCall("PROC_RebuildStockBalance", database)
	call.registerInParam("partId", system.db.INTEGER, partId)
	system.db.execSProcCall(call)


def seedOpeningBalance(partId, site=1, database=DATABASE):
	"""CUTOVER backfill (thread b / option 2): record this part's CURRENT IN_QTY as a single
	OPENING_BALANCE ledger movement WITHOUT bumping IN_QTY (the materialized column already holds the
	cutover value). After seeding, SUM(qty_delta) == IN_QTY by construction; forward post()s keep the
	invariant. Idempotent on (partId, 'OPENING:part=<id>') — safe to re-run. The from-zero replay is
	impossible on the purged data (receiving IN-moves aged out, F2), so this opening balance + forward
	parity IS the cutover sign-off, not a full-history reconstruction."""
	call = system.db.createSProcCall("SEED_OpeningBalance", database)
	call.registerInParam("partId", system.db.INTEGER, partId)
	call.registerInParam("site",   system.db.INTEGER, site)
	system.db.execSProcCall(call)


def seedAllOpeningBalances(site=1, database=DATABASE):
	"""Seed an opening balance for EVERY part (the one-time cutover backfill) via the SET-BASED
	SEED_AllOpeningBalances proc — one statement / one transaction, so it scales to thousands of parts
	(not a per-part round-trip loop). Genesis-only + idempotent: only parts with no ledger row yet are
	seeded, so a re-run is safe. This is the exact path the validator exercises."""
	call = system.db.createSProcCall("SEED_AllOpeningBalances", database)
	call.registerInParam("site", system.db.INTEGER, site)
	system.db.execSProcCall(call)
