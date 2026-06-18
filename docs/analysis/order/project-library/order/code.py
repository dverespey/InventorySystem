# order — Project Library service for the Order commit/write path (the "Create order(s) from
# worksheet" action, legacy TOrder_Form.ProcessOrder_ButtonClick, Order.pas:628-893).
#
# Source spec: docs/analysis/order/legacy-order-spec.md; domain: project-order-renban-domain memory.
# Turns a worksheet line (a part + a lot count + the computed qty) into INV_OPEN_ORDER_INF order
# records, branching on the lot-sized flag and assigning FRS# + renban.
#
# KEY DOMAIN FACT — the BIT_LOT_SIZE_ORDERS flag is stored INVERTED: 0 = lot-sized TRUE,
# 1 = palletized (David 2026-06-15). The legacy keys its 'LotSizeOrders' string off
# BIT_LOT_SIZE_ORDERS.AsBoolean (Order.pas:548), so the code's 'TRUE' branch is actually the
# PALLETIZED case and 'FALSE' is the LOT-SIZED case. Net behavior (verified against Order.pas:683-849):
#   * LOT-SIZED (bit 0): ONE record PER LOT — loop 1..lotCount, FRS suffix '01','02',… , qty = lotQty,
#     renban = kanban + 3-digit counter (sequential, e.g. 16H006/16H007); the counter bumps per record.
#   * PALLETIZED (bit 1): ONE record total, FRS suffix '01', qty = Q (the whole order), renban BLANK
#     (palletized parts carry a renban group -> renban is assigned later by the RenbanOrder grouping
#     form; born blank-renban). If a palletized part has NO renban group, it falls back to the counter.
# Renban counter (Order.pas:709-745): read IN_RENBAN_COUNT, renban uses the CURRENT value, then the
# counter is INC'd with rollover (>999 -> 1). FRS# = copy(yymmdd,2,5)+suffix (the 5-char YMMDD of the
# order-by date + a 2-char trailer); the INSERT_OpenOrder proc derives VC_FRS_DATE (year-rollover) itself.
#
# FRS-SUFFIX OWNERSHIP (verified in review): INSERT_OpenOrder does NOT persist the caller's suffix as-is —
# it RE-COMPUTES the trailer server-side from max(VC_FRS_NUMBER) for the part and relies on the
# VC_FRS_NUMBER varchar(7) column truncating the over-long concat. The suffix this service computes
# matches the proc's only because both increment in lockstep from a clean FRS series; if rows already
# exist under the same prefix for the part, the proc's assignment wins. This is faithful to the legacy
# (Order.pas passes the same 7-char value to the same proc), but the proc owns the final suffix.
# IG83-TODO (architects): decide whether the rebuilt write should pass only the 5-char prefix and let
# the proc own the trailer, rather than computing a suffix the proc overwrites.
#
# Jython 2.7 (8.1 and 8.3). Pure logic (computeOrderRecords) imports nothing and is CPython-importable
# for unit test. Only commitOrders touches the gateway globals `system`.
# IG81-COMPAT: createSProcCall / runPrepQuery identical on 8.1.52 and 8.3.

DATABASE = "Inventory_Spike"


def _frsPrefix(frsDateYmd):
	"""The 5-char FRS prefix = copy(formatdatetime('yymmdd', orderByDate), 2, 5). From an 8-char
	yyyymmdd that is yymmdd = yyyymmdd[2:] then drop its first char -> yyyymmdd[3:8] (Y + MMDD)."""
	return str(frsDateYmd)[3:8]


def _bumpRenban(count):
	"""The legacy renban counter step: returns the new counter after one use (INC, rollover >999 -> 1)."""
	nxt = count + 1
	return 1 if nxt > 999 else nxt


def computeOrderRecords(supCode, partNum, kanban, lotSizedBit, renbanGroup,
                        lotCount, qtyQ, lotQty, frsDateYmd, renbanCount):
	"""Pure: one worksheet line -> the order records to insert + the new renban counter. No DB.

	supCode/partNum/kanban : the part's codes.
	lotSizedBit  : INV_PARTS_STOCK_MST.BIT_LOT_SIZE_ORDERS as 0/1 (INVERTED: 0 = lot-sized, 1 = palletized).
	renbanGroup  : VC_RENBAN_GROUP_CODE ('' = none -> counter renban; non-'' -> blank renban, grouped later).
	lotCount     : worksheet R (number of lots) — drives the per-lot loop for lot-sized.
	qtyQ         : worksheet Q (= lotCount * lotQty) — the palletized single-record qty.
	lotQty       : IN_1LOTQTY — the per-lot record qty for lot-sized.
	frsDateYmd   : the order-by date as 8-char yyyymmdd (the worksheet FRS-date cell).
	renbanCount  : the part's current IN_RENBAN_COUNT (starting value).

	Returns (records, newRenbanCount) where each record is a dict ready for INSERT_OpenOrder.
	"""
	prefix = _frsPrefix(frsDateYmd)
	hasGroup = str(renbanGroup or "").strip() != ""
	count = int(renbanCount)
	records = []

	def _emit(suffix, qty):
		# assign renban + advance the counter (only when there is no renban group)
		renban = ""
		newcount = count
		if not hasGroup:
			renban = "%s%03d" % (kanban, count)
			newcount = _bumpRenban(count)
		records.append({"supCode": supCode, "partNum": partNum, "kanban": kanban,
		                "frsNum": prefix + suffix, "renbanNum": renban, "qty": int(qty)})
		return newcount

	lotSized = (int(lotSizedBit) == 0)   # INVERTED flag
	if lotSized:
		# ONE record per lot: FRS '01','02',… (2-char zero-padded), qty = lotQty.
		for j in range(1, int(lotCount) + 1):
			count = _emit("%02d" % j, lotQty)
	else:
		# PALLETIZED: ONE record, FRS '01', qty = Q; renban blank when grouped (the usual case).
		count = _emit("01", qtyQ)

	return records, count


# ---------------------------------------------------------------------------------------------
# Gateway driver — read the part's renban counter, compute, INSERT each order, persist the counter.
# (Touches `system`; not exercised by the pure unit test.)
# ---------------------------------------------------------------------------------------------
def _scalar(sql, args, db):
	rows = system.db.runPrepQuery(sql, args, db)
	return rows[0][0] if len(rows) else None


def commitOrders(line, database=None):
	"""Write the order records for one worksheet line and persist the renban counter.

	`line` is a dict: supCode, partNum, kanban, lotSizedBit, renbanGroup, lotCount, qtyQ, lotQty,
	frsDateYmd. Reads INV_PARTS_STOCK_MST.IN_RENBAN_COUNT, computes the records (computeOrderRecords),
	inserts each via INSERT_OpenOrder, and writes the advanced counter back via UPDATE_PartsStockRenban.

	D11#7 (flagged): the legacy read-then-write of IN_RENBAN_COUNT races under concurrent order-creates
	(two specialists could read the same count -> duplicate renbans). This mirrors the legacy (single
	read, batched write of the final count); the cutover fix is an atomic counter allocation. Escalate.
	"""
	db = database if database is not None else DATABASE
	count0 = _scalar("SELECT IN_RENBAN_COUNT FROM INV_PARTS_STOCK_MST WHERE VC_PART_NUMBER = ?",
	                 [line["partNum"]], db)
	if count0 is None:
		# The part must exist (the worksheet sourced it). Fail loudly rather than silently seed 0.
		raise ValueError("order.commitOrders: unknown part %r (no INV_PARTS_STOCK_MST row)" % (line["partNum"],))
	records, newCount = computeOrderRecords(
		line["supCode"], line["partNum"], line["kanban"], line["lotSizedBit"], line["renbanGroup"],
		line["lotCount"], line["qtyQ"], line["lotQty"], line["frsDateYmd"], int(count0))

	# Atomic across ALL records + the counter write, matching the legacy BeginTrans/CommitTrans
	# (Order.pas:686/757) — a mid-loop failure must not leave partial orders or a stale renban counter.
	# EXEC via runPrepUpdate on the shared txId (createSProcCall can't join a gateway transaction);
	# every value is a bound ? param.
	tx = system.db.beginTransaction(db)
	try:
		for r in records:
			system.db.runPrepUpdate(
				"EXEC INSERT_OpenOrder @SupCode=?, @PartNum=?, @KanbanNum=?, @FRSNum=?, @RenbanNum=?, @Qty=?",
				[r["supCode"], r["partNum"], r["kanban"], r["frsNum"], r["renbanNum"], r["qty"]], db, tx=tx)
		if records and newCount != int(count0):
			system.db.runPrepUpdate("EXEC UPDATE_PartsStockRenban @PartNum=?, @RenbanCount=?",
			                        [line["partNum"], newCount], db, tx=tx)
		system.db.commitTransaction(tx)
	except:
		system.db.rollbackTransaction(tx)
		raise
	finally:
		system.db.closeTransaction(tx)
	return records
