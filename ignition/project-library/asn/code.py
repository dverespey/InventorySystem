# asn — Project Library: the ASN-creation ratio fan-out (CalculateASNFRS, hand-written Delphi
# DataModule.pas:5106-5319) reimplemented as PURE Jython-2.7/CPython logic, the producer-recipe
# pattern (twin of inventory-stock/shipping.computePosts).
#
# Source truth:
#   docs/analysis/production-readiness/m1-asn-creation-spec.md  (the decoded create-ASN chain)
#   docs/analysis/production-readiness/AD_FRSPULL-shared.sql    (the GALC vehicle-count qty source)
#   DataModule.pas:5180-5268                                    (the EXACT branch/round/manifest logic)
#
# ROLE: turn the GALC build data (AD_FRSPull rows, per broadcast code: BC/Orders/Vehicles) + the
# Inventory forecast-detail ratios (SELECT_ForecastDetailBCASN rows: part + IN_TIRE_RATIO/
# IN_WHEEL_RATIO + IN_ASSY_QTY + VC_ASSY_MANIFEST_NUMBER) into the N ASN-detail rows that feed
# INSERT_ASNDetail (the per-manifest accumulating upsert, §5 / PART C).
#
# This module holds BOTH the PURE decision logic (computeAsnDetails) AND the thin gateway driver
# (create_asn), the order/code.py pattern (computeOrderRecords + commitOrders side by side).
#   * The PURE part imports nothing and references no gateway global at import time -> it stays
#     CPython-importable for the unit test (scripts/e2e/test_asn_fanout.py).
#   * create_asn references `system` (and AD_FRSPull on the ALC datasource) only at CALL time, exactly
#     as order.commitOrders does, so importing the module never needs the gateway runtime. It is driven
#     headlessly through scripts/e2e/jython_shim.py (the R8 system.db shim) for end-to-end parity
#     (scripts/e2e/test_create_asn_parity.py) against the live VehicleOrder + Inventory backups.
#
# Jython 2.7 (8.1 and 8.3, no delta). The fan-out abort conditions (missing manifest cost; BC with no
# forecast detail) are surfaced as exceptions for the driver to roll back on, matching the Delphi
# `raise` -> CalculateASNFRS returns False -> CreateASNEntries_ButtonClick rolls back (spec §3).


class AsnFanoutError(Exception):
	"""Aborts the whole create (the Delphi `raise` inside CalculateASNFRS -> result:=False -> the
	caller rolls the Inv_Connection transaction back, spec §3 'Abort conditions')."""
	pass


# ---------------------------------------------------------------------------------------------
# Delphi-faithful rounding. DataModule.pas:5234 uses Pascal `Round`, which is BANKER'S rounding
# (round-half-to-even) — NOT round-half-up/away-from-zero. Python 3's built-in round() is also
# banker's, but Jython 2.7's round() is round-half-AWAY-from-zero, so we MUST implement it
# explicitly to be correct on BOTH runtimes (the known parity trap, spec §9.3).
#
# We round the EXACT rational value num/den (den = 100 here) to avoid binary-float drift at the .5
# boundary: a value like 2.5 represented as a float can already be slightly off, which would make a
# float-based round silently pick the wrong neighbour. Integer arithmetic on (num, den) is exact.
# ---------------------------------------------------------------------------------------------
def _bankers_div_round(num, den):
	"""round(num/den) with round-half-to-even, computed exactly in integers. den > 0.
	Mirrors Delphi Round() applied to (VEHICLES*IN_ASSY_QTY*IN_TIRE_RATIO)/100."""
	if den <= 0:
		raise ValueError("asn._bankers_div_round: den must be > 0, got %r" % (den,))
	# Work on magnitude, restore sign at the end (banker's rounding is symmetric about zero).
	neg = num < 0
	n = -num if neg else num
	q, r = divmod(n, den)          # n = q*den + r, 0 <= r < den
	twice = 2 * r
	if twice > den:
		q += 1                      # remainder > half -> round up
	elif twice == den:
		if q % 2 == 1:              # exactly half -> round to EVEN neighbour
			q += 1
	# twice < den -> round down (q unchanged)
	return -q if neg else q


def _manifest(productionDate, assyManifestNumber):
	"""'7' + copy(fProductionDate,4,5) + VC_ASSY_MANIFEST_NUMBER (DataModule.pas:5186 / :5239).

	Pascal copy(s,4,5) is 1-based: chars 4..8 of yyyymmdd = last-digit-of-year + MM + DD (the
	spec's '1-digit year' correction, §6). E.g. '20260618' -> '60618'; with assy id '57' ->
	'76061857' (matches the daily-log manifest 76061857). assyManifestNumber is the 2-char id from
	INV_FORECAST_DETAIL_INF.VC_ASSY_MANIFEST_NUMBER."""
	productionDate = str(productionDate)
	if len(productionDate) != 8:
		raise ValueError("asn._manifest: productionDate must be yyyymmdd (8 chars), got %r"
		                 % (productionDate,))
	return "7" + productionDate[3:8] + str(assyManifestNumber)


def _int(v):
	return int(v)


# ---------------------------------------------------------------------------------------------
# computeAsnDetails — the PURE fan-out. Reproduces DataModule.pas:5180-5268 exactly.
# ---------------------------------------------------------------------------------------------
def computeAsnDetails(frsRows, forecastByBc, productionDate, asnId):
	"""Fan one ASN's build data + forecast ratios out into the detail rows feeding INSERT_ASNDetail.

	Inputs
	------
	frsRows : list of dicts, one per broadcast code from AD_FRSPull (the GALC vehicle-count pull).
	          Keys: 'BC' (broadcast code, str), 'Orders' (int — vehicles*4 ground / vehicles spare;
	          the <=5 No-Ratio trigger), 'VEHICLES' (int — the qty multiplier). The driver passes
	          AD_FRSPull's result rows verbatim.
	forecastByBc : dict BC -> list of forecast-detail dicts from SELECT_ForecastDetailBCASN for that
	          BC. Each dict's keys: 'VC_ASSY_PART_NUMBER_CODE', 'VC_ASSY_MANIFEST_NUMBER',
	          'IN_ASSY_QTY', 'IN_TIRE_RATIO', 'IN_WHEEL_RATIO', 'IN_MANIFEST_COST_ID' (None = missing
	          cost -> abort).
	          *** ORDER CONTRACT (FIX 2): the caller MUST pass each BC's rows in a DETERMINISTIC order.
	          The No-Ratio branch (Orders<=5) emits ONE row from fcRows[0] and `break`s — so list order
	          decides WHICH assy/manifest a small-volume BC ships — and the ratio branch emits in list
	          order. SELECT_ForecastDetailBCASN has no ORDER BY (two HEAPs), so the DRIVER (create_asn,
	          step 3) sorts the rows by ID_FORECAST_DETAIL ascending (first-configured) before calling
	          this. This pure function does NOT re-sort (it would need an extra key + couples ordering
	          policy here); it RELIES on the caller's deterministic order. Tests must pass rows pre-sorted
	          the same way. Pass rows in heap/arbitrary order and the No-Ratio pick is nondeterministic.
	productionDate : 'yyyymmdd' (fProductionDate). Drives the manifest prefix.
	asnId : the ASN header id (fRecordID from INSERT_ASNInfo OUTPUT). Stamped on every detail so the
	          re-keyed INSERT_ASNDetail (PART C) accumulates per-ASN.

	Returns
	-------
	list of detail dicts: {'asnId', 'manifest', 'partNumber', 'qty', 'noRatio'} — one per
	INSERT_ASNDetail call, in emission order. 'noRatio' flags the small-volume single-row branch
	(the log's 'INSERT ASN entry(No Ratio)' rows) purely for traceability/logging parity; it does
	not change the upsert.

	Faithful to DataModule.pas:5180-5268:
      * No-Ratio branch  (Orders <= 5, :5183): emit ONE row from the FIRST forecast-detail row,
        qty = VEHICLES * IN_ASSY_QTY (no ratio), then `break` (so only the first part of that BC
        ships). This is the legacy small-volume case.
      * Ratio branch     (Orders >  5, :5214): one row PER forecast-detail row. If BOTH ratios are
        100 (:5226) qty = VEHICLES * IN_ASSY_QTY (no rounding); else
        qty = round(VEHICLES * IN_ASSY_QTY * IN_TIRE_RATIO / 100) with banker's rounding (:5234).
      * Manifest = _manifest(productionDate, VC_ASSY_MANIFEST_NUMBER) in BOTH branches.
      * Abort if any forecast-detail row of a BC has IN_MANIFEST_COST_ID is None (:5163-5174), OR a
        BC produced by AD_FRSPull has NO forecast-detail rows (:5270-5275).
        NB: this is the PRE-loop manifest-cost RAISE (:5160-5175) — abort the whole BC before any emit.
        It is NOT the post-loop SELECT_ASNMissingCost (:5285-5308), which is a warn-and-continue audit
        that belongs to the DRIVER (deferred) — do not turn that into an abort here.

	NOTE: the cross-DB AD_FRSPull read, the per-BC SELECT_ForecastDetailBCASN read, the header
	INSERT_ASNInfo, the INSERT_ASNDetail upsert calls, and the post-loop SELECT_ASNMissingCost audit
	are the DRIVER's job (deferred). This function does the pure translation only.
	"""
	details = []
	for frs in frsRows:
		bc = frs["BC"]
		orders = _int(frs["Orders"])
		vehicles = _int(frs["VEHICLES"])

		fcRows = forecastByBc.get(bc)
		if not fcRows:
			# DataModule.pas:5270-5275 — AD_FRSPull returned a BC with no forecast detail -> abort.
			raise AsnFanoutError(
				"Missing Broadcast Code Information(%s), ASN create failed" % (bc,))

		# Manifest-cost pre-check over ALL rows of the BC (DataModule.pas:5160-5175): any part with a
		# NULL IN_MANIFEST_COST_ID aborts the WHOLE create before any detail of this BC is emitted.
		missing = [r["VC_ASSY_PART_NUMBER_CODE"] for r in fcRows
		           if r.get("IN_MANIFEST_COST_ID") is None]
		if missing:
			raise AsnFanoutError(
				"Missing Manifest Cost Information BCode(%s) Assy Part Number ( %s )ASN create failed"
				% (bc, " ".join(str(m) for m in missing)))

		if orders <= 5:
			# --- No-Ratio branch (DataModule.pas:5183-5212): one row from the FIRST fc row, break.
			r = fcRows[0]
			qty = vehicles * _int(r["IN_ASSY_QTY"])
			details.append({
				"asnId": asnId,
				"manifest": _manifest(productionDate, r["VC_ASSY_MANIFEST_NUMBER"]),
				"partNumber": r["VC_ASSY_PART_NUMBER_CODE"],
				"qty": qty,
				"noRatio": True,
			})
			# the Delphi `break` — stop after the first forecast-detail row for this BC.
			continue

		# --- Ratio branch (DataModule.pas:5214-5265): one row per forecast-detail row.
		for r in fcRows:
			tire = _int(r["IN_TIRE_RATIO"])
			wheel = _int(r["IN_WHEEL_RATIO"])
			base = vehicles * _int(r["IN_ASSY_QTY"])
			if tire == 100 and wheel == 100:
				qty = base                                  # :5229 both 100% -> full qty, no round
			else:
				qty = _bankers_div_round(base * tire, 100)  # :5234 round(base*tire/100), banker's
			details.append({
				"asnId": asnId,
				"manifest": _manifest(productionDate, r["VC_ASSY_MANIFEST_NUMBER"]),
				"partNumber": r["VC_ASSY_PART_NUMBER_CODE"],
				"qty": qty,
				"noRatio": False,
			})

	return details


# =================================================================================================
# create_asn — the GATEWAY DRIVER. Wires the pure fan-out (computeAsnDetails) to the DB, reproducing
# the ordered create-ASN chain (CreateASNEntries_ButtonClick + InsertASNInfo + CalculateASNFRS,
# m1-asn-creation-spec §0/§7). This is the "Create ASN entries only" path (NOT Create ASN/Files — the
# 856 build + status flip to 'S' is M1 Rank 2, decoupled per Q2).
#
# IG81-COMPAT: every gateway API used here (runPrepQuery / runPrepUpdate(...,tx=) / beginTransaction /
# getLogger) is identical on 8.1.52 and 8.3 — no version guard needed.
# =================================================================================================

# Logical datasource names. In the gateway these are connection names; the headless shim maps them to
# the spike DBs (Inventory / VehicleOrder). DATABASE = the Inventory rebuild; ALC_DATABASE = the shared
# VehicleOrder DB that hosts AD_FRSPULL (a cross-DB READ, like AD_GetSpecialDate — not relocated).
from db_shared import CONNECTION as DATABASE   # centralized DB-conn name (default Inventory_Spike; single prod-rename point)
ALC_DATABASE = "VehicleOrder"

# FIX 1 (BLOCKER) — the filtered UNIQUE index that backstops the idempotency guard.
# spike-asn-unique-guard.sql creates UX_INV_ASN_MST_LINE_PDATE_NORMAL on
# (VC_LINE_NAME, VC_PRODUCTION_DATE) WHERE VC_START_SEQ_NUMBER <> '-1'. The driver's read-back guard
# (SELECT_ASNSeq) handles the common case; this index catches the concurrent-create RACE the read-back
# guard cannot see (two gateway sessions both pass the guard, both INSERT). create_asn catches the
# resulting unique-violation and treats it as the same idempotent "already created" skip.
ASN_UNIQUE_INDEX = "UX_INV_ASN_MST_LINE_PDATE_NORMAL"


def _isUniqueViolation(exc):
	"""True if `exc` is the SQL-Server unique-constraint violation for the ASN (line, prodDate) key.

	Runtime-portable across BOTH the live gateway and the headless test shim by matching on the error
	MESSAGE TEXT, which is stable on every path:
	  * Gateway (real JDBC): the duplicate-key insert surfaces as a java.sql.SQLException whose message
	    is the SQL-Server text 'Cannot insert duplicate key row ... with unique index
	    "UX_INV_ASN_MST_LINE_PDATE_NORMAL" ...' (SQL error 2601). Ignition wraps it but preserves the
	    message in str(exc).
	  * Shim (jython_shim._TxSession): raises SqlError carrying the same 'Msg 2601 ... Cannot insert
	    duplicate key row ... unique index 'UX_INV_ASN_MST_LINE_PDATE_NORMAL'' text.
	We match on EITHER the SQL-Server error number 2601 OR the specific index name — so we never swallow
	an unrelated error (a different constraint, a deadlock, a connection drop) as an idempotent skip.
	(SQL 2601 = duplicate key in a unique INDEX; 2627 = unique CONSTRAINT — accept both for safety, but
	this artifact creates an INDEX so 2601 is the live one, proven by PROBE-A: 'Msg 2601'.)"""
	msg = str(exc)
	if ASN_UNIQUE_INDEX in msg:
		return True
	if "2601" in msg or "2627" in msg:
		# guard the bare-number match against a coincidental substring: require the duplicate-key text.
		return "duplicate key" in msg.lower()
	return False


def _effMonth(productionDate):
	"""@EffMonth for SELECT_ForecastDetailBCASN = copy(prodDate,1,4)+'/'+copy(prodDate,5,2)
	(DataModule.pas:5154) = 'yyyy/MM' (7-char). The forecast rows are stored with this exact format or
	'' (the proc's OR VC_EFFECTIVE_MONTH = '' branch). Verified on the spike: '2026/06'."""
	productionDate = str(productionDate)
	return productionDate[0:4] + "/" + productionDate[4:6]


def create_asn(line, prodDate, seqStart, seqLast, beginDate, endDate, shipQty,
               site=1, database=None, alcDatabase=None):
	"""Create one ASN header + its manifest-detail fan-out in a single Inventory transaction.

	Reproduces the "Create ASN entries only" chain (spec §0/§7):
	  1. SELECT_ASNSeq idempotency guard  — block if an ASN already exists for (line, prodDate).
	     (Fast common-path no-op; NOT atomic on its own — the step-5 unique index backstops the race.)
	  2. AD_FRSPULL on the ALC datasource — per-BC vehicle counts (READ, no tx; VehicleOrder DB).
	  3. per BC: SELECT_ForecastDetailBCASN on Inventory — parts + ratios (READ, no tx). Rows are sorted
	     by ID_FORECAST_DETAIL ascending here (FIX 2 — deterministic No-Ratio pick; the proc has no
	     ORDER BY). computeAsnDetails relies on this caller-side ordering (see its ORDER CONTRACT).
	  4. computeAsnDetails — the pure fan-out (branch / banker's round / manifest gen). RAISES
	     AsnFanoutError on a missing manifest cost or a BC with no forecast detail (the pre-loop
	     abort) BEFORE the transaction is opened, so nothing is written on abort.
	  5. ONE Inventory transaction: INSERT_ASNInfo (status 'C', OUTPUT @ASNID) then per detail
	     INSERT_ASNDetail (the re-keyed accumulate upsert). Commit; rollback + raise on error.
	     FIX 1 (BLOCKER): the header insert is RACE-SAFE — if a concurrent gateway session won the
	     create, the filtered UNIQUE index (spike-asn-unique-guard.sql) makes our INSERT fail with the
	     duplicate-key violation; we catch it and return the idempotent skip (writes nothing), the same
	     outcome as the step-1 guard. Any OTHER tx error still rolls back + re-raises.
	  6. post-loop SELECT_ASNMissingCost audit — WARN only (does NOT abort), matching the Delphi
	     warn-and-continue (DataModule.pas:5285-5308).

	Params
	------
	line      : VC_LINE_NAME (e.g. 'COROLLA').
	prodDate  : 'yyyymmdd' production date (drives manifest prefix + the idempotency key + effMonth).
	seqStart  : start broadcast seq number (VC_START_SEQ_NUMBER, 4-char). Passed to AD_FRSPULL as
	            @Start but UNUSED there (verified — AD_FRSPULL filters purely by date+line); written
	            onto the ASN header.
	seqLast   : end seq number (VC_END_SEQ_NUMBER). Same: @Last unused by AD_FRSPULL, header only.
	beginDate : DT_START_SEQ — the start of the build datetime window (the active AD_FRSPULL filter).
	            A datetime/ISO string the JDBC layer accepts.
	endDate   : DT_END_SEQ — the end of the build window. AD_FRSPULL counts vehicles in [begin,end].
	shipQty   : IN_QTY on the ASN header — the Check-button vehicle count (legacy fQty, sourced from
	            AD_ProductionSeq, spec §1b; ASNSelect.pas:382 StrToInt(ShipQty_MaskEdit.Text)). The
	            caller supplies it; it is NOT derived from AD_FRSPULL (whose SPARE rows are 1-per-vehicle
	            and GROUND rows *4, neither of which is the distinct vehicle count). Header-only; the
	            detail-row parity does not depend on it.
	site      : site id. KEPT in the signature for the M4 multi-site re-key; NOT yet a column on
	            INV_ASN_MST / INV_ASN_DETAIL_MST (added at the M4 schema surgery — see spec §1/§5),
	            so it does not touch SQL today beyond the idempotency-guard intent.
	database     : Inventory connection name (defaults to DATABASE).
	alcDatabase  : ALC/VehicleOrder connection name for AD_FRSPULL (defaults to ALC_DATABASE).

	Returns a dict: {'asnId', 'details', 'qty', 'skipped'(bool), 'missingCost'(list of audit rows)}.
	If an ASN already exists for (line, prodDate) returns {'skipped': True, ...} and writes nothing
	(the legacy locks the UI; here we no-op idempotently). The same {'skipped': True} shape is returned
	if a concurrent session won the create and our header insert lost the unique-index race (FIX 1).

	EIN HANDLING — NOT allocated at create. The legacy stamps SiteEIN+1 onto the header at create and
	bumps the ALC Site counter inside the create tx (spec §2). The M1 build decision is the cleaner
	at-SEND model: the per-site EIN is allocated atomically from INV_SITES.IN_EIN_SEQ when the 856 is
	sent (M1 Rank 2 / the 856-send increment), NOT here. So we write IN_ASN_EIN = 0 (unset) at create;
	the send increment claims and stamps the real EIN. This is an INTENDED divergence from legacy
	parity (flagged in spec §2/§8) — the ASN_DETAIL parity diff therefore ignores IN_ASN_EIN.
	"""
	db = database if database is not None else DATABASE
	alcDb = alcDatabase if alcDatabase is not None else ALC_DATABASE
	log = system.util.getLogger("SPIKE.create_asn")

	# --- 1. idempotency guard (SELECT_ASNSeq, spec §1a) -----------------------------------------
	# Re-keyed intent (site_id, line, prodDate) at M4; today (line, prodDate). VC_START_SEQ_NUMBER<>-1
	# in the proc excludes the placeholder/hot-call header rows (e.g. the -1/-1 row at id 4722).
	existing = system.db.runPrepQuery(
		"EXEC SELECT_ASNSeq @LineName=?, @PDate=?", [line, str(prodDate)], db)
	if len(existing):
		log.warn("ASN already exists for line=%s prodDate=%s -> skip (idempotent)" % (line, prodDate))
		return {"asnId": None, "details": [], "qty": 0, "skipped": True, "missingCost": []}

	# --- 2. AD_FRSPULL on the ALC datasource (READ, no transaction; VehicleOrder DB) -------------
	# @Start/@Last are passed for signature fidelity but UNUSED in the proc body (verified) — the
	# active filter is @begindate/@enddate + @LineName (AD_FRSPULL-shared.sql NOTE).
	frsDs = system.db.runPrepQuery(
		"EXEC AD_FRSPULL @begindate=?, @enddate=?, @Start=?, @Last=?, @LineName=?",
		[beginDate, endDate, int(seqStart), int(seqLast), line], alcDb)
	frsRows = []
	for row in frsDs:
		bc = row["BC"]
		bc = bc.strip() if bc is not None else bc   # AD_FRSPULL BC is char(3) -> may be space-padded
		frsRows.append({"BC": bc, "Orders": int(row["ORDERS"]), "VEHICLES": int(row["VEHICLES"])})

	# --- 3. per BC: SELECT_ForecastDetailBCASN on Inventory (READ, no transaction) ---------------
	effMonth = _effMonth(prodDate)
	forecastByBc = {}
	for frs in frsRows:
		bc = frs["BC"]
		if bc in forecastByBc:
			continue
		fcDs = system.db.runPrepQuery(
			"EXEC SELECT_ForecastDetailBCASN @BCode=?, @EffMonth=?", [bc, effMonth], db)
		fcRows = []
		for r in fcDs:
			fcRows.append({
				# FIX 2 (No-Ratio determinism): capture the PK/identity so the rows can be sorted into a
				# STABLE order before the fan-out. ID_FORECAST_DETAIL is INV_FORECAST_DETAIL_INF's
				# IDENTITY(1,1) PK (col 1, verified live) — ascending = first-configured order. The proc
				# SELECT_ForecastDetailBCASN has NO ORDER BY over two HEAPs, so without this the row order
				# is nondeterministic heap-scan order; see the sort below + the computeAsnDetails contract.
				"ID_FORECAST_DETAIL": int(r["ID_FORECAST_DETAIL"]),
				"VC_ASSY_PART_NUMBER_CODE": r["VC_ASSY_PART_NUMBER_CODE"],
				"VC_ASSY_MANIFEST_NUMBER": r["VC_ASSY_MANIFEST_NUMBER"],
				"IN_ASSY_QTY": r["IN_ASSY_QTY"],
				"IN_TIRE_RATIO": r["IN_TIRE_RATIO"],
				"IN_WHEEL_RATIO": r["IN_WHEEL_RATIO"],
				"IN_MANIFEST_COST_ID": r["IN_MANIFEST_COST_ID"],
			})
		# FIX 2 (SHOULD-FIX, adversary-findings SHOULD-FIX-1): the shared legacy proc
		# SELECT_ForecastDetailBCASN returns rows in nondeterministic heap order (no ORDER BY over two
		# HEAPs). The No-Ratio branch in computeAsnDetails takes fcRows[0] and `break`s — so WHICH
		# assy/manifest a small-volume BC ships was luck-of-allocation (it FIRES on live BC PEE:
		# FEL1000/m36 vs FEL2000/m37). We do NOT ALTER the shared proc (other callers depend on it);
		# instead the REBUILD sorts the rows here, in the driver, by ID_FORECAST_DETAIL ascending so the
		# No-Ratio pick (and the ratio-branch emission order) is deterministic and reproducible.
		# TIEBREAK = lowest ID_FORECAST_DETAIL (= first-configured assy). CONFIRMED by David 2026-06-20
		#     (faithful-deterministic: it is the stable proxy for the legacy's undefined "first row" and
		#     reproduces what production actually shipped for No-Ratio BC PEE in ASN 4721 — manifest
		#     76061836 / FEL1000). NOT a ratio-based pick (that would be a behavior change vs legacy).
		fcRows.sort(key=lambda r: r["ID_FORECAST_DETAIL"])
		forecastByBc[bc] = fcRows

	# --- 4. pure fan-out. Raises AsnFanoutError (missing cost / no forecast detail) BEFORE we open
	#        the transaction, so an abort writes nothing — matching the Delphi pre-insert raise. The
	#        asnId is unknown until INSERT_ASNInfo runs, so pass a placeholder and re-stamp at write. -
	details = computeAsnDetails(frsRows, forecastByBc, str(prodDate), None)

	# --- 5. ONE Inventory transaction: header + details (spec §7 — fixes the legacy split where the
	#        ALC steps ran outside the Inv_Connection tx) -----------------------------------------
	tx = system.db.beginTransaction(db)
	raced = False           # set if the header insert lost the concurrent-create race (FIX 1).
	try:
		# INSERT_ASNInfo (status 'C', @Ein=0 — EIN allocated at SEND, see docstring). @ASNID is an
		# OUTPUT param, NOT an auto-generated key: getKey/SCOPE_IDENTITY() after a bare EXEC is NULL
		# (the proc's INSERT is a child scope). createSProcCall+registerOutParam can't join a gateway
		# transaction (see order.commitOrders), so we capture the OUTPUT explicitly with a DECLARE +
		# EXEC ... @ASNID=@id OUTPUT + SELECT @id, read back via runScalarPrepQuery on the SAME tx — a
		# real 8.1+ API that accepts the tx id so it shares the open BEGIN TRAN.
		# @AssyLine: VC_ASSEMBLY_LINE varchar(1); the morning entries-only create uses '' (non-key).
		#
		# FIX 1 (BLOCKER) — RACE-SAFE header insert. The step-1 SELECT_ASNSeq guard is non-atomic: a
		# concurrent gateway session can pass the same guard and reach this INSERT too. The filtered
		# UNIQUE index UX_INV_ASN_MST_LINE_PDATE_NORMAL (spike-asn-unique-guard.sql) makes the SECOND
		# committed normal header for (line, prodDate) fail with SQL error 2601 (duplicate key). We catch
		# THAT specific violation (via _isUniqueViolation) inside the transaction handling and treat it as
		# the idempotent "already created" skip — matching the guard's intent rather than erroring. Note
		# this fires on INSERT_ASNInfo (the header carries the unique key); the detail inserts are not
		# subject to it. Because the violation aborts before any detail is written, the rollback below
		# leaves nothing behind (all-or-nothing, same as any other tx error).
		asnId = system.db.runScalarPrepQuery(
			"DECLARE @id int; "
			"EXEC INSERT_ASNInfo @ASNID=@id OUTPUT, @LineName=?, @AssyLine=?, @StartSeq=?, "
			"@DTStartSeq=?, @EndSeq=?, @DTEndSeq=?, @Qty=?, @PDate=?, @Ein=?; "
			"SELECT @id",
			[line, "", str(seqStart), beginDate, str(seqLast), endDate,
			 int(shipQty), str(prodDate), 0],
			db, tx)
		if asnId is None:
			raise AsnFanoutError("create_asn: INSERT_ASNInfo returned no ASN id")
		asnId = int(asnId)

		# INSERT_ASNDetail per detail row — the re-keyed accumulate upsert (@HotCall defaults 0).
		# Re-stamp asnId now that the header exists (computeAsnDetails carried a placeholder None).
		for d in details:
			system.db.runPrepUpdate(
				"EXEC INSERT_ASNDetail @ASNID=?, @EIN=?, @Manifest=?, @PartNumber=?, @Qty=?",
				[asnId, 0, d["manifest"], d["partNumber"], int(d["qty"])], db, tx=tx)
			d["asnId"] = asnId

		system.db.commitTransaction(tx)
	except Exception as exc:
		# Always roll back the partial header (the rollback is safe even if SQL Server already doomed the
		# tx — the shim/JDBC guard @@TRANCOUNT). Then split: a unique-violation is the concurrent-create
		# RACE -> idempotent skip; anything else is a real failure -> re-raise (unchanged behaviour).
		system.db.rollbackTransaction(tx)
		if _isUniqueViolation(exc):
			raced = True
			log.warn("ASN concurrent-create race for line=%s prodDate=%s (unique index %s caught the "
			         "second insert) -> skip (idempotent)" % (line, prodDate, ASN_UNIQUE_INDEX))
		else:
			raise
	finally:
		system.db.closeTransaction(tx)

	if raced:
		# The other session already created this ASN; we wrote nothing. Same shape as the step-1 skip.
		return {"asnId": None, "details": [], "qty": 0, "skipped": True, "missingCost": []}

	# --- 6. post-loop missing-cost audit (SELECT_ASNMissingCost) — WARN, do NOT abort (the abort
	#        already happened pre-insert in step 4). Read-only; runs after commit, autocommit. --------
	missingCost = []
	try:
		auditDs = system.db.runPrepQuery("EXEC SELECT_ASNMissingCost @ASNID=?", [asnId], db)
		for r in auditDs:
			rowd = {"manifest": r["Manifest"], "partNumber": r["PartNumber"], "errorMsg": r["ErrorMsg"]}
			missingCost.append(rowd)
			log.warn("ASN %s missing-cost audit: Part(%s) Manifest(%s): %s"
			         % (asnId, rowd["partNumber"], rowd["manifest"], rowd["errorMsg"]))
	except Exception:
		# the Delphi wraps this audit in its own try/except and only logs the failure (it must not undo
		# a committed ASN). DataModule.pas:5310-5314.
		log.warn("ASN %s missing-cost audit failed (non-fatal)" % (asnId,))

	log.info("create_asn line=%s prodDate=%s -> ASN %s, %d detail rows, qty=%d"
	         % (line, prodDate, asnId, len(details), _detailQty(details)))
	return {"asnId": asnId, "details": details, "qty": _detailQty(details),
	        "skipped": False, "missingCost": missingCost}


def _detailQty(details):
	return sum(int(d["qty"]) for d in details)
