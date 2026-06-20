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
	          cost -> abort). The order of rows matters for the No-Ratio branch (it takes the FIRST).
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
DATABASE = "Inventory_Spike"
ALC_DATABASE = "VehicleOrder"


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
	  2. AD_FRSPULL on the ALC datasource — per-BC vehicle counts (READ, no tx; VehicleOrder DB).
	  3. per BC: SELECT_ForecastDetailBCASN on Inventory — parts + ratios (READ, no tx).
	  4. computeAsnDetails — the pure fan-out (branch / banker's round / manifest gen). RAISES
	     AsnFanoutError on a missing manifest cost or a BC with no forecast detail (the pre-loop
	     abort) BEFORE the transaction is opened, so nothing is written on abort.
	  5. ONE Inventory transaction: INSERT_ASNInfo (status 'C', OUTPUT @ASNID) then per detail
	     INSERT_ASNDetail (the re-keyed accumulate upsert). Commit; rollback + raise on error.
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
	(the legacy locks the UI; here we no-op idempotently).

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
				"VC_ASSY_PART_NUMBER_CODE": r["VC_ASSY_PART_NUMBER_CODE"],
				"VC_ASSY_MANIFEST_NUMBER": r["VC_ASSY_MANIFEST_NUMBER"],
				"IN_ASSY_QTY": r["IN_ASSY_QTY"],
				"IN_TIRE_RATIO": r["IN_TIRE_RATIO"],
				"IN_WHEEL_RATIO": r["IN_WHEEL_RATIO"],
				"IN_MANIFEST_COST_ID": r["IN_MANIFEST_COST_ID"],
			})
		forecastByBc[bc] = fcRows

	# --- 4. pure fan-out. Raises AsnFanoutError (missing cost / no forecast detail) BEFORE we open
	#        the transaction, so an abort writes nothing — matching the Delphi pre-insert raise. The
	#        asnId is unknown until INSERT_ASNInfo runs, so pass a placeholder and re-stamp at write. -
	details = computeAsnDetails(frsRows, forecastByBc, str(prodDate), None)

	# --- 5. ONE Inventory transaction: header + details (spec §7 — fixes the legacy split where the
	#        ALC steps ran outside the Inv_Connection tx) -----------------------------------------
	tx = system.db.beginTransaction(db)
	try:
		# INSERT_ASNInfo (status 'C', @Ein=0 — EIN allocated at SEND, see docstring). @ASNID is an
		# OUTPUT param, NOT an auto-generated key: getKey/SCOPE_IDENTITY() after a bare EXEC is NULL
		# (the proc's INSERT is a child scope). createSProcCall+registerOutParam can't join a gateway
		# transaction (see order.commitOrders), so we capture the OUTPUT explicitly with a DECLARE +
		# EXEC ... @ASNID=@id OUTPUT + SELECT @id, read back via runScalarPrepQuery on the SAME tx — a
		# real 8.1+ API that accepts the tx id so it shares the open BEGIN TRAN.
		# @AssyLine: VC_ASSEMBLY_LINE varchar(1); the morning entries-only create uses '' (non-key).
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
	except:
		system.db.rollbackTransaction(tx)
		raise
	finally:
		system.db.closeTransaction(tx)

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
