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
# This module is the PURE decision logic ONLY. The gateway driver `create_asn` (the transaction
# orchestration that reads AD_FRSPull + SELECT_ForecastDetailBCASN, inserts the header, calls
# INSERT_ASNDetail, and allocates the EIN AT SEND) is DEFERRED — it needs the live VehicleOrder
# backup (the GALC Vehicle/Model/vehicledata/DataItem tables that feed AD_FRSPull are NOT on the
# spike yet). Live end-to-end parity re-runs against that backup. See spec §3/§9 + §7.
#
# Jython 2.7 (8.1 and 8.3, no delta). Imports nothing -> CPython-importable for the unit test
# (scripts/e2e/test_asn_fanout.py). The fan-out abort conditions (missing manifest cost; BC with no
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
