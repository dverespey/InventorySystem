# renban — Project Library service for the renban breakdown / trailer-grouping stage
# (legacy TGroupRenbanOrder_Form, RenbanOrder.pas; the live unit per InventorySystem.dpr:32).
#
# This is the MIDDLE of the daily order chain: Order.pas commits blank-renban placeholder orders
# for palletized parts -> THIS stage splits those placeholders across N trailers, assigning each a
# unique FRS# and a unique renban number -> OrderFormCreateF.pas serializes the now-eligible rows to
# supplier order files. It SHARES NO CODE with the Order commit (project-library/order/code.py) — a
# separate service, distinct lifecycle (delete-then-reinsert vs insert).
#
# Source spec: docs/analysis/order/renban-breakdown-spec.md (esp. §12 the STEP-0 rebuild algorithm,
# extracted verbatim from RenbanOrder.pas:255-301 / :746-799) + renban-data-analysis.md (proc bodies
# + the live-data proofs). Domain: project-order-renban-domain memory.
#
# WHAT IT REPRODUCES (faithful to RenbanOrder.pas):
#   * compute_trailer_breakdown(parts, trailers, palletsPerTrailer): pure in-memory distribution of
#     each part's lots across T equal-capacity trailers (lots div T base share + forward-spill of
#     overflow + lots mod T remainder dribbled one-lot-at-a-time round-robin), merging repeated parts
#     on a truck (SUM lots — the R3 sum-all, FAITHFUL), then the per-(truck,part) FRS# + renban + qty.
#   * commit_renban_breakdown(...): per part DELETE_OrderRenban(@FRS='',@Renban='') to clear ALL still-
#     blank placeholders, then INSERT_OpenOrder per trailer-row (renban NON-blank, qty>0, skip qty=0),
#     then once UPDATE_RenbanGroupCount(next_count). ONE transaction, rollback on any failure.
#
# KEY DOMAIN FACTS (proved — renban-data-analysis.md / re-proved 2026-06-21 on mssql-spike):
#   * The blank renban (VC_RENBAN_NUMBER='') is BOTH the selection flag (SELECT_OrderNoRenban pulls
#     only =='') AND the downstream eligibility gate (the emitter pulls only <>''). The breakdown's
#     whole job is to flip blank -> assigned, so the re-inserted rows ship.
#   * FRS-SUFFIX OWNERSHIP — for the renban breakdown the proc's server-side recompute is a NO-OP:
#     @FRSNum is varchar(7); we send the FULL 7-char FRS (5-char prefix + 2-digit trailer ordinal);
#     INSERT_OpenOrder's `@FRSNum = @FRSNum + <suffix>` produces 9 chars that silently TRUNCATE back
#     to 7 = our value. PROVED both branches: '6090102'+'01' -> '6090102', '6090103'+max+1 ->
#     '6090103' (both len 7). So the proc HONORS Pascal's TruckNumber+1 suffix; we do NOT reimplement
#     max+1 here (that recompute is only live for the original 5-char Order path). (data-analysis §3.2)
#   * Stock-neutral: the re-inserted rows are status-empty, so the INSERT/DELETE triggers (gated on
#     VC_STATUS_SUPPLIER_SHIPPING<>'') do not move INV_PARTS_STOCK_MST.IN_QTY. (data-analysis §3.1)
#   * The inverted BIT_LOT_SIZE_ORDERS flag is consumed UPSTREAM (Order.pas) — RenbanOrder never reads
#     it; the rows reaching here are already the palletized/blank-renban class. Not a hazard here.
#
# DO NOT resurrect the commented-out update-in-place path (RenbanOrder.pas:482-539 — abandoned because
# a non-first single-item truck was left without a renban). Delete-then-reinsert ONLY.
#
# Jython 2.7 (8.1 and 8.3). compute_trailer_breakdown is PURE (imports nothing, CPython-importable for
# the unit test). Only commit_renban_breakdown touches the gateway globals `system`.
# IG81-COMPAT: beginTransaction / runPrepUpdate / runPrepQuery identical on 8.1.52 and 8.3.

DATABASE = "Inventory_Spike"


# ---------------------------------------------------------------------------------------------
# PURE distribution — no DB. Faithful to RenbanOrder.pas:255-301 (distribution) + :746-799 (read-out).
# ---------------------------------------------------------------------------------------------
class _Truck(object):
    """One trailer: fixed capacity `size`, running `current` (lots loaded across ALL parts), and a
    per-part-ordered list of (part -> lots) so repeated parts MERGE (sum). Mirrors TTruck (:32-62,
    AddOrder :155-182)."""
    def __init__(self, size):
        self.size = size
        self.current = 0
        self._order_keys = []        # part numbers in insertion order (TStringList ordering)
        self._orders = {}            # part -> dict(kanban,supplier,part,frs,lotqty,lots,renban)

    def add(self, kanban, supplier, part, frs, lotqty, lots, renban):
        # TTruck.AddOrder :155-182 — merge repeated parts (SUM lots), bump CurrentCount unconditionally.
        if part in self._orders:
            self._orders[part]["lots"] += lots          # :160-163 sum, do NOT split into a 2nd line
        else:
            self._orders[part] = {"kanban": kanban, "supplier": supplier, "part": part,
                                  "frs": frs, "lotqty": lotqty, "lots": lots, "renban": renban}
            self._order_keys.append(part)
        self.current += lots                            # :179

    def orders(self):
        """The merged order lines in insertion order (Truck.First/Next walk, :184-228)."""
        return [self._orders[k] for k in self._order_keys]


def _distribute_part(trucks, kanban, supplier, part, frs, lotqty, lots, renban_seed):
    """TGroupRenban.AddOrder distribution math (:255-301) for ONE part with `lots` lots, distributed
    across `trucks` and ADDED to them in place (each truck.add bumps .current immediately, exactly as
    the Pascal TTruck.AddOrder does at :179). CRITICAL: the distribution decisions read trucks[i].current
    AS IT MUTATES within the loop — later Phase-A iterations and Phase B see the running counts of the
    SAME part's earlier adds. (An earlier delta-then-add-later refactor diverged from the .pas here.)
    `leftover` carries Phase A -> Phase B (:282)."""
    T = len(trucks)
    leftover = 0

    def _add(i, n):
        trucks[i].add(kanban, supplier, part, frs, lotqty, n, renban_seed)

    # --- Phase A: even base share, forward-spill overflow ----------------------------  :262-277
    if (lots // T) != 0:
        for i in range(T):
            share = (lots // T) + leftover
            if trucks[i].current + share <= trucks[i].size:                 # fits
                _add(i, share)                                              # :268 (bumps current)
                leftover = 0
            else:                                                           # truck full -> spill
                # :273 — NB uses trucks[0].size; identical to trucks[i].size since all share one size
                leftover = ((lots // T) + leftover) - (trucks[0].size - trucks[i].current)
                topoff = trucks[i].size - trucks[i].current                 # :274
                _add(i, topoff)                                             # bumps current to .size

    # --- Phase B: remainder one-lot-at-a-time, round-robin ---------------------------  :279-300
    if (lots % T) != 0:
        remainder = (lots % T) + leftover                                   # :282
        guard = 0
        max_passes = remainder + T + 1   # bound the loop: capacity gate should preclude a stall
        while remainder != 0:
            placed_this_pass = False
            for i in range(T):
                if trucks[i].current + 1 <= trucks[i].size:                 # :289
                    _add(i, 1)                                              # :291 (bumps current)
                    remainder -= 1
                    placed_this_pass = True
                if remainder == 0:
                    break                                                   # :295-296
            guard += 1
            if remainder != 0 and not placed_this_pass:
                # all trucks full but remainder left — the legacy would spin forever (:285). The
                # capacity gate (compute_trailer_breakdown) prevents this; fail loudly if it slips.
                raise ValueError("renban breakdown: %d lot(s) could not be placed (all trailers full) "
                                 "— capacity gate should have caught this" % remainder)
            if guard > max_passes:
                raise ValueError("renban breakdown: remainder distribution did not converge")


def _frs_suffix(frs, truck_number):
    """FRS = first 5 chars of the original FRS + the 2-digit trailer ordinal (TruckNumber+1).
    RenbanOrder.pas:763-767 — truck>8 emits the raw 2+ digits (10,11,...), else zero-padded 01..09."""
    prefix = str(frs)[:5]
    ordinal = truck_number + 1                       # 1-based
    if truck_number > 8:
        return prefix + str(ordinal)                 # :765 — '10','11',... (already 2 digits)
    return prefix + "0" + str(ordinal)               # :767 — '01'..'09'


def _renban_number(group_code, renban_seed, truck_number):
    """Renban = group_code + %.3d(seed3 + TruckNumber). RenbanOrder.pas:775-779.
    renban_seed is group_code||group_count (e.g. 'CMWA288'); the 3-digit tail is the seed count.

    Returns (renban_string, rcount) where rcount is the RAW (unwrapped) seed3 + truck_number — the
    string is ring-wrapped (%03d % (rcount % 1000)) but the rcount stays RAW.

    P4 CLEAN WRAP (divergence D-RNB-1, renban-collision-design.md §2.2): wrap the renban-number TAIL to
    the 000-999 ring so a run that straddles 999 emits CMWA998/CMWA999/CMWA000 (3 digits) instead of
    CMWA1000 (4 digits). Replaces the legacy str(N)[:3]-companion artifact (the legacy emitted the literal
    4-digit CMWA1000). Keeps every renban exactly group_code + 3 digits, fitting VC_RENBAN_NUMBER(8) AND
    avoiding a 4-digit renban Toyota could see. KEEP rcount RAW (unwrapped) so next_count = max(rcount)+1
    carries the TRUE running count for the single authoritative `% 1000` persist in step (c) — the two
    wraps compose (raw count forward, wrap on emit + wrap on persist). A run NOT crossing 999 is byte-for-
    byte identical to the legacy (rcount<1000 -> rcount % 1000 == rcount)."""
    seed3 = int(str(renban_seed)[-3:])               # rightstr(...,3) :775
    rcount = seed3 + truck_number                     # :777 (0-based truck index) — RAW, may exceed 999
    wrapped = rcount % 1000                           # D-RNB-1: ring-wrap the rendered tail only
    return group_code + ("%03d" % wrapped), rcount    # :779 (string wrapped; rcount RAW for next_count)


def compute_trailer_breakdown(parts, trailers, palletsPerTrailer, group_code):
    """Pure: distribute each part's lots across `trailers` equal-capacity trucks and assign per-trailer
    FRS#/renban/qty. No DB. Faithful to RenbanOrder.pas:255-301 + :746-799.

    parts            : list of dicts, one per blank-renban open-order row (from SELECT_OrderNoRenban,
                       ALIASED) with keys: kanban, supplier, part, frs, order_qty, lotqty,
                       group_code, renban_seed (= group_code||group_count, e.g. 'CMWA288').
    trailers         : T = number of trailers (operator input; the combo allows 1..6).
    palletsPerTrailer: P = each trailer's pallet (=lot) capacity (operator input).
    group_code       : the renban group code (the renban-number prefix, RenbanGroups_ComboBox.Text).

    Returns a dict:
      {"rows": [ {truck, part, kanban, supplier, frs, lots, lotqty, qty, renban}, ... ],  # skip qty==0
       "truck_counts": [int, ...],     # per-truck total lots (for the listbox summary)
       "next_count": int,              # the advanced group counter (UPDATE_RenbanGroupCount value)
       "total_lots": int}
    Raises ValueError on a bad trailer/pallet count, a capacity-gate violation (T*P < TotalLots),
    or a div-by-zero lotqty (H7). Rows with qty==0 are SKIPPED (legacy :540 "don't send 0 qty").
    """
    T = int(trailers)
    P = int(palletsPerTrailer)
    if T <= 0:
        raise ValueError("renban breakdown: trailer count must be >= 1 (got %r)" % (trailers,))
    if P <= 0:
        raise ValueError("renban breakdown: pallets/trailer must be >= 1 (got %r)" % (palletsPerTrailer,))

    # per-row lots = order_qty div lotqty (:622); guard div-by-zero (H7).
    enriched = []
    total_lots = 0
    for row in parts:
        lotqty = int(row["lotqty"]) if row["lotqty"] not in (None, "") else 0
        if lotqty <= 0:
            raise ValueError("renban breakdown: part %r has lotqty %r (0/NULL) — cannot compute lots "
                             "(misconfigured palletized part)" % (row.get("part"), row.get("lotqty")))
        lots = int(row["order_qty"]) // lotqty
        enriched.append((row, lotqty, lots))
        total_lots += lots

    # capacity gate (:709) — refuse if the lots will not fit.
    if T * P < total_lots:
        raise ValueError("renban breakdown: %d trailer(s) x %d pallets = %d capacity < %d total lots "
                         "(will not fit)" % (T, P, T * P, total_lots))

    trucks = [_Truck(P) for _ in range(T)]

    # distribute each part across the trucks, in the input row order (H10: SELECT_OrderNoRenban has no
    # ORDER BY, but each part's fill is independent so the truck loads are stable). _distribute_part
    # ADDS to the trucks in place (mutating .current as it goes — faithful to the interleaved Pascal).
    for row, lotqty, lots in enriched:
        _distribute_part(trucks, row["kanban"], row["supplier"], row["part"], row["frs"],
                         lotqty, lots, row.get("renban_seed", ""))

    # read-out: walk trucks outer, merged orders inner (:746-794). One row per (truck, part).
    out_rows = []
    truck_counts = []
    last_rcount = -1
    for ti in range(T):
        truck = trucks[ti]
        truck_counts.append(truck.current)
        for order in truck.orders():
            lots = order["lots"]
            qty = lots * order["lotqty"]
            frs_out = _frs_suffix(order["frs"], ti)
            renban_out, rcount = _renban_number(group_code, order["renban"], ti)
            if rcount > last_rcount:
                last_rcount = rcount                  # track max emitted rcount (:798 reads the last)
            if qty <= 0:
                continue                              # skip 0-qty trailer-rows (:540) — but rcount tracked
            out_rows.append({"truck": ti, "part": order["part"], "kanban": order["kanban"],
                             "supplier": order["supplier"], "frs": frs_out, "lots": lots,
                             "lotqty": order["lotqty"], "qty": qty, "renban": renban_out})

    # next group counter = last emitted rcount + 1 (:798 fNewMaxRenban := rcount+1). Robust to a
    # trailing-empty truck: == seed3 + (highest non-empty TruckNumber) + 1.
    if last_rcount < 0:
        # no rows at all (empty group) — leave the counter unchanged (the legacy never reaches here
        # because LoadScreen aborts on 0 records, :648). Re-derive the seed from the first part.
        seed3 = int(str(parts[0]["renban_seed"])[-3:]) if parts else 0
        next_count = seed3
    else:
        next_count = last_rcount + 1

    return {"rows": out_rows, "truck_counts": truck_counts,
            "next_count": next_count, "total_lots": total_lots}


# ---------------------------------------------------------------------------------------------
# Gateway driver — read the blank-renban orders, compute, then DELETE-then-reINSERT + bump the counter
# in ONE transaction. (Touches `system`; not exercised by the pure unit test.)
# ---------------------------------------------------------------------------------------------
# IG-DB-NOTE: SELECT_OrderNoRenban is `SELECT *` over a 3-table join with DUPLICATE column names
# (IN_QTY appears as both o.IN_QTY = ORDER qty (ord 6) and p.IN_QTY = on-hand stock (ord 42); same for
# VC_KANBAN_NUMBER / VC_PART_NUMBER). A Named Query / runPrepQuery MUST alias every column explicitly
# and take o.IN_QTY (the order qty), never "the IN_QTY column". We therefore DO NOT call the proc for the
# read; we issue the equivalent aliased SELECT so the duplicate trap can't bite. (data-analysis §2.2)
_SELECT_NO_RENBAN = """
SELECT  o.VC_KANBAN_NUMBER   AS kanban,
        o.VC_SUPPLIER_CODE   AS supplier,
        o.VC_PART_NUMBER     AS part,
        o.VC_FRS_NUMBER      AS frs,
        o.IN_QTY             AS order_qty,
        p.IN_1LOTQTY         AS lotqty,
        r.VC_RENBAN_GROUP_CODE  AS group_code,
        r.VC_RENBAN_GROUP_COUNT AS group_count
  FROM INV_OPEN_ORDER_INF o
  JOIN INV_PARTS_STOCK_MST p  ON o.VC_PART_NUMBER = p.VC_PART_NUMBER
  JOIN INV_RENBAN_GROUP_MST r ON r.IN_RENBAN_ID = p.IN_RENBAN_ID
                             AND r.VC_RENBAN_GROUP_CODE = ?
 WHERE o.VC_RENBAN_NUMBER = ''
"""


def loadBlankRenbanOrders(groupCode, database=None):
    """Read the blank-renban open orders for a renban group, ALIASED (avoids the SELECT_OrderNoRenban
    duplicate-IN_QTY trap). Returns a list of dicts shaped for compute_trailer_breakdown. Read-only."""
    db = database if database is not None else DATABASE
    ds = system.db.runPrepQuery(_SELECT_NO_RENBAN, [groupCode], db)
    rows = []
    for r in ds:
        gc = r["group_code"]
        cnt = r["group_count"]
        rows.append({"kanban": r["kanban"], "supplier": r["supplier"], "part": r["part"],
                     "frs": r["frs"], "order_qty": int(r["order_qty"]), "lotqty": int(r["lotqty"]),
                     "group_code": gc, "renban_seed": "%s%s" % (gc, cnt)})
    return rows


# =============================================================================================
# P4 — COLLISION-AWARE ALLOCATOR (WARN -> GUIDE -> FIX). renban-collision-design.md §3 +
# renban-collision-sourcetruth.md §2-§4. The clean wrap (above) makes the count VALID at rollover;
# this allocator is the LOAD-BEARING backstop that keeps a wrap from REUSING a still-resident number
# (CMWA: lap ~= retention -> ~863/1000 occupied, ~1mo free runway). All read-only checks run BEFORE
# beginTransaction (never hold a tx across a human decision); an in-tx RE-CHECK on EVERY commit path
# closes the TOCTOU window vs the concurrent Order commit (order/code.py:130 writes non-blank renbans
# for lot-sized parts with no unique-constraint backstop).
# =============================================================================================

# IN-USE predicate (sql-analyst, source-truth §2): RESIDENT-ROWS, STATUS-INDEPENDENT, EXACT equality.
# A renban is in use iff ANY resident INV_OPEN_ORDER_INF row carries it (it frees only when
# DELETE_AutoPurge removes the row). NO status filter (an ordered/shipped-but-unpurged renban is still
# occupied -> a status filter would UNDER-detect it). NO own-row exclusion: the predicate is SELF-SAFE
# by construction — this breakdown's INPUT rows are BLANK (VC_RENBAN_NUMBER=''), the candidates are
# non-blank, so `= @candidate` can never match the very rows we are about to delete-and-reinsert.
# `= @renban` (EXACT) not `LIKE grp+'%'` (which over-matches a residual 4-digit CMWA1000 artifact).
_RENBAN_IN_USE = ("SELECT VC_RENBAN_NUMBER AS renban, MIN(VC_FRS_NUMBER) AS open_frs, "
                  "       MIN(VC_PART_NUMBER) AS open_part, MAX(VC_ADD) AS last_add, COUNT(*) AS rows_open "
                  "  FROM INV_OPEN_ORDER_INF "
                  " WHERE VC_RENBAN_NUMBER IN (%s) "
                  " GROUP BY VC_RENBAN_NUMBER")


def check_renban_collisions(rows, database=None, tx=None):
    """For each DISTINCT candidate renban in `rows`, test the resident-rows in-use predicate in ONE
    batched query. Returns a list of collision dicts (one per colliding renban):
      {renban, candidate_part, open_frs, open_part, last_add, rows_open}
    Empty list = clear to commit. The breakdown's own (blank) input rows never match, so this is
    self-safe (source-truth §2). Consumes the predicate as a defined input; does NOT re-derive 'open'.

    `tx` (SHOULD-FIX-3): when given (a beginTransaction handle), the query runs ON THAT TX CONNECTION so
    the in-tx re-check truly shares the open transaction with the subsequent INSERTs (and sees committed
    concurrent writers under READ COMMITTED). The pre-commit GATE caller passes tx=None (read-only, no tx
    open yet)."""
    db = database if database is not None else DATABASE
    # distinct candidate renbans, preserve first-seen order + map renban -> its candidate part (for GUIDE)
    cand_order = []
    cand_part = {}
    for r in rows:
        rb = r["renban"]
        if rb and rb not in cand_part:
            cand_part[rb] = r["part"]
            cand_order.append(rb)
    if not cand_order:
        return []
    placeholders = ",".join(["?"] * len(cand_order))
    ds = system.db.runPrepQuery(_RENBAN_IN_USE % placeholders, list(cand_order), db, tx=tx)
    hit = {}
    for row in ds:
        hit[row["renban"]] = {"open_frs": row["open_frs"], "open_part": row["open_part"],
                              "last_add": row["last_add"], "rows_open": int(row["rows_open"])}
    collisions = []
    for rb in cand_order:                          # report in candidate order (stable)
        if rb in hit:
            h = hit[rb]
            collisions.append({"renban": rb, "candidate_part": cand_part[rb],
                               "open_frs": h["open_frs"], "open_part": h["open_part"],
                               "last_add": h["last_add"], "rows_open": h["rows_open"]})
    return collisions


# NEXT-FREE RUN-of-N (sql-analyst source-truth §3, extended to a contiguous RUN). A multi-trailer
# breakdown issues N SEQUENTIAL renbans (base, base+1, ... base+N-1, all % 1000). GUIDE must suggest a
# base whose ENTIRE run is free, not just a single TOP-1 number. Forward-ring scan from the current
# count: the first start S in [count .. count+999] such that {(S+k) % 1000 : k in 0..N-1} are ALL not
# in use. Exhaustion (no free run of N anywhere in the 000-999 ring) -> return None (hard WARN-cancel;
# NEVER silently wrap into a live number, source-truth §4 / edge 6).
def _all_resident_suffixes(group_code, database=None):
    """The set of 0..999 suffixes currently in use by a resident row of this group (EXACT group + 3-digit
    match, excluding the residual 4-digit CMWA1000 artifact). One read; read-only.

    SHOULD-FIX-1 (sql-adversary): take the 3-digit suffix as the EXACT positional field right after the
    group prefix (SUBSTRING from len(group)+1, length 3), NOT RIGHT(...,3). For a renban stored with a
    trailing space ('ZZRB289 ', LEN 7 / DATALENGTH 8) the old RIGHT(...,3) read '89 ' -> int('89 ')==89,
    recording 289 as 089 (so next_free_run thought 289 was FREE -> GUIDE-onto-occupied). SUBSTRING(.,len+1,
    3) reads the digits '289' regardless of any trailing space.

    The match is structural (NOT a `LIKE group+'[0-9][0-9][0-9]'` whose trailing-space semantics depend on
    varchar-vs-NVARCHAR collation — under the gateway JDBC driver string params bind as nvarchar, and an
    NVARCHAR bracket pattern does NOT trailing-space-match a varchar column the way a varchar pattern does;
    verified the LIKE form silently missed 'ZZRB289 ' through the JDBC-faithful shim). Instead:
      * LEFT(.,len) = group           -> exact group prefix (the LEFT result has no trailing space),
      * SUBSTRING(.,len+1,3) is digits (NOT LIKE '%[^0-9]%'),
      * LTRIM(RTRIM(.)) = group + that 3-digit field -> the renban (trimmed) is EXACTLY group + 3 digits,
        which ACCEPTS the trailing-space form ('ZZRB289 ' trims to 'ZZRB289') and REJECTS the residual
        4-digit artifact ('ZZRB1000' != 'ZZRB'+'100'). Latent (0 trailing-space rows live) but closed."""
    db = database if database is not None else DATABASE
    g = len(group_code)
    ds = system.db.runPrepQuery(
        "SELECT DISTINCT SUBSTRING(VC_RENBAN_NUMBER, ?, 3) AS suf FROM INV_OPEN_ORDER_INF "
        "WHERE LEFT(VC_RENBAN_NUMBER, ?) = ? "
        "  AND SUBSTRING(VC_RENBAN_NUMBER, ?, 3) NOT LIKE '%[^0-9]%' "
        "  AND LTRIM(RTRIM(VC_RENBAN_NUMBER)) = ? + SUBSTRING(VC_RENBAN_NUMBER, ?, 3)",
        [g + 1, g, group_code, g + 1, group_code, g + 1], db)
    used = set()
    for r in ds:
        suf = r["suf"]
        try:
            used.add(int(suf))                 # suf is the exact 3-digit field; no trailing space to trip int()
        except (ValueError, TypeError):
            continue
    return used


def next_free_run(group_code, start_count, n, database=None):
    """Forward-ring scan: the lowest base S at/after `start_count` (wrapping 999->000) such that the N
    contiguous suffixes [S, S+1, ..., S+N-1] (each % 1000) are ALL free (no resident row). Returns the
    base suffix (0..999) or None if NO free run of N exists anywhere in the ring (exhaustion -> the
    caller hard-WARNs with cancel-only; never silently wrap into a live number). Read-only."""
    db = database if database is not None else DATABASE
    used = _all_resident_suffixes(group_code, db)
    if len(used) >= 1000:
        return None                                # ring fully occupied — exhausted
    start = int(start_count) % 1000
    for step in range(1000):
        base = (start + step) % 1000
        run = [(base + k) % 1000 for k in range(n)]
        if all(s not in used for s in run):
            # also require the run not to overlap itself wrapping past 999 in a degenerate huge-N case
            if len(set(run)) == n:
                return base
    return None                                    # no free run of N anywhere -> exhaustion


def _remap_rows_onto_base(rows, group_code, base):
    """PURE re-map: re-issue every trailer row's renban onto a chosen base suffix, preserving each row's
    POSITION in the run. Returns (new_rows, new_next_count).

    BLOCKER-1 FIX (sql-adversary, was code.py:384-407): the per-row offset MUST come from the row's
    SEQUENTIAL POSITION in the run (the trailer/renban index, 0-based in emission order), NOT from the
    rendered renban TAIL value. Deriving the offset from the tail as `(tail - min(tails)) % 1000` is
    correct ONLY for a non-wrapped run; for a run that straddles 999->000 (tails e.g. [997,998,999,0,1],
    min=0) it yields offsets [997,998,999,0,1] and re-seats the rows onto [base-3..base+1] — a block
    `next_free_run` NEVER validated (silent GUIDE-onto-occupied collision) AND persists a wrong count.

    `next_free_run(n=<distinct renbans>)` validates the CONTIGUOUS block [base..base+n-1] (% 1000). So
    the remap must write EXACTLY that block: rank the DISTINCT renbans in EMISSION ORDER (rows come out
    truck-outer, so the first-seen distinct renban is the run's forward-lowest = truck 0), and map the
    i-th distinct renban (0-based) onto (base + i) % 1000. Rows sharing a truck (same original renban)
    get the SAME new suffix, exactly as they shared the original. new_next_count = base + (N-1) + 1 (RAW;
    step (c) wraps on persist via % 1000) — the TRUE highest-written suffix + 1, so the next breakdown
    seeds PAST every number this run wrote (no self-collision)."""
    if not rows:
        return [], None
    # distinct original renbans in EMISSION order (truck-outer scan -> first-seen = forward-lowest = the
    # validated run base). Their 0-based rank IS the run offset, so the rewrite is exactly [base..base+N-1].
    rank = {}
    for r in rows:
        rb = r["renban"]
        if rb not in rank:
            rank[rb] = len(rank)          # 0,1,2,... in first-seen (emission) order
    N = len(rank)
    new_rows = []
    for r in rows:
        offset = rank[r["renban"]]        # 0-based position in the run, NOT the tail value
        new_suffix = (base + offset) % 1000
        nr = dict(r)
        nr["renban"] = group_code + ("%03d" % new_suffix)
        new_rows.append(nr)
    new_next_count = base + (N - 1) + 1            # = base + N; RAW, step (c) wraps on persist (% 1000)
    return new_rows, new_next_count


def _write_renban_alarm(db, siteId, summary, renban, candidatePart, errorText, tx=None):
    """Write ONE RENBAN_COLLISION alarm row (INV_EDI_ALARM_REJ) — the DATA/flag side of the home-hub
    attention surface (mirrors edi_inbound._write_alarm). The colliding renban goes in VC_MANIFEST_NUMBER
    (varchar(8) — exact fit for a renban string). BIT_RESOLVED defaults 0 (active). Returns the new
    IN_ALARM_ID. IG83-TODO: wire system.alarm off these rows (prod). A sibling of the inbound helper (the
    INSERT is 6 lines; the two services have independent lifecycles, so no cross-import)."""
    return system.db.runPrepUpdate(
        "INSERT INTO INV_EDI_ALARM_REJ "
        "(IN_SITE_ID, VC_ALARM_TYPE, VC_SUMMARY, VC_MANIFEST_NUMBER, VC_ASSY_PART_NUMBER, "
        " VC_ERROR_TEXT, IN_EIN, BIT_RESOLVED, DT_RAISED) "
        "VALUES (?, 'RENBAN_COLLISION', ?, ?, ?, ?, NULL, 0, GETDATE())",
        [siteId, (summary or "")[:200], renban, candidatePart,
         (errorText or "")[:200] if errorText is not None else None], db, tx=tx, getKey=True)


def _candidate_base(rows):
    """The breakdown's STARTING count = the lowest candidate renban suffix (the group count at breakdown
    start, before the trailers advanced it). The next-free scan starts HERE (forward-from-the-candidates,
    source-truth §3), NOT from next_count (which already sits PAST the last candidate and would skip the
    free numbers among the candidates themselves)."""
    if not rows:
        return 0
    return min(int(r["renban"][-3:]) for r in rows)


def _emitted_parts(rows):
    """The distinct EMITTED parts (first-seen order) — the DELETE scope (partial-lot parity, see below)."""
    parts_seen = []
    seen = set()
    for r in rows:
        if r["part"] not in seen:
            seen.add(r["part"]); parts_seen.append(r["part"])
    return parts_seen


def _commit_rows_tx(db, log, groupCode, rows, next_count, alarm_id=None, override_note=None,
                    acknowledged=None):
    """The ATOMIC one-transaction write (delete-then-reinsert + counter bump), shared by EVERY commit
    path (no-collision straight-commit, use_next_free, override). RUNS A TOCTOU RE-CHECK *inside* the tx
    against the candidate renbans BEFORE the first INSERT — if any was claimed since the pre-tx check
    (the concurrent Order commit at order/code.py:130 is a non-blank renban writer with no unique
    constraint), the tx is ABORTED and a fresh COLLISION payload is returned (no partial write).

    `acknowledged` (override path): the SET of renbans the operator DELIBERATELY chose to reuse (already
    known in-use at the pre-check). The in-tx re-check excludes these from the abort decision — it only
    aborts on a renban taken NEWLY since the pre-check (a genuine lost race), NOT on the numbers the
    operator explicitly overrode. For use_next_free / straight-commit `acknowledged` is empty, so ANY
    in-use candidate aborts (the chosen-free number must still be free at commit).

    When `alarm_id` is given (a resolution path), the RENBAN_COLLISION alarm is ack'd (BIT_RESOLVED=1) IN
    THE SAME tx, so a rolled-back commit also rolls back the ack (the alarm correctly stays active on
    failure). Returns {"status":"COMMITTED", ...} on success, or {"status":"COLLISION", ...} on a lost race.

    The DELETE scope = the EMITTED parts, NOT the full feed (partial-lot parity, RenbanOrder.pas:417/506):
    a 0<qty<lotqty part lands on NO truck -> emits NO row -> its blank placeholder SURVIVES (deleting it
    with no re-insert = silent order loss)."""
    parts_seen = _emitted_parts(rows)
    ack = set(acknowledged or ())
    tx = system.db.beginTransaction(db)
    try:
        # (0) IN-TX TOCTOU RE-CHECK — covers EVERY path (use_next_free / override / no-collision straight
        # commit). The pre-tx check ran BEFORE this tx; a concurrent writer could have grabbed a chosen
        # renban in the gap. SHOULD-FIX-3: re-test the resident-rows predicate ON THE SAME TX CONNECTION
        # (tx=tx) — the re-check and the INSERTs below now run on ONE connection in ONE transaction (it was
        # previously passed `db` only, hitting a SEPARATE autocommit connection, contra this comment). It
        # sees committed rows from other connections under READ COMMITTED. EXCLUDE the operator-acknowledged
        # renbans (override path) so only a NEWLY-taken number aborts; on a lost race -> abort + return a
        # fresh COLLISION payload so the dialog re-WARNs (it does NOT silently issue a live number). NB this
        # narrows the TOCTOU window to ~0 but does not hold a row lock (no unique constraint on
        # VC_RENBAN_NUMBER — an architect call: a filtered unique index or UPDLOCK,HOLDLOCK would close it
        # fully). Bounded: with ~2 operators/site this almost never loops.
        recheck = [c for c in check_renban_collisions(rows, db, tx=tx) if c["renban"] not in ack]
        if recheck:
            system.db.rollbackTransaction(tx)
            log.warn("renban breakdown: TOCTOU re-check FAILED in-tx (lost race) group=%s renbans=%s — aborting"
                     % (groupCode, [c["renban"] for c in recheck]))
            return {"status": "COLLISION", "collisions": recheck, "rows": rows,
                    "next_count": next_count, "alarm_id": alarm_id, "lost_race": True}

        # (a) DELETE all still-blank placeholders, per EMITTED part (@FRSNumber='' -> blank-renban branch).
        for part in parts_seen:
            system.db.runPrepUpdate(
                "EXEC DELETE_OrderRenban @PartNumber=?, @FRSNumber=?, @RenbanNumber=?",
                [part, "", ""], db, tx=tx)

        # (b) INSERT each trailer-row (qty>0 already filtered by compute). RenbanNum NON-blank; the
        # 7-char FRS is honored as-is (proc recompute is a varchar(7) truncation no-op).
        for r in rows:
            system.db.runPrepUpdate(
                "EXEC INSERT_OpenOrder @SupCode=?, @PartNum=?, @KanbanNum=?, @FRSNum=?, @RenbanNum=?, @Qty=?",
                [r["supplier"], r["part"], r["kanban"], r["frs"], r["renban"], r["qty"]], db, tx=tx)

        # (c) advance the group's roll-over counter — P4 CLEAN WRAP (% 1000); see the design notes above.
        if next_count is not None:
            persisted_count = "%03d" % (next_count % 1000)
            system.db.runPrepUpdate(
                "EXEC UPDATE_RenbanGroupCount @RenbanCode=?, @RenbanCount=?",
                [groupCode, persisted_count], db, tx=tx)

        # (d) ACK the RENBAN_COLLISION alarm IN THE SAME TX (resolution paths only). A rolled-back commit
        # rolls back the ack -> the alarm stays active if the write fails. override stamps an audit note.
        if alarm_id is not None:
            if override_note is not None:
                system.db.runPrepUpdate(
                    "UPDATE INV_EDI_ALARM_REJ SET BIT_RESOLVED=1, VC_ERROR_TEXT=? WHERE IN_ALARM_ID=?",
                    [override_note[:200], alarm_id], db, tx=tx)
            else:
                system.db.runPrepUpdate(
                    "UPDATE INV_EDI_ALARM_REJ SET BIT_RESOLVED=1 WHERE IN_ALARM_ID=?",
                    [alarm_id], db, tx=tx)

        system.db.commitTransaction(tx)
        log.info("renban breakdown committed: group=%s parts=%d trailer_rows=%d next_count=%s alarm_ack=%s"
                 % (groupCode, len(parts_seen), len(rows), next_count, alarm_id))
    except:
        system.db.rollbackTransaction(tx)
        raise
    finally:
        system.db.closeTransaction(tx)
    return {"status": "COMMITTED", "rows": rows, "next_count": next_count}


def commit_renban_breakdown(groupCode, trailers, palletsPerTrailer, database=None, _orders=None,
                            resolution=None, siteId=None, actor=None, session=None):
    """The renban breakdown write-back, COLLISION-AWARE (P4 WARN->GUIDE->FIX). Loads the blank-renban
    orders, computes the trailer distribution, checks each candidate renban against the resident-rows
    in-use predicate, then commits the delete-then-reinsert + counter bump in ONE atomic tx.

    BLOCKER-2 / SERVER-SIDE WRITE GATE (code-reviewer; the R25/P15 hole class): this driver WRITES
    (INV_OPEN_ORDER_INF delete-then-reinsert + UPDATE_RenbanGroupCount + the alarm ack). Every write
    MUST be authorized SERVER-SIDE from the SESSION's roles, not a client prop (the legacy H3 hole). We
    call auth.requireWrite(session) FIRST — BEFORE loadBlankRenbanOrders / compute / any read-only
    collision check / any write — so a forged client prop or an anonymous session is rejected on ALL
    THREE resolution branches (None / use_next_free / override) with no read, no compute, no write. The
    sibling RenbanGroup master Save is already session-gated; this matches it. `session` is the Perspective
    SESSION (self.session), threaded from the view button scripts. (A None session resolves to no roles ->
    AuthError, so an un-threaded caller fails CLOSED.) The headless e2e passes a write-role session shape.

    `resolution` (the FIX choice; backward-compatible — positional e2e calls pass None):
      None
          Compute + check_renban_collisions (pre-tx, read-only).
          * No collisions  -> straight commit (today's behavior), {"status":"COMMITTED", ...}. The in-tx
            re-check still runs (TOCTOU) and can return COLLISION on a lost race.
          * Collisions     -> WRITE a RENBAN_COLLISION alarm (WARN, autocommit) + the next-free RUN-of-N
            base (GUIDE) and RETURN WITHOUT WRITING the breakdown:
              {"status":"COLLISION", "collisions":[...], "rows":[...], "next_count":N,
               "next_free":<base or None>, "alarm_id":<id>}
      {"action":"use_next_free", "base":<int>, "alarm_id":<id>}
          Re-map rows onto `base` (PURE), then commit + in-tx re-check + ack the alarm. A lost race
          returns a fresh COLLISION payload (re-GUIDE).
      {"action":"override", "alarm_id":<id>}
          Commit the ORIGINAL candidate rows (operator confirmed) + in-tx re-check + ack the alarm with
          an audit note (VC_ERROR_TEXT = "OVERRIDE by <actor> <ts>"). A lost race still aborts (the
          re-check is on EVERY path) — override does NOT bypass a number that was JUST taken.
      (cancel is purely CLIENT-SIDE: no driver call; the alarm row stays BIT_RESOLVED=0 / active.)

    Atomicity (RenbanOrder.pas:416/436/473): delete-then-reinsert + counter bump are ONE tx; the human
    decision + the alarm WARN happen OUTSIDE the tx; only the resolution's commit + alarm-ack are in-tx.
    """
    db = database if database is not None else DATABASE
    log = system.util.getLogger("SPIKE.renban")        # IG-DEBUG (keep: low-volume, operator action)
    # SERVER-SIDE WRITE GATE (BLOCKER-2): authorize FROM THE SESSION before any read/compute/write, on
    # EVERY path. requireWrite reads session.props.auth.user.roles (server-side) + raises AuthError on a
    # missing write role; a None/anon/forged-prop session fails CLOSED here (no DB touched).
    import auth as A
    A.requireWrite(session)
    orders = _orders if _orders is not None else loadBlankRenbanOrders(groupCode, db)
    if not orders:
        # mirror the legacy "No records for this Renban Group" (:648) — nothing to do, do not write.
        log.info("renban breakdown: no blank-renban orders for group %s — nothing to commit" % groupCode)
        return {"status": "EMPTY", "rows": [], "truck_counts": [], "next_count": None, "total_lots": 0}

    result = compute_trailer_breakdown(orders, trailers, palletsPerTrailer, groupCode)
    rows = result["rows"]
    next_count = result["next_count"]

    # ---- resolution paths (the operator FIX) ------------------------------------------------
    if isinstance(resolution, dict) and resolution.get("action") == "use_next_free":
        base = int(resolution["base"])
        alarm_id = resolution.get("alarm_id")
        new_rows, new_next = _remap_rows_onto_base(rows, groupCode, base)
        log.info("renban FIX use_next_free: group=%s base=%03d rows=%d" % (groupCode, base, len(new_rows)))
        committed = _commit_rows_tx(db, log, groupCode, new_rows, new_next, alarm_id=alarm_id)
        committed["truck_counts"] = result["truck_counts"]
        committed["total_lots"] = result["total_lots"]
        return committed

    if isinstance(resolution, dict) and resolution.get("action") == "override":
        alarm_id = resolution.get("alarm_id")
        ts = system.date.format(system.date.now(), "yyyy-MM-dd HH:mm:ss")
        note = "OVERRIDE by %s %s (renban reused by operator decision)" % (actor or "operator", ts)
        # ACKNOWLEDGED = ONLY the renbans that were COLLIDING when the operator saw the WARN (the ones they
        # explicitly chose to reuse). The CONTRACT (SHOULD-FIX-2, sql-adversary): the dialog MUST pass this
        # set back from the COLLISION payload (gen_renban_views.py's override script does:
        # acknowledged=list(p.acknowledged)). The in-tx re-check then aborts on a renban that was FREE at the
        # WARN but got taken NEWLY since (a genuine lost race on a number the operator never agreed to
        # reuse) — it does NOT abort on the numbers being deliberately overridden.
        #
        # When `acknowledged` is ABSENT (None) we now default to the EMPTY set — acknowledge NOTHING unless
        # the dialog names it. The old behavior recomputed check_renban_collisions(rows) AT OVERRIDE TIME,
        # which auto-acked any number grabbed by a concurrent writer BETWEEN the WARN and the override click
        # (the operator silently reused a number they never saw colliding). Empty-default fails CLOSED: an
        # override with no named set re-checks ALL candidates and aborts on ANY in-use number — so a missing
        # contract can never silently reuse an un-seen number. (Acknowledging ALL candidate rows would make
        # override never abort; the seen-colliding set is what keeps the TOCTOU guard meaningful.)
        acknowledged = resolution.get("acknowledged")
        acknowledged = set() if acknowledged is None else set(acknowledged)
        log.warn("renban FIX override: group=%s rows=%d acknowledged=%s note=%s"
                 % (groupCode, len(rows), sorted(acknowledged), note))
        committed = _commit_rows_tx(db, log, groupCode, rows, next_count, alarm_id=alarm_id,
                                    override_note=note, acknowledged=acknowledged)
        committed["truck_counts"] = result["truck_counts"]
        committed["total_lots"] = result["total_lots"]
        return committed

    # ---- resolution is None: the PRE-COMMIT GATE (read-only check, outside any tx) -----------
    collisions = check_renban_collisions(rows, db)
    if collisions:
        # WARN: write a RENBAN_COLLISION alarm row (autocommit — the home-hub flag persists even if the
        # operator walks away) + GUIDE: the next-free RUN-of-N base (scan FROM the breakdown's starting
        # count, so the run can include free candidate numbers).
        n = len(set(r["renban"] for r in rows))
        next_free = next_free_run(groupCode, _candidate_base(rows), n, db)
        c0 = collisions[0]
        summary = "Renban %s collides with open release (FRS %s)" % (c0["renban"], c0["open_frs"])
        err = "open part %s FRS %s, %d resident row(s), last %s; %d candidate(s) in use" % (
            c0["open_part"], c0["open_frs"], c0["rows_open"], c0["last_add"], len(collisions))
        alarm_id = _write_renban_alarm(db, siteId, summary, c0["renban"], c0["candidate_part"], err)
        log.warn("renban COLLISION group=%s collisions=%s next_free=%s alarm_id=%s"
                 % (groupCode, [c["renban"] for c in collisions], next_free, alarm_id))
        return {"status": "COLLISION", "collisions": collisions, "rows": rows, "next_count": next_count,
                "next_free": next_free, "alarm_id": alarm_id,
                "truck_counts": result["truck_counts"], "total_lots": result["total_lots"]}

    # no collision -> straight commit (today's behavior). The in-tx re-check still guards the TOCTOU race.
    committed = _commit_rows_tx(db, log, groupCode, rows, next_count)
    if committed["status"] == "COLLISION":
        # a concurrent writer claimed a candidate between the pre-check and the tx: WARN + GUIDE now.
        n = len(set(r["renban"] for r in rows))
        next_free = next_free_run(groupCode, _candidate_base(rows), n, db)
        c0 = committed["collisions"][0]
        summary = "Renban %s collides with open release (FRS %s) [lost race]" % (c0["renban"], c0["open_frs"])
        err = "lost race: candidate taken between pre-check and commit; %d in use" % len(committed["collisions"])
        alarm_id = _write_renban_alarm(db, siteId, summary, c0["renban"], c0["candidate_part"], err)
        committed["next_free"] = next_free
        committed["alarm_id"] = alarm_id
    committed["truck_counts"] = result["truck_counts"]
    committed["total_lots"] = result["total_lots"]
    return committed
