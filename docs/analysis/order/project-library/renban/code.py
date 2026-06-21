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
    renban_seed is group_code||group_count (e.g. 'CMWA288'); the 3-digit tail is the seed count."""
    seed3 = int(str(renban_seed)[-3:])               # rightstr(...,3) :775
    rcount = seed3 + truck_number                     # :777 (0-based truck index)
    return group_code + ("%03d" % rcount), rcount     # :779


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


def commit_renban_breakdown(groupCode, trailers, palletsPerTrailer, database=None, _orders=None):
    """The renban breakdown write-back. Loads the blank-renban orders for `groupCode`, computes the
    trailer distribution, then in ONE transaction: per part DELETE_OrderRenban(@FRS='',@Renban='')
    (clear ALL still-blank placeholders), INSERT_OpenOrder per trailer-row with qty>0 (assigned renban,
    full 7-char FRS), and once UPDATE_RenbanGroupCount(next_count). Rollback on any failure.

    Returns the compute result dict (rows + next_count + truck_counts + total_lots).
    `_orders` is a test seam: pass a pre-loaded order list to skip the DB read (the compute is still
    real); production passes None -> loadBlankRenbanOrders.

    Atomicity (RenbanOrder.pas:416/436/473): the delete-then-reinsert + the counter bump are ONE
    transaction; a mid-loop failure must not leave a part with its placeholder deleted but no grouped
    rows (the exact failure that motivated delete-all-then-reinsert, :489-492).
    """
    db = database if database is not None else DATABASE
    log = system.util.getLogger("SPIKE.renban")        # IG-DEBUG (keep: low-volume, operator action)
    orders = _orders if _orders is not None else loadBlankRenbanOrders(groupCode, db)
    if not orders:
        # mirror the legacy "No records for this Renban Group" (:648) — nothing to do, do not write.
        log.info("renban breakdown: no blank-renban orders for group %s — nothing to commit" % groupCode)
        return {"rows": [], "truck_counts": [], "next_count": None, "total_lots": 0}

    result = compute_trailer_breakdown(orders, trailers, palletsPerTrailer, groupCode)
    rows = result["rows"]
    next_count = result["next_count"]

    # the distinct parts to clear. CRITICAL (partial-lot parity): derive this from the EMITTED trailer
    # rows, NOT the full loaded feed. The legacy commit loop iterates ONLY the grid rows
    # (RenbanOrder.pas:417 `for i:=1 to fAvailableCount`, fAvailableCount = grid RowCount-1, :799), and
    # NewFRSOrder deletes by the EMITTED row's part (`@PartNumber := AvailableGrid.Cells[2,Row]`, :506).
    # A part with 0 < qty < lotqty has lots = qty div lotqty = 0 -> lands on NO truck -> emits NO grid
    # row -> DELETE_OrderRenban is NEVER called for it -> its blank-renban placeholder SURVIVES (it waits
    # for a future breakdown once its qty grows, or for manual handling). Deriving parts_seen from the
    # full feed instead would DELETE that placeholder with no re-insert -> silent order loss.
    # DELETE_OrderRenban is part-wide on blank renban, so one call per distinct emitted part suffices
    # (repeat calls for a part with multiple trailer-rows are benign no-ops). Preserve first-seen order.
    parts_seen = []
    seen = set()
    for r in rows:
        if r["part"] not in seen:
            seen.add(r["part"]); parts_seen.append(r["part"])

    tx = system.db.beginTransaction(db)
    try:
        # (a) DELETE all still-blank placeholders, per part (@FRSNumber='' -> part+blank-renban branch).
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

        # (c) advance the group's roll-over counter. PARALLEL-RUN PARITY: reproduce the legacy's
        # persisted value EXACTLY. Legacy sends Format('%.3d',[next_count]) -- a MINIMUM-width zero-pad
        # that NEVER caps (1002 -> '1002', 4 chars) -- into @RenbanCount varchar(3), which the proc
        # LEFT-TRUNCATES to the first 3 chars ('1002' -> '100'). So the reduction is str(N)[:3], NOT
        # N % 1000. ('%03d' % N)[:3] reproduces it: 1002->'100', 1000->'100', 634->'634', 5->'005'.
        # PROVEN on the live proc (mssql-spike, rolled back): EXEC ...@RenbanCount='1002' -> stored '100'.
        # IG83-TODO: the rollover itself is a LATENT LEGACY BUG -- at next_count>=1000 the persisted
        #   count collapses to '100' (1000->'100', 1002->'100'), so the NEXT run's renban numbers COLLIDE
        #   with the earlier CMWA100x block. We faithfully reproduce it for parallel-run parity; do NOT
        #   silently "fix" the wrap here. POST-CUTOVER fix (like the GetShip-calendar carry): widen the
        #   count column / param, or alert+block the operator at 999 before the counter can roll. See
        #   renban-breakdown-spec.md §12.7 (rollover-latent-bug carry).
        # The renban NUMBER itself (group_code + '%03d' % rcount, into VC_RENBAN_NUMBER varchar(8)) is
        #   UNAFFECTED -- CMWA1002 is 8 chars and fits; only this persisted varchar(3) COUNT truncates.
        if next_count is not None:
            system.db.runPrepUpdate(
                "EXEC UPDATE_RenbanGroupCount @RenbanCode=?, @RenbanCount=?",
                [groupCode, ("%03d" % next_count)[:3]], db, tx=tx)

        system.db.commitTransaction(tx)
        log.info("renban breakdown committed: group=%s parts=%d trailer_rows=%d next_count=%s"
                 % (groupCode, len(parts_seen), len(rows), next_count))
    except:
        system.db.rollbackTransaction(tx)
        raise
    finally:
        system.db.closeTransaction(tx)
    return result
