# Live-Session Test Plan — `Order/OrderSpike` (faithful 4-row ledger)

**Target:** Ignition 8.1.52 gateway, project `spike`, view `Order/OrderSpike`, route
`/order` (confirmed in `page-config/config.json`). DB connection `Inventory_Spike` →
database `Inventory` (container `mssql-spike`).
**Anchor:** Today=2026-06-15, Line=COROLLA, FillDays=25.
**Why a human runs this:** the rebuilt `custom.gridModel` script transform only
*deserializes* on disk — its binding fires **only when a session opens the view**.
The code-reviewer confirmed the pooling math reconciles to the proc cell-for-cell but
flagged that the transform has **never executed at runtime**. These steps drive that
execution and assert on the `SPIKE` log facts, not on eyeballing.

**Scope:** parity/behavior of the rendered ledger + edit-lock + live recompute.
Commit path is DEFERRED/untouched (Commit button is disabled by design).
**Non-destructive:** no DB writes (typed orders live in `view.custom.editQty` only;
commit is deferred). No fixtures to restore. The `Reset Typed Orders` button is the
only "undo" needed and is part of the test.

> Already verified by QA without a browser (do not re-run, context only):
> - **Render-parity 14/20 groups PASS cell-for-cell**, order-by/peach **22/22**, against
>   the golden ledger, using a harness that replicates the *view's own* PAB recompute
>   (`/tmp/view_render_parity.py`). The 6 mismatches are the tracked out-of-scope gaps
>   (FILM ×4 R1, WHEEL M1 R3, WHEEL 17D1 R2 demo edit) — see the QA report.
> - TIRE 18DL anchors: pooled Beg[0]=47216, End[0]=47885, Usage[0]=706; DUNLOP receipts
>   1375@fp0,fp2 + peach fp5 (06-22); MICHELIN 1320@fp3 + peach fp6 (06-23).
> - VALVE RV safety=9220, End goes red at fp20–24 (8993,7757,6521,5285,4261).
> - No OrderSpike deserialize error in `wrapper.log` (the one stale error is for
>   `PartsStockMaster/List`, 2026-06-12, unrelated).

---

## Pre-state (run once, before opening the session)

**P0. Reload resources and confirm OrderSpike parses (QA could not run `gwcmd -r`; you run it).**
```bash
/usr/local/ignition/gwcmd.sh -r
# wait ~3s for the reload, then assert no NEW deserialize error for OrderSpike:
grep "Unable to deserialize" /usr/local/ignition/logs/wrapper.log | grep -i OrderSpike
```
- **Expected:** the grep prints **nothing** (OrderSpike deserialized clean).
- **Fail:** any line mentioning `Order/OrderSpike` → stop, hand to ignition-developer.
  (A pre-existing line for `PartsStockMaster/List` is fine — ignore it.)

**P1. Mark the log so you only read THIS run's lines.** Note the current end of the log:
```bash
wc -l /usr/local/ignition/logs/wrapper.log
```
Remember that number `N`. In every step below, read only new lines with:
```bash
tail -n +$((N+1)) /usr/local/ignition/logs/wrapper.log | grep "SPIKE"
```
(or just `grep "SPIKE" /usr/local/ignition/logs/wrapper.log | tail -20` for the latest.)

**P2. Pre-state guard (anti-affirming-the-consequent).** The transform has not run yet,
so there must be **no** `SPIKE grid:` line from this session:
```bash
grep "SPIKE grid:" /usr/local/ignition/logs/wrapper.log | tail -1
```
- **Expected:** either nothing, or only lines older than `N`. If a fresh `SPIKE grid:`
  line already exists before you click anything, a prior session is feeding the log —
  close other sessions/tabs so your clicks are attributable.

---

## Test 1 — Transform RUNS, and TIRE 18DL renders as the faithful pooled ledger

This is the headline assertion: the rebuilt transform executes at runtime.

**Steps**
1. Open a browser to **`http://localhost:8088/data/perspective/client/spike/order`**.
   (Do NOT type the `view.json` path; navigate to the *session* URL above.)
2. In the **SelectOrderBar**: confirm **Line = COROLLA**, set **Part Type = TIRE**,
   **Fill Days = 25**. Today shows **2026-06-15** (read-only label).
3. Click the blue **Simulate** button once.
4. In a terminal:
   ```bash
   grep "SPIKE grid:" /usr/local/ignition/logs/wrapper.log | tail -1
   ```

**Expected result (assert all)**
- The grep prints a line of the form:
  `SPIKE grid: A=25 B=26 C=250 rows (line=COROLLA type=TIRE fd=25 edits=0)`
  - `edits=0` **proves the rebuilt transform ran with a clean (un-typed) overlay** —
    this is the runtime execution the code-reviewer flagged as never-before-exercised.
  - `A=25` (25 production day columns), `type=TIRE`, `fd=25`.
- In the **PhasedGrid** table, the **18DL** block renders as exactly:
  - **Beg Balance** row, **two Receipts rows** (Brand/Supplier = DUNLOP `4265202S1000`
    and MICHELIN `4265202S2000`), **Usage** row, **End Balance** row, then a gray
    separator row.
  - Day-0 column (`Mon 06-15`): **Beg = 47216, Usage = 706, End = 47885.**
  - DUNLOP Receipts: **1375** under `06-15` and `06-17`; the **06-22** cell is the peach
    editable cell (blank/0 qty).
  - MICHELIN Receipts: **1320** under `06-18`; the **06-23** cell is peach.

**Pass/Fail**
- PASS if the `SPIKE grid:` line appears with `edits=0` AND the 18DL numbers above match.
- FAIL → what it tells you: no `SPIKE grid:` line = the binding/transform did not fire
  (route, binding, or DB connection issue → ignition-developer). Line present but numbers
  wrong = a runtime pooling bug the static reconcile missed (re-open this with QA + the
  `/tmp/view_render_parity.py` baseline). It does NOT by itself implicate the proc — the
  proc output is already parity-validated.

---

## Test 2 — Ledger is numbers-only / color-only (no glyph clutter)

Asserts NEW-behavior requirement #1 (drop `★ [LT] [OT] 🚚 📦 ⚠`), color-only semantics.

**Steps (visual, on the TIRE grid from Test 1)**
1. Scan every data cell in the 18DL block (and a few other groups).
2. Note where peach (`#FFCC99`) and red (`#FF0000`) appear.

**Expected result**
- **No glyphs anywhere** — every value cell contains a number or is blank. No
  `★`, `[LT]`, `[OT]`, 🚚, 📦, ⚠.
- **Peach** appears on **exactly two cells in 18DL**: DUNLOP `06-22` and MICHELIN `06-23`
  (each supplier's single order-by cell). No other peach in the block.
- **Red** font appears on **no** TIRE End cell (every tire group has safety_stock=0, so
  End is never below safety — verified: TIRE 15D safety=0).
- To see red, switch **Part Type → VALVE**, click **Simulate**, find the **RV** group:
  its **End Balance** row is red on the last five day columns (~`07-20`…`07-24`,
  values 8993, 7757, 6521, 5285, 4261 — all below safety 9220). TPMSS is never red.

**Pass/Fail**
- PASS if no glyphs, peach only on the two 18DL order-by cells, and VALVE RV End reds at
  fp20–24 while no tire group reds.
- FAIL → glyphs present = wrong/old transform loaded (reload, re-check Test 1). Peach on a
  non-order-by cell = `orderby_col_index` placement bug (but QA validated 22/22 statically,
  so suspect a stale view) → ignition-developer. Red where safety=0 = the `safety>0` guard
  regressed.

---

## Test 3 — Type into the peach cell → ACCEPTED, End recomputes forward, red re-evaluates

Settles **RISK-3 (qualified-value-access)** and requirement #4 (live recompute / row-23).
Use VALVE RV (the red demonstrator) so the recompute visibly clears red.

**Steps**
1. With **Part Type = VALVE** loaded (Test 2), locate the **RV** group's single Receipts
   row and its **peach** order-by cell (the one peach cell in that supplier's row).
2. Double-click the peach cell, type a **large qty** (e.g. **8000**), press Enter to commit.
3. Read the log:
   ```bash
   grep "SPIKE edit" /usr/local/ignition/logs/wrapper.log | tail -2
   grep "SPIKE grid:" /usr/local/ignition/logs/wrapper.log | tail -1
   ```

**Expected result (assert all)**
- A line: `SPIKE edit ACCEPTED: 900804500600|<fp> = 8000 (nonce=…, 1 total edits)`
  - The committed cell is RV's order-by `fill_pos` for part `900804500600`.
- Immediately after, a new `SPIKE grid: … edits=1` line (the `editNonce` bump re-ran the
  transform with the typed overlay folded in).
- In the grid: RV's **Receipts** peach cell now shows **8000**; the **End Balance** row
  recomputes **forward from that day** (every End at/after the order-by fill_pos increases
  by ~8000); the previously-red End cells at fp20–24 turn **black** (now ≥ safety 9220).
  Earlier-day End cells (before the order-by day) are unchanged.

**Pass/Fail**
- PASS if `SPIKE edit ACCEPTED` + `edits=1` grid re-run + forward End recompute + red
  cleared. This proves the qualified-value-access path (reading `view.custom.editQty`,
  the hidden `part`/`_obfp` fields) works at runtime — RISK-3 settled.
- FAIL → no `SPIKE edit ACCEPTED` (commit handler didn't fire or rejected a valid cell);
  ACCEPTED but no `edits=1` grid line (nonce→binding chain broken); End didn't move
  (pooled recompute not folding `editQty`). Each → ignition-developer with the exact line.

---

## Test 4 — Non-peach day cell is LOCKED (edit rejected)

Settles requirement #2 (order entry locked to the one valid cell per supplier).

**Steps**
1. Still on a loaded grid (TIRE or VALVE), pick a Receipts row and a **non-peach** day
   cell (any day that is NOT that supplier's order-by date — e.g. DUNLOP's `06-15` cell).
2. Try to edit it: double-click and attempt to type a number, press Enter.
3. Read the log:
   ```bash
   grep "SPIKE edit" /usr/local/ignition/logs/wrapper.log | tail -3
   ```

**Expected result**
- Preferred: the cell is **not editable** — double-click does nothing, no edit cursor
  (the builder set `editable:false` on every non-order-by cell).
- If the table still routes a commit (column is editable at column level), the handler
  rejects it: a line `SPIKE edit REJECTED (locked): row=… col=d<fp> label=Receipts obfp=…`
  appears, and **NO** `SPIKE edit ACCEPTED` follows for that cell. The cell value does not
  persist into `editQty` (no `edits` count increase on the next `SPIKE grid:` line).
- Editing a **non-Receipts** row (Beg/Usage/End) similarly yields `REJECTED (locked)` or
  is non-editable.

**Pass/Fail**
- PASS if the non-peach cell cannot be saved AND (if a commit fired) `REJECTED (locked)`
  logged with no matching ACCEPTED. FAIL → an `ACCEPTED` on a non-order-by cell = the lock
  leaks (the `int(obfp) != fp` guard is wrong) → ignition-developer.

---

## Test 5 — Reset Typed Orders clears the overlay

**Steps**
1. After Test 3 (RV has a typed 8000), click the gray **Reset Typed Orders** button.
2. Read the log:
   ```bash
   grep "SPIKE" /usr/local/ignition/logs/wrapper.log | tail -3
   ```

**Expected result**
- A line `Reset typed orders -> nonce=…`, followed by a `SPIKE grid: … edits=0` line.
- In the grid: RV's peach cell returns to blank/0; the End Balance row returns to the
  baseline (red again at fp20–24); no typed value survives.

**Pass/Fail**
- PASS if `edits=0` after reset and the ledger returns to baseline. FAIL → overlay not
  cleared (`editQty` not emptied or nonce not bumped) → ignition-developer.

---

## Teardown
- None required — no DB rows or files were mutated (commit deferred; typed orders are
  session-local in `view.custom.editQty`). Test 5 already clears the overlay. Closing the
  session tab discards all state. Leave the gateway and `mssql-spike` running.

## Out-of-scope reminders (do NOT fail the view for these)
- FILM (all 4 colors): Usage=0 / flat balance — tracked proc gap **R1** (forecast
  week-number mapping). The view renders what the proc returns.
- WHEEL M1 (`4261102Q8000`): receipt overcount — tracked proc gap **R3**.
- WHEEL 17D1 day5=25: David's manual row-23 demo edit (**R2**), not a bug.
- These affect numbers in those specific groups, not the layout/lock/recompute behavior
  under test here.
