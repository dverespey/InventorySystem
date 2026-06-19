# Adversarial pre-cutover review — stock-ledger / D6 / Order cutover

Reviewer stance: refute the cutover plan. Findings ranked by blast radius (plant-floor IN_QTY
corruption / mis-bill first). Each: severity, claim, evidence (file:line or scenario), NEW carry vs
flaw in an existing carry. Verdict at the end.

Ground truth read: `cutover-runbook.md`; the 4 producers + `order` lib; `spike-post-stockmovement-proc.sql`,
`spike-seed-opening-balance.sql`, `spike-stock-ledger-table.sql`, `spike-report-procs-d6.sql`; the recon
harnesses; `DB Schema/CreateInventory.sql` (UTF-16LE, decoded).

---

## BLOCKER 1 — The `UPDATE_PartNumber` trigger is a 13th, un-retired live writer that double-writes the audit ledger on every ledger post (and the runbook never mentions it)

**Claim under test:** "Cutover = flip reads + DROP the 12 legacy qty-triggers. Seams presume triggers gone."
The runbook (§5) and every producer docstring assume exactly twelve triggers touch on-hand and that
dropping them makes the ledger the sole owner of `IN_QTY`.

**Evidence — there is a 13th trigger on `INV_PARTS_STOCK_MST` that fires on EVERY update of it:**
`CreateInventory.sql:4095` `CREATE TRIGGER [dbo].[UPDATE_PartNumber] ON [dbo].[INV_PARTS_STOCK_MST] FOR UPDATE`.
Its body (4104–4116) does, on `@numrows>0`:
- `INSERT into INV_PARTS_STOCK_MST_HIST SELECT * from deleted` (full-row snapshot), and
- `INSERT INV_PART_QTY_INF(...) SELECT i.VC_PART_NUMBER, d.IN_QTY-i.IN_QTY, i.IN_QTY, 'U', i.VC_LAST_UPDATE FROM inserted i, deleted d WHERE i.IN_QTY <> d.IN_QTY`.

`POST_StockMovement` mutates on-hand with `UPDATE dbo.INV_PARTS_STOCK_MST SET IN_QTY = IN_QTY + @delta, VC_LAST_UPDATE = @ts WHERE IN_PART_ID = @partId`
(`spike-post-stockmovement-proc.sql:71-74`). That UPDATE **fires `UPDATE_PartNumber`**. So after cutover,
*every forward ledger post* also writes a row to `INV_PART_QTY_INF` and a full-row snapshot to
`INV_PARTS_STOCK_MST_HIST` — the exact legacy audit path the design says the ledger *replaces*
(`IGNITION-stock-ledger-design.md:81` "formalizes / replaces `INV_PART_QTY_INF`"; :108 "We do NOT
migrate or rely on" those rows).

**This is not theoretical — the harness already had to work around it.** `test_ledger_opening_balance.py:78`
must `DISABLE TRIGGER UPDATE_PartNumber ON INV_PARTS_STOCK_MST` before any forward `post()`, with the
comment "forward posts bump IN_QTY -> the legacy UPDATE_PartNumber audit trigger would fire; suppress it."
The test only passes because it disables the trigger. **Production cutover has no such DISABLE step in the
runbook.**

**Why it's a BLOCKER, not a NIT:**
1. Post-cutover the ledger is supposed to be the *single* source of stock-movement truth, but
   `INV_PART_QTY_INF` keeps silently accumulating a parallel, partial audit (`'U'` rows only, change =
   `d.IN_QTY - i.IN_QTY` = `-@delta`). Two ledgers drift in meaning; reports/queries pointed at the old
   audit table will be half-right.
2. `INSERT ... SELECT * from deleted` into `INV_PARTS_STOCK_MST_HIST` runs on **every post** — a hot path
   for a shipment burst, writing a full wide-row snapshot per single-part stock move. Unbounded growth +
   write amplification on the plant-floor write path.
3. Most dangerous: `DELETE_AutoPurge` deletes `INV_PARTS_STOCK_MST_HIST` rows by `VC_ADD <= cutoff`
   (`design:153`). With every post now generating HIST rows, the purge surface and timing assumptions of
   §3.1 change.

**The runbook is silent.** `grep UPDATE_PartNumber docs/analysis/cutover-runbook.md` → nothing;
`grep PART_QTY_INF` → nothing. The "12 triggers" enumeration (`design:138-140`) explicitly counts only
RecConfStat/Reject/Stocktaking/PartShipping and folds `DeleteShipDate` in as the 13th — it never accounts
for `UPDATE_PartNumber`, which is a *different* trigger on a *different* table that the ledger's own write
fires.

**Direction of fix (bounce to ignition-architect + delphi-architect):** decide explicitly whether
`UPDATE_PartNumber` is dropped, or split (keep the rename-cascade legs at `@numrows=1`, drop the
`INV_PART_QTY_INF`/HIST qty-audit legs), or whether `POST_StockMovement` should suppress it
(`SET CONTEXT_INFO` guard / session flag). This MUST be a named carry with a cutover step.

**Status:** NEW carry (absent from the runbook's 9).

---

## BLOCKER 2 — Two contradictory backfill definitions; the shipped one (`SEED_AllOpeningBalances`) silently DISCARDS the D8(3)/D12#3/F3/F5 bug-fix corrections the design promised at cutover

**Claim under test:** Runbook §4 — "Run `SEED_AllOpeningBalances` ONCE at cutover … After it,
`SUM(ledger) == IN_QTY` for every part … this opening balance + forward parity IS the sign-off."

**The contradiction.** The stock-ledger design §9 defines the cutover backfill differently:
`IGNITION-stock-ledger-design.md:357` "**Backfill:** at cutover, seed the ledger by **deriving one
movement per live source row** (the §6 derivation, run for-real once) so the ledger's `SUM` equals the
cutover `IN_QTY` exactly." And §9 immediately after (the "Bug fixes that intentionally change on-hand at
cutover" para, ~:360) says the backfill will **produce the *corrected* on-hand** for D8(3)/D12#3/F3/F5 and
"Reconcile the delta explicitly in the cutover runbook."

The shipped proc does the opposite. `SEED_AllOpeningBalances` (`spike-seed-opening-balance.sql:85-91`)
records ONE `OPENING_BALANCE` row per part = the part's **current legacy `IN_QTY`**, by construction:
`SELECT p.IN_PART_ID, p.IN_QTY, ... FROM dbo.INV_PARTS_STOCK_MST p`. So `SUM(ledger) == IN_QTY` is true
**tautologically** — it copies the legacy (buggy) balance forward verbatim.

**These are not the same plan and they produce different cutover balances.** For any part the design
flagged as "ledger correct, legacy buggy" (D8(3) arrival-reversal overstated; D12#3 yard under-count; F3
multi-row under-count; F5 part-change no-op), the **derive-per-source-row** backfill writes the
*corrected* on-hand, while the shipped **opening-balance** backfill writes the *uncorrected legacy*
on-hand. The whole §6 parity apparatus — four EXPECTED-DIVERGENT classes, the GO/NO-GO that "data
adjudicates" — exists to prove the corrected number is right, and then the shipped cutover throws that
corrected number away.

**This is the harness lying to itself.** `test_ledger_opening_balance.py:138` asserts "after backfill,
ledger SUM == IN_QTY for ALL parts (0 drift) — cutover parity closed." That check is **vacuously true** by
construction (it copied IN_QTY in) — it proves nothing about correctness. This is precisely the
`feedback-parity-fixture-fidelity` failure mode the project already burned itself on: green parity that
isn't faithfulness.

**Decide which plan is real.** Either:
- (a) opening-balance copy is the intended cutover and the D8(3)/D12#3/F3/F5 "corrections at cutover"
  language in design §9 is dead/aspirational — then say so, retire the four divergence classes from the
  cutover scope, and stop calling the all-parts opening-balance check a parity sign-off; or
- (b) the corrections are real, and the cutover must run the derive-per-source-row backfill (design §9 /
  the §6 derivation) for the parts with surviving source rows and only fall back to opening-balance for
  the purge-aged parts — a hybrid the runbook does not describe.

Note the harness honesty caveat compounds this: `test_ledger_fullhistory_recon.py:7` and
`test_stock_ledger_parity.py` banner both state the restored `.bak` is post-purge, so the
derive-per-source-row reconstruction has **never been run against real surviving history that equals
`IN_QTY`** — it is proven only on a hand-built controlled part (PART 16). So plan (b) is also unproven on
real data.

**Direction of fix (bounce to ignition-architect):** reconcile design §9 vs runbook §4 + the shipped
proc; pick one backfill semantics; if corrections are in scope, the runbook needs the explicit
delta-reconciliation step §9 itself demands and it is currently missing.

**Status:** flaw in existing carry §4 (its premise contradicts design §9 and the divergence-class work).

---

## SHOULD-FIX 3 — A 4th/5th direct writer of `INV_PARTS_STOCK_MST.IN_QTY` survives cutover: `UPDATE_PartsStockInfo` (absolute set) and `UPDATE_PartsStockInfoCount` (additive)

**Claim under test:** "POST_StockMovement — the ONLY writer of INV_PARTS_STOCK_MST.IN_QTY"
(`spike-post-stockmovement-proc.sql:2`); the producers are the 4 writers.

**Evidence — grep of `IN_QTY =` against `INV_PARTS_STOCK_MST` in the live dump finds non-trigger procs:**
- `UPDATE_PartsStockInfo` (`CreateInventory.sql:~5747`) — the PartsStockMaster edit proc —
  `UPDATE INV_PARTS_STOCK_MST SET ... IN_QTY=@QTY ... WHERE IN_PART_ID=@PartID`. An **absolute overwrite**,
  not additive. It is **live** (called from `DataModule.pas:1482` `'dbo.UPDATE_PartsStockInfo;1'`), and
  was carried into the rebuilt master-data module (`ignition-spike-log.md:104`).
- `UPDATE_PartsStockInfoCount` (`CreateInventory.sql:~4086`) — `SET IN_QTY = IN_QTY-@QTY ... WHERE
  VC_PART_NUMBER = @PartNumber`. The spec itself flags it: `parts-stock-master.md:183` "**A second writer
  of `IN_QTY` keyed by number** … confirms `IN_QTY` is mutated from multiple places."

**Why it matters at cutover.** After the ledger owns `IN_QTY`, any path that still does an *absolute*
`SET IN_QTY=@QTY` (or an additive `-@QTY`) **outside** `POST_StockMovement` silently clobbers / desyncs
the materialized balance from `SUM(ledger)` — a lost-update with no ledger row. The parts-stock-master
analysis argues the risk is "effectively closed by `ReadOnly`" on the form's qty box
(`parts-stock-master.md:226-234`) — but that is a *legacy-UI* mitigation, not a *cutover* guarantee:
(1) the rebuilt Perspective form, not the Delphi `.dfm`, is what ships, and the spike-log says the Save
button calls `UPDATE_PartsStockInfo` with the loaded value — fine *only* as long as the rebuild also keeps
qty read-only and never passes a hand-keyed `@QTY`; (2) `UPDATE_PartsStockInfoCount` is a genuine
qty-mover (line-pull/count flow) that has **no ledger post at all** and is not in the 4-producer set.

Both of these also fire **`UPDATE_PartNumber`** (BLOCKER 1), so even the "harmless" same-value rewrite
emits a `d.IN_QTY-i.IN_QTY = 0` row that is filtered (`i.IN_QTY <> d.IN_QTY`), but a real change writes an
`INV_PART_QTY_INF` row with no matching `INV_STOCK_LEDGER` row → the two ledgers diverge.

**Direction of fix (bounce to ignition-architect):** at cutover, either neuter the `IN_QTY` assignment in
`UPDATE_PartsStockInfo` (drop the `IN_QTY=@QTY` clause — it's a master-data edit, not a stock move) and
re-home `UPDATE_PartsStockInfoCount` onto `stockLedger.post()`, or add a DB guard that forbids non-ledger
writes to `IN_QTY`. Confirm the rebuilt PartsStock form never sends a mutated `@QTY`.

**Status:** NEW carry — the "5th writer" the brief asked us to hunt for. `UPDATE_PartsStockInfoCount` is a
genuine uncovered qty-mover; `UPDATE_PartsStockInfo` is a latent clobber.

---

## SHOULD-FIX 4 — Sequence hazard: backfill genesis-guard vs read-flip vs trigger-drop has an unguarded double-count / gap window, and the runbook gives no ordering

**Claim under test (the brief's headline):** is there a window where `IN_QTY` double-counts (triggers +
seams both live) or a post lands before/after the backfill genesis guard?

**The genesis guard is correct but narrow.** `SEED_OpeningBalance` (`spike-seed-opening-balance.sql:36-38`)
THROWs if any non-`OPENING_BALANCE` ledger row exists for the part — good, it stops seeding *after* a
forward post. But it does **nothing** about the inverse danger: the materialized `IN_QTY` it reads
(:40 `SELECT IN_QTY ... INV_PARTS_STOCK_MST`) must be the *quiesced* cutover value. There is no fencing
that the legacy triggers are *already stopped* when seeding runs. The only safe order is:

1. quiesce legacy writers (stop the Delphi app) → 2. DROP/disable the 12 qty-triggers → 3. run
`SEED_AllOpeningBalances` → 4. flip reads to the ledger → 5. enable the new seam write paths.

The runbook never states this order. §4 says "before any forward post"; §5 says "Sequence the drop with
the read-flip" — vaguely. The actual hazards if mis-ordered:

- **Triggers live + a seam write also live (overlap of steps 2 and 5):** `IN_QTY` double-counts — the
  trigger does `IN_QTY ± x` AND `POST_StockMovement` does `IN_QTY += delta`. Every producer docstring warns
  "do NOT run against a trigger-live DB or IN_QTY double-counts" — but nothing *enforces* the gap; it's
  prose discipline on a 5am cutover.
- **Seed runs while triggers still firing (step 3 before step 2):** a legacy trigger bumps `IN_QTY`
  between the seed `SELECT IN_QTY` of two parts → the opening balances are a torn read of a moving target.
  `SEED_AllOpeningBalances` is one statement/one transaction (good) but at READ COMMITTED a concurrent
  trigger UPDATE to a not-yet-seeded part is still visible inconsistently relative to the app's reads.
- **In-flight edit during the flip:** an order/shipment written by the legacy app at T-0 (trigger moved
  `IN_QTY`) but whose ledger replay hasn't happened → after the flip the ledger is missing that movement;
  conversely a seam post for a row the trigger already moved → double. The runbook has no "drain in-flight
  / freeze window" step.

**Direction of fix (bounce to ignition-architect):** the runbook needs an explicit, ordered,
single-writer cutover sequence with an enforced quiesce window (app down), not per-file prose warnings.
Make the trigger-drop and the seam-enable atomic w.r.t. app downtime.

**Status:** strengthens existing carry §5 (trigger-retire) — it currently lacks the ordering that makes it
safe.

---

## SHOULD-FIX 5 — `IN_BALANCE_AFTER` is computed from the pre-post `IN_QTY` and is not race-safe; it can record a wrong running balance under concurrency

**Claim under test:** `POST_StockMovement` "additive `IN_QTY += @delta` is commutative … needs NO
SERIALIZABLE" (`spike-post-stockmovement-proc.sql:8-11`). True for the `IN_QTY` *column*. But the proc
also stores a snapshot:

**Evidence:** `:59-60` `DECLARE @balanceAfter int = (SELECT IN_QTY FROM ... WHERE IN_PART_ID=@partId) + @delta;`
then inserts that as `IN_BALANCE_AFTER` (:62-67), and only afterward does the additive UPDATE (:71). Two
concurrent posts on the same part at READ COMMITTED both read the same pre-value `B`, and write
`IN_BALANCE_AFTER = B+d1` and `B+d2` — but the true serialized balances are `B+d1` and `B+d1+d2`. The
final `IN_QTY` is correct (additive UPDATEs serialize), but the **`IN_BALANCE_AFTER` audit column is
wrong** for at least one row.

Blast radius is limited (it's a snapshot/diagnostic; design:91 calls it "Optional but cheap; lets the
harness verify monotonic replay"). But a reviewer "verifying monotonic replay" off `IN_BALANCE_AFTER`
would see non-monotonic / duplicated balances under concurrency and chase a phantom. Either drop the
column, or compute it inside the same UPDATE via `OUTPUT inserted.IN_QTY`.

**Status:** NEW (minor) — a flaw in the §1 atomicity area.

---

## NIT 6 — Order commit posts nothing to the ledger; correct today, a latent gap at cutover

`order.commitOrders` (`order/code.py:126-138`) inserts open-order rows via `INSERT_OpenOrder` and updates
the renban counter — but never calls `stockLedger.post()`. Today that's fine: a freshly created order is
un-shipped/un-arrived (effect 0), and the legacy `INSERT_RecConfStatPartsStockMstQTY` trigger would also
move 0. **But** if an order is ever created already-stamped as shipped/arrived (a back-dated entry), the
legacy trigger *would* move stock and the rebuilt order path would not (triggers dropped at cutover, no
ledger post in this path). The receiving seam (`receiving.insertOpenOrder`) DOES post; the Order
worksheet commit path does not. Confirm the worksheet can never emit a counted order, or route its inserts
through the receiving seam. Low likelihood, flagging for completeness.

**Status:** NEW (minor); relates to §5 (what replaces the dropped INSERT trigger on the Order path).

---

## Things I tried to break and could NOT (SOUND)

- **`:to=<effect>:v=<stamp>` amend idempotency.** The brief's worry (two same-centisecond programmatic
  amends collide, 2nd delta dropped by the UNIQUE backstop) is real *only* for two amends to the SAME
  target value in the same centisecond — and that case has `delta = effect(new)-effect(old) = 0` and posts
  nothing anyway (stocktaking `code.py:91-99` reasoning verified across all four producers). Two amends to
  *different* targets differ in `:to=`. SOUND. The one residual: a sub-centisecond A→B then B→A round-trip
  *within the same stamp* would share `(stamp, to=A)`... but the intermediate B→A net is non-zero and the
  final key `:to=A:v=<stamp>` collides with the first `:to=A` — the 2nd is swallowed. This needs the
  stamp to advance; it's the documented "16-char stamp resolution" limit (§6) and is already a named carry.
  Acceptable for a solo-dev hand-edit workflow, not a high-frequency programmatic one.
- **`VC_SOURCE_EVENT varchar(100)` width.** `test_eventkey_integrity.py` exercises the REAL longest keys
  (RECEIVING_REVERSAL F5, ~64 chars) and round-trips them intact. The earlier varchar(40) truncation
  blocker is genuinely closed. SOUND.
- **Reject/shipping purge-inflation (old D11#9).** Verified live: `DELETE_AutoPurge` touches only
  open-order + HIST tables (design:153, Q2 RESOLVED), never reject/shipping — so reject/shipping deletes
  are always genuine restores. SOUND.
- **D6 manifest-cost CROSS APPLY.** `spike-report-procs-d6.sql` correctly replaces window-blind JOINs;
  the per-row 8-char date is passed (not the 6-char month) in `REPORT_MonthlyINVOICESSummary:51`. The
  only carry is *applying* it at cutover (§3) + the DB-level non-overlap constraint — already named. SOUND
  as written.
- **No dead-code-as-live or live-as-missing found** in the procs cited by the cutover artifacts. The
  triggers, `fn_ManifestCostAt`, and the four producer source tables all exist in the live dump.

---

## VERDICT

**NOT SOUND to flip yet — two BLOCKERs must be resolved first.**

1. **BLOCKER 1 (`UPDATE_PartNumber`)** — the ledger's own `IN_QTY` UPDATE fires a live 13th trigger that
   double-writes `INV_PART_QTY_INF` + `INV_PARTS_STOCK_MST_HIST` on every post. The runbook does not
   mention it; the harness only passes because it disables it. This is a real post-cutover data path with
   no cutover step. **NEW carry.**
2. **BLOCKER 2 (backfill contradiction)** — design §9 (derive-per-source-row, corrected on-hand) and the
   shipped `SEED_AllOpeningBalances` + runbook §4 (copy legacy `IN_QTY` forward) are different plans with
   different results for exactly the parts the whole §6 divergence-class effort was built to correct. The
   all-parts "0 drift" parity check is vacuous by construction. Pick one; if corrections are in scope, add
   the delta-reconciliation step §9 itself demands.

Plus three SHOULD-FIX (a 5th `IN_QTY` writer `UPDATE_PartsStockInfoCount` + the latent
`UPDATE_PartsStockInfo` clobber; the unstated cutover ordering/quiesce window; the racy `IN_BALANCE_AFTER`)
and two NITs (Order-commit ledger gap; minor).

The build itself is solid — idempotency keys, event-key width, the additive-post atomicity argument, the
D6 migration, and the divergence taxonomy all survive scrutiny. The danger is entirely at the **flip**:
the runbook treats cutover as "drop 12 triggers + seed + flip reads," but the real DB has a 13th trigger,
extra `IN_QTY` writers, and two competing backfill definitions. Resolve BLOCKER 1 and BLOCKER 2, add the
ordered quiesce sequence (SHOULD-FIX 4), and re-run the parity harness with a *non-vacuous* backfill before
David flips.
