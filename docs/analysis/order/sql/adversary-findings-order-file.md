# Adversarial Findings — Order-FILE Generator rebuild vs legacy `.ord` + feed/stamp/ship

Adversary review of the rebuilt order-file generator
(`docs/analysis/order/project-library/order_file/code.py`) against the legacy
`OrderFormCreateF.pas` (the live `TOrderFormCreate_Form.FormActivate`, .ord built
at `:556-585`). Goal: find the input where the rebuild's bytes / rows / files
differ from the legacy SQL + Pascal on the SAME inputs.

Method: legacy proc bodies read from the LIVE DB (`Inventory_Live` via
`sp_helptext`), not the spec; Pascal `Format`/`DayOfTheWeek`/`AsBoolean` semantics
proved with a real Free Pascal 3.2.2 compile (not from docs); live data probed
bounded + READ-ONLY on `Inventory_Live`/`VehicleOrder`; the e2e ran against the
`Inventory` spike DB and was restored as-found.

Verdict up front: the `.ord` LINE BODY is byte-faithful for the live value ranges,
and the feed/qty-alias + emit-then-stamp + logistics + ship-date semantics are
faithful. BUT there is **one live-reachable defect (the timestamp-bit filename)**
and **one latent defect (negative qty `%05d` vs `%.5d`)**, plus parity-method
blind spots that let both pass. Equivalence is therefore NOT fully proven.

---

## BLOCKER 1 — Non-timestamped suppliers get a TIMESTAMPED `.ord` filename (bit-coercion bug). LIVE-REACHABLE.

**Claim under test:** "the rebuild reproduces the legacy filename: `<name>-<sup>.ord`
when `BIT_ORDER_FILE_TIMESTAMP=0`, `<name>-<sup><fts>.ord` when =1."

**Counterexample (supplier 11111 PACIFIC_MFG, `BIT_ORDER_FILE_TIMESTAMP=0`):**

| | filename produced |
|---|---|
| **legacy** (`AsBoolean` of bit 0 = False → `:321`) | `PACIFIC_MFG-11111.ord` |
| **rebuild** (this run, observed in the e2e) | `PACIFIC_MFG-111112026062200000000.ord` |

Root cause — `code.py:444`:
```python
timestamp = bool(info["Order File Timestamp"])
```
`SELECT_SupplierInfo`'s `Order File Timestamp` is the `bit` `BIT_ORDER_FILE_TIMESTAMP`.
Through the spike feed it arrives as the **string `'0'`**, and `bool('0') is True`
in Python (any non-empty string is truthy). So a supplier flagged "no timestamp"
is treated as timestamped, and `_ord_filename(... timestamp=True ...)` appends
`yyyymmddhhmmss00`.

Proof:
- Live config: `SELECT VC_SUPPLIER_CODE, BIT_ORDER_FILE_TIMESTAMP FROM INV_SUPPLIER_MST WHERE VC_SUPPLIER_CODE='11111'` → `11111 | 0` (and `Inventory_Live` agrees). 15 of 16 live suppliers are `ts=0`; only `38844 SUPERIOR` is `ts=1`.
- Reproduced through the real shim: `_supplier_info('11111')['Order File Timestamp']` = `'0'` (str), `bool('0')` = `True`, `_ord_filename(...)` = `PACIFIC_MFG-111112026062200000000.ord`.
- Legacy filename forms proved by FPC compile: `ts=0` → `PACIFIC_MFG-11111.ord`; `ts=1` → `PACIFIC_MFG-111112026062200000000.ord`.

**Why it matters:** the legacy intent for `ts=0` is a STABLE filename that
overwrites each run (the dispatch/FTP pickup keys on `<name>-<sup>.ord`). The
rebuild emits a fresh, ever-changing timestamped name on every run → the pickup
process never finds the expected stable name, and the supplier dir accumulates one
new file per run. This is wrong for 15/16 live suppliers.

**Classification:** code defect (`code.py:444`), masked by a parity-method flaw.
On the real Ignition gateway, JDBC returns a `bit` as a Java Boolean/Integer, so
`bool(False)`/`bool(0)` would be correct THERE — i.e. the bug fully manifests under
the spike shim and *may* not manifest on the gateway. That ambiguity is itself the
problem: the code relies on driver type fidelity it never asserts, the spike test
exercises the wrong-typed path, and nothing catches the divergence. Fix:
`str(info["Order File Timestamp"]).strip() in ("1","True","true")` or coerce the
bit explicitly. Same fragile pattern sits on `Site Number in Order`
(`bool('1')`=True too) — harmless only because the rebuild never reads it (sendsite
deliberately unbuilt).

---

## BLOCKER 2 — Parity method is blind to the filename + cannot prove the ship-date field. METHOD FLAW.

**Claim under test:** "the e2e proves the published `.ord` equals the legacy."

Two gaps make a green e2e NOT a proof of equivalence:

1. **The filename is never compared to the legacy.** `test_order_file_e2e.py:170`
   only checks `published.endswith(".ord")` and existence. The wrong timestamped
   name (BLOCKER 1) sailed through with all 25 assertions PASS. A green run here
   does not prove the published artifact's NAME matches legacy — and it doesn't.

2. **The ship-date field is verified circularly (rebuild vs rebuild).**
   `test_order_file_e2e.py:152` computes `shipOffset = of.get_ship_offset(...)`
   (the rebuild's OWN function) and then builds the expected bytes with
   `of.format_ord_line(... shipDate ...)` (the rebuild's OWN formatter, line 176).
   So the ship-date bytes are checked against the code that produced them — a
   self-consistency check, not equivalence to legacy `GetShip`. The pure test
   (`test_order_file_build.py` §6) checks `compute_ship_offset` against
   hand-rolled expectations, but those expectations were written by the rebuild
   author, not derived from the legacy.

   I broke the circularity independently (see SHOULD-FIX/PROVEN below): a faithful
   FPC re-compile of `GetShip` on the REAL `VehicleOrder` holiday window gives the
   same offset 18 → so the ship logic IS faithful — but the TEST as written does
   not establish that.

**Classification:** test/parity-method flaw. The .ord LINE BODY assertion (bytes
per row) IS non-vacuous and correct; the FILENAME and SHIP-DATE assertions are not
load-bearing against legacy.

---

## SHOULD-FIX 3 — Negative qty: `%05d` (Python) ≠ `%.5d` (Pascal). LATENT (not in current data).

**Claim under test:** "`%05d` reproduces the Pascal `format('%.5d', [IN_QTY])`."

**Counterexample (qty = −5), proved by FPC compile vs Python:**

| qty | legacy `format('%.5d')` | rebuild `"%05d"` | match |
|---|---|---|---|
| 240 | `00240` | `00240` | yes |
| 1200 | `01200` | `01200` | yes |
| 0 | `00000` | `00000` | yes |
| 123456 | `123456` | `123456` | yes |
| **−5** | **`-00005`** (6 chars) | **`-0005`** (5 chars) | **NO** |
| −50 | `-00050` | `-0050` | NO |
| −12345 | `-12345` | `-12345` | yes |

Pascal `.5` is a PRECISION (minimum digit count, sign prepended OUTSIDE the
5 zero-padded digits); Python `05` is a FIELD WIDTH that the sign CONSUMES. They
diverge for any negative with magnitude < 10000.

Live reachability: `INV_OPEN_ORDER_INF.IN_QTY` is `int NOT NULL` with **no check
constraint** (proved: 0 check constraints on the table), so a negative is
*possible*; on the current snapshot `MIN(IN_QTY)=4, MAX=21600, neg_cnt=0,
over5_cnt=0` — so it is NOT exercised today. The rebuild's own test
(`test_order_file_build.py:71-73`) asserts `'-0005'` — i.e. it ENSHRINES the
divergent Python behavior as "expected." If a negative order qty ever occurs
(adjustment/return), the rebuild's fixed-width line shifts one byte vs legacy and a
position-keyed sub-supplier parser misaligns.

**Classification:** code defect (`code.py:92`), latent on current data; the pure
test bakes in the wrong expectation. Fix: emit the Delphi shape
(`('-' if q<0 else '') + '%05d' % abs(q)`).

---

## NIT 4 — `fts` is a fixed midnight placeholder (the one timestamped supplier).

`code.py:414`: `fts = runDate.strftime("%Y%m%d") + "000000" + "00"`. The legacy
`fts = formatdatetime('yyyymmddhhmmss00', now)` uses wall-clock time. For the one
`ts=1` supplier (38844 SUPERIOR) the rebuild writes `...20260622 00000000.ord`
instead of `...20260622HHMMSS00.ord`. Self-flagged in the code as a gateway-passes-
real-now placeholder; not wrong logic, but the deterministic value collides on
repeat same-day runs (the legacy's 1-second resolution is also collision-prone —
legacy H7 — so this is no worse, just different). Classification: noted gap, not a
defect.

## NIT 5 — Archive path separator mixing on a non-Windows gateway.

`code.py:287` uses `os.path.join(supplierDir, "Archive")`; the legacy uses
`lastdirectory + '\Archive\'`. On a Linux gateway `os.path.join('S:\\PACIFICMFG',
'Archive')` yields `S:\PACIFICMFG/Archive` (mixed `\`/`/`). Cosmetic only — the
spec (§12) already says these become configured per-site shares, not literal
Windows paths. Flag, do not block.

---

## What IS proven faithful (the attacks that FAILED to break it)

These were attacked with live counterexamples and held:

- **`.ord` line body, byte-exact (live ranges).** Independent FPC re-compile of
  `:568-574` for the e2e fixture row (supplier 11111, FRS 9800001, renban 98R001,
  part 426070205000, qty 1200, ship 20260710) →
  `111119800001  98R0014260702050000120020260710` — byte-identical to the rebuild's
  published file. `%8s` right-justify (`H006`→`    H006`), `>8` renban not truncated
  (`16GY12345` stays 9 wide), `%.5d` zero-pad all match for non-negative qty. Live
  renban width is 3..8 (none >8); qty 4..21600 (none >5 digits) — so the unpadded
  supplier/FRS/part raw-concat (legacy H3) and the >8 / >99999 hazards are NOT
  exercised on this vintage (faithful carries, golden-`.ord` still pending to
  confirm the receiving parser expects full-width columns).

- **Feed = aliased `SELECT_OrderNotOrdered`; qty is the OPEN-ORDER `IN_QTY`.**
  `_FEED_SQL` matches `spike-order-file-feed.sql` (test enforces no-drift) and the
  proc body verbatim (live `sp_helptext`): same `WHERE ((VC_ORDER_DATE IS NULL) OR
  ='') AND VC_RENBAN_NUMBER <> ''`, same `ORDER BY s.VC_SUPPLIER_CODE,
  i.VC_RENBAN_NUMBER`. The aliasing fix is real and observable: e2e emitted qty
  `01200` (the open-order `i.IN_QTY`), NOT the parts-stock on-hand `p.IN_QTY` (the
  fixture part's on-hand is 0; the legacy `SELECT *` `fieldbyname('IN_QTY')` first-
  match is also `i.`, so this MATCHES legacy and FIXES the H10 latent ambiguity).
  No fan-out: `i→p→s` 1:1, `r` LEFT JOIN cannot fan (renban-group PK unique). The
  feed-supplier invariant `i.VC_SUPPLIER_CODE == s.VC_SUPPLIER_CODE` holds 4284/4284
  on live, so the supplier-break grouping (rebuild breaks on `i.`, sorts on `s.`,
  same as legacy) aligns; CI_AS collation on both = same group collapse.

- **Emit-from-snapshot-then-stamp + re-emit guard (H11).** `UPDATE_ORDEROrderDate`
  body verbatim from live: keys on part+FRS only (NO renban filter), guard
  `AND VC_ORDER_DATE <> @OrderDate`. e2e fixture = 2 renbans / 1 part+FRS: one stamp
  marked BOTH rows ordered (matches legacy), emitted from the pre-read snapshot so
  no sibling is dropped, and a 2nd run emitted NOTHING (idempotent). De-duping the
  stamp on distinct (part,FRS) is a net-identical optimization of the legacy per-row
  call. ATOMICITY (H1 fix) verified: a mid-stamp RAISERROR rolled back, deleted the
  staged `.tmp`, published NO final `.ord`, and left the rows un-stamped.

- **Ship date (`SELECT_PartShipDays` + `GetShip`), proved INDEPENDENTLY of the
  rebuild.** Proc body verbatim (renban-group override: `@gc=IN_RENBAN_ID`; group's
  ship-days if non-null else part's). Fixture part is in renban group 7
  (base Ship=0, weekday overrides 13). A faithful FPC re-compile of `GetShip`
  (lead=13, Monday 2026-06-22) against the REAL `VehicleOrder` 'H' holidays
  (2026-07-03 + the 07-13..17 summer shutdown) gives offset **18 → ship 20260710**,
  identical to the rebuild's `compute_ship_offset` on the same holiday set and to
  the e2e's published bytes. `DateUtils.DayOfTheWeek` is ISO 1=Mon..7=Sun (FPC-
  confirmed), so the weekday-pick map {1:M..6:S} and the weekend test
  (`<6` legacy = `>=6` weekend = rebuild `isoweekday()>=6`) match. The GetShip
  calendar-inconsistency (skips weekends + 'H' only; ignores 'O'/'X'/'W', unlike the
  forecast day-spread) is reproduced FAITHFULLY (a carry, flagged in code for
  cutover adjudication — NOT a new divergence). Ship-day columns are nullable but
  `pick_ship_lead`'s `int(v or 0)` matches Delphi `AsInteger`=0 on NULL/0.

- **Logistics `'NONE'`/`''`/NULL ladder.** Proc body verbatim (INNER join → 0 rows
  for an unlinked part). On live: all 47 parts have `IN_LOGISTICS_ID=NULL` (part
  level dormant), 3 suppliers link `TLDL S:\TLDL`, 13 → `'NONE'`. The rebuild's
  `resolve_logistics_dir` reproduces part→supplier→`'NONE'`, keeps `''`≠`'NONE'`≠NULL
  distinct, and `_supplier_destinations` skips the logistics copy only on `'NONE'`.
  e2e proved LocalFTP=False → 1 copy; True → supplier+logistics+archive (3,
  byte-identical); True+`'NONE'` → supplier+archive only. (The legacy H8 empty-string
  root-relative write is deliberately SKIPPED in the rebuild with a warning — a
  documented, safer divergence, not a faithfulness break.)

- **Forecast→order supplier consistency.** Re-confirmed on live: breakdown vs
  parts-stock supplier 959/959 match; open-order vs parts-stock supplier 4284/4284
  match. No supplier mismatch across the chain.

---

## Verdict

The rebuild's **`.ord` line body, the aliased feed (open-order `IN_QTY`, no
fan-out, same WHERE/ORDER BY), the emit-from-snapshot-then-stamp + re-emit guard,
the atomicity fix, the part→supplier→`'NONE'` logistics ladder + LocalFTP fan-out,
and the `SELECT_PartShipDays`/`GetShip` ship-date computation are PROVEN equivalent
to the legacy** on the SAME inputs (line body and ship date proved independently of
the rebuild via FPC compiles, not just self-consistency).

It is **NOT fully proven equivalent** because of two real divergences and a
parity-method blind spot:

- **BLOCKER 1 (live):** `bool(info["Order File Timestamp"])` mis-coerces the
  no-timestamp bit, so 15/16 live suppliers get a wrong timestamped filename under
  the spike; the gateway JDBC type *may* save it, but the code never asserts that
  and the test never compares the filename to legacy.
- **BLOCKER 2 (method):** the e2e never compares the filename to legacy and verifies
  the ship-date field circularly (rebuild vs rebuild) — a green run does not prove
  filename or ship-date equivalence (I had to prove ship-date out-of-band).
- **SHOULD-FIX 3 (latent):** `%05d` ≠ Pascal `%.5d` for negative qty (`-0005` vs
  `-00005`); not in current data but the column allows it and the test bakes in the
  wrong expectation.

Two pieces remain UNPROVABLE from available data (call it out, don't paper over):
the receiving sub-supplier parser's full-width expectation for the raw
supplier/FRS/part columns, and the timestamped supplier's wall-clock `fts`, both
pending a GOLDEN production `.ord`. Net: the `.ord` payload semantics are sound;
fix the filename bit-coercion + the negative-qty format and add a legacy-anchored
filename/ship-date assertion before this can be called PROVEN equivalent.

---

# RE-VERIFY (round 2) — 2026-06-21

Re-attack of the developer's fixes on branch `m2-order-file-gen` (claims: BLOCKER 1
filename bit, SHOULD-FIX 3 negative-qty `%.5d`, the feed/stamp NULL drop, and the
circular ship-date test all fixed). Method unchanged: legacy proc bodies read from
the LIVE/spike DB (`sp_helptext`), Pascal semantics proved by a real FPC 3.2.2
compile, live data probed bounded + READ-ONLY on `Inventory_Live`/`VehicleOrder`,
the e2e run against the `Inventory` spike and restored as-found (0 synthetic rows
remaining, verified). Each prior finding is re-attacked → RESOLVED-with-proof or
STILL-OPEN-with-counterexample.

## BLOCKER 1 (filename bit) — RESOLVED.

The fix is a dedicated `_coerce_bit` (`code.py:368-395`) called at `code.py:503`
in place of the old `bool(info["Order File Timestamp"])`. It reproduces Delphi
`TField.AsBoolean = (value <> 0)` across BOTH transports + NULL/empty:

- FPC ground truth (compiled): `AsBoolean` of bit `0` = FALSE (stable name), bit `1`
  = TRUE (timestamped).
- `_coerce_bit` over every transport (run directly): shim-string `'0'`→False,
  `'1'`→True; JDBC int `0`→False, `1`→True; JDBC bool `False/True`→False/True;
  `None`(NULL)→False; `''`→False; padded `' 0 '`→False. 0 mismatches.
- **Genuine bug-path proof (not vacuous):** the spike shim transports
  `SELECT_SupplierInfo['Order File Timestamp']` for supplier 11111 as the **string
  `'0'`** (`type=str`) — the exact original bug input. `bool('0')` is still `True`
  (the old defect); `_coerce_bit('0')` is `False`. End-to-end filenames on that real
  row: OLD-bug → `PACIFIC_MFG-111112026062200000000.ord`; FIXED →
  `PACIFIC_MFG-11111.ord`.
- **Live re-prove on a real supplier:** `INV_SUPPLIER_MST` (spike + `Inventory_Live`)
  has `BIT_ORDER_FILE_TIMESTAMP` = `bit`, nullable; per-supplier 15× `0`, 1× `1`
  (only 38844 SUPERIOR = 1). Supplier 11111 = `PACIFIC_MFG`, bit `0`. The published
  e2e file for 11111 is `PACIFIC_MFG-11111.ord` — flag 0 → stable, no timestamp.
  Flag-1 form (38844) → `SUPERIOR-38844<fts>.ord`. Both match legacy `:321`/`:290`.
- **Filename test is non-circular:** `test_order_file_e2e.py:194` pins
  `legacy_name = "PACIFIC_MFG-11111.ord"` (the legacy `:321` form, hand-derived, NOT
  the rebuild's `_ord_filename` output) and asserts `os.path.basename(published) ==
  legacy_name` (`:195-197`) plus "no yyyymmdd in the tail" (`:198-200`). A regression
  to `bool('0')` (timestamped) FAILS this assertion (proven: the OLD-bug filename
  differs).

No bit value/transport found where it still diverges from `AsBoolean`. The only
`_coerce_bit` tokens that differ from naive truthiness (`'1.0'`/`'0.0'`/`'abc'`→False)
are non-bit garbage a `bit` column can never emit — not reachable, not a parity gap.
Classification: **code defect — FIXED**, and the test that masked it is now a real
guard. Carry: the archive copy is still always timestamped even for a bit-0 supplier
(`code.py:540` `timestamp or kind=="archive"`), matching legacy `:341` (e2e emitted
`PACIFIC_MFG-111112026062200000000.ord` in `Archive/`) — faithful.

## SHOULD-FIX 3 (negative qty `%.5d`) — RESOLVED.

`_format_qty` (`code.py:58-69`) now emits the Pascal shape: `("-" + "%05d" % -n)`
for negatives (sign OUTSIDE the 5-digit zero-pad), `"%05d" % n` otherwise. Verified
field-by-field against an FPC `format('%.5d',[q])` compile:

| qty | FPC `%.5d` | rebuild `_format_qty` | match |
|---|---|---|---|
| 240 | `00240` | `00240` | yes |
| 0 | `00000` | `00000` | yes |
| 99999 | `99999` | `99999` | yes |
| 100000 | `100000` | `100000` | yes |
| 123456 | `123456` | `123456` | yes |
| **−5** | **`-00005`** | **`-00005`** | **yes (was `-0005`)** |
| −50 | `-00050` | `-00050` | yes |
| −12345 | `-12345` | `-12345` | yes |
| −99999 | `-99999` | `-99999` | yes |
| −100000 | `-100000` | `-100000` | yes |

0 mismatches across the whole set. The pure unit test now asserts the CORRECT
`-00005` (`test_order_file_build.py:73-81`) and would FAIL on the old `%05d`
behavior (proven: `"%05d" % -5 == "-0005" != "-00005"`). Classification:
**code defect — FIXED**; the test no longer enshrines the wrong expectation. (Still
latent on data — live `IN_QTY` min is positive — but now correct if a negative/return
qty ever occurs.)

## Feed/stamp NULL — RESOLVED (faithful + safe).

`_FEED_SQL` (`code.py:317-338`) and the canonical `spike-order-file-feed.sql` both
carry `WHERE i.VC_ORDER_DATE = '' AND i.VC_RENBAN_NUMBER <> ''` — the legacy
`SELECT_OrderNotOrdered`'s `IS NULL` arm is DROPPED. Verified against live bodies:

- Legacy feed `SELECT_OrderNotOrdered` (spike `sp_helptext`): `WHERE
  ((i.VC_ORDER_DATE is null) or (i.VC_ORDER_DATE = '')) AND i.VC_RENBAN_NUMBER <>
  '' ORDER BY s.VC_SUPPLIER_CODE, i.VC_RENBAN_NUMBER`.
- Legacy stamp `UPDATE_ORDEROrderDate` (spike `sp_helptext`): `WHERE VC_PART_NUMBER
  = @PartNumber AND VC_FRS_NUMBER = @FRSNumber AND VC_ORDER_DATE <> @OrderDate` — the
  rebuild calls this exact proc unchanged (`code.py:586-588`); same 4 params.
- **Schema proves the `IS NULL` arm unreachable:** `INV_OPEN_ORDER_INF.VC_ORDER_DATE`
  is `is_nullable = 0` with `DEFAULT ('')` on the spike, and `is_nullable = 0` on
  `Inventory_Live`. Live data: `null_cnt = 0`, total 4238 rows — no NULL exists or
  can be inserted.
- **The drop is the SAFE direction (matches the stamp, not just the schema):** the
  stamp clears via `VC_ORDER_DATE <> @OrderDate`; under ANSI_NULLS ON a NULL row's
  `NULL <> '...'` is UNKNOWN → never stamped → would re-emit forever. The legacy
  feed's `IS NULL` arm could surface a row the stamp can never clear; the rebuild
  selects ONLY what the stamp can clear (`= ''`). On the NOT-NULL schema the two
  predicates are row-for-row identical; the rebuild additionally cannot trip the
  legacy's latent re-emit-forever hazard.
- **Drift-guard in sync:** the e2e normalizes (strips comments/whitespace) and
  asserts `_norm(of._FEED_SQL) == _norm(canonical .sql)` — re-ran: **True**
  (`test_order_file_e2e.py:135-137`). The `ORDER BY s.VC_SUPPLIER_CODE,
  i.VC_RENBAN_NUMBER` and the aliasing-off-`i.` (H10 fix) are byte-identical to the
  proc.

Classification: **faithful on the live schema + safer w.r.t. the stamp** — RESOLVED.

## Ship-date non-circular — RESOLVED.

The e2e anchors the ship date to an OUT-OF-BAND legacy constant, not rebuild-vs-
rebuild. `test_order_file_e2e.py:164-171` pins `LEGACY_SHIP_OFFSET = 18` /
`LEGACY_SHIP_DATE = "20260710"` (GetShip, lead 13, Mon 2026-06-22, real `'H'`
window) and asserts the rebuild's `get_ship_offset` AND the published .ord's
trailing-8-bytes equal it (`:166-171`, `:217-218`). I re-derived that constant
independently (NOT via the rebuild's `compute_ship_offset`):

- Part 426070205000 `SELECT_PartShipDays` (spike): `Ship=0, ShipM=13, ShipT=13,
  ShipW=13, ShipTh=13, ShipF=13, ShipS=0` → Monday lead 13 (renban-group override).
- `VehicleOrder.AD_GetSpecialDate` `'H'` rows in the window (READ-ONLY): 2026-07-03
  (Fri) + the 2026-07-13..17 (Mon–Fri) summer shutdown.
- A hand calendar walk of GetShip's x/y bookkeeping (skip weekend + `'H'`, from
  tomorrow) lands the 13th working day on **2026-07-10 (Fri), offset 18** — matching
  the pinned constant exactly. The rebuild's `get_ship_offset` returned 18 / 20260710
  on the same input (e2e PASS). If the spike `VehicleOrder` `'H'` window changes, the
  constant must be re-derived (the test comment says so); it is a legacy anchor, not
  a self-consistency loop.

Classification: **parity-method flaw — FIXED** (filename + ship-date both now
legacy-anchored; the e2e fails on a regression of either).

## No regression — CONFIRMED.

- **.ord line body still byte-exact.** e2e published bytes for the fixture row:
  `111119800001  98R0014260702050000120020260710` = supplier `11111` raw + FRS
  `9800001` raw + `  98R001` (`%8s` renban, left-pad to 8) + part `426070205000` raw
  + `01200` (`%05d` qty) + `20260710` (ship). Matches legacy `:568-574` field order;
  `format_ord_line` (`code.py:72-115`) unchanged in layout. Pure test: 53 PASS / 0
  FAIL (incl. `%8s`, >8-renban no-truncate, >5-digit qty widen, raw-field shift H3,
  CRLF-after-every-line).
- **Feed / qty-alias intact.** Emitted qty is the OPEN-ORDER `i.IN_QTY` (1200), NOT
  parts-stock on-hand (0) — aliasing H10 fix holds (e2e `:210-213`).
- **Emit-then-stamp intact.** Both same part+FRS rows stamped on one call (H11), 2nd
  run emits nothing (idempotent), atomicity (H1) + RISK-1 all-or-nothing publish all
  green. e2e: **34 PASS / 0 FAIL / 0 SKIP**; spike restored as found.

## RE-VERIFY verdict

All four prior findings are **RESOLVED with proof** (no counterexample found on
re-attack):

- **BLOCKER 1 (filename bit) — FIXED.** `_coerce_bit` reproduces `AsBoolean =
  (value <> 0)` on the shim-string, JDBC-int, JDBC-bool, and NULL/empty transports;
  the genuine `'0'`-string bug path is exercised and the published name is the stable
  `PACIFIC_MFG-11111.ord`; the test is non-circular and fails on a `bool('0')`
  regression.
- **SHOULD-FIX 3 (negative qty) — FIXED.** `_format_qty` matches FPC `%.5d` for
  every value including `-5 → -00005`; the test asserts the correct shape.
- **Feed/stamp NULL — RESOLVED.** The `IS NULL` arm is dropped from both `_FEED_SQL`
  and the canonical `.sql` (drift-guard green); justified by the NOT-NULL `DEFAULT
  ('')` schema and safer w.r.t. the stamp's `<>` clause; the stamp proc is the legacy
  body verbatim.
- **Ship-date — RESOLVED.** Anchored to an independently FPC/calendar-derived legacy
  GetShip constant (offset 18 → 20260710), not rebuild-vs-rebuild.

The rebuild now reproduces the legacy `.ord` line body + filename (both timestamp
branches) + feed/stamp semantics + ship-date on the SAME inputs, with the divergences
removed and the parity method made legacy-anchored (non-circular) for both the
filename and the ship-date.

**Equivalence verdict:** PROVEN equivalent to the legacy `OrderFormCreateF.pas` /
`SELECT_OrderNotOrdered` / `UPDATE_ORDEROrderDate` for all tested inputs and the live
value ranges, **modulo the still-pending GOLDEN production `.ord`** — the two
genuinely UNPROVABLE-from-available-data items remain and are unchanged by this round:
(1) whether the receiving sub-supplier parser expects always-full-width
supplier/FRS/part columns (the legacy + rebuild both write them RAW — H3), and (2)
the one timestamped supplier's wall-clock `fts` (the rebuild uses a deterministic
midnight placeholder; the gateway timer passes real `now()`). Neither is a new
divergence; both await a golden artifact to confirm, and should be called out, not
papered over.
