# Adversary Findings — Outbound EDI 856 builder vs legacy `EDI856Object.pas` / `REPORT_EDI856`

Goal: refute that the rebuilt 856 builder + SQL feed reproduce the legacy 856 byte-for-byte. Every finding
carries a file:line or a live counterexample (legacy vs rebuild) with its output. Spike left as-found (all
destructive probes were rolled-back transactions on `Inventory`; `Inventory_Live`/`VehicleOrder` read-only).

Artifacts under review:
- Feed: `docs/analysis/edi/856/spike-edi856-feed.sql`, inlined as `_FEED_SQL` in
  `docs/analysis/edi/856/project-library/edi856/code.py:286-304`.
- Builder: `docs/analysis/edi/856/project-library/edi856/code.py` (`build_856`, `send_856`).
- Spec: `docs/analysis/edi/856/edi856-wire-format.md`, `report-edi856-data-analysis.md`.
- Legacy: `EDI856Object.pas` (live per `InventorySystem.dpr:48`); `REPORT_EDI856` (live OBJECT_DEFINITION on
  `Inventory`).
- Tests: `scripts/e2e/test_edi856_build.py`, `scripts/e2e/test_edi856_e2e.py`.

---

## BLOCKER 1 — `TD1` segment byte-divergence: rebuild emits `TD1*`, legacy emits `TD1**`

Legacy `EDI856Object.pas:294-296`:
```
TD1Item := 'TD1' + fSepElement;   // 'TD1*'
TD1Item := TD1Item + fSepElement; // 'TD1**'   <- a SECOND empty element
HLList.Add(TD1Item);
```
Rebuild `code.py:203`: `hl.append(sep.join(["TD1", ""]))` -> **`TD1*`** (one separator, one empty element).

Counterexample (reconstructing legacy from the `.pas` vs the rebuild's `sep.join`, sep `*`):
```
TD1  legacy='TD1**'   rebuild='TD1*'   MATCH=False
```
The rebuild drops one empty trailing element. This is a hard byte difference on a segment present in EVERY
856. The locked decision is "byte-faithful to the TEMA-accepted legacy output"; this is not faithful.

Type: **code defect** (the port deviates from the legacy AND from its own spec — `edi856-wire-format.md:39`
correctly records `TD1**`). The unit test froze the WRONG value (`test_edi856_build.py:100` `"TD1*"`), so the
test is green against the bug. See PARITY-METHOD finding below.

---

## BLOCKER 2 — `LIN` segment byte-divergence: legacy emits a TRAILING element separator, rebuild does not

Legacy `EDI856Object.pas:347-353` — note line 352 ends with `+fSepElement` AFTER the kanban:
```
LINItem := 'LIN' + fSepElement;                         // 'LIN*'
LINItem := LINItem + fSepElement;                       // 'LIN**'
LINItem := LINItem + 'BP' + fSepElement;
LINItem := LINItem + ...PartNumber... + fSepElement;
LINItem := LINItem + 'RC' + fSepElement;
LINItem := LINItem + ...Kanban... + fSepElement;        // <- TRAILING separator after kanban
HLList.Add(LINItem);
```
So legacy LIN = `LIN**BP*<part>*RC*<kanban>*` (trailing `*`).
Rebuild `code.py:231`: `sep.join(["LIN","","BP",part,"RC",kanban])` -> `LIN**BP*<part>*RC*<kanban>` (NO
trailing separator).

Counterexample (part `42600F261100`, kanban `JZV9`, sep `*`):
```
LIN  legacy='LIN**BP*42600F261100*RC*JZV9*'   rebuild='LIN**BP*42600F261100*RC*JZV9'   MATCH=False
```
One byte different on EVERY item line of EVERY 856. (`SN1` is fine — legacy `:355-358` ends `+'PC'` with no
trailing sep; rebuild matches: `SN1**32*PC` == `SN1**32*PC`.)

Type: **code defect, propagated from a spec defect.** `edi856-wire-format.md:41` ALSO dropped the trailing
sep (`LIN**BP*<PartNumber>*RC*<Kanban>`), so the port faithfully reproduced the spec's error — but the spec is
wrong vs the `.pas`. Either way the emitted bytes differ from the legacy stream.

Note: BLOCKER 1+2 do NOT change SE01/CTT01 (same number of segments, just different intra-segment element
counts), so the counts stay self-consistent — which is exactly why a structure-only / self-consistency test
misses them. They WILL change the byte stream a strict X12 parser (TEMA) consumes.

---

## BLOCKER 3 (latent, data-gated) — feed `MO_PRICE` GROUP-BY collapse diverges from legacy on a real-world shape

Legacy `REPORT_EDI856` `@EIN=0` branch GROUP BY (live body) includes **`m.MO_PRICE`**,
`a.VC_START_SEQ_NUMBER`, `a.VC_LINE_NAME`, `a.IN_ASN_EIN`. The rebuild feed
(`spike-edi856-feed.sql:65-69`, `code.py:301-302`) GROUP BY = `Manifest, Part, IN_QTY, prodDate, Kanban`
ONLY — it **dropped `MO_PRICE`**. `START_SEQ`/`LINE_NAME`/`EIN` are header-constant per ASN (proven:
ASN 2298 has one distinct value of each) so they cannot add rows; `MO_PRICE` comes from the cost join and
CAN fan out.

Counterexample (rolled-back txn on `Inventory`, ASN 2298 / part `426200ZR2000` / date `20180925`; added a
SECOND cost window `20180101..20191231` also covering the date, price `99.99` alongside the existing
`81.25`):
```
LEGACY  | 78092551 | 426200ZR2000 | 81.2500 | 1 | JLN9
LEGACY  | 78092551 | 426200ZR2000 | 99.9900 | 1 | JLN9     <- legacy keeps BOTH (MO_PRICE in GROUP BY)
REBUILD | 78092551 | 426200ZR2000 |         | 1 | JLN9     <- rebuild collapses to ONE (no MO_PRICE)
```
Legacy emits **2** detail rows (= 2 Item HLs / 2 LIN / 2 SN1) for the same shipped line; the rebuild emits
**1**. Different CTT01, different SE01, different shipped-line breakdown. `MO_PRICE` is never emitted on the
wire, yet legacy uses it to SPLIT the rows — so the rebuild's "project only what the wire needs" reasoning
(`spike-edi856-feed.sql:33`) is the bug: dropping a GROUP-BY column the wire doesn't carry still changes the
ROW COUNT the wire does carry.

On the **current snapshot this does not fire**: every part has exactly one cost row (45 cost parts, 0 with
>1 row / >1 price). Full-snapshot diff legacy-9col vs rebuild-5col over all ASNs: **14,356 == 14,356 feed
rows, 0 ASNs differ.** So the rebuild is feed-equivalent on TODAY's data but diverges the moment a part gets
a second price-distinct overlapping cost window — which the schema permits (`INV_MANIFEST_COST_MST` is a HEAP
with no uniqueness on part+window).

Type: **code defect, latent / data-gated.** Currently masked by a sparse single-price snapshot.

---

## SHOULD-FIX 4 — forecast `CROSS APPLY (SELECT TOP 1 …)` has NO `ORDER BY` (nondeterministic) and DROPS rows legacy keeps (the #2 keystone risk)

Rebuild `code.py:295-297` / `spike-edi856-feed.sql:59-61`:
```
CROSS APPLY (SELECT TOP 1 f1.VC_ASSY_KANBAN_NUMBER
             FROM INV_FORECAST_DETAIL_INF f1
             WHERE f1.VC_ASSY_PART_NUMBER_CODE = d.VC_ASSY_PART_NUMBER) f
```
Legacy uses `JOIN INV_FORECAST_DETAIL_INF f ON d.VC_ASSY_PART_NUMBER = f.VC_ASSY_PART_NUMBER_CODE` (a real
INNER join) then 9-col GROUP BY. `INV_FORECAST_DETAIL_INF` is a **HEAP with no clustered index** on both
`Inventory` and `Inventory_Live` (verified via `sys.indexes` -> `NULL | HEAP`), and the TOP 1 has no
`ORDER BY`, so the surviving row is whatever the storage engine scans first — undefined and unstable across
inserts / stats / page reorg / parallel plans.

Counterexample A — when a part has >1 forecast row, legacy keeps ALL distinct kanbans, rebuild keeps ONE
(rolled-back txn on `Inventory`, part `426200ZR2000`, added a 2nd forecast row kanban `AAAA` alongside
`JLN9`):
```
LEGACY  | 78092551 | 426200ZR2000 | 1 | AAAA
LEGACY  | 78092551 | 426200ZR2000 | 1 | JLN9     <- legacy fans out then GROUP-BY keeps both kanbans
REBUILD | 78092551 | 426200ZR2000 | 1 | JLN9     <- rebuild TOP 1 keeps ONE
```
Counterexample B — which one TOP 1 keeps is nondeterministic (same 2-row data):
```
TOP1 plain                       -> JLN9
TOP1 ORDER BY kanban ASC         -> AAAA
TOP1 ORDER BY ID_FORECAST_DETAIL DESC -> AAAA   (engine can surface either as "first")
```
So even in the "matches legacy by collapsing to 1 row" case, the kanban on the surviving LIN is a coin flip.

On the **current snapshot this does not fire**: forecast is 1:1 per part (Inventory 50 parts, Inventory_Live
all parts: 0 parts with >1 row, 0 with >1 distinct kanban). So TOP 1 == the unique row == the legacy
GROUP-BY row TODAY. The rebuild's own comment (`spike-edi856-feed.sql:18-21`) claims TOP 1 "could only ever
have inflated the 856, never matched it" for the legacy fan-out — that is the wrong way round: when a part has
2 forecast rows with DISTINCT kanbans, the legacy correctly emits 2 lines and the rebuild drops one (loses a
shipped line), and even with identical kanbans the legacy GROUP BY would collapse to 1 too, so the legacy is
NOT "always inflated." The fan-out the rebuild "tightened" is a fan-out legacy resolves via GROUP BY, not a
pure inflation.

Type: **code defect, latent / data-gated + nondeterministic.** Add `ORDER BY` at minimum; the TOP-1 vs
INNER+GROUP-BY semantics differ whenever a part has >1 distinct forecast kanban.

---

## SHOULD-FIX 5 — the parity METHOD is self-referential on the two divergent segments (vacuous byte-parity)

`test_edi856_build.py:87` labels the expected segment list "hand-computed from EDI856Object.pas," but two
entries do NOT match the `.pas`:
- line 100: `"TD1*"` — legacy emits `TD1**` (`.pas:294-295`).
- lines 106/109/114/227: `"LIN**BP*…*RC*…"` — legacy emits `…*RC*…*` with a trailing sep (`.pas:347-352`).

The "expected" was computed from the rebuild's `sep.join` model, not the legacy byte stream, so the test
asserts the rebuild agrees with ITSELF on exactly the segments where it disagrees with legacy. That is the
fixture-fidelity / vacuous-parity trap (memory: feedback-parity-fixture-fidelity).

`test_edi856_e2e.py` is honest about what it does NOT prove (no golden 856 -> byte-parity UNPROVABLE,
docstring `:32`/`:44`/`:115`), and its feed-row parity is a REAL cross-check against the legacy SELECT — but:
- it compares only the **4 wire columns** Manifest/Part/Qty/Kanban (`:191`,`:194`), so the legacy 9-col GROUP
  BY vs rebuild 5-col GROUP BY ROW-COUNT divergence (BLOCKER 3) only surfaces if the data contains a
  multi-price / multi-forecast part — and ASN 4721's parts do not, so the test passes vacuously on that axis.
- it never builds with the REAL `VehicleOrder.Site` values (uses placeholder `INV_SITES`), so the ISA byte
  layout is never exercised against TEMA-real input.

Type: **test / parity-method flaw.** The structure assertions (SE01/CTT01/EIN-in-7-positions/HL chain) and
the feed-row parity are sound; the byte-parity claim on TD1/LIN is self-referential.

---

## NIT 6 — ISA `_fixed` hard-truncation is a documented, deliberate divergence (harmless on real data)

`code.py:62-74` hard-truncates over-width fields to keep ISA byte-width; legacy `Format('%-Ns')` PADS only,
never cuts (`.pas:147,149,151,153`). Tested with REAL `VehicleOrder.Site` (`abbr 'MAS'`, DUNS `969009112`,
`969009112-71930` = exactly 15, TMM `808369495`): truncation NEVER fires; ISA is byte-clean (widths
10/10/15/15, ISA09 `260618` yymmdd, ISA15 `P`, ISA16 `#`). Truncation only fires on schema-max pathological
input (`SiteAbbr varchar(25)`, `SiteSupplierCode varchar(50)` in `VehicleOrder.Site`) that legacy would have
mis-emitted anyway. Honestly documented as a "Trap 6 fix." Not a real-data divergence.

Type: deliberate fix, correctly scoped.

---

## NIT 7 — placeholder `INV_SITES` would emit an invalid ISA15; builder does not validate

The spike `INV_SITES` rows carry `VC_EDI_MODE = 'PROD'` (and `VC_SEP_SUBELEMENT '>'`, `VC_SEP_SEGMENT '~'`)
vs the real `VehicleOrder.Site` (`SiteEDIMode 'P'`, sub-elem `#`, seg-term `$`). With the placeholder values
`build_856` emits `…*P*…*PROD*>` — ISA15 = `PROD` (4 chars where X12 mandates single `T`/`P`); the builder
performs no single-char validation. This is the documented placeholder-identity gap (`code.py:313-320`,
`test_edi856_e2e.py:34-38`), correctly scoped as "bytes not wire-correct until real site values load" — but
flag that the real cutover must load `VC_EDI_MODE='P'` (the real value is `P`, so the data is fine; the
builder just won't catch a bad load).

Type: data-loading caveat, honestly documented; builder hardening optional.

---

## What I could NOT refute (rebuild PROVEN equivalent / correct here)

- **EIN allocation (#3 in brief): atomic + site-scoped — CONFIRMED.** `UPDATE INV_SITES SET IN_EIN_SEQ =
  IN_EIN_SEQ + 1 OUTPUT INSERTED.IN_EIN_SEQ WHERE IN_SITE_ID = ?` (`code.py:391-393`) is one statement under
  the row X-lock; two concurrent sends serialize and cannot collide. Rolled-back probe: site 1 0->1, site 2
  stayed 0 (no cross-site bump — fixes legacy `AD_UpdateEIN` no-WHERE D1 hazard). EIN is allocated, stamped
  on the header (`code.py:399-401`), THEN passed to `build_856`, so all 7 control positions share one value.
- **Byte structure (#4): SE01/CTT01/EIN/dates — CONFIRMED with REAL site values.** SE01 = actual ST..SE
  count (independently recomputed = 22), CTT01 = HL count (= 6), BSN02 = 17 chars (`yyyymmdd`+`%09d`), all 7
  EIN positions identical `%09d`, ISA09 `yymmdd` while GS04/BSN03/DTM02 `yyyymmdd`, ISA widths 10/10/15/15,
  CRLF-only, segment terminator never emitted, trailing CRLF present. `_fmt_ein` widen-not-truncate matches
  legacy `%9.9d`; real EIN range stays < 9 digits.
- **Feed shape on the current snapshot: EQUIVALENT.** Legacy 9-col vs rebuild 5-col over ALL ASNs: 14,356 ==
  14,356 rows, 0 ASNs differ. Cost INNER drop, forecast INNER drop, inclusive window (`<=`/`>=`) all match
  legacy on this data.
- **Status flip + side-effect-free feed: CONFIRMED.** The feed SELECT is pure (no UPDATE); the C->S flip is
  per-ASN `WHERE IN_ASN_ID = ?` (`code.py:453-454`), NOT `REPORT_EDI856`'s self-flip
  (`WHERE IN_ASN_EIN=@EIN`) nor the blanket `UPDATE_ASNStatus` (`WHERE VC_ASN_STATUS='C'`). The e2e decoy
  ASN proves a non-target 'C' ASN stays 'C'.

---

## VERDICT

**NOT proven byte-equivalent to the legacy 856.** Two unconditional byte-divergences are present on every
856: `TD1*` vs legacy `TD1**` (BLOCKER 1) and `LIN**BP*<part>*RC*<kanban>` vs legacy
`LIN**BP*<part>*RC*<kanban>*` (BLOCKER 2). The "byte-faithful" claim and the unit test's hand-computed
`expected` array encode the rebuild's own (wrong) bytes, not the legacy `.pas` stream, so the green test is
self-referential on exactly these segments.

The **feed** is equivalent to legacy ON THE CURRENT SNAPSHOT (14,356 == 14,356 rows, 0 ASNs differ), but
carries two LATENT, data-gated divergences that the sparse snapshot masks: dropping `MO_PRICE` from the
GROUP BY (BLOCKER 3) and replacing the legacy forecast INNER join with a `CROSS APPLY TOP 1` that has NO
`ORDER BY` (SHOULD-FIX 4). Both fire the instant a part gains a second overlapping cost window (different
price) or a second forecast row — and the schema (heaps, no uniqueness) permits both. So feed equivalence is
PROVEN only for this data vintage, NOT in general; and the forecast TOP-1 is additionally nondeterministic.

EIN allocation (atomic, site-scoped, stamp-before-build), the byte STRUCTURE with REAL site values
(SE01/CTT01/EIN×7/dates/widths/CRLF), the side-effect-free pure feed, and the decoupled per-ASN C->S flip are
all CONFIRMED correct.

### Key results requested
- **#2 (forecast TOP-1 vs INNER):** the `CROSS APPLY (SELECT TOP 1 …)` has **no `ORDER BY`** over a **HEAP**;
  when a part has >1 forecast row the legacy INNER+GROUP-BY keeps every distinct kanban (proven: 2 rows
  `AAAA`,`JLN9`) while the rebuild keeps ONE, and WHICH one is nondeterministic (plain->`JLN9`, ordered->
  `AAAA`). Equivalent on today's 1:1 forecast data only.
- **#5 (real-site widths):** built with REAL `VehicleOrder.Site` values — ISA is byte-clean, the
  `_fixed` truncation never fires; truncation is a deliberate, documented divergence that only triggers on
  schema-max pathological input legacy would have mis-emitted anyway.

---

# RE-VERIFY (round 2) — 2026-06-20

Re-attack of the developer's claim that **all** round-1 findings are fixed. Method: rebuilt the legacy byte
stream INDEPENDENTLY from `EDI856Object.pas` (a hand-transcribed oracle of every `.pas` string concatenation,
cited per source line — `/tmp/legacy_856_sim.py`), diffed it against `build_856`'s actual output segment-by-
segment; re-read the LIVE `REPORT_EDI856` body from `Inventory_Live`; re-ran the pure unit test
(`scripts/e2e/test_edi856_build.py`, 49/0) and the live-data e2e (`scripts/e2e/test_edi856_e2e.py`, 51/0); and
empirically PATCHED the rebuild back to the old wrong bytes to prove the test is non-vacuous.

Files re-read: `EDI856Object.pas` (the 12,634-byte legacy builder, lines 89-459); `code.py` (the rebuild,
working-tree modified); `edi856-wire-format.md`; `spike-edi856-feed.sql`; both test files; live
`REPORT_EDI856`. NOTE: the fixes live in the WORKING TREE (uncommitted) on `m1-edi856-outbound`; HEAD is still
`18a9963 "...KNOWN BLOCKERS, fix next"`.

## Per-finding resolution

### BLOCKER 1 — TD1 (`TD1*` vs legacy `TD1**`): RESOLVED
`code.py:222` now emits `sep.join(["TD1", "", ""])` → `TD1**` (two empty trailing elements). My independent
`.pas` oracle (lines 294-296: `'TD1'+fSepElement` then `+fSepElement`, nothing appended) produces `TD1**`.
Segment-by-segment diff of rebuild vs oracle: **MATCH** at index [6]. Non-vacuity PROVEN: patched the rebuild
back to `sep.join(["TD1",""])` → emits `TD1*` → the unit-test assertion `td1 == "TD1**" and td1 != "TD1*"`
goes False (test FAILS). The fix holds.

### BLOCKER 2 — LIN (no trailing sep vs legacy `…*RC*<kanban>*`): RESOLVED
`code.py:253-254` now emits `sep.join(["LIN","","BP",part,"RC",kanban,""])` → `LIN**BP*<part>*RC*<kanban>*`
(trailing element separator after the kanban). My `.pas` oracle (lines 347-352: line 352 appends the kanban
`+fSepElement`, so the segment ends in a separator) produces the same. Diff vs oracle: **MATCH** at indices
[12],[15],[20]; all three LINs end in `*`. Non-vacuity PROVEN: patched the rebuild to drop the trailing `""`
→ emits `LIN**BP*42600FEK5000*RC*JZUX` → the assertion `all(s != "…JZUX" …)` goes False (test FAILS). Holds.

### BLOCKER 3 — feed dropped `MO_PRICE` from GROUP BY: RESOLVED
`code.py:343` and `spike-edi856-feed.sql:71` both keep `m.MO_PRICE` in the GROUP BY. Live counterexample
(rolled-back, sentinel-tagged, swept): injected a SECOND price-distinct cost window covering the same prod
date for part `42600FEK5000` on a single-line probe ASN. Legacy `REPORT_EDI856 @EIN=0` SELECT → 2 rows;
rebuild feed → **2 rows, identical** (e2e case #5 PASS). Before the inject both return 1; after sweeping the
2nd window both return 1. The legacy proc body confirms `m.MO_PRICE` is the 4th GROUP-BY column in BOTH
branches. The MO_PRICE split is reproduced.

### SHOULD-FIX 4 — forecast `CROSS APPLY TOP 1` (nondeterministic, dropped lines): RESOLVED (with a NIT)
`CROSS APPLY` is GONE; `code.py:339` / `spike-edi856-feed.sql:67` use `JOIN INV_FORECAST_DETAIL_INF f ON
d.VC_ASSY_PART_NUMBER = f.VC_ASSY_PART_NUMBER_CODE` — the SAME INNER join the live proc uses (verified against
the live body). Live counterexample (rolled-back): injected a 2nd forecast row for `42600FEK5000` with a
DISTINCT kanban `ZZZ9`. Legacy SELECT → 2 rows `{JZUX, ZZZ9}`; rebuild feed → **2 rows, both kanbans,
identical** (e2e case #6 PASS). The drift guard (e2e #8) additionally asserts `"CROSS APPLY" not in _FEED_SQL`.
The line-dropping + heap-TOP-1 nondeterminism is gone. (Residual ordering NIT below.)

### SHOULD-FIX 5 (== BLOCKER 3) — MO_PRICE: see BLOCKER 3. RESOLVED.

### #6 full feed parity on real data: CONFIRMED on this snapshot
E2E (1) ran the legacy `REPORT_EDI856 @EIN=0` SELECT vs the rebuild `_FEED_SQL` on real ASN 4721's 16 detail
rows (copied into the rebuild DB): **16 == 16 rows, row-for-row identical** on Manifest/Part/Qty/Kanban. The
4-table INNER join, inclusive cost window (`<=`/`>=`), and 6-col GROUP BY all match. PROVEN that the rebuild's
6-col GROUP BY is equivalent to legacy's 9-col GROUP BY: `IN_ASN_EIN`, `VC_START_SEQ_NUMBER`, `VC_LINE_NAME`
all come from table `a` filtered to `IN_ASN_ID = ?`, and `IN_ASN_ID` is the PK of `INV_ASN_MST` (verified:
`PK_INV_ASN_MST`), so those three are single-valued per ASN → dropping them is provably cardinality-neutral.

### #5 the unit test is genuinely legacy-derived + non-vacuous: CONFIRMED
The `expected[]` array (`test_edi856_build.py:111-140`) now encodes the LEGACY bytes (`TD1**`, `LIN…*`), and I
independently re-derived the same 26 segments straight from the `.pas` — they match `expected[]` exactly.
Non-vacuity PROVEN twice: (a) the anti-regression block (`:259-265`) rejects the old `TD1*` / `LIN…JZUX`
forms; (b) I patched the live rebuild back to the old bytes and the assertions flip to False (test FAILS).
This is no longer self-referential on these segments.

### #7a ISA15 single-char + the `_FEED_SQL`==.sql drift guard: CONFIRMED
INV_SITES seed is now `VC_EDI_MODE='P'` (was the 4-char `'PROD'` defect) and `VC_SEP_SUBELEMENT='>'` (1 char)
— verified on the spike. ISA15 emits a valid 1-char `P`. The drift guard (e2e #8) PASSES and is real: the
inline `_FEED_SQL` is char-identical to the `.sql` marked body (modulo `@ASNID`→`?`); mutating either (e.g.
dropping `m.MO_PRICE`) makes the compare fail.

## Remaining divergences / caveats (none are byte-defects vs legacy on real data)

- **NIT — residual fan-out ordering (qualifies the "NO nondeterminism" claim for #4).** The feed
  `ORDER BY Manifest, PartNumber` does NOT tie-break sibling fan-out rows that share the same
  (Manifest, PartNumber) — e.g. the two kanbans `JZUX`/`ZZZ9` for one manifest+part have IDENTICAL ORDER BY
  keys (shown on the spike), so their relative order (and thus the HL-id-to-LIN assignment) is engine-defined.
  This is NOT worse than legacy (the proc has NO ORDER BY at all), and the 856 wire is indifferent to which
  sibling gets which HL id as long as each LIN/SN1 carries the right part/kanban/qty — but the developer's
  "with NO nondeterminism" wording is slightly overstated: cross-manifest order is now pinned, intra-tie order
  is not. Code-precision NIT, not a code defect.
- **NIT — `_one_char` / `_fixed` are defensive deviations, not faithful ports, on PATHOLOGICAL input.** Legacy
  `.pas:161` emits `SiteEDIMode` raw; on a 4-char `'PROD'` legacy would emit `PROD` (malformed), the rebuild
  emits `P`. Likewise `_fixed` hard-truncates over-wide ISA fields legacy would have over-run. On REAL data
  (1-char mode, in-range DUNS/abbr) both are byte-identical; these only diverge on input legacy would have
  mis-emitted anyway. Documented in the code as Trap-6 / defect-#7. Acceptable, but it means the port is
  "byte-faithful on real data," not "a literal transcription on all input."
- **DATA-VINTAGE GAP (unchanged from round 1) — full LEGACY BYTE parity remains UNPROVABLE.** No golden 856
  exists for any ASN, and the spike `INV_SITES` carries PLACEHOLDER identity (abbr `MAS`, DUNS `000000001`,
  supplier `MAS`, TMM `000000011`, sub-elem `>`) — NOT the real TEMA wire values (DUNS `969009112`, supplier
  `71930`, TMM `808369495`, sub-elem `#`). So the ISA/GS IDENTITY bytes the spike emits are NOT the TEMA wire.
  The STRUCTURE is value-independent and faithful; the exact identity bytes can only be confirmed once the
  real site row is loaded at cutover and diffed against a captured legacy 856. This is an honest, documented
  gap (test docstrings state it plainly), not a defect.

---

## VERDICT (round 2)

All four round-1 byte/feed findings are **RESOLVED**, each with a counterexample-grade proof:

- **BLOCKER 1 (TD1):** RESOLVED — `TD1**` matches the `.pas` oracle; reverting fails the test.
- **BLOCKER 2 (LIN):** RESOLVED — `LIN…*RC*<kanban>*` matches the `.pas` oracle; reverting fails the test.
- **BLOCKER 3 / SHOULD-FIX 5 (MO_PRICE):** RESOLVED — `m.MO_PRICE` back in GROUP BY; live 2-cost-window probe
  yields 2 rows in BOTH legacy and rebuild.
- **SHOULD-FIX 4 (forecast TOP-1):** RESOLVED — INNER `JOIN INV_FORECAST_DETAIL_INF` (matches the live proc);
  live 2-kanban probe yields 2 rows / both kanbans in BOTH; CROSS APPLY gone; drift-guarded.

The rebuild is now **byte-faithful to the legacy 856 in STRUCTURE + every segment + the data feed, on real
data, MODULO the documented byte-parity-pending-golden gap** (no golden file + placeholder site identity on
the spike). I attempted to refute byte-faithfulness with an independent `.pas`-derived oracle (26/26 segments
match), a non-vacuity patch (old bytes now fail the test), the live proc body (feed joins/GROUP-BY/window all
match), and rolled-back fan-out/split counterexamples (legacy == rebuild row-for-row) — and could NOT.

Two NITs remain, neither a byte-defect vs legacy on real data: (1) the "NO nondeterminism" claim is slightly
overstated — intra-(manifest,part) fan-out sibling order is still engine-defined (no worse than legacy);
(2) `_one_char`/`_fixed` are defensive deviations on pathological input legacy would have mis-emitted.

**Bottom line:** the specific fixes HOLD. Equivalence is PROVEN for structure + segments + feed on real data;
true end-to-end LEGACY BYTE parity stays UNPROVABLE until a golden 856 + real site values are loaded at
cutover — say so; do not call the green spike tests "TEMA byte parity."

---

## RE-VERIFY (round 2) — INDEPENDENT SECOND-ADVERSARY CORROBORATION — 2026-06-20

A second adversary pass re-attacked the four findings by a DIFFERENT method than the harness above (raw
`sqlcmd` rolled-back probes + direct `.pas` line reads, not the Python test fixtures), to guard against the
test and the code agreeing with each other vacuously. Result: every round-1 finding independently RESOLVED;
no new divergence found. New corroborating evidence:

### Segment bytes (BLOCKER 1/2 + the full envelope) — re-read the `.pas` directly
Read the legacy emitters line-by-line and matched them to `build_856`'s `sep.join` arguments:
- TD1 `EDI856Object.pas:294-296` (`'TD1'+fSepElement` then `+fSepElement`, nothing appended) -> `TD1**`;
  rebuild `code.py:222` `sep.join(["TD1","",""])` -> `TD1**`. MATCH.
- LIN `:347-352` (line 352 appends `Kanban+fSepElement`, trailing sep) -> `...*RC*<kanban>*`; rebuild
  `code.py:253-254` appends a trailing `""` -> same. MATCH.
- Envelope cross-checked against the `.pas` too (not just the harness): ISA `:144-162` (ISA09 `copy(,3,6)`
  yymmdd; ISA16 `:162 fSepSubElement`, NO trailing sep), GS `:179-187` (`+'004010'`, no trailing), ST
  `:204-206`, BSN `:225-229` (`:229` hhmm with `//+fSepElement` COMMENTED OUT -> no trailing sep), DTM
  `:248-252` (`+'ET'`), CTT `:386-387`, SE `:406-409` (`INC(fSegCount)` at `:407` BEFORE the append -> SE
  counts itself), GE `:427-429`, IEA `:447-449`. Emission order `Execute:116-126` == the rebuild's `segs`
  order, and the legacy `RecordCount>0` guard (`:113`) == the driver's `if not detailRows: raise`. No envelope
  segment diverges. Non-vacuity re-confirmed independently: patched a throwaway copy of `code.py` back to
  `["TD1",""]` / LIN-without-trailing-`""` and the unit test's anti-regression assertions
  (`test_edi856_build.py:261-265`) evaluate to False (i.e. would FAIL).

### Feed MO_PRICE + forecast join (BLOCKER 3 / SHOULD-FIX 4) — raw sqlcmd fan-out probe, NOT the harness
Ran ONE rolled-back `sqlcmd` transaction on `Inventory` (single-line probe ASN, prodDate `20260618`, part
`42600FEK5000`) injecting BOTH a 2nd price-distinct cost window AND a 2nd distinct-kanban (`ZZK9`) forecast
row, then counted rows four ways:
- legacy proc-body **9-col** GROUP BY (transcribed from the LIVE `REPORT_EDI856` `@EIN=0` branch) -> **4 rows**
- rebuild feed **6-col** GROUP BY (`m.MO_PRICE` kept, INNER forecast join) -> **4 rows** (MATCH: 2 prices x 2 kanbans)
- OLD bug A (MO_PRICE dropped from GROUP BY) -> **2 rows** (collapses the price split)
- OLD bug B (`CROSS APPLY (SELECT TOP 1 ...)`, no ORDER BY) -> **2 rows** (drops the kanban fan-out)

Against the legacy truth of 4, the fixed feed = 4 while BOTH old bugs = 2 — the fixes are the difference,
proven OUTSIDE the test harness. Tran rolled back; sentinel rows swept; verified 0 leftover and
`INV_SITES.IN_EIN_SEQ` restored to its as-found value (0).

### Structural facts the GROUP-BY-equivalence + nondeterminism arguments rest on — verified live
- `PK_INV_ASN_MST` is on `IN_ASN_ID` ALONE (queried `sys.indexes`/`sys.index_columns`). So under
  `WHERE IN_ASN_ID = ?` table `a` is exactly one row -> `IN_ASN_EIN`/`VC_START_SEQ_NUMBER`/`VC_LINE_NAME` are
  single-valued -> dropping them from the rebuild's GROUP BY is provably cardinality-neutral (legacy 9-col ==
  rebuild 6-col). The round-1 BLOCKER-3 collapse came ONLY from dropping `m.MO_PRICE` (from the cost join,
  which CAN multi-value), never from these three header columns.
- `INV_FORECAST_DETAIL_INF` is a **HEAP** (no clustered index) on `Inventory_Live` — confirming the retired
  `CROSS APPLY (SELECT TOP 1 ...)` had genuinely engine-defined ("nondeterministic") row choice, the round-1
  SHOULD-FIX-4 hazard. The INNER join removes the TOP-1 choice entirely.

### Harness re-run (independent invocation)
`test_edi856_build.py` -> **49 PASS / 0 FAIL**; `test_edi856_e2e.py` -> **51 PASS / 0 FAIL** (feed-row parity
16==16 on real ASN 4721; multi-price #5 and multi-kanban #6 fan-out parity; EIN-from-`INV_SITES`; per-ASN
decoupled C->S flip; temp-then-rename atomicity). The e2e's "legacy SELECT" oracle (`:230-242`, `:398-413`) was
checked column-for-column against the LIVE `REPORT_EDI856 @EIN=0` body and is a faithful transcription (same 4
INNER joins, same inclusive `<=`/`>=` window, same 9-col GROUP BY) — the parity is a real legacy-vs-rebuild
diff, not self-comparison.

### Verdict (independent corroboration)
I tried to refute byte-faithfulness a second time, by hand off the `.pas` and by raw rolled-back SQL against
the live proc, and **could not**. All four round-1 findings are RESOLVED; the two NITs above stand (intra-tie
fan-out order engine-defined — no worse than the legacy proc, which has NO `ORDER BY` at all;
`_one_char`/`_fixed` defensive on pathological input legacy would have mis-emitted). RECOMMENDATION: the
byte/feed logic is sound to SHIP — with the explicit, unchanged caveat that **TRUE end-to-end legacy BYTE
parity is still UNPROVABLE** until (a) a captured golden 856 and (b) the REAL `INV_SITES` identity values
(DUNS/supplier/TMM/sub-elem `#`, EDI mode) replace the spike placeholders at cutover. The green tests prove
structure + segments + feed equivalence ON REAL DATA; they do NOT prove "TEMA byte parity" and must not be
labeled as such. Housekeeping note (not a code defect): the fixes are UNCOMMITTED working-tree changes; HEAD is
still `18a9963 "...KNOWN BLOCKERS, fix next"` — commit them so the SHIP state is the recorded state.
