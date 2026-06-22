# Adversary findings — hot-call ASN gap build (branch `m1-hotcall-entry`)

**Reviewer:** sql-adversary · **Date:** 2026-06-22 · **Default stance:** the reimplementation is wrong
until proven equivalent to the legacy on the SAME inputs.

**Scope of refutation:**
1. Is the `8HC` filename byte-faithful to the legacy?
2. Does `create_hotcall_asn` produce correct hot-call ASN rows that flow through the 856 as `8HC…` / M390?

**Artifacts under review:**
- Legacy anchor: `ASNInvoice.pas:817-826` (the *recreate-button* 8HC/856 branch)
- Legacy create: `HotCallEntry.pas` (the "One Cycle Entry" form)
- Rebuild filename: `docs/analysis/edi/856/project-library/edi856/code.py` (`_filename_856`, `send_856`)
- Rebuild create: `docs/analysis/edi/project-library/hotcall/code.py` (`create_hotcall_asn`)
- Tests: `scripts/e2e/test_hotcall_build.py`, `scripts/e2e/test_hotcall_e2e.py`
- Re-key: `docs/analysis/edi/spike-asndetail-rekey.sql`

All spike work was bounded, READ-ONLY on `Inventory_Live`/`VehicleOrder`, rolled-back / self-cleaning on
`Inventory`, `SET QUOTED_IDENTIFIER ON` (the R18 filtered-index trap), and left the spike as-found.

---

## BLOCKER-1 — The 8HC filename is faithful to the WRONG legacy branch; the operational send path appends the LineName (and a per-batch counter), the rebuild does not.

**Claim under test:** "the `8HC` filename is byte-faithful to the legacy" and "send_856 produces an `8HC…`
file that flows like the legacy."

**Counterexample (proven end-to-end on the spike, rolled back):**

There are THREE legacy code paths that build an 856 hot-call filename, and they are NOT identical:

| Path | file:line | Hot-call (`StartSeq='-1'`) filename | Flips C→S? |
|---|---|---|---|
| Recreate button | `ASNInvoice.pas:825` | `'8HC'+copy(pd,4,5)+'1.txt'` | **NO** |
| **Operational bulk send** | `MainMenu.pas:2722-2723` | `'8HC'+copy(pd,4,5)+IntToStr(y)+LineName+'.txt'` | **YES (`UpdateASNStatus`, :2747-2748)** |
| (normal counterpart) | `MainMenu.pas:2718` | `'856'+copy(pd,5,4)+LineName+'.txt'` | YES |

`send_856` allocates the EIN, writes the file, and **flips the ASN C→S** (`edi856/code.py:537-538`). The
C→S flip is the discriminator: the *recreate button* does NOT flip status (verified: no status update in
`ASNInvoice.pas:828-856`); the *operational `MainMenu.ResendMarkedEDIs`* path DOES (`MainMenu.pas:2747-2748`).
**Therefore `send_856` reproduces the OPERATIONAL path (`MainMenu:2722`), but its `_filename_856` was
anchored on the RECREATE button (`ASNInvoice.pas:825`).**

The operational path appends two things the rebuild omits:
- `IntToStr(y)` — a **per-batch counter** that starts at 1 and is `INC(y)`'d per hot-call in the run
  (`MainMenu.pas:2724`). The rebuild's `1` is hardcoded as a literal suffix; it is correct ONLY for the
  first hot-call in a batch.
- `EDI856.LineName` — the ASN's `VC_LINE_NAME`, sourced by the report `REPORT_EDI856` which projects
  `a.VC_LINE_NAME AS 'LineName'` (verified live), set onto the object at `MainMenu.pas:2712`.

Real hot-call ASNs carry a non-empty line name. Live `Inventory_Live.INV_ASN_MST` (READ-ONLY):
all 315 `VC_START_SEQ_NUMBER='-1'` headers carry `VC_LINE_NAME='COROLLA'` (e.g. ids 4722, 4712, 4704…).

End-to-end probe on the spike (created a hot-call ASN, line `COROLLA`, prodDate `20211201`, ran the REAL
`send_856`, then cleaned up; the spike's only residual `COROLLA/20211201/-1` row is the pre-existing LIVE
id 3184, status `'A'`, untouched):

```
HEADER VC_LINE_NAME  = 'COROLLA'
HEADER START_SEQ     = '-1'  (hot-call sentinel)
REBUILD filename     = '8HC112011.txt'
LEGACY recreate(:825)= '8HC112011.txt'           match: True
LEGACY operational(MainMenu:2722, y=1) = '8HC112011COROLLA.txt'   match: False
```

**Input → legacy vs rebuild:** prodDate `20211201`, line `COROLLA`, first hot-call in a batch →
- legacy (operational, the path send_856 reproduces): `8HC112011COROLLA.txt`
- rebuild: `8HC112011.txt`  ← **missing the `COROLLA` line-name suffix.**

For the second hot-call in a batch the gap widens to BOTH the counter and the line:
legacy `8HC112012COROLLA.txt` vs rebuild `8HC112011.txt` (rebuild would also COLLIDE on the filename for
two same-date hot-calls because its `1` never increments — the legacy `IntToStr(y)` exists precisely to
keep multiple hot-calls in one batch from overwriting each other's file).

**Why it matters:** the build's own analysis says the dispatcher/TEMA receiving side "may key on" the
hot-call filename and calls it "customer-visible" (`hotcall-coverage-analysis.md:178-185`). A filename that
drops the line name and never increments the counter is a different byte string than what TEMA has been
receiving for years, and two same-day hot-calls would overwrite to one file on disk. The 8HC prefix and the
`copy(pd,4,5)` offset ARE correct; the suffix is not.

**Classification:** code defect (rebuild anchored on the wrong legacy branch) + a documentation root-cause
(`hotcall-coverage-analysis.md:178-185` cites only `ASNInvoice.pas:817-825`, never `MainMenu.pas:2718/2722`).

**Caveat (honest):** which legacy filename is "the" target is, strictly, a David decision — the recreate
button and the operational sender genuinely disagree in the legacy itself, and TEMA may ignore the filename
entirely. But the rebuild cannot claim "byte-faithful to the legacy" while matching only the path that does
NOT do the C→S flip it reproduces. At minimum this must be raised as a decision, not shipped silently.

---

## BLOCKER-2 — The parity tests pass GREEN on the BLOCKER-1 divergence: the EXPECTED 8HC string is anchored on the wrong branch and never asserts the line/counter suffix.

**Claim under test:** "the test derives the expected FROM the .pas (not the rebuild) + is non-vacuous."

The test *is* non-vacuous against the pre-fix always-`'856'` revert (`test_hotcall_build.py:119-127`), and
the EXPECTED IS transcribed from a `.pas` line (`_legacy_8hc_from_pas`, `:57-69`) rather than from the
rebuild — so it is not self-referential in the classic sense. **But it transcribes the WRONG `.pas` line.**
`_legacy_8hc_from_pas` mirrors `ASNInvoice.pas:825` (recreate, `'8HC'+copy(pd,4,5)+'1.txt'`), which is NOT
the operational path `send_856` reproduces. Consequences:

- `test_hotcall_build.py:99` asserts `_filename_856('20260606','-1') == '8HC606061.txt'` — green, but that
  is the recreate byte, not the operational `8HC606061COROLLA.txt`.
- `test_hotcall_e2e.py:225-228` asserts `sent['filename'] == '8HC'+PDATE[3:8]+'1.txt'` — it builds the
  ASN with line `ZZHCE2E` and runs the REAL `send_856`, yet asserts a filename with NO line suffix. The
  test would PASS even though the operational legacy would have produced `…1ZZHCE2E.txt`. The single check
  that could have caught BLOCKER-1 (assert the filename carries the line name) is absent.

I ran both tests: `test_hotcall_build.py` = **26 PASS / 0 FAIL**; the divergence is invisible to them.

**Non-vacuity proof of MY finding:** the spike probe above shows `_filename_856` returns `8HC112011.txt`
where the operational legacy returns `8HC112011COROLLA.txt` — a string the current assertions never compare
against. Revert the line/counter into the test's expected and the e2e test FAILS; that is the missing check.

**Classification:** test / parity-method flaw (correct oracle technique, wrong oracle source line).

**Required fix:** re-derive the 8HC oracle from `MainMenu.pas:2722` (`'8HC'+copy(pd,4,5)+IntToStr(y)+
LineName+'.txt'`), add the LineName + counter to `_filename_856`'s signature, and assert the e2e filename
carries the line name + the per-send counter. Then prove non-vacuity (revert the suffix → e2e FAILS).

---

## SHOULD-FIX-1 — The NORMAL 856 filename also diverges from the operational legacy (decision-E vs MainMenu:2718).

Not the hot-call focus, but it compounds. `_filename_856` normal branch = `'856'+copy(pd,4,5)+'.txt'`
(decision E, the CREATE pattern `ASNSelect.pas:457`). The operational normal send is
`MainMenu.pas:2718` = `'856'+copy(pd,5,4)+LineName+'.txt'`. For `20260606`:

- rebuild: `85660606.txt`  (`copy(,4,5)` = `60606`)
- operational legacy: `8560606COROLLA.txt`  (`copy(,5,4)` = `0606`, plus LineName)

So both the OFFSET (`4,5` vs `5,4`) and the LineName differ on the normal path too. The build deliberately
locked decision-E (`copy(,4,5)`) as "the one deterministic normal choice", which is defensible IF David
ratified that the normal filename changes — but the same omitted-LineName issue as BLOCKER-1 applies.

**Classification:** ratified divergence (offset) layered with an unratified omission (LineName) — surface
to David alongside BLOCKER-1 as one filename decision.

---

## CONFIRMED-FAITHFUL (refutation attempts FAILED — these are correct)

These I actively tried to break and could not:

- **8HC prefix + `copy(pd,4,5)` offset + the `1` position.** `copy(s,4,5)` is 1-based start 4, length 5 →
  chars 4..8 = `prodDate[3:8]` 0-based. Verified char-by-char for `20260606`→`60606`, `20260529`→`60529`,
  `20271225`→`71225`. The `8HC` prefix, the 5-char date slice, and a `1` immediately before `.txt` are all
  byte-correct vs `ASNInvoice.pas:825`. (The DEFECT is the MISSING LineName after the `1`, not these.)
- **`@HotCall=1` always-INSERT keeps distinct/duplicate parts.** Live proc body (`spike-asndetail-rekey.sql`
  matches) — `@HotCall=1` → unconditional INSERT, no manifest accumulate. e2e proven: two PART_A rows under
  one manifest persist as qty 1 and qty 3 (NOT a single accumulated 4).
- **The `-1` sentinel + `INSERT_ASNInfo`.** Live proc: `@StartSeq`/`@EndSeq` are `varchar(4)` → passing the
  `'-1'` STRING is correct; status hardcoded `'C'` in `VALUES(@Ein,'C',…)`; `SET @ASNID = SCOPE_IDENTITY()`
  OUTPUT → the rebuild's DECLARE/EXEC `@ASNID=@id OUTPUT`/`SELECT @id` capture is the right pattern (a bare
  EXEC's SCOPE_IDENTITY would be NULL across the child scope).
- **Header `IN_QTY` = SUM(detail).** Live `INSERT_ASNInfo` writes `@Qty` into `IN_QTY`. The legacy passes
  the stale loop var (`HotCallEntry.pas:246-247`, `@QTY := qty`) — genuine garbage (last-validated qty). The
  P14 fix (sum) is a benign divergence: 856/810 read DETAIL qty, not header qty. No wire impact. FAITHFUL-
  enough + safer.
- **EIN-at-send bumps once.** `send_856` does `UPDATE INV_SITES SET IN_EIN_SEQ=IN_EIN_SEQ+1 OUTPUT
  INSERTED…` site-scoped, atomic, on the tx. Probe: seeded 9300 → allocated 9301. `create_hotcall_asn`
  correctly writes `IN_ASN_EIN=0` at create and drops the legacy `AD_UpdateEIN` (would double-allocate
  under the at-send model). Ratified M1 divergence, consistent.
- **Unique guard never trips for hot-calls.** `UX_INV_ASN_MST_LINE_PDATE_NORMAL` has
  `filter_definition = ([VC_START_SEQ_NUMBER]<>'-1')` (verified). Rolled-back probe: TWO hot-calls on the
  same line+pdate both inserted (count=2); a NORMAL (real-seq) duplicate on the same line+pdate was BLOCKED
  (`Cannot insert duplicate key … UX_INV_ASN_MST_LINE_PDATE_NORMAL`). The filtered index behaves exactly as
  claimed.
- **Flows through the 856 feed by `IN_ASN_ID` alone.** Feed SQL is `WHERE a.IN_ASN_ID = ?` with no status /
  seq / `'7'` filter; a hot-call line survives the INNER cost + forecast joins (live: id 4712 →
  `Manifest=52089698, Part=42600FEL2000, ShipQty=1`). e2e: the 856 carries LIN/SN1 per part + PRF carrying
  the operator manifest. CONFIRMED.
- **M390 routing.** NOTE the task framing "flows through the 856 as … M390" conflates two layers: M390/M391
  is the **810** `IT101` (`edi810/code.py:287`, `M391 if manifest[:1]=='7' else 'M390'`); the **856 emits NO
  IT1/M390 segment at all** (verified: neither `EDI856Object.pas` nor the rebuild 856 builder references
  M390/M391/IT1 — the manifest rides the 856 only via PRF). A hot-call's operator-typed non-'7' manifest →
  M390 in the 810, correctly; the `_validate_manifest` guard (len≥8, reject `'7'`-prefix) enforces it at
  entry. FAITHFUL.
- **Validation matches the pre-BeginTrans exits.** `_validate_manifest` (len≥8, `HotCallEntry.pas:157`) and
  `_validate_item` (qty numeric>0 `:184-197`, part-required `:204-218`) all run BEFORE
  `beginTransaction`, mirroring the Delphi `exit` guards that precede `:221 BeginTrans`. The numeric-manifest
  check is correctly NOT added (commented out in the .pas, `:150-155`).

---

## VERDICTS

**(1) Is the 8HC filename byte-faithful (modulo golden-8HC pending)?**
**NO — not for the path `send_856` reproduces.** The prefix `8HC`, the `copy(pd,4,5)` date slice, and the
`1` are byte-correct vs `ASNInvoice.pas:825` (the recreate button). But `send_856` does the C→S flip, so it
reproduces the OPERATIONAL sender `MainMenu.pas:2722`, whose filename is `'8HC'+copy(pd,4,5)+IntToStr(y)+
LineName+'.txt'`. The rebuild OMITS the `LineName` suffix and hardcodes the counter as a literal `1`.
Proven on the spike: rebuild `8HC112011.txt` vs operational legacy `8HC112011COROLLA.txt` (line `COROLLA`).
**8HC byte-confirmation:** `_filename_856('20260606','-1') == '8HC606061.txt'` is correct ONLY against the
recreate branch; the operational byte is `8HC606061<LineName>.txt` (and `…<y><LineName>.txt` for batch
position `y>1`).

**(2) Does create_hotcall_asn produce correct hot-call ASNs that flow through the 856 as 8HC/M390?**
**Rows: YES. Filename: NO.** The created header (status `'C'`, `-1/-1` sentinel, `VC_ASSEMBLY_LINE=''`,
`IN_ASN_EIN=0`, `IN_QTY=SUM(detail)`) and the per-part `@HotCall=1` always-INSERT detail rows are faithful
to `HotCallEntry.pas` and persist correctly; they survive the 856 feed by `IN_ASN_ID`, emit LIN/SN1 + a PRF
carrying the operator manifest, route to M390 in the 810 (the 856 carries no M390 segment), bypass the
unique guard, and allocate the EIN once at send. The ONLY break is the outbound **filename** (BLOCKER-1),
which the parity tests fail to catch (BLOCKER-2).

**Fix needed:**
1. Re-derive the 8HC oracle from `MainMenu.pas:2722` (operational sender), not `ASNInvoice.pas:825`
   (recreate). Add `lineName` + a per-send/per-batch counter to `_filename_856` and have `send_856` pass
   `VC_LINE_NAME`. (Bounce the build to `ignition-developer`; the offset/normal-pattern decision-E and
   "which legacy filename is canonical / does TEMA key on it" go to the architects / David.)
2. Make the e2e test assert the filename carries the LineName + counter, with a non-vacuity revert.
3. Correct `hotcall-coverage-analysis.md:178-185` to cite the operational `MainMenu.pas:2718/2722` paths,
   not only the recreate button.

**Equivalence status:** the hot-call **rows** are PROVEN equivalent to the legacy create; the hot-call
**8HC filename** is PROVEN to DIVERGE from the operational legacy send path on real (line-bearing) hot-call
data. The "which filename is canonical" question is a David decision the build silently assumed.

---

## RE-VERIFY (round 2) — 2026-06-22 · re-anchor of BOTH filename branches to the operational sender

**Trigger:** BLOCKER-1/BLOCKER-2 (8HC anchored on the recreate button) + SHOULD-FIX-1 (normal branch wrong
offset + missing LineName). The dev re-anchored BOTH branches to the OPERATIONAL sender
(`MainMenu.pas:2718` normal / `:2722-2724` hot-call) and added the LineName + a deterministic counter.
Re-attacked each item; every claim is RESOLVED-with-proof or flagged. Spike left as-found (all probes
rolled back / self-cleaning; READ-ONLY on `Inventory_Live`).

### Anchor re-read (the legacy operational sender, `MainMenu.ResendMarkedEDIsClick`)
- `:2718` NORMAL: `'\856'+copy(EDI856.PickupDate,5,4)+EDI856.LineName+'.txt'`
- `:2722-2723` HOT-CALL: `'\8HC'+copy(EDI856.PickupDate,4,5)+IntToStr(y)+EDI856.LineName+'.txt'`; `:2724 INC(y)`
- `:2702` `y:=1` (per-batch init); `:2712` `EDI856.LineName := EDI856DataSet.FieldByName('LineName')`;
  `:2747-2748` the C->S flip — confirming this (not the recreate button) is the path `send_856` reproduces.

### 1. NORMAL filename — RESOLVED (byte-proof)
`_filename_856` normal branch (`edi856/code.py:330`) = `"856" + prodDate[4:8] + lineName + ".txt"`.
Independent Pascal-`copy` derivation: `copy('20260618',5,4)` returns chars 5..8 = `'0618'` = `pd[4:8]`
0-based (verified char-identical, NOT the old `[3:8]`/`copy(,4,5)`). For `20260618`/COROLLA the rebuild
returns **`8560618COROLLA.txt`** == the `:2718` legacy byte. Matrix `20260606/20260529/20271225` all match.
- ASSERTED + non-vacuous: `test_edi856_build.py:298-314` asserts the exact byte AND that the old
  `85660618.txt` (wrong `[3:8]` + no LineName) is no longer emitted; `test_hotcall_build.py:107-111`
  diffs the rebuild against an INDEPENDENT `.pas` port `_legacy_856_from_pas` (`:53-68`). The real-driver
  e2e (`test_edi856_e2e.py:357`) emits `8560618ZZ856E2E.txt` (MMDD + LineName) on the copied ASN 4721.
- Non-vacuity re-proven by me: reverting the branch to the PR#29 form (`856`+`[3:8]`+no-LineName) yields
  `85660618.txt != 8560618COROLLA.txt` -> the assertion FAILS. The green is meaningful.

### 2. HOT-CALL filename — RESOLVED (byte-proof) + asymmetry preserved
`_filename_856` hot-call branch (`:328`) = `"8HC" + prodDate[3:8] + str(int(counter)) + lineName + ".txt"`.
`copy('20260618',4,5)` returns chars 4..8 = `'60618'` = `pd[3:8]` (verified char-identical). For
`20260618`/y=1/COROLLA the rebuild returns **`8HC606181COROLLA.txt`** == the `:2722-2723` legacy byte.
- OFFSET ASYMMETRY PRESERVED: normal `pd[4:8]` (=`0618`, MMDD) vs hot-call `pd[3:8]` (=`60618`, Y+MMDD) —
  the leading `6` (last year digit) appears in the 8HC name only, exactly the legacy asymmetry.
- ASSERTED + non-vacuous: `test_hotcall_build.py:114-123` diffs against the INDEPENDENT `.pas` port
  `_legacy_8hc_from_pas` (`:71-87`); `:152-168` runs the reverted (recreate-anchored, no-LineName) form
  through the SAME assertion and proves it FAILS. The real-driver e2e (`test_hotcall_e2e.py:229-234`)
  emits `8HC606181ZZHCE2E.txt` (incl. LineName + y=1) from a created hot-call ASN.
- Non-vacuity re-proven by me: reverting to `8HC`+`[3:8]`+literal-`1`+no-LineName yields `8HC606181.txt`
  (or with LineName-but-no-counter `8HC60618COROLLA.txt`) != `8HC606181COROLLA.txt` -> FAILS.

### 3. LineName source — RESOLVED
`send_856` reads `VC_LINE_NAME` in the SAME header SELECT as `VC_START_SEQ_NUMBER`
(`edi856/code.py:498-508`: `SELECT VC_PRODUCTION_DATE, VC_START_SEQ_NUMBER, VC_LINE_NAME FROM INV_ASN_MST
WHERE IN_ASN_ID=?`), on the tx, and passes it to `_filename_856`. This is the ASN HEADER's `VC_LINE_NAME`.
- Matches the operational sender's source: `EDI856DataSet.FieldByName('LineName')` (`MainMenu.pas:2712`) is
  fed by `REPORT_EDI856`, whose `@EIN=0` branch projects (verified live, `Inventory_Live`):
  `a.VC_LINE_NAME 'LineName' FROM INV_ASN_MST a` — the header alias, the same column the rebuild reads.
- Real-data flow CONFIRMED: all 316 live hot-call ASNs (`VC_START_SEQ_NUMBER='-1'`) carry
  `VC_LINE_NAME='COROLLA'` (e.g. ids 4722/4712/4704). NULL is mapped to `''` (`:508`); empty/None
  lineName -> the same byte the Pascal `+''` would emit. So `COROLLA` flows verbatim into the 8HC byte.

### 4. The counter — RESOLVED (deterministic, collision-free, correctly scoped)
`send_856` (`:546-551`) for a hot-call sets
`counter = 1 + COUNT(*) WHERE VC_START_SEQ_NUMBER='-1' AND VC_ASN_STATUS='S' AND VC_LINE_NAME=? AND
VC_PRODUCTION_DATE=? AND IN_ASN_ID<>?`, read ON the tx (the in-flight 'C' ASN is excluded by `IN_ASN_ID<>?`).
- Rolled-back fixture probe on `Inventory` (3 in-flight 'C' ASNs + 1 prior 'S' on ZZL1/20260618):
  - ZZL1/20260618 (one prior 'S' same line+day) -> **counter 2** (the 2nd file; NO overwrite of the y=1 file)
  - ZZL2/20260618 (different line, same day)     -> **counter 1** (cross-LINE does NOT bleed)
  - ZZL1/20260619 (same line, different day)     -> **counter 1** (cross-DAY does NOT bleed)
  - post-rollback: 0 probe rows remained (spike as-found).
- SCOPE is same-day AND same-line — the right granularity. The legacy per-batch `y` resets each
  `ResendMarkedEDIs` run; the rebuild's "Nth hot-call of the day for THIS line" gives the same property the
  `y` existed to guarantee: two same-day same-line hot-calls get distinct counters 1 then 2 ->
  collision-free (the e2e `test_hotcall_e2e.py:259-271` proves it end-to-end: `8HC606181ZZHCE2E.txt` then
  `8HC606182ZZHCE2E.txt`). A single hot-call/day/line -> counter 1, matching legacy single-send `y=1`.
- FLAG (P13, golden-pending) — UNCHANGED from round 1: the exact `y` RANGE the legacy actually produced
  (vs send order, multi-batch days, lines other than COROLLA) is UNVERIFIABLE until a golden `8HC` file
  exists. The PATTERN, the offset, the LineName, and the no-collision property are proven; the precise
  counter VALUE in a multi-batch/multi-hot-call legacy day is the one residual cutover check. (Today live
  data is single-line COROLLA, typically one hot-call/day -> counter 1, the dominant real case.)

### 5. No regression — RESOLVED
- Normal 856 still flows: `test_edi856_e2e.py` = **52 PASS / 0 FAIL** (ASN-4721 copy -> `8560618ZZ856E2E.txt`,
  feed-row parity vs the legacy `REPORT_EDI856 @EIN=0` SELECT 16/16, EIN-at-send 9101, per-ASN decoupled
  flip, MO_PRICE split #5, kanban fan-out #6, post-write atomicity).
- Hot-call flows through `send_856` -> 8HC + counter + LineName: `test_hotcall_e2e.py` = **32 PASS / 0 FAIL**.
- Pure builders: `test_edi856_build.py` **52/0**, `test_hotcall_build.py` **31/0** (incl. the new normal-
  filename assertion + both non-vacuity reverts).
- The original BLOCKER-1 (missing LineName / wrong anchor) and BLOCKER-2 (wrong oracle line) are RESOLVED:
  the oracle is now re-derived from `MainMenu.pas:2718/2722` and both branches assert LineName + counter
  with a proven-failing revert.

### RE-VERIFY VERDICT
**BOTH the normal and the hot-call 856 filenames are now BYTE-FAITHFUL to the operational sender**
(`MainMenu.pas:2718` / `:2722-2724`), modulo the single golden-`y`-range cutover check (P13):
- NORMAL: `"856" + prodDate[4:8] + LineName + ".txt"` == `copy(pd,5,4)` byte — PROVEN (`8560618COROLLA.txt`).
- HOT-CALL: `"8HC" + prodDate[3:8] + str(counter) + LineName + ".txt"` == `copy(pd,4,5)+IntToStr(y)+LineName`
  byte — PROVEN for the verifiable counter values (`8HC606181COROLLA.txt`; offset asymmetry preserved).
- LineName source PROVEN (= the header `a.VC_LINE_NAME`, the `REPORT_EDI856` projection the sender uses).
- Counter PROVEN deterministic, correctly scoped (same-day + same-line), and collision-free; only the exact
  multi-batch `y` VALUE remains golden-pending (a cutover check, not a code defect).

BLOCKER-1, BLOCKER-2 and SHOULD-FIX-1 are all RESOLVED. The only residual is the data-vintage gap (no golden
8HC -> the precise legacy counter range in a multi-hot-call day is unprovable from available data; SAY SO,
do not call self-consistency parity). Equivalence on the filename is PROVEN for every input verifiable today.
