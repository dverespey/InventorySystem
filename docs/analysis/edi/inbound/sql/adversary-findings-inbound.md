# Adversary findings — M1 inbound 997/824 processor vs the legacy EDIUpload.pas

**Stance:** the reimplementation is wrong until proven equivalent. A finding is only real with a
file:line or a counterexample (input → legacy vs rebuild) + its output.

**Under review:** `docs/analysis/edi/inbound/project-library/edi_inbound/code.py` (the rebuild) against the
legacy `EDIUpload.pas:186-305` (997/824 branches) + `UPDATE_EINStatus` (`/tmp/inv_utf8.sql:1711-1730`;
`CreateInventory.sql:1711-1730`) + `AD_GetSiteTMMDUNS` (live `VehicleOrder`).

**Evidence base (this session):** live `mssql-spike` — `Inventory` (rebuild, rolled-back probes),
read-only `Inventory_Live` (parity snapshot), read-only `VehicleOrder` (legacy ALC DUNS proc + Site DUNS
data). Pure parsers exercised in CPython; the driver exercised via `scripts/e2e/jython_shim.py`
(`Inventory_Spike` → physical `Inventory`). **Spike restored as-found** (all sentinel-row counts = 0;
all destructive probes were `ROLLBACK`-ed or sentinel-cleaned).

---

## KEY FINDING #2 — the DUNS-guard column (VC_DUNS vs VC_TMM_DUNS)

### Result: the rebuild's `VC_TMM_DUNS` choice is FAITHFUL TO THE LEGACY PROC — but BOTH match a column that, on OUR OWN outbound ISA layout, does NOT hold what `delSL[4]` carries. The element question is genuinely UNPROVABLE without a captured inbound ISA.

**What the legacy does (proven, live `VehicleOrder`):** `AD_GetSiteTMMDUNS` body:

```sql
CREATE PROCEDURE [dbo].[AD_GetSiteTMMDUNS] @SiteTMMDUNS varchar(10)
AS  SELECT * FROM Site s WHERE s.SiteTMMDUNS = @SiteTMMDUNS;
```

`EDIUpload.pas:71-79` splits the ISA on `*` and passes `delSL[4]` → `@SiteTMMDUNS`, matched against
**`Site.SiteTMMDUNS`** (the Toyota/TEMA DUNS), NOT `Site.SiteDUNS` (our own DUNS). The rebuild
(`code.py:265-280`, `_resolve_site_by_duns`) matches `delSL[4]` against **`INV_SITES.VC_TMM_DUNS`**.
**Same column semantics → the rebuild is faithful to the legacy proc.** Verdict on the parity question
narrowly asked ("VC_DUNS vs VC_TMM_DUNS"): **VC_TMM_DUNS is the correct, legacy-faithful choice.**

**BUT the load-bearing risk is the ELEMENT `delSL[4]` lands on, and it is unproven — proven against the live data:**

Live DUNS values (read-only `VehicleOrder.dbo.Site`):
- `SiteDUNS` (ours) = `969009112`
- `SiteTMMDUNS` (Toyota) = `808369495` / `961659588`

Our OUTBOUND ISA builder (`EDI856Object.pas:145-153`, `EDI810Object.pas:140-148`) writes, element by
element (verified `legacy_splitString` == `str.split('*')` for indices 0-8):

| split idx | ISA element | outbound value (856/810 builder) |
|-----------|-------------|----------------------------------|
| `delSL[4]` | ISA04 security-information | **`SiteDUNS` (our `969009112`)** ← the index the guard reads |
| `delSL[6]` | ISA06 sender ID | `SiteDUNS-SupplierCode` (ours) |
| `delSL[8]` | ISA08 receiver ID | **`SiteTMMDUNS` (Toyota `808369495`)** |

So on a file WE generate, `delSL[4]` = OUR DUNS, and `SiteTMMDUNS` is at `delSL[8]`. The legacy guard
matches `delSL[4]` against the **TMM** column — i.e. it matches the ISA04-security slot against the
Toyota DUNS. For that to ever match an inbound file, TEMA's inbound ISA must place its OWN (Toyota) DUNS
at the ISA04-security slot — a **non-standard** layout (ISA04 is "Security Information," normally blank
/ zeros; the routing parties are ISA05–08).

**Counterexample (rebuild `_split_isa` + `_resolve_site_by_duns` against the three plausible inbound conventions, real DUNS):**

| inbound ISA convention | `delSL[4]` | matches `VC_TMM_DUNS=808369495`? |
|------------------------|-----------|----------------------------------|
| X12-standard (ISA04 = security 00s; parties at ISA05-08) | `''` | **NO → QUARANTINE (every TEMA file mis-routed)** |
| Mirror-of-our-quirk (TEMA copies sender DUNS into ISA04) | `808369495` | YES → routes |
| Our-own-outbound shape (ISA04 = sender's own DUNS) | `969009112`(if echoed) | **NO → QUARANTINE** |

Only ONE of three plausible conventions routes correctly, and it is itself non-standard. **The same risk
exists in the legacy** (identical column + index), so this is not a rebuild *regression* — but it means
"the rebuild reproduces the legacy" is true while "the routing is correct" is **UNPROVABLE without a
captured inbound TEMA ISA.** If the legacy ever actually consumed inbound files successfully, TEMA must
use the mirror-of-our-quirk layout and the rebuild inherits that correctly; if the legacy inbound path
was never exercised in anger (plausible — see #5, the 6 stuck-`S` rows suggesting 997s often did not land),
the rebuild faithfully carries a latent mis-route.

**Classification:** code is faithful-to-legacy (NOT a defect vs legacy); the routing CORRECTNESS is an
**unfixable data-vintage / golden-sample gap** (need a captured inbound ISA to pin the element).

**Parity-METHOD flaw piggy-backing here (BLOCKER for the test, not the code):** both test suites build the
inbound ISA with `build_inbound_isa(tmmDuns)` / `inbound_isa(tmmDuns)` that **hard-code the TMM DUNS into
the `delSL[4]` slot** (`test_edi_inbound_build.py:86-92`, `test_edi_inbound_e2e.py:84-87`). The DUNS-guard
assertions (`build:246-247`, `e2e:158-160, 273-276`) then verify the guard matches that same fixture. This
is **self-consistent, not equivalent**: the fixture is built to the rebuild's own interpretation of
`delSL[4]`, so the green DUNS tests prove nothing about which element a real inbound ISA uses. The 42/42
and 29/29 green do **not** retire the element question.

---

## BLOCKER / SHOULD-FIX / NIT findings

### BLOCKER-1 (parity-method) — the DUNS-element is tested against a fixture built to the answer
See KEY FINDING #2. The test fixtures encode the rebuild's `delSL[4]=TMM-DUNS` assumption and then assert
the guard agrees. `test_edi_inbound_build.py:86-92,246-247`; `test_edi_inbound_e2e.py:84-87,158-160`.
**Flaw type:** test/parity-method (vacuous on the load-bearing unknown). The code itself is legacy-faithful;
the TEST overclaims. Gate: do not treat the DUNS tests as proof until a captured inbound ISA fixes the element.

### SHOULD-FIX-1 (code defect) — idempotency guard is defeated by a rename; 824 then DOUBLE-flips + writes DUPLICATE alarms
`_hash_text` docstring (`code.py:306-309`) claims "the same file content re-dropped is detected **even
under a different name**." But `_already_processed` (`code.py:288-291`) keys on
`VC_FILE_NAME = ? AND VC_FILE_HASH = ?`. The name is ANDed in, so a same-content file under a new name is
**not** a ledger hit and is fully re-processed.

**Live counterexample (driver via shim, rolled-back/cleaned on `Inventory`):** an 824 (manifest `ZZRENM01`)
dropped as `zzren_fileA.edi` then re-dropped with identical content as `zzren_fileB.edi`:
- after drop #1: alarm rows for manifest = **1**
- drop #2 outcome = **PROCESSED** (not `SKIPPED_ALREADY_PROCESSED`)
- after drop #2: alarm rows for manifest = **2** (duplicate), and the ASN re-flipped to `R`.

This is realistic — VAN/mailer redelivery commonly re-stamps the filename. For a 997 the re-apply is
idempotent (same status), but for an 824 it produces duplicate main-screen alarms and a redundant `R`
flip. Either dedupe on `VC_FILE_HASH` alone (matching the docstring) or fix the docstring to say the guard
is name-scoped. **Defect type:** code/spec mismatch. (Legacy de-duped implicitly via content-derived
archive names `824<ISA13>.EDI` / `997<EIN>.EDI`, so the rebuild is arguably weaker here than legacy against
a renamed redelivery.)

### SHOULD-FIX-2 (correctness of the NEW Q10 824 behavior) — one 824 reject line flips UP TO 3 ASNs (manifest fan-out), untested
The 824 flip (`code.py:451-455`) is `UPDATE a SET a.VC_ASN_STATUS='R' FROM INV_ASN_MST a JOIN
INV_ASN_DETAIL_MST d ON a.IN_ASN_ID=d.IN_ASN_ID WHERE d.VC_MANIFEST_NUMBER = ?` — it flips **every** ASN
whose detail carries that manifest.

**Proven on Live (read-only):** a manifest maps to up to **3 distinct `IN_ASN_ID`** —
`MAX(distinct ASN per manifest) = 3`; e.g. manifest `52066074` → 7 detail rows across **3** ASN IDs.

**Rolled-back counterexample on `Inventory`:** two distinct ASNs (`4934`, `4935`) sharing manifest
`ZZFAN001`; the rebuild's single-NTE UPDATE flipped **both** to `R`, immediate `@@ROWCOUNT = 2`. So one
824 reject line records `asnsFlagged = 2` (or 3 on real data) and flips multiple shipments.

The e2e test only ever asserts the single-ASN case (`asnsFlagged == 1`, `e2e:246-247`), so the fan-out is
**unexercised**. Whether flipping all N is *correct* is a Q10 intent question (a manifest legitimately
spanning multiple ASNs may warrant flipping all — but it may also over-reject). **Not a legacy divergence**
(legacy flipped NOTHING — `EDIUpload.pas:253-305` is Excel-only; confirmed). **Type:** correctness gap in
NEW behavior + a test-coverage hole. Decide + test the multi-ASN policy.

### SHOULD-FIX-3 (incomplete fix) — AK9 E/P stored distinctly but render BLANK; the "required follow-on" is NOT on disk
`ak9_to_status` (`code.py:76-102`) correctly maps `A/E/P/R` distinct + `M/W/X/garbage/empty → R` (Q6, a
strict superset of the legacy binary A-vs-rest). Verified in the pure test (A→A, E→E, P→P, R→R, M/W/X/Z/''
/None→R, case-insensitive). **But** the only status-render decode that exists anywhere decodes **A/S/C/R
only** — legacy `SELECT_ASNStatus` (`/tmp/inv_utf8.sql:1955-1959`, `:1979-1983`) and invoice
`:3254-3257`. A repo-wide search found **no rebuild status-render NQ/view** that decodes `E`/`P`
(only the spec, the analysis, and the `code.py` docstrings mention them). So a stored `E`/`P` renders
**BLANK** downstream — exactly the legacy bug Q6 set out to fix. `code.py:90-93` itself flags this as a
"REQUIRED M1 follow-on … recorded as a REQUIRED M1 follow-on, not silently dropped." **Until that
render arm exists, the E/P fix is half-done.** Confirm against a golden whether TEMA ever sends E/P to
this supplier (if only A/R, the gap is latent). **Type:** acknowledged incomplete fix / pending golden.

### NIT-1 (latent crash) — `siteId=None` interpolated into the `-- M4` comment raises TypeError before the DB call
`_apply_997_ack` (`code.py:329,335`) and the 824 flip (`code.py:454`) build the M4-marker comment with
`% (siteId,)` where the format is `%d`. `("...%d..." % (None,))` → `TypeError: %d format: a real number is
required, not NoneType`. The per-file path guards this (`process_one_file` quarantines on `siteId is None`
before dispatch), so it is unreachable today. But `process_997`/`process_824` called with `tx` set and
`siteId=None` (a future direct caller) would crash inside the try and roll back. Defensive only.
**Type:** latent code smell. Cosmetic until a caller violates the contract.

---

## Items CONFIRMED FAITHFUL (attacked, did not break)

- **#1 EIN extraction byte-faithful.** Legacy `copy(fcl,5,2)` chars 5-6 → `ak1[4:6]`; `copy(fcl,8,9)`
  chars 8-16 → `ak1[7:16]` (`code.py:158-159`). Verified on `AK1*SH*123456789` (EIN 123456789), a short
  `AK1*SH*42` (EIN 42, same as legacy `StrToInt`), and a 10-digit control # (both truncate to 9). A 3-char
  fnId (`AK1*SHX*...`) makes the EIN slice start mid-junk → `EdiInboundError`; legacy `StrToInt` would
  also fail there. Match.
- **#1 AK9 tolerance is a deliberate, correct fix.** Rebuild scans forward to the real `AK9`, tolerating
  `AK2/AK3/AK4` (`code.py:166-185`); legacy reads char 5 of the immediately-following line blindly
  (`EDIUpload.pas:196-197`). Rebuild is MORE correct; for the simple `AK1`+`AK9` case both agree. (Byte
  parity to a REAL TEMA 997 remains golden-pending — not claimed.)
- **#3 site-scoping marker + @@ROWCOUNT alarm.** Legacy `UPDATE_EINStatus` has NO site column and discards
  `@@ROWCOUNT` (confirmed: `INV_ASN_MST`/`INV_INV_MST` carry no `IN_SITE_ID`; proc body at
  `/tmp/inv_utf8.sql:1711-1730`). Rebuild threads `siteId` into a `-- M4` marker and CHECKS rows: an
  unknown EIN (0 rows) writes a `997_UNKNOWN_EIN` alarm, not a silent success (`code.py:367-377`).
  E2E-verified: unknown EIN `889999` → exactly 1 alarm row, `BIT_RESOLVED=0`.
- **#4 manifest join is structurally correct.** Manifest lives on `INV_ASN_DETAIL_MST.VC_MANIFEST_NUMBER`
  (verified — `INV_ASN_MST` has NO manifest column); the join `a.IN_ASN_ID = d.IN_ASN_ID` → status `R` is
  right (see SHOULD-FIX-2 for the fan-out scope question). 824 writes one alarm row per reject line with
  manifest/part/errorText (`code.py:465-469`), e2e-verified (2 lines → 2 alarm rows, fields verbatim).
- **#4 SE-trailer termination is a deliberate fix.** Rebuild stops at the first real `SE` (`code.py:221`);
  legacy `while data <> 'SE*'` compared a 3-char `copy(fcl,1,3)` to the 4-char literal `'SE*'` — never
  equal — so it ran to EOF. For a single-ST 824 they agree; if NTE-shaped lines follow the SE (a 2nd ST
  loop), legacy over-collects and the rebuild correctly stops. Documented (`code.py:203-206`).
- **#6 997 re-apply idempotent for SAME content+name.** E2E-verified: re-drop of the identical file
  (same name+hash) → `SKIPPED_ALREADY_PROCESSED`, no re-flip, no duplicate alarm, one ledger row. (The
  rename hole is SHOULD-FIX-1.)
- **#6 re-ack terminal-state.** Rebuild's flip is an unconditional `UPDATE … SET status=?` (`code.py:327-336`),
  matching legacy's unconditional `UPDATE_EINStatus` — A→R and R→A overwrites are both allowed, faithful.
  (Spec §4.5 leaves terminal-state locking as an open decision; rebuild = legacy = no lock.)
- **#7 status domain.** Rebuild writes A/E/P/R (Q6); legacy wrote the raw AK char into a `varchar(1)`.
  Live snapshot is all `A` (no C/S/R/E/P), so C/S/R are exercised only via synthetic fixtures — the
  reject/create branches are **unobservable on parity data** (a data-vintage limit, stated honestly in the
  e2e docstring). SH→ASN / IN(else)→invoice routing matches `UPDATE_EINStatus`'s `if @EINType='SH' … else`.

---

## VERDICT

**The inbound processor is NOT proven equivalent — but the gaps are (a) golden-sample/data-vintage limits
that are equally true of the legacy, plus (b) a small number of rebuild-side fixables. It is faithful to
the legacy on every byte-parsing and routing decision I could test.**

Specifically:

1. **DUNS column (#2):** `VC_TMM_DUNS` is the LEGACY-FAITHFUL choice (legacy `AD_GetSiteTMMDUNS` matches
   `delSL[4]` against `Site.SiteTMMDUNS`). **Not a defect vs legacy.** But which X12 element `delSL[4]`
   is on an inbound TEMA ISA is **UNPROVABLE from available data** — and on our own outbound layout
   `delSL[4]` is OUR `SiteDUNS`, with the TMM DUNS at `delSL[8]`. If TEMA's inbound ISA is X12-standard,
   BOTH legacy and rebuild quarantine every file. This is the single biggest load-bearing unknown and is
   blocked on a captured inbound 997/824 ISA.

2. **Equivalence is UNPROVABLE in three places by data vintage:** the real inbound AK1/NTE/ISA byte
   offsets (no golden TEMA file), the C/S/R/E/P status branches (Live is 100% `A`), and the DUNS element.
   The green test suites (42/42, 29/29) are **self-consistent against fixtures built to the rebuild's own
   assumptions** — they do NOT close these gaps, and the DUNS test in particular is vacuous on the
   element question (BLOCKER-1).

3. **Genuine rebuild-side items to fix before prod:** SHOULD-FIX-1 (rename defeats idempotency → duplicate
   824 alarms + double `R` flip — live-proven), SHOULD-FIX-2 (one 824 line flips up to 3 ASNs, untested —
   live-proven, a NEW-behavior correctness/coverage gap), SHOULD-FIX-3 (E/P stored but no render arm
   exists → renders blank, the Q6 fix is half-done).

**Bottom line:** the decided fixes (AK9 A/E/P/R distinct, 824→R+alarm, site-marker, @@ROWCOUNT-checked,
idempotent, no-bug-carry) are PRESENT and behave as specified on the cases that are observable — but
"correct" cannot be asserted over the legacy until (i) a captured inbound TEMA 997/824 pins the DUNS
element + the AK1/NTE offsets, and (ii) the idempotency rename hole, the 824 fan-out policy, and the E/P
render arm are closed. Until then: **faithful-to-legacy where testable; equivalence UNPROVABLE on the
golden-pending + DUNS-element gaps; three fixable rebuild divergences proven with counterexamples.**
