# ASN-Creation Keystone — Design / Documentation Review (completeness + honesty)

**Reviewer stance:** adversarial. Default verdict is "not a sign-off basis until it survives." This is the
design/doc angle; the sql-adversary owns empirical counterexamples on the live DB.

**Inputs reviewed (all read in full):**
- `docs/analysis/production-readiness/sql/{AD_FRSPULL-analysis, SELECT_ForecastDetailBCASN-analysis,
  asn-write-chain-analysis, delphi-fanout-confirmation, equivalence-map}.md`
- `docs/analysis/production-readiness/m1-asn-creation-spec.md`
- `docs/analysis/edi/project-library/asn/code.py`
- `docs/analysis/edi/spike-asndetail-rekey.sql`, `docs/analysis/production-readiness/AD_FRSPULL-shared.sql`
- `/Users/apple/Documents/FP docs/SQL/VehicleOrder.sql` (the live ALC dump, UTF-16, Script Date 06/10/2026)
- `scripts/e2e/test_create_asn_parity.py`

**Headline:** the doc set is *unusually honest* about its parity ceiling (the parity test prose is exemplary —
it refuses to sell self-consistency as parity). But it is **NOT trustworthy as a production correctness sign-off
yet**, for one reason that dwarfs the rest: the foundational source artifact for the qty inputs
(`AD_FRSPULL-shared.sql`) **materially disagrees with the live `VehicleOrder.sql` dump on three load-bearing
facts**, the equivalence map propagated one of those wrong facts (char(3)) into the rebuild code, and the two
internal analyses already contradict each other on it without resolution. The fan-out qty is computed off a BC
formula nobody in this doc set has reconciled to a single authoritative source.

---

## BLOCKERS (plant-floor / revenue: wrong BC -> wrong manifest -> mis-ship / mis-bill)

### B1. The AD_FRSPULL source artifact contradicts the live dump on THREE load-bearing facts; the analysis's "no SQL drift" claim is false.

This is the worst finding and it sits under everything else, because `VEHICLES`/`ORDERS`/`BC` from this proc
drive every detail qty and every manifest match.

`AD_FRSPULL-analysis.md §0` asserts: *"`OBJECT_DEFINITION(...)` on the live VehicleOrder is **byte-identical in
its SQL** to the decoded body in `AD_FRSPULL-shared.sql` ... **No behavioral drift.**"* That is refuted by the
live dump on disk:

| Fact | `AD_FRSPULL-shared.sql` (analysis basis) | Live `VehicleOrder.sql:4722/4742/4745/4765` | Consequence |
|---|---|---|---|
| BC width | `convert(char(3), …)` (lines 46,66,69,85) | `convert(char(21), …)` (lines 4722,4742,4745,4765) | The entire char(3)/trailing-space "namespace separator" theory (AD_FRSPULL-analysis §3.4/§7.2, equivalence-map #2/#5) is built on a width that the live proc does not use. |
| Ground concat order | `vd2+vd1` with vd2=GROUNDWHEEL, vd1=GROUNDTIRE = **WHEEL+TIRE** | `vd1+vd2` with vd1=GROUNDTIRE (i1, line 4733), vd2=GROUNDWHEEL (i2, line 4738) = **TIRE+WHEEL** | The two artifacts produce a **different BC string** for the same vehicle. AD_FRSPULL-analysis §2/§7.1 insists "WHEEL before TIRE"; the live dump is TIRE before WHEEL. A reversed BC matches the wrong `VC_BROADCAST_CODE` pattern or none -> wrong recipe or "Missing Broadcast Code" abort. |
| Spare `<> 'M'` filter | **present** (line 84) | **ABSENT** — 0 occurrences of `<> 'M'` in the live proc range (verified by grep over 4714-4767) | AD_FRSPULL-analysis §3.3/§7.5 declares the `<> 'M'` exclusion "real and load-bearing ... The rebuild MUST keep `spare value <> 'M'`." The live dump does not have it. If the rebuild keeps it, the rebuild **drops spare rows the live proc keeps**; if it drops it, it diverges from the analysis. Either way the analysis's MUST is wrong against this dump. |

Worse, the doc set is **internally** inconsistent and never closes it: `delphi-fanout-confirmation §f` reads
`char(21)` from "the live `VehicleOrder.sql`" (matching the dump), while `AD_FRSPULL-analysis §0` swears the
live proc is `char(3)`. The equivalence-map (#5) *names* this contradiction ("char(3) vs char(21) ... must be
adjudicated against the running proc") — good — but then it leaves it open AND ships `code.py:278` with the
comment `# AD_FRSPULL BC is char(3)` and a `bc.strip()` justified by the char(3) theory. **The rebuild code
encodes the unverified side of an unresolved contradiction.**

There are effectively **three** AD_FRSPULL variants implied by this doc set: (a) `shared.sql` (char(3),
WHEEL+TIRE, has `<> 'M'`); (b) the live dump (char(21), TIRE+WHEEL, no `<> 'M'`); (c) the running mssql-spike
backup the sql-adversary tested, which produced `[NN ]` 2-char-plus-space spare BCs and a firing `<> 'M'`
exclusion — i.e. it behaves like (a), NOT like (b). The doc set treats "the live proc" as a single settled
artifact when at least two and probably three live in the wild.

> Direction of fix (delphi-architect + sql-adversary, jointly): pull `OBJECT_DEFINITION('dbo.AD_FRSPULL')`
> from the **one DB the rebuild will actually read at runtime**, declare it canonical, and reconcile shared.sql,
> AD_FRSPULL-analysis §0/§2/§3.3, delphi-fanout §f, equivalence-map #2/#5, and `code.py:278` to it in one pass.
> Until then no BC-dependent parity claim is admissible.

### B2. `bc.strip()` (code.py:278) is a correctness change, not a fidelity choice — and it is justified by the refuted char(3) theory.

Even setting B1 aside: T-SQL `LIKE` does **not** trailing-trim its **left** operand
(SELECT_ForecastDetailBCASN-analysis §1b is explicit; AD_FRSPULL-analysis §6 re-states it). The legacy feeds
the BC as `@BCode varchar(20)` with whatever padding `TField.AsString` produced from a fixed-width CHAR column —
the doc set itself flags (AD_FRSPULL-analysis §6 closing line; delphi-fanout residual #3) that **whether Delphi
trims is unconfirmed**. `code.py:278` unconditionally strips. If the legacy did NOT trim, the rebuild changes
what every spare BC matches under `LIKE`. The equivalence map correctly rates this ❌ (#2/#5) — but it is filed
as a "GAP to surface," when its blast radius (silent wrong-recipe match on spare BCs) makes it a BLOCKER for
sign-off, not a tracked risk. The map's own ordering ("by how directly they threaten qty parity") should have
put #2/#5 above the doc-only items; it does, but the severity tag (❌ GAP) under-weights it relative to the
⚠️ items that got their own "sign-off needed" section.

> Direction: resolve via B1 first (which width/padding the canonical proc emits + a one-line Delphi read of
> whether `AsString` trims), then make `code.py` reproduce that exactly instead of unconditionally stripping.
> delphi-architect owns the `AsString`-trims question.

### B3. Nondeterministic recipe order is inherited verbatim into a revenue path, and the prescribed fix was not applied.

`SELECT_ForecastDetailBCASN` has no `ORDER BY` over two heaps
(SELECT_ForecastDetailBCASN-analysis §3, proven `fc_indexes=0 fc_heap=1`). The No-Ratio branch takes
`fcRows[0]` (code.py:150). `computeAsnDetails` appends rows in proc-returned (heap-scan) order with no re-sort
(verified: grep for sort/order/sorted in code.py returns NONE). The analysis prescribes
`ORDER BY ID_FORECAST_DETAIL` + a David decision on which assy the single-vehicle case should pick; the
equivalence map (#6) honestly records that **neither was done**. For the multi-row No-Ratio BCs the analysis
enumerates ([KLM]CC, [MNP]BB, etc. — up to 3 candidate assys with different part numbers and ratios) the
shipped part is allocation-order-dependent. On a table reload, page reuse after deletes, stats change, or a
parallel plan, the rebuilt ASN can pick a *different part number* than last run — and the parity test
(self-consistency) would still pass because it recomputes over the same unordered read. This is a silent
wrong-part hazard on the morning revenue path.

> Direction: delphi-architect/developer add the deterministic ORDER BY and get David's call on the
> single-vehicle assy; this is a must-fix, not a "surface it."

---

## SHOULD-FIX (real divergences needing David sign-off, or gaps under-weighted)

### S1. EIN-at-send vs EIN-at-create — sign-off REQUIRED, and one downstream consumer is unverified.

The map (#10/#11) and spec §2/§8 correctly flag this as an intended divergence and correctly identify the two
legacy bugs it removes (out-of-txn ALC `SiteEIN+1` not reverted on rollback; unscoped `UPDATE Site SET
SiteEIN+1` with no WHERE = multi-site BLOCKER, confirmed in delphi-fanout §e against `vo_utf8.sql:623`). That
reasoning is sound. **But the doc set never checks whether any consumer reads `IN_ASN_EIN` between create and
send.** The legacy stamps `fEIN+1` on header AND every detail at create (delphi-fanout §e; asn-write-chain §1);
the rebuild writes 0 to both until send. asn-write-chain §1 asserts "nothing in these four procs interprets
`IN_ASN_EIN`" — true for *these four procs*, but that is a narrower claim than "no consumer assumes an ASN has
an EIN at create." A report, a re-print, an 856 regeneration, or a between-create-and-send query that filters
or displays `IN_ASN_EIN` would see 0. **This must be a David sign-off** (he knows the floor workflow), backed by
a grep of every `IN_ASN_EIN` reader, before it ships. The map's "verify the parity harness ignores
`IN_ASN_EIN`" is necessary but not sufficient — it protects the test, not production.

### S2. Single-transaction (vs legacy no-txn partial ASN) — sign-off REQUIRED that no workflow depends on partial ASNs.

The map (#16) and asn-write-chain §0/§6.9 are honest that this changes failure semantics from "partial ASN
persists" to all-or-nothing, and frame it as an improvement. Almost certainly correct. The unproven half: does
any legacy recovery workflow **depend** on a half-written ASN surviving a mid-loop abort (e.g. an operator who
re-runs and relies on the earlier-BC rows already being there, or a cleanup proc that expects orphan headers)?
The doc set asserts the improvement without checking for a dependency on the bug. Low probability given the
single-user desktop, but it is a behavior change on the revenue keystone and belongs on the sign-off list, not
in the "improvements, assumed safe" bucket.

### S3. Concurrency: no DB unique constraint behind the idempotency guard — correctly flagged, severity right, but it is a NEW exposure the rebuild creates.

The map (#14/#16) and asn-write-chain §3/§7 correctly note the read-back guard (`SELECT_ASNSeq` then skip,
code.py:263-267) has no DB unique constraint and that two concurrent `create_asn` calls can both pass and
double-insert. This is genuinely new: the legacy was single-user desktop; the gateway is multi-session. The
severity (concurrency double-ASN = duplicate shipment) is BLOCKER-class in production even though it is
"low-probability today." It is filed as a ❌ GAP, which is right, but it should be explicitly on the
**must-fix-before-sign-off** list (add unique index on `(VC_LINE_NAME, VC_PRODUCTION_DATE[, site_id])` with
`VC_START_SEQ_NUMBER <> -1` semantics, or serialize), not deferred — because the rebuild is what introduces the
concurrency, so "the legacy never hit it" is not a defense.

### S4. The EIN-at-send relocation correctness (`AD_GetSite`/`AD_UpdateEIN` -> `INV_SITES.IN_EIN_SEQ`) is asserted but not specified or tested.

code.py docstring (:249-254) says the real EIN is "allocated atomically from `INV_SITES.IN_EIN_SEQ` ... at SEND
(M1 Rank 2)." None of the five analyses or the spec contains `INV_SITES` or `IN_EIN_SEQ` — they describe the
legacy ALC `Site.SiteEIN`. So the target of the relocation (the new per-site sequence column/table) is named
only in a docstring and is **out of scope of every analysis here and untested by the parity harness** (which
asserts EIN=0 at create and stops). The map (#11) treats the at-send allocation as settled; it is not — it is a
forward reference to unbuilt M1-Rank-2 work. Honest framing should say: "create-side EIN=0 is verified; the
send-side per-site atomic allocation is a future deliverable, not yet designed or tested here."

### S5. Multi-site `site_id` deferral threads through writes that are not yet site-scoped — a real partial-state hazard.

`create_asn(site=1, ...)` carries `site` but "does not touch SQL today" (code.py:239-241). The idempotency
guard (`SELECT_ASNSeq`), the re-keyed detail upsert, and `DELETE_ASNItem` are all described as re-keyed to
`(site_id, ...)` "at M4" (spec §1/§5, map #12/#14). Until M4, the rebuild is single-site-correct but the doc
set repeatedly writes "re-keyed `(site_id, IN_ASN_ID, manifest)`" as if done (equivalence-map #12 status ✅
"re-key applied") when the site_id half is **not** in the spike proc (spike-asndetail-rekey keys on
`(IN_ASN_ID, VC_MANIFEST_NUMBER)` only; asn-write-chain §2b confirms two-column). A reader skimming the ✅ could
conclude multi-site dedup is done. It is two-thirds done. Tag it ✅-single-site / ⬜-site_id-deferred so the ✅
does not over-claim.

### S6. Cross-DB AD_FRSPULL read failure mode mid-create is undocumented.

`create_asn` does the AD_FRSPULL read (step 2, code.py:272) and the per-BC SELECT_ForecastDetailBCASN reads
(step 3) **before** opening the Inventory transaction, then opens the txn for writes. Good ordering. But the
failure mode if the ALC datasource is down or the cross-gateway connection drops between the read and the write
is never discussed. Legacy ran both on a live ADO connection in the same click; the rebuild splits Inventory
and VehicleOrder into separate Ignition datasources (`DATABASE` vs `ALC_DATABASE`, code.py:195-196). A
VehicleOrder timeout yields an empty/partial `frsRows` -> a silently smaller ASN (fewer BCs) with no abort,
because nothing asserts AD_FRSPULL returned the expected vehicle count. The window-is-the-filter cross-check
exists only in the *test* (test_create_asn_parity.py:157-163: spare BC sum == header qty), **not in the
driver.** A short-read from a flaky ALC connection ships an under-counted ASN. Worth a guard + a finding.

### S7. NOLOCK dirty-read in AD_FRSPULL (live dump) is not carried into the analysis's fidelity list.

The live dump uses `with(NOLOCK)` on every table in AD_FRSPULL (VehicleOrder.sql:4724-4754).
`AD_FRSPULL-shared.sql` *also* has NOLOCK (lines 48-78), but **neither the AD_FRSPULL-analysis nor the
equivalence map mentions NOLOCK at all.** Dirty reads on the GALC `Vehicle` heap mid-broadcast can count an
in-flight (uncommitted) vehicle or miss one, perturbing `VEHICLES` and thus every downstream qty. The analysis
proved the one-row-per-vehicle invariant (§3.7) but did not flag that NOLOCK can read a partially-written
re-broadcast. Low frequency, but it is a legacy behavior the analyses claim to enumerate and they missed it.

---

## NITS (doc hygiene; correct but mis-stated or stale)

- **N1. Stale "14-char" stamp comments.** asn-write-chain §1 proves the stamp is **16 chars** (style-114, first
  2 ms digits); `spike-asndetail-rekey.sql:25-26` and `m1-asn-creation-spec.md §2` still say "14-char". The map
  (#17) catches this and labels it doc-only. Correct, low severity, but fix the comments so the parity reviewer
  is not misled (cosmetic — nothing parses the stamp back).
- **N2. The spec's central caveat is dead but still prominent.** `m1-asn-creation-spec.md` §3/§9 and the boxed
  "single biggest correction" still declare AD_FRSPULL "the one true source gap / BLOCKER for parity, body NOT
  in any dump." `delphi-fanout-confirmation` (newer, same author) refutes this ("**WRONG** ... fully present at
  `/tmp/vo_utf8.sql:4714`") and I confirmed the dump is on disk. The spec was never updated. A reader hitting
  the spec first sees a false blocker; a reader hitting the confirmation first sees it resolved. Reconcile —
  the spec is the doc David will read for the chain, and its headline is now false. (This is the *inverse*
  source-gap bug from the review checklist: something declared "missing" that EXISTS in the repo. It is fixed in
  one doc and stale in another.)
- **N3. Equivalence-map sources line says "all complete on disk" (line 7)** — true for the four analysis docs,
  but the AD_FRSPULL *body* those rest on is the contested artifact (B1). The "complete on disk" framing
  understates the open contradiction.

---

## Answers to the five charge questions

1. **Completeness of the map.** Mostly complete and the 17-row table maps every legacy behavior to a rebuild
   location. Gaps in coverage: NOLOCK (S7) is absent; the cross-DB read failure mode (S6) is absent; the
   at-send relocation target `INV_SITES.IN_EIN_SEQ` (S4) is named only in a docstring and unmapped. The
   flagged-but-under-weighted items are correctly listed but mis-ranked: `bc.strip()`/char-width (#2/#5) and
   nondeterministic order (#6) are BLOCKER-class, not "GAPs to surface." No ✅ is *fabricated* — but #1 ("✅
   faithful: wrap-the-proc, BC formula never re-derived") rests on a proc whose canonical body is unresolved
   (B1), so its ✅ is contingent, not earned.

2. **⚠️ intended divergences — sign-off vs safe-by-construction.**
   - **David sign-off REQUIRED:** EIN-at-send (S1 — needs a grep of all `IN_ASN_EIN` readers + his workflow
     call); single-transaction all-or-nothing (S2 — needs confirmation no recovery flow depends on partial
     ASNs).
   - **Safe-by-construction (still note, but no sign-off):** headless idempotency no-op vs UI form-lock (#3 —
     same intent, no UI; reasonable) — *conditional* on S3 (the concurrency constraint) landing, because the
     no-op is only safe if a second concurrent create can't slip past it.
   - **Doc-only (no behavior sign-off):** 16-char-vs-14-char comment (N1).
   - **Not actually settled, mislabeled as a decided divergence:** positional VALUES retained for Q1 (#13) is
     fine now but the M4 hazard is real; the at-send EIN *mechanism* (S4) is unbuilt, not a divergence yet.

3. **❌ GAP severities + missing ones.** char3/char21 (#5) and `bc.strip()` (#2): under-rated — BLOCKER (B1/B2).
   Nondeterministic order (#6): under-rated — BLOCKER (B3). Truncation alarm absent (#3): severity right
   (latent, low live-probability). No DB unique constraint (#14): right that it's a gap, but should be on the
   must-fix list because the rebuild *creates* the concurrency (S3). Positional-INSERT M4 (#13): right
   (tracked, not today). Stale 14-char (#17): right (nit). **Missing gaps:** EIN-at-send allocation race at the
   856-send step (S4 — the new atomic sequence is undesigned/untested here); cross-DB AD_FRSPULL short-read
   mid-create (S6); multi-site site_id half-done but shown ✅ (S5); NOLOCK dirty read (S7); the inverse
   source-gap staleness (N2).

4. **Does the documentation oversell correctness?** The **parity test prose does not** — it is a model of
   honesty (test_create_asn_parity.py:6-64 explicitly: self-consistency ≠ legacy parity; names the
   recipe-vintage drift; proves only the total-qty invariant 4240==4240 and per-manifest is *reported, not
   gated*; even leaves an IG-TODO to upgrade to gated parity when a same-vintage ASN exists). That is exactly
   right and should be preserved verbatim. The **equivalence map and AD_FRSPULL-analysis DO oversell** in two
   spots: (a) "✅ faithful" rows that depend on the unresolved canonical proc (B1); (b) AD_FRSPULL-analysis §0's
   flat "byte-identical / no SQL drift," which is false against the dump. A reader who reads the equivalence map
   alone — which is its stated purpose ("Feeds the Verify gate and David's sign-off") — could mistake this for
   "ASN creation is verified correct." It is not; what is proven is total-qty conservation + driver
   self-consistency under one contested set of inputs.

5. **Can this serve as a correctness sign-off?** **No — not yet.** It is an excellent *analysis* and a candid
   *parity-ceiling statement*, but a sign-off basis requires (a) one canonical AD_FRSPULL adjudicated and
   propagated (B1), (b) the BC-padding/`strip` reconciled to it (B2), (c) deterministic recipe order + David's
   assy call (B3), (d) the concurrency constraint (S3), and (e) a true row-parity oracle — which the doc set
   itself admits does not exist today (no reproducible legacy ASN shares the current recipe vintage). Without an
   oracle, "correct" is unprovable by construction; the most this set can sign off today is "qty-total-conserving
   and internally self-consistent," which is not the same as "ships the right manifests."

---

## VERDICT

**Is the doc set TRUSTWORTHY as a production correctness sign-off basis? NO.**

It is trustworthy as a *behavioral analysis* and as an *honest parity-ceiling statement* (the parity test is
exemplary and must be kept as-is). It is **not** a correctness sign-off because the qty/BC inputs rest on an
unresolved three-way contradiction about the one proc that produces them, the rebuild code encodes the
unverified side of that contradiction, a known nondeterminism is carried into the revenue path unfixed, and —
by the doc set's own admission — there is no legacy row-parity oracle. Self-consistency + total-qty conservation
is what is actually proven; the docs should not let a Verify-gate reader read it as "ASN creation is correct."

### David sign-offs needed (consolidated)
1. **EIN timing** — accept EIN-at-send over legacy EIN-at-create, AFTER a grep confirms no consumer reads
   `IN_ASN_EIN` between create and send (S1).
2. **Failure semantics** — accept single-transaction all-or-nothing over legacy partial-ASN persistence,
   confirming no recovery workflow depends on partial ASNs surviving (S2).
3. **Single-vehicle No-Ratio assy** — which assy the `Orders<=5` "first row" should deterministically pick once
   an ORDER BY is added (B3; SELECT_ForecastDetailBCASN-analysis §3).
4. **Spare `<> 'M'` semantics** — whether the canonical runtime proc includes the `<> 'M'` exclusion (it is in
   shared.sql, absent in the dump), i.e. whether the rebuild keeps it (S1/B1).

### Must-fix BEFORE sign-off (ranked by blast radius)
1. **B1** — Adjudicate ONE canonical `AD_FRSPULL` (width char(3) vs char(21); ground concat TIRE+WHEEL vs
   WHEEL+TIRE; `<> 'M'` present vs absent) against the DB the rebuild will actually read; reconcile shared.sql,
   AD_FRSPULL-analysis §0/§2/§3.3, delphi-fanout §f, equivalence-map #2/#5, and `code.py:278` to it.
2. **B2** — Replace the unconditional `bc.strip()` (code.py:278) with the exact padding/trim the canonical proc
   + the Delphi `TField.AsString` actually produce; confirm spare-BC `LIKE` matches end-to-end.
3. **B3** — Add deterministic `ORDER BY ID_FORECAST_DETAIL` to the recipe read and reflect David's
   single-vehicle assy decision in `computeAsnDetails`.
4. **S3** — Add the DB unique constraint (or serialization) behind the idempotency guard; the gateway is
   multi-session and the rebuild is what introduces the concurrency.
5. **Oracle** — Stand up a true row-parity oracle: either create a legacy ASN under the current forecast-recipe
   vintage and gate `rebuilt == legacy` (the test's own IG-TODO at line 289), or freeze a contemporaneous
   recipe snapshot to replay an old ASN row-for-row. Until one exists, sign-off can only certify
   total-qty-conservation, not manifest-level correctness.

### Lower-priority fixes (after sign-off gate)
S4 (design+test the at-send `INV_SITES.IN_EIN_SEQ` allocation), S5 (tag the site_id half as deferred, not ✅),
S6 (guard the cross-DB AD_FRSPULL short-read), S7 (note NOLOCK dirty-read), N1/N2/N3 (doc staleness: 14-char
comments; the dead "AD_FRSPull is the one source gap" caveat in the spec; "complete on disk" framing).
