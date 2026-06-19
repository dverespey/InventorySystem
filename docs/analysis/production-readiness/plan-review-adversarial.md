# Adversarial Review — Production Implementation Plan

**Reviewer:** adversarial architecture reviewer · **Date:** 2026-06-19
**Target:** `docs/analysis/production-readiness/implementation-plan.md` (M1–M5 + Q1–Q17)
**Method:** every material claim verified against `/tmp/inv_utf8.sql` (UTF-8 of `DB Schema/CreateInventory.sql`,
live dump 2026-06-13), the Delphi `.pas`/`.dfm` source, and the calendar proc body. Default stance skeptical.

Verdict legend: BLOCKER (data-loss / mis-bill / wrong-site write before M1 is safe) > SHOULD-FIX > NIT.
Each finding tagged: **[plan flaw]** / **[decision consequence]** / **[new risk]**.

---

## BLOCKER-1 — `REPORT_EDI856` is NOT a read proc; it self-flips ASN status. Breaks Q2 AND shadow-mode parity. [plan flaw]

**Claim under test:** the plan treats `REPORT_EDI856` as the data feed for the 856 builder (Rank 2, NQ
`edi/report_856`) and treats status-flip as a *separate* step coupled to send-commit (Q2: "don't mark `'S'`
until that ASN's 856 file write + transmit commits").

**Evidence — the proc writes:** `/tmp/inv_utf8.sql:3695`, inside `REPORT_EDI856`'s `@EIN<>0` branch:
```
UPDATE INV_ASN_MST set VC_ASN_STATUS = 'S' WHERE IN_ASN_EIN = @EIN
```
So calling the proc to *generate the 856 data* ALSO flips status to `'S'` as a side effect, keyed by EIN, for
ALL ASN rows with that EIN — before any file is written or transmitted.

**Why it's a BLOCKER:**
1. **Q2 is unimplementable by wrapping this proc.** You cannot "wrap the proc as-is first" (the plan's stated
   principle, line 53–54) and *also* defer the flip to post-send. The wrap-it path flips on read; the Q2 design
   flips on send-commit. They are mutually exclusive. The plan must port the SELECT out of this proc and drop
   the embedded UPDATE — that is logic-porting, not wrapping, and it is not flagged as such anywhere in Rank 2.
2. **Shadow mode (M1) is not read-only against the shared DB.** M1 says "Ignition writes EDI files to a separate
   staging dir... Legacy stays the system of record" (line 270–272). But if Ignition calls `REPORT_EDI856` to
   build its shadow file, it flips real `INV_ASN_MST.VC_ASN_STATUS` to `'S'` in the shared DB — so the legacy
   app then sees those ASNs as already-sent and skips its own 856. **Double-source corruption during the very
   parallel-run the plan relies on as its safety net.** (Failure-mode checklist #4: parallel-run double-write.)

**What settles it:** the Rank 2 design must explicitly state that `REPORT_EDI856` is split into a pure SELECT
(for the NQ) and the status-flip is removed from it. Until then M1 cannot shadow safely.

---

## BLOCKER-2 — `UPDATE_EINStatus` is not site-scoped; per-site EIN sequences (Q4) make 997 acks flip the WRONG site. [decision consequence]

**Claim under test:** Q4 makes EIN "a per-site atomic sequence in the Inventory DB"; Q6/Q7 route inbound 997s
to flip the matching ASN/INV status via `UPDATE_EINStatus`.

**Evidence:** `/tmp/inv_utf8.sql:1722-1729`:
```
if @EINType = 'SH'  UPDATE INV_ASN_MST  SET VC_ASN_STATUS=@EINStatus WHERE IN_ASN_EIN = @EIN
else                UPDATE INV_INV_MST  SET VC_INV_STATUS=@EINStatus WHERE IN_INV_EIN = @EIN
```
The WHERE clause is `IN_ASN_EIN = @EIN` only — **no site predicate.**

**Failure scenario:** Q4 says EIN is now per-site (each site's sequence starts independently). So site A and
site B both legitimately have an ASN with `IN_ASN_EIN = 9069`. A 997 arrives acking EIN 9069 for site A. The
DUNS guard (Q11) confirms the *file* belongs to site A — but `UPDATE_EINStatus(9069,...)` flips **both** sites'
ASN 9069 to Accepted. Site B's ASN gets a phantom acceptance it never received from TEMA. (This is exactly the
InventorySystem shared-`RecordID` pattern, P9, carried forward — checklist #8.)

**Why BLOCKER:** mis-acked ASN/INV = compliance/payment-state corruption on the Toyota exchange. The moment Q4
(per-site EIN) and multi-site coexist, this proc is wrong. The plan parameterizes the hardcoded `6440` (good)
but never adds `site_id`/EIN-namespace scoping to `UPDATE_EINStatus`. Direction of fix → delphi-architect +
ignition-architect: EIN lookup must be `(site_id, EIN)` or EIN must be globally unique across sites (which
contradicts Q4's "per-site sequence"). Pick one; the plan currently implies both.

---

## BLOCKER-3 — M1↔M2 is a circular dependency: the outbound loop cannot be *validated* without M2's inbound. [plan flaw]

**Claim under test:** M1 = outbound ASN→856→810 (parallel-run, gate = "TEMA test-accept of an Ignition-built
856 and 810", line 274). M2 = inbound 997/862/824 (line 276–286).

**The circularity:**
- M1's own acceptance gate (criterion #2, line 256) requires "997-accepted by TEMA." The 997 is an **inbound**
  artifact that only M2 can ingest (`edi_inbound.py`, Rank 3, M2).
- Q2's per-ASN status flip to Accepted (`'A'`) is driven by the 997 (Q6), which is M2.
- Q5's unsend/recreate "no ack received" trigger and the 824-reject auto-flag (Q10) are inbound (M2).

So in M1-only, an Ignition-built ASN can reach `'S'` (sent) but **can never reach `'A'`/`'R'`** — there is no
ack ingester. The M1 ASNs sit in `'S'` indefinitely; the unacked-past-threshold alarm (§4) fires on every one.
You cannot demonstrate M1's "byte-parity + TEMA-accept" gate without standing up at least the 997 path.

**Consequence:** M1 and M2 are not cleanly separable at the boundary the plan draws. Either (a) the 997
ingester moves into M1 (re-scoping M1 upward, worsening the already-tight 5–7wk estimate), or (b) M1's gate is
downgraded to "byte-diff vs legacy only" and TEMA-accept slips to M2 — in which case the plan should say so.
Right now M1's stated gate is unreachable within M1. **[plan flaw — milestone boundary]**

---

## SHOULD-FIX-1 — Q11/Q14 vs D1: "the gateway's configured site" is incoherent under one-gateway-many-sites. [decision consequence]

**Claim under test:** Q11 — the poller "confirms each inbound file's DUNS matches **the gateway's configured
site('s)** DUNS." Q14 — single gateway. D1 — many sites on one shared app/DB.

**The contradiction:** one gateway serves N sites (D1 + Q14). There is no single "the gateway's configured
site." The inbound DUNS must be matched against the DUNS of **every** configured site and the file **routed**
to the matched site (this is what `delSL[4]→site` in Rank 3 actually requires). Q11's singular phrasing
("the gateway's configured site('s)") reads as one-site-per-gateway, which silently contradicts D1.

**Failure scenario:** site B's 830 forecast lands in the shared drop. The poller, checking against "the
gateway's site" (= site A), rejects it as non-matching ("non-matching files are not consumed", Q11). Site B's
forecast is silently dropped; the ≥8-day staleness alarm (Q11 added-scope) eventually fires, but a day of
forecast is lost. **The plan's inbound routing is written for one site even though every other decision is
multi-site.** Direction of fix → ignition-architect: Q11 must read "match DUNS against ALL `sites` rows and
route per match," not "the gateway's site." The Rank 3 `delSL[4]→site (D1)` text is correct; Q11's prose
contradicts it — reconcile.

---

## SHOULD-FIX-2 — Q9 cross-DB calendar read is incompatible with D1 "fully isolated sites" if a site's DB is on another server. [decision consequence]

**Claim under test:** Q9 — calendar `AD_GetSpecialDate` stays in `VehicleOrder`, read cross-DB; D1 — sites are
independent/isolated on one shared DB.

**Evidence the dependency is real and hard-wired:** the Delphi calls it by bare name on the **ALC connection**
(`DataModule.pas:3767` `CommandText := 'dbo.AD_GetSpecialDate'` inside `with ALC_Dataset`), and the proc body
itself uses 3-part names `VehicleOrder.dbo.F_ISO_WEEK_OF_YEAR` (`AD_GetSpecialDate-shared.sql:38,56,75`). So the
read assumes the calendar DB is reachable on the **same SQL Server instance** as the caller.

**Failure scenario:** a future site (D1) whose Inventory DB lives on a *different* SQL Server has no
`VehicleOrder` on its instance — the cross-DB / 3-part read fails. "Shared calendar in VehicleOrder" only works
if every site's Inventory DB co-locates with one canonical VehicleOrder. The plan never says *whose*
VehicleOrder multiple sites read, nor that co-location is a deployment constraint. **This is a latent
single-point-of-failure that quietly defeats D1's isolation claim.** Direction → ignition-architect: state the
co-location constraint explicitly, or replicate the (small, slow-changing) calendar per instance. Acceptable
for the current single-instance reality; must be a documented assumption, not silent.

---

## SHOULD-FIX-3 — The `sites`-table relocation premise is UNVERIFIED in the source; it may not live where the plan says. [plan flaw / needs evidence]

**Claim under test (§4, line 233):** "the authoritative site/line configuration currently lives in the
**VehicleOrder** DB... the order/forecast `LINE` lookups + site config read across databases today."

**What I could NOT confirm:** there is **no reference to a `sites` table anywhere in the Delphi source** — `grep
-rni "FROM sites|dbo\.sites|\.\.sites"` over all `.pas`/`.dfm` returns nothing, and no `dbo.LINE`/`..LINE`
cross-DB read appears either. The site identity in the legacy is read from **INI** (`[SITE]`/`[INIT]`, per
`SiteInfo.pas` / CLAUDE.md), not a `sites` table. The cross-DB things that ARE provable are the ALC procs
(`AD_GetSpecialDate`, `AD_GetNextASN`, `AD_UpdateEIN`) — all on `ALC_Connection` (`DataModule.dfm:469,533,692`).

**Why it matters:** the entire M4 "relocate `VehicleOrder.sites` → `Inventory.sites`, repoint every cross-DB
reference, retire the VehicleOrder copy" workstream (lines 233, 299–300) is predicated on a `sites` table that I
cannot find being read by this app. Either (a) `sites` is a NEW table being introduced and the legacy never had
one (in which case "relocate + repoint + retire" is the wrong framing — it's "create + populate," with no
existing readers in *this* app to repoint), or (b) it exists and is consumed by **sibling apps** (GALC/MES),
not InventorySystem — which makes "retire the VehicleOrder copy" a cross-app blast-radius decision (see
SHOULD-FIX-4). The plan asserts (a-as-relocation) without a single file:line. **Send back to delphi-architect:
prove `VehicleOrder.sites` exists and enumerate its readers before scoping M4 around moving it.**

---

## SHOULD-FIX-4 — Retiring `VehicleOrder.sites`/`LINE` could break GALC + MES; dependency not checked. [new risk]

**Claim under test:** M4 "retire the VehicleOrder copy" of `sites` (line 233).

**Evidence of shared ownership:** the calendar decision itself (Q9) establishes that VehicleOrder is the
**shared** DB across Inventory, GALC, and MES, keyed by `Line.LineName` (`AD_GetSpecialDate-shared.sql:3,8`).
The calendar's `Line` table is explicitly kept shared. If `sites`/`LINE` config is co-resident in VehicleOrder
and also read by GALC/MES, **relocating or retiring it from InventorySystem's side breaks the siblings.** The
plan even half-acknowledges this ("only SITE config relocates; the calendar's `Line` table stays shared", Q9)
but never audits *who else* reads `VehicleOrder.sites`. Per the failure-mode checklist #9, this must be checked
against the `Delphi-VCL-Components` repo (GALC Session / NUMMI Tools), not assumed Inventory-only. Direction →
delphi-architect: enumerate cross-app readers of `VehicleOrder.sites`/`Line` before any retire step.

---

## SHOULD-FIX-5 — Q5 in-place recompute re-prices across a window boundary; the customer already has the original. [decision consequence]

**Claim under test:** Q5 — unsend recomputes costs in place via `fn_ManifestCostAt` (D6) on recreate, "keep the
header row," "recompute costs in place."

**The hazard:** `fn_ManifestCostAt` is window-aware and inclusive (Q3). If a manifest-cost window boundary
passes between the original send and the recreate (e.g. a price change effective the next day), the recreated
810 picks up the **current** window's price, not the price on the original `VC_PRODUCTION_DATE`. Two failure
shapes:
- **Re-price drift:** TEMA already received the original 810 at price P1; the recreate re-transmits at P2 →
  invoice mismatch / dispute. The "no ack received" use case (Q5) is precisely the situation where TEMA may
  have the original even though no 997 came back (997 lost, not the 810).
- **Intended vs hazard is undecided.** Q5 calls recompute "David's better solution" but does not state whether
  recompute keys on the *original production date* (correct: same price) or *recompute time* (re-price). Because
  `fn_ManifestCostAt` keys on the production/manifest date, it is *probably* stable — but only if the recreate
  passes the original `VC_PRODUCTION_DATE`, not `getdate()`. **The plan never pins this.** Direction →
  ignition-architect: assert recompute uses the original production date as the pricing instant; add a parity
  test that crosses a window boundary. Until pinned, RISK of silent re-pricing.

---

## SHOULD-FIX-6 — Q4 EIN gap/burn on send failure vs Q2 allocate-before-send ordering. [decision consequence]

**Claim under test:** Q4 — EIN is "the authoritative outbound control number tracked through the VAN... allocated
atomically and uniquely per site." Q2 — don't flip status until send commits.

**The ordering trap:** the legacy allocates the EIN up front via `AD_UpdateEIN` *during ASN create*
(`ASNSelect.pas:388,471`, before any 856 exists). If Ignition keeps allocate-on-create and the later send
fails, the EIN is **burned** (gap in the VAN control sequence). If instead it allocates at send-time to avoid
gaps, that contradicts the legacy timing (the EIN is stamped onto `INV_ASN_MST.IN_ASN_EIN` at create and is the
join key `REPORT_EDI856` reads). **The plan asserts "allocate atomically" but never resolves allocate-vs-commit
ordering or whether TEMA tolerates non-contiguous control numbers.** Most VANs/AS2 partners tolerate gaps in
ISA13 (it must be unique + ascending, not gapless) — but that's an assumption the plan should *state and verify
with TEMA*, because if TEMA requires contiguity, atomic-allocate-on-create + send-failure = a stuck sequence.
Direction → ignition-architect + delphi-architect: confirm TEMA's ISA13 contiguity requirement; decide
allocate timing; document the burn behavior.

---

## SHOULD-FIX-7 — C1 (remove 200-row cap) makes Ignition output DIFFER from legacy → false parity failures in Q16. [decision consequence]

**Claim under test:** Q15 — build faithful (Option A) so it parity-matches the Q16 dev-mirror; **C1 included**
(remove the silent ≤200-row truncation, surface dropped rows).

**The conflict, in the plan's own words:** post-cutover-enhancements.md:9 says the deferred items are deferred
*because* "(a) [they] would break parallel-run / dev-mirror parity (the cutover validation gate, Q16)." C1 is
**not** deferred — it ships in the faithful build (line 18) — yet C1 does exactly that: whenever a legacy run
had >200 rows, legacy silently dropped the overflow and Ignition will now show it. The Q16 diff will flag every
such case as a divergence. **C1 is held to a parity bar it is designed to violate.** Either the Q16 harness
must special-case "rows beyond legacy's 200-cap are expected-extra" (so a real regression in those rows isn't
masked), or C1 ships post-cutover with C2–C6. The plan acknowledges C1's defect-fix nature but **not its parity
impact.** Direction → ignition-qa: the dev-mirror diff needs a documented C1 exception band; otherwise C1
manufactures false failures and can mask real ones in the >200 region.

---

## SHOULD-FIX-8 — F1 `_HIST` `SELECT *` trigger hazard: enumerable NOW, will surprise mid-M4 if not pinned. [new risk]

**Claim under test:** §4/M4 — "adding `site_id` breaks `SELECT *` `_HIST` triggers unless the `_HIST` table gets
the column too." Staged additively.

**Evidence — the full blast radius IS knowable today (so enumerate it, don't discover it):**
- 4 `_HIST` tables: `INV_FORECAST_DETAIL_INF_HIST`, `INV_PARTS_STOCK_MST_HIST`, `INV_ASSY_BUILD_HIST`,
  `INV_OPEN_ORDER_INF_HIST` (`/tmp/inv_utf8.sql:43,138,676,809`).
- `INSERT ... _HIST SELECT * FROM inserted/deleted` triggers that break the instant the base table gains a
  column the `_HIST` table lacks: `:2664` (FORECAST del), `:3440` (FORECAST ins), `:3617` (FORECAST ins),
  `:4107` (PARTS_STOCK del), `:4269` (PARTS_STOCK ins), `:5475` (OPEN_ORDER del), `:7496` (OPEN_ORDER ins).
  **7 confirmed `SELECT *` history triggers across 3 tables.**
- `INV_ASSY_BUILD_HIST` is the safe one — it inserts via explicit `VALUES` (`:2514`), not `SELECT *`.

**Why it's only SHOULD-FIX not BLOCKER:** the hazard is real and proven on the spike, but it is fully
enumerable and additively fixable (add `site_id` to each `_HIST` table in lockstep). The risk is that the plan
says "every `_HIST` table + every `SELECT *` trigger" without the list — so a solo dev mid-M4 may miss one and
corrupt a history insert silently (a `SELECT *` with mismatched column count throws at runtime, halting the
base DML). Direction → ignition-architect: bake the 7-trigger / 3-table list above into the M4 site_id
checklist as a hard pre-flight. (Bonus: 25 triggers total — re-audit the other 18 for any `SELECT *` on tables
gaining `site_id`.)

---

## SHOULD-FIX-9 — `DELETE_ASNItem` / `INSERT_ASNDetail` dedup-by-manifest-ALONE is worse than the plan states. [decision consequence — verifies Q1]

**Claim under test:** Q1 — re-key the upsert + delete to `(site_id, IN_ASN_ID, manifest)`.

**Evidence the legacy is exactly as bad as Q1 says (good — Q1 is well-founded):**
- `INSERT_ASNDetail` (`/tmp/inv_utf8.sql:2704,2711`): existence-check and the `IN_QTY += @Qty` upsert key on
  `WHERE VC_MANIFEST_NUMBER = @Manifest` **alone** — no ASN, no site.
- `DELETE_ASNItem` (`:2800-2805`): takes only `@ManifestNumber`, deletes `WHERE VC_MANIFEST_NUMBER = @Manifest`
  — wipes that manifest from **every** ASN.

So Q1's `(site_id, IN_ASN_ID, manifest)` fix is correct and necessary. **One caveat the plan misses:** the
existence-check uses `SELECT * FROM ... IF @@rowcount` (`:2704`) — a non-atomic read-then-write. Under
parallel-run concurrency (two operators / two sites, which the legacy single-user app never had — checklist #4),
two `INSERT_ASNDetail` calls for the same key can both see rowcount 0 and both INSERT (duplicate manifest row),
or interleave the `+= @Qty` and lose an add. The re-key fixes *collision* but not the *race*. Direction →
delphi-architect: the re-keyed upsert needs the same SERIALIZABLE/UPDLOCK atomic-claim treatment the plan
already mandates for the renban counter (Carry 2) — the plan applies that rigor to renban but not to the ASN
detail upsert, which is now multi-user.

---

## SHOULD-FIX-10 — Solo-dev scope: M2 is underestimated; reporting (M3) is the safe cut, not the EDI loop. [new risk]

**Claim under test:** ~18–26 dev-weeks total; M1 5–7, M2 5–7, M3 3–4, M4 4–6, M5 1–2.

**Where it breaks:**
- **M2 is the most underestimated.** It bundles: a real X12 parser (ISA-separator-honoring, not byte-offset),
  delSL[4]→site routing, 997/AK9 with AK2/AK3/AK4 tolerance, 830 forecast ingest, 862/824 server-render, the
  processed-files idempotency ledger, the renban breakdown algorithm with atomic counter, the byte-exact `.ord`
  generator (supplier/logistics/archive), AND the order worksheet forecast-fill. That is **four distinct
  hard subsystems** (inbound EDI, renban, order-files, forecast-fill) in one 5–7wk box — each comparable in
  size to M1's whole outbound builder. M2 is realistically 8–11wk solo. The plan even flags its inputs as
  "inferred from call sites, not fully read" (Confidence note, line 465–469) — every one of those unverified
  procs (`CalculateASNFRS`, `SELECT_ASNSeq`, FRS-renban inserts, all ALC `AD_*`) sits in M2's critical path.
- **Single biggest timeline risk:** the unread proc bodies. The plan repeatedly defers "confirm with
  delphi-architect" (`CalculateASNFRS` body, ALC `AD_*` bodies, AK9 semantics, delSL[4] index). If any of
  these turns out to do more than assumed (as `REPORT_EDI856`'s embedded UPDATE just did — BLOCKER-1), the
  estimate moves. **Resolve the proc-body unknowns BEFORE committing to the M2 estimate.**
- **What can be cut to ship sooner:** M3 (reporting) — the plan already says it's additive read-only and can
  run alongside legacy (line 292). Reports are the lowest revenue-criticality (Rank 7) and the operator
  tolerated them failing 3× on the sample day. **M3 is the correct thing to defer/parallelize, not to sequence
  as a blocker.** The plan gets this right (line 318); just make it explicit that M3 is droppable from the
  critical path if the timeline slips.

---

## NIT-1 — Naming: "VehicleOrder DB" (calendar) vs "ALC DB" (EIN) may be the same physical DB. [plan flaw — clarity]

The Delphi calls `AD_GetSpecialDate`, `AD_GetNextASN`, and `AD_UpdateEIN` all by bare `dbo.` name on the **same
`ALC_Connection`** (`DataModule.dfm:469,533,692`; `DataModule.pas:3767`, `ASNSelect.pas:388`). The plan treats
the calendar as "VehicleOrder DB" (Q9) and EIN/site as "ALC DB" (Q4/§4) as if they're distinct. They resolve
through one connection today. Whether ALC and VehicleOrder are two DBs or one (or ALC has 3-part references into
VehicleOrder) is not established. This doesn't change the decisions but the plan should state the actual
physical topology so Q4's "move EIN to Inventory" and Q9's "leave calendar in VehicleOrder" don't accidentally
split one DB the wrong way. → delphi-architect: confirm what `ALC_Connection`'s Initial Catalog is.

## NIT-2 — `REPORT_EDI810` also has the `@EIN=0` C-status branch; confirm it does NOT self-flip like 856. [needs evidence]

`REPORT_EDI810` (`:3734`) mirrors `REPORT_EDI856`'s `@EIN=0`/`<>0` structure and filters `VC_INV_STATUS='C'`
(`:3753`). I did not fully read its `@EIN<>0` branch for an embedded `UPDATE INV_INV_MST` twin of BLOCKER-1.
Given 856 self-flips, **assume 810 may too** until read. → delphi-architect: read `REPORT_EDI810` lines
3760–3800 for an embedded status UPDATE before wrapping it.

---

## Coverage check — daily capabilities vs plan

Every one of the 17 daily capabilities in `daily-workflow-usage.md §5` maps to a Rank in §2. No daily capability
is uncovered. The `LogActLog` audit trail (capability #15) is correctly carried to §4 (audit profile). One
under-specified item: **SELECT SFT context selectors** (capability #14) — the order/ASN-context comboboxes (24
fired during the hot-call alone, log rows 138–163) are "PARTIAL" and not explicitly designed in any Rank; they
ride implicitly on the ASN/Order views. Low risk, but name them so they aren't forgotten in M1/M2.

---

## Decision-vs-legacy contradiction scan (Q1–Q17)

- **Q2 contradicts the legacy proc it claims to preserve** — see BLOCKER-1 (`REPORT_EDI856` self-flips).
- **Q4 + multi-site contradicts `UPDATE_EINStatus`** — see BLOCKER-2 (no site scope).
- **Q11 prose contradicts D1/Q14** — see SHOULD-FIX-1 (singular "the gateway's site").
- **Q1 is faithful and correct** — verified against `:2704/:2711/:2800` (SHOULD-FIX-9 adds the race caveat).
- **Q3 is sound** — `fn_ManifestCostAt` inclusive window is already migrated (PR #10/#14); nothing to build.
- **Q8 (logistics skip-by-config) is verified** — the source decode cites `OrderFormCreateF.pas:111/217-222`
  (`'NONE'` → no logistics file); consistent with config-driven, not part-type rule. SOUND.
- **Q5, Q6** — sound in direction; Q5 needs the re-pricing instant pinned (SHOULD-FIX-5).

---

# VERDICT

**The plan is NOT yet sound to start M1.** The foundation reuse, the faithful-build stance (Q15), the Excel-layer
retirement, and most of the 17 decisions are well-reasoned and evidence-backed. But three findings will corrupt
plant-floor / Toyota-exchange data the moment M1 parallel-runs, and they all stem from treating two
status-writing procs as read-only and from multi-site EIN collisions — exactly the failure classes this review
exists to catch.

**Top 3 David must fix before building M1:**

1. **BLOCKER-1 — Split the status-flip out of `REPORT_EDI856` (`:3695`).** As written, "wrapping" the proc to
   build the shadow 856 flips real `INV_ASN_MST` status to `'S'` in the shared DB, which makes legacy skip its
   own send. Shadow mode is not read-only until this UPDATE is removed and the flip is reimplemented per-ASN,
   coupled to send-commit (Q2). Also read `REPORT_EDI810` (`:3734`) for the same twin (NIT-2).

2. **BLOCKER-2 — Site-scope `UPDATE_EINStatus` (`:1722-1729`) before per-site EIN (Q4) goes live.** With per-site
   EIN sequences, `WHERE IN_ASN_EIN=@EIN` flips every site's matching EIN — a 997 for site A acks site B. Decide
   EIN namespacing (global-unique vs `(site_id, EIN)`) and make the proc match.

3. **BLOCKER-3 — Resolve the M1↔M2 boundary.** M1's gate ("997-accepted by TEMA") is unreachable without M2's
   997 ingester. Either pull the 997 path into M1 (and re-estimate) or downgrade M1's gate to byte-diff-only and
   move TEMA-accept to M2 — and say which.

Secondary, before M2/M4: prove `VehicleOrder.sites` actually exists and enumerate its GALC/MES readers
(SHOULD-FIX-3/4); resolve the unread proc bodies (`CalculateASNFRS`, ALC `AD_*`) that gate the M2 estimate
(SHOULD-FIX-10); and give the Q16 dev-mirror a documented C1 >200-row exception band (SHOULD-FIX-7).
