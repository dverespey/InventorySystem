# EDI 810 — Adversarial RE-VERIFY (byte + feed + money + D6 + Carry-5)

Independent attempt to REFUTE that the rebuilt 810 builder/drivers reproduce the legacy 810 (wire + feed)
with the decided divergences (D6 window-aware pricing, clean money, Carry-5 in-place unsend). Drivers/
atomicity/unsend were SHIP'd by the ignition-code-reviewer; this pass owns the independent BYTE + FEED +
MONEY confirmation. Default posture: reimplementation is wrong until proven equivalent on the SAME inputs.

Anchors (read, not assumed):
- Legacy bytes: `/Users/apple/Documents/GitHub/InventorySystem/EDI810Object.pas` (T810EDI).
- Legacy create driver: `/Users/apple/Documents/GitHub/InventorySystem/MainMenu.pas:2587-2657`
  (`CreateINVOICEClick`).
- Legacy recreate driver: `/Users/apple/Documents/GitHub/InventorySystem/ASNInvoice.pas:857-895`.
- Legacy procs (live dump): `/tmp/inv_utf8.sql` — `REPORT_EDI810` :3734, `REPORT_EDI810Recreate` :3706,
  `INSERT_INVInfo` :2567, `UPDATE_INVItems` :3406, `UPDATE_INVUnsend` :3387.
- Live legacy proc bodies on the spike (`Inventory` DB): `REPORT_EDI810` (window-blind, INNER JOIN), and
  `fn_ManifestCostAt` (D6 windowed table fn). D6 proc text:
  `/Users/apple/Documents/GitHub/InventorySystem/docs/analysis/reporting/spike-report-procs-d6.sql:60-92`.
- Rebuild: `/Users/apple/Documents/GitHub/InventorySystem/docs/analysis/edi/810/project-library/edi810/code.py`;
  feed `.../docs/analysis/edi/810/spike-edi810-feed.sql`; tests `scripts/e2e/test_edi810_build.py` (61 PASS),
  `scripts/e2e/test_edi810_e2e.py` (53 PASS, spike left as-found).

Method note: the spike's live `REPORT_EDI810` is still the LEGACY window-blind body (verified via
`OBJECT_DEFINITION`); the D6 body lives only in the `.sql` file + the rebuild's inlined SELECT. So legacy-vs-
D6 was diffed by running BOTH SELECT forms by hand (no proc EXEC of the `@EIN<>0` branch — it self-flips).

---

## What I PROVED equivalent (counterexamples that did NOT refute)

### BYTE faithfulness vs EDI810Object.pas — CLEAN, no hidden trailing-sep trap
Per-segment byte audit of a multi-manifest/multi-item build (parts A,B on manifest `76061857` '7'-prefix →
M391; part C on `80061900` → M390):
- ISA = 17 elements, ends on ISA16 (`>`); GS01=`IN` (not the 856's `SH`), GS ends on `004010`; ST=`ST*810*<EIN>`;
  BIG ends on `SiteSupplierCode` (BIG02). IT1 = 12 populated elements, ends on the dock code `D01` (NO trailing
  sep). REF/DTM/TDS/CTT/SE/GE/IEA all end on a value.
- NO `~` anywhere (CRLF-only — segment terminator read but never emitted, faithful legacy bug). NO `**` (no
  empty interior element). File ends on one trailing CRLF (legacy `Writeln`). The 856 TD1/LIN trailing-sep
  quirk did NOT bleed in.
- IT101 M391/M390 by manifest `'7'` prefix; interior REF carries the PREVIOUS manifest, trailing REF the
  FINAL manifest; interior DTM = current row's pickup, trailing DTM = last pickup — all match the .pas loop.
- EIN `%09d` in all 7 control positions (ISA13/GS06/ST02/SE02/GE02/IEA02), all consistent.
- **SE01 independently re-derived from `EDI810Object.pas` fSECount** (ST :210, BIG :231, IT1 ×3 :310, interior
  REF :278 + DTM :286, trailing REF :325 + DTM :334, TDS :353, CTT :381, SE :395) = 12 → matches the rebuild's
  counted SE01=12 and the test's `expected` SE01=12. CTT01 = #IT1 = 3. Counted, not a magic offset.
- **The unit test's `expected[]` is .pas-derived, not rebuild-derived, and is non-vacuous:** injecting a
  GS01→`SH` regression and a TDS-scale regression into a copy of `code.py` produced `GS*SH*...` and `TDS*1004`,
  both of which the `expected[]` array rejects. The test would FAIL on a real regression.

### CLEAN MONEY — correct, not just different (exact arithmetic, no float drift)
Ran the rebuild's `_tds_total` / `_it104` (Decimal, ROUND_HALF_UP) against a faithful simulation of the legacy
Delphi TDS surgery (`EDI810Object.pas:339-350`):

| total      | legacy TDS (buggy) | rebuild TDS (clean) | note |
|------------|--------------------|---------------------|------|
| `1234.5`   | `12345`            | `12345000`          | 1-digit frac: legacy off by 10000x; rebuild correct |
| `2292.40`  | `22924`            | `22924000`          | same class (the unit-test fixture) |
| `0.5`      | `05`               | `5000`              | legacy corrupts |
| `100`      | `1000`             | `1000000`           | whole-dollar: legacy `pos('.')=0` malformed; rebuild correct |
| `1000`     | `1000`             | `10000000`          | whole-dollar malformed |
| `12.55`    | `125500`           | `125500`            | 2-digit frac: AGREE |
| `319.68`   | `3196800`          | `3196800`           | 2-digit frac: AGREE |
| `58093.75` | `580937500`        | `580937500`         | 2-digit frac: AGREE (the D6 EIN-5692 total) |

The legacy bug fires ONLY on 1-digit-fraction and whole-dollar totals; the rebuild is correct on all and
IDENTICAL on the 2/3/4-digit-fraction cases. The divergence is bounded to exactly the two documented bugs
(DECISION 810-1). IT104 = fixed scale-4 with a literal `.`, locale-independent (e.g. `12.5`→`12.5000`),
matching `MO_PRICE`'s native money/scale-4. Rounding boundary: inputs are `money`/scale-4 and `qty` is int, so
`price*qty` stays scale-4 — the 5th-decimal tiebreak never fires on real data; where it would, ROUND_HALF_UP
== T-SQL `ROUND` (half-away-from-zero), so no banker's-rounding divergence.

### FEED == legacy D6, BOTH branches; D6 over-bill RE-PROVEN on real EIN 5692
- The recreate branch the LIVE app actually calls is `REPORT_EDI810` `@EIN<>0` (`ASNInvoice.pas:860`), NOT the
  dead `REPORT_EDI810Recreate` (different WHERE: `i.VC_INV_STATUS='C'`, no EIN filter). The rebuild's
  `_RECREATE_FEED_SQL` (`WHERE iim.IN_INV_EIN=@EIN`) matches the proc that is really invoked. Confirmed.
- Drift guard holds: `_CREATE_FEED_SQL`/`_RECREATE_FEED_SQL` are char-identical to the marked `.sql` bodies
  (e2e (1)); both are pure SELECTs (no self-flip / `UPDATE_INVRecreate` / any DML).
- Feed-row parity vs the legacy `@EIN=0` SELECT proven row-for-row on synthetic unbilled data (e2e (1)).
- **D6 over-bill RE-PROVEN, real EIN 5692** (read-only; SELECT forms run by hand, no self-flip):
  - Legacy window-blind (INNER JOIN on part code alone) → 2 lines: `715×81.25 = 58,093.75` (window
    `20180901..20260901` covers shipment `20201209`) + `45×82.00 = 3,690.00` (window `20220719..20280229`
    does NOT cover `20201209`). Total `$61,783.75`.
  - D6 (`CROSS APPLY fn_ManifestCostAt`) → 1 line, `$58,093.75`; drops the wrong-window line.
  - Over-bill = exactly `$3,690.00`. The rebuild uses D6 → corrected. (Note: 5692's correct total has a
    2-digit fraction, so this is a pure WINDOW fix; the TDS-format fix is independent and also correct.)

### EIN — per-site, atomic at create, reused at recreate, shared 856/810 sequence
e2e (2)/(3): create allocates `IN_EIN_SEQ+1` atomically (`UPDATE ... OUTPUT INSERTED.IN_EIN_SEQ`), bumps the
counter by exactly 1; recreate REUSES the EIN and does NOT bump. The rebuild's atomic OUTPUT is strictly safer
than legacy's read-then-bump (`SiteEIN+1`) and faithful on the single-threaded daily run. `INV_INV_MST`
column order (`IN_INV_ID identity, IN_INV_EIN, VC_INV_STATUS, VC_LAST_UPDATE, VC_ADD`) confirms the rebuild's
named-column INSERT maps exactly to the legacy positional `INSERT_INVInfo VALUES(@Ein,'S',@AddDate,@AddDate)`.

### Self-flip excluded + Carry-5 in-place unsend
Feeds are side-effect-free; `UPDATE_INVRecreate` never run. e2e (5): unsend reverts status to `'C'` IN-PLACE,
re-pools detail (`IN_INV_ID=null`), KEEPS the header + EIN + audit — NOT the legacy `UPDATE_INVUnsend`
hard-DELETE (`/tmp/inv_utf8.sql:3391-3392`), and faithful to that proc's commented-out original intent
(`--UPDATE INV_INV_MST SET VC_INV_STATUS='C'`, :3390). Atomicity (e2e (4)): a post-`.tmp`-write commit fault
rolls back the EIN bump + header + link and leaves NO final 810 (only a swept `.tmp`).

---

## FINDINGS

### SHOULD-FIX 1 — CREATE driver does NOT split files per pickup-date (legacy did; structural + filename byte gap)
- **Claim under test:** the rebuild reproduces the legacy 810 CREATE.
- **Legacy:** `MainMenu.CreateINVOICEClick` (`MainMenu.pas:2613-2654`) is a `while not Eof` loop over the whole
  `REPORT_EDI810 @EIN=0` result; each iteration builds ONE 810 whose IT1 loop BREAKS on a pickup-date change
  (`EDI810Object.pas:263-266`), then writes a SEPARATE file and allocates a NEW EIN + `INSERT_INVInfo`. Net:
  **one invoice / one EIN / one file per pickup date** across the unbilled set.
- **Rebuild:** `create_invoice` builds exactly ONE invoice / ONE EIN / ONE file for the ENTIRE unbilled feed,
  regardless of pickup-date boundaries (`code.py:531-575`). `build_810` does not break on a pickup-date change;
  a mixed-pickup feed collapses to one envelope whose ONLY DTM (single manifest) carries the LAST pickup date,
  silently dropping the earlier one(s).
- **Reachable?** YES. The unbilled-'A' detail spans **2,312 distinct pickup dates** over 39,707 rows on the
  live snapshot (counterexample query below). A daily run with unbilled detail crossing a pickup-date boundary
  would: emit 1 file where legacy emits N; burn 1 EIN where legacy burns N; and put wrong DTM dates on lines.
  - Counterexample (live, read-only):
    `SELECT COUNT(DISTINCT a.VC_PRODUCTION_DATE) FROM INV_ASN_MST a JOIN INV_ASN_DETAIL_MST d ON
    a.IN_ASN_ID=d.IN_ASN_ID AND a.VC_ASN_STATUS='A';` → **2312**.
  - Pure-builder counterexample (two pickup dates, one manifest): build_810 emits a single `DTM*050*20260611`
    and drops `20260610`; legacy would have emitted two files.
- **Classification:** code defect in the CREATE driver's grouping (the BYTE content + EIN count + filename all
  depend on it). NOTE: the RECREATE path is safe — no real invoice spans >1 pickup date (0 of 2,550 invoices
  with linked detail; query below), so the legacy IT1 file-break never fires on recreate. The gap is
  CREATE-only.

### SHOULD-FIX 2 — CREATE filename drops the LineName the legacy create filename carried
- **Legacy create filename** (`MainMenu.pas:2623`): `'810' + copy(PickUpDate,5,4) + LineName + '.txt'`
  → e.g. `8100610<LineName>.txt`.
- **Legacy recreate filename** (`ASNInvoice.pas:872`): `'810' + copy(PickUpDate,5,4) + '.txt'` → `8100610.txt`.
- **Rebuild** `_filename_810` (`code.py:339-347`): `'810' + mmdd + '.txt'` = `8100610.txt` — matches RECREATE
  but NOT CREATE (no LineName). Only 1 distinct line name among 'A' ASNs today, so the collision/mismatch is
  low-impact now, but it is a real create-path filename byte difference (and couples to SHOULD-FIX 1: legacy
  separated files by pickup-date AND line name).
- **Classification:** code defect (create filename), low severity at current data cardinality.

### NIT 1 — true TEMA byte-parity + exact IT104 scale remain UNPROVABLE (data-vintage / no golden)
- The spike `INV_SITES` carries PLACEHOLDER identity (abbr/DUNS/supplier/dock/sep/EDI-mode), so the ISA/GS/BIG
  BYTES the driver emits are NOT the TEMA wire — only STRUCTURE is asserted. No golden 810 exists, so
  byte-for-byte TEMA parity and the exact IT104 scale (4 vs trimmed) cannot be proven from available data. The
  tests honestly disclaim this (`test_edi810_e2e.py` docstring). This is an unfixable data-vintage gap until
  real site values + a golden 810 are loaded — correctly flagged, not a defect.

### RISK-1 assessment (priceless-line under-bill) — FAITHFUL PORT, not a regression
- The CROSS-APPLY inner drop of un-priced lines mirrors LEGACY behavior: legacy `REPORT_EDI810` uses an INNER
  `JOIN INV_MANIFEST_COST_MST`, and **24,586 / 39,707** detail rows have NO cost row at all for their part —
  legacy's inner join already drops those. So "drop the priceless line" is a faithful port, not a new defect.
- D6 vs legacy drop DIFFERENT sets on billed detail: legacy keeps **15,121** (any cost row exists), D6 keeps
  **14,356** (a covering WINDOW exists at the production date). The 765 extra lines D6 drops are exactly the
  wrong-window (mis-priced) lines D6 is designed to exclude — the over-bill family, correct to drop. The
  documented mitigation (run `priceless-lines-diagnostic.sql`, EDI810 branch, expect 0 before each run) is the
  right operational guard.
- **Classification:** faithful legacy port + intended D6 improvement; keep the pre-run diagnostic gate.

---

## VERDICT

For a SINGLE invoice / SINGLE pickup date (the RECREATE path, and any CREATE run whose unbilled detail shares
one pickup date), the rebuilt 810 is **byte-faithful** — segment structure, ordering, separators (CRLF-only),
EIN `%09d`×7, GS01=`IN`, BIG02=SupplierCode, IT101 M391/M390, the per-manifest REF/DTM breaks, counted
SE01/CTT01, and no trailing-sep/empty-element traps — with **correct clean money** (TDS01 implied-decimal
scale-4 integer and IT104 fixed scale-4, exact Decimal arithmetic, the two legacy TDS bugs fixed and bounded)
and **correct D6 window-aware pricing** (over-bill RE-PROVEN at exactly $3,690.00 on real EIN 5692). EIN
provenance/reuse, the pure (non-self-flipping) feeds, atomicity, and the Carry-5 in-place unsend all hold.
The only currently UNPROVABLE item is true TEMA byte-parity + the exact IT104 scale (no golden 810; placeholder
site values) — honestly disclaimed (NIT 1).

HOWEVER there is a real DIVERGENCE on the CREATE path: **the rebuild does not split the invoice file per
pickup-date the way the legacy `CreateINVOICEClick` loop does** (SHOULD-FIX 1) — reachable (2,312 distinct
pickup dates in the unbilled-'A' set) and affecting file count, EIN count, and the DTM bytes — plus the create
filename drops the LineName (SHOULD-FIX 2). These are CREATE-driver grouping/naming defects, not builder or
money defects; the byte builder, feed, money, D6, EIN, and unsend are equivalent.

So: **NOT yet PROVEN equivalent on the CREATE path** (the per-pickup-date file split must be reproduced before
cutover); **byte/feed/money/D6/EIN/unsend PROVEN equivalent** for the single-pickup-date (recreate + same-day
create) case, modulo the documented TEMA-golden / IT104-scale-pending gap.

---

# RE-VERIFY 2 — the per-pickup-date REWORK (FIX-1: one invoice/EIN/file per distinct pickup date)

Focused adversarial pass on the reworked `create_invoice` (`docs/analysis/edi/810/project-library/edi810/
code.py` — `_group_by_pickup_date` / `_create_one_invoice` / the date-scoped link). SHOULD-FIX 1 from the
prior pass was reworked; this pass asks whether the rework is feed/grouping-equivalent to the legacy, or
whether the new group-by-distinct-date can diverge from the legacy break-on-change. Read: code.py (reworked),
`scripts/e2e/test_edi810_e2e.py`, legacy `MainMenu.pas:2587-2657` + `EDI810Object.pas:241-369`,
`InsertINVInfo`/`UPDATE_INVItems` (DataModule.pas:5000-5056; procs CreateInventory.sql:2567/3406),
`REPORT_EDI810` (CreateInventory.sql:3734). Read-only on `Inventory_Live`; rolled-back tx on `Inventory`.

## How the legacy actually groups (re-read, not assumed)

- `MainMenu.pas:2613-2654` walks the `REPORT_EDI810` @EIN=0 cursor (ORDER BY `VC_MANIFEST_NUMBER`,
  CreateInventory.sql:3753). Each `EDI810.Execute` consumes rows until `EDI810Object.pas:263-266` **BREAKS the
  IT1 loop on a `PickUpDate` CHANGE** — i.e. legacy splits on **contiguous runs of equal PickUpDate in
  manifest order**, NOT on distinct date value. One file + one new EIN (`SiteEIN+1`, MainMenu.pas:2619) +
  one `InsertINVInfo` per run.
- The rebuild's `_group_by_pickup_date` (code.py:473-501) partitions by **DISTINCT PickUpDate value** — all of
  a date's rows in ONE invoice, regardless of manifest adjacency.

These two definitions agree IFF every pickup date forms a SINGLE contiguous run in manifest order. They
DIVERGE when a date's rows are split into NON-ADJACENT runs by an intervening other-date row.

## BLOCKER-class divergence (data-reachable): a date split into non-adjacent manifest runs

The driving fact (`Inventory_Live`, 39,791 billed detail rows — bounded, the small Inventory side):
manifest numbers come in three sort-disjoint **prefix families** — `7…` (39,290 rows), `5…` (488), `T…` (13).
All `5…` sort before all `7…` before all `T…`, so a single pickup date carrying two families has its rows
NON-ADJACENT in `ORDER BY VC_MANIFEST_NUMBER`.

Counterexample 1 (observed, real billed data, `Inventory_Live`, read-only):
- **203 of 2,319 distinct pickup dates carry >1 manifest-prefix family** (query: `perdate` HAVING
  COUNT(DISTINCT LEFT(VC_MANIFEST_NUMBER,1))>1 = 203). Independently, a LAG-run analysis over the full
  manifest-ordered feed finds **203 dates forming >1 contiguous run** (max 3 runs; e.g. date `20240111`).
- Concrete: pickup date **`20260606`** (the latest billed date) carries family `5` (manifest `52089698`) AND
  family `7` (manifests `760606xx`). In manifest order, `52089698` (date 20260606) is followed by
  `52089913` (date 20260619) then a long `70010xxx` block (dates 2020-01-02…) before the `760606xx`
  (date 20260606) rows. So `20260606` occupies TWO non-adjacent runs.

Counterexample 2 (constructed on the rebuild `Inventory` DB, rolled-back tx, left as-found — leftover=0):
three unbilled status-'A' ASNs — A(manifest `59990001`, date `20991201`), B(`79990001`, date `20991201`),
C(`69990001`, date `20991202`). The unbilled-only @EIN=0 feed (ORDER BY manifest) is:

| feed order | manifest | PickUpDate | legacy break-on-change |
|---|---|---|---|
| 1 | 59990001 | 20991201 | BREAK -> file 1 (date 20991201) |
| 2 | 69990001 | 20991202 | BREAK -> file 2 (date 20991202) |
| 3 | 79990001 | 20991201 | BREAK -> file 3 (date 20991201 **AGAIN**) |

- **LEGACY result:** 3 invoices / 3 EINs / 3 files; pickup date `20991201` is billed across **TWO** separate
  810 files (file 1 + file 3), each with its own EIN, each carrying its own ISA/GS/BIG envelope.
- **REBUILD result:** `_group_by_pickup_date` -> 2 groups -> **2 invoices / 2 EINs / 2 files**; pickup date
  `20991201` is billed in **ONE** 810 (both manifests `5…` and `7…` in one IT1 list under one EIN).

So on this input the two produce a DIFFERENT number of invoices, a DIFFERENT number of EINs consumed from
`INV_SITES.IN_EIN_SEQ`, and DIFFERENT file contents/count for the same pickup date. This is exactly the
"same pickup date appears in two non-adjacent manifest runs" case the task asked me to prove — and it is
**data-reachable**: 203 real dates already exhibit the multi-family pattern; any create run in which an
unbilled `5…` and an unbilled `7…` (or `T…`) row share a pickup date, with an intervening other-date
unbilled row, triggers it.

**Honest scope of the proof:** the live snapshot is ALL-billed, so I could not observe a *real* unbilled
multi-family batch in flight; the 203-date figure is the structural prevalence and the 3-file split is proven
on a constructed (but schema-faithful, rolled-back) unbilled batch. Whether a single nightly create run ever
holds two families for one date unbilled simultaneously depends on the shipping cadence I cannot replay from
an all-billed snapshot — so the *frequency* is a data-vintage unknown, but the *reachability* is proven.

- **Classification:** behavioral divergence (CREATE grouping). It is the rebuild being arguably MORE correct
  (one date -> one invoice is the natural billing unit; the legacy double-files a date as an artifact of
  manifest-family sort order). But "more correct" is NOT "equivalent": file count, EIN consumption, and the
  per-file byte content differ on a reachable input. Whether to (a) accept the rebuild's group-by-distinct as
  an intended improvement (then the prior SHOULD-FIX is closed and this is a LOCKED divergence to document
  for TEMA, like D6/clean-money), or (b) reproduce the legacy break-on-change exactly, is a David/architect
  decision — NOT something I can call "proven equivalent." Bounce to the architects.

## Second, independent divergence: legacy links only the FIRST ASN of a run (under-bills); rebuild links ALL

Re-reading the legacy link path: `MainMenu.pas:2615` sets `Data_Module.ASN := EDI810DataSet.FieldByName
('ASNid')` ONCE, from the cursor row at the START of each `Execute` (the run's first row). `InsertINVInfo`
(DataModule.pas:5043-5048) then calls `UPDATE_INVItems(@INVID, @ASNID := fASN)` — a SINGLE ASN id —
and `UPDATE_INVItems` (CreateInventory.sql:3421) does `… WHERE IN_INV_ID is null AND IN_ASN_ID = @ASNID`.
So the legacy marks **only the first ASN's** detail as billed, even though the IT1 loop EMITS lines for every
ASN sharing that pickup-date run.

Counterexample (real, `Inventory_Live`): pickup date `20220824`, family `7`, spans THREE ASNs whose manifest
ranges interleave — ASN 3389 (`72082401..72082452`, 22 rows), ASN 3421 (`72082431`, 1 row), ASN 3422
(`72082435`, 1 row). One contiguous same-date run, first row's ASN = 3389. Legacy `UPDATE_INVItems(_,3389)`
bills ONLY ASN 3389; ASNs 3421/3422 get their lines EMITTED in the 810 but stay `IN_INV_ID IS NULL`
(re-billable on the next run). The rebuild iterates `asnIds = sorted(set(...))` over the whole group
(code.py:534, 575-580) and links ALL three ASNs.

- **Classification:** the rebuild FIXES a legacy under-link bug (legacy could re-emit the unlinked ASNs'
  lines on a later create = duplicate bill). Again a correctness improvement, again NOT byte/state-equivalent
  to legacy on a reachable input (post-run `IN_INV_ID` state differs). Document as an intended divergence or
  bounce to architects; do not call it equivalent silently.

## What the rework got RIGHT (re-confirmed)

- **Date-scoped link is safe (attack #2 — CLEAN).** `VC_PRODUCTION_DATE` exists ONLY on `INV_ASN_MST`
  (header); `INV_ASN_DETAIL_MST` has no such column (sys.columns check: detail=0, header=1). `IN_ASN_ID` is
  unique in the header (rowcount = distinct count) and **0 ASNs carry >1 distinct production date**. So an
  ASN has exactly one pickup date; the rebuild's `WHERE d.IN_ASN_ID=? AND a.VC_PRODUCTION_DATE=?` selects the
  SAME detail as a pure per-ASN link — the date predicate is redundant-but-correct and can NEVER mis-link
  across dates. The e2e "NO cross-date detail leak" check (asn1->inv2=0, asn2->inv1=0) confirms it on a
  2-date batch. No cross-date leak.
- **EIN per date (attack #3 — CLEAN).** `_create_one_invoice` allocates each group's EIN with an atomic
  `UPDATE INV_SITES SET IN_EIN_SEQ=IN_EIN_SEQ+1 OUTPUT INSERTED.IN_EIN_SEQ` (code.py:544-546), one per group.
  e2e multi-date: 2 dates -> EINs 9211, 9212 (distinct + sequential), `IN_EIN_SEQ` bumped by exactly 2, and
  the atomicity case proves a failed group's bump rolls back. No reuse/collision. (This is the per-Execute
  `SiteEIN+1` / `AD_UpdateEIN` of the legacy loop — one EIN per emitted file. NB: it follows that wherever
  the legacy double-files a date, legacy consumes one MORE EIN than the rebuild for that batch — a corollary
  of the grouping divergence above, not a separate defect.)
- **Per-group file content (attack #1 detail).** Each group's `build_810` gets ONLY that date's rows + the
  right DTM*050; e2e confirms inv1 DTM==[PDATE], inv2 DTM==[PDATE2], CTT01 per file == its own #IT1, distinct
  filenames `8100610.txt`/`8100611.txt`. The builder/feed/money/D6/EIN-reuse/unsend did NOT regress: the full
  e2e is **72 PASS / 0 FAIL** (feed-row parity vs legacy SELECT, GS01='IN', %09d EIN×7, SE01/CTT01 counted,
  CLEAN TDS/IT104, recreate EIN-reuse, atomicity, Carry-5 unsend).

## Parity-METHOD flaw (test does NOT exercise the divergence)

`test_edi810_e2e.py` step (6) claims the multi-date feed "INTERLEAVES the two dates' rows … proving the
grouping splits by date, not by feed adjacency." It does not. Its rows sort (by manifest) as
`76061857`/20260611, `76061857`/20260610, `80061900`/20260610 — the two dates are already adjacency-aligned
(legacy would also emit 2 files), so legacy and rebuild AGREE on this fixture. The test therefore proves the
rebuild self-consistently splits by distinct date, but it **never constructs the divergent input** (one date
in two non-adjacent runs) — the one case where legacy != rebuild. Green here is self-consistency, not
legacy-equivalence on the hard case. (Fixture-fidelity discipline: the test cannot prove equivalence on an
input it does not contain.)

## RE-VERIFY 2 VERDICT

The per-date rework is **feed/data-correct and did not regress** the builder/feed/money/D6/EIN/unsend (72/72
e2e green; date-scoped link proven leak-free; EIN-per-date atomic, distinct, sequential, no collision; one
ASN = one pickup date so the date-scope is correctness-neutral).

But the rework is **NOT grouping-equivalent to the legacy**. The legacy splits on **contiguous PickUpDate runs
in manifest order**; the rebuild splits on **distinct PickUpDate value**. These DIVERGE on a data-reachable
input — proven by counterexample: 3 unbilled ASNs (manifests `5/6/7`, dates `D/D2/D`) yield **legacy = 3
files/3 EINs (date D double-filed)** vs **rebuild = 2 files/2 EINs (date D single-filed)**; and 203 of 2,319
real pickup dates already carry the multi-manifest-family pattern that triggers it. A second divergence:
legacy links only the FIRST ASN of a run (under-bills, leaving emitted lines re-billable), while the rebuild
links every ASN in the group. Both divergences are the rebuild being arguably *more correct* — but "more
correct" is not "equivalent," and both change file count / EIN consumption / post-run billed-state on
reachable inputs.

So: **YES, there is a non-adjacent-date-recurrence divergence.** The per-date rework is NOT proven equivalent
to the legacy break-on-change; it is a deliberate-looking improvement that must be ACCEPTED+LOCKED+documented
(like D6/clean-money) by David/the architects, or replaced with a faithful break-on-change, before cutover.
It is NOT a silent equivalence. (The e2e test's green does not cover this — it omits the divergent fixture.)
