# M1 ASN-Creation — End-to-End Architecture + Implementation

**Keystone:** the morning revenue starter (M1 Rank 1). One operator click — "Create ASN entries
only" — turns a window of GALC-built vehicles into an ASN header + N manifest-detail rows + the
"(No Ratio)" remainders, under one ASN id. **Revenue-critical, can't-have-issues:** a wrong
broadcast-code conversion ships the wrong manifest → mis-ship / mis-bill to a Toyota plant.

**Status of this doc.** It is the *single assembled* architecture + implementation reference for the
keystone: legacy Delphi behavior → the current Ignition rebuild, centred on the `AD_FRSPULL`
broadcast-code→numbers conversion. It is an **honest production-readiness assessment, not a
correctness sign-off** (see §5/§6). Read §6 (MUST-FIX) and §7 (sign-offs) before shipping.

**Authoritative basis (live wins over dump).** The conversion math is documented on the
**live running `AD_FRSPULL`** as captured in `docs/analysis/production-readiness/AD_FRSPULL-shared.sql`
— `char(3)`, ground concat `ModelYearCode + WHEEL(vd2) + TIRE(vd1)`, spare `<> 'M'` filter. This is
the proc the rebuild actually reads and was adjudicated against the restored backup
(`adversary-findings.md` "RESOLUTION OF THE #1/#2 COUPLED GAP — DEFINITIVE"). The on-disk
`/Users/apple/Documents/FP docs/SQL/VehicleOrder.sql` (UTF-16, Script Date 06/10/2026) is a **stale,
different** dump (`char(21)`, reversed concat, no `<> 'M'`) and is **quarantined** — see the
version-conflict box in §2.

**Sources synthesized (all on disk, all read in full):**
- `docs/analysis/production-readiness/sql/AD_FRSPULL-analysis.md` (live VehicleOrder, 2.33M Vehicle rows)
- `docs/analysis/production-readiness/sql/SELECT_ForecastDetailBCASN-analysis.md`
- `docs/analysis/production-readiness/sql/asn-write-chain-analysis.md`
- `docs/analysis/production-readiness/sql/delphi-fanout-confirmation.md`
- `docs/analysis/production-readiness/sql/equivalence-map.md`
- `docs/analysis/production-readiness/sql/adversary-findings.md` (sql-adversary, live-DB counterexamples)
- `docs/analysis/production-readiness/sql/design-review-findings.md` (doc/design adversary)
- `docs/analysis/production-readiness/m1-asn-creation-spec.md`
- `docs/analysis/production-readiness/AD_FRSPULL-shared.sql`
- Rebuild: `docs/analysis/edi/project-library/asn/code.py`
  (`computeAsnDetails` :85-179, `create_asn` :207-363)
- Parity harness: `scripts/e2e/test_create_asn_parity.py`, unit: `scripts/e2e/test_asn_fanout.py`

---

## 1. Overview — the end-to-end create-ASN dataflow

One "Create ASN entries only" click (`ASNSelect.pas:369 CreateASNEntries_ButtonClick`) fans a
datetime window of built vehicles into ASN rows across **two databases**: `VehicleOrder` (the GALC
build data, ALC side) and `Inventory` (the parts/forecast/ASN side).

```
 OPERATOR: line + production date + [start,end] datetime window + start/last seq + ship qty
     |
     v
 [ ASNSelect.pas ]  CreateASNEntries_ButtonClick (:369)
     |
     | (0) idempotency guard
     |      SELECT_ASNSeq(@LineName,@PDate)            [Inventory]   asn-write-chain §3
     |        row exists for (line, prodDate, START_SEQ<>-1) -> LOCK UI, disable Create
     |
     | (1) header first (needs its IDENTITY for the details)
     |      INSERT_ASNInfo(... OUTPUT @ASNID)          [Inventory]   status 'C', SCOPE_IDENTITY()
     |        fRecordID := @ASNID
     |
     | (2) CalculateASNFRS  (Delphi method, NOT a proc; DataModule.pas:5106)   the fan-out
     |      |
     |      |--(2a) per (line, window):
     |      |        AD_FRSPULL(@begindate,@enddate,@LineName)   [VehicleOrder, char(3)]
     |      |          -> one row per BROADCAST CODE: (BC, ORDERS, VEHICLES)      <== §2 centrepiece
     |      |
     |      |--(2b) per BC:
     |      |        SELECT_ForecastDetailBCASN(@BCode, @EffMonth) [Inventory]
     |      |          WHERE @BCode LIKE VC_BROADCAST_CODE  (column is the PATTERN)
     |      |          -> recipe rows: part + IN_ASSY_QTY + IN_TIRE/WHEEL_RATIO + cost id
     |      |        PRE-LOOP COST GUARD: any IN_MANIFEST_COST_ID IS NULL -> RAISE, whole create fails
     |      |
     |      |--(2c) per recipe row:
     |      |        branch on ORDERS<=5 (No-Ratio: 1 row, no ratio, break) else ratio split
     |      |        qty = VEHICLES*IN_ASSY_QTY  or  banker_round(VEHICLES*IN_ASSY_QTY*IN_TIRE_RATIO/100)
     |      |        manifest = '7' + prodDate[year-digit+MM+DD] + VC_ASSY_MANIFEST_NUMBER
     |      |        INSERT_ASNDetail(@ASNID, manifest, part, qty)  [Inventory]  accumulate-on-repeat
     |      |
     |      `--(2d) POST-LOOP COST WARN:
     |               SELECT_ASNMissingCost(@ASNID)        [Inventory]  date-windowed; log only, NEVER aborts
     |
     | (3) AD_UpdateEIN  (legacy)  UPDATE Site SET SiteEIN+1  [VehicleOrder]  no WHERE; OUTSIDE the txn
     |
     v (4) commit
```

**Critical legacy structural fact — NO surrounding transaction (the partial-ASN bug).** In the
legacy, `BeginTrans`/`CommitTrans` is on **`Inv_Connection` only** (`ASNSelect.pas:372,391`); each
`ExecProc` otherwise auto-commits, and steps 0b/2a/3 run on the **un-transacted ALC_Connection**
(`delphi-fanout-confirmation §e`; `asn-write-chain §0`). A mid-loop abort (a missing cost on BC #3)
leaves the header + BC #1/#2 detail rows persisted — an orphan partial ASN — and the EIN bump (3)
can advance even when the Inventory side rolls back (EIN gap). **The rebuild wraps the header + all
details in one Gateway transaction** and moves EIN allocation in-transaction at-send (a deliberate
fix; see §3 PART E and §6).

---

## 2. The `AD_FRSPULL` conversion — the centrepiece

`AD_FRSPULL` (DB **VehicleOrder**, the shared GALC build DB; read-only cross-DB from the rebuild) is
the proc that produces every number the fan-out multiplies. It is called **once per (line,
production-date window)** and returns, **per broadcast code**, a vehicle count and an implied order
count. Body verified byte-identical (SQL) to `AD_FRSPULL-shared.sql` on the live VehicleOrder backup
(`AD_FRSPULL-analysis §0`).

### 2.1 Signature

```sql
CREATE PROCEDURE [dbo].[AD_FRSPULL]
    @begindate datetime, @enddate datetime,
    @Start int, @Last int,            -- DECLARED BUT NEVER REFERENCED (inert — §2.7)
    @LineName varchar(50)
```
Returns 4 columns: **`BC char(3)`, `ORDERS int`, `VEHICLES int`, and a literal `''`** (4th column is
always blank — the Delphi caller ignores it). Two `SELECT … GROUP BY` blocks `UNION`'d, `ORDER BY BC`.

### 2.2 BC composition (AUTHORITATIVE live char(3) basis)

GALC stores vehicle attributes EAV-style (`vehicledata` → `DataItem.DataItemDescription`). The proc
pivots three attributes by self-joining `vehicledata`/`DataItem` once each:

| Block | BC formula | VEHICLES | ORDERS | filter |
|---|---|---|---|---|
| Ground | `CONVERT(char(3), ModelYearCode + GROUNDWHEEL + GROUNDTIRE)` | `COUNT(*)` | `COUNT(*) * 4` | — |
| Spare | `CONVERT(char(3), ModelYearCode + SPARETIRE)` | `COUNT(*)` | `COUNT(*)` (1:1) | `SPARETIRE <> 'M'` |

- **Concat order = ModelYear, then WHEEL, then TIRE** for ground. The alias numbering is the *reverse*
  of the read order: in `AD_FRSPULL-shared.sql:46/57/62`, `vd2` = GROUNDWHEEL and `vd1` = GROUNDTIRE,
  and the expression is `ModelYearCode + vd2 + vd1` = MY + WHEEL + TIRE. Do not assume vd1=first
  attribute. (`AD_FRSPULL-analysis §2/§7.1`)
- **`ORDERS = VEHICLES×4` is the 4-corners assumption (ground only).** Spare is one unit → spare
  ORDERS = VEHICLES. **ORDERS is NOT a quantity** — it is purely the gate for the downstream
  "(No Ratio)" small-volume branch (`ORDERS <= 5`). `VEHICLES` is the qty multiplier.
  (`AD_FRSPULL-analysis §2/§3.2`)

### 2.3 char(3) padding semantics (LIKE-significant, but strip-safe — see §4)

- **Ground BC = 3 real chars** (e.g. `NBB`). **Spare BC = 2 chars + ONE trailing space** (e.g. `NN `).
  Proven on live proc output (`AD_FRSPULL-analysis §3.1`): `DATALENGTH('NN ')=3, LEN=2`.
- The trailing space is the **namespace separator** between ground (3-char) and spare (2-char+space)
  codes and feeds the downstream `LIKE`. Today the namespaces are structurally disjoint (zero blank
  ground values in-window, `AD_FRSPULL-analysis §3.4`), so `UNION` removes 0 rows and behaves like
  `UNION ALL`.
- **The rebuild strips the space (`bc.strip()`, code.py:278) and this is PROVEN SAFE** —
  see §4 GAP-2-RESOLVED. T-SQL `LIKE` ignores trailing spaces on the *left (matched)* operand, so
  padded vs stripped BC return byte-identical forecast matches (`adversary-findings.md` RESOLUTION).

### 2.4 char(3) silent-truncation latent risk (safe today, must alarm)

`CONVERT(char(3), …)` **silently truncates** to 3 chars with no error. It is safe **only because
every component is exactly 1 char in current data** (GROUNDTIRE/WHEEL/SPARETIRE
`MAX(LEN(DataValue))=1`, `AD_FRSPULL-analysis §3.1`). If GALC ever emits a 2-char wheel/tire value,
the ground BC becomes 4 chars and is chopped to 3 — a **wrong BC that still looks valid** and matches
the wrong (or no) recipe. `AD_FRSPULL-analysis §7.3` calls this "the single highest-severity latent
defect." The rebuild reproduces the truncation faithfully (proc-resident) but **has no >1-char alarm**
(`adversary-findings.md` SHOULD-FIX-3) — see §6 NIT.

### 2.5 Spare `<> 'M'` exclusion (load-bearing on historical data)

`M` = "vehicle with no spare tire" ("Quick fix for no spare broadcast"). The filter is on the spare
block only — a vehicle whose spare is 'M' still contributes its **ground** BC, it just produces no
spare BC. It removes **0 rows in current production** but **2–9 spare rows/month historically**
(2018–2020). The rebuild MUST keep it. (`AD_FRSPULL-analysis §3.3/§7.5`) **This filter is one of the
three facts that differ between the live and the stale dump — see the version-conflict box.**

### 2.6 Other live-verified semantics the rebuild inherits via wrap-the-proc

- **Inclusive-both-ends date window + the midnight trap.** `DateCreated >= @begindate AND <= @enddate`;
  `DateCreated` carries a real time-of-day. `@enddate = '<day> 00:00:00'` returns **ZERO rows**;
  adjacent windows sharing an endpoint **double-count** the boundary instant. (`AD_FRSPULL-analysis §3.8`)
- **One vehicledata row per vehicle per attribute is convention, not enforced** (heap, no unique key).
  No fan-out today across 5 days + a heavy historical month; a duplicate write would silently double
  VEHICLES/ORDERS for that BC. (`§3.7`)
- **NULL in any BC component → NULL BC**, routed to a single non-matching bucket (silent no-line).
  No NULLs in-window today. (`§3.6`)
- **CI collation** on the GROUP BY / BC key; **`NOLOCK`** on every table (dirty-read perturbation
  possible mid-broadcast — flagged but unmentioned by the analyses; `design-review §S7`).

### 2.7 `@Start` / `@Last` are inert

The body never references them; identical output with absurd values (`§3.9`). This corrects the M1
spec's inference that S/E were an ASN sequence range. The rebuild passes them for signature fidelity
only (code.py:274).

### 2.8 Worked example (real line, 1-day window)

`EXEC AD_FRSPULL @begindate='2026-06-18 00:00:00', @enddate='2026-06-18 23:59:59.997', @LineName='COROLLA'`
(819 COROLLA vehicles built that day; `AD_FRSPULL-analysis §5`):

| BC | kind | VEHICLES | ORDERS | decode |
|---|---|---|---|---|
| `NBB` | ground | 507 | 2028 (=507×4) | MY=N, WHEEL=B, TIRE=B |
| `NEE` | ground | 194 | 776 (=194×4) | MY=N, WHEEL=E, TIRE=E |
| `PEE` | ground | 1 | 4 (=1×4) | MY=P, WHEEL=E, TIRE=E **← No-Ratio (4≤5)** |
| `NN ` | spare | 798 | 798 (1:1) | MY=N, SPARE=N |
| `NP ` | spare | 20 | 20 (1:1) | MY=N, SPARE=P |
| `PN ` | spare | 1 | 1 (1:1) | MY=P, SPARE=N **← No-Ratio (1≤5)** |

Block totals reconcile: ground ΣVEHICLES = spare ΣVEHICLES = 819 = vehicles built (each vehicle once
per block). The ×4 makes a **ground** BC trip No-Ratio at ≈1 vehicle (1×4=4≤5; 2→8>5) while a
**spare** BC trips at ≤5 vehicles — a real behavioral coupling the rebuild preserves by passing ORDERS
through unchanged. For `NBB` (VEHICLES=507): each matched recipe row yields
`banker_round(507 × IN_ASSY_QTY × IN_TIRE_RATIO / 100)`.

### 2.9 VERSION-CONFLICT — the live char(3) proc is authoritative; the on-disk dump is quarantined

> **DOCUMENTED FINDING (adjudicated by the orchestrator against the live restored backup).**
> There are **two different `AD_FRSPULL` versions**. They are NOT the same proc.
>
> | Fact | **LIVE / AUTHORITATIVE** (`AD_FRSPULL-shared.sql`; what the rebuild reads) | **STALE / QUARANTINED** (`/Users/apple/Documents/FP docs/SQL/VehicleOrder.sql`, Script Date 06/10/2026, lines 4714–4767) |
> |---|---|---|
> | BC width | `convert(char(3), …)` | `convert(char(21), …)` (verified lines 4722/4742) |
> | Ground concat | `MY + WHEEL(vd2) + TIRE(vd1)` | `MY + TIRE(vd1) + WHEEL(vd2)` — **REVERSED** (vd1=GROUNDTIRE line 4733, vd2=GROUNDWHEEL line 4738) |
> | Spare `<> 'M'` | **present** | **ABSENT**; spare wheel join `/*+vd2.DataValue*/` commented out |
>
> **Resolution:** "live wins over dump." `adversary-findings.md` adjudicated this DEFINITIVELY against
> the running proc on the restored backup:
> `sys.dm_exec_describe_first_result_set` → **`BC char(3)`**; `OBJECT_DEFINITION` → **two
> `convert(char(3),…)`, zero `char(21)`**; spare BCs come back as `'NN '` (2 chars + ONE space) and the
> `<> 'M'` exclusion fires on historical data. The downstream column `VC_BROADCAST_CODE` and param
> `@BCode` are both `varchar(20)` — there is **no `char(21)` anywhere in the live chain**.
>
> **Therefore `design-review-findings.md` "B1" (and its echo in `delphi-fanout-confirmation.md §f`) is
> REFUTED.** The arch-reviewer read the stale 06/10 dump and concluded "the rebuild encodes the wrong
> side / char(21), TIRE+WHEEL, no `<> 'M'`." That is wrong: the rebuild encodes the **LIVE** side
> correctly (it reads the live char(3) proc verbatim and never re-derives the BC formula). The stale
> dump must be treated as a foreign artifact — **never analyze the keystone from it**.
>
> **David confirmation still needed (§7 item 1):** formally confirm the live-backup `AD_FRSPULL`
> (char(3) / WHEEL+TIRE / `<> 'M'`) is the **production-canonical** runtime proc, and quarantine the
> 06/10 dump so no future analysis regresses to it.

---

## 3. Layer-by-layer — Delphi → SQL proc → rebuild

### PART A — `AD_FRSPULL` (the BC→numbers pull)

| | |
|---|---|
| **Delphi** | `CalculateASNFRS` calls it once per (line, window) on `ALC_StoredProc` (`DataModule.pas:5125`); reads `BC`, `Orders`, `VEHICLES` per row; the BC is read `FieldByName('BC').AsString` **without TRIM** (`adversary-findings.md` legacy cross-check, `DataModule.pas:5152`). |
| **Proc** | §2 above. char(3); ground ×4 / spare ×1; spare `<> 'M'`; inclusive window. |
| **Rebuild** | `create_asn` step 2 (code.py:272-279): `runPrepQuery("EXEC AD_FRSPULL @begindate=?, @enddate=?, @Start=?, @Last=?, @LineName=?", …, alcDb)`; per row `bc.strip()`, keep `ORDERS`/`VEHICLES` as ints. **Wrap-the-proc**: the BC formula is never re-derived in Ignition, so concat order / spare filter / ×4 can't drift. |

### PART B — `SELECT_ForecastDetailBCASN` (BC→parts/ratios)

| | |
|---|---|
| **Delphi** | per BC, `SELECT_ForecastDetailBCASN(@BCode, @EffMonth)` (`DataModule.pas:5149`); `@EffMonth = yyyy/MM` from prodDate (`:5154`); guards `recordcount > 0` (else "Missing Broadcast Code" abort, `:5273`). |
| **Proc** | `SELECT * FROM INV_FORECAST_DETAIL_INF LEFT JOIN INV_MANIFEST_COST_MST WHERE @BCode LIKE VC_BROADCAST_CODE AND ((VC_EFFECTIVE_MONTH=@EffMonth OR ='') AND IN_TIRE_RATIO<>0 AND IN_WHEEL_RATIO<>0)`. **The column is the LIKE pattern** (`[KLM]CC`), inbound BC is the literal left operand, CI collation (`SELECT_ForecastDetailBCASN-analysis §1a`). `@EffMonth` is **dead** (all eff months = `' '`, `' '=''` is TRUE, `§1c`). Only **tire ratio** drives qty; wheel ratio is a redundant copy used only in the both-100 gate (`§2c`). Zero rows = skip the BC (`§1d`). **No ORDER BY over two heaps** → nondeterministic row order (`§3`). |
| **Rebuild** | `create_asn` step 3 (code.py:288-300): same `EXEC` with the stripped BC as `@BCode`; explicit projection of the 6 columns Delphi reads (avoids the `SELECT *` duplicate-column trap, `§2e`). LIKE direction preserved by leaving the proc unchanged. **Does NOT add the prescribed `ORDER BY ID_FORECAST_DETAIL`** — see §6 SHOULD-FIX. |

### PART C — `CalculateASNFRS` fan-out → `computeAsnDetails` (the qty/manifest math)

| | |
|---|---|
| **Delphi** | `DataModule.pas:5180-5268`. Per BC: PRE-loop cost guard, then branch on `Orders<=5`. **No-Ratio** (`:5183`): one row from the FIRST forecast row, `qty = VEHICLES*IN_ASSY_QTY` (no ratio), `break`. **Ratio** (`:5214`): one row per forecast row; both ratios 100 → `qty = VEHICLES*IN_ASSY_QTY`; else `qty = round(VEHICLES*IN_ASSY_QTY*IN_TIRE_RATIO/100)`. Delphi `round` = **banker's** (round-half-to-even). |
| **Manifest** | `'7' + copy(prodDate,4,5) + VC_ASSY_MANIFEST_NUMBER` (`:5186`/`:5239`) = `'7'` + **1-digit year** + MM + DD + 2-char assy id (8 chars). E.g. `20260618` + id `57` → `76061857`. (`delphi-fanout §c`) |
| **Rebuild** | `computeAsnDetails` (code.py:85-179): branch on `orders<=5` → `fcRows[0]`, `qty = vehicles*IN_ASSY_QTY`, `continue` (the Pascal `break`); else per row both-100 full-qty or `_bankers_div_round(base*tire, 100)`. **Banker's rounding is implemented in exact integer arithmetic** (`_bankers_div_round`, code.py:45-61) to be correct on BOTH Jython 2.7 (whose builtin `round()` is half-away) AND Python 3. `_manifest` (code.py:64-75) = `"7" + prodDate[3:8] + assyId`, with a length-8 assertion. PRE-loop cost guard (code.py:139-146): scans ALL fc rows, raises `AsnFanoutError` if any `IN_MANIFEST_COST_ID is None`, **before the transaction opens** (so an abort writes nothing). |

### PART D — the write chain (`INSERT_ASNInfo` / `INSERT_ASNDetail`) → `create_asn`

| | |
|---|---|
| **`INSERT_ASNInfo`** | Header first (its `SCOPE_IDENTITY()` is the `IN_ASN_ID` every detail writes). Status hard-coded `'C'`; OUTPUT `@ASNID = SCOPE_IDENTITY()` (NOT `@@IDENTITY`/`IDENT_CURRENT`); 16-char `VC_ADD`/`VC_LAST_UPDATE` stamp (`yyyymmddHHmmss` + first 2 ms digits — **16, not 14**, `asn-write-chain §1`). Rebuild: code.py:318-328, captures the OUTPUT via `DECLARE @id; EXEC … @ASNID=@id OUTPUT; SELECT @id` on the SAME open tx through `runScalarPrepQuery(…, tx)` (a real 8.1+ API). `@Ein` passed as **0** (at-send EIN, PART E). |
| **`INSERT_ASNDetail`** | `@HotCall=0` accumulate (`IN_QTY += @Qty`), `@HotCall=1` always-insert. **Legacy keys on `VC_MANIFEST_NUMBER` ALONE** → a later ASN reusing a manifest accumulates into the *first* ASN's row (cross-ASN collision; real spike data has manifests in 3 distinct ASNs, `asn-write-chain §2d`). **Rebuild Q1 re-key to `(IN_ASN_ID, VC_MANIFEST_NUMBER)`** (`spike-asndetail-rekey.sql`; supporting composite index added) preserves the within-ASN accumulate (one manifest summed from several recipe rows) and removes the cross-ASN collision. Rebuild driver: code.py:333-335. |

### PART E — orchestration deltas the rebuild introduces (intended divergences)

| Delta | Legacy | Rebuild | Why |
|---|---|---|---|
| Transaction | Inv_Connection only; ALC steps outside → partial-ASN + EIN-gap | **One** Gateway txn wraps header + all details (code.py:309-343); pure fan-out (incl. PRE-loop abort) runs before the tx opens | All-or-nothing; no orphan partial ASN |
| EIN | `SiteEIN+1` stamped at create; `AD_UpdateEIN` (`UPDATE Site SET SiteEIN+1`, **no WHERE**) bumped outside the txn | `IN_ASN_EIN = 0` at create; real per-site EIN allocated atomically **at SEND** (M1 Rank 2) | Removes the EIN-gap, the read-then-bump race, and the unscoped-bump multi-site BLOCKER |
| Idempotency | UI form-lock (single-user desktop, structurally serial) | `SELECT_ASNSeq` read-back → no-op skip (code.py:263-267) | Headless driver; same intent — **but now concurrency-exposed, see §6 BLOCKER** |
| Post-loop audit | `try/except` swallow, log only | `SELECT_ASNMissingCost` after commit, warn-only, try/except (code.py:347-358) | Faithful — never aborts a committed ASN |

---

## 4. The equivalence table (corrected per the adjudication)

Status: ✅ faithful · ✅* faithful-with-note (a previously-open GAP now resolved) · ⚠️ intended
divergence · ❌ open GAP/risk. **GAP #1 (char-width) and GAP #2 (`bc.strip`) flip from ❌ to ✅\* per
the orchestrator's live-DB adjudication.**

| # | Legacy behavior | Rebuild location | Status |
|---|---|---|---|
| 1 | BC composition (MY+WHEEL+TIRE ground / MY+SPARE spare, `<> 'M'`) | Read verbatim from live `AD_FRSPULL` (code.py:272-279); never re-derived | ✅ |
| 2 | char(3) trailing-space padding (spare `'NN '`) | `bc.strip()` (code.py:278) | **✅\*** — **GAP RESOLVED.** T-SQL `LIKE` ignores trailing spaces on the left operand: `'NN ' LIKE '[MNP][N]'` and `'NN' LIKE '[MNP][N]'` both MATCH; `'NNX'` does NOT (real char ≠ space → strip ≠ chop). Padded vs stripped give byte-identical matches end-to-end on both DBs (`adversary-findings.md` proof a/b). Ground/spare keys stay disjoint (3 vs 2 chars), `collisions_after_strip=0`. NOT a defect today. |
| 3 | char(3) silent truncation | Proc-resident (faithful); **no >1-char alarm in rebuild** | ❌ NIT (latent; safe today, alarm absent — §6) |
| 4 | ORDERS=×4 ground/×1 spare; VEHICLES=count | Passed through unchanged; ORDERS only gates `<=5`, VEHICLES only multiplies (code.py:148-166) | ✅ |
| 5 | BC→forecast `@BCode LIKE VC_BROADCAST_CODE` (column=pattern), CI | Proc unchanged; stripped BC as left operand (code.py:289) | **✅\*** — **char(3)-vs-char(21) source conflict RESOLVED: char(3) is live-authoritative** (§2.9). LIKE direction + CI preserved. The strip is safe (#2). |
| 6 | `SELECT_ForecastDetailBCASN` nondeterministic order → No-Ratio `fcRows[0]` pick | `fcRows[0]` in heap-scan order; **no ORDER BY added** | ❌ SHOULD-FIX — **fires on live data** (BC `PEE`→m36-vs-m37, `adversary-findings.md` SHOULD-FIX-1) |
| 7 | No-Ratio: first fc row, `qty=VEHICLES*IN_ASSY_QTY`, break | code.py:148-160 | ✅ (modulo #6) |
| 8 | Ratio qty = banker_round(VEHICLES*IN_ASSY_QTY*IN_TIRE_RATIO/100); both-100 full-qty; tire-only numerator | code.py:163-177, `_bankers_div_round` | ✅ highest-fidelity item |
| 9 | Manifest `'7'+1-digit-yr+MM+DD+2-char id` | `_manifest` code.py:64-75 | ✅ |
| 10 | `INSERT_ASNInfo` status 'C'; SCOPE_IDENTITY OUTPUT | code.py:318-328 | ✅ (status/identity); EIN → #11 |
| 11 | EIN `fEIN+1` at create + bump | `IN_ASN_EIN=0` at create; real EIN at SEND | ⚠️ intended divergence — harness must **exclude `IN_ASN_EIN`** from parity diff |
| 12 | `INSERT_ASNDetail` accumulate; Q1 re-key `(IN_ASN_ID, manifest)` | re-keyed proc + composite index; code.py:333-335 | ✅ (single-site; site_id half deferred to M4) |
| 13 | Positional VALUES / IDENTITY-skip fragility | re-key kept positional VALUES (M4 markers) | ⚠️→❌ latent — M4 `IN_SITE_ID` add will break the positional INSERT unless `@Site` prepend lands atomically |
| 14 | `SELECT_ASNSeq` idempotency = UI dedup, **no DB unique constraint** | read-back → no-op skip; no constraint | ❌ **BLOCKER under concurrency** (§6) |
| 15 | TWO cost checks (PRE hard-abort on NULL cost id, not date-aware; POST date-windowed warn) | PRE in `computeAsnDetails` pre-tx (code.py:139-146); POST after commit warn-only (code.py:347-358) | ✅ both implemented + correctly distinguished |
| 16 | No surrounding transaction (partial-ASN) | one Gateway txn; abort writes nothing | ⚠️ intended divergence (a fix) — failure-path is NOT byte-parity |
| 17 | 16-char audit stamp (not 14) | proc emits 16; re-key emits 16 | ✅ value faithful; ⚠️ stale "14-char" comments (§6 NIT) |

**Genuinely-open items (kept):** #3 (truncation alarm), #6 (No-Ratio nondeterminism — firing),
#13 (M4 positional INSERT), #14 (concurrency constraint). Intended divergences needing sign-off:
#11 (EIN-at-send), #16 (single-transaction).

---

## 5. Verification status — honest (fixture-fidelity discipline)

**PROVEN equivalent (the MATH, on the live char(3) basis)** — `adversary-findings.md`
"WHAT IS PROVEN EQUIVALENT", verified against the live running procs/data on `mssql-spike`:

- **Ratio-branch qty** (the revenue-critical path): `qty = VEHICLES*IN_ASSY_QTY` when both ratios=100,
  else `banker_round(VEHICLES*IN_ASSY_QTY*IN_TIRE_RATIO/100)` — tire-only numerator, literal-100
  denominator, both-100 gate reads both ratios. Verified live (NEE V=199 → m36=557/m37=239; NBB V=526
  → 842/421/842). Matches `DataModule.pas:5226-5235`.
- **No-Ratio branch first-row + break**, `qty=VEHICLES*IN_ASSY_QTY` (`DataModule.pas:5183-5212`),
  modulo the SHOULD-FIX-1 order hazard.
- **Banker's rounding** matching Delphi `Round` (half-to-even), correct on Jython 2.7 and Python 3
  (exact integer arithmetic). Unit-tested at exact halves; diverges from T-SQL half-away by design.
- **Manifest scheme** `'7'+1-digit-yr+MM+DD+id`.
- **BC→forecast LIKE handoff incl. `.strip()`** — padded vs stripped byte-identical on both DBs; LIKE
  direction + CI preserved.
- **BOTH cost guards** — PRE-loop NULL-cost-id hard-abort (pre-tx); POST-loop date-windowed
  `SELECT_ASNMissingCost` warn (incl. the "out of date" EXPIRED-cost case, proven live on
  `42670FET9000`, window expired vs prodDate 20260618 → rebuild WARNs and prices the line, does not
  abort/drop).
- **Header write** (status 'C', SCOPE_IDENTITY) and the re-keyed accumulate upsert.

**NOT proven — row-for-row legacy parity (UNPROVABLE from available data).** Every reproducible
legacy ASN (4718–4721; 4722 is a hot-call/manual) was frozen under an **older forecast-recipe
vintage** in `Inventory_Live`'s history; the per-BC qtys imply mutually-inconsistent vehicle counts
under today's recipe (e.g. NBB legacy 80/900/1124 vs today's assy=4, tire 40/20/40). **No legacy ASN
built under the current recipe exists**, so there is no row-parity oracle. What the harness proves is
(A) driver self-consistency and (B) a total-qty match — and even the total is **±1-per-BC-residuals
cancel coincidental, not a conservation law** (`adversary-findings.md` SHOULD-FIX-2): for ASN 4721 the
grand total 4240==4240 only because NBB's +1 and NJJ's −1 cancel; `sum(round(base*ratio_i/100))` can
drift ±1 even when ratios sum to 100. The honest invariant is **per-BC ≈ base ± rounding residual**,
not grand-total-exact.

> **STATE PLAINLY:** do NOT read this as "ASN creation is certified correct." What is certified is
> total-qty-conserving + internally self-consistent qty/manifest math on one (live, authoritative) set
> of inputs. Manifest-level row correctness vs the legacy is **unproven by construction** until a true
> oracle exists (§7 item 5).

---

## 6. MUST-FIX before production sign-off

Real defects both adversaries found, on the live char(3) basis. Ranked by blast radius.

### BLOCKER — no DB unique constraint behind the idempotency guard (concurrency double-insert)

The only unique index on `INV_ASN_MST` is `PK_INV_ASN_MST` on the IDENTITY column; there is **no
unique constraint on `(VC_LINE_NAME, VC_PRODUCTION_DATE)`** (`adversary-findings.md` BLOCKER-1, via
`sys.index_columns`). `create_asn` does check-then-act with no serialization (code.py:263-267): two
concurrent gateway `create_asn` calls for the same (line, prodDate) both read 0 rows, both commit a
**full ASN** (header + complete fan-out) → **duplicate shipment**. **This risk is CREATED by the
rebuild** — the legacy was a single-user desktop with a UI form-lock, structurally serial and never
concurrency-exposed. The headless multi-session gateway inherits the non-atomic guard with nothing to
enforce it.
**Fix:** add a unique index on `(VC_LINE_NAME, VC_PRODUCTION_DATE[, IN_SITE_ID])` (with the
`VC_START_SEQ_NUMBER <> -1` exclusion of placeholder rows), and handle the unique-violation race
(catch → treat as the idempotent skip). Do not rely on the read-back guard alone.

### SHOULD-FIX — `SELECT_ForecastDetailBCASN` nondeterministic row order (No-Ratio pick is luck-of-allocation, and it FIRES)

Both base tables are heaps with no `ORDER BY` (`SELECT_ForecastDetailBCASN-analysis §3`). The No-Ratio
branch takes `fcRows[0]` in heap-scan order. **This is not theoretical — it fires on live data:** ASN
4721, BC `PEE` (ground, VEHICLES=1, ORDERS=4≤5) matches the 2-row pattern `[MNP]EE` →
`42600FEL1000`/m36 vs `42600FEL2000`/m37; the single-vehicle 4 units land on whichever the heap
returns first (today m36, stable across 6 reruns — but incidental allocation). A table reload / page
reuse / replan can pick a **different part number** silently, and the self-consistency test would still
pass. (`adversary-findings.md` SHOULD-FIX-1; `design-review §B3`.)
**Fix:** add deterministic `ORDER BY ID_FORECAST_DETAIL` to the recipe read, and get David's decision
on which assy the single-vehicle case is *supposed* to pick (§7 item 4).

### SHOULD-FIX — the parity test oversells the total-qty invariant

`test_create_asn_parity.py` sells the grand-total match (4240==4240) as a "REAL parity signal that
survives the vintage drift." It can pass by **±1 per-BC residual cancellation** (§5). The fan-out does
NOT conserve `VEHICLES*assy` per BC; it conserves it per-BC only `≈ base ± residual`.
**Fix:** soften the NOTE's language; if a total gate is kept, gate **per-BC** (or document that the
grand total can match by cancellation). The qty math itself is faithful — this is a test-method flaw,
not a code defect.

### NIT (low-severity, fix before sign-off review)

- **char(3) truncation alarm absent** (#3). Faithful-to-legacy today (both wrap the same proc), but the
  "assert/alarm if any BC component > 1 char" that `AD_FRSPULL-analysis §7.3` calls the highest-severity
  latent defect is not implemented. Add a width assertion on the BC components.
- **Stale "14-char" stamp comments.** The true value is **16** (`asn-write-chain §1`);
  `spike-asndetail-rekey.sql:25-26` and `m1-asn-creation-spec.md §2` still say 14. Doc-only; correct so
  the parity reviewer isn't misled.
- **The spec still calls `AD_FRSPULL` "the one true source gap / BLOCKER for parity, body NOT in any
  dump"** (`m1-asn-creation-spec.md §3/§9`). This is **dead and false** — the body is on disk
  (`AD_FRSPULL-shared.sql`) and the conflict is refuted (§2.9). Reconcile the spec's headline.

---

## 7. Open items + David sign-offs needed

> **✅ SIGN-OFF STATUS (David, 2026-06-20):** #1 CONFIRMED (live-backup `AD_FRSPULL` is production-canonical;
> quarantine the stale 06/10 dump). #2 EIN-at-send CONFIRMED — reader trace done: only `UPDATE_EINStatus`
> (ack, post-send) + `REPORT_EDI856` (the send) touch `IN_ASN_EIN`; `SELECT_ASNList` only DISPLAYS it (shows
> 0 pre-send — render as "—"); no break. #3 single-txn all-or-nothing CONFIRMED. #4 No-Ratio tiebreak
> CONFIRMED = **lowest `ID_FORECAST_DETAIL`** (faithful; reproduces ASN 4721's FEL1000/76061836) — already
> implemented in `code.py`. #5 true row-parity oracle: APPROACH ACCEPTED (certify total-qty + self-consistency
> now; a true oracle / the dev-mirror multi-week comparison comes later). The BLOCKER + nondeterminism are
> FIXED (PR #28). **The ASN-creation keystone is signed off** on the faithful-on-the-live-basis standard.

1. **Confirm the live-backup `AD_FRSPULL` is production-canonical** (char(3) / WHEEL+TIRE / `<> 'M'`)
   and **quarantine the stale 06/10 dump** (`/Users/apple/Documents/FP docs/SQL/VehicleOrder.sql`) so
   no future analysis regresses to the char(21)/reversed/no-`M` version. (§2.9; refutes design-review B1.)
2. **EIN-at-send** — accept EIN-at-send over legacy EIN-at-create **after a grep confirms no consumer
   reads `IN_ASN_EIN` between create and send** (a report, re-print, 856 regen, or between-create-and-send
   query would see 0). The "nothing interprets `IN_ASN_EIN`" claim was proven only for the four write
   procs, not for all consumers (`design-review §S1`). David knows the floor workflow.
3. **Single-transaction all-or-nothing** — accept over legacy partial-ASN persistence, confirming **no
   recovery workflow depends on a half-written ASN surviving a mid-loop abort** (`design-review §S2`).
4. **No-Ratio deterministic assy pick** — which assy the `Orders<=5` "first row" should pick once an
   `ORDER BY ID_FORECAST_DETAIL` is added (the legacy result may itself have been arbitrary;
   `SELECT_ForecastDetailBCASN-analysis §3`). Blocks SHOULD-FIX (No-Ratio nondeterminism).
5. **A true row-parity oracle** — none exists today (every reproducible legacy ASN is an older
   recipe-vintage; §5). Either create a legacy ASN under the current forecast-recipe vintage and gate
   `rebuilt == legacy`, or freeze a contemporaneous recipe snapshot to replay an old ASN row-for-row.
   Until one exists, sign-off can only certify total-qty-conservation + self-consistency, **not
   manifest-level correctness**.

**Lower-priority (post sign-off gate):** design+test the at-send `INV_SITES.IN_EIN_SEQ` allocation
(named only in a docstring, untested — `design-review §S4`); tag the multi-site `site_id` half as
deferred (it is two-thirds done — `§S5`); guard the cross-DB `AD_FRSPULL` short-read (a flaky ALC
connection ships an under-counted ASN with no abort — `§S6`); note the `NOLOCK` dirty-read perturbation
(`§S7`).

---

## Overall honest correctness verdict for this keystone

**The rebuild's per-row qty / manifest MATH is PROVEN equivalent to the legacy on the live
(authoritative) char(3) basis** — ratio branch, both-100 full-qty, No-Ratio first-row+break, banker's
rounding matching Delphi `Round`, the manifest scheme, the BC→forecast `LIKE` handoff including the
`.strip()`, and BOTH cost guards (incl. the post-loop EXPIRED-cost warn). The two coupled fears
(char-width and `bc.strip()`) are **RESOLVED**: char width is **char(3)** (the char(21) claim came from
a stale dump and is refuted), and `bc.strip()` is **safe** (T-SQL `LIKE` trims the left operand;
ground/spare namespaces stay disjoint).

**But this keystone is NOT yet production-equivalent.** Three things gate sign-off: the **concurrency
BLOCKER** (no DB unique constraint — a risk the rebuild itself introduces by moving to a multi-session
gateway), the **No-Ratio nondeterminism** (already firing on live data, no `ORDER BY`, David's pick
open), and the **absence of a row-parity oracle** (recipe-vintage drift makes legacy row-for-row parity
unprovable). What can be signed off today is "total-qty-conserving and internally self-consistent" —
which is **not** the same as "ships the right manifests." Close §6 BLOCKER + SHOULD-FIXes and stand up
an oracle (§7.5) before treating this as certified.
