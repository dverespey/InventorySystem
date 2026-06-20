# ASN-create write + guard chain — authoritative T-SQL behavioral analysis

**Scope:** the four procs that *write* an ASN and the two cost guards around them, in DB `Inventory`:
`INSERT_ASNInfo`, `INSERT_ASNDetail` (Q1-re-keyed on the spike), `SELECT_ASNSeq`, `SELECT_ASNMissingCost`.
**Sibling analysis (the BC→parts fan-out + pre-loop cost abort):**
`docs/analysis/production-readiness/sql/SELECT_ForecastDetailBCASN-analysis.md`.
**Re-key artifact applied to spike `Inventory`:** `docs/analysis/edi/spike-asndetail-rekey.sql`.

**Verification.** Every claim below was proven on the live `mssql-spike` container against DB `Inventory`
(rebuild target, re-key applied) and `Inventory_Live` (legacy, READ-ONLY). All destructive probes ran inside
`BEGIN TRAN … ROLLBACK` on `Inventory`; each ended with a zero-row "after rollback" check (shown). Proc bodies
pulled live via `OBJECT_DEFINITION` from both DBs; legacy dump cross-checked at
`DB Schema/CreateInventory.sql` (UTF-16LE). Verified-vs-inferred stated per item.

---

## 0. The write-chain order (what executes, in what order)

Caller: `ASNSelect.pas` (the ASN UI). Two phases, header-first by necessity (the detail rows need the
header's IDENTITY).

```
ASNSelect form load
  └─ LoadSeqNumbers           ASNSelect.pas:137  → SELECT_ASNSeq   (idempotency GUARD, §3)
                                                    if a row exists → lock fields + DISABLE Create buttons

User clicks Create
  └─ Data_Module.InsertASNInfo               DataModule.pas:5321
       ├─ EXEC INSERT_ASNInfo (OUTPUT @ASNID) DataModule.pas:5337   → header row, status 'C', SCOPE_IDENTITY
       │       fRecordID := @ASNID            DataModule.pas:5364   (the new IN_ASN_ID, threaded into details)
       └─ CalculateASNFRS                     DataModule.pas:5106 / called :5381
            ├─ EXEC AD_FRSPull (ALC DB)       DataModule.pas:5125   → one row per broadcast code (BC)
            └─ for each BC row:
                 ├─ EXEC SELECT_ForecastDetailBCASN  :5149         → recipe rows for the BC
                 ├─ if 0 rows → RAISE "Missing Broadcast Code …"  :5273  (hard abort)
                 ├─ PRE-LOOP COST GUARD (hard abort, §4a):        :5160-5175
                 │     walk ALL recipe rows; if ANY IN_MANIFEST_COST_ID IS NULL → RAISE, ASN create fails
                 └─ for each recipe row:
                      └─ EXEC INSERT_ASNDetail (@ASNID=fRecordID) :5191 / :5243  → detail upsert (§2)
            └─ after ALL BCs done:
                 └─ POST-LOOP COST WARN (does NOT abort, §4b):    :5285-5308
                      EXEC SELECT_ASNMissingCost @ASNID=fRecordID → date-windowed warn; ShowMessage + log only
```

**Load-bearing ordering facts**
- `INSERT_ASNInfo` MUST run first: its `SCOPE_IDENTITY()` is the `IN_ASN_ID` every `INSERT_ASNDetail` writes
  into `INV_ASN_DETAIL_MST.IN_ASN_ID` (and, post-re-key, the *key* the detail upsert dedups on).
- The header is committed **before** the detail loop; the detail loop can abort (pre-loop cost RAISE, missing
  BC RAISE) **after** the header exists. There is **no surrounding transaction** in the Delphi — each `ExecProc`
  is its own auto-commit. So a pre-loop cost abort on BC #3 leaves the header + BC #1/#2 detail rows persisted
  (orphan-ish partial ASN). A faithful rebuild that wraps the whole create in one transaction would change this
  failure semantics — call it out before "improving" it.
- The post-loop `SELECT_ASNMissingCost` runs only after the loop completes normally; it is wrapped in its own
  Delphi `try/except` that **swallows** any error (`DataModule.pas:5303-5307`) — it can never fail the create.

---

## 1. INSERT_ASNInfo — the header insert (VERIFIED; no drift)

**Param signature (catalog-verified):**

| # | Param | Type | Dir |
|---|---|---|---|
| 1 | `@ASNID` | int | **OUTPUT** |
| 2 | `@LineName` | varchar(50) | in |
| 3 | `@AssyLine` | varchar(1) | in |
| 4 | `@StartSeq` | varchar(4) | in |
| 5 | `@DTStartSeq` | datetime | in |
| 6 | `@EndSeq` | varchar(4) | in |
| 7 | `@DTEndSeq` | datetime | in |
| 8 | `@Qty` | int | in |
| 9 | `@PDate` | varchar(8) | in |
| 10 | `@Ein` | int | in |

**Body (live `Inventory`, identical to `Inventory_Live` and to the dump):**

```sql
SET @Now = getdate()
SET @AddDate = CONVERT(char(8), @Now, 112) + SUBSTRING(CONVERT(varchar,@Now,114),1,2)
             + SUBSTRING(CONVERT(varchar,@Now,114),4,2) + SUBSTRING(CONVERT(varchar,@Now,114),7,2)
             + SUBSTRING(CONVERT(varchar,@Now,114),10,2)
INSERT INTO INV_ASN_MST
VALUES( @Ein, 'C', @LineName, @AssyLine, @StartSeq, @DTStartSeq, @EndSeq, @DTEndSeq, @Qty, @PDate, @AddDate, @AddDate)
SET @ASNID = SCOPE_IDENTITY()
```

**Columns it sets (positional VALUES — IDENTITY-skip).** `INV_ASN_MST` has 13 columns; col 1
`IN_ASN_ID` is IDENTITY, so the 12-value list maps cols 2–13 in order:

| col | column | value | note |
|---|---|---|---|
| 2 | `IN_ASN_EIN` | `@Ein` | EIN placeholder — see below |
| 3 | `VC_ASN_STATUS` | `'C'` | **hard-coded literal** (Created) |
| 4 | `VC_LINE_NAME` | `@LineName` | |
| 5 | `VC_ASSEMBLY_LINE` | `@AssyLine` | varchar(1) |
| 6 | `VC_START_SEQ_NUMBER` | `@StartSeq` | varchar(4) |
| 7 | `DT_START_SEQ` | `@DTStartSeq` | |
| 8 | `VC_END_SEQ_NUMBER` | `@EndSeq` | varchar(4) |
| 9 | `DT_END_SEQ` | `@DTEndSeq` | |
| 10 | `IN_QTY` | `@Qty` | |
| 11 | `VC_PRODUCTION_DATE` | `@PDate` | varchar(8) yyyymmdd |
| 12 | `VC_LAST_UPDATE` | `@AddDate` | the stamp |
| 13 | `VC_ADD` | `@AddDate` | same stamp |

**Status `'C'` is hard-coded** — there is no parameter for it; every freshly created ASN header is status `C`.
(VERIFIED — live header IN_ASN_ID=4745 created in the §4 probe came back `VC_ASN_STATUS='C'`.)

**`@ASNID` / SCOPE_IDENTITY mechanism (VERIFIED).** `SET @ASNID = SCOPE_IDENTITY()` returns the IDENTITY value
generated by *this* INSERT in *this* scope (not `@@IDENTITY`, which would also see any trigger-generated
identity on another table). Delphi reads it back as the `pdOutput` param and stores it in `fRecordID`
(`DataModule.pas:5340/5364`). **Trap:** a faithful rebuild MUST use `SCOPE_IDENTITY()` / the JDBC
generated-keys API, not `@@IDENTITY` and not `IDENT_CURRENT('INV_ASN_MST')` (race under concurrency). There is
no INSERT trigger on `INV_ASN_MST` today, so `@@IDENTITY` would agree by accident — do not rely on that.

**EIN handling (`IN_ASN_EIN`).**
- **Legacy + current Delphi:** the caller passes `@Ein := fEIN+1` (`DataModule.pas:5359`) — the *next* EDI
  Interchange Control Number, bumped at create time. The same `fEIN+1` is also passed into every
  `INSERT_ASNDetail` (`@EIN`, `:5196/:5248`) so header and detail carry a matching EIN.
- **Rebuild decision (at-SEND EIN):** the rebuild passes **`@Ein = 0` at create**; the real EIN is assigned
  when the EDI is actually transmitted. **`0` is the create-time placeholder** in `IN_ASN_EIN` (and
  `INV_ASN_DETAIL_MST.IN_ASN_EIN`). (VERIFIED — the §4 probe created the header with `@Ein=0` and read back
  `IN_ASN_EIN=0`.) The column itself is `int NOT NULL`, so `0` is a sentinel, not NULL.
  - *Faithful-reimplementation note:* nothing in these four procs interprets `IN_ASN_EIN` — they only store it.
    The "0 means not-yet-sent" meaning lives entirely in the send path. The rebuild must (a) write 0 here at
    create and (b) stamp the real EIN at send (the old `fEIN+1` bump moves to the send step).

**The `@AddDate` stamp is 16 chars, not 14 (VERIFIED — correction to the re-key file's comment).**
`CONVERT(char(8),@Now,112)` = `yyyymmdd` (8); style 114 = `hh:mi:ss:mmm` with a **colon** before ms; the four
2-char substrings pull HH(1-2), mm(4-5), ss(7-8), and **the first 2 of the 3 millisecond digits (10-11)**:

```
$ DECLARE @Now datetime='2026-06-15T09:23:33.673';
  SELECT CONVERT(varchar,@Now,114)                 → 09:23:33:673
  SELECT <the concatenation>, LEN(...)             → 2026061509233367 | 16
```

So `VC_ADD`/`VC_LAST_UPDATE` = `yyyymmddHHmmssNN` where `NN` = the first two **milliseconds** digits, NOT
seconds-then-blank. Live detail rows confirm 16-char stamps (`2026061109233367`, `LEN=16`). A rebuild that
emits a 14-char `yyyymmddHHmmss` would not byte-match the legacy stamp. (Cosmetic — nothing parses it back —
but it is the kind of silent diff a parity harness flags.)

---

## 2. INSERT_ASNDetail — the re-keyed upsert (DRIFT: Inventory ≠ Inventory_Live)

**Param signature (catalog-verified):** `@ASNID int, @EIN int, @Manifest varchar(8), @PartNumber varchar(12),
@Qty int, @HotCall bit = 0`. The re-key added **no new parameter** — `@ASNID` was already passed (it is
`fRecordID`, `DataModule.pas:5194/:5246`), and `IN_ASN_ID` is already a column on `INV_ASN_DETAIL_MST`.

### 2a. Legacy body (Inventory_Live) — keys on manifest ALONE (the bug)

```sql
if @HotCall = 0
begin
    SELECT * FROM INV_ASN_DETAIL_MST WHERE VC_MANIFEST_NUMBER = @Manifest    -- manifest ONLY
    if @@rowcount = 0
        INSERT INTO INV_ASN_DETAIL_MST VALUES(@ASNID, null, @EIN, @Manifest, @PartNumber, @Qty, @AddDate, @AddDate)
    else
        UPDATE INV_ASN_DETAIL_MST SET IN_QTY = IN_QTY + @Qty WHERE VC_MANIFEST_NUMBER = @Manifest  -- manifest ONLY
end
else  -- @HotCall=1
    INSERT INTO INV_ASN_DETAIL_MST VALUES(@ASNID, null, @EIN, @Manifest, @PartNumber, @Qty, @AddDate, @AddDate)
```

The legacy existence check + UPDATE are scoped to `VC_MANIFEST_NUMBER` **globally across all ASNs**. A later
ASN that reuses a manifest id accumulates into the *first* ASN's row (cross-ASN collision), and the legacy
companion `DELETE_ASNItem(@ManifestNumber)` deletes that manifest from *every* ASN. Within one create the
accumulate is *desired* (one manifest hit by several BC/parts rows must sum). The legacy `SELECT *` upsert
probe also leaks a result set (it `SELECT *`s rather than `IF EXISTS`), harmless to Delphi which ignores it.

### 2b. Re-keyed body (Inventory) — keys on `(IN_ASN_ID, VC_MANIFEST_NUMBER)`

```sql
IF @HotCall = 0
BEGIN
    IF EXISTS (SELECT 1 FROM INV_ASN_DETAIL_MST
               WHERE IN_ASN_ID = @ASNID AND VC_MANIFEST_NUMBER = @Manifest)        -- (ASN, manifest)
        UPDATE INV_ASN_DETAIL_MST SET IN_QTY = IN_QTY + @Qty
         WHERE IN_ASN_ID = @ASNID AND VC_MANIFEST_NUMBER = @Manifest;              -- (ASN, manifest)
    ELSE
        INSERT INTO INV_ASN_DETAIL_MST VALUES(@ASNID, null, @EIN, @Manifest, @PartNumber, @Qty, @AddDate, @AddDate);
END
ELSE
    INSERT INTO INV_ASN_DETAIL_MST VALUES(@ASNID, null, @EIN, @Manifest, @PartNumber, @Qty, @AddDate, @AddDate);
```

Preserved: `@HotCall=0` accumulate; `@HotCall=1` always-insert; positional VALUES; the 16-char stamp.
Changed: dedup scope is now per-ASN. Supporting index `IX_INV_ASN_DETAIL_MST_ASN_MANIFEST (IN_ASN_ID,
VC_MANIFEST_NUMBER)` was added (legacy `IX_INV_ASN_DETAIL_MST` on `VC_MANIFEST_NUMBER` alone remains).
(VERIFIED via `sys.indexes` — both indexes present on spike `Inventory`.)

### 2c. Positional VALUES vs the 9-column table (the IDENTITY-skip trap, VERIFIED)

`INV_ASN_DETAIL_MST` has **9** columns; col 1 `IN_ASN_DETAIL_ID` is IDENTITY. The `INSERT … VALUES` lists **8**
values and **no column list**, so SQL Server maps them positionally to cols 2–9:

| value | col | column |
|---|---|---|
| `@ASNID` | 2 | `IN_ASN_ID` |
| `null` | 3 | `IN_INV_ID` (nullable) |
| `@EIN` | 4 | `IN_ASN_EIN` |
| `@Manifest` | 5 | `VC_MANIFEST_NUMBER` |
| `@PartNumber` | 6 | `VC_ASSY_PART_NUMBER` |
| `@Qty` | 7 | `IN_QTY` |
| `@AddDate` | 8 | `VC_LAST_UPDATE` |
| `@AddDate` | 9 | `VC_ADD` (nullable) |

**Trap (fragile to schema change):** because there is no explicit column list and the leading IDENTITY is
skipped by position, *any* column added/reordered on `INV_ASN_DETAIL_MST` silently shifts the mapping or throws
a count mismatch. This is exactly the M4 multi-site hazard — adding `IN_SITE_ID` will break the positional
INSERT unless `@Site` is prepended in the VALUES (and the WHEREs gain `IN_SITE_ID = @Site AND`, the `-- M4:`
markers in `spike-asndetail-rekey.sql`). A faithful rebuild should convert to an **explicit column-list
INSERT** to make the IDENTITY-skip and column mapping non-positional.

### 2d. Re-key proof on spike `Inventory` (rolled-back tran)

Synthetic manifest `ZTEST999`, two REAL distinct existing headers `@asnA=1804`, `@asnB=4715`:

```
--- (a) ACCUMULATE WITHIN ONE ASN: 3 calls @asnA, Qty 10/25/7, HotCall=0
1804 | ZTEST999 | IN_QTY=42 | rows_for_man=1        ← one row, 10+25+7=42  ✔
--- (b) NO CROSS-ASN COLLISION: same manifest, @asnB, Qty 100, HotCall=0
1804 | ZTEST999 | PARTX | 42
4715 | ZTEST999 | PARTY | 100                         ← SEPARATE row in ASN B  ✔
   total_rows_for_man=2  total_qty=142
--- (c) HOTCALL=1 always-insert: 4th call @asnA, Qty 3, HotCall=1
1804 | rows_in_asn=2 | qty_in_asn=45                  ← NEW row (was 1 row/42) ✔
4715 | rows_in_asn=1 | qty_in_asn=100
--- ROLLED BACK → rows_after_rollback = 0
```

Contrast probe (same seed, comparing the two WHEREs directly):

```
--- LEGACY manifest-only check when inserting for ASN B:  WHERE VC_MANIFEST_NUMBER='ZTEST999'  → 1
        (>0 ⇒ legacy would UPDATE ASN A's row, accumulating B's qty into A — the collision)
--- RE-KEYED check for ASN B:  WHERE IN_ASN_ID=4715 AND VC_MANIFEST_NUMBER='ZTEST999' → 0
        (0 ⇒ inserts a new row for B — no collision)
--- ROLLED BACK → 0
```

**This is not a theoretical bug.** Real spike data already has manifests living in multiple distinct ASNs that
would have collided under the legacy key:

```
$ SELECT TOP 5 VC_MANIFEST_NUMBER, COUNT(DISTINCT IN_ASN_ID) FROM INV_ASN_DETAIL_MST
  GROUP BY VC_MANIFEST_NUMBER HAVING COUNT(DISTINCT IN_ASN_ID)>1 ORDER BY 2 DESC:
52081285 | 3      52066074 | 3      52038818 | 3      52068238 | 3      52062102 | 2
```

---

## 3. SELECT_ASNSeq — the idempotency / re-create guard (VERIFIED; no drift)

**Signature:** `@LineName varchar(50), @PDate varchar(8)`. **Body (identical in both DBs + dump):**

```sql
SELECT * FROM INV_ASN_MST
WHERE VC_PRODUCTION_DATE = @PDate
AND VC_LINE_NAME = @LineName
AND VC_START_SEQ_NUMBER <> -1
```

**Result.** It is `SELECT *` over `INV_ASN_MST` (all 13 columns). The caller (`ASNSelect.pas:164-178`) reads
the **S / E / Q** triplet plus the two dates:
`VC_START_SEQ_NUMBER` (S), `VC_END_SEQ_NUMBER` (E), `IN_QTY` (Q), `DT_START_SEQ`, `DT_END_SEQ`.
Live shape for an existing (COROLLA, 20260610):

```
4715|9057|A|COROLLA||0413|2026-06-09 23:37:29.487|0220|2026-06-11 00:29:17.323|808|20260610|2026061109563121|2026061109233291
       ^S=0413                                        ^E=0220                    ^Q=808 ^PDate
```

**Idempotency-guard semantics (the rebuild's "prevent re-create for a line+prodDate").** `SELECT_ASNSeq` is
called on form load (`LoadSeqNumbers`). If it returns **≥1 row**, an ASN already exists for that
`(VC_LINE_NAME, VC_PRODUCTION_DATE)` and the form (a) populates start/last/qty/dates from it, (b) sets those
edits **ReadOnly**, and (c) **disables** `Check_Button`, `CreateASN_Button`, `CreateASNEntries_Button`
(`ASNSelect.pas:180-182`) — so the user cannot create a second ASN for the same line+date. If it returns **0
rows**, the create buttons are enabled (`:198-200`). So the guard is a **UI-level dedup**, not a DB constraint:
there is no unique index on `(VC_LINE_NAME, VC_PRODUCTION_DATE)` enforcing it.

**Where the sequence RANGE originates (S/E).** Not from this proc — `SELECT_ASNSeq` only *reads back* an
already-created header. On a fresh (no-row) case the operator **types** the start/last seq numbers into the
edits (`ASNSelect.pas:344-345`, zero-padded to 4 via `format('%.4d', …)`), and the production date comes from
`ASN_DateTimePicker`. Those typed values flow into `INSERT_ASNInfo` `@StartSeq/@EndSeq`. The `IN_QTY` (Q) on
read-back is the value `INSERT_ASNInfo` stored (operator-entered `fQTY`, `ASNSelect.pas:382/450`). So the
"sequence range" is operator input bounded only by the truck-seq length (`fiTruckSeqLength`); the proc's job is
purely the existence read-back.

**Edge-case matrix (SELECT_ASNSeq):**

| edge | behavior | proof / note |
|---|---|---|
| `VC_START_SEQ_NUMBER <> -1` filter | excludes "sentinel/cancelled" headers (start seq = `-1`) from the guard, so they don't block re-create | (VERIFIED filter present) — note `-1` is an int literal compared to a `varchar(4)` column → **implicit varchar→int** conversion of the column for the comparison. With CI_AS collation and clean numeric data this is fine, but a non-numeric `VC_START_SEQ_NUMBER` would raise a conversion error. The rebuild should compare as a string (`<> '-1'`) to avoid the implicit cast. The Delphi *write* side uses `'-1'` string (`DataModule.pas:3753`). |
| multiple rows for one (line,date) | the form binds to the **first** row (`fieldbyname(...)`) — there is no `ORDER BY`, and the table is read with no guaranteed order | possible only if the UI guard were bypassed; today the guard prevents creating a 2nd, so 0-or-1 in practice |
| trailing space / case on `@LineName` | `=` is CI_AS + ANSI trailing-space-trim, so `'COROLLA '` matches `'COROLLA'` | inherited collation behavior |

---

## 4. The TWO cost checks are genuinely different mechanisms (VERIFIED)

There are two cost guards in the chain. They are **not** the same query and catch **different** conditions.

### 4a. PRE-loop guard — hard ABORT on `IN_MANIFEST_COST_ID IS NULL`

Runs *before* any detail insert, per BC, in Delphi (`DataModule.pas:5160-5175`) over the result of
`SELECT_ForecastDetailBCASN` (which LEFT JOINs `INV_MANIFEST_COST_MST`). `IN_MANIFEST_COST_ID` is that cost
table's PK; it is NULL **only via a LEFT-JOIN miss** = the assy part has **no cost-master row at all**. If ANY
recipe row for the BC is NULL, Delphi RAISEs `"Missing Manifest Cost Information BCode(…) … ASN create failed"`
and the whole create fails. (Full lineage in `SELECT_ForecastDetailBCASN-analysis.md §4`.)

Catches: **assy part with no manifest-cost master entry whatsoever** (un-manifestable; the manifest number
would otherwise be built from NULL). Live: 5 distinct forecast assy parts qualify:

```
$ forecast parts with NO cost master (IN_MANIFEST_COST_ID null via LEFT JOIN miss):
42600F277100  42670F290200  42670F2X4100  42670FEB7000  42670FEL2000      (count = 5)
```

### 4b. POST-loop guard — non-aborting WARN, date-windowed (SELECT_ASNMissingCost)

**Signature:** `@ASNID INTEGER`. **Body (identical in both DBs + dump):**

```sql
SELECT d.VC_MANIFEST_NUMBER 'Manifest', d.VC_ASSY_PART_NUMBER 'PartNumber',
       CASE WHEN m2.MO_PRICE IS NULL THEN 'Missing Manifest Cost Entry'
            ELSE 'Manifest Cost Entry is out of date' END AS 'ErrorMsg'
FROM INV_ASN_MST a
JOIN INV_ASN_DETAIL_MST d ON a.IN_ASN_ID = d.IN_ASN_ID
LEFT JOIN INV_MANIFEST_COST_MST m                          -- m = DATE-WINDOWED join
   ON d.VC_ASSY_PART_NUMBER = m.VC_ASSY_PART_NUMBER_CODE
  AND m.VC_START_MANIFEST <= a.VC_PRODUCTION_DATE
  AND m.VC_END_MANIFEST   >= a.VC_PRODUCTION_DATE
LEFT JOIN INV_MANIFEST_COST_MST m2                         -- m2 = part-only join (no date window)
   ON d.VC_ASSY_PART_NUMBER = m2.VC_ASSY_PART_NUMBER_CODE
WHERE a.IN_ASN_ID = @ASNID
AND m.MO_PRICE IS NULL                                     -- the date-windowed cost is MISSING
```

Runs *after* all detail rows are written, once for the whole ASN. The Delphi handler
(`DataModule.pas:5293-5302`) only `ShowMessage`s + logs each returned row and **does not raise** — and the
whole call is in a `try/except` that swallows errors (`:5303-5307`). **It cannot abort the create.**

Predicate logic — it returns a row when the **date-windowed** cost (`m`) is absent
(`m.MO_PRICE IS NULL`), then the CASE splits by whether a part-only cost (`m2`) exists:
- `m2.MO_PRICE IS NULL` → **"Missing Manifest Cost Entry"** (no cost master at all — same population as 4a).
- `m2.MO_PRICE` not null → **"Manifest Cost Entry is out of date"** (a cost master exists, but the ASN's
  `VC_PRODUCTION_DATE` falls **outside** `[VC_START_MANIFEST, VC_END_MANIFEST]`).

`VC_START_MANIFEST`/`VC_END_MANIFEST`/`VC_PRODUCTION_DATE` are all `varchar(8)` `yyyymmdd`; the `<=`/`>=`
comparison is **lexicographic string** comparison, which for zero-padded `yyyymmdd` equals chronological order
(VERIFIED safe for this format; a rebuild using real dates must keep the same boundary inclusivity — both ends
inclusive).

### 4c. The distinction proven end-to-end (rolled-back tran, prod date 20260615)

Header created with `INSERT_ASNInfo` (`@Ein=0`), two detail rows:
- `42670FET9000` — **has** a cost master, window `20230404..20251227` (expired before 20260615).
- `42670FEB7000` — has **no** cost master at all.

```
--- POST-loop SELECT_ASNMissingCost @ASNID=4745:
70615YY | 42670FEB7000 | Missing Manifest Cost Entry            (m2 null)
70615ZZ | 42670FET9000 | Manifest Cost Entry is out of date     (m2 not null, date-window miss)
--- PRE-loop equivalent (IN_MANIFEST_COST_ID IS NULL?):
42670FET9000 | has id  -> pre-loop PASSES        ← slips PAST the hard abort
42670FEB7000 | NULL    -> pre-loop WOULD abort
--- ROLLED BACK → hdr_after_rollback = 0
```

**Conclusion — different mechanisms, different catch sets:**

| | PRE-loop (4a) | POST-loop (4b, SELECT_ASNMissingCost) |
|---|---|---|
| where | Delphi, per BC, before inserts | proc, whole ASN, after inserts |
| keyed on | `IN_MANIFEST_COST_ID IS NULL` (part has no cost master) | `m.MO_PRICE IS NULL` (no cost in the **date window**) |
| date-aware? | **No** | **Yes** (`VC_START/END_MANIFEST` window) |
| outcome | **RAISE → ASN create fails** | ShowMessage + log; **never aborts** (try/except swallows) |
| catches | un-manifestable part (no cost master) | both "no cost master" AND **"cost exists but expired/not-yet-effective"** |

The **"out of date"** case (`42670FET9000`) is caught **only** by the post-loop warn — the pre-loop hard abort
passes it through because the part *does* have a cost-master row (its `IN_MANIFEST_COST_ID` is non-NULL). The
ASN is created with that detail line; the operator is merely warned. Conversely the pre-loop abort stops a
"no cost master" part *before* a detail row is even written, so for that population the post-loop warn is a
belt-and-suspenders re-detection of the rows that survived (it can still report it because the abort is
per-BC — if BC #1 had a missing-cost part the create already RAISEd and never reached the post-loop).

---

## 5. Drift check: Inventory vs Inventory_Live (VERIFIED)

Byte-level diff of `OBJECT_DEFINITION` for each proc:

| proc | Inventory vs Inventory_Live |
|---|---|
| `INSERT_ASNInfo` | **IDENTICAL** (and matches dump) |
| `INSERT_ASNDetail` | **DRIFT** — the Q1 re-key (manifest-only → `(IN_ASN_ID, manifest)`). §2. |
| `SELECT_ASNSeq` | **IDENTICAL** |
| `SELECT_ASNMissingCost` | **IDENTICAL** |

Plus the index drift on `INV_ASN_DETAIL_MST`: spike `Inventory` has the added
`IX_INV_ASN_DETAIL_MST_ASN_MANIFEST (IN_ASN_ID, VC_MANIFEST_NUMBER)`; legacy has only
`IX_INV_ASN_DETAIL_MST (VC_MANIFEST_NUMBER)`. The companion `DELETE_ASNItem` is likewise re-keyed on the spike
(gains `@ASNID`) per `spike-asndetail-rekey.sql` — out of this analysis's four procs but part of the same Q1
cutover and noted so the rebuild keeps the pair consistent.

---

## 6. What a faithful reimplementation MUST reproduce (and the traps that silently break it)

1. **Order: header first, then details.** `INSERT_ASNInfo` must commit and return its IDENTITY before any
   `INSERT_ASNDetail`, because that IDENTITY is the `IN_ASN_ID` the details write *and* (post-re-key) the dedup
   key. (§0)
2. **`SCOPE_IDENTITY()`, not `@@IDENTITY` / `IDENT_CURRENT`.** Use the scoped value or JDBC generated-keys.
   No trigger today makes `@@IDENTITY` agree by luck — don't rely on luck. (§1)
3. **Status `'C'` is hard-coded** on create; there is no status param. (§1)
4. **EIN: write `0` at create** (placeholder in `IN_ASN_EIN` on both header and detail); assign the real EIN at
   send. The legacy `fEIN+1` bump moves to the send path. (§1)
5. **Re-keyed detail upsert keys on `(IN_ASN_ID, VC_MANIFEST_NUMBER)`** for `@HotCall=0`: accumulate
   (`IN_QTY += @Qty`) within one ASN, never collide across ASNs. `@HotCall=1` always inserts. Do NOT reduce
   this to manifest-only (legacy bug) and do NOT "fix" the within-ASN accumulate away — it is load-bearing
   (one manifest summed from several recipe rows). (§2, proof §2d)
6. **Prefer an explicit column-list INSERT** for the detail row. The legacy positional 8-value VALUES against
   the 9-column IDENTITY table is fragile: any column add/reorder (esp. M4 `IN_SITE_ID`) silently breaks it.
   Map: `(IN_ASN_ID, IN_INV_ID=null, IN_ASN_EIN, VC_MANIFEST_NUMBER, VC_ASSY_PART_NUMBER, IN_QTY,
   VC_LAST_UPDATE, VC_ADD)`. (§2c)
7. **Idempotency guard = existence read-back of `(line, prodDate)` with `START_SEQ <> -1`, enforced by
   disabling create.** There is no DB unique constraint; if the rebuild allows concurrent creates it must add
   one or keep the read-back guard. Compare the sentinel as a **string** (`<> '-1'`) to avoid the legacy
   implicit varchar→int cast on `VC_START_SEQ_NUMBER`. (§3)
8. **Two cost checks, kept separate.** Pre-loop: hard-abort the create when any recipe row's
   `IN_MANIFEST_COST_ID IS NULL` (no cost master), surfacing the assy part number(s). Post-loop:
   non-aborting warn from `SELECT_ASNMissingCost`, which is **date-windowed** and additionally catches
   "cost exists but the production date is outside `[VC_START_MANIFEST, VC_END_MANIFEST]`" → "out of date".
   Do not collapse them into one check; the post-loop's date window is the only thing that catches expired
   cost. Keep the boundary inclusive on both ends and keep the comparison on zero-padded `yyyymmdd`. (§4)
9. **No surrounding transaction today** → a mid-loop abort leaves a partial ASN (header + earlier BC details).
   If the rebuild wraps the create in one transaction (recommended), that *changes* the failure semantics from
   "partial persists" to "all-or-nothing" — an improvement, but flag it as a behavior change, don't assume
   parity. The post-loop cost call's swallow-all `try/except` must also be preserved: it can warn but must
   never fail a create. (§0, §4b)
10. **16-char `VC_ADD`/`VC_LAST_UPDATE` stamp** = `yyyymmddHHmmss` + first 2 ms digits (style-114 quirk), not
    14 chars. Match it if byte-level parity on the stamp matters; nothing parses it back. (§1)

---

## Appendix — query log (bounded; destructive probes rolled back)

- Proc bodies: `OBJECT_DEFINITION(OBJECT_ID('dbo.<proc>'))` in `Inventory` and `Inventory_Live` →
  `/tmp/asn_procs/*.sql`; diffed (§5).
- Schemas: `sys.columns`/`sys.types` for `INV_ASN_MST` (13 cols, col1 IDENTITY),
  `INV_ASN_DETAIL_MST` (9 cols, col1 IDENTITY), `INV_MANIFEST_COST_MST`; `sys.indexes`/`sys.index_columns`
  for the re-key index. Param catalogs via `sys.parameters`.
- Re-key proof: `BEGIN TRAN … ROLLBACK` on `Inventory`, synthetic manifest `ZTEST999`, real headers
  1804/4715; accumulate=42, separate B=100, HotCall new row, legacy-vs-rekey WHERE counts 1-vs-0;
  rows_after_rollback=0. (§2d)
- Two-cost proof: `BEGIN TRAN … ROLLBACK`, header @ASNID=4745 (`@Ein=0`, status 'C'), parts
  42670FET9000 (out-of-date) + 42670FEB7000 (no master); SELECT_ASNMissingCost returned both rows with the
  two distinct ErrorMsgs; pre-loop NULL-check showed FET9000 passes / FEB7000 aborts; hdr_after_rollback=0.
  (§4c)
- Stamp decomposition: style-114 `09:23:33:673` → 16-char `2026061509233367`. (§1)
- Real cross-ASN manifest reuse: top manifests in 3 distinct ASNs (52081285, 52066074, …). (§2d)
