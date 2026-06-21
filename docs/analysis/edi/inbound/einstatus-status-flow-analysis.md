# M1 inbound (997/824) — `UPDATE_EINStatus` + the ASN/invoice status flow (SQL source-truth)

**Scope:** decode the DB side of the inbound ack — `UPDATE_EINStatus` (the proc that flips ASN/invoice
status when a 997/824 arrives) and the full `VC_ASN_STATUS` / `VC_INV_STATUS` state machine, so the
rebuild's inbound processor writes the right **site-scoped** status.

**Evidence base:** live `Inventory` + read-only `Inventory_Live` parity snapshot on `mssql-spike` (SQL Server
2019); live `VehicleOrder` (ALC) for the cross-DB EIN counter. Proc bodies confirmed LIVE ==
`/tmp/inv_utf8.sql`. Caller = `EDIUpload.pas` (997 branch :186-252, 824 branch :253-305) +
`DataModule.pas:6753 UpdateEINStatus`. Builds on `../810/report-edi810-data-analysis.md:56` and
`../856/report-edi856-data-analysis.md:45`.

---

## 1. `UPDATE_EINStatus(@EIN int, @EINStatus varchar(1), @EINType varchar(2))`

**Live body (verified identical to `/tmp/inv_utf8.sql:1711-1730`; OBJECT_DEFINITION on `Inventory` matches
byte-for-byte):**

```sql
CREATE PROCEDURE [dbo].[UPDATE_EINStatus]
    @EIN        integer,
    @EINStatus  varchar(1),
    @EINType    varchar(2)
AS
BEGIN
    SET NOCOUNT ON;
    if @EINType = 'SH'
    begin
        UPDATE INV_ASN_MST SET VC_ASN_STATUS = @EINStatus WHERE IN_ASN_EIN = @EIN   -- 856/ASN ack
    end
    else
    begin
        UPDATE INV_INV_MST SET VC_INV_STATUS = @EINStatus WHERE IN_INV_EIN = @EIN   -- 810/invoice ack (DEFAULT)
    end
END
```

**Behavior:** a two-branch UPDATE keyed on EIN.
- `@EINType = 'SH'` (exact, case-insensitive default collation) → flip the **ASN** row(s).
- **anything else** (incl. `'IN'`, `''`, NULL, garbage) → flip the **invoice** row(s). The `else` is the
  default — there is no explicit `'IN'` test. (`@EINType` is `varchar(2)`; the caller passes `copy(fcl,5,2)`
  from the AK1 segment = `'SH'` or `'IN'`, see §5.)
- Returns **no result set and no status** to the caller. `@@ROWCOUNT` is discarded inside the proc; the
  Delphi wrapper (`DataModule.pas:6755`) hard-codes `result := TRUE` before the call and only flips to error
  on an ADO `Errors.Count > 0` (a SQL-level failure), **not** on "0 rows updated." So an ack for a
  non-existent EIN silently no-ops and reports success (see §5 edge cases).

### The site-scoping gap (the headline fix)
Both UPDATEs filter on **EIN only** — `WHERE IN_ASN_EIN = @EIN` / `WHERE IN_INV_EIN = @EIN`. **No site
predicate**, because the legacy schema is single-site:

- **Proven:** neither table has a site column.
  `INV_ASN_MST` cols incl. `IN_ASN_EIN int NOT NULL`, `VC_ASN_STATUS varchar(1) NOT NULL` — **no
  `IN_SITE_ID`**. Same for `INV_INV_MST` (`IN_INV_EIN int NOT NULL`, `VC_INV_STATUS varchar(1) NOT NULL`,
  no site col). Query: `sys.columns` filtered to those tables returned only the EIN + status (+ id/no)
  columns; no site column exists.
- **Why it's latent today (proven):** EIN is globally unique within each table — on `Inventory_Live`,
  `INV_ASN_MST` = 2557 rows / 2557 distinct EIN (0 dup-EIN groups, 0 null/0); `INV_INV_MST` = 2941 rows /
  2941 distinct EIN (0 dup, 0 null/0). And **zero overlap across the two tables** (`ASN JOIN INV on
  EIN` = 0 rows) — so the `@EINType` branch is currently redundant for routing, but the two namespaces are
  NOT formally disjoint (see below).
- **Why it becomes a real collision under multi-site (proven cross-DB):** the EIN is allocated from **one
  shared per-site counter** — `VehicleOrder.dbo.Site.SiteEIN` (`int`), bumped by
  `AD_UpdateEIN` = `UPDATE Site SET SiteEIN = SiteEIN+1` (full body verified on live `VehicleOrder`;
  **no WHERE** — a second site-scoping hazard: it would bump every site's counter). The app does
  `EIN := SiteEIN + 1` then calls `AD_UpdateEIN` (`MainMenu.pas:2619, 2636`). **The same counter feeds BOTH
  856 (ASN) and 810 (invoice)** — proven by the numeric ranges interleaving on Live: ASN EIN 3502-9071,
  INV EIN 29-9072. So once a second site exists, (a) site X EIN 5000 and site Y EIN 5000 both exist →
  `WHERE IN_ASN_EIN = 5000` flips **both sites**; and (b) within one site the shared counter means an ASN and
  an invoice can carry the same EIN value → only `@EINType` keeps them apart. Both hazards are dormant only
  because there is exactly one site and one global sequence today.

**Rebuild contract:** key the ack write by **(site_id, EIN, type)** —
`UPDATE <asn|inv> SET status = @status WHERE site_id = @site AND ein = @ein`. Carry the `@EINType`→table
routing (`'SH'`→ASN, else→invoice). Add `IN_SITE_ID` to both master tables and to the EIN counter
(per-site `IN_EIN_SEQ`, atomic, site-scoped — replacing the WHERE-less `AD_UpdateEIN`).

---

## 2. Status domain + flow (`VC_ASN_STATUS` / `VC_INV_STATUS`, both `varchar(1) NOT NULL`)

**Authoritative value→name decode** (from the report procs' `CASE` — ASN `/tmp/inv_utf8.sql:1955-1960`,
invoice `:3254-3257`). Note the invoice decode lists only A/S/R; "R" is misspelled `Rejecteed` in the proc
text (cosmetic, doesn't affect storage):

| code | meaning            | set by                                                                 |
|------|--------------------|------------------------------------------------------------------------|
| `C`  | Create File        | ASN insert `INSERT_ASNInfo` (`:2551` literal `'C'`); invoice "recreate" only |
| `S`  | Sent               | send-flip C→S; invoice **insert** `INSERT_INVInfo` (`:2581` literal `'S'`) |
| `A`  | Accepted           | 997 ack via `UPDATE_EINStatus` (`AK5 = 'A'`)                            |
| `R`  | Rejected           | 997 ack via `UPDATE_EINStatus` (`AK5 != 'A'`)                           |

**The flow differs by document type:**

- **ASN (856):** `INSERT_ASNInfo` writes `'C'` (`:2551`). Send-flip `'C'→'S'`: per-ASN
  `REPORT_EDI856 @EIN<>0` self-flip (`:3695`) — or the blanket hazard `UPDATE_ASNStatus`
  (`SET 'S' WHERE 'C'`, all rows, `:1695`; **rebuild must NOT use**). Ack `'S'→A/R` via `UPDATE_EINStatus`.
  So: **C → S → A/R**.
- **Invoice (810):** `INSERT_INVInfo` writes `'S'` **directly** (`:2581`) — invoices are born Sent (the
  810 file is created in the same action; the `'C'` invoice insert at `:3390` is **commented out**). The
  `'C'` invoice state exists only on a **recreate/resend** path (`UPDATE_INVRecreate` sets `'C'→'S'` `:1568`;
  `REPORT_EDI810Recreate` joins `i.VC_INV_STATUS='C'` `:3723`, with its ASN already `'A'`). Ack `'S'→A/R`.
  So normal: **S → A/R**; recreate: **(A) → C → S → A/R**.

**Proven distribution (read-only `Inventory_Live`):**

```
INV_ASN_MST   VC_ASN_STATUS  A = 2557        (100% Accepted)
INV_INV_MST   VC_INV_STATUS  A = 2935 / S = 6   (99.8% Accepted, 6 still Sent)
```

(Live `Inventory` rebuild target nearly identical: ASN A=2550; INV A=2928, S=6.) **No `C` and no `R` appear
anywhere on the snapshot.** Consequence for parity testing: the snapshot cannot exercise the `C` send branch
or the `R` reject branch — these must be tested with synthetic fixtures (tag e.g. `VC_ADD` so they're
distinguishable). The 6 stuck `S` invoices = sent-but-never-acked (no 997 came back, or it failed to update).

---

## 3. AK9 code → status mapping (Q6 recommendation)

**What the legacy actually reads (important):** the caller does NOT read AK9. In `EDIUpload.pas` the 997
loop reads `AK1` for routing then reads the **next line** and takes `Status := copy(fcl,5,1)` (`:197`) — i.e.
a single char at offset 5 of the segment after AK1. In a 997 that segment is **AK5** (transaction-set
response, element AK501) when present, or AK9 (functional-group response, AK901) for group-level. Either way
legacy collapses it to a **single char** and tests **`if Status = 'A'` → Accepted, `else` → Rejected**
(`:205` vs `:218`). So legacy is **binary: A vs everything-else-as-R** (it writes the raw char as
`@EINStatus`, but only the `'A'` history message vs the reject history message branch on it). Any AK code
that is not literally `'A'` (E/P/R/M/W…) is funneled to the "Rejected" history line, and whatever single char
it was gets stored verbatim in `VC_*_STATUS`.

**EDI standard AK5/AK9 acknowledgment codes:** `A` Accepted, `E` Accepted-with-errors, `P` Partially-accepted,
`R` Rejected, `M` Rejected-msg-auth, `W` Rejected-assurance, `X` Rejected-content.

**Q6 recommendation — write a distinct status per AK code (superset of legacy):**

| AK5/AK9 | meaning                 | rebuild `@EINStatus` | legacy did      |
|---------|-------------------------|----------------------|-----------------|
| `A`     | Accepted                | `A` (Accepted)       | `A`             |
| `E`     | Accepted, with errors   | `E` (Accepted/errors)| collapsed → R-history (stored `E`) |
| `P`     | Partially accepted      | `P` (Partial)        | collapsed → R-history (stored `P`) |
| `R`     | Rejected                | `R` (Rejected)       | collapsed → R-history (stored `R`) |
| `M/W/X` | Rejected (auth/content) | `R` (Rejected)       | collapsed → R-history |

This is a strict **superset**: `A`/`R` keep their legacy meaning; `E`/`P` become first-class instead of being
lumped under "Rejected." Widen the status domain accordingly (the `varchar(1)` column already fits a single
char; the report `CASE` decodes must add `E`/`P` arms or they render NULL). Per Q6, also tolerate AK2/AK3/AK4
detail segments (transaction-set / data-segment / data-element notes) when present — read them for the reject
report, don't let them break the parse — and map A/E/P/R distinctly as above.

---

## 4. What the rebuild's inbound status-write MUST reproduce

1. **Two-branch routing on `@EINType`:** `'SH'` → ASN table, else → invoice table. Keep the `else`-is-default
   semantics (treat `'IN'` and anything-non-`SH` as invoice), OR tighten to explicit `'SH'`/`'IN'` and
   reject unknown — but document the change; legacy routes all non-`SH` to invoices.
2. **Site-scoped UPDATE:** `WHERE site_id = @site AND ein = @ein` on the correct table. This is the one
   behavioral *fix* vs legacy (legacy is EIN-only). Key the whole operation by **(site_id, EIN, type)**.
3. **AK-mapped status:** write the §3 mapping (A/E/P/R distinct; M/W/X→R). The 824 reject path → write the
   **Rejected** status `R` (legacy never does this — see trap below) plus set an auto-flag for follow-up
   (Q10: 824/NTE error → flag the ASN/invoice + surface the manifest+part+error text for the operator).
4. **No-op visibility:** capture `@@ROWCOUNT`. 0 rows = "ack for unknown (site,EIN,type)" → log/alarm
   (legacy hides this; the Delphi wrapper falsely returns success). 
5. **Re-ack tolerance:** an ack landing on an already-A/R row simply overwrites (idempotent UPDATE). Decide
   whether to allow A→R / R→A downgrades-upgrades or to lock terminal states; legacy allows any overwrite.
6. **EIN allocation:** preserve the **shared per-site** counter (one sequence feeds both 856 and 810) but make
   it **atomic + site-scoped** (`INV_SITES.IN_EIN_SEQ` / per-site row), replacing the WHERE-less
   `AD_UpdateEIN` blanket bump. The ack `WHERE` must match the EIN that was actually allocated/sent.

### The traps that will silently break it
- **No site filter (legacy)** → under multi-site, a cross-site ack flips the wrong site's row (and possibly
  both). The single biggest latent bug.
- **Shared counter ⇒ ASN/invoice EIN values can coincide** within a site → dropping `@EINType` from the WHERE
  key would cross-contaminate. Keep the type→table routing.
- **824 does NOT flip status in legacy.** The 824 branch (`EDIUpload.pas:253-305`) only writes an Excel error
  report and appends history — it **never calls `UpdateEINStatus`**. So a hard reject delivered via 824 leaves
  the row stuck at `'S'` (this likely explains the 6 stuck `S` invoices on Live). The rebuild MUST make 824 a
  real status event (→`R` + flag), not just a report.
- **Binary A-vs-rest collapse** (`copy(fcl,5,1)` + `if = 'A'`) loses E/P granularity and reads a single char
  at a fixed offset — fragile to segment layout. Parse AK5/AK9 properly (segment-aware), don't offset-slice.
- **Report `CASE` arms are incomplete:** any status the decode doesn't list renders NULL (e.g. an invoice in
  `'C'` against the A/S/R-only invoice decode). Adding E/P statuses requires adding decode arms everywhere
  status is displayed.
- **Wrapper always reports success** (`result:=TRUE`, error-only on ADO exception) → don't model "did the
  ack land?" on the proc's return; model it on `@@ROWCOUNT`.

---

## 5. Edge cases (with the evidence)

- **Ack for a non-existent EIN:** UPDATE matches 0 rows, proc returns nothing, wrapper still reports success
  (`DataModule.pas:6755 result:=TRUE`; only ADO `Errors.Count>0` flips it). **Silent.** Rebuild: check
  `@@ROWCOUNT`, log/alarm on 0.
- **Re-ack of an already-A/R row:** plain overwrite — idempotent for same value; a later differing ack
  silently changes the stored status (A→R or R→A). Legacy permits it. Decide terminal-state policy in rebuild.
- **EIN-type disambiguation (how inbound knows SH vs INV):** from the **997 `AK1` functional-group code** —
  `EDIType := copy(fcl,5,2)` (`EDIUpload.pas:194`), which is the AK1 element carrying the original group's
  functional identifier (`'SH'` = 856 Ship Notice group, `'IN'` = 810 Invoice group). That string is passed
  straight through as `@EINType`. Rebuild: read AK101 and map the FG code to ASN/invoice; do not infer type
  from the EIN value (the namespaces are not disjoint under the shared counter).
- **`@EINType` NULL/blank/unknown:** falls into the `else` → treated as an **invoice** ack. A malformed 997
  with a missing/garbled AK1 functional code would flip an invoice row, not error. Rebuild should validate.
- **Snapshot blind spots:** Live has **no `C` and no `R`** in either table (§2) — the send branch and the
  reject branch are unobservable on the parity data; test them with tagged synthetic fixtures.
- **The 6 stuck `S` invoices on Live:** sent, never acked to A/R — consistent with the 824-doesn't-flip trap
  and/or a lost 997. Useful as a fixture for the "ack never arrived / arrived as 824" path.

---

## Cross-references / dependencies
- EIN allocation + the per-site counter, and the send-flip / `UPDATE_ASNStatus` & `UPDATE_INVRecreate`
  blanket hazards: `../856/report-edi856-data-analysis.md:30-47`,
  `../810/report-edi810-data-analysis.md:37-58`.
- Cross-DB: `VehicleOrder.dbo.Site.SiteEIN` (`int`) + `AD_UpdateEIN` (`UPDATE Site SET SiteEIN=SiteEIN+1`,
  no WHERE) — the shared counter feeding both 856 and 810. Caller: `MainMenu.pas:2619, 2636`.
- Caller of `UPDATE_EINStatus`: `DataModule.pas:6753` (wrapper) ← `EDIUpload.pas:203` (997 branch).
