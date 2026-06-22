# M3 Report Usage — PRUNE LIST (rebuild vs skip)

> **SUPERSEDED (2026-06-22):** the conservative "DECIDE most reports" verdict below was based on the 1-day
> `DailyWorkLog.csv`. It is superseded by David's full 1-year exports (`ReportLog.csv` = REPORT trans;
> `AllLog.csv` = all trans), which showed only **4 of 22 reports ran in a year** (18 never, 3 once) → M3
> rebuilt exactly those 4.

**Phase:** M3 source-truth (2026-06-21) **Analyst:** Claude (SQL-semantics) **Scope:** decide which of
the 22 reports (R1–R22) + 8 Excel companions (C1–C8) in `m3-report-inventory.md` to REBUILD vs SKIP,
driven by *actual usage signal* — David's directive: "most of the 22 reports are NOT in use; prune the
dead ones before we rebuild any."

> **Bottom line first:** the spike's activity-log table is **NOT real production history** (synthetic,
> 78 rows, this week, zero `REPORT` rows). The only real usage signal available is David's
> `DailyWorkLog.csv` — and it is **one production day** (2026-06-19). One day is a *thin* prune signal,
> so this doc is **conservative**: it recommends **DECIDE** for almost everything and proposes asking
> David for a multi-month `Act_Log` export to make the cut data-driven. The single hard data point: on
> a real production day, exactly **one** report was run (Daily Shipping Assy, R3/R4).

---

## 1. The usage signal: where report-runs are logged

### 1a. The logging mechanism (Delphi → Activity DB)

`Data_Module.LogActLog('REPORT', 'Do <name>')` is the usage emitter. Source:
`DataModule.pas:6020` (`procedure TData_Module.LogActLog`). It writes through `Act_StoredProc`, which the
DFM binds to proc `InsertDetailedAct_Log;1` on `Act_Connection` (`DataModule.dfm:41-43`). The Delphi call
binds 11 params (`@App_From, @IP_Address, @Trans, @DT_Sender, @ComputerName, @Description, @NTUserName,
@DB_Time, @VIN, @Sequence_Number, @Last_Modified`) — i.e. the *detailed* activity-log contract.

Every report handler in `MainMenu.pas` calls `LogActLog('REPORT','Do <report name>')` in its **success
path** (after the data proc opens / around the OLE write). The 21 distinct strings are the per-report
usage key (full map in §3). All 24 `REPORT`-type call sites are in `MainMenu.pas` only — no other unit
logs `REPORT` (`grep LogActLog('REPORT'` → MainMenu.pas exclusively).

### 1b. The target table on the spike — Activity.dbo.Act_Log — IS SYNTHETIC (NOT usable)

`Activity` DB exists on `mssql-spike`. It contains exactly one table + one proc:

| Object | Type |
|---|---|
| `dbo.Act_Log` | USER_TABLE |
| `dbo.InsertAct_Log` | PROC |

`Act_Log` columns: `id int`, `VC_TYPE varchar(50)`, `VC_SOURCE varchar(50)`, `VC_MESSAGE varchar(200)`,
`DT_WHEN datetime`. **Verdict on its data:**

```
total_rows  min_dt                   max_dt
78          2026-06-17 21:52:35.550  2026-06-21 21:59:10.803     -- this week (spike build window)
VC_TYPE:  INVENTORY = 78  (ONLY value)
VC_SOURCE: TRIGGER = 78   (ONLY value)
REPORT-type rows: 0
```

This is **spike-generated** data, not a restored production Activity snapshot: 78 rows, all dated inside
the 2026-06-17→06-21 spike work window, all `VC_TYPE='INVENTORY'` / `VC_SOURCE='TRIGGER'` (the stock-ledger
trigger firing during the spike's ledger work). There are **zero `REPORT` rows**. There is also a
proc/schema **drift**: the spike's `InsertAct_Log` takes only `(@Type, @Source, @Message)` and inserts
3 columns, whereas the live Delphi target is the 11-param `InsertDetailedAct_Log`. The spike Activity DB
is a stub, not the production log.

> Cross-checked the Inventory + Inventory_Live DBs for any alternate report-usage/history table:
> only domain-history tables exist (`INV_ASSY_BUILD_HIST`, `INV_FORECAST_DETAIL_INF_HIST`,
> `INV_OPEN_ORDER_INF_HIST`, `INV_PARTS_STOCK_MST_HIST`, `INV_EDI_INBOUND_LOG`) — none log report runs.

**Conclusion:** the activity-log signal is **not available on the spike**. Per the directive, fall back
to the DailyWorkLog.

**Queries used** (all `SET NOCOUNT ON`, bounded):
`Activity`: `SELECT name,type_desc FROM sys.objects WHERE type IN ('P','U')`;
`SELECT COUNT(*),MIN(DT_WHEN),MAX(DT_WHEN) FROM dbo.Act_Log`;
`SELECT VC_TYPE,COUNT(*) FROM dbo.Act_Log GROUP BY VC_TYPE`;
`SELECT COUNT(*) FROM dbo.Act_Log WHERE VC_TYPE='REPORT'` → 0.

### 1c. Fallback signal: DailyWorkLog.csv (REAL, but ONE DAY)

`DB Schema/DailyWorkLog.csv` (gitignored real client data — **summarized only, not echoed**) is a real
production export in the exact activity-log shape `App_From, Trans, Message, User, DT_WHEN`:

- **174 rows, a single day:** 2026-06-19 07:52:11 → 14:57:29.
- **Two operators:** `lktur` (KTURNER, 98 rows) and `Administrator` (GHEATH, 76 rows).
- App version logged: 2.9.4.1.

Trans-type distribution (this is the real "what the operators actually do all day" signal):

| Trans | rows | Trans | rows |
|---|---|---|---|
| SELECT SFT | 32 | EDI | 4 |
| ORDERF | 27 | LOGIN / START / STOP | 3/3/2 |
| ASN | 23 | UPD ASNSta | 2 |
| RENBAN BD | 19 | INS ASN | 2 |
| ORDER | 13 | HOTCALL | 2 |
| ASNINV | 11 | **REPORT** | **1** |
| EDIIMP | 11 | ERROR | 5 |
| ORDERS | 9 | UPD EINSta | 5 |

The day's work is dominated by the **daily ASN→856→810 loop + order/renban breakdown** (ASN+ASNINV+EDI+
EDIIMP+ORDER+ORDERF+ORDERS+RENBAN BD = ~123 of 174 rows). **Reports are a rounding error: 1 of 174.**

---

## 2. The single hard usage data point

The day's **only** `REPORT` row:

```
Trans=REPORT  Message="Do Daily Shipping Assy Report"  User=lktur  TS=2026-06-19 08:48:55.820
```

…preceded by **three** `ERROR` rows: `"Failed on Daily Shipping Assy Report, PrintOut method of
Worksheet class failed"` at 08:34, 08:35, 08:47. So the operator tried the Daily Shipping Assy report
**4 times**, the first 3 failing on the Excel `PrintOut`, the 4th succeeding (the 08:48 REPORT row).

Two things follow:
1. **Daily Shipping Assy (R3/R4) is genuinely a daily operational report** — corroborates the
   `m3-report-inventory.md` flag (R1–R3 are P0 daily shipping) and the known Excel-`PrintOut` failure
   class (same `PrintOut method ... failed` already in scope as the FAILING path).
2. **The REPORT row only fires on success.** A report attempted-and-failed leaves *ERROR* rows, not a
   REPORT row. So a raw COUNT of REPORT rows **undercounts demand** for any report whose Excel/OLE write
   is flaky. Treat low/zero REPORT counts as "low *successful* usage," not proven "never wanted."

---

## 3. Message-string → report map (the usage key)

Each report's `LogActLog('REPORT',...)` string from `MainMenu.pas` (the GROUP-BY key you'd run against a
real `Act_Log`). **Caveats baked in:**

- **Shared strings (COUNT cannot separate them):**
  - `"Do Daily Shipping Assy Report"` → **R3 (Daily ASN) AND R4 (Monthly ASN)** — both handlers log the
    same string (`MainMenu.pas:3210` and `:3435`). A run count on this string is R3+R4 combined.
  - `"Do Daily Supplier Order"` → **R5** (two call sites :1149, :1259 = the ±cost variants).
  - `"Do Supplier Order"` → **R6** (two call sites :1384, :1494 = ±cost variants).
- **`"Do ASN with Cost"`** (:3324) → **R-ASNwithCost, OUT of scope (D9 deprecated)** — but note its
  REPORT rows would still appear in a real log; exclude when counting.
- **No `REPORT`-type logging (usage is UNKNOWN from the report signal):**
  - **R22 (InvMgmt QReport)** — on-screen `TQuickRep.Preview`, no `LogActLog('REPORT')` (its usage would
    show only as the InvMgmt screen's own trans, not a REPORT row).
  - **C1–C8 (Excel companions)** — these do NOT log under `REPORT`. They log under their transactional
    types (`ORDERF/ORDERS` for the order sheet C1, `EDI/EDIIMP` for the 856/810 + echoes C4–C7,
    `ORDER` for sim C2, forecast trans for C3). On 2026-06-19 those types were *heavily* used
    (ORDERF=27, ORDERS=9, EDI/EDIIMP=15) — i.e. **the C-companions' parent workflows are daily-critical**,
    even though the .xls companion itself isn't separately counted.

| R# | Report | REPORT message string (exact) | MainMenu.pas |
|---|---|---|---|
| R1 | Daily Shipping (T/W) | `Do Daily Shipping Report` | :3103 |
| R2 | Daily Shipping Range (T/W) | `Do Daily Shipping Range Report` | :2997 |
| R3 | Daily Shipping ASN (Assy) | `Do Daily Shipping Assy Report` (shared w/ R4) | :3210 |
| R4 | Monthly Shipping ASN (Assy) | `Do Daily Shipping Assy Report` (shared w/ R3) | :3435 |
| R5 | Daily Supplier Order | `Do Daily Supplier Order` | :1149 / :1259 |
| R6 | Monthly Supplier Order (±cost) | `Do Supplier Order` | :1384 / :1494 |
| R7 | Monthly Logistics Order | `Do Logistics` | :1582 |
| R8 | Monthly Supplier Invoice | `Do Supplier Invoice` | :1715 |
| R9 | INVOICE Summary (D6) | `Do INVOICE Summary Report` | :3683 |
| R10 | Monthly INVOICE Summary (D6) | `Do Monthly INVOICE Summary Report` | :3560 |
| R11 | Logical Inventory | `Do Logical Inventory` | :1028 |
| R12 | Lot Location (PLANT) | `Do Lot Location` | :882 |
| R13 | Empty Container | `Do Empty Container` | :1833 |
| R14 | Past-Due / Late FRS | `Do Past Due FRS` | :2481 |
| R15 | PO Report | `Do PO Report` | :2867 |
| R16 | Forecast Parts Summary | `Do Forecast Parts Summary` | :1970 |
| R17 | Forecast Assy Summary | `Do Forecast Assy Summary` | :2065 |
| R18 | Forecast Detail | `Do Forecast Detail Report` | :3810 |
| R19 | Forecast vs Usage | `Do Forecast vs Usage` | :2395 |
| R20 | Unused Tire Part Numbers | `Do Unused Tire report` | :649 |
| R21 | Unused Wheel Part Numbers | `Do Unused Wheel report` | :699 |
| R22 | InvMgmt QReport | *(none — on-screen QuickReport)* | — |
| — | ASN with Cost (D9 OUT) | `Do ASN with Cost` | :3324 |

---

## 4. THE PRUNE LIST

**Usage column key:** the spike `Act_Log` gives **no** report data (synthetic). The only positive
evidence is the 2026-06-19 DailyWorkLog (1 day). So usage is "DailyWorkLog-1day" where seen, else
"NO LOG SIGNAL AVAILABLE."

**Verdict key:**
- **REBUILD** = positive real-usage evidence (run, or its parent daily workflow is run).
- **DECIDE** = no usage evidence either way (one-day log can't prove a monthly/occasional report is dead);
  ask David / get a multi-month `Act_Log` export before pruning.
- **SKIP** = already out of scope by prior decision (D9), OR clear, well-evidenced non-use. *Note: with
  only one day of real log, there is NO report for which "dead" is proven by the usage signal alone* —
  so SKIP here is reserved for the already-decided D9 deprecations, not new prune calls.

| # | Report | Usage signal | Verdict | Evidence / reason |
|---|---|---|---|---|
| R1 | Daily Shipping (T/W) | none in 1-day log | **DECIDE→lean REBUILD** | Daily-shipping family; R3 sibling ran. P0 in inventory doc. Likely daily; confirm cadence. |
| R2 | Daily Shipping Range (T/W) | none in 1-day log | **DECIDE** | Range variant of R1; may be ad-hoc. Confirm. |
| R3 | Daily Shipping ASN (Assy) | **RUN 2026-06-19** (4 attempts, 3 fail + 1 ok) | **REBUILD** | Only report run on a real prod day; daily operational. Hard evidence. |
| R4 | Monthly Shipping ASN (Assy) | shares R3 string (can't isolate) | **DECIDE→lean REBUILD** | Monthly cadence won't show in 1 day; same proc family as the proven R3. |
| R5 | Daily Supplier Order | none in 1-day log | **DECIDE** | "Daily" name but not run 06-19 (orders that day went via ORDER/ORDERF flow, not this report). Confirm. |
| R6 | Monthly Supplier Order (±cost) | none (monthly) | **DECIDE** | Monthly → invisible in 1 day. Don't prune blind. |
| R7 | Monthly Logistics Order | none (monthly) | **DECIDE** | Monthly → invisible in 1 day. |
| R8 | Monthly Supplier Invoice | none (monthly) | **DECIDE** | Monthly → invisible in 1 day. |
| R9 | INVOICE Summary (D6) | none (monthly/billing) | **DECIDE→lean REBUILD** | Billing numbers Toyota sees (P1 business risk in inventory doc); monthly cadence. High consequence if dropped. |
| R10 | Monthly INVOICE Summary (D6) | none (monthly) | **DECIDE→lean REBUILD** | Same billing-risk class as R9. |
| R11 | Logical Inventory | none in 1-day log | **DECIDE** | Periodic/ad-hoc inventory check. Confirm. |
| R12 | Lot Location (PLANT) | none in 1-day log | **DECIDE** | Print-template report; cadence unknown. Confirm. |
| R13 | Empty Container | none in 1-day log | **DECIDE→lean SKIP** | P3, niche; candidate for prune — but needs multi-month log to confirm zero use. |
| R14 | Past-Due / Late FRS | none in 1-day log | **DECIDE** | Exception report; runs only when there's a problem → naturally sparse. Don't prune on absence. |
| R15 | PO Report | none in 1-day log | **DECIDE→lean SKIP** | CAMEX-PO oriented (CAMEX decommissioned per D9 context); likely low/zero use. Confirm relevance. |
| R16 | Forecast Parts Summary | none in 1-day log | **DECIDE** | Forecast review; periodic. Forecast data WAS touched 06-19 (SELECT SFT on forecast detail) but report not run. |
| R17 | Forecast Assy Summary | none in 1-day log | **DECIDE** | Same as R16. |
| R18 | Forecast Detail | none in 1-day log | **DECIDE** | Same as R16. |
| R19 | Forecast vs Usage | none in 1-day log | **DECIDE→lean SKIP** | Analytical/ad-hoc; candidate for prune. Confirm with multi-month log. |
| R20 | Unused Tire Part Numbers | none in 1-day log | **DECIDE→lean SKIP** | Master-data hygiene report; run rarely (when cleaning part master). Candidate prune. |
| R21 | Unused Wheel Part Numbers | none (and has D11 col bug) | **DECIDE→lean SKIP** | Same as R20; plus the known wheel/tire column bug. Candidate prune. |
| R22 | InvMgmt QReport | no REPORT log; on-screen | **DECIDE→lean SKIP** | On-screen QuickReport, no Excel to retire, no usage signal. Lowest-value rebuild. |
| — | ASN with Cost | (D9) | **SKIP** | Already deprecated (D9). Out of scope. |
| — | NUMMI Lot Location[W] | (D9) | **SKIP** | Already deprecated (D9). |
| C8 | Forecast CAMEX | (D9) | **SKIP** | Already deprecated (D9, CAMEX gone). |
| C1 | Order sheet companion | parent ran heavily (ORDERF=27/ORDERS=9) | **REBUILD (companion .xls only)** | Parent order-sheet workflow is daily-critical. (Machine .ord artifact already built; retire .xls.) |
| C2 | Order simulation | none in 1-day log | **DECIDE→lean SKIP** | On-screen sim; .xls is layout-only. Low value. |
| C3 | Forecast breakdown (.frc) | parent forecast flow touched 06-19 | **DECIDE→lean REBUILD** | .frc machine artifact + P6 crash to fix; companion .xls cadence unknown. |
| C4 | 856/810 CSV companions | parent ran (EDI/EDIIMP=15) | **REBUILD (companion only)** | EDI 856/810 loop is the revenue-critical daily path; CSV companion supports it. |
| C5 | 862 FirmOrder echo | EDIIMP shows 862 processed 06-19 | **REBUILD** | 862 was processed on the real day (`EDI 862 Processed` rows). P1, deferred from M1/M2. |
| C6 | 861 Receiving Advice echo | none in 1-day log | **DECIDE→lean SKIP** | Inbound echo; cadence unknown. |
| C7 | 820 Remittance echo | none in 1-day log | **DECIDE→lean SKIP** | Inbound echo; cadence unknown. |

---

## 5. Recommended REBUILD-vs-SKIP split (conservative)

**Hard evidence only lets us firmly classify a handful.** With one day of real log:

- **REBUILD (evidence-backed, ~4 reports + 3 companions):**
  - **R3** (proven run), **R1/R4** (same daily-shipping family as R3), **R9/R10** (billing — drop = Toyota-visible risk).
  - Companions whose parent daily workflow is proven-active: **C1** (order sheet), **C4** (856/810 CSV), **C5** (862 echo).
- **SKIP (already decided, not new prunes):** ASN-with-Cost, NUMMI Lot Location[W], **C8** Forecast CAMEX (all D9).
- **DECIDE — everything else (the bulk: ~13 reports + several companions):** monthly/exception/ad-hoc
  reports that *cannot* be proven dead from one day. Several lean SKIP (R13 Empty Container, R15 PO,
  R19 Forecast vs Usage, R20/R21 Unused part-number hygiene, R22 InvMgmt QR, C2/C6/C7) — but none has
  *clear* zero-use evidence yet.

**So: the data does NOT yet support the "most reports are dead, prune them" cut.** What it supports is:
on a real day the operators ran the daily ASN/order/EDI loop intensively and exactly **1 report
(Daily Shipping Assy)**. That is consistent with David's hypothesis that report usage is low — but it is
**one day**, and the safe-to-prune set cannot be drawn from it without conflating "monthly/exception/
not-run-today" with "dead."

## 6. The blocker + the ask (to make this data-driven)

The prune list is only as good as the usage history. To turn the DECIDE rows into REBUILD/SKIP:

1. **Get a real `Act_Log` export from production** — ideally 6–12 months — and run, per §3 message string:
   ```sql
   SET NOCOUNT ON;
   SELECT VC_MESSAGE, COUNT(*) AS runs, MAX(DT_WHEN) AS last_run,
          SUM(CASE WHEN DT_WHEN >= DATEADD(month,-3,GETDATE()) THEN 1 ELSE 0 END) AS runs_last_3mo
   FROM   dbo.Act_Log               -- or InsertDetailedAct_Log's target on the live Activity DB
   WHERE  VC_TYPE = 'REPORT'        -- adjust column name to the live schema (Trans column)
   GROUP  BY VC_MESSAGE
   ORDER  BY runs DESC;
   ```
   (Note the live target is the **detailed** log via `InsertDetailedAct_Log`; column names differ from the
   spike's 3-column `Act_Log` — confirm the live schema first.) That yields the real
   run-count + last-run + 3-month-frequency the prune list needs. A report with **0 runs over 12 months**
   = safe SKIP; this is the bar to actually prune.
2. **Add the 3 ERROR strings** (`Failed on <X> Report`) to the count for any flaky-Excel report, so a
   report wanted-but-broken isn't mistaken for unused.
3. **Treat shared/no-log reports specially:** R3/R4 share a string (split by handler context or add a
   distinct string in the rebuild); R22 + C-companions have no REPORT row (count via their screen/trans).

Until that export exists, **rebuild the evidence-backed set (R1/R3/R4, R9/R10, C1/C4/C5) first** (which
also matches the M3 build order in `m3-report-inventory.md §4`), and **hold the DECIDE set** rather than
either rebuilding all 22 or pruning blind.

---

### Provenance
- Logging mechanism: `DataModule.pas:6020` (`LogActLog`), `DataModule.dfm:41-43` (`Act_StoredProc` →
  `InsertDetailedAct_Log;1` on `Act_Connection`). All `REPORT` call sites: `MainMenu.pas` (§3 table).
- Spike Activity DB (synthetic): `Activity.dbo.Act_Log` — 78 rows, 2026-06-17→06-21, `VC_TYPE` all
  `INVENTORY`, 0 `REPORT` rows (queries in §1b).
- Real usage fallback: `DB Schema/DailyWorkLog.csv` — 174 rows, 2026-06-19 only, 2 operators, **1 REPORT
  row** (Daily Shipping Assy) + 3 ERROR rows for the same report (§2). Raw client rows NOT reproduced.
