# REPORT_EDI856 — Data-Feed Analysis (M1 Rank-2 outbound 856)

What `REPORT_EDI856` reads to populate the 856 segments, so the rebuild reimplements the SAME read as a
**side-effect-free SELECT** (NOT a wrap — the proc self-flips status). Proven on live `Inventory` (mssql-spike,
read-only; 2550 ASN headers / 39,707 detail rows). Sources: live `REPORT_EDI856` (OBJECT_DEFINITION) vs
`CreateInventory.sql` (~:3623-3697; self-flip :3695; `6440` :3683) vs D6 copy `spike-report-procs-d6.sql:95-131`.

## The SELECT shape (what the rebuild reimplements)
Lineage: `INV_ASN_MST a JOIN INV_ASN_DETAIL_MST d ON a.IN_ASN_ID=d.IN_ASN_ID
JOIN INV_MANIFEST_COST_MST m ON d.VC_ASSY_PART_NUMBER=m.VC_ASSY_PART_NUMBER_CODE
JOIN INV_FORECAST_DETAIL_INF f ON d.VC_ASSY_PART_NUMBER=f.VC_ASSY_PART_NUMBER_CODE`.

9-column projection → the 856 fields:
| col | source | → segment |
|---|---|---|
| Manifest | `d.VC_MANIFEST_NUMBER` varchar(8) | PRF (HL Order) |
| PartNumber | `d.VC_ASSY_PART_NUMBER` varchar(12) | LIN BP |
| UnitPrice | `m.MO_PRICE` | **NOT EMITTED** (referenced 0× in EDI856Object.pas) |
| ShipQty | `d.IN_QTY` int | SN1 |
| PickUpDate | `a.VC_PRODUCTION_DATE` varchar(8) yyyymmdd | BSN/DTM + filename |
| Kanban | `f.VC_ASSY_KANBAN_NUMBER` | LIN RC |
| SiteEIN | `a.IN_ASN_EIN` int | control numbers |
| StartSeq | `a.VC_START_SEQ_NUMBER` varchar(4) (header; unqualified in proc) | filename/hot-call branch |
| LineName | `a.VC_LINE_NAME` (only in the `@EIN=0` preview branch) | — |

Cardinality: one row per (header × surviving detail line), then a `GROUP BY` over all 9 columns (poor-man's
DISTINCT). `@EIN=0` branch: `WHERE VC_ASN_STATUS='C'` + cost-window predicate, **no mutation** (preview).
`@EIN<>0` branch: `WHERE IN_ASN_EIN=6440`, cost-window commented out, then the self-flip.

## The self-flip — do NOT wrap (`:3695`, `@EIN<>0` branch only)
`UPDATE INV_ASN_MST set VC_ASN_STATUS='S' WHERE IN_ASN_EIN=@EIN`. Reading the feed mutates state → the rebuild
reads via a pure SELECT and does the C→S flip **separately, per-ASN, at send**.
**Worse sibling — `UPDATE_ASNStatus` (`:1695`, called DataModule.pas:5077):** `UPDATE INV_ASN_MST SET
VC_ASN_STATUS='S' WHERE VC_ASN_STATUS='C'` — ALL EINs, ALL sites, no filter. The biggest cross-site
"mark-everything-sent" hazard. The rebuild must NEVER use it (per-ASN flip only).

## EIN + the 6440 hardcode
`@EIN int=0`. `=0` → preview (no mutation). `<>0` → send: its `WHERE` is the **literal `6440`, not `@EIN`**
(`:3683`) → returns EIN-6440 rows regardless of the EIN passed, then flips status for the *actual* `@EIN`
(the two disagree). Live: `WHERE IN_ASN_EIN=6440` → 0 rows on this snapshot (latent). D6 already fixed it to
`= @EIN`. `IN_ASN_EIN` is int NOT NULL, range 3502-9057, **unique per ASN** (an allocated sequence, not a site
code). No proc touches `IN_EIN_SEQ`; the app does `fEIN+1`. **Rebuild:** allocate EIN per-site at send from
`INV_SITES.IN_EIN_SEQ` (atomic, site-scoped); filter the read by the passed `@EIN`; never a literal.

## UPDATE_EINStatus — the inbound ack, site-scoping gap (`:1711`, DataModule.pas:6760)
`if @EINType='SH' UPDATE INV_ASN_MST SET VC_ASN_STATUS=@EINStatus WHERE IN_ASN_EIN=@EIN` (no site filter).
Under per-site EIN reuse an ack for site X's EIN N also flips site Y's. Rebuild: scope by `site_id` (or key
acks off a globally-unique shipment id).

## DECISIONS the 856 build needs (proven artifacts of the legacy read)
1. **The 856 carries NO price** — drop `UnitPrice`/`MO_PRICE` from the emit. But the cost JOIN currently
   *filters*: **25,351 / 39,707 (64%) detail lines have no covering cost window and silently vanish** (likely a
   sparse-snapshot artifact, but the principle stands: should the 856 report a shipped line that has no manifest
   cost?). → **DECISION A: drop the cost join (report all shipped lines) vs keep it (active/billable parts only).**
2. **Forecast JOIN (Kanban):** inner join drops **24,357 / 39,707** lines whose part has no forecast row; no
   index/uniqueness on the forecast part code (fan-out structurally possible). → **DECISION B: LEFT vs INNER for
   Kanban + cap to one row (`OUTER APPLY TOP 1`).**
3. **GROUP BY collapse:** distinct detail rows sharing all 9 columns merge — **14,356 / 28,712 (half) collapse**.
   A correct ASN arguably SUMs `IN_QTY` over merged lines rather than dropping them (changes shipped qty). →
   **DECISION C: SUM qty on collapse vs legacy drop.**
4. **Segment terminator:** legacy emits CRLF only (no `~`). → **DECISION D: match the TEMA-accepted legacy
   output (CRLF-only) for parallel-run parity, vs emit standard-X12 `~`.** (Parity argues match-legacy unless
   TEMA requires `~`.)
5. **Filename:** create vs recreate use different `copy` offsets → pick one deterministic pattern (DECISION E).

## What the rebuild's 856 read MUST reproduce + traps
The 9-col read, one row per surviving (header × detail); EXCLUDE the self-flip (pure SELECT; flip C→S per-ASN
at send); NEVER the `6440` literal or the blanket `UPDATE_ASNStatus`; EIN per-site from `INV_SITES.IN_EIN_SEQ`
at send; site-scope `UPDATE_EINStatus`; resolve DECISIONS A-E before/at build. Edge cases: `VC_ASN_STATUS`
varchar(1) NOT NULL, snapshot is 100% 'A' (no 'C'/'S' to observe the preview branch); `IN_ASN_EIN` NOT NULL
(no NULL EIN; define the pre-send sentinel under EIN-at-send); ASN with no detail rows → inner join drops it
(caller guards on RecordCount>0).
