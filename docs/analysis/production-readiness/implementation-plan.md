# Production Implementation Plan — InventorySystem Delphi → Ignition

**Status:** PLANNING deliverable for David's review. Nothing is built here — this is the roadmap.
**Date:** 2026-06-19  ·  **Author:** ignition-architect
**Inputs:** `daily-workflow-usage.md` (real operator-day decode + coverage gap), `cutover-readiness-checkpoint.md`
(the spike foundation), the functional specs under `docs/analysis/{edi,order,reporting,forecasting,admin,...}`,
the version constraint (dev 8.1.52 / prod 8.3), and the headless-authoring realities
(`[[reference-headless-ignition-authoring-limits]]`).

---

## 1. Framing — what the spike proved vs what production requires

### 1.1 The reframe, stated plainly
The cutover-readiness checkpoint called the rebuild "feature-complete." Measured against a **real operator-day**
(`DailyWorkLog.csv`, 174 rows), that is true **only of the data/ledger/master-data foundation plus the Order
worksheet display.** The thing the operator actually does all day — **ASN entry → EDI 856 out → 997/862/810
import → ASN invoicing → order-file generation** — is **~49% of the day's log rows and almost entirely NOT BUILT.**
The two fully-built areas (master-data CRUD, stock ledger) are exactly the two the operator **barely or never
touched** on a normal day.

So: the spike is a **strong, de-risked foundation and outline — NOT production-ready.** It validated the hardest
*data* problems. The revenue-critical *operational* loop is the bulk of the remaining build.

### 1.2 What the spike genuinely PROVED (reuse, don't rebuild)
| Proven asset | Why it de-risks production |
|---|---|
| **Data layer + Named-Query/proc-wrap pattern** | The "wrap the proc, mirror the schema" approach works headless and editable (`gen_perspective_view.py`, master-data CRUD). Every NOT-BUILT capability reuses this exact pattern. |
| **Stock-ledger correctness** (`INV_STOCK_LEDGER` + 4 producers + write-then-post seams + parity proof) | The single hardest *data-integrity* problem (on-hand truth, trigger retirement, double-count avoidance) is solved and dress-rehearsed. |
| **Order commit path** (`INSERT_OpenOrder` → `INV_OPEN_ORDER_INF`, with the SERIALIZABLE/renban-claim design) | The write side of replenishment exists; the order *file* generator can sit on top. |
| **D6 report-proc data layer** (`fn_ManifestCostAt` + 4 window-aware procs + no-overlap guard) | The *data* behind reports + 810/856 pricing is already correct and window-aware. Reporting/EDI build on top of it for **presentation only**. |
| **Cutover mechanics** (4-phase sequence, 11 carries, adversarial review, dress-rehearsal, GO/NO-GO zero-drift gate) | The flip itself is designed and rehearsed. New subsystems extend this sequence, they don't reinvent it. |
| **Landing hub + theme + headless authoring + e2e harness (Playwright)** | The UI shell, auth scaffold, and test rig are live. Every new view inherits them. |

### 1.3 What production REQUIRES that the spike does NOT have
The operational EDI/ASN/order-file daily loop, in full, plus the cross-cutting prod gaps (security, multi-site,
file-share access, alerting, deployment). The Order spike's own checkpoint confirms **Check C (EDI re-scope +
atomic poller) was never started** (`ignition-spike-log.md:232`). The EDI builders (`EDI856Object`/`EDI810Object`),
the inbound X12 poller/parser, the `.ord` fixed-width order-file generator, the renban breakdown algorithm, and
the hot-call path have **zero spike coverage.**

**Bottom line:** the spike answered "can on-hand and pricing be made correct, and can we author/cut over safely?"
Production must now answer "can the operator run a full revenue day in Ignition." That second question is the
remaining ~70% of the build by operator-facing surface area, but it sits on a foundation that has retired the
worst data risks.

---

## 2. Capability → Ignition design (prioritized by production-critical rank)

Each subsystem below is designed as: **Perspective view(s) · Gateway scripts (Project Library) · Named
Queries/proc wrappers · Tags/UDTs.** Priorities follow `daily-workflow-usage.md §6`. The **wrap-the-proc**
principle holds throughout: parallel-run calls the existing procs as-is first; logic ports into Ignition only
where the proc is confirmed buggy (D6 pricing) or the legacy logic was client-side (X12 build, file I/O).

> **Cross-cutting design rule for ALL file I/O (EDI + order files):** there is **no client** anymore. Every
> read/write the legacy did on a mapped drive (`X:\EDIOut\`, `S:\<carrier>\`, the `<EDIIn>` drop) becomes
> **Gateway-side file access** (`system.file.*` on 8.1; `system.file`/Event-Streams on 8.3) against
> gateway-mounted shares whose paths come from the **`sites` table + gateway config**, never a client path.
> Builders that were Delphi `T856EDI`/`T810EDI` objects become **Gateway Jython in a `edi/` Project Library
> package** (8.1-safe Jython 2.7), called by timer/event scripts. This is the single biggest structural shift.

### Rank 1 — ASN management (ASN entry creation) — *NOT BUILT*
The morning revenue keystone. One operator click fans into a header + N manifest-detail inserts, split by
assembly ratio with a "(No Ratio)" remainder row, inside one transaction.

- **Perspective views:**
  - `asn/Create` — line + production-date + sequence-range selector; **ASN sequence-check panel** (Start/Last/Qty,
    the GALC-fed broadcast range) bound to `SELECT_ASNSeq`; a "Create ASN" CTA that previews the split-by-ratio
    detail rows **before** commit (the legacy commits blind; we add a review step).
  - `asn/Manifest` — the manifest-number generator combo (`'7' + yy-digit + MM + DD + 2-char id`, per
    `asn-invoice.md §4.6`) for manual/hot-call adds.
- **Gateway scripts (Project Library `asn/`):**
  - `create_asn(site, line, prod_date, seq_start, seq_end)` — orchestrates the **single transaction**:
    `INSERT_ASNInfo` (OUTPUT `@ASNID`, status `'C'`) → `CalculateASNFRS` (stock decrement; cross-module, body
    unverified — **confirm with delphi-architect**) → N× `INSERT_ASNDetail`. Use the spike's persistent-session
    transaction shim pattern (the one proven for Order's `commitOrders`).
- **Named Queries (`asn/`):** `asn/seq_check` (`SELECT_ASNSeq`), `asn/insert_info` (`INSERT_ASNInfo`),
  `asn/insert_detail` (`INSERT_ASNDetail` — **re-scope dedup to `(IN_ASN_ID, manifest)`**, open Q below),
  `asn/list`, `asn/items` (`SELECT_ASNItems`).
- **Tags/UDTs:** none required (transactional, request-scoped). Optionally a per-line "last ASN sequence" status
  tag for the landing hub.
- **Parity:** the 20-detail + 2-No-Ratio fan-out (rows 4–25 of the log) must reproduce row-for-row against a
  legacy run on the same seq range.

### Rank 2 — EDI outbound 856 + status flip — *NOT BUILT*
The ASN *is* the shipment notice. Daily per line + hot-call (`8HC`).

- **Perspective views:** an **EDI Outbound** panel on `asn/Manage` (recreate / send / unsend buttons; status
  filter C/S/A/R). No new heavy screen — bolt onto the ASN browser.
- **Gateway scripts (Project Library `edi/edi_outbound.py`):**
  - `build_856(site, asn_ein)` — port the `EDI856Object.T856EDI` segment map **verbatim** (`asn-invoice.md §4.2`:
    ISA/GS/ST/BSN/DTM/HL-loop/PRF/LIN/SN1/TD5/CTT/SE/GE/IEA; separators from `sites`; the hardcoded `TD3`
    truck-id and HL parent quirks preserved unless David says otherwise). Replace the inconsistent `fSegCount`
    increment with a **computed** SE01 count (byte-exact SE count is a TEMA-reject risk).
  - File write → `<sites.edi_out_path>\856<date><line>.txt` (normal) or `8HC<...>` (hot-call sentinel
    `StartSeq=-1`). **Transmission stays external** (the VAN mailer), as today.
  - `mark_sent(site, ...)` → `UPDATE_ASNStatus('S')` — **re-scope to per-ASN** (legacy flips ALL `'C'` rows; a
    multi-user/multi-site landmine — open Q).
- **Named Queries:** `edi/report_856` (window-aware `REPORT_EDI856`, made inclusive `>=/<=`),
  `asn/update_status` (`UPDATE_ASNStatus`), `asn/unsend` (`UPDATE_ASNUnsend`).
- **Tags/UDTs:** none. Generation is request-scoped.

### Rank 3 — EDI inbound import / 997 ack — *NOT BUILT (procs exist, no caller)*
11 EDIIMP rows/day, serviced in ~4 bursts. Closes the accept/reject loop. **The single strongest
gateway-Python-service candidate** (`edi-upload.md §6`).

- **Perspective views:** `edi/Inbound` — a status/results table (replacing the blocking Delphi log + busy-wait):
  per-file Found / Trading-Partner / type / accept-reject / archive, backed by a DB-tracked **processed-files**
  table; a "Run poll now" button + the scheduled cadence indicator.
- **Gateway scripts (Project Library `edi/edi_inbound.py` + a Gateway Timer/Scheduled script):**
  - `poll_edi_in(site)` — replaces the manual button + filesystem scan. List `<sites.edi_in_path>`, sniff `ISA`,
    parse the interchange header with a **real X12 split honoring ISA-declared separators** (NOT `copy(,n,len)`
    byte offsets), resolve **`delSL[4]` → site (D1)**, dispatch by transaction-set id, then move to archive only
    **after** the DB side-effect commits (idempotent + crash-safe; fixes the legacy re-ingest + `EDIFileNumber`
    carry-over bugs).
  - Dispatch: **997** → `UPDATE_EINStatus` (the only DB-writing inbound path); **830** → call the shared
    forecast-ingest service (validate DUNS **once**); **862/824** → server-side report render (no Excel);
    **820** → REPORT-ONLY (D12 — site doesn't use it).
  - **997 hardening:** parse **AK9 explicitly** (map `A`/`E`/`P`/`R`), tolerate `AK2/AK3/AK4` detail between AK1
    and AK9 — the legacy blindly reads char 5 of the next segment (`edi-upload.md §4.4`, open Q).
- **Named Queries:** `ein/ack` (`UPDATE_EINStatus`: `@EINType='SH'`→ASN status, else→invoice status, site-scoped),
  plus a new `edi/processed_files` insert/select for the idempotency ledger.
- **Tags/UDTs:** an **EDI inbound health** tag group (last-poll time, files-processed count, unacked-856/810
  count) feeding the landing hub + an alarm (see §3 alerting). This replaces the operator's manual "did the acks
  come in?" vigilance.

### Rank 4 — ASN invoice reconciliation + EDI 810 outbound — *NOT BUILT*
Billing accuracy → payment. Qty corrections feed the invoice; the 810 is the daily-revenue pair to the 856.

- **Perspective views:**
  - `asn/Invoice` — the per-manifest qty edit grid (9 Updates + 1 Delete in the sample day): `UPDATE_ASNItem` /
    `DELETE_ASNItem`. **Scope delete to `IN_ASN_DETAIL_ID`**, not the global manifest key (legacy deletes by
    manifest globally — can hit other ASNs; open Q).
  - `edi/Invoice810` panel — create/recreate/unsend 810; a **priceless-lines pre-check** (the D6 diagnostic
    `REPORT_EDI810_PricelessLines`) surfaced **before** billing so gaps are caught, not shipped as $0.
- **Gateway scripts (`edi/edi_outbound.py`):**
  - `build_810(site, inv_ein)` — port `EDI810Object.T810EDI` segment map verbatim (`asn-invoice.md §4.3`:
    ISA/GS/ST/BIG/IT1-loop/REF/DTM/TDS/CTT/SE/GE/IEA; `M391` broadcast vs `M390` hot-call rule). **Replace the
    hand-rolled `FloatToStr`-split TDS money formatting with explicit integer-cents** (locale-fragile in legacy).
  - **D6 fix:** the 810 line price uses the **same window-aware pricing function** as the 856 (`fn_ManifestCostAt`)
    so 810 and 856 agree. This is the one place we *port logic* rather than wrap a buggy proc.
- **Named Queries:** `inv/create` (`INSERT_INVInfo`+`UPDATE_INVItems`), `inv/unsend` (`UPDATE_INVUnsend` — note it
  **hard-deletes**; reconsider under D3, open Q), `edi/report_810` (window-aware), `asn/update_item`,
  `asn/delete_item` (re-scoped).

### Rank 5 — Order worksheet + order-FILE generation + renban breakdown — *PARTIAL / NOT BUILT*
The replenishment side. 36 ORDERF/ORDERS rows/day. The Order worksheet **display + commit** are spike-built
(Option A faithful calc); the **forecast-fill compute, renban breakdown, and `.ord` file generation** are not.

- **Order worksheet (PARTIAL → finish):** the read-only PhasedGrid + commit path exist (`order/option-a.md`).
  Remaining: the forecast-fill path (the legacy Excel-Cells write that **failed** in the log, rows 46–47) becomes
  the server-side `SIM_OrderSimulation` proc/Named-Query feeding the grid — **no Excel.** This also retires the
  most fragile ERROR path. Calc stays faithful (Option A); B §7 changes deferred behind David's sign-off.
- **Renban breakdown (PARTIAL → build the algorithm):** RenbanGroup CRUD exists; the **breakdown** (FRS lots →
  trailer/pallet count → renban numbers, advancing the group counter `CMWA296`→`297`) does not.
  - **Gateway script `order/renban.py`:** `breakdown(site, group, lots)` — port the `RenbanOrder` algorithm; the
    counter advance MUST be the **atomic claim** (SERIALIZABLE + UPDLOCK/HOLDLOCK, Carry 2) — parallel-run adds
    concurrency the single-user legacy never had. Never emit a blank renban (per `[[project-order-renban-domain]]`).
  - **NQ:** `renban/group` (`SELECT_RenbanGroup`), `renban/bump_counter` (`UPDATE_RenbanGroupCount`, atomic),
    FRS-renban insert/delete.
- **Order-FILE generation (NOT BUILT — the real gap):** per supplier, the legacy writes an Order Sheet `.xls` +
  **3 fixed-width `.ord` text files**: supplier (`S:\CMX\CMWA-0572B.ord`), logistics (`S:\TLDL\...`), and an
  archive copy (`S:\CMX\Archive\...<ts>.ord`). **Downstream supplier/logistics systems parse these byte-exactly.**
  - **Gateway script `order/order_files.py`:** `generate_order_files(site, order_run)` — build the **fixed-width
    positional layout** verbatim (e.g. `0572B6062601 CMWA2964261102Q51000120020260625` = supplier+forecast+
    renban+part+qty(7,zero-pad)+FRS-date(8), `daily-workflow-usage.md §3`), write supplier + logistics + archive
    copies to the `sites`/supplier-master-driven paths. **Replace the Order Sheet Excel** with a server-side
    CSV/PDF (the operator's human-readable copy) — the *machine* `.ord` files are the contract; the `.xls` is not.
  - **Open data-fidelity Q (flagged in the source decode):** the TIRE run wrote supplier+archive but **no
    logistics** file. Confirm whether DUNLOP (`07451`) is configured `logistics=none` (intended) vs a part-type
    code branch — drives whether the generator skips logistics by config or by rule.
  - **NQ:** `order/sheet_source` (the proc result the legacy `OrderFormCreateF` reads).

### Rank 6 — Hot call — *NOT BUILT (depends on Rank 1 + Rank 2)*
Ad-hoc but urgent (line-down avoidance). One operator adds a part/qty against a manifest, then an `8HC` 856 goes
out immediately.

- **Perspective view:** `asn/HotCall` — a proper line grid (the legacy's positional control-pairing is fragile):
  up to 12 assy-part/qty pairs against manifest + line + date.
- **Gateway script:** reuse `create_asn` with the **hot-call branch** (`@StartSeq=-1`, `@EndSeq=-1`,
  `@EIN=SiteEIN+1`, `@HotCall=1` forces insert past dedup), then `build_856` emits the `8HC` filename.
- **NQ:** reuses `asn/insert_info`, `asn/insert_detail` (HotCall=1), `edi/report_856`.

### Rank 7 — Daily reports — *PARTIAL (data migrated, no view; the FAILING path)*
The Daily Shipping Assy Report and the other `REPORT_*` reports. **Every ERROR row in the log was this Excel/OLE
path** (§3). Data procs are largely migrated under D6.

- **Perspective views:** a `reports/` area — a shared date-range/supplier/part param picker (replacing
  `MonthlyReportSelect`) + one parameterized result view per report family (supplier-order, logistics, invoice,
  inventory/lot-location, shipping/ASN, forecast, exception). The ~29 `REPORT_*` procs become Named Queries
  feeding tables.
- **Rendering replacement:** see §3 — **server-side render, no desktop Excel.**

### Rank 8 — Master-data CRUD + stock ledger — *BUILT*
Done (PRs #3–#9, #12). Carry forward; only the multi-site `site_id` scoping (§3) and the auth/role gating remain.

---

## 3. Replacing the fragile Excel/COM layer (feature gap AND reliability win)

**Every one of the 5 ERROR rows in the operator day was Excel OLE:** 3× `Worksheet.PrintOut` failure on the Daily
Shipping Assy Report, and 2× `Cells[...].value` "wrong type / out of range" that **aborted the entire Order
worksheet** on a single bad cell (`daily-workflow-usage.md §4`). The legacy drives a hidden `Excel.Application`
via `CreateOleObject` for **all** reports, the 862/824/820 inbound reports, and the order worksheet/order sheet.
This is the single most fragile dependency in the app — it fails on printer, Excel, and locale conditions the
server cannot control, and one bad cell loses the whole run.

**Production design — eliminate desktop Excel entirely:**

| Legacy Excel use | Production replacement |
|---|---|
| `REPORT_*` → Excel template → `SaveAs .xls` + optional `PrintOut` | **Server-side render.** On **8.3**: the Ignition **Reporting module** (templated, scheduled, PDF/CSV/XLSX export, no client). On **8.1.52 dev**: Perspective table view + **Perspective PDF/CSV export** (`system.perspective`/report-data Named Query → CSV via gateway file write). Print becomes "export PDF + send to a network printer queue server-side," never a client printer driver. |
| Order worksheet / Order Sheet `.xls` (the aborting path) | The worksheet is a **Perspective PhasedGrid** (already spiked, display) fed by `SIM_OrderSimulation`. The human-readable "Order Sheet" copy is a **server-side CSV/PDF**. A bad/NULL cell is a rendered blank, **not** an exception that aborts the run. |
| 862 firm-order `.xls`, 824 advice `.xls`, 820 remittance `.xls` (inbound reports) | Parsed server-side into **real data** (or a rendered report) — never an Excel file. 862 → an order-view; 824 → an exception list (decision-gated: flag the ASN, open Q); 820 → report-only (D12). |
| **The machine-readable `.ord` order files** | **NOT Excel** — these are fixed-width text the legacy already wrote as text. Reproduced byte-exact in `order/order_files.py` (Rank 5). The `.ord` files are a downstream *contract*, kept; only the *Order Sheet `.xls`* is replaced. |

**Reliability framing:** call this out as a **production win, not just parity.** Server-side rendering means
reports cannot be blocked by an offline printer, the order run cannot be lost by one locale-mangled cell, and
there is no per-client Excel install to maintain. The D6 work already moved the report *data* server-side; this
plan finishes the job by moving the *presentation/print* path server-side too.

> **Version note:** prefer the Reporting module on 8.3 (richer, scheduled). Guard it `# IG83-ONLY:` and keep an
> 8.1-runnable Perspective-export fallback so the whole plan stays runnable on the dev box.

---

## 4. Cross-cutting production-readiness gaps (mandatory, not in the daily log)

| Gap | Production design | Carry/decision ref |
|---|---|---|
| **Security — plaintext passwords, single `BIT_ADMIN`** | Retire `INV_USERS` + the `SELECT/INSERT/UPDATE/DELETE_UserInfo` procs. Use an **Ignition User Source** (internal for dev, AD/LDAP/OAuth/SAML for prod). Map `BIT_ADMIN=1`→`Admin` role, else `User`; gate Administration + EDI/order views by **role-based component security**. Introduce the **per-feature permission model** the legacy lacks. Seed from `INV_USERS` once; **force a password reset on first login** (cannot migrate plaintext→hash). | `admin/auth-users.md §6` |
| **Multi-site (D1) — spike is single-site** | Every `INV_*` table gains `site_id`; the INI `[SITE]`/`[INIT]`/`[DISPLAY]` flags + the DB `TSiteInfo` identity (DUNS, EIN, ISA/GS separators, TMM DUNS, delivery-method, supplier code) merge into one **`sites` row**. **Site comes from the session/user, never a client param** (`siteScopedQuery()`). **F1 hazard:** adding `site_id` breaks `SELECT *` `_HIST` triggers unless the `_HIST` table gets the column too (proven on the spike for `INV_PARTS_STOCK_MST`). The EDI `delSL[4]` DUNS lookup (`AD_GetSiteTMMDUNS`, ALC DB) collapses into a per-site `sites` attribute. **Parameterize the hardcoded `IN_ASN_EIN=6440` in `REPORT_EDI856`** (a baked-in site EIN — D1 blocker). **OWN THE `sites` TABLE IN THE INVENTORY DB (new scope, David 2026-06-19):** the authoritative site/line configuration currently lives in the **VehicleOrder** DB (a cross-DB dependency — the order/forecast `LINE` lookups + site config read across databases today). Relocate it to be the canonical **`Inventory.dbo.sites`** table so multi-site config is self-contained and the cross-DB read is eliminated; migration = copy `VehicleOrder.sites`→`Inventory.sites`, repoint every cross-DB reference (order/forecast `LINE`, `AD_GetSite`/`AD_GetSiteTMMDUNS` DUNS lookups), then retire the VehicleOrder copy. Add a **Sites master screen** — the **8th master-data CRUD** (Admin-gated), managing each site row: plant/assembler/supplier codes, DUNS + TMM DUNS, EIN seed/sequence, ISA/GS separators, delivery method, and the `[INIT]`/`[DISPLAY]` flags (`fill_days`, `forecast_usage_compare`, `use_first_production_day`, …). It reuses the proven master-CRUD pattern (combined view + Named Queries + refCount delete-gate); since `site_id` FKs every table, this delete-gate is the strictest of all. | `admin/configuration-site.md §6`, `edi/asn-invoice.md §4.4`, `decisions.md` D1 |
| **File-share access + secrets/paths config** | The legacy mapped-drive UNC paths (`X:\EDIOut\`, `S:\<carrier>\`, `<EDIIn>`) become **gateway-mounted shares**; paths live in `sites`/gateway config, **DB connection strings + credentials in the gateway secret store** (never per-client INI). Gateway needs read/write to these shares — a deployment prerequisite. | `configuration-site.md §6` |
| **Error handling / resilience / alerting** | Replace the blocking Delphi log + busy-wait with **native Ignition Alarming**: alarms on EDI poll failure, unacked 856/810 past a threshold (compliance/payment risk), 824 rejects, order-file write failure, and a stale-inbound watchdog. Notification pipelines (email/SMS) replace the operator manually noticing. The EDI inbound health tag group (Rank 3) drives these. | new |
| **At-least-once / idempotency on EDI I/O** | Outbound: don't mark sent until the file write + status flip commit. Inbound: don't archive until the DB side-effect commits; a DB-tracked **processed-files** ledger replaces move-to-archive as the idempotency key (fixes the legacy re-ingest + `EDIFileNumber` carry-over bugs). **This is NOT Store-and-Forward** (that's tag history only). | `edi-upload.md §4.8` |
| **Atomicity (Carry 1) + renban race (Carry 2) + trigger retirement** | Already tracked in the cutover runbook. Fold source-write + ledger post into one transaction at cutover; atomic renban-counter claim; drop the 13 qty-triggers as the seams become the live `IN_QTY` writer (dress-rehearsed). | `cutover-runbook.md`, checkpoint §3 |
| **Backup/restore** | Gateway config + project backup (`.gwbk`) on a schedule; DB backup remains SQL Server's job. Document a restore runbook. Spike infra (Colima/docker mssql) is dev-only; prod is the existing SQL Server. | new |
| **Deployment topology** | One Gateway (prod 8.3) serving Perspective; SQL Server as today. Decide **redundancy/failover** (Ignition redundant gateway) given EDI is revenue-critical — recommend a redundant pair for prod, single gateway acceptable for parallel-run. | new (open Q) |
| **EDI EIN / site-scoping hooks** | `AD_UpdateEIN` (the EIN counter, ALC DB) becomes a **per-site sequence**; the ALC `AD_GetSite`/`AD_GetSiteTMMDUNS` bodies must be confirmed before relying on them. EIN is the outbound control number — it must be allocated atomically per site. | `asn-invoice.md §4.4`, open Q |
| **Audit trail (`LogActLog`)** | The per-action audit trail (every ASN/EDI/order action) → an Ignition **audit profile** or an `audit_log` table written by the gateway services, carrying the Perspective username + client address. Keep the human-readable EDI per-file trail. | `auth-users.md §6` |

---

## 5. Phased roadmap to production

Sized for a **solo dev**. Effort is rough order-of-magnitude (dev-weeks), assuming the spike foundation is reused.
"Parallel-run / shadow" = both Delphi and Ignition hit the same SQL Server, procs mediate. "Hard cutover" = a
quiesced flip.

### Definition of "production-ready" (the acceptance bar)
1. **Full-day operability:** the operator can run an entire production day in Ignition (ASN build → 856 out →
   inbound acks → ASN invoice → 810 out → order run → order files → daily report → hot-call) **matching the
   legacy**, with **zero Excel/OLE** in the path.
2. **EDI accepted by TEMA:** generated 856/810 files are **byte-diff'd against known-good legacy files** AND
   **997-accepted by TEMA** in a real exchange.
3. **Parallel-run parity:** for a fixed sample, Ignition output (ASN details, EDI files, `.ord` files, order
   numbers, report numbers) matches a same-data legacy run.
4. **Security:** Ignition User Source + role gating live; no plaintext passwords; per-feature permissions.
5. **Multi-site:** `site_id` end-to-end; session-scoped; F1-safe; no hardcoded site EIN.
6. **Resilience:** native alarming on the EDI/order failure modes; idempotent EDI I/O; backup runbook.

### Milestones

**M1 — The revenue-critical ASN → 856 → 810 daily loop (parallel-run shadow). ~5–7 dev-weeks.**
Ranks 1, 2, 4 + the outbound half of the loop.
- ASN entry (create transaction + sequence check + split-by-ratio review), ASN invoice reconciliation, 856 + 810
  builders (Gateway Jython, byte-exact), status flip/unsend, D6 window-aware pricing shared by 810/856.
- **Mode:** **shadow** — Ignition writes ASN/INV rows + EDI files to a **separate `EDIOut` staging dir**; diff
  against legacy files; do NOT transmit Ignition's files until byte-parity + a TEMA test-accept. Legacy stays the
  system of record.
- **Depends on:** ASN-detail dedup scope decision; `CalculateASNFRS` body; ALC `AD_*` bodies; the persistent-tx
  shim (proven).
- **Gate:** byte-diff 856/810 vs legacy on a sample day; TEMA test-accept of an Ignition-built 856 and 810.

**M2 — EDI inbound + order-file generation. ~5–7 dev-weeks.**
Ranks 3, 5, 6.
- `edi_inbound.py` poller (X12 parse, `delSL[4]`→site, 997 ack with AK9, 830→forecast, 862/824 server-render,
  820 report-only), processed-files idempotency ledger.
- Renban breakdown algorithm (atomic counter), order-file generator (`.ord` byte-exact to supplier/logistics/
  archive), order worksheet forecast-fill finished (no Excel), hot-call path.
- **Mode:** inbound can **shadow** (parse + log, write 997 status to a shadow column first, then enable the real
  `UPDATE_EINStatus`); order files written to a staging dir and byte-diff'd before pointing at the real shares.
- **Depends on:** M1 (hot-call needs ASN+856); `AD_GetSpecialDate` body (calendar walk, blocks order forecast-
  fill); the TIRE-logistics config Q; AK9 semantics Q.
- **Gate:** 997 round-trip flips the right ASN/INV status; `.ord` files byte-diff vs legacy and parse downstream.

**M3 — Reporting + Excel-layer retirement. ~3–4 dev-weeks.**
Rank 7 + §3.
- Server-side render of the report families (8.3 Reporting module / 8.1 Perspective-export fallback); retire the
  last Excel/OLE paths; the Daily Shipping Assy Report (the FAILING path) renders + "prints" server-side.
- **Mode:** purely additive; reports are read-only — can run live alongside legacy immediately.
- **Gate:** report numbers match legacy for a sample; no Excel in any path.

**M4 — Security / multi-site / hardening. ~4–6 dev-weeks.**
§4.
- User Source + roles + per-feature gating; `site_id` end-to-end (schema surgery + F1-safe `_HIST` + session
  scoping + parameterize the hardcoded EIN); **relocate the `sites` table VehicleOrder→Inventory DB + build the
  Sites master screen (8th master, Admin-gated) + repoint the cross-DB `LINE`/DUNS reads**; native alarming +
  notification pipelines; secrets/paths to gateway config + secret store; backup runbook; redundancy decision;
  transactional DATAPURGE scoped by site.
- **Mode:** schema surgery (`site_id`) is the riskiest — stage it additively (nullable, default current site)
  during parallel-run, enforce NOT-NULL at cutover. Auth flip is a hard cutover for operators (they get Ignition
  logins).
- **Gate:** the 6 production-ready criteria above all green.

**M5 — Cutover. ~1–2 dev-weeks (mostly the rehearsed flip + parallel-run soak).**
- Extend the existing 4-phase cutover sequence (Phase A/B/C/D) to additionally cut over the EDI/ASN/order-file
  subsystems: stop the legacy as the EDI system of record, repoint transmission to Ignition's `EDIOut`, retire
  the legacy app. Soak parallel-run for an agreed window (e.g. 2 weeks) with both producing EDI, transmitting
  only Ignition's, legacy as hot fallback.
- **Mode:** **the EDI transmission flip is a HARD cutover** (only one app can be the system of record for the
  TEMA exchange at a time — you cannot have both transmitting 856/810). Everything up to it parallel-runs.
- **Gate:** an agreed parallel-run soak passes (every day's EDI accepted by TEMA from Ignition, parity green),
  then flip; legacy retained as fallback for the soak window.

**Rough total: ~18–26 dev-weeks** for a solo dev, foundation reused. M1 is the highest-value, highest-risk chunk
and should be sequenced first; M3 can slip in parallel with M1/M2 since it's additive read-only.

### What parallel-runs/shadows vs what is a hard cutover
- **Parallel-run / shadow (most of it):** ASN/invoice writes (procs dedup), inbound parse-to-shadow, all
  reporting (read-only), order-file staging, `site_id` additive schema. Both apps share the DB; procs mediate.
- **Hard cutover (the few):** (1) **EDI transmission** — only one app transmits 856/810 to TEMA; (2) **the qty-
  trigger drop / seam-as-live-writer** (already designed/rehearsed in Phase C); (3) **operator auth** (they
  switch to Ignition logins). Everything else can shadow until these flip.

---

## 6. Open questions for David (decisions needed before / early in the build)

**Blocking M1:**
1. **ASN-detail dedup/delete scope.** Legacy `INSERT_ASNDetail` dedups (and `DELETE_ASNItem` deletes) by
   `VC_MANIFEST_NUMBER` **globally**. Scope to `(IN_ASN_ID, manifest)` + per-site under D1? (Recommend: yes.)
2. **`UPDATE_ASNStatus('S')` flips ALL `'C'` ASNs** (no id filter). Per-ASN send, or intentional batch? (Multi-
   user/multi-site breaks the batch assumption.)
3. **D6 window boundary inclusivity** — confirm `>=`/`<=` (inclusive), and that 810 must price identically to 856.
4. **ALC-DB procs** (`AD_GetSite`, `AD_GetSiteTMMDUNS`, `AD_UpdateEIN`, `CalculateASNFRS`): confirm bodies + that
   the EIN counter is the authoritative outbound control number to replicate as a **per-site sequence**.
5. **Unsend semantics** — `UPDATE_INVUnsend` **hard-deletes** the invoice header. Make it a recoverable status
   revert per D3?

**Blocking M2:**
6. **997 AK9 semantics** — map AK9 group codes (`A`/`E`/`P`/`R`) to distinct statuses and tolerate `AK2/AK3/AK4`
   detail between AK1 and AK9? (Recommend: yes — the legacy is fragile here.)
7. **`delSL[4]` exact X12 element + DUNS provenance** — confirm against a real inbound ISA which element it lands
   on (code index is authoritative), and that `AD_GetSiteTMMDUNS` is the allow-list to migrate to `sites`.
8. **TIRE order-file logistics omission** — is DUNLOP (`07451`) configured `logistics=none` (intended) or is there
   a part-type code branch skipping logistics? Drives the generator's skip-by-config vs skip-by-rule design.
9. **`AD_GetSpecialDate` body + status domain** (ALC `TireOrder` DB) — needed for the order forecast-fill calendar
   walk; do not re-derive the O/X/holiday rules.
10. **824 application-advice action** — auto-flag/reject the named ASN, or operator report only?
11. **EDI ingest trigger + cadence** — Gateway scheduled poll (recommend) vs on-demand; cadence; one shared
    inbound drop fanned out by DUNS vs per-site dirs.

**Blocking M4 / M5:**
12. **Roles granularity** — Admin/User split enough, or a finer per-feature permission set (EDI-only, receiving-
    only operators)?
13. **First-login password reset** — OK to force every user to set a new password at first Ignition login (we will
    NOT preserve plaintext)?
14. **Deployment redundancy** — redundant Ignition gateway pair for prod (recommend for revenue-critical EDI), or
    single gateway?
15. **Order calc changes (B §7 C1–C6)** — stay deferred behind sign-off (recommend), build faithful (Option A)?
16. **Parallel-run soak window** — how long must Ignition produce TEMA-accepted EDI alongside legacy before the
    transmission hard-cutover? (Recommend ≥2 weeks of clean days.)
17. **DATAPURGE retention scope + per-site** — keep the legacy 3-table scope or extend; per-site or global policy?

---

## Confidence / assumptions to flag
- **Source behavior:** ASN/EDI/order legacy semantics are taken from the functional specs
  (`docs/analysis/edi/`, `order/`, `reporting/`) which themselves flag several proc bodies as **inferred from
  call sites, not fully read** (`REPORT_EDI856`/`810` branches, `SELECT_ASNSeq`, FRS-renban inserts,
  `CalculateASNFRS`, all ALC `AD_*`). **Confirm these with delphi-architect before porting** — do not let a clean
  Ignition design guess legacy behavior.
- **Effort sizing** assumes the spike's NQ-generation + persistent-tx shim + e2e harness are reused and that the
  hardest data problems (ledger, pricing, cutover) stay solved. New X12/file-I/O work is the dominant cost.
- **The EDI subsystem is the strongest Gateway-Python-service candidate** in the system; it deliberately does
  **not** live in Perspective bindings or Named Queries — those wrap the DB procs only.
