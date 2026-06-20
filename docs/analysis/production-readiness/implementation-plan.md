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
  - `build_856(site, asn_id)` — port the `EDI856Object.T856EDI` segment map **verbatim** (`asn-invoice.md §4.2`:
    ISA/GS/ST/BSN/DTM/HL-loop/PRF/LIN/SN1/TD5/CTT/SE/GE/IEA; separators from `sites`; the hardcoded `TD3`
    truck-id and HL parent quirks preserved unless David says otherwise). Replace the inconsistent `fSegCount`
    increment with a **computed** SE01 count (byte-exact SE count is a TEMA-reject risk).
  - **DO NOT wrap `REPORT_EDI856` to feed this builder.** Verified (`CreateInventory.sql:3695`,
    `/tmp/inv_utf8.sql:3695`): the proc's `@EIN<>0` branch executes
    `UPDATE INV_ASN_MST set VC_ASN_STATUS='S' WHERE IN_ASN_EIN=@EIN` — calling it to *read* the 856 data flips
    real ASN status to `'S'` for every ASN sharing that EIN, in the shared DB, before any file is written. Under
    parallel-run/shadow that makes the legacy app see those ASNs as already-sent and skip its own 856
    (double-source corruption). **Resolution: build the 856 from a NEW pure-SELECT** (`edi/asn_856_data` NQ — the
    SELECT lifted out of the `@EIN<>0` branch, with the embedded `UPDATE` dropped and the hardcoded `6440`/EIN
    predicate parameterized per-site). The builder ports the `EDI856Object` segment logic over that SELECT.
  - File write → `<sites.edi_out_path>\856<date><line>.txt` (normal) or `8HC<...>` (hot-call sentinel
    `StartSeq=-1`). **Transmission stays external** (the VAN mailer), as today.
  - `mark_sent(site, asn_id, ein)` → the **decoupled per-ASN status flip** (Q2): set `VC_ASN_STATUS='S'` for that
    one ASN **only after** its 856 file write + transmit commits (the at-least-once rule, §4). This is the flip
    that BLOCKER-1 removes from the report proc — it now lives explicitly here, per-ASN, never as a side effect of
    reading data. Re-scope the underlying update to `WHERE IN_ASN_ID=@ASNID AND VC_ASN_STATUS='C' AND site_id=@site`.
- **Named Queries:** `edi/asn_856_data` (the **pure SELECT** lifted from `REPORT_EDI856`, embedded UPDATE
  removed, window-aware inclusive `>=/<=`, EIN/site parameterized — **not** a wrap of the mutating proc),
  `asn/update_status` (the re-scoped per-ASN `UPDATE_ASNStatus`), `asn/unsend` (`UPDATE_ASNUnsend`).
- **Tags/UDTs:** none. Generation is request-scoped.
- **Principle (applies to ALL report procs):** the wrap-the-proc shadow strategy is **only valid for
  pure-SELECT** procs. Any proc that mutates state — status flips, inserts, the embedded `UPDATE`s in
  `REPORT_EDI856`/`REPORT_EDI810` — **must be reimplemented, not wrapped.** Before relying on any other `REPORT_*`
  or `EDI*` proc as a shadow data feed, audit its body for hidden writes (see §4 "wrap-vs-reimplement audit").

### Rank 3 — EDI inbound import / 997 + 824 ack (M1) + 830/862 forecast import (M2) — *NOT BUILT (procs exist, no caller)*
11 EDIIMP rows/day, serviced in ~4 bursts. Closes the accept/reject loop. **The single strongest
gateway-Python-service candidate** (`edi-upload.md §6`).

> **Milestone split (BLOCKER-3 boundary redraw):** the **minimal inbound that closes the outbound loop — 997
> functional-group ack + 824 application-advice — is M1** (without it, an Ignition-built ASN can reach `'S'` but
> never `'A'`/`'R'`, so M1's "TEMA-accept" gate is unreachable). The **remaining inbound — 830/862 forecast
> import + the forecast-fill — is M2.** The poller shell (`poll_edi_in`, X12 parse, DUNS→site routing, archive,
> processed-files ledger) is built in M1 (it must run to ingest 997/824) and **extended** in M2 to dispatch
> 830/862. The dispatch table below tags each transaction set with its milestone.

- **Perspective views:** `edi/Inbound` — a status/results table (replacing the blocking Delphi log + busy-wait):
  per-file Found / Trading-Partner / type / accept-reject / archive, backed by a DB-tracked **processed-files**
  table; a "Run poll now" button + the scheduled cadence indicator.
- **Gateway scripts (Project Library `edi/edi_inbound.py` + a Gateway Timer/Scheduled script):**
  - `poll_edi_in()` — replaces the manual button + filesystem scan. **One gateway poller serves all sites**
    (single gateway, Q14 + multi-site, D1). List the shared `<edi_in drop>`, sniff `ISA`, parse the interchange
    header with a **real X12 split honoring ISA-declared separators** (NOT `copy(,n,len)` byte offsets), resolve
    the file's sender **DUNS** (`delSL[4]`) and **match it against ALL configured `sites` rows' DUNS, routing the
    file to the matching site** (D1). A file whose DUNS matches **no** configured site is **dropped/quarantined**
    (logged + alarmed, not silently consumed). Then dispatch by transaction-set id, then move to archive only
    **after** the DB side-effect commits (idempotent + crash-safe; fixes the legacy re-ingest + `EDIFileNumber`
    carry-over bugs).
  - Dispatch (milestone-tagged): **997 (M1)** → the site-scoped ack flip (see `ein/ack` below — the only
    DB-writing inbound path in M1); **824 (M1)** → auto-flag the named ASN rejected + raise the home-screen alarm
    (Q10); **830 (M2)** → call the shared forecast-ingest service (validate DUNS **once**); **862 (M2)** →
    server-side report render (no Excel); **820** → REPORT-ONLY (D12 — site doesn't use it).
  - **997 hardening:** parse **AK9 explicitly** (map `A`/`E`/`P`/`R`), tolerate `AK2/AK3/AK4` detail between AK1
    and AK9 — the legacy blindly reads char 5 of the next segment (`edi-upload.md §4.4`, open Q).
- **Named Queries:** `ein/ack` (`UPDATE_EINStatus`: `@EINType='SH'`→ASN status, else→invoice status). **Must be
  site-scoped (BLOCKER-2).** Verified the proc keys on `WHERE IN_ASN_EIN=@EIN` / `IN_INV_EIN=@EIN` **alone**
  (`/tmp/inv_utf8.sql:1724/1728`) — with Q4's per-site EIN sequences, site A and site B can both hold EIN 9069, so
  a 997 for site A would flip **both** sites' ASN 9069. **Add `site_id` to the WHERE** so the flip scopes to
  `(site_id, EIN)`; the inbound file already resolved to a site by DUNS (Q7/Q11), so pass that `site_id` through.
  (General rule: every EIN-keyed update must be re-checked for the same un-scoped pattern — see §4.) Plus a new
  `edi/processed_files` insert/select for the idempotency ledger.
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
  - `build_810(site, inv_id)` — port `EDI810Object.T810EDI` segment map verbatim (`asn-invoice.md §4.3`:
    ISA/GS/ST/BIG/IT1-loop/REF/DTM/TDS/CTT/SE/GE/IEA; `M391` broadcast vs `M390` hot-call rule). **Replace the
    hand-rolled `FloatToStr`-split TDS money formatting with explicit integer-cents** (locale-fragile in legacy).
  - **DO NOT wrap `REPORT_EDI810` to feed this builder** — it has the same self-mutation as 856 (BLOCKER-1 twin,
    NIT-2 **confirmed**). Verified (`/tmp/inv_utf8.sql`, `REPORT_EDI810` `@EIN<>0` branch):
    `UPDATE INV_INV_MST SET VC_INV_STATUS='S' WHERE IN_INV_EIN=@EIN`. Build the 810 from a NEW pure-SELECT
    (`edi/inv_810_data` NQ — the SELECT lifted out, embedded UPDATE dropped, EIN/site parameterized).
  - **Decoupled per-INV status flip (Q2 parallel for 810):** set `VC_INV_STATUS='S'` for the one invoice only
    after its 810 file write + transmit commits — never as a side effect of reading data.
  - **D6 fix:** the 810 line price uses the **same window-aware pricing function** as the 856 (`fn_ManifestCostAt`)
    so 810 and 856 agree. This is one of the places we *port logic* rather than wrap a buggy proc.
- **Named Queries:** `inv/create` (`INSERT_INVInfo`+`UPDATE_INVItems`), `inv/unsend` (the **in-place**
  recreate-flag rebuild of `UPDATE_INVUnsend` per Q5/D3 — no hard-delete), `edi/inv_810_data` (the **pure SELECT**
  lifted from `REPORT_EDI810`, embedded UPDATE removed, window-aware, EIN/site parameterized — **not** a wrap of
  the mutating proc), `inv/mark_sent` (the decoupled per-INV flip), `asn/update_item`, `asn/delete_item`
  (re-scoped to `(site_id, IN_ASN_ID, manifest)`).

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
- **NQ:** reuses `asn/insert_info`, `asn/insert_detail` (HotCall=1), `edi/asn_856_data` (the pure-SELECT feed).

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
| **Security — plaintext passwords, single `BIT_ADMIN`** | Retire `INV_USERS` + the `SELECT/INSERT/UPDATE/DELETE_UserInfo` procs. Use an **Ignition User Source** (internal for dev, AD/LDAP/OAuth/SAML for prod). Map `BIT_ADMIN=1`→`Admin` role, else `User`; gate Administration + EDI/order views by **role-based component security**. **Admin/User two-role split for now** (Q12 — the finer per-feature model, e.g. EDI-only/receiving-only, is deferred). Seed from `INV_USERS` once; **force a password reset on first login** (cannot migrate plaintext→hash). | `admin/auth-users.md §6` |
| **Multi-site (D1) — spike is single-site** | Every `INV_*` table gains `site_id`; the INI `[SITE]`/`[INIT]`/`[DISPLAY]` flags + the DB `TSiteInfo` identity (DUNS, EIN, ISA/GS separators, TMM DUNS, delivery-method, supplier code) merge into one **`sites` row**. **Site comes from the session/user, never a client param** (`siteScopedQuery()`). **F1 hazard — full blast radius enumerated (bake into M4, don't discover mid-build):** adding `site_id` to a base table breaks any `INSERT ... _HIST SELECT * FROM inserted/deleted` trigger unless the `_HIST` table gains the same column in lockstep. **7 confirmed `SELECT *` history triggers across 3 tables** (verified `/tmp/inv_utf8.sql`): FORECAST — `INSERT INTO INV_FORECAST_DETAIL_INF_HIST SELECT *` at lines **2664** (del), **3440** (ins), **3617** (ins); PARTS_STOCK — `INV_PARTS_STOCK_MST_HIST SELECT *` at **4107** (del), **4269** (ins); OPEN_ORDER — `INV_OPEN_ORDER_INF_HIST SELECT *` at **5475** (del), **7496** (ins). (`INV_ASSY_BUILD_HIST` is safe — explicit `VALUES`, not `SELECT *`.) M4 pre-flight: add `site_id` to each of the 3 `_HIST` tables before/with the base-table change; a mismatched-column `SELECT *` throws at runtime and halts the base DML. (Bonus: re-audit the remaining triggers — 25 total — for any other `SELECT *` on a table gaining `site_id`.) The EDI `delSL[4]` DUNS lookup (`AD_GetSiteTMMDUNS`, ALC DB) collapses into a per-site `sites` attribute. **Parameterize the hardcoded `IN_ASN_EIN=6440` in `REPORT_EDI856`** (a baked-in site EIN — D1 blocker). **OWN THE `sites` TABLE IN THE INVENTORY DB (new scope, David 2026-06-19):** the authoritative site/line configuration currently lives in the **VehicleOrder** DB (a cross-DB dependency — the order/forecast `LINE` lookups + site config read across databases today). Relocate it to be the canonical **`Inventory.dbo.sites`** table so multi-site config is self-contained and the cross-DB read is eliminated; migration = copy `VehicleOrder.sites`→`Inventory.sites`, repoint every cross-DB reference (order/forecast `LINE`, `AD_GetSite`/`AD_GetSiteTMMDUNS` DUNS lookups), then retire the VehicleOrder copy. **M4 PRE-TASK (verify-before-relocate) — `VehicleOrder.sites` existence + readers are UNVERIFIED in the InventorySystem source** (the adversarial review found no `FROM sites`/`dbo.sites` reference in any `.pas`/`.dfm`; site identity in *this* app is read from INI per `SiteInfo.pas`). David asserts the table lives in VehicleOrder — so **before relocating/retiring it, confirm with delphi-architect: (1) the real table name + which DB/schema it actually lives in, (2) the exact cross-DB references to repoint, and (3) which sibling apps read it** (GALC/MES in `Delphi-VCL-Components` — the Q9 calendar decision already establishes VehicleOrder is shared across Inventory/GALC/MES, so a retire step has cross-app blast radius). If it has sibling readers, "retire the VehicleOrder copy" becomes "keep shared / dual-read," not a clean move. Do not break siblings. (This is verify-before-relocate, not a contradiction of David's call.) Add a **Sites master screen** — the **8th master-data CRUD** (Admin-gated), managing each site row: plant/assembler/supplier codes, DUNS + TMM DUNS, EIN seed/sequence, ISA/GS separators, delivery method, and the `[INIT]`/`[DISPLAY]` flags (`fill_days`, `forecast_usage_compare`, `use_first_production_day`, …). It reuses the proven master-CRUD pattern (combined view + Named Queries + refCount delete-gate); since `site_id` FKs every table, this delete-gate is the strictest of all. | `admin/configuration-site.md §6`, `edi/asn-invoice.md §4.4`, `decisions.md` D1 |
| **File-share access + secrets/paths config** | The legacy mapped-drive UNC paths (`X:\EDIOut\`, `S:\<carrier>\`, `<EDIIn>`) become **gateway-mounted shares**; paths live in `sites`/gateway config, **DB connection strings + credentials in the gateway secret store** (never per-client INI). Gateway needs read/write to these shares — a deployment prerequisite. | `configuration-site.md §6` |
| **Error handling / resilience / alerting** | Replace the blocking Delphi log + busy-wait with **native Ignition Alarming**: alarms on EDI poll failure, unacked 856/810 past a threshold (compliance/payment risk), 824 rejects, order-file write failure, and a stale-inbound watchdog. Notification pipelines (email/SMS) replace the operator manually noticing. The EDI inbound health tag group (Rank 3) drives these. | new |
| **At-least-once / idempotency on EDI I/O** | Outbound: don't mark sent until the file write + status flip commit. Inbound: don't archive until the DB side-effect commits; a DB-tracked **processed-files** ledger replaces move-to-archive as the idempotency key (fixes the legacy re-ingest + `EDIFileNumber` carry-over bugs). **This is NOT Store-and-Forward** (that's tag history only). | `edi-upload.md §4.8` |
| **Atomicity (Carry 1) + renban race (Carry 2) + trigger retirement** | Already tracked in the cutover runbook. Fold source-write + ledger post into one transaction at cutover; atomic renban-counter claim; drop the 13 qty-triggers as the seams become the live `IN_QTY` writer (dress-rehearsed). | `cutover-runbook.md`, checkpoint §3 |
| **Backup/restore** | Gateway config + project backup (`.gwbk`) on a schedule; DB backup remains SQL Server's job. Document a restore runbook. Spike infra (Colima/docker mssql) is dev-only; prod is the existing SQL Server. | new |
| **Deployment topology** | One Gateway (prod 8.3) serving Perspective; SQL Server as today. Decide **redundancy/failover** (Ignition redundant gateway) given EDI is revenue-critical — recommend a redundant pair for prod, single gateway acceptable for parallel-run. | new (open Q) |
| **EDI EIN / site-scoping hooks** | `AD_UpdateEIN` (the EIN counter, ALC DB) becomes a **per-site sequence** (Q4 — the VAN control number per site/DUNS); the ALC `AD_GetSite`/`AD_GetSiteTMMDUNS` bodies must be confirmed before relying on them. EIN is the outbound control number — it must be allocated atomically per site. **BLOCKER-2 — site-scope every EIN-keyed update:** `UPDATE_EINStatus` keys on `WHERE IN_ASN_EIN=@EIN` / `IN_INV_EIN=@EIN` **alone** (verified `/tmp/inv_utf8.sql:1724/1728`) — with per-site EIN sequences (Q4), two sites can hold the same EIN, so an unscoped flip acks the wrong site. **Add `site_id` to the WHERE** → `(site_id, EIN)`; the inbound 997/824 already resolves a site by DUNS (Q7/Q11) so the `site_id` is in hand. **Audit ALL EIN-keyed updates** for the same un-scoped pattern before per-site EIN goes live (this row + `REPORT_EDI856`/`810`'s now-removed embedded flips were the same class of bug). | `asn-invoice.md §4.4`, decisions Q4, BLOCKER-2 |
| **Wrap-vs-reimplement audit (report procs)** | The shadow strategy's "wrap the proc as-is" is **only safe for pure-SELECT procs.** Two supposed read procs mutate state: `REPORT_EDI856` (`:3695`, flips ASN→`'S'`) and `REPORT_EDI810` (flips INV→`'S'`) — both are now **reimplemented as pure-SELECT NQs with the embedded UPDATE removed** and the status flip decoupled to send-commit (Rank 2/4, Q2). **Before relying on any other `REPORT_*`/`EDI*` proc as a shadow data feed, read its body for hidden `UPDATE`/`INSERT`/`DELETE`.** A proc that writes cannot be wrapped read-only during parallel-run. | §2 Rank 2/4, BLOCKER-1, NIT-2 |
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

> **M1↔M2 boundary (redrawn — BLOCKER-3, David 2026-06-19):** M1's old gate ("997-accepted by TEMA") was
> unreachable inside the old M1 because the 997 ingester sat in M2 — a circular dependency. **Resolution: M1 = the
> FULL ASN→856/810→997/824 loop** (outbound generation + transmit-staging **plus** the minimal inbound ack
> ingestion that closes the loop and makes M1 gate-able on a real TEMA accept). **M2 = the remaining inbound (830/862
> forecast import + forecast-fill) + order-file generation + renban breakdown.** M1 grows (absorbs minimal inbound);
> M2 shrinks (loses 997/824) but still bundles three hard subsystems, so it is re-sized up, not down, from the old
> 5–7wk underestimate.

**M1 — The revenue-critical ASN → 856/810 → 997/824 daily loop (parallel-run shadow). ~7–10 dev-weeks.**
Ranks 1, 2, 4 + the **minimal-inbound (997/824) slice of Rank 3**.
- **Outbound:** ASN entry (create transaction + sequence check + split-by-ratio review), ASN invoice
  reconciliation, 856 + 810 builders (Gateway Jython, byte-exact, **built over NEW pure-SELECT NQs** — NOT the
  self-mutating `REPORT_EDI856`/`810`, BLOCKER-1/NIT-2), **decoupled per-ASN/per-INV status flip on send-commit**
  (Q2), in-place unsend/recreate-flag (Q5/D3), D6 window-aware pricing shared by 810/856.
- **Minimal inbound (the loop-closer):** the `poll_edi_in` poller shell — X12 parse honoring ISA separators,
  DUNS→site routing against ALL `sites` (D1), processed-files idempotency ledger, archive-after-commit — wired to
  dispatch **only 997 (AK9 ack, site-scoped `ein/ack`) and 824 (auto-flag rejected + home-screen alarm, Q10)**.
  830/862 dispatch is stubbed/deferred to M2. The poller is built here because M1's TEMA-accept gate requires it.
- **Mode:** **shadow** — Ignition writes ASN/INV rows + EDI files to a **separate `EDIOut` staging dir**; diff
  against legacy files; do NOT transmit Ignition's files to TEMA until byte-parity. **Inbound shadow:** ingest a
  copy of the 997/824 stream and write ack status to a **shadow column** first, then enable the real site-scoped
  `ein/ack` flip. Legacy stays the system of record. (Critical: never feed the builders from `REPORT_EDI856`/`810`
  — wrapping them would flip real status in the shared DB and make legacy skip its own send.)
- **Depends on:** ASN-detail key `(site_id, IN_ASN_ID, manifest)` (Q1); the embedded-UPDATE removal from
  `REPORT_EDI856`/`810` (BLOCKER-1); site-scoping of `UPDATE_EINStatus` (BLOCKER-2 — needed for the 997 flip);
  `CalculateASNFRS` body; ALC `AD_*` bodies + the per-site EIN sequence (Q4); AK9 semantics (Q6); the persistent-tx
  shim (proven). **Resolve the unread proc bodies before locking the estimate** (a hidden write already cost us
  BLOCKER-1).
- **Gate:** byte-diff 856/810 vs legacy on a sample day **AND** a real **TEMA test-accept** — an Ignition-built
  856/810 transmitted in a test exchange, the returned **997 ingested by M1's own poller** flipping the ASN/INV to
  `'A'`/`'R'` correctly (and an 824 reject auto-flagging its ASN). The loop is gate-able end-to-end within M1.

**M2 — Remaining inbound (830/862 forecast import + forecast-fill) + order-file generation + renban breakdown. ~8–11 dev-weeks.**
Ranks 5, 6 + the **830/862 + forecast-fill slice of Rank 3**.
- **Inbound (extends M1's poller):** add the **830 forecast import** (→ shared forecast-ingest service, per-site
  `auto`/`manual` mode + the home-hub Forecast Import box + ≥8-day staleness alarm, Q11) and **862** server-render;
  820 stays report-only. The poller, X12 parser, DUNS-routing, and processed-files ledger are reused from M1, not
  rebuilt.
- **Order/renban:** renban breakdown algorithm (atomic SERIALIZABLE counter claim, Carry 2), order-file generator
  (`.ord` byte-exact to supplier/logistics/archive — logistics emitted only when configured, Q8), order worksheet
  forecast-fill finished via `SIM_OrderSimulation` (no Excel), hot-call path (reuses M1's `create_asn` + `build_856`).
- **Mode:** inbound can **shadow** (830 parse + log, forecast written to a shadow set first); order files written
  to a staging dir and byte-diff'd before pointing at the real shares.
- **Depends on:** M1 (hot-call needs ASN+856; the poller shell + ledger are M1 assets); `AD_GetSpecialDate` body
  (shared calendar in VehicleOrder, read-only — calendar walk blocks the order forecast-fill, Q9); the TIRE-logistics
  config (resolved Q8); the unread `CalculateASNFRS`/FRS-renban insert bodies (resolve before locking this estimate).
- **Gate:** 830 round-trip lands the right per-site forecast; the forecast-fill matches legacy; `.ord` files
  byte-diff vs legacy and parse downstream.
- **Why re-sized up to 8–11wk:** even after shedding 997/824 to M1, M2 still bundles **three hard subsystems**
  (830/862 inbound + forecast-fill, renban breakdown with atomic counter, byte-exact `.ord` generation), several
  riding on proc bodies flagged "inferred, not fully read." The old 5–7wk estimate was optimistic; this is the
  honest size.

**M3 — Reporting + Excel-layer retirement. ~3–4 dev-weeks.**
Rank 7 + §3.
- Server-side render of the report families (8.3 Reporting module / 8.1 Perspective-export fallback); retire the
  last Excel/OLE paths; the Daily Shipping Assy Report (the FAILING path) renders + "prints" server-side.
- **Mode:** purely additive; reports are read-only — can run live alongside legacy immediately.
- **Gate:** report numbers match legacy for a sample; no Excel in any path.

**M4 — Security / multi-site / hardening. ~4–6 dev-weeks.**
§4.
- ✅ **PRE-TASK DONE (verify-before-relocate, 2026-06-19 — see `vehicleorder-sites-verification.md`).**
  Verified: **InventorySystem reads NO `sites` table** (source grep + the spike VehicleOrder, which holds only
  `LINE`); its site config comes from the INI (`SiteInfo.pas`). InventorySystem's real VehicleOrder deps are
  `LINE`, `AD_GetSpecialDate(s)` (calendar — stays shared, Q9), `AD_UpdateEIN` (EIN → per-site `INV_SITES`, Q4),
  + Reports. **So `INV_SITES` adoption is SAFE — net-new authoritative source, breaks no InventorySystem
  reader (relocate half already satisfied).** The production `VehicleOrder.sites` + its GALC/MES/Admin readers
  are **not inspectable here** (spike VehicleOrder is a `LINE`-only stub; no real dump in repo; siblings are
  separate codebases) → **do NOT retire the shared `VehicleOrder.sites` as part of InventorySystem work**;
  treat it as a shared external table that stays put. Physical retirement is a cross-system decision needing
  the real VehicleOrder schema / David's confirmation of readers.
- User Source + roles + per-feature gating; `site_id` end-to-end (schema surgery + **F1-safe `_HIST`: add `site_id`
  to the 3 history tables in lockstep with the base tables — the 7 enumerated `SELECT *` triggers at
  `:2664/3440/3617/4107/4269/5475/7496`, §4** + session scoping + parameterize the hardcoded EIN `6440`);
  **site-scope every EIN-keyed update (BLOCKER-2 — `UPDATE_EINStatus` → `(site_id, EIN)`)**; relocate the `sites`
  table VehicleOrder→Inventory DB (pending the pre-task) + build the Sites master screen (8th master, Admin-gated)
  + repoint the cross-DB `LINE`/DUNS reads; native alarming + notification pipelines; secrets/paths to gateway
  config + secret store; backup runbook; redundancy decision; transactional DATAPURGE scoped by site.
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

**Rough total: ~23–33 dev-weeks** for a solo dev, foundation reused (M1 7–10, M2 8–11, M3 3–4, M4 4–6, M5 1–2 —
re-sized after the BLOCKER-3 boundary redraw moved minimal inbound into M1 and the reviewer's honest re-estimate of
M2's bundled subsystems). M1 is the highest-value, highest-risk chunk and should be sequenced first; M3 (reporting)
is additive read-only and can slip in parallel with M1/M2 — it is the safe thing to defer if the timeline slips,
never a critical-path blocker.

### What parallel-runs/shadows vs what is a hard cutover
- **Parallel-run / shadow (most of it):** ASN/invoice *insert* writes (the `INSERT_ASN*`/`INSERT_INV*` procs
  dedup), inbound parse-to-shadow (997/824 ack to a shadow column first), all reporting (read-only), order-file
  staging, `site_id` additive schema. Both apps share the DB; procs mediate. **Caveat (BLOCKER-1):** shadow is
  only read-safe because the 856/810 builders feed from **pure-SELECT NQs**, not the self-mutating
  `REPORT_EDI856`/`810` — wrapping those would flip real status in the shared DB and break parallel-run.
- **Hard cutover (the few):** (1) **EDI transmission** — only one app transmits 856/810 to TEMA; (2) **the qty-
  trigger drop / seam-as-live-writer** (already designed/rehearsed in Phase C); (3) **operator auth** (they
  switch to Ignition logins). Everything else can shadow until these flip.

---

## 6. Open questions for David (decisions needed before / early in the build)

**Blocking M1:**
1. ✅ **RESOLVED (David 2026-06-19) — ASN-detail dedup/delete scope.**
   - **Preserve the accumulate-on-repeat upsert.** `INSERT_ASNDetail` (@HotCall=0) is an UPSERT: manifest
     exists → `IN_QTY += @Qty`, else INSERT. It is **called multiple times during ASN creation and the
     manifest qty accumulates** — this is intentional and must be kept. `@HotCall=1` always INSERTs (hot
     calls never dedup) — keep that branch too.
   - **Add `site_id` to the key** ("yes, add the site info"): the upsert existence-check AND the delete must
     include `site_id` so two sites' identical manifest numbers never collide/accumulate into one row.
   - **Also scope to `IN_ASN_ID` (recommended, pending David's nod).** Today the upsert WHERE and
     `DELETE_ASNItem` key on `VC_MANIFEST_NUMBER` **alone** (`DELETE_ASNItem` takes only `@ManifestNumber`)
     — so a later ASN reusing a manifest would accumulate into the OLD ASN's row, and deleting one ASN's
     line wipes that manifest from every ASN. Scoping both to **`(site_id, IN_ASN_ID, manifest)`** preserves
     the within-ASN accumulate while removing the cross-ASN/cross-site collision. **Final key (confirmed
     David 2026-06-19): `(site_id, IN_ASN_ID, manifest)`** — the upsert existence-check and `DELETE_ASNItem`
     both key on all three.
2. ✅ **RESOLVED (David 2026-06-19) — per-ASN status flip, not bulk.** Legacy `UPDATE_ASNStatus 'S'` does
   `UPDATE INV_ASN_MST SET VC_ASN_STATUS='S' WHERE VC_ASN_STATUS='C'` (flips EVERY open ASN at once,
   verified). **Rebuild: flip one ASN at a time, immediately after that individual ASN's 856 send.** Add
   `@ASNID` (+ `site_id`) → `WHERE IN_ASN_ID=@ASNID AND VC_ASN_STATUS='C' [AND site_id=@site]`. Couple the
   flip to the send (don't mark `'S'` until that ASN's 856 file write + transmit commits — the at-least-once
   idempotency rule, §4). This is also required for multi-user/multi-site (the bulk flip would send another
   operator's / another site's open ASNs). **Implementation note (BLOCKER-1):** this flip is now the *only* place
   ASN status moves to `'S'`. The 856 data feed is a NEW pure-SELECT NQ (`edi/asn_856_data`), **not** a wrap of
   `REPORT_EDI856` — that proc's `@EIN<>0` branch self-flips status to `'S'` on read (`/tmp/inv_utf8.sql:3695`),
   which would corrupt the shared DB during shadow/parallel-run. `REPORT_EDI810` has the identical twin (NIT-2
   confirmed). Both procs' embedded `UPDATE`s are removed and reimplemented as this decoupled per-row flip.
3. ✅ **RESOLVED (David 2026-06-19) — inclusive boundaries.** Window is inclusive (`VC_START <= prodDate <=
   VC_END`; `VC_END` = last effective day). Re-confirms the earlier D6 decision (David 2026-06-18, PR #10:
   strict `SELECT_ManifestCost` was the bug, superseded by `fn_ManifestCostAt`; gap convention for
   no-overlap). **810 prices identically to 856 — already guaranteed by construction:** all four billing
   procs were migrated to `CROSS APPLY fn_ManifestCostAt` (PR #14), so they share one inclusive window rule.
   Nothing further to build — already implemented + dress-rehearsed.
4. ✅ **RESOLVED (David 2026-06-19) — site lookups → Inventory.sites; EIN = the VAN tracking id.**
   - **`AD_GetSite` / `AD_GetSiteTMMDUNS` are eliminated as cross-DB ALC reads** — per the Q1 relocation,
     site + TMM DUNS are columns on the new `Inventory.dbo.sites` table, so the ASN/EDI code reads them
     locally. No ALC cross-DB site/DUNS lookup remains.
   - **`AD_UpdateEIN` → a per-site atomic EIN sequence in the Inventory DB.** Confirmed: **EIN is the
     authoritative outbound control number tracked through the VAN**, so it must be allocated atomically and
     uniquely per site (the per-site sequence replaces the ALC counter; seed lives on the `sites` row).
   - **BLOCKER-2 — site-scope the EIN-keyed status update.** Keeping per-site EIN sequences creates a collision in
     the ack path: `UPDATE_EINStatus` keys on `WHERE IN_ASN_EIN=@EIN` / `IN_INV_EIN=@EIN` **alone** (verified
     `/tmp/inv_utf8.sql:1724/1728`), so site A and site B both holding EIN 9069 means a 997 for site A flips
     **both** sites. **Add `site_id` to the WHERE** → the flip scopes to `(site_id, EIN)`. The inbound 997/824
     already resolves to a site by DUNS (Q7/Q11), so the `site_id` is in hand at flip time. **Audit every other
     EIN-keyed update for the same un-scoped pattern** before per-site EIN goes live.
   - **`CalculateASNFRS` — still a behavior-port item** (date/FRS logic, not site config): confirm the body
     with delphi-architect before porting; decide then whether it moves to Inventory or stays ALC. (The
     calendar proc `AD_GetSpecialDate` is tracked separately in Q9.)
5. ✅ **RESOLVED (David 2026-06-19) — in-place status + recreate flag, no hard-delete.** Legacy
   `UPDATE_INVUnsend` detaches the lines (`INV_ASN_DETAIL_MST.IN_INV_ID = null`) then **hard-deletes the
   `INV_INV_MST` header** — today the only way to flag an invoice for recreate, used when no ack/acceptance
   message ever comes back; the delete forces a full cost recreate on rebuild. (The proc even has a
   commented-out `SET VC_INV_STATUS='C'` — the status approach was contemplated.) **Rebuild (David's better
   solution):** update **in place** — set `VC_INV_STATUS` back to an unsent state + a clear **"recreate the
   810 file" flag**, keep the header row (recoverable, D3 — no destructive delete), and **recompute costs in
   place via `fn_ManifestCostAt`** (D6) rather than delete-and-rebuild. The flag drives the system to
   regenerate + retransmit the 810; the "no ack received" use case is preserved (operator action or an
   unacked-past-threshold alarm, §4). Audit trail of the original send is retained.
   - **Re-pricing instant pinned (SHOULD-FIX-5).** `fn_ManifestCostAt` is window-aware; the in-place recompute
     MUST key on the **original production date** (the ASN/line `VC_PRODUCTION_DATE` that priced the first send),
     **NOT `getdate()`/current.** Otherwise a recreate after a manifest-cost window boundary passes would silently
     re-price the 810 at the new window's price while TEMA may already hold the original (the "no ack" case is
     exactly when the 810 reached TEMA but the 997 was lost) → invoice mismatch/dispute. Add a parity test that
     recreates an 810 **across a cost-window boundary** and asserts the price is unchanged.

**Blocking M2:**
6. ✅ **RESOLVED (David 2026-06-19) — yes.** Map the AK9 functional-group ack codes (`A` accepted / `E`
   accepted-with-errors / `P` partial / `R` rejected) to **distinct invoice/ASN statuses**, and **tolerate
   the optional `AK2/AK3/AK4` transaction-set detail** between `AK1` and `AK9` (the legacy parse is fragile
   here). The 997 processor records the specific outcome per transaction set rather than a binary
   accepted/not.
7. ✅ **RESOLVED (David 2026-06-19) — yes, migrate the DUNS allow-list to `sites`.** The legacy
   `AD_GetSiteTMMDUNS` allow-list becomes per-site DUNS attributes on the relocated `Inventory.sites` table
   (consistent with Q4); inbound EDI resolves a file to a site by matching its DUNS against `sites`. **Build-
   time verification (not a decision):** confirm the exact `delSL[4]` X12 element index against a **real
   inbound ISA** before relying on it (the code index is authoritative; this is the per-site routing key) —
   pull a sample inbound file at M2 start (delphi-architect / a captured ISA).
8. ✅ **RESOLVED (David 2026-06-19) — skip-by-CONFIG (confirmed in code).** "No logistics" is a valid
   configured option — the supplier delivers on their own. Verified in `OrderFormCreateF.pas`: the logistics
   `.ord` file is written only `if lastlogisticsdirectory <> 'NONE'` (lines 111/135/147/301), and
   `lastlogisticsdirectory` is sourced from `SELECT_PartsStockLogistics.LogisticsDirectory` — **NULL → set to
   `'NONE'` → no logistics file** (lines 217-222). DUNLOP (`07451`) has no logistics directory → self-deliver
   → no file. The TIRE/WHEEL branch (line 244/255) only selects the Excel *template*, NOT logistics. **Rebuild
   generator: emit supplier + archive files always; emit the logistics file only when the site/supplier has a
   logistics destination configured** (a Logistics-master / `sites`-driven attribute, not a part-type rule).
9. ✅ **RESOLVED (David 2026-06-19) — calendar stays SHARED in VehicleOrder; read it, don't relocate.** Body
   provided + captured at `docs/analysis/production-calendar/AD_GetSpecialDate-shared.sql`. The production
   calendar (`AD_GetSpecialDate` + `SpecialDate`/`Line`/`ProductionStatus` + `F_ISO_WEEK_OF_YEAR`) lives in
   **`VehicleOrder` and is shared across ALL apps** (Inventory, GALC, MES) — so, unlike `sites`, it is **NOT
   relocated** into Inventory. The rebuild reads it **read-only via a Named Query / cross-DB proc wrapper**.
   **Architectural rule this sets:** app-specific config (`sites`) moves INTO Inventory; genuinely
   cross-app shared reference data (the plant calendar) STAYS shared in VehicleOrder. Specifics to honor: the
   **status domain is data-driven** (`ProductionStatus.ProductionStatus` + `ProductionStatusAbv` — do NOT
   hardcode O/X/holiday); **line resolution** = a specific `@LineName` returns its line-specific dates UNION
   the all-lines (`LineID IS NULL`) global dates, blank line returns all; week# via `F_ISO_WEEK_OF_YEAR`
   (ISO), day# via `DATEPART(DW, date + @@DATEFIRST - 1)`. **Reconcile (M2):** the Inventory `sites`/site-line
   config must reference the shared `VehicleOrder.Line.LineName` values (the calendar is keyed by LineName) —
   this scopes Q4's "repoint LINE": the calendar's `Line` table stays shared; only SITE config relocates.
10. ✅ **RESOLVED (David 2026-06-19) — auto-flag rejected + main-screen alarm + click-through to detail.** The
    824 inbound processor **auto-flags the named ASN as rejected** (a recoverable status flag per D3 — flag,
    don't delete) and **raises a native Ignition alarm surfaced on the main/home screen**; **clicking the
    alarm opens the 824 report detail** (the rejection reason per referenced transaction). Not operator-report-
    only. Reinforces §4 (native alarming) + the home hub gets an alarm indicator/banner; the alarm's
    associated data carries the ASN id + 824 detail so the click navigates straight to the report.
11. ✅ **RESOLVED (David 2026-06-19) — scheduled gateway poll + DUNS guard; forecast import is per-site
    auto/manual with a home-hub box.**
    - **Scheduled gateway poll** (not on-demand). Cadence configurable (TBD interval).
    - **DUNS guard + multi-site routing (SHOULD-FIX-1).** The **single** gateway poller (one gateway, Q14)
      serves **all** sites (D1), so there is no "the gateway's configured site." For each inbound file the poller
      reads the sender DUNS and **matches it against the DUNS of ALL configured `sites` rows, routing the file to
      the matching site** and processing it site-scoped. A file whose DUNS matches **no** configured site is
      **dropped/quarantined** (logged + alarmed), never consumed. (This is how Q7's per-site routing lands — DUNS
      is the routing key, not a single-site filter; it reconciles with the `delSL[4]→site` routing in §2 Rank 3.)
    - **Exception — forecast import has a per-site `auto`/`manual` config** (new `sites` attribute,
      `forecast_import_mode`): **AUTO** = the poll imports the forecast automatically; **MANUAL** = **leave the
      file in place** and raise a **home-screen alert on the Forecast Import button/box** so the operator runs
      it manually.
    - **ADDED SCOPE (David 2026-06-19) — a "Forecast Import" box on the home hub:** shows the **last import
      date/time**, the manual-run action (in manual mode) + a waiting-file alert, and **raises a staleness
      alert when ≥ 8 days have elapsed since the last forecast import** (native alarm + the box flags). Folds
      into §4 alarming + the home hub; `forecast_import_mode` + the last-import timestamp are managed/shown via
      the Sites master and the hub.

**Blocking M4 / M5:**
12. ✅ **RESOLVED (David 2026-06-19) — Admin/User split is enough for now.** Two roles: map legacy
    `BIT_ADMIN=1 → Admin`, else `User`. Gate Administration (Sites master, user mgmt) to Admin; operational
    views to User. The finer per-feature permission model (EDI-only / receiving-only operators) is **deferred**
    (future, not M4). Keeps the role design simple for a solo dev.
13. ✅ **RESOLVED (David 2026-06-19) — yes.** No plaintext migration; **force a password reset at first
    Ignition login.** Seed usernames/roles from `INV_USERS` once; users set a new password on first login
    (legacy plaintext is discarded, never hashed-in).
14. ✅ **RESOLVED (David 2026-06-19) — single gateway; redundancy is a transparent infra concern.** Build/plan
    for a **single gateway**. Redundancy will be discussed with the user/IT separately, but Ignition gateway
    redundancy is **invisible to the application** — so the app architecture does NOT design around it (no
    app-level failover logic); it's a deployment decision layered on later without code changes.
15. ✅ **RESOLVED (David 2026-06-19) — build FAITHFUL (Option A); defer C2–C6 to a post-cutover doc.** *"Build
    as legacy, don't confuse the users any more than we have to."* The rebuild reproduces the legacy order
    calc exactly so it parity-matches under the dev-mirror comparison (Q16). **C1 (silent ≤200-row truncation
    → remove cap + warn) IS included** — a defect fix that prevents silent data loss, no order-math change.
    **C2–C6 (MRP modernizations) are deferred** and tracked in
    `docs/analysis/production-readiness/post-cutover-enhancements.md`, to be adopted one-at-a-time post-cutover
    behind sign-off + a per-site toggle, validated against the faithful baseline.
16. ✅ **RESOLVED (David 2026-06-19) — dev mirror compared over a multi-week timeframe.** Rather than a prod
    parallel-run only, **run a MIRROR in dev** and compare Ignition output against legacy **over a multi-week
    window** before the transmission cutover. The comparison harness (the existing parity-diff approach +
    TEMA-accept checks) runs against the mirror; cutover gated on a clean multi-week comparison.
17. ✅ **RESOLVED (David 2026-06-19) — per-site retention policy.** DATAPURGE retention is configured/applied
    **per site** (a `sites` attribute, not a global policy); each site sets its own retention horizon. Keep the
    legacy table scope for now (extend later if needed); the purge runs site-scoped + transactional (Carry).

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
