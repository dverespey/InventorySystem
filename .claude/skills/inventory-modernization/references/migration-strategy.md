# Migration Strategy

How we rebuild the Delphi InventorySystem as a modern web app. Living document —
update the phase checklist as modules are completed.

## Constraints & decisions (from the user)
- **Full-system** rebuild, web-based.
- **Parallel run:** the Delphi app keeps operating during the transition. The new app
  and old app point at the **same SQL Server database** initially.
- **Database stays SQL Server** for now. **Postgres is a later phase.**
- **App first, DB migration later.**
- **Solo developer**, Ruby + Python skillset. Keep it simple and one-person-maintainable.

## Recommended stack (working assumption — confirm with user)
- **Primary web app: Ruby on Rails.** Batteries-included, fast CRUD scaffolding for the
  many master/transaction screens, mature auth, good for a solo dev.
  - DB access: ActiveRecord against **SQL Server** via the `activerecord-sqlserver-adapter`
    + `tiny_tds` gems. Run `rails db:schema:dump` against the existing DB rather than
    writing migrations at first (DB is owned by the legacy app during parallel run).
- **Python services for the data-heavy edges:**
  - **EDI X12** parse/generate (810/856/830) + FTP transport.
  - **Forecasting / FRS breakdown** math.
  - Expose as a small FastAPI service Rails calls, or run as scheduled jobs.
- **Reporting:** the 29 `REPORT_*` procs → call directly at first, render HTML/PDF
  (replaces QuickReport). Migrate proc logic into the app later.

> Open question for the user: confirm Rails-as-primary vs Python-as-primary. The plan
> below is stack-shaped but the *sequence* and *findings* are stack-neutral.

## The central pattern: procs are the spec
Because business logic lives in 179 procs + 24 triggers (not the Delphi UI), migration
of a module = **understand its procs/triggers, then re-express them in app code.**

Two-stage approach per proc:
1. **Wrap (parallel-run phase):** call the existing stored proc from Rails/Python so the
   new UI behaves identically to the old app. Lowest risk; both apps share one DB.
2. **Reimplement (post-cutover / Postgres phase):** translate the proc/trigger logic into
   ActiveRecord + service objects / model callbacks, so the DB becomes "dumb" and
   portable to Postgres.

Triggers (esp. the stock-quantity ones) become **model callbacks or service-object
transactions** — capture each trigger's exact rule in the module spec before touching it.

## Parallel-run / data-integrity rules
- New app is **read-mostly first**; enable writes module-by-module after parity is proven.
- Never let new and old app *both* own the same write path simultaneously without the
  shared procs/triggers mediating — that's why stage-1 "wrap the proc" matters.
- Keep `INV_PROGRAM_VERSION` semantics in mind (legacy app gates on version).

## Phase checklist

### Phase 0 — Foundations
- [ ] Confirm stack (Rails primary?) and hosting (on-prem vs cloud).
- [ ] Rails app skeleton + connect to SQL Server (read-only) via `tiny_tds`.
- [ ] `db:schema:dump` the 41 tables; generate baseline models.
- [ ] Auth: port `INV_USERS` / Logon (note legacy password scheme — audit it).
- [ ] Sanitized config (no plaintext DB passwords; use Rails credentials/ENV).

### Phase 1 — Master data (lowest risk, proves the pattern)
- [ ] Supplier, Size, Logistics, ManifestCost, Renban group, Part type — CRUD modules.

### Phase 2 — Calendar & inventory
- [ ] Production calendar (first production day, overtime/holiday).
- [ ] Parts stock / inventory views (read), Stocktaking.

### Phase 3 — Core transactions
- [ ] Receiving (RecConfStat, rejects) — careful with stock triggers.
- [ ] Shipping.
- [ ] **Ordering & Renban** (highest value/risk — `Order.pas` + ordering procs).
- [ ] **Forecasting / FRS** (Python service for the math).

### Phase 4 — EDI & billing
- [ ] EDI 856 ASN generation (Python).
- [ ] EDI 810 invoice generation (Python).
- [ ] EDI 830 forecast ingestion + FTP transport.

### Phase 5 — Reporting
- [ ] Port the 29 `REPORT_*` procs to web reports (HTML/PDF/export).

### Phase 6 — DB modernization (Postgres)
- [ ] Reimplement remaining procs/triggers in app layer.
- [ ] Schema translate SQL Server → Postgres; data migration + cutover.
- [ ] Decommission the Delphi app.

## Where analysis artifacts go
Per-module deep-dives: `docs/analysis/<area>/<module>.md` using
[../templates/module-analysis-template.md](../templates/module-analysis-template.md).
Update [database-objects.md](database-objects.md) and [module-map.md](module-map.md)
as understanding deepens.
