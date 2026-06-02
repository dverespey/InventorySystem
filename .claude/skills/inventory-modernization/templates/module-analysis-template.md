# Module Analysis: <Module Name>

> Copy to `docs/analysis/<area>/<module>.md` and fill in. One file per module.
> Status: ⬜ not started · 🟨 in progress · ✅ spec complete · 🚀 rebuilt

**Area:** <e.g. Ordering>  **Status:** ⬜  **Analyst:** <name/date>

## 1. Legacy surface
- **Form(s):** `<Foo.pas>` + `<Foo.dfm>`  (size / complexity: ___)
- **Entry point:** how it's reached from `MainMenu.pas`.
- **Purpose (one paragraph):** what business job this screen does.

## 2. Data touched
| Table | Read | Write | Notes |
|-------|:----:|:-----:|-------|
| `INV_...` |  |  |  |

**Triggers on these tables:** (from database-objects.md) — capture each one's exact rule:
- `<trigger>` → <invariant it enforces>

## 3. Stored procedures used
(Grep the `.pas`/`DataModule.pas` for the proc names; read each with `sql.sh proc NAME`.)

| Proc | Operation | Business rule (summarized from the body) |
|------|-----------|------------------------------------------|
| `<PROC>` | SELECT/INSERT/... | ... |

## 4. Business rules & edge cases
- Validation, calculations, sequencing, status transitions.
- Date math / calendar dependencies (first production day, holidays).
- Anything implicit in trigger or proc logic that the UI relies on.

## 5. UI / UX notes
- Key fields, grids, filters, actions/buttons.
- Workflows / multi-step flows; printing/report hooks.
- What to keep vs. modernize (this is a chance to improve UX).

## 6. Target design
- **Models:** <ActiveRecord models + associations>
- **Controllers/routes:** <RESTful resources>
- **Views/components:** <screens>
- **Services:** <Python service? background job? proc-wrap vs reimplement?>
- **Reports:** <which REPORT_* procs, output format>

## 7. Migration plan for this module
- [ ] Stage 1 — wrap existing proc(s), read-only parity.
- [ ] Stage 2 — enable writes (mediated by shared procs/triggers).
- [ ] Stage 3 — reimplement proc/trigger logic in app (Postgres-ready).

## 8. Open questions for the user (domain expert)
- ...

## 9. Test cases / parity checks
- Inputs → expected outputs to validate new vs old against the same DB.
