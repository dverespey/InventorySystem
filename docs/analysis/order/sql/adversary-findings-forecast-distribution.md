# Adversary Findings — Forecast-Distribution Feed (P6, branch `p6-forecast-distribution`)

**Reviewer stance:** the reimplementation is WRONG until proven equivalent. Every finding carries a
counterexample (input → legacy vs rebuild) or a file:line. Destructive probes ran inside rolled-back
transactions on the rebuild `Inventory`; `Inventory_Live` / `VehicleOrder` were READ-ONLY; the spike was
left as-found (the e2e teardown ran and cleaned its sentinels — see NIT-4).

**Artifacts under review**
- Rebuild feed: `docs/analysis/order/project-library/forecast_distribution/code.py`
- Legacy oracle: `ForecastBreakdownF.pas:340-568` (emit at `:484-508`)
- Procs (live dump, UTF-16→UTF-8 `/tmp/createinv_utf8.sql`): `SELECT_ForecastSupplier` (`CreateInventory.sql:2058`),
  `SELECT_SupplierInfo` (`:5993`); table `INV_BREAKDOWN_FC_INF` (`:578`)
- Test: `scripts/e2e/test_forecast_distribution_e2e.py` (36 PASS / 0 FAIL, incl. `FD_REVERT_PROOF=1`)
- `.ord` consistency reference: `docs/analysis/order/project-library/order_file/code.py`

---

## VERDICT

**For a SINGLE-SITE deployment: the feed reproduces `SELECT_ForecastSupplier` + the `.frc` byte format
EXACTLY — no query, format, width, channel, filename, or sendsite divergence found.** The query is the same
proc on the same single table, the `>` filter is strict-future-only, the per-supplier break key is the same
case-sensitive column, the 12 columns are read from the same table, the `%.2d`/`%.5d` formatters match
Pascal `Format` byte-for-byte (verified across negative / overflow / NULL / zero), CRLF terminates every
line incl. the last, the `'/'`-only filename clean and the no-separator archive date are reproduced, and
NULL/blank/unknown Output File Type → BOTH (the no-ELSE CASE).

**One genuine, bounded gap (SHOULD-FIX-1): the sendsite prefix is NOT site-scoped.** The rebuild prepends
`SELECT TOP 1 VC_SUPPLIER_CODE FROM INV_SITES ORDER BY IN_SITE_ID` — the first site's code — to EVERY
sendsite supplier's line, regardless of which site the forecast belongs to. On the spike (`INV_SITES` = 2
rows: MAS, HERO) this unconditionally emits `MAS`. This is **faithful for single-site** (the legacy's
`fiSupplierCode` is one INI-global value; there is no per-site forecast data to disambiguate — proven: no
site column on `INV_BREAKDOWN_FC_INF`, no site FK on `INV_SUPPLIER_MST`), but it is **a latent
multi-site wire defect and is UNPROVABLE either way against the legacy** (the legacy cannot do multi-site,
so no oracle exists). It must not ship to a genuinely multi-site gateway without an M4 site-scoping fix.

No BLOCKER. The byte/number parity is PROVEN for single-site; the only divergence is the multi-site prefix,
which is correctly flagged GOLDEN-PENDING/M4 in the code but whose `TOP 1` resolution is an over-commitment
worth tightening before any multi-site cutover.

---

## Answers to the six attack vectors

### 1. Driving query == `SELECT_ForecastSupplier` — PROVEN EQUIVALENT
- The live proc body on the spike == the dump == the rebuild's `_FEED_SQL`
  (`EXEC dbo.SELECT_ForecastSupplier @WeekDate = ?`):
  `SELECT * FROM INV_BREAKDOWN_FC_INF WHERE VC_WEEK_DATE > @WeekDate ORDER BY vc_supplier_code,
  vc_part_number, vc_week_date`.
- **Strict `>` (future-only, today excluded) — PROVEN by counterexample.** Rolled-back tran, floor
  `20260601`: a `20260601` row (== floor) was EXCLUDED; a `20260602` row was INCLUDED. The rebuild passes the
  same `weekDate` floor straight to the proc, so the boundary is the proc's.
- **`VC_WEEK_DATE` varchar(8) lexicographic compare — SAFE.** `yyyymmdd` is lexicographically chronological;
  column collation `SQL_Latin1_General_CP1_CI_AS` (verified) doesn't affect digit ordering.
- **Supplier break — IDENTICAL.** Legacy break key (`ForecastBreakdownF.pas:366`):
  `fieldbyname('VC_SUPPLIER_CODE').AsString <> lastsupplier` (Delphi case-SENSITIVE string `<>`). Rebuild
  (`code.py:496`): `row["supCode"] != curSup` (Python case-sensitive). Same column, same case-sensitivity,
  same single-pass-over-sorted-stream break. No missed/extra break.
- **Same 12 columns, same single table — CONFIRMED.** `INV_BREAKDOWN_FC_INF` is a HEAP with 12 columns and
  NO index (verified `sys.indexes` → `HEAP`); the rebuild reads exactly the 12 named columns; no join, no
  derived column. The ORDER BY is the ONLY ordering source (relevant to the shared NIT-1 below).

### 2. `.frc` byte format == `ForecastBreakdownF.pas:484-508` — PROVEN EQUIVALENT
Field order/widths verified against the live `fText`/`fBoth` path, byte-for-byte:
- `[siteSupplierCode if sendsite]` + `VC_SUPPLIER_CODE`(raw) + `VC_PART_NUMBER`(raw) + `VC_WEEK_DATE`(raw)
  + `Format('%.2d',WeekNo)` + `Format('%.5d',IN_QTY1..7)`. The rebuild's `format_frc_line` assembles this in
  the same byte order (`code.py:103-112`).
- **`%.Nd` min-width semantics — PROVEN across all hazardous edges** (independent Python re-derivation of the
  Pascal `Format` vs the rebuild's `_format_qty`): `-5`→`-00005` (6 chars, sign OUTSIDE pad, NOT Python
  `'%05d'`'s `-0005`); `100000`→`100000` (widens, NO truncation); `0`/NULL→`00000` (`AsInteger` of NULL = 0,
  reproduced by `_as_int(None)=0`); week `Format('%.2d',[5])`→`05`, `[100]`→`100` (widens). All matched.
- **RAW (unpadded) supplier/part/weekdate — reproduced.** A positional-parse shift risk if the sub-supplier
  parser expects fixed widths, flagged GOLDEN-PENDING in both the spec and the code; faithful to the legacy.
- **CRLF after EVERY line incl. last — CONFIRMED** (`format_frc_file` joins with `\r\n` then appends a
  trailing `\r\n`; the legacy `Writeln` writes `\r\n` after every line). Test read back `'rb'` and asserted
  the trailing CRLF; shim `writeFile` is `wb`/no-translation (`jython_shim.py:488-497`).
- **Filename `<name '/'-stripped>-<code>.frc`, ONLY `/` stripped — CONFIRMED** (`_clean_frc_name`,
  `code.py:198-203`; legacy `:450` `ANSIReplaceStr(...,'/','')`). DELIBERATELY narrower than `.ord`'s
  `_clean_name` (which also strips `,`/`.`/`\`). Proven live in the trigger run: `EPC, INC-10011.frc` retains
  the comma (the `.ord` would strip it) and `PACIFIC_MFG-11111.frc` retains the underscore.
- **Archive `<name>-<code><yyyymmdd>.frc`, NO separator before the date — CONFIRMED**
  (`_frc_paths`, `code.py:332`; legacy `:456`). Test: `TextOnly Co-ZZP6T20260601.frc`.

### 3. Output-type CASE (TEXT/EXCEL/BOTH; NULL/blank/unknown → BOTH) — PROVEN EQUIVALENT
- The proc's `CASE VC_OUTPUT_FILE WHEN 'T' THEN 'TEXT' WHEN 'E' THEN 'EXCEL' WHEN 'B' THEN 'BOTH' END` has
  **no ELSE**, so any non-T/E/B value (incl. `'X'`, `''`, `' '`) and NULL → SQL `NULL`. Verified on the
  spike: the alias is ALWAYS one of the literal uppercase strings `TEXT`/`EXCEL`/`BOTH` or `NULL` — it can
  never be lowercase/mixed.
- Legacy (`:425-430`): `'TEXT'`→fText, `'EXCEL'`→fExcel, `else`(NULL or `'BOTH'`)→fBoth. Rebuild `_file_kind`
  (`code.py:206-216`): `.strip().upper()` then `TEXT`/`EXCEL`/`else BOTH`. Because the proc output is already
  uppercase-or-NULL, the `.upper()` is a harmless no-op and the four possible proc outputs map identically:
  TEXT→(text), EXCEL→(excel), BOTH→(both), **NULL→(both)** (H2). A NULL-output-type supplier gets BOTH, not
  skipped/errored — proven by the `SUP_NULL` e2e case (`['frc-main','xlsx']`). No divergence.

### 4. Sendsite prefix — FAITHFUL for single-site; UNPROVABLE/latent-defect for multi-site (SHOULD-FIX-1)
- The legacy commented-out original line `tcl := Data_Module.fiSupplierCode.AsString` (`:489`) reads
  `fiSupplierCode`, which is a **`TCIniField`** (`DataModule.pas:94`) — a single INI-file global, the one
  site's own supplier code (`Configuration.pas:113/164` read/write it; `EDIUpload.pas:96` consumes it). The
  rebuild maps this to `INV_SITES.VC_SUPPLIER_CODE` (`_site_supplier_code`, `code.py:287-298`), which the
  `.ord` `_M4` note resolves identically — so the CHOICE of column is consistent with the `.ord` rebuild's
  `BIT_SITE_NUMBER_IN_ORDER` leading-field stance. It does NOT reproduce the phantom-`SiteSupplierCode`
  crash and does NOT drop the prefix for all (which would break sendsite suppliers' positional parse). All
  of that is correct.
- **The defect:** `_site_supplier_code` resolves the prefix with `SELECT TOP 1 VC_SUPPLIER_CODE FROM
  INV_SITES ORDER BY IN_SITE_ID` — the FIRST site, unconditionally, for EVERY sendsite supplier. On the
  spike (`INV_SITES`: site 1 = `MAS`, site 2 = `HERO`) it always emits `MAS` (verified live + in the e2e:
  `siteSupplierCode='MAS'`, and the trigger run prefixed all 13 sendsite live suppliers with `MAS`).
- **Why this is faithful for single-site but a latent multi-site wire defect:** there is NO way to scope the
  prefix per-site from the forecast data — `INV_BREAKDOWN_FC_INF` has NO site column and `INV_SUPPLIER_MST`
  has NO site FK (only the `BIT_SITE_NUMBER_IN_ORDER` flag) (both verified via `sys.columns`). The legacy
  itself is single-site (one INI `fiSupplierCode`), so for a single-site deployment `TOP 1` == the one site's
  code == the legacy value — equivalent. But against a GENUINELY multi-site DB (the spike already has 2 sites)
  the rebuild emits the wrong site's supplier code for any supplier belonging to site 2+, and there is **no
  legacy oracle** to diff against (the legacy can't do multi-site). So this is simultaneously (a) faithful
  for the stated single-site target and (b) UNPROVABLE and wire-wrong for multi-site. The non-sendsite path
  (plain `VC_SUPPLIER_CODE`, no lead) is correct unconditionally.

### 5. Qty/edge cases — PROVEN; the divergence-prone edges are unreachable in production
- `>99999` widen, negative `-00005`, all-zero week, NULL qty → all match Pascal `Format` (vector 2).
- **Production reachability (Inventory_Live, 959 rows, READ-ONLY):** `IN_WEEK_NUMBER` ∈ [1,52] (never ≥100,
  so `%.2d` never widens in practice); max day-qty ≈ 3239 (never > 99999, H-OVF never triggers); zero
  negative qtys. So the H-OVF / H-NEG / 3-digit-week hazards are faithfully reproduced but **do not occur in
  current data** — parity holds and the edges are moot in practice.
- **Empty/whitespace supplier code, trailing-space part:** `ANSI_PADDING ON` on the table (verified) →
  varchar trailing spaces ARE retained (`'PT   '` round-trips with `DATALENGTH=5`). The rebuild's `_s()`
  passes the value through verbatim, matching Delphi `AsString` (which does not trim). This shifts the
  positional output IDENTICALLY in legacy and rebuild — no divergence. (See NIT-2: the e2e shim's `-W` flag
  TRIMS trailing spaces, so the TEST cannot catch a trailing-space shift; the rebuild logic is still
  faithful.)

### 6. Test integrity (R15) — PASSES
- The `.frc` expectations are built from an INDEPENDENT source re-derivation: `src_frc_line` / `_src_qty`
  (`test:217-232`) re-implement the Pascal `%.Nd` from the spec (`"%0*d" % (width, abs(n))` + sign outside),
  NOT a call into the rebuild's `_format_qty`. Separate code, same spec.
- **Non-vacuity PROVEN:** `FD_REVERT_PROOF=1` monkeypatches the rebuild's `_format_qty` to the wrong Python
  `'%05d'` and the `.frc`-byte check FAILS (`-0005 != -00005`) — confirmed in this run. The other format
  assertions (week `%.2d`, CRLF, NULL→`00000`, sendsite lead, Excel `AsString`) are each anchored to the
  source-derived constant, not the rebuild output. Not circular.
- The trigger test exercises the OPERATIONAL path (`import_830(emitFeed=True)` fires the feed POST-commit,
  `forecast/code.py:725`), matching the EDIUpload operational sender discriminator (spec §6) — the right
  source path.

---

## Findings

### SHOULD-FIX-1 — sendsite prefix is not site-scoped (`TOP 1 ... ORDER BY IN_SITE_ID`)
- **Claim:** for a sendsite supplier the rebuild leads every `.frc` line with the FIRST site's
  `VC_SUPPLIER_CODE`, not the supplier's actual site.
- **Counterexample:** spike `INV_SITES` = {1:MAS, 2:HERO}. `_site_supplier_code` returns `MAS`. The e2e
  trigger run prefixed all 13 sendsite live suppliers (GDYR, DICASTAL, DUNLOP, …) with `MAS` regardless of
  their site. No site column exists on `INV_BREAKDOWN_FC_INF` and no site FK on `INV_SUPPLIER_MST` to do
  otherwise.
- **Classification:** faithful for the single-site target (one INI `fiSupplierCode`), but an
  **unfixable-from-data / UNPROVABLE multi-site gap** and a latent wire defect if ever run multi-site. The
  code flags the field GOLDEN-PENDING/M4, but the `TOP 1` resolution silently commits to one site. Before any
  multi-site cutover this must be re-derived from the session's site (M4), not `TOP 1`, and proven against a
  golden `.frc` for the sendsite leading-field width/justification (still GOLDEN-PENDING — no captured
  downstream `.frc` exists).

### NIT-1 — case-variant supplier codes cause a shared file-break/overwrite bug (legacy == rebuild)
- Under the CI collation, `ZZX` and `zzx` sort as EQUAL on the supplier key, so the secondary key (part)
  decides order. Counterexample (rolled-back): rows `ZZX/AAAA`, `zzx/BBBB`, `ZZX/CCCC` sort to that exact
  order; the case-SENSITIVE break logic (in BOTH legacy `:366` and rebuild `:496`) then opens `ZZX`, breaks
  to `zzx`, breaks BACK to `ZZX` — re-opening/overwriting the first `ZZX` file and losing its line. This is a
  **shared latent bug, NOT a parity divergence** (identical in both), and supplier codes don't vary by case
  in practice. Flagged for completeness, not a parity finding.

### NIT-2 — the e2e shim's `-W` strips trailing spaces; the test cannot catch a trailing-space positional shift
- The table is `ANSI_PADDING ON` (varchar trailing spaces retained), and Delphi `AsString` + the gateway
  JDBC both preserve them, but the e2e `sql()` and the shim row-parse run sqlcmd with `-W`
  (`jython_shim.py:54`), which TRIMS trailing whitespace. So a `VC_PART_NUMBER='PT   '` reaches the rebuild as
  `'PT'` under the test. The rebuild's `_s()` is faithful (it passes through whatever it gets); this is a
  **test-fidelity gap**, not a rebuild defect — the test would GREEN a trailing-space shift that the gateway
  would actually emit. A golden `.frc` (or a gateway-transport assertion) is the only way to close it.

### NIT-3 — `system.file.writeFile` newline/charset is gateway-dependent (carried from the 856/810 seam)
- The shim encodes ASCII + `wb` (no newline translation), so `\r\n` is verbatim; the rebuild embeds explicit
  `\r\n`. On the real gateway, `system.file.writeFile(path, str)` writes the platform charset and must not
  re-translate newlines. This is the same managed assumption the already-shipped 856/810 lanes rely on — not
  a new divergence, but it is a gateway-only assertion the spike cannot prove.

### NIT-4 — stale sentinel rows observed pre-run (cleaned by this run's teardown)
- At session start the rebuild `Inventory` held leftover `ZZP6*` synthetic supplier/breakdown rows from a
  PRIOR aborted e2e run (a real feed run would have emitted spurious `ZZP6*` `.frc` files). The clean e2e run
  here executed its teardown and the post-run assertion confirmed `leftover=0` and `INV_SITES` unchanged.
  Spike left as-found (restored). Flagged so operators know an aborted P6 e2e can leave sentinels that a
  later real emit would pick up.

---

## Equivalence statement
- **`SELECT_ForecastSupplier` reproduction: PROVEN equivalent** (same proc, same table, strict-`>` boundary
  proven, identical break key/order, 12 named columns).
- **`.frc` byte format: PROVEN equivalent** for all reachable and hazardous inputs (widths, negatives,
  overflow, NULL, CRLF, `'/'`-only clean, no-separator archive date, channel selection incl. NULL→BOTH).
- **Sendsite prefix: PROVEN faithful for single-site; UNPROVABLE + latent wire defect for multi-site**
  (SHOULD-FIX-1).
- No BLOCKER. The number/byte the sub-supplier parses is correct under the single-site target; the only way
  to ship a wrong byte is the multi-site sendsite prefix or an untested trailing-space shift — both flagged.
