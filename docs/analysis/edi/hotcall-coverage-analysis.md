# Hot-Call Entry — Coverage Analysis (punch-list P12)

**Author:** delphi-architect · **Date:** 2026-06-21
**Live unit:** `HotCallEntry.pas` / `.dfm` — confirmed live at `InventorySystem.dpr:55`
(`HotCallEntry in 'HotCallEntry.pas' {HotCallEntryForm}`). Form caption "One Cycle Entry".
**Frequency:** ~112 runs/yr (~weekly), per the punch-list. Urgent, out-of-cycle shipment the operator
keys by hand.
**Evidence base:** `HotCallEntry.pas` read in full; M1 artifacts
(`docs/analysis/edi/project-library/asn/code.py`, `spike-asndetail-rekey.sql`,
`production-readiness/m1-asn-creation-spec.md`); the 856/810 builders
(`856/project-library/edi856/code.py`, `810/project-library/edi810/code.py`,
`856/spike-edi856-feed.sql`); `INSERT_ASNInfo` body dumped **live** from the spike DB; hot-call ASN
rows queried **live** (315 headers, feed run against a real hot-call ASN).

---

## VERDICT (one line)

**PARTIAL GAP.** The **write seam is COVERED** — hot-calls call the *same* `INSERT_ASNInfo` /
`INSERT_ASNDetail` / `AD_UpdateEIN` procs the M1 ASN keystone uses, with the M1 Q1 re-key and EIN
model serving them unchanged. The **entry path is a GAP** — there is no Perspective hot-call screen
and no `create_hotcall_asn` driver yet. The gap is **small**: a 12-row manual-entry view + a thin
driver that *reuses the M1 ASN write seam* (it does NOT touch the forecast/ratio fan-out at all).
Hot-call ASNs **flow through the SAME 856/810 we built**, verified against live data (the M390 path
is already implemented).

---

## 1. The full hot-call entry flow (cite `HotCallEntry.pas`)

### UI (`.dfm` + `Execute`)
- **Header (`GroupBox2`)**: `Line_ComboBox` (line), `ASN_DateTimePicker` (ASN date),
  `ManifestNumber_Edit` (8-char manifest, `MaxLength=8`, `.dfm:373`).
- **Items (`ASNItems_GroupBox`)**: **12 fixed rows**, each a part `ComboBox` (`AssyPartsCodeN_ComboBox`,
  upper-cased) + a qty `Edit` (`QtyN_Edit`, `MaxLength=3`). So **max 12 part/qty pairs**.
- `Execute` (`:92-125`) loads the line list from **`AD_GetLines;1`** on the ALC dataset (`:101`), then
  `ClearEntries` + `ShowModal`. `ClearEntries` (`:72-90`) populates each part combo from
  `INV_FORECAST_DETAIL_INF.VC_ASSY_PART_NUMBER_CODE` via `SelectSingleField` (`:83`) — so the part
  picklist is the configured assy-part master, but the operator chooses the **final part numbers + qtys
  directly**.

### Manifest number — OPERATOR-TYPED, not generated, NOT `'7'`-prefixed
The manifest is **typed by the operator** into `ManifestNumber_Edit` and passed verbatim to
`INSERT_ASNDetail` `@Manifest` (`:273-274`). The only validation is **length ≥ 8** (`:157-162`,
`ShowMessage('Manifest number must 8 characters')`); the numeric-check is **commented out** (`:150-155`).
This is the **structural difference from the forecast path**: the M1 `create_asn` *generates* the
manifest as `'7' + copy(prodDate,4,5) + assyManifestId` (`asn/code.py:_manifest`, → `M391`). Hot-call
manifests do **not** start with `'7'`, so they become **M390** on the wire (§4). **Live-confirmed**:
hot-call detail manifests are operator values like `52089698`, `52089578`, `52089154` — none start `'7'`.

### The writes, in order (one Inv_Connection transaction, `:221-295`)

| # | Call | Line | Key params |
|---|---|---|---|
| 0 | `Inv_Connection.BeginTrans` | `:221` | the only transaction (Inv side) |
| 0a | `SiteDataset.Close/Open` | `:223-224` | reads `SiteEIN` (the ALC Site row) |
| 1 | `INSERT_ASNInfo;1` (OUTPUT `@ASNID`) | `:229-255` | see below |
| 2 | per filled row: `INSERT_ASNDetail;1` | `:265-281` | `@HotCall=1` (always-insert) |
| 3 | `AD_UpdateEIN` | `:290-292` | `UPDATE Site SET SiteEIN=SiteEIN+1` (ALC) |
| 4 | `Inv_Connection.CommitTrans` | `:295` | commit header + details |

**`INSERT_ASNInfo` params (`:231-253`)** — the **hot-call sentinel values**:
- `@StartSeq = -1` (`:238-239`) and `@EndSeq = -1` (`:242-243`) — the **`-1` sentinel**. This is what
  marks the header as a hot-call: `SELECT_ASNSeq`'s `VC_START_SEQ_NUMBER <> -1` filter excludes these
  from the morning-create idempotency lock, and the legacy 856 filename branch tests `StartSeq='-1'`.
- `@LineName = Line_ComboBox.Text` (`:234-235`); `@AssyLine = ''` (`:236-237`).
- `@DTStartSeq = now`, `@DTEndSeq = now` (`:240-245`) — the build-window datetimes are just *now* (no
  real broadcast window — there is none).
- `@QTY = qty` (`:246-247`) — **note the bug**: `qty` here is whatever the last validation-loop
  iteration left in the var (`:184/:264`), NOT a deliberate header total. The header `IN_QTY` is
  effectively garbage/last-row. Harmless in practice (the 856/810 read **detail** `IN_QTY`, not header
  qty), but reproduce-or-fix is a decision (see §4 hazards).
- `@PDate = formatdatetime('yyyymmdd', ASN_DateTimePicker.DateTime)` (`:248-249`).
- `@EIN = SiteDataset.SiteEIN + 1` (`:250-251`) — **EIN allocated at create**, SiteEIN+1, same as the
  legacy morning ASN.
- **Status is hardcoded `'C'`** inside `INSERT_ASNInfo` (live body: `VALUES( @Ein, 'C', ...)`).
  `@ASNID` returned via `SCOPE_IDENTITY()`. **`RecordID := @ASNID`** (`:255`) → stamped on every detail.

**`INSERT_ASNDetail` per part (`:258-285`)** — loop walks `ASNItems_GroupBox.Controls`; for every
non-empty qty `Edit` it pairs `Controls[i-1]` (the part combo) with `Controls[i]` (the qty):
- `@ASNID = RecordID` (`:269-270`); `@EIN = SiteEIN+1` (`:271-272`); `@Manifest = ManifestNumber_Edit`
  (`:273-274`); `@PartNumber = Controls[i-1].Text` (`:275-276`); `@Qty = qty` (`:277-278`);
- **`@Hotcall = 1`** (`:279-280`) → in the proc the `@HotCall=1` branch is **always-INSERT, never the
  manifest accumulate** (`spike-asndetail-rekey.sql:85-90`). Every entered line is a discrete row even if
  two share the manifest.
- Logs `LogActLog('HOTCALL', ...)` per row (`:282`).

**`AD_UpdateEIN` (`:290-292`)** — runs on `UpdateReportCommand` (ALC connection) → bumps the **same
per-site `SiteEIN` counter** the normal ASN/856 path bumps (`m1-asn-creation-spec.md` step 3:
`UPDATE Site SET SiteEIN = SiteEIN+1`, no WHERE). So the hot-call **shares the single per-site EIN
sequence** with the morning ASN — it is not a separate counter.

---

## 2. How it DIFFERS from the forecast-driven `create_asn` (M1)

The difference is entirely the **INPUT side**. The write seam is identical; the explode is replaced by
manual entry.

| Aspect | M1 `create_asn` (forecast/broadcast) | Hot-call (`HotCallEntry`) |
|---|---|---|
| Qty source | `AD_FRSPULL` (GALC vehicle counts) × forecast tire/wheel ratios, banker's-rounded | **Operator types final qty** per part. **No AD_FRSPULL, no ratio explode, no rounding.** |
| Parts | fanned out from `SELECT_ForecastDetailBCASN` per broadcast code | **Operator picks ≤12 parts** from the assy-part picklist |
| Manifest | **generated** `'7'+copy(pd,4,5)+assyId` → **M391** | **operator-typed 8-char** → **M390** |
| Seq | real broadcast S/E seq numbers on the header | **`-1` / `-1` sentinel** |
| Build window | real `DT_START_SEQ`/`DT_END_SEQ` from the build | `now` / `now` (no real window) |
| Detail upsert | `INSERT_ASNDetail @HotCall=0` (accumulate per (ASN,manifest)) | `INSERT_ASNDetail @HotCall=1` (**always insert**) |
| Idempotency | `SELECT_ASNSeq` lock + unique index `WHERE VC_START_SEQ <> '-1'` | **none** — `-1` rows are deliberately allowed to repeat per (line,prodDate); two clicks = two ASNs (legacy parity) |
| Abort conditions | missing manifest cost / BC with no forecast → raise | none in Pascal (no cost pre-check) |
| Header / details / EIN procs | `INSERT_ASNInfo`, `INSERT_ASNDetail`, EIN | **SAME procs** |

So a hot-call is `create_asn` **with the entire AD_FRSPULL→ratio→manifest-gen stage deleted** and a
manual `[(part, qty)] + manifest` substituted, plus `@HotCall=1` and the `-1` seq. Nothing
hot-call-specific lives in the procs except the `@HotCall=1` always-insert branch and the `-1` sentinel
semantics (which the M1 idempotency index already accounts for —
`spike-asn-unique-guard.sql:22-27,49`).

---

## 3. COVERAGE verdict (split)

### (a) WRITE SEAM — **COVERED**
The hot-call calls **exactly the procs the M1 build already drives**, with compatible params:
- `INSERT_ASNInfo` — same OUTPUT-`@ASNID` header insert, status `'C'`. M1's `create_asn` already issues
  this via `runScalarPrepQuery("DECLARE @id int; EXEC INSERT_ASNInfo @ASNID=@id OUTPUT, ...; SELECT @id")`
  (`asn/code.py:395-402`). The hot-call just passes `@StartSeq=-1, @EndSeq=-1, @AssyLine=''` — all
  ordinary param values the same EXEC accepts.
- `INSERT_ASNDetail` — **same proc, same re-key.** The Q1 re-key
  (`spike-asndetail-rekey.sql`) **explicitly preserves the `@HotCall=1` always-insert branch**
  (`:85-90`, comment "@HotCall=1: hot calls never dedup -> always INSERT (unchanged)"). M1 calls it with
  `@HotCall` defaulted to 0; the hot-call driver just passes `@HotCall=1`. **No proc change needed.**
- `AD_UpdateEIN` — same ALC per-site counter bump. (Note: M1's chosen model is **EIN-at-send**, not
  EIN-at-create — see hazards; the seam still uses the same counter.)
- The Q1 per-(ASN,manifest) re-key and the unique-index race guard both already handle hot-call rows
  correctly (the index filters out `VC_START_SEQ_NUMBER='-1'` so hot-calls never trip it —
  `spike-asn-unique-guard.sql:49`).

**Conclusion:** M1's write seam serves the hot-call as-is. A `create_hotcall_asn` driver reuses the
*same two EXEC statements* M1 wrote; it does not need new SQL.

### (b) ENTRY PATH — **GAP**
There is **no** Perspective hot-call view and **no** `create_hotcall_asn` driver. M1 built only the
**forecast** producer (`create_asn` + `computeAsnDetails`, the AD_FRSPULL fan-out). The manual-entry
screen and the thin driver that turns 12 manual rows into the header + N `INSERT_ASNDetail(@HotCall=1)`
calls do not exist yet. This is the buildable gap.

---

## 4. The GAP + build scope

### What needs building
1. **Perspective view `HotCallEntry`** (small): line dropdown (from `AD_GetLines`), ASN-date picker,
   8-char manifest input (validate length ≥ 8), and a **repeatable part/qty line grid** (replace the
   legacy fixed-12 positional `Controls[i±1]` coupling — flagged fragile in `asn-invoice.md:211`).
   Part picklist from `INV_FORECAST_DETAIL_INF.VC_ASSY_PART_NUMBER_CODE`. Client-side validation: qty
   numeric & > 0 (`HotCallEntry.pas:184-197`), part required when qty present (`:204-218`), manifest
   length (`:157`).
2. **Driver `create_hotcall_asn(line, prodDate, manifest, items[], site=...)`** (small — reuses the M1
   seam): ONE Inventory transaction →
   - `INSERT_ASNInfo @StartSeq='-1', @EndSeq='-1', @AssyLine='', @PDate=prodDate, @Qty=<total or 0>,
     @Ein=<0 per M1 at-send model>` → capture `@ASNID` (the M1 DECLARE/SELECT pattern verbatim).
   - per item: `INSERT_ASNDetail @ASNID, @EIN=0, @Manifest=manifest, @PartNumber, @Qty, @HotCall=1`.
   - commit. **No AD_FRSPULL, no SELECT_ForecastDetailBCASN, no computeAsnDetails, no rounding.** It is
     a ~30-line driver — the simplest producer in the suite.

### Size
**Small.** The view is one form; the driver is a thin wrapper over the already-built write seam (no new
procs, no fan-out logic, no ratio math). Materially smaller than `create_asn`.

### Hazards to decide (record in a hot-call divergence ledger if any diverge)
- **EIN-at-create vs EIN-at-send (DIVERGENCE — flag for David).** Legacy stamps `SiteEIN+1` at create
  and bumps the ALC counter in the create tx (`HotCallEntry.pas:250-251,290-292`). M1's `create_asn`
  deliberately writes `IN_ASN_EIN=0` at create and allocates the EIN atomically from
  `INV_SITES.IN_EIN_SEQ` **at 856 send** (`asn/code.py:296-303`). **The hot-call driver must follow the
  M1 at-send model for consistency** (otherwise hot-calls and morning ASNs allocate EINs from different
  mechanisms). This changes nothing TEMA sees on the wire (the EIN is still on the 856), but it is an
  intended divergence from legacy already ratified for the ASN — extend the same decision to hot-calls.
- **856 filename `8HC…` branch (RESOLVED — 2-part fix, operational-sender-faithful).** Legacy emits `8HC…`
  instead of `856…` for hot-call files. The original 8HC work was anchored on the RECREATE button
  (`ASNInvoice.pas:817-825`) — which the sql-adversary + delphi-architect found is the WRONG anchor: the
  live daily path is the OPERATIONAL SENDER `MainMenu.ResendMarkedEDIsClick` (`MainMenu.pas:2691-2771`, the
  C→S-flip send path). The operational sender's filenames differ from the recreate button: BOTH branches
  append `LineName`, and the offsets differ. The fix had **TWO parts** (both branches were wrong):
  - **HOT-CALL** (`MainMenu.pas:2722-2724`): `8HC + copy(PickupDate,4,5)=Y+MMDD + IntToStr(y) + LineName +
    .txt`. e.g. `20260618`/y=1/COROLLA → `8HC606181COROLLA.txt`. (The recreate-anchored build had no
    LineName and a literal `1` instead of the `y` counter.)
  - **NORMAL** (already-merged PR #29; `MainMenu.pas:2718`): `856 + copy(PickupDate,5,4)=MMDD + LineName +
    .txt`. e.g. `20260618`/COROLLA → `8560618COROLLA.txt`. The merged normal filename was ALSO wrong — it
    omitted `LineName` AND used the wrong date offset (`[3:8]` instead of `[4:8]`). Fixed here too.
  - `y` is a per-send-batch counter (`:2702`/`:2724`); the rebuild reproduces it per-ASN as
    `1 + count of same-day same-line hot-call ASNs already flipped to 'S'` (deterministic, collision-free).
    Single hot-call/day → 1. **Exact `y` range golden-pending (P13 cutover check).** See `_filename_856` /
    `send_856` (`edi856/code.py`) + decision E in `edi856-wire-format.md`.
- **Header `@QTY` garbage (`HotCallEntry.pas:246-247`).** Legacy sets header `IN_QTY` to a stale loop
  var. Reproduce-or-fix decision; **no wire impact** (856/810 read detail qty). Recommend writing the
  sum of entered qtys (safer, no customer-visible change) and decide-and-flag in the ledger.
- **No idempotency / double-click = 2 ASNs.** Legacy has no guard for hot-calls (the `-1` rows
  legitimately repeat). The rebuild should NOT add a unique guard here (`spike-asn-unique-guard.sql`
  already excludes `-1`); a debounce/confirm on the view is the right place — faithful behavior.
- **Multi-site.** Same M4 re-key as the ASN: `INSERT_ASNDetail` gains `@Site`, the EIN comes from
  `INV_SITES.IN_EIN_SEQ` per site, `INSERT_ASNInfo` gets `IN_SITE_ID`. The driver passes `site`; no
  hot-call-specific multi-site work beyond what M1 already scoped.

### Milestone
**M1-follow-on** (a sibling of `create_asn`, before or alongside the M1 856/810 send work). It is part
of the daily revenue loop (a hot-call ships and bills like any ASN) and is tiny because the seam exists.
Not M4. Treat it as "the second ASN producer" — the same rank as `create_asn`, reusing its plumbing.

---

## 5. Do hot-call ASNs flow through the SAME 856/810? — **YES (live-verified)**

- **The 856 feed is keyed by `IN_ASN_ID` ALONE** (`spike-edi856-feed.sql:64-73`) — no status filter, no
  `VC_START_SEQ` filter, no `'7'`-manifest filter. Once a hot-call ASN has header + detail rows it reads
  identically. **PROVEN**: running the exact M1 feed SQL against live hot-call ASN **id 4712** returns
  the hot-call detail row `Manifest=52089698, Part=42600FEL2000, ShipQty=1, PickUpDate=20260606,
  Kanban=JZV5`. The hot-call line is selected and built with no special-casing.
- **The 810/856 M390 handling already exists.** Both builders compute `IT101 = "M391" if manifest
  starts '7' else "M390"` (`edi810/code.py:287`; `edi810-wire-format.md:24-25`;
  `asn-invoice.md:186-187`). A hot-call's operator-typed non-`'7'` manifest → **M390** automatically.
  **No new wire work** for hot-calls in the 856/810.
- **Live data scale:** 315 hot-call headers (`VC_START_SEQ_NUMBER='-1'`), all now status `'A'`
  (created `'C'` → flipped at send → archived); detail qtys are small (1-3) — consistent with urgent
  single-unit out-of-cycle shipments. These have been flowing through the legacy 856/810 as M390 for
  years; the M1 builders reproduce that path.

**Caveat to verify against the golden (data-dependent):** the only behavioral divergence on the
out-bound side is the **filename** (`8HC…` vs `856…`, hazard above). The *content* (segments, M390,
EIN, qty) of a hot-call 856/810 is covered by the M1 builders; confirm a sample hot-call 856 file's
**filename and BSN/ST contents** against a golden legacy `8HC…` file before cutover.

---

## Bottom line for the punch-list

- **Verdict:** **PARTIAL GAP** — seam COVERED, entry path is the GAP.
- **Covered:** `INSERT_ASNInfo` / `INSERT_ASNDetail` (incl. the Q1 re-key + the preserved `@HotCall=1`
  always-insert branch) / the EIN counter; and the 856 **and** 810 outbound (M390 path proven on live
  hot-call data, feed keyed by `IN_ASN_ID`).
- **Needs building:** a Perspective hot-call entry view (line + date + 8-char manifest + ≤12 part/qty
  grid) and a thin `create_hotcall_asn` driver reusing the M1 ASN write seam.
- **Scope/size:** **small** — the driver is the smallest producer in the suite (no fan-out, no ratio,
  no rounding); the view is one form.
- **Flag-for-David divergences:** (1) the `8HC…` 856 filename branch M1 dropped (customer-visible →
  decision); (2) extend the EIN-at-send model to hot-calls (already-ratified ASN decision); (3) header
  `@QTY` garbage (no-wire-impact, decide-and-flag safer fix).
- **Milestone:** **M1-follow-on** (second ASN producer), not M4.
