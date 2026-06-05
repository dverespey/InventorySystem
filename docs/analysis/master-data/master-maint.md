# Module Analysis: Master-Maintenance Menu (Hub)

**Area:** Master data  **Status:** ✅ spec complete  **Analyst:** Claude / 2026-06-04

> **Not a CRUD module.** `MasterMaint` is the *navigation hub* for the master-data area —
> a one-screen button menu that opens the individual master editors (Supplier, Logistics,
> Size, Parts/Stock, Renban Group, Assembly Detail, Assembly-Ratio, Monthly-PO /
> Manifest-Cost, ASN/Invoice). It owns **no table, calls no stored procedure, and writes
> nothing**. Its only "logic" is two `[SITE]` INI feature flags that toggle which buttons
> appear and what one button is labeled/does. Read this spec as the **router/IA contract**
> for the master-data section of the rebuilt app, not as a data spec.

## 1. Legacy surface
- **Form:** `MasterMaint.pas` (5.2 KB / 198 lines) + `MasterMaint.dfm` (2.8 KB / 123 lines).
  `TMasterMaint_Form`, Caption **"Master Data Maintenance Menu"**, header label
  "Master Data Maintenance Menu". `BorderStyle = bsDialog`, `Position = poDesktopCenter`.
  Author: Aaron Huge, 2002-10-25; later edits: David Verespey added Monthly PO (2005-04-12)
  and "BC ratio" (2006-09-27). **Tiny** even by this repo's standards — there is no grid,
  no edit panel, no dataset, no ADO call.
- **Registered live** in `InventorySystem.dpr` line 13:
  `MasterMaint in 'MasterMaint.pas' {MasterMaint_Form}`.
- **Entry point:** `MainMenu.pas` → **`DateMaint_ButtonClick`** (lines 360-367):
  `Hide; MasterMaint_Form := TMasterMaint_Form.Create(self); MasterMaint_Form.Execute;
  MasterMaint_Form.Free; Show;`. (The MainMenu handler is named `DateMaint_*` and the
  form's header label var is `MastDateMaintMenu_Label` / its error text says "Master Date
  Maintenance screen" — all stale "Date" typos for **Data**. Cosmetic, but note it: the
  menu is *Master **Data*** maintenance, not date maintenance — the production-calendar
  date editors live elsewhere, e.g. `OvertimeHoliday`.)
- **Purpose (one paragraph):** Present a button grid; each button **modally** opens one
  master-data editor child form and returns here when it closes. `Execute` first runs the
  feature-flag gating (§4 — **pattern P13**), then `ShowModal`. The hub follows the universal
  Hide→Create→Execute→Free→Show child-launch idiom (**pattern P14**) used throughout `MainMenu.pas`
  (the child is created on demand and freed on return). **Close** (`ModalResult = mrCancel`)
  ends the menu; `Execute` returns `False` only when closed via Close, `True` otherwise —
  a return value the caller (`MainMenu`) ignores.

## 2. Data touched
| Table | Read | Write | Notes |
|-------|:----:|:-----:|-------|
| *(none directly)* |  |  | The hub opens **no** dataset and issues **no** ADO command. |

**It reads two configuration values, not tables** — both `TCIniField` objects on the data
module, bound to the **`[SITE]`** section of `InventorySystem.INI` (confirmed in
`DataModule.dfm`):

| INI field (Pascal) | Section / Key | Default | Effect in this hub |
|--------------------|---------------|:-------:|--------------------|
| `data_module.fiPOEDISupport` | `[SITE] POEDISupport` | `True` | When **true**, the **Monthly PO** button is visible; when false it is hidden. |
| `data_module.fiGenerateEDI` | `[SITE] GenerateEDI` | `True` | When **true**: relabels Monthly-PO→**"Manifest Cost"** and reroutes it; hides Assy-Ratio; **shows** the **ASN/Invoice** button. |

These are the **only** data dependencies. All table/proc work happens *inside the child
forms* — see their own specs (`supplier.md`, `logistics.md`, and the not-yet-written
PartsStock/Size/Renban/Assembly/ManifestCost/ASNInvoice specs).

**Triggers on these tables:** none — there are no tables. (The hub is invisible to the
24-trigger inventory.)

## 3. Stored procedures used
**None.** `MasterMaint.pas` contains **zero** references to `Inv_StoredProc`,
`ProcedureName`, `Inv_DataSet`, `ExecProc`, or any `SELECT_/INSERT_/UPDATE_/DELETE_` proc
name (grep-verified). This is the headline structural difference from the CRUD masters:

**Difference vs Supplier / Logistics:** Supplier and Logistics each drive four procs
(`{SELECT,INSERT,UPDATE,DELETE}_…Info`) through the shared ADO object (P6) with the
retry-on-recursion harness (P8) and key off the shared `RecordID` (P9). The hub has
**none** of P1/P2/P3/P4/P5/P6/P7/P8/P9 — it neither queries, dup-checks, timestamps,
resolves names, nor holds a record key. It is the parent that *launches* those modules.

### Procs reachable transitively (for completeness — owned by the children, not the hub)
Each button instantiates one child form; the procs below belong to those children's specs,
not to MasterMaint. Listed so the router contract is traceable end-to-end:

| Button (caption) | Child form (`.pas`) | dpr line | Owns procs (see child spec) |
|------------------|---------------------|:--------:|-----------------------------|
| `&Supplier Master` | `SupplierMaster` | 15 | `*_SupplierInfo`, `SELECT_PartsSupplier` → `supplier.md` |
| `&Logistics Master` (`Button1`) | `LogisticsMaster` | 29 | `*_LogisticsInfo`, `REPORT_MonthlyLogisticsOrders` → `logistics.md` |
| `Si&ze Master` | `SizeMaster` | 16 | `*_SizeInfo`, `SELECT/UPDATE_SizeUsage` |
| `&Parts / Stock Master` | `PartsStockMaster` | 18 | `*_PartsStockInfo`, etc. |
| `&Renban Group Master` | `RenbanGroupMaster` | 31 | `*_RenbanGroup` |
| `&Assembly Detail` (`ForecastDetail_Button`) | `ForecastDetail` | 27 | Assembly-Detail master (form Caption **"Assembly Detail Master"**) |
| `&ASSY / Ratio Master` | `AssyRatioMaster` | 17 | `*_AssyRatioInfo` — **but always hidden** (§4) |
| `&Monthly PO` *(or)* `Manifest Cost` | `MonthlyPOMaster` *(or)* `ManifestCostMaster` | 45 / 49 | `*_AssyMonthlyPO` / `*_ManifestCost` — **flag-routed** (§4) |
| `&ASN/Invoice ` | `ASNInvoice` | 52 | ASN/Invoice resend (810/856) |

> Naming traps to carry into the rebuild: (a) the **"Assembly Detail"** button is wired to
> the unit named **`ForecastDetail`** whose form Caption is "Assembly Detail Master" — the
> unit name is misleading legacy baggage, not a forecast screen. (b) The Logistics button
> is the un-renamed default **`Button1`** / `Button1Click`. (c) The ASN/Invoice control is
> spelled `ASNINVOIVE_Button` (typo'd) in code.

## 4. Business rules & edge cases
The *entire* behavioral content of this unit is the flag-gating in `Execute` plus one
flag re-check in the Monthly-PO click handler. Exact rules:

**A. Monthly-PO visibility** — `if fiPOEDISupport then MonthlyPO_Button.Visible := True
else False`. So a site without PO-EDI support never sees the Monthly-PO / Manifest-Cost
entry at all. (Mirror of the MainMenu gating at `MainMenu.pas:380` — same flag governs the
ASN/Invoice/PO-report buttons there.)

**B. EDI-generating site reconfiguration** — `if fiGenerateEDI then begin
MonthlyPO_Button.Caption := 'Manifest Cost'; AssyRatioMaster_Button.Visible := False;
ASNINVOIVE_Button.Visible := True; end`. I.e. on an EDI-generating site the **Monthly PO**
button is **relabeled "Manifest Cost"** and (per the click handler) opens a *different*
form, the Assy-Ratio button is hidden, and the **ASN/Invoice** button is shown. Note the
old commented gate `//if data_module.fiAssemblerName.AsString = 'CAMEX'` — historically
this branch keyed on the **plant name CAMEX**; DMV replaced it (2005-04-15) with the
`fiGenerateEDI` flag. That is a **latent multi-site signal**: behavior that *was*
site-name-specific is now a per-site boolean (good for the rebuild).

**C. Assy-Ratio is dead in the UI** — immediately after the flag block:
`AssyRatioMaster_Button.Visible := False; // not used yet`. This runs **unconditionally**,
so the **`&ASSY / Ratio Master` button is ALWAYS hidden** regardless of flags. The
`AssyRatioMaster_ButtonClick` handler and the live `AssyRatioMaster` form (dpr line 17)
exist and compile, but are **unreachable from this menu**. (Confirm before "fixing": the
form is live code, just not surfaced here — possibly reachable from another menu, or truly
orphaned. Flag for §8.)

**D. Button overlap bug (cosmetic/latent).** In the `.dfm`, **`AssyRatioMaster_Button` and
`ASNINVOIVE_Button` occupy the identical rectangle** (both `Left=200, Top=174,
Width=121, Height=33`). Because Assy-Ratio is force-hidden (rule C) and ASN/Invoice starts
`Visible=False` and is only shown on EDI sites (rule B), they never both show — so the
overlap is masked at runtime. Do not preserve this collision in the rebuilt layout; it is
an artifact of the Assy-Ratio→ASN/Invoice repurposing.

**E. Monthly-PO click is itself flag-routed** — `MonthlyPO_ButtonClick` re-reads
`fiGenerateEDI`: **true → open `ManifestCostMaster_Form`**, **false → open
`MonthlyPOMaster_Form`**. So one button, two destinations, decided by the same flag that
set its caption in `Execute`. (The flag is read **twice** — once for the caption in
`Execute`, once for the route here. If the INI were edited mid-session the two could
diverge, but in practice the menu is re-created each visit.)

**F. Child-launch contract** — every button handler is the same shape:
`Hide; <Child>_Form := T<Child>_Form.Create(self); <Child>_Form.Execute; <Child>_Form.Free;
Show;`. The hub **hides itself** while a child is modal-open and **shows** again on return.
`SupMaster` and `ASNINVOIVE` wrap this in `try…finally Show`; the other six do **not** —
so if a child's `Execute` raises, those six leave the hub **hidden** (the exception bubbles
to `MainMenu`'s `Application.OnException` `CatchAll`, which `ShowMessage`s, logs, and calls
`Application.Terminate`). Inconsistent error safety worth normalizing (§5/§8).

**G. No table/calendar/date math, no status transitions, no validation.** Nothing to port
beyond the flag gating and the routing table. Audit-log: the hub itself writes **no**
`LogActLog` entry (children do).

## 5. UI / UX notes
- **Layout:** a fixed 2-column × ~5-row button grid on a non-resizable dialog
  (`bsDialog`, 370×323 client). Buttons (label → handler → destination):
  - Col 1: **Supplier Master**, **Logistics Master**, **Assembly Detail**, **Renban Group
    Master**; Col 2: **Parts / Stock Master**, **Size Master**, **ASSY / Ratio Master**
    *(hidden)* / **ASN/Invoice** *(same slot)*, **Monthly PO / Manifest Cost**.
  - A full-width **Close** button (`ModalResult=2`/`mrCancel`).
  - Accelerator keys via `&` (Alt-S Supplier, Alt-Z Size, Alt-P Parts, Alt-L Logistics,
    Alt-A Assembly Detail, Alt-R Renban, Alt-M Monthly, Alt-C Close). Tab order set.
- **No filters, no search, no grid, no entry fields.** It is purely a launcher.
- **Conditional surface:** the visible button set is per-site (flags A/B/C). A non-EDI,
  non-PO site shows the smallest menu (no Monthly-PO, no ASN/Invoice, no Assy-Ratio).
- **Keep vs modernize:**
  - *Keep* the information architecture: master-data lives under one section with these
    sub-resources. This is effectively the **nav/sidebar for the masters area**.
  - *Modernize:* replace the modal Hide/Show dance with normal routed pages
    (`/masters/...`); drive button/link visibility from **policy + feature flags** instead
    of `.Visible` toggles; fix the dead Assy-Ratio entry and the overlapping-button bug;
    normalize the inconsistent `try…finally` so a child error never strands the menu;
    rename the misleading identifiers (`Button1`, `ASNINVOIVE`, the "Date" typos) in the
    new code.

## 6. Target design  *(Rails primary)*
This module becomes **routing + navigation + feature-flag policy**, not a model/controller
pair of its own.

- **Models:** **none for the hub.** (Its children own the models: `Supplier`, `Logistics`,
  `Size`, `PartsStock`, `RenbanGroup`, `AssemblyDetail`, `AssyRatio`, `MonthlyPo` /
  `ManifestCost`, `AsnInvoice`.)
- **Routes/controllers:** a `namespace :masters` (or a `MastersController#index` landing
  page) listing links to the child resources:
  `resources :suppliers, :logistics, :sizes, :parts_stocks, :renban_groups,
  :assembly_details, :assy_ratios, :monthly_pos, :manifest_costs` and an
  ASN/Invoice-resend action. The hub itself is just `masters#index`.
- **Views/components:** one **index/landing** (or a persistent left-nav for the masters
  area) rendering the link set, with each link gated by feature flags (below). Replace the
  10 buttons with links/cards.
- **Feature flags (the real port target):** model `[SITE] POEDISupport` and
  `[SITE] GenerateEDI` as **per-site settings** (a `Site#po_edi_support?` /
  `Site#generates_edi?`, or a settings table) — *not* global constants, because the
  rebuild is multi-site (see §8 / [[project-multisite]]). The view rules become:
  - `po_edi_support?` → show **Monthly PO** link.
  - `generates_edi?` → Monthly-PO link **label = "Manifest Cost"** and **target =
    ManifestCost** (else label "Monthly PO", target MonthlyPo); show **ASN/Invoice** link;
    hide Assy-Ratio (which is already hidden anyway).
  - Centralize these in one policy object so MainMenu and the masters nav stay consistent
    (today the same flags are re-checked in both `MasterMaint.Execute` and
    `MainMenu.Configure` at `:380` — a single source of truth removes that duplication).
- **Services / Reports:** none for the hub.

## 7. Migration plan for this module
- [ ] **Stage 1 — read-only parity:** build the masters **landing/nav** page with links to
      each child module (which can themselves be stage-1 proc-wraps). Reproduce the exact
      visible-link set per the two flags. No writes — there were none to begin with.
- [ ] **Stage 2 — n/a (no writes).** Instead: wire the **Monthly-PO ⇆ Manifest-Cost**
      flag routing and the **ASN/Invoice** visibility through the feature-flag policy so it
      matches legacy per-site behavior; decide the fate of the dead **Assy-Ratio** entry
      (surface it or formally retire it — §8).
- [ ] **Stage 3 — reimplement / cleanup:** the hub has no proc/trigger logic to re-home;
      "stage 3" here is **IA + policy hardening** — single feature-flag source of truth,
      per-site scoping, fix the overlap/dead-button artifacts, drop the modal Hide/Show
      pattern entirely (RESTful pages), rename the misleading legacy identifiers.

## 8. Open questions for the user (domain expert)
1. **Per-site feature flags:** `POEDISupport` and `GenerateEDI` live in `[SITE]` of the
   single-site INI (both default `True`). In multi-site, are these **per-site** toggles
   (most likely) — i.e. some plants generate EDI / use PO-EDI and others don't — and where
   should they live (a `sites` row, a settings table)? This decides whether the masters nav
   is computed per-current-site. (Same multi-site lens as supplier.md §8.1 / logistics.md
   §8.1.)
2. **Dead Assy-Ratio entry:** `&ASSY / Ratio Master` is **force-hidden** here
   (`Visible := False; // not used yet`) yet `AssyRatioMaster` is a **live** unit (dpr 17)
   with its own `*_AssyRatioInfo` procs. Is Assy-Ratio maintenance reachable from another
   menu, is it deprecated, or was it meant to come back? Should the rebuilt masters section
   include it?
3. **"Assembly Detail" vs "Forecast Detail":** the button labeled **Assembly Detail** opens
   the unit `ForecastDetail` (form Caption "Assembly Detail Master"). Confirm this screen is
   an **assembly/BOM detail master**, not a forecast screen, so the rebuilt resource is
   named correctly (`assembly_details`, not `forecast_details`).
4. **Monthly-PO ⇆ Manifest-Cost relationship:** on an EDI-generating site the single
   Monthly-PO button becomes **Manifest Cost** and opens a *different* form. Are these two
   genuinely alternative tools (an EDI site never needs the Monthly-PO editor and vice
   versa), or should the rebuilt app expose **both** independently rather than overloading
   one nav slot?
5. **Error-safety inconsistency:** six of the eight launch handlers omit the `try…finally
   Show`, so a child exception leaves the hub hidden and bubbles to a terminate-the-app
   handler. Is the app-terminate-on-unhandled-exception behavior intended to be preserved,
   or should the rebuild simply show an error and keep the user on the masters page?
6. **Menu naming:** the MainMenu entry/handler is `DateMaint_*` and labels say "Master
   Date Maintenance" — confirmed typos for "**Data**". Safe to standardize on "Master Data"
   in the rebuild?

## 9. Test cases / parity checks
Because the hub has no data, parity is about **which entries appear, how one entry is
labeled, and where each entry leads** under the two flags.

- **Both flags true** (`POEDISupport=True`, `GenerateEDI=True` — the default site): visible
  master entries = Supplier, Logistics, Size, Parts/Stock, Renban Group, Assembly Detail,
  **Manifest Cost** (relabeled), **ASN/Invoice**. **Hidden:** Assy-Ratio (always),
  Monthly-PO label is *not* shown as "Monthly PO" (it shows as "Manifest Cost").
- **`POEDISupport=True`, `GenerateEDI=False`:** Monthly-PO entry visible **labeled "Monthly
  PO"**; ASN/Invoice **hidden**; Assy-Ratio hidden. Clicking Monthly PO opens
  `MonthlyPOMaster` (not ManifestCost).
- **`POEDISupport=False`** (regardless of `GenerateEDI`): the Monthly-PO/Manifest-Cost
  entry is **absent**. (If `GenerateEDI=True`, ASN/Invoice is still shown and the Monthly-PO
  caption is set to "Manifest Cost" by `Execute`, but the button is hidden by rule A — so
  net: no Monthly/Manifest entry; ASN/Invoice present.)
- **Routing parity:** clicking each visible entry opens exactly the mapped child form
  (table in §3); **Monthly PO/Manifest Cost** must route by `GenerateEDI`
  (`true→ManifestCost`, `false→MonthlyPO`) — assert both branches.
- **Assy-Ratio entry never visible** from this menu under any flag combination (rule C).
- **Close** ends the menu and returns control to the main menu; legacy `Execute` returns
  `False` on Close / `True` otherwise (caller ignores it) — the rebuilt nav simply returns
  to the main app shell.
- **No DB side effects:** opening, navigating, and closing the masters menu writes **no**
  rows and emits **no** `LogActLog` audit entry from the hub itself (only child forms log).
