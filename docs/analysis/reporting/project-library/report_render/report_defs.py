# report_defs — the CONFIG for the M3 render harness. A report = a render def + a .sql (data-driven;
# adding a flat-table report = one entry + one .sql). Mirrors m3-render-architecture.md §2.2.
#
# THE LIVE SET (data-driven prune — m3-report-usage-prune.md): of the 22 report families, 18 NEVER RAN in
# the available usage signal. These 4 are the entire live set actually built here:
#   daily_shipping_assy  R3  REPORT_DailyShippingAssy   285 runs/yr (the one report run on a real prod day)
#   invoice_summary      R9  REPORT_INVOICESSummary       1 run/yr  (BILLING — D6 window-aware, corrected)
#   forecast_detail      R18 REPORT_ForecastDetail        1 run/yr  (faithful)
#   lot_location         R12 REPORT_PLANTLotLocation[W]   1 run/yr  (faithful; CONFIRMED LIVE, not D9 NUMMI)
# The 18 SKIP reports (deprecate-by-deletion at legacy retirement, NOT ported) are listed in SKIP_REPORTS.
#
# Each def has:
#   key            report key (matches the SPIKE log + the test).
#   title          legacy cell [1,1] title literal.
#   out_name       legacy SaveAs <name>; the yyyymmddhhmmss00 + '.xlsx' suffix is appended by write_output.
#   proc           the legacy REPORT_* proc this wraps (provenance).
#   query          the .sql KEY (sql/<key>.sql); the faithful↔corrected seam resolves it (resolve_query).
#   params         ordered param names the .sql ? markers bind, in declared order (NQ name-binding port).
#   render         "declarative" (build_declarative_plan) | a custom-plan function NAME in code.py.
#   header_cells / header_row / columns / first_data_row / totals  — the declarative layout (see code.py).
#
# Faithful↔corrected (m3-render-architecture §2.4): QUERY_VARIANT picks faithful vs corrected per the
# `query` KEY -> resolve_query maps "<key>" (corrected/default) or "<key>_faithful" (legacy). Only
# invoice_summary has a _faithful variant (the D6 window-blind legacy); the other 3 are faithful as-is.

QUERY_VARIANT = "corrected"          # "corrected" (default, ships) | "faithful" (parity proof / parallel)

# Per-report override so faithful and corrected can run SIMULTANEOUSLY for the SAME report during parallel
# run (the parity proof needs both at once — m3-render-architecture §9). resolve_query(key, override=...).
_FAITHFUL_EXISTS = {"invoice_summary"}   # keys that have a sql/<key>_faithful.sql


def resolve_query(query_key, override=None):
    """Map a def's `query` KEY to the .sql KEY to load, honoring QUERY_VARIANT (or a per-call override).
    'faithful' -> '<key>_faithful' iff that variant exists, else the base key (the 3 already-faithful
    reports have no _faithful twin -> base). 'corrected' -> the base key."""
    variant = override or QUERY_VARIANT
    if variant == "faithful" and query_key in _FAITHFUL_EXISTS:
        return query_key + "_faithful"
    return query_key


# Excel number formats lifted from the .pas (cited per use).
_MONEY = "$#,##0.0000_);[Red]($#,##0.0000)"     # INVOICE Summary Unit Cost/Item Total (MainMenu.pas:3648)
_PNFMT = "############"                          # part-number column format (legacy Columns[c].NumberFormat)


REPORTS = {

    # ---------------------------------------------------------------------------------------------
    # R3 — Daily Shipping ASN (Assy). FAITHFUL. Layout: MainMenu.pas:3165-3199 (DailyASNReportClick).
    # The SOUND shipping path (clean IN_ASN_ID join, no fan-out). 2-col detail (Desc write commented out).
    # ---------------------------------------------------------------------------------------------
    "daily_shipping_assy": {
        "key": "daily_shipping_assy",
        "title": "ASN (Assy Part Numbers)",          # cell [1,1] (MainMenu.pas:3172)
        "out_name": "DailyShippingAssy",             # SaveAs (MainMenu.pas:3199)
        "proc": "REPORT_DailyShippingAssy",
        "query": "daily_shipping_assy",
        "params": ["PDate"],                          # @PDate only; the caller's extra @Line is DROPPED (port)
        "render": "declarative",
        "header_cells": [
            (1, 1, "title"),                                          # [1,1] title
            (2, 1, ("param", "PDate", "yyyy/mm/dd")),                 # [2,1] production date, yyyy/mm/dd
            (2, 2, ("header", "StartEnd")),                           # [2,2] 'Start Seq:..../End Seq:....'
            (2, 3, ("header", "Vehicle Count", "Vehicle Count:")),    # [2,3] 'Vehicle Count:<n>'
        ],
        "header_row": 3,                              # column-title row (legacy row 3)
        "columns": [                                  # (header_text, query_col, width, num_fmt)
            ("Part Number", "Part Number", 17, _PNFMT),   # [3,1] w17 fmt ############ (MainMenu.pas:3180-3182)
            ("Qty",         "PQty",        30, None),      # [3,2] 'Qty' w30 (Desc col commented out, :3190)
        ],
        "first_data_row": 4,                          # z := 4
        "totals": None,
        "header_query": "daily_shipping_header_assy",  # the Start/End/Vehicle-Count header band (single row)
        "pdf_template": None,
    },

    # ---------------------------------------------------------------------------------------------
    # R9 — INVOICE Summary. CORRECTED (D6 window-aware, DEFAULT). Layout: MainMenu.pas:3624-3667.
    # D6 DIVERGENCE (documented, David-LOCKED): corrected via fn_ManifestCostAt (invoice_summary.sql);
    # faithful window-blind legacy behind the seam (invoice_summary_faithful.sql). BILLING -> numbers matter.
    # 5-col detail + a per-row Item Total (=Qty*UnitCost) + a grand INVOICE TOTAL band (2 blank rows below).
    # ---------------------------------------------------------------------------------------------
    "invoice_summary": {
        "key": "invoice_summary",
        "title": "INVOICE (Assy Part Numbers)",      # cell [1,1] (MainMenu.pas:3632)
        "out_name": "INVOICESummary",                # SaveAs (MainMenu.pas:3680)
        "proc": "REPORT_INVOICESSummary",
        "query": "invoice_summary",                  # resolve_query -> _faithful under QUERY_VARIANT='faithful'
        "params": ["PDate"],
        "render": "declarative",
        "header_cells": [
            (1, 1, "title"),                                         # [1,1] title
            (2, 1, ("param", "PDate", "INVOICE DATE:", "p")),        # [2,1] 'INVOICE DATE:'+date (MainMenu:3636)
        ],
        "header_row": 3,
        "columns": [                                  # legacy [3,1..5] (MainMenu.pas:3638-3653)
            ("Part Number",     "PartNumber", 17, _PNFMT),   # [3,1] w17 ############
            ("Manifest Number", "Manifest",   17, None),     # [3,2] w17
            ("Qty",             "ShipQty",    30, None),     # [3,3] w30
            ("Unit Cost",       "UnitPrice",  30, _MONEY),   # [3,4] w30 money
            ("Item Total",      None,         30, _MONEY),   # [3,5] w30 money — COMPUTED (totals.item)
        ],
        "first_data_row": 4,                          # z := 4
        "totals": {
            "item": {"col": 5, "factors": ("ShipQty", "UnitPrice"), "num_fmt": _MONEY},  # E=C*D per row
            "gap": 2,                                 # INC(z);INC(z) -> 2 blank rows (MainMenu.pas:3658-3659)
            "cells": [(3, "INVOICE TOTAL"), (5, "__grand__")],   # [tz,3]='INVOICE TOTAL'; [tz,5]=SUM
            "grand_num_fmt": _MONEY,
        },
        "header_query": None,
        "pdf_template": None,
    },

    # ---------------------------------------------------------------------------------------------
    # R18 — Forecast Detail. FAITHFUL. CUSTOM render (group-break + blank separator — build_forecast_detail_plan).
    # Layout + loop: MainMenu.pas:3744-3794. 8 cols; part#/eff-month only on a group's first row; one blank
    # separator row before each part group; ratios are integers.
    # ---------------------------------------------------------------------------------------------
    "forecast_detail": {
        "key": "forecast_detail",
        "title": "Forecast Detail Report",           # cell [1,1] (MainMenu.pas:3746)
        "out_name": "ForecastDetail",                # SaveAs (MainMenu.pas:3799)
        "proc": "REPORT_ForecastDetail",
        "query": "forecast_detail",
        "params": [],                                 # no params
        "render": "build_forecast_detail_plan",
        "header_row": 3,
        "columns": [                                  # (header_text, query_col, width, num_fmt) — MainMenu:3748-3766
            ("Assembly Part#",    "VC_ASSY_PART_NUMBER_CODE", 17, None),
            ("Effective Month",   "VC_EFFECTIVE_MONTH",       17, None),
            ("BC Code",           "VC_BROADCAST_CODE",        10, None),
            ("Tire Part Number",  "VC_TIRE_PART_NUMBER_CODE", 17, _PNFMT),
            ("Tire Share Ratio",  "IN_TIRE_RATIO",            17, None),
            ("Wheel Part Number", "VC_WHEEL_PART_NUMBER_CODE",17, _PNFMT),
            ("Wheel Share Ratio", "IN_WHEEL_RATIO",           17, None),
            ("Forecast Share",    "IN_RATIO",                 17, None),
        ],
        "first_data_row": 4,                          # z := 4 (first INC -> z=5)
        "header_query": None,
        "pdf_template": None,
    },

    # ---------------------------------------------------------------------------------------------
    # R12 — Lot Location (PLANT). FAITHFUL. CONFIRMED LIVE (REPORT_PLANTLotLocation[W], NOT the D9 NUMMI
    # twin — the live LotLocationClick handler calls the PLANT procs). CUSTOM render (two procs + line/size
    # group headers + conditional cells + date reformat — build_lot_location_plan). Layout MainMenu:716-861.
    # 7 cols. Runs the WHEELS proc then the TIRES proc into one sheet. Print chrome (page breaks/borders/
    # alignment) is NOT reproduced (presentation, non-gating; m3-render-architecture §4).
    # ---------------------------------------------------------------------------------------------
    "lot_location": {
        "key": "lot_location",
        "title": "Lot Location",                     # actual [1,1] = '<PlantName> Lot Location' (built in plan)
        "out_name": "LotLocation",                   # SaveAs <PlantName>LotLocation (plant prefix from params)
        "proc": "REPORT_PLANTLotLocation / REPORT_PLANTLotLocationW",
        "query": "lot_location",                     # special: resolves to TWO .sql (wheels + tires) in driver
        "queries": {"wheels": "lot_location_wheels", "tires": "lot_location_tires"},
        "params": [],                                 # no proc params; PlantName/AssemblerName from gateway cfg
        "render": "build_lot_location_plan",
        "header_query": None,
        "pdf_template": None,
    },
}


# Header-band query KEYS (single-row aggregates feeding header_cells; not a detail table). Daily Shipping
# Assy reads Start/End/Vehicle Count from the ASN header for the date (the proc returns them per detail row,
# but the header band wants ONE aggregated row — MIN start / MAX end / SUM qty, like daily_shipping_header).
HEADER_QUERIES = {"daily_shipping_header_assy"}


# =================================================================================================
# SKIP LIST — the 18 dead reports + their Excel/OLE paths. NOT ported (deprecate-by-deletion at the
# legacy-retirement step). Source: m3-report-usage-prune.md (18 of 22 never ran) + m3-report-inventory.md
# (the per-report Excel mechanism) + the D9 deprecations. This is a NOTE (don't delete legacy code here).
# =================================================================================================
SKIP_REPORTS = [
    # (R#, report, legacy proc, Excel/OLE path retired, why-skip)
    ("R1",  "Daily Shipping (T/W)",        "REPORT_DailyShipping",            "OLE-template ReportTemplate.xls -> DailyShippingTW<ts>.xls",        "T/W fan-out + 72% orphan-drop bug; no run in usage signal; superseded by the sound Assy path (R3)"),
    ("R2",  "Daily Shipping Range (T/W)",  "REPORT_DailyShippingRange",       "OLE-template -> DailyShippingRangeTW<ts>.xls",                      "same T/W defects; ad-hoc; no run signal"),
    ("R4",  "Monthly Shipping ASN (Assy)", "REPORT_MonthlyShippingAssy",      "OLE-template -> DailyShippingAssy<ts>.xls (name collides w/ R3)",   "monthly; no isolated run signal (shares R3 string); same proc family as R3 — fold in later if needed"),
    ("R5",  "Daily Supplier Order",        "REPORT_DailySupplierOrders/Cost", "OLE-template -> DailySupplierOrder<ts>.xls",                        "not run 06-19 (orders went via the ORDER/ORDERF flow); no run signal"),
    ("R6",  "Monthly Supplier Order",      "REPORT_MonthlySupplierOrders/Cost","OLE-template -> SupplierOrder<ts>.xls",                            "monthly; no run signal; D12 ship-date-range bug"),
    ("R7",  "Monthly Logistics Order",     "REPORT_MonthlyLogisticsOrders",   "OLE-template -> Logistics<ts>.xls",                                 "monthly; no run signal; D12"),
    ("R8",  "Monthly Supplier Invoice",    "REPORT_MonthlySupplierInvoices",  "OLE-template -> SupplierInvoice<ts>.xls",                           "monthly; no run signal"),
    ("R10", "Monthly INVOICE Summary",     "REPORT_MonthlyINVOICESSummary",   "OLE-template -> MonthlyINVOICESummary<ts>.xls",                     "monthly; D6 (same class as R9); no run signal — defer (R9 proves the pattern)"),
    ("R11", "Logical Inventory",           "REPORT_LogicalInventory",         "OLE-template -> LogicalInventory<ts>.xls",                          "ad-hoc; SELECT* exposes ledger IN_QTY (M2 trap); no run signal"),
    ("R13", "Empty Container",             "REPORT_EmptyContainer",           "OLE-template -> EmptyContainer<ts>.xls",                            "P3 niche; lean-SKIP in prune; no run signal"),
    ("R14", "Past-Due / Late FRS",         "REPORT_LATEFRS",                  "OLE-template -> PastDueFRS<ts>.xls",                                "exception report (sparse by nature); SELECT*; no run signal"),
    ("R15", "PO Report",                   "REPORT_PO",                       "OLE-template -> POReport<ts>.xls",                                  "CAMEX-PO oriented (CAMEX decommissioned, D9 context); lean-SKIP"),
    ("R16", "Forecast Parts Summary",      "REPORT_ForecastPartsSummary",     "OLE-template -> ForecastPartsSummary<ts>.xls",                      "periodic; WINDOW (GetDate-relative); SUM-NULL; no run signal"),
    ("R17", "Forecast Assy Summary",       "REPORT_ForecastSummary",          "OLE-template -> ForecastAssySummary<ts>.xls",                       "periodic; WINDOW (GetDate-relative); SELECT*; no run signal"),
    ("R19", "Forecast vs Usage",           "(forecast/usage proc)",           "OLE-template -> ForecastvsUsage<ts>.xls",                           "analytical/ad-hoc; lean-SKIP; no run signal"),
    ("R20", "Unused Tire Part Numbers",    "REPORT_UnusedTirePartNumbers",    "OLE-template -> TireWithoutAssembly<ts>.xls",                       "master-data hygiene (rare); lean-SKIP"),
    ("R21", "Unused Wheel Part Numbers",   "REPORT_UnusedWheelPartNumbers",   "OLE-template -> WheelWithoutAssembly<ts>.xls",                      "master-data hygiene (rare); has the D11 tire/wheel-column bug; lean-SKIP"),
    ("R22", "InvMgmt QReport",             "(none — Grid_ClientDataSet)",     "TQuickRep.Preview (on-screen, NO Excel)",                           "on-screen QuickReport; no .xls to retire; no REPORT log signal; lowest-value rebuild"),
]

# Already-decided D9 deprecations (not counted among the 18 — out of rebuild scope entirely):
SKIP_DEPRECATED_D9 = [
    ("—", "ASN with Cost",            "REPORT_ASNWithCost",            "ASNwithCost<ts>.xls",     "D9 deprecated"),
    ("—", "NUMMI Lot Location[W]",    "REPORT_NUMMILotLocation[W]",    "NUMMILotLocation<ts>.xls","D9 deprecated (NUMMI decommissioned) — NOT the live R12 Lot Location"),
    ("C8","Forecast CAMEX",           "(ForecastCamexreport.pas)",     "<sup>-<code>-CFForecast.xls","D9 deprecated (CAMEX gone)"),
]
