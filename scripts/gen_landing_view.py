#!/usr/bin/env python3
"""
gen_landing_view.py — build the Perspective landing-page HUB view on the tai-* theme,
with REAL KPI data bound from the live spike DB.

Emits Home/Home: site-aware top bar -> KPI strip -> module-card grid + a "needs
attention" rail, styled through the suite theme tokens (var(--tai-*)). KPI/rail/module
counts are loaded by ONE view-level binding (custom.kpi) whose script transform runs
the Home/kpiSummary + kpiTopShort Named Queries (single-combined-load pattern), then
labels bind to view.custom.kpi.<key>. Also patches the /home route + session theme,
and wires a Theme toggle button.

KPIs are honest to what the data actually supports (verified against the live DB
2026-06-18): open orders (INV_OPEN_ORDER_INF), active ASNs (INV_ASN_MST status A),
shipments (INV_SHIPPING_INF), parts short (INV_PARTS_STOCK_MST IN_QTY<=0), stocktake
entries (INV_STOCKTAKING_INF). There is NO EDI-error / receipts-pending / ASN-due-date
signal in the schema, so those mockup tiles were dropped rather than faked.

# IG81-COMPAT: ia.* components as on 8.1.52; meta.domId/name; props.style CSS object;
#              data via system.db.runNamedQuery (Home/kpiSummary + kpiTopShort; created in Designer).
# IG83-TODO: move the KPI load to a Named Query; revisit component prop schemas on 8.3.
"""
import json, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _ignenv import GATEWAY_PROJECT_DIR, PERSPECTIVE_DIR, DB_CONN   # noqa: E402 — centralized (repo-split-plan §4.C/§4.D)

PROJ_DIR  = GATEWAY_PROJECT_DIR
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PERSP     = PERSPECTIVE_DIR
DB        = DB_CONN

# ---- token shorthands ------------------------------------------------------------------
BG, SURF, SURF2 = "var(--tai-bg)", "var(--tai-surface)", "var(--tai-surface-2)"
BORDER          = "var(--tai-border)"
TXT, MUT, FAINT = "var(--tai-text)", "var(--tai-text-muted)", "var(--tai-text-faint)"
ACCENT, ASOFT, ONACC = "var(--tai-accent)", "var(--tai-accent-soft)", "var(--tai-on-accent)"
DANGER, WARN, INFO, OK = ("var(--tai-status-danger)", "var(--tai-status-warning)",
                          "var(--tai-status-info)", "var(--tai-status-success)")
RAD, RADC = "var(--tai-radius)", "var(--tai-radius-control)"
MONO       = "var(--tai-font-mono)"
SHADOW     = "var(--tai-shadow-card)"
CARD = {"backgroundColor": SURF, "border": "1px solid " + BORDER, "borderRadius": RAD,
        "padding": "16px", "boxShadow": SHADOW}

# ---- KPI data load: one binding -> view.custom.kpi (script transform, runNamedQuery) ----
KPI_KEYS = ["openOrders_num", "openOrders_foot", "asns_num", "asns_foot", "shipments_num",
            "partsShort_num", "partsShort_foot", "stocktake_num",
            "rail_short_title", "rail_short_meta", "rail_belowlot_title", "rail_belowlot_meta",
            "rail_stocktake_title", "rail_stocktake_meta",
            "rail_renban_title", "rail_renban_meta",      # P4 RENBAN_COLLISION attention surface
            "mod_orders", "mod_shipasn", "mod_stock", "mod_master"]

KPI_TRANSFORM = (
    "\tsystem.util.getLogger(\"KPI\").info(\"Home KPI load via runNamedQuery\")\n"
    "\tdata = system.db.runNamedQuery(\"Home/kpiSummary\", {})\n"
    "\ttop  = system.db.runNamedQuery(\"Home/kpiTopShort\", {})\n"
    "\tdef g(col): return int(data.getValueAt(0, col))\n"
    "\topenOrders=g(\"OpenOrders\"); openSuppliers=g(\"OpenSuppliers\"); asns=g(\"ActiveAsns\")\n"
    "\tasnLines=g(\"AsnLines\"); shipments=g(\"Shipments\"); partsTotal=g(\"PartsTotal\")\n"
    "\tpartsShort=g(\"PartsShort\"); belowLot=g(\"PartsBelowLot\"); stocktake=g(\"StocktakeEntries\"); suppliers=g(\"Suppliers\")\n"
    "\trenbanCollisions=g(\"RenbanCollisions\")\n"
    "\ttopShort = (\"%s (qty %d)\" % (top.getValueAt(0,\"PartName\"), int(top.getValueAt(0,\"Qty\")))) if top.rowCount else \"none\"\n"
    "\tdef c(n): return \"{:,}\".format(n)\n"
    "\treturn {\n"
    "\t\t\"openOrders_num\": c(openOrders), \"openOrders_foot\": \"across %d suppliers\" % openSuppliers,\n"
    "\t\t\"asns_num\": c(asns), \"asns_foot\": \"%s ASN lines\" % c(asnLines),\n"
    "\t\t\"shipments_num\": c(shipments),\n"
    "\t\t\"partsShort_num\": c(partsShort), \"partsShort_foot\": \"of %d parts tracked\" % partsTotal,\n"
    "\t\t\"stocktake_num\": c(stocktake),\n"
    "\t\t\"rail_short_title\": \"%d parts at/below zero stock\" % partsShort, \"rail_short_meta\": \"Lowest: %s\" % topShort,\n"
    "\t\t\"rail_belowlot_title\": \"%d parts below one lot\" % belowLot, \"rail_belowlot_meta\": \"Master Data - review reorder\",\n"
    "\t\t\"rail_stocktake_title\": \"%d stocktake entries to review\" % stocktake, \"rail_stocktake_meta\": \"Stocktaking - cycle counts\",\n"
    "\t\t\"rail_renban_title\": (\"%d renban collision(s) to resolve\" % renbanCollisions) if renbanCollisions else \"No renban collisions\", \"rail_renban_meta\": \"Order - Renban Breakdown\" if renbanCollisions else \"all renbans clear\",\n"
    "\t\t\"mod_orders\": \"%s open\" % c(openOrders), \"mod_shipasn\": \"%s ASNs\" % c(asns),\n"
    "\t\t\"mod_stock\": \"%d entries\" % stocktake, \"mod_master\": \"%d parts - %d suppliers\" % (partsTotal, suppliers),\n"
    "\t}\n"
)

# ---- component helpers -----------------------------------------------------------------
def lbl(name, text, style=None, position=None, domId=None, bind=None):
    meta = {"name": name}
    if domId: meta["domId"] = domId
    props = {"text": text}
    if style: props["style"] = style
    c = {"type": "ia.display.label", "meta": meta, "props": props}
    if bind:
        c["propConfig"] = {"props.text": {"binding": {"type": "property",
            "config": {"path": "view.custom.kpi.%s" % bind}}}}
    if position: c["position"] = position
    return c

def flex(name, children, direction="column", gap=0, style=None, justify=None,
         align=None, wrap=None, position=None, domId=None):
    props = {"direction": direction}
    if gap: props["gap"] = gap
    if justify: props["justify"] = justify
    if align: props["alignItems"] = align
    if wrap: props["wrap"] = wrap
    if style: props["style"] = style
    meta = {"name": name}
    if domId: meta["domId"] = domId
    c = {"type": "ia.container.flex", "meta": meta, "props": props, "children": children}
    if position: c["position"] = position
    return c

def icon(name, path, color=ACCENT, size="22px"):
    return {"type": "ia.display.icon", "meta": {"name": name},
            "props": {"path": path, "color": color, "style": {"width": size, "height": size}}}

def spacer(name="spacer"):
    return {"type": "ia.container.flex", "meta": {"name": name}, "props": {}, "position": {"grow": 1}}

# ---- top bar (polished spacing) --------------------------------------------------------
def topbar():
    logo = flex("logo", [lbl("mark", "TT", {"color": ONACC, "fontWeight": "700", "fontSize": "15px"})],
                align="center", justify="center",
                style={"backgroundColor": ACCENT, "borderRadius": "10px",
                       "width": "38px", "height": "38px"}, position={"shrink": 0, "basis": "38px"})
    brand = flex("brand", [
        logo,
        flex("brandtext", [
            lbl("name", "Inventory", {"fontWeight": "700", "fontSize": "16px", "color": TXT}),
            lbl("sub", "Supply & EDI", {"fontSize": "11px", "color": MUT}),
        ], gap=1),
    ], direction="row", gap=12, align="center", position={"shrink": 0, "basis": "auto"})

    site = flex("sitePill", [
        flex("dot", [], style={"backgroundColor": OK, "width": "7px", "height": "7px",
                               "borderRadius": "50%"}, position={"shrink": 0, "basis": "7px"}),
        lbl("siteName", "Tupelo, MS", {"fontWeight": "500", "color": TXT, "fontSize": "13px"}),
        lbl("siteCode", "· TT", {"color": MUT, "fontSize": "13px"}),
    ], direction="row", gap=8, align="center",
        style={"backgroundColor": SURF, "border": "1px solid " + BORDER, "borderRadius": "999px",
               "padding": "7px 14px", "marginLeft": "6px"}, position={"shrink": 0, "basis": "auto"})

    search = {"type": "ia.input.text-field", "meta": {"name": "search", "domId": "home-search"},
              "props": {"placeholder": "Search parts, orders, ASNs…", "style": {"borderRadius": RADC}},
              "position": {"shrink": 0, "basis": "250px"}}

    theme_btn = {"type": "ia.input.button",
                 "meta": {"name": "themeToggle", "domId": "home-theme-toggle"},
                 "props": {"text": "◐ Theme"},
                 "position": {"shrink": 0, "basis": "104px"},
                 "events": {"component": {"onActionPerformed": {"type": "script", "scope": "G",
                     "config": {"script":
                         "\tcur = self.session.props.theme\n"
                         "\tself.session.props.theme = 'tai-light' if cur == 'tai-dark' else 'tai-dark'\n"}}}}}

    avatar = flex("avatar", [lbl("ini", "DV", {"fontWeight": "600", "color": TXT, "fontSize": "13px"})],
                  align="center", justify="center",
                  style={"backgroundColor": SURF2, "border": "1px solid " + BORDER,
                         "borderRadius": "50%", "width": "38px", "height": "38px"},
                  position={"shrink": 0, "basis": "38px"})

    return flex("topbar", [brand, site, spacer("tb_sp"), search, theme_btn, avatar],
                direction="row", gap=16, align="center",
                style={"padding": "16px 0 18px"}, position={"shrink": 0, "basis": "auto"})

# ---- KPI strip (real, bound) -----------------------------------------------------------
# (label, num_key, num_color, foot_static, foot_bind, foot_color, tick)
KPIS = [
    ("Open orders",  "openOrders", TXT,    None,                 "openOrders_foot", MUT,    INFO),
    ("Active ASNs",  "asns",       TXT,    None,                 "asns_foot",       FAINT,  INFO),
    ("Shipments",    "shipments",  TXT,    "shipping records",   None,              FAINT,  OK),
    ("Parts short",  "partsShort", DANGER, None,                 "partsShort_foot", DANGER, DANGER),
    ("Stocktake",    "stocktake",  TXT,    "entries to review",  None,              WARN,   WARN),
]
def kpi_tile(label_text, num_key, num_color, foot_static, foot_bind, foot_color, tick):
    head = flex("h", [
        flex("tick", [], style={"backgroundColor": tick, "width": "8px", "height": "8px",
                                "borderRadius": "2px"}, position={"shrink": 0, "basis": "8px"}),
        lbl("l", label_text, {"fontSize": "12px", "color": MUT, "fontWeight": "500"}),
    ], direction="row", gap=7, align="center")
    foot = lbl("f", foot_static or "—",
               {"fontSize": "11.5px", "color": foot_color,
                "fontWeight": "600" if foot_color in (DANGER, WARN) else "400"},
               bind=foot_bind)
    return flex("kpi_" + num_key, [
        head,
        lbl("n", "—", {"fontSize": "27px", "fontWeight": "700", "fontFamily": MONO,
                       "color": num_color, "margin": "8px 0 4px"}, bind=num_key + "_num"),
        foot,
    ], style=CARD, position={"grow": 1, "basis": "0px"})

def kpi_strip():
    return flex("kpis", [kpi_tile(*k) for k in KPIS], direction="row", gap=14,
                domId="home-kpis", position={"shrink": 0, "basis": "auto"})

# ---- module grid -----------------------------------------------------------------------
# (title, desc, icon, status_color, status_static, status_bind)
MODULES = [
    ("Orders & Renban", "Renban worksheet, lot sizing, releases", "material/list_alt",       WARN,  None,            "mod_orders"),
    ("Shipping & ASN",  "Build shipments, EDI 856 notices",       "material/local_shipping",  WARN,  None,            "mod_shipasn"),
    ("Receiving",       "Confirm receipts, status & variances",   "material/move_to_inbox",   None,  None,            None),
    ("Stocktaking",     "Cycle counts, variance approval",        "material/fact_check",      INFO,  None,            "mod_stock"),
    ("Forecast",        "Demand breakdown & planning horizon",    "material/insights",        FAINT, "Updated 06:00", None),
    ("Master Data",     "Suppliers, sizes, logistics, parts",     "material/storage",         MUT,   None,            "mod_master"),
    ("Invoicing & EDI", "EDI 810 invoices, uploads & acks",       "material/receipt_long",    FAINT, "EDI 810",       None),
    ("Reports",         "Inventory, ledger & shipment reports",   "material/bar_chart",       FAINT, "23 saved",      None),
]
def mcard(i, title, desc, ipath, scolor, sstatic, sbind):
    ibox = flex("ibox%d" % i, [icon("ic%d" % i, ipath)], align="center", justify="center",
                style={"backgroundColor": ASOFT, "borderRadius": "10px", "width": "40px", "height": "40px"},
                position={"shrink": 0, "basis": "40px"})
    kids = [ibox,
            lbl("t%d" % i, title, {"fontWeight": "600", "fontSize": "14.5px", "color": TXT}),
            lbl("d%d" % i, desc, {"fontSize": "12.5px", "color": MUT})]
    if sstatic or sbind:
        kids.append(lbl("st%d" % i, sstatic or "—",
                        {"fontSize": "11.5px", "fontWeight": "600" if scolor not in (FAINT, MUT, None) else "500",
                         "color": scolor or FAINT, "margin": "auto 0 0 0"}, bind=sbind))
    return flex("card%d" % i, kids, gap=10, style=dict(CARD, **{"minHeight": "150px"}),
                position={"grow": 1, "basis": "31%"})

def module_section():
    header = flex("modhdr", [
        lbl("mh", "MODULES", {"fontSize": "13px", "fontWeight": "600", "letterSpacing": ".4px", "color": MUT}),
        spacer("mh_sp"),
        lbl("mh_link", "Customize ›", {"color": ACCENT, "fontWeight": "600", "fontSize": "12.5px", "cursor": "pointer"}),
    ], direction="row", align="center", style={"padding": "2px 2px 12px"})
    grid = flex("modgrid", [mcard(i, *m) for i, m in enumerate(MODULES)],
                direction="row", gap=14, wrap="wrap", domId="home-modules")
    return flex("modules", [header, grid], position={"grow": 1, "basis": "0px"})

# ---- attention rail (bound to real counts) ---------------------------------------------
# (color, title_bind, meta_bind)
ATTN = [
    (DANGER, "rail_renban_title",    "rail_renban_meta"),     # P4 RENBAN_COLLISION (top — revenue-path)
    (DANGER, "rail_short_title",     "rail_short_meta"),
    (WARN,   "rail_belowlot_title",  "rail_belowlot_meta"),
    (INFO,   "rail_stocktake_title", "rail_stocktake_meta"),
]
def attn_item(i, color, title_bind, meta_bind):
    dot = flex("ad%d" % i, [], style={"backgroundColor": color, "width": "8px", "height": "8px",
                                      "borderRadius": "50%", "margin": "5px 0 0 0"},
               position={"shrink": 0, "basis": "8px"})
    body = flex("ab%d" % i, [
        lbl("at%d" % i, "—", {"fontSize": "13px", "fontWeight": "500", "color": TXT}, bind=title_bind),
        lbl("am%d" % i, "—", {"fontSize": "11.5px", "color": FAINT, "margin": "2px 0 0 0"}, bind=meta_bind),
    ], position={"grow": 1, "basis": "0px"})
    return flex("attn%d" % i, [dot, body], direction="row", gap=11,
                style={"padding": "13px 0", "borderBottom": "1px solid " + BORDER})

def rail():
    head = flex("rhead", [lbl("rh", "Needs attention", {"fontWeight": "600", "fontSize": "13px", "color": TXT}),
                          spacer("rh_sp"),
                          lbl("rc", str(len(ATTN)), {"fontSize": "11px", "color": MUT})],
                direction="row", align="center", style={"padding": "0 0 6px"})
    items = [attn_item(i, *a) for i, a in enumerate(ATTN)]
    foot = lbl("rf", "View all activity ›", {"color": ACCENT, "fontWeight": "600",
                                             "fontSize": "12.5px", "cursor": "pointer", "margin": "10px 0 0 0"})
    return flex("rail", [head] + items + [foot], style=CARD, position={"shrink": 0, "basis": "320px"})

# ---- assemble --------------------------------------------------------------------------
def build_view():
    body = flex("body", [module_section(), rail()], direction="row", gap=22,
                position={"shrink": 0, "basis": "auto"})
    context = flex("context", [
        lbl("h1", "Home", {"fontSize": "21px", "fontWeight": "700", "color": TXT}),
        lbl("meta", "Thursday, June 18 · day shift · live counts from %s" % DB,
            {"fontSize": "13px", "color": MUT}),
    ], direction="row", gap=12, align="baseline", style={"padding": "6px 2px 18px"})
    footer = flex("footer", [
        lbl("ft", "Inventory · Tupelo, MS (TT) · v0.9-dev · Ignition Perspective",
            {"fontSize": "12px", "color": FAINT}),
    ], direction="row", style={"padding": "24px 0 0", "borderTop": "1px solid " + BORDER, "margin": "26px 0 0"})

    root = flex("root", [topbar(), context, kpi_strip(), body, footer],
                direction="column", gap=22, domId="home-root",
                style={"backgroundColor": BG, "padding": "8px 24px 40px", "minHeight": "100%"})

    return {
        "custom": {"kpi": {k: "" for k in KPI_KEYS}},   # pre-declare so sub-path bindings resolve
        "params": {},
        "propConfig": {"custom.kpi": {"binding": {
            "type": "expr", "config": {"expression": "now(60000)"},   # load on open + refresh 60s
            "transforms": [{"type": "script", "code": KPI_TRANSFORM}]}}},
        "props": {"defaultSize": {"width": 1300, "height": 1000}},
        "root": root,
    }

# ---- writers + patches -----------------------------------------------------------------
def write_view(base, view):
    d = os.path.join(base, "Home", "Home")
    os.makedirs(d, exist_ok=True)
    json.dump(view, open(os.path.join(d, "view.json"), "w"), indent=2)
    json.dump({"scope": "G", "version": 1, "restricted": False, "overridable": True,
               "files": ["view.json"], "attributes": {}}, open(os.path.join(d, "resource.json"), "w"), indent=2)

def patch_page_config():
    p = os.path.join(PERSP, "page-config", "config.json")
    cfg = json.load(open(p))
    cfg.setdefault("pages", {})["/home"] = {"viewPath": "Home/Home", "title": "Home"}
    json.dump(cfg, open(p, "w"), indent=2)

def patch_session_theme(theme="tai-light"):
    p = os.path.join(PERSP, "session-props", "props.json")
    sp = json.load(open(p))
    sp.setdefault("props", {})["theme"] = theme
    json.dump(sp, open(p, "w"), indent=2)

if __name__ == "__main__":
    view = build_view()
    write_view(os.path.join(PERSP, "views"), view)                                      # live gateway
    write_view(os.path.join(REPO_ROOT, "docs", "design", "perspective-views"), view)    # repo copy
    patch_page_config()
    patch_session_theme("tai-light")
    print("OK  Home/Home written (gateway + repo); /home route + theme=tai-light")
    print("    modules=%d  kpis=%d  rail=%d  | KPI load: 1 binding -> custom.kpi (runNamedQuery: kpiSummary+kpiTopShort)" %
          (len(MODULES), len(KPIS), len(ATTN)))
