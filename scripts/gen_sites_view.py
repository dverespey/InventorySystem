#!/usr/bin/env python3
"""Generate the Sites-master combined master-detail Perspective view.json.

EIGHTH master-data module. Mirrors the PROVEN canonical Size view
(Master/Size/Size) field-for-field — single combined view (grid left, detail form
right), inline system.db.runPrepQuery, config.bidirectional:true on every input
binding, RESTRICT refCount delete-gate, 16-char yyyymmddHHMMSSff audit stamp,
same-view recordId prop-write (NO navigation).

DIRECTION REVERSAL (David 2026-06-22): each site runs on its OWN gateway + DB
(single-site deployments, NOT shared-DB multi-tenancy). So INV_SITES holds the ONE
deployment's site config (one row); there is NO site_id surgery. This view is the
config editor for that deployment's site. Practically:
  * The grid shows the site row(s) present (a single-site deployment has 1; the
    spike seed has 2 placeholder rows MAS/HERO so the list+detail mechanics are
    exercised). On a real deployment it is one row.
  * RULE #2 "NOT SITE-SCOPED" is moot under single-site (no session site to scope
    by) — the list simply selects every INV_SITES row with no site filter. There
    is deliberately NO site predicate; do not add one.
  * "support edit + optionally insert-if-missing": the screen edits the existing
    row and can insert one (Save with recordId=0) for the rare bootstrap case.

SITES-MASTER-SPECIFIC (differs from every other master — see the SQL header in
docs/analysis/master-data/master-crud-namedqueries.sql):
  #1 ROLE-GATED  -> the WRITE (Save/Delete) is enforced SERVER-SIDE: each write
                     button calls auth.requireWrite(self.session) FIRST, which resolves
                     the caller's roles FROM THE SESSION (gateway-side) and authorizes
                     ProductionControl|Admin, raising AuthError on deny — BEFORE any
                     system.db write. This closes the reintroduced legacy H3 hole (the
                     write was previously authorized CLIENT-SIDE only via custom.isAdmin).
                     The DETAIL FORM + ACTION BAR ALWAYS RENDER — same as the other 7
                     masters (David 2026-06-22: show it like the other masters). There is
                     NO client-side visibility gate: the old custom.mayEdit prop, the
                     qaAdmin URL hatch, the form/ActionBar meta.visible bindings, and the
                     RESTRICTED AdminBanner were removed. They were UI-only and never the
                     boundary; the server-side requireWrite call is the sole enforcement.
                     (Page-level role-permission in the Designer remains the authoritative
                     UI gate to make the whole /sites page unreachable, distinct from the
                     per-write server enforcement — IG83-TODO.)
  #2 NO SITE FILTER on list/get (moot under single-site; see above).
  #3 read-only system fields shown disabled (IN_SITE_ID, IN_EIN_SEQ,
                     VC_LAST_FORECAST_IMPORT, VC_LAST_UPDATE, VC_ADD).
  #4 typed inputs: BIT_* -> checkbox; VC_FORECAST_IMPORT_MODE -> AUTO/MANUAL
                     dropdown; IN_* -> numeric (CHECKs enforced client-side:
                     fill_days<=50, retention>=12); VC_SEP_* / VC_EDI_MODE ->
                     exactly-1-char (load-bearing positional ISA); DUNS format
                     (9 or 9+4 digits); rest text. Path cols (M4 piece 1) -> text.
  #6 refCount delete-gate counts the throwaway INV_PARTS_STOCK_MST.site_id (the
                     only wired site reference today). Under single-site a site
                     should rarely be deleted; the gate stays as a safety net.

This writes view.json to BOTH the deployed gateway path AND the committed repo path (so the
reviewable/redeployable artifact == runtime), plus the committed repo resource.json (Size shape).
"""
import json
import os

# PROVENANCE: write BOTH copies so the committed (reviewable/redeployable) artifact == runtime.
#   GW_OUT  = the DEPLOYED gateway view (what the running session serves).
#   REPO_OUT = the committed source-of-truth artifact (re-deploying from it must NOT drop anything).
# Before this fix the generator wrote only GW_OUT, so the committed view.json went stale (had none of
# the 7 path columns nor the load-bearing ISA/DUNS validation). Both are now generated from this one file.
GW_OUT = "/usr/local/ignition/data/projects/InventorySystem/com.inductiveautomation.perspective/views/Master/Sites/Sites/view.json"
_REPO = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
REPO_OUT = os.path.join(_REPO, "docs", "analysis", "master-data", "perspective-views",
                        "Master", "Sites", "Sites", "view.json")

# resource.json companion for the committed repo copy. Only scope/version/files are load-bearing on
# disk-load; the gateway fills `attributes` (lastModification + signature) and re-signs on restart, so
# the committed copy deliberately omits them (do NOT chase the gateway-managed signature).
REPO_RESOURCE = os.path.join(os.path.dirname(REPO_OUT), "resource.json")
RESOURCE_JSON = {
    "scope": "G",
    "version": 1,
    "restricted": False,
    "overridable": True,
    "files": ["view.json"],
}

DB = "Inventory_Spike"

# 16-char yyyymmddHHMMSSff audit recipe (byte-identical to the other masters)
AUDIT = ("CONVERT(char(8),GETDATE(),112) "
         "+ SUBSTRING(CONVERT(varchar,GETDATE(),114),1,2) + SUBSTRING(CONVERT(varchar,GETDATE(),114),4,2) "
         "+ SUBSTRING(CONVERT(varchar,GETDATE(),114),7,2) + SUBSTRING(CONVERT(varchar,GETDATE(),114),10,2)")


# ---------------------------------------------------------------------------
# Field model: (custom_key, domId, label, kind, dbcol, default, maxlen, group)
#   kind: text | char1 | num | bit | mode | ro_text | ro_num
#   ro_* = read-only system field (rule #3) shown disabled, never written.
# ---------------------------------------------------------------------------
F = [
    # identity / address
    ("form_id",        "sites-id",        "Site ID (identity)",      "ro_num",  "IN_SITE_ID",   0,  None, "identity"),
    ("form_name",      "sites-name",      "Site Name *",             "text",    "VC_SITE_NAME", "", 50,   "identity"),
    ("form_abbr",      "sites-abbr",      "Site Abbr (ISA) *",       "text",    "VC_SITE_ABBR", "", 10,   "identity"),
    ("form_street",    "sites-street",    "Street",                  "text",    "VC_STREET",    "", 50,   "identity"),
    ("form_city",      "sites-city",      "City",                    "text",    "VC_CITY",      "", 50,   "identity"),
    ("form_state",     "sites-state",     "State",                   "text",    "VC_STATE",     "", 2,    "identity"),
    ("form_country",   "sites-country",   "Country",                 "text",    "VC_COUNTRY",   "", 3,    "identity"),
    ("form_zip",       "sites-zip",       "Zip",                     "text",    "VC_ZIP",       "", 10,   "identity"),
    # EDI / trading
    ("form_duns",      "sites-duns",      "DUNS",                    "text",    "VC_DUNS",      "", 13,   "edi"),
    ("form_supcode",   "sites-supcode",   "Supplier Code",           "text",    "VC_SUPPLIER_CODE","",5, "edi"),
    ("form_dock",      "sites-dock",      "Dock Code",               "text",    "VC_DOCK_CODE", "", 10,   "edi"),
    ("form_einseq",    "sites-einseq",    "EIN Seq (system)",        "ro_num",  "IN_EIN_SEQ",   0,  None, "edi"),
    ("form_edimode",   "sites-edimode",   "EDI Mode",                "text",    "VC_EDI_MODE",  "", 10,   "edi"),
    ("form_sepseg",    "sites-sepseg",    "Sep: Segment",            "char1",   "VC_SEP_SEGMENT","",1,    "edi"),
    ("form_sepelem",   "sites-sepelem",   "Sep: Element",            "char1",   "VC_SEP_ELEMENT","",1,    "edi"),
    ("form_sepsub",    "sites-sepsub",    "Sep: Sub-element",        "char1",   "VC_SEP_SUBELEMENT","",1, "edi"),
    ("form_tmmname",   "sites-tmmname",   "TMM Name",                "text",    "VC_TMM_NAME",  "", 50,   "edi"),
    ("form_tmmabbr",   "sites-tmmabbr",   "TMM Abbr",                "text",    "VC_TMM_ABBR",  "", 10,   "edi"),
    ("form_tmmduns",   "sites-tmmduns",   "TMM DUNS",                "text",    "VC_TMM_DUNS",  "", 13,   "edi"),
    ("form_maxseq",    "sites-maxseq",    "Max Sequence",            "num",     "IN_MAX_SEQUENCE",0,None, "edi"),
    ("form_acceptasn", "sites-acceptasn", "Accept Any Order ASN",    "bit",     "BIT_ACCEPT_ANY_ORDER_ASN",0,None,"edi"),
    ("form_delivery",  "sites-delivery",  "Delivery Method (TD5)",   "text",    "VC_DELIVERY_METHOD_CODE","",5,"edi"),
    # order / forecast
    ("form_filldays",  "sites-filldays",  "Fill Days (<= 50)",       "num",     "IN_FILL_DAYS", 0,  None, "order"),
    ("form_fcusage",   "sites-fcusage",   "Forecast Usage Compare",  "num",     "IN_FORECAST_USAGE_COMPARE",0,None,"order"),
    ("form_usefpd",    "sites-usefpd",    "Use First Production Day", "bit",     "BIT_USE_FIRST_PRODUCTION_DAY",0,None,"order"),
    ("form_fcmode",    "sites-fcmode",    "Forecast Import Mode",     "mode",    "VC_FORECAST_IMPORT_MODE","AUTO",None,"order"),
    ("form_lastfc",    "sites-lastfc",    "Last Forecast Import (system)","ro_text","VC_LAST_FORECAST_IMPORT","",None,"order"),
    # purge
    ("form_enpurge",   "sites-enpurge",   "Enable Data Purge",       "bit",     "BIT_ENABLE_DATA_PURGE",0,None,"purge"),
    ("form_prpurge",   "sites-prpurge",   "Prompt Data Purge",       "bit",     "BIT_PROMPT_DATA_PURGE",1,None,"purge"),
    ("form_retention", "sites-retention", "Data Retention (>= 12 mo)","num",    "IN_DATA_RETENTION",0,None,"purge"),
    # directory paths (M4 piece 1 — relocated from the legacy INI [DIRECTORIES])
    ("form_ediout",    "sites-ediout",    "EDI Out Dir (shared)",    "text",    "VC_EDIOUT_DIR", "", 260, "paths"),
    ("form_ediin",     "sites-ediin",     "EDI In Dir (shared)",     "text",    "VC_EDIIN_DIR",  "", 260, "paths"),
    ("form_forecastd", "sites-forecastd", "Forecast Input Dir",      "text",    "VC_FORECAST_DIR","", 260, "paths"),
    ("form_logisticsd","sites-logisticsd","Logistics Input Dir",     "text",    "VC_LOGISTICS_DIR","",260, "paths"),
    ("form_reportsd",  "sites-reportsd",  "Reports Output Dir",      "text",    "VC_REPORTS_DIR","", 260, "paths"),
    ("form_shippingd", "sites-shippingd", "Shipping File Dir",       "text",    "VC_SHIPPING_DIR","", 260, "paths"),
    ("form_templated", "sites-templated", "Template Dir",            "text",    "VC_TEMPLATE_DIR","", 260, "paths"),
    # audit (read-only)
    ("form_lastupd",   "sites-lastupd",   "Last Update (audit)",     "ro_text", "VC_LAST_UPDATE","",None, "audit"),
    ("form_add",       "sites-add",       "Created (audit)",         "ro_text", "VC_ADD",       "", None, "audit"),
]

GROUPS = [
    ("identity", "Identity / Address"),
    ("edi",      "EDI / Trading Identity"),
    ("order",    "Order / Forecast Config"),
    ("purge",    "Data Purge"),
    ("paths",    "Directory Paths (from legacy INI [DIRECTORIES])"),
    ("audit",    "Audit (read-only)"),
]

# Writable fields in the canonical save order (insert/update column order) — must
# match the Sites/insert + Sites/update SQL param order. Excludes read-only (#3).
WRITABLE = [f for f in F if f[3] not in ("ro_text", "ro_num")]


def custom_defaults():
    # NO `mayEdit` UI-visibility prop: the Sites detail ALWAYS shows, consistent with the other 7 masters
    # (David 2026-06-22 — show it like the other masters). The ONLY authorization boundary is the
    # SERVER-SIDE auth.requireWrite(self.session) call in Save/Delete (resolves SESSION roles gateway-side,
    # authorizes ProductionControl|Admin, raises AuthError on deny). The old client-side visibility gate
    # (mayEdit prop + form/ActionBar meta.visible + RESTRICTED AdminBanner) was removed — it was UI-only,
    # never the enforcement boundary, and made Sites inconsistent with the other masters.
    c = {"recordId": 0, "runNonce": 0, "searchTerm": "", "statusMsg": ""}
    for key, _, _, kind, _, default, _, _ in F:
        c[key] = default
    return c


def value_prop(kind):
    if kind in ("bit",):
        return "props.selected"
    if kind in ("num", "ro_num", "mode"):
        return "props.value"
    return "props.text"  # text, char1, ro_text


def comp_type(kind):
    if kind == "bit":
        return "ia.input.checkbox"
    if kind in ("num", "ro_num"):
        return "ia.input.numeric-entry-field"
    if kind == "mode":
        return "ia.input.dropdown"
    return "ia.input.text-field"


# ---------------------------------------------------------------------------
# LAYOUT (NIT 1 — David click-through 2026-06-24): the ~45-field Sites detail was
# vertically COMPRESSED — every label+input row lived flat in a column flex with
# overflow:auto + maxHeight, but no row carried flex-shrink:0, so when the 45 rows
# overflowed the cap the column SHRANK each row to nothing instead of scrolling
# (unreadable). FIX:
#   * every field ROW gets position.shrink=0 + style.minHeight (ROW_MIN_HEIGHT) so a
#     row is a DEFINITE legible height the flex can never squish; the Form's
#     overflow:auto then SCROLLS instead of compressing.
#   * 2-COLUMN layout: each group's rows live in a wrapping flex (flexWrap:wrap) and
#     each row is ~half width (ROW_BASIS, grow:1) so two fit per line — halving the
#     scroll and improving scan-ability. Group header spans full width (its own line).
#   * section grouping is KEPT (the existing GROUPS) — Identity, EDI, Order/Forecast,
#     Purge, Paths, Audit — each a labelled section.
# The Form keeps overflow:auto and grows to fill the RightPane height (see build_view),
# so the whole detail scrolls as one readable, full-height column.
ROW_MIN_HEIGHT = "34px"   # definite single-line height per row; flex can never shrink below this
ROW_BASIS = "360px"       # ~half a ~760px-wide form line -> two columns; grow:1 fills the remainder
LABEL_BASIS = "168px"     # field label width inside a row (narrower than the old 200 to fit 2-up)


def field_row(key, domid, label, kind, maxlen):
    vprop = value_prop(kind)
    readonly = kind in ("ro_text", "ro_num")
    inner_props = {}
    if kind in ("text", "char1", "ro_text") and maxlen:
        inner_props["maxLength"] = maxlen
    if readonly:
        inner_props["enabled"] = False
    if kind == "mode":
        inner_props["options"] = [
            {"value": "AUTO", "label": "AUTO"},
            {"value": "MANUAL", "label": "MANUAL"},
        ]
    if kind == "bit":
        # blank the checkbox's own caption: ia.input.checkbox defaults props.text to the literal "text",
        # which rendered a stray "text" word beside every bit field (the row LABEL already names it). The
        # value binding lives on props.selected (untouched) — this only clears the cosmetic caption.
        inner_props["text"] = ""
    binding = {
        "config": {"bidirectional": True, "path": "view.custom.%s" % key},
        "type": "property",
    }
    # read-only system fields: one-way (no bidirectional write-back) — rule #3.
    if readonly:
        binding = {"config": {"path": "view.custom.%s" % key}, "type": "property"}
    field = {
        "meta": {"domId": domid, "name": "fld_" + key},
        # the input fills the rest of the (now narrower, 2-up) row after the label
        "position": {"basis": "0px", "grow": 1},
        "propConfig": {vprop: {"binding": binding}},
        "type": comp_type(kind),
    }
    if inner_props:
        field["props"] = inner_props
    lbl = {
        "meta": {"name": "lbl_" + key},
        "position": {"basis": LABEL_BASIS, "shrink": 0},
        "props": {"style": {"fontSize": "12px"}, "text": label},
        "type": "ia.display.label",
    }
    return {
        "meta": {"name": "row_" + key},
        # shrink:0 = the row keeps ROW_MIN_HEIGHT and is NEVER compressed by the column
        # overflow; basis ROW_BASIS + grow:1 = two rows per wrapping line (2 columns).
        "position": {"basis": ROW_BASIS, "grow": 1, "shrink": 0},
        "props": {"alignItems": "center",
                  "style": {"gap": "10px", "minHeight": ROW_MIN_HEIGHT}},
        "type": "ia.container.flex",
        "children": [lbl, field],
    }


def group_block(gkey, gtitle):
    """One labelled section: a full-width header + a 2-column wrapping grid of its rows.
    Returned as a SINGLE section container so the Form is a clean column of sections (the
    header always starts a new section; the rows wrap 2-up within it)."""
    rows = [field_row(key, domid, label, kind, maxlen)
            for (key, domid, label, kind, dbcol, default, maxlen, grp) in F if grp == gkey]
    header = {
        "meta": {"name": "grp_" + gkey},
        "position": {"shrink": 0},
        "props": {"style": {"color": "#1976D2", "fontSize": "12px", "fontWeight": "bold",
                            "marginTop": "4px", "borderBottom": "1px solid #D0D4DA"},
                  "text": gtitle},
        "type": "ia.display.label",
    }
    # the rows wrap 2-up inside this grid; the grid never shrinks (shrink:0) so the
    # Form's overflow:auto scrolls the whole readable column.
    grid = {
        "meta": {"name": "grpgrid_" + gkey},
        "position": {"shrink": 0},
        "props": {"wrap": "wrap", "style": {"gap": "8px 16px"}},
        "type": "ia.container.flex",
        "children": rows,
    }
    return [{
        "meta": {"name": "section_" + gkey},
        "position": {"shrink": 0},
        "props": {"direction": "column", "style": {"gap": "4px"}},
        "type": "ia.container.flex",
        "children": [header, grid],
    }]


# ---------------------------------------------------------------------------
# Scripts (Jython 2.7; IG81-COMPAT).
# ---------------------------------------------------------------------------

GRID_SCRIPT = r'''	# ---- Master/Sites combined grid builder — inline runPrepQuery ----
	# RULE #2: NOT SITE-SCOPED. Admin manages ALL sites -> NO site filter, every row.
	# IG81-COMPAT: plain parameterized T-SQL; identical on 8.1.52 and 8.3.
	log = system.util.getLogger("SPIKE")
	DB = "Inventory_Spike"
	parts = unicode(value).split("|")
	searchTerm = parts[0]

	sql = ("SELECT  IN_SITE_ID              AS RecordID, "
	       "        VC_SITE_ABBR            AS Abbr, "
	       "        VC_SITE_NAME            AS SiteName, "
	       "        VC_STATE                AS State, "
	       "        VC_DUNS                 AS DUNS, "
	       "        VC_FORECAST_IMPORT_MODE AS FCMode "
	       "FROM    INV_SITES "
	       "WHERE  (? = '' OR VC_SITE_ABBR LIKE '%' + ? + '%' OR VC_SITE_NAME LIKE '%' + ? + '%' OR VC_DUNS LIKE '%' + ? + '%') "
	       "ORDER BY VC_SITE_ABBR")   # RULE #2: deliberately NO site_id predicate
	ds = system.db.runPrepQuery(sql, [searchTerm, searchTerm, searchTerm, searchTerm], DB)
	log.info("SPIKE Sites/list: %d rows (search='%s')" % (ds.rowCount, searchTerm))

	columns = [
		{"field": "Abbr",     "header": {"title": "Abbr"}, "strictWidth": 90, "sort": "ascending"},
		{"field": "SiteName", "header": {"title": "Site Name"}},
		{"field": "State",    "header": {"title": "State"}, "strictWidth": 70},
		{"field": "DUNS",     "header": {"title": "DUNS"}, "strictWidth": 120},
		{"field": "FCMode",   "header": {"title": "FC Mode"}, "strictWidth": 90},
		{"field": "RecordID", "header": {"title": "id"}, "visible": False}
	]
	rows = []
	for r in xrange(ds.rowCount):
		rows.append({
			"RecordID": ds.getValueAt(r, "RecordID"),
			"Abbr":     ds.getValueAt(r, "Abbr"),
			"SiteName": ds.getValueAt(r, "SiteName"),
			"State":    ds.getValueAt(r, "State"),
			"DUNS":     ds.getValueAt(r, "DUNS"),
			"FCMode":   ds.getValueAt(r, "FCMode")
		})
	return {"data": rows, "columns": columns}
'''


def _get_assign(key, dbcol, kind):
    if kind == "bit":
        return '\tc.%s = 1 if (g("%s") in (1, True, "1")) else 0' % (key, dbcol)
    if kind in ("num", "ro_num"):
        return ('\t_v = ds.getValueAt(0, "%s")\n\tc.%s = int(_v) if _v is not None else 0' % (dbcol, key))
    return '\tc.%s = g("%s")' % (key, dbcol)


def recordid_onchange():
    cols = ", ".join(dbcol for (_, _, _, _, dbcol, _, _, _) in F)
    assigns = "\n".join(_get_assign(key, dbcol, kind)
                        for (key, _, _, kind, dbcol, _, _, _) in F)
    return (
        "\t# ---- Sites combined view: recordId change -> load detail ----\n"
        "\t# Grid onRowClick sets self.view.custom.recordId (same-view prop write, NO\n"
        "\t# navigation). recordId==0 -> insert mode. RULE #2: get is NOT site-scoped.\n"
        "\t# IG81-COMPAT: plain change script + T-SQL.\n"
        "\tlog = system.util.getLogger(\"SPIKE\")\n"
        "\tv = self.view if hasattr(self, \"view\") else self\n"
        "\tDB = \"Inventory_Spike\"\n"
        "\tc = v.custom\n"
        "\tcv = getattr(currentValue, \"value\", currentValue)\n"
        "\trecId = int(cv or 0)\n"
        "\tlog.info(\"SPIKE Sites Detail recordId onChange fired: recId=%s origin=%s\" % (recId, origin))\n"
        "\tif recId <= 0:\n"
        "\t\tc.statusMsg = \"New site — enter details and Save.\"\n"
        "\t\tlog.info(\"SPIKE Sites Detail: insert mode (recordId=0)\")\n"
        "\t\treturn\n"
        "\tgetSql = (\"SELECT " + cols + " FROM INV_SITES WHERE IN_SITE_ID = ?\")  # RULE #2: no site predicate\n"
        "\tds = system.db.runPrepQuery(getSql, [recId], DB)\n"
        "\tif ds.rowCount == 0:\n"
        "\t\tc.statusMsg = \"Site id %d not found.\" % recId\n"
        "\t\tlog.warn(\"SPIKE Sites Detail: id %d not found\" % recId); return\n"
        "\tdef g(col):\n"
        "\t\tval = ds.getValueAt(0, col)\n"
        "\t\treturn val if val is not None else \"\"\n"
        + assigns + "\n"
        "\tc.statusMsg = \"Editing site %s (id %d).\" % (c.form_abbr, recId)\n"
        "\tlog.info(\"SPIKE Sites Detail loaded id=%d abbr=%s name=%s\" % (recId, c.form_abbr, c.form_name))\n"
    )


def field_clears(indent):
    """Reset-to-default assignments for every WRITABLE form field (read-only system
    fields rule #3 are skipped). Single source for the 3 places that clear the form:
    new_script(), delete_script() (2-tab, inside try), and the Clear button (1-tab).
    `indent` is the leading whitespace (e.g. "\\t" or "\\t\\t")."""
    return "\n".join(
        "%sc.%s = %s" % (indent, key,
                         ('""' if kind in ("text", "char1", "mode") and default == "" else repr(default)))
        for (key, _, _, kind, _, default, _, _) in F if kind not in ("ro_text", "ro_num"))


def new_script():
    clears = field_clears("\t")
    # NIT (P16): New is a NO-DB-WRITE form reset — IDENTICAL in effect to Clear (blank the form +
    # recordId=0). It is NOT a write, so it carries NO server-side write gate (gate the WRITES — Save/Delete
    # — not the resets). Gating New here was inconsistent with the ungated Clear and would deny a read-only
    # viewer from even starting a new entry; the actual security boundary is auth.requireWrite on Save/Delete,
    # which is unchanged. Decision documented in m4-auth-design.md / the P16 deliverable note: New + Clear are
    # gated CONSISTENTLY = NEITHER (both no-DB-write resets); the Save/Delete writes remain server-side gated.
    return (
        "\t# New record: same-view prop write (NO navigation, NO DB write). recordId=0 -> insert mode.\n"
        "\t# No write gate: this is a form reset (identical to Clear), not a DB write — the WRITES\n"
        "\t# (Save/Delete) carry the server-side auth.requireWrite gate. (P16 Clear/New consistency NIT.)\n"
        "\tlog = system.util.getLogger(\"SPIKE\")\n"
        "\tc = self.view.custom\n"
        + clears + "\n"
        "\tc.recordId = 0\n"
        "\tc.statusMsg = \"New site — enter details and Save.\"\n"
        "\tlog.info(\"SPIKE Sites New -> recordId=0 (insert mode)\")\n"
    )


# ---------------------------------------------------------------------------
# SHARED SQL builders (module-level, pure) — the SINGLE source of truth for the
# insert/update/get column lists. Both save_script() (the view) AND the e2e test
# (test_sites_master.py) build SQL from these, so the test exercises the IDENTICAL
# SQL the view runs (no drift). The AUDIT expr is appended by the view-side script
# (the test stamps VC_ADD/VC_LAST_UPDATE explicitly with the same recipe).
# ---------------------------------------------------------------------------

# Nullable "unset = NULL" INT columns (schema: all NULL-able; CHECKs allow NULL;
# the table treats NULL as legacy "unset"). Blank/0 binds NULL, never literal 0 —
# 0 is rejected by CK_INV_SITES_DATA_RETENTION (>=12) and is not the "unset" marker
# for the others (B1/N1 fix).
NULLABLE_UNSET = {"form_retention", "form_filldays", "form_fcusage", "form_maxseq"}

# ordered (form_key, kind) args for insert (skip IN_EIN_SEQ -> literal 0; VC_ADD in SQL).
INS_ORDER = [
    ("form_name", "text"), ("form_abbr", "text"), ("form_street", "text"), ("form_city", "text"),
    ("form_state", "text"), ("form_country", "text"), ("form_zip", "text"),
    ("form_duns", "text"), ("form_supcode", "text"), ("form_dock", "text"),
    ("form_edimode", "text"),
    ("form_sepseg", "text"), ("form_sepelem", "text"), ("form_sepsub", "text"),
    ("form_tmmname", "text"), ("form_tmmabbr", "text"), ("form_tmmduns", "text"), ("form_maxseq", "num"),
    ("form_acceptasn", "bit"), ("form_delivery", "text"), ("form_filldays", "num"),
    ("form_fcusage", "num"), ("form_usefpd", "bit"), ("form_fcmode", "mode"),
    ("form_enpurge", "bit"), ("form_prpurge", "bit"), ("form_retention", "num"),
    # directory paths (M4 piece 1) — same order in insert + update column lists.
    ("form_ediout", "text"), ("form_ediin", "text"), ("form_forecastd", "text"),
    ("form_logisticsd", "text"), ("form_reportsd", "text"), ("form_shippingd", "text"),
    ("form_templated", "text"),
]

INS_COLS = ("VC_SITE_NAME, VC_SITE_ABBR, VC_STREET, VC_CITY, VC_STATE, VC_COUNTRY, VC_ZIP, "
            "VC_DUNS, VC_SUPPLIER_CODE, VC_DOCK_CODE, IN_EIN_SEQ, VC_EDI_MODE, "
            "VC_SEP_SEGMENT, VC_SEP_ELEMENT, VC_SEP_SUBELEMENT, "
            "VC_TMM_NAME, VC_TMM_ABBR, VC_TMM_DUNS, IN_MAX_SEQUENCE, "
            "BIT_ACCEPT_ANY_ORDER_ASN, VC_DELIVERY_METHOD_CODE, IN_FILL_DAYS, "
            "IN_FORECAST_USAGE_COMPARE, BIT_USE_FIRST_PRODUCTION_DAY, VC_FORECAST_IMPORT_MODE, "
            "BIT_ENABLE_DATA_PURGE, BIT_PROMPT_DATA_PURGE, IN_DATA_RETENTION, "
            "VC_EDIOUT_DIR, VC_EDIIN_DIR, VC_FORECAST_DIR, VC_LOGISTICS_DIR, "
            "VC_REPORTS_DIR, VC_SHIPPING_DIR, VC_TEMPLATE_DIR, VC_ADD")
# VALUES placeholders: IN_EIN_SEQ is the literal 0 (rule #3); the 7 path cols are ?; VC_ADD is the audit expr.
INS_VALS = "?,?,?,?,?,?,?, ?,?,?, 0, ?, ?,?,?, ?,?,?,?, ?,?, ?, ?,?,?, ?,?,?, ?,?,?,?, ?,?,?, "

UPD_SET = ("VC_SITE_NAME=?, VC_SITE_ABBR=?, VC_STREET=?, VC_CITY=?, VC_STATE=?, VC_COUNTRY=?, VC_ZIP=?, "
           "VC_DUNS=?, VC_SUPPLIER_CODE=?, VC_DOCK_CODE=?, VC_EDI_MODE=?, "
           "VC_SEP_SEGMENT=?, VC_SEP_ELEMENT=?, VC_SEP_SUBELEMENT=?, "
           "VC_TMM_NAME=?, VC_TMM_ABBR=?, VC_TMM_DUNS=?, IN_MAX_SEQUENCE=?, "
           "BIT_ACCEPT_ANY_ORDER_ASN=?, VC_DELIVERY_METHOD_CODE=?, IN_FILL_DAYS=?, "
           "IN_FORECAST_USAGE_COMPARE=?, BIT_USE_FIRST_PRODUCTION_DAY=?, VC_FORECAST_IMPORT_MODE=?, "
           "BIT_ENABLE_DATA_PURGE=?, BIT_PROMPT_DATA_PURGE=?, IN_DATA_RETENTION=?, "
           "VC_EDIOUT_DIR=?, VC_EDIIN_DIR=?, VC_FORECAST_DIR=?, VC_LOGISTICS_DIR=?, "
           "VC_REPORTS_DIR=?, VC_SHIPPING_DIR=?, VC_TEMPLATE_DIR=?, ")

# the full ordered column list for Sites/get (every column incl. read-only + paths) — F order.
GET_COLS = ", ".join(dbcol for (_, _, _, _, dbcol, _, _, _) in F)


def save_script():
    def pyval(key, kind):
        if kind == "bit":
            return "_bit(c.%s)" % key
        if kind == "num":
            if key in NULLABLE_UNSET:
                return "_intn(c.%s)" % key
            return "_int(c.%s)" % key
        if kind == "mode":
            return "_mode(c.%s)" % key
        return "_str(c.%s)" % key

    ins_cols = INS_COLS
    ins_vals = INS_VALS + AUDIT
    ins_order = INS_ORDER
    ins_args = "[" + ", ".join(pyval(k, kd) for k, kd in ins_order) + "]"
    upd_set = UPD_SET + "VC_LAST_UPDATE = " + AUDIT
    upd_args = "[" + ", ".join(pyval(k, kd) for k, kd in ins_order) + ", recId]"

    return (
        "\t# ---- Sites Save (insert-or-update) ----\n"
        "\t# recordId==0 -> insert (IDENTITY assigns id, VC_ADD stamped); >0 -> update\n"
        "\t# (VC_LAST_UPDATE stamped). RULE #3: IN_SITE_ID/IN_EIN_SEQ/VC_LAST_FORECAST_IMPORT\n"
        "\t# are NEVER written by the form. RULE #2: no site predicate.\n"
        "\timport auth as A\n"
        "\tlog = system.util.getLogger(\"SPIKE\")\n"
        "\tDB = \"Inventory_Spike\"\n"
        "\tc = self.view.custom\n"
        "\trecId = int(c.recordId or 0)\n"
        "\t# ---- SERVER-SIDE WRITE GATE (rule #1, the H3 hole-closer) ----\n"
        "\t# auth.requireWrite resolves the SESSION's roles GATEWAY-SIDE (not the client mayEdit prop)\n"
        "\t# and authorizes ProductionControl|Admin, raising AuthError on deny — BEFORE any system.db\n"
        "\t# write. A forged client prop (mayEdit=true via devtools) or an anon session is rejected HERE.\n"
        "\t# The c.mayEdit / form-hide is UI defense-in-depth only; this call is the enforcement boundary.\n"
        "\ttry:\n"
        "\t\tA.requireWrite(self.session)\n"
        "\texcept A.AuthError, e:\n"
        "\t\tc.statusMsg = \"DENIED (server-side): %s\" % unicode(e)\n"
        "\t\tlog.warn(\"SPIKE Sites save DENIED (server-side gate): %s\" % unicode(e)); return\n"
        "\tdef _str(v):\n"
        "\t\ts = (unicode(v) if v is not None else u\"\").strip()\n"
        "\t\treturn s\n"
        "\tdef _int(v):\n"
        "\t\ttry:\n"
        "\t\t\treturn int(v) if v not in (None, \"\") else 0\n"
        "\t\texcept Exception:\n"
        "\t\t\treturn 0\n"
        "\tdef _intn(v):\n"
        "\t\t# nullable \"unset\" INT: blank/0 -> None (binds SQL NULL), else the typed int.\n"
        "\t\t# 0 is NOT a valid stored value for these cols (CK rejects retention 0; the\n"
        "\t\t# others use NULL as the legacy \"unset\" marker) -> coerce blank/0 to NULL.\n"
        "\t\tn = _int(v)\n"
        "\t\treturn n if n != 0 else None\n"
        "\tdef _bit(v):\n"
        "\t\treturn 1 if v in (1, True, \"1\", \"true\", \"True\") else 0\n"
        "\tdef _mode(v):\n"
        "\t\tm = (_str(v) or \"AUTO\").upper()\n"
        "\t\treturn m if m in (\"AUTO\", \"MANUAL\") else \"AUTO\"\n"
        "\tname = _str(c.form_name)\n"
        "\tabbr = _str(c.form_abbr)\n"
        "\t# ---- validation (client-side fast-fail; CHECK constraints are the backstop) ----\n"
        "\tif name == \"\":\n"
        "\t\tc.statusMsg = \"Site name is required.\"\n"
        "\t\tlog.info(\"SPIKE Sites save REJECTED: blank name\"); return\n"
        "\tif abbr == \"\":\n"
        "\t\tc.statusMsg = \"Site abbr is required.\"\n"
        "\t\tlog.info(\"SPIKE Sites save REJECTED: blank abbr\"); return\n"
        "\tif len(abbr) > 10:\n"
        "\t\tc.statusMsg = \"Site abbr must be 10 characters or fewer.\"\n"
        "\t\tlog.info(\"SPIKE Sites save REJECTED: abbr len=%d\" % len(abbr)); return\n"
        "\tfillDays = _int(c.form_filldays)\n"
        "\tif fillDays > 50:\n"
        "\t\tc.statusMsg = \"Fill Days must be 50 or fewer (CK_INV_SITES_FILL_DAYS).\"\n"
        "\t\tlog.info(\"SPIKE Sites save REJECTED: fillDays=%d > 50\" % fillDays); return\n"
        "\t# Retention: blank/0 = \"unset\" -> stored as NULL (allowed by the CHECK); a\n"
        "\t# TYPED value 1..11 is invalid and must be rejected with the clear message.\n"
        "\tretention = _int(c.form_retention)\n"
        "\tif retention != 0 and retention < 12:\n"
        "\t\tc.statusMsg = \"Data Retention must be >= 12 months, or left blank/0 for 'unset' (CK_INV_SITES_DATA_RETENTION).\"\n"
        "\t\tlog.info(\"SPIKE Sites save REJECTED: retention=%d (typed 1..11 invalid)\" % retention); return\n"
        "\tmode = _mode(c.form_fcmode)\n"
        "\tif mode not in (\"AUTO\", \"MANUAL\"):\n"
        "\t\tc.statusMsg = \"Forecast Import Mode must be AUTO or MANUAL.\"\n"
        "\t\tlog.info(\"SPIKE Sites save REJECTED: bad mode=%s\" % mode); return\n"
        "\t# ---- LOAD-BEARING positional-ISA validation (source-truth §4) ----\n"
        "\t# VC_EDI_MODE -> ISA15 verbatim; the 3 separators are ISA/GS structural\n"
        "\t# chars. A 2+ char value emits a MALFORMED positional ISA (EDI856/810Object).\n"
        "\t# Each must be EXACTLY 1 char if set. Blank is allowed (col is NULLable).\n"
        "\tedimode = _str(c.form_edimode)\n"
        "\tif edimode != \"\" and len(edimode) != 1:\n"
        "\t\tc.statusMsg = \"EDI Mode must be exactly 1 character (positional ISA15, e.g. 'P' or 'T').\"\n"
        "\t\tlog.info(\"SPIKE Sites save REJECTED: edimode len=%d\" % len(edimode)); return\n"
        "\tfor _lbl, _v in ((\"Segment\", _str(c.form_sepseg)), (\"Element\", _str(c.form_sepelem)), (\"Sub-element\", _str(c.form_sepsub))):\n"
        "\t\tif _v != \"\" and len(_v) != 1:\n"
        "\t\t\tc.statusMsg = \"Separator (%s) must be exactly 1 character (ISA/GS structural).\" % _lbl\n"
        "\t\t\tlog.info(\"SPIKE Sites save REJECTED: sep %s len=%d\" % (_lbl, len(_v))); return\n"
        "\t# DUNS format: 9 digits (or 9+4 = 13 digits), if set. (Real DUNS load at cutover;\n"
        "\t# the spike seed uses 9-digit placeholders.)\n"
        "\tdef _duns_ok(d):\n"
        "\t\treturn d == \"\" or (d.isdigit() and len(d) in (9, 13))\n"
        "\tduns = _str(c.form_duns)\n"
        "\tif not _duns_ok(duns):\n"
        "\t\tc.statusMsg = \"DUNS must be 9 digits (or 9+4 = 13 digits).\"\n"
        "\t\tlog.info(\"SPIKE Sites save REJECTED: bad DUNS=%r\" % duns); return\n"
        "\ttmmduns = _str(c.form_tmmduns)\n"
        "\tif not _duns_ok(tmmduns):\n"
        "\t\tc.statusMsg = \"TMM DUNS must be 9 digits (or 9+4 = 13 digits).\"\n"
        "\t\tlog.info(\"SPIKE Sites save REJECTED: bad TMM DUNS=%r\" % tmmduns); return\n"
        "\ttry:\n"
        "\t\tif recId == 0:\n"
        "\t\t\tinsSql = (\"INSERT INTO INV_SITES (" + ins_cols + ") \"\n"
        "\t\t\t          \"VALUES (" + ins_vals + "); \"\n"
        "\t\t\t          \"SELECT CAST(SCOPE_IDENTITY() AS int) AS newId\")\n"
        "\t\t\trs = system.db.runPrepQuery(insSql, " + ins_args + ", DB)\n"
        "\t\t\tnewId = int(rs.getValueAt(0, 0))\n"
        "\t\t\tc.statusMsg = \"Inserted site %s (id %d).\" % (abbr, newId)\n"
        "\t\t\tlog.info(\"SPIKE Sites INSERT ok: abbr=%s newId=%d\" % (abbr, newId))\n"
        "\t\t\tc.runNonce = (c.runNonce or 0) + 1\n"
        "\t\t\tc.recordId = newId\n"
        "\t\telse:\n"
        "\t\t\tupdSql = (\"UPDATE INV_SITES SET " + upd_set + " \"\n"
        "\t\t\t          \"WHERE IN_SITE_ID = ?\")  # RULE #2: no site predicate\n"
        "\t\t\tn = system.db.runPrepUpdate(updSql, " + upd_args + ", DB)\n"
        "\t\t\tc.statusMsg = \"Updated site %s (%d row).\" % (abbr, n)\n"
        "\t\t\tlog.info(\"SPIKE Sites UPDATE ok: id=%d abbr=%s rows=%d\" % (recId, abbr, n))\n"
        "\t\t\tc.runNonce = (c.runNonce or 0) + 1\n"
        "\texcept Exception, e:\n"
        "\t\tc.statusMsg = \"Save failed: %s\" % unicode(e)\n"
        "\t\tlog.warn(\"SPIKE Sites save FAILED: %s\" % unicode(e))\n"
    )


def delete_script():
    # NOTE: these clears are emitted INSIDE the try: block -> TWO tabs of indent.
    clears = field_clears("\t\t")
    return (
        "\t# ---- Sites Delete — RESTRICT gate (rule #6) ----\n"
        "\t# refCount counts the THROWAWAY INV_PARTS_STOCK_MST.site_id (no FK yet).\n"
        "\t# ⚠️ MUST be EXTENDED to every IN_SITE_ID child once M4 wires the FKs.\n"
        "\t# Block on any non-zero total; never delete a referenced site.\n"
        "\timport auth as A\n"
        "\tlog = system.util.getLogger(\"SPIKE\")\n"
        "\tDB = \"Inventory_Spike\"\n"
        "\tc = self.view.custom\n"
        "\trecId = int(c.recordId or 0)\n"
        "\tabbr = (unicode(c.form_abbr) if c.form_abbr is not None else u\"\").strip()\n"
        "\t# ---- SERVER-SIDE WRITE GATE (rule #1, the H3 hole-closer) ----\n"
        "\t# Roles from the SESSION (gateway-side), NOT the client mayEdit prop. Rejected before the\n"
        "\t# refCount read + the DELETE. A forged prop / anon session cannot delete a site.\n"
        "\ttry:\n"
        "\t\tA.requireWrite(self.session)\n"
        "\texcept A.AuthError, e:\n"
        "\t\tc.statusMsg = \"DENIED (server-side): %s\" % unicode(e)\n"
        "\t\tlog.warn(\"SPIKE Sites DELETE DENIED (server-side gate): %s\" % unicode(e)); return\n"
        "\tif recId == 0:\n"
        "\t\tc.statusMsg = \"Nothing to delete (unsaved new record).\"\n"
        "\t\treturn\n"
        "\trefSql = \"SELECT (SELECT COUNT(*) FROM INV_PARTS_STOCK_MST WHERE site_id = ?) AS n\"\n"
        "\trs = system.db.runPrepQuery(refSql, [recId], DB)\n"
        "\tn = int(rs.getValueAt(0, 0))\n"
        "\tlog.info(\"SPIKE Sites refCount: id=%d abbr=%s n=%d\" % (recId, abbr, n))\n"
        "\tif n > 0:\n"
        "\t\tc.statusMsg = (\"Cannot delete site %s — still referenced by %d part(s). \"\n"
        "\t\t               \"Reassign or archive those rows first.\") % (abbr, n)\n"
        "\t\tlog.info(\"SPIKE Sites DELETE BLOCKED: id=%d abbr=%s refs=%d\" % (recId, abbr, n)); return\n"
        "\ttry:\n"
        "\t\tdelN = system.db.runPrepUpdate(\"DELETE FROM INV_SITES WHERE IN_SITE_ID = ?\", [recId], DB)\n"
        "\t\tlog.info(\"SPIKE Sites DELETE ok: id=%d abbr=%s rows=%d\" % (recId, abbr, delN))\n"
        "\t\tc.statusMsg = \"Deleted site %s.\" % abbr\n"
        + clears + "\n"
        "\t\tc.runNonce = (c.runNonce or 0) + 1\n"
        "\t\tc.recordId = 0\n"
        "\texcept Exception, e:\n"
        "\t\tc.statusMsg = \"Delete failed: %s\" % unicode(e)\n"
        "\t\tlog.warn(\"SPIKE Sites DELETE FAILED: %s\" % unicode(e))\n"
    )


def build_view():
    children_form = []
    for gkey, gtitle in GROUPS:
        children_form.extend(group_block(gkey, gtitle))

    view = {
        "custom": custom_defaults(),
        "params": {},
        "propConfig": {
            "custom.gridModel": {
                "binding": {
                    "config": {"expression": "{view.custom.searchTerm} + '|' + {view.custom.runNonce}"},
                    "transforms": [{"code": GRID_SCRIPT, "type": "script"}],
                    "type": "expr",
                }
            },
            # NO UI-VISIBILITY GATE. The Sites detail (Form + ActionBar) ALWAYS renders, consistent with
            # the other 7 masters (David 2026-06-22). The ONLY authorization boundary is the SERVER-SIDE
            # auth.requireWrite(self.session) call in Save/Delete (resolves SESSION roles gateway-side,
            # authorizes ProductionControl|Admin, raises AuthError on deny BEFORE any system.db write). A
            # forged client prop or an anon session is rejected THERE — the visibility gate was never the
            # boundary, so removing it (the old custom.mayEdit prop + qaAdmin URL hatch + form/ActionBar
            # meta.visible + RESTRICTED AdminBanner) does not weaken security. (IG83-TODO: a Designer
            # page-level role-permission can make the whole /sites page unreachable to non-ProductionControl
            # users; that is the authoritative UI gate, distinct from the per-write server enforcement.)
            "custom.recordId": {
                "onChange": {"enabled": None, "script": recordid_onchange()}
            },
        },
        "props": {"defaultSize": {"height": 900, "width": 1440}},
        "root": {
            "meta": {"name": "root"},
            "props": {"style": {"backgroundColor": "#F5F6F8", "gap": "12px", "padding": "12px"}},
            "type": "ia.container.flex",
            "children": [
                # ---- LEFT PANE: list ----
                {
                    "meta": {"name": "LeftPane"},
                    "position": {"basis": "560px", "grow": 1},
                    "props": {"direction": "column", "style": {"gap": "10px"}},
                    "type": "ia.container.flex",
                    "children": [
                        {
                            "meta": {"name": "Title"},
                            "props": {"style": {"color": "#222222", "fontSize": "15px", "fontWeight": "bold"},
                                      "text": "Sites Master — List + Detail  ·  INV_SITES (Inventory_Spike)  ·  ProductionControl (site config)  ·  single-site deployment"},
                            "type": "ia.display.label",
                        },
                        # (no AdminBanner — the RESTRICTED banner was part of the removed UI-visibility gate)
                        # filter bar
                        {
                            "meta": {"name": "FilterBar"},
                            "position": {"shrink": 0},
                            "props": {"alignItems": "flex-end",
                                      "style": {"backgroundColor": "#FFFFFF", "border": "1px solid #D0D4DA",
                                                "borderRadius": "4px", "gap": "12px", "padding": "10px"}},
                            "type": "ia.container.flex",
                            "children": [
                                {
                                    "meta": {"name": "SearchWrap"},
                                    "props": {"direction": "column", "style": {"gap": "2px"}},
                                    "type": "ia.container.flex",
                                    "children": [
                                        {"meta": {"name": "SearchCap"},
                                         "props": {"style": {"color": "#666666", "fontSize": "11px"},
                                                   "text": "Search (abbr, name or DUNS)"},
                                         "type": "ia.display.label"},
                                        {"meta": {"domId": "sites-search", "name": "SearchField"},
                                         "position": {"basis": "240px"},
                                         "propConfig": {"props.text": {"binding": {"config": {
                                             "bidirectional": True, "path": "view.custom.searchTerm"},
                                             "type": "property"}}},
                                         "type": "ia.input.text-field"},
                                    ],
                                },
                                {"events": {"component": {"onActionPerformed": {"config": {
                                    "script": "\tself.view.custom.runNonce = (self.view.custom.runNonce or 0) + 1\n\tsystem.util.getLogger(\"SPIKE\").info(\"Sites search -> runNonce=\" + str(self.view.custom.runNonce))\n"},
                                    "scope": "G", "type": "script"}}},
                                 "meta": {"domId": "sites-search-btn", "name": "SearchButton"},
                                 "props": {"style": {"backgroundColor": "#1976D2", "color": "#FFFFFF", "fontWeight": "bold"},
                                           "text": "Search"},
                                 "type": "ia.input.button"},
                                {"events": {"component": {"onActionPerformed": {"config": {
                                    "script": new_script()}, "scope": "G", "type": "script"}}},
                                 "meta": {"domId": "sites-new-btn", "name": "NewButton"},
                                 "props": {"style": {"backgroundColor": "#388E3C", "color": "#FFFFFF", "fontWeight": "bold"},
                                           "text": "New Site"},
                                 "type": "ia.input.button"},
                            ],
                        },
                        # grid
                        {
                            "events": {"component": {"onRowClick": {"config": {
                                "script": "\t# Row select -> SAME-VIEW prop write (NO navigation).\n\tlog = system.util.getLogger(\"SPIKE\")\n\ttry:\n\t\trecId = event.value[\"RecordID\"]\n\texcept Exception:\n\t\trecId = event.value[\"RecordID\"] if hasattr(event, \"value\") else None\n\tif recId is None:\n\t\tlog.warn(\"SPIKE Sites list row click: no RecordID in event\")\n\t\treturn\n\tlog.info(\"SPIKE Sites list -> open Detail recordId=%s\" % recId)\n\tself.view.custom.recordId = int(recId)\n"},
                                "scope": "G", "type": "script"}}},
                            "meta": {"domId": "sites-grid", "name": "SitesGrid"},
                            "position": {"basis": "500px", "grow": 1},
                            "propConfig": {
                                "props.columns": {"binding": {"config": {"path": "view.custom.gridModel.columns"}, "type": "property"}},
                                "props.data": {"binding": {"config": {"path": "view.custom.gridModel.data"}, "type": "property"}},
                            },
                            "props": {
                                "filtering": {"enabled": False},
                                "pager": {"activeOption": 1000, "bottom": False, "options": [1000], "top": False},
                                "selection": {"enabled": True},
                                "sortOrder": ["0:Abbr"],
                            },
                            "type": "ia.display.table",
                        },
                        {"meta": {"name": "FooterNote"},
                         "props": {"style": {"color": "#37474F", "fontSize": "11px", "fontStyle": "italic"},
                                   "text": "ProductionControl-editable (site config). Single-site deployment (one gateway = the site). Single combined view; row-select is an in-view prop write, no navigation. IN_SITE_ID / IN_EIN_SEQ / last-forecast-import / audit fields are read-only (system-maintained)."},
                         "type": "ia.display.label"},
                    ],
                },
                # ---- RIGHT PANE: detail ----
                {
                    "meta": {"name": "RightPane"},
                    "position": {"basis": "640px", "grow": 1},
                    # minHeight:0 lets the column-flex Form child scroll instead of forcing the pane taller
                    # than the page (the NIT-1 scroll fix needs an established height bound on this pane).
                    "props": {"direction": "column", "style": {"gap": "8px", "minHeight": "0px"}},
                    "type": "ia.container.flex",
                    "children": [
                        {"meta": {"name": "DetailTitle"},
                         "props": {"style": {"color": "#222222", "fontSize": "15px", "fontWeight": "bold"}, "text": "Detail"},
                         "type": "ia.display.label"},
                        {"meta": {"domId": "sites-status", "name": "StatusMsg"},
                         "propConfig": {"props.text": {"binding": {"config": {"path": "view.custom.statusMsg"}, "type": "property"}}},
                         "props": {"style": {"color": "#B71C1C", "fontSize": "13px", "fontWeight": "bold", "padding": "6px 8px"}},
                         "type": "ia.display.label"},
                        # form (ALWAYS visible — like the other masters; the write is gated server-side, rule #1)
                        # NIT 1: the Form GROWS to fill the RightPane height between StatusMsg and the
                        # ActionBar (grow:1, basis:0) and SCROLLS its 45 readable rows (overflow:auto). The
                        # minHeight:0 is LOAD-BEARING — a column-flex child won't shrink below its content
                        # (so overflow:auto won't scroll) WITHOUT it; with it the Form is bounded by the pane
                        # and scrolls. No fixed maxHeight cap any more (the pane bounds the height), so the
                        # detail uses the full available height instead of being squished into 640px.
                        {
                            "meta": {"domId": "sites-form", "name": "Form"},
                            "position": {"basis": "0px", "grow": 1, "shrink": 1},
                            "props": {"direction": "column",
                                      "style": {"backgroundColor": "#FFFFFF", "border": "1px solid #D0D4DA",
                                                "borderRadius": "4px", "gap": "8px", "padding": "12px",
                                                "overflow": "auto", "minHeight": "0px"}},
                            "type": "ia.container.flex",
                            "children": children_form,
                        },
                        # action bar (ALWAYS visible — like the other masters; Save/Delete are gated server-side)
                        {
                            "meta": {"name": "ActionBar"},
                            "position": {"shrink": 0},
                            "props": {"alignItems": "center", "style": {"gap": "12px", "padding": "8px 0"}},
                            "type": "ia.container.flex",
                            "children": [
                                {"events": {"component": {"onActionPerformed": {"config": {"script": save_script()},
                                    "scope": "G", "type": "script"}}},
                                 "meta": {"domId": "sites-save-btn", "name": "SaveButton"},
                                 "props": {"style": {"backgroundColor": "#1976D2", "color": "#FFFFFF", "fontWeight": "bold"}, "text": "Save"},
                                 "type": "ia.input.button"},
                                {"events": {"component": {"onActionPerformed": {"config": {"script": delete_script()},
                                    "scope": "G", "type": "script"}}},
                                 "meta": {"domId": "sites-delete-btn", "name": "DeleteButton"},
                                 "props": {"style": {"backgroundColor": "#C62828", "color": "#FFFFFF", "fontWeight": "bold"}, "text": "Delete"},
                                 "type": "ia.input.button"},
                                {"events": {"component": {"onActionPerformed": {"config": {
                                    "script": "\tc = self.view.custom\n" +
                                              field_clears("\t") +
                                              "\n\tc.recordId = 0\n\tc.statusMsg = \"Cleared — ready for a new site.\"\n\tsystem.util.getLogger(\"SPIKE\").info(\"SPIKE Sites form cleared\")\n"},
                                    "scope": "G", "type": "script"}}},
                                 "meta": {"domId": "sites-clear-btn", "name": "ClearButton"},
                                 "props": {"style": {"backgroundColor": "#607D8B", "color": "#FFFFFF"}, "text": "Clear"},
                                 "type": "ia.input.button"},
                            ],
                        },
                    ],
                },
            ],
        },
    }
    return view


def _write_view(path, view):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w") as f:
        json.dump(view, f, indent=1, sort_keys=True)
        f.write("\n")
    print("wrote", path)


def main():
    view = build_view()
    # 1) deployed gateway copy (what the running session serves)
    _write_view(GW_OUT, view)
    # 2) committed repo copy (the redeployable source of truth — must == runtime)
    _write_view(REPO_OUT, view)
    # 2b) the repo resource.json companion (scope/version/files only; gateway re-signs attributes)
    with open(REPO_RESOURCE, "w") as f:
        json.dump(RESOURCE_JSON, f, indent=2)
        f.write("\n")
    print("wrote", REPO_RESOURCE)


if __name__ == "__main__":
    main()
