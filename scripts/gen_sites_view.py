#!/usr/bin/env python3
"""Generate the Sites-master combined master-detail Perspective view.json.

EIGHTH master-data module. Mirrors the PROVEN canonical Size view
(Master/Size/Size) field-for-field — single combined view (grid left, detail form
right), inline system.db.runPrepQuery, config.bidirectional:true on every input
binding, RESTRICT refCount delete-gate, 16-char yyyymmddHHMMSSff audit stamp,
same-view recordId prop-write (NO navigation).

SITES-MASTER-SPECIFIC (differs from every other master — see the SQL header in
docs/analysis/master-data/master-crud-namedqueries.sql):
  #1 ADMIN-GATED  -> in-view custom.isAdmin gate (defense-in-depth; the enforced
                     role-gate needs Designer view-permissions + an IdP this spike
                     lacks — reported, not faked).
  #2 NOT SITE-SCOPED (INVERTED IG-SITE seam) -> list/get have NO site filter.
  #3 read-only system fields shown disabled (IN_SITE_ID, IN_EIN_SEQ,
                     VC_LAST_FORECAST_IMPORT, VC_LAST_UPDATE, VC_ADD).
  #4 typed inputs: BIT_* -> checkbox; VC_FORECAST_IMPORT_MODE -> AUTO/MANUAL
                     dropdown; IN_* -> numeric (CHECKs enforced client-side:
                     fill_days<=50, retention>=12); VC_SEP_* -> 1-char; rest text.
  #6 refCount counts the throwaway INV_PARTS_STOCK_MST.site_id (no FK yet).

This writes view.json only; resource.json is written separately (Size shape).
"""
import json
import os

OUT = "/usr/local/ignition/data/projects/spike/com.inductiveautomation.perspective/views/Master/Sites/Sites/view.json"

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
    # audit (read-only)
    ("form_lastupd",   "sites-lastupd",   "Last Update (audit)",     "ro_text", "VC_LAST_UPDATE","",None, "audit"),
    ("form_add",       "sites-add",       "Created (audit)",         "ro_text", "VC_ADD",       "", None, "audit"),
]

GROUPS = [
    ("identity", "Identity / Address"),
    ("edi",      "EDI / Trading Identity"),
    ("order",    "Order / Forecast Config"),
    ("purge",    "Data Purge"),
    ("audit",    "Audit (read-only)"),
]

# Writable fields in the canonical save order (insert/update column order) — must
# match the Sites/insert + Sites/update SQL param order. Excludes read-only (#3).
WRITABLE = [f for f in F if f[3] not in ("ro_text", "ro_num")]


def custom_defaults():
    c = {"recordId": 0, "runNonce": 0, "searchTerm": "", "statusMsg": "", "isAdmin": False}
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
    binding = {
        "config": {"bidirectional": True, "path": "view.custom.%s" % key},
        "type": "property",
    }
    # read-only system fields: one-way (no bidirectional write-back) — rule #3.
    if readonly:
        binding = {"config": {"path": "view.custom.%s" % key}, "type": "property"}
    field = {
        "meta": {"domId": domid, "name": "fld_" + key},
        "position": {"basis": "320px"},
        "propConfig": {vprop: {"binding": binding}},
        "type": comp_type(kind),
    }
    if inner_props:
        field["props"] = inner_props
    lbl = {
        "meta": {"name": "lbl_" + key},
        "position": {"basis": "200px"},
        "props": {"style": {"fontSize": "12px"}, "text": label},
        "type": "ia.display.label",
    }
    return {
        "meta": {"name": "row_" + key},
        "props": {"alignItems": "center", "style": {"gap": "10px"}},
        "type": "ia.container.flex",
        "children": [lbl, field],
    }


def group_block(gkey, gtitle):
    rows = [field_row(key, domid, label, kind, maxlen)
            for (key, domid, label, kind, dbcol, default, maxlen, grp) in F if grp == gkey]
    header = {
        "meta": {"name": "grp_" + gkey},
        "props": {"style": {"color": "#1976D2", "fontSize": "12px", "fontWeight": "bold",
                            "marginTop": "4px", "borderBottom": "1px solid #D0D4DA"},
                  "text": gtitle},
        "type": "ia.display.label",
    }
    return [header] + rows


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


def new_script():
    clears = "\n".join(
        "\tc.%s = %s" % (key, ('""' if kind in ("text", "char1", "ro_text", "mode") and default == "" else repr(default)))
        for (key, _, _, kind, _, default, _, _) in F if kind not in ("ro_text", "ro_num"))
    # mode default 'AUTO' should be a string literal
    return (
        "\t# New record: same-view prop write (NO navigation). recordId=0 -> insert mode.\n"
        "\t# Admin-gate (rule #1, defense-in-depth): non-admins cannot start a new site.\n"
        "\tc = self.view.custom\n"
        "\tif not c.isAdmin:\n"
        "\t\tc.statusMsg = \"Admin only — you do not have permission to add sites.\"\n"
        "\t\tsystem.util.getLogger(\"SPIKE\").info(\"SPIKE Sites New BLOCKED: not admin\"); return\n"
        + clears + "\n"
        "\tc.recordId = 0\n"
        "\tc.statusMsg = \"New site — enter details and Save.\"\n"
        "\tsystem.util.getLogger(\"SPIKE\").info(\"SPIKE Sites New -> recordId=0 (insert mode)\")\n"
    )


def save_script():
    # ordered param lists matching Sites/insert and Sites/update SQL
    # Nullable "unset = NULL" INT columns (schema: all four INT cols are NULL-able and
    # their CHECKs allow NULL; the table comment treats NULL as legacy "unset"). Blank/0
    # MUST bind NULL, never literal 0 — 0 is rejected by CK_INV_SITES_DATA_RETENTION
    # (>= 12) and is not the legacy "unset" value for the others (B1/N1 fix).
    NULLABLE_UNSET = {"form_retention", "form_filldays", "form_fcusage", "form_maxseq"}

    def pyval(key, kind):
        if kind == "bit":
            return "_bit(c.%s)" % key
        if kind == "num":
            # nullable-unset INTs: blank/0 -> NULL; otherwise the typed value.
            if key in NULLABLE_UNSET:
                return "_intn(c.%s)" % key
            return "_int(c.%s)" % key
        if kind == "mode":
            return "_mode(c.%s)" % key
        return "_str(c.%s)" % key

    ins_cols = ("VC_SITE_NAME, VC_SITE_ABBR, VC_STREET, VC_CITY, VC_STATE, VC_COUNTRY, VC_ZIP, "
                "VC_DUNS, VC_SUPPLIER_CODE, VC_DOCK_CODE, IN_EIN_SEQ, VC_EDI_MODE, "
                "VC_SEP_SEGMENT, VC_SEP_ELEMENT, VC_SEP_SUBELEMENT, "
                "VC_TMM_NAME, VC_TMM_ABBR, VC_TMM_DUNS, IN_MAX_SEQUENCE, "
                "BIT_ACCEPT_ANY_ORDER_ASN, VC_DELIVERY_METHOD_CODE, IN_FILL_DAYS, "
                "IN_FORECAST_USAGE_COMPARE, BIT_USE_FIRST_PRODUCTION_DAY, VC_FORECAST_IMPORT_MODE, "
                "BIT_ENABLE_DATA_PURGE, BIT_PROMPT_DATA_PURGE, IN_DATA_RETENTION, VC_ADD")
    # the VALUES placeholders: IN_EIN_SEQ is the literal 0 (rule #3); VC_ADD is the audit expr.
    ins_vals = ("?,?,?,?,?,?,?, ?,?,?, 0, ?, ?,?,?, ?,?,?,?, ?,?, ?, ?,?,?, ?,?,?, " + AUDIT)
    # ordered python args for insert (skip IN_EIN_SEQ which is literal 0; VC_ADD is in SQL)
    ins_order = [
        ("form_name", "text"), ("form_abbr", "text"), ("form_street", "text"), ("form_city", "text"),
        ("form_state", "text"), ("form_country", "text"), ("form_zip", "text"),
        ("form_duns", "text"), ("form_supcode", "text"), ("form_dock", "text"),
        ("form_edimode", "text"),
        ("form_sepseg", "text"), ("form_sepelem", "text"), ("form_sepsub", "text"),
        ("form_tmmname", "text"), ("form_tmmabbr", "text"), ("form_tmmduns", "text"), ("form_maxseq", "num"),
        ("form_acceptasn", "bit"), ("form_delivery", "text"), ("form_filldays", "num"),
        ("form_fcusage", "num"), ("form_usefpd", "bit"), ("form_fcmode", "mode"),
        ("form_enpurge", "bit"), ("form_prpurge", "bit"), ("form_retention", "num"),
    ]
    ins_args = "[" + ", ".join(pyval(k, kd) for k, kd in ins_order) + "]"

    upd_set = ("VC_SITE_NAME=?, VC_SITE_ABBR=?, VC_STREET=?, VC_CITY=?, VC_STATE=?, VC_COUNTRY=?, VC_ZIP=?, "
               "VC_DUNS=?, VC_SUPPLIER_CODE=?, VC_DOCK_CODE=?, VC_EDI_MODE=?, "
               "VC_SEP_SEGMENT=?, VC_SEP_ELEMENT=?, VC_SEP_SUBELEMENT=?, "
               "VC_TMM_NAME=?, VC_TMM_ABBR=?, VC_TMM_DUNS=?, IN_MAX_SEQUENCE=?, "
               "BIT_ACCEPT_ANY_ORDER_ASN=?, VC_DELIVERY_METHOD_CODE=?, IN_FILL_DAYS=?, "
               "IN_FORECAST_USAGE_COMPARE=?, BIT_USE_FIRST_PRODUCTION_DAY=?, VC_FORECAST_IMPORT_MODE=?, "
               "BIT_ENABLE_DATA_PURGE=?, BIT_PROMPT_DATA_PURGE=?, IN_DATA_RETENTION=?, "
               "VC_LAST_UPDATE = " + AUDIT)
    upd_args = "[" + ", ".join(pyval(k, kd) for k, kd in ins_order) + ", recId]"

    return (
        "\t# ---- Sites Save (insert-or-update) ----\n"
        "\t# recordId==0 -> insert (IDENTITY assigns id, VC_ADD stamped); >0 -> update\n"
        "\t# (VC_LAST_UPDATE stamped). RULE #3: IN_SITE_ID/IN_EIN_SEQ/VC_LAST_FORECAST_IMPORT\n"
        "\t# are NEVER written by the form. RULE #2: no site predicate.\n"
        "\tlog = system.util.getLogger(\"SPIKE\")\n"
        "\tDB = \"Inventory_Spike\"\n"
        "\tc = self.view.custom\n"
        "\trecId = int(c.recordId or 0)\n"
        "\t# Admin-gate (rule #1, defense-in-depth):\n"
        "\tif not c.isAdmin:\n"
        "\t\tc.statusMsg = \"Admin only — you do not have permission to save sites.\"\n"
        "\t\tlog.info(\"SPIKE Sites save BLOCKED: not admin\"); return\n"
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
    clears = "\n".join(
        "\t\tc.%s = %s" % (key, ('""' if kind in ("text", "char1", "mode") and default == "" else repr(default)))
        for (key, _, _, kind, _, default, _, _) in F if kind not in ("ro_text", "ro_num"))
    return (
        "\t# ---- Sites Delete — RESTRICT gate (rule #6) ----\n"
        "\t# refCount counts the THROWAWAY INV_PARTS_STOCK_MST.site_id (no FK yet).\n"
        "\t# ⚠️ MUST be EXTENDED to every IN_SITE_ID child once M4 wires the FKs.\n"
        "\t# Block on any non-zero total; never delete a referenced site.\n"
        "\tlog = system.util.getLogger(\"SPIKE\")\n"
        "\tDB = \"Inventory_Spike\"\n"
        "\tc = self.view.custom\n"
        "\trecId = int(c.recordId or 0)\n"
        "\tabbr = (unicode(c.form_abbr) if c.form_abbr is not None else u\"\").strip()\n"
        "\tif not c.isAdmin:\n"
        "\t\tc.statusMsg = \"Admin only — you do not have permission to delete sites.\"\n"
        "\t\tlog.info(\"SPIKE Sites DELETE BLOCKED: not admin\"); return\n"
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
            # ---- Admin-gate (rule #1, defense-in-depth) ----
            # AUTHORITATIVE prod gate is isAuthorized(false,"Authenticated/Roles/Admin"),
            # which needs an IdP + an "Admin" security level THIS HEADLESS SPIKE LACKS
            # (anonymous sessions, no IdP) -> it fails CLOSED (false) here, the correct
            # prod behavior. The OR-branch is a SPIKE-ONLY escape hatch: the e2e harness
            # opens /sites?qaAdmin=1 (page.props.urlParams.qaAdmin). In prod the page is
            # role-gated in the Designer so the URL hatch is unreachable by non-admins.
            # IG83-TODO: drop the qaAdmin branch + enforce via Designer view permissions.
            "custom.isAdmin": {
                "binding": {
                    "config": {
                        "expression": "isAuthorized(false, \"Authenticated/Roles/Admin\") || {page.props.urlParams.qaAdmin} = \"1\""
                    },
                    "type": "expr",
                }
            },
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
                                      "text": "Sites Master — List + Detail  ·  INV_SITES (Inventory_Spike)  ·  ADMIN-ONLY · ALL SITES (not site-scoped)"},
                            "type": "ia.display.label",
                        },
                        # admin banner (visible only to non-admins)
                        {
                            "meta": {"domId": "sites-admin-banner", "name": "AdminBanner"},
                            "propConfig": {
                                "props.text": {"binding": {"config": {
                                    "expression": "if({view.custom.isAdmin}, '', 'ADMIN ONLY — this screen is restricted. (Enforced role-gate pending Designer view permissions + IdP.)')"},
                                    "type": "expr"}},
                                "meta.visible": {"binding": {"config": {
                                    "expression": "!{view.custom.isAdmin}"}, "type": "expr"}},
                            },
                            "props": {"style": {"backgroundColor": "#FFF3CD", "border": "1px solid #FFC107",
                                                "color": "#7A5C00", "fontSize": "12px", "fontWeight": "bold",
                                                "padding": "6px 8px"}},
                            "type": "ia.display.label",
                        },
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
                                   "text": "Admin-only. NOT site-scoped — shows ALL sites (inverted IG-SITE seam). Single combined view; row-select is an in-view prop write, no navigation. IN_SITE_ID / IN_EIN_SEQ / last-forecast-import / audit fields are read-only (system-maintained)."},
                         "type": "ia.display.label"},
                    ],
                },
                # ---- RIGHT PANE: detail ----
                {
                    "meta": {"name": "RightPane"},
                    "position": {"basis": "640px", "grow": 1},
                    "props": {"direction": "column", "style": {"gap": "8px"}},
                    "type": "ia.container.flex",
                    "children": [
                        {"meta": {"name": "DetailTitle"},
                         "props": {"style": {"color": "#222222", "fontSize": "15px", "fontWeight": "bold"}, "text": "Detail"},
                         "type": "ia.display.label"},
                        {"meta": {"domId": "sites-status", "name": "StatusMsg"},
                         "propConfig": {"props.text": {"binding": {"config": {"path": "view.custom.statusMsg"}, "type": "property"}}},
                         "props": {"style": {"color": "#B71C1C", "fontSize": "13px", "fontWeight": "bold", "padding": "6px 8px"}},
                         "type": "ia.display.label"},
                        # form (hidden from non-admins; rule #1)
                        {
                            "meta": {"domId": "sites-form", "name": "Form"},
                            "propConfig": {"meta.visible": {"binding": {"config": {
                                "expression": "{view.custom.isAdmin}"}, "type": "expr"}}},
                            "props": {"direction": "column",
                                      "style": {"backgroundColor": "#FFFFFF", "border": "1px solid #D0D4DA",
                                                "borderRadius": "4px", "gap": "6px", "padding": "12px",
                                                "overflow": "auto", "maxHeight": "640px"}},
                            "type": "ia.container.flex",
                            "children": children_form,
                        },
                        # action bar
                        {
                            "meta": {"name": "ActionBar"},
                            "position": {"shrink": 0},
                            "propConfig": {"meta.visible": {"binding": {"config": {
                                "expression": "{view.custom.isAdmin}"}, "type": "expr"}}},
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
                                              "\n".join("\tc.%s = %s" % (key, ('""' if kind in ("text","char1","mode") and default=="" else repr(default))) for (key,_,_,kind,_,default,_,_) in F if kind not in ("ro_text","ro_num")) +
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


def main():
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    view = build_view()
    with open(OUT, "w") as f:
        json.dump(view, f, indent=1, sort_keys=True)
        f.write("\n")
    print("wrote", OUT)


if __name__ == "__main__":
    main()
