#!/usr/bin/env python3
"""gen_hotcall_view.py — emit the Perspective HOT-CALL ENTRY view (project-as-code JSON) for the
"One Cycle Entry" screen (the P12 entry path), replacing the legacy fixed-12 positional
`Controls[i±1]` coupling (HotCallEntry.pas:79-89/:258-285) with a proper editable part/qty TABLE.

What this view is (the SHELL — see the Designer-finish notes at the bottom):
  * a LINE dropdown (options from AD_GetLines — HotCallEntry.pas:101/112) bound to view.custom.line.
  * an ASN production-date input bound to view.custom.prodDate (yyyymmdd).
  * an 8-char MANIFEST text field bound to view.custom.manifest (the >=8 / non-'7' validation lives in
    the driver hotcall.create_hotcall_asn; the field caps display length at 8).
  * a part/qty TABLE (ia.display.table, editable) bound to view.custom.items — N rows of {part, qty},
    NOT 12 fixed positional controls. An "Add row" button appends a blank row; onEditCellCommit writes
    the edited cell back into view.custom.items.
  * a "Create Hot-Call ASN" BUTTON whose onActionPerformed posts a gateway message
    ("create_hotcall_asn") carrying {line, prodDate, manifest, items} — the gateway message handler
    (project-library/hotcall/message-handler.py) calls the REAL create_hotcall_asn driver in gateway
    scope (the DB write must NOT run in the browser/session scope).

HEADLESS-AUTHORABLE (proven patterns, memory reference-headless-ignition-authoring-limits):
  * the view.json structure, the custom props, the table columns, the button onActionPerformed script,
    and the bidirectional bindings (config.bidirectional:true) are all hand-authorable on disk.
  * resource.json carries scope/version/files + a placeholder attributes block (the gateway re-signs).

NEEDS THE DESIGNER TO FINISH (do NOT fake these — they are NOT reliably hand-authorable headless):
  * the AD_GetLines dropdown OPTIONS binding as a real Named Query (the data.bin is Ignition XML
    serialization — SAXParseException on a hand-authored file). Authored here as a script-transform on a
    custom prop that calls system.db.runPrepQuery("EXEC AD_GetLines") (the proven Order/Sites workaround);
    promotable to a real NQ in the Designer.
  * the PAGE/SESSION config (mounting the view at a route so it appears in the Session Launcher) — a
    page-config resource is Designer-authored (the param-mapping structure a hand route doesn't reproduce).
  * the message-handler RESOURCE wiring (a Gateway Message Handler is bound in the Designer's
    Project Properties / Scripting > Message Handlers; the script body is in
    project-library/hotcall/message-handler.py, ready to paste/bind).

Usage:  python3 scripts/gen_hotcall_view.py
  -> writes docs/analysis/edi/project-library/hotcall/perspective-views/HotCall/HotCallEntry/{view.json,resource.json}
"""
import json, os

ROOT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
OUT_DIR = os.path.join(ROOT, "ignition", "perspective-views", "HotCall", "HotCallEntry")

# --- the button onActionPerformed script: gather the form, post the gateway message, surface the result.
# IG81-COMPAT: system.util.sendRequest + system.perspective.* are identical on 8.1.52 and 8.3.
CREATE_SCRIPT = (
    "\t# Create Hot-Call ASN: gather the form, post the 'create_hotcall_asn' gateway message (the DB\n"
    "\t# write MUST run in gateway scope, not the session). The message handler calls the REAL driver.\n"
    "\tlog = system.util.getLogger(\"SPIKE.hotcall.view\")\n"
    "\tc = self.view.custom\n"
    "\titems = []\n"
    "\tfor r in (c.items or []):\n"
    "\t\tpart = (r.get(\"part\") or \"\").strip() if hasattr(r, \"get\") else \"\"\n"
    "\t\tqty = r.get(\"qty\") if hasattr(r, \"get\") else None\n"
    "\t\tif part == \"\" and (qty is None or str(qty).strip() == \"\"):\n"
    "\t\t\tcontinue   # skip blank rows (the driver also skips, but trim here for a clean payload)\n"
    "\t\titems.append({\"part\": part, \"qty\": qty})\n"
    "\tpayload = {\"line\": c.line, \"prodDate\": c.prodDate, \"manifest\": c.manifest, \"items\": items}\n"
    "\tlog.info(\"SPIKE hotcall view -> create_hotcall_asn payload=%s\" % payload)\n"
    "\ttry:\n"
    "\t\t# project name comes from the session; sendRequest blocks for the handler's return value.\n"
    "\t\tproj = self.session.props.projectName\n"
    "\t\tres = system.util.sendRequest(proj, \"create_hotcall_asn\", payload)\n"
    "\t\tc.statusMsg = \"Created hot-call ASN %s (%d parts, qty %s).\" % (\n"
    "\t\t\tres.get(\"asnId\"), len(res.get(\"details\", [])), res.get(\"qty\"))\n"
    "\t\tc.items = [{\"part\": \"\", \"qty\": \"\"}]   # reset the grid (ClearEntries parity)\n"
    "\t\tlog.info(\"SPIKE hotcall view -> created ASN %s\" % res.get(\"asnId\"))\n"
    "\texcept Exception as e:\n"
    "\t\tc.statusMsg = \"Failed: %s\" % e\n"
    "\t\tlog.warn(\"SPIKE hotcall view -> create FAILED: %s\" % e)\n"
)

# --- "Add row" appends a blank {part, qty} to the editable table.
ADDROW_SCRIPT = (
    "\tc = self.view.custom\n"
    "\trows = list(c.items or [])\n"
    "\trows.append({\"part\": \"\", \"qty\": \"\"})\n"
    "\tc.items = rows\n"
)

# --- onEditCellCommit writes the edited cell back into view.custom.items (replace-whole-list so the
#     table re-binds; sub-path write-back does not work — headless-authoring-limits memo).
EDITCELL_SCRIPT = (
    "\tc = self.view.custom\n"
    "\trows = list(c.items or [])\n"
    "\tidx = event.row\n"
    "\tcol = event.column            # the column id ('part' or 'qty')\n"
    "\tif 0 <= idx < len(rows):\n"
    "\t\trow = dict(rows[idx])\n"
    "\t\trow[col] = event.value\n"
    "\t\trows[idx] = row\n"
    "\t\tc.items = rows             # replace the whole list so the table re-keys (no sub-path write)\n"
)

# --- the AD_GetLines dropdown options: a custom-prop script-transform calling runPrepQuery (the proven
#     headless workaround; promotable to a real NQ in the Designer). Returns [{value,label}].
LINEOPTS_SCRIPT = (
    "\t# AD_GetLines (HotCallEntry.pas:101, on the ALC dataset) -> dropdown options. Headless workaround\n"
    "\t# for the NQ data.bin wall (reference-headless-ignition-authoring-limits): inline runPrepQuery in a\n"
    "\t# script transform. AD_GetLines lives in the VehicleOrder (ALC) DB, NOT Inventory — same connection\n"
    "\t# create_asn uses for the ALC cross-DB read (asn/code.py ALC_DATABASE='VehicleOrder').\n"
    "\tlog = system.util.getLogger(\"SPIKE.hotcall.view\")\n"
    "\tout = []\n"
    "\ttry:\n"
    "\t\tds = system.db.runPrepQuery(\"EXEC AD_GetLines\", [], \"VehicleOrder\")\n"
    "\t\tfor r in ds:\n"
    "\t\t\tname = r[\"LineName\"]\n"
    "\t\t\tout.append({\"value\": name, \"label\": name})\n"
    "\texcept Exception as e:\n"
    "\t\tlog.warn(\"SPIKE hotcall view -> AD_GetLines failed: %s\" % e)\n"
    "\treturn out\n"
)


def _bidi(path):
    return {"binding": {"type": "property", "config": {"path": path, "bidirectional": True}}}


def _label(name, text, color="#37474F", size="12px", bold=False):
    style = {"color": color, "fontSize": size}
    if bold:
        style["fontWeight"] = "bold"
    return {"meta": {"name": name}, "props": {"text": text, "style": style}, "type": "ia.display.label"}


def build_view():
    return {
        "custom": {
            "line": "",
            "prodDate": "",
            "manifest": "",
            "items": [{"part": "", "qty": ""}],
            "lineOptions": [],
            "statusMsg": "",
        },
        "params": {},
        "propConfig": {
            # AD_GetLines options via a script-transform on the custom prop (runs at load).
            "custom.lineOptions": {"binding": {
                "type": "expr", "config": {"expression": "now(0)"},
                "transforms": [{"type": "script", "code": LINEOPTS_SCRIPT}]}},
        },
        "props": {"defaultSize": {"width": 900, "height": 640}},
        "root": {
            "meta": {"name": "root"},
            "props": {"direction": "column", "style": {"gap": "10px", "padding": "16px"}},
            "type": "ia.container.flex",
            "children": [
                _label("Title", "Hot-Call Entry (One Cycle Entry)", color="#222222", size="18px",
                       bold=True),
                _label("Subtitle",
                       "Urgent out-of-cycle shipment. Manifest 8 chars, non-'7' (M390). "
                       "EIN allocated at 856 send. File goes out as 8HC...",
                       color="#666666", size="12px"),
                # --- header bar: line / date / manifest ---
                {
                    "meta": {"name": "HeaderBar"},
                    "props": {"direction": "row", "style": {"gap": "16px"}, "wrap": "wrap"},
                    "type": "ia.container.flex",
                    "children": [
                        {"meta": {"name": "LineWrap"},
                         "props": {"direction": "column", "style": {"gap": "2px"}},
                         "type": "ia.container.flex",
                         "children": [
                             _label("LineCap", "Line", color="#666666", size="11px"),
                             {"meta": {"domId": "hotcall-line", "name": "LineDropdown"},
                              "position": {"basis": "180px"},
                              "propConfig": {
                                  "props.value": _bidi("view.custom.line"),
                                  "props.options": {"binding": {"type": "property", "config": {
                                      "path": "view.custom.lineOptions"}}},
                              },
                              "props": {"placeholder": "Select a line"},
                              "type": "ia.input.dropdown"},
                         ]},
                        {"meta": {"name": "DateWrap"},
                         "props": {"direction": "column", "style": {"gap": "2px"}},
                         "type": "ia.container.flex",
                         "children": [
                             _label("DateCap", "ASN production date (yyyymmdd)", color="#666666",
                                    size="11px"),
                             {"meta": {"domId": "hotcall-pdate", "name": "ProdDateField"},
                              "position": {"basis": "140px"},
                              "propConfig": {"props.text": _bidi("view.custom.prodDate")},
                              "props": {"placeholder": "20260618"},
                              "type": "ia.input.text-field"},
                         ]},
                        {"meta": {"name": "ManifestWrap"},
                         "props": {"direction": "column", "style": {"gap": "2px"}},
                         "type": "ia.container.flex",
                         "children": [
                             _label("ManifestCap", "Manifest (8 chars)", color="#666666", size="11px"),
                             {"meta": {"domId": "hotcall-manifest", "name": "ManifestField"},
                              "position": {"basis": "160px"},
                              "propConfig": {"props.text": _bidi("view.custom.manifest")},
                              "props": {"maxLength": 8, "placeholder": "52089698"},
                              "type": "ia.input.text-field"},
                         ]},
                    ],
                },
                # --- the part/qty TABLE (replaces the legacy fixed-12 positional controls) ---
                _label("ItemsCap", "Parts (add a row per part/qty — NOT 12 fixed slots)",
                       color="#37474F", size="12px", bold=True),
                {
                    "events": {"component": {"onEditCellCommit": {"config": {
                        "script": EDITCELL_SCRIPT}, "scope": "G", "type": "script"}}},
                    "meta": {"domId": "hotcall-items", "name": "ItemsTable"},
                    "position": {"basis": "300px", "grow": 1},
                    "propConfig": {
                        "props.data": {"binding": {"type": "property", "config": {
                            "path": "view.custom.items"}}},
                    },
                    "props": {
                        "columns": [
                            {"field": "part", "header": {"title": "Part Number"},
                             "editable": True, "sort": "none",
                             "render": "auto", "viewParams": {}},
                            {"field": "qty", "header": {"title": "Qty"},
                             "editable": True, "sort": "none",
                             "render": "auto", "viewParams": {}},
                        ],
                        "filtering": {"enabled": False},
                        "pager": {"activeOption": 1000, "bottom": False, "options": [1000], "top": False},
                        "selection": {"enabled": False},
                    },
                    "type": "ia.display.table",
                },
                # --- action bar ---
                {
                    "meta": {"name": "ActionBar"},
                    "props": {"alignItems": "center", "style": {"gap": "12px", "padding": "8px 0"}},
                    "type": "ia.container.flex",
                    "children": [
                        {"events": {"component": {"onActionPerformed": {"config": {
                            "script": ADDROW_SCRIPT}, "scope": "G", "type": "script"}}},
                         "meta": {"domId": "hotcall-addrow-btn", "name": "AddRowButton"},
                         "props": {"style": {"backgroundColor": "#607D8B", "color": "#FFFFFF"},
                                   "text": "Add row"},
                         "type": "ia.input.button"},
                        {"events": {"component": {"onActionPerformed": {"config": {
                            "script": CREATE_SCRIPT}, "scope": "G", "type": "script"}}},
                         "meta": {"domId": "hotcall-create-btn", "name": "CreateButton"},
                         "props": {"style": {"backgroundColor": "#388E3C", "color": "#FFFFFF",
                                             "fontWeight": "bold"},
                                   "text": "Create Hot-Call ASN"},
                         "type": "ia.input.button"},
                    ],
                },
                {"meta": {"domId": "hotcall-status", "name": "StatusMsg"},
                 "propConfig": {"props.text": {"binding": {"type": "property", "config": {
                     "path": "view.custom.statusMsg"}}}},
                 "props": {"style": {"color": "#1B5E20", "fontSize": "13px", "fontWeight": "bold",
                                     "padding": "6px 8px"}},
                 "type": "ia.display.label"},
                {"meta": {"name": "FooterNote"},
                 "props": {"style": {"color": "#37474F", "fontSize": "11px", "fontStyle": "italic"},
                           "text": "SHELL view. AD_GetLines options via inline runPrepQuery (promotable "
                                   "to a Named Query in the Designer). The 'Create' button posts the "
                                   "create_hotcall_asn gateway message (handler in "
                                   "project-library/hotcall/message-handler.py). Page/session mounting + "
                                   "the message-handler binding need a Designer pass — see "
                                   "gen_hotcall_view.py header."},
                 "type": "ia.display.label"},
            ],
        },
    }


RESOURCE = {
    "scope": "G",
    "version": 1,
    "restricted": False,
    "overridable": True,
    "files": ["view.json"],
    "attributes": {
        "lastModification": {"actor": "external", "timestamp": "2026-06-22T00:00:00Z"},
        # placeholder signature; the gateway re-signs on load (headless-authoring-limits #5).
        "lastModificationSignature":
            "0000000000000000000000000000000000000000000000000000000000000000",
    },
}


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    with open(os.path.join(OUT_DIR, "view.json"), "w") as fh:
        json.dump(build_view(), fh, indent=2)
        fh.write("\n")
    with open(os.path.join(OUT_DIR, "resource.json"), "w") as fh:
        json.dump(RESOURCE, fh, indent=2)
        fh.write("\n")
    print("wrote %s/{view.json,resource.json}" % OUT_DIR)


if __name__ == "__main__":
    main()
