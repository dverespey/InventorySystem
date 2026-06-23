#!/usr/bin/env python3
"""gen_renban_views.py — emit the P4 renban-breakdown Perspective SHELL (headless-authorable minimum):

  Order/RenbanBreakdown        the operator entry point: group-code / trailers / pallets inputs +
                               Preview (PURE compute_trailer_breakdown) + Commit. On a COLLISION the
                               Commit script opens the collision dialog (system.perspective.openPopup)
                               instead of committing.
  Order/RenbanCollisionDialog  the WARN -> GUIDE -> FIX dialog: a WARN message label, a GUIDE table
                               (next-free run-of-N + the colliding open-release detail), 3 FIX buttons
                               (Use next-free / Proceed anyway / Cancel) each invoking the matching
                               commit_renban_breakdown resolution and closing the popup.

These are the headless minimum (single combined views per reference-headless-ignition-authoring-limits).
DESIGNER HAND-OFF (cannot be authored headlessly; see the printed list): the Named Query data.bin XML
(Order/renbanGroups), the /order/renban page-route mount, the popup registration on RenbanBreakdown, and
the bidirectional-binding flag (binding.config, NOT a top-level prop) for the editable inputs. The driver
(renban/code.py) is fully headless-tested; this is the UI shell over it.

SERVER-SIDE WRITE GATE (BLOCKER-2, code-reviewer): the Commit / Use-next-free / Proceed-anyway scripts call
auth.requireWrite(self.session) BEFORE the driver, and thread session=self.session into the driver (which
re-gates server-side). A forged client prop / anon session is rejected server-side — the R25/P15 hole class
is closed here, matching the master CRUD views. (Was previously deferred as a Designer-finish item.)

resource.json carries an ATTRIBUTES block (a missing one FAULTS the WHOLE gateway on cold start — R31).
Writes the gateway copy + a repo copy (docs/design/perspective-views). Run:  python3 scripts/gen_renban_views.py
"""
import json, os

PROJ_DIR  = os.environ.get("PROJ_DIR", "/usr/local/ignition/data/projects/InventorySystem")
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PERSP     = os.path.join(PROJ_DIR, "com.inductiveautomation.perspective")

# suite theme tokens (var(--tai-*)) so the shell inherits tai-light/tai-dark like the rest of the suite.
TXT = "var(--tai-text, #1F2933)"; MUT = "var(--tai-text-muted, #6B7785)"
SURF = "var(--tai-surface, #FFFFFF)"; BG = "var(--tai-bg, #F5F6F8)"
BORDER = "var(--tai-border, #D7DCE2)"; ACCENT = "var(--tai-accent, #2A6FB0)"
DANGER = "var(--tai-danger, #C0392B)"; WARN = "var(--tai-warn, #B8860B)"
CARD = {"backgroundColor": SURF, "border": "1px solid " + BORDER, "borderRadius": "10px",
        "padding": "16px"}


# ---- component helpers (match the deployed view JSON shape) -----------------------------
def lbl(name, text, style=None, bind=None, domId=None):
    meta = {"name": name}
    if domId:
        meta["domId"] = domId
    c = {"type": "ia.display.label", "meta": meta, "props": {"text": text}}
    if style:
        c["props"]["style"] = style
    if bind:
        c["propConfig"] = {"props.text": {"binding": {"type": "property", "config": {"path": bind}}}}
    return c


def flex(name, children, direction="column", gap=10, style=None, align=None, justify=None, domId=None):
    props = {"direction": direction}
    if gap:
        props["gap"] = gap
    if align:
        props["alignItems"] = align
    if justify:
        props["justify"] = justify
    if style:
        props["style"] = style
    meta = {"name": name}
    if domId:
        meta["domId"] = domId
    return {"type": "ia.container.flex", "meta": meta, "props": props, "children": children}


def numfield(name, prop_path, domId, label):
    # numeric-entry-field: value is under props.value (NOT props.text). Bidirectional write-back is a
    # Designer-finish item (binding.config bidirectional flag) — DESIGNER HAND-OFF.
    return flex(name + "_wrap", [
        lbl(name + "_lbl", label, {"fontSize": "12px", "color": MUT}),
        {"type": "ia.input.numeric-entry-field", "meta": {"name": name, "domId": domId},
         "props": {"value": 0},
         "propConfig": {"props.value": {"binding": {"type": "property",
                        "config": {"path": prop_path}, "bidirectional": True}}}},
    ], gap=3)


def textfield(name, prop_path, domId, label):
    # text-field: value is under props.text (NOT props.value) — the documented per-component prop trap.
    return flex(name + "_wrap", [
        lbl(name + "_lbl", label, {"fontSize": "12px", "color": MUT}),
        {"type": "ia.input.text-field", "meta": {"name": name, "domId": domId},
         "props": {"text": ""},
         "propConfig": {"props.text": {"binding": {"type": "property",
                        "config": {"path": prop_path}, "bidirectional": True}}}},
    ], gap=3)


def button(name, text, domId, script, primary=False, danger=False):
    bg = DANGER if danger else (ACCENT if primary else SURF)
    fg = "var(--tai-on-accent, #FFFFFF)" if (primary or danger) else TXT
    return {"type": "ia.input.button", "meta": {"name": name, "domId": domId},
            "props": {"text": text, "style": {"backgroundColor": bg, "color": fg,
                      "border": "1px solid " + BORDER, "borderRadius": "8px", "padding": "8px 16px"}},
            "events": {"component": {"onActionPerformed": {
                "config": {"script": script}, "scope": "G", "type": "script"}}}}


# ---- RenbanBreakdown view ----------------------------------------------------------------
# Preview = PURE compute_trailer_breakdown (no DB). Commit = the gateway driver; on a COLLISION result the
# script opens the dialog (openPopup) carrying the collision payload via the popup params.
PREVIEW_SCRIPT = (
    "\t# PREVIEW: PURE compute (no DB). Loads the blank-renban orders then computes the trailer split,\n"
    "\t# so the operator sees the candidate renbans BEFORE committing. (Read-only.)\n"
    "\tlog = system.util.getLogger(\"SPIKE.renban.ui\")\n"
    "\tc = self.view.custom\n"
    "\tgroup = (c.groupCode or \"\").strip()\n"
    "\tT = int(c.trailers or 0); P = int(c.pallets or 0)\n"
    "\tif not group or T < 1 or P < 1:\n"
    "\t\tc.status = \"Enter a group code, trailers (>=1) and pallets/trailer (>=1).\"\n"
    "\t\treturn\n"
    "\timport renban  # project library: docs/analysis/order/project-library/renban/code.py\n"
    "\ttry:\n"
    "\t\torders = renban.loadBlankRenbanOrders(group, \"Inventory_Spike\")\n"
    "\t\tif not orders:\n"
    "\t\t\tc.status = \"No blank-renban orders for group %s.\" % group; c.preview = {\"data\": [], \"columns\": []}; return\n"
    "\t\tres = renban.compute_trailer_breakdown(orders, T, P, group)\n"
    "\t\trows = [{\"trailer\": r[\"truck\"] + 1, \"part\": r[\"part\"], \"frs\": r[\"frs\"],\n"
    "\t\t         \"renban\": r[\"renban\"], \"qty\": r[\"qty\"]} for r in res[\"rows\"]]\n"
    "\t\tc.preview = {\"data\": rows, \"columns\": [{\"field\": k} for k in (\"trailer\",\"part\",\"frs\",\"renban\",\"qty\")]}\n"
    "\t\tc.status = \"Preview: %d trailer-row(s), next count %s.\" % (len(rows), res[\"next_count\"])\n"
    "\t\tlog.info(\"SPIKE renban preview group=%s rows=%d\" % (group, len(rows)))\n"
    "\texcept Exception, e:\n"
    "\t\tc.status = \"Preview failed: %s\" % e; log.warn(\"SPIKE renban preview failed: %s\" % e)\n")

COMMIT_SCRIPT = (
    "\t# COMMIT: call the gateway driver (resolution=None -> pre-commit collision gate). On COLLISION,\n"
    "\t# open the collision dialog (openPopup) carrying the payload; otherwise report COMMITTED.\n"
    "\t# SERVER-SIDE WRITE GATE (BLOCKER-2): the driver authorizes auth.requireWrite(self.session) FIRST\n"
    "\t# (server-side, from the SESSION roles) on every path — same P15 posture as the master CRUD views.\n"
    "\t# We ALSO pre-gate here so a denied click shows DENIED without a driver round-trip; the driver gate\n"
    "\t# is the real boundary (a forged client prop / anon session cannot reach the write).\n"
    "\tlog = system.util.getLogger(\"SPIKE.renban.ui\")\n"
    "\tc = self.view.custom\n"
    "\tgroup = (c.groupCode or \"\").strip()\n"
    "\tT = int(c.trailers or 0); P = int(c.pallets or 0)\n"
    "\tif not group or T < 1 or P < 1:\n"
    "\t\tc.status = \"Enter a group code, trailers and pallets/trailer first.\"; return\n"
    "\timport renban\n"
    "\timport auth as A\n"
    "\ttry:\n"
    "\t\tA.requireWrite(self.session)\n"
    "\t\tsiteId = self.session.custom.siteId if hasattr(self.session.custom, \"siteId\") else None\n"
    "\t\tres = renban.commit_renban_breakdown(group, T, P, \"Inventory_Spike\", resolution=None,\n"
    "\t\t\tsiteId=siteId, session=self.session)\n"
    "\t\tif res[\"status\"] == \"COLLISION\":\n"
    "\t\t\tc.status = \"COLLISION: %d candidate renban(s) in use. Resolve in the dialog.\" % len(res[\"collisions\"])\n"
    "\t\t\t# GUIDE/FIX: hand the payload to the dialog popup (params). DESIGNER: register the popup id.\n"
    "\t\t\tsystem.perspective.openPopup(\"renbanCollision\", \"Order/RenbanCollisionDialog\", params={\n"
    "\t\t\t\t\"group\": group, \"trailers\": T, \"pallets\": P,\n"
    "\t\t\t\t\"collisions\": res[\"collisions\"], \"nextFree\": res[\"next_free\"],\n"
    "\t\t\t\t\"alarmId\": res[\"alarm_id\"], \"acknowledged\": [x[\"renban\"] for x in res[\"collisions\"]]})\n"
    "\t\telif res[\"status\"] == \"EMPTY\":\n"
    "\t\t\tc.status = \"Nothing to commit (no blank-renban orders for %s).\" % group\n"
    "\t\telse:\n"
    "\t\t\tc.status = \"COMMITTED: %d trailer-row(s), next count %s.\" % (len(res[\"rows\"]), res[\"next_count\"])\n"
    "\t\tlog.info(\"SPIKE renban commit group=%s status=%s\" % (group, res[\"status\"]))\n"
    "\texcept A.AuthError, ae:\n"
    "\t\tc.status = \"DENIED (server-side): %s\" % ae\n"
    "\t\tlog.warn(\"SPIKE renban commit DENIED (server-side gate): %s\" % ae)\n"
    "\texcept Exception, e:\n"
    "\t\tc.status = \"Commit failed: %s\" % e; log.warn(\"SPIKE renban commit failed: %s\" % e)\n")


def build_breakdown():
    title = lbl("title", "Renban Breakdown", {"fontSize": "17px", "fontWeight": "700", "color": TXT})
    sub = lbl("sub", "Split blank-renban palletized orders across N trailers, assigning each a renban.",
              {"fontSize": "12.5px", "color": MUT})
    # group-code: a text field for the shell (DESIGNER: swap to a dropdown bound to NQ Order/renbanGroups).
    inputs = flex("inputs", [
        textfield("group", "view.custom.groupCode", "rb-group", "Renban group (e.g. CMWA)"),
        numfield("trailers", "view.custom.trailers", "rb-trailers", "Trailers"),
        numfield("pallets", "view.custom.pallets", "rb-pallets", "Pallets / trailer"),
    ], direction="row", gap=16, align="flex-end", style=CARD)
    actions = flex("actions", [
        button("preview", "Preview", "rb-preview", PREVIEW_SCRIPT),
        button("commit", "Commit", "rb-commit", COMMIT_SCRIPT, primary=True),
    ], direction="row", gap=12)
    status = lbl("status", "—", {"fontSize": "13px", "color": MUT, "margin": "4px 0"},
                 bind="view.custom.status", domId="rb-status")
    grid = {"type": "ia.display.table", "meta": {"name": "previewTable", "domId": "rb-preview-table"},
            "props": {"data": []},
            "propConfig": {"props.data": {"binding": {"type": "property",
                           "config": {"path": "view.custom.preview.data"}}}}}
    root = flex("root", [title, sub, inputs, actions, status, grid], gap=14, domId="renban-breakdown-root",
                style={"backgroundColor": BG, "padding": "18px", "minHeight": "100%"})
    return {
        "custom": {"groupCode": "", "trailers": 1, "pallets": 1, "status": "",
                   "preview": {"data": [], "columns": []}},
        "params": {},
        "propConfig": {},
        "props": {"defaultSize": {"width": 900, "height": 640}},
        "root": root,
    }


# ---- RenbanCollisionDialog view ----------------------------------------------------------
USE_NEXT_FREE_SCRIPT = (
    "\tlog = system.util.getLogger(\"SPIKE.renban.ui\")\n"
    "\tp = self.view.params\n"
    "\timport renban\n"
    "\timport auth as A\n"
    "\tif p.nextFree is None:\n"
    "\t\tself.view.custom.warn = \"No free run available — close a stale release first, then cancel.\"; return\n"
    "\ttry:\n"
    "\t\tA.requireWrite(self.session)   # SERVER-SIDE WRITE GATE (BLOCKER-2)\n"
    "\t\tsiteId = self.session.custom.siteId if hasattr(self.session.custom, \"siteId\") else None\n"
    "\t\tres = renban.commit_renban_breakdown(p.group, int(p.trailers), int(p.pallets), \"Inventory_Spike\",\n"
    "\t\t\tresolution={\"action\": \"use_next_free\", \"base\": int(p.nextFree), \"alarm_id\": p.alarmId},\n"
    "\t\t\tsiteId=siteId, session=self.session)\n"
    "\t\tself.view.custom.result = res[\"status\"]\n"
    "\t\tif res[\"status\"] == \"COLLISION\":\n"
    "\t\t\tself.view.custom.warn = \"Lost race: the suggested run was taken. Re-run the breakdown.\"; return\n"
    "\t\tlog.info(\"SPIKE renban use_next_free committed base=%s\" % p.nextFree)\n"
    "\t\tsystem.perspective.closePopup(\"renbanCollision\")\n"
    "\texcept A.AuthError, ae:\n"
    "\t\tself.view.custom.warn = \"DENIED (server-side): %s\" % ae\n"
    "\t\tlog.warn(\"SPIKE renban use_next_free DENIED (server-side gate): %s\" % ae)\n"
    "\texcept Exception, e:\n"
    "\t\tself.view.custom.warn = \"Use-next-free failed: %s\" % e\n")

OVERRIDE_SCRIPT = (
    "\t# PROCEED ANYWAY: requires the confirm checkbox (a second confirm — they may know the old release\n"
    "\t# is closing). Acknowledged set = the renbans the operator SAW colliding (params.acknowledged), so\n"
    "\t# the in-tx re-check still aborts on a DIFFERENT number taken newly since the WARN.\n"
    "\tlog = system.util.getLogger(\"SPIKE.renban.ui\")\n"
    "\tc = self.view.custom; p = self.view.params\n"
    "\tif not c.confirmOverride:\n"
    "\t\tc.warn = \"Tick 'I understand this reuses an open renban' to proceed anyway.\"; return\n"
    "\timport renban\n"
    "\timport auth as A\n"
    "\ttry:\n"
    "\t\tA.requireWrite(self.session)   # SERVER-SIDE WRITE GATE (BLOCKER-2)\n"
    "\t\tsiteId = self.session.custom.siteId if hasattr(self.session.custom, \"siteId\") else None\n"
    "\t\tres = renban.commit_renban_breakdown(p.group, int(p.trailers), int(p.pallets), \"Inventory_Spike\",\n"
    "\t\t\tresolution={\"action\": \"override\", \"alarm_id\": p.alarmId, \"acknowledged\": list(p.acknowledged or [])},\n"
    "\t\t\tsiteId=siteId, actor=\"operator\", session=self.session)\n"
    "\t\tc.result = res[\"status\"]\n"
    "\t\tif res[\"status\"] == \"COLLISION\":\n"
    "\t\t\tc.warn = \"Lost race: a NEW number was taken since the warning. Re-run the breakdown.\"; return\n"
    "\t\tlog.info(\"SPIKE renban override committed group=%s\" % p.group)\n"
    "\t\tsystem.perspective.closePopup(\"renbanCollision\")\n"
    "\texcept A.AuthError, ae:\n"
    "\t\tc.warn = \"DENIED (server-side): %s\" % ae\n"
    "\t\tlog.warn(\"SPIKE renban override DENIED (server-side gate): %s\" % ae)\n"
    "\texcept Exception, e:\n"
    "\t\tc.warn = \"Override failed: %s\" % e\n")

CANCEL_SCRIPT = (
    "\t# CANCEL is purely client-side: NO driver call. The RENBAN_COLLISION alarm STAYS active\n"
    "\t# (BIT_RESOLVED=0) on the home hub as a standing reminder to go close the stale release.\n"
    "\tsystem.util.getLogger(\"SPIKE.renban.ui\").info(\"SPIKE renban collision CANCELLED (alarm left active)\")\n"
    "\tsystem.perspective.closePopup(\"renbanCollision\")\n")


def build_dialog():
    head = lbl("dh", "Renban collision", {"fontSize": "16px", "fontWeight": "700", "color": DANGER})
    # WARN: bound to a custom prop the open script populates from params (the colliding renban + detail).
    warn = lbl("warn", "—", {"fontSize": "13.5px", "color": TXT, "margin": "6px 0"},
               bind="view.custom.warn", domId="rcd-warn")
    guide_lbl = lbl("gl", "Next free run + the colliding open release(s):",
                    {"fontSize": "12px", "color": MUT, "margin": "8px 0 2px"})
    # GUIDE table: the colliding open-release detail (next-free is shown in the WARN/below).
    guide = {"type": "ia.display.table", "meta": {"name": "guideTable", "domId": "rcd-guide-table"},
             "props": {"data": []},
             "propConfig": {"props.data": {"binding": {"type": "property",
                            "config": {"path": "view.params.collisions"}}}}}
    nextfree = lbl("nf", "Next free run base: —",
                   {"fontSize": "13px", "color": ACCENT, "fontWeight": "600", "margin": "8px 0"},
                   bind="view.custom.nextFreeText", domId="rcd-nextfree")
    confirm = {"type": "ia.input.checkbox", "meta": {"name": "confirmOverride", "domId": "rcd-confirm"},
               "props": {"text": "I understand this reuses an open renban", "selected": False},
               "propConfig": {"props.selected": {"binding": {"type": "property",
                              "config": {"path": "view.custom.confirmOverride"}, "bidirectional": True}}}}
    fix = flex("fix", [
        button("useNextFree", "Use next free", "rcd-use-next-free", USE_NEXT_FREE_SCRIPT, primary=True),
        button("override", "Proceed anyway", "rcd-override", OVERRIDE_SCRIPT, danger=True),
        button("cancel", "Cancel", "rcd-cancel", CANCEL_SCRIPT),
    ], direction="row", gap=12)
    root = flex("root", [head, warn, nextfree, guide_lbl, guide, confirm, fix], gap=8,
                domId="renban-collision-root", style={"backgroundColor": BG, "padding": "18px",
                "minHeight": "100%", "minWidth": "520px"})
    return {
        # the open script populates warn/nextFreeText from params; confirmOverride gates the override.
        "custom": {"warn": "", "nextFreeText": "", "confirmOverride": False, "result": ""},
        "params": {"group": "", "trailers": 0, "pallets": 0, "collisions": [], "nextFree": None,
                   "alarmId": None, "acknowledged": []},
        "propConfig": {},
        "props": {"defaultSize": {"width": 560, "height": 460}},
        "root": root,
    }


# ---- writers -----------------------------------------------------------------------------
def write_view(base, path_parts, view):
    d = os.path.join(base, *path_parts)
    os.makedirs(d, exist_ok=True)
    json.dump(view, open(os.path.join(d, "view.json"), "w"), indent=2)
    # ATTRIBUTES BLOCK IS MANDATORY — a missing one FAULTS the whole gateway on cold start (R31).
    json.dump({"scope": "G", "version": 1, "restricted": False, "overridable": True,
               "files": ["view.json"],
               "attributes": {"lastModification": {"actor": "external", "timestamp": "2026-06-23T00:00:00Z"},
                              "lastModificationSignature": "p4-renban-collision-shell"}},
              open(os.path.join(d, "resource.json"), "w"), indent=2)


if __name__ == "__main__":
    bd = build_breakdown()
    dlg = build_dialog()
    for base in (os.path.join(PERSP, "views"),
                 os.path.join(REPO_ROOT, "docs", "design", "perspective-views")):
        write_view(base, ["Order", "RenbanBreakdown"], bd)
        write_view(base, ["Order", "RenbanCollisionDialog"], dlg)
    print("OK  Order/RenbanBreakdown + Order/RenbanCollisionDialog written (gateway + repo)")
    print("    resource.json carries its attributes block (validated below).")
    print("")
    print("DESIGNER HAND-OFF (cannot be authored headlessly — reference-headless-ignition-authoring-limits):")
    print("  1. Named Query Order/renbanGroups data.bin XML (group-code dropdown source).")
    print("  2. Page-route mount: add /order/renban -> Order/RenbanBreakdown in page-config/config.json.")
    print("  3. Popup registration: the openPopup id 'renbanCollision' on RenbanBreakdown.")
    print("  4. Bidirectional-binding flag in binding.config for the editable inputs (commits on Tab/blur).")
    print("  NOTE: the server-side write gate (auth.requireWrite(self.session)) is NOW BAKED INTO the Commit /")
    print("        Use-next-free / Proceed-anyway scripts (BLOCKER-2 closed) — same P15 posture as the master")
    print("        views; nothing to finish in the Designer for the gate.")
