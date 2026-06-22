# hotcall/message-handler.py — the GATEWAY MESSAGE HANDLER body for the "create_hotcall_asn" message the
# Hot-Call Entry view's "Create Hot-Call ASN" button posts (via system.util.sendRequest).
#
# WHY a gateway message handler (not an inline session script): the create_hotcall_asn DB write must run
# in GATEWAY scope (the transaction + the DB connection live on the gateway), and the site id must derive
# from the gateway/session context, NEVER a client param (multi-site D1). sendRequest blocks for and
# returns this handler's value, so the view can show the created ASN id / qty.
#
# HOW TO BIND IN THE DESIGNER (the one Designer-finish step for the handler):
#   Project Properties > Scripting > Gateway Event Scripts > Message Handlers > Add
#     Message Type: create_hotcall_asn
#     Threading:    Dedicated (a DB transaction; keep it off the shared thread)
#     Script:       paste the body below (or `import hotcall` and call hotcall.create_hotcall_asn — once
#                   project-library/hotcall/code.py is promoted to a Project Library script `hotcall`).
#
# The handler signature is the standard Ignition message-handler one: def handleMessage(payload).
# `payload` is the dict the view posts: {"line", "prodDate", "manifest", "items":[{"part","qty"},...]}.

import hotcall   # the Project Library module (project-library/hotcall/code.py) — promoted in the Designer.


def handleMessage(payload):
    """Invoke the REAL create_hotcall_asn driver in gateway scope. Returns the driver's result dict
    (asnId / details / qty / manifest / line / prodDate) — sendRequest hands it back to the view.

    Site id derives from the gateway context (the session's site), NEVER the payload (multi-site D1).
    For the single-site spike this is site 1; at M4 the handler resolves the session's IN_SITE_ID."""
    log = system.util.getLogger("SPIKE.hotcall.handler")

    line = payload["line"]
    prodDate = payload["prodDate"]
    manifest = payload["manifest"]
    items = [(it["part"], it["qty"]) for it in (payload.get("items") or [])]

    # site derives from the session/gateway, not the client payload (D1). Spike = 1.
    site = 1

    log.info("create_hotcall_asn message: line=%s prodDate=%s manifest=%s items=%d site=%d"
             % (line, prodDate, manifest, len(items), site))

    # HotcallError (validation) and any DB error propagate back to the view as the sendRequest exception,
    # which the button's onActionPerformed catches and shows in view.custom.statusMsg.
    result = hotcall.create_hotcall_asn(line=line, prodDate=prodDate, manifest=manifest,
                                        items=items, site=site, database="Inventory_Spike")
    log.info("create_hotcall_asn message -> ASN %s, qty %s" % (result["asnId"], result["qty"]))
    return result
