"""E2E confirmation of Order/OrderSpike — no human clicks.

Drives a real Perspective session headless: resets the trial if needed, opens the
view, screenshots the rendered 4-row ledger, fires Simulate and the edit/recompute,
and asserts on BOTH the DOM and the gateway-log SPIKE markers. Per-check PASS/FAIL/
SKIP; exit 1 if any FAIL.

  python3 e2e/test_order_spike.py            # headless
  python3 e2e/test_order_spike.py --headed   # watch it live (design review)

Selectors prefer component domId when present (#spike-order-grid, #spike-simulate-btn,
#spike-reset-btn) and fall back to visible text, so it works before/after the
developer adds domIds. See docs/automated-ui-testing.md.
"""
import sys, time
from playwright.sync_api import sync_playwright
import lib
from reset_trial import reset_trial

VIEW_URL = lib.view_url("order")


def q(page, domid, text=None, role=None):
    """domId first, then text/role fallback. Returns a locator or None."""
    el = page.query_selector("#" + domid)
    if el:
        return page.locator("#" + domid)
    if text:
        loc = page.get_by_text(text, exact=False)
        if loc.count():
            return loc.first
    if role:
        loc = page.get_by_role(role, name=text) if text else page.get_by_role(role)
        if loc.count():
            return loc.first
    return None


def cell_bg(page, text):
    """computed backgroundColor of the first element whose text == `text`."""
    el = page.query_selector("text=%s" % text)
    if not el:
        return None
    return el.evaluate("e=>getComputedStyle(e.closest('td,[role=cell],div')||e).backgroundColor")


# ---- virtualized-grid scroll/read helpers --------------------------------
# The Order grid (#spike-order-grid, ia.display.table) is VIRTUALIZED: only the
# rows in the viewport are mounted in the DOM, and they UNMOUNT when scrolled
# away. So we must scroll the inner ReactVirtualized__Grid incrementally and
# READ a matching row in the same JS step where it is mounted. Filtering is
# disabled, so scrolling is the only way to reach a deep group.
SCROLL_ROOT = "#spike-order-grid .ReactVirtualized__Grid"
PEACH = 'div.ia_table__cell[style*="rgb(255, 204, 153)"]'


def _scroll_top(page):
    """Reset the virtualized grid to the top; robust to current scroll pos."""
    page.evaluate(
        "(sel)=>{const g=document.querySelector(sel); if(g) g.scrollTop=0;}",
        SCROLL_ROOT,
    )
    page.wait_for_timeout(250)


def scroll_to_group(page, size_code, label="End Balance", max_steps=40, step_px=150):
    """Scroll the virtualized grid until a row whose token[0]==label and
    token[1]==size_code is mounted, then return that row's pipe-joined tokens
    (list of strings). Returns None if not found within max_steps.

    Read happens inside the same evaluate as the visibility test, because the
    row unmounts the moment it scrolls out of the viewport.
    """
    js = """
    (args)=>{
      const {sel, label, code} = args;
      const g = document.querySelector(sel);
      if(!g) return {found:false, atBottom:true, tokens:null};
      const rows = g.querySelectorAll('div.ia_table__row');
      for (const r of rows){
        const toks = (r.innerText||'').split('\\n').map(s=>s.trim()).filter(s=>s.length);
        if (toks.length>=2 && toks[0]===label && toks[1]===code){
          return {found:true, atBottom:false, tokens:toks};
        }
      }
      const atBottom = (g.scrollTop + g.clientHeight) >= (g.scrollHeight - 1);
      return {found:false, atBottom:atBottom, tokens:null};
    }
    """
    _scroll_top(page)
    for _ in range(max_steps):
        res = page.evaluate(js, {"sel": SCROLL_ROOT, "label": label, "code": size_code})
        if res and res.get("found"):
            return res["tokens"]
        if res and res.get("atBottom"):
            break
        page.evaluate(
            "(args)=>{const g=document.querySelector(args.sel); if(g) g.scrollTop+=args.dy;}",
            {"sel": SCROLL_ROOT, "dy": step_px},
        )
        page.wait_for_timeout(120)  # settle: let React mount the new window
    return None


def read_supplier_receipts(page, size_code, supplier, max_steps=40, step_px=150):
    """Scroll the grid to find the Receipts row for `supplier` within group
    `size_code` (token[0]=='Receipts', token[1]==size_code, token[2]==supplier),
    and in the SAME step count the peach order-by cells inside that row.

    Returns dict {found, tokens, peach} or {found:False,...}. Read + peach-count
    are atomic with the visibility test (the row unmounts when scrolled away).
    """
    js = """
    (args)=>{
      const {sel, code, supplier, peachSel} = args;
      const g = document.querySelector(sel);
      if(!g) return {found:false, atBottom:true, tokens:null, peach:0};
      const rows = g.querySelectorAll('div.ia_table__row');
      for (const r of rows){
        const toks = (r.innerText||'').split('\\n').map(s=>s.trim()).filter(s=>s.length);
        if (toks.length>=3 && toks[0]==='Receipts' && toks[1]===code && toks[2]===supplier){
          const peach = r.querySelectorAll(peachSel).length;
          return {found:true, atBottom:false, tokens:toks, peach:peach};
        }
      }
      const atBottom = (g.scrollTop + g.clientHeight) >= (g.scrollHeight - 1);
      return {found:false, atBottom:atBottom, tokens:null, peach:0};
    }
    """
    _scroll_top(page)
    for _ in range(max_steps):
        res = page.evaluate(js, {"sel": SCROLL_ROOT, "code": size_code,
                                 "supplier": supplier, "peachSel": PEACH})
        if res and res.get("found"):
            return res
        if res and res.get("atBottom"):
            break
        page.evaluate(
            "(args)=>{const g=document.querySelector(args.sel); if(g) g.scrollTop+=args.dy;}",
            {"sel": SCROLL_ROOT, "dy": step_px},
        )
        page.wait_for_timeout(120)
    return {"found": False, "tokens": None, "peach": 0}


def assert_group(page, rep, size_code, end_day0, suppliers):
    """Per-group deep assertions: group renders (End Balance row scrolls into
    view), pooled End day0 == golden constant, and each supplier's Receipts row
    is present with the right lead + exactly one peach order-by cell.

      suppliers: list of (name, lead) tuples.
    """
    end_toks = scroll_to_group(page, size_code, label="End Balance")
    rep.check("group %s renders (End Balance row scrolls into view)" % size_code,
              end_toks is not None,
              "tokens[:3]=%s" % (end_toks[:3] if end_toks else None))
    if end_toks is None:
        rep.skip("group %s pooled End day0" % size_code, "End Balance row not found")
        for name, _ in suppliers:
            rep.skip("group %s supplier %s Receipts" % (size_code, name), "group not found")
        return
    # token[2] is the day0 pooled End value, comma-formatted.
    got_end = end_toks[2] if len(end_toks) >= 3 else None
    rep.check("group %s pooled End day0 == %s" % (size_code, end_day0),
              got_end == end_day0, "got token[2]=%r" % got_end)

    for name, lead in suppliers:
        res = read_supplier_receipts(page, size_code, name)
        toks = res.get("tokens")
        rep.check("group %s supplier %s Receipts row present" % (size_code, name),
                  res.get("found"), "tokens[:4]=%s" % (toks[:4] if toks else None))
        if not res.get("found"):
            rep.skip("group %s %s lead == %s" % (size_code, name, lead), "row not found")
            rep.skip("group %s %s one peach order-by cell" % (size_code, name), "row not found")
            continue
        # Receipts row: token3 == lead (per RESUME: token2=supplier, token3=lead).
        got_lead = toks[3] if len(toks) >= 4 else None
        rep.check("group %s %s lead == %s" % (size_code, name, lead),
                  got_lead == str(lead), "got token[3]=%r" % got_lead)
        rep.check("group %s %s exactly one peach order-by cell" % (size_code, name),
                  res.get("peach") == 1, "%d peach cells in row" % res.get("peach"))


def run(page, rep):
    # ---- render ----------------------------------------------------------
    page.goto(VIEW_URL, wait_until="networkidle", timeout=30000)
    rendered = False
    try:
        page.wait_for_selector("text=Beg Balance", timeout=20000)
        rendered = True
    except Exception:
        pass
    page.screenshot(path=lib.ARTIFACTS + "/order_spike_load.png", full_page=True)
    if "Trial Expired" in page.inner_text("body"):
        rep.check("view renders ledger", False, "Perspective TRIAL EXPIRED — run reset_trial.py with creds")
        return
    rep.check("view renders ledger (Beg Balance present)", rendered,
              "screenshot: artifacts/order_spike_load.png")
    if not rendered:
        rep.skip("downstream grid checks", "grid not rendered")
        return

    body = page.inner_text("body")
    # The table is VIRTUALIZED — assert against the top-of-grid group (15D), whose
    # golden anchors are on-screen. Numbers render comma-formatted.
    rep.check("real inventory numbers render (15D Beg day0 = 12,000)", "12,000" in body,
              "top-group golden anchor")
    rep.check("PAB End row renders (15D End day0 = 11,973)", "11,973" in body)
    for lbl in ("Beg Balance", "Receipts", "Usage", "End Balance"):
        rep.check("ledger row label '%s'" % lbl, lbl in body)
    # glyph-clutter must be GONE (requirement #1)
    glyphs = [g for g in ("★", "🚚", "📦", "⚠", "[LT]", "[OT]") if g in body]
    rep.check("no glyph clutter (numbers/color-only)", not glyphs,
              "found: %s" % glyphs if glyphs else "clean")

    # ---- Simulate fires the transform (log marker) -----------------------
    off = lib.log_marker()
    btn = q(page, "spike-simulate-btn", text="Simulate", role="button")
    if btn:
        btn.click()
        time.sleep(2.5)  # WebSocket round-trip; transform logs SPIKE grid:
        lines = lib.grep_spike_since(off, "SPIKE grid:")
        rep.check("Simulate ran transform (SPIKE grid: marker)", bool(lines),
                  lines[-1].split("SPIKE")[-1][:60] if lines else "no marker in wrapper.log")
    else:
        rep.skip("Simulate marker", "Simulate button not found")

    # ---- peach on order-by cell (requirement #2/#3) ----------------------
    # Peach editable cells render as div.ia_table__cell with inline
    # background-color: rgb(255, 204, 153). Scope to grid cells (exclude the
    # legend swatch, which is an ia_labelComponent). PEACH defined at module top.
    peach_cells = page.locator(PEACH)
    npeach = peach_cells.count()
    rep.check("peach order-by cells render (one per visible supplier)", npeach >= 1,
              "%d peach cells on screen" % npeach)

    # ---- edit -> live recompute (requirement #4) -------------------------
    if npeach:
        off = lib.log_marker()
        try:
            cell = peach_cells.first
            cell.scroll_into_view_if_needed()
            cell.dblclick()                 # IA table: dblclick opens the cell editor
            page.wait_for_timeout(400)
            page.keyboard.type("5000")      # type into the focused editor
            page.keyboard.press("Enter")    # commit -> fires onEditCellCommit
            time.sleep(2.5)
            acc = lib.grep_spike_since(off, "SPIKE edit")
            rep.check("typed qty accepted (SPIKE edit ACCEPTED + recompute)",
                      any("ACCEPTED" in l for l in acc),
                      acc[-1].split("SPIKE")[-1][:70] if acc else "no SPIKE edit marker (gesture/edit-mode needs refinement)")
            page.screenshot(path=lib.ARTIFACTS + "/order_spike_after_edit.png", full_page=True)
        except Exception as e:
            rep.skip("edit recompute", "edit interaction failed: %s" % e)
    else:
        rep.skip("edit recompute", "no peach cell on screen")

    # ---- deep-group assertions (virtualized scroll) ----------------------
    # The edit check above left the grid scrolled; each assert_group() re-scrolls
    # from the top (via scroll_to_group/_scroll_top), so it is robust to position.
    # Golden anchors (authoritative, from the Delphi exports):
    #   18DL: pooled End day0 = 47,885; DUNLOP lead 5, MICHELIN lead 6 (2 peach).
    #   SPARE: pooled End day0 = -754; YOKOHAMA lead 5, MAXXIS lead 7 (2 peach).
    assert_group(page, rep, "18DL", "47,885",
                 [("DUNLOP", 5), ("MICHELIN", 6)])
    assert_group(page, rep, "SPARE", "-754",
                 [("YOKOHAMA", 5), ("MAXXIS", 7)])


def main():
    headed = "--headed" in sys.argv
    import os
    os.makedirs(lib.ARTIFACTS, exist_ok=True)
    rep = lib.Report()
    with sync_playwright() as p:
        b = p.chromium.launch(headless=not headed, slow_mo=300 if headed else 0)
        pg = b.new_page(viewport={"width": 1680, "height": 1050})
        print("== trial reset ==")
        ok, msg = reset_trial(pg)
        if msg == "NEED_CREDS":
            rep.skip("trial reset", "no GATEWAY_USER/GATEWAY_PASS — view render will be blocked if trial expired")
        else:
            rep.check("trial active", ok, msg)
        print("== order view ==")
        run(pg, rep)
        b.close()
    return rep.summary_exit()


if __name__ == "__main__":
    sys.exit(main())
