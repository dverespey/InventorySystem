# Automated UI confirmation for Ignition Perspective (no manual clicking)

> **STATUS: PROVEN (2026-06-15).** `scripts/e2e/test_order_spike.py` runs **12/12
> PASS, repeatable, zero human clicks** against `Order/OrderSpike`: auto-resets the
> trial, renders the 4-row ledger, asserts numbers/labels/no-glyphs/peach cells,
> fires Simulate + a live cell edit, and checks the `SPIKE` log markers — capturing
> `artifacts/order_spike_load.png` + `_after_edit.png`. It resolved the
> code-reviewer's two open runtime RISKs (transform-never-ran; `onEditCellCommit`
> qualified-value access — log: `edit ACCEPTED: 4265202R6000|5 = …`).
>
> Concrete findings baked into the harness: IdP login is **two-step** (username →
> Enter → `input[name=password]` reveals; ignore the unnamed decoy password
> inputs). Peach editable cells = `div.ia_table__cell` with inline
> `background-color: rgb(255, 204, 153)`. Editing a cell = **double-click → type →
> Enter** (fires `onEditCellCommit`). Numbers render **comma-formatted** (`12,000`).
> The table is **virtualized** → assert against the top-visible group (15D) or
> scroll/filter to reach deeper rows (e.g. 18DL); that scroll-to-row is the next
> enhancement. `domId`s (`#spike-order-grid`, `#spike-simulate-btn`, `#spike-reset-btn`,
> `#spike-parttype`, `#spike-line`) are **added and loaded** — a gateway restart
> activated them (restart is now permitted via `.claude/settings.json`); the harness
> prefers `domId` and falls back to text/structure. The gateway does NOT auto-scan
> external file edits, so adding more `domId`s later needs another restart (or a
> Designer save) to take effect.


Goal: stop development waiting on a human to open a browser and click. Drive the
Perspective session programmatically — fire the bindings/transforms, assert the
result, and capture **screenshots** — so the dev loop self-confirms functionality
and David reviews *design* from screenshots/recordings on his own schedule, not as
a blocking gate.

## Recommendation: Playwright (Python), headless by default

**Why Playwright over Selenium/Cypress** (2026 consensus for a modern React SPA,
which Perspective is): built-in **auto-waiting** (waits for elements to be ready —
critical because Perspective renders cells asynchronously over a WebSocket, so a
fixed `sleep` is flaky), first-class **headless + headed**, **screenshots**,
**video**, and a **trace viewer** (a full DOM-snapshot timeline you can scrub).
Selenium needs manual explicit-waits and extra config; Cypress is fine but
Playwright has better cross-context control and is the higher-momentum tool.

**Python**, not Node — the repo already does its data work in Python
(`openpyxl`, `parity_diff.py`). A Playwright-Python script can drive the browser
**and** cross-check the on-screen numbers against the proc/golden in one file.

- Install (native on the Intel Mac — no Colima needed; it hits `localhost:8088`):
  `pip install playwright && playwright install chromium` (~150 MB one-time).
- Default headless for the dev loop; `--headed --slowmo 400` when David wants to
  watch a run live; trace/video artifacts for after-the-fact review.

## The linchpin: `domId` makes components addressable

Perspective has an official meta property **`domId`** — docs verbatim: *"allows you
to set the DOM 'id' of the output element. This property is intended for testing
purposes only, such as using a framework like Selenium to test a page."* It is
hidden by default; add it under a component's **Meta** category.

→ Convention: give every component a test needs to touch a stable `domId`
(`spike-order-grid`, `spike-simulate-btn`, `spike-reset-btn`, …). Then Playwright
targets `#spike-order-grid` — robust against layout churn. **Do NOT** rely on
Perspective's internal CSS classes or the undocumented/unsupported JS API; `domId`
is the supported, stable hook. (Add this as a standing step in ignition-developer's
build: "assign domId to anything a test asserts on.")

## Dual-channel assertions (the strong part)

Combine two independent signals so a green run really means it works:

1. **DOM / computed style via Playwright** — read cell text values; read
   `getComputedStyle(cell).backgroundColor` to assert the **peach `#FFCC99`** lands
   only on order-by cells and **red `#FF0000`** only on below-safety End cells
   (this automates the *color-parity* check too, not just values); confirm
   non-peach cells are not editable.
2. **Gateway-log markers we already emit** — the view logs `SPIKE grid: A=.. B=..
   C=.. edits=N`, `SPIKE edit ACCEPTED/REJECTED`. Playwright drives the click that
   *triggers* them; the script then greps `/usr/local/ignition/logs/wrapper.log`.
   This is exactly the existing `live-session-test-plan.md`, with the human click
   replaced by `page.click('#spike-simulate-btn')`.

A run can also diff the rendered grid against `/tmp/golden/*.xlsx` — closing the
loop from "proc parity" → "what's actually painted on screen parity".

## Screenshots replace "I need to click to see it"

`page.screenshot(path=..., full_page=True)` (or element-scoped) after each step →
David reviews design + functionality from PNGs (and the trace viewer's scrubbable
timeline / video) asynchronously. He still eyeballs to sign off design; he just
isn't the one driving the clicks, and dev never blocks on his availability.

## Auth

The spike route `…/client/spike/order` returns **200 anonymously** (verified) — no
login step needed for spike work. (Production/authed views: Playwright logs in once
and reuses `storage_state` — a stored cookie/session — so auth is a one-time
fixture, not per-test friction.)

## Proposed shape

```
scripts/e2e/
  conftest.py            # Playwright fixtures: base_url, headless flag, artifacts dir
  test_order_spike.py    # open → wait #spike-order-grid → screenshot →
                         #   assert 18DL block (4 pooled rows + 2 receipts, values) →
                         #   assert peach@order-by / red@below-safety (computed style) →
                         #   click Simulate → grep SPIKE grid edits=0 →
                         #   type into DUNLOP peach cell → grep SPIKE edit ACCEPTED →
                         #     assert End recompute + red cleared →
                         #   type into a locked cell → assert non-editable
  artifacts/             # screenshots, video, trace.zip (gitignored)
```

Run by **ignition-qa** after each build; it reports pass/fail + attaches
screenshots. David is pinged only to approve design, from the images.

## Caveats / gotchas (Perspective-specific)

- **Async rendering:** always `wait_for_selector`/`expect` on a `domId` or text —
  never assert right after `goto`. Playwright's auto-wait handles most of it.
- **Undocumented JS API:** Perspective's in-browser component JS API is explicitly
  unsupported — avoid it; stay on DOM + `domId` + computed styles.
- **Version:** Playwright + Chromium run fine on 8.1.52 and macOS Intel (x86_64);
  it drives the browser, independent of the gateway version.
- **`gwcmd -r`** (resource reload) can be sandbox-blocked as "disruptive"; the
  harness should reload via the script runner or fold it into a setup step David
  approves, not assume it.
