# Order spike — exported Perspective view resources

On-disk export of the `spike` Ignition project's Perspective view(s) for the Order
keystone, so the gateway-only resources are captured in git (RESUME-order-spike.md
NEXT STEP 4). The dev gateway is throwaway; without this, a rebuild loses the view.

## `Order/OrderSpike/`
The faithful pooled 4-row ledger view (Beg / Receipts-per-supplier / Usage / End +
gray separator; numbers + color only; peach editable order-by cell; red below-safety
End; client-side live End recompute). Source of truth for layout/behavior is
`../option-a.md` + `../build-spec.md`; the data contract is `dbo.SIM_OrderSimulation`
(`../SIM_OrderSimulation.sql`).

- `view.json` — the Perspective view (the `PhasedGrid` builder transform calls
  `SIM_OrderSimulation` once per section A/B/C via `system.db.runPrepQuery`). **Carries
  the R1 fix**: the binding passes `@UseFirstProductionDay=1` (the golden client config;
  with `=0` the forecast week-offset is dropped and tire/FILM regress — see `../sc1-parity-results.md`).
- `resource.json` — Ignition resource manifest (scope G). The `lastModificationSignature`
  is gateway-specific and will differ after any Designer edit; it is captured as-is.

### Redeploy to a gateway
Copy back into the project and reload:
```
cp -r Order /usr/local/ignition/data/projects/spike/com.inductiveautomation.perspective/views/
/usr/local/ignition/gwcmd.sh -r
```
Then confirm with the E2E harness: `python3 ../../../../scripts/e2e/test_order_spike.py` (28/28).

### Provenance
Exported 2026-06-15 from gateway path
`data/projects/spike/com.inductiveautomation.perspective/views/Order/OrderSpike/`
(Ignition 8.1.52). Re-export after Designer edits to keep git in sync (the view is still
edited in the Designer, not from git — this is a snapshot, not a two-way sync).
