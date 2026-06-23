# Adversary findings — M4 P15 master-write-gate sweep (branch `m4-master-write-gates`)

Reviewer verdict on whether the P15 systemic hole (7 master CRUD views writing with NO server-side gate)
is now fully closed. Methodology: static enumeration of EVERY script in each deployed view (not just the
3 the injector targets), gate-position trace, the live e2e suite, and a DB read against the live spike.

## Per-view write-path coverage

Across all 7 views the component structure is IDENTICAL and minimal: `SearchButton`, `NewButton`,
`<Name>Grid` (onRowClick), `SaveButton`, `DeleteButton`, `ClearButton`, plus the view-level
`custom.recordId.onChange` detail-load. I scanned every one of those scripts for a REAL write primitive
(`runPrepUpdate(...)` OR a `runPrepQuery` carrying `INSERT INTO` / `DELETE FROM`) vs. the gate marker
`A.requireWrite(self.session)`.

The ONLY scripts that reach a real write primitive in every view are **SaveButton** (INSERT via
SCOPE_IDENTITY + UPDATE via runPrepUpdate) and **DeleteButton** (DELETE via runPrepUpdate). Both are
GATED, in every view. NewButton is a pure form-reset (no write) but carries the gate anyway. Search,
onRowClick, Clear, and recordId.onChange are read/prop-only — confirmed no write primitive.

| View | SaveButton | DeleteButton | NewButton | Search / onRowClick / Clear / recordId.onChange |
|---|---|---|---|---|
| Size           | WRITE → GATED | WRITE → GATED | no-write (gated) | no write primitive |
| Supplier       | WRITE → GATED | WRITE → GATED | no-write (gated) | no write primitive |
| PartsStock     | WRITE → GATED | WRITE → GATED | no-write (gated) | no write primitive |
| ManifestCost   | WRITE → GATED | WRITE → GATED | no-write (gated) | no write primitive |
| RenbanGroup    | WRITE → GATED | WRITE → GATED | no-write (gated) | no write primitive |
| AssemblyDetail | WRITE → GATED | WRITE → GATED | no-write (gated) | no write primitive |
| Logistics      | WRITE → GATED | WRITE → GATED | no-write (gated) | no write primitive |

No inline grid-edit / onEditCellCommit / row-action / secondary write button exists in any of the 7 —
these are list+detail views; the only DB mutations are Save and Delete, both gated. The injector's
worry-case (a missed write path) does not exist in these views because there are no write paths beyond
Save/Delete.

## Gate is genuinely server-side and runs BEFORE the write — VERDICTS

- **SOUND — gate position.** Spot-checked PartsStock / ManifestCost / AssemblyDetail Save scripts: the
  `try: A.requireWrite(self.session) except A.AuthError, e: ... return` block sits at top-level indent
  (single tab), immediately after `c = self.view.custom`, and strictly precedes the first write
  (PartsStock gate line 15 vs first-write line 83; ManifestCost 17 vs 79; AssemblyDetail 15 vs 62). The
  deny path `return`s before any validation or `system.db` call. Same injected block in Delete (gate
  before the refCount read and the DELETE).
- **SOUND — server-side, not client-trust.** `auth.requireWrite(session)` →
  `sessionRoles(session)` reads `session.props.auth.user.roles` (gateway-populated, not client-writable)
  → `authorizeAny(roles, (ProductionControl, Admin))`, raising `AuthError` on deny
  (`project-library/auth/code.py:102-117, 135-156`). The decision never reads `view.custom.mayEdit` or any
  client prop. A forged `mayEdit=true` and an anonymous (roles=[]) session are both rejected — proven by
  the e2e suite (forged `mayEdit=true` + anon/viewer → `write=False`, statusMsg "DENIED") and by the
  "SESSION not prop" case (ProductionControl with `mayEdit=FALSE` still writes).
- **SOUND — no qaAdmin / client-only hatch in any write path.** `test_no_client_trust_in_write_path`
  asserts `qaAdmin` absent + session-gated on Save+Delete for all 7 (14 PASS). The qaAdmin URL hatch lives
  only in the Sites `custom.mayEdit` UI-visibility binding (per gen_sites_view.py), which the write gate
  never reads; it does not exist in these 7 views' scripts at all.
- **SOUND — revert-proof.** Neutering `auth.requireWrite` to a no-op makes the forged-prop anon SAVE slip
  through to the write on ALL 7 (test §2), proving the gate (not the harness) is load-bearing.

## Idempotent + committed==deployed — SOUND

- `gen_master_write_gates.py --check` → all 7 `Save=GATED, Delete=GATED, New=GATED`, exit 0.
- Re-running the injector in inject mode reports `gated=[] skipped(already)=[Save,Delete,New]` / "no
  changes"; file md5s are byte-identical before/after → truly idempotent (skips on `A.requireWrite(`
  marker, and the deterministic `json.dump(sort_keys=True)` re-serialize produces identical bytes).
- Committed repo view.json == deployed gateway view.json for all 7 (byte-identical `diff`), so the
  reviewable artifact equals the runtime.
- e2e: `test_master_write_gates.py` → **86 PASS / 0 FAIL / 0 SKIP**.

## PartsStock test failures — READ-side + pre-existing (confirmed)

`test_partsstock_crud.py` → 16 PASS / 6 FAIL / 3 SKIP. All 6 failures are in the detail-LOAD read path
(`open_detail_via_row`, driven by the `recordId.onChange` SELECT) or a cross-DB SELECT — NONE touch
Save/Delete/New (the only scripts the injector modified):

1. `Cross-DB VehicleOrder.dbo.LINE readable (CAMRY..TUNDRA)` (test line 554) — `VehicleOrder.dbo.LINE`
   holds exactly ONE row (COROLLA, id 1) on the live spike; the test expects 5. Pure read / seed-data gap.
2-6. `recordId onChange loaded the row` / `5 FK combos populated` / `part-number field populated` /
   `IN_QTY read-only shows 28133` / `BIT_LOT_SIZE_ORDERS inverted` (lines 251, 256, 266, 291, 297) — all
   downstream of the detail-load SELECT + combo population, which depend on the same single-row Line table
   (the combos check requires `line=5`, but only line id 1 exists). The anchored part `4261102Q5100` IS
   present (DB query: anchored_present=1; 47 parts total, matching the test's own "restored to 47 parts"
   teardown), so the row is not missing — the read-side rendering/cross-DB seed is the cause.

The injector explicitly left `recordId.onChange` untouched (read-only SELECT; injector docstring scope
note) — confirmed UNGATED + write-free in my scan. Because the injection only edited Save/Delete/New,
none of which is on the read/detail-load path, it CANNOT have introduced these. "Restore the ungated view
reproduces them" is structurally sound: reverting the gate leaves the read path byte-identical. Verified
read-side and pre-existing — not a regression from the sweep.

## END VERDICT — SOUND. P15 systemic hole CLOSED across all 7 views.

Every DB-write path in all 7 views (Save + Delete; New carries the gate defensively) is server-side
gated via `auth.requireWrite(self.session)`, which runs before the write, keys off the gateway-populated
session roles, and rejects forged props and anon sessions. There is NO additional write path (no inline
grid-edit / cell-commit / secondary write button) in these list+detail views, so no path was skipped by
the injector's Save/Delete/New targeting. Committed==deployed, idempotent, 86/0/0 e2e. The 6 PartsStock
failures are read-side cross-DB/seed gaps, pre-existing, and outside the swept write paths. The
cutover-blocker for the master-write-gate hole is cleared. (Residual, NOT a blocker: the PartsStock Line
dropdown / detail-load needs the cross-DB VehicleOrder.dbo.LINE seed populated beyond COROLLA for that
view's read path to render fully — bounce to the data-seed / cutover owner, separate from auth.)
