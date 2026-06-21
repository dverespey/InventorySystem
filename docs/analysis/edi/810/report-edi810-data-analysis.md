# REPORT_EDI810 — Data-Feed Analysis (M1 Rank-4 outbound 810 invoice)

What `REPORT_EDI810` reads to populate the 810 — incl. the **price** (the 810 carries it; the 856 doesn't).
Rebuild reads via a side-effect-free SELECT (NOT a wrap; the proc self-flips). Proven on live `Inventory`
(mssql-spike, RO/rolled-back). Sources: D6 `spike-report-procs-d6.sql:60-92`; legacy `REPORT_EDI810`
(live == `/tmp/inv_utf8.sql:3734-3776`, self-flip `:3775`, window-blind `:3749`); `fn_ManifestCostAt`
(`spike-manifest-cost-lookup.sql`); consumer `EDI810Object.pas`; `priceless-lines-diagnostic.sql`;
`test_report_procs_d6.py`.

## SELECT shape — TWO branches
- **`@EIN=0` (CREATE preview, 6 cols, NO mutation):** `INV_ASN_MST a JOIN INV_ASN_DETAIL_MST d ON
  a.IN_ASN_ID=d.IN_ASN_ID AND a.VC_ASN_STATUS='A' AND d.IN_INV_ID IS NULL` (the unbilled flag is on the
  DETAIL) `CROSS APPLY fn_ManifestCostAt(d.VC_ASSY_PART_NUMBER, a.VC_PRODUCTION_DATE) m ORDER BY
  d.VC_MANIFEST_NUMBER`. This is the daily-invoicing read (unbilled shipments → a new invoice).
- **`@EIN<>0` (RECREATE, 7 cols, +`JOIN INV_INV_MST iim`, SELF-FLIP):** `WHERE iim.IN_INV_EIN=@EIN` then
  `UPDATE INV_INV_MST SET VC_INV_STATUS='S' WHERE IN_INV_EIN=@EIN`.
- Projection → segments: Manifest→REF/MK + IT101 M391-vs-M390 (manifest `'7'` prefix); part→IT1 PN;
  **UnitPrice = `m.MO_PRICE` (money 4-dec) → IT104 unit price + the line amount + the TDS total**; qty=`IN_QTY`
  →IT102; PickUpDate=`VC_PRODUCTION_DATE`→DTM/filename/new-file boundary. One row per (header × surviving
  detail); **no GROUP BY** (unlike the 856); CROSS APPLY TOP 1 prevents fan-out on a clean cost master.

## D6 window-aware pricing — the decided divergence (OVER-BILL PROVEN on real data)
Legacy joins the cost master on **part code ALONE, no date window** → arbitrary/duplicate price. D6 →
`CROSS APPLY fn_ManifestCostAt` (inclusive `<=`/`>=`, TOP-1 newest-start). **The rebuild uses D6.** Proof:
- **EIN 5692 (REAL):** legacy **$61,783.75** (2 lines — prices a 2020 shipment with a 2022-onward window) vs
  D6 **$58,093.75** (1 line) → legacy **over-billed $3,690.00**; D6 drops the wrong-window line.
- **Synthesized 2nd window:** legacy doubles to **$128,878.75** vs D6 **$70,785.00** (newest-start wins).
`test_report_procs_d6.py` covers the 810 (the headline over-bill test). This is an accepted billing-correctness
divergence (same decision family as D6).

## Priceless lines (CROSS APPLY drops un-priced → under-bill)
Decision: keep CROSS APPLY (inner) + run the pre-invoice diagnostic. Live: **25,351 / 39,707 (64%) of billed
lines have no covering window** (sparse snapshot: 45 cost rows / 128 distinct billed parts) → would drop. The
`@EIN=0` create-feed priceless count is 0 only because the create feed is empty here (all detail already
billed). Rebuild: keep inner; run `priceless-lines-diagnostic.sql` (EDI810 branch) before every run; expect 0.

## Self-flip — do NOT wrap (`:3775`)
`UPDATE INV_INV_MST SET VC_INV_STATUS='S' WHERE IN_INV_EIN=@EIN` (810 twin of the 856 `:3695`). Also AVOID
**`UPDATE_INVRecreate`** = blanket `UPDATE INV_INV_MST SET VC_INV_STATUS='S' WHERE VC_INV_STATUS='C'` (the
invoice twin of the `UPDATE_ASNStatus` hazard). Create path doesn't self-flip in SQL (`INSERT_INVInfo` inserts
status `'S'`, `UPDATE_INVItems` links detail — deliberate app writes). Rebuild: pure SELECT, flip per-invoice.

## EIN
`IN_INV_EIN` int NOT NULL, unique-per-invoice (live 29–9058, all distinct). Allocated at invoice-CREATE
(`MainMenu.CreateINVOICEClick` → read SiteEIN → EIN=SiteEIN+1 → `AD_UpdateEIN` bumps the per-site counter →
`INSERT_INVInfo @EIN`). Reused at recreate. NO hardcoded literal. **Rebuild: allocate per-site from
`INV_SITES.IN_EIN_SEQ` at invoice-create; reuse at recreate; filter the read by it.** Legacy shares ONE per-site
counter for BOTH 856 and 810 → keep the SHARED `INV_SITES.IN_EIN_SEQ` (faithful; interleaved control numbers).

## Unsend (Carry 5) — legacy HARD-DELETEs; rebuild in-place
`UPDATE_INVUnsend @INVid` (`:3387`): `UPDATE INV_ASN_DETAIL_MST SET IN_INV_ID=null …` (re-pool) then
`DELETE FROM INV_INV_MST …` (destroys header+EIN+audit; a commented-out line shows the original intent was a
status revert). **Rebuild (Carry 5): in-place status revert (keep header+EIN+audit) + a "recreate the 810 file"
flag + recompute costs via fn_ManifestCostAt + re-pool detail — NOT hard-delete.**

## UPDATE_EINStatus — invoice ack (997/824), site-scoping gap
`@EINType<>'SH'` branch: `UPDATE INV_INV_MST SET VC_INV_STATUS=@EINStatus WHERE IN_INV_EIN=@EIN` (no site
filter). Under multi-site EIN reuse, a cross-site ack collides. Rebuild: scope by `site_id` (or a globally-
unique invoice id).

## What the rebuild's 810 read MUST reproduce + traps
The 6-col create read (`@EIN=0`, unbilled `IN_INV_ID IS NULL`, no mutation) + the 7-col recreate read
(`@EIN<>0`, EXCLUDE the self-flip); **D6 window-aware `MO_PRICE`**; line amt = price×qty; TDS = Σ (DECISION
810-1 on the format/bug); EIN per-site at invoice-create, reused at recreate; in-place unsend (Carry 5);
site-scoped ack. Traps: the self-flip + `UPDATE_INVRecreate` (mutate-on-read); window-blind pricing (the
over-bill); priceless drop (under-bill — run the diagnostic); **UnitPrice IS emitted** (a missing price drops
the whole line); the TDS string-format bug (DECISION 810-1). Snapshot edges: all detail billed → the `@EIN=0`
create feed returns 0 here (test with synthetic unbilled lines).
