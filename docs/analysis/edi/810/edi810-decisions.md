# EDI 810 — Locked Decisions (the intended divergences from legacy)

The 810 rebuild is **byte-faithful to the legacy WIRE FORMAT** but deliberately diverges from the legacy on
a set of **billing-correctness** points David has explicitly locked (same discipline as D6 + clean-money).
These are NOT silent — each changes invoice content/granularity/billed-state vs the legacy and is recorded
here. True TEMA byte-parity is still pending a golden 810 + the real `INV_SITES` site values (placeholders on
the spike) — confirm at cutover.

## D-810-1 — Clean / correct money (David 2026-06-20)
- **TDS01** = `round(total × 10000)` as an implied-decimal-4 integer — FIXES the legacy hand-rolled
  string-surgery bugs (1-digit-fraction off-by-10000×; whole-dollar malformed). The rebuild emits the
  CORRECT value; the unit test asserts correct + asserts the buggy value absent.
- **IT104** = fixed scale-4 decimal, locale-independent (NOT `FloatToStr`). Exact scale pending golden 810.

## D-810-2 — D6 window-aware pricing (David 2026-06-18, applied here)
The 810 prices each line via `fn_ManifestCostAt(part, VC_PRODUCTION_DATE)` (inclusive window, newest-start),
NOT the legacy window-blind assy-code-only join. PROVEN over-bill fix on real EIN 5692: legacy $61,783.75 →
D6 $58,093.75 (the $3,690 wrong-window line dropped). Priceless lines (no covering window) are dropped by the
inner CROSS APPLY — run `priceless-lines-diagnostic.sql` (EDI810 branch) before each run; expect 0.

## D-810-3 — Invoice grouping: ONE invoice per DISTINCT pickup date (David 2026-06-21)
The legacy splits invoices on **contiguous runs** of equal pickup-date in manifest order, so a date whose
manifests are NON-ADJACENT (the `5…`/`7…`/`T…` prefix families sort-disjoint — **203 of 2,319 real pickup
dates** carry >1 family) gets **multiple invoices/EINs/files** (double-filed). **The rebuild groups by
DISTINCT pickup date** → one invoice/EIN/file per date. Cleaner ("one day's shipments = one invoice"); avoids
the double-filing quirk. **Divergence: file/EIN count differs from legacy on those 203 dates.** Locked as
intended.

## D-810-4 — Detail-link scope: bill ALL ASNs in the date group (David 2026-06-21)
The legacy `UPDATE_INVItems` links only the **first ASN** of a run (`WHERE IN_ASN_ID=@ASNID`) → on a multi-ASN
run it emits all lines on the wire but BILLS only the first ASN, leaving the rest un-billed (re-billable next
run = an under-bill / double-bill hazard). **The rebuild links EVERY unbilled line in the invoice's date
group.** Correct; no dangling/re-billable lines. **Divergence: post-run billed state differs from legacy on
multi-ASN runs.** Locked as intended.

## D-810-5 — Create filename = `810<mmdd>.txt` (no LineName); the legacy create path crashes
The legacy `MainMenu.pas:2623` create-810 filename interpolates `EDI810DataSet.FieldByName('LineName')`, but
the `REPORT_EDI810 @EIN=0` feed returns NO LineName column (it's a copy-paste from the 856, whose feed DOES
have it) → `FieldByName` raises `EDatabaseError` → caught → "Unable to create INVOICE" → **the legacy
create-810-with-EDI path writes NO file** (the recreate path, `ASNInvoice.pas:872`, is the working one, and it
uses `810<mmdd>.txt` with no LineName). **The rebuild uses `810<mmdd>.txt` (no LineName) and WORKS** — both
faithful to the only filename the legacy can actually produce AND an improvement (the create path no longer
crashes). `mmdd` = chars 5-8 of `VC_PRODUCTION_DATE` (drops the year — a known legacy trait; cross-year
collision risk carried forward).

## Carried (decided earlier / mechanism, not divergence)
- **Unsend (Carry 5):** in-place status revert + recreate flag + re-pool detail + cost-recompute — NOT the
  legacy hard-delete of the invoice header.
- **EIN:** per-site `INV_SITES.IN_EIN_SEQ`, allocated at invoice-create (atomic, site-scoped), REUSED at
  recreate (so the 997-ack lands); shared 856/810 sequence (interleaved, faithful).
- **Self-flip / blanket `UPDATE_INVRecreate`:** never used; status flipped per-invoice at send.
- **`UPDATE_EINStatus` (invoice ack, 997/824):** needs `site_id` scoping at multi-site (carry).
