# EDI 856 ASN — Wire-Format Spec (`EDI856Object.T856EDI`)

Source-truth for the M1 Rank-2 outbound-856 builder. Decoded from `EDI856Object.pas` (the `T856EDI`
builder; LIVE per `InventorySystem.dpr:48`), drivers `ASNSelect.pas` (create) + `ASNInvoice.pas:792`
(recreate). Site identity from `AD_GetSite` on the **VehicleOrder** DB (`ALC_Connection`). The rebuild
builds this as **NEW gateway Jython** (porting the segment logic) — it does **NOT** wrap `REPORT_EDI856`
(which self-flips status; see `report-edi856-data-analysis.md`).

## Format conventions
- **Dates:** `fPickupDate` = `INV_ASN_MST.VC_PRODUCTION_DATE` = varchar(8) `yyyymmdd`. `f810Time = now` captured
  once per file → all `hhmm` fields share one timestamp.
- **EIN (`fein`):** formatted `%9.9d` (9 chars, zero-padded) EVERYWHERE it appears.
- Each segment is `Writeln` → trailing **CRLF**. `fSepSegment` is read but **never emitted** (legacy bug —
  Trap 1).

## Segment list (emission order) — `Execute:116-126`
ISA · GS · ST · BSN · DTM · HL(Shipment) · TD1 · TD5 · TD3 · [per manifest: HL(Order) · PRF] · [per detail:
HL(Item) · LIN · SN1] … · CTT · SE · GE · IEA.

### ISA (`:140-164`) — positional, `%-Ns` padded
`ISA*00*<10>*01*<DUNS %-10>*ZZ*<DUNS-Supplier %-15>*01*<TMMDUNS %-15>*<yymmdd>*<hhmm>*U*00400*<%9.9d EIN>*0*<EDIMode>*<subElemSep>`
- ISA09 = `copy(fPickupDate,3,6)` = **`yymmdd`** (century dropped) — UNLIKE GS04/BSN03/DTM02 which are full `yyyymmdd`.
- ISA13 = the EIN (interchange control #). ISA15 = `Site.SiteEDIMode` (`T`/`P`). ISA16 = sub-element sep.
- Fixed widths: ISA02/04 = 10, ISA06/08 = 15 → the rebuild MUST hard-truncate/validate (Trap 6).

### GS (`:174-189`)
`GS*SH*<DUNS>*<TMMDUNS>*<yyyymmdd>*<hhmm>*<%9.9d EIN>*X*004010` — GS06 = EIN (group control #).

### ST (`:199-208`)
`ST*856*<%9.9d EIN>` — ST02 = EIN (transaction-set control #), NOT a 1-based counter. `fSegCount`→1.

### BSN (`:219-231`)
`BSN*00*<yyyymmdd + %9.9d EIN (17 chars)>*<yyyymmdd>*<hhmm>` — BSN02 = shipment id = pickup-date + 9-digit EIN. `fSegCount`→2.

### DTM (`:242-254`)
`DTM*011*<yyyymmdd>*<hhmm>*ET` — shipped date/time, time-zone `ET` hardcoded. `fSegCount`→3.

### HL hierarchy (S → O → I)  `:265-378`
- **Shipment** (`:285`): `HL*1**S*1` (HL02 empty = no parent). Then `TD1**`, `TD5*B*25*00000*<SiteDeliveryMethodCode>`, `TD3*TL**1234567890` (**trailer id literal `1234567890`** — Trap 5).
- **Order** per distinct Manifest (`:320`): `HL*<id>*1*O*1` (HL02 hardcoded `1` → always parents the shipment). Then `PRF*<Manifest>-<Manifest>` (the `-` is data).
- **Item** per detail (`:338`): `HL*<id>*<parent>*I*0`. Then `LIN**BP*<PartNumber>*RC*<Kanban>*` (note the **TRAILING element separator** after the kanban — `.pas:352` appends `+fSepElement` after the kanban too, so the segment ends in one empty trailing element) and `SN1**<ShipQty>*PC` (SN1 has NO trailing sep — `.pas:358` ends `+'PC'`; qty = `IN_QTY`, no leading zeros, PC = pieces).
- All HL-loop segments buffered in `HLList`, flushed with `+1 fSegCount` each (`:367`).

### CTT (`:380`) `CTT*<fHLCount>` — count of **HL segments only** (S + #orders + #items), NOT line count.
### SE (`:400`) `SE*<fSegCount>*<%9.9d EIN>` — segment count ST..SE inclusive (SE counts itself).
### GE (`:421`) `GE*1*<%9.9d EIN>` (GE02 must = GS06). ### IEA (`:441`) `IEA*1*<%9.9d EIN>` (IEA02 must = ISA13).

## Separators (from the site row → `INV_SITES.VC_SEP_*`)
element = `VC_SEP_ELEMENT` (prod `*`), sub-element = `VC_SEP_SUBELEMENT` (ISA16 only), segment terminator =
`VC_SEP_SEGMENT` (prod typically `~`) — **legacy READS but never emits it** (relies on CRLF). See Trap 1 / the
terminator decision in the data doc.

## Control numbers + EIN — all four are the EIN
ISA13, GS06, ST02, and the BSN02 suffix are every one `%9.9d` of the EIN; GE02/IEA02/SE02 echo it. Legacy
allocated the EIN AT CREATE (`SiteEIN+1`) then `AD_UpdateEIN` bumped `Site.SiteEIN` AFTER the build (no WHERE
→ D1 hazard). **Rebuild: EIN = `INV_SITES.IN_EIN_SEQ`, allocated per-site AT SEND, `%09d` in all 7 positions;
BSN02 = `yyyymmdd` + `%09d`.**

## SE01 / CTT01 count rule (the TEMA-reject one)
**SE01 = count of segments ST..SE inclusive** (3 [ST,BSN,DTM] + `HLList.Count` + 1 [CTT] + 1 [SE]). The prior
"inconsistent `fSegCount`" flag in asn-invoice.md §4.2 is **WRONG — the legacy count is correct** (the per-
`HLList[i]` INC at `:367` catches TD1/TD5/TD3/PRF/LIN/SN1). **CTT01 = HL count only.** Rebuild: compute both
by counting actual emitted segments — never copy a magic offset.

## File output
- Dir = `[DIRECTORIES] EDIOut` (the log's `X:\EDIOut\` is that INI value). → gateway-side file I/O, path from `sites`/gateway config.
- **Decision E (REVISED — anchored on the OPERATIONAL SENDER, not the recreate button).** The rebuild reproduces
  `MainMenu.ResendMarkedEDIsClick` (`MainMenu.pas:2691-2771`), the live daily C→S-flip send path — NOT the
  `ASNInvoice` recreate button the build was originally anchored on (a different code path with a latently-
  inconsistent offset). The earlier decision E was WRONG for both branches: it dropped `LineName` AND used the
  recreate offset. The canonical operational-sender patterns:
  - **NORMAL** (`VC_START_SEQ_NUMBER <> '-1'`, `MainMenu.pas:2718`):
    `856 + copy(PickupDate,5,4) + LineName + .txt` = `856` + `prodDate[4:8]` (MMDD, 4 chars) + `LineName` + `.txt`.
    e.g. `20260618`/COROLLA → `8560618COROLLA.txt`.
  - **HOT-CALL** (`VC_START_SEQ_NUMBER = '-1'`, `MainMenu.pas:2722-2724`):
    `8HC + copy(PickupDate,4,5) + IntToStr(y) + LineName + .txt` = `8HC` + `prodDate[3:8]` (year-digit+MMDD,
    5 chars) + `y` + `LineName` + `.txt`. e.g. `20260618`/y=1/COROLLA → `8HC606181COROLLA.txt`.
  - **Note the deliberate legacy asymmetry**: NORMAL uses `[4:8]` (MMDD); HOT-CALL uses `[3:8]` (year+MMDD).
    Preserve it verbatim. `y` is a per-send-batch counter (init 1 `:2702`, `INC` only in the hot-call branch
    `:2724`); the rebuild sends per-ASN, so the driver derives a deterministic, collision-free per-ASN
    equivalent (`1 + count of same-day same-line hot-call ASNs already flipped to 'S'`). Single hot-call/day → 1.
    **The exact `y` RANGE is golden-pending (P13 cutover check)** — byte-faithful to the source pattern, range
    unverified until a golden `8HC` file exists.
  - (For reference, the recreate button `ASNInvoice:817-825` used `856<copy(PickupDate,5,4)>.txt` (no LineName)
    and `8HC<copy(PickupDate,4,5)>1.txt` (literal '1', no LineName) — NOT the operational pattern; do not port it.)

## Hardcoded / multi-site hazards
1. **No segment terminator emitted** (CRLF only). 2. **TD3 trailer id literal `1234567890`**. 3. Order HL02
hardcoded `1`. 4. Fixed qualifiers (ISA `00/01/ZZ/01/U/00400`, GS `SH/X/004010`, ST `856`, DTM `011…ET`, TD5
`B/25/00000`, TD3 `TL`, LIN `BP/RC`, SN1 `PC`) — correct to keep. 5. ISA09 `yymmdd` vs `yyyymmdd` elsewhere.
6. ISA `%-Ns` fixed widths (truncation risk if SiteAbbr/SupplierCode wider than 10/15). 7. `REPORT_EDI856`
hardcodes `WHERE IN_ASN_EIN=6440` + self-flips status (the reason to rebuild new, not wrap). 8. Window is now
**inclusive** (`<=`/`>=`) in the live proc — supersedes asn-invoice.md §4.1 strict-compare.

## What the rebuild MUST emit byte-for-byte
Segments in exact order; separators from the site row (and EMIT the real segment terminator — the one legacy
bug to fix only if TEMA requires it, see the parity decision); EIN `%09d` per-site-at-send in all 7 positions;
SE01 = ST..SE count, CTT01 = HL count; ISA widths 10/10/15/15 + ISA09 `yymmdd`; one shared `hhmm`; keep all
fixed qualifiers; parameterize the literals (TD3 trailer id, the `6440`).

## Top byte-exact traps
1. Segment terminator (CRLF-only legacy) — see parity decision. 2. SE01/CTT01 counts. 3. EIN `%09d` ×7 +
BSN02 17-char. 4. ISA09 `yymmdd` + ISA fixed widths. 5. Window now inclusive. 6. Don't wrap `REPORT_EDI856`.
