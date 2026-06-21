# EDI 810 Invoice — Wire-Format Spec (`EDI810Object.T810EDI`)

Source-truth for the M1 Rank-4 outbound-810 builder. LIVE builder = `EDI810Object.pas` (`T810EDI`,
`InventorySystem.dpr:47`); LIVE driver = `ASNInvoice.pas:857-895` (recreate) + `MainMenu.pas:2587`
(create). **`Write810File.pas` is DEAD** (entire EDI body is a `{...}` comment — do NOT port from it).
Twin of the 856: `REPORT_EDI810` self-flips status → the rebuild is **NEW gateway Jython**, NOT a wrap.

## Conventions (same as the 856)
- EIN `%9.9d` (9-digit zero-pad) in all 7 control positions. Separators from the site row
  (`INV_SITES.VC_SEP_*`); **segment terminator read but NEVER emitted — CRLF only** (same legacy bug).
- `f810Time = now` captured once → all date/time fields share it.

## Segment list (emission order) — `Execute:112-120`
**ISA · GS · ST · BIG · [per-manifest break: REF · DTM] · IT1(×N) · trailing REF · DTM · TDS · CTT · SE ·
GE · IEA.** NO N1 / TXI / SAC / CAD / ISS (none emitted; summary is just TDS).

- **ISA** (`:135-165`): positional `%-Ns` (ISA02/04=10, ISA06/08=15). **ISA09 = `yymmdd` of NOW** (the 810
  dates off *now*, NOT the pickup date — a 856 difference). ISA13=EIN, ISA15=EDIMode, ISA16=sub-elem sep.
- **GS** (`:175-188`): `GS*IN*<DUNS>*<TMMDUNS>*<yyyymmdd now>*<hhmm>*<%9.9d EIN>*X*004010` — **GS01=`IN`** (856 uses `SH`); GS04 = now.
- **ST** (`:200`): `ST*810*<%9.9d EIN>`.
- **BIG** (`:220`): `BIG*<yyyymmdd now>*<SiteSupplierCode>` — BIG01 = invoice date (now), **BIG02 = SiteSupplierCode** (the "invoice number" — constant per site; real uniqueness is the EIN in ISA13/ST02).
- **REF** (`:271/:322`): `REF*MK*<Manifest>` (MK = manifest key). Emitted on each manifest break + a trailing one.
- **DTM** (`:281/:330`): `DTM*050*<PickUpDate yyyymmdd>` (050 = received date).
- **IT1** (`:289-310`, per detail row): `IT1*<M391|M390>*<qty>*EA*<unitPrice>*QT*PN*<part>*PK*1*ZZ*<SiteDockCode>`.
  IT101 = `M391` if Manifest starts `'7'` (broadcast) else `M390` (hot-call). IT102 = `IN_QTY` (no leading zeros).
  **IT104 = unit price (MONEY)**. IT107 = part, IT111 = SiteDockCode. **NO trailing sep** (ends on dock code —
  do NOT copy the 856 LIN trailing-sep quirk here).
- **TDS** (`:337-353`): `TDS*<implied-decimal total>` (invoice total = Σ(unitPrice×qty)). See MONEY below.
- **CTT** (`:371`): `CTT*<IT1 line count>`. **SE** (`:391`): `SE*<segment count ST..SE>*<%9.9d EIN>`.
  **GE** (`:412`): `GE*1*<EIN>`. **IEA** (`:432`): `IEA*1*<EIN>`.

## Trailing-sep / empty-element audit (the 856 lesson — every Write call checked)
**The 810 is CLEAN: zero trailing separators, zero empty interior elements, no `**` anywhere.** Each segment
ends on a populated value (ISA on ISA16, IT1 on dock code, etc.). The byte-exact risks here are the MONEY
fields + the SE/CTT counts, NOT separators. (Verified line-by-line vs `EDI810Object.pas`.)

## MONEY (the 810-specific trap — TWO inconsistent formats)
1. **IT104 unit price** (`:298`): legacy `FloatToStr(MO_PRICE.AsCurrency)` → a **literal decimal point**, the
   **OS-locale separator**, and a **variable** fraction count (`12.5`, `12.55`, `12.5500`). Hazard: locale-
   sensitive + variable digits. Rebuild: emit a clean FIXED-scale decimal (decision/golden — DECISION 810-2).
2. **TDS01 total** (`:337-350`): legacy hand-rolls an **implied-decimal scale-4 (NO point)**: integer part +
   4-digit fraction concatenated, e.g. `1234.5` → `12345000`. **WITH BUGS:** a 1-digit fraction is NOT padded
   (`length=1` falls through → emits `12345`, an implied `1.2345` = **off by 10000×**); a whole-dollar total
   (no `.` in FloatToStr) yields a malformed `temp=''`/`totalstr=copy(...,1,4)`. **DECISION 810-1:** fix
   (compute `round(total×10000)` integer, correct scale-4) vs reproduce the buggy surgery. Recommend FIX (a
   wrong invoice TOTAL is a TEMA-reject / mis-bill — same spirit as the D6 fix). Values come from the D6 feed.

## Control numbers + EIN
All 7 positions = the **invoice** EIN (`INV_INV_MST.IN_INV_EIN`, int NOT NULL, unique-per-invoice), `%9.9d`.
**Allocated at invoice-CREATE** (`MainMenu.CreateINVOICEClick` → `INSERT_INVInfo`), **reused at recreate** (do
NOT re-allocate, or the 997-ack `UPDATE_EINStatus WHERE IN_INV_EIN=@EIN` won't land). NO hardcoded literal
(unlike the 856's `6440`). **Rebuild: allocate per-site from `INV_SITES.IN_EIN_SEQ` at invoice-create.** NOTE:
legacy shares ONE per-site counter (`AD_UpdateEIN`/`Site.SiteEIN`) for BOTH 856 and 810 → control numbers
interleave; the rebuild keeps the SHARED `INV_SITES.IN_EIN_SEQ` (faithful) unless we split (not recommended).

## SE01 / CTT01 (compute, don't copy — the 856 lesson)
SE01 = `2 (ST,BIG) + #IT1 + #REF + #DTM + 1 (TDS) + 1 (CTT) + 1 (SE itself)`, where #REF = #DTM = #distinct
manifests + 1 trailing. CTT01 = #IT1. Compute by counting emitted segments; verify on a multi-manifest file.

## File output
Path = `[DIRECTORIES] EDIOut` (→ gateway file I/O). Filename (recreate, `ASNInvoice.pas:872`) =
`'810' + copy(PickUpDate,5,4) + '.txt'` = `810<mmdd>.txt` (drops the year → cross-year collisions; pick one
deterministic pattern). One invoice file per pickup date (the IT1 loop breaks on a pickup-date change).

## Hazards / hardcoded
GS01=`IN`; BIG02=SupplierCode (constant); IT101 `M391`/`M390` by manifest `'7'` prefix; fixed qualifiers
(EA/QT/PN/PK/ZZ, REF `MK`, DTM `050`); ISA `%-Ns` widths; ISA09/GS04/BIG01 off NOW not pickup; CRLF-only;
`REPORT_EDI810` self-flips + is window-blind (→ rebuild new + D6 price). `Write810File.pas` DEAD.

## What the rebuild MUST emit byte-for-byte
Segments in order; CRLF-only; EIN `%09d` ×7 = the invoice EIN (allocate at invoice-create, reuse at recreate);
GS01=`IN`, dates off now, BIG02=SupplierCode; IT1 per the table (no trailing sep); CTT/SE computed; ISA widths.
MONEY per DECISION 810-1/810-2 (values from the D6 window-aware feed). NEW Jython, not a wrap.
