# Form-UX semantics: `SiteInfo.pas` — NOT A FORM (explicit finding)

**`SiteInfo.pas` is a plain `TObject` data-holder class, `TSiteInfo`, not a `TForm`.** There is no
`SiteInfo.dfm` in the repo (confirmed: `find . -iname "SiteInfo*"` returns only `SiteInfo.pas`), and
none is expected — the unit declares no visual component, no `ShowModal`/`Execute`, no event
handlers, no `uses ... Forms ...`/`Dialogs` at all (`SiteInfo.pas:1-83`). It is registered live in
`InventorySystem.dpr:58` (`SiteInfo in 'SiteInfo.pas'`, no form-class suffix like the other units in
this family — the `.dpr` entries for actual forms carry a `{TXxx_Form}` designer-class comment;
`SiteInfo`'s entry has none), consistent with it being a non-visual unit.

`TSiteInfo` is a read-only property bag (`SiteName`, `SiteAbbr`, `SiteStreet`, `SiteCity`,
`SiteState`, `SiteCountry`, `SiteZip`, `SiteDUNS`, `SiteSupplierCode`, `SiteDockCode`, `SiteEIN`,
`SitePassword`, `SiteEDIMode`, `SiteSepSegment`, `SiteSepElement`, `SiteSepSubElement`,
`SiteTMMName`, `SiteTMMAbbr`, `SiteTMMDUNS`, `SiteMaxSequence`, `AcceptAnyOrderASN`,
`SiteDeliveryMethodCode` — all `read`-only, no setters exposed on the public interface,
`SiteInfo.pas:33-76`) backing private fields (`fSiteName` etc., `:8-29`) that are never assigned
anywhere in this unit (`implementation` section is empty except for the class boilerplate,
`:80-82`) — this class is a passive value-object shape; whoever constructs and populates it (almost
certainly `Data_Module`, per this repo's convention of DB-config-driven site properties) lives
elsewhere and is out of scope for a form-UX extraction.

**Conclusion for this sweep: there is no UX semantics layer to extract for `SiteInfo` — no dialogs,
no confirmations, no field-clear/repopulate behavior, no focus/keyboard handling, no enable/disable
state machine, no form-level error surfacing.** If the issue's intent was the *site configuration
screen* (an admin UI for editing these site properties), that screen is a DIFFERENT unit not named
`SiteInfo.pas` in this repo — confirm with whoever filed #141 whether a `SiteMaint`/`SiteConfig`-
style form exists under another name before assuming this data class was the intended target.

## Cross-refs
- None applicable (non-visual unit; no proc/data spec doc exists for it either).
