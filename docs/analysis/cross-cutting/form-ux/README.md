# Form-UX semantics quote-sheets (issue inventory#141, R109-S2)

The FORM-UX layer of every legacy form behind a built Ignition screen — dialogs,
armed confirms, field clear/repopulate, focus, enable/disable state machines,
error surfacing — extracted with `.pas`/`.dfm` `file:line` citations. This is the
layer the per-module proc/wire specs never covered; both inventory#134 (missing
armed delete) and inventory#135 (stale field values) lived here.

**Contract:** every claim cites source; `[UNVERIFIED]` marks anything not
provable from the source alone. One file per form unit, uniform sections
(Dialogs & confirmations / Field clear-repopulate / Focus & keyboard /
Enable-disable state machine / Error surfacing / Cross-refs).

**Forward norm (inventory CLAUDE.md, build recipe step 1):** a screen is not
done without its form-UX extraction landing here first.

## Coverage map (screen family → forms)

| Family | Forms | Notes |
|---|---|---|
| Masters | MasterMaint, SizeMaster, SupplierMaster, LogisticsMaster, ManifestCostMaster, PartsStockMaster, RenbanGroupMaster, ForecastDetail, AssyRatioMaster, SiteInfo | **The rebuild's AssemblyDetail screen maps to ForecastDetail.pas** (the live INV_FORECAST_DETAIL_INF editor), NOT AssyRatioMaster — AssyRatioMaster is DEAD in legacy (menu button hidden, MasterMaint.pas:78; assembly/assy-ratio-master.md rules "do NOT port") and documented for the verdict only. SiteInfo is a property bag, not a form |
| Order / HotCall | Order, OrderFormCreateF, RenbanOrder, HotCallEntry | OrderFormCreate + OrderQty are DEAD (documented for the verdict only) |
| Ship / Receive / Stock | Shipping, ManualShipping, ModifyShipping, ASNInvoice, ASNSelect, RecConfStat, RecReject, Stocktaking | |
| Forecast / EDI / Admin | ManualForecast, ForecastBreakdownF, EDIUpload, UserAdmin, NewPassword, ConfirmPassword | ForecastBreakDown + ForecastUploadBreakDown are DEAD; ConfirmPassword is DORMANT (call site commented out — no re-auth gate IS legacy behavior) |

## Cross-cutting conventions (bigger than any one form)

- **Combo "no selection" is a literal single-space string at `ItemIndex:=0`**, not
  `-1`/null (`DataModule.pas:5767-6020` ClearControls/SelectSingleField/SearchCombo)
  — and forms override it inconsistently (RecConfStat forces `-1`, Stocktaking
  doesn't, ManualShipping bypasses the helper). Any rebuild treating "empty" as one
  semantic diverges from at least one form.
- **Empty-date behavior differs across siblings**: Stocktaking defaults to `now`,
  RecReject keeps the prior value, RecConfStat goes true-blank.
- **Module-level `Data_Module` detail-panel state is shared across master forms** —
  the source of the stale-panel (#135-class) family of behaviors.
