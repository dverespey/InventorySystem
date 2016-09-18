program InventorySystem;

uses
  Forms,
  MainMenu in 'MainMenu.pas' {MainMenu_Form},
  DataModule in 'DataModule.pas' {Data_Module: TDataModule},
  UploadBreakDown in 'UploadBreakDown.pas' {UpBreakDown_Form},
  Order in 'Order.pas' {Order_Form},
  RecConfStat in 'RecConfStat.pas' {RecConfStat_Form},
  RecReject in 'RecReject.pas' {RecRej_Form},
  InvMgmt in 'InvMgmt.pas' {InvMgmt_Form},
  Stocktaking in 'Stocktaking.pas' {Stocktaking_Form},
  MasterMaint in 'MasterMaint.pas' {MasterMaint_Form},
  Shipping in 'Shipping.pas' {Shipping_Form},
  SupplierMaster in 'SupplierMaster.pas' {SupplierMaster_Form},
  SizeMaster in 'SizeMaster.pas' {SizeMaster_Form},
  AssyRatioMaster in 'AssyRatioMaster.pas' {AssyRatioMaster_Form},
  PartsStockMaster in 'PartsStockMaster.pas' {PartsStockMaster_Form},
  UserInfo in 'UserInfo.pas',
  VersionInfo in 'VersionInfo.pas',
  InvMgmtQReport in 'InvMgmtQReport.pas' {InvMgmtQReport_Form},
  ConfirmPassword in 'ConfirmPassword.pas' {ConfirmPassword_Form},
  UserAdmin in 'UserAdmin.pas' {UserAdmin_Form},
  Logon in 'Logon.pas' {Logon_Form},
  About in 'ABOUT.PAS' {AboutBox},
  OvertimeHoliday in 'OvertimeHoliday.pas' {OvertimeHoliday_Form},
  ForecastDetail in 'ForecastDetail.pas' {ForecastDetail_Form},
  ForecastBreakdownF in 'ForecastBreakdownF.pas' {ForecastBreakdown_Form},
  LogisticsMaster in 'LogisticsMaster.pas' {LogisticsMaster_Form},
  DirectorySelect in 'DirectorySelect.pas' {SelectDirectoryDlg},
  RenbanGroupMaster in 'RenbanGroupMaster.pas' {RenbanGroupMaster_Form},
  RenbanOrder in 'RenbanOrder.pas' {GroupRenbanOrder_Form},
  OrderFormCreateF in 'OrderFormCreateF.pas' {OrderFormCreate_Form},
  FRSBreakdown in 'FRSBreakdown.pas' {FRSBreakdownDlg},
  LogisticsBreakdown in 'LogisticsBreakdown.pas' {LogisticsBreakdown_Form},
  Configuration in 'Configuration.pas' {ConfigurationDlg},
  MonthlyReportSelect in 'MonthlyReportSelect.pas' {DateSelectDlg},
  MonthlySupplerOrderReport in 'MonthlySupplerOrderReport.pas' {MonthlySupplierOrderSummary_Form},
  MonthlyLogiticsOrderReport in 'MonthlyLogiticsOrderReport.pas' {MonthlyLogisticsOrderSummary_Form},
  InvoiceBreakdown in 'InvoiceBreakdown.pas' {InvoiceBreakdown_Form},
  MonthlySupplerInvoiceReport in 'MonthlySupplerInvoiceReport.pas' {MonthlySupplierInvoice_Form},
  FirstProductiionDay in 'FirstProductiionDay.pas' {FirstProductionDay_Form},
  ManualForecast in 'ManualForecast.pas' {ManualForecast_Form},
  DailyBuildTotal in 'DailyBuildTotal.pas' {DailyBuildtotalForm},
  MonthlyPOMaster in 'MonthlyPOMaster.pas' {MonthlyPOMaster_Form},
  ManualShipping in 'ManualShipping.pas' {ManualShipping_Form},
  EDI810Object in 'EDI810Object.pas',
  EDI856Object in 'EDI856Object.pas',
  ManifestCostMaster in 'ManifestCostMaster.pas' {ManifestCostMaster_Form},
  ProductionDates in 'ProductionDates.pas' {ProductionDateSelectDlg},
  ASNSelect in 'ASNSelect.pas' {ASNSelect_Form},
  ASNInvoice in 'ASNInvoice.pas' {ASNInvoice_Form},
  EDIUpload in 'EDIUpload.pas' {EDIUpload_Form},
  ModifyShipping in 'ModifyShipping.pas' {ModifyShipping_Form},
  HotCallEntry in 'HotCallEntry.pas' {HotCallEntryForm},
  NewPassword in 'NewPassword.pas' {NewPasswordDlg},
  Write810File in 'Write810File.pas',
  SiteInfo in 'SiteInfo.pas',
  ForecastCamexreport in 'ForecastCamexreport.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'Inventory Control';
  Application.CreateForm(TData_Module, Data_Module);
  Application.CreateForm(TMainMenu_Form, MainMenu_Form);
  Application.Run;
end.
