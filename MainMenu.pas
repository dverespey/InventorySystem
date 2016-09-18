//****************************************************************
//
//       Inventory Control
//
//       Copyright (c) 2002-2008 Failproof Manufacturing Systems.
//
//****************************************************************
//
//  10/25/2002  Aaron Huge      Initial creation
//  lot of stuff.....
//  05/06/2008  David Verespey  Data Purge

unit MainMenu;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Menus, DataModule, Shipping, Logon,
  ConfirmPassword, UserAdmin, ADOdb, ComObj, Excel2000, ADOVersion,DateUtils,
  VersionInfo;

type
  TMainMenu_Form = class(TForm)
    MainMenu_Label: TLabel;
    Forecast_GroupBox: TGroupBox;
    Forecast_Button: TButton;
    Order_GroupBox: TGroupBox;
    Image1: TImage;
    Order_Button: TButton;
    RecConf_GroupBox: TGroupBox;
    RecStatMgmt_Button: TButton;
    MenuBar_MainMenu: TMainMenu;
    File_MenuItem: TMenuItem;
    File_Exit_MenuItem: TMenuItem;
    Window_MenuItem: TMenuItem;
    Window_ForecastUpload_MenuItem: TMenuItem;
    Window_RecCon_MenuItem: TMenuItem;
    Window_Order_MenuItem: TMenuItem;
    Window_ProdInst_MenuItem: TMenuItem;
    Window_RecConf_RecStatMngmt_MenuItem: TMenuItem;
    Window_RecConf_RecRej_MenuItem: TMenuItem;
    Shipping_GroupBox: TGroupBox;
    ProdRej_Button: TButton;
    InvMgmt_GroupBox: TGroupBox;
    InvMgmt_Button: TButton;
    Stocktaking_Button: TButton;
    RecRej_Button: TButton;
    MonthlyProc_GroupBox: TGroupBox;
    DataMaint_GroupBox: TGroupBox;
    DeleteData_Button: TButton;
    DateMaint_Button: TButton;
    Quit_Panel: TPanel;
    Quit_Button: TButton;
    Image2: TImage;
    Image3: TImage;
    Image5: TImage;
    Window_Shipping_MenuItem: TMenuItem;
    Window_Shipping_Car_MenuItem: TMenuItem;
    Window_Shipping_Truck_MenuItem: TMenuItem;
    Window_Shipping_ProdRej_MenuItem: TMenuItem;
    Window_InvMgmt_MenuItem: TMenuItem;
    Window_InvMgmt_InvMgmt_MenuItem: TMenuItem;
    Window_InvMgmt_Stocktaking_MenuItem: TMenuItem;
    Window_UploadInvoices_MenuItem: TMenuItem;
    Window_DataMaint_MenuItem: TMenuItem;
    Administration_MenuItem: TMenuItem;
    Administration_UserAdmin_MenuItem: TMenuItem;
    N1: TMenuItem;
    OvertimeHoliday1: TMenuItem;
    Help1: TMenuItem;
    About1: TMenuItem;
    CreateOrderForm_Button: TButton;
    CreateOrderForms1: TMenuItem;
    RecUpl_Button: TButton;
    ReceivingUpload1: TMenuItem;
    GroupRenbanOrderButton: TButton;
    Configuration1: TMenuItem;
    N2: TMenuItem;
    Reports1: TMenuItem;
    Partswithoutassembly1: TMenuItem;
    Tirewithoutassembly: TMenuItem;
    Wheelwithoutassembly: TMenuItem;
    LotLocation: TMenuItem;
    LogicalInventoryReport1: TMenuItem;
    MonthlyOrderSummary: TMenuItem;
    MonthlyLogistics: TMenuItem;
    MonthlySupplierInvoice1: TMenuItem;
    EmptyContainer: TMenuItem;
    ForecastvsUsage1: TMenuItem;
    PastDueFRS1: TMenuItem;
    FirstProductionDayoftheYear1: TMenuItem;
    N3: TMenuItem;
    ManualForecast_Button: TButton;
    EDIBREAK: TMenuItem;
    CreateASN: TMenuItem;
    CreateINVOICE: TMenuItem;
    N4: TMenuItem;
    VersionInfo: TVersionInfo;
    POReport: TMenuItem;
    ForecastSummary1: TMenuItem;
    ForecastPartSummary1: TMenuItem;
    EDIUploadBox: TGroupBox;
    EDIUpload_Button: TButton;
    N5: TMenuItem;
    DailyShipping: TMenuItem;
    DailyASNReport: TMenuItem;
    MonthlyASN: TMenuItem;
    INVOICEReport: TMenuItem;
    ResendMarkedEDIs: TMenuItem;
    ForecastDetail1: TMenuItem;
    ProcessPanel: TPanel;
    Panel2: TPanel;
    Label1: TLabel;
    Button1: TButton;
    DailySupplierOrders1: TMenuItem;
    MonthlyINVOICEAssyPartNumbers1: TMenuItem;
    DailyShippingRangeTireWheelPartNumbers: TMenuItem;
    HotCallASN8561: TMenuItem;
    DailySupplierOrderCostSummary1: TMenuItem;
    MonthlySupplierOrderCostSummary1: TMenuItem;
    ASNReportWithCostAssyPartNumbers1: TMenuItem;
    N6: TMenuItem;
    CamexTest1: TMenuItem;
    procedure RecStatMgmt_ButtonClick(Sender: TObject);
    procedure Quit_ButtonClick(Sender: TObject);
    procedure Forecast_ButtonClick(Sender: TObject);
    procedure Order_ButtonClick(Sender: TObject);
    procedure RecProdRej_ButtonClick(Sender: TObject);
    procedure CarOrTruckShip_ButtonClick(Sender: TObject);
    procedure InvMgmt_ButtonClick(Sender: TObject);
    procedure Stocktaking_ButtonClick(Sender: TObject);
    procedure DateMaint_ButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Administration_UserAdmin_MenuItemClick(Sender: TObject);
    procedure OvertimeHoliday1Click(Sender: TObject);
    procedure About1Click(Sender: TObject);
    procedure CreateOrderForm_ButtonClick(Sender: TObject);
    procedure RecUpl_ButtonClick(Sender: TObject);
    procedure GroupRenbanOrderButtonClick(Sender: TObject);
    procedure Configuration1Click(Sender: TObject);
    procedure TirewithoutassemblyClick(Sender: TObject);
    procedure WheelwithoutassemblyClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure LotLocationClick(Sender: TObject);
    procedure DoBoarders(topedge,bottomedge:integer;startcolumn,endcolumn:string;fill:boolean=FALSE);
    procedure LogicalInventoryReport1Click(Sender: TObject);
    procedure MonthlyOrderSummaryClick(Sender: TObject);
    procedure MonthlyLogisticsClick(Sender: TObject);
    procedure Window_UploadInvoices_MenuItemClick(Sender: TObject);
    procedure MonthlySupplierInvoice1Click(Sender: TObject);
    procedure EmptyContainerClick(Sender: TObject);
    procedure ForecastvsUsage1Click(Sender: TObject);
    procedure PastDueFRS1Click(Sender: TObject);
    procedure FirstProductionDayoftheYear1Click(Sender: TObject);
    procedure ManualForecast_ButtonClick(Sender: TObject);
    procedure CreateASNClick(Sender: TObject);
    procedure CreateINVOICEClick(Sender: TObject);
    procedure POReportClick(Sender: TObject);
    procedure ForecastSummary1Click(Sender: TObject);
    procedure ForecastPartSummary1Click(Sender: TObject);
    procedure EDIUpload_ButtonClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure DailyShippingClick(Sender: TObject);
    procedure DailyASNReportClick(Sender: TObject);
    procedure MonthlyASNClick(Sender: TObject);
    procedure INVOICEReportClick(Sender: TObject);
    procedure ResendMarkedEDIsClick(Sender: TObject);
    procedure ForecastDetail1Click(Sender: TObject);
    procedure DailySupplierOrders1Click(Sender: TObject);
    procedure MonthlyINVOICEAssyPartNumbers1Click(Sender: TObject);
    procedure DailyShippingRangeTireWheelPartNumbersClick(Sender: TObject);
    procedure HotCallASN8561Click(Sender: TObject);
    procedure DailySupplierOrderCostSummary1Click(Sender: TObject);
    procedure MonthlySupplierOrderCostSummary1Click(Sender: TObject);
    procedure ASNReportWithCostAssyPartNumbers1Click(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure CamexTest1Click(Sender: TObject);
  private
    { Private declarations }
	  fAppSigHandle	:	THandle;
    procedure CatchAll (Sender: TObject; E: Exception);
    procedure Configure;
  public
    { Public declarations }
  end;

var
  MainMenu_Form: TMainMenu_Form;

implementation


uses UploadBreakDown, Order, RecConfStat, RecReject, InvMgmt,
  Stocktaking, MasterMaint, OvertimeHoliday, About,
  RenbanOrder, OrderFormCreateF, Configuration, MonthlyReportSelect,
  MonthlySupplerOrderReport, MonthlyLogiticsOrderReport,
  MonthlySupplerInvoiceReport, ManualForecast, FirstProductiionDay,
  DailyBuildTotal, ManualShipping, EDI810Object, EDI856Object,
  ProductionDates, ASNSelect, EDIUpload, HotCallEntry, ForecastCamexreport;

{$R *.dfm}


function LockSignature (Signature	:	string; var SigHandle: THandle)	:	boolean;
begin
  SigHandle := CreateSemaphore (nil, 0, 1, PChar (Signature));
  result    := (SigHandle <> 0) and (GetLastError() <> ERROR_ALREADY_EXISTS);

  if (NOT result) then
  	if (SigHandle <> 0) then
			CloseHandle (SigHandle);
end;


procedure TMainMenu_Form.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Data_Module.LogActLog('STOP','Program stop');
end;

procedure TMainMenu_Form.Quit_ButtonClick(Sender: TObject);
begin
  Close;
end;

procedure TMainMenu_Form.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  // Check for renban order not created
  // Confirm all orders are sent
  if Data_Module.fiCreatePOPriorToClose.AsBoolean then
  begin
    if Data_Module.OpenPOOrders then
    begin
      if MessageDlg('There are unprocessed orders, are you sure you want to quit?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
        CanClose:=TRUE;
      end
      else
      begin
        CanClose:=FALSE;
      end;
    end;
  end;
end;

procedure TMainMenu_Form.Forecast_ButtonClick(Sender: TObject);
begin
  Hide;
  UpBreakDown_Form:=TUpBreakDown_Form.Create(self);
  UpBreakDown_Form.BreakdownKind:=bForecast;
  UpBreakDown_Form.Execute;
  UpBreakDown_Form.Free;
  Show;
end;

procedure TMainMenu_Form.Order_ButtonClick(Sender: TObject);
begin
  Hide;
  Order_Form:=TOrder_Form.Create(self);
  Order_Form.Execute;
  Order_Form.Free;
  Show;
end;

procedure TMainMenu_Form.CreateOrderForm_ButtonClick(Sender: TObject);
begin
  if Data_Module.fiConfirmOrderFileCreation.AsBoolean then
  begin
    if MessageDlg('Would you like to create the order files now?', mtConfirmation, [mbYes, mbNo], 0) = mrNo then
    begin
      exit;
    end;
  end;
  OrderFormCreate_Form:=TOrderFormCreate_Form.Create(self);
  OrderFormCreate_Form.FileKind:=fText;
  OrderFormCreate_Form.Execute;
  OrderFormCreate_Form.Free;
end;

procedure TMainMenu_Form.RecStatMgmt_ButtonClick(Sender: TObject);
begin
  Hide;
  RecConfStat_Form:=TRecConfStat_Form.Create(self);
  if RecConfStat_Form.Execute then
    SHowMessage('Order file generation completed');
  RecConfStat_Form.Free;
  Show;
end;

procedure TMainMenu_Form.RecProdRej_ButtonClick(Sender: TObject);
//var
  //fButton: String;
begin
{
The RecRej_Button and the Window_RecConf_RecRej_MenuItem have a tag value of 1
The ProdRej_Button and the Window_Shipping_Truck_MenuItem have a tag value of 2
}
  Case TControl(Sender).Tag of
    1: Data_Module.Division := '1';
    2: Data_Module.Division := '2';
  else
    Data_Module.Division := '1';
  end;

  Hide;
  RecRej_Form:=TRecRej_Form.Create(self);
  RecRej_Form.Execute;
  RecRej_Form.Free;
  Show;
end;

procedure TMainMenu_Form.CarOrTruckShip_ButtonClick(Sender: TObject);
begin
  if Data_Module.fiFileALC.AsBoolean then
  begin
    //
    //  ToDo Fix Manual Shipping form
    //
    //
    Hide;
    ManualShipping_Form:=TManualShipping_Form.Create(self);
    ManualShipping_Form.Execute;
    ManualShipping_Form.Free;
    {UpBreakDown_Form:=TUpBreakDown_Form.Create(self);
    UpBreakDown_Form.BreakdownKind:=bDailyBuildT;
    UpBreakDown_Form.Execute;
    UpBreakDown_Form.Free;}
    Show;
  end
  else
  begin
    Hide;
    Shipping_Form:=TShipping_Form.Create(self);
    Shipping_Form.Execute;
    Shipping_Form.Free;
    Show;
  end;
end;

procedure TMainMenu_Form.InvMgmt_ButtonClick(Sender: TObject);
begin
  Hide;
  InvMgmt_Form:=TInvMgmt_Form.Create(self);
  InvMgmt_Form.Execute;
  InvMgmt_Form.Free;
  Show;
end;

procedure TMainMenu_Form.Stocktaking_ButtonClick(Sender: TObject);
begin
  Hide;
  Stocktaking_Form:=TStocktaking_Form.Create(self);
  Stocktaking_Form.Execute;
  Stocktaking_Form.Free;
  Show;
end;

procedure TMainMenu_Form.DateMaint_ButtonClick(Sender: TObject);
begin
  Hide;
  MasterMaint_Form:=TMasterMaint_Form.Create(self);
  MasterMaint_Form.Execute;
  MasterMaint_Form.Free;
  Show;
end;

procedure TMainMenu_Form.CatchAll (Sender: TObject; E: Exception);
begin
  ShowMessage('Unhandled exception from ('+Sender.ClassName+'), '+e.message);
  Data_Module.LogActLog('ERROR','Unhandled Exception from ('+Sender.ClassName+')'+e.message);
	Application.Terminate;
end;

procedure TMainMenu_Form.Configure;
begin
  ManualForecast_Button.Visible:=Data_Module.fiBuildOut.AsBoolean;

  if data_module.fiPOEDISupport.AsBoolean then
  begin
    HotCallASN8561.Visible:=TRUE;
    CreateASN.Visible:=TRUE;
    CreateINVOICE.Visible:=TRUE;
    ResendMarkedEDIs.Visible:=TRUE;
    DailyASNReport.Visible:=TRUE;
    MonthlyASN.Visible:=TRUE;
    INVOICEReport.Visible:=TRUE;
    if data_module.fiGenerateEDI.AsBoolean then
      POReport.Visible:=FALSE
    else
      POReport.Visible:=TRUE;
  end
  else
  begin
    HotCallASN8561.Visible:=FALSE;
    CreateASN.Visible:=FALSE;
    CreateINVOICE.Visible:=FALSE;
    ResendMarkedEDIs.Visible:=FALSE;
    DailyASNReport.Visible:=FALSE;
    MonthlyASN.Visible:=FALSE;
    INVOICEReport.Visible:=FALSE;
    POReport.Visible:=FALSE;
  end;

  Caption:='Inventory System for '+Data_Module.fiAssemblerName.AsString+'/'+Data_Module.fiPlantName.AsString;
end;

procedure TMainMenu_Form.FormCreate(Sender: TObject);
var
  purge:boolean;
  Year,Month,Day:word;
  MaxDays:integer;
begin
  purge:=TRUE;
 	Application.OnException := CatchAll;

	if (NOT LockSignature ('InventorySystem',fAppSigHandle)) then
  begin
  	ShowMessage('Inventory System already running');
  	Application.Terminate;
    exit;
  end;
  //
  //  Check program version if not valid then close
  //
  //  replace with the check version component later
  //
  VersionInfo.FilePath:=Application.ExeName;

  With Data_Module.Inv_DataSet do
  Begin
    Close;
    CommandType := CmdStoredProc;
    CommandText := 'dbo.SELECT_ProgramVersion;1';
    Parameters.Clear;

    Open;
    if fieldbyname('Program_Version').AsString <> VersionInfo.GetVersion then
    begin
      ShowMessage('Current database program version is '+fieldbyname('Program_Version').AsString+', you are running version '+VersionInfo.GetVersion+', please upgrade. Program terminating');
      application.Terminate;
      exit;
    end;
    Close;
  end;

  Hide;
  Logon_Form:=TLogon_Form.Create(self);
  If Not Logon_Form.Execute Then
  begin
    Application.Terminate;
    Data_Module.LogActLog('LOGIN','User Login failed');
  end;
  Logon_Form.Free;
  If Not Data_Module.gobjUser.AppUserAdmin Then
    MenuBar_MainMenu.Items.Remove(Administration_MenuItem);


  Data_Module.LogActLog('START','Program start Ver '+VersionInfo.GetVersion+' '+DateTimeTostr(FileDateToDateTime(FileAge(application.ExeName))));

  Configure;

  // Add Data Purge Code DMV
  //
  //
  if Data_Module.fiEnableDataPurge.AsBoolean then
  begin
    //check if purge time is here
    Data_Module.NT.NummiTime:=Data_Module.fiLastPurge.AsString;

    DecodeDate((now - Data_Module.NT.Time),Year,Month,Day);

    if uppercase(Data_Module.fiPurgeRate.AsString) = 'DAILY' then
    begin
      if now - Data_Module.NT.Time < 1 then
        purge:=FALSE;
    end
    else if uppercase(Data_Module.fiPurgeRate.AsString) = 'WEEKLY' then
    begin
      if now - Data_Module.NT.Time < 7 then
        purge:=FALSE
      else
      begin
        if uppercase(formatdatetime('dddd',now)) <> uppercase(Data_Module.fiPurgeDayWeekly.AsString) then
          purge:=FALSE;
      end;
    end
    else if uppercase(Data_Module.fiPurgeRate.AsString) = 'MONTHLY' then
    begin
      if now - Data_Module.NT.Time < 28 then //Cheat make it easy
        purge:=FALSE
      else
      begin
        purge:=FALSE;
        //Check for Last Day of month
        MaxDays:=MonthDays[IsLeapYear(StrToInt(formatDateTime('yyyy',now))), StrToInt(formatdatetime('M',now))];
        if MaxDays = StrToInt(formatdatetime('d',now)) then
        begin
          if uppercase(Data_Module.fiPurgeDayMonthly.AsString) = 'LAST' then
            Purge:=TRUE;
        end

        //Check for 15th
        else if uppercase(formatdatetime('D',now)+'TH') = uppercase(Data_Module.fiPurgeDayMonthly.AsString) then
          Purge:=TRUE

        //Check for 1st
        else if uppercase(formatdatetime('D',now)+'ST') = uppercase(Data_Module.fiPurgeDayMonthly.AsString) then
          Purge:=TRUE;

      end;
    end;


    if purge then
    begin
      if Data_Module.fiPromptDataPurge.AsBoolean then
      begin
        if Data_Module.fiPromptDataPurge.AsBoolean then
        begin
          if MessageDlg('Would you like to run the Data Purge now?', mtConfirmation, [mbYes, mbNo], 0) = mrNo then
          begin
            exit;
          end;
        end;
      end;

      // Do Purge
      if Data_Module.AutoPurge then
      begin
        // Set Last Time
        Data_Module.NT.Time:=now;
        Data_Module.fiLastPurge.AsString:=Data_Module.NT.NummiTime;
        ShowMessage('Auto database purge complete');
      end
      else
      begin
        ShowMessage('Auto database purge failed, will try again on next program start up');
      end;
    end;

  end;

  //Show;
end;

procedure TMainMenu_Form.Administration_UserAdmin_MenuItemClick(
  Sender: TObject);
begin
  //ConfirmPassword_Form := TConfirmPassword_Form.Create(self);
  //If ConfirmPassword_Form.Execute Then
  //Begin
    //ConfirmPassword_Form.Free;
    UserAdmin_Form := TUserAdmin_Form.Create(Self);
    UserAdmin_Form.Execute;
    UserAdmin_Form.Free;
  //End
  //Else
    //ConfirmPassword_Form.Free;
end;        //Administration_UserAdmin_MenuItemClick

procedure TMainMenu_Form.OvertimeHoliday1Click(Sender: TObject);
begin
  try
    OvertimeHoliday_Form:=TOvertimeHoliday_Form.Create(self);
    OvertimeHoliday_Form.Execute;
    OvertimeHoliday_Form.Free;
  except
    on e:exception do
    begin
      ShowMessage('Unable to access Overtime/Holiday data, '+e.Message);
    end;
  end;
end;

procedure TMainMenu_Form.About1Click(Sender: TObject);
begin
  AboutBox:=TAboutBox.Create(self);
  AboutBox.ProgramVersion:=VersionInfo.GetVersion;
  AboutBox.Execute;
  AboutBox.Free;
end;


procedure TMainMenu_Form.RecUpl_ButtonClick(Sender: TObject);
begin
  Hide;
  UpBreakDown_Form:=TUpBreakDown_Form.Create(self);
  UpBreakDown_Form.BreakdownKind:=bReceiving;
  UpBreakDown_Form.Execute;
  UpBreakDown_Form.Free;
  Show;
end;

procedure TMainMenu_Form.GroupRenbanOrderButtonClick(Sender: TObject);
begin
  Hide;
  GroupRenbanOrder_Form:=TGroupRenbanOrder_Form.Create(self);
  GroupRenbanOrder_Form.Execute;
  GroupRenbanOrder_Form.Free;
  Show;
end;

procedure TMainMenu_Form.Configuration1Click(Sender: TObject);
begin
  ConfigurationDlg:=TConfigurationDlg.Create(self);
  ConfigurationDlg.Execute;
  if not ConfigurationDlg.Cancel and ConfigurationDlg.DataChanged then
  begin
    ShowMessage('Configuration update complete');
    Configure;
  end;
end;

procedure TMainMenu_Form.TirewithoutassemblyClick(Sender: TObject);
var
  z:integer;
begin
  with Data_Module.Inv_DataSet do
  begin
    try
      ProcessPanel.Visible:=TRUE;
      application.ProcessMessages;
      Close;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.REPORT_UnusedTirePartNumbers;1';
      Parameters.Clear;
      Open;
      excel := createOleObject('Excel.Application');
      excel.visible := False;
      excel.DisplayAlerts := False;
      excel.workbooks.open(Data_Module.TemplateDir+'ReportTemplate.xls');
      mysheet := excel.workSheets[1];
      mysheet.cells[1,1].value:='Tire Part Numbers Not Assigned to Assembly Part Number';
      z:=4;
      while not EOF do
      begin
        mysheet.Cells[z,1].value := fieldbyname('VC_PART_NUMBER').AsString;
        INC(z);
        Next;
      end;
      ProcessPanel.Visible:=FALSE;
      excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\TireWithoutAssembly'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
      excel.Workbooks.Close;
      excel.Quit;
      excel:=Unassigned;
      mysheet:=Unassigned;
      Data_Module.LogActLog('REPORT','Do Unused Tire report');
      ShowMessage('Report Complete');
    except
      on e:exception do
      begin
        Data_Module.LogActLog('ERROR','Failed on Unused Tire Report, '+e.Message);
        ShowMessage('Failed on Unused Tire Report, '+e.Message);
        excel.Workbooks.Close;
        excel.Quit;
        excel:=Unassigned;
        mysheet:=Unassigned;
        ProcessPanel.Visible:=FALSE;
      end;
    end;
  end;
end;

procedure TMainMenu_Form.WheelwithoutassemblyClick(Sender: TObject);
var
  z:integer;
begin
  with Data_Module.Inv_DataSet do
  begin
    try
      ProcessPanel.Visible:=TRUE;
      application.ProcessMessages;
      Close;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.REPORT_UnusedWheelPartNumbers;1';
      Parameters.Clear;
      Open;
      excel := createOleObject('Excel.Application');
      excel.visible := False;
      excel.DisplayAlerts := False;
      excel.workbooks.open(Data_Module.TemplateDir+'ReportTemplate.xls');
      mysheet := excel.workSheets[1];
      mysheet.cells[1,1].value:='Wheel Part Numbers Not Assigned to Assembly Part Number';
      z:=4;
      while not EOF do
      begin
        mysheet.Cells[z,1].value := fieldbyname('VC_PART_NUMBER').AsString;
        INC(z);
        Next;
      end;
      ProcessPanel.Visible:=FALSE;
      excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\WheelWithoutAssembly'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
      excel.Workbooks.Close;
      excel.Quit;
      excel:=Unassigned;
      mysheet:=Unassigned;
      Data_Module.LogActLog('REPORT','Do Unused Wheel report');
      ShowMessage('Report Complete');
    except
      on e:exception do
      begin
        Data_Module.LogActLog('ERROR','Failed on Unused Wheel Report, '+e.Message);
        ShowMessage('Failed on Unused Wheel Report, '+e.Message);
        excel.Workbooks.Close;
        excel.Quit;
        excel:=Unassigned;
        mysheet:=Unassigned;
        ProcessPanel.Visible:=FALSE;
      end;
    end;
  end;
end;

procedure TMainMenu_Form.LotLocationClick(Sender: TObject);
var
  lastSize,lastWheel:string;
  i,z,lastbreak:integer;
begin
  lastSize:='X';
  lastWheel:='X';
  with Data_Module.Inv_DataSet do
  begin
    try
      ProcessPanel.Visible:=TRUE;
      application.ProcessMessages;
      excel := createOleObject('Excel.Application');
      excel.visible := FALSE;
      excel.DisplayAlerts := False;
      excel.workbooks.open(Data_Module.TemplateDir+'ReportTemplate.xls');
      mysheet := excel.workSheets[1];

      mysheet.Range['A1:G500'].Font.Size := 10;

      mysheet.cells[1,1].value:=Data_Module.fiPlantName.AsString+' Lot Location';
      mysheet.Range['A1','F1'].Interior.ColorIndex:=Data_Module.XlColToInt('AI');
      mysheet.Range['A1:A1'].Font.Size := 14;
      mysheet.PageSetup.PrintTitleRows := '$1:$4';
      mysheet.Rows[3].HorizontalAlignment := xlCenter;
      mysheet.Rows[3].VerticalAlignment := xlCenter;

      mysheet.cells[3,1].value:='Size';
      mysheet.cells[3,2].value:='Parking';
      mysheet.cells[3,3].value:='Trailer';
      mysheet.cells[3,4].value:='Renban';
      mysheet.cells[3,5].value:='Arrived';
      mysheet.cells[3,6].value:=Data_Module.fiAssemblerName.AsString+' Loc';
      mysheet.cells[3,7].value:=Data_Module.fiPlantName.AsString+' Loc';

      mysheet.Range['A3','G3'].Interior.ColorIndex:=Data_Module.XlColToInt('O');

      mysheet.Columns[1].ColumnWidth:=11;
      mysheet.Columns[1].VerticalAlignment := xlCenter;

      mysheet.Columns[2].ColumnWidth:=11;
      mysheet.Columns[2].HorizontalAlignment := xlCenter;
      mysheet.Columns[2].VerticalAlignment := xlCenter;

      mysheet.Columns[3].ColumnWidth:=14;
      mysheet.Columns[3].HorizontalAlignment := xlCenter;
      mysheet.Columns[3].VerticalAlignment := xlCenter;

      mysheet.Columns[4].ColumnWidth:=11;
      mysheet.Columns[4].HorizontalAlignment := xlCenter;
      mysheet.Columns[4].VerticalAlignment := xlCenter;

      mysheet.Columns[5].ColumnWidth:=10;
      mysheet.Columns[5].HorizontalAlignment := xlCenter;
      mysheet.Columns[5].VerticalAlignment := xlCenter;

      mysheet.Columns[6].ColumnWidth:=12;
      mysheet.Columns[6].HorizontalAlignment := xlCenter;
      mysheet.Columns[6].VerticalAlignment := xlCenter;

      mysheet.Columns[7].ColumnWidth:=12;
      mysheet.Columns[7].HorizontalAlignment := xlCenter;
      mysheet.Columns[7].VerticalAlignment := xlCenter;

      for i:=1 to 150 do
      begin
        mysheet.Rows[inttostr(i)].RowHeight:=20;
      end;
      z:=5;

      //
      //  Process wheels
      //
      Close;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.REPORT_PLANTLotLocationW;1';
      Parameters.Clear;
      Open;
      lastwheel := '';
      while not EOF do
      begin
        if lastwheel <> fieldbyname('vc_line_name').AsString then
        begin
          if fieldbyname('vc_line_name').AsString = 'Passenger' then
            mysheet.cells[z,1].value:='Car Wheels'
          else
          begin
            inc(z);
            mysheet.cells[z,1].value:='Truck Wheels';
          end;
          lastwheel:=fieldbyname('vc_line_name').AsString;
        end;
        if fieldbyname('vc_assembler_location').AsString = '' then
          mysheet.Cells[z,2].value:=fieldbyname('vc_PLANT_parking').AsString
        else
          mysheet.Cells[z,2].value:=fieldbyname('vc_ASSEMBLER_location').AsString;

        mysheet.Cells[z,3].value := fieldbyname('vc_trailer_number').AsString;
        mysheet.Cells[z,4].value :=fieldbyname('vc_renban_number').AsString;
        mysheet.Cells[z,5].value := copy(fieldbyname('vc_status_PLANT_yard').AsString,1,4)+'/'+copy(fieldbyname('vc_status_PLANT_yard').AsString,5,2)+'/'+copy(fieldbyname('vc_status_PLANT_yard').AsString,7,2);
        DoBoarders(z,z,'F','F',FALSE);
        DoBoarders(z,z,'G','G',FALSE);
        INC(z);

        Next;
      end;

      //
      //  Process Tires
      //
      lastbreak:=z;
      Close;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.REPORT_PLANTLotLocation;1';
      Parameters.Clear;
      Open;
      while not EOF do
      begin
        // tires are after wheels fill sheet
        if FieldByName('vc_size_code').AsString <> lastsize then
        begin
          // Header
          INC(z);
          INC(lastbreak);
          if FieldByName('vc_size_code').AsString = '' then
            mysheet.cells[z,1].value:='<NO SIZE>'
          else
            mysheet.cells[z,1].value:=FieldByName('vc_size_code').AsString;
          lastsize:=FieldByName('vc_size_code').AsString;

          // page break if split over page
          if ((lastbreak mod 33) >= 24) then
          begin
            mysheet.HPageBreaks.Add( mysheet.Cells.Item[z, 1] );
            lastbreak:=1;
          end;
        end;

        if fieldbyname('vc_assembler_location').AsString = '' then
          mysheet.Cells[z,2]:=fieldbyname('vc_PLANT_parking').AsString
        else
          mysheet.Cells[z,2]:=fieldbyname('vc_ASSEMBLER_location').AsString;
        mysheet.Cells[z,3].value := fieldbyname('vc_trailer_number').AsString;
        mysheet.Cells[z,4].value := fieldbyname('vc_renban_number').AsString;
        mysheet.Cells[z,5].value := copy(fieldbyname('vc_status_PLANT_yard').AsString,1,4)+'/'+copy(fieldbyname('vc_status_PLANT_yard').AsString,5,2)+'/'+copy(fieldbyname('vc_status_PLANT_yard').AsString,7,2);
        DoBoarders(z,z,'F','F',FALSE);
        DoBoarders(z,z,'G','G',FALSE);
        INC(z);
        INC(lastbreak);
        Next;
      end;


      excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\'+Data_Module.fiPlantName.AsString+'LotLocation'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');

      ProcessPanel.Visible:=FALSE;

      if MessageDlg('Print this report?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
        mysheet.PrintOut (EmptyParam, EmptyParam, EmptyParam,
                          EmptyParam, EmptyParam, EmptyParam, EmptyParam);
      end;
      excel.Workbooks.Close;
      excel.Quit;
      excel:=Unassigned;
      mysheet:=Unassigned;
      Data_Module.LogActLog('REPORT','Do Lot Location');
      ShowMessage('Report Complete');
    except
      on e:exception do
      begin
        Data_Module.LogActLog('ERROR','Failed on Lot Location, '+e.Message);
        ShowMessage('Failed on Lot Location, '+e.Message);
        excel.Workbooks.Close;
        excel.Quit;
        excel:=Unassigned;
        mysheet:=Unassigned;
        ProcessPanel.Visible:=FALSE;
      end;
    end;
  end;
end;

procedure TMainMenu_Form.DoBoarders(topedge,bottomedge:integer;startcolumn,endcolumn:string;fill:boolean=FALSE);
begin
  mysheet.Range[startcolumn+inttostr(topedge), endcolumn+inttostr(bottomedge)].Borders.Item[xlEdgeTop].Linestyle
                    :=  xlContinuous;
  mysheet.Range[startcolumn+inttostr(topedge), endcolumn+inttostr(bottomedge)].Borders.Item[xlEdgeBottom].Linestyle
                    :=  xlContinuous;
  mysheet.Range[startcolumn+inttostr(topedge), endcolumn+inttostr(bottomedge)].Borders.Item[xlEdgeRight].Linestyle
                    :=  xlContinuous;
  mysheet.Range[startcolumn+inttostr(topedge), endcolumn+inttostr(bottomedge)].Borders.Item[xlEdgeLeft].Linestyle
                    :=  xlContinuous;
  if fill then
  begin
    if startcolumn <> endcolumn then
      mysheet.Range[startcolumn+inttostr(topedge), endcolumn+inttostr(bottomedge)].Borders.Item[xlInsideVertical].Linestyle
                       :=  xlContinuous;
      mysheet.Range[startcolumn+inttostr(topedge), endcolumn+inttostr(bottomedge)].Borders.Item[xlInsideHorizontal].Linestyle
                    :=  xlContinuous;
  end;
end;

procedure TMainMenu_Form.LogicalInventoryReport1Click(Sender: TObject);
var
  z,sizeCount,sumIT,sumOH,sumTotal:integer;
  lastSize:string;
begin
  //get inventory report
  lastSize:='X';
  sizeCount:=1;
  with Data_Module.Inv_DataSet do
  begin
    try
      ProcessPanel.Visible:=TRUE;
      application.ProcessMessages;
      Close;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.REPORT_LogicalInventory;1';
      Parameters.Clear;
      Open;
      excel := createOleObject('Excel.Application');
      excel.visible := False;
      excel.DisplayAlerts := False;
      excel.workbooks.open(Data_Module.TemplateDir+'ReportTemplate.xls');
      mysheet := excel.workSheets[1];
      mysheet.cells[1,1].value:='Logical Inventory';

      mysheet.cells[3,1].value:='Size';
      mysheet.Columns[1].ColumnWidth:=10;

      mysheet.cells[3,2].value:='Part Number';
      mysheet.Columns[2].ColumnWidth:=18;
      mysheet.Columns[2].HorizontalAlignment := xlHAlignLeft;
      mysheet.Columns[2].NumberFormat := '############';

      mysheet.cells[3,3].value:='Kanban';

      mysheet.cells[3,4].value:='In Transit';

      mysheet.cells[3,5].value:='On Hand';

      mysheet.cells[3,6].value:='Total';

      z:=4;
      while not EOF do
      begin
        if lastSize<>FieldByName('VC_SIZE_CODE').AsString then
        begin
          //check last size for multiple parts
          if (sizeCount>1) and (lastSize<>'') then
          begin
            //mulitple do a summary
            mysheet.cells[z,3].value:='Size Total';
            mysheet.cells[z,4].value:=sumIT;
            mysheet.cells[z,5].value:=sumOH;
            mysheet.cells[z,6].value:=sumTotal;
            INC(z);
          end;

          sizeCount:=0;

          INC(z);
          if FieldByName('VC_SIZE_CODE').AsString = '' then
            mysheet.cells[z,1].value:='<No Size>'
          else
            mysheet.cells[z,1].value:=FieldByName('VC_SIZE_CODE').AsString;

          lastSize:=FieldByName('VC_SIZE_CODE').AsString;
          sumIT:=0;
          sumOH:=0;
          sumTotal:=0;
        end;

        mysheet.cells[z,2].value:=FieldByName('VC_PART_NUMBER').AsString;
        mysheet.cells[z,3].value:=FieldByName('VC_KANBAN_NUMBER').AsString;
        mysheet.cells[z,6].value:=FieldByName('IN_QTY').AsString;
        sumTotal:=sumTotal+FieldByName('IN_QTY').AsInteger;
        // Get InTransit
        with Data_Module.Inv_Check_DataSet do
        begin
          Close;
          CommandType := CmdStoredProc;
          CommandText := 'dbo.SELECT_OrderInTransit;1';
          Parameters.Clear;
          Parameters.AddParameter.Name := '@PartNumber';
          Parameters.ParamValues['@PartNumber'] := Data_Module.Inv_DataSet.FieldByName('VC_PART_NUMBER').AsString;
          Open;
          mysheet.cells[z,4].value:=FieldByName('QTY').AsString;
          sumIT:=sumIT+FieldByName('QTY').AsInteger;
          mysheet.cells[z,5].value:=inttostr(Data_Module.Inv_DataSet.FieldByName('IN_QTY').AsInteger-FieldByName('QTY').AsInteger);
          sumOH:=sumOH+(Data_Module.Inv_DataSet.FieldByName('IN_QTY').AsInteger-FieldByName('QTY').AsInteger);
        end;
        next;
        INC(z);
        INC(sizeCount);
      end;
      DoBoarders(4,z,'A','F',TRUE);
      excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\LogicalInventory'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');

      ProcessPanel.Visible:=FALSE;

      if MessageDlg('Print this report?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin

        mysheet.PrintOut (EmptyParam, EmptyParam, EmptyParam,
                          EmptyParam, EmptyParam, EmptyParam, EmptyParam);
      end;
      excel.Workbooks.Close;
      excel.Quit;
      excel:=Unassigned;
      mysheet:=Unassigned;
      Data_Module.LogActLog('REPORT','Do Logical Inventory');
      ShowMessage('Report Complete');
    except
      on e:exception do
      begin
        Data_Module.LogActLog('ERROR','Failed on Logical Inventory, '+e.Message);
        ShowMessage('Failed on Logical Inventory, '+e.Message);
        excel.Workbooks.Close;
        excel.Quit;
        excel:=Unassigned;
        mysheet:=Unassigned;
        ProcessPanel.Visible:=FALSE;
      end;
    end;
  end;
end;

procedure TMainMenu_Form.DailySupplierOrderCostSummary1Click(
  Sender: TObject);
var
  z:integer;
begin
  //get daily order summary
  DateSelectDlg:=TDateSelectDlg.Create(self);
  DateSelectDlg.DoSupplier:=True;
  DateSelectDlg.JustStart:=TRUE;
  DateSelectDlg.execute;
  if not DateSelectDlg.Cancel then
  begin
    with Data_Module.Inv_DataSet do
    begin
      try
        ProcessPanel.Visible:=TRUE;
        application.ProcessMessages;
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'dbo.REPORT_DailySupplierOrdersCost;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@FromDate';
        Parameters.ParamValues['@FromDate'] := FormatdateTime('yyyymmdd',DateSelectDlg.FromDate);;
        Parameters.AddParameter.Name := '@Supplier';
        Parameters.ParamValues['@Supplier'] := DateSelectDlg.Supplier;
        Open;
        excel := createOleObject('Excel.Application');
        excel.visible := False;
        excel.DisplayAlerts := False;
        excel.workbooks.open(Data_Module.TemplateDir+'ReportTemplate.xls');
        mysheet := excel.workSheets[1];
        mysheet.cells[1,1].value:='Daily Supplier Order';

        mysheet.cells[3,1].value:='Supplier';
        mysheet.Columns[1].ColumnWidth:=28;

        mysheet.cells[3,2].value:='Invoice No';
        mysheet.Columns[2].ColumnWidth:=28;

        mysheet.cells[3,3].value:='Ship Date';
        mysheet.Columns[3].ColumnWidth:=11;
        mysheet.Columns[3].NumberFormat := 'yy/mm/dd';

        mysheet.cells[3,4].value:='RENBAN';
        mysheet.Columns[4].ColumnWidth:=11;

        mysheet.cells[3,5].value:='FRS';
        mysheet.Columns[5].ColumnWidth:=11;
        mysheet.Columns[5].HorizontalAlignment := xlHAlignLeft;

        mysheet.cells[3,6].value:='PART NO';
        mysheet.Columns[6].ColumnWidth:=20;
        mysheet.Columns[6].HorizontalAlignment := xlHAlignLeft;
        mysheet.Columns[6].NumberFormat := '############';

        mysheet.cells[3,7].value:='QTY';
        mysheet.Columns[7].ColumnWidth:=11;

        mysheet.cells[3,8].value:='Unit Price';
        mysheet.Columns[8].ColumnWidth:=11;
        mysheet.Columns[8].NumberFormat := '$#,##0.0000_);[Red]($#,##0.0000)';

        mysheet.cells[3,9].value:='Total Amount';
        mysheet.Columns[9].ColumnWidth:=14;
        mysheet.Columns[9].NumberFormat := '$#,##0.0000_);[Red]($#,##0.0000)';

        //mysheet.cells[3,2].value:='Order Date';
        //mysheet.Columns[2].ColumnWidth:=11;
        //mysheet.Columns[2].NumberFormat := 'yy/mm/dd';

        //Columns("B:B").Select
        //Selection.NumberFormat = "mm/dd/yy"

        z:=4;
        while not EOF do
        begin
          mysheet.Cells[z,1].value := fieldbyname('vc_supplier_name').AsString;
          //mysheet.Cells[z,2].value := StrToDate(copy(fieldbyname('vc_order_date').AsString,5,2)+'/'+copy(fieldbyname('vc_order_date').AsString,7,2)+'/'+copy(fieldbyname('vc_order_date').AsString,1,4));

          if fieldbyname('SHIPPING').AsString <> '' then
            mysheet.Cells[z,3].value := StrToDate(copy(fieldbyname('SHIPPING').AsString,5,2)+'/'+copy(fieldbyname('SHIPPING').AsString,7,2)+'/'+copy(fieldbyname('SHIPPING').AsString,1,4))
          else
            mysheet.Cells[z,3].value := '';

          mysheet.Cells[z,4].value := fieldbyname('vc_renban_number').AsString;
          mysheet.Cells[z,5].value := fieldbyname('vc_frs_number').AsString;
          mysheet.Cells[z,6].value := fieldbyname('vc_part_number').AsString;
          mysheet.Cells[z,7].value := fieldbyname('in_qty').AsString;
          mysheet.Cells[z,8].value := fieldbyname('mo_part_cost').AsString;
          mysheet.Cells[z,9].value := fieldbyname('total').AsString;
          INC(z);
          next;
        end;
        excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\DailySupplierOrder'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
        ProcessPanel.Visible:=FALSE;
        if MessageDlg('Print this report?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
          mysheet.PrintOut (EmptyParam, EmptyParam, EmptyParam,
                            EmptyParam, EmptyParam, EmptyParam, EmptyParam);
        end;
        excel.Workbooks.Close;
        excel.Quit;
        excel:=Unassigned;
        mysheet:=Unassigned;
        Data_Module.LogActLog('REPORT','Do Daily Supplier Order');
        ShowMessage('Report Daily Supplier Order');
      except
        on e:exception do
        begin
          Data_Module.LogActLog('ERROR','Failed on Daily Supplier Order, '+e.Message);
          ShowMessage('Failed on Daily Supplier Order, '+e.Message);
          excel.Workbooks.Close;
          excel.Quit;
          excel:=Unassigned;
          mysheet:=Unassigned;
          ProcessPanel.Visible:=FALSE;
        end;
      end;
    end;
  end;
  DateSelectDlg.Free;
end;


procedure TMainMenu_Form.DailySupplierOrders1Click(Sender: TObject);
var
  z:integer;
begin
  //get daily order summary
  DateSelectDlg:=TDateSelectDlg.Create(self);
  DateSelectDlg.DoSupplier:=True;
  DateSelectDlg.JustStart:=TRUE;
  DateSelectDlg.execute;
  if not DateSelectDlg.Cancel then
  begin
    with Data_Module.Inv_DataSet do
    begin
      try
        ProcessPanel.Visible:=TRUE;
        application.ProcessMessages;
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'dbo.REPORT_DailySupplierOrders;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@FromDate';
        Parameters.ParamValues['@FromDate'] := FormatdateTime('yyyymmdd',DateSelectDlg.FromDate);;
        Parameters.AddParameter.Name := '@Supplier';
        Parameters.ParamValues['@Supplier'] := DateSelectDlg.Supplier;
        Open;
        excel := createOleObject('Excel.Application');
        excel.visible := False;
        excel.DisplayAlerts := False;
        excel.workbooks.open(Data_Module.TemplateDir+'ReportTemplate.xls');
        mysheet := excel.workSheets[1];
        mysheet.cells[1,1].value:='Daily Supplier Order';

        mysheet.cells[3,1].value:='Supplier';
        mysheet.Columns[1].ColumnWidth:=28;

        mysheet.cells[3,2].value:='Order Date';
        mysheet.Columns[2].ColumnWidth:=11;
        mysheet.Columns[2].NumberFormat := 'yy/mm/dd';

        mysheet.cells[3,3].value:='Ship Date';
        mysheet.Columns[3].ColumnWidth:=11;
        mysheet.Columns[3].NumberFormat := 'yy/mm/dd';
    //Columns("B:B").Select
    //Selection.NumberFormat = "mm/dd/yy"

        mysheet.cells[3,4].value:='RENBAN';
        mysheet.Columns[4].ColumnWidth:=11;

        mysheet.cells[3,5].value:='FRS';
        mysheet.Columns[5].ColumnWidth:=11;
        mysheet.Columns[5].HorizontalAlignment := xlHAlignLeft;

        mysheet.cells[3,6].value:='PART NO';
        mysheet.Columns[6].ColumnWidth:=20;
        mysheet.Columns[6].HorizontalAlignment := xlHAlignLeft;
        mysheet.Columns[6].NumberFormat := '############';


        mysheet.cells[3,7].value:='QTY';
        mysheet.Columns[7].ColumnWidth:=11;

        z:=4;
        while not EOF do
        begin
          mysheet.Cells[z,1].value := fieldbyname('vc_supplier_name').AsString;
          mysheet.Cells[z,2].value := StrToDate(copy(fieldbyname('vc_order_date').AsString,5,2)+'/'+copy(fieldbyname('vc_order_date').AsString,7,2)+'/'+copy(fieldbyname('vc_order_date').AsString,1,4));

          if fieldbyname('SHIPPING').AsString <> '' then
            mysheet.Cells[z,3].value := StrToDate(copy(fieldbyname('SHIPPING').AsString,5,2)+'/'+copy(fieldbyname('SHIPPING').AsString,7,2)+'/'+copy(fieldbyname('SHIPPING').AsString,1,4))
          else
            mysheet.Cells[z,3].value := '';

          mysheet.Cells[z,4].value := fieldbyname('vc_renban_number').AsString;
          mysheet.Cells[z,5].value := fieldbyname('vc_frs_number').AsString;
          mysheet.Cells[z,6].value := fieldbyname('vc_part_number').AsString;
          mysheet.Cells[z,7].value := fieldbyname('in_qty').AsString;
          INC(z);
          next;
        end;
        excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\DailySupplierOrder'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
        ProcessPanel.Visible:=FALSE;
        if MessageDlg('Print this report?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
          mysheet.PrintOut (EmptyParam, EmptyParam, EmptyParam,
                            EmptyParam, EmptyParam, EmptyParam, EmptyParam);
        end;
        excel.Workbooks.Close;
        excel.Quit;
        excel:=Unassigned;
        mysheet:=Unassigned;
        Data_Module.LogActLog('REPORT','Do Daily Supplier Order');
        ShowMessage('Report Daily Supplier Order');
      except
        on e:exception do
        begin
          Data_Module.LogActLog('ERROR','Failed on Daily Supplier Order, '+e.Message);
          ShowMessage('Failed on Daily Supplier Order, '+e.Message);
          excel.Workbooks.Close;
          excel.Quit;
          excel:=Unassigned;
          mysheet:=Unassigned;
          ProcessPanel.Visible:=FALSE;
        end;
      end;
    end;
  end;
  DateSelectDlg.Free;
end;

procedure TMainMenu_Form.MonthlySupplierOrderCostSummary1Click(
  Sender: TObject);
var
  z:integer;
begin
  //get monthly order summary
  DateSelectDlg:=TDateSelectDlg.Create(self);
  DateSelectDlg.DoSupplier:=True;
  DateSelectDlg.execute;
  if not DateSelectDlg.Cancel then
  begin

    with Data_Module.Inv_DataSet do
    begin
      try
        ProcessPanel.Visible:=TRUE;
        application.ProcessMessages;
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'dbo.REPORT_MonthlySupplierOrdersCost;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@FromDate';
        Parameters.ParamValues['@FromDate'] := FormatdateTime('yyyymmdd',DateSelectDlg.FromDate);;
        Parameters.AddParameter.Name := '@ToDate';
        Parameters.ParamValues['@ToDate'] := FormatdateTime('yyyymmdd',DateSelectDlg.ToDate);;
        Parameters.AddParameter.Name := '@Supplier';
        Parameters.ParamValues['@Supplier'] := DateSelectDlg.Supplier;
        Open;
        excel := createOleObject('Excel.Application');
        excel.visible := False;
        excel.DisplayAlerts := False;
        excel.workbooks.open(Data_Module.TemplateDir+'ReportTemplate.xls');
        mysheet := excel.workSheets[1];
        mysheet.cells[1,1].value:='Supplier Order';

        mysheet.cells[3,1].value:='Supplier';
        mysheet.Columns[1].ColumnWidth:=28;

        mysheet.cells[3,2].value:='Invoice No';
        mysheet.Columns[2].ColumnWidth:=28;

        mysheet.cells[3,3].value:='Ship Date';
        mysheet.Columns[3].ColumnWidth:=11;
        mysheet.Columns[3].NumberFormat := 'yy/mm/dd';

        mysheet.cells[3,4].value:='RENBAN';
        mysheet.Columns[4].ColumnWidth:=11;

        mysheet.cells[3,5].value:='FRS';
        mysheet.Columns[5].ColumnWidth:=11;
        mysheet.Columns[5].HorizontalAlignment := xlHAlignLeft;

        mysheet.cells[3,6].value:='PART NO';
        mysheet.Columns[6].ColumnWidth:=20;
        mysheet.Columns[6].HorizontalAlignment := xlHAlignLeft;
        mysheet.Columns[6].NumberFormat := '############';

        mysheet.cells[3,7].value:='QTY';
        mysheet.Columns[7].ColumnWidth:=11;

        mysheet.cells[3,8].value:='Unit Price';
        mysheet.Columns[8].ColumnWidth:=11;
        mysheet.Columns[8].NumberFormat := '$#,##0.0000_);[Red]($#,##0.0000)';

        mysheet.cells[3,9].value:='Total Amount';
        mysheet.Columns[9].ColumnWidth:=14;
        mysheet.Columns[9].NumberFormat := '$#,##0.0000_);[Red]($#,##0.0000)';

        //mysheet.cells[3,2].value:='Order Date';
        //mysheet.Columns[2].ColumnWidth:=11;
        //mysheet.Columns[2].NumberFormat := 'yy/mm/dd';

        //Columns("B:B").Select
        //Selection.NumberFormat = "mm/dd/yy"

        z:=4;
        while not EOF do
        begin
          mysheet.Cells[z,1].value := fieldbyname('vc_supplier_name').AsString;
          //mysheet.Cells[z,2].value := StrToDate(copy(fieldbyname('vc_order_date').AsString,5,2)+'/'+copy(fieldbyname('vc_order_date').AsString,7,2)+'/'+copy(fieldbyname('vc_order_date').AsString,1,4));

          if fieldbyname('SHIPPING').AsString <> '' then
            mysheet.Cells[z,3].value := StrToDate(copy(fieldbyname('SHIPPING').AsString,5,2)+'/'+copy(fieldbyname('SHIPPING').AsString,7,2)+'/'+copy(fieldbyname('SHIPPING').AsString,1,4))
          else
            mysheet.Cells[z,3].value := '';

          mysheet.Cells[z,4].value := fieldbyname('vc_renban_number').AsString;
          mysheet.Cells[z,5].value := fieldbyname('vc_frs_number').AsString;
          mysheet.Cells[z,6].value := fieldbyname('vc_part_number').AsString;
          mysheet.Cells[z,7].value := fieldbyname('in_qty').AsString;
          mysheet.Cells[z,8].value := fieldbyname('mo_part_cost').AsString;
          mysheet.Cells[z,9].value := fieldbyname('total').AsString;
          INC(z);
          next;
        end;
        excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\SupplierOrder'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
        ProcessPanel.Visible:=FALSE;
        if MessageDlg('Print this report?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
          mysheet.PrintOut (EmptyParam, EmptyParam, EmptyParam,
                            EmptyParam, EmptyParam, EmptyParam, EmptyParam);
        end;
        excel.Workbooks.Close;
        excel.Quit;
        excel:=Unassigned;
        mysheet:=Unassigned;
        Data_Module.LogActLog('REPORT','Do Supplier Order');
        ShowMessage('Report Supplier Order');
      except
        on e:exception do
        begin
          Data_Module.LogActLog('ERROR','Failed on Supplier Order, '+e.Message);
          ShowMessage('Failed on Supplier Order, '+e.Message);
          excel.Workbooks.Close;
          excel.Quit;
          excel:=Unassigned;
          mysheet:=Unassigned;
          ProcessPanel.Visible:=FALSE;
        end;
      end;
    end;
  end;
  DateSelectDlg.Free;
end;

procedure TMainMenu_Form.MonthlyOrderSummaryClick(Sender: TObject);
var
  z:integer;
begin
  //get monthly order summary
  DateSelectDlg:=TDateSelectDlg.Create(self);
  DateSelectDlg.DoSupplier:=True;
  DateSelectDlg.execute;
  if not DateSelectDlg.Cancel then
  begin
    {MonthlySupplierOrderSummary_Form:=TMonthlySupplierOrderSummary_Form.Create(self);
    MonthlySupplierOrderSummary_Form.Fromdate:=FormatdateTime('yyyymmdd',DateSelectDlg.FromDate);
    MonthlySupplierOrderSummary_Form.Todate:=FormatdateTime('yyyymmdd',DateSelectDlg.ToDate);
    MonthlySupplierOrderSummary_Form.Execute;
    MonthlySupplierOrderSummary_Form.Free;}
    with Data_Module.Inv_DataSet do
    begin
      try
        ProcessPanel.Visible:=TRUE;
        application.ProcessMessages;
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'dbo.REPORT_MonthlySupplierOrders;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@FromDate';
        Parameters.ParamValues['@FromDate'] := FormatdateTime('yyyymmdd',DateSelectDlg.FromDate);;
        Parameters.AddParameter.Name := '@ToDate';
        Parameters.ParamValues['@ToDate'] := FormatdateTime('yyyymmdd',DateSelectDlg.ToDate);;
        Parameters.AddParameter.Name := '@Supplier';
        Parameters.ParamValues['@Supplier'] := DateSelectDlg.Supplier;
        Open;
        excel := createOleObject('Excel.Application');
        excel.visible := False;
        excel.DisplayAlerts := False;
        excel.workbooks.open(Data_Module.TemplateDir+'ReportTemplate.xls');
        mysheet := excel.workSheets[1];
        mysheet.cells[1,1].value:='Supplier Order';

        mysheet.cells[3,1].value:='Supplier';
        mysheet.Columns[1].ColumnWidth:=28;

        mysheet.cells[3,2].value:='Order Date';
        mysheet.Columns[2].ColumnWidth:=11;
        mysheet.Columns[2].NumberFormat := 'yy/mm/dd';

        mysheet.cells[3,3].value:='Ship Date';
        mysheet.Columns[3].ColumnWidth:=11;
        mysheet.Columns[3].NumberFormat := 'yy/mm/dd';
    //Columns("B:B").Select
    //Selection.NumberFormat = "mm/dd/yy"

        mysheet.cells[3,4].value:='RENBAN';
        mysheet.Columns[4].ColumnWidth:=11;

        mysheet.cells[3,5].value:='FRS';
        mysheet.Columns[5].ColumnWidth:=11;
        mysheet.Columns[5].HorizontalAlignment := xlHAlignLeft;

        mysheet.cells[3,6].value:='PART NO';
        mysheet.Columns[6].ColumnWidth:=20;
        mysheet.Columns[6].HorizontalAlignment := xlHAlignLeft;
        mysheet.Columns[6].NumberFormat := '############';


        mysheet.cells[3,7].value:='QTY';
        mysheet.Columns[7].ColumnWidth:=11;

        z:=4;
        while not EOF do
        begin
          mysheet.Cells[z,1].value := fieldbyname('vc_supplier_name').AsString;
          mysheet.Cells[z,2].value := StrToDate(copy(fieldbyname('VC_ORDER_DATE').AsString,5,2)+'/'+copy(fieldbyname('VC_ORDER_DATE').AsString,7,2)+'/'+copy(fieldbyname('VC_ORDER_DATE').AsString,1,4));
          mysheet.Cells[z,3].value := StrToDate(copy(fieldbyname('SHIPPING').AsString,5,2)+'/'+copy(fieldbyname('SHIPPING').AsString,7,2)+'/'+copy(fieldbyname('SHIPPING').AsString,1,4));
          mysheet.Cells[z,4].value := fieldbyname('vc_renban_number').AsString;
          mysheet.Cells[z,5].value := fieldbyname('vc_frs_number').AsString;
          mysheet.Cells[z,6].value := fieldbyname('vc_part_number').AsString;
          mysheet.Cells[z,7].value := fieldbyname('in_qty').AsString;
          INC(z);
          next;
        end;
        excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\SupplierOrder'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
        ProcessPanel.Visible:=FALSE;
        if MessageDlg('Print this report?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
          mysheet.PrintOut (EmptyParam, EmptyParam, EmptyParam,
                            EmptyParam, EmptyParam, EmptyParam, EmptyParam);
        end;
        excel.Workbooks.Close;
        excel.Quit;
        excel:=Unassigned;
        mysheet:=Unassigned;
        Data_Module.LogActLog('REPORT','Do Supplier Order');
        ShowMessage('Report Supplier Order');
      except
        on e:exception do
        begin
          Data_Module.LogActLog('ERROR','Failed on Supplier Order, '+e.Message);
          ShowMessage('Failed on Supplier Order, '+e.Message);
          excel.Workbooks.Close;
          excel.Quit;
          excel:=Unassigned;
          mysheet:=Unassigned;
          ProcessPanel.Visible:=FALSE;
        end;
      end;
    end;
  end;
  DateSelectDlg.Free;
end;

procedure TMainMenu_Form.MonthlyLogisticsClick(Sender: TObject);
var
  z:integer;
begin
  // get monthly logistics report
  DateSelectDlg:=TDateSelectDlg.Create(self);
  DateSelectDlg.DoLogistics:=TRUE;
  DateSelectDlg.execute;
  if not DateSelectDlg.Cancel then
  begin
    {MonthlyLogisticsOrderSummary_Form:= TMonthlyLogisticsOrderSummary_Form.Create(self);
    MonthlyLogisticsOrderSummary_Form.Fromdate:=FormatdateTime('yyyymmdd',DateSelectDlg.FromDate);
    MonthlyLogisticsOrderSummary_Form.Todate:=FormatdateTime('yyyymmdd',DateSelectDlg.ToDate);
    MonthlyLogisticsOrderSummary_Form.Execute;
    MonthlyLogisticsOrderSummary_Form.Free;}
    with Data_Module.Inv_DataSet do
    begin
      try
        ProcessPanel.Visible:=TRUE;
        application.ProcessMessages;
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'dbo.REPORT_MonthlyLogisticsOrders;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@FromDate';
        Parameters.ParamValues['@FromDate'] := FormatdateTime('yyyymmdd',DateSelectDlg.FromDate);;
        Parameters.AddParameter.Name := '@ToDate';
        Parameters.ParamValues['@ToDate'] := FormatdateTime('yyyymmdd',DateSelectDlg.ToDate);;
        Parameters.AddParameter.Name := '@Logistics';
        Parameters.ParamValues['@Logistics'] := DateSelectDlg.Logistics;
        Open;
        excel := createOleObject('Excel.Application');
        excel.visible := False;
        excel.DisplayAlerts := False;
        excel.workbooks.open(Data_Module.TemplateDir+'ReportTemplate.xls');
        mysheet := excel.workSheets[1];
        mysheet.cells[1,1].value:='Monthly Logistics';

        mysheet.cells[3,1].value:='Logistics';
        mysheet.Columns[1].ColumnWidth:=28;

        mysheet.cells[3,2].value:='Ship Date';
        mysheet.Columns[2].ColumnWidth:=11;
        mysheet.Columns[2].NumberFormat := 'yy/mm/dd';

        mysheet.cells[3,3].value:='RENBAN';
        mysheet.Columns[3].ColumnWidth:=11;
        mysheet.Columns[3].HorizontalAlignment := xlHAlignLeft;

        z:=4;
        while not EOF do
        begin
          mysheet.Cells[z,1].value := fieldbyname('vc_logistics_name').AsString;
          mysheet.Cells[z,2].value := StrToDate(copy(fieldbyname('SHIPPING').AsString,5,2)+'/'+copy(fieldbyname('SHIPPING').AsString,7,2)+'/'+copy(fieldbyname('SHIPPING').AsString,1,4));
          mysheet.Cells[z,3].value := fieldbyname('vc_renban_number').AsString;
          INC(z);
          next;
        end;
        excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\Logistics'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
        ProcessPanel.Visible:=FALSE;
        if MessageDlg('Print this report?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
          mysheet.PrintOut (EmptyParam, EmptyParam, EmptyParam,
                            EmptyParam, EmptyParam, EmptyParam, EmptyParam);
        end;
        excel.Workbooks.Close;
        excel.Quit;
        excel:=Unassigned;
        mysheet:=Unassigned;
        Data_Module.LogActLog('REPORT','Do Logistics');
        ShowMessage('Report Logistics');
      except
        on e:exception do
        begin
          Data_Module.LogActLog('ERROR','Failed on Logistics, '+e.Message);
          ShowMessage('Failed on Logistics, '+e.Message);
          excel.Workbooks.Close;
          excel.Quit;
          excel:=Unassigned;
          mysheet:=Unassigned;
          ProcessPanel.Visible:=FALSE;
        end;
      end;
    end;
  end;
  DateSelectDlg.Free;
end;

procedure TMainMenu_Form.Window_UploadInvoices_MenuItemClick(
  Sender: TObject);
begin
  Hide;
  UpBreakDown_Form:=TUpBreakDown_Form.Create(self);
  UpBreakDown_Form.BreakdownKind:=bInvoice;
  UpBreakDown_Form.Execute;
  UpBreakDown_Form.Free;
  Show;
end;

procedure TMainMenu_Form.MonthlySupplierInvoice1Click(Sender: TObject);
var
  z:integer;
begin
  // get monthly logistics report
  DateSelectDlg:=TDateSelectDlg.Create(self);
  DateSelectDlg.DoSupplier:=TRUE;
  DateSelectDlg.execute;
  if not DateSelectDlg.Cancel then
  begin
    with Data_Module.Inv_DataSet do
    begin
      try
        ProcessPanel.Visible:=TRUE;
        application.ProcessMessages;
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'dbo.REPORT_MonthlySupplierInvoices;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@FromDate';
        Parameters.ParamValues['@FromDate'] := FormatdateTime('yyyymmdd',DateSelectDlg.FromDate);;
        Parameters.AddParameter.Name := '@ToDate';
        Parameters.ParamValues['@ToDate'] := FormatdateTime('yyyymmdd',DateSelectDlg.ToDate);;
        Parameters.AddParameter.Name := '@Supplier';
        Parameters.ParamValues['@Supplier'] := DateSelectDlg.Supplier;
        Open;
        excel := createOleObject('Excel.Application');
        excel.visible := False;
        excel.DisplayAlerts := False;
        excel.workbooks.open(Data_Module.TemplateDir+'ReportTemplate.xls');
        mysheet := excel.workSheets[1];
        mysheet.cells[1,1].value:='Supplier Invoice';
        mysheet.PageSetup.PrintTitleRows := '$1:$3';

        mysheet.cells[3,1].value:='Supplier';
        mysheet.Columns[1].ColumnWidth:=28;

        mysheet.cells[3,2].value:='Invoice No';
        mysheet.Columns[2].ColumnWidth:=20;
        mysheet.Columns[2].HorizontalAlignment := xlHAlignLeft;
        mysheet.Columns[2].NumberFormat := '############';

        mysheet.cells[3,3].value:='Part No';
        mysheet.Columns[3].ColumnWidth:=20;
        mysheet.Columns[3].HorizontalAlignment := xlHAlignLeft;
        mysheet.Columns[3].NumberFormat := '############';

        mysheet.cells[3,4].value:='FRS';
        mysheet.Columns[4].ColumnWidth:=11;
        mysheet.Columns[4].HorizontalAlignment := xlHAlignLeft;

        mysheet.cells[3,5].value:='RENBAN';
        mysheet.Columns[5].ColumnWidth:=11;
        mysheet.Columns[5].HorizontalAlignment := xlHAlignLeft;

        mysheet.cells[3,6].value:='QTY';
        mysheet.Columns[6].ColumnWidth:=11;

        mysheet.cells[3,7].value:='Unit Price';
        mysheet.Columns[7].ColumnWidth:=11;
        mysheet.Columns[7].NumberFormat := '$#,##0.0000_);[Red]($#,##0.0000)';

        mysheet.cells[3,8].value:='Total Amount';
        mysheet.Columns[8].ColumnWidth:=14;
        mysheet.Columns[8].NumberFormat := '$#,##0.0000_);[Red]($#,##0.0000)';

        mysheet.cells[3,9].value:='Invoice Date';
        mysheet.Columns[9].ColumnWidth:=11;
        mysheet.Columns[9].NumberFormat := 'yy/mm/dd';

        mysheet.cells[3,10].value:='Due Date';
        mysheet.Columns[10].ColumnWidth:=11;
        mysheet.Columns[10].NumberFormat := 'yy/mm/dd';

        mysheet.Range['A3','J3'].Interior.ColorIndex:=Data_Module.XlColToInt('O');

        z:=4;
        while not EOF do
        begin
          mysheet.Cells[z,1].value := fieldbyname('vc_supplier_name').AsString;
          mysheet.Cells[z,2].value := fieldbyname('vc_invoice_number').AsString;
          mysheet.Cells[z,3].value := fieldbyname('vc_part_number').AsString;
          mysheet.Cells[z,4].value := fieldbyname('vc_frs_number').AsString;
          mysheet.Cells[z,5].value := fieldbyname('vc_renban_number').AsString;
          mysheet.Cells[z,6].value := fieldbyname('in_qty').AsString;
          mysheet.Cells[z,7].value := fieldbyname('money_unit_price').AsString;
          mysheet.Cells[z,8].value := fieldbyname('money_total_amount').AsString;
          mysheet.Cells[z,9].value := StrToDate(copy(fieldbyname('invoicedate').AsString,5,2)+'/'+copy(fieldbyname('invoicedate').AsString,7,2)+'/'+copy(fieldbyname('invoicedate').AsString,1,4));
          mysheet.Cells[z,10].value := StrToDate(copy(fieldbyname('duedate').AsString,5,2)+'/'+copy(fieldbyname('duedate').AsString,7,2)+'/'+copy(fieldbyname('duedate').AsString,1,4));
          INC(z);
          next;
        end;
        excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\SupplierInvoice'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
        ProcessPanel.Visible:=FALSE;
        if MessageDlg('Print this report?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
          mysheet.PrintOut (EmptyParam, EmptyParam, EmptyParam,
                            EmptyParam, EmptyParam, EmptyParam, EmptyParam);
        end;
        excel.Workbooks.Close;
        excel.Quit;
        excel:=Unassigned;
        mysheet:=Unassigned;
        Data_Module.LogActLog('REPORT','Do Supplier Invoice');
        ShowMessage('Report Supplier Invoice');
      except
        on e:exception do
        begin
          Data_Module.LogActLog('ERROR','Failed on Supplier Invoice, '+e.Message);
          ShowMessage('Failed on Supplier Invoice, '+e.Message);
          excel.Workbooks.Close;
          excel.Quit;
          excel:=Unassigned;
          mysheet:=Unassigned;
          ProcessPanel.Visible:=FALSE;
        end;
      end;
    end;
  end;
  DateSelectDlg.Free;
end;

procedure TMainMenu_Form.EmptyContainerClick(Sender: TObject);
var
  z:integer;
  lastRenban:string;
begin
  // get monthly logistics report
  lastREnban:='XXX';
  DateSelectDlg:=TDateSelectDlg.Create(self);
  DateSelectDlg.DoSupplier:=TRUE;
  DateSelectDlg.DoLogistics:=TRUE;
  DateSelectDlg.execute;
  if not DateSelectDlg.Cancel then
  begin
    with Data_Module.Inv_DataSet do
    begin
      try
        ProcessPanel.Visible:=TRUE;
        application.ProcessMessages;
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'dbo.REPORT_EmptyContainer;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@FromDate';
        Parameters.ParamValues['@FromDate'] := FormatdateTime('yyyymmdd',DateSelectDlg.FromDate);;
        Parameters.AddParameter.Name := '@ToDate';
        Parameters.ParamValues['@ToDate'] := FormatdateTime('yyyymmdd',DateSelectDlg.ToDate);;
        Parameters.AddParameter.Name := '@Supplier';
        Parameters.ParamValues['@Supplier'] := DateSelectDlg.Supplier;
        Parameters.AddParameter.Name := '@Logistics';
        Parameters.ParamValues['@Logistics'] := DateSelectDlg.Logistics;
        Open;
        excel := createOleObject('Excel.Application');
        excel.visible := False;
        excel.DisplayAlerts := False;
        excel.workbooks.open(Data_Module.TemplateDir+'ReportTemplate.xls');
        mysheet := excel.workSheets[1];
        mysheet.cells[1,1].value:='Empty Container';

        mysheet.cells[3,1].value:='Supplier';
        mysheet.Columns[1].ColumnWidth:=28;

        mysheet.cells[3,2].value:='Logistics';
        mysheet.Columns[2].ColumnWidth:=28;

        mysheet.cells[3,3].value:='RENBAN';
        mysheet.Columns[3].ColumnWidth:=11;

        mysheet.cells[3,4].value:='Lot Location';
        mysheet.Columns[4].ColumnWidth:=11;

        mysheet.cells[3,5].value:='Kanban';
        mysheet.Columns[5].ColumnWidth:=11;

        mysheet.cells[3,6].value:='Order Date';
        mysheet.Columns[6].ColumnWidth:=11;

        mysheet.cells[3,7].value:='In Transit';
        mysheet.Columns[7].ColumnWidth:=11;

        mysheet.cells[3,8].value:='Empty ';
        mysheet.Columns[8].ColumnWidth:=11;

        mysheet.cells[3,9].value:='Terminated';
        mysheet.Columns[9].ColumnWidth:=11;

        mysheet.Range['A3','I3'].Interior.ColorIndex:=Data_Module.XlColToInt('O');
        z:=4;
        while not EOF do
        begin
          if lastRenban <> fieldbyname('vc_renban_number').AsString then
          begin
            mysheet.Cells[z,1].value := fieldbyname('vc_supplier_name').AsString;
            mysheet.Cells[z,2].value := fieldbyname('vc_logistics_name').AsString;
            mysheet.Cells[z,3].value := fieldbyname('vc_renban_number').AsString;
            mysheet.Cells[z,4].value := fieldbyname('vc_PLANT_parking').AsString;
            if UPPERCASE(fieldbyname('vc_part_type').AsString) = 'TIRE' then
              mysheet.Cells[z,5].value := fieldbyname('vc_kanban_number').AsString
            else
              mysheet.Cells[z,5].value := '';
            mysheet.Cells[z,6].value := fieldbyname('orderdate').AsString;
            mysheet.Cells[z,7].value := fieldbyname('shipping').AsString;
            mysheet.Cells[z,8].value := fieldbyname('emptytrailer').AsString;
            mysheet.Cells[z,9].value := fieldbyname('terminated').AsString;
            INC(z);
            lastRenban:=fieldbyname('vc_renban_number').AsString;
          end;
          next;
        end;
        ProcessPanel.Visible:=FALSE;
        excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\EmptyContainer'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
        if MessageDlg('Print this report?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
          mysheet.PrintOut (EmptyParam, EmptyParam, EmptyParam,
                            EmptyParam, EmptyParam, EmptyParam, EmptyParam);
        end;
        excel.Workbooks.Close;
        excel.Quit;
        excel:=Unassigned;
        mysheet:=Unassigned;
        Data_Module.LogActLog('REPORT','Do Empty Container');
        ShowMessage('Report Complete');
      except
        on e:exception do
        begin
          Data_Module.LogActLog('ERROR','Failed on Empty Container, '+e.Message);
          ShowMessage('Failed on Empty Container, '+e.Message);
          excel.Workbooks.Close;
          excel.Quit;
          excel:=Unassigned;
          mysheet:=Unassigned;
          ProcessPanel.Visible:=FALSE;
        end;
      end;
    end;
  end;
  DateSelectDlg.Free;
end;

procedure TMainMenu_Form.ForecastPartSummary1Click(Sender: TObject);
var
  z,r,x:integer;
  firstWeek,lastWeek:integer;
  PartList:TStringList;
begin
  try
    ProcessPanel.Visible:=TRUE;
    application.ProcessMessages;
    excel := createOleObject('Excel.Application');
    excel.visible := False;
    //excel.visible := TRUE;
    excel.DisplayAlerts := False;
    excel.workbooks.open(Data_Module.TemplateDir+'ReportTemplate.xls');
    mysheet := excel.workSheets[1];

    PartList:=TStringList.Create;


    with Data_Module.Inv_DataSet do
    begin
      Close;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.REPORT_ForecastPartsSummary;1';
      Parameters.Clear;
      Open;

      // Do Headers
      mysheet.cells[1,1].value:='Forecast Parts Summary';

      mysheet.cells[3,1].value:='Part Number';
      mysheet.Columns[1].ColumnWidth:=18;
      mysheet.Columns[1].HorizontalAlignment := xlHAlignLeft;
      mysheet.Columns[1].NumberFormat := '############';

      mysheet.cells[3,2].value:='Part Description';
      mysheet.Columns[2].ColumnWidth:=40;
      mysheet.Columns[2].HorizontalAlignment := xlHAlignLeft;

      lastWeek:=-1;
      r:=3;
      While not EOF do
      begin
        if lastWeek <> FieldByName('IN_WEEK_NUMBER').AsInteger then
        begin
          // do week Number
          mysheet.cells[3,r].value:='Week '+IntTOStr(fieldbyname('IN_WEEK_NUMBER').AsInteger);
          mysheet.Columns[r].HorizontalAlignment := xlHAlignLeft;
          INC(r);
          lastWeek:=fieldbyname('IN_WEEK_NUMBER').AsInteger;
        end;
        // Fill PartList
        x:=PartList.IndexOf(trim(FieldByName('VC_PART_NUMBER').AsString));
        if x = -1  then
        begin
          PartList.Add(Trim(FieldByName('VC_PART_NUMBER').AsString))
        end;
        Next;
      end;
      First;


      // Do Part Numbers
      z:=4;
      for x:=0 to PartList.Count-1 do
      begin
        mysheet.cells[z,1].value:=PartList[x];
        INC(z);
      end;

      First;
      z:=4;
      while not EOF do
      begin
        x:=PartList.IndexOf(trim(FieldByName('VC_PART_NUMBER').AsString));
        if x <> -1  then
        begin
          mysheet.cells[z+x,2].value:=FieldByName('VC_PARTS_NAME').AsString;
        end;
        Next;
      end;

      // Fill data for weeks
      First;
      firstWeek:=FieldByName('IN_WEEK_NUMBER').AsInteger;
      r:=3;
      z:=4;
      while not EOF do
      begin
        if FieldByName('IN_WEEK_NUMBER').AsInteger <> firstweek then
        begin
          firstWeek:=FieldByName('IN_WEEK_NUMBER').AsInteger;
          inc(r);
        end;
        x:=PartList.IndexOf(trim(FieldByName('VC_PART_NUMBER').AsString));
        if x <> -1  then
        begin
          mysheet.cells[z+x,r].value:=IntToStr(fieldbyname('qty').AsInteger);
        end;
        Next;
      end;


      ProcessPanel.Visible:=FALSE;
      //DoBoarders(4,z,'A',Data_Module.IntToXLCol(i),TRUE);
      excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\ForecastPartsSummary'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
            ProcessPanel.Visible:=FALSE;

      if MessageDlg('Print this report?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin

        mysheet.PrintOut (EmptyParam, EmptyParam, EmptyParam,
                          EmptyParam, EmptyParam, EmptyParam, EmptyParam);
      end;
      excel.Workbooks.Close;
      excel.Quit;
      excel:=Unassigned;
      mysheet:=Unassigned;
      Data_Module.LogActLog('REPORT','Do Forecast Parts Summary');
      ShowMessage('Forecast Parts Summary Report Complete');
    end;
  except
    on e:exception do
    begin
      Data_Module.LogActLog('ERROR','Failed on Forecast Parts Summary, '+e.Message);
      ShowMessage('Failed on Forecast Parts Summary, '+e.Message);
      excel.Workbooks.Close;
      excel.Quit;
      excel:=Unassigned;
      mysheet:=Unassigned;
      ProcessPanel.Visible:=FALSE;
    end;
  end;
end;

procedure TMainMenu_Form.ForecastSummary1Click(Sender: TObject);
var
  z,r:integer;
  lastPart:string;
begin
  try
    ProcessPanel.Visible:=TRUE;
    application.ProcessMessages;
    excel := createOleObject('Excel.Application');
    excel.visible := False;
    //excel.visible := TRUE;
    excel.DisplayAlerts := False;
    excel.workbooks.open(Data_Module.TemplateDir+'ReportTemplate.xls');
    mysheet := excel.workSheets[1];


    with Data_Module.Inv_DataSet do
    begin
      Close;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.REPORT_ForecastSummary;1';
      Parameters.Clear;
      Open;

      // Do Headers
      mysheet.cells[1,1].value:='Forecast Assy Summary';

      mysheet.cells[3,1].value:='Assy Part Number';
      mysheet.Columns[1].ColumnWidth:=18;
      mysheet.Columns[1].HorizontalAlignment := xlHAlignLeft;
      mysheet.Columns[1].NumberFormat := '############';

      mysheet.cells[3,2].value:='Kanban';
      mysheet.Columns[2].HorizontalAlignment := xlHAlignLeft;

      lastPart:=FieldByName('VC_Part_NUMBER').AsString;
      r:=2;
      while lastPart =  FieldByName('VC_Part_NUMBER').AsString do
      begin
        // do week Number
        mysheet.cells[3,r].value:='Week '+IntTOStr(fieldbyname('IN_WEEK_NUMBER').AsInteger);
        mysheet.Columns[r].HorizontalAlignment := xlHAlignLeft;
        INC(r);
        Next;
      end;
      First;

      z:=4;
      while not EOF do
      begin
        lastPart:=FieldByName('VC_PART_NUMBER').AsString;
        mysheet.cells[z,1].value:=FieldByName('VC_PART_NUMBER').AsString;
        r:=2;
        while (lastPart =  FieldByName('VC_PART_NUMBER').AsString) and (not EOF) do
        begin
          // do week Number
          mysheet.cells[z,r].value:=IntTOStr(fieldbyname('IN_COUNT').AsInteger);
          INC(r);
          Next;
        end;
        INC(z);
      end;

      ProcessPanel.Visible:=FALSE;
      //DoBoarders(4,z,'A',Data_Module.IntToXLCol(i),TRUE);
      excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\ForecastAssySummary'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
      ProcessPanel.Visible:=FALSE;

      if MessageDlg('Print this report?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin

        mysheet.PrintOut (EmptyParam, EmptyParam, EmptyParam,
                          EmptyParam, EmptyParam, EmptyParam, EmptyParam);
      end;
      excel.Workbooks.Close;
      excel.Quit;
      excel:=Unassigned;
      mysheet:=Unassigned;
      Data_Module.LogActLog('REPORT','Do Forecast Assy Summary');
      ShowMessage('Forecast Assy Summary Report Complete');
    end;
  except
    on e:exception do
    begin
      Data_Module.LogActLog('ERROR','Failed on Forecast Assy Summary, '+e.Message);
      ShowMessage('Failed on Forecast Assy Summary, '+e.Message);
      excel.Workbooks.Close;
      excel.Quit;
      excel:=Unassigned;
      mysheet:=Unassigned;
      ProcessPanel.Visible:=FALSE;
    end;
  end;
end;

procedure TMainMenu_Form.ForecastvsUsage1Click(Sender: TObject);
var
  weekoffset,i,x,z,fillCols,sizeCount:integer;
  fDates:array[0..100] of Double;
  fForecast:array[0..100] of integer;
  fUsage:array[0..100] of integer;
  lastSize:string;
begin
  //get inventory report
  lastSize:='X';
  sizeCount:=1;
  fillCols:=0;
  DateSelectDlg:=TDateSelectDlg.Create(self);
  DateSelectDlg.DoSupplier:=FALSE;
  DateSelectDlg.DoLogistics:=FALSE;
  DateSelectDlg.DoPartNumber:=TRUE;
  DateSelectDlg.execute;
  if not DateSelectDlg.Cancel then
  begin
    try
      ProcessPanel.Visible:=TRUE;
      application.ProcessMessages;
      excel := createOleObject('Excel.Application');
      excel.visible := False;
      //excel.visible := TRUE;
      excel.DisplayAlerts := False;
      excel.workbooks.open(Data_Module.TemplateDir+'ReportTemplate.xls');
      mysheet := excel.workSheets[1];

      with Data_Module.ALC_StoredProc do
      begin
        Close;
        ProcedureName := 'dbo.AD_GetSpecialDate';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@BeginDate';
        Parameters.ParamValues['@BeginDate'] := formatdatetime('yyyy-mm-dd 06:00:00',DateSelectDlg.FromDate);
        Parameters.AddParameter.Name := '@EndDate';
        Parameters.ParamValues['@EndDate'] := formatdatetime('yyyy-mm-dd 05:00:00',DateSelectDlg.ToDate);
        Parameters.AddParameter.Name := '@Line';
        Parameters.ParamValues['@Line'] := '';
        Open;

        if recordcount > 0 then
        begin
          //check for holiday/overtime
          i:=3;
          for x:=0 to trunc(DateSelectDlg.ToDate-DateSelectDlg.FromDate) do
          begin
            if not eof then
            begin
              if fieldbyname('DATE').AsDateTime =  trunc(DateSelectDlg.FromDate+x) then
              begin
                if fieldbyname('Date Status Abrv').AsString = 'O' then
                begin
                  INC(fillCols);
                  INC(i);
                  mysheet.Cells[3,i].value := formatdatetime('m/dd',DateSelectDlg.FromDate+x);
                  fDates[i-4]:=DateSelectDlg.FromDate+x;
                end;
                next;
              end
              else
              begin
                if DayOfTheWeek(DateSelectDlg.FromDate+x) < 6 then
                begin
                  INC(fillCols);
                  INC(i);
                  mysheet.Cells[3,i].value := formatdatetime('m/dd',DateSelectDlg.FromDate+x);
                  fDates[i-4]:=DateSelectDlg.FromDate+x;
                end;
              end;
            end
            else
            begin
              if DayOfTheWeek(DateSelectDlg.FromDate+x) < 6 then
              begin
                INC(fillCols);
                INC(i);
                mysheet.Cells[3,i].value := formatdatetime('m/dd',DateSelectDlg.FromDate+x);
                fDates[i-4]:=DateSelectDlg.FromDate+x;
              end;
            end;
          end;
        end
        else
        begin
          //normal run
          i:=3;
          for x:=0 to trunc(DateSelectDlg.ToDate-DateSelectDlg.FromDate) do
          begin
            if DayOfTheWeek(DateSelectDlg.FromDate+x) < 6 then
            begin
              INC(fillCols);
              INC(i);
              mysheet.Cells[3,i].value := formatdatetime('m/dd',DateSelectDlg.FromDate+x);
              fDates[i-4]:=DateSelectDlg.FromDate+x;
            end;
          end;
        end;
      end;
      mysheet.Range['A3',Data_Module.IntToXLCol(i)+'3'].Interior.ColorIndex:=Data_Module.XlColToInt('O');
      mysheet.Columns[i+1].ColumnWidth:=18;

      with Data_Module.Inv_DataSet do
      begin
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'dbo.REPORT_LogicalInventory;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@PartNo';
        Parameters.ParamValues['@PartNo'] := DateSelectDlg.Partnumber;
        Open;
        mysheet.cells[1,1].value:='Forecast vs Usage';

        mysheet.cells[3,1].value:='Size';
        mysheet.Columns[1].ColumnWidth:=10;

        mysheet.cells[3,2].value:='Part Number';
        mysheet.Columns[2].ColumnWidth:=18;
        mysheet.Columns[2].HorizontalAlignment := xlHAlignLeft;
        mysheet.Columns[2].NumberFormat := '############';

        mysheet.cells[3,3].value:='Kanban';
        mysheet.Columns[3].HorizontalAlignment := xlHAlignLeft;

        z:=4;
        while not EOF do
        begin
          if lastSize<>FieldByName('VC_SIZE_CODE').AsString then
          begin
            //check last size for multiple parts
            if (sizeCount>1) and (lastSize<>'') then
            begin
              //mulitple do a summary
              for x:=0 to fillcols-1 do
              begin
                mysheet.cells[z,4+x].value:=fForecast[x];
              end;
              mysheet.cells[z,4+fillcols].value:='Forecast Total';
              INC(z);

              for x:=0 to fillcols-1 do
              begin
                mysheet.cells[z,4+x].value:=fUsage[x];
              end;
              mysheet.cells[z,4+fillcols].value:='Usage Total';
              INC(z);
            end
            else
            begin
              if sizeCount>1 then
              begin
                mysheet.cells[z-2,4+fillcols].value:='Forecast Total';
                mysheet.cells[z-1,4+fillcols].value:='Forecast Total';
              end
              else
              begin
                if lastSize<>'X' then
                  mysheet.cells[z-1,4+fillcols].value:='Forecast Total';
              end;
              if lastSize<>'X' then
              begin
                for x:=0 to fillcols-1 do
                begin
                  mysheet.cells[z,4+x].value:=fUsage[x];
                end;
                mysheet.cells[z,4+fillcols].value:='Usage Total';
                INC(z);
              end;
            end;

            sizeCount:=0;
            for x:=0 to fillcols do
            begin
              fForecast[x]:=0;
              fUsage[x]:=0;
            end;


            if lastSize<>'X' then
              INC(z);
            if FieldByName('VC_SIZE_CODE').AsString = '' then
              mysheet.cells[z,1].value:='<No Size>'
            else
              mysheet.cells[z,1].value:=FieldByName('VC_SIZE_CODE').AsString;

            lastSize:=FieldByName('VC_SIZE_CODE').AsString;
          end;

          mysheet.cells[z,2].value:=FieldByName('VC_PART_NUMBER').AsString;
          mysheet.cells[z,3].value:=FieldByName('VC_KANBAN_NUMBER').AsString;

          with data_module.INV_Forecast_DataSet do
          begin

            if Data_Module.fiUseFirstProductionDay.AsBoolean then
            begin
              Close;
              CommandType := CmdStoredProc;
              CommandText := 'dbo.SELECT_FirstProductionDay;1';
              Parameters.Clear;
              Parameters.AddParameter.Name := '@ProdYear';
              Parameters.ParamValues['@ProdYear'] := formatdatetime('yyyy',now);
              Open;

              weekoffset:= FieldByName('First Week Number').AsInteger-1;
            end
            else
            begin
              weekoffset:=0;
            end;


            for x:=0 to fillcols-1 do
            begin
              // Get Forecast for day
              Close;
              CommandType := CmdStoredProc;
              CommandText := 'dbo.SELECT_ForecastPartNumberWeek;1';
              Parameters.Clear;
              Parameters.AddParameter.Name := '@WeekNo';



              if WeekOfTheYear(fDates[x])-weekoffset < 1 then
                Parameters.ParamValues['@WeekNo'] := 54-WeekOfTheYear(fDates[x])
              else
                Parameters.ParamValues['@WeekNo'] := WeekOfTheYear(fDates[x])-weekoffset;

              //Parameters.ParamValues['@WeekNo'] := WeekOfTheYear(fDates[x]);
              Parameters.AddParameter.Name := '@DayNo';
              Parameters.ParamValues['@DayNo'] := DayOfTheWeek(fDates[x]);
              Parameters.AddParameter.Name := '@PartNo';
              Parameters.ParamValues['@PartNo'] := Data_Module.Inv_DataSet.FieldByName('VC_PART_NUMBER').AsString;//Data_Module.Inv_StoredProc.FieldByName('VC_PART_NUMBER').Value;
              Open;
              mysheet.Cells[z,4+x].value  := fieldbyname('Qty').AsInteger;
              fForecast[x]:=fForecast[x]+fieldbyname('Qty').AsInteger;
              //Get Usage
              Close;
              CommandType := CmdStoredProc;
              CommandText := 'dbo.SELECT_UsageDay;1';
              Parameters.Clear;
              Parameters.AddParameter.Name := '@Date';
              Parameters.ParamValues['@Date'] := formatdatetime('yyyymmdd',fDates[x]);
              Parameters.AddParameter.Name := '@PartNo';
              Parameters.ParamValues['@PartNo'] := Data_Module.Inv_DataSet.FieldByName('VC_PART_NUMBER').AsString;//Data_Module.Inv_StoredProc.FieldByName('VC_PART_NUMBER').Value;
              Open;
              fUsage[x]:=fUsage[x]+fieldbyname('Qty').asInteger;
            end;
          end;

          next;
          INC(z);
          INC(sizeCount);
        end;

        //Last One
        if (sizeCount>1) and (lastSize<>'') then
        begin
          //mulitple do a summary
          for x:=0 to fillcols-1 do
          begin
            mysheet.cells[z,4+x].value:=fForecast[x];
          end;
          mysheet.cells[z,4+fillcols].value:='Forecast Total';
          INC(z);
          for x:=0 to fillcols-1 do
          begin
            mysheet.cells[z,4+x].value:=fUsage[x];
          end;
          mysheet.cells[z,4+fillcols].value:='Usage Total';
          INC(z);
        end
        else
        begin
          if sizeCount>1 then
          begin
            mysheet.cells[z-2,4+fillcols].value:='Forecast Total';
            mysheet.cells[z-1,4+fillcols].value:='Forecast Total';
          end
          else
          begin
            if lastSize<>'X' then
              mysheet.cells[z-1,4+fillcols].value:='Forecast Total';
          end;
          if lastSize<>'X' then
          begin
            for x:=0 to fillcols-1 do
            begin
              mysheet.cells[z,4+x].value:=fUsage[x];
            end;
            mysheet.cells[z,4+fillcols].value:='Usage Total';
            INC(z);
          end;
        end;


        ProcessPanel.Visible:=FALSE;
        DoBoarders(4,z,'A',Data_Module.IntToXLCol(i),TRUE);
        excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\ForecastvsUsage'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');

        ProcessPanel.Visible:=FALSE;

        if MessageDlg('Print this report?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin

          mysheet.PrintOut (EmptyParam, EmptyParam, EmptyParam,
                            EmptyParam, EmptyParam, EmptyParam, EmptyParam);
        end;
        excel.Workbooks.Close;
        excel.Quit;
        excel:=Unassigned;
        mysheet:=Unassigned;
        Data_Module.LogActLog('REPORT','Do Forecast vs Usage');
        ShowMessage('Report Complete');
      end;
    except
      on e:exception do
      begin
        Data_Module.LogActLog('ERROR','Failed on Forecast vs Usage, '+e.Message);
        ShowMessage('Failed on Forecast vs Usage, '+e.Message);
        excel.Workbooks.Close;
        excel.Quit;
        excel:=Unassigned;
        mysheet:=Unassigned;
        ProcessPanel.Visible:=FALSE;
      end;
    end;
  end;
  DateSelectDlg.Free;
end;

procedure TMainMenu_Form.PastDueFRS1Click(Sender: TObject);
var
  z:integer;
begin
  with Data_Module.Inv_DataSet do
  begin
    try
      ProcessPanel.Visible:=TRUE;
      application.ProcessMessages;
      Close;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.REPORT_LATEFRS;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@FRSNumber';
      Parameters.ParamValues['@FRSNumber'] := FormatdateTime('yyyymmdd',now);
      Open;

      if recordcount > 0 then
      begin

        excel := createOleObject('Excel.Application');
        excel.visible := False;
        excel.DisplayAlerts := False;
        excel.workbooks.open(Data_Module.TemplateDir+'ReportTemplate.xls');
        mysheet := excel.workSheets[1];
        mysheet.cells[1,1].value:='Past Due FRS';

        mysheet.cells[3,1].value:='FRS Number';
        mysheet.Columns[1].ColumnWidth:=11;

        mysheet.cells[3,2].value:='RENBAN';
        mysheet.Columns[2].ColumnWidth:=11;

        mysheet.cells[3,3].value:='Kanban';
        mysheet.Columns[3].ColumnWidth:=10;

        mysheet.cells[3,4].value:='Order Date';
        mysheet.Columns[4].ColumnWidth:=17;

        mysheet.cells[3,5].value:='Shipped';
        mysheet.Columns[5].ColumnWidth:=17;

        z:=4;
        while not eof do
        begin
          mysheet.Cells[z,1].value := fieldbyname('VC_FRS_NUMBER').AsString;
          mysheet.Cells[z,2].value := fieldbyname('VC_RENBAN_NUMBER').AsString;
          mysheet.Cells[z,3].value := fieldbyname('VC_KANBAN_NUMBER').AsString;
          mysheet.Cells[z,4].value := fieldbyname('VC_ORDER_DATE').AsString;
          mysheet.Cells[z,5].value := fieldbyname('VC_STATUS_SUPPLIER_SHIPPING').AsString;
          INC(z);
          next;
        end;


        ProcessPanel.Visible:=FALSE;
        excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\PastDueFRS'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
        if MessageDlg('Print this report?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin

          mysheet.PrintOut (EmptyParam, EmptyParam, EmptyParam,
                            EmptyParam, EmptyParam, EmptyParam, EmptyParam);
        end;
        excel.Workbooks.Close;
        excel.Quit;
        excel:=Unassigned;
        mysheet:=Unassigned;
        Data_Module.LogActLog('REPORT','Do Past Due FRS');
        ShowMessage('Report Complete');
      end
      else
      begin
        ShowMessage('No past due records');
      end;
      ProcessPanel.Visible:=FALSE;
    except
      on e:exception do
      begin
        Data_Module.LogActLog('ERROR','Failed on Past Due FRS, '+e.Message);
        ShowMessage('Failed on Past Due FRS, '+e.Message);
        excel.Workbooks.Close;
        excel.Quit;
        excel:=Unassigned;
        mysheet:=Unassigned;
        ProcessPanel.Visible:=FALSE;
      end;
    end;
  end;
end;

procedure TMainMenu_Form.FirstProductionDayoftheYear1Click(
  Sender: TObject);
begin
  try
    FirstProductionDay_Form:=TFirstProductionDay_Form.Create(self);
    FirstProductionDay_Form.Execute;
    FirstProductionDay_Form.Free;
  except
    on e:exception do
    begin
      ShowMessage('Unable to access First Production Day data, '+e.Message);
    end;
  end;
end;

procedure TMainMenu_Form.ManualForecast_ButtonClick(Sender: TObject);
begin
  try
    Hide;
    UpBreakDown_Form:=TUpBreakDown_Form.Create(self);
    UpBreakDown_Form.BreakdownKind:=bBuildout;
    UpBreakDown_Form.Execute;
    UpBreakDown_Form.Free;
    Show;
  except
    on e:exception do
    begin
      ShowMessage('Unable to access Buildout Forecast data, '+e.Message);
    end;
  end;
end;

procedure TMainMenu_Form.HotCallASN8561Click(Sender: TObject);
begin
  // Hot Call Entry Form
  try
    HotCallEntryForm:=THotCallEntryForm.Create(self);
    HotCallEntryForm.Execute;
  except
    on e:exception do
    begin
      ShowMessage('Unable to create Hotcall ASN, '+e.Message);
    end;
  end;
end;

procedure TMainMenu_Form.CreateASNClick(Sender: TObject);
begin
  // Do ASN Process
  try
    if Data_Module.fiGenerateEDI.AsBoolean then
    begin
      // Generate True EDI files
      ASNSelect_Form:=TASNSelect_Form.Create(self);
      ASNSelect_Form.Execute;
      ASNSelect_Form.Free;
    end
    else
    begin
      DateSelectDlg:=TDateSelectDlg.Create(self);
      DateSelectDlg.DoSupplier:=FALSE;
      DateSelectDlg.DoLogistics:=FALSE;
      DateSelectDlg.DoPartNumber:=FALSE;
      DateSelectDlg.execute;
      if not DateSelectDlg.Cancel then
      begin
        DailyBuildtotalForm:=TDailyBuildtotalForm.Create(self);
        DailyBuildTotalForm.FormMode:=fmASN;
        DailyBuildTotalForm.FromDate:=DateSelectDlg.FromDate;
        DailyBuildTotalForm.ToDate:=DateSelectDlg.ToDate;
        DailyBuildtotalForm.Execute;
        DailyBuildtotalForm.Free;
      end;
      DateSelectDlg.Free;
    end;
  except
    on e:exception do
    begin
      ShowMessage('Unable to create ASN, '+e.Message);
    end;
  end;
end;

procedure TMainMenu_Form.CreateINVOICEClick(Sender: TObject);
var
  EDI810:T810EDI;
  i:integer;
  fcf:TextFile;
  line:string;
begin
  // Do Invoice Process
  try
    if Data_Module.fiGenerateEDI.AsBoolean then
    begin
      // Generate True EDI files
      //
      //  Check for Normal 810 Creates (ASN complete)
      //
      Data_Module.EDI810DataSet.Close;
      Data_Module.EDI810DataSet.CommandText:='REPORT_EDI810';
      Data_Module.EDI810DataSet.Parameters.Clear;
      Data_Module.EDI810DataSet.Open;

      if Data_Module.EDI810DataSet.RecordCount > 0 then
      begin
        if MessageDlg('Create INVOICE Files now?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
          ProcessPanel.Visible:=TRUE;
          application.ProcessMessages;
          while not Data_Module.EDI810DataSet.Eof do
          begin
            Data_Module.ASN:=Data_Module.EDI810DataSet.FieldByName('ASNid').AsInteger;
            Data_Module.SiteDataSet.Close;
            Data_Module.SiteDataSet.Open;
            EDI810:=T810EDI.Create;
            EDI810.EIN:=Data_Module.SiteDataset.FieldByName('SiteEIN').AsInteger+1;
            if EDI810.Execute then
            begin
              Data_Module.EIN:=Data_Module.SiteDataSet.fieldByName('SiteEIN').AsInteger;
              AssignFile(fcf, Data_Module.fiEDIOut.AsString+'\810'+copy(EDI810.PickUpDate,5,4)+'.txt');
              Rewrite(fcf);

              // Loop through report and save
              for i:=0 to EDI810.EDIRecord.Count-1 do
              begin
                line:=EDI810.EDIRecord[i];
                Writeln(fcf,line);
              end;
              CloseFile(fcf);

              try
                Data_Module.UpdateReportCommand.CommandType:=cmdStoredProc;
                Data_Module.UpdateReportCommand.CommandText:='AD_UpdateEIN';
                Data_Module.UpdateReportCommand.Execute;

                if not Data_Module.InsertINVInfo then
                  ShowMessage('Failed on EIN update after EDI810 create');

              except
                on e:exception do
                begin
                  ShowMessage('Failed on create status after EDI810 create'+e.Message);
                end;
              end;

            end
            else
              ShowMessage('Unable to create EDI810 for ('+EDI810.PickUpDate+')');

            EDI810.Free;
          end;
          ProcessPanel.Visible:=FALSE;
          ShowMessage('Create EDI 810 files complete');
        end;
      end
      else
      begin
        ShowMessage('No Invoice files to create');
      end;
    end
    else
    begin
      // Generate TAI EDI files
      DateSelectDlg:=TDateSelectDlg.Create(self);
      DateSelectDlg.DoSupplier:=FALSE;
      DateSelectDlg.DoLogistics:=FALSE;
      DateSelectDlg.DoPartNumber:=FALSE;
      DateSelectDlg.execute;
      if not DateSelectDlg.Cancel then
      begin
        DailyBuildtotalForm:=TDailyBuildtotalForm.Create(self);
        DailyBuildTotalForm.FormMode:=fmINVOICE;
        DailyBuildTotalForm.FromDate:=DateSelectDlg.FromDate;
        DailyBuildTotalForm.ToDate:=DateSelectDlg.ToDate;
        DailyBuildtotalForm.Execute;
        DailyBuildtotalForm.Free;
      end;
    end;
  except
    on e:exception do
    begin
      ProcessPanel.Visible:=FALSE;
      ShowMessage('Unable to create INVOICE, '+e.Message);
    end;
  end;
end;

procedure TMainMenu_Form.ResendMarkedEDIsClick(Sender: TObject);
var
  i,y:integer;
  EDI856:T856EDI;
  fcf:TextFile;
  line:string;
begin
  try
    Data_Module.EDI856DataSet.Close;
    Data_Module.EDI856DataSet.Parameters.Clear;
    Data_Module.EDI856DataSet.Open;
    y:=1;
    if Data_Module.EDI856DataSet.RecordCount > 0 then
    begin
      if MessageDlg('Create ASN Files now?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
        while not Data_Module.EDI856DataSet.Eof do
        begin
          ProcessPanel.Visible:=TRUE;
          application.ProcessMessages;
          EDI856:=T856EDI.Create;
          if EDI856.Execute then
          begin
            if Data_Module.EDI856DataSet.FieldByName('StartSeq').AsString <> '-1' then
            begin
              Data_Module.LogActLog('EDI','Create EDI 856 filename ('+Data_Module.fiEDIOut.AsString+'\856'+EDI856.PickupDate+'.txt)');
              AssignFile(fcf, Data_Module.fiEDIOut.AsString+'\856'+copy(EDI856.PickupDate,5,4)+'.txt');
            end
            else
            begin
              Data_Module.LogActLog('EDI',Data_Module.fiEDIOut.AsString+'\8HC'+copy(EDI856.PickupDate,4,5)+InttoStr(y)+'.txt');
              AssignFile(fcf, Data_Module.fiEDIOut.AsString+'\8HC'+copy(EDI856.PickupDate,4,5)+InttoStr(y)+'.txt');
              INC(y);
            end;

            Rewrite(fcf);

            // Loop through report and save
            for i:=0 to EDI856.EDIRecord.Count-1 do
            begin
              line:=EDI856.EDIRecord[i];
              Writeln(fcf,line);
            end;
            CloseFile(fcf);
            Data_Module.LogActLog('EDI','Create EDI 856 for ('+EDI856.PickupDate+')');
          end
          else
          begin
            ShowMessage('Unable to create EDI856 for ('+EDI856.PickupDate+')');
            Data_Module.LogActLog('EDI','Unable to create EDI856 for ('+EDI856.PickupDate+')');
          end;

          EDI856.Free;
        end;
        try
          Data_Module.ASNStatus:='S';
          Data_Module.UpdateASNStatus;
        except
          on e:exception do
          begin
            ShowMessage('Failed on status update after EDI856 create, '+e.Message);
            Data_Module.LogActLog('ERROR','Failed on status update after EDI856 create, '+e.Message);
          end;
        end;
        ProcessPanel.Visible:=FALSE;
        ShowMessage('856 files complete');
      end;
    end
    else
    begin
      ShowMessage('No 856 records to create');
    end;
  except
    on e:exception do
    begin
      ShowMessage('Exception: Unable to create 856 files, '+e.Message);
      Data_Module.LogActLog('ERROR','Exception: Unable to create 856 files, '+e.Message);
    end;
  end;
end;

procedure TMainMenu_Form.POReportClick(Sender: TObject);
var
  lastPO:string;
  z:integer;
begin
  try
    DateSelectDlg:=TDateSelectDlg.Create(self);
    DateSelectDlg.DoSupplier:=FALSE;
    DateSelectDlg.DoLogistics:=FALSE;
    DateSelectDlg.DoPartNumber:=FALSE;
    DateSelectDlg.execute;
    if not DateSelectDlg.Cancel then
    begin
      with Data_Module.Inv_DataSet do
      begin
        try
          ProcessPanel.Visible:=TRUE;
          application.ProcessMessages;
          Close;
          CommandType := CmdStoredProc;
          CommandText := 'dbo.REPORT_PO;1';
          Parameters.Clear;
          Parameters.AddParameter.Name := '@BeginDate';
          Parameters.ParamValues['@BeginDate'] := FormatDateTime('yyyymmdd',DateSelectDlg.FromDate);
          Parameters.AddParameter.Name := '@EndDate';
          Parameters.ParamValues['@EndDate'] := FormatDateTime('yyyymmdd',DateSelectDlg.ToDate);
          Open;

          if recordcount > 0 then
          begin

            excel := createOleObject('Excel.Application');
            excel.visible := False;
            excel.DisplayAlerts := False;
            excel.workbooks.open(Data_Module.TemplateDir+'ReportTemplate.xls');
            mysheet := excel.workSheets[1];
            mysheet.cells[1,1].value:='PO Report';

            mysheet.cells[3,1].value:='Manifest No';
            mysheet.Columns[1].ColumnWidth:=11;

            mysheet.cells[3,2].value:='Period';
            mysheet.Columns[2].ColumnWidth:=25;

            mysheet.cells[3,3].value:='Pick Up Date';
            mysheet.Columns[3].ColumnWidth:=10;

            mysheet.cells[3,4].value:='Part Number';
            mysheet.Columns[4].ColumnWidth:=17;
            mysheet.Columns[4].NumberFormat := '############';

            mysheet.cells[3,5].value:='Qty Requested';
            mysheet.Columns[5].ColumnWidth:=17;

            mysheet.cells[3,6].value:='Qty Complete';
            mysheet.Columns[6].ColumnWidth:=17;

            mysheet.cells[3,7].value:='Ship Period U/P';
            mysheet.Columns[7].ColumnWidth:=17;
            mysheet.Columns[7].NumberFormat := '$#,##0.0000_);[Red]($#,##0.0000)';

            lastPO:='';
            z:=4;
            while not eof do
            begin
              if LastPO <> fieldbyname('VC_PO_NUMBER').AsString then
              begin
                mysheet.Cells[z,1].value := fieldbyname('VC_PO_NUMBER').AsString;
                mysheet.Cells[z,2].value := copy(fieldbyname('VC_PO_MONTH_START').AsString,5,2)+'/'+copy(fieldbyname('VC_PO_MONTH_START').AsString,7,2)+'/'+copy(fieldbyname('VC_PO_MONTH_START').AsString,1,4)+' to '+
                                            copy(fieldbyname('VC_PO_MONTH_END').AsString,5,2)+'/'+copy(fieldbyname('VC_PO_MONTH_END').AsString,7,2)+'/'+copy(fieldbyname('VC_PO_MONTH_END').AsString,1,4);
                mysheet.Cells[z,3].value := copy(fieldbyname('VC_PICK_UP_DATE').AsString,5,2)+'/'+copy(fieldbyname('VC_PICK_UP_DATE').AsString,7,2)+'/'+copy(fieldbyname('VC_PICK_UP_DATE').AsString,1,4);
                LastPO:=fieldbyname('VC_PO_NUMBER').AsString;
              end;
              mysheet.Cells[z,4].value := fieldbyname('VC_ASSY_PART_NUMBER_CODE').AsString;
              mysheet.Cells[z,5].value := fieldbyname('IN_PO_QTY').AsString;
              mysheet.Cells[z,6].value := fieldbyname('IN_PO_CHARGED').AsString;
              mysheet.Cells[z,7].value := fieldbyname('MO_ASSEMBLY_COST').AsString;
              INC(z);
              next;
            end;


            ProcessPanel.Visible:=FALSE;
            excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\POReport'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
            if MessageDlg('Print this report?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
            begin

              mysheet.PrintOut (EmptyParam, EmptyParam, EmptyParam,
                                EmptyParam, EmptyParam, EmptyParam, EmptyParam);
            end;
            excel.Workbooks.Close;
            excel.Quit;
            excel:=Unassigned;
            mysheet:=Unassigned;
            Data_Module.LogActLog('REPORT','Do PO Report');
            ShowMessage('Report Complete');
          end
          else
          begin
            ShowMessage('No past due records');
          end;
        except
          on e:exception do
          begin
            Data_Module.LogActLog('ERROR','Failed on PO Report, '+e.Message);
            ShowMessage('Failed on PO Report, '+e.Message);
            excel.Workbooks.Close;
            excel.Quit;
            excel:=Unassigned;
            mysheet:=Unassigned;
            ProcessPanel.Visible:=FALSE;
          end;
        end;
      end;
    end;
    DateSelectDlg.Free;
  except
    on e:exception do
    begin
      ShowMessage('Unable to create PO report, '+e.Message);
    end;
  end;
end;

procedure TMainMenu_Form.EDIUpload_ButtonClick(Sender: TObject);
begin
  // Upload EDI documents
  //
  // If forecast prompt user to use forecast function
  // If 824 compile error report for printing
  // If 997 log EDI transaction as complete
  EDIUpload_Form:=TEDIUpload_Form.Create(self);
  EDIUpload_Form.Show;
  Application.ProcessMessages;
  EDIUpload_Form.Execute;
  EDIUpload_Form.Free;
end;

procedure TMainMenu_Form.FormShow(Sender: TObject);
begin
  if Data_Module.fiGenerateEDI.AsBoolean then
  begin
    EDIUploadBox.Visible:=TRUE;
  end;
end;

procedure TMainMenu_Form.DailyShippingRangeTireWheelPartNumbersClick(
  Sender: TObject);
var
  z:integer;
begin
  try
    ProductionDateSelectDlg:=TProductionDateSelectDlg.Create(self);
    ProductionDateSelectDlg.INVOICE:=FALSE;
    ProductionDateSelectDlg.ASN:=FALSE;
    ProductionDateSelectDlg.Range:=TRUE;
    ProductionDateSelectDlg.execute;
    if not ProductionDateSelectDlg.Cancel then
    begin
      with Data_Module.Inv_DataSet do
      begin
        try
          ProcessPanel.Visible:=TRUE;
          application.ProcessMessages;
          Close;
          CommandType := CmdStoredProc;
          CommandText := 'dbo.REPORT_DailyShippingRange;1';
          Parameters.Clear;
          Parameters.AddParameter.Name := '@BeginPDate';
          Parameters.ParamValues['@BeginPDate'] := copy(ProductionDateSelectDlg.ProductionDate,1,4)+copy(ProductionDateSelectDlg.ProductionDate,6,2)+copy(ProductionDateSelectDlg.ProductionDate,9,2);
          Parameters.AddParameter.Name := '@EndPDate';
          Parameters.ParamValues['@EndPDate'] := copy(ProductionDateSelectDlg.ToProductionDate,1,4)+copy(ProductionDateSelectDlg.ToProductionDate,6,2)+copy(ProductionDateSelectDlg.ToProductionDate,9,2);
          Open;

          if recordcount > 0 then
          begin

            excel := createOleObject('Excel.Application');
            excel.visible := False;
            excel.DisplayAlerts := False;
            excel.workbooks.open(Data_Module.TemplateDir+'ReportTemplate.xls');
            mysheet := excel.workSheets[1];
            mysheet.cells[1,1].value:='Daily Shipping Range(Tire/Wheel Part Numbers)';

            mysheet.cells[2,1].NumberFormat := 'yyyy/mm/dd';
            mysheet.cells[2,1].value:=ProductionDateSelectDlg.ProductionDate;

            mysheet.cells[2,2].NumberFormat := 'yyyy/mm/dd';
            mysheet.cells[2,2].value:=ProductionDateSelectDlg.ToProductionDate;

            mysheet.cells[3,1].value:='Part Number';
            mysheet.Columns[1].ColumnWidth:=17;
            mysheet.Columns[1].NumberFormat := '############';

            mysheet.cells[3,2].value:='Part Description';
            mysheet.Columns[2].ColumnWidth:=30;

            mysheet.cells[3,3].value:='Qty';
            mysheet.Columns[3].ColumnWidth:=5;

            z:=4;
            while not eof do
            begin
              mysheet.Cells[z,1].value := fieldbyname('Part Number').AsString;
              mysheet.Cells[z,2].value := fieldbyname('Desc').AsString;
              mysheet.Cells[z,3].value := fieldbyname('PQTY').AsString;
              INC(z);
              next;
            end;


            ProcessPanel.Visible:=FALSE;

            excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\DailyShippingRangeTW'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
            if MessageDlg('Print this report?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
            begin

              mysheet.PrintOut (EmptyParam, EmptyParam, EmptyParam,
                                EmptyParam, EmptyParam, EmptyParam, EmptyParam);
            end;
            excel.Workbooks.Close;
            excel.Quit;
            excel:=Unassigned;
            mysheet:=Unassigned;
            Data_Module.LogActLog('REPORT','Do Daily Shipping Range Report');
            ShowMessage('Report Complete');
          end
          else
          begin
            ShowMessage('No daily records');
            ProcessPanel.Visible:=FALSE;
          end;
        except
          on e:exception do
          begin
            Data_Module.LogActLog('ERROR','Failed on Daily Shipping Range Report, '+e.Message);
            ShowMessage('Failed on Daily Shipping Range Report, '+e.Message);
            excel.Workbooks.Close;
            excel.Quit;
            excel:=Unassigned;
            mysheet:=Unassigned;
            ProcessPanel.Visible:=FALSE;
          end;
        end;
      end;
    end;
    ProductionDateSelectDlg.Free;
  except
    on e:exception do
    begin
      ShowMessage('Unable to create Daily Shipping Range Report, '+e.Message);
      ProcessPanel.Visible:=FALSE;
    end;
  end;
end;

procedure TMainMenu_Form.DailyShippingClick(Sender: TObject);
var
  z:integer;
begin
  try
    ProductionDateSelectDlg:=TProductionDateSelectDlg.Create(self);
    ProductionDateSelectDlg.INVOICE:=FALSE;
    ProductionDateSelectDlg.ASN:=FALSE;
    ProductionDateSelectDlg.execute;
    if not ProductionDateSelectDlg.Cancel then
    begin
      with Data_Module.Inv_DataSet do
      begin
        try
          ProcessPanel.Visible:=TRUE;
          application.ProcessMessages;
          Close;
          CommandType := CmdStoredProc;
          CommandText := 'dbo.REPORT_DailyShipping;1';
          Parameters.Clear;
          Parameters.AddParameter.Name := '@PDate';
          Parameters.ParamValues['@PDate'] := copy(ProductionDateSelectDlg.ProductionDate,1,4)+copy(ProductionDateSelectDlg.ProductionDate,6,2)+copy(ProductionDateSelectDlg.ProductionDate,9,2);
          Open;

          if recordcount > 0 then
          begin

            excel := createOleObject('Excel.Application');
            excel.visible := False;
            excel.DisplayAlerts := False;
            excel.workbooks.open(Data_Module.TemplateDir+'ReportTemplate.xls');
            mysheet := excel.workSheets[1];
            mysheet.cells[1,1].value:='Daily Shipping (Tire/Wheel Part Numbers)';

            mysheet.cells[2,1].NumberFormat := 'yyyy/mm/dd';
            mysheet.cells[2,1].value:=ProductionDateSelectDlg.ProductionDate;
            mysheet.cells[2,2].value:='Start Seq:'+fieldbyname('Start').AsString+'/End Seq:'+fieldbyname('End').AsString;
            mysheet.cells[2,3].value:='Vehicle Count:'+fieldbyname('Vehicle Count').AsString;


            mysheet.cells[3,1].value:='Part Number';
            mysheet.Columns[1].ColumnWidth:=17;
            mysheet.Columns[1].NumberFormat := '############';

            mysheet.cells[3,2].value:='Part Description';
            mysheet.Columns[2].ColumnWidth:=30;

            mysheet.cells[3,3].value:='Qty';
            mysheet.Columns[3].ColumnWidth:=5;

            z:=4;
            while not eof do
            begin
              mysheet.Cells[z,1].value := fieldbyname('Part Number').AsString;
              mysheet.Cells[z,2].value := fieldbyname('Desc').AsString;
              mysheet.Cells[z,3].value := fieldbyname('PQTY').AsString;
              INC(z);
              next;
            end;


            ProcessPanel.Visible:=FALSE;

            excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\DailyShippingTW'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
            if MessageDlg('Print this report?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
            begin

              mysheet.PrintOut (EmptyParam, EmptyParam, EmptyParam,
                                EmptyParam, EmptyParam, EmptyParam, EmptyParam);
            end;
            excel.Workbooks.Close;
            excel.Quit;
            excel:=Unassigned;
            mysheet:=Unassigned;
            Data_Module.LogActLog('REPORT','Do Daily Shipping Report');
            ShowMessage('Report Complete');
          end
          else
          begin
            ShowMessage('No daily records');
            ProcessPanel.Visible:=FALSE;
          end;
        except
          on e:exception do
          begin
            Data_Module.LogActLog('ERROR','Failed on Daily Shipping Report, '+e.Message);
            ShowMessage('Failed on Daily Shipping Report, '+e.Message);
            excel.Workbooks.Close;
            excel.Quit;
            excel:=Unassigned;
            mysheet:=Unassigned;
            ProcessPanel.Visible:=FALSE;
          end;
        end;
      end;
    end;
    ProductionDateSelectDlg.Free;
  except
    on e:exception do
    begin
      ShowMessage('Unable to create Daily Shipping Report, '+e.Message);
      ProcessPanel.Visible:=FALSE;
    end;
  end;
end;

procedure TMainMenu_Form.DailyASNReportClick(
  Sender: TObject);
var
  z:integer;
begin
  try
    ProductionDateSelectDlg:=TProductionDateSelectDlg.Create(self);
    ProductionDateSelectDlg.INVOICE:=FALSE;
    ProductionDateSelectDlg.ASN:=TRUE;
    ProductionDateSelectDlg.execute;
    if not ProductionDateSelectDlg.Cancel then
    begin
      with Data_Module.Inv_DataSet do
      begin
        try
          ProcessPanel.Visible:=TRUE;
          application.ProcessMessages;
          Close;
          CommandType := CmdStoredProc;
          CommandText := 'dbo.REPORT_DailyShippingAssy;1';
          Parameters.Clear;
          Parameters.AddParameter.Name := '@PDate';
          Parameters.ParamValues['@PDate'] := copy(ProductionDateSelectDlg.ProductionDate,1,4)+copy(ProductionDateSelectDlg.ProductionDate,6,2)+copy(ProductionDateSelectDlg.ProductionDate,9,2);
          Open;

          if recordcount > 0 then
          begin

            excel := createOleObject('Excel.Application');
            excel.visible := FALSE;
            excel.DisplayAlerts := False;
            excel.workbooks.open(Data_Module.TemplateDir+'ReportTemplate.xls');
            mysheet := excel.workSheets[1];
            mysheet.cells[1,1].value:='ASN (Assy Part Numbers)';


            mysheet.cells[2,1].NumberFormat := 'yyyy/mm/dd';
            mysheet.cells[2,1].value:=ProductionDateSelectDlg.ProductionDate;
            mysheet.cells[2,2].value:='Start Seq:'+fieldbyname('Start').AsString+'/End Seq:'+fieldbyname('End').AsString;
            mysheet.cells[2,3].value:='Vehicle Count:'+fieldbyname('Vehicle Count').AsString;


            mysheet.cells[3,1].value:='Part Number';
            mysheet.Columns[1].ColumnWidth:=17;
            mysheet.Columns[1].NumberFormat := '############';


            mysheet.cells[3,2].value:='Qty';
            mysheet.Columns[2].ColumnWidth:=30;

            z:=4;
            while not eof do
            begin
              mysheet.Cells[z,1].value := fieldbyname('Part Number').AsString;
              //mysheet.Cells[z,2].value := fieldbyname('Desc').AsString;
              mysheet.Cells[z,2].value := fieldbyname('PQTY').AsString;
              INC(z);
              next;
            end;


            ProcessPanel.Visible:=FALSE;

            excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\DailyShippingAssy'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
            if MessageDlg('Print this report?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
            begin

              mysheet.PrintOut (EmptyParam, EmptyParam, EmptyParam,
                                EmptyParam, EmptyParam, EmptyParam, EmptyParam);
            end;
            excel.Workbooks.Close;
            excel.Quit;
            excel:=Unassigned;
            mysheet:=Unassigned;
            Data_Module.LogActLog('REPORT','Do Daily Shipping Assy Report');
            ShowMessage('Report Complete');
          end
          else
          begin
            ShowMessage('No Daily records');
            ProcessPanel.Visible:=FALSE;
          end;
        except
          on e:exception do
          begin
            Data_Module.LogActLog('ERROR','Failed on Daily Shipping Assy Report, '+e.Message);
            ShowMessage('Failed on Daily Shipping Assy Report, '+e.Message);
            excel.Workbooks.Close;
            excel.Quit;
            excel:=Unassigned;
            mysheet:=Unassigned;
            ProcessPanel.Visible:=FALSE;
          end;
        end;
      end;
    end;
    ProductionDateSelectDlg.Free;
  except
    on e:exception do
    begin
      ShowMessage('Unable to create Daily Shipping Assy Report, '+e.Message);
      ProcessPanel.Visible:=FALSE;
    end;
  end;
end;

procedure TMainMenu_Form.ASNReportWithCostAssyPartNumbers1Click(
  Sender: TObject);
var
  z:integer;
begin
  try
    ProductionDateSelectDlg:=TProductionDateSelectDlg.Create(self);
    ProductionDateSelectDlg.INVOICE:=FALSE;
    ProductionDateSelectDlg.ASN:=TRUE;
    ProductionDateSelectDlg.execute;
    if not ProductionDateSelectDlg.Cancel then
    begin
      with Data_Module.Inv_DataSet do
      begin
        try
          ProcessPanel.Visible:=TRUE;
          application.ProcessMessages;
          Close;
          CommandType := CmdStoredProc;
          CommandText := 'dbo.REPORT_ASNWithCost;1';
          Parameters.Clear;
          Parameters.AddParameter.Name := '@PDate';
          Parameters.ParamValues['@PDate'] := copy(ProductionDateSelectDlg.ProductionDate,1,4)+copy(ProductionDateSelectDlg.ProductionDate,6,2)+copy(ProductionDateSelectDlg.ProductionDate,9,2);
          Open;

          if recordcount > 0 then
          begin

            excel := createOleObject('Excel.Application');
            excel.visible := FALSE;
            excel.DisplayAlerts := False;
            excel.workbooks.open(Data_Module.TemplateDir+'ReportTemplate.xls');
            mysheet := excel.workSheets[1];
            mysheet.cells[1,1].value:='ASN (Assy Part Numbers) with Cost';


            mysheet.cells[2,1].NumberFormat := 'yyyy/mm/dd';
            mysheet.cells[2,1].value:=ProductionDateSelectDlg.ProductionDate;
            mysheet.cells[2,2].value:='Start Seq:'+fieldbyname('Start').AsString+'/End Seq:'+fieldbyname('End').AsString;
            mysheet.cells[2,3].value:='Vehicle Count:'+fieldbyname('Vehicle Count').AsString;


            mysheet.cells[3,1].value:='Part Number';
            mysheet.Columns[1].ColumnWidth:=17;
            mysheet.Columns[1].NumberFormat := '############';


            mysheet.cells[3,2].value:='Qty';
            mysheet.Columns[2].ColumnWidth:=30;

            mysheet.cells[3,3].value:='Unit Cost';
            mysheet.Columns[3].ColumnWidth:=30;
            mysheet.Columns[3].NumberFormat := '$#,##0.0000_);[Red]($#,##0.0000)';

            mysheet.cells[3,4].value:='Item Total';
            mysheet.Columns[4].ColumnWidth:=30;
            mysheet.Columns[4].NumberFormat := '$#,##0.0000_);[Red]($#,##0.0000)';

            z:=4;
            while not eof do
            begin
              mysheet.Cells[z,1].value := fieldbyname('Part Number').AsString;
              //mysheet.Cells[z,2].value := fieldbyname('Desc').AsString;
              mysheet.Cells[z,2].value := fieldbyname('PQTY').AsString;
              INC(z);
              next;
            end;


            ProcessPanel.Visible:=FALSE;

            excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\ASNwithCost'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
            if MessageDlg('Print this report?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
            begin

              mysheet.PrintOut (EmptyParam, EmptyParam, EmptyParam,
                                EmptyParam, EmptyParam, EmptyParam, EmptyParam);
            end;
            excel.Workbooks.Close;
            excel.Quit;
            excel:=Unassigned;
            mysheet:=Unassigned;
            Data_Module.LogActLog('REPORT','Do ASN with Cost');
            ShowMessage('Report Complete');
          end
          else
          begin
            ShowMessage('No ASN with Cost records');
            ProcessPanel.Visible:=FALSE;
          end;
        except
          on e:exception do
          begin
            Data_Module.LogActLog('ERROR','Failed on ASN with Cost Report, '+e.Message);
            ShowMessage('Failed on ASN with Cost Report, '+e.Message);
            excel.Workbooks.Close;
            excel.Quit;
            excel:=Unassigned;
            mysheet:=Unassigned;
            ProcessPanel.Visible:=FALSE;
          end;
        end;
      end;
    end;
    ProductionDateSelectDlg.Free;
  except
    on e:exception do
    begin
      ShowMessage('Unable to create Daily Shipping Assy Report, '+e.Message);
      ProcessPanel.Visible:=FALSE;
    end;
  end;
end;

procedure TMainMenu_Form.MonthlyASNClick(
  Sender: TObject);
var
  z:integer;
begin
  try
    ProductionDateSelectDlg:=TProductionDateSelectDlg.Create(self);
    ProductionDateSelectDlg.INVOICE:=FALSE;
    ProductionDateSelectDlg.ASN:=TRUE;
    ProductionDateSelectDlg.Month:=TRUE;
    ProductionDateSelectDlg.execute;
    if not ProductionDateSelectDlg.Cancel then
    begin
      with Data_Module.Inv_DataSet do
      begin
        try
          ProcessPanel.Visible:=TRUE;
          application.ProcessMessages;
          Close;
          CommandType := CmdStoredProc;
          CommandText := 'dbo.REPORT_MonthlyShippingAssy;1';
          Parameters.Clear;
          Parameters.AddParameter.Name := '@PDate';
          Parameters.ParamValues['@PDate'] := copy(ProductionDateSelectDlg.ProductionDate,1,4)+copy(ProductionDateSelectDlg.ProductionDate,6,2);
          Open;

          if recordcount > 0 then
          begin

            excel := createOleObject('Excel.Application');
            excel.visible := FALSE;
            excel.DisplayAlerts := False;
            excel.workbooks.open(Data_Module.TemplateDir+'ReportTemplate.xls');
            mysheet := excel.workSheets[1];
            mysheet.cells[1,1].value:='Monthly ASN''s (Assy Part Numbers)';


            mysheet.cells[2,1].NumberFormat := 'yyyy/mm';
            mysheet.cells[2,1].value:=ProductionDateSelectDlg.ProductionDate;
            //mysheet.cells[2,2].value:='Start Seq:'+fieldbyname('Start').AsString+'/End Seq:'+fieldbyname('End').AsString;
            //mysheet.cells[2,3].value:='Vehicle Count:'+fieldbyname('Vehicle Count').AsString;


            mysheet.cells[3,1].value:='Part Number';
            mysheet.Columns[1].ColumnWidth:=17;
            mysheet.Columns[1].NumberFormat := '############';


            mysheet.cells[3,2].value:='Qty';
            mysheet.Columns[2].ColumnWidth:=10;

            mysheet.cells[3,3].value:='Production Date';
            mysheet.Columns[3].ColumnWidth:=20;

            z:=4;
            while not eof do
            begin
              mysheet.Cells[z,1].value := fieldbyname('Part Number').AsString;
              //mysheet.Cells[z,2].value := fieldbyname('Desc').AsString;
              mysheet.Cells[z,2].value := fieldbyname('PQTY').AsString;
              mysheet.Cells[z,3].value := copy(fieldbyname('PDate').AsString,1,4)+'/'+copy(fieldbyname('PDate').AsString,5,2)+'/'+copy(fieldbyname('PDate').AsString,7,2);
              INC(z);
              next;
            end;


            ProcessPanel.Visible:=FALSE;

            excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\DailyShippingAssy'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
            if MessageDlg('Print this report?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
            begin

              mysheet.PrintOut (EmptyParam, EmptyParam, EmptyParam,
                                EmptyParam, EmptyParam, EmptyParam, EmptyParam);
            end;
            excel.Workbooks.Close;
            excel.Quit;
            excel:=Unassigned;
            mysheet:=Unassigned;
            Data_Module.LogActLog('REPORT','Do Daily Shipping Assy Report');
            ShowMessage('Report Complete');
          end
          else
          begin
            ShowMessage('No Daily records');
            ProcessPanel.Visible:=FALSE;
          end;
        except
          on e:exception do
          begin
            Data_Module.LogActLog('ERROR','Failed on Daily Shipping Assy Report, '+e.Message);
            ShowMessage('Failed on Daily Shipping Assy Report, '+e.Message);
            excel.Workbooks.Close;
            excel.Quit;
            excel:=Unassigned;
            mysheet:=Unassigned;
            ProcessPanel.Visible:=FALSE;
          end;
        end;
      end;
    end;
    ProductionDateSelectDlg.Free;
  except
    on e:exception do
    begin
      ShowMessage('Unable to create Daily Shipping Assy Report, '+e.Message);
      ProcessPanel.Visible:=FALSE;
    end;
  end;
end;

procedure TMainMenu_Form.MonthlyINVOICEAssyPartNumbers1Click(
  Sender: TObject);
var
  z:integer;
begin
  try
    ProductionDateSelectDlg:=TProductionDateSelectDlg.Create(self);
    ProductionDateSelectDlg.INVOICE:=TRUE;
    ProductionDateSelectDlg.ASN:=FALSE;
    ProductionDateSelectDlg.Month:=TRUE;
    ProductionDateSelectDlg.execute;
    if not ProductionDateSelectDlg.Cancel then
    begin
      with Data_Module.Inv_DataSet do
      begin
        try
          ProcessPanel.Visible:=TRUE;
          application.ProcessMessages;
          Close;
          CommandType := CmdStoredProc;
          CommandText := 'dbo.REPORT_MonthlyINVOICESSummary;1';
          Parameters.Clear;
          Parameters.AddParameter.Name := '@PDate';
          Parameters.ParamValues['@PDate'] := copy(ProductionDateSelectDlg.ProductionDate,1,4)+
                                              copy(ProductionDateSelectDlg.ProductionDate,6,2)+
                                              copy(ProductionDateSelectDlg.ProductionDate,9,2);
                                              //copy(ProductionDateSelectDlg.ProductionDate,12,2)+
                                              //copy(ProductionDateSelectDlg.ProductionDate,15,2);
          Open;

          if recordcount > 0 then
          begin

            excel := createOleObject('Excel.Application');
            excel.visible := FALSE;
            excel.DisplayAlerts := False;
            excel.workbooks.open(Data_Module.TemplateDir+'ReportTemplate.xls');
            mysheet := excel.workSheets[1];
            mysheet.cells[1,1].value:='INVOICE (Assy Part Numbers)';


            mysheet.cells[2,1].value:='INVOICE DATE:'+ProductionDateSelectDlg.ProductionDate;

            mysheet.cells[3,1].value:='Part Number';
            mysheet.Columns[1].ColumnWidth:=17;
            mysheet.Columns[1].NumberFormat := '############';

            mysheet.cells[3,2].value:='Manifest Number';
            mysheet.Columns[2].ColumnWidth:=17;
            mysheet.Columns[2].NumberFormat := '@';

            mysheet.cells[3,3].value:='Qty';
            mysheet.Columns[3].ColumnWidth:=30;

            mysheet.cells[3,4].value:='Unit Cost';
            mysheet.Columns[4].ColumnWidth:=30;
            mysheet.Columns[4].NumberFormat := '$#,##0.0000_);[Red]($#,##0.0000)';

            mysheet.cells[3,5].value:='Item Total';
            mysheet.Columns[5].ColumnWidth:=30;
            mysheet.Columns[5].NumberFormat := '$#,##0.0000_);[Red]($#,##0.0000)';


            z:=4;
            while not eof do
            begin
              mysheet.Cells[z,1].value := fieldbyname('PartNumber').AsString;
              mysheet.Cells[z,2].value := fieldbyname('Manifest').AsString;
              mysheet.Cells[z,3].value := fieldbyname('ShipQty').AsString;
              mysheet.Cells[z,4].value := fieldbyname('UnitPrice').AsFloat;
              mysheet.Range['E'+IntToStr(z)].Formula  := '=C'+IntToStr(z)+'*D'+IntToStr(z);
              INC(z);
              next;
            end;

            INC(z);
            INC(z);
            mysheet.Cells[z,3].value := 'INVOICE TOTAL';
            mysheet.Range['E'+IntToStr(z)].Formula  := '=SUM(E4:E'+IntToStr(z-2)+')';

            ProcessPanel.Visible:=FALSE;

            excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\MonthlyINVOICESummary'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
            if MessageDlg('Print this report?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
            begin

              mysheet.PrintOut (EmptyParam, EmptyParam, EmptyParam,
                                EmptyParam, EmptyParam, EmptyParam, EmptyParam);
            end;
            excel.Workbooks.Close;
            excel.Quit;
            excel:=Unassigned;
            mysheet:=Unassigned;
            Data_Module.LogActLog('REPORT','Do Monthly INVOICE Summary Report');
            ShowMessage('Report Complete');
          end
          else
          begin
            ShowMessage('No INVOICE Summary records');
            ProcessPanel.Visible:=FALSE;
          end;
        except
          on e:exception do
          begin
            Data_Module.LogActLog('ERROR','Failed on Monthly INVOICE Summary Report, '+e.Message);
            ShowMessage('Failed on Monthly INVOICE Summary Report, '+e.Message);
            excel.Workbooks.Close;
            excel.Quit;
            excel:=Unassigned;
            mysheet:=Unassigned;
            ProcessPanel.Visible:=FALSE;
          end;
        end;
      end;
    end;
    ProductionDateSelectDlg.Free;
  except
    on e:exception do
    begin
      ProcessPanel.Visible:=FALSE;
      ShowMessage('Unable to create Monthly INVOICE Summary Report, '+e.Message);
    end;
  end;
end;

procedure TMainMenu_Form.INVOICEReportClick(
  Sender: TObject);
var
  z:integer;
begin
  try
    ProductionDateSelectDlg:=TProductionDateSelectDlg.Create(self);
    ProductionDateSelectDlg.INVOICE:=TRUE;
    ProductionDateSelectDlg.ASN:=FALSE;
    ProductionDateSelectDlg.execute;
    if not ProductionDateSelectDlg.Cancel then
    begin
      with Data_Module.Inv_DataSet do
      begin
        try
          ProcessPanel.Visible:=TRUE;
          application.ProcessMessages;
          Close;
          CommandType := CmdStoredProc;
          CommandText := 'dbo.REPORT_INVOICESSummary;1';
          Parameters.Clear;
          Parameters.AddParameter.Name := '@PDate';
          Parameters.ParamValues['@PDate'] := copy(ProductionDateSelectDlg.ProductionDate,1,4)+
                                              copy(ProductionDateSelectDlg.ProductionDate,6,2)+
                                              copy(ProductionDateSelectDlg.ProductionDate,9,2);
                                              //copy(ProductionDateSelectDlg.ProductionDate,12,2)+
                                              //copy(ProductionDateSelectDlg.ProductionDate,15,2);
          Open;

          if recordcount > 0 then
          begin

            excel := createOleObject('Excel.Application');
            excel.visible := FALSE;
            excel.DisplayAlerts := False;
            excel.workbooks.open(Data_Module.TemplateDir+'ReportTemplate.xls');
            mysheet := excel.workSheets[1];
            mysheet.cells[1,1].value:='INVOICE (Assy Part Numbers)';


            mysheet.cells[2,1].value:='INVOICE DATE:'+ProductionDateSelectDlg.ProductionDate;

            mysheet.cells[3,1].value:='Part Number';
            mysheet.Columns[1].ColumnWidth:=17;
            mysheet.Columns[1].NumberFormat := '############';

            mysheet.cells[3,2].value:='Manifest Number';
            mysheet.Columns[2].ColumnWidth:=17;

            mysheet.cells[3,3].value:='Qty';
            mysheet.Columns[3].ColumnWidth:=30;

            mysheet.cells[3,4].value:='Unit Cost';
            mysheet.Columns[4].ColumnWidth:=30;
            mysheet.Columns[4].NumberFormat := '$#,##0.0000_);[Red]($#,##0.0000)';

            mysheet.cells[3,5].value:='Item Total';
            mysheet.Columns[5].ColumnWidth:=30;
            mysheet.Columns[5].NumberFormat := '$#,##0.0000_);[Red]($#,##0.0000)';


            z:=4;
            while not eof do
            begin
              mysheet.Cells[z,1].value := fieldbyname('PartNumber').AsString;
              mysheet.Cells[z,2].value := fieldbyname('Manifest').AsString;
              mysheet.Cells[z,3].value := fieldbyname('ShipQty').AsString;
              mysheet.Cells[z,4].value := fieldbyname('UnitPrice').AsFloat;
              mysheet.Range['E'+IntToStr(z)].Formula  := '=C'+IntToStr(z)+'*D'+IntToStr(z);
              INC(z);
              next;
            end;

            INC(z);
            INC(z);
            mysheet.Cells[z,3].value := 'INVOICE TOTAL';
            mysheet.Range['E'+IntToStr(z)].Formula  := '=SUM(E4:E'+IntToStr(z-2)+')';

            ProcessPanel.Visible:=FALSE;

            excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\INVOICESummary'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
            if MessageDlg('Print this report?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
            begin

              mysheet.PrintOut (EmptyParam, EmptyParam, EmptyParam,
                                EmptyParam, EmptyParam, EmptyParam, EmptyParam);
            end;
            excel.Workbooks.Close;
            excel.Quit;
            excel:=Unassigned;
            mysheet:=Unassigned;
            Data_Module.LogActLog('REPORT','Do INVOICE Summary Report');
            ShowMessage('Report Complete');
          end
          else
          begin
            ShowMessage('No INVOICE Summary records');
            ProcessPanel.Visible:=FALSE;
          end;
        except
          on e:exception do
          begin
            Data_Module.LogActLog('ERROR','Failed on INVOICE Summary Report, '+e.Message);
            ShowMessage('Failed on INVOICE Summary Report, '+e.Message);
            excel.Workbooks.Close;
            excel.Quit;
            excel:=Unassigned;
            mysheet:=Unassigned;
            ProcessPanel.Visible:=FALSE;
          end;
        end;
      end;
    end;
    ProductionDateSelectDlg.Free;
  except
    on e:exception do
    begin
      ProcessPanel.Visible:=FALSE;
      ShowMessage('Unable to create INVOICE Summary Report, '+e.Message);
    end;
  end;
end;


procedure TMainMenu_Form.ForecastDetail1Click(Sender: TObject);
var
  z:integer;
  lastpn,lastmn:string;
begin
  with Data_Module.Inv_DataSet do
  begin
    try
      ProcessPanel.Visible:=TRUE;
      application.ProcessMessages;
      Close;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.REPORT_ForecastDetail;1';
      Parameters.Clear;
      Open;

      if recordcount > 0 then
      begin

        excel := createOleObject('Excel.Application');
        excel.visible := False;
        excel.DisplayAlerts := False;
        excel.workbooks.open(Data_Module.TemplateDir+'ReportTemplate.xls');
        mysheet := excel.workSheets[1];
        mysheet.cells[1,1].value:='Forecast Detail Report';

        mysheet.cells[3,1].value:='Assembly Part#';
        mysheet.Columns[1].ColumnWidth:=17;

        mysheet.cells[3,2].value:='Effective Month';
        mysheet.Columns[2].ColumnWidth:=17;

        mysheet.cells[3,3].value:='BC Code';
        mysheet.Columns[3].ColumnWidth:=10;

        mysheet.cells[3,4].value:='Tire Part Number';
        mysheet.Columns[4].ColumnWidth:=17;
        mysheet.Columns[4].NumberFormat := '############';

        mysheet.cells[3,5].value:='Tire Share Ratio';
        mysheet.Columns[5].ColumnWidth:=17;

        mysheet.cells[3,6].value:='Wheel Part Number';
        mysheet.Columns[6].ColumnWidth:=17;
        mysheet.Columns[6].NumberFormat := '############';

        mysheet.cells[3,7].value:='Wheel Share Ratio';
        mysheet.Columns[7].ColumnWidth:=17;

        mysheet.cells[3,8].value:='Forecast Share';
        mysheet.Columns[8].ColumnWidth:=17;

        z:=4;
        lastpn:='';
        while not eof do
        begin
          if LastPN <> fieldbyname('VC_ASSY_PART_NUMBER_CODE').AsString then
          begin
            INC(z);
            mysheet.Cells[z,1].value := fieldbyname('VC_ASSY_PART_NUMBER_CODE').AsString;
            mysheet.Cells[z,2].value := fieldbyname('VC_EFFECTIVE_MONTH').AsString;
            lastpn:=fieldbyname('VC_ASSY_PART_NUMBER_CODE').AsString;
            lastmn:=fieldbyname('VC_EFFECTIVE_MONTH').AsString;
          end;
          if lastmn <> fieldbyname('VC_EFFECTIVE_MONTH').AsString then
          begin
            mysheet.Cells[z,2].value := fieldbyname('VC_EFFECTIVE_MONTH').AsString;
            lastmn:=fieldbyname('VC_EFFECTIVE_MONTH').AsString;
          end;

          mysheet.Cells[z,3].value := fieldbyname('VC_bROADCAST_CODE').AsString;
          mysheet.Cells[z,4].value := fieldbyname('VC_TIRE_PART_NUMBER_CODE').AsString;
          mysheet.Cells[z,5].value := fieldbyname('IN_TIRE_RATIO').AsInteger;
          mysheet.Cells[z,6].value := fieldbyname('VC_WHEEL_PART_NUMBER_CODE').AsString;
          mysheet.Cells[z,7].value := fieldbyname('IN_WHEEL_RATIO').AsInteger;
          mysheet.Cells[z,8].value := fieldbyname('IN_RATIO').AsInteger;
          INC(z);
          next;
        end;


        ProcessPanel.Visible:=FALSE;

        excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\ForecastDetail'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
        if MessageDlg('Print this report?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin

          mysheet.PrintOut (EmptyParam, EmptyParam, EmptyParam,
                            EmptyParam, EmptyParam, EmptyParam, EmptyParam);
        end;
        excel.Workbooks.Close;
        excel.Quit;
        excel:=Unassigned;
        mysheet:=Unassigned;
        Data_Module.LogActLog('REPORT','Do Forecast Detail Report');
        ShowMessage('Report Complete');
      end
      else
      begin
        ShowMessage('No Forecast Detail records');
      end;
      ProcessPanel.Visible:=FALSE;
    except
      on e:exception do
      begin
        Data_Module.LogActLog('ERROR','Failed on Forecast Detail Report, '+e.Message);
        ShowMessage('Failed on Forecast Detail Report, '+e.Message);
        excel.Workbooks.Close;
        excel.Quit;
        excel:=Unassigned;
        mysheet:=Unassigned;
        ProcessPanel.Visible:=FALSE;
      end;
    end;
  end;
end;



procedure TMainMenu_Form.CamexTest1Click(Sender: TObject);
var
  Forecastreport:TForecastCAMEXReport;
begin
        ForecastReport := TForecastCAMEXReport.Create();
        if not ForecastReport.Execute then
        begin
          Data_Module.LogActLog('FORECAST','Failed on Camex forecast report');
        end;
end;

end.
