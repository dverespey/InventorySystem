//****************************************************************
//
//       Inventory Control
//
//       Copyright (c) 2002-2008 Failproof Manufacturing Systems.
//
//****************************************************************
//

unit Configuration;

interface

uses Windows, SysUtils, Classes, Graphics, Forms, Controls, StdCtrls,
  Buttons, ExtCtrls, Filectrl, ComCtrls;

type
  TConfigurationDlg = class(TForm)
    OKBtn: TButton;
    Button1: TButton;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    Label8: TLabel;
    ForecastUsageCompare_Edit: TEdit;
    UsageUpdateCompare_Edit: TEdit;
    ForecastDir_SpeedButton: TSpeedButton;
    Label3: TLabel;
    Label4: TLabel;
    LogisticsDir_SpeedButton: TSpeedButton;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    ReportsDir_SpeedButton: TSpeedButton;
    ForecastFileName_Edit: TEdit;
    ForecastFileDir_Edit: TEdit;
    LogisticsFileName_Edit: TEdit;
    LogisticsFileDir_Edit: TEdit;
    ReportsOutputDir_Edit: TEdit;
    Label2: TLabel;
    SupplierCode_Edit: TEdit;
    LocalFTP_CheckBox: TCheckBox;
    Label9: TLabel;
    PlantName_Edit: TEdit;
    Label10: TLabel;
    AssemblerName_Edit: TEdit;
    EDIPOSUpport_CheckBox: TCheckBox;
    BuildOutButton_CheckBox: TCheckBox;
    Label1: TLabel;
    EDIInputDir_Edit: TEdit;
    EDIInputDir_SpeedButton: TSpeedButton;
    Label11: TLabel;
    EDIOutputDir_Edit: TEdit;
    EDIOutputDir_SpeedButton: TSpeedButton;
    Label12: TLabel;
    ExcelOrderSheet_CheckBox: TCheckBox;
    TabSheet4: TTabSheet;
    EnableDataPurgeCheckBox: TCheckBox;
    DataRetention_Edit: TEdit;
    PromptDataPurgeCheckBox: TCheckBox;
    Label13: TLabel;
    PurgeRateComboBox: TComboBox;
    Label14: TLabel;
    PurgeDayWeekly_Group: TRadioGroup;
    PurgeDayMonthly_Group: TRadioGroup;
    OrderFillDays_Edit: TEdit;
    Label15: TLabel;
    ConfirmOrderFileCreationCheckBox: TCheckBox;
    Label16: TLabel;
    TemplateDir_Edit: TEdit;
    TemplateDir_SpeedButton: TSpeedButton;
    UseApplicationDir_CheckBox: TCheckBox;
    CreatePOPriorToCloseCheckBox: TCheckBox;
    procedure ForecastDir_SpeedButtonClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure ServerName_EditChange(Sender: TObject);
    procedure LocalFTP_CheckBoxClick(Sender: TObject);
    procedure PurgeRateComboBoxChange(Sender: TObject);
    procedure UseApplicationDir_CheckBoxClick(Sender: TObject);
  private

    { Private declarations }
    fCancel:boolean;
    fChanged:boolean;

    procedure ShowPurgeRateDetail;
  public
    { Public declarations }
    procedure Execute;

    property Cancel:boolean
    read fCancel
    write fCancel;

    property DataChanged:boolean
    read fChanged;
  end;

var
  ConfigurationDlg: TConfigurationDlg;

implementation

uses DataModule{, DirectorySelect};

{$R *.dfm}

procedure TConfigurationDlg.Execute;
var
  i:integer;
begin
  SupplierCode_Edit.Text:=Data_Module.fiSupplierCode.AsString;

  ForecastFilename_Edit.Text:=Data_Module.fiForecastFilename.AsString;
  ForecastFileDir_Edit.Text:=Data_Module.fiForecastInputDir.AsString;

  LogisticsFilename_edit.Text:=Data_Module.fiLogisticsFilename.AsString;
  LogisticsFileDir_Edit.Text:=Data_Module.fiLogisticsInputDir.AsString;

  EDIInputDir_Edit.Text:=Data_Module.fiEDIIn.AsString;
  EDIOutputDir_Edit.Text:=Data_Module.fiEDIOut.AsString;

  ReportsOutputDir_Edit.Text:=Data_Module.fiReportsOutputDir.AsString;

  LocalFTP_CheckBox.Checked:=Data_Module.fiLocalFTP.AsBoolean;

  ForecastUsageCompare_Edit.Text:=Data_Module.fiForecastUsageCompare.AsString;

  UsageUpdateCompare_Edit.Text:=Data_Module.fiUsageUpdateCompare.AsString;


  TemplateDir_Edit.Text:=Data_Module.fiTemplateDir.AsString;
  UseApplicationDir_CheckBox.Checked:=Data_Module.fiUseApplicationDir.AsBoolean;

  PlantName_Edit.Text:=Data_Module.fiPlantName.AsString;
  AssemblerName_Edit.Text:=Data_Module.fiAssemblerName.AsString;

  LocalFTP_CheckBox.Checked:=Data_Module.fiLocalFTP.AsBoolean;
  EDIPOSUpport_CheckBox.Checked:=DATA_Module.fiPOEDISupport.AsBoolean;
  BuildOutButton_CheckBox.Checked:=Data_Module.fiBuildOut.AsBoolean;
  ExcelOrderSheet_CheckBox.Checked:=Data_Module.fiExcelOrderSheet.AsBoolean;

  EnableDataPurgeCheckBox.Checked:=Data_Module.fiEnableDataPurge.AsBoolean;
  PromptDataPurgeCheckBox.Checked:=Data_Module.fiPromptDataPurge.AsBoolean;
  DataRetention_Edit.Text:=Data_Module.fiDataRetention.AsString;
  Data_Module.SearchCombo(PurgeRateComboBox,Data_Module.fiPurgeRate.AsString);

  OrderFillDays_Edit.Text:=Data_Module.fiFillDays.AsString;

  ConfirmOrderFileCreationCheckBox.Checked := Data_Module.fiConfirmOrderFileCreation.AsBoolean;
  CreatePOPriorToCloseCheckBox.Checked := Data_Module.fiCreatePOPriorToClose.AsBoolean;

  ShowPurgeRateDetail;


  Cancel:=False;
  fChanged:=FALSE;

  ShowModal;

  if (not cancel) and fChanged then
  begin
    Data_Module.fiSupplierCode.AsString:=SupplierCode_Edit.Text;

    Data_Module.fiForecastFilename.AsString:=ForecastFilename_Edit.Text;
    Data_Module.fiForecastInputDir.AsString:=ForecastFileDir_Edit.Text;

    Data_Module.fiLogisticsFilename.AsString:=LogisticsFilename_edit.Text;
    Data_Module.fiLogisticsInputDir.AsString:=LogisticsFileDir_Edit.Text;

    Data_Module.fiReportsOutputDir.AsString:=ReportsOutputDir_Edit.Text;
    Data_Module.fiLocalFTP.AsBoolean:=LocalFTP_CheckBox.Checked;

    Data_Module.fiEDIIn.AsString:=EDIInputDir_Edit.Text;
    Data_Module.fiEDIOut.AsString:=EDIOutputDir_Edit.Text;

    Data_Module.fiTemplateDir.AsString:=TemplateDir_Edit.Text;
    Data_Module.fiUseApplicationDir.AsBoolean:=UseApplicationDir_CheckBox.Checked;


    if TryStrToInt(ForecastUsageCompare_Edit.Text,i) then
      Data_Module.fiForecastUsageCompare.AsString:=ForecastUsageCompare_Edit.Text;
    if TryStrToInt(UsageUpdateCompare_Edit.Text,i) then
      Data_Module.fiUsageUpdateCompare.AsString:=UsageUpdateCompare_Edit.Text;

    Data_Module.fiPlantName.AsString:=PlantName_Edit.Text;
    Data_Module.fiAssemblerName.AsString:=AssemblerName_Edit.Text;

    Data_Module.fiLocalFTP.AsBoolean:=LocalFTP_CheckBox.Checked;
    DATA_Module.fiPOEDISupport.AsBoolean:=EDIPOSUpport_CheckBox.Checked;
    Data_Module.fiBuildOut.AsBoolean:=BuildOutButton_CheckBox.Checked;
    Data_Module.fiExcelOrderSheet.AsBoolean:=ExcelOrderSheet_CheckBox.Checked;

    Data_Module.fiEnableDataPurge.AsBoolean:=EnableDataPurgeCheckBox.Checked;
    Data_Module.fiPromptDataPurge.AsBoolean:=PromptDataPurgeCheckBox.Checked;
    Data_Module.fiDataRetention.AsString:=DataRetention_Edit.Text;
    Data_Module.fiPurgeRate.AsString:=PurgeRateComboBox.Text;

    if PurgeDayWeekly_Group.ItemIndex >= 0 then
      Data_Module.fiPurgeDayWeekly.AsString:=PurgeDayWeekly_Group.Items[PurgeDayWeekly_Group.ItemIndex];
    if PurgeDayMonthly_Group.ItemIndex >= 0 then
      Data_Module.fiPurgeDayMonthly.AsString:=PurgeDayMonthly_Group.Items[PurgeDayMonthly_Group.ItemIndex];

    if TryStrToInt(OrderFillDays_Edit.Text, i) then
      if i<=50 then
        Data_Module.fiFillDays.AsString:=OrderFillDays_Edit.Text;

    Data_Module.fiConfirmOrderFileCreation.AsBoolean :=ConfirmOrderFileCreationCheckBox.Checked;
    Data_Module.fiCreatePOPriorToClose.AsBoolean:=CreatePOPriorToCloseCheckBox.Checked;
  end;
end;

procedure TConfigurationDlg.ForecastDir_SpeedButtonClick(Sender: TObject);
var
  dir:string;
begin
  if TComponent(Sender).Name='ForecastDir_SpeedButton' then
  begin
    dir:=ForecastFileDir_Edit.Text;
    if SelectDirectory('Select A Directory','My Computer',dir) then
      ForecastFileDir_Edit.Text:=Dir;
  end
  else if TComponent(Sender).Name='LogisticsDir_SpeedButton' then
  begin
    dir:=LogisticsFileDir_Edit.Text;
    if SelectDirectory('Select A Directory','My Computer',dir) then
      LogisticsFileDir_Edit.Text:=Dir;
  end
  else if TComponent(Sender).Name='EDIInputDir_SpeedButton' then
  begin
    dir:=EDIInputDir_Edit.Text;
    if SelectDirectory('Select A Directory','My Computer',dir) then
      EDIInputDir_Edit.Text:=Dir;
  end
  else if TComponent(Sender).Name='EDIOutputDir_SpeedButton' then
  begin
    dir:=EDIOutputDir_Edit.Text;
    if SelectDirectory('Select A Directory','My Computer',dir) then
      EDIOutputDir_Edit.Text:=Dir;
  end
  else if TComponent(Sender).Name='ReportsDir_SpeedButton' then
  begin
    dir:=ReportsOutputDir_Edit.Text;
    if SelectDirectory('Select A Directory','My Computer',dir) then
      ReportsOutputDir_Edit.Text:=Dir;
  end
  else if TComponent(Sender).Name='TemplateDir_SpeedButton' then
  begin
    dir:=TemplateDir_Edit.Text;
    if SelectDirectory('Select A Directory','My Computer',dir) then
      TemplateDir_Edit.Text:=Dir;
  end;
end;

procedure TConfigurationDlg.Button1Click(Sender: TObject);
begin
  FCancel:=True;
end;

procedure TConfigurationDlg.ServerName_EditChange(Sender: TObject);
begin
  fChanged:=TRUE;
end;

procedure TConfigurationDlg.LocalFTP_CheckBoxClick(Sender: TObject);
begin
  fChanged:=TRUE;
end;

procedure TConfigurationDlg.PurgeRateComboBoxChange(Sender: TObject);
begin
  fChanged:=TRUE;

  ShowPurgeRateDetail;
end;

procedure TConfigurationDlg.ShowPurgeRateDetail;
begin
  PurgeDayWeekly_Group.Visible:=FALSE;
  PurgeDayMonthly_Group.Visible:=FALSE;

  case PurgeRateComboBox.ItemIndex of
    0:
    begin
    end;
    1:
    begin
      PurgeDayWeekly_Group.Visible:=TRUE;

       if UpperCase(Data_Module.fiPurgeDayWeekly.AsString)='MONDAY' then
       begin
        PurgeDayWeekly_Group.ItemIndex:=0;
       end
       else if UpperCase(Data_Module.fiPurgeDayWeekly.AsString)='TUESDAY' then
       begin
        PurgeDayWeekly_Group.ItemIndex:=1;
       end
       else if UpperCase(Data_Module.fiPurgeDayWeekly.AsString)='WEDNESDAY' then
       begin
        PurgeDayWeekly_Group.ItemIndex:=2;
       end
       else if UpperCase(Data_Module.fiPurgeDayWeekly.AsString)='THURSDAY' then
       begin
        PurgeDayWeekly_Group.ItemIndex:=3;
       end
       else if UpperCase(Data_Module.fiPurgeDayWeekly.AsString)='FRIDAY' then
       begin
        PurgeDayWeekly_Group.ItemIndex:=4;
       end;
    end;
    2:
    begin
      PurgeDayMonthly_Group.Visible:=TRUE;

       if UpperCase(Data_Module.fiPurgeDayMonthly.AsString)='1ST' then
       begin
        PurgeDayMonthly_Group.ItemIndex:=0;
       end
       else if UpperCase(Data_Module.fiPurgeDayMonthly.AsString)='15TH' then
       begin
        PurgeDayMonthly_Group.ItemIndex:=1;
       end
       else if UpperCase(Data_Module.fiPurgeDayMonthly.AsString)='LAST' then
       begin
        PurgeDayMonthly_Group.ItemIndex:=2;
       end;
    end;
  end;

end;

procedure TConfigurationDlg.UseApplicationDir_CheckBoxClick(
  Sender: TObject);
begin
  TemplateDir_Edit.ReadOnly:=UseApplicationDir_CheckBox.Checked;
  TemplateDir_SpeedButton.Enabled:= not UseApplicationDir_CheckBox.Checked;
  fChanged:=TRUE;
end;

end.
