//****************************************************************
//
//       Inventory Control
//
//       Copyright (c) 2002-2003 Failproof Manufacturing Systems.
//
//****************************************************************
//
// Change History
//
//  03/10/2003  David Verespey  Create Form
//
unit UploadBreakDown;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, DataModule, ComCtrls, StdCtrls,
  DirOutln, OleServer, Excel2000, Excel97;

type
  TBreakdownKind = (bForecast,bReceiving,bInvoice,bBuildout,bDailyBuildT,bDailyBuildP);

  TUpBreakDown_Form = class(TForm)
    ForeUpBreakDown_Label: TLabel;
    FileName_Label: TLabel;
    FileName_Edit: TEdit;
    Browse_Button: TButton;
    ForecastFilleNameDialog: TOpenDialog;
    Start_Button: TButton;
    Close_Button: TButton;
    procedure Browse_ButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Start_ButtonClick(Sender: TObject);
    procedure FileName_EditChange(Sender: TObject);
  private
    { Private declarations }
    fFileSelected:boolean;
    fBreakdownKind:TBreakdownKind;
  public
    { Public declarations }
    function Execute: boolean;

    property BreakdownKind : TBreakdownKind
    read fBreakdownKind
    write fBreakdownKind;
  end;

var
  UpBreakDown_Form: TUpBreakDown_Form;

implementation

uses ForecastBreakdownF, LogisticsBreakdown, InvoiceBreakdown,
  ManualForecast, DailyBuildTotal;

{$R *.dfm}

function TUpBreakDown_Form.Execute: boolean;
Begin
  Result:= True;
  try
    try
      fFileSelected:=False;
      if fBreakdownKind = bForecast then
      begin
        ForeUpBreakDown_Label.Caption:='Forecast '+ForeUpBreakDown_Label.Caption;
        UpBreakDown_Form.Caption:='Forecast '+UpBreakDown_Form.Caption;
      end
      else if fBreakdownKind = bInvoice then
      begin
        ForeUpBreakDown_Label.Caption:='Invoice '+ForeUpBreakDown_Label.Caption;
        UpBreakDown_Form.Caption:='Invoice '+UpBreakDown_Form.Caption;
      end
      else if fBreakdownKind = bReceiving then
      begin
        ForeUpBreakDown_Label.Caption:='Receiving '+ForeUpBreakDown_Label.Caption;
        UpBreakDown_Form.Caption:='Receiving '+UpBreakDown_Form.Caption;
      end
      else if fBreakdownKind = bBuildout then
      begin
        ForeUpBreakDown_Label.Caption:='Build out '+ForeUpBreakDown_Label.Caption;
        UpBreakDown_Form.Caption:='Build out '+UpBreakDown_Form.Caption;
      end
      else if fBreakdownKind = bDailyBuildT then
      begin
        ForeUpBreakDown_Label.Caption:='Daily Build Worksheet(Truck) '+ForeUpBreakDown_Label.Caption;
        UpBreakDown_Form.Caption:='Daily Build Worksheet(Truck) '+UpBreakDown_Form.Caption;
      end;
      ShowModal;
    except
      On E:Exception do
      begin
        showMessage('Unable to generate, Upload & Break Down screen.'
          + #13 + 'ERROR:'
          + #13 + E.Message);
      end;
    end;
  finally
    If ModalResult = mrCancel Then
      Result:= False;
  end;

end;      //Execute

procedure TUpBreakDown_Form.Browse_ButtonClick(Sender: TObject);
begin
  with ForecastFilleNameDialog do
  begin
    if fBreakdownKind = bForecast then
    begin
      if Data_Module.fiAssemblerName.AsString = 'WQS' then
      begin
        Filter := 'Text Forecast File (*.prelftp)|*.prelftp';
        InitialDir:=Data_Module.fiForecastInputDir.AsString;
        DefaultExt := 'prelftp';
        Filename := Data_Module.fiForecastFilename.AsString;
      end
      else if Data_Module.fiAssemblerName.AsString = 'CAMEX' then
      begin
        Filter := 'Text Forecast File (*.txt)|*.txt';
        InitialDir:=Data_Module.fiForecastInputDir.AsString;
        DefaultExt := 'txt';
        Filename := '';
      end
    end
    else if fBreakdownKind = bInvoice then
    begin
      Filter := 'Text Invoice File (*.inv)|*.inv|Text files (*.txt)|*.TXT';
      InitialDir:='D:\_Inventory_Control\Suppliers';
      DefaultExt := 'inv';
      Filename := '';
    end
    else if fBreakdownKind = bReceiving then
    begin
      Filter := 'Text Logistics file (*.txt)|*.TXT|Text Logistics File (*.log)|*.log';
      InitialDir:=Data_Module.fiLogisticsInputDir.AsString;
      DefaultExt := 'txt';
      Filename := Data_Module.fiLogisticsFilename.AsString;
    end
    else if fBreakdownKind = bBuildout then
    begin
      Filter := 'Excel Buildout file (*.xls)|*.XLS|Excel Buildout File (*.xls)|*.xls';
      InitialDir:=Data_Module.fiForecastInputDir.AsString;
      DefaultExt := 'xls';
      Filename := '';
    end
    else if (fBreakdownKind = bDailyBuildT) or (fBreakdownKind = bDailyBuildP) then
    begin
      Filter := 'Excel Daily Build Worksheet file (*.xls)|*.XLS|Excel Daily Build Worksheet File (*.xls)|*.xls';
      InitialDir:=Data_Module.fiTextShippingFileDir.AsString;
      DefaultExt := 'xls';
      Filename := '';
    end;

    Options := [ofHideReadOnly, ofFileMustExist,
      ofPathMustExist];
    if Execute then
    begin
      FileName_Edit.Text := ForecastFilleNameDialog.FileName;
      fFileselected:=TRUE;
      Start_Button.Enabled:=TRUE;
    end;
  end;
end;

procedure TUpBreakDown_Form.FormCreate(Sender: TObject);
begin
  FileName_Edit.Text := '[Type file path and name here or click Browse]';
  Start_Button.Enabled:=FALSE;
  fFileSelected:=FALSE;
end;

procedure TUpBreakDown_Form.Start_ButtonClick(Sender: TObject);
var
  breakdown:TForecastBreakdown_Form;
begin
  try
    if fFileSelected then
    begin
      if fBreakdownKind = bForecast then
      begin
        breakdown:=TForecastBreakdown_Form.Create(self);
        breakdown.filename:=ForecastFilleNameDialog.FileName;
        breakdown.SupplierCode:=data_Module.fiSupplierCode.AsString;
        breakdown.Show;
        if not breakdown.Execute then
          Showmessage('Forecast has not been added, please retry');
        breakdown.Free;
        close;
      end
      else if fBreakdownKind = bInvoice then
      begin
        InvoiceBreakdown_Form:=TInvoiceBreakdown_Form.Create(self);
        InvoiceBreakdown_Form.filename:=ForecastFilleNameDialog.FileName;
        InvoiceBreakdown_Form.Execute;
        InvoiceBreakdown_Form.Free;
      end
      else if fBreakdownKind = bReceiving then
      begin
        LogisticsBreakdown_Form:=TLogisticsBreakdown_Form.Create(self);
        LogisticsBreakdown_Form.filename:=ForecastFilleNameDialog.FileName;
        LogisticsBreakdown_Form.Execute;
        LogisticsBreakdown_Form.Free;
      end
      else if fBreakdownKind = bBuildout then
      begin
        ManualForecast_Form:=TManualForecast_Form.Create(self);
        ManualForecast_Form.filename:=ForecastFilleNameDialog.FileName;
        //ManualForecast_Form.Show;
        ManualForecast_Form.Execute;
        ManualForecast_Form.Free;
      end
      else if fBreakdownKind = bDailyBuildT then
      begin
        DailyBuildtotalForm:=TDailyBuildtotalForm.Create(self);
        DailyBuildtotalForm.filename:=ForecastFilleNameDialog.FileName;
        DailyBuildtotalForm.Line:='T';
        DailyBuildTotalForm.FormMode:=fmDaily;
        DailyBuildtotalForm.Execute;
        DailyBuildtotalForm.Free;
      end
      else if fBreakdownKind = bDailyBuildP then
      begin
        DailyBuildtotalForm:=TDailyBuildtotalForm.Create(self);
        DailyBuildtotalForm.filename:=ForecastFilleNameDialog.FileName;
        DailyBuildtotalForm.Line:='P';
        DailyBuildTotalForm.FormMode:=fmDaily;
        DailyBuildtotalForm.Execute;
        DailyBuildtotalForm.Free;
      end;
    end
    else
      ShowMessage('Please select a valid file first');

    FileName_Edit.Text:='[Type file path and name here or click Browse]';
    Start_Button.Enabled:=FALSE;
  except
    on e:exception do
      ShowMessage('Err: ' + e.Message);
  end;
end;

procedure TUpBreakDown_Form.FileName_EditChange(Sender: TObject);
begin
  fFileSelected:=True;
end;

end.
