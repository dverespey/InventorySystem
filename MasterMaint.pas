//***********************************************************
//  InventorySystem
//
//  Copyright (c) 2002-2006 Failproof Manufacturing Systems
//
//***********************************************************
//
//  10/25/2002  Aaron Huge      Initial creation
//  04/12/2005  David Verespey  Add Monthly PO
//  09/27/2006  David Verespey  Add BC ratio

unit MasterMaint;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TMasterMaint_Form = class(TForm)
    SupMaster_Button: TButton;
    SizeMaster_Button: TButton;
    AssyRatioMaster_Button: TButton;
    PartsStockMaster_Button: TButton;
    MastDateMaintMenu_Label: TLabel;
    Close_Button: TButton;
    ForecastDetail_Button: TButton;
    Button1: TButton;
    RenbanGroupMaster_Button: TButton;
    MonthlyPO_Button: TButton;
    ASNINVOIVE_Button: TButton;
    procedure SupMaster_ButtonClick(Sender: TObject);
    procedure SizeMaster_ButtonClick(Sender: TObject);
    procedure AssyRatioMaster_ButtonClick(Sender: TObject);
    procedure PartsStockMaster_ButtonClick(Sender: TObject);
    procedure ForecastDetail_ButtonClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure RenbanGroupMaster_ButtonClick(Sender: TObject);
    procedure MonthlyPO_ButtonClick(Sender: TObject);
    procedure ASNINVOIVE_ButtonClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    function Execute: boolean;
  end;

var
  MasterMaint_Form: TMasterMaint_Form;

implementation

uses SupplierMaster, SizeMaster, AssyRatioMaster, PartsStockMaster,
  ForecastDetail, LogisticsMaster, RenbanGroupMaster, MonthlyPOMaster,DataModule,
  ManifestCostMaster, ASNInvoice;

{$R *.dfm}

function TMasterMaint_Form.Execute: boolean;
Begin
  Result:= True;
  try
    try
      //if data_Module.fiAssemblerName.AsString = 'CAMEX' then  //DMV 04/15/2005
      if data_module.fiPOEDISupport.AsBoolean then
        MonthlyPO_Button.Visible:=TRUE
      else
        MonthlyPO_Button.Visible:=FALSE;

      if Data_Module.fiGenerateEDI.AsBoolean then
      begin
        MonthlyPO_Button.Caption:='Manifest Cost';
        AssyRatioMaster_Button.Visible:=FALSE;
        ASNINVOIVE_Button.Visible:=TRUE;
      end;

      AssyRatioMaster_Button.Visible:=FALSE; // not used yet

      ShowModal;
    except
      On E:Exception do
        showMessage('Unable to generate Master Date Maintenance screen.'
          + #13 + 'ERROR:'
          + #13 + E.Message);
    end;
  finally
    If ModalResult = mrCancel Then
      Result:= False;
  end;
end;        //Execute

procedure TMasterMaint_Form.SupMaster_ButtonClick(Sender: TObject);
begin
  try
    Hide;
    SupplierMaster_Form:=TSupplierMaster_Form.Create(self);
    SupplierMaster_Form.Execute;
    SupplierMaster_Form.Free;
  finally
    Show;
  end;

end;

procedure TMasterMaint_Form.SizeMaster_ButtonClick(Sender: TObject);
begin
  try
    Hide;
    SizeMaster_Form:=TSizeMaster_Form.Create(self);
    SizeMaster_Form.Execute;
    SizeMaster_Form.Free;
  finally
    Show;
  end;

end;

procedure TMasterMaint_Form.AssyRatioMaster_ButtonClick(Sender: TObject);
begin
  Hide;
  AssyRatioMaster_Form:=TAssyRatioMaster_Form.Create(self);
  AssyRatioMaster_Form.Execute;
  AssyRatioMaster_Form.Free;
  Show;
end;

procedure TMasterMaint_Form.PartsStockMaster_ButtonClick(Sender: TObject);
begin
  Hide;
  PartsStockMaster_Form:=TPartsStockMaster_Form.Create(self);
  PartsStockMaster_Form.Execute;
  PartsStockMaster_Form.Free;
  Show;
end;

procedure TMasterMaint_Form.ForecastDetail_ButtonClick(Sender: TObject);
begin
  Hide;
  ForecastDetail_Form:=TForecastDetail_Form.Create(self);
  ForecastDetail_Form.Execute;
  ForecastDetail_Form.Free;
  Show;

end;

procedure TMasterMaint_Form.Button1Click(Sender: TObject);
begin
  Hide;
  LogisticsMaster_Form:=TLogisticsMaster_Form.Create(self);
  LogisticsMaster_Form.Execute;
  LogisticsMaster_Form.Free;
  Show;
end;

procedure TMasterMaint_Form.RenbanGroupMaster_ButtonClick(Sender: TObject);
begin
  Hide;
  RenbanGroupMaster_Form:=TRenbanGroupMaster_Form.Create(self);
  RenbanGroupMaster_Form.Execute;
  RenbanGroupMaster_Form.Free;
  Show;
end;

procedure TMasterMaint_Form.MonthlyPO_ButtonClick(Sender: TObject);
begin
  if Data_Module.fiGenerateEDI.AsBoolean then
  begin
    Hide;
    ManifestCostMaster_Form:=TManifestCostMaster_Form.Create(self);
    ManifestCostMaster_Form.Execute;
    ManifestCostMaster_Form.Free;
    Show;
  end
  else
  begin
    Hide;
    MonthlyPOMaster_Form:=TMonthlyPOMaster_Form.Create(self);
    MonthlyPOMaster_Form.Execute;
    MonthlyPOMaster_Form.Free;
    Show;
  end;
end;

procedure TMasterMaint_Form.ASNINVOIVE_ButtonClick(Sender: TObject);
begin
  //Handle ASN and Invoice resends
  try
    Hide;
    ASNInvoice_Form:=TASNInvoice_Form.Create(self);
    ASNInvoice_Form.Execute;
    ASNInvoice_Form.Free
  finally
    Show;
  end;
end;

end.
