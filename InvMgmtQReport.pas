//***********************************************************
//  InventorySystem
//
//  Copyright (c) 2002 Failproof Manufacturing Systems
//
//***********************************************************
//
//  10/25/2002  Aaron Huge  Initial creation

unit InvMgmtQReport;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, DataModule, QuickRpt, QRCtrls;

type
  TInvMgmtQReport_Form = class(TForm)
    InvMgmt_QuickRep: TQuickRep;
    PageHeaderBand1: TQRBand;
    DateTime_QRSysData: TQRSysData;
    TitleBand1: TQRBand;
    Title_QRSysData: TQRSysData;
    ColumnHeaderBand1: TQRBand;
    SupCode_QRLabel: TQRLabel;
    PartCode_QRLabel: TQRLabel;
    SizeCode_QRLabel: TQRLabel;
    KANBAN_QRLabel: TQRLabel;
    TireWheel_QRLabel: TQRLabel;
    CarTruck_QRLabel: TQRLabel;
    QTY_QRLabel: TQRLabel;
    DetailBand1: TQRBand;
    SupCode_QRDBText: TQRDBText;
    PartCode_QRDBText: TQRDBText;
    SizeCode_QRDBText: TQRDBText;
    Kanban_QRDBText: TQRDBText;
    TireWheel_QRDBText: TQRDBText;
    CarTruck_QRDBText: TQRDBText;
    QTY_QRDBText: TQRDBText;
    PageFooterBand1: TQRBand;
    PageNum_QRSysData: TQRSysData;
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    function Execute: boolean;
  end;

var
  InvMgmtQReport_Form: TInvMgmtQReport_Form;

implementation

{$R *.dfm}
function TInvMgmtQReport_Form.Execute: boolean;
Begin
  Result:= True;
  try
    try
//      data_module.ADODataSet1.Open;
      ShowModal;
    except
      On E:Exception do
        showMessage('Unable to generate Inventory Management report.'
          + #13 + 'ERROR:'
          + #13 + E.Message);
    end;
  finally
{
    data_module.ADODataSet1.Close;
    data_module.ADOConnection1.Close;
}
    If ModalResult = mrCancel Then
      Result:= False;
  end;
end;        //Execute


procedure TInvMgmtQReport_Form.FormCreate(Sender: TObject);
begin
  try
    With InvMgmt_QuickRep, Data_Module.Grid_ClientDataSet do
    Begin
      {SupCode_QRDBText.DataField := 'Supplier';   //ShowMessage(Fields[0].Name);
      PartCode_QRDBText.DataField := 'Parts';     //Fields[1].AsString;
      SizeCode_QRDBText.DataField := 'Size';      //Fields[2].AsString;
      Kanban_QRDBText.DataField := 'KANBAN';      //Fields[3].AsString;
      TireWheel_QRDBText.DataField := 'Tire / Wheel';   //Fields[4].AsString;
      CarTruck_QRDBText.DataField := 'Car / Truck';     //Fields[5].AsString;
      QTY_QRDBText.DataField := 'QTY';            //Fields[6].AsString;
      Price_QRDBText.DataField := 'Report Price';        //Fields[7].AsString;
      Price_QRDBText.Alignment := taRightJustify;
      TotalCost_QRDBText.DataField := 'Report Total';     //Fields[8].AsString;
      TotalCost_QRDBText.Alignment := taRightJustify;}
    End;    //With
  except
    On e:exception do
    Begin
      ShowMessage('Unable to generate report at this time.');
      Data_Module.LogActLog('ERROR', 'InvMgmtReport: ' + e.Message + ' Class:' + e.ClassName, 0);
    End;
  end;      //try...except
end;

end.
