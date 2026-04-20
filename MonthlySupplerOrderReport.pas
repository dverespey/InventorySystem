unit MonthlySupplerOrderReport;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, QuickRpt, QRCtrls, ADODb;

type
  TMonthlySupplierOrderSummary_Form = class(TForm)
    MonthlySupplierOrder_QuickRep: TQuickRep;
    ColumnHeaderBand1: TQRBand;
    DetailBand1: TQRBand;
    PageFooterBand1: TQRBand;
    TitleBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRDBText2: TQRDBText;
    QRDBText3: TQRDBText;
    QRDBText4: TQRDBText;
    QRDBText5: TQRDBText;
    QRDBText6: TQRDBText;
    QRSysData1: TQRSysData;
    QRSysData2: TQRSysData;
    QRLabel8: TQRLabel;
    QRLabel9: TQRLabel;
    FromLabel: TQRLabel;
    ToLabel: TQRLabel;
    QRDBText7: TQRDBText;
    QRLabel10: TQRLabel;
    QRLabel11: TQRLabel;
  private
    { Private declarations }
    fFromDate:string;
    fToDAte:string;
  public
    { Public declarations }
    procedure Execute;

    property Fromdate:string
    read fFromDate
    write fFromDate;

    property ToDate:string
    read fToDate
    write fToDate;
  end;

var
  MonthlySupplierOrderSummary_Form: TMonthlySupplierOrderSummary_Form;

implementation

uses DataModule;

{$R *.dfm}

procedure TMonthlySupplierOrderSummary_Form.Execute;
begin
  FromLabel.Caption:=copy(fFromDate,3,2)+'/'+copy(fFromDate,5,2)+'/'+copy(fFromDate,7,2);
  ToLabel.Caption:=copy(fToDate,3,2)+'/'+copy(fToDate,5,2)+'/'+copy(fToDate,7,2);;
  try
    with Data_Module.Inv_DataSet do
    begin
      Close;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.REPORT_MonthlySupplierOrders;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@FromDate';
      Parameters.ParamValues['@FromDate'] := fFromDate;
      Parameters.AddParameter.Name := '@ToDate';
      Parameters.ParamValues['@ToDate'] := fToDate;
      Open;
      //MonthlySupplierOrder_QuickRep.Print;
      MonthlySupplierOrder_QuickRep.Preview;
      Close;
    end;
  except
    on e:exception do
    begin
      ShowMessage('Unable tp get report data, '+e.Message);
      Data_Module.LogActLog('ERROR','Unable to get report data, '+e.Message);
    end;
  end;
end;

end.
