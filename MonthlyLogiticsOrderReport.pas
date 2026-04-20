unit MonthlyLogiticsOrderReport;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, QuickRpt, QRCtrls, ExtCtrls, ADODb;

type
  TMonthlyLogisticsOrderSummary_Form = class(TForm)
    MonthlyLogisticsOrderSummary_QRep: TQuickRep;
    TitleBand1: TQRBand;
    PageFooterBand1: TQRBand;
    DetailBand1: TQRBand;
    ColumnHeaderBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel8: TQRLabel;
    QRLabel9: TQRLabel;
    FromLabel: TQRLabel;
    ToLabel: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel4: TQRLabel;
    QRSysData1: TQRSysData;
    QRSysData2: TQRSysData;
    QRDBText1: TQRDBText;
    QRLabel3: TQRLabel;
    QRDBText2: TQRDBText;
    QRDBText3: TQRDBText;
    QRLabel5: TQRLabel;
    QRLabel6: TQRLabel;
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
  MonthlyLogisticsOrderSummary_Form: TMonthlyLogisticsOrderSummary_Form;

implementation

uses DataModule;

{$R *.dfm}
procedure TMonthlyLogisticsOrderSummary_Form.Execute;
begin
  FromLabel.Caption:=copy(fFromDate,3,2)+'/'+copy(fFromDate,5,2)+'/'+copy(fFromDate,7,2);
  ToLabel.Caption:=copy(fToDate,3,2)+'/'+copy(fToDate,5,2)+'/'+copy(fToDate,7,2);;
  try
    with Data_Module.Inv_DataSet do
    begin
      Close;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.REPORT_MonthlyLogisticsOrders;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@FromDate';
      Parameters.ParamValues['@FromDate'] := fFromDate;
      Parameters.AddParameter.Name := '@ToDate';
      Parameters.ParamValues['@ToDate'] := fToDate;
      Open;
      //MonthlySupplierOrder_QuickRep.Print;
      MonthlyLogisticsOrderSummary_QRep.Preview;
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
