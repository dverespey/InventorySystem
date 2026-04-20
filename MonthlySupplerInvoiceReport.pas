unit MonthlySupplerInvoiceReport;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, QRCtrls, QuickRpt, ExtCtrls, ADODB, StdCtrls;

type
  TMonthlySupplierInvoice_Form = class(TForm)
    MonthlySupplierInvoices_QuickRep: TQuickRep;
    ColumnHeaderBand1: TQRBand;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    DetailBand1: TQRBand;
    QRDBText2: TQRDBText;
    QRDBText3: TQRDBText;
    QRDBText4: TQRDBText;
    QRDBText5: TQRDBText;
    QRDBText6: TQRDBText;
    QRDBText7: TQRDBText;
    PageFooterBand1: TQRBand;
    QRSysData1: TQRSysData;
    QRSysData2: TQRSysData;
    TitleBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel8: TQRLabel;
    QRLabel9: TQRLabel;
    FromLabel: TQRLabel;
    ToLabel: TQRLabel;
    QRLabel10: TQRLabel;
    SupplierLabel: TQRLabel;
    QRLabel11: TQRLabel;
    QRDBText1: TQRDBText;
    QRLabel12: TQRLabel;
    QRLabel13: TQRLabel;
    QRLabel14: TQRLabel;
    QRLabel15: TQRLabel;
    QRDBText8: TQRDBText;
    QRDBText9: TQRDBText;
  private
    { Private declarations }
    fFromDate:string;
    fToDAte:string;
    fSupplier:string;
  public
    { Public declarations }
    procedure Execute;

    property Fromdate:string
    read fFromDate
    write fFromDate;

    property ToDate:string
    read fToDate
    write fToDate;

    property Supplier:string
    read fSupplier
    write fSupplier;
  end;

var
  MonthlySupplierInvoice_Form: TMonthlySupplierInvoice_Form;

implementation

uses DataModule;

{$R *.dfm}

procedure TMonthlySupplierInvoice_Form.Execute;
begin
  FromLabel.Caption:=copy(fFromDate,3,2)+'/'+copy(fFromDate,5,2)+'/'+copy(fFromDate,7,2);
  ToLabel.Caption:=copy(fToDate,3,2)+'/'+copy(fToDate,5,2)+'/'+copy(fToDate,7,2);
  SupplierLabel.Caption:=fSupplier;

  try
    with Data_Module.Inv_DataSet do
    begin
      Close;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.REPORT_MonthlySupplierInvoices;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@FromDate';
      Parameters.ParamValues['@FromDate'] := fFromDate;
      Parameters.AddParameter.Name := '@ToDate';
      Parameters.ParamValues['@ToDate'] := fToDate;
      Parameters.AddParameter.Name := '@Supplier';
      Parameters.ParamValues['@Supplier'] := fSupplier;

      Open;
      //MonthlySupplierInvoices_QuickRep.Print;
      MonthlySupplierInvoices_QuickRep.Preview;
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
