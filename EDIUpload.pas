unit EDIUpload;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, History, CopyFile, ComObj, Excel2000, StrUtils;

type
  TEDIUpload_Form = class(TForm)
    Hist: THistory;
    OKButton: TButton;
    CopyFile: TCopyFile;
    procedure OKButtonClick(Sender: TObject);
  private
    { Private declarations }
    fClosed:boolean;
  public
    { Public declarations }
    procedure Execute;
  end;

var
  EDIUpload_Form: TEDIUpload_Form;

implementation

uses DataModule, ForecastBreakdownF;

{$R *.dfm}

procedure TEDIUpload_Form.Execute;
var
  sr: TSearchRec;
  fcf:Textfile;
  EIN, EDIType, status, data, header, fcl, EDIFileNumber:string;
  manifest,partnumber, proddate,remitdate:string;
  Qty,x:integer;
  cost,total,remittotal:extended;
  excel,mysheet:variant;
  sl:TStringList;
begin
  try
    EIN:='';
    fClosed:=FALSE;
    Hist.Append('Seraching '+Data_Module.fiEDIIn.AsString+' for EDI documents');

    if FindFirst(Data_Module.fiEDIIn.AsString+'\*.*', faAnyFile, sr) = 0 then
    begin
      repeat
        if ((sr.Attr and faAnyFile) = sr.Attr) and (sr.Attr <> 16) then
        begin
          Hist.Append('Found File: '+sr.Name);
          //Open and check if EDI

          AssignFile(fcf, Data_Module.fiEDIIn.AsString+'\'+sr.Name);
          Reset(fcf);

          Readln(fcf, fcl);
          if pos('ISA',fcl) > 0 then
          begin
            Readln(fcf, fcl);
            Readln(fcf, fcl);
            data:=copy(fcl,4,3);
            if data='830' then
            begin
              EDIFileNumber:=data;
              Hide;
              Application.ProcessMessages;
              ForecastBreakdown_Form:=TForecastBreakdown_Form.Create(self);
              ForecastBreakdown_Form.filename:=Data_Module.fiEDIIn.AsString+'\'+sr.Name;
              ForecastBreakdown_Form.SupplierCode:=Data_Module.fiSupplierCode.AsString;
              ForecastBreakdown_Form.Show;
              ForecastBreakdown_Form.Execute;
              ForecastBreakdown_Form.Free;
              Show;
              Application.ProcessMessages;
              Hist.Append('EDI 830 Imported: '+sr.Name);
              Data_Module.LogActLog('EDIIMP','EDI 830 Imported: '+sr.Name);
            end
            else if data='862' then
            begin
              EDIFileNumber:=data;
              //Firm Order
              //
              // One day forecast
              Readln(fcf,fcl);
              data:=copy(fcl,17,8);
              remitdate:=copy(data,1,4)+'/'+copy(data,5,2)+'/'+copy(data,7,2);
              Readln(fcf,fcl); //ignore line
              Readln(fcf,fcl); //ignore line

              x:=4;

              excel := createOleObject('Excel.Application');
              excel.visible := FALSE;
              excel.DisplayAlerts := False;
              excel.workbooks.open(Data_Module.TemplateDir+'ReportTemplate.xls');
              mysheet := excel.workSheets[1];
              mysheet.cells[1,1].value:='862 Firm Order';

              mysheet.cells[2,1].value:='Order Date:';
              mysheet.cells[2,2].NumberFormat := 'yyyy/mm/dd';
              mysheet.cells[2,2].value:=remitdate;

              mysheet.cells[3,1].value:='Part Number';
              mysheet.Columns[1].ColumnWidth:=17;

              mysheet.cells[3,2].value:='Qty';
              mysheet.Columns[2].ColumnWidth:=17;

              mysheet.cells[3,3].value:='Prod Date';
              mysheet.Columns[3].ColumnWidth:=17;


              Readln(fcf,fcl);
              data:=copy(fcl,1,3);
              while data <> 'CTT' do
              begin
                mysheet.cells[x,1].value:=copy(fcl,9,12);

                while data <> 'SHP' do
                begin
                  Readln(fcf,fcl);
                  data:=copy(fcl,1,3);
                end;

                fcl:=copy(fcl,8,length(fcl));
                data:=copy(fcl,1,pos('*',fcl)-1);

                mysheet.cells[x,2].value:=data;

                fcl:=copy(fcl,pos('*',fcl)+1,length(fcl));
                fcl:=copy(fcl,pos('*',fcl)+1,length(fcl));
                data:=copy(fcl,1,8);
                data:=copy(data,1,4)+'/'+copy(data,5,2)+'/'+copy(data,7,2);

                mysheet.cells[x,3].value:=data;
                mysheet.cells[x,3].NumberFormat := 'yyyy/mm/dd';


                while data <> 'TD5' do
                begin
                  Readln(fcf,fcl);
                  data:=copy(fcl,1,3);
                end;
                Readln(fcf,fcl);
                data:=copy(fcl,1,3);
                INC(x)
              end;



              excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\FirmOrder'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
              excel.Workbooks.Close;
              excel.Quit;
              excel:=Unassigned;
              mysheet:=Unassigned;
              Hist.Append('EDI 862 Processed: '+sr.Name);
              Data_Module.LogActLog('EDIIMP','EDI 862 Processed: '+sr.Name);
            end
            else if data='997' then
            begin
              EDIFileNumber:=data;
              //Ack
              Readln(fcf, fcl);
              data:=copy(fcl,1,3);
              while data = 'AK1' do
              begin
                EDIType:=copy(fcl,5,2);
                EIN:=copy(fcl,8,9);
                Readln(fcf, fcl);
                Status:=copy(fcl,5,1);

                Data_Module.EIN:=StrToInt(EIN);
                Data_Module.EINStatus:=Status;
                Data_Module.EINType:=EDIType;

                if Data_Module.UpdateEINStatus then
                begin
                  if Status = 'A' then
                  begin
                    if EDIType = 'SH' then
                    begin
                      Hist.Append('EDI EIN('+EIN+') 856 Accepted');
                      Data_Module.LogActLog('EDIIMP','EDI EIN('+EIN+') 856 Accepted');
                    end
                    else
                    begin
                      Hist.Append('EDI EIN('+EIN+') 810 Accepted');
                      Data_Module.LogActLog('EDIIMP','EDI EIN('+EIN+') 810 Accepted');
                    end;
                  end
                  else
                  begin
                    if EDIType = 'SH' then
                    begin
                      Hist.Append('EDI EIN('+EIN+') 856 Rejected');
                      Data_Module.LogActLog('EDIIMP','EDI EIN('+EIN+') 856 Rejected');
                    end
                    else
                    begin
                      Hist.Append('EDI EIN('+EIN+') 810 Rejected');
                      Data_Module.LogActLog('EDIIMP','EDI EIN('+EIN+') 810 Rejected');
                    end
                  end;
                end
                else
                begin
                  if EDIType = 'SH' then
                  begin
                    Hist.Append('Unable to update 856 EIN('+EIN+')');
                    Data_Module.LogActLog('ERROR','Unable to update 856 EIN('+EIN+')');
                  end
                  else
                  begin
                    Hist.Append('Unable to update 810 EIN('+EIN+')');
                    Data_Module.LogActLog('ERROR','Unable to update 810 EIN('+EIN+')');
                  end;
                end;

                Readln(fcf, fcl);
                data:=copy(fcl,1,3);
              end;

              Hist.Append('EDI 997 Processed: '+sr.Name);
              Data_Module.LogActLog('EDIIMP','EDI 997 Imported: '+sr.Name);
            end
            else if data='824' then
            begin

              //CREATE REPORT SHEET AND POST TO USER
              excel := createOleObject('Excel.Application');
              excel.visible := FALSE;
              excel.DisplayAlerts := False;
              excel.workbooks.open(Data_Module.TemplateDir+'ReportTemplate.xls');
              mysheet := excel.workSheets[1];
              mysheet.cells[1,1].value:='Remittance';

              mysheet.cells[3,1].value:='Manifest Number';
              mysheet.Columns[1].ColumnWidth:=17;

              mysheet.cells[3,2].value:='Part Number';
              mysheet.Columns[2].ColumnWidth:=17;

              mysheet.cells[3,3].value:='Error Text';
              mysheet.Columns[3].ColumnWidth:=60;


              EDIFileNumber:=data;
              // Error Do report and flag
              Readln(fcf, fcl);
              data:=copy(fcl,1,3);
              x:=4;
              while data <> 'SE*' do
              begin
                if data = 'NTE' then
                begin
                  mysheet.cells[x,3].value:=copy(fcl,9,50);
                  mysheet.cells[x,1].value:=copy(fcl,60,8);
                  mysheet.cells[x,2].value:=copy(fcl,69,12);
                  INC(x);
                end;

                Readln(fcf, fcl);
                data:=copy(fcl,1,3);
              end;


              excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\ReceivingAdvice'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
              excel.Workbooks.Close;
              excel.Quit;
              excel:=Unassigned;
              mysheet:=Unassigned;

              Hist.Append('EDI 824 Processed: '+sr.Name);
              Hist.Append('ASN ERRORS: '+IntToStr(x-3)+ 'total, view the file ('+Data_Module.fiReportsOutputDir.AsString+'\ReceivingAdvice'+formatdatetime('yyyymmddhhmmss00',now)+'.xls'+') for details');
            end
            else if data='820' then
            begin
              EDIFileNumber:=data;
              //Remittance
              //
              // Store in Table and print report
              //
              // Currently only a report is available
              //
              header:='';
              while header <> 'BPR' do
              begin
                Readln(fcf,fcl); //ignore line
                header:=copy(fcl,1,pos('*',fcl)-1);
              end;

              sl:=TStringList.Create;
              sl.Delimiter:='*';
              sl.DelimitedText:=fcl;

              remittotal:=StrToFloat(sl[2]);
              remitdate:=sl[16];

              {Readln(fcf,fcl); //ignore line
              fcl:=copy(fcl,7,length(fcl));
              header:=copy(fcl,1,pos('*',fcl)-1);
              remittotal:=StrToFloat(data);
              fcl:=copy(fcl,pos('*',fcl)+21,length(fcl));
              remitdate:=copy(fcl,1,8);
              }


              excel := createOleObject('Excel.Application');
              excel.visible := FALSE;
              excel.DisplayAlerts := False;
              excel.workbooks.open(Data_Module.TemplateDir+'ReportTemplate.xls');
              mysheet := excel.workSheets[1];
              mysheet.cells[1,1].value:='Remittance';


              mysheet.cells[2,1].value:='Remittance Date:';
              mysheet.cells[2,2].value:=remitdate;
              mysheet.cells[2,5].value:='Remittance Value:';
              mysheet.cells[2,6].value:=remittotal;


              mysheet.cells[3,1].value:='Manifest Number';
              mysheet.Columns[1].ColumnWidth:=17;

              mysheet.cells[3,2].value:='Part Number';
              mysheet.Columns[2].ColumnWidth:=17;

              mysheet.cells[3,3].value:='Prod Date';
              mysheet.Columns[3].ColumnWidth:=17;

              mysheet.cells[3,4].value:='Qty';
              mysheet.Columns[4].ColumnWidth:=17;

              mysheet.cells[3,5].value:='Unit Cost';
              mysheet.Columns[5].ColumnWidth:=30;
              mysheet.Columns[5].NumberFormat := '$#,##0.0000_);[Red]($#,##0.0000)';

              mysheet.cells[3,6].value:='Item Total';
              mysheet.Columns[6].ColumnWidth:=30;
              mysheet.Columns[6].NumberFormat := '$#,##0.0000_);[Red]($#,##0.0000)';


              total := 0;
              qty := 0;
              cost :=0;
              x:=4;
              while header <> 'SE' do
              begin

                Readln(fcf,fcl);
                header := LeftStr(fcl,Pos('*', fcl)-1);

                if header =  'RMR' then
                begin
                  manifest:=copy(fcl,8,8); //Manifest #
                  data:=copy(fcl,18,length(fcl));
                  total:=StrToFloat(data); //total Paid
                end
                else if header = 'IT1' then
                begin
                  fcl:=copy(fcl,10,length(fcl)); //ignore begining
                  data:=copy(fcl,1,pos('*',fcl)-1); //Qty
                  Qty:=StrToInt(data);
                  fcl:=copy(fcl,pos('*',fcl)+4,length(fcl)); //skip to cost
                  data:=copy(fcl,1,pos('*',fcl)-1); //Cost
                  cost:=StrToFloat(data);
                  fcl:=copy(fcl,pos('*',fcl)+7,length(fcl)); //skip to partnumber
                  partnumber:=copy(fcl,1,12);
                end
                else if header = 'DTM' then
                begin
                  proddate:=copy(fcl,9,8);

                  mysheet.Cells[x,1].value := manifest;
                  mysheet.Cells[x,2].value := partnumber;
                  mysheet.Cells[x,3].value := proddate;
                  mysheet.Cells[x,4].value := IntToStr(Qty);
                  mysheet.Cells[x,5].value := cost;
                  mysheet.Cells[x,6].value := total;
                  inc(x);

                  Hist.Append('Manifest('+manifest+'), PartNumber('+partnumber+'), ProdDate('+proddate+'), Qty('+IntToStr(Qty)+'), Cost('+FloatToStr(cost)+'), Total('+FloatToStr(total)+')');
                end;
{
                Readln(fcf,fcl);
                Readln(fcf,fcl);
                Readln(fcf,fcl);
                Readln(fcf,fcl);
                manifest:=copy(fcl,8,8); //Manifest #
                data:=copy(fcl,18,length(fcl));
                total:=StrToFloat(data); //total Paid
                Readln(fcf,fcl);
                // Get line info
                fcl:=copy(fcl,10,length(fcl)); //ignore begining
                data:=copy(fcl,1,pos('*',fcl)-1); //Qty
                Qty:=StrToInt(data);
                fcl:=copy(fcl,pos('*',fcl)+4,length(fcl)); //skip to cost
                data:=copy(fcl,1,pos('*',fcl)-1); //Cost
                cost:=StrToFloat(data);
                fcl:=copy(fcl,pos('*',fcl)+7,length(fcl)); //skip to partnumber
                partnumber:=copy(fcl,1,12);


                Readln(fcf,fcl);
                Readln(fcf,fcl);
                proddate:=copy(fcl,9,8);
                Readln(fcf,fcl);

                mysheet.Cells[x,1].value := manifest;
                mysheet.Cells[x,2].value := partnumber;
                mysheet.Cells[x,3].value := proddate;
                mysheet.Cells[x,4].value := IntToStr(Qty);
                mysheet.Cells[x,5].value := cost;
                mysheet.Cells[x,6].value := total;
                inc(x);

                Hist.Append('Manifest('+manifest+'), PartNumber('+partnumber+'), ProdDate('+proddate+'), Qty('+IntToStr(Qty)+'), Cost('+FloatToStr(cost)+'), Total('+FloatToStr(total)+')');
                data:=copy(fcl,1,3);}

              end;


              excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\Remittance'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
              excel.Workbooks.Close;
              excel.Quit;
              excel:=Unassigned;
              mysheet:=Unassigned;
              Hist.Append('EDI 820 Processed: '+sr.Name);

            end
            else
            begin
              Hist.Append('ERROR: File('+sr.Name+') is an unexpected EDI('+data+') document, it cannot be processed');
            end;
            // Move File
            CloseFile(fcf);
            CopyFile.Movefile:=TRUE;
            CopyFile.CopyFrom:=Data_Module.fiEDIIn.AsString+'\'+sr.Name;
            // Change move to file name to give easier detail by filename
            //CopyFile.CopyTo:=Data_Module.fiEDIIn.AsString+'\Archive\'+sr.Name;
            if EIN <> '' then
              CopyFile.CopyTo:=Data_Module.fiEDIIn.AsString+'\Archive\'+EDIFileNumber+'-'+EIN+'.EDI'
            else
              CopyFile.CopyTo:=Data_Module.fiEDIIn.AsString+'\Archive\'+EDIFileNumber+'.EDI';
            CopyFile.Copynow;

          end
          else
          begin
            Hist.Append('ERROR: File('+sr.Name+') is not a valid EDI document');
            // Move file
            CloseFile(fcf);
            CopyFile.Movefile:=TRUE;
            CopyFile.CopyFrom:=Data_Module.fiEDIIn.AsString+'\'+sr.Name;
            CopyFile.CopyTo:=Data_Module.fiEDIIn.AsString+'\Archive\NOTEDI'+sr.Name;
            CopyFile.Copynow;
          end;
        end;
      until FindNext(sr) <> 0;
      FindClose(sr);
    end;
    // Find Files
    // Check if EDI
    // If so process and then copy to Archive
    // Else just move and log
  finally
    Hist.Append('Completed EDI Upload');
    OKButton.Visible:=TRUE;
    while not fclosed do
    begin
      application.ProcessMessages;
      sleep(500);
    end;
  end;
end;


procedure TEDIUpload_Form.OKButtonClick(Sender: TObject);
begin
  fClosed:=TRUE;
end;

end.
