//****************************************************************
//
//       Inventory Control
//
//       Copyright (c) 2002-2004 Failproof Manufacturing Systems.
//
//****************************************************************
//
// Change History
//
//  09/14/2004  David Verespey  Create Form
//
unit DailyBuildTotal;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, History, ComObj, Excel2000, DB, DBTables;

type
  TDailyBuildtotalForm = class(TForm)
    Hist: THistory;
    DoneButton: TButton;
    procedure DoneButtonClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
    fFileName:string;
  public
    { Public declarations }
    procedure Execute;

    property FileName:string
    read fFileName
    write fFileName;
  end;

var
  DailyBuildtotalForm: TDailyBuildtotalForm;
  excel,mysheet:variant;
  excelASN,mysheetASN:variant;
  excelINV,mysheetINV:variant;

implementation

uses DataModule;

{$R *.dfm}
procedure TDailyBuildtotalForm.Execute;
begin
  ShowModal;
end;

procedure TDailyBuildtotalForm.DoneButtonClick(Sender: TObject);
begin
  Close;
end;

procedure TDailyBuildtotalForm.FormShow(Sender: TObject);
var
  z,x,i,count:integer;
  invoicetotal:extended;
begin
  try
    invoicetotal:=0;
    Data_Module.LogActLog('ALC PULL','Start ALC Pull');
    Hist.Append('Start ALC Pull');

    application.ProcessMessages;

    excel := createOleObject('Excel.Application');
    excel.visible := False;
    excel.DisplayAlerts := False;
    excel.workbooks.open(fFilename);
    Data_Module.LogActLog('ALC PULL','Filename='+fFilename);
    Hist.Append('Filename='+fFilename);
    mysheet := excel.workSheets[1];

    if (not VarIsEmpty(mysheet.cells[4,4].value)) and (not VarIsEmpty(mysheet.cells[5,4].value)) and (not VarIsEmpty(mysheet.cells[5,6].value)) then
    begin
      //
      //  Get build infomation and process
      //
      Data_Module.ProductionDate:=formatdatetime('yyyymmdd',StrToDate(mysheet.cells[4,4].value));
      Data_Module.AssyCode := 'T';
      Data_Module.Inv_dataset.Filter:='';
      Data_Module.Inv_dataset.Filtered:=FALSE;
      Data_Module.GetShippingInfo;
      if Data_Module.ProductionDate = Data_Module.Inv_DataSet.FieldByName('VC_PRODUCTION_DATE').AsString then
      begin
        Data_Module.LogActLog('ALC PULL','Production Date:'+Data_Module.ProductionDate+', Already processed');
        Hist.Append('Production Date:'+Data_Module.ProductionDate+', Already processed');
      end
      else
      begin
        Data_Module.StartSeq:=mysheet.cells[5,4].value;
        Data_Module.LastSeq:=mysheet.cells[5,6].value;

        Data_Module.LogActLog('ALC PULL','Production Date:'+Data_Module.ProductionDate+', Start Seq:'+Data_Module.StartSeq+', End Seq:'+Data_Module.LastSeq);
        Hist.Append('Production Date:'+Data_Module.ProductionDate+', Start Seq:'+Data_Module.StartSeq+', End Seq:'+Data_Module.LastSeq);

        if StrToInt(Data_Module.StartSeq) > StrToInt(Data_Module.LastSeq) then
        begin
          count:=StrToInt(Data_Module.LastSeq)+(StrToInt(Data_Module.StartSeq)-Data_Module.fiHighSequence.AsInteger);
          Data_Module.LogActLog('ALC PULL','Count='+IntToStr(count));
          Hist.Append('Count='+IntToStr(count));
        end
        else
        begin
          count:=(StrToInt(Data_Module.LastSeq)-StrToInt(Data_Module.StartSeq))+1;
          Data_Module.LogActLog('ALC PULL','Count='+IntToStr(count));
          Hist.Append('Count='+IntToStr(count));
        end;
        DAta_Module.Quantity:=Count;
        Data_Module.Continuation:=0;

        //
        // Get assembly part numbers and pass to process
        //
        try


          //
          //  Create ASN
          //
          excelASN := createOleObject('Excel.Application');
          excelASN.visible := FALSE;
          excelASN.DisplayAlerts := False;
          excelASN.workbooks.add;
          mysheetASN := excelASN.workSheets[1];

          //ASN Header
          mysheetASN.Cells[1,1].value  := 'H'; //Qualifier
          mysheetASN.Cells[1,2].value  := formatdatetime('yyyymmdd',StrToDate(mysheet.cells[4,4].value)); //ShipDate
          mysheetASN.Cells[1,3].value  := formatdatetime('hhmm',now); // Ship Time, use now unless something else is needed
          mysheetASN.Cells[1,4].value  := 'PT'; //Ship Code
          mysheetASN.Cells[1,5].value  := 'ASM'; //Packaging Code
          mysheetASN.Cells[1,6].value  := ''; //Lading Qty
          mysheetASN.Cells[1,7].value  := ''; //Carrier Supplier Code
          mysheetASN.Cells[1,8].value  := ''; //Equipment Description Code
          mysheetASN.Cells[1,9].value  := ''; //Equipment Number
          x:=2;


          //
          //  Create INVOICE
          //
          excelINV := createOleObject('Excel.Application');
          excelINV.visible := FALSE;
          excelINV.DisplayAlerts := False;
          excelINV.workbooks.add;
          mysheetINV := excelINV.workSheets[1];

          //INVOICE Header
          mysheetINV.Cells[1,1].value  := 'H'; //Qualifier
          mysheetINV.Cells[1,2].value  := formatdatetime('yyyymmdd',StrToDate(mysheet.cells[4,4].value)); //Invoice Number
          mysheetINV.Cells[1,3].value  := formatdatetime('01ymmdd',StrToDate(mysheet.cells[4,4].value)); //Invoice Number
          z:=2;


          Data_Module.Inv_Connection.BeginTrans;
          i:=10;

          while not VarIsEmpty(mysheet.cells[i,7].value) do
          begin
            Data_Module.AssyCode:='42600'+mysheet.cells[i,7].value+'00';
            Data_Module.Quantity:=mysheet.cells[i,11].value;

            if Data_Module.Quantity > 0 then
            begin
              // Remove inventory
              if not Data_Module.InsertExcelShippingInfo then
              begin
                Data_Module.Inv_Connection.RollbackTrans;
                Data_Module.LogActLog('ALC PULL','Import Failed');
                Hist.Append('Import Failed');
                break;
              end
              else
              begin
                Data_Module.LogActLog('ALC PULL','Import Assy Code = '+Data_Module.AssyCode+', Count = '+IntToStr(Data_Module.Quantity));
                Hist.Append('Import Assy Code = '+Data_Module.AssyCode+', Count = '+IntToStr(Data_Module.Quantity));
                application.ProcessMessages;
              end;

              //ASN PO line
              mysheetASN.Cells[x,1].value  := 'O'; //Qualifier
              mysheetASN.Cells[x,2].value  := Data_module.Inv_StoredProc.FieldByName('VC_BLANKET_PO').AsString; //PO Number
              INC(X);

              //ASN item line
              mysheetASN.Cells[x,1].value  := 'I'; //Qualifier
              mysheetASN.Cells[x,2].value  := Data_Module.AssyCode; // PartNumber
              mysheetASN.Cells[x,3].value  := Data_Module.AssyCode; // Customer Partnumber
              mysheetASN.Cells[x,4].value  := Data_Module.Quantity; //Units Shipped
              mysheetASN.Cells[x,5].value  := 'EA'; //Basis Code
              INC(X);

              // Add Line to INVOICE
              mysheetINV.Cells[z,1].value  := 'I'; //Qualifier
              mysheetINV.Cells[z,2].value  := Data_Module.Quantity; //Units Shipped
              mysheetINV.Cells[z,3].value  := ''; // Basis Code
              mysheetINV.Cells[z,4].value  := Data_module.Inv_StoredProc.FieldByName('MO_ASSEMBLY_COST').AsString; // Unit Price
              mysheetINV.Cells[z,5].value  := ''; // Basis of unit price Code
              mysheetINV.Cells[z,6].value  := Data_Module.AssyCode; // Customer Partnumber
              mysheetINV.Cells[z,7].value  := ''; // Product B
              mysheetINV.Cells[z,8].value  := ''; // Product C
              mysheetINV.Cells[z,9].value  := Data_module.Inv_StoredProc.FieldByName('VC_BLANKET_PO').AsString; // Reference  USE This to send PO
              mysheetINV.Cells[z,10].value  := formatdatetime('yyyymmdd',StrToDate(mysheet.cells[4,4].value)); // Date
              INC(z);

              invoicetotal:=invoicetotal+Data_Module.Quantity*Data_module.Inv_StoredProc.FieldByName('MO_ASSEMBLY_COST').AsFloat;
            end;

            INC(i);
          end;
          Data_Module.CarTruck:='T';
          Data_Module.InsertExcelShippingEndInfo;
          Data_Module.Inv_Connection.CommitTrans;
        except
          on e:exception do
          begin
            Data_Module.Inv_Connection.RollbackTrans;
            Hist.Append('Import Failed');
          end;
        end;
        Hist.Append('Pull Complete............');

        //
        //  Do Scrap
        //
        mysheet := excel.workSheets[2];

        for i:=6 to 15 do
        begin
          if not VarIsEmpty(mysheet.cells[i,5].value) then
          begin
            if mysheet.cells[i,5].value <> 0 then
            begin
              //process scrap
              //mysheet.cells[i,4].value+
              Data_Module.LogActLog('SCRAP','Partnumber = '+mysheet.cells[i,3].value+', Count = '+mysheet.cells[i,5].value);
              Hist.Append('Scrap Partnumber = '+mysheet.cells[i,3].value+', Count = '+mysheet.cells[i,5].value);
            end;
          end;
        end;
        Hist.Append('Scrap Complete............');




        //
        //  Save ASN
        //
        excelASN.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\ASN'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
        excelASN.Workbooks.Close;
        excelASN.Quit;
        excelASN:=Unassigned;
        mysheetASN:=Unassigned;
        Hist.Append('ASN Complete............');


        //
        //  Save Invoice
        //
        mysheetINV.Cells[z,1].value  := 'T';      // Qualifier
        mysheetINV.Cells[z,2].value  := floattostrf(invoicetotal,ffFixed,12,4); // Invoice total

        excelINV.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\INV'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
        excelINV.Workbooks.Close;
        excelINV.Quit;
        excelINV:=Unassigned;
        mysheetINV:=Unassigned;
        Hist.Append('INVOICE Complete............');


        Hist.Append('File Processed............');
        DoneButton.Visible:=TRUE;
      end;
    end
    else
    begin
      Data_Module.LogActLog('ERROR','Missing data on build total form');
      Hist.Append('Missing data on build total form');
    end;


    excel.Workbooks.Close;
    excel.Quit;
    excel:=Unassigned;
    mysheet:=Unassigned;

    Data_Module.LogActLog('ALC PULL','End ALC Pull');
    Hist.Append('End ALC Pull');

  except
    on e:exception do
    begin
      Data_Module.LogActLog('ERROR','Failed on Daily Build, '+e.Message);
      Hist.Append('Failed on Daily Build, '+e.Message);
      excel.Workbooks.Close;
      excel.Quit;
      excel:=Unassigned;
      mysheet:=Unassigned;
    end;
  end;
end;

end.
