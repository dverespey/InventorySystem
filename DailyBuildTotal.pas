//****************************************************************
//
//       Inventory Control
//
//       Copyright (c) 2002-2005 TAI, Failproof Manufacturing Systems.
//
//****************************************************************
//
// Change History
//
//  09/14/2004  David Verespey  Create Form
//  05/27/2005  David verespey  Break form into parts Read Daily Build/ ASN report/ Invoice report
//
unit DailyBuildTotal;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, History, ComObj, Excel2000, DB, DBTables, ExtCtrls;

type
  TFormMode=(fmDaily,fmASN,fmINVOICE);

  TDailyBuildtotalForm = class(TForm)
    Hist: THistory;
    DoneButton: TButton;
    RunTimer: TTimer;
    procedure DoneButtonClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure RunTimerTimer(Sender: TObject);
  private
    { Private declarations }
    fToDate:TDate;
    fFromDate:TDate;
    fLine:string;
    fFileName:string;
    fFormMode:TFormMode;
  public
    { Public declarations }
    procedure Execute;

    property FileName:string
    read fFileName
    write fFileName;

    property FormMode:TFormMode
    read fFormMode
    write fFormMode;

    property Line:string
    read fLine
    write fLine;

    property FromDate:Tdate
    read fFromDate
    write fFromDate;

    property ToDate:TDate
    read fToDate
    write fToDate;
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
begin
  RunTimer.Enabled:=TRUE;
end;

procedure TDailyBuildtotalForm.RunTimerTimer(Sender: TObject);
var
  z,x,i,count:integer;
  //devdatedone:boolean;
  invoicetotal:extended;
  lastpickup:string;
begin
  RunTimer.Enabled:=FALSE;
  //devdatedone:=FALSE;
  application.ProcessMessages;
  lastpickup:='';

  case fFormMode of
    //
    //  Daily Process of input, update inventory and add information for later ASN/INV processing
    //
    //
    fmDaily:
    begin
      try
        Data_Module.LogActLog('ALC PULL','Start ALC Pull');
        Hist.Append('Start ALC Pull');
        application.ProcessMessages;

        //  Open Worksheet with build info
        //
        excel := createOleObject('Excel.Application');
        excel.visible := False;
        excel.DisplayAlerts := False;
        excel.workbooks.open(fFilename);
        Data_Module.LogActLog('ALC PULL','Filename='+fFilename);
        Hist.Append('Filename='+fFilename);
        mysheet := excel.workSheets[1];

        //  Check if today has already been processed
        //
        Data_Module.ProductionDate:=formatdatetime('yyyymmdd',StrToDate(mysheet.cells[3,2].value));
        Data_Module.AssyCode := fLine;
        Data_Module.Inv_dataset.Filter:='';
        Data_Module.Inv_dataset.Filtered:=FALSE;
        Data_Module.GetShippingInfo;
        if Data_Module.ProductionDate = Data_Module.Inv_DataSet.FieldByName('VC_PRODUCTION_DATE').AsString then
        begin
          //  Already done post error
          //
          Data_Module.LogActLog('ALC PULL','Production Date:'+Data_Module.ProductionDate+', Already processed');
          Hist.Append('Production Date:'+Data_Module.ProductionDate+', Already processed');
        end
        else
        begin
          Data_Module.Inv_Connection.BeginTrans;

          // Find last entry
          i:=8;
          while not VarIsEmpty(mysheet.cells[i,8].value) do
          begin
            INC(i);
          end;

          count:=mysheet.cells[i-1,8].value;

          Data_Module.LogActLog('ALC PULL','Count='+IntToStr(count));
          Hist.Append('Count='+IntToStr(count));
          DAta_Module.Quantity:=Count;
          Data_Module.Continuation:=0;


          //  If not, Move through worksheet add to daily build table
          //
          i:=8;
          while not VarIsEmpty(mysheet.cells[i,1].value) do
          begin
            Data_Module.AssyCode:=mysheet.cells[i,1].value; //Assembly Code
            Data_Module.Quantity:=mysheet.cells[i,8].value; // Build Qty

            if Data_Module.Quantity > 0 then
            begin
              // Remove inventory
              if not Data_Module.InsertExcelShippingInfo then
              begin
                if Data_Module.Inv_Connection.InTransaction then
                  Data_Module.Inv_Connection.RollbackTrans;
                Data_Module.LogActLog('ALC PULL','Inventory update failure for Assembly('+Data_Module.AssyCode+')');
                Hist.Append('Inventory update failure for Assembly('+Data_Module.AssyCode+')');
                ShowMessage('Unable to process ALC pull, fix error and retry');
                excel.Workbooks.Close;
                excel.Quit;
                excel:=Unassigned;
                mysheet:=Unassigned;
                DoneButton.Visible:=TRUE;
                exit;
              end
              else
              begin
                Data_Module.LogActLog('ALC PULL','Import Assy Code = '+Data_Module.AssyCode+', Count = '+IntToStr(Data_Module.Quantity));
                Hist.Append('Import Assy Code = '+Data_Module.AssyCode+', Count = '+IntToStr(Data_Module.Quantity));
                application.ProcessMessages;
              end;

              // Add to Build History
              if not Data_Module.InsertBuildHist then
              begin
                if Data_Module.Inv_Connection.InTransaction then
                  Data_Module.Inv_Connection.RollbackTrans;
                Data_Module.LogActLog('BUILD HIST','Build Hist update failure for Assembly('+Data_Module.AssyCode+')');
                Hist.Append('Build Hist update failure for Assembly('+Data_Module.AssyCode+')');
                ShowMessage('Unable to process ALC pull, fix error and retry');
                excel.Workbooks.Close;
                excel.Quit;
                excel:=Unassigned;
                mysheet:=Unassigned;
                DoneButton.Visible:=TRUE;
                exit;
              end
              else
              begin
                Data_Module.LogActLog('BUILD HIST','BUILD HIST Assy Code = '+Data_Module.AssyCode+', Count = '+IntToStr(Data_Module.Quantity));
                Hist.Append('BUILD HIST Assy Code = '+Data_Module.AssyCode+', Count = '+IntToStr(Data_Module.Quantity));
                application.ProcessMessages;
              end;
            end;
            INC(i);
          end;




          //
          //  Do Scrap
          //
          Hist.Append('Do Scrap............');

          i:=31;
          while not VarIsEmpty(mysheet.cells[i,8].value) do
          begin
            INC(i);
          end;
          x:=i-1;

          for i:=31 to x do
          begin
            if not VarIsEmpty(mysheet.cells[i,1].value) then
            begin
              if mysheet.cells[i,10].value <> 0 then
              begin
                //  process scrap
                //
                //  compare current with new and update if different
                //
                //
                Data_Module.PartNum:=copy(mysheet.cells[i,1].value,1,5)+copy(mysheet.cells[i,1].value,7,5)+'00';
                Data_Module.ScrapCount:=mysheet.cells[i,10].value;

                z:=Data_Module.InsertAutoScrap;
                if z > 0 then
                begin
                  Data_Module.LogActLog('SCRAP','Partnumber = '+Data_Module.PartNum+', Count = '+IntToStr(z));
                  Hist.Append('Scrap Partnumber = '+Data_Module.PartNum+', Count = '+IntToStr(z));
                end
                else if z < 0 then
                begin
                  Data_Module.LogActLog('ERROR','Unable to Scrap Partnumber = '+Data_Module.PartNum+', Count = '+IntToStr(z));
                  Hist.Append('Failed to Scrap Partnumber = '+Data_Module.PartNum+', Count = '+IntToStr(z));
                end;
              end;
            end;
          end;
          Hist.Append('Scrap Complete............');



          Data_Module.Quantity:=count;
          Data_Module.CarTruck:=fLine;
          Data_Module.InsertExcelShippingEndInfo;
          Data_Module.Inv_Connection.CommitTrans;

        end;

        excel.Workbooks.Close;
        excel.Quit;
        excel:=Unassigned;
        mysheet:=Unassigned;

      except
        on e:exception do
        begin
          Data_Module.LogActLog('ERROR','Failed on Daily Build, '+e.Message);
          Hist.Append('Failed on Daily Build, '+e.Message);
          if Data_Module.Inv_Connection.InTransaction then
            Data_Module.Inv_Connection.RollbackTrans;
          excel.Workbooks.Close;
          excel.Quit;
          excel:=Unassigned;
          mysheet:=Unassigned;
        end;
      end;
    end;
    fmASN:
    begin
      try
        Data_Module.LogActLog('ASN CREATE','Start ASN');
        Hist.Append('Start ASN');
        application.ProcessMessages;
        //
        // Get Data based on range
        Data_Module.BeginDate:=fFromDate;
        Data_Module.EndDate:=fToDate;
        Data_Module.LogActLog('ASN CREATE','ASN From='+FormatDateTime('yyyymmdd',fFromDate)+'To='+FormatDateTime('yyyymmdd',fToDate));
        Hist.Append('ASN From='+FormatDateTime('yyyymmdd',fFromDate)+'To='+FormatDateTime('yyyymmdd',fToDate));
        Data_Module.ASN:=1;
        Data_Module.INVOICE:=0;
        Data_Module.GetBuildHist;

        if Data_Module.Inv_DataSet.RecordCount > 0 then
        begin
          Data_Module.Inv_Connection.BeginTrans;
          count:=0;

          // Loop records
          while not Data_Module.Inv_DataSet.Eof do
          begin
            if lastpickup <> Data_module.Inv_DataSet.FieldByName('VC_PICK_UP_DATE').AsString then
            begin
              if lastpickup <> '' then
              begin
                // close old
                //
                //  Save ASN
                //
                excelASN.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\ASN'+formatdatetime('yyyymmddhhmmss0',now)+IntToStr(count)+'.csv',xlCSV);
                INC(count);
                excelASN.Workbooks.Close;
                excelASN.Quit;
                excelASN:=Unassigned;
                mysheetASN:=Unassigned;
                Data_Module.LogActLog('ASN CREATE','Finish ASN PickUpDate='+lastpickup+'PickUptime='+Data_module.Inv_DataSet.FieldByName('VC_PICK_UP_TIME').AsString);
                Hist.Append('Finish ASN PickUpDate='+lastpickup+'PickUptime='+Data_module.Inv_DataSet.FieldByName('VC_PICK_UP_TIME').AsString);
              end;

              lastpickup:=Data_module.Inv_DataSet.FieldByName('VC_PICK_UP_DATE').AsString;

              //
              //  Create ASN
              //
              excelASN := createOleObject('Excel.Application');
              excelASN.visible := FALSE;
              excelASN.DisplayAlerts := False;
              excelASN.workbooks.add;
              mysheetASN := excelASN.workSheets[1];
              mysheetASN.Columns[3].NumberFormat := '@';
              mysheetASN.Columns[2].NumberFormat := '@';

              //ASN Header
              mysheetASN.Cells[1,1].value  := 'H';    //Qualifier
              mysheetASN.Cells[1,2].value  := Data_module.Inv_DataSet.FieldByName('VC_PICK_UP_DATE').AsString; //Pickup Date
              mysheetASN.Cells[1,3].value  := Data_module.Inv_DataSet.FieldByName('VC_PICK_UP_TIME').AsString; //Pickup Time
              mysheetASN.Cells[1,4].value  := 'PT';   //Ship Code
              mysheetASN.Cells[1,5].value  := 'ASM';  //Packaging Code
              mysheetASN.Cells[1,6].value  := '';     //Lading Qty
              mysheetASN.Cells[1,7].value  := '';     //Carrier Supplier Code
              mysheetASN.Cells[1,8].value  := '';     //Equipment Description Code
              mysheetASN.Cells[1,9].value  := '';     //Equipment Number
              x:=2;

              Data_Module.LogActLog('ASN CREATE','ASN PickUpDate='+Data_module.Inv_DataSet.FieldByName('VC_PICK_UP_DATE').AsString+' PickUptime='+Data_module.Inv_DataSet.FieldByName('VC_PICK_UP_TIME').AsString);
              Hist.Append('ASN PickUpDate='+Data_module.Inv_DataSet.FieldByName('VC_PICK_UP_DATE').AsString+' PickUptime='+Data_module.Inv_DataSet.FieldByName('VC_PICK_UP_TIME').AsString);
            end;


            //ASN PO line
            mysheetASN.Cells[x,1].value  := 'O'; //Qualifier
            mysheetASN.Cells[x,2].value  := Data_module.Inv_DataSet.FieldByName('VC_PO_NUMBER').AsString; //PO Number
            INC(X);


            //ASN item line
            mysheetASN.Cells[x,1].value  := 'I'; //Qualifier
            mysheetASN.Cells[x,2].NumberFormat := '@';
            mysheetASN.Cells[x,2].value  := Data_module.Inv_DataSet.FieldByName('VC_ASSY_PART_NUMBER_CODE').AsString; // PartNumber
            mysheetASN.Cells[x,3].NumberFormat := '@';
            mysheetASN.Cells[x,3].value  := Data_module.Inv_DataSet.FieldByName('VC_ASSY_PART_NUMBER_CODE').AsString; // Customer Partnumber

            //mysheetASN.Cells[x,4].value  := Data_module.Inv_DataSet.FieldByName('Qty').AsString; // Units Shipped

            // Use the lesser requested value, PO will not be fully charged against
            {if Data_module.Inv_DataSet.FieldByName('IN_QTY').AsInteger < Data_module.Inv_DataSet.FieldByName('Available').AsInteger then
            begin
              mysheetASN.Cells[x,4].value  := Data_module.Inv_DataSet.FieldByName('IN_QTY').AsString; // Units Shipped
              Data_Module.Quantity := Data_module.Inv_DataSet.FieldByName('IN_QTY').AsInteger;
            end
            // Else use the PO amount. If equal the PO will be filled, if greater than it has already been charge against or this line fills it completely
            else
            begin
              mysheetASN.Cells[x,4].value  := Data_module.Inv_DataSet.FieldByName('Available').AsString; // Units Shipped
              Data_Module.Quantity := Data_module.Inv_DataSet.FieldByName('Available').AsInteger;
            end;}

            mysheetASN.Cells[x,4].value  := Data_module.Inv_DataSet.FieldByName('IN_PO_QTY').AsString; // Units Shipped
            Data_Module.Quantity := Data_module.Inv_DataSet.FieldByName('IN_PO_QTY').AsInteger;


            mysheetASN.Cells[x,5].value  := 'EA'; //Basis Code
            INC(X);

            Data_Module.LogActLog('ASN CREATE','ASN AssyCode='+Data_module.Inv_DataSet.FieldByName('VC_ASSY_PART_NUMBER_CODE').AsString+
                                                    'PONumber='+Data_module.Inv_DataSet.FieldByName('VC_PO_NUMBER').AsString+
                                                    'Qty='+Data_module.Inv_DataSet.FieldByName('IN_PO_QTY').AsString);
            Hist.Append('ASN AssyCode='+Data_module.Inv_DataSet.FieldByName('VC_ASSY_PART_NUMBER_CODE').AsString+
                                                    ' PONumber='+Data_module.Inv_DataSet.FieldByName('VC_PO_NUMBER').AsString+
                                                    ' Qty='+Data_module.Inv_DataSet.FieldByName('IN_PO_QTY').AsString);


            // Add to Charged
            Data_Module.AssyCode:=Data_module.Inv_DataSet.FieldByName('VC_ASSY_PART_NUMBER_CODE').AsString;
            Data_Module.ProductionDate:=Data_module.Inv_DataSet.FieldByName('VC_PRODUCTION_DATE').AsString;
            Data_Module.PickUp:=Data_module.Inv_DataSet.FieldByName('VC_PICK_UP_DATE').AsString;
            //Data_Module.Quantity:=Data_module.Inv_DataSet.FieldByName('Available').AsInteger;
            Data_Module.PONumber:=Data_module.Inv_DataSet.FieldByName('VC_PO_NUMBER').AsString;

            if not Data_Module.InsertPOCharged then
            begin
              if data_Module.Inv_Connection.InTransaction then
                Data_Module.Inv_Connection.RollbackTrans;
              excelASN.Workbooks.Close;
              excelASN.Quit;
              excelASN:=Unassigned;
              mysheetASN:=Unassigned;
              Data_Module.LogActLog('ASN CREATE','Failed ASN');
              Hist.Append('Fail on ASN Create');
              ShowMessage('Fail on ASN Create');
              DoneButton.Visible:=TRUE;
              exit;
            end;

            Data_Module.Inv_DataSet.Next;
          end;

          //
          //  Save ASN
          //
          excelASN.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\ASN'+formatdatetime('yyyymmddhhmmss',now)+IntToStr(count)+'.csv',xlCSV);
          excelASN.Workbooks.Close;
          excelASN.Quit;
          excelASN:=Unassigned;
          mysheetASN:=Unassigned;
          Data_Module.LogActLog('ASN CREATE','Finish ASN PickUpDate='+Data_module.Inv_DataSet.FieldByName('VC_PICK_UP_DATE').AsString+'PickUptime='+Data_module.Inv_DataSet.FieldByName('VC_PICK_UP_TIME').AsString);
          Hist.Append('Finish ASN PickUpDate='+Data_module.Inv_DataSet.FieldByName('VC_PICK_UP_DATE').AsString+' PickUptime='+Data_module.Inv_DataSet.FieldByName('VC_PICK_UP_TIME').AsString);

          Data_Module.Inv_Connection.CommitTrans;

          //
          //  Update ASN Bit
          //Data_Module.UpdateASNDone;

        end
        else
        begin
          Data_Module.LogActLog('ASN','No ASN records for range');
          Hist.Append('No ASN records for range');
        end;
      except
        on e:exception do
        begin
          if data_Module.Inv_Connection.InTransaction then
            Data_Module.Inv_Connection.RollbackTrans;
          Data_Module.LogActLog('ERROR','Failed on ASN, '+e.Message);
          Hist.Append('Failed on ASN, '+e.Message);
          ShowMessage('Fail on ASN Create');
          excelASN.Workbooks.Close;
          excelASN.Quit;
          excelASN:=Unassigned;
          mysheetASN:=Unassigned;
        end;
      end;
    end;
    fmINVOICE:
    begin
      try
        Data_Module.LogActLog('INV CREATE','Start INVOICE');
        Hist.Append('Start INVOICE');
        application.ProcessMessages;
        //
        // Get Data based on range
        Data_Module.BeginDate:=fFromDate;
        Data_Module.EndDate:=fToDate;
        Data_Module.LogActLog('INV CREATE','INVOICE From='+FormatDateTime('yyyymmdd',fFromDate)+' To='+FormatDateTime('yyyymmdd',fToDate));
        Hist.Append('INVOICE From='+FormatDateTime('yyyymmdd',fFromDate)+' To='+FormatDateTime('yyyymmdd',fToDate));
        Data_Module.ASN:=0;
        Data_Module.INVOICE:=1;
        Data_Module.GetBuildHist;
        invoicetotal:=0;
        count:=0;

        if Data_Module.Inv_DataSet.RecordCount > 0 then
        begin
          // Loop records
          Data_Module.Inv_Connection.BeginTrans;
          while not Data_Module.Inv_DataSet.Eof do
          begin
            if lastpickup <> Data_module.Inv_DataSet.FieldByName('VC_PICK_UP_DATE').AsString then
            begin
              if lastpickup <> '' then
              begin
                //
                //  Save Invoice
                //
                mysheetINV.Cells[z,1].value  := 'T';                                    // Qualifier
                mysheetINV.Cells[z,2].value  := floattostrf(invoicetotal,ffFixed,12,4); // Invoice total

                excelINV.ActiveSheet.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\INV'+formatdatetime('yyyymmddhhmmss0',now)+IntToStr(count)+'.csv',xlCSV );
                INC(count);
                excelINV.Workbooks.Close;
                excelINV.Quit;
                excelINV:=Unassigned;
                mysheetINV:=Unassigned;
                Hist.Append('Finish INVOICE PickUpDate='+Data_module.Inv_DataSet.FieldByName('VC_PICK_UP_DATE').AsString+' PickUptime='+Data_module.Inv_DataSet.FieldByName('VC_PICK_UP_TIME').AsString);
                Data_Module.LogActLog('INV CREATE','Finish ASN PickUpDate='+Data_module.Inv_DataSet.FieldByName('VC_PICK_UP_DATE').AsString+' PickUptime='+Data_module.Inv_DataSet.FieldByName('VC_PICK_UP_TIME').AsString);
              end;

              //
              //  Create INVOICE
              //
              excelINV := createOleObject('Excel.Application');
              excelINV.visible := FALSE;
              excelINV.DisplayAlerts := False;
              excelINV.workbooks.add;
              mysheetINV := excelINV.workSheets[1];
              mysheetINV.Columns[6].NumberFormat := '@';

              //INVOICE Header
              mysheetINV.Range['B1','C1'].NumberFormat  := '@';
              mysheetINV.Range['B1','B1'].Value:=copy(Data_module.Inv_DataSet.FieldByName('VC_PICK_UP_DATE').AsString,3,6); //Invoice Date
              mysheetINV.Range['C1','C1'].Value:='01'+copy(Data_module.Inv_DataSet.FieldByName('VC_PICK_UP_DATE').AsString,3,6); //Invoice Number
              mysheetINV.Cells[1,1].value  := 'H'; //Qualifier
              z:=2;

              invoicetotal:=0;
              lastpickup:=Data_module.Inv_DataSet.FieldByName('VC_PICK_UP_DATE').AsString;
            end;

            // Add Line to INVOICE
            mysheetINV.Cells[z,1].value  := 'I';                                                                      // Qualifier
            mysheetINV.Cells[z,2].value  := Data_module.Inv_DataSet.FieldByName('IN_QTY_CHARGED').AsString;           // Units Shipped
            mysheetINV.Cells[z,3].value  := '';                                                                       // Basis Code
            mysheetINV.Cells[z,4].value  := Data_module.Inv_DataSet.FieldByName('MO_ASSEMBLY_COST').AsString;         // Unit Price
            mysheetINV.Cells[z,5].value  := '';                                                                       // Basis of unit price Code
            mysheetINV.Cells[z,6].NumberFormat := '@';
            mysheetINV.Cells[z,6].value  := Data_module.Inv_DataSet.FieldByName('VC_ASSY_PART_NUMBER_CODE').AsString; // Customer Partnumber
            mysheetINV.Cells[z,7].value  := '';                                                                       // Product B
            mysheetINV.Cells[z,8].value  := '';                                                                       // Product C
            mysheetINV.Cells[z,9].value  := Data_module.Inv_DataSet.FieldByName('VC_PO_NUMBER').AsString;             // Reference  USE This to send PO
            mysheetINV.Cells[z,10].value := Data_module.Inv_DataSet.FieldByName('VC_PICK_UP_DATE').AsString;          // Date
            INC(z);

            Data_Module.LogActLog('INV CREATE','INVOICE AssyCode='+Data_module.Inv_DataSet.FieldByName('VC_ASSY_PART_NUMBER_CODE').AsString+
                                                    ' PONumber='+Data_module.Inv_DataSet.FieldByName('VC_PO_NUMBER').AsString+
                                                    ' Qty='+Data_module.Inv_DataSet.FieldByName('IN_QTY_CHARGED').AsString);
            Hist.Append('INVOICE AssyCode='+Data_module.Inv_DataSet.FieldByName('VC_ASSY_PART_NUMBER_CODE').AsString+
                                                    ' PONumber='+Data_module.Inv_DataSet.FieldByName('VC_PO_NUMBER').AsString+
                                                    ' Qty='+Data_module.Inv_DataSet.FieldByName('IN_QTY_CHARGED').AsString);

            invoicetotal:=invoicetotal+Data_module.Inv_DataSet.FieldByName('totalcost').AsFloat;


            //
            //  Update INVOICE Bit
            Data_Module.INVOICE:=Data_module.Inv_DataSet.FieldByName('IN_CHARGE_ID').AsInteger;
            if not Data_Module.UpdateINVDone then
            begin
              if data_Module.Inv_Connection.InTransaction then
                Data_Module.Inv_Connection.RollbackTrans;
              excelINV.Workbooks.Close;
              excelINV.Quit;
              excelINV:=Unassigned;
              mysheetINV:=Unassigned;
              Data_Module.LogActLog('ERROR','Failed on INVOICE');
              Hist.Append('Failed on INVOICE');
              ShowMessage('Fail on INVOICE Create');
              DoneButton.Visible:=TRUE;
              exit;
            end;

            Data_Module.Inv_DataSet.Next;
          end;


          //
          //  Save Invoice
          //
          mysheetINV.Cells[z,1].value  := 'T';                                    // Qualifier
          mysheetINV.Cells[z,2].value  := floattostrf(invoicetotal,ffFixed,12,4); // Invoice total

          excelINV.ActiveSheet.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\INV'+formatdatetime('yyyymmddhhmmss0',now)+IntToStr(count)+'.csv',xlCSV );
          excelINV.Workbooks.Close;
          excelINV.Quit;
          excelINV:=Unassigned;
          mysheetINV:=Unassigned;
          Hist.Append('Finish INVOICE PickUpDate='+Data_module.Inv_DataSet.FieldByName('VC_PICK_UP_DATE').AsString+'PickUptime='+Data_module.Inv_DataSet.FieldByName('VC_PICK_UP_TIME').AsString);
          Data_Module.LogActLog('INV CREATE','Finish ASN PickUpDate='+Data_module.Inv_DataSet.FieldByName('VC_PICK_UP_DATE').AsString+'PickUptime='+Data_module.Inv_DataSet.FieldByName('VC_PICK_UP_TIME').AsString);

          Data_Module.Inv_Connection.CommitTrans;

        end
        else
        begin
          Data_Module.LogActLog('INV','No INVOICE records for range');
          Hist.Append('No INVOICE records for range');
        end;
      except
        on e:exception do
        begin
          if data_Module.Inv_Connection.InTransaction then
            Data_Module.Inv_Connection.RollbackTrans;
          excelINV.Workbooks.Close;
          excelINV.Quit;
          excelINV:=Unassigned;
          mysheetINV:=Unassigned;
          Data_Module.LogActLog('ERROR','Failed on INVOICE, '+e.Message);
          Hist.Append('GetBuildHistFailed on INVOICE, '+e.Message);
          ShowMessage('Fail on INVOICE Create');
        end;
      end;
    end;
  end;

  DoneButton.Visible:=TRUE;
end;

end.
