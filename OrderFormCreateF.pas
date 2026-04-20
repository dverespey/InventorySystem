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
//  06/24/2005  David Verespey  Add Partnumber logistics
//
unit OrderFormCreateF;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, History,DataModule, ADODB, StrUtils, ComObj,DateUtils;

type
  TOrderFormCreate_Form = class(TForm)
    OK_Button: TButton;
    Hist: THistory;
    procedure OK_ButtonClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
  private
    { Private declarations }
    fFileKind:TFileKind;
    function GetShip(lead:integer):integer;
  public
    { Public declarations }
    function Execute:boolean;
  published
    property FileKind:TFileKind
    read fFileKind
    write fFileKind;
  end;

var
  OrderFormCreate_Form: TOrderFormCreate_Form;

implementation


{$R *.dfm}

function TOrderFormCreate_Form.Execute:boolean;
begin
  result:=true;
  ShowModal;
end;

procedure TOrderFormCreate_Form.FormActivate(Sender: TObject);
var
  tcf,taf,tlf:textfile;
  tcl,fn,year:string;
  fopen,lasttimestamp,sendsite:boolean;
  lastrenban, lastsupplier,lastsuppliername,lastdirectory,lastlogisticsdirectory:string;
  excel:        variant;
  mySheet:      variant;
  excelOrder:        variant;
  mySheetOrder:      variant;
  page,i,o,ship:integer;
begin
  // Init both file types;
  // Init both file types;
  excel:=Unassigned;
  fopen:=False;
  lastsupplier:='';
  lastRenban:='';
  lasttimestamp:=TRUE;
  sendsite:=FALSE;
  page:=1;

  //  //  Produce output files to send to suppliers
  //
  try
    Hist.Append('Start order file creation');
    Data_Module.LogActLog('ORDERF','Start order file creation');
    Data_Module.Inv_Connection.BeginTrans;
    with Data_Module.Inv_DataSet do
    begin
      //  Get non-ordered orders
      Close;
      Filter:='';
      Filtered:=FALSE;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.SELECT_OrderNotOrdered;1';
      Parameters.Clear;
      Open;

      if recordcount > 0 then
      begin
        //
        //  New Selection Code
        //
        while not eof do
        begin
          if fieldbyname('VC_SUPPLIER_CODE').AsString <> lastsupplier then
          begin
            // Special Order Form Close if open
            if not VarIsEmpty(excelOrder) then
            begin
              //  Create file in directory specified for each supplier
              excelOrder.ActiveWorkbook.SaveAs(lastdirectory+'\OS'+ANSIReplaceStr(lastsuppliername,'\','')+'-'+lastsupplier+'-'+lastRenban);
              if Data_Module.fiLocalFTP.AsBoolean then
              begin
                //logistics
                if lastlogisticsdirectory <> 'NONE' then
                  excelOrder.ActiveWorkbook.SaveAs(lastlogisticsdirectory+'\OS'+ANSIReplaceStr(lastsuppliername,'\','')+'-'+lastsupplier+'-'+lastRenban);
                //Archive
                excelOrder.ActiveWorkbook.SaveAs(lastdirectory+'\'+'Archive'+'\OS'+ANSIReplaceStr(lastsuppliername,'\','')+'-'+lastsupplier+'-'+lastRenban);
              end;

              excelOrder.Workbooks.Close;
              excelOrder.Quit;
              excelOrder:=Unassigned;
              page:=1;
            end;


            // last is Excel then close
            if not VarIsEmpty(excel) then
            begin
              //  Create file in directory specified for each supplier
              //  Use supplier name+WeekDate for filename
              if lasttimestamp then
              begin
                excel.ActiveWorkbook.SaveAs(lastdirectory+'\'+ANSIReplaceStr(lastsuppliername,'\','')+'-'+lastsupplier+'-'+formatdatetime('yyyymmddhhmmss00',now));
                if Data_Module.fiLocalFTP.AsBoolean then
                begin
                  //logistics
                  if lastlogisticsdirectory <> 'NONE' then
                    excel.ActiveWorkbook.SaveAs(lastlogisticsdirectory+'\'+ANSIReplaceStr(lastsuppliername,'\','')+'-'+lastsupplier+'-'+formatdatetime('yyyymmddhhmmss00',now));
                  //Archive
                  excel.ActiveWorkbook.SaveAs(lastdirectory+'\'+'Archive'+'\'+ANSIReplaceStr(lastsuppliername,'\','')+'-'+lastsupplier+'-'+formatdatetime('yyyymmddhhmmss00',now));
                end;
              end
              else
              begin
                excel.ActiveWorkbook.SaveAs(lastdirectory+'\'+ANSIReplaceStr(lastsuppliername,'\','')+'-'+lastsupplier);
                if Data_Module.fiLocalFTP.AsBoolean then
                begin
                  //logistics
                  if lastlogisticsdirectory <> 'NONE' then
                    excel.ActiveWorkbook.SaveAs(lastlogisticsdirectory+'\'+ANSIReplaceStr(lastsuppliername,'\','')+'-'+lastsupplier);
                  //Archive
                  excel.ActiveWorkbook.SaveAs(lastdirectory+'\'+'Archive'+'\'+ANSIReplaceStr(lastsuppliername,'\','')+'-'+lastsupplier+'-'+formatdatetime('yyyymmddhhmmss00',now));
                end;
              end;

              excel.Workbooks.Close;
              excel.Quit;
              excel:=Unassigned;
            end;
            // or text
            if fopen then
            begin
              CloseFile(tcf);
              if Data_Module.fiLocalFTP.AsBoolean then
              begin
                if lastlogisticsdirectory <> 'NONE' then
                  CloseFile(tlf);
                CloseFile(taf);
              end;
              fopen:=FALSE;
            end;

            // new supplier code
            lastsupplier:=fieldbyname('VC_SUPPLIER_CODE').AsString;


            //
            //  Get PartNumber Logistics if available
            //              DMV 06/24/2005
            //
            With Data_Module.Inv_StoredProc do
            Begin
              Close;
              ProcedureName := 'dbo.SELECT_PartsStockLogistics;1';
              Parameters.Clear;
              Parameters.AddParameter.Name := '@PartCode';
              Parameters.ParamValues['@PartCode'] := Data_Module.Inv_DataSet.fieldbyname('VC_PART_NUMBER').AsString;
              Open;

              if recordcount > 0 then
                lastlogisticsdirectory:=fieldbyname('LogisticsDirectory').AsString
              else
                lastlogisticsdirectory:='';
            End;    //With

            //
            // Get FileKind from SUPPLIER table
            //
            With Data_Module.Inv_StoredProc do
            Begin
              Close;
              ProcedureName := 'dbo.SELECT_SupplierInfo;1';
              Parameters.Clear;
              Parameters.AddParameter.Name := '@SupCode';
              Parameters.ParamValues['@SupCode'] := lastsupplier;
              Parameters.AddParameter.Name := '@Logistics';
              Parameters.ParamValues['@Logistics'] := 1;
              Open;

              if FieldByName('Output File Type').AsString = 'TEXT' then
                fFileKind := fText
              else if FieldByName('Output File Type').AsString = 'EXCEL' then
                fFileKind := fExcel
              else
                fFileKind := fBoth;


              lastdirectory:=fieldbyname('Directory').AsString;
              if lastlogisticsdirectory = '' then
              begin
                if not fieldbyname('LogisticsDirectory').IsNull then
                  lastlogisticsdirectory:=fieldbyname('LogisticsDirectory').AsString
                else
                  lastlogisticsdirectory:='NONE';

              end;

              //filter out bad characters
              lastsuppliername:=AnsiReplaceText(AnsiReplaceText(fieldbyname('Supplier Name').AsString,',',''),'.','');
              lasttimestamp:=fieldbyname('Order File Timestamp').AsBoolean;
              sendsite:=fieldbyname('Site Number in Order').AsBoolean;

            End;    //With

            //
            //  Special Order Sheet
            //
            if (Data_Module.Inv_StoredProc.fieldbyname('Create Order Sheet').AsString <> '') and
            (Data_Module.Inv_StoredProc.fieldbyname('Create Order Sheet').AsString <> ' ') then
            begin
              //Special Order Sheet Format
              excelOrder := createOleObject('Excel.Application');
              excelOrder.visible := FALSE; //DEBUG
              excelOrder.DisplayAlerts := False;

              if Data_Module.Inv_StoredProc.fieldbyname('Create Order Sheet').AsString = 'WHEEL' then
                excelOrder.workbooks.open(Data_Module.TemplateDir+'OrderSheetTemplateWheel.xls')
              else
                excelOrder.workbooks.open(Data_Module.TemplateDir+'OrderSheetTemplateTire.xls');

              mysheetOrder := excelOrder.workSheets[1];
              Hist.Append('Create Order Sheet('+Data_Module.Inv_StoredProc.fieldbyname('Create Order Sheet').AsString+') file for supplier, '+lastsupplier);
              Data_Module.LogActLog('ORDERS','Create Order Sheet('+Data_Module.Inv_StoredProc.fieldbyname('Create Order Sheet').AsString+') file for supplier, '+lastsupplier);

              mysheetOrder.Cells[11,1].value  := Data_Module.Inv_StoredProc.fieldbyname('Supplier Name').AsString;
              mysheetOrder.Cells[11,2].value  := Data_Module.Inv_StoredProc.fieldbyname('Supplier Code').AsString;
              if Data_Module.Inv_StoredProc.fieldbyname('Create Order Sheet').AsString = 'WHEEL' then
                mysheetOrder.Cells[11,3].value  := fieldbyname('VC_RENBAN_NUMBER').AsString;
              year:=copy(formatdatetime('yyyy',now),1,3)+copy(fieldbyname('VC_FRS_NUMBER').AsString,1,1);
              mysheetOrder.Cells[8,8].value  := copy(fieldbyname('VC_FRS_NUMBER').AsString,2,2)+'/'+copy(fieldbyname('VC_FRS_NUMBER').AsString,4,2)+'/'+year;
              mysheetOrder.Cells[11,5].value  := page;
              mysheetOrder.Cells[13,2].value  := Data_Module.Inv_StoredProc.fieldbyname('Address').AsString;
              mysheetOrder.Cells[14,2].value  := Data_Module.Inv_StoredProc.fieldbyname('City').AsString+', '+Data_Module.Inv_StoredProc.fieldbyname('State').AsString+', '+Data_Module.Inv_StoredProc.fieldbyname('zip').AsString;

              lastRenban:=fieldbyname('VC_RENBAN_NUMBER').AsString;
              page:=1;
              o:=16;
            end;

            if (fFileKind = fExcel) or (fFileKind = fBoth) then
            begin
              excel := createOleObject('Excel.Application');
              excel.visible := False;
              excel.DisplayAlerts := False;

              excel.workbooks.open(Data_Module.TemplateDir+'OrderTemplate.xls');
              mysheet := excel.workSheets[1];
              Hist.Append('Create excel file for supplier, '+lastsupplier);
              Data_Module.LogActLog('ORDERF','Create excel file for supplier, '+lastsupplier);

              mysheet.Cells[1,6].value  := Data_Module.Inv_StoredProc.fieldbyname('Supplier Name').AsString;
              mysheet.Cells[2,6].value  := Data_Module.Inv_StoredProc.fieldbyname('Address').AsString;

              i:=10;
            end;
            if (fFileKind = fText) or (fFileKind = fBoth) then
            begin
              if lasttimestamp then
              begin
                // Add Timestamp to filename
                //
                fn:=lastdirectory+'\'+ANSIReplaceStr(lastsuppliername,'\','')+'-'+lastsupplier+formatdatetime('yyyymmddhhmmss00',now)+'.ord';
                AssignFile(tcf,fn);
                Rewrite(tcf);
                Hist.Append('Create text file for supplier('+lastsupplier+'), '+fn);
                Data_Module.LogActLog('ORDERF','Create text file for supplier('+lastsupplier+'), '+fn);
                if Data_Module.fiLocalFTP.AsBoolean then
                begin
                  //
                  //  FTP writes auto archive and logistics file
                  //
                  //logistics
                  if lastlogisticsdirectory <> 'NONE' then
                  begin
                    fn:=lastlogisticsdirectory+'\'+ANSIReplaceStr(lastsuppliername,'\','')+'-'+lastsupplier+formatdatetime('yyyymmddhhmmss00',now)+'.ord';
                    AssignFile(tlf,fn);
                    Rewrite(tlf);
                    Hist.Append('Create text file for logistics('+lastsupplier+'), '+fn);
                    Data_Module.LogActLog('ORDERF','Create text file for logistics('+lastsupplier+'), '+fn);
                  end;
                  //Archive
                  fn:=lastdirectory+'\'+'Archive'+'\'+ANSIReplaceStr(lastsuppliername,'\','')+'-'+lastsupplier+formatdatetime('yyyymmddhhmmss00',now)+'.ord';
                  AssignFile(taf,fn);
                  Rewrite(taf);
                  Hist.Append('Create text file for archive('+lastsupplier+'), '+fn);
                  Data_Module.LogActLog('ORDERF','Create text file for archive('+lastsupplier+'), '+fn);
                end;
              end
              else
              begin
                //  No timestamp in file name
                //
                fn:=lastdirectory+'\'+ANSIReplaceStr(lastsuppliername,'\','')+'-'+lastsupplier+'.ord';
                AssignFile(tcf,fn);
                Rewrite(tcf);
                Hist.Append('Create text file for supplier('+lastsupplier+'), '+fn);
                Data_Module.LogActLog('ORDERF','Create text file for supplier('+lastsupplier+'), '+fn);
                if Data_Module.fiLocalFTP.AsBoolean then
                begin
                  //
                  //  FTP writes auto archive and logistics file
                  //
                  //logistics
                  if lastlogisticsdirectory <> 'NONE' then
                  begin
                    fn:=lastlogisticsdirectory+'\'+ANSIReplaceStr(lastsuppliername,'\','')+'-'+lastsupplier+'.ord';
                    AssignFile(tlf,fn);
                    Rewrite(tlf);
                    Hist.Append('Create text file for logistics('+lastsupplier+'), '+fn);
                    Data_Module.LogActLog('ORDERF','Create text file for logistics('+lastsupplier+'), '+fn);
                  end;
                  //Archive (Force date time for new file name)
                  fn:=lastdirectory+'\'+'Archive'+'\'+ANSIReplaceStr(lastsuppliername,'\','')+'-'+lastsupplier+formatdatetime('yyyymmddhhmmss00',now)+'.ord';
                  AssignFile(taf,fn);
                  Rewrite(taf);
                  Hist.Append('Create text file for archive('+lastsupplier+'), '+fn);
                  Data_Module.LogActLog('ORDERF','Create text file for archive('+lastsupplier+'), '+fn);
                end;
              end;
              fopen:=true;
            end;
          end;

          //
          //  Check for new renban Number, if so close last supplier sheet and create another, ONLY FOR WHEEL!!
          //
          if (lastRenban <> fieldbyname('VC_RENBAN_NUMBER').AsString) and (Data_Module.Inv_StoredProc.fieldbyname('Create Order Sheet').AsString = 'WHEEL') then
          begin
            // Special Order Form Close if open
            if not VarIsEmpty(excelOrder) then
            begin
              //  Create file in directory specified for each supplier
              excelOrder.ActiveWorkbook.SaveAs(lastdirectory+'\OS'+ANSIReplaceStr(lastsuppliername,'\','')+'-'+lastsupplier+'-'+lastRenban);
              if Data_Module.fiLocalFTP.AsBoolean then
              begin
                //logistics
                if lastlogisticsdirectory <> 'NONE' then
                  excelOrder.ActiveWorkbook.SaveAs(lastlogisticsdirectory+'\OS'+ANSIReplaceStr(lastsuppliername,'\','')+'-'+lastsupplier+'-'+lastRenban);
                //Archive
                excelOrder.ActiveWorkbook.SaveAs(lastdirectory+'\'+'Archive'+'\OS'+ANSIReplaceStr(lastsuppliername,'\','')+'-'+lastsupplier+'-'+lastRenban);
              end;

              excelOrder.Workbooks.Close;
              excelOrder.Quit;
              excelOrder:=Unassigned;
              page:=1;
            end;



            //Special Order Sheet Format
            excelOrder := createOleObject('Excel.Application');
            excelOrder.visible := FALSE; //debug
            excelOrder.DisplayAlerts := False;

            excelOrder.workbooks.open(Data_Module.TemplateDir+'OrderSheetTemplateWheel.xls');

            mysheetOrder := excelOrder.workSheets[1];
            Hist.Append('Create Order Sheet file for supplier, '+lastsupplier);
            Data_Module.LogActLog('ORDERS','Create Order Sheet file for supplier, '+lastsupplier);

            mysheetOrder.Cells[11,1].value  := Data_Module.Inv_StoredProc.fieldbyname('Supplier Name').AsString;
            mysheetOrder.Cells[11,2].value  := Data_Module.Inv_StoredProc.fieldbyname('Supplier Code').AsString;
            mysheetOrder.Cells[11,3].value  := fieldbyname('VC_RENBAN_NUMBER').AsString;
            year:=copy(formatdatetime('yyyy',now),1,3)+copy(fieldbyname('VC_FRS_NUMBER').AsString,1,1);
            mysheetOrder.Cells[8,8].value  := copy(fieldbyname('VC_FRS_NUMBER').AsString,2,2)+'/'+copy(fieldbyname('VC_FRS_NUMBER').AsString,4,2)+'/'+year;
            mysheetOrder.Cells[11,5].value  := page;
            mysheetOrder.Cells[13,2].value  := Data_Module.Inv_StoredProc.fieldbyname('Address').AsString;
            mysheetOrder.Cells[14,2].value  := Data_Module.Inv_StoredProc.fieldbyname('City').AsString+', '+Data_Module.Inv_StoredProc.fieldbyname('State').AsString+', '+Data_Module.Inv_StoredProc.fieldbyname('zip').AsString;

            lastRenban:=fieldbyname('VC_RENBAN_NUMBER').AsString;
            page:=1;
            o:=16;
          end;




          // Get Shipping date, skip weekends
          with Data_Module.INV_Check_DataSet do
          begin
            Close;
            CommandType := CmdStoredProc;
            CommandText := 'dbo.SELECT_PartShipDays;1';
            Parameters.Clear;
            Parameters.AddParameter.Name := '@PartNumber';
            Parameters.ParamValues['@PartNumber'] := Data_Module.Inv_DataSet.fieldbyname('VC_PART_NUMBER').AsString;
            Open;

            case DayOfTheWeek(now) of
            1:
              begin
                if fieldbyname('ShipM').AsInteger <> 0 then
                  ship:=GetShip(fieldbyname('ShipM').AsInteger)
                else
                  ship:=GetShip(fieldbyname('Ship').AsInteger);
              end;
            2:
              begin
                if fieldbyname('ShipT').AsInteger <> 0 then
                  ship:=GetShip(fieldbyname('ShipT').AsInteger)
                else
                  ship:=GetShip(fieldbyname('Ship').AsInteger);
              end;
            3:
              begin
                if fieldbyname('ShipW').AsInteger <> 0 then
                  ship:=GetShip(fieldbyname('ShipW').AsInteger)
                else
                  ship:=GetShip(fieldbyname('Ship').AsInteger);
              end;
            4:
              begin
                if fieldbyname('ShipTh').AsInteger <> 0 then
                  ship:=GetShip(fieldbyname('ShipTh').AsInteger)
                else
                  ship:=GetShip(fieldbyname('Ship').AsInteger);
              end;
            5:
              begin
                if fieldbyname('ShipF').AsInteger <> 0 then
                  ship:=GetShip(fieldbyname('ShipF').AsInteger)
                else
                  ship:=GetShip(fieldbyname('Ship').AsInteger);
              end;
            6:
              begin
                if fieldbyname('ShipS').AsInteger <> 0 then
                  ship:=GetShip(fieldbyname('ShipS').AsInteger)
                else
                  ship:=GetShip(fieldbyname('Ship').AsInteger);
              end;
            else
              ship:=GetShip(fieldbyname('Ship').AsInteger);
            end;
          end;

          //
          //  If the special order sheet fills
          //
          if ((Data_Module.Inv_StoredProc.fieldbyname('Create Order Sheet').AsString <> '') or
            (Data_Module.Inv_StoredProc.fieldbyname('Create Order Sheet').AsString <> ' ')) and (o>23) then
          begin
            //Save Current Sheet
            // Special Order Form Close if open
            if not VarIsEmpty(excelOrder) then
            begin
              //  Create file in directory specified for each supplier
              excelOrder.ActiveWorkbook.SaveAs(lastdirectory+'\OS'+ANSIReplaceStr(lastsuppliername,'\','')+'-'+lastsupplier+'-'+lastRenban+IntToStr(page));
              if Data_Module.fiLocalFTP.AsBoolean then
              begin
                //logistics
                if lastlogisticsdirectory <> 'NONE' then
                  excelOrder.ActiveWorkbook.SaveAs(lastlogisticsdirectory+'\OS'+ANSIReplaceStr(lastsuppliername,'\','')+'-'+lastsupplier+'-'+lastRenban);
                //Archive
                excelOrder.ActiveWorkbook.SaveAs(lastdirectory+'\'+'Archive'+'\OS'+ANSIReplaceStr(lastsuppliername,'\','')+'-'+lastsupplier+'-'+lastRenban+IntToStr(page));
              end;

              excelOrder.Workbooks.Close;
              excelOrder.Quit;
              excelOrder:=Unassigned;
              INC(page);
            end;

            //Create Another Sheet
            excelOrder := createOleObject('Excel.Application');
            excelOrder.visible := False;
            excelOrder.DisplayAlerts := False;

            if Data_Module.Inv_StoredProc.fieldbyname('Create Order Sheet').AsString = 'WHEEL' then
              excelOrder.workbooks.open(Data_Module.TemplateDir+'OrderSheetTemplateWheel.xls')
            else
              excelOrder.workbooks.open(Data_Module.TemplateDir+'OrderSheetTemplateTire.xls');
            mysheetOrder := excelOrder.workSheets[1];
            Hist.Append('Create Order Sheet file for supplier, '+lastsupplier);
            Data_Module.LogActLog('ORDERS','Create Order Sheet file for supplier, '+lastsupplier);

            mysheetOrder.Cells[11,1].value  := Data_Module.Inv_StoredProc.fieldbyname('Supplier Name').AsString;
            mysheetOrder.Cells[11,2].value  := Data_Module.Inv_StoredProc.fieldbyname('Supplier Code').AsString;
            mysheetOrder.Cells[11,3].value  := fieldbyname('VC_RENBAN_NUMBER').AsString;
            year:=copy(formatdatetime('yyyy',now),1,3)+copy(fieldbyname('VC_FRS_NUMBER').AsString,1,1);
            mysheetOrder.Cells[8,8].value  := copy(fieldbyname('VC_FRS_NUMBER').AsString,2,2)+'/'+copy(fieldbyname('VC_FRS_NUMBER').AsString,4,2)+'/'+year;
            mysheetOrder.Cells[11,5].value  := page;
            mysheetOrder.Cells[13,2].value  := Data_Module.Inv_StoredProc.fieldbyname('Address').AsString;
            mysheetOrder.Cells[14,2].value  := Data_Module.Inv_StoredProc.fieldbyname('City').AsString+', '+Data_Module.Inv_StoredProc.fieldbyname('State').AsString+', '+Data_Module.Inv_StoredProc.fieldbyname('zip').AsString;

            inc(page);
            o:=16;
          end;

          //
          //  Special Order Sheet
          //
          if not VarIsEmpty(excelOrder) then
          begin
            Data_Module.LogActLog('ORDERS','Order Sheet add item, '+fieldbyname('VC_PART_NUMBER').AsString);
            mysheetOrder.Cells[8,6].value  := FormatDateTime('mm/dd/yy',now+ship);
            mysheetOrder.Cells[o,1].value  := fieldbyname('VC_PART_NUMBER').AsString;
            mysheetOrder.Cells[o,2].value  := fieldbyname('VC_PARTS_NAME').AsString;
            mysheetOrder.Cells[o,4].value  := fieldbyname('VC_KANBAN_NUMBER').AsString;
            if Data_Module.Inv_StoredProc.fieldbyname('Create Order Sheet').AsString = 'WHEEL' then
            begin
              mysheetOrder.Cells[o,5].value  := fieldbyname('IN_1LOTQTY').AsString;
              mysheetOrder.Cells[o,7].value  := fieldbyname('IN_QTY').AsString;
            end
            else
            begin
              mysheetOrder.Cells[o,5].value  := fieldbyname('VC_RENBAN_NUMBER').AsString;
              mysheetOrder.Cells[o,7].value  := fieldbyname('IN_QTY').AsString;
            end;
            INC(o);
          end;


          if (fFileKind = fExcel) or (fFileKind = fBoth) then
          begin
            mysheet.Cells[i,1].value  := fieldbyname('VC_PART_NUMBER').AsString;
            mysheet.Cells[i,2].value  := fieldbyname('VC_FRS_NUMBER').AsString;
            mysheet.Cells[i,3].value  := fieldbyname('VC_RENBAN_NUMBER').AsString;
            mysheet.Cells[i,4].value  := fieldbyname('IN_QTY').AsString;
            mysheet.Cells[i,5].value  := FormatDateTime('mm/dd/yyyy',now+ship);

            Hist.Append('Add line: '+fieldbyname('VC_PART_NUMBER').AsString+','+fieldbyname('VC_FRS_NUMBER').AsString+','+fieldbyname('VC_RENBAN_NUMBER').AsString+','+fieldbyname('IN_QTY').AsString);
            Data_Module.LogActLog('ORDERF','Add line: '+fieldbyname('VC_PART_NUMBER').AsString+','+fieldbyname('VC_FRS_NUMBER').AsString+','+fieldbyname('VC_RENBAN_NUMBER').AsString+','+fieldbyname('IN_QTY').AsString);

            INC(i);
          end;
          if (fFileKind = fText) or (fFileKind = fBoth) then
          begin
            tcl:='';

            if sendsite then                              // Mod to add site supplier
            begin
              // Get supplier code from query
              //tcl:=Data_Module.fiSupplierCode.AsString;
              tcl:=fieldbyname('SiteSupplierCode').AsString;
              tcl:=tcl+fieldbyname('VC_SUPPLIER_CODE').AsString;
            end
            else
              tcl:=fieldbyname('VC_SUPPLIER_CODE').AsString;

            tcl:=tcl+fieldbyname('VC_FRS_NUMBER').AsString;
            tcl:=tcl+format('%8s',[fieldbyname('VC_RENBAN_NUMBER').AsString]);
            tcl:=tcl+fieldbyname('VC_PART_NUMBER').AsString;
            tcl:=tcl+format('%.5d',[fieldbyname('IN_QTY').AsInteger]);
            tcl:=tcl+FormatDateTime('yyyymmdd',now+ship);

            Writeln(tcf,tcl);
            if Data_Module.fiLocalFTP.AsBoolean then
            begin
              if lastlogisticsdirectory <> 'NONE' then
                Writeln(tlf,tcl);
              Writeln(taf,tcl);
            end;
            Hist.Append('Add line: '+tcl);
            Data_Module.LogActLog('ORDERF','Add line: '+tcl);
          end;


          // Update order record
          with Data_Module.INV_Order_StoredProc do
          begin
            Close;
            ProcedureName := 'dbo.UPDATE_ORDEROrderDate;1';
            Parameters.Clear;
            Parameters.AddParameter.Name := '@PartNumber';
            Parameters.ParamValues['@PartNumber'] := Data_Module.Inv_DataSet.fieldbyname('VC_PART_NUMBER').AsString;
            Parameters.AddParameter.Name := '@FRSNumber';
            Parameters.ParamValues['@FRSNumber'] := Data_Module.Inv_DataSet.fieldbyname('VC_FRS_NUMBER').AsString;
            Parameters.AddParameter.Name := '@OrderDate';
            Parameters.ParamValues['@OrderDate'] := FormatDateTime('yyyymmdd',now);
            Parameters.AddParameter.Name := '@ShipDate';
            Parameters.ParamValues['@ShipDate'] := FormatDateTime('yyyymmdd',now+ship);
            ExecProc;
          end;

          next;
        end;


        // Special Order Form Close if open
        if not VarIsEmpty(excelOrder) then
        begin
          //  Create file in directory specified for each supplier
          excelOrder.ActiveWorkbook.SaveAs(lastdirectory+'\OS'+ANSIReplaceStr(lastsuppliername,'\','')+'-'+lastsupplier+'-'+lastRenban+IntToStr(page));
          if Data_Module.fiLocalFTP.AsBoolean then
          begin
            //logistics
            if lastlogisticsdirectory <> 'NONE' then
              excelOrder.ActiveWorkbook.SaveAs(lastlogisticsdirectory+'\OS'+ANSIReplaceStr(lastsuppliername,'\','')+'-'+lastsupplier+'-'+lastRenban);
            //Archive
            excelOrder.ActiveWorkbook.SaveAs(lastdirectory+'\'+'Archive'+'\OS'+ANSIReplaceStr(lastsuppliername,'\','')+'-'+lastsupplier+'-'+lastRenban);
          end;

          excelOrder.Workbooks.Close;
          excelOrder.Quit;
          excelOrder:=Unassigned;
        end;

        if (fFileKind = fExcel) or (fFileKind = fBoth) then
        begin
          if not VarIsEmpty(excel) then
          begin
            //  Create file in directory specified for each supplier
            //  Use supplier name+WeekDate for filename
            if lasttimestamp then
            begin
              excel.ActiveWorkbook.SaveAs(lastdirectory+'\'+ANSIReplaceStr(lastsuppliername,'\','')+'-'+lastsupplier+'-'+formatdatetime('yyyymmddhhmmss00',now));
              if Data_Module.fiLocalFTP.AsBoolean then
              begin
                //logistics
                if lastlogisticsdirectory <> 'NONE' then
                  excel.ActiveWorkbook.SaveAs(lastlogisticsdirectory+'\'+ANSIReplaceStr(lastsuppliername,'\','')+'-'+lastsupplier+'-'+formatdatetime('yyyymmddhhmmss00',now));
                //Archive
                excel.ActiveWorkbook.SaveAs(lastdirectory+'\'+'Archive'+'\'+ANSIReplaceStr(lastsuppliername,'\','')+'-'+lastsupplier+'-'+formatdatetime('yyyymmddhhmmss00',now));
              end;
            end
            else
            begin
              excel.ActiveWorkbook.SaveAs(lastdirectory+'\'+ANSIReplaceStr(lastsuppliername,'\','')+'-'+lastsupplier);
              if Data_Module.fiLocalFTP.AsBoolean then
              begin
                //logistics
                if lastlogisticsdirectory <> 'NONE' then
                  excel.ActiveWorkbook.SaveAs(lastlogisticsdirectory+'\'+ANSIReplaceStr(lastsuppliername,'\','')+'-'+lastsupplier);
                //Archive
                excel.ActiveWorkbook.SaveAs(lastdirectory+'\'+'Archive'+'\'+ANSIReplaceStr(lastsuppliername,'\','')+'-'+lastsupplier+'-'+formatdatetime('yyyymmddhhmmss00',now));
              end;
            end;
            excel.Workbooks.Close;
            excel.Quit;
            excel:=Unassigned;
          end;
        end;
        if (fFileKind = fText) or (fFileKind = fBoth) then
        begin
          if fopen then
          begin
            CloseFile(tcf);
            if Data_Module.fiLocalFTP.AsBoolean then
            begin
              if lastlogisticsdirectory <> 'NONE' then
                CloseFile(tlf);
              CloseFile(taf);
            end;
            fopen:=FALSE;
          end;
        end;
        //
      end
      else
      begin
        Hist.Append('No orders to process');
        Data_Module.LogActLog('ORDERF','No orders to process');
      end;
    end;
    Data_Module.Inv_Connection.CommitTrans;
    OK_Button.Visible:=True;
    Data_Module.LogActLog('ORDERF','Order file generation complete');
    Hist.Append('Order file generation complete');
  except
    on e:exception do
    begin
      Data_Module.Inv_Connection.RollbackTrans;
      OK_Button.Visible:=True;
      Data_Module.LogActLog('ERROR','Unable to create order output files, '+e.message);
      Hist.Append('Unable to create order output files, '+e.message);
      Data_Module.LogActLog('ORDERF','Failed on order file creation');
      if not VarIsEmpty(excel) then
      begin
        excel.Workbooks.Close;
        excel.Quit;
        excel:=Unassigned;
      end;
      if fopen then
        CloseFile(tcf);
    end;
  end;
end;

function TOrderFormCreate_Form.GetShip(lead:integer):integer;
var
  x,y:integer;
begin
  result:=0;
  try
    // scan forward to find the first non-weekend non-holiday date
    with Data_Module.ALC_StoredProc do
    begin
      Close;
      ProcedureName := 'dbo.AD_GetSpecialDate';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@BeginDate';
      Parameters.ParamValues['@BeginDate'] := formatdatetime('yyyy-mm-dd 00:00:00',now);
      Parameters.AddParameter.Name := '@EndDate';
      Parameters.ParamValues['@EndDate'] := formatdatetime('yyyy-mm-dd 00:00:00',now+result+59);
      Parameters.AddParameter.Name := '@LineName';
      Parameters.ParamValues['@LineName'] := '';//Line_ComboBox.Text;
      Open;

      x:=0;
      y:=0;
      while y <> lead do
      begin
        // check for weekend
        if DayOfTheWeek((now+1)+x) < 6 then
        begin
          // check for holiday
          first;
          while not eof do
          begin
            if (trunc((now+1)+x) = trunc(fieldbyname('DATE').AsDateTime)) AND ( trim(fieldbyname('Date Status Abrv').AsString) = 'H' ) then
              break
            else
              next;
          end;

          if eof then
          begin
            // if not holiday and not weekend then increment as valid
            INC(y);
            INC(x);
          end
          else
            INC(x);

        end
        else
        begin
          INC(X);
        end;
        result:=x;
      end;
    end;
  except
    on e:exception do
    begin
      Data_Module.LogActLog('ERROR','Unable to get ship date offset, '+e.message);
      Hist.Append('Unable to get ship date offset, '+e.message);
    end;
  end;
end;

procedure TOrderFormCreate_Form.OK_ButtonClick(Sender: TObject);
begin
  Close;
end;

end.
