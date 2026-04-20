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
//  01/10/2003  David Verespey  Create Unit
//
unit OrderFormCreate;

interface

uses Sysutils,dialogs,controls,ADOdb,ComObj,Variants,strutils,DataModule;

type
  TOrderFormCreate = class(TObject)
  private
    fFileKind:TFileKind;
  public
    function Execute:boolean;
  published
    property FileKind:TFileKind
    read fFileKind
    write fFileKind;
  end;
implementation


function TOrderFormCreate.Execute:boolean;
var
  tcf:textfile;
  tcl:string;
  fopen:boolean;
  lastsupplier:string;
  excel:        variant;
  mySheet:      variant;
  i:integer;
begin
  lastsupplier:='';

  //
  //  Produce output files to send to suppliers
  //
  try
    with Data_Module.Inv_DataSet do
    begin
      //  Get non-ordered orders
      Close;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.SELECT_OrderNotOrdered;1';
      Parameters.Clear;
      Open;

      if recordcount > 0 then
      begin
        if fFileKind = fExcel then
        begin
          excel:=Unassigned;
          ShowMessage('There are ('+IntToStr(recordcount)+') records to process');
          while not eof do
          begin
            if fieldbyname('VC_SUPPLIER_CODE').AsString <> lastsupplier then
            begin
              lastsupplier:=fieldbyname('VC_SUPPLIER_CODE').AsString;
              if not VarIsEmpty(excel) then
              begin
                //  Create file in directory specified for each supplier
                //  Use supplier name+OrderDate(16 varchar) for filename
                excel.ActiveWorkbook.SaveAs(Data_Module.INV_Forecast_DataSet.FieldbyName('Directory').AsString+'/'+ANSIReplaceStr(Data_Module.INV_Forecast_DataSet.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.INV_Forecast_DataSet.fieldbyname('Supplier Code').AsString+'-'+formatdatetime('yyyymmddhhmmss00',now));
                excel.Workbooks.Close;
                excel.Quit;
                excel:=Unassigned;
              end;

              excel := createOleObject('Excel.Application');
              excel.visible := False;
              excel.DisplayAlerts := False;
              excel.workbooks.open(Data_Module.fiInputFileDir.AsString+'OrderTemplate.xls');
              mysheet := excel.workSheets[1];

              i:=10;

              Data_Module.INV_Forecast_DataSet.Close;
              Data_Module.INV_Forecast_DataSet.CommandType := CmdStoredProc;
              Data_Module.INV_Forecast_DataSet.CommandText := 'dbo.SELECT_SupplierInfo;1';
              Data_Module.INV_Forecast_DataSet.Parameters.Clear;
              Data_Module.INV_Forecast_DataSet.Parameters.AddParameter.Name := '@SupCode';
              Data_Module.INV_Forecast_DataSet.Parameters.ParamValues['@SupCode'] := fieldbyname('VC_SUPPLIER_CODE').AsString;
              Data_Module.INV_Forecast_DataSet.Open;

              mysheet.Cells[1,6].value  := Data_Module.INV_Forecast_DataSet.fieldbyname('Supplier Name').AsString;
              mysheet.Cells[2,6].value  := Data_Module.INV_Forecast_DataSet.fieldbyname('Address').AsString;
            end;

            mysheet.Cells[i,1].value  := fieldbyname('VC_PART_NUMBER').AsString;
            mysheet.Cells[i,2].value  := fieldbyname('VC_FRS_NUMBER').AsString;
            mysheet.Cells[i,3].value  := fieldbyname('VC_RENBAN_NUMBER').AsString;
            mysheet.Cells[i,4].value  := fieldbyname('IN_QTY').AsString;

            //update record with order date
            with Data_Module.Inv_StoredProc do
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
              ExecProc;
            end;

            INC(i);
            next;
          end;
          if not VarIsEmpty(excel) then
          begin
            //  Create file in directory specified for each supplier
            //  Use supplier name+WeekDate for filename
            excel.ActiveWorkbook.SaveAs(Data_Module.INV_Forecast_DataSet.FieldbyName('Directory').AsString+'/'+ANSIReplaceStr(Data_Module.INV_Forecast_DataSet.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.INV_Forecast_DataSet.fieldbyname('Supplier Code').AsString+'-'+formatdatetime('yyyymmddhhmmss00',now));
            excel.Workbooks.Close;
            excel.Quit;
            excel:=Unassigned;
          end;
        end
        else
        begin
          ShowMessage('There are ('+IntToStr(recordcount)+') records to process');
          fopen:=False;
          while not eof do
          begin
            if fieldbyname('VC_SUPPLIER_CODE').AsString <> '' then
            begin
              if fieldbyname('VC_SUPPLIER_CODE').AsString <> lastsupplier then
              begin
                lastsupplier:=fieldbyname('VC_SUPPLIER_CODE').AsString;


                if fopen then
                begin
                  CloseFile(tcf);
                end;

                Data_Module.INV_Forecast_DataSet.Close;
                Data_Module.INV_Forecast_DataSet.CommandType := CmdStoredProc;
                Data_Module.INV_Forecast_DataSet.CommandText := 'dbo.SELECT_SupplierInfo;1';
                Data_Module.INV_Forecast_DataSet.Parameters.Clear;
                Data_Module.INV_Forecast_DataSet.Parameters.AddParameter.Name := '@SupCode';
                Data_Module.INV_Forecast_DataSet.Parameters.ParamValues['@SupCode'] := fieldbyname('VC_SUPPLIER_CODE').AsString;
                Data_Module.INV_Forecast_DataSet.Open;


                AssignFile(tcf,Data_Module.INV_Forecast_DataSet.FieldbyName('Directory').AsString+'/'+ANSIReplaceStr(Data_Module.INV_Forecast_DataSet.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.INV_Forecast_DataSet.fieldbyname('Supplier Code').AsString+formatdatetime('yyyymmdd',now)+'.ord');
                Rewrite(tcf);
                fopen:=true;
                //Hist.Append('Create file for supplier, '+fieldbyname('VC_SUPPLIER_CODE').AsString);
              end;

              tcl:='';

              tcl:=fieldbyname('VC_SUPPLIER_CODE').AsString;
              tcl:=tcl+fieldbyname('VC_FRS_NUMBER').AsString;
              tcl:=tcl+fieldbyname('VC_RENBAN_NUMBER').AsString;
              tcl:=tcl+fieldbyname('VC_PART_NUMBER').AsString;
              tcl:=tcl+format('%.5d',[fieldbyname('IN_QTY').AsInteger]);

              Writeln(tcf,tcl);
            end
            else
            begin
              ShowMessage('Blank Supplier, invalid order');
            end;

            //update record with order date
            with Data_Module.Inv_StoredProc do
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
              ExecProc;
            end;

            Next;
          end;
          if fopen then
            CloseFile(tcf);
        end;
      end
      else
      begin
        ShowMessage('No orders to process');
      end;
    end;
    result:=True;
  except
    on e:exception do
    begin
      Data_Module.LogActLog('ERROR','Unable to create order output files, '+e.message);
      result:=false;

      if not VarIsEmpty(excel) then
      begin
        excel.Workbooks.Close;
        excel.Quit;
        excel:=Unassigned;
      end;
    end;
  end;
end;

end.
