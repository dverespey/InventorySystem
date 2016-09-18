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
//  10/11/2002  David Verespey  Create Unit
//  12/16/2002  David Verespey  Update for new forecast information
//
unit ForecastBreakDown;


interface

uses Sysutils,dialogs,controls,ADOdb,ComObj,Variants,strutils;

type
  TWeekData = record
    WeekNumber    :integer;
    WeekDate      :string;
    WeekCount     :integer;
  end;

  TEntryRec = record
    Supplier      :string;
    Partnumber    :string;
    KanbanNumber  :string;
    Skip          :boolean;
    Weeks         :array[1..7] of TWeekData;
  end;

  TForecastBreakDown = class(TObject)
  private
    fFilename:string;
    fEntries:array of TEntryRec;
    fSupplierCode:string;
    fFirstWeekNumber:integer;
    fFirstWeekDate:string;
    fFileKind:TFileKind;
    function ScanLine(line:string;count:integer):boolean;
    function ScanPartnumber:boolean;
    procedure UpdateForecast;
    procedure DoPartNumberForecast(PN,WeekDate:string;FCCount,WeekNumber:integer;size:string);
  public
    function Execute:boolean;
  published
    property filename:string
    read fFilename
    write fFilename;

    property SupplierCode:string
    read fSupplierCode
    write fSupplierCode;

    property FileKind:TFileKind
    read fFileKind
    write fFileKind;
  end;

const
  SUPPLIER_OFF=1;
  SUPPLIER_SIZE=5;
  PARTNUMBER_OFF=6;
  PARTNUMBER_SIZE=12;
  KANBAN_OFF=18;
  KANBAN_SIZE=4;
  FORECASTW1_OFF=22;
  FORECASTW2_OFF=36;
  FORECASTW3_OFF=50;
  FORECASTW4_OFF=64;
  FORECASTW5_OFF=78;
  FORECASTW6_OFF=92;
  FORECASTW7_OFF=106;
  WEEKNUMBER_SIZE=2;
  WEEKDATE_SIZE=6;
  WEEKCOUNT_SIZE=6;

implementation

uses DataModule;

function TForecastBreakDown.Execute:boolean;
var
  fcf:Textfile;
  fcl:string;
  fopen:boolean;
  count,counter:integer;
  excel:        variant;
  lastsupplier: string;
  i:            integer;
  mySheet:      variant;
begin
  result:=TRUE;
  count:=0;
  counter:=0;
  lastsupplier:='';
  //open file
  try
    AssignFile(fcf, fFileName);
    Reset(fcf);
    //count records
    while not Seekeof(fcf) do
    begin
      Readln(fcf, fcl);
      INC(count)
    end;
    reset(fcf);
    SetLength(fEntries,count);
    //get all records, check for bad records on the fly
    try
      while not Seekeof(fcf) do
      begin
        Readln(fcf, fcl);
        if ScanLine(fcl,counter) then
        begin
          INC(counter);
        end;
      end;
      //scan for all suppliers, partnumbers and kanbannumber are in our database
      showmessage(IntToStr(counter)+' total records to process');
      SetLength(fEntries,counter);

      if ScanPartnumber then
      begin
        //clear forecast information from new week onward
        With Data_Module.Inv_StoredProc do
        Begin
          Close;
          ProcedureName := 'dbo.DELETE_ForecastInfoWeekDate;1';
          Parameters.Clear;
          Parameters.AddParameter.Name := '@WeekNumber';
          Parameters.ParamValues['@WeekNumber'] := fFirstWeekNumber;
          Parameters.AddParameter.Name := '@WeekDate';
          Parameters.ParamValues['@WeekDate'] := '20'+fFirstWeekDate;
          ExecProc;
        End;    //With

        //add to forecast table
        UpdateForecast;

        //
        //  Produce output files to send to suppliers
        //

        try
          with Data_Module.Inv_DataSet do
          begin
            //  Get week number only process files for the next week out
            Close;
            CommandType := CmdStoredProc;
            CommandText := 'dbo.SELECT_ForecastSupplier;1';
            Parameters.Clear;
            Parameters.AddParameter.Name := '@WeekDate';
            Parameters.ParamValues['@WeekDate'] := formatdatetime('yyyymmdd',now);
            Open;

            if fFileKind = fExcel then
            begin
              //
              //  Excel File
              //
              excel:=Unassigned;
              while not eof do
              begin
                if fieldbyname('VC_SUPPLIER_CODE').AsString <> lastsupplier then
                begin
                  lastsupplier:=fieldbyname('VC_SUPPLIER_CODE').AsString;
                  if not VarIsEmpty(excel) then
                  begin
                    //  Create file in directory specified for each supplier
                    //  Use supplier name+WeekDate for filename
                    if FileExists(Data_Module.INV_Forecast_DataSet.FieldbyName('Directory').AsString+'/'+ANSIReplaceStr(Data_Module.INV_Forecast_DataSet.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.INV_Forecast_DataSet.fieldbyname('Supplier Code').AsString+'-'+'Forecast') then
                      DeleteFile(Data_Module.INV_Forecast_DataSet.FieldbyName('Directory').AsString+'/'+ANSIReplaceStr(Data_Module.INV_Forecast_DataSet.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.INV_Forecast_DataSet.fieldbyname('Supplier Code').AsString+'-'+'Forecast');
                    excel.ActiveWorkbook.SaveAs(Data_Module.INV_Forecast_DataSet.FieldbyName('Directory').AsString+'/'+ANSIReplaceStr(Data_Module.INV_Forecast_DataSet.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.INV_Forecast_DataSet.fieldbyname('Supplier Code').AsString+'-'+'Forecast');
                    excel.Workbooks.Close;
                    excel.Quit;
                    excel:=Unassigned;
                  end;

                  excel := createOleObject('Excel.Application');
                  excel.visible := False;
                  excel.DisplayAlerts := False;
                  //excel.workbooks.add;
                  excel.workbooks.open(Data_Module.fiInputFileDir.AsString+'ForecastTemplate.xls');
                  mysheet := excel.workSheets[1];

                  i:=2;

                  Data_Module.INV_Forecast_DataSet.Close;
                  Data_Module.INV_Forecast_DataSet.CommandType := CmdStoredProc;
                  Data_Module.INV_Forecast_DataSet.CommandText := 'dbo.SELECT_SupplierInfo;1';
                  Data_Module.INV_Forecast_DataSet.Parameters.Clear;
                  Data_Module.INV_Forecast_DataSet.Parameters.AddParameter.Name := '@SupCode';
                  Data_Module.INV_Forecast_DataSet.Parameters.ParamValues['@SupCode'] := fieldbyname('VC_SUPPLIER_CODE').AsString;
                  Data_Module.INV_Forecast_DataSet.Open;
                end;

                mysheet.Cells[i,1].value  := fieldbyname('VC_SIZE_CODE').AsString;
                mysheet.Cells[i,2].value  := fieldbyname('VC_PART_NUMBER').AsString;
                mysheet.Cells[i,3].value  := fieldbyname('VC_WEEK_DATE').AsString;
                mysheet.Cells[i,4].value  := fieldbyname('IN_WEEK_NUMBER').AsString;
                mysheet.Cells[i,5].value  := fieldbyname('IN_QTY1').AsString;
                mysheet.Cells[i,6].value  := fieldbyname('IN_QTY2').AsString;
                mysheet.Cells[i,7].value  := fieldbyname('IN_QTY3').AsString;
                mysheet.Cells[i,8].value  := fieldbyname('IN_QTY4').AsString;
                mysheet.Cells[i,9].value  := fieldbyname('IN_QTY5').AsString;
                mysheet.Cells[i,10].value := fieldbyname('IN_QTY6').AsString;
                mysheet.Cells[i,11].value := fieldbyname('IN_QTY7').AsString;

                INC(i);
                next;
              end;
              if not VarIsEmpty(excel) then
              begin
                //  Create file in directory specified for each supplier
                //  Use supplier name+WeekDate for filename
                if FileExists(Data_Module.INV_Forecast_DataSet.FieldbyName('Directory').AsString+'/'+ANSIReplaceStr(Data_Module.INV_Forecast_DataSet.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.INV_Forecast_DataSet.fieldbyname('Supplier Code').AsString+'-'+'Forecast') then
                  DeleteFile(Data_Module.INV_Forecast_DataSet.FieldbyName('Directory').AsString+'/'+ANSIReplaceStr(Data_Module.INV_Forecast_DataSet.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.INV_Forecast_DataSet.fieldbyname('Supplier Code').AsString+'-'+'Forecast');
                excel.ActiveWorkbook.SaveAs(Data_Module.INV_Forecast_DataSet.FieldbyName('Directory').AsString+'/'+ANSIReplaceStr(Data_Module.INV_Forecast_DataSet.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.INV_Forecast_DataSet.fieldbyname('Supplier Code').AsString+'-'+'Forecast');
                excel.Workbooks.Close;
                excel.Quit;
                excel:=Unassigned;
              end;
            end
            else
            begin
              //
              //  Text file CR/LF delimited
              //
              fopen:=False;
              while not eof do
              begin

                if fieldbyname('VC_SUPPLIER_CODE').AsString <> lastsupplier then
                begin
                  lastsupplier:=fieldbyname('VC_SUPPLIER_CODE').AsString;

                  if fopen then
                  begin
                    CloseFile(fcf);
                  end;


                  Data_Module.INV_Forecast_DataSet.Close;
                  Data_Module.INV_Forecast_DataSet.CommandType := CmdStoredProc;
                  Data_Module.INV_Forecast_DataSet.CommandText := 'dbo.SELECT_SupplierInfo;1';
                  Data_Module.INV_Forecast_DataSet.Parameters.Clear;
                  Data_Module.INV_Forecast_DataSet.Parameters.AddParameter.Name := '@SupCode';
                  Data_Module.INV_Forecast_DataSet.Parameters.ParamValues['@SupCode'] := fieldbyname('VC_SUPPLIER_CODE').AsString;
                  Data_Module.INV_Forecast_DataSet.Open;


                  AssignFile(fcf,Data_Module.INV_Forecast_DataSet.FieldbyName('Directory').AsString+'/'+ANSIReplaceStr(Data_Module.INV_Forecast_DataSet.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.INV_Forecast_DataSet.fieldbyname('Supplier Code').AsString+'.frc');
                  Rewrite(fcf);
                  fopen:=true;

                end;

                fcl:='';

                fcl:=fieldbyname('VC_SUPPLIER_CODE').AsString;
                fcl:=fcl+fieldbyname('VC_PART_NUMBER').AsString;
                fcl:=fcl+fieldbyname('VC_WEEK_DATE').AsString;
                fcl:=fcl+Format('%.2d',[fieldbyname('IN_WEEK_NUMBER').AsInteger]);
                fcl:=fcl+Format('%.5d',[fieldbyname('IN_QTY1').AsInteger]);
                fcl:=fcl+Format('%.5d',[fieldbyname('IN_QTY2').AsInteger]);
                fcl:=fcl+Format('%.5d',[fieldbyname('IN_QTY3').AsInteger]);
                fcl:=fcl+Format('%.5d',[fieldbyname('IN_QTY4').AsInteger]);
                fcl:=fcl+Format('%.5d',[fieldbyname('IN_QTY5').AsInteger]);
                fcl:=fcl+Format('%.5d',[fieldbyname('IN_QTY6').AsInteger]);
                fcl:=fcl+Format('%.5d',[fieldbyname('IN_QTY7').AsInteger]);

                Writeln(fcf,fcl);

                Next;
              end;
              if fopen then
                CloseFile(fcf);
            end;
          end;
        except
          on e:exception do
          begin
            Data_Module.LogActLog('ERROR','Unable to create forecast output files, '+e.message);
            result:=false;

            if not VarIsEmpty(excel) then
            begin
              excel.Workbooks.Close;
              excel.Quit;
              excel:=Unassigned;
            end;
          end;
        end;
      end
      else
        result:=FALSE;
    except
      on e:exception do
      begin
        Data_Module.LogActLog('ERROR','Unable to load forecast, '+e.message);
        result:=false;
      end;
    end;
  finally
    CloseFile(fcf);
  end;
  //put data into forecast table
end;


function TForecastBreakdown.ScanLine(line:string;count:integer):boolean;
begin
  result:=TRUE;
  try
    if copy(line,SUPPLIER_OFF,SUPPLIER_SIZE) = fSupplierCode then
    begin
      fEntries[count].Supplier:=copy(line,SUPPLIER_OFF,SUPPLIER_SIZE);
      fEntries[count].Partnumber:=copy(line,PARTNUMBER_OFF,PARTNUMBER_SIZE);
      fEntries[count].KanbanNumber:=copy(line,KANBAN_OFF,KANBAN_SIZE);
      fEntries[count].Skip:=False;
      fEntries[count].Weeks[1].WeekNumber:=StrToInt(copy(line,FORECASTW1_OFF,WEEKNUMBER_SIZE));
      fEntries[count].Weeks[1].WeekDate:=copy(line,FORECASTW1_OFF+WEEKNUMBER_SIZE,WEEKDATE_SIZE);
      fFirstWeekNumber:=StrToInt(copy(line,FORECASTW1_OFF,WEEKNUMBER_SIZE));
      fFirstWeekDate:=copy(line,FORECASTW1_OFF+WEEKNUMBER_SIZE,WEEKDATE_SIZE);
      if copy(line,FORECASTW1_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE) <> '     -' then
        fEntries[count].Weeks[1].WeekCount:=StrToInt(copy(line,FORECASTW1_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE))
      else
        fEntries[count].Weeks[1].WeekCount:=0;

      fEntries[count].Weeks[2].WeekNumber:=StrToInt(copy(line,FORECASTW2_OFF,WEEKNUMBER_SIZE));
      fEntries[count].Weeks[2].WeekDate:=copy(line,FORECASTW2_OFF+WEEKNUMBER_SIZE,WEEKDATE_SIZE);
      if copy(line,FORECASTW2_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE) <> '     -' then
        fEntries[count].Weeks[2].WeekCount:=StrToInt(copy(line,FORECASTW2_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE))
      else
        fEntries[count].Weeks[2].WeekCount:=0;

      fEntries[count].Weeks[3].WeekNumber:=StrToInt(copy(line,FORECASTW3_OFF,WEEKNUMBER_SIZE));
      fEntries[count].Weeks[3].WeekDate:=copy(line,FORECASTW3_OFF+WEEKNUMBER_SIZE,WEEKDATE_SIZE);
      if copy(line,FORECASTW3_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE) <> '     -' then
        fEntries[count].Weeks[3].WeekCount:=StrToInt(copy(line,FORECASTW3_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE))
      else
        fEntries[count].Weeks[3].WeekCount:=0;

      fEntries[count].Weeks[4].WeekNumber:=StrToInt(copy(line,FORECASTW4_OFF,WEEKNUMBER_SIZE));
      fEntries[count].Weeks[4].WeekDate:=copy(line,FORECASTW4_OFF+WEEKNUMBER_SIZE,WEEKDATE_SIZE);
      if copy(line,FORECASTW4_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE) <> '     -' then
        fEntries[count].Weeks[4].WeekCount:=StrToInt(copy(line,FORECASTW4_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE))
      else
        fEntries[count].Weeks[4].WeekCount:=0;

      fEntries[count].Weeks[5].WeekNumber:=StrToInt(copy(line,FORECASTW5_OFF,WEEKNUMBER_SIZE));
      fEntries[count].Weeks[5].WeekDate:=copy(line,FORECASTW5_OFF+WEEKNUMBER_SIZE,WEEKDATE_SIZE);
      if copy(line,FORECASTW5_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE) <> '     -' then
        fEntries[count].Weeks[5].WeekCount:=StrToInt(copy(line,FORECASTW5_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE))
      else
        fEntries[count].Weeks[5].WeekCount:=0;

      fEntries[count].Weeks[6].WeekNumber:=StrToInt(copy(line,FORECASTW6_OFF,WEEKNUMBER_SIZE));
      fEntries[count].Weeks[6].WeekDate:=copy(line,FORECASTW6_OFF+WEEKNUMBER_SIZE,WEEKDATE_SIZE);
      if copy(line,FORECASTW6_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE) <> '     -' then
        fEntries[count].Weeks[6].WeekCount:=StrToInt(copy(line,FORECASTW6_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE))
      else
        fEntries[count].Weeks[6].WeekCount:=0;

      fEntries[count].Weeks[7].WeekNumber:=StrToInt(copy(line,FORECASTW7_OFF,WEEKNUMBER_SIZE));
      fEntries[count].Weeks[7].WeekDate:=copy(line,FORECASTW7_OFF+WEEKNUMBER_SIZE,WEEKDATE_SIZE);
      if copy(line,FORECASTW7_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE) <> '     -' then
        fEntries[count].Weeks[7].WeekCount:=StrToInt(copy(line,FORECASTW7_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE))
      else
        fEntries[count].Weeks[7].WeekCount:=0;

    end
    else
      result:=false;
  except
    on e:exception do
    begin
      Showmessage('File read error, '+e.Message+', import failed');
      Data_Module.LogActLog('ERROR','File read error, '+e.Message+', import failed');
      Raise;
    end;
  end;
end;


function TForecastBreakdown.ScanPartnumber:boolean;
var
  i,x,z,skip:integer;
  excel,mysheet:variant;
begin
  result:=true;
  skip:=0;

  with Data_Module.Inv_DataSet do
  begin
    try
      Close;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.SELECT_ForecastDetail;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@AssyCode';
      Parameters.ParamValues['@AssyCode'] := '';
      //Parameters.ParamValues['@AssyCode'] := fEntries[i].PArtNumber;
      Open;
      for i:=0 to High(fEntries) do
      begin
        if not Locate('Assembly Part Number Code',fEntries[i].PArtNumber,[]) then
        begin
          // Partnumber does not exist, skip
          fEntries[i].Skip:=True;
          INC(skip);
        end;
      end;
    except
      on e:exception do
      begin
        ShowMessage('Error on INV_FORECAST_DETAIL_INF table access, '+e.Message);
        Data_Module.LogActLog('ERROR','FORECAST: Error on INV_FORECAST_DETAIL_INF table access, '+e.Message);
        result:=false;
      end;
    end;
  end;

  if skip <> 0 then
  begin
    if messagedlg('There are '+IntToStr(skip)+', skipped forecast records. Continue processing?',
                    mtConfirmation, [mbYes, mbNo], 0) = mrNo then
    begin
      result:=False;
      exit;
    end;

    // create error xls form
    excel := createOleObject('Excel.Application');
    excel.visible := False;
    excel.DisplayAlerts := False;
    //excel.workbooks.add;
    excel.workbooks.open(Data_Module.fiInputFileDir.AsString+'ForecastError.xls');
    mysheet := excel.workSheets[1];
    z:=4;
    for x:=0 to high(fEntries) do
    begin
      if fEntries[x].Skip then
      begin
        mysheet.Cells[z,1].value := fEntries[x].Partnumber;
        INC(z);
      end;
    end;
    excel.ActiveWorkbook.SaveAs(Data_Module.fiInputFileDir.AsString+'ForecastError'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
    excel.Workbooks.Close;
    excel.Quit;
    excel:=Unassigned;
  end;

  ShowMessage(IntToStr(high(fEntries)-skip)+' forecast records to be added');
end;

procedure TForecastBreakdown.UpdateForecast;
var
  i,j:integer;
  sizecode:string;
begin
  for i:=0 to High(fEntries) do
  begin
    if (not fEntries[i].Skip) then
    begin
      //insert or update record
      try
        for j:=1 to 7 do
          With Data_Module.Inv_StoredProc do
          Begin
            begin
            Close;
            ProcedureName := 'dbo.INSERTUPDATE_ForecastInfo;1';
            Parameters.Clear;
            Parameters.AddParameter.Name := '@Supplier';
            Parameters.ParamValues['@Supplier'] := fEntries[i].Supplier;
            Parameters.AddParameter.Name := '@PartNumber';
            Parameters.ParamValues['@PartNumber'] := fEntries[i].Partnumber;
            Parameters.AddParameter.Name := '@Kanban';
            Parameters.ParamValues['@Kanban'] := fEntries[i].KanbanNumber;
            Parameters.AddParameter.Name := '@WeekNumber';
            Parameters.ParamValues['@WeekNumber'] := fEntries[i].Weeks[j].WeekNumber;
            Parameters.AddParameter.Name := '@WeekDate';
            Parameters.ParamValues['@WeekDate'] := '20'+fEntries[i].Weeks[j].WeekDate;
            Parameters.AddParameter.Name := '@Count';
            Parameters.ParamValues['@Count'] := fEntries[i].Weeks[j].WeekCount;
            ExecProc;
          End;    //With

          try
            with Data_Module.INV_DataSet do
            begin
              //Get ratio data
              Close;
              CommandType := CmdStoredProc;
              CommandText := 'dbo.SELECT_ForecastDetail;1';
              Parameters.Clear;
              Parameters.AddParameter.Name := '@AssyCode';
              Parameters.ParamValues['@AssyCode'] := fEntries[i].PArtNumber;
              Open;

              sizecode:=FieldByName('Size Code').AsString;

              // scan through the part numbers for this assembly code and assign
              if length(FieldByName('Tire Part Number Code').AsString) = 12 then
              begin
                DoPartNumberForecast( FieldByName('Tire Part Number Code').AsString,
                                      fEntries[i].Weeks[j].WeekDate,
                                      fEntries[i].Weeks[j].WeekCount,
                                      fEntries[i].Weeks[j].WeekNumber,
                                      sizeCode);

              end;

              if length(FieldByName('Wheel Part Number Code').AsString) = 12 then
              begin
                DoPartNumberForecast( FieldByName('Wheel Part Number Code').AsString,
                                      fEntries[i].Weeks[j].WeekDate,
                                      fEntries[i].Weeks[j].WeekCount,
                                      fEntries[i].Weeks[j].WeekNumber,
                                      sizecode);
              end;

            end;
          except
            on e:exception do
            begin
              ShowMessage('Error on INV_FORECAST_DETAIL_INF table access, '+e.Message);
              Data_Module.LogActLog('ERROR','FORECAST: Error on INV_FORECAST_DETAIL_INF table access, '+e.Message);
              break;
            end
          end;

          Data_Module.LogActLog('FORECAST','Insert/Update forecast information for part number, '+fEntries[i].Partnumber);
        end;
      except
        on e:exception do
        Begin
          Data_Module.LogActLog('ERROR', 'FAILED Insert/Update forecast information for part number, '+fEntries[i].Partnumber + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName, 0);
          exit;
        End;      //Except
      end;
      //process breakdown
    end;
  end;
end;

procedure TForecastBreakdown.DoPartNumberForecast(PN,WeekDate:string;FCCount,Weeknumber:integer;size:string);
var
  workday: array[1..7] of boolean;
  dayforecast: array[1..7] of integer;
  line,supplier:string;
  days,ratiocount,i:integer;
begin
  workday[1]:=true;
  workday[2]:=true;
  workday[3]:=true;
  workday[4]:=true;
  workday[5]:=true;
  workday[6]:=false;
  workday[7]:=false;
  days:=5;

  try
    // Get PartMaster Information
    with Data_Module.INV_Forecast_DataSet do
    begin
      //
      //  Get assembly line
      //
      Close;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.SELECT_PartsStockInfo;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@InvMgmtReport';
      Parameters.ParamValues['@InvMgmtReport'] := 'N';
      Parameters.AddParameter.Name := '@Supplier';
      Parameters.ParamValues['@Supplier'] := '';
      Parameters.AddParameter.Name := '@PartNum';
      Parameters.ParamValues['@PartNum'] := PN;
      Open;
      if FieldByName('Car / Truck').AsString = 'Car' then
        line:='P'
      else
        line:='T';

      supplier:=FieldByName('Supplier Code').AsString;
      //
      //  Find any changes to a normal weekly schedule
      //
      Close;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.SELECT_OvertimeHolidayWeek;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@WeekNumber';
      Parameters.ParamValues['@WeekNumber'] := WeekNumber;
      Parameters.AddParameter.Name := '@Line';
      Parameters.ParamValues['@Line'] := Line;
      Open;
      if recordCount > 0 then
      begin
        while not eof do
        begin
          if fieldByName('VC_OVERTIME_HOLIDAY').AsString = 'H' then
          begin
            workday[FieldByName('IN_DAY_NUMBER').AsInteger]:=False;
            DEC(days);
          end
          else
          begin
            workday[FieldByName('IN_DAY_NUMBER').AsInteger]:=True;
            INC(days);
          end;
          next;
        end;
      end;
      Close;
      if days > 0 then
        ratiocount:=FCCount div days
      else
        ratiocount := 0;
      for i:=1 to 7 do
      begin
        if workday[i] then
          dayforecast[i]:=ratiocount
        else
          dayforecast[i]:=0;
      end;
    end;

    //
    //  Insert/Update record for this partnumber on this week
    //
    With Data_Module.Inv_StoredProc do
    Begin
      Close;
      ProcedureName := 'dbo.INSERTUPDATE_BreakdownForecastInfo;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@WeekNumber';
      Parameters.ParamValues['@WeekNumber'] := WeekNumber;
      Parameters.AddParameter.Name := '@WeekDate';
      Parameters.ParamValues['@WeekDate'] := '20'+WeekDate;
      Parameters.AddParameter.Name := '@Supplier';
      Parameters.ParamValues['@Supplier'] := Supplier;
      Parameters.AddParameter.Name := '@PartNumber';
      Parameters.ParamValues['@PartNumber'] := PN;
      Parameters.AddParameter.Name := '@SizeCode';
      Parameters.ParamValues['@SizeCode'] := size;
      Parameters.AddParameter.Name := '@Qty1';
      Parameters.ParamValues['@Qty1'] := dayforecast[1];
      Parameters.AddParameter.Name := '@Qty2';
      Parameters.ParamValues['@Qty2'] := dayforecast[2];
      Parameters.AddParameter.Name := '@Qty3';
      Parameters.ParamValues['@Qty3'] := dayforecast[3];
      Parameters.AddParameter.Name := '@Qty4';
      Parameters.ParamValues['@Qty4'] := dayforecast[4];
      Parameters.AddParameter.Name := '@Qty5';
      Parameters.ParamValues['@Qty5'] := dayforecast[5];
      Parameters.AddParameter.Name := '@Qty6';
      Parameters.ParamValues['@Qty6'] := dayforecast[6];
      Parameters.AddParameter.Name := '@Qty7';
      Parameters.ParamValues['@Qty7'] := dayforecast[7];
      ExecProc;
    End;

  except
    on e:exception do
    begin
      Data_Module.LogActLog('ERROR','Failure on partnumber forecast update/insert, '+PN);
      Raise;
    end
  end;
end;

end.
