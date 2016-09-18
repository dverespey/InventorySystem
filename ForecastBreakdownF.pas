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
//
unit ForecastBreakdownF;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, History, Datamodule,ADOdb,ComObj,strutils,dateutils;

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
    Weeks         :array[1..14] of TWeekData;
  end;

  TForecastBreakdown_Form = class(TForm)
    Hist: THistory;
    OKButton: TButton;
    procedure OKButtonClick(Sender: TObject);
  private
    { Private declarations }
    fFilename:string;
    fEntries:array of TEntryRec;
    fSupplierCode:string;
    fFirstWeekNumber:integer;
    fFirstWeekDate:string;
    fHistDate:string;
    fFileKind: TFileKind;
    fClosed:boolean;
    function ScanLine(line:string;count:integer):boolean;
    function ScanPartnumber:boolean;
    procedure UpdateForecast;
    procedure UpdateUsage;
    procedure DoPartNumberForecast(PN,WeekDate:string;FCCount,WeekNumber:integer);
    function HistoryForecast(partno:string):integer;
    procedure DeleteBreakdown(partnum:string);

  public
    { Public declarations }
    function Execute:boolean;
  published
    property filename:string
    read fFilename
    write fFilename;

    property SupplierCode:string
    read fSupplierCode
    write fSupplierCode;
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
  FORECASTW8_OFF=120;
  FORECASTW9_OFF=134;
  FORECASTW10_OFF=148;
  FORECASTW11_OFF=162;
  FORECASTW12_OFF=176;
  FORECASTW13_OFF=190;
  FORECASTW14_OFF=204;
  WEEKNUMBER_SIZE=2;
  WEEKDATE_SIZE=6;
  WEEKCOUNT_SIZE=6;


var
  ForecastBreakdown_Form: TForecastBreakdown_Form;

implementation

uses ForecastCamexreport;

{$R *.dfm}

procedure TForecastBreakdown_Form.DeleteBreakdown(partnum:string);
begin
    With Data_Module.Inv_StoredProc do
    Begin
      Close;
      ProcedureName := 'dbo.DELETE_ForecastInfo;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@WeekDate';
      Parameters.ParamValues['@WeekDate'] := '20'+fFirstWeekDate;
      Parameters.AddParameter.Name := '@HistWeekDate';
      Parameters.ParamValues['@HistWeekDate'] := fHistDate;
      Parameters.AddParameter.Name := '@PartNumber';
      Parameters.ParamValues['@PartNumber'] := partnum;
      ExecProc;
    End;    //With
end;

function TForecastBreakdown_Form.Execute:boolean;
var
  fcf,tcf,taf:Textfile;
  fcl,tcl,fn,data:string;
  fopen:boolean;
  count,counter:integer;
  excel:        variant;
  lastsupplier: string;
  i:            integer;
  mySheet:      variant;
  sendsite:     boolean;
  EDIfile:      boolean;
  EDILine:      string;
  EDIcount,EDIdate,EDIweek: string;
  ForecastReport: TForecastCAMEXReport;
begin
  result:=TRUE;
  count:=0;
  counter:=0;
  lastsupplier:='';
  fclosed:=False;
  OKButton.Visible:=False;
  sendsite:=FALSE;

  //open file
  try
    AssignFile(fcf, fFileName);
    Reset(fcf);

    //determine file type
    //
    //  Set EDI flag if EDI 830 file
    //
    Readln(fcf, fcl);
    if pos('ISA',fcl) > 0 then
    begin
      Readln(fcf, fcl);
      Readln(fcf, fcl);
      data:=copy(fcl,4,3);
      if data='830' then
        EDIfile:=TRUE
      else
      begin
        Hist.Append('EDI file type='+data+', expected type=830. Import fail.');
        Data_Module.LogActLog('FORECAST','EDI file type='+data+', expected type=830. Import fail.');
        exit;
      end;
    end
    else
      EDIfile:=FALSE;

    Reset(fcf);
    //count records
    while not Seekeof(fcf) do
    begin
      Readln(fcf, fcl);
      if EDIfile then
      begin
        //if EDI then LIN identifier indicates a partnumber loop
        data:=copy(fcl,1,3);
        if data='LIN' then
          INC(count);
      end
      else
        INC(count);
    end;

    reset(fcf);
    SetLength(fEntries,count);
    //get all records, check for bad records on the fly
    try
      if EDIfile then
      begin

        while true do
        begin
          while (data <> 'LIN') and (data <> 'CTT') do
          begin
            Readln(fcf, fcl);
            data:=copy(fcl,1,3);
          end;

          if data = 'CTT' then
          begin
            // end of records
            break;
          end;

          // 5 digit supplier + 12 digit assy part number + 4 digit bogus kanban number
          EDILine:=Data_Module.fiSupplierCode.AsString+copy(fcl,9,12)+copy(fcl,25,4);
          // Build the old style line
          //
          // Point to data line
          data:='';
          while data <> 'FST' do
          begin
            Readln(fcf, fcl);
            data:=copy(fcl,1,3);
          end;
          // Get all data points
          while data = 'FST' do
          begin
            // Get count
            fcl:=copy(fcl,pos('*',fcl)+1,length(fcl)); //strip line header
            EDICount:=format('%.6d',[StrToInt(copy(fcl,1,pos('*',fcl)-1))]);
            fcl:=copy(fcl,pos('*',fcl)+1,length(fcl)); //strip count
            fcl:=copy(fcl,5,length(fcl)); //strip more characters
            EDIdate:=copy(fcl,3,6);
            EDIweek:=copy(fcl,26,2);
            EDILine:=EDILine+EDIWeek+EDIDate+EDICount;
            // Get next line
            Readln(fcf, fcl);
            data:=copy(fcl,1,3);
          end;
          if ScanLine(EDILine,counter) then
          begin
            INC(counter);
          end;
        end;
      end
      else
      begin
        while not Seekeof(fcf) do
        begin
          // Non-EDI pass the line and parse it out
          Readln(fcf, fcl);
          if ScanLine(fcl,counter) then
          begin
            INC(counter);
          end;
        end;
      end;

      //scan for all suppliers, partnumbers and kanbannumber are in our database
      Hist.Append(IntToStr(counter)+' total records to process');
      Data_Module.LogActLog('FORECAST',IntToStr(counter)+' total records to process');
      SetLength(fEntries,counter);

      fHistDate:=formatdatetime('yyyymmdd',now-(data_module.fiHistoricalForecast.AsInteger*7));

      if ScanPartnumber then
      begin
        for i:=0 to High(fEntries) do
        begin
          if (not fEntries[i].Skip) then
          begin
            // Delete data that will be forecast this time and clear history, keep anything that isn't forecast this time
            DeleteBreakdown(fEntries[i].PartNumber);
          end;
        end;

        //add to forecast table
        UpdateForecast;

        // Update Usage in Size Master
        UpdateUsage;

        //
        //  Produce output files to send to suppliers
        //
        try
          with Data_Module.Inv_DataSet do
          begin
            //  Get week number only process files for the next week out

            Data_Module.Inv_DataSet.Close;
            Filter:='';
            Filtered:=FALSE;
            Close;
            CommandType := CmdStoredProc;
            CommandText := 'dbo.SELECT_ForecastSupplier;1';
            Parameters.Clear;
            Parameters.AddParameter.Name := '@WeekDate';
            Parameters.ParamValues['@WeekDate'] := formatdatetime('yyyymmdd',now);
            Open;

            //
            //  Change to select file output type based on Supplier selection
            //
            //
            // Init both file types;
            excel:=Unassigned;
            fopen:=False;

            while not eof do
            begin
              if fieldbyname('VC_SUPPLIER_CODE').AsString <> lastsupplier then
              begin
                try
                  if not VarIsEmpty(excel) then
                  begin
                    //  Create file in directory specified for each supplier
                    //  Use supplier name+WeekDate for filename
                    try
                      if FileExists(Data_Module.Inv_StoredProc.FieldbyName('Directory').AsString+'\'+ANSIReplaceStr(Data_Module.Inv_StoredProc.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.Inv_StoredProc.fieldbyname('Supplier Code').AsString+'-'+'Forecast') then
                        DeleteFile(Data_Module.Inv_StoredProc.FieldbyName('Directory').AsString+'\'+ANSIReplaceStr(Data_Module.Inv_StoredProc.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.Inv_StoredProc.fieldbyname('Supplier Code').AsString+'-'+'Forecast');
                      excel.ActiveWorkbook.SaveAs(Data_Module.Inv_StoredProc.FieldbyName('Directory').AsString+'/'+ANSIReplaceStr(Data_Module.Inv_StoredProc.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.Inv_StoredProc.fieldbyname('Supplier Code').AsString+'-'+'Forecast');
                    except
                      on e:exception do
                      begin
                        Data_Module.LogActLog('ERROR','Failed on delete and save excel files, '+e.message+', for supplier('+Data_Module.Inv_StoredProc.FieldbyName('Directory').AsString+'\'+ANSIReplaceStr(Data_Module.Inv_StoredProc.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.Inv_StoredProc.fieldbyname('Supplier Code').AsString+'-'+'Forecast'+')');
                        Hist.Append('Failed on delete and save excel files, '+e.message+', for supplier('+Data_Module.Inv_StoredProc.FieldbyName('Directory').AsString+'\'+ANSIReplaceStr(Data_Module.Inv_StoredProc.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.Inv_StoredProc.fieldbyname('Supplier Code').AsString+'-'+'Forecast'+')');
                      end;
                    end;
                    excel.Workbooks.Close;
                    excel.Quit;
                    excel:=Unassigned;
                  end;
                  if fopen then
                  begin
                    CloseFile(tcf);
                    if Data_Module.fiLocalFTP.AsBoolean then
                    begin
                      CloseFile(taf)
                    end;
                    fOpen:=FALSE;
                  end;
                except
                  on e:exception do
                  begin
                    Data_Module.LogActLog('ERROR','Failed on close output files, '+e.message+', for supplier('+fieldbyname('VC_SUPPLIER_CODE').AsString+')');
                    Hist.Append('Failed on close output files, '+e.message+', for supplier('+fieldbyname('VC_SUPPLIER_CODE').AsString+')');

                    if not VarIsEmpty(excel) then
                    begin
                      excel.Workbooks.Close;
                      excel.Quit;
                      excel:=Unassigned;
                    end;
                  end;
                end;

                lastsupplier:=fieldbyname('VC_SUPPLIER_CODE').AsString;
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
                  Open;

                  if FieldByName('Output File Type').AsString = 'TEXT' then
                    fFileKind := fText
                  else if FieldByName('Output File Type').AsString = 'EXCEL' then
                    fFileKind := fExcel
                  else
                    fFileKind := fBoth;

                  sendsite:=fieldbyname('Site Number in Order').AsBoolean;
                End;    //With

                if (fFileKind = fExcel) or (fFileKind = fBoth) then
                begin
                  excel := createOleObject('Excel.Application');
                  excel.visible := False;
                  excel.DisplayAlerts := False;

                  excel.workbooks.open(Data_Module.TemplateDir+'ForecastTemplate.xls');
                  mysheet := excel.workSheets[1];
                  Hist.Append('Create excel file for supplier, '+lastsupplier);
                  Data_Module.LogActLog('FORECAST','Create excel file for supplier, '+lastsupplier);

                  i:=2;
                end;
                if (fFileKind = fText) or (fFileKind = fBoth) then
                begin
                  fn:=Data_Module.Inv_StoredProc.FieldbyName('Directory').AsString+'\'+ANSIReplaceStr(Data_Module.Inv_StoredProc.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.Inv_StoredProc.fieldbyname('Supplier Code').AsString+'.frc';
                  AssignFile(tcf,Data_Module.Inv_StoredProc.FieldbyName('Directory').AsString+'\'+ANSIReplaceStr(Data_Module.Inv_StoredProc.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.Inv_StoredProc.fieldbyname('Supplier Code').AsString+'.frc');
                  Data_Module.LogActLog('FORECAST','Create text file :'+Data_Module.Inv_StoredProc.FieldbyName('Directory').AsString+'\'+ANSIReplaceStr(Data_Module.Inv_StoredProc.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.Inv_StoredProc.fieldbyname('Supplier Code').AsString+'.frc');
                  Rewrite(tcf);
                  if Data_Module.fiLocalFTP.AsBoolean then
                  begin
                    AssignFile(taf,Data_Module.Inv_StoredProc.FieldbyName('Directory').AsString+'\Archive\'+ANSIReplaceStr(Data_Module.Inv_StoredProc.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.Inv_StoredProc.fieldbyname('Supplier Code').AsString+formatdatetime('yyyymmdd',now)+'.frc');
                    Data_Module.LogActLog('FORECAST','Create archive text file :'+Data_Module.Inv_StoredProc.FieldbyName('Directory').AsString+'\Archive\'+ANSIReplaceStr(Data_Module.Inv_StoredProc.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.Inv_StoredProc.fieldbyname('Supplier Code').AsString+formatdatetime('yyyymmdd',now)+'.frc');
                    Rewrite(taf);
                  end;
                  fopen:=true;
                  Hist.Append('Create text file for supplier, '+lastsupplier);
                  Data_Module.LogActLog('FORECAST','Create text file for supplier, '+lastsupplier);
                end;
              end;

              if (fFileKind = fExcel) or (fFileKind = fBoth) then
              begin
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
              end;
              if (fFileKind = fText) or (fFileKind = fBoth) then
              begin
                tcl:='';

                if sendsite then                              // Mod to add site supplier
                begin
                  tcl:=Data_Module.fiSupplierCode.AsString;
                  tcl:=tcl+fieldbyname('VC_SUPPLIER_CODE').AsString;
                end
                else
                begin
                  tcl:=fieldbyname('VC_SUPPLIER_CODE').AsString;
                end;

                tcl:=tcl+fieldbyname('VC_PART_NUMBER').AsString;
                tcl:=tcl+fieldbyname('VC_WEEK_DATE').AsString;
                tcl:=tcl+Format('%.2d',[fieldbyname('IN_WEEK_NUMBER').AsInteger]);
                tcl:=tcl+Format('%.5d',[fieldbyname('IN_QTY1').AsInteger]);
                tcl:=tcl+Format('%.5d',[fieldbyname('IN_QTY2').AsInteger]);
                tcl:=tcl+Format('%.5d',[fieldbyname('IN_QTY3').AsInteger]);
                tcl:=tcl+Format('%.5d',[fieldbyname('IN_QTY4').AsInteger]);
                tcl:=tcl+Format('%.5d',[fieldbyname('IN_QTY5').AsInteger]);
                tcl:=tcl+Format('%.5d',[fieldbyname('IN_QTY6').AsInteger]);
                tcl:=tcl+Format('%.5d',[fieldbyname('IN_QTY7').AsInteger]);

                Writeln(tcf,tcl);
                if Data_Module.fiLocalFTP.AsBoolean then
                begin
                  Writeln(taf,tcl);
                end;
              end;
              next;
            end;

            if not VarIsEmpty(excel) then
            begin
              //  Create file in directory specified for each supplier
              //  Use supplier name+WeekDate for filename
              if FileExists(Data_Module.Inv_StoredProc.FieldbyName('Directory').AsString+'/'+ANSIReplaceStr(Data_Module.Inv_StoredProc.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.Inv_StoredProc.fieldbyname('Supplier Code').AsString+'-'+'Forecast') then
                DeleteFile(Data_Module.Inv_StoredProc.FieldbyName('Directory').AsString+'/'+ANSIReplaceStr(Data_Module.Inv_StoredProc.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.Inv_StoredProc.fieldbyname('Supplier Code').AsString+'-'+'Forecast');
              excel.ActiveWorkbook.SaveAs(Data_Module.Inv_StoredProc.FieldbyName('Directory').AsString+'/'+ANSIReplaceStr(Data_Module.Inv_StoredProc.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.Inv_StoredProc.fieldbyname('Supplier Code').AsString+'-'+'Forecast');

              if Data_Module.fiLocalFTP.AsBoolean then
              begin
                if FileExists(Data_Module.Inv_StoredProc.FieldbyName('Directory').AsString+'\Archive\'+ANSIReplaceStr(Data_Module.Inv_StoredProc.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.Inv_StoredProc.fieldbyname('Supplier Code').AsString+'-'+'Forecast') then
                  DeleteFile(Data_Module.Inv_StoredProc.FieldbyName('Directory').AsString+'\Archive\'+ANSIReplaceStr(Data_Module.Inv_StoredProc.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.Inv_StoredProc.fieldbyname('Supplier Code').AsString+'-'+'Forecast');
                //Archive
                excel.ActiveWorkbook.SaveAs(Data_Module.Inv_StoredProc.FieldbyName('Directory').AsString+'\Archive\'+ANSIReplaceStr(Data_Module.Inv_StoredProc.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.Inv_StoredProc.fieldbyname('Supplier Code').AsString+'-'+'Forecast');
              end;

              excel.Workbooks.Close;
              excel.Quit;
              excel:=Unassigned;
            end
            else if fopen then
            begin
              CloseFile(tcf);
              if Data_Module.fiLocalFTP.AsBoolean then
              begin
                CloseFile(taf);
              end;
            end;
            //
            // End new file selection code
            //
            //
          end;
        except
          on e:exception do
          begin
            Data_Module.LogActLog('ERROR','Unable to create forecast output files, '+e.message+', for supplier('+lastsupplier+')');
            Hist.Append('Unable to create forecast output files, '+e.message+', for supplier('+lastsupplier+')');
            result:=false;

            if not VarIsEmpty(excel) then
            begin
              excel.Workbooks.Close;
              excel.Quit;
              excel:=Unassigned;
            end;
          end;
        end;


        ForecastReport := TForecastCAMEXReport.Create();
        if not ForecastReport.Execute then
        begin
          Data_Module.LogActLog('FORECAST','Failed on Camex forecast report');
          Hist.Append('Camex forecast excel file create failed');
        end
        else
          Hist.Append('Camex forecast excel create complete');


        Data_Module.LogActLog('FORECAST','Forecast processing complete');
        Hist.Append('Forecast processing complete, Press OK to continue');
      end //DEBUG
      else
        result:=FALSE; //DEBUG
    except
      on e:exception do
      begin
        Data_Module.LogActLog('ERROR','Unable to load forecast, '+e.message);
        Hist.Append('Unable to load forecast, '+e.Message);
        result:=false;
      end;
    end;
  finally
    CloseFile(fcf);
  end;
  OKButton.Visible:=True;
  while not fclosed do
  begin
    application.ProcessMessages;
    sleep(500);
  end;
  //put data into forecast table
end;


function TForecastBreakdown_Form.ScanLine(line:string;count:integer):boolean;
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

      fEntries[count].Weeks[8].WeekNumber:=StrToInt(copy(line,FORECASTW8_OFF,WEEKNUMBER_SIZE));
      fEntries[count].Weeks[8].WeekDate:=copy(line,FORECASTW8_OFF+WEEKNUMBER_SIZE,WEEKDATE_SIZE);
      if copy(line,FORECASTW8_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE) <> '     -' then
        fEntries[count].Weeks[8].WeekCount:=StrToInt(copy(line,FORECASTW8_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE))
      else
        fEntries[count].Weeks[8].WeekCount:=0;

      fEntries[count].Weeks[9].WeekNumber:=StrToInt(copy(line,FORECASTW9_OFF,WEEKNUMBER_SIZE));
      fEntries[count].Weeks[9].WeekDate:=copy(line,FORECASTW9_OFF+WEEKNUMBER_SIZE,WEEKDATE_SIZE);
      if copy(line,FORECASTW9_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE) <> '     -' then
        fEntries[count].Weeks[9].WeekCount:=StrToInt(copy(line,FORECASTW9_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE))
      else
        fEntries[count].Weeks[9].WeekCount:=0;

      fEntries[count].Weeks[10].WeekNumber:=StrToInt(copy(line,FORECASTW10_OFF,WEEKNUMBER_SIZE));
      fEntries[count].Weeks[10].WeekDate:=copy(line,FORECASTW10_OFF+WEEKNUMBER_SIZE,WEEKDATE_SIZE);
      if copy(line,FORECASTW10_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE) <> '     -' then
        fEntries[count].Weeks[10].WeekCount:=StrToInt(copy(line,FORECASTW10_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE))
      else
        fEntries[count].Weeks[10].WeekCount:=0;

      fEntries[count].Weeks[11].WeekNumber:=StrToInt(copy(line,FORECASTW11_OFF,WEEKNUMBER_SIZE));
      fEntries[count].Weeks[11].WeekDate:=copy(line,FORECASTW11_OFF+WEEKNUMBER_SIZE,WEEKDATE_SIZE);
      if copy(line,FORECASTW11_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE) <> '     -' then
        fEntries[count].Weeks[11].WeekCount:=StrToInt(copy(line,FORECASTW11_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE))
      else
        fEntries[count].Weeks[11].WeekCount:=0;

      fEntries[count].Weeks[12].WeekNumber:=StrToInt(copy(line,FORECASTW12_OFF,WEEKNUMBER_SIZE));
      fEntries[count].Weeks[12].WeekDate:=copy(line,FORECASTW12_OFF+WEEKNUMBER_SIZE,WEEKDATE_SIZE);
      if copy(line,FORECASTW12_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE) <> '     -' then
        fEntries[count].Weeks[12].WeekCount:=StrToInt(copy(line,FORECASTW12_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE))
      else
        fEntries[count].Weeks[12].WeekCount:=0;

      fEntries[count].Weeks[13].WeekNumber:=StrToInt(copy(line,FORECASTW13_OFF,WEEKNUMBER_SIZE));
      fEntries[count].Weeks[13].WeekDate:=copy(line,FORECASTW13_OFF+WEEKNUMBER_SIZE,WEEKDATE_SIZE);
      if copy(line,FORECASTW13_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE) <> '     -' then
        fEntries[count].Weeks[13].WeekCount:=StrToInt(copy(line,FORECASTW13_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE))
      else
        fEntries[count].Weeks[13].WeekCount:=0;

      if Data_Module.fiAssemblerName.AsString='WQS' then
      begin
        fEntries[count].Weeks[14].WeekNumber:=StrToInt(copy(line,FORECASTW14_OFF,WEEKNUMBER_SIZE));
        fEntries[count].Weeks[14].WeekDate:=copy(line,FORECASTW14_OFF+WEEKNUMBER_SIZE,WEEKDATE_SIZE);
        if copy(line,FORECASTW14_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE) <> '     -' then
          fEntries[count].Weeks[14].WeekCount:=StrToInt(copy(line,FORECASTW14_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE))
        else
          fEntries[count].Weeks[14].WeekCount:=0;
      end;

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


function TForecastBreakdown_Form.ScanPartnumber:boolean;
var
  i,x,z,y,skip:integer;
  excel,mysheet:variant;
  dbmissing:TStringList;
  foundx:boolean;
begin
  result:=true;
  skip:=0;
  dbmissing:=TSTringList.Create;

  Hist.Append('Scan part number list');
  Data_Module.LogActLog('FORECAST','Scan part number list');

  with Data_Module.Inv_DataSet do
  begin
    try
      for i:=0 to High(fEntries) do
      begin
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'dbo.SELECT_ForecastDetail;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@AssyCode';
        Parameters.ParamValues['@AssyCode'] := fEntries[i].PArtNumber;
        Open;
        if recordcount = 0 then //IsEmpty then//  not Locate('Assembly part number Code',fEntries[i].PArtNumber,[]) then
        begin
          // Partnumber does not exist, skip
          Hist.Append('Part Number not found in DB, '+fEntries[i].PartNumber);
          Data_Module.LogActLog('FORECAST','Part Number not found in DB, '+fEntries[i].PartNumber);
          fEntries[i].Skip:=True;
          INC(skip);
        end;
      end;
      Close;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.SELECT_ForecastDetail;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@AssyCode';
      Parameters.ParamValues['@AssyCode'] := '';
      Parameters.AddParameter.Name := '@ForecastNotZero';
      Parameters.ParamValues['@ForecastNotZero'] := 1;
      Parameters.AddParameter.Name := '@EffectiveMonth';
      Parameters.ParamValues['@EffectiveMonth'] := formatdatetime('yyyy/mm',now);
      Open;
      while not eof do
      begin
        foundx:=False;
        for i:=0 to High(fEntries) do
        begin
          if fEntries[i].Partnumber = fieldbyname('Assembly Part Number Code').AsString then
          begin
            foundx:=true;
            break;
          end;
        end;
        if not foundx then
        begin
          dbmissing.Add(fieldbyname('Assembly Part Number Code').AsString);
          Hist.Append('Part Number not found in Forecast, '+fieldbyname('Assembly Part Number Code').AsString);
          Data_Module.LogActLog('FORECAST','Part Number not found in Forecast, '+fieldbyname('Assembly Part Number Code').AsString);
        end;
        next;
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

  //
  //  Forecast report
  //
  try
    excel := createOleObject('Excel.Application');
    excel.visible := False;
    excel.DisplayAlerts := False;
    //excel.workbooks.add;
    excel.workbooks.open(Data_Module.TemplateDir+'ReportTemplate.xls');
    mysheet := excel.workSheets[1];

    mysheet.cells[1,1].value:='Forecast Part Numbers';
    z:=4;

    if Data_Module.fiAssemblerName.AsString='WQS' then
    begin
      for y:=1 to 14 do
      begin
        mysheet.Cells[z-1,Y+1].value := 'Week '+IntToStr(fEntries[1].Weeks[y].WeekNumber);
      end;
      for x:=0 to high(fEntries) do
      begin
        mysheet.Cells[z,1].value := fEntries[x].Partnumber;
        for y:=1 to 14 do
        begin
          mysheet.Cells[z,Y+1].value := fEntries[x].Weeks[y].WeekCount;
        end;
        INC(z);
      end;
    end
    else
    begin
      for y:=1 to 13 do
      begin
        mysheet.Cells[z-1,Y+1].value := 'Week '+IntToStr(fEntries[1].Weeks[y].WeekNumber);
      end;
      for x:=0 to high(fEntries) do
      begin
        mysheet.Cells[z,1].value := fEntries[x].Partnumber;
        for y:=1 to 13 do
        begin
          mysheet.Cells[z,Y+1].value := fEntries[x].Weeks[y].WeekCount;
        end;
        INC(z);
      end;
    end;

    excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\ForecastReport'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
    excel.Workbooks.Close;
    excel.Quit;
    excel:=Unassigned;
  except
    on e:exception do
    begin
      Showmessage('Cannot save excel report template('+Data_Module.fiReportsOutputDir.AsString+'\ForecastReport'+formatdatetime('yyyymmddhhmmss00',now)+'.xls'+'), '+e.Message);
      Data_Module.LogActLog('ERROR','Cannot save excel report template('+Data_Module.fiReportsOutputDir.AsString+'\ForecastReport'+formatdatetime('yyyymmddhhmmss00',now)+'.xls'+'), '+e.Message);
      //raise;
    end;
  end;

  //
  // record in database and not in forecast
  //
  try
    if dbmissing.Count > 0 then
    begin
      if messagedlg('There are '+IntToStr(dbMissing.count)+', in the database and not in the forecast. Continue processing?',
                      mtConfirmation, [mbYes, mbNo], 0) = mrNo then
      begin
        result:=False;
        exit;
      end;

      hist.append('Create skipped database report, '+IntToStr(skip)+' records');
      Data_Module.LogActLog('FORECAST','Create skipped database report, '+IntToStr(skip)+' records');
      // create error xls form
      excel := createOleObject('Excel.Application');
      excel.visible := False;
      excel.DisplayAlerts := False;
      excel.workbooks.open(Data_Module.TemplateDir+'ReportTemplate.xls');
      mysheet := excel.workSheets[1];
      mysheet.cells[1,1].value:='Forecast Part Numbers not in forecast';
      z:=4;
      for x:=0 to dbmissing.Count-1 do
      begin
        mysheet.Cells[z,1].value := dbmissing[x];
        INC(z);
      end;
      excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\ForecastDBError'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
      excel.Workbooks.Close;
      excel.Quit;
      excel:=Unassigned;
    end;
    dbmissing.Free;
  except
    on e:exception do
    begin
      Showmessage('Cannot save excel report template('+Data_Module.fiReportsOutputDir.AsString+'\ForecastReport'+formatdatetime('yyyymmddhhmmss00',now)+'.xls'+'), '+e.Message);
      Data_Module.LogActLog('ERROR','Cannot save excel report template('+Data_Module.fiReportsOutputDir.AsString+'\ForecastReport'+formatdatetime('yyyymmddhhmmss00',now)+'.xls'+'), '+e.Message);
    end;
  end;

  //
  // record in forecast and not in database
  //
  try
    if skip <> 0 then
    begin
      if messagedlg('There are '+IntToStr(skip)+', forecast records not in the database. Continue processing?',
                      mtConfirmation, [mbYes, mbNo], 0) = mrNo then
      begin
        result:=False;
        exit;
      end;

      hist.append('Create skipped records report, '+IntToStr(skip)+' records');
      Data_Module.LogActLog('FORECAST','Create skipped records report, '+IntToStr(skip)+' records');
      // create error xls form
      excel := createOleObject('Excel.Application');
      excel.visible := False;
      excel.DisplayAlerts := False;
      //excel.workbooks.add;
      excel.workbooks.open(Data_Module.TemplateDir+'ReportTemplate.xls');
      mysheet := excel.workSheets[1];
      mysheet.cells[1,1].value:='Forecast Part Numbers not in database';
      z:=4;
      for x:=0 to high(fEntries) do
      begin
        if fEntries[x].Skip then
        begin
          mysheet.Cells[z,1].value := fEntries[x].Partnumber;
          INC(z);
        end;
      end;
      excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\ForecastRecError'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
      excel.Workbooks.Close;
      excel.Quit;
      excel:=Unassigned;
    end;
  except
    on e:exception do
    begin
      Showmessage('Cannot save excel report template('+Data_Module.fiReportsOutputDir.AsString+'\ForecastReport'+formatdatetime('yyyymmddhhmmss00',now)+'.xls'+'), '+e.Message);
      Data_Module.LogActLog('ERROR','Cannot save excel report template('+Data_Module.fiReportsOutputDir.AsString+'\ForecastReport'+formatdatetime('yyyymmddhhmmss00',now)+'.xls'+'), '+e.Message);
      //raise;
    end;
  end;

  Hist.append(IntToStr(high(fEntries)-skip)+' forecast records to be added');
  Data_Module.LogActLog('FORECAST',IntToStr(high(fEntries)-skip)+' forecast records to be added');
end;

procedure TForecastBreakdown_Form.UpdateUsage;
var
  lastsize:string;
  usage:integer;
begin
  Hist.Append('Update usage DB');
  Data_Module.LogActLog('FORECAST','Update usage DB');
  application.ProcessMessages;
  try
    With Data_Module.INV_DataSet do
    Begin
      Close;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.SELECT_SizeUsage;1';
      Parameters.Clear;
      Open;

      lastsize:=fieldbyname('VC_SIZE_CODE').AsString;
      usage:=0;
      while not EOF do
      begin
        if (lastsize <> fieldbyname('VC_SIZE_CODE').AsString) then
        begin
          // Update last size
          with Data_Module.Inv_StoredProc do
          begin
            //Get ratio data
            if usage <> 0 then
            begin
              Close;
              ProcedureName := 'dbo.UPDATE_SizeUsage;1';
              Parameters.Clear;
              Parameters.AddParameter.Name := '@SizeCode';
              Parameters.ParamValues['@SizeCode'] := lastsize;
              Parameters.AddParameter.Name := '@Usage';
              Parameters.ParamValues['@Usage'] := usage;
              ExecProc;
            end;
          end;
          // Set next
          lastsize:=fieldbyname('VC_SIZE_CODE').AsString;
          usage:=0;
        end;

        // Get week forecast
        usage:=usage+HistoryForecast(fieldbyname('VC_PART_NUMBER').AsString);

        next;
        if eof then
        begin
          // do last
          with Data_Module.Inv_StoredProc do
          begin
            //Get ratio data
            if usage <> 0 then
            begin
              Close;
              ProcedureName := 'dbo.UPDATE_SizeUsage;1';
              Parameters.Clear;
              Parameters.AddParameter.Name := '@SizeCode';
              Parameters.ParamValues['@SizeCode'] := lastsize;
              Parameters.AddParameter.Name := '@Usage';
              Parameters.ParamValues['@Usage'] := usage;
              ExecProc;
            end;
          end;
        end;
      end;
    end;
  except
    on e:exception do
    begin
      Hist.append('Failed to update usage in INV_SIZE_MST,'+e.message);
      Data_Module.LogActLog('ERROR','FORECAST: Failed to update usage in INV_SIZE_MST,'+e.message);
    end;
  end;
  Hist.Append('Finished update usage DB');
  Data_Module.LogActLog('FORECAST','Finished update usage DB');
end;

function TForecastBreakdown_Form.HistoryForecast(partno:string):integer;
var
  x,z,y,total:integer;
begin
  //
  //  Change to 7 days instead of 30
  //
  x:=0;
  total:=0;
  try
    with data_module.INV_Forecast_DataSet do
    begin
      for z:=0 to Data_Module.fiUsageUpdateCompare.AsInteger do
      begin
        for y:=0 to 7 do
        begin
          Close;
          CommandType := CmdStoredProc;
          CommandText := 'dbo.SELECT_ForecastPartNumberWeek;1';
          Parameters.Clear;
          Parameters.AddParameter.Name := '@WeekNo';
          Parameters.ParamValues['@WeekNo'] := WeekOfTheYear(now+z);
          Parameters.AddParameter.Name := '@DayNo';
          Parameters.ParamValues['@DayNo'] := DayOfTheWeek(now+y);
          Parameters.AddParameter.Name := '@PartNo';
          Parameters.ParamValues['@PartNo'] := partno;
          Open;
          if (not FieldByName('Qty').IsNull) and (FieldByName('Qty').Value <> 0) then
          begin
            total:=total+FieldByName('Qty').Value;
            INC(x);
          end;
        end;
      end;
    end;
    if x<>0 then
      result:=total div x
    else
      result:=0;
  except
    on e:exception do
    begin
      Data_Module.LogActLog('ERROR','Unable to get usage forecast, '+e.Message);
      ShowMessage('Unable to get usage forecast, '+e.Message);
      raise;
    end;
  end;
end;


procedure TForecastBreakdown_Form.UpdateForecast;
var
  i,j,count,tirecount,wheelcount,weekcount,wheelratio,tireratio,forecastratio:integer;
  wm,tm:string;
  bd:boolean;
begin
  Hist.Append('Update forecast DB');
  Data_Module.LogActLog('FORECAST','Update forecast DB');
  if Data_Module.fiAssemblerName.AsString = 'WQS' then
    count:=14
  else
    count:=13;
  for i:=0 to High(fEntries) do
  begin
    if (not fEntries[i].Skip) then
    begin
      //insert or update record
      try
        for j:=1 to count do
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
              Parameters.ParamValues['@AssyCode'] := fEntries[i].PartNumber;
              Parameters.AddParameter.Name := '@ForecastNotZero';
              Parameters.ParamValues['@ForecastNotZero'] := 1;
              Open;

              //
              //  Add adjustment for ratio here
              //
              //      Date rule:
              //                  Record with a blank effective date is the default
              //                  Record the contains an effective month overrides if that is the month of the current data
              //                  if No blank record and no effective month for this time period, set value to zero for this time
              //
              tirecount:=0; // init everything
              wheelcount:=0;
              tireratio:=0;
              wheelratio:=0;
              forecastratio:=0;
              weekcount:=fEntries[i].Weeks[j].WeekCount; // for reading purposes
              bd:=FALSE; // flag for warning

              while not eof do
              begin
                if (FieldByName('Active Date').AsString = '') or (FieldByName('Active Date').AsString = ' ') then
                begin
                  // calculate default
                  wheelratio:=FieldByName('Wheel Ratio').AsInteger;
                  tireratio:=FieldByName('Tire Ratio').AsInteger;
                  forecastratio:=FieldByName('Forecast Ratio').AsInteger;
                  bd:=TRUE;
                end
                else
                begin
                  tm:=copy(FieldByName('Active Date').AsString,3,2)+copy(FieldByName('Active Date').AsString,6,2);
                  wm:=copy(fEntries[i].Weeks[j].WeekDate,1,4);
                  if tm = wm then
                  begin
                    // Month record match
                    wheelratio:=FieldByName('Wheel Ratio').AsInteger;
                    tireratio:=FieldByName('Tire Ratio').AsInteger;
                    forecastratio:=FieldByName('Forecast Ratio').AsInteger;
                    bd:=TRUE;
                    break;
                  end;
                end;
                next;
              end;

              if not bd then
              begin
                Data_Module.LogActLog('ERROR','No breakdown for part number('+fEntries[i].PartNumber+') on week('+fEntries[i].Weeks[j].WeekDate+') with count('+IntToStr(weekcount)+')');
                //ShowMessage('No breakdown for part number('+fEntries[i].PartNumber+') on week('+fEntries[i].Weeks[j].WeekDate+') with count('+IntToStr(weekcount)+'), count will be ignored.');
                Hist.Append('No breakdown for part number('+fEntries[i].PartNumber+') on week('+fEntries[i].Weeks[j].WeekDate+') with count('+IntToStr(weekcount)+'), count will be ignored.');
              end;

              // If any ratio is zero the total is zero
              if (forecastratio <> 0) and
                 (tireratio <> 0) and
                 (wheelratio <> 0) then
              begin
                // Changes to reflect ratios over 100 percent
                //
                // Added to support non-production values shifted to
                // other suppliers
                //

                if (tireratio <> 100) and (wheelratio <> 100) then
                begin
                    tirecount:=((((WeekCount div 2) * forecastratio) div 100) * tireratio) div 100;
                    wheelcount:=((((WeekCount div 2) * forecastratio) div 100) * wheelratio) div 100;
                end
                else if tireratio <> 100 then
                begin
                    tirecount:=(((WeekCount * forecastratio) div 100) * tireratio) div 100;
                    wheelcount:=(((WeekCount div 2) * forecastratio) div 100);
                end
                else if wheelratio <> 100 then
                begin
                    tirecount:=(((WeekCount div 2) * forecastratio) div 100);
                    wheelcount:=(((WeekCount * forecastratio) div 100) * wheelratio) div 100;
                end
                else //everything 100
                begin
                    tirecount:=((WeekCount * forecastratio) div 100);
                    wheelcount:=((WeekCount * forecastratio) div 100);
                end
              end;


              // scan through the part numbers for this assembly code and assign
              if length(FieldByName('Tire Part Number Code').AsString) = 12 then
              begin

                DoPartNumberForecast( FieldByName('Tire Part Number Code').AsString,
                                      fEntries[i].Weeks[j].WeekDate,
                                      tirecount,//fEntries[i].Weeks[j].WeekCount,
                                      fEntries[i].Weeks[j].WeekNumber);

              end;

              if length(FieldByName('Wheel Part Number Code').AsString) = 12 then
              begin
                DoPartNumberForecast( FieldByName('Wheel Part Number Code').AsString,
                                      fEntries[i].Weeks[j].WeekDate,
                                      wheelcount,//fEntries[i].Weeks[j].WeekCount,
                                      fEntries[i].Weeks[j].WeekNumber);
              end;

              if length(FieldByName('Valve Part Number').AsString) = 12 then
              begin
                DoPartNumberForecast( FieldByName('Valve Part Number').AsString,
                                      fEntries[i].Weeks[j].WeekDate,
                                      wheelcount,//fEntries[i].Weeks[j].WeekCount,
                                      fEntries[i].Weeks[j].WeekNumber);
              end;

              if length(FieldByName('Film Part Number').AsString) = 12 then
              begin
                DoPartNumberForecast( FieldByName('Film Part Number').AsString,
                                      fEntries[i].Weeks[j].WeekDate,
                                      wheelcount,//fEntries[i].Weeks[j].WeekCount,
                                      fEntries[i].Weeks[j].WeekNumber);
              end;


              if length(FieldByName('Label Part Number').AsString) = 12 then
              begin
                DoPartNumberForecast( FieldByName('Label Part Number').AsString,
                                      fEntries[i].Weeks[j].WeekDate,
                                      wheelcount,//fEntries[i].Weeks[j].WeekCount,
                                      fEntries[i].Weeks[j].WeekNumber);
              end;

              if length(FieldByName('Misc1 Part Number').AsString) = 12 then
              begin
                DoPartNumberForecast( FieldByName('Misc1 Part Number').AsString,
                                      fEntries[i].Weeks[j].WeekDate,
                                      wheelcount,//fEntries[i].Weeks[j].WeekCount,
                                      fEntries[i].Weeks[j].WeekNumber);
              end;

              if length(FieldByName('Misc2 Part Number').AsString) = 12 then
              begin
                DoPartNumberForecast( FieldByName('Misc2 Part Number').AsString,
                                      fEntries[i].Weeks[j].WeekDate,
                                      wheelcount,//fEntries[i].Weeks[j].WeekCount,
                                      fEntries[i].Weeks[j].WeekNumber);
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
  Hist.Append('Finished update forecast');
  Data_Module.LogActLog('FORECAST','Finished update forecast');
end;

procedure TForecastBreakdown_Form.DoPartNumberForecast(PN,WeekDate:string;FCCount,Weeknumber:integer);
var
  workday: array[1..7] of boolean;
  dayforecast: array[1..7] of integer;
  line,supplier,size:string;
  days,ratiocount,leftover,i, checkweeknumber:integer;
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
      Parameters.AddParameter.Name := '@PartNum';
      Parameters.ParamValues['@PartNum'] := PN;
      Open;
      if FieldByName('Line Name').AsString <> '' then
        line:=FieldByName('Line Name').AsString
      else
        line := 'ALL LINES';

      supplier:=FieldByName('Supplier Code').AsString;
      size:=FieldByName('Size Code').AsString;
      Close;

      //
      //  Modify the week number based on using the first production day
      //  week number value
      //
      //
      checkweeknumber:=WeekNumber;

      with data_module.Inv_StoredProc do
      begin
        if data_module.fiUseFirstProductionDay.AsBoolean then
        begin
          Close;
          ProcedureName := 'dbo.SELECT_FirstProductionDay;1';
          Parameters.Clear;
          Parameters.AddParameter.Name := '@ProdYear';
          Parameters.ParamValues['@ProdYear'] := '20'+copy(weekdate,1,2);// formatdatetime('yyyy',EventDate_NUMMIBmDateEdit.date);
          Open;

          if FieldByName('First Week Number').AsInteger <> 1 then
          begin
            // Go in the opposite direction to get to julian date if first week is greater than production week
            WeekNumber := WeekNumber + FieldByName('First Week Number').AsInteger - 1;
          end;
        end;
      end;


      //
      //  Find any changes to a normal weekly schedule
      //
      with Data_Module.ALC_DataSet do
      begin
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'dbo.AD_GetSpecialDateWeek';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@Week';
        Parameters.ParamValues['@Week'] := WeekNumber;
        Parameters.AddParameter.Name := '@Line';
        Parameters.ParamValues['@Line'] := Line;
        Open;
        if recordCount > 0 then
        begin
          while not eof do
          begin
            if (trim(fieldByName('Date Status Abrv').AsString) = 'H')  or (trim(fieldByName('Date Status Abrv').AsString) = 'X') then
            begin
              workday[FieldByName('Day Number').AsInteger]:=False;
              DEC(days);
            end
            else
            begin
              workday[FieldByName('Day Number').AsInteger]:=True;
              INC(days);
            end;
            next;
          end;
        end;
        Close;
      end;

      if days > 0 then
      begin
        ratiocount:=FCCount div days;
        leftover:= FCCount mod days;
      end
      else
      begin
        ratiocount := 0;
        leftover:=0;
      end;
      for i:=1 to 7 do
      begin
        if workday[i] then
        begin
          dayforecast[i]:=ratiocount+leftover;
          // add any extra to the first day and then reset
          leftover:=0;
        end
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
      Parameters.ParamValues['@WeekNumber'] := checkweeknumber;//WeekNumber;
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
      Data_Module.LogActLog('ERROR','Failure on partnumber forecast update/insert PN('+PN+') '+e.Message);
      Hist.Append('Failed on Part Number Update PN('+PN+'), '+e.Message);
      Raise;
    end
  end;
end;

procedure TForecastBreakdown_Form.OKButtonClick(Sender: TObject);
begin
  fclosed:=true;
end;

end.
