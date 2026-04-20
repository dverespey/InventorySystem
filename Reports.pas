//****************************************************************
//
//       Sequential Supplier Program
//
//       Copyright (c) 2002-2006 Failproof Manufacturing Systems.
//
//****************************************************************
//
//
//  08/19/2002  David Verespey  Add support for FRS report
//  08/28/2002  David Verespey  Fix data/sequence selection errors
//  09/05/2002  David Verespey  Change controls for date select
//  09/26/2002  David Verespey  Fix Rollover bug
//  10/07/2002  David Verespey  Changes for FPC support
//  11/11/2002  David Verespey  Add resort for truck FRS
//  02/24/2006  David Verespey  Dice and slice for new Admin program
//
unit Reports;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, DBGrids, StdCtrls, Buttons, ExtCtrls, DBCtrls, Printers, Db,
  DBTables, ComCtrls, IniFiles, BmDtEdit,
  NUMMIBmDateEdit, ADODB, ComObj, Excel2000,Variants;

type
  TReportType=(rtWQSFRS,rtWQSSHIP,rtINJEXGB,rtINJEXDP,rtEDI810,rtEDI856);
const
  sReportType:array[tReportType] of string
              = ('WQS_FRS', 'WQS_SHIP', 'INJEX Glovebox', 'INJEX Door Panel', 'EDI 810 INVOICE', 'EDI 856 ASN');
type
  TReportsForm = class(TForm)
    SelectPanel: TPanel;
    FirstPanel: TPanel;
    LastPanel: TPanel;
    End_Seq_Edit: TEdit;
    Selected_DataSource: TDataSource;
    Label6: TLabel;
    Printer_Combo: TComboBox;
    Label7: TLabel;
    ButtonPanel: TPanel;
    Print_Button: TBitBtn;
    Close_Button: TBitBtn;
    StatusBar: TStatusBar;
    Label8: TLabel;
    Report_Query: TADOQuery;
    Selected_Query: TADOQuery;
    UpdateReportQuery: TADOQuery;
    UpdateReportQuerywf: TADOQuery;
    Line_Combo: TComboBox;
    Label2: TLabel;
    Report_Edit: TEdit;
    Line_DataSet: TADODataSet;
    GetLastPrint_DataSet: TADODataSet;
    Label1: TLabel;
    StartDate: TDateTimePicker;
    StartTime: TDateTimePicker;
    Label3: TLabel;
    EndDate: TDateTimePicker;
    EndTime: TDateTimePicker;
    Report_DataSet: TADODataSet;
    UpdateReportCommand: TADOCommand;
    StartBox: TComboBox;
    EndBox: TComboBox;
    FRSPanel: TPanel;
    Label9: TLabel;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    FRSBox: TComboBox;
    Label4: TLabel;
    Start_Seq_Edit: TEdit;
    SelectTypeBox: TComboBox;
    SelectTypeLabel: TLabel;
    FRSDataSet: TADODataSet;
    GetDataPanel: TPanel;
    Panel2: TPanel;
    procedure Close_ButtonClick(Sender: TObject);
    procedure Print_ButtonClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
    procedure Start_Seq_EditChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure StartBoxChange(Sender: TObject);
    procedure SelectTypeBoxChange(Sender: TObject);
    procedure Start_Seq_EditExit(Sender: TObject);
    procedure Line_ComboChange(Sender: TObject);
  private
    FCollector: String;
    FLocalAlias: String;
    FOffset: Integer;
    FSystemDisplay: String;
    fLineName:string;
    fConnection:TADOConnection;
    fReportType:TReportType;
    fInitSeq:Boolean;
    //procedure KanBan_Display(Act_Start, Act_End: Int64);
    //procedure Manifest_Display(Act_Start, Act_End: Int64);
    procedure Print_INJEXGB;
    procedure Print_INJEXDP;
    procedure Print_Assembly_Summary;
    procedure Print_FRS_Summary;  //DMV  8/19
    procedure Print_EDI810;
    procedure Print_EDI856;
    //procedure Print_Delivery_Summary;
    //procedure Print_Shipping_Summary;
    procedure Set_Connection(Value:TADOConnection);
    procedure Set_Formats(Value: TStringList);
  public
    constructor Create(AOwner: TComponent; ReportType:TReportType); reintroduce;
    property Connection: TADOConnection read fConnection write Set_Connection;
    property Collector: String write FCollector;
    property Formats: TstringList write Set_Formats;
    property LocalAlias: String write FLocalAlias;
    property Offset: Integer write FOffset;
    property SystemDisplay: String write FSystemDisplay;
    property LineName: string read fLineName write fLineName;
  end;

var
  ReportsForm: TReportsForm;

implementation

{$R *.DFM}

uses Shift_Summary, FRS_Summary, Data, EDI810Object, EDI856Object,
  DailySummay_Summary;

const
  FULL_HEIGHT_C = 495;
  HALF_HEIGHT_C = 226;

constructor TReportsForm.Create(AOwner: TComponent; ReportType:TReportType);
var
  i:integer;
begin
  inherited Create(Aowner);
  try
    Line_Dataset.Open;
    Line_Combo.Items.Clear;
    while not Line_Dataset.Eof do
    begin
      Line_Combo.Items.Add(Line_Dataset.FieldByName('LineName').AsString);
      Line_Dataset.Next;
    end;

    Line_Combo.ItemIndex:=0;


    With Printer Do
      for i:=0 to Printers.Count - 1 do
        Printer_Combo.Items.Add(Printers[i]);

    fInitSeq:=TRUE;
    Enddate.Date        := Now;
    End_Seq_Edit.Text   := '';
    startdate.Date      := Now;
    Start_Seq_Edit.Text := '';
    fInitSeq:=FALSE;


    StatusBar.Font.Size := 9;
    StatusBar.Font.Style:= [fsBold];

    fReportType:=ReportType;
    Report_Edit.Text:=sReportType[ReportType];


  except
    on e:exception do
    begin
      ShowMessage('Error on form create, '+e.Message);
    end;
  end;
end;

procedure TReportsForm.Print_Assembly_Summary;
var
  Shift_Summary: TShiftSummary;
begin
  try
    with Report_DataSet do
    begin
      GetDataPanel.Visible:=TRUE;
      application.processmessages;
      close;
      CommandType:=cmdStoredProc;
      CommandText:='AD_WQSSUMMARY';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@begindate';
      Parameters.ParamValues['@begindate'] :=  StartBox.Text;
      Parameters.AddParameter.Name := '@enddate';
      Parameters.ParamValues['@enddate'] := EndBox.Text;
      Parameters.AddParameter.Name := '@Start';
      Parameters.ParamValues['@Start'] := StrToInt(Start_Seq_Edit.Text);
      Parameters.AddParameter.Name := '@Last';
      Parameters.ParamValues['@Last'] := StrToInt(End_Seq_Edit.Text);
      Parameters.AddParameter.Name:='@LineName';
      Parameters.ParamValues['@LineName']:=Line_Combo.Text;
      Open;
      GetDataPanel.Visible:=FALSE;
      if recordcount = 0 then
      begin
        ShowMessage('No records');
        exit;
      end;
    end;
  except
    on e:exception do
    begin
      GetDataPanel.Visible:=FALSE;
      ShowMessage('Unable to get FRS data, '+e.Message);
      exit;
    end;
  end;

  Shift_Summary := TShiftSummary.Create(Self);
  With Shift_Summary Do
    Try
      DataSet       := Report_DataSet;
      DateEnd       := StartBox.Text;
      DateStart     := EndBox.Text;
      SequenceEnd   := End_Seq_Edit.Text;
      SequenceStart := Start_Seq_Edit.Text;
      LineName      := Line_Combo.Text;
      Site          := '';//AdminDataModule.SiteDateset.fieldByName('SiteAbbr').AsString;
      Preview;
    Finally
      Close;
      Screen.Cursor := crDefault;
    End;
end;

procedure TReportsForm.Print_EDI856;
var
  EDI856:T856EDI;
  i,j:integer;
  fcf:TextFile;
  line,templine:string;
begin
  EDI856:=T856EDI.Create;

  EDI856.LineName:=Line_Combo.Text;
  EDI856.StartDate:=StartDate.Date;
  EDI856.EndDate:=EndDate.Date;
  EDI856.StartSeq:=StrToInt(Start_Seq_Edit.Text);
  EDI856.EndSeq:=StrToInt(End_Seq_Edit.Text);

  if EDI856.Execute then
  begin
    AssignFile(fcf, AdminDataModule.fiReportDir.AsString+'\EDI856'+formatdatetime('yyyymmddhhmm',now)+'.txt');
    Rewrite(fcf);

    // Loop through report and save
    for i:=0 to EDI856.EDIRecord.Count-1 do
    begin
      if length(EDI856.EDIRecord[i]) > 80 then
      begin
        // break it up
        templine:=EDI856.EDIRecord[i];
        while length(templine) > 80 do
        begin
          line:=copy(templine,1,80);
          Writeln(fcf,line);
          templine:=copy(templine,81,length(templine));
        end;
        line:=templine;
        if length(line) <> 80 then
        begin
          for j:=length(line)+1 to 80 do
          begin
            line:=line+' ';
          end;
        end;
        Writeln(fcf,line);
      end
      else
      begin
        // write it out
        line:=EDI856.EDIRecord[i];
        if length(line)+1 <> 80 then
        begin
          for j:=length(line) to 80 do
          begin
            line:=line+' ';
          end;
        end;
        Writeln(fcf,line);
      end;
    end;
    CloseFile(fcf);
    ShowMessage('EDI 856 complete');
    try
      // Log report and update EIN
      UpdateReportCommand.CommandType:=cmdStoredProc;
      UpdateReportCommand.CommandText:='AD_UpdateEIN';
      UpdateReportCommand.Execute;
    except
      on e:exception do
      begin
        ShowMessage('Failed on EIN update after EDI810 create, '+e.Message);
      end;
    end;
  end
  else
    ShowMessage('Unable to create EDI856');

  EDI856.Free;
end;

procedure TReportsForm.Print_EDI810;
var
  EDI810:T810EDI;
  i,j:integer;
  fcf:TextFile;
  line,templine:string;
begin
  EDI810:=T810EDI.Create;

  EDI810.LineName:=Line_Combo.Text;
  EDI810.StartDate:=StartDate.Date;
  EDI810.EndDate:=EndDate.Date;
  EDI810.StartSeq:=StrToInt(Start_Seq_Edit.Text);
  EDI810.EndSeq:=StrToInt(End_Seq_Edit.Text);

  if EDI810.Execute then
  begin
    AssignFile(fcf, AdminDataModule.fiReportDir.AsString+'\EDI810'+formatdatetime('yyyymmddhhmm',now)+'.txt');
    Rewrite(fcf);

    // Loop through report and save
    for i:=0 to EDI810.EDIRecord.Count-1 do
    begin
      if length(EDI810.EDIRecord[i]) > 80 then
      begin
        // break it up
        templine:=EDI810.EDIRecord[i];
        while length(templine) > 80 do
        begin
          line:=copy(templine,1,80);
          Writeln(fcf,line);
          templine:=copy(templine,81,length(templine));
        end;
        line:=templine;
        if length(line) <> 80 then
        begin
          for j:=length(line)+1 to 80 do
          begin
            line:=line+' ';
          end;
        end;
        Writeln(fcf,line);
      end
      else
      begin
        // write it out
        line:=EDI810.EDIRecord[i];
        if length(line)+1 <> 80 then
        begin
          for j:=length(line) to 80 do
          begin
            line:=line+' ';
          end;
        end;
        Writeln(fcf,line);
      end;
    end;
    CloseFile(fcf);
    ShowMessage('EDI 810 complete');
    try
      UpdateReportCommand.CommandType:=cmdStoredProc;
      UpdateReportCommand.CommandText:='AD_UpdateEIN';
      UpdateReportCommand.Execute;
    except
      on e:exception do
      begin
        ShowMessage('Failed on EIN update after EDI810 create, '+e.Message);
      end;
    end;
  end
  else
    ShowMessage('Unable to create EDI810');

  EDI810.Free;
end;

procedure TReportsForm.Print_INJEXGB;
begin
  try
    with Report_DataSet do
    begin
      close;
      GetDataPanel.Visible:=TRUE;
      application.processmessages;

      CommandType:=cmdStoredProc;
      CommandText:='dbo.AD_INJEXSummary';
      Parameters.Clear;

      Parameters.AddParameter.Name:='@FirstFRSASN';
      Parameters.ParamByName('@FirstFRSASN').Direction:=pdOutput;
      Parameters.ParamValues['@FirstFRSASN']:=-1;
      Parameters.AddParameter.Name:='@LastFRSASN';
      Parameters.ParamByName('@LastFRSASN').Direction:=pdOutput;
      Parameters.ParamValues['@LastFRSASN']:=-1;

      Parameters.AddParameter.Name := '@RackVehicleDescription';
      Parameters.ParamValues['@RackVehicleDescription'] := 'Passenger Glove Box';
      Parameters.AddParameter.Name := '@DataItemDescription';
      Parameters.ParamValues['@DataItemDescription'] := 'GloveBox';
      Parameters.AddParameter.Name:='@LineName';
      Parameters.ParamValues['@LineName']:=Line_Combo.Text;

      if SelectTypeBox.ItemIndex = 0 then
      begin
        // FRS
        Parameters.AddParameter.Name := '@RangeType';
        Parameters.ParamValues['@RangeType'] := 'F';
        Parameters.AddParameter.Name:='@FRSNumber';
        Parameters.ParamValues['@FRSNumber']:=FRSBox.Text+'%';
        Parameters.AddParameter.Name:='@BeginDate';
        Parameters.ParamValues['@BeginDate']:=now;
        Parameters.AddParameter.Name:='@EndDate';
        Parameters.ParamValues['@EndDate']:=now;
        Parameters.AddParameter.Name := '@BeginASN';
        Parameters.ParamValues['@BeginASN'] := 0;
        Parameters.AddParameter.Name := '@EndASN';
        Parameters.ParamValues['@EndASN'] := 999;
      end
      else
      begin
        // Range
        Parameters.AddParameter.Name := '@RangeType';
        Parameters.ParamValues['@RangeType'] := 'S';
        Parameters.AddParameter.Name:='@FRSNumber';
        Parameters.ParamValues['@FRSNumber']:='';
        Parameters.AddParameter.Name := '@begindate';
        Parameters.ParamValues['@begindate'] := StartBox.Text;
        Parameters.AddParameter.Name := '@enddate';
        Parameters.ParamValues['@enddate'] := EndBox.Text;
        Parameters.AddParameter.Name := '@BeginASN';
        Parameters.ParamValues['@BeginASN'] := StrToInt(Start_Seq_Edit.Text);
        Parameters.AddParameter.Name := '@EndASN';
        Parameters.ParamValues['@EndASN'] := StrToInt(End_Seq_Edit.Text);
      end;
      Open;
      GetDataPanel.Visible:=FALSE;
      if recordcount = 0 then
      begin
        ShowMessage('No records');
        exit;
      end;
      DailySummaryForm := TDailySummaryForm.Create(Self);
      With DailySummaryForm Do
        Try
          DataSet       := Report_DataSet;
          ReportTypeName:= 'Passenger Glove Box';

          if SelectTypeBox.ItemIndex = 1 then
          begin
            FRSDate:='';

            SequenceBegin := Start_Seq_Edit.Text;
            SequenceEnd   := End_Seq_Edit.Text;

            DateTimeBegin:=StartBox.Text;
            DateTimeEnd:=EndBox.Text;
          end
          else
          begin
            FRSDate:=FRSBox.Text;
            SequenceBegin   := Parameters.ParamValues['@FirstFRSASN'];
            SequenceEnd := Parameters.ParamValues['@LastFRSASN'];
            DateTimeBegin:=DateTimeToStr(Report_DataSet.FieldbyNAme('BeginDate').AsDateTime);
            DateTimeEnd:=DateTimeToStr(Report_DataSet.FieldbyNAme('EndDate').AsDateTime);
          end;


          Site := AdminDataModule.SiteDateset.FieldByName('SiteAbbr').AsString;
          Preview;
        Finally
          Close;
          Screen.Cursor := crDefault;
        End;
        end;
  except
    on e:exception do
    begin
      GetDataPanel.Visible:=FALSE;
      ShowMessage('Unable to get Summary data, '+e.Message);
      exit;
    end;
  end;
end;

procedure TReportsForm.Print_INJEXDP;
begin
  try
    with Report_DataSet do
    begin
      close;
      GetDataPanel.Visible:=TRUE;
      application.processmessages;

      CommandType:=cmdStoredProc;
      CommandText:='dbo.AD_INJEXSummary';
      Parameters.Clear;

      Parameters.AddParameter.Name:='@FirstFRSASN';
      Parameters.ParamByName('@FirstFRSASN').Direction:=pdOutput;
      Parameters.ParamValues['@FirstFRSASN']:=-1;
      Parameters.AddParameter.Name:='@LastFRSASN';
      Parameters.ParamByName('@LastFRSASN').Direction:=pdOutput;
      Parameters.ParamValues['@LastFRSASN']:=-1;

      Parameters.AddParameter.Name := '@RackVehicleDescription';
      Parameters.ParamValues['@RackVehicleDescription'] := 'Passenger Door Panel';
      Parameters.AddParameter.Name := '@DataItemDescription';
      Parameters.ParamValues['@DataItemDescription'] := 'FrDoorPanel';
      Parameters.AddParameter.Name:='@LineName';
      Parameters.ParamValues['@LineName']:=Line_Combo.Text;

      if SelectTypeBox.ItemIndex = 0 then
      begin
        // FRS
        Parameters.AddParameter.Name := '@RangeType';
        Parameters.ParamValues['@RangeType'] := 'F';
        Parameters.AddParameter.Name:='@FRSNumber';
        Parameters.ParamValues['@FRSNumber']:=FRSBox.Text+'%';
        Parameters.AddParameter.Name:='@BeginDate';
        Parameters.ParamValues['@BeginDate']:=now;
        Parameters.AddParameter.Name:='@EndDate';
        Parameters.ParamValues['@EndDate']:=now;
        Parameters.AddParameter.Name := '@BeginASN';
        Parameters.ParamValues['@BeginASN'] := 0;
        Parameters.AddParameter.Name := '@EndASN';
        Parameters.ParamValues['@EndASN'] := 999;
      end
      else
      begin
        // Range
        Parameters.AddParameter.Name := '@RangeType';
        Parameters.ParamValues['@RangeType'] := 'S';
        Parameters.AddParameter.Name:='@FRSNumber';
        Parameters.ParamValues['@FRSNumber']:='';
        Parameters.AddParameter.Name := '@begindate';
        Parameters.ParamValues['@begindate'] := StartBox.Text;
        Parameters.AddParameter.Name := '@enddate';
        Parameters.ParamValues['@enddate'] := EndBox.Text;
        Parameters.AddParameter.Name := '@BeginASN';
        Parameters.ParamValues['@BeginASN'] := StrToInt(Start_Seq_Edit.Text);
        Parameters.AddParameter.Name := '@EndASN';
        Parameters.ParamValues['@EndASN'] := StrToInt(End_Seq_Edit.Text);
      end;
      Open;
      GetDataPanel.Visible:=FALSE;
      if recordcount = 0 then
      begin
        ShowMessage('No records');
        exit;
      end;
      DailySummaryForm := TDailySummaryForm.Create(Self);
      With DailySummaryForm Do
        Try
          DataSet       := Report_DataSet;
          ReportTypeName:= 'Passenger Door Panel';

          if SelectTypeBox.ItemIndex = 1 then
          begin
            FRSDate:='';

            SequenceBegin := Start_Seq_Edit.Text;
            SequenceEnd   := End_Seq_Edit.Text;

            DateTimeBegin:=StartBox.Text;
            DateTimeEnd:=EndBox.Text;
          end
          else
          begin
            FRSDate:=FRSBox.Text;
            SequenceBegin   := Parameters.ParamValues['@FirstFRSASN'];
            SequenceEnd := Parameters.ParamValues['@LastFRSASN'];
            DateTimeBegin:=DateTimeToStr(Report_DataSet.FieldbyNAme('BeginDate').AsDateTime);
            DateTimeEnd:=DateTimeToStr(Report_DataSet.FieldbyNAme('EndDate').AsDateTime);
          end;


          Site := AdminDataModule.SiteDateset.FieldByName('SiteAbbr').AsString;
          Preview;
        Finally
          Close;
          Screen.Cursor := crDefault;
        End;
        end;
  except
    on e:exception do
    begin
      GetDataPanel.Visible:=FALSE;
      ShowMessage('Unable to get Summary data, '+e.Message);
      exit;
    end;
  end;
end;

procedure TReportsForm.Print_FRS_Summary; //DMV  8/19
var
  FRS_Summary: TFRSSummary;
  row,col:integer;
  excel,mysheet : Variant;
begin

  try
    with Report_DataSet do
    begin
      close;
      CommandType:=cmdStoredProc;
      GetDataPanel.Visible:=TRUE;
      application.processmessages;
      CommandText:='AD_WQSFRS';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@begindate';
      Parameters.ParamValues['@begindate'] := StartBox.Text;
      Parameters.AddParameter.Name := '@enddate';
      Parameters.ParamValues['@enddate'] := EndBox.Text;
      Parameters.AddParameter.Name := '@Start';
      Parameters.ParamValues['@Start'] := Start_Seq_Edit.Text;
      Parameters.AddParameter.Name := '@Last';
      Parameters.ParamValues['@Last'] := StrToInt(End_Seq_Edit.Text);
      Parameters.AddParameter.Name:='@LineName';
      Parameters.ParamValues['@LineName']:=Line_Combo.Text;
      Open;
      GetDataPanel.Visible:=FALSE;
      if recordcount = 0 then
      begin
        ShowMessage('No records');
        exit;
      end;
    end;
  except
    on e:exception do
    begin
      GetDataPanel.Visible:=FALSE;
      ShowMessage('Unable to get FRS data, '+e.Message);
      exit;
    end;
  end;



  Screen.Cursor := crHourGlass;
  FRS_Summary := TFRSSummary.Create(Self);
  With FRS_Summary Do
    Try
      //DataSet       := Report_Query;
      DataSet       := Report_DataSet;
      DateEnd       := StartBox.Text;
      DateStart     := EndBox.Text;
      SequenceEnd   := End_Seq_Edit.Text;
      SequenceStart := Start_Seq_Edit.Text;
      LineName      := Line_Combo.Text;
      Site          := AdminDataModule.SiteDateset.fieldByName('SiteAbbr').AsString;
      Preview;
    Finally
      Close;
      Screen.Cursor := crDefault;
    End;

  if uppercase(Line_Combo.Text) = 'PASSENGER' then
  begin
    // fill nummi pass sheet
    try
      excel := createOleObject('Excel.Application');
      excel.visible := False;
      excel.DisplayAlerts := False;
      excel.workbooks.open(ExtractFilePath(application.ExeName)+'BlankCarFRS.xls');
      mysheet := excel.workSheets[1];

      // DMV is this still OK?
      //mysheet.Cells[6,4].value := formatdatetime('YMMDD',starttime+1)+'10'; //FRS

      mysheet.Cells[6,4].value := copy(EndBox.Text,4,1)+copy(EndBox.Text,6,2)+copy(EndBox.Text,9,2)+'10';//formatdatetime('YMMDD',StrToDateTime(EndBox.Text))+'10';
      mysheet.Cells[8,4].value := formatdatetime('m/dd/yyyy',now);  //Processed date


      with Report_DataSet do
      begin
        active:=True;
        first;
        while not eof do
        begin
          mysheet.cells[fieldbyname('ReportCell').AsInteger,7].value := fieldbyname('orders').AsInteger;
          next;
        end;
        active:=False;
      end;

      excel.ActiveWorkbook.SaveAs(AdminDataModule.fiReportDir.AsString+'\PassFRS'+formatdatetime('yyyymmddhhmm',now)+'.xls');
      if MessageDlg('Print the excel report?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
        mysheet.PageSetup.PrintArea := '$A$1:$F$49';
        mysheet.PrintOut (EmptyParam, EmptyParam, EmptyParam,
                          EmptyParam, EmptyParam, EmptyParam, EmptyParam);
      end;
      excel.Workbooks.Close;
      excel.Quit;
      excel:=Unassigned;
      mysheet:=Unassigned;
    except
      on e:exception do
      begin
        ShowMessage('Error on tire('+Report_DataSet.fieldbyname('Tire').AsString+') FRS Failed, '+e.Message);
        excel.Workbooks.Close;
        excel.Quit;
        excel:=Unassigned;
        mysheet:=Unassigned;
      end;
    end;
  end
  else
  begin
    try
      // fill nummi truck sheet
      excel := createOleObject('Excel.Application');
      excel.visible := False;
      excel.DisplayAlerts := False;
      excel.workbooks.open(ExtractFilePath(application.ExeName)+'BlankTruckFRS.xls');
      mysheet := excel.workSheets[1];

      // DMV is this still OK?
      //mysheet.Cells[6,4].value := formatdatetime('YMMDD',starttime+1)+'10'; //FRS
      mysheet.Cells[8,4].value := copy(EndBox.Text,4,1)+copy(EndBox.Text,6,2)+copy(EndBox.Text,9,2)+'10';//formatdatetime('YMMDD',StrToDateTime(StartBox.Text))+'10';
      //mysheet.Cells[8,5].value := formatdatetime('YMMDD',starttime+1)+'10'; //FRS
      mysheet.Cells[9,4].value := formatdatetime('m/dd/yyyy',now);  //Processed date


      with Report_DataSet do
      begin
        //active:=True;
        first;
        while not eof do
        begin
          if pos(',',fieldbyname('ReportCell').AsString) > 0 then
          begin
            row:=StrToInt(copy(fieldbyname('ReportCell').AsString,1,pos(',',fieldbyname('ReportCell').AsString)-1));
            col:=StrToInt(copy(fieldbyname('ReportCell').AsString,pos(',',fieldbyname('ReportCell').AsString)+1,2));
            mysheet.cells[row,col].value := fieldbyname('orders').AsInteger
          end
          else
            mysheet.cells[fieldbyname('ReportCell').AsInteger,8].value := fieldbyname('orders').AsInteger;
          next;
        end;
        //active:=False;
      end;

      excel.ActiveWorkbook.SaveAs(AdminDataModule.fiReportDir.AsString+'\TruckFRS'+formatdatetime('yyyymmddhhmm',now)+'.xls');
      if MessageDlg('Print the excel report?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
        mysheet.PageSetup.PrintArea := '$A$1:$G$47';
        mysheet.PrintOut (EmptyParam, EmptyParam, EmptyParam,
                          EmptyParam, EmptyParam, EmptyParam, EmptyParam);
      end;
      excel.Workbooks.Close;
      excel.Quit;
      excel:=Unassigned;
      mysheet:=Unassigned;
    except
      on e:exception do
      begin
        ShowMessage('Error on tire('+Report_DataSet.fieldbyname('Tire').AsString+') FRS Failed, '+e.Message);
        excel.Workbooks.Close;
        excel.Quit;
        excel:=Unassigned;
        mysheet:=Unassigned;
      end;
    end;
  end;
end;

{procedure TReportsForm.Print_Shipping_Summary;
var
  //CarpetShipping: TCarpetShipping;
  starttime,endtime:TDateTime;
  temptime:string;
begin

  temptime:=DateToStr(Startdate.Date)+' 03:00:00AM';
  starttime:=StrToDateTime(temptime);
  if StartDate.date = Enddate.Date then
  begin
    endtime:=starttime+1;
  end
  else
  begin
    temptime:=DateToStr(EndDate.Date)+' 03:00:00AM';
    endtime:=StrToDateTime(temptime);
  end;

  With Report_Query, Report_Query.SQL Do
  begin
    Close;
    Clear;
    Add('SELECT SUBSTRING(CONVERT(CHAR, Printed, 102), 1, 10) AS SORT_DATE,');
    Add('       SUBSTRING(CONVERT(CHAR, Printed, 1), 1, 10) AS SHIP_DATE,');
    Add('       SUBSTRING(Input, 6, 4) AS KANBAN,');
    Add('       COUNT(*) AS DELIVERED');
    Add('FROM Activity');
    if Start_Seq_Edit.Text > End_Seq_Edit.Text then
    begin
      Add('WHERE (SUBSTRING(input,2,4) >= :STARTSEQ ' );
      Add('OR SUBSTRING(input,2,4) <= :ENDSEQ) ');
    end
    else
    begin
      Add('WHERE SUBSTRING(input,2,4) >= :STARTSEQ ' );
      Add('AND SUBSTRING(input,2,4) <= :ENDSEQ ');
    end;
    Add('and printed between :STARTTIME and :ENDTIME ');
    Add('  AND Input IS NOT NULL');
    Add('  AND Printed IS NOT NULL');
    Add('GROUP BY SUBSTRING(CONVERT(CHAR, Printed, 102), 1, 10),');
    Add('         SUBSTRING(CONVERT(CHAR, Printed, 1), 1, 10),');
    Add('         SUBSTRING(Input, 6, 4)');
    Add('ORDER BY SORT_DATE, KANBAN');
    Parameters.ParamByName('STARTSEQ').Value:=Start_Seq_Edit.Text;
    Parameters.ParamByName('ENDSEQ').Value:=End_Seq_Edit.Text;
    Parameters.ParamByName('STARTTIME').Value:=starttime;
    Parameters.ParamByName('ENDTIME').Value:=endtime;
    Open;
  end;

  Screen.Cursor := crHourGlass;

  {CarpetShipping := TCarpetShipping.Create(Self);
  With CarpetShipping Do
    Try
      DataSet       := Report_Query;
      DateEnd       := DateToStr(EndDate.Date);
      DateStart     := DateToStr(StartDate.Date);
      LocalAlias    := FLocalAlias;
      SequenceEnd   := End_Seq_Edit.Text;
      SequenceStart := Start_Seq_Edit.Text;
      Preview;
    Finally
      Close;
      Screen.Cursor := crDefault;
    End;
end;}


{procedure TReportsForm.Print_Delivery_Summary;
var
  //CarpetDelivery: TCarpetDelivery;
  starttime,endtime:TDateTime;
  temptime:string;
begin


  temptime:=DateToStr(Startdate.Date)+' 03:00:00AM';
  starttime:=StrToDateTime(temptime);
  if StartDate.Date = EndDate.Date then
  begin
    endtime:=starttime+1;
  end
  else
  begin
    temptime:=DateToStr(EndDate.Date)+' 03:00:00AM';
    endtime:=StrToDateTime(temptime);
  end;

  With Report_Query, Report_Query.SQL Do
  begin
    Close;
    Clear;
    Add('SELECT SUBSTRING(CONVERT(CHAR, Printed, 102), 1, 10) AS SORT_DATE,');
    Add('       SUBSTRING(CONVERT(CHAR, Printed, 1), 1, 10) AS DATE,');
    Add('       MIN(SUBSTRING(Input, ' + IntToStr(FOffSet) + ', 4)) AS MIN_SEQ,');
    Add('       MAX(SUBSTRING(Input, ' + IntToStr(FOffSet) + ', 4)) AS MAX_SEQ,');
    Add('       COUNT(*) AS DELIVERED');
    Add('FROM Activity');
    Add('WHERE input IS NOT NULL');
    if Start_Seq_Edit.Text > End_Seq_Edit.Text then
    begin
      Add('AND (SUBSTRING(input,2,4) >= :STARTSEQ ' );
      Add('OR SUBSTRING(input,2,4) <= :ENDSEQ) ');
    end
    else
    begin
      Add('AND SUBSTRING(input,2,4) >= :STARTSEQ ' );
      Add('AND SUBSTRING(input,2,4) <= :ENDSEQ ');
    end;
    Add('and printed between :STARTTIME and :ENDTIME ');
    Add('GROUP BY SUBSTRING(CONVERT(CHAR, Printed, 102), 1, 10),');
    Add('         SUBSTRING(CONVERT(CHAR, Printed, 1), 1, 10)');
    Add('ORDER BY SORT_DATE');
    Parameters.ParamByName('STARTSEQ').Value:=Start_Seq_Edit.Text;
    Parameters.ParamByName('ENDSEQ').Value:=End_Seq_Edit.Text;
    Parameters.ParamByName('STARTTIME').Value:=starttime;
    Parameters.ParamByName('ENDTIME').Value:=endtime;
    Open;
  end;

  Screen.Cursor := crHourGlass;

  {CarpetDelivery := TCarpetDelivery.Create(Self);
  With CarpetDelivery Do
    Try
      DataSet       := Report_Query;
      DateEnd       := DateToStr(EndDate.Date);
      DateStart     := DateToStr(StartDate.Date);
      SequenceEnd   := End_Seq_Edit.Text;
      SequenceStart := Start_Seq_Edit.Text;
      Preview;
    Finally
      Close;
      Screen.Cursor := crDefault;
    End;
end;}


procedure TReportsForm.Set_Formats(Value: TStringList);
begin
  {With Format_Combo Do
  begin
    Items := Value;
    if Items.Count = 1 then
      ItemIndex := 0;
  end;}
end;


{procedure TReportsForm.KanBan_Display(Act_Start, Act_End: Int64);
const
  SelectSQL=  'SELECT SUBSTRING(input, 2, 4) AS Sequence, '+
              'SUBSTRING(Input, 6, 4) AS Kanban, '+
              'SUBSTRING(Output, 1, 4) AS Container, '+
              'Printed, Input, Output FROM Activity '+
              'WHERE SUBSTRING(input,2,4) >= :STARTSEQ '+
	            'AND SUBSTRING(input,2,4) <= :ENDSEQ '+
              'and printed between :STARTTIME and :ENDTIME '+
              'ORDER BY Printed';
  SelectSQLR= 'SELECT SUBSTRING(input, 2, 4) AS Sequence, '+
              'SUBSTRING(Input, 6, 4) AS Kanban, '+
              'SUBSTRING(Output, 1, 4) AS Container, '+
              'Printed, Input, Output FROM Activity '+
              'WHERE (SUBSTRING(input,2,4) >= :STARTSEQ '+
	            'OR SUBSTRING(input,2,4) <= :ENDSEQ) '+
              'and printed between :STARTTIME and :ENDTIME '+
              'ORDER BY Printed';
var
  starttime,endtime:TDateTime;
  temptime:string;
begin
  temptime:=DateToStr(Startdate.Date)+' 03:00:00AM';
  starttime:=StrToDateTime(temptime);
  if StartDate.Date = EndDate.Date then
  begin
    endtime:=starttime+1;
  end
  else
  begin
    temptime:=DateToStr(EndDate.Date)+' 03:00:00AM';
    endtime:=StrToDateTime(temptime);
  end;

  With Selected_Query, Selected_Query.SQL Do
  begin
    Close;
    Clear;

    if Start_Seq_Edit.Text > End_Seq_Edit.Text then
      Add(SelectSQLR)
    else
      Add(SelectSQL);
    Parameters.ParamByName('STARTSEQ').Value:=Start_Seq_Edit.Text;
    Parameters.ParamByName('ENDSEQ').Value:=End_Seq_Edit.Text;
    Parameters.ParamByName('STARTTIME').Value:=starttime;
    Parameters.ParamByName('ENDTIME').Value:=endtime;
    Open;
  end;
 { With Selected_Query, Selected_Query.SQL Do
  begin
    Close;
    Clear;

    Add('SELECT SUBSTRING(Input, 2, 4) AS Sequence,');
    Add('       SUBSTRING(Input, 6, 4) AS Kanban,');
    Add('       SUBSTRING(Output, 1, 4) AS Container,');
    Add('       Printed, Input, Output');
    Add('FROM Activity');
    Add('WHERE input IS NOT NULL');
    if Start_Seq_Edit.Text > End_Seq_Edit.Text then
    begin
      Add('AND (SUBSTRING(input,1,4) >= :STARTSEQ ' );
      Add('OR SUBSTRING(input,1,4) <= :ENDSEQ) ');
    end
    else
    begin
      Add('AND SUBSTRING(input,1,4) >= :STARTSEQ ' );
      Add('AND SUBSTRING(input,1,4) <= :ENDSEQ ');
    end;
    Add('and printed between :STARTTIME and :ENDTIME ');
    Add('ORDER BY Printed');
    ParamByName('STARTSEQ').Value:=Start_Seq_Edit.Text;
    ParamByName('ENDSEQ').Value:=End_Seq_Edit.Text;
    ParamByName('STARTTIME').Value:=starttime;
    ParamByName('ENDTIME').Value:=endtime;
    Open;
  end;

  Selected_Grid.Columns[0].Width := 84;
  Selected_Grid.Columns[1].Width := 84;
  Selected_Grid.Columns[2].Width := 84;
  Selected_Grid.Columns[3].Width := 144;
  Selected_Grid.Columns[4].Visible := False;
  Selected_Grid.Columns[5].Visible := False;

end;}


{procedure TReportsForm.Manifest_Display(Act_Start, Act_End: Int64); //DMV all change
const
  SelectSQL='SELECT SUBSTRING(input, 1, 4) AS Sequence,SUBSTRING(Input, 5, 4) AS Assembly, Printed, Input, Output FROM Activity '+
              'WHERE SUBSTRING(input,1,4) >= :STARTSEQ '+
	            'AND SUBSTRING(input,1,4) <= :ENDSEQ '+
              'and printed between :STARTTIME and :ENDTIME '+
              'ORDER BY Printed';
  SelectSQLR='SELECT SUBSTRING(input, 1, 4) AS Sequence,SUBSTRING(Input, 5, 4) AS Assembly, Printed, Input, Output FROM Activity '+
              'WHERE (SUBSTRING(input,1,4) >= :STARTSEQ '+
	            'OR SUBSTRING(input,1,4) <= :ENDSEQ) '+
              'and printed between :STARTTIME and :ENDTIME '+
              'ORDER BY Printed';
var
  starttime,endtime:TDateTime;
  temptime:string;
begin

  temptime:=DateToStr(Startdate.Date)+' 03:00:00AM';
  starttime:=StrToDateTime(temptime);
  if StartDate.Date = EndDate.Date then
  begin
    endtime:=starttime+1;
  end
  else
  begin
    temptime:=DateToStr(EndDate.Date)+' 03:00:00AM';
    endtime:=StrToDateTime(temptime);
  end;


  With Selected_Query, Selected_Query.SQL Do
  begin
    Close;
    Clear;

    if Start_Seq_Edit.Text > End_Seq_Edit.Text then
      Add(SelectSQLR)
    else
      Add(SelectSQL);
    Parameters.ParamByName('STARTSEQ').Value:=Start_Seq_Edit.Text;
    Parameters.ParamByName('ENDSEQ').Value:=End_Seq_Edit.Text;
    Parameters.ParamByName('STARTTIME').Value:=starttime;
    Parameters.ParamByName('ENDTIME').Value:=endtime;
    Open;
  end;

  {Selected_Grid.Columns[0].Width := 84;
  Selected_Grid.Columns[1].Width := 84;
  Selected_Grid.Columns[2].Width := 144;
  Selected_Grid.Columns[3].Visible := False;
  Selected_Grid.Columns[4].Visible := False;

end;}


procedure TReportsForm.Set_Connection(Value: TADOConnection);
begin
  FConnection := Value;
  Report_Query.Connection   := Value;
  UpdateReportQuery.Connection   := Value;
  UpdateReportQuerywf.Connection   := Value;
  Selected_Query.Connection := Value;
end;

procedure TReportsForm.Close_ButtonClick(Sender: TObject);
begin
  Close;
end;


procedure TReportsForm.Print_ButtonClick(Sender: TObject);
begin

  StatusBar.SimpleText := '';
  StatusBar.Refresh;

  if (printer_combo.ItemIndex = -1) and (printer_combo.Visible) then
  begin
    MessageDlg('Please select a printer.', mtWarning, [mbOk], 0);
    printer_combo.SetFocus;
    Exit;
  end;


  if SelectTypeBox.ItemIndex = 1 then
  begin
    if Start_Seq_Edit.Text = '' then
    begin
      MessageDlg('Please select a valid start seq.', mtWarning, [mbOk], 0);
      Start_Seq_Edit.SetFocus;
      Exit;
    end;

    if End_Seq_Edit.Text = '' then
    begin
      MessageDlg('Please select a valid end seq.', mtWarning, [mbOk], 0);
      End_Seq_Edit.SetFocus;
      Exit;
    end;
  end
  else
  begin
    if FRSPanel.Visible then
      begin
      if FRSBox.Text = '' then
      begin
        MessageDlg('Please select a valid FRS Number.', mtWarning, [mbOk], 0);
        FRSBox.SetFocus;
      end;
    end;
  end;

  case fReportType of
    rtWQSFRS:
      Print_FRS_Summary;
    rtWQSSHIP:
      Print_Assembly_Summary;
    rtEDI810:
      Print_EDI810;
    rtEDI856:
      Print_EDI856;
    rtINJEXGB:
      Print_INJEXGB;
    rtINJEXDP:
      Print_INJEXDP;
    end;
end;


procedure TReportsForm.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action:=caFree;
end;

procedure TReportsForm.FormDestroy(Sender: TObject);
begin
  ReportsForm:=nil;
end;

procedure TReportsForm.StartBoxChange(Sender: TObject);
begin
  Start_Seq_EditChange(End_Seq_Edit);
end;

procedure TReportsForm.Start_Seq_EditChange(Sender: TObject);
begin
  if fInitSeq then
    exit;

  if (length(TEdit(Sender).Text) =  TEdit(Sender).MaxLength) and (Line_Combo.Text <> '')  then
  begin
    // Get Date Time of seq
    try
      GetLastPrint_DataSet.Close;
      GetLastPrint_DataSet.CommandType:=cmdStoredProc;
      GetLastPrint_DataSet.CommandText:='AD_GetLastPrint';
      GetLastPrint_DataSet.Parameters.Clear;
      GetLastPrint_DataSet.Parameters.AddParameter.Name:='@ASN';
      GetLastPrint_DataSet.Parameters.ParamValues['@ASN']:=TEdit(Sender).Text;
      GetLastPrint_DataSet.Parameters.AddParameter.Name:='@LineName';
      GetLastPrint_DataSet.Parameters.ParamValues['@LineName']:=Line_Combo.Text;
      GetLastPrint_DataSet.Parameters.AddParameter.Name:='@RevCount';
      GetLastPrint_DataSet.Parameters.ParamValues['@RevCount']:=AdminDataModule.fiReverseCheckCount.AsInteger;
      GetLastPrint_DataSet.Open;
      if GetLastPrint_DataSet.RecordCount > 0 then
      begin
        if TEdit(Sender).Name='Start_Seq_Edit' then
        begin
          // Fill Start drop down with all available
          StartBox.Items.Clear;
          while not GetLastPrint_DataSet.Eof do
          begin
            StartBox.Items.Add(FormatDateTime('yyyy-mm-dd hh:mm:ss.zzz', GetLastPrint_DataSet.fieldbyname('LastTime').AsDateTime));
            GetLastPrint_DataSet.Next;
          end;
          StartBox.ItemIndex:=0;

          // Clear End
          EndBox.Items.Clear;
          End_Seq_Edit.Text:='';
          End_Seq_Edit.SetFocus;
        end
        else
        begin
          EndBox.Items.Clear;
          while not GetLastPrint_DataSet.Eof do
          begin
            if StartBox.Text < FormatDateTime('yyyy-mm-dd hh:mm:ss.zzz',GetLastPrint_DataSet.fieldbyname('LastTime').AsDateTime) then
            begin
              EndBox.Items.Add(FormatDateTime('yyyy-mm-dd hh:mm:ss.zzz',GetLastPrint_DataSet.fieldbyname('LastTime').AsDateTime));
            end;
            GetLastPrint_DataSet.Next;
          end;
          if EndBox.Items.Count > 0 then
            EndBox.ItemIndex:=0
          else
          begin
            showMessage('Start date is after end date, changing to previous date for starting sequence.');
            StartBox.ItemIndex:=StartBox.ItemIndex+1;
            Start_Seq_EditChange(sender);
          end;
        end;
      end
      else
        ShowMessage('Sequence not found');

    except
      on e:exception do
      begin
        GetLastPrint_DataSet.Connection.Close;
        ShowMessage('Unable to access sequence number date/time, '+e.Message);
      end;
    end;
  end;
end;

procedure TReportsForm.FormShow(Sender: TObject);
begin
  if (fReportType = rtINJEXGB) or (fReportType = rtINJEXDP) then
  begin
    SelectTypeLabel.Visible:=TRUE;
    SelectTypeBox.Visible:=TRUE;
    SelectTypeBox.ItemIndex:=0;
    FRSPanel.Visible:=TRUE;
    FirstPanel.Visible:=FALSE;
    LAstPanel.Visible:=FALSE;

    try
      FRSDataSet.Close;
      FRSDataSet.Open;

      FRSBox.Items.Clear;

      while not FRSDataSet.Eof do
      begin
        FRSBox.Items.Add(FRSDataSet.fieldbyname('FRS').AsString);
        FRSDataSet.Next;
      end;
    except
      on e:exception do
      begin
        ShowMessage('Unable to get FRS Data,'+e.message);
      end;
    end;

  end;
  Line_Combo.SetFocus;
end;

procedure TReportsForm.SelectTypeBoxChange(Sender: TObject);
begin
  if SelectTypeBox.itemindex = 0 then
  begin
    FRSPanel.Visible:=TRUE;
    FRSBox.SetFocus;
    FirstPanel.Visible:=FALSE;
    LAstPanel.Visible:=FALSE;
  end
  else
  begin
    FRSPanel.Visible:=FALSE;
    FirstPanel.Visible:=TRUE;
    LAstPanel.Visible:=TRUE;
    Start_Seq_Edit.SetFocus;
  end;
end;

procedure TReportsForm.Start_Seq_EditExit(Sender: TObject);
begin
  if length(TEDIT(Sender).Text) < 3 then
  begin
    if TEDIT(Sender).Text <> '' then
      TEDIT(Sender).Text:=format('%.3d',[StrToInt(TEDIT(Sender).Text)]);
    //ShowMessage('Must use a 3 digit sequence number');
    //TEDIT(Sender).SetFocus;
  end;
end;

procedure TReportsForm.Line_ComboChange(Sender: TObject);
begin
  if firstpanel.Visible then
  begin
    endbox.Items.Clear;
    end_seq_edit.Text:='';
    startbox.Items.Clear;
    start_seq_edit.Text:='';
    start_seq_edit.SetFocus;
  end;
end;

end.
