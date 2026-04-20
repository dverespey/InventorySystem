unit ManualForecast;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Grids, DBGrids, History, BmDtEdit,
  NUMMIBmDateEdit, ComObj, Excel2000, DB, DBTables,DateUtils,ForecastBreakdownF;

type

  TManualForecast_Form = class(TForm)
    Hist: THistory;
    Button1: TButton;
    StartButton: TButton;
    Startdate: TNUMMIBmDateEdit;
    Enddate: TNUMMIBmDateEdit;
    Label1: TLabel;
    Label2: TLabel;
    procedure Button1Click(Sender: TObject);
    procedure StartButtonClick(Sender: TObject);
    procedure EnddateChange(Sender: TObject);
    procedure StartdateChange(Sender: TObject);
  private
    { Private declarations }
    fFilename:string;
  public
    { Public declarations }
    procedure Execute;

    property filename:string
    read fFilename
    write fFilename;
  end;

var
  ManualForecast_Form: TManualForecast_Form;
  excel,mysheet:variant;

implementation

uses DataModule;

{$R *.dfm}

procedure TManualForecast_Form.Execute;
begin
  // Ask for dates
  // Load up into array like other forcast and then run import code
  ShowModal;
end;

procedure TManualForecast_Form.Button1Click(Sender: TObject);
begin
  Close;
end;

procedure TManualForecast_Form.StartButtonClick(Sender: TObject);
var
  i,weeks,x,weekvalue,j:integer;
  tcf:Textfile;
  tcl,pn:string;
  start,endd,current:TDateTime;
begin

  try
    excel := createOleObject('Excel.Application');
    excel.visible := False;
    excel.DisplayAlerts := False;
    excel.workbooks.open(fFilename);
    mysheet := excel.workSheets[1];

    AssignFile(tcf,Data_Module.fiForecastInputDir.AsString+'\buildout.prelftp');
    Rewrite(tcf);

    //look for partnumbers and create a standard forecast file
    for i:=1 to 100 do
    begin
      if not VarIsEmpty(mysheet.cells[i,1].value) then
      begin
        if pos('-',mysheet.cells[i,1].value) > 0 then
        begin
          pn:='';
          pn:=copy(mysheet.cells[i,1].value,1,5);
          pn:=pn+copy(mysheet.cells[i,1].value,7,5);
          pn:=pn+copy(mysheet.cells[i,1].value,13,2);
          Hist.Append(pn);
          tcl:='';
          tcl:=Data_Module.fiSupplierCode.AsString+pn+mysheet.cells[i,3].value; //fake kanban

          // Move start and end date to Monday and get week number and count
          //
          Start:=StartDate.Date;
          endd:=EndDate.Date;

          // Get week count
          x:=trunc(endd-start);
          weeks:=x div 7;
          weekvalue:=mysheet.cells[i,2].value div weeks;

          current:=Start;

          for j:=1 to 14 do
          begin
            if j<=weeks then
              tcl:=tcl+format('%.2d',[WeekoftheYear(current+1)])+formatdatetime('yymmdd',current)+format('%.6d',[weekvalue])
            else
              tcl:=tcl+format('%.2d',[WeekoftheYear(current+1)])+formatdatetime('yymmdd',current)+format('%.6d',[0]);

            current:=current+7; // next monday
          end;


          // write forecast line
          Hist.Append(tcl);
          Writeln(tcf,tcl);
        end;
      end;
    end;

    excel.Workbooks.Close;
    excel.Quit;
    excel:=Unassigned;
    mysheet:=Unassigned;
    CloseFile(tcf);

    // Call normal breakdown form
    {breakdown:=TForecastBreakdown_Form.Create(self);
    breakdown.filename:=Data_Module.fiForecastInputDir.AsString+'\buildout.prelftp';
    breakdown.SupplierCode:=data_Module.fiSupplierCode.AsString;
    breakdown.Show;
    if not breakdown.Execute then
      Hist.Append('Forecast has not been added, please retry');
    breakdown.Free;
    close;}
    // Display end
    StartButton.Visible:=False;
    startdate.Clear;
    enddate.Clear;
  except
    on e:exception do
    begin
      Data_Module.LogActLog('ERROR','Failed on Manual Forecast, '+e.Message);
      Hist.Append('Failed on Manual Forecast, '+e.Message);
      excel.Workbooks.Close;
      excel.Quit;
      excel:=Unassigned;
      mysheet:=Unassigned;
    end;
  end;
end;

procedure TManualForecast_Form.EnddateChange(Sender: TObject);
begin
  if enddate.Date > startdate.Date then
  begin
    if DayOfTheWeek(EndDate.Date) <> 7 then
      ShowMessage('End date must be a Sunday')
    else
      startbutton.Visible:=TRUE;
  end
  else
  begin
    if startdate.Text <> '' then
      ShowMessage('End date must be greater than start date');
    startbutton.Visible:=FALSE;
  end;
end;

procedure TManualForecast_Form.StartdateChange(Sender: TObject);
begin
  if enddate.Date > startdate.Date then
  begin
    if DayOfTheWeek(StartDate.Date) <> 7 then
      ShowMessage('Start date must be a Sunday')
    else
      startbutton.Visible:=TRUE;
  end
  else
  begin
    if enddate.Text <> '' then
      ShowMessage('Start date must be less than end date');
    startbutton.Visible:=FALSE;
  end;
end;

end.
