unit FirstProductiionDay;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, BmDtEdit, StdCtrls, Grids, DBGrids, DateUtils, ExtCtrls, DB;

type
  TFirstProductionDay_Form = class(TForm)
    ManagementButtons_Panel: TPanel;
    Insert_Button: TButton;
    Search_Button: TButton;
    Clear_Button: TButton;
    Close_Button: TButton;
    Delete_Button: TButton;
    Event_DBGrid: TDBGrid;
    Event_Panel: TPanel;
    Label1: TLabel;
    Label4: TLabel;
    Year_Edit: TEdit;
    EventDate_NUMMIBmDateEdit: TBmDateEdit;
    FirstProductionDay_DataSource: TDataSource;
    Label2: TLabel;
    WeekNumber_Edit: TEdit;
    procedure Clear_ButtonClick(Sender: TObject);
    procedure Search_ButtonClick(Sender: TObject);
    procedure Delete_ButtonClick(Sender: TObject);
    procedure Insert_ButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Event_DBGridMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Event_DBGridKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure EventDate_NUMMIBmDateEditChange(Sender: TObject);
  private
    { Private declarations }
    function HoldDetails(fFromGrid: Boolean): String;
    procedure SetDetailBoxes;
  public
    { Public declarations }
    procedure ClearForm;
    procedure Execute;
  end;

var
  FirstProductionDay_Form: TFirstProductionDay_Form;

implementation

uses DataModule;

{$R *.dfm}

procedure TFirstProductionDay_Form.ClearForm;
begin
  EventDate_NUMMIBmDateEdit.Clear;
  Year_Edit.Text:='';
  EventDate_NUMMIBmDateEdit.SetFocus;
end;

procedure TFirstProductionDay_Form.FormCreate(Sender: TObject);
begin
  //Get the grid data and set it to the grid
  With Data_Module do
  Begin
    Inv_DataSet.Filter:='';
    Inv_DataSet.Filtered:=FALSE;
    ProdYear:='';
    GetFirstProductionDayInfo;
    FirstProductionDay_DataSource.DataSet:=Inv_DataSet;
    JustifyColumns(Event_DBGrid);
    HoldDetails(True);
    SetDetailBoxes;
  End;
end;

procedure TFirstProductionDay_Form.Execute;
begin
  Showmodal;
end;

procedure TFirstProductionDay_Form.Clear_ButtonClick(Sender: TObject);
begin
  ClearForm;
end;

procedure TFirstProductionDay_Form.Search_ButtonClick(Sender: TObject);
begin
  ShowMessage('Search not available');
end;

procedure TFirstProductionDay_Form.Delete_ButtonClick(Sender: TObject);
begin
  if Data_Module.Inv_DataSet.RecordCount > 0 then
  begin
    HoldDetails(False);
    If MessageDlg('Are you sure you wish to delete' + #13 +
                  'First Production Date ' + FormatDateTime('mm/dd/yyyy',Data_Module.EventDate) + #13 +
                  'from the database?',
                   mtWarning, [mbYes, mbNo], 0) = mrYes Then
    Begin
      Data_Module.DeleteFirstProductiondayInfo;
      Data_Module.ProdYear:='';
      Data_Module.GetFirstProductionDayInfo;
      Data_Module.JustifyColumns(Event_DBGrid);
      HoldDetails(True);
      SetDetailBoxes;
    End;
  end;
end;

procedure TFirstProductionDay_Form.Insert_ButtonClick(Sender: TObject);
var
  fErrMsg: String;
begin
  fErrMsg := '';

  fErrMsg := HoldDetails(False);

  If (fErrMsg = '') Then
  Begin
    If Not Data_Module.InsertFirstProductionDayInfo Then
      MessageDlg('Unable to INSERT ' + #13 +
                'Date ' + formatdatetime('mm/dd/yyyy',Data_Module.EventDate) + #13 +
                'with event ' + Data_Module.EventType + #13 + 'on line ' + Data_Module.EventLine +
                'It already exists in the database.', mtInformation, [mbOk], 0);

      Data_Module.ProdYear:='';
      Data_Module.GetFirstProductionDayInfo;
      Data_Module.JustifyColumns(Event_DBGrid);
  End
  Else
    MessageDlg(fErrMsg, mtError, [mbOK], 0);

  SetDetailBoxes;
  EventDate_NUMMIBmDateEdit.SetFocus;

end;

function TFirstProductionDay_Form.HoldDetails(fFromGrid: Boolean): String;
var
  fErrMsg: String;
Begin
  fErrMsg := '';

  If fFromGrid Then
  Begin
    with Event_DBGrid.DataSource.DataSet do
    begin
      with Data_Module do
      begin
        EventDate := Fields[1].AsDateTime;
        ProdYear  := Fields[0].AsString;
        WeekNumber:= Fields[2].AsInteger;
      end;
    End;      //With
  End
  Else
  Begin
    With Data_Module do
    Begin
      if EventDate_NUMMIBmDateEdit.Date < StrToDate('01/01/2002') then
        fErrMsg := #13 + 'Invalid Date'
      Else
      begin
        EventDate := EventDate_NUMMIBmDateEdit.Date;
        ProdYear := Year_Edit.Text;
        WeekNumber := strToInt(WeekNumber_Edit.text);
      end;

    End;      //With
  End;
  If Not (fErrMsg = '') Then
    fErrMsg := 'The following fields need to be corrected:' + fErrMsg;
  result := fErrMsg;
End;          //HoldDetails

procedure TFirstProductionDay_Form.SetDetailBoxes;
Begin
  With Data_Module do
  Begin
    EventDate_NUMMIBmDateEdit.Date:=EventDate;
    Year_Edit.Text:=ProdYear;
    WeekNumber_Edit.text := IntToStr(WeekNumber);
  End;      //With
End;          //SetDetailBoxes

procedure TFirstProductionDay_Form.Event_DBGridMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  HoldDetails(True);
  SetDetailBoxes;
end;

procedure TFirstProductionDay_Form.Event_DBGridKeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  HoldDetails(True);
  SetDetailBoxes;
end;

procedure TFirstProductionDay_Form.EventDate_NUMMIBmDateEditChange(
  Sender: TObject);
begin
  Year_Edit.Text:=formatdatetime('yyyy',EventDate_NUMMIBmDateEdit.Date);
  WeekNumber_Edit.Text:=IntToStr(WeekoftheYear(EventDate_NUMMIBmDateEdit.date));

end;

end.
