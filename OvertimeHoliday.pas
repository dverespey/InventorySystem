//****************************************************************
//
//       Inventory Control
//
//       Copyright (c) 2002 Failproof Manufacturing Systems.
//
//****************************************************************
//
//
//  12/16/2002  David Verespey  Initial Create
//
unit OvertimeHoliday;

interface

uses Windows, SysUtils, Classes, Graphics, Forms, Controls, StdCtrls,
  Buttons, ExtCtrls, Grids, DBGrids, BmDtEdit, NUMMIBmDateEdit, AdvCombo,
  ColCombo, NUMMIColumnComboBox,DateUtils,dialogs, DB;

type
  TOvertimeHoliday_Form = class(TForm)
    ManagementButtons_Panel: TPanel;
    Insert_Button: TButton;
    Search_Button: TButton;
    Clear_Button: TButton;
    Close_Button: TButton;
    Delete_Button: TButton;
    Event_DBGrid: TDBGrid;
    Event_Panel: TPanel;
    Label1: TLabel;
    EventType_NUMMIColumnComboBox: TNUMMIColumnComboBox;
    Label2: TLabel;
    Label3: TLabel;
    EventDate_NUMMIBmDateEdit: TBmDateEdit;
    Label4: TLabel;
    Label5: TLabel;
    WeekNumber_Edit: TEdit;
    DayNumber_Edit: TEdit;
    LineNameComboBox: TComboBox;
    Overtime_DataSource: TDataSource;
    Label6: TLabel;
    EventDescription_Edit: TEdit;
    WeekNumberCaculatedLabel: TLabel;
    WeekNumberCaculatedEdit: TEdit;
    procedure Clear_ButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure EventDate_NUMMIBmDateEditChange(Sender: TObject);
    procedure Delete_ButtonClick(Sender: TObject);
    procedure Insert_ButtonClick(Sender: TObject);
    procedure Search_ButtonClick(Sender: TObject);
    procedure Overtime_DataSourceDataChange(Sender: TObject;
      Field: TField);
  private
    { Private declarations }
    fDateChange:boolean;
    fCalculateWeek:boolean;
    function HoldDetails(fFromGrid: Boolean): String;
    procedure SetDetailBoxes;
    procedure UpdateCalculatedWeekNumber;
  public
    { Public declarations }
    procedure ClearForm;
    procedure Execute;
  end;

var
  OvertimeHoliday_Form: TOvertimeHoliday_Form;

implementation

uses DataModule;

{$R *.dfm}

procedure TOvertimeHoliday_Form.ClearForm;
begin
  EventDate_NUMMIBmDateEdit.Clear;
  LineNameComboBox.ItemIndex:=0;
  EventType_NUMMIColumnComboBox.ItemIndex:=0;
  WeekNumber_Edit.Text:='';
  DayNumber_Edit.Text:='';
  EventDate_NUMMIBmDateEdit.SetFocus;
end;

procedure TOvertimeHoliday_Form.Execute;
begin
  fDateChange:=TRUE;
  Showmodal;
end;


procedure TOvertimeHoliday_Form.Clear_ButtonClick(Sender: TObject);
begin
  ClearForm;
end;

procedure TOvertimeHoliday_Form.FormCreate(Sender: TObject);
begin
  //Get the grid data and set it to the grid
  With Data_Module do
  Begin
    ALC_DataSet.Filter:='';
    ALC_DataSet.Filtered:=FALSE;
    SelectSingleFieldALC('LINE', 'LineName', LineNameComboBox);
    // Add ALL lines option
    if LineNameComboBox.Items.Count > 2 then
    begin
      LineNameComboBox.Items.Add('ALL LINES');
    end;
    SelectMultiFieldALC('ProductionStatus', 'ProductionStatusAbv, ProductionStatus', EventType_NUMMIColumnComboBox, 'SpecialDateUse = 1');
    fCalculateWeek:=FALSE;
    GetOvertimeHolidayInfo;
    Overtime_DataSource.DataSet:=ALC_DataSet;
    JustifyColumns(Event_DBGrid);
    fCalculateWeek:=TRUE;
  End;
end;

procedure TOvertimeHoliday_Form.FormShow(Sender: TObject);
begin
  ClearForm;
  HoldDetails(True);
  SetDetailBoxes;

  if Data_Module.fiUseFirstProductionDay.AsBoolean then
  begin
    WeekNumberCaculatedEdit.Visible:=TRUE;
    WeekNumberCaculatedLabel.Visible:=TRUE;
  end
  else
  begin
    WeekNumberCaculatedEdit.Visible:=FALSE;
    WeekNumberCaculatedLabel.Visible:=FALSE;
  end;
end;

function TOvertimeHoliday_Form.HoldDetails(fFromGrid: Boolean): String;
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
        RecordID  := Fields[0].AsInteger;
        Fields[0].Visible:=FALSE;
        EventDate := Fields[1].AsDateTime;
        Fields[1].DisplayWidth:=10;
        TDateTimeField(Fields[1]).displayFormat:='mm/dd/yyyy';
        EventDescription := Fields[2].AsString;
        Fields[2].DisplayWidth:=20;
        LineName  := Fields[3].AsString;
        EventType := Fields[4].AsString;
        Fields[4].DisplayWidth:=10;
        WeekNumber:= Fields[5].AsInteger;
        DayNumber := Fields[6].AsInteger;
        EventTypeAbv := Fields[7].AsString;
      end;
    End;      //With
  End
  Else
  Begin
    With Data_Module do
    Begin
      if EventDate_NUMMIBmDateEdit.Date < (now-1) then
        fErrMsg := #13 + 'Invalid Date'
      Else
      begin
        EventDate := EventDate_NUMMIBmDateEdit.Date;
        WeekNumber := StrToInt(WeekNumber_Edit.Text);
        DayNumber := StrToInt(DayNumber_Edit.Text);
      end;

      EventDescription:=EventDescription_Edit.Text;

      if LineNameComboBox.ItemIndex <= 0 then
        fErrMsg := fErrMsg + #13 + 'Invalid line'
      else
        LineName:=LineNameComboBox.Text;

      if EventType_NUMMIColumnComboBox.ItemIndex <= 0 then
        fErrMsg := fErrMsg + #13 + 'Invalid event'
      else
      begin
        EventTypeAbv:=EventType_NUMMIColumnComboBox.ColumnItems[EventType_NUMMIColumnComboBox.ItemIndex,0];
        EventType:=EventType_NUMMIColumnComboBox.ColumnItems[EventType_NUMMIColumnComboBox.ItemIndex,1];
      end;
    End;      //With
  End;
  If Not (fErrMsg = '') Then
    fErrMsg := 'The following fields need to be corrected:' + fErrMsg;
  result := fErrMsg;
End;          //HoldDetails

procedure TOvertimeHoliday_Form.SetDetailBoxes;
Begin
  With Data_Module do
  Begin
    fDateChange:=FALSE;

    EventDate_NUMMIBmDateEdit.Date:=EventDate;

    SearchCombo(LineNameComboBox, LineName);

    SearchMultiCombo(EventType_NUMMIColumnComboBox, EventTypeAbv);

    EventDescription_Edit.Text:=EventDescription;

    WeekNumber_Edit.Text:=IntToStr(WeekNumber);

    DayNumber_Edit.Text:=IntToStr(DayNumber);

    fDateChange:=TRUE;
  End;      //With
End;          //SetDetailBoxes

procedure TOvertimeHoliday_Form.EventDate_NUMMIBmDateEditChange(
  Sender: TObject);
begin
  // Set day of week, week of year
  if fDateChange then //don't do if datasource change
  begin
    DayNumber_Edit.Text:=IntToStr(DayofTheWeek(EventDate_NUMMIBmDateEdit.date));
    WeekNumber_Edit.Text:=IntToStr(WeekoftheYear(EventDate_NUMMIBmDateEdit.date));
    UpdateCalculatedWeekNumber;
  end;
end;

procedure TOvertimeHoliday_Form.Delete_ButtonClick(Sender: TObject);
begin
  if Data_Module.ALC_DataSet.RecordCount > 0 then
  begin
    HoldDetails(False);
    If MessageDlg('Are you sure you wish to delete' + #13 +
                  'event date ' + FormatDateTime('mm/dd/yyyy',Data_Module.EventDate) + #13 +
                  'and event line ' + Data_Module.EventLine + #13 +
                  'and event type ' + Data_Module.EventType + #13 +
                  'from the database?',
                   mtWarning, [mbYes, mbNo], 0) = mrYes Then
    Begin
      Data_Module.DeleteOvertimeHolidayInfo;
      Data_Module.GetOvertimeHolidayInfo;
      Data_Module.JustifyColumns(Event_DBGrid);
      ClearForm;
      HoldDetails(True);
      SetDetailBoxes;
    End;
  end;
end;

procedure TOvertimeHoliday_Form.Insert_ButtonClick(Sender: TObject);
var
  fErrMsg: String;
begin
  fErrMsg := '';

  fErrMsg := HoldDetails(False);

  If (fErrMsg = '') Then
  Begin
    If Not Data_Module.InsertOvertimeHolidayInfo Then
      MessageDlg('Unable to INSERT ' + #13 +
                'Date ' + formatdatetime('mm/dd/yyyy',Data_Module.EventDate) + #13 +
                'with event ' + Data_Module.EventType + #13 + 'on line ' + Data_Module.LineName, mtWarning, [mbOk], 0);

      Data_Module.GetOvertimeHolidayInfo;
      Data_Module.JustifyColumns(Event_DBGrid);
  End
  Else
    MessageDlg(fErrMsg, mtError, [mbOK], 0);

  SetDetailBoxes;
  EventDate_NUMMIBmDateEdit.SetFocus;

  Event_DBGrid.Datasource.DataSet.Last;

end;        //Insert_ButtonClick

procedure TOvertimeHoliday_Form.Search_ButtonClick(Sender: TObject);
begin
  ShowMessage('Search not available');
end;


procedure TOvertimeHoliday_Form.Overtime_DataSourceDataChange(
  Sender: TObject; Field: TField);
begin
  HoldDetails(True);
  SetDetailBoxes;
  UpdateCalculatedWeekNumber;
end;

procedure TOvertimeHoliday_Form.UpdateCalculatedWeekNumber;
begin
  if fCalculateWeek then
  begin
    with data_module do
    begin
      if fiUseFirstProductionDay.AsBoolean then
      begin
        try
          Inv_StoredProc.Close;
          Inv_StoredProc.ProcedureName := 'dbo.SELECT_FirstProductionDay;1';
          Inv_StoredProc.Parameters.Clear;
          Inv_StoredProc.Parameters.AddParameter.Name := '@ProdYear';
          if (WeekoftheYear(EventDate_NUMMIBmDateEdit.date) >= 52) and (formatdatetime('yyyy',EventDate_NUMMIBmDateEdit.date) <> formatdatetime('yyyy',now)) then
            Inv_StoredProc.Parameters.ParamValues['@ProdYear'] := formatdatetime('yyyy',now)
          else
            Inv_StoredProc.Parameters.ParamValues['@ProdYear'] := formatdatetime('yyyy',EventDate_NUMMIBmDateEdit.date);

          Inv_StoredProc.Open;

          if Inv_StoredProc.FieldByName('First Week Number').AsInteger <> 1 then
          begin
            WeekNumberCaculatedEdit.Text:=IntToStr( WeekoftheYear(EventDate_NUMMIBmDateEdit.date) - (Inv_StoredProc.FieldByName('First Week Number').AsInteger-1) );
          end
          else
          begin
            WeekNumberCaculatedEdit.Text:=IntToStr(WeekoftheYear(EventDate_NUMMIBmDateEdit.date));
          end;

        except
          on e:exception do
          begin
            ShowMessage('Failed on week number get, '+e.Message);
          end;
        end;
      end;
    end;
  end;
end;

end.
