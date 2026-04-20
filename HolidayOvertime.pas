//****************************************************************
//
//       Admin
//
//       Copyright (c) 2006 Failproof Manufacturing Systems.
//
//****************************************************************
//
//  02/17/2006  David Verespey  Initial creation
unit HolidayOvertime;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ADODB, DB, StdCtrls, ComCtrls, Grids, DBGrids, ExtCtrls,
  AdvCombo, ColCombo, NUMMIColumnComboBox ;

type
  THolidayOvertimeForm = class(TForm)
    DBGrid1: TDBGrid;
    Panel1: TPanel;
    ClearButton: TButton;
    InsertButton: TButton;
    UpdateButton: TButton;
    DeleteButton: TButton;
    Panel2: TPanel;
    LineNameComboBox: TComboBox;
    SpecialDate: TDateTimePicker;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    DescriptionnEdit: TEdit;
    SearchButton: TButton;
    CloseButton: TButton;
    EventType_NUMMIColumnComboBox: TNUMMIColumnComboBox;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ClearButtonClick(Sender: TObject);
    procedure DataSource1DataChange(Sender: TObject; Field: TField);
    procedure FormDestroy(Sender: TObject);
    procedure SearchButtonClick(Sender: TObject);
    procedure InsertButtonClick(Sender: TObject);
    procedure UpdateButtonClick(Sender: TObject);
    procedure DeleteButtonClick(Sender: TObject);
    procedure LineNameComboBoxChange(Sender: TObject);
    procedure SpecialDateChange(Sender: TObject);
    procedure CloseButtonClick(Sender: TObject);
  private
    { Private declarations }
    fChanged:Boolean;
    procedure Clear;
  public
    { Public declarations }
    procedure Execute;
  end;

var
  HolidayOvertimeForm: THolidayOvertimeForm;

implementation

uses DataModule;

{$R *.dfm}
procedure Execute;
begin
end;

//
//  Buttons
//
procedure THolidayOvertimeForm.Clear;
begin
  LineNameComboBox.Text:='';
  LineNameComboBox.ItemIndex:=0;
  ProductionDateStatusComboBox.Text:='';
  ProductionDateStatusComboBox.ItemIndex:=0;
  DescriptionnEdit.Text:='';
  SpecialDate.DateTime:=now;
end;

procedure THolidayOvertimeForm.ClearButtonClick(Sender: TObject);
begin
  Clear;
  InsertButton.Enabled:=TRUE;
  UpdateButton.Enabled:=FALSE;
  DeleteButton.Enabled:=FALSE;
  SearchButton.Enabled:=TRUE;
end;

procedure THolidayOvertimeForm.SearchButtonClick(Sender: TObject);
var
  bm:TBookMark;
  found:boolean;
begin
  try
    found:=TRUE;
    bm:=SpecialDatesDataSet.GetBookmark;
    SpecialDatesDataSet.DisableControls;
    SpecialDatesDataSet.First;
    while not SpecialDatesDataSet.Eof do
    begin
      if formatdatetime('yyyymmdd',SpecialDate.DateTime) = formatdateTime('yyyymmdd',SpecialDatesDataSet.FieldByName('Date').AsDateTime) then
      begin
        found:=TRUE;
        break;
      end;
      SpecialDatesDataSet.Next;
    end;
    if not found then
    begin
      SpecialDatesDataSet.GotoBookmark(bm);
      ShowMessage('Date not found');
    end;
    fChanged:=FALSE;
  finally
    SpecialDatesDataSet.EnableControls;
  end;
end;

procedure THolidayOvertimeForm.InsertButtonClick(Sender: TObject);
begin
//                INSERT
  if fChanged then
  begin
    if LineNameComboBox.Text = '' then
    begin
      ShowMessage('Must select a Line');
      LineNameComboBox.SetFocus;
      exit;
    end;
    if ProductionDateStatusComboBox.Text = '' then
    begin
      ShowMessage('Must select a Date Status');
      ProductionDateStatusComboBox.SetFocus;
      exit;
    end;
    if formatdatetime('yyyymmdd',SpecialDate.DateTime) < formatdatetime('yyyymmdd',now) then
    begin
      ShowMessage('Date cannot have already passed');
      SpecialDate.SetFocus;
      exit;
    end;
    try
      SpecialDatesCommand.CommandType:=cmdStoredProc;
      SpecialDatesCommand.CommandText:='AD_InsertSpecialDate';

      SpecialDatesCommand.Parameters.Clear;
      SpecialDatesCommand.Parameters.AddParameter.Name:='@LineName';
      SpecialDatesCommand.Parameters.ParamValues['@LineName']:=LineNameComboBox.Text;
      SpecialDatesCommand.Parameters.AddParameter.Name:='@ProductionStatus';
      SpecialDatesCommand.Parameters.ParamValues['@ProductionStatus']:=ProductionDateStatusComboBox.Text;
      SpecialDatesCommand.Parameters.AddParameter.Name:='@SpecialDateName';
      SpecialDatesCommand.Parameters.ParamValues['@SpecialDateName']:=DescriptionnEdit.Text;
      SpecialDatesCommand.Parameters.AddParameter.Name:='@SpecialDate';
      SpecialDatesCommand.Parameters.ParamValues['@SpecialDate']:=SpecialDate.DateTime;
      SpecialDatesCommand.Execute;
      Refresh;
    except
      on e:exception do
      begin
        ShowMessage('Unable to insert record, '+e.Message);
        Refresh;
      end;
    end;
  end
  else
  begin
    ShowMessage('Nothing to Insert');
  end;
end;

procedure THolidayOvertimeForm.UpdateButtonClick(Sender: TObject);
begin
//                UPDATE
  if fChanged then
  begin
    if LineNameComboBox.Text = '' then
    begin
      ShowMessage('Must select a Line');
      LineNameComboBox.SetFocus;
      exit;
    end;
    if ProductionDateStatusComboBox.Text = '' then
    begin
      ShowMessage('Must select a Date Status');
      ProductionDateStatusComboBox.SetFocus;
      exit;
    end;
    if formatdatetime('yyyymmdd',SpecialDate.DateTime) < formatdatetime('yyyymmdd',now) then
    begin
      ShowMessage('Date cannot have already passed');
      SpecialDate.SetFocus;
      exit;
    end;
    try
      SpecialDatesCommand.CommandType:=cmdStoredProc;
      SpecialDatesCommand.CommandText:='AD_UpdateSpecialDate';

      SpecialDatesCommand.Parameters.Clear;
      SpecialDatesCommand.Parameters.AddParameter.Name:='@SpecialDateID';
      SpecialDatesCommand.Parameters.ParamValues['@SpecialDateID']:=SpecialDatesDataSet.FieldByName('SpecialDateID').AsInteger;
      SpecialDatesCommand.Parameters.AddParameter.Name:='@LineName';
      SpecialDatesCommand.Parameters.ParamValues['@LineName']:=LineNameComboBox.Text;
      SpecialDatesCommand.Parameters.AddParameter.Name:='@ProductionStatus';
      SpecialDatesCommand.Parameters.ParamValues['@ProductionStatus']:=ProductionDateStatusComboBox.Text;
      SpecialDatesCommand.Parameters.AddParameter.Name:='@SpecialDateName';
      SpecialDatesCommand.Parameters.ParamValues['@SpecialDateName']:=DescriptionnEdit.Text;
      SpecialDatesCommand.Parameters.AddParameter.Name:='@SpecialDate';
      SpecialDatesCommand.Parameters.ParamValues['@SpecialDate']:=SpecialDate.DateTime;
      SpecialDatesCommand.Execute;
      Refresh;
    except
      on e:exception do
      begin
        ShowMessage('Unable to update record, '+e.Message);
        Refresh;
      end;
    end;
  end
  else
  begin
    ShowMessage('Nothing to Update');
  end;
end;

procedure THolidayOvertimeForm.DeleteButtonClick(Sender: TObject);
begin
//              DELETE
  try
    if MessageDlg('Delete the record['+SpecialDatesDataSet.FieldByName('Description').AsString+':'+FormatDateTime('mm/dd/yyyy',SpecialDatesDataSet.FieldByName('Date').AsDateTime)+'] for Line('+LineNameComboBox.Text+')?',
      mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      SpecialDatesCommand.CommandType:=cmdStoredProc;
      SpecialDatesCommand.CommandText:='AD_DeleteSpecialDate';

      SpecialDatesCommand.Parameters.Clear;
      SpecialDatesCommand.Parameters.AddParameter.Name:='@SpecialDateID';
      SpecialDatesCommand.Parameters.ParamValues['@SpecialDateID']:=SpecialDatesDataSet.FieldByName('SpecialDateID').AsInteger;
      SpecialDatesCommand.Execute;
      Refresh;
    end;
  except
    on e:exception do
    begin
      ShowMessage('Unable to delete record, '+e.Message);
    end;
  end;
end;

procedure THolidayOvertimeForm.FormCreate(Sender: TObject);
begin
  // load combos
  try
    With Data_Module do
    Begin
      Inv_DataSet.Filter:='';
      Inv_DataSet.Filtered:=FALSE;
      SelectSingleFieldALC('LINE', 'LineName', LineNameComboBox);
      GetOvertimeHolidayInfo;
      JustifyColumns(Event_DBGrid);
    End;
  except
    on e:exception do
    begin
      ShowMessage('Unable to get Holiday/Overtime Data, '+e.Message);
      Close;
    end;
  end;
end;

procedure THolidayOvertimeForm.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action:=caFree;
end;


procedure THolidayOvertimeForm.FormDestroy(Sender: TObject);
begin
  HolidayOvertimeForm:=nil;
end;


procedure THolidayOvertimeForm.LineNameComboBoxChange(Sender: TObject);
begin
  fChanged:=TRUE;
end;

procedure THolidayOvertimeForm.SpecialDateChange(Sender: TObject);
begin
  LineNameComboBoxChange(self);
end;

procedure THolidayOvertimeForm.CloseButtonClick(Sender: TObject);
begin
  Close;
end;

end.
