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
unit LogisticsMaster;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, ExtCtrls, Buttons, DB, DBGrids, Filectrl;

type
  TLogisticsMaster_Form = class(TForm)
    SupplierMaster_Label: TLabel;
    LogisticsMaster_Panel: TPanel;
    Address_Label: TLabel;
    PhoneNum_Label: TLabel;
    Person_Label: TLabel;
    SupplierName_Label: TLabel;
    FaxNum_Label: TLabel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Address_Edit: TEdit;
    PhoneNum_Edit: TEdit;
    Person_Edit: TEdit;
    LogisticsName_Edit: TEdit;
    FaxNum_Edit: TEdit;
    Directory_Edit: TEdit;
    Email_Edit: TEdit;
    City_Edit: TEdit;
    State_Edit: TEdit;
    Zip_Edit: TEdit;
    ManagementButtons_Panel: TPanel;
    Insert_Button: TButton;
    Update_Button: TButton;
    Search_Button: TButton;
    Clear_Button: TButton;
    Close_Button: TButton;
    Delete_Button: TButton;
    Breakdown_SpeedButton: TSpeedButton;
    LogisticsMaster_DBGrid: TDBGrid;
    LogisticsMAster_DataSource: TDataSource;
    procedure FormCreate(Sender: TObject);
    procedure Search_ButtonClick(Sender: TObject);
    procedure Clear_ButtonClick(Sender: TObject);
    procedure Delete_ButtonClick(Sender: TObject);
    procedure Update_ButtonClick(Sender: TObject);
    procedure Insert_ButtonClick(Sender: TObject);
    procedure Breakdown_SpeedButtonClick(Sender: TObject);
    procedure LogisticsMaster_DBGridKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure LogisticsMaster_DBGridMouseUp(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure LogisticsMAster_DataSourceDataChange(Sender: TObject;
      Field: TField);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
    procedure HoldDetails(fFromGrid: Boolean);
    procedure SetDetailBoxes;
    function SearchGrid(fLogName: String): Boolean;
  public
    { Public declarations }
    function Execute: boolean;

  end;

var
  LogisticsMaster_Form: TLogisticsMaster_Form;

implementation

uses DataModule;

{$R *.dfm}
function TLogisticsMaster_Form.Execute: boolean;
Begin
  Result:= True;
  try
    try
      ShowModal;
    except
      On E:Exception do
        showMessage('Unable to generate Logistics Master screen.'
          + #13 + 'ERROR:'
          + #13 + E.Message);
    end;
  finally
    If ModalResult = mrCancel Then
      Result:= False;
  end;
end;    //Execute

procedure TLogisticsMaster_Form.FormCreate(Sender: TObject);
begin
  Data_Module.Inv_DataSet.Filter:='';
  Data_Module.Inv_DataSet.Filtered:=FALSE;
  Data_Module.GetLogisticsInfo;
  LogisticsMaster_DataSource.DataSet:=Data_Module.Inv_DataSet;
  LogisticsName_Edit.Text := '';
  Data_Module.ClearControls(LogisticsMaster_Panel);
end;

procedure TLogisticsMaster_Form.SetDetailBoxes;
Begin
  With Data_Module do
  Begin
    LogisticsName_Edit.Text := LogisticsName;
    Address_Edit.Text       := LogisticsAddress;
    PhoneNum_Edit.Text      := LogisticsTelephone;
    FaxNum_Edit.Text        := LogisticsFax;
    Person_Edit.Text        := LogisticsPerson;
    Directory_Edit.Text     := LogisticsDirectory;
    City_Edit.Text          := LogisticsCity;
    State_Edit.Text         := LogisticsState;
    Zip_edit.Text           := LogisticsZip;
    Email_Edit.Text         := LogisticsEmail;
  End;      //With
End;          //SetDetailBoxes

procedure TLogisticsMaster_Form.HoldDetails(fFromGrid: Boolean);
Begin
  If fFromGrid Then
  Begin
    with LogisticsMaster_DBGrid.DataSource.DataSet do
    Begin
      Data_Module.LogisticsName      := Fields[0].AsString;
      Data_Module.LogisticsAddress   := Fields[1].AsString;
      Data_Module.LogisticsCity      := Fields[2].AsString;
      Data_Module.LogisticsState     := Fields[3].AsString;
      Data_Module.LogisticsZip       := Fields[4].AsString;
      Data_Module.LogisticsTelephone := Fields[5].AsString;
      Data_Module.LogisticsFax       := Fields[6].AsString;
      Data_Module.LogisticsPerson    := Fields[7].AsString;
      Data_Module.LogisticsEmail     := Fields[8].AsString;
      Data_Module.LogisticsDirectory := Fields[9].AsString;
      Data_Module.RecordID           := Fields[10].AsInteger;
      Fields[10].Visible:=FALSE;
    End;      //With
  End
  Else
  Begin
    With Data_Module do
    Begin
      LogisticsName      := LogisticsName_Edit.Text;
      LogisticsAddress   := Address_Edit.Text;
      LogisticsCity      := City_Edit.Text;
      LogisticsState     := State_Edit.Text;
      LogisticsZip       := Zip_edit.Text;
      LogisticsTelephone := PhoneNum_Edit.Text;
      LogisticsFax       := FaxNum_Edit.Text;
      LogisticsPerson    := Person_edit.Text;
      LogisticsEmail     := Email_Edit.Text;
      LogisticsDirectory := Directory_Edit.Text;
    End;      //With
  End;
End;          //HoldDetails


procedure TLogisticsMaster_Form.Search_ButtonClick(Sender: TObject);
var
  fFound: Boolean;
Begin

  fFound := SearchGrid(LogisticsName_Edit.Text);
  If fFound Then
  Begin
    SetDetailBoxes;
  End
  Else
  Begin
    ShowMessage('No matches were found for your query.');
  End;

  LogisticsName_Edit.SelectAll;
  LogisticsName_Edit.SetFocus;
end;

function TLogisticsMaster_Form.SearchGrid(fLogName: String): Boolean;
var
  fBookmark: String;
begin
  Result := False;
  Data_Module.ClearControls(LogisticsMaster_Panel);

  try
    With Data_Module.Inv_DataSet do
    Begin
      fBookmark := Bookmark;
      First;
      While Not EOF do
      Begin
        If fLogName = Trim(FieldByName('LOGISTICS NAME').AsString) Then
        Begin
          Result := True;
          HoldDetails(True);
          fBookmark := Bookmark;
          Break;
        End;
        Next;
      End;      //While
    End;      //With
  except
    on e:exception do
      ShowMessage('Error in Search' + #13 + e.Message);
  end;      //try...except
  Data_Module.Inv_DataSet.Bookmark := fBookmark;
End;        //SearchGrid

procedure TLogisticsMaster_Form.Clear_ButtonClick(Sender: TObject);
begin
  LogisticsName_Edit.Text := '';
  LogisticsName_Edit.SetFocus;
  Data_Module.ClearControls(LogisticsMaster_Panel);
end;

procedure TLogisticsMaster_Form.Delete_ButtonClick(Sender: TObject);
begin
  HoldDetails(False);
  If MessageDlg('Are you sure you wish to delete' + #13 +
                 Data_Module.LogisticsName + ' from the database?',
                 mtWarning, [mbYes, mbNo], 0) = mrYes Then
  Begin
    Data_Module.DeleteLogisticsInfo;
    LogisticsName_Edit.Text := '';
    Data_Module.GetLogisticsInfo;
    SearchGrid(Data_Module.LogisticsName);
  End;
  LogisticsName_Edit.SelectAll;
  LogisticsName_Edit.SetFocus;
end;

procedure TLogisticsMaster_Form.Update_ButtonClick(Sender: TObject);
begin
  HoldDetails(False);
  Data_Module.UpdateLogisticsInfo;
  Data_Module.GetLogisticsInfo;
  Data_Module.Inv_DataSet.Locate('LOGISTICS NAME',LogisticsName_Edit.Text,[]);
  SetDetailBoxes;
end;

procedure TLogisticsMaster_Form.Insert_ButtonClick(Sender: TObject);
begin
  HoldDetails(False);
  If Not Data_Module.InsertLogisticsInfo Then
    MessageDlg('Unable to INSERT ' + Data_Module.LogisticsName, mtInformation, [mbOk], 0);

  Data_Module.GetLogisticsInfo;
  Data_Module.Inv_DataSet.Locate('LOGISTICS NAME',LogisticsName_Edit.Text,[]);
  SetDetailBoxes;
  LogisticsName_Edit.SelectAll;
  LogisticsName_Edit.SetFocus;
end;

procedure TLogisticsMaster_Form.Breakdown_SpeedButtonClick(Sender: TObject);
var
  dir:string;
begin
  if DirectoryExists(Directory_edit.Text) then
    dir:=Directory_edit.Text
  else
    dir:='c:\';

  if SelectDirectory('Select A Directory','My Computer',dir) then
    Directory_Edit.Text:=Dir;

end;


procedure TLogisticsMaster_Form.LogisticsMaster_DBGridKeyUp(
  Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  HoldDetails(True);
  SetDetailBoxes;
end;

procedure TLogisticsMaster_Form.LogisticsMaster_DBGridMouseUp(
  Sender: TObject; Button: TMouseButton; Shift: TShiftState; X,
  Y: Integer);
begin
  HoldDetails(True);
  SetDetailBoxes;
end;

procedure TLogisticsMaster_Form.LogisticsMAster_DataSourceDataChange(
  Sender: TObject; Field: TField);
begin
  HoldDetails(True);
  SetDetailBoxes;
end;

procedure TLogisticsMaster_Form.FormShow(Sender: TObject);
begin
  SetDetailBoxes;
  LogisticsName_Edit.SetFocus;
end;

end.

