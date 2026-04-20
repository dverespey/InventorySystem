//***********************************************************
//  InventorySystem
//
//  Copyright (c) 2003-2005 TAI, Failproof Manufacturing Systems
//
//***********************************************************
//
//  03/18/2003  David Verespey  Initial creation

unit RenbanGroupMaster;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, DBGrids, ExtCtrls, DB, Mask;

type
  TRenbanGroupMaster_Form = class(TForm)
    RenbanGroupMaster_Panel: TPanel;
    SizeCode_Label: TLabel;
    RenbanGroupCode_Edit: TEdit;
    RenbanGroupMaster_DBGrid: TDBGrid;
    ManagementButtons_Panel: TPanel;
    Insert_Button: TButton;
    Update_Button: TButton;
    Search_Button: TButton;
    Clear_Button: TButton;
    Close_Button: TButton;
    Delete_Button: TButton;
    Label1: TLabel;
    RenbanGroupCount_Edit: TEdit;
    Renban_DataSource: TDataSource;
    ShipDays_Edit: TEdit;
    Label4: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    Label18: TLabel;
    ShipDaysMonday_MaskEdit: TMaskEdit;
    ShipDaysTuesday_MaskEdit: TMaskEdit;
    ShipDaysWednesday_MaskEdit: TMaskEdit;
    ShipDaysThursday_MaskEdit: TMaskEdit;
    ShipDaysFriday_MaskEdit: TMaskEdit;
    ShipDaysSaturday_MaskEdit: TMaskEdit;
    PartsStockMaster_Label: TLabel;
    procedure RenbanGroupMaster_DBGridMouseUp(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure RenbanGroupMaster_DBGridKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Insert_ButtonClick(Sender: TObject);
    procedure Update_ButtonClick(Sender: TObject);
    procedure Delete_ButtonClick(Sender: TObject);
    procedure Search_ButtonClick(Sender: TObject);
    procedure Clear_ButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure RenbanGroupCode_EditChange(Sender: TObject);
    procedure Renban_DataSourceDataChange(Sender: TObject; Field: TField);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
    procedure SetDetailBoxes;
    procedure HoldDetails(fFromGrid: Boolean);
    function SearchGrid(RenbanGroupCode: String): Boolean;
    function Validate: boolean;
  public
    { Public declarations }
    function Execute:boolean;
  end;

var
  RenbanGroupMaster_Form: TRenbanGroupMaster_Form;

implementation

uses DataModule;

{$R *.dfm}

function TRenbanGroupMaster_Form.Execute:boolean;
begin
  Result:= True;
  try
    try
      ShowModal;
    except
      On E:Exception do
        showMessage('Unable to generate Renban Group screen.'
          + #13 + 'ERROR:'
          + #13 + E.Message);
    end;
  finally
    If ModalResult = mrCancel Then
      Result:= False;
  end;
end;

procedure TRenbanGroupMaster_Form.SetDetailBoxes;
begin
  With Data_Module do
  Begin
    RenbanGroupCode_Edit.Text:=Renbancode;
    RenbanGroupCount_Edit.Text:=RenbanCount;
    ShipDays_Edit.Text:=IntToStr(Shipdays);
    ShipDaysMonday_MaskEdit.Text:=IntToStr(ShipDaysMonday);
    ShipDaysTuesday_MaskEdit.Text:=IntToStr(ShipDaysTuesday);
    ShipDaysWednesday_MaskEdit.Text:=IntToStr(ShipDaysWednesday);
    ShipDaysThursday_MaskEdit.Text:=IntToStr(ShipDaysThursday);
    ShipDaysFriday_MaskEdit.Text:=IntToStr(ShipDaysFriday);
    ShipDaysSaturday_MaskEdit.Text:=IntToStr(ShipDaysSaturday);
  End;      //With
end;

procedure TRenbanGroupMaster_Form.HoldDetails(fFromGrid: Boolean);
var
  ftempint,days:integer;
begin
  If fFromGrid Then
  Begin
    with RenbanGroupMaster_DBGrid.DataSource.DataSet do
    begin
      Data_Module.RenbanCode        := Fields[0].AsString;
      Data_Module.RenbanCount       := Fields[1].AsString;
      Data_Module.Shipdays          := Fields[2].AsInteger;
      Data_Module.ShipDaysMonday    := Fields[3].AsInteger;
      Data_Module.ShipDaysTuesday   := Fields[4].AsInteger;
      Data_Module.ShipDaysWednesday := Fields[5].AsInteger;
      Data_Module.ShipDaysThursday  := Fields[6].AsInteger;
      Data_Module.ShipDaysFriday    := Fields[7].AsInteger;
      Data_Module.ShipDaysSaturday  := Fields[8].AsInteger;
      Data_Module.RecordID          := Fields[9].AsInteger;
      Fields[9].Visible:=FALSE;
    end;
  End
  Else
  Begin
    With Data_Module do
    Begin
      RenbanCode    := RenbanGroupCode_Edit.Text;
      if length(RenbanGroupCount_Edit.Text) = 3 then
        RenbanCount   := RenbanGroupCount_Edit.Text
      else if length(RenbanGroupCount_Edit.Text) = 2 then
        RenbanCount   := '0'+RenbanGroupCount_Edit.Text
      else if length(RenbanGroupCount_Edit.Text) = 1 then
        RenbanCount   := '00'+RenbanGroupCount_Edit.Text;

      if not TryStrToInt(ShipDays_Edit.Text,days) then
        ShipDays:=0
      else
        ShipDays:=days;

      If Not TryStrToInt(Trim(ShipDaysMonday_MaskEdit.Text), fTempInt) Then
        ShipDaysMonday:= 0
      Else
        ShipDaysMonday:= fTempInt;

      If Not TryStrToInt(Trim(ShipDaysTuesday_MaskEdit.Text), fTempInt) Then
        ShipDaysTuesday:= 0
      Else
        ShipDaysTuesday:= fTempInt;

      If Not TryStrToInt(Trim(ShipDaysWednesday_MaskEdit.Text), fTempInt) Then
        ShipDaysWednesday:= 0
      Else
        ShipDaysWednesday:= fTempInt;

      If Not TryStrToInt(Trim(ShipDaysThursday_MaskEdit.Text), fTempInt) Then
        ShipDaysThursday:= 0
      Else
        ShipDaysThursday:= fTempInt;

      If Not TryStrToInt(Trim(ShipDaysFriday_MaskEdit.Text), fTempInt) Then
        ShipDaysFriday:= 0
      Else
        ShipDaysFriday:= fTempInt;

      If Not TryStrToInt(Trim(ShipDaysSaturday_MaskEdit.Text), fTempInt) Then
        ShipDaysSaturday:= 0
      Else
        ShipDaysSaturday:= fTempInt;
    End;      //With
  End;
end;

function TRenbanGroupMaster_Form.SearchGrid(RenbanGroupCode: String): Boolean;
begin
  Result := False;
  Data_Module.ClearControls(RenbanGroupMaster_Panel);
  try
    With Data_Module.Inv_DataSet Do
    Begin
      Filtered := False;
      Filter := '[Renban Group Code] LIKE ' + QuotedStr(RenbanGroupCode);
      Filtered := True;
      If RecordCount > 0 Then
      Begin
        HoldDetails(True);
        Result := True;
      End;
    End;    //With
  except
    on e:exception do
      ShowMessage('Error in Search' + #13 + e.Message);
  end;      //try...except
End;        //SearchGrid

procedure TRenbanGroupMaster_Form.RenbanGroupMaster_DBGridMouseUp(
  Sender: TObject; Button: TMouseButton; Shift: TShiftState; X,
  Y: Integer);
begin
  HoldDetails(True);
  SetDetailBoxes;
end;

procedure TRenbanGroupMaster_Form.RenbanGroupMaster_DBGridKeyUp(
  Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  HoldDetails(True);
  SetDetailBoxes;
end;

procedure TRenbanGroupMaster_Form.Insert_ButtonClick(Sender: TObject);
begin
  if Validate then
  begin
    HoldDetails(False);
    If Not Data_Module.InsertRenbanGroupInfo Then
      MessageDlg('Unable to INSERT ' + Data_Module.RenbanCode + #13 +
                'It already exists in the database.', mtInformation, [mbOk], 0);

    Data_Module.GetRenbanGroupInfo;
    Data_Module.Inv_DataSet.Filtered := False;
    RenbanGroupCode_Edit.Text := '';
    RenbanGroupCode_Edit.SetFocus;
  end;
end;

function TRenbanGroupMaster_Form.Validate: boolean;
var
  fErrMsg:string;
  fTempInt:integer;
begin
  // confirm data
  result:=True;
  fErrMsg:='';


  If Not TryStrToInt(Trim(RenbanGroupCount_Edit.Text), fTempInt) Then
    fErrMsg := #13 + 'Renban count must be a numeric';

  If Not TryStrToInt(Trim(ShipDays_Edit.Text), fTempInt) Then
    fErrMsg := #13 + 'Invalid Monday Ship Days';

  If Not TryStrToInt(Trim(ShipDaysMonday_MaskEdit.Text), fTempInt) Then
    fErrMsg := #13 + 'Invalid Monday Ship Days';

  If Not TryStrToInt(Trim(ShipDaysTuesday_MaskEdit.Text), fTempInt) Then
    fErrMsg := #13 + 'Invalid Tuesday Ship Days';

  If Not TryStrToInt(Trim(ShipDaysWednesday_MaskEdit.Text), fTempInt) Then
    fErrMsg := #13 + 'Invalid Wednesday Ship Days';

  If Not TryStrToInt(Trim(ShipDaysThursday_MaskEdit.Text), fTempInt) Then
    fErrMsg := #13 + 'Invalid Thursday Ship Days';

  If Not TryStrToInt(Trim(ShipDaysFriday_MaskEdit.Text), fTempInt) Then
    fErrMsg := #13 + 'Invalid Friday Ship Days';

  If Not TryStrToInt(Trim(ShipDaysSaturday_MaskEdit.Text), fTempInt) Then
    fErrMsg := #13 + 'Invalid Saturday Ship Days';

  if fErrMsg <> '' then
  begin
    fErrMsg := 'The following fields need to be corrected:' + fErrMsg;
    ShowMessage(fErrMsg);
    RenbanGroupCount_Edit.SetFocus;
    result:=False;
  end;

end;

procedure TRenbanGroupMaster_Form.Update_ButtonClick(Sender: TObject);
var
  renban:string;
begin
  if Validate then
  begin
    HoldDetails(False);
    renban:=RenbanGroupCode_Edit.Text;
    Data_Module.UpdateRenbanGroupInfo;
    Data_Module.GetRenbanGroupInfo;
    Data_Module.Inv_DataSet.Locate('Renban Group Code',renban,[]);
    SetDetailBoxes;
  end;
end;

procedure TRenbanGroupMaster_Form.Delete_ButtonClick(Sender: TObject);
begin
  HoldDetails(False);
  If MessageDlg('Are you sure you wish to delete' + #13 +
                 Data_Module.RenbanCode + ' from the database?',
                 mtWarning, [mbYes, mbNo], 0) = mrYes Then
  Begin
    Data_Module.DeleteRenbanGroupInfo;
    Data_Module.GetRenbanGroupInfo;
  End;
  RenbanGroupCode_Edit.SelectAll;
  RenbanGroupCode_Edit.SetFocus;
end;

procedure TRenbanGroupMaster_Form.Search_ButtonClick(Sender: TObject);
var
  fFound: Boolean;
Begin
  fFound := SearchGrid(RenbanGroupCode_Edit.Text);
  If fFound Then
  Begin
    SetDetailBoxes;
  End
  Else
  Begin
    ShowMessage('No matches were found for your query.');
  End;

  RenbanGroupCode_Edit.SelectAll;
  RenbanGroupCode_Edit.SetFocus;
end;        //Search_ButtonClick

procedure TRenbanGroupMaster_Form.Clear_ButtonClick(Sender: TObject);
begin
  Data_Module.Inv_DataSet.Filtered := False;
  RenbanGroupCode_Edit.Text       :='';
  RenbanGroupCount_Edit.Text      :='0';
  ShipDays_Edit.Text              :='0';
  ShipDaysMonday_MaskEdit.Text    :='0';
  ShipDaysTuesday_MaskEdit.Text   :='0';
  ShipDaysWednesday_MaskEdit.Text :='0';
  ShipDaysThursday_MaskEdit.Text  :='0';
  ShipDaysFriday_MaskEdit.Text    :='0';
  ShipDaysSaturday_MaskEdit.Text  :='0';
  RenbanGroupCode_Edit.SetFocus;
end;

procedure TRenbanGroupMaster_Form.FormCreate(Sender: TObject);
begin
  With Data_Module do
  Begin
    Inv_DataSet.Filter:='';
    Inv_DataSet.Filtered:=FALSE;
    GetRenbanGroupInfo;
    Renban_DataSource.DataSet:=Inv_DataSet;
    Data_Module.ClearControls(RenbanGroupMaster_Panel);
    Data_Module.Inv_DataSet.Filtered := False;
    RenbanGroupCode_Edit.Text := '';
  End;      //With
end;

procedure TRenbanGroupMaster_Form.RenbanGroupCode_EditChange(
  Sender: TObject);
begin
  If Length(RenbanGroupCode_Edit.Text) <= 1 Then
    Data_Module.ClearControls(RenbanGroupMaster_Panel);
end;

procedure TRenbanGroupMaster_Form.Renban_DataSourceDataChange(
  Sender: TObject; Field: TField);
begin
  HoldDetails(True);
  SetDetailBoxes;
end;

procedure TRenbanGroupMaster_Form.FormShow(Sender: TObject);
begin
  RenbanGroupCode_Edit.SetFocus;
  Data_Module.Inv_DataSet.First;
end;


end.
