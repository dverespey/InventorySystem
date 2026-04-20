//***********************************************************
//  InventorySystem
//
//  Copyright (c) 2002 Failproof Manufacturing Systems
//
//***********************************************************
//
//  10/25/2002  Aaron Huge  Initial creation

unit SizeMaster;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Grids, DBGrids, DataModule, ADODB, Mask,
  dbcgrids, DB;

type
  TSizeMaster_Form = class(TForm)
    ManagementButtons_Panel: TPanel;
    Insert_Button: TButton;
    Update_Button: TButton;
    Search_Button: TButton;
    Clear_Button: TButton;
    Close_Button: TButton;
    Delete_Button: TButton;
    SizeMaster_Label: TLabel;
    SizeMaster_Panel: TPanel;
    SizeCode_Label: TLabel;
    SizeCode_Edit: TEdit;
    DailyUsage_Label: TLabel;
    SizeName_Label: TLabel;
    SizeName_Edit: TEdit;
    SafetyDays_Label: TLabel;
    DailyUsage_MaskEdit: TMaskEdit;
    SafetyDays_MaskEdit: TMaskEdit;
    SizeMaster_DBGrid: TDBGrid;
    Size_DataSource: TDataSource;
    procedure Insert_ButtonClick(Sender: TObject);
    procedure Update_ButtonClick(Sender: TObject);
    procedure Delete_ButtonClick(Sender: TObject);
    procedure Search_ButtonClick(Sender: TObject);
    procedure Clear_ButtonClick(Sender: TObject);
    procedure SizeCode_EditChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure TextChange(Sender: TObject);
    procedure SizeMaster_DBGridMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure SizeMaster_DBGridKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Size_DataSourceDataChange(Sender: TObject; Field: TField);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
    procedure SetDetailBoxes;
    procedure HoldDetails(fFromGrid: Boolean);
    function SearchGrid(fSizeCode: String): Boolean;
  public
    { Public declarations }
    function Execute: boolean;
  end;

var
  SizeMaster_Form: TSizeMaster_Form;

implementation

{$R *.dfm}

function TSizeMaster_Form.Execute: boolean;
Begin
  Result:= True;
  try
    try
      ShowModal;
    except
      On E:Exception do
        showMessage('Unable to generate Size Master screen.'
          + #13 + 'ERROR:'
          + #13 + E.Message);
    end;
  finally
    If ModalResult = mrCancel Then
      Result:= False;
  end;
end;    //Execute


procedure TSizeMaster_Form.Insert_ButtonClick(Sender: TObject);
begin
  HoldDetails(False);
  If Not Data_Module.InsertSizeInfo Then
    MessageDlg('Unable to INSERT ' + Data_Module.SizeCode + #13 +
              'It already exists in the database.', mtInformation, [mbOk], 0);

  Data_Module.GetSizeInfo;
  Data_Module.Inv_DataSet.Locate('Size Code',SizeCode_Edit.Text,[]);
  SetDetailBoxes;
  SizeCode_Edit.SelectAll;
  SizeCode_Edit.SetFocus;
end;        //Insert_ButtonClick

procedure TSizeMaster_Form.Update_ButtonClick(Sender: TObject);
var
  size:string;
begin
  HoldDetails(False);
  size:=SizeCode_Edit.Text;
  Data_Module.UpdateSizeInfo;
  Data_Module.GetSizeInfo;
  Data_Module.Inv_DataSet.Locate('Size Code',size,[]);
  SetDetailBoxes;
end;        //Update_ButtonClick

procedure TSizeMaster_Form.Delete_ButtonClick(Sender: TObject);
begin
  HoldDetails(False);
  If MessageDlg('Are you sure you wish to delete' + #13 +
                 Data_Module.SizeName + ' (' + Data_Module.SizeCode + ') from the database?',
                 mtWarning, [mbYes, mbNo], 0) = mrYes Then
  Begin
    Data_Module.DeleteSizeInfo;
    SizeCode_Edit.Text := '';
    Data_Module.GetSizeInfo;
    SearchGrid(Data_Module.SizeCode);
    Data_Module.Inv_DataSet.Filtered := False;
  End;

  SizeCode_Edit.SetFocus;
end;        //Delete_ButtonClick

procedure TSizeMaster_Form.Search_ButtonClick(Sender: TObject);
var
  fFound: Boolean;
Begin

  //fFound := False;
  fFound := SearchGrid(SizeCode_Edit.Text);
  If fFound Then
  Begin
    SetDetailBoxes;
  End
  Else
  Begin
    ShowMessage('No matches were found for your query.');
  End;

  SizeCode_Edit.SelectAll;
  SizeCode_Edit.SetFocus;
end;        //Search_ButtonClick

procedure TSizeMaster_Form.Clear_ButtonClick(Sender: TObject);
begin
  Data_Module.Inv_DataSet.Filtered := False;
  SizeCode_Edit.Text := '';
  SizeCode_Edit.SetFocus;
end;        //Clear_ButtonClick

procedure TSizeMaster_Form.SizeCode_EditChange(Sender: TObject);
begin
  If Length(SizeCode_Edit.Text) <= 1 Then
    Data_Module.ClearControls(SizeMaster_Panel);

end;

function TSizeMaster_Form.SearchGrid(fSizeCode: String): Boolean;
begin
  Result := False;
  Data_Module.ClearControls(SizeMaster_Panel);

  try
    With Data_Module.Inv_DataSet Do
    Begin
      Filtered := False;
      Filter := '[Size Code] = ' + QuotedStr(fSizeCode);
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

procedure TSizeMaster_Form.SetDetailBoxes;
Begin
  With Data_Module do
  Begin
    SizeCode_Edit.Text := SizeCode;
    SizeName_Edit.Text := SizeName;
    DailyUsage_MaskEdit.Text := IntToStr(DailyUsage);
    SafetyDays_MaskEdit.Text := IntToStr(SafetyDays);
  End;      //With
End;          //SetDetailBoxes

procedure TSizeMaster_Form.HoldDetails(fFromGrid: Boolean);
Begin
  If fFromGrid Then
  Begin
    with SizeMaster_DBGrid.DataSource.DataSet do
    begin
      Data_Module.SizeCode        := Fields[0].AsString;
      Data_Module.SizeName        := Fields[1].AsString;
      Data_Module.DailyUsage      := Fields[2].AsInteger;
      Data_Module.SafetyDays      := Fields[3].AsInteger;
      Data_Module.RecordID        := Fields[4].AsInteger;
      Fields[4].Visible:=FALSE;
    end;
  End
  Else
  Begin
    With Data_Module do
    Begin
      SizeCode    := SizeCode_Edit.Text;
      SizeName    := SizeName_Edit.Text;
      DailyUsage  := StrToInt(Trim(DailyUsage_MaskEdit.Text));
      SafetyDays  := StrToInt(Trim(SafetyDays_MaskEdit.Text));
    End;      //With
  End;
End;          //HoldDetails

procedure TSizeMaster_Form.FormCreate(Sender: TObject);
begin
  Data_Module.Inv_DataSet.Filter:='';
  Data_Module.Inv_DataSet.Filtered:=FALSE;
  Data_Module.GetSizeInfo;
  Size_DataSource.DataSet:=Data_Module.Inv_DataSet;
  Data_Module.Inv_DataSet.Filtered := False;
  SizeCode_Edit.Text := '';
end;

procedure TSizeMaster_Form.TextChange(Sender: TObject);
begin
  If Sender.ClassName = 'TMaskEdit' Then
    If Length(Trim(TMaskEdit(Sender).Text)) < 1 Then
      TMaskEdit(Sender).Text := '0';

end;

procedure TSizeMaster_Form.SizeMaster_DBGridMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  HoldDetails(True);
  SetDetailBoxes;
end;


procedure TSizeMaster_Form.SizeMaster_DBGridKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  HoldDetails(True);
  SetDetailBoxes;
end;

procedure TSizeMaster_Form.Size_DataSourceDataChange(Sender: TObject;
  Field: TField);
begin
  HoldDetails(True);
  SetDetailBoxes;
end;

procedure TSizeMaster_Form.FormShow(Sender: TObject);
begin
  SetDetailBoxes;
  SizeCode_Edit.SetFocus;
end;

end.
