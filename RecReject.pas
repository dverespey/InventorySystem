//***********************************************************
//  InventorySystem
//
//  Copyright (c) 2002 Failproof Manufacturing Systems
//
//***********************************************************
//
//  10/25/2002  Aaron Huge  Initial creation

unit RecReject;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Grids, DBGrids, DataModule, Mask, ComCtrls,
  DB, AdvCombo, ColCombo, NUMMIColumnComboBox;

type
  TRecRej_Form = class(TForm)
    RecProdReject_Label: TLabel;
    RecProdRej_DBGrid: TDBGrid;
    ManagementButtons_Panel: TPanel;
    Insert_Button: TButton;
    Update_Button: TButton;
    Search_Button: TButton;
    Clear_Button: TButton;
    Close_Button: TButton;
    Delete_Button: TButton;
    RecProdRej_Panel: TPanel;
    Reason_Memo: TMemo;
    Reason_Label: TLabel;
    RejQty_Label: TLabel;
    PartsCode_Label: TLabel;
    PartsCode_ComboBox: TComboBox;
    SupplierCode_Label: TLabel;
    DiscnDiv_RadioGroup: TRadioGroup;
    DiscernDiv_Label: TLabel;
    Date_Label: TLabel;
    RejQty_MaskEdit: TMaskEdit;
    Edit_DateTimePicker: TDateTimePicker;
    RecRej_DataSource: TDataSource;
    Supplier_NUMMIColumnComboBox: TNUMMIColumnComboBox;
    procedure RecProdRej_DBGridKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure RecProdRej_DBGridMouseUp(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure TextChange(Sender: TObject);
    procedure MaskEditExit(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure Insert_ButtonClick(Sender: TObject);
    procedure Update_ButtonClick(Sender: TObject);
    procedure Delete_ButtonClick(Sender: TObject);
    procedure Search_ButtonClick(Sender: TObject);
    procedure Clear_ButtonClick(Sender: TObject);
    procedure RecRej_DataSourceDataChange(Sender: TObject; Field: TField);
    procedure FormShow(Sender: TObject);
    procedure Supplier_NUMMIColumnComboBoxChange(Sender: TObject);
    procedure PartsCode_ComboBoxChange(Sender: TObject);
  private
    { Private declarations }
    fNoChange:boolean;

    function HoldDetails(fFromGrid: Boolean): String;
    procedure SetDetailBoxes;
    function SearchGrid(fSupCode: String; fPartNum: String): Boolean;
  public
    { Public declarations }
    function Execute: boolean;
  end;

var
  RecRej_Form: TRecRej_Form;

implementation

{$R *.dfm}

function TRecRej_Form.Execute: boolean;
Begin
  Result:= True;
  try
    try
      If Data_Module.Division = '' Then
      Begin
        Caption := 'Receiving Reject';
        DiscnDiv_RadioGroup.ItemIndex := 0;
      End
      Else
      Begin
        Caption := 'Production Reject';
        DiscnDiv_RadioGroup.ItemIndex := StrToInt(Data_Module.Division) - 1;
      End;
      ShowModal;
    except
      On E:Exception do
        showMessage('Unable to generate Receiving Reject screen.'
          + #13 + 'ERROR:'
          + #13 + E.Message);
    end;
  finally
    If ModalResult = mrCancel Then
      Result:= False;
  end;
end;        //Execute

procedure TRecRej_Form.RecProdRej_DBGridKeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  HoldDetails(True);
  SetDetailBoxes;
end;

procedure TRecRej_Form.RecProdRej_DBGridMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  HoldDetails(True);
  SetDetailBoxes;
end;

procedure TRecRej_Form.TextChange(Sender: TObject);
Begin
  If Sender.ClassName = 'TMaskEdit' Then
    If Length(Trim(TMaskEdit(Sender).Text)) < 1 Then
      TMaskEdit(Sender).Text := '0';
End;          //TextChange

procedure TRecRej_Form.MaskEditExit(Sender: TObject);
var
  fPos: Integer;
  fTempValue: String;
begin
  If Sender.ClassName = 'TMaskEdit' Then
  Begin
    fTempValue := Trim(TMaskEdit(Sender).Text);
    //fPos := -1;
    fPos := Pos(' ', fTempValue);
    While fPos <> 0 do
    Begin
      fTempValue := Copy(fTempValue, 1, fPos - 1) + Copy(fTempValue, fPos + 1, Length(fTempValue));
      fPos := Pos(' ', fTempValue);
    End;
    TMaskEdit(Sender).Text := fTempValue;
  End;
end;      //MaskEditExit

procedure TRecRej_Form.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Data_Module.Inv_Dataset.Filtered := False;
end;

procedure TRecRej_Form.FormCreate(Sender: TObject);
begin
  With Data_Module do
  Begin
    fNoChange:=True;
    SetTodaysDate(RecProdRej_Panel);
    Inv_DataSet.Filter:='';
    Inv_DataSet.Filtered:=FALSE;
    GetRecProdRejInfo;
    RecRej_DataSource.DataSet:=Inv_DataSet;
    ClearControls(RecProdRej_Panel);
    SelectMultiField('INV_SUPPLIER_MST', 'VC_SUPPLIER_CODE, VC_SUPPLIER_NAME', Supplier_NUMMIColumnComboBox);
    Inv_Dataset.Filtered := False;
    fNoChange:=False;
  End;      //With
end;

function TRecRej_Form.HoldDetails(fFromGrid: Boolean): String;
var
  fErrMsg: String;
Begin
  fErrMsg := '';

  If fFromGrid Then
  Begin
    with RecProdRej_DBGrid.DataSource.DataSet do
    begin
      Data_Module.EditDate := Copy(Fields[0].AsString, 1, 4) + Copy(Fields[0].AsString, 6, 2) + Copy(Fields[0].AsString, 9, 2);
      If Fields[1].AsString = 'Receiving' Then
        Data_Module.Division := '1'
      Else If Fields[1].AsString = 'Assembler' Then
        Data_Module.Division := '2'
      Else If Fields[1].AsString = 'Plant' Then
        Data_Module.Division := '3'
      Else
        Data_Module.Division := '1';

      Data_Module.SupplierCode := Fields[2].AsString;
      Data_Module.PartNum := Fields[3].AsString;
      Data_Module.PartName := Fields[4].AsString;
      Data_Module.Quantity := Fields[5].AsInteger;
      Data_Module.Comments := Fields[6].AsString;
      Data_Module.RecordID := Fields[7].AsInteger;
      Fields[7].Visible:=FALSE;
    End;      //With
  End
  Else
  Begin
    With Data_Module do
    Begin
      EditDate := FormatDateTime('yyyymmdd', Edit_DateTimePicker.Date);
      case DiscnDiv_RadioGroup.ItemIndex of
        0: Division := '1';
        1: Division := '2';
        2: Division := '3';
      else
        Division := '1';
      end;

      if PartsCode_ComboBox.Text <> '' then
        PartNum := PartsCode_ComboBox.Text
      else
        fErrMsg:='Part Number must not be blank';
        
      Quantity := StrToInt(Trim(RejQty_MaskEdit.Text));

      Comments := Reason_Memo.Text;
    End;      //With
  End;
  If Not (fErrMsg = '') Then
    fErrMsg := 'The following fields need to be corrected:' + fErrMsg;
  result := fErrMsg;
End;          //HoldDetails

procedure TRecRej_Form.SetDetailBoxes;
var
  fTempDate: String;
  td:TDateTime;
Begin
  try
    With Data_Module do
    Begin
      fTempDate := EditDate;
      fTempDate := Copy(fTempDate, 5, 2) + '/' + Copy(fTempDate, 7, 2) + '/' + Copy(fTempDate, 1, 4);
      if TryStrToDate(fTempDate,td) then
        Edit_DateTimePicker.Date:=td;
      If Division = '' Then
        DiscnDiv_RadioGroup.ItemIndex := 0
      Else
        DiscnDiv_RadioGroup.ItemIndex := StrToInt(Division) - 1;

      SearchMultiCombo(Supplier_NUMMIColumnComboBox, SupplierCode);
      Supplier_NUMMIColumnComboBoxChange(self);

      SearchCombo(PartsCode_ComboBox, PartNum);

      RejQty_MaskEdit.Text := IntToStr(Quantity);

      Reason_Memo.Text := Comments;
    End;      //With
  except
  end;
End;          //SetDetailBoxes

function TRecRej_Form.SearchGrid(fSupCode: String; fPartNum: String): Boolean;
var
  fSupCodeBlank, fPartNumBlank: Boolean;
begin
//  data_Module.Inv_DataSet.FilterGroup
  Data_Module.ClearControls(RecProdRej_Panel);
  Result := False;
  fSupCodeBlank := False;
  fPartNumBlank := False;

  If Length(fSupCode) = 0 Then
    fSupCodeBlank := True;
  If Length(fPartNum) = 0 Then
    fPartNumBlank := True;

  try
    With Data_Module.Inv_DataSet Do
    Begin
      Filtered := False;
      If Not fSupCodeBlank Then
        Filter := '[Supplier Code] = ' + QuotedStr(fSupCode);

      If Not fPartNumBlank Then
        If Length(Filter) > 0 Then
          Filter := Filter + ' AND [Parts Code]= ' + QuotedStr(fPartNum)
        Else
          Filter := 'Parts = ' + QuotedStr(fPartNum);

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


procedure TRecRej_Form.Insert_ButtonClick(Sender: TObject);
begin
  HoldDetails(False);
  If Not Data_Module.InsertRecProdRejInfo Then
    ShowMessage('Unable to INSERT Reject Data');
  Data_Module.GetRecProdRejInfo;
  SetDetailBoxes;
  Supplier_NUMMIColumnComboBox.SetFocus;
end;          //Insert_ButtonClick

procedure TRecRej_Form.Update_ButtonClick(Sender: TObject);
var
  ID:integer;
begin
  HoldDetails(False);
  ID:=Data_Module.RecordID;
  Data_Module.UpdateRecProdRejInfo;
  Data_Module.GetRecProdRejInfo;
  Data_Module.Inv_DataSet.Locate('RecordID',ID,[]);
  SetDetailBoxes;
  Supplier_NUMMIColumnComboBox.SetFocus;
end;          //Update_ButtonClick

procedure TRecRej_Form.Delete_ButtonClick(Sender: TObject);
begin
  HoldDetails(False);
  If MessageDlg('Are you sure you wish to delete' + #13 +
                 'supplier code ' + Data_Module.SupplierCode + #13 +
                 'and part code ' + Data_Module.PartNum + #13 +
                 ' from the database?',
                 mtWarning, [mbYes, mbNo], 0) = mrYes Then
  Begin
    Data_Module.DeleteRecProdRejInfo;
    Data_module.ClearControls(RecProdRej_Panel);
    Data_Module.GetRecProdRejInfo;
  End;

  Supplier_NUMMIColumnComboBox.SetFocus;
end;          //Delete_ButtonClick

procedure TRecRej_Form.Search_ButtonClick(Sender: TObject);
Begin
  if (Supplier_NUMMIColumnComboBox.Text<> '') and (PartsCode_ComboBox.Text<>'') then
    SearchGrid(Supplier_NUMMIColumnComboBox.ColumnItems[Supplier_NUMMIColumnComboBox.ItemIndex,0],PartsCode_ComboBox.Text);
  Supplier_NUMMIColumnComboBox.SetFocus;
end;      //Search_ButtonClick     InsertRecProdRejInfo


procedure TRecRej_Form.Clear_ButtonClick(Sender: TObject);
begin
  Supplier_NUMMIColumnComboBox.ItemIndex:=0;
  Data_Module.ClearControls(RecProdRej_Panel);
  Data_Module.SetTodaysDate(RecProdRej_Panel);
  Supplier_NUMMIColumnComboBox.SetFocus;
end;      //Clear_ButtonClick

procedure TRecRej_Form.RecRej_DataSourceDataChange(Sender: TObject;
  Field: TField);
begin
  if not fNoChange then
  begin
    HoldDetails(True);
    SetDetailBoxes;
  end;
end;

procedure TRecRej_Form.FormShow(Sender: TObject);
begin
  Supplier_NUMMIColumnComboBox.SetFocus;
  DiscnDiv_RadioGroup.Items.Clear;
  DiscnDiv_RadioGroup.Items.Add('Receiving Reject');
  DiscnDiv_RadioGroup.Items.Add(Data_Module.fiAssemblerName.AsString+' Production Reject');
  DiscnDiv_RadioGroup.Items.Add(Data_Module.fiPlantName.AsString+' Production Reject');
  DiscnDiv_RadioGroup.ItemIndex:=0;
end;

procedure TRecRej_Form.Supplier_NUMMIColumnComboBoxChange(Sender: TObject);
begin

  Data_Module.SelectDependantSingleField( 'SELECT_DependantPartNumber_Supplier',
                                          '@SupplierCode',
                                          'VC_PART_NUMBER',
                                          Supplier_NUMMIColumnComboBox.ColumnItems[Supplier_NUMMIColumnComboBox.ItemIndex,0],
                                          PartsCode_ComboBox);

  PartsCode_ComboBox.SetFocus;
end;

procedure TRecRej_Form.PartsCode_ComboBoxChange(Sender: TObject);
begin
  RejQty_MaskEdit.SetFocus;
end;

end.
