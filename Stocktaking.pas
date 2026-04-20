//***********************************************************
//  InventorySystem
//
//  Copyright (c) 2002 Failproof Manufacturing Systems
//
//***********************************************************
//
//  10/25/2002  Aaron Huge  Initial creation

unit Stocktaking;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Grids, DBGrids, DataModule, ComCtrls, Mask, ADODB,
  DB, AdvCombo, ColCombo, NUMMIColumnComboBox;

type
  TStocktaking_Form = class(TForm)
    RecProdReject_Label: TLabel;
    Stocktaking_DBGrid: TDBGrid;
    ManagementButtons_Panel: TPanel;
    Insert_Button: TButton;
    Update_Button: TButton;
    Search_Button: TButton;
    Clear_Button: TButton;
    Close_Button: TButton;
    Delete_Button: TButton;
    Stocktaking_Panel: TPanel;
    Reason_Label: TLabel;
    Reason_Memo: TMemo;
    Qty_Label: TLabel;
    PartsCode_Label: TLabel;
    PartsCode_ComboBox: TComboBox;
    SupplierCode_Label: TLabel;
    Date_Label: TLabel;
    Edit_DateTimePicker: TDateTimePicker;
    Qty_MaskEdit: TMaskEdit;
    Stocktaking_DataSource: TDataSource;
    Supplier_NUMMIColumnComboBox: TNUMMIColumnComboBox;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure MaskEditExit(Sender: TObject);
    procedure PartsCode_ComboBoxChange(Sender: TObject);
    procedure Insert_ButtonClick(Sender: TObject);
    procedure Update_ButtonClick(Sender: TObject);
    procedure Delete_ButtonClick(Sender: TObject);
    procedure Search_ButtonClick(Sender: TObject);
    procedure Clear_ButtonClick(Sender: TObject);
    procedure TextChange(Sender: TObject);
    procedure Stocktaking_DBGridKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Stocktaking_DBGridMouseUp(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure Stocktaking_DataSourceDataChange(Sender: TObject;
      Field: TField);
    procedure FormShow(Sender: TObject);
    procedure Supplier_NUMMIColumnComboBoxChange(Sender: TObject);
  private
    { Private declarations }
    fNoChange:Boolean;
    procedure SetDetailBoxes;
    function HoldDetails(fFromGrid: Boolean): String;
    function SearchGrid(fSupCode: String; fPartNum: String): Boolean;
  public
    { Public declarations }
    function Execute: boolean;
  end;

var
  Stocktaking_Form: TStocktaking_Form;

implementation

{$R *.dfm}

function TStocktaking_Form.Execute: boolean;
Begin
  Result:= True;
  try
    try
      ShowModal;
    except
      On E:Exception do
        showMessage('Unable to generate Stocktaking screen.'
          + #13 + 'ERROR:'
          + #13 + E.Message);
    end;
  finally
    If ModalResult = mrCancel Then
      Result:= False;
  end;
end;        //Execute

procedure TStocktaking_Form.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Data_Module.Inv_Dataset.Filtered := False;
end;

procedure TStocktaking_Form.FormCreate(Sender: TObject);
begin
  With Data_Module do
  Begin
    fNochange:=True;
    SetTodaysDate(Stocktaking_Panel);
    Inv_DataSet.Filter:='';
    Inv_DataSet.Filtered:=FALSE;
    GetStocktakingInfo;
    Stocktaking_DataSource.DataSet:=Inv_DataSet;
    ClearControls(Stocktaking_Panel);
    SelectMultiField('INV_SUPPLIER_MST', 'VC_SUPPLIER_CODE, VC_SUPPLIER_NAME', Supplier_NUMMIColumnComboBox);
    Inv_Dataset.Filtered := False;
    fNochange:=False;
  End;      //With
end;

procedure TStocktaking_Form.MaskEditExit(Sender: TObject);
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

function TStocktaking_Form.HoldDetails(fFromGrid: Boolean): String;
var
  temp,fErrMsg: String;
Begin
  fErrMsg := '';

  If fFromGrid Then
  Begin
    with Stocktaking_DBGrid.DataSource.DataSet do
    begin
      Data_Module.EditDate := Copy(Fields[0].AsString, 1, 4) + Copy(Fields[0].AsString, 6, 2) + Copy(Fields[0].AsString, 9, 2);
      Data_Module.SupplierCode := Fields[1].AsString;
      Data_Module.PartNum := Fields[2].AsString;
      Data_Module.PartName := Fields[3].AsString;
      Data_Module.Quantity := Fields[4].AsInteger;
      Data_Module.Comments := Fields[5].AsString;
      Data_Module.RecordID := Fields[6].AsInteger;
      Fields[6].Visible:=FALSE;

    End;      //With
  End
  Else
  Begin
    With Data_Module do
    Begin

      temp:=FormatDateTime('yyyymmdd', Edit_DateTimePicker.Date);
      EditDate := FormatDateTime('yyyymmdd', Edit_DateTimePicker.Date);

      If Length(trim(PartsCode_ComboBox.Text)) <> 12 Then
        fErrMsg := #13 + 'Invalid Part Code'
      Else
        PartNum := PartsCode_ComboBox.Text;

      Quantity := StrToInt(Trim(Qty_MaskEdit.Text));

      Comments := Reason_Memo.Text;
    End;      //With
  End;
  If Not (fErrMsg = '') Then
    fErrMsg := 'The following fields need to be corrected:' + fErrMsg;
  result := fErrMsg;
End;          //HoldDetails

procedure TStocktaking_Form.SetDetailBoxes;
var
  fTempDate: String;
Begin
  With Data_Module do
  Begin
    if EditDate <> '' then
    begin
      fTempDate := EditDate;
      fTempDate := Copy(fTempDate, 5, 2) + '/' + Copy(fTempDate, 7, 2) + '/' + Copy(fTempDate, 1, 4);
      Edit_DateTimePicker.Date := StrToDate(fTempDate);
    end
    else
      Edit_DateTimePicker.Date := now;

    SearchMultiCombo(Supplier_NUMMIColumnComboBox, SupplierCode);
    Supplier_NUMMIColumnComboBoxChange(self);
    SearchCombo(PartsCode_ComboBox, PartNum);
    Qty_MaskEdit.Text := IntToStr(Quantity);
    Reason_Memo.Text := Comments;
  End;      //With
End;          //SetDetailBoxes

function TStocktaking_Form.SearchGrid(fSupCode: String; fPartNum: String): Boolean;
begin
  Data_Module.ClearControls(Stocktaking_Panel);
  Result := False;
  try
    Data_Module.Inv_DataSet.Filter:='';
    Data_Module.Inv_DataSet.Filtered := False;

    If fSupCode <> ''  Then
      Data_Module.Inv_DataSet.Filter := '[Supplier Code] = ' + QuotedStr(fSupCode);

    If fPartNum <> '' Then
      If Length(Data_Module.Inv_DataSet.Filter) > 0 Then
        Data_Module.Inv_DataSet.Filter := Data_Module.Inv_DataSet.Filter + ' AND [Parts Code]= ' + QuotedStr(fPartNum)
      Else
        Data_Module.Inv_DataSet.Filter := '[Parts Code] = ' + QuotedStr(fPartNum);

    Data_Module.Inv_DataSet.Filtered := True;

    If Data_Module.Inv_DataSet.RecordCount > 0 Then
    Begin
      HoldDetails(True);
      Result := True;
    End
    else
    begin
      Data_Module.Inv_Dataset.Filter:='';
      Data_Module.Inv_Dataset.Filtered := False;
      Data_Module.ClearControls(Stocktaking_Panel);
    end;
  except
    on e:exception do
      ShowMessage('Error in Search' + #13 + e.Message);
  end;      //try...except
End;        //SearchGrid

procedure TStocktaking_Form.TextChange(Sender: TObject);
Begin
  If Sender.ClassName = 'TMaskEdit' Then
    If Length(Trim(TMaskEdit(Sender).Text)) < 1 Then
      TMaskEdit(Sender).Text := '0';
End;          //TextChange

procedure TStocktaking_Form.PartsCode_ComboBoxChange(Sender: TObject);
begin
  Qty_MaskEdit.SetFocus;
end;

procedure TStocktaking_Form.Insert_ButtonClick(Sender: TObject);
begin
  HoldDetails(False);
  with Data_Module do
  begin
    If Not Data_Module.InsertStocktakingInfo Then
      MessageDlg('Unable to INSERT ' + Data_Module.PartNum +'('+ Data_Module.SupplierCode +')', mtInformation, [mbOk], 0);
    GetStocktakingInfo;
  end;
  SetDetailBoxes;
  Supplier_NUMMIColumnComboBox.SetFocus;
end;        //Insert_ButtonClick

procedure TStocktaking_Form.Update_ButtonClick(Sender: TObject);
var
  ID:integer;
begin
  HoldDetails(False);

  with Data_Module do
  begin
    ID:=RecordID;
    UpdateStocktakingInfo;
    GetStocktakingInfo;
    Inv_Dataset.Filtered := False;
    Inv_DataSet.Locate('RecordID',ID,[])
  end;
  SetDetailBoxes;
  Supplier_NUMMIColumnComboBox.SetFocus;
end;          //Update_ButtonClick

procedure TStocktaking_Form.Delete_ButtonClick(Sender: TObject);
begin
  HoldDetails(False);
  If MessageDlg('Are you sure you wish to delete' + #13 +
                 'part code ' + Data_Module.PartNum + #13 +
                 ' from the database?',
                 mtWarning, [mbYes, mbNo], 0) = mrYes Then
  Begin
    Data_Module.DeleteStocktakingInfo;
    Data_Module.GetStocktakingInfo;
  End;
  Data_module.ClearControls(Stocktaking_Panel);
  Supplier_NUMMIColumnComboBox.SetFocus;
end;      //Delete_ButtonClick

procedure TStocktaking_Form.Search_ButtonClick(Sender: TObject);
Begin
  if (Supplier_NUMMIColumnComboBox.Text<> '') and (PartsCode_ComboBox.Text<>'') then
    SearchGrid(Supplier_NUMMIColumnComboBox.ColumnItems[Supplier_NUMMIColumnComboBox.ItemIndex,0],PartsCode_ComboBox.Text);

  Supplier_NUMMIColumnComboBox.SetFocus;
end;      //Search_ButtonClick

procedure TStocktaking_Form.Clear_ButtonClick(Sender: TObject);
begin
  with Data_Module do
  begin
    SetTodaysDate(Stocktaking_Panel);
    Inv_Dataset.Filtered := False;
    ClearControls(Stocktaking_Panel);
  End;    //If
  Supplier_NUMMIColumnComboBox.SetFocus;
end;      //Clear_ButtonClick

procedure TStocktaking_Form.Stocktaking_DBGridKeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  HoldDetails(True);
  SetDetailBoxes;
end;

procedure TStocktaking_Form.Stocktaking_DBGridMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  HoldDetails(True);
  SetDetailBoxes;
end;

procedure TStocktaking_Form.Stocktaking_DataSourceDataChange(
  Sender: TObject; Field: TField);
begin
  if not fNoChange then
  begin
    HoldDetails(True);
    SetDetailBoxes;
  end;
end;

procedure TStocktaking_Form.FormShow(Sender: TObject);
begin
  HoldDetails(True);
  SetDetailBoxes;
  Supplier_NUMMIColumnComboBox.SetFocus;
end;

procedure TStocktaking_Form.Supplier_NUMMIColumnComboBoxChange(
  Sender: TObject);
begin
  Data_Module.SelectDependantSingleField( 'SELECT_DependantPartNumber_Supplier',
                                          '@SupplierCode',
                                          'VC_PART_NUMBER',
                                          Supplier_NUMMIColumnComboBox.ColumnItems[Supplier_NUMMIColumnComboBox.ItemIndex,0],
                                          PartsCode_ComboBox);

  PartsCode_ComboBox.SetFocus;

end;

end.
