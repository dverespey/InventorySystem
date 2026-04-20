//***********************************************************
//  InventorySystem
//
//  Copyright (c) 2002 Failproof Manufacturing Systems
//
//***********************************************************
//
//  10/25/2002  Aaron Huge  Initial creation

unit InvMgmt;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, DBGrids, ExtCtrls, DataModule, InvMgmtQReport,
  DB, AdvCombo, ColCombo, NUMMIColumnComboBox;

type
  TInvMgmt_Form = class(TForm)
    InvMgmt_Label: TLabel;
    SearchKey_GroupBox: TGroupBox;
    SupplierCode_Label: TLabel;
    PartsCode_Label: TLabel;
    TireWheel_Label: TLabel;
    Label1: TLabel;
    InvMgmt_DBGrid: TDBGrid;
    ManagementButtons_Panel: TPanel;
    Print_Button: TButton;
    Search_Button: TButton;
    Clear_Button: TButton;
    Close_Button: TButton;
    Inventory_DataSource: TDataSource;
    Supplier_NUMMIColumnComboBox: TNUMMIColumnComboBox;
    PartNum_ComboBox: TComboBox;
    Line_ComboBox: TComboBox;
    PartType_ComboBox: TComboBox;
    PrintExcel_Button: TButton;
    procedure InvMgmt_DBGridMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure InvMgmt_DBGridKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormCreate(Sender: TObject);
    procedure Print_ButtonClick(Sender: TObject);
    procedure Clear_ButtonClick(Sender: TObject);
    procedure Search_ButtonClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure Supplier_NUMMIColumnComboBoxChange(Sender: TObject);
    procedure PartNum_ComboBoxChange(Sender: TObject);
    procedure Line_ComboBoxChange(Sender: TObject);
    procedure PartType_ComboBoxChange(Sender: TObject);
  private
    { Private declarations }
    //function SearchGrid(fSupCode: String; fPartNum: String; fSizeCode: String; fKanban: String): Boolean;
    procedure HoldDetails(fFromGrid: Boolean);
    procedure SetDetailBoxes;
    procedure SetCombos;
  public
    { Public declarations }
    function Execute: boolean;
  end;

var
  InvMgmt_Form: TInvMgmt_Form;

implementation
                                      
{$R *.dfm}
function TInvMgmt_Form.Execute: boolean;
Begin
  Result:= True;
  try
    try
      SetCombos;
      ShowModal;
    except
      On E:Exception do
        showMessage('Unable to generate Inventory Management screen.'
          + #13 + 'ERROR:'
          + #13 + E.Message);
    end;
  finally
    If ModalResult = mrCancel Then
      Result:= False;
  end;
end;        //Execute

{function TInvMgmt_Form.SearchGrid(fSupCode: String; fPartNum: String; fSizeCode: String; fKanban: String): Boolean;
begin
  Result := False;
  try
    With Data_Module.Inv_DataSet Do
    Begin
      Filtered := False;
      Filter := '';

      If Length(fSupCode) <> 0 Then
        Filter := '[Supplier Code] = ' + QuotedStr(fSupCode);

      If Length(fPartNum) <> 0 Then
        If Length(Filter) > 0 Then
          Filter := Filter + ' AND [Parts Code] = ' + QuotedStr(fPartNum)
        Else
          Filter := '[Parts Code] = ' + QuotedStr(fPartNum);

      If Length(fSizeCode) <> 0 Then
        If Length(Filter) > 0 Then
          Filter := Filter + ' AND [Size Code] = ' + QuotedStr(fSizeCode)
        Else
          Filter := '[Size Code] = ' + QuotedStr(fSizeCode);

      If Length(fKanban) <> 0 Then
        If Length(Filter) > 0 Then
          Filter := Filter + ' AND [KANBAN] = ' + QuotedStr(fKanban)
        Else
          Filter := '[KANBAN] = ' + QuotedStr(fKanban);


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
end;        //SearchGrid
 }
procedure TInvMgmt_Form.HoldDetails(fFromGrid: Boolean);
var
  fErrMsg: String;
Begin
  fErrMsg := '';

  If fFromGrid Then
  Begin
    with InvMgmt_DBGrid.DataSource.DataSet do
    begin
      Data_Module.SupplierCode := Fields[0].AsString;
      Data_Module.PartNum := Fields[1].AsString;
      Data_Module.PartType := Fields[5].AsString;
      Data_Module.LineName := Fields[6].AsString;
      Data_Module.SizeCode := Fields[7].AsString;
      Data_Module.Kanban := Fields[8].AsString;
      Data_Module.Quantity := Fields[26].AsInteger;
    End;      //With
  End
  Else
  Begin
    With Data_Module do
    Begin
      {If Length(SupplierCode_Edit.Text) < 5 Then
        fErrMsg := #13 + 'Invalid Supplier Code'
      Else
        SupplierCode := SupplierCode_Edit.Text;

      If Length(trim(PartsCode_Edit.Text)) < 1 Then
        fErrMsg := #13 + 'Invalid Part Code'
      Else
        PartNum := PartsCode_Edit.Text;

      SizeCode := SizeCode_Edit.Text;
      Kanban := Kanban_Edit.Text;}

    End;      //With
  End;
End;          //HoldDetails

procedure TInvMgmt_Form.SetDetailBoxes;
Begin
  With Data_Module do
  Begin
    {SupplierCode_Edit.Text := SupplierCode;
    PartsCode_Edit.Text := PartNum;
    SizeCode_Edit.Text := SizeCode;
    PartType_Edit.Text := PartType;
    LineName_Edit.Text := LineName;
    Kanban_Edit.Text := Kanban;
    Qty_Edit.Text := IntToStr(Quantity);}
  End;      //With
End;          //SetDetailBoxes

procedure TInvMgmt_Form.InvMgmt_DBGridMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  HoldDetails(True);
  SetDetailBoxes;
end;

procedure TInvMgmt_Form.InvMgmt_DBGridKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  HoldDetails(True);
  SetDetailBoxes;
end;

procedure TInvMgmt_Form.FormCreate(Sender: TObject);
begin
  Data_Module.Inv_DataSet.Filter:='';
  Data_Module.Inv_DataSet.Filtered:=FALSE;
  Data_Module.GetInventoryInfo;
  Inventory_DataSource.DataSet:=Data_Module.Inv_DataSet;
  Data_Module.JustifyColumns(InvMgmt_DBGrid);
  Data_Module.Grid_ClientDataSet.Filtered := False;
  Data_Module.ClearControls(SearchKey_GroupBox);
end;

procedure TInvMgmt_Form.SetCombos;
begin
  with Data_Module do
  begin
    SelectMultiField('INV_SUPPLIER_MST', 'VC_SUPPLIER_CODE, VC_SUPPLIER_NAME', Supplier_NUMMIColumnComboBox);
    SelectSingleField('INV_PARTS_STOCK_MST', 'VC_PART_NUMBER', PartNum_ComboBox);
    SelectSingleField('INV_PART_TYPE_MST', 'VC_PART_TYPE', PartType_ComboBox);
    SelectSingleFieldALC('LINE', 'LineName', Line_ComboBox);
  end;
end;

procedure TInvMgmt_Form.Print_ButtonClick(Sender: TObject);
begin
  Hide;
  InvMgmtQReport_Form:=TInvMgmtQReport_Form.Create(self);
  InvMgmtQReport_Form.InvMgmt_QuickRep.Preview;
  InvMgmtQReport_Form.Free;
  Show;
end;     //Print_ButtonClick

procedure TInvMgmt_Form.Clear_ButtonClick(Sender: TObject);
begin
  With Data_Module do
  Begin
    Grid_ClientDataSet.Filtered := False;
    GetInventoryInfo;
    ClearControls(SearchKey_GroupBox);
    HoldDetails(True);
    SetDetailBoxes;
  End;      //With Data_Module
  //SupplierCode_Edit.SetFocus;
end;

procedure TInvMgmt_Form.Search_ButtonClick(Sender: TObject);
begin

  {If (Trim(SupplierCode_Edit.Text ) = '') And (Trim(PartsCode_Edit.Text) = '')
        And  (Trim(SizeCode_Edit.Text ) = '') And (Trim(Kanban_Edit.Text) = '') then
    ShowMessage('Please enter some data in the' + #13 +
                'search key above and try again.')
  Else
  Begin
    //fFound := False;
    fFound := SearchGrid(Trim(SupplierCode_Edit.Text), Trim(PartsCode_Edit.Text), Trim(SizeCode_Edit.Text), Trim(Kanban_Edit.Text));
    If fFound Then
    Begin
      SetDetailBoxes;
    End
    Else
    Begin
      ShowMessage('No matches were found for your query.');
    End;
  End;

  SupplierCode_Edit.SelectAll;
  SupplierCode_Edit.SetFocus;}
end;

procedure TInvMgmt_Form.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Data_Module.Grid_ClientDataSet.Filtered := False;
end;

procedure TInvMgmt_Form.FormShow(Sender: TObject);
begin
  //SupplierCode_Edit.SetFocus;
  HoldDetails(True);
  SetDetailBoxes;
end;

procedure TInvMgmt_Form.Supplier_NUMMIColumnComboBoxChange(
  Sender: TObject);
begin
  Data_Module.SelectDependantSingleField('SELECT_PartsSupplier', '@VC_SUPPLIER_CODE', 'VC_PART_NUMBER', Supplier_NUMMIColumnComboBox.ColumnItems[Supplier_NUMMIColumnComboBox.ItemIndex,0], PartNum_ComboBox);
end;

procedure TInvMgmt_Form.PartNum_ComboBoxChange(Sender: TObject);
begin
  //Clear other items and update list
end;

procedure TInvMgmt_Form.Line_ComboBoxChange(Sender: TObject);
begin
  //Clear other items and update list
end;

procedure TInvMgmt_Form.PartType_ComboBoxChange(Sender: TObject);
begin
  //Clear other items and update list
end;

end.

