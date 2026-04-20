//***********************************************************
//  InventorySystem
//
//  Copyright (c) 2002 Failproof Manufacturing Systems
//
//***********************************************************
//
//  10/25/2002  Aaron Huge  Initial creation
//  11/14/2002  Aaron Huge  Alter to use two tire part numbers and two ratios
//                          and two wheel part numbers and two wheel ratios

unit AssyRatioMaster;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, StdCtrls, ExtCtrls, DataModule, Mask, DB;

type
  TAssyRatioMaster_Form = class(TForm)
    Buttons_Panel: TPanel;
    Insert_Button: TButton;
    Update_Button: TButton;
    Search_Button: TButton;
    Clear_Button: TButton;
    Close_Button: TButton;
    Delete_Button: TButton;
    AssyRatioMaster_Label: TLabel;
    AssyRatioMaster_Panel: TPanel;
    AssyCode_Label: TLabel;
    TireQty_Label: TLabel;
    TirePartsCode1_Label: TLabel;
    TirePartsCode2_Label: TLabel;
    WheelPartsCode1_Label: TLabel;
    WheelPartsCode2_Label: TLabel;
    WheelQty_Label: TLabel;
    SpareTirePartsCode_Label: TLabel;
    SpareTirePartsCode_Edit: TEdit;
    WheelQty_RadioGroup: TRadioGroup;
    TireQty_RadioGroup: TRadioGroup;
    AssyName_Label: TLabel;
    TireRatio1_Label: TLabel;
    TireRatio2_Label: TLabel;
    WheelRatio1_Label: TLabel;
    TireRatio4_Label: TLabel;
    SpareTireQty_Label: TLabel;
    SpareWheelPartsCode_Label: TLabel;
    SpareWheelPartsCode_Edit: TEdit;
    SpareTireQty_RadioGroup: TRadioGroup;
    AssyName_Edit: TEdit;
    TireRatio1_MaskEdit: TMaskEdit;
    TireRatio2_MaskEdit: TMaskEdit;
    WheelRatio1_MaskEdit: TMaskEdit;
    WheelRatio2_MaskEdit: TMaskEdit;
    BroadcastCode_Label: TLabel;
    BroadcastCode_Edit: TEdit;
    TirePartNum1_ComboBox: TComboBox;
    TirePartNum2_ComboBox: TComboBox;
    WheelPartNum1_ComboBox: TComboBox;
    WheelPartNum2_ComboBox: TComboBox;
    ASSYRatioMaster_DBGrid: TDBGrid;
    AssyCode_ComboBox: TComboBox;
    AssyRatio_DataSource: TDataSource;
    Label1: TLabel;
    WheelPartNum3_ComboBox: TComboBox;
    Label2: TLabel;
    WheelRatio3_MaskEdit: TMaskEdit;
    Label3: TLabel;
    TirePartNum3_ComboBox: TComboBox;
    Label4: TLabel;
    TireRatio3_MaskEdit: TMaskEdit;
    TotalTireRatio_Edit: TEdit;
    Label5: TLabel;
    Label6: TLabel;
    TotalWheelRatio_Edit: TEdit;
    procedure Insert_ButtonClick(Sender: TObject);
    procedure Update_ButtonClick(Sender: TObject);
    procedure Delete_ButtonClick(Sender: TObject);
    procedure Search_ButtonClick(Sender: TObject);
    procedure Clear_ButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure ASSYRatioMaster_DBGridKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ASSYRatioMaster_DBGridMouseUp(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure AssyRatio_DataSourceDataChange(Sender: TObject;
      Field: TField);
    procedure FormShow(Sender: TObject);
    procedure TotalTireRatio_EditChange(Sender: TObject);
    procedure TirePartNum3_ComboBoxChange(Sender: TObject);
    procedure TirePartNum2_ComboBoxChange(Sender: TObject);
    procedure WheelPartNum3_ComboBoxChange(Sender: TObject);
    procedure WheelPartNum2_ComboBoxChange(Sender: TObject);
    procedure TireRatio3_MaskEditChange(Sender: TObject);
    procedure TireRatio2_MaskEditChange(Sender: TObject);
    procedure TireRatio1_MaskEditChange(Sender: TObject);
    procedure WheelRatio1_MaskEditChange(Sender: TObject);
    procedure WheelRatio2_MaskEditChange(Sender: TObject);
    procedure WheelRatio3_MaskEditChange(Sender: TObject);
    procedure TirePartNum1_ComboBoxChange(Sender: TObject);
    procedure WheelPartNum1_ComboBoxChange(Sender: TObject);
  private
    { Private declarations }
    function Verify:boolean;
    procedure SetDetailBoxes;
    procedure HoldDetails(fFromGrid: Boolean);
    function SearchGrid(BroadCode: String): Boolean;
    procedure GetParts(tablename, selectfieldname, fPartType, returnfieldname: String; ComboBox: TComboBox);
  public
    { Public declarations }
    function Execute: boolean;
  end;

var
  AssyRatioMaster_Form: TAssyRatioMaster_Form;

implementation

{$R *.dfm}

function TAssyRatioMaster_Form.Verify:boolean;
begin
  result:=TRUE;

  if StrToInt(TotalTireRatio_Edit.Text) <> 100 then
  begin
    ShowMessage('Tire ratio must total to 100');
    TireRatio1_MaskEdit.SetFocus;
    result:=FALSE;
  end;

  if StrToInt(TotalWheelRatio_Edit.Text) <> 100 then
  begin
    ShowMessage('Wheel ratio must total to 100');
    WheelRatio1_MaskEdit.SetFocus;
    result:=FALSE;
  end;
end;

function TAssyRatioMaster_Form.Execute: boolean;
Begin
  Result:= True;
  try
    try
      ShowModal;
    except
      On E:Exception do
        showMessage('Unable to generate ASSY Ratio Master screen.'
          + #13 + 'ERROR:'
          + #13 + E.Message);
    end;
  finally
    If ModalResult = mrCancel Then
      Result:= False;
  end;
end;    //Execute

procedure TAssyRatioMaster_Form.Insert_ButtonClick(Sender: TObject);
var
  broad:string;
begin
  if Verify then
  begin
    HoldDetails(False);
    with Data_Module do
    begin
      If Not InsertAssyRatioInfo Then
        MessageDlg('Unable to INSERT ' + Data_Module.AssyCode, mtInformation, [mbOk], 0)
      else
      begin
        broad:=BroadcastCode;
        GetAssyRatioInfo;
        Inv_Dataset.Filtered := False;
        Inv_DataSet.Locate('Broadcast Code',broad,[]);
      end;
    end;
    SetDetailBoxes;
    BroadcastCode_Edit.SetFocus;
  end;
end;        //Insert_ButtonClick

procedure TAssyRatioMaster_Form.Update_ButtonClick(Sender: TObject);
var
  broad:string;
begin
  if Verify then
  begin
    HoldDetails(False);
    with Data_Module do
    begin
      Data_Module.UpdateAssyRatioInfo;
      broad:=BroadcastCode;
      GetAssyRatioInfo;
      Inv_Dataset.Filtered := False;
      Inv_DataSet.Locate('Broadcast Code',broad,[]);
    end;
    SetDetailBoxes;
    BroadcastCode_Edit.SetFocus;
  end;
end;        //Update_ButtonClick

procedure TAssyRatioMaster_Form.Delete_ButtonClick(Sender: TObject);
begin
  HoldDetails(False);
  If MessageDlg('Are you sure you wish to delete' + #13 +
                 Data_Module.BroadcastCode + ' from the database?',
                 mtWarning, [mbYes, mbNo], 0) = mrYes Then
  Begin
    Data_Module.DeleteAssyRatioInfo;
    Data_Module.GetAssyRatioInfo;
    Data_Module.ClearControls(AssyRatioMaster_Panel);
  End;
  BroadcastCode_Edit.SetFocus;
end;        //Delete_ButtonClick

procedure TAssyRatioMaster_Form.Search_ButtonClick(Sender: TObject);
var
  fFound: Boolean;
Begin
  If (Trim(BroadcastCode_Edit.Text ) = '') Then
    ShowMessage('Please enter a broadcast code before searching.')
  Else
  Begin
    fFound := SearchGrid(BroadcastCode_Edit.Text);
    If fFound Then
    Begin
      SetDetailBoxes;
    End
    Else
    Begin
      ShowMessage('No matches were found for your query.');
    End;
  end;
  BroadcastCode_Edit.SetFocus;
end;        //Search_ButtonClick

procedure TAssyRatioMaster_Form.Clear_ButtonClick(Sender: TObject);
begin
  with Data_Module do
  begin
    Inv_Dataset.Filtered:=False;
    ClearControls(AssyRatioMaster_Panel);
    TireRatio3_MaskEdit.Text:='0';
    TireRatio2_MaskEdit.Text:='0';
    TireRatio1_MaskEdit.Text:='0';
    WheelRatio3_MaskEdit.Text:='0';
    WheelRatio2_MaskEdit.Text:='0';
    WheelRatio1_MaskEdit.Text:='0';
  end;
  BroadcastCode_Edit.SetFocus;
end;      //Clear_ButtonClick

function TAssyRatioMaster_Form.SearchGrid(BroadCode: String): Boolean;
begin
  Result := False;
  Data_Module.ClearControls(AssyRatioMaster_Panel);
  TireRatio3_MaskEdit.Text:='0';
  TireRatio2_MaskEdit.Text:='0';
  TireRatio1_MaskEdit.Text:='0';
  WheelRatio3_MaskEdit.Text:='0';
  WheelRatio2_MaskEdit.Text:='0';
  WheelRatio1_MaskEdit.Text:='0';

  try
    With Data_Module.Inv_DataSet do
    Begin
      Filtered := False;
      Filter := '[Broadcast Code] = ' + QuotedStr(BroadCode);
      Filtered := True;
      If RecordCount > 0 Then
      Begin
        HoldDetails(True);
        Result := True;
      End
      else
      begin
        Filtered := False;
        Data_Module.ClearControls(AssyRatioMaster_Panel);
      end;
    End;      //With
  except
    on e:exception do
      ShowMessage('Error in Search' + #13 + e.Message);
  end;      //try...except
End;        //SearchGrid

procedure TAssyRatioMaster_Form.SetDetailBoxes;
Begin
  With Data_Module do
  Begin
    SearchCombo(AssyCode_ComboBox, AssyCode);
    AssyName_Edit.Text := AssyName;

//  11-14-2002
    BroadcastCode_Edit.Text := BroadcastCode;

    case TireQty of
      0: TireQty_RadioGroup.ItemIndex := 0;
      4: TireQty_RadioGroup.ItemIndex := 1;
      5: TireQty_RadioGroup.ItemIndex := 2;
      Else
        TireQty_RadioGroup.ItemIndex := 0;
    End;


//  11-14-2002
    SearchCombo(TirePartNum1_ComboBox, TirePartNum1);
    TireRatio1_MaskEdit.Text := IntToStr(TireRatio1);
    SearchCombo(TirePartNum2_ComboBox, TirePartNum2);
    TireRatio2_MaskEdit.Text := IntToStr(TireRatio2);
    SearchCombo(TirePartNum3_ComboBox, TirePartNum3);
    TireRatio3_MaskEdit.Text := IntToStr(TireRatio3);
    SearchCombo(WheelPartNum1_ComboBox, WheelPartNum1);
    WheelRatio1_MaskEdit.Text := IntToStr(WheelRatio1);
    SearchCombo(WheelPartNum2_ComboBox, WheelPartNum2);
    WheelRatio2_MaskEdit.Text := IntToStr(WheelRatio2);
    SearchCombo(WheelPartNum3_ComboBox, WheelPartNum3);
    WheelRatio3_MaskEdit.Text := IntToStr(WheelRatio3);

    case WheelQTY of
      1: WheelQTY_RadioGroup.ItemIndex := 0;
      4: WheelQTY_RadioGroup.ItemIndex := 1;
      5: WheelQTY_RadioGroup.ItemIndex := 2;
      Else
        WheelQTY_RadioGroup.ItemIndex := 0;
    end;      //case

    SpareTireQTY_RadioGroup.ItemIndex := SpareTireQty;
    SpareTirePartsCode_Edit.Text := SpareTirePartNum;
    SpareWheelPartsCode_Edit.Text := SpareWheelPartNum;
  End;      //With
End;          //SetDetailBoxes

procedure TAssyRatioMaster_Form.HoldDetails(fFromGrid: Boolean);
Begin
  If fFromGrid Then
  Begin
    With AssyRatioMaster_DBGrid.DataSource.DataSet do
    Begin
//  11-14-2002
      Data_Module.BroadcastCode     := Fields[0].AsString;
      Data_Module.BroadcastCodePrev := Fields[0].AsString;
      Data_Module.AssyCode          := Fields[1].AsString;
      Data_Module.AssyName          := Fields[2].AsString;
      Data_Module.TireQty           := Fields[3].AsInteger;
      Data_Module.TirePartNum1      := Fields[4].AsString;
      Data_Module.TireRatio1        := Fields[5].AsInteger;
      Data_Module.TirePartNum2      := Fields[6].AsString;
      Data_Module.TireRatio2        := Fields[7].AsInteger;
      Data_Module.TirePartNum3      := Fields[8].AsString;
      Data_Module.TireRatio3        := Fields[9].AsInteger;
      Data_Module.WheelQty          := Fields[10].AsInteger;
      Data_Module.WheelPartNum1     := Fields[11].AsString;
      Data_Module.WheelRatio1       := Fields[12].AsInteger;
      Data_Module.WheelPartNum2     := Fields[13].AsString;
      Data_Module.WheelRatio2       := Fields[14].AsInteger;
      Data_Module.WheelPartNum3     := Fields[15].AsString;
      Data_Module.WheelRatio3       := Fields[16].AsInteger;
      Data_Module.SpareTireQty      := Fields[17].AsInteger;
      Data_Module.SpareTirePartNum  := Fields[18].AsString;  
      Data_Module.SpareWheelPartNum := Fields[19].AsString;
    End;      //With
  End
  Else
  Begin
    With Data_Module do
    Begin
      if AssyCode_ComboBox.Text = ' ' then
        AssyCode :=''
      else
        AssyCode := AssyCode_ComboBox.Text;

      AssyName := AssyName_Edit.Text;
      BroadcastCode := BroadcastCode_Edit.Text;
//  11-14-2002
      case TireQty_RadioGroup.ItemIndex of
        0: TireQty := 1;
        1: TireQty := 4;
        2: TireQty := 5;
        Else
           TireQty := 1;
      end;      //case
//  11-14-2002
      if TirePartNum1_ComboBox.Text = ' ' then
        TirePartNum1 := ''
      else
        TirePartNum1 := TirePartNum1_ComboBox.Text;
      TireRatio1 := StrToInt(Trim(TireRatio1_MaskEdit.Text));

      if TirePartNum2_ComboBox.Text= ' ' then
        TirePartNum2 := ''
      else
        TirePartNum2 := TirePartNum2_ComboBox.Text;
      TireRatio2 := StrToInt(Trim(TireRatio2_MaskEdit.Text));

      if TirePartNum3_ComboBox.Text= ' ' then
        TirePartNum3 := ''
      else
        TirePartNum3 := TirePartNum3_ComboBox.Text;
      TireRatio3 := StrToInt(Trim(TireRatio3_MaskEdit.Text));

      if WheelPartNum1_ComboBox.Text = ' ' then
        WheelPartNum1 := ''
      else
        WheelPartNum1 := WheelPartNum1_ComboBox.Text;
      WheelRatio1 := StrToInt(Trim(WheelRatio1_MaskEdit.Text));

      if WheelPartNum2_ComboBox.Text=' ' then
        WheelPartNum2 := ''
      else
        WheelPartNum2 := WheelPartNum2_ComboBox.Text;
      WheelRatio2 := StrToInt(Trim(WheelRatio2_MaskEdit.Text));

      if WheelPartNum3_ComboBox.Text=' ' then
        WheelPartNum3 := ''
      else
        WheelPartNum3 := WheelPartNum3_ComboBox.Text;
      WheelRatio3 := StrToInt(Trim(WheelRatio3_MaskEdit.Text));

      case WheelQty_RadioGroup.ItemIndex of
        0: WheelQty := 1;
        1: WheelQty := 4;
        2: WheelQty := 5;
        Else
           WheelQty := 1;
      end;      //case
      SpareTireQty := SpareTireQty_RadioGroup.ItemIndex;
      SpareTirePartNum := SpareTirePartsCode_Edit.Text;
      SpareWheelPartNum := SpareWheelPartsCode_Edit.Text;
    End;      //With
  End;
End;          //HoldDetails

procedure TAssyRatioMaster_Form.FormCreate(Sender: TObject);
begin
  Data_Module.Inv_DataSet.Filter:='';
  Data_Module.Inv_DataSet.Filtered:=FALSE;
  Data_Module.GetAssyRatioInfo;
  AssyRatio_DataSource.DataSet:=Data_Module.Inv_DataSet;
  GetParts('INV_PARTS_STOCK_MST', 'VC_TIRE_WHEEL', 'T', 'VC_PART_NUMBER', TirePartNum1_ComboBox);
  GetParts('INV_PARTS_STOCK_MST', 'VC_TIRE_WHEEL', 'T', 'VC_PART_NUMBER', TirePartNum2_ComboBox);
  GetParts('INV_PARTS_STOCK_MST', 'VC_TIRE_WHEEL', 'T', 'VC_PART_NUMBER', TirePartNum3_ComboBox);
  GetParts('INV_PARTS_STOCK_MST', 'VC_TIRE_WHEEL', 'W', 'VC_PART_NUMBER', WheelPartNum1_ComboBox);
  GetParts('INV_PARTS_STOCK_MST', 'VC_TIRE_WHEEL', 'W', 'VC_PART_NUMBER', WheelPartNum2_ComboBox);
  GetParts('INV_PARTS_STOCK_MST', 'VC_TIRE_WHEEL', 'W', 'VC_PART_NUMBER', WheelPartNum3_ComboBox);
  GetParts('INV_FORECAST_DETAIL_INF', '', '', 'VC_ASSY_PART_NUMBER_CODE', AssyCode_ComboBox);
  AssyCode_Combobox.Text := '';
  Data_Module.ClearControls(AssyRatioMaster_Panel);
  TireRatio3_MaskEdit.Text:='0';
  TireRatio2_MaskEdit.Text:='0';
  TireRatio1_MaskEdit.Text:='0';
  WheelRatio3_MaskEdit.Text:='0';
  WheelRatio2_MaskEdit.Text:='0';
  WheelRatio1_MaskEdit.Text:='0';
end;

procedure TAssyRatioMaster_Form.GetParts(tablename, selectfieldname, fPartType, returnfieldname: String; ComboBox: TComboBox);
var
  fTableAndWhere: String;
Begin
  try
    try
      if selectfieldname <> '' then
        fTableAndWhere := tablename+' WHERE '+selectfieldname+' = ' + QuotedStr(fPartType)
      else
        fTableAndWhere := tablename;

        Data_Module.SelectSingleField(fTableAndWhere, returnfieldname, ComboBox);
    except
      On e:exception do
      Begin
        MessageDlg('Unable to get a list of parts.', mtError, [mbOK], 0);
      End;
    end;      //try...except
  finally
  end;      //try...finally
End;      //GetParts

procedure TAssyRatioMaster_Form.ASSYRatioMaster_DBGridKeyUp(
  Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  HoldDetails(True);
  SetDetailBoxes;
end;

procedure TAssyRatioMaster_Form.ASSYRatioMaster_DBGridMouseUp(
  Sender: TObject; Button: TMouseButton; Shift: TShiftState; X,
  Y: Integer);
begin
  HoldDetails(True);
  SetDetailBoxes;
end;

procedure TAssyRatioMaster_Form.AssyRatio_DataSourceDataChange(
  Sender: TObject; Field: TField);
begin
  HoldDetails(True);
  SetDetailBoxes;
end;

procedure TAssyRatioMaster_Form.FormShow(Sender: TObject);
begin
  BroadcastCode_Edit.SetFocus;
end;


procedure TAssyRatioMaster_Form.TotalTireRatio_EditChange(Sender: TObject);
begin
  if (StrToInt(TEdit(Sender).Text) = 100) OR (StrToInt(TEdit(Sender).Text) = 0)then
    TEdit(Sender).Font.Color:=clBlack
  else
    TEdit(Sender).Font.Color:=clRed;
end;

procedure TAssyRatioMaster_Form.TirePartNum3_ComboBoxChange(
  Sender: TObject);
begin
  if TirePartNum2_ComboBox.ItemIndex = 0 then
  begin
    TirePartNum3_ComboBox.ItemIndex:=0;
    TirePartNum2_ComboBox.SetFocus;
    exit;
  end;

  if TirePartNum3_ComboBox.ItemIndex = 0 then
  begin
    TireRatio3_MaskEdit.Text:='0';
  end;
end;

procedure TAssyRatioMaster_Form.TirePartNum2_ComboBoxChange(
  Sender: TObject);
begin
  if TirePartNum1_ComboBox.ItemIndex = 0 then
  begin
    TirePartNum2_ComboBox.ItemIndex:=0;
    TirePartNum1_ComboBox.SetFocus;
    exit;
  end;

  if TirePartNum2_ComboBox.ItemIndex = 0 then
  begin
    TirePartNum3_ComboBox.ItemIndex:=0;
    TireRatio3_MaskEdit.Text:='0';
    TireRatio2_MaskEdit.Text:='0';
  end;
end;

procedure TAssyRatioMaster_Form.TirePartNum1_ComboBoxChange(
  Sender: TObject);
begin
  if TirePartNum1_ComboBox.ItemIndex = 0 then
  begin
    TirePartNum3_ComboBox.ItemIndex:=0;
    TireRatio3_MaskEdit.Text:='0';
    TirePartNum2_ComboBox.ItemIndex:=0;
    TireRatio2_MaskEdit.Text:='0';
    TireRatio1_MaskEdit.Text:='0';
  end;
end;

procedure TAssyRatioMaster_Form.WheelPartNum3_ComboBoxChange(
  Sender: TObject);
begin
  if WheelPartNum2_ComboBox.ItemIndex = 0 then
  begin
    WheelPartNum3_ComboBox.ItemIndex:=0;
    WheelPartNum2_ComboBox.SetFocus;
    exit;
  end;

  if WheelPartNum3_ComboBox.ItemIndex = 0 then
  begin
    WheelRatio3_MaskEdit.Text:='0';
  end;
end;

procedure TAssyRatioMaster_Form.WheelPartNum2_ComboBoxChange(
  Sender: TObject);
begin
  if WheelPartNum1_ComboBox.ItemIndex = 0 then
  begin
    WheelPartNum2_ComboBox.ItemIndex:=0;
    WheelPartNum1_ComboBox.SetFocus;
    exit;
  end;

  if WheelPartNum2_ComboBox.ItemIndex = 0 then
  begin
    WheelPartNum3_ComboBox.ItemIndex:=0;
    WheelRatio3_MaskEdit.Text:='0';
    WheelRatio2_MaskEdit.Text:='0';
  end;
end;

procedure TAssyRatioMaster_Form.WheelPartNum1_ComboBoxChange(
  Sender: TObject);
begin
  if WheelPartNum1_ComboBox.ItemIndex = 0 then
  begin
    WheelPartNum3_ComboBox.ItemIndex:=0;
    WheelRatio3_MaskEdit.Text:='0';
    WheelPartNum2_ComboBox.ItemIndex:=0;
    WheelRatio2_MaskEdit.Text:='0';
    WheelRatio1_MaskEdit.Text:='0';
  end;
end;

procedure TAssyRatioMaster_Form.TireRatio1_MaskEditChange(Sender: TObject);
begin
  if TirePartNum1_ComboBox.ItemIndex = 0 then
    TMaskEdit(Sender).Text:='0';
  if TMaskEdit(Sender).Text = '' then
  begin
    TMaskEdit(Sender).Text:='0';
  end;

  try
    if trim(TireRatio1_MaskEdit.Text) <> '' then
    begin
      TotalTireRatio_Edit.Text:=trim(TireRatio1_MaskEdit.Text);
      if trim(TireRatio2_MaskEdit.Text) <> '' then
      begin
        TotalTireRatio_Edit.Text:=IntToStr(StrToInt(trim(TireRatio2_MaskEdit.Text))+StrToInt(TotalTireRatio_Edit.Text));
        if trim(TireRatio3_MaskEdit.Text) <> '' then
        begin
          TotalTireRatio_Edit.Text:=IntToStr(StrToInt(trim(TireRatio3_MaskEdit.Text))+StrToInt(TotalTireRatio_Edit.Text));
        end;
      end;
    end
  except
  end;
end;

procedure TAssyRatioMaster_Form.TireRatio2_MaskEditChange(Sender: TObject);
begin
  if TirePartNum2_ComboBox.ItemIndex = 0 then
    TMaskEdit(Sender).Text:='0';
  if TMaskEdit(Sender).Text = '' then
  begin
    TMaskEdit(Sender).Text:='0';
  end;

  try
    if trim(TireRatio1_MaskEdit.Text) <> '' then
    begin
      TotalTireRatio_Edit.Text:=trim(TireRatio1_MaskEdit.Text);
      if trim(TireRatio2_MaskEdit.Text) <> '' then
      begin
        TotalTireRatio_Edit.Text:=IntToStr(StrToInt(trim(TireRatio2_MaskEdit.Text))+StrToInt(TotalTireRatio_Edit.Text));
        if trim(TireRatio3_MaskEdit.Text) <> '' then
        begin
          TotalTireRatio_Edit.Text:=IntToStr(StrToInt(trim(TireRatio3_MaskEdit.Text))+StrToInt(TotalTireRatio_Edit.Text));
        end;
      end;
    end
  except
  end;
end;

procedure TAssyRatioMaster_Form.TireRatio3_MaskEditChange(Sender: TObject);
begin
  if TirePartNum3_ComboBox.ItemIndex = 0 then
    TMaskEdit(Sender).Text:='0';

  if TMaskEdit(Sender).Text = '' then
  begin
    TMaskEdit(Sender).Text:='0';
  end;

  try
    if trim(TireRatio1_MaskEdit.Text) <> '' then
    begin
      TotalTireRatio_Edit.Text:=trim(TireRatio1_MaskEdit.Text);
      if trim(TireRatio2_MaskEdit.Text) <> '' then
      begin
        TotalTireRatio_Edit.Text:=IntToStr(StrToInt(trim(TireRatio2_MaskEdit.Text))+StrToInt(TotalTireRatio_Edit.Text));
        if trim(TireRatio3_MaskEdit.Text) <> '' then
        begin
          TotalTireRatio_Edit.Text:=IntToStr(StrToInt(trim(TireRatio3_MaskEdit.Text))+StrToInt(TotalTireRatio_Edit.Text));
        end;
      end;
    end
  except
  end;
end;

procedure TAssyRatioMaster_Form.WheelRatio1_MaskEditChange(
  Sender: TObject);
begin
  if WheelPartNum1_ComboBox.ItemIndex = 0 then
    TMaskEdit(Sender).Text:='0';
  if TMaskEdit(Sender).Text = '' then
  begin
    TMaskEdit(Sender).Text:='0';
  end;

  try
    if trim(WheelRatio1_MaskEdit.Text) <> '' then
    begin
      TotalWheelRatio_Edit.Text:=trim(WheelRatio1_MaskEdit.Text);
      if trim(WheelRatio2_MaskEdit.Text) <> '' then
      begin
        TotalWheelRatio_Edit.Text:=IntToStr(StrToInt(trim(WheelRatio2_MaskEdit.Text))+StrToInt(TotalWheelRatio_Edit.Text));
        if trim(TireRatio3_MaskEdit.Text) <> '' then
        begin
          TotalWheelRatio_Edit.Text:=IntToStr(StrToInt(trim(WheelRatio3_MaskEdit.Text))+StrToInt(TotalWheelRatio_Edit.Text));
        end;
      end;
    end
  except
  end;
end;

procedure TAssyRatioMaster_Form.WheelRatio2_MaskEditChange(
  Sender: TObject);
begin
  if WheelPartNum2_ComboBox.ItemIndex = 0 then
    TMaskEdit(Sender).Text:='0';

  if TMaskEdit(Sender).Text = '' then
  begin
    TMaskEdit(Sender).Text:='0';
  end;
  try
    if trim(WheelRatio1_MaskEdit.Text) <> '' then
    begin
      TotalWheelRatio_Edit.Text:=trim(WheelRatio1_MaskEdit.Text);
      if trim(WheelRatio2_MaskEdit.Text) <> '' then
      begin
        TotalWheelRatio_Edit.Text:=IntToStr(StrToInt(trim(WheelRatio2_MaskEdit.Text))+StrToInt(TotalWheelRatio_Edit.Text));
        if trim(TireRatio3_MaskEdit.Text) <> '' then
        begin
          TotalWheelRatio_Edit.Text:=IntToStr(StrToInt(trim(WheelRatio3_MaskEdit.Text))+StrToInt(TotalWheelRatio_Edit.Text));
        end;
      end;
    end
  except
  end;
end;

procedure TAssyRatioMaster_Form.WheelRatio3_MaskEditChange(
  Sender: TObject);
begin
  if WheelPartNum3_ComboBox.ItemIndex = 0 then
    TMaskEdit(Sender).Text:='0';

  if TMaskEdit(Sender).Text = '' then
  begin
    TMaskEdit(Sender).Text:='0';
  end;
  try
    if trim(WheelRatio1_MaskEdit.Text) <> '' then
    begin
      TotalWheelRatio_Edit.Text:=trim(WheelRatio1_MaskEdit.Text);
      if trim(WheelRatio2_MaskEdit.Text) <> '' then
      begin
        TotalWheelRatio_Edit.Text:=IntToStr(StrToInt(trim(WheelRatio2_MaskEdit.Text))+StrToInt(TotalWheelRatio_Edit.Text));
        if trim(TireRatio3_MaskEdit.Text) <> '' then
        begin
          TotalWheelRatio_Edit.Text:=IntToStr(StrToInt(trim(WheelRatio3_MaskEdit.Text))+StrToInt(TotalWheelRatio_Edit.Text));
        end;
      end;
    end
  except
  end;
end;



end.
