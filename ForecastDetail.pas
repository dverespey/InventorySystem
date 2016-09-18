//***********************************************************
//  InventorySystem
//
//  Copyright (c) 2003 Failproof Manufacturing Systems
//
//***********************************************************
//
//  02/12/2003  David verespey  Initial creation

unit ForecastDetail;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Mask, StdCtrls, Grids, DBGrids, ExtCtrls, DB;

type
  TForecastDetail_Form = class(TForm)
    ManagementButtons_Panel: TPanel;
    Insert_Button: TButton;
    Update_Button: TButton;
    Search_Button: TButton;
    Clear_Button: TButton;
    Close_Button: TButton;
    Delete_Button: TButton;
    ForecastDetail_DBGrid: TDBGrid;
    ForecastDetail_Panel: TPanel;
    SizeCode_Label: TLabel;
    ASSYCode_Label: TLabel;
    DailyUsage_Label: TLabel;
    TirePartNum_ComboBox: TComboBox;
    WheelPartNum_ComboBox: TComboBox;
    ForecastDetail_DataSource: TDataSource;
    Label1: TLabel;
    ForecastRatio_Edit: TEdit;
    Label2: TLabel;
    Kanban_Edit: TEdit;
    Label3: TLabel;
    TireForecastRatio_Edit: TEdit;
    Label4: TLabel;
    WheelForecastRatio_Edit: TEdit;
    Label5: TLabel;
    EffectiveMonth_ComboBox: TComboBox;
    AssyRatioMaster_Label: TLabel;
    TireQty_Label: TLabel;
    AssyQty_RadioGroup: TRadioGroup;
    BroadcastCode_Label: TLabel;
    BroadcastCode_Edit: TEdit;
    ForecastPartsCode_ComboBox: TComboBox;
    Label6: TLabel;
    ValvePartNum_ComboBox: TComboBox;
    Label7: TLabel;
    FilmPArtNum_ComboBox: TComboBox;
    LabelPartNum_ComboBox: TComboBox;
    Label8: TLabel;
    Label9: TLabel;
    Misc1PartNum_ComboBox: TComboBox;
    Label10: TLabel;
    Misc2PartNum_ComboBox: TComboBox;
    Label11: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure Clear_ButtonClick(Sender: TObject);
    procedure Search_ButtonClick(Sender: TObject);
    procedure ForecastDetail_DBGridKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ForecastDetail_DBGridMouseUp(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure ForecastPartsCode_EditChange(Sender: TObject);
    procedure Delete_ButtonClick(Sender: TObject);
    procedure Update_ButtonClick(Sender: TObject);
    procedure Insert_ButtonClick(Sender: TObject);
    procedure ForecastDetail_DataSourceDataChange(Sender: TObject;
      Field: TField);
    procedure FormShow(Sender: TObject);
    procedure ForecastPartsCode_ComboBoxSelect(Sender: TObject);
  private
    { Private declarations }
    procedure GetAssy(ComboBox: TComboBox);
    procedure SetDetailBoxes;
    procedure HoldDetails(fFromGrid: Boolean);
    function SearchGrid: Boolean;
    function Validate: boolean;
  public
    { Public declarations }
    function Execute: boolean;
  end;

var
  ForecastDetail_Form: TForecastDetail_Form;

implementation

uses DataModule;

{$R *.dfm}

function TForecastDetail_Form.Execute: boolean;
Begin
  Result:= True;
  try
    try
      ShowModal;
    except
      On E:Exception do
        showMessage('Unable to generate Forecast Detail screen.'
          + #13 + 'ERROR:'
          + #13 + E.Message);
    end;
  finally
    If ModalResult = mrCancel Then
      Result:= False;
  end;
end;    //Execute

procedure TForecastDetail_Form.FormCreate(Sender: TObject);
var
  i,z,m,y:integer;
begin
  EffectiveMonth_ComboBox.Items.Clear;
  m:=StrToInt(formatdatetime('m',now));
  y:=StrToInt(formatdatetime('yyyy',now));
  EffectiveMonth_ComboBox.Items.Add(' ');
  for i:=-1 to 10 do
  begin
    if (i+m) <> 12 then
    begin
      z:=(i+m) mod 12;
      if i+m = 13 then
        INC(y);
    end
    else
      z:=12;
    if z < 10 then
      EffectiveMonth_ComboBox.Items.Add(IntToStr(y)+'/0'+IntToStr(z))
    else
      EffectiveMonth_ComboBox.Items.Add(IntToStr(y)+'/'+IntToStr(z));
  end;

  Data_Module.Inv_DataSet.Filter:='';
  Data_Module.GetForecastDetailInfo;
  Data_Module.Inv_DataSet.Filtered:=FALSE;
  ForecastDetail_DataSource.DataSet:=Data_Module.Inv_DataSet;
  Data_Module.SelectDependantSingleField('SELECT_DependantPartNumber_PartType', '@PArtType', 'VC_PART_NUMBER', 'TIRE', TirePartNum_ComboBox);
  Data_Module.SelectDependantSingleField('SELECT_DependantPartNumber_PartType', '@PArtType', 'VC_PART_NUMBER', 'WHEEL', WheelPartNum_ComboBox);
  Data_Module.SelectDependantSingleField('SELECT_DependantPartNumber_PartType', '@PArtType', 'VC_PART_NUMBER', 'VALVE', ValvePartNum_ComboBox);
  Data_Module.SelectDependantSingleField('SELECT_DependantPartNumber_PartType', '@PArtType', 'VC_PART_NUMBER', 'FILM', FilmPartNum_ComboBox);
  Data_Module.SelectDependantSingleField('SELECT_DependantPartNumber_PartType', '@PArtType', 'VC_PART_NUMBER', 'LABEL', LabelPartNum_ComboBox);
  Data_Module.SelectDependantSingleField('SELECT_DependantPartNumber_PartType', '@PArtType', 'VC_PART_NUMBER', 'MISC', Misc1PartNum_ComboBox);
  Data_Module.SelectDependantSingleField('SELECT_DependantPartNumber_PartType', '@PArtType', 'VC_PART_NUMBER', 'MISC', Misc2PartNum_ComboBox);
  GetAssy(ForecastPartsCode_ComboBox);
  Data_Module.Inv_DataSet.Filtered := False;
end;


procedure TForecastDetail_Form.GetAssy(ComboBox: TComboBox);
var
  fTableAndWhere: String;
Begin
  try
    try
      fTableAndWhere := 'INV_FORECAST_DETAIL_INF';
      Data_Module.SelectSingleField(fTableAndWhere, 'VC_ASSY_PART_NUMBER_CODE', ComboBox);
    except
      On e:exception do
      Begin
        MessageDlg('Unable to get a list of parts.', mtError, [mbOK], 0);
      End;
    end;      //try...except
  finally
  end;      //try...finally
End;      //GetParts

procedure TForecastDetail_Form.Clear_ButtonClick(Sender: TObject);
begin
  with Data_Module do
  begin
    Inv_Dataset.Filtered:=False;
    ClearControls(ForecastDetail_Panel);
  end;
  ForecastPartsCode_ComboBox.SetFocus;
end;


procedure TForecastDetail_Form.HoldDetails(fFromGrid: Boolean);
var
  fr:integer;
Begin
  If fFromGrid Then
  Begin
    with ForecastDetail_DBGrid.DataSource.DataSet do
    begin
      Data_Module.RecordID        := Fields[0].AsInteger;
      Fields[0].Visible:=FALSE;
      Data_Module.AssyCode        := Fields[1].AsString;
      Data_Module.AssyCodePrev    := Fields[1].AsString;
      Data_Module.BeginDatestr    := Fields[2].AsString;
      Data_Module.TirePartNum1    := Fields[3].AsString;
      Data_Module.TireRatio1      := Fields[4].AsInteger;
      Data_Module.WheelPartNum1   := Fields[5].AsString;
      Data_Module.WheelRatio1     := Fields[6].AsInteger;
      Data_Module.ForecastRatio   := Fields[7].AsInteger;
      Data_Module.Kanban          := Fields[8].AsString;
      Data_Module.BroadcastCode   := Fields[9].AsString;
      Data_Module.ValvePartNum    := Fields[10].AsString;
      Data_Module.FilmPartNum     := Fields[11].AsString;
      Data_Module.Quantity        := Fields[12].AsInteger;
      Data_Module.LabelPartNum    := Fields[13].AsString;
      Data_Module.Misc1PartNum    := Fields[14].AsString;
      Data_Module.Misc2PartNum    := Fields[15].AsString;
    end;
  End
  Else
  Begin
    With Data_Module do
    Begin
      AssyCode        := ForecastPartsCode_ComboBox.Text;
      BeginDatestr    := EffectiveMonth_ComboBox.Text;
      BroadcastCode   := BroadcastCode_Edit.Text;


      case AssyQty_RadioGroup.ItemIndex of
        0: Data_Module.Quantity := 1;
        1: Data_Module.Quantity := 2;
        2: Data_Module.Quantity := 4;
        3: Data_Module.Quantity := 5;
        Else
           Data_Module.Quantity := 1;
      end;

      if TirePartNum_ComboBox.Text = ' ' then
        TirePartNum1:=''
      else
        TirePartNum1    := TirePartNum_ComboBox.Text;

      if WheelPartNum_ComboBox.Text = ' ' then
        WheelPartNum1   := ''
      else
        WheelPartNum1   := WheelPartNum_ComboBox.Text;

      if ValvePartNum_ComboBox.Text = ' ' then
        ValvePartNum   := ''
      else
        ValvePartNum   := ValvePartNum_ComboBox.Text;

      if FilmPartNum_ComboBox.Text = ' ' then
        FilmPartNum   := ''
      else
        FilmPartNum   := FilmPartNum_ComboBox.Text;

      if LabelPartNum_ComboBox.Text = ' ' then
        LabelPartNum   := ''
      else
        LabelPartNum   := LabelPartNum_ComboBox.Text;

      if Misc1PartNum_ComboBox.Text = ' ' then
        Misc1PartNum   := ''
      else
        Misc1PartNum   := Misc1PartNum_ComboBox.Text;

      if Misc2PartNum_ComboBox.Text = ' ' then
        Misc2PartNum   := ''
      else
        Misc2PartNum   := Misc2PartNum_ComboBox.Text;



      if not TryStrToInt(ForecastRatio_Edit.Text, fr) then
      begin
        ShowMessage('Invalid ratio');
        ForecastRatio_Edit.SetFocus;
      end
      else
        ForecastRatio:=fr;
      if (ForecastRatio > 100) or (ForecastRatio < 0) then
      begin
        ShowMessage('Ratio must be between 0 and 100');
        ForecastRatio_Edit.SetFocus;
      end;

      if not TryStrToInt(TireForecastRatio_Edit.Text, fr) then
      begin
        ShowMessage('Invalid ratio');
        TireForecastRatio_Edit.SetFocus;
      end
      else
        TireRatio1:=fr;
      if (TireRatio1 > 100) or (TireRatio1 < 0) then
      begin
        ShowMessage('Ratio must be between 0 and 100');
        TireForecastRatio_Edit.SetFocus;
      end;

      if not TryStrToInt(WheelForecastRatio_Edit.Text, fr) then
      begin
        ShowMessage('Invalid ratio');
        WheelForecastRatio_Edit.SetFocus;
      end
      else
        WheelRatio1:=fr;
      if (WheelRatio1 > 100) or (WheelRatio1 < 0) then
      begin
        ShowMessage('Ratio must be between 0 and 100');
        WheelForecastRatio_Edit.SetFocus;
      end;
      Kanban:=Kanban_Edit.Text;
    End;      //With
  End;
End;          //HoldDetails

procedure TForecastDetail_Form.SetDetailBoxes;
Begin
  With Data_Module do
  Begin
    SearchCombo(ForecastPartsCode_ComboBox, Assycode);
    SearchCombo(EffectiveMonth_ComboBox, BeginDatestr);
    SearchCombo(TirePartNum_ComboBox, TirePartNum1);
    SearchCombo(WheelPartNum_ComboBox, WheelPartNum1);
    SearchCombo(ValvePartNum_ComboBox, ValvePartNum);
    SearchCombo(FilmPartNum_ComboBox, FilmPartNum);
    SearchCombo(LabelPartNum_ComboBox, LabelPartNum);
    SearchCombo(Misc1PartNum_ComboBox, Misc1PartNum);
    SearchCombo(Misc2PartNum_ComboBox, Misc2PartNum);

    ForecastRatio_Edit.Text:=IntToStr(ForecastRatio);
    TireForecastRatio_Edit.Text:=IntToStr(TireRatio1);
    WheelForecastRatio_Edit.Text:=IntToStr(WheelRatio1);
    Kanban_Edit.Text:=Kanban;
    BroadcastCode_Edit.Text:=BroadcastCode;
    case Quantity of
      0: AssyQty_RadioGroup.ItemIndex := 0;
      2: AssyQty_RadioGroup.ItemIndex := 1;
      4: AssyQty_RadioGroup.ItemIndex := 2;
      5: AssyQty_RadioGroup.ItemIndex := 3;
      Else
        AssyQty_RadioGroup.ItemIndex := 0;
    End;
  End;      //With
End;          //SetDetailBoxes

function TForecastDetail_Form.SearchGrid: Boolean;
begin
  Result := False;

  try
    With Data_Module.Inv_DataSet Do
    Begin
      if (ForecastPartsCode_ComboBox.Text <> '') and (ForecastPartsCode_ComboBox.Text <> ' ') then
      begin
        if (EffectiveMonth_ComboBox.Text <> ' ') and (EffectiveMonth_ComboBox.Text <> '') then
        begin
          Filtered := False;
          Filter := '[Assembly Part Number Code] LIKE ' + QuotedStr(ForecastPartsCode_ComboBox.Text) + ' AND [Active Date] LIKE ' + QuotedStr(EffectiveMonth_ComboBox.Text)
        end
        else
        begin
          Filtered := False;
          Filter := '[Assembly Part Number Code] = ' + QuotedStr(ForecastPartsCode_ComboBox.Text);
        end
      end
      else if (BroadcastCode_Edit.Text <> '') and (BroadcastCode_Edit.Text <> ' ') then
      begin
        if EffectiveMonth_ComboBox.Text <> ' ' then
        begin
          Filtered := False;
          Filter := '[Broadcast Code] = ' + QuotedStr(BroadcastCode_Edit.Text) + ' AND [Active Date] LIKE ' + QuotedStr(EffectiveMonth_ComboBox.Text)
        end
        else
        begin
          Filtered := False;
          Filter := '[Broadcast Code] = ' + QuotedStr(BroadcastCode_Edit.Text);
        end
      end
      else if (EffectiveMonth_ComboBox.Text <> ' ') and (EffectiveMonth_ComboBox.Text <> '') then
      begin
        Filtered := False;
        Filter := '[Active Date] LIKE ' + QuotedStr(EffectiveMonth_ComboBox.Text)
      end
      else
      begin
        ShowMessage('Search on Forecast(Assembly) Part Number, Broadcast Code and/or Effective Month');
        ForecastPartsCode_ComboBox.SetFocus;
      end;


      Data_Module.ClearControls(ForecastDetail_Panel);

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

procedure TForecastDetail_Form.Search_ButtonClick(Sender: TObject);
var
  fFound: Boolean;
Begin
  fFound := SearchGrid;

  If fFound Then
  Begin
    SetDetailBoxes;
  End
  Else
  Begin
    ShowMessage('No matches were found for your query.');
  End;

  ForecastPartsCode_ComboBox.SelectAll;
  ForecastPartsCode_ComboBox.SetFocus;
end;        //Search_ButtonClick

procedure TForecastDetail_Form.ForecastDetail_DBGridKeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  HoldDetails(True);
  SetDetailBoxes;
end;

procedure TForecastDetail_Form.ForecastDetail_DBGridMouseUp(
  Sender: TObject; Button: TMouseButton; Shift: TShiftState; X,
  Y: Integer);
begin
  HoldDetails(True);
  SetDetailBoxes;
end;

procedure TForecastDetail_Form.ForecastPartsCode_EditChange(
  Sender: TObject);
begin
 // If Length(ForecastPartsCode_Edit.Text) <= 1 Then
   // Data_Module.ClearControls(ForecastDetail_Panel);
end;

procedure TForecastDetail_Form.Delete_ButtonClick(Sender: TObject);
begin
  HoldDetails(False);
  If MessageDlg('Are you sure you wish to delete' + #13 +
                 Data_Module.AssyCode + ' (' + Data_Module.BeginDatestr + ') from the database?',
                 mtWarning, [mbYes, mbNo], 0) = mrYes Then
  Begin
    Data_Module.DeleteForecastDetailInfo;
    Data_Module.GetForecastDetailInfo;
  End;
  Data_Module.ClearControls(ForecastDetail_Panel);
  ForecastPartsCode_ComboBox.SetFocus;
end;

procedure TForecastDetail_Form.Update_ButtonClick(Sender: TObject);
var
  ID:integer;
begin
  if Validate then
  begin
    HoldDetails(False);
    with Data_Module do
    begin
      Id:=REcordID;
      UpdateForecastDetailInfo;
      GetForecastDetailInfo;
      Data_Module.Inv_DataSet.Locate('ID',ID,[]);
    end;
    SetDetailBoxes;
    ForecastPartsCode_ComboBox.SetFocus;
  end;
end;

procedure TForecastDetail_Form.Insert_ButtonClick(Sender: TObject);
var
  assy,eff:string;
begin
  if Validate then
  begin
    HoldDetails(False);
    with Data_Module do
    begin
      If Not InsertForecastDetailInfo Then
        MessageDlg('Unable to INSERT ' + Data_Module.AssyCode, mtInformation, [mbOk], 0)
      else
      begin
        assy:=AssyCode;
        eff:=BeginDateStr;
        GetAssy(ForecastPartsCode_ComboBox);
        GetForecastDetailInfo;
        Data_Module.Inv_DataSet.Locate('Assembly Part Number Code; Active Date',VarArrayOf([Assy,eff]),[]);
      end
    end;
    SetDetailBoxes;
    ForecastPartsCode_ComboBox.SetFocus;
  end;
end;

function TForecastDetail_Form.Validate:boolean;
begin
  // confirm data
  result:=True;

  if Length(ForecastPartsCode_ComboBox.Text) < 12 then
  begin
    ShowMessage('Forecast Part Code must be 12 characters');
    ForecastPartsCode_ComboBox.SetFocus;
    result:=False;
  end;
end;

procedure TForecastDetail_Form.ForecastDetail_DataSourceDataChange(
  Sender: TObject; Field: TField);
begin
  HoldDetails(True);
  SetDetailBoxes;
end;

procedure TForecastDetail_Form.FormShow(Sender: TObject);
begin
  SetDetailBoxes;
  ForecastPartsCode_ComboBox.SetFocus;
end;

procedure TForecastDetail_Form.ForecastPartsCode_ComboBoxSelect(
  Sender: TObject);
begin
  // Update boxes with new data, use search
  try
    if ForecastPartsCode_ComboBox.Text <> ' ' then
    begin
      ForecastDetail_DBGrid.DataSource.DataSet.DisableControls;
      ForecastDetail_DBGrid.DataSource.DataSet.FindFirst;

      While NOT ForecastDetail_DBGrid.DataSource.DataSet.Eof do
      begin
        if ForecastDetail_DBGrid.DataSource.DataSet.Fields[1].AsString = ForecastPartsCode_ComboBox.Text then
        begin
          HoldDetails(True);
          break;
        end;
        ForecastDetail_DBGrid.DataSource.DataSet.FindNext;
      end;

      SetDetailBoxes;
    end
    else
    begin
      with Data_Module do
      begin
        Inv_Dataset.Filtered:=False;
        ClearControls(ForecastDetail_Panel);
      end;
    end
  finally
    ForecastDetail_DBGrid.DataSource.DataSet.EnableControls;
  end;
end;

end.

