//***********************************************************
//  InventorySystem
//
//  Copyright (c) 2002-2005 TAI, Failproof Manufacturing Systems
//
//***********************************************************
//
//  10/25/2002  Aaron Huge  Initial creation

unit RecConfStat;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Grids, DBGrids, DataModule, ComCtrls,
  BmDtEdit, NUMMIBmDateEdit, DateUtils, Mask, DB;

type
  TRecConfStat_Form = class(TForm)
    RecConfStat_Label: TLabel;
    ManagementButtons_Panel: TPanel;
    Insert_Button: TButton;
    Update_Button: TButton;
    Search_Button: TButton;
    Clear_Button: TButton;
    Close_Button: TButton;
    RecConfStat_Panel: TPanel;
    SearchKey_GroupBox: TGroupBox;
    SupplierCode_Label: TLabel;
    FRS_Label: TLabel;
    PartsCode_Label: TLabel;
    RENBAN_Label: TLabel;
    FRS_Edit: TEdit;
    RENBAN_Edit: TEdit;
    InTransit_Label: TLabel;
    Arrival_Label: TLabel;
    TrailerNo_Label: TLabel;
    TrailerNo_Edit: TEdit;
    PlantYard_Label: TLabel;
    ParkingSpot_Label: TLabel;    
    ParkingSpot_Edit: TEdit;
    AssemblerYard_Label: TLabel;
    AssemblerLocation_Label: TLabel;
    AssemblerLocation_Edit: TEdit;
    EmptyTrailer_Label: TLabel;
    Detention_Label: TLabel;
    Detention_Edit: TEdit;
    RecConfStat_DBGrid: TDBGrid;
    InTransit_NUMMIBmDateEdit: TNUMMIBmDateEdit;
    ArrivalDate_NUMMIBmDateEdit: TNUMMIBmDateEdit;
    AssemblerYard_NUMMIBmDateEdit: TNUMMIBmDateEdit;
    EmptyTrailer_NUMMIBmDateEdit: TNUMMIBmDateEdit;
    PlantYard_NUMMIBmDateEdit: TNUMMIBmDateEdit;
    SupplierCode_ComboBox: TComboBox;
    PartsCode_ComboBox: TComboBox;
    Quantity_Label: TLabel;
    Label1: TLabel;
    Warehouse_NUMMIBmDateEdit: TNUMMIBmDateEdit;
    Label2: TLabel;
    Terminated_NUMMIBmDateEdit: TNUMMIBmDateEdit;
    RecConf_DataSource: TDataSource;
    AssemblerLocation_ComboBox: TComboBox;
    Label3: TLabel;
    Order_NUMMIBmDateEdit: TNUMMIBmDateEdit;
    RenbanUpdate_CheckBox: TCheckBox;
    Label4: TLabel;
    Ship_NUMMIBmDateEdit: TNUMMIBmDateEdit;
    Label5: TLabel;
    KanbanCode_ComboBox: TComboBox;
    Panel1: TPanel;
    HideTerminated_CheckBox: TCheckBox;
    SortBy_ComboBox: TComboBox;
    Label6: TLabel;
    Delete_Button: TButton;
    Unordered_Box: TCheckBox;
    Quantity_Edit: TEdit;
    procedure Insert_ButtonClick(Sender: TObject);
    procedure Update_ButtonClick(Sender: TObject);
    procedure Search_ButtonClick(Sender: TObject);
    procedure Clear_ButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure RecConfStat_DBGridKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure RecConfStat_DBGridMouseUp(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure InTransit_NUMMIBmDateEditExit(Sender: TObject);
    procedure SupplierCode_ComboBoxChange(Sender: TObject);
    procedure MaskEditExit(Sender: TObject);
    procedure TextChange(Sender: TObject);
    procedure RecConf_DataSourceDataChange(Sender: TObject; Field: TField);
    procedure FormShow(Sender: TObject);
    procedure ASSEMBLERLocation_ComboBoxChange(Sender: TObject);
    procedure TrailerNo_EditChange(Sender: TObject);
    procedure PartsCode_ComboBoxChange(Sender: TObject);
    procedure HideTerminated_CheckBoxClick(Sender: TObject);
    procedure SortBy_ComboBoxChange(Sender: TObject);
    procedure Delete_ButtonClick(Sender: TObject);
  private
    { Private declarations }
    fSupCode: String;
    fPartNum: String;
    fKanban: string;
    fFRS: String;
    fRenban: String;
    fOrderDate:string;
    fShipDate:string;
    fCheck:boolean;
    procedure HoldDetails(fFromGrid: Boolean);
    function  Validate:boolean;
    procedure SetDetailBoxes;
    function SearchGrid: Boolean;
  public
    { Public declarations }
    function Execute: boolean;
  end;

var
  RecConfStat_Form: TRecConfStat_Form;

implementation

{$R *.dfm}

function TRecConfStat_Form.Execute: boolean;
Begin
  Result:= True;
  try
    try
      Data_Module.SelectSingleField('INV_DOCK_INF', 'VC_DOCK_NAME', AssemblerLocation_ComboBox);

      ShowModal;
      Data_Module.Inv_Dataset.Filtered := False;
      RecConf_DataSource.DataSet:=nil;
    except
      On E:Exception do
        showMessage('Unable to generate Receiving Confirmation Status screen.'
          + #13 + 'ERROR:'
          + #13 + E.Message);
    end;
  finally
    If ModalResult = mrCancel Then
      Result:= False;
  end;
end;        //Execute

procedure TRecConfStat_Form.Delete_ButtonClick(Sender: TObject);
begin
  HoldDetails(False);
  If MessageDlg('Are you sure you wish to delete' + #13 +
                 Data_Module.FRSNo + ' (' + Data_Module.PartNum + ') from the database?',
                 mtWarning, [mbYes, mbNo], 0) = mrYes Then
  Begin
    Data_Module.DeleteRecConfStatInfo;
    Data_Module.ClearControls(RecConfStat_Panel);
    Data_Module.ClearControls(SearchKey_GroupBox);
    Data_Module.GetRecConfStatInfo;
    Data_Module.Inv_DataSet.Filtered := False;
  End;
end;

procedure TRecConfStat_Form.Insert_ButtonClick(Sender: TObject);
var
  part,FRS,REN:string;
begin
  if validate then
  begin
    HoldDetails(False);
    with Data_Module do
    begin
      If Not Data_Module.InsertRecConfStatInfo Then
        MessageDlg('Unable to INSERT Supplier Code: ' + Data_Module.SupplierCode + #13 +
                  'Parts Code: ' + Data_Module.PartNum + #13 +
                  'FRS Number: ' + Data_Module.FRSNo + #13 +
                  'RENBAN: ' + Data_Module.Renban + #13 +
                  'It already exists in the database.', mtInformation, [mbOk], 0)
      else
      begin
        part:=PartNum;
        FRS:=FRSNo;
        RENBAN:=Renban;
        GetRecConfStatInfo;
        Inv_Dataset.Filtered := False;
        Inv_DataSet.Locate('Parts;FRSNo;RENBAN',VarArrayOf([part,FRS,REN]),[])
      end;
    End;
    SetDetailBoxes;
    Renban_Edit.SetFocus;
  end;
end;        //Insert_ButtonClick

procedure TRecConfStat_Form.Update_ButtonClick(Sender: TObject);
var
  part,FRS,REN:string;
begin
  if validate then
  begin
    HoldDetails(False);
    With Data_Module do
    Begin
      part:=PartNum;
      FRS:=FRSNo;
      REN:=Renban;
      if RenbanUpdate_CheckBox.Checked then
        UpdateRecConfStatRenbanInfo
      else
        UpdateRecConfStatInfo;
      GetRecConfStatInfo;
      Inv_DataSet.Locate('Parts;FRSNo;RENBAN',VarArrayOf([part,FRS,REN]),[])
    End;
    SetDetailBoxes;
    Renban_Edit.SetFocus;
  end;
end;        //Update_ButtonClick

procedure TRecConfStat_Form.Search_ButtonClick(Sender: TObject);
var
  fFound: Boolean;
Begin
  If (Trim(SupplierCode_ComboBox.Text ) = '') And (Trim(PartsCode_ComboBox.Text) = '')
    And (Trim(FRS_Edit.Text) = '') And (Trim(Renban_Edit.Text) = '') and (Order_NUMMIBmDateEdit.Text = '')
    and (Ship_NUMMIBmDateEdit.Text = '') and (Trim(KanbanCode_ComboBox.Text ) = '') and (not Unordered_Box.Checked) Then
    ShowMessage('At least one of the following fields' + #13 +
                'must contain data before searching:' + #13 +
                '    Supplier Code' + #13 +
                '    Parts Code' + #13 +
                '    Kanban Code' + #13 +
                '    FRS Number' + #13 +
                '    Order Date' + #13 +
                '    Ship Date' + #13 +
                '    RENBAN Number' +#13+
                'OR the No Order Search must be selected')
  Else
  Begin
    if not Unordered_Box.Checked then
    begin
      if Order_NUMMIBmDateEdit.Text = '' then
        fOrderDate:=Order_NUMMIBmDateEdit.Text
      else
        fOrderDate:=FormatDateTime('yyyymmdd',Order_NUMMIBmDateEdit.date);
    end
    else
    begin
      fOrderDate:='X';
    end;

    if Ship_NUMMIBmDateEdit.Text = '' then
      fShipDate:=Ship_NUMMIBmDateEdit.Text
    else
      fShipDate:=FormatDateTime('yyyymmdd',Ship_NUMMIBmDateEdit.date);

    fSupCode:=Trim(SupplierCode_ComboBox.Text);
    fPartNum:=Trim(PartsCode_ComboBox.Text);
    fKanban:=Trim(KanbanCode_ComboBox.Text);
    fFRS:=Trim(FRS_Edit.Text);
    fRenban:=Trim(Renban_Edit.Text);

    fFound := SearchGrid;
    If fFound Then
    Begin
      SetDetailBoxes;
    End
    Else
    Begin
      ShowMessage('No matches were found for your query.');
    End;
  End;
  Renban_Edit.SetFocus;
end;        //Search_ButtonClick

procedure TRecConfStat_Form.Clear_ButtonClick(Sender: TObject);
begin
  With Data_Module Do
  Begin
    SetTodaysDate(RecConfStat_Panel);
    Inv_Dataset.Filtered := False;
    //SearchGrid;
    //SetComboBoxesWithObj(PartsCode_ComboBox, 'SELECT VC_PART_NUMBER, VC_PARTS_NAME FROM INV_PARTS_STOCK_MST ORDER BY VC_PART_NUMBER');
    //SetComboBoxesWithObj(KanbanCode_ComboBox, 'SELECT VC_KANBAN_NUMBER, VC_PARTS_NAME FROM INV_PARTS_STOCK_MST ORDER BY VC_KANBAN_NUMBER');
    ClearControls(RecConfStat_Panel);
    ClearControls(SearchKey_GroupBox);
  End;        //With
  SupplierCode_ComboBox.ItemIndex:=-1;
  PartsCode_ComboBox.ItemIndex:=-1;
  KanbanCode_ComboBox.ItemIndex:=-1;
  AssemblerLocation_ComboBox.ItemIndex:=-1;
  Renban_Edit.SetFocus;
  RenbanUpdate_CheckBox.Enabled:=FALSE;
  REnbanUpdate_CheckBox.Checked:=FALSE;
  Unordered_Box.Enabled:=TRUE;
  Unordered_Box.checked:=FALSE;
end;        //Clear_ButtonClick

procedure TRecConfStat_Form.FormCreate(Sender: TObject);
begin
  With Data_Module do
  Begin
    Inv_DataSet.Filter:='';
    Inv_DataSet.Filtered:=FALSE;
    SetTodaysDate(RecConfStat_Panel);
    GetRecConfStatInfo;
    Inv_Dataset.Filtered := False;
    fSupCode:='';
    fPartNum:='';
    fKanban:='';
    fFRS:='';
    fRenban:='';
    fOrderDate:='';
    fShipDate:='';
    fCheck:=False;
    HideTerminated_CheckBox.Checked:=Data_Module.fiHideTerminated.AsBoolean;
    SortBy_ComboBox.ItemIndex:=0;
    fCheck:=True;
    RecConf_DataSource.DataSet:=Inv_DataSet;
    SelectSingleField('INV_SUPPLIER_MST', 'VC_SUPPLIER_CODE', SupplierCode_ComboBox);
    SelectSingleField('INV_PARTS_STOCK_MST', 'VC_PART_NUMBER', PartsCode_ComboBox);
    SelectSingleField('INV_PARTS_STOCK_MST', 'VC_KANBAN_NUMBER', KanbanCode_ComboBox);
    ClearControls(RecConfStat_Panel);
    ClearControls(SearchKey_GroupBox);
  End;      //With
end;        //FormCreate

procedure TRecConfStat_Form.HoldDetails(fFromGrid: Boolean);
Begin
  If fFromGrid Then
  Begin
    With RecConfStat_DBGrid do
    Begin
      Data_Module.SupplierCode  := Fields[0].AsString;
      Data_Module.PartNum       := Fields[1].AsString;
      Data_Module.FRSNo         := Fields[2].AsString;
      RecConfStat_DBGrid.Fields[3].DisplayWidth:=12;
      Data_Module.Renban        := Fields[3].AsString;

      Data_Module.fSupplierCodePrev  := Fields[0].AsString;
      Data_Module.PartNumPrev       := Fields[1].AsString;
      Data_Module.FRSNoPrev         := Fields[2].AsString;
      Data_Module.fRenbanCodePrev    := Fields[3].AsString;


      Data_Module.Quantity      := Fields[4].AsInteger;
      Data_Module.Order         := Fields[5].AsString;
      Data_Module.InTransit     := Fields[6].AsString;
      Data_Module.Arrival       := Fields[7].AsString;
      Data_Module.TrailerNo     := Fields[8].AsString;
      Data_Module.PlantYard     := Fields[9].AsString;
      Data_Module.ParkingSpot   := Fields[10].AsString;
      Data_Module.AssemblerYard := Fields[11].AsString;
      Data_module.AssemblerLocation := Fields[12].AsString;
      Data_module.EmptyTrailer  := Fields[13].AsString;
      Data_Module.Detention     := Fields[14].AsString;
      Data_Module.Warehouse     := Fields[15].AsString;
      Data_Module.Terminated    := Fields[16].AsString;
      Data_Module.Kanban        := Fields[17].AsString;
      Data_Module.Ship          := Fields[18].AsString;
    End;      //With
  End
  Else
  Begin
    With Data_Module do
    Begin
      if SupplierCode_ComboBox.Text = ' ' then
        SupplierCode := ''
      else
        SupplierCode := SupplierCode_ComboBox.Text;

      if PartsCode_ComboBox.Text = ' ' then
        PartNum := ''
      else
        PartNum := PartsCode_ComboBox.Text;

      if KanbanCode_ComboBox.Text = ' ' then
        Kanban := ''
      else
        Kanban := KanbanCode_ComboBox.Text;


      FRSNo := FRS_Edit.Text;
      Renban := Renban_Edit.Text;
      If InTransit_NUMMIBmDateEdit.Text = '' Then
        InTransit := ''
      Else
        InTransit := FormatDateTime('yyyymmdd', InTransit_NUMMIBmDateEdit.Date);
      If ArrivalDate_NUMMIBmDateEdit.Text = '' Then
        Arrival := ''
      Else
        Arrival := FormatDateTime('yyyymmdd', ArrivalDate_NUMMIBmDateEdit.Date);
      TrailerNo := TrailerNo_Edit.Text;

      If Warehouse_NUMMIBmDateEdit.Text = '' Then
        Warehouse := ''
      Else
        Warehouse := FormatDateTime('yyyymmdd', Warehouse_NUMMIBmDateEdit.Date);

      If Order_NUMMIBmDateEdit.Text = '' Then
        Order := ''
      Else
        Order := FormatDateTime('yyyymmdd', Order_NUMMIBmDateEdit.Date);

      If Terminated_NUMMIBmDateEdit.Text = '' Then
        Terminated := ''
      Else
        Terminated := FormatDateTime('yyyymmdd', Terminated_NUMMIBmDateEdit.Date);

      If Ship_NUMMIBmDateEdit.Text = '' Then
        Ship := ''
      Else
        Ship := FormatDateTime('yyyymmdd', Ship_NUMMIBmDateEdit.Date);

      If PlantYard_NUMMIBmDateEdit.Text = '' Then
        PlantYard := ''
      Else
        PlantYard := FormatDateTime('yyyymmdd', PlantYard_NUMMIBmDateEdit.Date);
      ParkingSpot := ParkingSpot_Edit.Text;

      If AssemblerYard_NUMMIBmDateEdit.Text = '' Then
        AssemblerYard := ''
      Else
        AssemblerYard := FormatDateTime('yyyymmdd', AssemblerYard_NUMMIBmDateEdit.Date);
      AssemblerLocation:=Assemblerlocation_combobox.Text;
      //wqslocation_combobox.Text := GSALocation_Edit.Text;

      If EmptyTrailer_NUMMIBmDateEdit.Text = '' Then
        EmptyTrailer := ''
      Else
        EmptyTrailer := FormatDateTime('yyyymmdd', EmptyTrailer_NUMMIBmDateEdit.Date);
      Detention := Detention_Edit.Text;


      Quantity := StrToInt(Trim(Quantity_Edit.Text));

    End;      //With
  End;
End;          //HoldDetails

procedure TRecConfStat_Form.SetDetailBoxes;
var fTempDate: String;
Begin
  With Data_Module do
  Begin
    SupplierCode_ComboBox.Text := SupplierCode;
    PartsCode_ComboBox.Text := PartNum;
    KanbanCode_ComboBox.Text := Kanban;
    FRS_Edit.Text := FRSNo;
    Renban_Edit.Text := Renban;
    If InTransit <> '' Then
    Begin
      fTempDate := InTransit;
      fTempDate := Copy(fTempDate, 5, 2) + '/' + Copy(fTempDate, 7, 2) + '/' + Copy(fTempDate, 1, 4);
      InTransit_NUMMIBmDateEdit.Date := StrToDate(fTempDate);
    End
    Else
      InTransit_NUMMIBmDateEdit.Text := '';

    If Arrival <> '' Then
    Begin
      fTempDate := Arrival;
      fTempDate := Copy(fTempDate, 5, 2) + '/' + Copy(fTempDate, 7, 2) + '/' + Copy(fTempDate, 1, 4);
      ArrivalDate_NUMMIBmDateEdit.Date := StrToDate(fTempDate);
    End
    Else
      ArrivalDate_NUMMIBmDateEdit.Text := '';

    TrailerNo_Edit.Text := TrailerNo;

    If PlantYard <> '' Then
    Begin
      fTempDate := PlantYard;
      fTempDate := Copy(fTempDate, 5, 2) + '/' + Copy(fTempDate, 7, 2) + '/' + Copy(fTempDate, 1, 4);
      PlantYard_NUMMIBmDateEdit.Date := StrToDate(fTempDate);
    End
    Else
      PlantYard_NUMMIBmDateEdit.Text := '';

    ParkingSpot_Edit.Text := ParkingSpot;

    If AssemblerYard <> '' Then
    Begin
      fTempDate := AssemblerYard;
      fTempDate := Copy(fTempDate, 5, 2) + '/' + Copy(fTempDate, 7, 2) + '/' + Copy(fTempDate, 1, 4);
      AssemblerYard_NUMMIBmDateEdit.Date := StrToDate(fTempDate);
    End
    Else
      AssemblerYard_NUMMIBmDateEdit.Text := '';

    //GSALocation_Edit.Text := GSALocation;

    Assemblerlocation_combobox.Text:= AssemblerLocation;

    If EmptyTrailer <> '' Then
    Begin
      fTempDate := EmptyTrailer;
      fTempDate := Copy(fTempDate, 5, 2) + '/' + Copy(fTempDate, 7, 2) + '/' + Copy(fTempDate, 1, 4);
      EmptyTrailer_NUMMIBmDateEdit.Date := StrToDate(fTempDate);
    End
    Else
      EmptyTrailer_NUMMIBmDateEdit.Text := '';

    If Warehouse <> '' Then
    Begin
      fTempDate := Warehouse;
      fTempDate := Copy(fTempDate, 5, 2) + '/' + Copy(fTempDate, 7, 2) + '/' + Copy(fTempDate, 1, 4);
      Warehouse_NUMMIBmDateEdit.Date := StrToDate(fTempDate);
    End
    Else
      Warehouse_NUMMIBmDateEdit.Text := '';

    If Order <> '' Then
    Begin
      fTempDate := Order;
      fTempDate := Copy(fTempDate, 5, 2) + '/' + Copy(fTempDate, 7, 2) + '/' + Copy(fTempDate, 1, 4);
      Order_NUMMIBmDateEdit.Date := StrToDate(fTempDate);
    End
    Else
      Order_NUMMIBmDateEdit.Text := '';

    If Ship <> '' Then
    Begin
      fTempDate := Ship;
      fTempDate := Copy(fTempDate, 5, 2) + '/' + Copy(fTempDate, 7, 2) + '/' + Copy(fTempDate, 1, 4);
      Ship_NUMMIBmDateEdit.Date := StrToDate(fTempDate);
    End
    Else
      Ship_NUMMIBmDateEdit.Text := '';

    If Terminated <> '' Then
    Begin
      fTempDate := Terminated;
      fTempDate := Copy(fTempDate, 5, 2) + '/' + Copy(fTempDate, 7, 2) + '/' + Copy(fTempDate, 1, 4);
      Terminated_NUMMIBmDateEdit.Date := StrToDate(fTempDate);
    End
    Else
      Terminated_NUMMIBmDateEdit.Text := '';

    Detention_Edit.Text := Detention;
    Quantity_Edit.Text := IntToStr(Quantity);

    RenbanUpdate_CheckBox.Enabled:=False;
    RenbanUpdate_CheckBox.Checked:=False;

    Unordered_Box.Enabled:=FALSE;
    Unordered_Box.Checked:=FALSE;



  End;      //With
End;          //SetDetailBoxes

function TRecConfStat_Form.SearchGrid: Boolean;
begin
  Result := False;
  try
    Data_Module.ClearControls(SearchKey_GroupBox);
    Data_Module.ClearControls(RecConfStat_Panel);

    Data_Module.Inv_Dataset.Filtered := False;
    Data_Module.Inv_Dataset.Filter := '';

    If fSupCode <> '' Then
      Data_Module.Inv_Dataset.Filter := '[Supplier] LIKE '+QuotedStr('%'+fSupCode+'%');

    if Data_Module.fiHideTerminated.AsBoolean then
      If Length(Data_Module.Inv_Dataset.Filter) > 0 Then
        Data_Module.Inv_Dataset.Filter := Data_Module.Inv_Dataset.Filter + ' AND [Terminated] = '+QuotedStr('')
      Else
        Data_Module.Inv_Dataset.Filter := '[Terminated] = '+QuotedStr('');

    If fPArtNum <> '' Then
      If Length(Data_Module.Inv_Dataset.Filter) > 0 Then
        Data_Module.Inv_Dataset.Filter := Data_Module.Inv_Dataset.Filter + ' AND [Parts] LIKE '+QuotedStr('%'+fPartNum+'%')
      Else
        Data_Module.Inv_Dataset.Filter := '[Parts] LIKE '+QuotedStr('%'+fPartNum+'%');

    If fKanban <> '' Then
      If Length(Data_Module.Inv_Dataset.Filter) > 0 Then
        Data_Module.Inv_Dataset.Filter := Data_Module.Inv_Dataset.Filter + ' AND [Kanban] LIKE '+QuotedStr('%'+fKanban+'%')
      Else
        Data_Module.Inv_Dataset.Filter := '[Kanban] LIKE '+QuotedStr('%'+fKanban+'%');

    If fFRS <>  '' Then
      If Length(Data_Module.Inv_Dataset.Filter) > 0 Then
        Data_Module.Inv_Dataset.Filter := Data_Module.Inv_Dataset.Filter + ' AND [FRSNo] LIKE '+QuotedStr('%'+fFRS+'%')
      Else
        Data_Module.Inv_Dataset.Filter := '[FRSNo] LIKE '+QuotedStr('%'+fFRS+'%');

    If fRenban <> '' Then
      If Length(Data_Module.Inv_Dataset.Filter) > 0 Then
        Data_Module.Inv_Dataset.Filter := Data_Module.Inv_Dataset.Filter + ' AND [RENBAN] LIKE '+QuotedStr('%'+fRenban+'%')
      Else
        Data_Module.Inv_Dataset.Filter := '[RENBAN] LIKE '+QuotedStr('%'+fRenban+'%');

    If (fOrderDate <> '') and (fOrderDate <> 'X') then
    begin
      If Length(Data_Module.Inv_Dataset.Filter) > 0 Then
        Data_Module.Inv_Dataset.Filter := Data_Module.Inv_Dataset.Filter + ' AND [Order] LIKE '+QuotedStr('%'+fOrderDate+'%')
      Else
        Data_Module.Inv_Dataset.Filter := '[Order] LIKE '+QuotedStr('%'+fOrderDate+'%');
    end
    else if fOrderDate = 'X' then
    begin
      fOrderDate:='';
      If Length(Data_Module.Inv_Dataset.Filter) > 0 Then
        Data_Module.Inv_Dataset.Filter := Data_Module.Inv_Dataset.Filter + ' AND [Order] LIKE '+QuotedStr('')
      Else
        Data_Module.Inv_Dataset.Filter := '[Order] LIKE '+QuotedStr('');
    end;

    If fShipDate <> '' Then
      If Length(Data_Module.Inv_Dataset.Filter) > 0 Then
        Data_Module.Inv_Dataset.Filter := Data_Module.Inv_Dataset.Filter + ' AND [Shipped] LIKE '+QuotedStr('%'+fShipDate+'%')
      Else
        Data_Module.Inv_Dataset.Filter := '[Shipped] LIKE '+QuotedStr('%'+fShipDate+'%');

    Data_Module.Inv_Dataset.Filtered := True;

    If Data_Module.Inv_Dataset.RecordCount > 0 Then
    Begin
      if SortBy_ComboBox.ItemIndex <> 0 then
        Data_Module.Inv_Dataset.Sort:=SortBy_ComboBox.Text
      else
        Data_Module.Inv_Dataset.Sort:='';

      HoldDetails(True);
      Result := True;
    End
    else
    begin
      Data_Module.Inv_Dataset.Filter:='';
      Data_Module.Inv_Dataset.Filtered := False;
      Data_Module.ClearControls(SearchKey_GroupBox);
      Data_Module.ClearControls(RecConfStat_Panel);
      SupplierCode_ComboBox.ItemIndex:=-1;
      PartsCode_ComboBox.ItemIndex:=-1;
      KanbanCode_ComboBox.ItemIndex:=-1;
      Renban_Edit.SetFocus;
      RenbanUpdate_CheckBox.Enabled:=FALSE;
      REnbanUpdate_CheckBox.Checked:=FALSE;
      Unordered_Box.Enabled:=TRUE;
      Unordered_Box.checked:=FALSE;
    end;
  except
    on e:exception do
      ShowMessage('Error in Search' + #13 + e.Message);
  end;      //try...except
End;        //SearchGrid

procedure TRecConfStat_Form.RecConfStat_DBGridKeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  If RecConfStat_DBGrid.DataSource.DataSet.RecordCount > 0 Then
  Begin
    HoldDetails(True);
    SetDetailBoxes;
  End;
end;

procedure TRecConfStat_Form.RecConfStat_DBGridMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  If RecConfStat_DBGrid.DataSource.DataSet.RecordCount > 0 Then
  Begin
    HoldDetails(True);
    SetDetailBoxes;
  End;
end;

procedure TRecConfStat_Form.InTransit_NUMMIBmDateEditExit(Sender: TObject);
var
  NummiDate: TNUMMIBmDateEdit;
begin
//Set the default TAG value to one
//Check the TAG value prior to INSERT or UPDATE...
  try
    try
      NummiDate := TNUMMIBmDateEdit(Sender);
      If NummiDate.Text <> '' Then
        If NummiDate.Date = 0 Then
        Begin
          MessageDlg('Please enter a valid date.', mtError, [mbOK], 0);
          NummiDate.SetFocus;
        End;
    except
      on e:exception do
    end;
  finally
  end;
end;

procedure TRecConfStat_Form.SupplierCode_ComboBoxChange(Sender: TObject);
begin
  If SupplierCode_ComboBox.ItemIndex >= 0 Then
  Begin
    If Trim(SupplierCode_ComboBox.Text) = '' Then
    begin
      Data_Module.SelectSingleField('INV_PART_STOCK_MST', 'VC_PART_NUMBER', PartsCode_ComboBox);
      Data_Module.SelectSingleField('INV_PART_STOCK_MST', 'VC_KANBAN_NUMBER', KanbanCode_ComboBox);
    end
    Else
    begin
      Data_Module.SelectDependantSingleField( 'SELECT_DependantPartNumber_Supplier',
                                          '@SupplierCode',
                                          'VC_PART_NUMBER',
                                          SupplierCode_ComboBox.Text,
                                          PartsCode_ComboBox);
      Data_Module.SelectDependantSingleField( 'SELECT_DependantPartNumber_Supplier',
                                          '@SupplierCode',
                                          'VC_KANBAN_NUMBER',
                                          SupplierCode_ComboBox.Text,
                                          KanbanCode_ComboBox);
    end;
  End;      //If
end;

procedure TRecConfStat_Form.PartsCode_ComboBoxChange(Sender: TObject);
begin
  If PartsCode_ComboBox.ItemIndex >= 0 Then
  Begin
    If Trim(PartsCode_ComboBox.Text) <> '' Then
    begin
      Data_Module.SelectDependantSingleField( 'SELECT_DependantKanbanNumber_PartNumber',
                                          '@PartNumber',
                                          'VC_KANBAN_NUMBER',
                                          PartsCode_ComboBox.Text,
                                          KanbanCode_ComboBox);
    end;
  End;      //If

end;

procedure TRecConfStat_Form.TextChange(Sender: TObject);
begin
  If Sender.ClassName = 'TMaskEdit' Then
    If Length(Trim(TMaskEdit(Sender).Text)) < 1 Then
      TMaskEdit(Sender).Text := '0';

end;

procedure TRecConfStat_Form.MaskEditExit(Sender: TObject);
var
  fPos: Integer;
  fTempValue: String;
begin
  If Sender.ClassName = 'TMaskEdit' Then
  Begin
    fTempValue := Trim(TMaskEdit(Sender).Text);
    fPos := Pos(' ', fTempValue);
    While fPos <> 0 do
    Begin
      fTempValue := Copy(fTempValue, 1, fPos - 1) + Copy(fTempValue, fPos + 1, Length(fTempValue));
      fPos := Pos(' ', fTempValue);
    End;
    TMaskEdit(Sender).Text := fTempValue;
  End;
end;

procedure TRecConfStat_Form.RecConf_DataSourceDataChange(Sender: TObject;
  Field: TField);
begin
  HoldDetails(True);
  SetDetailBoxes;
end;

procedure TRecConfStat_Form.FormShow(Sender: TObject);
begin
  SearchGrid;
  RenbanUpdate_CheckBox.Enabled:=FALSE;
  RenbanUpdate_CheckBox.Checked:=FALSE;
  Unordered_Box.Enabled:=TRUE;
  Unordered_Box.checked:=FALSE;
  Data_Module.ClearControls(RecConfStat_Panel);
  Data_Module.ClearControls(SearchKey_GroupBox);
  Renban_Edit.SetFocus;

  AssemblerLocation_Label.Caption:=DAta_Module.fiAssemblerName.AsString+' Location';
  AssemblerYard_Label.Caption:=DAta_Module.fiAssemblerName.AsString+' Yard';
  PlantYard_Label.Caption:=Data_Module.fiPlantName.AsString+' Yard';

  //SetDetailBoxes;
  Clear_ButtonClick(self);
end;

function  TRecConfStat_Form.Validate:boolean;
var
  i:integer;
begin
  result:=False;
  //confirm data
  if length(SupplierCode_ComboBox.Text) < 5 then
  begin
    ShowMessage('Supplier Code must be 5 characters');
    Renban_Edit.SetFocus;
    exit;
  end;

  if length(PartsCode_ComboBox.Text) < 12 then
  begin
    ShowMessage('Parts Code must be 12 characters');
    PartsCode_ComboBox.SetFocus;
    exit;
  end;

  if length(FRS_Edit.Text) < 7 then
  begin
    ShowMessage('FRS number must be 7 characters');
    FRS_Edit.SetFocus;
    exit;
  end;

  if not TryStrToInt(Trim(Quantity_Edit.Text),i) then
  begin
    ShowMessage('Quantity must be a numeric');
    Quantity_Edit.SetFocus;
    exit;
  end;

  if InTransit_NUMMIBmDateEdit.Text = '' then
  begin
    if ArrivalDate_NUMMIBmDateEdit.Text <> '' then
    begin
      ShowMessage('Order must be marked In Transit when arrival is set');
      ArrivalDate_NUMMIBmDateEdit.SetFocus;
      exit;
    end;

    if Warehouse_NUMMIBmDateEdit.Text <> '' then
    begin
      ShowMessage('Order must be marked In Transit when warehouse is set');
      Warehouse_NUMMIBmDateEdit.SetFocus;
      exit;
    end;

    if PlantYard_NUMMIBmDateEdit.Text <> '' then
    begin
      ShowMessage('Order must be marked In Transit when plant yard is set');
      PlantYard_NUMMIBmDateEdit.SetFocus;
      exit;
    end;

    if AssemblerYard_NUMMIBmDateEdit.Text <> '' then
    begin
      ShowMessage('Order must be marked In Transit when assembler yard is set');
      AssemblerYard_NUMMIBmDateEdit.SetFocus;
      exit;
    end;
  end;

  result:=true;
end;

procedure TRecConfStat_Form.ASSEMBLERLocation_ComboBoxChange(Sender: TObject);
begin
  ParkingSpot_Edit.Text:='';
  AssemblerYard_NUMMIBmDateEdit.Date:=Data_Module.GetLastProductionDate;
end;

procedure TRecConfStat_Form.TrailerNo_EditChange(Sender: TObject);
begin
  RenbanUpdate_CheckBox.Enabled:=True;
  if Sender is TComboBox then
  begin
    if TComboBox(Sender).Name = 'AssemblerLocation_ComboBox' then
      ASSEMBLERLocation_ComboBoxChange(sender);
  end;

  {if Sender is TNUMMIBMDateEdit then
  begin
    if TNUMMIBMDateEdit(Sender).Name = 'Terminated_NUMMIBmDateEdit' then
    begin
      if EmptyTrailer_NUMMIBmDateEdit.Text = '' then
        EmptyTrailer_NUMMIBmDateEdit.Date:=TNUMMIBMDateEdit(Sender).Date;
    end;
  end;}
end;


procedure TRecConfStat_Form.HideTerminated_CheckBoxClick(Sender: TObject);
begin
  if fCheck then
  begin
    Data_Module.fiHideTerminated.AsBoolean:=HideTerminated_CheckBox.Checked;
    SearchGrid;
    Data_Module.ClearControls(RecConfStat_Panel);
    Data_Module.ClearControls(SearchKey_GroupBox);
    SupplierCode_ComboBox.ItemIndex:=-1;
    PartsCode_ComboBox.ItemIndex:=-1;
    KanbanCode_ComboBox.ItemIndex:=-1;
    RenbanUpdate_CheckBox.Enabled:=FALSE;
    REnbanUpdate_CheckBox.Checked:=FALSE;
    Unordered_Box.Enabled:=TRUE;
    Unordered_Box.checked:=FALSE;
    Renban_Edit.SetFocus;
  end;
end;

procedure TRecConfStat_Form.SortBy_ComboBoxChange(Sender: TObject);
begin
  SearchGrid;
  Data_Module.ClearControls(RecConfStat_Panel);
  Data_Module.ClearControls(SearchKey_GroupBox);
  Renban_Edit.SetFocus;
end;


end.
