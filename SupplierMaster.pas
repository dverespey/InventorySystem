//***********************************************************
//  InventorySystem
//
//  Copyright (c) 2002 Failproof Manufacturing Systems
//
//***********************************************************
//
//  10/25/2002  Aaron Huge  Initial creation

unit SupplierMaster;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Grids, DBGrids, DataModule, Buttons, DB, Filectrl;

type
  TSupplierObj = class(TTwoFieldObj)
  private
    fAddress: String;
    fTelephone: String;
    fFax: String;
    fContact: String;
  public
    constructor Create (); overload;
    destructor Destroy; override;
    property Address: String read fAddress write fAddress;
    property Telephone: String read fTelephone write fTelephone;
    property Fax: String read fFax write fFax;
    property Contact: String read fContact write fContact;
end;


type
  TSupplierMaster_Form = class(TForm)
    ManagementButtons_Panel: TPanel;
    Insert_Button: TButton;
    Update_Button: TButton;
    Search_Button: TButton;
    Clear_Button: TButton;
    Close_Button: TButton;
    Delete_Button: TButton;
    SupplierMaster_Label: TLabel;
    SupplierMaster_Panel: TPanel;
    SupplierCode_Label: TLabel;
    Address_Label: TLabel;
    PhoneNum_Label: TLabel;
    Person_Label: TLabel;
    SupplierCode_Edit: TEdit;
    Address_Edit: TEdit;
    PhoneNum_Edit: TEdit;
    Person_Edit: TEdit;
    SupplierName_Label: TLabel;
    SupplierName_Edit: TEdit;
    FaxNum_Label: TLabel;
    FaxNum_Edit: TEdit;
    Label1: TLabel;
    Directory_Edit: TEdit;
    Label2: TLabel;
    Email_Edit: TEdit;
    Label3: TLabel;
    City_Edit: TEdit;
    Label4: TLabel;
    State_Edit: TEdit;
    Label5: TLabel;
    Zip_Edit: TEdit;
    Logistics_ComboBox: TComboBox;
    Breakdown_SpeedButton: TSpeedButton;
    Label6: TLabel;
    OutputFileType_RadioGroup: TRadioGroup;
    Label7: TLabel;
    OrderFileTimestamp_CheckBox: TCheckBox;
    SupplierMaster_DBGrid: TDBGrid;
    Supplier_DataSource: TDataSource;
    SiteNumberinOrder_CheckBox: TCheckBox;
    Label8: TLabel;
    CreateOrderSheet_ComboBox: TComboBox;
    Label9: TLabel;
    InventoryAddPoint_ComboBox: TComboBox;
    procedure FormCreate(Sender: TObject);
    procedure Search_ButtonClick(Sender: TObject);
    procedure Insert_ButtonClick(Sender: TObject);
    procedure Clear_ButtonClick(Sender: TObject);
    procedure Update_ButtonClick(Sender: TObject);
    procedure Delete_ButtonClick(Sender: TObject);
    procedure Breakdown_SpeedButtonClick(Sender: TObject);
    procedure SupplierMaster_DBGridKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure SupplierMaster_DBGridMouseUp(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure Supplier_DataSourceDataChange(Sender: TObject;
      Field: TField);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
    procedure SetDetailBoxes;
    procedure HoldDetails(fFromGrid: Boolean);
    function SearchGrid(fSupCode: String): Boolean;
    function Validate:boolean;
  public
    { Public declarations }
    function Execute: boolean;
  end;

var
  SupplierMaster_Form: TSupplierMaster_Form;

implementation


{$R *.dfm}
constructor TSupplierObj.Create ();
Begin
 inherited create;
End;

destructor TSupplierObj.Destroy;
Begin
End;

function TSupplierMaster_Form.Execute: boolean;
Begin
  Result:= True;
  try
    try
      ShowModal;
    except
      On E:Exception do
        showMessage('Unable to generate Supplier Master screen.'
          + #13 + 'ERROR:'
          + #13 + E.Message);
    end;
  finally
    If ModalResult = mrCancel Then
      Result:= False;
  end;
end;    //Execute

procedure TSupplierMaster_Form.FormCreate(Sender: TObject);
begin
  Data_Module.Inv_DataSet.Filter:='';
  Data_Module.Inv_DataSet.Filtered:=FALSE;
  Data_Module.SelectSingleField('INV_LOGISTICS_MST', 'VC_LOGISTICS_NAME', Logistics_ComboBox);
  Data_Module.SelectSingleField('INV_PART_TYPE_MST', 'VC_PART_TYPE', CreateOrderSheet_ComboBox);
  Data_Module.SelectSingleField('INV_ADD_POINT_INF', 'VC_ADD_POINT', InventoryAddPoint_ComboBox);
  Data_Module.GetSupplierInfo;
  Supplier_DataSource.DataSet:=Data_Module.Inv_DataSet;
  SupplierCode_Edit.Text := '';
  Data_Module.ClearControls(SupplierMaster_Panel);
end;

procedure TSupplierMaster_Form.Insert_ButtonClick(Sender: TObject);
begin
  if Validate then
  begin
    HoldDetails(False);
    If Not Data_Module.InsertSupplierInfo Then
      MessageDlg('Unable to INSERT ' + Data_Module.SupplierName + '('+ Data_Module.SupplierCode +')', mtInformation, [mbOk], 0);

    Data_Module.GetSupplierInfo;
    Data_Module.Inv_DataSet.Locate('Supplier Code',SupplierCode_Edit.Text,[]);
    SetDetailBoxes;
    SupplierCode_Edit.SelectAll;
    SupplierCode_Edit.SetFocus;
  end;
end;

function TSupplierMaster_Form.Validate:boolean;
begin
  result:=True;

  with SupplierCode_Edit do
  begin
    if length(text) < 5 then
    begin
      ShowMessage('Supplier Code must be 5 characters');
      SetFocus;
      result:=False;
      exit;
    end;
  end;
end;

procedure TSupplierMaster_Form.Update_ButtonClick(Sender: TObject);
var
  SupCode:string;
begin
  if Validate then
  begin
    HoldDetails(False);
    with Data_Module do
    begin
      SupCode:=SupplierCode_Edit.Text;
      UpdateSupplierInfo;
      GetSupplierInfo;
      Inv_DataSet.Locate('Supplier Code',SupCode,[lopartialkey]);
    end;
    SetDetailBoxes;
  end;
end;

procedure TSupplierMaster_Form.Delete_ButtonClick(Sender: TObject);
begin
  HoldDetails(False);
  If MessageDlg('Are you sure you wish to delete' + #13 +
                 Data_Module.SupplierName + ' (' + Data_Module.SupplierCode + ') from the database?',
                 mtWarning, [mbYes, mbNo], 0) = mrYes Then
  Begin
    Data_Module.DeleteSupplierInfo;
    SupplierCode_Edit.Text := '';
    Data_Module.GetSupplierInfo;
    SearchGrid(Data_Module.SupplierCode);
  End;

  SupplierCode_Edit.SelectAll;
  SupplierCode_Edit.SetFocus;

end;

procedure TSupplierMaster_Form.Search_ButtonClick(Sender: TObject);
var
  fFound: Boolean;
Begin

  fFound := SearchGrid(SupplierCode_Edit.Text);
  If fFound Then
  Begin
    SetDetailBoxes;
  End
  Else
  Begin
    ShowMessage('No matches were found for your query.');
  End;

  SupplierCode_Edit.SelectAll;
  SupplierCode_Edit.SetFocus;
end;

procedure TSupplierMaster_Form.Clear_ButtonClick(Sender: TObject);
begin
  SupplierCode_Edit.Text := '';
  SupplierCode_Edit.SetFocus;
  Data_Module.ClearControls(SupplierMaster_Panel);
end;

function TSupplierMaster_Form.SearchGrid(fSupCode: String): Boolean;
var
  fBookmark: String;
begin
  Result := False;
  Data_Module.ClearControls(SupplierMaster_Panel);

  try
    With Data_Module.Inv_DataSet do
    Begin
      fBookmark := Bookmark;
      First;
      While Not EOF do
      Begin
        If fSupCode = Trim(FieldByName('SUPPLIER CODE').AsString) Then
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

procedure TSupplierMaster_Form.SetDetailBoxes;
Begin
  try
    With Data_Module do
    Begin
      SupplierCode_Edit.Text := SupplierCode;
      SupplierName_Edit.Text := SupplierName;
      Address_Edit.Text := SupplierAddress;
      PhoneNum_Edit.Text := SupplierTelephone;
      FaxNum_Edit.Text := SupplierFax;
      Person_Edit.Text := SupplierPerson;
      Directory_Edit.Text:=SupplierDirectory;
      SearchCombo(Logistics_ComboBox, LogisticsName);
      City_Edit.Text:=SupplierCity;
      State_Edit.Text:=SupplierState;
      Zip_edit.Text:=SupplierZip;
      Email_Edit.Text:=SupplierEmail;

      if SupplierOutputFileType = 'TEXT' then
        OutputFileType_RadioGroup.ItemIndex:=0
      else if SupplierOutputFileType = 'EXCEL' then
        OutputFileType_RadioGroup.ItemIndex:=1
      else
        OutputFileType_RadioGroup.ItemIndex:=2;

      OrderFileTimestamp_CheckBox.checked:=OrderFileTimestamp;
      SiteNumberInOrder_CheckBox.Checked:=SiteNumberInOrder;
      SearchCombo(CreateOrderSheet_ComboBox, CreateOrderSheet);
      SearchCombo(InventoryAddPoint_ComboBox, InvAddPoint);

    End;      //With
  except
    on e:exception do
      ShowMessage('Error in SetDetailBoxes' + #13 + e.Message);
  end;      //try...except
End;          //SetDetailBoxes

procedure TSupplierMaster_Form.HoldDetails(fFromGrid: Boolean);
Begin
  If fFromGrid Then
  Begin
    with SupplierMaster_DBGrid.DataSource.DataSet do
    Begin
      Data_Module.SupplierCode            := Fields[0].AsString;
      Data_Module.SupplierName            := Fields[1].AsString;
      Data_Module.SupplierAddress         := Fields[2].AsString;
      Data_Module.SupplierCity            := Fields[3].AsString;
      Data_Module.SupplierState           := Fields[4].AsString;
      Data_Module.SupplierZip             := Fields[5].AsString;
      Data_Module.SupplierTelephone       := Fields[6].AsString;
      Data_Module.SupplierFax             := Fields[7].AsString;
      Data_Module.SupplierPerson          := Fields[8].AsString;
      Data_Module.SupplierEmail           := Fields[9].AsString;
      Data_Module.SupplierDirectory       := Fields[10].AsString;
      Data_Module.LogisticsName           := Fields[11].AsString;
      Data_Module.SupplierOutputFileType  := Fields[12].AsString;
      Data_Module.OrderFileTimestamp      := Fields[13].AsBoolean;
      Data_Module.SiteNumberInOrder       := Fields[14].AsBoolean;
      Data_Module.CreateOrderSheet        := Fields[15].AsString;
      Data_Module.InvAddPoint              := Fields[16].AsString;
      Data_Module.RecordID                := Fields[17].AsInteger;
      Fields[17].Visible:=FALSE;
      Fields[18].Visible:=FALSE;
    End;      //With
  End
  Else
  Begin
    With Data_Module do
    Begin
      SupplierCode      := SupplierCode_Edit.Text;
      SupplierName      := SupplierName_Edit.Text;
      SupplierAddress   := Address_Edit.Text;
      SupplierCity      := City_Edit.Text;
      SupplierState     := State_Edit.Text;
      SupplierZip       := Zip_edit.Text;
      SupplierTelephone := PhoneNum_Edit.Text;
      SupplierFax       := FaxNum_Edit.Text;
      SupplierPerson    := Person_edit.Text;
      SupplierEmail     := Email_Edit.Text;
      SupplierDirectory := Directory_Edit.Text;

      // dose this due to empty string bug
      if (Logistics_ComboBox.Text=' ') or (Logistics_ComboBox.Text='') then
        LogisticsName := ''
      else
        LogisticsName := Logistics_ComboBox.text;

      If OutputFileType_RadioGroup.ItemIndex = 0 Then
        SupplierOutputFileType := 'T'
      Else If OutputFileType_RadioGroup.ItemIndex = 1 then
        SupplierOutputFileType := 'E'
      else
        SupplierOutputFileType := 'B';

      OrderFileTimestamp := OrderFileTimestamp_CheckBox.checked;
      SiteNumberInOrder := SiteNumberInOrder_CheckBox.Checked;
      CreateOrderSheet := CreateOrderSheet_ComboBox.Text;
      InvAddPoint := InventoryAddPoint_ComboBox.Text;

    End;      //With
  End;
End;          //HoldDetails


procedure TSupplierMaster_Form.Breakdown_SpeedButtonClick(Sender: TObject);
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

procedure TSupplierMaster_Form.SupplierMaster_DBGridKeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  HoldDetails(True);
  SetDetailBoxes;

end;

procedure TSupplierMaster_Form.SupplierMaster_DBGridMouseUp(
  Sender: TObject; Button: TMouseButton; Shift: TShiftState; X,
  Y: Integer);
begin
  HoldDetails(True);
  SetDetailBoxes;
end;

procedure TSupplierMaster_Form.Supplier_DataSourceDataChange(
  Sender: TObject; Field: TField);
begin
  HoldDetails(True);
  SetDetailBoxes;
end;

procedure TSupplierMaster_Form.FormShow(Sender: TObject);
begin
  SupplierCode_Edit.SetFocus;
  HoldDetails(True);
  SetDetailBoxes;
end;

end.
