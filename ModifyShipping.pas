//****************************************************************
//
//       Inventory Control
//
//       Copyright (c) 2002-2007 TAI-Failproof Manufacturing Systems.
//
//****************************************************************
//
//  05/31/2007  David Verespey      Initial creation

unit ModifyShipping;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, ExtCtrls, Grids, DBGrids, DB;

type
  TModifyShipping_Form = class(TForm)
    ManagementButtons_Panel: TPanel;
    Insert_Button: TButton;
    Update_Button: TButton;
    Search_Button: TButton;
    Clear_Button: TButton;
    Close_Button: TButton;
    Delete_Button: TButton;
    GroupBox1: TGroupBox;
    Label3: TLabel;
    Date_Label: TLabel;
    Line_ComboBox: TComboBox;
    Production_DateTimePicker: TDateTimePicker;
    StartSeqNo_Label: TLabel;
    Label4: TLabel;
    StartSeqNo_Edit: TEdit;
    LastSeqNo_Edit: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    Qty_Edit: TEdit;
    ModifyShipping_DataSource: TDataSource;
    ModifyShipping_DBGrid: TDBGrid;
    PartNumber_ComboBox: TComboBox;
    procedure Close_ButtonClick(Sender: TObject);
    procedure Update_ButtonClick(Sender: TObject);
    procedure ModifyShipping_DataSourceDataChange(Sender: TObject;
      Field: TField);
    procedure FormShow(Sender: TObject);
    procedure Clear_ButtonClick(Sender: TObject);
    procedure Insert_ButtonClick(Sender: TObject);
  private
    { Private declarations }
    InUpdate:boolean;
    fLineName:string;
    fProductionDate:TDateTime;
    procedure SetDetailBoxes;
    function HoldDetails(fFromGrid: Boolean):boolean;
  public
    { Public declarations }
    procedure Execute;

    property LineName:string
    read fLineName
    write fLineName;

    property ProductionDate:TDateTime
    read fProductionDate
    write fProductionDate;
  end;

var
  ModifyShipping_Form: TModifyShipping_Form;

implementation

uses DataModule;

{$R *.dfm}

function TModifyShipping_Form.HoldDetails(fFromGrid: Boolean):boolean;
var
  fr:integer;
Begin
  result:=TRUE;
  If fFromGrid Then
  Begin
    with Data_Module.Inv_DataSet do
    begin
      Data_Module.RecordID        := Fields[0].AsInteger;
      Fields[0].Visible:=FALSE;
      Fields[1].Visible:=FALSE;
      Data_Module.PartNum         := Fields[2].AsString;
      Fields[2].Displaylabel:='Part Number';
      Fields[3].Visible:=FALSE;
      Data_Module.Quantity        := Fields[4].AsInteger;
      Fields[4].Displaylabel:='Qty';
      Fields[5].Visible:=FALSE;
    end;
  End
  Else
  Begin
    With Data_Module do
    Begin
      if (PartNumber_ComboBox.Text = ' ') or  (PartNumber_ComboBox.Text = '') then
      begin
        ShowMessage('Part Number invalid');
        PartNumber_ComboBox.SetFocus;
        result:=FALSE;
        exit;
      end;
      Partnum:=PartNumber_ComboBox.Text;

      if not TryStrToInt(Qty_Edit.Text, fr) then
      begin
        ShowMessage('Quantity invalid');
        Qty_Edit.SetFocus;
        result:=FALSE;
        exit;
      end;
      Quantity:=fr;
    End;      //With
  End;
End;          //HoldDetails

procedure TModifyShipping_Form.SetDetailBoxes;
begin
  With Data_Module do
  Begin
    SearchCombo(PartNumber_ComboBox, PartNum);
    Qty_Edit.Text:=IntToStr(Quantity);
  end;
end;

procedure TModifyShipping_Form.Execute;
begin

  try
    InUpdate:=FALSE;
    with Data_Module do
    begin
      ProductionDate:=formatdatetime('yyyymmdd',fProductiondate);
      LineName:=fLineName;

      GetShippingInfo;

      StartSeqNo_Edit.Text := Inv_DataSet.FieldByName('VC_START_SEQ_NUMBER').AsString;
      LastSeqNo_Edit.Text := Inv_DataSet.FieldByName('VC_END_SEQ_NUMBER').AsString;
      Production_DateTimePicker.Date:=fProductionDate;
      Line_ComboBox.Items.Clear;
      Line_ComboBox.Items.Add(fLineName);
      Line_ComboBox.ItemIndex:=0;

      RecordID:=Inv_DataSet.FieldByName('IN_SHIPPING_ID').AsInteger;

      Inv_DataSet.Filter:='';
      GetShippingInfoDetail;
      Inv_DataSet.Filtered:=FALSE;
      ModifyShipping_DataSource.DataSet:=Inv_DataSet;

      SelectSingleField('INV_PARTS_STOCK_MST', 'VC_PART_NUMBER', PartNumber_ComboBox);


      SetDetailBoxes;
    end;

    ShowModal;
  except
    on e:exception do
    begin
    end;
  end;
end;

procedure TModifyShipping_Form.Close_ButtonClick(Sender: TObject);
begin
  Close;
end;

procedure TModifyShipping_Form.Update_ButtonClick(Sender: TObject);
var
  Id:integer;
begin
  // Update Record
  InUpdate:=TRUE;
  if HoldDetails(False) then
  begin
    with Data_Module do
    begin
      Id:=RecordID;
      UpdateShippingInfoDetail;
      GetShippingInfo;
      RecordID:=Inv_DataSet.FieldByName('IN_SHIPPING_ID').AsInteger;
      InUpdate:=FALSE;
      GetShippingInfoDetail;
      Data_Module.Inv_DataSet.Locate('IN_PART_SHIPPING_ID',ID,[]);
    end;
    SetDetailBoxes;
    Qty_Edit.SetFocus;
  end;
end;

procedure TModifyShipping_Form.ModifyShipping_DataSourceDataChange(
  Sender: TObject; Field: TField);
begin
  if not InUpdate then
  begin
    HoldDetails(True);
    SetDetailBoxes;
    Update_Button.Visible:=TRUE;
    Insert_button.Visible:=FALSE;
    Clear_button.Visible:=TRUE;
  end;
end;

procedure TModifyShipping_Form.FormShow(Sender: TObject);
begin
  Qty_Edit.SetFocus;
end;

procedure TModifyShipping_Form.Clear_ButtonClick(Sender: TObject);
begin
  Update_Button.Visible:=FALSE;
  Insert_button.Visible:=TRUE;
  Clear_button.Visible:=FALSE;

  PartNumber_ComboBox.ItemIndex:=0;
  Qty_Edit.Text:='';
end;

procedure TModifyShipping_Form.Insert_ButtonClick(Sender: TObject);
begin
  // Update Record
  InUpdate:=TRUE;
  if HoldDetails(False) then
  begin
    with Data_Module do
    begin
      GetShippingInfo;
      RecordID:=Inv_DataSet.FieldByName('IN_SHIPPING_ID').AsInteger;
      InsertShippingInfoDetail;
      InUpdate:=FALSE;
      GetShippingInfoDetail;
    end;
    SetDetailBoxes;
    Qty_Edit.SetFocus;
  end;
end;

end.
