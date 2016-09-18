unit ASNInvoice;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, Grids, DBGrids, StdCtrls, ExtCtrls, ADODB, Buttons;

type
  TFormMode = (fmASN, fmINV);
  TASNInvoice_Form = class(TForm)
    CarTruckShip_Label: TLabel;
    ASNListDataSource: TDataSource;
    List_GroupBox: TGroupBox;
    ListDBGrid: TDBGrid;
    ASNStatus_ComboBox: TComboBox;
    Label1: TLabel;
    Items_GroupBox: TGroupBox;
    ItemsDBGrid: TDBGrid;
    Buttons_Panel: TPanel;
    Insert_Button: TButton;
    Update_Button: TButton;
    Search_Button: TButton;
    Clear_Button: TButton;
    Close_Button: TButton;
    Delete_Button: TButton;
    ASNList_DataSet: TADODataSet;
    ASNItems_DataSet: TADODataSet;
    ASNItemsDataSource: TDataSource;
    Label2: TLabel;
    ASNorInvoice_ComboBox: TComboBox;
    Invoice_DataSource: TDataSource;
    InvoiceList_DataSet: TADODataSet;
    INVOICEItems_DataSet: TADODataSet;
    INVOICEItemsDataSource: TDataSource;
    ASNItem_Box: TGroupBox;
    ManifestNumber_Edit: TEdit;
    AssemblyPartNumber_Edit: TEdit;
    Qty_Edit: TEdit;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    ASNItem_Command: TADOCommand;
    DeleteASN_Button: TButton;
    UnsendASN_Button: TButton;
    AssemblyPartNumber_Combo: TComboBox;
    AssyManifest_DataSet: TADODataSet;
    Cancel_Button: TButton;
    Label6: TLabel;
    DataView: TComboBox;
    RecreateFile_Button: TButton;
    SearchEdit: TEdit;
    SpeedButton1: TSpeedButton;
    procedure FormShow(Sender: TObject);
    procedure ASNListDataSourceDataChange(Sender: TObject; Field: TField);
    procedure ASNStatus_ComboBoxChange(Sender: TObject);
    procedure ASNorInvoice_ComboBoxChange(Sender: TObject);
    procedure Invoice_DataSourceDataChange(Sender: TObject; Field: TField);
    procedure ASNItemsDataSourceDataChange(Sender: TObject; Field: TField);
    procedure Clear_ButtonClick(Sender: TObject);
    procedure Delete_ButtonClick(Sender: TObject);
    procedure Update_ButtonClick(Sender: TObject);
    procedure DeleteASN_ButtonClick(Sender: TObject);
    procedure UnsendASN_ButtonClick(Sender: TObject);
    procedure Insert_ButtonClick(Sender: TObject);
    procedure AssemblyPartNumber_ComboChange(Sender: TObject);
    procedure Cancel_ButtonClick(Sender: TObject);
    procedure RecreateFile_ButtonClick(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
  private
    { Private declarations }
    fFormMode:TFormMode;
    procedure GetASNs;
    procedure GetINVOICEs;
  public
    { Public declarations }
    procedure Execute;
  end;

var
  ASNInvoice_Form: TASNInvoice_Form;

implementation

uses DataModule,EDI810Object, EDI856Object;

{$R *.dfm}

procedure TASNInvoice_Form.GetINVOICEs;
begin
  try

    InvoiceList_DataSet.Close;
    case ASNStatus_ComboBox.ItemIndex of
      0:
      begin
        InvoiceList_DataSet.Parameters.Clear;
        InvoiceList_DataSet.Parameters.AddParameter.Name := '@List';
        InvoiceList_DataSet.Parameters.ParamValues['@List']:='X';
        InvoiceList_DataSet.Parameters.AddParameter.Name := '@Range';
        InvoiceList_DataSet.Parameters.ParamValues['@Range']:=DataView.ItemIndex;
        InvoiceList_DataSet.Open;
        ASNItem_Box.Visible:=FALSE;
        Insert_Button.Visible:=FALSE;
        Update_Button.Visible:=FALSE;
        Delete_Button.Visible:=FALSE;
        DeleteASN_Button.Visible:=FALSE;
        RecreateFile_Button.Visible:=FALSE;
        UnsendASN_Button.Visible:=FALSE;
        Clear_Button.Visible:=FALSE;
        Cancel_Button.Visible:=FALSE;
        AssemblyPartNumber_Combo.Visible:=FALSE;
        AssemblyPartNumber_Edit.Visible:=TRUE;
      end;
      1:
      begin
        ASNStatus_ComboBox.ItemIndex:=0;
        InvoiceList_DataSet.Parameters.Clear;
        InvoiceList_DataSet.Parameters.AddParameter.Name := '@List';
        InvoiceList_DataSet.Parameters.ParamValues['@List']:='X';
        InvoiceList_DataSet.Parameters.AddParameter.Name := '@Range';
        InvoiceList_DataSet.Parameters.ParamValues['@Range']:=DataView.ItemIndex;
        InvoiceList_DataSet.Open;
        ASNItem_Box.Visible:=FALSE;
        Insert_Button.Visible:=FALSE;
        Update_Button.Visible:=FALSE;
        Delete_Button.Visible:=FALSE;
        DeleteASN_Button.Visible:=FALSE;
        RecreateFile_Button.Visible:=FALSE;
        UnsendASN_Button.Visible:=FALSE;
        Clear_Button.Visible:=FALSE;
        Cancel_Button.Visible:=FALSE;
        AssemblyPartNumber_Combo.Visible:=FALSE;
        AssemblyPartNumber_Edit.Visible:=TRUE;
      end;
      2:
      begin
        InvoiceList_DataSet.Parameters.Clear;
        InvoiceList_DataSet.Parameters.AddParameter.Name := '@List';
        InvoiceList_DataSet.Parameters.ParamValues['@List']:='S';
        InvoiceList_DataSet.Parameters.AddParameter.Name := '@Range';
        InvoiceList_DataSet.Parameters.ParamValues['@Range']:=DataView.ItemIndex;
        InvoiceList_DataSet.Open;
        ASNItem_Box.Visible:=FALSE;
        Insert_Button.Visible:=FALSE;
        Update_Button.Visible:=FALSE;
        Delete_Button.Visible:=FALSE;
        DeleteASN_Button.Visible:=FALSE;
        RecreateFile_Button.Visible:=FALSE;
        UnsendASN_Button.Visible:=TRUE;
        UnsendASN_Button.Caption:='&Unsend INV';
        AssemblyPartNumber_Combo.Visible:=FALSE;
        Clear_Button.Visible:=FALSE;
        Cancel_Button.Visible:=FALSE;
        AssemblyPartNumber_Combo.Visible:=FALSE;
        AssemblyPartNumber_Edit.Visible:=TRUE;
      end;
      3:
      begin
        InvoiceList_DataSet.Parameters.Clear;
        InvoiceList_DataSet.Parameters.AddParameter.Name := '@List';
        InvoiceList_DataSet.Parameters.ParamValues['@List']:='A';
        InvoiceList_DataSet.Parameters.AddParameter.Name := '@Range';
        InvoiceList_DataSet.Parameters.ParamValues['@Range']:=DataView.ItemIndex;
        InvoiceList_DataSet.Open;
        ASNItem_Box.Visible:=FALSE;
        Insert_Button.Visible:=FALSE;
        Update_Button.Visible:=FALSE;
        Delete_Button.Visible:=FALSE;
        DeleteASN_Button.Visible:=FALSE;
        RecreateFile_Button.Visible:=TRUE;
        UnsendASN_Button.Visible:=TRUE;
        UnsendASN_Button.Caption:='&Unsend INV';
        Clear_Button.Visible:=FALSE;
        Cancel_Button.Visible:=FALSE;
        AssemblyPartNumber_Combo.Visible:=FALSE;
        AssemblyPartNumber_Edit.Visible:=TRUE;
      end;
      4:
      begin
        InvoiceList_DataSet.Parameters.Clear;
        InvoiceList_DataSet.Parameters.AddParameter.Name := '@List';
        InvoiceList_DataSet.Parameters.ParamValues['@List']:='R';
        InvoiceList_DataSet.Parameters.AddParameter.Name := '@Range';
        InvoiceList_DataSet.Parameters.ParamValues['@Range']:=DataView.ItemIndex;
        InvoiceList_DataSet.Open;
        ASNItem_Box.Visible:=FALSE;
        Insert_Button.Visible:=FALSE;
        Update_Button.Visible:=FALSE;
        Delete_Button.Visible:=FALSE;
        DeleteASN_Button.Visible:=FALSE;
        RecreateFile_Button.Visible:=FALSE;
        UnsendASN_Button.Visible:=FALSE;
        Clear_Button.Visible:=FALSE;
        Cancel_Button.Visible:=FALSE;
        AssemblyPartNumber_Combo.Visible:=FALSE;
        AssemblyPartNumber_Edit.Visible:=TRUE;
      end;
    end;
    InvoiceList_DataSet.Fields[0].Visible:=FALSE;
    ListDBGrid.DataSource:=Invoice_DataSource;
    List_GroupBox.Caption:='INVOICE List';
    Items_GroupBox.Caption:= 'INVOICE Items';
    fFormMode:=fmINV;
  except
    on e:exception do
    begin
      ShowMessage('Exception: unable to get INVOICE list, '+e.Message);
      Data_Module.LogActLog('ERROR','Exception: unable to get INVOICE list, '+e.Message);
      Close;
    end;
  end;
end;

procedure TASNInvoice_Form.GetASNs;
begin
  try
    ASNList_DataSet.Close;
    case ASNStatus_ComboBox.ItemIndex of
      0:
      begin
        ASNList_DataSet.Parameters.Clear;
        ASNList_DataSet.Parameters.AddParameter.Name := '@List';
        ASNList_DataSet.Parameters.ParamValues['@List']:='X';
        ASNList_DataSet.Parameters.AddParameter.Name := '@Range';
        ASNList_DataSet.Parameters.ParamValues['@Range']:=DataView.ItemIndex;
        ASNList_DataSet.Open;
        ASNItem_Box.Visible:=FALSE;
        Insert_Button.Visible:=FALSE;
        Update_Button.Visible:=FALSE;
        Delete_Button.Visible:=FALSE;
        DeleteASN_Button.Visible:=FALSE;
        RecreateFile_Button.Visible:=FALSE;
        UnsendASN_Button.Visible:=FALSE;
        Clear_Button.Visible:=FALSE;
        Cancel_Button.Visible:=FALSE;
        AssemblyPartNumber_Combo.Visible:=FALSE;
        AssemblyPartNumber_Edit.Visible:=TRUE;
      end;
      1:
      begin
        ASNList_DataSet.Parameters.Clear;
        ASNList_DataSet.Parameters.AddParameter.Name := '@List';
        ASNList_DataSet.Parameters.ParamValues['@List']:='C';
        ASNList_DataSet.Parameters.AddParameter.Name := '@Range';
        ASNList_DataSet.Parameters.ParamValues['@Range']:=DataView.ItemIndex;
        ASNList_DataSet.Open;
        ASNItem_Box.Visible:=TRUE;
        Insert_Button.Visible:=FALSE;
        Update_Button.Visible:=TRUE;
        Delete_Button.Visible:=TRUE;
        DeleteASN_Button.Visible:=TRUE;
        RecreateFile_Button.Visible:=FALSE;
        UnsendASN_Button.Visible:=FALSE;
        Clear_Button.Visible:=FALSE;
        Cancel_Button.Visible:=FALSE;
        AssemblyPartNumber_Combo.Visible:=FALSE;
        AssemblyPartNumber_Edit.Visible:=TRUE;
      end;
      2:
      begin
        ASNList_DataSet.Parameters.Clear;
        ASNList_DataSet.Parameters.AddParameter.Name := '@List';
        ASNList_DataSet.Parameters.ParamValues['@List']:='S';
        ASNList_DataSet.Parameters.AddParameter.Name := '@Range';
        ASNList_DataSet.Parameters.ParamValues['@Range']:=DataView.ItemIndex;
        ASNList_DataSet.Open;
        ASNItem_Box.Visible:=FALSE;
        Insert_Button.Visible:=FALSE;
        Update_Button.Visible:=FALSE;
        Delete_Button.Visible:=FALSE;
        DeleteASN_Button.Visible:=FALSE;
        RecreateFile_Button.Visible:=FALSE;
        UnsendASN_Button.Visible:=TRUE;
        UnsendASN_Button.Caption:='&Unsend ASN';
        Clear_Button.Visible:=FALSE;
        Cancel_Button.Visible:=FALSE;
        AssemblyPartNumber_Combo.Visible:=FALSE;
        AssemblyPartNumber_Edit.Visible:=TRUE;
      end;
      3:
      begin
        ASNList_DataSet.Parameters.Clear;
        ASNList_DataSet.Parameters.AddParameter.Name := '@List';
        ASNList_DataSet.Parameters.ParamValues['@List']:='A';
        ASNList_DataSet.Parameters.AddParameter.Name := '@Range';
        ASNList_DataSet.Parameters.ParamValues['@Range']:=DataView.ItemIndex;
        ASNList_DataSet.Open;
        ASNItem_Box.Visible:=FALSE;
        Insert_Button.Visible:=FALSE;
        Update_Button.Visible:=FALSE;
        Delete_Button.Visible:=FALSE;
        DeleteASN_Button.Visible:=FALSE;
        RecreateFile_Button.Visible:=TRUE;
        UnsendASN_Button.Visible:=TRUE;
        UnsendASN_Button.Caption:='&Unsend ASN';
        Clear_Button.Visible:=FALSE;
        Cancel_Button.Visible:=FALSE;
        AssemblyPartNumber_Combo.Visible:=FALSE;
        AssemblyPartNumber_Edit.Visible:=TRUE;
      end;
      4:
      begin
        ASNList_DataSet.Parameters.Clear;
        ASNList_DataSet.Parameters.AddParameter.Name := '@List';
        ASNList_DataSet.Parameters.ParamValues['@List']:='R';
        ASNList_DataSet.Parameters.AddParameter.Name := '@Range';
        ASNList_DataSet.Parameters.ParamValues['@Range']:=DataView.ItemIndex;
        ASNList_DataSet.Open;
        ASNItem_Box.Visible:=TRUE;
        Insert_Button.Visible:=FALSE;
        Update_Button.Visible:=TRUE;
        Delete_Button.Visible:=TRUE;
        DeleteASN_Button.Visible:=TRUE;
        RecreateFile_Button.Visible:=FALSE;
        UnsendASN_Button.Visible:=FALSE;
        Clear_Button.Visible:=FALSE;
        Cancel_Button.Visible:=FALSE;
        AssemblyPartNumber_Combo.Visible:=FALSE;
        AssemblyPartNumber_Edit.Visible:=TRUE;
      end;
    end;
    ASNList_DataSet.Fields[0].Visible:=FALSE;
    ListDBGrid.DataSource:=ASNListDataSource;
    ListDBGrid.Columns[3].Width:=100;
    ListDBGrid.Columns[5].Width:=150;
    ListDBGrid.Columns[7].Width:=150;
    List_GroupBox.Caption:='ASN List';
    Items_GroupBox.Caption:= 'ASN Items';
    fFormMode:=fmASN;
  except
    on e:exception do
    begin
      ShowMessage('Exception: unable to get ASN list, '+e.Message);
      Close;
    end;
  end;
end;

procedure TASNInvoice_Form.Execute;
begin
  try
    ASNorInvoice_ComboBox.ItemIndex:=0;
    ASNStatus_ComboBox.ItemIndex:=0;
    DataView.ItemIndex:=0;
    ShowModal;
  except
    on e:exception do
    begin
      ShowMessage('Exception: unable to get ASN/Invoice info, '+e.Message);
      close;
    end;
  end;
end;

procedure TASNInvoice_Form.FormShow(Sender: TObject);
begin
  case ASNorInvoice_ComboBox.ItemIndex of
    0:
      GetASNs;
    1:
      GetINVOICEs;
  end;
end;

procedure TASNInvoice_Form.ASNListDataSourceDataChange(Sender: TObject;
  Field: TField);
begin
  try
    ASNItems_DataSet.Close;
    ASNItems_DataSet.Parameters.Clear;
    ASNItems_DataSet.Parameters.AddParameter.Name := '@ASNID';
    ASNItems_DataSet.Parameters.ParamValues['@ASNID']:=ASNList_DataSet.FieldByName('IN_ASN_ID').AsInteger;
    ASNItems_DataSet.Open;
    ASNItems_DataSet.Fields[0].Visible:=FALSE;
    ItemsDBGrid.DataSource:=ASNItemsDataSource;

    if ASNStatus_ComboBox.ItemIndex  = 1 then
    begin
        ASNItem_Box.Visible:=TRUE;
        Insert_Button.Visible:=FALSE;
        Update_Button.Visible:=TRUE;
        Delete_Button.Visible:=TRUE;
        DeleteASN_Button.Visible:=TRUE;
        UnsendASN_Button.Visible:=FALSE;
        Clear_Button.Visible:=FALSE;
        Cancel_Button.Visible:=FALSE;
        AssemblyPartNumber_Combo.Visible:=FALSE;
        AssemblyPartNumber_Edit.Visible:=TRUE;
    end;

    if ASNList_DataSet.RecordCount > 0 then
    begin
      if (ASNList_DataSet.FieldByName('Start Seq').AsInteger = -1) and (ASNStatus_ComboBox.ItemIndex=1) then
        Clear_Button.Visible:=FALSE
      else if (ASNList_DataSet.FieldByName('Start Seq').AsInteger <>  -1) and (ASNStatus_ComboBox.ItemIndex=1) then
        Clear_Button.Visible:=TRUE;
    end;
  except
    on e:exception do
    begin
      ShowMessage('Exception: unable to get ASN detail, '+e.Message);
      Close;
    end;
  end;
end;


procedure TASNInvoice_Form.ASNStatus_ComboBoxChange(Sender: TObject);
begin
  case ASNorInvoice_ComboBox.ItemIndex of
    0:
      GetASNs;
    1:
      GetINVOICEs;
  end;
end;

procedure TASNInvoice_Form.ASNorInvoice_ComboBoxChange(Sender: TObject);
begin
  case ASNorInvoice_ComboBox.ItemIndex of
    0:
      begin
        ASNStatus_ComboBox.Items.Clear;
        ASNStatus_ComboBox.Items.Add('ALL');
        ASNStatus_ComboBox.Items.Add('NOT CREATED');
        ASNStatus_ComboBox.Items.Add('SENT');
        ASNStatus_ComboBox.Items.Add('ACCEPTED');
        ASNStatus_ComboBox.Items.Add('REJECTED');
        ASNStatus_ComboBox.ItemIndex:=0;
        GetASNs;
      end;
    1:
      begin
        ASNStatus_ComboBox.Items.Clear;
        ASNStatus_ComboBox.Items.Add('ALL');
        ASNStatus_ComboBox.Items.Add('NOT CREATED');
        ASNStatus_ComboBox.Items.Add('SENT');
        ASNStatus_ComboBox.Items.Add('ACCEPTED');
        ASNStatus_ComboBox.Items.Add('REJECTED');
        ASNStatus_ComboBox.ItemIndex:=0;
        GetINVOICEs;
      end;
  end;
end;

procedure TASNInvoice_Form.Invoice_DataSourceDataChange(Sender: TObject;
  Field: TField);
begin
  try
    INVOICEItems_DataSet.Close;
    INVOICEItems_DataSet.Parameters.Clear;
    INVOICEItems_DataSet.Parameters.AddParameter.Name := '@INVOICEID';
    INVOICEItems_DataSet.Parameters.ParamValues['@INVOICEID']:=InvoiceList_DataSet.FieldByName('IN_INV_ID').AsInteger;
    INVOICEItems_DataSet.Open;
    ItemsDBGrid.DataSource:=INVOICEItemsDataSource;
  except
    on e:exception do
    begin
      ShowMessage('Exception: unable to get INVOICE detail, '+e.Message);
      Close;
    end;
  end;
end;

procedure TASNInvoice_Form.ASNItemsDataSourceDataChange(Sender: TObject;
  Field: TField);
begin
    if not (Insert_Button.Visible) then
    begin
      case ASNStatus_ComboBox.ItemIndex of
        1:
        begin
          ManifestNumber_Edit.Text:=ASNItems_DataSet.fieldbyname('Manifest Number').asString;
          AssemblyPartNumber_Edit.Text:=ASNItems_DataSet.fieldbyname('Assy Part Number').AsString;
          Qty_Edit.Text:=ASNItems_DataSet.fieldbyname('Qty').AsString;
        end;
        4:
        begin
          ManifestNumber_Edit.Text:=ASNItems_DataSet.fieldbyname('Manifest Number').asString;
          AssemblyPartNumber_Edit.Text:=ASNItems_DataSet.fieldbyname('Assy Part Number').AsString;
          Qty_Edit.Text:=ASNItems_DataSet.fieldbyname('Qty').AsString;
        end;
      end;
    end;
end;

procedure TASNInvoice_Form.Clear_ButtonClick(Sender: TObject);
  var
    found:boolean;
begin
  ManifestNumber_Edit.Text:='';
  AssemblyPartNumber_Edit.Visible:=FALSE;
  AssemblyPartNumber_Combo.Visible:=TRUE;
  Cancel_Button.Visible:=TRUE;


  // Load Combo
  AssemblyPartNumber_Combo.Items.Clear;
  AssyManifest_DataSet.Close;
  AssyManifest_DataSet.Open;

  while not  AssyManifest_DataSet.Eof do
  begin
    ASNItems_DataSet.First;
    found:=FALSE;
    while not ASNItems_DataSet.Eof do
    begin
      if (ASNItems_DataSet.FieldByName('Assy Part Number').AsString = AssyManifest_DataSet.FieldByName('Assy').AsString) then
      begin
        found:=TRUE;
        break;
      end;
      ASNItems_DataSet.Next;
    end;
    if not found then
      AssemblyPartNumber_Combo.Items.Add(AssyManifest_DataSet.FieldByName('Assy').AsString);
    AssyManifest_DataSet.Next;
  end;

  ASNItems_DataSet.First;

  AssemblyPartNumber_Combo.ItemIndex:=0;
  AssemblyPartNumber_ComboChange(self);

  Update_Button.Visible:=FALSE;
  Delete_Button.Visible:=FALSE;
  DeleteASN_Button.Visible:=FALSE;
  Clear_Button.Visible:=FALSE;
  Insert_Button.Visible:=TRUE;

  AssemblyPartNumber_Combo.SetFocus;
  Qty_Edit.Text:='0';
end;

procedure TASNInvoice_Form.AssemblyPartNumber_ComboChange(Sender: TObject);
begin
  //generate manifest
  AssyManifest_DataSet.Locate('Assy', AssemblyPartNumber_Combo.Text, [loPartialKey]);
  ManifestNumber_Edit.Text:='7'+
                            copy(ASNList_DataSet.FieldByName('Production Date').AsString,4,1)+
                            copy(ASNList_DataSet.FieldByName('Production Date').AsString,6,2)+
                            copy(ASNList_DataSet.FieldByName('Production Date').AsString,9,2)+
                            AssyManifest_DataSet.FieldByName('Manifest ID').AsString;
  Qty_Edit.SetFocus;

end;

procedure TASNInvoice_Form.Insert_ButtonClick(Sender: TObject);
var
  i:integer;
begin
  try
    if not (tryStrToInt(Qty_Edit.Text, i)) then
    begin
      ShowMessage('Must have numberic value for Qty');
      Qty_Edit.SetFocus;
      exit;
    end;
    ASNItem_Command.CommandText:='dbo.INSERT_ASNDetail';
    ASNItem_Command.Parameters.Clear;
    ASNItem_Command.Parameters.AddParameter.Name:='@ASNid';
    ASNItem_Command.Parameters.ParamValues['@ASNid']:=ASNList_DataSet.FieldByName('IN_ASN_ID').AsInteger;
    ASNItem_Command.Parameters.AddParameter.Name:='@EIN';
    ASNItem_Command.Parameters.ParamValues['@EIN']:=ASNList_DataSet.FieldByName('EIN Number').AsInteger;
    ASNItem_Command.Parameters.AddParameter.Name:='@ManifestNumber';
    ASNItem_Command.Parameters.ParamValues['@ManifestNumber']:=ManifestNumber_Edit.Text;
    ASNItem_Command.Parameters.AddParameter.Name:='@PartNumber';
    ASNItem_Command.Parameters.ParamValues['@PartNumber']:=AssemblyPartNumber_Combo.Text;
    ASNItem_Command.Parameters.AddParameter.Name:='@Qty';
    ASNItem_Command.Parameters.ParamValues['@Qty']:=Qty_Edit.Text;
    ASNItem_Command.Execute;

    Data_Module.LogActLog('ASNINV','Insert ASNID('+IntToStr(ASNList_DataSet.FieldByName('IN_ASN_ID').AsInteger)+') part('+AssemblyPartNumber_Combo.Text+') manifest('+ManifestNumber_Edit.Text+') Qty('+Qty_Edit.Text+')');

    ASNItems_DataSet.Close;
    ASNItems_DataSet.Open;
    ASNItems_DataSet.Fields[0].Visible:=FALSE;

    AssemblyPartNumber_Edit.Visible:=TRUE;
    AssemblyPartNumber_Combo.Visible:=FALSE;

    Update_Button.Visible:=TRUE;
    Delete_Button.Visible:=TRUE;
    DeleteASN_Button.Visible:=TRUE;
    Clear_Button.Visible:=TRUE;
    Insert_Button.Visible:=FALSE;
  except
    on e:exception do
    begin
      ShowMessage('Unable to insert new ASN item, '+e.Message);
      AssemblyPartNumber_Edit.Visible:=TRUE;
      AssemblyPartNumber_Combo.Visible:=FALSE;

      Update_Button.Visible:=TRUE;
      Delete_Button.Visible:=TRUE;
      DeleteASN_Button.Visible:=TRUE;
      Clear_Button.Visible:=TRUE;
      Insert_Button.Visible:=FALSE;
      ManifestNumber_Edit.Text:=ASNItems_DataSet.fieldbyname('Manifest Number').asString;
      AssemblyPartNumber_Edit.Text:=ASNItems_DataSet.fieldbyname('Assy Part Number').AsString;
      Qty_Edit.Text:=ASNItems_DataSet.fieldbyname('Qty').AsString;
    end;
  end;
end;

procedure TASNInvoice_Form.Delete_ButtonClick(Sender: TObject);
begin
  try
    if MessageDlg('Delete ASN item with Manifest Number('+ManifestNumber_Edit.Text+')',
      mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      ASNItem_Command.CommandText:='dbo.DELETE_ASNItem';
      ASNItem_Command.Parameters.Clear;
      ASNItem_Command.Parameters.AddParameter.Name:='@ManifestNumber';
      ASNItem_Command.Parameters.ParamValues['@ManifestNumber']:=ManifestNumber_Edit.Text;
      ASNItem_Command.Execute;

      Data_Module.LogActLog('ASNINV','Delete ASNID('+IntToStr(ASNList_DataSet.FieldByName('IN_ASN_ID').AsInteger)+') manifest('+ManifestNumber_Edit.Text+') Qty('+Qty_Edit.Text+')');

      ASNItems_DataSet.Close;
      ASNItems_DataSet.Parameters.Clear;
      ASNItems_DataSet.Parameters.AddParameter.Name := '@ASNID';
      ASNItems_DataSet.Parameters.ParamValues['@ASNID']:=ASNList_DataSet.FieldByName('IN_ASN_ID').AsInteger;
      ASNItems_DataSet.Open;
      ASNItems_DataSet.Fields[0].Visible:=FALSE;
    end;

  except
    on e:exception do
    begin
      ShowMessage('Exception: unable to DELETE ASN Item('+ManifestNumber_Edit.Text+'), '+e.Message);
      Data_Module.LogActLog('ERROR','Exception: unable to DELETE ASN Item('+ManifestNumber_Edit.Text+'), '+e.Message);
    end;
  end;
end;

procedure TASNInvoice_Form.Update_ButtonClick(Sender: TObject);
var
  bm:TBookMark;
  i:integer;
begin
  try
    if not TryStrToInt(Qty_Edit.Text,i) then
    begin
      ShowMessage('Quantity must be a numeric');
      Qty_Edit.Text:=ASNItems_DataSet.fieldbyname('Qty').AsString;
      Qty_Edit.SetFocus;
      exit;
    end;

    if i<0 then
    begin
      ShowMessage('Quantity must be greater than 0');
      Qty_Edit.Text:=ASNItems_DataSet.fieldbyname('Qty').AsString;
      Qty_Edit.SetFocus;
      exit;
    end;

    ASNItem_Command.CommandText:='dbo.UPDATE_ASNItem';
    ASNItem_Command.Parameters.Clear;
    ASNItem_Command.Parameters.AddParameter.Name:='@IN_ASN_DETAIL_ID';
    ASNItem_Command.Parameters.ParamValues['@IN_ASN_DETAIL_ID']:=ASNItems_DataSet.FieldByNAme('IN_ASN_DETAIL_ID').AsInteger;
    ASNItem_Command.Parameters.AddParameter.Name:='@Qty';
    ASNItem_Command.Parameters.ParamValues['@Qty']:=Qty_Edit.Text;
    ASNItem_Command.Execute;

    Data_Module.LogActLog('ASNINV','Update ASNID('+IntToStr(ASNList_DataSet.FieldByName('IN_ASN_ID').AsInteger)+') manifest('+ManifestNumber_Edit.Text+') Qty('+Qty_Edit.Text+')');

    bm:=ASNItems_DataSet.GetBookmark;
    ASNItems_DataSet.DisableControls;
    ASNItems_DataSet.Close;
    ASNItems_DataSet.Parameters.Clear;
    ASNItems_DataSet.Parameters.AddParameter.Name := '@ASNID';
    ASNItems_DataSet.Parameters.ParamValues['@ASNID']:=ASNList_DataSet.FieldByName('IN_ASN_ID').AsInteger;
    ASNItems_DataSet.Open;
    ASNItems_DataSet.Fields[0].Visible:=FALSE;

    ASNItems_DataSet.EnableControls;
    ASNItems_DataSet.GotoBookmark(bm);

  except
    on e:exception do
    begin
      ShowMessage('Exception: unable to UPDATE ASN Item('+ManifestNumber_Edit.Text+'), '+e.Message);
      Data_Module.LogActLog('ERROR','Exception: unable to UPDATE ASN Item('+ManifestNumber_Edit.Text+'), '+e.Message);
    end;
  end;
end;

procedure TASNInvoice_Form.DeleteASN_ButtonClick(Sender: TObject);
begin
  try
    if MessageDlg('Delete complete ASN ('+ASNList_DataSet.fieldByName('Production Date').AsString+')',
      mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      ASNItem_Command.CommandText:='dbo.DELETE_ASNList';
      ASNItem_Command.Parameters.Clear;
      ASNItem_Command.Parameters.AddParameter.Name:='@ASNid';
      ASNItem_Command.Parameters.ParamValues['@ASNid']:=ASNList_DataSet.FieldByName('IN_ASN_ID').AsInteger;
      ASNItem_Command.Execute;

      GetASNs;

      Data_Module.LogActLog('ASNINV','Delete ASNID('+IntToStr(ASNList_DataSet.FieldByName('IN_ASN_ID').AsInteger)+') Production Date('+ASNList_DataSet.fieldByName('Production Date').AsString+')');
    end;
  except
    on e:exception do
    begin
      ShowMessage('Exception: unable to DELETE ASN ('+ASNList_DataSet.fieldByName('Production Date').AsString+'), '+e.Message);
      Data_Module.LogActLog('ERROR','Exception: unable to DELETE ASN ('+ASNList_DataSet.fieldByName('Production Date').AsString+'), '+e.Message);
    end;
  end;
end;

procedure TASNInvoice_Form.UnsendASN_ButtonClick(Sender: TObject);
begin
  if fFormMode = fmASN then
  begin
    try
      if ASNList_DataSet.RecordCount > 0 then
      begin
        if MessageDlg('Unsend ASN ('+ASNList_DataSet.fieldByName('Production Date').AsString+')',
          mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
          ASNItem_Command.CommandText:='dbo.UPDATE_ASNUnsend';
          ASNItem_Command.Parameters.Clear;
          ASNItem_Command.Parameters.AddParameter.Name:='@ASNid';
          ASNItem_Command.Parameters.ParamValues['@ASNid']:=ASNList_DataSet.FieldByName('IN_ASN_ID').AsInteger;
          ASNItem_Command.Execute;

          GetASNs;

          Data_Module.LogActLog('ASNINV','Unsend ASN('+IntToStr(ASNList_DataSet.FieldByName('IN_ASN_ID').AsInteger)+') Production Date('+ASNList_DataSet.fieldByName('Production Date').AsString+')');
        end;
      end;
    except
      on e:exception do
      begin
        ShowMessage('Exception: unable to unsend ASN('+IntToStr(ASNList_DataSet.FieldByName('IN_ASN_ID').AsInteger)+'), '+e.Message);
        Data_Module.LogActLog('ERROR','Exception: unable to unsend ASN('+IntToStr(ASNList_DataSet.FieldByName('IN_ASN_ID').AsInteger)+'), '+e.Message);
      end;
    end;
  end
  else
  begin
    try
      if InvoiceList_DataSet.RecordCount > 0 then
      begin
        if MessageDlg('Unsend INV ('+IntToStr(InvoiceList_DataSet.FieldByName('EIN Number').AsInteger)+')',
          mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
          ASNItem_Command.CommandText:='dbo.UPDATE_INVUnsend';
          ASNItem_Command.Parameters.Clear;
          ASNItem_Command.Parameters.AddParameter.Name:='@INVid';
          ASNItem_Command.Parameters.ParamValues['@INVid']:=InvoiceList_DataSet.FieldByName('IN_INV_ID').AsInteger;
          ASNItem_Command.Execute;

          GetInvoices;

          Data_Module.LogActLog('ASNINV','Unsend INVOICE('+IntToStr(InvoiceList_DataSet.FieldByName('IN_INV_ID').AsInteger)+')');
        end;
      end;
    except
      on e:exception do
      begin
        ShowMessage('Exception: unable to unsend Invoice('+IntToStr(InvoiceList_DataSet.FieldByName('IN_INV_ID').AsInteger)+'), '+e.Message);
        Data_Module.LogActLog('ERROR','Exception: unable to unsend Invoice('+IntToStr(InvoiceList_DataSet.FieldByName('IN_INV_ID').AsInteger)+'), '+e.Message);
      end;
    end;
  end;
end;



procedure TASNInvoice_Form.Cancel_ButtonClick(Sender: TObject);
begin
  case ASNorInvoice_ComboBox.ItemIndex of
    0:
      GetASNs;
    1:
      GetINVOICEs;
  end;
end;

procedure TASNInvoice_Form.RecreateFile_ButtonClick(Sender: TObject);
var
  EDI856:T856EDI;
  EDI810:T810EDI;
  fcf:TextFile;
  i:integer;
  line:string;
begin
  try
  //Create File Based on EIN
    case ASNorInvoice_ComboBox.ItemIndex of
      0:
      begin
        Data_Module.EDI856DataSet.Close;
        Data_Module.EDI856DataSet.CommandText:='REPORT_EDI856';
        Data_Module.EDI856DataSet.Parameters.Clear;
        Data_Module.EDI856DataSet.Parameters.AddParameter.Name := '@EIN';
        Data_Module.EDI856DataSet.Parameters.ParamValues['@EIN'] := ASNList_DataSet.FieldByName('EIN Number').AsString;
        Data_Module.EDI856DataSet.Open;

        if Data_Module.EDI856DataSet.RecordCount > 0 then
        begin
          EDI856:=T856EDI.Create;
          if EDI856.Execute then
          begin
            if Data_Module.EDI856DataSet.FieldByName('StartSeq').AsString <> '-1' then
            begin
              Data_Module.LogActLog('EDI','Recreate EDI 856 filename ('+Data_Module.fiEDIOut.AsString+'\856'+EDI856.PickupDate+'.txt)');
              AssignFile(fcf, Data_Module.fiEDIOut.AsString+'\856'+copy(EDI856.PickupDate,5,4)+'.txt');
            end
            else
            begin
              Data_Module.LogActLog('EDI',Data_Module.fiEDIOut.AsString+'\8HC'+copy(EDI856.PickupDate,4,5)+'1.txt');
              AssignFile(fcf, Data_Module.fiEDIOut.AsString+'\8HC'+copy(EDI856.PickupDate,4,5)+'1.txt');
            end;

            Rewrite(fcf);

            // Loop through report and save
            for i:=0 to EDI856.EDIRecord.Count-1 do
            begin
              line:=EDI856.EDIRecord[i];
              Writeln(fcf,line);
            end;
            CloseFile(fcf);
            Data_Module.LogActLog('EDI','Create EDI 856 for ('+EDI856.PickupDate+')');
          end
          else
          begin
            ShowMessage('Unable to Recreate EDI856 for ('+EDI856.PickupDate+')');
            Data_Module.LogActLog('EDI','Unable to Recreate EDI856 for ('+EDI856.PickupDate+')');
          end;

          EDI856.Free;

          ASNList_DataSet.Close;
          ASNList_DataSet.Open;

          ShowMessage('Recreate 856 file, complete');
        end
        else
        begin
          ShowMessage('Unable to recreate 856 file');
        end;
      end;
      1:
      begin
        Data_Module.EDI810DataSet.Close;
        Data_Module.EDI810DataSet.CommandText:='REPORT_EDI810';
        Data_Module.EDI810DataSet.Parameters.Clear;
        Data_Module.EDI810DataSet.Parameters.AddParameter.Name := '@EIN';
        Data_Module.EDI810DataSet.Parameters.ParamValues['@EIN'] := InvoiceList_DataSet.FieldByName('EIN Number').AsString;
        Data_Module.EDI810DataSet.Open;

        if Data_Module.EDI810DataSet.RecordCount > 0 then
        begin
          EDI810:=T810EDI.Create;
          EDI810.EIN:=InvoiceList_DataSet.FieldByName('EIN Number').AsInteger;
          if EDI810.Execute then
          begin
            AssignFile(fcf, Data_Module.fiEDIOut.AsString+'\810'+copy(EDI810.PickUpDate,5,4)+'.txt');
            Rewrite(fcf);

            // Loop through report and save
            for i:=0 to EDI810.EDIRecord.Count-1 do
            begin
              line:=EDI810.EDIRecord[i];
              Writeln(fcf,line);
            end;
            CloseFile(fcf);
          end;

          InvoiceList_DataSet.Close;
          InvoiceList_DataSet.Open;

          EDI810.Free;

          ShowMessage('Recreate 810 file, complete');
        end
        else
        begin
          ShowMessage('Unable to recreate 810 file');
        end;
      end;
    end;
  except
    on e:exception do
    begin
      ShowMessage('Exception: '+e.Message);
      Data_Module.LogActLog('ERROR','Exception: '+e.Message);
    end;
  end;

end;

procedure TASNInvoice_Form.SpeedButton1Click(Sender: TObject);
var
  pdate:string;
begin
  //Search
  //  Production date, entry starts with a 2 and is 8 characters
  //  Manifest, entry starts with a 7 and is 8 characters

  if ((pos('2',SearchEdit.Text) = 1) or (pos('7',SearchEdit.Text) = 1)) then
  begin
    case ASNorInvoice_ComboBox.ItemIndex of
      0:
        begin
          if pos('2',SearchEdit.Text)=1 then
          begin
            //search for production date
            if not ASNList_DataSet.Locate('Production Date',SearchEdit.Text,[loPartialKey]) then
            begin
              ShowMessage('Production Date not found');
            end;
          end
          else
          begin
            //search for manifest
            //
            //  First break manifest down to get production data and search for that
            //
            pdate:=copy(formatdatetime('yyyy',now),1,3)+copy(SearchEdit.Text,2,1)+'/'+copy(SearchEdit.Text,3,2)+'/'+copy(SearchEdit.Text,5,2);
            if not ASNList_DataSet.Locate('Production Date',pdate,[loPartialKey]) then
            begin
              ShowMessage('Manifest date not found');
            end
            else
            begin
              //Then the manifest
              if not ASNItems_DataSet.Locate('Manifest Number',SearchEdit.Text,[loPartialKey]) then
              begin
                ShowMessage('Manifest number not found');
              end;
            end;
          end;
        end;
      1:
        begin
          if pos('2',SearchEdit.Text)=1 then
          begin
            if not InvoiceList_DataSet.Locate('Production Date',SearchEdit.Text,[loPartialKey]) then
            begin
              ShowMessage('Production Date not found');
            end;
          end
          else
          begin
            //search for manifest
            //
            //  First break manifest down to get production data and search for that
            //
            pdate:=copy(formatdatetime('yyyy',now),1,3)+copy(SearchEdit.Text,2,1)+'/'+copy(SearchEdit.Text,3,2)+'/'+copy(SearchEdit.Text,5,2);
            if not InvoiceList_DataSet.Locate('Production Date',pdate,[loPartialKey]) then
            begin
              ShowMessage('Manifest date not found');
            end
            else
            begin
              //Then the manifest
              if not INVOICEItems_DataSet.Locate('Manifest',SearchEdit.Text,[loPartialKey]) then
              begin
                ShowMessage('Manifest number not found');
              end;
            end;
          end;
        end;
    end;
  end
  else
  begin
    ShowMessage('Incorrect search data, can search for Production Date or Manifest Number');
    SearchEdit.SetFocus;
  end;

end;

end.
