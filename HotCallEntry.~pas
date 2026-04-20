unit HotCallEntry;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, ADOdb, ComCtrls;

type
  THotCallEntryForm = class(TForm)
    Button_Panel: TPanel;
    Close_Button: TButton;
    Add_Button: TButton;
    Clear_Button: TButton;
    ASNItems_GroupBox: TGroupBox;
    SizeCode_Label: TLabel;
    Label2: TLabel;
    AssyPartsCode1_ComboBox: TComboBox;
    Qty1_Edit: TEdit;
    AssyPartsCode2_ComboBox: TComboBox;
    Qty2_Edit: TEdit;
    AssyPartsCode3_ComboBox: TComboBox;
    Qty3_Edit: TEdit;
    AssyPartsCode4_ComboBox: TComboBox;
    Qty4_Edit: TEdit;
    AssyPartsCode5_ComboBox: TComboBox;
    Qty5_Edit: TEdit;
    AssyPartsCode6_ComboBox: TComboBox;
    Qty6_Edit: TEdit;
    Label3: TLabel;
    Label4: TLabel;
    AssyPartsCode7_ComboBox: TComboBox;
    Qty7_Edit: TEdit;
    AssyPartsCode8_ComboBox: TComboBox;
    Qty8_Edit: TEdit;
    AssyPartsCode9_ComboBox: TComboBox;
    Qty9_Edit: TEdit;
    AssyPartsCode10_ComboBox: TComboBox;
    Qty10_Edit: TEdit;
    AssyPartsCode11_ComboBox: TComboBox;
    Qty11_Edit: TEdit;
    AssyPartsCode12_ComboBox: TComboBox;
    Qty12_Edit: TEdit;
    GroupBox2: TGroupBox;
    CarTruck_Label: TLabel;
    Line_ComboBox: TComboBox;
    Date_Label: TLabel;
    ASN_DateTimePicker: TDateTimePicker;
    Label1: TLabel;
    ManifestNumber_Edit: TEdit;
    procedure Clear_ButtonClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Add_ButtonClick(Sender: TObject);
    procedure Close_ButtonClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure Execute;
    procedure ClearEntries;
  end;

var
  HotCallEntryForm: THotCallEntryForm;

implementation

uses DataModule;

{$R *.dfm}

procedure THotCallEntryForm.ClearEntries;
var
  i:integer;
begin
  Line_ComboBox.ItemIndex:=0;
  ASN_DateTimePicker.DateTime:=now;
  ManifestNumber_Edit.Text:='';
  for i:=0 to ASNItems_GroupBox.ControlCount-1 do
  begin
    if ASNItems_GroupBox.Controls[i] is TComboBox then
    begin
      Data_Module.SelectSingleField('INV_FORECAST_DETAIL_INF', 'VC_ASSY_PART_NUMBER_CODE', TComboBox(ASNItems_GroupBox.Controls[i]));
    end
    else if ASNItems_GroupBox.Controls[i] is TEdit then
    begin
      TEdit(ASNItems_GroupBox.Controls[i]).Text:='';
    end;
  end;
end;

procedure THotCallEntryForm.Execute;
begin
  try
    //Get Line
    with Data_Module.ALC_DataSet do
    begin
      Line_ComboBox.Items.Clear;
      Close;
      CommandType := CmdStoredProc;
      CommandText := 'AD_GetLines;1';
      Parameters.Clear;
      Open;
      if Data_Module.Inv_Connection.Errors.Count > 0 then
      begin
        ShowMessage('Failed on ASN select, '+DAta_Module.Inv_Connection.Errors.Item[0].Get_Description);
        DAta_Module.LogActLog('ERROR','Unable to get ASN line information, '+DAta_Module.Inv_Connection.Errors.Item[0].Get_Description);
        exit;
      end;
      while not eof do
      begin
        Line_ComboBox.Items.Add(fieldByName('LineName').AsString);
        next;
      end;
      ClearEntries;
      ShowModal;
    end;
  except
    on e:exception do
    begin
      Data_Module.LogActLog('ERROR','Failed on ASN select, '+e.Message);
      ShowMessage('Failed on ASN select, '+e.Message);
    end;
  end;
end;
procedure THotCallEntryForm.Clear_ButtonClick(Sender: TObject);
begin
  ClearEntries;
end;

procedure THotCallEntryForm.FormShow(Sender: TObject);
begin
  Line_ComboBox.SetFocus;
end;

procedure THotCallEntryForm.Add_ButtonClick(Sender: TObject);
var
  REcordID, qty:integer;
  i:integer;
begin
  //Insert ASN entry
  try
    if (Line_ComboBox.Text = '') or (Line_ComboBox.Text=' ') then
    begin
      ShowMessage('Select a line');
      Line_ComboBox.SetFocus;
      exit;
    end;

    if not trystrtoint(ManifestNumber_Edit.text, qty) then
    begin
      ShowMessage('Manifest number must be numeric');
      ManifestNumber_Edit.SetFocus;
      exit;
    end;

    if length(ManifestNumber_Edit.text) < 8 then
    begin
      ShowMessage('Manifest number must 8 characters');
      ManifestNumber_Edit.SetFocus;
      exit;
    end;

    for i:=0 to ASNItems_GroupBox.ControlCount-1 do
    begin
      if ASNItems_GroupBox.Controls[i] is TComboBox then
      begin

      end
      else if ASNItems_GroupBox.Controls[i] is TEdit then
      begin
      end;
    end;

    //
    //  Check for correct qty's
    //
    for i:=0 to ASNItems_GroupBox.ControlCount-1 do
    begin
      if ASNItems_GroupBox.Controls[i] is TComboBox then
      begin
        if (TComboBox(ASNItems_GroupBox.Controls[i]).Text <> '') and  (TComboBox(ASNItems_GroupBox.Controls[i]).Text <> ' ') then
        begin
          if not TryStrToInt(TEdit(ASNItems_GroupBox.Controls[i+1]).Text, qty) then
          begin
            ShowMessage('Qty must be numeric');
            TEdit(ASNItems_GroupBox.Controls[i+1]).Text:='0';
            TEdit(ASNItems_GroupBox.Controls[i+1]).SetFocus;
            exit;
          end;
          if Qty <= 0 then
          begin
            ShowMessage('Qty must be greater than 0');
            TEdit(ASNItems_GroupBox.Controls[i+1]).Text:='0';
            TEdit(ASNItems_GroupBox.Controls[i+1]).SetFocus;
            exit;
          end
        end
      end
    end;
    //
    //  Check for part numbers
    //
    for i:=0 to ASNItems_GroupBox.ControlCount-1 do
    begin
      if ASNItems_GroupBox.Controls[i] is TEdit then
      begin
        if TEdit(ASNItems_GroupBox.Controls[i]).text <> '' then
        begin
          if (TComboBox(ASNItems_GroupBox.Controls[i-1]).Text = '') or (TComboBox(ASNItems_GroupBox.Controls[i-1]).Text = ' ') then
          begin
            ShowMessage('Part number required');
            TComboBox(ASNItems_GroupBox.Controls[i-1]).SetFocus;
            exit;
          end
        end
      end
    end;

    //Start Trans
    Data_Module.Inv_Connection.BeginTrans;

    Data_Module.SiteDataset.Close;
    Data_Module.SiteDataset.Open;

    With Data_Module.Inv_StoredProc do
    Begin
      Close;
      ProcedureName := 'dbo.INSERT_ASNInfo;1';
      Parameters.Clear;
      Parameters.AddParameter.Name:='@ASNID';
      Parameters.ParamByName('@ASNID').Direction:=pdOutput;
      Parameters.ParamValues['@ASNID']:=0;
      Parameters.AddParameter.Name := '@LineName';
      Parameters.ParamValues['@LineName'] := Line_ComboBox.Text;
      Parameters.AddParameter.Name := '@AssyLine';
      Parameters.ParamValues['@AssyLine'] := '';
      Parameters.AddParameter.Name := '@StartSeq';
      Parameters.ParamValues['@StartSeq'] := -1;
      Parameters.AddParameter.Name := '@DTStartSeq';
      Parameters.ParamValues['@DTStartSeq'] := now;
      Parameters.AddParameter.Name := '@EndSeq';
      Parameters.ParamValues['@EndSeq'] := -1;
      Parameters.AddParameter.Name := '@DTEndSeq';
      Parameters.ParamValues['@DTEndSeq'] := now;
      Parameters.AddParameter.Name := '@QTY';
      Parameters.ParamValues['@QTY'] := qty;
      Parameters.AddParameter.Name := '@PDate';
      Parameters.ParamValues['@PDate'] := formatdatetime('yyyymmdd', ASN_DateTimePicker.DateTime);
      Parameters.AddParameter.Name := '@EIN';
      Parameters.ParamValues['@EIN'] := Data_Module.SiteDataset.FieldByNAme('SiteEIN').AsInteger+1;

      ExecProc;

      RecordID:=Parameters.ParamByName('@ASNID').Value;


      for i:=0 to ASNItems_GroupBox.ControlCount-1 do
      begin
        if ASNItems_GroupBox.Controls[i] is TEdit then
        begin
          if TEdit(ASNItems_GroupBox.Controls[i]).text <> '' then
          begin
            TryStrToInt(TEdit(ASNItems_GroupBox.Controls[i]).Text, qty);
            Close;
            Parameters.Clear;
            ProcedureName:='INSERT_ASNDetail;1';
            Parameters.Clear;
            Parameters.AddParameter.Name := '@ASNID';
            Parameters.ParamValues['@ASNID'] := RecordID;
            Parameters.AddParameter.Name := '@EIN';
            Parameters.ParamValues['@EIN'] := Data_Module.SiteDataset.FieldByNAme('SiteEIN').AsInteger+1;
            Parameters.AddParameter.Name := '@Manifest';
            Parameters.ParamValues['@Manifest'] := ManifestNumber_Edit.Text;
            Parameters.AddParameter.Name := '@PartNumber';
            Parameters.ParamValues['@PartNumber'] := TComboBox(ASNItems_GroupBox.Controls[i-1]).Text;
            Parameters.AddParameter.Name := '@Qty';
            Parameters.ParamValues['@Qty'] := qty;
            Parameters.AddParameter.Name := '@Hotcall';
            Parameters.ParamValues['@Hotcall'] := 1;
            ExecProc;
            Data_Module.LogActLog('HOTCALL','Added PartNumber('+TComboBox(ASNItems_GroupBox.Controls[i-1]).Text+') Qty('+TEdit(ASNItems_GroupBox.Controls[i]).Text+') HotCall manifest('+ManifestNumber_Edit.Text+')');
          end
        end
      end;
    end;


    //Set new EIN
    Data_Module.UpdateReportCommand.CommandType:=cmdStoredProc;
    Data_Module.UpdateReportCommand.CommandText:='AD_UpdateEIN';
    Data_Module.UpdateReportCommand.Execute;

    //End Trans
    Data_Module.Inv_Connection.CommitTrans;

    Data_Module.LogActLog('HOTCALL','Added HotCall manifest('+ManifestNumber_Edit.Text+')');
    ShowMessage('Added HotCall manifest('+ManifestNumber_Edit.Text+')');


    ClearEntries;
    Line_ComboBox.SetFocus;
  except
    on e:exception do
    begin
      if Data_Module.Inv_Connection.InTransaction then
        Data_Module.Inv_Connection.RollbackTrans;
      Data_Module.LogActLog('ERROR','Failed on HotCall ASN add, '+e.Message);
      ShowMessage('Failed on HotCall ASN add, '+e.Message);
    end;
  end;
end;

procedure THotCallEntryForm.Close_ButtonClick(Sender: TObject);
begin
  Close;
end;

end.
