unit ASNSelect;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, NummiTime, Seqnum, StdCtrls, ComCtrls, ExtCtrls, Mask, Adodb,
  NUMMIComboBox;

type
  TASNSelect_Form = class(TForm)
    CarTruckShip_Label: TLabel;
    Check_Button: TButton;
    Authentication_GroupBox: TGroupBox;
    ShipQty_Label: TLabel;
    ShipQty_MaskEdit: TMaskEdit;
    Button_Panel: TPanel;
    CreateASNEntries_Button: TButton;
    Close_Button: TButton;
    Shipping_Panel: TPanel;
    Label4: TLabel;
    StartSeqNo_Label: TLabel;
    CarTruck_Label: TLabel;
    Date_Label: TLabel;
    Label3: TLabel;
    Label5: TLabel;
    StartSeqNo_Edit: TEdit;
    LastSeqNo_Edit: TEdit;
    ASN_DateTimePicker: TDateTimePicker;
    SeqNum: TSequenceNumber;
    NT: TNummiTime;
    Line_ComboBox: TComboBox;
    StartBox: TNUMMIComboBox;
    EndBox: TNUMMIComboBox;
    CreateASN_Button: TButton;
    procedure Line_ComboBoxChange(Sender: TObject);
    procedure ASN_DateTimePickerChange(Sender: TObject);
    procedure StartSeqNo_EditChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Check_ButtonClick(Sender: TObject);
    procedure CreateASN_ButtonClick(Sender: TObject);
    procedure StartBoxChange(Sender: TObject);
    procedure CreateASNEntries_ButtonClick(Sender: TObject);
    procedure LastSeqNo_EditChange(Sender: TObject);
  private
    { Private declarations }
    fInit:boolean;
    fChecked:boolean;
    fCurrentDateTime:TDateTime;
    function LoadASNDates:boolean;
    procedure LoadSeqNumbers;
    procedure ClearCheck;
  public
    { Public declarations }
    procedure Execute;
  end;

var
  ASNSelect_Form: TASNSelect_Form;

implementation

uses DataModule, EDI856Object;

{$R *.dfm}
procedure TASNSelect_Form.Execute;
var
  lines:TStringList;
  i:integer;
begin
  try
    //Get Line
    fInit:=FALSE;

    StartSeqNo_Edit.MaxLength:=Data_Module.fiTruckSeqLength.AsInteger;
    LastSeqNo_Edit.MaxLength:=Data_Module.fiTruckSeqLength.AsInteger;

    lines:=Data_Module.GetLines;
    for i:=0 to lines.Count-1 do
    begin
      Line_ComboBox.Items.Add(lines[i]);
    end;
    Line_ComboBox.ItemIndex:=0;

    LoadASNDates;

    fInit:=TRUE;
    ShowModal;
  except
    on e:exception do
    begin
      Data_Module.LogActLog('ERROR','Failed on ASN select, '+e.Message);
      ShowMessage('Failed on ASN select, '+e.Message);
    end;
  end;
end;

procedure TASNSelect_Form.Line_ComboBoxChange(Sender: TObject);
begin
  LoadASNDates;
end;

function TASNSelect_Form.LoadASNDates;
begin
  result:=FALSE;
  Data_Module.LineName:=Line_ComboBox.Text;
  Data_Module.GetNextASNDate;
  ASN_DateTimePicker.DateTime:=Data_Module.ProductionDateDT;
  fCurrentDateTime:=ASN_DateTimePicker.DateTime;

  StartSeqNo_Edit.Text:='';
  StartBox.Items.Clear;

  if Data_Module.StartSeq <> '-1' then
  begin
    StartSeqNo_Edit.Text:=Format('%0.3d', [StrToInt(Data_Module.StartSeq)]);//Data_Module.StartSeq;
    StartBox.Items.Add(FormatDateTime('yyyy-mm-dd hh:mm:ss.zzz',Data_Module.BeginDate));
    StartBox.ItemIndex:=0;
  end;

  fChecked:=FALSE;
  CreateASN_Button.Enabled:=FALSE;
  CreateASNEntries_Button.Enabled:=FALSE;
  Check_Button.Enabled:=FALSE;
end;

procedure TASNSelect_Form.LoadSeqNumbers;
var
  init:boolean;
begin
  try
    with Data_Module.Inv_DataSet do
    begin
      Close;
      CommandType := CmdStoredProc;
      CommandText := 'SELECT_ASNSeq;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@LineName';
      Parameters.ParamValues['@LineName'] := Line_ComboBox.Text;
      Parameters.AddParameter.Name := '@PDate';
      Parameters.ParamValues['@PDate'] := formatdatetime('yyyymmdd',ASN_DateTimePicker.Date);
      Open;
      if DAta_Module.Inv_Connection.Errors.Count > 0 then
      begin
        ShowMessage('Failed on ASN select, '+DAta_Module.Inv_Connection.Errors.Item[0].Get_Description);
        DAta_Module.LogActLog('ERROR','Unable to get ASN line information, '+DAta_Module.Inv_Connection.Errors.Item[0].Get_Description);
        exit;
      end;

      init:=FInit;
      fInit:=FALSE;
      if recordcount > 0 then
      begin
        StartSeqNo_Edit.Text:=fieldbyname('VC_START_SEQ_NUMBER').AsString;
        StartBox.Text:=fieldbyname('DT_START_SEQ').AsString;
        StartSeqNo_Edit.ReadOnly:=TRUE;
        StartBox.Items.Clear;
        StartBox.Items.Add(fieldbyname('DT_START_SEQ').AsString);
        StartBox.ItemIndex:=0;
        StartBox.ReadOnly:=TRUE;
        LastSeqNo_Edit.Text:=fieldbyname('VC_END_SEQ_NUMBER').AsString;
        EndBox.Text:=fieldbyname('DT_END_SEQ').AsString;
        LastSeqNo_Edit.ReadOnly:=TRUE;
        EndBox.Items.Clear;
        EndBox.Items.Add(fieldbyname('DT_END_SEQ').AsString);
        EndBox.ItemIndex:=0;
        EndBox.ReadOnly:=TRUE;
        ShipQty_MaskEdit.Text:=IntToStr(fieldbyname('IN_QTY').AsInteger);

        Check_Button.Enabled:=FALSE;
        CreateASN_Button.Enabled:=FALSE;
        CreateASNEntries_Button.Enabled:=FALSE;
      end
      else
      begin
        StartSeqNo_Edit.Text:='';
        StartSeqNo_Edit.ReadOnly:=FALSE;
        StartBox.Items.Clear;
        StartBox.Text:='';
        StartBox.ReadOnly:=FALSE;
        LastSeqNo_Edit.Text:='';
        LastSeqNo_Edit.ReadOnly:=FALSE;
        EndBox.Items.Clear;
        EndBox.Text:='';
        EndBox.ReadOnly:=FALSE;
        ShipQty_MaskEdit.Text:='0';

        Check_Button.Enabled:=TRUE;
        CreateASN_Button.Enabled:=FALSE;
        CreateASNEntries_Button.Enabled:=FALSE;
      end;
      fInit:=init;
    end;
  except
    on e:exception do
    begin
      Data_Module.LogActLog('ERROR','Failed on ASN select, '+e.Message);
      ShowMessage('Failed on ASN select, '+e.Message);
    end;
  end;
end;

procedure TASNSelect_Form.ASN_DateTimePickerChange(Sender: TObject);
begin
  if fInit and (fCurrentDateTime<>ASN_DateTimePicker.DateTime) then
  begin
    LoadSeqNumbers;
  end;
  fCurrentDateTime:=ASN_DateTimePicker.DateTime;
end;

procedure TASNSelect_Form.ClearCheck;
begin
  fChecked:=FALSE;
  Check_Button.Enabled:=FALSE;
  ShipQty_MaskEdit.Text:='0';
  CreateASN_Button.Enabled:=FALSE;
  CreateASNEntries_Button.Enabled:=FALSE;
end;

procedure TASNSelect_Form.LastSeqNo_EditChange(Sender: TObject);
begin
  if fInit then
  begin
    if LastSeqNo_Edit.Text <> '' then
    begin
      try
        Data_Module.StartSeq:=LastSeqNo_Edit.Text;
        if Data_Module.GetLastSeqDate(Data_Module.fiRevSeqLookup.AsInteger) then
        begin
          EndBox.Items.Clear;
          while not Data_Module.ALC_StoredProc.Eof do
          begin
            //only add is later than start
            if StartBox.Text < FormatDateTime('yyyy-mm-dd hh:mm:ss.zzz',Data_Module.ALC_StoredProc.fieldbyname('LastTime').AsDateTime) then
            begin
              EndBox.Items.Add(FormatDateTime('yyyy-mm-dd hh:mm:ss.zzz',Data_Module.ALC_StoredProc.fieldbyname('LastTime').AsDateTime));
            end;
            Data_Module.ALC_StoredProc.Next;

            EndBox.ItemIndex:=0;

            Check_Button.Enabled:=TRUE;
            ShipQty_MaskEdit.Text:='0';
            CreateASN_Button.Enabled:=FALSE;
            CreateASNEntries_Button.Enabled:=FALSE;
          end;
        end
        else
          ShowMessage('Sequence not found');
      except
        on e:exception do
        begin
          ShowMessage('Unable to access sequence number date/time, '+e.Message);
        end;
      end;
    end
    else
    begin
      EndBox.Items.Clear;
      EndBox.Text:='';
      ClearCheck;
    end;
  end;
end;

procedure TASNSelect_Form.StartSeqNo_EditChange(Sender: TObject);
begin
  if fInit then
  begin
    if StartSeqNo_Edit.Text <> '' then
    begin
      try
        Data_Module.StartSeq:=StartSeqNo_Edit.Text;
        if Data_Module.GetLastSeqDate(Data_Module.fiRevSeqLookup.AsInteger) then
        begin
            StartBox.Items.Clear;
            while not Data_Module.ALC_StoredProc.Eof do
            begin
              StartBox.Items.Add(FormatDateTime('yyyy-mm-dd hh:mm:ss.zzz',Data_Module.ALC_StoredProc.fieldbyname('LastTime').AsDateTime));
              Data_Module.ALC_StoredProc.Next;
            end;
            StartBox.ItemIndex:=0;

            // Clear End
            EndBox.Items.Clear;
            LastSeqNo_Edit.Text:='';
            ClearCheck;
        end
        else
          ShowMessage('Sequence not found');

      except
        on e:exception do
        begin
          ShowMessage('Unable to access sequence number date/time, '+e.Message);
        end;
      end;
      CreateASN_Button.Enabled:=FALSE;
      CreateASNEntries_Button.Enabled:=FALSE;
    end
    else
    begin
      StartBox.Items.Clear;
      StartBox.Text:='';
      EndBox.Items.Clear;
      EndBox.Text:='';
      LastSeqNo_Edit.Text:='';
      ClearCheck;
    end;
  end;
end;

procedure TASNSelect_Form.FormShow(Sender: TObject);
begin
  if StartSeqNo_Edit.Text='' then
    StartSeqNo_Edit.SetFocus
  else if LastSeqNo_Edit.Text='' then
    LastSeqNo_Edit.SetFocus
  else
    ASN_DateTimePicker.SetFocus;
end;

procedure TASNSelect_Form.Check_ButtonClick(Sender: TObject);
begin
  try
    // Get count of vehicle requested from tire database
    if (StartSeqNo_Edit.Text <> '') and (LastSeqNo_Edit.Text <> '') and (StartBox.Text <> '') and  (EndBox.Text <> '') then
    begin
      with Data_Module do
      begin
        BeginDateStr:=StartBox.Text;
        EndDatestr:=EndBox.Text;
        StartSeq:=format('%.4d',[StrToInt(StartSeqNo_Edit.Text)]);
        LastSeq:=format('%.4d',[StrToInt(LastSeqNo_Edit.Text)]);

        CheckShippingInfo;

        with ALC_StoredProc do
        begin
          ShipQty_MaskEdit.Text:=IntToStr(recordCount);
        end;
      end;
      fchecked:=true;
      Check_Button.Enabled:=FALSE;
      CreateASN_Button.Enabled:=TRUE;
      CreateASNEntries_Button.Enabled:=TRUE;
      Data_Module.LogActLog('ASN','ASN Sequence number check, P:'+formatdatetime('mm/dd/yyyy',ASN_DateTimePicker.DateTime)+' L:'+Data_Module.LineName+' S:'+StartSeqNo_Edit.Text+' E:'+LastSeqNo_Edit.Text+' Q:'+ShipQty_MaskEdit.Text);
    end;
  except
    on e:exception do
    begin
      Data_Module.LogActLog('ERROR','Failed on ASN check, '+e.Message);
      ShowMessage('Failed on ASN check, '+e.Message);
    end;
  end;
end;

procedure TASNSelect_Form.CreateASNEntries_ButtonClick(Sender: TObject);
begin
  try
    Data_Module.Inv_Connection.BeginTrans;

    Data_Module.LogActLog('ASN','Create ASN entries only');
    Data_Module.BeginDatestr:=StartBox.Text;
    Data_Module.EndDateStr:=EndBox.Text;
    Data_Module.ProductionDate:=formatdatetime('yyyymmdd',ASN_DateTimePicker.DateTime);
    Data_Module.SiteDataset.Close;
    Data_Module.SiteDataset.Open;
    Data_Module.EIN:=Data_Module.SiteDataset.FieldByNAme('SiteEIN').AsInteger;
    Data_Module.LineName:=Line_ComboBox.Text;
    Data_Module.Quantity:=StrToInt(trim(ShipQty_MaskEdit.Text));

    If Data_Module.InsertASNInfo Then
    begin
      //ShowMessage('EDI 856 complete');
      Data_Module.UpdateReportCommand.CommandType:=cmdStoredProc;
      Data_Module.UpdateReportCommand.CommandText:='AD_UpdateEIN';
      Data_Module.UpdateReportCommand.Execute;

      Data_Module.Inv_Connection.CommitTrans;

      fInit:=FALSE;
      StartSeqNo_Edit.Text:='';
      StartBox.items.clear;

      LoadASNDates;

      fInit:=TRUE;

      LastSeqNo_Edit.Text:='';
      EndBox.items.clear;
      EndBox.ItemIndex:=0;


      Check_Button.Enabled:=FALSE;
      ShipQty_MaskEdit.Text:='0';
      CreateASN_Button.Enabled:=FALSE;
      CreateASNEntries_Button.Enabled:=FALSE;
      LastSeqNo_Edit.SetFocus;

      Data_Module.LogActLog('ASN','Finish create ASN entries');
    end
    else
    begin
      ShowMessage('Failed to create ASN entries');
      if Data_Module.Inv_Connection.InTransaction then
        Data_Module.Inv_Connection.RollbackTrans;
      Data_Module.LogActLog('ASN','Failed to create ASN entries');
    end;
  except
    on e:exception do
    begin
      Data_Module.LogActLog('ERROR','Failed on ASN entries create, '+e.Message);
      ShowMessage('Failed on ASN entries create, '+e.Message);
      if Data_Module.Inv_Connection.InTransaction then
        Data_Module.Inv_Connection.RollbackTrans;
    end;
  end;
end;

procedure TASNSelect_Form.CreateASN_ButtonClick(Sender: TObject);
var
  i:integer;
  EDI856:T856EDI;
  fcf:TextFile;
  line:string;
begin
  try
    Data_Module.Inv_Connection.BeginTrans;

    Data_Module.LogActLog('ASN','Create ASN with files');
    Data_Module.BeginDatestr:=StartBox.Text;
    Data_Module.EndDateStr:=EndBox.Text;
    Data_Module.ProductionDate:=formatdatetime('yyyymmdd',ASN_DateTimePicker.DateTime);
    Data_Module.SiteDataset.Close;
    Data_Module.SiteDataset.Open;
    Data_Module.EIN:=Data_Module.SiteDataset.FieldByNAme('SiteEIN').AsInteger;
    Data_Module.AssyName:='T';
    Data_Module.Quantity:=StrToInt(trim(ShipQty_MaskEdit.Text));

    If Data_Module.InsertASNInfo Then
    begin
      EDI856:=T856EDI.Create;
      if EDI856.Execute then
      begin
        AssignFile(fcf, Data_Module.fiEDIOut.AsString+'\856'+copy(Data_Module.ProductionDate,4,5)+'.txt');
        Rewrite(fcf);

        // Loop through report and save
        for i:=0 to EDI856.EDIRecord.Count-1 do
        begin
          line:=EDI856.EDIRecord[i];
          Writeln(fcf,line);
        end;
        CloseFile(fcf);

        //ShowMessage('EDI 856 complete');

        Data_Module.UpdateReportCommand.CommandType:=cmdStoredProc;
        Data_Module.UpdateReportCommand.CommandText:='AD_UpdateEIN';
        Data_Module.UpdateReportCommand.Execute;

        Data_Module.ASNStatus:='S';
        Data_Module.UpdateASNStatus;

        Data_Module.Inv_Connection.CommitTrans;


        fInit:=FALSE;
        StartSeqNo_Edit.Text:='';
        StartBox.items.clear;

        LoadASNDates;

        fInit:=TRUE;

        LastSeqNo_Edit.Text:='';
        EndBox.items.clear;
        EndBox.ItemIndex:=0;


        Check_Button.Enabled:=FALSE;
        ShipQty_MaskEdit.Text:='0';
        CreateASN_Button.Enabled:=FALSE;
        CreateASNEntries_Button.Enabled:=FALSE;
        LastSeqNo_Edit.SetFocus;
      end
      else
      begin
        ShowMessage('Unable to create EDI856');
        if Data_Module.Inv_Connection.InTransaction then
          Data_Module.Inv_Connection.RollbackTrans;
        Data_Module.LogActLog('ASN','FAiled ASN with files');
      end;

      EDI856.Free;
    end
    else
    begin
      ShowMessage('Unable to create EDI856');
      if Data_Module.Inv_Connection.InTransaction then
        Data_Module.Inv_Connection.RollbackTrans;
      Data_Module.LogActLog('ASN','FAiled ASN with files');
    end;
  except
    on e:exception do
    begin
      Data_Module.LogActLog('ERROR','Failed on ASN create file, '+e.Message);
      ShowMessage('Failed on ASN create file, '+e.Message);
      if Data_Module.Inv_Connection.InTransaction then
        Data_Module.Inv_Connection.RollbackTrans;
    end;
  end;
end;

procedure TASNSelect_Form.StartBoxChange(Sender: TObject);
begin
  LastSeqNo_Edit.Text:='';
  EndBox.Items.Clear;
  LastSeqNo_Edit.SetFocus;
end;



end.
