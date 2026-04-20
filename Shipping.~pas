//****************************************************************
//
//       Inventory Control
//
//       Copyright (c) 2002 Failproof Manufacturing Systems.
//
//****************************************************************
//
//  10/25/2002  Aaron Huge      Initial creation
//  12/19/2002  David Verespey  Changes for inventory removal
//  01/05/2005  David Verespey  Mods for GALC sequence number changes
//  03/24/2005  David Verespey  More Mods for GALC

unit Shipping;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, StdCtrls, DataModule, ExtCtrls, Mask, Seqnum, DateUtils,
  NummiTime,ADOdb, NUMMIComboBox ;

type
  TShipping_Form = class(TForm)
    CarTruckShip_Label: TLabel;
    Check_Button: TButton;
    Authentication_GroupBox: TGroupBox;
    ShipQty_Label: TLabel;
    ContinuationNo_Label: TLabel;
    Button_Panel: TPanel;
    Insert_Button: TButton;
    Close_Button: TButton;
    Shipping_Panel: TPanel;
    Label4: TLabel;
    StartSeqNo_Label: TLabel;
    CarTruck_Label: TLabel;
    Date_Label: TLabel;
    StartSeqNo_Edit: TEdit;
    LastSeqNo_Edit: TEdit;
    ShipQty_MaskEdit: TMaskEdit;
    ContNo_MaskEdit: TMaskEdit;
    Production_DateTimePicker: TDateTimePicker;
    SeqNum: TSequenceNumber;
    NT: TNummiTime;
    Label3: TLabel;
    Label5: TLabel;
    Line_ComboBox: TComboBox;
    StartBox: TNUMMIComboBox;
    EndBox: TNUMMIComboBox;
    UpdateShipping_Button: TButton;
    procedure Insert_ButtonClick(Sender: TObject);
    procedure TextChange(Sender: TObject);
    procedure MaskEditExit(Sender: TObject);
    procedure Check_ButtonClick(Sender: TObject);
    procedure Production_DateTimePickerChange(Sender: TObject);
    procedure Refill_ButtonClick(Sender: TObject);
    procedure StartSeqNo_EditChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure StartBoxChange(Sender: TObject);
    procedure Line_ComboBoxChange(Sender: TObject);
    procedure UpdateShipping_ButtonClick(Sender: TObject);
  private
    { Private declarations }
    fCheck:boolean;
    fNoDateTimeUpdate:boolean;
    function HoldDetails: String;
    procedure SetDetailBoxes;
    procedure GetProductionDate;
  public
    { Public declarations }
    function Execute: boolean;
  end;

var
  Shipping_Form: TShipping_Form;

implementation

uses ManualShipping, ModifyShipping;

{$R *.dfm}

procedure TShipping_Form.GetProductionDate;
begin
  //
  // Get next production date to process, base on last process and overtime/holiday info
  //
  Data_module.GetNextProductionDate;
  Production_DateTimePicker.Date := Data_Module.ProductionDateDT;

  SetDetailBoxes;
end;

function TShipping_Form.Execute: boolean;
Begin
  Result:= True;
  fNoDateTimeUpdate:=FALSE;
  try
    try
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
        Line_ComboBox.ItemIndex:=0;
      end;

      Data_Module.LineName:=Line_ComboBox.Text;

      GetProductionDate;

      fCheck:=False;
      ShowModal;
    except
      On E:Exception do
        showMessage('Error on get information shipping screen.'
          + #13 + 'ERROR:'
          + #13 + E.Message);
    end;
  finally
    If ModalResult = mrCancel Then
      Result:= False;
  end;
end;      //Execute

procedure TShipping_Form.Insert_ButtonClick(Sender: TObject);
begin
  if fcheck then
  begin
    HoldDetails;
    Data_Module.Refill:=False;
    If Not Data_Module.InsertShippingInfo Then
      ShowMessage('Failed to update shipping info.')
    else
      ShowMessage('Shipping info update is complete.');

    //Data_Module.ClearControls(Shipping_Panel);
    //Data_Module.ClearControls(Authentication_GroupBox);
    SetDetailBoxes;
    StartSeqNo_Edit.SelectAll;
    if Shipping_Form.Visible then
      StartSeqNo_Edit.SetFocus;
  end
  else
  begin
    ShowMessage('Please run check first');
  end;
end;          //Insert_ButtonClick

function TShipping_Form.HoldDetails: String;
var
  fErrMsg: String;
Begin
  fErrMsg := '';

  With Data_Module do
  Begin
    BeginDatestr:=StartBox.Text;
    EndDateStr:=EndBox.Text;
    StartSeq:=format('%.4d',[StrToInt(StartSeqNo_Edit.Text)]);
    LastSeq:=format('%.4d',[StrToInt(LastSeqNo_Edit.Text)]);

    Quantity := StrToInt(Trim(ShipQty_MaskEdit.Text));
    Continuation := StrToInt(Trim(ContNo_MaskEdit.Text));
  End;      //With

  If Not (fErrMsg = '') Then
    fErrMsg := 'The following fields need to be corrected:' + fErrMsg;
  result := fErrMsg;
End;          //HoldDetails

procedure TShipping_Form.SetDetailBoxes;
var
  seq:integer;
Begin
  With Data_Module do
  Begin

    ProductionDate:=formatdatetime('yyyymmdd',Production_DateTimePicker.Date);
    Inv_dataset.Filter:='';
    Inv_dataset.Filtered:=FALSE;

    GetShippingInfo;

    if ProductionDate = Inv_DataSet.FieldByName('VC_PRODUCTION_DATE').AsString then
    begin
      //date already processed, fill in data and show (lock controls)
      StartSeqNo_Edit.Text := Inv_DataSet.FieldByName('VC_START_SEQ_NUMBER').AsString;
      LastSeqNo_Edit.Text := Inv_DataSet.FieldByName('VC_END_SEQ_NUMBER').AsString;
      StartBox.Items.Clear;
      StartBox.Items.Add(FormatDateTime('yyyy-mm-dd hh:mm:ss.zzz',Inv_DataSet.FieldByName('DT_START_SEQ_NUMBER').AsDateTime));
      StartBox.ItemIndex:=0;
      endBox.Items.Clear;
      endBox.Items.Add(FormatDateTime('yyyy-mm-dd hh:mm:ss.zzz',Inv_DataSet.FieldByName('DT_END_SEQ_NUMBER').AsDateTime));
      endBox.ItemIndex:=0;
      ShipQty_MaskEdit.Text := IntToStr(Inv_DataSet.FieldByName('IN_QTY').AsInteger);
      ContNo_MaskEdit.Text := IntToStr(Inv_DataSet.FieldByName('IN_CONTINUE_NUMBER').AsInteger);
      StartSeqNo_Edit.ReadOnly:=True;
      LastSeqNo_Edit.ReadOnly:=True;
      StartBox.ReadOnly:=TRUE;
      endBox.ReadOnly:=TRUE;
      Insert_Button.Enabled:=False;
      Check_Button.Enabled:=False;

      //
      //  Add Edit Button enable, to change shipping values
      //
      UPdateShipping_Button.Visible:=TRUE;


    end
    else
    begin
      // Last date found
      if Inv_DataSet.FieldByName('VC_END_SEQ_NUMBER').AsString <> '' then
      begin
        seq:=StrToInt(Inv_DataSet.FieldByName('VC_END_SEQ_NUMBER').AsString);
        inc(seq);
        if seq > 999 then
          seq:=0;
        StartSeqNo_Edit.Text := format('%.3d',[seq]);

        if Shipping_Form.Visible then
          LastSeqNo_Edit.SetFocus;
      end
      else
      begin
        StartSeqNo_Edit.Text:='';
        StartBox.Items.Clear;

        if Shipping_Form.Visible then
          StartSeqNo_Edit.SetFocus;
      end;

      EndBox.Items.Clear;
      LastSeqNo_Edit.Text := '';
      ShipQty_MaskEdit.Text := '';
      ContNo_MaskEdit.Text := '';
      StartSeqNo_Edit.ReadOnly:=False;
      LastSeqNo_Edit.ReadOnly:=False;
      StartBox.ReadOnly:=FALSE;
      endBox.ReadOnly:=FALSE;
      Insert_Button.Enabled:=True;
      Check_Button.Enabled:=True;
      UPdateShipping_Button.Visible:=FALSE;
    end;
  End;        //With
End;        //SetDetailBoxes

procedure TShipping_Form.TextChange(Sender: TObject);
Begin
  If Sender.ClassName = 'TMaskEdit' Then
    If Length(Trim(TMaskEdit(Sender).Text)) < 1 Then
      TMaskEdit(Sender).Text := '0';
End;          //TextChange

procedure TShipping_Form.MaskEditExit(Sender: TObject);
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

procedure TShipping_Form.Check_ButtonClick(Sender: TObject);
begin
  // Get count of vehicle requested from tire database
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
      Last;
      ContNo_MaskEdit.Text:=FieldByName('VehicleId').AsString;
    end;

  end;
  fcheck:=true;
  Data_Module.LogActLog('SHIPPING','Sequence number check, L:'+Data_Module.AssyCode+' S:'+StartSeqNo_Edit.Text+' E:'+LastSeqNo_Edit.Text+' Q:'+ShipQty_MaskEdit.Text);
end;          //Check_ButtonClick

procedure TShipping_Form.Production_DateTimePickerChange(Sender: TObject);
begin
  // Check if that date has already been run
  if not fNoDateTimeUpdate then
  begin
    fNoDateTimeUpdate:=TRUE;

    //
    // Get next production date to process, base on last process and overtime/holiday info
    //
    Data_module.GetNextProductionDate;
    Data_Module.ProductionDateDT:=Production_DateTimePicker.Date;

    SetDetailBoxes;

    fCheck:=False;
    fNoDateTimeUpdate:=FALSE;
  end;
end;

procedure TShipping_Form.Refill_ButtonClick(Sender: TObject);
begin
  //Refill without reduction in inventory or shipping record
  HoldDetails;
  Data_Module.BeginDateStr:=StartBox.Text;
  Data_Module.EndDatestr:=EndBox.Text;
  Data_Module.Refill:=True;
  If Not Data_Module.CalculateFRS Then
    ShowMessage('Failed to refill shipping info.')
  else
    ShowMessage('ReFill info update is complete.');

  Data_Module.ClearControls(Shipping_Panel);
  Data_Module.ClearControls(Authentication_GroupBox);
  SetDetailBoxes;
  StartSeqNo_Edit.SelectAll;
  StartSeqNo_Edit.SetFocus;
end;

procedure TShipping_Form.StartBoxChange(Sender: TObject);
begin
  StartSeqNo_EditChange(LastSeqNo_Edit);
end;

procedure TShipping_Form.StartSeqNo_EditChange(Sender: TObject);
begin
  if length(TEdit(Sender).Text) = TEdit(Sender).MaxLength then
  begin
    // Get Date Time of seq
    try
      Data_Module.StartSeq:=format('%.4d',[StrToInt(TEdit(Sender).Text)]);;
      if Data_Module.GetLastSeqDate(Data_Module.fiRevSeqLookup.AsInteger) then
      begin
        if TEdit(Sender).Name='StartSeqNo_Edit' then
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
          if Shipping_Form.Visible then
            LastSeqNo_Edit.SetFocus;
        end
        else
        begin
          EndBox.Items.Clear;
          while not Data_Module.ALC_StoredProc.Eof do
          begin
            if {StrToDateTime(}StartBox.Text{)} < FormatDateTime('yyyy-mm-dd hh:mm:ss.zzz',Data_Module.ALC_StoredProc.fieldbyname('LastTime').AsDateTime) then
            begin
              EndBox.Items.Add(FormatDateTime('yyyy-mm-dd hh:mm:ss.zzz',Data_Module.ALC_StoredProc.fieldbyname('LastTime').AsDateTime));
            end;
            Data_Module.ALC_StoredProc.Next;
          end;
          if EndBox.Items.Count > 0 then
            EndBox.ItemIndex:=0
          else
          begin
            showMessage('Start date is after end date, changing to previous date for starting sequence.');
            StartBox.ItemIndex:=StartBox.ItemIndex+1;
            StartSeqNo_EditChange(sender);
          end;
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
  end;
end;

procedure TShipping_Form.FormShow(Sender: TObject);
begin
  if StartSeqNo_Edit.Text = '' then
    StartSeqNo_Edit.SetFocus
  else
    LastSeqNo_Edit.SetFocus;
end;


procedure TShipping_Form.Line_ComboBoxChange(Sender: TObject);
begin
  //
  // Get next production date to process, base on last process and overtime/holiday info
  //
  Data_Module.LineName:=Line_ComboBox.Text;
  SetDetailBoxes;
end;

procedure TShipping_Form.UpdateShipping_ButtonClick(Sender: TObject);
begin
  try
    ModifyShipping_Form:=TModifyShipping_Form.Create(self);
    ModifyShipping_Form.LineName:=Line_ComboBox.Text;
    ModifyShipping_Form.ProductionDate:=Production_DateTimePicker.Date;
    ModifyShipping_Form.Execute;
    ModifyShipping_Form.Free;
  except
    on e:exception do
    begin
      ShowMessage('Unable to update shipping, '+e.Message);
      Data_Module.LogActLog('SHIPPING','Unable to update shipping, '+e.Message);
    end;
  end;
end;

end.
