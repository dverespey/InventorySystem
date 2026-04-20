//****************************************************************
//
//       Inventory Control
//
//       Copyright (c) 2006 Failproof Manufacturing Systems.
//
//****************************************************************
//
//  10/25/2002  David Verespey      Initial creation

unit ManualShipping;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, ComCtrls, adodb;

type
  TManualShipping_Form = class(TForm)
    Parts_StringGrid: TStringGrid;
    PartNumber_Edit: TEdit;
    Label1: TLabel;
    Count_Edit: TEdit;
    Label2: TLabel;
    Update_Button: TButton;
    Close_Button: TButton;
    Post_Button: TButton;
    Statistics_GroupBox: TGroupBox;
    Date_Label: TLabel;
    Production_DateTimePicker: TDateTimePicker;
    StartSeqNo_Label: TLabel;
    StartSeqNo_Edit: TEdit;
    Label4: TLabel;
    LastSeqNo_Edit: TEdit;
    Label3: TLabel;
    Label5: TLabel;
    DailyTotal_Edit: TEdit;
    IrregularShipCount_Label: TLabel;
    IrregularShipCount_Edit: TEdit;
    IrregularShip_Button: TButton;
    Line_ComboBox: TComboBox;
    procedure Post_ButtonClick(Sender: TObject);
    procedure Close_ButtonClick(Sender: TObject);
    procedure Production_DateTimePickerChange(Sender: TObject);
    procedure Update_ButtonClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Parts_StringGridSelectCell(Sender: TObject; ACol,
      ARow: Integer; var CanSelect: Boolean);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure IrregularShip_ButtonClick(Sender: TObject);
  private
    { Private declarations }
    fUpdateShipping:boolean;
    fLineName:string;
    fProductionDate:TDateTime;
    fChanged:boolean;

    fNoDateTimeUpdate:boolean;
    procedure SetDetailBoxes;
  public
    { Public declarations }
    function Execute: boolean;

    property UpdateShipping:boolean
    read fUpdateShipping
    write fUpdateShipping
    Default FALSE;

    property LineName:string
    read fLineName
    write fLineName;

    property ProductionDate:TDateTime
    read fProductionDate
    write fProductionDate;
  end;

var
  ManualShipping_Form: TManualShipping_Form;

implementation

uses DataModule;

{$R *.dfm}

function TManualShipping_Form.Execute: boolean;
begin
  Result:= True;
  fChanged:=FALSE;

  Parts_StringGrid.Cells[0,0]:='Part Number';
  Parts_StringGrid.Cells[1,0]:='Supplier';
  Parts_StringGrid.Cells[2,0]:='Part Description';
  Parts_StringGrid.Cells[3,0]:='Daily Build Count';

  try
    try
    // If redo then just add stuff
      if fUpdateShipping then
      begin
        Line_ComboBox.Items.Clear;
        Line_ComboBox.Items.Add(fLineName);
        Line_ComboBox.ItemIndex:=0;

        Production_DateTimePicker.Date := fProductionDate;
        Data_Module.ProductionDateDT:= fProductionDate;
        SetDetailBoxes;
        Statistics_GroupBox.Enabled:=FALSE;

        IrregularShipCount_Label.Visible:=FALSE;
        IrregularShipCount_Edit.Visible:=FALSE;
        IrregularShip_Button.Visible:=FALSE;
        Post_Button.Visible:=FALSE;


      end
      else
      begin
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

        //
        // Get next production date to process, base on last process and overtime/holiday info
        //
        Data_Module.LineName:=Line_ComboBox.Text;

        Data_module.GetNextProductionDate;
        Production_DateTimePicker.Date := Data_Module.ProductionDateDT;
        SetDetailBoxes;
      end;

      ShowModal;
    except
      On E:Exception do
        showMessage('Error on get information manual shipping screen.'
          + #13 + 'ERROR:'
          + #13 + E.Message);
    end;
  finally
    If ModalResult = mrCancel Then
      Result:= False;
  end;

end;

procedure TManualShipping_Form.SetDetailBoxes;
var
  i,seq:integer;
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
      DailyTotal_Edit.Text := IntToStr(Inv_DataSet.FieldByName('IN_QTY').AsInteger);



      // Fill Parts List
      GetPartsListCount;

      if INV_Dataset.RecordCount = 0 then
        Parts_StringGrid.RowCount:=2
      else
        Parts_StringGrid.RowCount:=INV_Dataset.RecordCount+1;

      if INV_Dataset.RecordCount <> 0 then
      begin
        i:=1;
        while not INV_Dataset.Eof do
        begin
          Parts_StringGrid.Cells[0,i]:=INV_Dataset.fieldByName('vc_part_number').AsString;
          Parts_StringGrid.Cells[1,i]:=INV_Dataset.fieldByName('vc_supplier_name').AsString;
          Parts_StringGrid.Cells[2,i]:=INV_Dataset.fieldByName('vc_parts_name').AsString;
          Parts_StringGrid.Cells[3,i]:=INV_Dataset.fieldByName('in_qty').AsString;
          INV_Dataset.Next;
          INC(i);
        end;
      end
      else
      begin
          Parts_StringGrid.Cells[0,1]:='';
          Parts_StringGrid.Cells[1,1]:='';
          Parts_StringGrid.Cells[2,1]:='';
          Parts_StringGrid.Cells[3,1]:='';
      end;


      StartSeqNo_Edit.ReadOnly:=True;
      LastSeqNo_Edit.ReadOnly:=True;
      Count_Edit.ReadOnly:=TRUE;

      Update_Button.Enabled:=FALSE;
      Post_Button.Enabled:=FALSE;

      PartNumber_Edit.Text:='';
      Count_Edit.Text:='';

      IrregularShipCount_Label.Visible:=TRUE;
      IrregularShipCount_Edit.Visible:=TRUE;
      IrregularShip_Button.Visible:=TRUE;
    end
    else
    begin
      // Last date found
      if Inv_DataSet.FieldByName('VC_END_SEQ_NUMBER').AsString <> '' then
      begin
        if StartSeqNo_Edit.MaxLength = 3 then
        begin
          seq:=StrToInt(Inv_DataSet.FieldByName('VC_END_SEQ_NUMBER').AsString);
          inc(seq);
          if seq > 999 then
            seq:=0;
          StartSeqNo_Edit.Text := format('%.3d',[seq]);
        end
        else
        begin
          //SeqNum.SequenceNumber:=Inv_DataSet.FieldByName('VC_END_SEQ_NUMBER').AsString;
          //SeqNum.Increment;
          //StartSeqNo_Edit.Text := SeqNum.SequenceNumber;
        end;

      end
      else
      begin
        StartSeqNo_Edit.Text := format('%.3d',[0]);
      end;


      // Fill Parts List

      GetPartsList;

      Parts_StringGrid.RowCount:=INV_Dataset.RecordCount+1;

      i:=1;
      while not INV_Dataset.Eof do
      begin
        Parts_StringGrid.Cells[0,i]:=INV_Dataset.fieldByName('vc_part_number').AsString;
        Parts_StringGrid.Cells[1,i]:=INV_Dataset.fieldByName('vc_supplier_name').AsString;
        Parts_StringGrid.Cells[2,i]:=INV_Dataset.fieldByName('vc_parts_name').AsString;
        Parts_StringGrid.Cells[3,i]:='0';
        INV_Dataset.Next;
        INC(i);
      end;

      PartNumber_Edit.Text:=Parts_StringGrid.Cells[0,1];
      Count_Edit.Text:='';

      if ManualShipping_Form.Visible then
        LastSeqNo_Edit.SetFocus;

      Parts_StringGrid.Row:=1;

      LastSeqNo_Edit.Text := '';
      DailyTotal_Edit.Text := '';
      StartSeqNo_Edit.ReadOnly:=False;
      LastSeqNo_Edit.ReadOnly:=False;
      Count_Edit.ReadOnly:=FALSE;

      Update_Button.Enabled:=TRUE;
      Post_Button.Enabled:=TRUE;

      IrregularShipCount_Label.Visible:=FALSE;
      IrregularShipCount_Edit.Visible:=FALSE;
      IrregularShip_Button.Visible:=FALSE;

    end;
  End;        //With
  fChanged:=FALSE;
End;        //SetDetailBoxes

procedure TManualShipping_Form.Post_ButtonClick(Sender: TObject);
var
  i,x:integer;
begin
  if fchanged then
  begin
    if StrToInt(DailyTotal_Edit.Text) > 0 then
    begin
      if not tryStrToInt(StartSeqNo_Edit.Text,i) then
      begin
        ShowMessage('Starting sequence invalid');
        StartSeqNo_Edit.Text:='';
        StartSeqNo_Edit.SetFocus;
        exit;
      end;

      if not tryStrToInt(LastSeqNo_Edit.Text,i) then
      begin
        ShowMessage('Last sequence invalid');
        LastSeqNo_Edit.Text:='';
        LastSeqNo_Edit.SetFocus;
        exit;
      end;

      With Data_Module do
      Begin
        Inv_Connection.BeginTrans;

        StartSeq:=format('%.3d',[StrToInt(StartSeqNo_Edit.Text)]);
        LastSeq:=format('%.3d',[StrToInt(LastSeqNo_Edit.Text)]);
        ProductionDate:=formatdatetime('yyyymmdd',Production_DateTimePicker.DateTime);
        Quantity := StrToInt(Trim(DailyTotal_Edit.Text));
        Continuation := 0;


        If Not InsertShippingInfoManual Then
        begin
          Inv_Connection.RollbackTrans;
          ShowMessage('Failed to update shipping info.')
        end
        else
        begin
          for x:=1 to Parts_StringGrid.RowCount-1 do
          begin
            // Add each item
            IrregulareQty:=0;
            PartNum:=Parts_StringGrid.Cells[0,x];
            Quantity:=StrToInt(Parts_StringGrid.Cells[3,x]);
            if not InsertShippingDetailManual then
            begin
              Inv_Connection.RollbackTrans;
              ShowMessage('Failed to update shipping info.');
              exit;
            end;
          end;

          Inv_Connection.CommitTrans;
          ShowMessage('Shipping info update is complete.');
          // reset screen
          Data_Module.ProductionDateDT:=Production_DateTimePicker.Date;
          SetDetailBoxes;
        end;
      End;      //With

    end
    else
    begin
        ShowMessage('Count is 0, no data to post');
        fchanged:=FALSE;
    end;
  end
  else
  begin
    ShowMessage('No updated entries');
  end;
end;

procedure TManualShipping_Form.Close_ButtonClick(Sender: TObject);
begin
  Close;
end;

procedure TManualShipping_Form.Production_DateTimePickerChange(
  Sender: TObject);
begin
  // Check if that date has already been run
  if not fNoDateTimeUpdate then
  begin
    fNoDateTimeUpdate:=TRUE;

    //
    // Get next production date to process, base on last process and overtime/holiday info
    //
    Data_Module.ProductionDateDT:=Production_DateTimePicker.Date;

    SetDetailBoxes;

    fNoDateTimeUpdate:=FALSE;
  end;
end;

procedure TManualShipping_Form.Update_ButtonClick(Sender: TObject);
var
  x,count:integer;
begin
  // Update Count
  if fUpdateShipping then
  begin
    try
    except
    end;
  end
  else
  begin
    Count_Edit.SetFocus;

    if trystrtoInt(Count_Edit.Text,x) then
    begin
      Parts_StringGrid.Cells[3,Parts_StringGrid.Row]:=Count_Edit.Text;

      if Parts_StringGrid.Row < Parts_StringGrid.RowCount-1 then
        Parts_StringGrid.Row:=Parts_StringGrid.Row+1
      else
        ShowMessage('End of parts list');

      Count_Edit.Text:='';
      Count_Edit.SetFocus;

      fChanged:=TRUE;

      count:=0;
      for x:=1 to Parts_StringGrid.RowCount-1 do
      begin
        count:=count+StrToInt(Parts_StringGrid.Cells[3,x]);
      end;

      DailyTotal_Edit.Text:=IntTostr(count);
    end
    else
    begin
      ShowMessage('Invalid count');
      Count_Edit.Text:='';
      Count_Edit.SetFocus;
    end;
  end;
end;

procedure TManualShipping_Form.FormShow(Sender: TObject);
begin
  Count_Edit.SetFocus;
end;

procedure TManualShipping_Form.Parts_StringGridSelectCell(Sender: TObject;
  ACol, ARow: Integer; var CanSelect: Boolean);
begin
  PartNumber_Edit.Text:=Parts_StringGrid.Cells[0,aRow];
  Count_Edit.Text:=Parts_StringGrid.Cells[3,aRow];
end;

procedure TManualShipping_Form.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  if fChanged then
  begin
    if messagedlg('Close and cancel changes?',mtConfirmation, [mbYes, mbNo], 0) = mrNo then
      CanClose:=FALSE
    else
      CanClose:=TRUE;
  end
end;

procedure TManualShipping_Form.IrregularShip_ButtonClick(Sender: TObject);
var
  i:integer;
begin
  With Data_Module do
  Begin
    if PartNumber_Edit.Text = '' then
    begin
      ShowMessage('Must select a partnumber');
      exit;
    end;
    if not tryStrToInt(IrregularShipCount_Edit.Text,i) then
    begin
      ShowMessage('Irregular ship count invalid');
      IrregularShipCount_Edit.Text:='';
      IrregularShipCount_Edit.SetFocus;
      exit;
    end;
    // Add each item
    ProductionDate:=formatdatetime('yyyymmdd',Production_DateTimePicker.Date);
    IrregulareQty:=1;
    PartNum:=PartNumber_Edit.Text;
    Quantity:=StrToInt(IrregularShipCount_Edit.Text);
    if not InsertShippingDetailManual then
    begin
      ShowMessage('Failed to update shipping info.');
      exit;
    end;

    IrregularShipCount_Edit.Text:='';
    SetDetailBoxes;
  end;
end;

end.
