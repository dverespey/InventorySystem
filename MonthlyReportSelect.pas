//****************************************************************
//
//       Inventory Control
//
//       Copyright (c) 2002-2003 Failproof Manufacturing Systems.
//
//****************************************************************
//
//  05/02/2003  David Verespey  Initial creation

unit MonthlyReportSelect;

interface

uses Windows, SysUtils, Classes, Graphics, Forms, Controls, StdCtrls,
  Buttons, ExtCtrls, BmDtEdit, NUMMIBmDateEdit, ADODB, Dialogs;

type
  TDateSelectDlg = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    Bevel1: TBevel;
    Start_NUMMIBmDateEdit: TNUMMIBmDateEdit;
    End_NUMMIBmDateEdit: TNUMMIBmDateEdit;
    Label1: TLabel;
    End_Label: TLabel;
    Supplier_Label: TLabel;
    Supplier_ComboBox: TComboBox;
    Logistics_Label: TLabel;
    Logistics_ComboBox: TComboBox;
    Partnumber_Label: TLabel;
    Partnumber_ComboBox: TComboBox;
    procedure CancelBtnClick(Sender: TObject);
    procedure Start_NUMMIBmDateEditExit(Sender: TObject);
  private
    { Private declarations }
    fStartDate:TDate;
    fEndDate:TDate;
    fCancel:boolean;
    fSupplier:string;
    fDoSupplier:boolean;
    fLogistics:string;
    fDoLogistics:boolean;
    fPartNumber:string;
    fDoPartNumber:boolean;
    fJustStart:boolean;
  public
    { Public declarations }
    procedure Execute;

    property FromDate:Tdate
    read fStartDate
    write fStartDate;

    property JustStart:boolean
    read fJustStart
    write fJustStart
    default FALSE;

    property ToDate:TDate
    read fEndDate
    write fEndDate;

    property Cancel:Boolean
    read fCancel
    write fCancel;

    property Supplier:string
    read fSupplier
    write fSupplier;

    property DoSupplier:boolean
    read fDoSupplier
    write fDoSupplier
    default FALSE;

    property Logistics:string
    read fLogistics
    write fLogistics;

    property DoLogistics:boolean
    read fDoLogistics
    write fDoLogistics
    default FALSE;

    property Partnumber:string
    read fPartNumber
    write fPartNumber;

    property DoPartNumber:boolean
    read fDoPartNumber
    write fDoPartNumber
    default FALSE;
  end;

var
  DateSelectDlg: TDateSelectDlg;

implementation

uses DataModule;

{$R *.dfm}

procedure TDateSelectDlg.Execute;
begin
  try
    fCancel:=False;
    if fDoPartNumber then
    begin
      Partnumber_ComboBox.Visible:=TRUE;
      Partnumber_Label.Visible:=TRUE;
      Partnumber_ComboBox.Items.Clear;
      Partnumber_ComboBox.Items.Add('ALL');


      with Data_Module.Inv_DataSet do
      begin
        Close;
        CommandType := CmdText;
        CommandText := 'SELECT VC_PART_NUMBER from INV_PARTS_STOCK_MST';
        Open;
        while not EOF do
        begin
          Partnumber_ComboBox.Items.Add(fieldbyname('VC_PART_NUMBER').AsString);
          Next;
        end;
      end;
      Partnumber_ComboBox.Text:='';
    end
    else
    begin
      Partnumber_ComboBox.Visible:=FALSE;
      PArtnumber_Label.Visible:=FALSE;
    end;

    if fDoSupplier then
    begin
      Supplier_Label.Visible:=True;
      Supplier_ComboBox.Visible:=True;
      Supplier_ComboBox.Items.Clear;
      Supplier_ComboBox.Items.Add('ALL');


      with Data_Module.Inv_DataSet do
      begin
        Close;
        CommandType := CmdText;
        CommandText := 'SELECT distinct(VC_SUPPLIER_NAME) from INV_SUPPLIER_MST';
        Open;
        while not EOF do
        begin
          Supplier_ComboBox.Items.Add(fieldbyname('VC_SUPPLIER_NAME').AsString);
          Next;
        end;
      end;
      Supplier_ComboBox.Text:='';
    end
    else
    begin
      Supplier_Label.Visible:=False;
      Supplier_ComboBox.Visible:=False;
    end;

    if fDoLogistics then
    begin
      Logistics_Label.Visible:=True;
      Logistics_ComboBox.Visible:=True;

      Logistics_ComboBox.Items.Clear;
      Logistics_ComboBox.Items.Add('ALL');

      with Data_Module.Inv_DataSet do
      begin
        Close;
        CommandType := CmdText;
        CommandText := 'SELECT distinct(VC_LOGISTICS_NAME) from INV_LOGISTICS_MST';
        Open;
        while not EOF do
        begin
          Logistics_ComboBox.Items.Add(fieldbyname('VC_LOGISTICS_NAME').AsString);
          Next;
        end;
      end;
      Logistics_ComboBox.Text:='';

    end
    else
    begin
      Logistics_Label.Visible:=False;
      Logistics_ComboBox.Visible:=False;
    end;

    if fJustStart then
    begin
      End_Label.Visible:=FALSE;
      End_NUMMIBmDateEdit.Visible:=FALSE;
    end;

    ShowModal;
    fStartDate:=Start_NUMMIBmDateEdit.Date;
    fEndDate:=End_NUMMIBmDateEdit.Date;
    if Supplier_ComboBox.Text = '' then
      fSupplier:='ALL'
    else
      fSupplier:=Supplier_ComboBox.Text;

    if Logistics_ComboBox.Text = '' then
      fLogistics:='ALL'
    else
      fLogistics:=Logistics_ComboBox.Text;

    if Partnumber_ComboBox.Text = '' then
      fPArtNumber:='ALL'
    else
      fPArtNumber:=Partnumber_ComboBox.Text;
  except
    on e:exception do
    begin
      ShowMessage('Unable to create select date box');
      Data_Module.LogActLog('ERROR','Unable to do get date form,'+e.Message);
      fCancel:=TRUE;
    end;
  end;
end;

procedure TDateSelectDlg.CancelBtnClick(Sender: TObject);
begin
  fCancel:=True;
end;

procedure TDateSelectDlg.Start_NUMMIBmDateEditExit(Sender: TObject);
begin
  End_NUMMIBmDateEdit.Date := Start_NUMMIBmDateEdit.Date;
end;

end.
