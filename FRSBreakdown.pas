unit FRSBreakdown;

interface

uses Windows, SysUtils, Classes, Graphics, Forms, Controls, StdCtrls,
  Buttons, ExtCtrls,Dialogs;

type
  TFRSBreakdownDlg = class(TForm)
    OK_Button: TButton;
    Cancel_Button: TButton;
    Bevel1: TBevel;
    Label1: TLabel;
    FRSCount_ComboBox: TComboBox;
    Label2: TLabel;
    PartNumber_Edit: TEdit;
    Label3: TLabel;
    qty_Edit: TEdit;
    Split_Memo: TMemo;
    Label4: TLabel;
    Label5: TLabel;
    LotSize_Edit: TEdit;
    Label6: TLabel;
    Lots_Edit: TEdit;
    procedure Cancel_ButtonClick(Sender: TObject);
    procedure OK_ButtonClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FRSCount_ComboBoxChange(Sender: TObject);
  private
    { Private declarations }
    fCancel:boolean;
    fpartnumber:string;
    fqty:integer;
    fLotSize:integer;
    fLots:integer;
    fSplit:string;
    fBreakdownCount:integer;
  public
    { Public declarations }
    procedure Execute;

    property Cancel:boolean
    read fCancel
    write fCancel;

    property BreakdownCount:integer
    read fBreakdownCount
    write fBreakdownCount;

    property partnumber:string
    read fpartnumber
    write fpartnumber;

    property qty: integer
    read fqty
    write fqty;

    property LotSize: integer
    read fLotSize
    write fLotSize;

    property Lots: integer
    read fLots
    write fLots;

    property Split: string
    read fSplit
    write fSplit;
  end;

var
  FRSBreakdownDlg: TFRSBreakdownDlg;

implementation

{$R *.dfm}

procedure TFRSBreakdownDlg.Execute;
var
  x,i:integer;
begin
  fCancel:=False;
  PartNumber_Edit.Text:=fpartnumber;
  qty_edit.Text:=IntToStr(fqty);
  LotSize_Edit.Text:=IntToStr(fLotSize);
  Lots_Edit.Text:=IntToStr(fLots);
  split_memo.Text:='';
  FRSCount_ComboBox.Items.Clear;
  x:=1;
  for i:=2 to fqty div flotsize do
  begin
    FRSCount_ComboBox.Items.Add(IntToStr(i));
    x:=i;
  end;
  if fqty mod flotsize <> 0 then
    FRSCount_ComboBox.Items.Add(IntToStr(x+1));
  ShowModal;
end;

procedure TFRSBreakdownDlg.Cancel_ButtonClick(Sender: TObject);
begin
  FCancel:=True;
  close;
end;

procedure TFRSBreakdownDlg.OK_ButtonClick(Sender: TObject);
begin
  try
    fBreakdownCount:=StrToInt(FRSCount_ComboBox.Text);
    Close;
  except
    on e:exception do
    begin
      SHowMessage('Value must be a numeric');
    end;
  end;
end;

procedure TFRSBreakdownDlg.FormShow(Sender: TObject);
begin
  FRSCount_ComboBox.SetFocus;
end;

procedure TFRSBreakdownDlg.FRSCount_ComboBoxChange(Sender: TObject);
var
  modt,i:integer;
begin
  Split_Memo.Text:='';
  fSplit:='';
  modt:=flots mod StrToInt(FRSCount_ComboBox.Text);

  for i:=1 to StrToInt(FRSCount_ComboBox.Text) do
  begin
    if fSplit='' then
      fSplit:=IntToStr((flots div StrToInt(FRSCount_ComboBox.Text)) + modt)
    else
      fSplit:=fSplit+','+IntToStr(flots div StrToInt(FRSCount_ComboBox.Text));
  end;

  split_memo.Text:=fSplit;
end;

end.
