unit SelectDateRange;

interface

uses Windows, SysUtils, Classes, Graphics, Forms, Controls, StdCtrls, 
  Buttons, ExtCtrls, ComCtrls;

type
  TSelectDateRangeDlg = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    Bevel1: TBevel;
    Label1: TLabel;
    Label2: TLabel;
    FromDateTimePicker: TDateTimePicker;
    ToDateTimePicker: TDateTimePicker;
    procedure OKBtnClick(Sender: TObject);
  private
    { Private declarations }
    fFromMonth: String;
    fToMonth: String;
    fCancel:boolean;
  public
    { Public declarations }
    property FromMonth: string
    read fFromMonth;

    property ToMonth:string
    read fToMonth;

    property Cancel:boolean
    read fCancel;

    procedure Execute();
  end;

var
  SelectDateRangeDlg: TSelectDateRangeDlg;

implementation

{$R *.dfm}

procedure TSelectDateRangeDlg.Execute();
begin
  fcancel:=TRUE;
  FromDateTimePicker.DateTime:=now;
  ToDateTimePicker.DateTime:=now;
  ShowModal;
end;


procedure TSelectDateRangeDlg.OKBtnClick(Sender: TObject);
begin
  fCancel:=FALSE;

  fFromMonth:=FormatDateTime('YYYY-MM-DD',FromDateTimePicker.DateTime);
  fToMonth:=FormatDateTime('YYYY-MM-DD',ToDateTimePicker.DateTime);
end;

end.
