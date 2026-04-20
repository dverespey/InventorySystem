unit NewPassword;

interface

uses Windows, SysUtils, Classes, Graphics, Forms, Controls, StdCtrls,
  Buttons, dialogs;

type
  TNewPasswordDlg = class(TForm)
    Label1: TLabel;
    NewPassword: TEdit;
    OKBtn: TButton;
    CancelBtn: TButton;
    Label2: TLabel;
    ConfirmPAssword: TEdit;
    procedure FormShow(Sender: TObject);
    procedure OKBtnClick(Sender: TObject);
  private
    { Private declarations }
    fCancel:boolean;
    fpassword:string;
  public
    { Public declarations }
    procedure Execute;

    property Password:string
    read fpassword
    write fpassword;

    property Cancel:boolean
    read fCancel
    write fCancel;
  end;

var
  NewPasswordDlg: TNewPasswordDlg;

implementation

{$R *.dfm}
procedure TNewPasswordDlg.Execute;
begin
  ShowModal;
end;

procedure TNewPasswordDlg.FormShow(Sender: TObject);
begin
  fCancel:=TRUE;
  NewPassword.SetFocus;
end;

procedure TNewPasswordDlg.OKBtnClick(Sender: TObject);
begin
  if (NewPassword.Text = ConfirmPassword.Text) AND (NewPAssword.Text<>'')then
  begin
    fPassword:=NewPAssword.Text;
    fCancel:=FALSE;
    Close;
  end
  else
  begin
    SHowMessage('Password does not match');
    ConfirmPAssword.Text:='';
    ConfirmPAssword.SetFocus;
  end;
end;

end.

