//***********************************************************
//  InventorySystem
//
//  Copyright (c) 2002 Failproof Manufacturing Systems
//
//***********************************************************
//
//  10/31/2002  Aaron Huge  Initial creation

unit ConfirmPassword;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, Buttons, DataModule;  //IniFiles;

type
  TConfirmPassword_Form = class(TForm)
    Panel2: TPanel;
    Password_Label: TLabel;
    bitBtnLogon: TBitBtn;
    BitBtnCancel: TBitBtn;
    Password_Edit: TEdit;
    procedure bitBtnLogonClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    function Execute: boolean;
  end;

var
  ConfirmPassword_Form: TConfirmPassword_Form;

implementation

{$R *.DFM}

procedure TConfirmPassword_Form.bitBtnLogonClick(Sender: TObject);
begin
  If (Trim(Password_Edit.Text) = '') Then
  Begin
    MessageDlg('The Password box is blank.' +
                #13 + 'Please try again.', mtInformation, [mbOK], 0);
    Password_Edit.SetFocus;
  End
  Else
  Begin
    If Data_Module.gobjUser.AppUserPass = Trim(Password_Edit.Text) Then
    Begin
      ModalResult := mrOK;
      Data_Module.LogActLog('USER ADMIN', 'Confirmed password to enter User Administration.', 0);
    End
    Else
    Begin
      Data_Module.LogActLog('ADMIN ERR', 'FAILED to match password to enter User Administration.', 0);
      Password_Edit.Text := '';
      Password_Edit.SetFocus;
      MessageDlg('The password you entered is invalid.' +
                  #13 + 'Please try again.', mtInformation, [mbOK], 0);
    End;
  End;
end;

function TConfirmPassword_Form.Execute: boolean;
Begin
  Result:= True;
  Password_Edit.Text:='';
  try
    try
      ShowModal;
    except
      On E:Exception do
        showMessage('Unable to generate logon screen.'
          + #13 + 'ERROR:'
          + #13 + E.Message);
    end;
  finally
    If ModalResult = mrCancel Then
      Result:= False;
  end;

end;

end.
