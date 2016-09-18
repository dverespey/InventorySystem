//***********************************************************
//  InventorySystem
//
//  Copyright (c) 2002 Failproof Manufacturing Systems
//
//***********************************************************
//
//  10/31/2002  Aaron Huge  Initial creation

unit Logon;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, Buttons, DataModule;  //IniFiles;

type
  TLogon_Form = class(TForm)
    Panel2: TPanel;
    Password_Label: TLabel;
    bitBtnLogon: TBitBtn;
    BitBtnCancel: TBitBtn;
    Password_Edit: TEdit;
    UserID_Label: TLabel;
    UserID_Edit: TEdit;
    procedure bitBtnLogonClick(Sender: TObject);
    procedure EditKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
  public
    { Public declarations }
    function Execute: boolean;
  end;

var
  Logon_Form: TLogon_Form;

implementation

{$R *.DFM}

procedure TLogon_Form.bitBtnLogonClick(Sender: TObject);
begin
  If (Trim(UserID_Edit.Text) = '') Or (Trim(Password_Edit.Text) = '') Then
  Begin
    MessageDlg('Either or both of the ID and Password boxes are blank.' +
                #13 + 'Please try again.', mtInformation, [mbOK], 0);
    UserID_Edit.SetFocus;
  End
  Else
  Begin
    If Data_Module.ValidateUser(Trim(UserID_Edit.Text), Trim(Password_Edit.Text)) Then
    Begin
      ModalResult := mrOK;
    End
    Else
    Begin
      UserID_Edit.Text := '';
      Password_Edit.Text := '';
      UserID_Edit.SetFocus;
      MessageDlg('The user ID and/or password you entered is invalid.' +
                  #13 + 'Please try again.', mtInformation, [mbOK], 0);
    End;
  End;
end;

function TLogon_Form.Execute: boolean;
Begin
  Result:= True;
  UserID_Edit.Text := '';
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

end;      //Execute

procedure TLogon_Form.EditKeyPress(Sender: TObject;
  var Key: Char);
begin
  If Key in ['a'..'z'] then Dec(Key,32); {Force Uppercase only!}
end;

end.
