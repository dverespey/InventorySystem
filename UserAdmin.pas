//***********************************************************
//  InventorySystem
//
//  Copyright (c) 2002 Failproof Manufacturing Systems
//
//***********************************************************
//
//  10/31/2002  Aaron Huge  Initial creation

unit UserAdmin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, DataModule;

type
  TUserAdmin_Form = class(TForm)
    Buttons_Panel: TPanel;
    Insert_Button: TButton;
    Update_Button: TButton;
    Clear_Button: TButton;
    Close_Button: TButton;
    Delete_Button: TButton;
    UserAdmin_Panel: TPanel;
    IsAdmin_CheckBox: TCheckBox;
    ConfirmPassword_Label: TLabel;
    Password_Label: TLabel;
    Label1: TLabel;
    ExistingUsers_Label: TLabel;
    ExistingUserIDs_ComboBox: TComboBox;
    UserID_Edit: TEdit;
    Password_Edit: TEdit;
    ConfirmPassword_Edit: TEdit;
    procedure EditKeyPress(Sender: TObject; var Key: Char);
    procedure FormCreate(Sender: TObject);
    procedure ExistingUserIDs_ComboBoxChange(Sender: TObject);
    procedure Clear_ButtonClick(Sender: TObject);
    procedure Insert_ButtonClick(Sender: TObject);
    procedure Update_ButtonClick(Sender: TObject);
    procedure Delete_ButtonClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    function Execute: boolean;
  end;

var
  UserAdmin_Form: TUserAdmin_Form;

implementation

{$R *.dfm}

function TUserAdmin_Form.Execute: boolean;
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

procedure TUserAdmin_Form.EditKeyPress(Sender: TObject;
  var Key: Char);
begin
  If Key in ['a'..'z'] then Dec(Key,32); {Force Uppercase only!}
end;

procedure TUserAdmin_Form.FormCreate(Sender: TObject);
begin
  Data_Module.SetComboBoxesWithUserObj(ExistingUserIDs_ComboBox, 'SELECT VC_USER_ID, VC_PASSWORD, BIT_ADMIN FROM INV_USERS ORDER BY VC_USER_ID, VC_PASSWORD');
end;

procedure TUserAdmin_Form.ExistingUserIDs_ComboBoxChange(Sender: TObject);
var
  UserAdminDetail: TUserAdminDetail;
begin
  If ExistingUserIDs_ComboBox.ItemIndex = 0 Then
    Data_Module.ClearControls(UserAdmin_Panel)
  Else
  Begin
    UserAdminDetail := ExistingUserIDs_ComboBox.Items.Objects [ExistingUserIDs_ComboBox.itemIndex] as TUserAdminDetail;
    UserID_Edit.Text := UserAdminDetail.ID;
    Password_Edit.Text := UserAdminDetail.Pass;
    IsAdmin_CheckBox.Checked := UserAdminDetail.Admin;
    ConfirmPassword_Edit.Text := '';
  End;          //IF...Else
end;

procedure TUserAdmin_Form.Clear_ButtonClick(Sender: TObject);
begin
  ExistingUserIDs_ComboBox.ItemIndex := 0;
  Data_Module.ClearControls(UserAdmin_Panel);
  ExistingUserIDs_ComboBox.SetFocus;
end;

procedure TUserAdmin_Form.Insert_ButtonClick(Sender: TObject);
begin

  If (Trim(UserID_Edit.Text) = '') Or (Trim(Password_Edit.Text) = '') Or (Trim(ConfirmPassword_Edit.Text) = '') Then
    MessageDlg('Please fill in all boxes.', mtInformation, [mbOK], 0)
  Else
  Begin
    If Trim(Password_Edit.Text) = Trim(ConfirmPassword_Edit.Text) Then
    Begin
      If Not Data_Module.InsertUser(Trim(UserID_Edit.Text), Trim(Password_Edit.Text), IsAdmin_CheckBox.Checked) Then
        MessageDlg('This User ID and Password already exist.' + #13 + 'Please try again.', mtError, [mbOK], 0)
      Else
      begin
        Data_Module.SetComboBoxesWithUserObj(ExistingUserIDs_ComboBox, 'SELECT VC_USER_ID, VC_PASSWORD, BIT_ADMIN FROM INV_USERS ORDER BY VC_USER_ID, VC_PASSWORD');
        ExistingUserIDs_ComboBox.ItemIndex := 0;
        Data_Module.ClearControls(UserAdmin_Panel);
        ExistingUserIDs_ComboBox.SetFocus;
      end;
    End
    Else
      MessageDlg('Please make sure that the password' + #13 + 'and confirm password fields match.', mtError, [mbOK], 0);
  End;
end;      //Insert_ButtonClick

procedure TUserAdmin_Form.Update_ButtonClick(Sender: TObject);
var
  UserAdminDetail: TUserAdminDetail;
begin
  If (Trim(UserID_Edit.Text) = '') Or (Trim(Password_Edit.Text) = '') Or (Trim(ConfirmPassword_Edit.Text) = '') Then
    MessageDlg('Please fill in all boxes.', mtInformation, [mbOK], 0)
  Else
  Begin
    If Trim(Password_Edit.Text) = Trim(ConfirmPassword_Edit.Text) Then
    Begin
      UserAdminDetail := ExistingUserIDs_ComboBox.Items.Objects [ExistingUserIDs_ComboBox.itemIndex] as TUserAdminDetail;
      Data_Module.UpdateUserInfo(UserAdminDetail.ID, UserAdminDetail.Pass, UserID_Edit.Text, Password_Edit.Text, IsAdmin_Checkbox.Checked);
      Data_Module.SetComboBoxesWithUserObj(ExistingUserIDs_ComboBox, 'SELECT VC_USER_ID, VC_PASSWORD, BIT_ADMIN FROM INV_USERS ORDER BY VC_USER_ID, VC_PASSWORD');
    End
    Else
      MessageDlg('Please make sure that the password' + #13 + 'and confirm password fields match.', mtError, [mbOK], 0);
  End;
end;        //Update_ButtonClick

procedure TUserAdmin_Form.Delete_ButtonClick(Sender: TObject);
var
  UserAdminDetail: TUserAdminDetail;
begin
  If Trim(ExistingUserIDs_ComboBox.Text) = '' Then
    MessageDlg('Please select a user from the drop down list prior to pressing delete.', mtInformation, [mbOK], 0)
  Else
  Begin
    UserAdminDetail := ExistingUserIDs_ComboBox.Items.Objects [ExistingUserIDs_ComboBox.itemIndex] as TUserAdminDetail;
    If MessageDlg('Are you sure you wish to delete ' + UserAdminDetail.ID + '?', mtConfirmation, [mbOK, mbCancel], 0) = mrOK Then
    Begin
      Data_Module.DeleteUserInfo(UserAdminDetail.ID, UserAdminDetail.Pass);
      Data_Module.ClearControls(UserAdmin_Panel);
      Data_Module.SetComboBoxesWithUserObj(ExistingUserIDs_ComboBox, 'SELECT VC_USER_ID, VC_PASSWORD, BIT_ADMIN FROM INV_USERS ORDER BY VC_USER_ID, VC_PASSWORD');
      ExistingUserIDs_ComboBox.ItemIndex := 0;
      ExistingUserIDs_ComboBox.SetFocus;
    End;
  End;
end;        //Delete_ButtonClick

end.
