object NewPasswordDlg: TNewPasswordDlg
  Left = 245
  Top = 108
  BorderStyle = bsDialog
  Caption = 'Password Dialog'
  ClientHeight = 171
  ClientWidth = 233
  Color = clBtnFace
  ParentFont = True
  OldCreateOrder = True
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 8
    Top = 9
    Width = 99
    Height = 13
    Caption = 'Enter new password:'
  end
  object Label2: TLabel
    Left = 8
    Top = 60
    Width = 109
    Height = 13
    Caption = 'Confirm new password:'
  end
  object NewPassword: TEdit
    Left = 8
    Top = 27
    Width = 217
    Height = 21
    CharCase = ecUpperCase
    MaxLength = 30
    PasswordChar = '*'
    TabOrder = 0
  end
  object OKBtn: TButton
    Left = 40
    Top = 138
    Width = 75
    Height = 25
    Caption = 'OK'
    Default = True
    TabOrder = 2
    OnClick = OKBtnClick
  end
  object CancelBtn: TButton
    Left = 120
    Top = 138
    Width = 75
    Height = 25
    Cancel = True
    Caption = 'Cancel'
    ModalResult = 2
    TabOrder = 3
  end
  object ConfirmPAssword: TEdit
    Left = 8
    Top = 78
    Width = 217
    Height = 21
    CharCase = ecUpperCase
    MaxLength = 30
    PasswordChar = '*'
    TabOrder = 1
  end
end
