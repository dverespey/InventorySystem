object UserAdmin_Form: TUserAdmin_Form
  Left = 612
  Top = 113
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsDialog
  Caption = 'Inventory System User Administration'
  ClientHeight = 304
  ClientWidth = 513
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object UserAdmin_Panel: TPanel
    Left = 60
    Top = 16
    Width = 385
    Height = 210
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 0
    object ConfirmPassword_Label: TLabel
      Left = 50
      Top = 130
      Width = 111
      Height = 20
      AutoSize = False
      Caption = 'Confirm Password:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      Layout = tlCenter
    end
    object Password_Label: TLabel
      Left = 50
      Top = 95
      Width = 89
      Height = 20
      AutoSize = False
      Caption = 'Password:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      Layout = tlCenter
    end
    object Label1: TLabel
      Left = 50
      Top = 60
      Width = 89
      Height = 20
      AutoSize = False
      Caption = 'User ID:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      Layout = tlCenter
    end
    object ExistingUsers_Label: TLabel
      Left = 50
      Top = 25
      Width = 109
      Height = 20
      AutoSize = False
      Caption = 'Existing User IDs:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      Layout = tlCenter
    end
    object IsAdmin_CheckBox: TCheckBox
      Left = 50
      Top = 165
      Width = 143
      Height = 17
      Alignment = taLeftJustify
      BiDiMode = bdLeftToRight
      Caption = 'Is Administrator:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentBiDiMode = False
      ParentFont = False
      TabOrder = 4
    end
    object ExistingUserIDs_ComboBox: TComboBox
      Left = 180
      Top = 25
      Width = 150
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 0
      OnChange = ExistingUserIDs_ComboBoxChange
      OnKeyPress = EditKeyPress
    end
    object UserID_Edit: TEdit
      Left = 180
      Top = 60
      Width = 150
      Height = 22
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = [fsBold]
      MaxLength = 30
      ParentFont = False
      TabOrder = 1
      OnKeyPress = EditKeyPress
    end
    object Password_Edit: TEdit
      Left = 180
      Top = 95
      Width = 150
      Height = 22
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = [fsBold]
      MaxLength = 30
      ParentFont = False
      PasswordChar = '*'
      TabOrder = 2
      OnKeyPress = EditKeyPress
    end
    object ConfirmPassword_Edit: TEdit
      Left = 180
      Top = 130
      Width = 150
      Height = 22
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = [fsBold]
      MaxLength = 30
      ParentFont = False
      PasswordChar = '*'
      TabOrder = 3
      OnKeyPress = EditKeyPress
    end
  end
  object Buttons_Panel: TPanel
    Left = 10
    Top = 242
    Width = 492
    Height = 50
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 1
    object Insert_Button: TButton
      Left = 10
      Top = 10
      Width = 90
      Height = 30
      Caption = '&Insert'
      TabOrder = 0
      OnClick = Insert_ButtonClick
    end
    object Update_Button: TButton
      Left = 105
      Top = 10
      Width = 90
      Height = 30
      Caption = '&Update'
      TabOrder = 1
      OnClick = Update_ButtonClick
    end
    object Clear_Button: TButton
      Left = 296
      Top = 10
      Width = 90
      Height = 30
      Caption = 'Cl&ear'
      TabOrder = 3
      OnClick = Clear_ButtonClick
    end
    object Close_Button: TButton
      Left = 391
      Top = 10
      Width = 90
      Height = 30
      Caption = '&Close'
      ModalResult = 2
      TabOrder = 4
    end
    object Delete_Button: TButton
      Left = 200
      Top = 10
      Width = 90
      Height = 30
      Caption = '&Delete'
      TabOrder = 2
      OnClick = Delete_ButtonClick
    end
  end
end
