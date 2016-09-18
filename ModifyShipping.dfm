object ModifyShipping_Form: TModifyShipping_Form
  Left = 663
  Top = 16
  BorderStyle = bsSingle
  Caption = 'Modify Shipping'
  ClientHeight = 458
  ClientWidth = 578
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Arial'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 14
  object Label1: TLabel
    Left = 25
    Top = 96
    Width = 59
    Height = 14
    Caption = 'Part Number'
  end
  object Label2: TLabel
    Left = 25
    Top = 125
    Width = 61
    Height = 14
    Caption = 'Shipping Qty'
  end
  object ManagementButtons_Panel: TPanel
    Left = 13
    Top = 402
    Width = 553
    Height = 50
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 3
    object Insert_Button: TButton
      Left = 5
      Top = 10
      Width = 90
      Height = 30
      Caption = '&Insert'
      TabOrder = 0
      Visible = False
      OnClick = Insert_ButtonClick
    end
    object Update_Button: TButton
      Left = 96
      Top = 10
      Width = 90
      Height = 30
      Caption = '&Update'
      TabOrder = 1
      OnClick = Update_ButtonClick
    end
    object Search_Button: TButton
      Left = 276
      Top = 10
      Width = 90
      Height = 30
      Caption = '&Search'
      TabOrder = 2
      Visible = False
    end
    object Clear_Button: TButton
      Left = 367
      Top = 10
      Width = 90
      Height = 30
      Caption = 'Cl&ear'
      TabOrder = 3
      Visible = False
      OnClick = Clear_ButtonClick
    end
    object Close_Button: TButton
      Left = 457
      Top = 12
      Width = 90
      Height = 28
      Caption = '&Close'
      ModalResult = 2
      TabOrder = 4
      OnClick = Close_ButtonClick
    end
    object Delete_Button: TButton
      Left = 186
      Top = 10
      Width = 90
      Height = 30
      Caption = '&Delete'
      TabOrder = 5
      Visible = False
    end
  end
  object GroupBox1: TGroupBox
    Left = 12
    Top = 6
    Width = 553
    Height = 71
    Caption = 'Production Date'
    Enabled = False
    TabOrder = 4
    object Label3: TLabel
      Left = 7
      Top = 19
      Width = 23
      Height = 14
      Caption = 'Line:'
    end
    object Date_Label: TLabel
      Left = 7
      Top = 41
      Width = 90
      Height = 21
      AutoSize = False
      Caption = 'Production Date:'
      Layout = tlCenter
    end
    object StartSeqNo_Label: TLabel
      Left = 271
      Top = 11
      Width = 99
      Height = 21
      AutoSize = False
      Caption = 'Start Sequence No.:'
      Layout = tlCenter
    end
    object Label4: TLabel
      Left = 271
      Top = 39
      Width = 97
      Height = 21
      AutoSize = False
      Caption = 'Last Sequence No.:'
      Layout = tlCenter
    end
    object Line_ComboBox: TComboBox
      Left = 114
      Top = 13
      Width = 145
      Height = 22
      ItemHeight = 14
      TabOrder = 0
      Text = 'Line_ComboBox'
    end
    object Production_DateTimePicker: TDateTimePicker
      Left = 114
      Top = 40
      Width = 100
      Height = 22
      Date = 37578.415469074100000000
      Time = 37578.415469074100000000
      TabOrder = 1
    end
    object StartSeqNo_Edit: TEdit
      Left = 378
      Top = 11
      Width = 100
      Height = 22
      CharCase = ecUpperCase
      MaxLength = 3
      TabOrder = 2
    end
    object LastSeqNo_Edit: TEdit
      Left = 378
      Top = 38
      Width = 100
      Height = 22
      CharCase = ecUpperCase
      MaxLength = 3
      TabOrder = 3
    end
  end
  object Qty_Edit: TEdit
    Left = 125
    Top = 121
    Width = 68
    Height = 22
    TabOrder = 1
    Text = 'Qty_Edit'
  end
  object ModifyShipping_DBGrid: TDBGrid
    Left = 13
    Top = 156
    Width = 551
    Height = 240
    DataSource = ModifyShipping_DataSource
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgRowSelect, dgConfirmDelete, dgCancelOnExit]
    TabOrder = 2
    TitleFont.Charset = ANSI_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Arial'
    TitleFont.Style = []
  end
  object PartNumber_ComboBox: TComboBox
    Left = 124
    Top = 93
    Width = 174
    Height = 22
    ItemHeight = 14
    TabOrder = 0
    Text = 'PartNumber_ComboBox'
  end
  object ModifyShipping_DataSource: TDataSource
    OnDataChange = ModifyShipping_DataSourceDataChange
    Left = 394
    Top = 120
  end
end
