object HotCallEntryForm: THotCallEntryForm
  Left = 970
  Top = 312
  BorderStyle = bsSingle
  Caption = 'One Cycle Entry'
  ClientHeight = 383
  ClientWidth = 806
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
  object Button_Panel: TPanel
    Left = 10
    Top = 318
    Width = 782
    Height = 50
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 2
    object Close_Button: TButton
      Left = 678
      Top = 9
      Width = 90
      Height = 30
      Caption = '&Close'
      ModalResult = 2
      TabOrder = 0
      OnClick = Close_ButtonClick
    end
    object Add_Button: TButton
      Left = 584
      Top = 10
      Width = 90
      Height = 30
      Caption = '&Add'
      TabOrder = 1
      OnClick = Add_ButtonClick
    end
    object Clear_Button: TButton
      Left = 489
      Top = 10
      Width = 90
      Height = 30
      Caption = '&Clear'
      TabOrder = 2
      OnClick = Clear_ButtonClick
    end
  end
  object ASNItems_GroupBox: TGroupBox
    Left = 10
    Top = 108
    Width = 782
    Height = 204
    Caption = 'ASN Items'
    TabOrder = 1
    object SizeCode_Label: TLabel
      Left = 38
      Top = 13
      Width = 110
      Height = 20
      AutoSize = False
      Caption = 'Assembly Part Number:'
      Layout = tlCenter
    end
    object Label2: TLabel
      Left = 231
      Top = 12
      Width = 45
      Height = 20
      AutoSize = False
      Caption = 'Qty:'
      Layout = tlCenter
    end
    object Label3: TLabel
      Left = 375
      Top = 13
      Width = 110
      Height = 20
      AutoSize = False
      Caption = 'Assembly Part Number:'
      Layout = tlCenter
    end
    object Label4: TLabel
      Left = 568
      Top = 12
      Width = 45
      Height = 20
      AutoSize = False
      Caption = 'Qty:'
      Layout = tlCenter
    end
    object AssyPartsCode1_ComboBox: TComboBox
      Left = 36
      Top = 37
      Width = 145
      Height = 22
      CharCase = ecUpperCase
      ItemHeight = 14
      TabOrder = 0
    end
    object Qty1_Edit: TEdit
      Left = 231
      Top = 38
      Width = 48
      Height = 22
      MaxLength = 3
      TabOrder = 1
      Text = 'Edit1'
    end
    object AssyPartsCode2_ComboBox: TComboBox
      Left = 36
      Top = 64
      Width = 145
      Height = 22
      CharCase = ecUpperCase
      ItemHeight = 14
      TabOrder = 2
    end
    object Qty2_Edit: TEdit
      Left = 231
      Top = 65
      Width = 48
      Height = 22
      MaxLength = 3
      TabOrder = 3
      Text = 'Qty2_Edit'
    end
    object AssyPartsCode3_ComboBox: TComboBox
      Left = 36
      Top = 90
      Width = 145
      Height = 22
      CharCase = ecUpperCase
      ItemHeight = 14
      TabOrder = 4
    end
    object Qty3_Edit: TEdit
      Left = 231
      Top = 91
      Width = 48
      Height = 22
      MaxLength = 3
      TabOrder = 5
      Text = 'Edit1'
    end
    object AssyPartsCode4_ComboBox: TComboBox
      Left = 36
      Top = 117
      Width = 145
      Height = 22
      CharCase = ecUpperCase
      ItemHeight = 14
      TabOrder = 6
    end
    object Qty4_Edit: TEdit
      Left = 231
      Top = 118
      Width = 48
      Height = 22
      MaxLength = 3
      TabOrder = 7
      Text = 'Edit1'
    end
    object AssyPartsCode5_ComboBox: TComboBox
      Left = 36
      Top = 143
      Width = 145
      Height = 22
      CharCase = ecUpperCase
      ItemHeight = 14
      TabOrder = 8
    end
    object Qty5_Edit: TEdit
      Left = 231
      Top = 144
      Width = 48
      Height = 22
      MaxLength = 3
      TabOrder = 9
      Text = 'Edit1'
    end
    object AssyPartsCode6_ComboBox: TComboBox
      Left = 36
      Top = 169
      Width = 145
      Height = 22
      CharCase = ecUpperCase
      ItemHeight = 14
      TabOrder = 10
    end
    object Qty6_Edit: TEdit
      Left = 231
      Top = 170
      Width = 48
      Height = 22
      MaxLength = 3
      TabOrder = 11
      Text = 'Edit1'
    end
    object AssyPartsCode7_ComboBox: TComboBox
      Left = 373
      Top = 37
      Width = 145
      Height = 22
      CharCase = ecUpperCase
      ItemHeight = 14
      TabOrder = 12
    end
    object Qty7_Edit: TEdit
      Left = 568
      Top = 38
      Width = 48
      Height = 22
      MaxLength = 3
      TabOrder = 13
      Text = 'Edit1'
    end
    object AssyPartsCode8_ComboBox: TComboBox
      Left = 373
      Top = 64
      Width = 145
      Height = 22
      CharCase = ecUpperCase
      ItemHeight = 14
      TabOrder = 14
    end
    object Qty8_Edit: TEdit
      Left = 568
      Top = 65
      Width = 48
      Height = 22
      MaxLength = 3
      TabOrder = 15
      Text = 'Edit1'
    end
    object AssyPartsCode9_ComboBox: TComboBox
      Left = 373
      Top = 90
      Width = 145
      Height = 22
      CharCase = ecUpperCase
      ItemHeight = 14
      TabOrder = 16
    end
    object Qty9_Edit: TEdit
      Left = 568
      Top = 91
      Width = 48
      Height = 22
      MaxLength = 3
      TabOrder = 17
      Text = 'Edit1'
    end
    object AssyPartsCode10_ComboBox: TComboBox
      Left = 373
      Top = 117
      Width = 145
      Height = 22
      CharCase = ecUpperCase
      ItemHeight = 14
      TabOrder = 18
    end
    object Qty10_Edit: TEdit
      Left = 568
      Top = 118
      Width = 48
      Height = 22
      MaxLength = 3
      TabOrder = 19
      Text = 'Edit1'
    end
    object AssyPartsCode11_ComboBox: TComboBox
      Left = 373
      Top = 143
      Width = 145
      Height = 22
      CharCase = ecUpperCase
      ItemHeight = 14
      TabOrder = 20
    end
    object Qty11_Edit: TEdit
      Left = 568
      Top = 144
      Width = 48
      Height = 22
      MaxLength = 3
      TabOrder = 21
      Text = 'Edit1'
    end
    object AssyPartsCode12_ComboBox: TComboBox
      Left = 373
      Top = 169
      Width = 145
      Height = 22
      CharCase = ecUpperCase
      ItemHeight = 14
      TabOrder = 22
    end
    object Qty12_Edit: TEdit
      Left = 568
      Top = 170
      Width = 48
      Height = 22
      MaxLength = 3
      TabOrder = 23
      Text = 'Edit1'
    end
  end
  object GroupBox2: TGroupBox
    Left = 10
    Top = 10
    Width = 782
    Height = 98
    Caption = 'ASN Details'
    TabOrder = 0
    object CarTruck_Label: TLabel
      Left = 35
      Top = 15
      Width = 75
      Height = 21
      AutoSize = False
      Caption = 'Line:'
      Layout = tlCenter
    end
    object Date_Label: TLabel
      Left = 35
      Top = 42
      Width = 90
      Height = 21
      AutoSize = False
      Caption = 'ASN Date:'
      Layout = tlCenter
    end
    object Label1: TLabel
      Left = 35
      Top = 70
      Width = 110
      Height = 20
      AutoSize = False
      Caption = 'Manifest Number:'
      Layout = tlCenter
    end
    object Line_ComboBox: TComboBox
      Left = 157
      Top = 13
      Width = 145
      Height = 22
      ItemHeight = 14
      TabOrder = 0
      Text = 'Line_ComboBox'
    end
    object ASN_DateTimePicker: TDateTimePicker
      Left = 157
      Top = 40
      Width = 100
      Height = 21
      Date = 37578.415469074100000000
      Time = 37578.415469074100000000
      TabOrder = 1
    end
    object ManifestNumber_Edit: TEdit
      Left = 157
      Top = 68
      Width = 121
      Height = 22
      MaxLength = 8
      TabOrder = 2
      Text = 'ManifestNumber_Edit'
    end
  end
end
