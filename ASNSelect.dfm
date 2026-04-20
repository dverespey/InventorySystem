object ASNSelect_Form: TASNSelect_Form
  Left = 707
  Top = 11
  BorderStyle = bsSingle
  Caption = 'ASN Create'
  ClientHeight = 392
  ClientWidth = 316
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
  object CarTruckShip_Label: TLabel
    Left = 20
    Top = 10
    Width = 96
    Height = 20
    Caption = 'ASN Create'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Check_Button: TButton
    Left = 110
    Top = 240
    Width = 81
    Height = 30
    Caption = 'C&heck'
    TabOrder = 0
    OnClick = Check_ButtonClick
  end
  object Authentication_GroupBox: TGroupBox
    Left = 40
    Top = 276
    Width = 230
    Height = 48
    Caption = 'Authentication'
    TabOrder = 1
    object ShipQty_Label: TLabel
      Left = 20
      Top = 15
      Width = 87
      Height = 21
      AutoSize = False
      Caption = 'Shipping Quantity:'
      Layout = tlCenter
    end
    object ShipQty_MaskEdit: TMaskEdit
      Left = 110
      Top = 15
      Width = 100
      Height = 22
      CharCase = ecUpperCase
      EditMask = '#######;1; '
      MaxLength = 7
      ReadOnly = True
      TabOrder = 0
      Text = '0      '
    end
  end
  object Button_Panel: TPanel
    Left = 17
    Top = 333
    Width = 281
    Height = 50
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 2
    object CreateASNEntries_Button: TButton
      Left = 8
      Top = 10
      Width = 102
      Height = 30
      Caption = '&Create ASN/Entries'
      TabOrder = 0
      OnClick = CreateASNEntries_ButtonClick
    end
    object Close_Button: TButton
      Left = 212
      Top = 10
      Width = 57
      Height = 30
      Caption = '&Close'
      ModalResult = 2
      TabOrder = 1
    end
    object CreateASN_Button: TButton
      Left = 114
      Top = 10
      Width = 94
      Height = 30
      Caption = '&Create ASN/Files'
      TabOrder = 2
      OnClick = CreateASN_ButtonClick
    end
  end
  object Shipping_Panel: TPanel
    Left = 17
    Top = 42
    Width = 281
    Height = 188
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 3
    object Label4: TLabel
      Left = 8
      Top = 126
      Width = 97
      Height = 21
      AutoSize = False
      Caption = 'Last Sequence No.:'
      Layout = tlCenter
    end
    object StartSeqNo_Label: TLabel
      Left = 8
      Top = 70
      Width = 99
      Height = 21
      AutoSize = False
      Caption = 'Start Sequence No.:'
      Layout = tlCenter
    end
    object CarTruck_Label: TLabel
      Left = 8
      Top = 12
      Width = 75
      Height = 21
      AutoSize = False
      Caption = 'Line:'
      Layout = tlCenter
    end
    object Date_Label: TLabel
      Left = 8
      Top = 44
      Width = 90
      Height = 21
      AutoSize = False
      Caption = 'ASN Date:'
      Layout = tlCenter
    end
    object Label3: TLabel
      Left = 8
      Top = 99
      Width = 62
      Height = 21
      AutoSize = False
      Caption = 'Start Time:'
      Layout = tlCenter
    end
    object Label5: TLabel
      Left = 8
      Top = 157
      Width = 62
      Height = 21
      AutoSize = False
      Caption = 'End Time:'
      Layout = tlCenter
    end
    object StartSeqNo_Edit: TEdit
      Left = 110
      Top = 70
      Width = 100
      Height = 22
      CharCase = ecUpperCase
      MaxLength = 4
      TabOrder = 0
      OnChange = StartSeqNo_EditChange
    end
    object LastSeqNo_Edit: TEdit
      Left = 110
      Top = 127
      Width = 100
      Height = 22
      CharCase = ecUpperCase
      MaxLength = 4
      TabOrder = 1
      OnChange = LastSeqNo_EditChange
    end
    object ASN_DateTimePicker: TDateTimePicker
      Left = 110
      Top = 43
      Width = 100
      Height = 21
      Date = 37578.415469074100000000
      Time = 37578.415469074100000000
      TabOrder = 2
      OnChange = ASN_DateTimePickerChange
    end
    object Line_ComboBox: TComboBox
      Left = 110
      Top = 13
      Width = 145
      Height = 22
      ItemHeight = 14
      TabOrder = 3
      Text = 'Line_ComboBox'
      OnChange = Line_ComboBoxChange
    end
    object StartBox: TNUMMIComboBox
      Left = 110
      Top = 97
      Width = 166
      Height = 22
      ItemHeight = 14
      TabOrder = 4
      OnChange = StartBoxChange
      ReadOnly = False
    end
    object EndBox: TNUMMIComboBox
      Left = 110
      Top = 157
      Width = 166
      Height = 22
      ItemHeight = 14
      TabOrder = 5
      ReadOnly = False
    end
  end
  object SeqNum: TSequenceNumber
    SequenceValue = 1
    SequenceNumber = '0001'
    RolloverWarning = False
    WarningBeforeOffset = 0
    WarningAfterOffset = 0
    Left = 278
    Top = 241
  end
  object NT: TNummiTime
    Time = 37719.706485810200000000
    NummiTime = '2003040816572000'
    TimeValue = 37719.706485810200000000
    Left = 280
    Top = 269
  end
end
