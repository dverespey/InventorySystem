object Shipping_Form: TShipping_Form
  Left = 716
  Top = 138
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsDialog
  Caption = 'Shipping'
  ClientHeight = 487
  ClientWidth = 308
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object CarTruckShip_Label: TLabel
    Left = 20
    Top = 10
    Width = 110
    Height = 20
    Caption = 'Line Shipping'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Check_Button: TButton
    Left = 117
    Top = 303
    Width = 81
    Height = 30
    Caption = 'C&heck'
    TabOrder = 1
    OnClick = Check_ButtonClick
  end
  object Authentication_GroupBox: TGroupBox
    Left = 40
    Top = 341
    Width = 230
    Height = 75
    Caption = 'Authentication'
    TabOrder = 2
    object ShipQty_Label: TLabel
      Left = 20
      Top = 15
      Width = 87
      Height = 21
      AutoSize = False
      Caption = 'Shipping Quantity:'
      Layout = tlCenter
    end
    object ContinuationNo_Label: TLabel
      Left = 20
      Top = 45
      Width = 87
      Height = 21
      AutoSize = False
      Caption = 'Continuation No.:'
      Layout = tlCenter
    end
    object ShipQty_MaskEdit: TMaskEdit
      Left = 110
      Top = 15
      Width = 100
      Height = 21
      CharCase = ecUpperCase
      EditMask = '#######;1; '
      MaxLength = 7
      ReadOnly = True
      TabOrder = 0
      Text = '0      '
      OnChange = TextChange
      OnExit = MaskEditExit
    end
    object ContNo_MaskEdit: TMaskEdit
      Left = 110
      Top = 45
      Width = 100
      Height = 21
      CharCase = ecUpperCase
      EditMask = '#######;1; '
      MaxLength = 7
      ReadOnly = True
      TabOrder = 1
      Text = '0      '
      OnChange = TextChange
      OnExit = MaskEditExit
    end
  end
  object Button_Panel: TPanel
    Left = 40
    Top = 426
    Width = 230
    Height = 50
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 3
    object Insert_Button: TButton
      Left = 8
      Top = 10
      Width = 98
      Height = 30
      Caption = '&Update Inventory'
      TabOrder = 0
      OnClick = Insert_ButtonClick
    end
    object Close_Button: TButton
      Left = 125
      Top = 10
      Width = 98
      Height = 30
      Caption = '&Close'
      ModalResult = 2
      TabOrder = 1
    end
  end
  object Shipping_Panel: TPanel
    Left = 17
    Top = 32
    Width = 281
    Height = 250
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 0
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
      Top = 41
      Width = 75
      Height = 21
      AutoSize = False
      Caption = 'Line:'
      Layout = tlCenter
    end
    object Date_Label: TLabel
      Left = 8
      Top = 13
      Width = 90
      Height = 21
      AutoSize = False
      Caption = 'Production Date:'
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
      Top = 69
      Width = 100
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 3
      TabOrder = 0
      OnChange = StartSeqNo_EditChange
    end
    object LastSeqNo_Edit: TEdit
      Left = 110
      Top = 127
      Width = 100
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 3
      TabOrder = 1
      OnChange = StartSeqNo_EditChange
    end
    object Production_DateTimePicker: TDateTimePicker
      Left = 110
      Top = 12
      Width = 100
      Height = 21
      Date = 37578.415469074100000000
      Time = 37578.415469074100000000
      TabOrder = 2
      OnChange = Production_DateTimePickerChange
    end
    object Line_ComboBox: TComboBox
      Left = 110
      Top = 40
      Width = 145
      Height = 21
      ItemHeight = 13
      TabOrder = 3
      OnChange = Line_ComboBoxChange
    end
    object StartBox: TNUMMIComboBox
      Left = 110
      Top = 98
      Width = 165
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 4
      OnChange = StartBoxChange
      ReadOnly = False
    end
    object EndBox: TNUMMIComboBox
      Left = 110
      Top = 157
      Width = 165
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 5
      ReadOnly = False
    end
  end
  object UpdateShipping_Button: TButton
    Left = 111
    Top = 303
    Width = 91
    Height = 30
    Caption = 'Update Shipping'
    TabOrder = 4
    Visible = False
    OnClick = UpdateShipping_ButtonClick
  end
  object SeqNum: TSequenceNumber
    SequenceValue = 1
    SequenceNumber = '0001'
    RolloverWarning = False
    WarningBeforeOffset = 0
    WarningAfterOffset = 0
    Left = 32
    Top = 215
  end
  object NT: TNummiTime
    Time = 37719.706485810200000000
    NummiTime = '2003040816572000'
    TimeValue = 37719.706485810200000000
    Left = 64
    Top = 215
  end
end
