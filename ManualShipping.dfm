object ManualShipping_Form: TManualShipping_Form
  Left = 404
  Top = 307
  Width = 852
  Height = 428
  Caption = 'Enter Daily Build'
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Arial'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCloseQuery = FormCloseQuery
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 14
  object Label1: TLabel
    Left = 295
    Top = 242
    Width = 62
    Height = 14
    Caption = 'Part Number:'
  end
  object Label2: TLabel
    Left = 514
    Top = 242
    Width = 53
    Height = 14
    Caption = 'Part Count:'
  end
  object IrregularShipCount_Label: TLabel
    Left = 514
    Top = 273
    Width = 53
    Height = 14
    Caption = 'Part Count:'
  end
  object Parts_StringGrid: TStringGrid
    Left = 0
    Top = 2
    Width = 843
    Height = 231
    ColCount = 4
    FixedCols = 0
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
    ScrollBars = ssVertical
    TabOrder = 0
    OnSelectCell = Parts_StringGridSelectCell
    ColWidths = (
      132
      134
      430
      103)
  end
  object PartNumber_Edit: TEdit
    Left = 388
    Top = 239
    Width = 121
    Height = 22
    ReadOnly = True
    TabOrder = 1
  end
  object Count_Edit: TEdit
    Left = 572
    Top = 239
    Width = 121
    Height = 22
    TabOrder = 2
  end
  object Update_Button: TButton
    Left = 706
    Top = 236
    Width = 45
    Height = 25
    Caption = 'Update'
    Default = True
    TabOrder = 3
    OnClick = Update_ButtonClick
  end
  object Close_Button: TButton
    Left = 754
    Top = 362
    Width = 81
    Height = 25
    Caption = 'Close'
    TabOrder = 4
    OnClick = Close_ButtonClick
  end
  object Post_Button: TButton
    Left = 764
    Top = 236
    Width = 75
    Height = 25
    Caption = 'Post Entries'
    TabOrder = 5
    OnClick = Post_ButtonClick
  end
  object Statistics_GroupBox: TGroupBox
    Left = 4
    Top = 238
    Width = 267
    Height = 152
    Caption = 'Daily Statistics'
    TabOrder = 6
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
      Left = 7
      Top = 67
      Width = 99
      Height = 21
      AutoSize = False
      Caption = 'Start Sequence No.:'
      Layout = tlCenter
    end
    object Label4: TLabel
      Left = 7
      Top = 95
      Width = 97
      Height = 21
      AutoSize = False
      Caption = 'Last Sequence No.:'
      Layout = tlCenter
    end
    object Label3: TLabel
      Left = 7
      Top = 19
      Width = 23
      Height = 14
      Caption = 'Line:'
    end
    object Label5: TLabel
      Left = 7
      Top = 124
      Width = 49
      Height = 14
      Caption = 'Daily Total'
    end
    object Production_DateTimePicker: TDateTimePicker
      Left = 114
      Top = 40
      Width = 100
      Height = 21
      Date = 37578.415469074100000000
      Time = 37578.415469074100000000
      TabOrder = 0
      OnChange = Production_DateTimePickerChange
    end
    object StartSeqNo_Edit: TEdit
      Left = 114
      Top = 67
      Width = 100
      Height = 22
      CharCase = ecUpperCase
      MaxLength = 3
      TabOrder = 1
    end
    object LastSeqNo_Edit: TEdit
      Left = 114
      Top = 94
      Width = 100
      Height = 22
      CharCase = ecUpperCase
      MaxLength = 3
      TabOrder = 2
    end
    object DailyTotal_Edit: TEdit
      Left = 114
      Top = 121
      Width = 98
      Height = 22
      ReadOnly = True
      TabOrder = 3
    end
    object Line_ComboBox: TComboBox
      Left = 114
      Top = 13
      Width = 145
      Height = 22
      ItemHeight = 14
      TabOrder = 4
      Text = 'Line_ComboBox'
    end
  end
  object IrregularShipCount_Edit: TEdit
    Left = 572
    Top = 270
    Width = 121
    Height = 22
    TabOrder = 7
  end
  object IrregularShip_Button: TButton
    Left = 706
    Top = 267
    Width = 133
    Height = 25
    Caption = 'Irregular Ship'
    Default = True
    TabOrder = 8
    OnClick = IrregularShip_ButtonClick
  end
end
