object AssyRatioMaster_Form: TAssyRatioMaster_Form
  Left = 829
  Top = 166
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsDialog
  Caption = 'ASSY / Ratio Master'
  ClientHeight = 613
  ClientWidth = 624
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Arial'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 14
  object AssyRatioMaster_Label: TLabel
    Left = 10
    Top = 10
    Width = 167
    Height = 20
    Caption = 'ASSY / Ratio Master'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Buttons_Panel: TPanel
    Left = 20
    Top = 556
    Width = 585
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
    object Search_Button: TButton
      Left = 295
      Top = 10
      Width = 90
      Height = 30
      Caption = '&Search'
      TabOrder = 3
      OnClick = Search_ButtonClick
    end
    object Clear_Button: TButton
      Left = 392
      Top = 10
      Width = 90
      Height = 30
      Caption = 'Cl&ear'
      TabOrder = 4
      OnClick = Clear_ButtonClick
    end
    object Close_Button: TButton
      Left = 485
      Top = 10
      Width = 90
      Height = 30
      Caption = '&Close'
      ModalResult = 2
      TabOrder = 5
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
  object AssyRatioMaster_Panel: TPanel
    Left = 20
    Top = 34
    Width = 585
    Height = 399
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 0
    object AssyCode_Label: TLabel
      Left = 12
      Top = 61
      Width = 95
      Height = 20
      AutoSize = False
      Caption = 'ASSY Code:'
      Layout = tlCenter
    end
    object TireQty_Label: TLabel
      Left = 10
      Top = 90
      Width = 95
      Height = 20
      AutoSize = False
      Caption = 'Tire Quantity:'
      Layout = tlCenter
    end
    object TirePartsCode1_Label: TLabel
      Left = 10
      Top = 115
      Width = 95
      Height = 20
      AutoSize = False
      Caption = 'Tire Parts Code 1:'
      Layout = tlCenter
    end
    object TirePartsCode2_Label: TLabel
      Left = 10
      Top = 140
      Width = 95
      Height = 20
      AutoSize = False
      Caption = 'Tire Parts Code 2:'
      Layout = tlCenter
    end
    object WheelPartsCode1_Label: TLabel
      Left = 10
      Top = 247
      Width = 95
      Height = 20
      AutoSize = False
      Caption = 'Wheel Parts Code 1:'
      Layout = tlCenter
    end
    object WheelPartsCode2_Label: TLabel
      Left = 10
      Top = 272
      Width = 95
      Height = 20
      AutoSize = False
      Caption = 'Wheel Parts Code 2:'
      Layout = tlCenter
    end
    object WheelQty_Label: TLabel
      Left = 10
      Top = 223
      Width = 95
      Height = 20
      AutoSize = False
      Caption = 'Wheel Quantity:'
      Layout = tlCenter
    end
    object SpareTirePartsCode_Label: TLabel
      Left = 10
      Top = 369
      Width = 115
      Height = 20
      AutoSize = False
      Caption = 'Spare Tire Parts Code:'
      Layout = tlCenter
      Visible = False
    end
    object AssyName_Label: TLabel
      Left = 12
      Top = 34
      Width = 95
      Height = 20
      AutoSize = False
      Caption = 'ASSY Name:'
      Layout = tlCenter
    end
    object TireRatio1_Label: TLabel
      Left = 280
      Top = 115
      Width = 95
      Height = 20
      AutoSize = False
      Caption = 'Tire Ratio 1:'
      Layout = tlCenter
    end
    object TireRatio2_Label: TLabel
      Left = 280
      Top = 140
      Width = 95
      Height = 20
      AutoSize = False
      Caption = 'Tire Ratio 2:'
      Layout = tlCenter
    end
    object WheelRatio1_Label: TLabel
      Left = 280
      Top = 247
      Width = 95
      Height = 20
      AutoSize = False
      Caption = 'Wheel Ratio 1:'
      Layout = tlCenter
    end
    object TireRatio4_Label: TLabel
      Left = 280
      Top = 272
      Width = 95
      Height = 20
      AutoSize = False
      Caption = 'Wheel Ratio 2:'
      Layout = tlCenter
    end
    object SpareTireQty_Label: TLabel
      Left = 10
      Top = 343
      Width = 95
      Height = 20
      AutoSize = False
      Caption = 'Spare Tire Quantity:'
      Layout = tlCenter
      Visible = False
    end
    object SpareWheelPartsCode_Label: TLabel
      Left = 280
      Top = 369
      Width = 122
      Height = 20
      AutoSize = False
      Caption = 'Spare Wheel Parts Code:'
      Layout = tlCenter
      Visible = False
    end
    object BroadcastCode_Label: TLabel
      Left = 10
      Top = 9
      Width = 95
      Height = 20
      AutoSize = False
      Caption = 'Broadcast Code:'
      Layout = tlCenter
    end
    object Label1: TLabel
      Left = 10
      Top = 297
      Width = 95
      Height = 20
      AutoSize = False
      Caption = 'Wheel Parts Code 2:'
      Layout = tlCenter
    end
    object Label2: TLabel
      Left = 280
      Top = 297
      Width = 95
      Height = 20
      AutoSize = False
      Caption = 'Wheel Ratio 3:'
      Layout = tlCenter
    end
    object Label3: TLabel
      Left = 10
      Top = 165
      Width = 95
      Height = 20
      AutoSize = False
      Caption = 'Tire Parts Code 3:'
      Layout = tlCenter
    end
    object Label4: TLabel
      Left = 280
      Top = 165
      Width = 95
      Height = 20
      AutoSize = False
      Caption = 'Tire Ratio 3:'
      Layout = tlCenter
    end
    object Label5: TLabel
      Left = 280
      Top = 190
      Width = 95
      Height = 20
      AutoSize = False
      Caption = 'Total Tire Ratio :'
      Layout = tlCenter
    end
    object Label6: TLabel
      Left = 280
      Top = 321
      Width = 95
      Height = 20
      AutoSize = False
      Caption = 'Total Wheel Ratio :'
      Layout = tlCenter
    end
    object SpareTireQty_RadioGroup: TRadioGroup
      Left = 133
      Top = 336
      Width = 135
      Height = 30
      Columns = 2
      ItemIndex = 0
      Items.Strings = (
        '0'
        '1')
      TabOrder = 17
      Visible = False
    end
    object SpareTirePartsCode_Edit: TEdit
      Left = 132
      Top = 369
      Width = 135
      Height = 22
      CharCase = ecUpperCase
      MaxLength = 12
      TabOrder = 18
      Visible = False
    end
    object WheelQty_RadioGroup: TRadioGroup
      Left = 133
      Top = 215
      Width = 135
      Height = 30
      Columns = 3
      ItemIndex = 0
      Items.Strings = (
        '1'
        '4'
        '5')
      TabOrder = 10
    end
    object TireQty_RadioGroup: TRadioGroup
      Left = 133
      Top = 82
      Width = 135
      Height = 30
      Columns = 3
      ItemIndex = 0
      Items.Strings = (
        '1'
        '4'
        '5')
      TabOrder = 3
    end
    object SpareWheelPartsCode_Edit: TEdit
      Left = 400
      Top = 368
      Width = 135
      Height = 22
      CharCase = ecUpperCase
      MaxLength = 12
      TabOrder = 19
      Visible = False
    end
    object AssyName_Edit: TEdit
      Left = 132
      Top = 34
      Width = 175
      Height = 22
      CharCase = ecUpperCase
      MaxLength = 50
      TabOrder = 1
    end
    object TireRatio1_MaskEdit: TMaskEdit
      Left = 400
      Top = 115
      Width = 126
      Height = 22
      CharCase = ecUpperCase
      EditMask = '999;1; '
      MaxLength = 3
      TabOrder = 5
      Text = '0  '
      OnChange = TireRatio1_MaskEditChange
    end
    object TireRatio2_MaskEdit: TMaskEdit
      Left = 400
      Top = 140
      Width = 126
      Height = 22
      CharCase = ecUpperCase
      EditMask = '999;1; '
      MaxLength = 3
      TabOrder = 7
      Text = '0  '
      OnChange = TireRatio2_MaskEditChange
    end
    object WheelRatio1_MaskEdit: TMaskEdit
      Left = 400
      Top = 247
      Width = 126
      Height = 22
      CharCase = ecUpperCase
      EditMask = '999;1; '
      MaxLength = 3
      TabOrder = 12
      Text = '0  '
      OnChange = WheelRatio1_MaskEditChange
    end
    object WheelRatio2_MaskEdit: TMaskEdit
      Left = 400
      Top = 272
      Width = 126
      Height = 22
      CharCase = ecUpperCase
      EditMask = '999;1; '
      MaxLength = 3
      TabOrder = 14
      Text = '0  '
      OnChange = WheelRatio2_MaskEditChange
    end
    object BroadcastCode_Edit: TEdit
      Left = 133
      Top = 9
      Width = 35
      Height = 22
      CharCase = ecUpperCase
      MaxLength = 2
      TabOrder = 0
    end
    object TirePartNum1_ComboBox: TComboBox
      Left = 133
      Top = 115
      Width = 135
      Height = 22
      Style = csDropDownList
      CharCase = ecUpperCase
      ItemHeight = 14
      TabOrder = 4
      OnChange = TirePartNum1_ComboBoxChange
    end
    object TirePartNum2_ComboBox: TComboBox
      Left = 133
      Top = 140
      Width = 135
      Height = 22
      Style = csDropDownList
      CharCase = ecUpperCase
      ItemHeight = 14
      TabOrder = 6
      OnChange = TirePartNum2_ComboBoxChange
    end
    object WheelPartNum1_ComboBox: TComboBox
      Left = 133
      Top = 247
      Width = 135
      Height = 22
      Style = csDropDownList
      CharCase = ecUpperCase
      ItemHeight = 14
      TabOrder = 11
      OnChange = WheelPartNum1_ComboBoxChange
    end
    object WheelPartNum2_ComboBox: TComboBox
      Left = 133
      Top = 272
      Width = 135
      Height = 22
      Style = csDropDownList
      CharCase = ecUpperCase
      ItemHeight = 14
      TabOrder = 13
      OnChange = WheelPartNum2_ComboBoxChange
    end
    object AssyCode_ComboBox: TComboBox
      Left = 132
      Top = 60
      Width = 136
      Height = 22
      Style = csDropDownList
      CharCase = ecUpperCase
      ItemHeight = 14
      TabOrder = 2
    end
    object WheelPartNum3_ComboBox: TComboBox
      Left = 133
      Top = 297
      Width = 135
      Height = 22
      Style = csDropDownList
      CharCase = ecUpperCase
      ItemHeight = 14
      TabOrder = 15
      OnChange = WheelPartNum3_ComboBoxChange
    end
    object WheelRatio3_MaskEdit: TMaskEdit
      Left = 400
      Top = 297
      Width = 126
      Height = 22
      CharCase = ecUpperCase
      EditMask = '999;1; '
      MaxLength = 3
      TabOrder = 16
      Text = '0  '
      OnChange = WheelRatio3_MaskEditChange
    end
    object TirePartNum3_ComboBox: TComboBox
      Left = 133
      Top = 165
      Width = 135
      Height = 22
      Style = csDropDownList
      CharCase = ecUpperCase
      ItemHeight = 14
      TabOrder = 8
      OnChange = TirePartNum3_ComboBoxChange
    end
    object TireRatio3_MaskEdit: TMaskEdit
      Left = 400
      Top = 165
      Width = 126
      Height = 22
      CharCase = ecUpperCase
      EditMask = '999;1; '
      MaxLength = 3
      TabOrder = 9
      Text = '0  '
      OnChange = TireRatio3_MaskEditChange
    end
    object TotalTireRatio_Edit: TEdit
      Tag = 1
      Left = 400
      Top = 189
      Width = 57
      Height = 22
      ReadOnly = True
      TabOrder = 20
      Text = '0'
      OnChange = TotalTireRatio_EditChange
    end
    object TotalWheelRatio_Edit: TEdit
      Tag = 1
      Left = 400
      Top = 320
      Width = 57
      Height = 22
      ReadOnly = True
      TabOrder = 21
      Text = '0'
      OnChange = TotalTireRatio_EditChange
    end
  end
  object ASSYRatioMaster_DBGrid: TDBGrid
    Left = 20
    Top = 436
    Width = 585
    Height = 117
    DataSource = AssyRatio_DataSource
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgRowSelect, dgConfirmDelete, dgCancelOnExit]
    TabOrder = 2
    TitleFont.Charset = ANSI_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Arial'
    TitleFont.Style = []
    OnKeyUp = ASSYRatioMaster_DBGridKeyUp
    OnMouseUp = ASSYRatioMaster_DBGridMouseUp
  end
  object AssyRatio_DataSource: TDataSource
    OnDataChange = AssyRatio_DataSourceDataChange
    Left = 586
    Top = 6
  end
end
