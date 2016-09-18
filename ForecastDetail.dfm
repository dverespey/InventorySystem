object ForecastDetail_Form: TForecastDetail_Form
  Left = 521
  Top = 174
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'Assembly Detail Master'
  ClientHeight = 556
  ClientWidth = 631
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object AssyRatioMaster_Label: TLabel
    Left = 10
    Top = 10
    Width = 189
    Height = 20
    Caption = 'Assembly Detail Master'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Label10: TLabel
    Left = 306
    Top = 214
    Width = 110
    Height = 20
    AutoSize = False
    Caption = 'Label Part Number'
    Layout = tlCenter
  end
  object ManagementButtons_Panel: TPanel
    Left = 20
    Top = 494
    Width = 585
    Height = 50
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 0
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
      Left = 390
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
  object ForecastDetail_DBGrid: TDBGrid
    Left = 20
    Top = 304
    Width = 585
    Height = 183
    DataSource = ForecastDetail_DataSource
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgRowSelect, dgConfirmDelete, dgCancelOnExit]
    TabOrder = 1
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    OnKeyUp = ForecastDetail_DBGridKeyUp
    OnMouseUp = ForecastDetail_DBGridMouseUp
  end
  object ForecastDetail_Panel: TPanel
    Left = 21
    Top = 31
    Width = 585
    Height = 266
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 2
    object SizeCode_Label: TLabel
      Left = 10
      Top = 14
      Width = 110
      Height = 20
      AutoSize = False
      Caption = 'Assembly Part Number'
      Layout = tlCenter
    end
    object ASSYCode_Label: TLabel
      Left = 10
      Top = 70
      Width = 110
      Height = 20
      AutoSize = False
      Caption = 'Tire Part Number'
      Layout = tlCenter
    end
    object DailyUsage_Label: TLabel
      Left = 10
      Top = 98
      Width = 110
      Height = 20
      AutoSize = False
      Caption = 'Wheel Part Number'
      Layout = tlCenter
    end
    object Label1: TLabel
      Left = 285
      Top = 14
      Width = 129
      Height = 20
      AutoSize = False
      Caption = 'Assembly Forecast Ratio'
      Layout = tlCenter
    end
    object Label2: TLabel
      Left = 287
      Top = 42
      Width = 110
      Height = 20
      AutoSize = False
      Caption = 'Forecast Kanban'
      Layout = tlCenter
    end
    object Label3: TLabel
      Left = 285
      Top = 69
      Width = 129
      Height = 20
      AutoSize = False
      Caption = 'Tire Forecast Ratio'
      Layout = tlCenter
    end
    object Label4: TLabel
      Left = 285
      Top = 99
      Width = 129
      Height = 20
      AutoSize = False
      Caption = 'Wheel Forecast Ratio'
      Layout = tlCenter
    end
    object Label5: TLabel
      Left = 10
      Top = 42
      Width = 110
      Height = 20
      AutoSize = False
      Caption = 'Effective Month'
      Layout = tlCenter
    end
    object TireQty_Label: TLabel
      Left = 10
      Top = 187
      Width = 95
      Height = 20
      AutoSize = False
      Caption = 'Assy Quantity:'
      Layout = tlCenter
    end
    object BroadcastCode_Label: TLabel
      Left = 10
      Top = 233
      Width = 95
      Height = 20
      AutoSize = False
      Caption = 'Broadcast Code:'
      Layout = tlCenter
    end
    object Label6: TLabel
      Left = 10
      Top = 126
      Width = 110
      Height = 20
      AutoSize = False
      Caption = 'Valve Part Number'
      Layout = tlCenter
    end
    object Label7: TLabel
      Left = 10
      Top = 155
      Width = 110
      Height = 20
      AutoSize = False
      Caption = 'Film Part Number'
      Layout = tlCenter
    end
    object Label8: TLabel
      Left = 285
      Top = 126
      Width = 110
      Height = 20
      AutoSize = False
      Caption = 'Label Part Number'
      Layout = tlCenter
    end
    object Label9: TLabel
      Left = 285
      Top = 154
      Width = 110
      Height = 20
      AutoSize = False
      Caption = 'Misc1 Part Number'
      Layout = tlCenter
    end
    object Label11: TLabel
      Left = 285
      Top = 184
      Width = 110
      Height = 20
      AutoSize = False
      Caption = 'Misc2 Part Number'
      Layout = tlCenter
    end
    object TirePartNum_ComboBox: TComboBox
      Left = 133
      Top = 69
      Width = 135
      Height = 21
      Style = csDropDownList
      CharCase = ecUpperCase
      ItemHeight = 13
      TabOrder = 3
    end
    object WheelPartNum_ComboBox: TComboBox
      Left = 133
      Top = 97
      Width = 135
      Height = 21
      Style = csDropDownList
      CharCase = ecUpperCase
      ItemHeight = 13
      TabOrder = 4
    end
    object ForecastRatio_Edit: TEdit
      Left = 419
      Top = 13
      Width = 132
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 12
      TabOrder = 0
      OnChange = ForecastPartsCode_EditChange
    end
    object Kanban_Edit: TEdit
      Left = 419
      Top = 41
      Width = 132
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 12
      TabOrder = 2
      OnChange = ForecastPartsCode_EditChange
    end
    object TireForecastRatio_Edit: TEdit
      Left = 419
      Top = 69
      Width = 132
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 12
      TabOrder = 5
      OnChange = ForecastPartsCode_EditChange
    end
    object WheelForecastRatio_Edit: TEdit
      Left = 419
      Top = 97
      Width = 132
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 12
      TabOrder = 6
      OnChange = ForecastPartsCode_EditChange
    end
    object EffectiveMonth_ComboBox: TComboBox
      Left = 133
      Top = 41
      Width = 135
      Height = 21
      Style = csDropDownList
      CharCase = ecUpperCase
      ItemHeight = 13
      TabOrder = 1
    end
    object AssyQty_RadioGroup: TRadioGroup
      Left = 133
      Top = 179
      Width = 136
      Height = 50
      Columns = 3
      ItemIndex = 0
      Items.Strings = (
        '1'
        '2'
        '4'
        '5')
      TabOrder = 7
    end
    object BroadcastCode_Edit: TEdit
      Left = 133
      Top = 233
      Width = 136
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 20
      TabOrder = 8
    end
    object ForecastPartsCode_ComboBox: TComboBox
      Left = 133
      Top = 13
      Width = 135
      Height = 21
      CharCase = ecUpperCase
      ItemHeight = 13
      TabOrder = 9
      OnSelect = ForecastPartsCode_ComboBoxSelect
    end
    object ValvePartNum_ComboBox: TComboBox
      Left = 133
      Top = 125
      Width = 135
      Height = 21
      Style = csDropDownList
      CharCase = ecUpperCase
      ItemHeight = 13
      TabOrder = 10
    end
    object FilmPArtNum_ComboBox: TComboBox
      Left = 133
      Top = 154
      Width = 135
      Height = 21
      Style = csDropDownList
      CharCase = ecUpperCase
      ItemHeight = 13
      TabOrder = 11
    end
    object LabelPartNum_ComboBox: TComboBox
      Left = 419
      Top = 125
      Width = 135
      Height = 21
      Style = csDropDownList
      CharCase = ecUpperCase
      ItemHeight = 13
      TabOrder = 12
    end
    object Misc1PartNum_ComboBox: TComboBox
      Left = 419
      Top = 154
      Width = 135
      Height = 21
      Style = csDropDownList
      CharCase = ecUpperCase
      ItemHeight = 13
      TabOrder = 13
    end
  end
  object Misc2PartNum_ComboBox: TComboBox
    Left = 440
    Top = 214
    Width = 135
    Height = 21
    Style = csDropDownList
    CharCase = ecUpperCase
    ItemHeight = 13
    TabOrder = 3
  end
  object ForecastDetail_DataSource: TDataSource
    OnDataChange = ForecastDetail_DataSourceDataChange
    Left = 476
    Top = 328
  end
end
