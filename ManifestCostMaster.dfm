object ManifestCostMaster_Form: TManifestCostMaster_Form
  Left = 888
  Top = 216
  BorderStyle = bsSingle
  Caption = 'Manifest Cost Master'
  ClientHeight = 385
  ClientWidth = 629
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Arial'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 14
  object PartsStockMaster_Label: TLabel
    Left = 20
    Top = 10
    Width = 172
    Height = 20
    Caption = 'Manifest Cost Master'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object MonthlyPOMaster_DBGrid: TDBGrid
    Left = 20
    Top = 160
    Width = 585
    Height = 158
    DataSource = MAnifestCost_DataSource
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgRowSelect, dgConfirmDelete, dgCancelOnExit]
    TabOrder = 0
    TitleFont.Charset = ANSI_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Arial'
    TitleFont.Style = []
  end
  object ManagementButtons_Panel: TPanel
    Left = 20
    Top = 324
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
  object ManifestCost_Panel: TPanel
    Left = 21
    Top = 39
    Width = 585
    Height = 116
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 2
    object BasisUnitPrice_Label: TLabel
      Left = 10
      Top = 60
      Width = 82
      Height = 20
      AutoSize = False
      Caption = 'Assembly Cost'
      Layout = tlCenter
    end
    object AssyCode_Label: TLabel
      Left = 10
      Top = 5
      Width = 95
      Height = 20
      AutoSize = False
      Caption = 'ASSY Code:'
      Layout = tlCenter
    end
    object InTransit_Label: TLabel
      Left = 10
      Top = 31
      Width = 88
      Height = 21
      AutoSize = False
      Caption = 'Cost Start Date'
      Layout = tlCenter
    end
    object Arrival_Label: TLabel
      Left = 275
      Top = 32
      Width = 83
      Height = 21
      AutoSize = False
      Caption = 'Cost End Date'
      Layout = tlCenter
    end
    object SizeCode_Label: TLabel
      Left = 8
      Top = 84
      Width = 96
      Height = 20
      AutoSize = False
      Caption = 'Assy Manifest ID'
      Layout = tlCenter
    end
    object AssyCode_ComboBox: TComboBox
      Left = 117
      Top = 4
      Width = 136
      Height = 22
      Style = csDropDownList
      CharCase = ecUpperCase
      ItemHeight = 14
      TabOrder = 0
    end
    object CostStart_NUMMIBmDateEdit: TNUMMIBmDateEdit
      Left = 117
      Top = 31
      Width = 121
      Height = 21
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      ColorCalendar.ColorValid = clBlue
      DayNames.Monday = 'Mo'
      DayNames.Tuesday = 'Tu'
      DayNames.Wednesday = 'We'
      DayNames.Thursday = 'Th'
      DayNames.Friday = 'Fr'
      DayNames.Saturday = 'Sa'
      DayNames.Sunday = 'Su'
      MonthNames.January = 'January'
      MonthNames.February = 'February'
      MonthNames.March = 'March'
      MonthNames.April = 'April'
      MonthNames.May = 'May'
      MonthNames.June = 'June'
      MonthNames.July = 'July'
      MonthNames.August = 'August'
      MonthNames.September = 'September'
      MonthNames.October = 'October'
      MonthNames.November = 'November'
      MonthNames.December = 'December'
      Options = [doButtonTabStop, doCanClear, doCanPopup, doIsMasked, doShowCancel, doShowToday]
      ReadOnly = False
    end
    object CostEnd_NUMMIBmDateEdit: TNUMMIBmDateEdit
      Left = 365
      Top = 32
      Width = 121
      Height = 21
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 2
      ColorCalendar.ColorValid = clBlue
      DayNames.Monday = 'Mo'
      DayNames.Tuesday = 'Tu'
      DayNames.Wednesday = 'We'
      DayNames.Thursday = 'Th'
      DayNames.Friday = 'Fr'
      DayNames.Saturday = 'Sa'
      DayNames.Sunday = 'Su'
      MonthNames.January = 'January'
      MonthNames.February = 'February'
      MonthNames.March = 'March'
      MonthNames.April = 'April'
      MonthNames.May = 'May'
      MonthNames.June = 'June'
      MonthNames.July = 'July'
      MonthNames.August = 'August'
      MonthNames.September = 'September'
      MonthNames.October = 'October'
      MonthNames.November = 'November'
      MonthNames.December = 'December'
      Options = [doButtonTabStop, doCanClear, doCanPopup, doIsMasked, doShowCancel, doShowToday]
      ReadOnly = False
    end
    object AssyCost_MaskEdit: TcurrEdit
      Left = 117
      Top = 57
      Width = 81
      Height = 22
      TabOrder = 3
      Text = '$0.0000'
      IntDigits = 0
      DecimalDigits = 4
      Prefix = '$'
    end
    object AssyManifestID_ComboBox: TComboBox
      Left = 117
      Top = 81
      Width = 55
      Height = 22
      Style = csDropDownList
      CharCase = ecUpperCase
      ItemHeight = 14
      TabOrder = 4
    end
  end
  object MAnifestCost_DataSource: TDataSource
    OnDataChange = MAnifestCost_DataSourceDataChange
    Left = 574
    Top = 44
  end
end
