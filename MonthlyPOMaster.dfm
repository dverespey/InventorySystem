object MonthlyPOMaster_Form: TMonthlyPOMaster_Form
  Left = 717
  Top = 227
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'Monthly PO Master'
  ClientHeight = 385
  ClientWidth = 628
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
  object PartsStockMaster_Label: TLabel
    Left = 20
    Top = 10
    Width = 152
    Height = 20
    Caption = 'Monthly PO Master'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object MonthlyPO_Panel: TPanel
    Left = 21
    Top = 39
    Width = 585
    Height = 141
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 0
    object BasisUnitPrice_Label: TLabel
      Left = 275
      Top = 84
      Width = 82
      Height = 20
      AutoSize = False
      Caption = 'Assembly Cost'
      Layout = tlCenter
    end
    object AssyCode_Label: TLabel
      Left = 8
      Top = 58
      Width = 95
      Height = 20
      AutoSize = False
      Caption = 'ASSY Code:'
      Layout = tlCenter
    end
    object InTransit_Label: TLabel
      Left = 8
      Top = 5
      Width = 88
      Height = 21
      AutoSize = False
      Caption = 'PO Start Date'
      Layout = tlCenter
    end
    object Arrival_Label: TLabel
      Left = 275
      Top = 6
      Width = 83
      Height = 21
      AutoSize = False
      Caption = 'PO End Date'
      Layout = tlCenter
    end
    object SizeCode_Label: TLabel
      Left = 8
      Top = 84
      Width = 110
      Height = 20
      AutoSize = False
      Caption = 'PO Number'
      Layout = tlCenter
    end
    object Label1: TLabel
      Left = 8
      Top = 31
      Width = 88
      Height = 21
      AutoSize = False
      Caption = 'Pick Up Date'
      Layout = tlCenter
    end
    object Label2: TLabel
      Left = 275
      Top = 32
      Width = 87
      Height = 20
      AutoSize = False
      Caption = 'Pick Up Time'
      Layout = tlCenter
    end
    object Label3: TLabel
      Left = 8
      Top = 110
      Width = 110
      Height = 20
      AutoSize = False
      Caption = 'PO Qty'
      Layout = tlCenter
    end
    object Label4: TLabel
      Left = 275
      Top = 110
      Width = 74
      Height = 20
      AutoSize = False
      Caption = 'PO Charged'
      Layout = tlCenter
    end
    object AssyCode_ComboBox: TComboBox
      Left = 117
      Top = 58
      Width = 136
      Height = 21
      Style = csDropDownList
      CharCase = ecUpperCase
      ItemHeight = 13
      TabOrder = 4
    end
    object POStart_NUMMIBmDateEdit: TNUMMIBmDateEdit
      Left = 117
      Top = 5
      Width = 121
      Height = 21
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
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
      OnChange = POStart_NUMMIBmDateEditChange
      ReadOnly = False
    end
    object POEnd_NUMMIBmDateEdit: TNUMMIBmDateEdit
      Left = 365
      Top = 6
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
    object PONumber_Edit: TEdit
      Left = 117
      Top = 83
      Width = 82
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 10
      TabOrder = 5
    end
    object PickUp_NUMMIBmDateEdit: TNUMMIBmDateEdit
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
    object POQty_Edit: TEdit
      Left = 117
      Top = 109
      Width = 82
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 10
      TabOrder = 7
    end
    object POCharged_Edit: TEdit
      Left = 364
      Top = 109
      Width = 82
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 10
      ReadOnly = True
      TabOrder = 8
    end
    object PickUpTime_Edit: TMaskEdit
      Left = 365
      Top = 30
      Width = 40
      Height = 21
      CharCase = ecUpperCase
      EditMask = '!00:00;0;_'
      MaxLength = 5
      TabOrder = 3
    end
    object AssyCost_MaskEdit: TcurrEdit
      Left = 365
      Top = 83
      Width = 81
      Height = 21
      TabOrder = 6
      Text = '$0.0000'
      IntDigits = 0
      DecimalDigits = 4
      Prefix = '$'
    end
  end
  object MonthlyPOMaster_DBGrid: TDBGrid
    Left = 20
    Top = 185
    Width = 585
    Height = 132
    DataSource = MonthlyPO_DataSource
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgRowSelect, dgConfirmDelete, dgCancelOnExit]
    TabOrder = 1
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
  end
  object ManagementButtons_Panel: TPanel
    Left = 20
    Top = 324
    Width = 585
    Height = 50
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 2
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
  object MonthlyPO_DataSource: TDataSource
    OnDataChange = MonthlyPO_DataSourceDataChange
    Left = 574
    Top = 44
  end
end
