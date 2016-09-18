object RecConfStat_Form: TRecConfStat_Form
  Left = 879
  Top = 238
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsDialog
  Caption = 'Receiving/Confirmation/Status Management'
  ClientHeight = 617
  ClientWidth = 580
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object RecConfStat_Label: TLabel
    Left = 20
    Top = 10
    Width = 373
    Height = 20
    Caption = 'Receiving / Confirmation / Status Management'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object ManagementButtons_Panel: TPanel
    Left = 13
    Top = 561
    Width = 553
    Height = 50
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 1
    object Insert_Button: TButton
      Left = 5
      Top = 10
      Width = 90
      Height = 30
      Caption = '&Insert'
      TabOrder = 0
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
      OnClick = Search_ButtonClick
    end
    object Clear_Button: TButton
      Left = 367
      Top = 10
      Width = 90
      Height = 30
      Caption = 'Cl&ear'
      TabOrder = 3
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
    end
    object Delete_Button: TButton
      Left = 186
      Top = 10
      Width = 90
      Height = 30
      Caption = '&Delete'
      TabOrder = 5
      OnClick = Delete_ButtonClick
    end
  end
  object RecConfStat_Panel: TPanel
    Left = 14
    Top = 77
    Width = 552
    Height = 355
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 0
    object InTransit_Label: TLabel
      Left = 286
      Top = 168
      Width = 88
      Height = 21
      AutoSize = False
      Caption = 'In Transit:'
      Layout = tlCenter
    end
    object Arrival_Label: TLabel
      Left = 286
      Top = 194
      Width = 83
      Height = 21
      AutoSize = False
      Caption = 'Arrival Date:'
      Layout = tlCenter
    end
    object TrailerNo_Label: TLabel
      Left = 21
      Top = 168
      Width = 90
      Height = 21
      AutoSize = False
      Caption = 'Trailer No.:'
      Layout = tlCenter
    end
    object PlantYard_Label: TLabel
      Left = 286
      Top = 247
      Width = 87
      Height = 21
      AutoSize = False
      Caption = 'NUMMI Yard:'
      Layout = tlCenter
    end
    object ParkingSpot_Label: TLabel
      Left = 21
      Top = 194
      Width = 90
      Height = 21
      AutoSize = False
      Caption = 'Parking Spot:'
      Layout = tlCenter
    end
    object AssemblerYard_Label: TLabel
      Left = 286
      Top = 273
      Width = 75
      Height = 21
      AutoSize = False
      Caption = 'WQS Yard:'
      Layout = tlCenter
    end
    object AssemblerLocation_Label: TLabel
      Left = 21
      Top = 221
      Width = 92
      Height = 21
      AutoSize = False
      Caption = 'WQS Location:'
      Layout = tlCenter
    end
    object EmptyTrailer_Label: TLabel
      Left = 286
      Top = 299
      Width = 73
      Height = 21
      AutoSize = False
      Caption = 'Empty Trailer:'
      Layout = tlCenter
    end
    object Detention_Label: TLabel
      Left = 21
      Top = 247
      Width = 82
      Height = 21
      AutoSize = False
      Caption = 'Detention:'
      Layout = tlCenter
    end
    object Quantity_Label: TLabel
      Left = 21
      Top = 274
      Width = 45
      Height = 20
      AutoSize = False
      Caption = 'Quantity:'
      Layout = tlCenter
    end
    object Label1: TLabel
      Left = 286
      Top = 220
      Width = 87
      Height = 21
      AutoSize = False
      Caption = 'Warehouse'
      Layout = tlCenter
    end
    object Label2: TLabel
      Left = 286
      Top = 326
      Width = 73
      Height = 21
      AutoSize = False
      Caption = 'Terminated:'
      Layout = tlCenter
    end
    object SearchKey_GroupBox: TGroupBox
      Left = 16
      Top = 6
      Width = 519
      Height = 159
      Caption = 'Searching Key'
      TabOrder = 0
      object SupplierCode_Label: TLabel
        Left = 5
        Top = 20
        Width = 69
        Height = 21
        AutoSize = False
        Caption = 'Supplier Code:'
        Layout = tlCenter
      end
      object FRS_Label: TLabel
        Left = 5
        Top = 49
        Width = 44
        Height = 21
        AutoSize = False
        Caption = 'FRS No.:'
        Layout = tlCenter
      end
      object PartsCode_Label: TLabel
        Left = 270
        Top = 20
        Width = 55
        Height = 21
        AutoSize = False
        Caption = 'Parts Code:'
        Layout = tlCenter
      end
      object RENBAN_Label: TLabel
        Left = 270
        Top = 78
        Width = 48
        Height = 21
        AutoSize = False
        Caption = 'RENBAN:'
        Layout = tlCenter
      end
      object Label3: TLabel
        Left = 5
        Top = 78
        Width = 87
        Height = 21
        AutoSize = False
        Caption = 'Order Date'
        Layout = tlCenter
      end
      object Label4: TLabel
        Left = 5
        Top = 132
        Width = 87
        Height = 21
        AutoSize = False
        Caption = 'Ship Date'
        Layout = tlCenter
      end
      object Label5: TLabel
        Left = 270
        Top = 49
        Width = 72
        Height = 21
        AutoSize = False
        Caption = 'Kanban Code:'
        Layout = tlCenter
      end
      object FRS_Edit: TEdit
        Left = 100
        Top = 49
        Width = 121
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 7
        TabOrder = 2
      end
      object RENBAN_Edit: TEdit
        Left = 360
        Top = 78
        Width = 121
        Height = 21
        CharCase = ecUpperCase
        MaxLength = 10
        TabOrder = 3
      end
      object SupplierCode_ComboBox: TComboBox
        Left = 100
        Top = 20
        Width = 121
        Height = 21
        CharCase = ecUpperCase
        ItemHeight = 13
        TabOrder = 0
        OnChange = SupplierCode_ComboBoxChange
      end
      object PartsCode_ComboBox: TComboBox
        Left = 360
        Top = 20
        Width = 121
        Height = 21
        CharCase = ecUpperCase
        ItemHeight = 13
        TabOrder = 1
        OnChange = PartsCode_ComboBoxChange
      end
      object Order_NUMMIBmDateEdit: TNUMMIBmDateEdit
        Left = 100
        Top = 78
        Width = 121
        Height = 21
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 4
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
        OnExit = InTransit_NUMMIBmDateEditExit
        ReadOnly = False
      end
      object RenbanUpdate_CheckBox: TCheckBox
        Left = 269
        Top = 108
        Width = 103
        Height = 17
        Alignment = taLeftJustify
        Caption = 'Renban Update?'
        TabOrder = 5
      end
      object Ship_NUMMIBmDateEdit: TNUMMIBmDateEdit
        Left = 100
        Top = 132
        Width = 121
        Height = 21
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 6
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
        OnExit = InTransit_NUMMIBmDateEditExit
        ReadOnly = False
      end
      object KanbanCode_ComboBox: TComboBox
        Left = 360
        Top = 49
        Width = 121
        Height = 21
        CharCase = ecUpperCase
        ItemHeight = 13
        TabOrder = 7
      end
      object Unordered_Box: TCheckBox
        Left = 5
        Top = 106
        Width = 108
        Height = 17
        Alignment = taLeftJustify
        Caption = 'No Order Date'
        TabOrder = 8
      end
    end
    object TrailerNo_Edit: TEdit
      Left = 116
      Top = 170
      Width = 121
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 10
      TabOrder = 8
      OnChange = TrailerNo_EditChange
    end
    object ParkingSpot_Edit: TEdit
      Left = 116
      Top = 196
      Width = 121
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 10
      TabOrder = 9
      OnChange = TrailerNo_EditChange
    end
    object AssemblerLocation_Edit: TEdit
      Left = 116
      Top = 222
      Width = 121
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 10
      TabOrder = 10
      Visible = False
    end
    object Detention_Edit: TEdit
      Left = 116
      Top = 248
      Width = 121
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 50
      TabOrder = 11
      OnChange = TrailerNo_EditChange
    end
    object InTransit_NUMMIBmDateEdit: TNUMMIBmDateEdit
      Left = 376
      Top = 170
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
      OnChange = TrailerNo_EditChange
      OnExit = InTransit_NUMMIBmDateEditExit
      ReadOnly = False
    end
    object ArrivalDate_NUMMIBmDateEdit: TNUMMIBmDateEdit
      Left = 376
      Top = 196
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
      OnChange = TrailerNo_EditChange
      OnExit = InTransit_NUMMIBmDateEditExit
      ReadOnly = False
    end
    object AssemblerYard_NUMMIBmDateEdit: TNUMMIBmDateEdit
      Left = 376
      Top = 274
      Width = 121
      Height = 21
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 5
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
      OnChange = TrailerNo_EditChange
      OnExit = InTransit_NUMMIBmDateEditExit
      ReadOnly = False
    end
    object EmptyTrailer_NUMMIBmDateEdit: TNUMMIBmDateEdit
      Left = 376
      Top = 300
      Width = 121
      Height = 21
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 6
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
      OnChange = TrailerNo_EditChange
      OnExit = InTransit_NUMMIBmDateEditExit
      ReadOnly = False
    end
    object PlantYard_NUMMIBmDateEdit: TNUMMIBmDateEdit
      Left = 376
      Top = 248
      Width = 121
      Height = 21
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 4
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
      OnChange = TrailerNo_EditChange
      OnExit = InTransit_NUMMIBmDateEditExit
      ReadOnly = False
    end
    object Warehouse_NUMMIBmDateEdit: TNUMMIBmDateEdit
      Left = 376
      Top = 222
      Width = 121
      Height = 21
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 3
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
      OnChange = TrailerNo_EditChange
      OnExit = InTransit_NUMMIBmDateEditExit
      ReadOnly = False
    end
    object Terminated_NUMMIBmDateEdit: TNUMMIBmDateEdit
      Left = 376
      Top = 326
      Width = 121
      Height = 21
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 7
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
      OnChange = TrailerNo_EditChange
      OnExit = InTransit_NUMMIBmDateEditExit
      ReadOnly = False
    end
    object AssemblerLocation_ComboBox: TComboBox
      Left = 116
      Top = 221
      Width = 123
      Height = 21
      CharCase = ecUpperCase
      ItemHeight = 13
      TabOrder = 12
      OnChange = TrailerNo_EditChange
    end
    object Quantity_Edit: TEdit
      Left = 116
      Top = 273
      Width = 121
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 6
      TabOrder = 13
      OnChange = TrailerNo_EditChange
    end
  end
  object RecConfStat_DBGrid: TDBGrid
    Left = 12
    Top = 437
    Width = 553
    Height = 120
    DataSource = RecConf_DataSource
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgRowSelect, dgConfirmDelete, dgCancelOnExit]
    TabOrder = 2
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    OnKeyUp = RecConfStat_DBGridKeyUp
    OnMouseUp = RecConfStat_DBGridMouseUp
  end
  object Panel1: TPanel
    Left = 15
    Top = 34
    Width = 551
    Height = 41
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 3
    object Label6: TLabel
      Left = 24
      Top = 11
      Width = 69
      Height = 21
      AutoSize = False
      Caption = 'Sort By:'
      Layout = tlCenter
    end
    object HideTerminated_CheckBox: TCheckBox
      Left = 388
      Top = 12
      Width = 112
      Height = 17
      Alignment = taLeftJustify
      Caption = 'Hide Terminated'
      TabOrder = 0
      OnClick = HideTerminated_CheckBoxClick
    end
    object SortBy_ComboBox: TComboBox
      Left = 116
      Top = 11
      Width = 159
      Height = 21
      ItemHeight = 13
      TabOrder = 1
      Text = 'SortBy_ComboBox'
      OnChange = SortBy_ComboBoxChange
      Items.Strings = (
        'Default'
        'Supplier'
        'Parts'
        'FRSNo'
        'Kanban'
        'Renban'
        'Order'
        'Shiped')
    end
  end
  object RecConf_DataSource: TDataSource
    OnDataChange = RecConf_DataSourceDataChange
    Left = 142
    Top = 429
  end
end
