object FirstProductionDay_Form: TFirstProductionDay_Form
  Left = 544
  Top = 269
  Width = 614
  Height = 496
  Caption = 'First Production Day of the Year'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object ManagementButtons_Panel: TPanel
    Left = 8
    Top = 402
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
    object Search_Button: TButton
      Left = 247
      Top = 10
      Width = 90
      Height = 30
      Caption = '&Search'
      TabOrder = 2
      OnClick = Search_ButtonClick
    end
    object Clear_Button: TButton
      Left = 366
      Top = 10
      Width = 90
      Height = 30
      Caption = 'Cl&ear'
      TabOrder = 3
      OnClick = Clear_ButtonClick
    end
    object Close_Button: TButton
      Left = 485
      Top = 10
      Width = 90
      Height = 30
      Caption = '&Close'
      ModalResult = 2
      TabOrder = 4
    end
    object Delete_Button: TButton
      Left = 128
      Top = 10
      Width = 90
      Height = 30
      Caption = '&Delete'
      TabOrder = 1
      OnClick = Delete_ButtonClick
    end
  end
  object Event_DBGrid: TDBGrid
    Left = 9
    Top = 131
    Width = 580
    Height = 261
    DataSource = FirstProductionDay_DataSource
    Options = [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs, dgConfirmDelete, dgCancelOnExit]
    TabOrder = 1
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    OnKeyUp = Event_DBGridKeyUp
    OnMouseUp = Event_DBGridMouseUp
  end
  object Event_Panel: TPanel
    Left = 8
    Top = 10
    Width = 584
    Height = 115
    BevelInner = bvRaised
    BevelOuter = bvNone
    BorderStyle = bsSingle
    TabOrder = 2
    object Label1: TLabel
      Left = 14
      Top = 18
      Width = 54
      Height = 13
      Caption = 'Event Date'
    end
    object Label4: TLabel
      Left = 360
      Top = 18
      Width = 76
      Height = 13
      Caption = 'Production Year'
    end
    object Label2: TLabel
      Left = 358
      Top = 58
      Width = 69
      Height = 13
      Caption = 'Week Number'
    end
    object Year_Edit: TEdit
      Left = 448
      Top = 13
      Width = 40
      Height = 21
      ReadOnly = True
      TabOrder = 0
      Text = 'Year_Edit'
    end
    object WeekNumber_Edit: TEdit
      Left = 448
      Top = 51
      Width = 29
      Height = 21
      ReadOnly = True
      TabOrder = 1
      Text = 'WeekNumber_Edit'
    end
  end
  object EventDate_NUMMIBmDateEdit: TBmDateEdit
    Left = 155
    Top = 24
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
    OnChange = EventDate_NUMMIBmDateEditChange
  end
  object FirstProductionDay_DataSource: TDataSource
    Left = 556
    Top = 18
  end
end
