object ManualForecast_Form: TManualForecast_Form
  Left = 568
  Top = 250
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'Manual Forecast'
  ClientHeight = 407
  ClientWidth = 668
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 433
    Top = 300
    Width = 92
    Height = 13
    Caption = 'Forecast Start Date'
  end
  object Label2: TLabel
    Left = 434
    Top = 327
    Width = 89
    Height = 13
    Caption = 'Forecast End Date'
  end
  object Hist: THistory
    Left = 0
    Top = 0
    Width = 668
    Height = 289
    Align = alTop
    TabOrder = 0
    ShowDate = False
    ShowTime = False
  end
  object Button1: TButton
    Left = 586
    Top = 379
    Width = 75
    Height = 25
    Caption = 'Close'
    TabOrder = 1
    OnClick = Button1Click
  end
  object StartButton: TButton
    Left = 505
    Top = 379
    Width = 75
    Height = 25
    Caption = 'Start'
    TabOrder = 2
    Visible = False
    OnClick = StartButtonClick
  end
  object Startdate: TNUMMIBmDateEdit
    Left = 534
    Top = 296
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
    OnChange = StartdateChange
    ReadOnly = False
  end
  object Enddate: TNUMMIBmDateEdit
    Left = 535
    Top = 324
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
    OnChange = EnddateChange
    ReadOnly = False
  end
end
