object Order_Form: TOrder_Form
  Left = 896
  Top = 293
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsDialog
  Caption = 'Order'
  ClientHeight = 243
  ClientWidth = 338
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
  object Order_Label: TLabel
    Left = 11
    Top = 12
    Width = 46
    Height = 20
    Caption = 'Order'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Label1: TLabel
    Left = 12
    Top = 45
    Width = 33
    Height = 13
    Caption = 'Today:'
    Layout = tlCenter
  end
  object CarTruck_Label: TLabel
    Left = 12
    Top = 82
    Width = 75
    Height = 21
    AutoSize = False
    Caption = 'Line:'
    Layout = tlCenter
  end
  object TireWheel_Label: TLabel
    Left = 12
    Top = 118
    Width = 75
    Height = 20
    AutoSize = False
    Caption = 'Part Type:'
    Layout = tlCenter
  end
  object SortBy_Label: TLabel
    Left = 12
    Top = 156
    Width = 75
    Height = 20
    AutoSize = False
    Caption = 'Sort By:'
    Layout = tlCenter
  end
  object Start_Button: TButton
    Left = 11
    Top = 210
    Width = 90
    Height = 23
    Caption = '&Start'
    TabOrder = 0
    OnClick = Start_ButtonClick
  end
  object Cancel_Button: TButton
    Left = 235
    Top = 210
    Width = 90
    Height = 23
    Caption = '&Cancel'
    TabOrder = 1
    OnClick = Cancel_ButtonClick
  end
  object Date_DateTimePicker: TDateTimePicker
    Left = 114
    Top = 42
    Width = 90
    Height = 20
    Date = 0.997056377316767000
    Format = 'yyyy/MM/dd'
    Time = 0.997056377316767000
    DateFormat = dfLong
    Enabled = False
    TabOrder = 2
  end
  object ProcessOrder_Button: TButton
    Left = 124
    Top = 210
    Width = 90
    Height = 23
    Caption = '&Order'
    TabOrder = 3
    OnClick = ProcessOrder_ButtonClick
  end
  object Line_ComboBox: TComboBox
    Left = 114
    Top = 80
    Width = 145
    Height = 21
    ItemHeight = 13
    TabOrder = 4
    Text = 'Line_ComboBox'
    OnChange = Line_ComboBoxChange
  end
  object PartType_ComboBox: TComboBox
    Left = 114
    Top = 118
    Width = 145
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 5
  end
  object SortBy_ComboBox: TComboBox
    Left = 114
    Top = 156
    Width = 145
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 6
    Items.Strings = (
      'LINE, RENBAN'
      'LINE, PART NUMBER'
      'PART NUMBER')
  end
end
