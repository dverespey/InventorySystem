object ReportsForm: TReportsForm
  Left = 723
  Top = 298
  BorderStyle = bsSingle
  Caption = '  Reports'
  ClientHeight = 328
  ClientWidth = 552
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -21
  Font.Name = 'Arial'
  Font.Style = [fsBold]
  FormStyle = fsMDIChild
  OldCreateOrder = False
  Position = poDefaultPosOnly
  Visible = True
  OnClose = FormClose
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 24
  object SelectPanel: TPanel
    Left = 0
    Top = 42
    Width = 552
    Height = 137
    Align = alTop
    Alignment = taLeftJustify
    BevelInner = bvRaised
    BevelOuter = bvLowered
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 0
    object Label7: TLabel
      Left = 196
      Top = 43
      Width = 50
      Height = 16
      Caption = 'Printer:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      Visible = False
    end
    object Label8: TLabel
      Left = 8
      Top = 43
      Width = 52
      Height = 16
      Caption = 'Report:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Label2: TLabel
      Left = 8
      Top = 1
      Width = 34
      Height = 16
      Caption = 'Line:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object SelectTypeLabel: TLabel
      Left = 8
      Top = 90
      Width = 89
      Height = 16
      Caption = 'Select Type:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      Visible = False
    end
    object Printer_Combo: TComboBox
      Left = 196
      Top = 59
      Width = 345
      Height = 24
      Style = csDropDownList
      ItemHeight = 16
      Sorted = True
      TabOrder = 1
      Visible = False
    end
    object Line_Combo: TComboBox
      Left = 8
      Top = 17
      Width = 177
      Height = 24
      Style = csDropDownList
      ItemHeight = 16
      TabOrder = 0
      OnChange = Line_ComboChange
    end
    object Report_Edit: TEdit
      Left = 8
      Top = 59
      Width = 175
      Height = 24
      TabStop = False
      ReadOnly = True
      TabOrder = 2
      Text = 'Report_Edit'
    end
    object SelectTypeBox: TComboBox
      Left = 8
      Top = 106
      Width = 176
      Height = 24
      ItemHeight = 16
      Sorted = True
      TabOrder = 3
      Visible = False
      OnChange = SelectTypeBoxChange
      Items.Strings = (
        'FRS'
        'SEQUENCE')
    end
  end
  object FirstPanel: TPanel
    Left = 0
    Top = 220
    Width = 552
    Height = 41
    Align = alTop
    BevelInner = bvRaised
    BevelOuter = bvLowered
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 2
    object Label1: TLabel
      Left = 8
      Top = 13
      Width = 39
      Height = 16
      Caption = 'First: '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Label4: TLabel
      Left = 313
      Top = 11
      Width = 75
      Height = 16
      Caption = 'Sequence:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object StartDate: TDateTimePicker
      Left = 48
      Top = 8
      Width = 102
      Height = 24
      Date = 38720.675901273150000000
      Time = 38720.675901273150000000
      TabOrder = 1
      TabStop = False
    end
    object StartTime: TDateTimePicker
      Left = 167
      Top = 8
      Width = 130
      Height = 24
      Date = 38720.676540960650000000
      Time = 38720.676540960650000000
      Kind = dtkTime
      TabOrder = 2
      TabStop = False
    end
    object StartBox: TComboBox
      Left = 49
      Top = 8
      Width = 256
      Height = 24
      Style = csDropDownList
      ItemHeight = 16
      TabOrder = 3
      OnChange = StartBoxChange
    end
    object Start_Seq_Edit: TEdit
      Left = 395
      Top = 6
      Width = 50
      Height = 24
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      MaxLength = 3
      ParentFont = False
      TabOrder = 0
      OnChange = Start_Seq_EditChange
      OnExit = Start_Seq_EditExit
    end
  end
  object LastPanel: TPanel
    Left = 0
    Top = 261
    Width = 552
    Height = 41
    Align = alTop
    BevelInner = bvRaised
    BevelOuter = bvLowered
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 3
    object Label6: TLabel
      Left = 314
      Top = 13
      Width = 79
      Height = 16
      Caption = 'Sequence: '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Label3: TLabel
      Left = 8
      Top = 13
      Width = 34
      Height = 16
      Caption = 'Last:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object End_Seq_Edit: TEdit
      Left = 396
      Top = 8
      Width = 50
      Height = 24
      MaxLength = 3
      TabOrder = 0
      OnChange = Start_Seq_EditChange
      OnExit = Start_Seq_EditExit
    end
    object EndDate: TDateTimePicker
      Left = 48
      Top = 9
      Width = 102
      Height = 24
      Date = 38720.675901273150000000
      Time = 38720.675901273150000000
      TabOrder = 1
      TabStop = False
    end
    object EndTime: TDateTimePicker
      Left = 167
      Top = 8
      Width = 130
      Height = 24
      Date = 38720.676540960650000000
      Time = 38720.676540960650000000
      Kind = dtkTime
      TabOrder = 2
      TabStop = False
    end
    object EndBox: TComboBox
      Left = 49
      Top = 8
      Width = 256
      Height = 24
      Style = csDropDownList
      ItemHeight = 16
      TabOrder = 3
    end
  end
  object ButtonPanel: TPanel
    Left = 0
    Top = 0
    Width = 552
    Height = 42
    Align = alTop
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 4
    object Print_Button: TBitBtn
      Left = 355
      Top = 8
      Width = 94
      Height = 25
      Caption = '&Print'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'Arial'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      OnClick = Print_ButtonClick
      Glyph.Data = {
        DE010000424DDE01000000000000760000002800000024000000120000000100
        0400000000006801000000000000000000001000000000000000000000000000
        80000080000000808000800000008000800080800000C0C0C000808080000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
        3333333333333333333333330000333333333333333333333333F33333333333
        00003333344333333333333333388F3333333333000033334224333333333333
        338338F3333333330000333422224333333333333833338F3333333300003342
        222224333333333383333338F3333333000034222A22224333333338F338F333
        8F33333300003222A3A2224333333338F3838F338F33333300003A2A333A2224
        33333338F83338F338F33333000033A33333A222433333338333338F338F3333
        0000333333333A222433333333333338F338F33300003333333333A222433333
        333333338F338F33000033333333333A222433333333333338F338F300003333
        33333333A222433333333333338F338F00003333333333333A22433333333333
        3338F38F000033333333333333A223333333333333338F830000333333333333
        333A333333333333333338330000333333333333333333333333333333333333
        0000}
      NumGlyphs = 2
    end
    object Close_Button: TBitBtn
      Left = 449
      Top = 8
      Width = 94
      Height = 25
      Caption = '&Close'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'Arial'
      Font.Style = []
      ModalResult = 1
      ParentFont = False
      TabOrder = 1
      OnClick = Close_ButtonClick
      Glyph.Data = {
        DE010000424DDE01000000000000760000002800000024000000120000000100
        0400000000006801000000000000000000001000000000000000000000000000
        80000080000000808000800000008000800080800000C0C0C000808080000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00388888888877
        F7F787F8888888888333333F00004444400888FFF444448888888888F333FF8F
        000033334D5007FFF4333388888888883338888F0000333345D50FFFF4333333
        338F888F3338F33F000033334D5D0FFFF43333333388788F3338F33F00003333
        45D50FEFE4333333338F878F3338F33F000033334D5D0FFFF43333333388788F
        3338F33F0000333345D50FEFE4333333338F878F3338F33F000033334D5D0FFF
        F43333333388788F3338F33F0000333345D50FEFE4333333338F878F3338F33F
        000033334D5D0EFEF43333333388788F3338F33F0000333345D50FEFE4333333
        338F878F3338F33F000033334D5D0EFEF43333333388788F3338F33F00003333
        4444444444333333338F8F8FFFF8F33F00003333333333333333333333888888
        8888333F00003333330000003333333333333FFFFFF3333F00003333330AAAA0
        333333333333888888F3333F00003333330000003333333333338FFFF8F3333F
        0000}
      NumGlyphs = 2
    end
  end
  object StatusBar: TStatusBar
    Left = 0
    Top = 309
    Width = 552
    Height = 19
    Panels = <>
    SimplePanel = True
  end
  object FRSPanel: TPanel
    Left = 0
    Top = 179
    Width = 552
    Height = 41
    Align = alTop
    BevelInner = bvRaised
    BevelOuter = bvLowered
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 1
    Visible = False
    object Label9: TLabel
      Left = 8
      Top = 13
      Width = 39
      Height = 16
      Caption = 'FRS: '
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object DateTimePicker1: TDateTimePicker
      Left = 48
      Top = 8
      Width = 102
      Height = 24
      Date = 38720.675901273150000000
      Time = 38720.675901273150000000
      TabOrder = 0
      TabStop = False
    end
    object DateTimePicker2: TDateTimePicker
      Left = 167
      Top = 8
      Width = 130
      Height = 24
      Date = 38720.676540960650000000
      Time = 38720.676540960650000000
      Kind = dtkTime
      TabOrder = 1
      TabStop = False
    end
    object FRSBox: TComboBox
      Left = 49
      Top = 8
      Width = 256
      Height = 24
      Style = csDropDownList
      ItemHeight = 16
      TabOrder = 2
      OnChange = StartBoxChange
    end
  end
  object GetDataPanel: TPanel
    Left = 24
    Top = 97
    Width = 515
    Height = 97
    Caption = 'GetDataPanel'
    Color = clMoneyGreen
    TabOrder = 6
    Visible = False
    object Panel2: TPanel
      Left = 18
      Top = 16
      Width = 481
      Height = 69
      Caption = 'Retrieving Data, please wait'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -24
      Font.Name = 'Arial'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 0
    end
  end
  object Selected_DataSource: TDataSource
    DataSet = Selected_Query
    Left = 520
    Top = 44
  end
  object Report_Query: TADOQuery
    Parameters = <>
    Left = 492
    Top = 71
  end
  object Selected_Query: TADOQuery
    Parameters = <>
    SQL.Strings = (
      'SELECT SUBSTRING(input, 1, 4) AS Sequence,'
      '              SUBSTRING(input, 5, 4) AS Tire, Printed'
      ''
      'FROM Activity'
      ''
      'WHERE input IS NOT NULL'
      '      AND Activity_Id >=0'
      '      AND Activity_Id <=100'
      ''
      'ORDER BY Activity_Id'
      '')
    Left = 492
    Top = 43
  end
  object UpdateReportQuery: TADOQuery
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM FRS'
      '')
    Left = 492
    Top = 100
  end
  object UpdateReportQuerywf: TADOQuery
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM FRSwf'
      '')
    Left = 492
    Top = 130
  end
  object Line_DataSet: TADODataSet
    Connection = AdminDataModule.VehicleOrderConnection
    CommandText = 'Select * FROM Line'
    Parameters = <>
    Left = 400
    Top = 44
  end
  object GetLastPrint_DataSet: TADODataSet
    Connection = AdminDataModule.VehicleOrderConnection
    Parameters = <>
    Left = 431
    Top = 44
  end
  object Report_DataSet: TADODataSet
    Connection = AdminDataModule.VehicleOrderConnection
    Parameters = <>
    Left = 463
    Top = 44
  end
  object UpdateReportCommand: TADOCommand
    Connection = AdminDataModule.VehicleOrderConnection
    Parameters = <>
    Left = 402
    Top = 74
  end
  object FRSDataSet: TADODataSet
    Connection = AdminDataModule.VehicleOrderConnection
    CommandText = 'AD_GetFRSList'
    Parameters = <>
    Left = 431
    Top = 72
  end
end
