object HolidayOvertimeForm: THolidayOvertimeForm
  Left = 537
  Top = 246
  BorderStyle = bsSingle
  Caption = 'Holiday/Overtime'
  ClientHeight = 349
  ClientWidth = 856
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Arial'
  Font.Style = []
  FormStyle = fsMDIChild
  Icon.Data = {
    0000010001002020100000000000E80200001600000028000000200000004000
    0000010004000000000080020000000000000000000000000000000000000000
    000000008000008000000080800080000000800080008080000080808000C0C0
    C0000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00CCCC
    CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCFF
    FFFFFFFFFFFFFFFFFFFFFFFFFFCCCCFFFFFFFFFFFFFFFFFFFFFCCCCFFFCCCCFF
    FFFFFFFFFFFFFFFFFFFFFFFCFFCCCCFFFFFFFFFCFFCFFCFFCFFFCCFCFFCCCCFF
    FFFFFFFCFFCFFCFFCFFCFFCCFFCCCCFFFFFFFFFCFFCFFCFFCFFCFFFCFFCCCCFF
    FFFFFFFCFFCFFCFFCFFCFFFCFFCCCCFFFFFFFFFCCFCFFCFFCFFCFFCCFFCCCCFF
    FFFFFFFCFCCCCFFCCCFFCCFCFFCCCCFFFFFFFFFFFFFFFFFFCFFFFFFFFFCCCCFF
    FFFFFFFFFFFFFFFFFCFFFFFFFFCCCCFFFFFFFFFFFFFFFFFFFFFFFFFFFFCCCCFF
    FFFFFFFFFFFFFFFFFFFFFFFFFFCCCCFFFFFFFFFFFFFFFFFFFFFFFFFFFFCCCCFF
    FFFFFFFFFFFFFFFFFFFFFFFFFFCCCCFFFFFFFFCFFFFFFFFFFFFFFFFFFFCCCCFF
    FFFFFFCFFFFFFFFFFFFFFFFFFFCCCCFFCFFFFFCFCCFFFFFFFFFFFFFFFFCCCCFF
    CFFFFFCCFFCFFFFFFFFFFFFFFFCCCCFFCFFFFFCFFFCFFFFFFFFFFFFFFFCCCCFF
    CFFFFFCFFFCFFFFFFFFFFFFFFFCCCCFFCCCCFFCCFFCFFFFFFFFFFFFFFFCCCCFF
    CFFFFFCFCCFFFFFFFFFFFFFFFFCCCCFFCFFFFFFFFFFFFFFFFFFFFFFFFFCCCCFF
    CCCCCFFFFFFFFFFFFFFFFFFFFFCCCCFFFFFFFFFFFFFFFFFFFFFFFFFFFFCCCCFF
    FFFFFFFFFFFFFFFFFFFFFFFFFFCCCCFFFFFFFFFFFFFFFFFFFFFFFFFFFFCCCCCC
    CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC0000
    0000000000000000000000000000000000000000000000000000000000000000
    0000000000000000000000000000000000000000000000000000000000000000
    0000000000000000000000000000000000000000000000000000000000000000
    000000000000000000000000000000000000000000000000000000000000}
  OldCreateOrder = False
  Position = poDefault
  Visible = True
  OnClose = FormClose
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 14
  object DBGrid1: TDBGrid
    Left = 0
    Top = 164
    Width = 856
    Height = 185
    Align = alBottom
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgConfirmDelete, dgCancelOnExit]
    TabOrder = 0
    TitleFont.Charset = ANSI_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Arial'
    TitleFont.Style = []
  end
  object Panel1: TPanel
    Left = 769
    Top = 0
    Width = 87
    Height = 164
    Align = alRight
    TabOrder = 1
    object ClearButton: TButton
      Left = 7
      Top = 3
      Width = 75
      Height = 25
      Caption = 'Clear'
      TabOrder = 0
      OnClick = ClearButtonClick
    end
    object InsertButton: TButton
      Left = 7
      Top = 55
      Width = 75
      Height = 25
      Caption = 'Insert'
      TabOrder = 1
      OnClick = InsertButtonClick
    end
    object UpdateButton: TButton
      Left = 7
      Top = 81
      Width = 75
      Height = 25
      Caption = 'Update'
      TabOrder = 2
      OnClick = UpdateButtonClick
    end
    object DeleteButton: TButton
      Left = 7
      Top = 107
      Width = 75
      Height = 25
      Caption = 'Delete'
      TabOrder = 3
      OnClick = DeleteButtonClick
    end
    object SearchButton: TButton
      Left = 7
      Top = 29
      Width = 75
      Height = 25
      Caption = 'Search'
      TabOrder = 4
      OnClick = SearchButtonClick
    end
    object CloseButton: TButton
      Left = 7
      Top = 133
      Width = 75
      Height = 25
      Caption = 'Close'
      TabOrder = 5
      OnClick = CloseButtonClick
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 769
    Height = 164
    Align = alClient
    TabOrder = 2
    object Label1: TLabel
      Left = 15
      Top = 18
      Width = 50
      Height = 14
      Caption = 'Line Name'
    end
    object Label2: TLabel
      Left = 15
      Top = 46
      Width = 56
      Height = 14
      Caption = 'Date Status'
    end
    object Label3: TLabel
      Left = 15
      Top = 74
      Width = 54
      Height = 14
      Caption = 'Description'
    end
    object Label4: TLabel
      Left = 15
      Top = 102
      Width = 22
      Height = 14
      Caption = 'Date'
    end
    object LineNameComboBox: TComboBox
      Left = 124
      Top = 14
      Width = 145
      Height = 22
      ItemHeight = 14
      TabOrder = 0
      Text = 'LineNameComboBox'
      OnChange = LineNameComboBoxChange
    end
    object SpecialDate: TDateTimePicker
      Left = 124
      Top = 98
      Width = 102
      Height = 22
      Date = 38720.675901273150000000
      Time = 38720.675901273150000000
      TabOrder = 1
      OnChange = SpecialDateChange
    end
    object DescriptionnEdit: TEdit
      Left = 124
      Top = 70
      Width = 327
      Height = 22
      MaxLength = 50
      TabOrder = 2
      Text = '12345678901234567890123456789012345678901234567890'
      OnChange = LineNameComboBoxChange
    end
    object EventType_NUMMIColumnComboBox: TNUMMIColumnComboBox
      Left = 124
      Top = 41
      Width = 58
      Height = 22
      Ctl3D = True
      Columns = <
        item
          Color = clWindow
          ColumnType = ctText
          Width = 100
          Alignment = taLeftJustify
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
        end
        item
          Color = clWindow
          ColumnType = ctText
          Width = 100
          Alignment = taLeftJustify
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
        end>
      ComboItems = <
        item
          ImageIndex = -1
          Strings.Strings = (
            'O'
            'OVERTIME')
          Tag = 0
        end
        item
          ImageIndex = -1
          Strings.Strings = (
            'H'
            'HOLIDAY')
          Tag = 0
        end>
      EditColumn = 0
      EditHeight = 16
      DropWidth = 200
      DropHeight = 200
      Etched = False
      Flat = False
      FocusBorder = False
      GridLines = False
      ItemIndex = 0
      LookupColumn = 0
      ParentCtl3D = False
      ShowItemHint = False
      SortColumn = 0
      Sorted = False
      TabOrder = 3
      ReadOnly = False
    end
  end
end
