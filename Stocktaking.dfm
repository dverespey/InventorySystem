object Stocktaking_Form: TStocktaking_Form
  Left = 685
  Top = 281
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsDialog
  Caption = 'Stocktaking Adjustment'
  ClientHeight = 494
  ClientWidth = 636
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object RecProdReject_Label: TLabel
    Left = 15
    Top = 10
    Width = 192
    Height = 20
    Caption = 'Stocktaking Adjustment'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Stocktaking_DBGrid: TDBGrid
    Left = 10
    Top = 234
    Width = 615
    Height = 189
    DataSource = Stocktaking_DataSource
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit]
    TabOrder = 1
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    OnKeyUp = Stocktaking_DBGridKeyUp
    OnMouseUp = Stocktaking_DBGridMouseUp
  end
  object ManagementButtons_Panel: TPanel
    Left = 10
    Top = 431
    Width = 615
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
      Left = 110
      Top = 10
      Width = 90
      Height = 30
      Caption = '&Update'
      TabOrder = 1
      OnClick = Update_ButtonClick
    end
    object Search_Button: TButton
      Left = 310
      Top = 10
      Width = 90
      Height = 30
      Caption = '&Search'
      TabOrder = 3
      OnClick = Search_ButtonClick
    end
    object Clear_Button: TButton
      Left = 412
      Top = 10
      Width = 90
      Height = 30
      Caption = 'Cl&ear'
      TabOrder = 4
      OnClick = Clear_ButtonClick
    end
    object Close_Button: TButton
      Left = 518
      Top = 10
      Width = 90
      Height = 30
      Caption = '&Close'
      ModalResult = 2
      TabOrder = 5
    end
    object Delete_Button: TButton
      Left = 210
      Top = 10
      Width = 90
      Height = 30
      Caption = '&Delete'
      TabOrder = 2
      OnClick = Delete_ButtonClick
    end
  end
  object Stocktaking_Panel: TPanel
    Left = 10
    Top = 34
    Width = 615
    Height = 198
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 0
    object Reason_Label: TLabel
      Left = 16
      Top = 130
      Width = 40
      Height = 13
      Caption = 'Reason:'
      Layout = tlCenter
    end
    object Qty_Label: TLabel
      Left = 16
      Top = 100
      Width = 42
      Height = 13
      Caption = 'Quantity:'
      Layout = tlCenter
    end
    object PartsCode_Label: TLabel
      Left = 16
      Top = 70
      Width = 55
      Height = 13
      Caption = 'Parts Code:'
      Layout = tlCenter
    end
    object SupplierCode_Label: TLabel
      Left = 16
      Top = 40
      Width = 69
      Height = 13
      Caption = 'Supplier Code:'
      Layout = tlCenter
    end
    object Date_Label: TLabel
      Left = 16
      Top = 10
      Width = 26
      Height = 13
      Caption = 'Date:'
      Layout = tlCenter
    end
    object Reason_Memo: TMemo
      Left = 146
      Top = 130
      Width = 420
      Height = 60
      Lines.Strings = (
        'Reason_Memo')
      TabOrder = 3
    end
    object PartsCode_ComboBox: TComboBox
      Left = 146
      Top = 68
      Width = 137
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 1
      OnChange = PartsCode_ComboBoxChange
    end
    object Edit_DateTimePicker: TDateTimePicker
      Left = 146
      Top = 6
      Width = 121
      Height = 21
      Date = 37540.580963750000000000
      Format = 'yyyy/MM/dd'
      Time = 37540.580963750000000000
      DateFormat = dfLong
      TabOrder = 0
    end
    object Qty_MaskEdit: TMaskEdit
      Left = 146
      Top = 99
      Width = 134
      Height = 21
      CharCase = ecUpperCase
      EditMask = '#######;1; '
      MaxLength = 7
      TabOrder = 2
      Text = '0      '
      OnChange = TextChange
      OnExit = MaskEditExit
    end
    object Supplier_NUMMIColumnComboBox: TNUMMIColumnComboBox
      Left = 145
      Top = 38
      Width = 258
      Height = 22
      Ctl3D = True
      Columns = <
        item
          Color = clWindow
          ColumnType = ctText
          Width = 50
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
          Width = 200
          Alignment = taLeftJustify
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = []
        end>
      ComboItems = <>
      EditColumn = -1
      EditHeight = 16
      DropWidth = 258
      DropHeight = 200
      Etched = False
      Flat = False
      FocusBorder = False
      GridLines = False
      ItemIndex = -1
      LookupColumn = 0
      ParentCtl3D = False
      ShowItemHint = False
      SortColumn = 0
      Sorted = False
      TabOrder = 4
      OnChange = Supplier_NUMMIColumnComboBoxChange
      ReadOnly = False
    end
  end
  object Stocktaking_DataSource: TDataSource
    OnDataChange = Stocktaking_DataSourceDataChange
    Left = 154
    Top = 328
  end
end
