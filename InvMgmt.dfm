object InvMgmt_Form: TInvMgmt_Form
  Left = 483
  Top = 119
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsDialog
  Caption = 'Inventory Management'
  ClientHeight = 517
  ClientWidth = 531
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
  object InvMgmt_Label: TLabel
    Left = 20
    Top = 10
    Width = 184
    Height = 20
    Caption = 'Inventory Management'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object SearchKey_GroupBox: TGroupBox
    Left = 21
    Top = 40
    Width = 489
    Height = 134
    Caption = 'Searching Key'
    TabOrder = 0
    object SupplierCode_Label: TLabel
      Left = 8
      Top = 15
      Width = 69
      Height = 20
      AutoSize = False
      Caption = 'Supplier:'
      Layout = tlCenter
    end
    object PartsCode_Label: TLabel
      Left = 8
      Top = 42
      Width = 55
      Height = 20
      AutoSize = False
      Caption = 'Parts Name:'
      Layout = tlCenter
    end
    object TireWheel_Label: TLabel
      Left = 8
      Top = 97
      Width = 76
      Height = 20
      AutoSize = False
      Caption = 'Part Type:'
      Layout = tlCenter
    end
    object Label1: TLabel
      Left = 8
      Top = 69
      Width = 76
      Height = 20
      AutoSize = False
      Caption = 'Line:'
      Layout = tlCenter
    end
    object Supplier_NUMMIColumnComboBox: TNUMMIColumnComboBox
      Left = 96
      Top = 15
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
      TabOrder = 0
      OnChange = Supplier_NUMMIColumnComboBoxChange
      ReadOnly = False
    end
    object PartNum_ComboBox: TComboBox
      Left = 96
      Top = 42
      Width = 150
      Height = 21
      Style = csDropDownList
      CharCase = ecUpperCase
      ItemHeight = 13
      TabOrder = 1
      OnChange = PartNum_ComboBoxChange
    end
    object Line_ComboBox: TComboBox
      Left = 96
      Top = 69
      Width = 150
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 2
      OnChange = Line_ComboBoxChange
    end
    object PartType_ComboBox: TComboBox
      Left = 96
      Top = 97
      Width = 150
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 3
      OnChange = PartType_ComboBoxChange
    end
  end
  object InvMgmt_DBGrid: TDBGrid
    Left = 20
    Top = 180
    Width = 489
    Height = 271
    DataSource = Inventory_DataSource
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit]
    TabOrder = 1
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    OnKeyUp = InvMgmt_DBGridKeyUp
    OnMouseUp = InvMgmt_DBGridMouseUp
  end
  object ManagementButtons_Panel: TPanel
    Left = 20
    Top = 458
    Width = 489
    Height = 50
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 2
    object Print_Button: TButton
      Left = 10
      Top = 10
      Width = 90
      Height = 30
      Caption = '&Print'
      TabOrder = 0
      OnClick = Print_ButtonClick
    end
    object Search_Button: TButton
      Left = 197
      Top = 10
      Width = 90
      Height = 30
      Caption = '&Search'
      TabOrder = 1
      OnClick = Search_ButtonClick
    end
    object Clear_Button: TButton
      Left = 290
      Top = 10
      Width = 90
      Height = 30
      Caption = 'Cl&ear'
      TabOrder = 2
      OnClick = Clear_ButtonClick
    end
    object Close_Button: TButton
      Left = 384
      Top = 10
      Width = 90
      Height = 30
      Caption = '&Close'
      ModalResult = 2
      TabOrder = 3
    end
    object PrintExcel_Button: TButton
      Left = 104
      Top = 10
      Width = 90
      Height = 30
      Caption = 'Print E&xcel'
      TabOrder = 4
      OnClick = Print_ButtonClick
    end
  end
  object Inventory_DataSource: TDataSource
    Left = 116
    Top = 252
  end
end
