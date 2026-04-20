object PartsStockMaster_Form: TPartsStockMaster_Form
  Left = 513
  Top = 88
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsDialog
  Caption = 'Parts Stock Master'
  ClientHeight = 607
  ClientWidth = 629
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
  object PartsStockMaster_Label: TLabel
    Left = 20
    Top = 10
    Width = 164
    Height = 20
    Caption = 'Parts / Stock Master'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object ManagementButtons_Panel: TPanel
    Left = 20
    Top = 548
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
      Left = 390
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
  object PartsStockMaster_Panel: TPanel
    Left = 20
    Top = 32
    Width = 585
    Height = 369
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 0
    object Quantity_Label: TLabel
      Left = 18
      Top = 200
      Width = 75
      Height = 20
      AutoSize = False
      Caption = 'Inventory Qty:'
      Layout = tlCenter
    end
    object Remarks_Label: TLabel
      Left = 20
      Top = 341
      Width = 75
      Height = 20
      AutoSize = False
      Caption = 'Remarks:'
      Layout = tlCenter
    end
    object OneLotQty_Label: TLabel
      Left = 18
      Top = 174
      Width = 75
      Height = 20
      AutoSize = False
      Caption = '1 Lot Qty:'
      Layout = tlCenter
    end
    object SizeCode_Label: TLabel
      Left = 18
      Top = 148
      Width = 75
      Height = 20
      AutoSize = False
      Caption = 'Size Code:'
      Layout = tlCenter
    end
    object CarTruck_Label: TLabel
      Left = 18
      Top = 100
      Width = 75
      Height = 20
      AutoSize = False
      Caption = 'Line:'
      Layout = tlCenter
    end
    object TireWheel_Label: TLabel
      Left = 279
      Top = 100
      Width = 75
      Height = 20
      AutoSize = False
      Caption = 'Part Type:'
      Layout = tlCenter
    end
    object PartsName_Label: TLabel
      Left = 18
      Top = 28
      Width = 75
      Height = 20
      AutoSize = False
      Caption = 'Parts Name:'
      Layout = tlCenter
    end
    object SupplierCode_Label: TLabel
      Left = 18
      Top = 52
      Width = 75
      Height = 20
      AutoSize = False
      Caption = 'Supplier:'
      Layout = tlCenter
    end
    object PartsCode_Label: TLabel
      Left = 18
      Top = 4
      Width = 69
      Height = 20
      AutoSize = False
      Caption = 'Parts Code:'
      Layout = tlCenter
    end
    object KanbanCode_Label: TLabel
      Left = 279
      Top = 4
      Width = 83
      Height = 20
      AutoSize = False
      Caption = 'KANBAN Code:'
      Layout = tlCenter
    end
    object Label1: TLabel
      Left = 18
      Top = 124
      Width = 75
      Height = 20
      AutoSize = False
      Caption = 'Renban Group'
      Layout = tlCenter
    end
    object Label2: TLabel
      Left = 18
      Top = 248
      Width = 97
      Height = 20
      AutoSize = False
      Caption = 'Lead Time Override'
      Layout = tlCenter
    end
    object Label3: TLabel
      Left = 279
      Top = 124
      Width = 99
      Height = 20
      AutoSize = False
      Caption = 'Next Renban Count'
      Layout = tlCenter
    end
    object Label4: TLabel
      Left = 279
      Top = 248
      Width = 95
      Height = 20
      AutoSize = False
      Caption = 'Ship Days Override'
      Layout = tlCenter
    end
    object Label5: TLabel
      Left = 18
      Top = 274
      Width = 97
      Height = 20
      AutoSize = False
      Caption = 'Lead Time Daily'
      Layout = tlCenter
    end
    object Label6: TLabel
      Left = 118
      Top = 274
      Width = 9
      Height = 13
      Caption = 'M'
    end
    object Label7: TLabel
      Left = 163
      Top = 274
      Width = 7
      Height = 13
      Caption = 'T'
    end
    object Label8: TLabel
      Left = 208
      Top = 274
      Width = 11
      Height = 13
      Caption = 'W'
    end
    object Label9: TLabel
      Left = 118
      Top = 300
      Width = 13
      Height = 13
      Caption = 'Th'
    end
    object Label10: TLabel
      Left = 165
      Top = 300
      Width = 6
      Height = 13
      Caption = 'F'
    end
    object Label11: TLabel
      Left = 210
      Top = 300
      Width = 7
      Height = 13
      Caption = 'S'
    end
    object Label12: TLabel
      Left = 279
      Top = 274
      Width = 97
      Height = 20
      AutoSize = False
      Caption = 'Ship Days Daily'
      Layout = tlCenter
    end
    object Label13: TLabel
      Left = 390
      Top = 274
      Width = 9
      Height = 13
      Caption = 'M'
    end
    object Label14: TLabel
      Left = 435
      Top = 274
      Width = 7
      Height = 13
      Caption = 'T'
    end
    object Label15: TLabel
      Left = 480
      Top = 274
      Width = 11
      Height = 13
      Caption = 'W'
    end
    object Label16: TLabel
      Left = 390
      Top = 300
      Width = 13
      Height = 13
      Caption = 'Th'
    end
    object Label17: TLabel
      Left = 437
      Top = 300
      Width = 6
      Height = 13
      Caption = 'F'
    end
    object Label18: TLabel
      Left = 482
      Top = 300
      Width = 7
      Height = 13
      Caption = 'S'
    end
    object Label19: TLabel
      Left = 18
      Top = 77
      Width = 89
      Height = 20
      AutoSize = False
      Caption = 'Override Logistics:'
      Layout = tlCenter
    end
    object Label20: TLabel
      Left = 18
      Top = 224
      Width = 75
      Height = 20
      AutoSize = False
      Caption = 'Part Cost:'
      Layout = tlCenter
    end
    object Label21: TLabel
      Left = 281
      Top = 200
      Width = 75
      Height = 20
      AutoSize = False
      Caption = 'Vendor Share:'
      Layout = tlCenter
      Visible = False
    end
    object VendorShare_SpeedButton: TSpeedButton
      Left = 465
      Top = 199
      Width = 23
      Height = 22
      Glyph.Data = {
        76010000424D7601000000000000760000002800000020000000100000000100
        04000000000000010000120B0000120B00001000000000000000000000000000
        800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333000000
        000033333377777777773333330FFFFFFFF03FF3FF7FF33F3FF700300000FF0F
        00F077F777773F737737E00BFBFB0FFFFFF07773333F7F3333F7E0BFBF000FFF
        F0F077F3337773F3F737E0FBFBFBF0F00FF077F3333FF7F77F37E0BFBF00000B
        0FF077F3337777737337E0FBFBFBFBF0FFF077F33FFFFFF73337E0BF0000000F
        FFF077FF777777733FF7000BFB00B0FF00F07773FF77373377373330000B0FFF
        FFF03337777373333FF7333330B0FFFF00003333373733FF777733330B0FF00F
        0FF03333737F37737F373330B00FFFFF0F033337F77F33337F733309030FFFFF
        00333377737FFFFF773333303300000003333337337777777333}
      NumGlyphs = 2
      Visible = False
    end
    object PartsNum_Edit: TEdit
      Left = 120
      Top = 4
      Width = 88
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 12
      TabOrder = 0
      Text = '000000000000'
    end
    object SizeCode_ComboBox: TComboBox
      Left = 120
      Top = 148
      Width = 150
      Height = 21
      AutoDropDown = True
      Style = csDropDownList
      CharCase = ecUpperCase
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ItemHeight = 13
      MaxLength = 6
      ParentFont = False
      TabOrder = 9
    end
    object Remarks_Edit: TEdit
      Left = 120
      Top = 340
      Width = 430
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 300
      TabOrder = 29
    end
    object KanbanNum_Edit: TEdit
      Left = 383
      Top = 4
      Width = 80
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 5
      TabOrder = 1
    end
    object PartsName_Edit: TEdit
      Left = 120
      Top = 28
      Width = 430
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 50
      TabOrder = 2
    end
    object OneLotQty_MaskEdit: TMaskEdit
      Left = 120
      Top = 174
      Width = 48
      Height = 21
      CharCase = ecUpperCase
      EditMask = '99999;1; '
      MaxLength = 5
      TabOrder = 10
      Text = '0    '
      OnChange = TextChange
      OnExit = MaskEditExit
    end
    object Quantity_MaskEdit: TMaskEdit
      Left = 120
      Top = 200
      Width = 150
      Height = 21
      CharCase = ecUpperCase
      EditMask = '99999;1; '
      MaxLength = 5
      ReadOnly = True
      TabOrder = 12
      Text = '0    '
      OnChange = TextChange
      OnExit = MaskEditExit
    end
    object RenbanCode_ComboBox: TComboBox
      Left = 120
      Top = 124
      Width = 73
      Height = 21
      AutoDropDown = True
      Style = csDropDownList
      CharCase = ecUpperCase
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ItemHeight = 13
      MaxLength = 6
      ParentFont = False
      TabOrder = 7
    end
    object LotSizeOrders_CheckBox: TCheckBox
      Left = 279
      Top = 174
      Width = 118
      Height = 17
      Alignment = taLeftJustify
      Caption = 'Lot Size Orders'
      TabOrder = 11
      OnClick = LotSizeOrders_CheckBoxClick
    end
    object LeadTime_MaskEdit: TMaskEdit
      Left = 120
      Top = 248
      Width = 28
      Height = 21
      CharCase = ecUpperCase
      EditMask = '99;1; '
      MaxLength = 2
      TabOrder = 15
      Text = '  '
      OnChange = TextChange
      OnExit = MaskEditExit
    end
    object RenbanCount_Edit: TMaskEdit
      Left = 383
      Top = 124
      Width = 37
      Height = 21
      CharCase = ecUpperCase
      EditMask = '999;1; '
      MaxLength = 3
      TabOrder = 8
      Text = '   '
      OnChange = TextChange
      OnExit = MaskEditExit
    end
    object ShipDays_Edit: TEdit
      Tag = 1
      Left = 387
      Top = 248
      Width = 28
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 3
      TabOrder = 22
    end
    object LeadtimeMonday_MaskEdit: TMaskEdit
      Left = 134
      Top = 274
      Width = 28
      Height = 21
      CharCase = ecUpperCase
      EditMask = '99;1; '
      MaxLength = 2
      TabOrder = 16
      Text = '  '
      OnChange = TextChange
      OnExit = MaskEditExit
    end
    object LeadtimeTuesday_MaskEdit: TMaskEdit
      Left = 174
      Top = 274
      Width = 28
      Height = 21
      CharCase = ecUpperCase
      EditMask = '99;1; '
      MaxLength = 2
      TabOrder = 17
      Text = '  '
      OnChange = TextChange
      OnExit = MaskEditExit
    end
    object LeadtimeWednesday_MaskEdit: TMaskEdit
      Left = 222
      Top = 274
      Width = 28
      Height = 21
      CharCase = ecUpperCase
      EditMask = '99;1; '
      MaxLength = 2
      TabOrder = 18
      Text = '  '
      OnChange = TextChange
      OnExit = MaskEditExit
    end
    object LeadtimeThursday_MaskEdit: TMaskEdit
      Left = 134
      Top = 300
      Width = 28
      Height = 21
      CharCase = ecUpperCase
      EditMask = '99;1; '
      MaxLength = 2
      TabOrder = 19
      Text = '  '
      OnChange = TextChange
      OnExit = MaskEditExit
    end
    object LeadtimeFriday_MaskEdit: TMaskEdit
      Left = 174
      Top = 300
      Width = 28
      Height = 21
      CharCase = ecUpperCase
      EditMask = '99;1; '
      MaxLength = 2
      TabOrder = 20
      Text = '  '
      OnChange = TextChange
      OnExit = MaskEditExit
    end
    object LeadtimeSaturday_MaskEdit: TMaskEdit
      Left = 222
      Top = 300
      Width = 28
      Height = 21
      CharCase = ecUpperCase
      EditMask = '99;1; '
      MaxLength = 2
      TabOrder = 21
      Text = '  '
      OnChange = TextChange
      OnExit = MaskEditExit
    end
    object ShipDaysMonday_MaskEdit: TMaskEdit
      Left = 406
      Top = 274
      Width = 28
      Height = 21
      CharCase = ecUpperCase
      EditMask = '99;1; '
      MaxLength = 2
      TabOrder = 23
      Text = '  '
      OnChange = TextChange
      OnExit = MaskEditExit
    end
    object ShipDaysTuesday_MaskEdit: TMaskEdit
      Left = 446
      Top = 274
      Width = 28
      Height = 21
      CharCase = ecUpperCase
      EditMask = '99;1; '
      MaxLength = 2
      TabOrder = 24
      Text = '  '
      OnChange = TextChange
      OnExit = MaskEditExit
    end
    object ShipDaysWednesday_MaskEdit: TMaskEdit
      Left = 494
      Top = 274
      Width = 28
      Height = 21
      CharCase = ecUpperCase
      EditMask = '99;1; '
      MaxLength = 2
      TabOrder = 25
      Text = '  '
      OnChange = TextChange
      OnExit = MaskEditExit
    end
    object ShipDaysThursday_MaskEdit: TMaskEdit
      Left = 406
      Top = 300
      Width = 28
      Height = 21
      CharCase = ecUpperCase
      EditMask = '99;1; '
      MaxLength = 2
      TabOrder = 26
      Text = '  '
      OnChange = TextChange
      OnExit = MaskEditExit
    end
    object ShipDaysFriday_MaskEdit: TMaskEdit
      Left = 446
      Top = 300
      Width = 28
      Height = 21
      CharCase = ecUpperCase
      EditMask = '99;1; '
      MaxLength = 2
      TabOrder = 27
      Text = '  '
      OnChange = TextChange
      OnExit = MaskEditExit
    end
    object ShipDaysSaturday_MaskEdit: TMaskEdit
      Left = 494
      Top = 300
      Width = 28
      Height = 21
      CharCase = ecUpperCase
      EditMask = '99;1; '
      MaxLength = 2
      TabOrder = 28
      Text = '  '
      OnChange = TextChange
      OnExit = MaskEditExit
    end
    object Logistics_ComboBox: TComboBox
      Left = 120
      Top = 76
      Width = 258
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      MaxLength = 25
      TabOrder = 4
    end
    object PartType_ComboBox: TComboBox
      Left = 383
      Top = 100
      Width = 145
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 6
    end
    object Line_ComboBox: TComboBox
      Left = 120
      Top = 100
      Width = 145
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 5
    end
    object Supplier_NUMMIColumnComboBox: TNUMMIColumnComboBox
      Left = 120
      Top = 52
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
      TabOrder = 3
      ReadOnly = False
    end
    object PArtCost_MaskEdit: TcurrEdit
      Left = 120
      Top = 224
      Width = 81
      Height = 21
      TabOrder = 14
      Text = '$0.0000'
      IntDigits = 0
      DecimalDigits = 4
      Prefix = '$'
    end
    object VendorShare_Edit: TEdit
      Left = 383
      Top = 200
      Width = 80
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 5
      TabOrder = 13
      Visible = False
    end
  end
  object PartsStockMaster_DBGrid: TDBGrid
    Left = 20
    Top = 407
    Width = 585
    Height = 134
    DataSource = Parts_DataSource
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgRowSelect, dgConfirmDelete, dgCancelOnExit]
    TabOrder = 2
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    OnKeyUp = PartsStockMaster_DBGridKeyUp
    OnMouseUp = PartsStockMaster_DBGridMouseUp
  end
  object Parts_DataSource: TDataSource
    OnDataChange = Parts_DataSourceDataChange
    Left = 46
    Top = 478
  end
end
