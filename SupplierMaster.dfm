object SupplierMaster_Form: TSupplierMaster_Form
  Left = 713
  Top = 307
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsDialog
  Caption = 'Supplier Master'
  ClientHeight = 492
  ClientWidth = 626
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
  object SupplierMaster_Label: TLabel
    Left = 10
    Top = 9
    Width = 127
    Height = 20
    Caption = 'Supplier Master'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object ManagementButtons_Panel: TPanel
    Left = 20
    Top = 437
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
  object SupplierMaster_Panel: TPanel
    Left = 21
    Top = 34
    Width = 585
    Height = 293
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 0
    object SupplierCode_Label: TLabel
      Left = 10
      Top = 13
      Width = 108
      Height = 20
      AutoSize = False
      Caption = 'Supplier Code:'
      Layout = tlCenter
    end
    object Address_Label: TLabel
      Left = 10
      Top = 57
      Width = 108
      Height = 20
      AutoSize = False
      Caption = 'Address:'
      Layout = tlCenter
    end
    object PhoneNum_Label: TLabel
      Left = 10
      Top = 123
      Width = 108
      Height = 20
      AutoSize = False
      Caption = 'Telephone Number:'
      Layout = tlCenter
    end
    object Person_Label: TLabel
      Left = 10
      Top = 145
      Width = 108
      Height = 20
      AutoSize = False
      Caption = 'Person:'
      Layout = tlCenter
    end
    object SupplierName_Label: TLabel
      Left = 226
      Top = 13
      Width = 75
      Height = 20
      AutoSize = False
      Caption = 'Supplier Name:'
      Layout = tlCenter
    end
    object FaxNum_Label: TLabel
      Left = 326
      Top = 123
      Width = 65
      Height = 20
      AutoSize = False
      Caption = 'Fax Number:'
      Layout = tlCenter
    end
    object Label1: TLabel
      Left = 10
      Top = 189
      Width = 108
      Height = 20
      AutoSize = False
      Caption = 'Breakdown Directory'
      Layout = tlCenter
    end
    object Label2: TLabel
      Left = 10
      Top = 167
      Width = 108
      Height = 20
      AutoSize = False
      Caption = 'Email:'
      Layout = tlCenter
    end
    object Label3: TLabel
      Left = 10
      Top = 79
      Width = 108
      Height = 20
      AutoSize = False
      Caption = 'City'
      Layout = tlCenter
    end
    object Label4: TLabel
      Left = 10
      Top = 101
      Width = 108
      Height = 20
      AutoSize = False
      Caption = 'State'
      Layout = tlCenter
    end
    object Label5: TLabel
      Left = 328
      Top = 101
      Width = 65
      Height = 20
      AutoSize = False
      Caption = 'Zip:'
      Layout = tlCenter
    end
    object Breakdown_SpeedButton: TSpeedButton
      Left = 552
      Top = 189
      Width = 23
      Height = 22
      Glyph.Data = {
        76010000424D7601000000000000760000002800000020000000100000000100
        04000000000000010000120B0000120B00001000000000000000000000000000
        800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00303333333333
        333337F3333333333333303333333333333337F33FFFFF3FF3FF303300000300
        300337FF77777F77377330000BBB0333333337777F337F33333330330BB00333
        333337F373F773333333303330033333333337F3377333333333303333333333
        333337F33FFFFF3FF3FF303300000300300337FF77777F77377330000BBB0333
        333337777F337F33333330330BB00333333337F373F773333333303330033333
        333337F3377333333333303333333333333337FFFF3FF3FFF333000003003000
        333377777F77377733330BBB0333333333337F337F33333333330BB003333333
        333373F773333333333330033333333333333773333333333333}
      NumGlyphs = 2
      OnClick = Breakdown_SpeedButtonClick
    end
    object Label6: TLabel
      Left = 10
      Top = 35
      Width = 108
      Height = 20
      AutoSize = False
      Caption = 'Logistics:'
      Layout = tlCenter
    end
    object Label7: TLabel
      Left = 10
      Top = 217
      Width = 108
      Height = 20
      AutoSize = False
      Caption = 'File Type'
      Layout = tlCenter
    end
    object Label8: TLabel
      Left = 11
      Top = 246
      Width = 91
      Height = 13
      Caption = 'Create Order Sheet'
    end
    object Label9: TLabel
      Left = 11
      Top = 271
      Width = 93
      Height = 13
      Caption = 'Inventory Add Point'
    end
    object SupplierCode_Edit: TEdit
      Tag = 1
      Left = 136
      Top = 13
      Width = 57
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 5
      TabOrder = 0
    end
    object Address_Edit: TEdit
      Left = 136
      Top = 57
      Width = 439
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 50
      TabOrder = 3
    end
    object PhoneNum_Edit: TEdit
      Left = 136
      Top = 123
      Width = 150
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 10
      TabOrder = 7
    end
    object Person_Edit: TEdit
      Left = 136
      Top = 145
      Width = 439
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 50
      TabOrder = 9
    end
    object SupplierName_Edit: TEdit
      Left = 335
      Top = 13
      Width = 240
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 25
      TabOrder = 1
    end
    object FaxNum_Edit: TEdit
      Left = 425
      Top = 123
      Width = 150
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 10
      TabOrder = 8
    end
    object Directory_Edit: TEdit
      Left = 136
      Top = 189
      Width = 415
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 50
      TabOrder = 11
    end
    object Email_Edit: TEdit
      Left = 136
      Top = 167
      Width = 439
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 255
      TabOrder = 10
    end
    object City_Edit: TEdit
      Left = 136
      Top = 79
      Width = 150
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 10
      TabOrder = 4
    end
    object State_Edit: TEdit
      Left = 136
      Top = 101
      Width = 150
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 10
      TabOrder = 5
    end
    object Zip_Edit: TEdit
      Left = 425
      Top = 101
      Width = 150
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 10
      TabOrder = 6
    end
    object Logistics_ComboBox: TComboBox
      Left = 136
      Top = 35
      Width = 295
      Height = 21
      ItemHeight = 13
      MaxLength = 25
      TabOrder = 2
    end
    object OutputFileType_RadioGroup: TRadioGroup
      Left = 136
      Top = 208
      Width = 243
      Height = 31
      BiDiMode = bdRightToLeftNoAlign
      Columns = 3
      ItemIndex = 0
      Items.Strings = (
        'Text'
        'Excel'
        'Both')
      ParentBiDiMode = False
      TabOrder = 12
    end
    object OrderFileTimestamp_CheckBox: TCheckBox
      Left = 395
      Top = 219
      Width = 142
      Height = 17
      Alignment = taLeftJustify
      Caption = 'Order File Timestamp'
      TabOrder = 13
    end
    object SiteNumberinOrder_CheckBox: TCheckBox
      Left = 395
      Top = 245
      Width = 142
      Height = 17
      Alignment = taLeftJustify
      Caption = 'Site Number in Order'
      TabOrder = 15
    end
    object CreateOrderSheet_ComboBox: TComboBox
      Left = 136
      Top = 241
      Width = 110
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 14
    end
    object InventoryAddPoint_ComboBox: TComboBox
      Left = 136
      Top = 266
      Width = 110
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 16
    end
  end
  object SupplierMaster_DBGrid: TDBGrid
    Left = 20
    Top = 330
    Width = 585
    Height = 103
    DataSource = Supplier_DataSource
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgRowSelect, dgConfirmDelete, dgCancelOnExit]
    TabOrder = 2
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    OnKeyUp = SupplierMaster_DBGridKeyUp
    OnMouseUp = SupplierMaster_DBGridMouseUp
  end
  object Supplier_DataSource: TDataSource
    OnDataChange = Supplier_DataSourceDataChange
    Left = 144
    Top = 338
  end
end
