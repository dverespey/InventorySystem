object RecRej_Form: TRecRej_Form
  Left = 804
  Top = 162
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsDialog
  Caption = 'Receiving / Productioin Reject'
  ClientHeight = 563
  ClientWidth = 637
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
    Width = 237
    Height = 20
    Caption = 'Receiving / Production Reject'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object RecProdRej_DBGrid: TDBGrid
    Left = 10
    Top = 320
    Width = 615
    Height = 177
    DataSource = RecRej_DataSource
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgConfirmDelete, dgCancelOnExit]
    TabOrder = 0
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    OnKeyUp = RecProdRej_DBGridKeyUp
    OnMouseUp = RecProdRej_DBGridMouseUp
  end
  object ManagementButtons_Panel: TPanel
    Left = 10
    Top = 505
    Width = 615
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
      TabOrder = 2
      OnClick = Search_ButtonClick
    end
    object Clear_Button: TButton
      Left = 412
      Top = 10
      Width = 90
      Height = 30
      Caption = 'Cl&ear'
      TabOrder = 3
      OnClick = Clear_ButtonClick
    end
    object Close_Button: TButton
      Left = 518
      Top = 10
      Width = 90
      Height = 30
      Caption = '&Close'
      ModalResult = 2
      TabOrder = 4
    end
    object Delete_Button: TButton
      Left = 210
      Top = 10
      Width = 90
      Height = 30
      Caption = '&Delete'
      TabOrder = 5
      OnClick = Delete_ButtonClick
    end
  end
  object RecProdRej_Panel: TPanel
    Left = 10
    Top = 34
    Width = 615
    Height = 277
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 2
    object Reason_Label: TLabel
      Left = 40
      Top = 232
      Width = 40
      Height = 13
      Caption = 'Reason:'
      Layout = tlCenter
    end
    object RejQty_Label: TLabel
      Left = 40
      Top = 179
      Width = 76
      Height = 13
      Caption = 'Reject Quantity:'
      Layout = tlCenter
    end
    object PartsCode_Label: TLabel
      Left = 40
      Top = 148
      Width = 62
      Height = 13
      Caption = 'Part Number:'
      Layout = tlCenter
    end
    object SupplierCode_Label: TLabel
      Left = 40
      Top = 117
      Width = 41
      Height = 13
      Caption = 'Supplier:'
      Layout = tlCenter
    end
    object DiscernDiv_Label: TLabel
      Left = 40
      Top = 37
      Width = 102
      Height = 13
      Caption = 'Discernment Division:'
      Layout = tlCenter
    end
    object Date_Label: TLabel
      Left = 40
      Top = 12
      Width = 26
      Height = 13
      Caption = 'Date:'
      Layout = tlCenter
    end
    object Reason_Memo: TMemo
      Left = 170
      Top = 211
      Width = 420
      Height = 60
      Lines.Strings = (
        'Reason_Memo')
      TabOrder = 5
    end
    object PartsCode_ComboBox: TComboBox
      Left = 170
      Top = 146
      Width = 256
      Height = 21
      ItemHeight = 13
      TabOrder = 3
      Text = '[SELECT A PART]'
      OnChange = PartsCode_ComboBoxChange
    end
    object DiscnDiv_RadioGroup: TRadioGroup
      Left = 170
      Top = 34
      Width = 187
      Height = 70
      ItemIndex = 0
      Items.Strings = (
        'Receiving Reject'
        'GSA Production Reject'
        'NUMMI Production Reject')
      TabOrder = 1
    end
    object RejQty_MaskEdit: TMaskEdit
      Left = 170
      Top = 179
      Width = 100
      Height = 21
      EditMask = '#######;1; '
      MaxLength = 7
      TabOrder = 4
      Text = '0      '
      OnChange = TextChange
      OnExit = MaskEditExit
    end
    object Edit_DateTimePicker: TDateTimePicker
      Left = 170
      Top = 12
      Width = 121
      Height = 21
      Date = 37540.580963750000000000
      Format = 'yyyy/MM/dd'
      Time = 37540.580963750000000000
      DateFormat = dfLong
      TabOrder = 0
    end
    object Supplier_NUMMIColumnComboBox: TNUMMIColumnComboBox
      Left = 169
      Top = 112
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
      TabOrder = 2
      OnChange = Supplier_NUMMIColumnComboBoxChange
      ReadOnly = False
    end
  end
  object RecRej_DataSource: TDataSource
    OnDataChange = RecRej_DataSourceDataChange
    Left = 100
    Top = 400
  end
end
