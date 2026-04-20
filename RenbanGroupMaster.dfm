object RenbanGroupMaster_Form: TRenbanGroupMaster_Form
  Left = 548
  Top = 308
  Width = 639
  Height = 383
  BorderIcons = [biSystemMenu, biMinimize]
  Caption = 'Renban Group Master'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PartsStockMaster_Label: TLabel
    Left = 20
    Top = 10
    Width = 179
    Height = 20
    Caption = 'Renban Group Master'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object RenbanGroupMaster_Panel: TPanel
    Left = 21
    Top = 35
    Width = 585
    Height = 93
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 0
    object SizeCode_Label: TLabel
      Left = 10
      Top = 14
      Width = 110
      Height = 20
      AutoSize = False
      Caption = 'Renban Group Code'
      Layout = tlCenter
    end
    object Label1: TLabel
      Left = 10
      Top = 40
      Width = 110
      Height = 20
      AutoSize = False
      Caption = 'Renban Group Count'
      Layout = tlCenter
    end
    object Label4: TLabel
      Left = 218
      Top = 13
      Width = 95
      Height = 20
      AutoSize = False
      Caption = 'Ship Days Override'
      Layout = tlCenter
    end
    object Label12: TLabel
      Left = 218
      Top = 33
      Width = 97
      Height = 20
      AutoSize = False
      Caption = 'Ship Days Daily'
      Layout = tlCenter
    end
    object Label13: TLabel
      Left = 318
      Top = 38
      Width = 9
      Height = 13
      Caption = 'M'
    end
    object Label14: TLabel
      Left = 363
      Top = 38
      Width = 7
      Height = 13
      Caption = 'T'
    end
    object Label15: TLabel
      Left = 408
      Top = 38
      Width = 11
      Height = 13
      Caption = 'W'
    end
    object Label16: TLabel
      Left = 318
      Top = 65
      Width = 13
      Height = 13
      Caption = 'Th'
    end
    object Label17: TLabel
      Left = 365
      Top = 65
      Width = 6
      Height = 13
      Caption = 'F'
    end
    object Label18: TLabel
      Left = 410
      Top = 65
      Width = 7
      Height = 13
      Caption = 'S'
    end
    object RenbanGroupCode_Edit: TEdit
      Tag = 1
      Left = 133
      Top = 14
      Width = 58
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 5
      TabOrder = 0
      OnChange = RenbanGroupCode_EditChange
    end
    object RenbanGroupCount_Edit: TEdit
      Tag = 1
      Left = 133
      Top = 40
      Width = 30
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 3
      TabOrder = 1
    end
    object ShipDays_Edit: TEdit
      Tag = 1
      Left = 334
      Top = 12
      Width = 30
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 3
      TabOrder = 2
    end
    object ShipDaysMonday_MaskEdit: TMaskEdit
      Tag = 1
      Left = 334
      Top = 37
      Width = 28
      Height = 21
      CharCase = ecUpperCase
      EditMask = '99;1; '
      MaxLength = 2
      TabOrder = 3
      Text = '  '
    end
    object ShipDaysTuesday_MaskEdit: TMaskEdit
      Tag = 1
      Left = 374
      Top = 37
      Width = 28
      Height = 21
      CharCase = ecUpperCase
      EditMask = '99;1; '
      MaxLength = 2
      TabOrder = 4
      Text = '  '
    end
    object ShipDaysWednesday_MaskEdit: TMaskEdit
      Tag = 1
      Left = 422
      Top = 37
      Width = 28
      Height = 21
      CharCase = ecUpperCase
      EditMask = '99;1; '
      MaxLength = 2
      TabOrder = 5
      Text = '  '
    end
    object ShipDaysThursday_MaskEdit: TMaskEdit
      Tag = 1
      Left = 334
      Top = 63
      Width = 28
      Height = 21
      CharCase = ecUpperCase
      EditMask = '99;1; '
      MaxLength = 2
      TabOrder = 6
      Text = '  '
    end
    object ShipDaysFriday_MaskEdit: TMaskEdit
      Tag = 1
      Left = 374
      Top = 63
      Width = 28
      Height = 21
      CharCase = ecUpperCase
      EditMask = '99;1; '
      MaxLength = 2
      TabOrder = 7
      Text = '  '
    end
    object ShipDaysSaturday_MaskEdit: TMaskEdit
      Tag = 1
      Left = 422
      Top = 63
      Width = 28
      Height = 21
      CharCase = ecUpperCase
      EditMask = '99;1; '
      MaxLength = 2
      TabOrder = 8
      Text = '  '
    end
  end
  object RenbanGroupMaster_DBGrid: TDBGrid
    Left = 20
    Top = 136
    Width = 585
    Height = 152
    DataSource = Renban_DataSource
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgRowSelect, dgConfirmDelete, dgCancelOnExit]
    TabOrder = 1
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    OnKeyUp = RenbanGroupMaster_DBGridKeyUp
    OnMouseUp = RenbanGroupMaster_DBGridMouseUp
  end
  object ManagementButtons_Panel: TPanel
    Left = 20
    Top = 295
    Width = 585
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
      Left = 392
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
  object Renban_DataSource: TDataSource
    OnDataChange = Renban_DataSourceDataChange
    Left = 136
    Top = 212
  end
end
