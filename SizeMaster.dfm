object SizeMaster_Form: TSizeMaster_Form
  Left = 631
  Top = 147
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsDialog
  Caption = 'Size Master'
  ClientHeight = 413
  ClientWidth = 632
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
  object SizeMaster_Label: TLabel
    Left = 10
    Top = 10
    Width = 96
    Height = 20
    Caption = 'Size Master'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object ManagementButtons_Panel: TPanel
    Left = 20
    Top = 354
    Width = 585
    Height = 51
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
  object SizeMaster_Panel: TPanel
    Left = 21
    Top = 33
    Width = 585
    Height = 79
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 0
    object SizeCode_Label: TLabel
      Left = 10
      Top = 12
      Width = 55
      Height = 20
      AutoSize = False
      Caption = 'Size Code:'
      Layout = tlCenter
    end
    object DailyUsage_Label: TLabel
      Left = 10
      Top = 42
      Width = 62
      Height = 20
      AutoSize = False
      Caption = 'Daily Usage:'
      Layout = tlCenter
    end
    object SizeName_Label: TLabel
      Left = 259
      Top = 12
      Width = 59
      Height = 20
      AutoSize = False
      Caption = 'Size Name:'
      Layout = tlCenter
    end
    object SafetyDays_Label: TLabel
      Left = 259
      Top = 42
      Width = 62
      Height = 20
      AutoSize = False
      Caption = 'Safety Days:'
      Layout = tlCenter
    end
    object SizeCode_Edit: TEdit
      Tag = 1
      Left = 91
      Top = 10
      Width = 150
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 6
      TabOrder = 0
      OnChange = SizeCode_EditChange
    end
    object SizeName_Edit: TEdit
      Left = 335
      Top = 10
      Width = 240
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 50
      TabOrder = 1
    end
    object DailyUsage_MaskEdit: TMaskEdit
      Left = 91
      Top = 41
      Width = 150
      Height = 21
      CharCase = ecUpperCase
      EditMask = '99999;1; '
      MaxLength = 5
      TabOrder = 2
      Text = '0    '
      OnChange = TextChange
    end
    object SafetyDays_MaskEdit: TMaskEdit
      Left = 335
      Top = 41
      Width = 150
      Height = 21
      CharCase = ecUpperCase
      EditMask = '99999;1; '
      MaxLength = 5
      TabOrder = 3
      Text = '0    '
      OnChange = TextChange
    end
  end
  object SizeMaster_DBGrid: TDBGrid
    Left = 20
    Top = 117
    Width = 585
    Height = 230
    DataSource = Size_DataSource
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgRowSelect, dgConfirmDelete, dgCancelOnExit]
    TabOrder = 2
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    OnKeyUp = SizeMaster_DBGridKeyUp
    OnMouseUp = SizeMaster_DBGridMouseUp
  end
  object Size_DataSource: TDataSource
    OnDataChange = Size_DataSourceDataChange
    Left = 176
    Top = 300
  end
end
