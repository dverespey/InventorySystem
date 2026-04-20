object LogisticsMaster_Form: TLogisticsMaster_Form
  Left = 635
  Top = 484
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'Logistics'
  ClientHeight = 454
  ClientWidth = 629
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
  object SupplierMaster_Label: TLabel
    Left = 10
    Top = 9
    Width = 132
    Height = 20
    Caption = 'Logistics Master'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object LogisticsMaster_Panel: TPanel
    Left = 21
    Top = 34
    Width = 585
    Height = 203
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 0
    object Address_Label: TLabel
      Left = 10
      Top = 35
      Width = 108
      Height = 20
      AutoSize = False
      Caption = 'Address:'
      Layout = tlCenter
    end
    object PhoneNum_Label: TLabel
      Left = 10
      Top = 100
      Width = 108
      Height = 20
      AutoSize = False
      Caption = 'Telephone Number:'
      Layout = tlCenter
    end
    object Person_Label: TLabel
      Left = 10
      Top = 121
      Width = 108
      Height = 20
      AutoSize = False
      Caption = 'Person:'
      Layout = tlCenter
    end
    object SupplierName_Label: TLabel
      Left = 10
      Top = 13
      Width = 75
      Height = 20
      AutoSize = False
      Caption = 'Logistics Name:'
      Layout = tlCenter
    end
    object FaxNum_Label: TLabel
      Left = 326
      Top = 100
      Width = 65
      Height = 20
      AutoSize = False
      Caption = 'Fax Number:'
      Layout = tlCenter
    end
    object Label1: TLabel
      Left = 10
      Top = 164
      Width = 108
      Height = 20
      AutoSize = False
      Caption = 'Breakdown Directory'
      Layout = tlCenter
    end
    object Label2: TLabel
      Left = 10
      Top = 143
      Width = 108
      Height = 20
      AutoSize = False
      Caption = 'Email:'
      Layout = tlCenter
    end
    object Label3: TLabel
      Left = 10
      Top = 56
      Width = 108
      Height = 20
      AutoSize = False
      Caption = 'City'
      Layout = tlCenter
    end
    object Label4: TLabel
      Left = 10
      Top = 80
      Width = 108
      Height = 20
      AutoSize = False
      Caption = 'State'
      Layout = tlCenter
    end
    object Label5: TLabel
      Left = 328
      Top = 80
      Width = 65
      Height = 20
      AutoSize = False
      Caption = 'Zip:'
      Layout = tlCenter
    end
    object Breakdown_SpeedButton: TSpeedButton
      Left = 554
      Top = 164
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
    object Address_Edit: TEdit
      Left = 136
      Top = 35
      Width = 439
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 50
      TabOrder = 1
    end
    object PhoneNum_Edit: TEdit
      Left = 136
      Top = 100
      Width = 150
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 10
      TabOrder = 5
    end
    object Person_Edit: TEdit
      Left = 136
      Top = 121
      Width = 439
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 50
      TabOrder = 7
    end
    object LogisticsName_Edit: TEdit
      Left = 136
      Top = 13
      Width = 240
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 25
      TabOrder = 0
    end
    object FaxNum_Edit: TEdit
      Left = 425
      Top = 100
      Width = 150
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 10
      TabOrder = 6
    end
    object Directory_Edit: TEdit
      Left = 136
      Top = 164
      Width = 417
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 50
      TabOrder = 9
    end
    object Email_Edit: TEdit
      Left = 136
      Top = 143
      Width = 439
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 255
      TabOrder = 8
    end
    object City_Edit: TEdit
      Left = 136
      Top = 56
      Width = 150
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 10
      TabOrder = 2
    end
    object State_Edit: TEdit
      Left = 136
      Top = 78
      Width = 150
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 10
      TabOrder = 3
    end
    object Zip_Edit: TEdit
      Left = 425
      Top = 78
      Width = 150
      Height = 21
      CharCase = ecUpperCase
      MaxLength = 10
      TabOrder = 4
    end
  end
  object ManagementButtons_Panel: TPanel
    Left = 20
    Top = 390
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
  object LogisticsMaster_DBGrid: TDBGrid
    Left = 20
    Top = 244
    Width = 585
    Height = 135
    DataSource = LogisticsMAster_DataSource
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgConfirmDelete, dgCancelOnExit]
    TabOrder = 2
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    OnKeyUp = LogisticsMaster_DBGridKeyUp
    OnMouseUp = LogisticsMaster_DBGridMouseUp
  end
  object LogisticsMAster_DataSource: TDataSource
    OnDataChange = LogisticsMAster_DataSourceDataChange
    Left = 178
    Top = 328
  end
end
