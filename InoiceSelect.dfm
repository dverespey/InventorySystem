object InvoiceSelect_Form: TInvoiceSelect_Form
  Left = 594
  Top = 367
  Width = 630
  Height = 355
  Caption = 'Invoice Select'
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Arial'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 14
  object CarTruckShip_Label: TLabel
    Left = 20
    Top = 10
    Width = 176
    Height = 20
    Caption = 'Select ASN to Invoice'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Button_Panel: TPanel
    Left = 198
    Top = 257
    Width = 230
    Height = 50
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 0
    object CreateASN_Button: TButton
      Left = 8
      Top = 10
      Width = 98
      Height = 30
      Caption = '&Create INVOICE'
      TabOrder = 0
    end
    object Close_Button: TButton
      Left = 125
      Top = 10
      Width = 98
      Height = 30
      Caption = '&Close'
      ModalResult = 2
      TabOrder = 1
    end
  end
  object GroupBox1: TGroupBox
    Left = 20
    Top = 36
    Width = 585
    Height = 207
    Caption = 'ASN'#39's Available to Invoice'
    TabOrder = 1
    object DBGrid1: TDBGrid
      Left = 2
      Top = 16
      Width = 581
      Height = 189
      Align = alClient
      TabOrder = 0
      TitleFont.Charset = ANSI_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -11
      TitleFont.Name = 'Arial'
      TitleFont.Style = []
    end
  end
  object DataSource1: TDataSource
    Left = 488
    Top = 274
  end
end
