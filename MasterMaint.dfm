object MasterMaint_Form: TMasterMaint_Form
  Left = 493
  Top = 167
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsDialog
  Caption = 'Master Data Maintenance Menu'
  ClientHeight = 323
  ClientWidth = 370
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  PixelsPerInch = 96
  TextHeight = 13
  object MastDateMaintMenu_Label: TLabel
    Left = 12
    Top = 16
    Width = 257
    Height = 20
    Caption = 'Master Data Maintenance Menu'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object SupMaster_Button: TButton
    Left = 50
    Top = 60
    Width = 121
    Height = 33
    Caption = '&Supplier Master'
    TabOrder = 0
    OnClick = SupMaster_ButtonClick
  end
  object SizeMaster_Button: TButton
    Left = 200
    Top = 117
    Width = 121
    Height = 33
    Caption = 'Si&ze Master'
    TabOrder = 1
    OnClick = SizeMaster_ButtonClick
  end
  object AssyRatioMaster_Button: TButton
    Left = 200
    Top = 174
    Width = 121
    Height = 33
    Caption = '&ASSY / Ratio Master'
    TabOrder = 2
    OnClick = AssyRatioMaster_ButtonClick
  end
  object PartsStockMaster_Button: TButton
    Left = 200
    Top = 60
    Width = 121
    Height = 33
    Caption = '&Parts / Stock Master'
    TabOrder = 3
    OnClick = PartsStockMaster_ButtonClick
  end
  object Close_Button: TButton
    Left = 50
    Top = 278
    Width = 271
    Height = 33
    Caption = '&Close'
    ModalResult = 2
    TabOrder = 4
  end
  object ForecastDetail_Button: TButton
    Left = 50
    Top = 174
    Width = 121
    Height = 33
    Caption = '&Assembly Detail'
    TabOrder = 5
    OnClick = ForecastDetail_ButtonClick
  end
  object Button1: TButton
    Left = 50
    Top = 117
    Width = 121
    Height = 33
    Caption = '&Logistics Master'
    TabOrder = 6
    OnClick = Button1Click
  end
  object RenbanGroupMaster_Button: TButton
    Left = 50
    Top = 227
    Width = 121
    Height = 33
    Caption = '&Renban Group Master'
    TabOrder = 7
    OnClick = RenbanGroupMaster_ButtonClick
  end
  object MonthlyPO_Button: TButton
    Left = 199
    Top = 227
    Width = 121
    Height = 33
    Caption = '&Monthly PO'
    TabOrder = 8
    OnClick = MonthlyPO_ButtonClick
  end
  object ASNINVOIVE_Button: TButton
    Left = 200
    Top = 174
    Width = 121
    Height = 33
    Caption = '&ASN/Invoice '
    TabOrder = 9
    Visible = False
    OnClick = ASNINVOIVE_ButtonClick
  end
end
