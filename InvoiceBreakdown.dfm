object InvoiceBreakdown_Form: TInvoiceBreakdown_Form
  Left = 318
  Top = 326
  BorderStyle = bsSingle
  Caption = 'Invoice Breakdown'
  ClientHeight = 330
  ClientWidth = 643
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Hist: THistory
    Left = 0
    Top = 0
    Width = 643
    Height = 287
    Align = alTop
    TabOrder = 0
    ShowDate = True
    ShowTime = True
  end
  object OK_Button: TButton
    Left = 284
    Top = 300
    Width = 75
    Height = 25
    Caption = 'OK'
    TabOrder = 1
    Visible = False
    OnClick = OK_ButtonClick
  end
end
