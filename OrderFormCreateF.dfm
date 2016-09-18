object OrderFormCreate_Form: TOrderFormCreate_Form
  Left = 239
  Top = 206
  BorderStyle = bsSingle
  Caption = 'Order Form Create'
  ClientHeight = 370
  ClientWidth = 555
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnActivate = FormActivate
  PixelsPerInch = 96
  TextHeight = 13
  object OK_Button: TButton
    Left = 240
    Top = 342
    Width = 75
    Height = 25
    Caption = 'OK'
    Default = True
    TabOrder = 0
    Visible = False
    OnClick = OK_ButtonClick
  end
  object Hist: THistory
    Left = 0
    Top = 0
    Width = 555
    Height = 333
    Align = alTop
    TabOrder = 1
    ShowDate = True
    ShowTime = True
  end
end
