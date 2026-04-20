object ForecastBreakdown_Form: TForecastBreakdown_Form
  Left = 138
  Top = 209
  BorderIcons = []
  BorderStyle = bsSingle
  Caption = 'Forecast Breakdown'
  ClientHeight = 345
  ClientWidth = 661
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object Hist: THistory
    Left = 0
    Top = 0
    Width = 661
    Height = 311
    Align = alTop
    TabOrder = 0
    ShowDate = True
    ShowTime = True
  end
  object OKButton: TButton
    Left = 323
    Top = 316
    Width = 75
    Height = 25
    Caption = 'OK'
    TabOrder = 1
    OnClick = OKButtonClick
  end
end
