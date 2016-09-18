object DailyBuildtotalForm: TDailyBuildtotalForm
  Left = 709
  Top = 293
  Width = 618
  Height = 401
  Caption = 'Daily Build Total Breakdown'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Hist: THistory
    Left = 0
    Top = 0
    Width = 610
    Height = 313
    Align = alTop
    TabOrder = 0
    ShowDate = True
    ShowTime = True
  end
  object DoneButton: TButton
    Left = 268
    Top = 330
    Width = 75
    Height = 25
    Caption = 'Done'
    TabOrder = 1
    Visible = False
    OnClick = DoneButtonClick
  end
  object RunTimer: TTimer
    Enabled = False
    Interval = 500
    OnTimer = RunTimerTimer
    Left = 110
    Top = 284
  end
end
