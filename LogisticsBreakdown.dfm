object LogisticsBreakdown_Form: TLogisticsBreakdown_Form
  Left = 297
  Top = 124
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'Logistics Breakdown'
  ClientHeight = 340
  ClientWidth = 644
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
  object OK_Button: TButton
    Left = 284
    Top = 300
    Width = 75
    Height = 25
    Caption = 'OK'
    TabOrder = 0
    Visible = False
    OnClick = OK_ButtonClick
  end
  object Hist: THistory
    Left = 0
    Top = 0
    Width = 644
    Height = 287
    Align = alTop
    TabOrder = 1
    ShowDate = True
    ShowTime = True
  end
  object CopyFile: TCopyFile
    Progress = False
    ShowFileNames = False
    Movefile = True
    TransferTimeDate = True
    Left = 84
    Top = 304
  end
end
