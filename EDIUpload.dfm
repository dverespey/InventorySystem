object EDIUpload_Form: TEDIUpload_Form
  Left = 110
  Top = 121
  Width = 983
  Height = 540
  Caption = 'EDIUpload'
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
    Width = 975
    Height = 446
    Align = alTop
    TabOrder = 0
    ShowDate = True
    ShowTime = True
  end
  object OKButton: TButton
    Left = 417
    Top = 472
    Width = 75
    Height = 25
    Caption = 'OK'
    TabOrder = 1
    Visible = False
    OnClick = OKButtonClick
  end
  object CopyFile: TCopyFile
    Progress = True
    ShowFileNames = False
    Movefile = True
    TransferTimeDate = True
    Left = 552
    Top = 468
  end
end
