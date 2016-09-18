object ForeUpBreakDown_Form: TForeUpBreakDown_Form
  Left = 302
  Top = 272
  Width = 496
  Height = 178
  Caption = 'Forecast Information Upload & Breakdown'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object ForeUpBreakDown_Label: TLabel
    Left = 24
    Top = 26
    Width = 295
    Height = 20
    Caption = 'Forecast Info. Upload && Break Down'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object FileName_Label: TLabel
    Left = 16
    Top = 82
    Width = 50
    Height = 13
    Caption = 'File Name:'
  end
  object FileName_Edit: TEdit
    Left = 74
    Top = 82
    Width = 399
    Height = 21
    TabOrder = 0
    Text = '[Type file path and name here or click Browse]'
    OnChange = FileName_EditChange
  end
  object Browse_Button: TButton
    Left = 14
    Top = 114
    Width = 153
    Height = 23
    Caption = '&Browse...'
    TabOrder = 1
    OnClick = Browse_ButtonClick
  end
  object Start_Button: TButton
    Left = 168
    Top = 114
    Width = 153
    Height = 23
    Caption = '&Start'
    TabOrder = 2
    OnClick = Start_ButtonClick
  end
  object Close_Button: TButton
    Left = 322
    Top = 114
    Width = 153
    Height = 23
    Caption = '&Close'
    ModalResult = 2
    TabOrder = 3
  end
  object ForecastFilleNameDialog: TOpenDialog
    DefaultExt = '.txt'
    Left = 336
    Top = 8
  end
end
