object ProductionDateSelectDlg: TProductionDateSelectDlg
  Left = 1172
  Top = 843
  BorderStyle = bsDialog
  Caption = 'Select Production Date'
  ClientHeight = 188
  ClientWidth = 426
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Arial'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 14
  object Bevel1: TBevel
    Left = 8
    Top = 8
    Width = 413
    Height = 136
    Shape = bsFrame
  end
  object DateTypeLabel: TLabel
    Left = 21
    Top = 70
    Width = 129
    Height = 14
    Caption = 'Available Production Dates'
  end
  object FRom_Label: TLabel
    Left = 206
    Top = 70
    Width = 24
    Height = 14
    Caption = 'From'
    Visible = False
  end
  object To_Label: TLabel
    Left = 206
    Top = 106
    Width = 12
    Height = 14
    Caption = 'To'
    Visible = False
  end
  object CarTruck_Label: TLabel
    Left = 21
    Top = 31
    Width = 75
    Height = 21
    AutoSize = False
    Caption = 'Line:'
    Layout = tlCenter
  end
  object OKBtn: TButton
    Left = 136
    Top = 154
    Width = 75
    Height = 25
    Caption = 'OK'
    Default = True
    ModalResult = 1
    TabOrder = 0
  end
  object CancelBtn: TButton
    Left = 216
    Top = 154
    Width = 75
    Height = 25
    Cancel = True
    Caption = 'Cancel'
    ModalResult = 2
    TabOrder = 1
    OnClick = CancelBtnClick
  end
  object ProductionDates_ComboBox: TComboBox
    Left = 254
    Top = 65
    Width = 145
    Height = 22
    Style = csDropDownList
    ItemHeight = 14
    TabOrder = 2
    OnChange = ProductionDates_ComboBoxChange
  end
  object ToProductionDate_ComboBox: TComboBox
    Left = 254
    Top = 101
    Width = 145
    Height = 22
    Style = csDropDownList
    ItemHeight = 14
    TabOrder = 3
    Visible = False
    OnChange = ToProductionDate_ComboBoxChange
  end
  object Line_ComboBox: TComboBox
    Left = 254
    Top = 30
    Width = 145
    Height = 22
    ItemHeight = 14
    TabOrder = 4
    OnChange = Line_ComboBoxChange
  end
end
