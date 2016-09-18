object BCRatioMaster_Form: TBCRatioMaster_Form
  Left = 766
  Top = 242
  BorderStyle = bsSingle
  Caption = 'Broadcast Ratio Master'
  ClientHeight = 570
  ClientWidth = 656
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Arial'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 14
  object AssyRatioMaster_Label: TLabel
    Left = 10
    Top = 10
    Width = 191
    Height = 20
    Caption = 'Broadcast Ratio Master'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object AssyRatioMaster_Panel: TPanel
    Left = 20
    Top = 36
    Width = 619
    Height = 484
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 0
    object BC_Panel: TPanel
      Left = 4
      Top = 3
      Width = 608
      Height = 246
      BevelInner = bvLowered
      TabOrder = 0
      object AssyCode_Label: TLabel
        Left = 346
        Top = 47
        Width = 95
        Height = 20
        AutoSize = False
        Caption = 'ASSY Code:'
        Layout = tlCenter
      end
      object BroadcastCode_Label: TLabel
        Left = 10
        Top = 9
        Width = 95
        Height = 20
        AutoSize = False
        Caption = 'Broadcast Code:'
        Layout = tlCenter
      end
      object Label7: TLabel
        Left = 346
        Top = 76
        Width = 95
        Height = 20
        AutoSize = False
        Caption = 'Assy Ratio:'
        Layout = tlCenter
      end
      object Label8: TLabel
        Left = 346
        Top = 105
        Width = 95
        Height = 20
        AutoSize = False
        Caption = 'Assy Quantity:'
        Layout = tlCenter
      end
      object BroadcastCode_Edit: TEdit
        Left = 133
        Top = 9
        Width = 35
        Height = 22
        CharCase = ecUpperCase
        MaxLength = 2
        TabOrder = 0
      end
      object AssyCode_ComboBox: TComboBox
        Left = 456
        Top = 46
        Width = 136
        Height = 22
        Style = csDropDownList
        CharCase = ecUpperCase
        ItemHeight = 14
        TabOrder = 1
      end
      object BCRatioMaster_DBGrid: TDBGrid
        Left = 12
        Top = 45
        Width = 323
        Height = 140
        DataSource = BCRatio_DataSource
        Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgRowSelect, dgConfirmDelete, dgCancelOnExit]
        TabOrder = 2
        TitleFont.Charset = ANSI_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Arial'
        TitleFont.Style = []
      end
      object MaskEdit1: TMaskEdit
        Left = 456
        Top = 72
        Width = 126
        Height = 22
        CharCase = ecUpperCase
        EditMask = '999;1; '
        MaxLength = 3
        TabOrder = 3
        Text = '0  '
      end
      object RadioGroup1: TRadioGroup
        Left = 455
        Top = 97
        Width = 135
        Height = 30
        Columns = 3
        ItemIndex = 0
        Items.Strings = (
          '1'
          '4'
          '5')
        TabOrder = 4
      end
      object Buttons_Panel: TPanel
        Left = 13
        Top = 189
        Width = 585
        Height = 50
        BevelInner = bvRaised
        BevelOuter = bvLowered
        TabOrder = 5
        object Insert_Button: TButton
          Left = 10
          Top = 10
          Width = 90
          Height = 30
          Caption = '&Insert'
          TabOrder = 0
        end
        object Update_Button: TButton
          Left = 105
          Top = 10
          Width = 90
          Height = 30
          Caption = '&Update'
          TabOrder = 1
        end
        object Search_Button: TButton
          Left = 295
          Top = 10
          Width = 90
          Height = 30
          Caption = '&Search'
          TabOrder = 3
        end
        object Clear_Button: TButton
          Left = 392
          Top = 10
          Width = 90
          Height = 30
          Caption = 'Cl&ear'
          TabOrder = 4
        end
        object Close_Button: TButton
          Left = 485
          Top = 10
          Width = 90
          Height = 30
          Caption = '&Close'
          ModalResult = 2
          TabOrder = 5
          Visible = False
        end
        object Delete_Button: TButton
          Left = 200
          Top = 10
          Width = 90
          Height = 30
          Caption = '&Delete'
          TabOrder = 2
        end
      end
    end
    object PN_Panel: TPanel
      Left = 4
      Top = 252
      Width = 609
      Height = 226
      BevelInner = bvLowered
      TabOrder = 1
      object Label9: TLabel
        Left = 346
        Top = 8
        Width = 95
        Height = 20
        AutoSize = False
        Caption = 'Part Code:'
        Layout = tlCenter
      end
      object Label1: TLabel
        Left = 346
        Top = 33
        Width = 95
        Height = 20
        AutoSize = False
        Caption = 'Part Description:'
        Layout = tlCenter
      end
      object Panel3: TPanel
        Left = 13
        Top = 171
        Width = 585
        Height = 51
        BevelInner = bvRaised
        BevelOuter = bvLowered
        TabOrder = 0
        object Button1: TButton
          Left = 10
          Top = 10
          Width = 90
          Height = 30
          Caption = '&Insert'
          TabOrder = 0
        end
        object Button2: TButton
          Left = 105
          Top = 10
          Width = 90
          Height = 30
          Caption = '&Update'
          TabOrder = 1
          Visible = False
        end
        object Button3: TButton
          Left = 295
          Top = 10
          Width = 90
          Height = 30
          Caption = '&Search'
          TabOrder = 3
        end
        object Button4: TButton
          Left = 392
          Top = 10
          Width = 90
          Height = 30
          Caption = 'Cl&ear'
          TabOrder = 4
        end
        object Button5: TButton
          Left = 485
          Top = 10
          Width = 90
          Height = 30
          Caption = '&Close'
          ModalResult = 2
          TabOrder = 5
          Visible = False
        end
        object Button6: TButton
          Left = 200
          Top = 10
          Width = 90
          Height = 30
          Caption = '&Delete'
          TabOrder = 2
        end
      end
      object PNRatio_DBGrid: TDBGrid
        Left = 12
        Top = 5
        Width = 323
        Height = 159
        DataSource = PNRatio_DataSource
        Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgRowSelect, dgConfirmDelete, dgCancelOnExit]
        TabOrder = 1
        TitleFont.Charset = ANSI_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Arial'
        TitleFont.Style = []
      end
      object PartNumber_ComboBox: TComboBox
        Left = 456
        Top = 7
        Width = 136
        Height = 22
        Style = csDropDownList
        CharCase = ecUpperCase
        ItemHeight = 14
        TabOrder = 2
      end
      object MaskEdit2: TMaskEdit
        Left = 456
        Top = 32
        Width = 144
        Height = 22
        CharCase = ecUpperCase
        EditMask = '999;1; '
        MaxLength = 3
        TabOrder = 3
        Text = '0  '
      end
    end
  end
  object Button7: TButton
    Left = 523
    Top = 532
    Width = 90
    Height = 30
    Caption = '&Close'
    ModalResult = 2
    TabOrder = 1
  end
  object BCRatio_DataSource: TDataSource
    Left = 294
    Top = 157
  end
  object PNRatio_DataSource: TDataSource
    Left = 316
    Top = 338
  end
end
