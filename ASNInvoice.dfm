object ASNInvoice_Form: TASNInvoice_Form
  Left = 457
  Top = 21
  BorderStyle = bsSingle
  Caption = 'ASN/Invoice'
  ClientHeight = 716
  ClientWidth = 955
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Arial'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 14
  object CarTruckShip_Label: TLabel
    Left = 20
    Top = 10
    Width = 99
    Height = 20
    Caption = 'ASN/Invoice'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Label1: TLabel
    Left = 20
    Top = 72
    Width = 31
    Height = 14
    Caption = 'Status'
  end
  object Label2: TLabel
    Left = 19
    Top = 44
    Width = 59
    Height = 14
    Caption = 'ASN/Invoice'
  end
  object Label6: TLabel
    Left = 278
    Top = 44
    Width = 26
    Height = 14
    Caption = 'View'
  end
  object SpeedButton1: TSpeedButton
    Left = 632
    Top = 40
    Width = 23
    Height = 22
    Glyph.Data = {
      76010000424D7601000000000000760000002800000020000000100000000100
      04000000000000010000130B0000130B00001000000000000000000000000000
      800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
      FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
      333333333333333333FF33333333333330003FF3FFFFF3333777003000003333
      300077F777773F333777E00BFBFB033333337773333F7F33333FE0BFBF000333
      330077F3337773F33377E0FBFBFBF033330077F3333FF7FFF377E0BFBF000000
      333377F3337777773F3FE0FBFBFBFBFB039977F33FFFFFFF7377E0BF00000000
      339977FF777777773377000BFB03333333337773FF733333333F333000333333
      3300333777333333337733333333333333003333333333333377333333333333
      333333333333333333FF33333333333330003333333333333777333333333333
      3000333333333333377733333333333333333333333333333333}
    NumGlyphs = 2
    OnClick = SpeedButton1Click
  end
  object List_GroupBox: TGroupBox
    Left = 20
    Top = 90
    Width = 909
    Height = 171
    Caption = 'ASN List'
    TabOrder = 0
    object ListDBGrid: TDBGrid
      Left = 2
      Top = 16
      Width = 905
      Height = 153
      Align = alClient
      DataSource = ASNListDataSource
      Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgConfirmDelete, dgCancelOnExit]
      TabOrder = 0
      TitleFont.Charset = ANSI_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -11
      TitleFont.Name = 'Arial'
      TitleFont.Style = []
    end
  end
  object ASNStatus_ComboBox: TComboBox
    Left = 89
    Top = 68
    Width = 145
    Height = 22
    Style = csDropDownList
    ItemHeight = 14
    TabOrder = 1
    OnChange = ASNStatus_ComboBoxChange
    Items.Strings = (
      'ALL'
      'NOT CREATED'
      'SENT'
      'ACCEPTED'
      'REJECTED')
  end
  object Items_GroupBox: TGroupBox
    Left = 20
    Top = 270
    Width = 909
    Height = 171
    Caption = 'ASN Items'
    TabOrder = 2
    object ItemsDBGrid: TDBGrid
      Left = 2
      Top = 16
      Width = 905
      Height = 153
      Align = alClient
      DataSource = ASNItemsDataSource
      Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgConfirmDelete, dgCancelOnExit]
      TabOrder = 0
      TitleFont.Charset = ANSI_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -11
      TitleFont.Name = 'Arial'
      TitleFont.Style = []
    end
  end
  object Buttons_Panel: TPanel
    Left = 62
    Top = 656
    Width = 833
    Height = 50
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 3
    object Insert_Button: TButton
      Left = 6
      Top = 10
      Width = 90
      Height = 30
      Caption = '&Insert'
      TabOrder = 0
      Visible = False
      OnClick = Insert_ButtonClick
    end
    object Update_Button: TButton
      Left = 97
      Top = 10
      Width = 90
      Height = 30
      Caption = '&Update'
      TabOrder = 1
      Visible = False
      OnClick = Update_ButtonClick
    end
    object Search_Button: TButton
      Left = 368
      Top = 10
      Width = 90
      Height = 30
      Caption = '&Search'
      TabOrder = 3
      Visible = False
    end
    object Clear_Button: TButton
      Left = 459
      Top = 10
      Width = 89
      Height = 30
      Caption = 'Cl&ear'
      TabOrder = 4
      Visible = False
      OnClick = Clear_ButtonClick
    end
    object Close_Button: TButton
      Left = 733
      Top = 10
      Width = 90
      Height = 30
      Caption = '&Close'
      ModalResult = 2
      TabOrder = 5
    end
    object Delete_Button: TButton
      Left = 189
      Top = 10
      Width = 90
      Height = 30
      Caption = '&Delete'
      TabOrder = 2
      Visible = False
      OnClick = Delete_ButtonClick
    end
    object DeleteASN_Button: TButton
      Left = 550
      Top = 10
      Width = 90
      Height = 30
      Caption = 'Delete &ASN'
      TabOrder = 6
      Visible = False
      OnClick = DeleteASN_ButtonClick
    end
    object UnsendASN_Button: TButton
      Left = 641
      Top = 10
      Width = 90
      Height = 30
      Caption = '&Unsend ASN'
      TabOrder = 7
      Visible = False
      OnClick = UnsendASN_ButtonClick
    end
    object Cancel_Button: TButton
      Left = 280
      Top = 10
      Width = 90
      Height = 30
      Caption = 'Cancel'
      TabOrder = 8
      Visible = False
      OnClick = Cancel_ButtonClick
    end
    object RecreateFile_Button: TButton
      Left = 549
      Top = 10
      Width = 90
      Height = 30
      Caption = '&ReCreate File'
      TabOrder = 9
      Visible = False
      OnClick = RecreateFile_ButtonClick
    end
  end
  object ASNorInvoice_ComboBox: TComboBox
    Left = 88
    Top = 40
    Width = 145
    Height = 22
    Style = csDropDownList
    ItemHeight = 14
    TabOrder = 4
    OnChange = ASNorInvoice_ComboBoxChange
    Items.Strings = (
      'ASN'
      'INVOICE')
  end
  object ASNItem_Box: TGroupBox
    Left = 24
    Top = 452
    Width = 905
    Height = 179
    Caption = 'ASN Item Values'
    TabOrder = 5
    object Label3: TLabel
      Left = 24
      Top = 30
      Width = 81
      Height = 14
      Caption = 'Manifest Number'
    end
    object Label4: TLabel
      Left = 24
      Top = 65
      Width = 110
      Height = 14
      Caption = 'Assembly Part Number'
    end
    object Label5: TLabel
      Left = 24
      Top = 98
      Width = 17
      Height = 14
      Caption = 'Qty'
    end
    object ManifestNumber_Edit: TEdit
      Left = 142
      Top = 26
      Width = 121
      Height = 22
      ReadOnly = True
      TabOrder = 0
    end
    object AssemblyPartNumber_Edit: TEdit
      Left = 142
      Top = 60
      Width = 121
      Height = 22
      ReadOnly = True
      TabOrder = 1
    end
    object Qty_Edit: TEdit
      Left = 142
      Top = 96
      Width = 121
      Height = 22
      TabOrder = 2
    end
    object AssemblyPartNumber_Combo: TComboBox
      Left = 142
      Top = 59
      Width = 168
      Height = 22
      ItemHeight = 14
      TabOrder = 3
      Text = 'AssemblyPartNumber_Combo'
      Visible = False
      OnChange = AssemblyPartNumber_ComboChange
    end
  end
  object DataView: TComboBox
    Left = 320
    Top = 40
    Width = 145
    Height = 22
    Style = csDropDownList
    ItemHeight = 14
    TabOrder = 6
    OnChange = ASNStatus_ComboBoxChange
    Items.Strings = (
      'ALL'
      'LAST YEAR'
      'LAST MONTH')
  end
  object SearchEdit: TEdit
    Left = 510
    Top = 40
    Width = 121
    Height = 22
    TabOrder = 7
  end
  object ASNListDataSource: TDataSource
    DataSet = ASNList_DataSet
    OnDataChange = ASNListDataSourceDataChange
    Left = 778
    Top = 122
  end
  object ASNList_DataSet: TADODataSet
    Connection = Data_Module.Inv_Connection
    CommandText = 'SELECT_ASNList'
    CommandType = cmdStoredProc
    Parameters = <
      item
        Name = '@RETURN_VALUE'
        DataType = ftInteger
        Direction = pdReturnValue
        Precision = 10
        Value = Null
      end
      item
        Name = '@List'
        Attributes = [paNullable]
        DataType = ftString
        Size = 1
        Value = Null
      end
      item
        Name = '@Range'
        DataType = ftInteger
        Value = Null
      end>
    Left = 808
    Top = 122
  end
  object ASNItems_DataSet: TADODataSet
    Connection = Data_Module.Inv_Connection
    CommandText = 'SELECT_ASNItems'
    CommandType = cmdStoredProc
    Parameters = <>
    Left = 872
    Top = 290
  end
  object ASNItemsDataSource: TDataSource
    DataSet = ASNItems_DataSet
    OnDataChange = ASNItemsDataSourceDataChange
    Left = 838
    Top = 290
  end
  object Invoice_DataSource: TDataSource
    DataSet = InvoiceList_DataSet
    OnDataChange = Invoice_DataSourceDataChange
    Left = 778
    Top = 162
  end
  object InvoiceList_DataSet: TADODataSet
    Connection = Data_Module.Inv_Connection
    CommandText = 'SELECT_INVOICEList'
    CommandType = cmdStoredProc
    Parameters = <
      item
        Name = '@RETURN_VALUE'
        DataType = ftInteger
        Direction = pdReturnValue
        Precision = 10
        Value = Null
      end
      item
        Name = '@List'
        Attributes = [paNullable]
        DataType = ftString
        Size = 1
        Value = Null
      end
      item
        Name = '@Range'
        DataType = ftInteger
        Value = Null
      end>
    Left = 812
    Top = 162
  end
  object INVOICEItems_DataSet: TADODataSet
    Connection = Data_Module.Inv_Connection
    CommandText = 'SELECT_INVOICEItems'
    CommandType = cmdStoredProc
    Parameters = <>
    Left = 872
    Top = 324
  end
  object INVOICEItemsDataSource: TDataSource
    DataSet = INVOICEItems_DataSet
    Left = 838
    Top = 324
  end
  object ASNItem_Command: TADOCommand
    CommandType = cmdStoredProc
    Connection = Data_Module.Inv_Connection
    Parameters = <>
    Left = 872
    Top = 356
  end
  object AssyManifest_DataSet: TADODataSet
    Connection = Data_Module.Inv_Connection
    CommandText = 'SELECT_ManifestCost'
    CommandType = cmdStoredProc
    Parameters = <>
    Left = 732
    Top = 354
  end
end
