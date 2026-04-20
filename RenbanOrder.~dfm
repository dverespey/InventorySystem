object GroupRenbanOrder_Form: TGroupRenbanOrder_Form
  Left = 913
  Top = 100
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'Renban Group Order'
  ClientHeight = 464
  ClientWidth = 480
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCloseQuery = FormCloseQuery
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 6
    Top = 14
    Width = 70
    Height = 13
    Caption = 'Renban Group'
  end
  object Label2: TLabel
    Left = 361
    Top = 380
    Width = 47
    Height = 13
    Caption = 'Total Lots'
  end
  object Label3: TLabel
    Left = 132
    Top = 380
    Width = 34
    Height = 13
    Caption = 'Trailers'
  end
  object Label4: TLabel
    Left = 218
    Top = 380
    Width = 89
    Height = 13
    Caption = 'Trailer Pallet Count'
  end
  object RenbanGroups_ComboBox: TComboBox
    Left = 82
    Top = 12
    Width = 145
    Height = 21
    CharCase = ecUpperCase
    ItemHeight = 13
    TabOrder = 0
    Text = 'RENBANGROUPS_COMBOBOX'
    OnChange = RenbanGroups_ComboBoxChange
  end
  object AvailableGrid: TStringGrid
    Left = 5
    Top = 46
    Width = 464
    Height = 325
    ColCount = 8
    FixedCols = 0
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine]
    TabOrder = 1
    ColWidths = (
      47
      51
      78
      68
      52
      39
      32
      82)
  end
  object CreateOrder_Button: TBitBtn
    Left = 282
    Top = 436
    Width = 111
    Height = 25
    Caption = 'Create Renban'
    TabOrder = 2
    OnClick = CreateOrder_ButtonClick
    Glyph.Data = {
      76010000424D7601000000000000760000002800000020000000100000000100
      04000000000000010000120B0000120B00001000000000000000000000000000
      800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
      FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00555555555555
      555555FFFFFFFFFF5F5557777777777505555777777777757F55555555555555
      055555555555FF5575F555555550055030555555555775F7F7F55555550FB000
      005555555575577777F5555550FB0BF0F05555555755755757F555550FBFBF0F
      B05555557F55557557F555550BFBF0FB005555557F55575577F555500FBFBFB0
      305555577F555557F7F5550E0BFBFB003055557575F55577F7F550EEE0BFB0B0
      305557FF575F5757F7F5000EEE0BFBF03055777FF575FFF7F7F50000EEE00000
      30557777FF577777F7F500000E05555BB05577777F75555777F5500000555550
      3055577777555557F7F555000555555999555577755555577755}
    NumGlyphs = 2
  end
  object OKButton: TButton
    Left = 394
    Top = 436
    Width = 75
    Height = 25
    Caption = 'OK'
    TabOrder = 3
    OnClick = OKButtonClick
  end
  object FRSBreakdown_Button: TBitBtn
    Left = 24
    Top = 436
    Width = 145
    Height = 25
    Caption = 'Create FRS Breakdown'
    TabOrder = 4
    OnClick = FRSBreakdown_ButtonClick
    Glyph.Data = {
      76010000424D7601000000000000760000002800000020000000100000000100
      04000000000000010000130B0000130B00001000000000000000000000000000
      800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
      FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
      3333333333333333FF3333333333333003333333333333377F33333333333307
      733333FFF333337773333C003333307733333777FF333777FFFFC0CC03330770
      000077777FF377777777C033C03077FFFFF077FF77F777FFFFF7CC00000F7777
      777077777777777777773CCCCC00000000003777777777777777333330030FFF
      FFF03333F77F7F3FF3F7333C0C030F00F0F03337777F7F77373733C03C030FFF
      FFF03377F77F7F3F333733C03C030F0FFFF03377F7737F733FF733C000330FFF
      0000337777F37F3F7777333CCC330F0F0FF0333777337F737F37333333330FFF
      0F03333333337FFF7F7333333333000000333333333377777733}
    NumGlyphs = 2
  end
  object TotalLots_Edit: TEdit
    Left = 419
    Top = 377
    Width = 50
    Height = 21
    ReadOnly = True
    TabOrder = 5
    Text = 'TotalLots_Edit'
  end
  object Trailers_ComboBox: TComboBox
    Left = 173
    Top = 377
    Width = 39
    Height = 21
    ItemHeight = 13
    TabOrder = 6
    OnChange = Trailers_ComboBoxChange
    Items.Strings = (
      '1'
      '2'
      '3'
      '4'
      '5'
      '6')
  end
  object ClearBreakdown_BitBtn: TBitBtn
    Left = 170
    Top = 436
    Width = 111
    Height = 25
    Caption = 'Clear Breakdown'
    TabOrder = 7
    OnClick = ClearBreakdown_BitBtnClick
    Glyph.Data = {
      76010000424D7601000000000000760000002800000020000000100000000100
      04000000000000010000120B0000120B00001000000000000000000000000000
      800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
      FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00500005000555
      555557777F777555F55500000000555055557777777755F75555005500055055
      555577F5777F57555555005550055555555577FF577F5FF55555500550050055
      5555577FF77577FF555555005050110555555577F757777FF555555505099910
      555555FF75777777FF555005550999910555577F5F77777775F5500505509990
      3055577F75F77777575F55005055090B030555775755777575755555555550B0
      B03055555F555757575755550555550B0B335555755555757555555555555550
      BBB35555F55555575F555550555555550BBB55575555555575F5555555555555
      50BB555555555555575F555555555555550B5555555555555575}
    NumGlyphs = 2
  end
  object TrailerPalletCount_Edit: TEdit
    Left = 317
    Top = 377
    Width = 39
    Height = 21
    MaxLength = 3
    TabOrder = 8
    Text = 'TrailerPalletCount_Edit'
    OnChange = TrailerPalletCount_EditChange
  end
  object TRailerCounts_ListBox: TListBox
    Left = 9
    Top = 375
    Width = 121
    Height = 56
    ItemHeight = 13
    TabOrder = 9
  end
end
