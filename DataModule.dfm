object Data_Module: TData_Module
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 572
  Top = 22
  Height = 886
  Width = 697
  object Inv_Connection: TADOConnection
    ConnectionString = 
      'Provider=SQLOLEDB.1;Password=noway;Persist Security Info=True;Us' +
      'er ID=Inventory;Initial Catalog=InventoryH;Data Source=dverespe-' +
      'cd6d24;Use Procedure for Prepare=1;Auto Translate=True;Packet Si' +
      'ze=4096;Workstation ID=FAILPROOF1;Use Encryption for Data=False;' +
      'Tag with column collation when possible=False'
    LoginPrompt = False
    Mode = cmRead
    Provider = 'SQLOLEDB.1'
    Left = 34
    Top = 22
  end
  object Inv_DataSet: TADODataSet
    Connection = Inv_Connection
    CursorType = ctStatic
    CommandText = 'SELECT_PartsStockInfo;1'
    CommandType = cmdStoredProc
    Parameters = <>
    Left = 118
    Top = 18
  end
  object Act_Connection: TADOConnection
    ConnectionString = 
      'Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security In' +
      'fo=False;User ID=Inventory;Initial Catalog=Activity;Data Source=' +
      'dverespe-cd6d24'
    LoginPrompt = False
    Provider = 'SQLOLEDB.1'
    Left = 34
    Top = 74
  end
  object Act_StoredProc: TADOStoredProc
    Connection = Act_Connection
    ProcedureName = 'InsertDetailedAct_Log;1'
    Parameters = <
      item
        Name = 'RETURN_VALUE'
        DataType = ftInteger
        Direction = pdReturnValue
        Precision = 10
        Value = Null
      end
      item
        Name = '@App_From'
        Attributes = [paNullable]
        DataType = ftString
        Size = 15
        Value = ''
      end
      item
        Name = '@IP_Address'
        Attributes = [paNullable]
        DataType = ftString
        Size = 15
        Value = ''
      end
      item
        Name = '@Trans'
        Attributes = [paNullable]
        DataType = ftString
        Size = 10
        Value = ''
      end
      item
        Name = '@DT_Sender'
        Attributes = [paNullable]
        DataType = ftString
        Size = 16
        Value = ''
      end
      item
        Name = '@ComputerName'
        Attributes = [paNullable]
        DataType = ftString
        Size = 35
        Value = ''
      end
      item
        Name = '@Description'
        Attributes = [paNullable]
        DataType = ftString
        Size = 100
        Value = ''
      end
      item
        Name = '@NTUserName'
        Attributes = [paNullable]
        DataType = ftString
        Size = 35
        Value = ''
      end
      item
        Name = '@Db_Time'
        Attributes = [paNullable]
        DataType = ftFloat
        Value = 0.000000000000000000
      end
      item
        Name = '@VIN'
        Attributes = [paNullable]
        DataType = ftString
        Size = 17
        Value = ''
      end
      item
        Name = '@Sequence_Number'
        Attributes = [paNullable]
        DataType = ftInteger
        Precision = 10
        Value = 0
      end
      item
        Name = '@Last_Modified'
        Attributes = [paNullable]
        DataType = ftString
        Size = 16
        Value = ''
      end>
    Left = 114
    Top = 74
  end
  object Inv_StoredProc: TADOStoredProc
    Connection = Inv_Connection
    ProcedureName = 'SELECT_AssyRatioInfo;1'
    Parameters = <
      item
        Name = '@RETURN_VALUE'
        DataType = ftInteger
        Direction = pdReturnValue
        Precision = 10
        Value = Null
      end
      item
        Name = '@AssyCode'
        Attributes = [paNullable]
        DataType = ftString
        Size = 12
        Value = Null
      end>
    Left = 200
    Top = 23
  end
  object Inv_Field_DataSet: TADODataSet
    Connection = Inv_Connection
    Parameters = <>
    Left = 200
    Top = 75
  end
  object Inv_DataSource: TDataSource
    DataSet = Grid_ClientDataSet
    Left = 34
    Top = 126
  end
  object Grid_ClientDataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'Inv_DataSetProvider'
    Left = 78
    Top = 192
  end
  object Inv_DataSetProvider: TDataSetProvider
    DataSet = Inv_DataSet
    Left = 200
    Top = 127
  end
  object INV_Forecast_DataSet: TADODataSet
    Connection = Inv_Connection
    CursorType = ctStatic
    CommandText = 
      'Select CH_SUPPLIER_CD '#39#39'SUPPLIER CODE'#39#39', CH_SUPPLIER_NA '#39#39'SUPPLI' +
      'ER NAME'#39#39', CH_ADDRESS '#39#39'ADDRESS'#39#39', CH_TEL '#39#39'TELEPHONE'#39#39', CH_FAX ' +
      #39#39'FAX'#39#39', CH_PERSON '#39#39'PERSON'#39#39#13#10'FROM INV_SUPPLIER_MST'
    Parameters = <>
    Left = 298
    Top = 24
  end
  object fiSupplierCode: TCIniField
    Section = 'SITE'
    Key = 'SupplierCode'
    Default = '05680'
    Rewrite = True
    Left = 298
    Top = 76
  end
  object NT: TNummiTime
    Time = 37609.595286504600000000
    NummiTime = '2002121914171200'
    TimeValue = 37609.595286504600000000
    Left = 298
    Top = 192
  end
  object INV_ShippingStoredProc: TADOStoredProc
    Connection = Inv_Connection
    ProcedureName = 'SELECT_AssyRatioInfo;1'
    Parameters = <
      item
        Name = '@RETURN_VALUE'
        DataType = ftInteger
        Direction = pdReturnValue
        Precision = 10
        Value = Null
      end
      item
        Name = '@AssyCode'
        Attributes = [paNullable]
        DataType = ftString
        Size = 12
        Value = Null
      end>
    Left = 416
    Top = 24
  end
  object INV_Order_StoredProc: TADOStoredProc
    Connection = Inv_Connection
    ProcedureName = 'SELECT_AssyRatioInfo;1'
    Parameters = <
      item
        Name = '@RETURN_VALUE'
        DataType = ftInteger
        Direction = pdReturnValue
        Precision = 10
        Value = Null
      end
      item
        Name = '@AssyCode'
        Attributes = [paNullable]
        DataType = ftString
        Size = 12
        Value = Null
      end>
    Left = 416
    Top = 132
  end
  object fiLogisticsInputDir: TCIniField
    Section = 'DIRECTORIES'
    Key = 'LogisticsInputDir'
    Default = 'c:\_Inventory_Control\'
    Rewrite = True
    Left = 416
    Top = 196
  end
  object fiForecastFilename: TCIniField
    Section = 'DIRECTORIES'
    Key = 'ForecastFilename'
    Default = 'nummi.prelftp'
    Rewrite = True
    Left = 416
    Top = 250
  end
  object fiReportsOutputDir: TCIniField
    Section = 'DIRECTORIES'
    Key = 'ReportsOutputDir'
    Default = 'c:\_Inventory_Control\Reports'
    Rewrite = True
    Left = 298
    Top = 251
  end
  object fiLogisticsFilename: TCIniField
    Section = 'DIRECTORIES'
    Key = 'LogisticsFilename'
    Default = 'nummi.prelftp'
    Rewrite = True
    Left = 416
    Top = 310
  end
  object fiForecastInputDir: TCIniField
    Section = 'DIRECTORIES'
    Key = 'ForecastInputDir'
    Default = 'c:\_Inventory_Control\Suppliers\NUMMI'
    Rewrite = True
    Left = 298
    Top = 314
  end
  object INV_Check_DataSet: TADODataSet
    Connection = Inv_Connection
    CursorType = ctStatic
    CommandText = 
      'Select CH_SUPPLIER_CD '#39#39'SUPPLIER CODE'#39#39', CH_SUPPLIER_NA '#39#39'SUPPLI' +
      'ER NAME'#39#39', CH_ADDRESS '#39#39'ADDRESS'#39#39', CH_TEL '#39#39'TELEPHONE'#39#39', CH_FAX ' +
      #39#39'FAX'#39#39', CH_PERSON '#39#39'PERSON'#39#39#13#10'FROM INV_SUPPLIER_MST'
    Parameters = <>
    Left = 416
    Top = 72
  end
  object fiLocalFTP: TCIniField
    Section = 'INIT'
    Key = 'LocalFTP'
    Default = 'False'
    Rewrite = True
    Left = 298
    Top = 366
  end
  object fiHideTerminated: TCIniField
    Section = 'DISPLAY'
    Key = 'HideTerminated'
    Default = 'True'
    Rewrite = True
    Left = 416
    Top = 360
  end
  object fiForecastUsageCompare: TCIniField
    Section = 'INIT'
    Key = 'ForecastUsageCompare'
    Default = '7'
    Rewrite = True
    Left = 200
    Top = 367
  end
  object fiUsageUpdateCompare: TCIniField
    Section = 'INIT'
    Key = 'UsageUpdateCompare'
    Default = '14'
    Rewrite = True
    Left = 38
    Top = 370
  end
  object fiPlantName: TCIniField
    Section = 'SITE'
    Key = 'PlantName'
    Default = 'NUMMI'
    Rewrite = True
    Left = 122
    Top = 434
  end
  object fiAssemblerName: TCIniField
    Section = 'SITE'
    Key = 'Assembler'
    Default = 'WQS'
    Rewrite = True
    Left = 96
    Top = 490
  end
  object fiFileALC: TCIniField
    Section = 'SITE'
    Key = 'FileALC'
    Default = 'True'
    Rewrite = True
    Left = 196
    Top = 434
  end
  object fiHighSequence: TCIniField
    Section = 'SITE'
    Key = 'HighSequence'
    Default = 'True'
    Rewrite = True
    Left = 196
    Top = 494
  end
  object fiDatabaseName: TCIniField
    Section = 'DATABASE'
    Key = 'DatabaseName'
    Default = 'INVENTORY'
    Rewrite = True
    Left = 300
    Top = 436
  end
  object fiHistoricalForecast: TCIniField
    Section = 'INIT'
    Key = 'HistoricalForecast'
    Default = '12'
    Rewrite = True
    Left = 418
    Top = 430
  end
  object fiUseFirstProductionDay: TCIniField
    Section = 'INIT'
    Key = 'UseFirstProductionDay'
    Default = 'True'
    Rewrite = True
    Left = 418
    Top = 484
  end
  object fiTextShippingFileDir: TCIniField
    Section = 'DIRECTORIES'
    Key = 'TextShippingFileDir'
    Default = 'D:\Daily Camex Results\'
    Rewrite = True
    Left = 298
    Top = 492
  end
  object fiPOEDISupport: TCIniField
    Section = 'SITE'
    Key = 'POEDISupport'
    Default = 'True'
    Rewrite = True
    Left = 124
    Top = 328
  end
  object fiBuildOut: TCIniField
    Section = 'DISPLAY'
    Key = 'BuildOut'
    Default = 'True'
    Rewrite = True
    Left = 40
    Top = 564
  end
  object fiInventoryConnection: TCIniField
    Section = 'DATABASE'
    Key = 'InventoryConnection'
    Default = 
      'Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security In' +
      'fo=False;User ID=Inventory;Initial Catalog=Inventory2;Data Sourc' +
      'e=Failproof1'
    Rewrite = True
    Left = 202
    Top = 560
  end
  object fiActivityConnection: TCIniField
    Section = 'DATABASE'
    Key = 'ActivityConnection'
    Default = 
      'Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security In' +
      'fo=False;User ID=Inventory;Initial Catalog=Activity;Data Source=' +
      'Failproof1'
    Rewrite = True
    Left = 110
    Top = 558
  end
  object fiTruckSeqLength: TCIniField
    Section = 'DISPLAY'
    Key = 'TruckSeqLength'
    Default = '3'
    Rewrite = True
    Left = 38
    Top = 616
  end
  object fiPassSeqLength: TCIniField
    Section = 'DISPLAY'
    Key = 'PassSeqLength'
    Default = '4'
    Rewrite = True
    Left = 104
    Top = 616
  end
  object fiGenerateEDI: TCIniField
    Section = 'SITE'
    Key = 'GenerateEDI'
    Default = 'True'
    Rewrite = True
    Left = 124
    Top = 374
  end
  object SiteDataSet: TADODataSet
    Connection = ALC_Connection
    CommandText = 'SELECT * FROM Site'
    Parameters = <>
    Left = 108
    Top = 254
  end
  object EDI810DataSet: TADODataSet
    Connection = Inv_Connection
    CommandText = 'REPORT_EDI810'
    CommandType = cmdStoredProc
    Parameters = <>
    Left = 180
    Top = 716
  end
  object UpdateReportCommand: TADOCommand
    Connection = ALC_Connection
    Parameters = <>
    Left = 272
    Top = 696
  end
  object EDI856DataSet: TADODataSet
    Connection = Inv_Connection
    CommandText = 'REPORT_EDI856'
    CommandType = cmdStoredProc
    Parameters = <
      item
        Name = '@RETURN_VALUE'
        DataType = ftInteger
        Direction = pdReturnValue
        Precision = 10
        Value = Null
      end>
    Left = 180
    Top = 668
  end
  object fiUseBCRatio: TCIniField
    Section = 'SITE'
    Key = 'UseBCRatio'
    Default = 'True'
    Rewrite = True
    Left = 408
    Top = 624
  end
  object fiRevSeqLookup: TCIniField
    Section = 'SITE'
    Key = 'RevSeqLookup'
    Default = '-30'
    Rewrite = True
    Left = 334
    Top = 624
  end
  object fiEDIOut: TCIniField
    Section = 'DIRECTORIES'
    Key = 'EDIOut'
    Default = 'c:\_Inventory_Control\EDIOut'
    Rewrite = True
    Left = 430
    Top = 556
  end
  object fiEDIIn: TCIniField
    Section = 'DIRECTORIES'
    Key = 'EDIIn'
    Default = 'c:\_Inventory_Control\EDIIn'
    Rewrite = True
    Left = 208
    Top = 604
  end
  object ALC_Connection: TADOConnection
    ConnectionString = 
      'Provider=SQLOLEDB.1;Password=nummidmv;Persist Security Info=True' +
      ';User ID=sa;Initial Catalog=VehicleOrderHERO;Data Source=FAILPRO' +
      'OF1'
    LoginPrompt = False
    Mode = cmRead
    Provider = 'SQLOLEDB.1'
    Left = 30
    Top = 260
  end
  object ALC_StoredProc: TADOStoredProc
    Connection = ALC_Connection
    ProcedureName = 'UPDATE_FRS;1'
    Parameters = <>
    Left = 188
    Top = 252
  end
  object fiALCConnection: TCIniField
    Section = 'DATABASE'
    Key = 'ALCConnection'
    Default = 
      'Provider=SQLOLEDB.1;Password=noway;Persist Security Info=True;Us' +
      'er ID=Inventory;Initial Catalog=TireOrder;Data Source=Failproof1'
    Rewrite = True
    Left = 306
    Top = 570
  end
  object ALC_DataSet: TADODataSet
    Connection = ALC_Connection
    CursorType = ctStatic
    CommandText = 'AD_GetSpecialDates;1'
    CommandType = cmdStoredProc
    ParamCheck = False
    Parameters = <>
    Left = 192
    Top = 186
  end
  object fiExcelOrderSheet: TCIniField
    Section = 'SITE'
    Key = 'ExcelOrderSheet'
    Default = 'True'
    Rewrite = True
    Left = 370
    Top = 690
  end
  object INV_Command: TADOCommand
    CommandType = cmdStoredProc
    Connection = Inv_Connection
    Parameters = <>
    Left = 112
    Top = 124
  end
  object fiEnableDataPurge: TCIniField
    Section = 'DATAPURGE'
    Key = 'EnableDataPurge'
    Default = 'FALSE'
    Rewrite = True
    Left = 34
    Top = 772
  end
  object fiPromptDataPurge: TCIniField
    Section = 'DATAPURGE'
    Key = 'PromptDataPurge'
    Default = 'FALSE'
    Rewrite = True
    Left = 80
    Top = 772
  end
  object fiDataRetention: TCIniField
    Section = 'DATAPURGE'
    Key = 'DataRetention'
    Default = '18'
    Rewrite = True
    Left = 134
    Top = 774
  end
  object fiPurgeRate: TCIniField
    Section = 'DATAPURGE'
    Key = 'PurgeRate'
    Default = 'Daily'
    Rewrite = True
    Left = 190
    Top = 774
  end
  object fiLastPurge: TCIniField
    Section = 'DATAPURGE'
    Key = 'LastPurge'
    Default = '2008050612000000'
    Rewrite = True
    Left = 244
    Top = 774
  end
  object fiPurgeDayWeekly: TCIniField
    Section = 'DATAPURGE'
    Key = 'PurgeDayWeekly'
    Default = 'Monday'
    Rewrite = True
    Left = 292
    Top = 776
  end
  object fiPurgeDayMonthly: TCIniField
    Section = 'DATAPURGE'
    Key = 'PurgeDayMonthly'
    Default = '1st'
    Rewrite = True
    Left = 336
    Top = 774
  end
  object fiFillDays: TCIniField
    Section = 'INIT'
    Key = 'FillDays'
    Default = '23'
    Rewrite = True
    Left = 34
    Top = 440
  end
  object fiConfirmOrderFileCreation: TCIniField
    Section = 'INIT'
    Key = 'ConfirmOrderFileCreation'
    Default = 'FALSE'
    Rewrite = True
    Left = 42
    Top = 328
  end
  object INV_StoredProc1: TADOStoredProc
    Connection = Inv_Connection
    ProcedureName = 'SELECT_AssyRatioInfo;1'
    Parameters = <
      item
        Name = '@RETURN_VALUE'
        DataType = ftInteger
        Direction = pdReturnValue
        Precision = 10
        Value = Null
      end
      item
        Name = '@AssyCode'
        Attributes = [paNullable]
        DataType = ftString
        Size = 12
        Value = Null
      end>
    Left = 298
    Top = 123
  end
  object fiTemplateDir: TCIniField
    Section = 'DIRECTORIES'
    Key = 'TemplateDir'
    Default = 'c:\_Inventory_Control\Templates'
    Rewrite = True
    Left = 488
    Top = 618
  end
  object fiUseApplicationDir: TCIniField
    Section = 'DIRECTORIES'
    Key = 'UseApplicationDir'
    Default = 'True'
    Rewrite = True
    Left = 496
    Top = 565
  end
  object fiCreatePOPriorToClose: TCIniField
    Section = 'INIT'
    Key = 'CreatePOPriorToClose'
    Default = 'TRUE'
    Rewrite = True
    Left = 206
    Top = 320
  end
end
