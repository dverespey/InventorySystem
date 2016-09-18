//****************************************************************
//
//       Inventory Control
//
//       Copyright (c) 2002-2008 TAI, Failproof Manufacturing Systems.
//
//****************************************************************
//
//  10/25/2002  Aaron Huge      Initial creation
//  11/14/2002  Aaron Huge      Alter to use two tire part numbers and two ratios
//                              and two wheel part numbers and two wheel ratios
//                              for ASSY Ratio form
//  12/13/2002  Aaron Huge      Altered ClearControls to clear the TNUMMIBmDateEdit controls
//  12/17/2002  Aaron Huge      Implemented ClientDataSet in GetRecConfStatInfo
//                              Implemented Quantity for RecConfStat...
//  12/18/2002  David Verespey  Merge changes
//  03/14/2002  David Verespey  Add Logistics
//  07/06/2004  David Verespey  Add First Production Day
//  04/12/2005  David Verespey  Add Monthly PO data
//  05/05/2005  David Verespey  Finish First Production Day/ CAMEX-WQS code merge
//  06/29/2005  David Verespey  Add auto scrap support
//  02/07/2006  David Verespey  Add Support for manual parts
//  05/07/2008  David Verespey  Add database auto purge

unit DataModule;

interface
uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, ADODB, StdCtrls, ExtCtrls, ComCtrls, UserInfo, VersionInfo, Variants,
  Mask, Grids, DBGrids, IdGlobal, DBClient, Provider, NUMMIBmDateEdit,
  Cinifld, NummiTime, DateUtils, NUMMIColumnComboBox;

type
  TColNum = 1..256;  // Excel columns only go up to IV
  TColString = string[2];

  TFileKind = (fText,fExcel,fBoth);

  TVehicleType = (vtPass,vtTruck);

type
  TInvUser = class(TUserInfo)
  private
    fAppUserID: String;
    fAppUserPass: String;
    fAppUserAdmin: Boolean;
  protected
  public
    property AppUserID: String read fAppUserID write fAppUserID;
    property AppUserPass: String read fAppUserPass write fAppUserPass;
    property AppUserAdmin: Boolean read fAppUserAdmin write fAppUserAdmin;
end;

type
  TTwoFieldObj = class(TObject)
  private
    fCode: String;
    fName: String;
  protected
  public
    constructor Create (Code: String; Name: String); overload;
    destructor Destroy; override;
    property Code: String read fCode write fCode;
    property Name: String read fName write fName;
end;

type
  TUserAdminDetail = class(TObject)
  private
    fID: String;
    fPass: String;
    fAdmin: Boolean;
  public
    constructor Create (fTempID: String; fTempPass: String; fTempAdmin: Boolean); overload;
    destructor Destroy; override;
    property ID: String read fID write fID;
    property Pass: String read fPass write fPass;
    property Admin: Boolean read fAdmin write fAdmin;
end;

type
  TData_Module = class(TDataModule)
    Inv_Connection: TADOConnection;
    Inv_DataSet: TADODataSet;
    Act_Connection: TADOConnection;
    Act_StoredProc: TADOStoredProc;
    Inv_StoredProc: TADOStoredProc;
    Inv_Field_DataSet: TADODataSet;
    Inv_DataSource: TDataSource;
    Grid_ClientDataSet: TClientDataSet;
    Inv_DataSetProvider: TDataSetProvider;
    INV_Forecast_DataSet: TADODataSet;
    fiSupplierCode: TCIniField;
    NT: TNummiTime;
    INV_ShippingStoredProc: TADOStoredProc;
    INV_Order_StoredProc: TADOStoredProc;
    fiLogisticsInputDir: TCIniField;
    fiForecastFilename: TCIniField;
    fiReportsOutputDir: TCIniField;
    fiLogisticsFilename: TCIniField;
    fiForecastInputDir: TCIniField;
    INV_Check_DataSet: TADODataSet;
    fiLocalFTP: TCIniField;
    fiHideTerminated: TCIniField;
    fiForecastUsageCompare: TCIniField;
    fiUsageUpdateCompare: TCIniField;
    fiPlantName: TCIniField;
    fiAssemblerName: TCIniField;
    fiFileALC: TCIniField;
    fiHighSequence: TCIniField;
    fiDatabaseName: TCIniField;
    fiHistoricalForecast: TCIniField;
    fiUseFirstProductionDay: TCIniField;
    fiTextShippingFileDir: TCIniField;
    fiPOEDISupport: TCIniField;
    fiBuildOut: TCIniField;
    fiInventoryConnection: TCIniField;
    fiActivityConnection: TCIniField;
    fiTruckSeqLength: TCIniField;
    fiPassSeqLength: TCIniField;
    fiGenerateEDI: TCIniField;
    SiteDataSet: TADODataSet;
    EDI810DataSet: TADODataSet;
    UpdateReportCommand: TADOCommand;
    EDI856DataSet: TADODataSet;
    fiUseBCRatio: TCIniField;
    fiRevSeqLookup: TCIniField;
    fiEDIOut: TCIniField;
    fiEDIIn: TCIniField;
    ALC_Connection: TADOConnection;
    ALC_StoredProc: TADOStoredProc;
    fiALCConnection: TCIniField;
    ALC_DataSet: TADODataSet;
    fiExcelOrderSheet: TCIniField;
    INV_Command: TADOCommand;
    fiEnableDataPurge: TCIniField;
    fiPromptDataPurge: TCIniField;
    fiDataRetention: TCIniField;
    fiPurgeRate: TCIniField;
    fiLastPurge: TCIniField;
    fiPurgeDayWeekly: TCIniField;
    fiPurgeDayMonthly: TCIniField;
    fiFillDays: TCIniField;
    fiConfirmOrderFileCreation: TCIniField;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
    fRecordID:integer;
    fASNStatus:string;

    fErrorCount: Integer;
    fSupCode: String;
    fSupName: String;
    fSupAddress: String;
    fSupCity: string;
    fSupState : string;
    fSupZip: string;
    fSupTel: String;
    fSupFax: String;
    fSupPerson: String;
    fSupEmail: string;
    fSupDirectory: string;
    fSupLogistics:string;
    fSupOutputFileType:string;
    fOrderFileTimestamp:boolean;
    fSiteNumberInOrder:boolean;
    fCreateOrderSheet:string;
    fInvAddPoint:string;


    fRenbanCode:string;
    fRenbanCount:string;
    fShipDays:integer;

    fLogName: String;
    fLogAddress: String;
    fLogCity: string;
    fLogState : string;
    fLogZip: string;
    fLogTel: String;
    fLogFax: String;
    fLogPerson: String;
    fLogEmail: string;
    fLogDirectory: string;

    fPartNum: String;
    fPartNumPrev: string;
    fPartName: String;
    fTireWheel: String;
    fCarTruck: String;
    fPartType:string;
    fKanban: String;
    f1LotQty: Integer;
    fPrice: Double;
    fQTY: Integer;
    fIrregulareQty:integer;
    fComments: String;
    fTotalPrice: Double;
    fLotSizeOrders: boolean;
    fLeadTime:integer;
    fLeadTimeMonday:integer;
    fLeadTimeTuesday:integer;
    fLeadTimeWednesday:integer;
    fLeadTimeThursday:integer;
    fLeadTimeFriday:integer;
    fLeadTimeSaturday:integer;

    fShipDaysMonday:integer;
    fShipDaysTuesday:integer;
    fShipDaysWednesday:integer;
    fShipDaysThursday:integer;
    fShipDaysFriday:integer;
    fShipDaysSaturday:integer;

    fLineName:string;

    fAssyManifestNo:string;
    fAssyCode: String;
    fAssyCodePrev: String;
    fAssyName: String;
    fTireQty: Integer;
//11-14-2002
    fValvePartNum:string;
    fFilmPartNum:string;

    fTirePartNum1: String;
    fTireRatio1: Integer;
    fTirePartNum2: String;
    fTireRatio2: Integer;
    fTirePartNum3: String;
    fTireRatio3: Integer;
    fWheelPartNum1: String;
    fWheelRatio1: Integer;
    fWheelPartNum2: String;
    fWheelRatio2: Integer;
    fWheelPartNum3: String;
    fWheelRatio3: Integer;
    fBroadCode: String;
    fBroadCodePrev: String;
    fWheelQty: Integer;
    fSpareTireQty: Integer;
    fSpareTirePartNum: String;
    fSpareWheelPartNum: String;

    fForecastRatio:integer;

    fSizeCode: String;
    fSizeName: String;
    fDailyUsage: Integer;
    fSafetyDays: Integer;

    fFRSNo: String;
    fFRSNoPrev: String;
    fRenban: String;
    fInTransit: String;
    fArrival: String;
    fOrder:string;
    fShip:string;
    fTrailerNo: String;
    fPlantYard: String;
    fWarehouse:string;
    fTerminated:string;
    fParkingSpot: String;
    fAssemblerYard: String;
    fAssemblerLocation: String;
    fEmptyTrailer: String;
    fDetention: String;
    fEditDate: String;
    fStartSeq: String;
    fLastSeq: String;
    fContinuation: Integer;
    fDivision: String;
    fBeforeDateTime, fAfterDateTime: TDateTime;
    fDiffDateTime: double;
    fAppFrom: String;
    fRefill:boolean;

    // DMV 12/18/2002
    fEventDate:       TDateTime;
    fEventType:       string;
    fEventTypeAbv:    string;
    fEventLine:       string;
    fEventDescription: string;
    fWeekNumber:      integer;
    fDayNumber:       integer;
    fProductionDate:  string;
    fProductionDateDT:  TDateTime;
    fBeginDate:       TDateTime;
    fEndDate:         TDateTime;
    fBeginDateStr:       string;
    fEndDateStr:         string;

    // DMV 07/06/2004
    fProdYear :string;

    //DMV 06/29/2005
    fScrapCount:integer;

    // DMV 04/12/2005
    fPOStart:string;
    fPOEnd:string;
    fPOQty:integer;
    fPOCharged:integer;
    fPOStartPrev:string;
    fPOEndPrev:string;
    fPickUp:string;
    fPickUpTime:string;
    fAssyCost:Double;
    fPONumber:string;
    fASN:integer;
    fINVOICE:integer;

    // DMV 05/23/2008
    fPartCost:Double;

    fEIN:integer;
    fEINStatus:string;
    fEINType:string;

    function LockSignature(Signature: String): Boolean;
    procedure CatchAll (Sender: TObject; E: Exception);
  public
    { Public declarations }
    fSupplierCodePrev:string;
    fRenbanCodePrev:string;
    gobjUser: TInvUser;
    property RecordID: integer read fRecordID write fRecordID;
    property ASNStatus: String read fASNStatus write fASNStatus;

    property SupplierCode: String read fSupCode write fSupCode;
    property SupplierName: String read fSupName write fSupName;
    property SupplierAddress: String read fSupAddress write fSupAddress;
    property SupplierCity: string read fSupCity write fSupCity;
    property SupplierState: string read fSupState write fSupState;
    property Supplierzip: string read fSupZip write fSupzip;
    property SupplierTelephone: String read fSupTel write fSupTel;
    property SupplierFax: String read fSupFax write fSupFax;
    property SupplierPerson: String read fSupPerson write fSupPerson;
    property SupplierEmail :string read fSupEmail write fSupEmail;
    property SupplierDirectory: String read fSupDirectory write fSupDirectory;
    property SupplierLogistics: string read fSupLogistics write fSupLogistics;
    property SupplierOutputFileType: string read fSupOutputFileType write fSupOutputFileType;
    property OrderFileTimestamp: boolean read fOrderFileTimestamp write fOrderFileTimestamp;
    property SiteNumberInOrder: boolean read fSiteNumberInOrder write fSiteNumberInOrder;
    property CreateOrderSheet: string read fCreateOrderSheet write fCreateOrderSheet;
    property InvAddPoint:string read fInvAddPoint write fInvAddPoint;


    property RenbanCode:string read fRenbancode write fRenbanCode;
    property RenbanCount:string read fRenbanCount write fRenbanCount;
    property Shipdays:integer read fShipDays write fShipDays;

    property LogisticsName: String read fLogName write fLogName;
    property LogisticsAddress: String read fLogAddress write fLogAddress;
    property LogisticsCity: string read fLogCity write fLogCity;
    property LogisticsState: string read fLogState write fLogState;
    property Logisticszip: string read fLogZip write fLogzip;
    property LogisticsTelephone: String read fLogTel write fLogTel;
    property LogisticsFax: String read fLogFax write fLogFax;
    property LogisticsPerson: String read fLogPerson write fLogPerson;
    property LogisticsEmail :string read fLogEmail write fLogEmail;
    property LogisticsDirectory: String read fLogDirectory write fLogDirectory;


    property PartNum: String read fPartNum write fPartNum;
    property PartNumPrev: String read fPartNumPrev write fPartNumPrev;
    property PartName: String read fPartName write fPartName;
    property TireWheel: String read fTireWheel write fTireWheel;
    property CarTruck: String read fCarTruck write fCarTruck;
    property PartType: String read fPartType write fPartType;

    property Kanban: String read fKanban write fKanban;
    property LotQty: Integer read f1LotQty write f1LotQty;
    property Price: Double read fPrice write fPrice;
    property Quantity: Integer read fQTY write fQTY;
    property IrregulareQty: integer read fIrregulareQty write fIrregulareQty;
    property TotalPrice: Double read fTotalPrice write fTotalPrice;
    property Comments: String read fComments write fComments;
    property LotSizeOrders: boolean read fLotSizeOrders write fLotSizeOrders;
    property LeadTime: integer read fLeadTime write fLeadTime;
    property LeadTimeMonday: integer read fLeadTimeMonday write fLeadTimeMonday;
    property LeadTimeTuesday: integer read fLeadTimeTuesday write fLeadTimeTuesday;
    property LeadTimeWednesday: integer read fLeadTimeWednesday write fLeadTimeWednesday;
    property LeadTimeThursday: integer read fLeadTimeThursday write fLeadTimeThursday;
    property LeadTimeFriday: integer read fLeadTimeFriday write fLeadTimeFriday;
    property LeadTimeSaturday: integer read fLeadTimeSaturday write fLeadTimeSaturday;
    property ShipDaysMonday: integer read fShipDaysMonday write fShipDaysMonday;
    property ShipDaysTuesday: integer read fShipDaysTuesday write fShipDaysTuesday;
    property ShipDaysWednesday: integer read fShipDaysWednesday write fShipDaysWednesday;
    property ShipDaysThursday: integer read fShipDaysThursday write fShipDaysThursday;
    property ShipDaysFriday: integer read fShipDaysFriday write fShipDaysFriday;
    property ShipDaysSaturday: integer read fShipDaysSaturday write fShipDaysSaturday;

//11-14-2002
    property AssyCode: String read fAssyCode write fAssyCode;
    property AssyCodePrev: String read fAssyCodePrev write fAssyCodePrev;
    property AssyName: String read fAssyName write fAssyName;
    property BroadcastCode: String read fBroadCode write fBroadCode;
    property BroadcastCodePrev: String read fBroadCodePrev write fBroadCodePrev;
    property TireQty: Integer read fTireQty write fTireQty;
    property ValvePartNum: String read fValvePartNum write fValvePartNum;
    property FilmPartNum: String read fFilmPartNum write fFilmPartNum;
    property TirePartNum1: String read fTirePartNum1 write fTirePartNum1;
    property TireRatio1: Integer read fTireRatio1 write fTireRatio1;
    property TirePartNum2: String read fTirePartNum2 write fTirePartNum2;
    property TireRatio2: Integer read fTireRatio2 write fTireRatio2;
    property TirePartNum3: String read fTirePartNum3 write fTirePartNum3;
    property TireRatio3: Integer read fTireRatio3 write fTireRatio3;
    property WheelPartNum1: String read fWheelPartNum1 write fWheelPartNum1;
    property WheelRatio1: Integer read fWheelRatio1 write fWheelRatio1;
    property WheelPartNum2: String read fWheelPartNum2 write fWheelPartNum2;
    property WheelRatio2: Integer read fWheelRatio2 write fWheelRatio2;
    property WheelPartNum3: String read fWheelPartNum3 write fWheelPartNum3;
    property WheelRatio3: Integer read fWheelRatio3 write fWheelRatio3;
    property WheelQty: Integer read fWheelQty write fWheelQty;
    property SpareTireQty: Integer read fSpareTireQty write fSpareTireQty;
    property SpareTirePartNum: String read fSpareTirePartNum write fSpareTirePartNum;
    property SpareWheelPartNum: String read fSpareWheelPartNum write fSpareWheelPartNum;

    property ForecastRatio: integer read fForecastRatio write fForecastRatio;

    property SizeCode: String read fSizeCode write fSizeCode;
    property SizeName: String read fSizeName write fSizeName;
    property DailyUsage: Integer read fDailyUsage write fDailyUsage;
    property SafetyDays: Integer read fSafetyDays write fSafetyDays;

    property FRSNo: String read fFRSNo write fFRSNo;
    property FRSNoPrev: String read fFRSNoPrev write fFRSNoPrev;
    property Order: string read fOrder write fOrder;
    property Ship: string read fShip write fShip;
    property Renban: String read fRenban write fRenban;
    property InTransit: String read fInTransit write fInTransit;
    property Arrival: String read fArrival write fArrival;
    property TrailerNo: String read fTrailerNo write fTrailerNo;
    property PlantYard: String read fPlantYard write fPlantYard;
    property Warehouse: string read fWarehouse write fWarehouse;
    property Terminated: string read fTerminated write fTerminated;
    property ParkingSpot: String read fParkingSpot write fParkingSpot;
    property AssemblerYard: String read fAssemblerYard write fAssemblerYard;
    property AssemblerLocation: String read fAssemblerLocation write fAssemblerLocation;
    property EmptyTrailer: String read fEmptyTrailer write fEmptyTrailer;
    property Detention: String read fDetention write fDetention;
    property EditDate: String read fEditDate write fEditDate;

    property StartSeq: String read fStartSeq write fStartSeq;
    property LastSeq: String read fLastSeq write fLastSeq;
    property Continuation: Integer read fContinuation write fContinuation;

    property Division: String read fDivision write fDivision;

    property Refill: boolean read fRefill write fRefill;

    property EIN: integer read fEIN write fEIN;
    property EINStatus: string read fEINStatus write fEINStatus;
    property EINType: string read fEINType write fEINType;

    property LineName: string read fLineName write fLineName;

    // DMV 12-18-2002
    property EventDate:       TDateTime read fEventDate       write fEventDate;
    property EventType:       string    read fEventType       write fEventType;
    property EventTypeAbv:    string    read fEventTypeAbv    write fEventTypeAbv;
    property EventLine:       string    read fEventLine       write fEventLine;
    property EventDescription:       string    read fEventDescription       write fEventDescription;
    property WeekNumber:      integer   read fWeekNumber      write fWeekNumber;
    property DayNumber:       integer   read fDayNumber       write fDayNumber;
    property ProductionDate:  string    read fProductionDate  write fProductionDate;
    property ProductionDateDT:  TDateTime    read fProductionDateDT  write fProductionDateDT;
    property BeginDate:       TDateTime read fBeginDate       write fBeginDate;
    property EndDate:         TDateTime read fEndDate         write fEndDate;
    property BeginDatestr:    String read fBeginDatestr          write fBeginDateStr;
    property EndDateStr:      String read fEndDateStr            write fEndDateStr;

    // DMV 07/06/2004
    property ProdYear:        string    read fProdYear        write fProdYear;

    // DMV 06/29/2005
    property ScrapCount:       integer   read fScrapCount      write fScrapCount;

    // DMV 04/12/2005
    property AssyManifestNo:  string    read fAssyManifestNo  write fAssyManifestNo;
    property POStart:         string    read fPOStart         write fPOStart;
    property POEnd:           string    read fPOEnd           write fPOEnd;
    property POQty:           integer   read fPOQty           write fPOQty;
    property POCharged:       integer   read fPOCharged       write fPOCharged;
    property POStartPrev:     string    read fPOStartPrev     write fPOStartPrev;
    property POEndPrev:       string    read fPOEndPrev       write fPOEndPrev;
    property PickUp:          string    read fPickUp          write fPickUp;
    property PickUpTime:      string    read fPickUpTime      write fPickUpTime;
    property AssyCost:        Double    read fAssyCost        write fAssyCost;
    property PONumber:        String    read fPONumber        write fPONumber;
    property ASN:             integer   read fASN             write fASN;
    property INVOICE:         integer   read fINVOICE         write fINVOICE;

    property PartCost:        Double    read fPArtCost        write fPartCost;

    //    procedure InitADOs;
    procedure FreeADOs;
    procedure LogActLog(fTrans: String; fDescription: String; fWithDB_Time: Integer=0);
    function ValidateUser(fUserID: String; fPassword: String): Boolean;
    procedure GetSupplierInfo;
    function InsertSupplierInfo: Boolean;
    procedure UpdateSupplierInfo;
    procedure DeleteSupplierInfo;

    // 03/14/2002 DMV
    procedure GetLogisticsInfo;
    function InsertLogisticsInfo: Boolean;
    procedure UpdateLogisticsInfo;
    procedure DeleteLogisticsInfo;

    procedure GetInventoryInfo;
    function InsertPartsStockInfo: Boolean;
    procedure UpdatePartsStockInfo;
    procedure DeletePartsStockInfo;

    procedure GetSizeInfo;
    function InsertSizeInfo: Boolean;
    procedure UpdateSizeInfo;
    procedure DeleteSizeInfo;

    // DMV 02-12-03
    procedure GetForecastDetailInfo;
    function InsertForecastDetailInfo: Boolean;
    procedure UpdateForecastDetailInfo;
    procedure DeleteForecastDetailInfo;

    // DMV 03-18-03
    procedure GetRenbanGroupInfo;
    function InsertRenbanGroupInfo: Boolean;
    procedure UpdateRenbanGroupInfo;
    procedure DeleteRenbanGroupInfo;

    // CAMEX PO Information
    procedure GetMonthlyPOInfo; // DMV 04/15/2005
    function InsertMonthlyPOInfo: Boolean;
    procedure UpdateMonthlyPOInfo;
    procedure DeleteMonthlyPOInfo;

    // EDI Manifest Information
    procedure GetManifestCostInfo; // DMV 11/13/2006
    function InsertManifestCostInfo: Boolean;
    procedure UpdateManifestCostInfo;
    procedure DeleteManifestCostInfo;

    procedure GetAssyRatioInfo;
    function InsertAssyRatioInfo: Boolean;
    procedure UpdateAssyRatioInfo;
    procedure DeleteAssyRatioInfo;

    // New BC Ratio Screen DMV 09/27/2006
    procedure GetBCRatioInfo;
    function InsertBCRatioInfo:boolean;
    procedure UpdateBCRatioInfo;
    procedure DeleteBCRatioInfo;

    procedure GetRecConfStatInfo;
    function InsertRecConfStatInfo: Boolean;
    procedure UpdateRecConfStatInfo;
    procedure UpdateRecConfStatRenbanInfo;
    procedure DeleteRecConfStatInfo;

    procedure GetStocktakingInfo;
    function InsertStocktakingInfo: Boolean;
    procedure UpdateStocktakingInfo;
    procedure DeleteStocktakingInfo;

    procedure GetNextASNDate;
    procedure GetNextProductionDate;
    procedure GetShippingInfo;
    procedure GetShippingInfoDetail;
    procedure UpdateShippingInfoDetail;
    procedure InsertShippingInfoDetail;

    //DMV 01/05/2005
    function GetLastSeqDate(offset:integer=0): Boolean;
    //DMV 09/16/2004
    function InsertExcelShippingInfo: Boolean;
    procedure InsertExcelShippingEndInfo;

    //DMV 06/02/2005
    function InsertBuildHist:Boolean;
    procedure GetBuildHist;
    function UpdateINVDone:boolean;

    // DMV 06/23/2005
    function InsertPOCharged: Boolean;

    //DMV 04/11/2005
    procedure GetExcelPOInfo;

    //DMV 06/29/2005
    function InsertAutoScrap: integer;

    //DMV 02/07/2006
    procedure GetPartsList;
    procedure GetPartsListCount;

    //DMV 05/07/2008
    function AutoPurge: boolean;

    function InsertShippingInfo: Boolean;
    function InsertShippingInfoManual: Boolean;
    function InsertShippingDetailManual: Boolean;
    function CheckShippingInfo: Boolean;

    function InsertINVInfo: Boolean;
    function InsertASNInfo: Boolean;
    function UpdateASNStatus: Boolean;

    procedure GetRecProdRejInfo;
    function InsertRecProdRejInfo: Boolean;
    procedure UpdateRecProdRejInfo;
    procedure DeleteRecProdRejInfo;

    procedure SetComboBoxesWithUserObj(ComboBox: TComboBox; fQryString: String);
    function InsertUser(fTempID: String; fTempPass: String; fTempAdmin: Boolean): Boolean;
    procedure UpdateUserInfo(fTempOldID: String; fTempOldPass: String; fTempNewID: String; fTempNewPass: String; fTempAdmin: Boolean);
    procedure DeleteUserInfo(fTempID: String; fTempPass: String);

    procedure SelectSingleFieldALC(fTableName: String; fFieldName: String; ComboBox: TComboBox; Distinct:boolean=TRUE);
    procedure SelectDependantSingleField(StoredProc: String; ParameterName: String; Field:string; Value: Variant; ComboBox: TComboBox);
    procedure SelectSingleField(fTableName: String; fFieldName: String; ComboBox: TComboBox; Distinct:boolean=TRUE);
    procedure SelectMultiField(fTableName: String; fFields: String; ComboBox: TNUMMIColumnComboBox);
    procedure SelectMultiFieldALC(fTableName: String; fFields: String; ComboBox: TNUMMIColumnComboBox; fWhereAnd: string = '');
    procedure SetComboBoxesWithObj(ComboBox: TComboBox; fQryString: String);
    procedure SearchCombo(ComboBox: TComboBox; value: String);
    procedure SearchMultiCombo(ComboBox: TNUMMIColumnComboBox; value: String; Column:integer=0);
//    procedure SetGrid(StringGrid: TStringGrid);
    procedure ClearControls(Control: TControl);
    procedure JustifyColumns(DBGrid: TDBGrid);
    procedure SetTodaysDate(Panel: TPanel);

    // DMV 12-18-2002
    procedure GetOvertimeHolidayInfo;
    procedure DeleteOvertimeHolidayInfo;
    function InsertOvertimeHolidayInfo: Boolean;

    // DMV 07/06/2004
    procedure GetFirstProductionDayInfo;
    // DMV 05/05/2005
    procedure DeleteFirstProductiondayInfo;
    function InsertFirstProductionDayInfo: boolean;

    //function GetProductionDataSeq(BeginSeq,EndSeq,count:integer;Line:TVehicleType):boolean;
    function CalculateFRS:boolean;
    function UpdateEINStatus:boolean;
    function CalculateASNFRS:boolean;
    procedure DoPartNumberInventory(partnumber:string;ratio,qty:integer);
    function GetLastProductionDate:TDate;
    function XlColToInt(const Col: TColString): TColNum;
    function IntToXlCol(ColNum: TColNum): TColString;
  end;

var
  Data_Module: TData_Module;
const
  MaxLetters      = 26;

implementation

{$R *.dfm}
constructor TTwoFieldObj.Create (Code: String; Name: String);
begin
	inherited Create;
	fCode := Code;
  fName := Name;
end;

destructor TTwoFieldObj.Destroy;
begin
end;

constructor TUserAdminDetail.Create(fTempID: String; fTempPass: String; fTempAdmin: Boolean);
Begin
	inherited Create;
  fID := fTempID;
  fPass := fTempPass;
  fAdmin := fTempAdmin;
End;

destructor TUserAdminDetail.Destroy;
Begin
End;


procedure TData_Module.DataModuleCreate(Sender: TObject);
var
  fTempExeName: string;
  //tempconnstr:string;
begin
  Application.OnException := CatchAll;
  gobjUser := TInvUser.Create(Self);

  fTempExeName := ExtractFileName(Application.ExeName);
  fAppFrom := Copy(fTempExeName, 1, Length(fTempExeName) - 4);
  If Not(LockSignature(fTempExeName)) Then
  Begin
    ShowMessage (fTempExeName + ' is already running.');
    Application.Terminate;
    Exit;
  End;

  // Use raw connection strings
  Inv_Connection.ConnectionString:=fiInventoryConnection.AsString;
  Act_Connection.ConnectionString:=fiActivityConnection.AsString;
  ALC_Connection.ConnectionString:=fiALCConnection.AsString;
end;

procedure TData_Module.CatchAll (Sender: TObject; E: Exception);
begin
  LogActLog('ERROR','Unhandled Exception '+e.message);
	Application.Terminate;
end;      //CatchAll

function TData_Module.LockSignature(Signature: String): Boolean;
var fSigHandle: Cardinal;
Begin
  fSigHandle := CreateSemaphore(nil, 0, 1, PChar(Signature));
  result := (fSigHandle <> 0) and (GetLastError() <> ERROR_ALREADY_EXISTS);

  If Not(result) Then
    If (fSigHandle <> 0) Then
      CloseHandle (fSigHandle);
End;      //LockSignature

procedure TData_Module.FreeADOs;
Begin
  Inv_DataSet.Close;
  Inv_StoredProc.Close;
  Inv_Connection.Close;
  Act_StoredProc.Close;
  Act_Connection.Close;
End;      //FreeADOs

procedure TData_Module.GetLogisticsInfo;
begin
  try
    try
      With Inv_DataSet do
      Begin
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'dbo.SELECT_LogisticsInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@LogCode';
        Parameters.ParamValues['@LogCode'] := '';

        fBeforeDateTime := Time;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to get logistics data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to get logistics data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to get logistics data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('GET LOG', 'SELECTED all logistics', 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          GetLogisticsInfo
        Else
          LogActLog('ERROR', 'FAILED to SELECT all logistics. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
end;

function TData_Module.InsertLogisticsInfo: Boolean;
var
  fDescription: String;
Begin
  Result := False;
  fDescription := 'INSERT ' + fLogName;

  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.SELECT_LogisticsInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@LogisticsName';
        Parameters.ParamValues['@LogisticsName'] := fLogName;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to get logistics dup data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to get logistics dup data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to get logistics dup data');
        end
        else
        begin
          If RecordCount = 0 Then
          Begin
            Close;
            ProcedureName := 'dbo.INSERT_LogisticsInfo;1';
            Parameters.Clear;
            Parameters.AddParameter.Name := '@LogName';
            Parameters.ParamValues['@LogName'] := fLogName;
            Parameters.AddParameter.Name := '@LogAdd';
            Parameters.ParamValues['@LogAdd'] := fLogAddress;
            Parameters.AddParameter.Name := '@LogCity';
            Parameters.ParamValues['@LogCity'] := fLogCity;
            Parameters.AddParameter.Name := '@LogState';
            Parameters.ParamValues['@LogState'] := fLogState;
            Parameters.AddParameter.Name := '@LogZip';
            Parameters.ParamValues['@LogZip'] := fLogZip;
            Parameters.AddParameter.Name := '@LogPhone';
            Parameters.ParamValues['@LogPhone'] := fLogTel;
            Parameters.AddParameter.Name := '@LogFax';
            Parameters.ParamValues['@LogFax'] := fLogFax;
            Parameters.AddParameter.Name := '@LogPerson';
            Parameters.ParamValues['@LogPerson'] := fLogPerson;
            Parameters.AddParameter.Name := '@LogEmail';
            Parameters.ParamValues['@LogEmail'] := fLogEmail;
            Parameters.AddParameter.Name := '@LogDirectory';
            Parameters.ParamValues['@LogDirectory'] := fLogDirectory;

            fBeforeDateTime := Time;
            ExecProc;
            if Inv_Connection.Errors.Count > 0 then
            begin
              ShowMessage('Unable to insert logistics data, '+Inv_Connection.Errors.Item[0].Get_Description);
              LogActLog('ERROR','Unable to insert logistics data, '+Inv_Connection.Errors.Item[0].Get_Description);
              raise EDatabaseError.Create('Unable to insert logistics data');
            end
            else
            begin
              fAfterDateTime := Time;
              fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
              Result := True;
            end;
          End
          Else
          begin
            fDescription := 'FAILED to ' + fDescription + ' (DUPLICATE)';
            ShowMessage('Unable to insert duplicate logistics name('+fLogName+')');
          end;
        end;
      End;    //With
      LogActLog('INSERT Log', fDescription, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          InsertLogisticsInfo
        Else
          LogActLog('ERROR', 'FAILED to INSERT ' + fLogName + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;      //InsertLogisticsInfo

procedure TData_Module.UpdateLogisticsInfo;
Begin
  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.UPDATE_LogisticsInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@LogName';
        Parameters.ParamValues['@LogName'] := fLogName;
        Parameters.AddParameter.Name := '@LogAdd';
        Parameters.ParamValues['@LogAdd'] := fLogAddress;
        Parameters.AddParameter.Name := '@LogCity';
        Parameters.ParamValues['@LogCity'] := fLogCity;
        Parameters.AddParameter.Name := '@LogState';
        Parameters.ParamValues['@LogState'] := fLogState;
        Parameters.AddParameter.Name := '@LogZip';
        Parameters.ParamValues['@LogZip'] := fLogZip;
        Parameters.AddParameter.Name := '@LogPhone';
        Parameters.ParamValues['@LogPhone'] := fLogTel;
        Parameters.AddParameter.Name := '@LogFax';
        Parameters.ParamValues['@LogFax'] := fLogFax;
        Parameters.AddParameter.Name := '@LogPerson';
        Parameters.ParamValues['@LogPerson'] := fLogPerson;
        Parameters.AddParameter.Name := '@LogEmail';
        Parameters.ParamValues['@LogEmail'] := fLogEmail;
        Parameters.AddParameter.Name := '@LogDirectory';
        Parameters.ParamValues['@LogDirectory'] := fLogDirectory;
        Parameters.AddParameter.Name := '@LogisticsID';
        Parameters.ParamValues['@LogisticsID'] := fRecordID;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to update logistics data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to update logistics data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to update logistics data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          UpdateLogisticsInfo
        Else
          LogActLog('ERROR', 'FAILED to UPDATE ' + fLogName + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName, 0)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;      //UpdateLogisticsInfo

procedure TData_Module.DeleteLogisticsInfo;
begin
  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.DELETE_LogisticsInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@LogisticsID';
        Parameters.ParamValues['@LogisticsID'] := fRecordID;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to delete logistics data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to delete logistics data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to delete logistics data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('DELETE LOG', 'DELETED ' + fLogName, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          DeleteLogisticsInfo
        Else
          LogActLog('ERROR', 'FAILED to DELETE ' + fLogName + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
end;

procedure TData_Module.GetSupplierInfo;
Begin
  try
    try
      With Inv_DataSet do
      Begin
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'dbo.SELECT_SupplierInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@SupCode';
        Parameters.ParamValues['@SupCode'] := '';

        fBeforeDateTime := Time;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to get supplier data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to get supplier data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to get supplier data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('GET SUPS', 'SELECTED all suppliers', 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          GetSupplierInfo
        Else
          LogActLog('ERROR', 'FAILED to SELECT all suppliers. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    fErrorCount := 0;
  end;
End;      //GetSupplierInfo

function TData_Module.InsertSupplierInfo: Boolean;
var
  fDescription: String;
Begin
  Result := False;
  fDescription := 'INSERT ' + fSupCode;

  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.SELECT_SupplierInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@SupCode';
        Parameters.ParamValues['@SupCode'] := fSupCode;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to get supplier dup data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to get supplier dup data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to get supplier dup data');
        end
        else
        begin
          If RecordCount = 0 Then
          Begin
            Close;
            ProcedureName := 'dbo.INSERT_SupplierInfo;1';
            Parameters.Clear;
            Parameters.AddParameter.Name := '@SupCode';
            Parameters.ParamValues['@SupCode'] := fSupCode;
            Parameters.AddParameter.Name := '@SupName';
            Parameters.ParamValues['@SupName'] := fSupName;
            Parameters.AddParameter.Name := '@SupAdd';
            Parameters.ParamValues['@SupAdd'] := fSupAddress;
            Parameters.AddParameter.Name := '@SupCity';
            Parameters.ParamValues['@SupCity'] := fSupCity;
            Parameters.AddParameter.Name := '@SupState';
            Parameters.ParamValues['@SupState'] := fSupState;
            Parameters.AddParameter.Name := '@SupZip';
            Parameters.ParamValues['@SupZip'] := fSupZip;
            Parameters.AddParameter.Name := '@SupPhone';
            Parameters.ParamValues['@SupPhone'] := fSupTel;
            Parameters.AddParameter.Name := '@SupFax';
            Parameters.ParamValues['@SupFax'] := fSupFax;
            Parameters.AddParameter.Name := '@SupPerson';
            Parameters.ParamValues['@SupPerson'] := fSupPerson;
            Parameters.AddParameter.Name := '@SupEmail';
            Parameters.ParamValues['@SupEmail'] := fSupEmail;
            Parameters.AddParameter.Name := '@SupDirectory';
            Parameters.ParamValues['@SupDirectory'] := fSupDirectory;
            Parameters.AddParameter.Name := '@SupLogistics';
            Parameters.ParamValues['@SupLogistics'] := fLogName;
            Parameters.AddParameter.Name := '@SupOutputFileType';
            Parameters.ParamValues['@SupOutputFileType'] := fSupOutputFileType;
            Parameters.AddParameter.Name := '@OrderFileTimestamp';
            Parameters.ParamValues['@OrderFileTimestamp'] := fOrderFileTimestamp;
            Parameters.AddParameter.Name := '@SiteNumberInOrder';
            Parameters.ParamValues['@SiteNumberInOrder'] := fSiteNumberInOrder;
            Parameters.AddParameter.Name := '@CreateOrderSheet';
            Parameters.ParamValues['@CreateOrderSheet'] := fCreateOrderSheet;
            Parameters.AddParameter.Name := '@InvAddPoint';
            Parameters.ParamValues['@InvAddPoint'] := fInvAddPoint;

            fBeforeDateTime := Time;
            ExecProc;
            if Inv_Connection.Errors.Count > 0 then
            begin
              ShowMessage('Unable to insert supplier data, '+Inv_Connection.Errors.Item[0].Get_Description);
              LogActLog('ERROR','Unable to insert supplier data, '+Inv_Connection.Errors.Item[0].Get_Description);
              raise EDatabaseError.Create('Unable to insert supplier data');
            end
            else
            begin
              fAfterDateTime := Time;
              fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
              Result := True;
            end;
          End
          Else
            fDescription := 'FAILED to ' + fDescription + ' (DUPLICATE)';
        end;
      End;    //With
      LogActLog('INSERT SUP', fDescription, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          InsertSupplierInfo
        Else
          LogActLog('ERROR', 'FAILED to INSERT ' + fSupCode + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;      //InsertSupplierInfo

procedure TData_Module.UpdateSupplierInfo;
Begin
  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.UPDATE_SupplierInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@SupCode';
        Parameters.ParamValues['@SupCode'] := fSupCode;
        Parameters.AddParameter.Name := '@SupName';
        Parameters.ParamValues['@SupName'] := fSupName;
        Parameters.AddParameter.Name := '@SupAdd';
        Parameters.ParamValues['@SupAdd'] := fSupAddress;
        Parameters.AddParameter.Name := '@SupCity';
        Parameters.ParamValues['@SupCity'] := fSupCity;
        Parameters.AddParameter.Name := '@SupState';
        Parameters.ParamValues['@SupState'] := fSupState;
        Parameters.AddParameter.Name := '@SupZip';
        Parameters.ParamValues['@Supzip'] := fSupZip;
        Parameters.AddParameter.Name := '@SupPhone';
        Parameters.ParamValues['@SupPhone'] := fSupTel;
        Parameters.AddParameter.Name := '@SupFax';
        Parameters.ParamValues['@SupFax'] := fSupFax;
        Parameters.AddParameter.Name := '@SupPerson';
        Parameters.ParamValues['@SupPerson'] := fSupPerson;
        Parameters.AddParameter.Name := '@SupEmail';
        Parameters.ParamValues['@SupEmail'] := fSupEmail;
        Parameters.AddParameter.Name := '@SupDirectory';
        Parameters.ParamValues['@SupDirectory'] := fSupDirectory;
        Parameters.AddParameter.Name := '@SupLogistics';
        Parameters.ParamValues['@SupLogistics'] := fLogName;
        Parameters.AddParameter.Name := '@SupOutputFileType';
        Parameters.ParamValues['@SupOutputFileType'] := fSupOutputFileType;
        Parameters.AddParameter.Name := '@OrderFileTimestamp';
        Parameters.ParamValues['@OrderFileTimestamp'] := fOrderFileTimestamp;
        Parameters.AddParameter.Name := '@SiteNumberInOrder';
        Parameters.ParamValues['@SiteNumberInOrder'] := fSiteNumberInOrder;
        Parameters.AddParameter.Name := '@CreateOrderSheet';
        Parameters.ParamValues['@CreateOrderSheet'] := fCreateOrderSheet;
        Parameters.AddParameter.Name := '@InvAddPoint';
        Parameters.ParamValues['@InvAddPoint'] := fInvAddPoint;
        Parameters.AddParameter.Name := '@SupplierID';
        Parameters.ParamValues['@SupplierID'] := fRecordID;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to update supplier data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to update supplier data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to update supplier data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('UPDATE SUP', 'UPDATE ' + fSupCode, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          UpdateSupplierInfo
        Else
          LogActLog('ERROR', 'FAILED to UPDATE ' + fSupCode + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;      //UpdateSupplierInfo

procedure TData_Module.DeleteSupplierInfo;
Begin
  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.DELETE_SupplierInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@SupplierID';
        Parameters.ParamValues['@SupplierID'] := fRecordID;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to delete supplier data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to delete supplier data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to delete supplier data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('DELETE SUP', 'DELETED ' + fSupCode, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          DeleteSupplierInfo
        Else
          LogActLog('ERROR', 'FAILED to DELETE ' + fSupCode + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;      //DeleteSupplierInfo

procedure TData_Module.GetInventoryInfo;
Begin
  try
    try
      With Inv_DataSet do
      Begin
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'dbo.SELECT_PartsStockInfo;1';

        Parameters.Clear;
        Parameters.AddParameter.Name := '@PartNum';
        Parameters.ParamValues['@PartNum'] := '';

        fBeforeDateTime := Time;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to get Part Master data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to get Part Master data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to get Part Master data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('GET PARTS', 'SELECTED all parts', 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          GetInventoryInfo
        Else
          LogActLog('ERROR', 'FAILED SELECT all parts. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    fErrorCount := 0;
  end;

End;        //GetInventoryInfo

function TData_Module.InsertPartsStockInfo: Boolean;
var
  fDescription: String;
Begin
  Result := False;
  fDescription := 'INSERT PartsStock Sup: ' + fSupCode + ' and Part: ' + fPartNum + ' Qty:'+IntToStr(fQty);

  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.SELECT_PartsStockInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@PartNum';
        Parameters.ParamValues['@PartNum'] := fPartNum;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Fail dup check Part Master data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Fail dup check Part Master data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Fail dup check Part Master data');
        end
        else
        begin
          If RecordCount = 0 Then
          Begin
            Close;
            ProcedureName := 'dbo.INSERT_PartsStockInfo;1';
            Parameters.Clear;
            Parameters.AddParameter.Name := '@SupCode';
            Parameters.ParamValues['@SupCode'] := fSupCode;
            Parameters.AddParameter.Name := '@LogisticsName';
            Parameters.ParamValues['@LogisticsName'] := fLogName;
            Parameters.AddParameter.Name := '@PartNum';
            Parameters.ParamValues['@PartNum'] := fPartNum;
            Parameters.AddParameter.Name := '@PartsName';
            Parameters.ParamValues['@PartsName'] := fPartName;
            Parameters.AddParameter.Name := '@SizeCode';
            Parameters.ParamValues['@SizeCode'] := fSizeCode;
            Parameters.AddParameter.Name := '@LineName';
            Parameters.ParamValues['@LineName'] := fLineName;
            Parameters.AddParameter.Name := '@Kanban';
            Parameters.ParamValues['@Kanban'] := fKanban;
            Parameters.AddParameter.Name := '@LotQty';
            Parameters.ParamValues['@LotQty'] := f1LotQty;
            Parameters.AddParameter.Name := '@QTY';
            Parameters.ParamValues['@QTY'] := fQTY;
            Parameters.AddParameter.Name := '@Comments';
            Parameters.ParamValues['@Comments'] := fComments;
            Parameters.AddParameter.Name := '@RenbanCode';
            Parameters.ParamValues['@RenbanCode'] := fRenbanCode;
            Parameters.AddParameter.Name := '@LotSizeOrders';
            Parameters.ParamValues['@LotSizeOrders'] := fLotSizeOrders;
            Parameters.AddParameter.Name := '@LeadTime';
            Parameters.ParamValues['@LeadTime'] := fLeadTime;
            Parameters.AddParameter.Name := '@LeadTimeMonday';
            Parameters.ParamValues['@LeadTimeMonday'] := fLeadTimeMonday;
            Parameters.AddParameter.Name := '@LeadTimeTuesday';
            Parameters.ParamValues['@LeadTimeTuesday'] := fLeadTimeTuesday;
            Parameters.AddParameter.Name := '@LeadTimeWednesday';
            Parameters.ParamValues['@LeadTimeWednesday'] := fLeadTimeWednesday;
            Parameters.AddParameter.Name := '@LeadTimeThursday';
            Parameters.ParamValues['@LeadTimeThursday'] := fLeadTimeThursday;
            Parameters.AddParameter.Name := '@LeadTimeFriday';
            Parameters.ParamValues['@LeadTimeFriday'] := fLeadTimeFriday;
            Parameters.AddParameter.Name := '@LeadTimeSaturday';
            Parameters.ParamValues['@LeadTimeSaturday'] := fLeadTimeSaturday;
            Parameters.AddParameter.Name := '@ShipDays';
            Parameters.ParamValues['@ShipDays'] := fShipDays;
            Parameters.AddParameter.Name := '@ShipDaysMonday';
            Parameters.ParamValues['@ShipDaysMonday'] := fShipDaysMonday;
            Parameters.AddParameter.Name := '@ShipDaysTuesday';
            Parameters.ParamValues['@ShipDaysTuesday'] := fShipDaysTuesday;
            Parameters.AddParameter.Name := '@ShipDaysWednesday';
            Parameters.ParamValues['@ShipDaysWednesday'] := fShipDaysWednesday;
            Parameters.AddParameter.Name := '@ShipDaysThursday';
            Parameters.ParamValues['@ShipDaysThursday'] := fShipDaysThursday;
            Parameters.AddParameter.Name := '@ShipDaysFriday';
            Parameters.ParamValues['@ShipDaysFriday'] := fShipDaysFriday;
            Parameters.AddParameter.Name := '@ShipDaysSaturday';
            Parameters.ParamValues['@ShipDaysSaturday'] := fShipDaysSaturday;
            Parameters.AddParameter.Name := '@RenbanCount';
            Parameters.ParamValues['@RenbanCount'] := fRenbanCount;
            Parameters.AddParameter.Name := '@PartType';
            Parameters.ParamValues['@PartType'] := fPartType;
            Parameters.AddParameter.Name := '@PartCost';
            Parameters.ParamValues['@PartCost'] := fPartCost;

            fBeforeDateTime := Time;
            ExecProc;
            if Inv_Connection.Errors.Count > 0 then
            begin
              ShowMessage('Fail insert Part Master data, '+Inv_Connection.Errors.Item[0].Get_Description);
              LogActLog('ERROR','Fail insert Part Master data, '+Inv_Connection.Errors.Item[0].Get_Description);
              raise EDatabaseError.Create('Fail insert Part Master data');
            end
            else
            begin
              fAfterDateTime := Time;
              fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
              Result := True;
            end;
          End
          Else
            fDescription := 'FAILED to ' + fDescription + ' (DUPLICATE)';
        end;
      End;    //With
      LogActLog('INS PART', fDescription, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          InsertPartsStockInfo
        Else
          LogActLog('ERROR', 'FAILED: ' + fDescription + ' Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;      //InsertPartsStockInfo

procedure TData_Module.UpdatePartsStockInfo;
Begin
  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.UPDATE_PartsStockInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@SupCode';
        Parameters.ParamValues['@SupCode'] := fSupCode;
        Parameters.AddParameter.Name := '@LogisticsName';
        Parameters.ParamValues['@LogisticsName'] := fLogName;
        Parameters.AddParameter.Name := '@PartNum';
        Parameters.ParamValues['@PartNum'] := fPartNum;
        Parameters.AddParameter.Name := '@PartsName';
        Parameters.ParamValues['@PartsName'] := fPartName;
        Parameters.AddParameter.Name := '@SizeCode';
        Parameters.ParamValues['@SizeCode'] := fSizeCode;
        Parameters.AddParameter.Name := '@LineName';
        Parameters.ParamValues['@LineName'] := fLineName;
        Parameters.AddParameter.Name := '@Kanban';
        Parameters.ParamValues['@Kanban'] := fKanban;
        Parameters.AddParameter.Name := '@LotQty';
        Parameters.ParamValues['@LotQty'] := f1LotQty;
        Parameters.AddParameter.Name := '@QTY';
        Parameters.ParamValues['@QTY'] := fQTY;
        Parameters.AddParameter.Name := '@Comments';
        Parameters.ParamValues['@Comments'] := fComments;
        Parameters.AddParameter.Name := '@RenbanCode';
        Parameters.ParamValues['@RenbanCode'] := fRenbanCode;
        Parameters.AddParameter.Name := '@LotSizeOrders';
        Parameters.ParamValues['@LotSizeOrders'] := fLotSizeOrders;
        Parameters.AddParameter.Name := '@LeadTime';
        Parameters.ParamValues['@LeadTime'] := fLeadTime;
        Parameters.AddParameter.Name := '@LeadTimeMonday';
        Parameters.ParamValues['@LeadTimeMonday'] := fLeadTimeMonday;
        Parameters.AddParameter.Name := '@LeadTimeTuesday';
        Parameters.ParamValues['@LeadTimeTuesday'] := fLeadTimeTuesday;
        Parameters.AddParameter.Name := '@LeadTimeWednesday';
        Parameters.ParamValues['@LeadTimeWednesday'] := fLeadTimeWednesday;
        Parameters.AddParameter.Name := '@LeadTimeThursday';
        Parameters.ParamValues['@LeadTimeThursday'] := fLeadTimeThursday;
        Parameters.AddParameter.Name := '@LeadTimeFriday';
        Parameters.ParamValues['@LeadTimeFriday'] := fLeadTimeFriday;
        Parameters.AddParameter.Name := '@LeadTimeSaturday';
        Parameters.ParamValues['@LeadTimeSaturday'] := fLeadTimeSaturday;
        Parameters.AddParameter.Name := '@ShipDays';
        Parameters.ParamValues['@ShipDays'] := fShipDays;
        Parameters.AddParameter.Name := '@ShipDaysMonday';
        Parameters.ParamValues['@ShipDaysMonday'] := fShipDaysMonday;
        Parameters.AddParameter.Name := '@ShipDaysTuesday';
        Parameters.ParamValues['@ShipDaysTuesday'] := fShipDaysTuesday;
        Parameters.AddParameter.Name := '@ShipDaysWednesday';
        Parameters.ParamValues['@ShipDaysWednesday'] := fShipDaysWednesday;
        Parameters.AddParameter.Name := '@ShipDaysThursday';
        Parameters.ParamValues['@ShipDaysThursday'] := fShipDaysThursday;
        Parameters.AddParameter.Name := '@ShipDaysFriday';
        Parameters.ParamValues['@ShipDaysFriday'] := fShipDaysFriday;
        Parameters.AddParameter.Name := '@ShipDaysSaturday';
        Parameters.ParamValues['@ShipDaysSaturday'] := fShipDaysSaturday;
        Parameters.AddParameter.Name := '@RenbanCount';
        Parameters.ParamValues['@RenbanCount'] := fRenbanCount;
        Parameters.AddParameter.Name := '@PartType';
        Parameters.ParamValues['@PartType'] := fPartType;
        Parameters.AddParameter.Name := '@PartCost';
        Parameters.ParamValues['@PartCost'] := fPartCost;
        Parameters.AddParameter.Name := '@PartID';
        Parameters.ParamValues['@PartID'] := fRecordID;


        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Fail update Part Master data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Fail update Part Master data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Fail update Part Master data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('UPDATE PRT', 'UPDATE S:' + fSupCode + ' P:' + fPartNum + ' Q:'+IntToStr(fQty), 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          UpdatePartsStockInfo
        Else
          LogActLog('ERROR', 'FAILED UPDATE S:' + fSupCode + ' P:' + fPartNum + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;      //UpdatePartsStockInfo

procedure TData_Module.DeletePartsStockInfo;
Begin
  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.DELETE_PartsStockInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@PartID';
        Parameters.ParamValues['@PartID'] := fRecordID;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Fail delete Part Master data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Fail delete Part Master data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Fail delete Part Master data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('DELETE PRT', 'DELETED S:' + fSupCode +  'P:' + fPartNum, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          DeletePartsStockInfo
        Else
          LogActLog('ERROR', 'FAILED DELETE PartsStockInfo S:' + fSupCode + ' P:' + fPartNum + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;      //DeletePartsStockInfo

procedure TData_Module.GetManifestCostInfo; // DMV 11/13/2006
begin
  try
    try
      With Inv_DataSet do
      Begin
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'dbo.SELECT_ManifestCost;1';
        Parameters.Clear;

        fBeforeDateTime := Time;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to get manifest cost data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to get manifest cost data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to get manifest cost data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('GET PO', 'SELECTED all manifest cost info', 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          GetSizeInfo
        Else
          LogActLog('ERROR', 'FAILED SELECT all manifest cost info. Err Msg: ' + E.Message + ' Err: ' + E.ClassName, 0)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
end;

function TData_Module.InsertManifestCostInfo: Boolean;
var
  fDescription: String;
Begin
  Result := False;
  fDescription := 'INSERT ' + fAssyCode + ',' + fAssyManifestNo + ',' + fPOstart + ',' + fPOEnd + ',' + FloatToStr(fAssyCost);

  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.INSERT_ManifestCost;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@AssyCode';
        Parameters.ParamValues['@AssyCode'] := fAssyCode;
        Parameters.AddParameter.Name := '@AssyManifestNo';
        Parameters.ParamValues['@AssyManifestNo'] := fAssyManifestNo;
        Parameters.AddParameter.Name := '@StartManifest';
        Parameters.ParamValues['@StartManifest'] := fPOStart;
        Parameters.AddParameter.Name := '@EndManifest';
        Parameters.ParamValues['@EndManifest'] := fPOEnd;
        Parameters.AddParameter.Name := '@AssyCost';
        Parameters.ParamValues['@AssyCost'] := fAssycost;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to insert PO data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to insert PO data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to insert PO data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
          Result := True;
        end;
      End;    //With
      LogActLog('INSERT PO', fDescription, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          InsertSizeInfo
        Else
          LogActLog('ERROR', 'FAILED to INSERT PO ' + fAssyCode + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
end;      //InsertMonthlyPOInfo

procedure TData_Module.UpdateManifestCostInfo;
begin
  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.UPDATE_ManifestCost;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@AssyCode';
        Parameters.ParamValues['@AssyCode'] := fAssyCode;
        Parameters.AddParameter.Name := '@AssyManifestNo';
        Parameters.ParamValues['@AssyManifestNo'] := fAssyManifestNo;
        Parameters.AddParameter.Name := '@StartManifest';
        Parameters.ParamValues['@StartManifest'] := fPOStart;
        Parameters.AddParameter.Name := '@EndManifest';
        Parameters.ParamValues['@EndManifest'] := fPOEnd;
        Parameters.AddParameter.Name := '@AssyCost';
        Parameters.ParamValues['@AssyCost'] := fAssycost;
        Parameters.AddParameter.Name := '@RecordID';
        Parameters.ParamValues['@RecordID'] := fRecordID;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to update PO data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to update PO data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to update PO data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('UPDATE PO', 'UPDATE PO:' + fAssyCode + ',' + fPOstart + ',' + fPOEnd + ',' + fPickUp + ',' + fPONumber + ',' + FloatToStr(fAssyCost), 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          UpdateSizeInfo
        Else
          LogActLog('ERROR', 'FAILED UPDATE PO:' + fAssyCode + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
end;

procedure TData_Module.DeleteManifestCostInfo;
begin
  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.DELETE_ManifestCost;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@RecordID';
        Parameters.ParamValues['@RecordID'] := fRecordID;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to delete manifest cost data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to delete manifest cost data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to delete manifest cost data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('DELETE PO', 'DELETED manifest cost:' + fAssyCode + 'From:' + fPOStart + 'To:' + fPOEnd , 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          DeleteSupplierInfo
        Else
          LogActLog('ERROR', 'FAILED DELETE manifest cost:' + fAssyCode + 'From:' + fPOStart + 'To:' + fPOEnd + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
end;


procedure TData_Module.GetMonthlyPOInfo; // DMV 04/15/2005
begin
  try
    try
      With Inv_DataSet do
      Begin
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'dbo.SELECT_AssyMonthlyPODisplay;1';
        Parameters.Clear;

        fBeforeDateTime := Time;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to get Monthly PO data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to get Monthly PO data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to get Monthly PO data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('GET PO', 'SELECTED all Monthly PO info', 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          GetSizeInfo
        Else
          LogActLog('ERROR', 'FAILED SELECT all Monthly PO info. Err Msg: ' + E.Message + ' Err: ' + E.ClassName, 0)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
end;      //GetMonthlyPOInfo

function TData_Module.InsertMonthlyPOInfo: Boolean;
var
  fDescription: String;
Begin
  Result := False;
  fDescription := 'INSERT ' + fAssyCode + ',' + fPOstart + ',' + fPOEnd + ',' + fPickUp + ',' + fPONumber + ',' + FloatToStr(fAssyCost);

  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.INSERT_AssyMonthlyPO;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@AssyCode';
        Parameters.ParamValues['@AssyCode'] := fAssyCode;
        Parameters.AddParameter.Name := '@POStart';
        Parameters.ParamValues['@POStart'] := fPOStart;
        Parameters.AddParameter.Name := '@POEnd';
        Parameters.ParamValues['@POEnd'] := fPOEnd;
        Parameters.AddParameter.Name := '@PickUp';
        Parameters.ParamValues['@PickUp'] := fPickUp;
        Parameters.AddParameter.Name := '@PickUpTime';
        Parameters.ParamValues['@PickUpTime'] := fPickUpTime;
        Parameters.AddParameter.Name := '@PONumber';
        Parameters.ParamValues['@PONumber'] := fPONumber;
        Parameters.AddParameter.Name := '@AssyCost';
        Parameters.ParamValues['@AssyCost'] := fAssycost;
        Parameters.AddParameter.Name := '@POQty';
        Parameters.ParamValues['@POQty'] := fPOQty;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to insert PO data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to insert PO data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to insert PO data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
          Result := True;
        end;
      End;    //With
      LogActLog('INSERT PO', fDescription, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          InsertSizeInfo
        Else
          LogActLog('ERROR', 'FAILED to INSERT PO ' + fAssyCode + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
end;      //InsertMonthlyPOInfo

procedure TData_Module.UpdateMonthlyPOInfo;
begin
  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.UPDATE_AssyMonthlyPO;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@AssyCode';
        Parameters.ParamValues['@AssyCode'] := fAssyCode;
        Parameters.AddParameter.Name := '@POStart';
        Parameters.ParamValues['@POStart'] := fPOStart;
        Parameters.AddParameter.Name := '@POEnd';
        Parameters.ParamValues['@POEnd'] := fPOEnd;
        Parameters.AddParameter.Name := '@PickUp';
        Parameters.ParamValues['@PickUp'] := fPickUp;
        Parameters.AddParameter.Name := '@PickUpTime';
        Parameters.ParamValues['@PickUpTime'] := fPickUpTime;
        Parameters.AddParameter.Name := '@PONumber';
        Parameters.ParamValues['@PONumber'] := fPONumber;
        Parameters.AddParameter.Name := '@AssyCost';
        Parameters.ParamValues['@AssyCost'] := fAssycost;
        Parameters.AddParameter.Name := '@POQty';
        Parameters.ParamValues['@POQty'] := fPOQty;
        Parameters.AddParameter.Name := '@AssyCodePrev';
        Parameters.ParamValues['@AssyCodePrev'] := fAssyCodePrev;
        Parameters.AddParameter.Name := '@POStartPrev';
        Parameters.ParamValues['@POStartPrev'] := fPOStartPrev;
        Parameters.AddParameter.Name := '@POEndPrev';
        Parameters.ParamValues['@POEndPrev'] := fPOEndPrev;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to update PO data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to update PO data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to update PO data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('UPDATE PO', 'UPDATE PO:' + fAssyCode + ',' + fPOstart + ',' + fPOEnd + ',' + fPickUp + ',' + fPONumber + ',' + FloatToStr(fAssyCost), 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          UpdateSizeInfo
        Else
          LogActLog('ERROR', 'FAILED UPDATE PO:' + fAssyCode + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
end;      //UpdateMonthlyPOInfo

procedure TData_Module.DeleteMonthlyPOInfo;
begin
  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.DELETE_AssyMonthlyPO;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@AssyCode';
        Parameters.ParamValues['@AssyCode'] := fAssyCode;
        Parameters.AddParameter.Name := '@POStart';
        Parameters.ParamValues['@POStart'] := fPOStart;
        Parameters.AddParameter.Name := '@POEnd';
        Parameters.ParamValues['@POEnd'] := fPOEnd;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to delete PO data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to delete PO data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to delete PO data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('DELETE PO', 'DELETED PO:' + fAssyCode + 'From:' + fPOStart + 'To:' + fPOEnd , 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          DeleteSupplierInfo
        Else
          LogActLog('ERROR', 'FAILED DELETE PO:' + fAssyCode + 'From:' + fPOStart + 'To:' + fPOEnd + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
end;      //DeleteMonthlyPOInfo


procedure TData_Module.GetRenbanGroupInfo;
begin
  try
    try
      With Inv_DataSet do
      Begin
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'dbo.SELECT_RenbanGroup;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@RenbanCode';
        Parameters.ParamValues['@RenbanCode'] := '';

        fBeforeDateTime := Time;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to get renban data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to get renban data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to get renban data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('GET REN', 'SELECTED all RENBAN GROUP info', 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          GetSizeInfo
        Else
          LogActLog('ERROR', 'FAILED SELECT all RENBAN GROUP info. Err Msg: ' + E.Message + ' Err: ' + E.ClassName, 0)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
end;

function TData_Module.InsertRenbanGroupInfo: Boolean;
var
  fDescription: String;
Begin
  Result := False;
  fDescription := 'INSERT ' + fRenbancode;

  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.INSERT_RenbanGroup;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@RenbanCode';
        Parameters.ParamValues['@RenbanCode'] := fRenbanCode;
        Parameters.AddParameter.Name := '@RenbanCount';
        Parameters.ParamValues['@RenbanCount'] := fRenbanCount;
        Parameters.AddParameter.Name := '@ShipDays';
        Parameters.ParamValues['@ShipDays'] := fShipDays;
        Parameters.AddParameter.Name := '@ShipDaysMonday';
        Parameters.ParamValues['@ShipDaysMonday'] := fShipDaysMonday;
        Parameters.AddParameter.Name := '@ShipDaysTuesday';
        Parameters.ParamValues['@ShipDaysTuesday'] := fShipDaysTuesday;
        Parameters.AddParameter.Name := '@ShipDaysWednesday';
        Parameters.ParamValues['@ShipDaysWednesday'] := fShipDaysWednesday;
        Parameters.AddParameter.Name := '@ShipDaysThursday';
        Parameters.ParamValues['@ShipDaysThursday'] := fShipDaysThursday;
        Parameters.AddParameter.Name := '@ShipDaysFriday';
        Parameters.ParamValues['@ShipDaysFriday'] := fShipDaysFriday;
        Parameters.AddParameter.Name := '@ShipDaysSaturday';
        Parameters.ParamValues['@ShipDaysSaturday'] := fShipDaysSaturday;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to insert renban data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to insert renban data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to insert renban data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
          Result := True;
        end;
      End;    //With
      LogActLog('INSERT REN', fDescription, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          InsertSizeInfo
        Else
          LogActLog('ERROR', 'FAILED to INSERT Renban Group ' + fAssyCode + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;        //InsertSizeInfo

procedure TData_Module.UpdateRenbanGroupInfo;
begin
  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.UPDATE_RenbanGroup;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@RenbanCode';
        Parameters.ParamValues['@RenbanCode'] := fRenbanCode;
        Parameters.AddParameter.Name := '@RenbanCount';
        Parameters.ParamValues['@RenbanCount'] := fRenbanCount;
        Parameters.AddParameter.Name := '@RenbanID';
        Parameters.ParamValues['@RenbanID'] := fRecordID;
        Parameters.AddParameter.Name := '@ShipDays';
        Parameters.ParamValues['@ShipDays'] := fShipDays;
        Parameters.AddParameter.Name := '@ShipDaysMonday';
        Parameters.ParamValues['@ShipDaysMonday'] := fShipDaysMonday;
        Parameters.AddParameter.Name := '@ShipDaysTuesday';
        Parameters.ParamValues['@ShipDaysTuesday'] := fShipDaysTuesday;
        Parameters.AddParameter.Name := '@ShipDaysWednesday';
        Parameters.ParamValues['@ShipDaysWednesday'] := fShipDaysWednesday;
        Parameters.AddParameter.Name := '@ShipDaysThursday';
        Parameters.ParamValues['@ShipDaysThursday'] := fShipDaysThursday;
        Parameters.AddParameter.Name := '@ShipDaysFriday';
        Parameters.ParamValues['@ShipDaysFriday'] := fShipDaysFriday;
        Parameters.AddParameter.Name := '@ShipDaysSaturday';
        Parameters.ParamValues['@ShipDaysSaturday'] := fShipDaysSaturday;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to update renban data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to update renban data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to update renban data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('UPDATE REN', 'UPDATE Renban Group:' + fAssyCode, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          UpdateSizeInfo
        Else
          LogActLog('ERROR', 'FAILED UPDATE Renban Group:' + fAssyCode + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
end;

procedure TData_Module.DeleteRenbanGroupInfo;
begin
  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.DELETE_RenbanGroup;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@RenbanID';
        Parameters.ParamValues['@RenbanID'] := fRecordID;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to delete renban data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to delete renban data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to delete renban data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('DELETE RENBAN', 'DELETED R:' + fRenbanCode , 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          DeleteSupplierInfo
        Else
          LogActLog('ERROR', 'FAILED DELETE Renban Code R:' + fRenbanCode + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
end;


// DMV 02-12-03
procedure TData_Module.GetForecastDetailInfo;
begin
  try
    try
      With Inv_DataSet do
      Begin
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'dbo.SELECT_ForecastDetail;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@AssyCode';
        Parameters.ParamValues['@AssyCode'] := '';

        fBeforeDateTime := Time;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to get Forecast Detail data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to get Forecast Detail data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to get Forecast Detail data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
          Inv_DataSource.DataSet := Inv_DataSet;
        end;
      End;    //With
      LogActLog('GET FORE', 'SELECTED all FORECAST DETAIL info', 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          GetSizeInfo
        Else
          LogActLog('ERROR', 'FAILED SELECT all FORECAST DETAIL info. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
end;

function TData_Module.InsertForecastDetailInfo: Boolean;
var
  fDescription: String;
Begin
  Result := False;
  fDescription := 'INSERT ' + fAssycode;

  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.INSERT_ForecastDetail;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@AssyCode';
        Parameters.ParamValues['@AssyCode'] := fAssyCode;
        Parameters.AddParameter.Name := '@EffectiveMonth';
        Parameters.ParamValues['@EffectiveMonth'] := fBeginDatestr;
        Parameters.AddParameter.Name := '@TireCode';
        Parameters.ParamValues['@TireCode'] := fTirePartNum1;
        Parameters.AddParameter.Name := '@TireRatio';
        Parameters.ParamValues['@TireRatio'] := fTireRatio1;
        Parameters.AddParameter.Name := '@WheelCode';
        Parameters.ParamValues['@WheelCode'] := fWheelPartNum1;
        Parameters.AddParameter.Name := '@WheelRatio';
        Parameters.ParamValues['@WheelRatio'] := fWheelRatio1;
        Parameters.AddParameter.Name := '@Ratio';
        Parameters.ParamValues['@Ratio'] := fForecastRatio;
        Parameters.AddParameter.Name := '@Kanban';
        Parameters.ParamValues['@Kanban'] := fKanban;
        Parameters.AddParameter.Name := '@BCode';
        Parameters.ParamValues['@BCode'] := fBroadCode;
        Parameters.AddParameter.Name := '@AssyQty';
        Parameters.ParamValues['@AssyQty'] := fQTY;
        Parameters.AddParameter.Name := '@ValveCode';
        Parameters.ParamValues['@ValveCode'] := fValvePartNum;
        Parameters.AddParameter.Name := '@FilmCode';
        Parameters.ParamValues['@FilmCode'] := fFilmPartNum;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to insert Forecast Detail data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to insert Forecast Detail data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to insert Forecast Detail data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
          Result := True;
        end;
      End;    //With
      LogActLog('INSERT FORE', fDescription, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          InsertForecastDetailInfo
        Else
          LogActLog('ERROR', 'FAILED to INSERT Forecast Detail ' + fAssyCode + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;        //InsertSizeInfo

procedure TData_Module.UpdateForecastDetailInfo;
begin
  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.UPDATE_ForecastDetail;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@AssyCode';
        Parameters.ParamValues['@AssyCode'] := fAssyCode;
        Parameters.AddParameter.Name := '@EffectiveMonth';
        Parameters.ParamValues['@EffectiveMonth'] := fBeginDatestr;
        Parameters.AddParameter.Name := '@TireCode';
        Parameters.ParamValues['@TireCode'] := fTirePartNum1;
        Parameters.AddParameter.Name := '@TireRatio';
        Parameters.ParamValues['@TireRatio'] := fTireRatio1;
        Parameters.AddParameter.Name := '@WheelCode';
        Parameters.ParamValues['@WheelCode'] := fWheelPartNum1;
        Parameters.AddParameter.Name := '@WheelRatio';
        Parameters.ParamValues['@WheelRatio'] := fWheelRatio1;
        Parameters.AddParameter.Name := '@Ratio';
        Parameters.ParamValues['@Ratio'] := fForecastRatio;
        Parameters.AddParameter.Name := '@Kanban';
        Parameters.ParamValues['@Kanban'] := fKanban;
        Parameters.AddParameter.Name := '@BCode';
        Parameters.ParamValues['@BCode'] := fBroadCode;
        Parameters.AddParameter.Name := '@AssyQty';
        Parameters.ParamValues['@AssyQty'] := fQTY;
        Parameters.AddParameter.Name := '@ValveCode';
        Parameters.ParamValues['@ValveCode'] := fValvePartNum;
        Parameters.AddParameter.Name := '@FilmCode';
        Parameters.ParamValues['@FilmCode'] := fFilmPartNum;
        Parameters.AddParameter.Name := '@RecordID';
        Parameters.ParamValues['@RecordID'] := fRecordID;                                                   

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to update Forecast Detail data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to update Forecast Detail data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to update Forecast Detail data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('UPDATE FORE', 'UPDATE Forecast Detail:' + fAssyCode, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          UpdateForecastDetailInfo
        Else
          LogActLog('ERROR', 'FAILED UPDATE Forecast Detail:' + fAssyCode + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
end;

procedure TData_Module.DeleteForecastDetailInfo;
begin
  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.DELETE_ForecastDetail;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@RecordID';
        Parameters.ParamValues['@RecordID'] := fRecordID;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to delete Forecast Detail data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to delete Forecast Detail data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to delete Forecast Detail data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('DELETE FORE', 'DELETED FORECAST DETAIL Info:' + fSizeCode, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          DeleteForecastDetailInfo
        Else
          LogActLog('ERROR', 'FAILED DELETE FORECAST DETAIL Info S:' + fSizeCode + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
end;
// DMV 02-12-03


procedure TData_Module.GetSizeInfo;
Begin
  try
    try
      With Inv_DataSet do
      Begin
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'dbo.SELECT_SizeInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@SizeCode';
        Parameters.ParamValues['@SizeCode'] := '';

        fBeforeDateTime := Time;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to get Size data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to get Size data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to get Size data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('GET SIZE', 'SELECTED all SIZE info', 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          GetSizeInfo
        Else
          LogActLog('ERROR', 'FAILED SELECT all SIZE info. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;        //GetSizeInfo

function TData_Module.InsertSizeInfo: Boolean;
var
  fDescription: String;
Begin
  Result := False;
  fDescription := 'INSERT ' + fSizeCode;

  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.SELECT_AssyRatioInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@SizeCode';
        Parameters.ParamValues['@SizeCode'] := fSizeCode;
        Open;

        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to get Size dup data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to get Size dup data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to get Size dup data');
        end
        else
        begin
          If RecordCount = 0 Then
          Begin
            Close;
            ProcedureName := 'dbo.INSERT_SizeInfo;1';
            Parameters.Clear;
            Parameters.AddParameter.Name := '@SizeCode';
            Parameters.ParamValues['@SizeCode'] := fSizeCode;
            Parameters.AddParameter.Name := '@SizeName';
            Parameters.ParamValues['@SizeName'] := fSizeName;
            Parameters.AddParameter.Name := '@Usage';
            Parameters.ParamValues['@Usage'] := fDailyUsage;
            Parameters.AddParameter.Name := '@Safety';
            Parameters.ParamValues['@Safety'] := fSafetyDays;

            fBeforeDateTime := Time;
            ExecProc;
            if Inv_Connection.Errors.Count > 0 then
            begin
              ShowMessage('Unable to insert Size data, '+Inv_Connection.Errors.Item[0].Get_Description);
              LogActLog('ERROR','Unable to insert Size data, '+Inv_Connection.Errors.Item[0].Get_Description);
              raise EDatabaseError.Create('Unable to insert Size data');
            end
            else
            begin
              fAfterDateTime := Time;
              fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
              Result := True;
            end;
          end
          else
              fDescription := 'FAILED to ' + fDescription + ' (DUPLICATE)';
        end;
      End;    //With
      LogActLog('INSERT SIZ', fDescription, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          InsertSizeInfo
        Else
          LogActLog('ERROR', 'FAILED to INSERT SIZE ' + fSizeCode + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;        //InsertSizeInfo

procedure TData_Module.UpdateSizeInfo;
Begin
  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.UPDATE_SizeInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@SizeCode';
        Parameters.ParamValues['@SizeCode'] := fSizeCode;
        Parameters.AddParameter.Name := '@SizeName';
        Parameters.ParamValues['@SizeName'] := fSizeName;
        Parameters.AddParameter.Name := '@Usage';
        Parameters.ParamValues['@Usage'] := fDailyUsage;
        Parameters.AddParameter.Name := '@Safety';
        Parameters.ParamValues['@Safety'] := fSafetyDays;
        Parameters.AddParameter.Name := '@SizeID';
        Parameters.ParamValues['@SizeID'] := FRecordID;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to update Size data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to update Size data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to update Size data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('UPDATE SIZ', 'UPDATE Size:' + fSizeCode, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          UpdateSizeInfo
        Else
          LogActLog('ERROR', 'FAILED UPDATE Size:' + fSizeCode + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;        //UpdateSizeInfo

procedure TData_Module.DeleteSizeInfo;
Begin
  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.DELETE_SizeInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@SizeID';
        Parameters.ParamValues['@SizeID'] := fRecordID;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to delete Size data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to delete Size data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to delete Size data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('DELETE SIZ', 'DELETED Size Info:' + fSizeCode, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          DeleteSizeInfo
        Else
          LogActLog('ERROR', 'FAILED DELETE Size Info S:' + fSizeCode + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;        //DeleteSizeInfo

procedure TData_Module.GetBCRatioInfo;
begin
  try
    try
      With Inv_DataSet do
      Begin
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'dbo.SELECT_BCRatioInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@BroadCode';
        Parameters.ParamValues['@BroadCode'] := '';

        fBeforeDateTime := Time;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to get AssyRatio data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to get AssyRatio data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to get AssyRatio data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
          Inv_DataSource.DataSet := Inv_DataSet;
        end;
      End;    //With
      LogActLog('GET ASSY', 'SELECTED all ASSY/RATIO info', 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          GetAssyRatioInfo
        Else
        begin
          LogActLog('ERROR', 'FAILED SELECT all ASSY/RATIO info. Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
          ShowMessage('FAILED SELECT all ASSY/RATIO info. Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
        end;
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
end;

function TData_Module.InsertBCRatioInfo:boolean;
begin
  result:=TRUE;
end;

procedure TData_Module.UpdateBCRatioInfo;
begin
end;

procedure TData_Module.DeleteBCRatioInfo;
begin
end;

procedure TData_Module.GetAssyRatioInfo;
Begin
  try
    try
      With Inv_DataSet do
      Begin
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'dbo.SELECT_AssyRatioInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@BroadCode';
        Parameters.ParamValues['@BroadCode'] := '';

        fBeforeDateTime := Time;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to get AssyRatio data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to get AssyRatio data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to get AssyRatio data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
          Inv_DataSource.DataSet := Inv_DataSet;
        end;
      End;    //With
      LogActLog('GET ASSY', 'SELECTED all ASSY/RATIO info', 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          GetAssyRatioInfo
        Else
        begin
          LogActLog('ERROR', 'FAILED SELECT all ASSY/RATIO info. Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
          ShowMessage('FAILED SELECT all ASSY/RATIO info. Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
        end;
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;        //GetAssyRatioInfo

function TData_Module.InsertAssyRatioInfo: Boolean;
var
  fDescription: String;
Begin
  Result := False;
  fDescription := 'INSERT ' + fBroadCode;

  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.SELECT_AssyRatioInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@BroadCode';
        Parameters.ParamValues['@BroadCode'] := fBroadCode;
        Open;

        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to get dup check, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to get dup check, '+Inv_Connection.Errors.Item[0].Get_Description);
        end
        else
        begin
          If RecordCount = 0 Then
          Begin
            Close;
            ProcedureName := 'dbo.INSERT_AssyRatioInfo;1';
            Parameters.Clear;
            Parameters.AddParameter.Name := '@AssyCode';
            Parameters.ParamValues['@AssyCode'] := fAssyCode;
            Parameters.AddParameter.Name := '@AssyName';
            Parameters.ParamValues['@AssyName'] := fAssyName;
            Parameters.AddParameter.Name := '@BroadCode';
            Parameters.ParamValues['@BroadCode'] := fBroadCode;
            Parameters.AddParameter.Name := '@TireQTY';
            Parameters.ParamValues['@TireQTY'] := fTireQTY;
  //11-14-2002
            Parameters.AddParameter.Name := '@TirePartNum1';
            Parameters.ParamValues['@TirePartNum1'] := fTirePartNum1;
            Parameters.AddParameter.Name := '@TireRatio1';
            Parameters.ParamValues['@TireRatio1'] := fTireRatio1;
            Parameters.AddParameter.Name := '@TirePartNum2';
            Parameters.ParamValues['@TirePartNum2'] := fTirePartNum2;
            Parameters.AddParameter.Name := '@TireRatio2';
            Parameters.ParamValues['@TireRatio2'] := fTireRatio2;
            Parameters.AddParameter.Name := '@TirePartNum3';
            Parameters.ParamValues['@TirePartNum3'] := fTirePartNum3;
            Parameters.AddParameter.Name := '@TireRatio3';
            Parameters.ParamValues['@TireRatio3'] := fTireRatio3;
            Parameters.AddParameter.Name := '@WheelQty';
            Parameters.ParamValues['@WheelQty'] := fWheelQty;
            Parameters.AddParameter.Name := '@WheelPartNum1';
            Parameters.ParamValues['@WheelPartNum1'] := fWheelPartNum1;
            Parameters.AddParameter.Name := '@WheelRatio1';
            Parameters.ParamValues['@WheelRatio1'] := fWheelRatio1;
            Parameters.AddParameter.Name := '@WheelPartNum2';
            Parameters.ParamValues['@WheelPartNum2'] := fWheelPartNum2;
            Parameters.AddParameter.Name := '@WheelRatio2';
            Parameters.ParamValues['@WheelRatio2'] := fWheelRatio2;
            Parameters.AddParameter.Name := '@WheelPartNum3';
            Parameters.ParamValues['@WheelPartNum3'] := fWheelPartNum3;
            Parameters.AddParameter.Name := '@WheelRatio3';
            Parameters.ParamValues['@WheelRatio3'] := fWheelRatio3;
            Parameters.AddParameter.Name := '@SpareTireQty';
            Parameters.ParamValues['@SpareTireQty'] := fSpareTireQty;
            Parameters.AddParameter.Name := '@SpareTirePartNum';
            Parameters.ParamValues['@SpareTirePartNum'] := fSpareTirePartNum;
            Parameters.AddParameter.Name := '@SpareWheelPartNum';
            Parameters.ParamValues['@SpareWheelPartNum'] := fSpareWheelPartNum;

            fBeforeDateTime := Time;
            ExecProc;
            if Inv_Connection.Errors.Count > 0 then
            begin
              ShowMessage('Unable to insert AssyRatio, '+Inv_Connection.Errors.Item[0].Get_Description);
              LogActLog('ERROR','Unable to insert AssyRatio, '+Inv_Connection.Errors.Item[0].Get_Description);
              raise EDatabaseError.Create('Unable to insert AssyRatio');
            end
            else
            begin
              fAfterDateTime := Time;
              fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
              Result := True;
            end;
          End
          Else
            fDescription := 'FAILED to ' + fDescription + ' (DUPLICATE)';
        end;
      End;    //With
      LogActLog('INSERT ASY', fDescription, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          InsertAssyRatioInfo
        Else
        begin
          LogActLog('ERROR', 'FAILED to INSERT ASSY/RATIO ' + fBroadCode + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
          ShowMessage('FAILED to INSERT ASSY/RATIO ' + fBroadCode + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
        end;
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;        //InsertAssyRatioInfo

procedure TData_Module.UpdateAssyRatioInfo;
Begin
  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.UPDATE_AssyRatioInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@AssyCode';
        Parameters.ParamValues['@AssyCode'] := fAssyCode;
        Parameters.AddParameter.Name := '@AssyName';
        Parameters.ParamValues['@AssyName'] := fAssyName;
        Parameters.AddParameter.Name := '@BroadCode';
        Parameters.ParamValues['@BroadCode'] := fBroadCode;
        Parameters.AddParameter.Name := '@TireQTY';
        Parameters.ParamValues['@TireQTY'] := fTireQTY;
        Parameters.AddParameter.Name := '@TirePartNum1';
        Parameters.ParamValues['@TirePartNum1'] := fTirePartNum1;
        Parameters.AddParameter.Name := '@TireRatio1';
        Parameters.ParamValues['@TireRatio1'] := fTireRatio1;
        Parameters.AddParameter.Name := '@TirePartNum2';
        Parameters.ParamValues['@TirePartNum2'] := fTirePartNum2;
        Parameters.AddParameter.Name := '@TireRatio2';
        Parameters.ParamValues['@TireRatio2'] := fTireRatio2;
        Parameters.AddParameter.Name := '@TirePartNum3';
        Parameters.ParamValues['@TirePartNum3'] := fTirePartNum3;
        Parameters.AddParameter.Name := '@TireRatio3';
        Parameters.ParamValues['@TireRatio3'] := fTireRatio3;
        Parameters.AddParameter.Name := '@WheelQty';
        Parameters.ParamValues['@WheelQty'] := fWheelQty;
        Parameters.AddParameter.Name := '@WheelPartNum1';
        Parameters.ParamValues['@WheelPartNum1'] := fWheelPartNum1;
        Parameters.AddParameter.Name := '@WheelRatio1';
        Parameters.ParamValues['@WheelRatio1'] := fWheelRatio1;
        Parameters.AddParameter.Name := '@WheelPartNum2';
        Parameters.ParamValues['@WheelPartNum2'] := fWheelPartNum2;
        Parameters.AddParameter.Name := '@WheelRatio2';
        Parameters.ParamValues['@WheelRatio2'] := fWheelRatio2;
        Parameters.AddParameter.Name := '@WheelPartNum3';
        Parameters.ParamValues['@WheelPartNum3'] := fWheelPartNum3;
        Parameters.AddParameter.Name := '@WheelRatio3';
        Parameters.ParamValues['@WheelRatio3'] := fWheelRatio3;
        Parameters.AddParameter.Name := '@SpareTireQty';
        Parameters.ParamValues['@SpareTireQty'] := fSpareTireQty;
        Parameters.AddParameter.Name := '@SpareTirePartNum';
        Parameters.ParamValues['@SpareTirePartNum'] := fSpareTirePartNum;
        Parameters.AddParameter.Name := '@SpareWheelPartNum';
        Parameters.ParamValues['@SpareWheelPartNum'] := fSpareWheelPartNum;
        Parameters.AddParameter.Name := '@BroadCodePrev';
        Parameters.ParamValues['@BroadCodePrev'] := fBroadCodePrev;


        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to update AssyRatio data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to update AssyRatio data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to update AssyRatio data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end
      End;    //With
      LogActLog('UPDATE ASY', 'UPDATE ASSY:' + fBroadCode, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          UpdateAssyRatioInfo
        Else
        begin
          LogActLog('ERROR', 'FAILED UPDATE ASSY:' + fBroadCode + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
          ShowMessage('FAILED UPDATE ASSY:' + fBroadCode + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
        end;
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;        //UpdateAssyRatioInfo

procedure TData_Module.DeleteAssyRatioInfo;
Begin
  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.DELETE_AssyRatioInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@BroadCode';
        Parameters.ParamValues['@BroadCode'] := fBroadCode;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to delete AssyRatio data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to delete AssyRatio data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to delete AssyRatio data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('DELETE ASY', 'DELETED ' + fBroadCode, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          DeleteAssyRatioInfo
        Else
        begin
          LogActLog('ERROR', 'FAILED DELETE ASSY:' + fBroadCode + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
          ShowMessage('FAILED DELETE ASSY:' + fBroadCode + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
        end;
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;        //DeleteAssyRatioInfo


procedure TData_Module.DeleteRecConfStatInfo;
Begin
  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.DELETE_RecConfStatInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@PartCode';
        Parameters.ParamValues['@PartCode'] := fPartNum;
        Parameters.AddParameter.Name := '@FRSNo';
        Parameters.ParamValues['@FRSNo'] := fFRSNo;
        Parameters.AddParameter.Name := '@Renban';
        Parameters.ParamValues['@Renban'] := fRenban;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to delete RecConfStat data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to delete RecConfStat data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to delete RecConfStat data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('DELETE REC', 'DELETED ' + fBroadCode, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          DeleteAssyRatioInfo
        Else
        begin
          LogActLog('ERROR', 'FAILED DELETE RECCONF:' + fBroadCode + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
          ShowMessage('FAILED DELETE ASSY:' + fBroadCode + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
        end;
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;        //DeleteAssyRatioInfo


procedure TData_Module.GetRecConfStatInfo;
Begin
  try
    try
      With Inv_DataSet do
      Begin
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'dbo.SELECT_RecConfStatInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@SupCode';
        Parameters.ParamValues['@SupCode'] := '';
        Parameters.AddParameter.Name := '@PartCode';
        Parameters.ParamValues['@PartCode'] := '';
        Parameters.AddParameter.Name := '@FRSNo';
        Parameters.ParamValues['@FRSNo'] := '';
        Parameters.AddParameter.Name := '@Renban';
        Parameters.ParamValues['@Renban'] := '';

        fBeforeDateTime := Time;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to get RecConf data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to get RecConf AssyRatio data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to get RecConf AssyRatio data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('GET RCS', 'SELECTED all RecConfStat info', 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          GetRecConfStatInfo
        Else
        begin
          LogActLog('ERROR', 'FAILED SELECT all RecConfStat info. Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
          ShowMessage('FAILED SELECT all RecConfStat info. Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
        end;
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;        //GetRecConfStatInfo

function TData_Module.InsertRecConfStatInfo: Boolean;
var
  fDescription: String;
Begin
  Result := False;
  fDescription := 'INSERT RCS S:' + fSupCode + ' P:' + fPartNum + ' F:' + fFRSNo + ' R:' + fRenban + ' Q:' + IntToStr(fQty);

  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.SELECT_RecConfStatInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@SupCode';
        Parameters.ParamValues['@SupCode'] := fSupCode;
        Parameters.AddParameter.Name := '@PartCode';
        Parameters.ParamValues['@PartCode'] := fPartNum;
        Parameters.AddParameter.Name := '@FRSNo';
        Parameters.ParamValues['@FRSNo'] := fFRSNo;
        Parameters.AddParameter.Name := '@Renban';
        Parameters.ParamValues['@Renban'] := fRenban;
        Open;

        If RecordCount = 0 Then
        Begin
          Close;
          ProcedureName := 'dbo.INSERT_RecConfStatInfo;1';
          Parameters.Clear;
          Parameters.AddParameter.Name := '@SupCode';
          Parameters.ParamValues['@SupCode'] := fSupCode;
          Parameters.AddParameter.Name := '@PartCode';
          Parameters.ParamValues['@PartCode'] := fPartNum;
          Parameters.AddParameter.Name := '@FRSNo';
          Parameters.ParamValues['@FRSNo'] := fFRSNo;
          Parameters.AddParameter.Name := '@Renban';
          Parameters.ParamValues['@Renban'] := fRenban;
//  12/17/2002  --AH
          Parameters.AddParameter.Name := '@Qty';
          Parameters.ParamValues['@Qty'] := fQty;
          Parameters.AddParameter.Name := '@StatusSupShip';
          Parameters.ParamValues['@StatusSupShip'] := fInTransit;
          Parameters.AddParameter.Name := '@Arrival';
          Parameters.ParamValues['@Arrival'] := fArrival;
          Parameters.AddParameter.Name := '@TrailerNo';
          Parameters.ParamValues['@TrailerNo'] := fTrailerNo;
          Parameters.AddParameter.Name := '@PlantYard';
          Parameters.ParamValues['@PlantYard'] := fPlantYard;
          Parameters.AddParameter.Name := '@PlantParking';
          Parameters.ParamValues['@PlantParking'] := fParkingSpot;
          Parameters.AddParameter.Name := '@AssemblerYard';
          Parameters.ParamValues['@AssemblerYard'] := fAssemblerYard;
          Parameters.AddParameter.Name := '@AssemblerLocation';
          Parameters.ParamValues['@AssemblerLocation'] := fAssemblerLocation;
          Parameters.AddParameter.Name := '@EmptyTrailer';
          Parameters.ParamValues['@EmptyTrailer'] := fEmptyTrailer;
          Parameters.AddParameter.Name := '@Detention';
          Parameters.ParamValues['@Detention'] := fDetention;
          Parameters.AddParameter.Name := '@Warehouse';
          Parameters.ParamValues['@Warehouse'] := fWarehouse;
          Parameters.AddParameter.Name := '@Terminated';
          Parameters.ParamValues['@Terminated'] := fTerminated;
          Parameters.AddParameter.Name := '@Order';
          Parameters.ParamValues['@Order'] := fOrder;

          fBeforeDateTime := Time;
          ExecProc;
          if Inv_Connection.Errors.Count > 0 then
          begin
            ShowMessage('Unable to insert RecConf data, '+Inv_Connection.Errors.Item[0].Get_Description);
            LogActLog('ERROR','Unable to insert RecConf AssyRatio data, '+Inv_Connection.Errors.Item[0].Get_Description);
            raise EDatabaseError.Create('Unable to insert RecConf AssyRatio data');
          end
          else
          begin
            fAfterDateTime := Time;
            fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
            Result := True;
          end;
        End
        Else
          fDescription := 'FAILED to ' + fDescription + ' (DUPLICATE)';
      End;    //With
      LogActLog('INSERT RCS', fDescription, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          InsertRecConfStatInfo
        Else
        begin
          LogActLog('ERROR', 'FAILED to INSERT RecConfStat S:' + fSizeCode + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
          ShowMessage('FAILED to INSERT RecConfStat S:' + fSizeCode + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
        end;
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;        //InsertRecConfStatInfo

procedure TData_Module.UpdateRecConfStatRenbanInfo;
begin
  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.UPDATE_RecConfStatRenbanInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@Renban';
        Parameters.ParamValues['@Renban'] := fRenban;
//  12/17/2002  --AH
        Parameters.AddParameter.Name := '@Qty';
        Parameters.ParamValues['@Qty'] := fQty;
        Parameters.AddParameter.Name := '@StatusSupShip';
        Parameters.ParamValues['@StatusSupShip'] := fInTransit;
        Parameters.AddParameter.Name := '@Arrival';
        Parameters.ParamValues['@Arrival'] := fArrival;
        Parameters.AddParameter.Name := '@TrailerNo';
        Parameters.ParamValues['@TrailerNo'] := fTrailerNo;
        Parameters.AddParameter.Name := '@PlantYard';
        Parameters.ParamValues['@PlantYard'] := fPlantYard;
        Parameters.AddParameter.Name := '@PlantParking';
        Parameters.ParamValues['@PlantParking'] := fParkingSpot;
        Parameters.AddParameter.Name := '@AssemblerYard';
        Parameters.ParamValues['@AssemblerYard'] := fAssemblerYard;
        Parameters.AddParameter.Name := '@AssemblerLocation';
        Parameters.ParamValues['@AssemblerLocation'] := fAssemblerLocation;
        Parameters.AddParameter.Name := '@EmptyTrailer';
        Parameters.ParamValues['@EmptyTrailer'] := fEmptyTrailer;
        Parameters.AddParameter.Name := '@Detention';
        Parameters.ParamValues['@Detention'] := fDetention;
        Parameters.AddParameter.Name := '@Warehouse';
        Parameters.ParamValues['@Warehouse'] := fWarehouse;
        Parameters.AddParameter.Name := '@Terminated';
        Parameters.ParamValues['@Terminated'] := fTerminated;
        Parameters.AddParameter.Name := '@Order';
        Parameters.ParamValues['@Order'] := fOrder;
        Parameters.AddParameter.Name := '@Kanban';
        Parameters.ParamValues['@Kanban'] := fKanban;
        Parameters.AddParameter.Name := '@Ship';
        Parameters.ParamValues['@Ship'] := fShip;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to update RecConf RENBAN data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to update RecConf AssyRatio RENBAN data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to update RecConf AssyRatio RENBAN data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('UPDATE RCS', 'UPDATE RCS RENBAN S:' + fSupCode + ' P:' + fPartNum + ' F:' + fFRSNo + ' R:' + fRenban + ' Q:'+IntToStr(fQty), 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          UpdateRecConfStatInfo
        Else
        begin
          LogActLog('ERROR', 'FAILED UPDATE RCS RENBAN Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
          ShowMessage('FAILED UPDATE RCS RENBAN Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
        end;
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
end;

procedure TData_Module.UpdateRecConfStatInfo;
Begin
  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.UPDATE_RecConfStatInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@SupCode';
        Parameters.ParamValues['@SupCode'] := fSupCode;
        Parameters.AddParameter.Name := '@PartCode';
        Parameters.ParamValues['@PartCode'] := fPartNum;
        Parameters.AddParameter.Name := '@FRSNo';
        Parameters.ParamValues['@FRSNo'] := fFRSNo;
        Parameters.AddParameter.Name := '@Renban';
        Parameters.ParamValues['@Renban'] := fRenban;
//  12/17/2002  --AH
        Parameters.AddParameter.Name := '@Qty';
        Parameters.ParamValues['@Qty'] := fQty;
        Parameters.AddParameter.Name := '@StatusSupShip';
        Parameters.ParamValues['@StatusSupShip'] := fInTransit;
        Parameters.AddParameter.Name := '@Arrival';
        Parameters.ParamValues['@Arrival'] := fArrival;
        Parameters.AddParameter.Name := '@TrailerNo';
        Parameters.ParamValues['@TrailerNo'] := fTrailerNo;
        Parameters.AddParameter.Name := '@PlantYard';
        Parameters.ParamValues['@PlantYard'] := fPlantYard;
        Parameters.AddParameter.Name := '@PlantParking';
        Parameters.ParamValues['@PlantParking'] := fParkingSpot;
        Parameters.AddParameter.Name := '@AssemblerYard';
        Parameters.ParamValues['@AssemblerYard'] := fAssemblerYard;
        Parameters.AddParameter.Name := '@AssemblerLocation';
        Parameters.ParamValues['@AssemblerLocation'] := fAssemblerLocation;
        Parameters.AddParameter.Name := '@EmptyTrailer';
        Parameters.ParamValues['@EmptyTrailer'] := fEmptyTrailer;
        Parameters.AddParameter.Name := '@Detention';
        Parameters.ParamValues['@Detention'] := fDetention;
        Parameters.AddParameter.Name := '@Warehouse';
        Parameters.ParamValues['@Warehouse'] := fWarehouse;
        Parameters.AddParameter.Name := '@Terminated';
        Parameters.ParamValues['@Terminated'] := fTerminated;
        Parameters.AddParameter.Name := '@Order';
        Parameters.ParamValues['@Order'] := fOrder;
        Parameters.AddParameter.Name := '@Kanban';
        Parameters.ParamValues['@Kanban'] := fKanban;
        Parameters.AddParameter.Name := '@Ship';
        Parameters.ParamValues['@Ship'] := fShip;

        Parameters.AddParameter.Name := '@SupCodePrev';
        Parameters.ParamValues['@SupCodePrev'] := fSupplierCodePrev;
        Parameters.AddParameter.Name := '@PartNumPrev';
        Parameters.ParamValues['@PartNumPrev'] := fPartNumPrev;
        Parameters.AddParameter.Name := '@FRSNoPrev';
        Parameters.ParamValues['@FRSNoPrev'] := fFRSNoPrev;
        Parameters.AddParameter.Name := '@RenbanCodePrev';
        Parameters.ParamValues['@RenbanCodePrev'] := fRenbanCodePrev;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to update RecConf data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to update RecConf AssyRatio data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to update RecConf AssyRatio data');
        end
        else
        begin
          fAfterDateTime := Time;
         fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('UPDATE RCS', 'UPDATE RCS S:' + fSupCode + ' P:' + fPartNum + ' F:' + fFRSNo + ' R:' + fRenban + ' Q:'+IntToStr(fQty), 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          UpdateRecConfStatInfo
        Else
        begin
          LogActLog('ERROR', 'FAILED UPDATE RCS Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
          ShowMessage('FAILED UPDATE RCS Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
        end;
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;        //UpdateRecConfStatInfo

procedure TData_Module.GetStocktakingInfo;
Begin
  try
    try
      With Inv_DataSet do
      Begin
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'dbo.SELECT_StocktakingInfo;1';
        Parameters.Clear;

        fBeforeDateTime := Time;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to get Stocktaking data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to get Stocktaking data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to get Stocktaking data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('GET StkTak', 'SELECTED all Stocktaking info', 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          GetStocktakingInfo
        Else
          LogActLog('ERROR', 'FAILED SELECT all Stocktaking info. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    fErrorCount := 0;
  end;
End;        //GetStocktakingInfo

function TData_Module.InsertStocktakingInfo: Boolean;
var
  fDescription: String;
Begin
  Result := False;
  fDescription := 'INSERT StkTak S:' + fSupCode + ' P:' + fPartNum + ' Q:'+IntToStr(fQty);

  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.INSERT_StocktakingInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@PartCode';
        Parameters.ParamValues['@PartCode'] := fPartNum;
        Parameters.AddParameter.Name := '@QTY';
        Parameters.ParamValues['@QTY'] := fQTY;
        Parameters.AddParameter.Name := '@Reason';
        Parameters.ParamValues['@Reason'] := fComments;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to insert Stocktaking data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to insert Stocktaking data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to insert Stocktaking data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
          Result := True;
        end;
      End;    //With
      LogActLog('INS StkTak', fDescription, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          InsertStocktakingInfo
        Else
          LogActLog('ERROR', 'FAILED to INSERT StkTak S:' + fSupCode + ' P:' + fPartNum + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;        //InsertStocktakingInfo

procedure TData_Module.UpdateStocktakingInfo;
var
  fDescription: String;
Begin
  fDescription := 'UPDATE StkTak S:' + fSupCode + ' P:' + fPartNum + ' Q:'+IntToStr(fQty);

  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.UPDATE_StocktakingInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@SupCode';
        Parameters.ParamValues['@SupCode'] := fSupCode;
        Parameters.AddParameter.Name := '@PartCode';
        Parameters.ParamValues['@PartCode'] := fPartNum;
        Parameters.AddParameter.Name := '@QTY';
        Parameters.ParamValues['@QTY'] := fQTY;
        Parameters.AddParameter.Name := '@Reason';
        Parameters.ParamValues['@Reason'] := fComments;
        Parameters.AddParameter.Name := '@StocktakingID';
        Parameters.ParamValues['@StocktakingID'] := fRecordID;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to update Stocktaking data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to update Stocktaking data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to update Stocktaking data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('UPD StkTak', fDescription, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          UpdateStocktakingInfo
        Else
          LogActLog('ERROR', 'FAILED ' + fDescription + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    fErrorCount := 0;
  end;
End;        //UpdateStocktakingInfo

procedure TData_Module.DeleteStocktakingInfo;
Begin
  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.DELETE_StocktakingInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@StocktakingID';
        Parameters.ParamValues['@StocktakingID'] := fRecordID;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to delete Stocktaking data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to delete Stocktaking data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to delete Stocktaking data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('DEL StkTak', 'DELETED S: ' + fSupCode + ' P: ' + fPartNum, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          DeleteStocktakingInfo
        Else
          LogActLog('ERROR', 'FAILED DELETE StkTak:S: ' + fSupCode + ' P: ' + fPartNum + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;        //DeleteStocktakingInfo

procedure TData_Module.GetPartsListCount;
begin
  try
    try
      With Inv_DataSet do
      Begin
        // Get last processed
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'SELECT_PartsDailyLinePullCount;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@AssyLine';
        Parameters.ParamValues['@AssyLine'] := fAssyCode;
        Parameters.AddParameter.Name := '@Date';
        Parameters.ParamValues['@Date'] := fProductionDate;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to get daily line pull count data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to get daily line pull count data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to get daily line pull count data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;

      End;    //With
      LogActLog('GET PROD', 'SELECTED GetPartsListCount info', 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          GetRecConfStatInfo
        Else
          LogActLog('ERROR', 'FAILED SELECTED GetPartsListCount info. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    fErrorCount := 0;
  end;

end;

procedure TData_Module.GetPartsList;
begin
  try
    try
      With Inv_DataSet do
      Begin
        // Get last processed
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'SELECT_PartsDailyLinePull;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@AssyLine';
        Parameters.ParamValues['@AssyLine'] := fLineName;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to get daily line pull data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to get daily line pull data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to get daily line pull data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;

      End;    //With
      LogActLog('GET PROD', 'SELECTED GetPartsList info', 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          GetRecConfStatInfo
        Else
          LogActLog('ERROR', 'FAILED SELECTED GetPartsList info. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    fErrorCount := 0;
  end;

end;

procedure TData_Module.GetNextASNDate;
var
  prod:string;
  done:boolean;
begin
  try
    with Data_Module.Inv_DataSet do
    begin
      Close;
      CommandType := CmdStoredProc;
      CommandText := 'SELECT_ASNMax;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@LineName';
      Parameters.ParamValues['@LineName'] := fLineName;
      Open;
      if Inv_Connection.Errors.Count > 0 then
      begin
        ShowMessage('Failed on ASNMAX select, '+Inv_Connection.Errors.Item[0].Get_Description);
        DAta_Module.LogActLog('ERROR','Unable to get ASNMax information, '+Inv_Connection.Errors.Item[0].Get_Description);
        raise EDatabaseError.Create('Unable to get ASNMax information, '+Inv_Connection.Errors.Item[0].Get_Description);
      end;
      if not fieldbyname('prod').IsNull then
      begin

        prod:=fieldbyname('prod').AsString;
        NT.NummiTime:=fieldbyname('prod').AsString+'00000000';  //get real

        fEndDate:=        ;
        with ALC_Dataset do
        begin
          Close;
          CommandType := CmdStoredProc;
          CommandText := 'dbo.AD_GetNextASN;1';
          Parameters.Clear;
          Parameters.AddParameter.Name := '@BeginDate';
          Parameters.ParamValues['@BeginDate'] := fEndDate;
          Parameters.AddParameter.Name := '@Line';
          Parameters.ParamValues['@Line'] := fLineName;
          Open;
          if Inv_Connection.Errors.Count > 0 then
          begin
            fStartSeq:='-1';
            ShowMessage('Unable to get production seq, '+Inv_Connection.Errors.Item[0].Get_Description);
            LogActLog('ERROR','Unable to get production seq, '+Inv_Connection.Errors.Item[0].Get_Description);
            raise EDatabaseError.Create('Unable to get production seq');
          end
          else
          begin
            fStartSeq:=FieldByName('ASN').AsString;
            fBeginDateStr:=FieldByName('DateCreated').AsString;
            fBeginDate:=FieldByName('DateCreated').AsDateTime;
          end;

          // Get future holiday and overtime
          Close;
          CommandType := CmdStoredProc;
          CommandText := 'dbo.AD_GetSpecialDate';
          Parameters.Clear;
          Parameters.AddParameter.Name := '@BeginDate';
          Parameters.ParamValues['@BeginDate'] := NT.Time;
          Parameters.AddParameter.Name := '@EndDate';
          Parameters.ParamValues['@EndDate'] := NT.Time+60;
          Parameters.AddParameter.Name := '@LineName';
          Parameters.ParamValues['@LineName'] := fLineName;
          Open;
          if DAta_Module.Inv_Connection.Errors.Count > 0 then
          begin
            ShowMessage('Unable to get production date data, '+DAta_Module.Inv_Connection.Errors.Item[0].Get_Description);
            DAta_Module.LogActLog('ERROR','Unable to get production date data, '+DAta_Module.Inv_Connection.Errors.Item[0].Get_Description);
            raise EDatabaseError.Create('Unable to get production date data');
          end
          else
          begin
            fAfterDateTime := Time;
            fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
          end;





          // Start with the next day
          done:=FALSE;
          repeat
            NT.Time:=NT.Time+1;
            if DayOfTheWeek(NT.Time) = 6 then
            begin
              // weekend check for overtime
              First;
              while not EOF do
              begin
                if NT.Time = fieldbyname('DATE').AsDateTime then
                begin
                  if fieldbyname('Date Status Abrv').AsString = 'O' then
                  begin
                    // work this day
                    done:=TRUE;
                  end;
                end;
                next;
              end;
            end
            else if DayOfTheWeek(NT.Time) < 6 then
            begin
              // weekday check for holiday
              First;
              while not EOF do
              begin
                if NT.Time = fieldbyname('DATE').AsDateTime then
                begin
                  if fieldbyname('Date Status Abrv').AsString = 'H' then
                  begin
                    // work this day
                    break;
                  end;
                end;
                next;
              end;
              if eof then
                done:=TRUE;
            end;
          until done = TRUE;

        end;
        fProductionDateDT:=NT.Time;
      end
      else
      begin
        fStartSeq:='-1';
        fProductionDateDT:=now;
      end;

    end;
  except
    on e:exception do
    begin
      Data_Module.LogActLog('ERROR','Failed on ASN select, '+e.Message);
      ShowMessage('Failed on ASN select, '+e.Message);
      fProductionDateDT:=now;
    end;
  end;
end;

procedure TData_Module.GetNextProductionDate;
var
  prod:string;
  done:boolean;
begin
  try
    try
      With Inv_DataSet do
      Begin
        // Get last processed
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'dbo.SELECT_ShipMax;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@AssyLine';
        Parameters.ParamValues['@AssyLine'] := fLineName;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to get production date data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to get production date data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to get production date data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;

        if not fieldbyname('prod').IsNull then
        begin
          prod:=fieldbyname('prod').AsString; //save
          NT.NummiTime:=prod+'00000000';  //get real
        end
        else
        begin
          NT.Time:=now;
        end;


        // Get future holiday and overtime
        with ALC_DAtaSet do
        begin
          Close;
          CommandType := CmdStoredProc;
          CommandText := 'dbo.AD_GetSpecialDate';
          Parameters.Clear;
          Parameters.AddParameter.Name := '@BeginDate';
          Parameters.ParamValues['@BeginDate'] := NT.Time;
          Parameters.AddParameter.Name := '@EndDate';
          Parameters.ParamValues['@EndDate'] := NT.Time+60;
          Parameters.AddParameter.Name := '@Line';
          Parameters.ParamValues['@Line'] := fLineName;
          Open;
          if Inv_Connection.Errors.Count > 0 then
          begin
            ShowMessage('Unable to get production date data, '+Inv_Connection.Errors.Item[0].Get_Description);
            LogActLog('ERROR','Unable to get production date data, '+Inv_Connection.Errors.Item[0].Get_Description);
            raise EDatabaseError.Create('Unable to get production date data');
          end
          else
          begin
            fAfterDateTime := Time;
            fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
          end;

          // Start with the next day
          done:=FALSE;
          repeat
            NT.Time:=NT.Time+1;
            if DayOfTheWeek(NT.Time) = 6 then
            begin
              // weekend check for overtime
              First;
              while not EOF do
              begin
                if NT.Time = fieldbyname('DATE').AsDateTime then
                begin
                  if fieldbyname('Date Status Abrv').AsString = 'O' then
                  begin
                    // work this day
                    done:=TRUE;
                  end;
                end;
                next;
              end;
            end
            else if DayOfTheWeek(NT.Time) < 6 then
            begin
              // weekday check for holiday
              First;
              while not EOF do
              begin
                if NT.Time = fieldbyname('DATE').AsDateTime then
                begin
                  if fieldbyname('Date Status Abrv').AsString = 'H' then
                  begin
                    // work this day
                    break;
                  end;
                end;
                next;
              end;
              if eof then
                done:=TRUE;
            end;
          until done = TRUE;

        end;
        fProductionDateDT:=NT.Time;
      End;    //With
      LogActLog('GET PROD', 'SELECTED GetNextProductionDate info', 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          GetRecConfStatInfo
        Else
          LogActLog('ERROR', 'FAILED SELECTED GetNextProductionDate info. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    fErrorCount := 0;
  end;
end;

procedure TData_Module.InsertShippingInfoDetail;
begin
  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.INSERT_ShippingDetail;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@PartShipID';
        Parameters.ParamValues['@PartShipID'] := fRecordID;
        Parameters.AddParameter.Name := '@PartNumber';
        Parameters.ParamValues['@PartNumber'] := fPartNum;
        Parameters.AddParameter.Name := '@ProductionDate';
        Parameters.ParamValues['@ProductionDate'] := fProductionDate;
        Parameters.AddParameter.Name := '@Qty';
        Parameters.ParamValues['@Qty'] := fQTY;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to insert ShippingDetail data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to insert ShippingDetail data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to insert ShippingDetail data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('UPDATE PO', 'INSERT ShippingDetail:' + fPartNum +'(' + IntToStr(fQTY) + ')', 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          INSERTShippingInfoDetail
        Else
          LogActLog('ERROR', 'FAILED INSERT ShippingDetail:' + fPartNum +'(' + IntToStr(fQTY) + ') . Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
end;      //InsertShippingInfoDetail

procedure TData_Module.UpdateShippingInfoDetail;
begin
  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.UPDATE_ShippingDetail;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@PartShipID';
        Parameters.ParamValues['@PartShipID'] := fRecordID;
        Parameters.AddParameter.Name := '@PartNumber';
        Parameters.ParamValues['@PartNumber'] := fPartNum;
        Parameters.AddParameter.Name := '@Qty';
        Parameters.ParamValues['@Qty'] := fQTY;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to update ShippingDetail data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to update ShippingDetail data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to update ShippingDetail data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('UPDATE PO', 'UPDATE ShippingDetail:' + fPartNum +'(' + IntToStr(fQTY) + ')', 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          UpdateShippingInfoDetail
        Else
          LogActLog('ERROR', 'FAILED UPDATE ShippingDetail:' + fPartNum +'(' + IntToStr(fQTY) + ') . Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
end;      //UpdateShippingInfoDetail

procedure TData_Module.GetShippingInfoDetail;
Begin
  try
    try
      With Inv_DataSet do
      Begin
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'dbo.SELECT_ShippingDetail;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@ShipID';
        Parameters.ParamValues['@ShipID'] := fRecordID;

        fBeforeDateTime := Time;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to get Shipping Detail data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to get Shipping Detail data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to get Shipping Detail data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('GET SHP', 'SELECTED GetShippingDetail info', 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          GetShippingInfoDetail
        Else
          LogActLog('ERROR', 'FAILED SELECTED GetShippingDetail info. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    fErrorCount := 0;
  end;
End;      //GetShippingInfo

procedure TData_Module.GetShippingInfo;
Begin
  try
    try
      With Inv_DataSet do
      Begin
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'dbo.SELECT_ShipLastSeq;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@LineName';
        Parameters.ParamValues['@LineName'] := fLineName;
        Parameters.AddParameter.Name := '@Date';
        Parameters.ParamValues['@Date'] := fProductionDate;

        fBeforeDateTime := Time;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to get Shipping data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to get Shipping data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to get Shipping data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('GET SHP', 'SELECTED GetShipping info', 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          GetShippingInfo
        Else
          LogActLog('ERROR', 'FAILED SELECTED GetShipping info. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    fErrorCount := 0;
  end;
End;      //GetShippingInfo

function TData_Module.GetLastSeqDate(offset:integer): Boolean;
begin
  result:=False;
  try
    With ALC_StoredProc do
    Begin
      Close;
      ProcedureName := 'dbo.AD_GetLastPrint;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@ASN';
      Parameters.ParamValues['@ASN'] := fStartSeq;
      Parameters.AddParameter.Name := '@LineName';
      Parameters.ParamValues['@LineName'] := fLineName;
      Parameters.AddParameter.Name := '@RevCount';
      Parameters.ParamValues['@RevCount'] := offset;
      Open;
      if Inv_Connection.Errors.Count > 0 then
      begin
        ShowMessage('Unable to check last seq date/time, '+Inv_Connection.Errors.Item[0].Get_Description);
        LogActLog('ERROR','Unable to check last seq date/time, '+Inv_Connection.Errors.Item[0].Get_Description);
        raise EDatabaseError.Create('Unable to check last seq date/time');
      end;
    end;

    result:=TRUE;
  except
    on e:exception do
    Begin
      LogActLog('ERROR', 'Failed to check last seq date/time. Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
      ALC_Connection.Connected:=FALSE;
    End;      //Except
  end;
end;

function TData_Module.CheckShippingInfo: Boolean;
begin
  result:=False;
  try
    With ALC_StoredProc do
    Begin
      Close;
      ProcedureName := 'dbo.AD_ProductionSeq;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@BeginDate';
      Parameters.ParamValues['@BeginDate'] := fBeginDateStr;
      Parameters.AddParameter.Name := '@EndDate';
      Parameters.ParamValues['@EndDate'] := fEndDateStr;
      Parameters.AddParameter.Name := '@Start';
      Parameters.ParamValues['@Start'] := fStartSeq;
      Parameters.AddParameter.Name := '@Last';
      Parameters.ParamValues['@Last'] := fLastSeq;
      Parameters.AddParameter.Name := '@LineName';
      Parameters.ParamValues['@LineName'] := fLineName;
      Open;
      if Inv_Connection.Errors.Count > 0 then
      begin
        ShowMessage('Unable to check '+fLineName+' Shipping data, '+Inv_Connection.Errors.Item[0].Get_Description);
        LogActLog('ERROR','Unable to check '+fLineName+' Shipping data, '+Inv_Connection.Errors.Item[0].Get_Description);
        raise EDatabaseError.Create('Unable to check '+fLineName+' Shipping data');
      end;
    end

  except
    on e:exception do
    Begin
      LogActLog('ERROR', 'Failed to check shipping information. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
    End;      //Except
  end;
end;      //CheckShippingInfo

procedure TData_Module.InsertExcelShippingEndInfo;
begin
  try
    With Inv_StoredProc do
    Begin
      Close;
      ProcedureName := 'dbo.INSERT_ShippingInfo;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@AssyLine';
      Parameters.ParamValues['@AssyLine'] := fCarTruck;
      Parameters.AddParameter.Name := '@StartSeq';
      Parameters.ParamValues['@StartSeq'] := fStartSeq;
      Parameters.AddParameter.Name := '@LastSeq';
      Parameters.ParamValues['@LastSeq'] := fLastSeq;
      Parameters.AddParameter.Name := '@QTY';
      Parameters.ParamValues['@QTY'] := fQTY;
      Parameters.AddParameter.Name := '@Continue';
      Parameters.ParamValues['@Continue'] := fContinuation;
      Parameters.AddParameter.Name := '@Date';
      Parameters.ParamValues['@Date'] := fProductionDate;

      fBeforeDateTime := Time;
      ExecProc;
      if Inv_Connection.Errors.Count > 0 then
      begin
        ShowMessage('Unable to insert shipping record, '+Inv_Connection.Errors.Item[0].Get_Description);
        LogActLog('ERROR','Unable to insert shipping record, '+Inv_Connection.Errors.Item[0].Get_Description);
        raise EDatabaseError.Create('Unable to insert shipping record');
      end
      else
      begin
        fAfterDateTime := Time;
        fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
      end;
    End;    //With
  except
    on e:exception do
    Begin
      LogActLog('ERROR', 'Failed insert shipping end. Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
      raise;
    End;      //Except
  end;
end;

procedure TData_Module.GetExcelPOInfo;
begin
  try
    try
      Inv_StoredProc.Close;
      Inv_StoredProc.ProcedureName := 'dbo.SELECT_AssyPOInfo;1';
      Inv_StoredProc.Parameters.Clear;
      Inv_StoredProc.Parameters.AddParameter.Name := '@AssyCode';
      Inv_StoredProc.Parameters.ParamValues['@assyCode'] := fAssyCode;
      Inv_StoredProc.Parameters.AddParameter.Name := '@ProdDate';
      Inv_StoredProc.Parameters.ParamValues['@ProdDate'] := fProductionDate;

      fBeforeDateTime := Time;
      Inv_StoredProc.Open;
      if Inv_Connection.Errors.Count > 0 then
      begin
        ShowMessage('Unable to get PO data, '+Inv_Connection.Errors.Item[0].Get_Description);
        LogActLog('ERROR','Unable to get PO data, '+Inv_Connection.Errors.Item[0].Get_Description);
        raise EDatabaseError.Create('Unable to get PO data');
      end
      else
      begin
        fAfterDateTime := Time;
        fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
      end;

      LogActLog('GET PO', 'SELECTED GetEXCELPO info', 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          GetRecConfStatInfo
        Else
          LogActLog('ERROR', 'FAILED SELECTED GetEXCELPO info. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    fErrorCount := 0;
  end;
end;

function TData_Module.UpdateINVDone:boolean;
begin
  result:=TRUE;
  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.UPDATE_AssyBuildHistINV;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@ChargedID';
        Parameters.ParamValues['@ChargedID'] := fInvoice;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to update Charge Hist INV data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to update Charge Hist INV data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to Charge Build Hist INV data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end
      End;    //With
      LogActLog('UPDATE INV', 'UPDATE Charge History ID=' + IntToStr(fInvoice), 1);
    except
      on e:exception do
      Begin
        result:=FALSE;
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          UpdateAssyRatioInfo
        Else
        begin
          LogActLog('ERROR', 'FAILED UPDATE Charge Hist ID=' + IntToStr(fInvoice) + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
          ShowMessage('FAILED UPDATE Charge Hist ID=' + IntToStr(fInvoice) + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
        end;
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
end;


function TData_Module.InsertPOCharged: Boolean;
begin
  result:=True;
  try
    With Inv_StoredProc do
    Begin
      Close;
      ProcedureName := 'dbo.INSERT_AssyPOCharged;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@AssyCode';
      Parameters.ParamValues['@AssyCode'] := fAssyCode;
      Parameters.AddParameter.Name := '@ProdDate';
      Parameters.ParamValues['@ProdDate'] := fProductionDate;
      Parameters.AddParameter.Name := '@PickUp';
      Parameters.ParamValues['@PickUp'] := fPickUp;
      Parameters.AddParameter.Name := '@QTY';
      Parameters.ParamValues['@QTY'] := fQTY;
      Parameters.AddParameter.Name := '@PONumber';
      Parameters.ParamValues['@PONumber'] := fPONumber;

      fBeforeDateTime := Time;
      ExecProc;
      if Inv_Connection.Errors.Count > 0 then
      begin
        ShowMessage('Unable to insert charge record, '+Inv_Connection.Errors.Item[0].Get_Description);
        LogActLog('ERROR','Unable to insert charge record, '+Inv_Connection.Errors.Item[0].Get_Description);
        raise EDatabaseError.Create('Unable to charge record');
      end
      else
      begin
        fAfterDateTime := Time;
        fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
      end;
      LogActLog('INS BuildHist', 'Insert charge info AssyCode='+fAssyCode+ 'PO='+ fPONumber +' Prod date='+fProductionDate+' Qty='+IntToStr(fQty), 1);
    End;    //With
  except
    on e:exception do
    begin
      LogActLog('ERROR', 'FAILED to INSERT Charge Production Date:'+ fProductionDate + ', AssyCode:'+ fAssyCode + 'PO='+ fPONumber + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
      result:=FALSE;
    end;
  end;
end;

procedure TData_Module.GetBuildHist;
begin
  try
    try
      With Inv_DataSet do
      Begin
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'dbo.SELECT_AssyBuildHist;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@BeginDate';
        Parameters.ParamValues['@BeginDate'] := FormatDateTime('yyyymmdd',fBeginDate);
        Parameters.AddParameter.Name := '@EndDate';
        Parameters.ParamValues['@EndDate'] := FormatDateTime('yyyymmdd',fEndDate);
        Parameters.AddParameter.Name := '@ASN';
        Parameters.ParamValues['@ASN'] := fASN;
        Parameters.AddParameter.Name := '@INVOICE';
        Parameters.ParamValues['@INVOICE'] := fINVOICE;


        fBeforeDateTime := Time;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to get build hist data, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to get build hist data, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to get build hist data');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('GET BuildHist', 'SELECTED build hist info', 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          GetStocktakingInfo
        Else
          LogActLog('ERROR', 'FAILED SELECT all build hist info. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    fErrorCount := 0;
  end;
end;

function TData_Module.InsertAutoScrap: integer;
begin
  result:=0;
  try
    With Inv_Dataset do
    begin
      Close;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.SELECT_PartsStockInfo;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@InvMgmtReport';
      Parameters.ParamValues['@InvMgmtReport'] := 'N';
      Parameters.AddParameter.Name := '@SupCode';
      Parameters.ParamValues['@SupCode'] := '';
      Parameters.AddParameter.Name := '@PartNum';
      Parameters.ParamValues['@PartNum'] := fPartNum;
      Open;

      if recordcount > 0 then
      begin
        if Inv_Dataset.FieldByName('Last Scrap Count').AsInteger <> fScrapCount then
        begin
          With Inv_StoredProc do
          Begin
            Close;
            ProcedureName := 'dbo.INSERT_StocktakingInfo;1';
            Parameters.Clear;
            Parameters.AddParameter.Name := '@SupCode';
            Parameters.ParamValues['@SupCode'] := Inv_Dataset.FieldByName('Supplier Code').AsString;
            Parameters.AddParameter.Name := '@PartCode';
            Parameters.ParamValues['@PartCode'] := fPartNum;
            Parameters.AddParameter.Name := '@QTY';
            Parameters.ParamValues['@QTY'] := 0-(fScrapCount-Inv_Dataset.FieldByName('Last Scrap Count').AsInteger);
            Parameters.AddParameter.Name := '@Reason';
            Parameters.ParamValues['@Reason'] := 'Auto Scrap Delete on '+DateTimeToStr(now);
            Parameters.AddParameter.Name := '@AutoScrap';
            Parameters.ParamValues['@AutoScrap'] := 1;

            fBeforeDateTime := Time;
            ExecProc;
            if Inv_Connection.Errors.Count > 0 then
            begin
              ShowMessage('Unable to insert auto scrap, '+Inv_Connection.Errors.Item[0].Get_Description);
              LogActLog('ERROR','Unable to insert auto scrap, '+Inv_Connection.Errors.Item[0].Get_Description);
              raise EDatabaseError.Create('Unable to insert auto scrap');
            end
            else
            begin
              fAfterDateTime := Time;
              fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
            end;
            LogActLog('INS BuildHist', 'Insert auto scrap PartNum='+PartNum+'Qty='+IntToStr(fScrapCount), 1);
            result:=fScrapCount-Inv_Dataset.FieldByName('Last Scrap Count').AsInteger;
          End;    //With
        end;
      end
      else
      begin
        LogActLog('ERROR', 'Part Number not found, Auto Scrap PartNum:'+ fPartNum );
        result:=-1;
      end;
    end;
  except
    on e:exception do
    begin
      LogActLog('ERROR', 'FAILED to INSERT Auto Scrap PartNum:'+ fPartNum + ', ' + E.Message + ' Err: ' + E.ClassName);
      result:=-1;
    end;
  end;
end;

function TData_Module.InsertBuildHist: Boolean;
begin
  result:=True;
  try
    With Inv_StoredProc do
    Begin
      Close;
      ProcedureName := 'dbo.INSERT_AssyBuildHist;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@AssyCode';
      Parameters.ParamValues['@AssyCode'] := fAssyCode;
      Parameters.AddParameter.Name := '@ProdDate';
      Parameters.ParamValues['@ProdDate'] := fProductionDate;
      Parameters.AddParameter.Name := '@QTY';
      Parameters.ParamValues['@QTY'] := fQTY;

      fBeforeDateTime := Time;
      ExecProc;
      if Inv_Connection.Errors.Count > 0 then
      begin
        ShowMessage('Unable to insert build record, '+Inv_Connection.Errors.Item[0].Get_Description);
        LogActLog('ERROR','Unable to insert build record, '+Inv_Connection.Errors.Item[0].Get_Description);
        raise EDatabaseError.Create('Unable to insert build record');
      end
      else
      begin
        fAfterDateTime := Time;
        fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
      end;
      LogActLog('INS BuildHist', 'Insert build hist info AssyCode='+fAssyCode+' Prod date='+fProductionDate+' Qty='+IntToStr(fQty), 1);
    End;    //With
  except
    on e:exception do
    begin
      LogActLog('ERROR', 'FAILED to INSERT Build Information Production Date:'+ fProductionDate + ', AssyCode:'+ fAssyCode + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
      result:=FALSE;
    end;
  end;
end;

function TData_Module.InsertExcelShippingInfo: Boolean;
begin
  result:=True;
  try
    Inv_StoredProc.Close;
    Inv_StoredProc.ProcedureName := 'dbo.SELECT_AssyRatioInfoAssy;1';
    Inv_StoredProc.Parameters.Clear;
    Inv_StoredProc.Parameters.AddParameter.Name := '@AssyCode';
    Inv_StoredProc.Parameters.ParamValues['@assyCode'] := fAssyCode;
    Inv_StoredProc.Open;
    if Inv_StoredProc.recordcount > 0 then
    begin
      // get all the part numbers
      // remove inventory
      if length(Inv_StoredProc.FieldByName('VC_TIRE_PART_NUMBER1_CODE').AsString) = 12 then
      begin
        DoPartNumberInventory(  Inv_StoredProc.FieldByName('VC_TIRE_PART_NUMBER1_CODE').AsString,
                                Inv_StoredProc.FieldByName('IN_TIRE_RATIO1').AsInteger,
                                fQty);
      end;

      if length(Inv_StoredProc.FieldByName('VC_TIRE_PART_NUMBER2_CODE').AsString) = 12 then
      begin
        DoPartNumberInventory( Inv_StoredProc.FieldByName('VC_TIRE_PART_NUMBER2_CODE').AsString,
                                Inv_StoredProc.FieldByName('IN_TIRE_RATIO2').AsInteger,
                                fQty);
      end;

      if length(Inv_StoredProc.FieldByName('VC_WHEEL_PART_NUMBER1_CODE').AsString) = 12 then
      begin
        DoPartNumberInventory( Inv_StoredProc.FieldByName('VC_WHEEL_PART_NUMBER1_CODE').AsString,
                                Inv_StoredProc.FieldByName('IN_WHEEL_RATIO1').AsInteger,
                                fQty);
      end;

      if length(Inv_StoredProc.FieldByName('VC_WHEEL_PART_NUMBER2_CODE').AsString) = 12 then
      begin
        DoPartNumberInventory( Inv_StoredProc.FieldByName('VC_WHEEL_PART_NUMBER2_CODE').AsString,
                                Inv_StoredProc.FieldByName('IN_WHEEL_RATIO2').AsInteger,
                                fQty);
      end;
    end;
  except
    on e:exception do
    begin
      LogActLog('ERROR', 'FAILED to INSERT Shipping information Production Date:'+ fProductionDate + ', Line:'+ fAssyCode + 'Start:'+fStartSeq+', Last:'+fLAstSeq + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
      result:=FALSE;
    end;
  end;
end;        //InsertExcelShippingInfo


function TData_Module.InsertShippingDetailManual: Boolean;
begin
  result:=FALSE;
  with INV_ShippingStoredProc do
  begin
    try
      Close;
      Parameters.Clear;
      ProcedureName:='INSERT_ShippingDetail;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@part';
      Parameters.ParamValues['@part'] := fPartNum;
      Parameters.AddParameter.Name := '@QTY';
      Parameters.ParamValues['@QTY'] := fQty;
      Parameters.AddParameter.Name := '@Date';
      Parameters.ParamValues['@Date'] := fProductionDate;
      Parameters.AddParameter.Name := '@AssyLine';
      Parameters.ParamValues['@AssyLine'] := fAssyCode;
      Parameters.AddParameter.Name := '@IrregularQty';
      Parameters.ParamValues['@IrregularQty'] := fIrregulareQty;

      ExecProc;
      if Inv_Connection.Errors.Count > 0 then
      begin
        ShowMessage('Unable to update part shipping part info, '+Inv_Connection.Errors.Item[0].Get_Description);
        LogActLog('ERROR','Unable to update part shipping part info, '+Inv_Connection.Errors.Item[0].Get_Description);
      end
      else
      begin
        result:=TRUE;
        LogActLog('INS Shp', 'INSERT Shipping detail information PartNumber: '+fPartNum+' Qty:'+ IntToStr(fQTY) + ', Irregular:'+ IntToStr(fIrregulareQty));
      end

    except
      on e:exception do
      begin
        Begin
          LogActLog('ERROR', 'FAILED to INSERT Shipping detail information PartNumber: '+fPartNum+' Production Date:'+ fProductionDate + ', Line:'+ fAssyCode +  '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
        End;      //Except
      end;
    end;
  end
end;


function TData_Module.InsertShippingInfoManual: Boolean;
var
  fDescription: String;
Begin
  Result := False;
  fDescription := 'INSERT Shipping information Production Date:'+ fProductionDate + ', Line:'+ fAssyCode + ', Start:'+fStartSeq+', Last:'+fLAstSeq;

  try
    try
        //
        //
        //  Insert status record
        //
        With Inv_StoredProc do
        Begin
          Close;
          ProcedureName := 'dbo.INSERT_ShippingInfo;1';
          Parameters.Clear;
          Parameters.AddParameter.Name := '@AssyLine';
          Parameters.ParamValues['@AssyLine'] := fAssyCode;
          Parameters.AddParameter.Name := '@StartSeq';
          Parameters.ParamValues['@StartSeq'] := fStartSeq;
          Parameters.AddParameter.Name := '@LastSeq';
          Parameters.ParamValues['@LastSeq'] := fLastSeq;
          Parameters.AddParameter.Name := '@QTY';
          Parameters.ParamValues['@QTY'] := fQTY;
          Parameters.AddParameter.Name := '@Continue';
          Parameters.ParamValues['@Continue'] := fContinuation;
          Parameters.AddParameter.Name := '@Date';
          Parameters.ParamValues['@Date'] := fProductionDate;

          fBeforeDateTime := Time;
          ExecProc;
          if Inv_Connection.Errors.Count > 0 then
          begin
            ShowMessage('Unable to insert shipping record, '+Inv_Connection.Errors.Item[0].Get_Description);
            LogActLog('ERROR','Unable to insert shipping record, '+Inv_Connection.Errors.Item[0].Get_Description);
            raise EDatabaseError.Create('Unable to insert shipping record');
          end
          else
          begin
            fAfterDateTime := Time;
            fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
            Result := True;
          end;
        End;    //With
        LogActLog('INS Shp', fDescription, 1);

    except
      on e:exception do
      Begin
        LogActLog('ERROR', 'FAILED to INSERT Shipping information Production Date:'+ fProductionDate + ', Line:'+ fAssyCode + 'Start:'+fStartSeq+', Last:'+fLAstSeq + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;        //InsertShippingInfo

function TData_Module.CalculateFRS:boolean;
var
  i:integer;
  TW : array[0..3] of string;
begin
  result:=true;
  TW[0]:='T';
  TW[1]:='W';
  TW[2]:='V';
  TW[3]:='F';
  //
  //  Pre process all select records into the FRS table
  //  then pass it into the report form
  //
  try
    //if fiGenerateEDI.AsBoolean then
    //begin
      //
      //  Pull data using the combo forecast table information
      //
      //
      with ALC_StoredProc do
      begin
        Close;
        ProcedureName:='AD_FRSPull;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@begindate';
        Parameters.ParamValues['@begindate'] := fbegindatestr;
        Parameters.AddParameter.Name := '@enddate';
        Parameters.ParamValues['@enddate'] := fenddatestr;
        Parameters.AddParameter.Name := '@Start';
        Parameters.ParamValues['@Start'] := fStartSeq;
        Parameters.AddParameter.Name := '@Last';
        Parameters.ParamValues['@Last'] := fLastSeq;
        Parameters.AddParameter.Name := '@LineName';
        Parameters.ParamValues['@LineName'] := fLineName;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to get '+fLineName+' FRS, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to get '+fLineName+' FRS , '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to get '+fLineName+' FRS');
        end
        else
        begin
          while not eof do
          begin
            for i:=low(TW) to high(TW) do
            begin
              Inv_StoredProc.Close;
              Inv_StoredProc.ProcedureName := 'dbo.SELECT_ForecastDetailBC;1';
              Inv_StoredProc.Parameters.Clear;
              Inv_StoredProc.Parameters.AddParameter.Name := '@BCode';
              Inv_StoredProc.Parameters.ParamValues['@BCode'] := FieldByName('BC').AsString;
              Inv_StoredProc.Parameters.AddParameter.Name := '@EffMonth';
              Inv_StoredProc.Parameters.ParamValues['@EffMonth'] := formatdatetime('yyyy/mm',fProductionDateDT);
              Inv_StoredProc.Parameters.AddParameter.Name := '@TireWheel';
              Inv_StoredProc.Parameters.ParamValues['@TireWheel'] := TW[i];
              Inv_StoredProc.Open;

              if Inv_StoredProc.recordcount > 0 then
              begin
                while not Inv_StoredProc.Eof do
                begin
                  if (FieldByName('Orders').AsInteger=4) or (FieldByName('Orders').AsInteger=5) then
                  begin
                    // if one wheel set as 100 percent for the first item
                    DoPartNumberInventory(  Inv_StoredProc.FieldByName('PartNo').AsString,
                                            100,
                                            FieldByName('Orders').AsInteger);
                    break;
                  end
                  else
                  begin
                    DoPartNumberInventory(  Inv_StoredProc.FieldByName('PartNo').AsString,
                                            Inv_StoredProc.FieldByName('ratio').AsInteger,
                                            FieldByName('Orders').AsInteger);
                    Inv_StoredProc.next;
                  end;
                end;
              end
              else
              begin
                // missing bc code fail on pull
                if i < 2 then
                begin
                  ShowMessage('Missing Broadcast Code Information('+FieldByName('BC').AsString+'), inventory pull failed');
                  LogActLog('ERROR','Missing Broadcast Code Information('+FieldByName('BC').AsString+'), inventory pull failed');
                  raise EDatabaseError.Create('Missing Broadcast Code Information('+FieldByName('BC').AsString+'), inventory pull failed');
                end;
              end;
            end;
            next; // get next BC
          end;
        end;
      end;
    {end
    else
    begin
      with ALC_StoredProc do
      begin
        Close;
        ProcedureName:='AD_WQSFRS;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@begindate';
        Parameters.ParamValues['@begindate'] := fbegindatestr;
        Parameters.AddParameter.Name := '@enddate';
        Parameters.ParamValues['@enddate'] := fenddatestr;
        Parameters.AddParameter.Name := '@Start';
        Parameters.ParamValues['@Start'] := fStartSeq;
        Parameters.AddParameter.Name := '@Last';
        Parameters.ParamValues['@Last'] := fLastSeq;
        Parameters.AddParameter.Name := '@LineName';
        Parameters.ParamValues['@LineName'] := fLineName;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to get '+fLineName+' FRS, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to get '+fLineName+' FRS , '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to get '+fLineName+' FRS');
        end
        else
        begin
          while not eof do
          begin
            Inv_StoredProc.Close;
            Inv_StoredProc.ProcedureName := 'dbo.SELECT_AssyRatioInfoRaw;1';
            Inv_StoredProc.Parameters.Clear;
            Inv_StoredProc.Parameters.AddParameter.Name := '@BroadcastCode';
            Inv_StoredProc.Parameters.ParamValues['@BroadcastCode'] := FieldByName('Tire').AsString;
            Inv_StoredProc.Parameters.AddParameter.Name := '@TirePartNum';
            Inv_StoredProc.Parameters.ParamValues['@TirePartNum'] := '';
            Inv_StoredProc.Parameters.AddParameter.Name := '@WheelPartNum';
            Inv_StoredProc.Parameters.ParamValues['@WheelPartNum'] := '';
            Inv_StoredProc.Open;
            if recordcount > 0 then
            begin
              // get all the part numbers
              // remove inventory
              if length(Inv_StoredProc.FieldByName('VC_TIRE_PART_NUMBER1_CODE').AsString) = 12 then
              begin
                DoPartNumberInventory(  Inv_StoredProc.FieldByName('VC_TIRE_PART_NUMBER1_CODE').AsString,
                                        Inv_StoredProc.FieldByName('IN_TIRE_RATIO1').AsInteger,
                                        FieldByName('Orders').AsInteger);
              end;

              if length(Inv_StoredProc.FieldByName('VC_TIRE_PART_NUMBER2_CODE').AsString) = 12 then
              begin
                DoPartNumberInventory( Inv_StoredProc.FieldByName('VC_TIRE_PART_NUMBER2_CODE').AsString,
                                        Inv_StoredProc.FieldByName('IN_TIRE_RATIO2').AsInteger,
                                        FieldByName('Orders').AsInteger);
              end;

              if length(Inv_StoredProc.FieldByName('VC_WHEEL_PART_NUMBER1_CODE').AsString) = 12 then
              begin
                DoPartNumberInventory( Inv_StoredProc.FieldByName('VC_WHEEL_PART_NUMBER1_CODE').AsString,
                                        Inv_StoredProc.FieldByName('IN_WHEEL_RATIO1').AsInteger,
                                        FieldByName('Orders').AsInteger);
              end;

              if length(Inv_StoredProc.FieldByName('VC_WHEEL_PART_NUMBER2_CODE').AsString) = 12 then
              begin
                DoPartNumberInventory( Inv_StoredProc.FieldByName('VC_WHEEL_PART_NUMBER2_CODE').AsString,
                                        Inv_StoredProc.FieldByName('IN_WHEEL_RATIO2').AsInteger,
                                        FieldByName('Orders').AsInteger);
              end;
            end;
            next;
          end;
        end;
      end;
    end;}
    //
    //  Compare Ship vs Forecast and send error is off by parameter percentage
    //


  except
    on e:exception do
    begin
      LogActLog('ERROR','Unable to do FRS data on Shipping, '+e.message);
      result:=False;
    end;
  end;
end;

function TData_Module.InsertShippingInfo: Boolean;
var
  fDescription: String;
Begin
  Result := False;
  fDescription := 'INSERT Shipping information Production Date:'+ fProductionDate + ', Line:'+ fAssyCode + ', Start:'+fStartSeq+', Last:'+fLAstSeq;

  try
    try
      Inv_Connection.BeginTrans;
      //
      //  Insert status record
      //
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.INSERT_ShippingInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name:='@ShippingID';
        Parameters.ParamByName('@ShippingID').Direction:=pdOutput;
        Parameters.ParamValues['@ShippingID']:=0;
        Parameters.AddParameter.Name := '@LineName';
        Parameters.ParamValues['@LineName'] := fLineName;
        Parameters.AddParameter.Name := '@StartSeq';
        Parameters.ParamValues['@StartSeq'] := fStartSeq;
        Parameters.AddParameter.Name := '@DTStartSeq';
        Parameters.ParamValues['@DTStartSeq'] := fBeginDateStr;
        Parameters.AddParameter.Name := '@EndSeq';
        Parameters.ParamValues['@EndSeq'] := fLastSeq;
        Parameters.AddParameter.Name := '@DTEndSeq';
        Parameters.ParamValues['@DTEndSeq'] := fEndDateStr;
        Parameters.AddParameter.Name := '@QTY';
        Parameters.ParamValues['@QTY'] := fQTY;
        Parameters.AddParameter.Name := '@Continue';
        Parameters.ParamValues['@Continue'] := fContinuation;
        Parameters.AddParameter.Name := '@Date';
        Parameters.ParamValues['@Date'] := fProductionDate;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to insert shipping record, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to insert shipping record, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to insert shipping record');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
          fRecordID:=Parameters.ParamByName('@ShippingID').Value;
        end;
      End;    //With
      LogActLog('INS Shp', fDescription, 1);


      //
      //  Calculate the number of assemblys built(USE FRS report info)
      //  and remove inventory
      //
      if CalculateFRS then
      begin
        if Inv_Connection.InTransaction then
          Inv_Connection.CommitTrans;
        Result := True;
      end
      else
      begin
        LogActLog('ERROR', 'Failed on update shipping info', 1);
        if Inv_Connection.InTransaction then
          Inv_Connection.RollbackTrans;
      end;
    except
      on e:exception do
      Begin
        LogActLog('ERROR', 'FAILED to INSERT Shipping information Production Date:'+ fProductionDate + ', Line:'+ fAssyCode + 'Start:'+fStartSeq+', Last:'+fLAstSeq + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
        if Inv_Connection.InTransaction then
          Inv_Connection.RollbackTrans;
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;        //InsertShippingInfo

function TData_Module.InsertINVInfo: Boolean;
var
  fDescription: String;
Begin
  Result := False;
  fDescription := 'INSERT INV ';

  try
    try
      Inv_Connection.BeginTrans;
      //
      //
      //  Insert status record
      //
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.INSERT_INVInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name:='@INVID';
        Parameters.ParamByName('@INVID').Direction:=pdOutput;
        Parameters.ParamValues['@INVID']:=0;
        Parameters.AddParameter.Name := '@EIN';
        Parameters.ParamValues['@EIN'] := fEIN+1;

        fBeforeDateTime := Time;
        ExecProc;

        fRecordID:=Parameters.ParamByName('@INVID').Value;

        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to insert ASN record, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to insert ASN record, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to insert ASN record');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
          LogActLog('INS ASN', 'Insert , 1');

          Close;
          ProcedureName := 'dbo.UPDATE_INVItems;1';
          Parameters.Clear;
          Parameters.AddParameter.Name := '@INVID';
          Parameters.ParamValues['@INVID'] := fRecordID;
          Parameters.AddParameter.Name := '@ASNID';
          Parameters.ParamValues['@ASNID'] := fASN;

          fBeforeDateTime := Time;
          ExecProc;

          Inv_Connection.CommitTrans;
          Result := True;
        end;
      End;    //With
    except
      on e:exception do
      Begin
        LogActLog('ERROR', 'FAILED to INSERT INV . Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
        Inv_Connection.RollbackTrans;
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;        //InsertINvInfo

function TData_Module.UpdateASNStatus: Boolean;
begin
  result:=True;
  try
    With Inv_StoredProc do
    Begin
      Close;
      ProcedureName := 'dbo.UPDATE_ASNStatus;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@ASNStatus';
      Parameters.ParamValues['@ASNStatus'] := fASNStatus;

      fBeforeDateTime := Time;
      ExecProc;
      if Inv_Connection.Errors.Count > 0 then
      begin
        ShowMessage('Unable to update ASN Status, '+Inv_Connection.Errors.Item[0].Get_Description);
        LogActLog('ERROR','Unable to update ASN Status, '+Inv_Connection.Errors.Item[0].Get_Description);
        raise EDatabaseError.Create('Unable to update ASN Status');
      end
      else
      begin
        fAfterDateTime := Time;
        fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
      end;
      LogActLog('UPD ASNStatus', 'Update ASN Status='+fASNStatus, 1);
    End;    //With
  except
    on e:exception do
    begin
      LogActLog('ERROR', 'FAILED to update ASN Status Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
      result:=FALSE;
    end;
  end;
end;

function TData_Module.CalculateASNFRS:boolean;
var
  count:integer;
  manifest:string;
begin
  result:=true;
  //
  //  Pre process all select records into the FRS table
  //  then pass it into the report form
  //
  try
    //
    //  Pull data using the combo forecast table information
    //
    //
    with ALC_StoredProc do
    begin
      Close;
      ProcedureName:='AD_FRSPull;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@begindate';
      Parameters.ParamValues['@begindate'] := fbegindatestr;
      Parameters.AddParameter.Name := '@enddate';
      Parameters.ParamValues['@enddate'] := fenddatestr;
      Parameters.AddParameter.Name := '@Start';
      Parameters.ParamValues['@Start'] := fStartSeq;
      Parameters.AddParameter.Name := '@Last';
      Parameters.ParamValues['@Last'] := fLastSeq;
      Parameters.AddParameter.Name := '@LineName';
      Parameters.ParamValues['@LineName'] := fLineName;
      Open;
      if Inv_Connection.Errors.Count > 0 then
      begin
        ShowMessage('Unable to get '+fLineName+' FRS, '+Inv_Connection.Errors.Item[0].Get_Description);
        LogActLog('ERROR','Unable to get '+fLineName+' FRS , '+Inv_Connection.Errors.Item[0].Get_Description);
        raise EDatabaseError.Create('Unable to get '+fLineName+' FRS');
      end
      else
      begin
        while not eof do
        begin
          Inv_StoredProc.Close;
          Inv_StoredProc.ProcedureName := 'dbo.SELECT_ForecastDetailBCASN;1';
          Inv_StoredProc.Parameters.Clear;
          Inv_StoredProc.Parameters.AddParameter.Name := '@BCode';
          Inv_StoredProc.Parameters.ParamValues['@BCode'] := ALC_StoredProc.FieldByName('BC').AsString;
          Inv_StoredProc.Parameters.AddParameter.Name := '@EffMonth';
          Inv_StoredProc.Parameters.ParamValues['@EffMonth'] := copy(fproductiondate,1,4)+'/'+copy(fproductiondate,5,2);
          Inv_StoredProc.Open;

          if Inv_StoredProc.recordcount > 0 then
          begin
            // Insert New ASN Record
            while not Inv_StoredProc.eof do
            begin
              count:=0;
              if ALC_StoredProc.FieldByName('Orders').AsInteger <= 5 then
              begin
                // no ratio on single vehicle
                manifest:='7'+copy(fProductionDate,4,5)+Inv_StoredProc.FieldByName('VC_ASSY_MANIFEST_NUMBER').AsString;

                INV_ShippingStoredProc.Close;
                INV_ShippingStoredProc.Parameters.Clear;
                INV_ShippingStoredProc.ProcedureName:='INSERT_ASNDetail;1';
                INV_ShippingStoredProc.Parameters.Clear;
                INV_ShippingStoredProc.Parameters.AddParameter.Name := '@ASNID';
                INV_ShippingStoredProc.Parameters.ParamValues['@ASNID'] := fRecordID;
                INV_ShippingStoredProc.Parameters.AddParameter.Name := '@EIN';
                INV_ShippingStoredProc.Parameters.ParamValues['@EIN'] := fEIN+1;
                INV_ShippingStoredProc.Parameters.AddParameter.Name := '@Manifest';
                INV_ShippingStoredProc.Parameters.ParamValues['@Manifest'] := manifest;
                INV_ShippingStoredProc.Parameters.AddParameter.Name := '@PartNumber';
                INV_ShippingStoredProc.Parameters.ParamValues['@PartNumber'] := Inv_StoredProc.FieldByName('VC_ASSY_PART_NUMBER_CODE').AsString;
                INV_ShippingStoredProc.Parameters.AddParameter.Name := '@Qty';
                INV_ShippingStoredProc.Parameters.ParamValues['@Qty'] := FieldByName('Orders').AsInteger;
                INV_ShippingStoredProc.ExecProc;
                if Inv_Connection.Errors.Count > 0 then
                begin
                  ShowMessage('Unable to update part shipping part info, '+Inv_Connection.Errors.Item[0].Get_Description);
                  LogActLog('ERROR','Unable to update part shipping part info, '+Inv_Connection.Errors.Item[0].Get_Description);
                  raise EDatabaseError.Create('Unable to update part shipping part info');
                end;

                break;
              end
              else
              begin
                if (Inv_StoredProc.FieldByName('IN_TIRE_RATIO').AsInteger = 100) and (Inv_StoredProc.FieldByName('IN_WHEEL_RATIO').AsInteger = 100) then
                begin
                  count:=FieldByName('Orders').AsInteger
                end
                else if (Inv_StoredProc.FieldByName('IN_TIRE_RATIO').AsInteger <> 100) and (Inv_StoredProc.FieldByName('IN_WHEEL_RATIO').AsInteger = 100) then
                begin
                  count:=round((FieldByName('Orders').AsInteger  * Inv_StoredProc.FieldByName('IN_TIRE_RATIO').AsInteger) / 100);

                  //count:=(FieldByName('Orders').AsInteger * Inv_StoredProc.FieldByName('IN_TIRE_RATIO').AsInteger) div 100;
                  //if (FieldByName('Orders').AsInteger * Inv_StoredProc.FieldByName('IN_TIRE_RATIO').AsInteger) mod 100 >= 50 then
                    //INC(count)
                end
                else if (Inv_StoredProc.FieldByName('IN_TIRE_RATIO').AsInteger = 100) and (Inv_StoredProc.FieldByName('IN_WHEEL_RATIO').AsInteger <> 100) then
                begin
                  count:=round((FieldByName('Orders').AsInteger  * Inv_StoredProc.FieldByName('IN_WHEEL_RATIO').AsInteger) / 100);

                  //count:=(FieldByName('Orders').AsInteger * Inv_StoredProc.FieldByName('IN_WHEEL_RATIO').AsInteger) div 100;
                  //if (FieldByName('Orders').AsInteger * Inv_StoredProc.FieldByName('IN_WHEEL_RATIO').AsInteger) mod 100 >= 50 then
                    //INC(count)
                end
                else if (Inv_StoredProc.FieldByName('IN_TIRE_RATIO').AsInteger <> 100) and (Inv_StoredProc.FieldByName('IN_WHEEL_RATIO').AsInteger <> 100) then
                begin

                  count:=round((round((FieldByName('Orders').AsInteger  * Inv_StoredProc.FieldByName('IN_TIRE_RATIO').AsInteger) / 100) * Inv_StoredProc.FieldByName('IN_WHEEL_RATIO').AsInteger) / 100);
                  //if (((FieldByName('Orders').AsInteger * Inv_StoredProc.FieldByName('IN_TIRE_RATIO').AsInteger) div 100) * Inv_StoredProc.FieldByName('IN_WHEEL_RATIO').AsInteger) mod 100 >= 50 then
                    //INC(count)
                end;

                // Create manifest number
                manifest:='7'+copy(fProductionDate,4,5)+Inv_StoredProc.FieldByName('VC_ASSY_MANIFEST_NUMBER').AsString;

                INV_ShippingStoredProc.Close;
                INV_ShippingStoredProc.Parameters.Clear;
                INV_ShippingStoredProc.ProcedureName:='INSERT_ASNDetail;1';
                INV_ShippingStoredProc.Parameters.Clear;
                INV_ShippingStoredProc.Parameters.AddParameter.Name := '@ASNID';
                INV_ShippingStoredProc.Parameters.ParamValues['@ASNID'] := fRecordID;
                INV_ShippingStoredProc.Parameters.AddParameter.Name := '@EIN';
                INV_ShippingStoredProc.Parameters.ParamValues['@EIN'] := fEIN+1;
                INV_ShippingStoredProc.Parameters.AddParameter.Name := '@Manifest';
                INV_ShippingStoredProc.Parameters.ParamValues['@Manifest'] := manifest;
                INV_ShippingStoredProc.Parameters.AddParameter.Name := '@PartNumber';
                INV_ShippingStoredProc.Parameters.ParamValues['@PartNumber'] := Inv_StoredProc.FieldByName('VC_ASSY_PART_NUMBER_CODE').AsString;
                INV_ShippingStoredProc.Parameters.AddParameter.Name := '@Qty';
                INV_ShippingStoredProc.Parameters.ParamValues['@Qty'] := count;
                INV_ShippingStoredProc.ExecProc;
                if Inv_Connection.Errors.Count > 0 then
                begin
                  ShowMessage('Unable to update part shipping part info, '+Inv_Connection.Errors.Item[0].Get_Description);
                  LogActLog('ERROR','Unable to update part shipping part info, '+Inv_Connection.Errors.Item[0].Get_Description);
                  raise EDatabaseError.Create('Unable to update part shipping part info');
                end;

                LogActLog('ASN', 'INSERT ASN entry ASNID('+IntToStr(fRecordID)+') Manifest('+manifest+') Qty('+IntToStr(count)+')');

              end;

              Inv_StoredProc.next;
            end;
          end
          else
          begin
            // missing bc code fail on pull
            ShowMessage('Missing Broadcast Code Information('+ALC_StoredProc.FieldByName('BC').AsString+'), ASN create failed');
            LogActLog('ERROR','Missing Broadcast Code Information('+ALC_StoredProc.FieldByName('BC').AsString+'), ASN create failed');
            raise EDatabaseError.Create('Missing Broadcast Code Information('+ALC_StoredProc.FieldByName('BC').AsString+'), ASN create failed');
          end;
          next; // get next BC
        end;
      end;
    end;

  except
    on e:exception do
    begin
      LogActLog('ERROR','Unable to do FRS data on Shipping, '+e.message);
      result:=False;
    end;
  end;
end;

function TData_Module.InsertASNInfo: Boolean;
var
  fDescription: String;
Begin
  Result := False;
  fDescription := 'INSERT ASN information Production Date:'+ fProductionDate + ', Line:'+ fLineName + ', Start:'+fStartSeq+', Last:'+fLAstSeq;

  try
    try
      //
      //
      //  Insert status record
      //
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.INSERT_ASNInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name:='@ASNID';
        Parameters.ParamByName('@ASNID').Direction:=pdOutput;
        Parameters.ParamValues['@ASNID']:=0;
        Parameters.AddParameter.Name := '@LineName';
        Parameters.ParamValues['@LineName'] := fLineName;
        Parameters.AddParameter.Name := '@AssyLine';
        Parameters.ParamValues['@AssyLine'] := fAssyName;
        Parameters.AddParameter.Name := '@StartSeq';
        Parameters.ParamValues['@StartSeq'] := fStartSeq;
        Parameters.AddParameter.Name := '@DTStartSeq';
        Parameters.ParamValues['@DTStartSeq'] := fBeginDateStr;
        Parameters.AddParameter.Name := '@EndSeq';
        Parameters.ParamValues['@EndSeq'] := fLastSeq;
        Parameters.AddParameter.Name := '@DTEndSeq';
        Parameters.ParamValues['@DTEndSeq'] := fEndDateStr;
        Parameters.AddParameter.Name := '@QTY';
        Parameters.ParamValues['@QTY'] := fQTY;
        Parameters.AddParameter.Name := '@PDate';
        Parameters.ParamValues['@PDate'] := fProductionDate;
        Parameters.AddParameter.Name := '@EIN';
        Parameters.ParamValues['@EIN'] := fEIN+1;

        fBeforeDateTime := Time;
        ExecProc;

        fRecordID:=Parameters.ParamByName('@ASNID').Value;

        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to insert ASN record, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to insert ASN record, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to insert ASN record');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.

          //
          //  Calculate the number of assemblys built(USE FRS report info)
          //  and remove inventory
          //
          if CalculateASNFRS then
          begin
            LogActLog('ASN', fDescription, 1);
            Result := True;
          end
          else
          begin
            LogActLog('ERROR', 'Failed on update ASN info', 1);
          end;
        end;
      End;    //With
    except
      on e:exception do
      Begin
        LogActLog('ERROR', 'FAILED to INSERT ASN information Production Date:'+ fProductionDate + ', Line:'+ fAssyCode + 'Start:'+fStartSeq+', Last:'+fLAstSeq + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName);
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;        //InsertASNInfo

procedure TData_Module.GetRecProdRejInfo;
Begin
  try
    try
      With Inv_DataSet do
      Begin
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'dbo.SELECT_RecProdRejInfo;1';

        Parameters.Clear;

        fBeforeDateTime := Time;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to get Prod Rej Info, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to get Prod Rej Info, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to get Prod Rej Info');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('GET RecRej', 'SELECTED all RecProdRej info', 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          GetRecProdRejInfo
        Else
          LogActLog('ERROR', 'FAILED SELECT all RecProdRej info. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    fErrorCount := 0;
  end;
End;        //GetRecProdRejInfo

function TData_Module.InsertRecProdRejInfo: Boolean;
var
  fDescription: String;
Begin
  Result := False;
  fDescription := 'INSERT RecProdRej S:' + fSupCode + ' P:' + fPartNum;

  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.INSERT_RecProdRejInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@Division';
        Parameters.ParamValues['@Division'] := fDivision;
        Parameters.AddParameter.Name := '@PartCode';
        Parameters.ParamValues['@PartCode'] := fPartNum;
        Parameters.AddParameter.Name := '@QTY';
        Parameters.ParamValues['@QTY'] := fQTY;
        Parameters.AddParameter.Name := '@Reason';
        Parameters.ParamValues['@Reason'] := fComments;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to insert Prod Rej Info, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to insert Prod Rej Info, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to insert Prod Rej Info');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
          Result := True;
        end;
      End;    //With
      LogActLog('INS RecRej', fDescription, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          InsertRecProdRejInfo
        Else
          LogActLog('ERROR', 'FAILED to INSERT RecProdRej S:' + fSupCode + ' P:' + fPartNum + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;        //InsertRecProdRejInfo

procedure TData_Module.UpdateRecProdRejInfo;
var
  fDescription: String;
Begin
  fDescription := 'UPDATE RecProdRej S:' + fSupCode + ' P:' + fPartNum;
  try
    try
      With Inv_StoredProc do
      Begin
        ProcedureName := 'dbo.UPDATE_RecProdRejInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@Division';
        Parameters.ParamValues['@Division'] := fDivision;
        Parameters.AddParameter.Name := '@SupCode';
        Parameters.ParamValues['@SupCode'] := fSupCode;
        Parameters.AddParameter.Name := '@PartCode';
        Parameters.ParamValues['@PartCode'] := fPartNum;
        Parameters.AddParameter.Name := '@QTY';
        Parameters.ParamValues['@QTY'] := fQTY;
        Parameters.AddParameter.Name := '@Reason';
        Parameters.ParamValues['@Reason'] := fComments;
        Parameters.AddParameter.Name := '@RejectID';
        Parameters.ParamValues['@RejectID'] := fRecordID;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to update Prod Rej Info, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to update Prod Rej Info, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to update Prod Rej Info');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('UPD RecRej', fDescription, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          UpdateRecProdRejInfo
        Else
          LogActLog('ERROR', 'FAILED ' + fDescription + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    fErrorCount := 0;
  end;
End;        //UpdateRecProdRejInfo

procedure TData_Module.DeleteRecProdRejInfo;
Begin
  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.DELETE_RecProdRejInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@RejectID';
        Parameters.ParamValues['@RejectID'] := fRecordID;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to delete Prod Rej Info, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to delete Prod Rej Info, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to delete Prod Rej Info');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('DEL RecRej', 'DELETED S: ' + fSupCode + ' P: ' + fPartNum, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          DeleteRecProdRejInfo
        Else
          LogActLog('ERROR', 'FAILED DELETE RecProdRej S: ' + fSupCode + ' P: ' + fPartNum + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;  //DeleteRecProdRejInfo

procedure TData_Module.SelectMultiField(fTableName: String; fFields: String; ComboBox: TNUMMIColumnComboBox);
var
  i:integer;
begin
  try
    try
      With Inv_Field_DataSet do
      Begin
        Close;
        CommandType := CmdText;
        CommandText := 'SELECT ' + fFields + ' FROM ' + fTableName;
        fBeforeDateTime := Time;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to select multi field, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to select multi field, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to select multi field');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.

          ComboBox.ComboItems.Clear;
          i:=0;
          ComboBox.ComboItems.Add;
          ComboBox.ComboItems.Items[i].Strings.Add(' ');
          ComboBox.ComboItems.Items[i].Strings.Add(' ');
          INC(i);
          While Not Eof do
          Begin
            ComboBox.ComboItems.Add;
            ComboBox.ComboItems.Items[i].Strings.Add(Fields[0].AsString);
            ComboBox.ComboItems.Items[i].Strings.Add(Fields[1].AsString);
            INC(i);
            Next;
          End;        //While
        end;
      End;    //With Inv_Query
      LogActLog('SELECT MULT', 'SELECTED ' + fFields + ' from ' + fTableName, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          SelectMultiField(fTableName, fFields, ComboBox)
        Else
          LogActLog('ERROR', 'FAILED SELECT MULTIFIELD:' + fFields + ' from ' + fTableName + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_Field_DataSet.Close;
    fErrorCount := 0;
  end;
end;  //SelectMultiField

procedure TData_Module.SelectMultiFieldALC(fTableName: String; fFields: String; ComboBox: TNUMMIColumnComboBox; fWhereAnd: string = '');
var
  i:integer;
begin
  try
    try
      With ALC_DataSet do
      Begin
        Close;
        CommandType := CmdText;
        if fWhereAnd = '' then
          CommandText := 'SELECT ' + fFields + ' FROM ' + fTableName
        else
          CommandText := 'SELECT ' + fFields + ' FROM ' + fTableName + ' WHERE '+fWhereAnd;

        fBeforeDateTime := Time;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to select multi field, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to select multi field, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to select multi field');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.

          ComboBox.ComboItems.Clear;
          i:=0;
          ComboBox.ComboItems.Add;
          ComboBox.ComboItems.Items[i].Strings.Add(' ');
          ComboBox.ComboItems.Items[i].Strings.Add(' ');
          INC(i);
          While Not Eof do
          Begin
            ComboBox.ComboItems.Add;
            ComboBox.ComboItems.Items[i].Strings.Add(Fields[0].AsString);
            ComboBox.ComboItems.Items[i].Strings.Add(Fields[1].AsString);
            INC(i);
            Next;
          End;        //While
        end;
      End;    //With Inv_Query
      LogActLog('SELECT MULT', 'SELECTED ' + fFields + ' from ' + fTableName, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          SelectMultiFieldALC(fTableName, fFields, ComboBox, fWhereAnd)
        Else
          LogActLog('ERROR', 'FAILED SELECT MULTIFIELD:' + fFields + ' from ' + fTableName + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_Field_DataSet.Close;
    fErrorCount := 0;
  end;
end;  //SelectMultiFieldALC

procedure TData_Module.SelectSingleFieldALC(fTableName: String; fFieldName: String; ComboBox: TComboBox; Distinct:boolean=TRUE);
Begin
  try
    try
      With ALC_DataSet do
      Begin
        Close;
        CommandType := CmdText;
        if Distinct then
          CommandText := 'SELECT DISTINCT(' + fFieldName + ') FROM ' + fTableName + ' ORDER BY ' + fFieldName
        else
          CommandText := 'SELECT ' + fFieldName + ' FROM ' + fTableName + ' ORDER BY ' + fFieldName;
        fBeforeDateTime := Time;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to select single field, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to select single field, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to select single field');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.

          TComboBox(ComboBox).Clear;
          TComboBox(ComboBox).Items.Add(' ');
          While Not Eof do
          Begin
            TComboBox(ComboBox).Items.Add(FieldByName(fFieldName).AsString);
            Next;
          End;        //While
        end;
      End;    //With Inv_Query
      LogActLog('SELECT SFT', 'SELECTED ' + fFieldName + ' from ' + fTableName, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          SelectSingleField(fTableName, fFieldName, ComboBox)
        Else
          LogActLog('ERROR', 'FAILED SELECT FIELD:' + fFieldName + ' from ' + fTableName + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_Field_DataSet.Close;
    fErrorCount := 0;
  end;
End;  //SelectSingleFieldALC

procedure TData_Module.SelectSingleField(fTableName: String; fFieldName: String; ComboBox: TComboBox; Distinct:boolean=TRUE);
Begin
  try
    try
      With Inv_Field_DataSet do
      Begin
        CommandType := CmdText;
        if Distinct then
          CommandText := 'SELECT DISTINCT(' + fFieldName + ') FROM ' + fTableName + ' ORDER BY ' + fFieldName
        else
          CommandText := 'SELECT ' + fFieldName + ' FROM ' + fTableName + ' ORDER BY ' + fFieldName;
        fBeforeDateTime := Time;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to select single field, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to select single field, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to select single field');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.

          TComboBox(ComboBox).Clear;
          TComboBox(ComboBox).Items.Add(' ');
          While Not Eof do
          Begin
            TComboBox(ComboBox).Items.Add(FieldByName(fFieldName).AsString);
            Next;
          End;        //While
        end;
      End;    //With Inv_Query
      LogActLog('SELECT SFT', 'SELECTED ' + fFieldName + ' from ' + fTableName, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          SelectSingleField(fTableName, fFieldName, ComboBox)
        Else
          LogActLog('ERROR', 'FAILED SELECT FIELD:' + fFieldName + ' from ' + fTableName + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_Field_DataSet.Close;
    fErrorCount := 0;
  end;
End;        //SelectSingleField

procedure TData_Module.SelectDependantSingleField(StoredProc: String; ParameterName: String; Field:string; Value: Variant; ComboBox: TComboBox);
Begin
  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.'+StoredProc;
        Parameters.Clear;
        Parameters.AddParameter.Name := ParameterName;
        Parameters.ParamValues[ParameterName] := Value;

        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to select dependant single field, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to select dependant single field, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to select dependant single field');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.

          TComboBox(ComboBox).Clear;
          TComboBox(ComboBox).Items.Add(' ');
          While Not Eof do
          Begin
            TComboBox(ComboBox).Items.Add(FieldByName(Field).AsString);
            Next;
          End;        //While
        end;
      End;    //With Inv_Query
      LogActLog('SELECT DSFT', 'SELECTED ' + Field + ' from ' + StoredProc, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          SelectDependantSingleField(StoredProc, ParameterName, Field, Value, ComboBox)
        Else
          LogActLog('ERROR', 'FAILED SELECT DEPENDANT SINGLE FIELD:' + Field + ' from ' + StoredProc + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;        //SelectDependantSingleField

procedure TData_Module.SetComboBoxesWithObj(ComboBox: TComboBox; fQryString: String);
var
  TwoFields: TTwoFieldObj;
Begin
  try
    try
      With Inv_Field_DataSet do
      Begin
        CommandType := cmdText;
        CommandText := fQryString;
        fBeforeDateTime := Time;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to set combobox, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to set combobox, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to set combobox');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
          TComboBox(ComboBox).Clear;
          While Not EOF do
          Begin
            TwoFields := TTwoFieldObj.Create (Fields[0].AsString, Fields[1].AsString);
            TComboBox(ComboBox).Items.AddObject (TwoFields.Code, TwoFields);
            Next;
          End;      //While
          Close;
        end;
      End;      //With
      LogActLog('SET CBwOBJ', 'SetCBw/OBJ:' + ComboBox.Name, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          SetComboBoxesWithObj(ComboBox, fQryString)
        Else
          LogActLog('ERROR', 'FAILED SetCBw/OBJ:' + ComboBox.Name + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    End;
  finally
    fErrorCount := 0;
    Inv_Field_DataSet.Close;
  End;
End;          //SetComboBoxesWithObj

procedure TData_Module.SearchMultiCombo(ComboBox: TNUMMIColumnComboBox; value: String; Column:integer);
var fCount: Integer;
Begin
  if value = '' then
    value := ' ';
  ComboBox.ItemIndex:=0;
  For fCount := 0 to ComboBox.ComboItems.Count - 1 do
  Begin
    If ComboBox.ComboItems.Items[fCount].Strings[0] = value Then
    Begin
      ComboBox.ItemIndex := fCount;
      Break;
    End;
  End;    //For
End;      //SearchMultiCombo

          //  ComboBox.ComboItems.Items[i].Strings.Add(Fields[0].AsString);
          //  ComboBox.ComboItems.Items[i].Strings.Add(Fields[1].AsString);

procedure TData_Module.SearchCombo(ComboBox: TComboBox; value: String);
var fCount: Integer;
Begin
  if value = '' then
    value := ' ';
  TComboBox(ComboBox).ItemIndex:=0;
  For fCount := 0 to TComboBox(ComboBox).Items.Count - 1 do
  Begin
    If TComboBox(ComboBox).Items[fCount] = value Then
    Begin
      TComboBox(ComboBox).ItemIndex := fCount;
      Break;
    End;
  End;    //For
End;      //SearchCombo

procedure TData_Module.JustifyColumns(DBGrid: TDBGrid);
var fCount: Integer;
Begin
  For fCount := 0 to DBGrid.Columns.Count - 1 do
  Begin
    If (DBGrid.Columns[fCount].FieldName = 'Price') Or (DBGrid.Columns[fCount].FieldName = 'Total Cost') Then
      DBGrid.Columns[fCount].Alignment := taRightJustify;
    If (Copy(DBGrid.Columns[fCount].FieldName, 1, 6) = 'Report') Then
      DBGrid.Columns[fCount].Visible := False;
  End;      //For
End;        //JustifyColumns


procedure TData_Module.SetTodaysDate(Panel: TPanel);
var fCount: Integer;
Begin
  For fCount := 0 to Panel.ControlCount - 1 do
  Begin
    If (Panel.Controls[fCount] is TDateTimePicker) Then
    Begin
      TDateTimePicker(Panel.Controls[fCount]).Date := Date;
    End;    //IF ...
  End;      //For
End;        //SetTodaysDate

procedure TData_Module.ClearControls(Control: TControl);
var
  fContCount: Integer;
Begin
  If Control is TPanel Then
    For fContCount := 0 to TPanel(Control).ControlCount - 1 do
    Begin
      If (TPanel(Control).Controls[fContCount] is TEdit) and Not(TPanel(Control).Controls[fContCount].Tag = 1) Then
        TEdit(TPanel(Control).controls[fContCount]).Text := '';
      If (TPanel(Control).Controls[fContCount] is TMaskEdit) and Not(TPanel(Control).Controls[fContCount].Tag = 1) Then
        TMaskEdit(TPanel(Control).controls[fContCount]).Text := '';
      If (TPanel(Control).Controls[fContCount] is TCheckBox) and Not(TPanel(Control).Controls[fContCount].Tag = 1) Then
        TCheckBox(TPanel(Control).controls[fContCount]).Checked := False;
      If (TPanel(Control).Controls[fContCount] is TMemo) and Not(TPanel(Control).Controls[fContCount].Tag = 1) Then
        TMemo(TPanel(Control).controls[fContCount]).Text := '';
      If (TPanel(Control).Controls[fContCount] is TComboBox) and Not(TPanel(Control).Controls[fContCount].Tag = 1) Then
      begin
        TComboBox(TPanel(Control).controls[fContCount]).ItemIndex:=0;
        //TComboBox(TPanel(Control).controls[fContCount]).Text := '';
      end;
      If (TPanel(Control).Controls[fContCount] is TNUMMIBmDateEdit) and Not(TPanel(Control).Controls[fContCount].Tag = 1) Then
        TNUMMIBmDateEdit(TPanel(Control).controls[fContCount]).Text := '';
    End
  Else If Control is TGroupBox Then
    For fContCount := 0 to TPanel(Control).ControlCount - 1 do
    Begin
      If (TGroupBox(Control).Controls[fContCount] is TEdit) and Not(TGroupBox(Control).Controls[fContCount].Tag = 1) Then
        TEdit(TGroupBox(Control).controls[fContCount]).Text := '';
      If (TGroupBox(Control).Controls[fContCount] is TMaskEdit) and Not(TGroupBox(Control).Controls[fContCount].Tag = 1) Then
        TMaskEdit(TGroupBox(Control).controls[fContCount]).Text := '';
      If (TGroupBox(Control).Controls[fContCount] is TCheckBox) and Not(TGroupBox(Control).Controls[fContCount].Tag = 1) Then
        TCheckBox(TGroupBox(Control).controls[fContCount]).Checked := False;
      If (TGroupBox(Control).Controls[fContCount] is TMemo) and Not(TGroupBox(Control).Controls[fContCount].Tag = 1) Then
        TMemo(TGroupBox(Control).controls[fContCount]).Text := '';
      If (TGroupBox(Control).Controls[fContCount] is TNUMMIBmDateEdit) and Not(TGroupBox(Control).Controls[fContCount].Tag = 1) Then
        TNUMMIBmDateEdit(TGroupBox(Control).controls[fContCount]).Text := '';
      If (TPanel(Control).Controls[fContCount] is TComboBox) and Not(TPanel(Control).Controls[fContCount].Tag = 1) Then
      begin
        TComboBox(TGroupBox(Control).controls[fContCount]).ItemIndex:=0;
        //TComboBox(TGroupBox(Control).controls[fContCount]).Text := '';
      end;
    End
End;      //ClearControls

procedure TData_Module.LogActLog(fTrans: String; fDescription: String; fWithDB_Time: Integer=0);
Begin
  With Act_StoredProc do
  Begin
    Parameters.ParamValues['@App_From'] := fAppFrom;
    Parameters.ParamValues['@IP_Address'] := gobjUser.IPAddress;      //IPAddress
    Parameters.ParamValues['@Trans'] := fTrans;                  //Trans
    Parameters.ParamValues['@DT_Sender'] := Copy(FormatDateTime('yyyymmddhhnnsszzz', Now), 1, 16);
    Parameters.ParamValues['@ComputerName'] := gobjUser.LocalName;      //ComputerName
    Parameters.ParamValues['@Description'] := fDescription;            //Description
    Parameters.ParamValues['@NTUserName'] := gobjUser.UserName;       //Win_UserID
    Case fWithDB_Time of
      0:
      Begin
        Parameters.ParamValues['@DB_Time'] := NULL;
      End;
      1:
      Begin
        Parameters.ParamValues['@DB_Time'] := fDiffDateTime;                //DB_Time
      End;
    End;     //Case fWithDB_Time
    Parameters.ParamValues['@VIN'] := '';       //Not used for this app
    Parameters.ParamValues['@Sequence_Number'] := 0;       //Not used for this app
    Parameters.ParamValues['@Last_Modified'] := gobjUser.AppUserID;       //App User ID from login
    Try
      ExecProc;
    Except
      On E:Exception Do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          LogActLog(fTrans, fDescription, fWithDB_Time);
      End;      //Except
    End;        //Try
  End;  //With sprocDBHistory

  fBeforeDateTime := 0;
  fAfterDateTime := 0;
  fDiffDateTime := 0;
  fErrorCount := 0;
End;      //LogActLog

function TData_Module.ValidateUser(fUserID: String; fPassword: String): Boolean;
var
  fTrans: String;
  fDescr: String;
Begin
  Result := False;
  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.SELECT_UserInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@UserID';
        Parameters.ParamValues['@UserID'] := fUserID;
        Parameters.AddParameter.Name := '@Pass';
        Parameters.ParamValues['@Pass'] := fPassword;
        //Parameters.AddParameter.Name := '@Admin';
        //Parameters.ParamValues['@Admin'] := fTempAdmin;
        Open;

        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to validate user, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to validate user, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to validate user');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
          If RecordCount > 0 Then
          Begin
            Result := True;
            fTrans := 'LOGIN';
            fDescr := fUserID + ' logged in.';
            gobjUser.AppUserID := FieldByName('VC_USER_ID').AsString;
            gobjUser.AppUserPass := FieldByName('VC_PASSWORD').AsString;
            gobjUser.AppUserAdmin := FieldByName('BIT_ADMIN').AsBoolean;
          End
          Else
          Begin
            Result := False;
            fTrans := 'LOGIN ERR';
            fDescr := fUserID + ' failed to log in.';
          End;
        end;
      End;
      LogActLog(fTrans, fDescr, 1);
    except
      On e:exception do
      Begin
        If fErrorCount < 3 Then
        Begin
          LogActLog('ERROR', 'Unable to ValidateUser: ' + E.Message);
          Result := ValidateUser(fUserID, fPassword);
          Inc(fErrorCount);
        End
        Else
          MessageDlg('Unable to validate your' + #13 + 'information at this time.', mtError, [mbOK], 0);
      End;
    end;
  finally
    fErrorCount := 0;
    Inv_DataSet.Close;
  end;
End;

procedure TData_Module.DataModuleDestroy(Sender: TObject);
begin
  FreeADOs;
end;

procedure TData_Module.SetComboBoxesWithUserObj(ComboBox: TComboBox; fQryString: String);
var
  UserAdminDetail: TUserAdminDetail;
Begin
  try
    try
      With Inv_Field_DataSet do
      Begin
        CommandType := cmdText;
        CommandText := fQryString;
        fBeforeDateTime := Time;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to set combobox, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to set combobox, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to set combobox');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
          TComboBox(ComboBox).Clear;
          TComboBox(ComboBox).Items.AddObject ('', nil);
          While Not EOF do
          Begin
            UserAdminDetail := TUserAdminDetail.Create (Fields[0].AsString, Fields[1].AsString, Fields[2].AsBoolean);
            TComboBox(ComboBox).Items.AddObject (UserAdminDetail.ID, UserAdminDetail);
            Next;
          End;      //While
          Close;
        end;
      End;      //With
      LogActLog('SET CBwOBJ', 'SetCBw/OBJ:' + ComboBox.Name, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          SetComboBoxesWithUserObj(ComboBox, fQryString)
        Else
          LogActLog('ERROR', 'FAILED SetCBw/OBJ:' + ComboBox.Name + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    End;
  finally
    fErrorCount := 0;
    Inv_Field_DataSet.Close;
  End;
End;          //SetComboBoxesWithUserObj

function TData_Module.InsertUser(fTempID: String; fTempPass: String; fTempAdmin: Boolean): Boolean;
var
  fDescription: String;
Begin
  Result := False;
  If fTempAdmin Then
    fDescription := 'INSERT ' + fTempID + ' as ADMIN'
  Else
    fDescription := 'INSERT ' + fTempID + '(NO ADMIN)';

  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.SELECT_UserInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@UserID';
        Parameters.ParamValues['@UserID'] := fTempID;
        Parameters.AddParameter.Name := '@Pass';
        Parameters.ParamValues['@Pass'] := fTempPass;
        Parameters.AddParameter.Name := '@Admin';
        Parameters.ParamValues['@Admin'] := fTempAdmin;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to get user dup, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to get user dup, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to get user dup');
        end
        else
        begin
          If RecordCount = 0 Then
          Begin
            Close;
            ProcedureName := 'dbo.INSERT_UserInfo;1';
            Parameters.Clear;
            Parameters.AddParameter.Name := '@UserID';
            Parameters.ParamValues['@UserID'] := fTempID;
            Parameters.AddParameter.Name := '@Pass';
            Parameters.ParamValues['@Pass'] := fTempPass;
            Parameters.AddParameter.Name := '@Admin';
            Parameters.ParamValues['@Admin'] := fTempAdmin;

            fBeforeDateTime := Time;
            ExecProc;
            if Inv_Connection.Errors.Count > 0 then
            begin
              ShowMessage('Unable to insert User, '+Inv_Connection.Errors.Item[0].Get_Description);
              LogActLog('ERROR','Unable to insert User, '+Inv_Connection.Errors.Item[0].Get_Description);
              raise EDatabaseError.Create('Unable to insert User');
            end
            else
            begin
              fAfterDateTime := Time;
              fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
              Result := True;
            end;
          End
          Else
            fDescription := 'FAILED to ' + fDescription + ' (DUPLICATE)';
        end;
      End;    //With
      LogActLog('INSERT USR', fDescription, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          Result := InsertUser(fTempID, fTempPass, fTempAdmin)
        Else
          LogActLog('ERROR', 'FAILED to ' + fDescription + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;

procedure TData_Module.UpdateUserInfo(fTempOldID: String; fTempOldPass: String; fTempNewID: String; fTempNewPass: String; fTempAdmin: Boolean);
var
  fDescription: String;
Begin
  fDescription := 'UPDATE User ID: ' + fTempOldID + ' New ID: ' + fTempNewID;

  try
    try
      With Inv_StoredProc do
      Begin
        ProcedureName := 'dbo.UPDATE_UserInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@OldUserID';
        Parameters.ParamValues['@OldUserID'] := fTempOldID;
        Parameters.AddParameter.Name := '@OldPass';
        Parameters.ParamValues['@OldPass'] := fTempOldPass;
        Parameters.AddParameter.Name := '@NewUserID';
        Parameters.ParamValues['@NewUserID'] := fTempNewID;
        Parameters.AddParameter.Name := '@NewPass';
        Parameters.ParamValues['@NewPass'] := fTempNewPass;
        Parameters.AddParameter.Name := '@Admin';
        Parameters.ParamValues['@Admin'] := ABS(StrToInt(BoolToStr(fTempAdmin, False)));

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to update User, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to update User, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to update User');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('UPD USER', fDescription, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          UpdateUserInfo(fTempOldID, fTempOldPass, fTempNewID, fTempNewPass, fTempAdmin)
        Else
          LogActLog('ERROR', 'FAILED ' + fDescription + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    fErrorCount := 0;
  end;
End;        //UpdateUserInfo

procedure TData_Module.DeleteUserInfo(fTempID: String; fTempPass: String);
Begin
  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.DELETE_UserInfo;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@UserID';
        Parameters.ParamValues['@UserID'] := fTempID;
        Parameters.AddParameter.Name := '@Pass';
        Parameters.ParamValues['@Pass'] := fTempPass;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to delete User, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to delete User, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to delete User');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('DEL USER', 'DELETED ID: ' + fTempID, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          DeleteUserInfo(fTempID, fTempPass)
        Else
          LogActLog('ERROR', 'FAILED DELETE User ID: ' + fTempID + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;        //DeleteUserInfo

//
//  DMV 07-06-2004
//
procedure TData_Module.GetFirstProductionDayInfo;
Begin
  try
    try
      With Inv_DataSet do
      Begin
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'dbo.SELECT_FirstProductionDay;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@ProdYear';
        Parameters.ParamValues['@ProdYear'] := fProdYear;

        fBeforeDateTime := Time;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to get First Production Day, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to get First Production Day, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to get First Production Day');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('GET OT/HOL', 'SELECTED all First Production Day', 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          GetOvertimeHolidayInfo
        Else
          LogActLog('ERROR', 'FAILED SELECT all First Production Day. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    fErrorCount := 0;
  end;
End;        //GetFirstProductionDayInfo

procedure TData_Module.DeleteFirstProductiondayInfo;
Begin
  try
    try
      With Inv_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.DELETE_FirstProductionDay;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@ProdYear';
        Parameters.ParamValues['@ProdYear'] := fProdYear;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to delete First Production Day, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to delete First Production Day, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to delete First Production Day');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('DEL PD', 'DELETED FirstProductionDay: ' + fProdYear, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          DeleteFirstProductionDayInfo
        Else
          LogActLog('ERROR', 'FAILED DELETE First Production Day: ' + fProdyear + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;        //DeleteFirstProductionDayInfo


function TData_Module.InsertFirstProductionDayInfo: Boolean;
Begin
  Result := False;

  try
    With Inv_StoredProc do
    Begin
      Close;
      ProcedureName := 'dbo.INSERT_FirstProductionDay;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@ProdYear';
      Parameters.ParamValues['@ProdYear'] := fProdYear;
      Parameters.AddParameter.Name := '@ProdDate';
      Parameters.ParamValues['@ProdDate'] := fEventDate;
      Parameters.AddParameter.Name := '@WeekNumber';
      Parameters.ParamValues['@WeekNumber'] := fWeekNumber;

      fBeforeDateTime := Time;
      ExecProc;
      if Inv_Connection.Errors.Count > 0 then
      begin
        ShowMessage('Unable to insert First Production Date, '+Inv_Connection.Errors.Item[0].Get_Description);
        LogActLog('ERROR','Unable to insert First Production Date, '+Inv_Connection.Errors.Item[0].Get_Description);
        raise EDatabaseError.Create('Unable to insert First Production Date');
      end
      else
      begin
        fAfterDateTime := Time;
        fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        Result := True;
      end;
    End;    //With
    LogActLog('INSERT PD', 'Insert first production date, '+fProdyear, 1);
  except
    on e:exception do
    Begin
      fErrorCount := fErrorCount + 1;
      If fErrorCount < 3 Then
        InsertSizeInfo
      Else
        LogActLog('ERROR', 'FAILED to INSERT First Production Date ' + fProdyear + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
    End;      //Except
  end;
End;        //InsertFirstProductiondayInfo




//
//  DMV 12-18-2002
//
procedure TData_Module.GetOvertimeHolidayInfo;
Begin
  try
    try
      With ALC_DataSet do
      Begin
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'dbo.AD_GetSpecialDates';
        Parameters.Clear;

        fBeforeDateTime := Time;
        Open;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to get Overtime/Holiday, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to get Overtime/Holiday, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to get Overtime/Holiday');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('GET OT/HOL', 'SELECTED all Overtime/Holiday', 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          GetOvertimeHolidayInfo
        Else
          LogActLog('ERROR', 'FAILED SELECT all Overtime/Holiday. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_DataSource.DataSet := Inv_DataSet;
    fErrorCount := 0;
  end;
End;        //GetOvertimeHolidayInfo

procedure TData_Module.DeleteOvertimeHolidayInfo;
Begin
  try
    try
      With ALC_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.AD_DeleteSpecialDate';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@SpecialDateID';
        Parameters.ParamValues['@SpecialDateID'] := fRecordID;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to delete Overtime/Holiday, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to delete Overtime/Holiday, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to delete Overtime/Holiday');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
        end;
      End;    //With
      LogActLog('DELETE OH', 'DELETED D:' + formatdatetime('mm/dd/yyyy',fEventDate) +  'L:' + fEventLine, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          DeleteSupplierInfo
        Else
          LogActLog('ERROR', 'FAILED DELETE PartsStockInfo D:' + formatdatetime('mm/dd/yyyy',fEventDate) +  'L:' + fEventLine + '. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;      //DeleteOvertimeHolidayInfo

function TData_Module.InsertOvertimeHolidayInfo: Boolean;
var
  fDescription: String;
Begin
  Result := False;
  fDescription := 'INSERT Overtime/Holiday Sup: ' + formatdatetime('mm/dd/yyyy',fEventDate) + ' and Line: ' + fLineName +' and Type:' +fEventType;

  try
    try
      With ALC_StoredProc do
      Begin
        Close;
        ProcedureName := 'dbo.AD_InsertSpecialDate';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@LineName';
        Parameters.ParamValues['@LineName'] := fLineName;
        Parameters.AddParameter.Name := '@ProductionStatus';
        Parameters.ParamValues['@ProductionStatus'] := fEventType;
        Parameters.AddParameter.Name := '@SpecialDateName';
        Parameters.ParamValues['@SpecialDateName'] := fEventDescription;
        Parameters.AddParameter.Name := '@SpecialDate';
        Parameters.ParamValues['@SpecialDate'] := fEventdate;

        fBeforeDateTime := Time;
        ExecProc;
        if Inv_Connection.Errors.Count > 0 then
        begin
          ShowMessage('Unable to insert Overtime/Holiday, '+Inv_Connection.Errors.Item[0].Get_Description);
          LogActLog('ERROR','Unable to insert Overtime/Holiday, '+Inv_Connection.Errors.Item[0].Get_Description);
          raise EDatabaseError.Create('Unable to insert Overtime/Holiday');
        end
        else
        begin
          fAfterDateTime := Time;
          fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
          Result := True;
        end;
      End;    //With
      LogActLog('INS Overtime/Holiday', fDescription, 1);
    except
      on e:exception do
      Begin
        fErrorCount := fErrorCount + 1;
        If fErrorCount < 3 Then
          InsertSupplierInfo
        Else
          LogActLog('ERROR', 'FAILED: ' + fDescription + ' Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
      End;      //Except
    end;
  finally
    Inv_StoredProc.Close;
    fErrorCount := 0;
  end;
End;      //InsertOvertimeHolidayInfo

procedure TData_Module.DoPartNumberInventory(partnumber:string;ratio,qty:integer);
var
  update:integer;
begin

  update:=round((qty*ratio) / 100);

  LogActLog('UPDATE', 'Update inventory with ALC partnumber('+partnumber+') count('+inttostr(update)+')');

  with INV_ShippingStoredProc do
  begin
    //
    //  Use trigger to update part counts
    //
    Close;
    Parameters.Clear;
    ProcedureName:='INSERT_ShippingPartInfo;1';
    Parameters.Clear;
    Parameters.AddParameter.Name := '@ShippingID';
    Parameters.ParamValues['@ShippingID'] := fRecordID;
    Parameters.AddParameter.Name := '@part';
    Parameters.ParamValues['@part'] := partnumber;
    Parameters.AddParameter.Name := '@QTY';
    Parameters.ParamValues['@QTY'] := update;
    Parameters.AddParameter.Name := '@Date';
    Parameters.ParamValues['@Date'] := fProductionDate;
    ExecProc;
    if Inv_Connection.Errors.Count > 0 then
    begin
      ShowMessage('Unable to update part shipping part info, '+Inv_Connection.Errors.Item[0].Get_Description);
      LogActLog('ERROR','Unable to update part shipping part info, '+Inv_Connection.Errors.Item[0].Get_Description);
      raise EDatabaseError.Create('Unable to update part shipping part info');
    end
  end;
end;

function TData_Module.UpdateEINStatus:boolean;
begin
  result:=TRUE;
  try
    With Inv_StoredProc do
    Begin
      Close;
      ProcedureName := 'dbo.UPDATE_EINStatus;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@EIN';
      Parameters.ParamValues['@EIN'] := fEIN;
      Parameters.AddParameter.Name := '@EINStatus';
      Parameters.ParamValues['@EINStatus'] := fEINStatus;
      Parameters.AddParameter.Name := '@EINType';
      Parameters.ParamValues['@EINType'] := fEINType;

      fBeforeDateTime := Time;
      ExecProc;
      if Inv_Connection.Errors.Count > 0 then
      begin
        ShowMessage('Unable to update Prod Rej Info, '+Inv_Connection.Errors.Item[0].Get_Description);
        LogActLog('ERROR','Unable to update Prod Rej Info, '+Inv_Connection.Errors.Item[0].Get_Description);
        raise EDatabaseError.Create('Unable to update Prod Rej Info');
      end
      else
      begin
        fAfterDateTime := Time;
        fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
      end;
    End;    //With
    LogActLog('UPD EINStatus', 'UPDated EIN('+IntToStr(fEIN)+')', 1);
  except
    on e:exception do
    Begin
      fErrorCount := fErrorCount + 1;
      If fErrorCount < 3 Then
        UpdateRecProdRejInfo
      Else
        LogActLog('ERROR', 'FAILED UPDated EIN('+IntToStr(fEIN)+'). Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
    End;      //Except
  end;
end;



function TData_Module.GetLastProductionDate:TDate;
var
  x:integer;
begin
  //roll to last week
  result:=now-1;
  try
    With ALC_DataSet do
    Begin
      Close;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.AD_GetSpecialDates';
      Parameters.Clear;

      fBeforeDateTime := Time;
      Open;
      if Inv_Connection.Errors.Count > 0 then
      begin
        ShowMessage('Unable to get Overtime/Holiday, '+Inv_Connection.Errors.Item[0].Get_Description);
        LogActLog('ERROR','Unable to get Overtime/Holiday, '+Inv_Connection.Errors.Item[0].Get_Description);
        raise EDatabaseError.Create('Unable to get Overtime/Holiday');
      end
      else
      begin
        if DayOfTheWeek(now) = 1 then
        begin
          //check if saturday production
          if locate('Date Status Abrv;Week Number;Day Number',VarArrayOf(['O',WeekOfTheYear(now-2),DayOfTheWeek(now-2)]),[]) then
          begin
            result:=now-2;
            exit;
          end
          else
          begin
            x:=3;
            //if not move to Friday
            while locate('Date Status Abrv;Week Number;Day Number',VarArrayOf(['H',WeekOfTheYear(now-x),DayOfTheWeek(now-x)]),[])  do
              INC(x);
            result:=now-x;
            exit;
          end;
        end
        else
        begin
          x:=1;
          //if not move to Friday
          while locate('Date Status Abrv; Week Number; Day Number',VarArrayOf(['O',WeekOfTheYear(now-x),DayOfTheYear(now-x)]),[]) or (DayOfTheWeek(now-x) >= 6) do
            INC(x);
          result:=now-x;
          exit;
        end;
      end;
    end;
  except
    on e:exception do
    Begin
      fErrorCount := fErrorCount + 1;
      If fErrorCount < 3 Then
        GetOvertimeHolidayInfo
      Else
        LogActLog('ERROR', 'FAILED SELECT all Overtime/Holiday. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
    End;      //Except
  end;
end;

function TData_Module.AutoPurge: boolean;
begin
  result:=TRUE;

  try
    if fiDataRetention.AsInteger < 12 then
    begin
      ShowMessage('Invalid data rentention value ms be greater than 12, '+fiDataRetention.AsString);
      LogActLog('ERROR','Invalid data rentention value ms be greater than 12, '+fiDataRetention.AsString);
      result:=FALSE;
      exit;
    end;

    With Inv_StoredProc do
    Begin
      Close;
      ProcedureName := 'dbo.DELETE_AutoPurge;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@DataRentention';
      Parameters.ParamValues['@DataRentention'] := 0-fiDataRetention.AsInteger;

      fBeforeDateTime := Time;

      ExecProc;
      if Inv_Connection.Errors.Count > 0 then
      begin
        ShowMessage('Failed to auto purge, '+Inv_Connection.Errors.Item[0].Get_Description);
        LogActLog('ERROR','Failed to auto purge, '+Inv_Connection.Errors.Item[0].Get_Description);
        result:=FALSE;
      end
      else
      begin
        fAfterDateTime := Time;
        fDiffDateTime := (fAfterDateTime - fBeforeDateTime) * 86400000;   //Multiply 86400 by 1000 to represent milliseconds.
      end;
    End;    //With
    LogActLog('PURGE', 'Purge Data Complete', 1);
  except
    on e:exception do
    Begin
      result:=FALSE;
      LogActLog('ERROR', 'FAILED SELECT Auto Purge. Err Msg: ' + E.Message + ' Err: ' + E.ClassName)
    End;      //Except
  end;
end;

function TData_Module.IntToXlCol(ColNum: TColNum): TColString;
const
  Letters: array[0..MaxLetters] of char
='ZABCDEFGHIJKLMNOPQRSTUVWXYZ';
var
  nQuot, nMod: integer;
Begin
  if ColNum <= MaxLetters then
  begin
    Result[0] := Chr(1);
    Result[1] := Letters[ColNum];
  end
  else
  begin
    nQuot := ((ColNum - 1) * 10083) shr 18;
    nMod := ColNum - (nQuot * MaxLetters);
    Result[0] := Chr(2);
    Result[1] := Letters[nQuot];
    Result[2] := Letters[nMod];
  end;
end;

// Turns a column identifier into the corresponding integer,
//  A=1, ..., AA = 27, ..., IV = 256

function TData_Module.XlColToInt(const Col: TColString): TColNum;
const
  ASCIIOffset = Ord('A') - 1;
var
  Len: cardinal;
begin
  Len := Length(Col);
  Result := Ord(UpCase(Col[Len])) - ASCIIOffset;
  if (Len > 1) then
    Inc(Result, (Ord(UpCase(Col[1])) - ASCIIOffset) *
MaxLetters);
end;

end.
