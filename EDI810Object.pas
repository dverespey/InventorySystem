//****************************************************************
//
//       Admin
//
//       Copyright (c) 2006 Failproof Manufacturing Systems.
//
//****************************************************************
//
//  03/01/2006  David Verespey  Initial creation

unit EDI810Object;

interface

uses SysUtils, DB, Dialogs, Classes;

type
  T810EDI = class(TObject)
  private
    fEDIRecord:TStringList;
    fISAHeader:string;
    fGSHeader:string;
    fSTHeader:string;
    fBigHeader:string;
    fCTTTrailer:string;
    fSETrailer:string;
    fGETrailer:string;
    fIEATrailer:string;
    fLineName:string;
    fStartDate:TDateTime;
    fEndDate:TDateTime;
    fStartSeq:integer;
    fEndSeq:integer;
    fSepSegment:string;
    fSepElement:string;
    fSepSubElement:string;
    fein:integer;
    f810Time:TDateTime;
    fID1count:integer;
    fSECount:integer;
    fPickUpDate:string;
    function ISAHeader:boolean;
    function GSHeader:boolean;
    function STHeader:boolean;
    function BigHeader:boolean;
    function IT1Loop:boolean;
    function CTTTrailer:boolean;
    function SETrailer:boolean;
    function GETTrailer:boolean;
    function IEATrailer:boolean;
  public
    function Execute:boolean;

    property EDIRecord:TStringList
    read fEDIRecord;

    property LineName:string
    read fLineName
    write fLineName;

    property StartDate:TDateTime
    read fStartDate
    write fStartDate;

    property EndDate:TDateTime
    read fEndDate
    write fEndDate;

    property StartSeq:integer
    read fStartSeq
    write fStartSeq;

    property EndSeq:integer
    read fEndSeq
    write fEndSeq;

    property PickUpDate:string
    read fPickUpDate
    write fPickUpDate;

    property EIN:integer
    read fEIN
    write fEIN;

  end;

implementation

uses DataModule;

function T810EDI.Execute:boolean;
begin
  result:=FALSE;
  try
    fEDIRecord:=TStringList.Create;
    f810Time:=now;
    fID1count:=0;
    fSECount:=0;

    //  Get Site Info
    if not Data_Module.SiteDataset.Active then
      Data_Module.SiteDataset.Open;

    // Get Seperators
    fSepSegment:=Data_Module.SiteDataset.FieldByName('SiteSepSegment').AsString;
    fSepElement:=Data_Module.SiteDataset.FieldByName('SiteSepElement').AsString;
    fSepSubElement:=Data_Module.SiteDataset.FieldByName('SiteSepSubElement').AsString;

    if Data_Module.EDI810DataSet.RecordCount > 0 then
    begin
      //  REcord
      if ISAHeader then
        if GSHeader then
          if STHeader then
            if BigHeader then
              if IT1Loop then
                if CTTTrailer then
                  if SETrailer then
                    if GETTrailer then
                      if IEATrailer then
                        result:=TRUE;
    end
    else
    begin
      ShowMessage('No records to process');
    end;
  except
    on e:exception do
    begin
      ShowMessage('Failed on EDI create, '+e.Message);
    end;
  end;
end;

function T810EDI.ISAHeader:boolean;
begin
  result:=FALSE;
  try
    fISAHeader:='';
    fISAHeader:='ISA'+fSepElement;
    fISAHeader:=fISAHeader+'00'+fSepElement;  //1
    fISAHeader:=fISAHeader+format('%-10s',[Data_Module.SiteDataset.FieldByName('SiteAbbr').AsString])+fSepElement;  //2
    fISAHeader:=fISAHeader+'01'+fSepElement;  //3
    fISAHeader:=fISAHeader+format('%-10s',[Data_Module.SiteDataset.FieldByName('SiteDUNS').AsString])+fSepElement;  //4
    fISAHeader:=fISAHeader+'ZZ'+fSepElement;  //5
    fISAHeader:=fISAHeader+format('%-15s',[Data_Module.SiteDataset.FieldByName('SiteDUNS').AsString+'-'+Data_Module.SiteDataset.FieldByName('SiteSupplierCode').AsString])+fSepElement; //6
    fISAHeader:=fISAHeader+'01'+fSepElement;  //7
    fISAHeader:=fISAHeader+format('%-15s',[Data_Module.SiteDataset.FieldByName('SiteTMMDUNS').AsString])+fSepElement;  //8
    fISAHeader:=fISAHeader+formatdatetime('yymmdd',f810Time)+fSepElement;  //9
    fISAHeader:=fISAHeader+formatdatetime('hhmm',f810Time)+fSepElement;  //10
    fISAHeader:=fISAHeader+'U'+fSepElement; //11
    fISAHeader:=fISAHeader+'00400'+fSepElement;   //12
    // Create Next 810 ein
    //
    //  For recreate use the orginal EIN, passed in on the object call, cancel that for now
    //
    //if fein = -1 then
    //fein:=Data_Module.SiteDataset.FieldByName('SiteEIN').AsInteger+1;
    //INC(fein);
    fISAHeader:=fISAHeader+Format('%9.9d', [fein])+fSepElement;  //13
    fISAHeader:=fISAHeader+'0'+fSepElement; //14
    fISAHeader:=fISAHeader+Data_Module.SiteDataset.FieldByName('SiteEDIMode').AsString+fSepElement; //15
    fISAHeader:=fISAHeader+fSepSubElement;  //16

    fEDIRecord.Add(fISAHeader);
    result:=TRUE;
  except
    on e:exception do
    begin
      ShowMessage('Failed on ISA Header create, '+e.Message);
    end;
  end;
end;

function T810EDI.GSHeader:boolean;
begin
  result:=FALSE;
  try
    fGSHeader:='';
    fGSHeader:='GS'+fSepElement;
    fGSHeader:=fGSHeader+'IN'+fSepElement;
    fGSHeader:=fGSHeader+Data_Module.SiteDataset.FieldByName('SiteDUNS').AsString+fSepElement;
    fGSHeader:=fGSHeader+Data_Module.SiteDataset.FieldByName('SiteTMMDUNS').AsString+fSepElement;
    fGSHeader:=fGSHeader+formatdatetime('yyyymmdd',f810Time)+fSepElement;
    fGSHeader:=fGSHeader+formatdatetime('hhmm',f810Time)+fSepElement;
    fGSHeader:=fGSHeader+Format('%9.9d', [fein])+fSepElement;
    fGSHeader:=fGSHeader+'X'+fSepElement;
    fGSHeader:=fGSHeader+'004010';

    fEDIRecord.Add(fGSHeader);
    result:=TRUE;
  except
    on e:exception do
    begin
      ShowMessage('Failed on ISA Header create, '+e.Message);
    end;
  end;
end;

function T810EDI.STHeader:boolean;
begin
  result:=FALSE;
  try
    fSTHeader:='';
    fSTHeader:='ST'+fSepElement;
    fSTHeader:=fSTHeader+'810'+fSepElement;
    fSTHeader:=fSTHeader+Format('%9.9d', [fein]);

    fEDIRecord.Add(fSTHeader);
    INC(fSECount);
    result:=TRUE;
  except
    on e:exception do
    begin
      ShowMessage('Failed on ST Header create, '+e.Message);
    end;
  end;
end;

function T810EDI.BigHeader:boolean;
begin
  result:=FALSE;
  try
    fBigHeader:='';

    fBigHeader:='BIG'+fSepElement;
    fBigHeader:=fBigHeader+formatdatetime('yyyymmdd',f810Time)+fSepElement;
    fBigHeader:=fBigHeader+Data_Module.SiteDataset.FieldByName('SiteSupplierCode').AsString;

    fEDIRecord.Add(fBIGHeader);
    INC(fSECount);
    result:=TRUE;
  except
    on e:exception do
    begin
      ShowMessage('Failed on Big Header create, '+e.Message);
    end;
  end;
end;

function T810EDI.IT1Loop:boolean;
var
  IT1List:TStringList;
  IT1Item:string;
  REFItem:string;
  DTMItem:string;
  TDSItem:string;
  lastManifest,lastPickUp:string;
  totalstr,temp:string;
  i:integer;
  total:extended;
begin
  total:=0;
  result:=FALSE;
  IT1List:=TStringList.Create;
  try
    IT1Item:='';
    lastManifest:=Data_Module.EDI810DataSet.FieldByName('Manifest').AsString;
    lastPickUp:=Data_Module.EDI810DataSet.FieldByName('PickUpDate').AsString;
    while not Data_Module.EDI810DataSet.Eof do
    begin
      // New pick date requires new Invoice file
      if lastPickUp <> Data_Module.EDI810DataSet.FieldByName('PickUpDate').AsString then
      begin
        break;
      end;

      if (lastManifest <> Data_Module.EDI810DataSet.FieldByName('Manifest').AsString) then
      begin
        // Add REF segment
        REFItem:='REF'+fSepElement;
        REFItem:=REFItem+'MK'+fSepElement;
        REFItem:=REFItem+lastManifest;//Data_Module.EDI810DataSet.FieldByName('Manifest').AsString;
        lastManifest:=Data_Module.EDI810DataSet.FieldByName('Manifest').AsString;
        lastPickUp:=Data_Module.EDI810DataSet.FieldByName('PickUpDate').AsString;

        IT1List.Add(REFItem);
        INC(fSECount);

        // Add DTM segment
        DTMItem:='DTM'+fSepElement;
        DTMItem:=DTMItem+'050'+fSepElement;
        DTMItem:=DTMItem+Data_Module.EDI810DataSet.FieldByName('PickUpDate').AsString;//formatdatetime('yyyymmdd',f810Time);

        IT1List.Add(DTMItem);
        INC(fSECount);
      end;

      IT1Item:='IT1'+fSepElement;

      if copy(Data_Module.EDI810DataSet.FieldByName('Manifest').AsString,1,1) = '7' then
        IT1Item:=IT1Item+'M391'+fSepElement // Use code 391 for broadcast data
      else
        IT1Item:=IT1Item+'M390'+fSepElement; // Use code 390 for non-broadcast parts(HotCall)

      IT1Item:=IT1Item+IntToStr(Data_Module.EDI810DataSet.FieldByName('ShipQty').AsInteger)+fSepElement;
      IT1Item:=IT1Item+'EA'+fSepElement;
      IT1Item:=IT1Item+FloatToStr(Data_Module.EDI810DataSet.FieldByName('UnitPrice').AsCurrency)+fSepElement;
      IT1Item:=IT1Item+'QT'+fSepElement;
      IT1Item:=IT1Item+'PN'+fSepElement;
      IT1Item:=IT1Item+Data_Module.EDI810DataSet.FieldByName('PartNumber').AsString+fSepElement;
      IT1Item:=IT1Item+'PK'+fSepElement;
      IT1Item:=IT1Item+'1'+fSepElement;
      IT1Item:=IT1Item+'ZZ'+fSepElement;
      IT1Item:=IT1Item+Data_Module.SiteDataset.FieldByName('SiteDockCode').AsString;

      total:=total+(Data_Module.EDI810DataSet.FieldByName('UnitPrice').AsCurrency*Data_Module.EDI810DataSet.FieldByName('ShipQty').AsInteger);
      IT1List.Add(IT1Item);
      INC(fID1Count);
      INC(fSECount);

      Data_Module.EDI810DataSet.Next;
      if Data_Module.EDI810DataSet.Eof then
      begin
        lastManifest:=Data_Module.EDI810DataSet.FieldByName('Manifest').AsString;
        lastPickUp:=Data_Module.EDI810DataSet.FieldByName('PickUpDate').AsString;
      end;
    end;

    REFItem:='REF'+fSepElement;
    REFItem:=REFItem+'MK'+fSepElement;
    REFItem:=REFItem+lastManifest;

    IT1List.Add(REFItem);
    INC(fSECount);

    // Add DTM segment
    DTMItem:='DTM'+fSepElement;
    DTMItem:=DTMItem+'050'+fSepElement;
    DTMItem:=DTMItem+lastPickUp;//formatdatetime('yyyymmdd',f810Time);
    fPickUpDate:=lastPickUp;

    IT1List.Add(DTMItem);
    INC(fSECount);

    // Do DTS total
    TDSItem:='TDS'+fSepElement;

    totalstr:=floattostr(total);
    temp:=copy(totalstr,1,pos('.',totalstr)-1);
    totalstr:=copy(totalstr,pos('.',totalstr)+1,4);
    if length(totalstr) < 4 then
    begin
      if length(totalstr) = 3 then
        totalstr:=totalstr+'0'
      else if length(totalstr) = 2 then
        totalstr:=totalstr+'00';
    end;

    TDSItem:=TDSItem+temp+totalstr;//floattostr(total);

    IT1List.Add(TDSItem);
    INC(fSECount);


    for i:=0 to IT1List.Count-1 do
    begin
      fEDIRecord.Add(IT1List[i]);
    end;

    IT1List.Free;
    result:=TRUE;
  except
    on e:exception do
    begin
      ShowMessage('Failed on IT1Loop create, '+e.Message);
    end;
  end;
end;

function T810EDI.CTTTrailer:boolean;
begin
  result:=FALSE;
  try
    fCTTTrailer:='';

    fCTTTrailer:='CTT'+fSepElement;
    fCTTTrailer:=fCTTTrailer+IntToStr(fID1Count);

    fEDIRecord.Add(fCTTTrailer);
    INC(fSECount);
    result:=TRUE;
  except
    on e:exception do
    begin
      ShowMessage('Failed on CTT Trailer create, '+e.Message);
    end;
  end;
end;

function T810EDI.SETrailer:boolean;
begin
  result:=FALSE;
  try
    INC(fSECount);
    fSETrailer:='';

    fSETrailer:='SE'+fSepElement;
    fSETrailer:=fSETrailer+IntToStr(fSECount)+fSepElement;
    fSETrailer:=fSETrailer+Format('%9.9d', [fein]);

    fEDIRecord.Add(fSETrailer);
    result:=TRUE;
  except
    on e:exception do
    begin
      ShowMessage('Failed on SE Trailer create, '+e.Message);
    end;
  end;
end;

function T810EDI.GETTrailer:boolean;
begin
  result:=FALSE;
  try
    fGETrailer:='';

    fGETrailer:='GE'+fSepElement;
    fGETrailer:=fGETrailer+'1'+fSepElement;
    fGETrailer:=fGETrailer+Format('%9.9d', [fein]);

    fEDIRecord.Add(fGETrailer);
    result:=TRUE;
  except
    on e:exception do
    begin
      ShowMessage('Failed on GE Trailer create, '+e.Message);
    end;
  end;
end;

function T810EDI.IEATrailer:boolean;
begin
  result:=FALSE;
  try
    fIEATrailer:='';

    fIEATrailer:='IEA'+fSepElement;
    fIEATrailer:=fIEATrailer+'1'+fSepElement;
    fIEATrailer:=fIEATrailer+Format('%9.9d', [fein]);

    fEDIRecord.Add(fIEATrailer);
    result:=TRUE;
  except
    on e:exception do
    begin
      ShowMessage('Failed on ISA Trailer create, '+e.Message);
    end;
  end;
end;
end.
