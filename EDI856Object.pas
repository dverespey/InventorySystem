//****************************************************************
//
//       Admin
//
//       Copyright (c) 2006 Failproof Manufacturing Systems.
//
//****************************************************************
//
//  03/01/2006  David Verespey  Initial creation

unit EDI856Object;

interface

uses SysUtils, DB, Dialogs, Classes;

type
  T856EDI = class(TObject)
  private
    fEDIRecord:TStringList;
    fISAHeader:string;
    fGSHeader:string;
    fSTHeader:string;
    fBSNHeader:string;
    fDTMHeader:string;
    fCTTTrailer:string;
    fSETrailer:string;
    fGETrailer:string;
    fIEATrailer:string;
    fLineName:string;
    fStartDate:string;
    fEndDate:string;
    fStartSeq:integer;
    fEndSeq:integer;
    fSepSegment:string;
    fSepElement:string;
    fSepSubElement:string;
    fein:integer;
    f810Time:TDateTime;
    fHLcount:integer;
    fSegCount:integer;
    fPickupDate:string;
    function ISAHeader:boolean;
    function GSHeader:boolean;
    function STHeader:boolean;
    function BSNHeader:boolean;
    function DTMHeader:boolean;
    function HLLoop:boolean;
    function CTTTrailer:boolean;
    function SETrailer:boolean;
    function GETTrailer:boolean;
    function IEATrailer:boolean;
  public
    function Execute:boolean;

    property EDIRecord:TStringList
    read fEDIRecord;

    property PickupDate:string
    read fPickupDate
    write fPickupDate;

    property LineName:string
    read fLineName
    write fLineName;

    property StartDate:string
    read fStartDate
    write fStartDate;

    property EndDate:string
    read fEndDate
    write fEndDate;

    property StartSeq:integer
    read fStartSeq
    write fStartSeq;

    property EndSeq:integer
    read fEndSeq
    write fEndSeq;
  end;

implementation

uses DataModule;


function T856EDI.Execute:boolean;
begin
  result:=FALSE;
  try
    fEDIRecord:=TStringList.Create;
    f810Time:=now;
    fHLcount:=0;
    fSegCount:=0;

    //  Get Site Info
    if not Data_Module.SiteDataset.Active then
      Data_Module.SiteDataset.Open;

    // Get Seperators
    fSepSegment:=Data_Module.SiteDataset.FieldByName('SiteSepSegment').AsString;
    fSepElement:=Data_Module.SiteDataset.FieldByName('SiteSepElement').AsString;
    fSepSubElement:=Data_Module.SiteDataset.FieldByName('SiteSepSubElement').AsString;

    //Pickup Date
    fPickupDate:=Data_Module.EDI856DataSet.FieldByName('PickUpDate').AsString;

    if Data_Module.EDI856DataSet.RecordCount > 0 then
    begin
      //  REcord
      if ISAHeader then
        if GSHeader then
          if STHeader then
            if BSNHeader then
              if DTMHeader then
                if HLLoop then
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

function T856EDI.ISAHeader:boolean;
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
    fISAHeader:=fISAHeader+copy(fPickupDate,3,6)+fSepElement;  //9
    fISAHeader:=fISAHeader+formatdatetime('hhmm',f810Time)+fSepElement;  //10
    fISAHeader:=fISAHeader+'U'+fSepElement; //11
    fISAHeader:=fISAHeader+'00400'+fSepElement;   //12
    fein:=Data_Module.EDI856DataSet.FieldByName('SiteEIN').AsInteger;      //EIN is incremented elsewhere now
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

function T856EDI.GSHeader:boolean;
begin
  result:=FALSE;
  try
    fGSHeader:='';
    fGSHeader:='GS'+fSepElement;
    fGSHeader:=fGSHeader+'SH'+fSepElement;
    fGSHeader:=fGSHeader+Data_Module.SiteDataset.FieldByName('SiteDUNS').AsString+fSepElement;
    fGSHeader:=fGSHeader+Data_Module.SiteDataset.FieldByName('SiteTMMDUNS').AsString+fSepElement;
    fGSHeader:=fGSHeader+fPickupDate+fSepElement;
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

function T856EDI.STHeader:boolean;
begin
  result:=FALSE;
  try
    fSTHeader:='';
    fSTHeader:='ST'+fSepElement;
    fSTHeader:=fSTHeader+'856'+fSepElement;
    fSTHeader:=fSTHeader+Format('%9.9d', [fein]);

    fEDIRecord.Add(fSTHeader);
    INC(fSegCount);
    result:=TRUE;
  except
    on e:exception do
    begin
      ShowMessage('Failed on ST Header create, '+e.Message);
    end;
  end;
end;

function T856EDI.BSNHeader:boolean;
begin
  result:=FALSE;
  try
    fBSNHeader:='';

    fBSNHeader:='BSN'+fSepElement;
    fBSNHeader:=fBSNHeader+'00'+fSepElement;
    fBSNHeader:=fBSNHeader+fPickupDate+Format('%9.9d', [fein])+fSepElement; //Shipment ID unique to each shipment
    fBSNHeader:=fBSNHeader+fPickupDate+fSepElement;
    fBSNHeader:=fBSNHeader+formatdatetime('hhmm',f810Time);//+fSepElement;

    fEDIRecord.Add(fBSNHeader);
    INC(fSegCount);
    result:=TRUE;
  except
    on e:exception do
    begin
      ShowMessage('Failed on BSN Header create, '+e.Message);
    end;
  end;
end;

function T856EDI.DTMHeader:boolean;
begin
  result:=FALSE;
  try
    fDTMHeader:='';

    fDTMHeader:='DTM'+fSepElement;
    fDTMHeader:=fDTMHeader+'011'+fSepElement;
    fDTMHeader:=fDTMHeader+fPickupDate+fSepElement;
    fDTMHeader:=fDTMHeader+formatdatetime('hhmm',f810Time)+fSepElement;
    fDTMHeader:=fDTMHeader+'ET';

    fEDIRecord.Add(fDTMHeader);
    INC(fSegCount);
    result:=TRUE;
  except
    on e:exception do
    begin
      ShowMessage('Failed on DTM Header create, '+e.Message);
    end;
  end;
end;

function T856EDI.HLLoop:boolean;
var
  HLList:TStringList;
  HLItem:string;
  TD1Item:String;
  TD5Item:string;
  TD3Item:string;
  PRFItem:string;
  LINItem:string;
  SN1Item:string;
  lastSiteEIN,lastManifest:string;
  fHID,fParent,i:integer;
begin
  fHID:=1;
  result:=FALSE;
  lastmanifest:='';
  HLList:=TStringList.Create;

  try
    // Do Shipping
    HLItem:='HL'+fSepElement;
    HLItem:=HLItem+IntToStr(fHID)+fSepElement+fSepElement;
    fParent:=fHID;
    INC(fHID);
    HLItem:=HLItem+'S'+fSepElement;
    HLItem:=HLItem+'1';
    HLList.Add(HLItem);
    INC(fHLCount);

    TD1Item:='TD1'+fSepElement;
    TD1Item:=TD1Item+fSepElement;
    HLList.Add(TD1Item);

    TD5Item:='TD5'+fSepElement;
    TD5Item:=TD5Item+'B'+fSepElement;
    TD5Item:=TD5Item+'25'+fSepElement;
    TD5Item:=TD5Item+'00000'+fSepElement;
    TD5Item:=TD5Item+Data_Module.SiteDataset.FieldByName('SiteDeliveryMethodCode').AsString;//  'CV';
    HLList.Add(TD5Item);

    TD3Item:='TD3'+fSepElement;
    TD3Item:=TD3Item+'TL'+fSepElement;
    TD3Item:=TD3Item+fSepElement;
    TD3Item:=TD3Item+'1234567890';
    HLList.Add(TD3Item);

    lastSiteEIN:=Data_Module.EDI856DataSet.FieldByName('SiteEIN').AsString;
    while not Data_Module.EDI856DataSet.Eof do
    begin
      if  lastSiteEIN <> Data_Module.EDI856DataSet.FieldByName('SiteEIN').AsString then
        break;

      if (lastManifest <> Data_Module.EDI856DataSet.FieldByName('Manifest').AsString) then
      begin
        // New Order
        HLItem:='HL'+fSepElement;
        HLItem:=HLItem+IntToStr(fHID)+fSepElement;
        HLItem:=HLItem+'1'{ Shipment is always order parent IntToStr(fParent)}+fSepElement;
        HLItem:=HLItem+'O'+fSepElement;
        HLItem:=HLItem+'1';
        fParent:=fHID;
        INC(fHID);
        HLList.Add(HLItem);
        INC(fHLCount);

        PRFItem:='PRF'+fSepElement;
        PRFItem:=PRFItem+Data_Module.EDI856DataSet.FieldByName('Manifest').AsString+'-'+Data_Module.EDI856DataSet.FieldByName('Manifest').AsString;
        HLList.Add(PRFItem);

        lastManifest:=Data_Module.EDI856DataSet.FieldByName('Manifest').AsString;

      end;
      // Order Items
      HLItem:='HL'+fSepElement;
      HLItem:=HLItem+IntToStr(fHID)+fSepElement;
      INC(fHID);
      HLItem:=HLItem+IntToStr(fParent)+fSepElement;
      HLItem:=HLItem+'I'+fSepElement;
      HLItem:=HLItem+'0';
      HLList.Add(HLItem);
      INC(fHLCount);

      LINItem:='LIN'+fSepElement;
      LINItem:=LINItem+fSepElement;
      LINItem:=LINItem+'BP'+fSepElement;
      LINItem:=LINItem+Data_Module.EDI856DataSet.FieldByName('PartNumber').AsString+fSepElement;
      LINItem:=LINItem+'RC'+fSepElement;
      LINItem:=LINItem+Data_Module.EDI856DataSet.FieldByName('Kanban').AsString+fSepElement;
      HLList.Add(LINItem);

      SN1Item:='SN1'+fSepElement;
      SN1Item:=SN1Item+fSepElement;
      SN1Item:=SN1Item+IntToStr(Data_Module.EDI856DataSet.FieldByName('ShipQty').AsInteger)+fSepElement;
      SN1Item:=SN1Item+'PC';
      HLList.Add(SN1Item);

      Data_Module.EDI856DataSet.Next;
    end;

    for i:=0 to HLList.Count-1 do
    begin
      fEDIRecord.Add(HLList[i]);
      INC(fSegCount);
    end;

    HLList.Free;
    result:=TRUE;
  except
    on e:exception do
    begin
      ShowMessage('Failed on HLLoop create, '+e.Message);
    end;
  end;
end;

function T856EDI.CTTTrailer:boolean;
begin
  result:=FALSE;
  try
    fCTTTrailer:='';

    fCTTTrailer:='CTT'+fSepElement;
    fCTTTrailer:=fCTTTrailer+IntToStr(fHLCount);

    fEDIRecord.Add(fCTTTrailer);
    INC(fSegCount);
    result:=TRUE;
  except
    on e:exception do
    begin
      ShowMessage('Failed on CTT Trailer create, '+e.Message);
    end;
  end;
end;

function T856EDI.SETrailer:boolean;
begin
  result:=FALSE;
  try
    fSETrailer:='';

    fSETrailer:='SE'+fSepElement;
    INC(fSegCount);
    fSETrailer:=fSETrailer+IntToStr(fSegCount)+fSepElement;
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

function T856EDI.GETTrailer:boolean;
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

function T856EDI.IEATrailer:boolean;
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
