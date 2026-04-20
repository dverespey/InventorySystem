unit SiteInfo;

interface

type
  TSiteInfo = class(TObject)
  private
    fSiteName:string;
    fSiteAbbr:string;
    fSiteStreet:string;
    fSiteCity:string;
    fSiteState:string;
    fSiteCountry:string;
    fSiteZip:string;
    fSiteDUNS:string;
    fSiteSupplierCode:string;
    fSiteDockCode:string;
    fSiteEIN:integer;
    fSitePassword:string;
    fSiteEDIMode:string;
    fSiteSepSegment:string;
    fSiteSepElement:string;
    fSiteSepSubElement:string;
    fSiteTMMName:string;
    fSiteTMMAbbr:string;
    fSiteTMMDUNS:string;
    fSiteMaxSequence:integer;
    fAcceptAnyOrderASN:integer;
    fSiteDeliveryMethodCode:string;
  public


    property SiteName:string
    read fSiteName;
    property SiteAbbr:string
    read fSiteAbbr;
    property SiteStreet:string
    read fSiteStreet;
    property SiteCity:string
    read fSiteCity;
    property SiteState:string
    read fSiteState;
    property SiteCountry:String
    read fSiteCountry;
    property SiteZip:string
    read fSiteZip;
    property SiteDUNS:string
    read fSiteDUNS;
    property SiteSupplierCode:string
    read fSiteSupplierCode;
    property SiteDockCode:string
    read fSiteDockCode;
    property SiteEIN:integer
    read fSiteEIN;
    property SitePassword:string
    read fSitePassword;
    property SiteEDIMode:string
    read fSiteEDIMode;
    property SiteSepSegment:string
    read fSiteSepSegment;
    property SiteSepElement:string
    read fSiteSepElement;
    property SiteSepSubElement:string
    read fSiteSepSubElement;
    property SiteTMMName:string
    read fSiteTMMName;
    property SiteTMMAbbr:string
    read fSiteTMMAbbr;
    property SiteTMMDUNS:string
    read fSiteTMMDUNS;
    property SiteMaxSequence:integer
    read fSiteMaxSequence;
    property AcceptAnyOrderASN:integer
    read fAcceptAnyOrderASN;
    property SiteDeliveryMethodCode:string
    read fSiteDeliveryMethodCode;

  end;

implementation

end.
 