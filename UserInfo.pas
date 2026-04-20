//***********************************************************
//  InventorySystem
//
//  Copyright (c) 2002 Failproof Manufacturing Systems
//
//***********************************************************
//
//  10/25/2002  Aaron Huge  Initial creation

unit UserInfo;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls,
  Forms, Dialogs, ScktComp;

type
  TUserInfo = class(TComponent)
  private
    FIPAddress: String;
    FLocalName: String;
    FUserName:  String;
    function Get_IPAddress : string;
    function Get_LocalName : string;
    function Get_UserName  : string;
  public
    constructor Create(AOwner: TComponent); override;

    procedure Refresh;

    property IPAddress: String read Get_IPAddress;
    property LocalName: String read Get_LocalName;
    property UserName:  String read Get_UserName;
  end;

implementation

function TUserInfo.Get_UserName : string;
var
  Name_Size:  DWORD;
  PName:      PChar;
begin
  if FuserName = '' then
  begin
    Name_Size := 250;
    PName     := Stralloc(Name_Size + 1);
    Try
      Try
        if GetUserName(PName, Name_Size) then
          FUserName := PName
        else
          FUserName := 'N/A';
      Finally
        StrDispose(PName);
      End;
    Except;
    End;
  end;

  Result := FUserName;

end;


function TUserInfo.Get_IPAddress : string;
begin
  {if FIPAddress = '' then
    Try
      With TPowersock.Create(Self) Do
        Try
          FIPAddress := LocalIP;
        Finally
         Free;
        End;
    Except
    End;
   }
  Result := FIPAddress;

end;


function TUserInfo.Get_LocalName : String;
begin
  if FLocalName = '' then
    Try
      With TServerSocket.Create(Self) Do
        Try
          Active     := True;
          FLocalName := Socket.LocalHost;
        Finally
          Free;
        End;
    Except
    End;

  Result := FLocalName;

end;


procedure TUserInfo.Refresh;
begin
  FIPAddress := '';
  FLocalName := '';
  FUSerName  := '';
end;

constructor TUserInfo.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  Refresh;
end;

end.
