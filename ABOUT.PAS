//****************************************************************
//
//       Inventory Control
//
//       Copyright (c) 2002 Failproof Manufacturing Systems.
//
//****************************************************************
//
//
//  12/16/2002  David Verespey  Initial Create
//

unit About;

interface

uses WinTypes, WinProcs, Classes, Graphics, Forms, Controls, StdCtrls,
  Buttons, ExtCtrls, ShellAPI, VersionInfo, sysutils;

type
  TAboutBox = class(TForm)
    Panel1: TPanel;
    ProgramIcon: TImage;
    ProductName: TLabel;
    Version: TLabel;
    Copyright: TLabel;
    Date: TLabel;
    Label2: TLabel;
    OkButton: TBitBtn;
    Label3: TLabel;
    procedure OKButtonClick(Sender: TObject);
    procedure ProgramIconClick(Sender: TObject);
    procedure Marquee1Done(Sender: TObject);
    procedure Label3Click(Sender: TObject);
  private
    { Private declarations }
    fprogramversion:string;
  public
    { Public declarations }
    procedure Execute;

    property ProgramVersion:string
    read fprogramversion
    write fprogramversion;
  end;

var
  AboutBox: TAboutBox;

implementation

{$R *.DFM}

procedure TAboutBox.Execute;
begin
  version.Caption:=version.Caption+fprogramversion;
  DAte.Caption:=DateTimeToStr(FileDateToDateTime(FileAge(application.ExeName)));
  Copyright.Caption:='Copyright 2002-'+formatdatetime('yyyy',now)+' TAI';
  ShowModal;
end;

procedure TAboutBox.OKButtonClick(Sender: TObject);
begin
  Close;
end;

procedure TAboutBox.ProgramIconClick(Sender: TObject);
begin
  Panel1.Visible:=False;
  OKButton.Visible:=False;
end;

procedure TAboutBox.Marquee1Done(Sender: TObject);
begin
  Panel1.Visible:=True;
  OKButton.Visible:=True;
end;

procedure TAboutBox.Label3Click(Sender: TObject);
begin

  ShellExecute(Application.Handle,
                      'open',
                      PChar('http://www.failproofmfg.com'),
                      nil,
                      nil,
                      SW_NORMAL);

end;

end.

