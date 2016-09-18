//***********************************************************
//  InventorySystem
//
//  Copyright (c) 2003 Failproof Manufacturing Systems
//
//***********************************************************
//
unit DirectorySelect;

interface

uses Windows, SysUtils, Classes, Graphics, Forms, Controls, StdCtrls, 
  Buttons, ExtCtrls, FileCtrl;

type
  TSelectDirectoryDlg = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    Bevel1: TBevel;
    Drive_ComboBox: TDriveComboBox;
    Directory_ListBox: TDirectoryListBox;
    directory_label: TLabel;
    procedure CancelBtnClick(Sender: TObject);
  private
    { Private declarations }
    fCancel:boolean;
    fdirectory:string;
  public
    { Public declarations }
    procedure Execute;

    property directory:string
    read fdirectory
    write fdirectory;

    property cancel:boolean
    read fcancel
    write fcancel;
  end;

var
  SelectDirectoryDlg: TSelectDirectoryDlg;

implementation

{$R *.dfm}

procedure TSelectDirectoryDlg.Execute;
var
  tempdir:string;
begin
  fcancel:=false;
  if fdirectory <> '' then
  begin
    tempdir:=ExtractFileDir(fdirectory);
    Drive_ComboBox.Drive:=tempdir[1];
    Directory_ListBox.Directory:=fdirectory;
  end;
  ShowModal;
  fdirectory:=directory_label.Caption;
end;
procedure TSelectDirectoryDlg.CancelBtnClick(Sender: TObject);
begin
  fcancel:=true;
end;

end.
