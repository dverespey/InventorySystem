unit ExploreFiles;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, FileCtrl, StdCtrls, ExtCtrls;

type
  TExploreFiles_Form = class(TForm)
    Panel2: TPanel;
    FilterComboBox1: TFilterComboBox;
    DirectoryListBox1: TDirectoryListBox;
    FileListBox1: TFileListBox;
    Panel3: TPanel;
    Panel1: TPanel;
    Label1: TLabel;
    DriveComboBox1: TDriveComboBox;
    Open_Button: TButton;
    Cancel_Button: TButton;
  private
    { Private declarations }
    fFilePathAndName: String;
  public
    { Public declarations }
    property FilePathAndName: String read fFilePathAndName;
    function Execute: Boolean;
  end;

var
  ExploreFiles_Form: TExploreFiles_Form;

implementation

{$R *.dfm}

function TExploreFiles_Form.Execute: boolean;
Begin
  Result:= True;
  try
    try
//      data_module.ADODataSet1.Open;
      ShowModal;
      fFilePathAndName := Label1.Caption + '\' + FileListBox1.Items.Text;
    except
      On E:Exception do
        showMessage('Unable to browse for files at this time.'
          + #13 + 'ERROR:'
          + #13 + E.Message);
    end;
  finally
{
    data_module.ADODataSet1.Close;
    data_module.ADOConnection1.Close;
}
    If ModalResult = mrCancel Then
      Result:= False;
  end;

end;      //Execute

end.
