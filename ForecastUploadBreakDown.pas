unit ForecastUploadBreakDown;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, DataModule, ComCtrls, StdCtrls,
  DirOutln, OleServer, Excel2000, Excel97;

type
  TForeUpBreakDown_Form = class(TForm)
    ForeUpBreakDown_Label: TLabel;
    FileName_Label: TLabel;
    FileName_Edit: TEdit;
    Browse_Button: TButton;
    ForecastFilleNameDialog: TOpenDialog;
    Start_Button: TButton;
    Close_Button: TButton;
    procedure Browse_ButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Start_ButtonClick(Sender: TObject);
    procedure FileName_EditChange(Sender: TObject);
  private
    { Private declarations }
    fDefaultFilePathLabel: String;
    fFilePathLabel: String;
    fFileSelected:boolean;
  public
    { Public declarations }
    property FilePathLabel: String read fFilePathLabel write fFilePathLabel;
    function Execute: boolean;
  end;

var
  ForeUpBreakDown_Form: TForeUpBreakDown_Form;

implementation

uses ForecastBreakdownF;

{$R *.dfm}

function TForeUpBreakDown_Form.Execute: boolean;
Begin
  Result:= True;
  try
    try
      fFileSelected:=False;
      ShowModal;
    except
      On E:Exception do
        showMessage('Unable to generate Forecast Info. Upload & Break Down screen.'
          + #13 + 'ERROR:'
          + #13 + E.Message);
    end;
  finally
    If ModalResult = mrCancel Then
      Result:= False;
  end;

end;      //Execute

procedure TForeUpBreakDown_Form.Browse_ButtonClick(Sender: TObject);
begin
  with ForecastFilleNameDialog do
  begin
    Filter := 'Text Forecast File (*.txt)|*.txt';
    DefaultExt := 'txt';
    Filename := 'prel';
    Options := [ofHideReadOnly, ofFileMustExist,
      ofPathMustExist];
    if Execute then
    begin
      if ofExtensionDifferent in Options then
        MessageDlg ('Not a file with the .TXT extension',
          mtError, [mbOK], 0)
      else
        FilePathLabel := FileName;

      FileName_Edit.Text := FilePathLabel;
      fFileselected:=TRUE;
    end;
  end;
end;

procedure TForeUpBreakDown_Form.FormCreate(Sender: TObject);
begin
  fDefaultFilePathLabel := '[Type file path and name here or click Browse]';
  FilePathLabel := fDefaultFilePathLabel;
  FileName_Edit.Text := FilePathLabel;
end;

procedure TForeUpBreakDown_Form.Start_ButtonClick(Sender: TObject);
var
  breakdown:TForecastBreakdown_Form;
begin
  try
    if fFileSelected then
    begin
      breakdown:=TForecastBreakdown_Form.Create(self);
      breakdown.filename:=ForecastFilleNameDialog.FileName;
      breakdown.SupplierCode:=data_Module.fiSupplierCode.AsString;
      breakdown.FileKind:=fText;
      breakdown.Show;
      if not breakdown.Execute then
        Showmessage('Forecast has not been added, please retry');
      breakdown.Free;
      close;
    end
    else
      ShowMessage('Please select a valid file first');
  except
    on e:exception do
      ShowMessage('Err: ' + e.Message);
  end;
end;

procedure TForeUpBreakDown_Form.FileName_EditChange(Sender: TObject);
begin
  fFileSelected:=True;
end;

end.
