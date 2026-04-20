unit BCRatioMaster;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Mask, ExtCtrls, Grids, DBGrids, DB;

type
  TBCRatioMaster_Form = class(TForm)
    AssyRatioMaster_Label: TLabel;
    AssyRatioMaster_Panel: TPanel;
    BC_Panel: TPanel;
    AssyCode_Label: TLabel;
    BroadcastCode_Label: TLabel;
    BroadcastCode_Edit: TEdit;
    AssyCode_ComboBox: TComboBox;
    BCRatioMaster_DBGrid: TDBGrid;
    Label7: TLabel;
    MaskEdit1: TMaskEdit;
    Label8: TLabel;
    RadioGroup1: TRadioGroup;
    Buttons_Panel: TPanel;
    Insert_Button: TButton;
    Update_Button: TButton;
    Search_Button: TButton;
    Clear_Button: TButton;
    Close_Button: TButton;
    Delete_Button: TButton;
    PN_Panel: TPanel;
    Panel3: TPanel;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    Button5: TButton;
    Button6: TButton;
    PNRatio_DBGrid: TDBGrid;
    Label9: TLabel;
    PartNumber_ComboBox: TComboBox;
    Label1: TLabel;
    MaskEdit2: TMaskEdit;
    Button7: TButton;
    BCRatio_DataSource: TDataSource;
    PNRatio_DataSource: TDataSource;
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
    procedure SetDetailBoxes;
    procedure HoldDetails(fFromGrid: Boolean);
    function SearchGrid(BroadCode: String): Boolean;
    procedure GetParts(tablename, selectfieldname, fPartType, returnfieldname: String; ComboBox: TComboBox);
  public
    { Public declarations }
    function Execute: boolean;
  end;

var
  BCRatioMaster_Form: TBCRatioMaster_Form;

implementation

uses DataModule;

{$R *.dfm}

function TBCRatioMaster_Form.Execute: boolean;
Begin
  Result:= True;
  try
    try
      ShowModal;
    except
      On E:Exception do
        showMessage('Unable to generate ASSY Ratio Master screen.'
          + #13 + 'ERROR:'
          + #13 + E.Message);
    end;
  finally
    If ModalResult = mrCancel Then
      Result:= False;
  end;
end;    //Execute

procedure TBCRatioMaster_Form.FormCreate(Sender: TObject);
begin
  Data_Module.Inv_DataSet.Filter:='';
  Data_Module.Inv_DataSet.Filtered:=FALSE;
  Data_Module.GetBCRatioInfo;
  BCRatio_DataSource.DataSet:=Data_Module.Inv_DataSet;
  PNRatio_DataSource.DataSet:=Data_Module.Inv_Field_DataSet; //Use another Dataset
  GetParts('INV_PARTS_STOCK_MST', '', '', 'VC_PART_NUMBER', PartNumber_ComboBox);
  GetParts('INV_FORECAST_DETAIL_INF', '', '', 'VC_ASSY_PART_NUMBER_CODE', AssyCode_ComboBox);
  AssyCode_Combobox.Text := '';
  Data_Module.ClearControls(BC_Panel);
  Data_Module.ClearControls(PN_Panel);
end;


procedure TBCRatioMaster_Form.SetDetailBoxes;
begin
end;

procedure TBCRatioMaster_Form.HoldDetails(fFromGrid: Boolean);
begin
end;

function TBCRatioMaster_Form.SearchGrid(BroadCode: String): Boolean;
begin
  Result:= True;
end;

procedure TBCRatioMaster_Form.GetParts(tablename, selectfieldname, fPartType, returnfieldname: String; ComboBox: TComboBox);
var
  fTableAndWhere: String;
Begin
  try
    try
      if selectfieldname <> '' then
        fTableAndWhere := tablename+' WHERE '+selectfieldname+' = ' + QuotedStr(fPartType)
      else
        fTableAndWhere := tablename;

      Data_Module.SelectSingleField(fTableAndWhere, returnfieldname, ComboBox);
    except
      On e:exception do
      Begin
        MessageDlg('Unable to get a list of parts.', mtError, [mbOK], 0);
      End;
    end;      //try...except
  finally
  end;      //try...finally
End;      //GetParts

end.
