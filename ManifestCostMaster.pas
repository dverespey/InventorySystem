unit ManifestCostMaster;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Grids, DBGrids, currEdit, BmDtEdit,
  NUMMIBmDateEdit, DB;

type
  TManifestCostMaster_Form = class(TForm)
    PartsStockMaster_Label: TLabel;
    MonthlyPOMaster_DBGrid: TDBGrid;
    ManagementButtons_Panel: TPanel;
    Insert_Button: TButton;
    Update_Button: TButton;
    Search_Button: TButton;
    Clear_Button: TButton;
    Close_Button: TButton;
    Delete_Button: TButton;
    ManifestCost_Panel: TPanel;
    BasisUnitPrice_Label: TLabel;
    AssyCode_Label: TLabel;
    InTransit_Label: TLabel;
    Arrival_Label: TLabel;
    SizeCode_Label: TLabel;
    AssyCode_ComboBox: TComboBox;
    CostStart_NUMMIBmDateEdit: TNUMMIBmDateEdit;
    CostEnd_NUMMIBmDateEdit: TNUMMIBmDateEdit;
    AssyCost_MaskEdit: TcurrEdit;
    MAnifestCost_DataSource: TDataSource;
    AssyManifestID_ComboBox: TComboBox;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Clear_ButtonClick(Sender: TObject);
    procedure Search_ButtonClick(Sender: TObject);
    procedure Delete_ButtonClick(Sender: TObject);
    procedure Update_ButtonClick(Sender: TObject);
    procedure Insert_ButtonClick(Sender: TObject);
    procedure MAnifestCost_DataSourceDataChange(Sender: TObject;
      Field: TField);
  private
    { Private declarations }
    procedure SetDetailBoxes;
    procedure HoldDetails(fFromGrid: Boolean);
    function SearchGrid(AssyCode: String): Boolean;
    procedure GetParts(tablename, selectfieldname, fPartType, returnfieldname: String; ComboBox: TComboBox);
  public
    { Public declarations }
    function Execute:boolean;
  end;

var
  ManifestCostMaster_Form: TManifestCostMaster_Form;

implementation

uses DataModule;

{$R *.dfm}

procedure TManifestCostMaster_Form.SetDetailBoxes;
begin
  With Data_Module do
  Begin
    try
      SearchCombo(AssyCode_ComboBox, AssyCode);
      //AssyCode_ComboBox.Text:=AssyCode;

      SearchCombo(AssyManifestID_ComboBox, AssyManifestNo);
      //AssyCode_ComboBox.Text:=AssyManifestNo;

      CostStart_NUMMIBmDateEdit.Date:=StrTodate(copy(POStart,5,2)+'/'+copy(POStart,7,2)+'/'+copy(POStart,1,4));
      CostEnd_NUMMIBmDateEdit.Date:=StrTodate(copy(POEnd,5,2)+'/'+copy(POEnd,7,2)+'/'+copy(POEnd,1,4));

      AssyCost_MaskEdit.Text:=FormatFloat('$#######0.0000', AssyCost);

    except
      on e:exception do
      begin
      end;
    end;
  End;      //With
end;

procedure TManifestCostMaster_Form.HoldDetails(fFromGrid: Boolean);
var
  ftempprice:string;
  ftempdouble:double;
begin
  If fFromGrid Then
  Begin
    with MonthlyPOMaster_DBGrid.DataSource.DataSet do
    begin
      Data_Module.RecordID        := Fields[0].AsInteger;
      Fields[0].Visible:=FALSE;
      Data_Module.AssyCode        := Fields[1].AsString;
      Data_Module.AssyManifestNo  := Fields[2].AsString;
      Data_Module.POStart         := Fields[3].AsString;
      Data_Module.POEnd           := Fields[4].AsString;
      Data_Module.AssyCost        := Fields[5].AsFloat;
    end;
  End
  Else
  Begin
    With Data_Module do
    Begin

      POStart:=formatdatetime('yyyymmdd',CostStart_NUMMIBmDateEdit.Date);
      POEnd:=formatdatetime('yyyymmdd',CostEnd_NUMMIBmDateEdit.Date);

      AssyCode:=AssyCode_ComboBox.Text;
      AssyManifestNo  := AssyManifestID_ComboBox.Text;

      fTempPrice := AssyCost_MaskEdit.Text;
      //the 2 in the copy statement is to skip the dollar sign
      fTempPrice := Trim(Copy(fTempPrice, 2, Length(fTempPrice)));
      If Not TryStrToFloat(fTempPrice, fTempDouble) Then
      begin
        ShowMessage('Invalid assembly cost');
        AssyCost_MaskEdit.SetFocus;
      end
      Else
        AssyCost := fTempDouble;   //StrToFloat(BasisUnitPrice_Edit.Text);
    End;      //With
  End;
end;

function TManifestCostMaster_Form.SearchGrid(AssyCode: String): Boolean;
begin
  Result := False;
  Data_Module.ClearControls(ManifestCost_Panel);
  try
    With Data_Module.Inv_DataSet Do
    Begin
      Filtered := False;
      if AssyCode <> ' ' then
        Filter := '[Assy] LIKE ' + QuotedStr(AssyCode);

      Filtered := True;
      If RecordCount > 0 Then
      Begin
        HoldDetails(True);
        Result := True;
      End;
    End;    //With
  except
    on e:exception do
      ShowMessage('Error in Search' + #13 + e.Message);
  end;      //try...except
end;

procedure TManifestCostMaster_Form.GetParts(tablename, selectfieldname, fPartType, returnfieldname: String; ComboBox: TComboBox);
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
end;

function TManifestCostMaster_Form.Execute:boolean;
begin
  Result:= True;
  try
    try
      ShowModal;
    except
      On E:Exception do
      begin
        showMessage('Unable to generate Manifest Cost screen.'
          + #13 + 'ERROR:'
          + #13 + E.Message);
      end;
    end;
  finally
    If ModalResult = mrCancel Then
      Result:= False;
  end;
end;

procedure TManifestCostMaster_Form.FormCreate(Sender: TObject);
var
  i:integer;
begin

  AssyManifestID_ComboBox.Items.Clear;
  AssyManifestID_ComboBox.Items.Add(' ');
  for i:=1 to 99 do
  begin
    if i<10 then
      AssyManifestID_ComboBox.Items.Add('0'+IntToStr(i))
    else
      AssyManifestID_ComboBox.Items.Add(IntToStr(i));
  end;

  With Data_Module do
  Begin
    Inv_DataSet.Filter:='';
    Inv_DataSet.Filtered:=FALSE;
    GetManifestCostInfo;
    ManifestCost_DataSource.DataSet:=Inv_DataSet;
    Data_Module.ClearControls(ManifestCost_Panel);
    Data_Module.Inv_DataSet.Filtered := False;
  End;      //With
end;

procedure TManifestCostMaster_Form.FormShow(Sender: TObject);
begin
  AssyCode_ComboBox.SetFocus;
  GetParts('INV_FORECAST_DETAIL_INF', '', '', 'VC_ASSY_PART_NUMBER_CODE', AssyCode_ComboBox);
  Data_Module.Inv_DataSet.First;
end;

procedure TManifestCostMaster_Form.Clear_ButtonClick(Sender: TObject);
begin
  with Data_Module do
  begin
    Inv_Dataset.Filtered:=False;
    ClearControls(ManifestCost_Panel);
  end;
  AssyCode_ComboBox.SetFocus;
end;

procedure TManifestCostMaster_Form.Search_ButtonClick(Sender: TObject);
var
  ffound:boolean;
begin
  If (Trim(AssyCode_ComboBox.Text ) = '') Then
    ShowMessage('Please enter a assembly code before searching.')
  Else
  Begin
    fFound := SearchGrid(AssyCode_ComboBox.Text);
    If fFound Then
    Begin
      SetDetailBoxes;
    End
    Else
    Begin
      ShowMessage('No matches were found for your query.');
    End;
  end;
  AssyCode_ComboBox.SetFocus;
end;

procedure TManifestCostMaster_Form.Delete_ButtonClick(Sender: TObject);
begin
  HoldDetails(False);
  If MessageDlg('Are you sure you wish to delete Assy Number '+ Data_Module.AssyCode + #13 +
                 'From '+ Data_Module.POStart + #13 +
                 'To '+ Data_Module.POEnd + #13 +
                 ' from the database?',
                 mtWarning, [mbYes, mbNo], 0) = mrYes Then
  Begin
    Data_Module.DeleteManifestCostInfo;
    Data_Module.GetManifestCostInfo;
    Data_Module.ClearControls(ManifestCost_Panel);
    Data_Module.Inv_DataSet.Filtered := False;
    Data_Module.Inv_DataSet.First;
  End;
  AssyCode_ComboBox.SetFocus;
end;

procedure TManifestCostMaster_Form.Update_ButtonClick(Sender: TObject);
var
  Assy,SMan,EMan:string;
begin
  with Data_Module do
  begin
    HoldDetails(False);
    UpdateManifestCostInfo;
    Assy:=AssyCode;
    SMan:=POStart;
    EMan:=POEnd;
    GetManifestCostInfo;
    Inv_Dataset.Filtered := False;
    Inv_DataSet.Locate('Assy;Start Manifest;End Manifest',VarArrayOf([Assy,SMan,EMan]),[]);
  end;
  SetDetailBoxes;
  AssyCode_ComboBox.SetFocus;
end;

procedure TManifestCostMaster_Form.Insert_ButtonClick(Sender: TObject);
var
  Assy,SMan,EMan:string;
begin
  HoldDetails(False);
  with Data_Module do
  begin
    If Not InsertManifestCostInfo Then
      MessageDlg('Unable to INSERT ' + Data_Module.AssyCode, mtInformation, [mbOk], 0)
    else
    begin
      Assy:=AssyCode;
      SMan:=POStart;
      EMan:=POEnd;
      GetManifestCostInfo;
      Inv_Dataset.Filtered := False;
      Inv_DataSet.Locate('Assy;Start Manifest;End Manifest',VarArrayOf([Assy,SMan,EMan]),[]);
    end;
  end;
  SetDetailBoxes;
  AssyCode_ComboBox.SetFocus;
end;

procedure TManifestCostMaster_Form.MAnifestCost_DataSourceDataChange(
  Sender: TObject; Field: TField);
begin
  HoldDetails(True);
  SetDetailBoxes;
end;

end.
