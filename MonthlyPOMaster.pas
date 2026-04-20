//***********************************************************
//  InventorySystem
//
//  Copyright (c) 2002-2005 TAI, Failproof Manufacturing Systems
//
//***********************************************************
//
//  03/30/2005  David Verespey  Initial creation

unit MonthlyPOMaster;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, StdCtrls, Grids, DBGrids, ExtCtrls, BmDtEdit,
  NUMMIBmDateEdit, Mask, Datamodule, currEdit;

type
  TMonthlyPOMaster_Form = class(TForm)
    PartsStockMaster_Label: TLabel;
    MonthlyPO_Panel: TPanel;
    MonthlyPOMaster_DBGrid: TDBGrid;
    ManagementButtons_Panel: TPanel;
    Insert_Button: TButton;
    Update_Button: TButton;
    Search_Button: TButton;
    Clear_Button: TButton;
    Close_Button: TButton;
    Delete_Button: TButton;
    MonthlyPO_DataSource: TDataSource;
    BasisUnitPrice_Label: TLabel;
    AssyCode_Label: TLabel;
    AssyCode_ComboBox: TComboBox;
    InTransit_Label: TLabel;
    POStart_NUMMIBmDateEdit: TNUMMIBmDateEdit;
    Arrival_Label: TLabel;
    POEnd_NUMMIBmDateEdit: TNUMMIBmDateEdit;
    PONumber_Edit: TEdit;
    SizeCode_Label: TLabel;
    Label1: TLabel;
    PickUp_NUMMIBmDateEdit: TNUMMIBmDateEdit;
    Label2: TLabel;
    Label3: TLabel;
    POQty_Edit: TEdit;
    Label4: TLabel;
    POCharged_Edit: TEdit;
    PickUpTime_Edit: TMaskEdit;
    AssyCost_MaskEdit: TcurrEdit;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure MonthlyPO_DataSourceDataChange(Sender: TObject;
      Field: TField);
    procedure Search_ButtonClick(Sender: TObject);
    procedure Clear_ButtonClick(Sender: TObject);
    procedure Delete_ButtonClick(Sender: TObject);
    procedure Insert_ButtonClick(Sender: TObject);
    procedure POStart_NUMMIBmDateEditChange(Sender: TObject);
    procedure Update_ButtonClick(Sender: TObject);
  private
    { Private declarations }
    procedure SetDetailBoxes;
    procedure HoldDetails(fFromGrid: Boolean);
    function SearchGrid(AssyCode, PONumber: String): Boolean;
    procedure GetParts(tablename, selectfieldname, fPartType, returnfieldname: String; ComboBox: TComboBox);
  public
    { Public declarations }
    function Execute:boolean;
  end;

var
  MonthlyPOMaster_Form: TMonthlyPOMaster_Form;

implementation

{$R *.dfm}

procedure TMonthlyPOMaster_Form.SetDetailBoxes;
begin
  With Data_Module do
  Begin
    try
      SearchCombo(AssyCode_ComboBox, AssyCode);

      POStart_NUMMIBmDateEdit.Date:=StrTodate(copy(POStart,5,2)+'/'+copy(POStart,7,2)+'/'+copy(POStart,1,4));
      POEnd_NUMMIBmDateEdit.Date:=StrTodate(copy(POEnd,5,2)+'/'+copy(POEnd,7,2)+'/'+copy(POEnd,1,4));
      PickUp_NUMMIBmDateEdit.Date:=StrTodate(copy(PickUp,5,2)+'/'+copy(PickUp,7,2)+'/'+copy(PickUp,1,4));
      AssyCode_ComboBox.Text:=AssyCode;
      PONumber_Edit.Text:=PONumber;
      AssyCost_MaskEdit.Text:=FormatFloat('$#######0.0000', AssyCost);
      PickUpTime_Edit.Text:=PickUpTime;
      POQty_Edit.Text:=InttoStr(POQty);
      POCharged_Edit.Text:=IntToStr(POCharged);
    except
      on e:exception do
      begin
      end;
    end;
  End;      //With
end;

procedure TMonthlyPOMaster_Form.HoldDetails(fFromGrid: Boolean);
var
  ftempprice:string;
  ftempdouble:double;
begin
  If fFromGrid Then
  Begin
    with MonthlyPOMaster_DBGrid.DataSource.DataSet do
    begin
      Data_Module.POStart     := Fields[0].AsString;
      Data_Module.POEnd       := Fields[1].AsString;
      Data_Module.PickUp      := Fields[2].AsString;
      Data_Module.PickUpTime  := Fields[3].AsString;
      Data_Module.AssyCode    := Fields[4].AsString;
      Data_Module.PONumber    := Fields[5].AsString;
      Data_Module.AssyCost    := Fields[6].AsFloat;
      Data_Module.POQty       := Fields[7].AsInteger;
      Data_Module.POCharged   := Fields[8].AsInteger;
    end;
  End
  Else
  Begin
    With Data_Module do
    Begin
      POStart:=formatdatetime('yyyymmdd',POStart_NUMMIBmDateEdit.Date);
      POEnd:=formatdatetime('yyyymmdd',POEnd_NUMMIBmDateEdit.Date);
      PickUp:=formatdatetime('yyyymmdd',PickUp_NUMMIBmDateEdit.Date);
      AssyCode:=AssyCode_ComboBox.Text;
      PONumber:=PONumber_Edit.Text;
      PickUpTime:=PickUpTime_Edit.Text;
      POQty:=StrToInt(POQty_Edit.Text);

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

function TMonthlyPOMaster_Form.SearchGrid(AssyCode, PONumber: String): Boolean;
begin
  Result := False;
  Data_Module.ClearControls(MonthlyPO_Panel);
  try
    With Data_Module.Inv_DataSet Do
    Begin
      Filtered := False;
      if AssyCode <> ' ' then
        Filter := '[Assy Code] LIKE ' + QuotedStr(AssyCode);

      if PONumber <> '' then
        if Filter <> '' then
          Filter := Filter + ' AND [PO Number] LIKE ' + QuotedStr(PONumber)
        else
          Filter := '[PO Number] LIKE ' + QuotedStr(PONumber);

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
End;        //SearchGrid

function TMonthlyPOMaster_Form.Execute:boolean;
begin
  Result:= True;
  try
    try
      ShowModal;
    except
      On E:Exception do
      begin
        showMessage('Unable to generate Monthly PO screen.'
          + #13 + 'ERROR:'
          + #13 + E.Message);
      end;
    end;
  finally
    If ModalResult = mrCancel Then
      Result:= False;
  end;
end;

procedure TMonthlyPOMaster_Form.FormCreate(Sender: TObject);
begin
  With Data_Module do
  Begin
    Inv_DataSet.Filter:='';
    Inv_DataSet.Filtered:=FALSE;
    GetMonthlyPOInfo;
    MonthlyPO_DataSource.DataSet:=Inv_DataSet;
    Data_Module.ClearControls(MonthlyPO_Panel);
    Data_Module.Inv_DataSet.Filtered := False;
  End;      //With
end;

procedure TMonthlyPOMaster_Form.FormShow(Sender: TObject);
begin
  POStart_NUMMIBmDateEdit.SetFocus;
  GetParts('INV_FORECAST_DETAIL_INF', '', '', 'VC_ASSY_PART_NUMBER_CODE', AssyCode_ComboBox);
  Data_Module.Inv_DataSet.First;
end;

procedure TMonthlyPOMaster_Form.MonthlyPO_DataSourceDataChange(
  Sender: TObject; Field: TField);
begin
  HoldDetails(True);
  SetDetailBoxes;
end;

procedure TMonthlyPOMaster_Form.GetParts(tablename, selectfieldname, fPartType, returnfieldname: String; ComboBox: TComboBox);
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

procedure TMonthlyPOMaster_Form.Search_ButtonClick(Sender: TObject);
var
  ffound:boolean;
begin
  If (Trim(AssyCode_ComboBox.Text ) = '') and (Trim(PoNumber_Edit.Text ) = '') Then
    ShowMessage('Please enter a assembly code or PO Number before searching.')
  Else
  Begin
    fFound := SearchGrid(AssyCode_ComboBox.Text, PONumber_Edit.Text);
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

procedure TMonthlyPOMaster_Form.Clear_ButtonClick(Sender: TObject);
begin
  with Data_Module do
  begin
    Inv_Dataset.Filtered:=False;
    ClearControls(MonthlyPO_Panel);
  end;
  POCharged_Edit.Text:='0';
  PickUpTime_Edit.Text:='11:30';
  POStart_NUMMIBmDateEdit.SetFocus;
end;

procedure TMonthlyPOMaster_Form.Delete_ButtonClick(Sender: TObject);
begin
  HoldDetails(False);
  If MessageDlg('Are you sure you wish to delete PO Number '+ Data_Module.PONumber + #13 +
                 'Assy Number '+ Data_Module.AssyCode + #13 +
                 'From '+ Data_Module.POStart + #13 +
                 'To '+ Data_Module.POEnd + #13 +
                 ' from the database?',
                 mtWarning, [mbYes, mbNo], 0) = mrYes Then
  Begin
    Data_Module.DeleteMonthlyPOInfo;
    Data_Module.GetMonthlyPOInfo;
    Data_Module.ClearControls(MonthlyPO_Panel);
    Data_Module.Inv_DataSet.Filtered := False;
    Data_Module.Inv_DataSet.First;
  End;
  POStart_NUMMIBmDateEdit.SetFocus;
end;

procedure TMonthlyPOMaster_Form.Insert_ButtonClick(Sender: TObject);
var
  Assy,StartD,EndD,PO:string;
begin
  HoldDetails(False);
  with Data_Module do
  begin
    If Not InsertMonthlyPOInfo Then
      MessageDlg('Unable to INSERT ' + Data_Module.AssyCode, mtInformation, [mbOk], 0)
    else
    begin
      Assy:=AssyCode;
      StartD:=POStart;
      EndD:=POEnd;
      PO:=PONumber;
      GetMonthlyPOInfo;
      Inv_Dataset.Filtered := False;
      Inv_DataSet.Locate('Assy Code;PO Start;PO End;PO Number',VarArrayOf([Assy,StartD,EndD,PO]),[]);
    end;
  end;
  SetDetailBoxes;
  POStart_NUMMIBmDateEdit.SetFocus;
end;

procedure TMonthlyPOMaster_Form.Update_ButtonClick(Sender: TObject);
var
  Assy,StartD,EndD:string;
begin
  with Data_Module do
  begin
    AssyCodePrev:=AssyCode;
    POStartPrev:=POStart;
    POEndPrev:=POEnd;
    HoldDetails(False);
    UpdateMonthlyPOInfo;
    Assy:=AssyCode;
    StartD:=POStart;
    EndD:=POEnd;
    GetMonthlyPOInfo;
    Inv_Dataset.Filtered := False;
    Inv_DataSet.Locate('Assy Code;PO Start;PO End',VarArrayOf([Assy,StartD,EndD]),[]);
  end;
  SetDetailBoxes;
  POStart_NUMMIBmDateEdit.SetFocus;
end;

procedure TMonthlyPOMaster_Form.POStart_NUMMIBmDateEditChange(
  Sender: TObject);
begin
  //POEnd_NUMMIBmDateEdit.Date:=POStart_NUMMIBmDateEdit.Date+4;
end;


end.
