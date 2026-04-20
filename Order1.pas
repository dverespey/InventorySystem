//***********************************************************
//  InventorySystem
//
//  Copyright (c) 2002 Failproof Manufacturing Systems
//
//***********************************************************
//
//  10/25/2002  Aaron Huge      Initial creation
//  12/30/2002  David Verespey  Add the guts

unit Order;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, ComCtrls, OleServer, Excel2000,ComObj,ADOdb,DateUtils;

type
  TColNum = 1..256;  // Excel columns only go up to IV
  TColString = string[2];

  TOrder_Form = class(TForm)
    Start_Button: TButton;
    Close_Button: TButton;
    Order_Label: TLabel;
    Label1: TLabel;
    TireWheel_RadioGroup: TRadioGroup;
    Date_DateTimePicker: TDateTimePicker;
    Line_RadioGroup: TRadioGroup;
    procedure FormCreate(Sender: TObject);
    procedure Start_ButtonClick(Sender: TObject);
    procedure Close_ButtonClick(Sender: TObject);
  private
    { Private declarations }
    fpartline : array[1..100,1..2] of string;
    fpartcount:integer;
    fDates:array[0..20] of Double;
    fForecast:array[0..20] of integer;
    fLine:string;
    procedure DoBoarders(topedge,bottomedge:integer;startcolumn,endcolumn:string;fill:boolean=FALSE);
    procedure UpdateForecast(partno:string);
    procedure FillForecast(line:integer);
    procedure ClearForecast;
    procedure DoFormulas(topedge,bottomedge:integer);
    function XlColToInt(const Col: TColString): TColNum;
    function IntToXlCol(ColNum: TColNum): TColString;
    procedure UpdateSizeInfo(size,assyline:string;line:integer);
    procedure DoLeadTime(Leadtime,Line:integer);
  public
    { Public declarations }
    function Execute: boolean;
  end;

var
  Order_Form: TOrder_Form;
  excel,mysheet:variant;

implementation

uses DataModule;

const
  MaxLetters = 26;

{$R *.dfm}

function TOrder_Form.Execute: boolean;
Begin
  Result:= True;
  try
    try
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
end;    //Execute

procedure TOrder_Form.FormCreate(Sender: TObject);
begin
    Date_DateTimePicker.MinDate := Date - 30;
    Date_DateTimePicker.MaxDate := Date;
    Date_DateTimePicker.Date := StrToDate(FormatDateTime('mm/dd/yyyy', Date));

end;

procedure TOrder_Form.Start_ButtonClick(Sender: TObject);
var
  i,x,topedge,bottomedge:integer;
  lastsup,lastsiz:string;
  firstheader,fst:boolean;
begin

  fpartcount:=0;
  firstheader:=TRUE;
  fst:=TRUE;
  excel := createOleObject('Excel.Application');
  excel.visible := True;
  excel.DisplayAlerts:=False;

  if Line_RadioGroup.ItemIndex = 0 then
    fLine := 'C'
  else
    fLine := 'T';

  if TireWheel_RadioGroup.ItemIndex = 0 then
    excel.workbooks.open(Data_Module.fiInputFileDir.AsString+'TireOrder.xls')
  else
    excel.workbooks.open(Data_Module.fiInputFileDir.AsString+'WheelOrder.xls');

  mysheet := excel.workSheets[1];

  {for i:=1 to 56 do
  begin
    mysheet.Range[IntToXLCol(i)+'1'].Interior.ColorIndex:=i;
  end;}

  // Do the day and date information starts at column 18; row 5 for date number; row 6 for day(3char)
  // run for 3 weeks, include extra work days and eliminate any holidays
  try
    with Data_Module.Inv_StoredProc do
    begin
      Close;
      ProcedureName := 'dbo.SELECT_OvertimeHolidayDate;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@BeginDate';
      Parameters.ParamValues['@BeginDate'] := formatdatetime('yyyy-mm-dd 00:00:00',now);
      Parameters.AddParameter.Name := '@EndDate';
      Parameters.ParamValues['@EndDate'] := formatdatetime('yyyy-mm-dd 00:00:00',now+21);
      Parameters.AddParameter.Name := '@Line';
      Parameters.ParamValues['@Line'] := fLine;
      Open;
      i:=19;
      if recordcount > 0 then
      begin
        //check for holiday/overtime
        for x:=0 to 20 do
        begin
          if not eof then
          begin
            if fieldbyname('DT_DATE').AsDateTime =  trunc(now+x) then
            begin
              if fieldbyname('VC_OVERTIME_HOLIDAY').AsString = 'O' then
              begin
                mysheet.Cells[5,i].value := formatdatetime('dd',now+x);
                mysheet.Cells[6,i].value := formatdatetime('ddd',now+x);
                fDates[x]:=now+x;
                INC(i);
              end;
              next;
            end
            else
            begin
              if DayOfTheWeek(now+x) < 6 then
              begin
                mysheet.Cells[5,i].value := formatdatetime('dd',now+x);
                mysheet.Cells[6,i].value := formatdatetime('ddd',now+x);
                fDates[x]:=now+x;
                INC(i);
              end;
            end;
          end
          else
          begin
            if DayOfTheWeek(now+x) < 6 then
            begin
              mysheet.Cells[5,i].value := formatdatetime('dd',now+x);
              mysheet.Cells[6,i].value := formatdatetime('ddd',now+x);
              fDates[x]:=now+x;
              INC(i);
            end;
          end;
        end;
      end
      else
      begin
        //normal run
        for x:=0 to 20 do
        begin
          if DayOfTheWeek(now+x) < 6 then
          begin
            mysheet.Cells[5,i].value := formatdatetime('dd',now+x);
            mysheet.Cells[6,i].value := formatdatetime('ddd',now+x);
            fDates[x]:=now+x;
            INC(i);
          end;
        end;
      end;
    end;
  except
    on e:exception do
    begin
      Data_Module.LogActLog('ERROR','Unable to get holiday/overtime infomation in order, '+e.Message);
      ShowMessage('Unable to get holiday/overtime infomation in order, '+e.Message);
      excel.Workbooks.Close;
      excel.Quit;
      excel:=Unassigned;
      mysheet:=Unassigned;
      exit;
    end;
  end;

  //Export Data
  try
    with Data_Module.Inv_StoredProc do
    begin
      //First Fill the Size,Supplier and Partnumber columns
      Close;
      ProcedureName := 'dbo.SELECT_PartsStockInfoOrder;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@AssyLine';
      Parameters.ParamValues['@AssyLine'] := fLine;
      Parameters.AddParameter.Name := '@TireWheel';
      if TireWheel_RadioGroup.ItemIndex = 0 then
        Parameters.ParamValues['@TireWheel'] := 'T'
      else
        Parameters.ParamValues['@TireWheel'] := 'W';
      Open;

      if recordcount > 0 then
      begin
        i:=7;
        lastsiz:='';
        lastsup:='';
        topedge:=i;
        while not eof do
        begin
          if lastsiz <> FieldByName('VC_SIZE_CODE').ASString then
          begin
            if firstheader then
            begin
              firstheader:=FALSE;
            end
            else
            begin
              mysheet.Cells[i,18].value  := 'Usage';
              // Get Forecast information and fill
              FillForecast(i);
              INC(i);
              mysheet.Cells[i,18].value  := 'End Balance';
              bottomedge:=i;
              DoBoarders(topedge,bottomedge,'B','B');
              DoBoarders(topedge,bottomedge,'C','C');
              DoBoarders(topedge,bottomedge,'D','D');
              DoFormulas(topedge,bottomedge);
              ClearForecast;
              INC(i);
              topedge:=i;
            end;
            lastsiz:=FieldByName('VC_SIZE_CODE').ASString;
            mysheet.Cells[i,2].value  := FieldByName('VC_SIZE_CODE').ASString;
            mysheet.Cells[i,18].value  := 'Beg Balance';
            //mysheet.Range['S'+IntToStr(i)].Formula:='=K'+IntToStr(i)+'+L'+IntToStr(i);
            UpdateSizeInfo(FieldByName('VC_SIZE_CODE').ASString,fLine,i);
            INC(i);
            fst:=TRUE;
          end;
          if lastsup <> FieldByName('VC_SUPPLIER_CODE').Value then
          begin
            with data_module.Inv_DataSet do
            begin
              Close;
              CommandType := CmdStoredProc;
              CommandText := 'dbo.SELECT_SupplierInfo;1';
              Parameters.Clear;
              Parameters.AddParameter.Name := '@SupCode';
              Parameters.ParamValues['@SupCode'] := Data_Module.Inv_StoredProc.fieldbyname('VC_SUPPLIER_CODE').ASString;
              Open;
            end;
            lastsup:=FieldByName('VC_SUPPLIER_CODE').Value;
          end;
          mysheet.Cells[i,3].value  := data_module.Inv_DataSet.FieldByName('Supplier Name').Value;
          mysheet.Cells[i,4].value  := FieldByName('VC_PART_NUMBER').Value;
          mysheet.Cells[i,15].value  := FieldByName('IN_1LOTQTY').Value;
          mysheet.Cells[i,16].value  := FieldByName('IN_LEADTIME').Value;
          mysheet.Cells[i,XLColToInt('BA')].value  := FieldByName('IN_QTY').Value;

          //mysheet.Range['K'+IntToStr(i)].Formula  := '=BA'+IntToStr(i)+'-(L'+IntToStr(i)+'+M'+IntToStr(i);

          // Get Nummi count

          // Get In Transit

          mysheet.Range['P'+IntToStr(i)].Interior.ColorIndex:=36;
          // set the color for the lead time range
          // final one matches the order area
          // add formula
          DoLeadTime(FieldByName('IN_LEADTIME').AsInteger,i);
          // Do mini header
          if fst then
          begin
            mysheet.Cells[i,8].value  := 'Forecast';
            mysheet.Cells[i,9].value  := 'Ship';
            mysheet.Cells[i,10].value  := 'Total';
            fst:=FALSE;
          end;
          mysheet.Cells[i,18].value  := 'Receipts';
          mysheet.Range['Q'+IntToStr(i)].Interior.ColorIndex:=40;
          fpartline[fpartcount,1]:=FieldByName('VC_PART_NUMBER').Value;
          fpartline[fpartcount,2]:=IntToStr(i);
          UpdateForecast(FieldByName('VC_PART_NUMBER').Value);
          INC(i);
          next;
        end;
        mysheet.Cells[i,18].value  := 'Usage';
        INC(i);
        mysheet.Cells[i,18].value  := 'End Balance';
        bottomedge:=i;
        DoBoarders(topedge,bottomedge,'B','B');
        DoBoarders(topedge,bottomedge,'C','C');
        DoBoarders(topedge,bottomedge,'D','D');
        DoBoarders(7,bottomedge,'E','AJ',TRUE);
        INC(i);
      end;
    end;
  except
    on e:exception do
    begin
      Data_Module.LogActLog('ERROR','Unable to export infomation in order, '+e.Message);
      ShowMessage('Unable to export infomation in order, '+e.Message);
      excel.Workbooks.Close;
      excel.Quit;
      excel:=Unassigned;
      mysheet:=Unassigned;
      exit;
    end;
  end;

  Start_Button.Enabled:=False;
end;

procedure TOrder_Form.Close_ButtonClick(Sender: TObject);
begin
  if Start_Button.Enabled = False then
  begin
    // get orders and process
    excel.Workbooks.Close;
    excel.Quit;
    excel:=Unassigned;
    mysheet:=Unassigned;
  end;
end;

procedure TOrder_Form.DoBoarders(topedge,bottomedge:integer;startcolumn,endcolumn:string;fill:boolean=FALSE);
begin
  mysheet.Range[startcolumn+inttostr(topedge), endcolumn+inttostr(bottomedge)].Borders.Item[xlEdgeTop].Linestyle
                    :=  xlContinuous;
  mysheet.Range[startcolumn+inttostr(topedge), endcolumn+inttostr(bottomedge)].Borders.Item[xlEdgeBottom].Linestyle
                    :=  xlContinuous;
  mysheet.Range[startcolumn+inttostr(topedge), endcolumn+inttostr(bottomedge)].Borders.Item[xlEdgeRight].Linestyle
                    :=  xlContinuous;
  mysheet.Range[startcolumn+inttostr(topedge), endcolumn+inttostr(bottomedge)].Borders.Item[xlEdgeLeft].Linestyle
                    :=  xlContinuous;
  if fill then
  begin
    mysheet.Range[startcolumn+inttostr(topedge), endcolumn+inttostr(bottomedge)].Borders.Item[xlInsideVertical].Linestyle
                    :=  xlContinuous;
    mysheet.Range[startcolumn+inttostr(topedge), endcolumn+inttostr(bottomedge)].Borders.Item[xlInsideHorizontal].Linestyle
                    :=  xlContinuous;
  end;
end;

procedure TOrder_Form.UpdateSizeInfo(size,assyline:string;line:integer);
begin
  try
    // Fill forecast information
    with data_module.INV_Forecast_DataSet do
    begin
      Close;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.SELECT_SizeInfo;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@SizeCode';
      Parameters.ParamValues['@SizeCode'] := size;
      Parameters.AddParameter.Name := '@AssyCode';
      Parameters.ParamValues['@AssyCode'] := '';
      Parameters.AddParameter.Name := '@AssyLine';
      Parameters.ParamValues['@AssyLine'] := assyline;
      Open;
      if recordcount > 0 then
      begin
        mysheet.Cells[line,8].value  := FieldByName('IN_USAGE').Value;
        mysheet.Cells[line,9].value  := FieldByName('IN_DAYS').Value;
        mysheet.Range['J'+IntToStr(line)].Formula:='=H'+IntToStr(line)+'*I'+IntToStr(line);
        mysheet.Range['H'+IntToStr(line),'I'+IntToStr(line)].Interior.ColorIndex:=34;
      end;
    end;
  except
    on e:exception do
    begin
      Data_Module.LogActLog('ERROR','Unable to get size info in order, '+e.Message);
      ShowMessage('Unable to get size info in order, '+e.Message);
      raise;
    end;
  end;
end;

procedure TOrder_Form.UpdateForecast(partno:string);
var
  j,z:integer;
begin
  try
    // Fill forecast information
    with data_module.INV_Forecast_DataSet do
    begin
      j:=0;
      for z:=0 to 20 do
      begin
        if fDates[z] <> 0 then
        begin
          Close;
          CommandType := CmdStoredProc;
          CommandText := 'dbo.SELECT_ForecastPartNumberWeek;1';
          Parameters.Clear;
          Parameters.AddParameter.Name := '@WeekNo';
          Parameters.ParamValues['@WeekNo'] := WeekOfTheYear(fDates[z]);
          Parameters.AddParameter.Name := '@DayNo';
          Parameters.ParamValues['@DayNo'] := DayOfTheWeek(fDates[z]);
          Parameters.AddParameter.Name := '@PartNo';
          Parameters.ParamValues['@PartNo'] := Data_Module.Inv_StoredProc.FieldByName('VC_PART_NUMBER').Value;
          Open;
          fForecast[j]  := fForecast[j]+FieldByName('Qty').Value;
          INC(j);
        end;
      end;
    end;
  except
    on e:exception do
    begin
      Data_Module.LogActLog('ERROR','Unable to get forecast in order, '+e.Message);
      ShowMessage('Unable to get forecast in order, '+e.Message);
      raise;
    end;
  end;
end;

procedure TOrder_Form.FillForecast(line:integer);
var
  i:integer;
begin
  for i:=0 to 20 do
  begin
    if fForecast[i] <> 0 then
      mysheet.Cells[line,i+19].value := fForecast[i];
  end;
end;

procedure TOrder_Form.ClearForecast;
var
  i:integer;
begin
  for i:=0 to 20 do
  begin
    fForecast[i]:=0;
  end;
end;

procedure TOrder_Form.DoFormulas(topedge,bottomedge:integer);
var
  i:integer;
begin

  // Total Inventory Formulas
  // GSA
  mysheet.Range['K'+IntToStr(topedge)].Formula:='=SUM(K'+IntToStr(topedge+1)+':K'+IntToStr(bottomedge)+')';
  // NUMMI
  mysheet.Range['L'+IntToStr(topedge)].Formula:='=SUM(L'+IntToStr(topedge+1)+':L'+IntToStr(bottomedge)+')';
  // In Transit
  mysheet.Range['M'+IntToStr(topedge)].Formula:='=SUM(M'+IntToStr(topedge+1)+':M'+IntToStr(bottomedge)+')';
  // Open Order
  mysheet.Range['N'+IntToStr(topedge)].Formula:='=SUM(N'+IntToStr(topedge+1)+':N'+IntToStr(bottomedge)+')';
  // begining balance  and formula for forcast time


  mysheet.Range['S'+IntToStr(topedge)].Formula:='=SUM(BA'+IntToStr(topedge+1)+':BA'+IntToStr(bottomedge)+')';

  for i:=0 to 20 do
  begin
    if fForecast[i] <> 0 then
    begin
      if fForecast[i+1] <> 0 then
        mysheet.Range[IntToXLCol(20+i)+IntToStr(topedge)].Formula:='='+IntToXLCol(19+i)+IntToStr(bottomedge);
      if bottomedge-topedge = 3 then
      begin
        mysheet.Range[IntToXLCol(19+i)+IntToStr(bottomedge)].Formula:='='+IntToXLCol(19+i)+IntToStr(topedge)+'+'+IntToXLCol(19+i)+IntToStr(topedge+1)+'-'+IntToXLCol(19+i)+IntToStr(topedge+2);
      end
      else
      begin
        mysheet.Range[IntToXLCol(19+i)+IntToStr(bottomedge)].Formula:='='+IntToXLCol(19+i)+IntToStr(topedge)+'+'+IntToXLCol(19+i)+IntToStr(topedge+1)+'+'+IntToXLCol(19+i)+IntToStr(topedge+2)+'-'+IntToXLCol(19+i)+IntToStr(topedge+3);
      end;
    end;
  end;
end;

procedure TOrder_Form.DoLeadTime(Leadtime,Line:integer);
begin
  mysheet.Range['S'+IntToStr(line),IntToXLCol(XLColToInt('S')+(LeadTime-1))+IntToStr(line)].Interior.ColorIndex:=36;
  mysheet.Range[IntToXLCol(XLColToInt('S')+LeadTime)+IntToStr(line)].Interior.ColorIndex:=40;
  mysheet.Range[IntToXLCol(XLColToInt('S')+LeadTime)+IntToStr(line)].Formula:='=Q'+IntToStr(line);
end;

// Turns a column number into the corresponding column letter

function TOrder_Form.IntToXlCol(ColNum: TColNum): TColString;
const
  Letters: array[0..MaxLetters] of char
='ZABCDEFGHIJKLMNOPQRSTUVWXYZ';
var
  nQuot, nMod: integer;
Begin
  if ColNum <= MaxLetters then
  begin
    Result[0] := Chr(1);
    Result[1] := Letters[ColNum];
  end
  else
  begin
    nQuot := ((ColNum - 1) * 10083) shr 18;
    nMod := ColNum - (nQuot * MaxLetters);
    Result[0] := Chr(2);
    Result[1] := Letters[nQuot];
    Result[2] := Letters[nMod];
  end;
end;

// Turns a column identifier into the corresponding integer,
//  A=1, ..., AA = 27, ..., IV = 256

function TOrder_Form.XlColToInt(const Col: TColString): TColNum;
const
  ASCIIOffset = Ord('A') - 1;
var
  Len: cardinal;
begin
  Len := Length(Col);
  Result := Ord(UpCase(Col[Len])) - ASCIIOffset;
  if (Len > 1) then
    Inc(Result, (Ord(UpCase(Col[1])) - ASCIIOffset) *
MaxLetters);
end;

end.
