//***********************************************************
//  InventorySystem
//
//  Copyright (c) 2002-2003 Failproof Manufacturing Systems
//
//***********************************************************
//
//  10/25/2002  Aaron Huge      Initial creation
//  12/30/2002  David Verespey  Add the guts

unit Order;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, ComCtrls, OleServer, Excel2000, ComObj, ADOdb, DateUtils,
  NummiTime;

type
  TColNum = 1..256;  // Excel columns only go up to IV
  TColString = string[2];

  {TPartLine = class(TObject)
    PartNumber:string;
    LineNumber:integer;
    SupplierCode:string;
    KanbanNumber:string;
    RenbanGroup:string;
    FRSDate:string;
    LotSizeOrders:string;
  end;}

  TOrder_Form = class(TForm)
    Start_Button: TButton;
    Cancel_Button: TButton;
    Order_Label: TLabel;
    Label1: TLabel;
    TireWheel_RadioGroup: TRadioGroup;
    Date_DateTimePicker: TDateTimePicker;
    Line_RadioGroup: TRadioGroup;
    ProcessOrder_Button: TButton;
    procedure FormCreate(Sender: TObject);
    procedure Start_ButtonClick(Sender: TObject);
    procedure Cancel_ButtonClick(Sender: TObject);
    procedure ProcessOrder_ButtonClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
    //fPartLineList : array [1..200] of TPartLine;
    fpartline : array[1..200,1..7] of string;
    fpartcount:integer;
    fDates:array[0..60] of Double;
    fForecast:array[0..60] of integer;
    fOvertimes:array[0..60] of integer;
    fLine:string;
    fTireWheel:string;
    fFillDays:integer;
    fOvertimeCount:integer;
    fSizeCount:integer;
    fdaterange:integer;
    procedure DoBoarders(topedge,bottomedge:integer;startcolumn,endcolumn:string;fill:boolean=FALSE);
    procedure UpdateForecast(partno:string);
    procedure FillForecast(line:integer);
    procedure ClearForecast;
    procedure DoFormulas(topedge,bottomedge:integer);
    function XlColToInt(const Col: TColString): TColNum;
    function IntToXlCol(ColNum: TColNum): TColString;
    procedure UpdateSizeInfo(size,assyline:string;line:integer);
    procedure DoLeadTime(Leadtime,Line:integer);
    procedure PutWQSCount(partnumber:string;line:integer);
    procedure PutNUMMICount(partnumber:string;line:integer);
    procedure PutIntransitCount(partnumber:string;line:integer);
    procedure PutOpenOrderCount(partnumber:string;leadtime,line:integer);
    procedure UpdateFRSInfo(partnumber:string;line:integer);
    procedure MonthForecast(partno:string; line:integer);
    procedure OrderHistory(partnumber:string;line:integer);
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
  // definitions for excel form placement
  MaxLetters      = 26;
  DateWeekCol     = 19;     // These two need to
  DateWeekLetter  = 'T';    // change in concert
  ForecastCol     = 20;
  FirstCol        = 2;
  QtyCol          = 17;
  LotCol          = 18;
  DateRangeCount  = 59;
  FillDays        = 20;

  // Definitions for Partline array
  PartNumber      = 1;
  LineNumber      = 2;
  SupplierCode    = 3;
  KanbanNumber    = 4;
  RenbanGroup     = 5;
  FRSDate         = 6;
  LotSizeOrders   = 7;

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
  Cancel_Button.Caption:='E&xit';
end;

procedure TOrder_Form.Start_ButtonClick(Sender: TObject);
var
  i,x,topedge,bottomedge,pcount,ots,leadtime:integer;
  Year, Month, Day: Word;
  lastsup,lastsiz:string;
  firstheader,fst:boolean;
begin
  Cancel_Button.Caption:='&Cancel';

  Data_Module.LogActLog('ORDER','Start Excel Order');
  Start_Button.Enabled:=False;
  ProcessOrder_Button.Enabled:=True;
  fpartcount:=0;
  firstheader:=TRUE;
  fst:=TRUE;
  excel := createOleObject('Excel.Application');
  excel.visible := True;
  excel.DisplayAlerts:=False;
  fdaterange:=DateRangeCount; //inital starting point

  if Line_RadioGroup.ItemIndex = 0 then
    fLine := 'C'
  else
    fLine := 'T';

  if TireWheel_RadioGroup.ItemIndex = 0 then
    fTireWheel := 'T'
  else
    fTireWheel := 'W';

  excel.workbooks.open(ExtractFilePath(application.ExeName)+'OrderSimulation.xls');

  mysheet := excel.workSheets[1];

  //Color debug do not delete
  {for i:=1 to 56 do
  begin
    mysheet.Range[IntToXLCol(i)+'1'].Interior.ColorIndex:=i;
  end;}
  //end debug

  // Do the day and date information starts at column 18; row 5 for date number; row 6 for day(3char)
  // run for 3 weeks, include extra work days and eliminate any holidays

  mysheet.Cells[6,XLColToInt('AQ')].value  := 'Total Inv';
  mysheet.Cells[6,XLColToInt('AR')].value  := 'In Transit';
  mysheet.Cells[6,XLColToInt('AS')].value  := 'Added Leadtime';


  mysheet.Range['T5','AR200'].Locked := True;


  try
    with Data_Module.Inv_StoredProc do
    begin
      //
      // Account for Christmas Shutdown
      //
      {DecodeDate(now+fdaterange, Year, Month, Day);
      if (month = 12) and (day >= 20) then
      begin
        // Get first production day
        Close;
        ProcedureName := 'dbo.SELECT_FirstProductionDay;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@ProdYear';
        Parameters.ParamValues['@ProdYear'] := formatdatetime('yyyy',now+60);
        Open;

        fdaterange := trunc(fieldbyname('DT_FIRST_PRODUCTION_DAY_OF_YEAR').asDateTime - now)+1;
      end;}

      Close;
      ProcedureName := 'dbo.SELECT_OvertimeHolidayDate;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@BeginDate';
      Parameters.ParamValues['@BeginDate'] := formatdatetime('yyyy-mm-dd 00:00:00',now);
      Parameters.AddParameter.Name := '@EndDate';
      Parameters.ParamValues['@EndDate'] := formatdatetime('yyyy-mm-dd 00:00:00',now+fdaterange);
      Parameters.AddParameter.Name := '@Line';

      Parameters.ParamValues['@Line'] := fLine;
      Open;

      i:=DateWeekCol+1;

      mysheet.Cells[5,DateWeekCol].value := 'Date';
      mysheet.Cells[6,DateWeekCol].value := 'Week';

      fFillDays:=0;
      for x:=0 to fdaterange do
      begin
        fDates[x]:=0;
        fOvertimes[x]:=0;
        fForecast[x]:=0;
      end;
      ots:=0;
      fOvertimecount:=0;

      //fOvertimeCount:=recordcount;
      if recordcount > 0 then
      begin
        //check for holiday/overtime
        //for x:=0 to fdaterange do
        x:=0;
        while fFillDays < FillDays do
        begin
          if not eof then
          begin
            if fieldbyname('DT_DATE').AsDateTime =  trunc(now+x) then
            begin
              if fieldbyname('VC_OVERTIME_HOLIDAY').AsString = 'O' then
              begin
                mysheet.Cells[5,i].value := formatdatetime('dd-mmm-yy',now+x);
                mysheet.Cells[6,i].value := formatdatetime('ddd',now+x);
                fDates[x]:=now+x;
                INC(i);
                INC(fFillDays);
                fOvertimes[ots]:=fFilldays;
                INC(ots);
                INC(fOvertimeCount);
              end;
              next;
            end
            else
            begin
              if DayOfTheWeek(now+x) < 6 then
              begin
                mysheet.Cells[5,i].value := formatdatetime('dd-mmm-yy',now+x);
                mysheet.Cells[6,i].value := formatdatetime('ddd',now+x);
                fDates[x]:=now+x;
                INC(i);
                INC(fFillDays);
              end;
            end;
          end
          else
          begin
            if DayOfTheWeek(now+x) < 6 then
            begin
              mysheet.Cells[5,i].value := formatdatetime('dd-mmm-yy',now+x);
              mysheet.Cells[6,i].value := formatdatetime('ddd',now+x);
              fDates[x]:=now+x;
              INC(i);
              INC(fFillDays);
            end;
          end;
          //
          // change to increment x and not stop until 20 total days are diplayed
          //
          INC(x);

        end;
      end
      else
      begin
        //normal run
        for x:=0 to fdaterange do
        begin
          if DayOfTheWeek(now+x) < 6 then
          begin
            mysheet.Cells[5,i].value := formatdatetime('dd-mmm-yy',now+x);
            mysheet.Cells[6,i].value := formatdatetime('ddd',now+x);
            fDates[x]:=now+x;
            INC(i);
            INC(fFillDays);
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
      Parameters.ParamValues['@TireWheel'] := fTireWheel;
      Open;

      if recordcount > 0 then
      begin
        //setlength(fPartLine,recordcount,Arraywidth);
        i:=7;
        lastsiz:='';
        lastsup:='';
        topedge:=i;
        pcount:=1;
        while not eof do
        begin
          if lastsiz <> FieldByName('VC_SIZE_CODE').ASString then
          begin
            pcount:=1;
            if firstheader then
            begin
              firstheader:=FALSE;
            end
            else
            begin
              mysheet.Cells[i,DateWeekCol].value  := 'Usage';
              // Get Forecast information and fill
              FillForecast(i);
              INC(i);
              mysheet.Cells[i,DateWeekCol].value  := 'End Balance';
              bottomedge:=i;
              DoBoarders(topedge,bottomedge,'B','B');
              DoBoarders(topedge,bottomedge,'C','C');
              DoBoarders(topedge,bottomedge,'D','D');
              DoFormulas(topedge,bottomedge);
              ClearForecast;
              INC(i);
              mysheet.Range['E'+IntToStr(i), 'AP'+IntToStr(i)].Interior.ColorIndex:=XLColToInt('AV');
              INC(i);
              topedge:=i;
            end;
            lastsiz:=FieldByName('VC_SIZE_CODE').ASString;
            mysheet.Cells[i,FirstCol].value  := FieldByName('VC_SIZE_CODE').ASString;
            mysheet.Cells[i,DateWeekCol].value  := 'Beg Balance';
            UpdateSizeInfo(FieldByName('VC_SIZE_CODE').ASString,fLine,i);
            INC(i);
            fst:=TRUE;
            fSizeCount:=1;
          end;
          if lastsup <> FieldByName('VC_SUPPLIER_CODE').Value then
          begin
            with data_module.Inv_DataSet do
            begin
              Close;
              Filter:='';
              Filtered:=FALSE;
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

          case DayOfTheWeek(now) of
            1:
              if FieldByName('IN_LEADTIME_MONDAY').AsInteger <> 0 then
                leadtime := FieldByName('IN_LEADTIME_MONDAY').AsInteger
              else
                leadtime := FieldByName('IN_LEADTIME').AsInteger;
            2:
              if FieldByName('IN_LEADTIME_TUESDAY').AsInteger <> 0 then
                leadtime := FieldByName('IN_LEADTIME_TUESDAY').AsInteger
              else
                leadtime := FieldByName('IN_LEADTIME').AsInteger;
            3:
              if FieldByName('IN_LEADTIME_WEDNESDAY').AsInteger <> 0 then
                leadtime := FieldByName('IN_LEADTIME_WEDNESDAY').AsInteger
              else
                leadtime := FieldByName('IN_LEADTIME').AsInteger;
            4:
              if FieldByName('IN_LEADTIME_THURSDAY').AsInteger <> 0 then
                leadtime := FieldByName('IN_LEADTIME_THURSDAY').AsInteger
              else
                leadtime := FieldByName('IN_LEADTIME').AsInteger;
            5:
              if FieldByName('IN_LEADTIME_FRIDAY').AsInteger <> 0 then
                leadtime := FieldByName('IN_LEADTIME_FRIDAY').AsInteger
              else
                leadtime := FieldByName('IN_LEADTIME').AsInteger;
            6:
              if FieldByName('IN_LEADTIME_SATURDAY').AsInteger <> 0 then
                leadtime := FieldByName('IN_LEADTIME_SATURDAY').AsInteger
              else
                leadtime := FieldByName('IN_LEADTIME').AsInteger;
            else
              leadtime := FieldByName('IN_LEADTIME').AsInteger;
          end;

          mysheet.Cells[i,16].value  := leadtime;
          mysheet.Cells[i,XLColToInt('AQ')].value  := FieldByName('IN_QTY').Value;

          // Do order percentage
          mysheet.Cells[i,5].value  := 0; //Qty
          OrderHistory(FieldByName('VC_PART_NUMBER').Value,i);
          if pcount=1 then
            mysheet.Cells[i,6].value  := 1 //% actual percentage is 100 if one part
          else
          begin
            //calculate order percentage if two parts
            mysheet.Range['F'+IntToStr(i-1)].Formula := '=E'+IntToStr(i-1)+'/(E'+IntToStr(i)+'+E'+IntToStr(i-1)+')';
            mysheet.Range['F'+IntToStr(i)].Formula := '=E'+IntToStr(i)+'/(E'+IntToStr(i)+'+E'+IntToStr(i-1)+')';
          end;

          // Do FRS
          UpdateFRSInfo(FieldByName('VC_PART_NUMBER').Value,i);

          // Monthly forecast vs Usage
          MonthForecast(FieldByName('VC_PART_NUMBER').Value,i);

          // Get WQS count
          mysheet.Range['K'+IntToStr(i)].Formula  := '=AQ'+IntToStr(i)+'-(L'+IntToStr(i)+'+M'+IntToStr(i)+')';
          PutWQSCount(FieldByName('VC_PART_NUMBER').AsString,i);

          // Get Nummi count
          PutNUMMICount(FieldByName('VC_PART_NUMBER').AsString,i);

          // Get In Transit
          PutIntransitCount(FieldByName('VC_PART_NUMBER').AsString,i);

          // Get Open Orders
          PutOpenOrderCount(FieldByName('VC_PART_NUMBER').AsString,leadtime,i);

          //Lead time box
          mysheet.Range['P'+IntToStr(i)].Interior.ColorIndex:=36;

          // set the color for the lead time range
          // final one matches the order area
          // add formula
          DoLeadTime(leadtime,i);

          // Do mini header (forecast/ship/total)
          if fst then
          begin
            mysheet.Cells[i,8].value  := 'Forecast';
            mysheet.Cells[i,9].value  := 'Ship';
            mysheet.Cells[i,10].value  := 'Total';
            fst:=FALSE;
          end;
          mysheet.Cells[i,DateWeekCol].value  := 'Receipts';
          mysheet.Range['Q'+IntToStr(i)].Interior.ColorIndex:=40;
          mysheet.Range['R'+IntToStr(i)].Interior.ColorIndex:=40;
          // Lot size formula
          mysheet.Range['Q'+IntToStr(i)].Formula  := '=R'+IntToStr(i)+'*O'+IntToStr(i);
          fpartline[i,PartNumber]:=FieldByName('VC_PART_NUMBER').Value;
          fpartline[i,LineNumber]:=IntToStr(i);
          fpartline[i,SupplierCode]:=fieldbyname('VC_SUPPLIER_CODE').ASString;
          fpartline[i,KanbanNumber]:=fieldbyname('VC_KANBAN_NUMBER').ASString;
          fpartline[i,RenbanGroup]:=fieldbyname('VC_RENBAN_GROUP_CODE').ASString;
          if fieldbyname('BIT_LOT_SIZE_ORDERS').AsBoolean then
            fpartline[i,LotSizeOrders]:='TRUE'
          else
            fpartline[i,LotSizeOrders]:='FALSE';

          INC(fPartCount);
          UpdateForecast(FieldByName('VC_PART_NUMBER').Value);
          INC(i);
          INC(pcount);
          next;
          INC(fSizeCount);
        end;
        mysheet.Cells[i,DateWeekCol].value  := 'Usage';
        // Get Forecast information and fill
        FillForecast(i);
        INC(i);
        mysheet.Cells[i,DateWeekCol].value  := 'End Balance';
        bottomedge:=i;
        DoBoarders(topedge,bottomedge,'B','B');
        DoBoarders(topedge,bottomedge,'C','C');
        DoBoarders(topedge,bottomedge,'D','D');
        DoFormulas(topedge,bottomedge);
        DoBoarders(7,bottomedge,'E','AP',TRUE);
        ClearForecast;
        INC(i);
      end;
    end;
    //mysheet.protect('OrderSimulation.xls');
  except
    on e:exception do
    begin
      Data_Module.LogActLog('ERROR','Unable to export infomation in order, '+e.Message);
      ShowMessage('Unable to export infomation in order, '+e.Message);
      excel.Workbooks.Close;
      excel.Quit;
      excel:=Unassigned;
      mysheet:=Unassigned;
      ProcessOrder_Button.Enabled:=False;
      Start_Button.Enabled:=True;
    end;
  end;
end;

procedure TOrder_Form.Cancel_ButtonClick(Sender: TObject);
begin
  try
    if Start_Button.Enabled = False then
    begin
      if MessageDlg('Cancel current worksheet, all changes will be lost?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
        // get orders and process
        excel.Workbooks.Close;
        excel.Quit;
        excel:=Unassigned;
        mysheet:=Unassigned;

        Cancel_Button.Caption:='E&xit';
        Data_Module.LogActLog('ORDER','Exit Order without save');
        Start_Button.Enabled:=true;
        ProcessOrder_Button.Enabled:=False;
      end;
    end
    else
    begin
      Close;
      Data_Module.LogActLog('ORDER','Exit Order');
    end;
  except
    on e:exception do
    begin
      ShowMessage('Unable to exit, check Excel form for active edit and retry, '+e.Message);
      Data_Module.LogActLog('ERROR','Unable to exit, check Excel form for active edit and retry, '+e.Message);
    end;
  end;
end;

procedure TOrder_Form.ProcessOrder_ButtonClick(Sender: TObject);
var
  i,j,rcount:integer;
  qtycount,lotcount:string;
begin
  try
    if MessageDlg('Create order files from worksheet?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      //Loop through order column and use the partnumber array to add an order record
      //for the part number
      //
      //
      //
      //  Add FRS code: The FRS number is YYMMDDXX where xx is a count of the number of trailers based on lot size
      //                YYMMDD= the expect date of arrival, order date plus lead time.
      //
      //  Add Renban logic: The renban numbeR is KKKKKXXX where KKKKK is the kanban number and XXX is an incrementing count
      //                    Except in the case of renban goups where no renban number is generated here, that is done on the
      //                    renban/order grouping screen.
      //

      for i:=1 to 200 do
      begin
        if fPartLine[i,PartNumber] <> '' then
        begin
          //if valid line
          qtycount:=mysheet.Cells[i,QtyCol].value;
          lotcount:=mysheet.Cells[i,LotCol].value;
          if (qtycount <> '') and (lotcount <> '') then
          begin
            // if a number entered then add order record
            try
              StrToInt(qtycount);
            except
              on e:exception do
              begin
                ShowMessage('Invalid data in order qty column Line Number('+fpartline[i,LineNumber]+'), must be numeric');
                exit;
              end;
            end;

            try
              StrToInt(lotcount);
            except
              on e:exception do
              begin
                ShowMessage('Invalid data in order lot column Line Number('+fpartline[i,LineNumber]+'), must be numeric');
                exit;
              end;
            end;

            if StrToInt(qtycount) <> 0 then // no zero count orders
            begin
              // Check for lot size orders, if true one order per qty
              // else FRS one per lot
              if fPartLine[i,LotSizeOrders] = 'TRUE' then
              begin
                try
                  Data_Module.Inv_Connection.BeginTrans;
                  with data_module.INV_Order_StoredProc do
                  begin
                    Close;
                    ProcedureName := 'dbo.INSERT_OpenOrder;1';
                    Parameters.Clear;
                    Parameters.AddParameter.Name := '@SupCode';
                    Parameters.ParamValues['@SupCode'] := fPartLine[i,SupplierCode];
                    Parameters.AddParameter.Name := '@PartNum';
                    Parameters.ParamValues['@PartNum'] := fPartLine[i,PartNumber];
                    Parameters.AddParameter.Name := '@KanbanNum';
                    Parameters.ParamValues['@KanbanNum'] := fPartLine[i,KanbanNumber];

                    //
                    Parameters.AddParameter.Name := '@FRSNum';
                    Parameters.ParamValues['@FRSNum'] := copy(formatdatetime('yymmdd',StrToDate(fPartLine[i,FRSDate])),2,5)+'01';

                    // Renban = Kanban + incremented value
                    // Unless it is a Renban Group then none assigned here
                    //
                    Parameters.AddParameter.Name := '@RenbanNum';
                    if fPartLine[i,RenbanGroup] = '' then
                    begin

                      // Get current count
                      with Data_Module.Inv_StoredProc do
                      begin
                        Close;
                        ProcedureName := 'dbo.SELECT_PartsStockRenban;1';
                        Parameters.Clear;
                        Parameters.AddParameter.Name := '@PartCode';
                        Parameters.ParamValues['@PartCode'] := fPartLine[i,PartNumber];
                        Open;
                      end;

                      //set renban
                      Parameters.ParamValues['@RenbanNum'] := fPartLine[i,KanbanNumber]+format('%.3d',[Data_Module.Inv_StoredProc.fieldbyname('IN_RENBAN_COUNT').asInteger]);

                      rcount:=Data_Module.Inv_StoredProc.fieldbyname('IN_RENBAN_COUNT').AsInteger;

                      // new renban
                      INC(rcount);
                      if rcount > 999 then
                        rcount:=1;

                      //update renban
                      with Data_Module.Inv_StoredProc do
                      begin
                        Close;
                        ProcedureName := 'dbo.UPDATE_PartsStockRenban;1';
                        Parameters.Clear;
                        Parameters.AddParameter.Name := '@PartCode';
                        Parameters.ParamValues['@PartCode'] := fPartLine[i,PartNumber];
                        Parameters.AddParameter.Name := '@RenbanCount';
                        Parameters.ParamValues['@RenbanCount'] := rcount;
                        ExecProc;
                      end;
                    end
                    else
                    begin
                      Parameters.ParamValues['@RenbanNum'] := '';
                    end;

                    Parameters.AddParameter.Name := '@Qty';
                    Parameters.ParamValues['@Qty'] := qtycount; //lot size
                    ExecProc;
                    Data_Module.LogActLog('ORDER','Create order, pn='+fPartLine[i,PartNumber]+',kanban='+fPartLine[i,KanbanNumber]+', qty='+qtycount);
                  end;
                  Data_Module.Inv_Connection.CommitTrans;
                except
                  on e:exception do
                  begin
                    Data_Module.LogActLog('ERROR','Unable to create order records, '+e.Message);
                    ShowMessage('Unable to create order records, '+e.Message);
                    excel.Workbooks.Close;
                    excel.Quit;
                    excel:=Unassigned;
                    mysheet:=Unassigned;
                    ProcessOrder_Button.Enabled:=False;
                    Start_Button.Enabled:=True;
                    if Data_Module.Inv_Connection.InTransaction then
                      Data_Module.Inv_Connection.RollbackTrans;
                  end;
                end;
              end
              else
              begin
              // Break down into FRS lot counts (1 lot per FRS number 00-99)
                for j:=1 to StrToInt(lotcount) do
                begin
                  try
                    Data_Module.Inv_Connection.BeginTrans;
                    with data_module.INV_Order_StoredProc do
                    begin
                      Close;
                      ProcedureName := 'dbo.INSERT_OpenOrder;1';
                      Parameters.Clear;
                      Parameters.AddParameter.Name := '@SupCode';
                      Parameters.ParamValues['@SupCode'] := fPartLine[i,SupplierCode];
                      Parameters.AddParameter.Name := '@PartNum';
                      Parameters.ParamValues['@PartNum'] := fPartLine[i,PartNumber];
                      Parameters.AddParameter.Name := '@KanbanNum';
                      Parameters.ParamValues['@KanbanNum'] := fPartLine[i,KanbanNumber];

                      //
                      Parameters.AddParameter.Name := '@FRSNum';
                      if j<10 then
                        Parameters.ParamValues['@FRSNum'] := copy(formatdatetime('yymmdd',StrToDate(fPartLine[i,FRSDate])),2,5)+'0'+IntToStr(j)
                      else
                        Parameters.ParamValues['@FRSNum'] := copy(formatdatetime('yymmdd',StrToDate(fPartLine[i,FRSDate])),2,5)+IntToStr(j);

                      // Renban = Kanban + incremented value
                      // Unless it is a Renban Group then none assigned here
                      //
                      Parameters.AddParameter.Name := '@RenbanNum';
                      if fPartLine[i,RenbanGroup] = '' then
                      begin

                        // Get current count
                        with Data_Module.Inv_StoredProc do
                        begin
                          Close;
                          ProcedureName := 'dbo.SELECT_PartsStockRenban;1';
                          Parameters.Clear;
                          Parameters.AddParameter.Name := '@PartCode';
                          Parameters.ParamValues['@PartCode'] := fPartLine[i,PartNumber];
                          Open;
                        end;

                        //set renban
                        Parameters.ParamValues['@RenbanNum'] := fPartLine[i,KanbanNumber]+format('%.3d',[Data_Module.Inv_StoredProc.fieldbyname('IN_RENBAN_COUNT').asInteger]);

                        rcount:=Data_Module.Inv_StoredProc.fieldbyname('IN_RENBAN_COUNT').AsInteger;

                        // new renban
                        INC(rcount);
                        if rcount > 999 then
                          rcount:=1;

                        //update renban
                        with Data_Module.Inv_StoredProc do
                        begin
                          Close;
                          ProcedureName := 'dbo.UPDATE_PartsStockRenban;1';
                          Parameters.Clear;
                          Parameters.AddParameter.Name := '@PartCode';
                          Parameters.ParamValues['@PartCode'] := fPartLine[i,PartNumber];
                          Parameters.AddParameter.Name := '@RenbanCount';
                          Parameters.ParamValues['@RenbanCount'] := rcount;
                          ExecProc;
                        end;
                      end
                      else
                      begin
                        Parameters.ParamValues['@RenbanNum'] := '';
                      end;

                      Parameters.AddParameter.Name := '@Qty';
                      Parameters.ParamValues['@Qty'] := mysheet.Cells[i,15].value; //lot size
                      qtycount:=mysheet.Cells[i,15].value;
                      ExecProc;
                      Data_Module.LogActLog('ORDER','Create order, pn='+fPartLine[i,PartNumber]+',kanban='+fPartLine[i,KanbanNumber]+',Qty='+qtycount);
                    end;
                    Data_Module.Inv_Connection.CommitTrans;
                  except
                    on e:exception do
                    begin
                      Data_Module.LogActLog('ERROR','Unable to create order records, '+e.Message);
                      ShowMessage('Unable to create order records, '+e.Message);
                      excel.Workbooks.Close;
                      excel.Quit;
                      excel:=Unassigned;
                      mysheet:=Unassigned;
                      ProcessOrder_Button.Enabled:=False;
                      Start_Button.Enabled:=True;
                      if Data_Module.Inv_Connection.InTransaction then
                        Data_Module.Inv_Connection.RollbackTrans;
                    end;
                  end;
                end;
              end;
            end;
          end;
        end;
      end;

      // Close workbook
      excel.Workbooks.Close;
      excel.Quit;
      excel:=Unassigned;
      mysheet:=Unassigned;
      ProcessOrder_Button.Enabled:=False;
      Start_Button.Enabled:=True;
      Cancel_Button.Caption:='E&xit';
    end;
  except
    on e:exception do
    begin
      Data_Module.LogActLog('ERROR','Unable to save orders from Excel sheet, '+e.Message);
      ShowMessage('Unable to save orders, check for active edit on worksheet and retry, '+e.Message);
    end;
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
    // Get usage information
    with data_module.INV_Forecast_DataSet do
    begin
      Close;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.SELECT_SizeInfo;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@SizeCode';
      Parameters.ParamValues['@SizeCode'] := size;
      Open;
      if recordcount > 0 then
      begin
        mysheet.Cells[line,8].value  := FieldByName('Daily Usage').Value;
        mysheet.Cells[line,9].value  := FieldByName('Safety Days').Value;
        mysheet.Range['J'+IntToStr(line)].Formula:='=H'+IntToStr(line)+'*I'+IntToStr(line);
        mysheet.Range['H'+IntToStr(line),'I'+IntToStr(line)].Interior.ColorIndex:=34;
      end;

      Close;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.SELECT_SizeInfo;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@SizeCode';
      Parameters.ParamValues['@SizeCode'] := size;
      Open;
    end;

  except
    on e:exception do
    begin
      Data_Module.LogActLog('ERROR','Unable to get size info in order, '+e.Message);
      Data_Module.LogActLog('ERROR','Size = '+size);
      ShowMessage('Unable to get size info in order, '+e.Message);
      raise;
    end;
  end;
end;

procedure TOrder_Form.OrderHistory(partnumber:string;line:integer);
begin
  try
    with data_module.INV_Forecast_DataSet do
    begin
      Close;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.SELECT_OrderHistory;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@PartNumber';
      Parameters.ParamValues['@PartNumber'] := partnumber;
      Open;

      mysheet.Cells[line,5].value  := fieldbyname('Qty').AsInteger;
    end;
  except
    on e:exception do
    begin
      Data_Module.LogActLog('ERROR','Unable to get order history info in order, '+e.Message);
      ShowMessage('Unable to get order history info in order, '+e.Message);
      raise;
    end;
  end;
end;

procedure TOrder_Form.UpdateFRSInfo(partnumber:string;line:integer);
var
  count,i:integer;
begin
  try
    with data_module.INV_Forecast_DataSet do
    begin
      Close;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.SELECT_AssyRatioInfoRaw;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@BroadcastCode';
      Parameters.ParamValues['@BroadcastCode'] := '';
      if fTireWheel = 'T' then
      begin
        Parameters.AddParameter.Name := '@TirePartNum';
        Parameters.ParamValues['@TirePartNum'] := partnumber;
        Parameters.AddParameter.Name := '@WheelPartNum';
        Parameters.ParamValues['@WheelPartNum'] := '';
      end
      else
      begin
        Parameters.AddParameter.Name := '@TirePartNum';
        Parameters.ParamValues['@TirePartNum'] := '';
        Parameters.AddParameter.Name := '@WheelPartNum';
        Parameters.ParamValues['@WheelPartNum'] := partnumber;
      end;
      Open;
      // Place ratio on form
      //% plan from ratio table
      if fTireWheel = 'T' then
      begin
        if partnumber = fieldbyname('VC_TIRE_PART_NUMBER1_CODE').AsString then
        begin
          mysheet.Cells[line,7].value  := fieldbyname('IN_TIRE_RATIO1').AsFloat/100;
        end
        else
        begin
          mysheet.Cells[line,7].value  := fieldbyname('IN_TIRE_RATIO2').AsFloat/100;
        end;
      end
      else
      begin
        if partnumber = fieldbyname('VC_WHEEL_PART_NUMBER1_CODE').AsString then
        begin
          mysheet.Cells[line,7].value  := fieldbyname('IN_WHEEL_RATIO1').AsFloat/100;
        end
        else
        begin
          mysheet.Cells[line,7].value  := fieldbyname('IN_WHEEL_RATIO2').AsFloat/100;
        end;
      end;

      //
      //  Get previous production count to compare to forecast
      //
      //
      count:=0;
      for i:=1 to Data_Module.fiForecastUsageCompare.AsInteger do
      begin
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'dbo.SELECT_UsageDay;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@Date';
        Parameters.ParamValues['@Date'] := formatdatetime('yyyymmdd',now-i);
        Parameters.AddParameter.Name := '@PartNo';
        Parameters.ParamValues['@PartNo'] := partnumber;
        Open;
        count:=count+fieldbyname('Qty').asInteger;
      end;
    end;


    //
    //  Set count and compare to forecast set red if overproduced
    //
    mysheet.Cells[line+1,8].value  := 0;
    mysheet.Cells[line+1,9].value  := count;
    mysheet.Range['J'+IntToStr(line+1)].Formula  := '=H'+IntToStr(line+1)+'-I'+IntToStr(line+1);
    mysheet.Range['J'+IntToStr(line+1)].FormatConditions.Add(xlExpression, EmptyParam, '=$J$'+IntToStr(line+1)+'<0', EmptyParam);
    mysheet.Range['J'+IntToStr(line+1)].FormatConditions[1].font.ColorIndex:=3;
    //
    // Put total count for vendor share
    //
    if fSizeCount > 1 then
    begin
      mysheet.Range['J'+IntToStr(line+2)].Formula  := '=J'+IntToStr(line+1)+'+J'+IntToStr(line);
      mysheet.Range['J'+IntToStr(line+2)].FormatConditions.Add(xlExpression, EmptyParam, '=$J$'+IntToStr(line+2)+'<0', EmptyParam);
      mysheet.Range['J'+IntToStr(line+2)].FormatConditions[1].font.ColorIndex:=3;
    end
  except
    on e:exception do
    begin
      Data_Module.LogActLog('ERROR','Unable to get FRS info in order, '+e.Message);
      ShowMessage('Unable to get FRS info in order, '+e.Message);
      raise;
    end;
  end;
end;


procedure TOrder_Form.MonthForecast(partno:string; line:integer);
var
  z,total:integer;
begin
  //
  //  Change to 7 days instead of 30
  //
  try
    with data_module.INV_Forecast_DataSet do
    begin
      for z:=1 to Data_Module.fiForecastUsageCompare.AsInteger do
      begin
        Close;
        CommandType := CmdStoredProc;
        CommandText := 'dbo.SELECT_ForecastPartNumberWeek;1';
        Parameters.Clear;
        Parameters.AddParameter.Name := '@WeekNo';
        Parameters.ParamValues['@WeekNo'] := WeekOfTheYear(now-z);
        Parameters.AddParameter.Name := '@DayNo';
        Parameters.ParamValues['@DayNo'] := DayOfTheWeek(now-z);
        Parameters.AddParameter.Name := '@PartNo';
        Parameters.ParamValues['@PartNo'] := partno;//Data_Module.Inv_StoredProc.FieldByName('VC_PART_NUMBER').Value;
        Open;
        if not FieldByName('Qty').IsNull then
          total  := total + FieldByName('Qty').Value;
      end;
      mysheet.Cells[line+1,8].value  := total;
    end;
  except
    on e:exception do
    begin
      Data_Module.LogActLog('ERROR','Unable to get month forecast in order, '+e.Message);
      ShowMessage('Unable to get month forecast in order, '+e.Message);
      raise;
    end;
  end;
end;

//
// Fill Forecast array
//
procedure TOrder_Form.UpdateForecast(partno:string);
var
  j,z:integer;
begin
  try
    // Fill forecast information
    with data_module.INV_Forecast_DataSet do
    begin
      j:=0;
      for z:=0 to fdaterange do
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
          Parameters.ParamValues['@PartNo'] := partno;//Data_Module.Inv_StoredProc.FieldByName('VC_PART_NUMBER').Value;
          Open;
          if recordcount > 0 then
          begin
            if FieldByName('Qty').IsNull then
              fForecast[j]  := 0++FieldByName('Qty').Value
            else
              fForecast[j]  := fForecast[j]+FieldByName('Qty').Value;
          end
          else
            fForecast[j]:=0;
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

procedure TOrder_Form.PutWQSCount(partnumber:string;line:integer);
begin
  try
    with data_module.INV_Forecast_DataSet do
    begin
      Close;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.SELECT_OrderAtWQS;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@PartNumber';
      Parameters.ParamValues['@PartNumber'] := partnumber;
      Open;
      if not FieldByName('Qty').IsNull then
        mysheet.Cells[line,XLColToInt('K')].value  := FieldByName('Qty').Value
      else
        mysheet.Cells[line,XLColToInt('K')].value  := 0;
    end;
  except
    on e:exception do
    begin
      Data_Module.LogActLog('ERROR','Unable to get WQS in order, '+e.Message);
      ShowMessage('Unable to get WQS in order, '+e.Message);
      raise;
    end;
  end;
end;

procedure TOrder_Form.PutNUMMICount(partnumber:string;line:integer);
begin
  try
    with data_module.INV_Forecast_DataSet do
    begin
      Close;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.SELECT_OrderAtNUMMI;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@PartNumber';
      Parameters.ParamValues['@PartNumber'] := partnumber;
      Open;
      if not FieldByName('Qty').IsNull then
        mysheet.Cells[line,XLColToInt('L')].value  := FieldByName('Qty').Value
      else
        mysheet.Cells[line,XLColToInt('L')].value  := 0;
    end;
  except
    on e:exception do
    begin
      Data_Module.LogActLog('ERROR','Unable to get NUMMI in order, '+e.Message);
      ShowMessage('Unable to get NUMMI in order, '+e.Message);
      raise;
    end;
  end;
end;

procedure TOrder_Form.PutIntransitCount(partnumber:string;line:integer);
var
  i,z,daycount:integer;
  temp,lastFRS,firstFRS:string;
begin
  try
    with data_module.INV_Forecast_DataSet do
    begin
      Close;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.SELECT_OrderInTransit;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@PartNumber';
      Parameters.ParamValues['@PartNumber'] := partnumber;
      Open;
      if not FieldByName('Qty').IsNull then
        mysheet.Cells[line,XLColToInt('M')].value  := FieldByName('Qty').Value
      else
        mysheet.Cells[line,XLColToInt('M')].value  := 0;

      temp:=formatdatetime('yyyymmdd',fDates[0]);
      firstFRS:=copy(temp,4,1)+copy(temp,5,4)+'00';

      // Get count for calculation
      Close;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.SELECT_OrderInTransit;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@PartNumber';
      Parameters.ParamValues['@PartNumber'] := partnumber;
      Parameters.AddParameter.Name := '@FirstFRS';
      Parameters.ParamValues['@FirstFRS'] := FirstFRS;
      Open;
      if not FieldByName('Qty').IsNull then
        mysheet.Cells[line,XLColToInt('AR')].value  := FieldByName('Qty').Value
      else
        mysheet.Cells[line,XLColToInt('AR')].value  := 0;

      close;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.SELECT_OrderInTransitList;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@PartNumber';
      Parameters.ParamValues['@PartNumber'] := partnumber;
      Parameters.AddParameter.Name := '@FirstFRS';
      Parameters.ParamValues['@FirstFRS'] := FirstFRS;
      Open;

      if recordcount > 0 then
      begin
        daycount:=0;
        z:=0;
        lastFRS:='';
        while not eof do
        begin
          if lastFRS <> copy(FieldByName('VC_FRS_NUMBER').AsString,1,5) then
          begin
            if (daycount <> 0)  and (z<20)then //DMV
            begin
              mysheet.Cells[line,XLColToInt(DateWeekLetter)+z].value := daycount;
              mysheet.Range[IntToXLCol(XLColToInt(DateWeekLetter)+z)+IntToStr(line)].font.ColorIndex:=23;
            end;

            //new day
            daycount:=0;
            lastFRS:=copy(FieldByName('VC_FRS_NUMBER').AsString,1,5);
            z:=0;
            for i:=0 to fDateRange do
            begin
              if trunc(fDates[i]) = StrToDate(copy(lastFRS,2,2)+'/'+copy(lastFRS,4,2)+'/'+formatdatetime('yyyy',fDates[i])) then
                break;
              if trunc(fDates[i]) <> 0 then
                INC(z);
            end;
          end;
          daycount:=daycount+FieldByName('IN_QTY').AsInteger;
          next;
        end;
        if (daycount <> 0)  and (z<fDateRange)then//DMV
        begin
          mysheet.Cells[line,XLColToInt(DateWeekLetter)+z].value := daycount;
          mysheet.Range[IntToXLCol(XLColToInt(DateWeekLetter)+z)+IntToStr(line)].font.ColorIndex:=23;
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

procedure TOrder_Form.PutOpenOrderCount(partnumber:string;leadtime,line:integer);
var
  i,z,daycount:integer;
  temp,firstFRS,lastFRS:string;
begin
  try
    with data_module.INV_Forecast_DataSet do
    begin
      // Get count
      Close;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.SELECT_OrderOpenOrder;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@PartNumber';
      Parameters.ParamValues['@PartNumber'] := partnumber;
      Open;
      if not FieldByName('Qty').IsNull then
        mysheet.Cells[line,XLColToInt('N')].value  := FieldByName('Qty').Value
      else
        mysheet.Cells[line,XLColToInt('N')].value  := 0;
      // Get records and place number in correct date
      temp:=formatdatetime('yyyymmdd',fDates[0]);
      firstFRS:=copy(temp,4,1)+copy(temp,5,4)+'00';

      Close;
      CommandType := CmdStoredProc;
      CommandText := 'dbo.SELECT_OrderOpenOrderList;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@PartNumber';
      Parameters.ParamValues['@PartNumber'] := partnumber;
      Parameters.AddParameter.Name := '@FirstFRS';
      Parameters.ParamValues['@FirstFRS'] := FirstFRS;
      Open;
      if recordcount > 0 then
      begin
        daycount:=0;
        z:=0;
        lastFRS:='';
        while not eof do
        begin
          if lastFRS <> copy(FieldByName('VC_FRS_NUMBER').AsString,1,5) then
          begin
            if (daycount <> 0)  and (z<fDateRange)then //dmv
            begin
              mysheet.Cells[line,XLColToInt(DateWeekLetter)+z].value := daycount;
              mysheet.Range[IntToXLCol(XLColToInt(DateWeekLetter)+z)+IntToStr(line)].font.ColorIndex:=10;
            end;

            //new day
            daycount:=0;
            lastFRS:=copy(FieldByName('VC_FRS_NUMBER').AsString,1,5);
            z:=0;
            for i:=0 to fDateRange do
            begin
              if trunc(fDates[i]) = StrToDate(copy(lastFRS,2,2)+'/'+copy(lastFRS,4,2)+'/'+formatdatetime('yyyy',now)) then
                break;
              if trunc(fDates[i]) <> 0 then
                INC(z);
            end;
          end;
          daycount:=daycount+FieldByName('IN_QTY').AsInteger;
          next;
        end;
        if (daycount <> 0)  and (z<fDateRange)then
        begin
          mysheet.Cells[line,XLColToInt(DateWeekLetter)+z].value := daycount;
          mysheet.Range[IntToXLCol(XLColToInt(DateWeekLetter)+z)+IntToStr(line)].font.ColorIndex:=10;
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
  for i:=0 to fFillDays-1 do
  begin
    mysheet.Cells[line,i+(DateWeekCol+1)].value := fForecast[i];
  end;
end;

procedure TOrder_Form.ClearForecast;
var
  i:integer;
begin
  for i:=0 to fDateRange do
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
  mysheet.Range[DateWeekLetter+IntToStr(topedge)].Formula:='=SUM(AQ'+IntToStr(topedge+1)+':AQ'+IntToStr(bottomedge)+')-SUM(AR'+IntToStr(topedge+1)+':AR'+IntToStr(bottomedge)+')';

  for i:=0 to fFillDays-1 do
  begin
//    if fForecast[i] <> 0 then
  //  begin
      //if fForecast[i+1] <> 0 then
      if i <> fFillDays-1 then
        mysheet.Range[IntToXLCol((ForecastCol+1)+i)+IntToStr(topedge)].Formula:='='+IntToXLCol(ForecastCol+i)+IntToStr(bottomedge);
      if bottomedge-topedge = 3 then
      begin
        mysheet.Range[IntToXLCol(ForecastCol+i)+IntToStr(bottomedge)].Formula:='='+IntToXLCol(ForecastCol+i)+IntToStr(topedge)+'+'+IntToXLCol(ForecastCol+i)+IntToStr(topedge+1)+'-'+IntToXLCol(ForecastCol+i)+IntToStr(topedge+2);
        mysheet.Range[IntToXLCol(ForecastCol+i)+IntToStr(bottomedge)].FormatConditions.Add(xlCellValue, xlLess, '=$J$'+IntToStr(topedge), EmptyParam);
        mysheet.Range[IntToXLCol(ForecastCol+i)+IntToStr(bottomedge)].FormatConditions[1].font.ColorIndex:=3;
      end
      else
      begin
        mysheet.Range[IntToXLCol(ForecastCol+i)+IntToStr(bottomedge)].Formula:='='+IntToXLCol(ForecastCol+i)+IntToStr(topedge)+'+'+IntToXLCol(ForecastCol+i)+IntToStr(topedge+1)+'+'+IntToXLCol(ForecastCol+i)+IntToStr(topedge+2)+'-'+IntToXLCol(ForecastCol+i)+IntToStr(topedge+3);
        mysheet.Range[IntToXLCol(ForecastCol+i)+IntToStr(bottomedge)].FormatConditions.Add(xlCellValue, xlLess, '=$J$'+IntToStr(topedge), EmptyParam);
        mysheet.Range[IntToXLCol(ForecastCol+i)+IntToStr(bottomedge)].FormatConditions[1].font.ColorIndex:=3;
      end;
    //end;
  end;
end;

procedure TOrder_Form.DoLeadTime(Leadtime,Line:integer);
var
  order:string;
  i,addedleadtime:integer;
begin
  addedleadtime:=0;
  for i:=0 to fOvertimeCount-1 do
  begin
    if  (fOvertimes[i]-1) <= (leadtime+i) then
      INC(addedleadtime)
    else
      break;
  end;
  mysheet.Cells[line,XLColToInt('AS')].value  := addedleadtime;

  mysheet.Range[DateWeekLetter+IntToStr(line),IntToXLCol(XLColToInt(DateWeekLetter)+((LeadTime-1)+addedleadtime))+IntToStr(line)].Interior.ColorIndex:=36;
  mysheet.Range[IntToXLCol(XLColToInt(DateWeekLetter)+LeadTime+addedleadtime)+IntToStr(line)].Interior.ColorIndex:=40;

  for i:=0 to fOvertimeCount-1 do
  begin
    mysheet.Range[IntToXLCol(XLColToInt(DateWeekLetter)+(fOvertimes[i]-1))+IntToStr(line)].Interior.ColorIndex:=3;
  end;

  fPartLine[line,FRSDate]:=mysheet.Cells[5,XLColToInt(DateWeekLetter)+LeadTime+addedleadtime].value;
  order:=mysheet.Cells[line,XLColToInt(DateWeekLetter)+LeadTime+addedleadtime].value;
  if order = ''  then
    mysheet.Range[IntToXLCol(XLColToInt(DateWeekLetter)+LeadTime+addedleadtime)+IntToStr(line)].Formula:='=Q'+IntToStr(line)
  else
    mysheet.Range['Q'+IntToStr(line),'R'+IntToStr(line)].locked:=TRUE;
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


procedure TOrder_Form.FormShow(Sender: TObject);
begin
  ProcessOrder_Button.Enabled:=False;
end;

end.
