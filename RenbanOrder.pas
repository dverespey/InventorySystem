  //****************************************************************
//
//       Inventory Control
//
//       Copyright (c) 2002-2003 Failproof Manufacturing Systems.
//
//****************************************************************
//
//  03/26/2003  David Verespey  Initial creation

unit RenbanOrder;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, Buttons, StrUtils;

type
  TOrder = class(TObject)
  private
  public
    kanban:string;
    supplier:string;
    partnumber:string;
    frsnumber:string;
    lotqty:string;
    lots:string;
    renban:string;
  end;

  TTruck = class(TObject)
  private
    fSize:integer;
    fCurrentCount:integer;
    fOrderCount:integer;
    fOrderList:TStringList;
    fOrder:TOrder;
    fEof:boolean;
    fOrderOffset:integer;
    fOrderName:string;
  public
    constructor Create; virtual;
    procedure AddOrder(kanban, supplier, partnumber, frsnumber, lotqty, lots, renban: string);

    property Size :integer
    read fSize
    write fSize;

    property CurrentCount:integer
    read fCurrentCount
    write fCurrentCount;

    function First: boolean;
    function Next:boolean;

    property Eof: boolean
    read fEof;

    property Order:TOrder
    read fOrder;
  end;

  TGroupRenban = class(TObject)
  private
    fTruckList:TStringList;
    fEof:boolean;
    fTruckOffset:integer;

    fTRuck:TTruck;
    fTruckName:string;
  public
    constructor Create; virtual;

    procedure SetTrucks(NumberOfTrucks, SizeOfTRuck :integer);
    procedure AddOrder(kanban, supplier, partnumber, frsnumber, orderqty, lotqty, lots, renban: string);

    function First: boolean;
    function Next:boolean;

    property Eof: boolean
    read fEof;

    property Truck:TTruck
    read fTruck;

    property TruckName:string
    read fTruckName;

    property TruckNumber: integer
    read fTruckOffset;
  end;

  TGroupRenbanOrder_Form = class(TForm)
    RenbanGroups_ComboBox: TComboBox;
    AvailableGrid: TStringGrid;
    CreateOrder_Button: TBitBtn;
    OKButton: TButton;
    Label1: TLabel;
    FRSBreakdown_Button: TBitBtn;
    Label2: TLabel;
    TotalLots_Edit: TEdit;
    Trailers_ComboBox: TComboBox;
    Label3: TLabel;
    ClearBreakdown_BitBtn: TBitBtn;
    Label4: TLabel;
    TrailerPalletCount_Edit: TEdit;
    TRailerCounts_ListBox: TListBox;
    procedure CreateOrder_ButtonClick(Sender: TObject);
    procedure RenbanGroups_ComboBoxChange(Sender: TObject);
    procedure OKButtonClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FRSBreakdown_ButtonClick(Sender: TObject);
    procedure ClearBreakdown_BitBtnClick(Sender: TObject);
    procedure Trailers_ComboBoxChange(Sender: TObject);
    procedure TrailerPalletCount_EditChange(Sender: TObject);
  private
    { Private declarations }
    fAvailableCount:integer;
    fBreakdownWaiting:boolean;
    fNewMaxRenban:integer;
    GroupRenban:TGroupRenban;
    fFRSOffset:string;
    fTRailerChange:boolean;
    function LoadScreen:boolean;
    procedure NewFRSOrder();
    procedure FreeList;
  public
    { Public declarations }
    procedure Execute;
  end;

var
  GroupRenbanOrder_Form: TGroupRenbanOrder_Form;

implementation

uses DataModule, FRSBreakdown;

{$R *.dfm}

constructor TTruck.Create;
begin
  inherited Create;

  fOrderList:=TStringList.Create;

  fOrderCount:=0;
  fOrder:=nil;
  fEof:=TRUE;
  fOrderOffset:=0;
  fOrderName:='';
end;

procedure TTruck.AddOrder(kanban, supplier, partnumber, frsnumber, lotqty, lots, renban: string);
var
  order:TOrder;
  x:integer;
begin
  x:=fOrderList.IndexOf(partnumber);
  if x <> -1 then
  begin
    TOrder(fOrderList.Objects[x]).lots:=IntToStr(StrToInt(TOrder(fOrderList.Objects[x]).lots)+StrToInt(lots));
  end
  else
  begin
    order:=TOrder.Create;
    order.kanban:=kanban;
    order.supplier:=supplier;
    order.partnumber:=partnumber;
    order.frsnumber:=frsnumber;
    order.lotqty:=lotqty;
    order.lots:=lots;
    order.renban:=renban;

    fOrderList.AddObject(partnumber, order);
  end;

  fCurrentCount:=fCurrentCount+StrToInt(lots);

  INC(fOrderCount);
end;

function TTruck.First: boolean;
begin
  result:=TRUE;

  try
    fOrderOffset:=0;

    if fOrderList.Count > 0 then
    begin
      fOrder:=TOrder(fOrderList.Objects[fOrderOffset]);
      fOrderName:=fOrderList[fOrderOffset];
      fEof:=FALSE;
    end
    else
    begin
      fEof:=TRUE;
      fOrder:=nil;
    end;
  except
    result:=FALSE;
  end;
end;

function TTruck.Next:boolean;
begin
  result:=TRUE;

  try
    INC(fOrderOffset);

    if fOrderOffset <= fOrderList.Count-1 then
    begin
      fOrder:=TOrder(fOrderList.Objects[fOrderOffset]);
      fOrderName:=fOrderList[fOrderOffset];
      fEof:=FALSE;
    end
    else
    begin
      fEof:=TRUE;
      fOrder:=nil;
    end;
  except
    result:=FALSE;
  end;
end;

constructor TGroupRenban.Create;
begin
  inherited Create;

  fEof:=TRUE;
  fTruckOffset:=0;
  fTruckName:='';
end;

procedure TGroupRenban.SetTrucks(NumberOfTrucks, SizeOfTRuck :integer);
var
  i:integer;
  truck:TTruck;
begin
  fTruckList:=TStringList.Create;

  for i:=1 to NumberOfTrucks do
  begin
    truck:=TTruck.Create;
    truck.Size:=SizeOfTruck;

    fTruckList.AddObject('Trailer'+IntToStr(i), truck);
  end;
end;

procedure TGroupRenban.AddOrder(kanban, supplier, partnumber, frsnumber, orderqty, lotqty, lots, renban:string);
var
  i,leftover, remainder:integer;
begin

  leftover:=0;

  if StrToInt(lots) div fTruckList.Count <> 0 then
  begin
    for i:=0 to fTruckList.Count-1 do
    begin
      if (TTruck(fTruckList.Objects[i]).CurrentCount) + (StrToInt(lots) div fTruckList.Count) + leftover <= (TTruck(fTruckList.Objects[i]).Size) then
      begin
        TTruck(fTruckList.Objects[i]).AddOrder(kanban, supplier, partnumber, frsnumber, lotqty, IntToStr((StrToInt(lots) div fTruckList.Count) + leftover), renban);
        leftover:=0;
      end
      else
      begin
        leftover := ((StrToInt(lots) div fTruckList.Count) + leftover) -  ((TTruck(fTruckList.Objects[0]).Size) -  TTruck(fTruckList.Objects[i]).CurrentCount);
        TTruck(fTruckList.Objects[i]).AddOrder(kanban, supplier, partnumber, frsnumber, lotqty, IntToStr((TTruck(fTruckList.Objects[i]).Size) -  TTruck(fTruckList.Objects[i]).CurrentCount), renban)
      end;
    end;
  end;

  if StrToInt(lots) mod fTruckList.Count <> 0 then
  begin
    // break up remainder
    remainder:=(StrToInt(lots) mod fTruckList.Count) + leftover;


    while remainder <> 0 do
    begin
      for i:=0 to fTruckList.Count-1 do
      begin
        if (TTruck(fTruckList.Objects[i]).CurrentCount) + 1 <= (TTruck(fTruckList.Objects[i]).Size) then
        begin
          TTruck(fTruckList.Objects[i]).AddOrder(kanban, supplier, partnumber, frsnumber, lotqty, '1', renban);
          remainder:=remainder-1;
        end;

        if remainder = 0 then
          break;
      end;
    end;

  end;
end;

function TGroupRenban.First: boolean;
begin
  result:=TRUE;

  try
    fTruckOffset:=0;

    if fTruckList.Count > 0 then
    begin
      fTruck:=TTruck(ftruckList.Objects[fTruckOffset]);
      fTruckName:=fTruckList[fTruckOffset];
      fEof:=FALSE;
    end
    else
    begin
      fEof:=TRUE;
      fTruck:=nil;
    end;
  except
    result:=FALSE;
  end;
end;

function TGroupRenban.Next:boolean;
begin
  result:=TRUE;

  try
    INC(fTruckOffset);

    if fTruckOffset <= fTruckList.Count-1 then
    begin
      fTruck:=TTruck(ftruckList.Objects[fTruckOffset]);
      fTruckName:=fTruckList[fTruckOffset];
    end
    else
    begin
      fEof:=TRUE;
      fTruck:=nil;
    end;
  except
    result:=FALSE;
  end;
end;

procedure TGroupRenbanOrder_Form.Execute;
begin
  // Fill renaban groups
  try
    RenbanGroups_ComboBox.Items.Clear;
    RenbanGroups_ComboBox.Text:='';

    with Data_Module.Inv_StoredProc do
    begin
      //Renban
      Close;
      ProcedureName := 'dbo.SELECT_RenbanGroup;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@RebanCode';
      Parameters.ParamValues['@RebanCode'] := '';
      Open;

      while not eof do
      begin
        RenbanGroups_ComboBox.Items.Add(FieldByName('Renban Group Code').AsString);
        next;
      end;
      Close;

      AvailableGrid.Cells[0,0]:='Kanban';
      AvailableGrid.Cells[1,0]:='Supplier';
      AvailableGrid.Cells[2,0]:='Part Number';
      AvailableGrid.Cells[3,0]:='FRS Number';
      AvailableGrid.Cells[4,0]:='Order Qty';
      AvailableGrid.Cells[5,0]:='Lot Qty';
      AvailableGrid.Cells[6,0]:='Lots';
      AvailableGrid.Cells[7,0]:='Renban';
      AvailableGrid.RowCount:=2;
      fAvailableCount:=0;
    end;

    fBreakdownWaiting:=FALSE;
    TrailerPalletCount_Edit.Text:='';
    TotalLots_Edit.Text:='';

    //Set buttons
    ClearBreakdown_BitBtn.Enabled:=False;
    CreateOrder_Button.Enabled:=False;
    FRSBreakdown_Button.Enabled:=False;
    RenbanGroups_ComboBox.Enabled:=True;
    TrailerCounts_ListBox.Visible:=False;

    ShowModal;

  except
    on e:exception do
    begin
      ShowMessage('Unable to get Renban Group Orders, '+e.message);
      Data_Module.LogActLog('ERROR','Unable to get Renban Group Orders, '+e.message);
    end;
  end;
end;

procedure TGroupRenbanOrder_Form.CreateOrder_ButtonClick(Sender: TObject);
var
  i:integer;
begin
  // create new order
  if fBreakdownWaiting then
  begin
    if MessageDlg('Update these records?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      try
        Data_Module.Inv_Connection.BeginTrans;
        for i:=1 to fAvailableCount do
        begin
          AvailableGrid.Row:=i;
          NewFRSOrder;
        end;
        // Update Renban Number
        With Data_Module.Inv_StoredProc do
        Begin
          Close;
          ProcedureName := 'dbo.UPDATE_RenbanGroupCount;1';
          Parameters.Clear;
          Parameters.AddParameter.Name := '@RenbanCode';
          Parameters.ParamValues['@RenbanCode'] := RenbanGroups_ComboBox.Text;
          Parameters.AddParameter.Name := '@RenbanCount';
          Parameters.ParamValues['@RenbanCount'] := Format('%.3d',[fNewMaxRenban]);
          ExecProc;
          Data_Module.LogActLog('RENBAN BD','Set next Renban Group Number Renban='+RenbanGroups_ComboBox.Text+', Number='+Format('%.3d',[fNewMaxRenban]));
        end;

        Data_Module.Inv_Connection.CommitTrans;

        //clear screen
        for i:=1 to fAvailableCount do
        begin
          AvailableGrid.Cells[0,i]:='';
          AvailableGrid.Cells[1,i]:='';
          AvailableGrid.Cells[2,i]:='';
          AvailableGrid.Cells[3,i]:='';
          AvailableGrid.Cells[4,i]:='';
          AvailableGrid.Cells[5,i]:='';
          AvailableGrid.Cells[6,i]:='';
          AvailableGrid.Cells[7,i]:='';
        end;
        AvailableGrid.RowCount:=2;
        //Clear List
        FreeList;

        //Set buttons
        ClearBreakdown_BitBtn.Enabled:=False;
        CreateOrder_Button.Enabled:=False;
        FRSBreakdown_Button.Enabled:=False;
        RenbanGroups_ComboBox.Enabled:=True;
        TrailerCounts_ListBox.Visible:=False;

        fBreakdownWaiting:=FALSE;

        TrailerPalletCount_Edit.Text:='';
        TotalLots_Edit.Text:='';
        RenbanGroups_ComboBox.Text:='';
        Trailers_ComboBox.Text:='';

      except
        on e:exception do
        begin
          Data_Module.LogActLog('ERROR','Unable to update these records,'+e.Message);
          ShowMessage('Unable to update these records,'+e.Message);
          if Data_Module.Inv_Connection.InTransaction then
            Data_Module.Inv_Connection.RollbackTrans;
        end;
      end;
    end;
  end;
end;

procedure TGroupRenbanOrder_Form.NewFRSOrder();
//var
  //OS:string;
begin
  try
    with Data_Module.Inv_StoredProc do
    begin
      //
      //  With new break down an error can occur where only one item order can be put on any FRS truck
      //  if it's not the first truck the inital order is left without a Renban number
      ///
      //  To rectify, delete all orders and re-add with renban number FRS
      //
      //
      //Update the old record
      //if copy(AvailableGrid.Cells[3,AvailableGrid.Row],6,2) = fFRSOffset { Update first record}  then
      //begin
        //if  StrToInt(AvailableGrid.Cells[4,AvailableGrid.Row]) = 0 then
        //begin
          // small runner don't send 0 qty orders
          // delete and insert later
          Close;
          ProcedureName := 'dbo.DELETE_OrderRenban;1';
          Parameters.Clear;
          Parameters.AddParameter.Name := '@PartNumber';
          Parameters.ParamValues['@PartNumber'] := AvailableGrid.Cells[2,AvailableGrid.Row];
          Parameters.AddParameter.Name := '@FRSNumber';
          Parameters.ParamValues['@FRSNumber'] := '';//copy(AvailableGrid.Cells[3,AvailableGrid.Row],1,5)+'01';
          Parameters.AddParameter.Name := '@RenbanNumber';
          Parameters.ParamValues['@RenbanNumber'] := '';
          ExecProc;
          Data_Module.LogActLog('RENBAN BD','Delete Temp FRS Renban Group Order, p:'
                                +AvailableGrid.Cells[2,AvailableGrid.Row]
                                +'f:'+AvailableGrid.Cells[3,AvailableGrid.Row]);
        {end
        else
        begin
          //Update inital Renban
          Close;
          ProcedureName := 'dbo.UPDATE_OrderRenbanQty;1';
          Parameters.Clear;
          Parameters.AddParameter.Name := '@PartNumber';
          Parameters.ParamValues['@PartNumber'] := AvailableGrid.Cells[2,AvailableGrid.Row];
          Parameters.AddParameter.Name := '@FRSNumber';
          Parameters.ParamValues['@FRSNumber'] := AvailableGrid.Cells[3,AvailableGrid.Row];
          Parameters.AddParameter.Name := '@RenbanNumber';
          Parameters.ParamValues['@RenbanNumber'] := AvailableGrid.Cells[7,AvailableGrid.Row];
          Parameters.AddParameter.Name := '@Qty';
          Parameters.ParamValues['@Qty'] := AvailableGrid.Cells[4,AvailableGrid.Row];
          ExecProc;
          Data_Module.LogActLog('RENBAN BD','Update record, p:'
                                +AvailableGrid.Cells[2,AvailableGrid.Row]
                                +'f:'+AvailableGrid.Cells[3,AvailableGrid.Row]
                                +'r:'+AvailableGrid.Cells[7,AvailableGrid.Row]
                                +'q:'+AvailableGrid.Cells[4,AvailableGrid.Row]);
        end;
      end
      else
      begin}
        if  StrToInt(AvailableGrid.Cells[4,AvailableGrid.Row]) > 0 then
        begin
          Close;
          ProcedureName := 'dbo.INSERT_OpenOrder;1';
          Parameters.Clear;
          Parameters.AddParameter.Name := '@SupCode';
          Parameters.ParamValues['@SupCode'] := AvailableGrid.Cells[1,AvailableGrid.Row];
          Parameters.AddParameter.Name := '@PartNum';
          Parameters.ParamValues['@PartNum'] := AvailableGrid.Cells[2,AvailableGrid.Row];
          Parameters.AddParameter.Name := '@KanbanNum';
          Parameters.ParamValues['@KanbanNum'] := AvailableGrid.Cells[0,AvailableGrid.Row];
          Parameters.AddParameter.Name := '@FRSNumber';
          Parameters.ParamValues['@FRSNumber'] := AvailableGrid.Cells[3,AvailableGrid.Row];
          Parameters.AddParameter.Name := '@RenbanNum';
          Parameters.ParamValues['@RenbanNum'] := AvailableGrid.Cells[7,AvailableGrid.Row];
          Parameters.AddParameter.Name := '@Qty';
          Parameters.ParamValues['@Qty'] := AvailableGrid.Cells[4,AvailableGrid.Row];
          ExecProc;
          Data_Module.LogActLog('RENBAN BD','Insert record, p:'
                                +AvailableGrid.Cells[2,AvailableGrid.Row]
                                +'f:'+AvailableGrid.Cells[3,AvailableGrid.Row]
                                +'r:'+AvailableGrid.Cells[7,AvailableGrid.Row]
                                +'q:'+AvailableGrid.Cells[4,AvailableGrid.Row]);
        end;
      //end;
    end;
  except
    on e:exception do
    begin
      Data_Module.LogActLog('ERROR','Failed to create new FRS numbers, '+e.Message);
      raise;
    end;
  end;
end;

function TGroupRenbanOrder_Form.LoadScreen:boolean;
var
  i,lots:integer;
begin
  // if in process question to save
  // Load available grid and clear selected
  fTrailerChange:=True;
  result:=True;
  try

    with Data_Module.Inv_StoredProc do
    begin
      Close;
      ProcedureName := 'dbo.SELECT_OrderNoRenban;1';
      Parameters.Clear;
      Parameters.AddParameter.Name := '@RenbanGroupCode';
      Parameters.ParamValues['@RenbanGroupCode'] := RenbanGroups_comboBox.Text;
      Open;

      //clear previous
      for i:=1 to AvailableGrid.RowCount-1 do
      begin
        AvailableGrid.Cells[0,i]:='';
        AvailableGrid.Cells[1,i]:='';
        AvailableGrid.Cells[2,i]:='';
        AvailableGrid.Cells[3,i]:='';
        AvailableGrid.Cells[4,i]:='';
        AvailableGrid.Cells[5,i]:='';
        AvailableGrid.Cells[6,i]:='';
        AvailableGrid.Cells[7,i]:='';
      end;

      if recordcount > 0 then
      begin
        fAvailableCount:=recordcount;
        AvailableGrid.RowCount:=recordcount+1;
        i:=1;
        lots:=0;
        while not eof do
        begin
          AvailableGrid.Cells[0,i]:=FieldByName('VC_KANBAN_NUMBER').AsString;
          AvailableGrid.Cells[1,i]:=FieldByName('VC_SUPPLIER_CODE').AsString;
          AvailableGrid.Cells[2,i]:=FieldByName('VC_PART_NUMBER').AsString;
          AvailableGrid.Cells[3,i]:=FieldByName('VC_FRS_NUMBER').AsString;
          fFRSOffset:=copy(AvailableGrid.Cells[3,AvailableGrid.Row],6,2);
          AvailableGrid.Cells[4,i]:=FieldByName('IN_QTY').AsString;
          AvailableGrid.Cells[5,i]:=FieldByName('IN_1LOTQTY').AsString;
          AvailableGrid.Cells[6,i]:=IntToStr(StrToInt(FieldByName('IN_QTY').AsString) div StrToInt(FieldByName('IN_1LOTQTY').AsString));
          lots:=lots+StrToInt(FieldByName('IN_QTY').AsString) div StrToInt(FieldByName('IN_1LOTQTY').AsString);
          AvailableGrid.Cells[7,i]:=FieldByName('VC_RENBAN_GROUP_CODE').AsString+FieldByName('VC_RENBAN_GROUP_COUNT').AsString;
          TotalLots_Edit.Text:=IntToStr(lots);
          INC(i);
          next;
        end;
        //Set buttons
        ClearBreakdown_BitBtn.Enabled:=False;
        CreateOrder_Button.Enabled:=False;
        FRSBreakdown_Button.Enabled:=True;
        RenbanGroups_ComboBox.Enabled:=True;
        TrailerCounts_ListBox.Visible:=False;
      end
      else
      begin
        AvailableGrid.RowCount:=2;
        AvailableGrid.Cells[0,1]:='';
        AvailableGrid.Cells[1,1]:='';
        AvailableGrid.Cells[2,1]:='';
        AvailableGrid.Cells[3,1]:='';
        AvailableGrid.Cells[4,1]:='';
        AvailableGrid.Cells[5,1]:='';
        AvailableGrid.Cells[6,1]:='';
        AvailableGrid.Cells[7,1]:='';

        ShowMessage('No records for this Renban Group');
        result:=False;
      end;

      Close;
    end;
  except
    on e:exception do
    begin
      Data_Module.LogActLog('ERROR','Unable to get Renban Group Orders, '+e.message);
      result:=False;
    end;
  end;
end;

procedure TGroupRenbanOrder_Form.RenbanGroups_ComboBoxChange(Sender: TObject);
begin
  //clear previous
  if LoadScreen then
  begin
    Trailers_ComboBox.SetFocus;
    //Set buttons
    ClearBreakdown_BitBtn.Enabled:=False;
    CreateOrder_Button.Enabled:=False;
    FRSBreakdown_Button.Enabled:=True;
    TrailerCounts_ListBox.Visible:=False;
  end;
end;

procedure TGroupRenbanOrder_Form.OKButtonClick(Sender: TObject);
begin
  Close;
end;

procedure TGroupRenbanOrder_Form.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  if fBreakdownWaiting then
  begin
    if MessageDlg('There is an unprocessed breakdown waiting, close anyway?', mtConfirmation, [mbYes, mbNo], 0) = mrNo then
      CanClose:=False
    else
    begin
      FreeList;
      CanClose:=True;
    end;
  end
  else
    CanClose:=True;
end;

procedure TGroupRenbanOrder_Form.FRSBreakdown_ButtonClick(Sender: TObject);
var
  {max, smax, modt, j,} i, x, {lcount,} tcount, {ttcount, }pcount, {modpass, }rcount:integer;

begin

  rcount:=0;
  //fRenban:=copy(AvailableGrid.Cells[7,1],1,5);
  if TryStrToInt(Trailers_ComboBox.Text, tcount) and TryStrToInt(TrailerPalletCount_Edit.Text, pcount) then
  begin
    if tcount*pcount >= StrToInt(TotalLots_Edit.text) then
    begin

      Data_Module.LogActLog('RENBAN BD','Total Lots:'+TotalLots_Edit.Text+', Trailer Pallet Count:'+TrailerPalletCount_Edit.Text+', Trailers:'+Trailers_ComboBox.Text);

      GroupRenban:=TGroupRenban.Create;

      GroupRenban.SetTrucks(tcount,pcount);


      for i:=1 to fAvailableCount do
      begin
        GroupRenban.AddOrder( AvailableGrid.Cells[0,i],
                              AvailableGrid.Cells[1,i],
                              AvailableGrid.Cells[2,i],
                              AvailableGrid.Cells[3,i],
                              AvailableGrid.Cells[4,i],
                              AvailableGrid.Cells[5,i],
                              AvailableGrid.Cells[6,i],
                              AvailableGrid.Cells[7,i]);
      end;

      //clear screen
      for i:=1 to AvailableGrid.RowCount-1 do
      begin
        AvailableGrid.Cells[0,i]:='';
        AvailableGrid.Cells[1,i]:='';
        AvailableGrid.Cells[2,i]:='';
        AvailableGrid.Cells[3,i]:='';
        AvailableGrid.Cells[4,i]:='';
        AvailableGrid.Cells[5,i]:='';
        AvailableGrid.Cells[6,i]:='';
        AvailableGrid.Cells[7,i]:='';
      end;

      AvailableGrid.RowCount:=2;

      x:=1;
      if GroupRenban.First then
      begin
        while not GroupRenban.Eof do
        begin
          if GroupRenban.Truck.First then
          begin
            while not GroupRenban.Truck.Eof do
            begin
              if x = AvailableGrid.RowCount then
                AvailableGrid.RowCount:=AvailableGrid.RowCount+1;

              AvailableGrid.Cells[0,x]:=GroupRenban.Truck.Order.kanban;
              AvailableGrid.Cells[1,x]:=GroupRenban.Truck.Order.supplier;
              AvailableGrid.Cells[2,x]:=GroupRenban.Truck.Order.partnumber;


              AvailableGrid.Cells[3,x]:=copy(GroupRenban.Truck.Order.frsnumber,1,5);
              if GroupRenban.TruckNumber > 8 Then
                AvailableGrid.Cells[3,x] := AvailableGrid.Cells[3,x] + IntToStr(GroupRenban.TruckNumber + 1)
              Else
                AvailableGrid.Cells[3,x] := AvailableGrid.Cells[3,x] + '0' + IntToStr(GroupRenban.TruckNumber + 1);


              AvailableGrid.Cells[4,x]:=IntToStr(StrToInt(GroupRenban.Truck.Order.lots)*StrToInt(GroupRenban.Truck.Order.lotqty));
              AvailableGrid.Cells[5,x]:=GroupRenban.Truck.Order.lotqty;
              AvailableGrid.Cells[6,x]:=GroupRenban.Truck.Order.lots;


              rcount:=StrToInt(rightstr(GroupRenban.Truck.Order.renban,3));
              //rcount:=StrToInt(copy(fTruckItems[i-1].Values['RENBAN'],6,3));
              rcount:=rcount+GroupRenban.TruckNumber;
              //AvailableGrid.Cells[7,x]:=copy(fTruckItems[i-1].Values['RENBAN'],1,5)+format('%.3d',[rcount]);
              AvailableGrid.Cells[7,x]:=RenbanGroups_ComboBox.Text+format('%.3d',[rcount]);

              Data_Module.LogActLog('RENBAN BD', 'New Renban breakdown item, s:'
                                    +GroupRenban.Truck.Order.supplier
                                    +' p:'+GroupRenban.Truck.Order.partnumber
                                    +' f:'+AvailableGrid.Cells[3,x]
                                    +' r:'+AvailableGrid.Cells[7,x]
                                    +' l:'+AvailableGrid.Cells[6,x]);
              INC(x);

              GroupRenban.Truck.Next;
            end;
          end;
          TrailerCounts_ListBox.Items.Add(GroupRenban.TruckName+' count ='+IntToStr(GroupRenban.Truck.CurrentCount));
          GRoupRenban.Next;
        end;
      end;


      fNewMaxRenban:=rcount+1;
      fAvailableCount:=AvailableGrid.RowCount-1;

      //Set buttons
      ClearBreakdown_BitBtn.Enabled:=True;
      CreateOrder_Button.Enabled:=True;
      FRSBreakdown_Button.Enabled:=False;
      RenbanGroups_ComboBox.Enabled:=False;
      TrailerCounts_ListBox.Visible:=True;
      TrailerCounts_ListBox.Visible:=True;

      fBreakdownwaiting:=True;

    end
    else
    begin
      Showmessage('The total lots will not fit on the selected number of trailers max('+IntToStr(tcount*pcount)+'):current('+TotalLots_Edit.Text+')');
    end;
  end
  else
  begin
    if not TryStrToInt(Trailers_ComboBox.Text, tcount) then
    begin
      SHowMessage('Please select trailer count');
      Trailers_ComboBox.SetFocus;
    end
    else if not TryStrToInt(TrailerPalletCount_Edit.Text, pcount) then
    begin
      SHowMessage('Please select trailer pallet count');
      TrailerPalletCount_Edit.SetFocus;
    end;
  end;
end;

procedure TGroupRenbanOrder_Form.ClearBreakdown_BitBtnClick(
  Sender: TObject);
var
  i:integer;
begin
  if fBreakdownWaiting then
  begin
    if MessageDlg('There is an unprocessed breakdown waiting, clear anyway?', mtConfirmation, [mbYes, mbNo], 0) = mrNo then
      exit;
  end;
  for i:=1 to AvailableGrid.RowCount-1 do
  begin
    AvailableGrid.Cells[0,i]:='';
    AvailableGrid.Cells[1,i]:='';
    AvailableGrid.Cells[2,i]:='';
    AvailableGrid.Cells[3,i]:='';
    AvailableGrid.Cells[4,i]:='';
    AvailableGrid.Cells[5,i]:='';
    AvailableGrid.Cells[6,i]:='';
    AvailableGrid.Cells[7,i]:='';
  end;
  FreeList;
  TRailerCounts_ListBox.Items.Clear;
  fBreakdownWaiting:=False;
  LoadScreen;
  Data_Module.LogActLog('RENBAN BD','Clear Renban breakdown');
end;

procedure TGroupRenbanOrder_Form.FreeList;
begin

  GroupRenban.Free;

end;

procedure TGroupRenbanOrder_Form.Trailers_ComboBoxChange(Sender: TObject);
begin
  //auto calculated trailer quantitiy
  try
    if Trailers_ComboBox.Text <> '' then
    begin
      fTrailerChange:=False;
      TrailerPalletCount_Edit.Text := IntToStr(StrToInt(TotalLots_Edit.Text) div StrToInt(Trailers_ComboBox.Text));
    end;
  finally
    fTrailerChange:=True;
  end;
end;

procedure TGroupRenbanOrder_Form.TrailerPalletCount_EditChange(
  Sender: TObject);
var
  x:integer;
begin
  //auto calculate trailers
  if fTrailerChange then
  begin
    if TryStrToInt(TrailerPalletCount_Edit.Text, x) then
    begin
      Trailers_ComboBox.ItemIndex := (StrToInt(TotalLots_Edit.Text) div x);
    end;
  end;
end;

end.
