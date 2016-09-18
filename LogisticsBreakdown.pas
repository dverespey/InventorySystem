unit LogisticsBreakdown;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, History,DB, CopyFile;

type
  TLogisticsBreakdown_Form = class(TForm)
    OK_Button: TButton;
    Hist: THistory;
    CopyFile: TCopyFile;
    procedure FormShow(Sender: TObject);
    procedure OK_ButtonClick(Sender: TObject);
  private
    { Private declarations }
    fFileName:string;
  public
    { Public declarations }
    procedure Execute;

    property filename:string
    read fFilename
    write fFilename;
  end;

var
  LogisticsBreakdown_Form: TLogisticsBreakdown_Form;

implementation

uses DataModule;

const
  InTransit     = 'INTRANSIT';
  ArrivedNUMMI  = 'ARRIVEDNUMMI';
  ArrivedMAN    = 'ARRIVEDMANUF';
  ArrivedUnion  = 'ARRIVEDUNION';
  ArrivedCONS   = 'ARRIVEDCONS';
  Terminated    = 'TERMINATED';

  Renban        = 1;
  Equipment     = 9;
  DT            = 20;
  Status        = 28;
  Lot           = 40;

  Renbanl       = 8;
  Equipmentl    = 11;
  DTl           = 8;
  Statusl       = 12;
  Lotl          = 5;

  {$R *.dfm}
procedure TLogisticsBreakdown_Form.Execute;
begin
  ShowModal;
end;

procedure TLogisticsBreakdown_Form.FormShow(Sender: TObject);
var
  fcf:Textfile;
  fcl:string;
begin
  Hist.Append('Start File Process');
  Data_Module.LogActLog('LOGISTICS','Start processing file,'+fFileNAme);
  try
    AssignFile(fcf, fFileName);
    Reset(fcf);
    //count records
    while not Seekeof(fcf) do
    begin
      Readln(fcf, fcl);
      //
      //Process record
      //
      // check length
      // break out data and update DB
      try
        With Data_Module.Inv_StoredProc do
        Begin
          Close;
          ProcedureName := 'dbo.SELECT_OrderOpenOrderLog;1';
          Parameters.Clear;
          Parameters.AddParameter.Name := '@RenbanNumber';
          Parameters.ParamValues['@RenbanNumber'] := copy(fcl,Renban,Renbanl);
          Open;
          if Data_Module.Inv_Connection.Errors.Count > 0 then
          begin
            ShowMessage('Unable to get Logistics in Order data, '+Data_Module.Inv_Connection.Errors.Item[0].Get_Description);
            Data_Module.LogActLog('ERROR','Unable to get Logistics in Order data, '+Data_Module.Inv_Connection.Errors.Item[0].Get_Description);
            raise EDatabaseError.Create('Unable to get Logistics in Order data');
          end;
          if not IsEmpty then
          begin
            Close;

            if trim(copy(fcl,status,statusl)) = InTransit then
            begin
              ProcedureName := 'dbo.UPDATE_OrderShipping;1';
              Parameters.Clear;
              Parameters.AddParameter.Name := '@RenbanNumber';
              Parameters.ParamValues['@RenbanNumber'] := trim(copy(fcl,Renban,Renbanl));
              Parameters.AddParameter.Name := '@Shipping';
              Parameters.ParamValues['@Shipping'] := trim(copy(fcl,DT,DTl));
              Parameters.AddParameter.Name := '@Trailer';
              Parameters.ParamValues['@Trailer'] := trim(copy(fcl,Equipment,Equipmentl));
              Data_Module.LogActLog('LOGISTICS','Update InTransit, R:'+copy(fcl,Renban,Renbanl)+' DT:'+copy(fcl,DT,DTl));
              Hist.Append('Update InTransit, R:'+copy(fcl,Renban,Renbanl)+' DT:'+copy(fcl,DT,DTl));
            end
            else if trim(copy(fcl,status,statusl)) = ArrivedMAN then
            begin
              ProcedureName := 'dbo.UPDATE_OrderPLANT;1';
              Parameters.Clear;
              Parameters.AddParameter.Name := '@RenbanNumber';
              Parameters.ParamValues['@RenbanNumber'] := trim(copy(fcl,Renban,Renbanl));
              Parameters.AddParameter.Name := '@PLANT';
              Parameters.ParamValues['@PLANT'] := trim(copy(fcl,DT,DTl));
              Parameters.AddParameter.Name := '@Trailer';
              Parameters.ParamValues['@Trailer'] := trim(copy(fcl,Equipment,Equipmentl));
              Parameters.AddParameter.Name := '@Lot';
              if (length(fcl) = 44) then
              begin
                //test:=trim(copy(fcl,lot,lotl+1));
                Parameters.ParamValues['@Lot'] := trim(copy(fcl,Lot,Lotl));
                Data_Module.LogActLog('LOGISTICS','Update Arrived'+Data_Module.fiPlantName.AsString+', R:'+copy(fcl,Renban,Renbanl)+' DT:'+copy(fcl,DT,DTl)+' PK:'+copy(fcl,Lot,Lotl));
                Hist.Append('Update Arrived'+Data_Module.fiPlantName.AsString+', R:'+copy(fcl,Renban,Renbanl)+' DT:'+copy(fcl,DT,DTl)+' PK:'+copy(fcl,Lot,Lotl));
              end
              else if (length(fcl) = 43) then
              begin
                //test:=trim(copy(fcl,lot,lotl+1));
                Parameters.ParamValues['@Lot'] := trim(copy(fcl,Lot,Lotl));
                Data_Module.LogActLog('LOGISTICS','Update Arrived'+Data_Module.fiPlantName.AsString+', R:'+copy(fcl,Renban,Renbanl)+' DT:'+copy(fcl,DT,DTl)+' PK:'+copy(fcl,Lot,Lotl));
                Hist.Append('Update Arrived'+Data_Module.fiPlantName.AsString+', R:'+copy(fcl,Renban,Renbanl)+' DT:'+copy(fcl,DT,DTl)+' PK:'+copy(fcl,Lot,Lotl));
              end
              else
              begin
                Parameters.ParamValues['@Lot'] := '';
                Data_Module.LogActLog('LOGISTICS','Update Arrived'+Data_Module.fiPlantName.AsString+', R:'+copy(fcl,Renban,Renbanl)+' DT:'+copy(fcl,DT,DTl));
                Hist.Append('Update Arrived'+Data_Module.fiPlantName.AsString+', R:'+copy(fcl,Renban,Renbanl)+' DT:'+copy(fcl,DT,DTl));
              end;
            end
            else if trim(copy(fcl,status,statusl)) = ArrivedNUMMI then
            begin
              ProcedureName := 'dbo.UPDATE_OrderPLANT;1';
              Parameters.Clear;
              Parameters.AddParameter.Name := '@RenbanNumber';
              Parameters.ParamValues['@RenbanNumber'] := trim(copy(fcl,Renban,Renbanl));
              Parameters.AddParameter.Name := '@PLANT';
              Parameters.ParamValues['@PLANT'] := trim(copy(fcl,DT,DTl));
              Parameters.AddParameter.Name := '@Trailer';
              Parameters.ParamValues['@Trailer'] := trim(copy(fcl,Equipment,Equipmentl));
              Parameters.AddParameter.Name := '@Lot';
              if (length(fcl) = 44) then
              begin
                //test:=trim(copy(fcl,lot,lotl+1));
                Parameters.ParamValues['@Lot'] := trim(copy(fcl,Lot,Lotl));
                Data_Module.LogActLog('LOGISTICS','Update Arrived'+Data_Module.fiPlantName.AsString+', R:'+copy(fcl,Renban,Renbanl)+' DT:'+copy(fcl,DT,DTl)+' PK:'+copy(fcl,Lot,Lotl));
                Hist.Append('Update Arrived'+Data_Module.fiPlantName.AsString+', R:'+copy(fcl,Renban,Renbanl)+' DT:'+copy(fcl,DT,DTl)+' PK:'+copy(fcl,Lot,Lotl));
              end
              else
              begin
                Parameters.ParamValues['@Lot'] := '';
                Data_Module.LogActLog('LOGISTICS','Update Arrived'+Data_Module.fiPlantName.AsString+', R:'+copy(fcl,Renban,Renbanl)+' DT:'+copy(fcl,DT,DTl));
                Hist.Append('Update Arrived'+Data_Module.fiPlantName.AsString+', R:'+copy(fcl,Renban,Renbanl)+' DT:'+copy(fcl,DT,DTl));
              end;
            end
            else if trim(copy(fcl,status,statusl)) = ArrivedCONS then
            begin
              ProcedureName := 'dbo.UPDATE_OrderWarehouse;1';
              Parameters.Clear;
              Parameters.AddParameter.Name := '@RenbanNumber';
              Parameters.ParamValues['@RenbanNumber'] := trim(copy(fcl,Renban,Renbanl));
              Parameters.AddParameter.Name := '@Warehouse';
              Parameters.ParamValues['@Warehouse'] := trim(copy(fcl,DT,DTl));
              Parameters.AddParameter.Name := '@Trailer';
              Parameters.ParamValues['@Trailer'] := trim(copy(fcl,Equipment,Equipmentl));
              Data_Module.LogActLog('LOGISTICS','Update CONS, R:'+copy(fcl,Renban,Renbanl)+' DT:'+copy(fcl,DT,DTl));
              Hist.Append('Update CONS, R:'+copy(fcl,Renban,Renbanl)+' DT:'+copy(fcl,DT,DTl));
            end
            else if trim(copy(fcl,status,statusl)) = ArrivedUnion then
            begin
              ProcedureName := 'dbo.UPDATE_OrderWarehouse;1';
              Parameters.Clear;
              Parameters.AddParameter.Name := '@RenbanNumber';
              Parameters.ParamValues['@RenbanNumber'] := trim(copy(fcl,Renban,Renbanl));
              Parameters.AddParameter.Name := '@Warehouse';
              Parameters.ParamValues['@Warehouse'] := trim(copy(fcl,DT,DTl));
              Parameters.AddParameter.Name := '@Trailer';
              Parameters.ParamValues['@Trailer'] := trim(copy(fcl,Equipment,Equipmentl));
              Data_Module.LogActLog('LOGISTICS','Update Warehouse, R:'+copy(fcl,Renban,Renbanl)+' DT:'+copy(fcl,DT,DTl));
              Hist.Append('Update Warehouse, R:'+copy(fcl,Renban,Renbanl)+' DT:'+copy(fcl,DT,DTl));
            end
            else if trim(copy(fcl,status,statusl)) = Terminated then
            begin
              ProcedureName := 'dbo.UPDATE_OrderTerminated;1';
              Parameters.Clear;
              Parameters.AddParameter.Name := '@RenbanNumber';
              Parameters.ParamValues['@RenbanNumber'] := trim(copy(fcl,Renban,Renbanl));
              Parameters.AddParameter.Name := '@Terminated';
              Parameters.ParamValues['@Terminated'] := trim(copy(fcl,DT,DTl));
              Parameters.AddParameter.Name := '@Trailer';
              Parameters.ParamValues['@Trailer'] := trim(copy(fcl,Equipment,Equipmentl));
              Data_Module.LogActLog('LOGISTICS','Update Terminated, R:'+copy(fcl,Renban,Renbanl)+' DT:'+copy(fcl,DT,DTl));
              Hist.Append('Update Terminated, R:'+copy(fcl,Renban,Renbanl)+' DT:'+copy(fcl,DT,DTl));
            end;
            ExecProc;
            if Data_Module.Inv_Connection.Errors.Count > 0 then
            begin
              ShowMessage('Unable to get Logistics in Order data, '+Data_Module.Inv_Connection.Errors.Item[0].Get_Description);
              Data_Module.LogActLog('ERROR','Unable to get Logistics in Order data, '+Data_Module.Inv_Connection.Errors.Item[0].Get_Description);
              raise EDatabaseError.Create('Unable to get Logistics in Order data');
            end;
          end
          else
          begin
            Data_Module.LogActLog('LOGISTICS','Renban not in database, R:'+copy(fcl,Renban,Renbanl)+' DT:'+copy(fcl,DT,DTl));
            Hist.Append('Renban not in database, R:'+copy(fcl,Renban,Renbanl)+' DT:'+copy(fcl,DT,DTl));
          end;
        end;
      except
        on e:exception do
        begin
          Hist.Append('Unable to update Order record, '+e.Message);
          Data_Module.LogActLog('ERROR','Unable to update Order record, '+e.Message);
          //
          // Create Error report<<<<
        end;
      end;
    end;
    CloseFile(fcf);

    // Move the file to Archive
    CopyFile.CopyFrom:=filename

  except
    on e:exception do
    begin
      Data_Module.LogActLog('ERROR','Unable to load logistics, '+e.message);
      Hist.Append('Unable to load logistics, '+e.message);
      CloseFile(fcf);
    end;
  end;
  Hist.Append('End File Process');
  Data_Module.LogActLog('LOGISTICS','Finished processing file,'+fFileNAme);
  OK_Button.Visible:=True;
end;

procedure TLogisticsBreakdown_Form.OK_ButtonClick(Sender: TObject);
begin
  Close;
end;

end.
