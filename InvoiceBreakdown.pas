unit InvoiceBreakdown;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, History,DB;

type
  TInvoiceBreakdown_Form = class(TForm)
    Hist: THistory;
    OK_Button: TButton;
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

const
  Supplier      = 1;
  Renban        = 6;
  FRS           = 14;
  Part          = 21;
  Kanban        = 33;
  Qty           = 38;
  UnitPrice     = 43;
  Total         = 48;
  InvoiceDate   = 56;
  DueDate       = 64;


  Suppliern      = 1;
  Invoicen       = 6;
  Renbann        = 18;
  FRSn           = 26;
  Partn          = 33;
  Kanbann        = 45;
  Qtyn           = 50;
  UnitPricen     = 55;
  Totaln         = 62;
  InvoiceDaten   = 72;
  DueDaten       = 80;



  Supplierl      = 5;
  Invoicel       = 12;
  Renbanl        = 8;
  FRSl           = 7;
  Partl          = 12;
  Kanbanl        = 5;
  Qtyl           = 5;
  UnitPricel     = 5;
  Totall         = 8;
  UnitPrice4l    = 7;
  Total4l        = 10;
  InvoiceDatel   = 8;
  DueDatel       = 8;


var
  InvoiceBreakdown_Form: TInvoiceBreakdown_Form;

implementation

uses DataModule;

{$R *.dfm}

procedure TInvoiceBreakdown_Form.Execute;
begin
  ShowModal;
end;

procedure TInvoiceBreakdown_Form.FormShow(Sender: TObject);
var
  fcf:Textfile;
  fcl,unitpricet,totalt,test:string;
begin
  Hist.Append('Start File Process');
  Data_Module.LogActLog('INVOICE','Start processing file,'+fFileNAme);
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
      if (length(fcl) <> 71) and (length(fcl) <> 87) then
      begin
        Hist.Append('Bad length record expected (71/87) received '+IntToStr(length(fcl)));
        Data_Module.LogActLog('ERROR','Bad length record expected (71/87) received '+IntToStr(length(fcl)));
      end
      else
      begin
        // break out data and update DB
        try
          With Data_Module.Inv_StoredProc do
          Begin
            Close;
            if length(fcl) = 87 then
            begin
              ProcedureName := 'dbo.INSERTUPDATE_Invoice;1';
              Parameters.Clear;
              Parameters.AddParameter.Name := '@Supplier';
              Parameters.ParamValues['@Supplier'] := copy(fcl,Suppliern,Supplierl);
              test:=copy(fcl,Suppliern,Supplierl);
              Parameters.AddParameter.Name := '@PartNumber';
              Parameters.ParamValues['@PartNumber'] := copy(fcl,Partn,Partl);
              test:= copy(fcl,Partn,Partl);
              Parameters.AddParameter.Name := '@FRS';
              Parameters.ParamValues['@FRS'] := copy(fcl,FRSn,FRSl);
              test:= copy(fcl,FRSn,FRSl);
              Parameters.AddParameter.Name := '@Renban';
              Parameters.ParamValues['@Renban'] := copy(fcl,Renbann,Renbanl);
              test:= copy(fcl,Renbann,Renbanl);
              Parameters.AddParameter.Name := '@Kanban';
              Parameters.ParamValues['@Kanban'] := copy(fcl,Kanbann,Kanbanl);
              test:= copy(fcl,Kanbann,Kanbanl);
              Parameters.AddParameter.Name := '@Qty';
              Parameters.ParamValues['@Qty'] := copy(fcl,Qtyn,Qtyl);
              test:= copy(fcl,Qtyn,Qtyl);
              // get money values
              unitpricet:=copy(fcl,UnitPricen,UnitPrice4l);
              unitpricet:=copy(unitpricet,1,3)+'.'+copy(unitpricet,4,4);
              totalt:=copy(fcl,Totaln,Total4l);
              totalt:=copy(totalt,1,6)+'.'+copy(totalt,7,4);

              Parameters.AddParameter.Name := '@UnitPrice';
              Parameters.ParamValues['@UnitPrice'] := StrToFloat(unitpricet);
              Parameters.AddParameter.Name := '@Total';
              Parameters.ParamValues['@Total'] := StrToFloat(totalt);
              Parameters.AddParameter.Name := '@InvoiceDate';
              Parameters.ParamValues['@InvoiceDate'] := copy(fcl,InvoiceDaten,InvoiceDatel);
              test:= copy(fcl,InvoiceDaten,InvoiceDatel);
              Parameters.AddParameter.Name := '@DueDate';
              Parameters.ParamValues['@DueDate'] := copy(fcl,DueDaten,DueDatel);
              test:= copy(fcl,DueDaten,DueDatel);
              Parameters.AddParameter.Name := '@InvoiceNum';
              Parameters.ParamValues['@InvoiceNum'] := copy(fcl,Invoicen,Invoicel);
              test:= copy(fcl,Invoicen,Invoicel);
            end
            else
            begin
              ProcedureName := 'dbo.INSERTUPDATE_Invoice;1';
              Parameters.Clear;
              Parameters.AddParameter.Name := '@Supplier';
              Parameters.ParamValues['@Supplier'] := copy(fcl,Supplier,Supplierl);
              test:=copy(fcl,Supplier,Supplierl);
              Parameters.AddParameter.Name := '@PartNumber';
              Parameters.ParamValues['@PartNumber'] := copy(fcl,Part,Partl);
              test:= copy(fcl,Part,Partl);
              Parameters.AddParameter.Name := '@FRS';
              Parameters.ParamValues['@FRS'] := copy(fcl,FRS,FRSl);
              test:= copy(fcl,FRS,FRSl);
              Parameters.AddParameter.Name := '@Renban';
              Parameters.ParamValues['@Renban'] := copy(fcl,Renban,Renbanl);
              test:= copy(fcl,Renban,Renbanl);
              Parameters.AddParameter.Name := '@Kanban';
              Parameters.ParamValues['@Kanban'] := copy(fcl,Kanban,Kanbanl);
              test:= copy(fcl,Kanban,Kanbanl);
              Parameters.AddParameter.Name := '@Qty';
              Parameters.ParamValues['@Qty'] := copy(fcl,Qty,Qtyl);
              test:= copy(fcl,Qty,Qtyl);
              // get money values
              unitpricet:=copy(fcl,UnitPrice,UnitPricel);
              unitpricet:=copy(unitpricet,1,3)+'.'+copy(unitpricet,4,2);
              totalt:=copy(fcl,Total,Totall);
              totalt:=copy(totalt,1,6)+'.'+copy(totalt,7,2);

              Parameters.AddParameter.Name := '@UnitPrice';
              Parameters.ParamValues['@UnitPrice'] := StrToFloat(unitpricet);
              Parameters.AddParameter.Name := '@Total';
              Parameters.ParamValues['@Total'] := StrToFloat(totalt);
              Parameters.AddParameter.Name := '@InvoiceDate';
              Parameters.ParamValues['@InvoiceDate'] := copy(fcl,InvoiceDate,InvoiceDatel);
              test:= copy(fcl,InvoiceDate,InvoiceDatel);
              Parameters.AddParameter.Name := '@DueDate';
              Parameters.ParamValues['@DueDate'] := copy(fcl,DueDate,DueDatel);
              test:= copy(fcl,DueDate,DueDatel);
            end;
            ExecProc;
            if Data_Module.Inv_Connection.Errors.Count > 0 then
            begin
              ShowMessage('Unable to get Logistics in Order data, '+Data_Module.Inv_Connection.Errors.Item[0].Get_Description);
              Data_Module.LogActLog('ERROR','Unable to get Logistics in Order data, '+Data_Module.Inv_Connection.Errors.Item[0].Get_Description);
              raise EDatabaseError.Create('Unable to get Logistics in Order data');
            end;
            if length(fcl) = 87 then
            begin
              Data_Module.LogActLog('INVOICE','Add/Update invoice Supplier:'+copy(fcl,Suppliern,Supplierl)+', Part:'+copy(fcl,Partn,Partl)+', FRS:'+copy(fcl,FRSn,FRSl)+', Renban:'+copy(fcl,renbann,renbanl));
              Hist.Append('Add/Update invoice Supplier:'+copy(fcl,Suppliern,Supplierl)+', Part:'+copy(fcl,Partn,Partl)+', FRS:'+copy(fcl,FRSn,FRSl)+', Renban:'+copy(fcl,renbann,renbanl));
            end
            else
            begin
              Data_Module.LogActLog('INVOICE','Add/Update invoice Supplier:'+copy(fcl,Supplier,Supplierl)+', Part:'+copy(fcl,Part,Partl)+', FRS:'+copy(fcl,FRS,FRSl)+', Renban:'+copy(fcl,renban,renbanl));
              Hist.Append('Add/Update invoice Supplier:'+copy(fcl,Supplier,Supplierl)+', Part:'+copy(fcl,Part,Partl)+', FRS:'+copy(fcl,FRS,FRSl)+', Renban:'+copy(fcl,renban,renbanl));
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
    end;
    CloseFile(fcf);
  except
    on e:exception do
    begin
      Data_Module.LogActLog('ERROR','Unable to load invoice, '+e.message);
      Hist.Append('Unable to load invoice, '+e.message);
      CloseFile(fcf);
    end;
  end;
  Hist.Append('End File Process');
  Data_Module.LogActLog('INVOICE','Finished processing file,'+fFileNAme);
  OK_Button.Visible:=True;
end;

procedure TInvoiceBreakdown_Form.OK_ButtonClick(Sender: TObject);
begin
  Close;
end;

end.
