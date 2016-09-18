unit Write810File;

interface

uses SysUtils, DB, Dialogs, Classes,EDI810Object,DataModule;

type
  T810EDIFile = class(TObject)
  private
    fedi810:T810EDI;
  public
    constructor Create(EDI810:T810EDI); overload;
    procedure Execute();

  end;

implementation

constructor T810EDIFile.Create(EDI810:T810EDI);
begin
  fedi810:=EDI810;
end;

procedure T810EDIFile.Execute();
var
  EDI810:T810EDI;
  i:integer;
  fcf:TextFile;
  line:string;
begin
  // Do Invoice Process
  try
    if Data_Module.fiGenerateEDI.AsBoolean then
    begin
    {
      // Generate True EDI files
      //
      //  Check for Normal 810 Creates (ASN complete)
      //
      Data_Module.EDI810DataSet.Close;
      Data_Module.EDI810DataSet.CommandText:='REPORT_EDI810';
      Data_Module.EDI810DataSet.Parameters.Clear;
      Data_Module.EDI810DataSet.Open;

      if Data_Module.EDI810DataSet.RecordCount > 0 then
      begin
        if MessageDlg('Create INVOICE Files now?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
          ProcessPanel.Visible:=TRUE;
          application.ProcessMessages;
          while not Data_Module.EDI810DataSet.Eof do
          begin
            Data_Module.ASN:=Data_Module.EDI810DataSet.FieldByName('ASNid').AsInteger;
            EDI810:=T810EDI.Create;
            EDI810.EIN:=-1;
            if EDI810.Execute then
            begin
              Data_Module.SiteDataSet.Close;
              Data_Module.SiteDataSet.Open;
              Data_Module.EIN:=Data_Module.SiteDataSet.fieldByName('SiteEIN').AsInteger;
              AssignFile(fcf, Data_Module.fiEDIOut.AsString+'\810'+copy(EDI810.PickUpDate,5,4)+'.txt');
              Rewrite(fcf);

              // Loop through report and save
              for i:=0 to EDI810.EDIRecord.Count-1 do
              begin
                line:=EDI810.EDIRecord[i];
                Writeln(fcf,line);
              end;
              CloseFile(fcf);

              try
                Data_Module.UpdateReportCommand.CommandType:=cmdStoredProc;
                Data_Module.UpdateReportCommand.CommandText:='AD_UpdateEIN';
                Data_Module.UpdateReportCommand.Execute;

                if not Data_Module.InsertINVInfo then
                  ShowMessage('Failed on EIN update after EDI810 create');

              except
                on e:exception do
                begin
                  ShowMessage('Failed on create status after EDI810 create'+e.Message);
                end;
              end;

            end
            else
              ShowMessage('Unable to create EDI810 for ('+EDI810.PickUpDate+')');

            EDI810.Free;
          end;
          ProcessPanel.Visible:=FALSE;
          ShowMessage('Create EDI 810 files complete');
        end;
      end
      else
      begin
        ShowMessage('No Invoice files to create');
      end;
    end
    else
    begin
      // Generate TAI EDI files
      DateSelectDlg:=TDateSelectDlg.Create(self);
      DateSelectDlg.DoSupplier:=FALSE;
      DateSelectDlg.DoLogistics:=FALSE;
      DateSelectDlg.DoPartNumber:=FALSE;
      DateSelectDlg.execute;
      if not DateSelectDlg.Cancel then
      begin
        DailyBuildtotalForm:=TDailyBuildtotalForm.Create(self);
        DailyBuildTotalForm.FormMode:=fmINVOICE;
        DailyBuildTotalForm.FromDate:=DateSelectDlg.FromDate;
        DailyBuildTotalForm.ToDate:=DateSelectDlg.ToDate;
        DailyBuildtotalForm.Execute;
        DailyBuildtotalForm.Free;
      end;}
    end;
  except
    on e:exception do
    begin
      ShowMessage('Unable to create INVOICE, '+e.Message);
    end;
  end;
end;

end.
 