//****************************************************************
//
//       Inventory Control
//
//       Copyright (c) 2002-2015 Failproof Manufacturing Systems.
//
//****************************************************************
//
// Change History
//
//  06/25/2015  David Verespey  Create Form
//
unit ForecastCamexreport;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, History, Datamodule,ADOdb,ComObj,strutils,dateutils;


type
  TColNum = 1..256;  // Excel columns only go up to IV
  TColString = string[2];

  TForecastCAMEXReport = class(TObject)
  private
    function IntToXlCol(ColNum: TColNum): TColString;
    //function XlColToInt(const Col: TColString): TColNum;
  public
    function Execute:boolean;
  end;

implementation


function TForecastCAMEXReport.IntToXlCol(ColNum: TColNum): TColString;
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

//function TForecastCAMEXReport.XlColToInt(const Col: TColString): TColNum;
//const
//  ASCIIOffset = Ord('A') - 1;
//var
//  Len: cardinal;
//begin
//  Len := Length(Col);
//  Result := Ord(UpCase(Col[Len])) - ASCIIOffset;
//  if (Len > 1) then
//    Inc(Result, (Ord(UpCase(Col[1])) - ASCIIOffset) *
//MaxLetters);
//end;

function TForecastCAMEXReport.Execute:boolean;
var
  excel:              variant;
  mySheet:            variant;
  lastsupplier,lastsuppliercode,lastsupplierdirectory,lastpart:        string;
  firstpart:          boolean;
  i,weekoffset, weekrow, daterow, rowoffset, startrow,startcol,maxweekoffset:integer;
begin
  result:=false;
  try

    Data_Module.LogActLog('CAMEXFC','Start order file creation');
    excel:=Unassigned;
    lastsupplier:='';
    weekrow:=2;
    daterow:=3;
    startrow:=4;
    startcol:=2;
    weekoffset:=0;
    rowoffset:=0;
    maxweekoffset:=0;
    firstpart:=true;


    //  Get non-ordered orders
    Data_Module.Inv_DataSet.Close;
    Data_Module.Inv_DataSet.Filter:='';
    Data_Module.Inv_DataSet.Filtered:=FALSE;
    Data_Module.Inv_DataSet.CommandType := CmdStoredProc;
    Data_Module.Inv_DataSet.CommandText := 'dbo.REPORT_ForecastCAMEXReport;1';
    Data_Module.Inv_DataSet.Parameters.Clear;
    Data_Module.Inv_DataSet.Parameters.AddParameter.Name := '@WeekDate';
    Data_Module.Inv_DataSet.Parameters.ParamValues['@WeekDate'] := formatdatetime('yyyymmdd',now);
    Data_Module.Inv_DataSet.Open;

    lastpart:=Data_Module.Inv_DataSet.FieldByName('Part Number').AsString;

    if Data_Module.Inv_DataSet.recordcount > 0 then
    begin
      while not Data_Module.Inv_DataSet.Eof do
      begin
        if lastpart <> Data_Module.Inv_DataSet.FieldByName('Part Number').AsString then
        begin
          lastpart:=Data_Module.Inv_DataSet.FieldByName('Part Number').AsString;
          if firstpart then
          begin
            mysheet.Cells[weekrow,startcol+weekoffset+1].value:='Total';
            maxweekoffset:=weekoffset;
            firstpart:=false;
          end;

          //Sum Formula
          mysheet.Range[IntToXlCol(startcol+weekoffset+1)+IntToStr(startrow+rowoffset)].Formula:='=SUM(B'+IntToStr(startrow+rowoffset)+':'+IntToXlCol(startcol+weekoffset)+IntToStr(startrow+rowoffset)+')';

          INC(rowoffset);
          weekoffset:=0;
        end;

        //check for supplier change and save file
        if lastsupplier <> Data_Module.Inv_DataSet.FieldByName('Supplier Name').AsString then
        begin
          if not VarIsEmpty(excel) then
          begin
            // TODO! Add totals macros at bottom of page
            rowoffset:=rowoffset+2;
            mysheet.Cells[startrow+rowoffset,1].value :=  'Total';
            for i:=0 to maxweekoffset+1 do
            begin
              mysheet.Range[IntToXlCol(startcol+i)+IntToStr(startrow+rowoffset)].Formula:='=SUM('+IntToXlCol(startcol+i)+IntToStr(startrow)+':'+IntToXlCol(startcol+i)+IntToStr(startrow+rowoffset-1)+')';
            end;

            //  Create file in directory specified for each supplier
            //  Use supplier name+WeekDate for filename
            try
              if FileExists(lastsupplierdirectory+'\'+ANSIReplaceStr(lastsupplier,'/','')+'-'+lastsuppliercode+'-'+'CFForecast') then
                DeleteFile(lastsupplierdirectory+'\'+ANSIReplaceStr(lastsupplier,'/','')+'-'+lastsuppliercode+'-'+'CFForecast');
              excel.ActiveWorkbook.SaveAs(lastsupplierdirectory+'/'+ANSIReplaceStr(lastsupplier,'/','')+'-'+lastsuppliercode+'-'+'CFForecast');
              Data_Module.LogActLog('CAMEXFC','Create camex forecast file '+lastsupplierdirectory+'/'+ANSIReplaceStr(lastsupplier,'/','')+'-'+lastsuppliercode+'-'+'CFForecast');
            except
              on e:exception do
              begin
                Data_Module.LogActLog('ERROR','Failed on delete and save excel files, '+e.message+', for supplier('+Data_Module.Inv_DataSet.FieldbyName('Directory').AsString+'\'+ANSIReplaceStr(Data_Module.Inv_DataSet.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.Inv_DataSet.fieldbyname('Supplier Code').AsString+'-'+'CFForecast'+')');
              end;
            end;
            excel.Workbooks.Close;
            excel.Quit;
            excel:=Unassigned;
          end;

          lastsupplier :=  Data_Module.Inv_DataSet.FieldByName('Supplier Name').AsString;
          lastsuppliercode := Data_Module.Inv_DataSet.fieldbyname('Supplier Code').AsString;
          lastsupplierdirectory := Data_Module.Inv_DataSet.FieldbyName('Directory').AsString;

          //open new file
          excel := createOleObject('Excel.Application');
          excel.visible := False;
          excel.DisplayAlerts := False;
          excel.workbooks.open(Data_Module.TemplateDir+'ForecastCamexTemplate.xls');
          mysheet := excel.workSheets[1];
          Data_Module.LogActLog('CAMEXFC','Create CAmex excel file for supplier, '+lastsupplier);


          firstpart:=TRUE;
          lastpart:=Data_Module.Inv_DataSet.FieldByName('Part Number').AsString;
          weekoffset:=0;
          rowoffset:=0;
        end;

        if firstpart then
        begin
          // do data and week headers
          mysheet.Cells[weekrow,startcol+weekoffset].value := 'Week '+ Data_Module.Inv_DataSet.FieldByName('IN_WEEK_NUMBER').AsString;
          mysheet.Cells[daterow,startcol+weekoffset].value := Data_Module.Inv_DataSet.FieldByName('VC_WEEK_DATE').AsString;
        end;
        if weekoffset=0 then
          mysheet.Cells[startrow+rowoffset,1].value :=  lastpart;

        mysheet.Cells[startrow+rowoffset,startcol+weekoffset].value :=  Data_Module.Inv_DataSet.FieldByName('Qty').AsString;

        INC(weekoffset);
        Data_Module.Inv_DataSet.Next;
      end;
          if not VarIsEmpty(excel) then
          begin
            // TODO! Add totals macros at bottom of page
            rowoffset:=rowoffset+2;
            mysheet.Cells[startrow+rowoffset,1].value :=  'Total';
            for i:=0 to maxweekoffset do
            begin
              mysheet.Range[IntToXlCol(startcol+i)+IntToStr(startrow+rowoffset)].Formula:='=SUM('+IntToXlCol(startcol+i)+IntToStr(startrow)+':'+IntToXlCol(startcol+i)+IntToStr(startrow+rowoffset-1)+')';
            end;

            //  Create file in directory specified for each supplier
            //  Use supplier name+WeekDate for filename
            try
              if FileExists(lastsupplierdirectory+'\'+ANSIReplaceStr(lastsupplier,'/','')+'-'+lastsuppliercode+'-'+'CFForecast') then
                DeleteFile(lastsupplierdirectory+'\'+ANSIReplaceStr(lastsupplier,'/','')+'-'+lastsuppliercode+'-'+'CFForecast');
              excel.ActiveWorkbook.SaveAs(lastsupplierdirectory+'/'+ANSIReplaceStr(lastsupplier,'/','')+'-'+lastsuppliercode+'-'+'CFForecast');
              Data_Module.LogActLog('CAMEXFC','Create camex forecast file '+lastsupplierdirectory+'/'+ANSIReplaceStr(lastsupplier,'/','')+'-'+lastsuppliercode+'-'+'CFForecast');
            except
              on e:exception do
              begin
                Data_Module.LogActLog('ERROR','Failed on delete and save excel files, '+e.message+', for supplier('+Data_Module.Inv_DataSet.FieldbyName('Directory').AsString+'\'+ANSIReplaceStr(Data_Module.Inv_DataSet.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.Inv_DataSet.fieldbyname('Supplier Code').AsString+'-'+'CFForecast'+')');
              end;
            end;
            excel.Workbooks.Close;
            excel.Quit;
            excel:=Unassigned;
          end;
    end
    else
    begin
      Data_Module.LogActLog('CAMEXFC','No Camex forecast records found');
    end;

    Data_Module.LogActLog('CAMEXFC','Camex forecast file generation complete');
    result:=true;

  except
    on e:exception do
    begin
      Data_Module.LogActLog('ERROR','Unable to create camex forecast files, '+e.message);
      Data_Module.LogActLog('CAMEXFC','Failed on CAMEX forecast create');
      if not VarIsEmpty(excel) then
      begin
        excel.Workbooks.Close;
        excel.Quit;
        excel:=Unassigned;
      end;
    end;
  end;
end;

end.
