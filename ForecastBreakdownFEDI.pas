//****************************************************************
//
//       Inventory Control
//
//       Copyright (c) 2002-2003 Failproof Manufacturing Systems.
//
//****************************************************************
//
// Change History
//
//  03/10/2003  David Verespey  Create Form
//
unit ForecastBreakdownF;

interface

uses
   Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
   Dialogs, StdCtrls, History, Datamodule,ADOdb,ComObj,strutils,dateutils;

type

   TWeekData = record
     WeekNumber    :integer;
     WeekDate      :string;
     WeekCount     :integer;
   end;

   TEntryRec = record
     Supplier      :string;
     Partnumber    :string;
     KanbanNumber  :string;
     Skip          :boolean;
     Weeks         :array[1..14] of TWeekData;
   end;

   TForecastBreakdown_Form = class(TForm)
     Hist: THistory;
     OKButton: TButton;
     procedure OKButtonClick(Sender: TObject);
   private
     { Private declarations }
     fFilename:string;
     fEntries:array of TEntryRec;
     fSupplierCode:string;
     fFirstWeekNumber:integer;
     fFirstWeekDate:string;
     fFileKind:TFileKind;
     fClosed:boolean;
     function ScanLine(line:string;count:integer):boolean;
     procedure MWReadLn( var fcf:Text; var fcl:string );
     function ScanPartnumber:boolean;
     procedure UpdateForecast;
     procedure UpdateUsage;
     procedure 
DoPartNumberForecast(PN,WeekDate:string;FCCount,WeekNumber:integer);
     function HistoryForecast(partno:string):integer;
   public
     { Public declarations }
     function Execute:boolean;
   published
     property filename:string
     read fFilename
     write fFilename;

     property SupplierCode:string
     read fSupplierCode
     write fSupplierCode;
   end;

const
   SUPPLIER_OFF=1;
   SUPPLIER_SIZE=5;
   PARTNUMBER_OFF=6;
   PARTNUMBER_SIZE=12;
   KANBAN_OFF=18;
   KANBAN_SIZE=4;
   FORECASTW1_OFF=22;
   FORECASTW2_OFF=36;
   FORECASTW3_OFF=50;
   FORECASTW4_OFF=64;
   FORECASTW5_OFF=78;
   FORECASTW6_OFF=92;
   FORECASTW7_OFF=106;
   FORECASTW8_OFF=120;
   FORECASTW9_OFF=134;
   FORECASTW10_OFF=148;
   FORECASTW11_OFF=162;
   FORECASTW12_OFF=176;
   FORECASTW13_OFF=190;
   FORECASTW14_OFF=204;
   WEEKNUMBER_SIZE=2;
   WEEKDATE_SIZE=6;
   WEEKCOUNT_SIZE=6;

   MW_WEEK_START_INDEX=18;
   MW_WEEK_INDEX_LENGTH=9;

var
   ForecastBreakdown_Form: TForecastBreakdown_Form;

implementation

{$R *.dfm}

// Meade Willis uses three Carriage Returns as a new line delimiter
Procedure TForecastBreakdown_Form.MWReadLn( var fcf:Text; var fcl:string );
Var
   Ch : char;
   NewLineCount : Integer;
Begin
   NewLineCount := 0;
   fcl := '';
   Repeat
     Read( fcf, Ch );
     // Do we have a Carriage Return?
     If( Ch = #13 ) Then
       inc(NewLineCount)
     Else
       fcl := fcl + Ch;
   Until (NewLineCount >= 3);
End;

function TForecastBreakdown_Form.Execute:boolean;
var
   fcf,tcf:Textfile;
   fcl,tcl,fn:string;
   fopen:boolean;
   count,counter:integer;
   excel:        variant;
   lastsupplier: string;
   i:            integer;
   mySheet:      variant;
   weekTotal:    integer;
begin
   result:=TRUE;
   count:=0;
   counter:=0;
   lastsupplier:='';
   fclosed:=False;
   OKButton.Visible:=False;
   //open file
   try
     AssignFile(fcf, fFileName);
     Reset(fcf);
     //count records
     while not Seekeof(fcf) do
     begin
       //Readln(fcf, fcl);
       MWReadLn(fcf, fcl);
       INC(count)
     end;
     reset(fcf);
     SetLength(fEntries,count);
     //get all records, check for bad records on the fly
     try
       while not Seekeof(fcf) do
       begin
         //Readln(fcf, fcl);
         MWReadLn(fcf, fcl);
         if ScanLine(fcl,counter) then
         begin
           INC(counter);
         end;
       end;
       //scan for all suppliers, partnumbers and kanbannumber are in 
our database
       Hist.Append(IntToStr(counter)+' total records to process');
       SetLength(fEntries,counter);

       if ScanPartnumber then
       begin
         //clear forecast information from new week onward
         With Data_Module.Inv_StoredProc do
         Begin
           Close;
           ProcedureName := 'dbo.DELETE_ForecastInfoWeekDate;1';
           Parameters.Clear;
           Parameters.AddParameter.Name := '@WeekNumber';
           Parameters.ParamValues['@WeekNumber'] := fFirstWeekNumber;
           Parameters.AddParameter.Name := '@WeekDate';
           Parameters.ParamValues['@WeekDate'] := '20'+fFirstWeekDate;
           ExecProc;
         End;    //With

         //add to forecast table
         UpdateForecast;

         // Update Usage in Size Master
         UpdateUsage;

         //
         //  Produce output files to send to suppliers
         //
         try
           with Data_Module.Inv_DataSet do
           begin
             //  Get week number only process files for the next week out
             Filter:='';
             Filtered:=FALSE;
             Close;
             CommandType := CmdStoredProc;
             CommandText := 'dbo.SELECT_ForecastSupplier;1';
             Parameters.Clear;
             Parameters.AddParameter.Name := '@WeekDate';
             Parameters.ParamValues['@WeekDate'] := 
formatdatetime('yyyymmdd',now);
             Open;

             //
             //  Change to select file output type based on Supplier selection
             //
             //
             // Init both file types;
             excel:=Unassigned;
             fopen:=False;

             while not eof do
             begin
               if fieldbyname('VC_SUPPLIER_CODE').AsString <> lastsupplier then
               begin
                 if not VarIsEmpty(excel) then
                 begin

                   if (fFileKind = fExcel) or (fFileKind = fBoth) then
                   begin
                     //  Create file in directory specified for each supplier
                     //  Use supplier name+WeekDate for filename
                     if 
FileExists(Data_Module.Inv_StoredProc.FieldbyName('Directory').AsString+'\'+ANSIReplaceStr(Data_Module.Inv_StoredProc.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.Inv_StoredProc.fieldbyname('Supplier Code').AsString+'-'+'Forecast') 
then
                       
DeleteFile(Data_Module.Inv_StoredProc.FieldbyName('Directory').AsString+'\'+ANSIReplaceStr(Data_Module.Inv_StoredProc.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.Inv_StoredProc.fieldbyname('Supplier 
Code').AsString+'-'+'Forecast');
                     
excel.ActiveWorkbook.SaveAs(Data_Module.Inv_StoredProc.FieldbyName('Directory').AsString+'/'+ANSIReplaceStr(Data_Module.Inv_StoredProc.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.Inv_StoredProc.fieldbyname('Supplier 
Code').AsString+'-'+'Forecast');
                   end;

                   // Keep copy of forcasts for local archives
                   if( DirectoryExists('c:\MACForecasts') ) then
                       
excel.ActiveWorkbook.SaveAs('c:\MACForecasts\'+ANSIReplaceStr(Data_Module.Inv_StoredProc.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.Inv_StoredProc.fieldbyname('Supplier 
Code').AsString+'-'+'Forecast');

                   excel.Workbooks.Close;
                   excel.Quit;
                   excel:=Unassigned;
                 end;

                 if fopen then
                 begin
                   CloseFile(tcf);
                   fOpen:=FALSE;
                 end;

                 lastsupplier:=fieldbyname('VC_SUPPLIER_CODE').AsString;

                 //
                 // Get FileKind from SUPPLIER table
                 //
                 With Data_Module.Inv_StoredProc do
                 Begin
                   Close;
                   ProcedureName := 'dbo.SELECT_SupplierInfo;1';
                   Parameters.Clear;
                   Parameters.AddParameter.Name := '@SupCode';
                   Parameters.ParamValues['@SupCode'] := lastsupplier;
                   Open;

                   if FieldByName('Output File Type').AsString = 'TEXT' then
                     fFileKind := fText
                   else if FieldByName('Output File Type').AsString = 
'EXCEL' then
                     fFileKind := fExcel
                   else
                     fFileKind := fBoth;

                 End;    //With

                 // Always create excel forcast file for archiving 
(KAM 05 2004)
                 //if (fFileKind = fExcel) or (fFileKind = fBoth) then
                 //begin
                   excel := createOleObject('Excel.Application');
                   excel.visible := False;
                   excel.DisplayAlerts := False;
                   //excel.workbooks.add;

                   
excel.workbooks.open(ExtractFilePath(application.ExeName)+'ForecastTemplate.xls');
                   mysheet := excel.workSheets[1];
                   Hist.Append('Create excel file for supplier, 
'+lastsupplier);

                   i:=2;
                 //end;
                 if (fFileKind = fText) or (fFileKind = fBoth) then
                 begin
                   
fn:=Data_Module.Inv_StoredProc.FieldbyName('Directory').AsString+'\'+ANSIReplaceStr(Data_Module.Inv_StoredProc.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.Inv_StoredProc.fieldbyname('Supplier 
Code').AsString+'.frc';
                   
AssignFile(tcf,Data_Module.Inv_StoredProc.FieldbyName('Directory').AsString+'\'+ANSIReplaceStr(Data_Module.Inv_StoredProc.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.Inv_StoredProc.fieldbyname('Supplier 
Code').AsString+'.frc');
                   Rewrite(tcf);
                   fopen:=true;
                   Hist.Append('Create text file for supplier, '+lastsupplier);
                 end;
               end;

               if (fFileKind = fExcel) or (fFileKind = fBoth) then
               begin
                 weekTotal := fieldbyname('IN_QTY1').AsInteger +
                 fieldbyname('IN_QTY2').AsInteger +
                 fieldbyname('IN_QTY3').AsInteger +
                 fieldbyname('IN_QTY4').AsInteger +
                 fieldbyname('IN_QTY5').AsInteger +
                 fieldbyname('IN_QTY6').AsInteger +
                 fieldbyname('IN_QTY7').AsInteger;

                 mysheet.Cells[i,1].value  := 
fieldbyname('VC_SIZE_CODE').AsString;
                 mysheet.Cells[i,2].value  := 
fieldbyname('VC_PART_NUMBER').AsString;
                 mysheet.Cells[i,3].value  := 
fieldbyname('VC_WEEK_DATE').AsString;
                 mysheet.Cells[i,4].value  := 
fieldbyname('IN_WEEK_NUMBER').AsString;
                 mysheet.Cells[i,5].value  := fieldbyname('IN_QTY1').AsString;
                 mysheet.Cells[i,6].value  := fieldbyname('IN_QTY2').AsString;
                 mysheet.Cells[i,7].value  := fieldbyname('IN_QTY3').AsString;
                 mysheet.Cells[i,8].value  := fieldbyname('IN_QTY4').AsString;
                 mysheet.Cells[i,9].value  := fieldbyname('IN_QTY5').AsString;
                 mysheet.Cells[i,10].value := fieldbyname('IN_QTY6').AsString;
                 mysheet.Cells[i,11].value := fieldbyname('IN_QTY7').AsString;
                 mysheet.Cells[i,12].value := IntToStr(weekTotal);

                 INC(i);
               end;
               if (fFileKind = fText) or (fFileKind = fBoth) then
               begin
                 tcl:='';

                 tcl:=fieldbyname('VC_SUPPLIER_CODE').AsString;
                 tcl:=tcl+fieldbyname('VC_PART_NUMBER').AsString;
                 tcl:=tcl+fieldbyname('VC_WEEK_DATE').AsString;
                 
tcl:=tcl+Format('%.2d',[fieldbyname('IN_WEEK_NUMBER').AsInteger]);
                 tcl:=tcl+Format('%.5d',[fieldbyname('IN_QTY1').AsInteger]);
                 tcl:=tcl+Format('%.5d',[fieldbyname('IN_QTY2').AsInteger]);
                 tcl:=tcl+Format('%.5d',[fieldbyname('IN_QTY3').AsInteger]);
                 tcl:=tcl+Format('%.5d',[fieldbyname('IN_QTY4').AsInteger]);
                 tcl:=tcl+Format('%.5d',[fieldbyname('IN_QTY5').AsInteger]);
                 tcl:=tcl+Format('%.5d',[fieldbyname('IN_QTY6').AsInteger]);
                 tcl:=tcl+Format('%.5d',[fieldbyname('IN_QTY7').AsInteger]);

                 Writeln(tcf,tcl);
               end;
               next;
             end;

             if not VarIsEmpty(excel) then
             begin
               if (fFileKind = fExcel) or (fFileKind = fBoth) then
               begin

                 //  Create file in directory specified for each supplier
                 //  Use supplier name+WeekDate for filename
                 if 
FileExists(Data_Module.Inv_StoredProc.FieldbyName('Directory').AsString+'/'+ANSIReplaceStr(Data_Module.Inv_StoredProc.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.Inv_StoredProc.fieldbyname('Supplier Code').AsString+'-'+'Forecast') 
then
                   
DeleteFile(Data_Module.Inv_StoredProc.FieldbyName('Directory').AsString+'/'+ANSIReplaceStr(Data_Module.Inv_StoredProc.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.Inv_StoredProc.fieldbyname('Supplier 
Code').AsString+'-'+'Forecast');

                 
excel.ActiveWorkbook.SaveAs(Data_Module.Inv_StoredProc.FieldbyName('Directory').AsString+'/'+ANSIReplaceStr(Data_Module.Inv_StoredProc.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.Inv_StoredProc.fieldbyname('Supplier 
Code').AsString+'-'+'Forecast');
               end;

               // Keep copy of forcasts for local archives
               if( DirectoryExists('c:\MACForecasts') ) then
                   
excel.ActiveWorkbook.SaveAs('c:\MACForecasts\'+ANSIReplaceStr(Data_Module.Inv_StoredProc.fieldbyname('Supplier Name').AsString,'/','')+'-'+Data_Module.Inv_StoredProc.fieldbyname('Supplier 
Code').AsString+'-'+'Forecast');

               excel.Workbooks.Close;
               excel.Quit;
               excel:=Unassigned;
             end;

             if fopen then
             begin
               CloseFile(tcf);
             end;
             //
             // End new file selection code
             //
             //
           end;
         except
           on e:exception do
           begin
             Data_Module.LogActLog('ERROR','Unable to create forecast 
output files, '+e.message);
             Hist.Append('Unable to create forecast output files, '+e.message);
             result:=false;

             if not VarIsEmpty(excel) then
             begin
               excel.Workbooks.Close;
               excel.Quit;
               excel:=Unassigned;
             end;
           end;
         end;
       end
       else
         result:=FALSE;
     except
       on e:exception do
       begin
         Data_Module.LogActLog('ERROR','Unable to load forecast, '+e.message);
         result:=false;
       end;
     end;
   finally
     CloseFile(fcf);
   end;
   OKButton.Visible:=True;
   while not fclosed do
   begin
     application.ProcessMessages;
     sleep(500);
   end;
   //put data into forecast table
end;


function TForecastBreakdown_Form.ScanLine(line:string;count:integer):boolean;
var
   LineData :  TStrings;
   NumWeeks, WeekOffset, i :  Integer;

begin
   result:=TRUE;

   LineData := TStringList.Create;
   Try
     LineData.CommaText := line;

     If( LineData[8] = fSupplierCode ) Then
     Begin
       fEntries[count].Supplier:=LineData[8];
       fEntries[count].Partnumber:=LineData[9];
       fEntries[count].KanbanNumber:=LineData[10];
       fEntries[count].Skip:=False;

       // Determine the number of weeks included with this entry
//      NumWeeks := ( LineData.Count - MW_WEEK_START_INDEX ) div 
MW_WEEK_INDEX_LENGTH - 1;
//  NJK:: Bug
// THIS SHOULD BE IT NO?????      NumWeeks := ( LineData.Count - 
MW_WEEK_START_INDEX ) div MW_WEEK_INDEX_LENGTH;
       NumWeeks := ( LineData.Count - MW_WEEK_START_INDEX ) div 
MW_WEEK_INDEX_LENGTH - 1;


      For i := 1 To NumWeeks Do
       Begin
         WeekOffset := MW_WEEK_START_INDEX + (MW_WEEK_INDEX_LENGTH * (i-1));

         // Get the week number from "year / week" field
         fEntries[count].Weeks[i].WeekNumber :=
           StrToInt( copy( LineData[ WeekOffset + 8 ], 3, 2 ) );

         fEntries[count].Weeks[i].WeekDate := LineData[ WeekOffset + 5 ];

         if( LineData[ WeekOffset ] <> '' ) Then
           fEntries[count].Weeks[i].WeekCount := StrToInt( LineData[ 
WeekOffset ] )
         else
           fEntries[count].Weeks[i].WeekCount:=0;

         If( i = 1 ) Then
         Begin
           fFirstWeekNumber := fEntries[count].Weeks[i].WeekNumber;
           fFirstWeekDate := fEntries[count].Weeks[i].WeekDate;
         End;

       End;
     End
     Else
     Begin
       Result := false;
     End;
   except
     on e:exception do
     begin
       LineData.Free;
       Showmessage('File read error, '+e.Message+', import failed');
       Data_Module.LogActLog('ERROR','File read error, '+e.Message+', 
import failed');
       Raise;
     end;
   end;

   LineData.Free;

end;


// This is the original ScanLine from WQS
{
function TForecastBreakdown_Form.ScanLine(line:string;count:integer):boolean;
begin
   result:=TRUE;
   try
     if copy(line,SUPPLIER_OFF,SUPPLIER_SIZE) = fSupplierCode then
     begin
       fEntries[count].Supplier:=copy(line,SUPPLIER_OFF,SUPPLIER_SIZE);
       fEntries[count].Partnumber:=copy(line,PARTNUMBER_OFF,PARTNUMBER_SIZE);
       fEntries[count].KanbanNumber:=copy(line,KANBAN_OFF,KANBAN_SIZE);
       fEntries[count].Skip:=False;
       
fEntries[count].Weeks[1].WeekNumber:=StrToInt(copy(line,FORECASTW1_OFF,WEEKNUMBER_SIZE));
       
fEntries[count].Weeks[1].WeekDate:=copy(line,FORECASTW1_OFF+WEEKNUMBER_SIZE,WEEKDATE_SIZE);
       fFirstWeekNumber:=StrToInt(copy(line,FORECASTW1_OFF,WEEKNUMBER_SIZE));
       fFirstWeekDate:=copy(line,FORECASTW1_OFF+WEEKNUMBER_SIZE,WEEKDATE_SIZE);
       if 
copy(line,FORECASTW1_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE) <> '    
  -' then
         
fEntries[count].Weeks[1].WeekCount:=StrToInt(copy(line,FORECASTW1_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE))
       else
         fEntries[count].Weeks[1].WeekCount:=0;

       
fEntries[count].Weeks[2].WeekNumber:=StrToInt(copy(line,FORECASTW2_OFF,WEEKNUMBER_SIZE));
       
fEntries[count].Weeks[2].WeekDate:=copy(line,FORECASTW2_OFF+WEEKNUMBER_SIZE,WEEKDATE_SIZE);
       if 
copy(line,FORECASTW2_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE) <> '    
  -' then
         
fEntries[count].Weeks[2].WeekCount:=StrToInt(copy(line,FORECASTW2_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE))
       else
         fEntries[count].Weeks[2].WeekCount:=0;

       
fEntries[count].Weeks[3].WeekNumber:=StrToInt(copy(line,FORECASTW3_OFF,WEEKNUMBER_SIZE));
       
fEntries[count].Weeks[3].WeekDate:=copy(line,FORECASTW3_OFF+WEEKNUMBER_SIZE,WEEKDATE_SIZE);
       if 
copy(line,FORECASTW3_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE) <> '    
  -' then
         
fEntries[count].Weeks[3].WeekCount:=StrToInt(copy(line,FORECASTW3_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE))
       else
         fEntries[count].Weeks[3].WeekCount:=0;

       
fEntries[count].Weeks[4].WeekNumber:=StrToInt(copy(line,FORECASTW4_OFF,WEEKNUMBER_SIZE));
       
fEntries[count].Weeks[4].WeekDate:=copy(line,FORECASTW4_OFF+WEEKNUMBER_SIZE,WEEKDATE_SIZE);
       if 
copy(line,FORECASTW4_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE) <> '    
  -' then
         
fEntries[count].Weeks[4].WeekCount:=StrToInt(copy(line,FORECASTW4_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE))
       else
         fEntries[count].Weeks[4].WeekCount:=0;

       
fEntries[count].Weeks[5].WeekNumber:=StrToInt(copy(line,FORECASTW5_OFF,WEEKNUMBER_SIZE));
       
fEntries[count].Weeks[5].WeekDate:=copy(line,FORECASTW5_OFF+WEEKNUMBER_SIZE,WEEKDATE_SIZE);
       if 
copy(line,FORECASTW5_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE) <> '    
  -' then
         
fEntries[count].Weeks[5].WeekCount:=StrToInt(copy(line,FORECASTW5_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE))
       else
         fEntries[count].Weeks[5].WeekCount:=0;

       
fEntries[count].Weeks[6].WeekNumber:=StrToInt(copy(line,FORECASTW6_OFF,WEEKNUMBER_SIZE));
       
fEntries[count].Weeks[6].WeekDate:=copy(line,FORECASTW6_OFF+WEEKNUMBER_SIZE,WEEKDATE_SIZE);
       if 
copy(line,FORECASTW6_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE) <> '    
  -' then
         
fEntries[count].Weeks[6].WeekCount:=StrToInt(copy(line,FORECASTW6_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE))
       else
         fEntries[count].Weeks[6].WeekCount:=0;

       
fEntries[count].Weeks[7].WeekNumber:=StrToInt(copy(line,FORECASTW7_OFF,WEEKNUMBER_SIZE));
       
fEntries[count].Weeks[7].WeekDate:=copy(line,FORECASTW7_OFF+WEEKNUMBER_SIZE,WEEKDATE_SIZE);
       if 
copy(line,FORECASTW7_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE) <> '    
  -' then
         
fEntries[count].Weeks[7].WeekCount:=StrToInt(copy(line,FORECASTW7_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE))
       else
         fEntries[count].Weeks[7].WeekCount:=0;

       
fEntries[count].Weeks[8].WeekNumber:=StrToInt(copy(line,FORECASTW8_OFF,WEEKNUMBER_SIZE));
       
fEntries[count].Weeks[8].WeekDate:=copy(line,FORECASTW8_OFF+WEEKNUMBER_SIZE,WEEKDATE_SIZE);
       if 
copy(line,FORECASTW8_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE) <> '    
  -' then
         
fEntries[count].Weeks[8].WeekCount:=StrToInt(copy(line,FORECASTW8_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE))
       else
         fEntries[count].Weeks[8].WeekCount:=0;

       
fEntries[count].Weeks[9].WeekNumber:=StrToInt(copy(line,FORECASTW9_OFF,WEEKNUMBER_SIZE));
       
fEntries[count].Weeks[9].WeekDate:=copy(line,FORECASTW9_OFF+WEEKNUMBER_SIZE,WEEKDATE_SIZE);
       if 
copy(line,FORECASTW9_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE) <> '    
  -' then
         
fEntries[count].Weeks[9].WeekCount:=StrToInt(copy(line,FORECASTW9_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE))
       else
         fEntries[count].Weeks[9].WeekCount:=0;

       
fEntries[count].Weeks[10].WeekNumber:=StrToInt(copy(line,FORECASTW10_OFF,WEEKNUMBER_SIZE));
       
fEntries[count].Weeks[10].WeekDate:=copy(line,FORECASTW10_OFF+WEEKNUMBER_SIZE,WEEKDATE_SIZE);
       if 
copy(line,FORECASTW10_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE) <> '   
   -' then
         
fEntries[count].Weeks[10].WeekCount:=StrToInt(copy(line,FORECASTW10_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE))
       else
         fEntries[count].Weeks[10].WeekCount:=0;

       
fEntries[count].Weeks[11].WeekNumber:=StrToInt(copy(line,FORECASTW11_OFF,WEEKNUMBER_SIZE));
       
fEntries[count].Weeks[11].WeekDate:=copy(line,FORECASTW11_OFF+WEEKNUMBER_SIZE,WEEKDATE_SIZE);
       if 
copy(line,FORECASTW11_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE) <> '   
   -' then
         
fEntries[count].Weeks[11].WeekCount:=StrToInt(copy(line,FORECASTW11_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE))
       else
         fEntries[count].Weeks[11].WeekCount:=0;

       
fEntries[count].Weeks[12].WeekNumber:=StrToInt(copy(line,FORECASTW12_OFF,WEEKNUMBER_SIZE));
       
fEntries[count].Weeks[12].WeekDate:=copy(line,FORECASTW12_OFF+WEEKNUMBER_SIZE,WEEKDATE_SIZE);
       if 
copy(line,FORECASTW12_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE) <> '   
   -' then
         
fEntries[count].Weeks[12].WeekCount:=StrToInt(copy(line,FORECASTW12_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE))
       else
         fEntries[count].Weeks[12].WeekCount:=0;

       
fEntries[count].Weeks[13].WeekNumber:=StrToInt(copy(line,FORECASTW13_OFF,WEEKNUMBER_SIZE));
       
fEntries[count].Weeks[13].WeekDate:=copy(line,FORECASTW13_OFF+WEEKNUMBER_SIZE,WEEKDATE_SIZE);
       if 
copy(line,FORECASTW13_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE) <> '   
   -' then
         
fEntries[count].Weeks[13].WeekCount:=StrToInt(copy(line,FORECASTW13_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE))
       else
         fEntries[count].Weeks[13].WeekCount:=0;

       
fEntries[count].Weeks[14].WeekNumber:=StrToInt(copy(line,FORECASTW14_OFF,WEEKNUMBER_SIZE));
       
fEntries[count].Weeks[14].WeekDate:=copy(line,FORECASTW14_OFF+WEEKNUMBER_SIZE,WEEKDATE_SIZE);
       if 
copy(line,FORECASTW14_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE) <> '   
   -' then
         
fEntries[count].Weeks[14].WeekCount:=StrToInt(copy(line,FORECASTW14_OFF+WEEKNUMBER_SIZE+WEEKDATE_SIZE,WEEKCOUNT_SIZE))
       else
         fEntries[count].Weeks[14].WeekCount:=0;

     end
     else
       result:=false;
   except
     on e:exception do
     begin
       Showmessage('File read error, '+e.Message+', import failed');
       Data_Module.LogActLog('ERROR','File read error, '+e.Message+', 
import failed');
       Raise;
     end;
   end;
end;
}

function TForecastBreakdown_Form.ScanPartnumber:boolean;
var
   i,x,z,y,skip:integer;
   excel,mysheet:variant;
   dbmissing:TStringList;
   foundx:boolean;
begin
   result:=true;
   skip:=0;
   dbmissing:=TSTringList.Create;

   Hist.Append('Scan part number list');

   with Data_Module.Inv_DataSet do
   begin
     try
       Close;
       CommandType := CmdStoredProc;
       CommandText := 'dbo.SELECT_ForecastDetail;1';
       Parameters.Clear;
       Parameters.AddParameter.Name := '@AssyCode';
       Parameters.ParamValues['@AssyCode'] := '';
       //Parameters.ParamValues['@AssyCode'] := fEntries[i].PArtNumber;
       Open;
       for i:=0 to High(fEntries) do
       begin
         if not Locate('Assembly part number 
Code',fEntries[i].PArtNumber,[]) then
         begin
           // Partnumber does not exist, skip
           Hist.Append('Part Number not found in DB, '+fEntries[i].PartNumber);
           fEntries[i].Skip:=True;
           INC(skip);
         end;
       end;
       First;
       while not eof do
       begin
         foundx:=False;
         for i:=0 to High(fEntries) do
         begin
           if fEntries[i].Partnumber = fieldbyname('Assembly Part 
Number Code').AsString then
           begin
             foundx:=true;
             break;
           end;
         end;
         if not foundx then
         begin
           dbmissing.Add(fieldbyname('Assembly Part Number Code').AsString);
           Hist.Append('Part Number not found in Forecast, 
'+fieldbyname('Assembly Part Number Code').AsString);
         end;
         next;
       end;
     except
       on e:exception do
       begin
         ShowMessage('Error on INV_FORECAST_DETAIL_INF table access, 
'+e.Message);
         Data_Module.LogActLog('ERROR','FORECAST: Error on 
INV_FORECAST_DETAIL_INF table access, '+e.Message);
         result:=false;
       end;
     end;
   end;

   //
   //  Forecast report
   //
   excel := createOleObject('Excel.Application');
   excel.visible := False;
   excel.DisplayAlerts := False;
   //excel.workbooks.add;
   
excel.workbooks.open(ExtractFilePath(application.ExeName)+'ReportTemplate.xls');
   mysheet := excel.workSheets[1];
   mysheet.cells[1,1].value:='Forecast Part Numbers';
   z:=4;
   for y:=1 to 14 do
   begin
     mysheet.Cells[z-1,Y+1].value := 'Week 
'+IntToStr(fEntries[1].Weeks[y].WeekNumber);
   end;
   for x:=0 to high(fEntries) do
   begin
     mysheet.Cells[z,1].value := fEntries[x].Partnumber;
     for y:=1 to 14 do
     begin
       mysheet.Cells[z,Y+1].value := fEntries[x].Weeks[y].WeekCount;
     end;
     INC(z);
   end;
   
excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\ForecastReport'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
   excel.Workbooks.Close;
   excel.Quit;
   excel:=Unassigned;

   //
   // record in database and not in forecast
   //
   if dbmissing.Count > 0 then
   begin
     if messagedlg('There are '+IntToStr(dbMissing.count)+', in the 
database and not in the forecast. Continue processing?',
                     mtConfirmation, [mbYes, mbNo], 0) = mrNo then
     begin
       result:=False;
       exit;
     end;

     hist.append('Create skipped database report, '+IntToStr(skip)+' records');
     // create error xls form
     excel := createOleObject('Excel.Application');
     excel.visible := False;
     excel.DisplayAlerts := False;
     //excel.workbooks.add;
     
excel.workbooks.open(ExtractFilePath(application.ExeName)+'ReportTemplate.xls');
     mysheet := excel.workSheets[1];
     mysheet.cells[1,1].value:='Forecast Part Numbers not in forecast';
     z:=4;
     for x:=0 to dbmissing.Count-1 do
     begin
       mysheet.Cells[z,1].value := dbmissing[x];
       INC(z);
     end;
     
excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\ForecastDBError'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
     excel.Workbooks.Close;
     excel.Quit;
     excel:=Unassigned;
   end;
   dbmissing.Free;

   //
   // record in forecast and not in database
   //
   if skip <> 0 then
   begin
     if messagedlg('There are '+IntToStr(skip)+', forecast records 
not in the database. Continue processing?',
                     mtConfirmation, [mbYes, mbNo], 0) = mrNo then
     begin
       result:=False;
       exit;
     end;

     hist.append('Create skipped records report, '+IntToStr(skip)+' records');
     // create error xls form
     excel := createOleObject('Excel.Application');
     excel.visible := False;
     excel.DisplayAlerts := False;
     //excel.workbooks.add;
     
excel.workbooks.open(ExtractFilePath(application.ExeName)+'ReportTemplate.xls');
     mysheet := excel.workSheets[1];
     mysheet.cells[1,1].value:='Forecast Part Numbers not in database';
     z:=4;
     for x:=0 to high(fEntries) do
     begin
       if fEntries[x].Skip then
       begin
         mysheet.Cells[z,1].value := fEntries[x].Partnumber;
         INC(z);
       end;
     end;
     
excel.ActiveWorkbook.SaveAs(Data_Module.fiReportsOutputDir.AsString+'\ForecastRecError'+formatdatetime('yyyymmddhhmmss00',now)+'.xls');
     excel.Workbooks.Close;
     excel.Quit;
     excel:=Unassigned;
   end;

   Hist.append(IntToStr(high(fEntries)-skip)+' forecast records to be added');
end;

procedure TForecastBreakdown_Form.UpdateUsage;
var
   lastsize:string;
   usage:integer;
begin
   Hist.Append('Update usage DB');
   application.ProcessMessages;
   try
     With Data_Module.INV_DataSet do
     Begin
       Close;
       CommandType := CmdStoredProc;
       CommandText := 'dbo.SELECT_SizeUsage;1';
       Parameters.Clear;
       Open;

       lastsize:=fieldbyname('VC_SIZE_CODE').AsString;
       usage:=0;
       while not EOF do
       begin
         if (lastsize <> fieldbyname('VC_SIZE_CODE').AsString) then
         begin
           // Update last size
           with Data_Module.Inv_StoredProc do
           begin
             //Get ratio data
             if usage <> 0 then
             begin
               Close;
               ProcedureName := 'dbo.UPDATE_SizeUsage;1';
               Parameters.Clear;
               Parameters.AddParameter.Name := '@SizeCode';
               Parameters.ParamValues['@SizeCode'] := lastsize;
               Parameters.AddParameter.Name := '@Usage';
               Parameters.ParamValues['@Usage'] := usage;
               ExecProc;
             end;
           end;
           // Set next
           lastsize:=fieldbyname('VC_SIZE_CODE').AsString;
           usage:=0;
         end;

         // Get week forecast
         usage:=usage+HistoryForecast(fieldbyname('VC_PART_NUMBER').AsString);

         next;
       end;
     end;
   except
     on e:exception do
     begin
       Hist.append('Failed to update usage in INV_SIZE_MST,'+e.message);
       Data_Module.LogActLog('ERROR','FORECAST: Failed to update 
usage in INV_SIZE_MST,'+e.message);
     end;
   end;
   Hist.Append('Finished update usage DB');
end;

function TForecastBreakdown_Form.HistoryForecast(partno:string):integer;
var
   x,z,total:integer;
begin
   //
   //  Change to 7 days instead of 30
   //
   x:=0;
   total:=0;
   try
     with data_module.INV_Forecast_DataSet do
     begin
       for z:=0 to Data_Module.fiUsageUpdateCompare.AsInteger do
       begin
         Close;
         CommandType := CmdStoredProc;
         CommandText := 'dbo.SELECT_ForecastPartNumberWeek;1';
         Parameters.Clear;
         Parameters.AddParameter.Name := '@WeekNo';
         Parameters.ParamValues['@WeekNo'] := WeekOfTheYear(now+z);
         Parameters.AddParameter.Name := '@DayNo';
         Parameters.ParamValues['@DayNo'] := DayOfTheWeek(now+z);
         Parameters.AddParameter.Name := '@PartNo';
         Parameters.ParamValues['@PartNo'] := partno;
         Open;
         if (not FieldByName('Qty').IsNull) and 
(FieldByName('Qty').Value <> 0) then
         begin
           total:=total+FieldByName('Qty').Value;
           INC(x);
         end;
       end;
     end;
     if x<>0 then
       result:=total div x
     else
       result:=0;
   except
     on e:exception do
     begin
       Data_Module.LogActLog('ERROR','Unable to get usage forecast, 
'+e.Message);
       ShowMessage('Unable to get usage forecast, '+e.Message);
       raise;
     end;
   end;
end;


procedure TForecastBreakdown_Form.UpdateForecast;
var
   i,j:integer;
begin
   Hist.Append('Update forecast DB');
   for i:=0 to High(fEntries) do
   begin
     if (not fEntries[i].Skip) then
     begin
       //insert or update record
       while(Hist.Lines.Count > 100) do begin
         Hist.Lines.Delete(0);
       end;
       Hist.Append(fEntries[i].PartNumber);

       try
         for j:=1 to 14 do
           With Data_Module.Inv_StoredProc do
           Begin
             begin
             Close;
             ProcedureName := 'dbo.INSERTUPDATE_ForecastInfo;1';
             Parameters.Clear;
             Parameters.AddParameter.Name := '@Supplier';
             Parameters.ParamValues['@Supplier'] := fEntries[i].Supplier;
             Parameters.AddParameter.Name := '@PartNumber';
             Parameters.ParamValues['@PartNumber'] := fEntries[i].Partnumber;
             Parameters.AddParameter.Name := '@Kanban';
             Parameters.ParamValues['@Kanban'] := fEntries[i].KanbanNumber;
             Parameters.AddParameter.Name := '@WeekNumber';
             Parameters.ParamValues['@WeekNumber'] := 
fEntries[i].Weeks[j].WeekNumber;
             Parameters.AddParameter.Name := '@WeekDate';
             Parameters.ParamValues['@WeekDate'] := 
'20'+fEntries[i].Weeks[j].WeekDate;
             Parameters.AddParameter.Name := '@Count';
             Parameters.ParamValues['@Count'] := 
fEntries[i].Weeks[j].WeekCount;
             ExecProc;
           End;    //With

           try
             with Data_Module.INV_DataSet do
             begin
               //Get ratio data
               Close;
               CommandType := CmdStoredProc;
               CommandText := 'dbo.SELECT_ForecastDetail;1';
               Parameters.Clear;
               Parameters.AddParameter.Name := '@AssyCode';
               Parameters.ParamValues['@AssyCode'] := fEntries[i].PartNumber;
               Open;


               // scan through the part numbers for this assembly 
code and assign
               if length(FieldByName('Tire Part Number 
Code').AsString) = 12 then
               begin
                 DoPartNumberForecast( FieldByName('Tire Part Number 
Code').AsString,
                                       fEntries[i].Weeks[j].WeekDate,
                                       fEntries[i].Weeks[j].WeekCount,
                                       fEntries[i].Weeks[j].WeekNumber);
               end;

               if length(FieldByName('Wheel Part Number 
Code').AsString) = 12 then
               begin
                 DoPartNumberForecast( FieldByName('Wheel Part Number 
Code').AsString,
                                       fEntries[i].Weeks[j].WeekDate,
                                       fEntries[i].Weeks[j].WeekCount,
                                       fEntries[i].Weeks[j].WeekNumber);
               end;

             end;
           except
             on e:exception do
             begin
               ShowMessage('Error on INV_FORECAST_DETAIL_INF table 
access, '+e.Message);
               Data_Module.LogActLog('ERROR','FORECAST: Error on 
INV_FORECAST_DETAIL_INF table access, '+e.Message);
               break;
             end
           end;

           Data_Module.LogActLog('FORECAST','Insert/Update forecast 
information for part number, '+fEntries[i].Partnumber);
         end;
       except
         on e:exception do
         Begin
           Data_Module.LogActLog('ERROR', 'FAILED Insert/Update 
forecast information for part number, '+fEntries[i].Partnumber + '. 
Err Msg: ' + E.Message + ' Err: ' + E.ClassName, 0);
           exit;
         End;      //Except
       end;
       //process breakdown
     end;
   end;
   Hist.Append('Finished update forecast');

   end;

procedure 
TForecastBreakdown_Form.DoPartNumberForecast(PN,WeekDate:string;FCCount,Weeknumber:integer);
var
   workday: array[1..7] of double;
   dayforecast: array[1..7] of integer;
   line,supplier,size:string;
   leftover,i:integer;
   days, ratiocount:double;
begin
   workday[1]:=1.0;
   workday[2]:=1.0;
   workday[3]:=1.0;
   workday[4]:=1.0;
   workday[5]:=1.0;
   workday[6]:=0.0;
   workday[7]:=0.0;
   days:=5.0;

   try
     // Get PartMaster Information
     with Data_Module.INV_Forecast_DataSet do
     begin
       //
       //  Get assembly line
       //
       Close;
       CommandType := CmdStoredProc;
       CommandText := 'dbo.SELECT_PartsStockInfo;1';
       Parameters.Clear;
       Parameters.AddParameter.Name := '@InvMgmtReport';
       Parameters.ParamValues['@InvMgmtReport'] := 'N';
       Parameters.AddParameter.Name := '@Supplier';
       Parameters.ParamValues['@Supplier'] := '';
       Parameters.AddParameter.Name := '@PartNum';
       Parameters.ParamValues['@PartNum'] := PN;
       Open;
       if FieldByName('Car / Truck').AsString = 'Car' then
         line:='C'
       else
         line:='T';

       supplier:=FieldByName('Supplier Code').AsString;
       size:=FieldByName('Size Code').AsString;
       //
       //  Find any changes to a normal weekly schedule
       //
       Close;
       CommandType := CmdStoredProc;
       CommandText := 'dbo.SELECT_OvertimeHolidayWeek;1';
       Parameters.Clear;
       Parameters.AddParameter.Name := '@WeekNumber';
       Parameters.ParamValues['@WeekNumber'] := WeekNumber;
       Parameters.AddParameter.Name := '@Line';
       Parameters.ParamValues['@Line'] := Line;
       Open;
       if recordCount > 0 then
       begin
         while not eof do
         begin
           if fieldByName('VC_OVERTIME_HOLIDAY').AsString = 'H' then
           begin
             workday[FieldByName('IN_DAY_NUMBER').AsInteger]:=0.0;
           end
           else
           begin
             
workday[FieldByName('IN_DAY_NUMBER').AsInteger]:=FieldByName('dec_load').AsFloat;
           end;
           next;
         end;
       end;
       Close;

       days := 0.0;
       for i := 1 to 7 do
         days := days + workday[i];

       if days > 0.0 then
       begin
         //change workday to be float.
         ratiocount := FCCount / days;
       end
       else
       begin
         ratiocount := 0.0;
       end;
       leftover := 0;
       for i:=1 to 7 do
       begin
         if workday[i] > 0.0 then
         begin
           dayforecast[i] := Trunc(ratiocount * workday[i]);
           leftover := leftover + dayforecast[i];
           // need to set leftover. sum of all dayforecast sub'bed from FCCount
         end
         else
           dayforecast[i]:=0;
       end;
       for i:= 1 to 7 do
       begin
         if workday[i] <> 0 then
         begin
           dayforecast[i] := dayforecast[i] + (FCCount - leftover);
           leftover := FCCount;
           break;
         end;
       end;
     end;

     // Sanity assertion.
     leftover := 0;
     for i:=1 to 7 do
     begin
       leftover := leftover + dayforecast[i];
     end;

     Assert(leftover = FCCount, 'Count is NOT equal to the sub of the 
working days');
     //
     //  Insert/Update record for this partnumber on this week
     //
     With Data_Module.Inv_StoredProc do
     Begin
       Close;
       ProcedureName := 'dbo.INSERTUPDATE_BreakdownForecastInfo;1';
       Parameters.Clear;
       Parameters.AddParameter.Name := '@WeekNumber';
       Parameters.ParamValues['@WeekNumber'] := WeekNumber;
       Parameters.AddParameter.Name := '@WeekDate';
       Parameters.ParamValues['@WeekDate'] := '20'+WeekDate;
       Parameters.AddParameter.Name := '@Supplier';
       Parameters.ParamValues['@Supplier'] := Supplier;
       Parameters.AddParameter.Name := '@PartNumber';
       Parameters.ParamValues['@PartNumber'] := PN;
       Parameters.AddParameter.Name := '@SizeCode';
       Parameters.ParamValues['@SizeCode'] := size;
       Parameters.AddParameter.Name := '@Qty1';
       Parameters.ParamValues['@Qty1'] := dayforecast[1];
       Parameters.AddParameter.Name := '@Qty2';
       Parameters.ParamValues['@Qty2'] := dayforecast[2];
       Parameters.AddParameter.Name := '@Qty3';
       Parameters.ParamValues['@Qty3'] := dayforecast[3];
       Parameters.AddParameter.Name := '@Qty4';
       Parameters.ParamValues['@Qty4'] := dayforecast[4];
       Parameters.AddParameter.Name := '@Qty5';
       Parameters.ParamValues['@Qty5'] := dayforecast[5];
       Parameters.AddParameter.Name := '@Qty6';
       Parameters.ParamValues['@Qty6'] := dayforecast[6];
       Parameters.AddParameter.Name := '@Qty7';
       Parameters.ParamValues['@Qty7'] := dayforecast[7];
       ExecProc;
     End;

   except
     on e:exception do
     begin
       Data_Module.LogActLog('ERROR','Failure on partnumber forecast 
update/insert, '+PN);
       Raise;
     end
   end;
end;

procedure TForecastBreakdown_Form.OKButtonClick(Sender: TObject);
begin
   fclosed:=true;
end;

end.


