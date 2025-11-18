unit ExcelUtils;
(*************************************************************
Copyright © 2012 Toby Allen (https://github.com/tobya)

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction,
including without limitation the rights to use, copy, modify, merge, publish, distribute, sub-license, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so,
subject to the following conditions:

The above copyright notice, and every other copyright notice found in this software, and all the attributions in every file, and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NON-INFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
****************************************************************)
interface

uses Classes,Sysutils, MainUtils, ResourceUtils,  ActiveX, ComObj, WinINet, Variants,
DynamicFileNameGenerator,

 Excel_TLB_Constants,StrUtils;

type

TExcelXLSConverter = Class(TDocumentConverter)
Private
    ExcelApp : OleVariant;
    FExcelVersion : String;

        olevar_FromPage,  olevar_ToPage   : OleVariant;

    function SingleFileExecuteConversion(fileToConvert, OutputFilename: String;   OutputFileFormat: Integer): TConversionInfo;
    procedure SaveAsPDF(OutputFilename : string) ;
    procedure SaveAsXPS(OutputFilename: string);
    procedure SaveAsCSV(OutputFilename: string);
    procedure ExportWorkSheetasPDF(ws :OleVariant; FileNameGen : TDynamicFileNameGenerator);
    procedure ExportWorkbookasPDF(OutputFileName: String);

    function isWorkSheetEmpty(WorkSheet : OleVariant): boolean;

public
    constructor Create() ;
    function CreateOfficeApp() : boolean;  override;
    function DestroyOfficeApp() : boolean; override;
    function ExecuteConversion(fileToConvert: String; OutputFilename: String; OutputFileFormat : Integer): TConversionInfo; override;
    function AvailableFormats() : TStringList; override;
    function FormatsExtensions(): TStringList; override;
    function OfficeAppVersion() : String; override;
End;

const
  xlTypePDF=50000 ;
  xlpdf=50000 ;
  xlTypeXPS=50001  ;
  xlXPS=50001 ;


implementation



function TExcelXLSConverter.AvailableFormats() : TStringList;

begin
  Formats := TResourceStrings.Create('EXCELFORMATS');
  result := Formats;
end;


{ TWordDocConverter }

constructor TExcelXLSConverter.Create;
begin
  inherited;
  //setup defaults
  InputExtension := '.xls*';
   OfficeAppName := 'Excel';
end;

function TExcelXLSConverter.CreateOfficeApp: boolean;
begin
    ExcelApp :=  CreateOleObject('Excel.Application');
   // ExcelApp.Visible := false;
   result := true;
end;

function TExcelXLSConverter.DestroyOfficeApp: boolean;
begin
  if not VarIsEmpty(ExcelApp) then
  begin
    ExcelApp.Quit;
  end;
  Result := true;
end;



//Useful Links:
//    https://docs.microsoft.com/en-us/office/vba/api/excel.workbooks.open
//    https://docs.microsoft.com/en-us/office/vba/api/excel.workbook.exportasfixedformat

function TExcelXLSConverter.ExecuteConversion(fileToConvert: String; OutputFilename: String; OutputFileFormat : Integer): TConversionInfo;
var
    NonsensePassword :OleVariant;
    FromPage, ToPage : OleVariant;
    activeSheet : OleVariant;
    dynamicoutputDir, dynamicoutputFile, dynamicoutputExt, dynamicOutputFileName, dynamicSheetName : String;
    ExitAction :TExitAction;
    Sheet : integer;
begin
      //Excel is particuarily sensitive to having \\ at end of filename, eg it won't create file.
      //so we remove any double \\
      OutputFilename := stringreplace(OutputFilename, '\\', '\', [rfReplaceAll]);
      log(OutputFilename, verbose);
      ExitAction := aSave;
      Result.InputFile := fileToConvert;
      Result.Successful := false;
      NonsensePassword := 'tfm554!ghAGWRDD';
       LogDebug('in conversion file');
        try
          ExcelApp.Workbooks.Open( FileToConvert,   //FileName					,
                                EmptyParam,    //UpdateLinks					,
                                False,         //ReadOnly					,
                                EmptyParam,      //Format						,
                                NonsensePassword  //Password					,
                                );
                                //WriteResPassword			,
                                //IgnoreReadOnlyRecommended,
                                //Origin						,
                                //Delimiter					,
                                //Editable					,
                                //,
                                //Notify						,
                                //Converter					,
                                //AddToMru					,
                                //Local						,
          Except on E : Exception do
          begin

                    // if Error contains EOleException The password is incorrect.
                    // then it is password protected and should be skipped.
                    if ContainsStr(E.Message, 'The password you supplied is not correct' ) then
                    begin
                       log('Error Occured:' +  E.Message + ' ' + E.ClassName, Verbose);
                       log('SKIPPED - Password Protected:' + fileToConvert, STANDARD);
                       Result.Successful := false;
                       Result.Error := 'SKIPPED - Password Protected:';
                       ExitAction := aExit;

                    end
                    else
                    begin
                      // fallback error log
                      logerror(ConvertErrorText(E.ClassName) + ' ' + ConvertErrorText(E.Message));
                      Result.Successful := false;
                      Result.OutputFile := '';
                      Result.Error := E.Message;
                      ExitAction := aExit;
                    end;
          end;

        end;

     case exitAction of
        aExit: // exit without Closing. (usually because no file is open)
        begin
          Result.Successful := false;
          Result.OutputFile := '';
        end;
        aClose: // Just Close
        begin
            ExcelApp.ActiveWorkbook.Close();
        end;
        aSave: // Go ahead and save
        begin

            Result :=    SingleFileExecuteConversion(fileToConvert, OutputFilename, OutputFileFormat);
        end;
    end;


end;


procedure TExcelXLSConverter.ExportWorkbookasPDF(OutputFileName: String);
var
activeWkBk: OleVariant;
begin

      activeWkBk :=      ExcelApp.ActiveWorkbook;
      ExcelApp.ActiveWorkBook.save;

      ExcelApp.Application.DisplayAlerts := False ;
        activeWkBk.ExportAsFixedFormat(XlFixedFormatType_xlTypePDF,
                                                   OutputFileName,
                                                  EmptyParam,         // Quality
                                                  IncludeDocProps,    // IncludeDocProperties,
                                                  False,              // IgnorePrintAreas,
                                                  olevar_FromPage ,          // From,
                                                  olevar_ToPage,             // To,
                                                  pdfOpenAfterExport, // OpenAfterPublish,  (default false);
                                                  EmptyParam          // FixedFormatExtClassPtr
                                                  ) ;
         fOutputFiles.Add(OutputFileName);

end;

procedure TExcelXLSConverter.ExportWorkSheetasPDF(ws: OleVariant; FileNameGen: TDynamicFileNameGenerator);
begin
        if self.isWorkSheetEmpty(ws) then
        begin
          logInfo('The worksheet "' +  ws.Name + '" is Empty and will not be output', STANDARD);

        end else
        begin


      ExcelApp.Application.DisplayAlerts := False ;
        ws.ExportAsFixedFormat(XlFixedFormatType_xlTypePDF,
                                                   FileNameGen.Generate(ws.Name),
                                                  EmptyParam,         // Quality
                                                  IncludeDocProps,    // IncludeDocProperties,
                                                  False,              // IgnorePrintAreas,
                                                  olevar_FromPage ,          // From,
                                                  olevar_ToPage,             // To,
                                                  pdfOpenAfterExport, // OpenAfterPublish,  (default false);
                                                  EmptyParam          // FixedFormatExtClassPtr
                                                  ) ;
        end;

        fOutputFiles.Add(FileNameGen.Generate(ws.Name));

end;

//Useful Links:
//    https://docs.microsoft.com/en-us/office/vba/api/excel.workbooks.open
//    https://docs.microsoft.com/en-us/office/vba/api/excel.workbook.exportasfixedformat

function TExcelXLSConverter.SingleFileExecuteConversion(fileToConvert: String; OutputFilename: String; OutputFileFormat : Integer): TConversionInfo;
var
    NonsensePassword :OleVariant;
    FromPage, ToPage : OleVariant;
    activeSheet : OleVariant;
    dynamicoutputDir, dynamicoutputFile, dynamicoutputExt, dynamicOutputFileName, dynamicSheetName : String;
    ExitAction :TExitAction;
    Sheet : integer;
begin

             logdebug('SingleFileExecuteConversion',VERBOSE);

            //Unlike Word, in Excel you must call a different function to save a pdf and XPS.
            if OutputFileFormat = xlTypePDF then
            begin

                SaveAsPDF(OutputFilename);


            end
            else if OutputFileFormat = xlTypeXPS then
            begin
                SaveAsXPS(OutputFilename);
            end
            else if OutputFileFormat = xlCSV then
             begin



                // to get sheets
                // Sheets(Array("Sheet4", "Sheet5")) or Sheets(3) or Sheets(Array(1,2))
                 ExcelApp.Application.DisplayAlerts := False ;
                SaveAsCSV(     OutputFilename);

//           ExcelApp.activeWorkbook.SaveAs( OutputFilename, OutputFileFormat);
  //              ExcelApp.ActiveWorkBook.saved := true;
             end
            else
            begin
              //Excel has a tendency to popup alerts so we don't want that.
              ExcelApp.Application.DisplayAlerts := False ;
              ExcelApp.activeWorkbook.SaveAs( OutputFilename, OutputFileFormat);
              ExcelApp.ActiveWorkBook.Saved := true;

            end;

            // Close Excel Sheet.
            Result.Successful := true;
            Result.OutputFile := OutputFilename;
            ExcelApp.ActiveWorkbook.Close();

end;


function TExcelXLSConverter.FormatsExtensions: TStringList;
var
  Extensions : TResourceStrings;

begin
  Extensions := TResourceStrings.Create('XLSEXTENSIONS');

  result := Extensions;
end;

function TExcelXLSConverter.isWorkSheetEmpty(WorkSheet: OleVariant): boolean;
begin

        //  this will tell if a sheet is empty.
      Result := ExcelApp.WorksheetFunction.CountA(WorkSheet.Cells) = 0;

end;

function TExcelXLSConverter.OfficeAppVersion(): String;
begin
  FExcelVersion :=  ReadOfficeAppVersion;
  if FExcelVersion = '' then
  begin
  CreateOfficeApp();
  FExcelVersion := ExcelApp.Version;
  WriteOfficeAppVersion(FExcelVersion);
  end;
  result := FExcelVersion;
end;


procedure TExcelXLSConverter.SaveAsPDF(OutputFilename : string) ;
var

    FromPage, ToPage, SheetList, ExcelSheets : OleVariant;
    Sheet1,Sheet2,Sheet3 , Workbook , SheetsArray: OleVariant;
    activeSheet, WorkSheets, ws, wsName : OleVariant;
    I,j, sheetNumber :integer;
    sheetName : olevariant;
    FileNameGen: TDynamicFileNameGenerator;
begin
 logdebug('Save as pdf',debug);
 ExcelApp.Application.DisplayAlerts := False ;

      if pdfPrintToPage > 0 then
      begin
        logdebug('PrintFromPage: ' + inttostr(pdfPrintFromPage),debug);
        logdebug('PrintToPage: ' + inttostr(pdfPrintToPage),debug);

        olevar_FromPage :=  pdfPrintFromPage;
        olevar_ToPage   :=  pdfPrintToPage;

      end else
      begin
          olevar_FromPage  := EmptyParam;
          olevar_ToPage := EmptyParam;

      end ;


        logdebug('SelectedSheets.Count:' + inttostr( SelectedSheets.Count), debug);
      if SelectedSheets.Count > 0 then
      begin


        logDebug('Selecting sheets' + SelectedSheets.Text,VERBOSE);


        WorkSheets := ExcelApp.Worksheets;

        logDebug('count:' + inttostr(WorkSheets.Count), verbose);
        FileNameGen := TDynamicFileNameGenerator.Create(OutputFilename);

        //for I := 1 to WorkSheets.Count  do
        for j := 0 to SelectedSheets.Count -1 do
        begin

            sheetName := SelectedSheets[j];


            if (TryStrToInt(sheetName,sheetNumber) )then
            begin
              ws := WorkSheets.Item[sheetNumber];
            end else
            begin

               ws := WorkSheets.Item[sheetName];

            end;







                           logDebug('worksheet:' + ws.Name, VERBOSE);

                           ExportWorkSheetasPDF(ws,FileNameGen);


        end;

      end else
      begin

         self.ExportWorkbookasPDF(OutputFileName);
      end;





       ExcelApp.ActiveWorkBook.Saved := True

end;


procedure TExcelXLSConverter.SaveAsXPS(OutputFilename : string) ;
begin

                ExcelApp.Application.DisplayAlerts := False ;
                ExcelApp.activeWorkbook.ExportAsFixedFormat(XlFixedFormatType_xlTypeXPS, OutputFilename  );
                ExcelApp.ActiveWorkBook.save;
end;

procedure TExcelXLSConverter.SaveAsCSV(OutputFilename: string);
var
    FromPage, ToPage : OleVariant;
    activeSheet : OleVariant;
    dynamicoutputDir, dynamicoutputFile, dynamicoutputExt, dynamicOutputFileName, dynamicSheetName : String;
    ExitAction :TExitAction;
    Sheet : integer;
    FileNameGen : TDynamicFileNameGenerator;
begin
                LogDebug('output to csv format');

              //CSV pops up alert. must be hidden for automation
                ExcelApp.Application.DisplayAlerts := False ;

               // if fSelectedSheets.Count > 0 then


                 FileNameGen := TDynamicFileNameGenerator.Create(OutputFilename);

                for Sheet := 1 to ExcelApp.ActiveWorkbook.WorkSheets.Count do
                begin
                 LogDebug('CSV Loop');
                 activeSheet := ExcelApp.ActiveWorkbook.Sheets[Sheet];
                 dynamicSheetName := activeSheet.Name;

                 LogDebug(dynamicSheetName);

                // dynamicSheetName := SafeFileName(dynamicSheetName);
                 dynamicOutputFilename := FileNameGen.Generate(dynamicSheetName);  //dynamicoutputDir + dynamicoutputFile + '_(' + inttostr(Sheet) + dynamicSheetName +  ')' + dynamicoutputExt;

                 LogDebug(dynamicOutputFileName);

                 activeSheet.SaveAs( dynamicoutputFilename, OutputFileFormat);
                end;


                ExcelApp.ActiveWorkBook.save;
end;

end.
