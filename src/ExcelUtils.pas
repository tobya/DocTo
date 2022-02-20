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

uses Classes,Sysutils, MainUtils, ResourceUtils,  ActiveX, ComObj, WinINet, Variants, Excel_TLB_Constants,StrUtils;

type

TExcelXLSConverter = Class(TDocumentConverter)
Private
    ExcelApp : OleVariant;
    FExcelVersion : String;

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
    ExitAction :TExitAction;
begin
      //Excel is particuarily sensitive to having \\ at end of filename, eg it won't create file.
      //so we remove any double \\
      OutputFilename := stringreplace(OutputFilename, '\\', '\', [rfReplaceAll]);
      log(OutputFilename, verbose);
      ExitAction := aSave;
      Result.InputFile := fileToConvert;
      Result.Successful := false;
      NonsensePassword := 'tfm554!ghAGWRDD';
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

            logdebug('PrintFromPage: ' + inttostr(pdfPrintFromPage),debug);
            logdebug('PrintFromPage: ' + inttostr(pdfPrintToPage),debug);
            //Unlike Word, in Excel you must call a different function to save a pdf and XPS.
            if OutputFileFormat = xlTypePDF then
            begin

                if pdfPrintToPage > 0 then
                begin
                  FromPage :=  pdfPrintFromPage;
                  ToPage   :=  pdfPrintToPage;
                end else
                begin
                  FromPage := EmptyParam;
                  ToPage   := EmptyParam;
                end ;

                ExcelApp.Application.DisplayAlerts := False ;
                ExcelApp.activeWorkbook.ExportAsFixedFormat(XlFixedFormatType_xlTypePDF,
                                                            OutputFilename,
                                                            EmptyParam, //Quality
                                                            IncludeDocProps, // IncludeDocProperties,
                                                            False,// IgnorePrintAreas,
                                                            FromPage , // From,
                                                            ToPage, //  To,
                                                            pdfOpenAfterExport, //   OpenAfterPublish,  (default false);
                                                            EmptyParam//    FixedFormatExtClassPtr
                                                            ) ;

                ExcelApp.ActiveWorkBook.save;

            end
            else if OutputFileFormat = xlTypeXPS then
            begin
                ExcelApp.Application.DisplayAlerts := False ;
                ExcelApp.activeWorkbook.ExportAsFixedFormat(XlFixedFormatType_xlTypeXPS, OutputFilename  );
                ExcelApp.ActiveWorkBook.save;
            end
            else if OutputFileFormat = xlCSV then
             begin
              //CSV pops up alert. must be hidden for automation
                ExcelApp.Application.DisplayAlerts := False ;
                ExcelApp.activeWorkbook.SaveAs( OutputFilename, OutputFileFormat);
                ExcelApp.ActiveWorkBook.save;
             end
            else
            begin
              //Excel has a tendency to popup alerts so we don't want that.
              ExcelApp.Application.DisplayAlerts := False ;
              ExcelApp.activeWorkbook.SaveAs( OutputFilename, OutputFileFormat);
              ExcelApp.ActiveWorkBook.Save;

            end;
            Result.Successful := true;
            Result.OutputFile := OutputFilename;
            ExcelApp.ActiveWorkbook.Close();
            end;
    end;
end;


function TExcelXLSConverter.FormatsExtensions: TStringList;
var
  Extensions : TResourceStrings;

begin
  Extensions := TResourceStrings.Create('XLSEXTENSIONS');

  result := Extensions;
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

end.
