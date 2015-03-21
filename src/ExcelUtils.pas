unit ExcelUtils;
(*************************************************************
Copyright © 2012 Toby Allen (http://github.com/tobya)

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction,
including without limitation the rights to use, copy, modify, merge, publish, distribute, sub-license, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so,
subject to the following conditions:

The above copyright notice, and every other copyright notice found in this software, and all the attributions in every file, and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NON-INFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
****************************************************************)
interface

uses Classes,Sysutils, MainUtils, ResourceUtils,  ActiveX, ComObj, WinINet, Variants,  Types;

type

TExcelXLSConverter = Class(TDocumentConverter)
Private
    ExcelApp : OleVariant;

public
    constructor Create() ;
    function CreateOfficeApp() : boolean;  override;
    function DestroyOfficeApp() : boolean; override;
    function ExecuteConversion(fileToConvert: String; OutputFilename: String; OutputFileFormat : Integer): string; override;
    function AvailableFormats() : TStringList; override;
    function FormatsExtensions(): TStringList; override;

End;




implementation



function TExcelXLSConverter.AvailableFormats() : TStringList;
var
  Formats : TStringList;

begin
  Formats := Tstringlist.Create();
  LoadStringListFromResource('EXCELFORMATS',Formats);

  result := Formats;
end;


{ TWordDocConverter }

constructor TExcelXLSConverter.Create;
begin
  inherited;
  //setup defaults
  InputExtension := '.xls';
  FLogFilename := 'XlsTo.Log';
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

function TExcelXLSConverter.ExecuteConversion(fileToConvert: String; OutputFilename: String; OutputFileFormat : Integer): string;
begin
            //Open doc and save in requested format.

            //Excel is particuarily sensitive to having \\ at end of filename, eg it won't create file.
            //so we remove any double \\
            log(OutputFilename, VERBOSE);
            OutputFilename := stringreplace(OutputFilename, '\\', '\', [rfReplaceAll]);
            log(OutputFilename, verbose);
            ExcelApp.Workbooks.Open( FileToConvert);

            //pdf is not an actual standard xls output format so we created our own.
            if OutputFileFormat = 50000 then //pdf
            begin
                ExcelApp.Application.DisplayAlerts := False ;
                //Unlike Word, in Excel you must call a different function to save a pdf. Ensure we export entire workbook.
                ExcelApp.activeWorkbook.ExportAsFixedFormat(0, OutputFilename  );
                ExcelApp.ActiveWorkBook.save;
            end
            else if OutputFileFormat = 50001 then //xps
            begin
                ExcelApp.Application.DisplayAlerts := False ;
                //Unlike Word, in Excel you must call a different function to save a pdf. Ensure we export entire workbook.
                ExcelApp.activeWorkbook.ExportAsFixedFormat(1, OutputFilename  );
                ExcelApp.ActiveWorkBook.save;
            end
            else if OutputFileFormat = 6 then //CSV
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

            end;

            ExcelApp.ActiveWorkbook.Close();
end;


function TExcelXLSConverter.FormatsExtensions: TStringList;
var
  Extensions : TStringList;

begin
  Extensions := Tstringlist.Create();
  LoadStringListFromResource('EXTENSIONS',Extensions);

  result := Extensions;
end;

end.
