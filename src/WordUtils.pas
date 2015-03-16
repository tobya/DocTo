unit WordUtils;
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
uses Classes, MainUtils, ResourceUtils,  ActiveX, ComObj, WinINet, Variants, sysutils, Types;

type

TWordDocConverter = Class(TDocumentConverter)
Private
    WordApp : OleVariant;
protected



public
    Constructor Create();
    function CreateOfficeApp() : boolean;  override;
    function DestroyOfficeApp() : boolean; override;
    function ExecuteConversion(fileToConvert: String; OutputFilename: String; OutputFileFormat : Integer): string; override;
    function AvailableFormats() : TStringList; override;
    function FormatsExtensions(): TStringList; override;
    function OfficeAppVersion() : String; override;
End;





implementation

function TWordDocConverter.AvailableFormats() : TStringList;
var
  Formats : TStringList;

begin
  Formats := Tstringlist.Create();
  LoadStringListFromResource('WORDFORMATS',Formats);

  result := Formats;
end;

function TWordDocConverter.FormatsExtensions() : TStringList;
var
  Extensions : TStringList;

begin
  Extensions := Tstringlist.Create();
  LoadStringListFromResource('EXTENSIONS',Extensions);

  result := Extensions;
end;





function TWordDocConverter.OfficeAppVersion: String;
begin
  CreateOfficeApp();
  result := Wordapp.Version;
end;

{ TWordDocConverter }

constructor TWordDocConverter.Create;
begin
  inherited;
  Extension := '.doc';
  LogFilename := 'DocTo.Log';
end;

function TWordDocConverter.CreateOfficeApp: boolean;
begin

  if  VarIsEmpty(WordApp) then
  begin
    Wordapp :=  CreateOleObject('Word.Application');
    Wordapp.Visible := false;
  end;
  Result := true;
end;

function TWordDocConverter.DestroyOfficeApp: boolean;
begin
  if not VarIsEmpty(WordApp) then
  begin
    WordApp.Quit();
  end;
  Result := true;
end;

function TWordDocConverter.ExecuteConversion(fileToConvert: String; OutputFilename: String; OutputFileFormat : Integer): string;
begin
        //Open doc and save in requested format.
        Wordapp.documents.Open( FileToConvert, false, true);

    try
        //SaveAs2 was introducted in 2010 V 14 by this list
        //http://stackoverflow.com/a/29077879/6244
        if (strtofloat(OfficeAppVersion) < 14) then
        begin

              Wordapp.activedocument.Saveas(OutputFilename ,OutputFileFormat);
        end
        else
        begin

          Wordapp.activedocument.Saveas2(OutputFilename ,OutputFileFormat,
                                        EmptyParam,  //LockComments
                                        EmptyParam,  //Password
                                        EmptyParam,  //AddToRecentFiles
                                        EmptyParam,  //WritePassword
                                        EmptyParam,  //ReadOnlyRecommended
                                        EmptyParam,  //EmbedTrueTypeFonts
                                        EmptyParam,  //SaveNativePictureFo
                                        EmptyParam,  //SaveFormsData
                                        EmptyParam,  //SaveAsAOCELetter
                                        EmptyParam,  //Encoding
                                        EmptyParam,  //InsertLineBreaks
                                        EmptyParam,  //AllowSubstitutions
                                        EmptyParam,  //LineEnding
                                        EmptyParam,  //AddBiDiMarks
                                        CompatibilityMode  //CompatibilityMode
                                        );
        end;
    finally

            Wordapp.activedocument.Close;
    end;


end;



end.
