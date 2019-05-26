unit WordUtils;
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
uses Classes, MainUtils, ResourceUtils,  ActiveX, ComObj, WinINet, Variants, sysutils, Types, StrUtils;

type

TWordDocConverter = Class(TDocumentConverter)
Private
    FWordVersion : String;
    WordApp : OleVariant;
public
    Constructor Create();
    function CreateOfficeApp() : boolean;  override;
    function DestroyOfficeApp() : boolean; override;
    function ExecuteConversion(fileToConvert: String; OutputFilename: String; OutputFileFormat : Integer): TConversionInfo; override;
    function AvailableFormats() : TStringList; override;
    function FormatsExtensions(): TStringList; override;
    function OfficeAppVersion() : String; override;
End;


const

wdDoNotSaveChanges    =	 0; //	Do not save pending changes.
wdSaveChanges         =	-1; //	Save pending changes automatically without prompting the user.
wdPromptToSaveChanges	= -2;	//  Prompt the user to save pending changes.




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
var
  WdVersion: String;
  decimalPos : integer;
begin
  if FWordVersion = '' then
  begin
    CreateOfficeApp();
    WdVersion := Wordapp.Version;
    log('WordVersion:' + WdVersion,VERBOSE);

    //Get Major version as that is all we are interested in and strtofloat causes errors Issue#31
    decimalPos := pos('.',WdVersion);
    FWordVersion  := LeftStr(WdVersion,decimalPos -1);
    log('WordVersion Major:' + FWordVersion,VERBOSE);
  end;
  result := FWordVersion;
end;

{ TWordDocConverter }

constructor TWordDocConverter.Create;
begin
  inherited;
  InputExtension := '.doc*';
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

function TWordDocConverter.ExecuteConversion(fileToConvert: String; OutputFilename: String; OutputFileFormat : Integer): TConversionInfo;
Type
  TWordExitAction = (aSave,aClose, aExit);
var
  wdEncoding : OleVariant;
  NonsensePassword : OleVariant;
  WordExitAction : TWordExitAction;

begin
        WordExitAction := aSave;
        Result.Successful := false;
        Result.InputFile := fileToConvert;
        log('ExecuteConversion:' + fileToConvert, Verbose);

        // Check if document has password as per
        // https://wordmvp.com/FAQs/MacrosVBA/CheckIfPWProtectB4Open.htm
        // Always open with password, if none it will be ignored,
        // if the file has a password set then we can catch error.
        NonsensePassword := 'tfm554!ghAGWRDD';

        try
          // Open doc and save in requested format.
          Wordapp.documents.Open( FileToConvert,  // FileName
                                false,          // ConfirmConversions
                                true,            // ReadOnly
                                EmptyParam,    // AddToRecentFiles,
                                NonsensePassword    // PasswordDocument,
                                );

          // For some reason if the document contains a TableofContents, it hangs Word.  In older
          // versions it popped up a dialog.  Until someone can find a work around, the docs will be skipped.
          // Issue  #40  - experimental as it gets some false positives.
          if SkipDocsWithTOC then
          begin
            if Wordapp.ActiveDocument.TablesOfContents.count > 0 then
            begin
             log('SKIPPED - Document has TOC: ' + fileToConvert , STANDARD);
             Result.Error := 'SKIPPED - Document has TOC:';
             WordExitAction := aClose;
            end;
          end;
        except
        on E: Exception do
        begin
          // if ErroR contains EOleException The password is incorrect.
          // then it is password protected and should be skipped.
          if ContainsStr(E.Message, 'The password is incorrect' ) then
          begin
             log('SKIPPED - Password Protected:' + fileToConvert, STANDARD);
             Result.Error := 'SKIPPED - Password Protected:';
             WordExitAction := aExit;
          end
          else
          begin
            // fallback error log
            logerror(ConvertErrorText(E.ClassName) + ' ' + ConvertErrorText(E.Message));
            Result.Successful := false;
            Result.OutputFile := '';
            Result.Error := E.Message;
            Exit();
          end;
        end;


        end;

        if Encoding = -1 then
        begin
           wdEncoding := EmptyParam;
        end
        else
        begin
           wdEncoding := Encoding;
        end;

      case WordExitAction of
      aExit :
      begin
        // document wasnt opened, so just exit function.
        Result.Successful := false;
        Result.OutputFile := '';
        Exit();
      end;
      aClose:
      begin
        WordApp.activeDocument.Close(wdDoNotSaveChanges);
      end;
      aSave:
      begin
        try
            //SaveAs2 was introducted in 2010 V 14 by this list
            //https://stackoverflow.com/a/29077879/6244
            if (strtoint( OfficeAppVersion) < 14) then
            begin
                  log('Version < 14 Using Saveas Function', VERBOSE);
                  Wordapp.activedocument.Saveas(OutputFilename ,
                                                OutputFileFormat,
                                                EmptyParam, //LockComments,
                                                EmptyParam, //Password,
                                                EmptyParam, //AddToRecentFiles,
                                                EmptyParam, //WritePassword,
                                                EmptyParam, //ReadOnlyRecommended,
                                                EmptyParam, //EmbedTrueTypeFonts,
                                                EmptyParam, //SaveNativePictureFormat,
                                                EmptyParam, //SaveFormsData,
                                                EmptyParam, //SaveAsAOCELetter,
                                                wdEncoding, //Encoding,
                                                EmptyParam, //InsertLineBreaks,
                                                EmptyParam, //AllowSubstitutions,
                                                EmptyParam, //LineEnding,
                                                EmptyParam //AddBiDiMarks
                                                );

            end
            else
            begin
                  log('Version >= 14 Using Saveas2 Function', VERBOSE);
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
                                            wdEncoding,  //Encoding
                                            EmptyParam,  //InsertLineBreaks
                                            EmptyParam,  //AllowSubstitutions
                                            EmptyParam,  //LineEnding
                                            EmptyParam,  //AddBiDiMarks
                                            CompatibilityMode  //CompatibilityMode
                                            );
            end;
            Result.Successful := true;
            Result.OutputFile := OutputFilename;
            Result.Error := '';
            log('FileCreated: ' + OutputFilename, STANDARD);
        finally
                // Close the document - do not save changes if doc has changed in any way.
                Wordapp.activedocument.Close(wdDoNotSaveChanges);
        end;
    end;
    end;


end;



end.
