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
uses Classes, MainUtils, ResourceUtils,  ActiveX, ComObj, WinINet, Variants, sysutils, Types, StrUtils,Word_TLB_Constants;

type

TWordDocConverter = Class(TDocumentConverter)
Private
    FWordVersion : String;
    WordApp : OleVariant;
    WarnBeforeSavingPrintingSendingMarkup_Origional : Boolean;


public
    Constructor Create();
    function CreateOfficeApp() : boolean;  override;
    function DestroyOfficeApp() : boolean; override;
    function ExecuteConversion(fileToConvert: String; OutputFilename: String; OutputFileFormat : Integer): TConversionInfo; override;
    function AvailableFormats() : TStringList; override;
    function FormatsExtensions(): TStringList; override;
    function WordConstants: TStringList;
    function OfficeAppVersion() : String; override;
    procedure BeforeListConvert(); override;
    Procedure AfterListConvert(); override;
End;


implementation



function TWordDocConverter.AvailableFormats() : TStringList;
begin
  Formats := TResourceStrings.Create('WORDFORMATS');
  result := Formats;
end;

function TWordDocConverter.FormatsExtensions() : TStringList;

begin
  fFormatsExtensions := TResourceStrings.Create('DOCEXTENSIONS');
  result := fFormatsExtensions;
end;

function TWordDocConverter.WordConstants() : TStringList;
  var
    Constants : TResourceStrings;
begin
  Constants := TResourceStrings.Create('WORDCONSTANTS');
  result := Constants;
end;




function TWordDocConverter.OfficeAppVersion(): String;
var
  WdVersion: String;
  decimalPos : integer;
begin

  FWordVersion := ReadOfficeAppVersion();

  if FWordVersion = '' then
  begin
    CreateOfficeApp();
    WdVersion := Wordapp.Version;

    //Get Major version as that is all we are interested in and strtofloat causes errors Issue#31
    decimalPos := pos('.',WdVersion);
    FWordVersion  := LeftStr(WdVersion,decimalPos -1);
    WriteOfficeAppVersion(FWordVersion);

  end;
  result := FWordVersion;
end;

{ TWordDocConverter }

procedure TWordDocConverter.BeforeListConvert;
begin
  inherited;
      // If  Word Options->Trust Center->Privacy Options-> "Warn before printing, saving or sending a file that contains tracked changes or comments"
      // is checked it will pop up a dialog on conversion.  Makes not sense for a commandline util to have this set to true
      // So we turn if off but reset it to origional value after.
      WarnBeforeSavingPrintingSendingMarkup_Origional := WordApp.Options.WarnBeforeSavingPrintingSendingMarkup;

      if WarnBeforeSavingPrintingSendingMarkup_Origional then
      begin
        WordApp.Options.WarnBeforeSavingPrintingSendingMarkup := false;
        LogDebug('[SETTING] Setting WordApp.Options.WarnBeforeSavingPrintingSendingMarkup = false', VERBOSE);
      end;
end;

procedure TWordDocConverter.AfterListConvert;
begin
  inherited;

  // set back to originoal value if required.
  if WarnBeforeSavingPrintingSendingMarkup_Origional then
  begin
    WordApp.Options.WarnBeforeSavingPrintingSendingMarkup := true;
    LogDebug('[SETTING] Restoring WordApp.Options.WarnBeforeSavingPrintingSendingMarkup = true', VERBOSE);
  end;
end;

constructor TWordDocConverter.Create;
begin
  inherited;
  InputExtension := '.doc*';
  LogFilename := 'DocTo.Log';
  OfficeAppName := 'Word';
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

var
  EncodingValue : OleVariant;
  NonsensePassword : OleVariant;

  ExitAction : TExitAction;


begin
        ExitAction := aSave;
        Result.Successful := false;
        Result.InputFile := fileToConvert;
        logInfo('ExecuteConversion:' + fileToConvert, Verbose);

        // disable auto macro
        if (fDisableMacros) then
        begin
          WordApp.WordBasic.DisableAutoMacros ;
        end;

        // Check if document has password as per
        // https://wordmvp.com/FAQs/MacrosVBA/CheckIfPWProtectB4Open.htm
        // Always open with password, if none it will be ignored,
        // if the file has a password set then we can catch error.
        NonsensePassword := 'tfm554!ghAGWRDD';

        try
          // Open doc and save in requested format.
          Wordapp.documents.OpenNoRepairDialog( FileToConvert,  // FileName
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
             logInfo('[SKIPPED] - Document has TOC: ' + fileToConvert , STANDARD);
             Result.Successful := false;
             Result.Error := '[SKIPPED] - Document has Table of Contents.';
             ExitAction := aClose;
            end;
          end;
        except
        on E: Exception do
        begin
          // if Error contains EOleException The password is incorrect.
          // then it is password protected and should be skipped.
          if ContainsStr(E.Message, 'The password is incorrect' ) then
          begin
             logInfo('[SKIPPED] - Password Protected:' + fileToConvert, STANDARD);
             Result.Successful := false;
             Result.Error := '[SKIPPED] - Password Protected:';
             ExitAction := aExit;
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



        // Encoding can be set.  If a HTML type, then additional values
        // need to be set for Weboptions.encoding
        if Encoding = -1 then
        begin
           EncodingValue := EmptyParam;
        end
        else
        begin
           if (OutputFileFormat = wdFormatHTML)
           or (OutputFileFormat = wdFormatFilteredHTML)
            or (OutputFileFormat = wdFormatWebArchive)
           then
           begin
             LogDebug('Setting WebOptions.Encoding ',Verbose);
             WordApp.ActiveDocument.WebOptions.Encoding := Encoding;

           end;

           EncodingValue := Encoding;
        end;

      case ExitAction of
      aExit :
      begin
        // document wasn't opened, so just exit function.
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
        if (OutputFileFormat = wdFormatPDF) or
            (OutputFileFormat = wdFormatXPS) then
        begin



        // Saveas works for PDF but github issue 79 requestes exporting bookmarks
        // also which requires ExportAsFixedFormat
        // https://docs.microsoft.com/en-us/office/vba/api/word.document.exportasfixedformat
        WordApp.ActiveDocument.ExportAsFixedFormat(
                   OutputFilename,  //   OutputFileName:=
                   OutputfileFormat, //   ExportFormat:=
                   PDFOpenAfterExport, // OpenAfterExport
                   pdfOptimizeFor ,//   OptimizeFor:= _wdExportOptimizeForPrint
                   pdfExportRange,//   Range
                   pdfPrintFromPage,//   From:=1,
                   pdfprintTopage,//   To:=1, _
                   ExportMarkup,//   Item:=
                   IncludeDocProps,//   IncludeDocProps:=True,
                   KeepIRM,//   KeepIRM:=True, _
                   BookmarkSource,//   CreateBookmarks
                   DocStructureTags,//   DocStructureTags:=True, _
                   BitmapMissingFonts,//   BitmapMissingFonts:=True,
                   useISO190051 //   UseISO19005_1:=False
         );
        end else
        begin

              //SaveAs2 was introduced in 2010 V 14 by this list
              //https://stackoverflow.com/a/29077879/6244
              if (strtoint( OfficeAppVersion) < 14) then
              begin
                    logDebug('Version < 14 Using Saveas Function', VERBOSE);


                    if ( OutputFileFormat = wdFormatPDF )then
                    begin
                      LogInfo('This version of Word does not appear to support saving as PDF.  You will need to install a later version.  Word 2010 is the first to support Native saving as PDF.');
                    end;


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
                                                  EncodingValue, //Encoding,
                                                  EmptyParam, //InsertLineBreaks,
                                                  EmptyParam, //AllowSubstitutions,
                                                  EmptyParam, //LineEnding,
                                                  EmptyParam //AddBiDiMarks
                                                  );

              end
              else
              begin
                    logDebug('Version >= 14 Using Saveas2 Function', VERBOSE);
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
                                              EncodingValue,  //Encoding
                                              EmptyParam,  //InsertLineBreaks
                                              EmptyParam,  //AllowSubstitutions
                                              EmptyParam,  //LineEnding
                                              EmptyParam,  //AddBiDiMarks
                                              CompatibilityMode  //CompatibilityMode
                                              );
              end;

          end;
              Result.Successful := true;
              Result.OutputFile := OutputFilename;
              Result.Error := '';
             // loginfo('FileCreated: ' + OutputFilename, STANDARD);
       finally

            // Close the document - do not save changes if doc has changed in any way.
            Wordapp.activedocument.Close(wdDoNotSaveChanges);
       end;
     end;
    end;


end;



end.
