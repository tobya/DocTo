unit MainUtils;
(*************************************************************
Copyright © 2012 Toby Allen (https://github.com/tobya)
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction,
including without limitation the rights to use, copy, modify, merge, publish, distribute, sub-license, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so,
subject to the following conditions:
The above copyright notice, and every other copyright notice found in this software, and all the attributions in every file, and this permission notice shall be included in all copies or substantial portions of the Software.
THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NON-INFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Intereting article
https://support.microsoft.com/en-gb/topic/considerations-for-server-side-automation-of-office-48bcfe93-8a89-47f1-0bce-017433ad79e2
****************************************************************)
interface
uses  classes, Windows, sysutils, ActiveX, ComObj, WinINet, Variants, iduri,
      Types,  ResourceUtils,
      PathUtils, ShellAPI, datamodssl, Word_TLB_Constants;

Const
  VERBOSE = 10;
  DEBUG = 9;
  HELP = 8;
  CHATTY = 5;
  STANDARD = 2;
  ERRORS = 1;
  SILENT = 0;

  // APP
  MSWORD = 1;
  MSEXCEL = 2;
  MSPOWERPOINT = 3;
  MSVISIO = 4;

  
  DOCTO_VERSION = '1.9.41';  // dont use 0x - choco needs incrementing versions.
  DOCTO_VERSION_NOTE = ' x64 Release ';
type


  TConsoleLog = class
  public
    procedure Log(Sender: TObject; Log : String);

  end;

  TConversionInfo =  Record
    InputFile , OutputFile : String;
    Successful : Boolean;
    Error : String;

  End;

  TExitAction = (aSave,aClose, aExit);

  TDocumentConverter = class
  private
    FIgnore_MACOSX: boolean;
    FFirstLogEntry: boolean;
    FList_ErrorDocs : boolean;
    FList_ErrorDocs_Seconds : Integer;
    FIgnore_ErrorDocs : boolean;
    FBookMarkSource : integer;
    FWordConstants : TResourceStrings;

    FNetHandle: HINTERNET;
    FOutputIsStdOut: Boolean;
    fExportMarkup: integer;
    FOfficeAppName: String;
    FIncludeDocProps: boolean;
    FKeepIRM: boolean;
    FDocStructureTags: boolean;
    FBitmapMissingFonts: boolean;


    procedure SetCompatibilityMode(const Value: Integer);
    procedure SetIgnore_MACOSX(const Value: boolean);
    procedure SetEncoding(const Value: Integer);
    procedure SetSkipDocsWithTOC(const Value: Boolean);
    procedure HaltWithConfigError(ErrorNo: Integer; Msg: String);

    procedure SetList_ErrorDocs(const Value: Boolean);
    procedure SetList_ErrorDocs_Seconds(const Value: Integer);
    procedure SetIgnore_ErrorDocs(const Value: Boolean);
    procedure SetPDFOpenAfterExport(const Value: Boolean);
    function getIsExcel: Boolean;
    function getIsPP: Boolean;
    function getIsWord: Boolean;
    procedure SetpdfExportRange(const Value: Integer);
    function getWordConstants: TResourceStrings;
    procedure LogMainHelp;
    procedure SetOutputIsStdOut(const Value: Boolean);
    function getIsVisio: Boolean;
    procedure SetIncludeDocProps(const Value: boolean);
    procedure SetKeepIRM(const Value: boolean);
    procedure SetDocStructureTags(const Value: boolean);
    procedure SetBitmapMissingFonts(const Value: boolean);


  protected
    Formats : TStringlist;
    fFormatsExtensions : TStringlist;
    FOutputFileFormatString: String;
    FOutputFileFormat: Integer;
    FOutputLog: Boolean;
    FOutputFile: String;
    FInputFile: String;
    FOutputLogFile: String;
    ConsoleLog : TConsoleLog;
    FVersionString: String;
    FLogLevel : Integer;
    FLogtoFile : Boolean;
    FLogFile : TStringlist;
    FInputFiles : TStringList;
    FLogFilename: String;
    FDoSubDirs: Boolean;
    FInputIsFile: Boolean;
    FInputIsDir: Boolean;
    FOutputExt: string;
    FWebHook : String;
    FInputExtension : String;
    FSkipDocsWithTOC : Boolean;
    fSkipDocsExist : Boolean;
    FCompatibilityMode: Integer;
    FEncoding : Integer;
    FPDFOpenAfterExport : boolean;

    FPDFPrintFromPage : integer;
    FPDFPrintToPage : integer;

    fpdfOptimizeFor : integer;

    FHaltOnWordError: Boolean;
    FRemoveFileOnConvert: boolean;

    FIgnoreErrorDocsFile : TStringList;
    FAppID : Integer;
    FPdfExportRange_Word: Integer;
    FuseISO190051 : Boolean;
    fDisableMacros : Boolean;

    FOutputIsFile: Boolean;
    FOutputIsDir: Boolean;
    procedure SetInputFile(const Value: String);
    procedure SetOutputFile(const Value: String);
    procedure SetOutputFileFormat(const Value: Integer);
    procedure SetOutputFileFormatString(const Value: String);
    procedure SetOutputLog(const Value: Boolean);
    procedure SetOutputLogFile(const Value: String);
    function IsValidFormat(FormatID : Integer): Boolean;

    procedure HaltWithError(ErrorNo:Integer; Msg : String);
    procedure SetLogToFile(const Value: Boolean);
    procedure SetLogFilename(const Value: String);
    procedure ListFiles(const PathName, FileName: string; const SubDir: boolean; outFiles: TStrings);
    procedure SetDoSubDirs(const Value: Boolean);
    procedure SetIsDirInput(const Value: Boolean);
    procedure SetIsFileInput(const Value: Boolean);
    function NewFileNameFromBase(OldBase, NewBase, FileName, NewExt: String): String;
    procedure SetOutputExt(const Value: string);
    function GetUrl(Url: string): String;
    function GetSSLUrl(Url: string): String;
    function URLEncode(Param : String): String;
    procedure SetHaltOnWordError(const Value: Boolean);
    procedure SetRemoveFileOnConvert(const Value: boolean);
    procedure SetIsDirOutput(const Value: Boolean);
    procedure SetIsFileOutput(const Value: Boolean);
    procedure SetLogLevel(const Value: integer);
    property InputIsFile : Boolean read FInputIsFile write SetIsFileInput;
    property InputIsDir : Boolean read FInputIsDir write SetIsDirInput;
    property OutputIsFile : Boolean read FOutputIsFile write SetIsFileOutput;
    property OutputIsDir : Boolean read FOutputIsDir write SetIsDirOutput;
    property OutputIsStdOut : Boolean read FOutputIsStdOut write SetOutputIsStdOut;
    property DoSubDirs : Boolean read FDoSubDirs write SetDoSubDirs;
    property OutputExt : string read FOutputExt write SetOutputExt;
    property LogLevel : integer read FLogLevel write SetLogLevel;
    property RemoveFileOnConvert: boolean read FRemoveFileOnConvert write SetRemoveFileOnConvert;
    property Ignore_MACOSX : boolean   read FIgnore_MACOSX write SetIgnore_MACOSX;
    property List_ErrorDocs : Boolean read FList_ErrorDocs write SetList_ErrorDocs ;
    property List_ErrorDocs_Seconds : Integer read FList_ErrorDocs_Seconds write SetList_ErrorDocs_Seconds ;
    property Ignore_ErrorDocs : Boolean read FIgnore_ErrorDocs write SetIgnore_ErrorDocs;
    property pdfOpenAfterExport: Boolean read FPDFOpenAfterExport write SetpdfOpenAfterExport;
    property pdfPrintFromPage : integer read FpdfPrintFromPage;
    property pdfPrintToPage : integer read FpdfPrintToPage;
    property useISO190051 : boolean read FuseISO190051;
    property pdfOptimizeFor : integer read fpdfOptimizeFor write fpdfOptimizeFor;

    property ExportMarkup : integer read fExportMarkup;
    property IncludeDocProps : boolean read FIncludeDocProps write SetIncludeDocProps;
    property KeepIRM : boolean read FKeepIRM write SetKeepIRM; //  XPS-no-IRM
    property DocStructureTags : boolean read FDocStructureTags write SetDocStructureTags;
    property BitmapMissingFonts : boolean read FBitmapMissingFonts write SetBitmapMissingFonts;


    property WordConstants : TResourceStrings read getWordConstants;
    property OfficeAppName : String read FOfficeAppName write FOfficeAppName;


    // Events
    procedure BeforeListConvert(); virtual;
    Procedure AfterListConvert(); virtual;

    procedure SetExtension(const Value: String); virtual;
    function GetExtension: String;  virtual;
    function OfficeAppVersion() : String; virtual; abstract;

    //Check files and folders
    function AllowDirectory(DirName : String; FullPath : String) : Boolean;
    function AllowFile(FileName : String; Fullpath : String): Boolean;

    // Timing
    procedure CheckDocumentTiming(StartTime, EndTime : cardinal; DocumentPath : String);

    procedure WriteOfficeAppVersion(Version : String);
    function  ReadOfficeAppVersion(): String;

    // Check Should Ignore
    function CheckShouldIgnore(DocumentPath : String): Boolean;


  public

    Constructor Create();
    Destructor Destroy(); override;

    // Load Config
    procedure LoadConfig(Params: TStrings);
    function ChooseConverter(Params: TStrings): integer;

    procedure ConfigLoggingLevel(Params: TStrings);

    function Execute() : string; virtual;
    function ExecuteConversion(fileToConvert: String; OutputFilename: String; OutputFileFormat : Integer): TConversionInfo; virtual; abstract;

    function DestroyOfficeApp() : boolean; virtual; abstract;
    function CreateOfficeApp() : boolean; virtual; abstract;
    function AvailableFormats() : TStringList;  virtual; abstract;
    function FormatsExtensions(): TStringList; virtual; abstract;

    procedure Log(Msg: String; Level  : Integer = ERRORS); overload;

    procedure Log(Msg: String; List:  TStrings; Level: Integer); overload;
        procedure LogInfo(Msg: String; Level  : Integer = ERRORS);
        procedure LogDebug(Msg: String; Level  : Integer = ERRORS);
    procedure LogError(Msg: String);
    function ConvertErrorText(Msg: String) : String;
    function CallWebHook(Params: String) : string;
    FUNCTION AfterConversion(InputFile, OutputFile: String):string;
    Function OnConversionError(InputFile, OutputFile, Error: String):string;


    procedure LogResourceHelp(HelpResName : String);
    procedure LogVersionInfo(ForceReload : boolean = true);



    procedure LogWordFormats();
    procedure LogExcelFormats();
    procedure LogPowerPointFormats();
    procedure ConfigLogHelp(Param, Value: String; AllValues: TStrings);

    function ConfigFileName : String;


    property OutputLog : Boolean    read FOutputLog write SetOutputLog;
    property OutputLogFile : String read FOutputLogFile write SetOutputLogFile;
    Property InputFile : String     read FInputFile write SetInputFile;
    Property OutputFile : String    read FOutputFile write SetOutputFile;
    Property OutputFileFormat : Integer read FOutputFileFormat write SetOutputFileFormat;
    Property OutputFileFormatString : String read FOutputFileFormatString write SetOutputFileFormatString;
    Property LogToFile : Boolean read FLogToFile write SetLogToFile;
    property LogFilename: String read FLogFilename write SetLogFilename;
    Property Version : String read FVersionString;
    property HaltOnWordError : Boolean read FHaltOnWordError write SetHaltOnWordError;
    property SkipDocsWithTOC : Boolean read FSkipDocsWithTOC write SetSkipDocsWithTOC;
    property SkipDocsExist : Boolean read FSkipDocsExist write FSkipDocsExist;
    property InputExtension: String read GetExtension write SetExtension;
    property CompatibilityMode : Integer read FCompatibilityMode write SetCompatibilityMode;
    property Encoding : Integer read FEncoding write SetEncoding;
    property BookMarkSource: Integer read FBookMarkSource;
    property pdfExportRange : Integer read FPdfExportRange_Word write SetPDfExportRange  ;


    property IsWord : Boolean read getIsWord;
    property IsExcel : Boolean read getIsExcel;
    Property IsPowerPoint : Boolean read getIsPP;
    Property IsVisio : Boolean read getIsVisio;
  end;



  function IsNumber(Str: String) : Boolean;


implementation


{ TConsoleLog }

procedure  TConsoleLog.Log(Sender: TObject; Log: String);
begin
  Writeln(Log);
end;

{TDocumentConverter}

function TDocumentConverter.CallWebHook(Params: String):string;
var url : string;
URLResponse : String;
QuestionMarkIndex : Integer;
begin

  try
    if FWebHook > '' then
    begin
      QuestionMarkIndex := pos('?',FWebHook);
      if QuestionMarkIndex = 0  then
      begin
        url := FWebHook + '?'  + Params;
      end
      else if QuestionMarkIndex = length(FWebHook) then  //last character
      begin
        url := FWebHook + Params;
      end
      else
      begin
        url := FWebHook + '&' + Params;
      end;


      URLResponse :=  GetURL(url);

      loginfo('Webhook Called:' + url, CHATTY);
      loginfo('Webhook Response:' + URLResponse, CHATTY);
    end;
  except on E: Exception do
  begin
    logerror(ConvertErrorText( E.ClassName) + ' ' + ConvertErrorText( E.Message));
  end;

  end;
end;


function TDocumentConverter.ConfigFileName: String;
var
 TempDir, ConfigFn, fullfn, value: String;
begin
  ConfigFn := 'docto.config';

  TempDir := GetEnvironmentVariable('TEMP');
  fullfn := TempDir  + '\' + configfn;
  result := fullfn;
end;

procedure TDocumentConverter.ConfigLoggingLevel(Params: TStrings);
var
  iParam : Integer;
  id, pstr, value : String;
begin
LogLevel := STANDARD;
iParam := 0;

While iParam <= Params.Count -1 do
  begin
    pstr := Params[iParam];

    id := UpperCase( pstr);
    if ParamCount -1  > iParam then
    begin
      try
        value := Trim(Params[iParam +1]);
      except on E: Exception do
        HaltWithError(202,E.message );
      end;
    end
    else
    begin
      value := '';
    end;


    if id  = '-L' then
    begin

      if isNumber(value) then
      begin
        LogLevel := strtoint(value);
      end
    end
    else if id  = '-Q' then
    begin
      FLogLevel := 0;
      OutputLog := false;
    end ;

    // Check every parameter as some do not have a value eg. -XL
    inc(iParam,1);
  end;

  Logdebug('Log Level Set To:' + IntToStr(FLogLevel),CHATTY);
end;


procedure TDocumentConverter.ConfigLogHelp(Param, Value: String;
  AllValues: TStrings);

begin

      Value := uppercase(Value);
      if trim(Value) = '' then
      begin

      LogMainhelp();
      end
      else if (Value = 'COMPATIBILITY')
      OR (Value = '-C') then
      BEGIN
        LogResourceHelp('HELPCOMPATIBILITY');
      END
      else if (Value = '-XL') or (Value = 'XL') then
       begin
        LogExcelFormats;
       end
       else if (Value = '-WD') or (Value = 'WD') then
       begin
        LogWordFormats;
       end
       else if (Value = '-PP') or (Value = 'PP') then
      begin
        LogPowerPointFormats;
      end
      else if (value = '-HW') or (value = 'HW') then
      begin
        LogResourceHelp('HELPWEBHOOK');

      end else
      begin
        LogMainHelp();
      end;

      halt(2);
end;




procedure TDocumentConverter.LogMainHelp();
  var
  HelpStrings : TResourceStrings;
begin
      HelpStrings := TResourceStrings.Create('HELP');
      try

        log(format( HelpStrings.Text, [DOCTO_VERSION, OfficeAppVersion()]),Help);
              log('');
        log('FILE FORMATS');
        log('--------------');
        log('To view application specific help and file formats for ');
        log('Word, Excel and Powerpoint use the commands below');
        log('docto -h WD');
        log('docto -h XL');
        log('docto -h PP');
        finally

          HelpStrings.Free;
        end;
end;

// ConvertErrorText removed a lone CR which can overwrite 1 error
// message with another if concatenation see issue #37
// https://github.com/tobya/DocTo/issues/37
function TDocumentConverter.ConvertErrorText(Msg: String): String;
var
  ErrorMessage : String;
begin
  ErrorMessage := StringReplace(Msg,#13,'--',[rfReplaceAll]);
  Result :=  ErrorMessage;
end;

constructor TDocumentConverter.Create;
begin
  ConsoleLog := TConsoleLog.Create();


  //Initial values
  FOutputFileFormatString := '';
  FOutputFileFormat := -1;
  FOutputFile := '';
  FInputFile := '';
  FLogLevel := STANDARD;
  FLogtoFile := false;
  FLogFilename := 'DocTo.Log';
  FRemoveFileOnConvert := false;
  FWebHook := '';
  FOutputExt := '';
  FInputIsFile := false;
  FInputIsDir := false;
  FOutputIsFile := false;
  FOutputIsDir := false;
  FCompatibilityMode := 0;
  FEncoding := -1;
  FDoSubDirs := true;
  FIgnore_MACOSX := true;
  fSkipDocsWithTOC := false;
  fSkipDocsExist :=  false;
  FFirstLogEntry := true;
  FBookMarkSource := 1; //wdExportCreateHeadingBookmarks
  fpdfOptimizeFor := 0; // wdExportOptimizeForPrint
  fPDFOpenAfterExport := false;
  FPdfExportRange_Word := wdExportAllDocument;
  FPDFPrintFromPage := 1;
  FPDFPrintTopage := -1;
  FuseISO190051 := false;
  FIncludeDocProps := true;
  FKeepIRM := true;
  FDocStructureTags := true;
  FBitmapMissingFonts := true;
  FInputFiles := TStringList.Create;
  fDisableMacros := true;


end;

destructor TDocumentConverter.Destroy;
begin
  ConsoleLog.Free;
  if assigned(FLogFile) then
  begin
    FLogFile.SaveToFile(FLogFilename);
    FLogFile.Free;
    FLogFile := nil;
  end;

  if assigned(FIgnoreErrorDocsFile) then
  begin
    FIgnoreErrorDocsFile.Free;
  end;


  FInputFiles.Free;

  if assigned(FNetHandle) then
  begin
    InternetCloseHandle(FNetHandle);
  end;

end;


(*
  Execute conversion on 1 or more files.
*)
function TDocumentConverter.Execute: string;
var

  DoExecute : Boolean;
  i : integer;
  FileToConvert, FileToCreate : String;
  OutputFilePath : String;
  ErrorMessage, EventMsg : String;
  ConversionInfo : TConversionInfo;
  StartTime , EndTime : cardinal;
  FileStream ,OutStream : TStream;



begin

    DoExecute := false;
    if (InputFile > '') and (OutputFile > '') and (OutputFileFormat > -1) then
    begin
      DoExecute := true;
    end;

    if not DoExecute  then HaltWithError(201, 'Input File, Output File and FileFormat must all be specified');

    // Set Output Filename if Dir Provided.
    if (InputIsFile and OutputIsDir) then
    begin
      if OutputExt = '' then
      begin
        OutputExt := '.' + FormatsExtensions.Values[OutputFileFormatString];
        loginfo('Output Extension is ' + outputExt, CHATTY);
      end;

      OutputFile :=  OutputFile  + ChangeFileExt( ExtractFileName(InputFile),OutputExt);
    end;

    // Add file to InputFiles List if only one.
    if FInputFiles.Count = 0 then
    begin
      FInputFiles.Add(FInputFile);
    end;

   try
    CreateOfficeApp();

    BeforeListConvert();

    for i := 0 to FInputFiles.Count -1 do
    begin
      FileToConvert := FInputFiles[i];

      // Check if we should ignore this file as it has previously provided an error.
      if CheckShouldIgnore(FileToConvert) then
      begin
        loginfo('Skipped: File on ignore list. ' + FileToConvert );

        // Jump to end of loop.
        Continue;
      end;

      if InputIsDir then
      begin
        FileToCreate := NewFileNameFromBase(FInputFile ,FOutputFile,FileToConvert, FOutputExt);
      end
      else
      begin
        FileToCreate :=  OutputFile;
      end;




        logdebug('Current Directory: ' + GetCurrentDir,10);

        // Ensure directory exists
        OutputFilePath := ExtractFilePath( FileToCreate);
        if (OutputFilePath = '') then
        begin
          OutputFilePath := GetCurrentDir();
          FileToCreate := OutputFilePath + '\' + FileToCreate;
        end;

        ForceDirectories(OutputFilePath);



      logdebug('Ready to Execute' , VERBOSE);
      if FileExists(FileToCreate) and (SkipDocsExist) then //Not working currently as file doesnt include .ext
      begin
        loginfo('Skipped: FileExists Cannot Create: ' + FileToCreate,Error);

        // Jump to end of loop
        Continue;
      end;
       try

            StartTime := GettickCount();
             logdebug('Executing Conversion ... ',VERBOSE);
            ConversionInfo :=  ExecuteConversion(FileToConvert, FileToCreate, OutputFileFormat);

            if ConversionInfo.Successful then
            begin
              EndTime := getTickCount();
              CheckDocumentTiming(StartTime, EndTime, FileToConvert);
            end;

            if ConversionInfo.Successful then
            begin

              logInfo('File Converted: ' + ConversionInfo.OutputFile);

              // Check if file needs to be deleted.
              if RemoveFileOnConvert then
              begin
                // Check file exists and Delete if requested
                if FileExists(FileToCreate) then
                begin
                  DeleteFile(FileToConvert);
                  Loginfo('Deleted:' + FileToConvert,STANDARD);
                end;
              end;


              // Make a call to webhook if it exists
              EventMsg := AfterConversion(FileToConvert, FileToCreate);

              // This is experimental.  Does not seem to work correctly for anything other than plain text
              // files.  Any binary info seems to be corrupted.  Might be a good project for someone.
              if OutputIsStdOut then
              begin

                FileStream := TFileStream.Create(FileToCreate,fmOpenRead); // This creates the input stream
                OutStream := THandleStream.Create(GetStdHandle(STD_OUTPUT_HANDLE)); // Here goes the output stream
                OutStream.CopyFrom(FileStream, FileStream.Size); // And the copy operation

              end;



            end
            else    // Conversion not successful
            begin
                OnConversionError(ConversionInfo.InputFile, ConversionInfo.OutputFile, ConversionInfo.Error);
            end;
        except
          on E: EOleSysError do
          begin
              // In some instances the error coming from Word has a #13 or Carriage Return which is not
              // interpreted correctly when output to console, replace here.  Github #37
            ErrorMessage := StringReplace(E.Message,#13,'--',[rfReplaceAll]);

            if pos('Invalid class string',E.Message) > 0 then
            begin
              HaltWithError(221,'Word Does not appear to be installed:' + E.ClassName + '  ' + ErrorMessage);
            end
            else
            begin

              CallWebHook('action=error&type='+ FOutputFileFormatString + '&outputfilename=' + URLEncode(FileToCreate)+ '&inputfilename=' + URLEncode(InputFile)
                          + '&error=' + URLEncode(E.ClassName + '  ' + ErrorMessage));

              if (HaltOnWordError) then
              begin

                loginfo('FileToConvert:' + FileToConvert);
                loginfo('OutputFile:' + FileToCreate);
                loginfo('Ext' + inttostr(OutputFileFormat));
                HaltWithError(220,E.ClassName + '  ' + ErrorMessage);
              end
              else
              begin
                loginfo('FileToConvert:' + FileToConvert);
                loginfo('OutputFile:' + FileToCreate);
                loginfo('Ext' + inttostr(OutputFileFormat));
                logerror(E.ClassName + '  ' + ErrorMessage);

              end;
            end;

          end;
          on E: Exception do
          begin
              ErrorMessage := StringReplace(E.Message,#13,'--',[rfReplaceAll]);
              if (HaltOnWordError) then
              begin
                HaltWithError(220,E.ClassName + '  ' + ErrorMessage + ' ' + FileToConvert + ':' + FileToCreate);
              end
              else
              begin
                LogError(E.ClassName + '  ' + ErrorMessage + ' ' + FileToConvert + ':' + FileToCreate);

              end;
          end;
        end;

    end;

    finally
      AfterListConvert();
      DestroyOfficeApp();
    end;


end;




procedure TDocumentConverter.HaltWithError(ErrorNo: Integer; Msg: String);
begin
  LogError(Msg);
  LogError('Exiting with Error Code : ' + inttostr(ErrorNo));
  // Ensure word is quit before halting.
  DestroyOfficeApp();
  Halt(ErrorNo);
end;



function TDocumentConverter.IsValidFormat(FormatID: Integer): Boolean;
var
  i : integer;
begin
  Result := false;
  for i := 0 to Formats.Count -1 do
  begin
    if Formats.ValueFromIndex[i] = inttostr(FormatID) then
    begin
      Result := true;
      break;
    end;
  end;
end;

function TDocumentConverter.ChooseConverter(Params: TStrings) : integer;
var  f , iParam, idx: integer;
pstr : string;
id, value, tmppath : string;

tmpext : String;
valueBool : Boolean;

begin
  // Initialise
  iParam := 0;

  ConfigLoggingLevel(Params);

  OutputLog := true;
  OutputLogFile := '';

  log('Loading ChooseConverter...',VERBOSE);
  log('Parameter Count is ' + inttostr(params.Count), VERBOSE);

  if Params.Count = 0 then
  begin
      log('Parameters Expected: -H for help');
      halt(1);
  end ;

  Result := MSWord;


  While iParam <= Params.Count -1 do
  begin
    pstr := Params[iParam];

    id := UpperCase( pstr);
    if ParamCount -1  > iParam then
    begin
      try
        value := Trim(Params[iParam +1]);
      except on E: Exception do
        HaltWithError(202,E.message);
      end;
    end
    else
    begin
      value := '';
    end;

    // jump to next id + value
    inc(iParam,2);


    if (id = '-XL') or
       (id = '--EXCEL') then
    begin
       Result := MSEXCEL;


    end
    else if (id = '-WD') or
            (id = '--WORD') then
    begin
      Result := MSWORD;

    end
    else if (id = '-PP') or
            (id = '--POWERPOINT') then
    begin
       Result := MSPOWERPOINT;
    end
    else if (id = '-VS') or
            (id = '--VISIO') then
    begin
       Result := MSVISIO;
    end;

    FAppID := Result;

  end;

end;

procedure TDocumentConverter.LoadConfig(Params: TStrings);
var  f , iParam, idx: integer;
pstr : string;
id, value, tmppath : string;
HelpStrings: TResourceStrings;
tmpext : String;
valueBool : Boolean;
  X: Integer;
  Sval : String;

begin
  // Initialise
  iParam := 0;
  Formats := AvailableFormats();
  fFormatsExtensions := FormatsExtensions();
  ConfigLoggingLevel(Params);

  OutputLog := true;
  OutputLogFile := '';

  HaltOnWordError := true;

  loginfo('Loading Configuration...',VERBOSE);
  logdebug('Parameter Count is ' + inttostr(params.Count), VERBOSE);

  if Params.Count = 0 then
  begin
      log('Parameters Expected: -H for help');
      halt(1);
  end ;



  While iParam <= Params.Count -1 do
  begin

    pstr := Params[iParam];

    id := UpperCase( pstr);
    if ParamCount -1  > iParam then
    begin
      try
        value := Trim(Params[iParam +1]);
      except on E: Exception do
        HaltWithError(202,E.message);
      end;
    end
    else
    begin
      value := '';
    end;

    // jump to next id + value
    inc(iParam,2);



if  (id = '-XL') or
        (id = '--EXCEL') or
        (id = '-WD') or
        (id = '--WORD') or
        (id = '-PP') or
        (id = '--POWERPOINT')   or
        (id = '-VS') or
        (id = '--VISIO') then

    begin
      // ignore here as these are checked in ChooseConverter
      dec(iparam);
    end
    else if (id = '-O') or
            (id = '--OUTPUTFILE') then
    begin
      // Before doing anything else expand file name to remove any relative paths.
      FOutputFile :=  ExpandFileName( value);

      tmpext := ExtractFileExt(FOutputFile);

      // if no extension then assume directory, otherwise no way to
      // tell file with no ext from dir.
      if (tmpext = '') then
      begin
        FOutputFile := IncludeTrailingBackslash(FOutputfile);
        OutputIsDir := true;
        OutputIsFile := false;
        ForceDirectories(FOutputFile);
        logInfo('Output directory: ' + FOutputFile,CHATTY);

      end
      else
      begin
        OutputIsFile := true;
        OutputIsDir := false;
        logInfo('Output file: ' + FOutputFile,CHATTY);
      end;


    end

    else if (id = '-OX') or
            (id = '--OUTPUTEXTENSION') then
    begin

     //If the first character isn't . add it.
     if value[1] = '.' then
     begin
        FOutputExt := value;
     end
     else
     begin
       FOutputExt := '.' + value;
     end;

    end
    else if (id = '-F') or
            (id = '--INPUTFILE') then
    begin
      // Before doing anything else expand file name to remove any relative paths.
      FInputFile := ExpandFileName(value);

      logdebug('Input File is: ' + FInputFile,CHATTY);

      tmppath := ExtractFilePath(FInputFile);

        // If we are given a filename with no path, get currentdir and add to file.
        if (tmppath = '') then
        begin
          tmppath := GetCurrentDir();
          tmppath := IncludeTrailingBackslash(tmppath);
          FInputFile := tmppath + FInputFile;
        end;

      if (FileExists(FInputFile) = false) and (DirectoryExists(FInputFile) = false) then
      begin
        HaltWithError(204,'EInput file ' + FInputFile + ' does not exist.');
      end
      else if (FileExists(FInputFile)) then
      begin
        InputIsFile := true;
        InputIsDir := false;
      end
      else if (DirectoryExists(FInputFile)) then
      begin
        InputIsFile := false;
        InputIsDir := true;

        // Create Absolute path from any relative path
        FInputFile := ExpandFileName(FInputFile);
      end;


      // Set Output Directory to Input Directry at this stage. This ensure if no
      // output directory  (-o) is specified, then it will default to same as
      // input dir. If output has been supplied as param it will overwrite later.
      if FOutputFile = '' then
      begin
        FOutputFile :=  IncludeTrailingBackslash(tmppath);
        OutputIsDir := true;
      end;

    end
    else if (id = '--NO-RECURSE') or
            (id = '--NO-SUBDIR') or
            (id = '--NO-SUBDIRS') then
    begin
        FDoSubDirs := false;
        LogInfo('Loading files from directory but not subdirectories',CHATTY);
        dec(iparam);
    end
    else if (id = '--STDOUT') then
    BEGIN
      OutPutIsStdOut := true;
      dec(iparam);
    END
    else if (id = '-FX') or
            (id = '--INPUTFILEEXTENSION') or
            (id = '--INPUTFILTER')
            then
    begin
      InputExtension := value;
    end
    else if ( id  = '-L')
         OR (id = '--LOGLEVEL') then
    begin
      if isNumber(value) then
      begin
        LogLevel := strtoint(value);
        LogInfo('Log Level Set To:' + IntToStr(LogLevel),LogLevel);
      end
    end
    else if (id  = '-Q') or
            (id = '--QUIET') then
    begin

      OutputLog := false;
      // Doesn't require a value
      dec(iParam);
    end
    else if (id = '-T') or (id = '-TF') or
            (id = '--FORMAT') or (id = '--FORCEFORMAT') then
    begin

      if IsNumber(value) then
      begin
        FOutputFileFormat :=  strtoint(value);
        //If not forcing, and the format is invalid by list, then raise error.
        if (not (id = '-TF')) and ( not IsValidFormat(FOutputFileFormat)) then
        begin
          HaltWithConfigError(200, 'File Format ' + value + ' is invalid, please see help. -h.  To force use, use -TF');
        end;
      end
      else  // string format such as 'XLcsv'
      begin
        FOutputFileFormatString := value;

        idx := formats.IndexOfName(FOutputFileFormatString);
        if  idx > -1 then
        begin
          OutputFileFormat := strtoint(formats.Values[OutputFileFormatString]);

        end
        else if idx = -1 then
        begin
          HaltWithConfigError(200,'File Format ' + OutputFileFormatString + ' is an invalid ' + OfficeAppName +  ' file extension , please see help. -h');

        end;
      end;
      logdebug('Type Integer is: ' + inttostr(FOutputFileFormat), VERBOSE);

    end
    else if (id = '-C') or
            (id = '--COMPATIBILITY') then
    begin

      CompatibilityMode := strtoint(value);
    end
    else if (id = '-E') or
            (id = '--ENCODING') then
    begin
      Encoding := strtoint(value);
    end
    else if (id = '-G') or
             (id = '--WRITELOGFILE') then
    begin
       LogToFile := true;
       dec(iParam);
    end
    else if (id = '-GL') or
            (id = '--LOGFILENAME') then
    begin
       FLogFilename := value;
       LogToFile := true;
    end
    else if (id = '-M')  or
              (id = '--IGNOREMACOS') then
    begin
      Ignore_MACOSX := StrToBool( value);
    end
    else if (id = '-N')  or
            (id = '--LISTLONGRUNNING') then
    begin
      List_ErrorDocs := True;
      List_ErrorDocs_Seconds := StrToInt(value);
    end
    else if (id = '-NX')  or
            (id = '--IGNORELONGRUNNINGLIST') then
    begin
      Ignore_ErrorDocs := True;
      dec(iParam);
    end
    else if (id = '--PDF-OPENAFTEREXPORT') then
    begin

      PDFOpenAfterExport := true;
      dec(iParam);
    end
    else if (id = '--USE-ISO190051') or
            (id = '--USE-ISO19005-1') then
    begin
       FuseISO190051 := true;
       dec(iparam);
    end
    else if (id = '-R')
         or (id = '--DELETEFILES') then
    begin

      // -R value must be followed by true or false.
      if TryStrToBool(value, valueBool)  then
      begin
        RemoveFileOnConvert  := valueBool;
      end
      else
      begin
          HaltWithConfigError(200,'If -R is used it must be followed by true or false');
      end;


    end
    else if (id = '--BOOKMARKSOURCE') or
            (id = '--PDF-BOOKMARKSOURCE') then
    begin
         if (WordConstants.Exists(value)) then
         begin
           FBookMarkSource := StrToInt( WordConstants.Values[value]);
           logdebug('Set Bookmark To: ' + InttoStr(FBookmarkSource), Verbose);
         end else
         begin
           HaltWithConfigError(205,'Invalid value for --PDF-BOOKMARKSOURCE :' + value);
         end;

    end
    else if (id = '--PDF-FROMPAGE') then
    begin

      // From value can also set the ExportRange setting.
      // Otherwise it is an integer that set the from page
      if (value = 'wdExportCurrentPage')
      or (value = 'wdExportSelection') then
      begin
        if isWord then
        begin
            FPdfExportRange_Word := WordConstants.ValueAsInt[value];
        end else
        begin HaltWithConfigError(202, value + ' is only valid with Microsoft Word (-WD)'); end;
      end
      else  // value must be an integer representing page number.
      begin
        try
         fpdfPrintFromPage := StrToInt(value);
         FPdfExportRange_Word := wdExportFromTo;
        except
              HaltWithConfigError( 202, '--PDF-FROMPAGE must be an integer');
        end;
      end;

    end
    else if (id = '--PDF-TOPAGE') then
    begin
      try
       fpdfPrintToPage := StrToInt(value);
       FPdfExportRange_Word := wdExportFromTo;
      except
            HaltWithConfigError( 202, '--PDF-TOPAGE must be an integer');
      end;


    end
    else if (id = '--PDF-OPTIMIZEFOR') or
            (id = '--PDF-OPTIMISEFOR') then
    BEGIN

         if (WordConstants.Exists(value)) then
         begin
           pdfOptimizeFor := StrToInt( WordConstants.Values[value]);
           logdebug('Set pdfOptimizeFor To: '  + value + ':' + InttoStr(pdfOptimizeFor), Verbose);
         end else
         begin
           HaltWithConfigError(205,'Invalid value for --PDF-OPTIMIZEFOR :' + value);
         end;
    END
    else if (id = '--EXPORTMARKUP') then
    begin
         if (WordConstants.Exists(value)) then
         begin
           fExportMarkup := StrToInt( WordConstants.Values[value]);
           logdebug('Set fExportMarkup To: ' + InttoStr(fExportMarkup), Verbose);
         end else
         begin
           HaltWithConfigError(205,'Invalid value for --EXPORTMARKUP :' + value);
         end;
    end
    else if (id = '--NO-INCLUDEDOCPROPERTIES')
         OR (id = '--NO-DOCPROP') then
    begin
      FIncludeDocProps := false;
      dec(iParam);
    end
    else if (id = '--PDF-NO-DOCSTRUCTURETAGS') then
    begin
      FDocStructureTags := false;
      dec(iParam);
    end
    else if (id = '--XPS-NO-IRM') then
    begin
      FKeepIRM := false;
      dec(iParam);
    end
    else if (id = '--PDF-NO-BITMAPMISSINGFONTS') then
    begin
      FBitmapMissingFonts := false;
      dec(iParam);
    end
    else if (id = '-W') or
            (id = '--WEBHOOK') then
    begin
      FWebHook := value;
    end
    else if (id = '-V') then
    begin
      LogVersionInfo(true);
      halt(2);

    end
    else if (id = '--ENABLE-MACROAUTORUN') then
    begin
      fDisableMacros := false;
      if (OfficeAppName <> 'Word')then
      begin
      // Excel   Application.EnableEvents = False
      //  HaltWithError(301,'Parameter '  + id + ' not Implemented for ' + OfficeAppName );
      end;

    end
    else if (id = '-X') or
            (id = '--HALTERROR') then
    begin
      HaltOnWordError := not(lowercase(value) = 'false');
    end
    // Long form only
    else if (id = '--SKIPDOCSWITHTOC') then
    begin
      fSkipDocsWithTOC := true;
      dec(iParam);
    end
    else if (id = '--DONOTOVERWRITE') or
            (id = '--NO-OVERWRITE')  then
    begin
      logInfo('DoNotOverwrite=True',Verbose);
      fSkipDocsExist := true;
      dec(iParam);
    end
    // Help etc
    else if (id = '-H') or
            (id = '-?') or
            (id = '?') OR
            (id = '--HELP') then
    begin
      ConfigLogHelp(ID,value,Params);
    end
    else if (id = '--HELP-EXCEL') then
    begin
      LogResourceHelp('EXCELFORMATS');
      halt(2);
    end
    else if (id = '-HJ') then
    begin
    log( 'hJ');
    LogResourceHelp('HELPJSON');
      halt(2);
    end
    else if (id = '-HW') then
    begin
    LogResourceHelp('HELPWEBHOOK');
  halt(2);
    end
    else if (id = '-HX') then
    begin
   LogResourceHelp('HELPERRORS');
      halt(2);
    end
    else if (length(pstr) > 3) then
    begin
      // Check if long param without --
      if pos(pstr , '--') = 0 then
      begin

        if pstr[1] = '-' then
        begin
            HaltWithConfigError(203,'Unknown Switch or Parameter :'''  + pstr + '''. ' +  ' Did you mean to provide the long param  ''-' + pstr + '''');
        end;

      end;

    end
    else
    begin
      HaltWithConfigError(203,'Unknown Switch or Parameter at index '+inttostr(iParam -1) + ':'  + pstr);
    end;




  end;

  // Code to run when all parameters have been loaded.
  // Get Files

   // IsFileInput := true;
    // If input is Dir rather than file, enumerate files.
    if DirectoryExists(InputFile) then
    begin
       InputIsDir := true;


       ListFiles(finputfile, '*' + InputExtension,DoSubDirs,FInputFiles);
       if FInputFiles.Count = 0 then
       begin
         HaltWithError(204, 'No File Matches in Input Directory: ' + finputfile + '*' + InputExtension );
       end;
       log('File List', FInputFiles,STANDARD);
       logInfo('Beginning to convert files....',STANDARD);
    end
    else
    begin
      InputIsFile := true;
    end;



end;



procedure TDocumentConverter.Log(Msg: String; Level : Integer = ERRORS );
var
  OutputLog, OutputTimeStamp : Boolean;
begin
  Outputlog := false;
 //   Outputlog := true;
  OutputTimeStamp := false;




  if Level <= FLogLevel then
  begin
    OutputLog := true;
  end;

  if FFirstLogEntry then
  begin
    OutputTimeStamp := true;
  end;

  if Level = HELP then
  begin
      OutputLog := true;
      OutputTimeStamp := false;
  end;

  if OutputTimeStamp then
  begin
    FFirstLogEntry := false;
    Msg := '[' + FormatDateTime('YYYYMMDD HH:NN:SS -' , now) +  ']: '  +  Msg;
  end;


  // Only actually output the log to console and file if levels match.
  if OutputLog = true then
  begin
    ConsoleLog.Log(self, Msg);

    if FLogtoFile then
    begin
      FLogFile.Add(Msg);
      FLogFile.SaveToFile(FLogFilename);
    end;

  end;
end;

procedure TDocumentConverter.Log(Msg: String; List:  TStrings; Level: Integer);
begin
  log(Msg, Level);
  log(  List.Text, Level);

end;

procedure TDocumentConverter.LogDebug(Msg: String; Level: Integer);
begin
  log('[DEBUG]  ' + Msg, Level);
end;

procedure TDocumentConverter.LogError(Msg: String);
begin

  Log('*******************************************', ERRORS);
  Log('Error: ' + Msg, ERRORS);
  Log('*******************************************', ERRORS);
end;

procedure TDocumentConverter.LogExcelFormats;
begin
    log('Format HELP');
   log('DOCTO');
    log('Listed below are all Accepted formats that Excel will export to and their associated file extensions:');
    log('--');
     LogResourceHelp('XLSEXTENSIONS');
end;

procedure TDocumentConverter.LogPowerPointFormats;
begin
    log('Format HELP');
   log('DOCTO');
    log('Listed below are all Accepted formats that PowerPoint will export to and their associated file extensions:');
    log('--');
     LogResourceHelp('PPEXTENSIONS');
end;

procedure TDocumentConverter.LogWordFormats;
begin
    log('Format HELP');
   log('DOCTO');
    log('Listed below are all Accepted formats that Word will export to and their associated file extensions:');
    log('--');
     LogResourceHelp('DOCEXTENSIONS');
end;

procedure TDocumentConverter.HaltWithConfigError(ErrorNo: Integer; Msg: String);
begin

  Log('*******************************************', ERRORS);
  Log('CONFIG ERROR: ' + Msg, ERRORS);
  Log('*******************************************', ERRORS);
  LOG('Please check help -h for further information.');
  LOG('halting without any conversions....');
  halt(200);
end;

procedure TDocumentConverter.LogResourceHelp(HelpResName: String);
var HelpStrings : TResourceStrings;
begin
      HelpStrings := TResourceStrings.Create(HelpResName);
      try

        log(HelpStrings.Text,Help);
      finally
        HelpStrings.Free;
      end;
end;

procedure TDocumentConverter.LogInfo(Msg: String; Level: Integer);
begin
  log('[INFO]   ' + Msg, Level);
end;



procedure TDocumentConverter.LogVersionInfo(ForceReload: Boolean = true);
begin

      // Version is cached in temp file.  Delete file to force check.
      if ForceReload then
      begin
        DeleteFile(ConfigFileName);
      end;

      // Prevent Date from Printing.
      FFirstLogEntry := false;

      // Log versions.
      log('DocTo Version:' + DOCTO_VERSION + DOCTO_VERSION_NOTE);
      log('OfficeApp Version:' +  OfficeAppVersion(),0);
      log('Source: https://github.com/tobya/DocTo/');

end;



function TDocumentConverter.NewFileNameFromBase(OldBase, NewBase,
  FileName, NewExt : String): String;
var
  BaseLessFN : String;
  NewFileName : String;
begin
  BaseLessFN :=  StringReplace(filename,oldbase,'',[rfReplaceAll]);
  NewFileName := IncludeTrailingBackslash(NewBase) + BaseLessFN;
  NewFileName := ChangeFileExt(NewFileName , NewExt);
  Result := NewFileName;
end;

function TDocumentConverter.OnConversionError(InputFile, OutputFile, Error: String):string;
var url_end : string;
begin

  url_end :=  'action=error&type='+ FOutputFileFormatString + '&outputfilename=' + URLEncode(OutputFile)
                                  + '&inputfilename=' + URLEncode(InputFile)
                                    + '&error=' + Error;
  CallWebHook(url_end);

end;


function TDocumentConverter.ReadOfficeAppVersion: String;
var
  ConfigFile : TStringlist;
  value, fullfn: String;
begin

  ConfigFile := TStringList.Create();
  fullfn := ConfigFileName;
  if FileExists(fullfn) then
  begin
    ConfigFile.LoadFromFile(fullfn);

    value := ConfigFile.Values[    OfficeAppName + '_' + FormatDateTime('yyyymmdd',now)];
    result := value;
  end
  else
  begin
  result := '';
  end;



end;

procedure TDocumentConverter.WriteOfficeAppVersion(Version: String);
var
  ConfigFile : TStringlist;
  TempDir, ConfigFn, fullfn, value: String;
begin
  ConfigFile := TStringList.Create();
  ConfigFile.Add( OfficeAppName + '_' + FormatDateTime('yyyymmdd',now) + '=' + Version);
  ConfigFile.SaveToFile(ConfigFileName);
  LogDebug('Writing Version to File:' + ConfigFileName,VERBOSE);
end;

procedure TDocumentConverter.SetBitmapMissingFonts(const Value: boolean);
begin
  FBitmapMissingFonts := Value;
end;

procedure TDocumentConverter.SetCompatibilityMode(const Value: Integer);
begin
  FCompatibilityMode := Value;
end;

procedure TDocumentConverter.SetDocStructureTags(const Value: boolean);
begin
  FDocStructureTags := Value;
end;

procedure TDocumentConverter.SetDoSubDirs(const Value: Boolean);
begin
  FDoSubDirs := Value;
end;

function TDocumentConverter.GetExtension: String;
begin
  Result := fInputExtension;
end;



function TDocumentConverter.getIsExcel: Boolean;
begin
    Result := MSExcel = FAppID;
end;

function TDocumentConverter.getIsPP: Boolean;
begin
    Result := MSPOWERPOINT = FAppID;
end;

function TDocumentConverter.getIsVisio: Boolean;
begin
  Result := MSVISIO = FAppID;
end;

function TDocumentConverter.getIsWord: Boolean;
begin
    Result := MSWord = FAppID;
end;



procedure TDocumentConverter.SetEncoding(const Value: Integer);
begin
  FEncoding := Value;
end;

procedure TDocumentConverter.SetExtension(const Value: String);
begin
  FInputExtension := Value;
end;

procedure TDocumentConverter.SetHaltOnWordError(const Value: Boolean);
begin
  FHaltOnWordError := Value;
end;

procedure TDocumentConverter.SetIgnore_ErrorDocs(const Value: Boolean);
begin
  FIgnore_ErrorDocs := Value;
end;

procedure TDocumentConverter.SetIgnore_MACOSX(const Value: boolean);
begin
  FIgnore_MACOSX := Value;
end;

procedure TDocumentConverter.SetList_ErrorDocs(const Value: Boolean);
begin
  FList_ErrorDocs := Value;
end;

procedure TDocumentConverter.SetList_ErrorDocs_Seconds(const Value: Integer);
begin
  FList_ErrorDocs_Seconds := Value;
end;

procedure TDocumentConverter.SetIncludeDocProps(const Value: boolean);
begin
  FIncludeDocProps := Value;
end;

procedure TDocumentConverter.SetInputFile(const Value: String);
begin
  FInputFile := Value;
end;

procedure TDocumentConverter.SetIsDirInput(const Value: Boolean);
begin
  FInputIsDir := Value;
end;

procedure TDocumentConverter.SetIsDirOutput(const Value: Boolean);
begin
  FOutputIsDir := Value;
end;

procedure TDocumentConverter.SetIsFileInput(const Value: Boolean);
begin
  FInputIsFile := Value;
end;

procedure TDocumentConverter.SetIsFileOutput(const Value: Boolean);
begin
  FOutputIsFile := Value;
end;

procedure TDocumentConverter.SetKeepIRM(const Value: boolean);
begin
  FKeepIRM := Value;
end;

procedure TDocumentConverter.SetLogFilename(const Value: String);
begin
  FLogFilename := Value;
end;

procedure TDocumentConverter.SetLogLevel(const Value: integer);
begin
  FLogLevel := Value;
  OutputLog := true;
  if FLogLevel = 0 then
  begin
    OutputLog := false;
    FLogLevel := ERRORS;
  end;
end;

procedure TDocumentConverter.SetLogToFile(const Value: Boolean);
begin
  FLogToFile := Value;

  // Set up logfile.
  if FLogtoFile then
  begin
    if not assigned(fLogFile) then
    begin
      FLogFile :=TStringList.Create;
    end;
    if fileexists(FLogFilename) then
    begin
      flogfile.LoadFromFile(fLogFileName);
    end;
    logInfo('['+FormatDateTime('YYYYMMDD HH:NN:SS -' , now )+ ']: LogFile Loaded', STANDARD);
  end
  else
  begin
    FLogFile.SaveToFile(FLogFilename);
    FLogFile.Free;
    FLogFile := nil;
  end;
end;

procedure TDocumentConverter.SetPDFOpenAfterExport(const Value: Boolean);
begin   FPDFOpenAfterExport := Value;  end;

procedure TDocumentConverter.SetOutputExt(const Value: string);
begin    FOutputExt := Value; end;

procedure TDocumentConverter.SetOutputFile(const Value: String);
begin
  FOutputFile := Value;
end;

procedure TDocumentConverter.SetOutputFileFormat(const Value: Integer);
begin
  FOutputFileFormat := Value;
end;

procedure TDocumentConverter.SetOutputFileFormatString(const Value: String);
begin
  FOutputFileFormatString := Value;
end;

procedure TDocumentConverter.SetOutputIsStdOut(const Value: Boolean);
begin
  FOutputIsStdOut := Value;
end;

procedure TDocumentConverter.SetOutputLog(const Value: Boolean);
begin
  FOutputLog := Value;
end;

procedure TDocumentConverter.SetOutputLogFile(const Value: String);
begin
  FOutputLogFile := Value;
end;

procedure TDocumentConverter.SetpdfExportRange(const Value: Integer);
begin
  FPdfExportRange_Word := Value;
end;

procedure TDocumentConverter.SetRemoveFileOnConvert(const Value: boolean);
begin
  FRemoveFileOnConvert := Value;
end;




procedure TDocumentConverter.SetSkipDocsWithTOC(const Value: Boolean);
begin
  FSkipDocsWithTOC := Value;
end;



function TDocumentConverter.URLEncode(Param: String): String;
begin
  Result :=  TIDURI.ParamsEncode(Param);
end;



function IsNumber(Str: String) : Boolean;
var
  i : integer;
begin
  Result := true;
  try
  i := strtoint(Str);

  except
    Result := false;
  end;
end;

procedure TDocumentConverter.ListFiles(const PathName, FileName : string; const SubDir : boolean; outFiles: TStrings);
var Rec  : TSearchRec;
    Path : string;
begin
Path := IncludeTrailingBackslash(PathName);
if FindFirst(Path + FileName, faAnyFile - faDirectory, Rec) = 0 then
 try
   repeat
     outFiles.Add(Path + Rec.Name);
   until FindNext(Rec) <> 0;
 finally
   FindClose(Rec);
 end;

If not SubDir then Exit;

if FindFirst(Path + '*.*', faDirectory, Rec) = 0 then
 try
   repeat

     if ((Rec.Attr and faDirectory) <> 0)  and (Rec.Name<>'.') and (Rec.Name<>'..') then
     BEGIN
      if AllowDirectory(Rec.Name, Path + Rec.Name) then
      begin
       ListFiles(Path + Rec.Name, FileName, True, outFiles);
      end;
     end;

   until FindNext(Rec) <> 0;
 finally
   FindClose(Rec);
 end;
end; //procedure FileSearch





function TDocumentConverter.GetUrl(Url: string): String;
var

  UrlHandle: HINTERNET;
  Buffer: array[0..1023] of byte;
  BytesRead: DWord;
  StrBuffer: UTF8String;
begin


  Result := '';
  // Intialising Win Internet Connection with InternetOpen only needs to be called once.
  if not assigned(FNetHandle) then
  begin
    FNetHandle := InternetOpen('github/tobya/DocTo', INTERNET_OPEN_TYPE_PRECONFIG, nil, nil, 0);
  end;
  if Assigned(FNetHandle) then
  begin
      UrlHandle := InternetOpenUrl(FNetHandle, PChar(Url), nil, 0, INTERNET_FLAG_RELOAD, 0);
      if Assigned(UrlHandle) then
        try
          repeat
            InternetReadFile(UrlHandle, @Buffer, SizeOf(Buffer), BytesRead);
            SetString(StrBuffer, PAnsiChar(@Buffer[0]), BytesRead);
            Result := Result + StrBuffer;
          until BytesRead = 0;
        finally
          InternetCloseHandle(UrlHandle);
        end
      else
      begin
        raise Exception.CreateFmt('Cannot open URL: %s', [Url]);
      end;
  end
  else
  begin
    raise Exception.Create('Unable to initialize Wininet');
  end;
end;


function TDocumentConverter.getWordConstants: TResourceStrings;
begin
  if FWordConstants = nil then
  begin
         fWordConstants := TResourceStrings.Create('WORDCONSTANTS');
         fWordConstants.Append('WORDCONSTANTS_EXTRA');
  end;
  Result := FWordConstants;

end;

function TDocumentConverter.GetSSLUrl(Url: string): String;
begin
  Result :=  dmssl.IdHTTP1.Get(Url);
end;


function TDocumentConverter.AfterConversion(InputFile, OutputFile: String):string;
var UrlToCall : String;
begin
  UrlToCall := 'action=convert&type='+ FOutputFileFormatString + '&outputfilename='
                + URLEncode(OutputFile)+ '&inputfilename='
                + URLEncode(InputFile);

  //Make a call to webhook if it exists
  Result := CallWebHook(UrlToCall);
end;

procedure TDocumentConverter.AfterListConvert;
begin

end;

function TDocumentConverter.AllowDirectory(DirName, FullPath: String): Boolean;
begin
    Result := true;
    if (Ignore_MACOSX) then
      if (DirName = '__MACOSX') then
      BEGIN
        Result := false;
      END;
end;

function TDocumentConverter.AllowFile(FileName, Fullpath: String): Boolean;
begin

end;



procedure TDocumentConverter.BeforeListConvert;
begin

end;

// Check howlong a document took to convert.  If greater > X then record in ignorelist.
// This can be used to find error documents as a modal window is displayed and execution does not
// continue until MANUALLY clicked. Suggested value of -N 5 seconds.
procedure TDocumentConverter.CheckDocumentTiming(StartTime, EndTime: cardinal; DocumentPath : String);
var
  sl : TStringList;
  ignorecount : integer;
const
  ignorelistfilename = 'docto.ignore.txt';
begin
  if (List_ErrorDocs) then
  begin

    // Check if the length of time taken to convert.
    if ((EndTime - StartTime) /1000) > List_ErrorDocs_Seconds then
    begin

      sl := TStringList.Create();
      try

        if  FileExists(ignorelistfilename) then
        begin
          sl.LoadFromFile(ignorelistfilename);
        end else begin
          sl.Add('[Comments]');
          sl.Add('COMMENT1=THIS FILE RECORDS ANY WORD DOCUMENTS THAT TOOK LONGER THAN X SECONDS TO COMPLETE.');
          sl.Add('COMMENT2=THIS SHOULD, OVER TIME ALLOW YOU TO TRACK DOWN ALL FILES CAUSING PROBLEMS');
          SL.Add('COMMENT3=THESE FILES WILL BE IGNORED ON SUBSEQUENT RUNS AS LONG AS "-NX" IS USED');
          SL.Add('COMMENT4=----DO NOT DELETE THIS FILE-----------');
          SL.Add('[FILES TO IGNORE]');
          SL.Add('IGNORECOUNT=0');
        end;

        log('Writing filename to Ignore List: ' + DocumentPath , CHATTY);

        ignorecount := StrToInt(sl.Values['IGNORECOUNT']);
        INC(ignorecount);

        sl.Add('IGNOREFILE' + INTTOSTR(ignorecount) + '=' + DocumentPath);
        sl.Values['IGNORECOUNT'] := INTTOSTR(ignorecount);
        sl.SaveToFile('docto.ignore.txt');

      finally
        sl.free;
      end;

    end;
  end;
end;


(*
  Check if document is to be ignored. Load ignore file and checklist.
*)
function TDocumentConverter.CheckShouldIgnore(DocumentPath: String): Boolean;
var
 // sl : TStringList;
  fn : string;
  ignorecount : integer;
  I: Integer;
const
  ignorelistfilename = 'docto.ignore.txt';
begin
  Result := False;
  if (Ignore_ErrorDocs) then
  begin

      //Create and Load on first call
      if FIgnoreErrorDocsFile = nil then
      begin
        FIgnoreErrorDocsFile := TStringList.Create();

        if  FileExists(ignorelistfilename) then
        begin
          FIgnoreErrorDocsFile.LoadFromFile(ignorelistfilename);
        end else begin
          log('No docto.ignore.txt file found.', CHATTY);
          Result := false;
          exit;
        end;
      end;


      if FIgnoreErrorDocsFile.Count > 0 then
      begin
        log('Checking Ignore List for filename:' + DocumentPath, VERBOSE);
        ignorecount := StrToInt(FIgnoreErrorDocsFile.Values['IGNORECOUNT']);

        for I := 1 to ignorecount do
        BEGIN
         fn := FIgnoreErrorDocsFile.Values['IGNOREFILE' + INTTOSTR(I)];
         log('Check Against: ' + fn, VERBOSE);
         if fn = DocumentPath then
         begin
           Result := true;
           break;
         end;
        END;
      end;

  end;

end;


end.
