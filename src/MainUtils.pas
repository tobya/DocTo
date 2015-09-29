unit MainUtils;
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
uses classes, Windows, sysutils, ActiveX, ComObj, WinINet, Variants,  Types,  ResourceUtils,
     PathUtils;

Const
  VERBOSE = 10;
  CHATTY = 5;
  STANDARD = 2;
  SILENT = 0;
  ERRORS = 1;

type


  TConsoleLog = class
  public
    procedure Log(Sender: TObject; Log : String);
    procedure LogError(Log: String);
  end;

  TDocumentConverter = class
  private
    FIgnore_MACOSX: boolean;

    procedure SetCompatibilityMode(const Value: Integer);
    procedure SetIgnore_MACOSX(const Value: boolean);
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
    FIsFileInput: Boolean;
    FIsDirInput: Boolean;
    FOutputExt: string;
    FWebHook : String;
    FInputExtension : String;
    FCompatibilityMode: Integer;

    FHaltOnWordError: Boolean;
    FRemoveFileOnConvert: boolean;


    FIsFileOutput: Boolean;
    FIsDirOutput: Boolean;
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
    procedure ListFiles(const PathName, FileName: string; const InDir: boolean; outFiles: TStrings);
    procedure SetDoSubDirs(const Value: Boolean);
    procedure SetIsDirInput(const Value: Boolean);
    procedure SetIsFileInput(const Value: Boolean);
    function NewFileNameFromBase(OldBase, NewBase, FileName, NewExt: String): String;
    procedure SetOutputExt(const Value: string);
    function GetUrl(Url: string): String;
    function URLEncode(Param : String): String;
    procedure SetHaltOnWordError(const Value: Boolean);
    procedure SetRemoveFileOnConvert(const Value: boolean);
    procedure SetIsDirOutput(const Value: Boolean);
    procedure SetIsFileOutput(const Value: Boolean);
    procedure SetLogLevel(const Value: integer);
    property IsFileInput : Boolean read FIsFileInput write SetIsFileInput;
    property IsDirInput : Boolean read FIsDirInput write SetIsDirInput;
    property IsFileOutput : Boolean read FIsFileOutput write SetIsFileOutput;
    property IsDirOutput : Boolean read FIsDirOutput write SetIsDirOutput;
    property DoSubDirs : Boolean read FDoSubDirs write SetDoSubDirs;
    property OutputExt : string read FOutputExt write SetOutputExt;
    property LogLevel : integer read FLogLevel write SetLogLevel;
    property RemoveFileOnConvert: boolean read FRemoveFileOnConvert write SetRemoveFileOnConvert;
    property Ignore_MACOSX : boolean   read FIgnore_MACOSX write SetIgnore_MACOSX;


    procedure SetExtension(const Value: String); virtual;
    function GetExtension: String;  virtual;
    function OfficeAppVersion() : String; virtual; abstract;

    //Check files and folders
    function AllowDirectory(DirName : String; FullPath : String) : Boolean;
    function AllowFile(FileName : String; Fullpath : String): Boolean;



  public

    Constructor Create();
    Destructor Destroy(); override;
    procedure LoadConfig(Params: TStrings);
    procedure ConfigLoggingLevel(Params: TStrings);

    function Execute() : string; virtual;
    function ExecuteConversion(fileToConvert: String; OutputFilename: String; OutputFileFormat : Integer): string; virtual; abstract;

    function DestroyOfficeApp() : boolean; virtual; abstract;
    function CreateOfficeApp() : boolean; virtual; abstract;
    function AvailableFormats() : TStringList;  virtual; abstract;
    function FormatsExtensions(): TStringList; virtual; abstract;

    procedure Log(Msg: String; Level  : Integer = ERRORS); overload;
    procedure Log(Msg: String; List:  TStrings; Level: Integer); overload;
    procedure LogError(Msg: String);
    function CallWebHook(Params: String) : string;
    procedure LogHelp(HelpResName : String);

    property OutputLog : Boolean read FOutputLog write SetOutputLog;
    property OutputLogFile : String read FOutputLogFile write SetOutputLogFile;
    Property InputFile : String read FInputFile write SetInputFile;
    Property OutputFile : String read FOutputFile write SetOutputFile;
    Property OutputFileFormat : Integer read FOutputFileFormat write SetOutputFileFormat;
    Property OutputFileFormatString : String read FOutputFileFormatString write SetOutputFileFormatString;
    Property LogToFile : Boolean read FLogToFile write SetLogToFile;
    property LogFilename: String read FLogFilename write SetLogFilename;
    Property Version : String read FVersionString;
    property HaltOnWordError : Boolean read FHaltOnWordError write SetHaltOnWordError;
    property InputExtension: String read GetExtension write SetExtension;
    property CompatibilityMode : Integer read FCompatibilityMode write SetCompatibilityMode;

  end;



  function IsNumber(Str: String) : Boolean;


implementation


{ TConsoleLog }

procedure  TConsoleLog.Log(Sender: TObject; Log: String);
begin
  Writeln(Log);
end;



function TDocumentConverter.CallWebHook(Params: String):string;
var url : string;
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


    GetURL(url);

    log(url, chatty);
  end;
  except on E: Exception do
  begin
    logerror(E.ClassName + ' ' + E.Message);
  end;

  end;
end;

procedure TDocumentConverter.ConfigLoggingLevel(Params: TStrings);
var
  iParam : Integer;
  id, pstr, value : String;
begin
LogLevel := STANDARD;
iParam := 0;
lOG(ID,VERBOSE);
While iParam <= Params.Count -1 do
  begin
    pstr := Params[iParam];
    log(inttostr(iparam), VERBOSE);
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
    inc(iParam,2);
    lOG(ID,VERBOSE);
    if id  = '-L' then
    begin
      if isNumber(value) then
      begin
        LogLevel := strtoint(value);

      end
    end
    else if id  = '-Q' then
    begin

      OutputLog := false;
      //Doesn't require a value
      dec(iParam);
    end
  end;
  Log('Log Level Set To:' + IntToStr(FLogLevel),CHATTY);
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
  FIsFileInput := false;
  FIsDirInput := false;
  FIsFileOutput := false;
  FIsDirOutput := false;
  FCompatibilityMode := 0;
  FIgnore_MACOSX := true;

  FInputFiles := TStringList.Create;
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

  FInputFiles.Free;
end;


(*
  Execute conversion on 1 or more files.
*)
function TDocumentConverter.Execute: string;
var

  Continue : Boolean;
  i : integer;
  FileToConvert, FileToCreate, UrlToCall : String;

begin

    Continue := false;
    if (InputFile > '') and (OutputFile > '') and (OutputFileFormat > -1) then
    begin
      Continue := true;
    end;

    if not Continue  then HaltWithError(201, 'Input File, Output File and FileFormat must all be specified');

    //Set Output Filename if Dir Provided.
    if (IsFileInput and IsDirOutput) then
    begin
      if OutputExt = '' then
      begin
        OutputExt := fFormatsExtensions.Values[OutputFileFormatString];
        log('Output Extension is ' + outputExt, CHATTY);
      end;

      OutputFile :=  OutputFile  + ChangeFileExt( ExtractFileName(InputFile), '.' + OutputExt);
    end;

    //Add file to InputFiles List if only one.
    if FInputFiles.Count = 0 then
    begin
      FInputFiles.Add(FInputFile);
    end;

   try
    CreateOfficeApp();

    for i := 0 to FInputFiles.Count -1 do
    begin
      FileToConvert := FInputFiles[i];
      if IsDirInput then
      begin
        FileToCreate := NewFileNameFromBase(FInputFile ,FOutputFile,FileToConvert, FOutputExt);
      end
      else
      begin
        FileToCreate :=  OutputFile;
      end;

        //Ensure directory exists
        ForceDirectories(ExtractFilePath( FileToCreate));



      log('Ready to Execute' , VERBOSE);
      (*if FileExists(FileToCreate) then //Not working currently as file doesnt include .ext
      begin
        raise Exception.Create('FileExists Cannot Create: ' + FileToCreate);

      end;*)
       try


            ExecuteConversion(FileToConvert, FileToCreate, OutputFileFormat);
            log('FileCreated: ' + FileToCreate, STANDARD);
            if RemoveFileOnConvert then
            begin
              //Check file exists and Delete if requested
              if FileExists(FileToCreate) then
              begin
                DeleteFile(FileToConvert);
                Log('Deleted:' + FileToConvert,STANDARD);
              end;
            end;



            UrlToCall := 'action=convert&type='+ FOutputFileFormatString + '&outputfilename=' + URLEncode(FileToCreate)+ '&inputfilename=' + URLEncode(InputFile);

            //Make a call to webhook if it exists
            CallWebHook(UrlToCall);

          log('Creating File: ' + FileToCreate,CHATTY);

        except
          on E: EOleSysError do
          begin
            if pos('Invalid class string',E.Message) > 0 then
            begin
              HaltWithError(221,'Word Does not appear to be installed:' +E.ClassName + '  ' + e.Message);
            end
            else
            begin

              CallWebHook('action=error&type='+ FOutputFileFormatString + '&outputfilename=' + URLEncode(FileToCreate)+ '&inputfilename=' + URLEncode(InputFile)
                          + '&error=' + URLEncode(E.ClassName + '  ' + e.Message));

              if (HaltOnWordError) then
              begin
                log('FileToConvert:' + FileToConvert);
                log('OutputFile:' + FileToCreate);
                log('Ext' + inttostr(OutputFileFormat));
              HaltWithError(220,E.ClassName + '  ' + e.Message);
              end
              else
              begin
                log('FileToConvert:' + FileToConvert);
                log('OutputFile:' + FileToCreate);
                log('Ext' + inttostr(OutputFileFormat));
                logerror(E.ClassName + '  ' + e.Message);

              end;
            end;

          end;
          on E: Exception do
          begin
              if (HaltOnWordError) then
              begin
                HaltWithError(220,E.ClassName + '  ' + e.Message + ' ' + FileToConvert + ':' + FileToCreate);
              end
              else
              begin
                LogError(E.ClassName + '  ' + e.Message + ' ' + FileToConvert + ':' + FileToCreate);

              end;
          end;
        end;

    end;

    finally

      DestroyOfficeApp();
    end;


end;




procedure TDocumentConverter.HaltWithError(ErrorNo: Integer; Msg: String);
begin
  LogError(Msg);
  LogError('Exiting with Error Code : ' + inttostr(ErrorNo));
  //Ensure word is quit before halting.
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

procedure TDocumentConverter.LoadConfig(Params: TStrings);
var  f , iParam, idx: integer;
pstr : string;
id, value : string;
HelpStrings : TStringList;
tmpext : String;

begin
  //Initialise
  iParam := 0;
  Formats := AvailableFormats();
  fFormatsExtensions := FormatsExtensions();
  ConfigLoggingLevel(Params);

  OutputLog := true;
  OutputLogFile := '';

  HaltOnWordError := true;

  log('Loading Configuration...',VERBOSE);
  log('Parameter Count is ' + inttostr(params.Count), VERBOSE);

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
    inc(iParam,2);


    if (id = '-O') or
       (id = '--OUTPUTFILE') then
    begin
      FOutputFile :=  value;


      tmpext := ExtractFileExt(FOutputFile);


      if (tmpext = '') then
      begin
        FOutputFile := IncludeTrailingBackslash(value);
        IsDirOutput := true;
        ForceDirectories(FOutputFile);
        log('Output directory is: ' + FOutputFile,CHATTY);
      end
      else
      begin
        IsFileOutput := true;
        log('Output file is: ' + FOutputFile,CHATTY);
      end;


    end
    else if (id = '-OX') or
            (id = '--OUTPUTEXTENSION') then
    begin
       FOutputExt := value;
    end
    else if (id = '-F') or
            (id = '--INPUTFILE') then
    begin
      FInputFile := value;



      log('Input File is: ' + FInputFile,CHATTY);
    end
    else if (id = '-FX') or
            (id = '--INPUTFILEEXTENSION') then
    begin
      InputExtension := value;
    end
    else if ( id  = '-L')
         OR (id = '--LOGLEVEL') then
    begin
      if isNumber(value) then
      begin
        LogLevel := strtoint(value);
        Log('Log Level Set To:' + IntToStr(LogLevel),LogLevel);
      end
    end
    else if (id  = '-Q') or
            (id = '--QUIET') then
    begin

      OutputLog := false;
      //Doesn't require a value
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
          LogError('File Format ' + value + ' is invalid, please see help. -h.  To force use, use -TF');
          halt(200);
        end;
      end
      else
      begin
        FOutputFileFormatString := value;

        idx := formats.IndexOfName(FOutputFileFormatString);
        if  idx > -1 then
        begin
          OutputFileFormat := strtoint(formats.Values[OutputFileFormatString]);

        end
        else if idx = -1 then
        begin
          Log('File Format ' + OutputFileFormatString + ' is invalid, please see help. -h');
          halt(200);
        end;
      end;
      log('Type Integer is: ' + inttostr(FOutputFileFormat), VERBOSE);

    end
    else if (id = '-C') or
            (id = '--COMPATIBILITY') then
    begin
      CompatibilityMode := strtoint(value);
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
    else if (id = '-R')
          or (id = '--DELETEFILES') then
    begin
      RemoveFileOnConvert  := lowercase(value) = 'true';
    end
    else if (id = '-W') or
            (id = '--WEBHOOK') then
    begin
      FWebHook := value;
    end
    else if (id = '-V') then
    begin
      log('Version:0.7.2');  //Move to ancestor class
      log('OfficeApp Version:' +  OfficeAppVersion,0);
      halt(2);

    end
    else if (id = '-X') or
            (id = '--halterror') then
    begin
      HaltOnWordError := not(lowercase(value) = 'false');
    end
    else if (id = '-H') or
            (id = '-?') or
            (id = '?') then
    begin
      HelpStrings := TStringList.Create;
      try
        LoadStringListFromResource('HELP',HelpStrings);
        log(HelpStrings.Text);
      finally
        HelpStrings.Free;
      end;
      log('');
      log('FILE FORMATS');
      for f := 0 to Formats.Count -1 do
      begin
        log(Formats.Names[f] + '=' + Formats.Values[Formats.Names[f]]);
      end;

      halt(2);
    end
    else if (id = '-HJ') then
    begin
    LogHelp('HELPJSON');
      halt(2);
    end
    else if (id = '-HW') then
    begin
    LogHelp('HELPWEBHOOK');
  halt(2);
    end
    else if (id = '-HX') then
    begin
   LogHelp('HELPERRORS');
      halt(2);
    end
    else
    begin
      HaltWithError(203,'Unknown Switch:' + pstr);
    end;




  end;

  //Code to run when all parameters have been loaded.
  //Get Files

   // IsFileInput := true;
    //If input is Dir rather than file, enumerate files.
    if DirectoryExists(InputFile) then
    begin
       IsDirInput := true;
       DoSubDirs := true;

       ListFiles(finputfile, '*' + InputExtension,true,FInputFiles);

       log('File List', FInputFiles,verbose);
    end
    else
    begin
      IsFileInput := true;
    end;



end;


procedure TDocumentConverter.Log(Msg: String; Level : Integer = ERRORS );
begin


  if Level <= FLogLevel then
  begin
    if OutputLog = true then
    begin
      ConsoleLog.Log(self, Msg);
    end;
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

procedure TDocumentConverter.LogError(Msg: String);
begin
  Log('*******************************************', ERRORS);
  Log('Error: ' + Msg, ERRORS);
end;

procedure TDocumentConverter.LogHelp(HelpResName: String);
var HelpStrings : TStringList;
begin
      HelpStrings := TStringList.Create;
      try
        LoadStringListFromResource(HelpResName,HelpStrings);
        log(HelpStrings.Text);
      finally
        HelpStrings.Free;
      end;
end;

function TDocumentConverter.NewFileNameFromBase(OldBase, NewBase,
  FileName, NewExt : String): String;
var
  BaseLessFN : String;
  NewFileName : String;
begin
  BaseLessFN :=  StringReplace(filename,oldbase,'',[rfReplaceAll]);
 // BaseLessFN := BaseLessFN + '\';
  NewFileName := NewBase + '\' + BaseLessFN;
  NewFileName := ChangeFileExt(NewFileName , NewExt);
  Result := NewFileName;
end;

procedure TDocumentConverter.SetCompatibilityMode(const Value: Integer);
begin
  FCompatibilityMode := Value;
end;

procedure TDocumentConverter.SetDoSubDirs(const Value: Boolean);
begin
  FDoSubDirs := Value;
end;

function TDocumentConverter.GetExtension: String;
begin
  Result := fInputExtension;
end;

procedure TDocumentConverter.SetExtension(const Value: String);
begin
  FInputExtension := Value;
end;

procedure TDocumentConverter.SetHaltOnWordError(const Value: Boolean);
begin
  FHaltOnWordError := Value;
end;

procedure TDocumentConverter.SetIgnore_MACOSX(const Value: boolean);
begin
  FIgnore_MACOSX := Value;
end;

procedure TDocumentConverter.SetInputFile(const Value: String);
begin
  FInputFile := Value;
end;

procedure TDocumentConverter.SetIsDirInput(const Value: Boolean);
begin
  FIsDirInput := Value;
end;

procedure TDocumentConverter.SetIsDirOutput(const Value: Boolean);
begin
  FIsDirOutput := Value;
end;

procedure TDocumentConverter.SetIsFileInput(const Value: Boolean);
begin
  FIsFileInput := Value;
end;

procedure TDocumentConverter.SetIsFileOutput(const Value: Boolean);
begin
  FIsFileOutput := Value;
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

  //Set up logfile.
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
  end
  else
  begin
    FLogFile.SaveToFile(FLogFilename);
    FLogFile.Free;
    FLogFile := nil;
  end;
end;

procedure TDocumentConverter.SetOutputExt(const Value: string);
begin
  FOutputExt := Value;
end;

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

procedure TDocumentConverter.SetOutputLog(const Value: Boolean);
begin
  FOutputLog := Value;
end;

procedure TDocumentConverter.SetOutputLogFile(const Value: String);
begin
  FOutputLogFile := Value;
end;

procedure TDocumentConverter.SetRemoveFileOnConvert(const Value: boolean);
begin
  FRemoveFileOnConvert := Value;
end;

function TDocumentConverter.URLEncode(Param: String): String;
begin
 result :=  param;
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

procedure TDocumentConverter.ListFiles(const PathName, FileName : string; const InDir : boolean; outFiles: TStrings);
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

If not InDir then Exit;

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

procedure TConsoleLog.LogError(Log: String);
begin

end;





function TDocumentConverter.GetUrl(Url: string): String;
var
  NetHandle: HINTERNET;
  UrlHandle: HINTERNET;
  Buffer: array[0..1023] of byte;
  BytesRead: DWord;
  StrBuffer: UTF8String;
begin
  Result := '';
  NetHandle := InternetOpen('Delphi XE', INTERNET_OPEN_TYPE_PRECONFIG, nil, nil, 0);
  if Assigned(NetHandle) then
    try
      UrlHandle := InternetOpenUrl(NetHandle, PChar(Url), nil, 0, INTERNET_FLAG_RELOAD, 0);
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
        raise Exception.CreateFmt('Cannot open URL %s', [Url]);
    finally
      InternetCloseHandle(NetHandle);
    end
  else
    raise Exception.Create('Unable to initialize Wininet');
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

end.
