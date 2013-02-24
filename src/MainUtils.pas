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
uses classes, WordUtils, sysutils, ActiveX, ComObj, WinINet,   Types,  ResourceUtils,
     PathUtils;

type
  TConsoleLog = class
  public
    procedure Log(Sender: TObject; Log : String);
    procedure LogError(Log: String);
  end;

  TDocumentConverter = class
  private
    Formats : TStringlist;
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

    FHaltOnWordError: Boolean;
    FRemoveFileOnConvert: boolean;

    WordApp : OleVariant;
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
    property IsFileInput : Boolean read FIsFileInput write SetIsFileInput;
    property IsDirInput : Boolean read FIsDirInput write SetIsDirInput;
    property DoSubDirs : Boolean read FDoSubDirs write SetDoSubDirs;
    property OutputExt : string read FOutputExt write SetOutputExt;
    property RemoveFileOnConvert: boolean read FRemoveFileOnConvert write SetRemoveFileOnConvert;
  public

    Constructor Create();
    Destructor Destroy(); override;
    procedure LoadConfig(Params: TStrings);


    function Execute() : string;

    procedure Log(Msg: String; Level  : Integer = 0);
    procedure LogError(Msg: String);
    procedure CallWebHook(Params: String);

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

  end;



  function IsNumber(Str: String) : Boolean;

  Const
  VERBOSE = 10;
  CHATTY = 5;
  STANDARD = 1;
  SILENT = 0;
  ERRORS = 0;
implementation


{ TConsoleLog }

procedure  TConsoleLog.Log(Sender: TObject; Log: String);
begin
  Writeln(Log);
end;

procedure TDocumentConverter.CallWebHook(Params: String);
begin
  if FWebHook > '' then
  begin
    GetURL(FWebHook + '?' + Params);
  end;
end;

constructor TDocumentConverter.Create;
begin
  ConsoleLog := TConsoleLog.Create();
  FVersionString := '0.3ALPHA';

  //Initial values
  FOutputFileFormatString := '';
  FOutputFileFormat := -1;
  FOutputFile := '';
  FInputFile := '';
  FLogLevel := 1;
  FLogtoFile := false;
  FLogFilename := 'DocTo.Log';
  FRemoveFileOnConvert := false;
  FWebHook := '';
  FOutputExt := '';

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
  FileToConvert, OutputFilename : String;

begin

    Continue := false;
    if (InputFile > '') and (OutputFile > '') and (OutputFileFormat > -1) then
    begin
      Continue := true;
    end;

    if not Continue  then HaltWithError(201, 'Input File, Output File and FileFormat must all be specified');

    //Add file to InputFiles List if only one.
    if FInputFiles.Count = 0 then
    begin
      FInputFiles.Add(FInputFile);
    end;

   try
    Wordapp :=  CreateOleObject('Word.Application');
    Wordapp.Visible := false;

    for i := 0 to FInputFiles.Count -1 do
    begin
      FileToConvert := FInputFiles[i];
      if IsDirInput then
      begin
        OutputFilename := NewFileNameFromBase(FInputFile ,FOutputFile,FileToConvert, FOutputExt);
      end
      else
      begin
        OutputFilename :=  OutputFile;
      end;

        //Ensure directory exists
        ForceDirectories(ExtractFilePath( OutputFilename));



      log('Ready to Execute' , VERBOSE);
       try

            //Open doc and save in requested format.
            Wordapp.documents.Open(FileToConvert, false, true);
            Wordapp.activedocument.Saveas(OutputFilename ,OutputFileFormat );

            Wordapp.activedocument.Close;

            if RemoveFileOnConvert then
            begin
              //Check file exists and Delete if requested
              if FileExists(OutputFilename) then
              begin
                DeleteFile(FileToConvert);
                Log('Deleted:' + FileToConvert,STANDARD);
              end;
            end;

            //Make a call to webhook if it exists
            CallWebHook('action=convert&type='+ FOutputFileFormatString + '&ouputfilename=' + URLEncode(OutputFilename)+ '&inputfilename=' + URLEncode(InputFile));

          log(OutputFilename,STANDARD);
          result := OutputFilename;
        except
          on E: EOleSysError do
          begin
            if pos('Invalid class string',E.Message) > 0 then
            begin
              HaltWithError(221,'Word Does not appear to be installed:' +E.ClassName + '  ' + e.Message);
            end
            else
            begin

              CallWebHook('action=error&type='+ FOutputFileFormatString + '&ouputfilename=' + URLEncode(OutputFilename)+ '&inputfilename=' + URLEncode(InputFile)
                          + '&error=' + URLEncode(E.ClassName + '  ' + e.Message));

              if (HaltOnWordError) then
              begin
                log('FileToConvert:' + FileToConvert);
                log('OutputFile:' + OutputFilename);
                log('Ext' + inttostr(OutputFileFormat));
              HaltWithError(220,E.ClassName + '  ' + e.Message);
              end
              else
              begin
                log('FileToConvert:' + FileToConvert);
                log('OutputFile:' + OutputFilename);
                log('Ext' + inttostr(OutputFileFormat));
                logerror(E.ClassName + '  ' + e.Message);

              end;
            end;

          end;
          on E: Exception do
          begin
              if (HaltOnWordError) then
              begin
                HaltWithError(220,E.ClassName + '  ' + e.Message + ' ' + FileToConvert + ':' + OutputFilename);
              end
              else
              begin
                LogError(E.ClassName + '  ' + e.Message + ' ' + FileToConvert + ':' + OutputFilename);

              end;
          end;
        end;

    end;

    finally

    wordapp.quit();
    end;


end;



procedure TDocumentConverter.HaltWithError(ErrorNo: Integer; Msg: String);
begin
  LogError(Msg);
  LogError('Exiting with Error Code : ' + inttostr(ErrorNo));

  if not VarIsEmpty(WordApp) then
  begin
    WordApp.Quit();
  end;
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
var i, f , iParam, idx: integer;
pstr : string;
id, value : string;
HelpStrings : TStringList;

begin
  //Initislise
  iParam := 0;
  Formats := AvailableWordFormats();


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


    if id = '-O' then
    begin
      FOutputFile := value;
      log('Output file is : ' + FOutputFile,CHATTY);
    end
    else if id = '-OX' then
    begin
       FOutputExt := value;
    end
    else if id = '-F' then
    begin
      FInputFile := value;
      IsFileInput := true;
      if DirectoryExists(FInputFile) then
      begin
         IsDirInput := true;
         DoSubDirs := true;
         ListFiles(finputfile, '*.doc',true,FInputFiles);
      end
      else
      begin

        IsFileInput := true;
      end;


      log('Input File is : ' + FInputFile,CHATTY);
    end
    else if id  = '-Q' then
    begin

      OutputLog := false;
      //Doesn't require a value
      dec(iParam);
    end
    else if id  = '-L' then
    begin
      if isNumber(value) then
      begin
        FLogLevel := strtoint(value);
        Log('Log Level Set To:' + IntToStr(FLogLevel),FLogLevel);
      end
    end
    else if (id = '-T') or (id = '-TF') then
    begin

      if IsNumber(value) then
      begin
        FOutputFileFormat :=  strtoint(value);
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
    else if (id = '-G') then
    begin
       LogToFile := true;
       dec(iParam);
    end
    else if (id = '-GL') then
    begin
       FLogFilename := value;
       LogToFile := true;
    end
    else if (id = '-R') then
    begin
      RemoveFileOnConvert  := lowercase(value) = 'true';
    end
    else if (id = '-W') then
    begin
      FWebHook := value;
    end
    else if (id = '-X') then
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
      HelpStrings := TStringList.Create;
      try
        LoadStringListFromResource('HELPJSON',HelpStrings);
        log(HelpStrings.Text);
      finally
        HelpStrings.Free;
      end;
      halt(2);
    end
    else if (id = '-HW') then
    begin
      HelpStrings := TStringList.Create;
      try
        LoadStringListFromResource('HELPWEBHOOK',HelpStrings);
        log(HelpStrings.Text);
      finally
        HelpStrings.Free;
      end;
      halt(2);
    end
    else
    begin
      HaltWithError(203,'Unknown Switch:' + pstr);
    end;




  end;



end;


procedure TDocumentConverter.Log(Msg: String; Level : Integer = 0 );
begin
  if Level <= FLogLevel then
  begin
    ConsoleLog.Log(self, Msg);
    if FLogtoFile then
    begin
      FLogFile.Add(Msg);
      FLogFile.SaveToFile(FLogFilename);
    end;
  end;
end;

procedure TDocumentConverter.LogError(Msg: String);
begin
  Log('*******************************************', ERRORS);
  Log('Error: ' + Msg, ERRORS);
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

procedure TDocumentConverter.SetDoSubDirs(const Value: Boolean);
begin
  FDoSubDirs := Value;
end;

procedure TDocumentConverter.SetHaltOnWordError(const Value: Boolean);
begin
  FHaltOnWordError := Value;
end;

procedure TDocumentConverter.SetInputFile(const Value: String);
begin
  FInputFile := Value;
end;

procedure TDocumentConverter.SetIsDirInput(const Value: Boolean);
begin
  FIsDirInput := Value;
end;

procedure TDocumentConverter.SetIsFileInput(const Value: Boolean);
begin
  FIsFileInput := Value;
end;

procedure TDocumentConverter.SetLogFilename(const Value: String);
begin
  FLogFilename := Value;
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
     ListFiles(Path + Rec.Name, FileName, True, outFiles);
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



end.
