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
uses classes, WordUtils, sysutils, ActiveX, ComObj;

type
  TConsoleLog = class
  public
    procedure Log(Sender: TObject; Log : String);

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
    procedure SetInputFile(const Value: String);
    procedure SetOutputFile(const Value: String);
    procedure SetOutputFileFormat(const Value: Integer);
    procedure SetOutputFileFormatString(const Value: String);
    procedure SetOutputLog(const Value: Boolean);
    procedure SetOutputLogFile(const Value: String);
    function IsValidFormat(FormatID : Integer): Boolean;
  public

    Constructor Create();
    Destructor Destroy();
    procedure LoadConfig(Params: TStrings);

    procedure Log(Msg: String);
    function Execute() : string;
    property OutputLog : Boolean read FOutputLog write SetOutputLog;
    property OutputLogFile : String read FOutputLogFile write SetOutputLogFile;
    Property InputFile : String read FInputFile write SetInputFile;
    Property OutputFile : String read FOutputFile write SetOutputFile;
    Property OutputFileFormat : Integer read FOutputFileFormat write SetOutputFileFormat;
    Property OutputFileFormatString : String read FOutputFileFormatString write SetOutputFileFormatString;
    Property Version : String read FVersionString;
  end;

  function IsNumber(Str: String) : Boolean;


implementation


{ TConsoleLog }

procedure  TConsoleLog.Log(Sender: TObject; Log: String);
begin


  Writeln(Log);

end;

constructor TDocumentConverter.Create;
begin
  ConsoleLog := TConsoleLog.Create();
  FVersionString := '0.1';
end;

destructor TDocumentConverter.Destroy;
begin
  ConsoleLog.Free;
end;

function TDocumentConverter.Execute: string;
var
  WordApp : OleVariant;
begin
    log('Ready to Execute');
    try
    try
    Wordapp :=  CreateOleObject('Word.Application');
    Wordapp.Visible := false;
    Wordapp.documents.Open(InputFile, false, true);
    Wordapp.activedocument.Saveas(OutputFile ,OutputFileFormat);
    Wordapp.activedocument.Close;
    wordapp.quit();
    result := OutputFile;
    except on E: Exception do
      log(E.ClassName + '  ' + e.Message);

    end;
    finally

    end;

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

begin

  Formats := AvailableWordFormats();

 (* TConsoleLog.Log(self,Formats.Text);

  for f := 0 to Formats.Count -1 do
  begin
   TConsoleLog.Log(self,Formats.Names[f] + '::');
  end;      *)

  //Defaults
  OutputLog := true;
  OutputLogFile := '';

  log('Loading Configuration...');
  log('Parameter Count is ' + inttostr(params.Count));

  if Params.Count = 0 then
  begin
      log('Parameters Expected: -H for help');
      halt(1);
  end ;


  While iParam <= Params.Count -1 do
  begin
    pstr := Params[iParam];

    id := UpperCase( pstr);
    if ParamCount -1 > iParam then
    begin
    value := Trim(Params[iParam +1]);
    end;
    inc(iParam,2);


    if id = '-O' then
    begin
      FOutputFile := value;
      log('Output file is : ' + FOutputFile);
    end
    else if id = '-F' then
    begin
      FInputFile := value;
      log('Input File is : ' + FInputFile);
    end
    else if (id = '-T') or (id = '-TF') then
    begin

      if IsNumber(value) then
      begin
        FOutputFileFormat :=  strtoint(value);
        if (id = '-TF') and ( not IsValidFormat(FOutputFileFormat)) then
        begin
          Log('File Format ' + OutputFileFormatString + ' is invalid, please see help. -h');
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
      log('Type is: ' + inttostr(FOutputFileFormat));
    end
    else if (id = '-H') then
    begin
      Log('Help');
      Log('Version:' + Version);
      Log('Command Line Params');
      log('Each Parameter should be followed by its value  -f "c:\Docs\MyDoc.doc" -O "C:\MyDir\MyFile" ');
      log('  -H This message');
      log('  -F Input File or Directory');
      log('  -O Output File or Directory to place converted Docs');
      log('  -T Format to convert file to, either integer or wdSaveFormat constant. ');
      log('     Available from http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat.aspx ');
      log('     See current List Below. ');
      log('  -TF Force Format.  -T values are checked against current list compiled in and not passed if unavailable.  To future proof, -TF will pass through value without checking.  Word will return an "EOleException  Value out of range" error if invalid.');
      log('     Use instead of -T not as well as.');
      log('  -L Log to file in directory');
      log(' ');
      log('FILE FORMATS');
      for f := 0 to Formats.Count -1 do
      begin
        log(Formats.Names[f] + '=' + Formats.Values[Formats.Names[f]]);
      end;

      log('ERROR CODES:');
      log('200 : Invalid File Format specified');
      halt(2);
    end
    else
    begin
      Log('Unknown:' + pstr);
    end;




  end;



end;


procedure TDocumentConverter.Log(Msg: String);
begin
  if OutputLog then
  begin
    ConsoleLog.Log(self, Msg);
  end;
end;

procedure TDocumentConverter.SetInputFile(const Value: String);
begin
  FInputFile := Value;
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

end.
