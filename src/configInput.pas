unit configInput;

interface

uses classes, MainUtils,System.Contnrs, Sysutils,
 baseConfig;

type
TParamInput = class(TParamLoader)
public
  procedure RegisterParams(List : TStrings);
  procedure Load(Converter : TDocumentConverter; Param, Value : String);
  function  ShouldDec : Boolean;

end;

implementation

{ TParamOutput }

procedure TParamInput.Load(Converter: TDocumentConverter; Param,
  Value: String);
var tmppath : string;
begin

      // Before doing anything else expand file name to remove any relative paths.
      Converter.InputFile := ExpandFileName(Value);

      Converter.logdebug('Input File is: ' + Converter.InputFile,CHATTY);

      tmppath := ExtractFilePath(Converter.InputFile);

        // If we are given a filename with no path, get currentdir and add to file.
        if (tmppath = '') then
        begin
          tmppath := GetCurrentDir();
          tmppath := IncludeTrailingBackslash(tmppath);
          Converter.InputFile := tmppath + Converter.InputFile;
        end;

      if (FileExists(Converter.InputFile) = false) and (DirectoryExists(Converter.InputFile) = false) then
      begin
        Converter.HaltWithError(204,'EInput file ' + Converter.InputFile + ' does not exist.');
      end
      else if (FileExists(Converter.InputFile)) then
      begin
        Converter.InputIsFile := true;
        Converter.InputIsDir := false;
      end
      else if (DirectoryExists(Converter.InputFile)) then
      begin
        Converter.InputIsFile := false;
        Converter.InputIsDir := true;

        // Create Absolute path from any relative path
        Converter.InputFile := ExpandFileName(Converter.InputFile);
      end;


      // Set Output Directory to Input Directry at this stage. This ensure if no
      // output directory  (-o) is specified, then it will default to same as
      // input dir. If output has been supplied as param it will overwrite later.
      if Converter.OutputFile = '' then
      begin
        Converter.OutputFile :=  IncludeTrailingBackslash(tmppath);
        Converter.OutputIsDir := true;
      end;




end;

procedure TParamInput.RegisterParams(List: TStrings);
begin
    List.Values['-F'] := Self.ClassName;
    List.Values['--INPUTFILE'] := Self.ClassName;
end;

function TParamInput.ShouldDec: Boolean;
begin
    Result := false;
end;

end.
