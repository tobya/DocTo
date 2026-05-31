unit configFormat;

interface

uses classes, MainUtils, System.Contnrs, SysUtils,
 baseConfig;

type
TParamFormat = class(TParamLoader)
public
  procedure RegisterParams(List : TStrings);  override;
  procedure Load(Converter : TDocumentConverter; Param, Value : String); override;
  function  ShouldDec : Boolean; override;
  class procedure RegisterParameters(List : TStrings);
end;

implementation

{ TParamFormat }

procedure TParamFormat.Load(Converter: TDocumentConverter; Param, Value: String);
var
  ForceFormat : Boolean;
  FormatInt   : Integer;
begin
  ForceFormat := (Param = '-TF') or (Param = '--FORCEFORMAT');

  if IsNumber(Value) then
  begin
    FormatInt := StrToInt(Value);
    Converter.OutputFileFormat := FormatInt;
    if (not ForceFormat) and (not Converter.IsValidFormat(FormatInt)) then
    begin
      Converter.HaltWithConfigError(200, 'File Format ' + Value +
        ' is invalid, please see help. -h.  To force use, use -TF');
    end;
  end
  else  // string format such as 'wdFormatPDF', 'xlCSV'
  begin
    Converter.OutputFileFormatString := Value;
    FormatInt := Converter.LookupFormatByName(Value);
    if FormatInt > -1 then
    begin
      Converter.OutputFileFormat := FormatInt;
    end
    else
    begin
      Converter.HaltWithConfigError(200, 'File Format ' + Value +
        ' is an invalid ' + Converter.OfficeAppName + ' file extension , please see help. -h');
    end;
  end;

  Converter.LogDebug('Type Integer is: ' + IntToStr(Converter.OutputFileFormat), VERBOSE);
end;

class procedure TParamFormat.RegisterParameters(List: TStrings);
begin
  List.AddPair('-T',            TParamFormat.ClassName, TObject(TParamFormat));
  List.AddPair('--FORMAT',      TParamFormat.ClassName, TObject(TParamFormat));
  List.AddPair('-TF',           TParamFormat.ClassName, TObject(TParamFormat));
  List.AddPair('--FORCEFORMAT', TParamFormat.ClassName, TObject(TParamFormat));
end;

procedure TParamFormat.RegisterParams(List: TStrings);
begin

end;

function TParamFormat.ShouldDec: Boolean;
begin
  Result := false;
end;

end.
