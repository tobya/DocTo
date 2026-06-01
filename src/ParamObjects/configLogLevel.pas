unit configLogLevel;

interface

uses classes, MainUtils, System.Contnrs, SysUtils,
 baseConfig;

type
TParamLogLevel = class(TParamLoader)
public

  procedure Load(Converter : TDocumentConverter; Param, Value : String); override;
  function  ShouldDec : Boolean; override;
         class procedure RegisterParameters(List : TStrings);
end;

implementation

{ TParamLogLevel }

procedure TParamLogLevel.Load(Converter: TDocumentConverter; Param,
  Value: String);
begin
  if IsNumber(Value) then
  begin
    Converter.LogLevel := StrToInt(Value);
    Converter.LogInfo('Log Level Set To:' + IntToStr(Converter.LogLevel), Converter.LogLevel);
    Converter.LogInfo('TConfigLogLevel Sets:' + IntToStr(Converter.LogLevel), Converter.LogLevel);
  end;
end;

class procedure TParamLogLevel.RegisterParameters(List: TStrings);
begin
  List.AddPair('-L', TParamLogLevel.ClassName, TObject(TParamLogLevel));
  List.AddPair('--LOGLEVEL', TParamLogLevel.ClassName, TObject(TParamLogLevel));
end;


function TParamLogLevel.ShouldDec: Boolean;
begin
  Result := false;
end;

end.
