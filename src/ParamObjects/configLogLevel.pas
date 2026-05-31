unit configLogLevel;

interface

uses classes, MainUtils, System.Contnrs, SysUtils,
 baseConfig;

type
TParamLogLevel = class(TParamLoader)
public
  procedure RegisterParams(List : TStrings);  override;
  procedure Load(Converter : TDocumentConverter; Param, Value : String); override;
  function  ShouldDec : Boolean; override;

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
    Converter.LogInfo('Config Log Level Sets:' + IntToStr(Converter.LogLevel), Converter.LogLevel);
  end;
end;

procedure TParamLogLevel.RegisterParams(List: TStrings);
begin
  List.Values['-L'] := Self.ClassName;
  List.Values['--LOGLEVEL'] := Self.ClassName;
end;

function TParamLogLevel.ShouldDec: Boolean;
begin
  Result := false;
end;

end.
