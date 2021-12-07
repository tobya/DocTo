unit baseConfig;

interface

uses Classes,System.Contnrs,

      MainUtils;

type

TParamLoader = class
  private
    fParamID : string;
  protected
    function GetParamID: String; virtual;
    procedure SetParamID(const Value: String); virtual;


public
  procedure RegisterParams(List : TStrings); virtual; abstract;
  procedure Load(Converter : TDocumentConverter; Param, Value : String); virtual; abstract;
  function  ShouldDec : Boolean; virtual; abstract;
  Property ParamID : String read GetParamID Write SetParamID;

end;


implementation

{ TParamLoader }

function TParamLoader.GetParamID: String;
begin
  result := FParamID;
end;

procedure TParamLoader.SetParamID(const Value: String);
begin
  FParamID := Value;
end;

end.
