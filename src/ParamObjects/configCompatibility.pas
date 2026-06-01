unit configCompatibility;

interface

uses classes, MainUtils, System.Contnrs, SysUtils,
 baseConfig;

type
TParamCompatibility = class(TParamLoader)
public

  procedure Load(Converter : TDocumentConverter; Param, Value : String); override;
  function  ShouldDec : Boolean; override;
  class procedure RegisterParameters(List : TStrings);
end;

implementation

{ TParamCompatibility }

procedure TParamCompatibility.Load(Converter: TDocumentConverter; Param, Value: String);
begin
    //
end;

class procedure TParamCompatibility.RegisterParameters(List: TStrings);
begin

            List.AddPair('-C', TParamCompatibility.ClassName, TObject(TParamCompatibility));
            List.AddPair('--COMPATIBILITY', TParamCompatibility.ClassName, TObject(TParamCompatibility));
    
end;


function TParamCompatibility.ShouldDec: Boolean;
begin
  Result := false;
end;

end.
