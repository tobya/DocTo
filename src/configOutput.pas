unit configOutput;

interface

uses classes, MainUtils,System.Contnrs,
 baseConfig;

type
TParamOutput = class(TParamLoader)
public
  procedure RegisterParams(List : TStrings);
  procedure Load(Converter : TDocumentConverter; Param, Value : String);
  function  ShouldDec : Boolean;

end;

implementation

{ TParamOutput }

procedure TParamOutput.Load(Converter: TDocumentConverter; Param,
  Value: String);
begin

     //If the first character isn't . add it.
     if value[1] = '.' then
     begin
        Converter.OutputExt := value;
     end
     else
     begin
       Converter.OutputExt := '.' + value;
     end;


end;

procedure TParamOutput.RegisterParams(List: TStrings);
begin
  List.Values['-OX'] := Self.ClassName;
    List.Values['--OUTPUTEXTENSION'] := Self.ClassName;
end;

function TParamOutput.ShouldDec: Boolean;
begin
    Result := false;
end;

end.
