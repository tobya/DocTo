unit configOutput;

interface

uses classes, MainUtils,System.Contnrs,
 baseConfig;

type
TParamOutputExtension = class(TParamLoader)
public
  procedure RegisterParams(List : TStrings);   override;
  procedure Load(Converter : TDocumentConverter; Param, Value : String);   override;
  function  ShouldDec : Boolean;   override;

end;

implementation

{ TParamOutput }

procedure TParamOutputExtension.Load(Converter: TDocumentConverter; Param,
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

procedure TParamOutputExtension.RegisterParams(List: TStrings);
begin
  List.Values['-OX'] := Self.ClassName;
    List.Values['--OUTPUTEXTENSION'] := Self.ClassName;
end;

function TParamOutputExtension.ShouldDec: Boolean;
begin
    Result := false;
end;

end.
