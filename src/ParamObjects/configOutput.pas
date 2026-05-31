unit configOutput;

interface

uses classes, MainUtils,System.Contnrs,
 baseConfig;

type
TParamOutputExtension = class(TParamLoader)
public

  procedure Load(Converter : TDocumentConverter; Param, Value : String);   override;
  function  ShouldDec : Boolean;   override;

 class procedure RegisterParameters(List : TStrings);

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

class procedure TParamOutputExtension.RegisterParameters(List: TStrings);
begin
  List.AddPair('-OX',TParamOutputExtension.Classname, TObject(TParamOutputExtension));
  List.AddPair('--OUTPUTEXTENSION',TParamOutputExtension.Classname, TObject(TParamOutputExtension));

end;



function TParamOutputExtension.ShouldDec: Boolean;
begin
    Result := false;
end;

end.
