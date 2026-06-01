unit config{{$name}};

interface

uses classes, MainUtils, System.Contnrs, SysUtils,
 baseConfig;

type
TParam{{$name}} = class(TParamLoader)
public

  procedure Load(Converter : TDocumentConverter; Param, Value : String); override;
  function  ShouldDec : Boolean; override;
  class procedure RegisterParameters(List : TStrings);
end;

implementation

{ TParam{{$name}} }

procedure TParam{{$name}}.Load(Converter: TDocumentConverter; Param, Value: String);
begin
    //
end;

class procedure TParam{{$name}}.RegisterParameters(List: TStrings);
begin

    @foreach($paramlist as $param )
        List.AddPair('{{$param}}', TParam{{$name}}.ClassName, TObject(TParam{{$name}}));
    @endforeach

end;


function TParam{{$name}}.ShouldDec: Boolean;
begin
  Result := false;
end;

end.
