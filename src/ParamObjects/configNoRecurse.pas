unit configNoRecurse;

interface

uses classes, MainUtils, System.Contnrs, SysUtils,
 baseConfig;

type
TParamNoRecurse = class(TParamLoader)
public

  procedure Load(Converter : TDocumentConverter; Param, Value : String); override;
  function  ShouldDec : Boolean; override;
  class procedure RegisterParameters(List : TStrings);
end;

implementation

{ TParamNoRecurse }

procedure TParamNoRecurse.Load(Converter: TDocumentConverter; Param, Value: String);
begin
        Converter.DoSubDirs := false;

        Converter.LogInfo('Loading files from directory but not subdirectories',CHATTY);

end;

class procedure TParamNoRecurse.RegisterParameters(List: TStrings);
begin

            List.AddPair('--NO-RECURSE', TParamNoRecurse.ClassName, TObject(TParamNoRecurse));
            List.AddPair('--NO-SUBDIR', TParamNoRecurse.ClassName, TObject(TParamNoRecurse));

           List.AddPair('--NO-SUBDIRS', TParamNoRecurse.ClassName, TObject(TParamNoRecurse));


end;


function TParamNoRecurse.ShouldDec: Boolean;
begin
  Result := true;
end;

end.
