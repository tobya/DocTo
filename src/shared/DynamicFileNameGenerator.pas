unit DynamicFileNameGenerator;

interface

uses Classes, MainUtils, ResourceUtils,  ActiveX, ComObj, WinINet, Variants, sysutils, Types, StrUtils;

type

TDynamicFileNameGenerator = Class(TObject)
protected

        fOrigionalFilename : string;
        dynamicoutputDir   :string;
         dynamicoutputFile :string;
         dynamicoutputExt  :string;
        fFilenameParts : TStringList;
        procedure Deconstruct();
public
        Constructor Create(FileName: String);
        function Generate(tag: String ) : String;
            function SafeFileName(FileName: String ) : String;
End;


implementation

{ TDynamicFileNameGenerator }

constructor TDynamicFileNameGenerator.Create(FileName: String);
begin

        fOrigionalFilename := Filename;
        self.Deconstruct;
end;

procedure TDynamicFileNameGenerator.Deconstruct;


begin
                 dynamicoutputDir := ExtractFilePath(fOrigionalFilename); // includes last \
                 dynamicoutputFile :=  ChangefileExt ( ExtractFileName(fOrigionalFilename),'');
                 dynamicoutputExt  := ExtractFileExt(fOrigionalFilename);


end;

function TDynamicFileNameGenerator.Generate(tag: String): String;
var
dynamicFileName : String;
dynamicOutputFilename : string;

begin
       dynamicFileName := SafeFileName(tag);
       result := dynamicoutputDir + dynamicoutputFile + '_(' + tag +  ')' + dynamicoutputExt;

end;

function TDynamicFileNameGenerator.SafeFileName(FileName: String): String;
begin
  Filename :=  StringReplace(Filename,'&','_',[rfReplaceAll,rfIgnoreCase]);
    Filename :=  StringReplace(Filename,'/','_',[rfReplaceAll,rfIgnoreCase]);
    Filename :=  StringReplace(Filename,'\\','_',[rfReplaceAll,rfIgnoreCase]);
 Filename :=  StringReplace(Filename,'<','_',[rfReplaceAll,rfIgnoreCase]);
  Filename :=  StringReplace(Filename,'>','_',[rfReplaceAll,rfIgnoreCase]);
   Filename :=  StringReplace(Filename,':','_',[rfReplaceAll,rfIgnoreCase]);
    Filename :=  StringReplace(Filename,'?','_',[rfReplaceAll,rfIgnoreCase]);
 Filename :=  StringReplace(Filename,'*','_',[rfReplaceAll,rfIgnoreCase]);
    result := Filename;
end;

end.
