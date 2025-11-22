unit DynamicFileNameGenerator;

interface

uses Classes, MainUtils, ResourceUtils,  ActiveX, ComObj, WinINet, Variants, sysutils, Types, StrUtils;

type

TDynamicFileNameGenerator = Class(TObject)
  private
    FIgnoreTag: Boolean;
    procedure SetIgnoreTag(const Value: Boolean);
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
        property IgnoreTag : Boolean read FIgnoreTag write SetIgnoreTag;
End;


implementation

{ TDynamicFileNameGenerator }

constructor TDynamicFileNameGenerator.Create(FileName: String);
begin

        FIgnoreTag := false;
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
       if FIgnoreTag then
       begin
                result := fOrigionalFilename;
       end else
       begin
                result := dynamicoutputDir + dynamicoutputFile + '_(' + tag +  ')' + dynamicoutputExt;
       end;

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

procedure TDynamicFileNameGenerator.SetIgnoreTag(const Value: Boolean);
begin
  FIgnoreTag := Value;
end;

end.
