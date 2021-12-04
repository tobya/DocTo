unit PowerPointUtils;

interface

uses Classes, MainUtils, ResourceUtils,  ActiveX, ComObj, WinINet, Variants, sysutils, Types, StrUtils;

type

TPowerPointConverter = Class(TDocumentConverter)
Private
    FPPVersion : String;
    PPApp : OleVariant;

public
             constructor Create();
    function CreateOfficeApp() : boolean;  override;
    function DestroyOfficeApp() : boolean; override;
    function ExecuteConversion(fileToConvert: String; OutputFilename: String; OutputFileFormat : Integer): TConversionInfo; override;
    function AvailableFormats() : TStringList; override;
    function FormatsExtensions(): TStringList; override;
    function OfficeAppVersion() : String; override;
End;


implementation

{ TPowerPointConverter }

function TPowerPointConverter.AvailableFormats: TStringList;
begin
  Formats := TResourceStrings.Create('PPFORMATS');
  result := Formats;

end;



constructor TPowerPointConverter.Create;
begin
inherited;
  //setup defaults
  InputExtension := '.ppt*';
  OfficeAppName := 'Powerpoint';
end;

function TPowerPointConverter.CreateOfficeApp: boolean;
begin

  if  VarIsEmpty(PPApp) then
  begin
    PPApp :=  CreateOleObject('PowerPoint.Application');
 //      ppApp.Visible := 0;
  end;
  Result := true;
end;

function TPowerPointConverter.DestroyOfficeApp: boolean;
begin
  if not VarIsEmpty(PPApp) then
  begin
    PPApp.Quit;
  end;
  Result := true;
end;

function TPowerPointConverter.ExecuteConversion(
                                fileToConvert,
                                OutputFilename: String;
                                OutputFileFormat: Integer): TConversionInfo;
begin
        logDebug('Trying to Convert: ' + fileToConvert, Debug);

        Result.Successful := false;
        Result.InputFile := fileToConvert;
        // Open file
        ppApp.Presentations.Open(fileToConvert);

        // Save as file and close
        PPApp.ActivePresentation.SaveAs(OutputFileName, OutputFileFormat, false);

        PPApp.ActivePresentation.Save;

        PPApp.ActivePresentation.Close;

        Result.InputFile := fileToConvert;
        Result.Successful := true;
        Result.OutputFile := OutputFilename;

end;

function TPowerPointConverter.FormatsExtensions: TStringList;
var
  Extensions : TStringList;

begin
  Extensions := Tstringlist.Create();
  LoadStringListFromResource('PPEXTENSIONS',Extensions);

  result := Extensions;
end;

function TPowerPointConverter.OfficeAppVersion(): String;
begin
  FPPVersion := ReadOfficeAppVersion;
  if FPPVersion = '' then
  begin
  CreateOfficeApp();
  FPPVersion := PPApp.Version;
  WriteOfficeAppVersion(FPPVersion);
  end;
  result := FPPVersion;
end;



end.              
