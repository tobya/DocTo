unit PowerPointUtils;

interface

uses Classes, MainUtils, ResourceUtils,  ActiveX, ComObj, WinINet, Variants, sysutils, Types, StrUtils;

type

TPowerPointConverter = Class(TDocumentConverter)
Private
    FPPVersion : String;
    PPApp : OleVariant;

public

    function CreateOfficeApp() : boolean;  override;
    function DestroyOfficeApp() : boolean; override;
    function ExecuteConversion(fileToConvert: String; OutputFilename: String; OutputFileFormat : Integer): TConversionInfo; override;
    function AvailableFormats() : TStringList; override;
    function FormatsExtensions(): TStringList; override;
    function WordConstants: TStringList;
    function OfficeAppVersion() : String; override;
End;


implementation

{ TPowerPointConverter }

function TPowerPointConverter.AvailableFormats: TStringList;
begin

end;



function TPowerPointConverter.CreateOfficeApp: boolean;
begin

  if  VarIsEmpty(PPApp) then
  begin
    PPApp :=  CreateOleObject('PowerPoint.Application');
//    PPApp.Visible := false;
  end;
  Result := true;
end;

function TPowerPointConverter.DestroyOfficeApp: boolean;
begin
  if not VarIsEmpty(PPApp) then
  begin
    PPApp.Quit();
  end;
  Result := true;
end;

function TPowerPointConverter.ExecuteConversion(
                                fileToConvert,
                                OutputFilename: String;
                                OutputFileFormat: Integer): TConversionInfo;
begin
        Result.Successful := false;
        Result.InputFile := fileToConvert;

        PPApp.ActivePresentation.SaveAs(OutputFileName, 17, false);

        Result.InputFile := fileToConvert;
        Result.Successful := true;

end;

function TPowerPointConverter.FormatsExtensions: TStringList;
begin
  Result := TStringList.create();
end;

function TPowerPointConverter.OfficeAppVersion: String;
begin

end;

function TPowerPointConverter.WordConstants: TStringList;
begin

end;

end.
