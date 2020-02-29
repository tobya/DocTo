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

        Result.Successful := false;
        Result.InputFile := fileToConvert;

        // Open file
        ppApp.Presentations.Open(fileToConvert);

        // Save as file and close
        PPApp.ActivePresentation.SaveAs(OutputFileName, OutputFileFormat, false);
        PPApp.ActivePresentation.Save;
        PPApp.ActivePresentation.Close();
        Result.InputFile := fileToConvert;
        Result.Successful := true;
        Result.OutputFile := OutputFilename;

end;

function TPowerPointConverter.FormatsExtensions: TStringList;
begin
  Result := TStringList.create();
end;

function TPowerPointConverter.OfficeAppVersion: String;
begin

end;



end.
