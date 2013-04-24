program XlsTo;

{$APPTYPE CONSOLE}
{$R 'xlsFormats.res' 'xlsFormats.rc'}
{$R *.res}
uses
  SysUtils,
  Classes,
  ActiveX,
  ExcelUtils in 'ExcelUtils.pas',
  MainUtils in 'MainUtils.pas',
  ResourceUtils in 'ResourceUtils.pas',
  PathUtils in 'PathUtils.pas';

var
  i : integer;
  paramlist : TStringlist;
  DocConv : TExcelXLSConverter;
  LogResult : String;
begin

  paramlist := TStringlist.create;
  DocConv := TExcelXLSConverter.Create;
  try
  try

  for i := 1 to ParamCount do
  begin
     paramlist.Add(ParamStr(i));
  end;

   DocConv.LoadConfig(paramlist);

   CoInitialize(nil);
   LogResult :=  DocConv.Execute;
   DocConv.log( LogResult );

   CoUninitialize;


  finally
    paramlist.Free;
  end;
  except on E: Exception do
    WriteLn('Error:' + E.Message);
  end;



end.
