program docto;
(*************************************************************
Copyright © 2012-2016 Toby Allen (https://github.com/tobya)

Permission is hereby granted, free of charge, to any person obtaining a copy of
this software and associated documentation files (the “Software”), to deal in the
Software without restriction,
including without limitation the rights to use, copy, modify, merge, publish,
distribute, sub-license, and/or sell copies of the Software, and to permit persons
to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice, and every other copyright notice found in this software, and all the attributions in every file, and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED
TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NON-INFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE
OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
****************************************************************)
{$APPTYPE CONSOLE}
{$R 'ExtraFiles.res' 'ExtraFiles.rc'}
{$R *.res}
uses
  SysUtils,
  Classes,
  ActiveX,
  WordUtils in 'WordUtils.pas',
  MainUtils in 'MainUtils.pas',
  ResourceUtils in 'ResourceUtils.pas',
  PathUtils in 'PathUtils.pas',
  datamodSSL in 'datamodSSL.pas' {dmSSL: TDataModule},
  ExcelUtils in 'ExcelUtils.pas',
  PowerPointUtils in 'PowerPointUtils.pas',
  Word_TLB_Constants in 'Word_TLB_Constants.pas',
  Excel_TLB_Constants in 'Excel_TLB_Constants.pas',
  PowerPoint_TLB_Constants in 'PowerPoint_TLB_Constants.pas',
  VisioUtils in 'VisioUtils.pas',
  Visio_TLB in 'Visio_TLB.pas';

var
  i, Converter : integer;
  paramlist : TStringlist;
  DocConv : TWordDocConverter;
  XLSConv : TExcelXLSConverter;
  PPConv : TPowerPointConverter;
  VSConv : TVisioConverter;
  LogResult : String;
begin

  paramlist := TStringlist.create;

  try
   try
     // Due to the way config is loaded I must create all of them.
     DocConv := TWordDocConverter.Create;
     XLSConv := TExcelXLSConverter.Create;
     PPConv :=  TPowerPointConverter.Create;
     VSConv :=  TVisioConverter.Create;
    try

      for i := 1 to ParamCount do
      begin
       paramlist.Add(ParamStr(i));
      end;

      CoInitialize(nil);

      Converter := DocConv.ChooseConverter(ParamList);

      //DocConv.LogVersionInfo;
      if Converter = MSWord then
      begin
        DocConv.Log('Converter:MS Word' ,CHATTY);
        DocConv.LoadConfig(paramlist);
        LogResult :=  DocConv.Execute;
        DocConv.log( LogResult );
      end
      else if Converter = MSEXCEL then
      begin
        // Config Level already set for DocConv but not for XLS.
        XLSConv.ChooseConverter(ParamList);
        XLSConv.Log('Converter:MS Excel' ,CHATTY);
        XLSConv.LoadConfig(ParamList);
        LogResult :=  XLSConv.Execute;
        XLSConv.log( LogResult );
      end
      ELSE if Converter = MSPOWERPOINT then
      begin
        PPConv.ChooseConverter(ParamList);
        PPConv.Log('Converter:MS Powerpoint' ,CHATTY);
        PPConv.LoadConfig(ParamList);
        LogResult := PPConv.Execute;
        PPConv.Log(LogResult);
      end
     ELSE if Converter = MSVISIO then
      begin
        VSConv.ChooseConverter(ParamList);
        VSConv.Log('Converter:MS Visio' ,CHATTY);
        VSConv.LoadConfig(ParamList);
        LogResult := VSConv.Execute;
        VSConv.Log(LogResult);
      end;

      CoUninitialize;
    finally
      DocConv.free;
      XLSConv.Free;
      PPConv.Free;

    end;
   except on E: Exception do
    begin
         Writeln('Exiting with Error 400: ' + E.ClassName );
         Writeln(E.Message);
         halt(400);
    end;

   end;
  finally
    paramlist.Free;
  end;

end.
