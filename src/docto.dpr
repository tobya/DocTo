program docto;
(*************************************************************
Copyright � 2012-2016 Toby Allen (https://github.com/tobya)

Permission is hereby granted, free of charge, to any person obtaining a copy of
this software and associated documentation files (the �Software�), to deal in the
Software without restriction,
including without limitation the rights to use, copy, modify, merge, publish,
distribute, sub-license, and/or sell copies of the Software, and to permit persons
to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice, and every other copyright notice found in this software, and all the attributions in every file, and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED �AS IS�, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED
TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NON-INFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE
OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
****************************************************************)
{$APPTYPE CONSOLE}
{$R 'wdFormats.res' 'wdFormats.rc'}
{$R *.res}
uses
  SysUtils,
  Classes,
  ActiveX,
  WordUtils in 'WordUtils.pas',
  MainUtils in 'MainUtils.pas',
  ResourceUtils in 'ResourceUtils.pas',
  PathUtils in 'PathUtils.pas',
  datamodSSL in 'datamodSSL.pas' {dmSSL: TDataModule};

var
  i : integer;
  paramlist : TStringlist;
  DocConv : TWordDocConverter;
  LogResult : String;
begin

  paramlist := TStringlist.create;

  try
   try
     DocConv := TWordDocConverter.Create;
    try

      for i := 1 to ParamCount do
      begin
       paramlist.Add(ParamStr(i));
      end;

      CoInitialize(nil);
      DocConv.LoadConfig(paramlist);


      LogResult :=  DocConv.Execute;
      DocConv.log( LogResult );

      CoUninitialize;
    finally
      DocConv.free;
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
