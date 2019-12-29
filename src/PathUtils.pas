unit PathUtils;
(*************************************************************
Copyright © 2012 Toby Allen (https://github.com/tobya)

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction,
including without limitation the rights to use, copy, modify, merge, publish, distribute, sub-license, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so,
subject to the following conditions:

The above copyright notice, and every other copyright notice found in this software, and all the attributions in every file, and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NON-INFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
****************************************************************)
interface
uses sysutils, Classes;

  type
  TPath = Class
  Private
    FPath : String;
    FParts : TStringList;

    function GetParts: TStringList;
    function GetExt: String;
    function GetFileExists: Boolean;

  public
    Constructor Create(Path : String);
    Destructor Destroy;  override;

    Property Parts : TStringList  read GetParts;
    Property Extension : String read GetExt;
    Property Exists : Boolean read GetFileExists;
    function NewFileName(NewExtension : String) : String;


  End;
implementation

{ TPath }

constructor TPath.Create(Path: String);
begin
    FPath := Path;
end;



destructor TPath.Destroy;
begin
  if Assigned(FParts) then FParts.Free;

  inherited destroy;
  
end;

function TPath.GetExt: String;
begin
   result := ExtractFileExt(FPath);
end;

function TPath.GetFileExists: Boolean;
begin
 Result :=  FileExists(FPath) ;
 
end;

function TPath.GetParts: TStringList;
var
  Drive, Folder, Filename : String;
begin
  Drive := ExtractFileDrive(FPath);
  Folder := ExtractFileDir(FPath);
  Filename := ExtractFileName(FPath);

  if not assigned(FParts) then FParts := TStringlist.Create();

  FParts.Add('Drive=' + Drive + ':\');
  FParts.Add('Folder=' + Folder);
  FParts.Add('Filename=' + Filename);
  FParts.Add('Ext=' + Extension);

  Result := FParts;
end;

function TPath.NewFileName(NewExtension: String): String;
begin

end;

end.
