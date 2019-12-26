unit ResourceUtils;
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

uses Classes;


type

TResourceStrings = class(TStringList)
private
  { private declarations }
protected
  { protected declarations }
public
  { public declarations }
  Constructor Create(ResourceName : String);
  Procedure Load(ResourceName: String);
  Function Exists(Key : String) : Boolean;

published
  { published declarations }
end;


procedure LoadStringListFromResource(const ResName: string;SL : TStringList);

implementation


procedure LoadStringListFromResource(const ResName: string;SL : TStringList);
var
  RS: TResourceStream;
begin
  RS := TResourceStream.Create(HInstance, ResName, 'Text');
  try
    SL.LoadFromStream(RS);
  finally
    RS.Free;
  end;
end;
{ TResourceStrings }



constructor TResourceStrings.Create(ResourceName: String);
begin
  inherited Create;
  Load(ResourceName);
end;

function TResourceStrings.Exists(Key: String): Boolean;
var idx : integer;
begin
  Result := false;
  idx :=  Self.IndexOfName(Key);
  if idx > -1 then
  begin
    Result :=  True;
  end;

end;

procedure TResourceStrings.Load(ResourceName: String);
begin
        LoadStringListFromResource(ResourceName,Self);
end;



end.
