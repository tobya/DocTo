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

uses Classes, strutils, sysutils;


type


TResourceStrings = class(TStringList)
private
    FValuesAreHex: Boolean;
    FValuesAreInt: Boolean;
    procedure SetValuesAreHex(const Value: Boolean);
    procedure SetValuesAreInt(const Value: Boolean);
    function GetValueasInt(Name: String): Integer;
    procedure SetValueasInt(Name: String; const Value: Integer);
    function GetValuebyIndexasInt(idx: integer): Integer;
    procedure SetValuebyIndexasInt(idx: integer; const Value: Integer);
  { private declarations }
protected
  { protected declarations }
public
  { public declarations }
  Constructor Create(ResourceName : String);
  Procedure Load(ResourceName: String);
  Procedure Append(ResourceName : String);
  Function Exists(Key : String) : Boolean;
  Property ValuesAreHex : Boolean read FValuesAreHex write SetValuesAreHex;
  property ValuesAreInt : Boolean  read FValuesAreInt write SetValuesAreInt;
  Property ValueasInt[Name: String]: Integer read GetValueasInt write SetValueasInt;
  Property ValuebyIndexasInt[idx: integer] : Integer read GetValuebyIndexasInt write SetValuebyIndexasInt;

published
  { published declarations }
end;


procedure LoadStringListFromResource(const ResName: string;SL : TStringList; Append: Boolean = false);

implementation


procedure LoadStringListFromResource(const ResName: string;SL : TStringList; Append: Boolean = false);
var
  RS: TResourceStream;
  tempSL : TStringList;
begin
  tempSL := TStringList.Create();
  RS := TResourceStream.Create(HInstance, ResName, 'Text');
  try
    tempSL.LoadFromStream(RS);
  finally
    RS.Free;
  end;
  if Append then
  begin
    SL.AddStrings(tempSL);
  end else
  begin
    SL.Text := tempSL.Text;
  end;
end;
{ TResourceStrings }



procedure TResourceStrings.Append(ResourceName: String);
begin
       LoadStringListFromResource(ResourceName,Self,true);
end;

constructor TResourceStrings.Create(ResourceName: String);
begin
  inherited Create;
  FValuesAreHex := false;
  FValuesAreInt := false;
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

function TResourceStrings.GetValueasInt(Name: String): Integer;
  var
    val : String;
begin
    val := Self.Values[Name];
    Result := Strtoint(val);
end;

function TResourceStrings.GetValuebyIndexasInt(idx: integer): Integer;
  var
    val : String;
begin
    val := Self.ValueFromIndex[idx];
    Result := Strtoint(val);

end;

procedure TResourceStrings.Load(ResourceName: String);
begin
        LoadStringListFromResource(ResourceName,Self);
end;



procedure TResourceStrings.SetValueasInt(Name: String; const Value: Integer);
begin

end;

procedure TResourceStrings.SetValuebyIndexasInt(idx: integer;
  const Value: Integer);
begin

end;

procedure TResourceStrings.SetValuesAreHex(const Value: Boolean);
begin
  FValuesAreHex := Value;
end;

procedure TResourceStrings.SetValuesAreInt(const Value: Boolean);
begin
  FValuesAreInt := Value;
end;

end.
