REM XLSTO fffff
REM XLSTO fffff


 REM Load csv files convert to xls
"../exe/docto.exe" -XL  -f "%~d0%~p0Inputfiles\"  -o "%~d0%~p0GeneratedFiles"    -T  20 -OX .txt -FX .csv -L 10

REM Full Dir
  "../exe/docto.exe"  -XL -L 10  -f "%~d0%~p0Inputfiles\"  -o "%~d0%~p0GeneratedFiles"    -T  xlExcel9795 -OX .xls 

