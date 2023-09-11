REM XLSTO fffff


 REM Load csv files convert to xls
REM "../exe/docto.exe" -XL  -f "%~d0%~p0Inputfiles\"  -o "%~d0%~p0GeneratedFiles"    -T  20 -OX .txt -FX .csv -L 10
"../exe/docto.exe" -XL  -f "%~d0%~p0Inputfiles\"  -o "%~d0%~p0GeneratedFiles\"    -T  xlPDF   -L 10

REM Full Dir
  "../exe/docto.exe"  -XL -L 10  -f "%~d0%~p0Inputfiles\"  -o "%~d0%~p0GeneratedFiles"    -T  xlHtml -OX .html
  "../exe/docto.exe"  -XL -L 10  -f "%~d0%~p0Inputfiles\"  -o "%~d0%~p0GeneratedFiles"    -T  xlExcel7 -OX .xls97 

