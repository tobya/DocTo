REM Try on Single no output file
cls
"../exe/docto.exe" -V

 REM "../exe/docto.exe"  -f "%~d0%~p0Inputfiles\Mayonnaise.doc"  -o "%~d0%~p0GeneratedFiles\"    -T  wdFormatPDF --PDF-FROMPAGE 2 --PDF-TOPAGE 3 --PDF-OpenAfterExport 
 "../exe/docto.exe"  -f "%~d0%~p0Inputfiles\Mayonnaise.doc"  -o "%~d0%~p0GeneratedFiles\mayo2.pdf"    -T  wdFormatPDF --PDF-FROMPAGE 2  --PDF-OpenAfterExport -L 10
 REM "../exe/docto.exe"  -f "%~d0%~p0Inputfiles\Mayonnaise.doc"  -o "%~d0%~p0GeneratedFiles\mayo3.pdf"    -T  wdFormatPDF  --PDF-TOPAGE 2 --PDF-OpenAfterExport 

 REM ..\exe\docto.exe -XL -f "InputFiles\Week 1 Test with PrintArea.xls" -o "%~d0%~p0GeneratedFiles\" -t xlPDF -l 10 --PDF-FROMPAGE 2 --PDF-TOPAGE 3



