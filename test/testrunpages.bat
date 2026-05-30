REM Try on Single no output file
cls
"../exe/32/docto.exe" -V

 "../exe/32/docto.exe" -wd -f "%~d0%~p0Inputfiles\Mayonnaise.doc"  -o "%~d0%~p0GeneratedFiles\"    -T  wdFormatPDF --PDF-FROMPAGE 2 --PDF-TOPAGE 3 
 "../exe/32/docto.exe" -wd -f "%~d0%~p0Inputfiles\Mayonnaise.doc"  -o "%~d0%~p0GeneratedFiles\mayo5.pdf"    -T  wdFormatPDF --PDF-FROMPAGE wdExportCurrentPage  
  "../exe/32/docto.exe" -wd -f "%~d0%~p0Inputfiles\Mayonnaise.doc"  -o "%~d0%~p0GeneratedFiles\mayo6.pdf"    -T  wdFormatPDF --PDF-FROMPAGE wdExportSelection  



 "../exe/32/docto.exe"  -f "%~d0%~p0Inputfiles\Mayonnaise.doc"  -o "%~d0%~p0GeneratedFiles\mayo2.pdf"    -T  wdFormatPDF --PDF-FROMPAGE 2  -L 10
 "../exe/32/docto.exe"  -f "%~d0%~p0Inputfiles\Mayonnaise.doc"  -o "%~d0%~p0GeneratedFiles\mayo3.pdf"    -T  wdFormatPDF  --PDF-TOPAGE 2 --PDF-OpenAfterExport 

  ..\exe\docto.exe -XL -f "InputFiles\Week 1 Test with PrintArea.xls" -o "%~d0%~p0GeneratedFiles\" -t xlPDF -l 10 --PDF-FROMPAGE 2 --PDF-TOPAGE 3



