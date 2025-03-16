REM If output Dir left out default to input
REM copy  ".\Inputfiles\pie3.doc"  ".\GeneratedFiles\PieNoOutputTest.doc"
REM Try on Single

del .\GeneratedFiles\*.* /Q
"../exe/32/docto.exe" -PP -f ".\Inputfilespp\"  -o ".\GeneratedFiles\"  -ox ".pdf"  -T  ppSaveAsPDF -l 10