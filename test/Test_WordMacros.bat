REM If output Dir left out default to input
REM copy  ".\Inputfiles\pie3.doc"  ".\GeneratedFiles\PieNoOutputTest.doc"
REM Try on Single
 "../exe/docto.exe"  -f ".\Inputfiles\DocWithMacro.docm"  -o ".\GeneratedFiles\DocWithMacro.pdf"    -T  wdFormatPDF -l 10

"../exe/docto.exe"  -f ".\Inputfiles\DocWithMacro_oldext.doc"  -o ".\GeneratedFiles\DocWithMacro_oldext.pdf"    -T  wdFormatPDF -l 10


 "../exe/docto.exe"  -f ".\Inputfiles\DocWithMacro.docm"  -o ".\GeneratedFiles\DocWithMacro2.pdf"    -T  wdFormatPDF -l 10 --enable-macroautorun

"../exe/docto.exe"  -f ".\Inputfiles\DocWithMacro_oldext.doc"  -o ".\GeneratedFiles\DocWithMacro_oldext3.pdf"    -T  wdFormatPDF -l 10 

REM Raises an error. 301

"../exe/docto.exe" -XL  -f ".\inputfilesxl\Week 1 Test.xls"  -o "C:\dev\github\docto\test\GeneratedFiles"    -T  xlPDF -l 10 --enable-macroautorun

