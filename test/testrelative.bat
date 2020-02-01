REM If output Dir left out default to input
REM copy  ".\Inputfiles\pie3.doc"  ".\GeneratedFiles\PieNoOutputTest.doc"
REM Try on Single
"../exe/docto.exe"  -f ".\Inputfiles\pie3.doc"  -o ".\GeneratedFiles\Pie3Single.rtf"    -T  wdFormatRTF -l 10