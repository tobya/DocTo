REM If output Dir left out default to input
REM copy  ".\Inputfiles\pie3.doc"  ".\GeneratedFiles\PieNoOutputTest.doc"
REM Try on Single
 "../exe/docto.exe"  -f ".\Inputfiles\pie3.doc"  -o ".\GeneratedFiles\\Pie3Single.rtf"    -T  wdFormatRTF -l 10

"../exe/docto.exe"  -f ".\Inputfiles\pie3.doc"  -o "..\test\GeneratedFiles234\Pie3Single.rtf"    -T  wdFormatRTF -l 10

 "../exe/docto.exe"  -f ".\Inputfiles\pie3.doc"  -o "..\test\Generatednewdir\Pie3Single.rtf"    -T  wdFormatRTF -l 10


"../exe/docto.exe"  -f ".\Inputfiles\"  -o "..\test\Generatednewdir2"    -T  wdFormatRTF -l 10
