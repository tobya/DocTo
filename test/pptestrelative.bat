REM If output Dir left out default to input
REM copy  ".\Inputfiles\pie3.doc"  ".\GeneratedFiles\PieNoOutputTest.doc"
REM Try on Single
"../exe/docto.exe" -PP -f ".\Inputfilespp\Presentation1.ppt"  -o ".\GeneratedFiles\res3.jpg"    -T  17 -l 10
"../exe/docto.exe" -PP -f ".\Inputfilespp\Presentation1.ppt"  -o ".\GeneratedFiles\res3.pdf"    -T  ppSaveAsPDF -l 10