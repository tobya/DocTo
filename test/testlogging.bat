REM Check Logging
"../exe/32/docto.exe" -v
"../exe/32/docto.exe"    -o "%~d0%~p0GeneratedFiles\Pie3.pdf"    -T  wdFormatPDF -f "%~d0%~p0Inputfiles\pie3.doc"  -G -L 10
"../exe/32/docto.exe"    -o "%~d0%~p0GeneratedFiles\Pie3.pdf"    -T  wdFormatPDF -f "%~d0%~p0Inputfiles\pie3.doc"  -GL outputlog.log -L 10