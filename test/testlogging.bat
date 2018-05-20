REM Check Logging
"../exe/docto.exe" -v
"../exe/docto.exe"    -o "%~d0%~p0GeneratedFiles\Pie3.pdf"    -T  wdFormatPDF -f "%~d0%~p0Inputfiles\pie3.doc"  -G -L 10
"../exe/docto.exe"    -o "%~d0%~p0GeneratedFiles\Pie3.pdf"    -T  wdFormatPDF -f "%~d0%~p0Inputfiles\pie3.doc"  -GL outputlog.log -L 10