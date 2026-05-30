REM Check Logging
"../exe/32/docto.exe" -v
"../exe/32/docto.exe"  -f "%~d0%~p0Inputfiles2\pie3.doc"   -o "%~d0%~p0GeneratedFiles\Pie3.pdf"    -T  wdFormatPDF -R -G -L 10  
cd "%~d0%~p0Inputfiles2\"
dir
Pause