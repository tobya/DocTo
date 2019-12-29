REM Check Logging
"../exe/docto.exe" -v
"../exe/docto.exe"  -f "%~d0%~p0Inputfiles2\pie3.doc"   -o "%~d0%~p0GeneratedFiles\Pie3.pdf"    -T  wdFormatPDF -R -G -L 10  
cd "%~d0%~p0Inputfiles2\"
dir
Pause