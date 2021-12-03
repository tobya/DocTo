REM ---------------------------------
REM Convert All files in Input Directory to PDF
REM ---------------------------------

"../exe/docto.exe"  -f "%~d0%~p0Inputfiles\"  -o "%~d0%~p0GeneratedFiles"    -T  wdFormatPDF -OX pdf  -q

