REM Try on Directory
 "../exe/32/docto.exe"  -f "%~d0%~p0Inputfiles\"  -o "%~d0%~p0GeneratedFiles"     -T  wdFormatPDF -OX .pdf   -M true


REM Try on Directory no recurse
"../exe/32/docto.exe"  -f "%~d0%~p0Inputfiles\"  -o "%~d0%~p0GeneratedFiles"     -T  wdFormatPDF -OX .pdf   -M true --no-subdirs
