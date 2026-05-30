REM Try on Directory
REM "../exe/32/docto.exe"  -f "%~d0%~p0Inputfiles\"  -o "%~d0%~p0GeneratedFiles"    -T  wdFormatPDF -OX pdf -W http://toflidium.com/webhooks/docto/webhook_test.php -Q

"../exe/32/docto.exe" -v
REM 
REM Try on Relative Directory
REM "../exe/32/docto.exe"  -f "..\test\Inputfiles\"  -o "%~d0%~p0GeneratedFiles"    -T  wdFormatPDF -OX pdf 

