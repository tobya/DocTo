REM Try on Directory
"../exe/docto.exe"  -f "%~d0%~p0Inputfiles\"  -o "%~d0%~p0GeneratedFiles"    -T  wdFormatPDF -OX pdf -W http://toflidium.com/webhooks/docto/webhook_test.php -Q

REM Try on Relative Directory
"../exe/docto.exe"  -f "..\test\Inputfiles\"  -o "%~d0%~p0GeneratedFiles"    -T  wdFormatPDF -OX pdf 

