REM Try on Directory
"../exe/docto.exe"  -f "%~d0%~p0Inputfiles\"  -o "%~d0%~p0GeneratedFiles"     -T  wdFormatPDF -OX .pdf -L 10 -W http://toflidium.com/webhooks/docto/webhook_test.php -M true
