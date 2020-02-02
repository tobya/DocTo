REM Test Webhook
REM ---------------------------------
REM To view visit http://toflidium.com/webhooks/docto/docto_test_values.txt
REM ---------------------------------
"../exe/docto.exe"  -f "%~d0%~p0Inputfiles\"  -o "%~d0%~p0GeneratedFiles\"  -T  wdFormatPDF -W https://toflidium.com/webhooks/docto/webhook_test.php -ox .pdf -L 10 

"../exe/docto.exe" -XL -f "%~d0%~p0Inputfiles\"  -o "%~d0%~p0GeneratedFiles\"  -T  XLTypePDF -W https://toflidium.com/webhooks/docto/webhook_test.php -ox .pdf -L 10 
