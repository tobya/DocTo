REM Test Webhook
REM ---------------------------------
REM To view visit http://toflidium.com/webhooks/docto/docto_test_values.txt
REM ---------------------------------
"../exe/docto.exe"  -f "%~d0%~p0Inputfiles\pie3.doc"  -o "%~d0%~p0GeneratedFiles\pie3.pdf"  -T  wdFormatPDF -W https://toflidium.com/webhooks/docto/webhook_test.php -L 10 
