REM Test Webhook
REM ---------------------------------
REM To view visit https://toflidium.com/webhooks/docto/docto_test_values.txt
REM ---------------------------------
"../exe/docto.exe"  -f "%~d0%~p0Inputfiles\pie3.doc"  -o "%~d0%~p0GeneratedFiles\Pie3.pdf"    -T  wdFormatPDF -W https://webhook.azure.cookingisfun.ie/doctotest.php -L 10