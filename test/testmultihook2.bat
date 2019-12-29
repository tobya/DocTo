REM Test Webhook
REM ---------------------------------
REM To view visit http://toflidium.com/webhooks/docto/docto_test_values.txt
REM ---------------------------------

REM Call webhook for each conversion.
FOR /F "eol=; tokens=1,2* delims=, " %%i in (doctoFormatList.txt) do "../exe/docto.exe"  -f "%~d0%~p0\Inputfiles\pie3.doc"  -o "%~d0%~p0GeneratedFiles\pie3out_%%i.%%j"  -T  %%i  -W https://toflidium.com/webhooks/docto/webhook_test.php -L 10

