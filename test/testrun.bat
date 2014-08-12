REM Explaination 
REM For use of FOR type FOR /? at prompt
REM The first part specifies ; lines to be ignored
REM I want the first 2 items (delimited by , or [space] on each line from the file specified.
REM Variables will be %%i (%%j is created automatically)
REM execute docto inserting variables 
REM %~d0 and %~p0 together give the full directory this batch file is executing in.



REM Individually try each format on Test Document
FOR /F "eol=; tokens=1,2* delims=, " %%i in (testdata.txt) do "../exe/docto.exe"  -f "%~d0%~p0\Inputfiles\pie3.doc"  -o "%~d0%~p0GeneratedFiles\pie3out_%%i.%%j"  -T  %%i

REM Try on Directory
"../exe/docto.exe"  -f "%~d0%~p0Inputfiles\"  -o "%~d0%~p0GeneratedFiles"    -T  wdFormatPDF -OX pdf

REM Try on Single
"../exe/docto.exe"  -f "%~d0%~p0Inputfiles\pie3.doc"  -o "%~d0%~p0GeneratedFiles\Pie3Single.pdf"    -T  wdFormatPDF

REM Try on Single no output file
"../exe/docto.exe"  -f "%~d0%~p0Inputfiles\pie3.doc"  -o "%~d0%~p0GeneratedFiles\Single"    -T  wdFormatPDF 


REM Should produce an error
"../exe/docto.exe"  -f "%~d0%~p0Inputfiles\pie3.doc"  -o "%~d0%~p0GeneratedFiles\Pie3.pdf"    -T  wdFormatTestPDF

REM Should produce an error
"../exe/docto.exe"  -f "%~d0%~p0Inputfiles\pie3_doesntexist.doc"  -o "%~d0%~p0GeneratedFiles\Pie3.pdf"    -T  wdFormatPDF

REM Test Webhook
REM *********************************
REM To view visit http://toflidium.com/webhooks/docto/docto_test_values.txt
REM *********************************
"../exe/docto.exe"  -f "%~d0%~p0Inputfiles\pie3.doc"  -o "%~d0%~p0GeneratedFiles\Pie3.pdf"    -T  wdFormatPDF -W http://toflidium.com/webhooks/docto/webhook_test.php

