REM Explanation
REM Test runs to check docto scenarios
REM execute docto inserting variables 
REM %~d0 and %~p0 together give the full directory this batch file is executing in.


REM Remove all generated files from output directory that may exist.
 del GeneratedFiles\*.* /q


REM Output Help Text
"../exe/docto.exe" -h


REM ---------------------------------
REM Main Test : 
REM Loop through each format provided in doctoFormatList and try to convert our
REM test file to each format.
REM ---------------------------------

FOR /F "eol=; tokens=1,2* delims=, " %%i in (doctoFormatList.txt) do "../exe/docto.exe"  -f "%~d0%~p0\Inputfiles\pie3.doc"  -o "%~d0%~p0GeneratedFiles\pie3out_%%i.%%j"  -T  %%i



REM ---------------------------------
REM Convert All files in Input Directory to PDF
REM ---------------------------------

"../exe/docto.exe"  -f "%~d0%~p0Inputfiles\"  -o "%~d0%~p0GeneratedFiles"    -T  wdFormatPDF -OX pdf


REM Try on Single
"../exe/docto.exe"  -f "%~d0%~p0Inputfiles\pie3.doc"  -o "%~d0%~p0GeneratedFiles\Pie3Single.pdf"    -T  wdFormatPDF

REM Try on Single no output file with Verbose Logging
"../exe/docto.exe" -L 10 -f "%~d0%~p0Inputfiles\pie3.doc"  -o "%~d0%~p0GeneratedFiles\SingleDir\"    -T  wdFormatPDF  

REM Try on Single no output file with Verbose Logging
"../exe/docto.exe" -L 10 -f "%~d0%~p0Inputfiles\pie3.doc"  -o "%~d0%~p0GeneratedFiles\SingleDirNoSlash"    -T  wdFormatXMLDocument 



REM Should produce an error incorrect format.
"../exe/docto.exe"  -f "%~d0%~p0Inputfiles\pie3.doc"  -o "%~d0%~p0GeneratedFiles\Pie3.pdf"    -T  wdFormatTestPDF

REM Should produce an error - input file does not exist
"../exe/docto.exe"  -f "%~d0%~p0Inputfiles\pie3_doesntexist.doc"  -o "%~d0%~p0GeneratedFiles\Pie3.pdf"    -T  wdFormatPDF

REM Test http Webhook
REM ---------------------------------
REM To view visit https://toflidium.com/webhooks/docto/docto_test_values.txt
REM ---------------------------------
"../exe/docto.exe"  -f "%~d0%~p0Inputfiles\pie3.doc"  -o "%~d0%~p0GeneratedFiles\Pie3.pdf"    -T  wdFormatPDF -W http://toflidium.com/webhooks/docto/webhook_test.php
REM Check https webhook.

"../exe/docto.exe"  -f "%~d0%~p0Inputfiles\pie3.doc"  -o "%~d0%~p0GeneratedFiles\Pie3.pdf"    -T  wdFormatPDF -W https://toflidium.com/webhooks/docto/webhook_test.php


REM If output Dir left out default to input
copy  "%~d0%~p0Inputfiles\pie3.doc"  "%~d0%~p0GeneratedFiles\PieNoOutputTest.doc"
"../exe/docto.exe"  -f "%~d0%~p0GeneratedFiles\PieNoOutputTest.doc"    -T  wdFormatPDF

REM Check that works with -o before -f
"../exe/docto.exe"    -o "%~d0%~p0GeneratedFiles\Pie3.pdf"    -T  wdFormatPDF -f "%~d0%~p0Inputfiles\pie3.doc" 

REM Check Unicode to txt conversion. issue #32
"../exe/docto.exe"  -f "%~d0%~p0Inputfiles\UnicodeTest.doc"  -o "%~d0%~p0GeneratedFiles\UnicodeTest.txt"    -T  wdFormatEncodedText -E 65001 

REM Check Logging
"../exe/docto.exe"    -o "%~d0%~p0GeneratedFiles\Pie3.pdf"    -T  wdFormatPDF -f "%~d0%~p0Inputfiles\pie3.doc"  -G -L 10
"../exe/docto.exe"    -o "%~d0%~p0GeneratedFiles\Pie3.pdf"    -T  wdFormatPDF -f "%~d0%~p0Inputfiles\pie3.doc"  -GL outputlog.log -L 10

