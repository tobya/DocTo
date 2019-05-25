REM Explanation 
REM For use of FOR type FOR /? at prompt
REM The first part specifies ; lines to be ignored
REM I want the first 2 items (delimited by , or [space] on each line from the file specified.
REM Variables will be %%i (%%j is created automatically)
REM execute docto inserting variables 
REM %~d0 and %~p0 together give the full directory this batch file is executing in.


REM Remove all generated files that may exist.
del GeneratedFiles\*.* /q
pause

"../exe/docto.exe" -h

REM Individually try each format on Test Document
REM FOR /F "eol=; tokens=1,2* delims=, " %%i in (testdata.txt) do "../exe/docto.exe"  --inputfile "%~d0%~p0\Inputfiles\pie3.doc"  --outputfile "%~d0%~p0GeneratedFiles\pie3out_%%i.%%j"  -T  %%i

REM Try on Directory
"../exe/docto.exe"  -f "%~d0%~p0Inputfiles\"  -o "%~d0%~p0GeneratedFiles"    --format  wdFormatPDF --outputextension pdf


REM Try on Single
"../exe/docto.exe"  -f "%~d0%~p0Inputfiles\pie3.doc"  -o "%~d0%~p0GeneratedFiles\Pie3Single.pdf"    --FORMAT  wdFormatPDF

REM Try on Single no output file with Verbose Logging
"../exe/docto.exe" --LOGLEVEL 10 -f "%~d0%~p0Inputfiles\pie3.doc"  -o "%~d0%~p0GeneratedFiles\SingleDir\"    --FORMAT  wdFormatPDF  

REM Try on Single no output file with Verbose Logging
"../exe/docto.exe" -L 10 -f "%~d0%~p0Inputfiles\pie3.doc"  -o "%~d0%~p0GeneratedFiles\SingleDirNoSlash"    -T  wdFormatXMLDocument 



REM Should produce an error
"../exe/docto.exe"  -f "%~d0%~p0Inputfiles\pie3.doc"  -o "%~d0%~p0GeneratedFiles\Pie3.pdf"    -T  wdFormatTestPDF

REM Should produce an error
"../exe/docto.exe"  -f "%~d0%~p0Inputfiles\pie3_doesntexist.doc"  -o "%~d0%~p0GeneratedFiles\Pie3.pdf"    -T  wdFormatPDF

REM Test Webhook
REM *********************************
REM To view visit http://toflidium.com/webhooks/docto/docto_test_values.txt
REM *********************************
"../exe/docto.exe"  -f "%~d0%~p0Inputfiles\pie3.doc"  -o "%~d0%~p0GeneratedFiles\Pie3.pdf"    -T  wdFormatPDF -W http://home.kinchik.ie/webhooks/docto/all.php
"../exe/docto.exe"  -f "%~d0%~p0Inputfiles\pie3.doc"  -o "%~d0%~p0GeneratedFiles\Pie3b.pdf"    -T  wdFormatPDF -W http://home.kinchik.ie/webhooks/docto/all.php?
"../exe/docto.exe"  -f "%~d0%~p0Inputfiles\pie3.doc"  -o "%~d0%~p0GeneratedFiles\Pie3c.pdf"    -T  wdFormatPDF -W http://home.kinchik.ie/webhooks/docto/all.php?Aparam=avalueIhave



REM XLSTO

REM Try on Directory
"../exe/xlsto.exe"  -f "%~d0%~p0Inputfiles\"  -o "%~d0%~p0GeneratedFiles"    -T  xlPDF -OX .pdf
