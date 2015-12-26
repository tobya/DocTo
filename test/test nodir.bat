REM Explaination 
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

REM If output Dir left out default to input
copy  "%~d0%~p0Inputfiles\pie3.doc"  "%~d0%~p0GeneratedFiles\PieNoOutputTest.doc"
"../exe/docto.exe"  -f "%~d0%~p0GeneratedFiles\PieNoOutputTest.doc"    -T  wdFormatPDF

