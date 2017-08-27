REM Explaination 
REM Test runs to check docto scenarios
REM execute docto inserting variables 
REM %~d0 and %~p0 together give the full directory this batch file is executing in.


REM Remove all generated files from output directory that may exist.
 del GeneratedFiles\*.* /q


REM Output Help Text
"../exe/docto.exe" -h




REM Try on Single
"../exe/docto.exe"  -f "%~d0%~p0Inputfiles\CullohillApplePie - Copy.doc"  -o "%~d0%~p0GeneratedFiles\Pie3Single.pdf"    -T  wdFormatPDF -GL Logfile.log -L 10

REM docto -f C:\docto-test\CullohillApplePie - Copy.doc -O c:\docto-test\testspace.pdf -T wdFormatPDF -GL Logfile.log -L 10
