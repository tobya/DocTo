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
REM Use a relative path
REM ---------------------------------

FOR /F "eol=; tokens=1,2* delims=, " %%i in (doctoFormatList.txt) do "../exe/docto.exe"  -f "Inputfiles\pie3.doc"  -o "GeneratedFiles\pie3out_%%i.%%j"  -T  %%i
