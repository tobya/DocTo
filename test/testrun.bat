REM Explaination 
REM For use of FOR type FOR /? at prompt
REM The first part specifies ; lines to be ignored
REM I want the first 2 items (delimited by , or [space] on each line from the file specified.
REM Variables will be %%i (%%j is created automatically)
REM execute docto inserting variables 
REM %~d0 and %~p0 together give the full directory this batch file is executing in.

FOR /F "eol=; tokens=1,2* delims=, " %%i in (testdata.txt) do "../exe/docto.exe"  -f "%~d0%~p0pie3.doc"  -o "%~d0%~p0GeneratedFiles\pie3out_%%i.%%j"  -T  %%i
