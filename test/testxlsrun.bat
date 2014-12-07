REM XLSTO

REM Try on Directory
REM "../exe/xlsto.exe" -L 10  -f "%~d0%~p0Inputfiles\"  -o "%~d0%~p0GeneratedFiles"    -T  xlPDF -OX .pdf
REM convert to csv
REM Try on Single no output file with Verbose Logging
REM "../exe/xlsto.exe" -L 10 -f "%~d0%~p0Inputfiles\Week 1 Test.xls"  -o "%~d0%~p0GeneratedFiles\myfile4742.csv"    -T  xlcsv -OX .csv
REM Full Dir
 "../exe/xlsto.exe"  -f "%~d0%~p0Inputfiles\"  -o "%~d0%~p0GeneratedFiles"    -T  xlcsv -OX .csv

