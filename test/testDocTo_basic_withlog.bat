REM Explanation
REM Test runs to check docto scenarios
REM execute docto inserting variables 
REM %~d0 and %~p0 together give the full directory this batch file is executing in.


REM Remove all generated files from output directory that may exist.
 del GeneratedFiles\*.* /q


REM Output Help Text
"../exe/32/docto.exe" -h


REM Try on Single
"../exe/32/docto.exe" -wd  -f "%~d0%~p0Inputfiles\pie3.doc"  -o "%~d0%~p0GeneratedFiles\Pie3Single.pdf"    -T  wdFormatPDF -G 

REM Try on Single no output file with Verbose Logging
"../exe/32/docto.exe" -WD  -f "%~d0%~p0Inputfiles\pie3.doc"  -o "%~d0%~p0GeneratedFiles\SingleDir\"    -T  wdFormatText  -G -L 2 


REM try xl
"../exe/32/docto.exe" -XL -f "%~d0%~p0Inputfilesxl\Week 1 Test.xls"  -o "%~d0%~p0GeneratedFiles\Week1.pdf"    -T  XLPDF -G -L 2

REM try xl
"../exe/32/docto.exe" -XL -f "%~d0%~p0Inputfilesxl\Week 1 Test.xls"  -o "%~d0%~p0GeneratedFiles\Week1.csv"    -T  XLcsv  -G -L 2

REM try xl
"../exe/32/docto.exe" -PP -f "%~d0%~p0Inputfilespp\Presentation1.ppt"  -o "%~d0%~p0GeneratedFiles\pres1.pdf"    -T  ppSaveasPDF -G -L 2

REM try xl
"../exe/32/docto.exe" -PP -f "%~d0%~p0Inputfilespp\Presentation1.ppt"  -o "%~d0%~p0GeneratedFiles\pres1.rtf"    -T  ppSaveasRTF   -G -L 2

