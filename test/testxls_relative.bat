 REM Load xls from a relative path and convert to csv
"../exe/docto.exe" -XL  -f "inputfilesxl\Week 1 Test.xls"  -o "GeneratedFiles\Week1.csv"    -T  xlCSV -l 10

"../exe/docto.exe" -XL  -f "inputfilesxl\"  -o "..\Test\GeneratedTestputFiles/Week3"    -T  xlCSV -l 10
