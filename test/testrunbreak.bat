REM Try on Single no output file
cls
"../exe/32/docto.exe" -V
"../exe/32/docto.exe"  -L 0 -G -f "%~d0%~p0Inputfiles\pie3.doc"  -o "%~d0%~p0GeneratedFiles\SingleDocFolderCreated5"    -T  wdFormatPDddd

"../exe/32/docto.exe"  -L 10 -G -f  "%~d0%~p0Inputfiles\pie3.doc"  -o "%~d0%~p0GeneratedFiles\SingleDocFolderCreated6"  --wpp  -T  wdFormatRTF

Pause