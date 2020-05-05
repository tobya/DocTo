REM ---------------------------------
REM Main Test : 
REM Loop through each format provided in doctoFormatList and try to convert our
REM test file to each format.
REM ---------------------------------

FOR /F "eol=; tokens=1,2* delims=, " %%i in (ppdoctoFormatList.txt) do "../exe/docto.exe" -PP  -f "%~d0%~p0\Inputfilespp\Presentation1.ppt"  -o "%~d0%~p0GeneratedFiles\pres_%%i.%%j"  -T  %%i -L 10


