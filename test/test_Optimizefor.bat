REM Try optimize options on large file.
 "../exe/docto.exe"  -f ".\Inputfiles\GEOCADx (1).docx"  -o ".\GeneratedFiles\GEOCADx_OptimizedforPrint.pdf"   -T  wdFormatPDF --PDF-OptimizeFor forPrint    -l 10
 "../exe/docto.exe"  -f ".\Inputfiles\GEOCADx (1).docx"  -o ".\GeneratedFiles\GEOCADx_OptimizedforScreen.pdf"   -T  wdFormatPDF -l 10 --PDF-OptimizeFor forOnScreen

REM Check works with differnt cases
"../exe/docto.exe"  -f ".\Inputfiles\GEOCADx (1).docx"  -o ".\GeneratedFiles\GEOCADx_OptimizedforScreen2.pdf"   -T  wdFormatPDF -l 10 --PDF-OptimizeFor foronscreen
