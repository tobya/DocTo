{
    "title" : "How do I Convert a Microsoft Excel Spreadsheet to a PDF and save it to a UNC network Drive?  " 
}

How do I Convert a Microsoft Excel Spreadsheet to a PDF and save it to a UNC network Drive? 
==

How do I Convert a Microsoft Excel Spreadsheet to a Adobe PDF Format (PDF) and write it to a UNC Path?         
-

It is very simple to convert a Microsoft Excel Spreadsheet to a PDF file  on the command line using [Docto](https://github.com/tobya/docto). You can also do this easily in Microsoft Excel, but sometimes it helps to be able to do it from the command line.  

If you want to save your output to a UNC `\\Server` path, there is one extra trick you need to do.  You need to double up the slashes (\\) for the output file or Directory.  At the begining of the file path there is usually 2 slashes, you need to put 4. 

The command line below shows how you can convert a Microsoft Excel Spreadsheet Document to a Adobe PDF Format file - PDF and save it to a UNC path.

Command Line 
-

 ````
 docto -XL -f '\\MyServer\path\Spreadsheet.xls' -o '\\\\myserver\path\Output\Spreadsheet.PDF' -t xlpdf
 ````

 or easier to read

  ````
 docto -XL  -f '\\MyServer\path\Spreadsheet.xls' 
            -o '\\\\myserver\path\Output\Spreadsheet.PDF'
            -t xlpdf
 ````

Command Line Explained 
-

 - `-XL`   This is a conversion using Microsoft Excel.  
 - `-f`   The File or directory to be converted 
 - `-o`   The Output File or Directory where you would like the converted file to be written to.
 - `-T`   The file format type that is being converted to




Some other interesting commands
-

You might find some of the following commands also interesting.

   

Other File Types Available for Conversion
-

The following values below can be used to convert a Microsoft Excel Spreadsheet to another file type.


````
xlAddIn=18
xlCSV=6
xlCSVMac=22
xlCSVMSDOS=24
xlCSVWindows=23
xlCurrentPlatformText=-4158
xlDBF2=7
xlDBF3=8
xlDBF4=11
xlDIF=9
xlExcel12=50
xlExcel2=16
xlExcel2FarEast=27
xlExcel3=29
xlExcel4=33
xlExcel4Workbook=35
xlExcel5=39
xlExcel7=39
xlExcel8=56
xlExcel9795=43
xlHtml=44
xlIntlAddIn=26
xlIntlMacro=25
xlOpenDocumentSpreadsheet=60
xlOpenXMLAddIn=55
xlOpenXMLTemplate=54
xlOpenXMLTemplateMacroEnabled=53
xlOpenXMLWorkbook=51
xlOpenXMLWorkbookMacroEnabled=52
xlSYLK=2
xlTemplate=17
xlTextMac=19
xlTextMSDOS=21
xlTextPrinter=36
xlTextWindows=20
xlUnicodeText=42
xlWebArchive=45
xlWJ2WD1=14
xlWJ3=40
xlWJ3FJ3=41
xlWK1=5
xlWK1ALL=31
xlWK1FMT=30
xlWK3=15
xlWK3FM3=32
xlWK4=38
xlWKS=4
xlWorkbookDefault=51
xlWorkbookNormal=-4143
xlWorks2FarEast=28
xlWQ1=34
xlXMLSpreadsheet=46
xlTypePDF=50000
xlpdf=50000
xlTypeXPS=50001
xlXPS=50001

```` 

