{
    "title" : "How do I Convert a Microsoft Excel Spreadsheet to a ODS?  " 
}

How do I Convert a Microsoft Excel Spreadsheet to a ODS? 
==

It is very simple to convert a Microsoft Excel Spreadsheet to a ODS file  on the command line using [Docto](https://github.com/tobya/docto). You can also do this easily in Microsoft Excel, but sometimes it helps to be able to do it from the command line.  

The command line below shows how you can convert a Microsoft Excel Spreadsheet Document to a Open Document Spreadsheet Format file - ODS.

Command Line 
-

 ````
 docto -XL -f 'c:\path\Spreadsheet.xls' -o 'c:\path\Spreadsheet.ODS' -t xlOpenDocumentSpreadsheet
 ````

 or easier to read

  ````
 docto -XL  -f 'c:\path\Spreadsheet.xls' 
            -o 'c:\path\Spreadsheet.ODS' 
            -t xlOpenDocumentSpreadsheet
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

