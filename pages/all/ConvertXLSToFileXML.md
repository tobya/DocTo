{
    "title" : "How do I Convert a Microsoft Word Doc to a Open XML Workbook Format? " 
}

Open XML Workbook Format 
==

How do I Convert a Microsoft Excel Spreadsheet to a Open XML Workbook Format (XML)?         
-

It is very simple to convert a Microsoft Excel Spreadsheet to a XML file  on the command line using [Docto](https://github.com/tobya/docto). You can also do this easily in Microsoft Excel, but sometimes it helps to be able to do it from the command line.  

The command line below shows how you can convert a Microsoft Excel Spreadsheet Document to a Open XML Workbook Format file - XML.

Command Line 
-

 ````
 docto -XL -f 'c:\path\Spreadsheet.xls' -o 'c:\path\Spreadsheet.XML' -t xlOpenXMLWorkbook
 ````

 or easier to read

  ````
 docto -XL  -f 'c:\path\Spreadsheet.xls' 
            -o 'c:\path\Spreadsheet.XML' 
            -t xlOpenXMLWorkbook
 ````

Command Line Explained 
-

 - `-XL`   This is a conversion using Microsoft Excel.  
 - `-f`   The File or directory to be converted 
 - `-o`   The File or Directory where you would like the converted file to be written to.
 - `-T`   The file format type that is being converted to




Some other interesting commands
-

You might find some of the following commands also interesting.

    

