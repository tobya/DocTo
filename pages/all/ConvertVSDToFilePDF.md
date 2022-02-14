{
    "title" : "How do I Convert a Microsoft Visio Document to a Adobe PDF Format?" 
}

How do I Convert a Microsoft Visio Document to a Adobe PDF Format?         
-

It is very simple to convert a Visio Document to a PDF file on the command line using [Docto](https://github.com/tobya/docto). You can also do this easily in Microsoft Visio, but sometimes it helps to be able to do it from the command line.  

The command line below shows how you can convert a Microsoft Visio Document to a Adobe PDF Format file - PDF.

Command Line 
-

 ````
 docto -VS -f 'c:\path\Document.vsd' -o 'c:\path\Document.PDF' -t visFixedFormatPDF 
 ````
 or easier to read
 ````
 docto -VS  -f 'c:\path\Document.vsd' 
            -o 'c:\path\Document.PDF' 
            -t visFixedFormatPDF
 ````

Command Line Explained 
-

 - `-VS` -  This is a conversion using Microsoft Visio.  
 - `-f` -  The File or directory to be converted 
 - `-o` -  The Output File or Directory where you would like the converted file to be written to.
 - `-T` -  The file format type that is being converted to


You can see a [list below](#OtherTypes) of other conversion types available.

Some other interesting commands
-

You might find some of the following commands also interesting.

- [Convert all Word Document in a folder to a PDF file](ConvertDirDocToFilePDF.md);

<a name="OtherTypes">Other File Types Available for Conversion</a>
-

The following values below can be used to convert a Word Document to another file type.


````

vsPDF=50000
vsXPS=50001
visFixedFormatPDF=50000
visFixedFormatXPS=50001


````


    

