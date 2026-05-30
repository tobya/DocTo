{
    "title" : "How do I Convert a Microsoft Word Doc to a Open Document Text Format?" 
}

How do I Convert a Microsoft Word Doc to a Open Document Text Format?         
-

It is very simple to convert a Word Document to a ODD file  on the command line using [Docto](https://github.com/tobya/docto). You can also do this easily in Microsoft Word, but sometimes it helps to be able to do it from the command line.  

The command line below shows how you can convert a Microsoft Word Document to a Open Document Text Format file - ODD.

Command Line 
-

 ````
 docto -WD -f 'c:\path\Document.doc' -o 'c:\path\Document.ODD' -t wdFormatStrictOpenXMLDocument
 ````
 or easier to read
 ````
 docto -WD  -f 'c:\path\Document.doc' 
            -o 'c:\path\Document.ODD' 
            -t wdFormatStrictOpenXMLDocument
 ````

Command Line Explained 
-

 - `-WD` -  This is a conversion using Microsoft Word. 
 - `-f` -  The File or directory to be converted 
 - `-o` -  The Output File or Directory where you would like the converted file to be written to.
 - `-T` -  The file format type that is being converted to


You can see a [list below](#OtherTypes) of other conversion types available.

Some other interesting commands
-

You might find some of the following commands also interesting.

- [Convert all Word Document in a folder to a ODD file](ConvertDirDocToFileODD.md);

<a name="OtherTypes">Other File Types Available for Conversion</a>
-

The following values below can be used to convert a Word Document to another file type.


````

wdFormatDOSTextLineBreaks=5
wdFormatEncodedText=7
wdFormatFilteredHTML=10
wdFormatOpenDocumentText=23
wdFormatHTML=8
wdFormatRTF=6
wdFormatStrictOpenXMLDocument=24
wdFormatTemplate=1
wdFormatText=2
wdFormatTextLineBreaks=3
wdFormatUnicodeText=7
wdFormatWebArchive=9
wdFormatXML=11
wdFormatDocument97=0
wdFormatDocumentDefault=16
wdFormatPDF=17
wdFormatTemplate97=1
wdFormatXMLDocument=12
wdFormatXMLDocumentMacroEnabled=13
wdFormatXMLTemplate=14
wdFormatXMLTemplateMacroEnabled=15
wdFormatXPS=18


````


    

