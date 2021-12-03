{
    "title" : "How do I change where bookmarks come from in a converted PDF ? " 
}

How do I ensure any comments I have made on my Word Document are included in the converted PDF ?      
-

When converting to a PDF by Default Word will not show comments you have added to the document.  You can ask Word to include these using docto.  

  

This is the command line to include Word Comments in your PDF With this option no comments are present in the output PDF file

wdExportDocumentContent

all options available are

- wdExportDocumentContent
- wdExportDocumentWithMarkup

Command Line 
-

 ````
 docto -WD -f 'c:\path\Document.doc' -o 'c:\path\Document.pdf' -t wdFormatPDF --ExportMarkup wdExportDocumentContent
 ````
 or easier to read
 ````
 docto  -WD 
        -f 'c:\path\Document.doc' 
        -o 'c:\path\Document.pdf' 
        -t wdFormatPDF 
        --ExportMarkup wdExportDocumentContent
 ````

Command Line Explained 
-

 - `-WD` :  This is a conversion using Microsoft Word.  This is not required but makes it easier to read
 - `-f` :  The File or directory to be converted 
 - `-o` :  The Output File or Directory where you would like the converted file to be written to.
 - `-T` :  The file format type that is being converted to
 - `--ExportMarkup` :  Export Comments and other markup from Word Document to PDF




Some other interesting commands
-

You might find some of the following commands also interesting.

- [Convert all Word Document in a folder to a pdf file](ConvertDirDocToFilepdf.md);
    

