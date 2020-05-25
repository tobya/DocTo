{
    "title" : "How do I change where bookmarks come from in a converted PDF ? " 
}

How do I change where bookmarks come from in a converted PDF ?      
-

When converting to a PDF by Default Word Headings are used for the bookmarks in the PDF. You can if you wish use other options

  - WordHeadings (default) 
  - WordBookmarks
  - None      

This is the command line to change the bookmark source to WordBookmarks

Command Line 
-

 ````
 docto -WD -f 'c:\path\Document.doc' -o 'c:\path\Document..pdf' -t wdFormatPDF --bookmarksource WordBookmarks
 ````
 or easier to read
 ````
 docto  -WD 
        -f 'c:\path\Document.doc' 
        -o 'c:\path\Document..pdf' 
        -t wdFormatPDF 
        --bookmarksource WordBookmarks
 ````

Command Line Explained 
-

 - `-WD` :  This is a conversion using Microsoft Word.  This is not required but makes it easier to read
 - `-f` :  The File or directory to be converted 
 - `-o` :  The Output File or Directory where you would like the converted file to be written to.
 - `-T` :  The file format type that is being converted to
 - `-bookmarksource` :  Where to get Bookmarks from for PDF




Some other interesting commands
-

You might find some of the following commands also interesting.

- [Convert all Word Document in a folder to a pdf file](ConvertDirDocToFilepdf.md);
    

