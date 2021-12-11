{
    "title" : "How do I convert word document to a ISO 19005-1 (PDF/A) standard PDF file ? " 
}

How do I convert word document to a ISO 19005-1 (PDF/A) standard PDF file ?      
-

When converting a document to a pdf, by default Microsoft Word does not output a PDF/A (PDF Archive) file.  If you wish to use this format, you can specify the `--use-ISO190051` parameter and the pdf will be saved as PDF/A format and be self contained.  Please note however the file size may be larger.  

  

This is the command line to use to output a pdf as iso 19005-1 format 



Command Line 
-

 ````
 docto -WD -f 'c:\path\' -o 'c:\output\' -t wdFormatPDF --use-ISO190051
 ````
 or easier to read
 ````
 docto  -WD 
        -f 'c:\path\' 
        -o 'c:\output\' 
        -t wdFormatPDF 
        --use-ISO190051
 ````

Command Line Explained 
-

 - `-WD` :  This is a conversion using Microsoft Word. 
 - `-f` :  The File or directory to be converted 
 - `-o` :  The Output File or Directory where you would like the converted file to be written to.
 - `-T` :  The file format type that is being converted to
 - `--use-iso19005-1` :  Output PDFs from Word as ISO standard 19005-1 for self contained PDFS. Sometimes also refered to as <a href="https://en.wikipedia.org/wiki/PDF/A">PDF/A</a> 




Some other interesting commands
-

You might find some of the following commands also interesting.

- [Convert all Word Document in a folder to a pdf file](ConvertDirDocToFilepdf.md);
    

