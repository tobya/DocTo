{
    "title" : "How do I Convert a Folder full of RTF (Rich Text) documents to Adobe PDF Format? " 
}

How do I Convert a Folder full of RTF (Rich Text) documents to Adobe PDF Format?          
-

Even though it is easy to convert one [file at a time](ConvertDocToFilePDF.md), Sometimes you need to convert an entire directory of RTF Files ([or txt files](ConvertDirTXTToFile.md)) to PDFs.  You can do this easily on the command line using [Docto](https://github.com/tobya/docto). 

The command line below shows how you can convert all the RTF Files in a folder to a Adobe PDF Format file - PDF.  If you provide a directory and an extension, docto knows to convert all the documents matching that extension in that directory.

Command Line 
-

 ````
 docto -WD -f 'c:\path\todocuments' -fx .rtf -o 'c:\path\toOutputfiles' -t wdFormatPDF
 ````
 or easier to read
 ````
 docto  -WD 
        -f 'c:\path\todocuments' 
        -fx '.rtf'
        -o 'c:\path\toOutputfiles' 
        -t wdFormatPDF
 ````

Command Line Explained 
-

 - `-WD` :  This is a conversion using Microsoft Word. 
 - `-f` :  The File or directory to be converted 
 - `-fx` :  File with this extension should be matched. 
 - `-o` :  The Output File or Directory where you would like the converted file to be written to.
 - `-T` :  The file format type that is being converted to

When providing an extension to search for do not include the initial `*`, so `*.rtf` should be `.rtf`


Some other interesting commands
-

You might find some of the following commands also interesting.

- [Convert a Word Document to a PDF file](ConvertDocToFilePDF.md);
    

