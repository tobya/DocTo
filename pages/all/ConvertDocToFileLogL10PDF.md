{
    "title" : "How do I get log out put when converting a file ? " 
}

How do I get log out put when converting a file ?         
-

When doing a conversion, it is useful to output log info.     

Here we set the log level to `10` which is  defined as 'Verbose' which can be explained as A large amount of information will be output with the conversion.  Useful for debugging and trying to track down an issue.

Command Line 
-

 ````
 docto -WD -f 'c:\path\Document.doc' -o 'c:\path\Document.PDF' -t wdFormatPDF -L 10
 ````
 or easier to read
 ````
 docto -WD  -f 'c:\path\Document.doc' 
            -o 'c:\path\Document.PDF' 
            -t wdFormatPDF 
            -L 10
 ````

Command Line Explained 
-

 - `-WD` :  This is a conversion using Microsoft Word. 
 - `-f` :  The File or directory to be converted 
 - `-o` :  The Output File or Directory where you would like the converted file to be written to.
 - `-T` :  The file format type that is being converted to
 - `-L` :  Set output Log level.

O


Some other interesting commands
-

You might find some of the following commands also interesting.

- [Convert all Word Document in a folder to a pdf file](ConvertDirDocToFilepdf.md);
    

