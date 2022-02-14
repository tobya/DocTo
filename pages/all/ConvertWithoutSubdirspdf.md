{
    "title" : "How do I convert a directory of file without doing the subdirectories ? " 
}

How do I convert just the directory I want and not all its subdirectories?      
-

When converting a folder, Docto will by default convert all subdirectories also.  By using the `--no-subdirs` parameter you can prevent this.  

  

This is the command line to restrict use to just the directory you want. 



Command Line 
-

 ````
 docto -WD -f 'c:\path\' -o 'c:\output\' -t wdFormatPDF --no-subdirs
 ````
 or easier to read
 ````
 docto  -WD 
        -f 'c:\path\' 
        -o 'c:\output\' 
        -t wdFormatPDF 
        --no-subdirs
 ````

Command Line Explained 
-

 - `-WD` :  This is a conversion using Microsoft Word. 
 - `-f` :  The File or directory to be converted 
 - `-o` :  The Output File or Directory where you would like the converted file to be written to.
 - `-T` :  The file format type that is being converted to
 - `--no-subdirs` :  Don't rescurse subdirs.  Only convert files in requested dir. 




Some other interesting commands
-

You might find some of the following commands also interesting.

- [Convert all Word Document in a folder to a pdf file](ConvertDirDocToFilepdf.md);
    

