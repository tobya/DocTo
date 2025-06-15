{
    "title" : "How do I delete a file after Converting it to  XPS?  " 
}

How do I delete a file after Converting it to  XPS?        
-

When you carry out a conversion, sometimes you would like to remove the origional file.  This is particuarily useful when converting a directory.       

Sometimes you need to convert an entire directory of Word Documents to XPSs and remove the files that have been converted.  You can do this easily on the command line using [Docto](https://github.com/tobya/docto). 

The command line below shows how you can convert all the Microsoft Word Documents in a folder to a Microsoft XPS Format (XPS) file and then remove the file.

Command Line 
-

 ````
 docto -WD -f 'c:\path\todocuments' -o 'c:\path\toOutputfiles' -t wdFormatXPS -R true
 ````
 or easier to read
 ````
 docto  -WD 
        -f 'c:\path\todocuments' 
        -o 'c:\path\toOutputfiles' 
        -t wdFormatXPS
        -R true
 ````

Command Line Explained 
-

 - `-WD` :  This is a conversion using Microsoft Word. 
 - `-f` :  The File or directory to be converted 
 - `-o` :  The Output File or Directory where you would like the converted file to be written to.
 - `-T` :  The file format type that is being converted to
 - `-R` :  Delete the converted file after conversion.




Some other interesting commands
-

You might find some of the following commands also interesting.

- [Convert a Word Document to a XPS file](ConvertDocToFileXPS.md);
   

