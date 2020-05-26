{
    "title" : "How do I Convert a Folder full of Microsoft Word Documents to Open Document Text Format? " 
}

How do I Convert a Folder full of Microsoft Word Documents to Open Document Text Format?          
-

Even though it is easy to convert one [file at a time](ConvertDocToFileODD.md), Sometimes you need to convert an entire directory of Word Documents to ODDs.  You can do this easily on the command line using [Docto](https://github.com/tobya/docto). 

The command line below shows how you can convert all the Microsoft Word Documents in a folder to a Open Document Text Format file - ODD.  If you provide a directory rather than a single file, docto knows to convert all the documents in that directory.

Command Line 
-

 ````
 docto -WD -f 'c:\path\todocuments' -o 'c:\path\toOutputfiles' -t wdFormatStrictOpenXMLDocument
 ````
 or easier to read
 ````
 docto  -WD 
        -f 'c:\path\todocuments' 
        -o 'c:\path\toOutputfiles' 
        -t wdFormatStrictOpenXMLDocument
 ````

Command Line Explained 
-

 - `-WD` :  This is a conversion using Microsoft Word.  This is not required but makes it easier to read
 - `-f` :  The File or directory to be converted 
 - `-o` :  The Output File or Directory where you would like the converted file to be written to.
 - `-T` :  The file format type that is being converted to




Some other interesting commands
-

You might find some of the following commands also interesting.

- [Convert a Word Document to a ODD file](ConvertDocToFileODD.md);
    

