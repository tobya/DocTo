{
    "title" : "How do I Convert a Folder full of Microsoft Word Documents to Text File? " 
}

How do I Convert a Folder full of Microsoft Word Documents to Text File?          
-

Even though it is easy to convert one [file at a time](ConvertDocToFileTXT.md), Sometimes you need to convert an entire directory of Word Documents to TXTs.  You can do this easily on the command line using [Docto](https://github.com/tobya/docto). 

The command line below shows how you can convert all the Microsoft Word Documents in a folder to a Text File file - TXT.  If you provide a directory rather than a single file, docto knows to convert all the documents in that directory.

Command Line 
-

 ````
 docto -WD -f 'c:\path\todocuments' -o 'c:\path\toOutputfiles' -t wdFormatText
 ````
 or easier to read
 ````
 docto  -WD 
        -f 'c:\path\todocuments' 
        -o 'c:\path\toOutputfiles' 
        -t wdFormatText
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

- [Convert a Word Document to a TXT file](ConvertDocToFileTXT.md);
    

