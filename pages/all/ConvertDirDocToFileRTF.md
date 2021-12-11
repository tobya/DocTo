{
    "title" : "How do I Convert a Folder full of Microsoft Word Documents to Windows Rich Text Format? " 
}

How do I Convert a Folder full of Microsoft Word Documents to Windows Rich Text Format?          
-

Even though it is easy to convert one [file at a time](ConvertDocToFileRTF.md), Sometimes you need to convert an entire directory of Word Documents to RTFs.  You can do this easily on the command line using [Docto](https://github.com/tobya/docto). 

The command line below shows how you can convert all the Microsoft Word Documents in a folder to a Windows Rich Text Format file - RTF.  If you provide a directory rather than a single file, docto knows to convert all the documents in that directory.

Command Line 
-

 ````
 docto -WD -f 'c:\path\todocuments' -o 'c:\path\toOutputfiles' -t wdFormatRTF
 ````
 or easier to read
 ````
 docto  -WD 
        -f 'c:\path\todocuments' 
        -o 'c:\path\toOutputfiles' 
        -t wdFormatRTF
 ````

Command Line Explained 
-

 - `-WD` :  This is a conversion using Microsoft Word. 
 - `-f` :  The File or directory to be converted 
 - `-o` :  The Output File or Directory where you would like the converted file to be written to.
 - `-T` :  The file format type that is being converted to




Some other interesting commands
-

You might find some of the following commands also interesting.

- [Convert a Word Document to a RTF file](ConvertDocToFileRTF.md);
    

