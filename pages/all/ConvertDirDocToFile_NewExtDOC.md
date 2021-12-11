{
    "title" : "How do I Convert a Folder full of Microsoft Word Documents to a Latest Microsoft Office Word 365 Version Format ? with differnt extension " 
}

How do I Convert a Folder full of Microsoft Word Documents to a Latest Microsoft Office Word 365 Version Format  with a specific Extension?         
-

Sometimes you wish to convert a file and give it a non standard extension. 

The command line below shows how you can convert all the Microsoft Word Documents in a folder to a Latest Microsoft Office Word 365 Version Format  file - DOC and give them the extension of .DOCX.

Command Line 
-

 ````
 docto -WD -f 'c:\path\todocuments' -o 'c:\path\toOutputfiles' -ox '.DOCX' -t wdFormatDocumentDefault 
 ````
 or easier to read
 ````
 docto  -WD 
        -f 'c:\path\todocuments' 
        -o 'c:\path\toOutputfiles' 
        -ox '.DOCX'
        -t wdFormatDocumentDefault
 ````

Command Line Explained 
-

 - `-WD` :  This is a conversion using Microsoft Word. 
 - `-f` :  The File or directory to be converted 
 - `-o` :  The Output File or Directory where you would like the converted file to be written to.
 - `-ox` :  The Output extension to be used for converted files
 - `-T` :  The file format type that is being converted to





Some other interesting commands
-

You might find some of the following commands also interesting.

- [Convert a Word Document to a DOC file](ConvertDocToFileDOC.md);
    

