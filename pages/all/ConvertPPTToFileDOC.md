{
    "title" : "How do I Convert a Microsoft Powerpoint presentation to a Latest Microsoft Office Word 365 Version Format ? " 
}

How do I Convert a Microsoft Powerpoint presentation to a Latest Microsoft Office Word 365 Version Format  (DOC)?         
-

It is very simple to convert a Presentation to a DOC file  on the command line using [Docto](https://github.com/tobya/docto). You can also do this easily in Microsoft Powerpoint, but sometimes it helps to be able to do it from the command line.  

The command line below shows how you can convert a Powerpoint Presentation to a Latest Microsoft Office Word 365 Version Format  file - DOC.

Command Line 
-

 ````
 docto -PP -f 'c:\path\Presentation.ppt' -o 'c:\path\Presentation.DOC' -t wdFormatDocumentDefault
 ````
 or easier to read
 ````
 docto  -PP  
        -f 'c:\path\Presentation.ppt' 
        -o 'c:\path\Presentation.DOC' 
        -t wdFormatDocumentDefault
 ````

Command Line Explained 
-

 - `-PP` -  This is a conversion using Microsoft PowerPoint.  
 - `-f` -  The File or directory to be converted 
 - `-o` -  The Output File or Directory where you would like the converted file to be written to.
 - `-T` -  The file format type that is being converted to




Some other interesting commands
-

You might find some of the following commands also interesting.

- [Convert all Powerpoint presentations in a folder to a DOC file](ConvertDirPPTToFileDOC.md);
    

