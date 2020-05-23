{
    "title" : "How do I Convert a Microsoft Powerpoint presentation to a PNG File format? " 
}

How do I Convert a Microsoft Powerpoint presentation to a PNG File format (png)?         
-

It is very simple to convert a Presentation to a png file  on the command line using [Docto](https://github.com/tobya/docto). You can also do this easily in Microsoft Powerpoint, but sometimes it helps to be able to do it from the command line.  

The command line below shows how you can convert a Powerpoint Presentation to a PNG File format file - png.

Command Line 
-

 ````
 docto -PP -f 'c:\path\Presentation.ppt' -o 'c:\path\Presentation.png' -t ppSaveAsPNG
 ````
 or easier to read
 ````
 docto  -PP  
        -f 'c:\path\Presentation.ppt' 
        -o 'c:\path\Presentation.png' 
        -t ppSaveAsPNG
 ````

Command Line Explained 
-

 - `-PP` -  This is a conversion using Microsoft PowerPoint.  
 - `-f` -  The File or directory to be converted 
 - `-o` -  The Output File or Directory where you would like the converted file to be written to.
 - `-T` -  The file format type that is being converted to


Other File Types Available for Conversion
-

The following values below can be used to convert a Powerpoint presentation file to another file type.


````
ppSaveAsAddIn=8
ppSaveAsAnimatedGIF=40
ppSaveAsBMP=19
ppSaveAsDefault=11
ppSaveAsEMF=23
ppSaveAsExternalConverter=64000
ppSaveAsGIF=16
ppSaveAsJPG=17
ppSaveAsMetaFile=15
ppSaveAsMP4=39
ppSaveAsOpenDocumentPresentation=35
ppSaveAsOpenXMLAddin=30
ppSaveAsOpenXMLPicturePresentation=36
ppSaveAsOpenXMLPresentation=24
ppSaveAsOpenXMLPresentationMacroEnabled=25
ppSaveAsOpenXMLShow=28
ppSaveAsOpenXMLShowMacroEnabled=29
ppSaveAsOpenXMLTemplate=26
ppSaveAsOpenXMLTemplateMacroEnabled=27
ppSaveAsOpenXMLTheme=31
ppSaveAsPDF=32
ppSaveAsPNG=18
ppSaveAsPresentation=1
ppSaveAsRTF=6
ppSaveAsShow=7
ppSaveAsStrictOpenXMLPresentation=38
ppSaveAsTemplate=5
ppSaveAsTIF=21
ppSaveAsWMV=37
ppSaveAsXMLPresentation=34
ppSaveAsXPS=33
````



Some other interesting commands
-

You might find some of the following commands also interesting.

- [Convert all Powerpoint presentations in a folder to a png file](ConvertDirPPTToFilepng.md);
    

