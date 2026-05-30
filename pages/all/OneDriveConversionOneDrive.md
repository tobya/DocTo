Converting Word Documents for use with Onedrive
=======================
If you need to upgrade a bunch of files to work *without* conversion on OneDrive /Office365 / Word 20XX then you can use DocTo.

If you run docto on a directory to the latest version you can upload to OneDrive and then you can open your Word files in OneDrive and edit without conversion.  You need to use the following additional parameters to your command line 



Command Line 
-

 ````
 docto -WD -f 'c:\path\todocuments' -o 'c:\path\toOutputfiles' -ox '.docx'  -T wdFormatDocumentDefault -C 65535
 ````
 or easier to read
 ````
 docto  -WD 
        -f 'c:\path\todocuments' 
        -o 'c:\path\toOutputfiles' 
        -ox '.docx'        
        -t wdFormatDocumentDefault 
        -C 65535
 ````

these will make your outputted `.docx` files compatible with the latest version of word and allow you to upload to onedrive without requiring conversion when editing.

Command Line Explained 
-

 - `-WD` :  This is a conversion using Microsoft Word. 
 - `-f` :  The File or directory to be converted 
 - `-o` :  The Output File or Directory where you would like the converted file to be written to.
 - `-ox` :  The Output extension to be used for converted files
 - `-T` :  The file format type that is being converted to
 - `-c` :  Compatibility Level with Office Versions


The important part is the `-C 65535` which ensures that the output doc is compatible with the latest version of word.

See [this Stackexchange](http://webapps.stackexchange.com/questions/74859/what-format-does-word-onedrive-use) question 

Other Compatibility Values

````
wdCurrent=65535
wdWord2003=11
wdWord2007=12
wdWord2010=14
wdWord2013=15

````



Some other interesting commands
-

You might find some of the following commands also interesting.

 
