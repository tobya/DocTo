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

 - `{[$Params.appwd.cmd]}` :  {[$Params.appwd.desc]}
 - `{[$Params.dashf.cmd]}` :  {[$Params.dashf.desc]} 
 - `{[$Params.dasho.cmd]}` :  {[$Params.dasho.desc]}
 - `{[$Params.dashox.cmd]}` :  {[$Params.dashox.desc]}
 - `{[$Params.dasht.cmd]}` :  {[$Params.dasht.desc]}
 - `{[$Params.dashc.cmd]}` :  {[$Params.dashc.desc]}


The important part is the `-C 65535` which ensures that the output doc is compatible with the latest version of word.

See [this Stackexchange](http://webapps.stackexchange.com/questions/74859/what-format-does-word-onedrive-use) question 

Other Compatibility Values

````
{[$ResourceFiles.WdCompatibilityMode.contents]}
````



Some other interesting commands
-

You might find some of the following commands also interesting.

 
