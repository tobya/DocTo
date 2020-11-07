{
    "title" : "{[$Command.Title]} " 
}

{[$Command.Title]}
==

It is very simple to convert a Microsoft Excel Spreadsheet to a {[$Command.FileTypeExt]} file  on the command line using [Docto](https://github.com/tobya/docto). You can also do this easily in Microsoft Excel, but sometimes it helps to be able to do it from the command line.  

The command line below shows how you can convert a Microsoft Excel Spreadsheet Document to a {[$Command.FileTypeDescription]} file - {[$Command.FileTypeExt]}.

Command Line 
-

 ````
 docto -XL -f 'c:\path\Spreadsheet.xls' -o 'c:\path\Spreadsheet.{[$Command.FileTypeExt]}' -t {[$Command.FileFormat]}
 ````

 or easier to read

  ````
 docto -XL  -f 'c:\path\Spreadsheet.xls' 
            -o 'c:\path\Spreadsheet.{[$Command.FileTypeExt]}' 
            -t {[$Command.FileFormat]}
 ````

Command Line Explained 
-

 - `{[$Params.appxl.cmd]}`   {[$Params.appxl.desc]}
 - `{[$Params.dashf.cmd]}`   {[$Params.dashf.desc]} 
 - `{[$Params.dasho.cmd]}`   {[$Params.dasho.desc]}
 - `{[$Params.dasht.cmd]}`   {[$Params.dasht.desc]}




Some other interesting commands
-

You might find some of the following commands also interesting.

{[foreach from=$Command.RelatedLinks key=LinkTitle item=L]}
 - [{[$LinkTitle]}]({[$L]})
{[/foreach]}   

Other File Types Available for Conversion
-

The following values below can be used to convert a Microsoft Excel Spreadsheet to another file type.


````
{[$ResourceFiles.xlsFormats.contents]}
```` 

