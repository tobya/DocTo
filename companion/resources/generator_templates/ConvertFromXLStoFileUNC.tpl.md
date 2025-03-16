{
    "title" : "{[$Command.Title]} " 
}

{[$Command.Title]}
==

How do I Convert a Microsoft Excel Spreadsheet to a {[$Command.FileTypeDescription]} ({[$Command.FileTypeExt]}) and write it to a UNC Path?         
-

It is very simple to convert a Microsoft Excel Spreadsheet to a {[$Command.FileTypeExt]} file  on the command line using [Docto](https://github.com/tobya/docto). You can also do this easily in Microsoft Excel, but sometimes it helps to be able to do it from the command line.  

If you want to save your output to a UNC `\\Server` path, there is one extra trick you need to do.  You need to double up the slashes (\\) for the output file or Directory.  At the begining of the file path there is usually 2 slashes, you need to put 4. 

The command line below shows how you can convert a Microsoft Excel Spreadsheet Document to a {[$Command.FileTypeDescription]} file - {[$Command.FileTypeExt]} and save it to a UNC path.

Command Line 
-

 ````
 docto -XL -f '\\MyServer\path\Spreadsheet.xls' -o '\\\\myserver\path\Output\Spreadsheet.{[$Command.FileTypeExt]}' -t {[$Command.FileFormat]}
 ````

 or easier to read

  ````
 docto -XL  -f '\\MyServer\path\Spreadsheet.xls' 
            -o '\\\\myserver\path\Output\Spreadsheet.{[$Command.FileTypeExt]}'
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

