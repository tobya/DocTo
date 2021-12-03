{
    "title" : "How do I Convert a Microsoft Powerpoint presentation to a {[$Command.FileTypeDescription]}? " 
}

How do I Convert a Microsoft Powerpoint presentation to a {[$Command.FileTypeDescription]} ({[$Command.FileTypeExt]})?         
-

It is very simple to convert a Presentation to a {[$Command.FileTypeExt]} file  on the command line using [Docto](https://github.com/tobya/docto). You can also do this easily in Microsoft Powerpoint, but sometimes it helps to be able to do it from the command line.  

The command line below shows how you can convert a Powerpoint Presentation to a {[$Command.FileTypeDescription]} file - {[$Command.FileTypeExt]}.

Command Line 
-

 ````
 docto -PP -f 'c:\path\Presentation.ppt' -o 'c:\path\Presentation.{[$Command.FileTypeExt]}' -t {[$Command.FileFormat]}
 ````
 or easier to read
 ````
 docto  -PP  
        -f 'c:\path\Presentation.ppt' 
        -o 'c:\path\Presentation.{[$Command.FileTypeExt]}' 
        -t {[$Command.FileFormat]}
 ````

Command Line Explained 
-

 - `{[$Params.apppp.cmd]}` -  {[$Params.apppp.desc]}
 - `{[$Params.dashf.cmd]}` -  {[$Params.dashf.desc]} 
 - `{[$Params.dasho.cmd]}` -  {[$Params.dasho.desc]}
 - `{[$Params.dasht.cmd]}` -  {[$Params.dasht.desc]}


Other File Types Available for Conversion
-

The following values below can be used to convert a Powerpoint presentation file to another file type.


````
{[$ResourceFiles.ppFormats.contents]}
````



Some other interesting commands
-

You might find some of the following commands also interesting.

- [Convert all Powerpoint presentations in a folder to a {[$Command.FileTypeExt]} file](ConvertDirPPTToFile{[$Command.FileTypeExt]}.md);
{[foreach from=$Command.RelatedLinks key=LinkTitle item=L]}
 - [{[$LinkTitle]}]({[$L]})
{[/foreach]}    

