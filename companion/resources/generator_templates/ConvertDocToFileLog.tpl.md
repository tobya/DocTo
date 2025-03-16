{
    "title" : "How do I get log out put when converting a file ? " 
}

How do I get log out put when converting a file ?         
-

When doing a conversion, it is useful to output log info.     

Here we set the log level to `{[$Command.LogLevel]}` which is  defined as '{[$Command.LogLevelDesc]}' which can be explained as {[$Command.LogLevelExtendedDesc]}

Command Line 
-

 ````
 docto -WD -f 'c:\path\Document.doc' -o 'c:\path\Document.{[$Command.FileTypeExt]}' -t wdFormatPDF -L {[$Command.LogLevel]}
 ````
 or easier to read
 ````
 docto -WD  -f 'c:\path\Document.doc' 
            -o 'c:\path\Document.{[$Command.FileTypeExt]}' 
            -t wdFormatPDF 
            -L {[$Command.LogLevel]}
 ````

Command Line Explained 
-

 - `{[$Params.appwd.cmd]}` :  {[$Params.appwd.desc]}
 - `{[$Params.dashf.cmd]}` :  {[$Params.dashf.desc]} 
 - `{[$Params.dasho.cmd]}` :  {[$Params.dasho.desc]}
 - `{[$Params.dasht.cmd]}` :  {[$Params.dasht.desc]}
 - `{[$Params.dashl.cmd]}` :  {[$Params.dashl.desc]}

O


Some other interesting commands
-

You might find some of the following commands also interesting.

- [Convert all Word Document in a folder to a pdf file](ConvertDirDocToFilepdf.md);
{[foreach from=$Command.RelatedLinks key=LinkTitle item=L]}
 - [{[$LinkTitle]}]({[$L]})
{[/foreach]}    

