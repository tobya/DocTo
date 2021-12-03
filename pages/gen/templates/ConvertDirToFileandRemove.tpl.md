{
    "title" : "{[$Command.Title]} " 
}

{[$Command.Title]}       
-

When you carry out a conversion, sometimes you would like to remove the origional file.  This is particuarily useful when converting a directory.       

Sometimes you need to convert an entire directory of Word Documents to {[$Command.FileTypeExt]}s and remove the files that have been converted.  You can do this easily on the command line using [Docto](https://github.com/tobya/docto). 

The command line below shows how you can convert all the Microsoft Word Documents in a folder to a {[$Command.FileTypeDescription]} ({[$Command.FileTypeExt]}) file and then remove the file.

Command Line 
-

 ````
 docto -WD -f 'c:\path\todocuments' -o 'c:\path\toOutputfiles' -t {[$Command.FileFormat]} -R true
 ````
 or easier to read
 ````
 docto  -WD 
        -f 'c:\path\todocuments' 
        -o 'c:\path\toOutputfiles' 
        -t {[$Command.FileFormat]}
        -R true
 ````

Command Line Explained 
-

 - `{[$Params.appwd.cmd]}` :  {[$Params.appwd.desc]}
 - `{[$Params.dashf.cmd]}` :  {[$Params.dashf.desc]} 
 - `{[$Params.dasho.cmd]}` :  {[$Params.dasho.desc]}
 - `{[$Params.dasht.cmd]}` :  {[$Params.dasht.desc]}
 - `{[$Params.dashr.cmd]}` :  {[$Params.dashr.desc]}




Some other interesting commands
-

You might find some of the following commands also interesting.

- [Convert a Word Document to a {[$Command.FileTypeExt]} file](ConvertDocToFile{[$Command.FileTypeExt]}.md);
{[foreach from=$Command.RelatedLinks key=LinkTitle item=L]}
 - [{[$LinkTitle]}]({[$L]})
{[/foreach]}    

