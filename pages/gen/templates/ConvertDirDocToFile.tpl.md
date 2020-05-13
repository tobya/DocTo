{
    "title" : "How do I Convert a Folder full of Microsoft Word Documents to a {[$Command.FileTypeDescription]}? " 
}

How do I Convert a Folder full of Microsoft Word Documents to a {[$Command.FileTypeDescription]}?         
-

Even though it is easy to convert one [file at a time](ConvertDocToFile{[$Command.FileTypeExt]}.md), Sometimes you need to convert an entire directory of Word Documents to {[$Command.FileTypeExt]}s.  You can do this easily on the command line using [Docto](https://github.com/tobya/docto). 

The command line below shows how you can convert all the Microsoft Word Documents in a folder to a {[$Command.FileTypeDescription]} file - {[$Command.FileTypeExt]}.

Command Line 
-

 ````
 docto -WD -f 'c:\path\todocuments' -o 'c:\path\toOutputfiles' -t {[$Command.FileFormat]}
 ````
 or easier to read
 ````
 docto  -WD 
        -f 'c:\path\todocuments' 
        -o 'c:\path\toOutputfiles' 
        -t {[$Command.FileFormat]}
 ````

Command Line Explained 
-

 - `{[$Params.appwd.cmd]}` :  {[$Params.appwd.desc]}
 - `{[$Params.dashf.cmd]}` :  {[$Params.dashf.desc]} 
 - `{[$Params.dasho.cmd]}` :  {[$Params.dasho.desc]}
 - `{[$Params.dasht.cmd]}` :  {[$Params.dasht.desc]}




Some other interesting commands
-

You might find some of the following commands also interesting.

- [Convert a Word Document to a {[$Command.FileTypeExt]} file](ConvertDocToFile{[$Command.FileTypeExt]}.md);
{[foreach from=$Command.RelatedLinks key=LinkTitle item=L]}
 - [{[$LinkTitle]}]({[$L]})
{[/foreach]}    

