{
    "title" : "How do I Convert a Folder full of Microsoft Word Documents to a {[$Command.FileTypeDescription]}? with differnt extension " 
}

How do I Convert a Folder full of Microsoft Word Documents to a {[$Command.FileTypeDescription]} with a specific Extension?         
-

Sometimes you wish to convert a file and give it a non standard extension. 

The command line below shows how you can convert all the Microsoft Word Documents in a folder to a {[$Command.FileTypeDescription]} file - {[$Command.FileTypeExt]} and give them the extension of {[$Command.FileNewExt]}.

Command Line 
-

 ````
 docto -WD -f 'c:\path\todocuments' -o 'c:\path\toOutputfiles' -ox '{[$Command.FileNewExt]}' -t {[$Command.FileFormat]} 
 ````
 or easier to read
 ````
 docto  -WD 
        -f 'c:\path\todocuments' 
        -o 'c:\path\toOutputfiles' 
        -ox '{[$Command.FileNewExt]}'
        -t {[$Command.FileFormat]}
 ````

Command Line Explained 
-

 - `{[$Params.appwd.cmd]}` :  {[$Params.appwd.desc]}
 - `{[$Params.dashf.cmd]}` :  {[$Params.dashf.desc]} 
 - `{[$Params.dasho.cmd]}` :  {[$Params.dasho.desc]}
 - `{[$Params.dashox.cmd]}` :  {[$Params.dashox.desc]}
 - `{[$Params.dasht.cmd]}` :  {[$Params.dasht.desc]}





Some other interesting commands
-

You might find some of the following commands also interesting.

- [Convert a Word Document to a {[$Command.FileTypeExt]} file](ConvertDocToFile{[$Command.FileTypeExt]}.md);
{[foreach from=$Command.RelatedLinks key=LinkTitle item=L]}
 - [{[$LinkTitle]}]({[$L]})
{[/foreach]}    

