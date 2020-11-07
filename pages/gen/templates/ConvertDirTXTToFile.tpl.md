{
    "title" : "{[$Command.Title]}" 
}

{[$Command.Title]}         
-

Even though it is easy to convert [one file at a time](ConvertDocToFile{[$Command.FileTypeExt]}.md), Sometimes you need to convert an entire directory of text Files ([or rtf files](ConvertDirRTFToFile.md)) to {[$Command.FileTypeExt]}s.  You can do this easily on the command line using [Docto](https://github.com/tobya/docto). 

The command line below shows how you can convert all the text files in a folder to a {[$Command.FileTypeDescription]} file - {[$Command.FileTypeExt]}.  If you provide a directory and an extension, docto knows to convert all the documents matching that extension in that directory.

Command Line 
-

 ````
 docto -WD -f 'c:\path\todocuments' -fx .txt -o 'c:\path\toOutputfiles' -t {[$Command.FileFormat]}
 ````
 or easier to read
 ````
 docto  -WD 
        -f 'c:\path\todocuments' 
        -fx '.txt'
        -o 'c:\path\toOutputfiles' 
        -t {[$Command.FileFormat]}
 ````

Command Line Explained 
-

 - `{[$Params.appwd.cmd]}` :  {[$Params.appwd.desc]}
 - `{[$Params.dashf.cmd]}` :  {[$Params.dashf.desc]} 
 - `{[$Params.dashfx.cmd]}` :  {[$Params.dashfx.desc]} 
 - `{[$Params.dasho.cmd]}` :  {[$Params.dasho.desc]}
 - `{[$Params.dasht.cmd]}` :  {[$Params.dasht.desc]}

When providing an extension to search for do not include the initial `*`, so `*.txt` should be `.txt`


Some other interesting commands
-

You might find some of the following commands also interesting.

- [Convert a Word Document to a {[$Command.FileTypeExt]} file](ConvertDocToFile{[$Command.FileTypeExt]}.md);
{[foreach from=$Command.RelatedLinks key=LinkTitle item=L]}
 - [{[$LinkTitle]}]({[$L]})
{[/foreach]}    

