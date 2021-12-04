{
    "title" : "How do I convert a directory of file without doing the subdirectories ? " 
}

How do I convert just the directory I want and not all its subdirectories?      
-

When converting a folder, Docto will by default convert all subdirectories also.  By using the `--no-subdirs` parameter you can prevent this.  

  

This is the command line to restrict use to just the directory you want. 



Command Line 
-

 ````
 docto -WD -f 'c:\path\' -o 'c:\output\' -t wdFormatPDF --no-subdirs
 ````
 or easier to read
 ````
 docto  -WD 
        -f 'c:\path\' 
        -o 'c:\output\' 
        -t wdFormatPDF 
        --no-subdirs
 ````

Command Line Explained 
-

 - `{[$Params.appwd.cmd]}` :  {[$Params.appwd.desc]}
 - `{[$Params.dashf.cmd]}` :  {[$Params.dashf.desc]} 
 - `{[$Params.dasho.cmd]}` :  {[$Params.dasho.desc]}
 - `{[$Params.dasht.cmd]}` :  {[$Params.dasht.desc]}
 - `{[$Params.ddnosubdirs.cmd]}` :  {[$Params.ddnosubdirs.desc]}




Some other interesting commands
-

You might find some of the following commands also interesting.

- [Convert all Word Document in a folder to a pdf file](ConvertDirDocToFilepdf.md);
{[foreach from=$Command.RelatedLinks key=LinkTitle item=L]}
 - [{[$LinkTitle]}]({[$L]})
{[/foreach]}    

