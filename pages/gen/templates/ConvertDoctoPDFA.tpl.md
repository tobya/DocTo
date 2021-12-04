{
    "title" : "How do I convert word document to a ISO 19005-1 (PDF/A) standard PDF file ? " 
}

How do I convert word document to a ISO 19005-1 (PDF/A) standard PDF file ?      
-

When converting a document to a pdf, by default Microsoft Word does not output a PDF/A (PDF Archive) file.  If you wish to use this format, you can specify the `--use-ISO190051` parameter and the pdf will be saved as PDF/A format and be self contained.  Please note however the file size may be larger.  

  

This is the command line to use to output a pdf as iso 19005-1 format 



Command Line 
-

 ````
 docto -WD -f 'c:\path\' -o 'c:\output\' -t wdFormatPDF --use-ISO190051
 ````
 or easier to read
 ````
 docto  -WD 
        -f 'c:\path\' 
        -o 'c:\output\' 
        -t wdFormatPDF 
        --use-ISO190051
 ````

Command Line Explained 
-

 - `{[$Params.appwd.cmd]}` :  {[$Params.appwd.desc]}
 - `{[$Params.dashf.cmd]}` :  {[$Params.dashf.desc]} 
 - `{[$Params.dasho.cmd]}` :  {[$Params.dasho.desc]}
 - `{[$Params.dasht.cmd]}` :  {[$Params.dasht.desc]}
 - `{[$Params.dduseISO190051.cmd]}` :  {[$Params.dduseISO190051.desc]}




Some other interesting commands
-

You might find some of the following commands also interesting.

- [Convert all Word Document in a folder to a pdf file](ConvertDirDocToFilepdf.md);
{[foreach from=$Command.RelatedLinks key=LinkTitle item=L]}
 - [{[$LinkTitle]}]({[$L]})
{[/foreach]}    

