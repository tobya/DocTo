{
    "title" : "How do I change where bookmarks come from in a converted PDF ? " 
}

How do I ensure any comments I have made on my Word Document are included in the converted PDF ?      
-

When converting to a PDF by Default Word will not show comments you have added to the document.  You can ask Word to include these using docto.  

  

This is the command line to include Word Comments in your PDF {[$Command.ContentExtra]}

{[$Command.ExportComments]}

all options available are

- wdExportDocumentContent
- wdExportDocumentWithMarkup

Command Line 
-

 ````
 docto -WD -f 'c:\path\Document.doc' -o 'c:\path\Document.{[$Command.FileTypeExt]}' -t wdFormatPDF --ExportMarkup {[$Command.ExportComments]}
 ````
 or easier to read
 ````
 docto  -WD 
        -f 'c:\path\Document.doc' 
        -o 'c:\path\Document.{[$Command.FileTypeExt]}' 
        -t wdFormatPDF 
        --ExportMarkup {[$Command.ExportComments]}
 ````

Command Line Explained 
-

 - `{[$Params.appwd.cmd]}` :  {[$Params.appwd.desc]}
 - `{[$Params.dashf.cmd]}` :  {[$Params.dashf.desc]} 
 - `{[$Params.dasho.cmd]}` :  {[$Params.dasho.desc]}
 - `{[$Params.dasht.cmd]}` :  {[$Params.dasht.desc]}
 - `{[$Params.ddexportmarkup.cmd]}` :  {[$Params.ddexportmarkup.desc]}




Some other interesting commands
-

You might find some of the following commands also interesting.

- [Convert all Word Document in a folder to a pdf file](ConvertDirDocToFilepdf.md);
{[foreach from=$Command.RelatedLinks key=LinkTitle item=L]}
 - [{[$LinkTitle]}]({[$L]})
{[/foreach]}    

