{
    "title" : "How do I change where bookmarks come from in a converted PDF ? " 
}

How do I change where bookmarks come from in a converted PDF ?      
-

When converting to a PDF by Default Word Headings are used for the bookmarks in the PDF. You can if you wish use other options

  - WordHeadings (default) 
  - WordBookmarks
  - None      

This is the command line to change the bookmark source to {[$Command.BookmarksSource]}

{[$Command.ContentExtra]}

Command Line 
-

 ````
 docto -WD -f 'c:\path\Document.doc' -o 'c:\path\Document.{[$Command.FileTypeExt]}' -t wdFormatPDF --bookmarksource {[$Command.BookmarksSource]}
 ````
 or easier to read
 ````
 docto  -WD 
        -f 'c:\path\Document.doc' 
        -o 'c:\path\Document.{[$Command.FileTypeExt]}' 
        -t wdFormatPDF 
        --bookmarksource {[$Command.BookmarksSource]}
 ````

Command Line Explained 
-

 - `{[$Params.appwd.cmd]}` :  {[$Params.appwd.desc]}
 - `{[$Params.dashf.cmd]}` :  {[$Params.dashf.desc]} 
 - `{[$Params.dasho.cmd]}` :  {[$Params.dasho.desc]}
 - `{[$Params.dasht.cmd]}` :  {[$Params.dasht.desc]}
 - `{[$Params.ddbookmarksource.cmd]}` :  {[$Params.ddbookmarksource.desc]}




Some other interesting commands
-

You might find some of the following commands also interesting.

- [Convert all Word Document in a folder to a pdf file](ConvertDirDocToFilepdf.md);
{[foreach from=$Command.RelatedLinks key=LinkTitle item=L]}
 - [{[$LinkTitle]}]({[$L]})
{[/foreach]}    

