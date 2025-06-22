# Parameter Overview

## Usage
3 Parameters are required

- -F Input File Name
- -O Output File Name
- -T Type to be converted to.

Parameters that take a value have a space seperating them from the value.  Some parameters do
not require a value.  All parameters are case insensitive.

### Input File or Directory
> -F --inputfile

The file or folder you wish docto to open.  If it is a folder, docto will load all files in that 
directory and its subdirectories. If you do not wish to load files from subdirectories see the `--no-subdirs`
parameter.

Conversion will be performed on each file in turn.

### Output File or Folder
> -O --outputfile

The filename or foldername where you would like the output files to be placed.  If Input is a file but
output is a folder then the output file will have the same name as the input but with the new extension.

### Conversion Type
> -T --format

Specify what format you wish to convert to such as `wdFormatPDF` or `wdFormatText` etc.

View possible [Word Formats](https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdsaveformat)
    and [Excel Formats](https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.xlfileformat).  Can also use the integer value


### Help
> -H , --Help

Display the help text listing all parameters and versions of docto and office applications

### Version
> -V --version

Display the version string of both DocTo and Microsoft Office.



### Application Selection
> -WD -XL -PP -VS

This parameter tells DocTo which of the applications you wish to use to load and save your document
For historical reasons DocTo defaults to -WD if no value is given, however it is a good habit to get
into to always use one of these values any time you use Docto.
 - -WD Microsoft Word
 - -XL Microsoft Excel
 - -PP Microsoft Powerpoint
 - -VS Microsoft Visio


### Input Folder Extension 
>-FX --inputextension

By default DocTo will load all files in the directory with the standard Application extension

eg.
- Word (*.doc*) matches .doc & .docx files 
- Excel (*.xls*) matches .xls & .xlsx files 
- Powerpoint (*.ppt*) matches .ppt & .pptx files
- Visio (*.vsd*)

If you wish to convert a differnt set of files eg *.rtf or *.txt you can specify it here by ext
such as `.rtf`


### Output Extension
> -OX --outputextension

The output extension on a conversion is pulled from a standard list, eg. if converting to wdFormatPDF the file
will be output with extension `.pdf`.  If you would like to specify your own extension (such as `.pdfx`) you can
with this parameter.

### Force Format Use
> -TF --forceformat

If -T is an integer if it is a value that wasnt available when DocTo was compiled it will raise an error.
If you use -TF it will pass the integer value of -T to the Office Application without checking.

### Logging
> -L --loglevel

Set level of log output.  -l 10 is useful for debugging. Use -l 0 or -Q to surpress logging.

####Levels
- 10 VERBOSE
- 9 CHATTY
- 5 STANDARD
- 1 ERRORS (default)
- 0 SILENT

### Document Compatibility
> -C --compatibility

Compatibility Mode Integer. Set to an INTEGER value from [msdn list](https://msdn.microsoft.com/en-us/library/office/ff192388.aspx) .

Set the compatibility mode of the version of word the document is to be compatible with.  Particuarily
useful when wishing to convert older documents to current version. Can be used to convert old 
word documents to be compatible with [onedrive](https://github.com/tobya/DocTo/wiki/OneDrive-Conversion).

### Document Encoding
> -E --encoding

Sets codepage Encoding.  See [MSDN](https://msdn.microsoft.com/en-us/library/office/ff860880.aspx)
 for more details and values.

### List Long running Files
> -N --ListLongRunning

Some files when being converted can cause a dialog box to pop up.  This can only be fixed by 
manual intervention.  By setting this parameter you can at least record the documents that are 
causing difficulty (to a file called `docto.ignore.txt`) and if you set `-NX` these documents will be skipped on subsequent executions.

### Skip Files in docto.ignore.txt file
> -NX --IgnoreLongRunningList {no-value-required}

When set any files listed in `docto.ignore.txt` in the same directory as DocTo.exe will be skipped.
This allows troublesome documents in a directory structure to be ignored.

### Logging

### Write to Log File
> -G --writelogfile [no value required]

Write the log to a file as well as stdout. `docto.log` by default.  

### Log File
> -GL --logfilename  {filename}

Specify the filename that you wish the logfile to be written to.

### Quiet Mode
> -Q --quiet [no value required]

No output to stdout.  Everything including errors are surpressed.  Use in conjunction with `-G`
to ensure you get errors.

### Delete Input Files
> -R --deletefiles {true|false}

If you would like for the inputfile to be deleted after conversion you can set this to true.

### Fire a Webhook
> -W --webhook

If you wish you can call a web url after each conversion or error.
The Webhook URL will be called on the following events with the following parameters

  - File Conversion
    - action=convert
    - type=wdFormatType (or int if no matching format type)
    - ouputfilename=File being written to.
    - inputfilename=File being converted.

  - Error
    - action=error
    - type=wdFormatType (or int if no matching format type)
    - ouputfilename=File being written to.
    - inputfilename=File being converted.
    - error=Error Message

Return value is logged in DocTo Log

### Halt on Errors
> -X --halterror {true|false}

Docto will halt when a COM error is raised.  If you wish to ignore the error and continue set this value
to true.

### Bookmark Source
> --BookmarkSource {source}

PDF conversions can take their bookmarks from WordBookmarks, WordHeadings (default) or None

### Overwrite Files
> --DoNotOverwrite | --no-overwrite  [no value required]
      
Existing files are overridden by default, if you do not wish a file to be over written 
use this option.

### Recurse SubDirectories
> --no-subdirs [no value required]

By default sub directories are converted. Use to only convert specified directory. Do not recurse sub directories

### Export Markup
>   --ExportMarkup {exportvalue}

Specifies 
- `wdExportDocumentContent`	Exports the document without markup. (default)
- `wdExportDocumentWithMarkup` Exports the document with markup.

use  wdExportDocumentWithMarkup to export all word comments with pdf

### Open after Export
>   --PDF-OpenAfterExport

If you wish for the converted PDF to be opened after creation. No value req.

### Convert Specific Pages
> --PDF-FromPage

> --PDF-ToPage

Only convert certain pages in the document.

### Document Properties
> --no-IncludeDocProperties  --no-DocProp [no value required]

Do not include Document Properties in the exported pdf file.     

### Optomize PDF for on Screen display
> --PDF-OptimizeFor {for}

Set the pdf/xps to be optimized for print or screen.
  
- ForPrint (default) 
- ForOnScreen

### Keep IRM
> --XPS-no-IRM
  
Do not copy IRM permissions to exported XPS document.  By default these permissions are copied to the XPS file.

### Document Structure Tags
> --PDF-No-DocStructureTags

Do not include DocStructureTags to help screen readers.  By default these are included in the document.

### Bitmap Missing Fonts
> --PDF-no-BitmapMissingFonts
      
Do not bitmap missing fonts, fonts will be substituted. 

### Use ISO19005-1 
> --use-ISO190051 

Create PDF to the ISO 19005-1 standard, also know as PDF-A or PDF Archive.


### Special Case Parameters

### Do not ignore __MACOSX Directory
> -M --ignoreMACOS {true|false}

By default DocTo ignores any files in a hidden `__MACOSX` directory that MACOS creates.  This directory is oftern 
present on an external disk that is shared between systesm.  If you wish to check this dir 
set this value. You must specify value eg `-M false`.

### Experimental

### Skip Table of Contents
> --skipdocswithtoc

EXPERIMENTAL.  Will skip any docs that contain a TOC to prevent hanging. Currently matches some false positives.  Default False.

### Output to stdout
> --stdout
  
Send file to Stdout after conversion. ( Does not work correctly for binary files)  Works ok with text based
files such as .txt, .rtf, .html













