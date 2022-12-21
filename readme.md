# DocTo

Document Converter
====

Simple utility for converting a Microsoft Word Document '.doc', Microsoft Excel '.xls' and Microsoft Powerpoint .ppt files to any other supported format 
such as .txt .csv .rtf .pdf.  

Can also be used to convert .txt, .rtf, .csv to .doc, .xls or .pdf format.

Can be used to convert older word documents to latest format.

Must have Microsoft Word, Excel or Powerpoint installed on host machine.

Download Release From Github Releases - https://github.com/tobya/DocTo/releases/
Further Information available at https://tobya.github.io/DocTo/
Further Examples available at https://docto.toflidium.com


## Features

  1. Convert Doc/RTF/Text file to any Word SaveAs Type Doc/Text/RTF/PDF
  1. Convert XLS/XLSX/CSV file to any Excel SaveAs Type CSV/Text/PDF
  1. Convert Text/CSV file to full fledged Word or Excel format.
  1. [Single File Conversion](https://github.com/tobya/DocTo/wiki/Converting-a-File-to-PDF)
  1. [Multiple / Directory File Conversion.](https://github.com/tobya/DocTo/wiki/Converting-a-directory-of-files)
  1. [Delete after conversion](https://github.com/tobya/DocTo/wiki/Delete-%5Cinput-File-after-conversion)
  1. [Fire https Webhook on each conversion.](https://github.com/tobya/DocTo#webhooks)
 
## Examples
More Examples available at 
- [View Examples](https://github.com/tobya/DocTo/blob/master/pages/all/index.md) 
- [https://docto.toflidium.com/](https://docto.toflidium.com/) 
- [Wiki](https://github.com/tobya/DocTo/wiki)
- [All Parameters Explained](/pages/gen/templates/AllParameters.md)

## Installation

Download .exe from Release https://github.com/tobya/docTo/releases

Package Managers
--

### Choco


Also Available for [installation via Chocolatey ](https://chocolatey.org/packages/docto)

> choco install docto

to upgrade to latest version before generally available (replace with current version)

> choco upgrade docto --version=1.8

### Node

Node Wrappers has been created by @KerimG & @brrd

https://www.npmjs.com/package/node-docto

https://github.com/brrd/msoconvert
  
## Bugs and Features

Please log an [issue](https://github.com/tobya/DocTo/issues) for any bugs, features or suggestions.


## Examples

### Single

Convert **Microsoft Word Document** to text

    docto -f C:\Directory\MyFile.doc -O "C:\Output Directory\MyTextFile.txt" -T wdFormatText

Convert **Microsoft Excel Document** to csv text

    docto -XL -f C:\Directory\MyFile.xls -O "C:\Output Directory\MyTextFile.csv" -T xlCSV    

Convert **Microsoft Word Document** to PDF (requires version of Microsoft Word that supports this).

     docto -f C:\Directory\MyFile.doc -O "C:\Output Directory\MyTextFile.pdf" -T wdFormatPDF

### Multiple Files and Folders

Convert All Microsoft Word Documents in Directory and its Sub Directories to PDF

    docto -f "C:\Dir with Spaces\FilesToConvert\" -O "C:\DirToOutput" -T wdFormatPDF  -OX .pdf

#### Delete Original File after Conversion ####

Delete Original Files after conversion (-R) . 

    docto -f "C:\Dir with Spaces\FilesToConvert\" -O "C:\DirToOutput" -T wdFormatPDF  -OX .pdf -R true

Webhooks
========

Add a Webhook to fire on each conversion (-W)

    docto -f "C:\Dir with Spaces\FilesToConvert\" -O "C:\DirToOutput" -T wdFormatPDF  -OX .pdf  -W https://toflidium.com/webhooks/docto/webhook_test.php
    
A Webhook is a url that can be called on each converstion to give you the ability to repond externally whenever a file is converted.  Currently `https` address is experimental so log an [issue](https://github.com/tobya/DocTo/issues/new) if you have any issues.

Use in the Wild
---------------

If you are using DocTo in the wild somewhere, please add details to this [wiki page](https://github.com/tobya/DocTo/wiki/Uses-of-DocTo-in-the-wild)


OneDrive Conversion
=======================
If you need to upgrade a bunch of files to work without conversion on OneDrive /Office365 / Word 20XX then you can use DocTo.
See this StackExchange question 

https://webapps.stackexchange.com/questions/74859/what-format-does-word-online-use

## Command Line Help
    Help
    Docto Version:%s
    Office Version : %s
    Open Source: https://github.com/tobya/DocTo/
    Description: DocTo converts Word Documents and Excel Spreadsheets to other formats.
    
    Command Line Parameters:
    Each Parameter should be followed by its value eg
            -f "c:\Docs\MyDoc.doc"
    Parameters markers are case insensitive.
    
      -H  This message
          --HELP -?
      -WD Use Word for Converstion (Default). Help '-h -wd'
          --word
      -XL Use Excel for Conversion. Help '-h -xl'
          --excel
      -PP Use Powerpoint for Conversion. help '-h -pp'
          --powerpoint
      -VS Use Visio for Conversion. 
          --visio
      -F  Input File or Directory
          --inputfile
      -FX Input file search for if -f is directory. Can use .rtf test*.txt etc
          Default ".doc*" (will find ".docx" also)
          --inputextension
      -O  Output File or Directory to place converted Docs
          --outputfile
      -OX Output Extension if -F is Directory. Please include '.' eg. '.pdf' .
          If not provided, pulled from standard list.
          --outputextension
      -T  Format(Type) to convert file to, either integer or wdSaveFormat constant.
          Available from
          https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdsaveformat
          or https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.xlfileformat
          or https://docs.microsoft.com/en-us/office/vba/api/powerpoint.presentation.saveas
          See current List Below.
          --format
      -TF Force Format. -T value if an integer, is checked against current list
          compiled in. It is not passed if unavailable.  -TF will pass through value
          without checking. Word will return an "EOleException  Value out of range"
          error if invalid. Use instead of -T.
          --forceformat
      -L  Log Level Integer: 1 ERRORS 2 STANDARD 5 CHATTY 9 DEBUG 10 VERBOSE. Default: 2=STANDARD
          --loglevel
      -C  Compatibility Mode Integer. Set to an INTEGER value from
          https://msdn.microsoft.com/en-us/library/office/ff192388.aspx.
          Set the compatibility mode when you want to convert documents to a later
          version of word. See help '-h -c' for further info.
          --compatibility
      -E  Encoding Integer: Sets codepage Encoding.  See
          https://msdn.microsoft.com/en-us/library/office/ff860880.aspx
          for more details and values.
          --encoding
      -M  Ignore all files in __MACOSX\ subdirectory if it exists.  Default True.
          --ignoremacos
      -N  Make list of files that take over n seconds to complete.
          Use number of seconds over that conversion takes and add to list.
          Outputs to filename 'docto.ignore.txt'
          --listlongrunning
      -NX Ignore any file listed in docto.ignore.txt, created by -N
          --ignorelongrunninglist
      -G  Write Log to file in directory
          --writelogfile
      -GL Log File Name to Use. Default 'DocTo.Log';
          --logfilename
      -Q  Quiet Mode: Nothing will be output to console.  To see any errors you must
          set -G or -GL. Equivalent to setting -L 0
          --quiet
      -R  Remove Files after successful conversion: Default false; To use specify
          value eg -R true
          --deletefiles
      -W  Webhook: Url to call on events. See help '-H -HW' for more details.
          --webhook
      -X  Halt on COM Error: Default True;  If you have trouble with some files
          not converting, set this to false to ignore errors and continue with
          batch job.
          --halterror
      -V  Show Versions.  DocTo and Word/Excel/Powerpoint
    
    Long Parameters:
    
      --BookmarkSource
          PDF conversions can take their bookmarks from
          WordBookmarks, WordHeadings (default) or None
      --DoNotOverwrite
      --no-overwrite
          Existing files are overridden by default, if you do not wish a file to be
          over written use this option.
      --no-subdirs Only convert specified directory. Do not recurse sub directories
      --ExportMarkup Value for wdExportItem - default wdExportDocumentContent.
          use    wdExportDocumentWithMarkup to export all word comments with pdf
      --no-IncludeDocProperties 
      --no-DocProp
          Do not include Document Properties in the exported pdf file.      
      --PDF-OpenAfterExport
          If you wish for a converted PDF to be opened after creation. No value req.
      --PDF-FromPage
          Save a range of pages to pdf. Integer/String. If integer --PDF-ToPage must also be set.
          Other values wdExportCurrentPage, wdExportSelection
      --PDF-ToPage
          Save a range of pages to pdf. Integer. --PDF-FromPage must also be set.
      --PDF-OptimizeFor
          Set the pdf/xps to be optimized for print or screen.
          Default  ForPrint | ForOnScreen
      --XPS-no-IRM
          Do not copy IRM permissions to exported XPS document.
      --PDF-No-DocStructureTags
          Do not include DocStructureTags to help screen readers.
      --PDF-no-BitmapMissingFonts
          Do not bitmap missing fonts, fonts will be substituted.   
      --use-ISO190051 
          Create PDF to the ISO 19005-1 standard.
    
    
    
    
    Experimental:
      --skipdocswithtoc
          EXPERIMENTAL.  Will skip any docs that contain a TOC to prevent hanging.
          Currently matches some false positives.  Default False.
      --stdout
          Send file to Stdout after conversion. ( Does not work correctly for binary files)
    
    ERROR CODES:
    200 : Invalid File Format specified
    201 : Insufficient Inputs.  Minimum of Input File, Output File & Type
    202 : Incorrect switches.  Switch requires value
    203 : Unknown switch in command
    204 : Input File does not exist
    205 : Invalid Parameter Value
    220 : Word or COM Error
    221 : Word not Installed
    400 : Unknown Error
        
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
> --DoNotOverwrite  --no-overwrite  [no value required]
      
Existing files are overridden by default, if you do not wish a file to be over written 
use this option.

### Recurse SubDirectories
> --no-subdirs 

By default sub directories are converted. Use to only convert specified directory. Do not recurse sub directories

### Export Markup
>   --ExportMarkup 

Specifies 
- wdExportDocumentContent	Exports the document without markup.
- wdExportDocumentWithMarkup Exports the document with markup.

use  wdExportDocumentWithMarkup to export all word comments with pdf

### Open after Export
>   --PDF-OpenAfterExport

If you wish for the converted PDF to be opened after creation. No value req.

### Convert Specific Pages
> --PDF-FromPage

> --PDF-ToPage

Only convert certain pages in the document.

### Use ISO19005-1 
> --use-ISO190051 

Create PDF to the ISO 19005-1 standard, also know as PDF-A or PDF Archive.


### Special Case Parameters

### Do not ignore __MACOSX Directory
> -M --ignoreMACOS {true|false}

By default DocTo ignores any files in a hidden `__MACOSX` directory that MACOS creates.  This directory is often 
present on an external disk that is shared between systems.  If you wish to check this dir set this value. You must specify value eg `-M false`.


Compiling
--
The project compiles with Delphi (I use 10.3 but it should compile with most versions including XE4 & 7). The project will not compile on Linux as it uses several Windows only components such as COM and Word and Excel do not have Linux versions anyway so there would be no point.

XLSTo
--

XLSTo is now incorporated into DocTo.  Previously XLSTo was a seperate EXE that was used to convert xls files to csv or pdf.  This can now be done with the main `DocTo.exe` by simply adding the -XL flag.


## Get Involved.

 I am happy to accept any PR anyone might like to submit.  If a large amount of work involved, please open an issue first to ensure the effort wont be wasted.

 The main branch name in the repo is `DocTo`

