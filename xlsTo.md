# XLSTo

Simple utility for converting an Excel Spreadsheet  '.xls' file to any other supported file format 
such as `.txt` `.csv` `.pdf` etc.  

Can also be used to convert .csv, .txt to .xls .

Can be used to convert older excel documents to latest format.

Must have Microsoft Excel installed on host machine.

Download Release From Github Releases - https://github.com/tobya/DocTo/releases/
Further Information available at https://tobya.github.io/DocTo/

DocTo
=====
Word version available - [Details here](readme.md)

## Features

  1. Convert XLS file to any Excel SaveAs Type XLS,CSV,TXT
  1. Single File Conversion
  1. Multiple / Directory File Conversion
  1. Delete after conversion
  1. Fire Webhook one each conversion.
  
## Bugs and Features

Please log an [issue](https://github.com/tobya/DocTo/issues) for any bugs, features or suggestions.


## Examples

### Single

Convert Microsoft Excel Document to csv text

    xlsto -f C:\Directory\MyFile.xls -O "C:\Output Directory\MyTextFile.csv" -T xlCSV

Convert Microsoft Excel Document to PDF (requires version of Microsoft Excel that supports this).

     xlsto -f C:\Directory\MyFile.xls -O "C:\Output Directory\MyTextFile.pdf" -T xlTypePDF

### Multiple Files and Folders

Convert All Microsoft Excel Documents in Directory and its Sub Directories to PDF

    xlsto -f "C:\Dir with Spaces\FilesToConvert\" -O "C:\DirToOutput" -T xlTypePDF  -OX .pdf

#### Delete Origional File after Conversion ####

Delete Origional Files after conversion (-R) . 

    xlsto -f "C:\Dir with Spaces\FilesToConvert\" -O "C:\DirToOutput" -T xlTypePDF  -OX .pdf -R

Webhooks
========

Add a Webhook to fire on each conversion (-W)

    xlsto -f "C:\Dir with Spaces\FilesToConvert\" -O "C:\DirToOutput" -T xlTypePDF  -OX .pdf  -W http://toflidium.com/webhooks/docto/webhook_test.php
    
A Webhook is a url that can be called on each converstion to give you the ability to repond externally whenever a file is converted.

Use in the Wild
---------------

If you are using XLSTo or DocTo in the wild somewhere, please add details to this [wiki page](https://github.com/tobya/DocTo/wiki/Uses-of-DocTo-in-the-wild)


## Command Line Help

    Help
    Version:0.7.9
    Source: https://github.com/tobya/DocTo/
    Command Line Parameters
    Each Parameter should be followed by its value  -f "c:\Docs\MyDoc.doc" -O "C:\MyDir\MyFile"
    Parameters markers are case insensitive. Short and Long can be used mixed.
      -H  This message
          --help -?
      -F  Input File or Directory
          --inputfile
      -FX Input Extension to search for if directory.  Default ".doc" (will find ".docx" also)
          --inputextension
      -O  Output File or Directory to place converted Docs
          --outputfile
      -OX Output Extension if -F is Directory. Please include '.' eg. '.pdf' .
          If not provided, pulled from standard list.
          --outputextension
      -T  Format(Type) to convert file to, either integer or wdSaveFormat constant.
          Available from https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdsaveformat
          or https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.xlfileformat
          See current List Below.
          --format
      -TF Force Format.  -T value if integer is checked against current list compiled in and not passed if unavailable.
          -TF will pass through value without checking. Word will return an "EOleException  Value out of range" error if invalid.
          Use instead of -T.
          --forceformat
      -L  Log Level Integer: 1 ERRORS Only, 2 STANDARD, 5 CHATTY, 10 VERBOSE
          Default: 2=STANDARD
          --loglevel
      -C  Compatibility Mode Integer. Set to an INTEGER value from https://msdn.microsoft.com/en-us/library/office/ff192388.aspx.
          Set the compatibility mode when you want to convert documents to a later version of word.
          See List Below
          --compatability
      -E  Encoding Integer: Sets codepage Encoding.  See
          https://msdn.microsoft.com/en-us/library/office/ff860880.aspx for more details and values.
          --encoding          
      -M  Ignore all files in __MACOSX\ subdirectory if it exists.  Default True.
          --ignoremacos
      -G  Write Log to file in directory
          --writelogfile
      -GL Log File Name to Use. Default 'DocTo.Log';
          --logfilename
      -Q  Quiet Mode: Nothing will be output to console.  To see any errors you must set -G or -GL
          Equivalent to setting -L 0
          --quiet
      -R  Remove Files after successful conversion: Default false;
          --deletefiles
      -W  Webhook: Url to call on events (plain url no params). See -HW for more details.
          --webhook
      -HW Webhook Help.
      -X  Halt on COM Error: Default True;  If you have trouble with some files not converting, set this to false to ignore
          errors and continue with batch job.
          --halterror
      -V  Show Versions.  DocTo and Word/Excel


    ERROR CODES:
    200 : Invalid File Format specified
    201 : Insufficient Inputs.  Minimum of Input File, Output File & Type
    202 : Incorrect switches.  Switch requires value
    203 : Unknown switch in command
    204 : Input File does not exist
    220 : Word or COM Error
    221 : Word not Installed
    400 : Unknown Error




