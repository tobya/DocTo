# DocTo & XLSTo

DocTo
====

Simple utility for converting a Microsoft Word Document '.doc' file to any other supported format 
such as .txt .rtf .pdf.  

Can also be used to convert .txt, .rtf to .doc or .pdf format.

Can be used to convert older word documents to latest format.

Must have Microsoft Word installed on host machine.

Download Release From Github Releases - https://github.com/tobya/DocTo/releases/
Further Information available at http://tobya.github.com/DocTo

XLSTo
=====

Simple utility for converting an Excel Spreadsheet '.xls' file to any other supported file format such as .txt .csv .pdf etc.

Can also be used to convert .csv, .txt to .xls .

Can be used to convert older excel documents to latest format.

Must have Microsoft Excel installed on host machine.

Further information available [here](xlsTo.md)


## Features

  1. Convert Doc/RTF/Text file to any Word SaveAs Type Doc/Text/RTF/PDF
  1. Single File Conversion
  1. Multiple / Directory File Conversion
  1. Delete after conversion
  1. [Fire Webhook on each conversion.](https://github.com/tobya/DocTo#webhooks)

  
## Bugs and Features

Please log an [issue](https://github.com/tobya/DocTo/issues) for any bugs, features or suggestions.


## Examples

### Single

Convert Microsoft Word Document to text

    docto -f C:\Directory\MyFile.doc -O "C:\Output Directory\MyTextFile.txt" -T wdFormatText

Convert Microsoft Word Document to PDF (requires version of Microsoft Word that supports this).

     docto -f C:\Directory\MyFile.doc -O "C:\Output Directory\MyTextFile.pdf" -T wdFormatPDF

### Multiple Files and Folders

Convert All Microsoft Word Documents in Directory and its Sub Directories to PDF

    docto -f "C:\Dir with Spaces\FilesToConvert\" -O "C:\DirToOutput" -T wdFormatPDF  -OX .pdf

#### Delete Origional File after Conversion ####

Delete Origional Files after conversion (-R) . 

    docto -f "C:\Dir with Spaces\FilesToConvert\" -O "C:\DirToOutput" -T wdFormatPDF  -OX .pdf -R

Webhooks
========

Add a Webhook to fire on each conversion (-W)

    docto -f "C:\Dir with Spaces\FilesToConvert\" -O "C:\DirToOutput" -T wdFormatPDF  -OX .pdf  -W http://toflidium.com/webhooks/docto/webhook_test.php
    
A Webhook is a url that can be called on each converstion to give you the ability to repond externally whenever a file is converted.  Currently only a `http` and not a `https` address can be specified. 

Use in the Wild
---------------

If you are using DocTo in the wild somewhere, please add details to this [wiki page](https://github.com/tobya/DocTo/wiki/Uses-of-DocTo-in-the-wild)


OneDrive Conversion
=======================
If you need to upgrade a bunch of files to work without conversion on OneDrive /Office365 / Word 20XX then you can use DocTo.
See this StackExchange question 

http://webapps.stackexchange.com/questions/74859/what-format-does-word-onedrive-use

## Command Line Help

    Help
    Version:0.7.9
    Source: http://github.com/tobya/DocTo/
    Command Line Parameters
    Each Parameter should be followed by its value  -f "c:\Docs\MyDoc.doc" -O "C:\MyDir\MyFile"
    Parameters markers are case insensitive. Short and Long can be used mixed.
      -H  This message
          --help -?
      -F  Input File or Directory.  Please use full path filename.
          --inputfile
      -FX Input Extension to search for if directory.  Default ".doc" (will find ".docx" also)
          --inputextension
      -O  Output File or Directory to place converted Docs
          --outputfile
      -OX Output Extension if -F is Directory. Please include '.' eg. '.pdf' .
          If not provided, pulled from standard list.
          --outputextension
      -T  Format(Type) to convert file to, either integer or wdSaveFormat constant.
          Available from http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat.aspx
          or http://msdn.microsoft.com/en-us/library/office/bb241279(v=office.12).aspx
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
    Extra Parameters

      --skipdocswithtoc
        EXPERIMENTAL.  Will skip any docs that contain a TOC to prevent hanging.
        Currently matches some false positives.  Default False.

    ERROR CODES:
    200 : Invalid File Format specified
    201 : Insufficient Inputs.  Minimum of Input File, Output File & Type
    202 : Incorrect switches.  Switch requires value
    203 : Unknown switch in command
    204 : Input File does not exist
    220 : Word or COM Error
    221 : Word not Installed
    400 : Unknown Error

    COMPATIBILITY MODES:
    FROM https://msdn.microsoft.com/en-us/library/office/ff836084.aspx

    wdCurrent   : 65535 Compatibility mode equivalent to the latest version of Microsoft Word.
    wdWord2003  : 11    Word 2010 is put into a mode that is most compatible with Word 2003.
                        Features new to Word 2010 are disabled in this mode.
    wdWord2007  : 12    Word 2010 is put into a mode that is most compatible with Office Word 2007.
                        Features new to Word 2010 are disabled in this mode.
    wdWord2010  : 14    Word 2013 is put into a mode that is most compatible with . Features new to
                        Word 2013 are disabled in this mode.
    wdWord2013  : 15    Default. All Word 2013 features are enabled.

## Get Involved.

The project complies with Delphi (I use XE4 but it should compile with most versions including 7). I am happy to accept any PR anyone might like to submit.

## Updates     

0.7     Added support for saveas2 function (with compatibility mode) added by microsoft in Word 2010.  Additional Switches. Added support for long parameters.

0.5.5   Changes made to logging.  -Q and -L 0 now work correctly ensuring nothing is output to console.  Must specify -G or -GL to get access to logs and errors.
                                Also -L 10 now outputs extra as logging param is loaded first.

