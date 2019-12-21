# DocTo & XLSTo

DocTo
====

Simple utility for converting a Microsoft Word Document '.doc' and Microsoft Excel '.xls' files to any other supported format 
such as .txt .csv .rtf .pdf.  

Can also be used to convert .txt, .rtf, .csv to .doc, .xls or .pdf format.

Can be used to convert older word documents to latest format.

Must have Microsoft Word or Excel installed on host machine.

Download Release From Github Releases - https://github.com/tobya/DocTo/releases/
Further Information available at https://tobya.github.io/DocTo/

XLSTo
=====

XLSTo is now incorporated into DocTo.  Previously XLSTo was a seperate EXE that was used to convert xls files to csv or pdf.  This can now be done with the main `DocTo.exe` by simply adding the -XL flag.


## Features

  1. Convert Doc/RTF/Text file to any Word SaveAs Type Doc/Text/RTF/PDF
  1. Convert XLS/XLSX/CSV file to any Excel SaveAs Type CSV/Text/PDF
  1. [Single File Conversion](https://github.com/tobya/DocTo/wiki/Converting-a-File-to-PDF)
  1. [Multiple / Directory File Conversion.](https://github.com/tobya/DocTo/wiki/Converting-a-directory-of-files)
  1. [Delete after conversion](https://github.com/tobya/DocTo/wiki/Delete-%5Cinput-File-after-conversion)
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

#### Delete Original File after Conversion ####

Delete Original Files after conversion (-R) . 

    docto -f "C:\Dir with Spaces\FilesToConvert\" -O "C:\DirToOutput" -T wdFormatPDF  -OX .pdf -R true

Webhooks
========

Add a Webhook to fire on each conversion (-W)

    docto -f "C:\Dir with Spaces\FilesToConvert\" -O "C:\DirToOutput" -T wdFormatPDF  -OX .pdf  -W http://toflidium.com/webhooks/docto/webhook_test.php
    
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

 https://github.com/tobya/DocTo/blob/05931f146a35e5c2d3054d77dd9d38aef4cc013f/src/res/HelpLog.txt#L1-L114
 
## Get Involved.

The project compiles with Delphi (I use XE4 but it should compile with most versions including 7). I am happy to accept any PR anyone might like to submit.

## Updates     

0.7     Added support for saveas2 function (with compatibility mode) added by microsoft in Word 2010.  Additional Switches. Added support for long parameters.

0.5.5   Changes made to logging.  -Q and -L 0 now work correctly ensuring nothing is output to console.  Must specify -G or -GL to get access to logs and errors.
                                Also -L 10 now outputs extra as logging param is loaded first.

