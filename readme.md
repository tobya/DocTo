# DocTo

Simple utility for converting a Microsoft Word Document '.doc' file to any other supported format 
such as .txt .rtf .pdf.  

Further Information available at http://tobya.github.com/DocTo

Must have Microsoft Word installed on host machine.

Licensced in source and binary form under MIT Open Source License, see License.txt for details

Download Release From Github Releases - https://github.com/tobya/DocTo/releases/

## Features

  1. Convert Doc/RTF/Text file to any Word SaveAs Type Doc/Text/RTF/PDF
  1. Single File Conversion
  1. Multiple / Directory File Conversion
  1. Delete after conversion
  1. Fire Webhook one each conversion.
  

## Updates

0.5.5 Changes made to logging.  -Q and -L 0 now work correctly ensuring nothing is output to console.  Must specify -G or -GL to get access to logs and errors.
                                Also -L 10 now outputs extra as logging param is loaded first.


## Examples

### Single

Convert Microsoft Word Document to text

    docto -f C:\Directory\MyFile.doc -O "C:\Output Directory\MyTextFile.txt" -T wdFormatText

Convert Microsoft Word Document to PDF (requires version of Microsoft Word that supports this).

     docto -f C:\Directory\MyFile.doc -O "C:\Output Directory\MyTextFile.pdf" -T wdFormatPDF

### Multiple Files and Folders

Convert All Microsoft Word Documents in Directory and its Sub Directories to PDF

    docto -f "C:\Dir with Spaces\FilesToConvert\" -O "C:\DirToOutput" -T wdFormatPDF  -OX .pdf

Delete Origional Files after conversion.

    docto -f "C:\Dir with Spaces\FilesToConvert\" -O "C:\DirToOutput" -T wdFormatPDF  -OX .pdf -R

Webhooks
========

Add a Webhook to fire on each conversion

    docto -f "C:\Dir with Spaces\FilesToConvert\" -O "C:\DirToOutput" -T wdFormatPDF  -OX .pdf  -R -W http://toflidium.com/webhooks/docto/webhook_test.php


## Command Line Help

     docto -h


	Help
	Version:0.4ALPHA
	Source: http://github.com/tobya/DocTo/
	Command Line Parameters
	Each Parameter should be followed by its value  -f "c:\Docs\MyDoc.doc" -O "C:\MyDir\MyFile"
	  -H  This message
	  -F  Input File or Directory
	  -O  Output File or Directory to place converted Docs
	  -OX Output Extension if -F is Directory.
	  -T  Format(Type) to convert file to, either integer or wdSaveFormat constant.
	      Available from http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat.aspx
	      See current List Below.
	  -TF Force Format.  -T value if integer is checked against current list compiled in and not passed if unavailable.
	      To future proof, -TF will pass through value without checking.
	      Word will return an "EOleException  Value out of range" error if invalid.
	      Use instead of -T.
	  -L  Log Level 0 Silent, 1 Standard, 10 VERBOSE
	      Default: 1 Standard
	  -G  Write Log to file in directory
	  -GL Log File Name to Use default 'DocTo.Log';
	  -Q  Quiet Mode: Equivalent to setting -L 0
	  -R  Remove Files after successful conversion: Default false;
	  -W  Webhook: Url to call on events (plain url no params). See -HW for more details.
	  -HW Webhook Help.
	  -X  Halt on Word Error: Default True;  If you have trouble with somefiles not converting, set this to false.
	
	
	ERROR CODES:
	200 : Invalid File Format specified
	201 : Insufficient Inputs.  Minimum of Input File, Output File & Type
	202 : Incorrect switches.  Switch requires value
	203 : Unknown switch in command
	220 : Word or COM Error
	221 : Word not Installed


        FILE FORMATS
        wdFormatDOSTextLineBreaks=5
        wdFormatEncodedText=7
        wdFormatFilteredHTML=10
        wdFormatOpenDocumentText=23
        wdFormatHTML=8
        wdFormatRTF=6
        wdFormatStrictOpenXMLDocument=24
        wdFormatTemplate=1
        wdFormatText=2
        wdFormatTextLineBreaks=3
        wdFormatUnicodeText=7
        wdFormatWebArchive=9
        wdFormatXML=11
        wdFormatDocument97=0
        wdFormatDocumentDefault=16
        wdFormatPDF=17
        wdFormatTemplate97=1
        wdFormatXMLDocument=12
        wdFormatXMLDocumentMacroEnabled=13
        wdFormatXMLTemplate=14
        wdFormatXMLTemplateMacroEnabled=15
        wdFormatXPS=18


