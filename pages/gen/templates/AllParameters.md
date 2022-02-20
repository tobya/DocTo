# Parameter Overview

## Usage
3 Parameters are required

- -F Input File Name
- -O Output File Name
- -T Type to be converted to.

Parameters that take a value have a space seperating them from the value.  Some parameters do
not require a value.

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
> -








