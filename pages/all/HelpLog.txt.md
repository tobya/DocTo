

    
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
  -FX Input Extension to search for if directory. (.rtf .txt etc)
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
  --PDF-OpenAfterExport
      If you wish for a converted PDF to be opened after creation. No value req.
  --PDF-FromPage
      Save a range of pages to pdf. Integer/String. If integer --PDF-ToPage must also be set.
      Other values wdExportCurrentPage, wdExportSelection
  --PDF-ToPage
      Save a range of pages to pdf. Integer. --PDF-FromPage must also be set.
  --use-ISO190051 Create PDF to the ISO 19005-1 standard.




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

    
        