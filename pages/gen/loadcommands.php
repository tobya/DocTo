<?php

$Commands = [



    "ConvertDocToFile" => [ 
        "Description" => "Convert Word Document to another file type",
        "Template" => "ConvertFromDocToFile.tpl.md",
        "Title" => 'How do I Convert a Microsoft Word Doc to a {[$Command.FileTypeDescription]}?'       , 
        "Blocks" => [
            [
                ""
            ]
        ],
        "Items" => [
            [


                "FileTypeExt" => 'RTF',
                "FileTypeDescription" => 'Windows Rich Text Format',
                "FileFormat" => 'wdFormatRTF',
                "RelatedLinks" => []
            ],

            [
                "FileTypeExt" => 'PDF',
                "FileTypeDescription" => 'Adobe PDF Format',
                "FileFormat" => 'wdFormatPDF',
                "RelatedLinks" => []
            ],
            [
                "FileTypeExt" => 'TXT',
                "FileTypeDescription" => 'Text File',
                "FileFormat" => 'wdFormatText',
                "RelatedLinks" => []
            ],
            [
                "FileTypeExt" => 'ODD',
                "FileTypeDescription" => 'Open Document Text Format',
                "FileFormat" => 'wdFormatStrictOpenXMLDocument',
                "RelatedLinks" => []
            ],
            [
                "FileTypeExt" => 'XML',
                "FileTypeDescription" => 'XML Document Format',
                "FileFormat" => 'wdFormatXMLDocument',
                "RelatedLinks" => []
            ],
            [
                "FileTypeExt" => 'HTML',
                "FileTypeDescription" => 'HTML File',
                "FileFormat" => 'wdFormatHTML',
                "RelatedLinks" => []
            ],
           
              [
                "FileTypeExt" => 'XPS',
                "FileTypeDescription" => 'Microsoft XPS Format',
                "FileFormat" => 'wdFormatXPS',
                "RelatedLinks" => []
            ],
              [
                "FileTypeExt" => 'DOC',
                "FileTypeDescription" => 'Latest Microsoft Office Word 365 Version Format ',
                "FileFormat" => 'wdFormatDocumentDefault',
                "RelatedLinks" => []
            ],
        ],
    ],



   "ConvertDirDocToFile" => [ 
        "Description" => "Convert all Word Documents in a Directory",
        "Template" => "ConvertDirDocToFile.tpl.md",
        "Title" => 'How do I Convert a Folder full of Microsoft Word Documents to {[$Command.FileTypeDescription]}? ',
        "Items" => [
            [


                "FileTypeExt" => 'RTF',
                "FileTypeDescription" => 'Windows Rich Text Format',
                "FileFormat" => 'wdFormatRTF',
                "RelatedLinks" => []
            ],

            [
                "FileTypeExt" => 'PDF',
                "FileTypeDescription" => 'Adobe PDF Format',
                "FileFormat" => 'wdFormatPDF',
                "RelatedLinks" => [
                    'Convert all files in directory without converting subdirectories.' => 'ConvertWithoutSubdirspdf.md',
                ]
            ],
           
            [
                "FileTypeExt" => 'HTML',
                "FileTypeDescription" => 'HTML File',
                "FileFormat" => 'wdFormatHTML',
                "RelatedLinks" => []
            ],
           
              [
                "FileTypeExt" => 'DOC',
                "FileTypeDescription" => 'Latest Microsoft Office Word 365 Version Format ',
                "FileFormat" => 'wdFormatDocumentDefault',
                "RelatedLinks" => []
            ],
        ],
    ],

   "ConvertDirRTFToFile" => [ 
        "Description" => "Convert all RTF Documents in a Directory",
        "Template" => "ConvertDirRTFToFile.tpl.md",
        "Title" => 'How do I Convert a Folder full of RTF (Rich Text) documents to {[$Command.FileTypeDescription]}? ',
        "Items" => [
            [


                "FileTypeExt" => 'Doc',
                "FileTypeDescription" => 'Microsoft Word Document',
                "FileFormat" => 'wdFormatDocumentDefault',
                "RelatedLinks" => []
            ],

            [
                "FileTypeExt" => 'PDF',
                "FileTypeDescription" => 'Adobe PDF Format',
                "FileFormat" => 'wdFormatPDF',
                "RelatedLinks" => []
            ],
            [
                "FileTypeExt" => 'TXT',
                "FileTypeDescription" => 'Text File',
                "FileFormat" => 'wdFormatText',
                "RelatedLinks" => []
            ],
            
           
            [
                "FileTypeExt" => 'HTML',
                "FileTypeDescription" => 'HTML File',
                "FileFormat" => 'wdFormatHTML',
                "RelatedLinks" => []
            ],
           
        ],
    ],
   "ConvertDirTXTToFile" => [ 
        "Description" => "Convert all Text Files in a Directory",
        "Template" => "ConvertDirTXTToFile.tpl.md",
        "Title" => 'How do I Convert a Folder full of text files to {[$Command.FileTypeDescription]}? ',
        "Items" => [
            [


                "FileTypeExt" => 'Doc',
                "FileTypeDescription" => 'Microsoft Word Document',
                "FileFormat" => 'wdFormatDocumentDefault',
                "RelatedLinks" => []
            ],

            [
                "FileTypeExt" => 'PDF',
                "FileTypeDescription" => 'Adobe PDF Format',
                "FileFormat" => 'wdFormatPDF',
                "RelatedLinks" => []
            ],
            
           
        ],
    ],
    "ConvertDirToFileRemove" => [ 
        "Description" => "Convert all Word Documents in a Directory and remove after conversion",
        "Template" => "ConvertDirToFileandRemove.tpl.md",
        'Title' => 'How do I delete a file after Converting it to  {[$Command.FileTypeExt]}? ',
        "Items" => [
            [


                "FileTypeExt" => 'RTF',
                "FileTypeDescription" => 'Windows Rich Text Format',
                "FileFormat" => 'wdFormatRTF',
                "RelatedLinks" => []
            ],

            [
                "FileTypeExt" => 'PDF',
                "FileTypeDescription" => 'Adobe PDF Format',
                "FileFormat" => 'wdFormatPDF',
                "RelatedLinks" => []
            ],
            [
                "FileTypeExt" => 'TXT',
                "FileTypeDescription" => 'Text File',
                "FileFormat" => 'wdFormatText',
                "RelatedLinks" => []
            ],
            [
                "FileTypeExt" => 'ODD',
                "FileTypeDescription" => 'Open Document Text Format',
                "FileFormat" => 'wdFormatStrictOpenXMLDocument',
                "RelatedLinks" => []
            ],
            [
                "FileTypeExt" => 'XML',
                "FileTypeDescription" => 'XML Document Format',
                "FileFormat" => 'wdFormatXMLDocument',
                "RelatedLinks" => []
            ],
            [
                "FileTypeExt" => 'HTML',
                "FileTypeDescription" => 'HTML File',
                "FileFormat" => 'wdFormatHTML',
                "RelatedLinks" => []
            ],
            [
                "FileTypeExt" => 'HTML',
                "FileTypeTitleExtra" => 'Filtered',
                "FileTypeDescription" => 'Filtered HTML File',
                "FileFormat" => 'wdFormatFilteredHTML',
                "RelatedLinks" => []
            ],

              [
                "FileTypeExt" => 'XPS',
                "FileTypeDescription" => 'Microsoft XPS Format',
                "FileFormat" => 'wdFormatXPS',
                "RelatedLinks" => []
            ],
              [
                "FileTypeExt" => 'DOC',
                "FileTypeDescription" => 'Latest Microsoft Office Word 365 Version Format ',
                "FileFormat" => 'wdFormatDocumentDefault',
                "RelatedLinks" => []
            ],
        ],
    ],
   "ConvertDirDocToFile_NewExt" => [ 
        "Description" => "Convert all Word Documents in a Directory and save with differnt extension",
        "Template" => "ConvertDirDocToFile_NewExt.tpl.md",
        "Items" => [

              [
                "FileTypeExt" => 'DOC',
                "FileTypeDescription" => 'Latest Microsoft Office Word 365 Version Format ',
                "FileFormat" => 'wdFormatDocumentDefault',
                "FileNewExt" => '.DOCX',
                "RelatedLinks" => []
            ],
        ],
    ],
   

    "ConvertXLSToFile" => [ 
        "Description" => "Convert Excel Spreadsheet to another file type",
        "Title" => 'How do I Convert a Microsoft Excel Spreadsheet to a {[$Command.FileTypeExt]}? ',
       "Template" => "ConvertFromXLSToFile.tpl.md",
        "Items" => [
            [


                "FileTypeExt" => 'CSV',
                "FileTypeDescription" => 'Comma Seperated Values file',
                "FileFormat" => 'xlCSV',
                "RelatedLinks" => []
            ],

            [
                "FileTypeExt" => 'PDF',
                "FileTypeDescription" => 'Adobe PDF Format',
                "FileFormat" => 'xlpdf',
                "RelatedLinks" => []
            ],
            [
                "FileTypeExt" => 'TXT',
                "FileTypeDescription" => 'Text File',
                "FileFormat" => 'xlTextWindows',
                "RelatedLinks" => []
            ],[
                "FileTypeExt" => 'TXT',
                "FileTypeDescription" => 'Unicode Text File',
                "FileTypeTitleExtra" => 'Unicode',
                "FileFormat" => 'xlUnicodeText',
                "RelatedLinks" => []
            ],
            [
                "FileTypeExt" => 'ODS',
                "FileTypeDescription" => 'Open Document Spreadsheet Format',
                "FileFormat" => 'xlOpenDocumentSpreadsheet',
                "RelatedLinks" => []
            ],
            [
                "FileTypeExt" => 'XML',
                "FileTypeDescription" => 'Open XML Workbook Format',
                "FileFormat" => 'xlOpenXMLWorkbook',
                "RelatedLinks" => []
            ],
            [
                "FileTypeExt" => 'HTML',
                "FileTypeDescription" => 'HTML File',
                "FileFormat" => 'xlHtml',
                "RelatedLinks" => []
            ],
            [
                "FileTypeExt" => 'xls',
                "FileTypeTitleExtra" => '9795',
                "FileTypeDescription" => 'Excel 97/95 format',
                "FileFormat" => 'xlExcel9795',
                "RelatedLinks" => []
            ],

              [
                "FileTypeExt" => 'XPS',
                "FileTypeDescription" => 'Microsoft XPS Format',
                "FileFormat" => 'xlXPS',
                "RelatedLinks" => []
            ]
        ],
    ],

    "ConvertXLSToFileUNC" => [ 
        "Description" => "Convert Excel Spreadsheet to another file type and save to UNC path",
        "Title" => 'How do I Convert a Microsoft Excel Spreadsheet to a {[$Command.FileTypeExt]} and save it to a UNC network Drive? ',
        "Items" => [
            [


                "FileTypeExt" => 'CSV',
                "FileTypeDescription" => 'Comma Seperated Values file',
                "FileFormat" => 'xlCSV',
                "RelatedLinks" => []
            ],

            [
                "FileTypeExt" => 'PDF',
                "FileTypeDescription" => 'Adobe PDF Format',
                "FileFormat" => 'xlpdf',
                "RelatedLinks" => []
            ],
            [
                "FileTypeExt" => 'TXT',
                "FileTypeDescription" => 'Text File',
                "FileFormat" => 'xlTextWindows',
                "RelatedLinks" => []
            ],
            [
                "FileTypeExt" => 'HTML',
                "FileTypeDescription" => 'HTML File',
                "FileFormat" => 'xlHtml',
                "RelatedLinks" => []
            ],
            [
                "FileTypeExt" => 'xls',
                "FileTypeTitleExtra" => '9795',
                "FileTypeDescription" => 'Excel 97/95 format',
                "FileFormat" => 'xlExcel9795',
                "RelatedLinks" => []
            ]

              
        ],
       "Template" => "ConvertFromXLSToFileUNC.tpl.md"
    ],
        "ConvertPPTToFile" => [ 
        "Description" => "Convert Microsoft PowerPoint Presentation to another file type",
        "Template" => "ConvertFromPPTToFile.tpl.md",
        "Blocks" => [
            [
                ""
            ]
        ],
        "Items" => [
  

            [
                "FileTypeExt" => 'PDF',
                "FileTypeDescription" => 'Adobe PDF Format',
                "FileFormat" => 'ppSaveAsPDF',
                "RelatedLinks" => []
            ],            [
                "FileTypeExt" => 'RTF',
                "FileTypeDescription" => 'Rich Text Format',
                "FileFormat" => 'ppSaveAsRTF',
                "RelatedLinks" => []
            ],
            [
                "FileTypeExt" => 'JPG',
                "FileTypeDescription" => 'Jpeg Image file',
                "FileFormat" => 'ppSaveAsJPG',
                "RelatedLinks" => []
            ],
            [
                "FileTypeExt" => 'pptx',
                "FileTypeDescription" => 'Open Document Presentation Format',
                "FileFormat" => 'ppSaveAsOpenXMLPresentation',
                "RelatedLinks" => []
            ],
            [
                "FileTypeExt" => 'png',
                "FileTypeDescription" => 'PNG File format',
                "FileFormat" => 'ppSaveAsPNG',
                "RelatedLinks" => []
            ],
            [
                "FileTypeExt" => 'gif',
                "FileTypeDescription" => 'Animated Gif',
                "FileFormat" => 'ppSaveAsAnimatedGIF',
                "RelatedLinks" => []
            ],
            [
                "FileTypeExt" => 'HTML',
                "FileTypeTitleExtra" => 'Filtered',
                "FileTypeDescription" => 'Filtered HTML File',
                "FileFormat" => 'wdFormatFilteredHTML',
                "RelatedLinks" => []
            ],

              [
                "FileTypeExt" => 'XPS',
                "FileTypeDescription" => 'Microsoft XPS Format',
                "FileFormat" => 'ppSaveAsXPS',
                "RelatedLinks" => []
            ],
              [
                "FileTypeExt" => 'PPTX',
                "FileTypeDescription" => 'Latest Microsoft Office PowerPoint 365 Version Format ',
                "FileFormat" => 'ppSaveAsOpenXMLPresentation',
                "RelatedLinks" => []
            ],
        ],
    ],

      "ConvertDocToFileLog" => [ 
        "Description" => "Output Log values while Converting Word Document to PDF",
        "Template" => "ConvertDocToFileLog.tpl.md",
        "Blocks" => [
            [
                ""
            ]
        ],
        "FileFormat" => [
                "FileTypeExt" => 'PDF',
                "FileTypeDescription" => 'Adobe PDF Format',
                "FileFormat" => 'wdFormatPDF',
        ],
        "Items" => [


            [

                "LogLevel" => '1',
                "LogLevelDesc" => 'Display Errors Only',
                "LogLevelExtendedDesc" => 'Display Errors only',
                'FileTypeTitleExtra' => 'L1',
                "RelatedLinks" => []
            ],
            [
                "LogLevel" => '2',
                "LogLevelDesc" => 'Standard Log output',
                "LogLevelExtendedDesc" => 'If you dont specify a value, this is the value.',
                'FileTypeTitleExtra' => 'L2',
                "RelatedLinks" => []
            ],
            [
                "LogLevel" => '5',
                "LogLevelDesc" => 'Chatty',
                'FileTypeTitleExtra' => 'L5',
                "LogLevelExtendedDesc" => 'Will provide some more information about the process of conversion',
                "RelatedLinks" => []
            ],
            [
                "LogLevel" => '9',
                'FileTypeTitleExtra' => 'L9',
                "LogLevelDesc" => 'Debug',
                "LogLevelExtendedDesc" => 'Some extra information above chatty',
                "RelatedLinks" => []
            ],
           
            [
                "LogLevel" => '10',
                'FileTypeTitleExtra' => 'L10',
                "LogLevelDesc" => 'Verbose',
                "LogLevelExtendedDesc" => 'A large amount of information will be output with the conversion.  Useful for debugging and trying to track down an issue.',
                "RelatedLinks" => []
            ],
           
        ],
    ],


    "ConvertDocToFileWithWebhook" => [ 
        "Description" => "Convert Word Document to another file type then fire a Webhook",
        "Items" => [
            [


                "FileTypeExt" => 'RTF',
                "FileTypeDescription" => 'Windows Rich Text Format',
                "FileFormat" => 'wdFormatRTF',
                "RelatedLinks" => []
            ],

            [
                "FileTypeExt" => 'PDF',
                "FileTypeDescription" => 'Adobe PDF Format',
                "FileFormat" => 'wdFormatPDF',
                "RelatedLinks" => []
            ],
            [
                "FileTypeExt" => 'ODD',
                "FileTypeDescription" => 'Open Document Text Format',
                "FileFormat" => 'wdFormatStrictOpenXMLDocument',
                "RelatedLinks" => []
            ],
          
            [
                "FileTypeExt" => 'HTML',
                "FileTypeDescription" => 'HTML File',
                "FileFormat" => 'wdFormatHTML',
                "RelatedLinks" => []
            ],
            [
                "FileTypeExt" => 'HTML',
                "FileTypeTitleExtra" => 'Filtered',
                "FileTypeDescription" => 'Filtered HTML File',
                "FileFormat" => 'wdFormatFilteredHTML',
                "RelatedLinks" => []
            ],

              [
                "FileTypeExt" => 'XPS',
                "FileTypeDescription" => 'Microsoft XPS Format',
                "FileFormat" => 'wdFormatXPS',
                "RelatedLinks" => []
            ],
              [
                "FileTypeExt" => 'DOC',
                "FileTypeDescription" => 'Latest Microsoft Office Word 365 Version Format ',
                "FileFormat" => 'wdFormatDocumentDefault',
                "RelatedLinks" => []
            ],
        ],
       "Template" => "ConvertDocToFileWithWebhook.tpl.md"
    ],

    "ShowVersion" => [
         "Description" => "Show Version of DocTo",
        "Items" => [
            [
                "FileTypeExt" => 'Version',
                "FileTypeDescription" => "Version"
             

              
            ]

           
        ],
       "Template" => "ShowVersion.tpl.md"
    ],

    "OneDriveConversion" => [
         "Description" => "Convert Word Docs for upload to Onedrive",
        "Items" => [
            [
                "FileTypeExt" => 'OneDrive',
                "FileTypeDescription" => "Onedrive"
             

              
            ]

           
        ],
       "Template" => "ConvertDirDocOneDrive.tpl.md"
    ],

    "BookmarksFromSource" => [
         "Description" => "Show boomarks from Word document in PDF",
        "Items" => [
            [
                "FileTypeExt" => 'pdf',
                "FileTypeDescription" => "pdf",
                "FileTypeTitleExtra" => "WordBookmarks",
                "BookmarksSource" => "WordBookmarks",
                "ContentExtra" => "",
                 "RelatedLinks" => []

            ],
            [
                "FileTypeExt" => 'pdf',
                "FileTypeDescription" => "pdf",
                "FileTypeTitleExtra" => "None",
                "BookmarksSource" => "None",
                "ContentExtra" => "With this option no bookmarks are present in the output PDF file",
                 "RelatedLinks" => []

            ],
              

           
        ],
       "Template" => "ConvertDocToPDFBookmarks.tpl.md"
    ],

     "CommentsFromWord" => [
         "Description" => "Show comments from Word document in PDF",
        "Items" => [
            [
                "FileTypeExt" => 'pdf',
                "FileTypeDescription" => "pdf",
                "FileTypeTitleExtra" => "WordComments",
                "ExportComments" => "wdExportDocumentWithMarkup",
                "ContentExtra" => "With this option comments are present in the output PDF file",
                 "RelatedLinks" => []

            ],
            [
                "FileTypeExt" => 'pdf',
                "FileTypeDescription" => "pdf",
                "FileTypeTitleExtra" => "NoComments",
                "ExportComments" => "wdExportDocumentContent",
                "ContentExtra" => "With this option no comments are present in the output PDF file",
                 "RelatedLinks" => []

            ],
              

           
        ],
       "Template" => "ConvertDocToPDFComments.tpl.md"
    ],

     "ConvertWithoutSubdirs" => [
         "Description" => "Do not recurse SubDirs",
        "Items" => [
            [
                "FileTypeExt" => 'pdf',
                "FileTypeDescription" => "pdf",
                "FileTypeTitleExtra" => "",
                
                "ContentExtra" => "With this option sub dirs will not be searched for files to convert",
                 "RelatedLinks" => []

            ],

              

           
        ],
       "Template" => "ConvertDirDocNoRecurse.tpl.md"
    ],
    
     "ConvertToPDFArchive" => [
         "Description" => "Convert Word Document to PDF Files at <a href='https://www.iso.org/standard/38920.html'>ISO 19005-1</a> standard (PDF/A) format.",
        "Items" => [
            [
                "FileTypeExt" => 'pdf',
                "FileTypeDescription" => "pdf",
                "FileTypeTitleExtra" => "iso19005-1",
                
                "ContentExtra" => "Save as PDF/A",
                 "RelatedLinks" => []

            ],

              

           
        ],
       "Template" => "ConvertDoctoPDFA.tpl.md"
    ],
    
    "ConvertVSDToFile" => [ 
        "Description" => "Convert Visio Document to another file type",
        "Template" => "ConvertFromVsdToFile.tpl.md",
        "Title" => 'How do I Convert a Microsoft Visio Document to a {[$Command.FileTypeDescription]}?'       , 
        "Blocks" => [
            [
                ""
            ]
        ],
        "Items" => [


            [
                "FileTypeExt" => 'PDF',
                "FileTypeDescription" => 'Adobe PDF Format',
                "FileFormat" => 'visFixedFormatPDF',
                "RelatedLinks" => []
            ],
           
              [
                "FileTypeExt" => 'XPS',
                "FileTypeDescription" => 'Microsoft XPS Format',
                "FileFormat" => 'visFixedFormatXPS',
                "RelatedLinks" => []
            ],

        ],
    ],


];

$Explain = [

    "appwd" => ['cmd' => '-WD' , 'desc' => "This is a conversion using Microsoft Word. "],
    "appxl" => ['cmd' => '-XL' , 'desc' => "This is a conversion using Microsoft Excel.  "],
    "apppp" => ['cmd' => '-PP' , 'desc' => "This is a conversion using Microsoft PowerPoint.  "],
    "appvs" => ['cmd' => '-VS' , 'desc' => "This is a conversion using Microsoft Visio.  "],
    "dashf" => ['cmd' => '-f' , 'desc' => "The File or directory to be converted"],
    "dashfx" => ['cmd' => '-fx' , 'desc' => "File with this extension should be matched."],
    "dasho" => ['cmd' => '-o' , 'desc' => "The Output File or Directory where you would like the converted file to be written to."],
    "dashox" => ['cmd' => '-ox' , 'desc' => "The Output extension to be used for converted files"],    
    "dasht" => ['cmd' => '-T' , 'desc' => "The file format type that is being converted to"],
    "dashw" => ['cmd' => '-W' , 'desc' => "Fire a webhook to the provided URL. See Below for more details"],
    "dashv" => ['cmd' => '-V' , 'desc' => "Show Version details for DocTo and loaded Office Applications"],
    "dashr" => ['cmd' => '-R' , 'desc' => "Delete the converted file after conversion."],
    "dashl" => ['cmd' => '-L' , 'desc' => "Set output Log level."],
    "dashc" => ['cmd' => '-c' , 'desc' => "Compatibility Level with Office Versions"],

    
    // -- Double Dash
    "dddeletefiles" => ['cmd' => '--deletefiles' , 'desc' => "Delete the converted file after conversion."],
    "ddbookmarksource" => ['cmd' => '--bookmarksource' , 'desc' => "Where to get Bookmarks from for PDF.  This parameter is only relevant for PDF files."],
    "ddexportmarkup" => ['cmd' => '--ExportMarkup', 'desc' => "Export Comments and other markup from Word Document to PDF"],
    "ddnosubdirs" => ['cmd' => '--no-subdirs', 'desc' => 'Don\'t rescurse subdirs.  Only convert files in requested dir. '],
    "dduseISO190051" => ['cmd' => '--use-iso19005-1', 'desc' => 'Output PDFs from Word as ISO standard 19005-1 for self contained PDFS. Sometimes also refered to as <a href="https://en.wikipedia.org/wiki/PDF/A">PDF/A</a> '],
    


];

