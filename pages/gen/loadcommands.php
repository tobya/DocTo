<?php

$Commands = [





    "ConvertDocToFile" => [ 
        "Description" => "Convert Word Document to another file type",
        "Template" => "ConvertFromDocToFile.tpl.md",
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
                "FileTypeDescription" => 'Adobe Acrobat Portable Document Format',
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
                "FileTypeDescription" => 'Adobe Acrobat Portable Document Format',
                "FileFormat" => 'wdFormatPDF',
        ],
        "Items" => [


            [

                "LogLevel" => '1',
                "LogLevelDesc" => 'Display Errors Only',
                "RelatedLinks" => []
            ],
            [
                "LogLevel" => '2',
                "LogLevelDesc" => 'Standard Log output',
                "LogLevelExtendedDesc" => 'If you dont specify a value, this is the value.',
                "RelatedLinks" => []
            ],
            [
                "LogLevel" => '5',
                "LogLevelDesc" => 'Chatty',
                "LogLevelExtendedDesc" => 'Will provide some more information about the process of conversion',
                "RelatedLinks" => []
            ],
            [
                "LogLevel" => '9',
                "LogLevelDesc" => 'Debug',
                "LogLevelExtendedDesc" => 'Some extra information above chatty',
                "RelatedLinks" => []
            ],
           
            [
                "LogLevel" => '10',
                "LogLevelDesc" => 'Verbose',
                "LogLevelExtendedDesc" => 'A large amount of information will be output with the conversion.  Useful for debugging and trying to track down an issue.',
                "RelatedLinks" => []
            ],
           
        ],
    ],
   "ConvertDirDocToFile" => [ 
        "Description" => "Convert all Word Documents in a Directory",
        "Template" => "ConvertDirDocToFile.tpl.md",
        "Items" => [
            [


                "FileTypeExt" => 'RTF',
                "FileTypeDescription" => 'Windows Rich Text Format',
                "FileFormat" => 'wdFormatRTF',
                "RelatedLinks" => []
            ],

            [
                "FileTypeExt" => 'PDF',
                "FileTypeDescription" => 'Adobe Acrobat Portable Document Format',
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
    "ConvertDirToFileRemove" => [ 
        "Description" => "Convert all Word Documents in a Directory and remove after conversion",
        "Template" => "ConvertDirToFileandRemove.tpl.md",
        "Items" => [
            [


                "FileTypeExt" => 'RTF',
                "FileTypeDescription" => 'Windows Rich Text Format',
                "FileFormat" => 'wdFormatRTF',
                "RelatedLinks" => []
            ],

            [
                "FileTypeExt" => 'PDF',
                "FileTypeDescription" => 'Adobe Acrobat Portable Document Format',
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
   

    "ConvertXLSToFile" => [ 
        "Description" => "Convert Excel Spreadsheet to another file type",
        "Items" => [
            [


                "FileTypeExt" => 'CSV',
                "FileTypeDescription" => 'Comma Seperated Values file',
                "FileFormat" => 'xlCSV',
                "RelatedLinks" => []
            ],

            [
                "FileTypeExt" => 'PDF',
                "FileTypeDescription" => 'Adobe Acrobat Portable Document Format',
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
       "Template" => "ConvertFromXLSToFile.tpl.md"
    ],
    "ConvertDocToFileWithWebhook" => [ 
        "Description" => "Convert Word Document to another file type with Webhook",
        "Items" => [
            [


                "FileTypeExt" => 'RTF',
                "FileTypeDescription" => 'Windows Rich Text Format',
                "FileFormat" => 'wdFormatRTF',
                "RelatedLinks" => []
            ],

            [
                "FileTypeExt" => 'PDF',
                "FileTypeDescription" => 'Adobe Acrobat Portable Document Format',
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
    ]
   

];

$Explain = [

    "appwd" => ['cmd' => '-WD' , 'desc' => "This is a conversion using Microsoft Word.  This is not required but makes it easier to read"],
    "appxl" => ['cmd' => '-XL' , 'desc' => "This is a conversion using Microsoft Excel.  "],
    "dashf" => ['cmd' => '-f' , 'desc' => "The File or directory to be converted"],
    "dasho" => ['cmd' => '-o' , 'desc' => "The Output File or Directory where you would like the converted file to be written to."],
    "dashl" => ['cmd' => '-L' , 'desc' => "The log level to be output."],
    "dasht" => ['cmd' => '-T' , 'desc' => "The file format type that is being converted to"],
    "dashw" => ['cmd' => '-W' , 'desc' => "Fire a webhook to the provided URL. See Below for more details"],
    "dashv" => ['cmd' => '-V' , 'desc' => "Show Version details for DocTo and loaded Office Applications"],
    "dashr" => ['cmd' => '-R' , 'desc' => "Delete the converted file after conversion."],
    "dashl" => ['cmd' => '-L' , 'desc' => "Set Log level for outpu."],

    "dddeletefiles" => ['cmd' => '--deletefiles' , 'desc' => "Delete the converted file after conversion."],


];
