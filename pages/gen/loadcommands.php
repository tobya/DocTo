<?php

$Commands = [





    "ConvertDocToFile" => [ 
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
       "Template" => "ConvertFromDocToFile.md"
    ],
   

    "ConvertXLSToFile" => [ 
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
       "Template" => "ConvertFromXLSToFile.md"
    ]
   

];

$Explain = [

    "appwd" => ['cmd' => '-WD' , 'desc' => "This is a conversion using Microsoft Word.  This is not required but makes it easier to read"],
    "appxl" => ['cmd' => '-XL' , 'desc' => "This is a conversion using Microsoft Excel.  "],
    "dashf" => ['cmd' => '-f' , 'desc' => "The File or directory to be converted"],
    "dasho" => ['cmd' => '-o' , 'desc' => "The File or Directory where you would like the converted file to be written to."],
    "dashl" => ['cmd' => '-L' , 'desc' => "The log level to be output."],
    "dasht" => ['cmd' => '-T' , 'desc' => "The file format type that is being converted to"]
];

/*


    "ConvertDocToPDF" => [ 
        "Title" => "How do I Convert a Word Doc to a PDF? ",
        
        "CMDValues" => [
            '-WD',
            '-f',
            "'c:\path\Document.doc'",
            "-o",
            "'c:\path\output\Document.pdf'"

        ],
        "Description" => "It is very simple to convert a Word Document to a  PDF on the command line using [Docto](https://github.com/tobya/docto). You can also do this easily in word, but sometimes it helps to be able to do it from the command line.  The command line below shows how you can convert a Microsoft Word Document to a Adobe Acrobat PDF file.",
        "RelatedLinks" => [
        "Convert a Directory to PDF" => "ConvertFolderToPDF"
        ]
    ],

    ,
    "ConvertDocToRTF" => [ 
    "Title" => "How do I Convert a Word Document to a RTF File? ",
    
    "CMDValues" => [
        '-WD',
        '-f',
        "'c:\path\Document.doc'",
        "-o",
        "'c:\path\output\Document.rtf'"

    ],
    "Description" => "It is very simple to convert a Word Document to a  RTF on the command line using [Docto](https://github.com/tobya/docto). You can also do this easily in word, but sometimes it helps to be able to do it from the command line.  The command line below shows how you can convert a Microsoft Word Document to a Windows Rich Text Format file - RTF.",
    "RelatedLinks" => [
    "Convert a Directory to RTF" => "ConvertFolderToRTF"
    ]
    ],
    "Name" => [ 
        "Title" => "",
        
        "CMDValues" => [
            '-WD',
        

        ],
        "Description" => "",
        "RelatedLinks" => [
        "" => ""
            ]
    ]

*/
